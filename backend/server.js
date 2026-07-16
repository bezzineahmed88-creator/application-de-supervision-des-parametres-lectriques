const express = require('express'); // Importation du module Express pour créer le serveur
const cors = require('cors'); // Importation du module CORS pour gérer les requêtes cross-origin
const XLSX = require('xlsx'); // Importation du module XLSX pour manipuler les fichiers Excel
const fs = require('fs'); // Importation du module fs pour gérer les fichiers
const path = require('path'); // Importation du module path pour gérer les chemins de fichiers
const snap7 = require('node-snap7'); // Importation du module node-snap7 pour communiquer avec les automates Siemens S7
const http = require('http'); // Pour créer un serveur HTTP autour d'Express (nécessaire pour Socket.IO)
const { Server } = require('socket.io'); // Pour la communication temps réel avec Angular
const { initializeApp, cert } = require('firebase-admin/app'); // Firebase Admin (initialisation)
const { getMessaging } = require('firebase-admin/messaging'); // Firebase Admin (envoi de notifications push)

// Chemin exact vers ta clé de service Firebase (doit être dans le même dossier que server.js)
const serviceAccount = require('./mesures-electriques-app-firebase-adminsdk-fbsvc-f2b7063962.json');

const app = express(); // Création d'une instance de l'application Express
const server = http.createServer(app); // Serveur HTTP qui enveloppe Express
const io = new Server(server, { // Socket.IO attaché au même serveur HTTP
  cors: {
    origin: '*', // Autorise Angular (localhost:4200, 10.0.2.2, etc.) à se connecter
  },
});

// Initialisation de Firebase Admin — DOIT être fait avant d'utiliser getMessaging()
initializeApp({
  credential: cert(serviceAccount),
});

// ============================================================
// MIDDLEWARE : Configuration d'Express
// ⚠️ DOIT ÊTRE PLACÉ ICI, AVANT TOUTES LES ROUTES (app.get / app.post...)
// ============================================================
app.use(cors()); // Autorise les requêtes depuis Angular (localhost:4200 ou 10.0.2.2)
app.use(express.json({ limit: '10mb' })); // Permet de lire le corps des requêtes JSON

// Définition du chemin vers le fichier Excel dans le dossier public du frontend
const frontendExcelPath = path.join(__dirname, '..', 'my-app', 'public', 'mesures.xlsx');

// ============================================================
// Stockage des tokens des appareils enregistrés pour les notifications push
// ============================================================
let tokensEnregistres = [];

app.post('/api/register-token', (req, res) => {
  const { token } = req.body;
  if (token && !tokensEnregistres.includes(token)) {
    tokensEnregistres.push(token);
    console.log('✅ Token enregistré pour notifications push :', token);
  }
  res.json({ message: 'Token enregistré' });
});

// ⚙️ Configuration de l'automate Siemens S7-1214C
const IP_AUTOMATE = '192.168.1.40'; // Adresse IP de l'automate Siemens S7
const RACK = 0; // Rack de l'automate (toujours 0 pour S7-1200)
const SLOT = 1; // Slot de l'automate (toujours 1 pour S7-1200)

// Création d'une instance du client Snap7
const client = new snap7.S7Client();

// ============================================================
// FONCTION : Connexion à l'automate Siemens S7
// ============================================================
function connecterAutomate() {
  try {
    if (client.Connected()) {
      return;
    }

    client.ConnectTo(IP_AUTOMATE, RACK, SLOT, (err) => {
      if (err) {
        console.error('❌ Erreur de connexion à l\'automate :', client.ErrorText(err));
      } else {
        console.log('✅ Connexion réussie à l\'automate Siemens S7-1214C');
      }
    });
  } catch (error) {
    console.error('❌ Exception lors de la connexion à l\'automate :', error.message);
  }
}

// ============================================================
// ROUTE GET /api/status : Vérifie si l'automate est connecté
// ============================================================
app.get(['/api/status', '/api/statut'], (req, res) => {
  if (client.Connected()) {
    res.json({ connecte: true });
  } else {
    res.json({ connecte: false });
  }
});

// ============================================================
// FONCTION : Formate une valeur flottante à 2 décimales
// ============================================================
function formaterReel(value) {
  if (typeof value !== 'number' || Number.isNaN(value)) {
    return null;
  }
  return Number(value.toFixed(2));
}

// ============================================================
// FONCTION : Lecture d'une valeur flottante depuis le résultat de lecture
// ============================================================
function lireValeurFloat(result) {
  if (!result || !result.Data || !Buffer.isBuffer(result.Data)) {
    console.error('❌ Résultat de lecture invalide ou vide :', result);
    return null;
  }
  try {
    const floatValue = result.Data.readFloatBE(0);
    return formaterReel(floatValue);
  } catch (error) {
    console.error('❌ Erreur de lecture du flottant depuis l\'automate :', error.message);
    return null;
  }
}

// ============================================================
// FONCTION : Lecture simultanée de plusieurs variables depuis l'automate
// ============================================================
function lireVariablesEnBloc(variables, callback) {
  const fallbackResults = [];
  let pending = variables.length;

  if (pending === 0) {
    callback(null, fallbackResults);
    return;
  }

  variables.forEach(function (variable, index) {
    client.ReadArea(
      variable.Area || client.S7AreaDB,
      variable.DBNumber || 0,
      variable.Start,
      variable.Amount,
      variable.WordLen || client.S7WLReal,
      function (readErr, data) {
        fallbackResults[index] = {
          Result: readErr ? 1 : 0,
          Data: readErr ? null : data,
          Error: readErr ? client.ErrorText(readErr) : null,
        };

        pending -= 1;
        if (pending === 0) {
          callback(null, fallbackResults);
        }
      }
    );
  });
}

// ============================================================
// Seuils de détection d'anomalie (à ajuster selon ton installation)
// ============================================================
const seuils = {
  tension:   { min: 200, max: 250 },   // exemple pour 230V nominal
  courant:   { min: 10,  max: 32 },    // exemple pour disjoncteur 32A
  puissance: { min: 5,   max: 7000 },  // exemple en W
};

// ============================================================
// FONCTION : Vérifie si une mesure est anormale
// ============================================================
function estAnormale(mesure) {
  const tensionsAnormales = [mesure.v1, mesure.v2, mesure.v3].some(
    v => v < seuils.tension.min || v > seuils.tension.max
  );

  const courantsAnormaux = [mesure.i1, mesure.i2, mesure.i3].some(
    i => i < seuils.courant.min || i > seuils.courant.max
  );

  const puissancesAnormales = [mesure.p1, mesure.p2, mesure.p3].some(
    p => p < seuils.puissance.min || p > seuils.puissance.max
  );

  return tensionsAnormales || courantsAnormaux || puissancesAnormales;
}

// ============================================================
// FONCTION : Envoie une notification push via Firebase Cloud Messaging
// ============================================================
async function envoyerNotificationPush(title, corps) {
  if (tokensEnregistres.length === 0) {
    console.log('⚠️ Aucun token enregistré pour les notifications push');
    return;
  }
  const message = {
    notification: { title: title, body: corps },
    tokens: tokensEnregistres,
  };
  try {
    const response = await getMessaging().sendEachForMulticast(message);
    console.log(`✅ ${response.successCount} notification(s) envoyée(s)`);
  } catch (error) {
    console.log('❌ Erreur envoi push :', error);
  }
}

// ============================================================
// FONCTION : Lecture des données depuis l'automate
// (Cette fonction lit les VRAIES valeurs de l'automate et les sauvegarde
//  telles quelles si elles sont anormales — pas de valeurs fictives ici)
// ============================================================
function lireDonneesAutomate() {
  if (!client.Connected()) {
    console.log('⚠️ Automate non connecté, lecture annulée');
    connecterAutomate();
    return;
  }
  const variables = [
    // --- Courants (offset 0, 4, 8 dans DB3) ---
    { Area: client.S7AreaDB, WordLen: client.S7WLReal, DBNumber: 3, Start: 0,   Amount: 1 }, // I1
    { Area: client.S7AreaDB, WordLen: client.S7WLReal, DBNumber: 3, Start: 4,   Amount: 1 }, // I2
    { Area: client.S7AreaDB, WordLen: client.S7WLReal, DBNumber: 3, Start: 8,   Amount: 1 }, // I3

    // --- Tensions (offset 56, 60, 64 dans DB3) ---
    { Area: client.S7AreaDB, WordLen: client.S7WLReal, DBNumber: 3, Start: 56,  Amount: 1 }, // V1
    { Area: client.S7AreaDB, WordLen: client.S7WLReal, DBNumber: 3, Start: 60,  Amount: 1 }, // V2
    { Area: client.S7AreaDB, WordLen: client.S7WLReal, DBNumber: 3, Start: 64,  Amount: 1 }, // V3

    // --- Puissances (offset 108, 112, 116 dans DB3) ---
    { Area: client.S7AreaDB, WordLen: client.S7WLReal, DBNumber: 3, Start: 108, Amount: 1 }, // P1
    { Area: client.S7AreaDB, WordLen: client.S7WLReal, DBNumber: 3, Start: 112, Amount: 1 }, // P2
    { Area: client.S7AreaDB, WordLen: client.S7WLReal, DBNumber: 3, Start: 116, Amount: 1 }, // P3

    // --- Puissance active totale (offset 120 dans DB3) ---
    { Area: client.S7AreaDB, WordLen: client.S7WLReal, DBNumber: 3, Start: 120, Amount: 1 }, // P_active_totale

    // --- Facteur de puissance (offset 168 dans DB3) ---
    { Area: client.S7AreaDB, WordLen: client.S7WLReal, DBNumber: 3, Start: 168, Amount: 1 }, // FC_PUISSANCE
  ];

  try {
    lireVariablesEnBloc(variables, function (err, results) {
      if (err) {
        console.error('❌ Erreur lecture variables :', client.ErrorText(err));
        return;
      }

      // ✅ Construction de l'objet mesure avec les VRAIES valeurs lues depuis l'automate
      const mesure = {
        date:         new Date().toISOString().slice(0, 16),
        i1:           lireValeurFloat(results[0]),
        i2:           lireValeurFloat(results[1]),
        i3:           lireValeurFloat(results[2]),
        v1:           lireValeurFloat(results[3]),
        v2:           lireValeurFloat(results[4]),
        v3:           lireValeurFloat(results[5]),
        p1:           lireValeurFloat(results[6]) * 1000,
        p2:           lireValeurFloat(results[7]) * 1000,
        p3:           lireValeurFloat(results[8]) * 1000,
        p_totale:     lireValeurFloat(results[9]) * 1000,
        fc_puissance: lireValeurFloat(results[10]),
      };

      console.log('📊 Mesure lue depuis DB3 :', mesure);

      // Envoi en temps réel au frontend via WebSocket (toujours, même si normale)
      io.emit('mesures', mesure);

      // ✅ Sauvegarde dans le fichier Excel UNIQUEMENT si la mesure est anormale
      //    → "mesure" contient les valeurs RÉELLEMENT lues depuis l'automate, pas des valeurs fictives
      if (estAnormale(mesure)) {
        console.log('🚨 Anomalie détectée — sauvegarde de la mesure');
        sauvegarderMesuresDansFichier([mesure]);

        io.emit('anomalie', mesure);
        envoyerNotificationPush(
          '⚠️ Anomalie détectée',
          `Valeurs anormales détectées à ${new Date().toLocaleTimeString()}`
        );
      } else {
        console.log('✅ Mesure normale — non sauvegardée');
      }
    });
  } catch (error) {
    console.error('❌ Exception lors de la lecture des variables :', error.message);
  }
}

// ============================================================
// FONCTION : Sauvegarde des mesures dans le fichier Excel
// Le fichier est dans public/ d'Angular pour être lu par le frontend
// ============================================================
function sauvegarderMesuresDansFichier(nouvellesMesures) {
  let workbook;
  let mesuresExistantes = [];

  if (fs.existsSync(frontendExcelPath)) {
    workbook = XLSX.readFile(frontendExcelPath);
    const sheetName = workbook.SheetNames[0] || 'Mesures';
    const sheet = workbook.Sheets[sheetName];
    mesuresExistantes = sheet ? XLSX.utils.sheet_to_json(sheet, { defval: '' }) : [];
  } else {
    workbook = XLSX.utils.book_new();
  }

  const toutesLesMesures = [...mesuresExistantes, ...nouvellesMesures];
  const worksheet = XLSX.utils.json_to_sheet(toutesLesMesures);

  if (workbook.SheetNames && workbook.SheetNames.length > 0) {
    const sheetName = workbook.SheetNames[0];
    workbook.Sheets[sheetName] = worksheet;
  } else {
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Mesures');
  }

  fs.mkdirSync(path.dirname(frontendExcelPath), { recursive: true });
  XLSX.writeFile(workbook, frontendExcelPath);
  console.log('✅ Mesure sauvegardée dans Excel à :', new Date().toISOString().slice(0, 16));
}

// ============================================================
// ROUTE GET /api/test-anomalie : route de TEST temporaire
// Permet de simuler une mesure anormale sans avoir besoin de l'automate
// ⚠️ À retirer une fois les tests terminés
// ============================================================
app.get('/api/test-anomalie', (req, res) => {
  const mesureTest = {
    date: new Date().toISOString().slice(0, 16),
    i1: 0,   // en dehors du seuil 10-32 → déclenche l'anomalie
    i2: 15,
    i3: 15,
    v1: 220, v2: 220, v3: 220,
    p1: 1000, p2: 1000, p3: 1000,
  };

  console.log('📊 Mesure de test :', mesureTest);
  io.emit('mesures', mesureTest);

  if (estAnormale(mesureTest)) {
    console.log('🚨 Anomalie détectée (test) — sauvegarde de la mesure');
    sauvegarderMesuresDansFichier([mesureTest]);
    io.emit('anomalie', mesureTest);
    envoyerNotificationPush(
      '⚠️ Anomalie détectée',
      `Valeurs anormales détectées à ${new Date().toLocaleTimeString()}`
    );
  }

  res.json({ success: true, mesure: mesureTest });
});

// ============================================================
// ROUTE GET /api/mesures : Lecture temps réel depuis l'automate
// Appelée par la page Mesures d'Angular pour afficher les valeurs actuelles
// ============================================================
app.get('/api/mesures', (req, res) => {
  if (!client.Connected()) {
    return res.status(503).json({ message: '⚠️ Automate non connecté' });
  }

  const variables = [
    { Area: client.S7AreaDB, WordLen: client.S7WLReal, DBNumber: 3, Start: 0,   Amount: 1 }, // I1
    { Area: client.S7AreaDB, WordLen: client.S7WLReal, DBNumber: 3, Start: 4,   Amount: 1 }, // I2
    { Area: client.S7AreaDB, WordLen: client.S7WLReal, DBNumber: 3, Start: 8,   Amount: 1 }, // I3
    { Area: client.S7AreaDB, WordLen: client.S7WLReal, DBNumber: 3, Start: 56,  Amount: 1 }, // V1
    { Area: client.S7AreaDB, WordLen: client.S7WLReal, DBNumber: 3, Start: 60,  Amount: 1 }, // V2
    { Area: client.S7AreaDB, WordLen: client.S7WLReal, DBNumber: 3, Start: 64,  Amount: 1 }, // V3
    { Area: client.S7AreaDB, WordLen: client.S7WLReal, DBNumber: 3, Start: 108, Amount: 1 }, // P1
    { Area: client.S7AreaDB, WordLen: client.S7WLReal, DBNumber: 3, Start: 112, Amount: 1 }, // P2
    { Area: client.S7AreaDB, WordLen: client.S7WLReal, DBNumber: 3, Start: 116, Amount: 1 }, // P3
    { Area: client.S7AreaDB, WordLen: client.S7WLReal, DBNumber: 3, Start: 120, Amount: 1 }, // P_totale
    { Area: client.S7AreaDB, WordLen: client.S7WLReal, DBNumber: 3, Start: 168, Amount: 1 }, // FC_PUISSANCE
  ];

  try {
    lireVariablesEnBloc(variables, function (err, results) {
      if (err) {
        return res.status(500).json({ error: client.ErrorText(err) });
      }

      res.json({
        i1:           lireValeurFloat(results[0]),
        i2:           lireValeurFloat(results[1]),
        i3:           lireValeurFloat(results[2]),
        v1:           lireValeurFloat(results[3]),
        v2:           lireValeurFloat(results[4]),
        v3:           lireValeurFloat(results[5]),
        p1:           lireValeurFloat(results[6]) * 1000,
        p2:           lireValeurFloat(results[7]) * 1000,
        p3:           lireValeurFloat(results[8]) * 1000,
        p_totale:     lireValeurFloat(results[9]) * 1000,
        fc_puissance: lireValeurFloat(results[10]),
      });
    });
  } catch (error) {
    return res.status(500).json({ error: error.message });
  }
});

// ============================================================
// ROUTE GET /api/mesures/excel : Lecture des mesures depuis le fichier Excel
// Appelée par la page Historique d'Angular
// ============================================================
app.get('/api/mesures/excel', (req, res) => {
  try {
    if (!fs.existsSync(frontendExcelPath)) {
      return res.json([]);
    }

    const workbook = XLSX.readFile(frontendExcelPath);
    const sheetName = workbook.SheetNames[0] || 'Mesures';
    const sheet = workbook.Sheets[sheetName];
    const mesures = sheet ? XLSX.utils.sheet_to_json(sheet, { defval: '' }) : [];

    res.json(mesures);
  } catch (error) {
    console.error('❌ Erreur lecture fichier Excel :', error);
    res.status(500).json({ message: 'Erreur lors de la lecture du fichier Excel' });
  }
});

// ============================================================
// ROUTE POST /api/mesures/excel : Sauvegarde manuelle dans le fichier Excel
// Appelée par le bouton "Sauvegarder" côté Angular
// ============================================================
app.post('/api/mesures/excel', (req, res) => {
  try {
    const nouvellesMesures = Array.isArray(req.body) ? req.body : [req.body];
    sauvegarderMesuresDansFichier(nouvellesMesures);
    res.status(200).json({ message: '✅ Mesures sauvegardées avec succès' });
  } catch (error) {
    console.error('❌ Erreur lors de la sauvegarde manuelle :', error);
    res.status(500).json({ message: 'Erreur lors de la sauvegarde' });
  }
});

// ============================================================
// DÉMARRAGE DU SERVEUR
// ⚠️ On démarre "server" (et pas "app") car Socket.IO est attaché dessus
// ⚠️ '0.0.0.0' est indispensable pour que l'émulateur/téléphone puisse accéder à l'API
// ============================================================
const PORT = 3000;
server.listen(PORT, '0.0.0.0', () => {
  console.log(`✅ Serveur lancé sur http://0.0.0.0:${PORT}`);
  console.log(`📱 À utiliser depuis l'émulateur Android : http://10.0.2.2:${PORT}`);

  // Connexion à l'automate au démarrage du serveur
  connecterAutomate();

  // ⏱️ Lecture et sauvegarde automatique toutes les 5 secondes
  setInterval(() => {
    if (client.Connected()) {
      lireDonneesAutomate();
    } else {
      console.log('⚠️ Automate déconnecté, tentative de reconnexion...');
      connecterAutomate();
    }
  }, 5000);
});

// ============================================================
// SCHÉMA DE FONCTIONNEMENT GLOBAL :
//
// Automate S7-1214C (DB3)
//         ↕ (node-snap7, toutes les 5s)
// Backend Node.js (server.js)
//         ├── Socket.IO → événements 'mesures' et 'anomalie' en temps réel
//         ├── Firebase Cloud Messaging → notifications push (même app fermée)
//         ├── Sauvegarde dans mesures.xlsx (public/) UNIQUEMENT si anomalie détectée
//         ├── GET  /api/status         → LED verte/rouge (état connexion automate)
//         ├── GET  /api/mesures        → page Mesures (lecture temps réel via HTTP)
//         ├── GET  /api/mesures/excel  → page Historique (depuis Excel)
//         ├── POST /api/mesures/excel  → sauvegarde manuelle
//         ├── POST /api/register-token → enregistre le token FCM d'un appareil
//         └── GET  /api/test-anomalie  → route de TEST (à retirer après usage)
//
// Angular (my-app)
//         ├── /login      → authentification
//         ├── /mesures    → affichage temps réel via Socket.IO + GET /api/mesures
//         └── /historique → lecture Excel via GET /api/mesures/excel
// ============================================================