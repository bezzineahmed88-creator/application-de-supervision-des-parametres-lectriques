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
app.use(cors());
app.use(express.json({ limit: '10mb' }));

// Chemin vers le fichier Excel dans le dossier public du frontend
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
const IP_AUTOMATE = '192.168.1.40';
const RACK = 0;
const SLOT = 1;

const client = new snap7.S7Client();

// ✅ Taille totale du DB3 à lire en une seule fois (dernier offset 168 + 4 octets)
const TAILLE_DB3 = 172;

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
// ROUTE GET /api/status
// ============================================================
app.get(['/api/status', '/api/statut'], (req, res) => {
  res.json({ connecte: client.Connected() });
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
// FONCTION : Découpe le buffer complet du DB3 en valeurs individuelles
// ✅ Une seule lecture réseau → plus de mélange entre variables
// ============================================================
function decouperMesureDepuisBuffer(data) {
  return {
    date:         new Date().toISOString().slice(0, 16),
    i1:           formaterReel(data.readFloatBE(0)),
    i2:           formaterReel(data.readFloatBE(4)),
    i3:           formaterReel(data.readFloatBE(8)),
    v1:           formaterReel(data.readFloatBE(56)),
    v2:           formaterReel(data.readFloatBE(60)),
    v3:           formaterReel(data.readFloatBE(64)),
    p1:           formaterReel(data.readFloatBE(108) * 1000),
    p2:           formaterReel(data.readFloatBE(112) * 1000),
    p3:           formaterReel(data.readFloatBE(116) * 1000),
    p_totale:     formaterReel(data.readFloatBE(120) * 1000),
    fc_puissance: formaterReel(data.readFloatBE(168)),
  };
}

// ============================================================
// Seuils de détection d'anomalie
// ============================================================
const seuils = {
  tension:   { min: 200, max: 350 },
  courant:   { min: 3,  max: 10 },
  puissance: { min: -600,   max: 150 },
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
// ✅ Une seule requête ReadArea sur tout le DB3
// ============================================================
function lireDonneesAutomate() {
  if (!client.Connected()) {
    console.log('⚠️ Automate non connecté, lecture annulée');
    connecterAutomate();
    return;
  }

  try {
    client.ReadArea(client.S7AreaDB, 3, 0, TAILLE_DB3, client.S7WLByte, function (err, data) {
      if (err) {
        console.error('❌ Erreur lecture DB3 :', client.ErrorText(err));
        return;
      }

      const mesure = decouperMesureDepuisBuffer(data);
      console.log('📊 Mesure lue depuis DB3 :', mesure);

      io.emit('mesures', mesure);

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
// ⚠️ À retirer une fois les tests terminés
// ============================================================
app.get('/api/test-anomalie', (req, res) => {
  const mesureTest = {
    date: new Date().toISOString().slice(0, 16),
    i1: 0, i2: 15, i3: 15,
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
// ✅ Une seule lecture en bloc
// ============================================================
app.get('/api/mesures', (req, res) => {
  if (!client.Connected()) {
    return res.status(503).json({ message: '⚠️ Automate non connecté' });
  }

  try {
    client.ReadArea(client.S7AreaDB, 3, 0, TAILLE_DB3, client.S7WLByte, function (err, data) {
      if (err) {
        return res.status(500).json({ error: client.ErrorText(err) });
      }
      res.json(decouperMesureDepuisBuffer(data));
    });
  } catch (error) {
    return res.status(500).json({ error: error.message });
  }
});

// ============================================================
// ROUTE GET /api/mesures/excel
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
// ROUTE POST /api/mesures/excel
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
// ============================================================
const PORT = 3000;
server.listen(PORT, '0.0.0.0', () => {
  console.log(`✅ Serveur lancé sur http://0.0.0.0:${PORT}`);
  console.log(`📱 À utiliser depuis l'émulateur Android : http://10.0.2.2:${PORT}`);

  connecterAutomate();

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