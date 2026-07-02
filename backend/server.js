const express = require('express'); // Importation du module Express pour créer le serveur
const cors = require('cors'); // Importation du module CORS pour gérer les requêtes cross-origin
const XLSX = require('xlsx'); // Importation du module XLSX pour manipuler les fichiers Excel
const fs = require('fs'); // Importation du module fs pour gérer les fichiers
const path = require('path'); // Importation du module path pour gérer les chemins de fichiers
const snap7 = require('node-snap7'); // Importation du module node-snap7 pour communiquer avec les automates Siemens S7

const app = express(); // Création d'une instance de l'application Express

// Définition du chemin vers le fichier Excel dans le dossier public du frontend
const frontendExcelPath = path.join(__dirname, '..', 'my-app', 'public', 'mesures.xlsx');

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
        // Affichage de l'erreur si la connexion échoue
        console.error('❌ Erreur de connexion à l\'automate :', client.ErrorText(err));
      } else {
        // Confirmation de la connexion réussie
        console.log('✅ Connexion réussie à l\'automate Siemens S7-1214C');
      }
    });
  } catch (error) {
    console.error('❌ Exception lors de la connexion à l\'automate :', error.message);
  }
}

function lireValeurFloat(result) {
  if (!result || !result.Data || !Buffer.isBuffer(result.Data)) {
    return null;
  }

  try {
    return result.Data.readFloatBE(0);
  } catch (error) {
    console.error('❌ Erreur de lecture du flottant depuis l\'automate :', error.message);
    return null;
  }
}

function lireVariablesEnBloc(variables, callback) {
  const fallbackResults = [];
  let pending = variables.length;

  if (pending === 0) {
    callback(null, fallbackResults);
    return;
  }

  variables.forEach(function(variable, index) {
    client.ReadArea(
      variable.Area || client.S7AreaDB,
      variable.DBNumber || 0,
      variable.Start,
      variable.Amount,
      variable.WordLen || client.S7WLReal,
      function(readErr, data) {
        fallbackResults[index] = {
          Result: readErr ? 1 : 0,
          Data: readErr ? null : data,
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
// FONCTION : Lecture des variables depuis le DB3 de l'automate
// Les offsets correspondent exactement à ton DB_Mesures_Elec [DB3]
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

  // Lecture simultanée de toutes les variables depuis l'automate
  try {
    lireVariablesEnBloc(variables, function(err, results) {
      if (err) {
        // Affichage de l'erreur si la lecture échoue
        console.error('❌ Erreur lecture variables :', client.ErrorText(err));
        return;
      }

      // Construction de l'objet mesure avec les valeurs lues
      const mesure = {
        date:         new Date().toISOString().slice(0, 16), // Date et heure actuelles au format 2025-06-29T08:00
        i1:           lireValeurFloat(results[0]),  // Courant phase 1 (A)
        i2:           lireValeurFloat(results[1]),  // Courant phase 2 (A)
        i3:           lireValeurFloat(results[2]),  // Courant phase 3 (A)
        v1:           lireValeurFloat(results[3]),  // Tension phase 1 (V)
        v2:           lireValeurFloat(results[4]),  // Tension phase 2 (V)
        v3:           lireValeurFloat(results[5]),  // Tension phase 3 (V)
        p1:           lireValeurFloat(results[6]),  // Puissance phase 1 (W)
        p2:           lireValeurFloat(results[7]),  // Puissance phase 2 (W)
        p3:           lireValeurFloat(results[8]),  // Puissance phase 3 (W)
        p_totale:     lireValeurFloat(results[9]),  // Puissance active totale (W)
        fc_puissance: lireValeurFloat(results[10]), // Facteur de puissance
      };

      console.log('📊 Mesure lue depuis DB3 :', mesure);

      // Sauvegarde automatique de la mesure dans le fichier Excel
      sauvegarderMesuresDansFichier([mesure]);
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
  let workbook; // Variable pour stocker le classeur Excel
  let mesuresExistantes = []; // Variable pour stocker les mesures déjà enregistrées

  // Vérification si le fichier Excel existe déjà
  if (fs.existsSync(frontendExcelPath)) {
    // Si oui → on lit les données existantes pour ne pas les écraser
    workbook = XLSX.readFile(frontendExcelPath);
    const sheetName = workbook.SheetNames[0] || 'Mesures'; // Nom de la première feuille
    const sheet = workbook.Sheets[sheetName]; // Récupération de la feuille
    mesuresExistantes = sheet ? XLSX.utils.sheet_to_json(sheet, { defval: '' }) : []; // Conversion en JSON
  } else {
    // Si non → on crée un nouveau classeur Excel vide
    workbook = XLSX.utils.book_new();
  }

  // Combinaison des mesures existantes avec les nouvelles mesures
  const toutesLesMesures = [...mesuresExistantes, ...nouvellesMesures];

  // Conversion de toutes les mesures en une feuille Excel
  const worksheet = XLSX.utils.json_to_sheet(toutesLesMesures);

  // Remplacement ou création de la feuille dans le classeur
  if (workbook.SheetNames && workbook.SheetNames.length > 0) {
    // Si une feuille existe déjà → on la remplace
    const sheetName = workbook.SheetNames[0];
    workbook.Sheets[sheetName] = worksheet;
  } else {
    // Sinon → on crée une nouvelle feuille appelée 'Mesures'
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Mesures');
  }

  // Création du répertoire si nécessaire (au cas où public/ n'existe pas)
  fs.mkdirSync(path.dirname(frontendExcelPath), { recursive: true });

  // Écriture du classeur Excel dans le fichier
  XLSX.writeFile(workbook, frontendExcelPath);
  console.log('✅ Mesure sauvegardée dans Excel à :', new Date().toISOString().slice(0, 16));
}

// ============================================================
// MIDDLEWARE : Configuration d'Express
// ============================================================
app.use(cors()); // Autorise les requêtes depuis Angular (localhost:4200)
app.use(express.json({ limit: '10mb' })); // Permet de lire le corps des requêtes JSON

// ============================================================
// ROUTE GET /api/mesures : Lecture temps réel depuis l'automate
// Appelée par la page Mesures d'Angular pour afficher les valeurs actuelles
// ============================================================
app.get('/api/mesures', (req, res) => {
  // Vérification que l'automate est bien connecté
  if (!client.Connected()) {
    return res.status(503).json({ message: '⚠️ Automate non connecté' });
  }

  // Lecture des variables en temps réel depuis DB3
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
    lireVariablesEnBloc(variables, function(err, results) {
      if (err) {
        return res.status(500).json({ error: client.ErrorText(err) });
      }

      // Retourne les valeurs au format JSON pour Angular
      res.json({
        i1:           lireValeurFloat(results[0]),
        i2:           lireValeurFloat(results[1]),
        i3:           lireValeurFloat(results[2]),
        v1:           lireValeurFloat(results[3]),
        v2:           lireValeurFloat(results[4]),
        v3:           lireValeurFloat(results[5]),
        p1:           lireValeurFloat(results[6]),
        p2:           lireValeurFloat(results[7]),
        p3:           lireValeurFloat(results[8]),
        p_totale:     lireValeurFloat(results[9]),
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
    // Vérification si le fichier Excel existe
    if (!fs.existsSync(frontendExcelPath)) {
      return res.json([]); // Si le fichier n'existe pas, renvoyer un tableau vide
    }

    const workbook = XLSX.readFile(frontendExcelPath); // Lecture du fichier Excel
    const sheetName = workbook.SheetNames[0] || 'Mesures'; // Nom de la première feuille
    const sheet = workbook.Sheets[sheetName]; // Récupération de la feuille
    const mesures = sheet ? XLSX.utils.sheet_to_json(sheet, { defval: '' }) : []; // Conversion en JSON

    res.json(mesures); // Envoi des mesures au format JSON
  } catch (error) {
    console.error('❌ Erreur lecture fichier Excel :', error);
    res.status(500).json({ message: 'Erreur lors de la lecture du fichier Excel' });
  }
});

// ============================================================
// ROUTE POST /api/mesures/excel : Sauvegarde manuelle depuis Angular
// Peut être appelée si tu veux sauvegarder depuis le frontend
// ============================================================
app.post('/api/mesures/excel', (req, res) => {
  try {
    const donneesRecues = req.body; // Récupération des données du corps de la requête
    // Vérification si les données sont un tableau, sinon les encapsuler dans un tableau
    const nouvellesMesures = Array.isArray(donneesRecues) ? donneesRecues : [donneesRecues];

    sauvegarderMesuresDansFichier(nouvellesMesures); // Sauvegarde dans le fichier Excel

    res.json({ message: '✅ Fichier Excel mis à jour avec succès' });
  } catch (error) {
    console.error('❌ Erreur mise à jour fichier Excel :', error);
    res.status(500).json({ message: 'Erreur lors de la mise à jour du fichier Excel' });
  }
});

// ============================================================
// DÉMARRAGE DU SERVEUR
// ============================================================
const PORT = 3000;
app.listen(PORT, () => {
  console.log(`✅ Serveur lancé sur http://localhost:${PORT}`);

  // Connexion à l'automate au démarrage du serveur
  connecterAutomate();

  // ⏱️ Lecture et sauvegarde automatique toutes les 5 secondes
  setInterval(() => {
    if (client.Connected()) {
      // Si l'automate est connecté → on lit les variables et on sauvegarde
      lireDonneesAutomate();
    } else {
      // Si l'automate est déconnecté → on tente de se reconnecter
      console.log('⚠️ Automate déconnecté, tentative de reconnexion...');
      connecterAutomate();
    }
  }, 5000); // 5000 ms = 5 secondes
});

// ============================================================
// SCHÉMA DE FONCTIONNEMENT GLOBAL :
//
// Automate S7-1214C (DB3)
//         ↕ (node-snap7, toutes les 5s)
// Backend Node.js (server.js)
//         ├── Sauvegarde dans mesures.xlsx (public/)
//         ├── GET /api/mesures       → page Mesures (temps réel)
//         ├── GET /api/mesures/excel → page Historique (depuis Excel)
//         └── POST /api/mesures/excel → sauvegarde manuelle
//
// Angular (my-app)
//         ├── /login      → authentification
//         ├── /mesures    → affichage temps réel via GET /api/mesures
//         └── /historique → lecture Excel via GET /api/mesures/excel
// ============================================================