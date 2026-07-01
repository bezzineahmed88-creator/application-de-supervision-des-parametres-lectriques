const express = require('express');// Importation du module Express pour créer le serveur
const cors = require('cors');// Importation du module CORS pour gérer les requêtes cross-origin
const XLSX = require('xlsx');// Importation du module XLSX pour manipuler les fichiers Excel
const fs = require('fs');// Importation du module fs pour gérer les fichiers
const path = require('path');// Importation du module path pour gérer les chemins de fichiers

const app = express();// Création d'une instance de l'application Express
// Définition du chemin vers le fichier Excel dans le dossier public du frontend
const frontendExcelPath = path.join(__dirname, '..', 'my-app', 'public', 'mesures.xlsx');

// Fonction pour sauvegarder les mesures dans un fichier Excel
function sauvegarderMesuresDansFichier(nouvellesMesures) {
  let workbook;// Variable pour stocker le classeur Excel
  let mesuresExistantes = [];// Variable pour stocker les mesures existantes

  // Vérification si le fichier Excel existe déjà
  if (fs.existsSync(frontendExcelPath)) {
    workbook = XLSX.readFile(frontendExcelPath);// Lecture du fichier Excel existant
    const sheetName = workbook.SheetNames[0] || 'Mesures';//  Récupération du nom de la première feuille ou utilisation de 'Mesures' par défaut
    const sheet = workbook.Sheets[sheetName];// Récupération de la feuille Excel
    mesuresExistantes = sheet ? XLSX.utils.sheet_to_json(sheet, { defval: '' }) : [];// Conversion de la feuille Excel en JSON, avec des valeurs par défaut vides pour les cellules vides
  } else {// Si le fichier Excel n'existe pas, création d'un nouveau classeur
    workbook = XLSX.utils.book_new();
  }

  const toutesLesMesures = [...mesuresExistantes, ...nouvellesMesures];// Combinaison des mesures existantes avec les nouvelles mesures
  const worksheet = XLSX.utils.json_to_sheet(toutesLesMesures);// Conversion des mesures combinées en une feuille Excel

  // Vérification si le classeur contient déjà des feuilles
  if (workbook.SheetNames && workbook.SheetNames.length > 0) {
    const sheetName = workbook.SheetNames[0];// Récupération du nom de la première feuille
    workbook.Sheets[sheetName] = worksheet;// Remplacement de la première feuille par la nouvelle feuille contenant toutes les mesures
  } else {
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Mesures');// Ajout de la nouvelle feuille au classeur avec le nom 'Mesures'
  }

  fs.mkdirSync(path.dirname(frontendExcelPath), { recursive: true });// Création du répertoire si nécessaire
  XLSX.writeFile(workbook, frontendExcelPath);// Écriture du classeur Excel dans le fichier
}

app.use(cors());
app.use(express.json({ limit: '10mb' }));

app.get('/api/mesures', (req, res) => {
  const i1 = 1111111111;
  res.json({ i1 });
});

app.get('/api/mesures/excel', (req, res) => {
  try {
    if (!fs.existsSync(frontendExcelPath)) {// Vérification si le fichier Excel existe
      return res.json([]);// Si le fichier n'existe pas, renvoyer un tableau vide
    }

    const workbook = XLSX.readFile(frontendExcelPath);//  Lecture du fichier Excel
    const sheetName = workbook.SheetNames[0] || 'Mesures';// Récupération du nom de la première feuille ou utilisation de 'Mesures' par défaut
    const sheet = workbook.Sheets[sheetName];// Récupération de la feuille Excel
    const mesures = sheet ? XLSX.utils.sheet_to_json(sheet, { defval: '' }) : [];// Conversion de la feuille Excel en JSON, avec des valeurs par défaut vides pour les cellules vides

    res.json(mesures);// Envoi des mesures au format JSON en réponse à la requête
  } catch (error) {
    console.error('Erreur lors de la lecture du fichier Excel :', error);
    res.status(500).json({ message: 'Erreur lors de la lecture du fichier Excel' });
  }
});

app.post('/api/mesures/excel', (req, res) => {
  try {
    const donneesRecu = req.body;// Récupération des données envoyées dans le corps de la requête
    const nouvellesMesures = Array.isArray(donneesRecu) ? donneesRecu : [donneesRecu];// Vérification si les données reçues sont un tableau, sinon les encapsuler dans un tableau

    sauvegarderMesuresDansFichier(nouvellesMesures);// Appel de la fonction pour sauvegarder les mesures dans le fichier Excel

    res.json({ message: 'Fichier Excel mis à jour avec succès' });
  } catch (error) {
    console.error('Erreur lors de la mise à jour du fichier Excel :', error);
    res.status(500).json({ message: 'Erreur lors de la mise à jour du fichier Excel' });
  }
});

// Démarrage du serveur
const PORT = 3000;
app.listen(PORT, () => {
  console.log(`Serveur lancé sur http://localhost:${PORT}`);
});
// Schéma de fonctionnement :
// 1. Le frontend envoie des mesures au backend via POST /api/mesures/excel.
// 2. Le backend vérifie si le fichier Excel existe déjà.
// 3. Si le fichier existe, il lit les mesures déjà enregistrées.
// 4. Si le fichier n'existe pas, il crée un nouveau classeur Excel vide.
// 5. Le backend ajoute les nouvelles mesures à la feuille Excel.
// 6. Le fichier Excel est enregistré dans le frontend (public/mesures.xlsx).
// 7. Le frontend peut ensuite lire ces données via GET /api/mesures/excel.