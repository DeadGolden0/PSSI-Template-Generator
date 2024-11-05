const express = require('express');
const fs = require('fs');
const path = require('path');
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');

const app = express();
app.use(express.urlencoded({ extended: true }));
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public'))); // Servir les fichiers statiques (index.html)

// Fonction pour supprimer uniquement le surlignage jaune dans le contenu XML
function removeHighlighting(content) {
    // Supprime uniquement les balises <w:shd> avec l'attribut de couleur de fond jaune sans toucher aux autres styles
    return content.replace(/<w:shd [^>]*w:fill="(?:yellow|ffe599|lightYellow)"[^>]*\/>/g, '');
}


// Fonction pour traiter tous les fichiers XML (corps, headers, footers)
function processXmlFiles(zip) {
    const filesToModify = [
        'word/document.xml',
        'word/header1.xml',
        'word/header2.xml',
        'word/footer1.xml',
        'word/footer2.xml'
        // Ajoutez d'autres fichiers si nécessaire (header3.xml, footer3.xml, etc.)
    ];

    filesToModify.forEach((fileName) => {
        if (zip.file(fileName)) {
            let content = zip.file(fileName).asText();
            content = removeHighlighting(content);
            zip.file(fileName, content);
        }
    });
}

// Route pour recevoir le formulaire et générer le DOCX
app.post('/submit', (req, res) => {
    const data = {
        'DATE': req.body.date,
        'NOM_ENTREPRISE': req.body.companyName,
        'SECTEUR_ACTIVITE': req.body.sector,
        'ANNEE_FONDATION': req.body.foundationYear,
        'NOMBRE_EMPLOYES': req.body.employeeCount,
        'NOMBRE_SITES': req.body.siteCount,
        'ZONE_GEOGRAPHIQUE': req.body.zone,
        'ADRESSE': req.body.address,
        'MAIL_DPO': req.body.dpoEmail,
        'SIREN': req.body.siren,
    };

    try {
        // Charger le modèle DOCX
        const templatePath = path.resolve(__dirname, 'Template', 'PSSI_Guardia.docx');
        const content = fs.readFileSync(templatePath, 'binary');
        const zip = new PizZip(content);

        const doc = new Docxtemplater(zip, {
            paragraphLoop: true,
            linebreaks: true,
        });

        // Remplacer les balises par les données du formulaire
        doc.render(data);

        // Générer le fichier DOCX et récupérer le contenu en tant que ZIP
        const buf = doc.getZip().generate({ type: 'nodebuffer' });
        const outputZip = new PizZip(buf);

        // Supprimer le surlignage jaune dans les fichiers pertinents
        processXmlFiles(outputZip);

        // Générer le fichier mis à jour
        const updatedBuf = outputZip.generate({ type: 'nodebuffer' });

        // Définir le type de réponse pour le téléchargement
        res.setHeader('Content-Disposition', 'attachment; filename=output.docx');
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
        
        // Envoyer le fichier au client
        res.send(updatedBuf);
    } catch (error) {
        console.error('Erreur lors de la génération du fichier DOCX :', error);
        res.status(500).send('Erreur lors de la génération du fichier DOCX');
    }
});

// Démarrer le serveur
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`Serveur en écoute sur http://localhost:${PORT}`);
});
