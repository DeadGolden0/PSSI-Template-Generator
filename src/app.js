const express = require('express');
const fs = require('fs');
const path = require('path');
const PizZip = require('pizzip');
const { PORT } = require('./config');
const Docxtemplater = require('docxtemplater');

const app = express();
app.use(express.urlencoded({ extended: true }));
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public'))); // Serve static files (index.html)

// Function to remove only the yellow highlighting in the XML content
function removeHighlighting(content) {
    // Remove only <w:shd> tags with yellow background color attribute without affecting other styles
    return content.replace(/<w:shd [^>]*w:fill="(?:yellow|ffe599|lightYellow)"[^>]*\/>/g, '');
}

// Function to process all XML files (body, headers, footers)
function processXmlFiles(zip) {
    const filesToModify = [
        'word/document.xml',
        'word/header1.xml',
        'word/header2.xml',
        'word/footer1.xml',
        'word/footer2.xml'
        // Add other files if necessary (header3.xml, footer3.xml, etc.)
    ];

    filesToModify.forEach((fileName) => {
        if (zip.file(fileName)) {
            let content = zip.file(fileName).asText();
            content = removeHighlighting(content);
            zip.file(fileName, content);
        }
    });
}

// Route to receive the form and generate the DOCX
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
        // Load the DOCX template
        const templatePath = path.resolve(__dirname, '../Template', 'PSSI_Template.docx');
        const content = fs.readFileSync(templatePath, 'binary');
        const zip = new PizZip(content);

        const doc = new Docxtemplater(zip, {
            paragraphLoop: true,
            linebreaks: true,
        });

        // Replace the tags with the form data
        doc.render(data);

        // Generate the DOCX file and get the content as ZIP
        const buf = doc.getZip().generate({ type: 'nodebuffer' });
        const outputZip = new PizZip(buf);

        // Remove the yellow highlighting in the relevant files
        processXmlFiles(outputZip);

        // Generate the updated file
        const updatedBuf = outputZip.generate({ type: 'nodebuffer' });

        // Set the response type for download
        res.setHeader('Content-Disposition', 'attachment; filename=output.docx');
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
        
        // Send the file to the client
        res.send(updatedBuf);
    } catch (error) {
        console.error('Error generating the DOCX file:', error);
        res.status(500).send('Error generating the DOCX file');
    }
});

// Start the server
app.listen(PORT, () => {
    console.log(`Server listening on http://localhost:${PORT}`);
});
