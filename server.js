const express = require('express')
const app = express()
const port = 3000
var bodyParser = require('body-parser')

app.use(bodyParser.urlencoded({ extended: false }))
app.use(bodyParser.json())

const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');

let fs = require('fs');
let path = require('path');

app.get('/', async (req, res) => {
    //Load the docx file as a binary
    var resume = fs.readFileSync(path.resolve(__dirname, 'resume.docx'), 'binary');
    var coverLetter = fs.readFileSync(path.resolve(__dirname, 'coverletter.docx'), 'binary');
    var resumeZip = new PizZip(resume);
    var coverLetterZip = new PizZip(coverLetter);
    var resumeDoc = new Docxtemplater();
    var coverDoc = new Docxtemplater();

    resumeDoc.loadZip(resumeZip);
    coverDoc.loadZip(coverLetterZip);
    //set the templateVariables

    let companyName = req.body.companyName
    let jobTitle = req.body.jobTitle

    resumeDoc.setData({
        company_name: companyName,
        job_title: jobTitle
    });
    coverDoc.setData({
        company_name: companyName,
        job_title: jobTitle
    });

    try {
        // render the document (replace all occurences of {first_name} by John, {last_name} by Doe, ...)
        resumeDoc.render()
        coverDoc.render()

    }
    catch (error) {
        var e = {
            message: error.message,
            name: error.name,
            stack: error.stack,
            properties: error.properties,
        }
        console.log(JSON.stringify({ error: e }));
        // The error thrown here contains additional information when logged with JSON.stringify (it contains a property object).
        throw error;
    }

    var resumeBuf = resumeDoc.getZip().generate({ type: 'nodebuffer' });
    var coverBuf = coverDoc.getZip().generate({ type: 'nodebuffer' });
    // buf is a nodejs buffer, you can either write it to a file or do anything else with it.
    fs.writeFileSync(path.resolve(__dirname, `Nick MacKenzie - Resume - ${companyName}.docx`), resumeBuf);
    fs.writeFileSync(path.resolve(__dirname, `Nick MacKenzie - Cover Letter - ${companyName}.docx`), coverBuf);
    res.send("done")
})

app.listen(port, () => {
    console.log(`Example app listening on port ${port}`)
})