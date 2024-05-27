const express = require('express');
const bodyParser = require('body-parser');
const app = express();
const port = 3000;
const cors = require('cors'); // Import the cors package
const ExcelService = require('./fileservice');

//Array
let participants = [];


//CORS CONFIG
const corsOptions = {
    origin: ['http://localhost:4200','https://feed-back-form.vercel.app'], // Allow only this origin
    methods: 'GET,HEAD,POST', // Allow specific methods
    credentials: true, // Allow cookies to be sent
    optionsSuccessStatus: 204 // Some legacy browsers (IE11, various SmartTVs) choke on 204
};


app.use(bodyParser.json());
app.use(cors(corsOptions)); // Use the cors middleware
app.use(express.static('public'));


//Port config
app.listen(port, () => {
    console.log(`Server is up and running on http://localhost:${port}`);
});




//APIS


//1) base url to test if server is up
app.get('/', (req, res) => {
    res.send('Server is up');
})




//PostMapping Request Setup

//save an individual participant
app.post('/participants/save', async (req, res) => {
    const participant = req.body;

    try {
        await ExcelService.addData(participant, "adcomsys.xlsx", String(participant.EventName).toUpperCase()) // Adding data to sheet
        res.send({status:"success"});
    } catch (error) {
        console.error('Error:', error);
        res.status(500).send({status:"failure"});
    }
});

// Get All Participants
app.get('/participants', (req, res) => {
    res.send(participants);
});


//Function for verifying token recieved from client
const getTokenScore = (req, res) => {
    const secretKey = "6LeF2eMpAAAAAHwNPUZE3cl3BDPE9Uf396K0r4wu"
    const token = req.body.token;
    const url = `https://www.google.com/recaptcha/api/siteverify?secret=${secretKey}&response=${token}`;

    fetch(url, { method: 'post' })
        .then(response => response.json())
        .then(google_response => {
            res.send({ "score": google_response['score'] })
        })
        .catch(error => res.json({ error }));
}
//verify captcha token
app.post('/captcha/verify', getTokenScore)



//FILE SAVING APIS

//method to create an excel file
app.get('/create-excel', async (req, res) => {

    const filename = req.body.filename;
    const worksheets = req.body.worksheets;
    try {
        await ExcelService.createWorkBook(filename, worksheets)
        res.send('Excel file created successfully!');
    } catch (error) {
        console.error('Error:', error);
        res.status(500).send('Error creating Excel file');
    }
});

//Saving participant to worksheet
app.post('/add-data', async (req, res) => {
    try {
        await ExcelService.addData(req.body, "adcomsys.xlsx", String(req.body.EventName).toUpperCase()) // Adding data to sheet
        res.send("success");
    } catch (error) {
        console.error('Error:', error);
        res.status(500).send('Error adding data');
    }
});


app.get('/read-excel/:sheetname', async (req, res) => {

    const pathParam = req.params.sheetname;
    
    console.log();
    const filename = 'adcomsys.xlsx';
    const sheetname = String(pathParam).toUpperCase();

    try {
        const rows = await ExcelService.readExcelFile(filename, sheetname);
        res.json(rows);
    } catch (error) {
        console.error('Error:', error);
        res.status(500).send('Error reading Excel file');
    }
});
module.exports = app
