const express = require('express');
const bodyParser = require('body-parser');

const app = express();
const port = 3000;
const cors = require('cors'); // Import the cors package

//CORS CONFIG
const corsOptions = {
    origin: 'http://localhost:4200', // Allow only this origin
    methods: 'GET,HEAD,POST', // Allow specific methods
    credentials: true, // Allow cookies to be sent
    optionsSuccessStatus: 204 // Some legacy browsers (IE11, various SmartTVs) choke on 204
};


app.use(bodyParser.json());
app.use(cors(corsOptions)); // Use the cors middleware
app.use(express.static('public'));

const getTokenScore = (req, res) => {
    const secretKey = '6LeF2eMpAAAAAHwNPUZE3cl3BDPE9Uf396K0r4wu';
    const token = req.body.token;
    const url = `https://www.google.com/recaptcha/api/siteverify?secret=${secretKey}&response=${token}`;

    fetch(url, { method: 'post' })
        .then(response => response.json())
        .then(google_response => {
            res.send( {"score":google_response['score']})
        })
        .catch(error => res.json({ error }));
}

app.get('/', (req, res) => {
    res.send('Server is up');
})

//Array
let participants = [];
let verifyToken;

//Port config
app.listen(port, () => {
    console.log(`Server is up and running on http://localhost:${port}`);
});

//PostMapping Request Setup


app.post('/participants/save', (req, res) => {
    const participant = req.body;
    participants.push(participant);
    res.status(201).send(participant);
});

// Get All Participants
app.get('/participants', (req, res) => {
    res.send(participants);
});

//verify captcha token
app.post('/captcha/verify', getTokenScore)