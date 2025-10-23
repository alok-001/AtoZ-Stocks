const fetch = require('node-fetch');
const express = require('express');
const multer = require('multer');
const bodyParser = require('body-parser');
const path = require('path');
const pdf = require('pdf-parse');
const axios = require('axios');
const say = require('say');
const fs = require('fs');
const gtts = require('gtts.js').gTTS;
const SummarizerManager = require("node-summarizer").SummarizerManager;
require('dotenv').config();
const pptxgen = require('pptxgenjs');

const app = express();
const port = 3000;

async function localSummarize(content) {
    const summarizerManager = new SummarizerManager(content, 20); // Adjust the number of sentences accordingly
    const summary = await summarizerManager.getSummaryByRank();
    return summary;
}


// Set up static and views directories
app.use('/static', express.static(path.join(__dirname, 'static')));
app.use('/downloads', express.static(path.join(__dirname))); // to serve the audio file

app.set('views', path.join(__dirname, 'templates'));
app.set('view engine', 'ejs');

// Use body-parser for parsing POST data
app.use(bodyParser.urlencoded({ extended: true }));

async function convertTextToSpeech(text) {
    return new Promise((resolve, reject) => {
        const outputFile = './audio/output.mp3';

        say.export(text, null, 1, outputFile, function (err) {
            if (err) {
                reject(err);
            } else {
                console.log(`Audio content written to file: ${outputFile}`);
                resolve(outputFile);
            }
        });
    });
}


const MAX_RETRIES = 3; // Maximum number of retries
const RETRY_DELAY = 5000; // Delay between retries in milliseconds (5 seconds)

async function chatWithGPT(apiKey, userMessage, retries = 0) {
    const url = "https://api.openai.com/v1/chat/completions";
    const headers = {
        "Content-Type": "application/json",
        "Authorization": `Bearer ${apiKey}`
    };

    const data = {
        model: "gpt-3.5-turbo",
        messages: [
            {
                role: "system",
                content: "You are a helpful assistant."
            },
            {
                role: "user",
                content: userMessage
            }
        ]
    };

    try {
        const response = await axios.post(url, data, { headers: headers });
        return response.data.choices[0].message.content;
    } catch (error) {
        console.error("There was a problem with the fetch operation:", error.message);
        if (error.response && error.response.status === 429 && retries < MAX_RETRIES) {
            console.log(`Rate limit hit. Retrying in ${RETRY_DELAY / 1000} seconds...`);
            await new Promise(resolve => setTimeout(resolve, RETRY_DELAY));
            return chatWithGPT(apiKey, userMessage, retries + 1);
        } else {
            throw new Error("Max retries reached or other error encountered. Please try again later.");
        }
    }
}

// Configure multer for file uploads
const storage = multer.diskStorage({
    destination: (req, file, cb) => {
        cb(null, './uploads');
    },
    filename: (req, file, cb) => {
        cb(null, file.originalname);
    }
});

const upload = multer({ storage: storage });

async function extractTextFromPDF(filePath) {
    const data = await pdf(filePath);
    return data.text;
}

// Routes
app.get('/', (req, res) => {
    res.render('index');
});
app.get('/audio/:filename', (req, res) => {
    const audioPath = path.join(__dirname, 'audio', req.params.filename);
    res.setHeader('Content-Type', 'audio/mpeg');
    res.sendFile(audioPath);
});
app.get('/ppt/:filename', (req, res) => {
    const pptPath = path.join(__dirname, 'ppt', req.params.filename);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.presentationml.presentation');
    res.sendFile(pptPath);
});


function generatePPTFromSummary(summary) {
    const pptx = new pptxgen();

    // Set the default layout for slides
    pptx.layout = 'LAYOUT_WIDE';

    // Set a black theme with white text
    pptx.defineSlideMaster({
        title: 'MASTER_SLIDE',
        background: { color: '000000' },
        color: 'FFFFFF',
        objects: [
            { 'placeholder': { options: { name: 'body', type: 'body', x: 0.5, y: 1.0, w: 8, h: 5.5, align: 'center', font_size: 20, color: 'FFFFFF' } } },
            { 'placeholder': { options: { name: 'title', type: 'title', x: 0.5, y: 0.5, w: 8, h: 1, align: 'center', font_size: 30, color: 'FFFFFF' } } }
        ]
    });

    // Split the summary into sentences
    const sentences = summary.split(/\.\s+/);

    sentences.forEach(sentence => {
        // Add random colored words
        const words = sentence.split(/\s+/);
        const coloredText = [];

        words.forEach((word, idx) => {
            if (Math.random() < 0.1) { // 10% chance to color a word
                coloredText.push({ text: word, options: { color: randomColor(), fontSize: 20 } });
            } else {
                coloredText.push({ text: word, options: { color: 'FFFFFF', fontSize: 20 } });
            }
            if (idx !== words.length - 1) {
                coloredText.push({ text: ' ', options: { color: 'FFFFFF', fontSize: 20 } }); // Add space after each word except the last one
            }
        });

        const slide = pptx.addSlide('MASTER_SLIDE');
        slide.title = 'Financial Summary';
        slide.addText(coloredText, { x: 1.5, y: 1, w: 7, h: 5.5, align: 'center', breakLine: true });
    });

    // Add the Thank You slide
    const thankYouSlide = pptx.addSlide('MASTER_SLIDE');
    thankYouSlide.addText('Thank You!\nDeepak Joshi', { x: 1.5, y: 2.5, w: 7, h: 3, align: 'center', fontSize: 36, color: 'FFFFFF' });

    const outputFile = path.join(__dirname, 'ppt', 'summary.pptx');
    pptx.writeFile(outputFile)
        .then(() => {
            console.log('PPT generated successfully!');
        })
        .catch(err => {
            console.error('Error generating PPT:', err);
        });

    return outputFile;
}

// Utility function to generate a random color
function randomColor() {
    const letters = '0123456789ABCDEF';
    let color = '';
    for (let i = 0; i < 6; i++) {
        color += letters[Math.floor(Math.random() * 16)];
    }
    return color;
}
















app.post('/upload', upload.single('file'), async (req, res) => {
    const filePath = path.join(__dirname, req.file.path);
    const extractedText = await extractTextFromPDF(filePath);

    // Extract key financial metrics from the content
    const financialData = {
        revenue: extractedText.match(/Revenue From Operations\n([\d,.]+)/i)?.[1],
        netProfitBeforeTax: extractedText.match(/Net Profit for the period \(before Tax and Exceptional items\)\n([\d,.]+)/i)?.[1],
        netProfitAfterTax: extractedText.match(/Net Profit for the period after Tax \(after Exceptional items\)\n([\d,.]+)/i)?.[1],
        totalComprehensiveIncome: extractedText.match(/Total Comprehensive Income for the period \[.*\]\n([\d,.]+)/i)?.[1],
        equityShareCapital: extractedText.match(/Paid up Equity Share Capital.*\n([\d,.]+)/i)?.[1],
        netWorth: extractedText.match(/Net Worth\n([\d,.]+)/i)?.[1],
        outstandingDebt: extractedText.match(/Paid up Debt Capital\/Outstanding Debt\n([\d,.]+)/i)?.[1],
        eps: extractedText.match(/Earning Per Share \(of Rs.10 each\)\nBasic \(Rs.\)\n([\d,.]+)/i)?.[1]
    };

    // Formulate a structured prompt for GPT-3 to summarize the extracted metrics
    const promptParts = [
        financialData.revenue && `- The Revenue from Operations is ${financialData.revenue}`,
        financialData.netProfitBeforeTax && `- The Net Profit before Tax is ${financialData.netProfitBeforeTax}`,
        financialData.netProfitAfterTax && `- The Net Profit after Tax is ${financialData.netProfitAfterTax}`,
        financialData.totalComprehensiveIncome && `- The Total Comprehensive Income is ${financialData.totalComprehensiveIncome}`,
        financialData.equityShareCapital && `- The Paid-up Equity Share Capital is ${financialData.equityShareCapital}`,
        financialData.netWorth && `- The Net Worth of the company is ${financialData.netWorth}`,
        financialData.outstandingDebt && `- The Outstanding Debt is ${financialData.outstandingDebt}`,
        financialData.eps && `- The Earnings Per Share is ${financialData.eps}`
    ].filter(Boolean);  // Filters out undefined or null values

    const prompt = `
        You are a Financial Analyst and Investor Relations Officer Based on this data :
        ${promptParts.join('\n')}
        Please provide me a concise summary for stakeholders and your suggestion also whether the investor should invest in this or not. 
    `;

    const apiKey = process.env.API_KEY;

    try {
        const summarizedResult = await chatWithGPT(apiKey, prompt);

        // Convert the summarized result into audio and provide a download link
        const audioFilePath = await convertTextToSpeech(summarizedResult);

        // Generate a PPT from the summarized result
        const pptFilePath = generatePPTFromSummary(summarizedResult);

        console.log("Final Summarized Result:", summarizedResult);
        res.setHeader('Content-Type', 'text/html');
        res.send(`
            <a href="/audio/${path.basename(audioFilePath)}" download>Click here to download the audio summary</a><br>
            <a href="/ppt/${path.basename(pptFilePath)}" download>Click here to download the PPT summary</a>
        `);

    } catch (error) {
        console.error(error.message);
        res.status(503).send('Error: Unable to process at this time. Please try again later.');
    }
});





app.post('/feedback', (req, res) => {
    const feedback = req.body.feedback;
    console.log(feedback);
    res.send('Thank you for your feedback!');
});

app.listen(port, () => {
    console.log(`Server started on http://localhost:${port}`);
});
