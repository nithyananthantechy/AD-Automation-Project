const express = require('express');
const bodyParser = require('body-parser');
const cors = require('cors');
const jwt = require('jsonwebtoken');
const bcrypt = require('bcryptjs');
const fs = require('fs');
const path = require('path');

const app = express();
const PORT = process.env.PORT || 3030;
const SECRET_KEY = 'desicrew-secret-key'; // Change this for production
const UPDATE_TOKEN = 'desicrew-update-token'; // Token for PowerShell script updates

app.use(cors());
app.use(bodyParser.json());
app.use(express.static(path.join(__dirname, 'public')));

// Root → always serve login page
app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'public', 'login.html'));
});

// Old /index.html → redirect to login
app.get('/index.html', (req, res) => {
    res.redirect('/login.html');
});

const DATA_FILE = path.join(__dirname, 'data.json');

// Initial data structure
const initialData = {
    system: {
        serverName: "AD-SERVER",
        uptime: "0d 0h 0m",
        lastUpdated: new Date().toISOString()
    },
    stats: {
        totalUsers: 0,
        enabledUsers: 0,
        disabledUsers: 0,
        lockedUsers: 0
    },
    services: {
        "AD Creation": { status: "unknown", lastRun: null },
        "Password Reset": { status: "unknown", lastRun: null },
        "Offboarding": { status: "unknown", lastRun: null },
        "Unlock": { status: "unknown", lastRun: null }
    },
    logs: []
};

// Ensure data file exists
if (!fs.existsSync(DATA_FILE)) {
    fs.writeFileSync(DATA_FILE, JSON.stringify(initialData, null, 2));
}

// Helper to read/write data
const getData = () => JSON.parse(fs.readFileSync(DATA_FILE, 'utf8'));
const saveData = (data) => fs.writeFileSync(DATA_FILE, JSON.stringify(data, null, 2));

// The login logic is now grouped under PUBLIC ENDPOINTS below

const verifyAdmin = (req, res, next) => {
    const token = req.headers['authorization']?.split(' ')[1];
    if (!token) return res.status(403).json({ message: 'No token provided' });

    jwt.verify(token, SECRET_KEY, (err, decoded) => {
        if (err || decoded.role !== 'admin') return res.status(401).json({ message: 'Unauthorized' });
        next();
    });
};

// --- PUBLIC ENDPOINTS ---
const ADMIN_HASH = bcrypt.hashSync('admin123', 10);

app.post('/api/login', (req, res) => {
    const { username, password } = req.body;
    if (username === 'admin' && bcrypt.compareSync(password, ADMIN_HASH)) {
        const token = jwt.sign({ role: 'admin' }, SECRET_KEY, { expiresIn: '1h' });
        return res.json({ token, role: 'admin' });
    }
    res.status(401).json({ message: 'Invalid credentials' });
});

// --- ADMIN ENDPOINTS (Protected) ---
app.get('/api/status', verifyAdmin, (req, res) => {
    res.json(getData());
});

// --- UPDATE ENDPOINT (Internal for PowerShell) ---
app.post('/api/update', (req, res) => {
    const token = req.headers['x-update-token'];
    if (token !== UPDATE_TOKEN) return res.status(403).json({ message: 'Forbidden' });

    const { type, payload } = req.body;
    const data = getData();

    if (type === 'stats') {
        data.stats = { ...data.stats, ...payload };
        data.system.lastUpdated = new Date().toISOString();
        if (payload.serverName) data.system.serverName = payload.serverName;
        if (payload.uptime) data.system.uptime = payload.uptime;
    } else if (type === 'log') {
        const logEntry = {
            id: Date.now(),
            time: new Date().toLocaleTimeString(),
            service: payload.service,
            status: payload.status,
            message: payload.message
        };
        data.logs.unshift(logEntry);
        if (data.logs.length > 50) data.logs.pop(); // Keep last 50 logs

        if (data.services[payload.service]) {
            data.services[payload.service].status = payload.status;
            data.services[payload.service].lastRun = new Date().toISOString();
        }
    }

    saveData(data);
    res.json({ message: 'Data updated successfully' });
});

// --- ADMIN ENDPOINTS ---
app.post('/api/create-user', verifyAdmin, (req, res) => {
    // In a real scenario, this would trigger the PowerShell script
    // For now, we'll just log it
    const data = getData();
    data.logs.unshift({
        id: Date.now(),
        time: new Date().toLocaleTimeString(),
        service: "Dashboard",
        status: "info",
        message: `Admin triggered user creation for ${req.body.username}`
    });
    saveData(data);
    res.json({ message: 'User creation triggered successfully' });
});

app.listen(PORT, () => {
    console.log(`Dashboard Server running at http://localhost:${PORT}`);
});
