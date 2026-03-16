const API_URL = '/api';

// --- STATE MANAGEMENT ---
let authData = {
    token: localStorage.getItem('token'),
    role: localStorage.getItem('role')
};

const updateUI = (data) => {
    // Basic Info
    document.getElementById('server-name').textContent = data.system.serverName;
    document.getElementById('uptime').textContent = `Uptime: ${data.system.uptime}`;
    document.getElementById('last-updated').textContent = `Updated: ${new Date(data.system.lastUpdated).toLocaleTimeString()}`;

    // Stats
    document.getElementById('total-users').textContent = data.stats.totalUsers;
    document.getElementById('enabled-users').textContent = data.stats.enabledUsers;
    document.getElementById('disabled-users').textContent = data.stats.disabledUsers;
    document.getElementById('locked-users').textContent = data.stats.lockedUsers;

    // Services List
    const servicesList = document.getElementById('services-list');
    servicesList.innerHTML = '';
    Object.entries(data.services).forEach(([name, info]) => {
        const item = document.createElement('div');
        item.className = 'service-item';
        const lastRunStr = info.lastRun ? new Date(info.lastRun).toLocaleString() : 'Never';
        item.innerHTML = `
            <div class="service-meta">
                <span>${name}</span>
                <span>Last run: ${lastRunStr}</span>
            </div>
            <span class="badge ${info.status === 'success' ? '' : info.status === 'error' ? 'danger' : 'warning'}">
                ${info.status.toUpperCase()}
            </span>
        `;
        servicesList.appendChild(item);
    });

    // Logs
    const logEntries = document.getElementById('log-entries');
    logEntries.innerHTML = '';
    data.logs.forEach(log => {
        const entry = document.createElement('div');
        entry.className = 'log-entry';
        entry.innerHTML = `
            <span class="log-time">${log.time}</span>
            <span class="status ${log.status}">${log.status.toUpperCase()}</span>
            <span class="log-msg">${log.message}</span>
        `;
        logEntries.appendChild(entry);
    });

    // Admin Panel Visibility
    const adminPanel = document.getElementById('admin-panel');
    const adminLoginBtn = document.getElementById('admin-login-btn');
    if (authData.token && authData.role === 'admin') {
        adminPanel.classList.remove('hidden');
        adminLoginBtn.classList.add('hidden');
    } else {
        adminPanel.classList.add('hidden');
        adminLoginBtn.classList.remove('hidden');
    }
};

const fetchData = async () => {
    try {
        const response = await fetch(`${API_URL}/status`);
        const data = await response.json();
        updateUI(data);
    } catch (err) {
        console.error('Failed to fetch dashboard status:', err);
    }
};

// --- AUTHENTICATION ---
const login = async (username, password) => {
    try {
        const response = await fetch(`${API_URL}/login`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ username, password })
        });

        if (response.ok) {
            const data = await response.json();
            authData = data;
            localStorage.setItem('token', data.token);
            localStorage.setItem('role', data.role);
            document.getElementById('login-modal').classList.add('hidden');
            fetchData();
        } else {
            alert('Invalid credentials');
        }
    } catch (err) {
        alert('Login failed. Server might be down.');
    }
};

const logout = () => {
    authData = { token: null, role: null };
    localStorage.removeItem('token');
    localStorage.removeItem('role');
    fetchData();
};

// --- EVENTS ---
document.getElementById('admin-login-btn').onclick = (e) => {
    e.preventDefault();
    document.getElementById('login-modal').classList.remove('hidden');
};

document.getElementById('close-modal').onclick = () => {
    document.getElementById('login-modal').classList.add('hidden');
};

document.getElementById('login-form').onsubmit = (e) => {
    e.preventDefault();
    const u = document.getElementById('username').value;
    const p = document.getElementById('password').value;
    login(u, p);
};

document.getElementById('logout-btn').onclick = logout;

document.getElementById('create-user-form').onsubmit = async (e) => {
    e.preventDefault();
    const username = document.getElementById('new-username').value;
    const name = document.getElementById('new-name').value;

    try {
        const response = await fetch(`${API_URL}/create-user`, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'Authorization': `Bearer ${authData.token}`
            },
            body: JSON.stringify({ username, name })
        });

        if (response.ok) {
            alert(`User creation triggered for ${username}`);
            document.getElementById('create-user-form').reset();
            fetchData();
        } else {
            alert('Unauthorized or failed to trigger');
        }
    } catch (err) {
        alert('Action failed');
    }
};

// --- INITIALIZATION ---
fetchData();
setInterval(fetchData, 5000); // Polling every 5 seconds
