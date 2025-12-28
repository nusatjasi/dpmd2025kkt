// API Wrapper untuk Google Apps Script
// Bergantung pada variabel global CONFIG dari config.js

const api = {
    // Helper untuk request
    async request(action, method = 'GET', payload = null) {
        if (CONFIG.API_URL.includes('YOUR_SCRIPT_ID_HERE')) {
            console.warn('API URL belum dikonfigurasi di js/config.js');
            throw new Error('Konfigurasi API URL belum diatur. Silakan deploy script dan update js/config.js.');
        }

        let url = `${CONFIG.API_URL}?action=${action}`;
        const options = {
            method: method,
            headers: {
                // 'Content-Type': 'application/json', // CORS issue fix usually requires text/plain for GAS
            }
        };

        if (method === 'POST' && payload) {
            // GAS `doPost` reads plain text body usually
            options.body = JSON.stringify({ action: action, payload: payload });
            // Gunakan no-cors jika perlu, tapi kita butuh response, jadi 'cors' standard
            // GAS Web App harus return JSONP atau CORS headers yang benar.
            // ContentService di GAS sudah setMimeType(JSON), biasanya cukup.
            // Namun, untuk menghindari preflight options request yang sering gagal di GAS,
            // kita kirim sebagai text/plain.
        }

        // Handling special case for file upload wrapper requiring distinct structuring
        if (action === 'uploadFile') {
            options.body = JSON.stringify(payload); // Payload structure: { action: 'uploadFile', fileData: ..., fileName: ..., mimeType: ... }
        }

        try {
            const response = await fetch(url, options);
            if (!response.ok) throw new Error(`HTTP Error: ${response.status}`);
            const data = await response.json();
            return data;
        } catch (error) {
            console.error(`API Error (${action}):`, error);
            throw error;
        }
    },

    // --- Pegawai ---
    async getPegawai() {
        return this.request('getPegawai');
    },
    async savePegawai(pegawaiData) {
        return this.request('savePegawai', 'POST', pegawaiData);
    },
    async deletePegawai(id) {
        // Wrapper payload to match generic structure if needed, or send ID as payload
        // Backend expects { action: 'deletePegawai', id: ... } in POST body
        // Helper `request` wraps `payload` in `payload` key by default for "POST". 
        // We need to override or adjust backend to look for `payload.id`.
        // Let's adjust front-end to match backend `doPost`:
        // Backend `doPost`: const data = JSON.parse(e.postData.contents); const action = data.action; ... return deletePegawai(data.id);
        // So we need to send { action: 'deletePegawai', id: id } directly.

        let url = CONFIG.API_URL; // POST to base URL
        const body = JSON.stringify({ action: 'deletePegawai', id: id });
        const response = await fetch(url, { method: 'POST', body: body });
        return await response.json();
    },

    // --- Surat Masuk ---
    async getSuratMasuk() {
        return this.request('getSuratMasuk');
    },
    async saveSuratMasuk(suratData) {
        return this.request('saveSuratMasuk', 'POST', suratData);
    },
    async deleteSuratMasuk(id) {
        let url = CONFIG.API_URL;
        const body = JSON.stringify({ action: 'deleteSuratMasuk', id: id });
        const response = await fetch(url, { method: 'POST', body: body });
        return await response.json();
    },

    // --- Data Master ---
    async getMasterData() {
        return this.request('getMasterData');
    },
    async saveMasterData(masterData) {
        return this.request('saveMasterData', 'POST', masterData);
    },

    // --- Dashboard ---
    async getDashboardStats() {
        return this.request('getDashboardStats');
    },

    // --- Settings & Absensi ---
    async getAppSettings() {
        return this.request('getAppSettings');
    },
    async saveAbsensi(absensiData) {
        return this.request('saveAbsensi', 'POST', absensiData);
    },
    async getAbsensi(date) {
        let url = `${CONFIG.API_URL}?action=getAbsensi&date=${date}`;
        const response = await fetch(url);
        return await response.json();
    },
    async updateAbsensiStatus(payload) {
        return this.request('updateAbsensiStatus', 'POST', payload);
    },
    async saveAppSettings(settings) {
        return this.request('saveAppSettings', 'POST', settings);
    },

    // --- File Upload ---
    async uploadFile(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.readAsDataURL(file);
            reader.onload = async () => {
                const base64Data = reader.result.split(',')[1];
                const payload = {
                    action: 'uploadFile',
                    fileName: file.name,
                    mimeType: file.type,
                    fileData: base64Data
                };

                try {
                    let url = CONFIG.API_URL;
                    const response = await fetch(url, { method: 'POST', body: JSON.stringify(payload) });
                    const result = await response.json();
                    if (result.status === 'success') {
                        resolve(result);
                    } else {
                        reject(result.message);
                    }
                } catch (e) {
                    reject(e);
                }
            };
            reader.onerror = error => reject(error);
        });
    }
};
