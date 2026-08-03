const express = require('express');
const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');
const { Pool } = require('pg');

const app = express();
const PORT = process.env.PORT || 3000;

app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.use(express.static(path.join(__dirname, 'public')));

const pool = new Pool({
    connectionString: process.env.DATABASE_URL,
    ssl: {
        rejectUnauthorized: false
    }
});

const TARGET_DIR = path.join(__dirname, 'public');
const JSON_FILE = path.join(TARGET_DIR, 'orders.json');
const EXCEL_FILE = path.join(TARGET_DIR, '訂單統計表.xlsx');
const USER_DB_FILE = path.join(TARGET_DIR, 'users.json');

const INITIAL_USER_DB = {
    '18445': '李珈豪', '601471': '陳永育', '11110': '林佳蘭', '11744': '施名娟',
    '10069': '許民芳', '13228': '宋筱湄', '12218': '沈佩琪', '10047': '許博捷',
    '6513': '李承州', '16661': '陳育倫', '601473': '李羽茹', '6800': '吳修文',
    '17020': '黃泓耀', '14778': '施憶宣', '13266': '梁薽予', '5514': '張淑娟',
    '9844': '許毓芬', '16696': '賴語婕', '16984': '梁慧如', '15150': '黃珮瑄',
    '16294': '梁婧盈', '16925': '李宜珊', '17528': '曾雅琴', 'TEST': '測試員'
};

async function ensureDirectoryExistence() {
    if (!fs.existsSync(TARGET_DIR)) {
        fs.mkdirSync(TARGET_DIR, { recursive: true });
        console.log(`[系統提示] 已自動建立目標資料夾：${TARGET_DIR}`);
    }
    if (!fs.existsSync(USER_DB_FILE)) {
        fs.writeFileSync(USER_DB_FILE, JSON.stringify(INITIAL_USER_DB, null, 2), 'utf-8');
        console.log(`[系統提示] 已自動建立員工資料庫檔案：${USER_DB_FILE}`);
    }
    if (!fs.existsSync(JSON_FILE)) {
        fs.writeFileSync(JSON_FILE, JSON.stringify([], null, 2), 'utf-8');
        console.log(`[系統提示] 已自動建立訂單檔案：${JSON_FILE}`);
    }

    try {
        const createTableQuery = `
            CREATE TABLE IF NOT EXISTS orders (
                id SERIAL PRIMARY KEY,
                order_id BIGINT UNIQUE NOT NULL,
                card_id VARCHAR(50) NOT NULL,
                name VARCHAR(100) NOT NULL,
                meal TEXT NOT NULL,
                spicy VARCHAR(50),
                note TEXT,
                total NUMERIC(10, 2),
                timestamp VARCHAR(100),
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            );
        `;
        await pool.query(createTableQuery);
        console.log(`[系統提示] PostgreSQL 訂單資料表檢查/建立成功！`);
    } catch (err) {
        console.error(`❌ 建立 PostgreSQL 資料表失敗：`, err.message);
    }
}

ensureDirectoryExistence();

function safeReadJSON(filePath, fallback = []) {
    try {
        const raw = fs.readFileSync(filePath, 'utf-8').trim();
        return raw ? JSON.parse(raw) : fallback;
    } catch (e) {
        console.error(`[檔案讀取錯誤] 解析 ${filePath} 失敗:`, e.message);
        return fallback;
    }
}

function loadUserDatabase() {
    const data = safeReadJSON(USER_DB_FILE, INITIAL_USER_DB);
    return (data && typeof data === 'object' && !Array.isArray(data)) ? data : INITIAL_USER_DB;
}

function saveUserDatabase(db) {
    fs.writeFileSync(USER_DB_FILE, JSON.stringify(db, null, 2), 'utf-8');
}

function isTodayInTaipei(orderTimestampMs) {
    const orderDate = new Date(Number(orderTimestampMs));
    if (isNaN(orderDate.getTime())) return false;

    const now = new Date();
    const options = { timeZone: 'Asia/Taipei', year: 'numeric', month: '2-digit', day: '2-digit' };
    
    const orderDateStr = new Intl.DateTimeFormat('zh-TW', options).format(orderDate);
    const todayDateStr = new Intl.DateTimeFormat('zh-TW', options).format(now);

    return orderDateStr === todayDateStr;
}

// 輔助函式：同步更新本機 orders.json 與 Excel
async function syncLocalFilesAndExcel() {
    try {
        const result = await pool.query('SELECT * FROM orders ORDER BY order_id ASC;');
        const orders = result.rows.map(row => ({
            orderId: Number(row.order_id),
            timestamp: row.timestamp,
            cardId: row.card_id,
            name: row.name,
            meal: row.meal,
            spicy: row.spicy,
            note: row.note,
            total: Number(row.total)
        }));

        // 1. 同步寫入本機 orders.json
        fs.writeFileSync(JSON_FILE, JSON.stringify(orders, null, 2), 'utf-8');

        // 2. 同步更新 Excel
        const workbook = new ExcelJS.Workbook();
        const sheet1 = workbook.addWorksheet('訂單明細');
        sheet1.columns = [
            { header: '訂單ID', key: 'orderId', width: 18 },
            { header: '時間', key: 'timestamp', width: 25 },
            { header: '員工卡號', key: 'cardId', width: 15 },
            { header: '員工姓名', key: 'name', width: 15 }, 
            { header: '點餐店家/品項', key: 'meal', width: 25 },
            { header: '醬料辣度', key: 'spicy', width: 15 },
            { header: '備註', key: 'note', width: 25 },
            { header: '金額', key: 'total', width: 12 }
        ];
        
        const sheet2 = workbook.addWorksheet('店家點餐統計表');
        sheet2.columns = [
            { header: '店家/品項名稱', key: 'meal', width: 25 },
            { header: '總點餐次數', key: 'count', width: 15 }
        ];

        orders.forEach(order => {
            sheet1.addRow({
                orderId: order.orderId,
                timestamp: order.timestamp,
                cardId: order.cardId,
                name: order.name || '未知', 
                meal: order.meal,
                spicy: order.spicy || '無',
                note: order.note || '無',
                total: !isNaN(order.total) ? order.total : 0
            });
        });

        const statistics = {};
        orders.forEach(order => {
            if (order.meal) {
                statistics[order.meal] = (statistics[order.meal] || 0) + 1;
            }
        });

        Object.keys(statistics).forEach(meal => {
            sheet2.addRow({
                meal: meal,
                count: statistics[meal]
            });
        });

        [sheet1, sheet2].forEach(sheet => {
            const headerRow = sheet.getRow(1);
            headerRow.height = 24;
            headerRow.font = { bold: true, color: { argb: 'FFFFFF' }, size: 11 };
            headerRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '2B4C23' } };
            
            sheet.eachRow((row) => {
                row.eachCell(cell => {
                    cell.alignment = { vertical: 'middle', horizontal: 'center' };
                    cell.border = {
                        top: { style: 'thin', color: { argb: 'DDDDDD' } },
                        bottom: { style: 'thin', color: { argb: 'DDDDDD' } },
                        left: { style: 'thin', color: { argb: 'DDDDDD' } },
                        right: { style: 'thin', color: { argb: 'DDDDDD' } }
                    };
                });
            });
        });

        await workbook.xlsx.writeFile(EXCEL_FILE);
        console.log(`📊 [系統提示] orders.json 與 訂單統計表.xlsx 已同步更新！`);
    } catch (err) {
        console.error('❌ 本機檔案與 Excel 同步失敗：', err.message);
    }
}

// ==================== API 路由 ====================

app.get('/api/orders', async (req, res) => {
    try {
        const result = await pool.query('SELECT order_id AS "orderId", card_id AS "cardId", name, meal, spicy, note, total, timestamp FROM orders ORDER BY order_id DESC;');
        res.json(result.rows);
    } catch (err) {
        console.error('獲取訂單失敗:', err.message);
        res.status(500).json({ success: false, message: '無法取得訂單資料' });
    }
});

app.get('/api/employees', (req, res) => {
    const db = loadUserDatabase();
    res.json(db);
});

app.post('/api/employees', (req, res) => {
    const { cardId, name } = req.body;
    const cleanCardId = cardId ? String(cardId).trim() : '';
    const cleanName = name ? String(name).trim() : '';

    if (!cleanCardId || !cleanName) {
        return res.status(400).json({ success: false, message: '卡號與姓名不可為空！' });
    }

    const db = loadUserDatabase();
    if (db[cleanCardId]) {
        return res.status(400).json({ success: false, message: '此卡號已經存在！' });
    }

    db[cleanCardId] = cleanName;
    saveUserDatabase(db);
    res.json({ success: true, message: '新增成功' });
});

app.delete('/api/employees/:cardId', (req, res) => {
    const { cardId } = req.params;
    const cleanCardId = cardId ? String(cardId).trim() : '';

    const db = loadUserDatabase();
    if (!db[cleanCardId]) {
        return res.status(404).json({ success: false, message: '找不到該卡號的員工' });
    }

    delete db[cleanCardId];
    saveUserDatabase(db);
    res.json({ success: true, message: '刪除成功' });
});

app.delete('/api/orders/:orderId', async (req, res) => {
    const { orderId } = req.params;
    const targetOrderIdStr = orderId ? String(orderId).trim() : '';

    if (!targetOrderIdStr) {
        return res.status(400).json({ success: false, message: '無效的訂單編號！' });
    }

    try {
        const deleteQuery = 'DELETE FROM orders WHERE order_id = $1 RETURNING *;';
        const result = await pool.query(deleteQuery, [targetOrderIdStr]);
        
        if (result.rowCount === 0) {
            return res.status(404).json({ success: false, message: `找不到訂單編號: ${targetOrderIdStr}` });
        }

        await syncLocalFilesAndExcel();
        res.json({ success: true, message: `訂單 ${targetOrderIdStr} 已成功刪除，JSON 與 Excel 亦同步更新。` });

    } catch (error) {
        console.error('後端處理刪除訂單發生錯誤:', error);
        res.status(500).json({ success: false, message: '伺服器刪除資料失敗。' });
    }
});

app.post('/api/login', (req, res) => {
    const { cardId } = req.body;
    const cleanCardId = cardId ? String(cardId).trim() : '';
    
    const db = loadUserDatabase();
    if (db[cleanCardId]) {
        res.json({ success: true, message: '登入成功', empName: db[cleanCardId] });
    } else {
        res.json({ success: false, message: '卡號無效，拒絕存取' });
    }
});

app.post('/api/order', async (req, res) => {
    const { cardId, meal, note, spicy, total } = req.body;
    const cleanCardId = cardId ? String(cardId).trim() : '';
    
    const db = loadUserDatabase();
    if (!cleanCardId || !db[cleanCardId]) {
        return res.status(400).json({ success: false, message: '卡號無效或未授權，拒絕下單！' });
    }

    const empName = db[cleanCardId];
    const numTotal = Number(total);
    const orderIdVal = Date.now();
    const timestampStr = new Date().toLocaleString('zh-TW', { timeZone: 'Asia/Taipei' });

    try {
        const insertQuery = `
            INSERT INTO orders (order_id, card_id, name, meal, spicy, note, total, timestamp)
            VALUES ($1, $2, $3, $4, $5, $6, $7, $8);
        `;
        const values = [
            orderIdVal,
            cleanCardId,
            empName,
            meal ? String(meal).trim() : "未知餐點",
            spicy ? String(spicy).trim() : "無",
            note ? String(note).trim() : "無",
            !isNaN(numTotal) ? numTotal : 0,
            timestampStr
        ];

        await pool.query(insertQuery, values);
        console.log(`[新訂單提示] 收到來自 ${empName} (${cleanCardId}) 的訂單，已同步至 PostgreSQL 與 orders.json！`);

        syncLocalFilesAndExcel().catch(e => console.error('同步檔案失敗:', e.message));

        res.json({ success: true, message: `🎉 訂單送出成功！` });

    } catch (error) {
        console.error('後端處理訂單發生錯誤:', error);
        res.status(500).json({ success: false, message: '伺服器寫入資料庫失敗。' });
    }
});

app.get('/api/order-history', async (req, res) => {
    const { cardId } = req.query;
    const cleanCardId = cardId ? String(cardId).trim() : '';

    if (!cleanCardId) {
        return res.json({ success: false, message: '缺少員工卡號', orders: [] });
    }

    try {
        const result = await pool.query(
            'SELECT order_id, meal, spicy, note, total, timestamp FROM orders WHERE card_id = $1 ORDER BY order_id DESC;',
            [cleanCardId]
        );

        const userOrders = result.rows
            .filter(order => isTodayInTaipei(order.order_id))
            .map(order => ({
                time: order.timestamp ? String(order.timestamp) : '-',
                meal: order.meal,
                note: `醬料辣度: ${order.spicy} | 備註: ${order.note} | 金額: $${order.total}`
            }));

        res.json({ success: true, orders: userOrders });

    } catch (error) {
        console.error('讀取紀錄失敗:', error);
        res.json({ success: false, message: '讀取歷史紀錄失敗', orders: [] });
    }
});

app.listen(PORT, () => {
    console.log(`================================================================`);
    console.log(` 🚀 訂餐系統後端已啟動！Port: ${PORT}`);
    console.log(` 📂 檔案同步路徑設定為：${TARGET_DIR}`);
    console.log(`================================================================`);
});
