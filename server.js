require('dotenv').config(); // 載入 .env 環境變數
const express = require('express');
const path = require('path');
const fs = require('fs');
const ExcelJS = require('exceljs');
const { Pool, types } = require('pg');

// 自動將 PostgreSQL BIGINT (OID 20) 解析為 JavaScript Number
types.setTypeParser(20, val => parseInt(val, 10));

const app = express();
const PORT = process.env.PORT || 3000;

app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.use(express.static(path.join(__dirname, 'public')));

// PostgreSQL (Neon) 連線設定
const pool = new Pool({
    connectionString: process.env.DATABASE_URL,
    ssl: {
        rejectUnauthorized: false // Neon 強制使用 SSL
    }
});

// 預設員工資料
const INITIAL_USER_DB = {
    "4860": "呂佳玟", "4931": "陳佳文", "5514": "張淑娟", "5853": "許家綾",
    "6154": "許庭芬", "6465": "呂盈瑩", "6513": "李承州", "6800": "吳修文",
    "9844": "許毓芬", "9864": "吳亞亭", "10047": "許博捷", "10069": "許民芳",
    "11110": "林佳蘭", "11258": "鄭秋燕", "11744": "施名娟", "12218": "沈佩琪",
    "12757": "黃娉娟", "12996": "鄭雅君", "13228": "宋筱湄", "13266": "梁薽予",
    "14253": "李沛珊", "14778": "施憶宣", "15150": "黃珮瑄", "15440": "施欣玫",
    "16294": "梁婧盈", "16508": "尹語璟", "16661": "陳育倫", "16696": "賴語婕",
    "16925": "李宜珊", "16984": "梁慧如", "17020": "黃泓耀", "17528": "曾雅琴",
    "18445": "李珈豪", "20133": "吳佳文", "601207": "郭勝明", "601471": "陳永育",
    "601473": "李羽茹", "TEST": "測試員"
};

// 輔助函式：同步寫入 public/orders.json (非同步寫入避免阻塞)
async function syncOrdersJsonFile() {
    try {
        const result = await pool.query(`
            SELECT 
                o.order_id AS "orderId", 
                o.card_id AS "cardId", 
                o.name, 
                o.meal, 
                o.spicy, 
                o.note, 
                o.total, 
                o.timestamp,
                COALESCE(u.is_paid, false) AS "isPaid"
            FROM orders o
            LEFT JOIN users u ON o.card_id = u.card_id
            ORDER BY o.order_id DESC;
        `);
        const jsonPath = path.join(__dirname, 'public', 'orders.json');
        await fs.promises.writeFile(jsonPath, JSON.stringify(result.rows, null, 2), 'utf-8');
        console.log('[系統提示] 已同步更新 public/orders.json 檔案');
    } catch (err) {
        console.error('❌ 同步 orders.json 失敗:', err.message);
    }
}

// 初始化資料庫
async function initDatabase() {
    try {
        await pool.query(`
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
        `);

        await pool.query(`
            CREATE TABLE IF NOT EXISTS users (
                card_id VARCHAR(50) PRIMARY KEY,
                name VARCHAR(100) NOT NULL,
                is_paid BOOLEAN DEFAULT FALSE,
                department VARCHAR(100) DEFAULT '未劃分'
            );
        `);

        // 新建系統設定資料表 (用於存放點餐截止時間、開放店家等設定)
        await pool.query(`
            CREATE TABLE IF NOT EXISTS settings (
                key VARCHAR(50) PRIMARY KEY,
                value TEXT NOT NULL
            );
        `);

        // 初始化預設鎖定時間與開放店家
        await pool.query(`
            INSERT INTO settings (key, value)
            VALUES ('order_lock_time', '10:30')
            ON CONFLICT (key) DO NOTHING;
        `);

        await pool.query(`
            INSERT INTO settings (key, value)
            VALUES ('active_store', '老聃飲食')
            ON CONFLICT (key) DO NOTHING;
        `);

        // BUG FIX 1: 初始化預設店家清單 JSON (若不存在)
        const DEFAULT_STORES = JSON.stringify([
            { id: 'store_1', name: '老聃飲食 (預設店家)', category: '美味餐點', isOpen: true },
            { id: 'store_2', name: '50嵐', category: '喝涼涼', isOpen: false },
            { id: 'store_3', name: '得正', category: '喝涼涼', isOpen: false },
            { id: 'store_4', name: '竑食泰式拌飯', category: '美味餐點', isOpen: false }
        ]);

        await pool.query(`
            INSERT INTO settings (key, value)
            VALUES ('store_list', $1)
            ON CONFLICT (key) DO NOTHING;
        `, [DEFAULT_STORES]);

        // 自動檢查並為已建立的 users 表格擴充 is_paid 與 department 欄位
        await pool.query(`
            ALTER TABLE users ADD COLUMN IF NOT EXISTS is_paid BOOLEAN DEFAULT FALSE;
            ALTER TABLE users ADD COLUMN IF NOT EXISTS department VARCHAR(100) DEFAULT '未劃分';
        `);

        const userCheck = await pool.query('SELECT COUNT(*) FROM users;');
        if (parseInt(userCheck.rows[0].count, 10) === 0) {
            console.log('[系統提示] 正在初始化預設員工名單至 Neon PostgreSQL...');
            for (const [cardId, name] of Object.entries(INITIAL_USER_DB)) {
                await pool.query(
                    'INSERT INTO users (card_id, name, is_paid, department) VALUES ($1, $2, false, $3) ON CONFLICT DO NOTHING;',
                    [cardId, name, '未劃分']
                );
            }
        }

        console.log(`[系統提示] Neon PostgreSQL 資料表與員工資料檢查完成！`);
        await syncOrdersJsonFile();
    } catch (err) {
        console.error(`❌ 初始化 Neon 資料庫失敗：`, err.message);
    }
}

initDatabase();

// 輔助函式：判斷時間戳記是否為台北時間的今天
function isTodayInTaipei(orderTimestampMs) {
    const orderDate = new Date(Number(orderTimestampMs));
    if (isNaN(orderDate.getTime())) return false;

    const now = new Date();
    const options = { timeZone: 'Asia/Taipei', year: 'numeric', month: '2-digit', day: '2-digit' };
    
    const orderDateStr = new Intl.DateTimeFormat('zh-TW', options).format(orderDate);
    const todayDateStr = new Intl.DateTimeFormat('zh-TW', options).format(now);

    return orderDateStr === todayDateStr;
}

// ==================== API 路由 ====================

// 0-1. 鎖定時間控制 API
app.get('/api/lock-time', async (req, res) => {
    try {
        const result = await pool.query("SELECT value FROM settings WHERE key = 'order_lock_time';");
        const lockTime = result.rows.length > 0 ? result.rows[0].value : "";
        res.json({ success: true, lockTime });
    } catch (err) {
        console.error('讀取鎖定時間失敗:', err.message);
        res.status(500).json({ success: false, message: '無法取得鎖定時間' });
    }
});

app.post('/api/lock-time', async (req, res) => {
    const { lockTime } = req.body;
    const newLockTime = lockTime ? String(lockTime).trim() : "";

    try {
        await pool.query(`
            INSERT INTO settings (key, value) 
            VALUES ('order_lock_time', $1)
            ON CONFLICT (key) DO UPDATE SET value = EXCLUDED.value;
        `, [newLockTime]);

        if (!newLockTime) {
            console.log(`[系統提示] 管理員已解除點餐時間限制`);
            return res.json({ 
                success: true, 
                message: '已成功解除點餐限制！目前系統無時間限制，隨時可進行下單。',
                lockTime: "" 
            });
        }

        console.log(`[系統提示] 管理員將點餐截止時間更新為：${newLockTime}`);
        return res.json({ 
            success: true, 
            message: `已成功將點餐截止時間設定為 ${newLockTime}`,
            lockTime: newLockTime
        });
    } catch (err) {
        console.error('更新鎖定時間失敗:', err.message);
        res.status(500).json({ success: false, message: '設定鎖定時間失敗' });
    }
});

// 0-2. 開放店家控制 API (存取與更新店家清單)
app.get('/api/stores', async (req, res) => {
    try {
        const result = await pool.query("SELECT value FROM settings WHERE key = 'store_list';");
        if (result.rows.length > 0) {
            res.json({ success: true, stores: JSON.parse(result.rows[0].value) });
        } else {
            res.json({ success: true, stores: [] });
        }
    } catch (err) {
        console.error('讀取店家設定失敗:', err.message);
        res.status(500).json({ success: false, message: '無法取得店家設定' });
    }
});

app.post('/api/stores', async (req, res) => {
    const { stores } = req.body;
    
    if (!Array.isArray(stores)) {
        return res.status(400).json({ success: false, message: '店家資料格式不正確！' });
    }

    try {
        // 1. 寫入完整店家 JSON 設定
        await pool.query(`
            INSERT INTO settings (key, value) 
            VALUES ('store_list', $1)
            ON CONFLICT (key) DO UPDATE SET value = EXCLUDED.value;
        `, [JSON.stringify(stores)]);

        // BUG FIX 2: 找出目前唯一開放的店家名稱，並自動過濾括號標籤 (例如: "老聃飲食 (預設店家)" -> "老聃飲食")
        const activeStoreObj = stores.find(s => s.isOpen);
        const activeStoreName = activeStoreObj 
            ? activeStoreObj.name.replace(/\s*\([^)]*\)/g, '').trim() 
            : "";

        // 3. 更新 active_store 供下單時比對
        await pool.query(`
            INSERT INTO settings (key, value) 
            VALUES ('active_store', $1)
            ON CONFLICT (key) DO UPDATE SET value = EXCLUDED.value;
        `, [activeStoreName]);

        console.log(`[系統提示] 管理員更新店家開放狀態，目前開放：${activeStoreName || '無限制 (全開放)'}`);
        return res.json({ 
            success: true, 
            message: activeStoreName ? `已成功將開放店家設定為：${activeStoreName}` : '已解除店家限制，目前開放所有店家訂購！',
            activeStore: activeStoreName,
            stores: stores
        });
    } catch (err) {
        console.error('更新開放店家失敗:', err.message);
        res.status(500).json({ success: false, message: '設定開放店家失敗' });
    }
});

// 0-3. 單獨開放店家設定 API
app.get('/api/active-store', async (req, res) => {
    try {
        const result = await pool.query("SELECT value FROM settings WHERE key = 'active_store';");
        const activeStore = result.rows.length > 0 ? result.rows[0].value : "";
        res.json({ success: true, activeStore });
    } catch (err) {
        console.error('讀取開放店家失敗:', err.message);
        res.status(500).json({ success: false, message: '無法取得開放店家設定' });
    }
});

app.post('/api/active-store', async (req, res) => {
    const { activeStore } = req.body;
    const storeName = activeStore ? String(activeStore).trim() : "";
    const cleanStoreName = storeName.replace(/\s*\([^)]*\)/g, '').trim();

    try {
        await pool.query(`
            INSERT INTO settings (key, value) 
            VALUES ('active_store', $1)
            ON CONFLICT (key) DO UPDATE SET value = EXCLUDED.value;
        `, [cleanStoreName]);

        console.log(`[系統提示] 管理員將開放訂餐店家更新為：${cleanStoreName || '全不限制'}`);
        return res.json({ 
            success: true, 
            message: cleanStoreName ? `已成功將開放店家設定為：${cleanStoreName}` : '已解除店家限制！',
            activeStore: cleanStoreName
        });
    } catch (err) {
        console.error('更新開放店家失敗:', err.message);
        res.status(500).json({ success: false, message: '設定開放店家失敗' });
    }
});

// 1. 取得所有訂單 (含關聯員工的繳費狀態 isPaid)
app.get('/api/orders', async (req, res) => {
    try {
        const result = await pool.query(`
            SELECT 
                o.order_id AS "orderId", 
                o.card_id AS "cardId", 
                o.name, 
                o.meal, 
                o.spicy, 
                o.note, 
                o.total, 
                o.timestamp,
                COALESCE(u.is_paid, false) AS "isPaid"
            FROM orders o
            LEFT JOIN users u ON o.card_id = u.card_id
            ORDER BY o.order_id DESC;
        `);
        res.json(result.rows);
    } catch (err) {
        console.error('獲取訂單失敗:', err.message);
        res.status(500).json({ success: false, message: '無法取得訂單資料' });
    }
});

// 修改單筆訂單的繳費狀態 API
app.patch('/api/orders/:orderId/payment', async (req, res) => {
    const { orderId } = req.params;
    const { isPaid } = req.body;
    const targetOrderIdStr = orderId ? String(orderId).trim() : '';
    const paidBool = isPaid === true || isPaid === 'true';

    if (!targetOrderIdStr) {
        return res.status(400).json({ success: false, message: '無效的訂單編號！' });
    }

    try {
        const orderRes = await pool.query('SELECT card_id FROM orders WHERE order_id = $1;', [targetOrderIdStr]);
        
        if (orderRes.rows.length === 0) {
            return res.status(404).json({ success: false, message: `找不到訂單編號: ${targetOrderIdStr}` });
        }

        const cardId = orderRes.rows[0].card_id;

        await pool.query(
            'UPDATE users SET is_paid = $1 WHERE card_id = $2;',
            [paidBool, cardId]
        );

        await syncOrdersJsonFile();

        res.json({ success: true, message: '訂單繳費狀態更新成功！' });
    } catch (err) {
        console.error('更新訂單繳費狀態失敗:', err.message);
        res.status(500).json({ success: false, message: '伺服器更新繳費狀態失敗。' });
    }
});

// 2. 取得所有員工名單 (包含 department)
app.get('/api/employees', async (req, res) => {
    try {
        const result = await pool.query('SELECT card_id, name, COALESCE(is_paid, false) AS is_paid, COALESCE(department, \'未劃分\') AS department FROM users;');
        const dbMap = {};
        result.rows.forEach(row => {
            dbMap[row.card_id] = {
                name: row.name,
                isPaid: row.is_paid,
                department: row.department
            };
        });
        res.json(dbMap);
    } catch (err) {
        console.error('讀取員工資料失敗:', err.message);
        res.status(500).json({ success: false, message: '無法取得員工資料' });
    }
});

// 3. 切換員工繳費狀態
app.patch('/api/employees/:cardId/payment', async (req, res) => {
    const { cardId } = req.params;
    const { isPaid } = req.body;
    const cleanCardId = cardId ? String(cardId).trim() : '';
    const paidBool = isPaid === true || isPaid === 'true';

    try {
        const result = await pool.query(
            'UPDATE users SET is_paid = $1 WHERE card_id = $2 RETURNING *;',
            [paidBool, cleanCardId]
        );

        if (result.rowCount === 0) {
            return res.status(404).json({ success: false, message: '找不到該卡號的員工' });
        }

        await syncOrdersJsonFile();
        res.json({ success: true, message: '繳費狀態更新成功' });
    } catch (err) {
        console.error('更新員工繳費狀態失敗:', err.message);
        res.status(500).json({ success: false, message: '更新失敗' });
    }
});

// 4. 編輯員工姓名與部門 API
app.put('/api/employees/:cardId', async (req, res) => {
    const { cardId } = req.params;
    const { name, department } = req.body;
    
    const cleanCardId = cardId ? String(cardId).trim() : '';
    const cleanName = name ? String(name).trim() : '';
    const cleanDept = department ? String(department).trim() : '未劃分';

    if (!cleanCardId || !cleanName) {
        return res.status(400).json({ success: false, message: '員工卡號與姓名不可為空！' });
    }

    try {
        const result = await pool.query(
            'UPDATE users SET name = $1, department = $2 WHERE card_id = $3 RETURNING *;',
            [cleanName, cleanDept, cleanCardId]
        );

        if (result.rowCount === 0) {
            return res.status(404).json({ success: false, message: '找不到該卡號的員工' });
        }

        await pool.query('UPDATE orders SET name = $1 WHERE card_id = $2;', [cleanName, cleanCardId]);
        
        await syncOrdersJsonFile();
        res.json({ success: true, message: '員工資料更新成功' });
    } catch (err) {
        console.error('更新員工資料失敗:', err.message);
        res.status(500).json({ success: false, message: '更新員工資料失敗' });
    }
});

// 5. 新增員工 (包含 department)
app.post('/api/employees', async (req, res) => {
    const { cardId, name, department, isPaid } = req.body;
    const cleanCardId = cardId ? String(cardId).trim() : '';
    const cleanName = name ? String(name).trim() : '';
    const cleanDept = department ? String(department).trim() : '未劃分';
    const paidBool = isPaid === true || isPaid === 'true';

    if (!cleanCardId || !cleanName) {
        return res.status(400).json({ success: false, message: '卡號與姓名不可為空！' });
    }

    try {
        const check = await pool.query('SELECT card_id FROM users WHERE card_id = $1;', [cleanCardId]);
        if (check.rows.length > 0) {
            return res.status(400).json({ success: false, message: '此卡號已經存在！' });
        }

        await pool.query(
            'INSERT INTO users (card_id, name, is_paid, department) VALUES ($1, $2, $3, $4);', 
            [cleanCardId, cleanName, paidBool, cleanDept]
        );
        res.json({ success: true, message: '新增成功' });
    } catch (err) {
        console.error('新增員工失敗:', err.message);
        res.status(500).json({ success: false, message: '新增員工失敗' });
    }
});

// 6. 刪除員工
app.delete('/api/employees/:cardId', async (req, res) => {
    const { cardId } = req.params;
    const cleanCardId = cardId ? String(cardId).trim() : '';

    try {
        const result = await pool.query('DELETE FROM users WHERE card_id = $1 RETURNING *;', [cleanCardId]);
        if (result.rowCount === 0) {
            return res.status(404).json({ success: false, message: '找不到該卡號的員工' });
        }
        res.json({ success: true, message: '刪除成功' });
    } catch (err) {
        console.error('刪除員工失敗:', err.message);
        res.status(500).json({ success: false, message: '刪除員工失敗' });
    }
});

// 7. 刪除訂單
app.delete('/api/orders/:orderId', async (req, res) => {
    const { orderId } = req.params;
    const targetOrderIdStr = orderId ? String(orderId).trim() : '';

    if (!targetOrderIdStr) {
        return res.status(400).json({ success: false, message: '無效的訂單編號！' });
    }

    try {
        const result = await pool.query('DELETE FROM orders WHERE order_id = $1 RETURNING *;', [targetOrderIdStr]);
        if (result.rowCount === 0) {
            return res.status(404).json({ success: false, message: `找不到訂單編號: ${targetOrderIdStr}` });
        }

        await syncOrdersJsonFile();
        res.json({ success: true, message: `訂單 ${targetOrderIdStr} 已成功刪除。` });
    } catch (error) {
        console.error('後端處理刪除訂單發生錯誤:', error);
        res.status(500).json({ success: false, message: '伺服器刪除資料失敗。' });
    }
});

// 8. 登入驗證
app.post('/api/login', async (req, res) => {
    const { cardId } = req.body;
    const cleanCardId = cardId ? String(cardId).trim() : '';
    
    try {
        const result = await pool.query('SELECT name FROM users WHERE card_id = $1;', [cleanCardId]);
        if (result.rows.length > 0) {
            res.json({ success: true, message: '登入成功', empName: result.rows[0].name });
        } else {
            res.json({ success: false, message: '卡號無效，拒絕存取' });
        }
    } catch (err) {
        console.error('登入驗證錯誤:', err.message);
        res.status(500).json({ success: false, message: '系統錯誤' });
    }
});

// 9. 新增訂單 (加入開放店家與時間阻擋邏輯)
app.post('/api/order', async (req, res) => {
    try {
        const { cardId, meal, note, spicy, total } = req.body;
        const cleanMeal = meal ? String(meal).trim() : "";

        // ---- 1. 開放店家檢查邏輯 (包含備註括號過濾) ----
        const storeRes = await pool.query("SELECT value FROM settings WHERE key = 'active_store';");
        let activeStore = storeRes.rows.length > 0 ? storeRes.rows[0].value.trim() : "";

        // 去除名稱內的備註括號 (例如 "老聃飲食 (預設店家)" 轉為 "老聃飲食")
        const cleanStoreKeyword = activeStore.replace(/\s*\([^)]*\)/g, '').trim();

        // 如果管理者有設定開放特定店家，但點購品項不包含該店家關鍵字，則阻擋下單
        if (cleanStoreKeyword && !cleanMeal.includes(cleanStoreKeyword)) {
            return res.status(403).json({
                success: false,
                message: `⚠️ 今日僅開放訂購「${cleanStoreKeyword}」，其他店家暫未開放！`
            });
        }

        // ---- 2. 點餐鎖定時間比對邏輯 ----
        const lockRes = await pool.query("SELECT value FROM settings WHERE key = 'order_lock_time';");
        const orderLockTime = lockRes.rows.length > 0 ? lockRes.rows[0].value : "";

        if (orderLockTime) {
            const now = new Date();
            const taipeiTimeStr = new Intl.DateTimeFormat('zh-TW', {
                timeZone: 'Asia/Taipei',
                hour: '2-digit',
                minute: '2-digit',
                hour12: false
            }).format(now);

            const [currentH, currentM] = taipeiTimeStr.split(':').map(Number);
            const currentTotalMinutes = currentH * 60 + currentM;

            const [lockH, lockM] = orderLockTime.split(':').map(Number);
            const lockTotalMinutes = lockH * 60 + lockM;

            if (currentTotalMinutes >= lockTotalMinutes) {
                return res.status(403).json({
                    success: false,
                    message: `⏰ 點餐已截止！今日點餐時間限制至 ${orderLockTime}，系統目前已鎖定無法再下單。`
                });
            }
        }

        const cleanCardId = cardId ? String(cardId).trim() : '';
        const userRes = await pool.query('SELECT name FROM users WHERE card_id = $1;', [cleanCardId]);
        if (!cleanCardId || userRes.rows.length === 0) {
            return res.status(400).json({ success: false, message: '卡號無效或未授權，拒絕下單！' });
        }

        const empName = userRes.rows[0].name;
        const numTotal = Number(total);
        const orderIdVal = Date.now();
        const timestampStr = new Date().toLocaleString('zh-TW', { timeZone: 'Asia/Taipei' });

        const insertQuery = `
            INSERT INTO orders (order_id, card_id, name, meal, spicy, note, total, timestamp)
            VALUES ($1, $2, $3, $4, $5, $6, $7, $8);
        `;
        const values = [
            orderIdVal,
            cleanCardId,
            empName,
            cleanMeal || "未知餐點",
            spicy ? String(spicy).trim() : "無",
            note ? String(note).trim() : "無",
            !isNaN(numTotal) ? numTotal : 0,
            timestampStr
        ];

        await pool.query(insertQuery, values);
        console.log(`[新訂單提示] 收到來自 ${empName} (${cleanCardId}) 的訂單，已寫入 Neon PostgreSQL！`);

        syncOrdersJsonFile().catch(err => console.error(err));

        res.json({ success: true, message: `🎉 訂單送出成功！` });
    } catch (error) {
        console.error('後端處理訂單發生錯誤:', error);
        res.status(500).json({ success: false, message: '伺服器寫入資料庫失敗。' });
    }
});

// 10. 取得個人歷史訂單紀錄
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

// 11. 動態匯出 Excel 下載 (包含繳費狀態)
app.get('/api/export-excel', async (req, res) => {
    try {
        const result = await pool.query(`
            SELECT 
                o.order_id, o.timestamp, o.card_id, o.name, o.meal, o.spicy, o.note, o.total,
                COALESCE(u.is_paid, false) AS is_paid
            FROM orders o
            LEFT JOIN users u ON o.card_id = u.card_id
            ORDER BY o.order_id ASC;
        `);
        
        const orders = result.rows.map(row => ({
            orderId: Number(row.order_id),
            timestamp: row.timestamp,
            cardId: row.card_id,
            name: row.name,
            meal: row.meal,
            spicy: row.spicy,
            note: row.note,
            total: Number(row.total),
            isPaid: row.is_paid ? '已繳費' : '未繳費'
        }));

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
            { header: '金額', key: 'total', width: 12 },
            { header: '繳費狀態', key: 'isPaid', width: 12 }
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
                total: !isNaN(order.total) ? order.total : 0,
                isPaid: order.isPaid
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

        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', `attachment; filename=orders_${Date.now()}.xlsx`);

        await workbook.xlsx.write(res);
        res.end();
    } catch (err) {
        console.error('導出 Excel 失敗:', err.message);
        res.status(500).json({ success: false, message: '無法產生 Excel 報表' });
    }
});

app.listen(PORT, () => {
    console.log(`================================================================`);
    console.log(` 🚀 訂餐系統後端已啟動！Port: ${PORT}`);
    console.log(`================================================================`);
});
