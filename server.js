require('dotenv').config(); // 載入 .env 環境變數[cite: 2]
const express = require('express');[cite: 2]
const path = require('path');[cite: 2]
const fs = require('fs');[cite: 2]
const ExcelJS = require('exceljs');[cite: 2]
const { Pool, types } = require('pg');[cite: 2]
const line = require('@line/bot-sdk'); // 引入 LINE SDK[cite: 2]

// 自動將 PostgreSQL BIGINT (OID 20) 解析為 JavaScript Number[cite: 2]
types.setTypeParser(20, val => parseInt(val, 10));[cite: 2]

const app = express();[cite: 2]
const PORT = process.env.PORT || 3000;[cite: 2]

app.use(express.json({ limit: '1mb' }));[cite: 2]
app.use(express.urlencoded({ extended: true, limit: '1mb' }));[cite: 2]
app.use(express.static(path.join(__dirname, 'public')));[cite: 2]

// LINE Messaging API 設定與初始化[cite: 2]
const lineConfig = {
    channelAccessToken: process.env.LINE_CHANNEL_ACCESS_TOKEN,
    channelSecret: process.env.LINE_CHANNEL_SECRET
};
const lineClient = new line.Client(lineConfig);[cite: 2]

// PostgreSQL (Neon) 連線設定[cite: 2]
const pool = new Pool({
    connectionString: process.env.DATABASE_URL,
    ssl: {
        rejectUnauthorized: false // Neon 強制使用 SSL[cite: 2]
    },
    max: 10,
    idleTimeoutMillis: 30000,
    connectionTimeoutMillis: 2000
});[cite: 2]

// 預設員工資料[cite: 2]
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
};[cite: 2]

// 輔助函式：同步寫入 public/orders.json (改為讀取 orders 表的 is_paid)[cite: 2]
async function syncOrdersJsonFile() {
    try {
        const publicDir = path.join(__dirname, 'public');
        if (!fs.existsSync(publicDir)) {
            await fs.promises.mkdir(publicDir, { recursive: true });
        }

        const result = await pool.query(`
            SELECT 
                order_id AS "orderId", 
                card_id AS "cardId", 
                name, 
                meal, 
                spicy, 
                note, 
                total, 
                timestamp,
                COALESCE(is_paid, false) AS "isPaid"
            FROM orders
            ORDER BY order_id DESC;
        `);
        const jsonPath = path.join(publicDir, 'orders.json');
        await fs.promises.writeFile(jsonPath, JSON.stringify(result.rows, null, 2), 'utf-8');
    } catch (err) {
        console.warn('⚠️ 警告：同步 orders.json 失敗 (雲端環境可能為唯讀檔案系統):', err.message);
    }
}[cite: 2]

// 初始化資料庫[cite: 2]
async function initDatabase() {
    try {
        // 在 orders 資料表建立時直接包含 is_paid[cite: 2]
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
                is_paid BOOLEAN DEFAULT FALSE,
                created_at TIMESTAMP WITH TIME ZONE DEFAULT CURRENT_TIMESTAMP
            );
        `);[cite: 2]

        // 自動補建舊 orders 資料表可能缺少的 is_paid 欄位[cite: 2]
        await pool.query(`
            ALTER TABLE orders ADD COLUMN IF NOT EXISTS is_paid BOOLEAN DEFAULT FALSE;
        `);[cite: 2]

        await pool.query(`
            CREATE TABLE IF NOT EXISTS users (
                card_id VARCHAR(50) PRIMARY KEY,
                name VARCHAR(100) NOT NULL,
                is_paid BOOLEAN DEFAULT FALSE,
                department VARCHAR(100) DEFAULT '未劃分'
            );
        `);[cite: 2]

        await pool.query(`
            CREATE TABLE IF NOT EXISTS settings (
                key VARCHAR(50) PRIMARY KEY,
                value TEXT NOT NULL
            );
        `);[cite: 2]

        await pool.query(`
            INSERT INTO settings (key, value)
            VALUES ('order_lock_time', '10:30')
            ON CONFLICT (key) DO NOTHING;
        `);[cite: 2]

        await pool.query(`
            INSERT INTO settings (key, value)
            VALUES ('active_store', '老聃飲食')
            ON CONFLICT (key) DO NOTHING;
        `);[cite: 2]

        const DEFAULT_STORES = [
            { id: 'store_1', name: '八方雲集', category: '美味餐點', isOpen: true },
            { id: 'store_2', name: '飯大廚', category: '美味餐點', isOpen: false },
            { id: 'store_3', name: '鼎森泰式料理', category: '美味餐點', isOpen: false },
            { id: 'store_4', name: '榮郁香廣式燒臘', category: '美味餐點', isOpen: false },
            { id: 'store_5', name: '顏舍關東煮', category: '美味餐點', isOpen: true },
            { id: 'store_6', name: '羹香鴨肉羹', category: '美味餐點', isOpen: false },
            { id: 'store_7', name: '徊香麵線糊', category: '美味餐點', isOpen: false },
            { id: 'store_8', name: '四海遊龍', category: '美味餐點', isOpen: false },
            { id: 'store_9', name: '北斗賓肉圓', category: '美味餐點', isOpen: true },
            { id: 'store_10', name: '百華味滷味', category: '美味餐點', isOpen: false },
            { id: 'store_11', name: '老聃飲食', category: '美味餐點', isOpen: false },
            { id: 'store_12', name: '八廚職人弁当', category: '美味餐點', isOpen: false },
            { id: 'store_13', name: '嵐 爌肉・豬腳飯', category: '美味餐點', isOpen: true },
            { id: 'store_14', name: '旺春豐傳統小吃', category: '美味餐點', isOpen: false },
            { id: 'store_15', name: '咕雞鹽水雞', category: '美味餐點', isOpen: false },
            { id: 'store_16', name: '家灶傳統爌肉豬腳專賣', category: '美味餐點', isOpen: false },
            { id: 'store_17', name: '菓蔬輕蒔坊', category: '美味餐點', isOpen: true },
            { id: 'store_18', name: '老蕭土魠魚羹麵館', category: '美味餐點', isOpen: false },
            { id: 'store_19', name: '小松丼丼食事處', category: '美味餐點', isOpen: false },
            { id: 'store_20', name: '滿座燒肉丼飯屋', category: '美味餐點', isOpen: false },
            { id: 'store_21', name: '地6攤排餐', category: '美味餐點', isOpen: true },
            { id: 'store_22', name: '八鮮雲吞', category: '美味餐點', isOpen: false },
            { id: 'store_23', name: '明美快餐', category: '美味餐點', isOpen: false },
            { id: 'store_24', name: '森晴餐盒製所', category: '美味餐點', isOpen: false },
            { id: 'store_25', name: '廈門沙茶麵', category: '美味餐點', isOpen: true },
            { id: 'store_26', name: '凡凡滷味', category: '美味餐點', isOpen: false },
            { id: 'store_27', name: '員湘園', category: '美味餐點', isOpen: false },
            { id: 'store_28', name: '志氣雞飯', category: '美味餐點', isOpen: false },
            { id: 'store_29', name: '竑食泰式拌飯', category: '美味餐點', isOpen: true },
            { id: 'store_30', name: '50嵐', category: '喝涼涼', isOpen: false },
            { id: 'store_31', name: 'TEA TOP', category: '喝涼涼', isOpen: false },
            { id: 'store_32', name: '得正', category: '喝涼涼', isOpen: false },
            { id: 'store_33', name: '尚淳草本茶', category: '喝涼涼', isOpen: false },
            { id: 'store_34', name: '烏弄', category: '喝涼涼', isOpen: false },
            { id: 'store_35', name: '可不可熟成紅茶', category: '喝涼涼', isOpen: false },
            { id: 'store_36', name: '鹿兒角', category: '喝涼涼', isOpen: false },
            { id: 'store_37', name: '八曜和茶', category: '喝涼涼', isOpen: false },
            { id: 'store_38', name: '一沐日', category: '喝涼涼', isOpen: false },
            { id: 'store_39', name: '粉圓伯', category: '喝涼涼', isOpen: false },
            { id: 'store_40', name: '先喝道', category: '喝涼涼', isOpen: false },
            { id: 'store_41', name: '李珍心綠豆沙專門店', category: '喝涼涼', isOpen: false },
            { id: 'store_42', name: '吳家紅茶冰', category: '喝涼涼', isOpen: false },
            { id: 'store_43', name: '有茶', category: '喝涼涼', isOpen: false },
            { id: 'store_44', name: '青山', category: '喝涼涼', isOpen: false },
            { id: 'store_45', name: '迷客夏', category: '喝涼涼', isOpen: false },
            { id: 'store_46', name: '茗沏', category: '喝涼涼', isOpen: false },
            { id: 'store_47', name: '李記紅茶冰', category: '喝涼涼', isOpen: false },
            { id: 'store_48', name: '旅人阿宏', category: '喝涼涼', isOpen: false },
            { id: 'store_49', name: '麻古茶坊', category: '喝涼涼', isOpen: false },
            { id: 'store_50', name: '茶湯會', category: '喝涼涼', isOpen: false },
            { id: 'store_51', name: '樂法', category: '喝涼涼', isOpen: false },
            { id: 'store_52', name: '一鳴', category: '美味餐點', isOpen: false },
            { id: 'store_53', name: '五路鍋聖', category: '美味餐點', isOpen: false }
        ];[cite: 2]

        const currentStoreRes = await pool.query("SELECT value FROM settings WHERE key = 'store_list';");[cite: 2]
        if (currentStoreRes.rows.length === 0) {
            await pool.query(`
                INSERT INTO settings (key, value)
                VALUES ('store_list', $1);
            `, [JSON.stringify(DEFAULT_STORES)]);[cite: 2]
        } else {
            try {
                let existingStores = JSON.parse(currentStoreRes.rows[0].value) || [];[cite: 2]
                if (!Array.isArray(existingStores)) existingStores = [];[cite: 2]
                
                const syncedStores = DEFAULT_STORES.map(defStore => {
                    const match = existingStores.find(s => s && s.id === defStore.id);
                    return {
                        ...defStore,
                        isOpen: match ? Boolean(match.isOpen) : defStore.isOpen
                    };
                });[cite: 2]

                await pool.query(`
                    INSERT INTO settings (key, value)
                    VALUES ('store_list', $1)
                    ON CONFLICT (key) DO UPDATE SET value = EXCLUDED.value;
                `, [JSON.stringify(syncedStores)]);[cite: 2]
            } catch (e) {
                console.error('解析現有店家清單失敗，重新寫入預設店家:', e);[cite: 2]
                await pool.query(`
                    INSERT INTO settings (key, value)
                    VALUES ('store_list', $1)
                    ON CONFLICT (key) DO UPDATE SET value = EXCLUDED.value;
                `, [JSON.stringify(DEFAULT_STORES)]);[cite: 2]
            }
        }

        await pool.query(`
            ALTER TABLE users ADD COLUMN IF NOT EXISTS is_paid BOOLEAN DEFAULT FALSE;
            ALTER TABLE users ADD COLUMN IF NOT EXISTS department VARCHAR(100) DEFAULT '未劃分';
        `);[cite: 2]

        const userCheck = await pool.query('SELECT COUNT(*) FROM users;');[cite: 2]
        if (parseInt(userCheck.rows[0].count, 10) === 0) {
            console.log('[系統提示] 正在初始化預設員工名單至 Neon PostgreSQL...');[cite: 2]
            const cardIds = Object.keys(INITIAL_USER_DB);[cite: 2]
            const names = Object.values(INITIAL_USER_DB);[cite: 2]

            await pool.query(`
                INSERT INTO users (card_id, name, is_paid, department)
                SELECT * FROM UNNEST($1::text[], $2::text[]) AS t(card_id, name)
                CROSS JOIN (SELECT false AS is_paid, '未劃分' AS department) d
                ON CONFLICT (card_id) DO NOTHING;
            `, [cardIds, names]);[cite: 2]
        }

        console.log(`[系統提示] Neon PostgreSQL 資料表與員工資料檢查完成！`);[cite: 2]
        await syncOrdersJsonFile();[cite: 2]
    } catch (err) {
        console.error(`❌ 初始化 Neon 資料庫失敗：`, err.message);[cite: 2]
        throw err;[cite: 2]
    }
}

// ==================== API 路由 ====================

// 新增：LINE Messaging API 公告推播 路由[cite: 2]
app.post('/api/send-line-notice', async (req, res) => {
    try {
        const targetId = process.env.LINE_TARGET_ID;

        if (!targetId || !process.env.LINE_CHANNEL_ACCESS_TOKEN) {
            return res.status(500).json({
                success: false,
                message: '伺服器未設定 LINE Messaging API 環境變數 (LINE_TARGET_ID 或 LINE_CHANNEL_ACCESS_TOKEN)！'
            });
        }

        // 1. 取得開放店家列表
        const storeRes = await pool.query("SELECT value FROM settings WHERE key = 'store_list';");
        let activeStoresText = '• 八方雲集 (美味餐點)';

        if (storeRes.rows.length > 0 && storeRes.rows[0].value) {
            try {
                const stores = JSON.parse(storeRes.rows[0].value) || [];
                const openStores = stores.filter(s => s && s.isOpen);
                if (openStores.length > 0) {
                    activeStoresText = openStores.map(s => `• ${s.name} (${s.category || '美味餐點'})`).join('\n');
                } else {
                    activeStoresText = '• 目前無開放店家';
                }
            } catch (e) {
                console.error('解析開放店家失敗:', e);
            }
        }

        // 2. 取得鎖定截止時間
        const lockRes = await pool.query("SELECT value FROM settings WHERE key = 'order_lock_time';");
        const lockTime = (lockRes.rows.length > 0 && lockRes.rows[0].value) ? lockRes.rows[0].value : '未設定';

        // 3. 組合公告訊息
        const messageText = `📢 【今日點餐公告通知】\n\n今日開放點餐店家：\n${activeStoresText}\n\n⏰ 點餐截止時間：${lockTime}\n🔗 前台點餐連結：https://cch1527097.onrender.com/index.html\n\n請大家務必於截止時間前完成下單，謝謝！`;

        // 4. 發送 LINE Push Message
        await lineClient.pushMessage(targetId, {
            type: 'text',
            text: messageText
        });

        res.json({ success: true, message: '🎉 已成功發送 LINE 點餐公告！' });
    } catch (err) {
        console.error('❌ 發送 LINE 公告失敗:', err);
        if (err.originalError && err.originalError.response) {
            console.error('LINE API 錯誤細節:', err.originalError.response.data);
        }
        res.status(500).json({ success: false, message: '發送 LINE 公告失敗，請檢查伺服器設定與 Log！' });
    }
});

app.get('/api/lock-time', async (req, res) => {
    try {
        const result = await pool.query("SELECT value FROM settings WHERE key = 'order_lock_time';");[cite: 2]
        const lockTime = result.rows.length > 0 ? result.rows[0].value : "";[cite: 2]
        res.json({ success: true, lockTime });[cite: 2]
    } catch (err) {
        console.error('讀取鎖定時間失敗:', err.message);[cite: 2]
        res.status(500).json({ success: false, message: '無法取得鎖定時間' });[cite: 2]
    }
});

app.post('/api/lock-time', async (req, res) => {
    const { lockTime } = req.body;[cite: 2]
    const newLockTime = lockTime ? String(lockTime).trim() : "";[cite: 2]

    try {
        await pool.query(`
            INSERT INTO settings (key, value) 
            VALUES ('order_lock_time', $1)
            ON CONFLICT (key) DO UPDATE SET value = EXCLUDED.value;
        `, [newLockTime]);[cite: 2]

        if (!newLockTime) {
            return res.json({ 
                success: true, 
                message: '已成功解除點餐限制！目前系統無時間限制，隨時可進行下單。',
                lockTime: "" 
            });[cite: 2]
        }

        return res.json({ 
            success: true, 
            message: `已成功將點餐截止時間設定為 ${newLockTime}`,
            lockTime: newLockTime
        });[cite: 2]
    } catch (err) {
        console.error('更新鎖定時間失敗:', err.message);[cite: 2]
        res.status(500).json({ success: false, message: '設定鎖定時間失敗' });[cite: 2]
    }
});

app.get('/api/stores', async (req, res) => {
    try {
        const result = await pool.query("SELECT value FROM settings WHERE key = 'store_list';");[cite: 2]
        if (result.rows.length > 0) {
            res.json({ success: true, stores: JSON.parse(result.rows[0].value) });[cite: 2]
        } else {
            res.json({ success: true, stores: [] });[cite: 2]
        }
    } catch (err) {
        console.error('讀取店家設定失敗:', err.message);[cite: 2]
        res.status(500).json({ success: false, message: '無法取得店家設定' });[cite: 2]
    }
});

app.post('/api/stores', async (req, res) => {
    const { stores } = req.body;[cite: 2]
    
    if (!Array.isArray(stores)) {
        return res.status(400).json({ success: false, message: '店家資料格式不正確！' });[cite: 2]
    }

    try {
        await pool.query(`
            INSERT INTO settings (key, value) 
            VALUES ('store_list', $1)
            ON CONFLICT (key) DO UPDATE SET value = EXCLUDED.value;
        `, [JSON.stringify(stores)]);[cite: 2]

        const activeStores = stores.filter(s => s && s.isOpen);[cite: 2]
        const activeStoreNames = activeStores
            .map(s => s.name.replace(/\s*\([^)]*\)/g, '').trim())
            .join('、');[cite: 2]

        await pool.query(`
            INSERT INTO settings (key, value) 
            VALUES ('active_store', $1)
            ON CONFLICT (key) DO UPDATE SET value = EXCLUDED.value;
        `, [activeStoreNames]);[cite: 2]

        return res.json({ 
            success: true, 
            message: activeStoreNames ? `已成功將開放店家設定為：${activeStoreNames}` : '已變更店家設定，目前無開放店家！',
            activeStore: activeStoreNames,
            stores: stores
        });[cite: 2]
    } catch (err) {
        console.error('更新開放店家失敗:', err.message);[cite: 2]
        res.status(500).json({ success: false, message: '設定開放店家失敗' });[cite: 2]
    }
});

app.get('/api/active-store', async (req, res) => {
    try {
        const result = await pool.query("SELECT value FROM settings WHERE key = 'active_store';");[cite: 2]
        const activeStore = result.rows.length > 0 ? result.rows[0].value : "";[cite: 2]
        res.json({ success: true, activeStore });[cite: 2]
    } catch (err) {
        console.error('讀取開放店家失敗:', err.message);[cite: 2]
        res.status(500).json({ success: false, message: '無法取得開放店家設定' });[cite: 2]
    }
});

app.post('/api/active-store', async (req, res) => {
    const { activeStore } = req.body;[cite: 2]
    const storeName = activeStore ? String(activeStore).trim() : "";[cite: 2]
    const cleanStoreName = storeName.replace(/\s*\([^)]*\)/g, '').trim();[cite: 2]

    try {
        await pool.query(`
            INSERT INTO settings (key, value) 
            VALUES ('active_store', $1)
            ON CONFLICT (key) DO UPDATE SET value = EXCLUDED.value;
        `, [cleanStoreName]);[cite: 2]

        return res.json({ 
            success: true, 
            message: cleanStoreName ? `已成功將開放店家設定為：${cleanStoreName}` : '已解除店家限制！',
            activeStore: cleanStoreName
        });[cite: 2]
    } catch (err) {
        console.error('更新開放店家失敗:', err.message);[cite: 2]
        res.status(500).json({ success: false, message: '設定開放店家失敗' });[cite: 2]
    }
});

// 獲取所有訂單（改為直接讀取 orders.is_paid）[cite: 2]
app.get('/api/orders', async (req, res) => {
    try {
        const result = await pool.query(`
            SELECT 
                order_id AS "orderId", 
                card_id AS "cardId", 
                name, 
                meal, 
                spicy, 
                note, 
                total, 
                timestamp,
                COALESCE(is_paid, false) AS "isPaid"
            FROM orders
            ORDER BY order_id DESC;
        `);[cite: 2]
        res.json(result.rows);[cite: 2]
    } catch (err) {
        console.error('獲取訂單失敗:', err.message);[cite: 2]
        res.status(500).json({ success: false, message: '無法取得訂單資料' });[cite: 2]
    }
});

// 切換特定「訂單」的繳費狀態（直接修改 orders 表）[cite: 2]
app.patch('/api/orders/:orderId/payment', async (req, res) => {
    const { orderId } = req.params;[cite: 2]
    const { isPaid } = req.body;[cite: 2]
    const targetOrderIdStr = orderId ? String(orderId).trim() : '';[cite: 2]
    const paidBool = isPaid === true || isPaid === 'true';[cite: 2]

    if (!targetOrderIdStr) {
        return res.status(400).json({ success: false, message: '無效的訂單編號！' });[cite: 2]
    }

    try {
        const result = await pool.query(
            'UPDATE orders SET is_paid = $1 WHERE order_id = $2 RETURNING *;',
            [paidBool, targetOrderIdStr]
        );[cite: 2]
        
        if (result.rowCount === 0) {
            return res.status(404).json({ success: false, message: `找不到訂單編號: ${targetOrderIdStr}` });[cite: 2]
        }

        syncOrdersJsonFile().catch(() => {});[cite: 2]

        res.json({ success: true, message: '訂單繳費狀態更新成功！' });[cite: 2]
    } catch (err) {
        console.error('更新訂單繳費狀態失敗:', err.message);[cite: 2]
        res.status(500).json({ success: false, message: '伺服器更新繳費狀態失敗。' });[cite: 2]
    }
});

app.get('/api/employees', async (req, res) => {
    try {
        const result = await pool.query('SELECT card_id, name, COALESCE(is_paid, false) AS is_paid, COALESCE(department, \'未劃分\') AS department FROM users;');[cite: 2]
        const dbMap = {};[cite: 2]
        result.rows.forEach(row => {
            dbMap[row.card_id] = {
                name: row.name,
                isPaid: row.is_paid,
                department: row.department
            };
        });[cite: 2]
        res.json(dbMap);[cite: 2]
    } catch (err) {
        console.error('讀取員工資料失敗:', err.message);[cite: 2]
        res.status(500).json({ success: false, message: '無法取得員工資料' });[cite: 2]
    }
});

app.patch('/api/employees/:cardId/payment', async (req, res) => {
    const { cardId } = req.params;[cite: 2]
    const { isPaid } = req.body;[cite: 2]
    const cleanCardId = cardId ? String(cardId).trim() : '';[cite: 2]
    const paidBool = isPaid === true || isPaid === 'true';[cite: 2]

    try {
        const result = await pool.query(
            'UPDATE users SET is_paid = $1 WHERE card_id = $2 RETURNING *;',
            [paidBool, cleanCardId]
        );[cite: 2]

        if (result.rowCount === 0) {
            return res.status(404).json({ success: false, message: '找不到該卡號的員工' });[cite: 2]
        }

        syncOrdersJsonFile().catch(() => {});[cite: 2]
        res.json({ success: true, message: '繳費狀態更新成功' });[cite: 2]
    } catch (err) {
        console.error('更新員工繳費狀態失敗:', err.message);[cite: 2]
        res.status(500).json({ success: false, message: '更新失敗' });[cite: 2]
    }
});

app.put('/api/employees/:cardId', async (req, res) => {
    const { cardId } = req.params;[cite: 2]
    const { name, department } = req.body;[cite: 2]
    
    const cleanCardId = cardId ? String(cardId).trim() : '';[cite: 2]
    const cleanName = name ? String(name).trim() : '';[cite: 2]
    const cleanDept = department ? String(department).trim() : '未劃分';[cite: 2]

    if (!cleanCardId || !cleanName) {
        return res.status(400).json({ success: false, message: '員工卡號與姓名不可為空！' });[cite: 2]
    }

    try {
        const result = await pool.query(
            'UPDATE users SET name = $1, department = $2 WHERE card_id = $3 RETURNING *;',
            [cleanName, cleanDept, cleanCardId]
        );[cite: 2]

        if (result.rowCount === 0) {
            return res.status(404).json({ success: false, message: '找不到該卡號的員工' });[cite: 2]
        }

        await pool.query('UPDATE orders SET name = $1 WHERE card_id = $2;', [cleanName, cleanCardId]);[cite: 2]
        
        syncOrdersJsonFile().catch(() => {});[cite: 2]
        res.json({ success: true, message: '員工資料更新成功' });[cite: 2]
    } catch (err) {
        console.error('更新員工資料失敗:', err.message);[cite: 2]
        res.status(500).json({ success: false, message: '更新員工資料失敗' });[cite: 2]
    }
});

app.post('/api/employees', async (req, res) => {
    const { cardId, name, department, isPaid } = req.body;[cite: 2]
    const cleanCardId = cardId ? String(cardId).trim() : '';[cite: 2]
    const cleanName = name ? String(name).trim() : '';[cite: 2]
    const cleanDept = department ? String(department).trim() : '未劃分';[cite: 2]
    const paidBool = isPaid === true || isPaid === 'true';[cite: 2]

    if (!cleanCardId || !cleanName) {
        return res.status(400).json({ success: false, message: '卡號與姓名不可為空！' });[cite: 2]
    }

    try {
        const check = await pool.query('SELECT card_id FROM users WHERE card_id = $1;', [cleanCardId]);[cite: 2]
        if (check.rows.length > 0) {
            return res.status(400).json({ success: false, message: '此卡號已經存在！' });[cite: 2]
        }

        await pool.query(
            'INSERT INTO users (card_id, name, is_paid, department) VALUES ($1, $2, $3, $4);', 
            [cleanCardId, cleanName, paidBool, cleanDept]
        );[cite: 2]
        res.json({ success: true, message: '新增成功' });[cite: 2]
    } catch (err) {
        console.error('新增員工失敗:', err.message);[cite: 2]
        res.status(500).json({ success: false, message: '新增員工失敗' });[cite: 2]
    }
});

app.delete('/api/employees/:cardId', async (req, res) => {
    const { cardId } = req.params;[cite: 2]
    const cleanCardId = cardId ? String(cardId).trim() : '';[cite: 2]

    try {
        const result = await pool.query('DELETE FROM users WHERE card_id = $1 RETURNING *;', [cleanCardId]);[cite: 2]
        if (result.rowCount === 0) {
            return res.status(404).json({ success: false, message: '找不到該卡號的員工' });[cite: 2]
        }
        res.json({ success: true, message: '刪除成功' });[cite: 2]
    } catch (err) {
        console.error('刪除員工失敗:', err.message);[cite: 2]
        res.status(500).json({ success: false, message: '刪除員工失敗' });[cite: 2]
    }
});

app.delete('/api/orders/:orderId', async (req, res) => {
    const { orderId } = req.params;[cite: 2]
    const targetOrderIdStr = orderId ? String(orderId).trim() : '';[cite: 2]

    if (!targetOrderIdStr) {
        return res.status(400).json({ success: false, message: '無效的訂單編號！' });[cite: 2]
    }

    try {
        const result = await pool.query('DELETE FROM orders WHERE order_id = $1 RETURNING *;', [targetOrderIdStr]);[cite: 2]
        if (result.rowCount === 0) {
            return res.status(404).json({ success: false, message: `找不到訂單編號: ${targetOrderIdStr}` });[cite: 2]
        }

        syncOrdersJsonFile().catch(() => {});[cite: 2]
        res.json({ success: true, message: `訂單 ${targetOrderIdStr} 已成功刪除。` });[cite: 2]
    } catch (error) {
        console.error('後端處理刪除訂單發生錯誤:', error);[cite: 2]
        res.status(500).json({ success: false, message: '伺服器刪除資料失敗。' });[cite: 2]
    }
});

app.post('/api/login', async (req, res) => {
    const { cardId } = req.body;[cite: 2]
    const cleanCardId = cardId ? String(cardId).trim() : '';[cite: 2]
    
    try {
        const result = await pool.query('SELECT name FROM users WHERE card_id = $1;', [cleanCardId]);[cite: 2]
        if (result.rows.length > 0) {
            res.json({ success: true, message: '登入成功', empName: result.rows[0].name });[cite: 2]
        } else {
            res.json({ success: false, message: '卡號無效，拒絕存取' });[cite: 2]
        }
    } catch (err) {
        console.error('登入驗證錯誤:', err.message);[cite: 2]
        res.status(500).json({ success: false, message: '系統錯誤' });[cite: 2]
    }
});

// 下訂單 API：預設 is_paid 為 false (未繳費)[cite: 2]
app.post('/api/order', async (req, res) => {
    try {
        const { cardId, meal, storeId, storeName, note, spicy, total } = req.body;[cite: 2]
        const cleanMeal = meal ? String(meal).trim() : "";[cite: 2]
        const reqStoreId = storeId ? String(storeId).trim() : "";[cite: 2]
        const reqStoreName = storeName ? String(storeName).trim() : "";[cite: 2]

        // 1. 店家開放狀態檢查[cite: 2]
        const storeListRes = await pool.query("SELECT value FROM settings WHERE key = 'store_list';");[cite: 2]
        
        if (storeListRes.rows.length > 0 && storeListRes.rows[0].value) {
            const stores = JSON.parse(storeListRes.rows[0].value) || [];[cite: 2]
            let matchedStore = null;[cite: 2]

            if (reqStoreId) {
                matchedStore = stores.find(s => s && s.id === reqStoreId);[cite: 2]
            }

            if (!matchedStore) {
                matchedStore = stores.find(s => {
                    if (!s || !s.name) return false;
                    const cleanName = s.name.replace(/\s*\([^)]*\)/g, '').trim();
                    if (!cleanName) return false;
                    if (reqStoreName && (reqStoreName.includes(cleanName) || cleanName.includes(reqStoreName))) return true;
                    return cleanMeal.includes(cleanName);
                });[cite: 2]
            }

            if (matchedStore && !matchedStore.isOpen) {
                const displayStoreName = matchedStore.name.replace(/\s*\([^)]*\)/g, '').trim();[cite: 2]
                return res.status(403).json({
                    success: false,
                    message: `⚠️ 店家「${displayStoreName}」目前尚未開放點餐！`
                });[cite: 2]
            }
        }

        // 2. 鎖定時間比對[cite: 2]
        const lockRes = await pool.query("SELECT value FROM settings WHERE key = 'order_lock_time';");[cite: 2]
        const orderLockTime = lockRes.rows.length > 0 ? lockRes.rows[0].value : "";[cite: 2]

        if (orderLockTime && orderLockTime.includes(':')) {
            const now = new Date();[cite: 2]
            const taipeiFormatter = new Intl.DateTimeFormat('zh-TW', {
                timeZone: 'Asia/Taipei',
                hour: '2-digit',
                minute: '2-digit',
                hour12: false
            });[cite: 2]
            const parts = taipeiFormatter.formatToParts(now);[cite: 2]
            const currentH = Number(parts.find(p => p.type === 'hour').value);[cite: 2]
            const currentM = Number(parts.find(p => p.type === 'minute').value);[cite: 2]
            const currentTotalMinutes = currentH * 60 + currentM;[cite: 2]

            const [lockH, lockM] = orderLockTime.split(':').map(Number);[cite: 2]
            const lockTotalMinutes = lockH * 60 + lockM;[cite: 2]

            if (currentTotalMinutes >= lockTotalMinutes) {
                return res.status(403).json({
                    success: false,
                    message: `⏰ 點餐已截止！今日點餐時間限制至 ${orderLockTime}，系統目前已鎖定無法再下單。`
                });[cite: 2]
            }
        }

        // 3. 員工驗證與寫入[cite: 2]
        const cleanCardId = cardId ? String(cardId).trim() : '';[cite: 2]
        const userRes = await pool.query('SELECT name FROM users WHERE card_id = $1;', [cleanCardId]);[cite: 2]
        if (!cleanCardId || userRes.rows.length === 0) {
            return res.status(400).json({ success: false, message: '卡號無效或未授權，拒絕下單！' });[cite: 2]
        }

        const empName = userRes.rows[0].name;[cite: 2]
        const numTotal = Number(total);[cite: 2]
        
        // 解決高併發碰撞：時間戳毫秒 + 4位隨機數[cite: 2]
        const orderIdVal = parseInt(`${Date.now()}${Math.floor(1000 + Math.random() * 9000)}`, 10);[cite: 2]
        const timestampStr = new Date().toLocaleString('zh-TW', { timeZone: 'Asia/Taipei' });[cite: 2]

        // 新增訂單預設將 is_paid 設為 false[cite: 2]
        const insertQuery = `
            INSERT INTO orders (order_id, card_id, name, meal, spicy, note, total, timestamp, is_paid)
            VALUES ($1, $2, $3, $4, $5, $6, $7, $8, false);
        `;[cite: 2]
        const values = [
            orderIdVal,
            cleanCardId,
            empName,
            cleanMeal || "未知餐點",
            spicy ? String(spicy).trim() : "無",
            note ? String(note).trim() : "無",
            !isNaN(numTotal) ? numTotal : 0,
            timestampStr
        ];[cite: 2]

        await pool.query(insertQuery, values);[cite: 2]
        syncOrdersJsonFile().catch(() => {});[cite: 2]

        res.json({ success: true, message: `🎉 訂單送出成功！` });[cite: 2]
    } catch (error) {
        console.error('後端處理訂單發生錯誤:', error);[cite: 2]
        res.status(500).json({ success: false, message: '伺服器寫入資料庫失敗。' });[cite: 2]
    }
});

app.get('/api/order-history', async (req, res) => {
    const { cardId } = req.query;[cite: 2]
    const cleanCardId = cardId ? String(cardId).trim() : '';[cite: 2]

    if (!cleanCardId) {
        return res.json({ success: false, message: '缺少員工卡號', orders: [] });[cite: 2]
    }

    try {
        const result = await pool.query(
            `SELECT order_id, meal, spicy, note, total, timestamp 
             FROM orders 
             WHERE card_id = $1 
               AND created_at >= ((CURRENT_DATE AT TIME ZONE 'Asia/Taipei')::timestamp AT TIME ZONE 'Asia/Taipei')
               AND created_at < (((CURRENT_DATE + INTERVAL '1 day') AT TIME ZONE 'Asia/Taipei')::timestamp AT TIME ZONE 'Asia/Taipei')
             ORDER BY order_id DESC;`,
            [cleanCardId]
        );[cite: 2]

        const userOrders = result.rows.map(order => ({
            time: order.timestamp ? String(order.timestamp) : '-',
            meal: order.meal,
            note: `醬料辣度: ${order.spicy} | 備註: ${order.note} | 金額: $${order.total}`
        }));[cite: 2]

        res.json({ success: true, orders: userOrders });[cite: 2]
    } catch (error) {
        console.error('讀取紀錄失敗:', error);[cite: 2]
        res.json({ success: false, message: '讀取歷史紀錄失敗', orders: [] });[cite: 2]
    }
});

// Excel 匯出（直接讀取 orders.is_paid）[cite: 2]
app.get('/api/export-excel', async (req, res) => {
    try {
        const result = await pool.query(`
            SELECT 
                order_id, timestamp, card_id, name, meal, spicy, note, total,
                COALESCE(is_paid, false) AS is_paid
            FROM orders
            ORDER BY order_id ASC;
        `);[cite: 2]
        
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
        }));[cite: 2]

        const workbook = new ExcelJS.Workbook();[cite: 2]
        const sheet1 = workbook.addWorksheet('訂單明細');[cite: 2]
        sheet1.columns = [
            { header: '訂單ID', key: 'orderId', width: 22 },
            { header: '時間', key: 'timestamp', width: 25 },
            { header: '員工卡號', key: 'cardId', width: 15 },
            { header: '員工姓名', key: 'name', width: 15 }, 
            { header: '點餐店家/品項', key: 'meal', width: 25 },
            { header: '醬料辣度', key: 'spicy', width: 15 },
            { header: '備註', key: 'note', width: 25 },
            { header: '金額', key: 'total', width: 12 },
            { header: '繳費狀態', key: 'isPaid', width: 12 }
        ];[cite: 2]
        
        const sheet2 = workbook.addWorksheet('店家點餐統計表');[cite: 2]
        sheet2.columns = [
            { header: '店家/品項名稱', key: 'meal', width: 25 },
            { header: '總點餐次數', key: 'count', width: 15 }
        ];[cite: 2]

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
        });[cite: 2]

        const statistics = {};[cite: 2]
        orders.forEach(order => {
            if (order.meal) {
                statistics[order.meal] = (statistics[order.meal] || 0) + 1;
            }
        });[cite: 2]

        Object.keys(statistics).forEach(meal => {
            sheet2.addRow({
                meal: meal,
                count: statistics[meal]
            });
        });[cite: 2]

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
        });[cite: 2]

        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');[cite: 2]
        res.setHeader('Content-Disposition', `attachment; filename=orders_${Date.now()}.xlsx`);[cite: 2]

        await workbook.xlsx.write(res);[cite: 2]
        res.end();[cite: 2]
    } catch (err) {
        console.error('導出 Excel 失敗:', err.message);[cite: 2]
        res.status(500).json({ success: false, message: '無法產生 Excel 報表' });[cite: 2]
    }
});

// 啟動伺服器[cite: 2]
async function startServer() {
    try {
        await initDatabase();[cite: 2]
        app.listen(PORT, () => {
            console.log(`================================================================`);[cite: 2]
            console.log(` 🚀 訂餐系統後端已啟動！Port: ${PORT}`);[cite: 2]
            console.log(`================================================================`);[cite: 2]
        });
    } catch (err) {
        console.error('❌ 伺服器啟動失敗，無法連線至 Neon PostgreSQL:', err);[cite: 2]
        process.exit(1);[cite: 2]
    }
}

startServer();[cite: 2]
