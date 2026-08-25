require('dotenv').config(); // 載入 .env 環境變數
const express = require('express');
const path = require('path');
const fs = require('fs');
const ExcelJS = require('exceljs');
const axios = require('axios');
const { Pool, types } = require('pg');

// 自動將 PostgreSQL BIGINT (OID 20) 解析為 JavaScript String (避免精度遺失)
types.setTypeParser(20, val => val);

const app = express();
const PORT = process.env.PORT || 3000;

// LINE Messaging API 設定
const LINE_CHANNEL_ACCESS_TOKEN = process.env.LINE_CHANNEL_ACCESS_TOKEN || 'YOUR_LINE_CHANNEL_ACCESS_TOKEN';
const TARGET_GROUP_ID = process.env.LINE_TARGET_ID || 'YOUR_TARGET_GROUP_OR_USER_ID';

app.use(express.json({ limit: '1mb' }));
app.use(express.urlencoded({ extended: true, limit: '1mb' }));
app.use(express.static(path.join(__dirname, 'public')));

// PostgreSQL (Neon) 連線設定
const pool = new Pool({
    connectionString: process.env.DATABASE_URL,
    ssl: {
        rejectUnauthorized: false // Neon 強制使用 SSL
    },
    max: 10,
    idleTimeoutMillis: 30000,
    connectionTimeoutMillis: 5000
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

// 輔助函式：發送 LINE 訊息
async function sendLinePushMessage(textMessage) {
    if (!LINE_CHANNEL_ACCESS_TOKEN || LINE_CHANNEL_ACCESS_TOKEN === 'YOUR_LINE_CHANNEL_ACCESS_TOKEN') {
        console.warn('⚠️ 未設定有效的 LINE_CHANNEL_ACCESS_TOKEN，跳過 LINE 通知發送。');
        return false;
    }
    try {
        await axios.post(
            'https://api.line.me/v2/bot/message/push',
            {
                to: TARGET_GROUP_ID,
                messages: [{ type: 'text', text: textMessage }]
            },
            {
                headers: {
                    'Content-Type': 'application/json',
                    'Authorization': `Bearer ${LINE_CHANNEL_ACCESS_TOKEN}`
                }
            }
        );
        return true;
    } catch (error) {
        console.error('LINE Messaging API 呼叫失敗:', error.response?.data || error.message);
        return false;
    }
}

// 輔助函式：同步寫入 public/orders.json
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
            ORDER BY id DESC;
        `);
        const jsonPath = path.join(publicDir, 'orders.json');
        await fs.promises.writeFile(jsonPath, JSON.stringify(result.rows, null, 2), 'utf-8');
    } catch (err) {
        console.warn('⚠️ 警告：同步 orders.json 失敗 (雲端環境可能為唯讀檔案系統):', err.message);
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
                is_paid BOOLEAN DEFAULT FALSE,
                created_at TIMESTAMP WITH TIME ZONE DEFAULT CURRENT_TIMESTAMP
            );
        `);

        await pool.query(`
            ALTER TABLE orders ADD COLUMN IF NOT EXISTS is_paid BOOLEAN DEFAULT FALSE;
        `);

        await pool.query(`
            CREATE TABLE IF NOT EXISTS users (
                card_id VARCHAR(50) PRIMARY KEY,
                name VARCHAR(100) NOT NULL,
                is_paid BOOLEAN DEFAULT FALSE,
                department VARCHAR(100) DEFAULT '未劃分'
            );
        `);

        await pool.query(`
            CREATE TABLE IF NOT EXISTS settings (
                key VARCHAR(50) PRIMARY KEY,
                value TEXT NOT NULL
            );
        `);

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
        ];

        const currentStoreRes = await pool.query("SELECT value FROM settings WHERE key = 'store_list';");
        if (currentStoreRes.rows.length === 0) {
            await pool.query(`
                INSERT INTO settings (key, value)
                VALUES ('store_list', $1);
            `, [JSON.stringify(DEFAULT_STORES)]);
        } else {
            try {
                let existingStores = JSON.parse(currentStoreRes.rows[0].value) || [];
                if (!Array.isArray(existingStores)) existingStores = [];
                
                const syncedStores = DEFAULT_STORES.map(defStore => {
                    const match = existingStores.find(s => s && s.id === defStore.id);
                    return {
                        ...defStore,
                        isOpen: match ? Boolean(match.isOpen) : defStore.isOpen
                    };
                });

                await pool.query(`
                    INSERT INTO settings (key, value)
                    VALUES ('store_list', $1)
                    ON CONFLICT (key) DO UPDATE SET value = EXCLUDED.value;
                `, [JSON.stringify(syncedStores)]);
            } catch (e) {
                console.error('解析現有店家清單失敗，重新寫入預設店家:', e);
                await pool.query(`
                    INSERT INTO settings (key, value)
                    VALUES ('store_list', $1)
                    ON CONFLICT (key) DO UPDATE SET value = EXCLUDED.value;
                `, [JSON.stringify(DEFAULT_STORES)]);
            }
        }

        await pool.query(`
            ALTER TABLE users ADD COLUMN IF NOT EXISTS is_paid BOOLEAN DEFAULT FALSE;
            ALTER TABLE users ADD COLUMN IF NOT EXISTS department VARCHAR(100) DEFAULT '未劃分';
        `);

        const userCheck = await pool.query('SELECT COUNT(*) FROM users;');
        if (parseInt(userCheck.rows[0].count, 10) === 0) {
            console.log('[系統提示] 正在初始化預設員工名單至 Neon PostgreSQL...');
            const cardIds = Object.keys(INITIAL_USER_DB);
            const names = Object.values(INITIAL_USER_DB);

            // 修正：修正 UNNEST 語法，明確指定常數資料型態
            await pool.query(`
                INSERT INTO users (card_id, name, is_paid, department)
                SELECT t.card_id, t.name, false, '未劃分'
                FROM UNNEST($1::text[], $2::text[]) AS t(card_id, name)
                ON CONFLICT (card_id) DO NOTHING;
            `, [cardIds, names]);
        }

        console.log(`[系統提示] Neon PostgreSQL 資料表與員工資料檢查完成！`);
        await syncOrdersJsonFile();
    } catch (err) {
        console.error(`❌ 初始化 Neon 資料庫失敗：`, err.message);
        throw err;
    }
}

// ==================== API 路由 ====================

// 手動發送 LINE 訊息 API
app.post('/api/send-line-message', async (req, res) => {
    const { message } = req.body;

    if (!message) {
        return res.status(400).json({ success: false, message: '訊息內容不可為空' });
    }

    const success = await sendLinePushMessage(message);
    if (success) {
        return res.json({ success: true, message: '訊息已推送至 LINE 群組' });
    } else {
        return res.status(500).json({ success: false, message: 'LINE API 發送失敗' });
    }
});

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
            return res.json({ 
                success: true, 
                message: '已成功解除點餐限制！目前系統無時間限制，隨時可進行下單。',
                lockTime: "" 
            });
        }

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
        await pool.query(`
            INSERT INTO settings (key, value) 
            VALUES ('store_list', $1)
            ON CONFLICT (key) DO UPDATE SET value = EXCLUDED.value;
        `, [JSON.stringify(stores)]);

        const activeStores = stores.filter(s => s && s.isOpen);
        const activeStoreNames = activeStores
            .map(s => s.name.replace(/\s*\([^)]*\)/g, '').trim())
            .join('、');

        await pool.query(`
            INSERT INTO settings (key, value) 
            VALUES ('active_store', $1)
            ON CONFLICT (key) DO UPDATE SET value = EXCLUDED.value;
        `, [activeStoreNames]);

        return res.json({ 
            success: true, 
            message: activeStoreNames ? `已成功將開放店家設定為：${activeStoreNames}` : '已變更店家設定，目前無開放店家！',
            activeStore: activeStoreNames,
            stores: stores
        });
    } catch (err) {
        console.error('更新開放店家失敗:', err.message);
        res.status(500).json({ success: false, message: '設定開放店家失敗' });
    }
});

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
            ORDER BY id DESC;
        `);
        res.json(result.rows);
    } catch (err) {
        console.error('獲取訂單失敗:', err.message);
        res.status(500).json({ success: false, message: '無法取得訂單資料' });
    }
});

app.patch('/api/orders/:orderId/payment', async (req, res) => {
    const { orderId } = req.params;
    const { isPaid } = req.body;
    const targetOrderIdStr = orderId ? String(orderId).trim() : '';
    const paidBool = isPaid === true || isPaid === 'true';

    if (!targetOrderIdStr) {
        return res.status(400).json({ success: false, message: '無效的訂單編號！' });
    }

    try {
        const result = await pool.query(
            'UPDATE orders SET is_paid = $1 WHERE order_id = $2 RETURNING *;',
            [paidBool, targetOrderIdStr]
        );
        
        if (result.rowCount === 0) {
            return res.status(404).json({ success: false, message: `找不到訂單編號: ${targetOrderIdStr}` });
        }

        syncOrdersJsonFile().catch(() => {});

        res.json({ success: true, message: '訂單繳費狀態更新成功！' });
    } catch (err) {
        console.error('更新訂單繳費狀態失敗:', err.message);
        res.status(500).json({ success: false, message: '伺服器更新繳費狀態失敗。' });
    }
});

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

        syncOrdersJsonFile().catch(() => {});
        res.json({ success: true, message: '繳費狀態更新成功' });
    } catch (err) {
        console.error('更新員工繳費狀態失敗:', err.message);
        res.status(500).json({ success: false, message: '更新失敗' });
    }
});

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
        
        syncOrdersJsonFile().catch(() => {});
        res.json({ success: true, message: '員工資料更新成功' });
    } catch (err) {
        console.error('更新員工資料失敗:', err.message);
        res.status(500).json({ success: false, message: '更新員工資料失敗' });
    }
});

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

        syncOrdersJsonFile().catch(() => {});
        res.json({ success: true, message: `訂單 ${targetOrderIdStr} 已成功刪除。` });
    } catch (error) {
        console.error('後端處理刪除訂單發生錯誤:', error);
        res.status(500).json({ success: false, message: '伺服器刪除資料失敗。' });
    }
});

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

// 下訂單 API
app.post('/api/order', async (req, res) => {
    try {
        const { cardId, meal, storeId, storeName, note, spicy, total } = req.body;
        const cleanMeal = meal ? String(meal).trim() : "";
        const reqStoreId = storeId ? String(storeId).trim() : "";
        const reqStoreName = storeName ? String(storeName).trim() : "";

        // 1. 店家開放狀態檢查
        const storeListRes = await pool.query("SELECT value FROM settings WHERE key = 'store_list';");
        
        if (storeListRes.rows.length > 0 && storeListRes.rows[0].value) {
            const stores = JSON.parse(storeListRes.rows[0].value) || [];
            let matchedStore = null;

            if (reqStoreId) {
                matchedStore = stores.find(s => s && s.id === reqStoreId);
            }

            if (!matchedStore) {
                matchedStore = stores.find(s => {
                    if (!s || !s.name) return false;
                    const cleanName = s.name.replace(/\s*\([^)]*\)/g, '').trim();
                    if (!cleanName) return false;
                    if (reqStoreName && (reqStoreName.includes(cleanName) || cleanName.includes(reqStoreName))) return true;
                    return cleanMeal.startsWith(cleanName);
                });
            }

            if (matchedStore && !matchedStore.isOpen) {
                const displayStoreName = matchedStore.name.replace(/\s*\([^)]*\)/g, '').trim();
                return res.status(403).json({
                    success: false,
                    message: `⚠️ 店家「${displayStoreName}」目前尚未開放點餐！`
                });
            }
        }

        // 2. 鎖定時間比對
        const lockRes = await pool.query("SELECT value FROM settings WHERE key = 'order_lock_time';");
        const orderLockTime = lockRes.rows.length > 0 ? lockRes.rows[0].value : "";

        if (orderLockTime && orderLockTime.includes(':')) {
            const now = new Date();
            const taipeiFormatter = new Intl.DateTimeFormat('zh-TW', {
                timeZone: 'Asia/Taipei',
                hour: '2-digit',
                minute: '2-digit',
                hour12: false
            });
            const parts = taipeiFormatter.formatToParts(now);
            const currentH = Number(parts.find(p => p.type === 'hour').value);
            const currentM = Number(parts.find(p => p.type === 'minute').value);
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

        // 3. 員工驗證與寫入
        const cleanCardId = cardId ? String(cardId).trim() : '';
        const userRes = await pool.query('SELECT name FROM users WHERE card_id = $1;', [cleanCardId]);
        if (!cleanCardId || userRes.rows.length === 0) {
            return res.status(400).json({ success: false, message: '卡號無效或未授權，拒絕下單！' });
        }

        const empName = userRes.rows[0].name;
        const numTotal = Number(total);
        
        const safeTimestamp = String(Date.now()).slice(-9);
        const randomNum = String(Math.floor(100 + Math.random() * 900));
        const orderIdVal = `${safeTimestamp}${randomNum}`;

        const timestampStr = new Date().toLocaleString('zh-TW', { timeZone: 'Asia/Taipei' });

        const insertQuery = `
            INSERT INTO orders (order_id, card_id, name, meal, spicy, note, total, timestamp, is_paid)
            VALUES ($1, $2, $3, $4, $5, $6, $7, $8, false);
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
        syncOrdersJsonFile().catch(() => {});

        // 下單成功後自動發送 LINE 群組通知
        const lineMsg = `🍱 【新訂單通知】\n姓名：${empName} (${cleanCardId})\n餐點：${cleanMeal}\n辣度：${spicy || '無'}\n備註：${note || '無'}\n金額：$${numTotal || 0}`;
        sendLinePushMessage(lineMsg).catch(() => {});

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
        // 修正：使用 NOW() - INTERVAL '24 hours' 來精準計算過去 24 小時的紀錄，避開 AT TIME ZONE 的轉型偏差
        const result = await pool.query(
            `SELECT order_id, meal, spicy, note, total, timestamp 
             FROM orders 
             WHERE card_id = $1 
               AND created_at >= NOW() - INTERVAL '24 hours'
             ORDER BY id DESC;`,
            [cleanCardId]
        );

        const userOrders = result.rows.map(order => ({
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

// Excel 匯出
app.get('/api/export-excel', async (req, res) => {
    try {
        const result = await pool.query(`
            SELECT 
                order_id, timestamp, card_id, name, meal, spicy, note, total,
                COALESCE(is_paid, false) AS is_paid
            FROM orders
            ORDER BY id ASC;
        `);
        
        const orders = result.rows.map(row => ({
            orderId: String(row.order_id),
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
            { header: '訂單ID', key: 'orderId', width: 22 },
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

// 啟動伺服器
async function startServer() {
    try {
        await initDatabase();
        app.listen(PORT, () => {
            console.log(`================================================================`);
            console.log(` 🚀 訂餐系統後端已啟動！Port: ${PORT}`);
            console.log(`================================================================`);
        });
    } catch (err) {
        console.error('❌ 伺服器啟動失敗，無法連線至 Neon PostgreSQL:', err);
        process.exit(1);
    }
}

startServer();
