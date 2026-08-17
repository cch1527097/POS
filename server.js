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

// 設定檔路徑
const settingsPath = path.join(__dirname, 'public', 'settings.json');

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

// 1. 取得所有訂單
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

// 2. 取得所有員工名單
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

// 5. 新增員工
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

// 9. 新增訂單 (修復版：即時檢查 settings.json 狀態)
app.post('/api/order', async (req, res) => {
    try {
        // === 🔐 檢查系統設定檔 (settings.json) ===
        if (fs.existsSync(settingsPath)) {
            const rawData = await fs.promises.readFile(settingsPath, 'utf-8');
            const settings = JSON.parse(rawData);

            const systemStatus = settings.systemStatus || settings.overrideStatus || 'auto';
            const cutoffTime = settings.cutoffTime || settings.autoCutoffTime || '';

            // 判斷 1: 是否為強制鎖定/停止接單
            if (systemStatus === 'locked' || systemStatus === 'closed' || systemStatus === 'manual_closed') {
                return res.status(403).json({ 
                    success: false, 
                    message: '⛔ 系統目前已設定為「強制鎖定」，暫停接受訂單！' 
                });
            }

            // 判斷 2: 處於自動模式且設定了截止時間
            if (systemStatus === 'auto' && cutoffTime) {
                const now = new Date();
                const twTime = new Date(now.toLocaleString("en-US", { timeZone: "Asia/Taipei" }));
                const currentHour = twTime.getHours();
                const currentMinute = twTime.getMinutes();

                const [cutoffHour, cutoffMinute] = cutoffTime.split(':').map(Number);

                if (currentHour > cutoffHour || (currentHour === cutoffHour && currentMinute >= cutoffMinute)) {
                    return res.status(403).json({ 
                        success: false, 
                        message: `⛔ 今日點餐已於 ${cutoffTime} 截止，無法再送出訂單。` 
                    });
                }
            }
        }

        const { cardId, meal, note, spicy, total } = req.body;
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
            meal ? String(meal).trim() : "未知餐點",
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

// 11. 動態匯出 Excel 下載
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

// 12. 取得系統設定 API
app.get('/api/settings', async (req, res) => {
    try {
        if (fs.existsSync(settingsPath)) {
            const data = await fs.promises.readFile(settingsPath, 'utf-8');
            return res.json(JSON.parse(data));
        }
        res.json({ cutoffTime: '10:30', systemStatus: 'auto' });
    } catch (err) {
        console.error('讀取系統設定失敗:', err.message);
        res.status(500).json({ success: false, message: '無法讀取系統設定' });
    }
});

// 13. 儲存系統設定 API
app.post('/api/settings', async (req, res) => {
    try {
        const { cutoffTime, systemStatus, autoCutoffTime, overrideStatus } = req.body;
        
        const settingsData = {
            cutoffTime: cutoffTime || autoCutoffTime || '10:30',
            systemStatus: systemStatus || overrideStatus || 'auto',
            updatedAt: new Date().toISOString()
        };

        await fs.promises.writeFile(settingsPath, JSON.stringify(settingsData, null, 2), 'utf-8');
        console.log('[系統提示] 系統設定更新成功：', settingsData);

        res.json({ success: true, message: '系統設定儲存成功！', settings: settingsData });
    } catch (err) {
        console.error('儲存系統設定失敗:', err.message);
        res.status(500).json({ success: false, message: '儲存系統設定失敗' });
    }
});

app.listen(PORT, () => {
    console.log(`================================================================`);
    console.log(` 🚀 訂餐系統後端已啟動！Port: ${PORT}`);
    console.log(`================================================================`);
});
