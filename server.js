const express = require('express');
const path = require('path');
const fs = require('fs');
const ExcelJS = require('exceljs');
const { Pool } = require('pg');

const app = express();
const PORT = process.env.PORT || 3000;

app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.use(express.static(path.join(__dirname, 'public')));

// PostgreSQL 連線設定
const pool = new Pool({
    connectionString: process.env.DATABASE_URL,
    ssl: process.env.DATABASE_URL ? { rejectUnauthorized: false } : false
});

// 預設員工資料
const INITIAL_USER_DB = {
    '18445': '李珈豪', '601471': '陳永育', '11110': '林佳蘭', '11744': '施名娟',
    '10069': '許民芳', '13228': '宋筱湄', '12218': '沈佩琪', '10047': '許博捷',
    '6513': '李承州', '16661': '陳育倫', '601473': '李羽茹', '6800': '吳修文',
    '17020': '黃泓耀', '14778': '施憶宣', '13266': '梁薽予', '5514': '張淑娟',
    '9844': '許毓芬', '16696': '賴語婕', '16984': '梁慧如', '15150': '黃珮瑄',
    '16294': '梁婧盈', '16925': '李宜珊', '17528': '曾雅琴', 'TEST': '測試員'
};

// 輔助函式：同步寫入 public/orders.json (相容舊後台)
async function syncOrdersJsonFile() {
    try {
        const result = await pool.query(
            'SELECT order_id AS "orderId", card_id AS "cardId", name, meal, spicy, note, total, timestamp FROM orders ORDER BY order_id DESC;'
        );
        const jsonPath = path.join(__dirname, 'public', 'orders.json');
        fs.writeFileSync(jsonPath, JSON.stringify(result.rows, null, 2), 'utf-8');
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
                name VARCHAR(100) NOT NULL
            );
        `);

        const userCheck = await pool.query('SELECT COUNT(*) FROM users;');
        if (parseInt(userCheck.rows[0].count, 10) === 0) {
            console.log('[系統提示] 正在初始化預設員工名單至 PostgreSQL...');
            for (const [cardId, name] of Object.entries(INITIAL_USER_DB)) {
                await pool.query(
                    'INSERT INTO users (card_id, name) VALUES ($1, $2) ON CONFLICT DO NOTHING;',
                    [cardId, name]
                );
            }
        }

        console.log(`[系統提示] PostgreSQL 資料表與員工資料檢查完成！`);
        
        // 啟動時即同步一次 JSON 檔
        await syncOrdersJsonFile();
    } catch (err) {
        console.error(`❌ 初始化 PostgreSQL 資料庫失敗：`, err.message);
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
        const result = await pool.query(
            'SELECT order_id AS "orderId", card_id AS "cardId", name, meal, spicy, note, total, timestamp FROM orders ORDER BY order_id DESC;'
        );
        res.json(result.rows);
    } catch (err) {
        console.error('獲取訂單失敗:', err.message);
        res.status(500).json({ success: false, message: '無法取得訂單資料' });
    }
});

// 2. 取得所有員工名單
app.get('/api/employees', async (req, res) => {
    try {
        const result = await pool.query('SELECT card_id, name FROM users;');
        const dbMap = {};
        result.rows.forEach(row => {
            dbMap[row.card_id] = row.name;
        });
        res.json(dbMap);
    } catch (err) {
        console.error('讀取員工資料失敗:', err.message);
        res.status(500).json({ success: false, message: '無法取得員工資料' });
    }
});

// 3. 新增員工
app.post('/api/employees', async (req, res) => {
    const { cardId, name } = req.body;
    const cleanCardId = cardId ? String(cardId).trim() : '';
    const cleanName = name ? String(name).trim() : '';

    if (!cleanCardId || !cleanName) {
        return res.status(400).json({ success: false, message: '卡號與姓名不可為空！' });
    }

    try {
        const check = await pool.query('SELECT card_id FROM users WHERE card_id = $1;', [cleanCardId]);
        if (check.rows.length > 0) {
            return res.status(400).json({ success: false, message: '此卡號已經存在！' });
        }

        await pool.query('INSERT INTO users (card_id, name) VALUES ($1, $2);', [cleanCardId, cleanName]);
        res.json({ success: true, message: '新增成功' });
    } catch (err) {
        console.error('新增員工失敗:', err.message);
        res.status(500).json({ success: false, message: '新增員工失敗' });
    }
});

// 4. 刪除員工
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

// 5. 刪除訂單
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

        // 同步更新 orders.json 檔案
        await syncOrdersJsonFile();

        res.json({ success: true, message: `訂單 ${targetOrderIdStr} 已成功刪除。` });
    } catch (error) {
        console.error('後端處理刪除訂單發生錯誤:', error);
        res.status(500).json({ success: false, message: '伺服器刪除資料失敗。' });
    }
});

// 6. 登入驗證
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

// 7. 新增訂單
app.post('/api/order', async (req, res) => {
    const { cardId, meal, note, spicy, total } = req.body;
    const cleanCardId = cardId ? String(cardId).trim() : '';
    
    try {
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
        console.log(`[新訂單提示] 收到來自 ${empName} (${cleanCardId}) 的訂單，已寫入 PostgreSQL！`);

        // 同步更新 orders.json 檔案供舊版前端讀取
        await syncOrdersJsonFile();

        res.json({ success: true, message: `🎉 訂單送出成功！` });
    } catch (error) {
        console.error('後端處理訂單發生錯誤:', error);
        res.status(500).json({ success: false, message: '伺服器寫入資料庫失敗。' });
    }
});

// 8. 取得個人歷史訂單紀錄
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

// 9. 動態匯出 Excel 下載
app.get('/api/export-excel', async (req, res) => {
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
