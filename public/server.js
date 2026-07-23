const express = require('express');
const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs'); // 引入 Excel 處理套件

const app = express();
const PORT = 3000;

// 允許解析 JSON 與表單資料
app.use(express.json());
app.use(express.urlencoded({ extended: true }));
// 靜態網頁檔案導向 public 資料夾
app.use(express.static('public'));

// 🎯 設定要同步的目標絕對路徑
const TARGET_DIR = path.join('C:', 'Users', '18445', 'Desktop', '李珈豪', 'pos-system', 'public');
const JSON_FILE = path.join(TARGET_DIR, 'orders.json');
const EXCEL_FILE = path.join(TARGET_DIR, '訂單統計表.xlsx');
// 🪪 新增：員工資料庫實體 JSON 檔案路徑
const USER_DB_FILE = path.join(TARGET_DIR, 'users.json');

// 🌟 初始員工資料庫（若 users.json 不存在時的預設值）
const INITIAL_USER_DB = {
    '18445': '李珈豪', '601471': '陳永育', '11110': '林佳蘭', '11744': '施名娟',
    '10069': '許民芳', '13228': '宋筱湄', '12218': '沈佩琪', '10047': '許博捷',
    '6513': '李承州', '16661': '陳育倫', '601473': '李羽茹', '6800': '吳修文',
    '17020': '黃泓耀', '14778': '施憶宣', '13266': '梁薽予', '5514': '張淑娟',
    '9844': '許毓芬', '16696': '賴語婕', '16984': '梁慧如', '15150': '黃珮瑄',
    '16294': '梁婧盈', '16925': '李宜珊', '17528': '曾雅琴', 'TEST': '測試員'
};

// 安全檢查函式：確保目標資料夾存在，並確保員工資料庫檔案存在
function ensureDirectoryExistence() {
    if (!fs.existsSync(TARGET_DIR)) {
        fs.mkdirSync(TARGET_DIR, { recursive: true });
        console.log(`[系統提示] 已自動建立目標資料夾：${TARGET_DIR}`);
    }
    // 檢查 users.json 是否存在，若不存在則建立初始檔案
    if (!fs.existsSync(USER_DB_FILE)) {
        fs.writeFileSync(USER_DB_FILE, JSON.stringify(INITIAL_USER_DB, null, 2), 'utf-8');
        console.log(`[系統提示] 已自動建立員工資料庫檔案：${USER_DB_FILE}`);
    }
}

// 輔助函式：從實體檔案讀取最新的員工名單
function loadUserDatabase() {
    ensureDirectoryExistence();
    try {
        const data = fs.readFileSync(USER_DB_FILE, 'utf-8');
        return JSON.parse(data);
    } catch (e) {
        console.error('讀取員工資料庫檔案失敗，使用內建資料:', e.message);
        return INITIAL_USER_DB;
    }
}

// 輔助函式：將最新的員工名單寫入實體檔案
function saveUserDatabase(db) {
    ensureDirectoryExistence();
    fs.writeFileSync(USER_DB_FILE, JSON.stringify(db, null, 2), 'utf-8');
}

// ==================== 自動生成並美化 Excel 的輔助函式 ====================
async function updateExcel(orders) {
    try {
        ensureDirectoryExistence(); // 確保路徑資料夾存在
        const workbook = new ExcelJS.Workbook();
        
        // 1. 建立第一個分頁：訂單明細
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
        
        // 2. 建立第二個分頁：店家點餐統計
        const sheet2 = workbook.addWorksheet('店家點餐統計表');
        sheet2.columns = [
            { header: '店家/品項名稱', key: 'meal', width: 25 },
            { header: '總點餐次數', key: 'count', width: 15 }
        ];

        // 寫入明細資料
        orders.forEach(order => {
            sheet1.addRow({
                orderId: order.orderId,
                timestamp: order.timestamp,
                cardId: order.cardId,
                name: order.name || '未知', 
                meal: order.meal,
                spicy: order.spicy || '無',
                note: order.note || '無',
                total: order.total !== undefined ? Number(order.total) : 0
            });
        });

        // 統計各店家的數量
        const statistics = {};
        orders.forEach(order => {
            if (order.meal) {
                statistics[order.meal] = (statistics[order.meal] || 0) + 1;
            }
        });

        // 寫入統計資料
        Object.keys(statistics).forEach(meal => {
            sheet2.addRow({
                meal: meal,
                count: statistics[meal]
            });
        });

        // 美化 Excel 樣式
        [sheet1, sheet2].forEach(sheet => {
            // 標題列樣式 (第 1 列)
            const headerRow = sheet.getRow(1);
            headerRow.font = { bold: true, color: { argb: 'FFFFFF' }, size: 11 };
            headerRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '2B4C23' } };
            
            // 全格線與置中
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
        console.log(`📊 [系統提示] 訂單統計表.xlsx 已自動更新至：${EXCEL_FILE}`);
    } catch (err) {
        console.error('❌ Excel 更新失敗（可能檔案被開啟中）：', err.message);
    }
}

// ==================== API 路由 ====================

// ⚙️ 新增 API 1：獲取完整的員工清單給後台表格
app.get('/api/employees', (req, res) => {
    const db = loadUserDatabase();
    res.json(db);
});

// ⚙️ 新增 API 2：後台新增員工卡號
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
    saveUserDatabase(db); // 寫入 users.json 永久儲存

    console.log(`[員工作業] 新增員工成功: ${cleanName} (${cleanCardId})`);
    res.json({ success: true, message: '新增成功' });
});

// ⚙️ 新增 API 3：後台刪除員工卡號
app.delete('/api/employees/:cardId', (req, res) => {
    const cardId = req.params.cardId;
    const cleanCardId = cardId ? String(cardId).trim() : '';

    const db = loadUserDatabase();
    if (!db[cleanCardId]) {
        return res.status(404).json({ success: false, message: '找不到該卡號的員工' });
    }

    const removedName = db[cleanCardId];
    delete db[cleanCardId];
    saveUserDatabase(db); // 寫入 users.json 永久儲存

    console.log(`[員工作業] 刪除員工成功: ${removedName} (${cleanCardId})`);
    res.json({ success: true, message: '刪除成功' });
});

// ⚙️ 新增 API 4：後台刪除特定訂單（連帶自動更新 JSON 與 Excel）
app.delete('/api/orders/:orderId', async (req, res) => {
    const { orderId } = req.params;
    const targetOrderId = orderId ? Number(orderId) : null;

    if (!targetOrderId) {
        return res.status(400).json({ success: false, message: '無效的訂單編號！' });
    }

    try {
        ensureDirectoryExistence(); // 確保路徑存在

        // 1. 檢查並讀取 orders.json
        if (!fs.existsSync(JSON_FILE)) {
            return res.status(404).json({ success: false, message: '目前尚無任何訂單紀錄檔案。' });
        }

        let orders = JSON.parse(fs.readFileSync(JSON_FILE, 'utf-8'));

        // 2. 檢查該訂單是否存在
        const orderExists = orders.some(order => Number(order.orderId) === targetOrderId);
        if (!orderExists) {
            return res.status(404).json({ success: false, message: `找不到訂單編號: ${targetOrderId}` });
        }

        // 3. 過濾掉被刪除的那筆訂單
        const updatedOrders = orders.filter(order => Number(order.orderId) !== targetOrderId);

        // 4. 全新陣列覆寫回 orders.json
        fs.writeFileSync(JSON_FILE, JSON.stringify(updatedOrders, null, 2), 'utf-8');
        console.log(`[後台管理] 訂單編號 ${targetOrderId} 已被管理員刪除，已同步至 JSON。`);

        // 5. 自動更新桌面 Excel 中的明細表與店家統計表
        await updateExcel(updatedOrders);

        // 6. 回傳成功狀態給前端
        res.json({ success: true, message: `訂單 ${targetOrderId} 已成功刪除，Excel 亦同步更新。` });

    } catch (error) {
        console.error('後端處理刪除訂單發生錯誤:', error);
        res.status(500).json({ success: false, message: '伺服器刪除資料或覆寫 Excel 失敗。' });
    }
});

// 1. 驗證卡號 API（已改為讀取實體檔案）
app.post('/api/login', (req, res) => {
    const { cardId } = req.body;
    const cleanCardId = cardId ? String(cardId).trim() : '';
    
    const db = loadUserDatabase();
    if (db[cleanCardId]) {
        res.json({ 
            success: true, 
            message: '登入成功', 
            empName: db[cleanCardId] 
        });
    } else {
        res.json({ success: false, message: '卡號無效，拒絕存取' });
    }
});

// 2. 接收訂餐資料 API（已改為讀取實體檔案）
app.post('/api/order', async (req, res) => {
    const { cardId, meal, note, spicy, total } = req.body;
    const cleanCardId = cardId ? String(cardId).trim() : '';
    
    const db = loadUserDatabase();
    const empName = db[cleanCardId] || "未知員工";

    const newOrder = {
        orderId: Date.now(),
        cardId: cleanCardId || "未登入卡號",
        name: empName, 
        meal: meal || "未知餐點",
        spicy: spicy || "無",
        note: note || "無",
        total: total !== undefined ? Number(total) : 0,
        timestamp: new Date().toLocaleString('zh-TW', { timeZone: 'Asia/Taipei' })
    };

    let orders = [];

    try {
        ensureDirectoryExistence(); // 確保路徑資料夾存在

        if (fs.existsSync(JSON_FILE)) {
            orders = JSON.parse(fs.readFileSync(JSON_FILE, 'utf-8'));
        }
        
        orders.push(newOrder);
        fs.writeFileSync(JSON_FILE, JSON.stringify(orders, null, 2), 'utf-8');
        console.log(`[新訂單提示] 收到來自 ${empName} (${cleanCardId}) 的訂單，已同步至 JSON！`);

        // 非同步更新 Excel，不阻塞前端回應
        updateExcel(orders).catch(e => console.error(e));

        res.json({ success: true, message: `🎉 訂單送出成功！` });

    } catch (error) {
        console.error('後端處理訂單發生錯誤:', error);
        res.status(500).json({ success: false, message: '伺服器寫入資料失敗。' });
    }
});

// 3. 獲取當日個人點餐紀錄 API（已改為讀取實體檔案）
app.get('/api/order-history', (req, res) => {
    const { cardId } = req.query;
    const cleanCardId = cardId ? String(cardId).trim() : '';

    if (!cleanCardId) {
        return res.json({ success: false, message: '缺少員工卡號', orders: [] });
    }

    try {
        ensureDirectoryExistence(); // 確保路徑資料夾存在
        
        let orders = [];
        if (fs.existsSync(JSON_FILE)) {
            orders = JSON.parse(fs.readFileSync(JSON_FILE, 'utf-8'));
        }

        // 精確台北時間日期抓取
        const now = new Date();
        const todayStr = new Intl.DateTimeFormat('zh-TW', {
            timeZone: 'Asia/Taipei',
            year: 'numeric', month: 'numeric', day: 'numeric'
        }).format(now);

        // 過濾出屬於該卡號，且時間為今天的紀錄
        const userOrders = orders
            .filter(order => {
                const isSameUser = String(order.cardId).trim() === cleanCardId;
                const isToday = order.timestamp && order.timestamp.includes(todayStr);
                return isSameUser && isToday;
            })
            .map(order => {
                const parts = order.timestamp ? order.timestamp.split(/\s+/) : [];
                const timePart = parts[1] || '-'; 
                
                return {
                    time: timePart, 
                    meal: order.meal,
                    note: `醬料辣度: ${order.spicy} | 備註: ${order.note} | 金額: $${order.total}`
                };
            });

        res.json({
            success: true,
            orders: userOrders
        });

    } catch (error) {
        console.error('讀取紀錄失敗:', error);
        res.json({ success: false, message: '讀取歷史紀錄失敗', orders: [] });
    }
});

app.listen(PORT, () => {
    console.log(`================================================================`);
    console.log(` 🚀 訂餐系統後端已啟動！`);
    console.log(` 🌐 請在瀏覽器輸入 http://localhost:${PORT}`);
    console.log(` 📂 檔案同步路徑設定為：${TARGET_DIR}`);
    console.log(`================================================================`);
});