const express = require('express');
const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs'); // 引入 Excel 處理套件

const app = express();
const PORT = process.env.PORT || 3000; // 支援 Render 自動分配的 PORT

// 允許解析 JSON 與表單資料
app.use(express.json());
app.use(express.urlencoded({ extended: true }));
// 靜態網頁檔案導向 public 資料夾
app.use(express.static(path.join(__dirname, 'public')));

// 🎯 改用跨平台的相對路徑 (相容 Render Linux 環境)
const TARGET_DIR = path.join(__dirname, 'public');
const JSON_FILE = path.join(TARGET_DIR, 'orders.json');
const EXCEL_FILE = path.join(TARGET_DIR, '訂單統計表.xlsx');
// 🪪 員工資料庫實體 JSON 檔案路徑
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

// 安全檢查函式：確保目標資料夾、users.json 與 orders.json 存在
function ensureDirectoryExistence() {
    if (!fs.existsSync(TARGET_DIR)) {
        fs.mkdirSync(TARGET_DIR, { recursive: true });
        console.log(`[系統提示] 已自動建立目標資料夾：${TARGET_DIR}`);
    }
    // 檢查 users.json 是否存在
    if (!fs.existsSync(USER_DB_FILE)) {
        fs.writeFileSync(USER_DB_FILE, JSON.stringify(INITIAL_USER_DB, null, 2), 'utf-8');
        console.log(`[系統提示] 已自動建立員工資料庫檔案：${USER_DB_FILE}`);
    }
    // 檢查 orders.json 是否存在，若不存在則自動初始化為空陣列 []
    if (!fs.existsSync(JSON_FILE)) {
        fs.writeFileSync(JSON_FILE, JSON.stringify([], null, 2), 'utf-8');
        console.log(`[系統提示] 已自動建立訂單檔案：${JSON_FILE}`);
    }
}

// 初始化資料夾與檔案
ensureDirectoryExistence();

// 輔助函式：安全解析 JSON 檔案
function safeReadJSON(filePath, fallback = []) {
    ensureDirectoryExistence();
    try {
        const raw = fs.readFileSync(filePath, 'utf-8').trim();
        return raw ? JSON.parse(raw) : fallback;
    } catch (e) {
        console.error(`[檔案讀取錯誤] 解析 ${filePath} 失敗:`, e.message);
        return fallback;
    }
}

// 輔助函式：從實體檔案讀取最新的員工名單
function loadUserDatabase() {
    const data = safeReadJSON(USER_DB_FILE, INITIAL_USER_DB);
    return (data && typeof data === 'object' && !Array.isArray(data)) ? data : INITIAL_USER_DB;
}

// 輔助函式：將最新的員工名單寫入實體檔案
function saveUserDatabase(db) {
    ensureDirectoryExistence();
    fs.writeFileSync(USER_DB_FILE, JSON.stringify(db, null, 2), 'utf-8');
}

// 輔助函式：判斷訂單時間是否為台灣時間的「今天」
function isTodayInTaipei(orderTimestampMs) {
    const orderDate = new Date(Number(orderTimestampMs));
    if (isNaN(orderDate.getTime())) return false;

    const now = new Date();
    const options = { timeZone: 'Asia/Taipei', year: 'numeric', month: '2-digit', day: '2-digit' };
    
    const orderDateStr = new Intl.DateTimeFormat('zh-TW', options).format(orderDate);
    const todayDateStr = new Intl.DateTimeFormat('zh-TW', options).format(now);

    return orderDateStr === todayDateStr;
}

// ==================== 自動生成並美化 Excel 的輔助函式 ====================
async function updateExcel(orders) {
    try {
        ensureDirectoryExistence();
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
            const numTotal = Number(order.total);
            sheet1.addRow({
                orderId: order.orderId,
                timestamp: order.timestamp,
                cardId: order.cardId,
                name: order.name || '未知', 
                meal: order.meal,
                spicy: order.spicy || '無',
                note: order.note || '無',
                total: !isNaN(numTotal) ? numTotal : 0
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
        console.log(`📊 [系統提示] 訂單統計表.xlsx 已自動更新至：${EXCEL_FILE}`);
    } catch (err) {
        console.error('❌ Excel 更新失敗（可能檔案被開啟中）：', err.message);
    }
}

// ==================== API 路由 ====================

// 🌟 API：取得完整 orders.json 資料
app.get('/api/orders', (req, res) => {
    const orders = safeReadJSON(JSON_FILE, []);
    res.json(orders);
});

// ⚙️ API 1：獲取完整的員工清單給後台表格
app.get('/api/employees', (req, res) => {
    const db = loadUserDatabase();
    res.json(db);
});

// ⚙️ API 2：後台新增員工卡號
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

    console.log(`[員工作業] 新增員工成功: ${cleanName} (${cleanCardId})`);
    res.json({ success: true, message: '新增成功' });
});

// ⚙️ API 3：後台刪除員工卡號
app.delete('/api/employees/:cardId', (req, res) => {
    const { cardId } = req.params;
    const cleanCardId = cardId ? String(cardId).trim() : '';

    const db = loadUserDatabase();
    if (!db[cleanCardId]) {
        return res.status(404).json({ success: false, message: '找不到該卡號的員工' });
    }

    const removedName = db[cleanCardId];
    delete db[cleanCardId];
    saveUserDatabase(db);

    console.log(`[員工作業] 刪除員工成功: ${removedName} (${cleanCardId})`);
    res.json({ success: true, message: '刪除成功' });
});

// ⚙️ API 4：後台刪除特定訂單
app.delete('/api/orders/:orderId', async (req, res) => {
    const { orderId } = req.params;
    const targetOrderIdStr = orderId ? String(orderId).trim() : '';

    if (!targetOrderIdStr) {
        return res.status(400).json({ success: false, message: '無效的訂單編號！' });
    }

    try {
        let orders = safeReadJSON(JSON_FILE, []);
        const orderExists = orders.some(order => String(order.orderId).trim() === targetOrderIdStr);
        
        if (!orderExists) {
            return res.status(404).json({ success: false, message: `找不到訂單編號: ${targetOrderIdStr}` });
        }

        const updatedOrders = orders.filter(order => String(order.orderId).trim() !== targetOrderIdStr);

        fs.writeFileSync(JSON_FILE, JSON.stringify(updatedOrders, null, 2), 'utf-8');
        console.log(`[後台管理] 訂單編號 ${targetOrderIdStr} 已被管理員刪除，已同步至 JSON。`);

        await updateExcel(updatedOrders);

        res.json({ success: true, message: `訂單 ${targetOrderIdStr} 已成功刪除，Excel 亦同步更新。` });

    } catch (error) {
        console.error('後端處理刪除訂單發生錯誤:', error);
        res.status(500).json({ success: false, message: '伺服器刪除資料或覆寫 Excel 失敗。' });
    }
});

// 1. 驗證卡號 API
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

// 2. 接收訂餐資料 API
app.post('/api/order', async (req, res) => {
    const { cardId, meal, note, spicy, total } = req.body;
    const cleanCardId = cardId ? String(cardId).trim() : '';
    
    const db = loadUserDatabase();
    if (!cleanCardId || !db[cleanCardId]) {
        return res.status(400).json({ success: false, message: '卡號無效或未授權，拒絕下單！' });
    }

    const empName = db[cleanCardId];
    const numTotal = Number(total);

    const newOrder = {
        orderId: Date.now(),
        cardId: cleanCardId,
        name: empName, 
        meal: meal ? String(meal).trim() : "未知餐點",
        spicy: spicy ? String(spicy).trim() : "無",
        note: note ? String(note).trim() : "無",
        total: !isNaN(numTotal) ? numTotal : 0,
        timestamp: new Date().toLocaleString('zh-TW', { timeZone: 'Asia/Taipei' })
    };

    try {
        let orders = safeReadJSON(JSON_FILE, []);
        orders.push(newOrder);
        
        fs.writeFileSync(JSON_FILE, JSON.stringify(orders, null, 2), 'utf-8');
        console.log(`[新訂單提示] 收到來自 ${empName} (${cleanCardId}) 的訂單，已同步至 JSON！`);

        updateExcel(orders).catch(e => console.error('同步 Excel 失敗:', e.message));

        res.json({ success: true, message: `🎉 訂單送出成功！` });

    } catch (error) {
        console.error('後端處理訂單發生錯誤:', error);
        res.status(500).json({ success: false, message: '伺服器寫入資料失敗。' });
    }
});

// 3. 獲取當日個人點餐紀錄 API
app.get('/api/order-history', (req, res) => {
    const { cardId } = req.query;
    const cleanCardId = cardId ? String(cardId).trim() : '';

    if (!cleanCardId) {
        return res.json({ success: false, message: '缺少員工卡號', orders: [] });
    }

    try {
        const orders = safeReadJSON(JSON_FILE, []);

        const userOrders = orders
            .filter(order => {
                const isSameUser = String(order.cardId).trim() === cleanCardId;
                // 使用訂單建立時間戳記做精確的「台灣當天」判斷
                const isToday = isTodayInTaipei(order.orderId);
                return isSameUser && isToday;
            })
            .map(order => {
                // 💡 修改處：相容不同格式的 timestamp，直接回傳完整的時間字串避免切割錯誤
                const displayTime = order.timestamp ? String(order.timestamp) : '-';
                
                return {
                    time: displayTime, 
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
    console.log(` 🚀 訂餐系統後端已啟動！Port: ${PORT}`);
    console.log(` 📂 檔案同步路徑設定為：${TARGET_DIR}`);
    console.log(`================================================================`);
});
