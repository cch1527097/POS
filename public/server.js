const express = require('express');
const fs = require('fs').promises; 
const path = require('path');
const ExcelJS = require('exceljs');

const app = express();
const PORT = 3000;

app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.use(express.static('public'));

const JSON_FILE = path.join(__dirname, 'orders.json');
const EXCEL_FILE = path.join(__dirname, '訂單統計表.xlsx');
const SHOPS_FILE = path.join(__dirname, 'shops.json'); // 🌟 店家動態資料庫

// ==================== 🛠️ 記憶體與資料檔案初始化 ====================

// 🌟 員工資料庫改用 let，允許後台動態增刪
let USER_DB = {
    '18445': '李珈豪', '601471': '陳永育', '11110': '林佳蘭', '11744': '施名娟',
    '10069': '許民芳', '13228': '宋筱湄', '12218': '沈佩琪', '10047': '許博捷',
    '6513': '李承州', '16661': '陳育倫', '601473': '李羽茹', '6800': '吳修文',
    '17020': '黃泓耀', '14778': '施憶宣', '13266': '梁薽予', '5514': '張淑娟',
    '9844': '許毓芬', '16696': '賴語婕', '16984': '梁慧如', '15150': '黃珮瑄',
    '16294': '梁婧盈', '16925': '李宜珊', '17528': '曾雅琴', 'TEST': '測試員'
};

// 預設合作店家資料
const DEFAULT_SHOPS = [
    { name: "八方雲集", file: "bafang.html", desc: "皮Q餡飽滿的招牌鍋貼與各式經典水餃。" },
    { name: "榮郁香廣式燒臘", file: "dafang.html", desc: "道地廣式燒臘，外酥內嫩的美味首選。" },
    { name: "飯大廚", file: "aafang.html", desc: "大火翻炒粒粒分明，炒飯料理的專家。" }
];

// 確保店家 json 存在
async function initShopsFile() {
    try {
        await fs.access(SHOPS_FILE);
    } catch {
        await fs.writeFile(SHOPS_FILE, JSON.stringify(DEFAULT_SHOPS, null, 2), 'utf-8');
    }
}
initShopsFile();

// ==================== 🔒 併發排隊管理佇列 ====================
let fileWriteQueue = Promise.resolve();
function queueFileWrite(operation) {
    fileWriteQueue = fileWriteQueue.then(operation).catch(err => console.error("📁 JSON 佇列執行錯誤:", err));
    return fileWriteQueue;
}

let excelWriteQueue = Promise.resolve();
function queueExcelWrite(operation) {
    excelWriteQueue = excelWriteQueue.then(operation).catch(err => console.error("📊 Excel 佇列執行錯誤:", err));
    return excelWriteQueue;
}

// 台灣時間標準化字串工具
function getTaiwanTimeParts() {
    const now = new Date();
    const twTime = new Date(now.getTime() + (8 * 60 * 60 * 1000));
    const isoStr = twTime.toISOString();
    return {
        datePart: isoStr.split('T')[0],
        timePart: isoStr.split('T')[1].substring(0, 8),
        fullLog: `${isoStr.split('T')[0]} ${isoStr.split('T')[1].substring(0, 8)}`
    };
}

// ==================== 📊 Excel 產生器 ====================
async function updateExcel(orders) {
    return queueExcelWrite(async () => {
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
                orderId: order.orderId, timestamp: order.timestamp, cardId: order.cardId,
                name: order.name || '未知', meal: order.meal, spicy: order.spicy || '無',
                note: order.note || '無', total: order.total !== undefined ? order.total : 0
            });
        });

        const statistics = {};
        orders.forEach(order => { statistics[order.meal] = (statistics[order.meal] || 0) + 1; });
        Object.keys(statistics).forEach(meal => { sheet2.addRow({ meal: meal, count: statistics[meal] }); });

        [sheet1, sheet2].forEach(sheet => {
            sheet.getRow(1).font = { bold: true, color: { argb: 'FFFFFF' } };
            sheet.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '2B4C23' } };
            sheet.eachRow((row) => {
                row.eachCell(cell => {
                    cell.alignment = { vertical: 'middle', horizontal: 'center' };
                    cell.border = {
                        top: { style: 'thin', color: { argb: 'DDDDDD' } }, bottom: { style: 'thin', color: { argb: 'DDDDDD' } },
                        left: { style: 'thin', color: { argb: 'DDDDDD' } }, right: { style: 'thin', color: { argb: 'DDDDDD' } }
                    };
                });
            });
        });

        try {
            await workbook.xlsx.writeFile(EXCEL_FILE);
            console.log('📊 [系統提示] 訂單統計表.xlsx 已自動更新！');
        } catch (err) {
            console.error('❌ Excel 寫入鎖定中，可能有人開啟了檔案。安全跳過本次更新。');
        }
    });
}

// ==================== 👤 前端員工端 API ====================

// 1. 驗證卡號 API
app.post('/api/login', (req, res) => {
    const { cardId } = req.body;
    if (USER_DB[cardId]) {
        res.json({ success: true, message: '登入成功', empName: USER_DB[cardId] });
    } else {
        res.json({ success: false, message: '卡號無效，拒絕存取' });
    }
});

// 2. 接收訂餐資料 API
app.post('/api/order', async (req, res) => {
    const { cardId, meal, note, spicy, total } = req.body;
    const empName = USER_DB[cardId] || "未知員工";
    const { fullLog } = getTaiwanTimeParts();

    const newOrder = {
        orderId: Date.now(),
        cardId: cardId || "未登入卡號",
        name: empName, 
        meal: meal || "未知餐點",
        spicy: spicy || "無",
        note: note || "無",
        total: total !== undefined ? Number(total) : 0,
        timestamp: fullLog
    };

    queueFileWrite(async () => {
        let orders = [];
        try {
            await fs.access(JSON_FILE);
            const fileData = await fs.readFile(JSON_FILE, 'utf-8');
            orders = JSON.parse(fileData || '[]');
        } catch (e) { orders = []; }
        
        orders.push(newOrder);
        await fs.writeFile(JSON_FILE, JSON.stringify(orders, null, 2), 'utf-8');
        console.log(`[新訂單提示] 收到來自 ${empName} (${cardId}) 的訂單！`);
        updateExcel(orders);
    });

    res.json({ success: true, message: `🎉 訂單送出成功！` });
});

// 3. 獲取當日個人點餐紀錄 API
app.get('/api/order-history', async (req, res) => {
    const { cardId } = req.query;
    if (!cardId) return res.json({ success: false, message: '缺少員工卡號' });

    try {
        let orders = [];
        try {
            const fileData = await fs.readFile(JSON_FILE, 'utf-8');
            orders = JSON.parse(fileData || '[]');
        } catch (e) {}

        const { datePart } = getTaiwanTimeParts();

        const userOrders = orders
            .filter(order => {
                const isSameUser = String(order.cardId).trim() === String(cardId).trim();
                const isToday = order.timestamp && order.timestamp.startsWith(datePart);
                return isSameUser && isToday;
            })
            .map(order => {
                const timePart = order.timestamp ? order.timestamp.split(' ')[1] : '-';
                return {
                    time: timePart,
                    meal: order.meal,
                    note: `醬料辣度: ${order.spicy} | 備註: ${order.note} | 金額: $${order.total}`
                };
            });

        res.json({ success: true, orders: userOrders });
    } catch (error) {
        res.json({ success: false, message: '讀取紀錄失敗', orders: [] });
    }
});


// ==================== ⚙️ 後台管理端 API (全新對接) ====================

// A-1. 後台管理員登入驗證
app.post('/api/admin/login', (req, res) => {
    const { username, password } = req.body;
    // 依據前端 HTML 寫死的預設備用帳密對接
    if (username === '18445' && password === 'zxc456456') {
        res.json({ success: true, message: '驗證成功' });
    } else {
        res.json({ success: false, message: '管理員帳號或密碼錯誤' });
    }
});

// A-2. 獲取所有人當日訂單統計資訊
app.get('/api/admin/orders', async (req, res) => {
    try {
        let orders = [];
        try {
            const fileData = await fs.readFile(JSON_FILE, 'utf-8');
            orders = JSON.parse(fileData || '[]');
        } catch (e) {}

        const { datePart } = getTaiwanTimeParts();

        // 過濾出今天的訂單，並包裝成前端對應欄位 (time, cardId, empName, meal, note)
        const todayOrders = orders
            .filter(order => order.timestamp && order.timestamp.startsWith(datePart))
            .map(order => ({
                time: order.timestamp.split(' ')[1].substring(0, 5), // 只取 "HH:mm"
                cardId: order.cardId,
                empName: order.name,
                meal: order.meal,
                note: `辣度: ${order.spicy} | 備註: ${order.note} | 金額: $${order.total}`
            }));

        res.json({ success: true, orders: todayOrders });
    } catch (err) {
        res.status(500).json({ success: false, message: '讀取統計資料失敗' });
    }
});

// B-1. 獲取員工卡號清單
app.get('/api/admin/employees', (req, res) => {
    // 將記憶體物件轉成前端需要的陣列格式 [{ cardId: "...", empName: "..." }]
    const employees = Object.keys(USER_DB).map(cardId => ({
        cardId: cardId,
        empName: USER_DB[cardId]
    }));
    res.json({ success: true, employees });
});

// B-2. 新增員工卡號
app.post('/api/admin/employees', (req, res) => {
    const { cardId, empName } = req.body;
    if (!cardId || !empName) return res.status(400).json({ success: false, message: '欄位缺失' });

    USER_DB[String(cardId).trim()] = String(empName).trim();
    console.log(`👤 [後台提示] 管理員新增了員工：${empName} (${cardId})`);
    res.json({ success: true });
});

// B-3. 刪除員工卡號
app.delete('/api/admin/employees/:cardId', (req, res) => {
    const { cardId } = req.params;
    if (USER_DB[cardId]) {
        console.log(`👤 [後台提示] 管理員移除了員工：${USER_DB[cardId]} (${cardId})`);
        delete USER_DB[cardId];
        res.json({ success: true });
    } else {
        res.status(404).json({ success: false, message: '找不到該員工' });
    }
});

// C-1. 獲取合作店家清單
app.get('/api/admin/shops', async (req, res) => {
    try {
        const fileData = await fs.readFile(SHOPS_FILE, 'utf-8');
        const shops = JSON.parse(fileData || '[]');
        res.json({ success: true, shops });
    } catch (e) {
        res.json({ success: true, shops: DEFAULT_SHOPS });
    }
});

// C-2. 新增合作店家
app.post('/api/admin/shops', async (req, res) => {
    const { name, file, desc } = req.body;
    try {
        const fileData = await fs.readFile(SHOPS_FILE, 'utf-8');
        const shops = JSON.parse(fileData || '[]');
        
        shops.push({ name, file, desc });
        await fs.writeFile(SHOPS_FILE, JSON.stringify(shops, null, 2), 'utf-8');
        console.log(`🏪 [後台提示] 成功上架新店家：${name}`);
        res.json({ success: true });
    } catch (e) {
        res.status(500).json({ success: false, message: '寫入店家資料失敗' });
    }
});

// C-3. 下架合作店家
app.delete('/api/admin/shops/:name', async (req, res) => {
    const shopName = decodeURIComponent(req.params.name);
    try {
        const fileData = await fs.readFile(SHOPS_FILE, 'utf-8');
        let shops = JSON.parse(fileData || '[]');
        
        shops = shops.filter(s => s.name !== shopName);
        await fs.writeFile(SHOPS_FILE, JSON.stringify(shops, null, 2), 'utf-8');
        console.log(`🏪 [後台提示] 成功下架店家：${shopName}`);
        res.json({ success: true });
    } catch (e) {
        res.status(500).json({ success: false, message: '移除店家資料失敗' });
    }
});

// ==================== 🚀 啟動伺服器 ====================
app.listen(PORT, () => {
    console.log(`🚀 訂餐系統全功能商用版已啟動！`);
    console.log(`👉 員工點餐大廳：http://localhost:${PORT}`);
    console.log(`👉 後台管理中心：http://localhost:${PORT}/admin.html (假設你的管理頁檔名為 admin.html)`);
});