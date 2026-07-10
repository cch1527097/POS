const express = require('express');
const cors = require('cors');
const app = express();
const PORT = 3000;

// ==================== 系統核心中間件 ====================
app.use(cors());
app.use(express.json());

// ==================== 模擬資料庫 (記憶體儲存) ====================
// 管理員帳密
const ADMIN_ACCOUNT = {
    username: "18445",
    password: "zxc456456"
};

// 預設員工資料
let employees = [
    { cardId: "18445", empName: "測試管理員" },
    { cardId: "E012", empName: "王小明" },
    { cardId: "E023", empName: "張大華" }
];

// 預設店家資料
let shops = [
    { name: "八方雲集", file: "bafang.html", desc: "皮Q餡飽滿的招牌鍋貼與各式經典水餃。" },
    { name: "榮郁香廣式燒臘", file: "dafang.html", desc: "道地廣式燒臘，外酥內嫩的美味首選。" }
];

// 預設今日點餐訂單
let orders = [
    { time: "11:24", cardId: "E012", empName: "王小明", meal: "八方雲集", note: "鍋貼 10 顆，不辣" },
    { time: "11:45", cardId: "E023", empName: "張大華", meal: "榮郁香廣式燒臘", note: "三寶飯，油蔥多" }
];

// ==================== ⭐ 【新增】點餐大廳專用 API ====================

// 1. 員工卡號認證
app.post('/api/login', (req, res) => {
    const { cardId } = req.body;
    if (!cardId) {
        return res.json({ success: false, message: "卡號不能為空" });
    }

    // 在員工清單中搜尋該卡號
    const employee = employees.find(emp => emp.cardId === cardId.trim());
    
    if (employee) {
        res.json({ success: true, empName: employee.empName });
    } else {
        res.json({ success: false, message: "查無此員工卡號，請聯絡管理員新增" });
    }
});

// 2. 接收大廳點餐/進入紀錄
app.post('/api/order', (req, res) => {
    const { cardId, meal, note } = req.body;
    
    if (!cardId || !meal) {
        return res.json({ success: false, message: "訂單資料不完整" });
    }

    // 撈取員工姓名，若找不到則用卡號代替
    const employee = employees.find(emp => emp.cardId === cardId);
    const empName = employee ? employee.empName : "未知員工";

    // 產生當前時間 (HH:MM)
    const now = new Date();
    const timeStr = `${String(now.getHours()).padStart(2, '0')}:${String(now.getMinutes()).padStart(2, '0')}`;

    // 建立新訂單並推入陣列
    const newOrder = {
        time: timeStr,
        cardId: cardId,
        empName: empName,
        meal: meal,
        note: note || ""
    };

    orders.push(newOrder);
    res.json({ success: true, message: "訂單送出成功" });
});

// 3. 查詢特定員工的歷史點餐紀錄
app.get('/api/order-history', (req, res) => {
    const { cardId } = req.query;
    if (!cardId) {
        return res.json({ success: false, message: "缺少卡號參數" });
    }

    // 篩選出屬於該員工的訂單
    const userOrders = orders.filter(order => order.cardId === cardId);
    res.json({ success: true, orders: userOrders });
});


// ==================== 🛠️ 管理員權限驗證 API ====================
app.post('/api/admin/login', (req, res) => {
    const { username, password } = req.body;
    
    if (username === ADMIN_ACCOUNT.username && password === ADMIN_ACCOUNT.password) {
        res.json({ success: true, message: "登入成功" });
    } else {
        res.json({ success: false, message: "帳號或密碼錯誤" });
    }
});

// ==================== 📊 訂單統計 API ====================
app.get('/api/admin/orders', (req, res) => {
    res.json({ success: true, orders: orders });
});

// ==================== 👥 員工卡號管理 API ====================
app.get('/api/admin/employees', (req, res) => {
    res.json({ success: true, employees: employees });
});

app.post('/api/admin/employees', (req, res) => {
    const { cardId, empName } = req.body;
    
    if (!cardId || !empName) {
        return res.json({ success: false, message: "欄位不能為空" });
    }
    
    const isDuplicate = employees.some(emp => emp.cardId === cardId);
    if (isDuplicate) {
        return res.json({ success: false, message: "該員工卡號已存在！" });
    }
    
    employees.push({ cardId, empName });
    res.json({ success: true });
});

app.delete('/api/admin/employees/:cardId', (req, res) => {
    const { cardId } = req.params;
    const initialLength = employees.length;
    
    employees = employees.filter(emp => emp.cardId !== cardId);
    
    if (employees.length < initialLength) {
        res.json({ success: true });
    } else {
        res.json({ success: false, message: "找不到該員工卡號" });
    }
});

// ==================== 🏪 合作店家管理 API ====================
app.get('/api/admin/shops', (req, res) => {
    res.json({ success: true, shops: shops });
});

app.post('/api/admin/shops', (req, res) => {
    const { name, file, desc } = req.body;
    
    if (!name || !file || !desc) {
        return res.json({ success: false, message: "店家資訊不完整" });
    }
    
    shops.push({ name, file, desc });
    res.json({ success: true });
});

app.delete('/api/admin/shops/:shopName', (req, res) => {
    const { shopName } = req.params;
    const decodedName = decodeURIComponent(shopName);
    const initialLength = shops.length;
    
    shops = shops.filter(shop => shop.name !== decodedName);
    
    if (shops.length < initialLength) {
        res.json({ success: true });
    } else {
        res.json({ success: false, message: "找不到該店家" });
    }
});

// ==================== 啟動伺服器 ====================
app.listen(PORT, () => {
    console.log(`=============================================`);
    console.log(` 美食訂餐系統後端伺服器已成功啟動！`);
    console.log(` 運行網址: http://localhost:${PORT}`);
    console.log(` 測試管理員帳號: ${ADMIN_ACCOUNT.username}`);
    console.log(` 測試管理員密碼: ${ADMIN_ACCOUNT.password}`);
    console.log(`=============================================`);
});