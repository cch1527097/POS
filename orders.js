// ==================== 新增：獲取個人當日點餐紀錄 API ====================
app.get('/api/order-history', (req, res) => {
    const { cardId } = req.query;

    if (!cardId) {
        return res.json({ success: false, message: '缺少員工卡號' });
    }

    try {
        let ordersArray = [];
        // 1. 檢查 json 檔案是否存在，存在就讀取出來
        if (fs.existsSync(JSON_FILE)) {
            const fileData = fs.readFileSync(JSON_FILE, 'utf-8');
            ordersArray = JSON.parse(fileData || '[]');
        }

        // 2. 獲取今天的日期字串（格式例如 "2026/7/8"），用來過濾當天訂單
        const todayStr = new Date().toLocaleDateString('zh-TW', { timeZone: 'Asia/Taipei' });

        // 3. 篩選出「屬於該卡號」且「是今天點的」訂單
        const userOrders = ordersArray
            .filter(order => {
                const isSameUser = String(order.cardId).trim() === String(cardId).trim();
                const isToday = order.timestamp && order.timestamp.includes(todayStr);
                return isSameUser && isToday;
            })
            .map(order => ({
                time: order.timestamp ? order.timestamp.split(' ')[1] : '-', // 只截取時間部分（如 12:30:15），畫面更乾淨
                meal: order.meal,
                note: `醬料辣度: ${order.spicy || '無'} | 備註: ${order.note || '無'} | 金額: $${order.total || 0}`
            }));

        // 4. 回傳給前端網頁
        res.json({
            success: true,
            orders: userOrders
        });

    } catch (error) {
        console.error('❌ 讀取訂單紀錄失敗:', error);
        res.json({ success: false, message: '讀取紀錄失敗', orders: [] });
    }
});