const TelegramBot = require('node-telegram-bot-api');
const XLSX = require('xlsx');


const bot = new TelegramBot('7510774911:AAFRh2cPNTXfGVXtkE_5cGCCgi7cGXo8Wbs', { polling: true });


function readExcelData(filePath) {
    const workbook = XLSX.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

    const result = {};
    let currentCategory = '';

    for (let i = 0; i < data.length; i++) {
        const row = data[i];
        if (row[0] && !row[1]) {
            currentCategory = row[0];
            result[currentCategory] = [];
        } else if (currentCategory && row[0]) {
            const product = {
                name: row[0],
                description: row[1],
                unit: row[2],
                quantity: row[3] || 0,
                price: row[4],
                sum: row[5] || 0,
            };
            result[currentCategory].push(product);
        }
    }

    return result;
}

const materialsData = readExcelData('./Калькулятор для бота .xlsx'); 
let userData = { products: [], length: 0 };


bot.onText(/\/start/, (msg) => {
    const chatId = msg.chat.id;
    bot.sendMessage(chatId, "Добро пожаловать! Для выбора материала нажмите кнопку ниже.");
    askMaterial(chatId);
});


function askMaterial(chatId) {
    const materials = [
        'Каркас забора', 'Жалюзи', 'Профнастил', 'Профнастил горизонт', 
        'Штакетник в 1 ряд', 'Штакетник в 1 ряд горизонт', 
        'Штакетник шахматка, горизонтальный', 'Штакетник дерево', 
        'Рабица', '3Д', 'Калитки, ворота', 'Допы', 
        'Монолит, сваи, кирпич и т.д.', 'Навесы', 'Доставка, наценка'
    ];

    const buttons = materials.map(material => [{ text: material }]);

    bot.sendMessage(chatId, "Выберите материал для забора:", {
        reply_markup: {
            keyboard: buttons,
            one_time_keyboard: true
        }
    });

    bot.once('message', (msg) => {
        const selectedMaterial = msg.text;
        if (materials.includes(selectedMaterial)) {
            askProduct(chatId, selectedMaterial);
        } else {
            bot.sendMessage(chatId, "Неизвестный материал. Попробуйте снова.");
            askMaterial(chatId);
        }
    });
}


function askProduct(chatId, material) {
    const products = materialsData[material];
    const productList = products.map((product, index) => `${index + 1}. ${product.name} (цена: ${product.price} ${product.unit})`).join('\n');
    bot.sendMessage(chatId, `Выберите продукт:\n${productList}\nВведите номер продукта для выбора.`);
    
    bot.once('message', (msg) => {
        const productIndex = parseInt(msg.text) - 1;
        if (products[productIndex]) {
            askQuantity(chatId, products[productIndex]);
        } else {
            bot.sendMessage(chatId, "Неверный выбор. Попробуйте еще раз.");
            askProduct(chatId, material);
        }
    });
}

function askQuantity(chatId, product) {
    bot.sendMessage(chatId, `Введите количество для ${product.name} (единица: ${product.unit}):`);
    bot.once('message', (msg) => {
        const quantity = parseFloat(msg.text);
        if (!isNaN(quantity) && quantity > 0) {
            userData.products.push({ ...product, quantity });
            askLength(chatId);
        } else {
            bot.sendMessage(chatId, "Введите корректное количество.");
            askQuantity(chatId, product);
        }
    });
}


function askLength(chatId) {
    bot.sendMessage(chatId, "Укажите длину забора (в метрах):");
    bot.once('message', (msg) => {
        const length = parseFloat(msg.text);
        if (!isNaN(length) && length > 0) {
            userData.length = length;
            askMoreProducts(chatId);
        } else {
            bot.sendMessage(chatId, "Введите корректную длину.");
            askLength(chatId);
        }
    });
}


function askMoreProducts(chatId) {
    bot.sendMessage(chatId, "Хотите выбрать еще один продукт? (да/нет)");
    bot.once('message', (msg) => {
        if (msg.text.toLowerCase() === 'да') {
            askMaterial(chatId);
        } else {
            calculateCost(chatId);
        }
    });
}


function calculateCost(chatId) {
    const { length, products } = userData;
    let totalCost = 0;

    products.forEach(product => {
        totalCost += product.price * product.quantity;
    });

    const resultMessage = `Итог:\nДлина забора: ${length} м\nВыбранные продукты:\n${products.map(product => `${product.name} - ${product.quantity} ${product.unit} (цена: ${product.price} за ${product.unit})`).join('\n')}\nОбщая стоимость: ${totalCost.toFixed(2)} руб.`;
    
    bot.sendMessage(chatId, resultMessage);
}


process.on('uncaughtException', (err) => {
    console.error("An error occurred:", err);
});
