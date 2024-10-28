const TelegramBot = require('node-telegram-bot-api');
const XLSX = require('xlsx');
const PDFDocument = require('pdfkit');
const fs = require('fs');

const bot = new TelegramBot('7510774911:AAFRh2cPNTXfGVXtkE_5cGCCgi7cGXo8Wbs', { polling: true }); // Replace with your token
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


const materialsData = readExcelData('./Калькулятор для бота  7 (1).xlsx');
let userData = { products: [] };

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
            one_time_keyboard: true,
            resize_keyboard: true
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

    const productList = products.map((product, index) => {
        const isBold = /1\.8м|2м/.test(product.name);
        const formattedName = isBold ? `*${product.name}*` : product.name;
        return `${index + 1}. ${formattedName} (цена: ${product.price} ${product.unit})`;
    }).join('\n');

    const buttons = products.map((product, index) => [{ text: `${index + 1}. ${product.name}` }]);

    bot.sendMessage(chatId, `Выберите продукт:\n${productList}`, {
        reply_markup: {
            keyboard: buttons,
            one_time_keyboard: true,
            resize_keyboard: true
        },
        parse_mode: 'Markdown'
    });

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


function askCustomPrice(chatId, product) {
    bot.sendMessage(chatId, `Введите свою цену для ${product.name}:`);
    bot.once('message', (msg) => {
        const price = parseFloat(msg.text);
        if (!isNaN(price) && price >= 0) {
            userData.products.push({ ...product, price });
            askMoreProducts(chatId);
        } else {
            bot.sendMessage(chatId, "Введите корректную цену.");
            askCustomPrice(chatId, product);
        }
    });
}

function askDiscount(chatId) {
    bot.sendMessage(chatId, "Введите сумму скидки:");
    bot.once('message', (msg) => {
        const discount = parseFloat(msg.text);
        if (!isNaN(discount) && discount >= 0) {
            userData.discount = discount;
            askMoreProducts(chatId);
        } else {
            bot.sendMessage(chatId, "Введите корректную сумму скидки.");
            askDiscount(chatId);
        }
    });
}
function askQuantity(chatId, product) {
    if (product.name === 'Доставка материала' || product.name === 'Доставка откатной' || product.name === 'Наценка') {
        // If product is one of the first three, allow custom sum input
        bot.sendMessage(chatId, `Введите сумму для ${product.name} (руб.):`);
        bot.once('message', (msg) => {
            const sum = parseFloat(msg.text);
            if (!isNaN(sum) && sum >= 0) {
                userData.products.push({ ...product, sum });
                askMoreProducts(chatId);
            } else {
                bot.sendMessage(chatId, "Введите корректную сумму.");
                askQuantity(chatId, product);
            }
        });
    } else {
        // For other products, ask for quantity
        bot.sendMessage(chatId, `Введите количество для ${product.name} (единица: ${product.unit}):`);
        bot.once('message', (msg) => {
            const quantity = parseInt(msg.text);
            if (!isNaN(quantity) && quantity > 0) {
                userData.products.push({ ...product, quantity });
                askMoreProducts(chatId);
            } else {
                bot.sendMessage(chatId, "Введите корректное количество.");
                askQuantity(chatId, product);
            }
        });
    }
}
function askMoreProducts(chatId) {
    bot.sendMessage(chatId, "Хотите выбрать еще один продукт?", {
        reply_markup: {
            keyboard: [
                [{ text: 'Да' }, { text: 'Нет' }]
            ],
            one_time_keyboard: true,
            resize_keyboard: true
        }
    });

    bot.once('message', (msg) => {
        const answer = msg.text.toLowerCase();
        if (answer === 'да') {
            askMaterial(chatId);
        } else if (answer === 'нет') {
            calculateCost(chatId);
        } else {
            bot.sendMessage(chatId, "Пожалуйста, выберите 'да' или 'нет'.");
            askMoreProducts(chatId);
        }
    });
}

function calculateCost(chatId) {
    const { products } = userData;
    let totalCost = 0;

    products.forEach(product => {
        if (product.sum) {
            // If sum is defined, use that for total cost
            totalCost += product.sum;
        } else {
            totalCost += product.price * product.quantity;
        }
    });

    bot.sendMessage(chatId, "Введите сумму скидки (если есть, иначе введите 0):");
    bot.once('message', (msg) => {
        const discount = parseFloat(msg.text) || 0;
        totalCost -= discount; // Subtract discount from total cost

        const resultMessage = `Итог:\nВыбранные продукты:\n${products.map(product => 
            `${product.name} - ${product.sum ? product.sum : product.quantity} ${product.unit} (цена: ${product.price} за ${product.unit})`).join('\n')}\nОбщая стоимость: ${totalCost.toFixed(2)} руб.`;

        bot.sendMessage(chatId, resultMessage);

        // Generate PDF and send to user
        generatePDF(chatId, products, totalCost);
        userData.products = [];
    });
}

function generatePDF(chatId, products, totalCost) {
    const doc = new PDFDocument();
    const filePath = `./invoice_${chatId}.pdf`;

    const writeStream = fs.createWriteStream(filePath);
    doc.pipe(writeStream);

    doc.font('./ArialRegular.ttf');
    doc.fontSize(20).text('Смета', { align: 'center' });
    doc.moveDown();

    doc.fontSize(12);
    const tableTop = doc.y;
    const itemX = 50;
    const quantityX = 250;
    const unitX = 300;
    const priceX = 350;
    const sumX = 400;

    // Заголовок таблицы
    doc.text('Продукт', itemX, tableTop, { width: 180 });
    doc.text('Кол-во', quantityX, tableTop, { width: 40 });
    doc.text('Ед.', unitX, tableTop, { width: 40 });
    doc.text('Цена', priceX, tableTop, { width: 40 });
    doc.text('Сумма', sumX, tableTop, { width: 50 });

    doc.moveTo(itemX, tableTop + 15)
       .lineTo(sumX + 50, tableTop + 15)
       .stroke();

    let positionY = tableTop + 25;

    products.forEach(product => {
        const fullName = `${product.name}${product.description ? ': ' + product.description : ''}`;
        
        // Вычисление высоты строки на основе текста
        const textOptions = { width: 180, height: 50, align: 'left' };
        const textHeight = doc.heightOfString(fullName, textOptions);
        
        doc.text(fullName, itemX, positionY, textOptions);
        doc.text(product.quantity || '', quantityX, positionY, textOptions);
        doc.text(product.unit, unitX, positionY, textOptions);
        doc.text(product.price, priceX, positionY, textOptions);
        doc.text(product.sum || (product.quantity * product.price).toFixed(2), sumX, positionY, textOptions);

        positionY += Math.max(textHeight, 20); // Set a minimum height for rows
    });

    doc.moveDown();
    doc.text(`Общая стоимость: ${totalCost.toFixed(2)} руб.`, { align: 'right' });
    doc.end();

    writeStream.on('finish', () => {
        bot.sendDocument(chatId, filePath).then(() => {
            fs.unlinkSync(filePath); // Clean up the generated PDF file
        });
    });
}