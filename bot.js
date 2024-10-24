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

const materialsData = readExcelData('./Калькулятор для бота .xlsx'); 
let userData = { products: [], length: 0, deliveryCost: 0 };

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

function askQuantity(chatId, product) {
    bot.sendMessage(chatId, `Введите количество для ${product.name} (единица: ${product.unit}):`);
    bot.once('message', (msg) => {
        const quantity = parseInt(msg.text);
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
            askDeliveryCost(chatId);
        } else {
            bot.sendMessage(chatId, "Введите корректную длину.");
            askLength(chatId);
        }
    });
}

function askDeliveryCost(chatId) {
    bot.sendMessage(chatId, "Введите стоимость доставки:");
    bot.once('message', (msg) => {
        const deliveryCost = parseFloat(msg.text);
        if (!isNaN(deliveryCost) && deliveryCost >= 0) {
            userData.deliveryCost = deliveryCost;
            askMoreProducts(chatId);
        } else {
            bot.sendMessage(chatId, "Введите корректную стоимость доставки.");
            askDeliveryCost(chatId);
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
    const { length, products, deliveryCost } = userData;
    let totalCost = deliveryCost;

    products.forEach(product => {
        totalCost += product.price * product.quantity;
    });

    const resultMessage = `Итог:\nДлина забора: ${length} м\nВыбранные продукты:\n${products.map(product => `${product.name} - ${product.quantity} ${product.unit} (цена: ${product.price} за ${product.unit})`).join('\n')}\nОбщая стоимость: ${totalCost.toFixed(2)} руб.`;

    bot.sendMessage(chatId, resultMessage);
    
    // Generate PDF and send to user
    generatePDF(chatId, length, products, totalCost);
}

function generatePDF(chatId, length, products, totalCost) {
    const doc = new PDFDocument();
    const filePath = `./invoice_${chatId}.pdf`;

    // Create write stream for PDF
    const writeStream = fs.createWriteStream(filePath);
    doc.pipe(writeStream);

    doc.fontSize(20).text('Счет', { align: 'center' });
    doc.moveDown();
    doc.fontSize(12).text(`Длина забора: ${length} м`);
    doc.moveDown();

    doc.text('Выбранные продукты:');
    products.forEach(product => {
        doc.text(`${product.name} - ${product.quantity} ${product.unit} (цена: ${product.price} за ${product.unit})`);
    });
    doc.moveDown();

    doc.text(`Общая стоимость: ${totalCost.toFixed(2)} руб.`);
    doc.end();

    // Send PDF to user after the PDF document has finished writing
    writeStream.on('finish', () => {
        bot.sendDocument(chatId, filePath)
            .then(() => {
                console.log(`PDF sent to chat ID ${chatId}`);
                // Optionally, delete the PDF file after sending
                fs.unlink(filePath, (err) => {
                    if (err) {
                        console.error("Failed to delete PDF:", err);
                    }
                });
            })
            .catch(err => {
                console.error("Failed to send PDF:", err);
            });
    });

    // Log any errors with the write stream
    writeStream.on('error', (err) => {
        console.error("Failed to write PDF:", err);
    });
}


process.on('uncaughtException', (err) => {
    console.error("An error occurred:", err);
});
