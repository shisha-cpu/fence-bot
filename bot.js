const TelegramBot = require('node-telegram-bot-api');
const XLSX = require('xlsx');
const PDFDocument = require('pdfkit');
const fs = require('fs');
const https = require('https');
const { log } = require('console');
const path = require('path');
const bot = new TelegramBot('7510774911:AAFRh2cPNTXfGVXtkE_5cGCCgi7cGXo8Wbs', { polling: true }); 
let accessList = new Set(); 
const adminId = 219764990;
const ACCESS_FILE_PATH = './accessList.json';
//219764990
// Загрузка списка доступа при запуске
loadAccessList();

// Функции для работы с JSON файлом
function saveAccessList() {
    const data = JSON.stringify(Array.from(accessList));
    fs.writeFileSync(ACCESS_FILE_PATH, data);
}

function loadAccessList() {
    if (fs.existsSync(ACCESS_FILE_PATH)) {
        const data = fs.readFileSync(ACCESS_FILE_PATH, 'utf8');
        accessList = new Set(JSON.parse(data));
    } else {
        accessList = new Set();
    }
}

bot.on("polling_error", (error) => {
    console.error("Ошибка в процессе polling:", error.code, error.response?.body || error);
});

function readExcelData(filePath) {
    try {
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
    } catch (error) {
        console.error("Ошибка при чтении данных из Excel:", error);
        return {};
    }
}

bot.onText(/\/admin/, (msg) => {
    const chatId = msg.chat.id;
    if (chatId === adminId) {       
        bot.sendMessage(chatId, "Отправьте новый Excel-файл для обновления данных.");
     
        bot.once('document', async (fileMsg) => {
            const fileId = fileMsg.document.file_id;
            const filePath = './updated_file.xlsx';

            try {
                const fileLink = await bot.getFileLink(fileId);
                const fileStream = fs.createWriteStream(filePath, { flags: 'wx' });
                
                fileStream.on('finish', () => {
                    bot.sendMessage(chatId, "Файл успешно загружен и обновлен.");
                    
                    const newMaterialsData = readExcelData(filePath);
                    if (newMaterialsData) {
                        materialsData = newMaterialsData;
                    }
                    
                    fs.unlinkSync(filePath);
                });

                https.get(fileLink, (response) => response.pipe(fileStream));

            } catch (error) {
                console.error("Ошибка при загрузке файла:", error);
                bot.sendMessage(chatId, "Произошла ошибка при загрузке файла.");
            }
        });
    } else {
        bot.sendMessage(chatId, "У вас нет прав для выполнения этой команды.");
    }
});

bot.onText(/\/access (.+)/, (msg, match) => {
    const chatId = msg.chat.id;
    if (chatId !== adminId) {
        return bot.sendMessage(chatId, "У вас нет прав для выполнения этой команды.");
    }

    const args = match[1].split(' ');
    const [username, action] = args;

    if (action === 'add') {
        accessList.add(username);
        saveAccessList();
        bot.sendMessage(chatId, `Пользователь ${username} добавлен в список доступа.`);
    } else if (action === 'remove') {
        accessList.delete(username);
        saveAccessList();
        bot.sendMessage(chatId, `Пользователь ${username} удалён из списка доступа.`);
    } else {
        bot.sendMessage(chatId, "Неверная команда. Используйте '/access <ник> add' или '/access <ник> remove'.");
    }
});

function checkAccess(msg, username) {
    if (!msg.from) {
        return false;
    }

    // Проверяем, является ли пользователь администратором
    if (msg.from.id === adminId) {
        return true;
    }

    // Проверяем, есть ли пользователь в списке доступа
    return accessList.has(username);
}
let materialsData = readExcelData('./Калькулятор для бота  7 (1).xlsx');
let userData = { products: [] };
bot.onText(/\/start/, (msg) => {
    const chatId = msg.chat.id;
    const username = msg.from.username;

    // Проверяем, является ли пользователь администратором
    if (msg.from.id === adminId) {
        // Администратору всегда доступен бот
        userData[chatId] = { products: [] };
        bot.sendMessage(chatId, "Добро пожаловать, администратор! Для выбора материала нажмите кнопку ниже.");
        askMaterial(chatId);
        return;
    }

    // Для остальных пользователей проверяем доступ
    if (!checkAccess(username)) {
        return bot.sendMessage(chatId, "У вас нет доступа к этому боту. Обратитесь к администратору.");
    }

    if (userData[chatId]) { 
        return bot.sendMessage(chatId, "Вы уже начали работу с ботом. Используйте меню для продолжения.");
    }

    userData[chatId] = { products: [] }; 
    bot.sendMessage(chatId, "Добро пожаловать! Для выбора материала нажмите кнопку ниже.");
    askMaterial(chatId);
});

function checkAccess(username) {
    return accessList.has(username);
}


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
        if (msg.text === '/admin') return;

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
    if (product.name === 'Доставка материала' || product.name === 'Доставка откатной' || product.name === 'Наценка') {
        bot.sendMessage(chatId, `Введите сумму для ${product.name} (руб.):`);
        bot.once('message', (msg) => {
            const sum = parseFloat(msg.text);
            if (!isNaN(sum) && sum >= 0) {
                userData[chatId].products.push({ ...product, sum });
                askMoreProducts(chatId);
            } else {
                bot.sendMessage(chatId, "Введите корректную сумму.");
                askQuantity(chatId, product);
            }
        });
    } else {
        bot.sendMessage(chatId, `Введите количество для ${product.name} (единица: ${product.unit}):`);
        bot.once('message', (msg) => {
            const quantity = parseInt(msg.text);
            if (!isNaN(quantity) && quantity > 0) {
                userData[chatId].products.push({ ...product, quantity });
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
    const { products } = userData[chatId];
    let totalCost = 0;

    products.forEach(product => {
        console.log(`Обрабатываем продукт: ${product.name}`);
        console.log(`Цена: ${product.price}, Количество: ${product.quantity}, Сумма: ${product.sum || Math.round(product.quantity * product.price)}`);

   if (!product.sum && !isNaN(product.price) && !isNaN(product.quantity)) {
            product.sum = Math.round(product.price * product.quantity);  
        }


        if (!isNaN(product.sum)) {
            totalCost += product.sum;
        }
    });

    bot.sendMessage(chatId, "Введите сумму скидки (если есть, иначе введите 0):");
    bot.once('message', (msg) => {
        const discount = Math.round(parseFloat(msg.text) || 0);
        totalCost -= discount;

        const resultMessage = `Итог:\nВыбранные продукты:\n${products.map(product => {
            const quantity = product.quantity || '';
            const pricePerUnit = Math.round(product.price);
            const totalPrice = Math.round(product.sum || product.quantity * product.price);
            return `${product.name} - ${quantity} ${product.unit} (цена: ${isNaN(pricePerUnit) ? 'н/д' : pricePerUnit} за ${product.unit}) - ${totalPrice}`;
        }).join('\n')}` +
        `\nОбщая стоимость: ${Math.round(totalCost)} руб.` +
        (discount > 0 ? `\nСкидка: ${discount} руб.` : '');

        bot.sendMessage(chatId, resultMessage);

        generatePDF(chatId, products, totalCost , discount);
        delete userData[chatId];
    });
}


function generatePDF(chatId, products, totalCost , discount) {
    const doc = new PDFDocument();
    const filePath = `./invoice_${chatId}.pdf`;

    const writeStream = fs.createWriteStream(filePath);
    doc.pipe(writeStream);

    doc.font('./ArialRegular.ttf');
    doc.fontSize(20).text('Смета', { align: 'center' });
    doc.moveDown();

    doc.fontSize(12);
    const tableTop = doc.y;
    const characteristicsX = 50;
    const quantityX = 300;
    const unitX = 350;
    const priceX = 400;
    const sumX = 450;
    const maxRowHeight = 50;


    doc.text('Характеристики', characteristicsX, tableTop, { width: 250, align: 'left' });
    doc.text('Кол-во', quantityX, tableTop, { width: 40, align: 'center' });
    doc.text('Ед.', unitX, tableTop, { width: 40, align: 'center' });
    doc.text('Цена', priceX, tableTop, { width: 40, align: 'center' });
    doc.text('Сумма', sumX, tableTop, { width: 50, align: 'center' });

    doc.moveTo(characteristicsX, tableTop + 15)
       .lineTo(sumX + 50, tableTop + 15)
       .stroke();

    let positionY = tableTop + 25;


    products.forEach(product => {
        const characteristics = `${product.name}${product.description ? ': ' + product.description : ''}`;
        const textOptions = { width: 250, align: 'left', lineGap: 2 };

        let lineHeight = doc.heightOfString(characteristics, textOptions);
        const rowHeight = Math.max(lineHeight, maxRowHeight);


        if (positionY + rowHeight + 30 > doc.page.height) {  
            doc.addPage();
            positionY = 50; 
            doc.fontSize(12); 

            doc.text('Характеристики', characteristicsX, positionY, { width: 250, align: 'left' });
            doc.text('Кол-во', quantityX, positionY, { width: 40, align: 'center' });
            doc.text('Ед.', unitX, positionY, { width: 40, align: 'center' });
            doc.text('Цена', priceX, positionY, { width: 40, align: 'center' });
            doc.text('Сумма', sumX, positionY, { width: 50, align: 'center' });

            doc.moveTo(characteristicsX, positionY + 15)
               .lineTo(sumX + 50, positionY + 15)
               .stroke();
            positionY += 20; 
        }


        doc.rect(characteristicsX, positionY - 5, sumX + 50 - characteristicsX, rowHeight + 10).stroke();


        const displayPrice = isNaN(product.price) ? '-' : Math.round(product.price);
        const displaySum = isNaN(product.sum) ? '-' : Math.round(product.sum || product.quantity * product.price);

        doc.text(characteristics, characteristicsX, positionY, textOptions);
        doc.text(product.quantity || '', quantityX, positionY, { width: 40, align: 'center' });
        doc.text(product.unit, unitX, positionY, { width: 40, align: 'center' });
        doc.text(displayPrice, priceX, positionY, { width: 40, align: 'center' });
        doc.text(displaySum, sumX, positionY, { width: 50, align: 'center' });

        positionY += rowHeight + 10;
    });

    positionY += 10;
    doc.moveTo(characteristicsX, positionY)
       .lineTo(sumX + 50, positionY)
       .stroke();

    positionY += 10;
    doc.fontSize(12).text(`Общая стоимость: ${isNaN(totalCost) ? '-' : Math.round(totalCost)} руб.`, sumX - 50, positionY, { align: 'right' });
    if (discount > 0) {
        doc.text(`Скидка: ${discount} руб.`, { align: 'right' });
    }
    doc.end();

    writeStream.on('finish', () => {
        bot.sendDocument(chatId, filePath).then(() => {
            fs.unlinkSync(filePath);
        });
    });
}





