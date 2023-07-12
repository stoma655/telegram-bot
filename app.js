
const { Telegraf } = require('telegraf');
const XLSX = require('xlsx');
import fetch from 'node-fetch';
const LocalSession = require('telegraf-session-local');
const fs = require('fs');

const bot = new Telegraf('6072826493:AAE22jKjF6NJMxMDt1YZQ8eIH5AEGMMR8Xw');

// Настройка локального хранилища сеансов
const localSession = new LocalSession({ database: 'sessions.json' });
bot.use(localSession.middleware());

function generateRandomFilename() {
  let randomString = '';
  const characters = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
  for (let i = 0; i < 10; i++) {
    randomString += characters.charAt(Math.floor(Math.random() * characters.length));
  }
  return `result-${randomString}.xlsx`;
}


bot.command('clear', (ctx) => {
  // Clear all variables
  json = null;
  wb = null;
  waitingForExcelFile = false;

  // Clear chat history
  ctx.session = null;

  // Send confirmation message to user
  ctx.reply('Все переменные очищены.');
});


let json = null;
let wb = null;



bot.on('document', async (ctx) => {
  try {
  const fileId = ctx.message.document.file_id;
  const fileUrl = await ctx.telegram.getFileLink(fileId);
  const response = await fetch(fileUrl);

  if (ctx.message.document.mime_type === 'application/json') {
    // Handle JSON file
    json = await response.json();

    setTimeout(() => {
      if (wb) {
        // Add data from JSON to existing Excel file
        addDataToExcel(json, wb);
        sendExcelFile(ctx, wb);
        json = null;
        wb = null;
      } else {
        // Create new Excel file
        wb = XLSX.utils.book_new();
        addDataToExcel(json, wb);
        sendExcelFile(ctx, wb);
        json = null;
        wb = null;
      }
    }, 2500);
  } else if (ctx.message.document.mime_type === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet') {
    // Handle Excel file
    const arrayBuffer = await response.arrayBuffer();
    const buffer = Buffer.from(arrayBuffer);
    wb = XLSX.read(buffer);

    // if (json) {
    //   // Add data from JSON to existing Excel file
    //   addDataToExcel(json, wb);
    //   sendExcelFile(ctx, wb);
    //   json = null;
    //   wb = null;
    // }
  }
} catch (error) {
  // Обработка ошибки и отправка сообщения об ошибке в Telegram
  ctx.reply(`Произошла ошибка: ${error.message}`);
}
});

function addDataToExcel(json, wb) {
  // Extract data from JSON
  const teamNames = Object.keys(json.votes);
  const data = [];
  for (const teamName of teamNames) {
    for (const vote of json.votes[teamName]) {
      const row = [vote.name];
      for (const tn of teamNames) {
        row.push(tn === teamName ? 'W' : '');
      }
      data.push(row);
    }
  }

  // Add data to Excel
  let ws;
  if (wb.SheetNames.length > 0) {
    ws = wb.Sheets[wb.SheetNames[0]];
    const range = XLSX.utils.decode_range(ws['!ref']);
    const nextCol = range.e.c + 2;

    // Extract names from first table
    const names = [];
    for (let row = 2; row <= range.e.r; row++) {
      const cell = ws[XLSX.utils.encode_cell({ c: 0, r: row })];
      if (cell) {
        names.push(cell.v);
      }
    }

    // Add data to Excel
    XLSX.utils.sheet_add_aoa(ws, [[json.description]], { origin: { r: 0, c: nextCol } });
    XLSX.utils.sheet_add_aoa(ws, [teamNames], { origin: { r: 1, c: nextCol } });
    for (let row = 0; row < names.length; row++) {
      const name = names[row];
      const rowData = data.find((d) => d[0] === name);
      if (rowData) {
        XLSX.utils.sheet_add_aoa(ws, [rowData.slice(1)], { origin: { r: row + 2, c: nextCol } });
      }
    }
  } else {
    ws = XLSX.utils.aoa_to_sheet([[json.description], ['Участники', ...teamNames], ...data]);
    ws['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 2 } }];
    XLSX.utils.cell_set_number_format(ws.A1, '@');
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
  }

  // Set column width for all columns
  const cols = [];
  for (let col = 0; col < ws['!ref'].split(':')[1].charCodeAt(0) - 'A'.charCodeAt(0) + 1; col++) {
    cols.push({ wch: 20 });
  }
  ws['!cols'] = cols;
}

function sendExcelFile(ctx, wb) {
  // Generate random filename
  const filename = generateRandomFilename();

  // Send Excel file to user
  const wbout = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
  ctx.replyWithDocument({ source: Buffer.from(wbout), filename });
}

bot.launch();









