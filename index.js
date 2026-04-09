const { Client, LocalAuth, MessageMedia } = require('whatsapp-web.js');
const qrcode = require('qrcode-terminal');
const { parseSoloLog, createReplyMessage, createFinalDayMessage, parseTimeToMinutes } = require('./splitter');

// Armazenamento em memória das hunts de cada chat (por grupo/chat-id)
const fs = require('fs');
const path = require('path');

const SESSIONS_FILE = path.join(__dirname, 'sessions.json');

function loadSessions() {
    if (fs.existsSync(SESSIONS_FILE)) {
        try {
            return JSON.parse(fs.readFileSync(SESSIONS_FILE, 'utf8'));
        } catch (e) {
            console.error('Erro ao ler sessions.json:', e);
            return {};
        }
    }
    return {};
}

function saveSessions(sessions) {
    try {
        fs.writeFileSync(SESSIONS_FILE, JSON.stringify(sessions, null, 2));
    } catch (e) {
        console.error('Erro ao salvar sessions.json:', e);
    }
}

let chatSessions = loadSessions();

// Inicializa o Client do WhatsApp
const client = new Client({
    authStrategy: new LocalAuth(), // Salva a sessão localmente
    puppeteer: {
        args: ['--no-sandbox', '--disable-setuid-sandbox'],
    }
});

// Gera o QR Code quando iniciado
client.on('qr', (qr) => {
    console.log('📌 LEIA O CÓDIGO QR ABAIXO PARA CONECTAR AO WHATSAPP:');
    qrcode.generate(qr, { small: true });
});

// Client pronto
client.on('ready', () => {
    console.log('✅ Tibia Analyzer Solo Bot está rodando e escutando mensagens!');
});

// Listener de mensagens recebidas
client.on('message_create', async (msg) => {
    
    // Ignora mensagens sem texto
    if (!msg.body) return;

    // Define corretamente o chat da conversa (se foi você que enviou, é 'to', se foi recebida, é 'from')
    const chatId = msg.fromMe ? msg.to : msg.from;

    // Comando do dono para encerrar o dia e zerar
    if (msg.body.toLowerCase().includes('por hoje é só') || msg.body.toLowerCase().includes('por hoje e so')) {
        console.log(`🧹 Comando de encerramento detectado em ${chatId}`);
        if (chatSessions[chatId] && (chatSessions[chatId].totalBalance !== 0 || chatSessions[chatId].totalMinutes !== 0)) {
            
            // --- SALVAR NA PLANILHA (XLSX) ---
            const ExcelJS = require('exceljs');
            const fs = require('fs');
            const path = require('path');
            const xlsxPath = path.join(__dirname, 'relatorio_mensal.xlsx');
            
            const workbook = new ExcelJS.Workbook();
            let sheet;
            
            // Lemos todo o arquivo atual se existir
            if (fs.existsSync(xlsxPath)) {
                await workbook.xlsx.readFile(xlsxPath);
                sheet = workbook.getWorksheet(1); // pega a primeira aba
            } else {
                sheet = workbook.addWorksheet('Relatório');
                sheet.addRow(['DIA', 'HORAS', 'FARM', 'TC/DIA']);
                sheet.addRow([]);
                sheet.addRow(['TOTAL', 'HORAS', '']);
                sheet.addRow(['', 'KK', '']);
                sheet.addRow(['', 'TC', '']);
            }

            // Localizando onde estão as colunas que o usuário desenhou no Excel dele
            let headerRow = -1;
            let colDia = 1, colHoras = 2, colFarm = 3; // Padrões caso não ache
            
            for (let i = 1; i <= Math.max(sheet.rowCount, 50); i++) {
                const r = sheet.getRow(i);
                for (let c = 1; c <= 30; c++) {
                    const val = r.getCell(c).value ? r.getCell(c).value.toString().toUpperCase().trim() : '';
                    if (val === 'DIA') { colDia = c; headerRow = i; }
                    if (val === 'HORAS') { colHoras = c; }
                    if (val === 'FARM') { colFarm = c; }
                }
                if (headerRow !== -1) break; // achou os cabeçalhos
            }

            if (headerRow === -1) headerRow = 1; // Se nem achar, usa o fallback

            // 1. Procuramos a primeira linha vazia exatamente na coluna de "DIA"
            let targetRow = headerRow + 1;
            for (let i = headerRow + 1; i <= Math.max(sheet.rowCount, 150); i++) {
                const r = sheet.getRow(i);
                const diaVal = r.getCell(colDia).value ? r.getCell(colDia).value.toString().trim().toUpperCase() : '';
                
                // Se a célula abaixo do DIA estiver vazia (ou for uma anotação de 'TOTAL'), paramos nela.
                if (diaVal === '' || diaVal.includes('TOTAL')) {
                    targetRow = i;
                    break;
                }
            }

            // Extrai a formatação do dia (da log lida, ou do dia atual caso não exista)
            const diaHoje = chatSessions[chatId].lastDay || new Date().getDate().toString();
            const hrs = Math.floor(chatSessions[chatId].totalMinutes / 60);
            const mins = chatSessions[chatId].totalMinutes % 60;
            const horasJogadas = `${hrs.toString().padStart(2, '0')}:${mins.toString().padStart(2, '0')}h`;
            
            // Farm do dia (Formatação pt-BR)
            const profitGoldStr = (chatSessions[chatId].totalBalance / 1000000).toFixed(2).replace('.', ',');

            // 2. Escrevemos inteligentemente nas posições exatas
            const tr = sheet.getRow(targetRow);
            tr.getCell(colDia).value = diaHoje;
            tr.getCell(colHoras).value = horasJogadas;
            tr.getCell(colFarm).value = parseFloat(profitGoldStr.replace(',', '.'));
            tr.commit();

            // 3. Recalculando todos os totais do mês baseado apenas no que foi preenchido na grade
            let sumFarm = 0;
            let sumMins = 0;
            
            for (let i = headerRow + 1; i <= Math.max(sheet.rowCount, 150); i++) {
                const r = sheet.getRow(i);
                const valDia = r.getCell(colDia).value ? r.getCell(colDia).value.toString().toUpperCase().trim() : '';
                
                // Evita processar a linha do rodapé ao fazer o cálculo
                if (valDia.includes('TOTAL')) continue;
                
                // Soma farm
                const farmText = r.getCell(colFarm).value ? r.getCell(colFarm).value.toString() : '';
                const farmNumber = parseFloat(farmText.replace(',', '.'));
                if (!isNaN(farmNumber)) sumFarm += farmNumber;
                
                // Soma horas
                const horasText = r.getCell(colHoras).value ? r.getCell(colHoras).value.toString() : '';
                if (horasText.includes('h')) {
                    const hParts = horasText.replace('h', '').split(':');
                    if (hParts.length === 2 && !isNaN(parseInt(hParts[0], 10))) {
                        sumMins += parseInt(hParts[0], 10) * 60 + parseInt(hParts[1], 10);
                    }
                }
            }

            const formatTotalFarm = sumFarm.toFixed(2).replace('.', ',');
            const sumH = Math.floor(sumMins / 60);
            const sumM = sumMins % 60;
            const formatTotalHours = `${sumH.toString().padStart(2, '0')}:${sumM.toString().padStart(2, '0')}h`;

            // 4. Preencher o rodapé - Varredura por marcadores 'KK' e 'HORAS'
            for (let i = headerRow; i <= Math.max(sheet.rowCount, 200); i++) {
                const r = sheet.getRow(i);
                let mudou = false;
                for (let c = 1; c <= 30; c++) {
                    const tag = r.getCell(c).value ? r.getCell(c).value.toString().toUpperCase().trim() : '';
                    if (tag === 'KK' || (tag === 'HORAS' && c !== colHoras)) {
                        const targetCell = r.getCell(c + 1);
                        // SÓ ATUALIZA SE NÃO FOR UMA FÓRMULA (Para não quebrar o projeto do usuário)
                        if (targetCell.type !== ExcelJS.ValueType.Formula) {
                            targetCell.value = (tag === 'KK') ? parseFloat(formatTotalFarm.replace(',', '.')) : formatTotalHours;
                            mudou = true;
                        }
                    }
                }
                if (mudou) r.commit();
            }

            // Salva de volta o arquivo
            await workbook.xlsx.writeFile(xlsxPath);
            console.log('✅ Planilha XLSX atualizada e recalculada com Auto-Scan de Estética!');
            // --------------------------------
            
            const xpParaPassar = chatSessions[chatId].totalXp || 0;
            const finalMsg = createFinalDayMessage(chatSessions[chatId].totalBalance, chatSessions[chatId].totalMinutes, xpParaPassar);
            await msg.reply(finalMsg + `\n\n*(📝 Salvo na sua planilha excel do PC!)*`);
            console.log('✅ Resumo final do dia enviado!');
            
            // Zera para o próximo dia
            delete chatSessions[chatId];
            saveSessions(chatSessions);
        } else {
            // Se ele invocar sem ter nenhuma hunt acumulada
            await msg.reply('🤷‍♂️ Não havia nenhuma hunt salva hoje em seu acumulado.');
        }
        return;
    }

    // Comando apenas para pedir a planilha
    if (msg.body.toLowerCase() === 'planilha' || msg.body.toLowerCase() === 'extrato') {
        const fs = require('fs');
        const path = require('path');
        const xlsxPath = path.join(__dirname, 'relatorio_mensal.xlsx');
        
        if (fs.existsSync(xlsxPath)) {
            const planMedia = MessageMedia.fromFilePath(xlsxPath);
            await client.sendMessage(chatId, planMedia, { sendMediaAsDocument: true, caption: '📲 Aqui está a sua planilha XLSX com tudo original!' });
            console.log('✅ Planilha XLSX enviada!');
        } else {
            await msg.reply('Nenhuma planilha foi criada ainda! Feche seu primeiro dia de hunt para que ela seja gerada.');
        }
        return;
    }

    // Comando para recalcular os totais da planilha (caso o usuário tenha editado manual)
    if (msg.body.toLowerCase() === 'recalcular') {
        const ExcelJS = require('exceljs');
        const fs = require('fs');
        const path = require('path');
        const xlsxPath = path.join(__dirname, 'relatorio_mensal.xlsx');

        if (!fs.existsSync(xlsxPath)) {
            await msg.reply('A planilha ainda não existe para recalcular!');
            return;
        }

        try {
            const workbook = new ExcelJS.Workbook();
            await workbook.xlsx.readFile(xlsxPath);
            const sheet = workbook.getWorksheet(1);

            // 1. Localizar cabeçalhos
            let headerRow = -1;
            let colDia = 1, colHoras = 2, colFarm = 3;
            for (let i = 1; i <= Math.max(sheet.rowCount, 50); i++) {
                const r = sheet.getRow(i);
                for (let c = 1; c <= 30; c++) {
                    const val = r.getCell(c).value ? r.getCell(c).value.toString().toUpperCase().trim() : '';
                    if (val === 'DIA') { colDia = c; headerRow = i; }
                    if (val === 'HORAS') { colHoras = c; }
                    if (val === 'FARM') { colFarm = c; }
                }
                if (headerRow !== -1) break;
            }

            if (headerRow === -1) headerRow = 1;

            // 2. Recalcular tudo baseado na grade
            let sumFarm = 0;
            let sumMins = 0;
            
            for (let i = headerRow + 1; i <= Math.max(sheet.rowCount, 150); i++) {
                const r = sheet.getRow(i);
                const valDia = r.getCell(colDia).value ? r.getCell(colDia).value.toString().toUpperCase().trim() : '';
                
                // Ignora linhas vazias ou o rodapé de TOTAL
                if (valDia === '' || valDia.includes('TOTAL')) continue;
                
                // Soma farm (KK)
                const farmText = r.getCell(colFarm).value ? r.getCell(colFarm).value.toString() : '';
                const farmNumber = parseFloat(farmText.replace(',', '.'));
                if (!isNaN(farmNumber)) sumFarm += farmNumber;
                
                // Soma horas (HH:MMh)
                const horasText = r.getCell(colHoras).value ? r.getCell(colHoras).value.toString() : '';
                if (horasText.includes('h')) {
                    const hParts = horasText.replace('h', '').split(':');
                    if (hParts.length === 2 && !isNaN(parseInt(hParts[0], 10))) {
                        sumMins += parseInt(hParts[0], 10) * 60 + parseInt(hParts[1], 10);
                    }
                }
            }

            const formatTotalFarm = sumFarm.toFixed(2).replace('.', ',');
            const sumH = Math.floor(sumMins / 60);
            const sumM = sumMins % 60;
            const formatTotalHours = `${sumH.toString().padStart(2, '0')}:${sumM.toString().padStart(2, '0')}h`;

            // 3. Atualizar o rodapé onde encontrar 'KK' e 'HORAS'
            for (let i = headerRow + 1; i <= Math.max(sheet.rowCount, 200); i++) {
                const r = sheet.getRow(i);
                let mudou = false;
                for (let c = 1; c <= 30; c++) {
                    const tag = r.getCell(c).value ? r.getCell(c).value.toString().toUpperCase().trim() : '';
                    if (tag === 'KK' || (tag === 'HORAS' && c !== colHoras)) {
                        const targetCell = r.getCell(c + 1);
                        if (targetCell.type !== ExcelJS.ValueType.Formula) {
                            targetCell.value = (tag === 'KK') ? parseFloat(formatTotalFarm.replace(',', '.')) : formatTotalHours;
                            mudou = true;
                        }
                    }
                }
                if (mudou) r.commit();
            }

            await workbook.xlsx.writeFile(xlsxPath);
            await msg.reply(`🔄 *Planilha Recalculada!*\n\n📊 Novos Totais:\n💰 *Farm:* ${formatTotalFarm} KK\n⏱️ *Horas:* ${formatTotalHours}\n\n*(Lembre-se de sempre fechar o Excel antes de pedir para o bot salvar!)*`);
            console.log('✅ Planilha recalculada com sucesso via comando!');
        } catch (error) {
            console.error('Erro ao recalcular planilha:', error);
            await msg.reply('❌ Erro ao recalcular! Verifique se a planilha não está aberta no seu PC.');
        }
        return;
    }

    // Comando para somar XP manualmente ao explit atual
    // Uso: "somar xp 1,5kk" | "somar xp 31 1,5kk" | "somar xp 31/03 1,5kk"
    const lowerBody = msg.body.toLowerCase();

    if (lowerBody.startsWith('somar xp ')) {
        const xpArgs = lowerBody.split('somar xp')[1].trim();

        // Tenta detectar se o primeiro token é um dia (ex: "31" ou "31/03")
        // Um "dia" válido é um número de 1-2 dígitos, com ou sem "/MM"
        const diaMatch = xpArgs.match(/^(\d{1,2}(?:\/\d{1,2})?)\s+(.+)$/);

        let diaEspecifico = null;
        let xpStr = xpArgs;

        if (diaMatch) {
            // Extrai apenas o número do dia (ignora o mês se vier "/03")
            diaEspecifico = diaMatch[1].split('/')[0].replace(/^0+/, '') || diaMatch[1].split('/')[0];
            xpStr = diaMatch[2].trim();
        }

        // Parseamento do valor de XP — aceita: 1500000 | 1.5kk | 1,5kk | 500k | 1.5m
        let xpValor = 0;
        if (xpStr.includes('kk')) {
            xpValor = Math.round(parseFloat(xpStr.replace(',', '.').replace('kk', '').trim()) * 1000000);
        } else if (xpStr.includes('k')) {
            xpValor = Math.round(parseFloat(xpStr.replace(',', '.').replace('k', '').trim()) * 1000);
        } else if (xpStr.includes('m')) {
            xpValor = Math.round(parseFloat(xpStr.replace(',', '.').replace('m', '').trim()) * 1000000);
        } else {
            xpValor = parseInt(xpStr.replace(/[.,]/g, ''), 10);
        }

        if (isNaN(xpValor) || xpValor <= 0) {
            await msg.reply('❌ Valor de XP inválido. Exemplos de uso:\n*somar xp 1,5kk*\n*somar xp 31 1,5kk*\n*somar xp 31/03 500k*');
            return;
        }

        if (!chatSessions[chatId]) {
            // Se não houver sessão, cria uma nova com valores zerados
            chatSessions[chatId] = {
                totalBalance: 0,
                totalMinutes: 0,
                totalXp: 0,
                lastDay: diaEspecifico || new Date().getDate().toString()
            };
            console.log(`🆕 Criando novo explit via comando manual de XP em ${chatId}`);
        }

        chatSessions[chatId].totalXp = (chatSessions[chatId].totalXp || 0) + xpValor;

        // Se um dia específico foi informado, atualiza o lastDay da sessão
        if (diaEspecifico) {
            chatSessions[chatId].lastDay = diaEspecifico;
        }

        saveSessions(chatSessions);

        const { formatTibiaCoins, calculateTC } = require('./splitter');
        const sess = chatSessions[chatId];
        const xpTotal = sess.totalXp;
        const profitTotal = sess.totalBalance;
        const minsTotal = sess.totalMinutes;
        const tcTotal = calculateTC(profitTotal);
        
        // Formata as horas
        const hrsSess = Math.floor(minsTotal / 60);
        const minSess = minsTotal % 60;
        const horasFormatadas = `${hrsSess.toString().padStart(2, '0')}:${minSess.toString().padStart(2, '0')}h`;
        
        const isTotalWaste = profitTotal < 0;
        const diaInfo = sess.lastDay ? ` *(dia ${sess.lastDay})*` : '';
        
        let msgReply = `✅ *XP Somada Manualmente!*${diaInfo}\n`;
        msgReply += `✨ Adicionado: +${formatTibiaCoins(xpValor)} XP\n\n`;
        msgReply += `📈 *RESUMO ACUMULADO ATÉ AGORA:*\n`;
        msgReply += `⏱️ *Tempo Total:* ${horasFormatadas}\n`;
        msgReply += `✨ *XP Acumulada:* ${formatTibiaCoins(xpTotal)}\n`;
        msgReply += `💎 *Profit Acumulado:* *${formatTibiaCoins(profitTotal)}* gold ${isTotalWaste ? '🔴' : '🟢'}\n`;
        msgReply += `💵 *TC Acumulado:* ${tcTotal} TC\n`;
        
        await msg.reply(msgReply);
        console.log(`✅ XP manual somada e resumo de sessão enviado: +${xpValor} → Total: ${xpTotal}`);
        return;
    }

    // Comando para apagar e recalcular um dia que deu errado

    if (lowerBody.startsWith('remover dia ') || lowerBody.startsWith('apagar dia ')) {
        const parts = lowerBody.includes('remover') ? lowerBody.split('remover dia') : lowerBody.split('apagar dia');
        const diaRemover = parts[1] ? parts[1].trim() : '';

        if (!diaRemover) {
            await msg.reply('Você precisa especificar o dia corretamente. Exemplo: *remover dia 27*');
            return;
        }

        const ExcelJS = require('exceljs');
        const fs = require('fs');
        const path = require('path');
        const xlsxPath = path.join(__dirname, 'relatorio_mensal.xlsx');

        if (!fs.existsSync(xlsxPath)) {
            await msg.reply('A planilha ainda não existe!');
            return;
        }

        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(xlsxPath);
        const sheet = workbook.getWorksheet(1);

        // Achar colunas mapeadas
        let headerRow = -1;
        let colDia = 1, colHoras = 2, colFarm = 3;
        for (let i = 1; i <= Math.max(sheet.rowCount, 50); i++) {
            const r = sheet.getRow(i);
            for (let c = 1; c <= 30; c++) {
                const val = r.getCell(c).value ? r.getCell(c).value.toString().toUpperCase().trim() : '';
                if (val === 'DIA') { colDia = c; headerRow = i; }
                if (val === 'HORAS') { colHoras = c; }
                if (val === 'FARM') { colFarm = c; }
            }
            if (headerRow !== -1) break;
        }
        if (headerRow === -1) headerRow = 1;

        // Procura a linha com esse dia
        let rowRemovida = -1;
        for (let i = headerRow + 1; i <= Math.max(sheet.rowCount, 150); i++) {
            const r = sheet.getRow(i);
            const valDia = r.getCell(colDia).value ? r.getCell(colDia).value.toString().toUpperCase().trim() : '';
            if (valDia === diaRemover && !valDia.includes('TOTAL')) {
                // Apaga APENAS o valor para preservar as caixinhas azuis, bordas e fórmulas (se houver)
                r.getCell(colDia).value = '';
                r.getCell(colHoras).value = '';
                r.getCell(colFarm).value = '';
                r.commit();
                rowRemovida = i;
                break;
            }
        }

        if (rowRemovida !== -1) {
            // Como apagamos os dados dessa caixa mágica, é hora de reler a tabela para somar o que sobrou!
            let sumFarm = 0;
            let sumMins = 0;
            for (let i = headerRow + 1; i <= Math.max(sheet.rowCount, 150); i++) {
                const r = sheet.getRow(i);
                const valDia = r.getCell(colDia).value ? r.getCell(colDia).value.toString().toUpperCase().trim() : '';
                if (valDia.includes('TOTAL') || valDia === '') continue;

                const farmText = r.getCell(colFarm).value ? r.getCell(colFarm).value.toString() : '';
                const farmNumber = parseFloat(farmText.replace(',', '.'));
                if (!isNaN(farmNumber)) sumFarm += farmNumber;

                const horasText = r.getCell(colHoras).value ? r.getCell(colHoras).value.toString() : '';
                if (horasText.includes('h')) {
                    const hParts = horasText.replace('h', '').split(':');
                    if (hParts.length === 2 && !isNaN(parseInt(hParts[0], 10))) {
                        sumMins += parseInt(hParts[0], 10) * 60 + parseInt(hParts[1], 10);
                    }
                }
            }

            const formatTotalFarm = sumFarm.toFixed(2).replace('.', ',');
            const sumH = Math.floor(sumMins / 60);
            const sumM = sumMins % 60;
            const formatTotalHours = `${sumH.toString().padStart(2, '0')}:${sumM.toString().padStart(2, '0')}h`;

            // Reflete a alteração visualmente atualizando as caixas 'KK' e 'HORAS' no rodapé
            for (let i = Math.max(headerRow, 1); i <= Math.max(sheet.rowCount, 200); i++) {
                const r = sheet.getRow(i);
                let mudou = false;
                for (let c = 1; c <= 30; c++) {
                    const tag = r.getCell(c).value ? r.getCell(c).value.toString().toUpperCase().trim() : '';
                    if (tag === 'KK' || (tag === 'HORAS' && c !== colHoras)) {
                        const targetCell = r.getCell(c + 1);
                        if (targetCell.type !== ExcelJS.ValueType.Formula) {
                            targetCell.value = (tag === 'KK') ? parseFloat(formatTotalFarm.replace(',', '.')) : formatTotalHours;
                            mudou = true;
                        }
                    }
                }
                if (mudou) r.commit();
            }

            await workbook.xlsx.writeFile(xlsxPath);
            await msg.reply(`✅ Feito! 🛡️ Os valores da hunt do dia *${diaRemover}* foram aspirados da sua planilha e os Totais foram atualizados!`);
            console.log(`🧹 Dia ${diaRemover} limpo sob demanda.`);
        } else {
            // Se ele pedir pra apagar um dia que já subiu no pó ou nem jogou
            await msg.reply(`🤷‍♂️ Procurei no seu excel e não encontrei nenhuma anotação de hunt no dia *${diaRemover}*.`);
        }
        return;
    }

    // Comando para adicionar a sessão atual a um dia específico
    if (lowerBody.startsWith('adicionar dia ') || lowerBody.startsWith('salvar dia ')) {
        const parts = lowerBody.includes('adicionar') ? lowerBody.split('adicionar dia') : lowerBody.split('salvar dia');
        const diaAdicionar = parts[1] ? parts[1].trim() : '';

        if (!diaAdicionar) {
            await msg.reply('Você precisa especificar o dia corretamente. Exemplo: *adicionar dia 27*');
            return;
        }

        console.log(`🧹 Comando de salvar detectado para o dia ${diaAdicionar} em ${chatId}`);
        if (chatSessions[chatId] && (chatSessions[chatId].totalBalance !== 0 || chatSessions[chatId].totalMinutes !== 0)) {
            // --- SALVAR NA PLANILHA (XLSX) ---
            const ExcelJS = require('exceljs');
            const fs = require('fs');
            const path = require('path');
            const xlsxPath = path.join(__dirname, 'relatorio_mensal.xlsx');
            
            const workbook = new ExcelJS.Workbook();
            let sheet;
            
            if (fs.existsSync(xlsxPath)) {
                await workbook.xlsx.readFile(xlsxPath);
                sheet = workbook.getWorksheet(1);
            } else {
                sheet = workbook.addWorksheet('Relatório');
                sheet.addRow(['DIA', 'HORAS', 'FARM', 'TC/DIA']);
                sheet.addRow([]);
                sheet.addRow(['TOTAL', 'HORAS', '']);
                sheet.addRow(['', 'KK', '']);
                sheet.addRow(['', 'TC', '']);
            }

            let headerRow = -1;
            let colDia = 1, colHoras = 2, colFarm = 3;
            for (let i = 1; i <= Math.max(sheet.rowCount, 50); i++) {
                const r = sheet.getRow(i);
                for (let c = 1; c <= 30; c++) {
                    const val = r.getCell(c).value ? r.getCell(c).value.toString().toUpperCase().trim() : '';
                    if (val === 'DIA') { colDia = c; headerRow = i; }
                    if (val === 'HORAS') { colHoras = c; }
                    if (val === 'FARM') { colFarm = c; }
                }
                if (headerRow !== -1) break;
            }
            if (headerRow === -1) headerRow = 1;

            let targetRow = headerRow + 1;
            let diaEncontrado = false;
            
            for (let i = headerRow + 1; i <= Math.max(sheet.rowCount, 150); i++) {
                const r = sheet.getRow(i);
                const valDia = r.getCell(colDia).value ? r.getCell(colDia).value.toString().trim().toUpperCase() : '';
                if (valDia === diaAdicionar && !valDia.includes('TOTAL')) {
                    targetRow = i;
                    diaEncontrado = true;
                    break;
                }
            }
            
            if (!diaEncontrado) {
                targetRow = headerRow + 1;
                for (let i = headerRow + 1; i <= Math.max(sheet.rowCount, 150); i++) {
                    const r = sheet.getRow(i);
                    const valDia = r.getCell(colDia).value ? r.getCell(colDia).value.toString().trim().toUpperCase() : '';
                    if (valDia === '' || valDia.includes('TOTAL')) {
                        targetRow = i;
                        break;
                    }
                }
            }

            const tr = sheet.getRow(targetRow);
            
            let currentMinutes = 0;
            let currentBalance = 0;
            
            if (diaEncontrado) {
                const horasText = tr.getCell(colHoras).value ? tr.getCell(colHoras).value.toString() : '';
                if (horasText.includes('h')) {
                    const hParts = horasText.replace('h', '').split(':');
                    if (hParts.length === 2 && !isNaN(parseInt(hParts[0], 10))) {
                        currentMinutes = parseInt(hParts[0], 10) * 60 + parseInt(hParts[1], 10);
                    }
                }
                const farmText = tr.getCell(colFarm).value ? tr.getCell(colFarm).value.toString() : '';
                const farmNumber = parseFloat(farmText.replace(',', '.'));
                if (!isNaN(farmNumber)) {
                    currentBalance = farmNumber * 1000000;
                }
            }
            
            const finalMinutes = currentMinutes + chatSessions[chatId].totalMinutes;
            const finalBalance = currentBalance + chatSessions[chatId].totalBalance;
            
            const hrs = Math.floor(finalMinutes / 60);
            const mins = finalMinutes % 60;
            const horasJogadas = `${hrs.toString().padStart(2, '0')}:${mins.toString().padStart(2, '0')}h`;
            const profitGoldStr = (finalBalance / 1000000).toFixed(2).replace('.', ',');
            
            tr.getCell(colDia).value = diaAdicionar;
            tr.getCell(colHoras).value = horasJogadas;
            tr.getCell(colFarm).value = parseFloat(profitGoldStr.replace(',', '.'));
            tr.commit();

            let sumFarm = 0;
            let sumMins = 0;
            for (let i = headerRow + 1; i <= Math.max(sheet.rowCount, 150); i++) {
                const r = sheet.getRow(i);
                const valDia = r.getCell(colDia).value ? r.getCell(colDia).value.toString().toUpperCase().trim() : '';
                if (valDia.includes('TOTAL') || valDia === '') continue;

                const farmText = r.getCell(colFarm).value ? r.getCell(colFarm).value.toString() : '';
                const farmNumber = parseFloat(farmText.replace(',', '.'));
                if (!isNaN(farmNumber)) sumFarm += farmNumber;

                const horasText = r.getCell(colHoras).value ? r.getCell(colHoras).value.toString() : '';
                if (horasText.includes('h')) {
                    const hParts = horasText.replace('h', '').split(':');
                    if (hParts.length === 2 && !isNaN(parseInt(hParts[0], 10))) {
                        sumMins += parseInt(hParts[0], 10) * 60 + parseInt(hParts[1], 10);
                    }
                }
            }

            const formatTotalFarm = sumFarm.toFixed(2).replace('.', ',');
            const sumH = Math.floor(sumMins / 60);
            const sumM = sumMins % 60;
            const formatTotalHours = `${sumH.toString().padStart(2, '0')}:${sumM.toString().padStart(2, '0')}h`;

            for (let i = Math.max(headerRow, 1); i <= Math.max(sheet.rowCount, 200); i++) {
                const r = sheet.getRow(i);
                let mudou = false;
                for (let c = 1; c <= 30; c++) {
                    const tag = r.getCell(c).value ? r.getCell(c).value.toString().toUpperCase().trim() : '';
                    if (tag === 'KK' || (tag === 'HORAS' && c !== colHoras)) {
                        const targetCell = r.getCell(c + 1);
                        if (targetCell.type !== ExcelJS.ValueType.Formula) {
                            targetCell.value = (tag === 'KK') ? parseFloat(formatTotalFarm.replace(',', '.')) : formatTotalHours;
                            mudou = true;
                        }
                    }
                }
                if (mudou) r.commit();
            }

            await workbook.xlsx.writeFile(xlsxPath);
            
            await msg.reply(`✅ Feito! 🛡️ As hunts salvas foram vinculadas ao dia *${diaAdicionar}* com sucesso!`);
            console.log(`✅ Resumo alocado no dia ${diaAdicionar}!`);
            
            delete chatSessions[chatId];
            saveSessions(chatSessions);
        } else {
            await msg.reply('🤷‍♂️ Não havia nenhuma hunt salva para adicionar a esse dia. Cole um registro do Hunt Analyzer primeiro.');
        }
        return;
    }

    // Comando para FECHAR o explit de um dia que não foi encerrado no momento
    // Uso: "fechar dia 30" → salva o acumulado atual no dia 30 da planilha e zera
    if (lowerBody.startsWith('fechar dia ')) {
        const diaFechar = lowerBody.split('fechar dia')[1].trim();

        if (!diaFechar) {
            await msg.reply('Você precisa especificar o dia. Exemplo: *fechar dia 30*');
            return;
        }

        console.log(`🔒 Comando "fechar dia ${diaFechar}" detectado em ${chatId}`);

        if (chatSessions[chatId] && (chatSessions[chatId].totalBalance !== 0 || chatSessions[chatId].totalMinutes !== 0)) {

            // --- SALVAR NA PLANILHA (XLSX) ---
            const ExcelJS = require('exceljs');
            const xlsxPath = path.join(__dirname, 'relatorio_mensal.xlsx');

            const workbook = new ExcelJS.Workbook();
            let sheet;

            if (fs.existsSync(xlsxPath)) {
                await workbook.xlsx.readFile(xlsxPath);
                sheet = workbook.getWorksheet(1);
            } else {
                sheet = workbook.addWorksheet('Relatório');
                sheet.addRow(['DIA', 'HORAS', 'FARM', 'TC/DIA']);
                sheet.addRow([]);
                sheet.addRow(['TOTAL', 'HORAS', '']);
                sheet.addRow(['', 'KK', '']);
                sheet.addRow(['', 'TC', '']);
            }

            // Localizar cabeçalhos
            let headerRow = -1;
            let colDia = 1, colHoras = 2, colFarm = 3;
            for (let i = 1; i <= Math.max(sheet.rowCount, 50); i++) {
                const r = sheet.getRow(i);
                for (let c = 1; c <= 30; c++) {
                    const val = r.getCell(c).value ? r.getCell(c).value.toString().toUpperCase().trim() : '';
                    if (val === 'DIA') { colDia = c; headerRow = i; }
                    if (val === 'HORAS') { colHoras = c; }
                    if (val === 'FARM') { colFarm = c; }
                }
                if (headerRow !== -1) break;
            }
            if (headerRow === -1) headerRow = 1;

            // Verifica se já existe o dia na planilha (para não duplicar)
            let targetRow = -1;
            let diaEncontrado = false;
            for (let i = headerRow + 1; i <= Math.max(sheet.rowCount, 150); i++) {
                const r = sheet.getRow(i);
                const valDia = r.getCell(colDia).value ? r.getCell(colDia).value.toString().trim().toUpperCase() : '';
                if (valDia === diaFechar && !valDia.includes('TOTAL')) {
                    targetRow = i;
                    diaEncontrado = true;
                    break;
                }
            }

            // Se não achou o dia, pega a primeira linha vazia disponível
            if (!diaEncontrado) {
                for (let i = headerRow + 1; i <= Math.max(sheet.rowCount, 150); i++) {
                    const r = sheet.getRow(i);
                    const valDia = r.getCell(colDia).value ? r.getCell(colDia).value.toString().trim().toUpperCase() : '';
                    if (valDia === '' || valDia.includes('TOTAL')) {
                        targetRow = i;
                        break;
                    }
                }
            }

            // Lê valores já existentes nessa linha (caso esteja atualizando um dia já salvo)
            const tr = sheet.getRow(targetRow);
            let currentMinutes = 0;
            let currentBalance = 0;

            if (diaEncontrado) {
                const horasText = tr.getCell(colHoras).value ? tr.getCell(colHoras).value.toString() : '';
                if (horasText.includes('h')) {
                    const hParts = horasText.replace('h', '').split(':');
                    if (hParts.length === 2 && !isNaN(parseInt(hParts[0], 10))) {
                        currentMinutes = parseInt(hParts[0], 10) * 60 + parseInt(hParts[1], 10);
                    }
                }
                const farmText = tr.getCell(colFarm).value ? tr.getCell(colFarm).value.toString() : '';
                const farmNumber = parseFloat(farmText.replace(',', '.'));
                if (!isNaN(farmNumber)) currentBalance = farmNumber * 1000000;
            }

            // Soma o acumulado da sessão com o que já estava na linha (se existia)
            const finalMinutes = currentMinutes + chatSessions[chatId].totalMinutes;
            const finalBalance = currentBalance + chatSessions[chatId].totalBalance;

            const hrs = Math.floor(finalMinutes / 60);
            const mins = finalMinutes % 60;
            const horasJogadas = `${hrs.toString().padStart(2, '0')}:${mins.toString().padStart(2, '0')}h`;
            const profitGoldStr = (finalBalance / 1000000).toFixed(2).replace('.', ',');

            tr.getCell(colDia).value = diaFechar;
            tr.getCell(colHoras).value = horasJogadas;
            tr.getCell(colFarm).value = parseFloat(profitGoldStr.replace(',', '.'));
            tr.commit();

            // Recalcular totais do mês
            let sumFarm = 0;
            let sumMins = 0;
            for (let i = headerRow + 1; i <= Math.max(sheet.rowCount, 150); i++) {
                const r = sheet.getRow(i);
                const valDia = r.getCell(colDia).value ? r.getCell(colDia).value.toString().toUpperCase().trim() : '';
                if (valDia.includes('TOTAL') || valDia === '') continue;

                const farmText = r.getCell(colFarm).value ? r.getCell(colFarm).value.toString() : '';
                const farmNumber = parseFloat(farmText.replace(',', '.'));
                if (!isNaN(farmNumber)) sumFarm += farmNumber;

                const horasText = r.getCell(colHoras).value ? r.getCell(colHoras).value.toString() : '';
                if (horasText.includes('h')) {
                    const hParts = horasText.replace('h', '').split(':');
                    if (hParts.length === 2 && !isNaN(parseInt(hParts[0], 10))) {
                        sumMins += parseInt(hParts[0], 10) * 60 + parseInt(hParts[1], 10);
                    }
                }
            }

            const formatTotalFarm = sumFarm.toFixed(2).replace('.', ',');
            const sumH = Math.floor(sumMins / 60);
            const sumM = sumMins % 60;
            const formatTotalHours = `${sumH.toString().padStart(2, '0')}:${sumM.toString().padStart(2, '0')}h`;

            // Atualizar rodapé KK e HORAS
            for (let i = Math.max(headerRow, 1); i <= Math.max(sheet.rowCount, 200); i++) {
                const r = sheet.getRow(i);
                let mudou = false;
                for (let c = 1; c <= 30; c++) {
                    const tag = r.getCell(c).value ? r.getCell(c).value.toString().toUpperCase().trim() : '';
                    if (tag === 'KK' || (tag === 'HORAS' && c !== colHoras)) {
                        const targetCell = r.getCell(c + 1);
                        if (targetCell.type !== ExcelJS.ValueType.Formula) {
                            targetCell.value = (tag === 'KK') ? parseFloat(formatTotalFarm.replace(',', '.')) : formatTotalHours;
                            mudou = true;
                        }
                    }
                }
                if (mudou) r.commit();
            }

            await workbook.xlsx.writeFile(xlsxPath);
            console.log(`✅ Explit do dia ${diaFechar} fechado e salvo na planilha!`);

            // Envia o resumo final como "por hoje é só"
            const xpParaPassar = chatSessions[chatId].totalXp || 0;
            const finalMsg = createFinalDayMessage(finalBalance, finalMinutes, xpParaPassar);
            await msg.reply(finalMsg + `\n\n*(📝 Salvo na sua planilha como dia ${diaFechar}!)*`);

            // Zera a sessão
            delete chatSessions[chatId];
            saveSessions(chatSessions);
        } else {
            await msg.reply('🤷‍♂️ Não havia nenhum explit aberto pra fechar. Cole os registros do Hunt Analyzer primeiro!');
        }
        return;
    }

    // Verifica se a mensagem se parece com o formato do Analyzer Solo
    if (msg.body.includes('Balance:') && msg.body.includes('Session:')) {
        console.log(`🔎 Recebido um possível Hunt Analyzer! Analisando...`);
        
        try {
            const stats = parseSoloLog(msg.body);
            if (stats) {
                // Inicia o acumulador local se for a primeira hunt da sessão ou reseta se mudou o dia (mas por segurança, vamos apenas atualizar os valores)
                if (!chatSessions[chatId]) {
                    chatSessions[chatId] = {
                        totalBalance: 0,
                        totalMinutes: 0,
                        totalXp: 0,
                        lastDay: stats.day || new Date().getDate().toString()
                    };
                } else if (stats.day && chatSessions[chatId].lastDay !== stats.day) {
                    // Update the day track if it wasn't tracked, or if the user started pasting a different day?
                    // We'll just force the day tracker to match the last inserted log
                    chatSessions[chatId].lastDay = stats.day;
                }

                // Soma o lucro e converte o tempo em minutos
                chatSessions[chatId].totalBalance += stats.balance;
                
                const playTimeMinutes = parseTimeToMinutes(stats.sessionTime);
                chatSessions[chatId].totalMinutes += playTimeMinutes;
                
                chatSessions[chatId].totalXp = (chatSessions[chatId].totalXp || 0) + (stats.xpGain || 0);

                // Salva o progresso logo após ler o novo log para que não perca
                saveSessions(chatSessions);

                // Cria a mensagem para o bot enviar
                const replyText = createReplyMessage(stats, chatSessions[chatId].totalBalance, chatSessions[chatId].totalMinutes, chatSessions[chatId].totalXp);
                
                // Responde na mesma conversa
                await msg.reply(replyText);
                console.log('✅ Resumo da Hunt + Acumulado enviado!');
            }
        } catch (error) {
            console.error('❌ Erro ao ler o hunt analyzer:', error);
        }
    }
});

// Inicia o bot
client.initialize();
