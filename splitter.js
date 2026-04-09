function parseSoloLog(log) {
    // Vamos ser mais flexíveis e não depender de "includes" exatos que podem quebrar no salto de linha do WhatsApp
    // Extrai o tempo da Sessão

    const sessionMatch = log.match(/Session:\s+(.+)/);
    // Extrai as informações de economia
    const lootMatch = log.match(/Loot:\s+([\d,.-]+)/);
    const suppliesMatch = log.match(/Supplies:\s+([\d,.-]+)/);
    const balanceMatch = log.match(/Balance:\s+([\d,.-]+)/);
    
    // Tenta extrair a XP Gain (Pode não existir, alguns colam logs antigos)
    const xpMatch = log.match(/XP Gain:\s+([\d,.-]+)/);
    
    // Tenta extrair a data do log (ex: "Session data: From 2024-03-31")
    const dateMatch = log.match(/(?:From|De)\s+(\d{4}-\d{2}-\d{2})/);
    let dayExtracted = null;
    if (dateMatch) {
        dayExtracted = parseInt(dateMatch[1].split('-')[2], 10).toString();
    }

    if (!sessionMatch || !balanceMatch) return null;

    const sessionTime = sessionMatch[1].trim();
    const balanceStr = balanceMatch[1].replace(/[,.]/g, '');
    const balance = parseInt(balanceStr, 10);
    
    let lootStr = lootMatch ? lootMatch[1].replace(/[,.]/g, '') : "0";
    let suppliesStr = suppliesMatch ? suppliesMatch[1].replace(/[,.]/g, '') : "0";
    let xpStr = xpMatch ? xpMatch[1].replace(/[,.]/g, '') : "0";
    
    return {
        sessionTime,
        balance,
        loot: parseInt(lootStr, 10),
        supplies: parseInt(suppliesStr, 10),
        xpGain: parseInt(xpStr, 10),
        day: dayExtracted
    };
}

function formatTibiaCoins(amount) {
    return amount.toLocaleString('en-US');
}

// Calcula quantos Tibia Coins (TC) equivalem ao profit
// Divide o profit (em gold) por 39.000
function calculateTC(profitGold) {
    return Math.floor(profitGold / 39000);
}

// Converte "02:15h" em minutos totais "135"
function parseTimeToMinutes(timeStr) {
    const parts = timeStr.replace('h', '').split(':');
    return parseInt(parts[0], 10) * 60 + parseInt(parts[1], 10);
}

// Converte 135 minutos em "02:15h"
function formatMinutesToTime(totalMinutes) {
    const hrs = Math.floor(totalMinutes / 60);
    const mins = totalMinutes % 60;
    return `${hrs.toString().padStart(2, '0')}:${mins.toString().padStart(2, '0')}h`;
}

function createReplyMessage(stats, currentTotalBalance, currentTotalMinutes, currentTotalXp = 0) {
    const isWaste = stats.balance < 0;
    const isTotalWaste = currentTotalBalance < 0;
    
    const tcHunt = calculateTC(stats.balance);
    const tcTotal = calculateTC(currentTotalBalance);

    let msg = `🕒 *Tempo (Essa Hunt):* ${stats.sessionTime}\n`;
    if (stats.xpGain > 0) msg += `✨ *XP Gain (Essa Hunt):* ${formatTibiaCoins(stats.xpGain)}\n`;
    msg += `⚖️ *Profit (Essa Hunt):* ${formatTibiaCoins(stats.balance)} gold ${isWaste ? '🔴' : '🟢'}\n`;
    msg += `💵 *TC (Essa Hunt):* ${tcHunt} TC\n`;
    msg += `------------------------\n`;
    msg += `📈 *RESUMO ACUMULADO HOJE*\n`;
    msg += `⏱️ *Tempo Jogado Total:* ${formatMinutesToTime(currentTotalMinutes)}\n`;
    if (currentTotalXp > 0) msg += `✨ *XP Acumulada:* ${formatTibiaCoins(currentTotalXp)}\n`;
    msg += `💎 *Profit Acumulado:* *${formatTibiaCoins(currentTotalBalance)}* gold ${isTotalWaste ? '🔴 (Waste)' : '🟢 (Profit)'}\n`;
    msg += `💵 *TC Acumulado:* ${tcTotal} TC\n`;

    return msg;
}

function createFinalDayMessage(totalBalance, totalMinutes, totalXp = 0) {
    const isTotalWaste = totalBalance < 0;
    
    let msg = `========================\n`;
    msg += `🏆 *RESUMO DO DIA FECHADO!* 🏆\n`;
    msg += `========================\n`;
    const tcFinal = calculateTC(totalBalance);

    msg += `⏱️ *Horas jogadas:* ${formatMinutesToTime(totalMinutes)}\n`;
    if (totalXp > 0) msg += `✨ *XP Total Ganha:* ${formatTibiaCoins(totalXp)}\n`;
    msg += `⚖️ *Saldo Final:* *${formatTibiaCoins(totalBalance)}* gold ${isTotalWaste ? '🔴' : '🟢'}\n`;
    msg += `💵 *TC do Dia:* ${tcFinal} TC\n\n`;
    msg += `A calculadora foi zerada para as hunts de amanhã! PARABENS JHONS`;
    
    return msg;
}

module.exports = { parseSoloLog, createReplyMessage, createFinalDayMessage, parseTimeToMinutes, calculateTC, formatTibiaCoins };
