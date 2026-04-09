# 🎮 Tibia WhatsApp Analyzer Bot

<div align="center">
  <img src="https://logodownload.org/wp-content/uploads/2017/04/whatsapp-logo-1.png" width="100"/>
  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
  <img src="https://upload.wikimedia.org/wikipedia/commons/4/41/Tibia_Logo.png" width="150"/>
</div>

<br/>

Este projeto é um **Bot para WhatsApp** desenvolvido em Node.js com a biblioteca `whatsapp-web.js`. Ele automatiza a leitura, o cálculo e o registro financeiro de sessões de caça (hunts) do jogo **Tibia**. O bot escuta mensagens contendo dados de *Hunt Analyzer*, processa essas informações e escreve o relatório consolidado automaticamente em uma planilha XLSX.

---

## 🚀 Funcionalidades

- **Leitura Automática do Hunt Analyzer**: Identifica logs de hunt colados pelo usuário no WhatsApp.
- **Processamento de Estatísticas**:
  - Consolida XP total acumulada na sessão.
  - Soma o Profit Total ou registra Wastage (prejuízo).
  - Calcula o tempo exato investido.
  - Converte instantaneamente o Profit para "Tibia Coins (TC)".
- **Gestão em Planilha (ExcelJS)**:
  - Gera um documento `.xlsx` de relatório mensal.
  - Posiciona a caçada de cada dia na respectiva coluna.
  - Auto-calcula tudo e atualiza o rodapé de "Totais" do mês.
- **Comandos via Chat**:
  - `por hoje é só` / `fechar dia [dia]`: Salva os logs parciais do dia na planilha.
  - `somar xp [valor]`: Adiciona experiência extra manualmente à conta provisória.
  - `planilha` / `extrato`: O bot responde enviando a planilha diretamente pelo WhatsApp.
  - `recalcular`: Força a releitura da planilha XLSX para atualizar células e totais.
  - `apagar dia [dia]`: Remove dados de um dia específico que deu errado.

## 🛠️ Tecnologias Utilizadas

- **JavaScript (Node.js)**
- [whatsapp-web.js](https://wwebjs.dev/) - Conexão e interface da API do WhatsApp Web.
- [ExcelJS](https://github.com/exceljs/exceljs) - Leitura, injeção e re-cálculo da planilha XLSX mantendo as formatações nativas.
- [qrcode-terminal](https://www.npmjs.com/package/qrcode-terminal) - Autenticação da sessão do WhatsApp direto no terminal via QR Code.

## ⚙️ Como executar

1. Clone o repositório em sua máquina:
   ```bash
   git clone https://github.com/SEU_USUARIO/tibia-whatsapp-bot.git
   ```

2. Na pasta do projeto, instale as dependências:
   ```bash
   npm install
   ```

3. Inicie a aplicação no terminal:
   ```bash
   npm start
   # ou: node index.js
   ```

4. **Conexão:** O terminal exibirá um QR Code. Abra o seu WhatsApp no celular e leia o código QR em "Dispositivos Conectados".  
5. **Pronto!** Envie uma mensagem com texto de "Session data" do Tibia para ver o robô atuar.

## 📌 Autor

Desenvolvido para gerenciar caçadas intensas de forma profissional 🧙‍♂️.
