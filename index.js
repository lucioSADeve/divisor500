const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const path = require('path');
const session = require('express-session');
const config = require('./config');

const app = express();

// Configuração da sessão
app.use(session(config.session));

// Configuração do Multer para memória
const upload = multer({
    storage: multer.memoryStorage(),
    limits: {
        fileSize: 50 * 1024 * 1024 // Limite de 50MB
    }
});

// Configurações
const LINHAS_POR_ARQUIVO = 5000;

// Configurar pasta de arquivos estáticos
app.use(express.static('public'));
app.use(express.json());

// Rota principal
app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// Rota para upload e processamento
app.post('/upload', upload.single('arquivo'), async (req, res) => {
    console.log('Iniciando upload...');
    
    if (!req.file) {
        return res.status(400).json({ error: 'Nenhum arquivo enviado' });
    }

    try {
        console.log('Processando arquivo:', req.file.originalname);
        
        // Lê o arquivo do buffer
        const workbook = XLSX.read(req.file.buffer);
        const planilha = workbook.Sheets[workbook.SheetNames[0]];
        const dados = XLSX.utils.sheet_to_json(planilha);

        console.log('Total de linhas:', dados.length);
        const totalArquivos = Math.ceil(dados.length / LINHAS_POR_ARQUIVO);
        const arquivosGerados = [];

        // Processa cada parte
        for (let i = 0; i < totalArquivos; i++) {
            const inicio = i * LINHAS_POR_ARQUIVO;
            const fim = Math.min((i + 1) * LINHAS_POR_ARQUIVO, dados.length);
            const dadosParte = dados.slice(inicio, fim);

            const novoWorkbook = XLSX.utils.book_new();
            const novaPlanilha = XLSX.utils.json_to_sheet(dadosParte);
            XLSX.utils.book_append_sheet(novoWorkbook, novaPlanilha, 'Planilha1');

            // Gera o buffer do arquivo
            const buffer = XLSX.write(novoWorkbook, { type: 'buffer', bookType: 'xlsx' });
            
            // Nome do arquivo
            const nomeArquivo = `${path.parse(req.file.originalname).name}_${i + 1}.xlsx`;
            arquivosGerados.push({
                nome: nomeArquivo,
                buffer: buffer
            });
        }

        // Armazena os buffers na sessão
        req.session.arquivos = arquivosGerados;

        // Retorna apenas os nomes dos arquivos
        res.json({
            success: true,
            arquivos: arquivosGerados.map(a => a.nome),
            total: totalArquivos
        });

    } catch (erro) {
        console.error('Erro no processamento:', erro);
        res.status(500).json({ 
            error: 'Erro ao processar arquivo',
            details: erro.message
        });
    }
});

// Rota para download
app.get('/download/:arquivo', (req, res) => {
    try {
        const nomeArquivo = req.params.arquivo;
        const arquivo = req.session?.arquivos?.find(a => a.nome === nomeArquivo);

        if (!arquivo) {
            return res.status(404).json({ error: 'Arquivo não encontrado' });
        }

        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', `attachment; filename="${arquivo.nome}"`);
        res.send(arquivo.buffer);

    } catch (erro) {
        console.error('Erro no download:', erro);
        res.status(500).json({ error: 'Erro ao baixar arquivo' });
    }
});

// Inicia o servidor
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`Servidor rodando na porta ${PORT}`);
});

// Tratamento de erros não capturados
process.on('unhandledRejection', (erro) => {
    console.error('Erro não tratado:', erro);
});