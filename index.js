const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const fs = require('fs-extra');
const path = require('path');

const app = express();

// Configuração do Multer com tratamento de erro
const storage = multer.diskStorage({
    destination: function (req, file, cb) {
        fs.ensureDir('/tmp/uploads').then(() => {
            cb(null, '/tmp/uploads');
        }).catch(err => {
            console.error('Erro ao criar diretório:', err);
            cb(err);
        });
    },
    filename: function (req, file, cb) {
        cb(null, Date.now() + '-' + file.originalname);
    }
});

const upload = multer({ storage: storage });

// Configurações
const LINHAS_POR_ARQUIVO = 5000;
const PASTA_ENTRADA = '/tmp/entrada';
const PASTA_SAIDA = '/tmp/saida';

// Configurar pasta de arquivos estáticos
app.use(express.static('public'));

// Middleware para processar JSON e form data
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

// Rota principal
app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// Rota para upload do arquivo
app.post('/upload', upload.single('arquivo'), async (req, res) => {
    console.log('Iniciando upload...');
    
    if (!req.file) {
        console.log('Nenhum arquivo recebido');
        return res.status(400).json({ error: 'Nenhum arquivo enviado' });
    }

    try {
        console.log('Arquivo recebido:', req.file.originalname);
        
        // Garante que os diretórios existam
        await fs.ensureDir(PASTA_ENTRADA);
        await fs.ensureDir(PASTA_SAIDA);
        
        // Move o arquivo para o diretório de entrada
        const caminhoEntrada = path.join(PASTA_ENTRADA, req.file.filename);
        await fs.move(req.file.path, caminhoEntrada);
        
        console.log('Iniciando processamento do arquivo...');
        const resultado = await dividirArquivo(req.file.filename);
        
        console.log('Processamento concluído:', resultado);
        res.json(resultado);
    } catch (erro) {
        console.error('Erro no processamento:', erro);
        res.status(500).json({ 
            error: 'Erro ao processar arquivo',
            details: erro.message
        });
    }
});

// Rota para download
app.get('/download/:arquivo', async (req, res) => {
    try {
        const arquivo = path.join(PASTA_SAIDA, req.params.arquivo);
        if (await fs.exists(arquivo)) {
            res.download(arquivo);
        } else {
            res.status(404).json({ error: 'Arquivo não encontrado' });
        }
    } catch (erro) {
        console.error('Erro no download:', erro);
        res.status(500).json({ error: 'Erro ao baixar arquivo' });
    }
});

// Função para dividir o arquivo
async function dividirArquivo(nomeArquivo) {
    try {
        console.log('Lendo arquivo Excel...');
        const workbook = XLSX.readFile(path.join(PASTA_ENTRADA, nomeArquivo));
        const planilha = workbook.Sheets[workbook.SheetNames[0]];
        const dados = XLSX.utils.sheet_to_json(planilha);

        console.log('Total de linhas:', dados.length);
        const totalArquivos = Math.ceil(dados.length / LINHAS_POR_ARQUIVO);
        const arquivosGerados = [];

        for (let i = 0; i < totalArquivos; i++) {
            console.log(`Processando parte ${i + 1} de ${totalArquivos}...`);
            const inicio = i * LINHAS_POR_ARQUIVO;
            const fim = Math.min((i + 1) * LINHAS_POR_ARQUIVO, dados.length);
            const dadosParte = dados.slice(inicio, fim);

            const novoWorkbook = XLSX.utils.book_new();
            const novaPlanilha = XLSX.utils.json_to_sheet(dadosParte);
            XLSX.utils.book_append_sheet(novoWorkbook, novaPlanilha, 'Planilha1');

            const nomeBase = path.basename(nomeArquivo, path.extname(nomeArquivo));
            const arquivoSaida = `${nomeBase}_${i + 1}.xlsx`;
            const caminhoCompleto = path.join(PASTA_SAIDA, arquivoSaida);

            XLSX.writeFile(novoWorkbook, caminhoCompleto);
            arquivosGerados.push(arquivoSaida);
        }

        // Limpa o arquivo de entrada
        await fs.unlink(path.join(PASTA_ENTRADA, nomeArquivo));

        return {
            success: true,
            arquivos: arquivosGerados,
            total: totalArquivos
        };

    } catch (erro) {
        console.error('Erro ao dividir arquivo:', erro);
        throw erro;
    }
}

// Inicia o servidor
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`Servidor rodando na porta ${PORT}`);
});

// Tratamento de erros não capturados
process.on('unhandledRejection', (erro) => {
    console.error('Erro não tratado:', erro);
});