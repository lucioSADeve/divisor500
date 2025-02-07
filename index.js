const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const fs = require('fs-extra');
const path = require('path');

const app = express();
const upload = multer({ dest: 'entrada/' });

// Configurações
const LINHAS_POR_ARQUIVO = 5000;
const PASTA_ENTRADA = './entrada';
const PASTA_SAIDA = './saida';

// Configurar pasta de arquivos estáticos
app.use(express.static('public'));

// Rota principal - página de upload
app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// Rota para upload do arquivo
app.post('/upload', upload.single('arquivo'), async (req, res) => {
    if (!req.file) {
        return res.status(400).json({ error: 'Nenhum arquivo enviado' });
    }

    try {
        const resultado = await dividirArquivo(req.file.filename);
        res.json(resultado);
    } catch (erro) {
        console.error(erro);
        res.status(500).json({ error: 'Erro ao processar arquivo' });
    }
});

// Rota para download dos arquivos
app.get('/download/:arquivo', (req, res) => {
    const arquivo = path.join(PASTA_SAIDA, req.params.arquivo);
    res.download(arquivo);
});

// Função para dividir o arquivo
async function dividirArquivo(nomeArquivo) {
    try {
        const workbook = XLSX.readFile(path.join(PASTA_ENTRADA, nomeArquivo));
        const planilha = workbook.Sheets[workbook.SheetNames[0]];
        const dados = XLSX.utils.sheet_to_json(planilha);

        const totalArquivos = Math.ceil(dados.length / LINHAS_POR_ARQUIVO);
        const arquivosGerados = [];

        await fs.ensureDir(PASTA_SAIDA);

        for (let i = 0; i < totalArquivos; i++) {
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
        throw erro;
    }
}

// Inicia o servidor
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`Servidor rodando em http://localhost:${PORT}`);
});