const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');

const app = express();
const port = process.env.PORT || 3000;

// Servir arquivos estáticos
app.use(express.static('public'));

// Configurar diretórios temporários para Vercel
const tmpDir = process.env.VERCEL ? '/tmp' : path.join(__dirname, 'uploads');
const downloadsDir = process.env.VERCEL ? '/tmp' : path.join(__dirname, 'downloads');

// Configuração do Multer
const storage = multer.diskStorage({
    destination: function (req, file, cb) {
        cb(null, tmpDir);
    },
    filename: function (req, file, cb) {
        cb(null, Date.now() + '-' + file.originalname);
    }
});

const upload = multer({ storage: storage });

// Criar diretórios necessários
[tmpDir, downloadsDir].forEach(dir => {
    if (!fs.existsSync(dir)) {
        fs.mkdirSync(dir, { recursive: true });
    }
});

app.post('/upload', upload.single('arquivo'), async (req, res) => {
    try {
        if (!req.file) {
            return res.status(400).json({ 
                error: 'Nenhum arquivo foi enviado',
                totalRegistros: 0,
                dominiosComBr: 0,
                arquivos: []
            });
        }

        // Ler o arquivo Excel
        const workbook = XLSX.readFile(req.file.path);
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const data = XLSX.utils.sheet_to_json(worksheet);

        // Filtrar domínios .com.br
        const dominiosComBr = data.filter(row => {
            return Object.values(row).some(value => 
                String(value).toLowerCase().includes('.com.br')
            );
        });

        // Dividir em grupos de 500
        const REGISTROS_POR_ARQUIVO = 500;
        const arquivosGerados = [];
        let numeroArquivo = 1;

        // Processar em lotes de 500
        for (let i = 0; i < dominiosComBr.length; i += REGISTROS_POR_ARQUIVO) {
            const registrosDoArquivo = dominiosComBr.slice(i, i + REGISTROS_POR_ARQUIVO);
            
            const novoWorkbook = XLSX.utils.book_new();
            const novaWorksheet = XLSX.utils.json_to_sheet(registrosDoArquivo);
            XLSX.utils.book_append_sheet(novoWorkbook, novaWorksheet, 'Sheet1');

            const nomeArquivo = `dominio_comBr_${numeroArquivo}_${registrosDoArquivo.length}_registros.xlsx`;
            const caminhoArquivo = path.join(downloadsDir, nomeArquivo);

            XLSX.writeFile(novoWorkbook, caminhoArquivo);

            arquivosGerados.push({
                id: numeroArquivo,
                nome: nomeArquivo,
                registros: registrosDoArquivo.length
            });

            numeroArquivo++;
        }

        // Limpar arquivo de upload
        fs.unlinkSync(req.file.path);

        // Enviar resposta
        return res.json({
            totalRegistros: data.length,
            dominiosComBr: dominiosComBr.length,
            arquivos: arquivosGerados
        });

    } catch (error) {
        console.error('Erro:', error);
        return res.status(500).json({ 
            error: 'Erro ao processar o arquivo: ' + error.message,
            totalRegistros: 0,
            dominiosComBr: 0,
            arquivos: []
        });
    }
});

// Rota para download
app.get('/download/:id/:nome', (req, res) => {
    try {
        const filePath = path.join(downloadsDir, req.params.nome);
        if (fs.existsSync(filePath)) {
            return res.download(filePath, req.params.nome);
        } else {
            return res.status(404).json({ error: 'Arquivo não encontrado' });
        }
    } catch (error) {
        return res.status(500).json({ error: 'Erro ao baixar o arquivo' });
    }
});

// Exportar para Vercel
module.exports = app;

// Iniciar servidor apenas se não estiver no Vercel
if (!process.env.VERCEL) {
    app.listen(port, () => {
        console.log(`Servidor rodando em http://localhost:${port}`);
    });
} 