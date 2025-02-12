const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');

const app = express();
const port = 3000;

// Configuração do Multer para upload de arquivos
const storage = multer.diskStorage({
    destination: function (req, file, cb) {
        const dir = './uploads';
        if (!fs.existsSync(dir)) {
            fs.mkdirSync(dir, { recursive: true });
        }
        cb(null, dir);
    },
    filename: function (req, file, cb) {
        cb(null, Date.now() + '-' + file.originalname);
    }
});

const upload = multer({ storage: storage });

// Servir arquivos estáticos
app.use(express.static('public'));
app.use('/downloads', express.static(path.join(__dirname, 'downloads')));

// Criar diretórios necessários
['uploads', 'downloads'].forEach(dir => {
    if (!fs.existsSync(dir)) {
        fs.mkdirSync(dir, { recursive: true });
    }
});

app.post('/upload', upload.single('arquivo'), async (req, res) => {
    try {
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

        // Definir tamanho exato de cada arquivo
        const REGISTROS_POR_ARQUIVO = 500;
        const arquivosGerados = [];
        
        // Calcular número total de arquivos necessários
        const totalArquivos = Math.ceil(dominiosComBr.length / REGISTROS_POR_ARQUIVO);

        // Dividir em arquivos de exatamente 500 registros
        for (let i = 0; i < totalArquivos; i++) {
            const inicio = i * REGISTROS_POR_ARQUIVO;
            const fim = Math.min(inicio + REGISTROS_POR_ARQUIVO, dominiosComBr.length);
            const chunk = dominiosComBr.slice(inicio, fim);

            const newWorkbook = XLSX.utils.book_new();
            const newWorksheet = XLSX.utils.json_to_sheet(chunk);
            XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, 'Sheet1');

            const fileName = `dominios_comBr_parte_${i + 1}_de_${totalArquivos}.xlsx`;
            const filePath = path.join(__dirname, 'downloads', fileName);

            XLSX.writeFile(newWorkbook, filePath);

            console.log(`Arquivo ${i + 1} criado com ${chunk.length} registros`); // Log para debug

            arquivosGerados.push({
                id: i + 1,
                nome: fileName,
                registros: chunk.length
            });
        }

        // Limpar arquivo de upload
        fs.unlinkSync(req.file.path);

        // Enviar resposta
        res.json({
            totalRegistros: data.length,
            dominiosComBr: dominiosComBr.length,
            arquivos: arquivosGerados
        });

    } catch (error) {
        console.error('Erro:', error);
        res.status(500).json({ 
            error: 'Erro ao processar o arquivo: ' + error.message,
            totalRegistros: 0,
            dominiosComBr: 0,
            arquivos: []
        });
    }
});

// Rota específica para download
app.get('/download/:id/:nome', (req, res) => {
    const fileName = req.params.nome;
    const filePath = path.join(__dirname, 'downloads', fileName);

    console.log(`Tentativa de download: ${filePath}`);

    if (fs.existsSync(filePath)) {
        res.download(filePath, fileName, (err) => {
            if (err) {
                console.error('Erro no download:', err);
                res.status(500).send('Erro ao baixar o arquivo');
            }
        });
    } else {
        console.error('Arquivo não encontrado:', filePath);
        res.status(404).send('Arquivo não encontrado');
    }
});

app.listen(port, () => {
    console.log(`Servidor rodando em http://localhost:${port}`);
    console.log(`Diretório atual: ${__dirname}`);
}); 