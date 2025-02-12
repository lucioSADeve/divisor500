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

        // Configuração fixa de 500 registros
        const REGISTROS_POR_ARQUIVO = 500;
        const arquivosGerados = [];

        // Dividir em grupos de 500
        const grupos = [];
        for (let i = 0; i < dominiosComBr.length; i += REGISTROS_POR_ARQUIVO) {
            grupos.push(dominiosComBr.slice(i, i + REGISTROS_POR_ARQUIVO));
        }

        // Criar um arquivo para cada grupo
        for (let i = 0; i < grupos.length; i++) {
            const grupo = grupos[i];
            
            // Criar novo workbook para este grupo
            const novoWorkbook = XLSX.utils.book_new();
            const novaWorksheet = XLSX.utils.json_to_sheet(grupo);
            XLSX.utils.book_append_sheet(novoWorkbook, novaWorksheet, 'Sheet1');

            // Nome do arquivo com número de registros
            const nomeArquivo = `parte_${i + 1}_de_${grupos.length}_${grupo.length}_registros.xlsx`;
            const caminhoArquivo = path.join(__dirname, 'downloads', nomeArquivo);

            // Salvar arquivo
            XLSX.writeFile(novoWorkbook, caminhoArquivo);

            // Adicionar às informações de arquivos gerados
            arquivosGerados.push({
                id: i + 1,
                nome: nomeArquivo,
                registros: grupo.length
            });

            console.log(`Arquivo ${i + 1} criado com ${grupo.length} registros`);
        }

        // Limpar arquivo de upload
        fs.unlinkSync(req.file.path);

        // Retornar informações completas
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