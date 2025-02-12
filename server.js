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

// Garantir que as pastas existem
['uploads', 'downloads'].forEach(dir => {
    const dirPath = path.join(__dirname, dir);
    if (!fs.existsSync(dirPath)) {
        fs.mkdirSync(dirPath, { recursive: true });
    }
});

app.post('/upload', upload.single('arquivo'), async (req, res) => {
    try {
        // Ler o arquivo Excel
        const workbook = XLSX.readFile(req.file.path);
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const data = XLSX.utils.sheet_to_json(worksheet);

        console.log('Total de registros lidos:', data.length);

        // Filtrar domínios .com.br
        const dominiosComBr = data.filter(row => {
            return Object.values(row).some(value => 
                String(value).toLowerCase().includes('.com.br')
            );
        });

        console.log('Total de domínios .com.br:', dominiosComBr.length);

        // Dividir em grupos de exatamente 500 registros
        const REGISTROS_POR_ARQUIVO = 500;
        const arquivosGerados = [];
        let numeroArquivo = 1;

        // Processar em lotes de 500
        for (let i = 0; i < dominiosComBr.length; i += REGISTROS_POR_ARQUIVO) {
            const registrosDoArquivo = dominiosComBr.slice(i, i + REGISTROS_POR_ARQUIVO);
            
            // Criar novo arquivo Excel
            const novoWorkbook = XLSX.utils.book_new();
            const novaWorksheet = XLSX.utils.json_to_sheet(registrosDoArquivo);
            XLSX.utils.book_append_sheet(novoWorkbook, novaWorksheet, 'Sheet1');

            // Nome do arquivo mais descritivo
            const nomeArquivo = `dominio_comBr_${numeroArquivo}_${registrosDoArquivo.length}_registros.xlsx`;
            const caminhoArquivo = path.join(__dirname, 'downloads', nomeArquivo);

            // Salvar arquivo
            XLSX.writeFile(novoWorkbook, caminhoArquivo);

            console.log(`Criado arquivo ${nomeArquivo} com ${registrosDoArquivo.length} registros`);

            // Adicionar informações do arquivo
            arquivosGerados.push({
                id: numeroArquivo,
                nome: nomeArquivo,
                registros: registrosDoArquivo.length
            });

            numeroArquivo++;
        }

        // Limpar arquivo de upload
        fs.unlinkSync(req.file.path);

        // Enviar resposta com todas as informações
        const resposta = {
            totalRegistros: data.length,
            dominiosComBr: dominiosComBr.length,
            arquivos: arquivosGerados
        };

        console.log('Resposta enviada:', resposta);
        res.json(resposta);

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