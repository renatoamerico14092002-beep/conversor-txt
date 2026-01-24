document.addEventListener('DOMContentLoaded', function() {
    // Elementos DOM
    const fileInput = document.getElementById('fileInput');
    const dropArea = document.getElementById('dropArea');
    const convertBtn = document.getElementById('convertBtn');
    const downloadBtn = document.getElementById('downloadBtn');
    const clearBtn = document.getElementById('clearBtn');
    const previewText = document.getElementById('previewText');
    const stats = document.getElementById('stats');
    
    // Instância do conversor
    const converter = new ExcelToTxtConverter();
    
    // Estado da aplicação
    let currentFile = null;
    let excelData = null;
    
    // Event Listeners para drag and drop
    ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
        dropArea.addEventListener(eventName, preventDefaults, false);
    });
    
    function preventDefaults(e) {
        e.preventDefault();
        e.stopPropagation();
    }
    
    ['dragenter', 'dragover'].forEach(eventName => {
        dropArea.addEventListener(eventName, highlight, false);
    });
    
    ['dragleave', 'drop'].forEach(eventName => {
        dropArea.addEventListener(eventName, unhighlight, false);
    });
    
    function highlight() {
        dropArea.classList.add('dragover');
    }
    
    function unhighlight() {
        dropArea.classList.remove('dragover');
    }
    
    // Processar arquivo dropado
    dropArea.addEventListener('drop', handleDrop, false);
    
    function handleDrop(e) {
        const dt = e.dataTransfer;
        const files = dt.files;
        
        if (files.length > 0) {
            handleFile(files[0]);
        }
    }
    
    // Processar seleção de arquivo
    fileInput.addEventListener('change', function(e) {
        if (e.target.files.length > 0) {
            handleFile(e.target.files[0]);
        }
    });
    
    function handleFile(file) {
        currentFile = file;
        
        // Verificar se é um arquivo Excel
        if (!file.name.match(/\.(xlsx|xls|csv)$/i)) {
            alert('Por favor, selecione um arquivo Excel (.xlsx, .xls) ou CSV (.csv)');
            return;
        }
        
        // Ler o arquivo
        const reader = new FileReader();
        
        reader.onload = function(e) {
            try {
                const data = e.target.result;
                let workbook;
                
                if (file.name.match(/\.csv$/i)) {
                    // Para arquivos CSV
                    const csvData = XLSX.read(data, { type: 'binary', cellDates: true });
                    workbook = csvData;
                } else {
                    // Para arquivos Excel
                    workbook = XLSX.read(data, { type: 'binary', cellDates: true });
                }
                
                // Pegar a primeira planilha
                const firstSheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[firstSheetName];
                
                // Converter para JSON
                excelData = XLSX.utils.sheet_to_json(worksheet, { defval: '' });
                
                // Log para debug
                console.log('Dados carregados:', excelData.length, 'linhas');
                if (excelData.length > 0) {
                    console.log('Primeira linha:', excelData[0]);
                }
                
                // Habilitar botão de conversão
                convertBtn.disabled = false;
                
                // Mostrar informações do arquivo
                const fileInfo = `${file.name} (${excelData.length} linhas)`;
                dropArea.querySelector('h3').textContent = `📁 ${fileInfo}`;
                
                // Resetar preview
                previewText.value = '';
                stats.textContent = '0 linhas processadas';
                downloadBtn.disabled = true;
                
            } catch (error) {
                console.error('Erro ao processar arquivo:', error);
                alert('Erro ao processar o arquivo. Verifique se é um arquivo Excel válido. Detalhes no console.');
            }
        };
        
        reader.onerror = function() {
            alert('Erro ao ler o arquivo.');
        };
        
        reader.readAsBinaryString(file);
    }
    
    // Converter para TXT
    convertBtn.addEventListener('click', function() {
        if (!excelData) {
            alert('Nenhum arquivo carregado.');
            return;
        }
        
        try {
            const result = converter.convertExcelData(excelData);
            
            // Mostrar preview (apenas as primeiras 20 linhas para performance)
            const previewLines = result.content.split('\n').slice(0, 20).join('\n');
            previewText.value = previewLines;
            
            if (result.lineCount > 20) {
                previewText.value += `\n... (${result.lineCount - 20} linhas adicionais)`;
            }
            
            // Atualizar estatísticas
            stats.textContent = `${result.lineCount} linhas processadas (101 caracteres cada)`;
            
            // Habilitar botão de download
            downloadBtn.disabled = false;
            
            // Rolar para o preview
            previewText.scrollIntoView({ behavior: 'smooth' });
            
            // Verificar comprimento da primeira linha
            const firstLine = result.content.split('\n')[0];
            if (firstLine && firstLine.length !== 101) {
                console.warn(`Atenção: A primeira linha tem ${firstLine.length} caracteres, esperado 101.`);
            }
            
        } catch (error) {
            console.error('Erro na conversão:', error);
            alert('Erro na conversão: ' + error.message);
        }
    });
    
    // Baixar arquivo TXT
    downloadBtn.addEventListener('click', function() {
        try {
            const filename = currentFile ? 
                currentFile.name.replace(/\.[^/.]+$/, "") + '_convertido.txt' : 
                'converted_file.txt';
            
            converter.downloadTxt(filename);
            
            // Feedback visual
            const originalText = stats.textContent;
            stats.textContent = '✅ Arquivo baixado com sucesso!';
            
            setTimeout(() => {
                stats.textContent = originalText;
            }, 3000);
            
        } catch (error) {
            alert('Erro ao baixar arquivo: ' + error.message);
        }
    });
    
    // Limpar tudo
    clearBtn.addEventListener('click', function() {
        // Resetar todos os elementos
        fileInput.value = '';
        previewText.value = '';
        stats.textContent = '0 linhas processadas';
        
        // Resetar botões
        convertBtn.disabled = true;
        downloadBtn.disabled = true;
        
        // Resetar área de upload
        dropArea.querySelector('h3').textContent = 'Arraste e solte seu arquivo Excel aqui';
        
        // Resetar estado
        currentFile = null;
        excelData = null;
        converter.txtContent = '';
        converter.lineCount = 0;
    });
    
    // Copiar para área de transferência (opcional)
    previewText.addEventListener('click', function() {
        if (previewText.value) {
            previewText.select();
            document.execCommand('copy');
            
            // Feedback visual
            const originalText = stats.textContent;
            stats.textContent = '📋 Copiado para área de transferência!';
            
            setTimeout(() => {
                stats.textContent = originalText;
            }, 2000);
        }
    });
    
    // Validar colunas do Excel
    function validateExcelColumns(firstRow) {
        const expectedColumns = ['categoria', 'codigo_item', 'parte_codigo', 'data', 'periodo', 'valor1', 'valor2', 'valor3', 'valor4'];
        const rowKeys = Object.keys(firstRow);
        
        console.log('Colunas encontradas:', rowKeys);
        console.log('Colunas esperadas:', expectedColumns);
        
        // Verifica se temos pelo menos algumas das colunas esperadas
        const missingColumns = expectedColumns.filter(col => !rowKeys.includes(col));
        
        if (missingColumns.length > 0) {
            console.warn('Colunas faltando:', missingColumns);
        }
        
        return true;
    }
});
