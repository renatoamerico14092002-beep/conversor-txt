document.addEventListener('DOMContentLoaded', function() {
    console.log('Aplicação iniciada');
    
    // Elementos DOM
    const fileInput = document.getElementById('fileInput');
    const dropArea = document.getElementById('dropArea');
    const convertBtn = document.getElementById('convertBtn');
    const downloadBtn = document.getElementById('downloadBtn');
    const clearBtn = document.getElementById('clearBtn');
    const previewText = document.getElementById('previewText');
    const stats = document.getElementById('stats');
    const fileInfo = document.getElementById('fileInfo');
    const loading = document.getElementById('loading');
    
    // Instância do conversor
    const converter = new ExcelToTxtConverter();
    
    // Estado da aplicação
    let currentFile = null;
    let excelData = null;
    
    // Verificar se XLSX está disponível
    if (typeof XLSX === 'undefined') {
        alert('Erro: Biblioteca para leitura de Excel não carregada. Verifique sua conexão com a internet.');
        return;
    }
    
    // Funções auxiliares
    function showLoading(show) {
        loading.style.display = show ? 'flex' : 'none';
    }
    
    function showMessage(message, type = 'info') {
        // Remove mensagens anteriores
        const existingMessages = document.querySelectorAll('.temp-message');
        existingMessages.forEach(msg => msg.remove());
        
        if (message) {
            const messageDiv = document.createElement('div');
            messageDiv.className = `temp-message ${type === 'error' ? 'error-message' : 'success-message'}`;
            messageDiv.textContent = message;
            dropArea.appendChild(messageDiv);
            
            // Remove após 5 segundos
            setTimeout(() => {
                if (messageDiv.parentNode) {
                    messageDiv.remove();
                }
            }, 5000);
        }
    }
    
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
    dropArea.addEventListener('drop', function(e) {
        const dt = e.dataTransfer;
        const files = dt.files;
        
        if (files.length > 0) {
            handleFile(files[0]);
        }
    });
    
    // Processar seleção de arquivo
    fileInput.addEventListener('change', function(e) {
        if (e.target.files.length > 0) {
            handleFile(e.target.files[0]);
        }
    });
    
    // Também permitir clicar em toda a área de drop
    dropArea.addEventListener('click', function(e) {
        if (e.target.tagName !== 'INPUT' && e.target.tagName !== 'BUTTON' && e.target.tagName !== 'A') {
            fileInput.click();
        }
    });
    
    function handleFile(file) {
        console.log('Processando arquivo:', file.name);
        currentFile = file;
        
        // Verificar se é um arquivo Excel
        const validExtensions = ['.xlsx', '.xls', '.csv'];
        const fileExt = file.name.toLowerCase().substring(file.name.lastIndexOf('.'));
        
        if (!validExtensions.includes(fileExt)) {
            showMessage('Formato de arquivo não suportado. Use .xlsx, .xls ou .csv.', 'error');
            return;
        }
        
        // Verificar tamanho do arquivo (limite de 10MB)
        if (file.size > 10 * 1024 * 1024) {
            showMessage('Arquivo muito grande. O limite é 10MB.', 'error');
            return;
        }
        
        // Mostrar loading
        showLoading(true);
        fileInfo.textContent = `Processando: ${file.name}`;
        
        // Ler o arquivo
        const reader = new FileReader();
        
        reader.onload = function(e) {
            try {
                const data = e.target.result;
                let workbook;
                
                // Determinar o tipo de leitura baseado na extensão
                if (fileExt === '.csv') {
                    workbook = XLSX.read(data, { 
                        type: 'binary', 
                        cellDates: true,
                        dateNF: 'dd/mm/yyyy'
                    });
                } else {
                    workbook = XLSX.read(data, { 
                        type: 'binary', 
                        cellDates: true,
                        dateNF: 'dd/mm/yyyy'
                    });
                }
                
                // Pegar a primeira planilha
                const firstSheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[firstSheetName];
                
                // Converter para JSON
                excelData = XLSX.utils.sheet_to_json(worksheet, { 
                    defval: '',
                    raw: false,
                    dateNF: 'dd/mm/yyyy'
                });
                
                console.log('Dados carregados:', excelData);
                console.log('Número de linhas:', excelData.length);
                
                if (excelData.length > 0) {
                    console.log('Primeira linha:', excelData[0]);
                    console.log('Colunas disponíveis:', Object.keys(excelData[0]));
                }
                
                // Habilitar botão de conversão
                convertBtn.disabled = false;
                
                // Atualizar informações do arquivo
                fileInfo.textContent = `${file.name} (${excelData.length} linhas carregadas)`;
                showMessage('Arquivo carregado com sucesso! Clique em "Converter para TXT" para processar.', 'success');
                
            } catch (error) {
                console.error('Erro ao processar arquivo:', error);
                showMessage('Erro ao processar o arquivo: ' + error.message, 'error');
                fileInfo.textContent = 'Arraste e solte seu arquivo Excel aqui';
            } finally {
                showLoading(false);
            }
        };
        
        reader.onerror = function() {
            showLoading(false);
            showMessage('Erro ao ler o arquivo. Tente novamente.', 'error');
            fileInfo.textContent = 'Arraste e solte seu arquivo Excel aqui';
        };
        
        reader.onprogress = function(e) {
            if (e.lengthComputable) {
                const percent = Math.round((e.loaded / e.total) * 100);
                fileInfo.textContent = `Carregando: ${percent}%`;
            }
        };
        
        // Usar readAsArrayBuffer para melhor compatibilidade
        reader.readAsArrayBuffer(file);
    }
    
    // Converter para TXT
    convertBtn.addEventListener('click', function() {
        if (!excelData) {
            showMessage('Nenhum arquivo carregado.', 'error');
            return;
        }
        
        showLoading(true);
        
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
            
            // Mostrar mensagem de sucesso
            showMessage(`Conversão concluída! ${result.lineCount} linhas processadas.`, 'success');
            
            // Rolar para o preview
            previewText.scrollIntoView({ behavior: 'smooth' });
            
            // Verificar comprimento da primeira linha
            const firstLine = result.content.split('\n')[0];
            if (firstLine && firstLine.length !== 101) {
                console.warn(`Atenção: A primeira linha tem ${firstLine.length} caracteres, esperado 101.`);
            }
            
        } catch (error) {
            console.error('Erro na conversão:', error);
            showMessage('Erro na conversão: ' + error.message, 'error');
        } finally {
            showLoading(false);
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
            showMessage('Arquivo baixado com sucesso!', 'success');
            
            setTimeout(() => {
                stats.textContent = originalText;
            }, 3000);
            
        } catch (error) {
            showMessage('Erro ao baixar arquivo: ' + error.message, 'error');
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
        fileInfo.textContent = 'Arraste e solte seu arquivo Excel aqui';
        
        // Remover mensagens
        showMessage('');
        
        // Resetar estado
        currentFile = null;
        excelData = null;
        converter.txtContent = '';
        converter.lineCount = 0;
        
        showMessage('Todos os dados foram limpos. Pronto para um novo arquivo.', 'success');
    });

    document.addEventListener('DOMContentLoaded', function() {
    // Elementos DOM
    const fileInput = document.getElementById('fileInput');
    const dropArea = document.getElementById('dropArea');
    const convertBtn = document.getElementById('convertBtn');
    const downloadBtn = document.getElementById('downloadBtn');
    const clearBtn = document.getElementById('clearBtn');
    const previewText = document.getElementById('previewText');
    const stats = document.getElementById('stats');
    const fileInfo = document.getElementById('fileInfo');
    
    // Instância do conversor
    const converter = new ExcelToTxtConverter();
    
    // Estado da aplicação
    let currentFile = null;
    let excelData = null;
    
    // 1. Evento de clique no botão "Selecione um arquivo" já está vinculado via label
    // 2. Evento de change no input file
    fileInput.addEventListener('change', function(e) {
        if (e.target.files.length > 0) {
            handleFile(e.target.files[0]);
        }
    });
    
    // 3. Drag and drop
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
    
    dropArea.addEventListener('drop', function(e) {
        const dt = e.dataTransfer;
        const files = dt.files;
        
        if (files.length > 0) {
            handleFile(files[0]);
        }
    });
    
    function handleFile(file) {
        currentFile = file;
        
        // Verificar se é um arquivo Excel
        if (!file.name.match(/\.(xlsx|xls|csv)$/i)) {
            alert('Por favor, selecione um arquivo Excel (.xlsx, .xls) ou CSV (.csv)');
            return;
        }
        
        // Atualizar informações do arquivo
        fileInfo.textContent = `Arquivo: ${file.name} (Carregando...)`;
        
        // Ler o arquivo
        const reader = new FileReader();
        
        reader.onload = function(e) {
            try {
                const data = e.target.result;
                const workbook = XLSX.read(data, { type: 'binary' });
                
                // Pegar a primeira planilha
                const firstSheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[firstSheetName];
                
                // Converter para JSON
                excelData = XLSX.utils.sheet_to_json(worksheet);
                
                console.log('Dados carregados:', excelData);
                
                // Habilitar botão de conversão
                convertBtn.disabled = false;
                
                // Atualizar informações do arquivo
                fileInfo.textContent = `Arquivo: ${file.name} (${excelData.length} linhas)`;
                
                // Resetar preview
                previewText.value = '';
                stats.textContent = '0 linhas processadas';
                downloadBtn.disabled = true;
                
                alert('Arquivo carregado com sucesso! Clique em "Converter para TXT" para gerar o arquivo.');
                
            } catch (error) {
                console.error('Erro ao processar arquivo:', error);
                alert('Erro ao processar o arquivo. Verifique se é um arquivo Excel válido.');
                fileInfo.textContent = 'Arraste e solte seu arquivo Excel aqui';
            }
        };
        
        reader.onerror = function() {
            alert('Erro ao ler o arquivo.');
            fileInfo.textContent = 'Arraste e solte seu arquivo Excel aqui';
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
        fileInfo.textContent = 'Arraste e solte seu arquivo Excel aqui';
        
        // Resetar estado
        currentFile = null;
        excelData = null;
        converter.txtContent = '';
        converter.lineCount = 0;
    });
});
    
    // Debug: Mostrar informações no console
    console.log('Aplicação configurada com sucesso');
    console.log('XLSX disponível:', typeof XLSX !== 'undefined');
});
