document.addEventListener('DOMContentLoaded', function() {
    // Elementos DOM
    const excelFileInput = document.getElementById('excelFile');
    const selectedFileName = document.getElementById('selectedFileName');
    const txtPreview = document.getElementById('txtPreview');
    const downloadBtn = document.getElementById('downloadBtn');
    const clearBtn = document.getElementById('clearBtn');
    const rowCount = document.getElementById('rowCount');
    const totalChars = document.getElementById('totalChars');
    const fileSize = document.getElementById('fileSize');
    const dropArea = document.getElementById('dropArea');
    const previewSection = document.getElementById('previewSection');
    const previewStats = document.getElementById('previewStats');
    
    // Estado da aplicação
    let txtContent = '';
    let currentFileName = '';
    
    // Configurar arrastar e soltar
    function setupDragAndDrop() {
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
            dropArea.style.borderColor = '#1f618d';
            dropArea.style.backgroundColor = '#e0f0ff';
        }
        
        function unhighlight() {
            dropArea.style.borderColor = '#3498db';
            dropArea.style.backgroundColor = '#ecf0f1';
        }
        
        dropArea.addEventListener('drop', handleDrop, false);
        
        function handleDrop(e) {
            const dt = e.dataTransfer;
            const files = dt.files;
            
            if (files.length) {
                handleFile(files[0]);
            }
        }
    }
    
    // Lidar com seleção de arquivo
    excelFileInput.addEventListener('change', function(e) {
        if (e.target.files.length) {
            handleFile(e.target.files[0]);
        }
    });
    
    // Função principal para processar arquivo
    function handleFile(file) {
        if (!file) return;
        
        currentFileName = file.name.replace(/\.[^/.]+$/, "");
        selectedFileName.innerHTML = `
            <strong>Arquivo:</strong> ${file.name}<br>
            <strong>Tamanho:</strong> ${formatFileSize(file.size)}<br>
            <strong>Tipo:</strong> ${file.type || getFileExtension(file.name)}
        `;
        
        const reader = new FileReader();
        
        reader.onload = function(e) {
            const data = e.target.result;
            processExcel(data);
        };
        
        if (file.name.endsWith('.csv')) {
            reader.readAsText(file, 'UTF-8');
        } else {
            reader.readAsArrayBuffer(file);
        }
        
        // Mostrar seção de preview
        previewSection.style.display = 'block';
        previewStats.style.display = 'flex';
    }
    
    // Processar Excel/CSV
    function processExcel(data) {
        try {
            let workbook;
            let excelData;
            
            if (typeof data === 'string') {
                // Processar CSV
                const lines = data.split('\n');
                const result = [];
                
                lines.forEach(line => {
                    if (line.trim()) {
                        const row = line.split(',').map(cell => cell.trim());
                        result.push(row);
                    }
                });
                
                excelData = result;
            } else {
                // Processar Excel
                workbook = XLSX.read(data, { type: 'array' });
                const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                excelData = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });
            }
            
            convertToSpecificTXT(excelData);
        } catch (error) {
            showError('Erro ao processar arquivo: ' + error.message);
            console.error('Erro:', error);
        }
    }
    
    // Converter para formato TXT específico
    function convertToSpecificTXT(data) {
        if (!data || data.length === 0) {
            showError('Arquivo vazio ou inválido');
            return;
        }
        
        // Detectar cabeçalho
        let startRow = 0;
        const headers = data[0];
        const hasHeader = headers && 
            (headers.some(h => h && h.toString().toLowerCase().includes('categoria')) ||
             headers.some(h => h && h.toString().toLowerCase().includes('codigo')));
        
        if (hasHeader) {
            startRow = 1;
        }
        
        const lines = [];
        
        // Processar cada linha
        for (let i = startRow; i < data.length; i++) {
            const row = data[i];
            if (!row || row.length === 0) continue;
            
            // Extrair dados (considerando diferentes estruturas possíveis)
            const categoria = (row[0] || 'AGUA').toString().padEnd(10, ' ');
            const codigoItem = (row[1] || '100000').toString().padStart(6, '0').substring(0, 6);
            const parteCodigo = row[2] || '1';
            const dataVal = row[3] || new Date().toISOString();
            const periodo = row[4] || '012025';
            
            // Valores (colunas 5-8)
            const valores = [];
            for (let j = 5; j <= 8; j++) {
                const val = row[j] || '0';
                valores.push(formatTo12Digits(val));
            }
            
            // Formatar linha
            const torCode = `TOR${parseInt(parteCodigo).toString().padStart(7, '0')}`;
            const dataFmt = formatDate(dataVal);
            const periodoFmt = periodo.toString().padStart(6, '0');
            
            const line = `${categoria}${codigoItem}${torCode}${dataFmt}${periodoFmt}${valores.join('')}`;
            lines.push(line);
        }
        
        // Gerar conteúdo final
        txtContent = lines.join('\n');
        
        // Atualizar preview
        updatePreview(lines);
        
        // Atualizar estatísticas
        updateStats(lines.length, txtContent.length);
        
        // Habilitar download
        downloadBtn.disabled = false;
        downloadBtn.innerHTML = '<i class="fas fa-download"></i> Baixar Arquivo TXT';
        
        showSuccess('Arquivo convertido com sucesso!');
    }
    
    // Atualizar pré-visualização
    function updatePreview(lines) {
        if (lines.length === 0) {
            txtPreview.innerHTML = '<p class="empty-preview">Nenhum dado encontrado no arquivo.</p>';
            return;
        }
        
        const previewLines = lines.slice(0, 5);
        const previewContent = previewLines.join('\n');
        
        if (lines.length > 5) {
            txtPreview.textContent = previewContent + `\n\n... (${lines.length - 5} linhas adicionais)`;
        } else {
            txtPreview.textContent = previewContent;
        }
    }
    
    // Atualizar estatísticas
    function updateStats(rows, chars) {
        rowCount.textContent = rows.toLocaleString('pt-BR');
        totalChars.textContent = chars.toLocaleString('pt-BR');
        
        const size = new Blob([txtContent]).size;
        fileSize.textContent = formatFileSize(size);
    }
    
    // Formatar número para 12 dígitos
    function formatTo12Digits(value) {
        try {
            // Converter para número
            let num = parseFloat(value);
            if (isNaN(num)) num = 0;
            
            // Multiplicar por 100 para converter para centavos (se for decimal)
            const isDecimal = value.toString().includes('.');
            if (isDecimal) {
                num = Math.round(num * 100);
            } else {
                num = Math.round(num);
            }
            
            // Formatar com 12 dígitos
            return num.toString().padStart(12, '0').substring(0, 12);
        } catch (e) {
            return '000000000000';
        }
    }
    
    // Formatar data para DD/MM/YYYY
    function formatDate(dateStr) {
        try {
            let date;
            
            if (typeof dateStr === 'string') {
                // Tentar diferentes formatos
                if (dateStr.includes('-')) {
                    const parts = dateStr.split('-');
                    if (parts.length === 3) {
                        date = new Date(parts[0], parts[1] - 1, parts[2]);
                    }
                } else if (dateStr.includes('/')) {
                    const parts = dateStr.split('/');
                    if (parts.length === 3) {
                        date = new Date(parts[2], parts[1] - 1, parts[0]);
                    }
                }
            }
            
            if (!date || isNaN(date.getTime())) {
                date = new Date(dateStr);
            }
            
            if (!date || isNaN(date.getTime())) {
                return '01/01/2025';
            }
            
            const day = date.getDate().toString().padStart(2, '0');
            const month = (date.getMonth() + 1).toString().padStart(2, '0');
            const year = date.getFullYear();
            
            return `${day}/${month}/${year}`;
        } catch (e) {
            return '01/01/2025';
        }
    }
    
    // Download do arquivo TXT
    downloadBtn.addEventListener('click', function() {
        if (!txtContent) {
            showError('Nenhum conteúdo para baixar');
            return;
        }
        
        try {
            const blob = new Blob([txtContent], { type: 'text/plain;charset=utf-8' });
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = `${currentFileName}_formatado.txt`;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(url);
            
            // Feedback visual
            downloadBtn.innerHTML = '<i class="fas fa-check"></i> Baixado!';
            downloadBtn.style.background = 'linear-gradient(135deg, #27ae60 0%, #229954 100%)';
            
            setTimeout(() => {
                downloadBtn.innerHTML = '<i class="fas fa-download"></i> Baixar Arquivo TXT';
                downloadBtn.style.background = 'linear-gradient(135deg, #27ae60 0%, #229954 100%)';
            }, 2000);
            
            showSuccess('Arquivo baixado com sucesso!');
        } catch (error) {
            showError('Erro ao baixar arquivo: ' + error.message);
        }
    });
    
    // Limpar dados
    clearBtn.addEventListener('click', function() {
        excelFileInput.value = '';
        txtContent = '';
        currentFileName = '';
        selectedFileName.textContent = '';
        txtPreview.innerHTML = '<p class="empty-preview">Nenhum arquivo processado ainda. Carregue uma planilha para ver a pré-visualização.</p>';
        downloadBtn.disabled = true;
        rowCount.textContent = '0';
        totalChars.textContent = '0';
        fileSize.textContent = '0 KB';
        
        showInfo('Dados limpos com sucesso');
    });
    
    // Utilitários
    function formatFileSize(bytes) {
        if (bytes === 0) return '0 Bytes';
        const k = 1024;
        const sizes = ['Bytes', 'KB', 'MB', 'GB'];
        const i = Math.floor(Math.log(bytes) / Math.log(k));
        return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
    }
    
    function getFileExtension(filename) {
        return filename.slice((filename.lastIndexOf(".") - 1 >>> 0) + 2);
    }
    
    function showError(message) {
        showNotification(message, 'error');
    }
    
    function showSuccess(message) {
        showNotification(message, 'success');
    }
    
    function showInfo(message) {
        showNotification(message, 'info');
    }
    
    function showNotification(message, type) {
        // Criar elemento de notificação
        const notification = document.createElement('div');
        notification.className = `notification ${type}`;
        notification.innerHTML = `
            <span class="notification-icon">
                ${type === 'error' ? '❌' : type === 'success' ? '✅' : 'ℹ️'}
            </span>
            <span class="notification-text">${message}</span>
        `;
        
        // Estilos da notificação
        notification.style.cssText = `
            position: fixed;
            top: 20px;
            right: 20px;
            padding: 15px 20px;
            background: ${type === 'error' ? '#e74c3c' : type === 'success' ? '#27ae60' : '#3498db'};
            color: white;
            border-radius: 5px;
            display: flex;
            align-items: center;
            gap: 10px;
            box-shadow: 0 4px 6px rgba(0,0,0,0.1);
            z-index: 1000;
            animation: slideIn 0.3s ease;
        `;
        
        // Adicionar ao DOM
        document.body.appendChild(notification);
        
        // Remover após 5 segundos
        setTimeout(() => {
            notification.style.animation = 'slideOut 0.3s ease';
            setTimeout(() => {
                if (notification.parentNode) {
                    notification.parentNode.removeChild(notification);
                }
            }, 300);
        }, 5000);
        
        // Adicionar animações CSS
        const style = document.createElement('style');
        style.textContent = `
            @keyframes slideIn {
                from { transform: translateX(100%); opacity: 0; }
                to { transform: translateX(0); opacity: 1; }
            }
            @keyframes slideOut {
                from { transform: translateX(0); opacity: 1; }
                to { transform: translateX(100%); opacity: 0; }
            }
        `;
        document.head.appendChild(style);
    }
    
    // Inicializar
    setupDragAndDrop();
    showInfo('Aplicação carregada com sucesso! Pronto para converter arquivos.');
});
