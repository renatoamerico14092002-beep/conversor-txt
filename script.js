// ========== CONVERSOR ==========
class ExcelToTxtConverter {
    constructor() {
        this.txtContent = '';
        this.lineCount = 0;
    }

    formatField(value, length, align = 'left', padChar = ' ') {
        let str = String(value || '');
        
        // Remove acentos mas mantém letras
        str = str.normalize('NFD').replace(/[\u0300-\u036f]/g, '');
        
        if (align === 'left') {
            return str.padEnd(length, padChar);
        } else {
            return str.padStart(length, padChar);
        }
    }

    formatNumber(value, length) {
        let num = parseFloat(value || 0);
        let intValue = Math.round(num * 100);
        intValue = Math.abs(intValue);
        return intValue.toString().padStart(length, '0');
    }

   formatDate(dateString) {
    try {
        if (!dateString) return '00/00/0000';
        
        // Remove hora se existir
        const datePart = dateString.toString().split(' ')[0];
        
        // Verifica formato DD/MM/YYYY
        const parts = datePart.split('/');
        if (parts.length === 3) {
            const day = parts[0].padStart(2, '0');
            const month = parts[1].padStart(2, '0');
            const year = parts[2];
            
            // Valida se é uma data real
            const testDate = new Date(`${year}-${month}-${day}`);
            if (isNaN(testDate.getTime())) {
                return '00/00/0000';
            }
            
            return `${day}/${month}/${year}`;
        }
        
        return '00/00/0000';
    } catch (error) {
        return '00/00/0000';
    }
}

    formatParteCodigo(parteCodigo) {
        let str = String(parteCodigo || '');
        
        if (/^[A-Za-z]/.test(str)) {
            const match = str.match(/^([A-Za-z])(\d*)$/);
            if (match) {
                const letra = match[1];
                const numeros = match[2] || '0';
                return letra + numeros.padStart(6, '0');
            }
        }
        
        const num = parseInt(str) || 0;
        return num.toString().padStart(7, '0');
    }

    formatPeriodo(periodo) {
        let str = String(periodo || '');
        return str.padStart(6, '0');
    }

    convertRow(row) {
        // Usa nomes de colunas mais comuns
        const categoria = this.formatField(row.categoria || row.CATEGORIA || 'AGUA', 10);
        const codigoItem = this.formatField(row.codigo_item || row['codigo item'] || row.CODIGO_ITEM || '', 9);
        const parteCodigo = this.formatParteCodigo(row.parte_codigo || row.PARTE_CODIGO || '');
        const data = this.formatDate(row.data || row.DATA || '');
        const periodo = this.formatPeriodo(row.periodo || row.PERIODO || '');
        
        const valor1 = this.formatNumber(row.valor1 || row.VALOR1 || 0, 14);
        const valor2 = this.formatNumber(row.valor2 || row.VALOR2 || 0, 15);
        const valor3 = this.formatNumber(row.valor3 || row.VALOR3 || 0, 14);
        const valor4 = this.formatNumber(row.valor4 || row.VALOR4 || 0, 15);
        
        return `${categoria}${codigoItem}${parteCodigo}${data}${periodo}${valor1}${valor2}${valor3}${valor4}`;
    }

    convertExcelData(data) {
        this.txtContent = '';
        this.lineCount = 0;
        
        if (!data || !Array.isArray(data)) {
            throw new Error('Dados inválidos');
        }

        data.forEach(row => {
            try {
                const txtLine = this.convertRow(row);
                this.txtContent += txtLine + '\n';
                this.lineCount++;
            } catch (error) {
                console.error('Erro na linha:', row, error);
            }
        });

        return {
            content: this.txtContent,
            lineCount: this.lineCount
        };
    }

    downloadTxt(filename = 'convertido.txt') {
        if (!this.txtContent) {
            throw new Error('Nada para baixar');
        }

        const blob = new Blob([this.txtContent], { type: 'text/plain;charset=utf-8' });
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        
        a.href = url;
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        window.URL.revokeObjectURL(url);
    }
}

// ========== APLICAÇÃO WEB ==========
document.addEventListener('DOMContentLoaded', function() {
    console.log('Aplicação iniciada');
    
    // Elementos
    const fileInput = document.getElementById('fileInput');
    const dropArea = document.getElementById('dropArea');
    const convertBtn = document.getElementById('convertBtn');
    const downloadBtn = document.getElementById('downloadBtn');
    const clearBtn = document.getElementById('clearBtn');
    const previewText = document.getElementById('previewText');
    const stats = document.getElementById('stats');
    const fileInfo = document.getElementById('fileInfo');
    
    // Conversor
    const converter = new ExcelToTxtConverter();
    let excelData = null;
    let currentFile = null;
    
    // ===== SELEÇÃO DE ARQUIVO SIMPLIFICADA =====
    
    // Clique no label
    document.querySelector('.btn-primary').addEventListener('click', function(e) {
        e.preventDefault();
        fileInput.click();
    });
    
    // Clique na área de drop
    dropArea.addEventListener('click', function(e) {
        if (e.target === dropArea || e.target === dropArea.querySelector('.upload-icon') || 
            e.target === dropArea.querySelector('h3') || e.target === dropArea.querySelector('p')) {
            fileInput.click();
        }
    });
    
    // Change no input
    fileInput.addEventListener('change', function(e) {
        if (e.target.files.length > 0) {
            handleFile(e.target.files[0]);
        }
    });
    
    // Drag and drop
    ['dragenter', 'dragover'].forEach(eventName => {
        dropArea.addEventListener(eventName, function(e) {
            e.preventDefault();
            dropArea.classList.add('dragover');
        });
    });
    
    ['dragleave', 'drop'].forEach(eventName => {
        dropArea.addEventListener(eventName, function(e) {
            e.preventDefault();
            dropArea.classList.remove('dragover');
            
            if (eventName === 'drop') {
                const files = e.dataTransfer.files;
                if (files.length > 0) {
                    handleFile(files[0]);
                }
            }
        });
    });
    
    // ===== PROCESSAMENTO DO ARQUIVO =====
    
    function handleFile(file) {
        console.log('Arquivo selecionado:', file.name);
        currentFile = file;
        
        // Validação básica
        if (!file.name.match(/\.(xlsx|xls|csv)$/i)) {
            alert('Selecione um arquivo Excel (.xlsx, .xls) ou CSV');
            return;
        }
        
        fileInfo.textContent = `Processando: ${file.name}`;
        
        const reader = new FileReader();
        
        reader.onload = function(e) {
            try {
                let data = e.target.result;
                let workbook;
                
                // Para CSV
                if (file.name.toLowerCase().endsWith('.csv')) {
                    workbook = XLSX.read(data, { type: 'binary', raw: true });
                } 
                // Para Excel
                else {
                    if (typeof data === 'string') {
                        workbook = XLSX.read(data, { type: 'binary' });
                    } else {
                        // ArrayBuffer
                        workbook = XLSX.read(new Uint8Array(data), { type: 'array' });
                    }
                }
                
                // Primeira planilha
                const firstSheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[firstSheetName];
                
                // Converter para JSON
                excelData = XLSX.utils.sheet_to_json(worksheet, { defval: '' });
                
                console.log('Dados extraídos:', excelData.length, 'linhas');
                
                if (excelData.length > 0) {
                    console.log('Colunas:', Object.keys(excelData[0]));
                }
                
                // Habilitar conversão
                convertBtn.disabled = false;
                fileInfo.textContent = `${file.name} (${excelData.length} linhas)`;
                stats.textContent = 'Pronto para converter';
                
            } catch (error) {
                console.error('Erro:', error);
                alert('Erro ao ler arquivo: ' + error.message);
                fileInfo.textContent = 'Erro ao processar arquivo';
            }
        };
        
        reader.onerror = function() {
            alert('Erro ao ler o arquivo');
            fileInfo.textContent = 'Erro de leitura';
        };
        
        // Ler como array buffer (mais compatível)
        reader.readAsArrayBuffer(file);
    }
    
    // ===== CONVERSÃO =====
    
    convertBtn.addEventListener('click', function() {
        if (!excelData) {
            alert('Nenhum arquivo carregado');
            return;
        }
        
        try {
            const result = converter.convertExcelData(excelData);
            
            // Mostrar preview
            const lines = result.content.split('\n');
            const preview = lines.slice(0, 15).join('\n');
            previewText.value = preview;
            
            if (lines.length > 15) {
                previewText.value += `\n... (mais ${lines.length - 15} linhas)`;
            }
            
            // Atualizar status
            stats.textContent = `${result.lineCount} linhas convertidas`;
            downloadBtn.disabled = false;
            
        } catch (error) {
            alert('Erro na conversão: ' + error.message);
        }
    });
    
    // ===== DOWNLOAD =====
    
    downloadBtn.addEventListener('click', function() {
        try {
            const filename = currentFile 
                ? currentFile.name.replace(/\.[^/.]+$/, "") + '_convertido.txt'
                : 'convertido.txt';
            
            converter.downloadTxt(filename);
            stats.textContent = '✅ Download concluído';
            
            setTimeout(() => {
                stats.textContent = `${converter.lineCount} linhas`;
            }, 2000);
            
        } catch (error) {
            alert('Erro ao baixar: ' + error.message);
        }
    });
    
    // ===== LIMPAR =====
    
    clearBtn.addEventListener('click', function() {
        fileInput.value = '';
        previewText.value = '';
        stats.textContent = '0 linhas';
        fileInfo.textContent = 'Arraste e solte seu arquivo Excel aqui';
        
        convertBtn.disabled = true;
        downloadBtn.disabled = true;
        
        excelData = null;
        currentFile = null;
        converter.txtContent = '';
        converter.lineCount = 0;
    });
    
    // ===== DICA INICIAL =====
    console.log('Aplicação pronta. Clique em "Selecione um arquivo" ou arraste um arquivo Excel.');
});
