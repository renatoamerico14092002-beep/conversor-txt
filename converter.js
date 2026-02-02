class ExcelToTxtConverter {
    constructor() {
        this.txtContent = '';
        this.lineCount = 0;
    }

    formatField(value, length, align = 'left', padChar = ' ') {
        let str = String(value || '');
        
        // Remove acentos para manter consistência, mas preserva letras
        str = str.normalize('NFD').replace(/[\u0300-\u036f]/g, '');
        
        if (align === 'left') {
            return str.padEnd(length, padChar);
        } else {
            return str.padStart(length, padChar);
        }
    }

    formatNumber(value, length) {
        // Converte o valor para número
        let num = parseFloat(value || 0);
        
        // Multiplica por 100 para converter para centavos (2 casas decimais)
        let intValue = Math.round(num * 100);
        
        // Garante que seja positivo
        intValue = Math.abs(intValue);
        
        // Formata com zeros à esquerda
        return intValue.toString().padStart(length, '0');
    }

    formatDate(dateString) {
        try {
            // Tenta converter a data
            let date;
            
            if (dateString instanceof Date) {
                date = dateString;
            } else if (typeof dateString === 'string') {
                // Remove qualquer hora que possa estar presente
                const datePart = dateString.split(' ')[0];
               
            }
            
            // Verifica se a data é válida
            if (isNaN(date.getTime())) {
                return '00/00/0000';
            }
            
            const day = date.getDate().toString().padStart(2, '0');
            const month = (date.getMonth() + 1).toString().padStart(2, '0');
            const year = date.getFullYear();
            return `${day}/${month}/${year}`;
        } catch (error) {
            console.error('Erro ao formatar data:', error, dateString);
            return '00/00/0000';
        }
    }

    formatParteCodigo(parteCodigo) {
        // Converte para string
        let str = String(parteCodigo || '');
        
        // Se for um número, formata com zeros à esquerda até 6 dígitos
        // Se começar com letra, mantém a letra e preenche com zeros
        if (/^[A-Za-z]/.test(str)) {
            // Extrai a letra e os números
            const match = str.match(/^([A-Za-z])(\d*)$/);
            if (match) {
                const letra = match[1];
                const numeros = match[2] || '0';
                return letra + numeros.padStart(6, '0');
            }
        }
        
        // Se for apenas números, preenche com zeros à esquerda até 7 dígitos
        const num = parseInt(str) || 0;
        return num.toString().padStart(7, '0');
    }

    formatPeriodo(periodo) {
        // Converte para string e preenche com zeros à esquerda até 6 dígitos
        let str = String(periodo || '');
        return str.padStart(6, '0');
    }

    convertRow(row) {
        // Formata cada campo conforme o formato desejado
        const categoria = this.formatField(row.categoria || 'AGUA', 10);
        const codigoItem = this.formatField(row.codigo_item || '', 9);
        const parteCodigo = this.formatParteCodigo(row.parte_codigo || row.parteCodigo || '');
        const data = this.formatDate(row.data || row.date || '');
        const periodo = this.formatPeriodo(row.periodo || '');
        
        // Formata os valores numéricos com tamanhos específicos
        // Baseado no formato que você mostrou:
        // Valor1: 14 caracteres
        // Valor2: 15 caracteres  
        // Valor3: 14 caracteres
        // Valor4: 15 caracteres
        const valor1 = this.formatNumber(row.valor1 || row.valor_1 || 0, 14);
        const valor2 = this.formatNumber(row.valor2 || row.valor_2 || 0, 15);
        const valor3 = this.formatNumber(row.valor3 || row.valor_3 || 0, 14);
        const valor4 = this.formatNumber(row.valor4 || row.valor_4 || 0, 15);
        
        // Monta a linha no formato especificado
        return `${categoria}${codigoItem}${parteCodigo}${data}${periodo}${valor1}${valor2}${valor3}${valor4}`;
    }

    convertExcelData(data) {
        this.txtContent = '';
        this.lineCount = 0;
        
        if (!data || !Array.isArray(data) || data.length === 0) {
            throw new Error('Dados inválidos ou vazios');
        }

        // Processa cada linha
        data.forEach(row => {
            try {
                const txtLine = this.convertRow(row);
                
                // Verifica se a linha tem exatamente 101 caracteres
                if (txtLine.length !== 101) {
                    console.warn(`Linha tem ${txtLine.length} caracteres, esperado 101:`, txtLine);
                }
                
                this.txtContent += txtLine + '\n';
                this.lineCount++;
            } catch (error) {
                console.error('Erro ao converter linha:', row, error);
                throw error;
            }
        });

        return {
            content: this.txtContent,
            lineCount: this.lineCount
        };
    }

    downloadTxt(filename = 'converted_file.txt') {
        if (!this.txtContent) {
            throw new Error('Nenhum conteúdo para baixar');
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
