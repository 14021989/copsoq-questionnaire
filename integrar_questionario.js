// Script para integrar o questionário COPSOQ II com a planilha Excel
// Este script deve ser adicionado ao final do arquivo HTML do questionário

// Função para gerar ID único para cada resposta
function generateResponseId() {
    return 'RESP_' + Date.now() + '_' + Math.random().toString(36).substr(2, 9);
}

// Função para converter resposta em valor numérico
function getResponseValue(questionName, responseText) {
    // Mapeamento das respostas para valores numéricos
    const scaleMapping = {
        // Escala 1-5 (Nunca/quase nunca até Sempre)
        'Nunca/quase nunca': 1,
        'Raramente': 2,
        'Às vezes': 3,
        'Frequentemente': 4,
        'Sempre': 5,
        
        // Escala alternativa (Nada/quase nada até Extremamente)
        'Nada/quase nada': 1,
        'Um pouco': 2,
        'Moderadamente': 3,
        'Muito': 4,
        'Extremamente': 5,
        
        // Escala de saúde
        'Excelente': 5,
        'Muito Boa': 4,
        'Boa': 3,
        'Razoável': 2,
        'Deficitária': 1
    };
    
    return scaleMapping[responseText] || responseText;
}

// Função para coletar todas as respostas do formulário
function collectFormData() {
    const form = document.getElementById('copsoqForm');
    const formData = new FormData(form);
    
    const responseData = {
        id_resposta: generateResponseId(),
        data_hora_submissao: new Date().toISOString(),
        empresa: formData.get('empresa') || '',
        genero: formData.get('genero') || '',
        faixa_etaria: formData.get('idade') || '',
        tempo_na_empresa: formData.get('tempo_empresa') || '',
        respostas: {}
    };
    
    // Coletar respostas das perguntas Q1-Q41
    for (let i = 1; i <= 41; i++) {
        const questionName = `q${i}`;
        const responseValue = formData.get(questionName);
        
        if (responseValue) {
            // Encontrar o texto da resposta
            const radioButton = document.querySelector(`input[name="${questionName}"][value="${responseValue}"]`);
            const responseText = radioButton ? radioButton.nextElementSibling.textContent.trim() : responseValue;
            
            responseData.respostas[`Q${i.toString().padStart(2, '0')}`] = {
                resposta: responseText,
                valor: getResponseValue(questionName, responseText)
            };
        }
    }
    
    return responseData;
}

// Função para converter dados para formato CSV (compatível com Excel)
function convertToCSV(data) {
    const headers = [
        'ID_Resposta', 'Data_Hora_Submissao', 'Empresa', 'Genero', 'Faixa_Etaria', 'Tempo_na_Empresa'
    ];
    
    // Adicionar cabeçalhos das perguntas
    for (let i = 1; i <= 41; i++) {
        const qNum = i.toString().padStart(2, '0');
        headers.push(`Q${qNum}_Resposta`, `Q${qNum}_Valor`);
    }
    
    // Criar linha de dados
    const row = [
        data.id_resposta,
        data.data_hora_submissao,
        data.empresa,
        data.genero,
        data.faixa_etaria,
        data.tempo_na_empresa
    ];
    
    // Adicionar respostas das perguntas
    for (let i = 1; i <= 41; i++) {
        const qNum = i.toString().padStart(2, '0');
        const questionData = data.respostas[`Q${qNum}`];
        
        if (questionData) {
            row.push(questionData.resposta, questionData.valor);
        } else {
            row.push('', '');
        }
    }
    
    // Converter para CSV
    const csvContent = headers.join(',') + '\n' + 
                      row.map(field => `"${field}"`).join(',');
    
    return csvContent;
}

// Função para fazer download do arquivo CSV
function downloadCSV(csvContent, filename) {
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    
    if (link.download !== undefined) {
        const url = URL.createObjectURL(blob);
        link.setAttribute('href', url);
        link.setAttribute('download', filename);
        link.style.visibility = 'hidden';
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
    }
}

// Função para salvar dados localmente (localStorage)
function saveToLocalStorage(data) {
    const existingData = JSON.parse(localStorage.getItem('copsoq_responses') || '[]');
    existingData.push(data);
    localStorage.setItem('copsoq_responses', JSON.stringify(existingData));
}

// Função para exportar todos os dados salvos
function exportAllData() {
    const allData = JSON.parse(localStorage.getItem('copsoq_responses') || '[]');
    
    if (allData.length === 0) {
        alert('Nenhuma resposta encontrada para exportar.');
        return;
    }
    
    // Criar cabeçalhos CSV
    const headers = [
        'ID_Resposta', 'Data_Hora_Submissao', 'Empresa', 'Genero', 'Faixa_Etaria', 'Tempo_na_Empresa'
    ];
    
    for (let i = 1; i <= 41; i++) {
        const qNum = i.toString().padStart(2, '0');
        headers.push(`Q${qNum}_Resposta`, `Q${qNum}_Valor`);
    }
    
    // Criar linhas de dados
    const rows = [headers.join(',')];
    
    allData.forEach(data => {
        const row = [
            data.id_resposta,
            data.data_hora_submissao,
            data.empresa,
            data.genero,
            data.faixa_etaria,
            data.tempo_na_empresa
        ];
        
        for (let i = 1; i <= 41; i++) {
            const qNum = i.toString().padStart(2, '0');
            const questionData = data.respostas[`Q${qNum}`];
            
            if (questionData) {
                row.push(questionData.resposta, questionData.valor);
            } else {
                row.push('', '');
            }
        }
        
        rows.push(row.map(field => `"${field}"`).join(','));
    });
    
    const csvContent = rows.join('\n');
    const timestamp = new Date().toISOString().slice(0, 19).replace(/:/g, '-');
    downloadCSV(csvContent, `copsoq_respostas_${timestamp}.csv`);
}

// Modificar o evento de submissão do formulário
document.addEventListener('DOMContentLoaded', function() {
    const form = document.getElementById('copsoqForm');
    
    if (form) {
        form.addEventListener('submit', function(e) {
            e.preventDefault();
            
            // Validar se todas as perguntas foram respondidas
            let allAnswered = true;
            for (let i = 1; i <= 41; i++) {
                const questionName = `q${i}`;
                const answered = document.querySelector(`input[name="${questionName}"]:checked`);
                if (!answered) {
                    allAnswered = false;
                    break;
                }
            }
            
            // Validar campos demográficos
            const requiredFields = ['empresa', 'genero', 'idade', 'tempo_empresa'];
            for (const field of requiredFields) {
                const fieldValue = document.getElementById(field).value;
                if (!fieldValue) {
                    allAnswered = false;
                    break;
                }
            }
            
            if (!allAnswered) {
                alert('Por favor, responda todas as perguntas antes de enviar.');
                return;
            }
            
            // Coletar dados
            const responseData = collectFormData();
            
            // Salvar localmente
            saveToLocalStorage(responseData);
            
            // Converter para CSV e fazer download
            const csvContent = convertToCSV(responseData);
            const timestamp = new Date().toISOString().slice(0, 19).replace(/:/g, '-');
            downloadCSV(csvContent, `copsoq_resposta_${timestamp}.csv`);
            
            // Mostrar mensagem de sucesso
            alert('Questionário enviado com sucesso! O arquivo CSV foi baixado automaticamente.');
            
            // Opcional: limpar formulário
            form.reset();
        });
    }
    
    // Adicionar botão para exportar todos os dados
    const exportButton = document.createElement('button');
    exportButton.textContent = 'Exportar Todas as Respostas';
    exportButton.type = 'button';
    exportButton.className = 'bg-green-600 text-white px-4 py-2 rounded-md hover:bg-green-700 ml-4';
    exportButton.onclick = exportAllData;
    
    // Encontrar o botão de envio e adicionar o botão de exportação ao lado
    const submitButton = document.querySelector('button[type="submit"]');
    if (submitButton && submitButton.parentNode) {
        submitButton.parentNode.appendChild(exportButton);
    }
});

// Função para limpar dados salvos (opcional)
function clearSavedData() {
    if (confirm('Tem certeza que deseja limpar todas as respostas salvas?')) {
        localStorage.removeItem('copsoq_responses');
        alert('Dados limpos com sucesso!');
    }
}

