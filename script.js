const abatimentos = [];
let parcelas = []; // Variável para armazenar as parcelas lidas do Excel

// Função para converter a data numérica do Excel para o formato legível
function converterDataExcelParaDataLegivel(numeroExcel) {
    const dataInicial = new Date(1900, 0, 1); // Excel considera que 1 de janeiro de 1900 é o dia 1
    const dataConvertida = new Date(dataInicial.getTime() + (numeroExcel - 1) * 24 * 60 * 60 * 1000);
    return dataConvertida.toLocaleDateString('pt-BR'); // Converte para formato DD/MM/YYYY
}

// Ler o arquivo Excel
document.getElementById('lerExcel').addEventListener('click', () => {
    const arquivo = document.getElementById('upload').files[0];
    if (!arquivo) {
        alert("Por favor, selecione um arquivo Excel.");
        return;
    }

    const leitor = new FileReader();
    leitor.onload = (e) => {
        const dados = new Uint8Array(e.target.result);
        const workbook = XLSX.read(dados, { type: "array" });
        const primeiraAba = workbook.Sheets[workbook.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(primeiraAba, { header: 1 });

        // Processar as linhas da planilha e armazenar as parcelas
        parcelas = json.slice(1).map(linha => ({
            vencimento: converterDataExcelParaDataLegivel(linha[0]), // Converte data de vencimento
            boleto: linha[1], // Número do boleto
            valor: parseFloat(linha[3]) // Valor da parcela
        }));

        // Ordenar as parcelas em ordem decrescente de vencimento
        parcelas.sort((a, b) => new Date(b.vencimento) - new Date(a.vencimento));

        alert('Planilha lida com sucesso!');
    };
    leitor.readAsArrayBuffer(arquivo);
});

// Adicionar abatimento
document.getElementById('adicionarAbatimento').addEventListener('click', () => {
    const data = document.getElementById('dataAbatimento').value;
    const valor = parseFloat(document.getElementById('valorAbatimento').value);

    if (data && !isNaN(valor)) {
        abatimentos.push({ data, valor });
        document.getElementById('dataAbatimento').value = '';
        document.getElementById('valorAbatimento').value = '';
        alert('Abatimento adicionado com sucesso!');
    } else {
        alert('Por favor, insira uma data e um valor válidos.');
    }
});

// Finalizar abatimentos
document.getElementById('finalizarAbatimentos').addEventListener('click', () => {
   
    
    if (abatimentos.length === 0) {
        alert('Nenhum abatimento adicionado.');
        return;
    }
    document.getElementById('finalizarAbatimentos').style.display = 'none';
    document.getElementById('adicionarAbatimento').style.display = 'none';

    const resultados = [];
    let valorRestanteAbatimento = 0;

    // Processar cada abatimento separadamente
    for (const abatimento of abatimentos) {
        valorRestanteAbatimento = abatimento.valor; // Reinicia para o valor do abatimento atual
        const parcelasAbatidas = [];

        // Processar as parcelas em ordem decrescente de vencimento
        for (let i = 0; i < parcelas.length; i++) {
            const parcela = parcelas[i];
            const valorParcela = parcela.valor;
           

            if (valorRestanteAbatimento >= valorParcela) {
                
                parcelasAbatidas.push({
                    ...parcela,
                    dataAbatimento: abatimento.data,
                    valorAbatido: valorParcela,
                    valorRestante: 0,
                    status: "Valor total da baixa"
                })
                   ;
                valorRestanteAbatimento -= valorParcela;
                parcelas.splice(i, 1); // Remove a parcela já completamente abatida
                i--; // Ajusta o índice após remover a parcela
            } else if (valorRestanteAbatimento > 0) {
                parcelasAbatidas.push({
                    ...parcela,
                    dataAbatimento: abatimento.data,
                    valorAbatido: valorRestanteAbatimento,
                    valorRestante: valorParcela - valorRestanteAbatimento,
                    status: `Abatido parcialmente: R$ ${valorRestanteAbatimento.toFixed(2)}`
                });
                parcela.valor -= valorRestanteAbatimento; // Atualiza o valor da parcela restante
                valorRestanteAbatimento = 0; // Termina o abatimento
                break; // Sai do loop quando não há mais abatimento restante
            }

            if (valorRestanteAbatimento <= 0) break;
        }

        resultados.push({ data: abatimento.data, parcelasAbatidas });
    }

    // Exibir resultados organizados por data de abatimento
    let detalhesHtml = '';
    for (const resultado of resultados) {
        detalhesHtml += `<h3>Abatimento realizado na data: ${resultado.data}</h3>`;
        for (const parcela of resultado.parcelasAbatidas) {
            detalhesHtml += `
                <p class="data">  Data de vencimento: ${parcela.vencimento}</p>
                <p>Número do boleto: ${parcela.boleto}</p>
                <p>Valor original da parcela: R$ ${parcela.valor.toFixed(2).replace('.', ',')}</p>
                <p>Valor abatido: R$ ${parcela.valorAbatido.toFixed(2).replace('.', ',')}</p>
                <p>Valor restante da parcela: R$ ${parcela.valorRestante.toFixed(2).replace('.', ',')}</p>
                <p id="statuss" >Status: ${parcela.status}</p>
                <hr>
            `;
        }
    }

    document.getElementById('detalhesAbatimentos').innerHTML = detalhesHtml;
});

