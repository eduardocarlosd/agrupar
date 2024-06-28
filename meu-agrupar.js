const XLSX = require('xlsx');

// Função para ler o arquivo Excel
function lerArquivoExcel(caminhoArquivo) {
    const workbook = XLSX.readFile('C:/Users/carlos.silva/Desktop/agrupar/agrupar.xlsx');
    const sheetName = 'agrupar'; // Especificar a planilha "agrupar"
    const worksheet = workbook.Sheets[sheetName];
    return XLSX.utils.sheet_to_json(worksheet, { defval: '' });
}

// Função para agrupar as peças
function agruparPecas(pecas) {
    const agrupados = {};

    pecas.forEach(peca => {
        const key = `${peca.Tipo}-${peca.Espessura}-${peca.Largura}`;
        if (!agrupados[key]) {
            agrupados[key] = [];
        }
        for (let i = 0; i < peca.Quantidade; i++) {
            agrupados[key].push({
                Posição: peca.Posição,
                Tipo: peca.Tipo,
                Espessura: peca.Espessura,
                Largura: peca.Largura,
                Comprimento: peca.Comprimento
            });
        }
    });

    return agrupados;
}

// Função para gerar a saída final
function gerarRelacao(agrupados) {
    const resultado = [];

    for (const key in agrupados) {
        const grupo = agrupados[key];
        grupo.sort((a, b) => a.Comprimento - b.Comprimento); // Ordenar por comprimento

        let totalComprimento = 0;
        let grupoAtual = [];
        let posicoes = [];

        grupo.forEach(peca => {
            if (totalComprimento + peca.Comprimento <= 12000) {
                grupoAtual.push(peca);
                posicoes.push(peca.Posição);
                totalComprimento += peca.Comprimento;
            } else {
                resultado.push({
                    Posição: posicoes.join(', '),
                    Tipo: grupoAtual[0].Tipo,
                    Espessura: grupoAtual[0].Espessura,
                    Largura: grupoAtual[0].Largura,
                    ComprimentoTotal: totalComprimento
                });
                grupoAtual = [peca];
                posicoes = [peca.Posição];
                totalComprimento = peca.Comprimento;
            }
        });

        if (grupoAtual.length > 0) {
            resultado.push({
                Posição: posicoes.join(', '),
                Tipo: grupoAtual[0].Tipo,
                Espessura: grupoAtual[0].Espessura,
                Largura: grupoAtual[0].Largura,
                ComprimentoTotal: totalComprimento
            });
        }
    }

    return resultado;
}

// Função para salvar os dados em um novo arquivo Excel
function salvarArquivoExcel(dados, caminhoArquivo) {
    const worksheet = XLSX.utils.json_to_sheet(dados);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'agrupados');
    XLSX.writeFile(workbook, caminhoArquivo);
}

// Caminho do arquivo Excel original
const caminhoArquivoOriginal = 'C:/Users/carlos.silva/Desktop/agrupar/agrupar.xlsx';
// Caminho do novo arquivo Excel
const caminhoArquivoNovo = 'C:/Users/carlos.silva/Desktop/agrupado/resultado_agrupado.xlsx';

// Ler e processar o arquivo Excel
const pecas = lerArquivoExcel(caminhoArquivoOriginal);
const agrupados = agruparPecas(pecas);
const relacao = gerarRelacao(agrupados);

// Salvar a relação final em um novo arquivo Excel
salvarArquivoExcel(relacao, caminhoArquivoNovo);

console.log('Arquivo Excel salvo com sucesso!');
