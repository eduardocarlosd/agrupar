const XLSX = require('xlsx');

// Função para ler o arquivo Excel
function lerArquivoExcel(caminhoArquivo) {
    const workbook = XLSX.readFile('C:/Users/Dudu/OneDrive/Área de Trabalho/agrupar/agrupar.xlsx');
    const sheetName = 'agrupar';
    const worksheet = workbook.Sheets[sheetName];
    return XLSX.utils.sheet_to_json(worksheet, { defval: '' });
}

// Função para agrupar as peças
function agruparPecas(pecas) {
    const agrupados = {};

    pecas.forEach(peca => {
        const key = `${peca.Posição}`;
        if (!agrupados[key]) {
            agrupados[key] = [];
        }
        for (let i = 0; i < peca.Quantidade; i++) {
            agrupados[key].push({
                Tipo: peca.Tipo,
                Espessura: peca.Espessura,
                Largura: peca.Largura,
                Alma: peca.Alma,
                MesaInferior: peca['Mesa Inferior'],
                MesaSuperior: peca['Mesa Superior'],
                Comprimento: peca.Comprimento
            });
        }
    });

    return agrupados;
}

// Função para gerar a saída final
function gerarRelacao(agrupados) {
    const resultado = [];
    const tiposOrdem = ['ALMA', 'MESA INFERIOR', 'MESA SUPERIOR'];

    for (const posicao in agrupados) {
        const grupo = agrupados[posicao];
        
        // Agrupar por combinação única de atributos
        const subGrupos = {};
        
        grupo.forEach(peca => {
            const key = `${peca.Tipo}-${peca.Espessura}-${peca.Largura}-${peca.Alma}-${peca.MesaInferior}-${peca.MesaSuperior}`;
            if (!subGrupos[key]) {
                subGrupos[key] = [];
            }
            subGrupos[key].push(peca);
        });

        // Processar cada subgrupo e ordenar pelo tipo
        for (const key in subGrupos) {
            const subGrupo = subGrupos[key];
            subGrupo.sort((a, b) => tiposOrdem.indexOf(a.Tipo) - tiposOrdem.indexOf(b.Tipo));

            let totalComprimento = 0;
            let grupoAtual = [];

            subGrupo.forEach(peca => {
                const novoComprimento = totalComprimento + peca.Comprimento;
                if (novoComprimento > 12000) {
                    resultado.push({
                        Posição: posicao,
                        Tipo: grupoAtual.map(p => p.Tipo).join(', '),
                        Espessura: peca.Espessura,
                        Alma: peca.Alma,
                        MesaInferior: peca.MesaInferior,
                        MesaSuperior: peca.MesaSuperior,
                        ComprimentoTotal: totalComprimento
                    });
                    grupoAtual = [];
                    totalComprimento = 0;
                }
                grupoAtual.push(peca);
                totalComprimento += peca.Comprimento;
            });

            if (grupoAtual.length > 0) {
                resultado.push({
                    Posição: posicao,
                    Tipo: grupoAtual.map(p => p.Tipo).join(', '),
                    Espessura: grupoAtual[0].Espessura,
                    Alma: grupoAtual[0].Alma,
                    MesaInferior: grupoAtual[0].MesaInferior,
                    MesaSuperior: grupoAtual[0].MesaSuperior,
                    ComprimentoTotal: totalComprimento
                });
            }
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
const caminhoArquivoOriginal = 'C:/Users/Dudu/OneDrive/Área de Trabalho/agrupar/agrupar.xlsx';
// Caminho do novo arquivo Excel
const caminhoArquivoNovo = 'C:/Users/Dudu/OneDrive/Área de Trabalho/agrupadas/resultado_agrupado.xlsx';

// Ler e processar o arquivo Excel
const pecas = lerArquivoExcel(caminhoArquivoOriginal);
const agrupados = agruparPecas(pecas);
const relacao = gerarRelacao(agrupados);

// Salvar a relação final em um novo arquivo Excel
salvarArquivoExcel(relacao, caminhoArquivoNovo);

console.log('Arquivo Excel salvo com sucesso!');
