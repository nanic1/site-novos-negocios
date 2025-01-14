(async function () {
    const fileUrl = 'BaseDados.xlsx'; // Caminho do arquivo na mesma pasta do HTML

    // Função para buscar o arquivo Excel e processá-lo
    async function fetchAndProcessExcel(url) {
        const response = await fetch(url);
        if (!response.ok) {
            throw new Error(`Erro ao carregar o arquivo: ${response.statusText}`);
        }

        const arrayBuffer = await response.arrayBuffer();
        const workbook = XLSX.read(arrayBuffer, { type: 'array' });

        // Seleciona a primeira planilha
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];

        // Lê os valores das células específicas
        const enderecoExterno = 'B2';
        const enderecoInterno = 'B3';

        const enderecoPriscylla = 'E2'
        const enderecoMauro = 'E3'
        const enderecoCedente = 'E4'
        const enderecoAlex = 'E5'

        const valorExterno = worksheet[enderecoExterno] ? worksheet[enderecoExterno].v : null;
        const valorInterno = worksheet[enderecoInterno] ? worksheet[enderecoInterno].v : null;
        const valorPriscylla = worksheet[enderecoPriscylla] ? worksheet[enderecoPriscylla].v : null;
        const valorMauro = worksheet[enderecoMauro] ? worksheet[enderecoMauro].v : null;
        const valorCedente = worksheet[enderecoCedente] ? worksheet[enderecoCedente].v : null;
        const valorAlex = worksheet[enderecoAlex] ? worksheet[enderecoAlex].v : null;

        return { valorExterno, valorInterno, valorPriscylla, valorMauro, valorCedente, valorAlex };
    }

    try {
        // Processa o arquivo Excel
        const { valorExterno, valorInterno, valorPriscylla, valorMauro, valorCedente, valorAlex } = await fetchAndProcessExcel(fileUrl);

        // Exibe os valores no console (opcional)
        console.log('Valor Externo:', valorExterno);
        console.log('Valor Interno:', valorInterno);

        // Atualiza o gráfico com os valores do Excel
        const chart1 = document.getElementById('grafico1');
        const chart2 = document.getElementById('grafico2');
        const chart3 = document.getElementById('grafico3');
        const chart4 = document.getElementById('grafico4')

        new Chart(chart1, {
            type: 'pie',
            data: {
                labels: ['Externo', 'Interno'],
                datasets: [
                    {
                        label: 'Status',
                        data: [valorExterno, valorInterno], // Usa os valores extraídos
                        borderWidth: 1,
                        hoverOffset: 10,
                        backgroundColor: ['#FF6384', '#36A2EB'], // Cores personalizadas
                    }
                ]
            },
            options: {
                responsive: true,
                plugins: {
                    legend: {
                        position: 'bottom',
                        labels: {
                            boxWidth: 20,
                            boxHeight: 20,
                            borderRadius: 5,
                        }
                    },
                }
            }
        });

        new Chart(chart2, {
            type: 'pie',
            data: {
                labels: ['PRISCYLLA', 'MAURO', 'ALEX', 'VISÃO CEDENTE'],
                datasets: [
                    {
                        label: 'Status',
                        data: [valorPriscylla, valorMauro, valorAlex, valorCedente],
                        borderWidth: 1,
                        hoverOffset: 10
                    }
                ]
            },
            options: {
                resposive: true,
                plugins: {
                    legend: {
                        position: 'bottom',
                        labels: {
                            boxWidth: 20,
                            boxHeight: 20,
                            borderRadius: 5,
                        }
                    },
                }
            }
        })

        new Chart(chart3, {
            type: 'pie',
            data: {
                labels: ['CADASTRO', 'CALL', 'COMITE', 'CONTRATO', 'PÓS CALL', 'PRÉ-ANALISE'],
                datasets: [
                    {
                        label: 'Status',
                        data: [1, 8, 3, 1, 8, 4],
                        borderWidth: 1,
                        hoverOffset: 10
                    }
                ]
            },
            options: {
                resposive: true,
                plugins: {
                    legend: {
                        position: 'bottom',
                        labels: {
                            boxWidth: 20,
                            boxHeight: 20,
                            borderRadius: 5,
                        }
                    },
                }
            }
        })

        new Chart(chart4, {
            type: 'pie',
            data: {
                labels: ['Primeira operação', 'Formalização do contrato', 'Comitê', 'Visita', 'Cadastro', 'Contato', 'Envio de documentação', 'Pós-call', 'Envio de proposta'],
                datasets: [
                    {
                        label: 'Status',
                        data: [27, 15, 7, 4, 3, 3, 3, 3, 3,],
                        borderWidth: 1,
                        hoverOffset: 10
                    }
                ]
            },
            options: {
                resposive: true,
                plugins: {
                    legend: {
                        position: 'bottom',
                        labels: {
                            boxWidth: 20,
                            boxHeight: 20,
                            borderRadius: 5,
                            padding: 4,
                            font: {
                                size: 12
                            }
                        }
                    },
                }
            }
        })
    } catch (error) {
        console.error('Erro:', error);
        document.getElementById('output').textContent = `Erro ao processar o arquivo: ${error.message}`;
    }
})();
