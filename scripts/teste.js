(async function () {
    const fileUrl = 'RELATORIO.xlsx'; // Caminho do arquivo na mesma pasta do HTML

    // Função para buscar o arquivo Excel e processá-lo
    async function fetchAndProcessExcel(url) {
        const response = await fetch(url);
        if (!response.ok) {
            throw new Error(`Erro ao carregar o arquivo: ${response.statusText}`);
        }

        const arrayBuffer = await response.arrayBuffer();
        const workbook = XLSX.read(arrayBuffer, { type: 'array' });

        // Seleciona a primeira planilha
        const guiaBackEnd = workbook.SheetNames[4];
        const worksheet = workbook.Sheets[guiaBackEnd];

        // Lê os valores das células específicas
        const endAtendidas = 'N4';
        const endAtendidasN = 'N5';

        const endOrigemN1 = 'P4'
        const endOrigemVal1 = 'Q4'
        const endOrigemN2 = 'P5'
        const endOrigemVal2 = 'Q5'
        const endOrigemN3 = 'P6'
        const endOrigemVal3 = 'Q6'
        const endOrigemN4 = 'P7'
        const endOrigemVal4 = 'Q7'

        const endEmpresaAtend = 'T4'
        const endEmpresaAtendN = 'T5'
        const endSemLig = 'T6'
        
        // ----------------------------------------------------------------------------------- //
        
        const valorAtendidas = worksheet[endAtendidas] ? worksheet[endAtendidas].v : null;
        const valorAtendidasN = worksheet[endAtendidasN] ? worksheet[endAtendidasN].v : null;
        
        const nomeOrigem1 = worksheet[endOrigemN1] ? worksheet[endOrigemN1].v : null;
        const valorOrigem1 = worksheet[endOrigemVal1] ? worksheet[endOrigemVal1].v : null;
        const nomeOrigem2 = worksheet[endOrigemN2] ? worksheet[endOrigemN2].v : null;
        const valorOrigem2 = worksheet[endOrigemVal2] ? worksheet[endOrigemVal2].v : null;
        const nomeOrigem3 = worksheet[endOrigemN3] ? worksheet[endOrigemN3].v : null;
        const valorOrigem3 = worksheet[endOrigemVal3] ? worksheet[endOrigemVal3].v : null;
        const nomeOrigem4 = worksheet[endOrigemN4] ? worksheet[endOrigemN4].v : null;
        const valorOrigem4 = worksheet[endOrigemVal4] ? worksheet[endOrigemVal4].v : null;

        const valorEmpresaAtend = worksheet[endEmpresaAtend] ? worksheet[endEmpresaAtend].v : null;
        const valorEmpresaAtendN = worksheet[endEmpresaAtendN] ? worksheet[endEmpresaAtendN].v : null;
        const valorSemLig = worksheet[endSemLig] ? worksheet[endSemLig].v : null;
        

        return { valorAtendidas, valorAtendidasN, nomeOrigem1, nomeOrigem2, nomeOrigem3, nomeOrigem4, valorOrigem1, valorOrigem2, valorOrigem3, valorOrigem4, valorEmpresaAtend, valorEmpresaAtendN, valorSemLig};
    }

    try {
        // Processa o arquivo Excel
        const { valorAtendidas, valorAtendidasN, nomeOrigem1, nomeOrigem2, nomeOrigem3, nomeOrigem4, valorOrigem1, valorOrigem2, valorOrigem3, valorOrigem4, valorEmpresaAtend, valorEmpresaAtendN, valorSemLig } = await fetchAndProcessExcel(fileUrl);

        // Atualiza o gráfico com os valores do Excel
        const chart1 = document.getElementById('grafico1');
        const chart2 = document.getElementById('grafico2');
        const chart3 = document.getElementById('grafico3');
        const chart4 = document.getElementById('grafico4')

        new Chart(chart1, {
            type: 'pie',
            data: {
                labels: ['Atendidas', 'Não atendidas'],
                datasets: [
                    {
                        label: 'Status',
                        data: [valorAtendidas, valorAtendidasN], // Usa os valores extraídos
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
                labels: [nomeOrigem1, nomeOrigem2, nomeOrigem3, nomeOrigem4],
                datasets: [
                    {
                        label: 'Status',
                        data: [valorOrigem1, valorOrigem2, valorOrigem3, valorOrigem4],
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
                labels: ['Empresas Atendidas', 'Empresas não atendidas', 'Sem ligação'],
                datasets: [
                    {
                        label: 'Status',
                        data: [valorEmpresaAtend, valorEmpresaAtendN, valorSemLig],
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
    } catch (error) {
        console.error('Erro:', error);
        document.getElementById('output').textContent = `Erro ao processar o arquivo: ${error.message}`;
    }
})();
