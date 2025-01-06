// Importando dados do excel

const XLSX = require("xlsx")

const arquivo = XLSX.readFile("BaseDados.xlsx"); // Selecionando arquivo de trabalho Excel

const selectPlanilha = arquivo.SheetNames[0] // Selecionando a planila; 0 = Folha1
const planilha = arquivo.Sheets[selectPlanilha] // Puxando todos os dados da planilha.

const enderecoExterno = "B2"; // Endereco da celula Externo
const enderecoInterno = "B3"; // Endereco da celula Interno
const celulaExterno = planilha[enderecoExterno];
const celulaInterno = planilha[enderecoInterno];

let valorExterno = celulaExterno ? celulaExterno.v : null; // Valor em INT Externo
let valorInterno = celulaInterno ? celulaInterno.v : null; // Valor em INT Interno

console.log(valorExterno)
console.log(valorInterno)