function gerarClientes() {
    return Array.from({ length: 20 }, () => {
        const codigo = Math.floor(Math.random() * 9000000) + 1000000;
        return `${codigo} - J BONFIM LINS`;
    });
}

function selecionarTerritorio(nome) {
    document.getElementById('territorios').style.display = 'none';
    document.getElementById('clientes').style.display = 'block';
    document.getElementById('territorio-selecionado').innerText = `Território: ${nome}`;
    const clientes = gerarClientes();
    document.getElementById('lista-clientes').innerHTML = clientes.map(cliente => 
        `<div class="cliente"><input type="checkbox" class="cliente-checkbox"> ${cliente}</div>`).join('');
}

function voltarMenu() {
// Obter as divs que representam os menus
const territorioMenu = document.getElementById('territorios');
const clientesMenu = document.getElementById('clientes');
const territorioPropostoMenu = document.getElementById('territorios-propostos');

// Se estamos na tela de 'territorios-propostos', volta para 'clientes'
if (territorioPropostoMenu && territorioPropostoMenu.style.display === 'block') {
territorioPropostoMenu.style.display = 'none';
clientesMenu.style.display = 'block';
} 
// Se estamos na tela de 'clientes', volta para 'territorios'
else if (clientesMenu.style.display === 'block') {
clientesMenu.style.display = 'none';
territorioMenu.style.display = 'block';
}
}


function mostrarMovimentacao() {
    document.getElementById('clientes').style.display = 'none';
    document.getElementById('movimentacao').style.display = 'block';
}

function voltarClientes() {
    document.getElementById('movimentacao').style.display = 'none';
    document.getElementById('clientes').style.display = 'block';
}

function selecionarMovimentacao(tipo) {
    document.getElementById('movimentacao').style.display = 'none';
    document.getElementById('territorio-proposto').style.display = 'block';
}

function selecionarTerritorioProposto(territorio) {
    document.getElementById('exportar-excel').style.display = 'block';
}

function exportarExcel() {
const wb = XLSX.utils.book_new();

// Cabeçalho da planilha
let ws_data = [['Código do Cliente', 'Código do Território de Origem', 'Código do Território de Destino']];

// Obter o território de origem e destino
const territorioOrigem = document.getElementById('territorio-selecionado').innerText.split(': ')[1];
const territorioDestino = document.querySelector('input[name="territorio"]:checked')?.value;
const tipoMovimentacao = document.querySelector('.movimentacao-container button.botao[disabled]')?.innerText.toLowerCase();

if (!territorioDestino) {
alert("Por favor, selecione um território proposto.");
return;
}

// Adicionar os dados dos clientes selecionados à planilha
const clientes = document.querySelectorAll('.cliente-checkbox:checked');

if (clientes.length === 0) {
alert("Selecione ao menos um cliente para exportar.");
return;
}

clientes.forEach(checkbox => {
const codigoCliente = checkbox.parentElement.innerText.trim().split(' ')[0]; // Captura o código do cliente

if (tipoMovimentacao === "multiatendimento") {
    // Multiatendimento: Deixar a célula do território de origem vazia
    ws_data.push([codigoCliente, "", territorioDestino.split(' ')[0]]);
} else {
    // Transferência: Preencher normalmente
    ws_data.push([codigoCliente, territorioOrigem.split(' ')[0], territorioDestino.split(' ')[0]]);
}
});

// Criar a planilha e adicioná-la ao arquivo Excel
const ws = XLSX.utils.aoa_to_sheet(ws_data);
XLSX.utils.book_append_sheet(wb, ws, 'Clientes');
XLSX.writeFile(wb, 'movimentacao.xlsx');
}
