function adicionarItem() {
  // Obter os valores dos campos de entrada
  var nome = document.getElementById("nome").value;
  var fornecedor = document.getElementById("fornecedor").value;
  var valor = document.getElementById("valor").value;
  var estado = document.getElementById("estado").value;
  var quantidade = document.getElementById("quantidade").value;

  // Validar se os campos estão preenchidos
  if (!nome || !fornecedor || !valor || !estado || !quantidade) {
    alert("Preencha todos os campos.");
    return;
  }

  // Criar uma linha na tabela
  var tabela = document
    .getElementById("tabela")
    .getElementsByTagName("tbody")[0];
  var novaLinha = tabela.insertRow(tabela.rows.length);
  var celulaNome = novaLinha.insertCell(0);
  var celulaFornecedor = novaLinha.insertCell(1);
  var celulaValor = novaLinha.insertCell(2);
  var celulaEstado = novaLinha.insertCell(3);
  var celulaQuantidade = novaLinha.insertCell(4);

  // Preencher as células com os valores
  celulaNome.innerHTML = nome;
  celulaFornecedor.innerHTML = fornecedor;
  celulaValor.innerHTML = valor;
  celulaEstado.innerHTML = estado;
  celulaQuantidade.innerHTML = quantidade;

  // Limpar os campos de entrada
  document.getElementById("nome").value = "";
  document.getElementById("fornecedor").value = "";
  document.getElementById("valor").value = "";
  document.getElementById("estado").value = "";
  document.getElementById("quantidade").value = "";
}

function exportarParaExcel() {
  var tabela = document.getElementById("tabela");
  var nomeArquivo = "tabela_produtos.xlsx";
  var wb = XLSX.utils.table_to_book(tabela, { sheet: "Tabela de Produtos" });
  XLSX.writeFile(wb, nomeArquivo);
}
