window.addEventListener("DOMContentLoaded", () => {
    const dataAtual = new Date();
    const dia = String(dataAtual.getDate()).padStart(2, "0");
    const mes = String(dataAtual.getMonth() + 1).padStart(2, "0");
    const ano = dataAtual.getFullYear();
    const dataFormatada = `${dia}/${mes}/${ano}`;
    document.getElementById("dataAtual").textContent = dataFormatada;
});

async function gerarPDF() {
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF();
    let y = 10;

    doc.setFont("helvetica", "bold");
    doc.setFontSize(16);
    doc.text("Relatório de Atracação / Desatracação", 10, y);
    y += 15;

    const boxs = document.querySelectorAll(".boxs, .box");

    doc.setFontSize(12);

    boxs.forEach((box) => {
        const label = box.querySelector("label")?.innerText || "";
        const input = box.querySelector("input")?.value || "";
        const select = box.querySelector("select")?.value || "";
        const divText = box.querySelector("div")?.innerText || "";
        const para = box.querySelector("p")?.innerText || "";

        let valor = input || select || divText || para;

        if (label && valor) {
            doc.setFont("helvetica", "bold");
            doc.text(`${label}:`, 10, y);
            y += 6;

            doc.setFont("helvetica", "normal");
            doc.text(`${valor}`, 10, y);
            y += 10;
        } else if (input || select || divText) {
            const dataLabel =
                box.getAttribute("data-label") || "Tipo de Operação";

            doc.setFont("helvetica", "bold");
            doc.text(`${dataLabel}:`, 10, y);
            y += 6;

            doc.setFont("helvetica", "normal");
            doc.text(`${input || select || divText}`, 10, y);
            y += 10;
        }
    });

    doc.save("relatorio_atracacao.pdf");
}

function gerarExcel() {
    const campos = document.querySelectorAll(".boxs input, .box select");

    let dados = {};
    campos.forEach((campo) => {
        let label = campo.closest("div").querySelector("label");
        let nomeCampo = label ? label.innerText.trim() : "Campo";
        let valorCampo = campo.value || campo.innerText || "";
        dados[nomeCampo] = valorCampo;
    });

    // Adiciona data e hora automática
    const agora = new Date();
    const dataHoraFormatada = agora.toLocaleString("pt-BR");
    dados["Data/Hora do Registro"] = dataHoraFormatada;

    // Define o mês atual como nome da aba (ex: "Maio")
    const nomeMes = agora.toLocaleDateString("pt-BR", { month: "long" });
    const nomeMesCapitalizado =
        nomeMes.charAt(0).toUpperCase() + nomeMes.slice(1);

    // Recupera dados do localStorage
    let historicoCompleto =
        JSON.parse(localStorage.getItem("dadosExcelPorMes")) || {};

    // Se não existir entrada para o mês, cria uma nova
    if (!historicoCompleto[nomeMesCapitalizado]) {
        historicoCompleto[nomeMesCapitalizado] = [];
    }

    // Adiciona nova entrada
    historicoCompleto[nomeMesCapitalizado].push(dados);

    // Atualiza localStorage
    localStorage.setItem("dadosExcelPorMes", JSON.stringify(historicoCompleto));

    // Cria workbook
    const workbook = XLSX.utils.book_new();

    // Cria uma aba para cada mês que tenha dados
    for (const [mes, registros] of Object.entries(historicoCompleto)) {
        let sheetData = [];

        registros.forEach((registro, index) => {
            const cabecalhos = Object.keys(registro);
            const valores = Object.values(registro);
            sheetData.push(cabecalhos);
            sheetData.push(valores);
            sheetData.push([]); // linha em branco
        });

        const worksheet = XLSX.utils.aoa_to_sheet(sheetData);
        XLSX.utils.book_append_sheet(workbook, worksheet, mes);
    }

    // Salva o arquivo
    XLSX.writeFile(workbook, "atracacao_mensal.xlsx");
}

function coletarDadosFormulario() {
    const dados = [];
    const campos = document.querySelectorAll(".boxs, .box");

    campos.forEach((box, index) => {
        const label = box.querySelector("label")?.innerText || ""; // se houver label
        const select = box.querySelector("select")?.value || ""; // valor do select (Atracação/Desatracação)
        const input = box.querySelector("input")?.value || ""; // valor do input (Navio)
        const div = box.querySelector("div")?.innerText || ""; // caso haja algum texto dentro de div
        const p = box.querySelector("p")?.innerText || ""; // caso haja algum parágrafo
        const dataLabel = box.getAttribute("data-label"); // atributo data-label para nome do campo

        let valor = input || select || div || p; // tenta pegar o valor do input, select, div ou p

        let nomeCampo = label || dataLabel || ""; // usa o label ou o data-label como nome do campo

        // Se for o campo do select (Atracação/Desatracação), o nomeCampo será o valor selecionado
        if (nomeCampo && (valor || valor === "")) {
            dados.push({ Campo: nomeCampo, Valor: valor });
        }
    });

    return dados;
}

// Converte array tipo: [{Campo: "X", Valor: "Y"}] → { X: "Y" }
function converterArrayParaObjeto(array) {
    const obj = {};
    array.forEach((item) => {
        obj[item.Campo] = item.Valor;
    });
    return obj;
}

// Formata a data no padrão dd/mm/aaaa
function formatarData(data) {
    const dia = String(data.getDate()).padStart(2, "0");
    const mes = String(data.getMonth() + 1).padStart(2, "0");
    const ano = data.getFullYear();
    return `${dia}/${mes}/${ano}`;
}

// Deixa a primeira letra maiúscula (ex: "maio" → "Maio")
function capitalizarPrimeiraLetra(texto) {
    return texto.charAt(0).toUpperCase() + texto.slice(1);
}

async function salvarArquivo(nome, conteudoBlob) {
    const options = {
        suggestedName: nome,
        types: [
            {
                description: "Arquivos PDF",
                accept: { "application/pdf": [".pdf"] },
            },
        ],
    };
    const handle = await window.showSaveFilePicker(options);
    const writable = await handle.createWritable();
    await writable.write(conteudoBlob);
    await writable.close();
}

if ("serviceWorker" in navigator) {
    window.addEventListener("load", () => {
        navigator.serviceWorker
            .register("/service-worker.js")
            .then((registration) => {
                console.log(
                    "Service Worker registrado com sucesso: ",
                    registration
                );
            })
            .catch((error) => {
                console.log("Falha ao registrar o Service Worker: ", error);
            });
    });
}

function limparDadosExcel() {
    // Limpa os dados armazenados no localStorage
    localStorage.removeItem("relatoriosPorMes");
    alert("Dados do Excel limpos com sucesso!");
}
