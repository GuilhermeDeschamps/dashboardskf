document.addEventListener("DOMContentLoaded", () => {

const btnLogin =
    document.getElementById("btn-login");

const dashboard =
    document.getElementById("dashboard");

const loginScreen =
    document.getElementById("login-screen");

const erroLogin =
    document.getElementById("erro-login");

if (btnLogin) {

    btnLogin.addEventListener(
        "click",
        () => {

            const usuario =
                document.getElementById("usuario").value;

            const senha =
                document.getElementById("senha").value;

            if (
                usuario === "admin" &&
                senha === "123456"
            ) {

                loginScreen.style.display =
                    "none";

                dashboard.style.display =
                    "block";

            } else {

                erroLogin.textContent =
                    "Usuário ou senha inválidos";

            }

        }
    );

}

let dados = {

    notas: [],

    meses: [],

    evolucaoMensal: {}

};

let notasFiltradasAtual = [];

const uploadExcel =
    document.getElementById(
        "upload-excel"
    );

const btnUpload =
    document.getElementById(
        "btn-upload"
    );

console.log("BTN:", btnUpload);
console.log("INPUT:", uploadExcel);

async function carregarExcel() {

    console.log("1 - Entrou na função");

    const resposta =
        await fetch("INDUSTRIAL.xlsx");

    console.log("2 - Fetch OK");

    const arquivo =
        await resposta.arrayBuffer();

    console.log("3 - ArrayBuffer OK");

    const workbook =
        XLSX.read(
            arquivo,
            { type: "array" }
        );

    console.log("4 - Workbook OK");

    const sheet =
        workbook.Sheets[
            workbook.SheetNames[0]
        ];

    const linhas =
        XLSX.utils.sheet_to_json(sheet);

    console.log("5 - Linhas:", linhas);

    console.log("Primeira linha:", linhas[0]);

    const nomesMeses = [
        "Janeiro",
        "Fevereiro",
        "Março",
        "Abril",
        "Maio",
        "Junho",
        "Julho",
        "Agosto",
        "Setembro",
        "Outubro",
        "Novembro",
        "Dezembro"
    ];

    dados.notas = linhas.map(item => {

        let data;

if (typeof item["Scan/Email Date"] === "number") {

    data = new Date(
        (item["Scan/Email Date"] - 25569)
        * 86400
        * 1000
    );

} else {

    data = new Date(
        item["Scan/Email Date"]
    );

}

return {

    ...item,

    mes:
        nomesMeses[data.getMonth()] ||
        "Sem mês",

    fornecedor:
        item["Fornecedor"] || "Não informado",

    requisitante:
        item["Task Owner"] || "Não informado",

    valor:
        Number(item["Net Amount"]) || 0,

    status:
        (item["Status NF"] || "Sem pendência")
        .toString()
        .trim()

};

    });

const ordemMeses = [
    "Janeiro",
    "Fevereiro",
    "Março",
    "Abril",
    "Maio",
    "Junho",
    "Julho",
    "Agosto",
    "Setembro",
    "Outubro",
    "Novembro",
    "Dezembro"
];

dados.meses = ordemMeses.filter(mes =>
    dados.notas.some(item => item.mes === mes)
);

console.log(
    [...new Set(dados.notas.map(n => n.status))]
);


}

function formatarValor(valor) {

    if (valor >= 1000000) {
        return `R$ ${(valor / 1000000).toFixed(1)}M`;
    }

    if (valor >= 1000) {
        return `R$ ${(valor / 1000).toFixed(0)}K`;
    }

    return `R$ ${valor.toLocaleString("pt-BR")}`;
}

function exportarExcel() {

    if (!notasFiltradasAtual.length) {

        alert(
            "Nenhum registro encontrado."
        );

        return;
    }

   const dadosExportacao = notasFiltradasAtual.map(item => ({

    "ASV Ref Id": item["ASV Ref Id"],

    Fornecedor: item.fornecedor,

    Requisitante: item.requisitante,

    Mês: item.mes,

    Valor: item.valor,

    Status: item.status,

    "Data Recebimento": item["Scan/Email Date"]

}));


const ws =
    XLSX.utils.json_to_sheet(
        dadosExportacao
    );

ws["!cols"] = [

    { wch: 30 }, // ASV Ref Id
    { wch: 40 }, // Fornecedor
    { wch: 30 }, // Requisitante
    { wch: 15 }, // Mês
    { wch: 20 }, // Valor
    { wch: 25 }, // Status
    { wch: 20 }  // Data

];

ws["!autofilter"] = {
    ref: ws["!ref"]
};

ws["!freeze"] = {
    xSplit: 0,
    ySplit: 1
};

    const wb =
        XLSX.utils.book_new();

    XLSX.utils.book_append_sheet(
        wb,
        ws,
        "Notas Filtradas"
    );

    XLSX.writeFile(
    wb,
    `Dashboard_Notas_Fiscais_${
        new Date()
            .toLocaleDateString("pt-BR")
            .replace(/\//g, "-")
    }.xlsx`
);

}

// =====================
// TOP FORNECEDORES
// =====================

Object.entries(dados.evolucaoMensal).forEach(([mes, valor]) => {

    const elemento =
        document.getElementById(`valor-${mes}`);

    if (elemento) {
        elemento.textContent = valor;
    }

});

const listaRequisitantes =
    document.getElementById("lista-requisitantes");

const filtroRequisitante =
    document.getElementById("filtro-requisitante");

const btnRequisitante =
    document.getElementById("btn-requisitante");

const filtroMes =
    document.getElementById("filtro-mes");

const btnMes =
    document.getElementById("btn-mes");

const filtroFornecedor =
    document.getElementById("filtro-fornecedor");

const btnFornecedor =
    document.getElementById("btn-fornecedor");

const filtroStatus =
    document.getElementById("filtro-status");

const btnStatus =
    document.getElementById("btn-status");

const tabelaFornecedores =
    document.getElementById("top-fornecedores");

const donut =
    document.getElementById("grafico-status");

const lblAprovadas =
    document.getElementById("lbl-aprovadas");

const lblPendentes =
    document.getElementById("lbl-pendentes");

const lblRejeitadas =
    document.getElementById("lbl-rejeitadas");

const btnExportar =
    document.getElementById(
        "btn-exportar"
    );

function preencherFiltros() {

filtroMes.innerHTML = `

    <input
        type="text"
        id="busca-mes"
        placeholder="🔍 Buscar mês..."
        style="
            width:100%;
            padding:6px;
            margin-bottom:8px;
            border-radius:6px;
            border:none;
        "
    >

    <div>
        <input
            type="checkbox"
            id="chk-todos-mes"
        >
        <strong>Selecionar Todos</strong>
    </div>

    <hr style="margin:6px 0; border-color:rgba(255,255,255,.2)">
`;

dados.meses.forEach(mes => {

    filtroMes.innerHTML += `
        <div>
            <input
                type="checkbox"
                class="chk-mes"
                value="${mes}"
            >
            ${mes}
        </div>
    `;

});

const buscaMes =
    document.getElementById(
        "busca-mes"
    );

if (buscaMes) {

    buscaMes.addEventListener(
        "input",
        () => {

            const texto =
                buscaMes.value
                .toLowerCase();

            document
            .querySelectorAll(
                "#filtro-mes .chk-mes"
            )
            .forEach(chk => {

                const linha =
                    chk.parentElement;

                linha.style.display =
                    linha.textContent
                        .toLowerCase()
                        .includes(texto)
                    ? "flex"
                    : "none";

            });

        }
    );

}


document
.querySelectorAll(".chk-mes")
.forEach(chk => {

    chk.addEventListener(
        "change",
        atualizarDashboard
    );

});

const chkTodosMes =
    document.getElementById(
        "chk-todos-mes"
    );

if (chkTodosMes) {

    chkTodosMes.addEventListener(
        "change",
        () => {

            document
            .querySelectorAll(
                ".chk-mes"
            )
            .forEach(chk => {

                chk.checked =
                    chkTodosMes.checked;

            });

            atualizarDashboard();

        }
    );

}

    filtroRequisitante.innerHTML = `
    <input
        type="text"
        id="busca-requisitante"
        placeholder="🔍 Buscar requisitante..."
        style="
            width:100%;
            padding:6px;
            margin-bottom:8px;
            border-radius:6px;
            border:none;
        "
    >

    <div>
        <input
            type="checkbox"
            id="chk-todos-requisitante"
        >
        <strong>Selecionar Todos</strong>
    </div>

    <hr style="margin:6px 0; border-color:rgba(255,255,255,.2)">
`;

const requisitantes = [
    ...new Set(
        dados.notas
            .map(item => item.requisitante)
            .filter(Boolean)
    )
].sort();

requisitantes.forEach(nome => {

    filtroRequisitante.innerHTML += `
        <div>
            <input
                type="checkbox"
                class="chk-requisitante"
                value="${nome}"
            >
            ${nome}
        </div>
    `;

});

document
.querySelectorAll(".chk-requisitante")
.forEach(chk => {

    chk.addEventListener(
        "change",
        atualizarDashboard
    );

});

const buscaRequisitante =
    document.getElementById(
        "busca-requisitante"
    );

if (buscaRequisitante) {

const chkTodosRequisitante =
    document.getElementById(
        "chk-todos-requisitante"
    );

if (chkTodosRequisitante) {

    chkTodosRequisitante.addEventListener(
        "change",
        () => {

            document
            .querySelectorAll(
                ".chk-requisitante"
            )
            .forEach(chk => {

                chk.checked =
                    chkTodosRequisitante.checked;

            });

            atualizarDashboard();

        }
    );

}

    buscaRequisitante.addEventListener(
        "input",
        () => {

            const texto =
                buscaRequisitante.value
                .toLowerCase();

            document
            .querySelectorAll(
                "#filtro-requisitante .chk-requisitante"
            )
            .forEach(chk => {

                const linha =
                    chk.parentElement;

                linha.style.display =
                    linha.textContent
                        .toLowerCase()
                        .includes(texto)
                    ? "flex"
                    : "none";

            });

        }
    );

}

filtroStatus.innerHTML = `
    <input
        type="text"
        id="busca-status"
        placeholder="🔍 Buscar status..."
        style="
            width:100%;
            padding:6px;
            margin-bottom:8px;
            border-radius:6px;
            border:none;
        "
    >

    <div>
        <input
            type="checkbox"
            id="chk-todos-status"
        >
        <strong>Selecionar Todos</strong>
    </div>

    <hr style="margin:6px 0; border-color:rgba(255,255,255,.2)">
`;

const statusUnicos = [
    ...new Set(
        dados.notas
            .map(item => item.status)
            .filter(Boolean)
    )
].sort();

statusUnicos.forEach(status => {

    filtroStatus.innerHTML += `
        <div>
            <input
                type="checkbox"
                class="chk-status"
                value="${status}"
            >
            ${status}
        </div>
    `;

});

document
.querySelectorAll(".chk-status")
.forEach(chk => {

    chk.addEventListener(
        "change",
        atualizarDashboard
    );

});

const buscaStatus =
    document.getElementById(
        "busca-status"
    );

if (buscaStatus) {

    buscaStatus.addEventListener(
        "input",
        () => {

            const texto =
                buscaStatus.value
                .toLowerCase();

            document
            .querySelectorAll(
                "#filtro-status .chk-status"
            )
            .forEach(chk => {

                const linha =
                    chk.parentElement;

                linha.style.display =
                    linha.textContent
                        .toLowerCase()
                        .includes(texto)
                    ? "flex"
                    : "none";

            });

        }
    );

}

const chkTodosStatus =
    document.getElementById(
        "chk-todos-status"
    );

if (chkTodosStatus) {

    chkTodosStatus.addEventListener(
        "change",
        () => {

            document
            .querySelectorAll(
                ".chk-status"
            )
            .forEach(chk => {

                chk.checked =
                    chkTodosStatus.checked;

            });

            atualizarDashboard();

        }
    );

}

filtroFornecedor.innerHTML = `
    <input
        type="text"
        id="busca-fornecedor"
        placeholder="🔍 Buscar fornecedor..."
        style="
            width:100%;
            padding:6px;
            margin-bottom:8px;
            border-radius:6px;
            border:none;
        "
    >

    <div>
        <input
            type="checkbox"
            id="chk-todos"
        >
        <strong>Selecionar Todos</strong>
    </div>

    <hr style="margin:6px 0; border-color:rgba(255,255,255,.2)">
`;

const fornecedores = [
    ...new Set(
        dados.notas
            .map(item => item.fornecedor)
            .filter(Boolean)
    )
].sort();

fornecedores.forEach(nome => {

filtroFornecedor.innerHTML += `
    <div>
        <input
            type="checkbox"
            class="chk-fornecedor"
            value="${nome}"
        >
        ${nome}
    </div>
`;

});

document
.querySelectorAll(".chk-fornecedor")
.forEach(chk => {

    chk.addEventListener(
        "change",
        atualizarDashboard
    );

});

const buscaFornecedor =
    document.getElementById(
        "busca-fornecedor"
    );

if (buscaFornecedor) {

    buscaFornecedor.addEventListener(
        "input",
        () => {

            const texto =
                buscaFornecedor.value
                .toLowerCase();

            document
            .querySelectorAll(
                "#filtro-fornecedor .chk-fornecedor"
            )
            .forEach(chk => {

                const linha =
                    chk.parentElement;

                linha.style.display =
                    linha.textContent
                        .toLowerCase()
                        .includes(texto)
                    ? "flex"
                    : "none";

            });

        }
    );

}



const chkTodos =
    document.getElementById("chk-todos");

if (chkTodos) {

    chkTodos.addEventListener(
        "change",
        () => {

            document
            .querySelectorAll(".chk-fornecedor")
            .forEach(chk => {

                chk.checked =
                    chkTodos.checked;

            });

            atualizarDashboard();

        }
    );

}

} 

function atualizarGraficoMensal(resumoMeses) {

    const meses = [
        "Janeiro",
        "Fevereiro",
        "Março",
        "Abril",
        "Maio",
        "Junho",
        "Julho",
        "Agosto"
    ];

const tooltip =
    document.getElementById(
        "tooltip-grafico"
    );

    const valores = meses.map(
        mes => resumoMeses[mes] || 0
    );

    const maiorValor =
        Math.max(...valores, 1);

    const pontos = [];

    const ids = [
    "jan",
    "fev",
    "mar",
    "abr",
    "mai",
    "jun",
    "jul",
    "ago"
];

valores.forEach((valor, index) => {

    const posicoesX = [
    20,
    110,
    200,
    290,
    380,
    470,
    560,
    630
];

const x = posicoesX[index];

    const y =
        200 -
        ((valor / maiorValor) * 160);

    const ponto =
        document.getElementById(
            `ponto-${ids[index]}`
        );

    if (ponto) {

    ponto.setAttribute(
        "cx",
        x
    );

    ponto.setAttribute(
        "cy",
        y
    );

    if (tooltip) {

        ponto.onmousemove = (e) => {

            tooltip.style.display =
                "block";

            tooltip.style.left =
                (e.clientX + 15) + "px";

            tooltip.style.top =
                (e.clientY - 15) + "px";

            tooltip.innerHTML = `
                <strong>${meses[index]}</strong>
                <br>
                ${formatarValor(valor)}
            `;

        };

        ponto.onmouseleave = () => {

            tooltip.style.display =
                "none";

        };

    }

}

const texto =
    document.getElementById(
        `valor-${ids[index]}`
    );

if (texto) {

    texto.setAttribute(
        "y",
        y - 25
    );

    texto.setAttribute(
        "x",
        x
    );

}

    pontos.push(
        `${x},${y}`
    );

});

    const linha =
        "M" +
        pontos.join(" L");

    const area =
        linha +
        " L630,220 L20,220 Z";

const linhaGrafico =
    document.getElementById(
        "grafico-linha"
    );

const areaGrafico =
    document.getElementById(
        "grafico-area"
    );

if (linhaGrafico) {
    linhaGrafico.setAttribute(
        "d",
        linha
    );
}

if (areaGrafico) {
    areaGrafico.setAttribute(
        "d",
        area
    );
}

}

function mostrarDetalhesFornecedor(nomeFornecedor) {

    const box =
        document.getElementById(
            "box-detalhes"
        );

    const tabela =
        document.getElementById(
            "detalhes-notas"
        );

    const titulo =
        document.getElementById(
            "titulo-detalhes"
        );

    const registros =
        dados.notas.filter(
            item =>
                item.fornecedor === nomeFornecedor
        );

    titulo.textContent =
        `Notas do fornecedor: ${nomeFornecedor}`;

    tabela.innerHTML = "";

    registros.forEach(item => {

        tabela.innerHTML += `
            <tr>
                <td>${item["ASV Ref Id"] || "-"}</td>
                <td>${item.fornecedor}</td>
                <td>${item.requisitante}</td>
                <td>${formatarValor(item.valor)}</td>
                <td>${item.status}</td>
            </tr>
        `;

    });

    box.style.display = "block";

}

function atualizarDashboard() {

const mesesSelecionados =
[...document.querySelectorAll(
    ".chk-mes:checked"
)]
.map(item => item.value);

    const requisitantesSelecionados =
    [...document.querySelectorAll(
        ".chk-requisitante:checked"
    )]
    .map(item => item.value);

    const fornecedoresSelecionados =
    [...document.querySelectorAll(
        ".chk-fornecedor:checked"
    )]
    .map(item => item.value);

const statusSelecionados =
[...document.querySelectorAll(
    ".chk-status:checked"
)]
.map(item => item.value);

if (btnFornecedor) {

    btnFornecedor.textContent =

        fornecedoresSelecionados.length === 0

        ? "Fornecedores ▼"

        : `Fornecedores (${fornecedoresSelecionados.length}) ▼`;

}

if (btnRequisitante) {

    btnRequisitante.textContent =

        requisitantesSelecionados.length === 0

        ? "Requisitantes ▼"

        : `Requisitantes (${requisitantesSelecionados.length}) ▼`;

}

if (btnStatus) {

    btnStatus.textContent =

        statusSelecionados.length === 0

        ? "Status ▼"

        : `Status (${statusSelecionados.length}) ▼`;

}

if (btnMes) {

    btnMes.textContent =

        mesesSelecionados.length === 0

        ? "Meses ▼"

        : `Meses (${mesesSelecionados.length}) ▼`;

}

    const notasFiltradas =
        dados.notas.filter(item => {

            const filtroMesOk =

    mesesSelecionados.length === 0 ||

    mesesSelecionados.includes(
        item.mes
    );

            const filtroReqOk =

    requisitantesSelecionados.length === 0 ||

    requisitantesSelecionados.includes(
        item.requisitante
    );

            const filtroFornecedorOk =

    fornecedoresSelecionados.length === 0 ||

    fornecedoresSelecionados.includes(
        item.fornecedor

    );

const filtroStatusOk =

    statusSelecionados.length === 0 ||

    statusSelecionados.includes(
        item.status
    );

            return (
    filtroMesOk &&
    filtroReqOk &&
    filtroFornecedorOk &&
    filtroStatusOk
);

        });

    const valorTotal =
        notasFiltradas.reduce(
            (soma, item) => soma + item.valor,
            0
        );

const resumoMeses = {};

notasFiltradas.forEach(item => {

    if (!resumoMeses[item.mes]) {
        resumoMeses[item.mes] = 0;
    }

    resumoMeses[item.mes] += item.valor;

});

notasFiltradasAtual = notasFiltradas;

let melhorMes = "-";
let maiorValorMes = 0;

Object.entries(resumoMeses).forEach(([mes, valor]) => {

    if (valor > maiorValorMes) {

        maiorValorMes = valor;
        melhorMes = mes;

    }

});

const cardMelhorMes =
    document.getElementById("melhor-mes");

if (cardMelhorMes) {

    cardMelhorMes.textContent =
        `Melhor mês: ${melhorMes}`;

}

const mapaMeses = {
    Janeiro: "jan",
    Fevereiro: "fev",
    Março: "mar",
    Abril: "abr",
    Maio: "mai",
    Junho: "jun",
    Julho: "jul",
    Agosto: "ago",
    Setembro: "set",
    Outubro: "out",
    Novembro: "nov",
    Dezembro: "dez"
};

Object.entries(mapaMeses).forEach(([mes, id]) => {

    const valor =
        resumoMeses[mes] || 0;

    const elemento =
        document.getElementById(`valor-${id}`);

    if (elemento) {

        elemento.textContent =
            formatarValor(valor)
                .replace("R$ ", "");

    }

});

atualizarGraficoMensal(
    resumoMeses
);

document
.querySelectorAll(".mes-grafico")
.forEach(mes => {

    mes.classList.remove("ativo");

});

mesesSelecionados.forEach(nomeMes => {

    const elemento =
        document.querySelector(
            `.mes-grafico[data-mes="${nomeMes}"]`
        );

    if (elemento) {

        elemento.classList.add(
            "ativo"
        );

    }

});

    document.getElementById("valor-total").textContent =
    formatarValor(valorTotal);

    const fornecedoresUnicos =
        new Set(
            notasFiltradas.map(
                item => item.fornecedor
            )
        );

    document.getElementById("fornecedores").textContent =
        fornecedoresUnicos.size;

    const requisitantesUnicos =
        new Set(
            notasFiltradas.map(
                item => item.requisitante
            )
        );

    document.getElementById("requisitantes").textContent =
        requisitantesUnicos.size;

const pendencias =
    notasFiltradas.filter(
        item => item.status &&
        item.status.trim() !== ""
    ).length;

document.getElementById("pendencias").textContent =
    pendencias;

    document.getElementById("status-total").innerHTML = `
        <div>
            <div>${notasFiltradas.length}</div>
            <small>NFs</small>
        </div>
    `;

const resumoFornecedores = {};

notasFiltradas.forEach(item => {

    if (!resumoFornecedores[item.fornecedor]) {
        resumoFornecedores[item.fornecedor] = 0;
    }

    resumoFornecedores[item.fornecedor] += item.valor;

});

const totalFornecedores =
    Math.max(
        Object.values(resumoFornecedores)
            .reduce((a, b) => a + b, 0),
        1
    );

let htmlFornecedores = "";

Object.entries(resumoFornecedores)
    .sort((a, b) => b[1] - a[1])
    .forEach(([nome, valor], index) => {

        const percentual =
            ((valor / totalFornecedores) * 100)
            .toFixed(0);

        htmlFornecedores += `
    <tr class="linha-fornecedor"
        data-fornecedor="${nome}">

        <td>${index + 1}</td>
        <td>${nome}</td>
        <td>${formatarValor(valor)}</td>
        <td>${percentual}%</td>

    </tr>
`;

    });

tabelaFornecedores.innerHTML =
    htmlFornecedores;

document
.querySelectorAll(".linha-fornecedor")
.forEach(linha => {

    linha.addEventListener(
        "click",
        () => {

            const fornecedor =
                linha.dataset.fornecedor;

            const checkbox =
                [...document.querySelectorAll(
                    ".chk-fornecedor"
                )]
                .find(
                    chk =>
                    chk.value === fornecedor
                );

            if (!checkbox) return;

            const jaSelecionado =
                checkbox.checked &&
                document.querySelectorAll(
                    ".chk-fornecedor:checked"
                ).length === 1;

            document
            .querySelectorAll(
                ".chk-fornecedor"
            )
            .forEach(chk => {

                chk.checked = false;

            });

            if (!jaSelecionado) {

                checkbox.checked = true;

            }

            atualizarDashboard();

            if (!jaSelecionado) {

                mostrarDetalhesFornecedor(
                    fornecedor
                );

            } else {

                document.getElementById(
                    "box-detalhes"
                ).style.display = "none";

            }

        }
    );

});

listaRequisitantes.innerHTML = "";

const resumoRequisitantes = {};

notasFiltradas.forEach(item => {

    if (!resumoRequisitantes[item.requisitante]) {
        resumoRequisitantes[item.requisitante] = 0;
    }

    resumoRequisitantes[item.requisitante] += item.valor;

});

const maiorValor =
    Math.max(
        ...Object.values(resumoRequisitantes),
        1
    );

Object.entries(resumoRequisitantes)
    .sort((a, b) => b[1] - a[1])
    .forEach(([nome, valor]) => {

        const percentual =
            (valor / maiorValor) * 100;

        listaRequisitantes.innerHTML += `
            <div class="progress">

                <div class="progress-header">
                    <span>${nome}</span>
                    <span>
                    ${formatarValor(valor)}
                    </span>
                </div>

                <div class="bar-bg">
                    <div class="bar"
                        style="width:${percentual}%">
                    </div>
                </div>

            </div>
        `;

    });

const resumoStatus = {};

notasFiltradas.forEach(item => {

    resumoStatus[item.status] =
        (resumoStatus[item.status] || 0) + 1;

});

const statusOrdenados =
    Object.entries(resumoStatus)
        .sort((a, b) => b[1] - a[1]);

const totalStatus =
    Math.max(
        notasFiltradas.length,
        1
    );

const s1 = statusOrdenados[0] || ["Sem Status", 0];
const s2 = statusOrdenados[1] || ["", 0];
const s3 = statusOrdenados[2] || ["", 0];

const p1 = (s1[1] / totalStatus) * 100;
const p2 = (s2[1] / totalStatus) * 100;

donut.style.background = `
    conic-gradient(
        #00d4ff 0% ${p1}%,
        #8b5cf6 ${p1}% ${p1 + p2}%,
        #ff5ad9 ${p1 + p2}% 100%
    )
`;

lblAprovadas.textContent =
    `${s1[0]} (${s1[1]})`;

lblPendentes.textContent =
    `${s2[0]} (${s2[1]})`;

lblRejeitadas.textContent =
    `${s3[0]} (${s3[1]})`;

    console.log(
        "Notas filtradas:",
        notasFiltradas
    );

}

// =====================
// EVENTO FILTRO
// =====================

if (filtroMes) {

    filtroMes.addEventListener(
        "change",
        atualizarDashboard
    );

}

if (filtroRequisitante) {

    filtroRequisitante.addEventListener(
        "change",
        atualizarDashboard
    );

}

if (filtroFornecedor) {

    filtroFornecedor.addEventListener(
        "change",
        atualizarDashboard
    );

}

if (filtroStatus) {

    filtroStatus.addEventListener(
        "change",
        atualizarDashboard
    );

}

if (btnFornecedor) {

    btnFornecedor.addEventListener(
        "click",
        () => {

            if (
                getComputedStyle(
                    filtroFornecedor
                ).display === "block"
            ) {

                filtroFornecedor.style.display =
                    "none";

            } else {

                filtroFornecedor.style.display =
                    "block";

            }

        }
    );

}

if (btnRequisitante) {

    btnRequisitante.addEventListener(
        "click",
        () => {

            if (
                getComputedStyle(
                    filtroRequisitante
                ).display === "block"
            ) {

                filtroRequisitante.style.display =
                    "none";

            } else {

                filtroRequisitante.style.display =
                    "block";

            }

        }
    );

}

if (btnStatus) {

    btnStatus.addEventListener(
        "click",
        () => {

            if (
                getComputedStyle(
                    filtroStatus
                ).display === "block"
            ) {

                filtroStatus.style.display =
                    "none";

            } else {

                filtroStatus.style.display =
                    "block";

            }

        }
    );

}

if (btnMes) {

    btnMes.addEventListener(
        "click",
        () => {

            if (
                getComputedStyle(
                    filtroMes
                ).display === "block"
            ) {

                filtroMes.style.display =
                    "none";

            } else {

                filtroMes.style.display =
                    "block";

            }

        }
    );

}

document.addEventListener(
    "click",
    (e) => {

        if (
            filtroFornecedor &&
            btnFornecedor
        ) {

            const clicouFornecedor =
                filtroFornecedor.contains(e.target);

            const clicouBtnFornecedor =
                btnFornecedor.contains(e.target);

            if (
                !clicouFornecedor &&
                !clicouBtnFornecedor
            ) {

                filtroFornecedor.style.display =
                    "none";

            }

        }


if (
    filtroRequisitante &&
    btnRequisitante
) {

    const clicouRequisitante =
        filtroRequisitante.contains(e.target);

    const clicouBtnRequisitante =
        btnRequisitante.contains(e.target);

    if (
        !clicouRequisitante &&
        !clicouBtnRequisitante
    ) {

        filtroRequisitante.style.display =
            "none";

    }

}

if (
    filtroStatus &&
    btnStatus
) {

    const clicouStatus =
        filtroStatus.contains(e.target);

    const clicouBtnStatus =
        btnStatus.contains(e.target);

    if (
        !clicouStatus &&
        !clicouBtnStatus
    ) {

        filtroStatus.style.display =
            "none";

    }

} // <-- FECHA O BLOCO STATUS

if (
    filtroMes &&
    btnMes
) {

    const clicouMes =
        filtroMes.contains(e.target);

    const clicouBtnMes =
        btnMes.contains(e.target);

    if (
        !clicouMes &&
        !clicouBtnMes
    ) {

        filtroMes.style.display =
            "none";

    }

}

    }
);

if (btnExportar) {

    btnExportar.addEventListener(
        "click",
        exportarExcel
    );

}

if (btnUpload && uploadExcel) {

    btnUpload.addEventListener(
        "click",
        () => {

            console.log("Abrindo seletor");

            uploadExcel.click();

        }
    );

    uploadExcel.addEventListener(
        "change",
        async (e) => {

            const arquivo =
                e.target.files[0];

           document.getElementById(
    "arquivo-carregado"
).textContent =
    `Base carregada: ${arquivo.name}`;


            if (!arquivo) return;

            const buffer =
                await arquivo.arrayBuffer();

            const workbook =
                XLSX.read(
                    buffer,
                    { type: "array" }
                );

            const sheet =
                workbook.Sheets[
                    workbook.SheetNames[0]
                ];

            const linhas =
                XLSX.utils.sheet_to_json(
                    sheet
                );

            dados.notas = [];

            const nomesMeses = [
                "Janeiro",
                "Fevereiro",
                "Março",
                "Abril",
                "Maio",
                "Junho",
                "Julho",
                "Agosto",
                "Setembro",
                "Outubro",
                "Novembro",
                "Dezembro"
            ];

            dados.notas = linhas.map(item => {

                let data;

                if (
                    typeof item["Scan/Email Date"] === "number"
                ) {

                    data = new Date(
                        (item["Scan/Email Date"] - 25569)
                        * 86400
                        * 1000
                    );

                } else {

                    data = new Date(
                        item["Scan/Email Date"]
                    );

                }

                return {

                    ...item,

                    mes:
                        nomesMeses[
                            data.getMonth()
                        ] || "Sem mês",

                    fornecedor:
                        item["Fornecedor"] ||
                        "Não informado",

                    requisitante:
                        item["Task Owner"] ||
                        "Não informado",

                    valor:
                        Number(
                            item["Net Amount"]
                        ) || 0,

                    status:
                        (
                            item["Status NF"] ||
                            "Sem pendência"
                        )
                        .toString()
                        .trim()

                };

            });

            preencherFiltros();

            atualizarDashboard();

        }
    );
}

document
.querySelectorAll(".mes-grafico")
.forEach(mes => {

    mes.addEventListener(
        "click",
        () => {

            const nomeMes =
                mes.dataset.mes;

            const checkbox =
                [...document.querySelectorAll(
                    ".chk-mes"
                )]
                .find(
                    chk =>
                    chk.value === nomeMes
                );

            if (!checkbox) return;

            const jaSelecionado =
                checkbox.checked &&
                document.querySelectorAll(
                    ".chk-mes:checked"
                ).length === 1;

            document
            .querySelectorAll(
                ".chk-mes"
            )
            .forEach(chk => {

                chk.checked = false;

            });

            if (!jaSelecionado) {

                checkbox.checked = true;

            }

            atualizarDashboard();

        }
    );

});

(async () => {

    await carregarExcel();

    preencherFiltros();

    atualizarDashboard();

})();

});
