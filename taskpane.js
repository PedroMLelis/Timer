console.log("VERSAO NOVA 1.0");

// Espera o DOM carregar (funciona em qualquer ambiente)
document.addEventListener("DOMContentLoaded", () => {
    const saveBtn = document.getElementById("save");

    if (!saveBtn) {
        console.error("Botão salvar não encontrado");
        return;
    }

    saveBtn.onclick = () => {
        const config = {
            startSlide: parseInt(document.getElementById("start").value),
            endSlide: parseInt(document.getElementById("end").value),
            duration: parseInt(document.getElementById("duration").value),
            color: document.getElementById("color").value,
            size: parseInt(document.getElementById("size").value),
            jumpTarget: parseInt(document.getElementById("jump").value)
        };

        console.log("Config capturada:", config);

        // 🔥 Se estiver dentro do PowerPoint
        if (typeof Office !== "undefined" && Office.context?.document) {
            try {
                Office.context.document.settings.set("timerConfig", config);

                Office.context.document.settings.saveAsync((res) => {
                    if (res.status === Office.AsyncResultStatus.Succeeded) {
                        alert("Configuração salva no PowerPoint!");
                    } else {
                        console.error("Erro ao salvar:", res.error);
                        alert("Erro ao salvar no PowerPoint.");
                    }
                });

            } catch (err) {
                console.error("Erro Office:", err);
                fallbackSave(config);
            }

        } else {
            // 🌐 Modo navegador
            fallbackSave(config);
        }
    };
});

// 🔁 Fallback para navegador / web limitado
function fallbackSave(config) {
    try {
        localStorage.setItem("timerConfig", JSON.stringify(config));
        alert("Configuração salva (modo navegador)");
    } catch (err) {
        console.error("Erro localStorage:", err);
        alert("Erro ao salvar localmente");
    }
}

// 🚀 Inserir timer no slide (chamado pelo manifest)
function insertTimer() {
    if (typeof Office === "undefined") {
        alert("Isso só funciona dentro do PowerPoint");
        return;
    }

    Office.context.document.setSelectedDataAsync(
        `<iframe src="https://pedromelis.github.io/Timer/timer.html" width="300" height="150" frameborder="0"></iframe>`,
        { coercionType: Office.CoercionType.Html },
        (res) => {
            if (res.status === Office.AsyncResultStatus.Succeeded) {
                console.log("Timer inserido com sucesso");
            } else {
                console.error("Erro ao inserir:", res.error);
                alert("Erro ao inserir timer");
            }
        }
    );
}
