console.log("VERSAO FINAL TASKPANE");

document.addEventListener("DOMContentLoaded", () => {
    const btn = document.getElementById("save");
    const status = document.getElementById("status");

    if (!btn) {
        console.error("Botão salvar não encontrado");
        return;
    }

    btn.onclick = () => {
        const config = {
            startSlide: parseInt(document.getElementById("start").value),
            endSlide: parseInt(document.getElementById("end").value),
            duration: parseInt(document.getElementById("duration").value),
            color: document.getElementById("color").value,
            size: parseInt(document.getElementById("size").value),
            jumpTarget: parseInt(document.getElementById("jump").value)
        };

        console.log("Config:", config);

        // 🔵 Dentro do PowerPoint
        if (typeof Office !== "undefined" && Office.context?.document) {

            try {
                Office.context.document.settings.set("timerConfig", config);

                Office.context.document.settings.saveAsync((res) => {
                    if (res.status === Office.AsyncResultStatus.Succeeded) {
                        console.log("Salvo no PowerPoint");
                        status.innerText = "✅ Salvo no PowerPoint";
                        status.style.color = "green";
                    } else {
                        console.error(res.error);
                        status.innerText = "❌ Erro ao salvar";
                        status.style.color = "red";
                    }
                });

            } catch (err) {
                console.error("Erro Office:", err);
                fallbackSave(config);
            }

        } else {
            // 🌐 Navegador
            fallbackSave(config);
        }
    };
});

// 🔁 fallback
function fallbackSave(config) {
    const status = document.getElementById("status");

    try {
        localStorage.setItem("timerConfig", JSON.stringify(config));
        console.log("Salvo localStorage");
        status.innerText = "💾 Salvo (modo navegador)";
        status.style.color = "blue";
    } catch (err) {
        console.error(err);
        status.innerText = "❌ Erro ao salvar localmente";
        status.style.color = "red";
    }
}

// 🚀 Inserir timer no slide
function insertTimer() {
    if (typeof Office === "undefined") return;

    Office.context.document.setSelectedDataAsync(
        `<iframe src="https://pedromelis.github.io/Timer/timer.html" width="300" height="150" frameborder="0"></iframe>`,
        { coercionType: Office.CoercionType.Html }
    );
}
