console.log("VERSAO NOVA 123");

Office.onReady(() => {
    document.getElementById("save").onclick = saveConfig;
});

function saveConfig() {
    const config = {
        startSlide: parseInt(start.value),
        endSlide: parseInt(end.value),
        duration: parseInt(duration.value),
        color: color.value,
        size: parseInt(size.value),
        jumpTarget: parseInt(jump.value)
    };

    localStorage.setItem("timerConfig", JSON.stringify(config));
    alert("Configuração salva!");
}

// CHAMADO PELO BOTÃO DO MANIFEST
function insertTimer() {
    Office.context.document.setSelectedDataAsync(
        `<iframe src="https://PedroMLelis.github.io/Timer/timer.html" width="300" height="150" frameborder="0"></iframe>`,
        { coercionType: Office.CoercionType.Html }
    );
}
if (typeof Office === "undefined") {
    console.log("Rodando fora do Office");

    document.getElementById("save").onclick = () => {
        const config = {
            startSlide: parseInt(document.getElementById("start").value),
            endSlide: parseInt(document.getElementById("end").value),
            duration: parseInt(document.getElementById("duration").value),
            color: document.getElementById("color").value,
            size: parseInt(document.getElementById("size").value),
            jumpTarget: parseInt(document.getElementById("jump").value)
        };

        localStorage.setItem("timerConfig", JSON.stringify(config));
        alert("Salvo (modo navegador)");
    };
}
