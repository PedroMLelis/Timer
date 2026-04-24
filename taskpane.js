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
