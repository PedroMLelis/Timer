let config;
let polling;

Office.onReady(() => {
    config = JSON.parse(localStorage.getItem("timerConfig"));
    if (!config) return;

    applyStyle();
    startPolling();
});

function applyStyle() {
    const el = document.getElementById("timer-display");
    el.style.color = config.color;
    el.style.fontSize = config.size + "px";
}

function startPolling() {
    polling = setInterval(() => {
        Office.context.document.getSelectedDataAsync(
            Office.CoercionType.SlideRange,
            (res) => {
                if (res.status !== Office.AsyncResultStatus.Succeeded) return;

                const index = res.value.slides[0].index;
                processSlide(index);
            }
        );
    }, 1000);
}

function processSlide(index) {
    if (index < config.startSlide || index > config.endSlide) {
        document.body.style.visibility = "hidden";
        return;
    }

    document.body.style.visibility = "visible";

    if (!localStorage.getItem("timerEnd")) {
        localStorage.setItem(
            "timerEnd",
            Date.now() + config.duration * 1000
        );
    }

    const end = parseInt(localStorage.getItem("timerEnd"));
    const left = Math.max(0, end - Date.now());

    update(left);

    if (left === 0 && config.jumpTarget > 0) {
        Office.context.document.goToByIdAsync(
            config.jumpTarget,
            Office.GoToType.Index
        );
    }
}

function update(ms) {
    const s = Math.ceil(ms / 1000);
    const m = Math.floor(s / 60);
    const sec = s % 60;

    document.getElementById("timer-display").innerText =
        `${m.toString().padStart(2,'0')}:${sec.toString().padStart(2,'0')}`;
}
