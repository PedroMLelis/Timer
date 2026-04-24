let timers = [];
let endTime = null;

// CARREGAR CONFIG
function loadTimers() {
    if (typeof Office !== "undefined" && Office.context?.document) {
        timers = Office.context.document.settings.get("timers") || [];
    } else {
        timers = JSON.parse(localStorage.getItem("timers") || "[]");
    }
}

// LOOP
setInterval(() => {

    if (typeof Office === "undefined") return;

    loadTimers();

    Office.context.document.getSelectedDataAsync(
        Office.CoercionType.SlideRange,
        (res) => {
            if (res.status !== Office.AsyncResultStatus.Succeeded) return;

            const slide = res.value.slides[0].index;

            const active = timers.find(t =>
                slide >= t.startSlide && slide <= t.endSlide
            );

            if (!active) {
                document.getElementById("timer").innerText = "";
                endTime = null;
                return;
            }

            if (!endTime) {
                endTime = Date.now() + active.duration * 1000;
            }

            const ms = Math.max(0, endTime - Date.now());

            const sec = Math.ceil(ms / 1000);
            const m = Math.floor(sec / 60);
            const s = sec % 60;

            document.getElementById("timer").innerText =
                `${m.toString().padStart(2,'0')}:${s.toString().padStart(2,'0')}`;
        }
    );

}, 1000);
