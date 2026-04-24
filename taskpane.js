console.log("TASKPANE MULTI TIMER");

let timers = [];

// 🔄 INIT
document.addEventListener("DOMContentLoaded", () => {
    loadTimers();
    renderList();

    document.getElementById("btn-add").onclick = showForm;
    document.getElementById("cancel").onclick = hideForm;
    document.getElementById("save").onclick = saveTimer;
});

// 💾 LOAD
function loadTimers() {
    timers = JSON.parse(localStorage.getItem("timers") || "[]");
}

// 💾 SAVE
function persistTimers() {
    localStorage.setItem("timers", JSON.stringify(timers));
}

// 📋 LISTA
function renderList() {
    const list = document.getElementById("list-view");

    if (timers.length === 0) {
        list.innerHTML = "<p>Nenhum timer criado</p>";
        return;
    }

    list.innerHTML = timers.map(t => `
        <div class="timer-item">
            Slides ${t.startSlide} - ${t.endSlide}<br>
            ${t.duration}s
        </div>
    `).join("");
}

// 👁️ FORM
function showForm() {
    document.getElementById("form-view").classList.remove("hidden");
}

function hideForm() {
    document.getElementById("form-view").classList.add("hidden");
}

// 🚨 VALIDAÇÃO DE SOBREPOSIÇÃO
function hasOverlap(newTimer) {
    return timers.some(t =>
        newTimer.startSlide <= t.endSlide &&
        newTimer.endSlide >= t.startSlide
    );
}

// 💾 SALVAR TIMER
function saveTimer() {
    const status = document.getElementById("status");

    const newTimer = {
        id: Date.now().toString(),
        startSlide: parseInt(document.getElementById("start").value),
        endSlide: parseInt(document.getElementById("end").value),
        duration: parseInt(document.getElementById("duration").value),
        color: document.getElementById("color").value,
        size: parseInt(document.getElementById("size").value),
        jumpTarget: parseInt(document.getElementById("jump").value)
    };

    // 🚨 validação básica
    if (newTimer.startSlide > newTimer.endSlide) {
        status.innerText = "❌ Slide inicial maior que final";
        status.style.color = "red";
        return;
    }

    // 🚨 validação de colisão
    const conflict = timers.find(t =>
        newTimer.startSlide <= t.endSlide &&
        newTimer.endSlide >= t.startSlide
    );

    if (conflict) {
        status.innerText =
            `❌ Conflito com intervalo ${conflict.startSlide}-${conflict.endSlide}`;
        status.style.color = "red";
        return;
    }

    // ✅ salvar
    timers.push(newTimer);
    persistTimers();

    status.innerText = "✅ Timer criado!";
    status.style.color = "green";

    hideForm();
    renderList();
}
