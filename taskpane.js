let timers = [];
let editingId = null;

// INIT
document.addEventListener("DOMContentLoaded", () => {
    loadTimers((data) => {
        timers = data;
        renderList();
    });

    document.getElementById("btn-add").onclick = () => {
        editingId = null;
        clearForm();
        showForm();
    };

    document.getElementById("cancel").onclick = hideForm;
    document.getElementById("save").onclick = saveTimer;
});

// ===============================
// STORAGE (PPT + WEB)
// ===============================

function saveTimers() {
    if (typeof Office !== "undefined" && Office.context?.document) {
        Office.context.document.settings.set("timers", timers);

        Office.context.document.settings.saveAsync((res) => {
            if (res.status === Office.AsyncResultStatus.Succeeded) {
                console.log("Salvo no PowerPoint");
            } else {
                console.error(res.error);
            }
        });
    } else {
        localStorage.setItem("timers", JSON.stringify(timers));
    }
}

function loadTimers(callback) {
    if (typeof Office !== "undefined" && Office.context?.document) {
        const data = Office.context.document.settings.get("timers");
        callback(data || []);
    } else {
        const data = JSON.parse(localStorage.getItem("timers") || "[]");
        callback(data);
    }
}

// ===============================
// UI
// ===============================

function renderList() {
    const list = document.getElementById("list-view");

    if (timers.length === 0) {
        list.innerHTML = "<p>Nenhum timer</p>";
        return;
    }

    list.innerHTML = timers.map(t => `
        <div class="timer-item">
            <b>${t.startSlide} → ${t.endSlide}</b> | ${t.duration}s

            <div class="row">
                <button onclick="insertTimer()">Inserir</button>
                <button onclick="editTimer('${t.id}')">Editar</button>
                <button onclick="deleteTimer('${t.id}')">Excluir</button>
            </div>
        </div>
    `).join("");
}

// FORM
function showForm() {
    document.getElementById("form-view").classList.remove("hidden");
}

function hideForm() {
    document.getElementById("form-view").classList.add("hidden");
}

function clearForm() {
    start.value = "";
    end.value = "";
    duration.value = "";
    color.value = "#000000";
    size.value = "60";
    jump.value = "0";
}

// ===============================
// CRUD
// ===============================

function saveTimer() {
    const status = document.getElementById("status");

    const newTimer = {
        id: editingId || Date.now().toString(),
        startSlide: parseInt(start.value),
        endSlide: parseInt(end.value),
        duration: parseInt(duration.value),
        color: color.value,
        size: parseInt(size.value),
        jumpTarget: parseInt(jump.value)
    };

    if (newTimer.startSlide > newTimer.endSlide) {
        status.innerText = "❌ Intervalo inválido";
        return;
    }

    const conflict = timers.find(t =>
        t.id !== editingId &&
        newTimer.startSlide <= t.endSlide &&
        newTimer.endSlide >= t.startSlide
    );

    if (conflict) {
        status.innerText = `❌ Conflito com ${conflict.startSlide}-${conflict.endSlide}`;
        return;
    }

    if (editingId) {
        timers = timers.map(t => t.id === editingId ? newTimer : t);
    } else {
        timers.push(newTimer);
    }

    saveTimers();
    hideForm();
    renderList();

    status.innerText = "✅ Salvo";
}

// EDITAR
function editTimer(id) {
    const t = timers.find(x => x.id === id);
    if (!t) return;

    editingId = id;

    start.value = t.startSlide;
    end.value = t.endSlide;
    duration.value = t.duration;
    color.value = t.color;
    size.value = t.size;
    jump.value = t.jumpTarget;

    showForm();
}

// EXCLUIR
function deleteTimer(id) {
    timers = timers.filter(t => t.id !== id);
    saveTimers();
    renderList();
}

// ===============================
// INSERIR NO SLIDE
// ===============================

function insertTimer() {
    if (typeof Office === "undefined") {
        console.log("Somente no PowerPoint");
        return;
    }

    Office.context.document.setSelectedDataAsync(
        `<iframe src="https://pedromlelis.github.io/Timer/timer.html" width="300" height="150" frameborder="0"></iframe>`,
        { coercionType: Office.CoercionType.Html }
    );
}
