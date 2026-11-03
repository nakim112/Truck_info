"use strict";

const state = {
  config: null,
  records: [],
  selected: new Set(),
};

const $ = (sel) => document.querySelector(sel);

async function api(path, options) {
  const res = await fetch(path, {
    headers: { "Content-Type": "application/json" },
    ...options,
  });
  let body = null;
  try {
    body = await res.json();
  } catch (_e) {
    body = null;
  }
  if (!res.ok) {
    const msg = (body && body.detail) || `요청 실패 (${res.status})`;
    throw new Error(msg);
  }
  return body;
}

function fillSelect(el, items, { keepValue = false } = {}) {
  const prev = el.value;
  el.innerHTML = "";
  items.forEach((it) => {
    const opt = document.createElement("option");
    opt.value = it;
    opt.textContent = it;
    el.appendChild(opt);
  });
  if (keepValue && items.includes(prev)) el.value = prev;
}

function fillDatalist(el, items) {
  el.innerHTML = "";
  items.forEach((it) => {
    const opt = document.createElement("option");
    opt.value = it;
    el.appendChild(opt);
  });
}

function updateCampOptions() {
  const type = $("#type").value;
  const camps = (state.config.camps_by_type[type] || []).slice();
  fillDatalist($("#camp-options"), camps);
}

function setMessage(text, kind) {
  const el = $("#form-msg");
  el.textContent = text || "";
  el.className = "form-msg" + (kind ? " " + kind : "");
}

function renderRows() {
  const body = $("#grid-body");
  body.innerHTML = "";
  state.records.forEach((rec) => {
    const tr = document.createElement("tr");
    tr.dataset.id = rec.id;
    if (state.selected.has(rec.id)) tr.classList.add("selected");

    const sel = document.createElement("td");
    sel.className = "col-sel";
    const cb = document.createElement("input");
    cb.type = "checkbox";
    cb.checked = state.selected.has(rec.id);
    cb.addEventListener("change", () => {
      if (cb.checked) state.selected.add(rec.id);
      else state.selected.delete(rec.id);
      tr.classList.toggle("selected", cb.checked);
      refreshToolbarState();
    });
    sel.appendChild(cb);
    tr.appendChild(sel);

    const cells = [
      { text: `${rec.emoji} ${rec.status}`, cls: "status-cell" },
      { text: rec.delivery_type },
      { text: rec.camp },
      { text: rec.dock },
      { text: rec.hocha },
      { text: rec.car_num },
      { text: rec.phone },
      { text: rec.time },
    ];
    cells.forEach((c) => {
      const td = document.createElement("td");
      td.textContent = c.text;
      if (c.cls) td.className = c.cls;
      tr.appendChild(td);
    });

    body.appendChild(tr);
  });

  $("#count").textContent = `${state.records.length}건`;
  $("#empty-msg").classList.toggle("hidden", state.records.length > 0);
  $("#select-all").checked =
    state.records.length > 0 && state.selected.size === state.records.length;
  refreshToolbarState();
}

function refreshToolbarState() {
  const has = state.selected.size > 0;
  $("#apply-status").disabled = !has;
  $("#delete-selected").disabled = !has;
}

async function loadRecords() {
  const data = await api("/api/records");
  state.records = data.records;
  const ids = new Set(state.records.map((r) => r.id));
  state.selected = new Set([...state.selected].filter((id) => ids.has(id)));
  renderRows();
}

async function onSubmit(ev) {
  ev.preventDefault();
  setMessage("", null);
  const payload = {
    delivery_type: $("#type").value,
    camp: $("#camp").value,
    dock: $("#dock").value,
    hocha: $("#hocha").value,
    car_num: $("#car_num").value,
    phone: $("#phone").value,
    time: $("#time").value,
  };
  try {
    const rec = await api("/api/records", {
      method: "POST",
      body: JSON.stringify(payload),
    });
    setMessage(`등록 완료: ${rec.delivery_type} / ${rec.camp} / ${rec.car_num}`, "ok");
    ["car_num", "phone", "time"].forEach((id) => ($("#" + id).value = ""));
    $("#car_num").focus();
    await loadRecords();
  } catch (err) {
    setMessage(err.message, "error");
  }
}

async function applyStatus() {
  const status = $("#status-select").value;
  const ids = [...state.selected];
  try {
    for (const id of ids) {
      await api(`/api/records/${id}/status`, {
        method: "PATCH",
        body: JSON.stringify({ status }),
      });
    }
    await loadRecords();
  } catch (err) {
    setMessage(err.message, "error");
  }
}

async function deleteSelected() {
  const ids = [...state.selected];
  if (ids.length === 0) return;
  if (!confirm(`선택한 ${ids.length}건을 삭제하시겠습니까?`)) return;
  try {
    for (const id of ids) {
      await api(`/api/records/${id}`, { method: "DELETE" });
    }
    state.selected.clear();
    await loadRecords();
  } catch (err) {
    setMessage(err.message, "error");
  }
}

async function clearAll() {
  if (!confirm("전체 차량 데이터를 삭제하시겠습니까?")) return;
  try {
    await api("/api/records", { method: "DELETE" });
    state.selected.clear();
    await loadRecords();
  } catch (err) {
    setMessage(err.message, "error");
  }
}

function onSelectAll(ev) {
  if (ev.target.checked) {
    state.records.forEach((r) => state.selected.add(r.id));
  } else {
    state.selected.clear();
  }
  renderRows();
}

async function init() {
  state.config = await api("/api/config");
  fillSelect($("#type"), state.config.delivery_types);
  fillSelect($("#status-select"), state.config.statuses);
  fillDatalist($("#dock-options"), state.config.docks);
  fillDatalist($("#hocha-options"), state.config.hochas);
  updateCampOptions();

  // 포커스 시 내용을 전체 선택해, 유지된 값 위에 바로 덮어쓸 수 있게 합니다.
  document.querySelectorAll(".record-form input").forEach((inp) => {
    inp.addEventListener("focus", () => inp.select());
  });

  $("#type").addEventListener("change", updateCampOptions);
  $("#record-form").addEventListener("submit", onSubmit);
  $("#apply-status").addEventListener("click", applyStatus);
  $("#delete-selected").addEventListener("click", deleteSelected);
  $("#clear-all").addEventListener("click", clearAll);
  $("#refresh").addEventListener("click", loadRecords);
  $("#select-all").addEventListener("change", onSelectAll);

  await loadRecords();
}

init().catch((err) => setMessage("초기화 실패: " + err.message, "error"));
