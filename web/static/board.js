"use strict";

function statusClass(status) {
  if (status.includes("접안") && !status.includes("미접안")) return "status-docked";
  if (status.includes("수배") || status.includes("미접안")) return "status-waiting";
  return "status-alert";
}

function tick() {
  const now = new Date();
  const pad = (n) => String(n).padStart(2, "0");
  document.getElementById("clock").textContent =
    `${pad(now.getHours())}:${pad(now.getMinutes())}:${pad(now.getSeconds())}`;
}

async function refresh() {
  try {
    const res = await fetch("/api/records?board=true");
    const data = await res.json();
    const body = document.getElementById("board-body");
    body.innerHTML = "";
    (data.records || []).forEach((rec) => {
      const tr = document.createElement("tr");
      const cells = [
        { text: `${rec.emoji} ${rec.status}`, cls: statusClass(rec.status) },
        { text: rec.delivery_type },
        { text: rec.camp },
        { text: rec.dock },
        { text: rec.hocha },
        { text: rec.car_num },
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
    document
      .getElementById("board-empty")
      .classList.toggle("hidden", (data.records || []).length > 0);
  } catch (_e) {
    /* 네트워크 오류 시 다음 주기에 재시도 */
  }
}

tick();
refresh();
setInterval(tick, 1000);
setInterval(refresh, 5000);
