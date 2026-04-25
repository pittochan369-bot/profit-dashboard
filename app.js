const sheetUrl =
  "https://docs.google.com/spreadsheets/d/1twuDfpuWtp4PEi847J1m6pnhXiEkUhwuxEQzUdLIu2o/edit?gid=1238786803#gid=1238786803";

const elements = {
  status: document.querySelector("#status"),
  sheetUpdatedAt: document.querySelector("#sheetUpdatedAt"),
  todayPeriod: document.querySelector("#todayPeriod"),
  monthPeriod: document.querySelector("#monthPeriod"),
  todayTarget: document.querySelector("#todayTarget"),
  todayProfit: document.querySelector("#todayProfit"),
  todayRate: document.querySelector("#todayRate"),
  todayRateBar: document.querySelector("#todayRateBar"),
  todayTargetBar: document.querySelector("#todayTargetBar"),
  todayProfitBar: document.querySelector("#todayProfitBar"),
  todayRemaining: document.querySelector("#todayRemaining"),
  monthTarget: document.querySelector("#monthTarget"),
  monthTargetBar: document.querySelector("#monthTargetBar"),
  monthProfit: document.querySelector("#monthProfit"),
  monthProfitBar: document.querySelector("#monthProfitBar"),
  monthRate: document.querySelector("#monthRate"),
  monthRateBar: document.querySelector("#monthRateBar"),
  monthProgress: document.querySelector("#monthProgress"),
  monthProgressBar: document.querySelector("#monthProgressBar"),
  monthShortage: document.querySelector("#monthShortage"),
  monthRemaining: document.querySelector("#monthRemaining"),
  todayStoreCount: document.querySelector("#todayStoreCount"),
  monthStoreCount: document.querySelector("#monthStoreCount"),
  todayStoreTable: document.querySelector("#todayStoreTable"),
  monthStoreTable: document.querySelector("#monthStoreTable"),
  todayStoreCards: document.querySelector("#todayStoreCards"),
  monthStoreCards: document.querySelector("#monthStoreCards"),
  refreshButton: document.querySelector("#refreshButton"),
};

function sheetUrlToCsvUrl(url) {
  const parsed = new URL(url);
  const id = parsed.pathname.match(/\/spreadsheets\/d\/([^/]+)/)?.[1];
  const gid = parsed.searchParams.get("gid") || parsed.hash.match(/gid=(\d+)/)?.[1] || "0";

  if (!id) {
    throw new Error("スプレッドシートIDをURLから取得できませんでした。");
  }

  return `https://docs.google.com/spreadsheets/d/${id}/export?format=csv&gid=${gid}`;
}

function parseCsv(text) {
  const rows = [];
  let row = [];
  let cell = "";
  let quoted = false;

  for (let index = 0; index < text.length; index += 1) {
    const char = text[index];
    const next = text[index + 1];

    if (char === '"' && quoted && next === '"') {
      cell += '"';
      index += 1;
    } else if (char === '"') {
      quoted = !quoted;
    } else if (char === "," && !quoted) {
      row.push(cell.trim());
      cell = "";
    } else if ((char === "\n" || char === "\r") && !quoted) {
      if (char === "\r" && next === "\n") index += 1;
      row.push(cell.trim());
      if (row.some(Boolean)) rows.push(row);
      row = [];
      cell = "";
    } else {
      cell += char;
    }
  }

  row.push(cell.trim());
  if (row.some(Boolean)) rows.push(row);

  return rows;
}

function normalize(text) {
  return String(text ?? "")
    .replace(/\s/g, "")
    .toLowerCase();
}

function toNumber(value) {
  const cleaned = String(value ?? "").replace(/[^\d.-]/g, "");
  return cleaned === "" ? null : Number(cleaned);
}

function formatNumber(value) {
  const number = toNumber(value);
  if (number === null || Number.isNaN(number)) return "--";
  return new Intl.NumberFormat("ja-JP").format(number);
}

function formatPercent(value) {
  const number = toNumber(value);
  if (number === null || Number.isNaN(number)) return "--";
  return `${Math.round(number)}%`;
}

function progressWidth(value) {
  const number = toNumber(value);
  if (number === null || Number.isNaN(number)) return "0%";
  return `${Math.max(0, Math.min(number, 100))}%`;
}

function ratioWidth(current, total) {
  const currentNumber = toNumber(current);
  const totalNumber = toNumber(total);
  if (!totalNumber || Number.isNaN(currentNumber) || Number.isNaN(totalNumber)) return "0%";
  return `${Math.max(0, Math.min((currentNumber / totalNumber) * 100, 100))}%`;
}

function getCell(rows, label, valueIndex = 1) {
  const target = normalize(label);
  const row = rows.find((currentRow) => normalize(currentRow[0]) === target);
  return row?.[valueIndex] ?? "";
}

function getPeriod(rows, marker) {
  const row = rows.find((currentRow) => normalize(currentRow[0]).startsWith(normalize(marker)));
  return row?.[0]?.replace(/^■\s*/, "") ?? "--";
}

function getKeyValueBlock(rows, marker) {
  const startIndex = rows.findIndex((row) => normalize(row[0]).startsWith(normalize(marker)));
  if (startIndex < 0) return [];

  const block = [];
  for (let index = startIndex + 1; index < rows.length; index += 1) {
    const row = rows[index];
    const firstCell = normalize(row[0]);

    if (firstCell.startsWith("店舗別") || firstCell.startsWith("■")) {
      if (block.length) break;
      continue;
    }

    if (row[0] && row[0] !== "項目") block.push(row);
  }

  return block;
}

function getTableBlock(rows, marker) {
  const startIndex = rows.findIndex((row) => normalize(row[0]).startsWith(normalize(marker)));
  if (startIndex < 0) return { headers: [], rows: [] };

  const headers = rows[startIndex + 1] ?? [];
  const body = [];
  for (let index = startIndex + 2; index < rows.length; index += 1) {
    const row = rows[index];
    const firstCell = normalize(row[0]);
    if (firstCell.startsWith("■") || firstCell.startsWith("店舗別")) break;
    if (row.some(Boolean)) body.push(row);
  }

  return { headers, rows: body };
}

function setText(element, value, formatter = (currentValue) => currentValue || "--") {
  if (!element) return;
  element.textContent = formatter(value);
}

function renderMetrics(rows) {
  const today = getKeyValueBlock(rows, "■ 本日");
  const month = getKeyValueBlock(rows, "■ 今月");

  if (elements.todayPeriod) elements.todayPeriod.textContent = getPeriod(rows, "■ 本日");
  if (elements.monthPeriod) elements.monthPeriod.textContent = getPeriod(rows, "■ 今月");

  setText(elements.todayTarget, getCell(today, "今日の粗利目標"), formatNumber);
  setText(elements.todayProfit, getCell(today, "現在の1日粗利"), formatNumber);
  setText(elements.todayRate, getCell(today, "達成率(%)"), formatPercent);
  setText(elements.todayRemaining, getCell(today, "目標残額"), formatNumber);
  if (elements.todayTargetBar) elements.todayTargetBar.style.width = "100%";
  if (elements.todayProfitBar) {
    elements.todayProfitBar.style.width = ratioWidth(
      getCell(today, "現在の1日粗利"),
      getCell(today, "今日の粗利目標"),
    );
  }
  if (elements.todayRateBar) elements.todayRateBar.style.width = progressWidth(getCell(today, "達成率(%)"));
  setText(elements.monthTarget, getCell(month, "今月の粗利目標"), formatNumber);
  setText(elements.monthProfit, getCell(month, "現在の月粗利"), formatNumber);
  setText(elements.monthRate, getCell(month, "達成率(%)"), formatPercent);
  setText(elements.monthProgress, getCell(month, "進捗率(%)"), formatPercent);
  setText(elements.monthShortage, getCell(month, "進捗不足粗利"), formatNumber);
  setText(elements.monthRemaining, getCell(month, "目標残額"), formatNumber);
  if (elements.monthTargetBar) elements.monthTargetBar.style.width = "100%";
  if (elements.monthProfitBar) {
    elements.monthProfitBar.style.width = ratioWidth(
      getCell(month, "現在の月粗利"),
      getCell(month, "今月の粗利目標"),
    );
  }
  if (elements.monthRateBar) elements.monthRateBar.style.width = progressWidth(getCell(month, "達成率(%)"));
  if (elements.monthProgressBar) {
    elements.monthProgressBar.style.width = progressWidth(getCell(month, "進捗率(%)"));
  }
}

function formatTableCell(header, value) {
  if (value === "") return "";
  if (header.includes("%") || header.includes("率")) return formatPercent(value);
  if (header.includes("店舗")) return value;
  return formatNumber(value);
}

function compactHeader(header) {
  const replacements = {
    "達成率(%)": "達成",
    "進捗率(%)": "進捗",
    "進捗不足粗利": "不足",
    "目標残額": "残額",
    "1日目標": "目標",
  };

  return replacements[header] || header;
}

function renderTable(table, tableData) {
  if (!table) return;
  const { headers, rows } = tableData;
  table.tHead.innerHTML = "";
  table.tBodies[0].innerHTML = "";

  const headerRow = table.tHead.insertRow();
  headers.forEach((header) => {
    const th = document.createElement("th");
    th.textContent = compactHeader(header);
    th.title = header;
    headerRow.append(th);
  });

  rows.forEach((row) => {
    const tr = table.tBodies[0].insertRow();
    headers.forEach((header, index) => {
      const td = tr.insertCell();
      const value = row[index] || "";
      td.textContent = formatTableCell(header, value);

      const numericValue = toNumber(value);
      if (header.includes("不足") || header.includes("残額")) {
        td.classList.toggle("positive", numericValue < 0);
        td.classList.toggle("negative", numericValue > 0);
      }
      if (header.includes("達成率") || header.includes("進捗率")) {
        td.classList.toggle("positive", numericValue >= 100);
      }
    });
  });
}

function getBadgeClass(value, header = "") {
  const numericValue = toNumber(value);
  if (header.includes("率")) return numericValue >= 100 ? "good" : "warn";
  if (header.includes("不足") || header.includes("残額")) return numericValue <= 0 ? "good" : "warn";
  return "";
}

function renderStoreCards(container, tableData, mode) {
  if (!container) return;
  container.innerHTML = "";
  const { headers, rows } = tableData;
  rows.forEach((row) => {
    const storeName = row[0] || "--";
    const badgeIndex = headers.findIndex((header) => header.includes("達成率"));
    const badgeValue = badgeIndex >= 0 ? formatTableCell(headers[badgeIndex], row[badgeIndex]) : "--";
    const badgeClass = badgeIndex >= 0 ? getBadgeClass(row[badgeIndex], headers[badgeIndex]) : "";
    const detailHeaders =
      mode === "today"
        ? ["売上", "粗利", "1日目標", "目標残額"]
        : ["粗利", "目標", "進捗率(%)", "進捗不足粗利"];

    const detailMarkup = detailHeaders
      .filter((header) => headers.includes(header))
      .map((header) => {
        const index = headers.indexOf(header);
        const value = formatTableCell(header, row[index]);
        const tone = getBadgeClass(row[index], header);
        return `<div><span class="store-stat-label">${compactHeader(header)}</span><div class="store-stat-value ${tone === "warn" ? "negative" : tone === "good" ? "positive" : ""}">${value}</div></div>`;
      })
      .join("");

    container.insertAdjacentHTML(
      "beforeend",
      `<article class="store-card">
        <div class="store-card-head">
          <h3>${storeName}</h3>
          <span class="store-badge ${badgeClass}">${badgeValue}</span>
        </div>
        <div class="store-stats">${detailMarkup}</div>
      </article>`,
    );
  });
}

function setStatus(message, isError = false) {
  elements.status.textContent = message;
  elements.status.classList.toggle("error", isError);
}

async function loadSheet() {
  try {
    setStatus("スプレッドシートから取得中...");
    const response = await fetch(sheetUrlToCsvUrl(sheetUrl), { cache: "no-store" });

    if (!response.ok) {
      throw new Error(`HTTP ${response.status}`);
    }

    const csv = await response.text();
    const rows = parseCsv(csv);

    if (rows.length === 0) {
      throw new Error("CSVに表示できる行がありません。");
    }

    renderMetrics(rows);
    const todayStores = getTableBlock(rows, "店舗別(本日)");
    const monthStores = getTableBlock(rows, "店舗別(今月)");
    renderTable(elements.todayStoreTable, todayStores);
    renderTable(elements.monthStoreTable, monthStores);
    renderStoreCards(elements.todayStoreCards, todayStores, "today");
    renderStoreCards(elements.monthStoreCards, monthStores, "month");
    if (elements.todayStoreCount) elements.todayStoreCount.textContent = `${todayStores.rows.length}店舗`;
    if (elements.monthStoreCount) elements.monthStoreCount.textContent = `${monthStores.rows.length}店舗`;
    if (elements.sheetUpdatedAt) elements.sheetUpdatedAt.textContent = getCell(rows, "最終更新日時") || "--";
    setStatus("取得しました。");
  } catch (error) {
    setStatus(
      `取得できませんでした。シートを「リンクを知っている全員が閲覧可」にするか、公開CSV/バックエンド経由に切り替えてください。詳細: ${error.message}`,
      true,
    );
  }
}

elements.refreshButton.addEventListener("click", loadSheet);
loadSheet();
