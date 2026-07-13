/**
 * 醫院自動排班系統 - 核心邏輯 (Pure JavaScript)
 * 包含：Excel 解析、模擬退火自動排班演算法、即時違規驗證、互動編輯與 Excel 匯出。
 */

// ==========================================================================
// 全域狀態
// ==========================================================================
let employees = [];        // 員工列表資料
let totalDays = 30;        // 當月天數 (預設 30)
let weekdays = [];         // 當月每日星期 (三, 四, 五...)
let headers = [];          // Excel 原始表頭資訊
let fileName = "";         // 上傳的檔案名稱
let currentSchedule = [];  // 當前班表 (2D Array: employees x days)
let originalExcelData = null; // 原始 Excel 工作表資料，用於匯出時保持結構
let statisticsRows = [];      // 底部統計列資訊 (格式: { rowIndex, label })

// 設定參數
const config = {
  minOff: 8,
  maxOff: 9,
  maxConsecutive: 6,
  demand: {
    day: 5,
    evening: 5,
    night: 5
  }
};

// 全域放假判斷函數 (大小寫無關)
function isOff(val) {
  if (!val) return false;
  const v = String(val).toUpperCase();
  return (v === "OFF" || v === "VV" || v === "W" || v === "FF" || v === "V" || v === "AV" || v === "VP");
}

// ==========================================================================
// 頁面載入與事件初始化
// ==========================================================================
document.addEventListener("DOMContentLoaded", () => {
  initTheme();
  initUploadEvents();
  initParamEvents();
  initActionEvents();
});

// 主題切換 (亮色/暗色)
function initTheme() {
  const themeToggle = document.getElementById("theme-toggle");
  const currentTheme = localStorage.getItem("theme") || "light";
  document.documentElement.setAttribute("data-theme", currentTheme);
  updateThemeIcon(currentTheme);

  themeToggle.addEventListener("click", () => {
    const theme = document.documentElement.getAttribute("data-theme");
    const newTheme = theme === "dark" ? "light" : "dark";
    document.documentElement.setAttribute("data-theme", newTheme);
    localStorage.setItem("theme", newTheme);
    updateThemeIcon(newTheme);
  });
}

function updateThemeIcon(theme) {
  const icon = document.querySelector("#theme-toggle i");
  if (theme === "dark") {
    icon.className = "fa-solid fa-sun";
  } else {
    icon.className = "fa-solid fa-moon";
  }
}

// 檔案拖放上傳
function initUploadEvents() {
  const uploadZone = document.getElementById("upload-zone");
  const fileInput = document.getElementById("file-input");

  uploadZone.addEventListener("click", () => fileInput.click());

  uploadZone.addEventListener("dragover", (e) => {
    e.preventDefault();
    uploadZone.classList.add("dragover");
  });

  uploadZone.addEventListener("dragleave", () => {
    uploadZone.classList.remove("dragover");
  });

  uploadZone.addEventListener("drop", (e) => {
    e.preventDefault();
    uploadZone.classList.remove("dragover");
    if (e.dataTransfer.files.length > 0) {
      handleFile(e.dataTransfer.files[0]);
    }
  });

  fileInput.addEventListener("change", (e) => {
    if (e.target.files.length > 0) {
      handleFile(e.target.files[0]);
    }
  });
}

// 參數即時更新與防呆
function initParamEvents() {
  const minOffInput = document.getElementById("min-off");
  const maxOffInput = document.getElementById("max-off");
  const maxConsecutiveInput = document.getElementById("max-consecutive");
  const demandDayInput = document.getElementById("demand-day");
  const demandEveningInput = document.getElementById("demand-evening");
  const demandNightInput = document.getElementById("demand-night");

  const updateConfig = () => {
    config.minOff = parseInt(minOffInput.value) || 8;
    config.maxOff = parseInt(maxOffInput.value) || 9;
    config.maxConsecutive = parseInt(maxConsecutiveInput.value) || 6;
    config.demand.day = parseInt(demandDayInput.value) || 5;
    config.demand.evening = parseInt(demandEveningInput.value) || 5;
    config.demand.night = parseInt(demandNightInput.value) || 5;
    
    document.getElementById("off-days-label").textContent = `${config.minOff} - ${config.maxOff} 天`;
    
    if (employees.length > 0) {
      renderScheduleAndValidate();
    }
  };

  [minOffInput, maxOffInput, maxConsecutiveInput, demandDayInput, demandEveningInput, demandNightInput].forEach(input => {
    input.addEventListener("input", updateConfig);
  });
}

// 自動排班與匯出按鈕
function initActionEvents() {
  const btnSchedule = document.getElementById("btn-schedule");
  const btnExport = document.getElementById("btn-export");
  const modalClose = document.getElementById("modal-close");
  const modalOverlay = document.getElementById("violation-modal");

  btnSchedule.addEventListener("click", () => {
    if (employees.length === 0) return;
    showToast("開始自動排班，正在尋找最佳解...", "info");
    
    // 使用 setTimeout 讓 UI 有時間渲染 Loading
    btnSchedule.disabled = true;
    setTimeout(() => {
      runAutoScheduling();
      btnSchedule.disabled = false;
    }, 100);
  });

  btnExport.addEventListener("click", () => {
    if (employees.length === 0) return;
    exportToExcel();
  });

  modalClose.addEventListener("click", () => {
    modalOverlay.classList.remove("show");
  });

  modalOverlay.addEventListener("click", (e) => {
    if (e.target === modalOverlay) {
      modalOverlay.classList.remove("show");
    }
  });
}

// ==========================================================================
// Excel 讀取與解析
// ==========================================================================
function handleFile(file) {
  fileName = file.name;
  document.getElementById("file-name").textContent = fileName;
  document.getElementById("file-info").classList.add("active");

  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: "array" });
      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];
      originalExcelData = sheet;
      
      const rawData = XLSX.utils.sheet_to_json(sheet, { header: 1 });
      parseExcelData(rawData);
      
      showToast("Excel 讀取成功！已解析員工名單與班表。", "success");
      
      // 啟用按鈕
      document.getElementById("btn-schedule").disabled = false;
      document.getElementById("btn-export").disabled = false;
      document.getElementById("empty-state").style.display = "none";
      document.getElementById("schedule-section").style.display = "flex";
    } catch (err) {
      console.error(err);
      showToast("讀取失敗，請確認檔案格式是否正確！", "error");
    }
  };
  reader.readAsArrayBuffer(file);
}

function parseExcelData(rawData) {
  employees = [];
  weekdays = [];
  headers = [];
  
  let headerRowIndex = -1;
  let weekdayRowIndex = -1;
  
  // 1. 尋找表頭日期列 (尋找包含數字 1 到 28 的列，且通常有 進階、職級、姓名、班別 欄位)
  for (let r = 0; r < rawData.length; r++) {
    const row = rawData[r];
    if (!row) continue;
    
    // 檢查是否有連續日期 (尋找是否有 1, 2, 3...)
    let consecutiveDates = 0;
    for (let c = 4; c < row.length; c++) {
      if (parseInt(row[c]) === consecutiveDates + 1) {
        consecutiveDates++;
      }
    }
    
    if (consecutiveDates >= 28) {
      headerRowIndex = r;
      totalDays = consecutiveDates; // 28, 29, 30 或 31
      break;
    }
  }
  
  if (headerRowIndex === -1) {
    throw new Error("找不到日期標頭列 (需要包含連續 1 到 28 的數字)");
  }
  
  // 2. 取得星期列 (通常在日期列的下一列)
  weekdayRowIndex = headerRowIndex + 1;
  const headerRow = rawData[headerRowIndex];
  const weekdayRow = rawData[weekdayRowIndex];
  
  for (let d = 0; d < totalDays; d++) {
    weekdays.push(weekdayRow[4 + d] || "");
  }
  
  // 更新月份徽章
  document.getElementById("month-badge").textContent = `當月共 ${totalDays} 天`;
  
  // 3. 解析員工列 (從星期列下方開始)
  const isIgnoredFirstDayText = (val) => {
    if (!val) return false;
    const str = String(val).trim();
    // 匹配如: 1大, 3大, 6白, 白1 等
    return /^[0-9]+[大小白休]|[大小白休]+[0-9]+|[0-9]+$/.test(str);
  };
  
  statisticsRows = [];
  for (let r = weekdayRowIndex + 1; r < rawData.length; r++) {
    const row = rawData[r];
    if (!row) continue;
    
    const level = String(row[0] || "").trim();
    const id = String(row[1] || "").trim();
    const name = String(row[2] || "").trim();
    const prefStr = String(row[3] || "").trim();
    
    // 必須有姓名與編號才算是有效員工列
    if (!id || !name) {
      // 檢查是否為底部統計行 (前 4 欄是否包含 "12", "8", "4", "85" 等班別標記)
      const cells = [level, id, name, prefStr];
      const matchLabel = cells.find(val => ["12", "8", "4", "85"].includes(val));
      if (matchLabel) {
        statisticsRows.push({
          rowIndex: r,
          label: matchLabel
        });
      }
      continue;
    }
    
    // 解析班別偏好
    let preferences = [];
    if (prefStr) {
      // 拆分偏好為字元數組，例如 "白大" -> ["白", "大"]
      preferences = Array.from(prefStr).filter(char => ["白", "小", "大", "休"].includes(char));
    }
    
    // 是否為護理長 (N4 且班別欄位為空白)
    const isHeadNurse = (level === "N4" && preferences.length === 0);
    
    // 解析排班歷程 (日期 1 到 totalDays)
    const originalSchedule = [];
    for (let d = 0; d < totalDays; d++) {
      let val = row[4 + d];
      if (val !== undefined && val !== null) {
        val = String(val).trim();
        
        // 規則 1：忽略第一天出現的 "1大"、"3大" 等文字，視為可重新排班 (null)
        if (d === 0 && isIgnoredFirstDayText(val)) {
          originalSchedule.push(null);
        } else {
          originalSchedule.push(val); // W, FF, SS, VV, V, AV, VP 等
        }
      } else {
        originalSchedule.push(null);
      }
    }
    
    employees.push({
      level,
      id,
      name,
      prefStr,
      preferences,
      isHeadNurse,
      originalSchedule,
      rowIndex: r // 記錄在原始 Excel 的列索引，方便匯出
    });
  }
  
  // 初始化目前班表為原始班表拷貝
  currentSchedule = employees.map(emp => [...emp.originalSchedule]);
  
  renderScheduleAndValidate();
}

// ==========================================================================
// 班表渲染與互動編輯
// ==========================================================================
function renderScheduleAndValidate() {
  const result = validateSchedule(currentSchedule);
  renderTable(result);
  renderStatsAndChart(result);
}

function renderTable(validationResult) {
  const table = document.getElementById("schedule-table");
  table.innerHTML = "";
  
  // --- 表頭第一列 (標題與日期 1~31) ---
  const tr1 = document.createElement("tr");
  
  const thLevel = document.createElement("th");
  thLevel.className = "sticky-col col-level";
  thLevel.textContent = "進階";
  tr1.appendChild(thLevel);
  
  const thId = document.createElement("th");
  thId.className = "sticky-col col-id";
  thId.textContent = "職級編號";
  tr1.appendChild(thId);
  
  const thName = document.createElement("th");
  thName.className = "sticky-col col-name";
  thName.textContent = "姓名";
  tr1.appendChild(thName);
  
  const thPref = document.createElement("th");
  thPref.className = "sticky-col col-pref";
  thPref.textContent = "班別";
  tr1.appendChild(thPref);
  
  // 日期列
  for (let d = 1; d <= totalDays; d++) {
    const thDay = document.createElement("th");
    thDay.textContent = d;
    // 檢查是否為週末 (日=7 或 6)
    const isWknd = isWeekend(d);
    if (isWknd) thDay.className = "weekend";
    tr1.appendChild(thDay);
  }
  
  // 右側統計列表頭
  const thWork = document.createElement("th");
  thWork.className = "sticky-right-col col-work-days";
  thWork.textContent = "班數";
  tr1.appendChild(thWork);
  
  const thNight = document.createElement("th");
  thNight.className = "sticky-right-col col-night-days";
  thNight.textContent = "夜班";
  tr1.appendChild(thNight);
  
  const thOff = document.createElement("th");
  thOff.className = "sticky-right-col col-off-days";
  thOff.textContent = "放假";
  tr1.appendChild(thOff);
  
  const thViolations = document.createElement("th");
  thViolations.className = "sticky-right-col col-violations";
  thViolations.textContent = "狀態";
  tr1.appendChild(thViolations);
  
  table.appendChild(tr1);
  
  // --- 表頭第二列 (星期 三、四、五...) ---
  const tr2 = document.createElement("tr");
  
  // 左側 4 個獨立的空白凍結 th
  const leftEmptyCols = ["col-level", "col-id", "col-name", "col-pref"];
  leftEmptyCols.forEach(cls => {
    const th = document.createElement("th");
    th.className = `sticky-col ${cls}`;
    tr2.appendChild(th);
  });
  
  for (let d = 1; d <= totalDays; d++) {
    const thWd = document.createElement("th");
    thWd.textContent = weekdays[d - 1] || "";
    if (isWeekend(d)) thWd.className = "weekend";
    tr2.appendChild(thWd);
  }
  
  // 右側 4 個獨立的空白凍結 th (RWD/AWD 修正：不可使用 colSpan=4 否則會產生大區塊遮擋)
  const rightEmptyCols = ["col-work-days", "col-night-days", "col-off-days", "col-violations"];
  rightEmptyCols.forEach(cls => {
    const th = document.createElement("th");
    th.className = `sticky-right-col ${cls}`;
    tr2.appendChild(th);
  });
  
  table.appendChild(tr2);

  // --- 員工排班資料行 ---
  employees.forEach((emp, empIdx) => {
    const tr = document.createElement("tr");
    
    // 進階
    const tdLvl = document.createElement("td");
    tdLvl.className = "sticky-col col-level";
    tdLvl.textContent = emp.level;
    tr.appendChild(tdLvl);
    
    // 職級編號
    const tdId = document.createElement("td");
    tdId.className = "sticky-col col-id";
    tdId.textContent = emp.id;
    tr.appendChild(tdId);
    
    // 姓名
    const tdName = document.createElement("td");
    tdName.className = "sticky-col col-name";
    tdName.textContent = emp.name;
    tr.appendChild(tdName);
    
    // 班別 (偏好)
    const tdPref = document.createElement("td");
    tdPref.className = "sticky-col col-pref";
    tdPref.textContent = emp.prefStr;
    tr.appendChild(tdPref);
    
    // 每日格子
    const empViolations = validationResult.violations.filter(v => v.employeeId === emp.id);
    
    for (let d = 1; d <= totalDays; d++) {
      const td = document.createElement("td");
      if (isWeekend(d)) td.className = "weekend";
      
      const val = currentSchedule[empIdx][d - 1];
      const isPrefilled = emp.originalSchedule[d - 1] !== null;
      
      const wrapper = document.createElement("div");
      wrapper.className = "cell-value";
      
      // 根據班別上色
      if (emp.isHeadNurse) {
        wrapper.className += " cell-head-nurse";
        wrapper.textContent = val || "";
      } else if (isPrefilled) {
        wrapper.className += " cell-prefilled";
        wrapper.textContent = val || "";
      } else {
        // 可編輯格子
        if (val === "8") {
          wrapper.className += " cell-day";
          wrapper.textContent = "8";
        } else if (val === "4") {
          wrapper.className += " cell-evening";
          wrapper.textContent = "4";
        } else if (val === "12") {
          wrapper.className += " cell-night";
          wrapper.textContent = "12";
        } else if (val === "OFF") {
          wrapper.className += " cell-off";
          wrapper.textContent = "OFF";
        } else {
          wrapper.textContent = "";
        }
        
        // 綁定編輯下拉選單
        td.className += " cell-dropdown";
        td.addEventListener("click", (e) => {
          e.stopPropagation();
          closeAllDropdowns();
          showCellDropdown(td, empIdx, d - 1);
        });
      }
      
      // 規則 4：檢查當前格子是否違規，若是，以亮紅色顯示並加 Tooltip
      const gridViol = empViolations.find(v => v.day === d);
      if (gridViol) {
        wrapper.className += " cell-violation";
        td.title = gridViol.reason;
      }
      
      td.appendChild(wrapper);
      tr.appendChild(td);
    }
    
    // 右側統計資料
    const stats = validationResult.employeeStats[emp.id] || { work: 0, night: 0, off: 0 };
    
    // 班數
    const tdWork = document.createElement("td");
    tdWork.className = "sticky-right-col col-work-days";
    tdWork.textContent = emp.isHeadNurse ? "-" : stats.work;
    tr.appendChild(tdWork);
    
    // 夜班
    const tdNight = document.createElement("td");
    tdNight.className = "sticky-right-col col-night-days";
    tdNight.textContent = emp.isHeadNurse ? "-" : stats.night;
    tr.appendChild(tdNight);
    
    // 放假
    const tdOff = document.createElement("td");
    tdOff.className = "sticky-right-col col-off-days";
    tdOff.textContent = emp.isHeadNurse ? "-" : stats.off;
    // 若休假天數違規，亮紅標記
    if (!emp.isHeadNurse && (stats.off < config.minOff || stats.off > config.maxOff)) {
      tdOff.style.color = "var(--color-off)";
      tdOff.style.fontWeight = "bold";
    }
    tr.appendChild(tdOff);
    
    // 狀態 / 違規統計
    const tdViol = document.createElement("td");
    tdViol.className = "sticky-right-col col-violations";
    
    if (emp.isHeadNurse) {
      tdViol.innerHTML = `<span style="color: var(--text-muted); font-size: 0.75rem;">護理長</span>`;
    } else if (empViolations.length > 0) {
      tdViol.innerHTML = `
        <span class="violation-badge" onclick="showEmployeeViolations('${emp.id}')">
          <i class="fa-solid fa-triangle-exclamation"></i> ${empViolations.length} 處違規
        </span>
      `;
    } else {
      tdViol.innerHTML = `<span style="color: #16a34a; font-weight: 600;"><i class="fa-solid fa-circle-check"></i> 正常</span>`;
    }
    tr.appendChild(tdViol);
    
    table.appendChild(tr);
  });
  
  // 點擊空白處關閉下拉選單
  document.addEventListener("click", closeAllDropdowns);
}

// 週末檢查輔助函數 (三、四、五...)
function isWeekend(dayIndex) {
  const wd = weekdays[dayIndex - 1];
  return wd === "六" || wd === "日" || wd === "六日" || wd === "日六";
}

// 關閉所有下拉選單
function closeAllDropdowns() {
  document.querySelectorAll(".dropdown-menu").forEach(el => el.remove());
}

// 顯示儲存格下拉編輯選單
function showCellDropdown(containerTd, empIdx, dayIdx) {
  const dropdown = document.createElement("div");
  dropdown.className = "dropdown-menu show";
  
  const options = [
    { text: "白 (8)", val: "8", color: "var(--bg-day)", txtColor: "var(--color-day)" },
    { text: "小 (4)", val: "4", color: "var(--bg-evening)", txtColor: "var(--color-evening)" },
    { text: "大 (12)", val: "12", color: "var(--bg-night)", txtColor: "var(--color-night)" },
    { text: "放 (OFF)", val: "OFF", color: "var(--bg-off)", txtColor: "var(--color-off)" },
    { text: "清除", val: null, color: "var(--bg-primary)", txtColor: "var(--text-secondary)" }
  ];
  
  options.forEach(opt => {
    const btn = document.createElement("button");
    btn.className = "dropdown-item";
    btn.textContent = opt.text;
    btn.style.backgroundColor = opt.color;
    btn.style.color = opt.txtColor;
    
    btn.addEventListener("click", (e) => {
      e.stopPropagation();
      currentSchedule[empIdx][dayIdx] = opt.val;
      dropdown.remove();
      renderScheduleAndValidate();
      showToast("班表已手動修改", "success", 1000);
    });
    
    dropdown.appendChild(btn);
  });
  
  containerTd.appendChild(dropdown);
}

// ==========================================================================
// 統計資訊與長條圖渲染
// ==========================================================================
function renderStatsAndChart(validationResult) {
  const totalDaysCount = totalDays;
  
  let dayMatch = 0;
  let eveningMatch = 0;
  let nightMatch = 0;
  
  for (let d = 1; d <= totalDaysCount; d++) {
    const stats = validationResult.dailyStats[d] || { day: 0, evening: 0, night: 0 };
    if (stats.day === config.demand.day) dayMatch++;
    if (stats.evening === config.demand.evening) eveningMatch++;
    if (stats.night === config.demand.night) nightMatch++;
  }
  
  document.getElementById("stat-day").textContent = `${dayMatch} / ${totalDaysCount} 天達標`;
  document.getElementById("stat-evening").textContent = `${eveningMatch} / ${totalDaysCount} 天達標`;
  document.getElementById("stat-night").textContent = `${nightMatch} / ${totalDaysCount} 天達標`;
  
  const violationCount = validationResult.violations.length;
  const violationBox = document.getElementById("stat-violations");
  violationBox.textContent = `${violationCount} 處違規`;
  
  if (violationCount > 0) {
    violationBox.parentElement.parentElement.style.borderColor = "var(--color-off)";
    violationBox.style.color = "var(--color-off)";
    violationBox.parentElement.parentElement.style.cursor = "pointer";
    violationBox.parentElement.parentElement.onclick = () => showAllViolations(validationResult.violations);
  } else {
    violationBox.parentElement.parentElement.style.borderColor = "var(--border-color)";
    violationBox.style.color = "var(--text-primary)";
    violationBox.parentElement.parentElement.style.cursor = "default";
    violationBox.parentElement.parentElement.onclick = null;
  }
  
  // 渲染長條圖
  const chartContainer = document.getElementById("chart-container");
  chartContainer.innerHTML = "";
  
  for (let d = 1; d <= totalDaysCount; d++) {
    const stats = validationResult.dailyStats[d] || { day: 0, evening: 0, night: 0, off: 0 };
    const totalWorking = stats.day + stats.evening + stats.night;
    
    const wrapper = document.createElement("div");
    wrapper.className = "chart-bar-wrapper";
    
    // 計算高度比例 (最高以 20 人計)
    const maxHeight = 100; // %
    const heightDay = (stats.day / 20) * maxHeight;
    const heightEvening = (stats.evening / 20) * maxHeight;
    const heightNight = (stats.night / 20) * maxHeight;
    
    const barDay = document.createElement("div");
    barDay.className = "chart-bar";
    barDay.style.height = `${heightDay}%`;
    barDay.style.backgroundColor = "var(--color-day)";
    wrapper.appendChild(barDay);
    
    const barEvening = document.createElement("div");
    barEvening.className = "chart-bar";
    barEvening.style.height = `${heightEvening}%`;
    barEvening.style.backgroundColor = "var(--color-evening)";
    wrapper.appendChild(barEvening);
    
    const barNight = document.createElement("div");
    barNight.className = "chart-bar";
    barNight.style.height = `${heightNight}%`;
    barNight.style.backgroundColor = "var(--color-night)";
    wrapper.appendChild(barNight);
    
    const label = document.createElement("div");
    label.className = "chart-bar-label";
    label.textContent = d;
    wrapper.appendChild(label);
    
    // Tooltip
    const tooltip = document.createElement("div");
    tooltip.className = "chart-tooltip";
    tooltip.innerHTML = `
      <strong>${d} 號 (星期${weekdays[d-1] || ""})</strong><br>
      白班: ${stats.day}人 (目標 ${config.demand.day})<br>
      小夜: ${stats.evening}人 (目標 ${config.demand.evening})<br>
      大夜: ${stats.night}人 (目標 ${config.demand.night})<br>
      放假: ${stats.off}人
    `;
    wrapper.appendChild(tooltip);
    
    // 檢查這天的人數是否不符合目標各 5 人，若是，加框
    const demandSum = config.demand.day + config.demand.evening + config.demand.night;
    const currentSum = stats.day + stats.evening + stats.night;
    if (stats.day !== config.demand.day || stats.evening !== config.demand.evening || stats.night !== config.demand.night) {
      wrapper.style.boxShadow = "0 0 0 1px var(--color-off) inset";
    }
    
    chartContainer.appendChild(wrapper);
  }
}

// 彈出員工違規清單
function showEmployeeViolations(empId) {
  const emp = employees.find(e => e.id === empId);
  const result = validateSchedule(currentSchedule);
  const empViolations = result.violations.filter(v => v.employeeId === empId);
  
  const modal = document.getElementById("violation-modal");
  const modalBody = document.getElementById("modal-body");
  
  modalBody.innerHTML = `
    <h4 style="margin-bottom: 0.5rem; color: var(--text-primary);">人員：${emp.name} (${emp.level} - ${emp.id})</h4>
    <p style="margin-bottom: 1rem; color: var(--text-secondary);">共有 ${empViolations.length} 處違反規則：</p>
    <ul>
      ${empViolations.map(v => `
        <li style="margin-bottom: 0.5rem; line-height: 1.4;">
          <strong style="color: var(--color-off);">${v.day} 號：</strong> ${v.reason}
        </li>
      `).join("")}
    </ul>
  `;
  
  modal.classList.add("show");
}

// 彈出所有全域違規清單
function showAllViolations(violations) {
  const modal = document.getElementById("violation-modal");
  const modalBody = document.getElementById("modal-body");
  
  modalBody.innerHTML = `
    <h4 style="margin-bottom: 0.5rem; color: var(--text-primary);">全域排班規則衝突列表</h4>
    <p style="margin-bottom: 1rem; color: var(--text-secondary);">全月共有 ${violations.length} 處違規項目，請參閱並手動調整：</p>
    <ul style="max-height: 300px; overflow-y: auto;">
      ${violations.map(v => `
        <li style="margin-bottom: 0.5rem; line-height: 1.4;">
          <strong style="color: var(--color-off);">${v.employeeName || "系統"} (天數：${v.day} 號)：</strong> ${v.reason}
        </li>
      `).join("")}
    </ul>
  `;
  
  modal.classList.add("show");
}

// ==========================================================================
// 核心排班規則驗證引擎 (Validator Engine)
// ==========================================================================
function validateSchedule(scheduleMatrix) {
  const violations = [];
  const employeeStats = {}; // { empId: { work, night, off } }
  const dailyStats = {};    // { day: { day, evening, night, off } }
  
  // 初始化每日統計
  for (let d = 1; d <= totalDays; d++) {
    dailyStats[d] = { day: 0, evening: 0, night: 0, off: 0 };
  }
  
  // 1. 驗證每個員工的個人限制
  employees.forEach((emp, empIdx) => {
    if (emp.isHeadNurse) return; // 護理長跳過
    
    let workCount = 0;
    let nightCount = 0;
    let offCount = 0;
    
    const empSchedule = scheduleMatrix[empIdx];
    
    // 計算假別天數與上班天數
    for (let d = 1; d <= totalDays; d++) {
      const val = empSchedule[d - 1];
      
      // 規則 1：VV 算放假，SS 不算放假。放假寫 OFF
      const isOffCell = isOff(val);
      
      if (isOffCell) {
        offCount++;
      } else if (val === "8") {
        workCount++;
      } else if (val === "4") {
        workCount++;
      } else if (val === "12") {
        workCount++;
        nightCount++;
      } else if (val === "SS") {
        workCount++; // SS 預輸入算上班
      } else if (val) {
        // 其他非假代號一律算上班
        workCount++;
      }
      
      // 每日班別統計 (排除護理長)
      if (val === "8") dailyStats[d].day++;
      else if (val === "4") dailyStats[d].evening++;
      else if (val === "12") dailyStats[d].night++;
      else if (isOffCell) dailyStats[d].off++;
    }
    
    employeeStats[emp.id] = { work: workCount, night: nightCount, off: offCount };
    
    // A. 驗證休假天數限制 (8 - 9 天)
    if (offCount < config.minOff) {
      violations.push({
        employeeId: emp.id,
        employeeName: emp.name,
        day: 1, // 關聯至第一天
        reason: `排班放假共 ${offCount} 天，少於規定的 ${config.minOff} 天假期。`
      });
    } else if (offCount > config.maxOff) {
      violations.push({
        employeeId: emp.id,
        employeeName: emp.name,
        day: totalDays,
        reason: `排班放假共 ${offCount} 天，超出規定的 ${config.maxOff} 天假期。`
      });
    }
    
    // B. 驗證最多連續上班 6 天限制
    let consecutiveWork = 0;
    for (let d = 1; d <= totalDays; d++) {
      const val = empSchedule[d - 1];
      const isOffCell = isOff(val);
      
      if (!isOffCell && val) {
        consecutiveWork++;
        if (consecutiveWork > config.maxConsecutive) {
          violations.push({
            employeeId: emp.id,
            employeeName: emp.name,
            day: d,
            reason: `連續上班第 ${consecutiveWork} 天，超出上限 ${config.maxConsecutive} 天限制。`
          });
        }
      } else {
        consecutiveWork = 0;
      }
    }
    
    // C. 驗證夜班銜接限制
    for (let d = 1; d <= totalDays; d++) {
      const val = empSchedule[d - 1];
      
      // 大夜 (12) 前一天不能是 VV、SS
      if (val === "12" && d > 1) {
        const prevVal = empSchedule[d - 2];
        if (prevVal === "VV" || prevVal === "SS") {
          violations.push({
            employeeId: emp.id,
            employeeName: emp.name,
            day: d,
            reason: `大夜班前一天不能是 ${prevVal}。`
          });
        }
      }
      
      // 小夜 (4) 後一天不能是 VV、SS
      if (val === "4" && d < totalDays) {
        const nextVal = empSchedule[d];
        if (nextVal === "VV" || nextVal === "SS") {
          violations.push({
            employeeId: emp.id,
            employeeName: emp.name,
            day: d,
            reason: `小夜班後一天不能是 ${nextVal}。`
          });
        }
      }
    }
    
    // D. 驗證班別偏好限制
    if (emp.preferences.length > 0) {
      for (let d = 1; d <= totalDays; d++) {
        const val = empSchedule[d - 1];
        if (!val || isOff(val)) continue;
        
        // 若該天為預填，不進行偏好違規驗證 (避免 "受訓", "勿小" 等造成偏好違規)
        if (emp.originalSchedule[d - 1] !== null) continue;
        
        // 若排了工作班，比對偏好
        let hasMatch = false;
        
        // "大" -> 12, "小" -> 4, "白" -> 8
        if (val === "12" && emp.preferences.includes("大")) hasMatch = true;
        else if (val === "4" && emp.preferences.includes("小")) hasMatch = true;
        else if (val === "8" && emp.preferences.includes("白")) hasMatch = true;
        else if (val === "SS" || val === "VV" || val === "FF" || val === "W") hasMatch = true; // 預輸入特殊字不屬於偏好違規
        
        if (!hasMatch) {
          violations.push({
            employeeId: emp.id,
            employeeName: emp.name,
            day: d,
            reason: `排班為班別 [${val}]，不符合其偏好：${emp.prefStr}`
          });
        }
      }
    }
  });
  
  // 2. 驗證全域限制 (每日人數)
  for (let d = 1; d <= totalDays; d++) {
    const stats = dailyStats[d];
    
    if (stats.day !== config.demand.day) {
      violations.push({
        employeeId: "system",
        employeeName: "系統限制",
        day: d,
        reason: `本日白班(8)人數為 ${stats.day} 人，不符合目標 ${config.demand.day} 人。`
      });
    }
    
    if (stats.evening !== config.demand.evening) {
      violations.push({
        employeeId: "system",
        employeeName: "系統限制",
        day: d,
        reason: `本日小夜(4)人數為 ${stats.evening} 人，不符合目標 ${config.demand.evening} 人。`
      });
    }
    
    if (stats.night !== config.demand.night) {
      violations.push({
        employeeId: "system",
        employeeName: "系統限制",
        day: d,
        reason: `本日大夜(12)人數為 ${stats.night} 人，不符合目標 ${config.demand.night} 人。`
      });
    }
  }
  
  return {
    violations,
    employeeStats,
    dailyStats
  };
}

// ==========================================================================
// 模擬退火排班演算法 (Simulated Annealing Scheduler)
// ==========================================================================
function runAutoScheduling() {
  const maxIterations = 60000;
  let temp = 120.0;
  const alpha = 0.9997;
  const minTemp = 0.005;
  
  // 1. 初始化班表矩陣，保證「每天白、小、大夜各 5 人」硬性限制
  let matrix = employees.map(emp => [...emp.originalSchedule]);
  
  // 計算每一天已被預先占用的班別名額，並動態分配剩餘名額
  for (let d = 1; d <= totalDays; d++) {
    let dayCount = 0;
    let eveningCount = 0;
    let nightCount = 0;
    
    // 計算當天預輸入的上班人數
    employees.forEach((emp, empIdx) => {
      if (emp.isHeadNurse) return;
      const val = matrix[empIdx][d - 1];
      if (val === "8") dayCount++;
      else if (val === "4") eveningCount++;
      else if (val === "12") nightCount++;
    });
    
    let neededDay = Math.max(0, config.demand.day - dayCount);
    let neededEvening = Math.max(0, config.demand.evening - eveningCount);
    let neededNight = Math.max(0, config.demand.night - nightCount);
    
    // 取得當天可以排班且沒預填的非護理長員工清單
    const candidateIndices = [];
    employees.forEach((emp, empIdx) => {
      if (!emp.isHeadNurse && matrix[empIdx][d - 1] === null) {
        candidateIndices.push(empIdx);
      }
    });
    
    // 打亂候選名單，增加隨機性
    shuffleArray(candidateIndices);
    
    // 指派白班
    for (let i = 0; i < neededDay && candidateIndices.length > 0; i++) {
      matrix[candidateIndices.pop()][d - 1] = "8";
    }
    // 指派小夜
    for (let i = 0; i < neededEvening && candidateIndices.length > 0; i++) {
      matrix[candidateIndices.pop()][d - 1] = "4";
    }
    // 指派大夜
    for (let i = 0; i < neededNight && candidateIndices.length > 0; i++) {
      matrix[candidateIndices.pop()][d - 1] = "12";
    }
    // 剩下的候選人一律放假 (OFF)
    while (candidateIndices.length > 0) {
      matrix[candidateIndices.pop()][d - 1] = "OFF";
    }
  }
  
  // 2. 進行模擬退火尋優
  let currentCost = calculateScheduleCost(matrix);
  let bestMatrix = matrix.map(row => [...row]);
  let bestCost = currentCost;
  
  let iter = 0;
  while (temp > minTemp && iter < maxIterations && bestCost > 0) {
    iter++;
    
    // 隨機選擇一天，以及這天中兩個「沒有預填」且非護理長的員工
    const d = Math.floor(Math.random() * totalDays) + 1;
    
    const rawCandidates = [];
    employees.forEach((emp, empIdx) => {
      if (!emp.isHeadNurse && emp.originalSchedule[d - 1] === null) {
        rawCandidates.push(empIdx);
      }
    });
    
    if (rawCandidates.length < 2) continue;
    
    // 隨機選兩個交換
    const idxA = Math.floor(Math.random() * rawCandidates.length);
    let idxB = Math.floor(Math.random() * rawCandidates.length);
    while (idxA === idxB) {
      idxB = Math.floor(Math.random() * rawCandidates.length);
    }
    
    const empIdxA = rawCandidates[idxA];
    const empIdxB = rawCandidates[idxB];
    
    // 交換他們的排班
    const valA = matrix[empIdxA][d - 1];
    const valB = matrix[empIdxB][d - 1];
    
    if (valA === valB) continue; // 相同班別交換無效，直接跳過
    
    matrix[empIdxA][d - 1] = valB;
    matrix[empIdxB][d - 1] = valA;
    
    // 計算新成本
    const newCost = calculateScheduleCost(matrix);
    const delta = newCost - currentCost;
    
    // 接受條件
    if (delta < 0 || Math.random() < Math.exp(-delta / temp)) {
      currentCost = newCost;
      if (newCost < bestCost) {
        bestCost = newCost;
        bestMatrix = matrix.map(row => [...row]);
      }
    } else {
      // 拒絕，還原交換
      matrix[empIdxA][d - 1] = valA;
      matrix[empIdxB][d - 1] = valB;
    }
    
    temp *= alpha;
  }
  
  // 3. 儲存結果並更新介面
  currentSchedule = bestMatrix;
  renderScheduleAndValidate();
  
  if (bestCost === 0) {
    showToast("完美排班成功！符合所有假別與人數規則。", "success");
  } else {
    showToast(`排班完成。剩餘衝突數為 ${bestCost} 分。已亮紅標記衝突。`, "warning");
  }
}

// 輔助洗牌函數
function shuffleArray(array) {
  for (let i = array.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [array[i], array[j]] = [array[j], array[i]];
  }
}

// ==========================================================================
// 演算法成本函數計算 (Cost / Energy Function)
// ==========================================================================
function calculateScheduleCost(scheduleMatrix) {
  let cost = 0;
  
  employees.forEach((emp, empIdx) => {
    if (emp.isHeadNurse) return;
    
    const empSchedule = scheduleMatrix[empIdx];
    let offCount = 0;
    let workCount = 0;
    let consecutiveWork = 0;
    
    // 用於多重班別統計
    let dayCount = 0;
    let eveningCount = 0;
    let nightCount = 0;
    
    for (let d = 1; d <= totalDays; d++) {
      const val = empSchedule[d - 1];
      const isOffCell = isOff(val);
      
      if (isOffCell) {
        offCount++;
        consecutiveWork = 0;
      } else {
        workCount++;
        consecutiveWork++;
        
        if (val === "8") dayCount++;
        else if (val === "4") eveningCount++;
        else if (val === "12") nightCount++;
        
        // 規則 3. 連續工作限制：超出的每天罰 40 分
        if (consecutiveWork > config.maxConsecutive) {
          cost += 40;
        }
      }
      
      // 規則 2. 夜班銜接限制大夜 (12) 前一天不能是 VV、SS
      if (val === "12" && d > 1) {
        const prevVal = empSchedule[d - 2];
        if (prevVal === "VV" || prevVal === "SS") {
          cost += 100;
        }
      }
      
      // 規則 2. 夜班銜接限制小夜 (4) 後一天不能是 VV、SS
      if (val === "4" && d < totalDays) {
        const nextVal = empSchedule[d];
        if (nextVal === "VV" || nextVal === "SS") {
          cost += 100;
        }
      }
      
      // 規則 3. 班別偏好約束 (若不符合，每天罰 50 分)
      if (emp.preferences.length > 0 && !isOff(val) && val) {
        // 若該天為預輸入，免除偏好處罰
        if (emp.originalSchedule[d - 1] !== null) continue;
        
        let isPrefValid = false;
        if (val === "12" && emp.preferences.includes("大")) isPrefValid = true;
        else if (val === "4" && emp.preferences.includes("小")) isPrefValid = true;
        else if (val === "8" && emp.preferences.includes("白")) isPrefValid = true;
        else if (val === "SS" || val === "VV" || val === "FF" || val === "W") isPrefValid = true; // 預輸入特殊字不屬於偏好違規
        
        if (!isPrefValid) {
          cost += 50;
        }
      }
    }
    
    // 規則 1. 月休天數限制 (少於 min 或多於 max 每天罰 100 分)
    if (offCount < config.minOff) {
      cost += (config.minOff - offCount) * 100;
    } else if (offCount > config.maxOff) {
      cost += (offCount - config.maxOff) * 100;
    }
    
    // 規則 3. 多重班別盡量平均
    // 比如：偏好 "小大"，則小夜天數與大夜天數越接近越好 (偏差 1 天以上時，每差 1 天罰 10 分)
    if (emp.preferences.includes("小") && emp.preferences.includes("大")) {
      const diff = Math.abs(eveningCount - nightCount);
      if (diff > 1) cost += (diff - 1) * 10;
    }
    if (emp.preferences.includes("白") && emp.preferences.includes("大")) {
      const diff = Math.abs(dayCount - nightCount);
      if (diff > 1) cost += (diff - 1) * 10;
    }
    if (emp.preferences.includes("白") && emp.preferences.includes("小")) {
      const diff = Math.abs(dayCount - eveningCount);
      if (diff > 1) cost += (diff - 1) * 10;
    }
  });
  
  // 注意：每日各 5 人的限制，因為我們初始化與交換機制已保證其完美滿足 (0 誤差)，所以此處不需額外罰分
  
  return cost;
}

// ==========================================================================
// Excel 匯出與結構保持 (Excel Export Module)
// ==========================================================================
function exportToExcel() {
  if (!originalExcelData) return;
  
  try {
    const wb = XLSX.utils.book_new();
    
    // 讀取 originalExcelData 並改寫其中的值
    const sheet = Object.assign({}, originalExcelData);
    
    // 定義細框線樣式
    const thinBorder = {
      top: { style: "thin", color: { rgb: "D9D9D9" } },
      bottom: { style: "thin", color: { rgb: "D9D9D9" } },
      left: { style: "thin", color: { rgb: "D9D9D9" } },
      right: { style: "thin", color: { rgb: "D9D9D9" } }
    };
    
    // 預先計算排班驗證結果以獲取統計數據
    const result = validateSchedule(currentSchedule);
    
    // 更新員工對應的排班格子值與樣式
    employees.forEach((emp, empIdx) => {
      const rowIndex = emp.rowIndex;
      const empSchedule = currentSchedule[empIdx];
      
      for (let d = 0; d < totalDays; d++) {
        const colIdx = 4 + d; // 日期從第 5 列 (E) 開始
        const cellRef = XLSX.utils.encode_cell({ r: rowIndex, c: colIdx });
        let val = empSchedule[d];
        const isPrefilled = emp.originalSchedule[d] !== null;
        
        // 若為自動排班放假，改為寫入 FF
        if (val === "OFF") {
          val = "FF";
        }
        
        // 初始化樣式，使用 Arial 10pt，置中，帶灰色細格線
        const cellStyle = {
          font: { name: "Arial", sz: 10, color: { rgb: "000000" } },
          alignment: { horizontal: "center", vertical: "center" },
          border: thinBorder
        };
        
        // 預輸入格子背景為標準黃
        if (isPrefilled) {
          cellStyle.fill = { fgColor: { rgb: "FFFF00" } };
        }
        
        // 小夜 (4) 與大夜 (12) 字體設為紅色
        if (val === "4" || val === "12") {
          cellStyle.font.color = { rgb: "FF0000" };
        }
        
        if (val !== null && val !== undefined) {
          sheet[cellRef] = { t: 's', v: val, s: cellStyle };
        } else {
          // 空白格子也加上置中與細格線，維持 Excel 對齊與美感
          sheet[cellRef] = { t: 's', v: "", s: {
            alignment: { horizontal: "center", vertical: "center" },
            border: thinBorder
          } };
        }
      }
      
      // 更新右側的統計欄位與樣式
      const stats = result.employeeStats[emp.id] || { work: 0, night: 0, off: 0 };
      const statStyle = {
        font: { name: "Arial", sz: 10, color: { rgb: "000000" } },
        alignment: { horizontal: "center", vertical: "center" },
        border: thinBorder
      };
      
      if (!emp.isHeadNurse) {
        // 班數 (在日期之後的欄位，索引 = 4 + totalDays)
        const workCellRef = XLSX.utils.encode_cell({ r: rowIndex, c: 4 + totalDays });
        sheet[workCellRef] = { t: 'n', v: stats.work, s: statStyle };
        
        // 夜班 (索引 = 5 + totalDays)
        const nightCellRef = XLSX.utils.encode_cell({ r: rowIndex, c: 5 + totalDays });
        sheet[nightCellRef] = { t: 'n', v: stats.night, s: statStyle };
      }
    });
    
    // 更新底部的統計列數值與樣式
    statisticsRows.forEach(stat => {
      const rowIndex = stat.rowIndex;
      const label = stat.label;
      
      for (let d = 0; d < totalDays; d++) {
        const colIdx = 4 + d;
        const cellRef = XLSX.utils.encode_cell({ r: rowIndex, c: colIdx });
        
        // 計算當天非護理長員工中，值等於標籤 (12, 8, 4, 85 等) 的總人數
        let count = 0;
        employees.forEach((emp, empIdx) => {
          if (!emp.isHeadNurse) {
            let val = currentSchedule[empIdx][d];
            if (val === "OFF") val = "FF";
            if (val === label) {
              count++;
            }
          }
        });
        
        const statCellStyle = {
          font: { name: "Arial", sz: 10, color: { rgb: "000000" } },
          alignment: { horizontal: "center", vertical: "center" },
          border: thinBorder
        };
        
        sheet[cellRef] = { t: 'n', v: count, s: statCellStyle };
      }
    });
    
    XLSX.utils.book_append_sheet(wb, sheet, "排班結果");
    
    // 輸出檔案
    const exportName = `排班結果_${fileName || "schedule.xlsx"}`;
    XLSX.writeFile(wb, exportName);
    showToast("Excel 匯出成功！檔案已開始下載。", "success");
  } catch (err) {
    console.error(err);
    showToast("匯出 Excel 發生錯誤，請重試！", "error");
  }
}

// ==========================================================================
// Toast 訊息元件
// ==========================================================================
function showToast(message, type = "success", duration = 3000) {
  const container = document.getElementById("toast-container");
  const toast = document.createElement("div");
  toast.className = `toast toast-${type}`;
  
  let icon = "fa-circle-check";
  if (type === "error") icon = "fa-circle-xmark";
  else if (type === "warning") icon = "fa-circle-exclamation";
  else if (type === "info") icon = "fa-circle-info";
  
  toast.innerHTML = `
    <i class="fa-solid ${icon}"></i>
    <span>${message}</span>
  `;
  
  container.appendChild(toast);
  
  setTimeout(() => {
    toast.style.animation = "slide-in 0.3s ease reverse forwards";
    setTimeout(() => toast.remove(), 300);
  }, duration);
}
