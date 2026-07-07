/* ==========================================================================
 * VN Office 인사·출장 관리 · Application Logic
 * ========================================================================== */

const STORAGE_KEY = "vn-office-v8";  // v8: 4명 닉네임 추가 (Holden, Mike, Vanessa, DK)
const PAGE_SIZE = 50;

// ==========================================================================
// 회사 표준 시업 시각
// ==========================================================================
const SCHEDULED_START_HOUR = 8;
const SCHEDULED_START_MIN = 15;   // 08:15 이후부터 1분 단위 지각 (08:16 = 1분 지각)

function computeLateMinutes(check_in) {
  if (!check_in) return 0;
  const m = String(check_in).match(/^(\d{1,2}):(\d{2})/);
  if (!m) return 0;
  const inMin = parseInt(m[1], 10) * 60 + parseInt(m[2], 10);
  const startMin = SCHEDULED_START_HOUR * 60 + SCHEDULED_START_MIN;
  return Math.max(0, inMin - startMin);
}

// 근태 기록 하나의 late_minutes / status 를 08:15 규칙으로 재계산
function recomputeAttendance(a) {
  const hasCheckIn = !!a.check_in;
  const hasCheckOut = !!a.check_out;
  if (!hasCheckIn && !hasCheckOut) {
    a.late_minutes = 0;
    if (a.status !== "HOLIDAY" && a.status !== "REMOTE" && a.status !== "BUSINESS_TRIP") {
      a.status = "ABSENT";
    }
    return;
  }
  const late = computeLateMinutes(a.check_in);
  a.late_minutes = late;
  if (late > 0) {
    if (a.status === "NORMAL" || a.status === "LATE") a.status = "LATE";
  } else {
    if (a.status === "LATE") a.status = "NORMAL";
  }
}

// 전체 근태 재계산
function recomputeAllAttendance() {
  (state.attendance || []).forEach(recomputeAttendance);
}

let state = {
  view: "overview",
  employees: [],
  attendance: [],
  trips: [],
  filter_dept: "ALL",
  filter_month: "ALL",
  filter_status: "ALL",
  page_att: 1,
  late_tab: null,  // 대시보드 부서별 지각 Top 5 선택 탭
  loaded: false,
};

// ==========================================================================
// Persistence
// ==========================================================================
function save() {
  const persist = {
    employees: state.employees,
    attendance: state.attendance,
    trips: state.trips,
  };
  try { localStorage.setItem(STORAGE_KEY, JSON.stringify(persist)); }
  catch (e) { console.warn("localStorage save failed:", e); }
}

async function load() {
  // Try localStorage first
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (raw) {
      const p = JSON.parse(raw);
      if (Array.isArray(p.employees) && p.employees.length > 0) {
        state.employees = p.employees;
        state.attendance = p.attendance || [];
        state.trips = p.trips || [];
        recomputeAllAttendance();  // 08:15 규칙으로 재계산
        state.loaded = true;
        return;
      }
    }
  } catch (e) { console.warn("localStorage load failed:", e); }

  // Fallback: fetch data.json
  try {
    const r = await fetch("data.json");
    if (!r.ok) throw new Error("data.json fetch failed: " + r.status);
    const data = await r.json();
    state.employees = data.employees || [];
    state.attendance = data.attendance || [];
    state.trips = data.trips || [];
    recomputeAllAttendance();  // 08:15 규칙으로 재계산
    state.loaded = true;
    save();
  } catch (e) {
    console.error("Failed to load data.json:", e);
    document.getElementById("app").innerHTML = `
      <div class="loading">
        <div class="loading-title" style="color:#dc2626">데이터 로드 실패</div>
        <div class="loading-text">${e.message}</div>
        <div style="font-size:12px; color:#94a3b8; margin-top:12px;">
          data.json 파일이 같은 폴더에 있어야 합니다.
        </div>
      </div>`;
    throw e;
  }
}

function resetAll() {
  if (!confirm("모든 로컬 저장 데이터를 삭제하고 서버에서 다시 불러옵니다. 계속?")) return;
  localStorage.removeItem(STORAGE_KEY);
  location.reload();
}

function exportBackup() {
  const data = { employees: state.employees, attendance: state.attendance, trips: state.trips };
  const blob = new Blob([JSON.stringify(data, null, 2)], { type: "application/json" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `VN_backup_${new Date().toISOString().slice(0,10)}.json`;
  a.click();
}

function exportSheet(kind) {
  const rows = state[kind];
  if (!rows.length) { alert("내보낼 데이터가 없습니다."); return; }
  const flat = rows.map(r => {
    const clone = { ...r };
    if (Array.isArray(clone.partners)) clone.partners = clone.partners.map(p => p.name).join(", ");
    if (Array.isArray(clone.itinerary)) clone.itinerary = clone.itinerary.map(d => `${d.day}: ${d.note}`).join(" | ");
    return clone;
  });
  const ws = XLSX.utils.json_to_sheet(flat);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, kind);
  XLSX.writeFile(wb, `VN_${kind}_${new Date().toISOString().slice(0,10)}.xlsx`);
}

// ==========================================================================
// Excel Import (KEYWATCH attendance format)
// ==========================================================================
function readWorkbook(file) {
  return new Promise((res, rej) => {
    const reader = new FileReader();
    reader.onload = e => {
      try { res(XLSX.read(new Uint8Array(e.target.result), { type: "array", cellDates: true })); }
      catch (err) { rej(err); }
    };
    reader.onerror = rej;
    reader.readAsArrayBuffer(file);
  });
}

function findHeaderRow(rows) {
  for (let i = 0; i < Math.min(rows.length, 20); i++) {
    const cells = rows[i].map(c => String(c || "").toLowerCase().trim());
    const hasPid = cells.some(c => c.includes("person id"));
    const hasName = cells.some(c => c === "name" || c === "이름");
    const hasDate = cells.some(c => c === "date" || c === "날짜");
    if (hasPid && hasName && hasDate) return { rowIdx: i, cells };
    if ((cells.includes("name") || cells.includes("이름")) && cells.includes("date")) return { rowIdx: i, cells };
  }
  return null;
}

function normTime(v) {
  if (!v) return "";
  const m = String(v).trim().match(/^(\d{1,2}):(\d{2})(?::\d{2})?$/);
  return m ? `${String(m[1]).padStart(2, "0")}:${m[2]}` : "";
}
function normDate(v) {
  if (!v) return "";
  if (v instanceof Date) return v.toISOString().slice(0, 10);
  const s = String(v).trim();
  const m = s.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
  if (m) return `${m[1]}-${String(m[2]).padStart(2,"0")}-${String(m[3]).padStart(2,"0")}`;
  const d = new Date(s);
  if (!isNaN(d)) return d.toISOString().slice(0, 10);
  return s.slice(0, 10);
}

async function importAttendance(file) {
  try {
    const wb = await readWorkbook(file);
    const sheetName = wb.SheetNames.find(n => /detail/i.test(n)) || wb.SheetNames[0];
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header: 1, defval: "", raw: false });
    const header = findHeaderRow(rows);
    if (!header) { alert("엑셀에서 헤더 행을 찾을 수 없습니다.\n(Person ID / Name / Date 필요)"); return; }

    const idx = {};
    header.cells.forEach((c, i) => {
      if (c.includes("person id")) idx.pid = i;
      else if (c === "name" || c === "이름") idx.name = i;
      else if (c === "department" || c === "부서") idx.dept = i;
      else if (c === "position" || c === "직책") idx.pos = i;
      else if (c === "date" || c === "날짜") idx.date = i;
      else if (c === "check-in" || c === "check in" || c === "출근") idx.cin = i;
      else if (c === "check-out" || c === "check out" || c === "퇴근") idx.cout = i;
      else if (c === "late" || c === "지각") idx.late = i;
      else if (c === "gender" || c === "성별") idx.gender = i;
    });
    if (idx.pid === undefined || idx.date === undefined) {
      alert("Person ID / Date 열을 인식하지 못했습니다."); return;
    }

    let added = 0, updated = 0, empAdded = 0;
    const nextEmp = () => state.employees.length ? Math.max(...state.employees.map(e => e.id)) + 1 : 1;
    const nextAtt = () => state.attendance.length ? Math.max(...state.attendance.map(a => a.id)) + 1 : 1;

    for (let i = header.rowIdx + 1; i < rows.length; i++) {
      const row = rows[i];
      const pid = String(row[idx.pid] || "").trim();
      const name = String(row[idx.name] || "").trim();
      if (!pid || !name) continue;

      const dept = String(row[idx.dept] || "").trim();
      let emp = state.employees.find(e => e.person_id === pid);
      if (!emp) {
        emp = {
          id: nextEmp(), person_id: pid, name, department: dept,
          position: idx.pos !== undefined ? String(row[idx.pos] || "").trim() : "",
          gender: idx.gender !== undefined ? String(row[idx.gender] || "").trim() : "",
          is_scm: dept.toUpperCase().includes("SCM"),
          annual_leave: 15, remaining_leave: 15,
        };
        state.employees.push(emp); empAdded++;
      }

      const date = normDate(row[idx.date]);
      if (!date) continue;
      const cin = normTime(row[idx.cin]);
      const cout = normTime(row[idx.cout]);
      // 08:15 규칙으로 지각 계산 (엑셀의 Late 컬럼 무시)
      const late = computeLateMinutes(cin);
      let status = "NORMAL";
      if (late > 0) status = "LATE";
      if (!cin && !cout) status = "ABSENT";

      const existing = state.attendance.find(a => a.person_id === pid && a.date === date);
      if (existing) {
        existing.check_in = cin || existing.check_in;
        existing.check_out = cout || existing.check_out;
        existing.late_minutes = late;
        existing.status = status;
        updated++;
      } else {
        state.attendance.push({
          id: nextAtt(), person_id: pid, name, department: dept,
          date, check_in: cin, check_out: cout,
          late_minutes: late, status, note: "",
        });
        added++;
      }
    }

    save(); render();
    alert(`가져오기 완료\n\n· 근태 신규 ${added.toLocaleString()}건\n· 근태 갱신 ${updated.toLocaleString()}건\n· 직원 자동 등록 ${empAdded}건`);
  } catch (e) {
    alert("엑셀 파일 처리 중 오류: " + e.message);
    console.error(e);
  }
}

async function importTripPlan(file) {
  try {
    const scmEmps = state.employees.filter(e => e.is_scm);
    if (scmEmps.length === 0) { alert("SCM 인원이 먼저 등록되어 있어야 합니다."); return; }

    const wb = await readWorkbook(file);
    const expenseSheet = wb.SheetNames.find(n => /expense/i.test(n));
    const reportSheet = wb.SheetNames.find(n => /report/i.test(n));

    const trip = {
      title: file.name.replace(/\.xlsx?$/i, ""),
      employee: "", destination: "", start_date: "", end_date: "",
      purpose: "SOURCING", status: "DRAFT",
      cost_planned: 0, cost_actual: 0, currency: "VND",
      partners: [], itinerary: [],
      outcome: "", roi: null, notes: "",
    };

    // Match employee from filename: "... - Andy.xlsx"
    const m = file.name.match(/-\s*([^.]+)\.xlsx?$/i);
    if (m) {
      const person = m[1].trim();
      const found = scmEmps.find(e => e.name.toLowerCase().includes(person.toLowerCase()));
      if (found) trip.employee = found.name;
    }

    if (expenseSheet) {
      const rows = XLSX.utils.sheet_to_json(wb.Sheets[expenseSheet], { header: 1, defval: "", raw: false });
      for (const row of rows) {
        const label = String(row[0] || "").toLowerCase();
        if (label.includes("destination")) trip.destination = String(row[2] || "").trim();
        if (label.includes("purpose")) trip.notes = String(row[2] || "").trim();
        if (label.includes("departure date")) {
          const d = row[2]; if (d) trip.start_date = (d instanceof Date ? d.toISOString().slice(0,10) : String(d).slice(0,10));
        }
        if (label.includes("return date")) {
          const d = row[2]; if (d) trip.end_date = (d instanceof Date ? d.toISOString().slice(0,10) : String(d).slice(0,10));
        }
        if (label === "total") {
          const amt = Number(row[2]); if (!isNaN(amt)) trip.cost_planned = amt;
        }
      }
    }

    if (reportSheet) {
      const rows = XLSX.utils.sheet_to_json(wb.Sheets[reportSheet], { header: 1, defval: "" });
      for (let i = 1; i < rows.length; i++) {
        const name = String(rows[i][0] || "").trim();
        if (name) trip.partners.push({
          name, district: "", bookings_2026: 0, visited: false,
          contract_signed: false, contact_person: "",
          meeting_summary: "", followup_1: "", followup_2: "",
        });
      }
    }

    if (trip.destination && trip.start_date) {
      const s = trip.start_date.slice(5).replace("-","/");
      const e = trip.end_date.slice(5).replace("-","/");
      trip.title = `${trip.destination} Biz Trip (${s}~${e})`;
    }

    // Confirm modal
    const empOpts = scmEmps.map(e => `<option value="${escHTML(e.name)}" ${trip.employee === e.name ? "selected" : ""}>${escHTML(e.name)}</option>`).join("");
    showModal("출장 계획서 파싱 결과", `
      <div class="badge b-primary" style="display:block; padding:8px;">${expenseSheet ? "✓ Expense" : ""} ${reportSheet ? "✓ Report" : ""}</div>
      <div><label class="field-label">제목</label><input id="tp_title" value="${escHTML(trip.title)}" /></div>
      <div class="grid-2">
        <div><label class="field-label">담당자 *</label><select id="tp_emp"><option value="">선택</option>${empOpts}</select></div>
        <div><label class="field-label">목적지</label><input id="tp_dest" value="${escHTML(trip.destination)}" /></div>
      </div>
      <div class="grid-2">
        <div><label class="field-label">시작일</label><input id="tp_start" type="date" value="${trip.start_date}" /></div>
        <div><label class="field-label">종료일</label><input id="tp_end" type="date" value="${trip.end_date}" /></div>
      </div>
      <div class="grid-3">
        <div><label class="field-label">통화</label><select id="tp_cur">${["VND","USD","KRW","THB"].map(c => `<option ${trip.currency===c?"selected":""}>${c}</option>`).join("")}</select></div>
        <div><label class="field-label">예산</label><input id="tp_cost" type="number" value="${trip.cost_planned}" /></div>
        <div><label class="field-label">상태</label><select id="tp_status">${["DRAFT","REQUESTED","APPROVED","IN_PROGRESS"].map(s => `<option ${trip.status===s?"selected":""}>${s}</option>`).join("")}</select></div>
      </div>
      <div class="card" style="padding:10px; background:#f8fafc;">
        <div style="font-size:11px; font-weight:600;">🏨 방문 파트너: ${trip.partners.length}개</div>
        <div style="font-size:10px; color:#94a3b8; max-height:80px; overflow-y:auto; margin-top:4px;">
          ${trip.partners.map(p => escHTML(p.name)).join(" · ")}
        </div>
      </div>
    `, () => {
      trip.title = val("tp_title");
      trip.employee = val("tp_emp");
      trip.destination = val("tp_dest");
      trip.start_date = val("tp_start");
      trip.end_date = val("tp_end");
      trip.currency = val("tp_cur");
      trip.cost_planned = +val("tp_cost");
      trip.status = val("tp_status");
      if (!trip.employee) { alert("SCM 담당자를 선택하세요."); return false; }
      const nextId = state.trips.length ? Math.max(...state.trips.map(t => t.id)) + 1 : 1;
      state.trips.push({ id: nextId, ...trip });
      save(); render();
      setTimeout(() => alert(`출장 등록 완료 · 파트너 ${trip.partners.length}개`), 100);
      return true;
    });
  } catch (e) {
    alert("계획서 처리 중 오류: " + e.message);
    console.error(e);
  }
}

// ==========================================================================
// UI Helpers
// ==========================================================================
function escHTML(s) {
  return String(s == null ? "" : s)
    .replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;");
}
function val(id) { return document.getElementById(id).value.trim(); }
function fmt(n) { return (n || 0).toLocaleString(); }
function curSym(c) { return c === "VND" ? "₫" : c === "USD" ? "$" : c === "KRW" ? "₩" : ""; }

function showModal(title, body, onSave, extraFooter = "") {
  const html = `
    <div class="modal-backdrop" id="modal">
      <div class="modal">
        <div class="modal-head">
          <div>${title}</div>
          <button class="modal-close" onclick="closeModal()">✕</button>
        </div>
        <div class="modal-body">${body}</div>
        <div class="modal-foot">
          ${extraFooter}
          <button class="btn btn-outline" onclick="closeModal()">취소</button>
          <button class="btn btn-primary" id="modal-save">저장</button>
        </div>
      </div>
    </div>`;
  document.body.insertAdjacentHTML("beforeend", html);
  document.getElementById("modal-save").onclick = () => { if (onSave() !== false) closeModal(); };
}
function closeModal() { document.getElementById("modal")?.remove(); }

// ==========================================================================
// Router
// ==========================================================================
const NAV = [
  { key: "overview",   label: "대시보드",   icon: "📊" },
  { key: "employees",  label: "인원 (VN)",  icon: "👥" },
  { key: "attendance", label: "근태",       icon: "🕘" },
  { key: "trips",      label: "SCM 출장",   icon: "✈️" },
  { key: "reports",    label: "리포트",     icon: "📈" },
];

function go(view) {
  state.view = view;
  state.page_att = 1;
  render();
  // Scroll content to top on view change
  const content = document.querySelector(".content");
  if (content) content.scrollTop = 0;
}

function renderShell() {
  const scm = state.employees.filter(e => e.is_scm).length;
  const active = state.trips.filter(t => ["APPROVED","IN_PROGRESS"].includes(t.status)).length;
  const stats = `직원 ${state.employees.length} (SCM ${scm}) · 근태 ${state.attendance.length.toLocaleString()} · 진행 출장 ${active}`;
  const titleMap = { overview: "대시보드", employees: "VN Office 인원", attendance: "근태", trips: "SCM 출장", reports: "리포트" };

  return `
    <div class="app">
      <aside class="sidebar">
        <div class="brand">
          <div class="brand-icon">VN</div>
          <div class="brand-text"><b>VN Office</b><small>인사·SCM 출장</small></div>
        </div>
        <nav class="nav">
          ${NAV.map(n => `
            <button class="nav-btn ${state.view === n.key ? "active" : ""}" onclick="go('${n.key}')">
              <span><span class="nav-icon">${n.icon}</span>${n.label}</span>
              ${n.key === "trips" ? `<span class="badge b-scm">${scm}</span>` : ""}
            </button>
          `).join("")}
        </nav>
        <div class="sidebar-footer">
          <div>v1.0 · GitHub Pages</div>
          <button onclick="exportBackup()">전체 백업 (JSON)</button>
        </div>
      </aside>
      <div class="main">
        <header class="header">
          <div class="page-title">${titleMap[state.view]}</div>
          <div class="stats">${stats}</div>
          <button class="btn btn-outline btn-sm" onclick="resetAll()">데이터 초기화</button>
        </header>
        <div class="content" id="content"></div>
      </div>
    </div>
  `;
}

function render() {
  if (!state.loaded) return;
  document.getElementById("app").innerHTML = renderShell();
  const content = document.getElementById("content");
  const views = {
    overview: viewOverview,
    employees: viewEmployees,
    attendance: viewAttendance,
    trips: viewTrips,
    reports: viewReports,
  };
  content.innerHTML = (views[state.view] || viewOverview)();
}

// ==========================================================================
// View: Overview
// ==========================================================================
function viewOverview() {
  const today = new Date().toISOString().slice(0, 10);
  const thisMonth = today.slice(0, 7);
  const monthAtt = state.attendance.filter(a => (a.date || "").startsWith(thisMonth));
  const monthLate = monthAtt.filter(a => a.status === "LATE").length;

  const scm = state.employees.filter(e => e.is_scm).length;
  const active = state.trips.filter(t => ["APPROVED","IN_PROGRESS"].includes(t.status));
  const completed = state.trips.filter(t => t.status === "COMPLETED");
  const rois = completed.filter(t => t.roi != null).map(t => t.roi);
  const avgRoi = rois.length ? rois.reduce((s,r) => s + r, 0) / rois.length : 0;

  const empByDept = {};
  state.employees.forEach(e => { empByDept[e.department] = (empByDept[e.department] || 0) + 1; });

  // 부서별 지각 Top 5 (누적)
  const lateByDeptPerson = {}; // { dept: { person_id: count } }
  state.attendance.forEach(a => {
    if (a.status !== "LATE") return;
    const dept = a.department || "미지정";
    if (!lateByDeptPerson[dept]) lateByDeptPerson[dept] = {};
    lateByDeptPerson[dept][a.person_id] = (lateByDeptPerson[dept][a.person_id] || 0) + 1;
  });
  // 부서별 Top 5 배열로 변환
  const deptTopLate = Object.entries(lateByDeptPerson)
    .map(([dept, personCounts]) => ({
      dept,
      isSCM: dept.toUpperCase().includes("SCM"),
      top: Object.entries(personCounts).sort((a,b) => b[1] - a[1]).slice(0, 5),
      totalDeptLate: Object.values(personCounts).reduce((s,n) => s + n, 0),
    }))
    .sort((a,b) => b.totalDeptLate - a.totalDeptLate);

  return `
    <div class="kpi-grid">
      ${kpi("VN 총 인원", state.employees.length, "명", "primary")}
      ${kpi("SCM 인원", scm, "명", "scm")}
      ${kpi(`${thisMonth} 지각`, monthLate, "건", monthLate > 10 ? "danger" : monthLate > 0 ? "warn" : "success")}
      ${kpi("SCM 진행 출장", active.length, "건", "primary")}
      ${kpi("SCM 평균 ROI", avgRoi ? avgRoi.toFixed(1) : "—", avgRoi ? "x" : "", avgRoi >= 3 ? "success" : "")}
    </div>

    <div class="grid-3">
      <div class="card">
        <h3>부서별 인원</h3>
        ${barChart(Object.entries(empByDept).sort((a,b) => b[1] - a[1]).map(([d, n]) => ({
          label: d, value: n, max: state.employees.length,
          scm: d.toUpperCase().includes("SCM"),
        })))}
      </div>

      <div class="card">
        <h3>부서별 누적 지각 Top 5</h3>
        ${deptTopLate.length === 0 ? empty("지각 기록 없음") : (() => {
          // 초기 선택 탭: state 에 없으면 SCM 우선, 없으면 총 지각 많은 부서
          const availableTabs = deptTopLate.filter(d => d.top.length > 0);
          if (!state.late_tab || !availableTabs.find(d => d.dept === state.late_tab)) {
            const scmTab = availableTabs.find(d => d.isSCM);
            state.late_tab = scmTab ? scmTab.dept : availableTabs[0].dept;
          }
          const selected = availableTabs.find(d => d.dept === state.late_tab) || availableTabs[0];
          const maxCnt = selected.top[0][1];

          return `
            <div style="display:flex; gap:4px; flex-wrap:wrap; margin-bottom:12px; border-bottom:1px solid #e2e8f0; padding-bottom:8px;">
              ${availableTabs.map(({dept, isSCM, totalDeptLate}) => {
                const active = dept === state.late_tab;
                return `
                  <button onclick="state.late_tab='${escHTML(dept)}'; render();"
                    style="border:none; padding:5px 10px; border-radius:6px; cursor:pointer; font-size:11px; font-weight:500;
                           ${active ? (isSCM ? 'background:#4f46e5; color:#fff;' : 'background:#2563eb; color:#fff;')
                                    : 'background:#f1f5f9; color:#475569;'}">
                    ${isSCM ? "⭐ " : ""}${escHTML(dept.replace("Office/", ""))} · ${totalDeptLate}
                  </button>
                `;
              }).join("")}
            </div>
            <div class="bar-chart">
              ${selected.top.map(([pid, cnt]) => {
                const emp = state.employees.find(e => e.person_id === pid);
                return `
                  <div class="bar-row">
                    <span class="bar-label ${emp && emp.is_scm ? "scm" : ""}">${escHTML(emp ? emp.name : pid)}</span>
                    <div class="bar-track"><div class="bar-fill warn" style="width:${(cnt/maxCnt)*100}%"></div></div>
                    <span class="bar-value">${cnt}회</span>
                  </div>
                `;
              }).join("")}
            </div>
          `;
        })()}
      </div>

      <div class="card">
        <h3>✈️ SCM 진행 중 출장</h3>
        ${active.length === 0 ? empty("진행 중 출장 없음") : `
          <div class="stack">
            ${active.map(t => {
              const dEnd = Math.ceil((new Date(t.end_date) - new Date(today)) / 86400000);
              const dStart = Math.ceil((new Date(t.start_date) - new Date(today)) / 86400000);
              const label = dStart > 0 ? `D-${dStart} 출발` : dEnd >= 0 ? `귀국 D-${dEnd}` : `귀국 D+${Math.abs(dEnd)}`;
              return `
                <div class="trip-card" onclick="go('trips')">
                  <div class="trip-card-title">${escHTML(t.title)}</div>
                  <div class="trip-card-meta">${escHTML(t.destination || "—")}</div>
                  <div class="trip-card-meta">${escHTML(t.employee)} · <span style="color:#2563eb;">${label}</span></div>
                </div>
              `;
            }).join("")}
          </div>
        `}
      </div>
    </div>

    ${completed.length > 0 ? `
      <div class="card mt-4">
        <h3>📈 완료된 SCM 출장 성과</h3>
        <div class="table-wrap"><table>
          <thead><tr>
            <th>Trip</th><th>담당자</th><th>목적지</th>
            <th class="right">예산</th><th class="right">실지출</th>
            <th class="right">ROI</th><th>결과 요약</th>
          </tr></thead>
          <tbody>
            ${completed.map(t => `
              <tr>
                <td>#${t.id} ${escHTML(t.title)}</td>
                <td>${escHTML(t.employee)}</td>
                <td>${escHTML(t.destination)}</td>
                <td class="right">${curSym(t.currency)}${fmt(t.cost_planned)}</td>
                <td class="right">${curSym(t.currency)}${fmt(t.cost_actual)}</td>
                <td class="right"><b>${t.roi ? t.roi.toFixed(1) + "x" : "—"}</b></td>
                <td>${escHTML(t.outcome || "—")}</td>
              </tr>
            `).join("")}
          </tbody>
        </table></div>
      </div>
    ` : ""}
  `;
}

function kpi(label, value, unit, tone = "primary") {
  return `
    <div class="kpi ${tone}">
      <div class="kpi-label">${label}</div>
      <div class="kpi-value">${value}<small>${unit}</small></div>
    </div>
  `;
}

function barChart(rows) {
  if (rows.length === 0) return empty("데이터 없음");
  const maxVal = Math.max(...rows.map(r => r.value));
  return `
    <div class="bar-chart">
      ${rows.map(r => `
        <div class="bar-row">
          <span class="bar-label ${r.scm ? "scm" : ""}">${escHTML(r.label)}${r.scm ? " ⭐" : ""}</span>
          <div class="bar-track"><div class="bar-fill ${r.scm ? "scm" : ""}" style="width:${(r.value/maxVal)*100}%"></div></div>
          <span class="bar-value">${r.value}${r.unit || ""}</span>
        </div>
      `).join("")}
    </div>
  `;
}

function empty(msg) { return `<div class="empty">${msg}</div>`; }

// ==========================================================================
// View: Employees
// ==========================================================================
function viewEmployees() {
  const depts = ["ALL", ...new Set(state.employees.map(e => e.department))];
  const filtered = state.employees.filter(e => state.filter_dept === "ALL" || e.department === state.filter_dept);

  return `
    <div class="flex center gap-3 wrap mt-2">
      <h2 style="margin:0; font-size:16px;">VN Office 인원 <span style="color:#94a3b8; font-size:13px;">(${filtered.length} / ${state.employees.length})</span></h2>
      <div class="ml-auto flex gap-2">
        <button class="btn btn-outline" onclick="exportSheet('employees')">📤 엑셀 내보내기</button>
      </div>
    </div>

    <div class="mt-4">
      <div class="chip-label">부서</div>
      <div class="chips">
        ${depts.map(d => `<button class="chip ${state.filter_dept === d ? "active" : ""}" onclick="state.filter_dept='${d}'; render();">
          ${d === "ALL" ? "전체" : escHTML(d)} ${d !== "ALL" ? `(${state.employees.filter(e => e.department === d).length})` : ""}
        </button>`).join("")}
      </div>
    </div>

    <div class="card mt-4" style="padding:0;">
      <div class="table-wrap"><table>
        <thead><tr>
          <th>Person ID</th><th>이름</th><th>부서</th><th>직책</th><th>성별</th>
          <th class="right">잔여 연차</th>
        </tr></thead>
        <tbody>
          ${filtered.map(e => `
            <tr>
              <td class="mono">${escHTML(e.person_id || "—")}</td>
              <td>
                <b>${escHTML(e.name)}</b>
                ${e.is_scm ? (e.position === "SCM Head" ? `<span class="badge b-scm" style="margin-left:6px; background:#4f46e5; color:#fff;">👑 SCM Head</span>` : `<span class="badge b-scm" style="margin-left:6px;">SCM</span>`) : ""}
              </td>
              <td>${escHTML(e.department || "—")}</td>
              <td>${escHTML(e.position || "—")}</td>
              <td>${escHTML(e.gender || "—")}</td>
              <td class="right">${e.remaining_leave} / ${e.annual_leave}일</td>
            </tr>
          `).join("")}
        </tbody>
      </table></div>
    </div>
  `;
}

// editEmployee / deleteEmployee 제거됨 (수기 입력 기능 제거)

// ==========================================================================
// View: Attendance
// ==========================================================================
function viewAttendance() {
  const depts = ["ALL", ...new Set(state.employees.map(e => e.department))];
  const months = ["ALL", ...new Set(state.attendance.map(a => (a.date || "").slice(0, 7)).filter(Boolean))].sort().reverse();
  const statuses = ["ALL","NORMAL","LATE","ABSENT"];

  let filtered = state.attendance.filter(a => {
    if (state.filter_dept !== "ALL" && a.department !== state.filter_dept) return false;
    if (state.filter_month !== "ALL" && !a.date.startsWith(state.filter_month)) return false;
    if (state.filter_status !== "ALL" && a.status !== state.filter_status) return false;
    return true;
  });
  filtered.sort((a,b) => (b.date || "").localeCompare(a.date || "") || (a.name || "").localeCompare(b.name || ""));

  const totalPages = Math.max(1, Math.ceil(filtered.length / PAGE_SIZE));
  if (state.page_att > totalPages) state.page_att = 1;
  const paginated = filtered.slice((state.page_att - 1) * PAGE_SIZE, state.page_att * PAGE_SIZE);

  return `
    <div class="flex center gap-3 wrap">
      <h2 style="margin:0; font-size:16px;">근태 <span style="color:#94a3b8; font-size:13px;">(${filtered.length.toLocaleString()} / ${state.attendance.length.toLocaleString()})</span></h2>
      <div class="ml-auto flex gap-2">
        <button class="btn btn-outline" onclick="exportSheet('attendance')">📤 필터 결과 내보내기</button>
      </div>
    </div>

    <div class="mt-3">
      ${dropzone("attendance", "근태 엑셀 (KEYWATCH 형식) 드래그")}
    </div>

    <div class="card mt-4">
      <div class="chip-label">부서</div>
      <div class="chips">
        ${depts.map(d => `<button class="chip ${state.filter_dept === d ? "active" : ""}" onclick="state.filter_dept='${d}'; state.page_att=1; render();">${d === "ALL" ? "전체" : escHTML(d)}</button>`).join("")}
      </div>
      <div class="chip-label mt-2">월</div>
      <div class="chips">
        ${months.map(m => `<button class="chip ${state.filter_month === m ? "active" : ""}" onclick="state.filter_month='${m}'; state.page_att=1; render();">${m === "ALL" ? "전체" : m}</button>`).join("")}
      </div>
      <div class="chip-label mt-2">상태</div>
      <div class="chips">
        ${statuses.map(s => `<button class="chip ${state.filter_status === s ? "active" : ""}" onclick="state.filter_status='${s}'; state.page_att=1; render();">${s === "ALL" ? "전체" : s}</button>`).join("")}
      </div>
    </div>

    <div class="card mt-4" style="padding:0;">
      <div class="table-wrap"><table>
        <thead><tr>
          <th>날짜</th><th>이름</th><th>부서</th><th>출근</th><th>퇴근</th>
          <th class="right">지각(분)</th><th>상태</th>
        </tr></thead>
        <tbody>
          ${paginated.map(a => `
            <tr>
              <td>${a.date}</td>
              <td>
                <b>${escHTML(a.name)}</b>
                ${a.department && a.department.toUpperCase().includes("SCM") ? `<span class="badge b-scm" style="margin-left:6px;">SCM</span>` : ""}
              </td>
              <td class="mono">${escHTML(a.department || "—")}</td>
              <td>${a.check_in || "—"}</td>
              <td>${a.check_out || "—"}</td>
              <td class="right ${a.late_minutes > 0 ? "text-late" : ""}">${a.late_minutes || "—"}</td>
              <td><span class="badge b-${a.status === "LATE" ? "warn" : a.status === "ABSENT" ? "danger" : "success"}">${a.status}</span></td>
            </tr>
          `).join("")}
        </tbody>
      </table></div>
    </div>

    ${totalPages > 1 ? `
      <div class="pagination">
        <button class="btn btn-outline btn-sm" ${state.page_att === 1 ? "disabled" : ""} onclick="state.page_att=Math.max(1,state.page_att-1); render();">이전</button>
        <span>${state.page_att} / ${totalPages} 페이지</span>
        <button class="btn btn-outline btn-sm" ${state.page_att === totalPages ? "disabled" : ""} onclick="state.page_att=Math.min(${totalPages},state.page_att+1); render();">다음</button>
      </div>
    ` : ""}
  `;
}

function dropzone(kind, label) {
  const handler = kind === "attendance" ? "handleAttFile" : "handleTripFile";
  return `
    <div class="dropzone" ondrop="handleDrop(event, '${kind}')" ondragover="event.preventDefault(); event.currentTarget.classList.add('dragover');" ondragleave="event.currentTarget.classList.remove('dragover');">
      <div class="dropzone-icon">📄</div>
      <div class="dropzone-text">${label}</div>
      <label class="btn btn-outline btn-sm" style="cursor:pointer;">
        파일 선택
        <input type="file" accept=".xlsx,.xls" onchange="${handler}(event)" />
      </label>
      <div class="dropzone-hint">${kind === "attendance" ? "Details 시트 자동 인식 · Person ID / Date / Late 헤더" : "Plan / Expense / Report 시트 자동 인식"}</div>
    </div>
  `;
}
function handleDrop(e, kind) {
  e.preventDefault();
  e.currentTarget.classList.remove("dragover");
  if (e.dataTransfer.files.length) {
    if (kind === "attendance") importAttendance(e.dataTransfer.files[0]);
    else importTripPlan(e.dataTransfer.files[0]);
  }
}
function handleAttFile(e) { if (e.target.files.length) importAttendance(e.target.files[0]); }
function handleTripFile(e) { if (e.target.files.length) importTripPlan(e.target.files[0]); }

// editAttendance / deleteAttendance 제거됨 (수기 입력 기능 제거)

// ==========================================================================
// View: Trips
// ==========================================================================
function viewTrips() {
  const cols = ["DRAFT","REQUESTED","APPROVED","IN_PROGRESS","COMPLETED"];
  const scmCount = state.employees.filter(e => e.is_scm).length;
  return `
    <div class="flex center gap-3 wrap">
      <h2 style="margin:0; font-size:16px;">SCM 출장 <span style="color:#94a3b8; font-size:13px;">(${state.trips.length}건 · 대상 ${scmCount}명)</span></h2>
      <div class="ml-auto flex gap-2">
        <button class="btn btn-outline" onclick="exportSheet('trips')">📤 엑셀 내보내기</button>
      </div>
    </div>

    <div class="card mt-3" style="background:#e0e7ff; border-color:#c7d2fe; padding:10px 14px; font-size:12px; color:#3730a3;">
      ℹ️ 출장·성과 tracking 은 <b>SCM 부서 인원(${scmCount}명)</b> 만 대상입니다.
    </div>

    <div class="mt-3">
      ${dropzone("tripplan", "출장 계획서 엑셀 드래그")}
    </div>

    <div class="kanban mt-4">
      ${cols.map(col => {
        const trips = state.trips.filter(t => t.status === col);
        return `
          <div class="kanban-col">
            <div class="kanban-head">
              <span>${col}</span>
              <span class="badge b-muted">${trips.length}</span>
            </div>
            <div class="kanban-body">
              ${trips.length === 0 ? `<div class="empty" style="padding:12px; font-size:11px;">없음</div>` :
                trips.map(t => {
                  const pCount = (t.partners || []).length;
                  const pVisited = (t.partners || []).filter(p => p.visited).length;
                  const sym = curSym(t.currency);
                  return `
                    <div class="trip-card" onclick="editTrip(${t.id})">
                      <div class="trip-card-id">#${t.id}</div>
                      <div class="trip-card-title">${escHTML(t.title)}</div>
                      <div class="trip-card-meta">${escHTML(t.destination || "—")}</div>
                      <div class="trip-card-meta">${t.start_date} ~ ${t.end_date}</div>
                      ${pCount > 0 ? `<div class="trip-card-meta" style="color:#4f46e5; margin-top:4px;">🏨 파트너 ${pVisited}/${pCount}</div>` : ""}
                      <div class="trip-card-footer">
                        <span class="trip-card-owner truncate">${escHTML(t.employee)}</span>
                        <span class="trip-card-cost">${sym}${fmt(t.cost_planned)}</span>
                      </div>
                      ${t.status === "COMPLETED" && !t.outcome ? '<div class="badge b-warn mt-2">결과보고 대기</div>' : ""}
                      ${t.roi ? `<div class="badge b-success mt-2">ROI ${t.roi.toFixed(1)}x</div>` : ""}
                    </div>
                  `;
                }).join("")
              }
            </div>
          </div>
        `;
      }).join("")}
    </div>
  `;
}

// Read-only trip detail viewer (수기 편집 기능 제거됨)
function editTrip(id) {
  const trip = state.trips.find(t => t.id === id);
  if (!trip) return;

  const partners = trip.partners || [];
  const itinerary = trip.itinerary || [];
  const cur = trip.currency || "USD";

  const row = (label, value) => `
    <div style="display:flex; padding:6px 0; border-bottom:1px solid #f1f5f9; font-size:13px;">
      <div style="width:110px; color:#64748b; font-size:12px;">${label}</div>
      <div style="flex:1;">${value}</div>
    </div>`;

  showModalReadOnly(`SCM 출장 상세 #${trip.id}`, `
    <div style="font-size:16px; font-weight:600; margin-bottom:10px;">${escHTML(trip.title || "—")}</div>

    ${row("담당자", escHTML(trip.employee || "—"))}
    ${row("목적지", escHTML(trip.destination || "—"))}
    ${row("기간", `${trip.start_date || "—"} ~ ${trip.end_date || "—"}`)}
    ${row("목적", escHTML(trip.purpose || "—"))}
    ${row("상태", `<span class="badge b-primary">${escHTML(trip.status || "—")}</span>`)}
    ${row("예산", `${curSym(cur)}${fmt(trip.cost_planned)} ${cur}`)}
    ${row("실지출", `${curSym(cur)}${fmt(trip.cost_actual)} ${cur}`)}
    ${trip.roi != null ? row("실제 ROI", `<b>${trip.roi.toFixed(1)}x</b>`) : ""}

    ${itinerary.length > 0 ? `
      <div class="card" style="padding:10px; background:#f8fafc; margin-top:14px;">
        <div style="font-size:11px; font-weight:600; margin-bottom:6px;">🗓️ 일정</div>
        <div style="font-size:12px; color:#475569;">
          ${itinerary.map(d => `<div style="padding:2px 0;"><b>${escHTML(d.day)}</b> · ${escHTML(d.note)}</div>`).join("")}
        </div>
      </div>
    ` : ""}

    ${partners.length > 0 ? `
      <div style="margin-top:14px;">
        <div style="font-size:11px; font-weight:600; margin-bottom:6px;">🏨 방문 파트너 (${partners.length})</div>
        <div class="partners-list">
          ${partners.map(p => `
            <div class="partner-row">
              <div style="width:20px; text-align:center; font-size:14px;">${p.visited ? "✓" : "·"}</div>
              <div style="flex:1;">
                <div class="partner-name">${escHTML(p.name)}</div>
                <div class="partner-meta">${escHTML(p.district || "—")}${p.bookings_2026 ? ` · YTD ${p.bookings_2026} 예약` : ""}</div>
              </div>
              ${p.contract_signed ? '<span class="badge b-success">SIGNED</span>' : ""}
            </div>
          `).join("")}
        </div>
      </div>
    ` : ""}

    ${trip.outcome ? `
      <div style="border-top:1px solid #e2e8f0; padding-top:12px; margin-top:14px;">
        <div style="font-size:11px; font-weight:600; margin-bottom:6px;">📋 출장 결과</div>
        <div style="font-size:13px; color:#475569; white-space:pre-wrap;">${escHTML(trip.outcome)}</div>
      </div>
    ` : ""}

    ${trip.notes ? `
      <div style="margin-top:12px;">
        <div style="font-size:11px; font-weight:600; margin-bottom:6px;">📝 노트</div>
        <div style="font-size:13px; color:#475569; white-space:pre-wrap;">${escHTML(trip.notes)}</div>
      </div>
    ` : ""}
  `);
}

// Read-only modal (닫기 버튼만 있는 뷰 전용 모달)
function showModalReadOnly(title, body) {
  const html = `
    <div class="modal-backdrop" id="modal">
      <div class="modal">
        <div class="modal-head">
          <div>${title}</div>
          <button class="modal-close" onclick="closeModal()">✕</button>
        </div>
        <div class="modal-body">${body}</div>
        <div class="modal-foot">
          <button class="btn btn-outline" onclick="closeModal()">닫기</button>
        </div>
      </div>
    </div>`;
  document.body.insertAdjacentHTML("beforeend", html);
}

// ==========================================================================
// View: Reports (pure CSS/SVG charts - no library)
// ==========================================================================
function viewReports() {
  // Monthly stats
  const months = [...new Set(state.attendance.map(a => (a.date || "").slice(0, 7)).filter(Boolean))].sort();
  const monthly = months.map(m => ({
    month: m,
    late: state.attendance.filter(a => a.date.startsWith(m) && a.status === "LATE").length,
    absent: state.attendance.filter(a => a.date.startsWith(m) && a.status === "ABSENT").length,
    total: state.attendance.filter(a => a.date.startsWith(m)).length,
  }));
  const monthlyMax = Math.max(1, ...monthly.map(m => Math.max(m.late, m.absent)));

  // Dept stats
  const deptStats = {};
  state.attendance.forEach(a => {
    if (!a.department) return;
    if (!deptStats[a.department]) deptStats[a.department] = { total: 0, late: 0, absent: 0 };
    deptStats[a.department].total++;
    if (a.status === "LATE") deptStats[a.department].late++;
    if (a.status === "ABSENT") deptStats[a.department].absent++;
  });

  // Attendance status distribution
  const attStatus = {};
  state.attendance.forEach(a => { attStatus[a.status] = (attStatus[a.status] || 0) + 1; });
  const attTotal = state.attendance.length;

  // Trip status distribution
  const tripStatus = {};
  state.trips.forEach(t => { tripStatus[t.status] = (tripStatus[t.status] || 0) + 1; });
  const tripTotal = state.trips.length;

  // Dept employees
  const deptEmp = {};
  state.employees.forEach(e => { deptEmp[e.department] = (deptEmp[e.department] || 0) + 1; });

  return `
    <div class="flex center gap-3">
      <h2 style="margin:0; font-size:16px;">리포트</h2>
      <button class="btn btn-outline ml-auto" onclick="exportBackup()">📤 전체 백업 (JSON)</button>
    </div>

    <div class="grid-2 mt-4">
      <div class="card">
        <h3>부서별 인원 분포</h3>
        ${barChart(Object.entries(deptEmp).sort((a,b) => b[1] - a[1]).map(([d, n]) => ({
          label: d, value: n, scm: d.toUpperCase().includes("SCM"),
        })))}
      </div>

      <div class="card">
        <h3>월별 지각·결근 추이</h3>
        <div class="monthly-chart">
          ${monthly.map(m => `
            <div class="month-col">
              <div class="month-bars">
                <div class="month-bar late" style="height:${(m.late/monthlyMax)*100}%;" title="지각 ${m.late}"></div>
                <div class="month-bar absent" style="height:${(m.absent/monthlyMax)*100}%;" title="결근 ${m.absent}"></div>
              </div>
              <div class="month-label">${m.month.slice(5)}</div>
              <div style="font-size:10px; color:#64748b;">L${m.late} · A${m.absent}</div>
            </div>
          `).join("")}
        </div>
        <div class="month-legend">
          <div class="legend-item"><span class="legend-dot" style="background:#f59e0b"></span>지각</div>
          <div class="legend-item"><span class="legend-dot" style="background:#ef4444"></span>결근</div>
        </div>
      </div>

      <div class="card">
        <h3>근태 상태 분포</h3>
        ${doughnut(attStatus, attTotal, { NORMAL: "#22c55e", LATE: "#f59e0b", ABSENT: "#ef4444", REMOTE: "#94a3b8", BUSINESS_TRIP: "#3b82f6", HOLIDAY: "#cbd5e1" })}
      </div>

      <div class="card">
        <h3>SCM 출장 상태</h3>
        ${doughnut(tripStatus, tripTotal, { DRAFT: "#94a3b8", REQUESTED: "#f59e0b", APPROVED: "#3b82f6", IN_PROGRESS: "#22c55e", COMPLETED: "#0ea5e9", CANCELLED: "#ef4444" })}
      </div>
    </div>

    <div class="card mt-4">
      <h3>부서별 근태 요약</h3>
      <div class="table-wrap"><table>
        <thead><tr>
          <th>부서</th><th class="right">총 근태</th>
          <th class="right">지각</th><th class="right">결근</th><th class="right">지각률</th>
        </tr></thead>
        <tbody>
          ${Object.entries(deptStats).sort((a,b) => (b[1].late/b[1].total) - (a[1].late/a[1].total)).map(([d, s]) => {
            const rate = ((s.late / s.total) * 100).toFixed(1);
            const isSCM = d.toUpperCase().includes("SCM");
            return `
              <tr>
                <td><b>${escHTML(d)}</b>${isSCM ? `<span class="badge b-scm" style="margin-left:6px;">SCM</span>` : ""}</td>
                <td class="right">${s.total.toLocaleString()}</td>
                <td class="right ${s.late > 20 ? "text-late" : ""}">${s.late}</td>
                <td class="right ${s.absent > 5 ? "text-absent" : ""}">${s.absent}</td>
                <td class="right"><b>${rate}%</b></td>
              </tr>
            `;
          }).join("")}
        </tbody>
      </table></div>
    </div>

    <div class="card mt-4">
      <h3>인원별 근태 요약</h3>
      <div class="table-wrap"><table>
        <thead><tr>
          <th>이름</th><th>부서</th><th class="right">근태</th>
          <th class="right">지각</th><th class="right">결근</th><th class="right">잔여 연차</th>
        </tr></thead>
        <tbody>
          ${state.employees.map(e => {
            const att = state.attendance.filter(a => a.person_id === e.person_id);
            const late = att.filter(a => a.status === "LATE").length;
            const absent = att.filter(a => a.status === "ABSENT").length;
            return `
              <tr>
                <td><b>${escHTML(e.name)}</b>${e.is_scm ? (e.position === "SCM Head" ? `<span class="badge b-scm" style="margin-left:6px; background:#4f46e5; color:#fff;">👑 SCM Head</span>` : `<span class="badge b-scm" style="margin-left:6px;">SCM</span>`) : ""}</td>
                <td class="mono">${escHTML(e.department || "—")}</td>
                <td class="right">${att.length}</td>
                <td class="right ${late > 3 ? "text-absent" : ""}">${late}</td>
                <td class="right ${absent > 0 ? "text-late" : ""}">${absent}</td>
                <td class="right">${e.remaining_leave}일</td>
              </tr>
            `;
          }).join("")}
        </tbody>
      </table></div>
    </div>
  `;
}

function doughnut(counts, total, colors) {
  const entries = Object.entries(counts).sort((a,b) => b[1] - a[1]);
  if (total === 0) return empty("데이터 없음");

  const r = 45, cx = 60, cy = 60;
  const c = 2 * Math.PI * r;
  let offset = 0;
  const paths = entries.map(([k, v]) => {
    const frac = v / total;
    const dash = frac * c;
    const color = colors[k] || "#94a3b8";
    const el = `<circle r="${r}" cx="${cx}" cy="${cy}" fill="transparent" stroke="${color}" stroke-width="18" stroke-dasharray="${dash} ${c}" stroke-dashoffset="${-offset}" transform="rotate(-90 ${cx} ${cy})" />`;
    offset += dash;
    return el;
  }).join("");

  const legend = entries.map(([k, v]) => {
    const pct = ((v / total) * 100).toFixed(1);
    const color = colors[k] || "#94a3b8";
    return `
      <div class="legend-item">
        <div><span class="legend-dot" style="background:${color}"></span>${k}</div>
        <b>${v.toLocaleString()} (${pct}%)</b>
      </div>
    `;
  }).join("");

  return `
    <div class="doughnut-wrap">
      <svg class="doughnut" viewBox="0 0 120 120">${paths}</svg>
      <div class="doughnut-legend">${legend}</div>
    </div>
  `;
}

// ==========================================================================
// Boot
// ==========================================================================
(async function() {
  try {
    await load();
    render();
  } catch (e) {
    console.error("Boot failed:", e)