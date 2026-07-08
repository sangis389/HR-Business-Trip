/* ==========================================================================
 * VN Office 인사·출장 관리 · Application Logic
 * ========================================================================== */

const STORAGE_KEY = "vn-office-v34";  // v34: 월/유형 필터 드롭다운 형식으로 변경 (근태·연차, SCM 출장, 통합 캘린더)
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
  employees: [], leaves: [],
  attendance: [],
  trips: [],
  filter_dept: "ALL",
  filter_month: "ALL",
  filter_status: "ALL",
  page_att: 1,
  late_tab: null,  // 대시보드 부서별 지각 Top 5 선택 탭
  att_tab: "daily",  // 근태 페이지 서브탭: daily / leaves / summary
  filter_leave_type: "ALL",  // 휴가 이력 타입 필터
  cal_month: null,           // 캘린더 뷰 월
  cal_scope: "SCM",          // 캘린더 범위: SCM | ALL
  filter_trip_month: "ALL",  // SCM 출장 월별 필터
  filter_trip_employee: "ALL",  // SCM 출장 담당자 필터
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
    leaves: state.leaves,
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
        state.leaves = p.leaves || [];
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
    state.leaves = data.leaves || [];
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
  const data = { employees: state.employees, attendance: state.attendance, trips: state.trips, leaves: state.leaves };
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

// 결과 텍스트 → 상태 분류
function classifyHotelResult(text) {
  const t = (text || "").toLowerCase();
  if (!t) return "PENDING";
  if (/done|setup|set up|signed|confirmed|updated|sign up|go live/.test(t)) return "DONE";
  if (/sent email|requested|reminded|pushed|propos|submitted|received/.test(t)) return "REQUESTED";
  if (/in discussing|in progress|under review|under consider|reviewing|being updated|keep updated|keep follow/.test(t)) return "IN_PROGRESS";
  if (/not support|reject|declined|not agree/.test(t)) return "REJECTED";
  return "PENDING";
}

// Report 시트 → hotels[] 파싱
function parseReportSheet(sheet) {
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "", raw: false });
  if (rows.length < 2) return [];
  // 헤더 행 찾기 (첫 3행 중 "hotel" 포함)
  let headerRow = 0;
  for (let r = 0; r < Math.min(3, rows.length); r++) {
    const joined = rows[r].join(" ").toLowerCase();
    if (joined.includes("hotel") && (joined.includes("contract") || joined.includes("contact") || joined.includes("purpose") || joined.includes("meeting") || joined.includes("follow"))) {
      headerRow = r; break;
    }
  }
  const header = rows[headerRow].map(h => String(h).toLowerCase());
  const findCol = (...kws) => header.findIndex(h => kws.every(kw => h.includes(kw)));
  let col_hotel = header.findIndex(h => h.includes("hotel") && h.includes("name"));
  if (col_hotel < 0) col_hotel = header.findIndex(h => h.includes("hotel"));
  if (col_hotel < 0) col_hotel = 0;
  const col_contract = findCol("contract");
  const col_contact = findCol("contact") >= 0 ? findCol("contact") : findCol("person") >= 0 ? findCol("person") : findCol("pic");
  const col_purpose = findCol("purpose");
  const col_summary = findCol("meeting");
  const col_follow = findCol("follow");
  const col_1w = findCol("week");
  const col_1m = findCol("month");

  const hotels = [];
  for (let r = headerRow + 1; r < rows.length; r++) {
    const row = rows[r];
    const hotel = String(row[col_hotel] || "").trim();
    if (!hotel || hotel.length < 3) continue;
    if (hotel.toLowerCase().startsWith("overall")) continue;
    const h = {
      hotel,
      contract: col_contract >= 0 ? String(row[col_contract] || "").trim() : "",
      contact:  col_contact >= 0 ? String(row[col_contact] || "").trim() : "",
      purpose:  col_purpose >= 0 ? String(row[col_purpose] || "").trim() : "",
      summary:  col_summary >= 0 ? String(row[col_summary] || "").trim() : "",
      followup: col_follow >= 0 ? String(row[col_follow] || "").trim() : "",
      result_1w: col_1w >= 0 ? String(row[col_1w] || "").trim() : "",
      result_1m: col_1m >= 0 ? String(row[col_1m] || "").trim() : "",
    };
    h.status = classifyHotelResult(h.result_1w || h.followup || h.summary);
    hotels.push(h);
  }
  return hotels;
}

// Expense 시트에서 Total 값 추출 (VND)
function parseExpenseTotal(sheet) {
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "", raw: false });
  for (const row of rows) {
    for (let c = 0; c < row.length; c++) {
      const v = String(row[c] || "").toLowerCase().trim();
      if (v === "total" || v === "grand total") {
        for (let c2 = c + 1; c2 < row.length; c2++) {
          const raw = row[c2];
          const num = typeof raw === "number" ? raw : parseInt(String(raw).replace(/[^\d]/g, ""));
          if (!isNaN(num) && num > 100000) return num;
        }
      }
    }
  }
  return null;
}

// 파일명에서 담당자 · 기간 · 목적지 추출
function parseTripFilename(fname) {
  const base = fname.replace(/\.xlsx?$/i, "").replace(/_/g, " ");
  const result = { employee: "", start: "", end: "", destination: "" };

  // 담당자: "- Andy" 또는 "Anna" 등 파일명 뒤쪽 또는 "_Anna" 형태
  const empMatch = base.match(/[-_]\s*([A-Za-z]+)\s*[.!]*\s*$/) || base.match(/\b(Anna|Andy|Mike|Nhi|Lukas|Hanny|Aerum|DK|Slena|Thu|Yoo|Sang)\b/i);
  if (empMatch) result.employee = empMatch[1];

  // 날짜: "13 - 15 Jan", "17.3 - 19.3", "26-28 May 2026", "(12- 14 May 25)" 등
  const monthNames = { jan:"01", feb:"02", mar:"03", apr:"04", may:"05", jun:"06", jul:"07", aug:"08", sep:"09", oct:"10", nov:"11", dec:"12" };
  // Pattern: DD - DD Mon YYYY
  const p1 = base.match(/(\d{1,2})\s*[-–]\s*(\d{1,2})\s+(\w{3})\s*(\d{2,4})?/i);
  if (p1) {
    const d1 = p1[1].padStart(2,"0"), d2 = p1[2].padStart(2,"0");
    const mo = monthNames[p1[3].toLowerCase().slice(0,3)];
    let yr = p1[4] || "2026";
    if (yr.length === 2) yr = "20" + yr;
    if (yr === "2025") yr = "2026";  // 파일명 오타 보정
    if (mo) {
      result.start = `${yr}-${mo}-${d1}`;
      result.end = `${yr}-${mo}-${d2}`;
    }
  }
  // Pattern: DD.M - DD.M (한국식)
  if (!result.start) {
    const p2 = base.match(/(\d{1,2})\.(\d{1,2})\s*[-–]\s*(\d{1,2})\.(\d{1,2})/);
    if (p2) {
      const yr = "2026";
      result.start = `${yr}-${p2[2].padStart(2,"0")}-${p2[1].padStart(2,"0")}`;
      result.end = `${yr}-${p2[4].padStart(2,"0")}-${p2[3].padStart(2,"0")}`;
    }
  }
  // Pattern: (02-03 Jun 2026)
  if (!result.start) {
    const p3 = base.match(/(\d{1,2})[-–](\d{1,2})\s+(\w{3})\s*(\d{4})?/i);
    if (p3) {
      const mo = monthNames[p3[3].toLowerCase().slice(0,3)];
      const yr = p3[4] || "2026";
      if (mo) {
        result.start = `${yr}-${mo}-${p3[1].padStart(2,"0")}`;
        result.end = `${yr}-${mo}-${p3[2].padStart(2,"0")}`;
      }
    }
  }
  // 목적지 힌트: 파일명 앞부분에서 도시 이름 추출
  const cities = ["Hanoi","Ha Noi","HCM","Ho Chi Minh","Da Nang","Danang","Da Lat","Dalat","Quy Nhon","Hoi An","Phan Thiet","Phu Quoc","Ha Long","Halong","Nha Trang"];
  for (const c of cities) {
    if (base.toLowerCase().includes(c.toLowerCase().replace(/\s/g,""))) { result.destination = c.replace(/^Ha Noi$/i,"Hanoi").replace(/^HCM$/i,"Ho Chi Minh").replace(/^Danang$/i,"Da Nang").replace(/^Dalat$/i,"Da Lat").replace(/^Halong$/i,"Ha Long"); break; }
    if (base.includes(c)) { result.destination = c; break; }
  }
  return result;
}

// SCM 직원 리스트에서 매칭
function matchEmployee(nameHint) {
  if (!nameHint) return null;
  const hint = nameHint.toLowerCase().trim();
  const emps = state.employees.filter(e => e.is_scm);
  // 1) 닉네임 완전 일치
  const nickHit = emps.find(e => {
    const nick = nickOnly(e.name).toLowerCase();
    return nick === hint;
  });
  if (nickHit) return nickHit.name;
  // 2) 이름 부분 일치
  const partHit = emps.find(e => e.name.toLowerCase().includes(hint));
  if (partHit) return partHit.name;
  // 3) Lukas 예외 케이스
  if (hint === "lukas") {
    const l = emps.find(e => e.name.includes("Lukas"));
    if (l) return l.name;
  }
  // 4) Sang Yoo / Yoo
  if (["sang", "yoo"].includes(hint)) {
    const y = emps.find(e => e.name.includes("Yoo") || e.name === "Sang Yoo");
    if (y) return y.name;
  }
  return null;
}

async function importTripPlan(file) {
  try {
    const scmEmps = state.employees.filter(e => e.is_scm);
    if (scmEmps.length === 0) { alert("SCM 인원이 먼저 등록되어 있어야 합니다."); return; }

    const wb = await readWorkbook(file);
    const expenseSheet = wb.SheetNames.find(n => /expense|application for expense/i.test(n));
    const reportSheet = wb.SheetNames.find(n => /^report$/i.test(n)) || wb.SheetNames.find(n => /report/i.test(n));

    // ▶ Recap 자동 처리 경로: Report 시트가 있으면 결과보고로 처리
    if (reportSheet) {
      const hotels = parseReportSheet(wb.Sheets[reportSheet]);
      // 호텔이 5개 이상이면 결과보고 확정
      if (hotels.length >= 3) {
        return await importTripRecap(file, wb, hotels, expenseSheet);
      }
    }

    // ▼ 기존 계획서 경로 (Report 없거나 호텔 <3)

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

// ============================================================================
// Leave Type 정의 (연차 신청 유형)
// ============================================================================
const LEAVE_TYPES = {
  "AL":    { label: "Annual Leave",         short: "AL",    days: 1.0, deduct_annual: true,  paid: true,  group: "annual",   color: "#3b82f6" },
  "AL/2":  { label: "Half Annual Leave",    short: "AL/2",  days: 0.5, deduct_annual: true,  paid: true,  group: "annual",   color: "#60a5fa" },
  "UP":    { label: "Unpaid Leave",         short: "UP",    days: 1.0, deduct_annual: false, paid: false, group: "unpaid",   color: "#94a3b8" },
  "UP/2":  { label: "Half Unpaid Leave",    short: "UP/2",  days: 0.5, deduct_annual: false, paid: false, group: "unpaid",   color: "#cbd5e1" },
  "H":     { label: "Holiday",              short: "H",     days: 1.0, deduct_annual: false, paid: true,  group: "paid_pl",  color: "#16a34a" },
  "CL":    { label: "Compensation Leave",   short: "CL",    days: 1.0, deduct_annual: false, paid: true,  group: "paid_pl",  color: "#22c55e" },
  "MR":    { label: "Marriage Leave",       short: "MR",    days: 1.0, deduct_annual: false, paid: true,  group: "paid_pl",  color: "#ec4899" },
  "FL":    { label: "Funeral Leave",        short: "FL",    days: 1.0, deduct_annual: false, paid: true,  group: "paid_pl",  color: "#64748b" },
  "BT":    { label: "Business Trip",        short: "BT",    days: 1.0, deduct_annual: false, paid: true,  group: "paid_pl",  color: "#f59e0b" },
  "SL":    { label: "Sick Leave",           short: "SL",    days: 1.0, deduct_annual: false, paid: false, group: "sick",     color: "#dc2626" },
  "MN":    { label: "Maternity Leave (SI)", short: "MN",    days: 1.0, deduct_annual: false, paid: false, group: "maternity",color: "#f472b6" },
};
function leaveInfo(type) { return LEAVE_TYPES[type] || { label: type, short: type, days: 1, deduct_annual: false, paid: false, group: "other", color: "#94a3b8" }; }

// Auto-generate BT leave records from state.trips (for leave view)
function autoGenerateBTLeaves() {
  const bt = [];
  (state.trips || []).forEach(t => {
    if (!["APPROVED","IN_PROGRESS","COMPLETED"].includes(t.status)) return;
    if (!t.start_date || !t.end_date) return;
    const start = new Date(t.start_date);
    const end = new Date(t.end_date);
    for (let d = new Date(start); d <= end; d.setDate(d.getDate() + 1)) {
      const dstr = d.toISOString().slice(0,10);
      bt.push({
        id: `bt-${t.id}-${dstr}`,
        person_id: (state.employees.find(e => e.name === t.employee) || {}).person_id,
        name: t.employee,
        date: dstr,
        type: "BT",
        days: 1.0,
        status: "APPROVED",
        note: `Trip #${t.id} · ${t.destination || ""}`,
        auto: true,
      });
    }
  });
  return bt;
}

// 담당자별 연차 잔여 계산 (annual - AL - AL/2*0.5)
function calcRemainingLeave(emp) {
  const annual = emp.annual_leave || 0;
  const myLeaves = (state.leaves || []).filter(l => l.person_id === emp.person_id && l.status === "APPROVED");
  const usedAnnual = myLeaves.reduce((s, l) => {
    const info = leaveInfo(l.type);
    return info.deduct_annual ? s + (l.days || info.days) : s;
  }, 0);
  return { annual, usedAnnual, remaining: annual - usedAnnual };
}
// 담당자 이름을 영문 닉네임만 표시 (예: "Le Thi Kim Anh (Aerum)" → "Aerum", "Yoo SangKyu" → "Yoo SangKyu")
function nickOnly(name) {
  const m = (name || "").match(/\(([^)]+)\)/);
  return m ? m[1].trim() : (name || "").trim();
}
// 지역명을 짧게 (예: "Ho Chi Minh, VN" → "Ho Chi Minh", "TBD (Timesheet BT)" → "TBD")
function destShort(dest) {
  if (!dest) return "—";
  let d = dest.replace(", VN", "").trim();
  if (d.startsWith("TBD")) return "TBD";
  return d;
}
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
  { key: "attendance", label: "근태·연차",  icon: "🕘" },
  { key: "trips",      label: "SCM 출장",   icon: "✈️" },
  { key: "calendar",   label: "통합 캘린더", icon: "📅" },
  { key: "reports",    label: "리포트",     icon: "📈" },
];

function go(view) {
  state.view = view;
  state.page_att = 1;
  render();
  const content = document.querySelector(".content");
  if (content) content.scrollTop = 0;
}

// 대시보드 담당자 카드에서 SCM 출장 → 특정 담당자 필터 세팅
function goTripsFor(employeeName) {
  state.filter_trip_month = "ALL";
  state.filter_trip_employee = employeeName;
  state.view = "trips";
  render();
  const content = document.querySelector(".content");
  if (content) content.scrollTop = 0;
}

// 출장 화면에서 담당자 필터 변경
function setTripEmpFilter(name) {
  state.filter_trip_employee = name;
  render();
}

function renderShell() {
  const scm = state.employees.filter(e => e.is_scm).length;
  const active = state.trips.filter(t => ["APPROVED","IN_PROGRESS"].includes(t.status)).length;
  const stats = `직원 ${state.employees.length} (SCM ${scm}) · 근태 ${state.attendance.length.toLocaleString()} · 진행 출장 ${active}`;
  const titleMap = { overview: "대시보드", employees: "VN Office 인원", attendance: "근태·연차", trips: "SCM 출장", calendar: "통합 캘린더", reports: "리포트" };

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
    calendar: viewCalendar,
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

    ${(() => {
      // 이달 SCM 트립 성과 KPI
      const now = new Date();
      const ym = now.toISOString().slice(0,7);
      const thisMonthTrips = state.trips.filter(t => (t.start_date || "").startsWith(ym));
      const allCompleted = state.trips.filter(t => t.status === "COMPLETED");
      const totalHotels = allCompleted.reduce((s,t) => s + (t.hotels?.length || 0), 0);
      const totalExpense = allCompleted.reduce((s,t) => s + (t.expense_vnd || 0), 0);
      const newContracts = allCompleted.reduce((s,t) => s + (t.hotels || []).filter(h => (h.contract || "").toLowerCase().includes("new") || (h.purpose || "").toLowerCase().includes("new account") || (h.purpose || "").toLowerCase().includes("new contract")).length, 0);
      const allHotels = allCompleted.flatMap(t => t.hotels || []);
      const doneCount = allHotels.filter(h => h.status === "DONE").length;
      const inProgCount = allHotels.filter(h => h.status === "IN_PROGRESS").length;
      const followUpRate = allHotels.length > 0 ? (doneCount / allHotels.length * 100).toFixed(0) : "0";
      const kpi = (label, val, sub, color) => `
        <div style="flex:1; min-width:150px; padding:14px; border-radius:10px; background:${color}12; border:1px solid ${color}30;">
          <div style="font-size:11px; color:#64748b; margin-bottom:6px;">${label}</div>
          <div style="font-size:22px; font-weight:700; color:${color};">${val}</div>
          <div style="font-size:11px; color:#94a3b8; margin-top:2px;">${sub}</div>
        </div>`;
      return `
        <div class="card mt-4">
          <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:12px;">
            <h3 style="margin:0;">🎯 SCM 출장 성과 지표</h3>
            <span style="font-size:11px; color:#64748b;">이달 ${thisMonthTrips.length}건 시작 · 전체 완료 ${allCompleted.length}건 기준</span>
          </div>
          <div style="display:flex; gap:10px; flex-wrap:wrap;">
            ${kpi("이달 트립", thisMonthTrips.length + "건", `${ym}`, "#0ea5e9")}
            ${kpi("방문 호텔 (누적)", totalHotels + "곳", `평균 ${allCompleted.length > 0 ? (totalHotels / allCompleted.length).toFixed(0) : 0}곳/트립`, "#4f46e5")}
            ${kpi("신규 계약 시도", newContracts + "건", `New Contract/Account`, "#16a34a")}
            ${kpi("Follow-up 완료율", followUpRate + "%", `Done ${doneCount} / In discussing ${inProgCount}`, "#f59e0b")}
            ${kpi("총 경비", (totalExpense/1000000).toFixed(1) + "M", `VND · 리포트 반영분`, "#dc2626")}
          </div>
        </div>
      `;
    })()}

    ${(() => {
      // 담당자별 출장 요약 (SCM 출장 담당자 기준 · Thu/Slena 제외)
      const scmEmps = state.employees.filter(e => e.is_scm && e.is_scm_traveler !== false);
      const summary = scmEmps.map(e => {
        const empTrips = state.trips.filter(t => t.employee === e.name);
        if (empTrips.length === 0) return null;
        const totalDays = empTrips.reduce((s, t) => {
          const d1 = new Date(t.start_date), d2 = new Date(t.end_date);
          return s + (isNaN(d1) || isNaN(d2) ? 0 : Math.max(1, Math.round((d2 - d1) / 86400000) + 1));
        }, 0);
        const sortedByEnd = [...empTrips].sort((a,b) => (b.end_date||"").localeCompare(a.end_date||""));
        const mostRecent = sortedByEnd[0];
        const completedCount = empTrips.filter(t => t.status === "COMPLETED").length;
        const missingOutcome = empTrips.filter(t => t.status === "COMPLETED" && !t.outcome).length;
        const activeCount = empTrips.filter(t => ["APPROVED","IN_PROGRESS"].includes(t.status)).length;
        return { emp: e, empTrips, totalDays, mostRecent, completedCount, missingOutcome, activeCount };
      }).filter(Boolean).sort((a,b) => b.empTrips.length - a.empTrips.length);

      if (summary.length === 0) return "";

      return `
        <div class="card mt-4">
          <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:12px;">
            <h3 style="margin:0;">📈 담당자별 SCM 출장 요약</h3>
            <span style="font-size:11px; color:#64748b;">${summary.length}명 · 총 ${state.trips.length}건</span>
          </div>
          <div style="display:grid; grid-template-columns:repeat(auto-fill, minmax(220px, 1fr)); gap:12px;">
            ${summary.map(s => {
              const isHead = s.emp.position === "SCM Head";
              return `
                <div style="border:1px solid #e2e8f0; border-radius:10px; padding:12px; background:#fff; cursor:pointer;" onclick='goTripsFor(${JSON.stringify(s.emp.name)})'>
                  <div style="display:flex; align-items:center; gap:6px; margin-bottom:8px;">
                    <span style="font-size:13px; font-weight:600; color:${isHead ? '#4f46e5' : '#0f172a'};">
                      ${isHead ? '👑 ' : ''}${escHTML(s.emp.name.replace(/\\s*\\([^)]+\\)/, ''))}
                    </span>
                  </div>
                  ${(() => {
                    const nick = (s.emp.name.match(/\\(([^)]+)\\)/) || [])[1];
                    return nick ? `<div style="font-size:10px; color:#94a3b8; margin-top:-4px; margin-bottom:6px;">${escHTML(nick)}</div>` : "";
                  })()}
                  <div style="display:flex; justify-content:space-between; font-size:12px; margin-bottom:4px;">
                    <span style="color:#64748b;">총 출장</span>
                    <span><b>${s.empTrips.length}건</b> · ${s.totalDays}일</span>
                  </div>
                  ${s.activeCount > 0 ? `
                    <div style="display:flex; justify-content:space-between; font-size:12px; margin-bottom:4px;">
                      <span style="color:#64748b;">진행 중</span>
                      <span class="badge b-primary">${s.activeCount}건</span>
                    </div>` : ""}
                  <div style="display:flex; justify-content:space-between; font-size:12px; margin-bottom:4px;">
                    <span style="color:#64748b;">최근 출장</span>
                    <span style="color:#0f172a;">${s.mostRecent.end_date}</span>
                  </div>
                  ${s.missingOutcome > 0 ? `
                    <div style="margin-top:8px; padding:4px 8px; background:#fef3c7; border-radius:6px; font-size:11px; color:#92400e; text-align:center;">
                      ⚠️ 결과 미기입 ${s.missingOutcome}건
                    </div>` : (s.completedCount > 0 ? `
                    <div style="margin-top:8px; padding:4px 8px; background:#dcfce7; border-radius:6px; font-size:11px; color:#166534; text-align:center;">
                      ✅ 결과보고 완료
                    </div>` : "")}
                </div>
              `;
            }).join("")}
          </div>
        </div>
      `;
    })()}
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
          <th class="right" style="background:#fef3c7;">부여</th>
          <th class="right" style="background:#fef3c7;">사용</th>
          <th class="right" style="background:#dcfce7;">잔여</th>
        </tr></thead>
        <tbody>
          ${filtered.map(e => {
            const annual = e.annual_leave || 0;
            const remaining = e.remaining_leave || 0;
            const used = Math.max(0, annual - remaining);
            return `
            <tr>
              <td class="mono">${escHTML(e.person_id || "—")}</td>
              <td>
                <b>${escHTML(e.name)}</b>
                ${e.is_scm ? (e.position === "SCM Head" ? `<span class="badge b-scm" style="margin-left:6px; background:#4f46e5; color:#fff;">👑 SCM Head</span>` : `<span class="badge b-scm" style="margin-left:6px;">SCM</span>`) : ""}
              </td>
              <td>${escHTML(e.department || "—")}</td>
              <td>${escHTML(e.position || "—")}</td>
              <td>${escHTML(e.gender || "—")}</td>
              <td class="right" style="background:#fef3c7;"><b>${annual.toFixed(1)}</b></td>
              <td class="right" style="background:#fef3c7;">${used.toFixed(1)}</td>
              <td class="right" style="background:#dcfce7;"><b>${remaining.toFixed(1)}</b></td>
            </tr>`;
          }).join("")}
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
  if (!state.att_tab) state.att_tab = "daily";
  const tabs = [
    { id: "daily",   label: "🗓️ 일별 근태" },
    { id: "leaves",  label: "🏖️ 휴가 이력" },
    { id: "summary", label: "📊 담당자별 요약" },
  ];
  return `
    <div class="flex center gap-3 wrap">
      <h2 style="margin:0; font-size:16px;">근태 · 휴가 관리</h2>
    </div>

    <div class="card mt-3" style="padding:6px; background:#f8fafc; display:flex; gap:4px;">
      ${tabs.map(t => `
        <button style="flex:1; padding:8px; border-radius:6px; border:none; cursor:pointer; background:${state.att_tab === t.id ? '#4f46e5' : 'transparent'}; color:${state.att_tab === t.id ? '#fff' : '#64748b'}; font-weight:${state.att_tab === t.id ? '600' : '400'}; font-size:13px;" onclick="state.att_tab='${t.id}'; state.page_att=1; render();">${t.label}</button>
      `).join("")}
    </div>

    ${state.att_tab === "daily" ? renderAttendanceDaily() : ""}
    ${state.att_tab === "leaves" ? renderLeavesTab() : ""}
    ${state.att_tab === "summary" ? renderAttendanceSummary() : ""}
  `;
}

// ===== 일별 근태 (기존 뷰) =====
function renderAttendanceDaily() {
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

  // Leave 데이터 join: 각 근태 record 의 name+date 로 leaves 매칭 → 상태를 leave type 으로 오버라이드
  const allLeaves = [...(state.leaves || []), ...autoGenerateBTLeaves()];
  const leaveMap = {};
  allLeaves.filter(l => l.status === "APPROVED").forEach(l => {
    leaveMap[`${l.name}|${l.date}`] = l;
  });

  return `
    <div class="mt-3">
      ${dropzone("attendance", "근태 엑셀 (KEYWATCH 형식) 드래그")}
    </div>

    <div class="card mt-3">
      <div class="chip-label">부서</div>
      <div class="chips">
        ${depts.map(d => `<button class="chip ${state.filter_dept === d ? "active" : ""}" onclick="state.filter_dept='${d}'; state.page_att=1; render();">${d === "ALL" ? "전체" : escHTML(d)}</button>`).join("")}
      </div>
      <div style="display:flex; gap:12px; margin-top:10px; flex-wrap:wrap;">
        <div style="flex:1; min-width:180px;">
          <div class="chip-label">월</div>
          <select class="select-filter" onchange="state.filter_month=this.value; state.page_att=1; render();">
            ${months.map(m => `<option value="${m}" ${state.filter_month === m ? "selected" : ""}>${m === "ALL" ? "전체" : m}</option>`).join("")}
          </select>
        </div>
        <div style="flex:1; min-width:180px;">
          <div class="chip-label">상태</div>
          <select class="select-filter" onchange="state.filter_status=this.value; state.page_att=1; render();">
            ${statuses.map(s => `<option value="${s}" ${state.filter_status === s ? "selected" : ""}>${s === "ALL" ? "전체" : s}</option>`).join("")}
          </select>
        </div>
      </div>
    </div>

    <div style="font-size:12px; color:#64748b; margin-top:8px;">
      총 ${filtered.length.toLocaleString()} / ${state.attendance.length.toLocaleString()}건
      <button class="btn btn-outline btn-sm ml-3" onclick="exportSheet('attendance')">📤 필터 결과 내보내기</button>
    </div>

    <div class="card mt-3" style="padding:0;">
      <div class="table-wrap"><table>
        <thead><tr>
          <th>날짜</th><th>이름</th><th>부서</th><th>출근</th><th>퇴근</th>
          <th class="right">지각(분)</th><th>상태 / 휴가</th>
        </tr></thead>
        <tbody>
          ${paginated.map(a => {
            const leaveMatched = leaveMap[`${a.name}|${a.date}`];
            let statusHtml;
            if (leaveMatched) {
              const info = leaveInfo(leaveMatched.type);
              statusHtml = `<span style="display:inline-block; padding:2px 8px; border-radius:4px; background:${info.color}20; color:${info.color}; font-size:11px; font-weight:600;" title="${escHTML(info.label)}${leaveMatched.note ? ' · '+escHTML(leaveMatched.note) : ''}">${leaveMatched.type}</span>`;
            } else {
              statusHtml = `<span class="badge b-${a.status === "LATE" ? "warn" : a.status === "ABSENT" ? "danger" : "success"}">${a.status}</span>`;
            }
            return `
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
                <td>${statusHtml}</td>
              </tr>`;
          }).join("")}
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

// ===== 휴가 이력 탭 =====
function renderLeavesTab() {
  const allLeaves = [...(state.leaves || []), ...autoGenerateBTLeaves()];
  const months = ["ALL", ...new Set(allLeaves.map(l => (l.date || "").slice(0,7)).filter(Boolean))].sort().reverse();
  const types = ["ALL", ...Object.keys(LEAVE_TYPES)];
  const depts = ["ALL", ...new Set(state.employees.map(e => e.department).filter(Boolean))].sort();

  if (!months.includes(state.filter_month)) state.filter_month = "ALL";
  const activeType = state.filter_leave_type || "ALL";
  if (!state.filter_dept || !depts.includes(state.filter_dept)) state.filter_dept = "ALL";

  // 부서 lookup (name -> dept)
  const empDept = {};
  state.employees.forEach(e => { empDept[e.name] = e.department || ""; });

  let filtered = allLeaves.filter(l => {
    if (state.filter_month !== "ALL" && !(l.date || "").startsWith(state.filter_month)) return false;
    if (activeType !== "ALL" && l.type !== activeType) return false;
    if (state.filter_dept !== "ALL" && empDept[l.name] !== state.filter_dept) return false;
    return true;
  });
  filtered.sort((a,b) => (b.date || "").localeCompare(a.date || ""));

  // 통계: 타입별 카운트 (부서 필터 반영)
  const typeCounts = {};
  Object.keys(LEAVE_TYPES).forEach(t => typeCounts[t] = 0);
  allLeaves.filter(l => l.status === "APPROVED" && (state.filter_dept === "ALL" || empDept[l.name] === state.filter_dept)).forEach(l => {
    if (typeCounts[l.type] !== undefined) typeCounts[l.type] += (l.days || leaveInfo(l.type).days);
  });

  return `
    <div class="card mt-3" style="background:#eff6ff; border-color:#bfdbfe; padding:10px 14px; font-size:12px; color:#1e3a8a;">
      ℹ️ 휴가는 <b>AL / AL/2</b> 만 연차 잔여에서 차감됩니다. <b>PL(H/CL/MR/FL/BT)</b> 은 유급, <b>UP/UP/2/SL/MN</b> 은 무급/특수. BT 는 SCM 트립에서 자동 반영됩니다.
    </div>

    <div class="card mt-3">
      <div style="display:flex; gap:6px; flex-wrap:wrap; margin-bottom:8px;">
        ${Object.entries(LEAVE_TYPES).map(([t, info]) => `
          <div style="padding:6px 10px; border-radius:6px; background:${info.color}15; border:1px solid ${info.color}40; font-size:11px;">
            <span style="color:${info.color}; font-weight:700;">${t}</span>
            <span style="color:#64748b; margin-left:4px;">${info.label}</span>
            <span style="color:${info.color}; font-weight:600; margin-left:6px;">${typeCounts[t]}</span>
          </div>
        `).join("")}
      </div>
    </div>

    <div class="card mt-3">
      <div class="chip-label">부서</div>
      <div class="chips">
        ${depts.map(d => `<button class="chip ${state.filter_dept === d ? "active" : ""}" onclick="state.filter_dept='${escHTML(d)}'; render();">${d === "ALL" ? "전체" : escHTML(d)}</button>`).join("")}
      </div>
      <div style="display:flex; gap:12px; margin-top:10px; flex-wrap:wrap;">
        <div style="flex:1; min-width:180px;">
          <div class="chip-label">월</div>
          <select class="select-filter" onchange="state.filter_month=this.value; render();">
            ${months.map(m => `<option value="${m}" ${state.filter_month === m ? "selected" : ""}>${m === "ALL" ? "전체" : m}</option>`).join("")}
          </select>
        </div>
        <div style="flex:1; min-width:180px;">
          <div class="chip-label">유형</div>
          <select class="select-filter" onchange="state.filter_leave_type=this.value; render();">
            ${types.map(t => `<option value="${t}" ${activeType === t ? "selected" : ""}>${t === "ALL" ? "전체" : `${t} · ${leaveInfo(t).label}`}</option>`).join("")}
          </select>
        </div>
      </div>
    </div>

    ${(() => {
      // 담당자별 연차 사용 요약 (부서 필터 반영)
      const empList = state.filter_dept === "ALL" ? state.employees : state.employees.filter(e => e.department === state.filter_dept);
      const perEmp = empList.map(e => {
        const myLeaves = allLeaves.filter(l => l.status === "APPROVED" && l.name === e.name);
        const annual = e.annual_leave || 0;
        const usedAL = myLeaves.filter(l => l.type === "AL").reduce((s,l) => s + (l.days || 1), 0);
        const usedAL2 = myLeaves.filter(l => l.type === "AL/2").reduce((s,l) => s + (l.days || 0.5), 0);
        const usedAnnual = usedAL + usedAL2;
        const usedUP = myLeaves.filter(l => l.type === "UP" || l.type === "UP/2").reduce((s,l) => s + (l.days || leaveInfo(l.type).days), 0);
        const usedSL = myLeaves.filter(l => l.type === "SL").reduce((s,l) => s + (l.days || 1), 0);
        const usedBT = myLeaves.filter(l => l.type === "BT").reduce((s,l) => s + (l.days || 1), 0);
        const usedOther = myLeaves.filter(l => ["H","CL","MR","FL","MN"].includes(l.type)).reduce((s,l) => s + (l.days || 1), 0);
        // 실제 이력 기반 잔여
        const remaining = (state.leaves && state.leaves.length > 0) ? annual - usedAnnual : (e.remaining_leave ?? 0);
        return { e, annual, usedAnnual, usedAL, usedAL2, usedUP, usedSL, usedBT, usedOther, remaining };
      }).filter(r => r.usedAnnual > 0 || r.usedBT > 0 || r.usedUP > 0 || r.usedSL > 0 || r.usedOther > 0);
      perEmp.sort((a,b) => b.usedAnnual - a.usedAnnual);
      if (perEmp.length === 0) return "";

      return `
        <div class="card mt-3" style="padding:0;">
          <div style="padding:10px 14px; border-bottom:1px solid #e2e8f0; background:#f8fafc; font-size:12px; font-weight:600;">
            📋 담당자별 사용 요약 <span style="color:#64748b; font-weight:400;">(부서: ${state.filter_dept === "ALL" ? "전체" : state.filter_dept} · ${perEmp.length}명)</span>
          </div>
          <div class="table-wrap"><table>
            <thead><tr>
              <th>담당자</th><th>부서</th>
              <th class="right" style="background:#eff6ff;">AL</th>
              <th class="right" style="background:#eff6ff;">AL/2</th>
              <th class="right">UP</th>
              <th class="right">SL</th>
              <th class="right">BT</th>
              <th class="right">기타 PL</th>
              <th class="right" style="background:#fef3c7;">부여</th>
              <th class="right" style="background:#fef3c7;">사용<br/><small>(연차)</small></th>
              <th class="right" style="background:#dcfce7;">잔여</th>
            </tr></thead>
            <tbody>
              ${perEmp.map(r => `
                <tr>
                  <td><b>${escHTML(nickOnly(r.e.name))}</b>${r.e.is_scm ? `<span class="badge b-scm" style="margin-left:6px;">SCM</span>` : ""}</td>
                  <td class="mono">${escHTML(r.e.department || "—")}</td>
                  <td class="right" style="background:#eff6ff;">${r.usedAL || "—"}</td>
                  <td class="right" style="background:#eff6ff;">${r.usedAL2 || "—"}</td>
                  <td class="right">${r.usedUP || "—"}</td>
                  <td class="right">${r.usedSL || "—"}</td>
                  <td class="right">${r.usedBT || "—"}</td>
                  <td class="right">${r.usedOther || "—"}</td>
                  <td class="right" style="background:#fef3c7;"><b>${r.annual.toFixed(1)}</b></td>
                  <td class="right" style="background:#fef3c7;">${r.usedAnnual.toFixed(1)}</td>
                  <td class="right" style="background:#dcfce7;"><b>${r.remaining.toFixed(1)}</b></td>
                </tr>`).join("")}
            </tbody>
          </table></div>
        </div>`;
    })()}

    <div class="card mt-3" style="padding:0;">
      <div class="table-wrap"><table>
        <thead><tr>
          <th>날짜</th><th>담당자</th><th>부서</th><th>유형</th><th class="right">일수</th><th>상태</th><th>메모</th>
        </tr></thead>
        <tbody>
          ${filtered.length === 0 ? `<tr><td colspan="7" style="text-align:center; padding:24px; color:#94a3b8;">이력 없음</td></tr>` :
            filtered.slice(0, 200).map(l => {
              const info = leaveInfo(l.type);
              const emp = state.employees.find(e => e.name === l.name);
              const dept = emp?.department || "—";
              return `
                <tr>
                  <td>${l.date}</td>
                  <td><b>${escHTML(nickOnly(l.name))}</b></td>
                  <td class="mono">${escHTML(dept)}</td>
                  <td><span style="display:inline-block; padding:2px 8px; border-radius:4px; background:${info.color}20; color:${info.color}; font-size:11px; font-weight:600;" title="${escHTML(info.label)}">${l.type}</span></td>
                  <td class="right">${l.days || info.days}</td>
                  <td><span class="badge b-${l.status === "APPROVED" ? "success" : l.status === "REQUESTED" ? "warn" : "muted"}">${l.status || "—"}</span></td>
                  <td style="font-size:11px; color:#64748b;">${escHTML(l.note || "")}${l.auto ? '<span class="badge b-muted" style="margin-left:4px; font-size:9px;">AUTO</span>' : ""}</td>
                </tr>`;
            }).join("")}
        </tbody>
      </table></div>
      ${filtered.length > 200 ? `<div style="padding:8px 12px; font-size:11px; color:#94a3b8; text-align:center;">200건 표시 · 필터로 좁혀서 보세요</div>` : ""}
    </div>
  `;
}

// ===== 담당자별 요약 탭 =====
function renderAttendanceSummary() {
  const allLeaves = [...(state.leaves || []), ...autoGenerateBTLeaves()];
  const depts = ["ALL", ...new Set(state.employees.map(e => e.department).filter(Boolean))].sort();
  if (!state.filter_dept || !depts.includes(state.filter_dept)) state.filter_dept = "ALL";
  const empList = state.filter_dept === "ALL" ? state.employees : state.employees.filter(e => e.department === state.filter_dept);
  const rows = empList.map(e => {
    const att = state.attendance.filter(a => a.person_id === e.person_id);
    const late = att.filter(a => a.status === "LATE").length;
    const absent = att.filter(a => a.status === "ABSENT").length;
    const my = allLeaves.filter(l => l.person_id === e.person_id && l.status === "APPROVED");
    const byType = {};
    Object.keys(LEAVE_TYPES).forEach(t => byType[t] = 0);
    my.forEach(l => { if (byType[l.type] !== undefined) byType[l.type] += (l.days || leaveInfo(l.type).days); });
    const usedAnnual = byType["AL"] + byType["AL/2"];
    const remaining = (e.annual_leave || 0) - usedAnnual;
    // fallback: baseline remaining if no leaves data yet
    const displayRemaining = (state.leaves && state.leaves.length > 0) ? remaining : (e.remaining_leave ?? 0);
    return { e, att, late, absent, byType, usedAnnual, remaining: displayRemaining };
  }).sort((a,b) => (b.late+b.absent) - (a.late+a.absent));

  return `
    <div class="card mt-3">
      <div class="chip-label">부서</div>
      <div class="chips">
        ${depts.map(d => `<button class="chip ${state.filter_dept === d ? "active" : ""}" onclick="state.filter_dept='${escHTML(d)}'; render();">${d === "ALL" ? "전체" : escHTML(d)} ${d !== "ALL" ? `(${state.employees.filter(x => x.department === d).length})` : ""}</button>`).join("")}
      </div>
    </div>

    <div class="card mt-3" style="background:#eff6ff; border-color:#bfdbfe; padding:10px 14px; font-size:12px; color:#1e3a8a;">
      ℹ️ <b>잔여 연차</b> = 부여 - (AL + AL/2 × 0.5). 휴가 이력이 아직 없으면 기존 baseline 값(2026-06 기준) 을 표시. 표시 인원: <b>${rows.length}명</b>
    </div>

    <div class="card mt-3" style="padding:0;">
      <div class="table-wrap"><table>
        <thead><tr>
          <th>담당자</th><th>부서</th>
          <th class="right">근태</th><th class="right">지각</th><th class="right">결근</th>
          <th class="right" style="background:#eff6ff;">AL</th>
          <th class="right" style="background:#eff6ff;">AL/2</th>
          <th class="right">UP</th><th class="right">SL</th><th class="right">BT</th>
          <th class="right">기타 PL</th>
          <th class="right" style="background:#fef3c7;">부여</th>
          <th class="right" style="background:#fef3c7;">사용</th>
          <th class="right" style="background:#dcfce7;">잔여</th>
        </tr></thead>
        <tbody>
          ${rows.map(r => {
            const otherPL = r.byType["H"] + r.byType["CL"] + r.byType["MR"] + r.byType["FL"];
            return `
              <tr>
                <td><b>${escHTML(nickOnly(r.e.name))}</b>${r.e.is_scm ? `<span class="badge b-scm" style="margin-left:6px;">SCM</span>` : ""}</td>
                <td class="mono">${escHTML(r.e.department || "—")}</td>
                <td class="right">${r.att.length}</td>
                <td class="right ${r.late > 3 ? "text-late" : ""}">${r.late || "—"}</td>
                <td class="right ${r.absent > 1 ? "text-absent" : ""}">${r.absent || "—"}</td>
                <td class="right" style="background:#eff6ff;">${r.byType["AL"] || "—"}</td>
                <td class="right" style="background:#eff6ff;">${r.byType["AL/2"] || "—"}</td>
                <td class="right">${(r.byType["UP"]+r.byType["UP/2"]) || "—"}</td>
                <td class="right">${r.byType["SL"] || "—"}</td>
                <td class="right">${r.byType["BT"] || "—"}</td>
                <td class="right">${otherPL || "—"}</td>
                <td class="right" style="background:#fef3c7;"><b>${(r.e.annual_leave || 0).toFixed(1)}</b></td>
                <td class="right" style="background:#fef3c7;">${r.usedAnnual.toFixed(1)}</td>
                <td class="right" style="background:#dcfce7;"><b>${r.remaining.toFixed(1)}</b></td>
              </tr>`;
          }).join("")}
        </tbody>
      </table></div>
    </div>
  `;
}

// ==========================================================================
// 출장 결과보고 (Recap) 자동 임포트: Report + Expense 파싱 → 기존 트립 병합 or 신규 생성
// ==========================================================================
async function importTripRecap(file, wb, hotels, expenseSheet) {
  const parsed = parseTripFilename(file.name);
  const empMatched = matchEmployee(parsed.employee);
  const expenseVnd = expenseSheet ? parseExpenseTotal(wb.Sheets[expenseSheet]) : null;

  // 통계
  const done = hotels.filter(h => h.status === "DONE").length;
  const inprog = hotels.filter(h => h.status === "IN_PROGRESS").length;
  const req = hotels.filter(h => h.status === "REQUESTED").length;
  const pend = hotels.filter(h => h.status === "PENDING").length;
  const rej = hotels.filter(h => h.status === "REJECTED").length;
  const newContracts = hotels.filter(h => /new/i.test((h.contract || "") + (h.purpose || ""))).length;
  const outcome = `방문 호텔 ${hotels.length}곳 · Done ${done} · In discussing ${inprog} · Requested ${req} · Rejected ${rej} · Pending ${pend} · 신규계약 시도 ${newContracts}`;

  // 기존 트립 매칭 (담당자 + 시작일 ±2일)
  let matchedTrip = null;
  if (empMatched && parsed.start) {
    const startD = new Date(parsed.start);
    matchedTrip = state.trips.find(t => {
      if (t.employee !== empMatched) return false;
      if (!t.start_date) return false;
      const ts = new Date(t.start_date);
      const diff = Math.abs((ts - startD) / 86400000);
      return diff <= 2;
    });
  }

  const scmEmps = state.employees.filter(e => e.is_scm && e.is_scm_traveler !== false);
  const empOpts = scmEmps.map(e =>
    `<option value="${escHTML(e.name)}" ${(matchedTrip?.employee || empMatched) === e.name ? "selected" : ""}>${escHTML(nickOnly(e.name))} — ${escHTML(e.name)}</option>`
  ).join("");

  const modeLabel = matchedTrip ? `🔗 기존 트립 병합 (#${matchedTrip.id})` : "🆕 신규 트립 생성";
  const modeColor = matchedTrip ? "#4f46e5" : "#16a34a";

  showModal("🏨 출장 결과보고 자동 인식", `
    <div style="padding:10px 14px; background:${modeColor}15; border:1px solid ${modeColor}40; border-radius:8px; margin-bottom:12px;">
      <div style="font-size:12px; color:${modeColor}; font-weight:700;">${modeLabel}</div>
      <div style="font-size:11px; color:#64748b; margin-top:2px;">${escHTML(file.name)}</div>
    </div>

    <div class="grid-2">
      <div><label class="field-label">담당자</label><select id="rc_emp"><option value="">선택</option>${empOpts}</select></div>
      <div><label class="field-label">목적지</label><input id="rc_dest" value="${escHTML(matchedTrip?.destination || parsed.destination ? (parsed.destination + ", VN") : "")}" /></div>
    </div>
    <div class="grid-2">
      <div><label class="field-label">시작일</label><input id="rc_start" type="date" value="${matchedTrip?.start_date || parsed.start || ""}" /></div>
      <div><label class="field-label">종료일</label><input id="rc_end" type="date" value="${matchedTrip?.end_date || parsed.end || ""}" /></div>
    </div>
    <div class="grid-2">
      <div><label class="field-label">경비 (VND)</label><input id="rc_exp" type="number" value="${expenseVnd || matchedTrip?.expense_vnd || 0}" /></div>
      <div><label class="field-label">상태</label><select id="rc_status">${["COMPLETED","IN_PROGRESS","APPROVED"].map(s => `<option ${(matchedTrip?.status || "COMPLETED")===s?"selected":""}>${s}</option>`).join("")}</select></div>
    </div>

    <div style="margin-top:12px; padding:10px; background:#f8fafc; border-radius:8px; border:1px solid #e2e8f0;">
      <div style="font-size:12px; font-weight:700; margin-bottom:6px;">📊 파싱 요약</div>
      <div style="font-size:11px; color:#475569;">
        <div>🏨 방문 호텔: <b>${hotels.length}곳</b></div>
        <div style="margin-top:4px; display:flex; gap:6px; flex-wrap:wrap;">
          <span style="padding:2px 8px; border-radius:4px; background:#16a34a20; color:#16a34a; font-weight:600;">Done ${done}</span>
          <span style="padding:2px 8px; border-radius:4px; background:#3b82f620; color:#3b82f6; font-weight:600;">In discussing ${inprog}</span>
          <span style="padding:2px 8px; border-radius:4px; background:#f59e0b20; color:#f59e0b; font-weight:600;">Requested ${req}</span>
          <span style="padding:2px 8px; border-radius:4px; background:#dc262620; color:#dc2626; font-weight:600;">Rejected ${rej}</span>
          <span style="padding:2px 8px; border-radius:4px; background:#94a3b820; color:#94a3b8; font-weight:600;">Pending ${pend}</span>
        </div>
        <div style="margin-top:6px;">🆕 신규 계약 시도: <b>${newContracts}건</b>${expenseVnd ? ` · 💰 ${expenseVnd.toLocaleString()} VND` : ""}</div>
      </div>
      <div style="margin-top:8px; font-size:10px; color:#94a3b8; max-height:80px; overflow-y:auto;">
        ${hotels.slice(0,10).map(h => escHTML(h.hotel)).join(" · ")}${hotels.length > 10 ? ` … 외 ${hotels.length-10}` : ""}
      </div>
    </div>
  `, () => {
    const emp = val("rc_emp");
    if (!emp) { alert("담당자를 선택하세요."); return false; }
    const dest = val("rc_dest");
    const start = val("rc_start");
    const end = val("rc_end");
    const exp = +val("rc_exp") || null;
    const status = val("rc_status");
    if (matchedTrip) {
      matchedTrip.employee = emp;
      matchedTrip.destination = dest;
      matchedTrip.start_date = start;
      matchedTrip.end_date = end;
      matchedTrip.status = status;
      matchedTrip.hotels = hotels;
      matchedTrip.expense_vnd = exp;
      matchedTrip.cost_planned = exp || 0;
      matchedTrip.cost_actual = exp || 0;
      matchedTrip.source_file = file.name;
      matchedTrip.outcome = outcome;
    } else {
      const nextId = state.trips.length ? Math.max(...state.trips.map(t => t.id)) + 1 : 1;
      const destShortName = (dest || "Biz Trip").replace(/,.*$/,"").trim();
      state.trips.push({
        id: nextId,
        title: `${destShortName} · ${start} (${end && start ? Math.round((new Date(end)-new Date(start))/86400000)+1 : 1}일)`,
        employee: emp,
        destination: dest,
        start_date: start, end_date: end,
        purpose: "SCM Business Trip",
        status,
        cost_planned: exp || 0, cost_actual: exp || 0, currency: "VND",
        partners: [], itinerary: [],
        hotels, expense_vnd: exp,
        source_file: file.name,
        outcome,
        roi: null, notes: "",
      });
    }
    save(); render();
    setTimeout(() => alert(`${matchedTrip ? "결과보고 병합" : "결과보고 신규 등록"} 완료 · 방문 호텔 ${hotels.length}곳`), 100);
    return true;
  });
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
// SCM 출장 Analytics (전체 필터일 때만 표시)
// ==========================================================================
function renderTripAnalytics() {
  const trips = state.trips;
  const totalTrips = trips.length;
  const completed = trips.filter(t => t.status === "COMPLETED");
  const allHotels = completed.flatMap(t => t.hotels || []);
  const totalHotels = allHotels.length;
  const totalExpense = completed.reduce((s,t) => s + (t.expense_vnd || 0), 0);
  const newContracts = allHotels.filter(h => ((h.contract||"") + (h.purpose||"")).toLowerCase().includes("new")).length;
  const doneCnt = allHotels.filter(h => h.status === "DONE").length;
  const fuRate = allHotels.length > 0 ? Math.round(doneCnt / allHotels.length * 100) : 0;

  const kpi = (label, val, sub, color) => `
    <div style="flex:1; min-width:130px; padding:12px; border-radius:10px; background:${color}12; border:1px solid ${color}30;">
      <div style="font-size:11px; color:#64748b; margin-bottom:4px;">${label}</div>
      <div style="font-size:20px; font-weight:700; color:${color};">${val}</div>
      <div style="font-size:10px; color:#94a3b8; margin-top:2px;">${sub}</div>
    </div>`;

  // 담당자별 트립 수 + 호텔 수
  const byEmp = {};
  trips.forEach(t => {
    const e = t.employee || "?";
    if (!byEmp[e]) byEmp[e] = { trips: 0, hotels: 0, expense: 0 };
    byEmp[e].trips++;
    byEmp[e].hotels += (t.hotels || []).length;
    byEmp[e].expense += (t.expense_vnd || 0);
  });
  const empRows = Object.entries(byEmp).sort((a,b) => b[1].hotels - a[1].hotels || b[1].trips - a[1].trips);
  const maxHotels = Math.max(1, ...empRows.map(r => r[1].hotels));

  // 지역별 트립 수
  const byDest = {};
  trips.forEach(t => {
    let d = (t.destination || "TBD").replace(", VN", "").trim();
    if (d.startsWith("TBD")) d = "TBD (Timesheet)";
    byDest[d] = (byDest[d] || 0) + 1;
  });
  const destRows = Object.entries(byDest).sort((a,b) => b[1] - a[1]);
  const maxDest = Math.max(1, ...destRows.map(r => r[1]));

  // 월 목록 (오름차순)
  const months = [...new Set(trips.map(t => (t.start_date||"").slice(0,7)).filter(Boolean))].sort();

  // 담당자 × 월 매트릭스: 트립 수 + 호텔 수 + 방문 지역 리스트
  const matrix = {};
  trips.forEach(t => {
    const emp = t.employee || "?";
    const m = (t.start_date||"").slice(0,7);
    if (!m) return;
    if (!matrix[emp]) matrix[emp] = {};
    if (!matrix[emp][m]) matrix[emp][m] = { trips: 0, hotels: 0, dests: [] };
    matrix[emp][m].trips++;
    matrix[emp][m].hotels += (t.hotels || []).length;
    const d = destShort(t.destination);
    if (d && !matrix[emp][m].dests.includes(d)) matrix[emp][m].dests.push(d);
  });

  // 담당자 리스트 (트립 수 기준 정렬, SCM Head 우선)
  const empList = Object.keys(matrix).sort((a,b) => {
    const ea = state.employees.find(e => e.name === a);
    const eb = state.employees.find(e => e.name === b);
    const aHead = ea && ea.position === "SCM Head" ? 1 : 0;
    const bHead = eb && eb.position === "SCM Head" ? 1 : 0;
    if (aHead !== bHead) return bHead - aHead;
    return (byEmp[b]?.trips || 0) - (byEmp[a]?.trips || 0);
  });

  const monthTotals = {};
  months.forEach(m => {
    monthTotals[m] = Object.values(matrix).reduce((s, empData) => s + (empData[m]?.trips || 0), 0);
  });

  const cellBg = (cnt) => {
    if (cnt === 0) return "background:#f8fafc; color:#cbd5e1;";
    if (cnt === 1) return "background:#dbeafe; color:#1e40af; font-weight:600;";
    if (cnt === 2) return "background:#93c5fd; color:#1e3a8a; font-weight:700;";
    return "background:#3b82f6; color:#fff; font-weight:700;";
  };

  const empShort = (name) => nickOnly(name);
  const nickOf = (name) => (name.match(/\(([^)]+)\)/) || [])[1] || "";

  return `
    <div class="card mt-4" style="background:#eff6ff; border-color:#bfdbfe;">
      <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:10px;">
        <h3 style="margin:0; color:#1e40af;">📊 SCM 출장 Analytics</h3>
        <span style="font-size:11px; color:#64748b;">전체 뷰 · 상세는 담당자/월 필터 선택</span>
      </div>
      <div style="display:flex; gap:8px; flex-wrap:wrap;">
        ${kpi("전체 트립", totalTrips + "건", `완료 ${completed.length} / 예정 ${totalTrips - completed.length}`, "#0ea5e9")}
        ${kpi("방문 호텔", totalHotels + "곳", `평균 ${completed.length > 0 ? Math.round(totalHotels / completed.length) : 0}곳/트립`, "#4f46e5")}
        ${kpi("신규 계약 시도", newContracts + "건", `New Contract/Account`, "#16a34a")}
        ${kpi("Follow-up 완료율", fuRate + "%", `Done ${doneCnt} / ${allHotels.length}`, "#f59e0b")}
        ${kpi("총 경비", (totalExpense/1000000).toFixed(1) + "M", "VND · 리포트 반영분", "#dc2626")}
      </div>
    </div>

    <div class="grid-2 mt-3">
      <div class="card">
        <h3>🌍 방문 지역 분포</h3>
        <div style="padding:8px 12px;">
          ${destRows.map(([d, cnt]) => {
            const pct = Math.round(cnt / maxDest * 100);
            const isTBD = d.startsWith("TBD");
            const color = isTBD ? "#94a3b8" : "#3b82f6";
            return `
              <div style="display:flex; align-items:center; gap:10px; margin-bottom:8px; font-size:12px; cursor:pointer;" title="${escHTML(d)} · ${cnt}건">
                <div style="width:100px; font-weight:${isTBD ? '400' : '500'}; color:${isTBD ? '#94a3b8' : '#0f172a'};">${escHTML(d)}</div>
                <div style="flex:1; height:16px; background:#f1f5f9; border-radius:4px; overflow:hidden;">
                  <div style="width:${pct}%; height:100%; background:${color}; border-radius:4px;"></div>
                </div>
                <div style="width:36px; text-align:right; font-weight:600;">${cnt}</div>
              </div>`;
          }).join("")}
        </div>
      </div>

      <div class="card">
        <h3>🏨 담당자별 방문 호텔 수 (누적)</h3>
        <div style="padding:8px 12px;">
          ${empRows.map(([emp, s]) => {
            const pct = Math.round(s.hotels / maxHotels * 100);
            const shortName = nickOnly(emp);
            const isHead = state.employees.find(e => e.name === emp)?.position === "SCM Head";
            return `
              <div style="display:flex; align-items:center; gap:10px; margin-bottom:8px; font-size:12px; cursor:pointer;" onclick='setTripEmpFilter(${JSON.stringify(emp)})' title="${escHTML(emp)} · ${s.trips}건 · ${s.hotels}곳 · 클릭하면 필터">
                <div style="width:110px; font-weight:500; color:${isHead ? '#4f46e5' : '#0f172a'};">${isHead ? '👑 ' : ''}${escHTML(shortName)}</div>
                <div style="flex:1; height:16px; background:#f1f5f9; border-radius:4px; overflow:hidden;">
                  <div style="width:${pct}%; height:100%; background:#4f46e5; border-radius:4px;"></div>
                </div>
                <div style="width:60px; text-align:right; font-weight:600; font-size:11px;">${s.trips}건·${s.hotels}곳</div>
              </div>`;
          }).join("")}
        </div>
      </div>
    </div>

    <div class="card mt-3">
      <h3>🗓️ 담당자 × 월 매트릭스 <span style="font-size:11px; color:#64748b; font-weight:400;">(셀 클릭 시 해당 담당자·월 필터 · 색이 진할수록 트립 수 많음)</span></h3>
      <div style="overflow-x:auto;">
        <table style="width:100%; border-collapse:collapse; font-size:12px; min-width:700px;">
          <thead>
            <tr style="background:#f8fafc;">
              <th style="text-align:left; padding:8px 10px; border-bottom:2px solid #e2e8f0; position:sticky; left:0; background:#f8fafc; z-index:1;">담당자</th>
              ${months.map(m => `<th style="text-align:center; padding:8px 6px; border-bottom:2px solid #e2e8f0; width:90px; font-weight:600;">${m.slice(5)}월</th>`).join("")}
              <th style="text-align:center; padding:8px 10px; border-bottom:2px solid #e2e8f0; background:#eff6ff;">합계</th>
              <th style="text-align:center; padding:8px 10px; border-bottom:2px solid #e2e8f0; background:#eff6ff;">호텔</th>
            </tr>
          </thead>
          <tbody>
            ${empList.map(emp => {
              const isHead = state.employees.find(e => e.name === emp)?.position === "SCM Head";
              const shortName = nickOnly(emp);
              const nick = nickOf(emp);
              const totalTr = byEmp[emp]?.trips || 0;
              const totalHo = byEmp[emp]?.hotels || 0;
              return `
                <tr style="border-bottom:1px solid #f1f5f9;">
                  <td style="padding:6px 10px; position:sticky; left:0; background:#fff; z-index:1; cursor:pointer;" onclick='setTripEmpFilter(${JSON.stringify(emp)})' title="담당자 필터">
                    <div style="font-weight:600; color:${isHead ? '#4f46e5' : '#0f172a'};">${isHead ? '👑 ' : ''}${escHTML(shortName)}</div>
                    ${nick ? `<div style="font-size:10px; color:#94a3b8;">${escHTML(nick)}</div>` : ""}
                  </td>
                  ${months.map(m => {
                    const cell = matrix[emp]?.[m];
                    const cnt = cell?.trips || 0;
                    const ho = cell?.hotels || 0;
                    const dests = cell?.dests || [];
                    const destStr = dests.join(", ");
                    return `<td style="text-align:center; padding:0; border:1px solid #f1f5f9;">
                      <div style="${cellBg(cnt)} padding:6px 4px; cursor:pointer; font-size:11px; min-height:44px; display:flex; flex-direction:column; justify-content:center; align-items:center;" onclick='setTripCellFilter(${JSON.stringify(emp)}, ${JSON.stringify(m)})' title="${escHTML(shortName)} · ${m} · ${cnt}건 · ${ho}곳 · ${escHTML(destStr)}">
                        ${cnt > 0 ? `
                          <div style="font-weight:700;">${cnt}건${ho > 0 ? ` · ${ho}곳` : ""}</div>
                          ${destStr ? `<div style="font-size:9px; opacity:0.85; margin-top:2px; line-height:1.1;">${escHTML(destStr)}</div>` : ""}
                        ` : '·'}
                      </div>
                    </td>`;
                  }).join("")}
                  <td style="text-align:center; padding:6px; background:#eff6ff; font-weight:700; color:#1e40af;">${totalTr}건</td>
                  <td style="text-align:center; padding:6px; background:#eff6ff; font-weight:600; color:#4f46e5;">${totalHo > 0 ? totalHo + '곳' : '—'}</td>
                </tr>`;
            }).join("")}
            <tr style="background:#eff6ff; font-weight:700; border-top:2px solid #bfdbfe;">
              <td style="padding:8px 10px; position:sticky; left:0; background:#eff6ff;">합계</td>
              ${months.map(m => `<td style="text-align:center; padding:6px; color:#1e40af;">${monthTotals[m]}</td>`).join("")}
              <td style="text-align:center; padding:6px; color:#1e40af;">${totalTrips}</td>
              <td style="text-align:center; padding:6px; color:#4f46e5;">${totalHotels}</td>
            </tr>
          </tbody>
        </table>
      </div>
    </div>
  `;
}

// 매트릭스 셀 클릭 → 담당자+월 동시 필터
function setTripCellFilter(emp, month) {
  state.filter_trip_employee = emp;
  state.filter_trip_month = month;
  render();
  const content = document.querySelector(".content");
  if (content) content.scrollTop = 0;
}

// ==========================================================================
// View: Trips
// ==========================================================================
function viewTrips() {
  const cols = ["DRAFT","REQUESTED","APPROVED","IN_PROGRESS","COMPLETED"];
  const scmCount = state.employees.filter(e => e.is_scm && e.is_scm_traveler !== false).length;

  // 트립 월 목록 (출장 시작월 기준)
  const tripMonths = [...new Set(state.trips.map(t => (t.start_date || "").slice(0, 7)).filter(Boolean))].sort().reverse();
  const monthTabs = ["ALL", ...tripMonths];
  if (!state.filter_trip_month || !monthTabs.includes(state.filter_trip_month)) {
    state.filter_trip_month = "ALL";
  }

  // 담당자 목록 (트립에 있는 담당자만)
  const tripEmployees = ["ALL", ...new Set(state.trips.map(t => t.employee).filter(Boolean))];
  if (!state.filter_trip_employee || !tripEmployees.includes(state.filter_trip_employee)) {
    state.filter_trip_employee = "ALL";
  }

  // 필터 적용 (월 + 담당자)
  const filtered = state.trips.filter(t => {
    if (state.filter_trip_month !== "ALL" && !(t.start_date || "").startsWith(state.filter_trip_month)) return false;
    if (state.filter_trip_employee !== "ALL" && t.employee !== state.filter_trip_employee) return false;
    return true;
  });

  return `
    <div class="flex center gap-3 wrap">
      <h2 style="margin:0; font-size:16px;">SCM 출장 <span style="color:#94a3b8; font-size:13px;">(${filtered.length} / ${state.trips.length}건 · 대상 ${scmCount}명)</span></h2>
      <div class="ml-auto flex gap-2">
        <button class="btn btn-outline" onclick="exportSheet('trips')">📤 엑셀 내보내기</button>
      </div>
    </div>

    <div class="card mt-3" style="background:#e0e7ff; border-color:#c7d2fe; padding:10px 14px; font-size:12px; color:#3730a3;">
      ℹ️ 출장·성과 tracking 은 <b>SCM 부서 인원(${scmCount}명)</b> 만 대상입니다.
    </div>

    <div class="mt-3">
      ${dropzone("tripplan", "출장 계획서 또는 결과보고 엑셀 드래그 (자동 인식)")}
    </div>

    <div class="card mt-3">
      <div style="display:flex; gap:12px; flex-wrap:wrap;">
        <div style="flex:1; min-width:200px;">
          <div class="chip-label">담당자 필터</div>
          <select class="select-filter" onchange="setTripEmpFilter(this.value)">
            ${tripEmployees.map(emp => {
              const label = emp === "ALL" ? "전체" : nickOnly(emp);
              return `<option value="${escHTML(emp)}" ${state.filter_trip_employee === emp ? "selected" : ""}>${escHTML(label)}</option>`;
            }).join("")}
          </select>
        </div>
        <div style="flex:1; min-width:200px;">
          <div class="chip-label">월 필터 (출장 시작월 기준)</div>
          <select class="select-filter" onchange="state.filter_trip_month=this.value; render();">
            ${monthTabs.map(m => `<option value="${m}" ${state.filter_trip_month === m ? "selected" : ""}>${m === "ALL" ? "전체" : m}</option>`).join("")}
          </select>
        </div>
      </div>
    </div>

    ${(state.filter_trip_month === "ALL" && state.filter_trip_employee === "ALL") ? renderTripAnalytics() : ""}

    <div class="kanban mt-4" style="${(state.filter_trip_month === "ALL" && state.filter_trip_employee === "ALL") ? "display:none;" : ""}">
      ${cols.map(col => {
        const trips = filtered.filter(t => t.status === col);
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
                      ${(t.hotels && t.hotels.length > 0) ? `<div class="trip-card-meta" style="color:#0ea5e9; margin-top:4px;">🏨 방문 ${t.hotels.length}곳${t.expense_vnd ? ` · ${(t.expense_vnd/1000).toFixed(0)}k VND` : ""}</div>` : ""}
                      ${pCount > 0 ? `<div class="trip-card-meta" style="color:#4f46e5; margin-top:4px;">🏨 파트너 ${pVisited}/${pCount}</div>` : ""}
                      <div class="trip-card-footer">
                        <span class="trip-card-owner truncate">${escHTML(nickOnly(t.employee))}</span>
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
    ${row("담당자", escHTML(nickOnly(trip.employee) || "—"))}
    ${row("목적지", escHTML(trip.destination || "—"))}
    ${row("기간", `${trip.start_date || "—"} ~ ${trip.end_date || "—"}`)}
    ${row("목적", escHTML(trip.purpose || "—"))}
    ${row("상태", `<span class="badge b-primary">${escHTML(trip.status || "—")}</span>`)}
    ${row("예산", `${curSym(cur)}${fmt(trip.cost_planned)} ${cur}`)}
    ${row("실지출", `${curSym(cur)}${fmt(trip.cost_actual)} ${cur}`)}
    ${trip.roi != null ? row("실제 ROI", `<b>${trip.roi.toFixed(1)}x</b>`) : ""}
    ${trip.outcome ? `
      <div style="border-top:1px solid #e2e8f0; padding-top:12px; margin-top:14px;">
        <div style="font-size:11px; font-weight:600; margin-bottom:6px;">📋 출장 결과 요약</div>
        <div style="font-size:13px; color:#475569; white-space:pre-wrap;">${escHTML(trip.outcome)}</div>
      </div>` : ""}
    ${(trip.hotels && trip.hotels.length > 0) ? `
      <div style="margin-top:14px;">
        <div style="font-size:11px; font-weight:600; margin-bottom:6px; display:flex; align-items:center; gap:8px;">
          <span>🏨 방문 호텔 Recap (${trip.hotels.length}곳)</span>
          ${trip.expense_vnd ? `<span class="badge b-muted">💰 ${trip.expense_vnd.toLocaleString()} VND</span>` : ""}
        </div>
        <div style="max-height:420px; overflow:auto; border:1px solid #e2e8f0; border-radius:8px;">
          <table style="width:100%; font-size:11px; border-collapse:collapse;">
            <thead style="position:sticky; top:0; background:#f8fafc;">
              <tr style="border-bottom:2px solid #e2e8f0;">
                <th style="text-align:left; padding:6px 8px; width:26px;">#</th>
                <th style="text-align:left; padding:6px 8px;">Hotel</th>
                <th style="text-align:left; padding:6px 8px; width:110px;">Contract</th>
                <th style="text-align:left; padding:6px 8px; width:140px;">Purpose</th>
                <th style="text-align:left; padding:6px 8px; width:140px;">Follow-Up</th>
                <th style="text-align:left; padding:6px 8px; width:110px;">1주 후 결과</th>
                <th style="text-align:left; padding:6px 8px; width:90px;">상태</th>
              </tr>
            </thead>
            <tbody>
              ${trip.hotels.map((h, i) => {
                const st = h.status || "PENDING";
                const stColor = {DONE:"#16a34a", REQUESTED:"#f59e0b", IN_PROGRESS:"#3b82f6", REJECTED:"#dc2626", PENDING:"#94a3b8"}[st] || "#94a3b8";
                return `<tr style="border-bottom:1px solid #f1f5f9; vertical-align:top;">
                    <td style="padding:6px 8px; color:#94a3b8;">${i+1}</td>
                    <td style="padding:6px 8px;"><b>${escHTML(h.hotel || "")}</b>${h.contact ? `<div style="color:#94a3b8; font-size:10px; margin-top:2px;">${escHTML(h.contact)}</div>` : ""}</td>
                    <td style="padding:6px 8px; color:#475569;">${escHTML(h.contract || "—")}</td>
                    <td style="padding:6px 8px; color:#475569;">${escHTML((h.purpose || "").slice(0, 80))}</td>
                    <td style="padding:6px 8px; color:#475569;">${escHTML((h.followup || "").slice(0, 80))}</td>
                    <td style="padding:6px 8px; color:#475569;">${escHTML((h.result_1w || "").slice(0, 60))}</td>
                    <td style="padding:6px 8px;"><span style="display:inline-block; padding:2px 6px; border-radius:4px; background:${stColor}20; color:${stColor}; font-size:10px; font-weight:600;">${st}</span></td>
                  </tr>`;
              }).join("")}
            </tbody>
          </table>
        </div>
        ${trip.source_file ? `<div style="font-size:10px; color:#94a3b8; margin-top:6px;">📎 원본 파일: ${escHTML(trip.source_file)}</div>` : ""}
      </div>` : ""}
    ${trip.notes ? `
      <div style="margin-top:12px;">
        <div style="font-size:11px; font-weight:600; margin-bottom:6px;">📝 노트</div>
        <div style="font-size:13px; color:#475569; white-space:pre-wrap;">${escHTML(trip.notes)}</div>
      </div>` : ""}
  `);
}

function showModalReadOnly(title, body) {
  const html = `
    <div class="modal-backdrop" id="modal">
      <div class="modal">
        <div class="modal-head">
          <h3 style="margin:0; font-size:15px;">${title}</h3>
          <button class="btn btn-outline" onclick="closeModal()">✕</button>
        </div>
        <div class="modal-body">${body}</div>
        <div class="modal-foot">
          <button class="btn btn-primary" onclick="closeModal()">닫기</button>
        </div>
      </div>
    </div>`;
  const div = document.createElement("div");
  div.innerHTML = html;
  document.body.appendChild(div.firstElementChild);
}

// ==========================================================================
// View: 통합 캘린더 (SCM 담당자별 월별 근무/출장/연차 통합 뷰)
// ==========================================================================
function viewCalendar() {
  // 월 선택 상태
  if (!state.cal_month) {
    const months = [...new Set(state.attendance.map(a => (a.date || "").slice(0,7)).filter(Boolean))].sort();
    state.cal_month = months[months.length - 1] || "2026-06";
  }
  if (!state.cal_scope || state.cal_scope === "SCM") state.cal_scope = "ALL"; // ALL | 부서명 (예: Office/SCM)

  const [yStr, mStr] = state.cal_month.split("-");
  const year = parseInt(yStr), month = parseInt(mStr);
  const daysInMonth = new Date(year, month, 0).getDate();
  const days = Array.from({length: daysInMonth}, (_, i) => `${state.cal_month}-${String(i+1).padStart(2,"0")}`);

  // 월 목록
  const allMonths = [...new Set([...state.attendance.map(a => (a.date || "").slice(0,7)), ...state.trips.map(t => (t.start_date || "").slice(0,7))].filter(Boolean))].sort();

  // 부서 목록 (자동 추출)
  const allDepts = [...new Set(state.employees.map(e => e.department).filter(Boolean))].sort();

  // 담당자 리스트
  let emps;
  if (state.cal_scope === "ALL") {
    emps = state.employees.slice();
  } else {
    // 특정 부서
    emps = state.employees.filter(e => e.department === state.cal_scope);
  }
  emps.sort((a,b) => {
    const aHead = a.position === "SCM Head" ? 1 : 0;
    const bHead = b.position === "SCM Head" ? 1 : 0;
    if (aHead !== bHead) return bHead - aHead;
    return (a.name || "").localeCompare(b.name || "");
  });

  // Leave 데이터 (auto BT 포함)
  const allLeaves = [...(state.leaves || []), ...autoGenerateBTLeaves()];
  const leaveMap = {};
  allLeaves.filter(l => l.status === "APPROVED").forEach(l => {
    const key = `${l.person_id}|${l.date}`;
    if (!leaveMap[key] || leaveInfo(l.type).deduct_annual) leaveMap[key] = l;
  });

  // 근태 lookup
  const attMap = {};
  state.attendance.forEach(a => {
    attMap[`${a.person_id}|${a.date}`] = a;
  });

  // 요일 배열 (0=일, 6=토)
  const dow = ["일","월","화","수","목","금","토"];
  const isWeekend = (dateStr) => {
    const d = new Date(dateStr);
    return d.getDay() === 0 || d.getDay() === 6;
  };

  // 셀 스타일 결정
  const cellStyle = (emp, date) => {
    const leave = leaveMap[`${emp.person_id}|${date}`];
    const att = attMap[`${emp.person_id}|${date}`];
    const weekend = isWeekend(date);

    if (leave) {
      const info = leaveInfo(leave.type);
      return { bg: info.color, fg: "#fff", label: leave.type.replace("/",""), title: `${info.label}${leave.note ? " · " + leave.note : ""}` };
    }
    if (weekend) return { bg: "#f1f5f9", fg: "#94a3b8", label: "·", title: "주말" };
    if (!att) return { bg: "#f8fafc", fg: "#cbd5e1", label: "", title: "데이터 없음" };
    if (att.status === "LATE") return { bg: "#f59e0b", fg: "#fff", label: att.late_minutes ? `+${att.late_minutes}` : "L", title: `지각 ${att.late_minutes}분` };
    if (att.status === "ABSENT") return { bg: "#dc2626", fg: "#fff", label: "A", title: "결근" };
    return { bg: "#dcfce7", fg: "#16a34a", label: "○", title: "정상 출근" };
  };

  // 담당자별 월 요약
  const empSummary = (emp) => {
    let normalD = 0, lateD = 0, absentD = 0;
    const byType = {};
    Object.keys(LEAVE_TYPES).forEach(t => byType[t] = 0);
    days.forEach(date => {
      const leave = leaveMap[`${emp.person_id}|${date}`];
      const att = attMap[`${emp.person_id}|${date}`];
      if (leave) { byType[leave.type] = (byType[leave.type] || 0) + (leave.days || leaveInfo(leave.type).days); return; }
      if (!att) return;
      if (att.status === "LATE") lateD++;
      else if (att.status === "ABSENT") absentD++;
      else normalD++;
    });
    return { normalD, lateD, absentD, byType };
  };

  return `
    <div class="flex center gap-3 wrap">
      <h2 style="margin:0; font-size:16px;">📅 통합 캘린더 <span style="color:#94a3b8; font-size:13px;">담당자별 근무 · 출장 · 연차 통합 뷰</span></h2>
    </div>

    <div class="card mt-3">
      <div style="display:flex; gap:12px; flex-wrap:wrap;">
        <div style="flex:1; min-width:180px;">
          <div class="chip-label">월</div>
          <select class="select-filter" onchange="state.cal_month=this.value; render();">
            ${allMonths.map(m => `<option value="${m}" ${state.cal_month === m ? "selected" : ""}>${m}</option>`).join("")}
          </select>
        </div>
        <div style="flex:1; min-width:220px;">
          <div class="chip-label">범위 · 부서</div>
          <select class="select-filter" onchange="state.cal_scope=this.value; render();">
            <option value="ALL" ${state.cal_scope === "ALL" ? "selected" : ""}>전체 인원</option>
            ${allDepts.map(d => `<option value="${escHTML(d)}" ${state.cal_scope === d ? "selected" : ""}>${escHTML(d)}</option>`).join("")}
          </select>
        </div>
      </div>
      <div style="margin-top:10px; display:flex; gap:8px; flex-wrap:wrap; font-size:11px;">
        <span style="padding:2px 8px; border-radius:4px; background:#dcfce7; color:#16a34a;">○ 정상</span>
        <span style="padding:2px 8px; border-radius:4px; background:#f59e0b; color:#fff;">지각 (분)</span>
        <span style="padding:2px 8px; border-radius:4px; background:#dc2626; color:#fff;">A 결근</span>
        <span style="padding:2px 8px; border-radius:4px; background:#3b82f6; color:#fff;">AL 연차</span>
        <span style="padding:2px 8px; border-radius:4px; background:#f59e0b; color:#fff;">BT 출장</span>
        <span style="padding:2px 8px; border-radius:4px; background:#dc2626; color:#fff;">SL 병가</span>
        <span style="padding:2px 8px; border-radius:4px; background:#94a3b8; color:#fff;">UP 무급</span>
        <span style="padding:2px 8px; border-radius:4px; background:#f1f5f9; color:#94a3b8;">· 주말</span>
      </div>
    </div>

    <div class="card mt-3" style="padding:0; overflow:auto;">
      <table style="border-collapse:collapse; font-size:10px; width:100%; min-width:900px;">
        <thead>
          <tr style="background:#f8fafc; border-bottom:2px solid #e2e8f0;">
            <th style="text-align:left; padding:8px 10px; position:sticky; left:0; background:#f8fafc; z-index:2; min-width:120px;">담당자</th>
            ${days.map(d => {
              const dObj = new Date(d);
              const day = dObj.getDate();
              const w = dow[dObj.getDay()];
              const weekend = dObj.getDay() === 0 || dObj.getDay() === 6;
              return `<th style="padding:4px 2px; min-width:26px; font-weight:600; color:${weekend ? '#94a3b8' : '#0f172a'};">
                <div style="font-size:10px;">${day}</div>
                <div style="font-size:9px; opacity:0.6;">${w}</div>
              </th>`;
            }).join("")}
            <th style="padding:8px; background:#eff6ff; min-width:80px;">요약</th>
          </tr>
        </thead>
        <tbody>
          ${emps.map(emp => {
            const summ = empSummary(emp);
            const isHead = emp.position === "SCM Head";
            const scmTraveler = emp.is_scm && emp.is_scm_traveler !== false;
            return `
              <tr style="border-bottom:1px solid #f1f5f9;">
                <td style="padding:6px 10px; position:sticky; left:0; background:#fff; z-index:1; border-right:1px solid #e2e8f0;">
                  <div style="font-weight:600; color:${isHead ? '#4f46e5' : '#0f172a'}; font-size:12px;">${isHead ? '👑 ' : ''}${escHTML(nickOnly(emp.name))}</div>
                  <div style="font-size:9px; color:#94a3b8;">${escHTML(emp.department || "")}${scmTraveler ? " · 출장 대상" : ""}</div>
                </td>
                ${days.map(date => {
                  const s = cellStyle(emp, date);
                  return `<td style="padding:0; text-align:center; border:1px solid #f1f5f9;">
                    <div style="background:${s.bg}; color:${s.fg}; padding:4px 2px; font-size:9px; font-weight:600; height:24px; display:flex; align-items:center; justify-content:center;" title="${escHTML(date + ' · ' + s.title)}">${escHTML(s.label)}</div>
                  </td>`;
                }).join("")}
                <td style="padding:6px; background:#eff6ff; font-size:10px; text-align:center;">
                  <div style="color:#16a34a; font-weight:600;">${summ.normalD}일 정상</div>
                  ${summ.lateD > 0 ? `<div style="color:#f59e0b;">지각 ${summ.lateD}</div>` : ""}
                  ${summ.absentD > 0 ? `<div style="color:#dc2626;">결근 ${summ.absentD}</div>` : ""}
                  ${summ.byType["AL"] > 0 ? `<div style="color:#3b82f6;">AL ${summ.byType["AL"]}</div>` : ""}
                  ${summ.byType["AL/2"] > 0 ? `<div style="color:#60a5fa;">AL/2 ${summ.byType["AL/2"]}</div>` : ""}
                  ${summ.byType["BT"] > 0 ? `<div style="color:#f59e0b;">BT ${summ.byType["BT"]}</div>` : ""}
                </td>
              </tr>`;
          }).join("")}
        </tbody>
      </table>
    </div>

    <div style="font-size:11px; color:#64748b; margin-top:10px;">
      💡 셀에 마우스를 올리면 상세 정보 (날짜 · 상태 · 지각 분 · 트립 목적지 등) 확인 가능. BT(출장) 은 SCM 트립 데이터에서 자동 반영됩니다.
    </div>
  `;
}

function viewReports() {
  const monthly = ["2026-03","2026-04","2026-05","2026-06"].map(m => ({
    m,
    late: state.attendance.filter(a => a.date && a.date.startsWith(m) && a.status === "LATE").length,
    absent: state.attendance.filter(a => a.date && a.date.startsWith(m) && a.status === "ABSENT").length,
  }));
  const monthlyMax = Math.max(1, ...monthly.map(m => Math.max(m.late, m.absent)));
  const deptStats = {};
  state.attendance.forEach(a => {
    if (!a.department) return;
    if (!deptStats[a.department]) deptStats[a.department] = { total: 0, late: 0, absent: 0 };
    deptStats[a.department].total++;
    if (a.status === "LATE") deptStats[a.department].late++;
    if (a.status === "ABSENT") deptStats[a.department].absent++;
  });
  const attStatus = {};
  state.attendance.forEach(a => { attStatus[a.status] = (attStatus[a.status] || 0) + 1; });
  const attTotal = state.attendance.length;
  const tripStatus = {};
  state.trips.forEach(t => { tripStatus[t.status] = (tripStatus[t.status] || 0) + 1; });
  const tripTotal = state.trips.length;
  const deptEmp = {};
  state.employees.forEach(e => { deptEmp[e.department] = (deptEmp[e.department] || 0) + 1; });
  return `
    <div class="flex center gap-3"><h2 style="margin:0; font-size:16px;">리포트</h2></div>
    <div class="grid-2 mt-4">
      <div class="card"><h3>월별 지각/결근</h3>
        <div style="padding:12px;">${monthly.map(m => `
          <div style="margin-bottom:10px;">
            <div style="display:flex; justify-content:space-between; font-size:12px; margin-bottom:2px;"><b>${m.m}</b><span style="color:#94a3b8;">지각 ${m.late} · 결근 ${m.absent}</span></div>
            <div style="display:flex; gap:2px; height:16px;">
              <div style="flex:${m.late}; background:#f59e0b; border-radius:2px;"></div>
              <div style="flex:${m.absent}; background:#ef4444; border-radius:2px;"></div>
              <div style="flex:${Math.max(0, monthlyMax - m.late - m.absent)}; background:#f1f5f9;"></div>
            </div>
          </div>`).join("")}
        </div>
      </div>
      <div class="card"><h3>근태 상태 분포</h3>${doughnut(attStatus, attTotal, { NORMAL: "#22c55e", PRESENT: "#22c55e", LATE: "#f59e0b", ABSENT: "#ef4444", LEAVE: "#3b82f6", HOLIDAY: "#94a3b8" })}</div>
      <div class="card"><h3>부서별 인원</h3>${doughnut(deptEmp, state.employees.length, { "Office/SCM": "#4f46e5", "Office/PD": "#7c3aed", "KR Manager": "#0ea5e9", "Office": "#94a3b8" })}</div>
      <div class="card"><h3>SCM 출장 상태</h3>${doughnut(tripStatus, tripTotal, { DRAFT: "#94a3b8", REQUESTED: "#f59e0b", APPROVED: "#3b82f6", IN_PROGRESS: "#22c55e", COMPLETED: "#0ea5e9", CANCELLED: "#ef4444" })}</div>
    </div>
    <div class="card mt-4"><h3>부서별 근태 요약</h3>
      <div class="table-wrap"><table>
        <thead><tr><th>부서</th><th class="right">총 근태</th><th class="right">지각</th><th class="right">결근</th><th class="right">지각률</th></tr></thead>
        <tbody>${Object.entries(deptStats).sort((a,b) => (b[1].late/b[1].total) - (a[1].late/a[1].total)).map(([d, s]) => {
          const rate = ((s.late / s.total) * 100).toFixed(1);
          const isSCM = d.toUpperCase().includes("SCM");
          return `<tr>
              <td><b>${escHTML(d)}</b>${isSCM ? `<span class="badge b-scm" style="margin-left:6px;">SCM</span>` : ""}</td>
              <td class="right">${s.total.toLocaleString()}</td>
              <td class="right ${s.late > 20 ? "text-late" : ""}">${s.late}</td>
              <td class="right ${s.absent > 5 ? "text-absent" : ""}">${s.absent}</td>
              <td class="right"><b>${rate}%</b></td>
            </tr>`;
        }).join("")}</tbody>
      </table></div>
    </div>`;
}

function doughnut(counts, total, colors) {
  const entries = Object.entries(counts).filter(([, v]) => v > 0);
  if (!entries.length || total === 0) return `<div style="text-align:center; padding:32px; color:#94a3b8;">데이터 없음</div>`;
  const C = 2 * Math.PI * 60;
  let acc = 0;
  const arcs = entries.map(([k, v]) => {
    const frac = v / total;
    const dash = frac * C;
    const off = -acc * C;
    acc += frac;
    return `<circle cx="80" cy="80" r="60" fill="none" stroke="${colors[k] || "#94a3b8"}" stroke-width="24" stroke-dasharray="${dash} ${C - dash}" stroke-dashoffset="${off}" transform="rotate(-90 80 80)"/>`;
  }).join("");
  const legend = entries.map(([k, v]) => `
    <div style="display:flex; align-items:center; gap:6px; font-size:12px; margin:2px 0;">
      <div style="width:12px; height:12px; background:${colors[k] || "#94a3b8"}; border-radius:3px;"></div>
      <span>${k}: <b>${v}</b> (${((v/total)*100).toFixed(0)}%)</span>
    </div>`).join("");
  return `
    <div style="display:flex; align-items:center; gap:16px; padding:12px;">
      <svg width="160" height="160" viewBox="0 0 160 160">${arcs}
        <text x="80" y="76" text-anchor="middle" font-size="12" fill="#64748b">총</text>
        <text x="80" y="94" text-anchor="middle" font-size="20" font-weight="bold" fill="#0f172a">${total}</text>
      </svg>
      <div style="flex:1;">${legend}</div>
    </div>`;
}

(async function() {
  try {
    await load();
    state.loaded = true;
    render();
  } catch (e) {
    console.error("Boot failed:", e);
  }
})();
