/**
 * 부산동성고등학교 성적 분석 대시보드 - 핵심 Logic (app.js)
 * 
 * 주요 기능:
 * 1. 엑셀 파일 (.xlsx, .csv) 드래그앤드롭 및 파일 로딩
 * 2. 복사/붙여넣기 텍스트 파싱 (CSV / TSV 자동 인식)
 * 3. 이름 기준 3개년 성적 데이터 병합 및 동명이인 예외 처리
 * 4. 단순 평균, 단위수 반영 가중 평균, 과목군별 평균 등급 계산
 * 5. Chart.js 연동 (역스케일 적용: 1등급이 가장 바깥쪽/길게 표시됨)
 * 6. 내장 샘플 데이터 제공
 */

// 애플리케이션 전역 상태 관리
const AppState = {
    rawGrades: [],     // 파싱된 원본 레코드 배열
    students: {},      // 학생별 가공 데이터 (Key: 학생식별자명)
    activeStudent: null, // 현재 선택된 학생 객체
    charts: {
        radar: null,
        bar: null,
        trend: null
    }
};

// 학생의 사용 가능한 가장 높은 학년 학번 추출 (정렬 기준 자동 선택용)
function getHighestStudentId(student) {
    if (!student || !student.studentIds) return "";
    if (student.studentIds[3]) return student.studentIds[3];
    if (student.studentIds[2]) return student.studentIds[2];
    if (student.studentIds[1]) return student.studentIds[1];
    return "";
}

// 과목군 정의 및 스타일 색상 매핑
const SUBJECT_GROUPS = {
    "국어": { class: "kor", color: "rgba(239, 68, 68, 0.8)", border: "rgba(239, 68, 68, 1)" },
    "수학": { class: "mat", color: "rgba(59, 130, 246, 0.8)", border: "rgba(59, 130, 246, 1)" },
    "영어": { class: "eng", color: "rgba(16, 185, 129, 0.8)", border: "rgba(16, 185, 129, 1)" },
    "사회": { class: "soc", color: "rgba(245, 158, 11, 0.8)", border: "rgba(245, 158, 11, 1)" },
    "과학": { class: "sci", color: "rgba(139, 92, 246, 0.8)", border: "rgba(139, 92, 246, 1)" },
    "기타": { class: "etc", color: "rgba(100, 116, 139, 0.8)", border: "rgba(100, 116, 139, 1)" }
};

// 웹 페이지 로드 시 이벤트 바인딩
document.addEventListener("DOMContentLoaded", () => {
    // Lucide 아이콘 초기화
    lucide.createIcons();
    
    // UI 요소 바인딩
    initEventBindings();
    
    // 아코디언 설정
    initAccordion();
});

// 이벤트 리스너 설정
function initEventBindings() {
    // 탭 전환
    const tabs = document.querySelectorAll(".tab-btn");
    tabs.forEach(tab => {
        tab.addEventListener("click", () => {
            tabs.forEach(t => t.classList.remove("active"));
            document.querySelectorAll(".tab-content").forEach(c => c.classList.remove("active"));
            
            tab.classList.add("active");
            document.getElementById(tab.dataset.tab).classList.add("active");
        });
    });

    // 드래그 앤 드롭 파일 업로드
    const dropZone = document.getElementById("dropZone");
    const fileInput = document.getElementById("fileInput");

    dropZone.addEventListener("click", () => fileInput.click());
    dropZone.addEventListener("dragover", (e) => {
        e.preventDefault();
        dropZone.classList.add("dragover");
    });
    dropZone.addEventListener("dragleave", () => {
        dropZone.classList.remove("dragover");
    });
    dropZone.addEventListener("drop", (e) => {
        e.preventDefault();
        dropZone.classList.remove("dragover");
        if (e.dataTransfer.files.length > 0) {
            handleMultipleFiles(e.dataTransfer.files);
        }
    });
    fileInput.addEventListener("change", (e) => {
        if (e.target.files.length > 0) {
            handleMultipleFiles(e.target.files);
        }
    });

    // 텍스트 붙여넣기 파싱
    document.getElementById("parsePasteBtn").addEventListener("click", handlePasteInput);

    // '성적파일' 폴더 자동화 일괄 로딩 바인딩
    const folderInput = document.getElementById("folderInput");
    document.getElementById("loadSampleBtn").addEventListener("click", () => folderInput.click());
    document.getElementById("welcomeSampleBtn").addEventListener("click", () => folderInput.click());
    folderInput.addEventListener("change", (e) => {
        if (e.target.files.length > 0) {
            handleMultipleFiles(e.target.files);
        }
    });

    // 엑셀 내보내기 버튼
    document.getElementById("exportExcelBtn").addEventListener("click", exportToExcel);

    // 대시보드 초기화 버튼
    document.getElementById("resetBtn").addEventListener("click", resetDashboard);

    // 검색창 입력 이벤트
    document.getElementById("searchBar").addEventListener("input", filterStudentList);

    // 필터 조건 변경
    document.getElementById("gradeFilterSelect").addEventListener("change", renderDetailedGrades);
    document.getElementById("groupFilterSelect").addEventListener("change", renderDetailedGrades);

    // 정렬 조건 변경
    document.getElementById("studentSortSelect").addEventListener("change", renderStudentList);
}

// 아코디언 기능 구현
function initAccordion() {
    const toggle = document.getElementById("mappingInfoToggle");
    if (!toggle) return;
    const accordion = toggle.parentElement;
    toggle.addEventListener("click", () => {
        accordion.classList.toggle("open");
    });
}

// 다중 엑셀 파일 처리 및 순차적 파싱
async function handleMultipleFiles(files) {
    let allParsedGrades = [];
    const fileNames = [];
    
    for (let i = 0; i < files.length; i++) {
        const file = files[i];
        
        // 시스템 임시 락 파일(~$로 시작) 및 비엑셀/비CSV 파일 건너뛰기
        const name = file.name;
        if (name.startsWith("~$") || (!name.toLowerCase().endsWith(".xlsx") && !name.toLowerCase().endsWith(".xls") && !name.toLowerCase().endsWith(".csv"))) {
            continue;
        }
        
        fileNames.push(name);
        try {
            const parsed = await parseSingleExcelFile(file);
            allParsedGrades = allParsedGrades.concat(parsed);
        } catch (err) {
            console.error(`Error parsing file ${file.name}:`, err);
            alert(`파일 [${file.name}] 파싱 중 오류가 발생했습니다: ${err.message}`);
        }
    }
    
    if (allParsedGrades.length === 0) {
        return;
    }
    
    // 신규 파싱된 데이터를 기존 데이터베이스에 결합 (중복 제거)
    addParsedGrades(allParsedGrades);
    
    // 파일 업로드 목록 UI 업데이트
    updateLoadedFilesUI(fileNames);
    
    // 3개년 병합 프로필 생성 및 대시보드 출력
    buildStudentProfiles();
}

// 개별 파일 읽기 (Promise 기반 비동기화)
function parseSingleExcelFile(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = function(e) {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const firstSheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[firstSheetName];
                const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
                
                const parsed = processRawRowsIntoArray(rows, file.name);
                resolve(parsed);
            } catch (err) {
                reject(err);
            }
        };
        reader.onerror = () => reject(reader.error);
        reader.readAsArrayBuffer(file);
    });
}

// 직접 붙여넣기 텍스트 파싱
function handlePasteInput() {
    const text = document.getElementById("pasteInput").value.trim();
    if (!text) {
        alert("성적 데이터를 붙여넣은 뒤 분석 버튼을 눌러주세요.");
        return;
    }

    const lines = text.split(/\r?\n/);
    const rows = [];
    for (const line of lines) {
        if (!line.trim()) continue;
        
        let delimiter = "\t";
        if (!line.includes("\t") && line.includes(",")) {
            delimiter = ",";
        }
        
        rows.push(line.split(delimiter).map(cell => cell.trim().replace(/^["']|["']$/g, '')));
    }

    try {
        const parsed = processRawRowsIntoArray(rows);
        if (parsed.length === 0) {
            alert("유효한 성적 레코드를 찾을 수 없습니다.");
            return;
        }
        addParsedGrades(parsed);
        updateLoadedFilesUI(["직접 붙여넣기 데이터"]);
        buildStudentProfiles();
        document.getElementById("pasteInput").value = ""; // 성공 시 텍스트 영역 클리어
    } catch (err) {
        console.error(err);
        alert("성적 데이터를 파싱하는 중 오류가 발생했습니다: " + err.message);
    }
}

// 업로드된 파일 정보 저장 및 UI 동기화
function updateLoadedFilesUI(newFileNames) {
    if (!AppState.uploadedFiles) {
        AppState.uploadedFiles = [];
    }
    
    newFileNames.forEach(name => {
        if (!AppState.uploadedFiles.includes(name)) {
            AppState.uploadedFiles.push(name);
        }
    });
    
    const container = document.getElementById("loadedFilesContainer");
    const countSpan = document.getElementById("fileCount");
    const listUl = document.getElementById("loadedFilesList");
    
    if (AppState.uploadedFiles.length > 0) {
        container.style.display = "block";
        countSpan.textContent = AppState.uploadedFiles.length + "개";
        listUl.innerHTML = "";
        
        AppState.uploadedFiles.forEach(name => {
            const li = document.createElement("li");
            li.innerHTML = `
                <span class="file-name" title="${name}">${name}</span>
                <span class="file-status">완료</span>
            `;
            listUl.appendChild(li);
        });
    } else {
        container.style.display = "none";
    }
}

// 중복 데이터 병합 결과 토스트 알림 표시
function showDuplicateNotification(count) {
    let container = document.getElementById("toastContainer");
    if (!container) {
        container = document.createElement("div");
        container.id = "toastContainer";
        container.style.position = "fixed";
        container.style.top = "20px";
        container.style.right = "20px";
        container.style.zIndex = "9999";
        container.style.display = "flex";
        container.style.flexDirection = "column";
        container.style.gap = "10px";
        document.body.appendChild(container);
    }
    
    // 토스트 애니메이션 스타일 주입
    if (!document.getElementById("toastStyles")) {
        const style = document.createElement("style");
        style.id = "toastStyles";
        style.textContent = `
            @keyframes slideIn {
                from { opacity: 0; transform: translateY(-10px); }
                to { opacity: 1; transform: translateY(0); }
            }
            .toast-alert {
                animation: slideIn 0.3s ease-out;
                border: 1px solid var(--border);
            }
        `;
        document.head.appendChild(style);
    }
    
    const toast = document.createElement("div");
    toast.className = "toast-alert";
    toast.style.backgroundColor = "var(--bg-card)";
    toast.style.borderLeft = "4px solid var(--primary-light)";
    toast.style.color = "var(--text-primary)";
    toast.style.padding = "12px 20px";
    toast.style.borderRadius = "var(--radius-sm)";
    toast.style.boxShadow = "var(--shadow-lg)";
    toast.style.display = "flex";
    toast.style.alignItems = "center";
    toast.style.gap = "10px";
    toast.style.fontFamily = "system-ui, sans-serif";
    toast.style.fontSize = "0.85rem";
    toast.style.fontWeight = "600";
    
    toast.innerHTML = `
        <span style="color: var(--primary-light); font-size: 1.1rem;">ℹ️</span>
        <span>중복 성적 데이터 ${count}건이 감지되어 최신 자료로 합산 및 업데이트되었습니다.</span>
    `;
    
    container.appendChild(toast);
    
    setTimeout(() => {
        toast.style.transition = "opacity 0.5s ease, transform 0.5s ease";
        toast.style.opacity = "0";
        toast.style.transform = "translateY(-10px)";
        setTimeout(() => {
            toast.remove();
        }, 500);
    }, 4500);
}

// 기존 데이터와 병합 시 중복 데이터(동일 학번 + 학년 + 과목) 덮어쓰기 로직
function addParsedGrades(newGrades) {
    let duplicateCount = 0;
    newGrades.forEach(newG => {
        const dupIdx = AppState.rawGrades.findIndex(g => {
            // 1. 학번이 완전히 일치하면 과목 정보 비교
            if (g.학번 === newG.학번) {
                return g.학년 === newG.학년 && g.학기 === newG.학기 && g.과목명 === newG.과목명;
            }
            
            // 2. 학번이 다르더라도 이름과 성적 정보가 일치하고, 둘 중 하나가 전입생 임시 반(99)인 경우 동일 학생으로 중복 교정 처리
            if (g.이름 === newG.이름 && g.학년 === newG.학년 && g.학기 === newG.학기 && g.과목명 === newG.과목명) {
                const isTempG = g.학번 && g.학번.substring(1, 3) === "99";
                const isTempNew = newG.학번 && newG.학번.substring(1, 3) === "99";
                if (isTempG || isTempNew) {
                    return true;
                }
            }
            return false;
        });
        
        if (dupIdx !== -1) {
            // 만약 기존 데이터의 학번이 이미 실제 학급으로 보정되어 있고(99반이 아님),
            // 새 데이터는 임시 학번(99반)을 갖고 있다면, 보정된 기존 학번을 새 데이터에 이식하여 병합!
            const oldId = AppState.rawGrades[dupIdx].학번;
            const newId = newG.학번;
            if (oldId && newId && oldId.substring(1, 3) !== "99" && newId.substring(1, 3) === "99") {
                newG.학번 = oldId;
            }
            AppState.rawGrades[dupIdx] = newG; // 중복 레코드는 최신 파일 내용으로 덮어씀
            duplicateCount++;
        } else {
            AppState.rawGrades.push(newG);
        }
    });
    
    if (duplicateCount > 0) {
        showDuplicateNotification(duplicateCount);
    }
}

// 2차원 배열 데이터의 레이아웃을 감지하여 알맞은 파서로 연동하는 라우터
function processRawRowsIntoArray(rows, fileName) {
    if (rows.length < 2) return [];

    let isNeisLedger = false;
    let hasNeisHeader = false;
    let isGridPivot = false;
    let isRecordFormat = false;

    // 생활기록부 형식 여부 감지 (첫 5개 행 중에서 '성명', '과목', '석차등급' / '성취도' / '학점' 동시 감지)
    for (let r = 0; r < Math.min(rows.length, 5); r++) {
        const row = rows[r];
        if (!row) continue;
        const joined = row.map(v => String(v || "").trim()).join("|").replace(/\s+/g, "");
        // 생활기록부 전입생 양식은 '과목' 컬럼명과 '성취도' / '학점'이 포함됨 (종합일람표의 '교과목'과는 구별)
        if (joined.includes("성명") && (joined.includes("과목") || joined.includes("과목명")) && (joined.includes("석차등급") || joined.includes("성취도") || joined.includes("학점")) && !joined.includes("교과목")) {
            isRecordFormat = true;
            break;
        }
    }

    if (isRecordFormat) {
        return parseRecordFormat(rows, fileName);
    }

    // 기존 나이스 종합일람표 및 그리드 피벗 양식 감지
    for (let r = 0; r < Math.min(rows.length, 15); r++) {
        const row = rows[r];
        if (!row) continue;
        for (let c = 0; c < row.length; c++) {
            const val = String(row[c] || "").trim();
            if (val.includes("학기말 성적 종합일람표")) {
                isNeisLedger = true;
                break;
            }
            if (val === "교과목") {
                hasNeisHeader = true;
            }
            if (val === "단위수") {
                isGridPivot = true;
            }
        }
        if (isNeisLedger || isGridPivot) break;
    }

    if (isGridPivot) {
        return parseGridFormat(rows, fileName);
    } else if (isNeisLedger || hasNeisHeader) {
        return parseNeisLedger(rows, fileName);
    } else {
        return parseStandardFormat(rows, fileName);
    }
}

// 퍼지 컬럼 매칭 함수 (한글 조어 특성 반영)
function findColumnIndex(headers, keys) {
    for (let i = 0; i < headers.length; i++) {
        if (!headers[i]) continue;
        const h = String(headers[i]).trim().replace(/\s+/g, '');
        for (const k of keys) {
            if (h.includes(k) || k.includes(h)) return i;
        }
    }
    return -1;
}

// 과목군 텍스트 정규화
function normalizeGroup(group) {
    if (!group) return "기타";
    const g = String(group).trim().replace(/\s+/g, '');
    
    // 사용자 명시 기타 과목군 매핑
    if (g.includes("글로벌이슈") || g.includes("미술감상") || g.includes("인간과경제") || g.includes("인공지능") || g.includes("음악감상") || g.includes("중국어")) {
        return "기타";
    }
    
    // 사용자 요구 명시 매핑 (한국사 -> 사회, 문학 -> 국어, 기하/대수 -> 수학)
    if (g.includes("한국사")) return "사회";
    if (g.includes("문학")) return "국어";
    if (g.includes("기하")) return "수학";
    if (g.includes("대수")) return "수학";
    
    if (g.includes("국어") || g.includes("국")) return "국어";
    if (g.includes("수학") || g.includes("수")) return "수학";
    if (g.includes("영어") || g.includes("영")) return "영어";
    if (g.includes("사회") || g.includes("사") || g.includes("역사") || g.includes("지리") || g.includes("도덕") || g.includes("윤리")) return "사회";
    if (g.includes("과학") || g.includes("과") || g.includes("물리") || g.includes("화학") || g.includes("생명") || g.includes("지학") || g.includes("지구")) return "과학";
    return "기타";
}

// 포맷 A: 단순 열 매핑형 (기본 템플릿용)
function parseStandardFormat(rows, fileName) {
    const headers = rows[0].map(h => String(h || "").trim());

    const idxGrade = findColumnIndex(headers, ["학년", "grade"]);
    const idxId = findColumnIndex(headers, ["학번", "id", "number"]);
    const idxName = findColumnIndex(headers, ["이름", "성명", "name"]);
    const idxSubject = findColumnIndex(headers, ["과목명", "과목", "subject"]);
    const idxUnits = findColumnIndex(headers, ["단위수", "단위", "시수", "unit"]);
    const idxRank = findColumnIndex(headers, ["등급", "석차등급", "rank"]);
    const idxGroup = findColumnIndex(headers, ["과목군", "분류", "group"]);
    const idxSemester = findColumnIndex(headers, ["학기", "semester"]);

    if (idxGrade === -1 || idxName === -1 || idxSubject === -1 || idxUnits === -1 || idxRank === -1) {
        throw new Error("필수 열(학년, 이름, 과목명, 단위수, 등급)을 찾을 수 없습니다.");
    }

    let defaultSemester = 1;
    if (fileName) {
        const semMatch = fileName.match(/([1-2])학기/);
        if (semMatch) {
            defaultSemester = parseInt(semMatch[1]);
        }
    }

    const parsedGrades = [];
    for (let i = 1; i < rows.length; i++) {
        const row = rows[i];
        if (!row || row.length === 0 || !row[idxName]) continue;

        const rawGrade = row[idxGrade];
        const rawId = row[idxId] ? String(row[idxId]).trim() : "";
        const rawName = String(row[idxName]).trim();
        const rawSubject = String(row[idxSubject]).trim();
        const rawUnits = row[idxUnits];
        const rawRank = row[idxRank];
        const rawGroup = idxGroup !== -1 && row[idxGroup] ? String(row[idxGroup]).trim() : "기타";
        
        let semester = defaultSemester;
        if (idxSemester !== -1 && row[idxSemester]) {
            const valSem = parseInt(String(row[idxSemester]).trim());
            if (!isNaN(valSem) && (valSem === 1 || valSem === 2)) {
                semester = valSem;
            }
        }

        const grade = parseInt(rawGrade);
        const units = parseFloat(rawUnits);
        
        if (isNaN(grade) || isNaN(units)) {
            continue;
        }

        const rawRankStr = String(rawRank).trim().toUpperCase();
        let rank;
        if (rawRankStr === "P") {
            rank = "P";
        } else {
            rank = parseInt(rawRank);
            if (isNaN(rank) || rank < 1 || rank > 9) {
                continue;
            }
        }

        parsedGrades.push({
            학년: grade,
            학기: semester,
            학번: rawId,
            이름: rawName,
            과목명: rawSubject,
            단위수: units,
            등급: rank,
            과목군: normalizeGroup(rawGroup)
        });
    }
    return parsedGrades;
}

// 포맷 B: 나이스 학기말 성적 종합일람표 형식 파서 (멀티 테이블 및 성취도 추출 지원)
function parseNeisLedger(rows, fileName) {
    const parsedData = [];
    
    // 1단계: 상단 메타데이터 파싱 (학년, 반 도출)
    let grade = 1;
    let classNum = 1;
    let semester = 1;
    
    if (fileName) {
        const semMatch = fileName.match(/([1-2])학기/);
        if (semMatch) semester = parseInt(semMatch[1]);
    }
    
    for (let r = 0; r < Math.min(rows.length, 10); r++) {
        const row = rows[r];
        if (!row) continue;
        for (let c = 0; c < row.length; c++) {
            const val = String(row[c] || "");
            
            const semMatch = val.match(/([1-2])학기/);
            if (semMatch) {
                semester = parseInt(semMatch[1]);
            }
            
            if (val.includes("학년") && val.includes("반")) {
                const tokens = val.split(/\s+/);
                tokens.forEach(token => {
                    if (token.includes("학년도")) return; // 연도 데이터 스킵
                    const gMatch = token.match(/([1-3])학년/);
                    if (gMatch) grade = parseInt(gMatch[1]);
                    const cMatch = token.match(/(\d+)반/);
                    if (cMatch) classNum = parseInt(cMatch[1]);
                });
            }
        }
    }
    
    // 2단계: '교과목' 헤더 행들 찾기 (하나의 시트에 세로로 여러 테이블이 있는 형태 지원)
    const headerRowIndices = [];
    for (let r = 0; r < rows.length; r++) {
        const row = rows[r];
        if (!row) continue;
        for (let c = 0; c < row.length; c++) {
            if (String(row[c] || "").trim() === "교과목") {
                headerRowIndices.push({ rowIdx: r, idxGubun: c });
                break; // 이 행에서는 더 이상 탐색하지 않음
            }
        }
    }
    
    if (headerRowIndices.length === 0) {
        throw new Error("성적 종합일람표 양식의 '교과목' 헤더 행을 찾을 수 없습니다.");
    }
    
    // 각 서브 테이블별로 순차적 파싱 진행
    for (let i = 0; i < headerRowIndices.length; i++) {
        const { rowIdx, idxGubun } = headerRowIndices[i];
        const nextHeaderIdx = (i + 1 < headerRowIndices.length) ? headerRowIndices[i + 1].rowIdx : rows.length;
        
        const headerRow = rows[rowIdx];
        const subjectCols = {}; // colIdx -> { name, units }
        const idxName = idxGubun - 1;
        const idxRoll = idxGubun - 2;
        const startSubjectCol = idxGubun + 1;
        
        // 과목 매핑
        for (let c = startSubjectCol; c < headerRow.length; c++) {
            const val = String(headerRow[c] || "").trim();
            if (val) {
                const match = val.match(/^([^\(]+)(?:\((\d+)\))?/);
                if (match) {
                    const subName = match[1].trim();
                    const units = match[2] ? parseInt(match[2]) : 1;
                    subjectCols[c] = { name: subName, units: units };
                }
            }
        }
        
        // 데이터 행 파싱
        for (let r = rowIdx + 1; r < nextHeaderIdx; r++) {
            const row = rows[r];
            if (!row) continue;
            
            // '석차등급' 줄 발견 시 파싱
            if (String(row[idxGubun] || "").trim() === "석차등급") {
                const achievementRowIdx = r - 1;
                const nameRowIdx = r - 2;
                
                const achievementRow = rows[achievementRowIdx];
                const nameRow = rows[nameRowIdx];
                
                if (nameRow && nameRow[idxName]) {
                    const studentName = String(nameRow[idxName]).trim();
                    const rollNumStr = String(nameRow[idxRoll]).trim();
                    let rollNum = parseInt(parseFloat(rollNumStr));
                    if (isNaN(rollNum)) rollNum = 1;
                    
                    const classStr = String(classNum).padStart(2, '0');
                    const rollStr = String(rollNum).padStart(2, '0');
                    const studentId = `${grade}${classStr}${rollStr}`;
                    
                    Object.keys(subjectCols).forEach(colIdx => {
                        const c = parseInt(colIdx);
                        const cellVal = String(row[c] || "").trim();
                        const achVal = achievementRow ? String(achievementRow[c] || "").trim() : "";
                        
                        // 석차등급 값이 있거나 성취도 값이 있는 경우 이수 과목으로 수집
                        if (cellVal || achVal) {
                            const valUpper = cellVal.toUpperCase();
                            let rank;
                            let achievement = achVal;
                            
                            if (valUpper === "P") {
                                rank = "P";
                            } else {
                                const parsedRank = parseInt(parseFloat(cellVal));
                                if (!isNaN(parsedRank) && parsedRank >= 1 && parsedRank <= 9) {
                                    rank = parsedRank;
                                } else {
                                    // 숫자가 아니거나 공백인 경우, 성취도(A, B, C 등)가 있다면 P(이수)로 대체 수집
                                    if (achVal && achVal !== "/") {
                                        rank = "P";
                                    } else {
                                        return; // 수강 데이터 없음
                                    }
                                }
                            }
                            
                            const subInfo = subjectCols[c];
                            parsedData.push({
                                학년: grade,
                                학기: semester,
                                학번: studentId,
                                이름: studentName,
                                과목명: subInfo.name,
                                단위수: subInfo.units,
                                등급: rank,
                                성취도: achievement,
                                과목군: normalizeGroup(subInfo.name)
                            });
                        }
                    });
                }
            }
        }
    }
    
    return parsedData;
}

// 포맷 D: 생활기록부 교과학습발달사항 형식 파서 (전입생 등 지원)
function parseRecordFormat(rows, fileName) {
    const parsedData = [];
    
    // 헤더 위치 분석 (첫 5행 중 탐색, 공백 및 공백문자 제거 후 검사)
    let headerRowIdx = -1;
    for (let r = 0; r < Math.min(rows.length, 5); r++) {
        const row = rows[r];
        if (!row) continue;
        const joined = row.map(v => String(v || "").trim()).join("|").replace(/\s+/g, "");
        if (joined.includes("성명") && (joined.includes("과목") || joined.includes("과목명")) && (joined.includes("석차등급") || joined.includes("성취도") || joined.includes("학점"))) {
            headerRowIdx = r;
            break;
        }
    }
    
    if (headerRowIdx === -1) {
        headerRowIdx = 0;
    }
    
    let headers = rows[headerRowIdx].map(v => String(v || "").trim().replace(/\s+/g, ""));
    let idxRoll = findColumnIndex(headers, ["번호", "번 호"]);
    let idxName = findColumnIndex(headers, ["성명", "이름"]);
    let idxGrade = findColumnIndex(headers, ["학년"]);
    let idxSemester = findColumnIndex(headers, ["학기"]);
    let idxSubject = findColumnIndex(headers, ["과목"]);
    let idxUnits = findColumnIndex(headers, ["학점", "단위수", "단위"]);
    let idxAchievement = findColumnIndex(headers, ["성취도"]);
    let idxRank = findColumnIndex(headers, ["석차등급", "등급"]);
    
    // 상태 상속 변수
    let activeRoll = "";
    let activeName = "";
    let activeGrade = 1;
    let activeSemester = 1;
    
    for (let r = headerRowIdx + 1; r < rows.length; r++) {
        const row = rows[r];
        if (!row || row.length === 0) continue;
        
        // 이수학점 합계나 빈 줄 등의 무시행 체크
        const firstCell = String(row[0] || "").trim();
        if (firstCell.includes("이수학점") || firstCell.includes("합계") || firstCell.includes("소계")) {
            continue;
        }
        
        // 중간 헤더 행 감지 시 동적 인덱스 맵 재계산
        const joined = row.map(v => String(v || "").trim()).join("|").replace(/\s+/g, "");
        if (joined.includes("성명") && (joined.includes("과목") || joined.includes("과목명")) && (joined.includes("석차등급") || joined.includes("성취도") || joined.includes("학점"))) {
            headers = row.map(v => String(v || "").trim().replace(/\s+/g, ""));
            idxRoll = findColumnIndex(headers, ["번호", "번 호"]);
            idxName = findColumnIndex(headers, ["성명", "이름"]);
            idxGrade = findColumnIndex(headers, ["학년"]);
            idxSemester = findColumnIndex(headers, ["학기"]);
            idxSubject = findColumnIndex(headers, ["과목"]);
            idxUnits = findColumnIndex(headers, ["학점", "단위수", "단위"]);
            idxAchievement = findColumnIndex(headers, ["성취도"]);
            idxRank = findColumnIndex(headers, ["석차등급", "등급"]);
            continue; // 헤더 행 자체는 데이터 파싱을 하지 않고 건너뜀
        }
        
        // 헤더 행 복사본 건너뛰기 필터링 (성명 셀이 "성명"이거나 번호 셀이 "번호"인 경우)
        if (idxName !== -1 && row[idxName] !== undefined) {
            const nameVal = String(row[idxName]).trim().replace(/\s+/g, "");
            if (nameVal === "성명" || nameVal === "이름") {
                continue;
            }
        }
        if (idxRoll !== -1 && row[idxRoll] !== undefined) {
            const rollVal = String(row[idxRoll]).trim().replace(/\s+/g, "");
            if (rollVal === "번호") {
                continue;
            }
        }
        
        // 값이 있을 때 상속 상태 갱신
        if (idxRoll !== -1 && row[idxRoll] !== undefined && String(row[idxRoll]).trim() !== "") {
            activeRoll = String(row[idxRoll]).trim();
        }
        if (idxName !== -1 && row[idxName] !== undefined && String(row[idxName]).trim() !== "") {
            activeName = String(row[idxName]).trim();
        }
        if (idxGrade !== -1 && row[idxGrade] !== undefined && String(row[idxGrade]).trim() !== "") {
            const parsedGrade = parseInt(parseFloat(String(row[idxGrade]).trim()));
            if (!isNaN(parsedGrade)) activeGrade = parsedGrade;
        }
        if (idxSemester !== -1 && row[idxSemester] !== undefined && String(row[idxSemester]).trim() !== "") {
            const parsedSemester = parseInt(parseFloat(String(row[idxSemester]).trim()));
            if (!isNaN(parsedSemester)) activeSemester = parsedSemester;
        }
        
        // 과목명과 석차등급/성취도 데이터 처리
        if (idxSubject !== -1 && row[idxSubject] !== undefined) {
            const subjectName = String(row[idxSubject]).trim();
            if (!subjectName || subjectName.replace(/\s+/g, "") === "과목") continue;
            
            const rawUnits = idxUnits !== -1 ? row[idxUnits] : 1;
            const units = parseFloat(String(rawUnits || "").trim()) || 1;
            
            const rawRank = idxRank !== -1 ? String(row[idxRank] || "").trim() : "";
            const rawAch = idxAchievement !== -1 ? String(row[idxAchievement] || "").trim() : "";
            
            if (rawRank || rawAch) {
                const rankUpper = rawRank.toUpperCase();
                let rank;
                
                if (rankUpper === "P") {
                    rank = "P";
                } else {
                    const parsedRank = parseInt(parseFloat(rawRank));
                    if (!isNaN(parsedRank) && parsedRank >= 1 && parsedRank <= 9) {
                        rank = parsedRank;
                    } else {
                        if (rawAch && rawAch !== "/" && rawAch !== "P") {
                            rank = "P"; // 석차등급이 숫자가 아니고 성취도가 있으면 P로 대체 수집
                        } else {
                            continue;
                        }
                    }
                }
                
                const classStr = "99"; // 전입생 전용 가상 반 코드
                const rollStr = String(activeRoll || "1").padStart(2, '0');
                const studentId = `${activeGrade}${classStr}${rollStr}`;
                
                parsedData.push({
                    학년: activeGrade,
                    학기: activeSemester,
                    학번: studentId,
                    이름: activeName,
                    과목명: subjectName,
                    단위수: units,
                    등급: rank,
                    성취도: rawAch,
                    과목군: normalizeGroup(subjectName)
                });
            }
        }
    }
    return parsedData;
}

// 포맷 C: 2학년 그리드 피벗 형태 파서 (단위수 행 분리형)
function parseGridFormat(rows, fileName) {
    const parsedData = [];
    
    // 1. "단위수" 셀 위치 찾기
    let unitRowIdx = -1;
    let unitColIdx = -1;
    for (let r = 0; r < rows.length; r++) {
        const row = rows[r];
        if (!row) continue;
        for (let c = 0; c < row.length; c++) {
            if (String(row[c] || "").trim() === "단위수") {
                unitRowIdx = r;
                unitColIdx = c;
                break;
            }
        }
        if (unitRowIdx !== -1) break;
    }
    
    if (unitRowIdx === -1) {
        throw new Error("그리드 양식에서 '단위수' 지표 셀을 찾을 수 없습니다.");
    }
    
    const subjectRowIdx = unitRowIdx + 1;
    const unitRow = rows[unitRowIdx];
    const subjectRow = rows[subjectRowIdx];
    if (!subjectRow) {
        throw new Error("단위수 행 아래에 과목명 행이 누락되었습니다.");
    }
    
    // 학기 탐지
    let semester = 1;
    if (fileName) {
        const semMatch = fileName.match(/([1-2])학기/);
        if (semMatch) semester = parseInt(semMatch[1]);
    }
    for (let r = 0; r < Math.min(rows.length, 5); r++) {
        const row = rows[r];
        if (!row) continue;
        for (let c = 0; c < row.length; c++) {
            const val = String(row[c] || "");
            const semMatch = val.match(/([1-2])학기/);
            if (semMatch) {
                semester = parseInt(semMatch[1]);
                break;
            }
        }
    }

    // 과목 열 매핑
    const subjectCols = {}; // colIdx -> { name, units }
    const startColIdx = unitColIdx + 1;
    
    for (let c = 0; c < subjectRow.length; c++) {
        if (c >= startColIdx) {
            const subName = String(subjectRow[c] || "").trim();
            const rawUnit = unitRow[c];
            let units = 1;
            if (rawUnit) {
                const parsedUnit = parseInt(parseFloat(String(rawUnit).trim()));
                if (!isNaN(parsedUnit)) {
                    units = parsedUnit;
                }
            }
            if (subName) {
                subjectCols[c] = { name: subName, units: units };
            }
        }
    }
    
    // 2. 학생 데이터 수집
    for (let r = subjectRowIdx + 1; r < rows.length; r++) {
        const row = rows[r];
        if (!row || row.length === 0) continue;
        
        // 학년(B, index 1), 반(C, index 2), 번호(D, index 3), 성명(E, index 4)
        if (row[1] !== undefined && row[4] !== undefined) {
            const grade = parseInt(parseFloat(String(row[1]).trim()));
            const clsNum = parseInt(parseFloat(String(row[2]).trim()));
            const rollNum = parseInt(parseFloat(String(row[3]).trim()));
            const studentName = String(row[4]).trim();
            
            if (isNaN(grade) || isNaN(clsNum) || isNaN(rollNum) || !studentName) {
                continue; // 헤더나 통계 요약행 제외
            }
            
            const studentId = `${grade}${String(clsNum).padStart(2, '0')}${String(rollNum).padStart(2, '0')}`;
            
            Object.keys(subjectCols).forEach(colIdx => {
                const c = parseInt(colIdx);
                const cellVal = String(row[c] || "").trim();
                if (cellVal) {
                    const valUpper = cellVal.toUpperCase();
                    if (valUpper === "P") {
                        const subInfo = subjectCols[c];
                        parsedData.push({
                            학년: grade,
                            학기: semester,
                            학번: studentId,
                            이름: studentName,
                            과목명: subInfo.name,
                            단위수: subInfo.units,
                            등급: "P",
                            과목군: normalizeGroup(subInfo.name)
                        });
                    } else {
                        const rank = parseInt(parseFloat(cellVal));
                        if (!isNaN(rank) && rank >= 1 && rank <= 9) {
                            const subInfo = subjectCols[c];
                            parsedData.push({
                                학년: grade,
                                학기: semester,
                                학번: studentId,
                                이름: studentName,
                                과목명: subInfo.name,
                                단위수: subInfo.units,
                                등급: rank,
                                과목군: normalizeGroup(subInfo.name)
                            });
                        }
                    }
                }
            });
        }
    }
    return parsedData;
}

/**
 * 이름 기반 데이터 병합 엔진 (동명이인 탐지 및 예외 처리 로직 포함)
 */
function buildStudentProfiles() {
    const studentsMap = {};
    // 0.5단계: 전입학 임시 학번(99반) 사전 보정
    // 임시 학번(99)과 실제 학번을 사전에 스캔하여 동명이인 오판을 예방하고 병합 정합성을 높입니다.
    const nameIdsMap = {}; // 이름 -> { 학년: Set(학번) }
    AppState.rawGrades.forEach(row => {
        const name = row.이름;
        const grade = row.학년;
        const id = row.학번;
        if (!name || !id) return;
        if (!nameIdsMap[name]) nameIdsMap[name] = {};
        if (!nameIdsMap[name][grade]) nameIdsMap[name][grade] = new Set();
        nameIdsMap[name][grade].add(id);
    });

    const idCorrectionMap = {}; // 이름 -> { 임시학번: 실제학번 }
    Object.keys(nameIdsMap).forEach(name => {
        const gradesMap = nameIdsMap[name];
        const allIds = [];
        Object.keys(gradesMap).forEach(grade => {
            gradesMap[grade].forEach(id => {
                allIds.push({ grade: parseInt(grade), id: id, isTemp: id.substring(1, 3) === "99" });
            });
        });
        
        const tempIds = allIds.filter(item => item.isTemp);
        const realIds = allIds.filter(item => !item.isTemp);
        
        if (tempIds.length > 0 && realIds.length > 0) {
            const sortedReal = realIds.sort((a, b) => b.grade - a.grade);
            const refItem = sortedReal[0];
            const actualClassAndRoll = refItem.id.substring(1); // 예: '0324'
            
            tempIds.forEach(item => {
                const sameGradeReal = realIds.find(r => r.grade === item.grade);
                let correctedId = "";
                if (sameGradeReal) {
                    correctedId = sameGradeReal.id;
                } else {
                    correctedId = item.grade + actualClassAndRoll; // 예: '1' + '0324'
                }
                
                if (!idCorrectionMap[name]) idCorrectionMap[name] = {};
                idCorrectionMap[name][item.id] = correctedId;
            });
        }
    });

    // 원본 레코드들의 학번을 사전에 일제히 정규화
    AppState.rawGrades.forEach(row => {
        const name = row.이름;
        const id = row.학번;
        if (name && id && idCorrectionMap[name] && idCorrectionMap[name][id]) {
            row.학번 = idCorrectionMap[name][id];
        }
    });

    // 1단계: 동명이인 사전 분석
    // 동일 이름 내에서 같은 학년(예: 1학년)인데 학번이 다른 케이스가 발생하면 동명이인으로 확정 처리
    const nameGradeIdMap = {}; // 이름 -> { 학년: Set(학번) }

    AppState.rawGrades.forEach(row => {
        const name = row.이름;
        const grade = row.학년;
        const id = row.학번;

        if (!nameGradeIdMap[name]) {
            nameGradeIdMap[name] = {};
        }
        if (!nameGradeIdMap[name][grade]) {
            nameGradeIdMap[name][grade] = new Set();
        }
        if (id) {
            nameGradeIdMap[name][grade].add(id);
        }
    });

    // 동명이인으로 감지되어 자동 레이블링이 필요한 목록
    const homonymNames = new Set();
    Object.keys(nameGradeIdMap).forEach(name => {
        const grades = nameGradeIdMap[name];
        let hasConflict = false;
        Object.keys(grades).forEach(grade => {
            if (grades[grade].size > 1) {
                hasConflict = true; // 같은 학년에 여러 학번이 매칭된 경우 동명이인 의심
            }
        });
        if (hasConflict) {
            homonymNames.add(name);
        }
    });

    // 2단계: 실제 학생 객체 생성 및 데이터 매칭
    AppState.rawGrades.forEach(row => {
        let uniqueKey = row.이름;
        let isHomonym = false;

        // 원본 데이터 이름 자체에 이미 (A), (B) 등 구분자가 있거나, 
        // 1단계 분석에서 동일 학년 다른 학번으로 중복 감지된 경우
        if (homonymNames.has(row.이름)) {
            isHomonym = true;
            // 학번 정보가 있는 경우 이름 뒤에 학번을 기재하여 독립된 고유 키를 자동 부여
            const suffix = row.학번 ? `(${row.학번})` : `(${row.학년}학년_동명이인)`;
            uniqueKey = `${row.이름}${suffix}`;
        }

        if (!studentsMap[uniqueKey]) {
            studentsMap[uniqueKey] = {
                name: row.이름,
                displayName: uniqueKey,
                isHomonym: isHomonym,
                studentIds: {}, // 학년별 학번 맵 { 1: "10101", 2: "20101" }
                grades: [],     // 성적 데이터 원본 배열
                metrics: {
                    simpleAverage: 0,
                    weightedAverage: 0,
                    totalUnits: 0,
                    groupAverages: {}
                }
            };
        }

        const student = studentsMap[uniqueKey];
        student.grades.push(row);
        
        if (row.학번) {
            student.studentIds[row.학년] = row.학번;
        }
    });

    // 2.5단계: 전입생 임시 학번(99반) 동적 보정 로직
    // 전입학 학년 이전/이후에 배정받은 실제 학급 정보가 다른 학년에 존재한다면, 그 반/번호를 임시 학번에 이식하여 동기화
    Object.keys(studentsMap).forEach(key => {
        const student = studentsMap[key];
        const ids = student.studentIds;
        
        // 99반 임시 학번이 존재하는 학년 찾기
        const tempYears = Object.keys(ids).filter(yr => ids[yr] && ids[yr].substring(1, 3) === "99");
        if (tempYears.length > 0) {
            // 실제 학반 정보를 갖고 있는 기준 학년 찾기 (2학년 우선, 없으면 3학년, 없으면 1학년)
            let referenceId = "";
            for (const yr of [2, 3, 1]) {
                const id = ids[yr];
                if (id && id.substring(1, 3) !== "99") {
                    referenceId = id;
                    break;
                }
            }
            
            if (referenceId && referenceId.length >= 5) {
                const actualClassAndRoll = referenceId.substring(1); // 예: '0324' (반 2자리 + 번호 2자리)
                tempYears.forEach(yr => {
                    const originalId = ids[yr];
                    const correctedId = yr + actualClassAndRoll; // 예: '1' + '0324' = '10324'
                    ids[yr] = correctedId;
                    
                    // 해당 학년의 성적 레코드 학번도 일괄 보정
                    student.grades.forEach(row => {
                        if (row.학년.toString() === yr && row.학번 === originalId) {
                            row.학번 = correctedId;
                        }
                    });
                });
            }
        }
    });

    // 3단계: 학생별 성적 지표 산출
    Object.keys(studentsMap).forEach(key => {
        const student = studentsMap[key];
        calculateStudentMetrics(student);
    });

    AppState.students = studentsMap;

    // 콘솔에 동명이인 탐지 리포팅
    if (homonymNames.size > 0) {
        const alertMsg = `⚠️ 동명이인 감지: [${Array.from(homonymNames).join(", ")}] 학생의 데이터에 여러 학번이 감지되었습니다. 구분 분석을 위해 이름 뒤에 학번 식별자가 자동으로 부여되었습니다.`;
        console.warn(alertMsg);
        alert(alertMsg + "\n\n정확한 3개년 병합을 위해 가급적 원본 파일에 '이름(A)', '이름(B)' 등으로 구분 표기해 주시는 것을 권장합니다.");
    }

    // UI 새로고침
    renderStudentList();
    
    // 화면 전환
    document.getElementById("welcomeContainer").style.display = "none";
    document.getElementById("reportContainer").style.display = "block";
    document.getElementById("resetBtn").style.display = "inline-flex";
    document.getElementById("exportExcelBtn").style.display = "inline-flex";
    
    // 첫 번째 학생 자동 선택
    const firstStudentKey = Object.keys(studentsMap)[0];
    if (firstStudentKey) {
        selectStudent(firstStudentKey);
    }
}

// 특정 학생 성적 산출 로직
function calculateStudentMetrics(student) {
    const grades = student.grades;
    if (grades.length === 0) return;

    let sumRanks = 0;
    let sumWeightedRanks = 0;
    let sumUnits = 0;
    let numericGradesCount = 0;

    // 과목군 계산용 변수 초기화
    const groupSums = {};
    const groupUnits = {};
    Object.keys(SUBJECT_GROUPS).forEach(g => {
        groupSums[g] = 0;
        groupUnits[g] = 0;
    });

    grades.forEach(row => {
        const rank = row.등급;
        const units = row.단위수;
        const group = row.과목군;

        if (rank === "P") {
            return; // Pass 과목은 등급 평균 연산에서 완전히 제외
        }

        sumRanks += rank;
        sumWeightedRanks += (rank * units);
        sumUnits += units;
        numericGradesCount++;

        if (groupSums[group] !== undefined) {
            groupSums[group] += (rank * units);
            groupUnits[group] += units;
        }
    });

    // 1. 단순 등급 평균 (등급제 과목만 계산)
    student.metrics.simpleAverage = numericGradesCount > 0 ? parseFloat((sumRanks / numericGradesCount).toFixed(2)) : 9.0;

    // 2. 단위수 반영 가중 평균 등급 (등급제 과목만 계산)
    student.metrics.weightedAverage = sumUnits > 0 ? parseFloat((sumWeightedRanks / sumUnits).toFixed(2)) : 9.0;
    
    // 3. 총 이수 단위수 계산 (P 이수 과목 단위수까지 전체 통합)
    let totalAllUnits = 0;
    grades.forEach(row => {
        totalAllUnits += row.단위수;
    });
    student.metrics.totalUnits = totalAllUnits;

    // 3. 과목군별 등급 평균 (단위수 반영)
    student.metrics.groupAverages = {};
    Object.keys(SUBJECT_GROUPS).forEach(g => {
        if (groupUnits[g] > 0) {
            student.metrics.groupAverages[g] = parseFloat((groupSums[g] / groupUnits[g]).toFixed(2));
        } else {
            student.metrics.groupAverages[g] = null; // 미이수 과목군
        }
    });
}

// 학생의 반 정보 추출 (2학년 기준, 없으면 타학년)
function getStudentClass(student, targetYear = 2) {
    const id = student.studentIds[targetYear];
    if (id && id.length >= 5) {
        return parseInt(id.substring(1, 3));
    }
    for (const yr of [1, 3]) {
        const fallbackId = student.studentIds[yr];
        if (fallbackId && fallbackId.length >= 5) {
            return parseInt(fallbackId.substring(1, 3));
        }
    }
    return 999; // 정보 부재 시 맨 아래로
}

// 학생의 누락된 학년/학기 데이터 추출
function getMissingSemestersForStudent(student, systemYearSemesters) {
    const missing = []; // { year, semester }
    const activeYears = Object.keys(student.studentIds).map(y => parseInt(y)).sort((a, b) => a - b);
    const latestYear = activeYears[activeYears.length - 1] || 1;
    
    for (let yr = 1; yr <= latestYear; yr++) {
        // 해당 학년이 아예 학번 자체가 없다면 (전체 누락)
        if (!student.studentIds[yr]) {
            const expectedSems = systemYearSemesters[yr] || new Set([1, 2]);
            expectedSems.forEach(sem => {
                missing.push({ year: yr, semester: sem, totalMissing: true });
            });
            continue;
        }
        
        // 학번은 있는데 특정 학기가 빠져 있는 경우 (부분 누락)
        const expectedSems = systemYearSemesters[yr] || new Set([1, 2]);
        const studentSems = new Set(
            student.grades
                .filter(row => row.학년 === yr && row.등급 !== "P")
                .map(row => row.학기 || 1)
        );
        
        expectedSems.forEach(sem => {
            if (!studentSems.has(sem)) {
                missing.push({ year: yr, semester: sem, totalMissing: false });
            }
        });
    }
    return missing;
}

// 학생 리스트 표 렌더링
function renderStudentList() {
    const tbody = document.getElementById("studentListBody");
    tbody.innerHTML = "";

    const studentsKeys = Object.keys(AppState.students);
    document.getElementById("studentCount").textContent = studentsKeys.length;

    if (studentsKeys.length === 0) {
        tbody.innerHTML = `<tr><td colspan="3" class="empty-state">학생 데이터가 없습니다.</td></tr>`;
        return;
    }

    // 시스템 전체에서 학년별로 존재하는 학기 파악
    const systemYearSemesters = {};
    Object.values(AppState.students).forEach(st => {
        st.grades.forEach(row => {
            const yr = row.학년;
            const sem = row.학기 || 1;
            if (!systemYearSemesters[yr]) {
                systemYearSemesters[yr] = new Set();
            }
            systemYearSemesters[yr].add(sem);
        });
    });

    // 정렬 알고리즘 적용 (학번순/이름순/성적순 등 정렬 적용, 단 전출생은 항상 하단 배치)
    const sortBy = document.getElementById("studentSortSelect") ? document.getElementById("studentSortSelect").value : "name";
    
    // 시스템 전체에서 가장 높은 학년(현재 학년) 구하기
    let systemMaxGrade = 1;
    Object.values(AppState.students).forEach(st => {
        Object.keys(st.studentIds).forEach(yr => {
            const yVal = parseInt(yr);
            if (yVal > systemMaxGrade) systemMaxGrade = yVal;
        });
    });

    studentsKeys.sort((aKey, bKey) => {
        const a = AppState.students[aKey];
        const b = AppState.students[bKey];
        
        // 최고 학년 학번을 가진 학생(재학생, 전입생)과 없는 학생(전출생)의 배치 구분
        const aHasMaxId = a.studentIds[systemMaxGrade] !== undefined;
        const bHasMaxId = b.studentIds[systemMaxGrade] !== undefined;
        
        if (aHasMaxId && !bHasMaxId) return -1;
        if (!aHasMaxId && bHasMaxId) return 1;
        
        // 둘 다 속한 그룹 내에서 선택한 기준으로 정렬
        if (sortBy === "id") {
            const aId = getHighestStudentId(a);
            const bId = getHighestStudentId(b);
            if (!aId && bId) return 1;
            if (aId && !bId) return -1;
            if (!aId && !bId) {
                return a.name.localeCompare(b.name, "ko");
            }
            return aId.localeCompare(bId);
        } else if (sortBy === "score") {
            return a.metrics.weightedAverage - b.metrics.weightedAverage;
        } else {
            return a.name.localeCompare(b.name, "ko");
        }
    });

    studentsKeys.forEach(key => {
        const student = AppState.students[key];
        const row = document.createElement("tr");
        row.dataset.key = key;
        
        if (AppState.activeStudent && AppState.activeStudent.displayName === key) {
            row.classList.add("active");
        }

        const activeYears = Object.keys(student.studentIds).map(y => parseInt(y)).sort((a, b) => a - b);
        const latestYear = activeYears[activeYears.length - 1] || 1;
        const latestId = student.studentIds[latestYear] || "";

        const missing = getMissingSemestersForStudent(student, systemYearSemesters);
        
        // 학년 이수 및 학기 누락 정보 뱃지 동적 생성
        let gradeBadgesHtml = "";
        const missingLabels = [];
        const isExcluded = latestYear < systemMaxGrade;

        for (let yr = 1; yr <= latestYear; yr++) {
            const isYearMissing = !student.studentIds[yr];
            const yrMissing = missing.filter(m => m.year === yr);
            
            if (!isYearMissing && yrMissing.length === 0) {
                gradeBadgesHtml += `<span class="grade-dot g${yr}">${yr}학년</span> `;
            } else {
                let badgeTitle = "";
                if (isYearMissing) {
                    badgeTitle = `${yr}학년 전체 성적 누락 (전입생 추정)`;
                    if (!isExcluded) missingLabels.push(`${yr}학년 전체`);
                } else {
                    const missingSemText = yrMissing.map(m => `${m.semester}학기`).join(", ");
                    badgeTitle = `${yr}학년 ${missingSemText} 성적 누락`;
                    if (!isExcluded) missingLabels.push(`${yr}학년 ${missingSemText}`);
                }
                
                // 전출생(제외)인 경우 빨간 불을 켜지 않고 평범하게 표시
                if (isExcluded) {
                    if (!isYearMissing) {
                        gradeBadgesHtml += `<span class="grade-dot g${yr}">${yr}학년</span> `;
                    }
                } else {
                    // 활성 학생이면서 누락된 경우 테두리 없이 옅은 빨간 배경과 빨간 글씨만 적용 (크기 일관성 유지)
                    gradeBadgesHtml += `<span class="grade-dot" style="border: none; background-color: rgba(239, 68, 68, 0.1); color: #ef4444; font-weight: bold;" title="${badgeTitle}">${yr}학년</span> `;
                }
            }
        }

        let statusBadgeHtml = "";
        if (isExcluded) {
            statusBadgeHtml = `<span class="badge" style="background-color:rgba(100,116,139,0.1);color:#64748b;font-size:0.65rem;border:1px solid rgba(100,116,139,0.2);" title="전출/자퇴 추정: 이후 학년 성적 없음">제외</span>`;
        } else if (missingLabels.length > 0) {
            statusBadgeHtml = `<span class="badge" style="background-color:rgba(239,68,68,0.1);color:#ef4444;font-size:0.65rem;border:1px solid rgba(239,68,68,0.2);cursor:help;" title="누락 내역: ${missingLabels.join(', ')}">누락</span>`;
        }

        row.innerHTML = `
            <td>
                <strong>${student.name}</strong>
                ${student.isHomonym ? `<span class="badge" style="background-color:rgba(239,68,68,0.1);color:var(--danger);font-size:0.65rem;">식별됨</span>` : ""}
                ${statusBadgeHtml}
                <div style="font-size:0.7rem;color:var(--text-secondary);margin-top:2px;">
                    ${latestId ? latestId : ""}
                </div>
            </td>
            <td><div class="student-badge-group">${gradeBadgesHtml}</div></td>
            <td><span class="score-text">${student.metrics.weightedAverage} 등급</span></td>
        `;

        row.addEventListener("click", () => {
            selectStudent(key);
        });

        tbody.appendChild(row);
    });
}

// 학생 리스트 검색 필터
function filterStudentList() {
    const query = document.getElementById("searchBar").value.toLowerCase().trim();
    const rows = document.querySelectorAll("#studentListBody tr");
    
    rows.forEach(row => {
        const key = row.dataset.key;
        if (!key) return;
        
        const student = AppState.students[key];
        // 이름 또는 학번 정보 검색 매칭
        const idsString = Object.values(student.studentIds).join(" ");
        const match = student.name.toLowerCase().includes(query) || 
                      key.toLowerCase().includes(query) ||
                      idsString.includes(query);
                      
        row.style.display = match ? "" : "none";
    });
}

// 학생 선택 동작
function selectStudent(key) {
    const student = AppState.students[key];
    if (!student) return;

    AppState.activeStudent = student;

    // 테이블 행 활성화 토글
    document.querySelectorAll("#studentListBody tr").forEach(row => {
        if (row.dataset.key === key) {
            row.classList.add("active");
        } else {
            row.classList.remove("active");
        }
    });

    // 1. 프로필 정보 렌더링
    document.getElementById("reportStudentName").textContent = student.name;
    document.getElementById("studentAvatar").textContent = student.name.charAt(0);

    // 학급 이수 학년 상태 뱃지 동적 업데이트
    const activeYears = Object.keys(student.studentIds).sort((a, b) => a - b);
    let statusText = "성적 분석 중";
    if (activeYears.length === 1) {
        statusText = `${activeYears[0]}학년 성적 분석 중`;
    } else if (activeYears.length > 1) {
        statusText = `${activeYears[0]}~${activeYears[activeYears.length - 1]}학년 통합 분석 완료`;
    }
    document.getElementById("reportStatusBadge").textContent = statusText;
    
    const idHistoryContainer = document.getElementById("reportStudentIdHistory");
    idHistoryContainer.innerHTML = "";
    
    // 시스템 전체 학기 수집
    const systemYearSemesters = {};
    Object.values(AppState.students).forEach(st => {
        st.grades.forEach(row => {
            const yr = row.학년;
            const sem = row.학기 || 1;
            if (!systemYearSemesters[yr]) {
                systemYearSemesters[yr] = new Set();
            }
            systemYearSemesters[yr].add(sem);
        });
    });

    const missing = getMissingSemestersForStudent(student, systemYearSemesters);
    const years = [1, 2, 3];
    const sortedDescYears = Object.keys(student.studentIds).sort((a, b) => b - a);
    const latestYear = parseInt(sortedDescYears[0] || "1");

    years.forEach(yr => {
        const id = student.studentIds[yr];
        const tag = document.createElement("span");
        tag.className = "id-tag";
        
        const yrMissing = missing.filter(m => m.year === yr);
        
        if (id && yrMissing.length === 0) {
            tag.innerHTML = `<strong>${yr}학년:</strong> ${id}`;
        } else {
            // 해당 학년에 누락이 있거나 학년 자체가 비어 있고 최고 학년 이하인 경우 경고 출력
            if (yr <= latestYear) {
                tag.style.borderColor = "#ef4444";
                tag.style.backgroundColor = "rgba(239, 68, 68, 0.03)";
                
                if (!id) {
                    tag.innerHTML = `<strong>${yr}학년:</strong> <span style="color:#ef4444;font-weight:600;">미입력 (자료 누락) ⚠️</span>`;
                } else {
                    const missingSemText = yrMissing.map(m => `${m.semester}학기`).join(", ");
                    tag.innerHTML = `<strong>${yr}학년:</strong> ${id} <span style="color:#ef4444;font-weight:600;">(${missingSemText} 누락) ⚠️</span>`;
                }
            } else {
                tag.innerHTML = `<strong>${yr}학년:</strong> 미기록`;
            }
        }
        idHistoryContainer.appendChild(tag);
    });

    // 2. 성적 요약 카드 데이터 바인딩
    document.getElementById("simpleAverageValue").textContent = `${student.metrics.simpleAverage} 등급`;
    document.getElementById("weightedAverageValue").textContent = `${student.metrics.weightedAverage} 등급`;
    document.getElementById("totalUnitsValue").textContent = `${student.metrics.totalUnits} 단위`;

    // 3. 상세 등급 내역 테이블 렌더링 (필터 유지하며 갱신)
    renderDetailedGrades();

    // 4. 시각화 차트 렌더링
    renderCharts(student);
}

// 상세 등급 내역 테이블 렌더링
function renderDetailedGrades() {
    const student = AppState.activeStudent;
    if (!student) return;

    const gradeFilter = document.getElementById("gradeFilterSelect").value;
    const groupFilter = document.getElementById("groupFilterSelect").value;

    const tbody = document.getElementById("detailedGradesBody");
    tbody.innerHTML = "";

    // 필터 조건에 맞춰 필터링
    const filteredGrades = student.grades.filter(row => {
        let matchGrade = true;
        if (gradeFilter !== "all") {
            const parts = gradeFilter.split("-");
            const filterGrade = parseInt(parts[0]);
            const filterSemester = parts[1]; // "all", "1", "2"
            
            const matchesYr = row.학년 === filterGrade;
            let matchesSem = true;
            if (filterSemester !== "all") {
                const semVal = row.학기 || 1;
                matchesSem = (semVal.toString() === filterSemester);
            }
            matchGrade = matchesYr && matchesSem;
        }
        const matchGroup = (groupFilter === "all") || (row.과목군 === groupFilter);
        return matchGrade && matchGroup;
    });

    if (filteredGrades.length === 0) {
        tbody.innerHTML = `<tr><td colspan="7" class="empty-state">해당 필터 조건에 부합하는 성적 내역이 없습니다.</td></tr>`;
        return;
    }

    // 학년, 학기, 과목군, 과목명 순으로 오름차순 정렬 (이수 P 과목도 자연스럽게 함께 정렬)
    filteredGrades.sort((a, b) => a.학년 - b.학년 || a.학기 - b.학기 || a.과목군.localeCompare(b.과목군) || a.과목명.localeCompare(b.과목명));

    // 등급별 고유 색상 매핑 표 (1등급-빨강, 2등급-노랑, 3등급-초록, 4등급-파랑, 5등급-보라 대응)
    const GRADE_COLORS = {
        1: "#ef4444", // 빨강
        2: "#eab308", // 노랑
        3: "#10b981", // 초록
        4: "#3b82f6", // 파랑
        5: "#8b5cf6", // 보라
        6: "#6366f1", // 남색/인디고
        7: "#ec4899", // 핑크
        8: "#475569", // 진회색
        9: "#64748b"  // 회색
    };

    // 성취도별 고유 색상 매핑 표 (A-빨강, B-노랑, C-초록, D-파랑, E-보라 대응)
    const ACHIEVEMENT_COLORS = {
        "A": "#ef4444", // 빨강
        "B": "#eab308", // 노랑
        "C": "#10b981", // 초록
        "D": "#3b82f6", // 파랑
        "E": "#8b5cf6", // 보라
        "P": "#64748b"  // Slate 회색
    };

    filteredGrades.forEach(row => {
        const tr = document.createElement("tr");
        
        // 등급별 진척도 시각화 계산 (P 과목 예외 처리 포함)
        const isPass = row.등급 === "P";
        const maxLimit = getMaxRankLimit();
        const percentage = isPass ? 0 : ((maxLimit + 1 - row.등급) / maxLimit) * 100;
        
        // 등급 색상과 성취도 색상을 상호 비교 가능하도록 매핑 적용
        const gradeColor = isPass ? "#64748b" : (GRADE_COLORS[row.등급] || "#cbd5e1");
        const achColor = ACHIEVEMENT_COLORS[row.성취도] || "#cbd5e1";
        
        const rankText = isPass ? "P" : `${row.등급}`;
        const scoreText = isPass ? "이수 (등급 제외)" : `${row.등급}등급`;

        const grpMeta = SUBJECT_GROUPS[row.과목군] || { class: "etc" };

        // 과목군 수정 기능이 추가된 드롭다운 구조
        const selectHtml = `
            <select class="subject-group-select ${grpMeta.class}" data-index="${student.grades.indexOf(row)}">
                <option value="국어" ${row.과목군 === "국어" ? "selected" : ""}>국어</option>
                <option value="수학" ${row.과목군 === "수학" ? "selected" : ""}>수학</option>
                <option value="영어" ${row.과목군 === "영어" ? "selected" : ""}>영어</option>
                <option value="사회" ${row.과목군 === "사회" ? "selected" : ""}>사회</option>
                <option value="과학" ${row.과목군 === "과학" ? "selected" : ""}>과학</option>
                <option value="기타" ${row.과목군 === "기타" ? "selected" : ""}>기타</option>
            </select>
        `;

        tr.innerHTML = `
            <td><strong>${row.학년}학년 ${row.학기 ? row.학기 + '학기' : '1학기'}</strong></td>
            <td><span class="text-muted">${row.학번 || "-"}</span></td>
            <td>${selectHtml}</td>
            <td>${row.과목명}</td>
            <td><strong>${row.단위수}</strong> 단위</td>
            <td><span class="achievement-badge" style="background-color: ${achColor};">${row.성취도 || '-'}</span></td>
            <td><span class="grade-badge ${isPass ? 'rank-P' : 'rank-' + row.등급}" style="background-color: ${gradeColor};">${rankText}</span></td>
            <td>
                <div class="progress-bar-container">
                    <div class="progress-bar-track">
                        <div class="progress-bar-fill" style="width: ${percentage}%; background-color: ${gradeColor};"></div>
                    </div>
                    <span class="score-text" style="font-size:0.75rem;">${scoreText}</span>
                </div>
            </td>
        `;
        tbody.appendChild(tr);
    });

    // 드롭다운 변경 이벤트 위임 처리 (글로벌 동기화 적용)
    tbody.querySelectorAll(".subject-group-select").forEach(select => {
        select.addEventListener("change", (e) => {
            const targetIdx = parseInt(e.target.dataset.index);
            const newGroup = e.target.value;
            const targetGrade = student.grades[targetIdx];
            
            if (targetGrade) {
                const subjectName = targetGrade.과목명;
                
                // 1. 전역 성적 레코드에서 동일 과목명을 가진 모든 항목의 과목군 업데이트
                AppState.rawGrades.forEach(g => {
                    if (g.과목명 === subjectName) {
                        g.과목군 = newGroup;
                    }
                });
                
                // 2. 현재 로드된 모든 학생 객체들의 개별 성적 배열도 동기화 및 메트릭 재산출
                Object.keys(AppState.students).forEach(key => {
                    const s = AppState.students[key];
                    s.grades.forEach(g => {
                        if (g.과목명 === subjectName) {
                            g.과목군 = newGroup;
                        }
                    });
                    calculateStudentMetrics(s);
                });
                
                // 3. UI 및 차트 리프레시
                renderStudentList();
                renderDetailedGrades(); // 모든 테이블 행 새로고침을 통해 드롭다운 상태를 동기화시킴
                renderCharts(student);
                
                // 평균 등급 등 카드 지표 새로고침
                document.getElementById("simpleAverageValue").textContent = `${student.metrics.simpleAverage} 등급`;
                document.getElementById("weightedAverageValue").textContent = `${student.metrics.weightedAverage} 등급`;
            }
        });
    });
}

// 최고 등급 한계값 판단 함수 (5등급제 vs 9등급제 자동 스케일)
function getMaxRankLimit() {
    let maxRank = 5;
    AppState.rawGrades.forEach(g => {
        if (g.등급 > 5) {
            maxRank = 9;
        }
    });
    return maxRank;
}

// 방사형 차트 포인트에 등급 숫자를 텍스트로 직접 출력하는 플러그인
const radarLabelsPlugin = {
    id: 'radarLabels',
    afterDatasetsDraw(chart, args, options) {
        const { ctx } = chart;
        ctx.save();
        ctx.font = 'bold 10px Inter, sans-serif';
        ctx.fillStyle = 'rgba(30, 58, 138, 0.9)'; // 남색 계열
        ctx.textAlign = 'center';
        ctx.textBaseline = 'bottom';
        
        chart.data.datasets.forEach((dataset, i) => {
            const meta = chart.getDatasetMeta(i);
            meta.data.forEach((point, index) => {
                const val = dataset.data[index];
                if (val === null || val === undefined) return;
                ctx.fillText(val.toFixed(2) + '등급', point.x, point.y - 6);
            });
        });
        ctx.restore();
    }
};

// 가로 막대 그래프 우측 빈 공간에 등급 텍스트를 출력하는 플러그인
const barLabelsPlugin = {
    id: 'barLabels',
    afterDatasetsDraw(chart, args, options) {
        const { ctx } = chart;
        ctx.save();
        ctx.font = 'bold 11px Inter, sans-serif';
        ctx.fillStyle = '#475569'; // 슬레이트 그레이
        ctx.textAlign = 'left';
        ctx.textBaseline = 'middle';
        
        const maxLimit = getMaxRankLimit();
        
        chart.data.datasets.forEach((dataset, i) => {
            const meta = chart.getDatasetMeta(i);
            meta.data.forEach((bar, index) => {
                const val = dataset.data[index];
                if (val === null || val === undefined) return;
                
                // 역산하여 실제 등급 텍스트 복구
                const actualRank = (maxLimit + 1) - val;
                const text = actualRank.toFixed(2) + '등급';
                
                // 막대 우측 8px 위치에 출력
                ctx.fillText(text, bar.x + 8, bar.y);
            });
        });
        ctx.restore();
    }
};

// Chart.js 시각화 렌더링 함수
function renderCharts(student) {
    const groups = ["국어", "수학", "영어", "사회", "과학", "기타"];
    const averages = groups.map(g => student.metrics.groupAverages[g]);
    const maxLimit = getMaxRankLimit();

    // 기존 차트 인스턴스 파괴 (메모리 누수 및 마우스 오버 시 깜빡임 방지)
    if (AppState.charts.radar) AppState.charts.radar.destroy();
    if (AppState.charts.bar) AppState.charts.bar.destroy();
    if (AppState.charts.trend) AppState.charts.trend.destroy();

    // 1. 방사형 차트 (Radar)
    const ctxRadar = document.getElementById("radarChart").getContext("2d");
    const validGroupCount = averages.filter(v => v !== null).length;
    
    if (validGroupCount < 3) {
        ctxRadar.clearRect(0, 0, 300, 300);
        document.getElementById("radarChart").style.display = "none";
        
        let placeholder = document.getElementById("radarPlaceholder");
        if (!placeholder) {
            placeholder = document.createElement("div");
            placeholder.id = "radarPlaceholder";
            placeholder.className = "empty-state";
            placeholder.style.position = "absolute";
            placeholder.textContent = "방사형 차트를 그리려면 이수한 과목군이 3개 이상이어야 합니다. (막대 그래프를 확인하세요)";
            document.getElementById("radarChart").parentElement.appendChild(placeholder);
        }
        placeholder.style.display = "block";
    } else {
        document.getElementById("radarChart").style.display = "block";
        const placeholder = document.getElementById("radarPlaceholder");
        if (placeholder) placeholder.style.display = "none";

        const chartData = averages.map(v => v === null ? null : v);

        AppState.charts.radar = new Chart(ctxRadar, {
            type: "radar",
            plugins: [radarLabelsPlugin], // 커스텀 등급 표시 플러그인 등록
            data: {
                labels: groups,
                datasets: [{
                    label: "과목군별 등급 (평균)",
                    data: chartData,
                    backgroundColor: "rgba(79, 70, 229, 0.15)",
                    borderColor: "rgba(79, 70, 229, 1)",
                    borderWidth: 2,
                    pointBackgroundColor: "rgba(79, 70, 229, 1)",
                    pointBorderColor: "#fff",
                    pointHoverBackgroundColor: "#fff",
                    pointHoverBorderColor: "rgba(79, 70, 229, 1)"
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                scales: {
                    r: {
                        min: 0.5,
                        max: maxLimit + 0.5,
                        reverse: true, // 아주 중요: 1등급이 바깥쪽(우수), 5등급/9등급이 중심(하위)
                        ticks: {
                            stepSize: 1,
                            font: { size: 10 },
                            callback: function(value) {
                                if (Number.isInteger(value) && value >= 1 && value <= maxLimit) {
                                    return value + "등급";
                                }
                                return "";
                            }
                        },
                        pointLabels: {
                            font: { size: 12, weight: 'bold' }
                        }
                    }
                },
                plugins: {
                    legend: { display: false }
                }
            }
        });
    }

    // 2. 가로 막대 차트 (Horizontal Bar) - 좌우 정방향 설계
    const ctxBar = document.getElementById("barChart").getContext("2d");
    
    // 데이터 역산: 1등급이 가장 긴 막대를 가질 수 있도록 대입
    // 예: 5등급제의 경우, 1등급 -> 5로 변환 (최장 막대), 5등급 -> 1로 변환 (최단 막대)
    const transformedAverages = averages.map(val => val === null ? null : (maxLimit + 1) - val);
    
    const barColors = groups.map(g => SUBJECT_GROUPS[g].color);
    const borderColors = groups.map(g => SUBJECT_GROUPS[g].border);
    
    AppState.charts.bar = new Chart(ctxBar, {
        type: "bar",
        plugins: [barLabelsPlugin], // 커스텀 우측 등급 텍스트 표시 플러그인 등록
        data: {
            labels: groups,
            datasets: [{
                data: transformedAverages,
                backgroundColor: barColors,
                borderColor: borderColors,
                borderWidth: 1.5,
                borderRadius: 4
            }]
        },
        options: {
            indexAxis: 'y', // 가로 막대 그래프
            responsive: true,
            maintainAspectRatio: false,
            scales: {
                x: {
                    min: 0.5,
                    max: maxLimit + 0.8, // 우측 등급 텍스트가 잘리지 않도록 0.8 여유를 부여
                    reverse: false, // 좌우 역방향 해제 (정방향: 왼쪽 시작 -> 오른쪽 성장)
                    ticks: {
                        stepSize: 1,
                        callback: function(value) {
                            if (Number.isInteger(value) && value >= 1 && value <= maxLimit) {
                                // 스케일 축 눈금을 5등급제/9등급제에 따라 역산 기입
                                const actualRank = (maxLimit + 1) - value;
                                return actualRank + "등급";
                            }
                            return "";
                        }
                    },
                    title: {
                        display: true,
                        text: "등급 (오른쪽으로 갈수록 우수)",
                        font: { size: 10 }
                    }
                },
                y: {
                    position: 'left', // y축을 왼쪽에 정상적으로 배치
                    grid: { display: false },
                    ticks: {
                        font: { weight: 'bold' }
                    }
                }
            },
            plugins: {
                legend: { display: false },
                tooltip: {
                    callbacks: {
                        label: function(context) {
                            const val = context.raw;
                            if (val === null) return "이수 과목 없음";
                            // 툴팁 출력 시 원본 등급값으로 역산하여 출력
                            const originalRank = (maxLimit + 1) - val;
                            return `평균 ${originalRank.toFixed(2)} 등급`;
                        }
                    }
                }
            }
        }
    });

    // 3. 학기별 등급 변화 추이 (Line Chart)
    const ctxTrend = document.getElementById("trendChart").getContext("2d");

    const semesters = [
        { grade: 1, sem: 1 },
        { grade: 1, sem: 2 },
        { grade: 2, sem: 1 },
        { grade: 2, sem: 2 },
        { grade: 3, sem: 1 }
    ];

    const trendData = {
        labels: ["1학년 1학기", "1학년 2학기", "2학년 1학기", "2학년 2학기", "3학년 1학기"],
        overall: [],
        groups: {
            "국어": [],
            "수학": [],
            "영어": [],
            "사회": [],
            "과학": [],
            "기타": []
        }
    };

    semesters.forEach(s => {
        const semGrades = student.grades.filter(row => row.학년 === s.grade && (row.학기 || 1) === s.sem && row.등급 !== "P");
        
        if (semGrades.length === 0) {
            trendData.overall.push(null);
            Object.keys(trendData.groups).forEach(g => {
                trendData.groups[g].push(null);
            });
        } else {
            // 전체 가중 평균
            let sumVal = 0;
            let sumUnits = 0;
            semGrades.forEach(row => {
                sumVal += (row.등급 * row.단위수);
                sumUnits += row.단위수;
            });
            const overallAvg = sumUnits > 0 ? parseFloat((sumVal / sumUnits).toFixed(2)) : null;
            trendData.overall.push(overallAvg);

            // 각 과목군별 가중 평균
            Object.keys(trendData.groups).forEach(g => {
                const groupGrades = semGrades.filter(row => row.과목군 === g);
                if (groupGrades.length === 0) {
                    trendData.groups[g].push(null);
                } else {
                    let gSum = 0;
                    let gUnits = 0;
                    groupGrades.forEach(row => {
                        gSum += (row.등급 * row.단위수);
                        gUnits += row.단위수;
                    });
                    const gAvg = gUnits > 0 ? parseFloat((gSum / gUnits).toFixed(2)) : null;
                    trendData.groups[g].push(gAvg);
                }
            });
        }
    });

    // Datasets 생성
    const trendDatasets = [
        {
            label: "전체 평균",
            data: trendData.overall,
            borderColor: "#000000",
            backgroundColor: "rgba(0, 0, 0, 0.05)",
            borderWidth: 3.5,
            pointBackgroundColor: "#000000",
            pointBorderColor: "#ffffff",
            pointBorderWidth: 1.5,
            pointRadius: 5,
            pointHoverRadius: 7,
            tension: 0.15,
            spanGaps: true
        }
    ];

    groups.forEach(g => {
        trendDatasets.push({
            label: g,
            data: trendData.groups[g],
            borderColor: SUBJECT_GROUPS[g].border,
            backgroundColor: SUBJECT_GROUPS[g].color.replace("0.8", "0.05"),
            borderWidth: 2,
            pointBackgroundColor: SUBJECT_GROUPS[g].border,
            pointBorderColor: "#ffffff",
            pointBorderWidth: 1,
            pointRadius: 4,
            pointHoverRadius: 6,
            tension: 0.15,
            spanGaps: true
        });
    });

    AppState.charts.trend = new Chart(ctxTrend, {
        type: "line",
        data: {
            labels: trendData.labels,
            datasets: trendDatasets
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            scales: {
                y: {
                    min: 0.8,
                    max: maxLimit + 0.2,
                    reverse: true, // 1등급이 최상단, 5/9등급이 하단
                    ticks: {
                        stepSize: 1,
                        callback: function(value) {
                            if (value >= 1 && value <= maxLimit && Number.isInteger(value)) {
                                return value + "등급";
                            }
                            return "";
                        }
                    },
                    title: {
                        display: true,
                        text: "등급 (위로 갈수록 우수)"
                    }
                },
                x: {
                    grid: { display: false }
                }
            },
            plugins: {
                tooltip: {
                    mode: "index",
                    intersect: false,
                    callbacks: {
                        label: function(context) {
                            const val = context.raw;
                            if (val === null) return `${context.dataset.label}: 미이수`;
                            return `${context.dataset.label}: ${val.toFixed(2)}등급`;
                        }
                    }
                },
                legend: {
                    position: "top",
                    labels: {
                        usePointStyle: true,
                        font: { weight: "bold", size: 11 }
                    }
                }
            }
        }
    });
}

// 내장 샘플 데이터 세트
const MOCK_DATA = [];

// 샘플 데이터 로딩 구현
function loadSampleData() {
    // 폴더 일괄 선택 방식으로 대체되었습니다.
}

// 대시보드 데이터 초기화 기능
function resetDashboard() {
    // 1. 상태 데이터 비우기
    AppState.rawGrades = [];
    AppState.students = {};
    AppState.activeStudent = null;
    AppState.uploadedFiles = [];
    
    // 2. 인스턴스 차트 해제
    if (AppState.charts.radar) {
        AppState.charts.radar.destroy();
        AppState.charts.radar = null;
    }
    if (AppState.charts.bar) {
        AppState.charts.bar.destroy();
        AppState.charts.bar = null;
    }
    if (AppState.charts.trend) {
        AppState.charts.trend.destroy();
        AppState.charts.trend = null;
    }
    
    // 3. 파일 선택기 및 텍스트창 초기화
    document.getElementById("fileInput").value = "";
    document.getElementById("pasteInput").value = "";
    document.getElementById("searchBar").value = "";
    document.getElementById("gradeFilterSelect").value = "all";
    document.getElementById("groupFilterSelect").value = "all";
    
    // 파일 업로드 목록 UI 리셋
    const container = document.getElementById("loadedFilesContainer");
    if (container) container.style.display = "none";
    const listUl = document.getElementById("loadedFilesList");
    if (listUl) listUl.innerHTML = "";
    
    // 4. 학생 목록 초기 상태 복구
    const tbody = document.getElementById("studentListBody");
    tbody.innerHTML = `<tr><td colspan="3" class="empty-state">데이터를 입력하면 학생 목록이 표시됩니다.</td></tr>`;
    document.getElementById("studentCount").textContent = "0";
    
    // 5. 화면 구성 원래대로 복구 및 초기화/엑셀 버튼 숨기기
    document.getElementById("welcomeContainer").style.display = "flex";
    document.getElementById("reportContainer").style.display = "none";
    document.getElementById("resetBtn").style.display = "none";
    document.getElementById("exportExcelBtn").style.display = "none";
    
    // 방사형 차트 대체 플레이스홀더 잔재 정리
    const placeholder = document.getElementById("radarPlaceholder");
    if (placeholder) placeholder.remove();
}

// 학생별 성적 데이터 요약 엑셀 내보내기 기능
function exportToExcel() {
    if (!AppState.students || Object.keys(AppState.students).length === 0) {
        alert("내보낼 성적 데이터가 없습니다.");
        return;
    }

    const data = [];
    const sortBy = document.getElementById("studentSortSelect") ? document.getElementById("studentSortSelect").value : "name";
    const studentsKeys = Object.keys(AppState.students);
    
    studentsKeys.sort((aKey, bKey) => {
        const a = AppState.students[aKey];
        const b = AppState.students[bKey];
        if (sortBy === "id") {
            const aId = a.studentIds[2] || "";
            const bId = b.studentIds[2] || "";
            if (!aId && bId) return 1;
            if (aId && !bId) return -1;
            return aId.localeCompare(bId);
        } else if (sortBy === "class") {
            const aClass = getStudentClass(a);
            const bClass = getStudentClass(b);
            if (aClass !== bClass) return aClass - bClass;
            const aId = a.studentIds[2] || Object.values(a.studentIds)[0] || "";
            const bId = b.studentIds[2] || Object.values(b.studentIds)[0] || "";
            return aId.localeCompare(bId);
        } else if (sortBy === "score") {
            return a.metrics.weightedAverage - b.metrics.weightedAverage;
        } else {
            return a.name.localeCompare(b.name, "ko");
        }
    });

    studentsKeys.forEach(key => {
        const student = AppState.students[key];
        const id1 = student.studentIds[1] || "";
        const id2 = student.studentIds[2] || "";
        const class2 = id2 && id2.length >= 5 ? parseInt(id2.substring(1, 3)) : "";
        
        const avgKor = student.metrics.groupAverages["국어"] || "-";
        const avgMat = student.metrics.groupAverages["수학"] || "-";
        const avgEng = student.metrics.groupAverages["영어"] || "-";
        const avgSoc = student.metrics.groupAverages["사회"] || "-";
        const avgSci = student.metrics.groupAverages["과학"] || "-";
        const avgEtc = student.metrics.groupAverages["기타"] || "-";
        
        data.push({
            "이름": student.name,
            "1학년 학번": id1,
            "2학년 학번": id2,
            "2학년 반": class2 ? `${class2}반` : "",
            "가중 평균 등급": student.metrics.weightedAverage,
            "단순 평균 등급": student.metrics.simpleAverage,
            "총 이수 단위수": student.metrics.totalUnits,
            "국어 평균": avgKor,
            "수학 평균": avgMat,
            "영어 평균": avgEng,
            "사회 평균": avgSoc,
            "과학 평균": avgSci,
            "기타 평균": avgEtc
        });
    });

    const worksheet = XLSX.utils.json_to_sheet(data);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "성적요약");
    
    // 열 너비 자동 조절
    const max_widths = [];
    data.forEach(row => {
        Object.keys(row).forEach((key, col_idx) => {
            const val = String(row[key]);
            const length = Math.max(val.length * 2, key.length * 2) + 2;
            if (!max_widths[col_idx] || length > max_widths[col_idx]) {
                max_widths[col_idx] = length;
            }
        });
    });
    worksheet["!cols"] = max_widths.map(w => ({ wch: w }));

    XLSX.writeFile(workbook, "부산동성고_학생_내신등급_요약.xlsx");
}
