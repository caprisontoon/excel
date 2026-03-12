import React, { useState, useMemo, useEffect, useRef } from 'react';
import { 
  Plus, 
  Minus, 
  Trash2, 
  PaintBucket, 
  Sigma, 
  Calculator,
  AlignLeft,
  ArrowRight,
  ArrowDown,
  X,
  BarChart,
  Image as ImageIcon,
  Save,
  Copy,
  Check,
  Bold,
  Underline,
  Highlighter,
  Table,
  ExternalLink
} from 'lucide-react';

// --- 유틸리티 함수 ---

// 숫자를 엑셀 열 문자로 변환 (0 -> A, 25 -> Z, 26 -> AA)
const numToLetter = (num) => {
  let letter = '';
  while (num >= 0) {
    letter = String.fromCharCode((num % 26) + 65) + letter;
    num = Math.floor(num / 26) - 1;
  }
  return letter;
};

// 엑셀 열 문자를 숫자로 변환 (A -> 0, AA -> 26)
const letterToNum = (colStr) => {
  let col = 0;
  for (let i = 0; i < colStr.length; i++) {
    col = col * 26 + (colStr.charCodeAt(i) - 64);
  }
  return col - 1;
};

// 셀 참조 문자열 파싱 (예: "A1" -> {r: 0, c: 0})
const getCellRef = (refStr) => {
  const match = refStr.match(/^([A-Z]+)([0-9]+)$/);
  if (!match) return null;
  return { c: letterToNum(match[1]), r: parseInt(match[2]) - 1 };
};

// 범위 문자열 확장 (예: "A1:B2" -> [{r:0, c:0}, {r:0, c:1}, {r:1, c:0}, {r:1, c:1}])
const expandRange = (rangeStr) => {
  const [start, end] = rangeStr.split(':');
  const startRef = getCellRef(start);
  const endRef = getCellRef(end);
  if (!startRef || !endRef) return [];
  const cells = [];
  for (let r = Math.min(startRef.r, endRef.r); r <= Math.max(startRef.r, endRef.r); r++) {
    for (let c = Math.min(startRef.c, endRef.c); c <= Math.max(startRef.c, endRef.c); c++) {
      cells.push({ r, c });
    }
  }
  return cells;
};

// 초기 그리드 생성
const createEmptyGrid = (rows, cols) => {
  return Array.from({ length: rows }, () =>
    Array.from({ length: cols }, () => ({ 
      value: '', 
      color: '#ffffff', 
      textColor: 'inherit',
      fontFamily: 'sans-serif', 
      fontSize: 14,
      fontWeight: 'normal',
      textDecoration: 'none'
    }))
  );
};

// 조건부 서식 스타일 옵션
const condFormatOptions = [
  { id: 'lightRedFillDarkRedText', label: '진한 빨강 텍스트가 있는 연한 빨강 채우기', bg: '#ffc7ce', color: '#9c0006' },
  { id: 'yellowFillDarkYellowText', label: '진한 노랑 텍스트가 있는 노랑 채우기', bg: '#ffeb9c', color: '#9c6500' },
  { id: 'greenFillDarkGreenText', label: '진한 녹색 텍스트가 있는 녹색 채우기', bg: '#c6efce', color: '#006100' },
  { id: 'lightRedFill', label: '연한 빨강 채우기', bg: '#ffc7ce', color: 'inherit' },
  { id: 'redText', label: '빨강 텍스트', bg: 'transparent', color: '#ff0000' },
];

// 표 서식 (Table Formats) 옵션
const tableFormats = [
  // 밝게 (Light)
  { category: '밝게', id: 'light-1', headerBg: '#f3f4f6', headerText: '#111827', oddRowBg: '#ffffff', oddRowText: '#111827', evenRowBg: '#f9fafb', evenRowText: '#111827' },
  { category: '밝게', id: 'light-2', headerBg: '#ebf8ff', headerText: '#1e3a8a', oddRowBg: '#ffffff', oddRowText: '#111827', evenRowBg: '#f0f9ff', evenRowText: '#111827' },
  { category: '밝게', id: 'light-3', headerBg: '#fff7ed', headerText: '#9a3412', oddRowBg: '#ffffff', oddRowText: '#111827', evenRowBg: '#fffbf5', evenRowText: '#111827' },
  { category: '밝게', id: 'light-4', headerBg: '#f0fdf4', headerText: '#166534', oddRowBg: '#ffffff', oddRowText: '#111827', evenRowBg: '#fcfdfd', evenRowText: '#111827' },
  { category: '밝게', id: 'light-5', headerBg: '#faf5ff', headerText: '#6b21a8', oddRowBg: '#ffffff', oddRowText: '#111827', evenRowBg: '#fdfbfb', evenRowText: '#111827' },

  // 중간 (Medium)
  { category: '중간', id: 'medium-1', headerBg: '#6b7280', headerText: '#ffffff', oddRowBg: '#ffffff', oddRowText: '#111827', evenRowBg: '#f3f4f6', evenRowText: '#111827' },
  { category: '중간', id: 'medium-2', headerBg: '#3b82f6', headerText: '#ffffff', oddRowBg: '#ffffff', oddRowText: '#111827', evenRowBg: '#eff6ff', evenRowText: '#111827' },
  { category: '중간', id: 'medium-3', headerBg: '#f97316', headerText: '#ffffff', oddRowBg: '#ffffff', oddRowText: '#111827', evenRowBg: '#fff7ed', evenRowText: '#111827' },
  { category: '중간', id: 'medium-4', headerBg: '#22c55e', headerText: '#ffffff', oddRowBg: '#ffffff', oddRowText: '#111827', evenRowBg: '#f0fdf4', evenRowText: '#111827' },
  { category: '중간', id: 'medium-5', headerBg: '#a855f7', headerText: '#ffffff', oddRowBg: '#ffffff', oddRowText: '#111827', evenRowBg: '#faf5ff', evenRowText: '#111827' },

  // 어둡게 (Dark)
  { category: '어둡게', id: 'dark-1', headerBg: '#1f2937', headerText: '#ffffff', oddRowBg: '#d1d5db', oddRowText: '#111827', evenRowBg: '#9ca3af', evenRowText: '#111827' },
  { category: '어둡게', id: 'dark-2', headerBg: '#1e3a8a', headerText: '#ffffff', oddRowBg: '#bfdbfe', oddRowText: '#111827', evenRowBg: '#93c5fd', evenRowText: '#111827' },
  { category: '어둡게', id: 'dark-3', headerBg: '#9a3412', headerText: '#ffffff', oddRowBg: '#fed7aa', oddRowText: '#111827', evenRowBg: '#fdba74', evenRowText: '#111827' },
  { category: '어둡게', id: 'dark-4', headerBg: '#14532d', headerText: '#ffffff', oddRowBg: '#bbf7d0', oddRowText: '#111827', evenRowBg: '#86efac', evenRowText: '#111827' },
  { category: '어둡게', id: 'dark-5', headerBg: '#581c87', headerText: '#ffffff', oddRowBg: '#e9d5ff', oddRowText: '#111827', evenRowBg: '#d8b4fe', evenRowText: '#111827' },
];

export default function App() {
  const [activeCell, setActiveCell] = useState({ r: 0, c: 0 });
  const [editingCell, setEditingCell] = useState(null);
  const [selectedRowIndex, setSelectedRowIndex] = useState(null); // 행 전체 선택 상태
  const [selectedColIndex, setSelectedColIndex] = useState(null); // 열 전체 선택 상태
  const editInputRef = useRef(null);
  const fileInputRef = useRef(null);
  const tableRef = useRef(null);

  // --- 다중 시트 상태 관리 ---
  const [sheets, setSheets] = useState([
    { id: 'sheet-1', name: 'Sheet1', grid: createEmptyGrid(5, 5), colTitles: {}, bgImage: null, bgOpacity: 0.5, conditionalRules: [] }
  ]);
  const [activeSheetId, setActiveSheetId] = useState('sheet-1');
  const [editingSheetId, setEditingSheetId] = useState(null);
  const [editSheetName, setEditSheetName] = useState('');

  // 현재 활성화된 시트 데이터 파생
  const currentSheet = useMemo(() => sheets.find(s => s.id === activeSheetId) || sheets[0], [sheets, activeSheetId]);

  const grid = currentSheet.grid;
  const colTitles = currentSheet.colTitles;
  const bgImage = currentSheet.bgImage;
  const bgOpacity = currentSheet.bgOpacity;

  // 상태 업데이트 래퍼 함수 (기존 로직 호환성 유지 및 현재 시트 업데이트)
  const setGrid = (updater) => {
    setSheets(prev => prev.map(s => s.id === activeSheetId ? { ...s, grid: typeof updater === 'function' ? updater(s.grid) : updater } : s));
  };
  const setColTitles = (updater) => {
    setSheets(prev => prev.map(s => s.id === activeSheetId ? { ...s, colTitles: typeof updater === 'function' ? updater(s.colTitles) : updater } : s));
  };
  const setBgImage = (updater) => {
    setSheets(prev => prev.map(s => s.id === activeSheetId ? { ...s, bgImage: typeof updater === 'function' ? updater(s.bgImage) : updater } : s));
  };
  const setBgOpacity = (updater) => {
    setSheets(prev => prev.map(s => s.id === activeSheetId ? { ...s, bgOpacity: typeof updater === 'function' ? updater(s.bgOpacity) : updater } : s));
  };
  const setConditionalRules = (updater) => {
    setSheets(prev => prev.map(s => s.id === activeSheetId ? { ...s, conditionalRules: typeof updater === 'function' ? updater(s.conditionalRules || []) : updater } : s));
  };

  // --- 시트 제어 함수 ---
  const addSheet = () => {
    const newId = `sheet-${Date.now()}`;
    setSheets(prev => [...prev, {
      id: newId,
      name: `Sheet${prev.length + 1}`,
      grid: createEmptyGrid(5, 5),
      colTitles: {},
      bgImage: null,
      bgOpacity: 0.5,
      conditionalRules: []
    }]);
    setActiveSheetId(newId);
    setActiveCell({ r: 0, c: 0 });
    setSelectedRowIndex(null);
    setSelectedColIndex(null);
    setEditingCell(null);
  };

  const deleteSheet = (id, e) => {
    e.stopPropagation();
    if (sheets.length <= 1) return; // 마지막 남은 시트는 삭제 불가
    const newSheets = sheets.filter(s => s.id !== id);
    setSheets(newSheets);
    if (activeSheetId === id) {
      setActiveSheetId(newSheets[newSheets.length - 1].id); // 삭제 시 가용한 마지막 탭으로 이동
      setActiveCell({ r: 0, c: 0 });
      setSelectedRowIndex(null);
      setSelectedColIndex(null);
      setEditingCell(null);
    }
  };

  const saveSheetName = (id) => {
    if (editSheetName.trim()) {
      setSheets(prev => prev.map(s => s.id === id ? { ...s, name: editSheetName.trim() } : s));
    }
    setEditingSheetId(null);
  };

  // --- 표 서식(Table Format) 관련 상태 ---
  const [showTableFormatModal, setShowTableFormatModal] = useState(false);

  const applyTableFormat = (format) => {
    setGrid(prev => prev.map((row, r) => {
      return row.map((cell, c) => {
        let bg, text, weight;
        if (r === 0) {
          bg = format.headerBg;
          text = format.headerText;
          weight = 'bold';
        } else {
          // 짝수 인덱스(1, 3, 5...)는 홀수행, 홀수 인덱스(2, 4, 6...)는 짝수행으로 취급 (0은 헤더)
          const isOddRow = r % 2 !== 0; 
          bg = isOddRow ? format.oddRowBg : format.evenRowBg;
          text = isOddRow ? format.oddRowText : format.evenRowText;
          weight = cell.fontWeight; // 기존 굵기 유지
        }
        return { ...cell, color: bg, textColor: text, fontWeight: weight };
      });
    }));
    setShowTableFormatModal(false);
  };

  // --- 데이터 입력 관리 모달 상태 ---
  const [showDataInputModal, setShowDataInputModal] = useState(false);
  const [tempGridData, setTempGridData] = useState([]);
  
  // 모달 드래그를 위한 상태 및 참조
  const [dataModalPos, setDataModalPos] = useState({ x: 0, y: 0 });
  const dragRef = useRef({ isDragging: false, startX: 0, startY: 0, initialX: 0, initialY: 0 });

  // --- VS 팀 모드 상태 ---
  const [isVsMode, setIsVsMode] = useState(false);

  const openDataInputModal = () => {
    // 모달을 열 때 현재 그리드의 값을 임시 상태로 복사하고 위치 초기화
    setTempGridData(grid.map(row => row.map(cell => cell.value)));
    setDataModalPos({ x: 0, y: 0 }); 
    setShowDataInputModal(true);
  };

  const handleDragStart = (e) => {
    dragRef.current.isDragging = true;
    dragRef.current.startX = e.clientX;
    dragRef.current.startY = e.clientY;
    dragRef.current.initialX = dataModalPos.x;
    dragRef.current.initialY = dataModalPos.y;
    e.currentTarget.setPointerCapture(e.pointerId);
  };

  const handleDragMove = (e) => {
    if (!dragRef.current.isDragging) return;
    const dx = e.clientX - dragRef.current.startX;
    const dy = e.clientY - dragRef.current.startY;
    setDataModalPos({
      x: dragRef.current.initialX + dx,
      y: dragRef.current.initialY + dy
    });
  };

  const handleDragEnd = (e) => {
    dragRef.current.isDragging = false;
    e.currentTarget.releasePointerCapture(e.pointerId);
  };

  const handleTempDataChange = (r, c, val) => {
    setTempGridData(prev => {
      const next = [...prev];
      next[r] = [...next[r]];
      next[r][c] = val;
      return next;
    });
  };

  const saveDataInput = () => {
    // 임시 상태의 값을 실제 그리드에 병합 (기존 스타일 및 포맷은 유지)
    setGrid(prevGrid => {
      return prevGrid.map((row, r) => {
        return row.map((cell, c) => {
          return { ...cell, value: tempGridData[r][c] !== undefined ? tempGridData[r][c] : cell.value };
        });
      });
    });
    setShowDataInputModal(false);
  };

  // 모달에 표출할 유효한 컬럼 추출 (본창에 세팅된 열의 속성값 노출)
  const getActiveInputCols = () => {
    const cols = [];
    for (let i = 1; i < colsCount; i++) {
      if (colTitles[i] && colTitles[i] !== "선택") {
        cols.push({ index: i, label: colTitles[i] });
      }
    }
    // 만약 설정된 타이틀이 하나도 없으면 기본으로 B~E열(1~4) 노출
    if (cols.length === 0) {
      for (let i = 1; i <= 4 && i < colsCount; i++) {
        cols.push({ index: i, label: numToLetter(i) + '열' });
      }
    }
    return cols;
  };

  // --- HTML 내보내기 관련 상태 ---
  const [showExportModal, setShowExportModal] = useState(false);
  const [exportHtml, setExportHtml] = useState('');
  const [copySuccess, setCopySuccess] = useState(false);
  const [showDonationSetupModal, setShowDonationSetupModal] = useState(false);

  // --- 조건부 서식 관련 상태 ---
  const [showCondFormatModal, setShowCondFormatModal] = useState(false);
  const [condRule, setCondRule] = useState({ type: 'greaterThan', value: '', format: 'lightRedFillDarkRedText' });

  const handleApplyCondRule = () => {
    const targetCol = selectedColIndex !== null ? selectedColIndex : activeCell.c;
    if (condRule.value.trim() === '') return;
    
    setConditionalRules(prev => [...prev, { ...condRule, targetCol }]);
    setShowCondFormatModal(false);
    setCondRule({ type: 'greaterThan', value: '', format: 'lightRedFillDarkRedText' }); // 리셋
  };

  const handleClearCondRules = () => {
    const targetCol = selectedColIndex !== null ? selectedColIndex : activeCell.c;
    setConditionalRules(prev => prev.filter(r => r.targetCol !== targetCol));
    setShowCondFormatModal(false);
  };

  // 조건부 서식 평가 함수
  const checkCondition = (cellValue, rule) => {
    if (cellValue === undefined || cellValue === null || cellValue === '') return false;
    const valStr = String(cellValue);
    const ruleValStr = String(rule.value);
    const valNum = Number(cellValue);
    const ruleNum = Number(rule.value);

    switch (rule.type) {
      case 'greaterThan':
        return !isNaN(valNum) && !isNaN(ruleNum) && valNum > ruleNum;
      case 'lessThan':
        return !isNaN(valNum) && !isNaN(ruleNum) && valNum < ruleNum;
      case 'equalTo':
        return valStr === ruleValStr || (!isNaN(valNum) && !isNaN(ruleNum) && valNum === ruleNum);
      case 'textContains':
        return valStr.includes(ruleValStr);
      default:
        return false;
    }
  };

  // --- 컬럼 헤더(타이틀) 속성 상태 ---
  const [editingColTitle, setEditingColTitle] = useState(null); // 헤더 속성 편집 모드 상태 추가
  const headerOptions = ["선택", "후원점수", "기여도", "직급", "기타점수", "팀"];

  // --- 크루 모달 관련 상태 ---
  const [showCrewModal, setShowCrewModal] = useState(false);
  const [crews, setCrews] = useState([
    { id: 1, name: '김코딩', avatar: 'https://i.pravatar.cc/150?u=11' },
    { id: 2, name: '박해커', avatar: 'https://i.pravatar.cc/150?u=22' },
    { id: 3, name: '이디자인', avatar: 'https://i.pravatar.cc/150?u=33' },
    { id: 4, name: '최기획', avatar: 'https://i.pravatar.cc/150?u=44' },
    { id: 5, name: '정개발', avatar: 'https://i.pravatar.cc/150?u=55' },
    { id: 6, name: '윤테스트', avatar: 'https://i.pravatar.cc/150?u=66' },
  ]);

  const rowsCount = grid.length;
  const colsCount = grid[0]?.length || 0;

  // --- 배경 이미지 핸들러 ---
  const handleImageUpload = (e) => {
    const file = e.target.files[0];
    if (file) {
      const reader = new FileReader();
      reader.onload = (event) => {
        setBgImage(event.target.result);
      };
      reader.readAsDataURL(file);
    }
    // 동일한 파일 재선택 가능하게 value 초기화
    e.target.value = null;
  };

  // --- 크루 제어 함수 ---
  const handleCrewSelect = (crew) => {
    updateActiveCell({ value: crew.name, avatar: crew.avatar });
    setShowCrewModal(false);
  };

  const handleDeleteCrew = (e, id) => {
    e.stopPropagation();
    setCrews(prev => prev.filter(c => c.id !== id));
  };

  const handleAddCrew = () => {
    const newId = Date.now();
    setCrews(prev => [...prev, {
      id: newId,
      name: `새 크루 ${prev.length + 1}`,
      avatar: `https://i.pravatar.cc/150?u=${newId}`
    }]);
  };

  // --- HTML 내보내기(저장) ---
  const handleSave = () => {
    // 1. 선택 상태 및 편집 상태 해제하여 깔끔한 표 준비
    setEditingCell(null);
    setActiveCell({ r: -1, c: -1 });
    setSelectedRowIndex(null);
    setSelectedColIndex(null);

    // 2. React 렌더링이 완료된 후 DOM에서 HTML 추출
    setTimeout(() => {
      if (tableRef.current) {
        setExportHtml(tableRef.current.outerHTML);
        setShowExportModal(true);
      }
    }, 50);
  };

  const handleCopyHtml = () => {
    const textArea = document.createElement("textarea");
    textArea.value = exportHtml;
    document.body.appendChild(textArea);
    textArea.select();
    try {
      document.execCommand('copy');
      setCopySuccess(true);
      setTimeout(() => setCopySuccess(false), 2000); // 2초 후 초기화
    } catch (err) {
      console.error('복사 실패', err);
    }
    document.body.removeChild(textArea);
  };

  // --- 수식 평가 로직 ---
  const evaluate = (value, r, c, currentGrid, visited) => {
    if (value === undefined || value === null || value === '') return '';
    const strVal = String(value);
    
    // 수식이 아니면 일반 값 반환
    if (!strVal.startsWith('=')) {
      const num = Number(strVal);
      return isNaN(num) ? strVal : num;
    }

    const cellKey = `${r},${c}`;
    if (visited.has(cellKey)) return "#REF!"; // 순환 참조 에러
    visited.add(cellKey);

    let formula = strVal.substring(1).toUpperCase();

    try {
      // 1. 함수 파싱 (SUM, AVG, MAX, MIN)
      formula = formula.replace(/(SUM|AVG|MAX|MIN)\(([A-Z0-9:]+)\)/g, (match, func, range) => {
        let refs = range.includes(':') ? expandRange(range) : [getCellRef(range)];
        let vals = refs.map(ref => {
          if (ref && currentGrid[ref.r] && currentGrid[ref.r][ref.c]) {
            const val = evaluate(currentGrid[ref.r][ref.c].value, ref.r, ref.c, currentGrid, new Set(visited));
            const num = Number(val);
            return isNaN(num) ? null : num;
          }
          return null;
        }).filter(v => v !== null);

        if (vals.length === 0) return 0;
        if (func === 'SUM') return vals.reduce((a, b) => a + b, 0);
        if (func === 'AVG') return vals.reduce((a, b) => a + b, 0) / vals.length;
        if (func === 'MAX') return Math.max(...vals);
        if (func === 'MIN') return Math.min(...vals);
        return 0;
      });

      // 2. 단일 셀 참조 파싱 (예: A1 + B2)
      formula = formula.replace(/[A-Z]+[0-9]+/g, (match) => {
        const ref = getCellRef(match);
        if (ref && currentGrid[ref.r] && currentGrid[ref.r][ref.c]) {
          const val = evaluate(currentGrid[ref.r][ref.c].value, ref.r, ref.c, currentGrid, new Set(visited));
          const num = Number(val);
          return isNaN(num) ? 0 : num;
        }
        return 0;
      });

      // 3. 허용되지 않은 문자 필터링 (보안 및 에러 방지)
      if (/[^0-9+\-*/().\s]/.test(formula)) {
        return "#NAME?";
      }

      // 4. 수식 계산
      // eslint-disable-next-line no-new-func
      const result = new Function(`return ${formula}`)();
      return isFinite(result) ? result : "#DIV/0!";
    } catch (e) {
      return "#ERROR!";
    }
  };

  // 실시간 미리보기를 위한 파생 그리드 (모달 입력 중 실시간 반영)
  const liveGrid = useMemo(() => {
    if (!showDataInputModal || !tempGridData || tempGridData.length === 0) return grid;
    return grid.map((row, r) =>
      row.map((cell, c) => {
        const tempVal = tempGridData[r]?.[c];
        return {
          ...cell,
          value: tempVal !== undefined ? tempVal : cell.value
        };
      })
    );
  }, [grid, showDataInputModal, tempGridData]);

  // 화면에 렌더링될 계산된 그리드 (liveGrid 변경 시 즉각 재계산)
  const displayGrid = useMemo(() => {
    return liveGrid.map((row, r) =>
      row.map((cell, c) => ({
        ...cell,
        computed: evaluate(cell.value, r, c, liveGrid, new Set())
      }))
    );
  }, [liveGrid]);

  // --- 셀 제어 함수 ---
  const updateActiveCell = (updates) => {
    if (!activeCell) return;
    setGrid(prev => {
      const newGrid = [...prev];
      newGrid[activeCell.r] = [...newGrid[activeCell.r]];
      newGrid[activeCell.r][activeCell.c] = { ...newGrid[activeCell.r][activeCell.c], ...updates };
      return newGrid;
    });
  };

  const handleCellChange = (r, c, value) => {
    setGrid(prev => {
      const newGrid = [...prev];
      newGrid[r] = [...newGrid[r]];
      newGrid[r][c] = { ...newGrid[r][c], value, avatar: null }; // 직접 편집 시 아바타 초기화
      return newGrid;
    });
  };

  // --- 행/열 추가 삭제 기능 ---
  const addRow = (direction) => {
    const targetIdx = direction === 'above' ? activeCell.r : activeCell.r + 1;
    const newRow = Array.from({ length: colsCount }, () => ({ 
      value: '', 
      color: '#ffffff',
      textColor: 'inherit',
      fontFamily: 'sans-serif', 
      fontSize: 14,
      fontWeight: 'normal',
      textDecoration: 'none'
    }));
    setGrid(prev => {
      const newGrid = [...prev];
      newGrid.splice(targetIdx, 0, newRow);
      return newGrid;
    });
    if (direction === 'above') setActiveCell(p => ({ ...p, r: p.r + 1 }));
    setSelectedRowIndex(null); // 행 추가 시 행 선택 초기화
  };

  const addCol = (direction) => {
    const targetIdx = direction === 'left' ? activeCell.c : activeCell.c + 1;
    setGrid(prev => prev.map(row => {
      const newRow = [...row];
      newRow.splice(targetIdx, 0, { 
        value: '', 
        color: '#ffffff',
        textColor: 'inherit',
        fontFamily: 'sans-serif', 
        fontSize: 14,
        fontWeight: 'normal',
        textDecoration: 'none'
      });
      return newRow;
    }));
    if (direction === 'left') setActiveCell(p => ({ ...p, c: p.c + 1 }));
  };

  const deleteRow = () => {
    // 선택된 행이 없으면 작동하지 않음
    if (selectedRowIndex === null) return;
    if (rowsCount <= 1) return;
    
    setGrid(prev => prev.filter((_, i) => i !== selectedRowIndex));
    setSelectedRowIndex(null);
    setActiveCell(p => ({ ...p, r: Math.max(0, selectedRowIndex - 1), c: 0 }));
  };

  const deleteCol = () => {
    // 선택된 열이 없으면 작동하지 않음
    if (selectedColIndex === null) return;
    if (colsCount <= 1) return;
    
    setGrid(prev => prev.map(row => row.filter((_, i) => i !== selectedColIndex)));
    setSelectedColIndex(null);
    setActiveCell(p => ({ ...p, c: Math.max(0, selectedColIndex - 1), r: 0 }));
  };

  // --- 정렬 기능 ---
  const sortBySponsorship = () => {
    // 1. '후원점수'가 선택된 컬럼의 인덱스 찾기
    const targetColStr = Object.keys(colTitles).find(key => colTitles[key] === "후원점수");
    if (!targetColStr) return;
    
    const colIdx = parseInt(targetColStr, 10);

    setGrid(prev => {
      // 2. 각 행의 후원점수 값을 계산하여 정렬용 배열 생성
      const rowScores = prev.map((row, r) => {
        const computed = evaluate(row[colIdx].value, r, colIdx, prev, new Set());
        const num = Number(computed);
        const isCrewEmpty = !row[0].value; // A열(크루 이름)이 비어있는지 확인
        return { row, score: isNaN(num) ? -999999 : num, isEmpty: isCrewEmpty };
      });

      // 3. 점수가 높은 순(내림차순)으로 정렬, 빈 행은 최하단으로 이동
      rowScores.sort((a, b) => {
        if (a.isEmpty && b.isEmpty) return 0;
        if (a.isEmpty) return 1;
        if (b.isEmpty) return -1;
        return b.score - a.score;
      });

      return rowScores.map(item => item.row);
    });
    
    setEditingCell(null);
    setSelectedRowIndex(null);
  };

  // --- 자동 함수 기능 ---
  const applyAutoFunction = (funcName) => {
    if (activeCell.r === 0) return; // 첫 번째 행이면 적용할 범위가 없음
    const colLetter = numToLetter(activeCell.c);
    const startCell = `${colLetter}1`;
    const endCell = `${colLetter}${activeCell.r}`;
    updateActiveCell({ value: `=${funcName}(${startCell}:${endCell})` });
  };

  // --- 렌더링 관련 이펙트 ---
  useEffect(() => {
    if (editingCell && editInputRef.current) {
      editInputRef.current.focus();
    }
  }, [editingCell]);

  // --- VS 팀 모드 데이터 계산 ---
  const getVsTeamsData = () => {
    const teamColStr = Object.keys(colTitles).find(key => colTitles[key] === "팀");
    // 후원점수가 없으면 기여도를 점수로 대체
    const scoreColStr = Object.keys(colTitles).find(key => colTitles[key] === "후원점수" || colTitles[key] === "기여도");
    
    const teamColIdx = teamColStr ? parseInt(teamColStr, 10) : -1;
    const scoreColIdx = scoreColStr ? parseInt(scoreColStr, 10) : -1;

    // 크루 이름이 있는 유효한 행만 필터링
    const validRows = displayGrid.filter(row => row[0].computed !== '');

    const teamsMap = {};

    if (teamColIdx !== -1) {
      // 1. '팀' 열이 지정된 경우 해당 값을 기준으로 그룹화
      validRows.forEach(row => {
        let teamName = row[teamColIdx].computed || '미지정팀';
        if (!teamsMap[teamName]) teamsMap[teamName] = { name: teamName, score: 0, members: [] };
        
        const score = scoreColIdx !== -1 ? Number(row[scoreColIdx].computed) || 0 : 0;
        teamsMap[teamName].score += score;
        teamsMap[teamName].members.push({
          name: row[0].computed,
          avatar: row[0].avatar,
          score: score
        });
      });
    } else {
      // 2. '팀' 열이 없는 경우 반으로 나누어 임의 배정
      teamsMap['A'] = { name: 'A', score: 0, members: [] };
      teamsMap['B'] = { name: 'B', score: 0, members: [] };
      validRows.forEach((row, idx) => {
        const teamName = idx % 2 === 0 ? 'A' : 'B';
        const score = scoreColIdx !== -1 ? Number(row[scoreColIdx].computed) || 0 : 0;
        teamsMap[teamName].score += score;
        teamsMap[teamName].members.push({
          name: row[0].computed,
          avatar: row[0].avatar,
          score: score
        });
      });
    }

    const teamKeys = Object.keys(teamsMap);
    const team1 = teamsMap[teamKeys[0]] || { name: 'A', score: 0, members: [] };
    const team2 = teamsMap[teamKeys[1]] || { name: 'B', score: 0, members: [] };

    return [team1, team2];
  };

  // --- 스타일 변경 제어 로직 ---
  const isStylePickerEnabled = selectedRowIndex !== null || selectedColIndex !== null;

  const getSelectedCellStyle = (key, defaultValue) => {
    if (selectedRowIndex !== null) return grid[selectedRowIndex]?.[0]?.[key] || defaultValue;
    if (selectedColIndex !== null) return grid[0]?.[selectedColIndex]?.[key] || defaultValue;
    return defaultValue;
  };

  const currentColor = getSelectedCellStyle('color', '#ffffff');
  const currentFontFamily = getSelectedCellStyle('fontFamily', 'sans-serif');
  const currentFontSize = getSelectedCellStyle('fontSize', 14);
  const currentFontWeight = getSelectedCellStyle('fontWeight', 'normal');
  const currentTextDecoration = getSelectedCellStyle('textDecoration', 'none');

  const handleStyleChange = (key, value) => {
    if (selectedRowIndex !== null) {
      // 행 전체 스타일 변경
      setGrid(prev => {
        const newGrid = [...prev];
        newGrid[selectedRowIndex] = newGrid[selectedRowIndex].map(cell => ({ ...cell, [key]: value }));
        return newGrid;
      });
    } else if (selectedColIndex !== null) {
      // 열 전체 스타일 변경
      setGrid(prev => prev.map(row => {
        const newRow = [...row];
        newRow[selectedColIndex] = { ...newRow[selectedColIndex], [key]: value };
        return newRow;
      }));
    }
  };

  return (
    <div className="flex h-screen bg-gray-50 text-gray-800 font-sans overflow-hidden">
      
      {/* 좌측 사이드바 (툴바) */}
      <div className="w-80 flex flex-col bg-white border-r shadow-sm z-10 overflow-y-auto">
        <div className="p-4 border-b flex items-center justify-between sticky top-0 bg-white z-10">
          <div className="flex items-center gap-3">
            <a href="/" className="flex items-center gap-0.5 hover:opacity-80 transition-opacity" title="홈으로 이동">
              <Plus className="w-5 h-5 text-blue-500 font-bold" strokeWidth={3} />
              <div className="flex -space-x-1">
                <div className="w-4 h-4 rounded-full border-2 border-blue-400"></div>
                <div className="w-4 h-4 rounded-full border-2 border-blue-400 bg-white"></div>
              </div>
            </a>
            <span className="font-bold text-green-700 text-xl flex items-center gap-2">
              <Calculator className="w-6 h-6" /> Crew Studio
            </span>
          </div>
        </div>

        <div className="p-4 flex flex-col gap-6">
          {/* 저장 및 안내 */}
          <div className="flex flex-col gap-2">
            <button 
              onClick={handleSave} 
              className="w-full py-2 px-4 bg-indigo-50 hover:bg-indigo-100 rounded-lg text-sm flex items-center justify-center gap-2 text-indigo-700 font-bold transition-colors border border-indigo-200" 
              title="현재 표를 HTML 코드로 추출하여 복사합니다"
            >
              <Save className="w-5 h-5" /> HTML 코드로 저장
            </button>
            <button 
              onClick={() => setShowDonationSetupModal(true)} 
              className="w-full py-2 px-4 bg-orange-50 hover:bg-orange-100 rounded-lg text-sm flex items-center justify-center gap-2 text-orange-700 font-bold transition-colors border border-orange-200" 
              title="후원페이지 URL 설정 안내"
            >
              <ExternalLink className="w-5 h-5" /> 후원 연동 안내
            </button>
          </div>

          <hr className="border-gray-200" />

          {/* 폰트/스타일/색상 설정 */}
          <div className="flex flex-col gap-3">
            <h3 className="text-xs font-bold text-gray-500 uppercase tracking-wider">스타일 설정</h3>
            
            <div className={`flex flex-col gap-2 ${!isStylePickerEnabled ? 'opacity-50' : ''}`}>
              <select
                value={currentFontFamily}
                onChange={(e) => handleStyleChange('fontFamily', e.target.value)}
                disabled={!isStylePickerEnabled}
                className={`w-full text-sm border border-gray-300 rounded-md p-2 outline-none ${isStylePickerEnabled ? 'cursor-pointer focus:border-blue-500 focus:ring-1 focus:ring-blue-500' : 'pointer-events-none bg-gray-50'}`}
              >
                <option value="sans-serif">고딕 (기본)</option>
                <option value="'Malgun Gothic', sans-serif">맑은 고딕</option>
                <option value="'Nanum Gothic', sans-serif">나눔 고딕</option>
                <option value="'Dotum', sans-serif">돋움</option>
                <option value="'Gulim', sans-serif">굴림</option>
                <option value="serif">명조</option>
                <option value="'Batang', serif">바탕</option>
                <option value="monospace">고정폭</option>
              </select>

              <div className="flex items-center gap-2">
                <select
                  value={currentFontSize}
                  onChange={(e) => handleStyleChange('fontSize', parseInt(e.target.value, 10))}
                  disabled={!isStylePickerEnabled}
                  className={`flex-1 text-sm border border-gray-300 rounded-md p-2 outline-none ${isStylePickerEnabled ? 'cursor-pointer focus:border-blue-500 focus:ring-1 focus:ring-blue-500' : 'pointer-events-none bg-gray-50'}`}
                >
                  {[10, 11, 12, 14, 16, 18, 20, 24, 28, 32].map(size => (
                    <option key={size} value={size}>{size}px</option>
                  ))}
                </select>

                <div className="flex items-center gap-1 border border-gray-300 rounded-md p-1 bg-white">
                  <button
                    onClick={() => handleStyleChange('fontWeight', currentFontWeight === 'bold' ? 'normal' : 'bold')}
                    disabled={!isStylePickerEnabled}
                    className={`p-1.5 rounded transition-colors ${currentFontWeight === 'bold' ? 'bg-gray-200 text-gray-900' : 'hover:bg-gray-100 text-gray-700'} ${!isStylePickerEnabled ? 'cursor-not-allowed' : ''}`}
                    title="굵게 (Bold)"
                  >
                    <Bold className="w-4 h-4" />
                  </button>
                  <button
                    onClick={() => handleStyleChange('textDecoration', currentTextDecoration === 'underline' ? 'none' : 'underline')}
                    disabled={!isStylePickerEnabled}
                    className={`p-1.5 rounded transition-colors ${currentTextDecoration === 'underline' ? 'bg-gray-200 text-gray-900' : 'hover:bg-gray-100 text-gray-700'} ${!isStylePickerEnabled ? 'cursor-not-allowed' : ''}`}
                    title="밑줄 (Underline)"
                  >
                    <Underline className="w-4 h-4" />
                  </button>
                </div>

                <div className="border border-gray-300 rounded-md p-1 bg-white flex items-center justify-center w-10 h-10 relative">
                  <PaintBucket className="w-4 h-4 text-gray-600 absolute pointer-events-none" />
                  <input 
                    type="color" 
                    className={`w-full h-full opacity-0 ${isStylePickerEnabled ? 'cursor-pointer' : 'pointer-events-none'}`}
                    value={currentColor}
                    onChange={(e) => handleStyleChange('color', e.target.value)}
                    disabled={!isStylePickerEnabled}
                    title="배경색 변경"
                  />
                  <div className="absolute bottom-1 right-1 w-3 h-3 rounded-full border border-gray-300 pointer-events-none" style={{ backgroundColor: currentColor }}></div>
                </div>
              </div>
            </div>
          </div>

          <hr className="border-gray-200" />

          {/* 고급 서식 */}
          <div className="flex flex-col gap-3">
            <h3 className="text-xs font-bold text-gray-500 uppercase tracking-wider">고급 서식</h3>
            <div className="grid grid-cols-2 gap-2">
              <button 
                onClick={() => setShowCondFormatModal(true)} 
                className="py-2 px-3 bg-white border border-gray-300 hover:border-pink-400 hover:bg-pink-50 rounded-lg text-sm flex flex-col items-center justify-center gap-1 text-gray-700 transition-colors" 
              >
                <Highlighter className="w-5 h-5 text-pink-500" />
                <span>조건부 서식</span>
              </button>
              <button 
                onClick={() => setShowTableFormatModal(true)} 
                className="py-2 px-3 bg-white border border-gray-300 hover:border-teal-400 hover:bg-teal-50 rounded-lg text-sm flex flex-col items-center justify-center gap-1 text-gray-700 transition-colors" 
              >
                <Table className="w-5 h-5 text-teal-500" />
                <span>표 서식</span>
              </button>
            </div>
            
            <div className="mt-2">
              <button 
                onClick={() => fileInputRef.current?.click()} 
                className="w-full py-2 px-4 bg-white border border-gray-300 hover:border-emerald-400 hover:bg-emerald-50 rounded-lg text-sm flex items-center justify-center gap-2 text-gray-700 transition-colors" 
              >
                <ImageIcon className="w-5 h-5 text-emerald-500" /> 배경 이미지 설정
              </button>
              {bgImage && (
                <div className="mt-2 p-3 bg-gray-50 rounded-lg border border-gray-200 flex flex-col gap-2">
                  <div className="flex justify-between items-center">
                    <span className="text-xs font-medium text-gray-600">투명도 조절</span>
                    <button 
                      onClick={() => {
                        setBgImage(null);
                        setBgOpacity(0.5);
                      }} 
                      className="text-xs text-red-500 hover:text-red-700 font-medium flex items-center gap-1"
                    >
                      <X className="w-3 h-3" /> 제거
                    </button>
                  </div>
                  <input 
                    type="range" 
                    min="0.1" 
                    max="1.0" 
                    step="0.05" 
                    value={bgOpacity} 
                    onChange={(e) => setBgOpacity(parseFloat(e.target.value))}
                    className="w-full accent-emerald-500 cursor-pointer"
                  />
                </div>
              )}
              <input 
                type="file" 
                accept="image/*" 
                ref={fileInputRef} 
                onChange={handleImageUpload} 
                className="hidden" 
              />
            </div>
          </div>

          <hr className="border-gray-200" />

          {/* 행/열 제어 */}
          <div className="flex flex-col gap-3">
            <h3 className="text-xs font-bold text-gray-500 uppercase tracking-wider">행/열 관리</h3>
            <div className="grid grid-cols-2 gap-2">
              <button onClick={() => addRow('below')} className="py-2 px-3 bg-white border border-gray-300 hover:bg-gray-50 rounded-lg text-sm flex items-center justify-center gap-1 text-gray-700 transition-colors">
                <ArrowDown className="w-4 h-4 text-blue-500" /> 행 추가
              </button>
              <button 
                onClick={deleteRow} 
                disabled={selectedRowIndex === null}
                className={`py-2 px-3 bg-white border border-gray-300 rounded-lg text-sm flex items-center justify-center gap-1 transition-colors ${selectedRowIndex !== null ? 'hover:bg-red-50 hover:border-red-300 text-red-600' : 'text-gray-400 cursor-not-allowed bg-gray-50'}`} 
              >
                <Trash2 className="w-4 h-4" /> 행 삭제
              </button>
              <button onClick={() => addCol('right')} className="py-2 px-3 bg-white border border-gray-300 hover:bg-gray-50 rounded-lg text-sm flex items-center justify-center gap-1 text-gray-700 transition-colors">
                <ArrowRight className="w-4 h-4 text-blue-500" /> 열 추가
              </button>
              <button 
                onClick={deleteCol} 
                disabled={selectedColIndex === null}
                className={`py-2 px-3 bg-white border border-gray-300 rounded-lg text-sm flex items-center justify-center gap-1 transition-colors ${selectedColIndex !== null ? 'hover:bg-red-50 hover:border-red-300 text-red-600' : 'text-gray-400 cursor-not-allowed bg-gray-50'}`} 
              >
                <Trash2 className="w-4 h-4" /> 열 삭제
              </button>
            </div>
          </div>

          <hr className="border-gray-200" />

          {/* 데이터 도구 */}
          <div className="flex flex-col gap-3">
            <h3 className="text-xs font-bold text-gray-500 uppercase tracking-wider">데이터 도구</h3>
            <button 
              onClick={sortBySponsorship} 
              disabled={!Object.values(colTitles).includes("후원점수")}
              className={`w-full py-2 px-4 border rounded-lg text-sm flex items-center justify-center gap-2 transition-colors ${Object.values(colTitles).includes("후원점수") ? 'bg-white border-gray-300 hover:border-orange-400 hover:bg-orange-50 text-gray-700' : 'bg-gray-50 border-gray-200 text-gray-400 cursor-not-allowed'}`} 
            >
              <BarChart className={`w-5 h-5 ${Object.values(colTitles).includes("후원점수") ? 'text-orange-500' : ''}`} /> 후원점수 내림차순 정렬
            </button>
            <div className="grid grid-cols-2 gap-2">
              <button onClick={() => applyAutoFunction('SUM')} className="py-2 px-3 bg-white border border-gray-300 hover:border-purple-400 hover:bg-purple-50 rounded-lg text-sm flex items-center justify-center gap-1 text-gray-700 transition-colors">
                <Sigma className="w-4 h-4 text-purple-500" /> 열 합계
              </button>
              <button onClick={() => applyAutoFunction('AVG')} className="py-2 px-3 bg-white border border-gray-300 hover:border-purple-400 hover:bg-purple-50 rounded-lg text-sm flex items-center justify-center gap-1 text-gray-700 transition-colors">
                <Sigma className="w-4 h-4 text-purple-500" /> 열 평균
              </button>
            </div>
          </div>
        </div>
      </div>

      {/* 우측 메인 영역 */}
      <div className="flex-1 flex flex-col min-w-0 bg-gray-100 relative">
        
        {/* 상단 컨트롤 바 (수식 입력줄, VS 모드 등) */}
        <div className="bg-white border-b shadow-sm z-10">
          {/* 시트 탭 영역 */}
          <div className="flex items-center bg-gray-50 border-b border-gray-200 overflow-x-auto select-none px-2 pt-2">
            {sheets.map(sheet => (
              <div
                key={sheet.id}
                onClick={() => {
                  if (editingSheetId) saveSheetName(editingSheetId);
                  setActiveSheetId(sheet.id);
                  setActiveCell({ r: 0, c: 0 });
                  setSelectedRowIndex(null);
                  setSelectedColIndex(null);
                  setEditingCell(null);
                }}
                onDoubleClick={() => {
                  setEditingSheetId(sheet.id);
                  setEditSheetName(sheet.name);
                }}
                className={`group flex items-center gap-2 px-4 py-2 min-w-[100px] max-w-[180px] rounded-t-lg cursor-pointer transition-colors
                  ${activeSheetId === sheet.id ? 'bg-white text-blue-700 font-bold border-t border-x border-gray-200 shadow-[0_2px_0_0_white]' : 'text-gray-600 hover:bg-gray-200 border-t border-x border-transparent'}
                `}
                style={{ marginBottom: '-1px' }}
              >
                {editingSheetId === sheet.id ? (
                  <input
                    autoFocus
                    value={editSheetName}
                    onChange={(e) => setEditSheetName(e.target.value)}
                    onBlur={() => saveSheetName(sheet.id)}
                    onKeyDown={(e) => { if (e.key === 'Enter') saveSheetName(sheet.id); }}
                    className="outline-none border border-blue-400 rounded px-1 w-full text-sm font-normal text-gray-800"
                  />
                ) : (
                  <span className="text-sm truncate flex-1">{sheet.name}</span>
                )}
                
                {sheets.length > 1 && (
                  <button
                    onClick={(e) => deleteSheet(sheet.id, e)}
                    className={`p-1 rounded-full hover:bg-red-100 text-gray-400 hover:text-red-600 transition-opacity ml-auto flex-shrink-0
                      ${activeSheetId === sheet.id ? 'opacity-100' : 'opacity-0 group-hover:opacity-100'}
                    `}
                    title="시트 삭제"
                  >
                    <X className="w-3 h-3" />
                  </button>
                )}
              </div>
            ))}
            <button 
              onClick={addSheet}
              className="p-2 ml-1 mb-1 hover:bg-gray-200 text-gray-600 rounded-md flex-shrink-0 transition-colors"
              title="새 시트 추가"
            >
              <Plus className="w-4 h-4" />
            </button>
          </div>

          <div className="flex items-center justify-between p-3">
            {/* 수식 입력줄 */}
            <div className="flex-1 flex items-center gap-3 max-w-3xl">
              <div className="bg-gray-100 px-3 py-1.5 font-mono text-sm border border-gray-300 rounded-md min-w-[60px] text-center font-bold text-gray-700">
                {numToLetter(activeCell.c)}{activeCell.r + 1}
              </div>
              <div className="flex-1 flex items-center gap-2 px-3 py-1.5 border border-gray-300 rounded-md bg-white focus-within:border-blue-500 focus-within:ring-1 focus-within:ring-blue-500 transition-all">
                <span className="text-gray-400 italic font-serif font-bold">fx</span>
                <input
                  className="flex-1 outline-none bg-transparent text-sm"
                  value={grid[activeCell.r]?.[activeCell.c]?.value || ''}
                  onChange={(e) => updateActiveCell({ value: e.target.value })}
                  placeholder="값이나 수식을 입력하세요 (예: =SUM(A1:A5) 또는 =A1+B1)"
                />
              </div>
            </div>

            <div className="flex items-center gap-4 ml-4">
              <button 
                onClick={openDataInputModal}
                className="bg-[#3b5998] hover:bg-[#2d4373] text-white px-4 py-2 rounded-lg text-sm font-bold shadow-sm transition-colors"
              >
                데이터 입력관리
              </button>
              
              <div className="w-px h-6 bg-gray-300"></div>

              {/* VS 팀 모드 토글 */}
              <label className="flex items-center cursor-pointer group bg-gray-50 hover:bg-gray-100 px-3 py-1.5 rounded-lg border border-gray-200 transition-colors" title="두 팀으로 나누어 점수를 경쟁하는 미리보기 화면을 활성화합니다.">
                <div className="relative">
                  <input 
                    type="checkbox" 
                    className="sr-only" 
                    checked={isVsMode} 
                    onChange={(e) => setIsVsMode(e.target.checked)} 
                  />
                  <div className={`block w-10 h-6 rounded-full transition-colors ${isVsMode ? 'bg-indigo-500' : 'bg-gray-300'}`}></div>
                  <div className={`absolute left-1 top-1 bg-white w-4 h-4 rounded-full transition-transform ${isVsMode ? 'transform translate-x-4' : ''} shadow-sm`}></div>
                </div>
                <span className="ml-3 text-sm font-bold text-gray-700 group-hover:text-indigo-600 transition-colors">VS 팀 모드</span>
              </label>
            </div>
          </div>
        </div>

        {/* 메인 콘텐츠 영역 (그리드 + 미리보기 좌우 배치) */}
        <div className="flex-1 flex overflow-hidden p-4 gap-4">
          
          {/* 좌측: 엑셀 그리드 */}
          <div className="flex-1 bg-white rounded-xl shadow-sm border border-gray-200 overflow-auto flex flex-col relative">
            <div className="p-3 border-b border-gray-100 bg-gray-50/50 flex justify-between items-center sticky top-0 left-0 z-30">
              <h2 className="text-sm font-bold text-gray-700 flex items-center gap-2">
                <Table className="w-4 h-4 text-gray-500" /> 데이터 설정
              </h2>
              <span className="text-xs text-gray-400">더블 클릭하여 편집 / A열 클릭 시 크루 선택</span>
            </div>
            
            <div className="flex-1 overflow-auto p-4">
              <div className="min-w-max inline-block rounded-lg border border-gray-300 overflow-hidden shadow-sm">
                <table 
                  className="border-collapse bg-white" 
                  style={{ 
                    tableLayout: 'fixed',
                    backgroundColor: bgImage ? 'transparent' : 'white',
                    ...(bgImage && {
                      backgroundImage: `url(${bgImage})`,
                      backgroundSize: 'cover',
                      backgroundPosition: 'center'
                    })
                  }}
                >
            {/* 열 헤더 (A, B, C...) */}
            <thead className="sticky top-0 z-20">
              <tr>
                <th className={`w-10 h-12 border border-gray-300 ${bgImage ? 'bg-gray-100/90 backdrop-blur-sm' : 'bg-gray-100'}`}></th>
                {Array.from({ length: colsCount }).map((_, c) => (
                  <th 
                    key={`col-${c}`} 
                    className={`w-28 border border-gray-300 text-center font-normal text-sm select-none p-1 cursor-pointer transition-colors
                      ${selectedColIndex === c 
                        ? (bgImage ? 'bg-blue-200/90 text-blue-800' : 'bg-blue-200 text-blue-800') 
                        : (bgImage ? 'bg-gray-100/90 hover:bg-gray-200/90 text-gray-800' : 'bg-gray-100 text-gray-600 hover:bg-gray-200')}
                      ${bgImage ? 'backdrop-blur-sm' : ''}
                    `}
                    title="열 전체 선택"
                    onClick={() => {
                      setSelectedColIndex(c);
                      setActiveCell({ r: 0, c }); // 열 선택 시 기준 셀을 해당 열의 첫 번째로 이동
                      setEditingCell(null);
                      setSelectedRowIndex(null); // 행 선택 초기화
                    }}
                  >
                    <div className={`text-[10px] leading-none mb-1 ${selectedColIndex === c ? 'text-blue-700 font-bold' : 'text-gray-500'}`}>
                      {numToLetter(c)}
                    </div>
                    {c === 0 ? (
                      <div className="font-semibold text-xs">크루</div>
                    ) : (c >= 1 && c <= 4) ? (
                      editingColTitle === c ? (
                        <select
                          autoFocus
                          value={colTitles[c] || "선택"}
                          onClick={(e) => e.stopPropagation()} // 셀렉트 박스 클릭 시 열 전체 선택 이벤트 전파 방지
                          onBlur={() => setEditingColTitle(null)} // 포커스를 잃으면 다시 텍스트 모드로 전환
                          onChange={(e) => {
                            setColTitles(prev => ({ ...prev, [c]: e.target.value }));
                            setEditingColTitle(null); // 선택을 완료하면 텍스트 모드로 전환
                          }}
                          className="w-[90%] mx-auto block bg-white border border-blue-400 rounded px-1 py-0.5 text-[11px] outline-none cursor-pointer text-center text-gray-800 font-normal shadow-sm"
                        >
                          {headerOptions.map(opt => {
                            // 현재 열에서 선택한 값이 아니면서, 이미 다른 열에서 선택된 값인지 확인
                            const isAlreadySelected = Object.values(colTitles).includes(opt) && colTitles[c] !== opt;
                            const isDisabled = opt === "선택" || isAlreadySelected;
                            return (
                              <option 
                                key={opt} 
                                value={opt} 
                                disabled={isDisabled}
                                className={isDisabled ? "text-gray-300" : "text-gray-800"}
                              >
                                {opt}
                              </option>
                            );
                          })}
                        </select>
                      ) : (
                        <div
                          onClick={(e) => {
                            e.stopPropagation(); // 열 전체 선택 방지
                            setEditingColTitle(c); // 셀렉트박스(편집) 모드로 전환
                          }}
                          className="w-[90%] mx-auto truncate border border-transparent hover:border-gray-300 hover:bg-white/80 rounded px-1 py-0.5 text-[11px] cursor-pointer text-center font-medium transition-colors"
                          title="클릭하여 헤더 속성 변경"
                        >
                          {colTitles[c] || "선택"}
                        </div>
                      )
                    ) : (
                      <div className="h-5"></div>
                    )}
                  </th>
                ))}
              </tr>
            </thead>
            
            {/* 본문 셀 */}
            <tbody>
              {displayGrid.map((row, r) => (
                <tr key={`row-${r}`}>
                  {/* 행 헤더 (1, 2, 3...) */}
                  <td 
                    className={`w-10 border border-gray-300 text-center text-sm select-none sticky left-0 z-10 cursor-pointer transition-colors
                      ${selectedRowIndex === r 
                        ? (bgImage ? 'bg-blue-200/90 text-blue-800 font-bold' : 'bg-blue-200 text-blue-800 font-bold') 
                        : (bgImage ? 'bg-gray-100/90 text-gray-800 hover:bg-gray-200/90' : 'bg-gray-100 text-gray-600 hover:bg-gray-200')}
                      ${bgImage ? 'backdrop-blur-sm' : ''}
                    `}
                    title="행 전체 선택"
                    onClick={() => {
                      setSelectedRowIndex(r);
                      setActiveCell({ r, c: 0 }); // 행 선택 시 기준 셀을 해당 행의 첫 번째로 이동
                      setEditingCell(null);
                      setSelectedColIndex(null); // 열 선택 초기화
                    }}
                  >
                    {r + 1}
                  </td>
                  
                  {/* 데이터 셀 */}
                  {row.map((cell, c) => {
                    const isActive = activeCell.r === r && activeCell.c === c;
                    const isEditing = editingCell?.r === r && editingCell?.c === c;
                    const isRowSelected = selectedRowIndex === r;
                    const isColSelected = selectedColIndex === c;
                    
                    // 배경 이미지가 있고, 셀 배경색이 기본값(#ffffff)인지 확인
                    const isDefaultColor = cell.color === '#ffffff' || !cell.color;
                    
                    // 조건부 서식 평가
                    let finalBgColor = cell.color;
                    let finalColor = cell.textColor !== 'inherit' && cell.textColor ? cell.textColor : (String(cell.computed).startsWith('#') ? 'red' : 'inherit');
                    let isCondFormatApplied = false;
                    
                    const rules = currentSheet.conditionalRules || [];
                    rules.filter(rule => rule.targetCol === c).forEach(rule => {
                      if (checkCondition(cell.computed, rule)) {
                        const formatStyle = condFormatOptions.find(opt => opt.id === rule.format);
                        if (formatStyle) {
                          if (formatStyle.bg !== 'transparent') {
                            finalBgColor = formatStyle.bg;
                            isCondFormatApplied = true;
                          }
                          if (formatStyle.color !== 'inherit') finalColor = formatStyle.color;
                        }
                      }
                    });

                    // 조건부 서식이 적용되지 않은 기본 셀에 배경 이미지가 있는 경우 반투명 처리
                    if (bgImage && isDefaultColor && !isCondFormatApplied) {
                      finalBgColor = `rgba(255, 255, 255, ${1 - bgOpacity})`;
                    }
                    
                    return (
                      <td
                        key={`cell-${r}-${c}`}
                        className={`border border-gray-300 relative cursor-cell
                          ${isActive ? 'outline outline-2 outline-blue-500 z-10' : ''}
                          ${(isRowSelected || isColSelected) && !isActive ? 'outline outline-1 outline-blue-400' : ''}
                          ${(isRowSelected || isColSelected) && !isActive && (isDefaultColor) && !isCondFormatApplied ? 'bg-blue-50/40' : ''}
                        `}
                        style={{ 
                          backgroundColor: finalBgColor, 
                          height: '32px' 
                        }}
                        onClick={() => {
                          if (!isEditing) {
                            setActiveCell({ r, c });
                            setEditingCell(null);
                            setSelectedRowIndex(null);
                            setSelectedColIndex(null);
                            
                            if (c === 0) {
                              setShowCrewModal(true);
                            }
                          }
                        }}
                        onDoubleClick={() => {
                          setActiveCell({ r, c });
                          setSelectedRowIndex(null);
                          setSelectedColIndex(null);
                          
                          if (c !== 0) {
                            setEditingCell({ r, c });
                          } else {
                            setShowCrewModal(true);
                          }
                        }}
                      >
                        {isEditing ? (
                          <input
                            ref={editInputRef}
                            className="absolute inset-0 w-full h-full px-1.5 outline-none shadow-inner"
                            style={{ 
                              backgroundColor: cell.color, // 편집 시에는 조건부서식/투명도 제거된 원본색 노출
                              fontFamily: cell.fontFamily || 'sans-serif',
                              fontSize: `${cell.fontSize || 14}px`,
                              fontWeight: cell.fontWeight || 'normal',
                              textDecoration: cell.textDecoration || 'none',
                              color: cell.textColor || 'inherit'
                            }} 
                            value={grid[r][c].value}
                            onChange={(e) => handleCellChange(r, c, e.target.value)}
                            onBlur={() => setEditingCell(null)}
                            onKeyDown={(e) => {
                              if (e.key === 'Enter') {
                                setEditingCell(null);
                                setActiveCell(prev => ({ ...prev, r: Math.min(prev.r + 1, rowsCount - 1) }));
                              }
                            }}
                          />
                        ) : (
                          <div 
                            className="w-full h-full px-1.5 flex items-center overflow-hidden whitespace-nowrap text-ellipsis cursor-cell gap-1.5"
                            style={{ 
                              color: finalColor, 
                              justifyContent: c === 0 || (isNaN(cell.computed) && cell.computed !== '' && !String(cell.computed).startsWith('#')) ? 'flex-start' : 'flex-end',
                              fontWeight: cell.fontWeight === 'bold' ? 'bold' : (bgImage && isDefaultColor ? '500' : 'normal'),
                              fontFamily: cell.fontFamily || 'sans-serif',
                              fontSize: `${cell.fontSize || 14}px`,
                              textDecoration: cell.textDecoration || 'none'
                            }}
                          >
                            {cell.avatar && (
                              <img src={cell.avatar} alt="crew" className="w-5 h-5 rounded-full object-cover flex-shrink-0 border border-gray-200" />
                            )}
                            {c === 0 && !cell.computed ? (
                              <span className="truncate text-gray-400 text-xs tracking-tight">크루원 선택</span>
                            ) : (
                              <span className="truncate">{cell.computed}</span>
                            )}
                          </div>
                        )}
                      </td>
                    );
                  })}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
      </div>

      {/* 우측: 미리보기 (Preview Table or VS Board) */}
      <div className="w-[45%] bg-white rounded-xl shadow-sm border border-gray-200 overflow-auto flex flex-col relative">
        <div className="p-3 border-b border-gray-100 bg-gray-50/50 flex justify-between items-center sticky top-0 left-0 z-30">
          <h2 className="text-sm font-bold text-gray-700 flex items-center gap-2">
            미리보기 {isVsMode && <span className="text-indigo-600 text-[10px] bg-indigo-50 px-1.5 py-0.5 rounded font-semibold border border-indigo-100 ml-1">VS 팀 모드</span>}
          </h2>
        </div>
        
        <div className="flex-1 overflow-auto p-4">
          {isVsMode ? (() => {
             const [team1, team2] = getVsTeamsData();
             const totalScore = team1.score + team2.score;
             
             // 점수 비율에 따른 너비 계산 (줄다리기 효과)
             // 한 쪽이 아예 안보이는 것을 방지하기 위해 최소 15%, 최대 85% 범위 지정
             let t1Ratio = 50;
             if (totalScore > 0) {
                t1Ratio = (team1.score / totalScore) * 100;
                t1Ratio = Math.max(15, Math.min(85, t1Ratio));
             }
             const t2Ratio = 100 - t1Ratio;

             return (
                <div 
                  ref={tableRef}
                  style={{ display: 'flex', minHeight: '350px', borderRadius: '16px', overflow: 'hidden', boxShadow: '0 10px 30px rgba(0,0,0,0.15)', color: 'white', position: 'relative', fontFamily: 'sans-serif' }}
                >
                  {/* VS Badge */}
                  <div style={{ 
                    position: 'absolute', 
                    left: `${t1Ratio}%`, 
                    top: '50%', 
                    transform: 'translate(-50%, -50%)', 
                    backgroundColor: '#facc15', // yellow
                    color: '#111827', 
                    width: '60px', 
                    height: '60px', 
                    display: 'flex', 
                    alignItems: 'center', 
                    justifyContent: 'center', 
                    borderRadius: '50%', 
                    fontWeight: '900', 
                    fontStyle: 'italic', 
                    fontSize: '20px', 
                    zIndex: 10, 
                    border: '6px solid #111827', 
                    boxShadow: '0 4px 10px rgba(0,0,0,0.4)',
                    transition: 'left 0.4s cubic-bezier(0.4, 0, 0.2, 1)'
                  }}>
                    VS
                  </div>

                  {/* Team 1 (Left - Blue) */}
                  <div style={{ 
                    width: `${t1Ratio}%`, 
                    backgroundColor: '#1d4ed8', // 짙은 파랑
                    padding: '24px', 
                    display: 'flex', 
                    flexDirection: 'column',
                    transition: 'width 0.4s cubic-bezier(0.4, 0, 0.2, 1)'
                  }}>
                    <h3 style={{ fontSize: '18px', fontWeight: 'bold', color: '#dbeafe', marginBottom: '8px', margin: 0 }}>{team1.name}</h3>
                    <div style={{ fontSize: '48px', fontWeight: '900', marginBottom: '24px', lineHeight: 1 }}>
                      {team1.score.toLocaleString()} <span style={{ fontSize: '16px', fontWeight: 'normal', color: '#93c5fd' }}>pts</span>
                    </div>
                    <div style={{ display: 'flex', flexWrap: 'wrap', gap: '8px' }}>
                      {team1.members.map((m, i) => (
                        <div key={i} style={{ display: 'flex', alignItems: 'center', gap: '8px', backgroundColor: 'rgba(0,0,0,0.3)', padding: '4px 12px 4px 4px', borderRadius: '9999px' }}>
                          {m.avatar ? <img src={m.avatar} style={{ width: '28px', height: '28px', borderRadius: '50%', objectFit: 'cover' }} alt="avatar" /> : <div style={{ width: '28px', height: '28px', borderRadius: '50%', backgroundColor: '#4b5563' }}></div>}
                          <div style={{ display: 'flex', flexDirection: 'column' }}>
                            <span style={{ fontSize: '12px', fontWeight: 'bold' }}>{m.name}</span>
                            <span style={{ fontSize: '10px', color: '#bfdbfe', fontWeight: '500' }}>{m.score.toLocaleString()}</span>
                          </div>
                        </div>
                      ))}
                    </div>
                  </div>

                  {/* Team 2 (Right - Red) */}
                  <div style={{ 
                    width: `${t2Ratio}%`, 
                    backgroundColor: '#b91c1c', // 짙은 빨강
                    padding: '24px', 
                    display: 'flex', 
                    flexDirection: 'column', 
                    alignItems: 'flex-end', 
                    textAlign: 'right',
                    transition: 'width 0.4s cubic-bezier(0.4, 0, 0.2, 1)'
                  }}>
                    <h3 style={{ fontSize: '18px', fontWeight: 'bold', color: '#fee2e2', marginBottom: '8px', margin: 0 }}>{team2.name}</h3>
                    <div style={{ fontSize: '48px', fontWeight: '900', marginBottom: '24px', lineHeight: 1 }}>
                      {team2.score.toLocaleString()} <span style={{ fontSize: '16px', fontWeight: 'normal', color: '#fca5a5' }}>pts</span>
                    </div>
                    <div style={{ display: 'flex', flexWrap: 'wrap', gap: '8px', justifyContent: 'flex-end' }}>
                      {team2.members.map((m, i) => (
                        <div key={i} style={{ display: 'flex', alignItems: 'center', gap: '8px', backgroundColor: 'rgba(0,0,0,0.3)', padding: '4px 4px 4px 12px', borderRadius: '9999px', flexDirection: 'row-reverse' }}>
                          {m.avatar ? <img src={m.avatar} style={{ width: '28px', height: '28px', borderRadius: '50%', objectFit: 'cover' }} alt="avatar" /> : <div style={{ width: '28px', height: '28px', borderRadius: '50%', backgroundColor: '#4b5563' }}></div>}
                          <div style={{ display: 'flex', flexDirection: 'column', textAlign: 'left' }}>
                            <span style={{ fontSize: '12px', fontWeight: 'bold' }}>{m.name}</span>
                            <span style={{ fontSize: '10px', color: '#fecaca', fontWeight: '500' }}>{m.score.toLocaleString()}</span>
                          </div>
                        </div>
                      ))}
                    </div>
                  </div>
                </div>
             );
          })() : (
            <div className="min-w-max inline-block rounded-lg border border-gray-300 overflow-hidden shadow-sm">
              <table 
                ref={tableRef}
                className="border-collapse bg-white" 
                style={{ 
                  tableLayout: 'fixed',
                  backgroundColor: bgImage ? 'transparent' : 'white',
                  ...(bgImage && {
                    backgroundImage: `url(${bgImage})`,
                    backgroundSize: 'cover',
                    backgroundPosition: 'center'
                  })
                }}
              >
                {/* 미리보기 열 헤더 */}
                <thead>
                  <tr>
                    <th className={`w-10 h-10 border border-gray-300 ${bgImage ? 'bg-gray-100/90 backdrop-blur-sm' : 'bg-gray-100'}`}></th>
                    {Array.from({ length: colsCount }).map((_, c) => (
                      <th 
                        key={`prev-col-${c}`} 
                        className={`w-28 border border-gray-300 text-center font-semibold text-sm select-none p-2
                          ${bgImage ? 'bg-gray-100/90 backdrop-blur-sm text-gray-800' : 'bg-gray-100 text-gray-700'}
                        `}
                      >
                        {c === 0 ? "크루" : (colTitles[c] !== "선택" ? colTitles[c] : "")}
                      </th>
                    ))}
                  </tr>
                </thead>
                
                {/* 미리보기 본문 셀 */}
                <tbody>
                  {displayGrid.map((row, r) => (
                    <tr key={`prev-row-${r}`}>
                      <td 
                        className={`w-10 border border-gray-300 text-center text-sm select-none
                          ${bgImage ? 'bg-gray-100/90 text-gray-800' : 'bg-gray-100 text-gray-600'}
                          ${bgImage ? 'backdrop-blur-sm' : ''}
                        `}
                      >
                        {r + 1}
                      </td>
                      
                      {row.map((cell, c) => {
                        const isDefaultColor = cell.color === '#ffffff' || !cell.color;
                        
                        let finalBgColor = cell.color;
                        let finalColor = cell.textColor !== 'inherit' && cell.textColor ? cell.textColor : (String(cell.computed).startsWith('#') ? 'red' : 'inherit');
                        let isCondFormatApplied = false;
                        
                        const rules = currentSheet.conditionalRules || [];
                        rules.filter(rule => rule.targetCol === c).forEach(rule => {
                          if (checkCondition(cell.computed, rule)) {
                            const formatStyle = condFormatOptions.find(opt => opt.id === rule.format);
                            if (formatStyle) {
                              if (formatStyle.bg !== 'transparent') {
                                finalBgColor = formatStyle.bg;
                                isCondFormatApplied = true;
                              }
                              if (formatStyle.color !== 'inherit') finalColor = formatStyle.color;
                            }
                          }
                        });

                        if (bgImage && isDefaultColor && !isCondFormatApplied) {
                          finalBgColor = `rgba(255, 255, 255, ${1 - bgOpacity})`;
                        }
                        
                        return (
                          <td
                            key={`prev-cell-${r}-${c}`}
                            className="border border-gray-300 relative"
                            style={{ backgroundColor: finalBgColor, height: '32px' }}
                          >
                            <div 
                              className="w-full h-full px-1.5 flex items-center overflow-hidden whitespace-nowrap text-ellipsis gap-1.5"
                              style={{ 
                                color: finalColor, 
                                justifyContent: c === 0 || (isNaN(cell.computed) && cell.computed !== '' && !String(cell.computed).startsWith('#')) ? 'flex-start' : 'flex-end',
                                fontWeight: cell.fontWeight === 'bold' ? 'bold' : (bgImage && isDefaultColor ? '500' : 'normal'),
                                fontFamily: cell.fontFamily || 'sans-serif',
                                fontSize: `${cell.fontSize || 14}px`,
                                textDecoration: cell.textDecoration || 'none'
                              }}
                            >
                              {cell.avatar && (
                                <img src={cell.avatar} alt="crew" className="w-5 h-5 rounded-full object-cover flex-shrink-0 border border-gray-200" />
                              )}
                              {c === 0 && !cell.computed ? (
                                <span className="truncate text-gray-400 text-xs tracking-tight">크루원 선택</span>
                              ) : (
                                <span className="truncate">{cell.computed}</span>
                              )}
                            </div>
                          </td>
                        );
                      })}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}
        </div>
      </div>
      </div>

      {/* 상태 표시줄 */}
      <div className="bg-gray-200 border-t p-1 px-4 text-xs text-gray-500 flex justify-between z-20">
        <span>더블 클릭하여 셀/시트 이름 편집 / <b>A열 클릭 시 크루 선택</b></span>
        <span>지원 함수: SUM(범위), AVG(범위), MAX(범위), MIN(범위) 및 기본 사칙연산(+,-,*,/)</span>
      </div>
    </div>

      {/* 크루 선택 모달 */}
      {showCrewModal && (
        <div className="fixed inset-0 bg-black/40 z-50 flex items-center justify-center p-4" onClick={() => setShowCrewModal(false)}>
          <div className="bg-white p-8 rounded-xl shadow-xl max-w-3xl w-full" onClick={e => e.stopPropagation()}>
            <div className="flex justify-between items-center mb-6">
              <h2 className="text-xl font-bold text-gray-800">크루 선택</h2>
              <button onClick={() => setShowCrewModal(false)} className="text-gray-500 hover:text-gray-800">
                <X className="w-6 h-6" />
              </button>
            </div>
            
            <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-x-6 gap-y-4">
              {/* 추가 버튼 */}
              <div 
                className="flex items-center gap-4 cursor-pointer group p-2 hover:bg-gray-50 rounded-lg transition-colors"
                onClick={handleAddCrew}
              >
                <div className="w-14 h-14 rounded-full border border-gray-300 border-dashed bg-white flex items-center justify-center group-hover:bg-gray-50 transition-colors">
                  <Plus className="text-gray-400 w-6 h-6" />
                </div>
                <span className="text-sm font-medium text-gray-600">추가</span>
              </div>

              {/* 크루 리스트 */}
              {crews.map(crew => (
                <div 
                  key={crew.id} 
                  className="flex items-center gap-4 cursor-pointer group p-2 hover:bg-gray-50 rounded-lg transition-colors"
                  onClick={() => handleCrewSelect(crew)}
                >
                  <img src={crew.avatar} alt={crew.name} className="w-14 h-14 rounded-full object-cover border border-gray-200 shadow-sm" />
                  <span className="text-sm font-medium text-gray-800 flex-1 truncate">{crew.name}</span>
                  <button 
                    className="text-red-400 hover:text-red-600 p-2 opacity-100 sm:opacity-0 sm:group-hover:opacity-100 transition-opacity"
                    onClick={(e) => handleDeleteCrew(e, crew.id)}
                    title="크루 삭제"
                  >
                    <Trash2 className="w-5 h-5" />
                  </button>
                </div>
              ))}
            </div>
          </div>
        </div>
      )}

      {/* 조건부 서식 설정 모달 */}
      {showCondFormatModal && (
        <div className="fixed inset-0 bg-black/40 z-50 flex items-center justify-center p-4" onClick={() => setShowCondFormatModal(false)}>
          <div className="bg-white p-6 rounded-xl shadow-xl max-w-[450px] w-full" onClick={e => e.stopPropagation()}>
            <div className="flex justify-between items-center mb-4">
              <h2 className="text-lg font-bold text-gray-800 flex items-center gap-2">
                <Highlighter className="w-5 h-5 text-pink-600" />
                셀 강조 규칙
              </h2>
              <button onClick={() => setShowCondFormatModal(false)} className="text-gray-500 hover:text-gray-800">
                <X className="w-5 h-5" />
              </button>
            </div>
            
            <div className="bg-blue-50 text-blue-800 text-xs px-3 py-2 rounded mb-4">
              현재 <strong className="font-bold">{numToLetter(selectedColIndex !== null ? selectedColIndex : activeCell.c)}열</strong>에 적용할 규칙을 설정합니다.
            </div>

            <div className="mb-4">
              <label className="block text-sm font-medium text-gray-700 mb-1">다음 규칙을 충족하는 셀의 서식 지정:</label>
              <div className="flex gap-2 mb-4">
                <select 
                  className="border border-gray-300 rounded p-2 text-sm outline-none focus:border-blue-400"
                  value={condRule.type}
                  onChange={e => setCondRule({ ...condRule, type: e.target.value })}
                >
                  <option value="greaterThan">다음 값보다 큼 ( &gt; )</option>
                  <option value="lessThan">다음 값보다 작음 ( &lt; )</option>
                  <option value="equalTo">다음 값과 같음 ( = )</option>
                  <option value="textContains">다음 텍스트 포함</option>
                </select>
                <input 
                  type="text" 
                  className="flex-1 border border-gray-300 rounded p-2 text-sm outline-none focus:border-blue-400" 
                  placeholder="값 입력..."
                  value={condRule.value}
                  onChange={e => setCondRule({ ...condRule, value: e.target.value })}
                />
              </div>

              <div className="flex items-center gap-2">
                <label className="text-sm font-medium text-gray-700 whitespace-nowrap">적용할 서식:</label>
                <select 
                  className="flex-1 border border-gray-300 rounded p-2 text-sm outline-none focus:border-blue-400"
                  value={condRule.format}
                  onChange={e => setCondRule({ ...condRule, format: e.target.value })}
                >
                  {condFormatOptions.map(opt => (
                    <option key={opt.id} value={opt.id}>{opt.label}</option>
                  ))}
                </select>
              </div>
            </div>

            <div className="flex justify-between mt-6 pt-4 border-t border-gray-100">
              <button 
                onClick={handleClearCondRules}
                className="px-3 py-1.5 text-sm text-red-600 hover:bg-red-50 rounded transition-colors"
                title="이 열에 적용된 모든 조건부 서식을 삭제합니다"
              >
                규칙 지우기
              </button>
              <div className="flex gap-2">
                <button 
                  onClick={() => setShowCondFormatModal(false)} 
                  className="px-4 py-1.5 rounded text-sm text-gray-600 hover:bg-gray-100 transition-colors"
                >
                  취소
                </button>
                <button 
                  onClick={handleApplyCondRule}
                  disabled={condRule.value.trim() === ''}
                  className={`px-4 py-1.5 rounded text-sm text-white transition-colors ${condRule.value.trim() === '' ? 'bg-gray-400 cursor-not-allowed' : 'bg-blue-600 hover:bg-blue-700'}`}
                >
                  확인
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* 표 서식 갤러리 모달 */}
      {showTableFormatModal && (
        <div className="fixed inset-0 bg-black/40 z-50 flex items-center justify-center p-4" onClick={() => setShowTableFormatModal(false)}>
          <div className="bg-white p-6 rounded-xl shadow-xl max-w-[550px] w-full max-h-[85vh] flex flex-col" onClick={e => e.stopPropagation()}>
            <div className="flex justify-between items-center mb-4">
              <h2 className="text-lg font-bold text-gray-800 flex items-center gap-2">
                <Table className="w-5 h-5 text-teal-600" />
                표 서식 선택
              </h2>
              <button onClick={() => setShowTableFormatModal(false)} className="text-gray-500 hover:text-gray-800">
                <X className="w-5 h-5" />
              </button>
            </div>
            
            <div className="bg-teal-50 text-teal-800 text-xs px-3 py-2 rounded mb-4">
              선택한 디자인이 <strong>현재 시트의 전체 표</strong>에 즉시 적용됩니다.
            </div>

            <div className="flex-1 overflow-y-auto pr-2 custom-scrollbar">
              {['밝게', '중간', '어둡게'].map(category => (
                <div key={category} className="mb-6">
                  <h3 className="text-sm font-bold text-gray-700 mb-3 border-b pb-1">{category}</h3>
                  <div className="grid grid-cols-5 gap-3">
                    {tableFormats.filter(tf => tf.category === category).map(tf => (
                      <div 
                        key={tf.id} 
                        onClick={() => applyTableFormat(tf)} 
                        className="cursor-pointer border border-gray-200 rounded overflow-hidden flex flex-col hover:ring-2 hover:ring-teal-500 transition-all shadow-sm bg-white h-[42px]"
                        title={`${category} 스타일 적용`}
                      >
                        {/* 썸네일 프리뷰 영역 (4행) */}
                        <div className="flex-1 flex" style={{ backgroundColor: tf.headerBg }}>
                          <div className="flex-1 border-r border-black/10"></div>
                          <div className="flex-1 border-r border-black/10"></div>
                          <div className="flex-1"></div>
                        </div>
                        <div className="flex-1 flex" style={{ backgroundColor: tf.oddRowBg }}>
                          <div className="flex-1 border-r border-black/5"></div>
                          <div className="flex-1 border-r border-black/5"></div>
                          <div className="flex-1"></div>
                        </div>
                        <div className="flex-1 flex" style={{ backgroundColor: tf.evenRowBg }}>
                          <div className="flex-1 border-r border-black/5"></div>
                          <div className="flex-1 border-r border-black/5"></div>
                          <div className="flex-1"></div>
                        </div>
                        <div className="flex-1 flex" style={{ backgroundColor: tf.oddRowBg }}>
                          <div className="flex-1 border-r border-black/5"></div>
                          <div className="flex-1 border-r border-black/5"></div>
                          <div className="flex-1"></div>
                        </div>
                      </div>
                    ))}
                  </div>
                </div>
              ))}
            </div>
            
            <div className="mt-4 pt-4 border-t border-gray-100 text-right">
               <button 
                  onClick={() => setShowTableFormatModal(false)} 
                  className="px-4 py-1.5 rounded text-sm text-gray-600 hover:bg-gray-100 transition-colors"
                >
                  닫기
                </button>
            </div>
          </div>
        </div>
      )}

      {/* 데이터 입력 관리 모달 */}
      {showDataInputModal && (() => {
        const targetCols = getActiveInputCols();
        return (
          <div className="fixed inset-0 z-50 flex items-center justify-center p-4 pointer-events-none">
            {/* 뒷배경 오버레이 */}
            <div className="absolute inset-0 bg-black/40 pointer-events-auto" onClick={() => setShowDataInputModal(false)}></div>
            
            {/* 드래그 가능한 모달창 */}
            <div 
              className="bg-white rounded-xl shadow-xl w-full max-w-[500px] flex flex-col h-auto max-h-[85vh] overflow-hidden relative pointer-events-auto" 
              onClick={e => e.stopPropagation()}
              style={{ transform: `translate(${dataModalPos.x}px, ${dataModalPos.y}px)` }}
            >
              {/* 드래그 핸들 영역 (헤더) */}
              <div 
                className="pt-6 px-6 pb-2 cursor-grab active:cursor-grabbing select-none"
                onPointerDown={handleDragStart}
                onPointerMove={handleDragMove}
                onPointerUp={handleDragEnd}
                onPointerCancel={handleDragEnd}
              >
                <button 
                  onClick={() => setShowDataInputModal(false)} 
                  className="absolute top-4 right-4 text-gray-400 hover:text-gray-600 cursor-pointer"
                  onPointerDown={e => e.stopPropagation()} // 드래그 이벤트 전파 방지
                >
                  <X className="w-6 h-6" />
                </button>
                <h2 className="text-xl font-extrabold text-gray-900 mt-2 mb-4 tracking-tight pointer-events-none">데이터 입력 관리</h2>
              </div>
              
              <div className="px-6 pb-4">
                <div className="overflow-auto custom-scrollbar max-h-[50vh] border border-black">
                  <table className="w-full border-collapse text-sm text-left">
                    <thead>
                      <tr>
                        <th className="border border-black bg-white p-2 font-bold text-gray-900 w-[30%]">크루</th>
                        {targetCols.map(col => (
                          <th key={`header-${col.index}`} className="border border-black bg-white p-2 font-bold text-gray-900">
                            {col.label}
                          </th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {tempGridData.map((rowVals, r) => (
                        <tr key={r}>
                          <td className="border border-black bg-white p-2 text-gray-800 font-medium truncate">
                            {grid[r][0].value || ''}
                          </td>
                          {targetCols.map(col => (
                            <td key={`cell-${r}-${col.index}`} className="border border-black p-0 relative bg-white h-[38px] min-w-[80px]">
                              <input
                                type="text"
                                className="w-full h-full p-2 outline-none text-left focus:bg-blue-50 focus:ring-1 focus:ring-blue-500 absolute inset-0"
                                value={rowVals[col.index] || ''}
                                onChange={(e) => handleTempDataChange(r, col.index, e.target.value)}
                              />
                            </td>
                          ))}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>

              <div className="p-6 pt-2 pb-6 bg-white">
                <button 
                  onClick={saveDataInput} 
                  className="w-full bg-[#0e73f6] hover:bg-[#0b5bc4] text-white py-3 rounded-xl text-base font-bold transition-colors shadow-sm"
                >
                  저장
                </button>
              </div>
            </div>
          </div>
        );
      })()}

      {/* 후원페이지 설정 안내 모달 */}
      {showDonationSetupModal && (
        <div className="fixed inset-0 bg-black/40 z-50 flex items-center justify-center p-4" onClick={() => setShowDonationSetupModal(false)}>
          <div className="bg-white p-6 rounded-xl shadow-xl max-w-md w-full flex flex-col" onClick={e => e.stopPropagation()}>
            <div className="flex justify-between items-center mb-4">
              <h2 className="text-xl font-bold text-gray-800 flex items-center gap-2">
                <ExternalLink className="w-6 h-6 text-orange-600" />
                후원페이지 URL 설정 안내
              </h2>
              <button onClick={() => setShowDonationSetupModal(false)} className="text-gray-500 hover:text-gray-800">
                <X className="w-6 h-6" />
              </button>
            </div>
            
            <div className="text-sm text-gray-700 space-y-3 mb-6 bg-orange-50 p-4 rounded-lg border border-orange-100">
              <p className="font-medium text-orange-800">방송에 후원페이지를 연동하려면 아래 단계를 따라 해보세요. 🎉</p>
              <ol className="list-decimal pl-5 space-y-2 mt-2">
                <li>왼쪽 메뉴에서 <strong>[계정설정]</strong>을 누르고 <strong>[내 후원페이지 URL 설정]</strong>을 클릭하세요.</li>
                <li>원하는 URL 주소를 입력하고 중복 확인을 진행하세요.</li>
                <li>설정이 완료된 URL을 복사하여 시청자들에게 공유하세요!</li>
              </ol>
            </div>

            <div className="flex justify-end gap-2">
              <button 
                onClick={() => setShowDonationSetupModal(false)} 
                className="px-4 py-2 rounded-lg text-sm font-medium text-gray-600 hover:bg-gray-100 transition-colors"
              >
                닫기
              </button>
              <button 
                onClick={() => {
                  alert("'계정설정 > 내 후원페이지 URL 설정' 메뉴로 이동합니다.");
                  setShowDonationSetupModal(false);
                }} 
                className="flex items-center gap-2 px-4 py-2 rounded-lg text-sm font-medium text-white bg-orange-500 hover:bg-orange-600 transition-colors"
              >
                다이렉트 이동 <ArrowRight className="w-4 h-4" />
              </button>
            </div>
          </div>
        </div>
      )}

      {/* HTML 내보내기 모달 */}
      {showExportModal && (
        <div className="fixed inset-0 bg-black/40 z-50 flex items-center justify-center p-4" onClick={() => setShowExportModal(false)}>
          <div className="bg-white p-6 rounded-xl shadow-xl max-w-2xl w-full flex flex-col max-h-[80vh]" onClick={e => e.stopPropagation()}>
            <div className="flex justify-between items-center mb-4">
              <h2 className="text-xl font-bold text-gray-800 flex items-center gap-2">
                <Save className="w-6 h-6 text-indigo-600" />
                HTML 코드 저장
              </h2>
              <button onClick={() => setShowExportModal(false)} className="text-gray-500 hover:text-gray-800">
                <X className="w-6 h-6" />
              </button>
            </div>
            
            <p className="text-sm text-gray-600 mb-4">
              아래 생성된 HTML 코드를 복사하여 웹사이트나 게시판 등에 붙여넣을 수 있습니다.
            </p>

            <div className="flex-1 overflow-hidden relative border border-gray-200 rounded-md bg-gray-50 mb-4 flex flex-col">
              <textarea 
                className="w-full h-full p-4 text-xs font-mono text-gray-700 bg-transparent outline-none resize-none"
                value={exportHtml}
                readOnly
              />
            </div>

            <div className="flex justify-end gap-2">
              <button 
                onClick={() => setShowExportModal(false)} 
                className="px-4 py-2 rounded-lg text-sm font-medium text-gray-600 hover:bg-gray-100 transition-colors"
              >
                닫기
              </button>
              <button 
                onClick={handleCopyHtml} 
                className={`flex items-center gap-2 px-4 py-2 rounded-lg text-sm font-medium text-white transition-colors ${copySuccess ? 'bg-green-600 hover:bg-green-700' : 'bg-blue-600 hover:bg-blue-700'}`}
              >
                {copySuccess ? <Check className="w-4 h-4" /> : <Copy className="w-4 h-4" />}
                {copySuccess ? '복사 완료!' : '코드 복사하기'}
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}