import React, { useState, useEffect, useRef, useMemo, useCallback } from 'react';
import { mepApi } from '../mepApi';

interface CellData {
  r: number;
  c: number;
  v: any;
  f?: string | null;
  bg?: string | null;
  fg?: string | null;
  b?: boolean;
  i?: boolean;
  align?: string;
  format?: string;
  warn?: 'red' | 'yellow' | 'green' | null;
  review?: {
    id: number;
    excel_row: number;
    excel_col: number;
    warning_color: string;
    status: string;
    boq_price: number;
    lib_price?: number;
    difference?: number;
    diff_percent?: number;
    lib_source?: string;
    lib_description?: string;
    user_action?: string;
    user_note?: string;
  };
}

interface MergedRange {
  s_r: number;
  s_c: number;
  e_r: number;
  e_c: number;
}

interface ExcelWorkspaceProps {
  workbookId: string | null;
  activeSheetName?: string;
  targetCellCoord?: { row: number; col: number } | null;
  onWorkbookLoaded?: (wbId: string) => void;
  onNavigateToReview?: () => void;
}

export function ExcelWorkspace({
  workbookId,
  activeSheetName,
  targetCellCoord,
  onWorkbookLoaded,
  onNavigateToReview
}: ExcelWorkspaceProps) {
  const [currentWbId, setCurrentWbId] = useState<string | null>(workbookId);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [sheetNames, setSheetNames] = useState<string[]>([]);
  const [currentSheet, setCurrentSheet] = useState<string>('');
  const [cells, setCells] = useState<Record<string, CellData>>({});
  const [merges, setMerges] = useState<MergedRange[]>([]);
  const [rowCount, setRowCount] = useState<number>(35);
  const [colCount, setColCount] = useState<number>(18);
  const [currentFilename, setCurrentFilename] = useState<string>('BoQ-MEP.xlsx');

  // Active Ribbon Tab: 'home' | 'review' | 'data'
  const [ribbonTab, setRibbonTab] = useState<'home' | 'review' | 'data'>('home');

  // Selection & Range Selection
  const [selectedCell, setSelectedCell] = useState<{ r: number; c: number }>({ r: 4, c: 9 });
  const [selectionRange, setSelectionRange] = useState<{ s_r: number; s_c: number; e_r: number; e_c: number }>({
    s_r: 4, s_c: 9, e_r: 4, e_c: 9
  });
  const [isSelecting, setIsSelecting] = useState(false);

  // Inline Cell Editing
  const [editingCell, setEditingCell] = useState<{ r: number; c: number } | null>(null);
  const [editValue, setEditValue] = useState<string>('');
  const [formulaBarValue, setFormulaBarValue] = useState<string>('');

  // Undo / Redo & Save tracking
  const [undoStack, setUndoStack] = useState<Array<Record<string, CellData>>>([]);
  const [redoStack, setRedoStack] = useState<Array<Record<string, CellData>>>([]);
  const [hasUnsavedChanges, setHasUnsavedChanges] = useState(false);
  const [saving, setSaving] = useState(false);

  // Search & Filters
  const [searchQuery, setSearchQuery] = useState('');
  const [searchMatches, setSearchMatches] = useState<string[]>([]);
  const [searchMatchIndex, setSearchMatchIndex] = useState(0);
  const [warningFilter, setWarningFilter] = useState<'all' | 'red' | 'yellow' | 'green'>('all');

  // Price Inspector Drawer
  const [inspectorOpen, setInspectorOpen] = useState(false);
  const [actionBusy, setActionBusy] = useState(false);
  const [rememberMemory, setRememberMemory] = useState(true);
  const [customNote, setCustomNote] = useState('');
  const [customPriceInput, setCustomPriceInput] = useState('');

  // Context Menu
  const [contextMenu, setContextMenu] = useState<{ x: number; y: number; r: number; c: number } | null>(null);

  // Zoom
  const [zoomLevel, setZoomLevel] = useState(100);

  const gridContainerRef = useRef<HTMLDivElement>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);

  // Column letters A, B, C...
  const colLetters = useMemo(() => {
    const letters: string[] = [];
    for (let i = 1; i <= Math.max(colCount, 26); i++) {
      let temp = i;
      let colName = '';
      while (temp > 0) {
        const mod = (temp - 1) % 26;
        colName = String.fromCharCode(65 + mod) + colName;
        temp = Math.floor((temp - mod) / 26);
      }
      letters.push(colName);
    }
    return letters;
  }, [colCount]);

  // Load sheet data
  const loadSheet = useCallback(async (wbId: string, sName?: string) => {
    setLoading(true);
    setError(null);
    try {
      const url = sName
        ? `/api/mep/workbooks/${wbId}/sheet?sheet_name=${encodeURIComponent(sName)}`
        : `/api/mep/workbooks/${wbId}/sheet`;
      const data = await mepApi(url);
      setSheetNames(data.sheet_names || []);
      setCurrentSheet(data.active_sheet);
      setCells(data.cells || {});
      setMerges(data.merges || []);
      setRowCount(Math.max(data.row_count, 30));
      setColCount(Math.max(data.col_count, 15));
      setUndoStack([]);
      setRedoStack([]);
      setHasUnsavedChanges(false);
    } catch (err: any) {
      setError(err.message || 'Không thể tải sheet Excel');
    } finally {
      setLoading(false);
    }
  }, []);

  // If no workbook is passed, automatically load sample BOQ
  useEffect(() => {
    if (!currentWbId) {
      mepApi('/api/mep/workbooks/sample')
        .then(wb => {
          if (wb && wb.id) {
            setCurrentWbId(wb.id);
            setCurrentFilename(wb.original_filename || 'BoQ-MEP-Mau-Can-Soat-Gia.xlsx');
            if (onWorkbookLoaded) onWorkbookLoaded(wb.id);
            loadSheet(wb.id);
          }
        })
        .catch(err => console.warn('Could not load sample workbook:', err));
    } else {
      loadSheet(currentWbId, activeSheetName);
    }
  }, [currentWbId, activeSheetName, loadSheet, onWorkbookLoaded]);

  // Jump to coordinate if provided
  useEffect(() => {
    if (targetCellCoord) {
      const r = targetCellCoord.row;
      const c = targetCellCoord.col;
      setSelectedCell({ r, c });
      setSelectionRange({ s_r: r, s_c: c, e_r: r, e_c: c });
      setInspectorOpen(true);
      setTimeout(() => {
        const el = document.getElementById(`cell-${r}-${c}`);
        if (el) {
          el.scrollIntoView({ behavior: 'smooth', block: 'center', inline: 'center' });
        }
      }, 150);
    }
  }, [targetCellCoord]);

  // Active cell data
  const activeKey = `${selectedCell.r}_${selectedCell.c}`;
  const activeCellData = cells[activeKey];
  const activeCoordLabel = `${colLetters[selectedCell.c - 1] || 'A'}${selectedCell.r}`;

  // Sync formula bar
  useEffect(() => {
    if (activeCellData) {
      setFormulaBarValue(activeCellData.f || String(activeCellData.v ?? ''));
      if (activeCellData.review) {
        setInspectorOpen(true);
      }
    } else {
      setFormulaBarValue('');
    }
  }, [selectedCell, cells]);

  // Search logic
  useEffect(() => {
    if (!searchQuery.trim()) {
      setSearchMatches([]);
      setSearchMatchIndex(0);
      return;
    }
    const q = searchQuery.toLowerCase();
    const matched = Object.keys(cells).filter(k => {
      const v = String(cells[k]?.v ?? '').toLowerCase();
      return v.includes(q);
    });
    setSearchMatches(matched);
    setSearchMatchIndex(0);
    if (matched.length > 0) {
      const [rStr, cStr] = matched[0].split('_');
      const r = parseInt(rStr);
      const c = parseInt(cStr);
      setSelectedCell({ r, c });
      setSelectionRange({ s_r: r, s_c: c, e_r: r, e_c: c });
      const el = document.getElementById(`cell-${r}-${c}`);
      if (el) el.scrollIntoView({ behavior: 'smooth', block: 'center', inline: 'center' });
    }
  }, [searchQuery, cells]);

  const handleNextSearchMatch = () => {
    if (searchMatches.length === 0) return;
    const nextIdx = (searchMatchIndex + 1) % searchMatches.length;
    setSearchMatchIndex(nextIdx);
    const [rStr, cStr] = searchMatches[nextIdx].split('_');
    const r = parseInt(rStr);
    const c = parseInt(cStr);
    setSelectedCell({ r, c });
    setSelectionRange({ s_r: r, s_c: c, e_r: r, e_c: c });
    const el = document.getElementById(`cell-${r}-${c}`);
    if (el) el.scrollIntoView({ behavior: 'smooth', block: 'center', inline: 'center' });
  };

  const handlePrevSearchMatch = () => {
    if (searchMatches.length === 0) return;
    const prevIdx = (searchMatchIndex - 1 + searchMatches.length) % searchMatches.length;
    setSearchMatchIndex(prevIdx);
    const [rStr, cStr] = searchMatches[prevIdx].split('_');
    const r = parseInt(rStr);
    const c = parseInt(cStr);
    setSelectedCell({ r, c });
    setSelectionRange({ s_r: r, s_c: c, e_r: r, e_c: c });
    const el = document.getElementById(`cell-${r}-${c}`);
    if (el) el.scrollIntoView({ behavior: 'smooth', block: 'center', inline: 'center' });
  };

  // Commit Cell Edit
  const commitEdit = (r: number, c: number, newVal: string) => {
    const key = `${r}_${c}`;
    const prevData = cells[key] || { r, c, v: null };

    setUndoStack(prev => [...prev.slice(-25), { ...cells }]);
    setRedoStack([]);

    let parsedVal: any = newVal;
    let formula: string | null = null;
    if (newVal.startsWith('=')) {
      formula = newVal;
      parsedVal = evaluateSimpleFormula(newVal, cells);
    } else {
      const num = Number(newVal.replace(/,/g, ''));
      if (!isNaN(num) && newVal.trim() !== '') {
        parsedVal = num;
      }
    }

    setCells(prev => ({
      ...prev,
      [key]: {
        ...prevData,
        v: parsedVal,
        f: formula,
      }
    }));

    setEditingCell(null);
    setHasUnsavedChanges(true);
  };

  // Evaluate simple formulas
  function evaluateSimpleFormula(expr: string, cellMap: Record<string, CellData>): any {
    try {
      const clean = expr.slice(1).trim().toUpperCase();
      // SUM(A1:A5)
      const sumMatch = clean.match(/^SUM\(([A-Z]+)(\d+):([A-Z]+)(\d+)\)$/);
      if (sumMatch) {
        const c1 = colLetterToNumber(sumMatch[1]);
        const r1 = parseInt(sumMatch[2]);
        const c2 = colLetterToNumber(sumMatch[3]);
        const r2 = parseInt(sumMatch[4]);
        let total = 0;
        for (let row = r1; row <= r2; row++) {
          for (let col = c1; col <= c2; col++) {
            const v = cellMap[`${row}_${col}`]?.v;
            if (typeof v === 'number') total += v;
          }
        }
        return total;
      }

      // Arithmetic replacement
      let evalExpr = clean.replace(/([A-Z]+)(\d+)/g, (_, colStr, rowStr) => {
        const colNum = colLetterToNumber(colStr);
        const rowNum = parseInt(rowStr);
        const val = cellMap[`${rowNum}_${colNum}`]?.v;
        return typeof val === 'number' ? String(val) : '0';
      });

      if (/^[0-9+\-*/().\s]+$/.test(evalExpr)) {
        // eslint-disable-next-line no-eval
        return Function(`'use strict'; return (${evalExpr})`)();
      }
      return expr;
    } catch {
      return expr;
    }
  }

  function colLetterToNumber(letter: string): number {
    let result = 0;
    for (let i = 0; i < letter.length; i++) {
      result = result * 26 + (letter.charCodeAt(i) - 64);
    }
    return result;
  }

  // Format Formatting Helpers
  const applyCellStyle = (prop: 'b' | 'i' | 'align' | 'bg' | 'fg', val: any) => {
    setUndoStack(prev => [...prev.slice(-25), { ...cells }]);
    setCells(prev => {
      const next = { ...prev };
      const min_r = Math.min(selectionRange.s_r, selectionRange.e_r);
      const max_r = Math.max(selectionRange.s_r, selectionRange.e_r);
      const min_c = Math.min(selectionRange.s_c, selectionRange.e_c);
      const max_c = Math.max(selectionRange.s_c, selectionRange.e_c);

      for (let r = min_r; r <= max_r; r++) {
        for (let c = min_c; c <= max_c; c++) {
          const k = `${r}_${c}`;
          const current = next[k] || { r, c, v: null };
          if (prop === 'b') next[k] = { ...current, b: !current.b };
          else if (prop === 'i') next[k] = { ...current, i: !current.i };
          else if (prop === 'align') next[k] = { ...current, align: val };
          else if (prop === 'bg') next[k] = { ...current, bg: val };
          else if (prop === 'fg') next[k] = { ...current, fg: val };
        }
      }
      return next;
    });
    setHasUnsavedChanges(true);
  };

  // Undo / Redo
  const handleUndo = () => {
    if (undoStack.length === 0) return;
    const prev = undoStack[undoStack.length - 1];
    setRedoStack(cur => [...cur, { ...cells }]);
    setCells(prev);
    setUndoStack(cur => cur.slice(0, cur.length - 1));
    setHasUnsavedChanges(true);
  };

  const handleRedo = () => {
    if (redoStack.length === 0) return;
    const next = redoStack[redoStack.length - 1];
    setUndoStack(cur => [...cur, { ...cells }]);
    setCells(next);
    setRedoStack(cur => cur.slice(0, cur.length - 1));
    setHasUnsavedChanges(true);
  };

  // Save changes to backend openpyxl
  const handleSaveToWorkbook = async () => {
    if (!currentWbId || !currentSheet) return;
    setSaving(true);
    try {
      const updates = Object.values(cells).map(c => ({
        r: c.r,
        c: c.c,
        v: c.f || c.v
      }));
      await mepApi(`/api/mep/workbooks/${currentWbId}/cells`, {
        method: 'POST',
        body: JSON.stringify({
          sheet_name: currentSheet,
          updates
        })
      });
      setHasUnsavedChanges(false);
      alert('Đã lưu tất cả thay đổi vào file Excel trên máy tính!');
    } catch (err: any) {
      alert('Lỗi lưu file: ' + err.message);
    } finally {
      setSaving(false);
    }
  };

  // Upload file from Workspace toolbar
  const handleToolbarUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setLoading(true);
    const form = new FormData();
    form.append('file', file);
    try {
      const res = await mepApi('/api/mep/workbooks/upload', {
        method: 'POST',
        body: form
      });
      setCurrentWbId(res.id);
      setCurrentFilename(res.filename);
      if (onWorkbookLoaded) onWorkbookLoaded(res.id);
      await loadSheet(res.id);
      alert(`Đã nạp file Excel: ${res.filename}`);
    } catch (err: any) {
      alert('Lỗi nạp file: ' + err.message);
    } finally {
      setLoading(false);
      e.target.value = '';
    }
  };

  // Structural Actions: Insert / Delete Row / Col
  const handleStructureAction = async (action: 'insert_row' | 'delete_row' | 'insert_col' | 'delete_col' | 'add_sheet') => {
    if (!currentWbId || !currentSheet) return;
    try {
      setLoading(true);
      await mepApi(`/api/mep/workbooks/${currentWbId}/structure`, {
        method: 'POST',
        body: JSON.stringify({
          sheet_name: currentSheet,
          action,
          index: action.includes('row') ? selectedCell.r : selectedCell.c
        })
      });
      await loadSheet(currentWbId, currentSheet);
    } catch (err: any) {
      alert('Lỗi thao tác: ' + err.message);
    } finally {
      setLoading(false);
      setContextMenu(null);
    }
  };

  // Keyboard navigation & Shortcuts
  const handleKeyDown = (e: React.KeyboardEvent) => {
    if (editingCell) return; // let editor handle keys

    const { r, c } = selectedCell;

    // Arrows
    if (e.key === 'ArrowUp') {
      e.preventDefault();
      const newR = Math.max(1, r - 1);
      setSelectedCell({ r: newR, c });
      setSelectionRange({ s_r: newR, s_c: c, e_r: newR, e_c: c });
    } else if (e.key === 'ArrowDown') {
      e.preventDefault();
      const newR = Math.min(rowCount, r + 1);
      setSelectedCell({ r: newR, c });
      setSelectionRange({ s_r: newR, s_c: c, e_r: newR, e_c: c });
    } else if (e.key === 'ArrowLeft') {
      e.preventDefault();
      const newC = Math.max(1, c - 1);
      setSelectedCell({ r, c: newC });
      setSelectionRange({ s_r: r, s_c: newC, e_r: r, e_c: newC });
    } else if (e.key === 'ArrowRight') {
      e.preventDefault();
      const newC = Math.min(colCount, c + 1);
      setSelectedCell({ r, c: newC });
      setSelectionRange({ s_r: r, s_c: newC, e_r: r, e_c: newC });
    } else if (e.key === 'Tab') {
      e.preventDefault();
      const newC = e.shiftKey ? Math.max(1, c - 1) : Math.min(colCount, c + 1);
      setSelectedCell({ r, c: newC });
      setSelectionRange({ s_r: r, s_c: newC, e_r: r, e_c: newC });
    } else if (e.key === 'Enter') {
      e.preventDefault();
      setEditingCell({ r, c });
      setEditValue(activeCellData?.f || String(activeCellData?.v ?? ''));
    } else if (e.key === 'Delete' || e.key === 'Backspace') {
      e.preventDefault();
      commitEdit(r, c, '');
    } else if (e.key === 'F2') {
      e.preventDefault();
      setEditingCell({ r, c });
      setEditValue(activeCellData?.f || String(activeCellData?.v ?? ''));
    } else if (e.ctrlKey && e.key.toLowerCase() === 'z') {
      e.preventDefault();
      handleUndo();
    } else if (e.ctrlKey && e.key.toLowerCase() === 'y') {
      e.preventDefault();
      handleRedo();
    } else if (e.ctrlKey && e.key.toLowerCase() === 'c') {
      // Copy to clipboard as TSV
      e.preventDefault();
      handleCopySelection();
    } else if (e.ctrlKey && e.key.toLowerCase() === 'v') {
      // Paste from clipboard
      e.preventDefault();
      handlePasteClipboard();
    } else if (e.key.length === 1 && !e.ctrlKey && !e.altKey && !e.metaKey) {
      // Direct typing starts edit immediately like Excel
      setEditingCell({ r, c });
      setEditValue(e.key);
    }
  };

  // Copy selection to system clipboard
  const handleCopySelection = async () => {
    const min_r = Math.min(selectionRange.s_r, selectionRange.e_r);
    const max_r = Math.max(selectionRange.s_r, selectionRange.e_r);
    const min_c = Math.min(selectionRange.s_c, selectionRange.e_c);
    const max_c = Math.max(selectionRange.s_c, selectionRange.e_c);

    const rows: string[] = [];
    for (let r = min_r; r <= max_r; r++) {
      const rowVals: string[] = [];
      for (let c = min_c; c <= max_c; c++) {
        const val = cells[`${r}_${c}`]?.v ?? '';
        rowVals.push(String(val));
      }
      rows.push(rowVals.join('\t'));
    }
    const tsv = rows.join('\n');
    try {
      await navigator.clipboard.writeText(tsv);
    } catch {
      // fallback
    }
  };

  // Paste TSV from clipboard
  const handlePasteClipboard = async () => {
    try {
      const text = await navigator.clipboard.readText();
      if (!text) return;
      const lines = text.split(/\r?\n/).filter(line => line.length > 0);
      setUndoStack(prev => [...prev.slice(-25), { ...cells }]);

      setCells(prev => {
        const next = { ...prev };
        lines.forEach((line, rOffset) => {
          const vals = line.split('\t');
          vals.forEach((val, cOffset) => {
            const targetR = selectedCell.r + rOffset;
            const targetC = selectedCell.c + cOffset;
            const k = `${targetR}_${targetC}`;
            let parsedVal: any = val;
            const num = Number(val.replace(/,/g, ''));
            if (!isNaN(num) && val.trim() !== '') parsedVal = num;
            next[k] = {
              ...(next[k] || { r: targetR, c: targetC }),
              v: parsedVal,
              f: val.startsWith('=') ? val : null
            };
          });
        });
        return next;
      });
      setHasUnsavedChanges(true);
    } catch (err) {
      console.warn('Clipboard paste failed:', err);
    }
  };

  // Next Alert Jump
  const handleJumpNextAlert = () => {
    const alertKeys = Object.keys(cells).filter(k => {
      const c = cells[k];
      if (warningFilter === 'all') {
        return c.warn === 'red' || c.warn === 'yellow';
      }
      return c.warn === warningFilter;
    });

    if (alertKeys.length === 0) {
      alert('Không có ô cảnh báo nào phù hợp với bộ lọc hiện tại.');
      return;
    }

    const currIndex = alertKeys.indexOf(activeKey);
    const nextKey = currIndex >= 0 && currIndex < alertKeys.length - 1
      ? alertKeys[currIndex + 1]
      : alertKeys[0];

    const targetCell = cells[nextKey];
    if (targetCell) {
      setSelectedCell({ r: targetCell.r, c: targetCell.c });
      setSelectionRange({ s_r: targetCell.r, s_c: targetCell.c, e_r: targetCell.r, e_c: targetCell.c });
      setInspectorOpen(true);
      const el = document.getElementById(`cell-${targetCell.r}-${targetCell.c}`);
      if (el) el.scrollIntoView({ behavior: 'smooth', block: 'center', inline: 'center' });
    }
  };

  // Inspector Action: Apply Library Price
  const handleApplyLibPrice = async () => {
    if (!activeCellData?.review) return;
    const rev = activeCellData.review;
    if (!rev.lib_price) {
      alert('Chưa có đơn giá thư viện tham chiếu để áp dụng.');
      return;
    }

    setActionBusy(true);
    try {
      await mepApi(`/api/mep/review/item/${rev.id}/action`, {
        method: 'POST',
        body: JSON.stringify({
          action: 'applied_lib',
          user_note: customNote || 'Áp dụng giá tham chiếu từ thư viện',
          remember_memory: rememberMemory
        })
      });

      commitEdit(selectedCell.r, selectedCell.c, String(rev.lib_price));
      setCells(prev => ({
        ...prev,
        [activeKey]: {
          ...prev[activeKey],
          warn: 'green',
          review: {
            ...prev[activeKey].review!,
            boq_price: rev.lib_price!,
            warning_color: 'green',
            user_action: 'applied_lib',
            user_note: customNote || 'Đã áp dụng giá thư viện'
          }
        }
      }));
      alert('Đã cập nhật đơn giá theo thư viện thành công!');
    } catch (err: any) {
      alert('Lỗi: ' + err.message);
    } finally {
      setActionBusy(false);
    }
  };

  // Inspector Action: Keep BOQ Price
  const handleKeepBoqPrice = async () => {
    if (!activeCellData?.review) return;
    setActionBusy(true);
    try {
      await mepApi(`/api/mep/review/item/${activeCellData.review.id}/action`, {
        method: 'POST',
        body: JSON.stringify({
          action: 'kept_boq',
          user_note: customNote || 'Xác nhận giữ nguyên giá BOQ sau soát xét',
          remember_memory: rememberMemory
        })
      });

      setCells(prev => ({
        ...prev,
        [activeKey]: {
          ...prev[activeKey],
          warn: 'green',
          review: {
            ...prev[activeKey].review!,
            warning_color: 'green',
            user_action: 'kept_boq',
            user_note: customNote || 'Đã xác nhận giữ nguyên giá'
          }
        }
      }));
      alert('Đã xác nhận giữ nguyên đơn giá BOQ!');
    } catch (err: any) {
      alert('Lỗi: ' + err.message);
    } finally {
      setActionBusy(false);
    }
  };

  // Selected range stats for status bar (SUM, COUNT, AVG)
  const rangeStats = useMemo(() => {
    const min_r = Math.min(selectionRange.s_r, selectionRange.e_r);
    const max_r = Math.max(selectionRange.s_r, selectionRange.e_r);
    const min_c = Math.min(selectionRange.s_c, selectionRange.e_c);
    const max_c = Math.max(selectionRange.s_c, selectionRange.e_c);

    let sum = 0;
    let numCount = 0;
    let totalCells = (max_r - min_r + 1) * (max_c - min_c + 1);

    for (let r = min_r; r <= max_r; r++) {
      for (let c = min_c; c <= max_c; c++) {
        const val = cells[`${r}_${c}`]?.v;
        if (typeof val === 'number') {
          sum += val;
          numCount += 1;
        }
      }
    }
    return {
      totalCells,
      sum,
      count: numCount,
      avg: numCount > 0 ? sum / numCount : 0
    };
  }, [selectionRange, cells]);

  return (
    <div
      className="excel-workspace-container"
      tabIndex={0}
      onKeyDown={handleKeyDown}
      onClick={() => setContextMenu(null)}
    >
      {/* Hidden File Input for Toolbar Upload */}
      <input
        type="file"
        ref={fileInputRef}
        style={{ display: 'none' }}
        accept=".xlsx,.xlsm"
        onChange={handleToolbarUpload}
      />

      {/* Top Application Header Bar */}
      <div className="excel-title-bar">
        <div className="excel-title-left">
          <span className="excel-logo-pill">EXCEL</span>
          <span className="excel-filename-text">{currentFilename}</span>
          {hasUnsavedChanges && <span className="excel-unsaved-dot">● Chưa lưu</span>}
        </div>
        <div className="excel-title-right">
          <button className="ribbon-btn-cta" onClick={onNavigateToReview}>
            ⚡ Mở Bảng Soát Xét Đơn Giá
          </button>
        </div>
      </div>

      {/* Excel Ribbon Tabs Navigation */}
      <div className="excel-ribbon-tabs">
        <button
          className={`ribbon-tab-btn ${ribbonTab === 'home' ? 'active' : ''}`}
          onClick={() => setRibbonTab('home')}
        >
          Trang chủ
        </button>
        <button
          className={`ribbon-tab-btn ${ribbonTab === 'review' ? 'active' : ''}`}
          onClick={() => setRibbonTab('review')}
        >
          Soát xét giá MEP
        </button>
        <button
          className={`ribbon-tab-btn ${ribbonTab === 'data' ? 'active' : ''}`}
          onClick={() => setRibbonTab('data')}
        >
          Dữ liệu & Tìm kiếm
        </button>
      </div>

      {/* Excel Ribbon Toolbar Contents */}
      <div className="excel-ribbon-body">
        {ribbonTab === 'home' && (
          <div className="ribbon-panel-row">
            {/* File Group */}
            <div className="ribbon-group">
              <div className="ribbon-buttons">
                <button
                  className="ribbon-tool-btn"
                  onClick={() => fileInputRef.current?.click()}
                  title="Tải lên file Excel (.xlsx, .xlsm)"
                >
                  <span className="btn-icon">📂</span>
                  <span className="btn-text">Nạp Excel</span>
                </button>
                <button
                  className={`ribbon-tool-btn ${hasUnsavedChanges ? 'highlight-save' : ''}`}
                  onClick={handleSaveToWorkbook}
                  disabled={saving}
                  title="Lưu các chỉnh sửa vào file Excel"
                >
                  <span className="btn-icon">💾</span>
                  <span className="btn-text">{saving ? 'Đang lưu…' : 'Lưu file'}</span>
                </button>
                <a
                  href={`/api/mep/workbooks/${currentWbId}/download`}
                  download
                  className="ribbon-tool-btn"
                  title="Xuất tải file Excel về máy"
                >
                  <span className="btn-icon">📥</span>
                  <span className="btn-text">Xuất Excel</span>
                </a>
              </div>
              <span className="group-label">Tệp</span>
            </div>

            <div className="ribbon-divider" />

            {/* Clipboard Group */}
            <div className="ribbon-group">
              <div className="ribbon-buttons">
                <button className="ribbon-tool-btn" onClick={handleCopySelection} title="Sao chép (Ctrl+C)">
                  <span className="btn-icon">📋</span>
                  <span className="btn-text">Sao chép</span>
                </button>
                <button className="ribbon-tool-btn" onClick={handlePasteClipboard} title="Dán (Ctrl+V)">
                  <span className="btn-icon">📄</span>
                  <span className="btn-text">Dán</span>
                </button>
                <button className="ribbon-tool-btn" onClick={handleUndo} disabled={undoStack.length === 0} title="Hoàn tác (Ctrl+Z)">
                  <span className="btn-icon">↩</span>
                  <span className="btn-text">Hoàn tác</span>
                </button>
                <button className="ribbon-tool-btn" onClick={handleRedo} disabled={redoStack.length === 0} title="Làm lại (Ctrl+Y)">
                  <span className="btn-icon">↪</span>
                  <span className="btn-text">Làm lại</span>
                </button>
              </div>
              <span className="group-label">Bộ nhớ tạm</span>
            </div>

            <div className="ribbon-divider" />

            {/* Font & Formatting */}
            <div className="ribbon-group">
              <div className="ribbon-formatting-grid">
                <button className="fmt-mini-btn bold-btn" onClick={() => applyCellStyle('b', true)} title="In đậm (Bold)">
                  <strong>B</strong>
                </button>
                <button className="fmt-mini-btn italic-btn" onClick={() => applyCellStyle('i', true)} title="In nghiêng (Italic)">
                  <em>I</em>
                </button>
                <button className="fmt-mini-btn" onClick={() => applyCellStyle('align', 'left')} title="Căn trái">
                  ⫷
                </button>
                <button className="fmt-mini-btn" onClick={() => applyCellStyle('align', 'center')} title="Căn giữa">
                  ≡
                </button>
                <button className="fmt-mini-btn" onClick={() => applyCellStyle('align', 'right')} title="Căn phải">
                  ⫸
                </button>
                <button className="fmt-mini-btn" onClick={() => applyCellStyle('bg', '#fef08a')} title="Tô nền vàng">
                  🎨
                </button>
                <button className="fmt-mini-btn" onClick={() => applyCellStyle('bg', null)} title="Xóa màu nền">
                  ⚪
                </button>
              </div>
              <span className="group-label">Định dạng ô</span>
            </div>

            <div className="ribbon-divider" />

            {/* Insert / Delete Row / Col */}
            <div className="ribbon-group">
              <div className="ribbon-buttons">
                <button className="ribbon-tool-btn" onClick={() => handleStructureAction('insert_row')} title="Chèn dòng phía trên">
                  <span className="btn-icon">➕</span>
                  <span className="btn-text">Thêm dòng</span>
                </button>
                <button className="ribbon-tool-btn" onClick={() => handleStructureAction('delete_row')} title="Xóa dòng hiện tại">
                  <span className="btn-icon">➖</span>
                  <span className="btn-text">Xóa dòng</span>
                </button>
                <button className="ribbon-tool-btn" onClick={() => handleStructureAction('insert_col')} title="Chèn cột bên trái">
                  <span className="btn-icon">➕</span>
                  <span className="btn-text">Thêm cột</span>
                </button>
                <button className="ribbon-tool-btn" onClick={() => handleStructureAction('delete_col')} title="Xóa cột hiện tại">
                  <span className="btn-icon">➖</span>
                  <span className="btn-text">Xóa cột</span>
                </button>
              </div>
              <span className="group-label">Dòng & Cột</span>
            </div>
          </div>
        )}

        {ribbonTab === 'review' && (
          <div className="ribbon-panel-row">
            {/* Review actions */}
            <div className="ribbon-group">
              <div className="ribbon-buttons">
                <button className="ribbon-tool-btn cta-review-btn" onClick={onNavigateToReview} title="Chạy soát giá tự động">
                  <span className="btn-icon">🚀</span>
                  <span className="btn-text">Chạy soát giá</span>
                </button>
                <button className="ribbon-tool-btn" onClick={() => setInspectorOpen(!inspectorOpen)} title="Bật/tắt ngăn đối chiếu giá">
                  <span className="btn-icon">🔍</span>
                  <span className="btn-text">Ngăn đối chiếu</span>
                </button>
              </div>
              <span className="group-label">Tác vụ soát giá</span>
            </div>

            <div className="ribbon-divider" />

            {/* Filter Alerts */}
            <div className="ribbon-group">
              <div className="ribbon-filter-controls">
                <select
                  value={warningFilter}
                  onChange={e => setWarningFilter(e.target.value as any)}
                  className="ribbon-select"
                >
                  <option value="all">Tất cả ô cảnh báo</option>
                  <option value="red">🔴 Chỉ ô Đỏ (Vượt ngưỡng)</option>
                  <option value="yellow">🟡 Chỉ ô Vàng (Cần kiểm tra)</option>
                  <option value="green">🟢 Chỉ ô Xanh (Đạt chuẩn)</option>
                </select>
                <button className="ribbon-pill-btn jump-btn" onClick={handleJumpNextAlert}>
                  ⏩ Đến ô cảnh báo tiếp theo
                </button>
              </div>
              <span className="group-label">Bộ lọc cảnh báo</span>
            </div>
          </div>
        )}

        {ribbonTab === 'data' && (
          <div className="ribbon-panel-row">
            {/* Search group */}
            <div className="ribbon-group">
              <div className="ribbon-search-controls">
                <input
                  type="text"
                  placeholder="Tìm kiếm nội dung ô…"
                  value={searchQuery}
                  onChange={e => setSearchQuery(e.target.value)}
                  className="ribbon-search-input"
                />
                {searchMatches.length > 0 && (
                  <span className="search-count-pill">
                    {searchMatchIndex + 1} / {searchMatches.length}
                  </span>
                )}
                <button className="fmt-mini-btn" onClick={handlePrevSearchMatch} title="Kết quả trước">
                  ▲
                </button>
                <button className="fmt-mini-btn" onClick={handleNextSearchMatch} title="Kết quả tiếp">
                  ▼
                </button>
              </div>
              <span className="group-label">Tìm kiếm</span>
            </div>

            <div className="ribbon-divider" />

            {/* Sheet structure */}
            <div className="ribbon-group">
              <div className="ribbon-buttons">
                <button className="ribbon-tool-btn" onClick={() => handleStructureAction('add_sheet')} title="Thêm sheet mới">
                  <span className="btn-icon">📄+</span>
                  <span className="btn-text">Thêm Sheet</span>
                </button>
              </div>
              <span className="group-label">Bảng tính</span>
            </div>
          </div>
        )}
      </div>

      {/* Formula Bar */}
      <div className="excel-formulabar">
        <div className="excel-name-box" title="Toạ độ ô đang chọn">{activeCoordLabel}</div>
        <div className="excel-fx-icon" title="Hàm công thức">fx</div>
        <input
          type="text"
          className="excel-formula-input"
          value={formulaBarValue}
          onChange={e => {
            setFormulaBarValue(e.target.value);
            commitEdit(selectedCell.r, selectedCell.c, e.target.value);
          }}
          placeholder="Nhập giá trị hoặc công thức (ví dụ: =H4*I4 hoặc =SUM(J4:J15))"
        />
      </div>

      {/* Main Grid Viewport */}
      <div className="workspace-main-area">
        <div
          className="grid-viewport"
          ref={gridContainerRef}
          style={{ transform: `scale(${zoomLevel / 100})`, transformOrigin: 'top left' }}
        >
          {loading ? (
            <div className="grid-loading">Đang tải bảng tính Excel…</div>
          ) : error ? (
            <div className="grid-error">{error}</div>
          ) : (
            <table className="excel-table">
              <thead>
                <tr>
                  <th className="corner-header">#</th>
                  {Array.from({ length: colCount }, (_, i) => (
                    <th key={i} className="col-header">
                      {colLetters[i] || '—'}
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {Array.from({ length: rowCount }, (_, rIdx) => {
                  const r = rIdx + 1;
                  return (
                    <tr key={r}>
                      <td className="row-header">{r}</td>
                      {Array.from({ length: colCount }, (_, cIdx) => {
                        const c = cIdx + 1;
                        const key = `${r}_${c}`;
                        const cell = cells[key];
                        const isSelected = selectedCell.r === r && selectedCell.c === c;
                        const isEditing = editingCell?.r === r && editingCell?.c === c;

                        // Check if inside range selection
                        const inRange =
                          r >= Math.min(selectionRange.s_r, selectionRange.e_r) &&
                          r <= Math.max(selectionRange.s_r, selectionRange.e_r) &&
                          c >= Math.min(selectionRange.s_c, selectionRange.e_c) &&
                          c <= Math.max(selectionRange.s_c, selectionRange.e_c);

                        // Search match highlight
                        const strVal = String(cell?.v ?? '');
                        const isSearchMatch = searchQuery && strVal.toLowerCase().includes(searchQuery.toLowerCase());

                        // Style attributes
                        const style: React.CSSProperties = {};
                        if (cell?.bg) style.backgroundColor = cell.bg;
                        if (cell?.fg) style.color = cell.fg;
                        if (cell?.b) style.fontWeight = 'bold';
                        if (cell?.i) style.fontStyle = 'italic';
                        if (cell?.align) style.textAlign = cell.align as any;

                        // Warning tint
                        let warnClass = '';
                        if (cell?.warn === 'red') warnClass = 'cell-warn-red';
                        else if (cell?.warn === 'yellow') warnClass = 'cell-warn-yellow';
                        else if (cell?.warn === 'green') warnClass = 'cell-warn-green';

                        // Number formatting
                        let displayVal = cell?.v;
                        if (typeof displayVal === 'number' && !cell?.f) {
                          displayVal = displayVal.toLocaleString('vi-VN');
                        }

                        return (
                          <td
                            id={`cell-${r}-${c}`}
                            key={c}
                            style={style}
                            className={`grid-cell ${isSelected ? 'selected' : ''} ${inRange && !isSelected ? 'in-range' : ''} ${warnClass} ${isSearchMatch ? 'search-match' : ''}`}
                            onMouseDown={() => {
                              setSelectedCell({ r, c });
                              setSelectionRange({ s_r: r, s_c: c, e_r: r, e_c: c });
                              setIsSelecting(true);
                            }}
                            onMouseEnter={() => {
                              if (isSelecting) {
                                setSelectionRange(prev => ({ ...prev, e_r: r, e_c: c }));
                              }
                            }}
                            onMouseUp={() => setIsSelecting(false)}
                            onDoubleClick={() => {
                              setEditingCell({ r, c });
                              setEditValue(cell?.f || String(cell?.v ?? ''));
                            }}
                            onContextMenu={e => {
                              e.preventDefault();
                              setSelectedCell({ r, c });
                              setContextMenu({ x: e.clientX, y: e.clientY, r, c });
                            }}
                          >
                            {isEditing ? (
                              <input
                                autoFocus
                                className="inline-editor"
                                value={editValue}
                                onChange={e => setEditValue(e.target.value)}
                                onBlur={() => commitEdit(r, c, editValue)}
                                onKeyDown={e => {
                                  if (e.key === 'Enter') commitEdit(r, c, editValue);
                                  if (e.key === 'Escape') setEditingCell(null);
                                }}
                              />
                            ) : (
                              <div className="cell-content">
                                <span>{displayVal !== null && displayVal !== undefined ? String(displayVal) : ''}</span>
                                {cell?.review && (
                                  <span className={`cell-badge badge-${cell.warn}`} title="Bấm để xem chi tiết đối chiếu">
                                    {cell.warn === 'red' ? '!' : cell.warn === 'yellow' ? '?' : '✓'}
                                  </span>
                                )}
                              </div>
                            )}
                          </td>
                        );
                      })}
                    </tr>
                  );
                })}
              </tbody>
            </table>
          )}
        </div>

        {/* Right Price Inspector Drawer */}
        {inspectorOpen && activeCellData?.review && (
          <aside className="price-inspector-drawer">
            <div className="inspector-header">
              <h3>🔍 Chi tiết đối chiếu đơn giá</h3>
              <button className="close-btn" onClick={() => setInspectorOpen(false)}>×</button>
            </div>

            <div className="inspector-body">
              <div className={`status-banner banner-${activeCellData.warn}`}>
                <strong>
                  {activeCellData.warn === 'red'
                    ? '🔴 CẢNH BÁO: ĐƠN GIÁ VƯỢT NGƯỠNG'
                    : activeCellData.warn === 'yellow'
                    ? '🟡 CẦN KIỂM TRA / THẤP HƠN'
                    : '🟢 ĐẠT (TRONG NGƯỠNG CHO PHÉP)'}
                </strong>
                <p>Vị trí ô: <strong>{activeCoordLabel}</strong> (Dòng {selectedCell.r}, Cột {colLetters[selectedCell.c - 1]})</p>
              </div>

              {/* Comparison Box */}
              <div className="comparison-cards">
                <div className="compare-card boq-card">
                  <span className="card-label">Giá dự thầu BOQ</span>
                  <strong className="price-number">
                    {activeCellData.review.boq_price.toLocaleString('vi-VN')} đ
                  </strong>
                </div>

                <div className="compare-divider">vs</div>

                <div className="compare-card lib-card">
                  <span className="card-label">Giá thư viện tham chiếu</span>
                  <strong className="price-number">
                    {activeCellData.review.lib_price
                      ? `${activeCellData.review.lib_price.toLocaleString('vi-VN')} đ`
                      : 'Chưa có giá'}
                  </strong>
                </div>
              </div>

              {/* Difference stats */}
              {activeCellData.review.difference !== undefined && (
                <div className="diff-stats">
                  <div className="stat-item">
                    <span>Chênh lệch:</span>
                    <strong className={activeCellData.review.difference > 0 ? 'text-red' : 'text-green'}>
                      {activeCellData.review.difference > 0 ? '+' : ''}
                      {activeCellData.review.difference.toLocaleString('vi-VN')} đ
                    </strong>
                  </div>
                  <div className="stat-item">
                    <span>Tỷ lệ chênh lệch:</span>
                    <strong className={activeCellData.review.diff_percent! > 0 ? 'text-red' : 'text-green'}>
                      {activeCellData.review.diff_percent! > 0 ? '+' : ''}
                      {activeCellData.review.diff_percent}%
                    </strong>
                  </div>
                </div>
              )}

              {/* Matched details */}
              <div className="matched-meta">
                <h4>Thông tin nguồn tham chiếu</h4>
                {activeCellData.review.lib_description && (
                  <p><strong>Vật tư:</strong> {activeCellData.review.lib_description}</p>
                )}
                <p><strong>Nguồn:</strong> {activeCellData.review.lib_source || 'Thư viện MEP'}</p>
                {activeCellData.review.user_note && (
                  <p className="note-text"><strong>Ghi chú:</strong> {activeCellData.review.user_note}</p>
                )}
              </div>

              {/* Action Buttons */}
              <div className="inspector-actions">
                <h4>Hành động xử lý</h4>

                {activeCellData.review.lib_price && (
                  <button
                    className="action-btn primary-btn"
                    disabled={actionBusy}
                    onClick={handleApplyLibPrice}
                  >
                    ✓ Áp dụng giá thư viện ({activeCellData.review.lib_price.toLocaleString('vi-VN')} đ)
                  </button>
                )}

                <button
                  className="action-btn secondary-btn"
                  disabled={actionBusy}
                  onClick={handleKeepBoqPrice}
                >
                  ⚡ Giữ nguyên giá BOQ
                </button>

                <div className="note-block">
                  <label>Ghi chú lý do xử lý:</label>
                  <textarea
                    rows={2}
                    value={customNote}
                    onChange={e => setCustomNote(e.target.value)}
                    placeholder="Nhập giải trình / căn cứ điều chỉnh..."
                  />
                </div>

                <label className="checkbox-label">
                  <input
                    type="checkbox"
                    checked={rememberMemory}
                    onChange={e => setRememberMemory(e.target.checked)}
                  />
                  <span>Ghi nhớ cặp tương đương này cho lần sau</span>
                </label>
              </div>
            </div>
          </aside>
        )}
      </div>

      {/* Right Click Context Menu */}
      {contextMenu && (
        <div
          className="excel-context-menu"
          style={{ top: contextMenu.y, left: contextMenu.x }}
          onClick={e => e.stopPropagation()}
        >
          <button onClick={() => { handleCopySelection(); setContextMenu(null); }}>
            📋 Sao chép (Ctrl+C)
          </button>
          <button onClick={() => { handlePasteClipboard(); setContextMenu(null); }}>
            📄 Dán (Ctrl+V)
          </button>
          <div className="menu-divider" />
          <button onClick={() => handleStructureAction('insert_row')}>
            ➕ Chèn dòng phía trên
          </button>
          <button onClick={() => handleStructureAction('delete_row')}>
            🗑 Xóa dòng này
          </button>
          <button onClick={() => handleStructureAction('insert_col')}>
            ➕ Chèn cột bên trái
          </button>
          <button onClick={() => handleStructureAction('delete_col')}>
            🗑 Xóa cột này
          </button>
          {activeCellData?.review && (
            <>
              <div className="menu-divider" />
              <button onClick={() => { setInspectorOpen(true); setContextMenu(null); }}>
                🔍 Xem chi tiết đối chiếu giá
              </button>
            </>
          )}
        </div>
      )}

      {/* Bottom Sheet Switcher Bar */}
      <div className="excel-bottom-bar">
        <div className="sheet-tabs-container">
          {sheetNames.map(name => (
            <button
              key={name}
              className={`excel-sheet-tab ${name === currentSheet ? 'active' : ''}`}
              onClick={() => loadSheet(currentWbId!, name)}
            >
              📄 {name}
            </button>
          ))}
          <button
            className="excel-add-sheet-btn"
            onClick={() => handleStructureAction('add_sheet')}
            title="Thêm sheet mới"
          >
            ＋
          </button>
        </div>

        {/* Status & Range summary */}
        <div className="excel-status-info">
          <span className="status-ready-tag">SẴN SÀNG</span>
          {rangeStats.totalCells > 1 && (
            <span className="range-summary-pill">
              {rangeStats.count > 0 && `TỔNG CỘNG: ${rangeStats.sum.toLocaleString('vi-VN')} đ | `}
              ĐẾM: {rangeStats.totalCells}
              {rangeStats.count > 0 && ` | TB: ${Math.round(rangeStats.avg).toLocaleString('vi-VN')} đ`}
            </span>
          )}
          <span className="dim-info">{rowCount} dòng × {colCount} cột</span>
          <div className="zoom-controls">
            <button onClick={() => setZoomLevel(Math.max(60, zoomLevel - 10))}>−</button>
            <span>{zoomLevel}%</span>
            <button onClick={() => setZoomLevel(Math.min(150, zoomLevel + 10))}>＋</button>
          </div>
        </div>
      </div>
    </div>
  );
}
