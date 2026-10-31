/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useEffect, useMemo, useCallback } from 'react';
import {
  TableSawFactoryState,
  SheetRecord,
  PartItem,
  PdfMatrixItem,
  OvertimeLog
} from './types/factory';
import { TableSawThermalLabel } from './components/TableSawThermalLabel';
import jsPDF from 'jspdf';
import html2canvas from 'html2canvas';
import {
  Layers,
  FileText,
  Clock,
  Moon,
  Bot,
  RefreshCw,
  Download,
  Upload,
  Database,
  Edit,
  Save,
  ShieldCheck,
  ShieldAlert,
  CheckCircle2,
  Printer,
  FileDown,
  Search,
  Filter,
  ArrowRight,
  AlertTriangle,
  Play,
  Check,
  KeyRound,
  X,
  ExternalLink,
  ChevronLeft,
  ChevronRight,
  Eye,
  Scissors,
  Compass,
  Box,
  QrCode,
  Scan,
  Camera,
  AlertOctagon,
  Hourglass,
  Ruler,
  ChevronDown,
  ChevronUp
} from 'lucide-react';
import * as XLSX from 'xlsx';
import { QrScannerModal } from './components/QrScannerModal';

export interface DimensionGroup {
  key: string;
  l: number;
  w: number;
  t: number;
  core: string;
  total_qty: number;
  cut_qty: number;
  remaining_qty: number;
  pct_cut: number;
  area_m2: number;
  total_area_m2: number;
  pdf_sets: string[];
  part_names: string[];
  part_codes: string[];
  sheets: Array<{
    sheet_id: string;
    sheet_title: string;
    qty_on_sheet: number;
    status: 'QUEUED' | 'ASSIGNED' | 'CUTTING' | 'CUT';
    assigned_saw: string;
    assigned_team: string;
    pdf_set: string;
    part_name: string;
    part_code: string;
  }>;
  remaining_sheets: Array<{
    sheet_id: string;
    sheet_title: string;
    qty_on_sheet: number;
    status: 'QUEUED' | 'ASSIGNED' | 'CUTTING' | 'CUT';
    assigned_saw: string;
    part_name: string;
    part_code: string;
  }>;
}

export interface GroupedPartItem {
  id: string;
  code: string;
  name: string;
  pdf_set: string;
  pdf_sets: string[];
  l: number;
  w: number;
  t: number;
  core: string;
  cab: string;
  villa: string;
  lip: string;
  total_needed: number;
  total_cut: number;
  remaining: number;
  pct_cut: number;
  sheets_count: number;
  pending_sheets_count: number;
  sheets: Array<{
    sheet_id: string;
    sheet_title: string;
    qty_on_sheet: number;
    status: string;
    assigned_saw: string;
    assigned_team: string;
    pdf_set: string;
  }>;
  remaining_sheets: Array<{
    sheet_id: string;
    sheet_title: string;
    qty_on_sheet: number;
    status: string;
    assigned_saw: string;
  }>;
}

interface StandardPartProgress {
  id: string;
  code: string;
  name: string;
  groupId: string;
  groupLabel: string;
  totalQty: number;
  cutQty: number;
  qtyPerSet: number | null;
  setsSupported: number | null;
}

interface VillaGroupProgress {
  id: string;
  label: string;
  partGroups: StandardPartProgress[];
  mappedPartCount: number;
  unmappedGroups: StandardPartProgress[];
  potentialSets: number;
  totalPieces: number;
  cutPieces: number;
}

interface TeamCutProgress {
  team: string;
  piecesCut: number;
}

const STANDARD_VILLA_COUNT = 40;
const VILLA_PDF_GROUPS = [
  { id: 'master-side-cabinet', label: 'Master Bedroom + Side Cabinet', pdfSets: ['PDF 1', 'PDF 4'] },
  { id: 'bedroom-03', label: 'Bedroom 03', pdfSets: ['PDF 2'] },
  { id: 'bedroom-02', label: 'Bedroom 02', pdfSets: ['PDF 3'] },
  { id: 'bedroom-01-drawers', label: 'Bedroom 01 & Drawers', pdfSets: ['PDF 5'] }
] as const;
// The source cutlist includes "PAKING" as a misspelling of packing.
const EXCLUDED_VILLA_PART_NAME = /pack|paking|skirt/i;

export default function App() {
  // State from server
  const [serverState, setServerState] = useState<TableSawFactoryState | null>(null);
  const [loading, setLoading] = useState<boolean>(true);
  const [activeTab, setActiveTab] = useState<'dispatch' | 'villa-readiness' | 'dimension-groups' | 'parts-explorer' | 'pdf-matrix' | 'time-forecast' | 'overtime' | 'database' | 'gemma'>('dispatch');

  // Exact Dimension Groups (Ignore Names) State
  const [dimSearch, setDimSearch] = useState<string>('');
  const [dimThicknessFilter, setDimThicknessFilter] = useState<'ALL' | '18mm' | '8mm'>('ALL');
  const [dimStatusFilter, setDimStatusFilter] = useState<'ALL' | 'PENDING' | 'CUT'>('PENDING');
  const [dimPdfFilter, setDimPdfFilter] = useState<string>('ALL');
  const [dimSortBy, setDimSortBy] = useState<'REMAINING' | 'TOTAL' | 'LENGTH' | 'WIDTH' | 'AREA' | 'PROGRESS_ASC' | 'PROGRESS_DESC'>('REMAINING');
  const [dimPage, setDimPage] = useState<number>(1);
  const [dimPageSize, setDimPageSize] = useState<number>(50);
  const [expandedDimKey, setExpandedDimKey] = useState<string | null>(null);

  // Sheet & Parts Sorting Options
  const [sheetSortBy, setSheetSortBy] = useState<'NAME_ASC' | 'NAME_DESC' | 'PDF' | 'STATUS' | 'PARTS_COUNT' | 'DEFAULT'>('NAME_ASC');
  const [partSortBy, setPartSortBy] = useState<'NAME' | 'CODE' | 'PDF' | 'LENGTH' | 'QTY' | 'CAB'>('NAME');

  // Parts Explorer Tab State (Enhanced with Grouping by Code & Name)
  const [partsExplorerViewMode, setPartsExplorerViewMode] = useState<'GROUPED' | 'FLAT'>('GROUPED');
  const [partsExplorerPdf, setPartsExplorerPdf] = useState<string>('ALL');
  const [partsExplorerCodeFilter, setPartsExplorerCodeFilter] = useState<string>('ALL');
  const [partsExplorerStatusFilter, setPartsExplorerStatusFilter] = useState<'ALL' | 'PENDING' | 'CUT'>('ALL');
  const [partsExplorerSearch, setPartsExplorerSearch] = useState<string>('');
  const [partsExplorerSortBy, setPartsExplorerSortBy] = useState<'CODE' | 'NAME' | 'REMAINING' | 'TOTAL' | 'CUT' | 'DIM' | 'LENGTH' | 'PDF' | 'CAB' | 'PROGRESS_ASC' | 'PROGRESS_DESC'>('CODE');
  const [partsExplorerPage, setPartsExplorerPage] = useState<number>(1);
  const [partsExplorerPageSize, setPartsExplorerPageSize] = useState<number>(50);
  const [expandedPartId, setExpandedPartId] = useState<string | null>(null);

  // Diagram Inspection Modal
  const [diagramModalSheet, setDiagramModalSheet] = useState<SheetRecord | null>(null);
  const [rawSvgContent, setRawSvgContent] = useState<string>('');
  const [loadingSvg, setLoadingSvg] = useState<boolean>(false);

  // Load raw SVG content whenever diagramModalSheet opens
  useEffect(() => {
    if (!diagramModalSheet?.diagram_url) {
      setRawSvgContent('');
      return;
    }

    let isMounted = true;
    setLoadingSvg(true);
    fetch(diagramModalSheet.diagram_url)
      .then((res) => {
        if (!res.ok) throw new Error(`HTTP ${res.status}`);
        return res.text();
      })
      .then((text) => {
        if (isMounted) {
          setRawSvgContent(text);
          setLoadingSvg(false);
        }
      })
      .catch((err) => {
        console.warn('Failed to load raw SVG diagram:', err);
        if (isMounted) {
          setRawSvgContent('');
          setLoadingSvg(false);
        }
      });

    return () => {
      isMounted = false;
    };
  }, [diagramModalSheet]);

  // QR Code Camera Scanner Modal
  const [isQrScannerOpen, setIsQrScannerOpen] = useState<boolean>(false);

  // Sheet Quick Edit Modal State
  const [editingSheet, setEditingSheet] = useState<SheetRecord | null>(null);
  const [editStatus, setEditStatus] = useState<string>('QUEUED');
  const [editDuration, setEditDuration] = useState<string>('');
  const [editSaw, setEditSaw] = useState<string>('Saw 1');
  const [editTeam, setEditTeam] = useState<string>('Team A (Ahmed)');
  const [isUpdatingSheet, setIsUpdatingSheet] = useState<boolean>(false);

  // Bulk Excel/CSV Import & Batch Edit State
  const [bulkImportJson, setBulkImportJson] = useState<string>('');
  const [bulkImportStatus, setBulkImportStatus] = useState<string>('');
  const [isBulkImporting, setIsBulkImporting] = useState<boolean>(false);

  // Database Manager Table Batch Selection & Edit
  const [dbSelectedSheetIds, setDbSelectedSheetIds] = useState<Set<string>>(new Set());
  const [dbSearchQuery, setDbSearchQuery] = useState<string>('');
  const [dbBatchStatus, setDbBatchStatus] = useState<string>('');
  const [dbBatchSaw, setDbBatchSaw] = useState<string>('');
  const [dbBatchDuration, setDbBatchDuration] = useState<string>('');
  const [isDbBatchUpdating, setIsDbBatchUpdating] = useState<boolean>(false);

  const handleDbToggleSelect = (sheet_id: string, checked: boolean) => {
    setDbSelectedSheetIds((prev) => {
      const next = new Set(prev);
      if (checked) next.add(sheet_id);
      else next.delete(sheet_id);
      return next;
    });
  };

  const handleDbSelectAll = (checked: boolean, sheets: SheetRecord[]) => {
    if (checked) {
      setDbSelectedSheetIds(new Set(sheets.map((s) => s.sheet_id)));
    } else {
      setDbSelectedSheetIds(new Set());
    }
  };

  const handleApplyDbBatchEdit = async () => {
    if (!isEffectiveAdmin) {
      setShowPasscodeModal(true);
      return;
    }
    if (dbSelectedSheetIds.size === 0) return;
    setIsDbBatchUpdating(true);

    try {
      const updates: any[] = [];
      dbSelectedSheetIds.forEach((sid) => {
        const item: any = { sheet_id: sid };
        if (dbBatchStatus) item.status = dbBatchStatus;
        if (dbBatchSaw) {
          item.assigned_saw = dbBatchSaw;
          item.assigned_team = dbBatchSaw === 'Saw 1' ? 'Team A (Ahmed)' : 'Team B';
        }
        if (dbBatchDuration !== '' && !isNaN(parseFloat(dbBatchDuration))) {
          item.duration_min = parseFloat(dbBatchDuration);
        }
        updates.push(item);
      });

      const res = await fetch('/api/saw/bulk-import-updates', {
        method: 'POST',
        headers: getAdminHeaders(),
        body: JSON.stringify({ updates })
      });
      const data = await res.json();
      if (data.success) {
        setBulkImportStatus(`✅ Batch updated ${data.count} selected sheets in database!`);
        setDbSelectedSheetIds(new Set());
        setDbBatchStatus('');
        setDbBatchSaw('');
        setDbBatchDuration('');
        await fetchStatus();
      } else {
        setBulkImportStatus(`Batch Edit Failed: ${data.error}`);
      }
    } catch (err: any) {
      setBulkImportStatus(`Batch Edit Error: ${err.message}`);
    } finally {
      setIsDbBatchUpdating(false);
    }
  };

  // Multi-selection
  const [selectedSheetIds, setSelectedSheetIds] = useState<Set<string>>(new Set());
  const [cutlistBatchStatus, setCutlistBatchStatus] = useState<string>('');
  const [cutlistBatchDuration, setCutlistBatchDuration] = useState<string>('');
  const [isCutlistBatchUpdating, setIsCutlistBatchUpdating] = useState<boolean>(false);

  const handleApplyCutlistBatchEdit = async () => {
    if (!isEffectiveAdmin) {
      setShowPasscodeModal(true);
      return;
    }
    const sheet_ids = Array.from(selectedSheetIds);
    if (sheet_ids.length === 0) return;
    setIsCutlistBatchUpdating(true);

    try {
      const updates: any[] = [];
      sheet_ids.forEach((sid) => {
        const item: any = { sheet_id: sid };
        if (cutlistBatchStatus) item.status = cutlistBatchStatus;
        if (cutlistBatchDuration !== '' && !isNaN(parseFloat(cutlistBatchDuration))) {
          item.duration_min = parseFloat(cutlistBatchDuration);
        }
        updates.push(item);
      });

      const res = await fetch('/api/saw/bulk-import-updates', {
        method: 'POST',
        headers: getAdminHeaders(),
        body: JSON.stringify({ updates })
      });
      const data = await res.json();
      if (data.success) {
        setCutlistBatchStatus('');
        setCutlistBatchDuration('');
        setSelectedSheetIds(new Set());
        await fetchStatus();
      }
    } catch (err: any) {
      console.error('Cutlist batch edit error:', err);
    } finally {
      setIsCutlistBatchUpdating(false);
    }
  };

  // Filters
  const [searchQuery, setSearchQuery] = useState<string>('');
  const [materialFilter, setMaterialFilter] = useState<'ALL' | '18mm' | '8mm'>('ALL');
  const [pdfFilter, setPdfFilter] = useState<string>('ALL');
  const [sawFilter, setSawFilter] = useState<string>('ALL');
  const [statusFilter, setStatusFilter] = useState<string>('ALL');
  const [pageSize, setPageSize] = useState<number>(50);
  const [currentPage, setCurrentPage] = useState<number>(1);

  // Time & Forecast Interactive Controls
  const [customCycleTime, setCustomCycleTime] = useState<number>(8.5);
  const [plannedOtHours, setPlannedOtHours] = useState<number>(2.5);

  // Overtime Form
  const [otDate, setOtDate] = useState<string>(() => new Date().toISOString().split('T')[0]);
  const [otTeam, setOtTeam] = useState<string>('Team A (Ahmed)');
  const [otHours, setOtHours] = useState<string>('2.5');

  // Admin Override & Passcode Modal
  const [adminPasscodeOverride, setAdminPasscodeOverride] = useState<boolean>(false);
  const [showPasscodeModal, setShowPasscodeModal] = useState<boolean>(false);
  const [passcodeInput, setPasscodeInput] = useState<string>('');
  const [passcodeError, setPasscodeError] = useState<string>('');

  // Print Modal
  const [isPrintModalOpen, setIsPrintModalOpen] = useState<boolean>(false);
  const [labelsToPrint, setLabelsToPrint] = useState<Array<{ part: PartItem; sheetStart: number; pieceIndex: number; totalPieces: number }>>([]);
  const [printModalTitle, setPrintModalTitle] = useState<string>('');
  const [isExportingPdf, setIsExportingPdf] = useState<boolean>(false);

  // PDF Export Function (100mm x 50mm exact page size per label)
  const handleExportPdf = async () => {
    if (labelsToPrint.length === 0) return;
    setIsExportingPdf(true);

    try {
      // Create PDF with exact 100mm width x 50mm height
      const doc = new jsPDF({
        orientation: 'landscape',
        unit: 'mm',
        format: [100, 50]
      });

      // Query hidden print elements rendered in DOM
      const labelElements = document.querySelectorAll('#print-area .thermal-label');

      for (let i = 0; i < labelElements.length; i++) {
        const elem = labelElements[i] as HTMLElement;
        const canvas = await html2canvas(elem, {
          scale: 3, // High DPI for barcodes
          useCORS: true,
          backgroundColor: '#ffffff'
        });

        const imgData = canvas.toDataURL('image/png');
        if (i > 0) {
          doc.addPage([100, 50], 'landscape');
        }
        doc.addImage(imgData, 'PNG', 0, 0, 100, 50);
      }

      const fileName = labelsToPrint.length === 1
        ? `thermal_label_S${labelsToPrint[0].sheetStart}_${labelsToPrint[0].part.code || 'part'}.pdf`
        : `batch_thermal_labels_${labelsToPrint.length}_stickers.pdf`;

      doc.save(fileName);
    } catch (err) {
      console.error('Failed to export PDF:', err);
    } finally {
      setIsExportingPdf(false);
    }
  };

  // Gemma AI State
  const [gemmaPrompt, setGemmaPrompt] = useState<string>('');
  const [gemmaResponse, setGemmaResponse] = useState<string>('');
  const [gemmaLoading, setGemmaLoading] = useState<boolean>(false);

  // Effective Admin Authorization
  const isEffectiveAdmin = Boolean(serverState?.isAdmin || adminPasscodeOverride);

  // Helper for admin headers
  const getAdminHeaders = useCallback(() => {
    const headers: Record<string, string> = { 'Content-Type': 'application/json' };
    if (adminPasscodeOverride) {
      headers['x-admin-key'] = 'admin-192-168-1-173';
    }
    return headers;
  }, [adminPasscodeOverride]);

  // Fetch status from /api/saw-factory-status
  const fetchStatus = useCallback(async () => {
    try {
      const res = await fetch('/api/saw-factory-status', {
        headers: getAdminHeaders()
      });
      if (res.ok) {
        const data = await res.json();
        setServerState(data);
        if (data.analytics?.avg_cycle_min && customCycleTime === 8.5) {
          setCustomCycleTime(data.analytics.avg_cycle_min);
        }
      }
    } catch (e) {
      console.warn('Unable to reach /api/saw-factory-status', e);
    } finally {
      setLoading(false);
    }
  }, [getAdminHeaders, customCycleTime]);

  useEffect(() => {
    fetchStatus();
    const interval = setInterval(fetchStatus, 4000);
    return () => clearInterval(interval);
  }, [fetchStatus]);

  // Unlock Passcode Handler
  const handleUnlockWithPasscode = () => {
    const clean = passcodeInput.trim().toLowerCase();
    if (clean === '192.168.1.173' || clean === 'admin' || clean === 'bazar3' || clean === 'mactech') {
      setAdminPasscodeOverride(true);
      setShowPasscodeModal(false);
      setPasscodeInput('');
      setPasscodeError('');
      fetchStatus();
    } else {
      setPasscodeError('Invalid passcode. Enter 192.168.1.173 or admin.');
    }
  };

  // Actions
  const handleToggleCut = async (sheet_id: string, currentStatus: string) => {
    if (!isEffectiveAdmin) {
      setShowPasscodeModal(true);
      return;
    }
    const endpoint = currentStatus === 'CUTTING' ? '/api/saw/finish-cut' : '/api/saw/start-cut';
    try {
      await fetch(endpoint, {
        method: 'POST',
        headers: getAdminHeaders(),
        body: JSON.stringify({ sheet_id })
      });
      await fetchStatus();
    } catch (e) {
      console.error(e);
    }
  };

  const handleToggleSaw2Repair = async () => {
    if (!isEffectiveAdmin) {
      setShowPasscodeModal(true);
      return;
    }
    try {
      await fetch('/api/saw/toggle-repair', {
        method: 'POST',
        headers: getAdminHeaders(),
        body: '{}'
      });
      await fetchStatus();
    } catch (e) {
      console.error(e);
    }
  };

  const handleRerouteAllToSaw1 = async () => {
    if (!isEffectiveAdmin) {
      setShowPasscodeModal(true);
      return;
    }
    if (!window.confirm('Reroute all pending sheets from Saw 2 to Saw 1 (Team A)?')) return;
    try {
      await fetch('/api/saw/reroute-to-saw1', {
        method: 'POST',
        headers: getAdminHeaders(),
        body: '{}'
      });
      await fetchStatus();
    } catch (e) {
      console.error(e);
    }
  };

  const handleBatchAssign = async (saw: string, team: string) => {
    if (!isEffectiveAdmin) {
      setShowPasscodeModal(true);
      return;
    }
    const sheet_ids = Array.from(selectedSheetIds);
    if (!sheet_ids.length) return;
    try {
      await fetch('/api/saw/batch-assign', {
        method: 'POST',
        headers: getAdminHeaders(),
        body: JSON.stringify({ sheet_ids, saw, team })
      });
      setSelectedSheetIds(new Set());
      await fetchStatus();
    } catch (e) {
      console.error(e);
    }
  };

  const handleLogOvertime = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!isEffectiveAdmin) {
      setShowPasscodeModal(true);
      return;
    }
    if (!otDate || !otHours) return;
    try {
      await fetch('/api/overtime/log', {
        method: 'POST',
        headers: getAdminHeaders(),
        body: JSON.stringify({ date: otDate, team: otTeam, hours: otHours })
      });
      setOtHours('');
      await fetchStatus();
    } catch (e) {
      console.error(e);
    }
  };

  const handleReparseCutlist = async () => {
    if (!isEffectiveAdmin) {
      setShowPasscodeModal(true);
      return;
    }
    try {
      setLoading(true);
      const res = await fetch('/api/reparse-cutlist', {
        method: 'POST',
        headers: getAdminHeaders()
      });
      if (res.ok) {
        await fetchStatus();
        alert('Cutlist re-parsed successfully from cutlist.html!');
      }
    } catch (e) {
      console.error(e);
    } finally {
      setLoading(false);
    }
  };

  // Open Quick Edit Modal for a sheet
  const handleOpenEditSheet = (sheet: SheetRecord) => {
    setEditingSheet(sheet);
    setEditStatus(sheet.status);
    setEditDuration(sheet.duration_min !== null && sheet.duration_min !== undefined ? sheet.duration_min.toString() : '');
    setEditSaw(sheet.assigned_saw || 'Saw 1');
    setEditTeam(sheet.assigned_team || 'Team A (Ahmed)');
  };

  // Submit Quick Edit Modal changes to server
  const handleSaveSheetEdit = async () => {
    if (!editingSheet) return;
    if (!isEffectiveAdmin) {
      setShowPasscodeModal(true);
      return;
    }
    setIsUpdatingSheet(true);
    try {
      const res = await fetch('/api/saw/update-sheet', {
        method: 'POST',
        headers: getAdminHeaders(),
        body: JSON.stringify({
          sheet_id: editingSheet.sheet_id,
          status: editStatus,
          duration_min: editDuration,
          assigned_saw: editSaw,
          assigned_team: editTeam
        })
      });
      if (res.ok) {
        setEditingSheet(null);
        await fetchStatus();
      }
    } catch (e) {
      console.error(e);
    } finally {
      setIsUpdatingSheet(false);
    }
  };

  // Handle Excel or CSV file upload for database update
  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = async (evt) => {
      try {
        const data = evt.target?.result;
        const workbook = XLSX.read(data, { type: 'binary' });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        const json = XLSX.utils.sheet_to_json<any>(worksheet);

        // Normalize rows to standard keys: sheet_id, status, duration_min, assigned_saw, assigned_team
        const normalized = json.map((row) => ({
          sheet_id: row['Sheet ID'] || row['sheet_id'] || row['Sheet'] || row['ID'],
          status: row['Status'] || row['status'],
          duration_min: row['Duration (Minutes)'] || row['Duration (min)'] || row['duration_min'] || row['Duration'] || row['Minutes'],
          assigned_saw: row['Assigned Saw'] || row['assigned_saw'] || row['Saw'],
          assigned_team: row['Assigned Team'] || row['assigned_team'] || row['Team']
        })).filter((x) => x.sheet_id);

        setBulkImportJson(JSON.stringify(normalized, null, 2));
        setBulkImportStatus(`Parsed ${normalized.length} valid sheet records from ${file.name}. Review below and click "Apply Updates to Database".`);
      } catch (err: any) {
        setBulkImportStatus(`Error parsing file: ${err.message}`);
      }
    };
    reader.readAsBinaryString(file);
  };

  // Apply parsed bulk updates to server
  const handleApplyBulkUpdates = async () => {
    if (!isEffectiveAdmin) {
      setShowPasscodeModal(true);
      return;
    }
    if (!bulkImportJson) return;
    setIsBulkImporting(true);
    try {
      const updates = JSON.parse(bulkImportJson);
      const res = await fetch('/api/saw/bulk-import-updates', {
        method: 'POST',
        headers: getAdminHeaders(),
        body: JSON.stringify({ updates })
      });
      const data = await res.json();
      if (data.success) {
        setBulkImportStatus(`✅ Successfully updated ${data.count} sheets in factory database!`);
        await fetchStatus();
      } else {
        setBulkImportStatus(`Failed: ${data.error}`);
      }
    } catch (err: any) {
      setBulkImportStatus(`Error: ${err.message}`);
    } finally {
      setIsBulkImporting(false);
    }
  };

  // Gemma AI Question
  const handleAskGemma = async (promptToAsk?: string) => {
    const q = promptToAsk || gemmaPrompt;
    if (!q.trim()) return;
    setGemmaLoading(true);
    setGemmaResponse('');
    try {
      const res = await fetch('/api/gemma/ask', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ prompt: q })
      });
      const data = await res.json();
      setGemmaResponse(data.reply || 'No response received.');
    } catch (err: any) {
      setGemmaResponse(`Error contacting Gemma: ${err.message}`);
    } finally {
      setGemmaLoading(false);
    }
  };

  // Printing single sheet or batch
  const prepareLabelsForSheet = (s: SheetRecord) => {
    const labels: Array<{ part: PartItem; sheetStart: number; pieceIndex: number; totalPieces: number }> = [];
    s.parts.forEach((p) => {
      for (let i = 1; i <= p.qty_per_sheet; i++) {
        labels.push({
          part: p,
          sheetStart: s.start_sheet,
          pieceIndex: i,
          totalPieces: p.qty_per_sheet
        });
      }
    });
    return labels;
  };

  const handlePrintSingleSheet = (s: SheetRecord) => {
    const labels = prepareLabelsForSheet(s);
    setLabelsToPrint(labels);
    setPrintModalTitle(`Print Sheet ${s.start_sheet} Labels (${labels.length} Stickers)`);
    setIsPrintModalOpen(true);
  };

  const handlePrintSelected = () => {
    if (selectedSheetIds.size === 0) return;
    const allLabels: Array<{ part: PartItem; sheetStart: number; pieceIndex: number; totalPieces: number }> = [];
    selectedSheetIds.forEach((sid) => {
      const s = (serverState?.sheets || []).find((x) => x.sheet_id === sid);
      if (s) {
        allLabels.push(...prepareLabelsForSheet(s));
      }
    });
    setLabelsToPrint(allLabels);
    setPrintModalTitle(`Batch Print: ${selectedSheetIds.size} Sheets (${allLabels.length} Total Stickers)`);
    setIsPrintModalOpen(true);
  };

  // Filtered Sheets
  const filteredSheets = useMemo(() => {
    let list = serverState?.sheets || [];

    if (materialFilter === '18mm') {
      list = list.filter((s) => s.thickness === 18 || !s.thickness);
    } else if (materialFilter === '8mm') {
      list = list.filter((s) => s.thickness === 8);
    }

    if (pdfFilter !== 'ALL') {
      list = list.filter((s) => s.pdf_sets.includes(pdfFilter));
    }
    if (sawFilter !== 'ALL') {
      list = list.filter((s) => s.assigned_saw === sawFilter);
    }
    if (statusFilter === 'STALE_QUEUED') {
      const nowSec = Math.floor(Date.now() / 1000);
      list = list.filter((s) => s.status === 'QUEUED' && s.queued_at && nowSec - s.queued_at >= 24 * 3600);
    } else if (statusFilter !== 'ALL') {
      list = list.filter((s) => s.status === statusFilter);
    }
    if (searchQuery.trim()) {
      const q = searchQuery.toLowerCase();
      list = list.filter((s) => {
        return (
          s.sheet_title.toLowerCase().includes(q) ||
          s.sheet_id.toLowerCase().includes(q) ||
          s.material.toLowerCase().includes(q) ||
          s.pdf_sets.some((pdf) => pdf.toLowerCase().includes(q)) ||
          s.rip_cuts.some((rc) => rc.toLowerCase().includes(q)) ||
          s.parts.some(
            (p) =>
              p.code.toLowerCase().includes(q) ||
              p.name.toLowerCase().includes(q) ||
              p.villa.toLowerCase().includes(q) ||
              p.cab.toLowerCase().includes(q)
          )
        );
      });
    }

    if (sheetSortBy === 'NAME_ASC') {
      list = [...list].sort((a, b) => a.sheet_title.localeCompare(b.sheet_title));
    } else if (sheetSortBy === 'NAME_DESC') {
      list = [...list].sort((a, b) => b.sheet_title.localeCompare(a.sheet_title));
    } else if (sheetSortBy === 'PDF') {
      list = [...list].sort((a, b) => (a.pdf_sets[0] || '').localeCompare(b.pdf_sets[0] || ''));
    } else if (sheetSortBy === 'STATUS') {
      const weight: Record<string, number> = { CUTTING: 1, QUEUED: 2, ASSIGNED: 3, CUT: 4 };
      list = [...list].sort((a, b) => (weight[a.status] || 5) - (weight[b.status] || 5));
    } else if (sheetSortBy === 'PARTS_COUNT') {
      list = [...list].sort((a, b) => b.parts.length - a.parts.length);
    }

    return list;
  }, [serverState?.sheets, materialFilter, pdfFilter, sawFilter, statusFilter, searchQuery, sheetSortBy]);

  // Paginated Sheets
  const totalPages = Math.ceil(filteredSheets.length / pageSize) || 1;
  const paginatedSheets = useMemo(() => {
    if (pageSize >= 9999) return filteredSheets;
    const start = (currentPage - 1) * pageSize;
    return filteredSheets.slice(start, start + pageSize);
  }, [filteredSheets, currentPage, pageSize]);

  // Helper to sort component parts inside a sheet row
  const getSortedParts = useCallback((parts: PartItem[], sortBy: 'NAME' | 'CODE' | 'PDF' | 'LENGTH' | 'QTY' | 'CAB'): PartItem[] => {
    const list = [...parts];
    if (sortBy === 'NAME') {
      list.sort((a, b) => a.name.localeCompare(b.name));
    } else if (sortBy === 'CODE') {
      list.sort((a, b) => a.code.localeCompare(b.code));
    } else if (sortBy === 'PDF') {
      list.sort((a, b) => a.pdf_set.localeCompare(b.pdf_set) || a.name.localeCompare(b.name));
    } else if (sortBy === 'LENGTH') {
      list.sort((a, b) => b.l - a.l);
    } else if (sortBy === 'QTY') {
      list.sort((a, b) => b.qty_per_sheet - a.qty_per_sheet);
    } else if (sortBy === 'CAB') {
      list.sort((a, b) => a.cab.localeCompare(b.cab) || a.name.localeCompare(b.name));
    }
    return list;
  }, []);

  // Flattened parts for Parts Explorer
  const allFlattenedParts = useMemo(() => {
    if (!serverState?.sheets) return [];
    const items: Array<{
      part: PartItem;
      sheet_id: string;
      sheet_title: string;
      material: string;
      thickness: number;
      sheet_status: string;
      assigned_saw: string;
      assigned_team: string;
    }> = [];

    serverState.sheets.forEach((s) => {
      s.parts.forEach((p) => {
        items.push({
          part: p,
          sheet_id: s.sheet_id,
          sheet_title: s.sheet_title,
          material: s.material,
          thickness: s.thickness || 18,
          sheet_status: s.status,
          assigned_saw: s.assigned_saw,
          assigned_team: s.assigned_team
        });
      });
    });
    return items;
  }, [serverState?.sheets]);

  const villaSetProgress = useMemo(() => {
    const groups = new Map<string, StandardPartProgress>();
    const teamCutTotals = new Map<string, number>();
    const groupByPdf = new Map<string, (typeof VILLA_PDF_GROUPS)[number]>();
    VILLA_PDF_GROUPS.forEach((group) => {
      group.pdfSets.forEach((pdfSet) => groupByPdf.set(pdfSet, group));
    });

    (serverState?.sheets || []).forEach((sheet) => {
      sheet.parts.forEach((part) => {
        if (
          part.villa.trim().toUpperCase() !== 'STANDARD 40' ||
          EXCLUDED_VILLA_PART_NAME.test(part.name)
        ) {
          return;
        }

        const villaGroup = groupByPdf.get(part.pdf_set);
        if (!villaGroup) return;

        const id = JSON.stringify([villaGroup.id, part.code, part.name, part.l, part.w, part.t || sheet.thickness || 18]);
        const qty = part.qty_per_sheet * (sheet.repeat_count || 1);
        let group = groups.get(id);

        if (!group) {
          group = {
            id,
            code: part.code,
            name: part.name,
            groupId: villaGroup.id,
            groupLabel: villaGroup.label,
            totalQty: 0,
            cutQty: 0,
            qtyPerSet: null,
            setsSupported: null
          };
          groups.set(id, group);
        }

        group.totalQty += qty;
        if (sheet.status === 'CUT') {
          group.cutQty += qty;
          const team = sheet.assigned_team.trim() || 'Unassigned';
          teamCutTotals.set(team, (teamCutTotals.get(team) || 0) + qty);
        }
      });
    });

    const partGroups = Array.from(groups.values()).map((group) => {
      const qtyPerSet = group.totalQty % STANDARD_VILLA_COUNT === 0
        ? group.totalQty / STANDARD_VILLA_COUNT
        : null;
      return {
        ...group,
        qtyPerSet,
        setsSupported: qtyPerSet === null ? null : Math.floor(group.cutQty / qtyPerSet)
      };
    });
    const mappedGroups = partGroups.filter((group) => group.qtyPerSet !== null);
    const unmappedGroups = partGroups.filter((group) => group.qtyPerSet === null);
    const villaGroups: VillaGroupProgress[] = VILLA_PDF_GROUPS.map((villaGroup) => {
      const groupParts = partGroups.filter((group) => group.groupId === villaGroup.id);
      const mappedPartGroups = groupParts.filter((group) => group.qtyPerSet !== null);
      const groupUnmappedGroups = groupParts.filter((group) => group.qtyPerSet === null);
      return {
        id: villaGroup.id,
        label: villaGroup.label,
        partGroups: groupParts,
        mappedPartCount: mappedPartGroups.length,
        unmappedGroups: groupUnmappedGroups,
        potentialSets: mappedPartGroups.length > 0
          ? Math.min(STANDARD_VILLA_COUNT, ...mappedPartGroups.map((group) => group.setsSupported || 0))
          : 0,
        totalPieces: groupParts.reduce((total, group) => total + group.totalQty, 0),
        cutPieces: groupParts.reduce((total, group) => total + group.cutQty, 0)
      };
    });
    const potentialSets = villaGroups.length > 0
      ? Math.min(...villaGroups.map((group) => group.potentialSets))
      : 0;
    const totalPieces = partGroups.reduce((total, group) => total + group.totalQty, 0);
    const cutPieces = partGroups.reduce((total, group) => total + group.cutQty, 0);
    const teams: TeamCutProgress[] = Array.from(teamCutTotals, ([team, piecesCut]) => ({
      team,
      piecesCut
    })).sort((a, b) => b.piecesCut - a.piecesCut);

    return {
      partGroups: partGroups.sort((a, b) => a.name.localeCompare(b.name)),
      villaGroups,
      mappedPartCount: mappedGroups.length,
      unmappedGroups,
      potentialSets,
      verifiedReadySets: partGroups.length > 0 && villaGroups.every((group) => group.partGroups.length > 0) && unmappedGroups.length === 0
        ? potentialSets
        : null,
      totalPieces,
      cutPieces,
      teams
    };
  }, [serverState?.sheets]);

  // Unique Part Codes across all parts
  const uniquePartCodes = useMemo(() => {
    const codes = new Set<string>();
    allFlattenedParts.forEach((item) => {
      if (item.part.code) codes.add(item.part.code.toUpperCase());
    });
    return Array.from(codes).sort();
  }, [allFlattenedParts]);

  // Grouped Parts (Aggregated by Part Code & Name)
  const groupedParts: GroupedPartItem[] = useMemo(() => {
    if (!serverState?.sheets) return [];

    const map = new Map<string, GroupedPartItem>();

    serverState.sheets.forEach((s) => {
      const sRepeat = s.repeat_count || 1;
      const isCut = s.status === 'CUT';

      s.parts.forEach((p) => {
        const code = (p.code || 'UNKNOWN').toUpperCase();
        const name = p.name || 'Unnamed Part';
        const key = `${code}___${name}___${p.l}x${p.w}x${p.t || s.thickness || 18}`;
        const qtyOnBoard = p.qty_per_sheet * sRepeat;

        let g = map.get(key);
        if (!g) {
          g = {
            id: key,
            code,
            name,
            pdf_set: p.pdf_set || '',
            pdf_sets: [],
            l: p.l,
            w: p.w,
            t: p.t || s.thickness || 18,
            core: p.core || s.material,
            cab: p.cab || '',
            villa: p.villa || '',
            lip: p.lip || '',
            total_needed: 0,
            total_cut: 0,
            remaining: 0,
            pct_cut: 0,
            sheets_count: 0,
            pending_sheets_count: 0,
            sheets: [],
            remaining_sheets: []
          };
          map.set(key, g);
        }

        g.total_needed += qtyOnBoard;
        if (isCut) {
          g.total_cut += qtyOnBoard;
        } else {
          g.remaining += qtyOnBoard;
          g.remaining_sheets.push({
            sheet_id: s.sheet_id,
            sheet_title: s.sheet_title,
            qty_on_sheet: qtyOnBoard,
            status: s.status,
            assigned_saw: s.assigned_saw
          });
        }

        if (p.pdf_set && !g.pdf_sets.includes(p.pdf_set)) {
          g.pdf_sets.push(p.pdf_set);
        }

        g.sheets.push({
          sheet_id: s.sheet_id,
          sheet_title: s.sheet_title,
          qty_on_sheet: qtyOnBoard,
          status: s.status,
          assigned_saw: s.assigned_saw,
          assigned_team: s.assigned_team,
          pdf_set: p.pdf_set
        });
      });
    });

    map.forEach((g) => {
      g.pct_cut = g.total_needed > 0 ? Math.round((g.total_cut / g.total_needed) * 100) : 0;
      g.sheets_count = g.sheets.length;
      g.pending_sheets_count = g.remaining_sheets.length;
      g.pdf_sets.sort();
    });

    return Array.from(map.values());
  }, [serverState?.sheets]);

  // Filtered & Sorted Grouped Parts
  const filteredGroupedParts = useMemo(() => {
    let list = groupedParts;

    if (partsExplorerPdf !== 'ALL') {
      list = list.filter((g) => g.pdf_sets.includes(partsExplorerPdf) || g.pdf_set === partsExplorerPdf);
    }

    if (partsExplorerCodeFilter !== 'ALL') {
      list = list.filter((g) => g.code.toUpperCase() === partsExplorerCodeFilter.toUpperCase());
    }

    if (partsExplorerStatusFilter === 'PENDING') {
      list = list.filter((g) => g.remaining > 0);
    } else if (partsExplorerStatusFilter === 'CUT') {
      list = list.filter((g) => g.remaining === 0);
    }

    if (partsExplorerSearch.trim()) {
      const q = partsExplorerSearch.toLowerCase().trim();
      list = list.filter((g) => {
        return (
          g.code.toLowerCase().includes(q) ||
          `[${g.code.toLowerCase()}]`.includes(q) ||
          g.name.toLowerCase().includes(q) ||
          g.cab.toLowerCase().includes(q) ||
          g.villa.toLowerCase().includes(q) ||
          g.core.toLowerCase().includes(q) ||
          `${g.l}x${g.w}`.includes(q) ||
          `${g.l} × ${g.w}`.includes(q)
        );
      });
    }

    list = [...list].sort((a, b) => {
      if (partsExplorerSortBy === 'CODE') {
        return a.code.localeCompare(b.code) || a.name.localeCompare(b.name);
      }
      if (partsExplorerSortBy === 'NAME') {
        return a.name.localeCompare(b.name) || a.code.localeCompare(b.code);
      }
      if (partsExplorerSortBy === 'REMAINING') {
        return b.remaining - a.remaining || b.total_needed - a.total_needed;
      }
      if (partsExplorerSortBy === 'TOTAL') {
        return b.total_needed - a.total_needed;
      }
      if (partsExplorerSortBy === 'CUT') {
        return b.total_cut - a.total_cut;
      }
      if (partsExplorerSortBy === 'DIM') {
        return b.l * b.w - a.l * a.w;
      }
      if (partsExplorerSortBy === 'LENGTH') {
        return b.l - a.l || b.w - a.w;
      }
      if (partsExplorerSortBy === 'PDF') {
        return a.pdf_set.localeCompare(b.pdf_set) || a.code.localeCompare(b.code);
      }
      if (partsExplorerSortBy === 'CAB') {
        return a.cab.localeCompare(b.cab) || a.code.localeCompare(b.code);
      }
      if (partsExplorerSortBy === 'PROGRESS_ASC') {
        return a.pct_cut - b.pct_cut || b.remaining - a.remaining;
      }
      if (partsExplorerSortBy === 'PROGRESS_DESC') {
        return b.pct_cut - a.pct_cut;
      }
      return 0;
    });

    return list;
  }, [groupedParts, partsExplorerPdf, partsExplorerCodeFilter, partsExplorerStatusFilter, partsExplorerSearch, partsExplorerSortBy]);

  const groupedTotalPages = Math.ceil(filteredGroupedParts.length / partsExplorerPageSize) || 1;
  const paginatedGroupedParts = useMemo(() => {
    if (partsExplorerPageSize >= 9999) return filteredGroupedParts;
    const start = (partsExplorerPage - 1) * partsExplorerPageSize;
    return filteredGroupedParts.slice(start, start + partsExplorerPageSize);
  }, [filteredGroupedParts, partsExplorerPage, partsExplorerPageSize]);

  // KPIs for Parts Explorer
  const partsKpis = useMemo(() => {
    const totalUnique = groupedParts.length;
    const pendingParts = groupedParts.filter((g) => g.remaining > 0).length;
    const completedParts = groupedParts.filter((g) => g.remaining === 0).length;
    const totalNeeded = groupedParts.reduce((a, b) => a + b.total_needed, 0);
    const totalCut = groupedParts.reduce((a, b) => a + b.total_cut, 0);
    const totalRemaining = groupedParts.reduce((a, b) => a + b.remaining, 0);
    const pct = totalNeeded > 0 ? Math.round((totalCut / totalNeeded) * 100) : 0;
    return { totalUnique, pendingParts, completedParts, totalNeeded, totalCut, totalRemaining, pct };
  }, [groupedParts]);

  // Filtered & Sorted Parts for Parts Explorer Tab (Flat mode)
  const filteredExplorerParts = useMemo(() => {
    let list = allFlattenedParts;

    if (partsExplorerPdf !== 'ALL') {
      list = list.filter((item) => item.part.pdf_set === partsExplorerPdf);
    }

    if (partsExplorerCodeFilter !== 'ALL') {
      list = list.filter((item) => item.part.code.toUpperCase() === partsExplorerCodeFilter.toUpperCase());
    }

    if (partsExplorerStatusFilter === 'PENDING') {
      list = list.filter((item) => item.sheet_status !== 'CUT');
    } else if (partsExplorerStatusFilter === 'CUT') {
      list = list.filter((item) => item.sheet_status === 'CUT');
    }

    if (partsExplorerSearch.trim()) {
      const q = partsExplorerSearch.toLowerCase();
      list = list.filter((item) => {
        const p = item.part;
        return (
          p.code.toLowerCase().includes(q) ||
          `[${p.code.toLowerCase()}]`.includes(q) ||
          p.name.toLowerCase().includes(q) ||
          p.cab.toLowerCase().includes(q) ||
          p.villa.toLowerCase().includes(q) ||
          p.pdf_set.toLowerCase().includes(q) ||
          item.sheet_title.toLowerCase().includes(q) ||
          item.sheet_id.toLowerCase().includes(q)
        );
      });
    }

    list = [...list].sort((a, b) => {
      if (partsExplorerSortBy === 'PDF') {
        return (
          a.part.pdf_set.localeCompare(b.part.pdf_set) ||
          a.part.code.localeCompare(b.part.code)
        );
      }
      if (partsExplorerSortBy === 'CODE') {
        return a.part.code.localeCompare(b.part.code);
      }
      if (partsExplorerSortBy === 'NAME') {
        return a.part.name.localeCompare(b.part.name);
      }
      if (partsExplorerSortBy === 'DIM') {
        return b.part.l * b.part.w - a.part.l * a.part.w;
      }
      if (partsExplorerSortBy === 'TOTAL') {
        return b.part.total_batch_qty - a.part.total_batch_qty;
      }
      if (partsExplorerSortBy === 'CAB') {
        return a.part.cab.localeCompare(b.part.cab);
      }
      return 0;
    });

    return list;
  }, [allFlattenedParts, partsExplorerPdf, partsExplorerCodeFilter, partsExplorerStatusFilter, partsExplorerSearch, partsExplorerSortBy]);

  const explorerTotalPages = Math.ceil(filteredExplorerParts.length / partsExplorerPageSize) || 1;
  const paginatedExplorerParts = useMemo(() => {
    if (partsExplorerPageSize >= 9999) return filteredExplorerParts;
    const start = (partsExplorerPage - 1) * partsExplorerPageSize;
    return filteredExplorerParts.slice(start, start + partsExplorerPageSize);
  }, [filteredExplorerParts, partsExplorerPage, partsExplorerPageSize]);

  // Export Grouped Parts to Excel
  const handleExportGroupedPartsExcel = () => {
    const rows = filteredGroupedParts.map((g) => ({
      'Part Code': `[${g.code}]`,
      'Part Name': g.name,
      'Dimensions': `${g.l} × ${g.w} × ${g.t} mm`,
      'Length (mm)': g.l,
      'Width (mm)': g.w,
      'Thickness (mm)': g.t,
      'Material / Core': g.core,
      'Cabinet': g.cab,
      'Villa': g.villa,
      'Edge Banding / Lip': g.lip,
      'Total Required (pcs)': g.total_needed,
      'Already Cut (pcs)': g.total_cut,
      'REMAINING TO CUT (pcs)': g.remaining,
      '% Completed': `${g.pct_cut}%`,
      'Total Sheets Count': g.sheets_count,
      'Pending Sheets Count': g.pending_sheets_count,
      'Drawing Sets': g.pdf_sets.join(', ')
    }));

    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Grouped Parts Summary');
    XLSX.writeFile(wb, `Grouped_Parts_Summary_${new Date().toISOString().split('T')[0]}.xlsx`);
  };

  // ==============================================================
  // EXACT SIZE & DIMENSION GROUPING (IGNORE PART NAMES)
  // ==============================================================
  const dimensionGroups: DimensionGroup[] = useMemo(() => {
    if (!serverState?.sheets) return [];

    const map = new Map<string, DimensionGroup>();

    serverState.sheets.forEach((s) => {
      const sRepeat = s.repeat_count || 1;
      const isCut = s.status === 'CUT';

      s.parts.forEach((p) => {
        const l = p.l;
        const w = p.w;
        const t = p.t || s.thickness || 18;
        const key = `${l}×${w}×${t}`;
        const qtyOnBoard = p.qty_per_sheet * sRepeat;

        let g = map.get(key);
        if (!g) {
          g = {
            key,
            l,
            w,
            t,
            core: p.core || s.material,
            total_qty: 0,
            cut_qty: 0,
            remaining_qty: 0,
            pct_cut: 0,
            area_m2: (l * w) / 1000000,
            total_area_m2: 0,
            pdf_sets: [],
            part_names: [],
            part_codes: [],
            sheets: [],
            remaining_sheets: []
          };
          map.set(key, g);
        }

        g.total_qty += qtyOnBoard;
        if (isCut) {
          g.cut_qty += qtyOnBoard;
        } else {
          g.remaining_qty += qtyOnBoard;
          g.remaining_sheets.push({
            sheet_id: s.sheet_id,
            sheet_title: s.sheet_title,
            qty_on_sheet: qtyOnBoard,
            status: s.status,
            assigned_saw: s.assigned_saw,
            part_name: p.name,
            part_code: p.code
          });
        }

        if (p.pdf_set && !g.pdf_sets.includes(p.pdf_set)) {
          g.pdf_sets.push(p.pdf_set);
        }
        if (p.name && !g.part_names.includes(p.name)) {
          g.part_names.push(p.name);
        }
        if (p.code && !g.part_codes.includes(p.code)) {
          g.part_codes.push(p.code);
        }

        g.sheets.push({
          sheet_id: s.sheet_id,
          sheet_title: s.sheet_title,
          qty_on_sheet: qtyOnBoard,
          status: s.status,
          assigned_saw: s.assigned_saw,
          assigned_team: s.assigned_team,
          pdf_set: p.pdf_set,
          part_name: p.name,
          part_code: p.code
        });
      });
    });

    map.forEach((g) => {
      g.pct_cut = g.total_qty > 0 ? Math.round((g.cut_qty / g.total_qty) * 100) : 0;
      g.total_area_m2 = parseFloat((g.area_m2 * g.total_qty).toFixed(2));
      g.pdf_sets.sort();
    });

    return Array.from(map.values());
  }, [serverState?.sheets]);

  // Filtered Dimension Groups
  const filteredDimGroups = useMemo(() => {
    let list = dimensionGroups;

    if (dimThicknessFilter === '18mm') {
      list = list.filter((g) => g.t === 18);
    } else if (dimThicknessFilter === '8mm') {
      list = list.filter((g) => g.t === 8);
    }

    if (dimStatusFilter === 'PENDING') {
      list = list.filter((g) => g.remaining_qty > 0);
    } else if (dimStatusFilter === 'CUT') {
      list = list.filter((g) => g.remaining_qty === 0);
    }

    if (dimPdfFilter !== 'ALL') {
      list = list.filter((g) => g.pdf_sets.includes(dimPdfFilter));
    }

    if (dimSearch.trim()) {
      const q = dimSearch.toLowerCase().trim();
      list = list.filter((g) => {
        return (
          g.key.toLowerCase().includes(q) ||
          String(g.l).includes(q) ||
          String(g.w).includes(q) ||
          String(g.t).includes(q) ||
          g.core.toLowerCase().includes(q) ||
          g.part_names.some((n) => n.toLowerCase().includes(q))
        );
      });
    }

    list = [...list].sort((a, b) => {
      if (dimSortBy === 'REMAINING') {
        return b.remaining_qty - a.remaining_qty || b.total_qty - a.total_qty;
      }
      if (dimSortBy === 'TOTAL') {
        return b.total_qty - a.total_qty;
      }
      if (dimSortBy === 'LENGTH') {
        return b.l - a.l || b.w - a.w;
      }
      if (dimSortBy === 'WIDTH') {
        return b.w - a.w || b.l - a.l;
      }
      if (dimSortBy === 'AREA') {
        return b.total_area_m2 - a.total_area_m2;
      }
      if (dimSortBy === 'PROGRESS_ASC') {
        return a.pct_cut - b.pct_cut || b.remaining_qty - a.remaining_qty;
      }
      if (dimSortBy === 'PROGRESS_DESC') {
        return b.pct_cut - a.pct_cut;
      }
      return 0;
    });

    return list;
  }, [dimensionGroups, dimThicknessFilter, dimStatusFilter, dimPdfFilter, dimSearch, dimSortBy]);

  const dimTotalPages = Math.ceil(filteredDimGroups.length / dimPageSize) || 1;
  const paginatedDimGroups = useMemo(() => {
    if (dimPageSize >= 9999) return filteredDimGroups;
    const start = (dimPage - 1) * dimPageSize;
    return filteredDimGroups.slice(start, start + dimPageSize);
  }, [filteredDimGroups, dimPage, dimPageSize]);

  // KPIs across all dimension groups
  const dimKpis = useMemo(() => {
    const totalSizes = dimensionGroups.length;
    const pendingSizes = dimensionGroups.filter((g) => g.remaining_qty > 0).length;
    const completedSizes = dimensionGroups.filter((g) => g.remaining_qty === 0).length;
    const totalPieces = dimensionGroups.reduce((acc, g) => acc + g.total_qty, 0);
    const cutPieces = dimensionGroups.reduce((acc, g) => acc + g.cut_qty, 0);
    const remainingPieces = dimensionGroups.reduce((acc, g) => acc + g.remaining_qty, 0);
    const pct = totalPieces > 0 ? Math.round((cutPieces / totalPieces) * 100) : 0;
    return { totalSizes, pendingSizes, completedSizes, totalPieces, cutPieces, remainingPieces, pct };
  }, [dimensionGroups]);

  // Export Dimensions Excel Summary
  const handleExportDimensionsExcel = () => {
    const rows = filteredDimGroups.map((g) => ({
      'Length (mm)': g.l,
      'Width (mm)': g.w,
      'Thickness (mm)': g.t,
      'Exact Dimension': `${g.l} × ${g.w} × ${g.t} mm`,
      'Material / Core': g.core,
      'Total Required (pcs)': g.total_qty,
      'Cut So Far (pcs)': g.cut_qty,
      'REMAINING TO CUT (pcs)': g.remaining_qty,
      '% Completed': `${g.pct_cut}%`,
      'Total Area (m²)': g.total_area_m2,
      'PDF Sets': g.pdf_sets.join(', '),
      'Shared Part Names Count': g.part_names.length,
      'All Part Names (Ignored)': g.part_names.join(' | ')
    }));

    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Exact Sizes Breakdown');
    XLSX.writeFile(wb, `Exact_Sizes_Cutlist_Summary_${new Date().toISOString().split('T')[0]}.xlsx`);
  };

  // Multi-select toggle
  const toggleSelect = (sheet_id: string, checked: boolean) => {
    setSelectedSheetIds((prev) => {
      const next = new Set(prev);
      if (checked) next.add(sheet_id);
      else next.delete(sheet_id);
      return next;
    });
  };

  const handleSelectAll = (checked: boolean) => {
    if (checked) {
      const allIds = new Set(filteredSheets.map((s) => s.sheet_id));
      setSelectedSheetIds(allIds);
    } else {
      setSelectedSheetIds(new Set());
    }
  };

  // Calculations for KPIs
  const totalSheetsVolume = useMemo(() => {
    return (serverState?.sheets || []).reduce((sum, s) => sum + s.repeat_count, 0);
  }, [serverState?.sheets]);

  const totalSheetsVolume18mm = useMemo(() => {
    return (serverState?.sheets || [])
      .filter((s) => s.thickness === 18 || !s.thickness)
      .reduce((sum, s) => sum + s.repeat_count, 0);
  }, [serverState?.sheets]);

  const totalSheetsVolume8mm = useMemo(() => {
    return (serverState?.sheets || [])
      .filter((s) => s.thickness === 8)
      .reduce((sum, s) => sum + s.repeat_count, 0);
  }, [serverState?.sheets]);

  const totalCutVolume = useMemo(() => {
    return (serverState?.sheets || [])
      .filter((s) => s.status === 'CUT')
      .reduce((sum, s) => sum + s.repeat_count, 0);
  }, [serverState?.sheets]);

  const staleQueuedSheetsVolume = useMemo(() => {
    const nowSec = Math.floor(Date.now() / 1000);
    return (serverState?.sheets || [])
      .filter((s) => s.status === 'QUEUED' && s.queued_at && nowSec - s.queued_at >= 24 * 3600)
      .reduce((sum, s) => sum + s.repeat_count, 0);
  }, [serverState?.sheets]);

  const cutPercent = totalSheetsVolume
    ? Math.round((totalCutVolume / totalSheetsVolume) * 100)
    : 0;

  const saw1 = serverState?.saws['Saw 1'] || {
    status: 'ONLINE',
    team: 'Team A (Ahmed)',
    assigned_sheets: 0,
    cut_sheets: 0
  };
  const saw2 = serverState?.saws['Saw 2'] || {
    status: 'IN_REPAIR',
    team: 'Team B',
    assigned_sheets: 0,
    cut_sheets: 0
  };

  // Live ETA Forecast Calculation
  const forecastData = useMemo(() => {
    const pendingSheets = (serverState?.sheets || [])
      .filter((s) => s.status !== 'CUT')
      .reduce((sum, s) => sum + s.repeat_count, 0);
    const activeSaws = saw2.status === 'IN_REPAIR' ? 1 : 2;
    const cycleMin = customCycleTime || 8.5;
    const totalMinNeeded = (pendingSheets * cycleMin) / activeSaws;
    const regularMinPerDay = 420; // 7.0h
    const otMinPerDay = (plannedOtHours || 0) * 60;
    const effectiveMinPerDay = regularMinPerDay + otMinPerDay;
    const workDaysNeeded = effectiveMinPerDay > 0 ? (totalMinNeeded / effectiveMinPerDay).toFixed(1) : '0';

    // Projected completion calendar date
    const target = new Date();
    let remainingWorkMin = totalMinNeeded;
    while (remainingWorkMin > 0) {
      target.setDate(target.getDate() + 1);
      // skip Fridays (day 5)
      if (target.getDay() !== 5) {
        remainingWorkMin -= effectiveMinPerDay;
      }
    }
    const dateStr = target.toLocaleDateString(undefined, {
      weekday: 'long',
      month: 'short',
      day: 'numeric'
    });

    return {
      pendingSheets,
      activeSaws,
      cycleMin,
      totalMinNeeded: Math.round(totalMinNeeded),
      effectiveMinPerDay,
      regularMinPerDay,
      workDaysNeeded,
      dateStr,
      dailyCapacity: Math.round((effectiveMinPerDay / cycleMin) * activeSaws)
    };
  }, [serverState?.sheets, saw2.status, customCycleTime, plannedOtHours]);

  return (
    <>
      <div id="main-ui-container" className="min-h-screen bg-[#070b14] text-slate-100 flex flex-col font-sans">
      {/* Topbar Header */}
      <header className="bg-[#0f172a] border-b border-[#223354] px-4 py-2.5 flex flex-wrap items-center justify-between gap-3 sticky top-0 z-40 shadow-lg">
        <div className="flex items-center gap-3">
          <div className="bg-gradient-to-br from-sky-600 to-sky-400 text-black font-black px-3 py-1.5 rounded-lg text-sm tracking-wide shadow-md">
            G4-SAW
          </div>
          <div>
            <div className="font-black text-sm sm:text-base tracking-tight text-white flex items-center gap-2">
              <span>BAZAR V3 — TABLE SAW DISPATCH &amp; PDF MATRIX MES</span>
            </div>
            <div className="text-[11px] text-slate-400 font-medium">
              Shift: 8AM–12PM &amp; 1PM–4PM (7h regular) · Rip-Cut Sequence · 100×50mm Thermal Labels
            </div>
          </div>
        </div>

        {/* Action Controls & Authorization Badge */}
        <div className="flex items-center gap-2.5 flex-wrap">
          {isEffectiveAdmin ? (
            <div className="inline-flex items-center gap-1.5 px-3 py-1 rounded-full text-xs font-black bg-emerald-500/20 text-emerald-300 border border-emerald-500/40 shadow-sm">
              <ShieldCheck className="w-3.5 h-3.5 text-emerald-400" />
              <span>
                🟢 ADMIN ({serverState?.clientIp || '192.168.1.173'}) — DISPATCH ACTIVE
              </span>
            </div>
          ) : (
            <button
              onClick={() => setShowPasscodeModal(true)}
              className="inline-flex items-center gap-1.5 px-3 py-1 rounded-full text-xs font-black bg-rose-500/20 text-rose-300 border border-rose-500/40 hover:bg-rose-500/30 transition cursor-pointer shadow-sm"
              title="Click to enter admin passcode"
            >
              <ShieldAlert className="w-3.5 h-3.5 text-rose-400" />
              <span>
                🔒 VIEW ONLY ({serverState?.clientIp || 'Floor LAN'}) — CLICK TO UNLOCK
              </span>
            </button>
          )}

          <button
            onClick={() => setIsQrScannerOpen(true)}
            className="px-3.5 py-1.5 bg-gradient-to-r from-sky-500 to-indigo-600 hover:from-sky-400 hover:to-indigo-500 text-white rounded-lg text-xs font-black flex items-center gap-2 shadow-lg shadow-sky-500/25 transition transform active:scale-95 cursor-pointer ring-1 ring-sky-300/40"
            title="Scan thermal part sticker barcode with camera to inspect specs & verify cut"
          >
            <Camera className="w-4 h-4 animate-pulse" />
            <span>📷 SCAN BARCODE</span>
          </button>

          <button
            onClick={fetchStatus}
            className="px-2.5 py-1.5 bg-[#131d33] hover:bg-[#1c2a4a] text-slate-200 border border-[#223354] rounded-lg text-xs font-bold flex items-center gap-1.5 transition"
            title="Refresh Factory Data"
          >
            <RefreshCw className="w-3.5 h-3.5 text-sky-400" />
            <span>↻ Refresh</span>
          </button>

          <a
            href="/api/export-master-excel"
            className="px-3 py-1.5 bg-emerald-600 hover:bg-emerald-500 text-white rounded-lg text-xs font-bold flex items-center gap-1.5 transition shadow"
            download
            title="Download full 3-sheet Excel database (.xlsx) with all cuts and parts"
          >
            <Download className="w-3.5 h-3.5" />
            <span>📊 Export Excel DB (.xlsx)</span>
          </a>

          <a
            href="/api/export-progress-csv"
            className="px-2.5 py-1.5 bg-sky-600 hover:bg-sky-500 text-white rounded-lg text-xs font-bold flex items-center gap-1.5 transition shadow"
            download
            title="Download CSV log"
          >
            <Download className="w-3.5 h-3.5" />
            <span>CSV</span>
          </a>

          <button
            onClick={handleReparseCutlist}
            className="px-2.5 py-1.5 bg-[#131d33] hover:bg-slate-700 text-slate-300 border border-[#223354] rounded-lg text-xs font-bold transition"
            title="Re-read and parse cutlist.html from data folder"
          >
            Re-parse HTML
          </button>
        </div>
      </header>

      {/* Machine Status Bar */}
      <section className="bg-[#0b1325] border-b border-[#223354] px-4 py-2 flex items-center gap-3 flex-wrap">
        <span className="text-xs font-extrabold text-slate-400 tracking-wider uppercase mr-1">
          SAWS STATUS:
        </span>

        {/* Saw 1 */}
        <div className="inline-flex items-center gap-2 px-3 py-1 rounded-lg text-xs font-extrabold bg-emerald-500/15 border border-emerald-500 text-emerald-300">
          <span>🟢 Saw 1: <strong>ONLINE</strong></span>
          <span>• Team A (Ahmed)</span>
          <span className="bg-black/40 px-2 py-0.5 rounded text-[11px] text-emerald-200 font-mono">
            {saw1.cut_sheets} / {saw1.assigned_sheets} cut
          </span>
        </div>

        {/* Saw 2 */}
        <div
          className={`inline-flex items-center gap-2 px-3 py-1 rounded-lg text-xs font-extrabold border ${
            saw2.status === 'IN_REPAIR'
              ? 'bg-rose-500/20 border-rose-500 text-rose-300'
              : 'bg-emerald-500/15 border-emerald-500 text-emerald-300'
          }`}
        >
          <span>
            {saw2.status === 'IN_REPAIR' ? '⚠️ Saw 2: UNDER REPAIR' : '🟢 Saw 2: ONLINE'}
          </span>
          <span>• Team B</span>
          <span className="bg-black/40 px-2 py-0.5 rounded text-[11px] font-mono">
            {saw2.cut_sheets} / {saw2.assigned_sheets} cut
          </span>
          <button
            onClick={handleToggleSaw2Repair}
            className="ml-1 px-2 py-0.5 bg-sky-600 hover:bg-sky-500 text-white rounded text-[10px] font-bold transition"
          >
            Toggle Status
          </button>
        </div>

        {/* Emergency Reroute Button */}
        <div className="ml-auto">
          <button
            onClick={handleRerouteAllToSaw1}
            className="px-3 py-1 bg-rose-600/90 hover:bg-rose-600 text-white rounded-lg text-xs font-bold flex items-center gap-1.5 transition shadow"
          >
            <span>⚡ Reroute All Saw 2 Sheets to Saw 1</span>
          </button>
        </div>
      </section>

      {/* Navigation Tabs */}
      <nav className="bg-[#0f172a] border-b border-[#223354] px-4 flex items-center gap-1 overflow-x-auto scrollbar-none">
        <button
          onClick={() => setActiveTab('dispatch')}
          className={`px-4 py-2.5 font-bold text-xs sm:text-sm flex items-center gap-2 border-b-2 whitespace-nowrap transition cursor-pointer ${
            activeTab === 'dispatch'
              ? 'border-sky-400 text-sky-400 bg-[#131d33]'
              : 'border-transparent text-slate-400 hover:text-white'
          }`}
        >
          <Layers className="w-4 h-4" />
          <span>📋 Table Saw Cutlist &amp; Rip Cuts</span>
        </button>

        <button
          onClick={() => setActiveTab('villa-readiness')}
          className={`px-4 py-2.5 font-bold text-xs sm:text-sm flex items-center gap-2 border-b-2 whitespace-nowrap transition cursor-pointer ${
            activeTab === 'villa-readiness'
              ? 'border-emerald-400 text-emerald-300 bg-[#131d33]'
              : 'border-transparent text-slate-400 hover:text-white'
          }`}
        >
          <CheckCircle2 className="w-4 h-4 text-emerald-400" />
          <span>🏠 Villa Set Readiness</span>
        </button>

        <button
          onClick={() => setActiveTab('dimension-groups')}
          className={`px-4 py-2.5 font-bold text-xs sm:text-sm flex items-center gap-2 border-b-2 whitespace-nowrap transition cursor-pointer ${
            activeTab === 'dimension-groups'
              ? 'border-emerald-400 text-emerald-300 bg-[#131d33]'
              : 'border-transparent text-slate-400 hover:text-white'
          }`}
        >
          <Ruler className="w-4 h-4 text-emerald-400" />
          <span>📐 Exact Sizes &amp; Remaining Cuts (Ignore Names)</span>
        </button>

        <button
          onClick={() => setActiveTab('parts-explorer')}
          className={`px-4 py-2.5 font-bold text-xs sm:text-sm flex items-center gap-2 border-b-2 whitespace-nowrap transition cursor-pointer ${
            activeTab === 'parts-explorer'
              ? 'border-amber-400 text-amber-400 bg-[#131d33]'
              : 'border-transparent text-slate-400 hover:text-white'
          }`}
        >
          <Box className="w-4 h-4 text-amber-400" />
          <span>🧩 Component Parts Explorer (Per PDF Set)</span>
        </button>

        <button
          onClick={() => setActiveTab('pdf-matrix')}
          className={`px-4 py-2.5 font-bold text-xs sm:text-sm flex items-center gap-2 border-b-2 whitespace-nowrap transition cursor-pointer ${
            activeTab === 'pdf-matrix'
              ? 'border-sky-400 text-sky-400 bg-[#131d33]'
              : 'border-transparent text-slate-400 hover:text-white'
          }`}
        >
          <FileText className="w-4 h-4" />
          <span>📑 PDF Completion Matrix (PDF 1–5)</span>
        </button>

        <button
          onClick={() => setActiveTab('time-forecast')}
          className={`px-4 py-2.5 font-bold text-xs sm:text-sm flex items-center gap-2 border-b-2 whitespace-nowrap transition cursor-pointer ${
            activeTab === 'time-forecast'
              ? 'border-sky-400 text-sky-400 bg-[#131d33]'
              : 'border-transparent text-slate-400 hover:text-white'
          }`}
        >
          <Clock className="w-4 h-4" />
          <span>⏱️ Shift, Cycle Time &amp; Live ETA</span>
        </button>

        <button
          onClick={() => setActiveTab('overtime')}
          className={`px-4 py-2.5 font-bold text-xs sm:text-sm flex items-center gap-2 border-b-2 whitespace-nowrap transition cursor-pointer ${
            activeTab === 'overtime'
              ? 'border-sky-400 text-sky-400 bg-[#131d33]'
              : 'border-transparent text-slate-400 hover:text-white'
          }`}
        >
          <Moon className="w-4 h-4 text-amber-400" />
          <span>🌙 Overtime Log (Next Morning)</span>
        </button>

        <button
          onClick={() => setActiveTab('database')}
          className={`px-4 py-2.5 font-bold text-xs sm:text-sm flex items-center gap-2 border-b-2 whitespace-nowrap transition cursor-pointer ${
            activeTab === 'database'
              ? 'border-emerald-400 text-emerald-400 bg-[#131d33]'
              : 'border-transparent text-slate-400 hover:text-white'
          }`}
        >
          <Database className="w-4 h-4 text-emerald-400" />
          <span>📁 Excel / CSV Database Manager</span>
        </button>

        <button
          onClick={() => setActiveTab('gemma')}
          className={`px-4 py-2.5 font-bold text-xs sm:text-sm flex items-center gap-2 border-b-2 whitespace-nowrap transition cursor-pointer ${
            activeTab === 'gemma'
              ? 'border-sky-400 text-sky-400 bg-[#131d33]'
              : 'border-transparent text-slate-400 hover:text-white'
          }`}
        >
          <Bot className="w-4 h-4 text-sky-400" />
          <span>🤖 Gemma 4 Director</span>
        </button>
      </nav>

      {/* Main Content Area */}
      <main className="flex-1 p-4 sm:p-6 max-w-[1920px] w-full mx-auto space-y-4">
        {activeTab === 'villa-readiness' && (
          <div className="space-y-4">
            <div className="bg-[#131d33] border border-[#223354] rounded-2xl p-4 sm:p-5 shadow">
              <h2 className="text-base sm:text-lg font-black text-white flex items-center gap-2">
                <CheckCircle2 className="w-5 h-5 text-emerald-400" />
                <span>40-Villa Complete Set Readiness</span>
              </h2>
              <p className="text-xs text-slate-400 mt-1 max-w-4xl leading-relaxed">
                Each complete set requires all four groups: Bedroom 03, Bedroom 02, Bedroom 01 &amp; Drawers,
                and Master Bedroom + Side Cabinet (PDF 1 and PDF 4 combined). Counts only parts marked
                <strong> STANDARD 40</strong> that are not packing or skirting. The cutlist does not assign
                these shared sets to specific Villa IDs.
              </p>
            </div>

            <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
              <div className="bg-[#131d33] border border-[#223354] rounded-xl p-3.5">
                <div className="text-[10px] font-bold text-slate-400 uppercase tracking-wider">
                  Verified Ready for Assembly
                </div>
                <div className={`text-2xl font-black mt-0.5 ${villaSetProgress.verifiedReadySets === null ? 'text-amber-300' : 'text-emerald-400'}`}>
                  {villaSetProgress.verifiedReadySets === null
                    ? 'BLOCKED'
                    : `${villaSetProgress.verifiedReadySets} / ${STANDARD_VILLA_COUNT}`}
                </div>
                <div className="text-[11px] text-slate-400 mt-0.5">
                  {villaSetProgress.verifiedReadySets === null
                    ? villaSetProgress.partGroups.length === 0
                      ? 'No eligible shared parts found in the cutlist'
                      : `${villaSetProgress.unmappedGroups.length} uneven part types need Villa ID mapping`
                    : 'All included part quantities map evenly to 40 sets'}
                </div>
              </div>

              <div className="bg-[#131d33] border border-[#223354] rounded-xl p-3.5">
                <div className="text-[10px] font-bold text-slate-400 uppercase tracking-wider">
                  Complete Sets Supported by All 4 Groups
                </div>
                <div className="text-2xl font-black text-sky-300 mt-0.5">
                  {villaSetProgress.potentialSets} / {STANDARD_VILLA_COUNT}
                </div>
                <div className="text-[11px] text-slate-400 mt-0.5">
                  Limited by the least-complete required group
                </div>
              </div>

              <div className="bg-[#131d33] border border-[#223354] rounded-xl p-3.5">
                <div className="text-[10px] font-bold text-slate-400 uppercase tracking-wider">
                  Included Shared Parts Cut
                </div>
                <div className="text-2xl font-black text-white mt-0.5">
                  {villaSetProgress.cutPieces.toLocaleString()} / {villaSetProgress.totalPieces.toLocaleString()}
                </div>
                <div className="text-[11px] text-slate-400 mt-0.5">
                  Pieces cut across all four required groups
                </div>
              </div>
            </div>

            <div className="grid grid-cols-1 sm:grid-cols-2 xl:grid-cols-4 gap-3">
              {villaSetProgress.villaGroups.map((group) => (
                <section key={group.id} className="bg-[#131d33] border border-[#223354] rounded-xl p-3.5">
                  <div className="text-[10px] font-bold text-slate-400 uppercase tracking-wider">
                    Required Group
                  </div>
                  <h3 className="text-sm font-extrabold text-white mt-1">{group.label}</h3>
                  <div className="text-xl font-black text-sky-300 mt-2">
                    {group.potentialSets} / {STANDARD_VILLA_COUNT} sets supported
                  </div>
                  <div className="text-[11px] text-slate-400 mt-1">
                    {group.cutPieces.toLocaleString()} / {group.totalPieces.toLocaleString()} pieces cut
                    {' · '}{group.partGroups.length} part types
                  </div>
                  {group.unmappedGroups.length > 0 && (
                    <div className="text-[10px] text-amber-300 font-bold mt-2">
                      {group.unmappedGroups.length} uneven part type(s) block verification
                    </div>
                  )}
                </section>
              ))}
            </div>

            {villaSetProgress.unmappedGroups.length > 0 && (
              <div className="bg-amber-950/30 border border-amber-700/60 rounded-xl p-4">
                <h3 className="text-sm font-extrabold text-amber-200">
                  Assembly-ready count is blocked by unmapped quantities
                </h3>
                <p className="text-xs text-amber-100/80 mt-1">
                  These part quantities do not divide evenly into 40 sets. They are not rounded or ignored;
                  provide their Villa ID allocations before any set can be verified as ready.
                </p>
                <div className="overflow-x-auto mt-3">
                  <table className="w-full text-left text-xs text-slate-200">
                    <thead className="text-amber-100/80 uppercase text-[10px] tracking-wider border-b border-amber-800/60">
                      <tr>
                        <th className="py-2 pr-3">Required group</th>
                        <th className="py-2 pr-3">Part</th>
                        <th className="py-2 px-3">Required</th>
                        <th className="py-2 px-3">Cut</th>
                        <th className="py-2 px-3">Remaining</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-amber-900/40">
                      {villaSetProgress.unmappedGroups.map((group) => (
                        <tr key={group.id}>
                          <td className="py-2 pr-3 text-slate-300">{group.groupLabel}</td>
                          <td className="py-2 pr-3 font-medium">{group.name}</td>
                          <td className="py-2 px-3 font-mono">{group.totalQty}</td>
                          <td className="py-2 px-3 font-mono text-emerald-300">{group.cutQty}</td>
                          <td className="py-2 px-3 font-mono text-rose-300">{group.totalQty - group.cutQty}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            )}

            {villaSetProgress.partGroups.length === 0 && (
              <div className="bg-amber-950/30 border border-amber-700/60 rounded-xl p-4">
                <h3 className="text-sm font-extrabold text-amber-200">No eligible shared parts found</h3>
                <p className="text-xs text-amber-100/80 mt-1">
                  Readiness cannot be verified until the cutlist contains eligible STANDARD 40 parts.
                </p>
              </div>
            )}

            <div className="grid grid-cols-1 lg:grid-cols-3 gap-4">
              <section className="lg:col-span-2 bg-[#131d33] border border-[#223354] rounded-xl p-4">
                <h3 className="text-sm font-extrabold text-white mb-3">Part Progress Within Each Required Group</h3>
                <div className="overflow-x-auto max-h-[560px] overflow-y-auto">
                  <table className="w-full text-left text-xs text-slate-200">
                    <thead className="bg-[#0b1325] text-slate-300 font-bold uppercase text-[10px] tracking-wider sticky top-0">
                      <tr>
                        <th className="py-2.5 px-3">Required group</th>
                        <th className="py-2.5 px-3">Part</th>
                        <th className="py-2.5 px-3">Required</th>
                        <th className="py-2.5 px-3">Cut</th>
                        <th className="py-2.5 px-3">Sets Supported</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-[#223354]/50">
                      {villaSetProgress.partGroups.map((group) => (
                        <tr key={group.id}>
                          <td className="py-2.5 px-3 text-slate-300">{group.groupLabel}</td>
                          <td className="py-2.5 px-3">
                            <div className="font-semibold text-slate-100">{group.name}</div>
                            <div className="text-[10px] text-slate-500 font-mono">{group.code}</div>
                          </td>
                          <td className="py-2.5 px-3 font-mono">{group.totalQty}</td>
                          <td className="py-2.5 px-3 font-mono">
                            <span className="text-emerald-300">{group.cutQty}</span>
                            <span className="text-slate-500"> / {group.totalQty}</span>
                          </td>
                          <td className="py-2.5 px-3">
                            {group.setsSupported === null ? (
                              <span className="text-amber-300 font-bold">Needs Villa mapping</span>
                            ) : (
                              <span className="font-mono">{group.setsSupported} / {STANDARD_VILLA_COUNT}</span>
                            )}
                          </td>
                        </tr>
                      ))}
                      {villaSetProgress.partGroups.length === 0 && (
                        <tr>
                          <td colSpan={5} className="py-8 text-center text-slate-400">
                            No eligible STANDARD 40 parts found in the current cutlist.
                          </td>
                        </tr>
                      )}
                    </tbody>
                  </table>
                </div>
              </section>

              <section className="bg-[#131d33] border border-[#223354] rounded-xl p-4">
                <h3 className="text-sm font-extrabold text-white">Teams on Completed Cut Sheets</h3>
                <p className="text-[11px] text-slate-400 mt-1 mb-3">
                  Counts eligible pieces on CUT sheets by assigned team. It does not identify a single team as owner of a complete shared set.
                </p>
                <div className="space-y-2">
                  {villaSetProgress.teams.map((team) => (
                    <div key={team.team} className="flex items-center justify-between gap-3 border-b border-[#223354]/50 pb-2 text-xs">
                      <span className="font-semibold text-slate-200">{team.team}</span>
                      <span className="font-mono text-emerald-300">{team.piecesCut.toLocaleString()} pcs cut</span>
                    </div>
                  ))}
                  {villaSetProgress.teams.length === 0 && (
                    <div className="text-xs text-slate-400">No eligible parts have been cut yet.</div>
                  )}
                </div>
              </section>
            </div>
          </div>
        )}

        {/* ============================================================== */}
        {/* TAB 1: TABLE SAW CUTLIST & RIP CUTS                            */}
        {/* ============================================================== */}
        {activeTab === 'dispatch' && (
          <div className="space-y-4">
            {/* KPI Row */}
            <div className="grid grid-cols-2 sm:grid-cols-5 gap-3">
              <div className="bg-[#131d33] border border-[#223354] rounded-xl p-3.5">
                <div className="text-[10px] font-bold text-slate-400 uppercase tracking-wider">
                  Total Cutlist Boards
                </div>
                <div className="text-2xl font-black text-sky-400 mt-0.5">
                  {totalSheetsVolume}
                </div>
                <div className="text-[11px] text-slate-400 flex items-center gap-1.5 flex-wrap mt-0.5">
                  <span className="text-sky-300 font-bold">18mm: {totalSheetsVolume18mm}</span>
                  <span>•</span>
                  <span className="text-amber-300 font-bold">8mm: {totalSheetsVolume8mm}</span>
                </div>
              </div>

              <div
                onClick={() => {
                  setStatusFilter(statusFilter === 'STALE_QUEUED' ? 'ALL' : 'STALE_QUEUED');
                  setCurrentPage(1);
                }}
                className={`bg-[#131d33] border rounded-xl p-3.5 cursor-pointer transition ${
                  staleQueuedSheetsVolume > 0
                    ? 'border-rose-500/60 hover:bg-rose-950/20 shadow-md shadow-rose-950/20'
                    : 'border-[#223354]'
                }`}
                title="Click to filter boards in QUEUED state for >24 hours"
              >
                <div className="text-[10px] font-bold text-rose-400 uppercase tracking-wider flex items-center justify-between">
                  <span>Stale Stock (&gt;24h)</span>
                  {staleQueuedSheetsVolume > 0 && <span className="animate-ping w-2 h-2 rounded-full bg-rose-500" />}
                </div>
                <div className="text-2xl font-black text-rose-400 mt-0.5 flex items-baseline gap-1.5">
                  <span>{staleQueuedSheetsVolume}</span>
                  <span className="text-xs font-semibold text-slate-400">boards</span>
                </div>
                <div className="text-[11px] text-rose-300/80 flex items-center gap-1 font-semibold">
                  <AlertTriangle className="w-3 h-3 text-rose-400" />
                  <span>{statusFilter === 'STALE_QUEUED' ? '✓ Filtering Stale' : 'Click to prioritize'}</span>
                </div>
              </div>

              <div className="bg-[#131d33] border border-[#223354] rounded-xl p-3.5">
                <div className="text-[10px] font-bold text-emerald-400 uppercase tracking-wider">
                  Saw 1 (Team A) Cut
                </div>
                <div className="text-2xl font-black text-emerald-400 mt-0.5">
                  {saw1.cut_sheets} <span className="text-sm font-semibold text-slate-400">/ {saw1.assigned_sheets}</span>
                </div>
                <div className="text-[11px] text-slate-400">Primary Saw Active</div>
              </div>

              <div className={`bg-[#131d33] border rounded-xl p-3.5 ${saw2.status === 'IN_REPAIR' ? 'border-rose-900/60' : 'border-amber-900/60'}`}>
                <div className="text-[10px] font-bold text-amber-400 uppercase tracking-wider">
                  Saw 2 Status
                </div>
                <div className="text-xl sm:text-2xl font-black mt-0.5 text-white">
                  {saw2.status === 'IN_REPAIR' ? '🛠️ UNDER REPAIR' : `${saw2.cut_sheets} / ${saw2.assigned_sheets}`}
                </div>
                <div className="text-[11px] text-slate-400">
                  {saw2.assigned_sheets - saw2.cut_sheets} Sheets Pending
                </div>
              </div>

              <div className="bg-[#131d33] border border-[#223354] rounded-xl p-3.5">
                <div className="text-[10px] font-bold text-purple-400 uppercase tracking-wider">
                  Overall Cut %
                </div>
                <div className="text-2xl font-black text-purple-400 mt-0.5">
                  {cutPercent}%
                </div>
                <div className="text-[11px] text-slate-400">
                  {totalCutVolume} Completed boards
                </div>
              </div>
            </div>

            {/* Industrial Bill of Materials & Rough Summary Banner */}
            {serverState?.material_summary && serverState.material_summary.length > 0 && (
              <div className="bg-[#0b1325] border border-[#223354] rounded-xl p-3.5 shadow-sm">
                <div className="flex items-center justify-between flex-wrap gap-2 mb-2.5">
                  <div className="flex items-center gap-2">
                    <Box className="w-4 h-4 text-sky-400" />
                    <span className="text-xs font-black text-slate-200 tracking-wide uppercase">
                      OpenCutList Factory Material &amp; Rough Bill of Materials (BOM)
                    </span>
                  </div>
                  <span className="text-[11px] text-slate-400 font-mono">
                    Industry Standard: 2800×2070mm Panels · 4.4mm Saw Kerf · Standard Guillotine Trims
                  </span>
                </div>

                <div className="grid grid-cols-1 md:grid-cols-3 gap-2.5">
                  {serverState.material_summary.map((m, idx) => (
                    <div
                      key={idx}
                      className="bg-[#131d33] border border-[#223354] rounded-lg p-2.5 flex items-center justify-between"
                    >
                      <div>
                        <div className="text-xs font-bold text-white flex items-center gap-1.5">
                          <span
                            className={`w-2.5 h-2.5 rounded-full ${
                              m.material.includes('8mm')
                                ? 'bg-amber-400'
                                : m.material.includes('18mm')
                                ? 'bg-sky-400'
                                : 'bg-rose-400'
                            }`}
                          />
                          <span>{m.material}</span>
                        </div>
                        <div className="text-[11px] text-slate-400 mt-0.5">
                          <strong className="text-slate-200 font-bold">{m.parts_qty.toLocaleString()}</strong> component parts
                        </div>
                      </div>

                      <div className="text-right text-[11px] font-mono">
                        <div className="text-emerald-400 font-bold">
                          Σ Area: {m.rough_area}
                        </div>
                        <div className="text-slate-400 text-[10px]">
                          Σ Len: {m.rough_length}
                        </div>
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            )}

            {/* Filter and Search Bar */}
            <div className="bg-[#131d33] border border-[#223354] rounded-xl p-3 flex flex-wrap items-center justify-between gap-3 shadow">
              <div className="flex items-center gap-2 flex-wrap flex-1">
                {/* Search */}
                <div className="relative min-w-[200px] flex-1 max-w-xs">
                  <Search className="w-3.5 h-3.5 text-slate-400 absolute left-2.5 top-1/2 -translate-y-1/2" />
                  <input
                    type="text"
                    placeholder="Search sheets, parts, codes..."
                    value={searchQuery}
                    onChange={(e) => {
                      setSearchQuery(e.target.value);
                      setCurrentPage(1);
                    }}
                    className="w-full bg-[#0b1325] border border-[#223354] rounded-lg pl-8 pr-3 py-1.5 text-xs text-white placeholder-slate-500 focus:outline-none focus:border-sky-500"
                  />
                </div>

                {/* Material Filter */}
                <select
                  value={materialFilter}
                  onChange={(e) => {
                    setMaterialFilter(e.target.value as 'ALL' | '18mm' | '8mm');
                    setCurrentPage(1);
                  }}
                  className="bg-[#0b1325] border border-[#223354] text-xs font-bold text-sky-400 rounded-lg px-2.5 py-1.5 focus:outline-none focus:border-sky-500"
                >
                  <option value="ALL">All Materials (18mm &amp; 8mm)</option>
                  <option value="18mm">18mm MDF Core ({totalSheetsVolume18mm} bds)</option>
                  <option value="8mm">8mm Back Panels ({totalSheetsVolume8mm} bds)</option>
                </select>

                {/* PDF Filter */}
                <select
                  value={pdfFilter}
                  onChange={(e) => {
                    setPdfFilter(e.target.value);
                    setCurrentPage(1);
                  }}
                  className="bg-[#0b1325] border border-[#223354] text-xs font-bold text-white rounded-lg px-2.5 py-1.5 focus:outline-none focus:border-sky-500"
                >
                  <option value="ALL">All Drawing Sets (PDF 1–5)</option>
                  <option value="PDF 1">PDF 1 (SCMB)</option>
                  <option value="PDF 2">PDF 2 (CB-03-SF)</option>
                  <option value="PDF 3">PDF 3 (CB-02-SF)</option>
                  <option value="PDF 4">PDF 4 (SVMB / SCMB)</option>
                  <option value="PDF 5">PDF 5 (CB-01 / Drawers)</option>
                </select>

                {/* Saw Filter */}
                <select
                  value={sawFilter}
                  onChange={(e) => {
                    setSawFilter(e.target.value);
                    setCurrentPage(1);
                  }}
                  className="bg-[#0b1325] border border-[#223354] text-xs font-bold text-white rounded-lg px-2.5 py-1.5 focus:outline-none focus:border-sky-500"
                >
                  <option value="ALL">All Saws</option>
                  <option value="Saw 1">Assigned: Saw 1 (Team A)</option>
                  <option value="Saw 2">Assigned: Saw 2 (Team B)</option>
                </select>

                {/* Status Filter */}
                <select
                  value={statusFilter}
                  onChange={(e) => {
                    setStatusFilter(e.target.value);
                    setCurrentPage(1);
                  }}
                  className="bg-[#0b1325] border border-[#223354] text-xs font-bold text-white rounded-lg px-2.5 py-1.5 focus:outline-none focus:border-sky-500"
                >
                  <option value="ALL">All Statuses</option>
                  <option value="STALE_QUEUED">⚠️ Stale Stock (&gt;24h Queued) ({staleQueuedSheetsVolume})</option>
                  <option value="QUEUED">Queued</option>
                  <option value="ASSIGNED">Assigned</option>
                  <option value="CUTTING">Cutting Now</option>
                  <option value="CUT">Cut Completed</option>
                </select>

                {/* Sort Boards / Sheets */}
                <select
                  value={sheetSortBy}
                  onChange={(e) => {
                    setSheetSortBy(e.target.value as any);
                    setCurrentPage(1);
                  }}
                  className="bg-[#0b1325] border border-sky-500/40 text-xs font-bold text-sky-300 rounded-lg px-2.5 py-1.5 focus:outline-none focus:border-sky-400"
                  title="Sort board/sheet items"
                >
                  <option value="NAME_ASC">🏷️ Sort Boards: Item Name (A–Z)</option>
                  <option value="NAME_DESC">🏷️ Sort Boards: Item Name (Z–A)</option>
                  <option value="PDF">📄 Sort Boards: PDF Drawing Set</option>
                  <option value="STATUS">🚦 Sort Boards: Status (Queued → Cut)</option>
                  <option value="PARTS_COUNT">📦 Sort Boards: Part Count</option>
                  <option value="DEFAULT">📋 Original Cutlist Order</option>
                </select>

                {/* Sort Component Parts inside Board Layout */}
                <select
                  value={partSortBy}
                  onChange={(e) => setPartSortBy(e.target.value as any)}
                  className="bg-[#0b1325] border border-amber-500/40 text-xs font-bold text-amber-300 rounded-lg px-2.5 py-1.5 focus:outline-none focus:border-amber-400"
                  title="Sort component parts inside each board layout"
                >
                  <option value="NAME">🏷️ Sort Parts: By Item Name (A–Z)</option>
                  <option value="CODE">🔤 Sort Parts: By Code (A–Z)</option>
                  <option value="PDF">🧩 Sort Parts: By PDF Set (PDF 1–5)</option>
                  <option value="CAB">🗄️ Sort Parts: By Cabinet Group</option>
                  <option value="LENGTH">📏 Sort Parts: By Length (Longest First)</option>
                  <option value="QTY">🔢 Sort Parts: By Quantity</option>
                </select>

                {/* Page Size */}
                <select
                  value={pageSize}
                  onChange={(e) => {
                    setPageSize(Number(e.target.value));
                    setCurrentPage(1);
                  }}
                  className="bg-[#0b1325] border border-[#223354] text-xs font-bold text-slate-300 rounded-lg px-2 py-1.5 focus:outline-none"
                >
                  <option value={25}>Show 25</option>
                  <option value={50}>Show 50</option>
                  <option value={100}>Show 100</option>
                  <option value={9999}>Show All ({filteredSheets.length})</option>
                </select>
              </div>

              {/* Multi-selection Buttons */}
              <div className="flex items-center gap-2">
                <button
                  onClick={() => handleSelectAll(true)}
                  className="px-2.5 py-1.5 bg-[#0b1325] hover:bg-[#1c2a4a] text-xs font-bold text-slate-300 rounded-lg border border-[#223354] transition"
                >
                  Check All Filtered ({filteredSheets.length})
                </button>
                <button
                  onClick={() => handleSelectAll(false)}
                  className="px-2.5 py-1.5 bg-[#0b1325] hover:bg-[#1c2a4a] text-xs font-bold text-slate-300 rounded-lg border border-[#223354] transition"
                >
                  Uncheck All
                </button>
              </div>
            </div>

            {/* Table */}
            <div className="bg-[#10172b] border border-[#223354] rounded-2xl overflow-hidden shadow-2xl">
              <div className="overflow-x-auto max-h-[calc(100vh-340px)] overflow-y-auto">
                <table className="w-full text-left text-xs text-slate-200 border-collapse table-auto">
                  <thead className="bg-[#0b1325] text-slate-300 font-black uppercase text-[11px] tracking-wider sticky top-0 z-20 border-b-2 border-[#223354] shadow-md">
                    <tr>
                      <th className="py-3 px-3 w-10 text-center bg-[#0b1325]">
                        <input
                          type="checkbox"
                          checked={
                            paginatedSheets.length > 0 &&
                            paginatedSheets.every((s) => selectedSheetIds.has(s.sheet_id))
                          }
                          onChange={(e) => {
                            const checked = e.target.checked;
                            paginatedSheets.forEach((s) => toggleSelect(s.sheet_id, checked));
                          }}
                          className="rounded cursor-pointer accent-sky-500"
                        />
                      </th>
                      <th className="py-3.5 px-3 min-w-[180px] bg-[#0b1325]">Sheet ID / Range</th>
                      <th className="py-3.5 px-3 min-w-[130px] bg-[#0b1325]">Diagram / Specs</th>
                      <th className="py-3.5 px-3 min-w-[120px] bg-[#0b1325]">Drawing Set</th>
                      <th className="py-3.5 px-3 min-w-[210px] bg-[#0b1325]">Primary Table Saw Rip Cuts</th>
                      <th className="py-3.5 px-3 min-w-[340px] bg-[#0b1325]">Parts on This Sheet</th>
                      <th className="py-3.5 px-3 min-w-[140px] bg-[#0b1325]">Assigned Saw</th>
                      <th className="py-3.5 px-3 min-w-[140px] bg-[#0b1325]">Status / Duration</th>
                      <th className="py-3.5 px-3 min-w-[210px] text-right bg-[#0b1325]">Action</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-[#223354]/60 font-medium">
                    {paginatedSheets.map((s) => {
                      const isChecked = selectedSheetIds.has(s.sheet_id);
                      return (
                        <tr
                          key={s.sheet_id}
                          className={`hover:bg-[#1a2642] transition-colors ${
                            isChecked ? 'bg-sky-950/30' : 'even:bg-[#0c1424]/40'
                          }`}
                        >
                          <td className="py-3.5 px-3 text-center align-top pt-4">
                            <input
                              type="checkbox"
                              checked={isChecked}
                              onChange={(e) => toggleSelect(s.sheet_id, e.target.checked)}
                              className="rounded cursor-pointer accent-sky-500"
                            />
                          </td>

                          <td className="py-3.5 px-3 align-top">
                            <div className="flex items-center gap-1.5 flex-wrap">
                              <strong className="text-sky-300 text-sm font-black tracking-tight">
                                {s.sheet_title}
                              </strong>
                              {s.thickness === 8 ? (
                                <span className="bg-amber-500/20 text-amber-300 border border-amber-500/50 text-[9.5px] font-black px-1.5 py-0.5 rounded shadow-sm">
                                  8MM BACK
                                </span>
                              ) : (
                                <span className="bg-sky-500/20 text-sky-300 border border-sky-500/50 text-[9.5px] font-black px-1.5 py-0.5 rounded shadow-sm">
                                  18MM CORE
                                </span>
                              )}
                              {s.status === 'QUEUED' && s.queued_at && Math.floor(Date.now() / 1000) - s.queued_at >= 24 * 3600 && (
                                <span
                                  className="stale-stock-badge cursor-help"
                                  title={`STALE STOCK WARNING: This board has been in QUEUED state for ${Math.round((Math.floor(Date.now() / 1000) - s.queued_at) / 3600)} hours (>24h). Prioritize cutting!`}
                                >
                                  <AlertTriangle className="w-3 h-3 text-rose-400 shrink-0" />
                                  <span>QUEUED &gt;24H</span>
                                </span>
                              )}
                            </div>
                            <div className="mt-1 flex items-center gap-2">
                              <span className="text-sky-300 font-bold text-[11px] bg-sky-950/40 border border-sky-800/40 px-1.5 py-0.5 rounded">
                                1 Full Board (2800×2070mm)
                              </span>
                            </div>
                            <div className="text-[11px] text-slate-400 mt-1 line-clamp-2 leading-relaxed" title={s.material}>
                              {s.material}
                            </div>
                          </td>

                          <td className="py-3.5 px-3 align-top">
                            <button
                              onClick={() => setDiagramModalSheet(s)}
                              className="group block relative w-24 h-16 bg-[#080d1a] border border-[#223354] rounded-lg overflow-hidden hover:border-sky-400 transition shadow-inner"
                              title="Click to view full-resolution cutting diagram & sequence"
                            >
                              {s.diagram_url ? (
                                <img
                                  src={s.diagram_url}
                                  alt={s.sheet_title}
                                  className="w-full h-full object-contain p-1 opacity-90 group-hover:opacity-100 transition"
                                />
                              ) : (
                                <div className="w-full h-full flex items-center justify-center text-[10px] text-slate-500">
                                  No SVG
                                </div>
                              )}
                              <span className="absolute inset-0 bg-black/60 opacity-0 group-hover:opacity-100 flex items-center justify-center text-white text-[11px] font-black transition">
                                <Eye className="w-3.5 h-3.5 mr-1" /> View
                              </span>
                            </button>
                            <div className="text-[10.5px] text-slate-400 mt-1.5 space-y-0.5 font-mono">
                              <div className="flex items-center gap-1 text-emerald-400 font-bold" title="Total length of saw blade passes">
                                <Scissors className="w-3 h-3 text-emerald-400" />
                                <span>{s.cut_length || '28.0 m'}</span>
                              </div>
                              <div className="text-[10px] text-slate-400 truncate max-w-[100px]" title={s.grain || 'Grain follows length'}>
                                {s.grain || 'Grain: L'}
                              </div>
                            </div>
                          </td>

                          <td className="py-3.5 px-3 align-top">
                            <div className="flex flex-wrap gap-1.5">
                              {s.pdf_sets.map((pdf) => (
                                <span
                                  key={pdf}
                                  className={`pdf-badge ${pdf.toLowerCase().replace(' ', '-')}`}
                                >
                                  {pdf}
                                </span>
                              ))}
                            </div>
                          </td>

                          <td className="py-3.5 px-3 align-top">
                            <div className="flex flex-wrap gap-1.5">
                              {s.rip_cuts && s.rip_cuts.length > 0 ? (
                                s.rip_cuts.map((rc, idx) => (
                                  <span key={idx} className="rip-cut-tag">
                                    Cut {idx + 1}: {rc}
                                  </span>
                                ))
                              ) : (
                                <span className="text-slate-500 text-xs italic">
                                  Follow Diagram
                                </span>
                              )}
                            </div>
                          </td>

                          <td className="py-3.5 px-3 align-top">
                            <div className="space-y-1.5 max-h-56 overflow-y-auto pr-1">
                              {getSortedParts(s.parts, partSortBy).map((p, idx) => (
                                <div key={idx} className="text-xs leading-snug bg-[#080d1a]/85 border border-[#223354]/70 rounded-lg p-2 hover:border-sky-500/50 transition-colors shadow-sm">
                                  <div className="flex items-center gap-1.5 flex-wrap">
                                    <span className={`pdf-badge text-[9.5px] py-0.5 px-1.5 ${p.pdf_set.toLowerCase().replace(' ', '-')}`}>
                                      {p.pdf_set}
                                    </span>
                                    <strong className="text-rose-400 font-mono font-black text-xs">
                                      [{p.code}]
                                    </strong>
                                    <strong className="text-white text-xs">{p.name}</strong>
                                    <span className="text-sky-400 font-black text-xs">
                                      — {p.qty_per_sheet} pcs
                                    </span>
                                    <span className="text-emerald-300 font-mono text-[11px] bg-emerald-950/40 px-1.5 py-0.5 rounded border border-emerald-800/40">
                                      {p.l} × {p.w} mm
                                    </span>
                                    {p.villa && p.villa !== 'STANDARD 40' && (
                                      <strong className="text-amber-300 bg-amber-950/40 border border-amber-800/50 text-[10px] font-bold px-1 py-0.5 rounded">
                                        [{p.villa}]
                                      </strong>
                                    )}
                                  </div>
                                </div>
                              ))}
                            </div>
                          </td>

                          <td className="py-3.5 px-3 align-top">
                            <div className="font-black text-white text-xs">{s.assigned_saw}</div>
                            <div className="text-[11px] text-slate-400 mt-0.5">{s.assigned_team}</div>
                          </td>

                          <td className="py-3.5 px-3 align-top">
                            <div className="flex flex-col items-start gap-1.5">
                              <span className={`st st-${s.status}`}>{s.status}</span>
                              {s.status === 'QUEUED' && s.queued_at && Math.floor(Date.now() / 1000) - s.queued_at >= 24 * 3600 && (
                                <span className="text-[10px] font-black text-rose-400 bg-rose-500/15 border border-rose-500/40 px-1.5 py-0.5 rounded flex items-center gap-1 font-mono">
                                  <Clock className="w-2.5 h-2.5 text-rose-400 shrink-0" />
                                  <span>{Math.round((Math.floor(Date.now() / 1000) - s.queued_at) / 3600)}h in queue</span>
                                </span>
                              )}
                              {s.duration_min ? (
                                <div className="text-[11px] text-emerald-400 font-mono font-bold">
                                  {s.duration_min} min taken
                                </div>
                              ) : null}
                            </div>
                          </td>

                          <td className="py-3.5 px-3 align-top text-right">
                            <div className="flex items-center justify-end gap-1.5 flex-wrap">
                              <button
                                onClick={() => handleOpenEditSheet(s)}
                                className="px-2.5 py-1.5 bg-[#0b1325] hover:bg-[#1c2a4a] text-slate-300 hover:text-white border border-[#223354] rounded-lg text-xs font-bold transition inline-flex items-center gap-1 shadow"
                                title="Edit status, actual time taken, assigned saw/team"
                              >
                                <Edit className="w-3.5 h-3.5 text-amber-400" />
                                <span>Edit</span>
                              </button>

                              <button
                                onClick={() => handlePrintSingleSheet(s)}
                                className="px-2.5 py-1.5 bg-sky-600 hover:bg-sky-500 text-white rounded-lg text-xs font-bold transition inline-flex items-center gap-1 shadow"
                                title="Print 100mm x 50mm thermal labels for this sheet"
                              >
                                <Printer className="w-3.5 h-3.5" />
                                <span>Labels</span>
                              </button>

                              {s.status !== 'CUT' && (
                                <button
                                  onClick={() => handleToggleCut(s.sheet_id, s.status)}
                                  className={`px-3 py-1.5 rounded-lg text-xs font-bold text-white transition inline-flex items-center gap-1 shadow ${
                                    s.status === 'CUTTING'
                                      ? 'bg-emerald-600 hover:bg-emerald-500'
                                      : 'bg-emerald-700 hover:bg-emerald-600'
                                  }`}
                                >
                                  {s.status === 'CUTTING' ? (
                                    <>
                                      <Check className="w-3.5 h-3.5" />
                                      <span>✓ Finish Cut</span>
                                    </>
                                  ) : (
                                    <>
                                      <Play className="w-3.5 h-3.5" />
                                      <span>▶ Start Saw</span>
                                    </>
                                  )}
                                </button>
                              )}
                            </div>
                          </td>
                        </tr>
                      );
                    })}

                    {paginatedSheets.length === 0 && (
                      <tr>
                        <td colSpan={9} className="py-12 text-center text-slate-400">
                          <div className="flex flex-col items-center justify-center space-y-2">
                            <Search className="w-8 h-8 text-slate-600" />
                            <div className="font-bold text-slate-300">No sheets found</div>
                            <div className="text-xs text-slate-500">Try adjusting your filters or search query.</div>
                          </div>
                        </td>
                      </tr>
                    )}
                  </tbody>
                </table>
              </div>

              {/* Pagination Controls */}
              {totalPages > 1 && (
                <div className="p-3 bg-[#0f172a] border-t border-[#223354] flex items-center justify-between text-xs text-slate-300">
                  <div>
                    Showing {(currentPage - 1) * pageSize + 1} to{' '}
                    {Math.min(currentPage * pageSize, filteredSheets.length)} of{' '}
                    {filteredSheets.length} sheets
                  </div>

                  <div className="flex items-center gap-2">
                    <button
                      onClick={() => setCurrentPage((p) => Math.max(1, p - 1))}
                      disabled={currentPage === 1}
                      className="px-2.5 py-1 bg-[#131d33] border border-[#223354] rounded text-slate-300 hover:text-white disabled:opacity-30 disabled:cursor-not-allowed transition"
                    >
                      <ChevronLeft className="w-3.5 h-3.5 inline" /> Prev
                    </button>
                    <span className="font-bold text-sky-400">
                      Page {currentPage} of {totalPages}
                    </span>
                    <button
                      onClick={() => setCurrentPage((p) => Math.min(totalPages, p + 1))}
                      disabled={currentPage === totalPages}
                      className="px-2.5 py-1 bg-[#131d33] border border-[#223354] rounded text-slate-300 hover:text-white disabled:opacity-30 disabled:cursor-not-allowed transition"
                    >
                      Next <ChevronRight className="w-3.5 h-3.5 inline" />
                    </button>
                  </div>
                </div>
              )}
            </div>
          </div>
        )}

        {/* ============================================================== */}
        {/* TAB: EXACT SIZE & REMAINING CUTS GROUPING (IGNORE PART NAMES)  */}
        {/* ============================================================== */}
        {activeTab === 'dimension-groups' && (
          <div className="space-y-4">
            {/* Header Banner */}
            <div className="bg-[#131d33] border border-[#223354] rounded-2xl p-4 sm:p-5 shadow flex flex-wrap items-center justify-between gap-4">
              <div>
                <h2 className="text-base sm:text-lg font-black text-white flex items-center gap-2">
                  <Ruler className="w-5 h-5 text-emerald-400" />
                  <span>Exact Size &amp; Remaining Cuts Aggregator (Ignore Names)</span>
                </h2>
                <p className="text-xs text-slate-400 mt-1 max-w-3xl leading-relaxed">
                  All component parts grouped strictly by <strong className="text-amber-300">exact physical cutting dimensions (Length × Width × Thickness)</strong>, ignoring part names. Perfect for batch-cutting identical panels on table saw rip fences and immediately seeing <strong className="text-rose-400">how many pieces of each size remain to be cut</strong>.
                </p>
              </div>

              <div className="flex items-center gap-2 flex-wrap">
                <button
                  onClick={handleExportDimensionsExcel}
                  className="px-3 py-1.5 bg-emerald-600 hover:bg-emerald-500 text-white rounded-lg text-xs font-bold flex items-center gap-1.5 transition shadow"
                  title="Export exact size cutlist breakdown to Excel"
                >
                  <Download className="w-3.5 h-3.5" />
                  <span>📊 Export Sizes (.xlsx)</span>
                </button>
              </div>
            </div>

            {/* KPI Highlights */}
            <div className="grid grid-cols-2 sm:grid-cols-4 gap-3">
              <div className="bg-[#131d33] border border-[#223354] rounded-xl p-3.5">
                <div className="text-[10px] font-bold text-slate-400 uppercase tracking-wider">
                  Distinct Panel Sizes
                </div>
                <div className="text-2xl font-black text-white mt-0.5">
                  {dimKpis.totalSizes}
                </div>
                <div className="text-[11px] text-slate-400 mt-0.5 flex gap-2">
                  <span className="text-rose-400 font-bold">{dimKpis.pendingSizes} Pending</span>
                  <span>•</span>
                  <span className="text-emerald-400 font-bold">{dimKpis.completedSizes} Done</span>
                </div>
              </div>

              <div
                onClick={() => {
                  setDimStatusFilter('PENDING');
                  setDimPage(1);
                }}
                className={`border rounded-xl p-3.5 cursor-pointer transition ${
                  dimStatusFilter === 'PENDING'
                    ? 'bg-rose-950/30 border-rose-500/60 shadow-lg shadow-rose-950/20'
                    : 'bg-[#131d33] border-[#223354] hover:border-slate-500'
                }`}
                title="Click to show only sizes with remaining pieces"
              >
                <div className="text-[10px] font-bold text-rose-400 uppercase tracking-wider flex items-center justify-between">
                  <span>Remaining Pieces to Cut</span>
                  <span className="animate-pulse w-2 h-2 rounded-full bg-rose-500" />
                </div>
                <div className="text-2xl font-black text-rose-400 mt-0.5">
                  {dimKpis.remainingPieces.toLocaleString()} pcs
                </div>
                <div className="text-[11px] text-slate-400 mt-0.5">
                  Across {dimKpis.pendingSizes} distinct dimensions
                </div>
              </div>

              <div
                onClick={() => {
                  setDimStatusFilter('CUT');
                  setDimPage(1);
                }}
                className={`border rounded-xl p-3.5 cursor-pointer transition ${
                  dimStatusFilter === 'CUT'
                    ? 'bg-emerald-950/30 border-emerald-500/60 shadow-lg shadow-emerald-950/20'
                    : 'bg-[#131d33] border-[#223354] hover:border-slate-500'
                }`}
                title="Click to show fully cut sizes"
              >
                <div className="text-[10px] font-bold text-emerald-400 uppercase tracking-wider">
                  Pieces Already Cut
                </div>
                <div className="text-2xl font-black text-emerald-400 mt-0.5">
                  {dimKpis.cutPieces.toLocaleString()} pcs
                </div>
                <div className="text-[11px] text-slate-400 mt-0.5">
                  {dimKpis.completedSizes} sizes 100% completed
                </div>
              </div>

              <div className="bg-[#131d33] border border-[#223354] rounded-xl p-3.5">
                <div className="text-[10px] font-bold text-sky-400 uppercase tracking-wider flex items-center justify-between">
                  <span>Batch Completion</span>
                  <span className="font-mono">{dimKpis.pct}%</span>
                </div>
                <div className="text-2xl font-black text-sky-300 mt-0.5">
                  {dimKpis.totalPieces.toLocaleString()} pcs
                </div>
                <div className="w-full bg-[#080d1a] h-2 rounded-full overflow-hidden mt-1.5">
                  <div
                    className="bg-emerald-400 h-full transition-all duration-500"
                    style={{ width: `${dimKpis.pct}%` }}
                  />
                </div>
              </div>
            </div>

            {/* Filter and Search Bar for Exact Sizes */}
            <div className="bg-[#131d33] border border-[#223354] rounded-xl p-3 flex flex-wrap items-center justify-between gap-3 shadow">
              <div className="flex items-center gap-2 flex-wrap flex-1">
                {/* Search Dimension */}
                <div className="relative min-w-[200px] flex-1 max-w-sm">
                  <Search className="w-3.5 h-3.5 text-slate-400 absolute left-2.5 top-1/2 -translate-y-1/2" />
                  <input
                    type="text"
                    placeholder="Search length, width, thickness (e.g. 2240, 580)..."
                    value={dimSearch}
                    onChange={(e) => {
                      setDimSearch(e.target.value);
                      setDimPage(1);
                    }}
                    className="w-full bg-[#0b1325] border border-[#223354] rounded-lg pl-8 pr-3 py-1.5 text-xs text-white placeholder-slate-500 focus:outline-none focus:border-emerald-400"
                  />
                </div>

                {/* Material / Thickness Filter */}
                <select
                  value={dimThicknessFilter}
                  onChange={(e) => {
                    setDimThicknessFilter(e.target.value as any);
                    setDimPage(1);
                  }}
                  className="bg-[#0b1325] border border-[#223354] text-xs font-bold text-sky-400 rounded-lg px-2.5 py-1.5 focus:outline-none focus:border-sky-500"
                >
                  <option value="ALL">All Materials (18mm &amp; 8mm)</option>
                  <option value="18mm">18mm MDF Core Only</option>
                  <option value="8mm">8mm Back Panels Only</option>
                </select>

                {/* Status Filter */}
                <select
                  value={dimStatusFilter}
                  onChange={(e) => {
                    setDimStatusFilter(e.target.value as any);
                    setDimPage(1);
                  }}
                  className="bg-[#0b1325] border border-rose-500/40 text-xs font-bold text-rose-300 rounded-lg px-2.5 py-1.5 focus:outline-none focus:border-rose-400"
                >
                  <option value="PENDING">⚠️ Pending Remaining Pieces Only (Remaining &gt; 0)</option>
                  <option value="CUT">✅ Completed Sizes Only (100% Done)</option>
                  <option value="ALL">All Sizes (Pending &amp; Cut)</option>
                </select>

                {/* PDF Drawing Set Filter */}
                <select
                  value={dimPdfFilter}
                  onChange={(e) => {
                    setDimPdfFilter(e.target.value);
                    setDimPage(1);
                  }}
                  className="bg-[#0b1325] border border-[#223354] text-xs font-bold text-white rounded-lg px-2.5 py-1.5 focus:outline-none focus:border-sky-500"
                >
                  <option value="ALL">All Drawing Sets (PDF 1–5)</option>
                  <option value="PDF 1">PDF 1 (SCMB)</option>
                  <option value="PDF 2">PDF 2 (CB-03-SF)</option>
                  <option value="PDF 3">PDF 3 (CB-02-SF)</option>
                  <option value="PDF 4">PDF 4 (SVMB Master)</option>
                  <option value="PDF 5">PDF 5 (CB-01 / Drawers)</option>
                </select>

                {/* Sort Sizes By */}
                <select
                  value={dimSortBy}
                  onChange={(e) => {
                    setDimSortBy(e.target.value as any);
                    setDimPage(1);
                  }}
                  className="bg-[#0b1325] border border-amber-500/40 text-xs font-bold text-amber-300 rounded-lg px-2.5 py-1.5 focus:outline-none focus:border-amber-400"
                >
                  <option value="REMAINING">⚠️ Most Remaining Pieces First</option>
                  <option value="TOTAL">🔢 Total Quantity (Highest First)</option>
                  <option value="LENGTH">📏 Length (Longest First - Rip Cut Setup)</option>
                  <option value="WIDTH">📐 Width (Widest First)</option>
                  <option value="AREA">🟩 Total Surface Area (m²)</option>
                  <option value="PROGRESS_ASC">⏱️ Least Completed First (% Done)</option>
                  <option value="PROGRESS_DESC">✅ Most Completed First (% Done)</option>
                </select>

                {/* Page Size */}
                <select
                  value={dimPageSize}
                  onChange={(e) => {
                    setDimPageSize(Number(e.target.value));
                    setDimPage(1);
                  }}
                  className="bg-[#0b1325] border border-[#223354] text-xs font-bold text-slate-300 rounded-lg px-2 py-1.5 focus:outline-none"
                >
                  <option value={25}>Show 25</option>
                  <option value={50}>Show 50</option>
                  <option value={100}>Show 100</option>
                  <option value={9999}>Show All ({filteredDimGroups.length})</option>
                </select>
              </div>

              {/* Reset Filter Button */}
              {(dimSearch || dimThicknessFilter !== 'ALL' || dimStatusFilter !== 'PENDING' || dimPdfFilter !== 'ALL' || dimSortBy !== 'REMAINING') && (
                <button
                  onClick={() => {
                    setDimSearch('');
                    setDimThicknessFilter('ALL');
                    setDimStatusFilter('PENDING');
                    setDimPdfFilter('ALL');
                    setDimSortBy('REMAINING');
                    setDimPage(1);
                  }}
                  className="px-2.5 py-1.5 bg-[#0b1325] hover:bg-[#1c2a4a] text-xs font-bold text-amber-400 rounded-lg border border-amber-500/30 transition flex items-center gap-1"
                >
                  <X className="w-3.5 h-3.5" /> Reset Filters
                </button>
              )}
            </div>

            {/* Exact Sizes Master Table */}
            <div className="bg-[#10172b] border border-[#223354] rounded-2xl overflow-hidden shadow-2xl">
              <div className="overflow-x-auto max-h-[calc(100vh-360px)] overflow-y-auto">
                <table className="w-full text-left text-xs text-slate-200 border-collapse table-auto">
                  <thead className="bg-[#0b1325] text-slate-300 font-black uppercase text-[11px] tracking-wider sticky top-0 z-20 border-b-2 border-[#223354] shadow-md">
                    <tr>
                      <th className="py-3.5 px-3">Exact Physical Cut Size</th>
                      <th className="py-3.5 px-3 text-center">Remaining to Cut</th>
                      <th className="py-3.5 px-3">Cut vs Total Qty</th>
                      <th className="py-3.5 px-3">Surface Area</th>
                      <th className="py-3.5 px-3">Drawing Sets</th>
                      <th className="py-3.5 px-3">Ignored Part Names (Grouped)</th>
                      <th className="py-3.5 px-3">Remaining Sheets to Cut</th>
                      <th className="py-3.5 px-3 text-right">Details</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-[#223354]/40 font-medium">
                    {paginatedDimGroups.map((g) => {
                      const isExpanded = expandedDimKey === g.key;

                      return (
                        <React.Fragment key={g.key}>
                          <tr
                            className={`transition hover:bg-[#131d33]/80 ${
                              isExpanded ? 'bg-[#15233e]' : ''
                            }`}
                          >
                            {/* Exact Dimensions */}
                            <td className="py-3 px-3 align-top">
                              <div className="flex items-center gap-2">
                                <span className="text-base font-black text-amber-300 font-mono tracking-tight">
                                  {g.l} × {g.w} mm
                                </span>
                                <span
                                  className={`text-[10px] font-black px-1.5 py-0.5 rounded ${
                                    g.t === 8
                                      ? 'bg-amber-500/20 text-amber-300 border border-amber-500/40'
                                      : 'bg-sky-500/20 text-sky-300 border border-sky-500/40'
                                  }`}
                                >
                                  {g.t}mm
                                </span>
                              </div>
                              <div className="text-[10.5px] text-slate-400 mt-0.5 font-mono">
                                Table Saw Setup: <strong className="text-slate-300">Rip: {g.w}mm</strong> | Stop: {g.l}mm
                              </div>
                              <div className="text-[10px] text-slate-500 truncate max-w-[200px]" title={g.core}>
                                {g.core}
                              </div>
                            </td>

                            {/* Remaining Pieces */}
                            <td className="py-3 px-3 align-top text-center">
                              {g.remaining_qty > 0 ? (
                                <div className="inline-flex flex-col items-center">
                                  <span className="px-3 py-1 bg-rose-500/20 text-rose-300 border border-rose-500/50 rounded-lg font-black font-mono text-sm shadow-sm flex items-center gap-1.5">
                                    <AlertTriangle className="w-3.5 h-3.5 text-rose-400" />
                                    <span>{g.remaining_qty} pcs LEFT</span>
                                  </span>
                                  <span className="text-[10px] text-rose-400/80 font-bold mt-0.5">
                                    in {g.remaining_sheets.length} sheet layouts
                                  </span>
                                </div>
                              ) : (
                                <div className="inline-flex flex-col items-center">
                                  <span className="px-3 py-1 bg-emerald-500/20 text-emerald-300 border border-emerald-500/50 rounded-lg font-black font-mono text-sm">
                                    ✅ ALL CUT
                                  </span>
                                  <span className="text-[10px] text-emerald-400/80 font-bold mt-0.5">
                                    0 pieces remaining
                                  </span>
                                </div>
                              )}
                            </td>

                            {/* Cut vs Total Qty */}
                            <td className="py-3 px-3 align-top">
                              <div className="flex items-center justify-between text-xs font-mono font-bold">
                                <span className="text-white">
                                  {g.cut_qty} / {g.total_qty} pcs
                                </span>
                                <span className="text-emerald-400">{g.pct_cut}%</span>
                              </div>
                              <div className="w-32 bg-[#080d1a] h-2 rounded-full overflow-hidden mt-1.5 border border-[#223354]">
                                <div
                                  className={`h-full transition-all duration-300 ${
                                    g.pct_cut === 100 ? 'bg-emerald-400' : 'bg-sky-400'
                                  }`}
                                  style={{ width: `${g.pct_cut}%` }}
                                />
                              </div>
                            </td>

                            {/* Surface Area */}
                            <td className="py-3 px-3 align-top font-mono">
                              <div className="text-xs font-bold text-slate-200">
                                {g.total_area_m2} m²
                              </div>
                              <div className="text-[10px] text-slate-400">
                                ({g.area_m2.toFixed(3)} m²/pc)
                              </div>
                            </td>

                            {/* Drawing Sets */}
                            <td className="py-3 px-3 align-top">
                              <div className="flex flex-wrap gap-1">
                                {g.pdf_sets.map((pdf) => (
                                  <span
                                    key={pdf}
                                    className={`pdf-badge text-[9.5px] py-0.5 px-1.5 ${pdf.toLowerCase().replace(' ', '-')}`}
                                  >
                                    {pdf}
                                  </span>
                                ))}
                              </div>
                            </td>

                            {/* Ignored Part Names Grouped */}
                            <td className="py-3 px-3 align-top">
                              <div className="text-[10.5px] font-bold text-slate-400 flex items-center gap-1">
                                <span className="bg-[#0b1325] px-1.5 py-0.5 rounded text-sky-300 font-mono border border-[#223354]">
                                  {g.part_names.length} names grouped:
                                </span>
                              </div>
                              <div className="flex flex-wrap gap-1 mt-1 max-w-[260px]">
                                {g.part_names.slice(0, 3).map((name, i) => (
                                  <span
                                    key={i}
                                    className="bg-[#080d1a] border border-[#223354] text-slate-300 text-[10px] px-1.5 py-0.5 rounded truncate max-w-[120px]"
                                    title={name}
                                  >
                                    {name}
                                  </span>
                                ))}
                                {g.part_names.length > 3 && (
                                  <span className="text-[9.5px] text-slate-400 bg-[#080d1a] px-1.5 py-0.5 rounded border border-[#223354]">
                                    +{g.part_names.length - 3} more
                                  </span>
                                )}
                              </div>
                            </td>

                            {/* Remaining Sheets to Cut */}
                            <td className="py-3 px-3 align-top">
                              {g.remaining_sheets.length > 0 ? (
                                <div className="space-y-1 max-w-[220px]">
                                  <div className="flex flex-wrap gap-1">
                                    {g.remaining_sheets.slice(0, 4).map((rs, idx) => (
                                      <span
                                        key={idx}
                                        className="bg-rose-950/40 border border-rose-800/40 text-rose-200 text-[10px] font-mono px-1.5 py-0.5 rounded"
                                        title={`${rs.sheet_title}: ${rs.qty_on_sheet} pcs (${rs.assigned_saw})`}
                                      >
                                        {rs.sheet_id} ({rs.qty_on_sheet}p)
                                      </span>
                                    ))}
                                    {g.remaining_sheets.length > 4 && (
                                      <span className="text-[9.5px] text-rose-300 bg-rose-950/20 px-1.5 py-0.5 rounded border border-rose-800/30">
                                        +{g.remaining_sheets.length - 4} more
                                      </span>
                                    )}
                                  </div>
                                </div>
                              ) : (
                                <span className="text-emerald-400 text-xs font-bold">
                                  All sheets cut
                                </span>
                              )}
                            </td>

                            {/* Details Action */}
                            <td className="py-3 px-3 align-top text-right">
                              <button
                                onClick={() => setExpandedDimKey(isExpanded ? null : g.key)}
                                className="px-2.5 py-1 bg-[#0b1325] hover:bg-[#1c2a4a] text-slate-300 border border-[#223354] rounded text-xs font-bold inline-flex items-center gap-1 transition"
                                title="Expand sheet occurrences"
                              >
                                {isExpanded ? (
                                  <>
                                    <span>Hide</span>
                                    <ChevronUp className="w-3.5 h-3.5" />
                                  </>
                                ) : (
                                  <>
                                    <span>{g.sheets.length} Sheets</span>
                                    <ChevronDown className="w-3.5 h-3.5" />
                                  </>
                                )}
                              </button>
                            </td>
                          </tr>

                          {/* Expanded Details Row */}
                          {isExpanded && (
                            <tr className="bg-[#0b1325]/90 border-b-2 border-sky-500/40">
                              <td colSpan={8} className="p-4">
                                <div className="space-y-3">
                                  <div className="flex items-center justify-between flex-wrap gap-2">
                                    <div className="font-extrabold text-sm text-white flex items-center gap-2">
                                      <span className="text-amber-300">{g.l} × {g.w} × {g.t} mm</span>
                                      <span className="text-slate-400">— All {g.sheets.length} Board Sheet Occurrences</span>
                                    </div>
                                    <div className="text-xs font-bold text-slate-400 font-mono">
                                      Total required: <strong className="text-white">{g.total_qty}</strong> | Cut: <strong className="text-emerald-400">{g.cut_qty}</strong> | Pending: <strong className="text-rose-400">{g.remaining_qty}</strong>
                                    </div>
                                  </div>

                                  <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-2">
                                    {g.sheets.map((sh, sIdx) => {
                                      const sRec = (serverState?.sheets || []).find((x) => x.sheet_id === sh.sheet_id);

                                      return (
                                        <div
                                          key={sIdx}
                                          className={`p-2.5 rounded-lg border text-xs flex items-center justify-between ${
                                            sh.status === 'CUT'
                                              ? 'bg-emerald-950/20 border-emerald-800/40'
                                              : 'bg-[#10172b] border-[#223354]'
                                          }`}
                                        >
                                          <div>
                                            <div className="font-mono font-bold text-white flex items-center gap-1.5">
                                              <span>{sh.sheet_title}</span>
                                              <span className={`st st-${sh.status} text-[9.5px]`}>
                                                {sh.status}
                                              </span>
                                            </div>
                                            <div className="text-[10.5px] text-slate-400 mt-0.5 font-mono">
                                              Part: <span className="text-rose-300 font-bold">[{sh.part_code}]</span> {sh.part_name}
                                            </div>
                                            <div className="text-[10px] text-slate-500">
                                              Saw: {sh.assigned_saw} • {sh.pdf_set}
                                            </div>
                                          </div>

                                          <div className="text-right">
                                            <div className="font-mono font-black text-amber-300 text-sm">
                                              {sh.qty_on_sheet} pcs
                                            </div>
                                            {sRec && (
                                              <button
                                                onClick={() => setDiagramModalSheet(sRec)}
                                                className="mt-1 px-2 py-0.5 bg-sky-600/30 hover:bg-sky-600 text-sky-200 rounded text-[10px] font-bold transition inline-flex items-center gap-1"
                                              >
                                                <Eye className="w-2.5 h-2.5" /> Diagram
                                              </button>
                                            )}
                                          </div>
                                        </div>
                                      );
                                    })}
                                  </div>
                                </div>
                              </td>
                            </tr>
                          )}
                        </React.Fragment>
                      );
                    })}

                    {filteredDimGroups.length === 0 && (
                      <tr>
                        <td colSpan={8} className="py-12 text-center text-slate-400">
                          <div className="flex flex-col items-center justify-center space-y-2">
                            <Search className="w-8 h-8 text-slate-600" />
                            <div className="font-bold text-slate-300">No exact sizes found</div>
                            <div className="text-xs text-slate-500">
                              Try clearing filters or changing the status filter from &apos;Pending&apos; to &apos;All Sizes&apos;.
                            </div>
                          </div>
                        </td>
                      </tr>
                    )}
                  </tbody>
                </table>
              </div>

              {/* Pagination Controls */}
              {dimTotalPages > 1 && (
                <div className="p-3 bg-[#0f172a] border-t border-[#223354] flex items-center justify-between text-xs text-slate-300">
                  <div>
                    Showing {(dimPage - 1) * dimPageSize + 1} to{' '}
                    {Math.min(dimPage * dimPageSize, filteredDimGroups.length)} of{' '}
                    {filteredDimGroups.length} distinct panel sizes
                  </div>

                  <div className="flex items-center gap-2">
                    <button
                      onClick={() => setDimPage((p) => Math.max(1, p - 1))}
                      disabled={dimPage === 1}
                      className="px-2.5 py-1 bg-[#131d33] border border-[#223354] rounded text-slate-300 hover:text-white disabled:opacity-30 disabled:cursor-not-allowed transition"
                    >
                      <ChevronLeft className="w-3.5 h-3.5 inline" /> Prev
                    </button>
                    <span className="font-bold text-emerald-400">
                      Page {dimPage} of {dimTotalPages}
                    </span>
                    <button
                      onClick={() => setDimPage((p) => Math.min(dimTotalPages, p + 1))}
                      disabled={dimPage === dimTotalPages}
                      className="px-2.5 py-1 bg-[#131d33] border border-[#223354] rounded text-slate-300 hover:text-white disabled:opacity-30 disabled:cursor-not-allowed transition"
                    >
                      Next <ChevronRight className="w-3.5 h-3.5 inline" />
                    </button>
                  </div>
                </div>
              )}
            </div>
          </div>
        )}

        {/* ============================================================== */}
        {/* TAB: COMPONENT PARTS EXPLORER (PER PDF DRAWING SET)            */}
        {/* ============================================================== */}
        {activeTab === 'parts-explorer' && (
          <div className="space-y-4">
            {/* Header Banner */}
            <div className="bg-[#131d33] border border-[#223354] rounded-2xl p-4 sm:p-5 shadow flex flex-wrap items-center justify-between gap-4">
              <div>
                <h2 className="text-base sm:text-lg font-black text-white flex items-center gap-2">
                  <Box className="w-5 h-5 text-amber-400" />
                  <span>Component Parts Explorer (Grouped by Part Code &amp; Drawing Set)</span>
                </h2>
                <p className="text-xs text-slate-400 mt-1 max-w-3xl leading-relaxed">
                  Group all components by unique <strong className="text-rose-400">Part Code (e.g. [DA], [DG], [AV])</strong> and existing part names to immediately identify each piece, see <strong className="text-sky-300">what is needed</strong>, <strong className="text-emerald-400">how many are cut</strong>, and <strong className="text-rose-300">how many remain</strong> across all 953 board sheets.
                </p>
              </div>

              <div className="flex items-center gap-2 flex-wrap">
                {/* View Mode Toggle */}
                <div className="bg-[#0b1325] p-1 rounded-xl border border-[#223354] flex items-center gap-1">
                  <button
                    onClick={() => {
                      setPartsExplorerViewMode('GROUPED');
                      setPartsExplorerPage(1);
                    }}
                    className={`px-3 py-1.5 rounded-lg text-xs font-black transition flex items-center gap-1.5 ${
                      partsExplorerViewMode === 'GROUPED'
                        ? 'bg-amber-500 text-slate-950 shadow'
                        : 'text-slate-400 hover:text-white'
                    }`}
                  >
                    <span>🏷️ Grouped by Part Code</span>
                    <span className="bg-black/30 px-1.5 py-0.2 rounded text-[10px] font-mono">
                      {groupedParts.length}
                    </span>
                  </button>

                  <button
                    onClick={() => {
                      setPartsExplorerViewMode('FLAT');
                      setPartsExplorerPage(1);
                    }}
                    className={`px-3 py-1.5 rounded-lg text-xs font-black transition flex items-center gap-1.5 ${
                      partsExplorerViewMode === 'FLAT'
                        ? 'bg-amber-500 text-slate-950 shadow'
                        : 'text-slate-400 hover:text-white'
                    }`}
                  >
                    <span>📄 All Sheet Occurrences</span>
                    <span className="bg-black/30 px-1.5 py-0.2 rounded text-[10px] font-mono">
                      {allFlattenedParts.length.toLocaleString()}
                    </span>
                  </button>
                </div>

                <button
                  onClick={handleExportGroupedPartsExcel}
                  className="px-3 py-1.5 bg-amber-600 hover:bg-amber-500 text-slate-950 font-black rounded-lg text-xs flex items-center gap-1.5 transition shadow"
                  title="Export grouped parts summary to Excel"
                >
                  <Download className="w-3.5 h-3.5" />
                  <span>📊 Export Parts (.xlsx)</span>
                </button>
              </div>
            </div>

            {/* Quick KPI Highlights */}
            <div className="grid grid-cols-2 sm:grid-cols-4 gap-3">
              <div className="bg-[#131d33] border border-[#223354] rounded-xl p-3.5">
                <div className="text-[10px] font-bold text-slate-400 uppercase tracking-wider">
                  Unique Component Parts
                </div>
                <div className="text-2xl font-black text-white mt-0.5">
                  {partsKpis.totalUnique.toLocaleString()}
                </div>
                <div className="text-[11px] text-slate-400 mt-0.5 flex gap-2">
                  <span className="text-rose-400 font-bold">{partsKpis.pendingParts} Pending</span>
                  <span>•</span>
                  <span className="text-emerald-400 font-bold">{partsKpis.completedParts} Cut</span>
                </div>
              </div>

              <div
                onClick={() => {
                  setPartsExplorerStatusFilter(partsExplorerStatusFilter === 'PENDING' ? 'ALL' : 'PENDING');
                  setPartsExplorerPage(1);
                }}
                className={`border rounded-xl p-3.5 cursor-pointer transition ${
                  partsExplorerStatusFilter === 'PENDING'
                    ? 'bg-rose-950/30 border-rose-500/60 shadow-lg shadow-rose-950/20'
                    : 'bg-[#131d33] border-[#223354] hover:border-slate-500'
                }`}
                title="Click to toggle showing only pending parts"
              >
                <div className="text-[10px] font-bold text-rose-400 uppercase tracking-wider flex items-center justify-between">
                  <span>Remaining Pieces to Cut</span>
                  <span className="animate-pulse w-2 h-2 rounded-full bg-rose-500" />
                </div>
                <div className="text-2xl font-black text-rose-400 mt-0.5">
                  {partsKpis.totalRemaining.toLocaleString()} pcs
                </div>
                <div className="text-[11px] text-slate-400 mt-0.5">
                  Across {partsKpis.pendingParts} distinct part codes
                </div>
              </div>

              <div
                onClick={() => {
                  setPartsExplorerStatusFilter(partsExplorerStatusFilter === 'CUT' ? 'ALL' : 'CUT');
                  setPartsExplorerPage(1);
                }}
                className={`border rounded-xl p-3.5 cursor-pointer transition ${
                  partsExplorerStatusFilter === 'CUT'
                    ? 'bg-emerald-950/30 border-emerald-500/60 shadow-lg shadow-emerald-950/20'
                    : 'bg-[#131d33] border-[#223354] hover:border-slate-500'
                }`}
                title="Click to toggle showing only completed parts"
              >
                <div className="text-[10px] font-bold text-emerald-400 uppercase tracking-wider">
                  Pieces Already Cut
                </div>
                <div className="text-2xl font-black text-emerald-400 mt-0.5">
                  {partsKpis.totalCut.toLocaleString()} pcs
                </div>
                <div className="text-[11px] text-slate-400 mt-0.5">
                  {partsKpis.completedParts} parts 100% completed
                </div>
              </div>

              <div className="bg-[#131d33] border border-[#223354] rounded-xl p-3.5">
                <div className="text-[10px] font-bold text-sky-400 uppercase tracking-wider flex items-center justify-between">
                  <span>Total Project Volume</span>
                  <span className="font-mono">{partsKpis.pct}%</span>
                </div>
                <div className="text-2xl font-black text-sky-300 mt-0.5">
                  {partsKpis.totalNeeded.toLocaleString()} pcs
                </div>
                <div className="w-full bg-[#080d1a] h-2 rounded-full overflow-hidden mt-1.5">
                  <div
                    className="bg-emerald-400 h-full transition-all duration-500"
                    style={{ width: `${partsKpis.pct}%` }}
                  />
                </div>
              </div>
            </div>

            {/* Quick PDF Set Summary Cards (PDF 1 to 5) */}
            <div className="grid grid-cols-1 sm:grid-cols-5 gap-3">
              {['PDF 1', 'PDF 2', 'PDF 3', 'PDF 4', 'PDF 5'].map((pdfKey) => {
                const pdfParts = allFlattenedParts.filter((x) => x.part.pdf_set === pdfKey);
                const pdfCutParts = pdfParts.filter((x) => x.sheet_status === 'CUT');
                const pct = pdfParts.length ? Math.round((pdfCutParts.length / pdfParts.length) * 100) : 0;
                const isSelected = partsExplorerPdf === pdfKey;

                return (
                  <div
                    key={pdfKey}
                    onClick={() => {
                      setPartsExplorerPdf(isSelected ? 'ALL' : pdfKey);
                      setPartsExplorerPage(1);
                    }}
                    className={`p-3.5 rounded-xl border transition cursor-pointer shadow ${
                      isSelected
                        ? 'bg-amber-950/30 border-amber-400 shadow-amber-500/10'
                        : 'bg-[#131d33] border-[#223354] hover:border-slate-500'
                    }`}
                  >
                    <div className="flex items-center justify-between">
                      <span className={`pdf-badge ${pdfKey.toLowerCase().replace(' ', '-')}`}>
                        {pdfKey}
                      </span>
                      <span className="text-xs font-mono font-bold text-emerald-400">{pct}%</span>
                    </div>

                    <div className="text-sm font-black text-white mt-2">
                      {pdfParts.length.toLocaleString()} Parts
                    </div>

                    <div className="w-full bg-[#080d1a] h-1.5 rounded-full overflow-hidden mt-2">
                      <div
                        className="bg-emerald-400 h-full transition-all duration-500"
                        style={{ width: `${pct}%` }}
                      />
                    </div>

                    <div className="text-[10.5px] text-slate-400 mt-1.5 flex justify-between">
                      <span>Cut: {pdfCutParts.length}</span>
                      <span>Pending: {pdfParts.length - pdfCutParts.length}</span>
                    </div>
                  </div>
                );
              })}
            </div>

            {/* Filter and Search Bar for Parts Explorer */}
            <div className="bg-[#131d33] border border-[#223354] rounded-xl p-3 flex flex-wrap items-center justify-between gap-3 shadow">
              <div className="flex items-center gap-2 flex-wrap flex-1">
                {/* Search Parts */}
                <div className="relative min-w-[200px] flex-1 max-w-sm">
                  <Search className="w-3.5 h-3.5 text-slate-400 absolute left-2.5 top-1/2 -translate-y-1/2" />
                  <input
                    type="text"
                    placeholder="Search part code [DG], name, cabinet, villa..."
                    value={partsExplorerSearch}
                    onChange={(e) => {
                      setPartsExplorerSearch(e.target.value);
                      setPartsExplorerPage(1);
                    }}
                    className="w-full bg-[#0b1325] border border-[#223354] rounded-lg pl-8 pr-3 py-1.5 text-xs text-white placeholder-slate-500 focus:outline-none focus:border-amber-400"
                  />
                </div>

                {/* Part Code Filter Dropdown (e.g. [DA], [DG], etc.) */}
                <select
                  value={partsExplorerCodeFilter}
                  onChange={(e) => {
                    setPartsExplorerCodeFilter(e.target.value);
                    setPartsExplorerPage(1);
                  }}
                  className="bg-[#0b1325] border border-rose-500/50 text-xs font-black text-rose-300 rounded-lg px-2.5 py-1.5 focus:outline-none focus:border-rose-400"
                >
                  <option value="ALL">🏷️ All Part Codes ({uniquePartCodes.length} codes)</option>
                  {uniquePartCodes.map((c) => (
                    <option key={c} value={c}>
                      [{c}] Part Code
                    </option>
                  ))}
                </select>

                {/* Status Filter (Pending vs Cut) */}
                <select
                  value={partsExplorerStatusFilter}
                  onChange={(e) => {
                    setPartsExplorerStatusFilter(e.target.value as any);
                    setPartsExplorerPage(1);
                  }}
                  className="bg-[#0b1325] border border-amber-500/40 text-xs font-bold text-amber-300 rounded-lg px-2.5 py-1.5 focus:outline-none focus:border-amber-400"
                >
                  <option value="ALL">All Statuses (Pending &amp; Cut)</option>
                  <option value="PENDING">⚠️ Pending Cuts Only (Remaining &gt; 0)</option>
                  <option value="CUT">✅ Completed Only (100% Done)</option>
                </select>

                {/* PDF Set Filter */}
                <select
                  value={partsExplorerPdf}
                  onChange={(e) => {
                    setPartsExplorerPdf(e.target.value);
                    setPartsExplorerPage(1);
                  }}
                  className="bg-[#0b1325] border border-[#223354] text-xs font-bold text-white rounded-lg px-2.5 py-1.5 focus:outline-none focus:border-amber-400"
                >
                  <option value="ALL">All Drawing Sets (PDF 1–5)</option>
                  <option value="PDF 1">PDF 1 (SCMB)</option>
                  <option value="PDF 2">PDF 2 (CB-03-SF)</option>
                  <option value="PDF 3">PDF 3 (CB-02-SF)</option>
                  <option value="PDF 4">PDF 4 (SVMB Master)</option>
                  <option value="PDF 5">PDF 5 (CB-01 / Drawers)</option>
                </select>

                {/* Sort Parts By */}
                <select
                  value={partsExplorerSortBy}
                  onChange={(e) => {
                    setPartsExplorerSortBy(e.target.value as any);
                    setPartsExplorerPage(1);
                  }}
                  className="bg-[#0b1325] border border-sky-500/40 text-xs font-bold text-sky-300 rounded-lg px-2.5 py-1.5 focus:outline-none focus:border-sky-400"
                >
                  <option value="CODE">🏷️ Sort: Part Code (A–Z)</option>
                  <option value="NAME">📛 Sort: Part Name (A–Z)</option>
                  <option value="REMAINING">⚠️ Sort: Most Remaining to Cut First</option>
                  <option value="TOTAL">🔢 Sort: Total Needed (Highest First)</option>
                  <option value="CUT">✅ Sort: Most Cut First</option>
                  <option value="LENGTH">📏 Sort: Length (Longest First)</option>
                  <option value="DIM">📐 Sort: Area / Size (Largest First)</option>
                  <option value="PDF">🧩 Sort: PDF Set (PDF 1–5)</option>
                  <option value="CAB">🗄️ Sort: Cabinet Group</option>
                  <option value="PROGRESS_ASC">⏱️ Sort: Least Completed (% Done)</option>
                  <option value="PROGRESS_DESC">🏆 Sort: Most Completed (% Done)</option>
                </select>

                {/* Page Size */}
                <select
                  value={partsExplorerPageSize}
                  onChange={(e) => {
                    setPartsExplorerPageSize(Number(e.target.value));
                    setPartsExplorerPage(1);
                  }}
                  className="bg-[#0b1325] border border-[#223354] text-xs font-bold text-slate-300 rounded-lg px-2 py-1.5 focus:outline-none"
                >
                  <option value={25}>Show 25</option>
                  <option value={50}>Show 50</option>
                  <option value={100}>Show 100</option>
                  <option value={9999}>
                    Show All (
                    {partsExplorerViewMode === 'GROUPED'
                      ? filteredGroupedParts.length
                      : filteredExplorerParts.length}
                    )
                  </option>
                </select>
              </div>

              {/* Reset Filter Button */}
              {(partsExplorerPdf !== 'ALL' ||
                partsExplorerCodeFilter !== 'ALL' ||
                partsExplorerStatusFilter !== 'ALL' ||
                partsExplorerSearch ||
                partsExplorerSortBy !== 'CODE') && (
                <button
                  onClick={() => {
                    setPartsExplorerPdf('ALL');
                    setPartsExplorerCodeFilter('ALL');
                    setPartsExplorerStatusFilter('ALL');
                    setPartsExplorerSearch('');
                    setPartsExplorerSortBy('CODE');
                    setPartsExplorerPage(1);
                  }}
                  className="px-2.5 py-1.5 bg-[#0b1325] hover:bg-[#1c2a4a] text-xs font-bold text-amber-400 rounded-lg border border-amber-500/30 transition flex items-center gap-1"
                >
                  <X className="w-3.5 h-3.5" /> Reset Filters
                </button>
              )}
            </div>

            {/* ============================================================== */}
            {/* VIEW MODE 1: GROUPED BY PART CODE & NAME                       */}
            {/* ============================================================== */}
            {partsExplorerViewMode === 'GROUPED' && (
              <div className="bg-[#10172b] border border-[#223354] rounded-2xl overflow-hidden shadow-2xl">
                <div className="overflow-x-auto max-h-[calc(100vh-360px)] overflow-y-auto">
                  <table className="w-full text-left text-xs text-slate-200 border-collapse table-auto">
                    <thead className="bg-[#0b1325] text-slate-300 font-black uppercase text-[11px] tracking-wider sticky top-0 z-20 border-b-2 border-[#223354] shadow-md">
                      <tr>
                        <th className="py-3.5 px-3">Part Code</th>
                        <th className="py-3.5 px-3">Part Name &amp; Lipping</th>
                        <th className="py-3.5 px-3">Cut Dimensions &amp; Material</th>
                        <th className="py-3.5 px-3">Drawing Sets</th>
                        <th className="py-3.5 px-3 text-center">Remaining to Cut</th>
                        <th className="py-3.5 px-3">Cut vs What We Need</th>
                        <th className="py-3.5 px-3">Parent Sheets</th>
                        <th className="py-3.5 px-3 text-right">Details</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-[#223354]/40 font-medium">
                      {paginatedGroupedParts.map((g) => {
                        const isExpanded = expandedPartId === g.id;

                        return (
                          <React.Fragment key={g.id}>
                            <tr
                              className={`transition hover:bg-[#131d33]/80 ${
                                isExpanded ? 'bg-[#15233e]' : ''
                              }`}
                            >
                              {/* Part Code */}
                              <td className="py-3 px-3 align-top">
                                <div className="inline-flex items-center gap-1.5 bg-rose-950/40 border border-rose-500/50 px-2 py-1 rounded-lg">
                                  <strong className="text-rose-300 font-mono font-black text-sm">
                                    [{g.code}]
                                  </strong>
                                </div>
                              </td>

                              {/* Part Name & Lipping */}
                              <td className="py-3 px-3 align-top">
                                <div className="font-extrabold text-white text-xs leading-snug">
                                  {g.name}
                                </div>
                                <div className="text-[10.5px] text-slate-400 font-mono mt-0.5">
                                  Lip: <span className="text-amber-300 font-bold">{g.lip || 'Standard'}</span>
                                  {g.cab && (
                                    <>
                                      {' '}
                                      • <span className="text-slate-300">{g.cab}</span>
                                    </>
                                  )}
                                  {g.villa && (
                                    <>
                                      {' '}
                                      • <span className="text-amber-400">{g.villa}</span>
                                    </>
                                  )}
                                </div>
                              </td>

                              {/* Dimensions & Material */}
                              <td className="py-3 px-3 align-top font-mono">
                                <div className="text-emerald-300 font-bold text-xs">
                                  {g.l} × {g.w} × {g.t} mm
                                </div>
                                <div className="text-[10px] text-slate-400 truncate max-w-[180px]" title={g.core}>
                                  {g.core}
                                </div>
                              </td>

                              {/* Drawing Sets */}
                              <td className="py-3 px-3 align-top">
                                <div className="flex flex-wrap gap-1">
                                  {g.pdf_sets.map((pdf) => (
                                    <span
                                      key={pdf}
                                      className={`pdf-badge text-[9.5px] py-0.5 px-1.5 ${pdf.toLowerCase().replace(' ', '-')}`}
                                    >
                                      {pdf}
                                    </span>
                                  ))}
                                </div>
                              </td>

                              {/* Remaining to Cut */}
                              <td className="py-3 px-3 align-top text-center">
                                {g.remaining > 0 ? (
                                  <div className="inline-flex flex-col items-center">
                                    <span className="px-3 py-1 bg-rose-500/20 text-rose-300 border border-rose-500/50 rounded-lg font-black font-mono text-xs shadow-sm flex items-center gap-1.5">
                                      <AlertTriangle className="w-3.5 h-3.5 text-rose-400" />
                                      <span>{g.remaining} pcs LEFT</span>
                                    </span>
                                    <span className="text-[10px] text-rose-400/80 font-bold mt-0.5">
                                      in {g.pending_sheets_count} sheets
                                    </span>
                                  </div>
                                ) : (
                                  <div className="inline-flex flex-col items-center">
                                    <span className="px-3 py-1 bg-emerald-500/20 text-emerald-300 border border-emerald-500/50 rounded-lg font-black font-mono text-xs">
                                      ✅ ALL CUT
                                    </span>
                                    <span className="text-[10px] text-emerald-400/80 font-bold mt-0.5">
                                      0 remaining
                                    </span>
                                  </div>
                                )}
                              </td>

                              {/* Cut vs What We Need */}
                              <td className="py-3 px-3 align-top">
                                <div className="flex items-center justify-between text-xs font-mono font-bold">
                                  <span className="text-white">
                                    {g.total_cut} / {g.total_needed} pcs
                                  </span>
                                  <span className="text-emerald-400">{g.pct_cut}%</span>
                                </div>
                                <div className="w-28 bg-[#080d1a] h-2 rounded-full overflow-hidden mt-1.5 border border-[#223354]">
                                  <div
                                    className={`h-full transition-all duration-300 ${
                                      g.pct_cut === 100 ? 'bg-emerald-400' : 'bg-sky-400'
                                    }`}
                                    style={{ width: `${g.pct_cut}%` }}
                                  />
                                </div>
                              </td>

                              {/* Parent Sheets */}
                              <td className="py-3 px-3 align-top">
                                <div className="text-slate-300 text-xs font-bold font-mono">
                                  {g.sheets_count} sheets ({g.pending_sheets_count} pending)
                                </div>
                                {g.remaining_sheets.length > 0 && (
                                  <div className="flex flex-wrap gap-1 mt-1 max-w-[200px]">
                                    {g.remaining_sheets.slice(0, 3).map((rs, rIdx) => (
                                      <span
                                        key={rIdx}
                                        className="bg-rose-950/40 border border-rose-800/40 text-rose-200 text-[9.5px] font-mono px-1.5 py-0.5 rounded truncate max-w-[100px]"
                                        title={`${rs.sheet_title}: ${rs.qty_on_sheet} pcs (${rs.assigned_saw})`}
                                      >
                                        {rs.sheet_id} ({rs.qty_on_sheet}p)
                                      </span>
                                    ))}
                                    {g.remaining_sheets.length > 3 && (
                                      <span className="text-[9.5px] text-rose-300 bg-rose-950/20 px-1 py-0.5 rounded border border-rose-800/30">
                                        +{g.remaining_sheets.length - 3}
                                      </span>
                                    )}
                                  </div>
                                )}
                              </td>

                              {/* Details Action */}
                              <td className="py-3 px-3 align-top text-right">
                                <button
                                  onClick={() => setExpandedPartId(isExpanded ? null : g.id)}
                                  className="px-2.5 py-1 bg-[#0b1325] hover:bg-[#1c2a4a] text-slate-300 border border-[#223354] rounded text-xs font-bold inline-flex items-center gap-1 transition"
                                  title="Expand sheet occurrences"
                                >
                                  {isExpanded ? (
                                    <>
                                      <span>Hide</span>
                                      <ChevronUp className="w-3.5 h-3.5" />
                                    </>
                                  ) : (
                                    <>
                                      <span>{g.sheets_count} Sheets</span>
                                      <ChevronDown className="w-3.5 h-3.5" />
                                    </>
                                  )}
                                </button>
                              </td>
                            </tr>

                            {/* Expanded Details Row */}
                            {isExpanded && (
                              <tr className="bg-[#0b1325]/90 border-b-2 border-amber-500/40">
                                <td colSpan={8} className="p-4">
                                  <div className="space-y-3">
                                    <div className="flex items-center justify-between flex-wrap gap-2">
                                      <div className="font-extrabold text-sm text-white flex items-center gap-2">
                                        <span className="text-rose-400 font-mono">[{g.code}]</span>
                                        <span>{g.name}</span>
                                        <span className="text-slate-400 font-normal">
                                          ({g.l} × {g.w} × {g.t} mm) — All {g.sheets_count} Board Sheets
                                        </span>
                                      </div>
                                      <div className="text-xs font-bold text-slate-400 font-mono">
                                        Total Needed: <strong className="text-white">{g.total_needed}</strong> | Already Cut: <strong className="text-emerald-400">{g.total_cut}</strong> | Remaining: <strong className="text-rose-400">{g.remaining}</strong>
                                      </div>
                                    </div>

                                    <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-2">
                                      {g.sheets.map((sh, sIdx) => {
                                        const sRec = (serverState?.sheets || []).find((x) => x.sheet_id === sh.sheet_id);

                                        return (
                                          <div
                                            key={sIdx}
                                            className={`p-2.5 rounded-lg border text-xs flex items-center justify-between ${
                                              sh.status === 'CUT'
                                                ? 'bg-emerald-950/20 border-emerald-800/40'
                                                : 'bg-[#10172b] border-[#223354]'
                                            }`}
                                          >
                                            <div>
                                              <div className="font-mono font-bold text-white flex items-center gap-1.5">
                                                <span>{sh.sheet_title}</span>
                                                <span className={`st st-${sh.status} text-[9.5px]`}>
                                                  {sh.status}
                                                </span>
                                              </div>
                                              <div className="text-[10px] text-slate-400 mt-0.5">
                                                Saw: {sh.assigned_saw} • {sh.pdf_set}
                                              </div>
                                            </div>

                                            <div className="text-right">
                                              <div className="font-mono font-black text-amber-300 text-sm">
                                                {sh.qty_on_sheet} pcs
                                              </div>
                                              {sRec && (
                                                <button
                                                  onClick={() => setDiagramModalSheet(sRec)}
                                                  className="mt-1 px-2 py-0.5 bg-sky-600/30 hover:bg-sky-600 text-sky-200 rounded text-[10px] font-bold transition inline-flex items-center gap-1"
                                                >
                                                  <Eye className="w-2.5 h-2.5" /> Diagram
                                                </button>
                                              )}
                                            </div>
                                          </div>
                                        );
                                      })}
                                    </div>
                                  </div>
                                </td>
                              </tr>
                            )}
                          </React.Fragment>
                        );
                      })}

                      {filteredGroupedParts.length === 0 && (
                        <tr>
                          <td colSpan={8} className="py-12 text-center text-slate-400">
                            <div className="flex flex-col items-center justify-center space-y-2">
                              <Search className="w-8 h-8 text-slate-600" />
                              <div className="font-bold text-slate-300">No parts found matching query</div>
                              <div className="text-xs text-slate-500">
                                Try adjusting your Part Code filter, status filter, or search query.
                              </div>
                            </div>
                          </td>
                        </tr>
                      )}
                    </tbody>
                  </table>
                </div>

                {/* Pagination Controls (Grouped Mode) */}
                {groupedTotalPages > 1 && (
                  <div className="p-3 bg-[#0f172a] border-t border-[#223354] flex items-center justify-between text-xs text-slate-300">
                    <div>
                      Showing {(partsExplorerPage - 1) * partsExplorerPageSize + 1} to{' '}
                      {Math.min(partsExplorerPage * partsExplorerPageSize, filteredGroupedParts.length)} of{' '}
                      {filteredGroupedParts.length} unique component parts
                    </div>

                    <div className="flex items-center gap-2">
                      <button
                        onClick={() => setPartsExplorerPage((p) => Math.max(1, p - 1))}
                        disabled={partsExplorerPage === 1}
                        className="px-2.5 py-1 bg-[#131d33] border border-[#223354] rounded text-slate-300 hover:text-white disabled:opacity-30 disabled:cursor-not-allowed transition"
                      >
                        <ChevronLeft className="w-3.5 h-3.5 inline" /> Prev
                      </button>
                      <span className="font-bold text-amber-400">
                        Page {partsExplorerPage} of {groupedTotalPages}
                      </span>
                      <button
                        onClick={() => setPartsExplorerPage((p) => Math.min(groupedTotalPages, p + 1))}
                        disabled={partsExplorerPage === groupedTotalPages}
                        className="px-2.5 py-1 bg-[#131d33] border border-[#223354] rounded text-slate-300 hover:text-white disabled:opacity-30 disabled:cursor-not-allowed transition"
                      >
                        Next <ChevronRight className="w-3.5 h-3.5 inline" />
                      </button>
                    </div>
                  </div>
                )}
              </div>
            )}

            {/* ============================================================== */}
            {/* VIEW MODE 2: FLAT INDIVIDUAL SHEET OCCURRENCES (10,183 ROWS)   */}
            {/* ============================================================== */}
            {partsExplorerViewMode === 'FLAT' && (
              <div className="bg-[#10172b] border border-[#223354] rounded-2xl overflow-hidden shadow-2xl">
                <div className="overflow-x-auto max-h-[calc(100vh-360px)] overflow-y-auto">
                  <table className="w-full text-left text-xs text-slate-200 border-collapse table-auto">
                    <thead className="bg-[#0b1325] text-slate-300 font-black uppercase text-[11px] tracking-wider sticky top-0 z-20 border-b-2 border-[#223354] shadow-md">
                      <tr>
                        <th className="py-3 px-3">PDF Drawing Set</th>
                        <th className="py-3 px-3">Code &amp; Part Name</th>
                        <th className="py-3 px-3">Dimensions &amp; Core</th>
                        <th className="py-3 px-3">Cabinet / Villa</th>
                        <th className="py-3 px-3">Qty / Sheet</th>
                        <th className="py-3 px-3">Parent Sheet Location</th>
                        <th className="py-3 px-3">Board Cut Status</th>
                        <th className="py-3 px-3 text-right">Actions</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-[#223354]/40 font-medium">
                      {paginatedExplorerParts.map((item, idx) => {
                        const p = item.part;
                        const sRecord = (serverState?.sheets || []).find((x) => x.sheet_id === item.sheet_id);

                        return (
                          <tr
                            key={`${item.sheet_id}_${p.code}_${idx}`}
                            className="hover:bg-[#131d33]/80 transition group"
                          >
                            {/* PDF Set */}
                            <td className="py-3 px-3 align-top">
                              <span className={`pdf-badge ${p.pdf_set.toLowerCase().replace(' ', '-')}`}>
                                {p.pdf_set}
                              </span>
                            </td>

                            {/* Code & Name */}
                            <td className="py-3 px-3 align-top">
                              <div className="flex items-center gap-1.5 flex-wrap">
                                <strong className="text-rose-400 font-mono font-black text-sm">
                                  [{p.code}]
                                </strong>
                                <span className="font-extrabold text-white text-xs">{p.name}</span>
                              </div>
                              <div className="text-[10.5px] text-slate-400 font-mono mt-0.5">
                                Lip: <span className="text-amber-300 font-bold">{p.lip || '2mm Front'}</span>
                              </div>
                            </td>

                            {/* Dimensions & Core */}
                            <td className="py-3 px-3 align-top font-mono">
                              <div className="text-emerald-300 font-bold text-xs">
                                {p.l} × {p.w} × {p.t} mm
                              </div>
                              <div className="text-[10px] text-slate-400 truncate max-w-[180px]" title={p.core}>
                                {p.core}
                              </div>
                            </td>

                            {/* Cabinet & Villa */}
                            <td className="py-3 px-3 align-top">
                              <div className="text-xs font-bold text-white truncate max-w-[160px]" title={p.cab}>
                                {p.cab}
                              </div>
                              <div className="text-[10.5px] text-amber-300 font-semibold">
                                {p.villa}
                              </div>
                            </td>

                            {/* Quantity */}
                            <td className="py-3 px-3 align-top">
                              <span className="text-sky-300 font-extrabold text-xs">
                                {p.qty_per_sheet} pcs / sheet
                              </span>
                              <div className="text-[10px] text-slate-400 font-mono">
                                Total Batch: {p.total_batch_qty} pcs
                              </div>
                            </td>

                            {/* Sheet Location */}
                            <td className="py-3 px-3 align-top">
                              <div className="font-mono text-sky-400 font-bold text-xs">
                                {item.sheet_title}
                              </div>
                              <div className="text-[10px] text-slate-400 font-mono">
                                {item.sheet_id} ({item.thickness}mm)
                              </div>
                            </td>

                            {/* Sheet Cut Status */}
                            <td className="py-3 px-3 align-top">
                              <span className={`st st-${item.sheet_status}`}>
                                {item.sheet_status}
                              </span>
                              <div className="text-[10px] text-slate-400 mt-0.5">
                                {item.assigned_saw}
                              </div>
                            </td>

                            {/* Actions */}
                            <td className="py-3 px-3 align-top text-right space-x-1.5">
                              {sRecord && (
                                <>
                                  <button
                                    onClick={() => handlePrintSingleSheet(sRecord)}
                                    className="px-2.5 py-1 bg-sky-600/20 hover:bg-sky-600 text-sky-300 hover:text-white border border-sky-500/40 rounded text-[11px] font-bold inline-flex items-center gap-1 transition"
                                    title="Print sticker for this component"
                                  >
                                    <Printer className="w-3 h-3" />
                                    <span>Print Sticker</span>
                                  </button>

                                  <button
                                    onClick={() => setDiagramModalSheet(sRecord)}
                                    className="px-2 py-1 bg-[#0b1325] hover:bg-[#1c2a4a] text-slate-300 border border-[#223354] rounded text-[11px] font-bold inline-flex items-center gap-1 transition"
                                    title="Inspect board diagram"
                                  >
                                    <Eye className="w-3 h-3 text-sky-400" />
                                    <span>Diagram</span>
                                  </button>
                                </>
                              )}
                            </td>
                          </tr>
                        );
                      })}

                      {filteredExplorerParts.length === 0 && (
                        <tr>
                          <td colSpan={8} className="py-12 text-center text-slate-400">
                            <div className="flex flex-col items-center justify-center space-y-2">
                              <Search className="w-8 h-8 text-slate-600" />
                              <div className="font-bold text-slate-300">No parts found matching query</div>
                              <div className="text-xs text-slate-500">
                                Try adjusting your Part Code filter, status filter, or search query.
                              </div>
                            </div>
                          </td>
                        </tr>
                      )}
                    </tbody>
                  </table>
                </div>

                {/* Pagination Controls (Flat Mode) */}
                {explorerTotalPages > 1 && (
                  <div className="p-3 bg-[#0f172a] border-t border-[#223354] flex items-center justify-between text-xs text-slate-300">
                    <div>
                      Showing {(partsExplorerPage - 1) * partsExplorerPageSize + 1} to{' '}
                      {Math.min(partsExplorerPage * partsExplorerPageSize, filteredExplorerParts.length)} of{' '}
                      {filteredExplorerParts.length} component parts
                    </div>

                    <div className="flex items-center gap-2">
                      <button
                        onClick={() => setPartsExplorerPage((p) => Math.max(1, p - 1))}
                        disabled={partsExplorerPage === 1}
                        className="px-2.5 py-1 bg-[#131d33] border border-[#223354] rounded text-slate-300 hover:text-white disabled:opacity-30 disabled:cursor-not-allowed transition"
                      >
                        <ChevronLeft className="w-3.5 h-3.5 inline" /> Prev
                      </button>
                      <span className="font-bold text-amber-400">
                        Page {partsExplorerPage} of {explorerTotalPages}
                      </span>
                      <button
                        onClick={() => setPartsExplorerPage((p) => Math.min(explorerTotalPages, p + 1))}
                        disabled={partsExplorerPage === explorerTotalPages}
                        className="px-2.5 py-1 bg-[#131d33] border border-[#223354] rounded text-slate-300 hover:text-white disabled:opacity-30 disabled:cursor-not-allowed transition"
                      >
                        Next <ChevronRight className="w-3.5 h-3.5 inline" />
                      </button>
                    </div>
                  </div>
                )}
              </div>
            )}
          </div>
        )}

        {/* ============================================================== */}
        {/* TAB 2: PDF COMPLETION MATRIX (PDF 1 to 5)                      */}
        {/* ============================================================== */}
        {activeTab === 'pdf-matrix' && (
          <div className="space-y-4">
            <div className="bg-[#131d33] border border-[#223354] rounded-xl p-4">
              <h2 className="text-sm font-extrabold text-white">
                Cross-Sheet PDF Dependency Matrix: How many sheets must be cut to complete each drawing set?
              </h2>
              <p className="text-xs text-slate-400 mt-1">
                A single PDF drawing set is often split across multiple sheets. Use this matrix to identify and prioritize blocking sheets for assembly.
              </p>
            </div>

            <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
              {Object.entries(serverState?.pdf_matrix || {}).map(([key, item]) => {
                const isFinished = item.pending_sheets === 0;
                return (
                  <div
                    key={key}
                    className={`bg-[#131d33] border rounded-xl p-4.5 space-y-3 shadow-md ${
                      isFinished ? 'border-emerald-500/80 bg-emerald-950/10' : 'border-[#223354]'
                    }`}
                  >
                    <div className="flex justify-between items-start">
                      <div>
                        <span className={`pdf-badge ${item.name.toLowerCase().replace(' ', '-')}`}>
                          {item.name}
                        </span>
                        <strong className="text-base text-white ml-1 font-bold">
                          {item.description}
                        </strong>
                      </div>
                      <span className={`st ${isFinished ? 'st-CUT' : 'st-CUTTING'}`}>
                        {item.percent_complete}%
                      </span>
                    </div>

                    {/* Progress Bar */}
                    <div className="w-full bg-[#0b1325] h-2.5 rounded-full overflow-hidden border border-[#223354]">
                      <div
                        className="bg-gradient-to-r from-sky-500 to-emerald-400 h-full transition-all duration-500"
                        style={{ width: `${item.percent_complete}%` }}
                      />
                    </div>

                    <div className="text-xs space-y-1 text-slate-300">
                      <div>
                        Required Sheets to Finish:{' '}
                        <strong className="text-white font-bold">{item.total_sheets} Sheets</strong>
                      </div>
                      <div>
                        Sheets Cut:{' '}
                        <strong className="text-emerald-400 font-bold">{item.cut_sheets}</strong> ·
                        Sheets Remaining:{' '}
                        <strong className="text-rose-400 font-bold">{item.pending_sheets}</strong>
                      </div>
                      <div>
                        Total Parts: <strong>{item.cut_parts} / {item.total_parts} cut</strong>
                      </div>
                    </div>

                    <div className="pt-2 border-t border-[#223354]/60">
                      {item.pending_sheets > 0 ? (
                        <button
                          onClick={() => {
                            setPdfFilter(item.name);
                            setStatusFilter('ALL');
                            setActiveTab('dispatch');
                          }}
                          className="w-full py-2 bg-sky-600 hover:bg-sky-500 text-white rounded-lg text-xs font-bold transition flex items-center justify-center gap-1.5 shadow"
                        >
                          <span>⚡ Filter {item.pending_sheets} Sheets Blocking {item.name}</span>
                        </button>
                      ) : (
                        <div className="text-emerald-400 font-bold text-xs flex items-center gap-1.5 py-1">
                          <CheckCircle2 className="w-4 h-4 text-emerald-400" />
                          <span>All parts cut! Ready for edge banding &amp; assembly.</span>
                        </div>
                      )}
                    </div>
                  </div>
                );
              })}
            </div>
          </div>
        )}

        {/* ============================================================== */}
        {/* TAB 3: SHIFT, CYCLE TIME & LIVE ETA FORECAST                   */}
        {/* ============================================================== */}
        {activeTab === 'time-forecast' && (
          <div className="space-y-4">
            <div className="grid grid-cols-2 sm:grid-cols-4 gap-3">
              <div className="bg-[#131d33] border border-[#223354] rounded-xl p-3.5">
                <div className="text-[10px] font-bold text-rose-400 uppercase tracking-wider">
                  Pending Boards
                </div>
                <div className="text-2xl font-black text-rose-400 mt-0.5">
                  {forecastData.pendingSheets}
                </div>
                <div className="text-[11px] text-slate-400">Total remaining volume</div>
              </div>

              <div className="bg-[#131d33] border border-[#223354] rounded-xl p-3.5">
                <div className="text-[10px] font-bold text-sky-400 uppercase tracking-wider">
                  Active Saws
                </div>
                <div className="text-2xl font-black text-sky-400 mt-0.5">
                  {forecastData.activeSaws} Saw{forecastData.activeSaws > 1 ? 's' : ''}
                </div>
                <div className="text-[11px] text-slate-400">
                  {forecastData.activeSaws === 1 ? 'Saw 2 In Repair Mode' : 'Both Online'}
                </div>
              </div>

              <div className="bg-[#131d33] border border-[#223354] rounded-xl p-3.5">
                <div className="text-[10px] font-bold text-emerald-400 uppercase tracking-wider">
                  Cycle Time
                </div>
                <div className="text-2xl font-black text-emerald-400 mt-0.5">
                  {customCycleTime} m
                </div>
                <div className="text-[11px] text-slate-400">Per board on saw table</div>
              </div>

              <div className="bg-[#131d33] border border-[#223354] rounded-xl p-3.5">
                <div className="text-[10px] font-bold text-amber-400 uppercase tracking-wider">
                  Days to Complete
                </div>
                <div className="text-2xl font-black text-amber-400 mt-0.5">
                  {forecastData.workDaysNeeded} d
                </div>
                <div className="text-[11px] text-slate-400">At selected OT rate</div>
              </div>
            </div>

            <div className="bg-[#131d33] border border-[#223354] rounded-xl p-5 space-y-4 shadow">
              <div>
                <h3 className="text-base font-extrabold text-white flex items-center gap-2">
                  <Clock className="w-5 h-5 text-sky-400" />
                  <span>Dynamic Completion Recalculator</span>
                </h3>
                <p className="text-xs text-slate-400 mt-1">
                  Standard Factory Shift: <strong>8:00 AM – 12:00 PM</strong> (4h) &amp;{' '}
                  <strong>1:00 PM – 4:00 PM</strong> (3h) ={' '}
                  <strong className="text-white">7.0 hours/day (420 productive minutes/saw)</strong>.
                </p>
              </div>

              {/* Slider & Select Controls */}
              <div className="flex gap-6 items-center flex-wrap bg-[#0b1325] p-4 rounded-lg border border-[#223354]">
                <div>
                  <label className="font-extrabold text-xs text-slate-300 block mb-1">
                    TABLE SAW CYCLE TIME (MINUTES / SHEET):
                  </label>
                  <div className="flex items-center gap-3">
                    <input
                      type="range"
                      min={4}
                      max={18}
                      step={0.5}
                      value={customCycleTime}
                      onChange={(e) => setCustomCycleTime(parseFloat(e.target.value))}
                      className="w-56 accent-sky-400 cursor-pointer"
                    />
                    <span className="font-black text-sm text-sky-400">
                      {customCycleTime} min/sheet
                    </span>
                  </div>
                </div>

                <div>
                  <label className="font-extrabold text-xs text-slate-300 block mb-1">
                    ESTIMATED OVERTIME PLANNED PER DAY:
                  </label>
                  <select
                    value={plannedOtHours}
                    onChange={(e) => setPlannedOtHours(parseFloat(e.target.value))}
                    className="bg-[#131d33] border border-[#223354] rounded-lg px-3 py-1.5 text-xs font-bold text-white focus:outline-none"
                  >
                    <option value={0}>0 Hours (Leave at 4:00 PM)</option>
                    <option value={1.5}>1.5 Hours OT (Leave at 5:30 PM)</option>
                    <option value={2.5}>2.5 Hours OT (Leave at 6:30 PM)</option>
                    <option value={3.5}>3.5 Hours OT (Leave at 7:30 PM)</option>
                  </select>
                </div>
              </div>

              {/* Forecast Result Card */}
              <div className="bg-[#0f172a] border border-sky-500/40 rounded-xl p-5 space-y-2">
                <div className="text-lg font-black text-sky-300 flex items-center gap-2">
                  <span>🎯 Projected Job Completion:</span>
                  <u className="text-sky-400 underline decoration-sky-400">{forecastData.dateStr}</u>
                </div>

                <div className="text-xs text-slate-300 space-y-1.5 font-medium leading-relaxed">
                  <div>
                    • <strong>{forecastData.pendingSheets} Pending Boards</strong> remaining across job.
                  </div>
                  <div>
                    • Active Cutting Capacity:{' '}
                    <strong>
                      {forecastData.activeSaws} Saw(s) running at {forecastData.cycleMin} min/sheet
                    </strong>.
                  </div>
                  <div>
                    • Daily Production Output:{' '}
                    <strong className="text-emerald-400">
                      ~{forecastData.dailyCapacity} boards/day
                    </strong>{' '}
                    (7h regular + {plannedOtHours}h overtime).
                  </div>
                  <div>
                    • Work Days Needed:{' '}
                    <strong className="text-amber-400">{forecastData.workDaysNeeded} Work Days</strong> (Fridays excluded).
                  </div>
                </div>
              </div>
            </div>
          </div>
        )}

        {/* ============================================================== */}
        {/* TAB 4: OVERTIME LOG (NEXT MORNING)                             */}
        {/* ============================================================== */}
        {activeTab === 'overtime' && (
          <div className="space-y-4">
            <div className="bg-[#131d33] border border-[#223354] rounded-xl p-5 max-w-xl shadow">
              <h3 className="text-sm font-extrabold text-white flex items-center gap-2">
                <Moon className="w-4 h-4 text-amber-400" />
                <span>Next-Morning Overtime Logger</span>
              </h3>
              <p className="text-xs text-slate-400 mt-1 mb-4">
                Confirm and log previous day overtime hours worked by each saw team:
              </p>

              <form onSubmit={handleLogOvertime} className="flex flex-wrap gap-2 items-center">
                <input
                  type="date"
                  value={otDate}
                  onChange={(e) => setOtDate(e.target.value)}
                  className="bg-[#0b1325] border border-[#223354] text-xs font-bold text-white rounded-lg px-3 py-2"
                  required
                />
                <select
                  value={otTeam}
                  onChange={(e) => setOtTeam(e.target.value)}
                  className="bg-[#0b1325] border border-[#223354] text-xs font-bold text-white rounded-lg px-3 py-2"
                >
                  <option value="Team A (Ahmed)">Team A (Ahmed)</option>
                  <option value="Team B">Team B</option>
                </select>
                <input
                  type="number"
                  step="0.5"
                  min="0"
                  placeholder="OT Hours (e.g. 2.5)"
                  value={otHours}
                  onChange={(e) => setOtHours(e.target.value)}
                  className="bg-[#0b1325] border border-[#223354] text-xs font-bold text-white rounded-lg px-3 py-2 w-36"
                  required
                />
                <button
                  type="submit"
                  className="px-4 py-2 bg-sky-600 hover:bg-sky-500 text-white rounded-lg text-xs font-bold transition shadow"
                >
                  Log Overtime
                </button>
              </form>
            </div>

            {/* Overtime History Table */}
            <div className="bg-[#131d33] border border-[#223354] rounded-xl overflow-hidden shadow">
              <div className="p-3 bg-[#0f172a] border-b border-[#223354] font-extrabold text-xs text-slate-300 uppercase tracking-wider">
                Confirmed Overtime History Log
              </div>
              <div className="overflow-x-auto">
                <table className="w-full text-left text-xs text-slate-200">
                  <thead className="bg-[#0b1325] text-slate-400 font-bold uppercase text-[10px] tracking-wider border-b border-[#223354]">
                    <tr>
                      <th className="py-2.5 px-4">Date</th>
                      <th className="py-2.5 px-4">Team</th>
                      <th className="py-2.5 px-4">Overtime Hours</th>
                      <th className="py-2.5 px-4">Logged Timestamp</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-[#223354]/40 font-medium">
                    {(serverState?.overtime_history || [])
                      .slice()
                      .reverse()
                      .map((log, idx) => (
                        <tr key={idx} className="hover:bg-[#1c2a4a]/40">
                          <td className="py-2.5 px-4 font-bold text-white">{log.date}</td>
                          <td className="py-2.5 px-4">{log.team}</td>
                          <td className="py-2.5 px-4">
                            <strong className="text-sky-400 font-mono font-bold">
                              {log.hours} Hours OT
                            </strong>
                          </td>
                          <td className="py-2.5 px-4 text-slate-400 font-mono text-[11px]">
                            {log.logged_at}
                          </td>
                        </tr>
                      ))}

                    {(!serverState?.overtime_history ||
                      serverState.overtime_history.length === 0) && (
                      <tr>
                        <td colSpan={4} className="py-6 text-center text-slate-500">
                          No overtime hours logged yet.
                        </td>
                      </tr>
                    )}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        )}

        {/* ============================================================== */}
        {/* TAB 5: EXCEL / CSV DATABASE MANAGER                            */}
        {/* ============================================================== */}
        {activeTab === 'database' && (
          <div className="space-y-4">
            <div className="bg-[#131d33] border border-[#223354] rounded-xl p-5 shadow space-y-4">
              <div className="flex flex-wrap items-center justify-between gap-3 border-b border-[#223354] pb-4">
                <div>
                  <h3 className="text-base font-extrabold text-white flex items-center gap-2">
                    <Database className="w-5 h-5 text-emerald-400" />
                    <span>Master Excel / CSV Factory Database</span>
                  </h3>
                  <p className="text-xs text-slate-400 mt-1">
                    Export full 3-sheet Excel spreadsheets (.xlsx) or CSV, update sheet statuses and actual times taken, or upload modified spreadsheets to bulk-sync the factory database.
                  </p>
                </div>

                <div className="flex items-center gap-2 flex-wrap">
                  <a
                    href="/api/export-master-excel"
                    className="px-4 py-2 bg-emerald-600 hover:bg-emerald-500 text-white rounded-lg text-xs font-bold flex items-center gap-2 shadow transition"
                    download
                  >
                    <Download className="w-4 h-4" />
                    <span>Download Master Excel (.xlsx)</span>
                  </a>

                  <a
                    href="/api/export-progress-csv"
                    className="px-3.5 py-2 bg-sky-600 hover:bg-sky-500 text-white rounded-lg text-xs font-bold flex items-center gap-1.5 shadow transition"
                    download
                  >
                    <Download className="w-3.5 h-3.5" />
                    <span>Download Production CSV</span>
                  </a>
                </div>
              </div>

              {/* Upload & Sync Box */}
              <div className="bg-[#0b1325] border border-[#223354] rounded-xl p-4 space-y-3">
                <div className="flex items-center justify-between flex-wrap gap-2">
                  <div className="text-xs font-bold text-slate-200 flex items-center gap-2">
                    <Upload className="w-4 h-4 text-sky-400" />
                    <span>Upload Modified Excel (.xlsx) or CSV to Update Database:</span>
                  </div>
                  <span className="text-[11px] text-slate-400">
                    Matches by "Sheet ID" column to update Status, Duration (min), and Assigned Saw/Team
                  </span>
                </div>

                <div className="flex items-center gap-3 flex-wrap">
                  <input
                    type="file"
                    accept=".xlsx, .xls, .csv"
                    onChange={handleFileUpload}
                    className="text-xs text-slate-300 file:mr-3 file:py-1.5 file:px-3 file:rounded-lg file:border-0 file:text-xs file:font-bold file:bg-sky-600 file:text-white hover:file:bg-sky-500 cursor-pointer"
                  />

                  {bulkImportJson && (
                    <button
                      onClick={handleApplyBulkUpdates}
                      disabled={isBulkImporting}
                      className="px-4 py-1.5 bg-emerald-600 hover:bg-emerald-500 disabled:opacity-40 text-white rounded-lg text-xs font-bold shadow flex items-center gap-1.5 transition"
                    >
                      <Save className="w-3.5 h-3.5" />
                      <span>{isBulkImporting ? 'Syncing...' : 'Apply Updates to Database'}</span>
                    </button>
                  )}
                </div>

                {bulkImportStatus && (
                  <div className="p-3 bg-[#131d33] border border-sky-500/40 rounded-lg text-xs text-sky-300">
                    {bulkImportStatus}
                  </div>
                )}
              </div>

              {/* Live Database Overview Table with Quick Edit & Batch Operations */}
              <div className="space-y-3">
                <div className="flex justify-between items-center flex-wrap gap-2">
                  <div className="text-xs font-bold text-slate-300 uppercase tracking-wider flex items-center gap-2">
                    <span>Live Database Records ({serverState?.sheets.length || 0} sheets / {totalSheetsVolume} boards)</span>
                    {dbSelectedSheetIds.size > 0 && (
                      <span className="bg-sky-500/20 text-sky-300 border border-sky-500/40 text-[11px] px-2 py-0.5 rounded-full font-bold">
                        {dbSelectedSheetIds.size} Selected
                      </span>
                    )}
                  </div>

                  {/* Table Search */}
                  <div className="relative min-w-[220px]">
                    <Search className="w-3.5 h-3.5 text-slate-400 absolute left-2.5 top-1/2 -translate-y-1/2" />
                    <input
                      type="text"
                      placeholder="Filter database rows..."
                      value={dbSearchQuery}
                      onChange={(e) => setDbSearchQuery(e.target.value)}
                      className="w-full bg-[#0b1325] border border-[#223354] rounded-lg pl-8 pr-3 py-1 text-xs text-white placeholder-slate-500 focus:outline-none focus:border-sky-500"
                    />
                  </div>
                </div>

                {/* Batch Edit Controls Toolbar for Selected DB Rows */}
                {dbSelectedSheetIds.size > 0 && (
                  <div className="p-3 bg-[#0c1a30] border-2 border-sky-500 rounded-xl flex flex-wrap items-center justify-between gap-3 shadow-xl animate-in fade-in">
                    <div className="flex items-center gap-2 flex-wrap">
                      <span className="text-xs font-black text-sky-300">
                        ⚡ BATCH EDIT ({dbSelectedSheetIds.size} sheets):
                      </span>

                      {/* Set Status */}
                      <select
                        value={dbBatchStatus}
                        onChange={(e) => setDbBatchStatus(e.target.value)}
                        className="bg-[#0b1325] border border-[#223354] rounded-lg px-2.5 py-1 text-xs font-bold text-amber-400 focus:outline-none"
                      >
                        <option value="">-- Change Status --</option>
                        <option value="QUEUED">QUEUED</option>
                        <option value="ASSIGNED">ASSIGNED</option>
                        <option value="CUTTING">CUTTING</option>
                        <option value="CUT">CUT (Completed)</option>
                      </select>

                      {/* Set Saw */}
                      <select
                        value={dbBatchSaw}
                        onChange={(e) => setDbBatchSaw(e.target.value)}
                        className="bg-[#0b1325] border border-[#223354] rounded-lg px-2.5 py-1 text-xs font-bold text-sky-400 focus:outline-none"
                      >
                        <option value="">-- Assign Saw --</option>
                        <option value="Saw 1">Saw 1 (Team A)</option>
                        <option value="Saw 2">Saw 2 (Team B)</option>
                      </select>

                      {/* Set Duration */}
                      <input
                        type="number"
                        placeholder="Time (min)"
                        value={dbBatchDuration}
                        onChange={(e) => setDbBatchDuration(e.target.value)}
                        className="w-24 bg-[#0b1325] border border-[#223354] rounded-lg px-2 py-1 text-xs text-emerald-400 placeholder-slate-500 font-bold focus:outline-none"
                        step="0.5"
                        min="0"
                      />

                      <button
                        onClick={handleApplyDbBatchEdit}
                        disabled={isDbBatchUpdating || (!dbBatchStatus && !dbBatchSaw && dbBatchDuration === '')}
                        className="px-4 py-1 bg-emerald-600 hover:bg-emerald-500 disabled:opacity-40 text-white rounded-lg text-xs font-bold shadow flex items-center gap-1.5 transition"
                      >
                        <Save className="w-3.5 h-3.5" />
                        <span>{isDbBatchUpdating ? 'Saving...' : 'Apply Batch Changes'}</span>
                      </button>
                    </div>

                    <button
                      onClick={() => setDbSelectedSheetIds(new Set())}
                      className="text-xs text-slate-400 hover:text-white font-bold transition"
                    >
                      Deselect All
                    </button>
                  </div>
                )}

                <div className="bg-[#0b1325] border border-[#223354] rounded-xl overflow-hidden shadow">
                  <div className="overflow-x-auto max-h-[550px]">
                    {(() => {
                      const filteredDbSheets = (serverState?.sheets || []).filter((s) => {
                        if (!dbSearchQuery.trim()) return true;
                        const q = dbSearchQuery.toLowerCase();
                        return (
                          s.sheet_id.toLowerCase().includes(q) ||
                          s.sheet_title.toLowerCase().includes(q) ||
                          s.material.toLowerCase().includes(q) ||
                          s.assigned_saw.toLowerCase().includes(q) ||
                          s.status.toLowerCase().includes(q) ||
                          s.parts.some((p) => p.name.toLowerCase().includes(q) || p.code.toLowerCase().includes(q))
                        );
                      });

                      const isAllDbSelected =
                        filteredDbSheets.length > 0 &&
                        filteredDbSheets.every((s) => dbSelectedSheetIds.has(s.sheet_id));

                      return (
                        <table className="w-full text-left text-xs text-slate-200">
                          <thead className="bg-[#0f172a] text-slate-400 font-bold uppercase text-[10px] tracking-wider border-b border-[#223354] sticky top-0 z-10">
                            <tr>
                              <th className="py-2.5 px-3 w-10 text-center">
                                <input
                                  type="checkbox"
                                  checked={isAllDbSelected}
                                  onChange={(e) => handleDbSelectAll(e.target.checked, filteredDbSheets)}
                                  className="rounded cursor-pointer accent-sky-500"
                                />
                              </th>
                              <th className="py-2.5 px-3">Sheet ID</th>
                              <th className="py-2.5 px-3">Thickness</th>
                              <th className="py-2.5 px-3">Board Count</th>
                              <th className="py-2.5 px-3">Material Specs</th>
                              <th className="py-2.5 px-3">Assigned Saw &amp; Team</th>
                              <th className="py-2.5 px-3">Status</th>
                              <th className="py-2.5 px-3">Time Taken (min)</th>
                              <th className="py-2.5 px-3 text-right">Direct Edit</th>
                            </tr>
                          </thead>
                          <tbody className="divide-y divide-[#223354]/40 font-medium">
                            {filteredDbSheets.map((s) => {
                              const isRowSelected = dbSelectedSheetIds.has(s.sheet_id);
                              return (
                                <tr
                                  key={s.sheet_id}
                                  className={`hover:bg-[#131d33]/60 transition ${
                                    isRowSelected ? 'bg-sky-950/40' : ''
                                  }`}
                                >
                                  <td className="py-2.5 px-3 text-center">
                                    <input
                                      type="checkbox"
                                      checked={isRowSelected}
                                      onChange={(e) => handleDbToggleSelect(s.sheet_id, e.target.checked)}
                                      className="rounded cursor-pointer accent-sky-500"
                                    />
                                  </td>
                                  <td className="py-2.5 px-3 font-mono font-bold text-sky-400">
                                    {s.sheet_id}
                                  </td>
                                  <td className="py-2.5 px-3">
                                    <span
                                      className={`px-2 py-0.5 rounded text-[10px] font-bold ${
                                        s.thickness === 8 ? 'bg-amber-500/20 text-amber-300' : 'bg-sky-500/20 text-sky-300'
                                      }`}
                                    >
                                      {s.thickness || 18}mm
                                    </span>
                                  </td>
                                  <td className="py-2.5 px-3 font-bold text-white">
                                    1 Board
                                  </td>
                                  <td className="py-2.5 px-3 text-slate-300 truncate max-w-[220px]" title={s.material}>
                                    {s.material}
                                  </td>
                                  <td className="py-2.5 px-3">
                                    <span className="font-bold text-white">{s.assigned_saw}</span>
                                    <span className="text-slate-400 text-[11px] block">{s.assigned_team}</span>
                                  </td>
                                  <td className="py-2.5 px-3">
                                    <span className={`st st-${s.status}`}>{s.status}</span>
                                  </td>
                                  <td className="py-2.5 px-3 font-mono font-bold text-emerald-400">
                                    {s.duration_min !== null && s.duration_min !== undefined ? `${s.duration_min} min` : '—'}
                                  </td>
                                  <td className="py-2.5 px-3 text-right">
                                    <button
                                      onClick={() => handleOpenEditSheet(s)}
                                      className="px-2.5 py-1 bg-[#131d33] hover:bg-[#1c2a4a] text-amber-400 border border-amber-500/30 rounded text-xs font-bold inline-flex items-center gap-1 transition shadow"
                                    >
                                      <Edit className="w-3 h-3" />
                                      <span>Edit Record</span>
                                    </button>
                                  </td>
                                </tr>
                              );
                            })}

                            {filteredDbSheets.length === 0 && (
                              <tr>
                                <td colSpan={9} className="py-8 text-center text-slate-500">
                                  No database records match "{dbSearchQuery}".
                                </td>
                              </tr>
                            )}
                          </tbody>
                        </table>
                      );
                    })()}
                  </div>
                </div>
              </div>
            </div>
          </div>
        )}

        {/* ============================================================== */}
        {/* TAB 6: GEMMA 4 PRODUCTION DIRECTOR                             */}
        {/* ============================================================== */}
        {activeTab === 'gemma' && (
          <div className="space-y-4">
            <div className="bg-[#131d33] border border-[#223354] rounded-xl p-5 shadow">
              <div className="flex items-center gap-3 mb-2">
                <div className="p-2 bg-sky-500/20 text-sky-400 rounded-lg">
                  <Bot className="w-5 h-5" />
                </div>
                <div>
                  <h3 className="text-sm font-extrabold text-white">
                    Gemma 4 Table Saw Production Director
                  </h3>
                  <p className="text-xs text-slate-400">
                    Industrial AI advisor grounded in your live OpenCutList data, rip-cut dimensions, and 5 PDF drawing sets.
                  </p>
                </div>
              </div>

              {/* Sample Quick Questions */}
              <div className="flex flex-wrap gap-2 my-3">
                <button
                  onClick={() => handleAskGemma('Which sheets are blocking PDF 2 (CB-03-SF)?')}
                  className="px-2.5 py-1 bg-[#0b1325] hover:bg-[#1c2a4a] text-slate-300 border border-[#223354] rounded text-[11px] transition"
                >
                  ⚡ Which sheets block PDF 2?
                </button>
                <button
                  onClick={() => handleAskGemma('What are the primary table saw rip cuts for Sheet 1 and Sheet 2?')}
                  className="px-2.5 py-1 bg-[#0b1325] hover:bg-[#1c2a4a] text-slate-300 border border-[#223354] rounded text-[11px] transition"
                >
                  📏 First rip cuts for Sheet 1 &amp; 2?
                </button>
                <button
                  onClick={() => handleAskGemma('How many days remaining if Saw 2 stays in repair with 2.5h overtime?')}
                  className="px-2.5 py-1 bg-[#0b1325] hover:bg-[#1c2a4a] text-slate-300 border border-[#223354] rounded text-[11px] transition"
                >
                  ⏱️ Days remaining with 1 saw?
                </button>
              </div>

              {/* Query Box */}
              <div className="space-y-2">
                <div className="flex gap-2">
                  <input
                    type="text"
                    placeholder="Ask Gemma a production or sequencing question..."
                    value={gemmaPrompt}
                    onChange={(e) => setGemmaPrompt(e.target.value)}
                    onKeyDown={(e) => {
                      if (e.key === 'Enter') handleAskGemma();
                    }}
                    className="flex-1 bg-[#0b1325] border border-[#223354] rounded-lg px-3.5 py-2 text-xs text-white placeholder-slate-500 focus:outline-none focus:border-sky-500"
                  />
                  <button
                    onClick={() => handleAskGemma()}
                    disabled={gemmaLoading || !gemmaPrompt.trim()}
                    className="px-4 py-2 bg-sky-600 hover:bg-sky-500 disabled:opacity-40 text-white rounded-lg text-xs font-bold transition shadow"
                  >
                    {gemmaLoading ? 'Asking...' : 'Ask Director'}
                  </button>
                </div>

                {/* Gemma Response Box */}
                {gemmaResponse && (
                  <div className="mt-4 p-4 bg-[#0b1325] border border-sky-500/40 rounded-xl space-y-2 text-xs text-slate-200">
                    <div className="font-extrabold text-sky-400 text-xs flex items-center gap-1.5">
                      <Bot className="w-3.5 h-3.5" />
                      <span>Gemma 4 Shop Floor Recommendation:</span>
                    </div>
                    <div className="whitespace-pre-wrap leading-relaxed font-sans">
                      {gemmaResponse}
                    </div>
                  </div>
                )}
              </div>
            </div>
          </div>
        )}
      </main>

      {/* Sticky Bottom Batch Toolbar for Cutlist & Rip Cuts Tab */}
      {selectedSheetIds.size > 0 && (
        <div className="fixed bottom-4 left-1/2 -translate-x-1/2 w-[95%] max-w-[1240px] bg-[#0c1a30] border-2 border-sky-500 p-3 sm:p-4 rounded-xl flex flex-wrap items-center justify-between gap-3 shadow-2xl z-50 animate-in slide-in-from-bottom-4">
          <div>
            <div className="font-black text-sm sm:text-base text-sky-300 flex items-center gap-2">
              <span>{selectedSheetIds.size} Sheets Selected</span>
              <span className="bg-sky-500/20 text-sky-200 border border-sky-500/40 text-[10px] px-2 py-0.5 rounded-full font-mono">
                Cutlist Batch Operations
              </span>
            </div>
            <div className="text-[11px] text-slate-400 font-medium mt-0.5">
              Change status, assign saw, set cycle time, or print 100mm × 50mm thermal labels in bulk
            </div>
          </div>

          <div className="flex items-center gap-2 flex-wrap">
            {/* Batch Print Labels */}
            <button
              onClick={handlePrintSelected}
              className="px-3 py-1.5 bg-sky-600 hover:bg-sky-500 text-white rounded-lg text-xs font-black flex items-center gap-1.5 transition shadow"
            >
              <Printer className="w-3.5 h-3.5" />
              <span>🖨️ Print Labels ({selectedSheetIds.size})</span>
            </button>

            {/* Change Status */}
            <select
              value={cutlistBatchStatus}
              onChange={(e) => setCutlistBatchStatus(e.target.value)}
              className="bg-[#0b1325] border border-[#223354] rounded-lg px-2.5 py-1.5 text-xs font-bold text-amber-400 focus:outline-none focus:border-sky-500"
            >
              <option value="">-- Status --</option>
              <option value="QUEUED">Set QUEUED</option>
              <option value="ASSIGNED">Set ASSIGNED</option>
              <option value="CUTTING">Set CUTTING</option>
              <option value="CUT">Set CUT (Completed)</option>
            </select>

            {/* Set Time Taken */}
            <input
              type="number"
              placeholder="Time (min)"
              value={cutlistBatchDuration}
              onChange={(e) => setCutlistBatchDuration(e.target.value)}
              className="w-24 bg-[#0b1325] border border-[#223354] rounded-lg px-2 py-1.5 text-xs text-emerald-400 placeholder-slate-500 font-bold focus:outline-none"
              step="0.5"
              min="0"
            />

            {(cutlistBatchStatus || cutlistBatchDuration !== '') && (
              <button
                onClick={handleApplyCutlistBatchEdit}
                disabled={isCutlistBatchUpdating}
                className="px-3 py-1.5 bg-emerald-600 hover:bg-emerald-500 disabled:opacity-40 text-white rounded-lg text-xs font-bold shadow flex items-center gap-1 transition"
              >
                <Save className="w-3.5 h-3.5" />
                <span>{isCutlistBatchUpdating ? 'Saving...' : 'Apply Edit'}</span>
              </button>
            )}

            {/* Quick Assign Buttons */}
            <button
              onClick={() => handleBatchAssign('Saw 1', 'Team A (Ahmed)')}
              className="px-2.5 py-1.5 bg-[#131d33] hover:bg-[#1c2a4a] text-slate-200 border border-[#223354] rounded-lg text-xs font-bold transition"
            >
              👉 Saw 1 (Team A)
            </button>

            <button
              onClick={() => handleBatchAssign('Saw 2', 'Team B')}
              className="px-2.5 py-1.5 bg-[#131d33] hover:bg-[#1c2a4a] text-slate-200 border border-[#223354] rounded-lg text-xs font-bold transition"
            >
              👉 Saw 2 (Team B)
            </button>

            <button
              onClick={() => setSelectedSheetIds(new Set())}
              className="p-1.5 text-slate-400 hover:text-white transition"
              title="Clear selection"
            >
              <X className="w-4 h-4" />
            </button>
          </div>
        </div>
      )}

      {/* Diagram Inspection Modal */}
      {diagramModalSheet && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/90 backdrop-blur-md p-4">
          <div className="bg-[#131d33] border border-[#223354] rounded-2xl shadow-2xl w-full max-w-5xl max-h-[92vh] flex flex-col overflow-hidden">
            <div className="p-4 border-b border-[#223354] flex justify-between items-center bg-[#0f172a]">
              <div>
                <h3 className="text-base font-black text-white flex items-center gap-2">
                  <Eye className="w-5 h-5 text-sky-400" />
                  <span>Cutting Diagram &amp; Rip Sequence: {diagramModalSheet.sheet_title}</span>
                </h3>
                <p className="text-xs text-slate-400 mt-0.5 flex items-center gap-2 flex-wrap">
                  <span className="font-mono text-sky-400">{diagramModalSheet.sheet_id}</span>
                  <span>•</span>
                  <span>{diagramModalSheet.sheet_dim}</span>
                  <span>•</span>
                  <span className="text-emerald-400 font-bold">Efficiency: {diagramModalSheet.efficiency}</span>
                  <span>•</span>
                  <span className="text-amber-400 font-bold">Cut Length: {diagramModalSheet.cut_length || '28.0 m'}</span>
                  <span>•</span>
                  <span className="text-slate-300 font-semibold">{diagramModalSheet.grain || 'Grain follows length'}</span>
                </p>
              </div>

              <div className="flex items-center gap-2">
                <a
                  href={diagramModalSheet.diagram_url}
                  target="_blank"
                  rel="noreferrer"
                  className="px-3 py-1.5 bg-[#0b1325] hover:bg-[#1c2a4a] text-slate-200 border border-[#223354] rounded-lg text-xs font-bold flex items-center gap-1.5 transition"
                >
                  <ExternalLink className="w-3.5 h-3.5 text-sky-400" />
                  <span>Open Raw SVG</span>
                </a>
                <button
                  onClick={() => setDiagramModalSheet(null)}
                  className="p-1.5 text-slate-400 hover:text-white rounded-lg transition"
                >
                  <X className="w-5 h-5" />
                </button>
              </div>
            </div>

            {/* Diagram Viewer Body */}
            <div className="p-4 overflow-y-auto flex-1 bg-[#070d19] flex flex-col items-center justify-center">
              {loadingSvg ? (
                <div className="p-12 text-slate-400 text-sm font-semibold flex items-center gap-2">
                  <RefreshCw className="w-5 h-5 animate-spin text-sky-400" />
                  <span>Loading Cutting Pattern Diagram...</span>
                </div>
              ) : rawSvgContent ? (
                <div
                  className="w-full flex justify-center bg-[#0b1325] p-3 rounded-xl border border-[#223354] shadow-inner overflow-hidden max-h-[64vh]"
                  dangerouslySetInnerHTML={{ __html: rawSvgContent }}
                />
              ) : diagramModalSheet.diagram_url ? (
                <div className="w-full flex justify-center bg-[#0b1325] p-3 rounded-xl border border-[#223354] shadow-inner">
                  <img
                    src={diagramModalSheet.diagram_url}
                    alt={diagramModalSheet.sheet_title}
                    className="max-w-full max-h-[60vh] object-contain rounded"
                  />
                </div>
              ) : (
                <div className="p-12 text-slate-500 text-sm font-semibold">
                  No SVG cutting diagram available for this sheet pattern.
                </div>
              )}

              {/* Primary Rip Cuts Sequence Footer */}
              <div className="w-full mt-3 p-3 bg-[#0f172a] border border-[#223354] rounded-xl flex items-center justify-between flex-wrap gap-2 text-xs">
                <div>
                  <span className="font-extrabold text-slate-300 uppercase tracking-wider text-[11px] block">
                    Table Saw Primary Rip Cut Sequence:
                  </span>
                  <div className="flex flex-wrap gap-1.5 mt-1">
                    {diagramModalSheet.rip_cuts && diagramModalSheet.rip_cuts.length > 0 ? (
                      diagramModalSheet.rip_cuts.map((rc, idx) => (
                        <span key={idx} className="rip-cut-tag font-bold text-xs py-1 px-2.5">
                          Cut {idx + 1}: {rc}
                        </span>
                      ))
                    ) : (
                      <span className="text-slate-400 italic">Follow pattern diagram labels</span>
                    )}
                  </div>
                </div>

                <div className="text-right">
                  <div className="text-slate-400 text-[11px]">Assigned to:</div>
                  <div className="font-bold text-white text-xs">{diagramModalSheet.assigned_saw} · {diagramModalSheet.assigned_team}</div>
                </div>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* Quick Edit Sheet Modal */}
      {editingSheet && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/85 backdrop-blur-sm p-4">
          <div className="bg-[#131d33] border border-[#223354] rounded-xl shadow-2xl w-full max-w-lg p-5 space-y-4">
            <div className="flex justify-between items-start border-b border-[#223354] pb-3">
              <div>
                <h3 className="text-base font-extrabold text-white flex items-center gap-2">
                  <Edit className="w-4 h-4 text-amber-400" />
                  <span>Update Sheet Record: {editingSheet.sheet_title}</span>
                </h3>
                <p className="text-xs text-slate-400 mt-0.5">
                  ID: <span className="font-mono text-sky-400">{editingSheet.sheet_id}</span> • {editingSheet.thickness || 18}mm • ×{editingSheet.repeat_count} boards
                </p>
              </div>
              <button
                onClick={() => setEditingSheet(null)}
                className="p-1 text-slate-400 hover:text-white rounded transition"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            <div className="space-y-3 text-xs">
              <div>
                <label className="block font-bold text-slate-300 mb-1">
                  Cutting Status:
                </label>
                <select
                  value={editStatus}
                  onChange={(e) => setEditStatus(e.target.value)}
                  className="w-full bg-[#0b1325] border border-[#223354] rounded-lg px-3 py-2 text-white font-bold focus:outline-none focus:border-sky-500"
                >
                  <option value="QUEUED">QUEUED (Pending)</option>
                  <option value="ASSIGNED">ASSIGNED</option>
                  <option value="CUTTING">CUTTING NOW (In Progress)</option>
                  <option value="CUT">CUT (Completed)</option>
                </select>
              </div>

              <div>
                <label className="block font-bold text-slate-300 mb-1">
                  Actual Time Taken (Minutes):
                </label>
                <input
                  type="number"
                  step="0.1"
                  min="0"
                  placeholder="e.g. 7.5 min"
                  value={editDuration}
                  onChange={(e) => setEditDuration(e.target.value)}
                  className="w-full bg-[#0b1325] border border-[#223354] rounded-lg px-3 py-2 text-white font-mono focus:outline-none focus:border-sky-500"
                />
                <span className="text-[10.5px] text-slate-500 mt-0.5 block">
                  Leave empty to keep current estimate or enter real clocked time.
                </span>
              </div>

              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="block font-bold text-slate-300 mb-1">
                    Assigned Saw:
                  </label>
                  <select
                    value={editSaw}
                    onChange={(e) => setEditSaw(e.target.value)}
                    className="w-full bg-[#0b1325] border border-[#223354] rounded-lg px-3 py-2 text-white font-bold focus:outline-none focus:border-sky-500"
                  >
                    <option value="Saw 1">Saw 1</option>
                    <option value="Saw 2">Saw 2</option>
                  </select>
                </div>

                <div>
                  <label className="block font-bold text-slate-300 mb-1">
                    Assigned Team:
                  </label>
                  <select
                    value={editTeam}
                    onChange={(e) => setEditTeam(e.target.value)}
                    className="w-full bg-[#0b1325] border border-[#223354] rounded-lg px-3 py-2 text-white font-bold focus:outline-none focus:border-sky-500"
                  >
                    <option value="Team A (Ahmed)">Team A (Ahmed)</option>
                    <option value="Team B">Team B</option>
                    <option value="Team B (Under Repair)">Team B (Under Repair)</option>
                  </select>
                </div>
              </div>
            </div>

            <div className="flex justify-end gap-2 pt-3 border-t border-[#223354]">
              <button
                onClick={() => setEditingSheet(null)}
                className="px-3.5 py-1.5 bg-[#0b1325] hover:bg-[#1c2a4a] text-slate-300 rounded-lg text-xs font-semibold"
              >
                Cancel
              </button>
              <button
                onClick={handleSaveSheetEdit}
                disabled={isUpdatingSheet}
                className="px-4 py-1.5 bg-emerald-600 hover:bg-emerald-500 disabled:opacity-40 text-white rounded-lg text-xs font-bold shadow flex items-center gap-1.5"
              >
                <Save className="w-3.5 h-3.5" />
                <span>{isUpdatingSheet ? 'Saving...' : 'Save & Update Database'}</span>
              </button>
            </div>
          </div>
        </div>
      )}

      {/* QR Code Camera Scanner Modal */}
      <QrScannerModal
        isOpen={isQrScannerOpen}
        onClose={() => setIsQrScannerOpen(false)}
        sheets={serverState?.sheets || []}
        onSelectSheet={(sheet) => {
          setDiagramModalSheet(sheet);
          setIsQrScannerOpen(false);
        }}
        onToggleCut={(sheetId, status) => {
          handleToggleCut(sheetId, status as any);
        }}
        onOpenEditSheet={(sheet) => {
          handleOpenEditSheet(sheet);
        }}
      />

      {/* Admin Passcode Unlock Modal */}
      {showPasscodeModal && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/85 backdrop-blur-sm p-4">
          <div className="bg-[#131d33] border border-[#223354] rounded-xl shadow-2xl w-full max-w-md p-5 space-y-4">
            <div className="flex items-center gap-3">
              <div className="p-2.5 bg-amber-500/20 text-amber-400 rounded-lg border border-amber-500/30">
                <KeyRound className="w-5 h-5" />
              </div>
              <div>
                <h3 className="text-sm font-bold text-white">Unlock Admin Edit Access</h3>
                <p className="text-xs text-slate-400">
                  Enter Admin IP (192.168.1.173) or manager PIN to authorize saw edits
                </p>
              </div>
            </div>

            <div className="space-y-2">
              <input
                type="password"
                value={passcodeInput}
                onChange={(e) => setPasscodeInput(e.target.value)}
                onKeyDown={(e) => {
                  if (e.key === 'Enter') handleUnlockWithPasscode();
                }}
                placeholder="Enter 192.168.1.173 or admin..."
                className="w-full bg-[#0b1325] border border-[#223354] rounded-lg px-3 py-2 text-xs text-white focus:outline-none focus:border-sky-500 font-mono"
                autoFocus
              />
              {passcodeError && (
                <p className="text-xs text-rose-400 font-semibold">{passcodeError}</p>
              )}
            </div>

            <div className="flex justify-end gap-2 pt-2 border-t border-[#223354]">
              <button
                onClick={() => {
                  setShowPasscodeModal(false);
                  setPasscodeInput('');
                  setPasscodeError('');
                }}
                className="px-3 py-1.5 bg-[#0b1325] hover:bg-[#1c2a4a] text-slate-300 rounded-lg text-xs font-semibold"
              >
                Cancel
              </button>
              <button
                onClick={handleUnlockWithPasscode}
                className="px-4 py-1.5 bg-sky-600 hover:bg-sky-500 text-white rounded-lg text-xs font-bold shadow"
              >
                Authorize
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Print Preview Modal */}
      {isPrintModalOpen && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/85 backdrop-blur-sm p-4">
          <div className="bg-[#131d33] border border-[#223354] rounded-xl shadow-2xl w-full max-w-4xl max-h-[90vh] flex flex-col overflow-hidden">
            <div className="p-4 border-b border-[#223354] flex justify-between items-center bg-[#0f172a]">
              <div>
                <h3 className="text-sm font-bold text-white flex items-center gap-2">
                  <Printer className="w-4 h-4 text-sky-400" />
                  <span>{printModalTitle}</span>
                </h3>
                <p className="text-xs text-slate-400 mt-0.5">
                  Standard 100mm × 50mm Industrial Thermal Label format with machine barcodes
                </p>
              </div>

              <div className="flex items-center gap-2">
                <button
                  onClick={handleExportPdf}
                  disabled={isExportingPdf}
                  className="px-4 py-1.5 bg-emerald-600 hover:bg-emerald-500 text-white rounded-lg text-xs font-bold flex items-center gap-1.5 transition shadow disabled:opacity-50"
                >
                  <FileDown className="w-3.5 h-3.5" />
                  <span>{isExportingPdf ? 'Generating PDF...' : 'Export to PDF (100x50mm)'}</span>
                </button>
                <button
                  onClick={() => window.print()}
                  className="px-4 py-1.5 bg-sky-600 hover:bg-sky-500 text-white rounded-lg text-xs font-bold flex items-center gap-1.5 transition shadow"
                >
                  <Printer className="w-3.5 h-3.5" />
                  <span>Print Labels Now</span>
                </button>
                <button
                  onClick={() => setIsPrintModalOpen(false)}
                  className="p-1.5 text-slate-400 hover:text-white rounded transition"
                >
                  <X className="w-5 h-5" />
                </button>
              </div>
            </div>

            {/* Sticker Preview Grid */}
            <div className="p-4 overflow-y-auto flex-1 bg-slate-950 flex flex-wrap gap-3 justify-center">
              {labelsToPrint.map((lbl, idx) => (
                <div key={idx} className="scale-95 origin-center">
                  <TableSawThermalLabel
                    part={lbl.part}
                    sheetStart={lbl.sheetStart}
                    pieceIndex={lbl.pieceIndex}
                    totalPieces={lbl.totalPieces}
                    printMode={false}
                  />
                </div>
              ))}
            </div>
          </div>
        </div>
      )}
      </div>

      {/* Hidden Thermal Print Mount Point for browser window.print() */}
      <div id="print-area" className="hidden print:block">
        {labelsToPrint.map((lbl, idx) => (
          <TableSawThermalLabel
            key={idx}
            part={lbl.part}
            sheetStart={lbl.sheetStart}
            pieceIndex={lbl.pieceIndex}
            totalPieces={lbl.totalPieces}
            printMode={true}
          />
        ))}
      </div>
    </>
  );
}
