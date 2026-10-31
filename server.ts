import express from 'express';
import { createServer as createViteServer } from 'vite';
import { GoogleGenAI } from '@google/genai';
import dotenv from 'dotenv';
import fs from 'fs';
import path from 'path';
import os from 'os';
import { fileURLToPath } from 'url';
import { execSync } from 'child_process';
import * as XLSX from 'xlsx';

dotenv.config();

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const PORT = 3000;
const ADMIN_IP = '192.168.1.173';
const STATE_FILE = path.join(__dirname, 'bazar_factory_saw_state.json');
const SHIFT_MINUTES_PER_DAY = 420; // 8:00 AM - 12:00 PM (240m) + 1:00 PM - 4:00 PM (180m) = 420m (7h)

// Initialize GoogleGenAI SDK on server
const ai = new GoogleGenAI({
  apiKey: process.env.GEMINI_API_KEY || '',
  httpOptions: {
    headers: {
      'User-Agent': 'aistudio-build',
    }
  }
});

const app = express();
app.use(express.json({ limit: '20mb' }));

// Helper: Normalize Client IP
function getClientIp(req: express.Request): string {
  let ip = (req.headers['x-forwarded-for'] as string) || req.socket.remoteAddress || '';
  if (ip.includes(',')) ip = ip.split(',')[0].trim();
  if (ip.startsWith('::ffff:')) ip = ip.replace('::ffff:', '');
  if (ip === '::1' || ip === '127.0.0.1' || ip === 'localhost') return '127.0.0.1';
  return ip;
}

function isAuthorizedAdmin(req: express.Request): boolean {
  const ip = getClientIp(req);
  const adminOverrideHeader = req.headers['x-admin-key'];
  if (adminOverrideHeader === 'admin-192-168-1-173') return true;
  return ip === ADMIN_IP || ip === '127.0.0.1';
}

// In-Memory Saw Factory State
interface SawStatus {
  status: string;
  team: string;
  assigned_sheets: number;
  cut_sheets: number;
}

interface PartItem {
  code: string;
  name: string;
  pdf_set: string;
  core: string;
  qty_per_sheet: number;
  total_batch_qty: number;
  l: number;
  w: number;
  t: number;
  cab: string;
  villa: string;
  lip: string;
}

interface SheetItem {
  sheet_id: string;
  sheet_title: string;
  thickness?: number;
  start_sheet: number;
  end_sheet: number;
  repeat_count: number;
  total_sheets: number;
  material: string;
  sheet_dim: string;
  efficiency: string;
  cut_length?: string;
  grain?: string;
  diagram_url?: string;
  rip_cuts: string[];
  pdf_sets: string[];
  parts: PartItem[];
  assigned_saw: string;
  assigned_team: string;
  status: 'QUEUED' | 'ASSIGNED' | 'CUTTING' | 'CUT';
  queued_at?: number | null;
  start_time?: number | null;
  cut_time?: number | null;
  duration_min?: number | null;
}

interface MaterialSummaryItem {
  material: string;
  parts_qty: number;
  rough_length: string;
  rough_area: string;
  rough_volume: string;
}

interface OvertimeEntry {
  date: string;
  team: string;
  hours: number;
  logged_at: string;
}

interface AuditEvent {
  ts: number;
  action: string;
  sheet_id?: string;
  saw?: string;
  duration?: number | null;
  op: string;
  note?: string;
}

let saws_status: Record<string, SawStatus> = {
  'Saw 1': { status: 'ONLINE', team: 'Team A (Ahmed)', assigned_sheets: 0, cut_sheets: 0 },
  'Saw 2': { status: 'IN_REPAIR', team: 'Team B (Under Repair)', assigned_sheets: 0, cut_sheets: 0 }
};

let sheets_registry: SheetItem[] = [];
let audit_events: AuditEvent[] = [];
let overtime_history: OvertimeEntry[] = [];
let material_summary_list: MaterialSummaryItem[] = [];

// Trigger Parser Script if needed
function ensureCutlistParsed(force = false) {
  if (force || !fs.existsSync(STATE_FILE)) {
    console.log('[PARSER] Executing python3 scripts/parse_cutlist.py ...');
    try {
      execSync('python3 scripts/parse_cutlist.py', { cwd: __dirname, stdio: 'inherit' });
    } catch (e: any) {
      console.error('[PARSER ERROR]', e.message);
    }
  }
}

function loadState() {
  ensureCutlistParsed();
  if (fs.existsSync(STATE_FILE)) {
    try {
      const raw = fs.readFileSync(STATE_FILE, 'utf-8');
      const data = JSON.parse(raw);
      if (data.saws_status) saws_status = data.saws_status;
      if (data.sheets && Array.isArray(data.sheets)) {
        // Ensure every sheet has a valid queued_at baseline (default 48h ago for initial queued items so foreman sees existing backlog)
        const defaultBaselineTime = Math.floor(Date.now() / 1000) - 36 * 3600;
        sheets_registry = data.sheets.map((s: SheetItem) => ({
          ...s,
          queued_at: s.queued_at || defaultBaselineTime
        }));
      }
      if (data.material_summary && Array.isArray(data.material_summary)) material_summary_list = data.material_summary;
      if (data.audit && Array.isArray(data.audit)) audit_events = data.audit;
      if (data.overtime && Array.isArray(data.overtime)) overtime_history = data.overtime;
      console.log(`[STATE] Loaded ${sheets_registry.length} sheets from ${STATE_FILE}`);
    } catch (e: any) {
      console.error('[STATE LOAD ERROR]', e.message);
    }
  }
}

function saveState() {
  try {
    const payload = {
      saws_status,
      sheets: sheets_registry,
      material_summary: material_summary_list,
      audit: audit_events.slice(-500),
      overtime: overtime_history,
      savedAt: new Date().toISOString()
    };
    fs.writeFileSync(STATE_FILE, JSON.stringify(payload, null, 2), 'utf-8');
  } catch (e: any) {
    console.error('[STATE SAVE ERROR]', e.message);
  }
}

// Compute PDF Completion Matrix
function buildPdfMatrix() {
  const descMap: Record<string, string> = {
    'PDF 1': 'SCMB (Side Cabinet)',
    'PDF 2': 'CB-03-SF (Bed 03)',
    'PDF 3': 'CB-02-SF (Bed 02)',
    'PDF 4': 'SVMB (Master Bedroom)',
    'PDF 5': 'CB-01 & Drawers'
  };

  const matrix: Record<string, {
    name: string;
    description: string;
    total_sheets: Set<string>;
    cut_sheets: Set<string>;
    pending_sheets: Set<string>;
    total_parts: number;
    cut_parts: number;
  }> = {};

  for (let i = 1; i <= 5; i++) {
    const k = `PDF ${i}`;
    matrix[k] = {
      name: k,
      description: descMap[k] || `Drawing Set ${i}`,
      total_sheets: new Set<string>(),
      cut_sheets: new Set<string>(),
      pending_sheets: new Set<string>(),
      total_parts: 0,
      cut_parts: 0
    };
  }

  sheets_registry.forEach((s) => {
    const isCut = s.status === 'CUT';
    s.parts.forEach((p) => {
      const pdf = p.pdf_set || 'PDF 5';
      if (!matrix[pdf]) {
        matrix[pdf] = {
          name: pdf,
          description: descMap[pdf] || pdf,
          total_sheets: new Set<string>(),
          cut_sheets: new Set<string>(),
          pending_sheets: new Set<string>(),
          total_parts: 0,
          cut_parts: 0
        };
      }
      matrix[pdf].total_sheets.add(s.sheet_id);
      matrix[pdf].total_parts += p.total_batch_qty;
      if (isCut) {
        matrix[pdf].cut_sheets.add(s.sheet_id);
        matrix[pdf].cut_parts += p.total_batch_qty;
      } else {
        matrix[pdf].pending_sheets.add(s.sheet_id);
      }
    });
  });

  const result: Record<string, any> = {};
  for (const [k, v] of Object.entries(matrix)) {
    const tSheets = v.total_sheets.size;
    const cSheets = v.cut_sheets.size;
    const pSheets = v.pending_sheets.size;
    const pct = tSheets > 0 ? Math.round((cSheets / tSheets) * 1000) / 10 : 0;
    result[k] = {
      name: v.name,
      description: v.description,
      total_sheets: tSheets,
      cut_sheets: cSheets,
      pending_sheets: pSheets,
      blocking_sheet_ids: Array.from(v.pending_sheets),
      total_parts: v.total_parts,
      cut_parts: v.cut_parts,
      percent_complete: pct
    };
  }
  return result;
}

// Initial state load
loadState();

// ==========================================
// API ROUTES
// ==========================================

// 1. Client Authorization & IP check
app.get('/api/client-info', (req, res) => {
  const ip = getClientIp(req);
  const isAdmin = isAuthorizedAdmin(req);
  res.json({
    clientIp: ip,
    adminIp: ADMIN_IP,
    isAdmin,
    mode: isAdmin ? 'ADMIN_FULL_ACCESS' : 'OPERATOR_SUPERVISED_READ_ONLY'
  });
});

// 2. Factory Status & Analytics
app.get('/api/saw-factory-status', (req, res) => {
  const clientIp = getClientIp(req);
  const isAdmin = isAuthorizedAdmin(req);

  // Compute average cycle time
  const durations = sheets_registry
    .map((s) => s.duration_min)
    .filter((d): d is number => typeof d === 'number' && d > 0);
  const avg_cycle_min = durations.length
    ? Math.round((durations.reduce((a, b) => a + b, 0) / durations.length) * 10) / 10
    : 8.5;

  // Saws stats
  const s1_assigned = sheets_registry
    .filter((s) => s.assigned_saw === 'Saw 1')
    .reduce((sum, s) => sum + s.repeat_count, 0);
  const s1_cut = sheets_registry
    .filter((s) => s.assigned_saw === 'Saw 1' && s.status === 'CUT')
    .reduce((sum, s) => sum + s.repeat_count, 0);

  const s2_assigned = sheets_registry
    .filter((s) => s.assigned_saw === 'Saw 2')
    .reduce((sum, s) => sum + s.repeat_count, 0);
  const s2_cut = sheets_registry
    .filter((s) => s.assigned_saw === 'Saw 2' && s.status === 'CUT')
    .reduce((sum, s) => sum + s.repeat_count, 0);

  saws_status['Saw 1'].assigned_sheets = s1_assigned;
  saws_status['Saw 1'].cut_sheets = s1_cut;
  saws_status['Saw 2'].assigned_sheets = s2_assigned;
  saws_status['Saw 2'].cut_sheets = s2_cut;

  const active_saws_count = saws_status['Saw 2'].status === 'IN_REPAIR' ? 1 : 2;
  const total_pending_sheets = sheets_registry
    .filter((s) => s.status !== 'CUT')
    .reduce((sum, s) => sum + s.repeat_count, 0);
  const total_minutes_needed = (total_pending_sheets * avg_cycle_min) / active_saws_count;
  const total_days_needed = Math.round((total_minutes_needed / SHIFT_MINUTES_PER_DAY) * 10) / 10;

  const pdf_matrix = buildPdfMatrix();

  res.json({
    saws: saws_status,
    sheets: sheets_registry,
    material_summary: material_summary_list,
    pdf_matrix,
    analytics: {
      avg_cycle_min,
      total_pending_sheets,
      total_minutes_needed: Math.round(total_minutes_needed),
      total_days_needed,
      active_saws: active_saws_count,
      shift_hours: '8:00 AM - 12:00 PM, 1:00 PM - 4:00 PM (7h regular)'
    },
    overtime_history,
    audit: audit_events.slice(-100),
    clientIp,
    adminIp: ADMIN_IP,
    isAdmin
  });
});

// 3. Export Progress CSV
app.get('/api/export-progress-csv', (req, res) => {
  res.setHeader('Content-Type', 'text/csv; charset=utf-8');
  res.setHeader('Content-Disposition', 'attachment; filename="factory_saw_production_log.csv"');

  let csvOut =
    'Sheet ID,Range,Thickness,Repeat Qty,Material,Status,Assigned Saw,Assigned Team,Primary Rip Cuts,PDF Sets,Duration (min),Started At,Cut Completed At\n';
  for (const s of sheets_registry) {
    const rips = (s.rip_cuts || []).join(' | ');
    const pdfs = (s.pdf_sets || []).join(' | ');
    const st_time = s.start_time ? new Date(s.start_time * 1000).toISOString() : '';
    const end_time = s.cut_time ? new Date(s.cut_time * 1000).toISOString() : '';
    const dur = s.duration_min !== null && s.duration_min !== undefined ? s.duration_min : '';
    const thick = s.thickness ? `${s.thickness}mm` : '18mm';
    csvOut += `"${s.sheet_id}","${s.sheet_title}","${thick}",${s.repeat_count},"${s.material}","${s.status}","${s.assigned_saw}","${s.assigned_team}","${rips}","${pdfs}",${dur},"${st_time}","${end_time}"\n`;
  }
  res.send(csvOut);
});

// 3B. Export Professional Master Excel (.xlsx) Database
app.get('/api/export-master-excel', (req, res) => {
  try {
    const wb = XLSX.utils.book_new();

    // Sheet 1: Master Production Log
    const sheetRows = sheets_registry.map((s) => ({
      'Sheet ID': s.sheet_id,
      'Sheet Title': s.sheet_title,
      'Thickness (mm)': s.thickness || 18,
      'Repeat Count (Boards)': s.repeat_count,
      'Material': s.material,
      'Dimensions': s.sheet_dim,
      'Efficiency': s.efficiency,
      'Length of Cuts': s.cut_length || '28.0 m',
      'Grain Direction': s.grain || 'Grain follows length',
      'Status': s.status,
      'Assigned Saw': s.assigned_saw,
      'Assigned Team': s.assigned_team,
      'Primary Rip Cuts': (s.rip_cuts || []).join(', '),
      'PDF Drawing Sets': (s.pdf_sets || []).join(', '),
      'Duration (Minutes)': s.duration_min ?? '',
      'Started At': s.start_time ? new Date(s.start_time * 1000).toISOString() : '',
      'Completed At': s.cut_time ? new Date(s.cut_time * 1000).toISOString() : ''
    }));
    const wsSheets = XLSX.utils.json_to_sheet(sheetRows);
    XLSX.utils.book_append_sheet(wb, wsSheets, 'Saw Production Sheets');

    // Sheet 2: Global Material Summary & Rough Bill of Materials (BOM)
    const summaryRows = material_summary_list.map((m) => ({
      'Material & Thickness': m.material,
      'Total Component Parts': m.parts_qty,
      'Total Rough Length': m.rough_length,
      'Total Rough Area': m.rough_area,
      'Total Rough Volume': m.rough_volume
    }));
    const wsSummary = XLSX.utils.json_to_sheet(summaryRows);
    XLSX.utils.book_append_sheet(wb, wsSummary, 'Material Rough BOM');

    // Sheet 3: Parts Breakdown Database
    const partRows: any[] = [];
    sheets_registry.forEach((s) => {
      s.parts.forEach((p) => {
        partRows.push({
          'Sheet ID': s.sheet_id,
          'Part Code': p.code,
          'Part Name': p.name,
          'PDF Set': p.pdf_set,
          'Cabinet Group': p.cab,
          'Villa': p.villa,
          'Length (mm)': p.l,
          'Width (mm)': p.w,
          'Thickness (mm)': p.t,
          'Qty Per Sheet': p.qty_per_sheet,
          'Total Batch Qty': p.total_batch_qty,
          'Edge Lip': p.lip,
          'Core Specs': p.core,
          'Sheet Status': s.status,
          'Assigned Saw': s.assigned_saw
        });
      });
    });
    const wsParts = XLSX.utils.json_to_sheet(partRows);
    XLSX.utils.book_append_sheet(wb, wsParts, 'All Parts Database');

    // Sheet 3: Overtime History
    const wsOvertime = XLSX.utils.json_to_sheet(
      overtime_history.map((o) => ({
        'Date': o.date,
        'Team': o.team,
        'Overtime Hours': o.hours,
        'Logged At': o.logged_at
      }))
    );
    XLSX.utils.book_append_sheet(wb, wsOvertime, 'Overtime Chrono');

    const buffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename="BAZAR_V3_Saw_Factory_Master_DB.xlsx"');
    res.send(buffer);
  } catch (err: any) {
    console.error('[EXCEL EXPORT ERROR]', err);
    res.status(500).json({ error: 'Failed to generate Excel database' });
  }
});

// 3C. Update Single Sheet Parameters (Status, Time Taken, Assigned Saw/Team, Notes)
app.post('/api/saw/update-sheet', (req, res) => {
  const clientIp = getClientIp(req);
  if (!isAuthorizedAdmin(req)) {
    return res.status(403).json({
      success: false,
      error: `ACCESS DENIED: Your IP (${clientIp}) is view-only. Only Admin (${ADMIN_IP}) has permission to edit.`
    });
  }

  const { sheet_id, status, duration_min, assigned_saw, assigned_team } = req.body;
  const s = sheets_registry.find((x) => x.sheet_id === sheet_id);
  if (!s) {
    return res.status(404).json({ success: false, error: 'Sheet not found' });
  }

  const oldStatus = s.status;
  if (status !== undefined) {
    s.status = status;
    if (status === 'QUEUED' && oldStatus !== 'QUEUED') {
      s.queued_at = Math.floor(Date.now() / 1000);
    }
  }
  if (duration_min !== undefined) s.duration_min = parseFloat(duration_min) || null;
  if (assigned_saw !== undefined) s.assigned_saw = assigned_saw;
  if (assigned_team !== undefined) s.assigned_team = assigned_team;

  if (s.status === 'CUT' && !s.cut_time) {
    s.cut_time = Math.floor(Date.now() / 1000);
  } else if (s.status === 'CUTTING' && !s.start_time) {
    s.start_time = Math.floor(Date.now() / 1000);
  }

  audit_events.push({
    ts: Date.now(),
    action: 'SHEET_MANUALLY_EDITED',
    sheet_id,
    saw: s.assigned_saw,
    duration: s.duration_min,
    op: `ADMIN (${ADMIN_IP})`,
    note: `Status ${oldStatus} -> ${s.status}, Duration: ${s.duration_min} min`
  });

  saveState();
  res.json({ success: true, sheet: s });
});

// 3D. Bulk Import / Update Database from CSV or JSON Rows
app.post('/api/saw/bulk-import-updates', (req, res) => {
  const clientIp = getClientIp(req);
  if (!isAuthorizedAdmin(req)) {
    return res.status(403).json({
      success: false,
      error: `ACCESS DENIED: Your IP (${clientIp}) is view-only. Only Admin (${ADMIN_IP}) has permission to edit.`
    });
  }

  const { updates = [] } = req.body; // Array of { sheet_id, status, duration_min, assigned_saw, assigned_team }
  let updatedCount = 0;

  for (const item of updates) {
    if (!item.sheet_id) continue;
    const s = sheets_registry.find((x) => x.sheet_id === item.sheet_id || x.sheet_title === item.sheet_id);
    if (s) {
      if (item.status && ['QUEUED', 'ASSIGNED', 'CUTTING', 'CUT'].includes(item.status.toUpperCase())) {
        s.status = item.status.toUpperCase();
        if (s.status === 'CUT' && !s.cut_time) s.cut_time = Math.floor(Date.now() / 1000);
      }
      if (item.duration_min !== undefined && item.duration_min !== '' && !isNaN(parseFloat(item.duration_min))) {
        s.duration_min = parseFloat(item.duration_min);
      }
      if (item.assigned_saw) s.assigned_saw = item.assigned_saw;
      if (item.assigned_team) s.assigned_team = item.assigned_team;
      updatedCount++;
    }
  }

  audit_events.push({
    ts: Date.now(),
    action: 'DATABASE_BULK_SYNCED',
    op: `ADMIN (${ADMIN_IP})`,
    note: `Bulk updated ${updatedCount} sheets from Excel/CSV sync`
  });

  saveState();
  res.json({ success: true, count: updatedCount });
});

// 4. Start Cutting Timer
app.post('/api/saw/start-cut', (req, res) => {
  const clientIp = getClientIp(req);
  if (!isAuthorizedAdmin(req)) {
    return res.status(403).json({
      success: false,
      error: `ACCESS DENIED: Your IP (${clientIp}) is view-only. Only Admin (${ADMIN_IP}) has permission to edit.`
    });
  }

  const { sheet_id } = req.body;
  const s = sheets_registry.find((x) => x.sheet_id === sheet_id);
  if (s) {
    s.status = 'CUTTING';
    s.start_time = Math.floor(Date.now() / 1000);
    saveState();
  }
  res.json({ success: true });
});

// 5. Finish Cut
app.post('/api/saw/finish-cut', (req, res) => {
  const clientIp = getClientIp(req);
  if (!isAuthorizedAdmin(req)) {
    return res.status(403).json({
      success: false,
      error: `ACCESS DENIED: Your IP (${clientIp}) is view-only. Only Admin (${ADMIN_IP}) has permission to edit.`
    });
  }

  const { sheet_id } = req.body;
  const s = sheets_registry.find((x) => x.sheet_id === sheet_id);
  if (s) {
    s.status = 'CUT';
    s.cut_time = Math.floor(Date.now() / 1000);
    if (s.start_time) {
      const elapsedMin = Math.round(((s.cut_time - s.start_time) / 60) * 10) / 10;
      s.duration_min = Math.max(elapsedMin, 1.0);
    } else {
      s.duration_min = 8.0;
    }

    audit_events.push({
      ts: Date.now(),
      action: 'SHEET_CUT_COMPLETED',
      sheet_id,
      saw: s.assigned_saw,
      duration: s.duration_min,
      op: `ADMIN (${ADMIN_IP})`
    });
    saveState();
  }
  res.json({ success: true });
});

// 6. Batch Assignment
app.post('/api/saw/batch-assign', (req, res) => {
  const clientIp = getClientIp(req);
  if (!isAuthorizedAdmin(req)) {
    return res.status(403).json({
      success: false,
      error: `ACCESS DENIED: Your IP (${clientIp}) is view-only. Only Admin (${ADMIN_IP}) has permission to edit.`
    });
  }

  const { sheet_ids = [], saw, team } = req.body;
  for (const sid of sheet_ids) {
    const s = sheets_registry.find((x) => x.sheet_id === sid);
    if (s) {
      s.assigned_saw = saw;
      s.assigned_team = team;
      s.status = 'ASSIGNED';
    }
  }
  saveState();
  res.json({ success: true, count: sheet_ids.length });
});

// 7. Toggle Saw 2 Repair Mode
app.post('/api/saw/toggle-repair', (req, res) => {
  const clientIp = getClientIp(req);
  if (!isAuthorizedAdmin(req)) {
    return res.status(403).json({
      success: false,
      error: `ACCESS DENIED: Your IP (${clientIp}) is view-only. Only Admin (${ADMIN_IP}) has permission to edit.`
    });
  }

  const cur = saws_status['Saw 2'].status;
  saws_status['Saw 2'].status = cur === 'IN_REPAIR' ? 'ONLINE' : 'IN_REPAIR';
  saveState();
  res.json({ success: true, status: saws_status['Saw 2'].status });
});

// 8. Reroute all Saw 2 pending sheets to Saw 1
app.post('/api/saw/reroute-to-saw1', (req, res) => {
  const clientIp = getClientIp(req);
  if (!isAuthorizedAdmin(req)) {
    return res.status(403).json({
      success: false,
      error: `ACCESS DENIED: Your IP (${clientIp}) is view-only. Only Admin (${ADMIN_IP}) has permission to edit.`
    });
  }

  let count = 0;
  for (const s of sheets_registry) {
    if (s.assigned_saw === 'Saw 2' && s.status !== 'CUT') {
      s.assigned_saw = 'Saw 1';
      s.assigned_team = 'Team A (Ahmed)';
      count++;
    }
  }
  saveState();
  res.json({ success: true, rerouted: count });
});

// 9. Log Previous Day Overtime
app.post('/api/overtime/log', (req, res) => {
  const clientIp = getClientIp(req);
  if (!isAuthorizedAdmin(req)) {
    return res.status(403).json({
      success: false,
      error: `ACCESS DENIED: Your IP (${clientIp}) is view-only. Only Admin (${ADMIN_IP}) has permission to edit.`
    });
  }

  const { date, team, hours } = req.body;
  const logEntry: OvertimeEntry = {
    date,
    team,
    hours: parseFloat(hours) || 0,
    logged_at: new Date().toISOString()
  };
  overtime_history.push(logEntry);
  saveState();
  res.json({ success: true });
});

// 10. Re-parse cutlist on demand
app.post('/api/reparse-cutlist', (req, res) => {
  const clientIp = getClientIp(req);
  if (!isAuthorizedAdmin(req)) {
    return res.status(403).json({
      success: false,
      error: `ACCESS DENIED: Only Admin (${ADMIN_IP}) can trigger cutlist re-parsing.`
    });
  }

  try {
    ensureCutlistParsed(true);
    loadState();
    res.json({ success: true, count: sheets_registry.length });
  } catch (err: any) {
    res.status(500).json({ success: false, error: err.message });
  }
});

// 11. AI Floor Assistant (Gemma / Gemini)
app.post('/api/gemma/ask', async (req, res) => {
  const { prompt } = req.body;
  if (!prompt) {
    return res.status(400).json({ error: 'Prompt is required' });
  }

  try {
    const totalSheets = sheets_registry.reduce((sum, s) => sum + s.repeat_count, 0);
    const cutSheets = sheets_registry.filter((s) => s.status === 'CUT').reduce((sum, s) => sum + s.repeat_count, 0);
    const s2Status = saws_status['Saw 2'].status;
    const sheets18 = sheets_registry.filter((s) => s.thickness === 18).reduce((sum, s) => sum + s.repeat_count, 0);
    const sheets8 = sheets_registry.filter((s) => s.thickness === 8).reduce((sum, s) => sum + s.repeat_count, 0);

    const response = await ai.models.generateContent({
      model: 'gemini-3.8-flash',
      contents: `You are Gemma Table Saw Production Director for BAZAR V3.
Context:
- Cutlist Total Volume: ${totalSheets} boards total (${sheets_registry.length} pattern groups across materials).
  * 18mm MDF Core Cutlist: ${sheets18} boards (from cutlist.html)
  * 8mm Back Panel Cutlist: ${sheets8} boards (from cutlist2.html)
- Current progress: ${cutSheets} sheets CUT out of ${totalSheets}.
- Saw 1: ONLINE (Team A - Ahmed).
- Saw 2: ${s2Status} (${s2Status === 'IN_REPAIR' ? 'Capacity halved to 1 saw' : 'Team B active'}).
- Shift hours: 8:00 AM - 12:00 PM (4h) and 1:00 PM - 4:00 PM (3h) = 7.0 h regular shift (420 min/saw).
- 5 PDF Drawing Sets: PDF 1 (SCMB), PDF 2 (CB-03-SF), PDF 3 (CB-02-SF), PDF 4 (SVMB Master), PDF 5 (CB-01 & Drawers).
- Admin IP: 192.168.1.173.

Question from Shop Floor Operator or Supervisor:
"${prompt}"

Answer concisely, authoritatively, and practically with clear recommendations for the table saw rip cuts and sheet sequencing.`,
      config: {
        systemInstruction:
          'You are Gemma, an industrial table saw production planner and MES director. Provide direct, concise shop-floor guidance.'
      }
    });

    const replyText = response.text || 'Unable to generate response from Gemma.';
    res.json({ reply: replyText });
  } catch (err: any) {
    console.error('[GEMINI API ERROR]', err.message);
    res.status(500).json({
      error: err.message,
      reply: `Notice: ${err.message}. Please consult the Cutlist table directly.`
    });
  }
});

// Setup Vite middleware in dev or serve static dist in prod
async function startServer() {
  const isProd = process.env.NODE_ENV === 'production' && fs.existsSync(path.join(__dirname, 'dist'));

  if (!isProd) {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: 'spa'
    });
    app.use(vite.middlewares);
  } else {
    app.use(express.static(path.join(__dirname, 'dist')));
    app.get('*', (req, res) => {
      res.sendFile(path.join(__dirname, 'dist', 'index.html'));
    });
  }

  app.listen(PORT, '0.0.0.0', () => {
    const ifaces = os.networkInterfaces();
    console.log(`================================================================`);
    console.log(`🏭 BAZAR V3 — TABLE SAW DISPATCH & PDF MATRIX MES ACTIVE`);
    console.log(`🔒 ADMIN RESTRICTED: Mutations locked to ${ADMIN_IP}`);
    console.log(`⚡ Saw 1: ONLINE (Team A - Ahmed)`);
    console.log(`🛠️ Saw 2: ${saws_status['Saw 2'].status} (Capacity auto-adjusted)`);
    console.log(`📋 Total Sheets in Cutlist: ${sheets_registry.length}`);
    console.log(`👉 Local Admin:   http://127.0.0.1:${PORT}`);
    for (const name of Object.keys(ifaces)) {
      for (const iface of ifaces[name] || []) {
        if (iface.family === 'IPv4' && !iface.internal) {
          console.log(`👉 Floor LAN:     http://${iface.address}:${PORT}`);
        }
      }
    }
    console.log(`================================================================`);
  });
}

startServer();
