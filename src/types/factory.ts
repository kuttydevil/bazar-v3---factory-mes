// ==========================================
// TABLE SAW DISPATCH & OPEN CUTLIST TYPES
// ==========================================
export interface PartItem {
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

export interface SheetRecord {
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

export interface MaterialSummaryItem {
  material: string;
  parts_qty: number;
  rough_length: string;
  rough_area: string;
  rough_volume: string;
}

export interface SawMachine {
  status: 'ONLINE' | 'IN_REPAIR';
  team: string;
  assigned_sheets: number;
  cut_sheets: number;
}

export interface PdfMatrixItem {
  name: string;
  description: string;
  total_sheets: number;
  cut_sheets: number;
  pending_sheets: number;
  blocking_sheet_ids: string[];
  total_parts: number;
  cut_parts: number;
  percent_complete: number;
}

export interface Analytics {
  avg_cycle_min: number;
  total_pending_sheets: number;
  total_minutes_needed: number;
  total_days_needed: number;
  active_saws: number;
  shift_hours: string;
}

export interface OvertimeLog {
  date: string;
  team: string;
  hours: number;
  logged_at: string;
}

export interface TableSawFactoryState {
  saws: Record<string, SawMachine>;
  sheets: SheetRecord[];
  material_summary?: MaterialSummaryItem[];
  pdf_matrix: Record<string, PdfMatrixItem>;
  analytics: Analytics;
  overtime_history: OvertimeLog[];
  audit: Array<{
    ts: number;
    action: string;
    sheet_id?: string;
    saw?: string;
    duration?: number;
    op: string;
    note?: string;
  }>;
  clientIp: string;
  adminIp: string;
  isAdmin: boolean;
}

