import React, { useEffect, useRef, useState, useCallback } from 'react';
import jsQR from 'jsqr';
import {
  Camera,
  X,
  RefreshCw,
  Flashlight,
  CheckCircle2,
  AlertCircle,
  ExternalLink,
  Layers,
  Scissors,
  Check,
  Play,
  Search,
  VideoOff
} from 'lucide-react';
import { SheetRecord, PartItem } from '../types/factory';

export interface ScannedQrData {
  serial?: string;
  code?: string;
  name?: string;
  pdf_set?: string;
  villa?: string;
  cab?: string;
  sheetStart?: number;
  l?: number;
  w?: number;
  t?: number;
  lip?: string;
  pieceIndex?: number;
  totalPieces?: number;
  rawText: string;
}

interface QrScannerModalProps {
  isOpen: boolean;
  onClose: () => void;
  sheets: SheetRecord[];
  onSelectSheet?: (sheet: SheetRecord) => void;
  onToggleCut?: (sheetId: string, currentStatus: string) => void;
  onOpenEditSheet?: (sheet: SheetRecord) => void;
}

export const QrScannerModal: React.FC<QrScannerModalProps> = ({
  isOpen,
  onClose,
  sheets,
  onSelectSheet,
  onToggleCut,
  onOpenEditSheet
}) => {
  const videoRef = useRef<HTMLVideoElement | null>(null);
  const canvasRef = useRef<HTMLCanvasElement | null>(null);
  const animFrameIdRef = useRef<number | null>(null);
  const streamRef = useRef<MediaStream | null>(null);

  const [cameraActive, setCameraActive] = useState<boolean>(false);
  const [cameraError, setCameraError] = useState<string | null>(null);
  const [torchOn, setTorchOn] = useState<boolean>(false);
  const [scannedResult, setScannedResult] = useState<ScannedQrData | null>(null);
  const [matchedSheet, setMatchedSheet] = useState<SheetRecord | null>(null);
  const [matchedPart, setMatchedPart] = useState<PartItem | null>(null);
  const [scanHistory, setScanHistory] = useState<Array<{ text: string; time: string }>>([]);
  const [audioBeepEnabled, setAudioBeepEnabled] = useState<boolean>(true);
  const [manualCodeInput, setManualCodeInput] = useState<string>('');

  // Play factory feedback beep on barcode detection
  const playBeep = useCallback(() => {
    if (!audioBeepEnabled) return;
    try {
      const audioCtx = new (window.AudioContext || (window as any).webkitAudioContext)();
      const osc = audioCtx.createOscillator();
      const gain = audioCtx.createGain();
      osc.type = 'sine';
      osc.frequency.setValueAtTime(880, audioCtx.currentTime);
      gain.gain.setValueAtTime(0.2, audioCtx.currentTime);
      gain.gain.exponentialRampToValueAtTime(0.01, audioCtx.currentTime + 0.15);
      osc.connect(gain);
      gain.connect(audioCtx.destination);
      osc.start();
      osc.stop(audioCtx.currentTime + 0.15);
    } catch {
      // Audio context might be restricted before interaction
    }
  }, [audioBeepEnabled]);

  // Match scanned text or JSON payload against factory database
  const processQrText = useCallback(
    (text: string) => {
      const trimmed = text.trim();
      if (!trimmed) return;

      let parsed: ScannedQrData = { rawText: trimmed };

      // Attempt parsing JSON payload if available
      try {
        const obj = JSON.parse(trimmed);
        if (typeof obj === 'object' && obj !== null) {
          parsed = {
            serial: obj.s,
            code: obj.c,
            name: obj.n,
            pdf_set: obj.p,
            villa: obj.v,
            cab: obj.cab,
            sheetStart: obj.sh,
            l: obj.l,
            w: obj.w,
            t: obj.t,
            lip: obj.lip,
            pieceIndex: obj.pc,
            totalPieces: obj.tot,
            rawText: trimmed
          };
        }
      } catch {
        // Standard Factory Barcode format: "18MM-S1-AV-01" or "8MM-S40-D-01"
        const m = trimmed.match(/(\d+)MM-S(\d+)-([A-Z0-9]+)-(\d+)/i);
        if (m) {
          parsed.t = parseInt(m[1], 10);
          parsed.sheetStart = parseInt(m[2], 10);
          parsed.code = m[3].toUpperCase();
          parsed.pieceIndex = parseInt(m[4], 10);
          parsed.serial = trimmed;
        } else {
          // Check for single Part Code e.g. "AV", "FK", "AT" or Sheet ID "S1_1"
          parsed.code = trimmed.toUpperCase();
          parsed.serial = trimmed;
        }
      }

      setScannedResult(parsed);
      playBeep();

      // Add to session scan history
      setScanHistory((prev) => [
        { text: parsed.serial || parsed.name || trimmed.slice(0, 30), time: new Date().toLocaleTimeString() },
        ...prev.slice(0, 8)
      ]);

      // Find matching sheet and part in factory registry
      let foundSheet: SheetRecord | null = null;
      let foundPart: PartItem | null = null;

      // 1. Direct match by sheetStart and thickness
      if (parsed.sheetStart !== undefined) {
        const t = parsed.t || 18;
        foundSheet =
          sheets.find(
            (s) =>
              s.thickness === t &&
              s.start_sheet <= parsed.sheetStart! &&
              s.end_sheet >= parsed.sheetStart!
          ) || null;
      }

      // 2. Fallback match by sheet ID or title in raw text
      if (!foundSheet) {
        foundSheet =
          sheets.find(
            (s) =>
              s.sheet_id.toLowerCase() === trimmed.toLowerCase() ||
              s.sheet_title.toLowerCase() === trimmed.toLowerCase() ||
              trimmed.toLowerCase().includes(s.sheet_id.toLowerCase())
          ) || null;
      }

      // 3. Fallback match by part code across all sheets
      if (parsed.code) {
        for (const s of sheets) {
          const p = s.parts.find((x) => x.code.toUpperCase() === parsed.code?.toUpperCase());
          if (p) {
            if (!foundSheet) foundSheet = s;
            foundPart = p;
            break;
          }
        }
      }

      if (foundSheet && !foundPart && parsed.code) {
        foundPart = foundSheet.parts.find((p) => p.code.toUpperCase() === parsed.code?.toUpperCase()) || null;
      }

      // If part found, populate additional details if missing
      if (foundPart) {
        parsed.name = parsed.name || foundPart.name;
        parsed.pdf_set = parsed.pdf_set || foundPart.pdf_set;
        parsed.villa = parsed.villa || foundPart.villa;
        parsed.cab = parsed.cab || foundPart.cab;
        parsed.l = parsed.l || foundPart.l;
        parsed.w = parsed.w || foundPart.w;
        parsed.t = parsed.t || foundPart.t;
        parsed.lip = parsed.lip || foundPart.lip;
        parsed.totalPieces = parsed.totalPieces || foundPart.qty_per_sheet;
      }

      setMatchedSheet(foundSheet);
      setMatchedPart(foundPart);
    },
    [sheets, playBeep]
  );

  // Frame processing loop for QR / 2D barcode detection via jsQR
  const tick = useCallback(() => {
    if (!videoRef.current || !canvasRef.current) return;

    const video = videoRef.current;
    const canvas = canvasRef.current;
    const ctx = canvas.getContext('2d', { willReadFrequently: true });

    if (video.readyState === video.HAVE_ENOUGH_DATA && ctx) {
      canvas.width = video.videoWidth;
      canvas.height = video.videoHeight;
      ctx.drawImage(video, 0, 0, canvas.width, canvas.height);

      const imageData = ctx.getImageData(0, 0, canvas.width, canvas.height);
      const code = jsQR(imageData.data, imageData.width, imageData.height, {
        inversionAttempts: 'dontInvert'
      });

      if (code && code.data) {
        if (!scannedResult || scannedResult.rawText !== code.data) {
          processQrText(code.data);
        }
      }
    }

    animFrameIdRef.current = requestAnimationFrame(tick);
  }, [scannedResult, processQrText]);

  // Start Camera Stream with comprehensive multi-tier fallback
  const startCamera = async () => {
    setCameraError(null);
    setScannedResult(null);
    setMatchedSheet(null);
    setMatchedPart(null);

    if (!navigator.mediaDevices || !navigator.mediaDevices.getUserMedia) {
      setCameraError('Camera access is not supported on this browser or requires a secure context (HTTPS / localhost). Use manual entry below.');
      setCameraActive(false);
      return;
    }

    // Tier 1: Try back/environment camera (phones/tablets)
    // Tier 2: Try default video stream (laptops/webcams)
    // Tier 3: Try minimal video stream
    const attempts: MediaStreamConstraints[] = [
      { video: { facingMode: { ideal: 'environment' }, width: { ideal: 1280 }, height: { ideal: 720 } }, audio: false },
      { video: { facingMode: 'user' }, audio: false },
      { video: true, audio: false }
    ];

    let stream: MediaStream | null = null;
    let lastErr: any = null;

    for (const constraints of attempts) {
      try {
        stream = await navigator.mediaDevices.getUserMedia(constraints);
        if (stream) break;
      } catch (err: any) {
        lastErr = err;
      }
    }

    if (stream && videoRef.current) {
      streamRef.current = stream;
      videoRef.current.srcObject = stream;
      videoRef.current.setAttribute('playsinline', 'true');
      try {
        await videoRef.current.play();
        setCameraActive(true);
        animFrameIdRef.current = requestAnimationFrame(tick);
      } catch (playErr: any) {
        console.warn('Video play error:', playErr);
        setCameraActive(true);
      }
    } else {
      const errorMsg = lastErr?.name === 'NotFoundError' || lastErr?.message?.includes('Requested device not found')
        ? 'No physical camera detected on this device. You can type or paste any part code (e.g. AV, FK, 18MM-S1-AV-01) below to inspect instantly.'
        : (lastErr?.message || 'Unable to access camera.');
      setCameraError(errorMsg);
      setCameraActive(false);
    }
  };

  // Stop Camera Stream
  const stopCamera = () => {
    if (animFrameIdRef.current) {
      cancelAnimationFrame(animFrameIdRef.current);
      animFrameIdRef.current = null;
    }
    if (streamRef.current) {
      streamRef.current.getTracks().forEach((track) => track.stop());
      streamRef.current = null;
    }
    setCameraActive(false);
    setTorchOn(false);
  };

  // Toggle Camera Flashlight/Torch if supported
  const toggleTorch = async () => {
    if (!streamRef.current) return;
    const track = streamRef.current.getVideoTracks()[0];
    if (track && (track.getCapabilities as any)?.torch) {
      try {
        const next = !torchOn;
        await track.applyConstraints({
          advanced: [{ torch: next } as any]
        });
        setTorchOn(next);
      } catch (e) {
        console.warn('Torch toggle failed', e);
      }
    }
  };

  // Lifecycle
  useEffect(() => {
    if (isOpen) {
      startCamera();
    } else {
      stopCamera();
    }
    return () => {
      stopCamera();
    };
  }, [isOpen]);

  if (!isOpen) return null;

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/90 backdrop-blur-md p-3 sm:p-4">
      <div className="bg-[#131d33] border border-[#223354] rounded-2xl shadow-2xl w-full max-w-4xl max-h-[95vh] flex flex-col overflow-hidden animate-in fade-in zoom-in-95 duration-200">
        {/* Top Header */}
        <div className="p-4 border-b border-[#223354] flex justify-between items-center bg-[#0f172a]">
          <div className="flex items-center gap-2.5">
            <div className="p-2 bg-sky-500/20 text-sky-400 rounded-xl border border-sky-500/30">
              <Camera className="w-5 h-5" />
            </div>
            <div>
              <h3 className="text-sm sm:text-base font-black text-white flex items-center gap-2">
                <span>Shop Floor Barcode &amp; Part Scanner</span>
                <span className="bg-emerald-500/20 text-emerald-400 border border-emerald-500/30 text-[10px] px-2 py-0.5 rounded-full font-bold uppercase">
                  {cameraActive ? 'Live Camera' : 'Ready'}
                </span>
              </h3>
              <p className="text-xs text-slate-400 mt-0.5">
                Scan or enter 100mm × 50mm thermal barcode stickers to inspect dimensions, core cm, and rip sequence
              </p>
            </div>
          </div>

          <div className="flex items-center gap-2">
            <button
              onClick={() => setAudioBeepEnabled(!audioBeepEnabled)}
              className={`p-2 rounded-lg text-xs font-bold transition border ${
                audioBeepEnabled
                  ? 'bg-sky-500/20 text-sky-300 border-sky-500/40'
                  : 'bg-[#0b1325] text-slate-400 border-[#223354]'
              }`}
              title="Toggle Audio Feedback"
            >
              🔔 Beep: {audioBeepEnabled ? 'ON' : 'OFF'}
            </button>

            <button
              onClick={onClose}
              className="p-2 text-slate-400 hover:text-white rounded-lg transition"
            >
              <X className="w-5 h-5" />
            </button>
          </div>
        </div>

        {/* Main Content Area */}
        <div className="p-4 overflow-y-auto flex-1 grid grid-cols-1 lg:grid-cols-12 gap-4 bg-[#070d19]">
          {/* Camera Viewport / Manual Search (Left/Top) */}
          <div className="lg:col-span-7 flex flex-col space-y-3">
            <div className="relative aspect-video w-full bg-black rounded-xl overflow-hidden border-2 border-sky-500/40 shadow-inner flex items-center justify-center">
              <video
                ref={videoRef}
                className="w-full h-full object-cover"
                autoPlay
                playsInline
                muted
              />
              <canvas ref={canvasRef} className="hidden" />

              {/* Scanning Target Crosshairs Overlay */}
              {cameraActive && (
                <div className="absolute inset-0 pointer-events-none flex items-center justify-center">
                  <div className="w-64 h-64 border-2 border-dashed border-sky-400/70 rounded-2xl relative flex items-center justify-center">
                    <div className="w-12 h-12 border-t-4 border-l-4 border-sky-400 absolute -top-1 -left-1 rounded-tl" />
                    <div className="w-12 h-12 border-t-4 border-r-4 border-sky-400 absolute -top-1 -right-1 rounded-tr" />
                    <div className="w-12 h-12 border-b-4 border-l-4 border-sky-400 absolute -bottom-1 -left-1 rounded-bl" />
                    <div className="w-12 h-12 border-b-4 border-r-4 border-sky-400 absolute -bottom-1 -right-1 rounded-br" />
                    <div className="w-full h-0.5 bg-gradient-to-r from-transparent via-rose-500 to-transparent shadow-[0_0_12px_#f43f5e] animate-bounce" />
                  </div>
                  <span className="absolute bottom-4 bg-black/75 px-3 py-1 rounded-full text-[11px] font-bold text-sky-300 border border-sky-500/40">
                    Align Barcode or Label inside target
                  </span>
                </div>
              )}

              {/* Camera Error / No Device Fallback Display */}
              {cameraError && (
                <div className="absolute inset-0 bg-[#0b1325]/95 p-6 flex flex-col items-center justify-center text-center space-y-3 z-20">
                  <div className="p-3 bg-amber-500/10 text-amber-400 rounded-full border border-amber-500/20">
                    <VideoOff className="w-8 h-8" />
                  </div>
                  <div className="text-xs font-bold text-slate-200 max-w-md">
                    {cameraError}
                  </div>
                  <button
                    onClick={startCamera}
                    className="px-3.5 py-1.5 bg-[#131d33] hover:bg-[#1c2a4a] text-sky-300 border border-sky-500/30 rounded-lg text-xs font-bold transition flex items-center gap-1.5 shadow"
                  >
                    <RefreshCw className="w-3.5 h-3.5" />
                    <span>Retry Camera</span>
                  </button>
                </div>
              )}
            </div>

            {/* Quick Barcode / Part Code Input Bar */}
            <div className="bg-[#0b1325] border border-[#223354] rounded-xl p-2.5 flex items-center gap-2">
              <Search className="w-4 h-4 text-sky-400 shrink-0" />
              <input
                type="text"
                value={manualCodeInput}
                onChange={(e) => setManualCodeInput(e.target.value)}
                onKeyDown={(e) => {
                  if (e.key === 'Enter') {
                    processQrText(manualCodeInput);
                  }
                }}
                placeholder="Type part code (e.g. AV, FK, D) or barcode serial..."
                className="bg-transparent text-white text-xs font-mono font-bold placeholder:text-slate-500 outline-none w-full"
              />
              <button
                onClick={() => processQrText(manualCodeInput)}
                className="px-3 py-1 bg-sky-600 hover:bg-sky-500 text-white rounded-lg text-xs font-bold transition shrink-0"
              >
                Inspect
              </button>
            </div>

            {/* Quick Test Samples */}
            <div className="flex items-center justify-between flex-wrap gap-2 text-xs">
              <div className="flex items-center gap-1">
                <span className="text-[11px] text-slate-400 font-bold">Quick Samples:</span>
                <button
                  onClick={() => processQrText('18MM-S1-AV-01')}
                  className="px-2 py-1 bg-[#131d33] hover:bg-[#1c2a4a] text-sky-300 border border-sky-500/30 rounded text-[10.5px] font-bold transition"
                >
                  [AV] 18mm Bed 01
                </button>
                <button
                  onClick={() => processQrText('18MM-S1-FK-01')}
                  className="px-2 py-1 bg-[#131d33] hover:bg-[#1c2a4a] text-amber-300 border border-amber-500/30 rounded text-[10.5px] font-bold transition"
                >
                  [FK] Villa 79
                </button>
                <button
                  onClick={() => processQrText('8MM-S1-D-01')}
                  className="px-2 py-1 bg-[#131d33] hover:bg-[#1c2a4a] text-emerald-300 border border-emerald-500/30 rounded text-[10.5px] font-bold transition"
                >
                  [D] 8mm Back Panel
                </button>
              </div>

              {cameraActive && (
                <button
                  onClick={toggleTorch}
                  className={`px-2.5 py-1 rounded-lg font-bold flex items-center gap-1 border transition text-[11px] ${
                    torchOn
                      ? 'bg-amber-500/20 text-amber-300 border-amber-500/50'
                      : 'bg-[#0b1325] text-slate-300 border-[#223354] hover:bg-[#1c2a4a]'
                  }`}
                >
                  <Flashlight className="w-3 h-3" />
                  <span>Torch</span>
                </button>
              )}
            </div>

            {/* Recent Scans Strip */}
            {scanHistory.length > 0 && (
              <div className="bg-[#0b1325] border border-[#223354] rounded-xl p-2.5">
                <div className="text-[10px] font-bold text-slate-400 uppercase tracking-wider mb-1.5">
                  Recent Scans This Session:
                </div>
                <div className="flex gap-2 overflow-x-auto scrollbar-none">
                  {scanHistory.map((h, idx) => (
                    <button
                      key={idx}
                      onClick={() => processQrText(h.text)}
                      className="px-2 py-0.5 bg-[#131d33] hover:bg-[#1c2a4a] border border-[#223354] text-[10px] font-mono text-slate-300 rounded whitespace-nowrap transition cursor-pointer"
                    >
                      {h.text} <span className="text-slate-500">({h.time})</span>
                    </button>
                  ))}
                </div>
              </div>
            )}
          </div>

          {/* Scanner Data Inspection Card (Right/Bottom) */}
          <div className="lg:col-span-5 flex flex-col space-y-3">
            {scannedResult ? (
              <div className="bg-[#0b1325] border-2 border-emerald-500/60 rounded-xl p-4 space-y-4 shadow-xl animate-in fade-in slide-in-from-right-4 duration-200">
                {/* Header status */}
                <div className="flex items-start justify-between border-b border-[#223354] pb-3">
                  <div>
                    <span className="inline-flex items-center gap-1.5 px-2.5 py-0.5 rounded-full text-[10px] font-black bg-emerald-500/20 text-emerald-300 border border-emerald-500/40">
                      <CheckCircle2 className="w-3 h-3 text-emerald-400" />
                      BARCODE IDENTIFIED
                    </span>
                    <h4 className="text-sm font-black text-white mt-1">
                      {scannedResult.name || scannedResult.serial}
                    </h4>
                  </div>
                  <div className="text-right">
                    <span className="font-mono text-xs font-bold text-sky-400">
                      [{scannedResult.code || 'CODE'}]
                    </span>
                    <div className="text-[10.5px] text-slate-400 font-mono">
                      Piece {scannedResult.pieceIndex || 1} of {scannedResult.totalPieces || 1}
                    </div>
                  </div>
                </div>

                {/* Component Specs Grid */}
                <div className="grid grid-cols-2 gap-2 text-xs">
                  <div className="bg-[#131d33] border border-[#223354] p-2.5 rounded-lg">
                    <div className="text-[10px] font-bold text-slate-400 uppercase">Cabinet / Villa</div>
                    <div className="font-bold text-white mt-0.5 truncate">{scannedResult.cab || 'Standard'}</div>
                    <div className="text-amber-400 font-extrabold text-[10.5px]">{scannedResult.villa || 'STANDARD 40'}</div>
                  </div>

                  <div className="bg-[#131d33] border border-[#223354] p-2.5 rounded-lg">
                    <div className="text-[10px] font-bold text-slate-400 uppercase">Drawing PDF Set</div>
                    <div className="font-bold text-sky-400 mt-0.5">{scannedResult.pdf_set || 'PDF'}</div>
                    <div className="text-slate-400 text-[10px]">Sheet #{scannedResult.sheetStart || (matchedSheet ? matchedSheet.start_sheet : '-')}</div>
                  </div>

                  <div className="bg-[#131d33] border border-[#223354] p-2.5 rounded-lg">
                    <div className="text-[10px] font-bold text-slate-400 uppercase">Finished Dimension</div>
                    <div className="font-mono font-black text-emerald-400 mt-0.5">
                      {scannedResult.l || (matchedPart ? matchedPart.l : 0)} × {scannedResult.w || (matchedPart ? matchedPart.w : 0)} × {scannedResult.t || (matchedPart ? matchedPart.t : 18)} mm
                    </div>
                    <div className="text-[10px] text-slate-400 font-mono">
                      Core: {(((scannedResult.l || (matchedPart ? matchedPart.l : 0))) / 10).toFixed(1)} × {(((scannedResult.w || (matchedPart ? matchedPart.w : 0))) / 10).toFixed(1)} cm
                    </div>
                  </div>

                  <div className="bg-[#131d33] border border-[#223354] p-2.5 rounded-lg">
                    <div className="text-[10px] font-bold text-slate-400 uppercase">Edge Banding (Lip)</div>
                    <div className="font-bold text-white mt-0.5">{scannedResult.lip || (matchedPart ? matchedPart.lip : '2mm Front')}</div>
                    <div className="text-[10px] text-slate-400">Standard Wood Grain</div>
                  </div>
                </div>

                {/* Linked Sheet Status & Saw Floor Actions */}
                {matchedSheet ? (
                  <div className="bg-[#131d33] border border-sky-500/40 rounded-xl p-3.5 space-y-2.5">
                    <div className="flex items-center justify-between">
                      <div className="font-extrabold text-xs text-slate-200 flex items-center gap-1.5">
                        <Layers className="w-3.5 h-3.5 text-sky-400" />
                        <span>Found in {matchedSheet.sheet_title}:</span>
                      </div>
                      <span className={`st st-${matchedSheet.status}`}>{matchedSheet.status}</span>
                    </div>

                    <div className="text-xs text-slate-300 space-y-1 font-medium">
                      <div>• Assigned: <strong>{matchedSheet.assigned_saw} ({matchedSheet.assigned_team})</strong></div>
                      <div>• Material: <span className="text-slate-400">{matchedSheet.material}</span></div>
                      <div>• Pattern Efficiency: <strong className="text-emerald-400">{matchedSheet.efficiency}</strong></div>
                    </div>

                    {/* Rip Cuts Sequence */}
                    {matchedSheet.rip_cuts && matchedSheet.rip_cuts.length > 0 && (
                      <div className="pt-2 border-t border-[#223354]">
                        <div className="text-[10.5px] font-bold text-slate-400 mb-1 flex items-center gap-1">
                          <Scissors className="w-3 h-3 text-emerald-400" />
                          <span>Primary Table Saw Rip Cuts:</span>
                        </div>
                        <div className="flex flex-wrap gap-1">
                          {matchedSheet.rip_cuts.map((rc, idx) => (
                            <span key={idx} className="rip-cut-tag text-[10px] py-0.5 px-2 font-bold">
                              Cut {idx + 1}: {rc}
                            </span>
                          ))}
                        </div>
                      </div>
                    )}

                    {/* Floor Action Buttons */}
                    <div className="pt-2 border-t border-[#223354] flex gap-2 flex-wrap">
                      {onToggleCut && (
                        <button
                          onClick={() => {
                            onToggleCut(matchedSheet.sheet_id, matchedSheet.status);
                            setMatchedSheet({
                              ...matchedSheet,
                              status: matchedSheet.status === 'CUTTING' ? 'CUT' : 'CUTTING'
                            });
                          }}
                          className={`px-3 py-1.5 rounded-lg text-xs font-bold text-white transition flex items-center gap-1 shadow ${
                            matchedSheet.status === 'CUTTING'
                              ? 'bg-emerald-600 hover:bg-emerald-500'
                              : matchedSheet.status === 'CUT'
                              ? 'bg-slate-700 hover:bg-slate-600'
                              : 'bg-emerald-700 hover:bg-emerald-600'
                          }`}
                        >
                          {matchedSheet.status === 'CUTTING' ? (
                            <>
                              <Check className="w-3.5 h-3.5" />
                              <span>✓ Mark Sheet CUT</span>
                            </>
                          ) : (
                            <>
                              <Play className="w-3.5 h-3.5" />
                              <span>▶ Start Cutting</span>
                            </>
                          )}
                        </button>
                      )}

                      {onOpenEditSheet && (
                        <button
                          onClick={() => {
                            onOpenEditSheet(matchedSheet);
                            onClose();
                          }}
                          className="px-3 py-1.5 bg-[#0b1325] hover:bg-[#1c2a4a] text-amber-400 border border-amber-500/30 rounded-lg text-xs font-bold transition shadow"
                        >
                          Edit Record
                        </button>
                      )}

                      {onSelectSheet && (
                        <button
                          onClick={() => {
                            onSelectSheet(matchedSheet);
                            onClose();
                          }}
                          className="px-3 py-1.5 bg-sky-600 hover:bg-sky-500 text-white rounded-lg text-xs font-bold transition shadow flex items-center gap-1"
                        >
                          <ExternalLink className="w-3.5 h-3.5" />
                          <span>View SVG Diagram</span>
                        </button>
                      )}
                    </div>
                  </div>
                ) : (
                  <div className="p-3 bg-[#131d33] border border-[#223354] rounded-lg text-xs text-slate-400">
                    Component data parsed, but matching sheet was not found in active batch.
                  </div>
                )}
              </div>
            ) : (
              <div className="bg-[#0b1325] border border-[#223354] rounded-xl p-8 flex flex-col items-center justify-center text-center space-y-3 h-full">
                <div className="p-3 bg-sky-500/10 text-sky-400 rounded-2xl border border-sky-500/20">
                  <Camera className="w-8 h-8 opacity-70 animate-pulse" />
                </div>
                <div className="text-xs font-bold text-slate-200">
                  Ready to Scan or Search
                </div>
                <p className="text-[11px] text-slate-400 max-w-xs">
                  Hold a printed 100mm × 50mm thermal part sticker in front of your camera, or type the part code/serial directly in the search bar above.
                </p>
              </div>
            )}
          </div>
        </div>
      </div>
    </div>
  );
};
