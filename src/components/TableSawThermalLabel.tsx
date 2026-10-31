import React from 'react';
import { PartItem } from '../types/factory';

interface TableSawLabelProps {
  part: PartItem;
  sheetStart: number;
  pieceIndex: number;
  totalPieces: number;
  printMode?: boolean;
}

export function generateBarcodeSvg(text: string): React.ReactElement {
  let hash = 0;
  for (let i = 0; i < text.length; i++) {
    hash = (hash << 5) - hash + text.charCodeAt(i);
    hash |= 0;
  }

  const bars: { x: number; width: number }[] = [];
  let currentX = 6;

  // Start pattern
  bars.push({ x: currentX, width: 2 });
  currentX += 4;
  bars.push({ x: currentX, width: 1.5 });
  currentX += 3;

  for (let i = 0; i < 45; i++) {
    const bit = Math.abs((hash * (i + 1) * 9301 + 49297) % 100);
    const w = (bit % 3) + 1;
    if (i % 2 === 0) {
      bars.push({ x: currentX, width: w });
    }
    currentX += w + 1;
  }

  // End pattern
  bars.push({ x: currentX, width: 2 });
  currentX += 10;

  return (
    <svg
      viewBox={`0 0 ${currentX} 34`}
      className="w-full h-8 max-h-8"
      preserveAspectRatio="none"
      xmlns="http://www.w3.org/2000/svg"
    >
      {bars.map((bar, idx) => (
        <rect
          key={idx}
          x={bar.x}
          y="0"
          width={bar.width}
          height="34"
          fill="#000000"
        />
      ))}
    </svg>
  );
}

export const TableSawThermalLabel: React.FC<TableSawLabelProps> = ({
  part,
  sheetStart,
  pieceIndex,
  totalPieces,
  printMode = false
}) => {
  const serial = `${part.t}MM-S${sheetStart}-${part.code || 'P'}-${pieceIndex.toString().padStart(2, '0')}`;
  const coreLengthCm = (part.l / 10).toFixed(1);
  const coreWidthCm = (part.w / 10).toFixed(1);

  return (
    <div
      className={`thermal-label ${
        printMode ? 'thermal-print-page' : 'border border-slate-700 bg-white text-black shadow-lg rounded-md'
      }`}
      style={{
        width: '100mm',
        height: '50mm',
        padding: '3mm 4mm',
        boxSizing: 'border-box',
        border: '1.5px solid #000',
        display: 'flex',
        flexDirection: 'column',
        justifyContent: 'space-between',
        fontFamily: 'sans-serif',
        pageBreakAfter: 'always',
        backgroundColor: '#ffffff',
        color: '#000000',
        margin: printMode ? 0 : '4px'
      }}
    >
      {/* Top Header */}
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', borderBottom: '1.5px solid #000', paddingBottom: '1mm' }}>
        <div>
          <span style={{ background: '#000', color: '#fff', fontWeight: 900, fontSize: '9pt', padding: '1px 5px', borderRadius: '3px' }}>
            {part.pdf_set}
          </span>
          <span style={{ fontSize: '10pt', fontWeight: 900, marginLeft: '2mm' }}>
            {part.villa !== 'STANDARD 40' ? part.villa.toUpperCase() : part.cab.toUpperCase()}
          </span>
          <div style={{ fontSize: '7.5pt', fontWeight: 800, color: '#222', marginTop: '1mm', maxWidth: '65mm', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
            {part.name}
          </div>
        </div>
        <div style={{ textAlign: 'right' }}>
          <div style={{ fontSize: '9.5pt', fontWeight: 900, background: '#000', color: '#fff', padding: '1px 6px', borderRadius: '3px' }}>
            SHEET {sheetStart}
          </div>
          <div style={{ fontSize: '6.5pt', fontWeight: 800, marginTop: '1mm', color: '#444' }}>
            {part.t === 8 ? '8MM BACK PANEL' : `${part.t}MM MDF CORE`}
          </div>
        </div>
      </div>

      {/* Barcode & Serial (Pure Clean Industrial Barcode) */}
      <div style={{ margin: '1mm 0', textAlign: 'center', display: 'flex', flexDirection: 'column', alignItems: 'center' }}>
        <div style={{ width: '85%' }}>{generateBarcodeSvg(serial)}</div>
        <div style={{ fontFamily: 'monospace', fontSize: '9.5pt', fontWeight: 900, letterSpacing: '1px', marginTop: '1px' }}>
          {serial} <span style={{ fontSize: '7.5pt', fontWeight: 700 }}>[Piece {pieceIndex} of {totalPieces}]</span>
        </div>
      </div>

      {/* Footer Specs */}
      <div style={{ borderTop: '1.5px solid #000', paddingTop: '1mm', display: 'flex', justifyContent: 'space-between', fontSize: '7.5pt', fontWeight: 800 }}>
        <span>DIM: {part.l} × {part.w} × {part.t}mm</span>
        <span>CORE: {coreLengthCm} × {coreWidthCm} cm</span>
        <span>LIP: {part.lip}</span>
      </div>
    </div>
  );
};
