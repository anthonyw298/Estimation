'use client';

import { useMemo } from 'react';
import type { DoorConfig } from '@/types';
import { calculateDloWidth, calculateDloHeight, DLO_EDGE_DEDUCTION, DLO_INTERIOR_DEDUCTION, DLO_SILL_DEDUCTION } from '@/lib/formulas';

interface BayDiagramProps {
  baysWide: number;
  baysTall: number;
  openingWidth: number;
  openingHeight: number;
  customBayWidths?: number[];
  customBayHeights?: number[];
  doors?: DoorConfig[];
}

export default function BayDiagram({
  baysWide,
  baysTall,
  openingWidth,
  openingHeight,
  customBayWidths,
  customBayHeights,
  doors = [],
}: BayDiagramProps) {
  const diagram = useMemo(() => {
    if (baysWide <= 0 || baysTall <= 0 || openingWidth <= 0 || openingHeight <= 0) {
      return null;
    }

    // Calculate bay dimensions
    const bayWidths: number[] =
      customBayWidths && customBayWidths.length === baysWide
        ? customBayWidths
        : Array(baysWide).fill(openingWidth / baysWide);

    const bayHeights: number[] =
      customBayHeights && customBayHeights.length === baysTall
        ? customBayHeights
        : Array(baysTall).fill(openingHeight / baysTall);

    // SVG layout constants
    const maxWidth = 500;
    const padding = { top: 58, right: 100, bottom: 60, left: 30 };
    const availableDrawWidth = maxWidth - padding.left - padding.right;
    const scale = availableDrawWidth / openingWidth;
    const drawHeight = openingHeight * scale;
    const svgWidth = maxWidth;
    const svgHeight = drawHeight + padding.top + padding.bottom;

    // Drawing origin
    const ox = padding.left;
    const oy = padding.top;
    const rectW = openingWidth * scale;
    const rectH = drawHeight;

    // Build vertical grid lines (bay widths)
    const verticalLines: { x: number; label: string }[] = [];
    let accX = 0;
    for (let i = 0; i < baysWide; i++) {
      const w = bayWidths[i];
      verticalLines.push({
        x: accX + w / 2,
        label: w.toFixed(1) + '"',
      });
      accX += w;
    }

    // Build horizontal grid lines (bay heights)
    const horizontalLines: { y: number; label: string }[] = [];
    let accY = 0;
    for (let i = 0; i < baysTall; i++) {
      const h = bayHeights[i];
      horizontalLines.push({
        y: accY + h / 2,
        label: h.toFixed(1) + '"',
      });
      accY += h;
    }

    // Accumulate positions for grid lines
    const vLinePositions: number[] = [];
    let vAcc = 0;
    for (let i = 0; i < baysWide - 1; i++) {
      vAcc += bayWidths[i];
      vLinePositions.push(vAcc);
    }

    const hLinePositions: number[] = [];
    let hAcc = 0;
    for (let i = 0; i < baysTall - 1; i++) {
      hAcc += bayHeights[i];
      hLinePositions.push(hAcc);
    }

    // Parse doors for drawing
    const doorRects: Array<{
      x: number; y: number; width: number; height: number;
      label: string;
      bayLabel?: string;
    }> = [];

    for (const door of doors) {
      const sizeParts = door.size.match(/(\d+)'\s*X\s*(\d+)'/i);
      if (!sizeParts) continue;
      const doorWidthInches = parseInt(sizeParts[1]) * 12;
      const doorHeightInches = parseInt(sizeParts[2]) * 12;

      // For each copy of this door
      for (let c = 0; c < door.count; c++) {
        let xInches: number;

        if (door.bayIndex != null && door.bayIndex < bayWidths.length) {
          // Position door(s) centered within the assigned bay
          const bayLeft = bayWidths.slice(0, door.bayIndex).reduce((a, b) => a + b, 0);
          const bayW = bayWidths[door.bayIndex];
          const totalDoorsW = door.count * doorWidthInches + (door.count - 1) * 2;
          const startX = bayLeft + bayW / 2 - totalDoorsW / 2;
          xInches = startX + c * (doorWidthInches + 2);
        } else if (door.x_positions && door.x_positions[c] != null) {
          xInches = door.x_positions[c];
        } else if (door.x_in != null) {
          xInches = door.x_in + c * (doorWidthInches + 2);
        } else {
          // Default: place from left edge, side by side with 2" gap
          xInches = c * (doorWidthInches + 2);
        }

        doorRects.push({
          x: xInches * scale,
          y: (openingHeight - doorHeightInches) * scale, // bottom-aligned
          width: doorWidthInches * scale,
          height: doorHeightInches * scale,
          label: `${sizeParts[1]}' x ${sizeParts[2]}'`,
          bayLabel: door.bayIndex != null ? `Bay ${door.bayIndex + 1}` : undefined,
        });
      }
    }

    return {
      svgWidth,
      svgHeight,
      ox,
      oy,
      rectW,
      rectH,
      scale,
      bayWidths,
      bayHeights,
      verticalLines,
      horizontalLines,
      vLinePositions,
      hLinePositions,
      doorRects,
    };
  }, [baysWide, baysTall, openingWidth, openingHeight, customBayWidths, customBayHeights, doors]);

  if (!diagram) {
    return (
      <div className="bg-[#111118] border border-[#1e1e2a] rounded-2xl p-8 text-center">
        <p className="text-sm text-[#ffffff]">
          Enter valid opening dimensions and bay counts to see the diagram.
        </p>
      </div>
    );
  }

  const {
    svgWidth,
    svgHeight,
    ox,
    oy,
    rectW,
    rectH,
    scale,
    bayWidths,
    bayHeights,
    verticalLines,
    horizontalLines,
    vLinePositions,
    hLinePositions,
    doorRects,
  } = diagram;

  return (
    <div className="bg-[#111118] border border-[#1e1e2a] rounded-2xl p-5">
      <svg
        viewBox={`0 0 ${svgWidth} ${svgHeight}`}
        className="w-full"
        style={{ maxWidth: svgWidth }}
      >
        {/* Background fill for the opening */}
        <rect
          x={ox}
          y={oy}
          width={rectW}
          height={rectH}
          fill="#0c0c12"
          stroke="#2a2a3a"
          strokeWidth={1.5}
        />

        {/* Outer frame — jambs (left/right 2"), head (top 2"), sill (bottom 2-7/16") */}
        {(() => {
          const frameW = 2 * scale; // 2" for jambs and head
          const frameFill = '#2a2a3c';
          const frameStroke = '#40405a';
          const labelFill = '#8888aa';
          return (
            <>
              {/* Left jamb — 2" */}
              <rect x={ox} y={oy} width={frameW} height={rectH}
                fill={frameFill} stroke={frameStroke} strokeWidth={0.5} />
              <text x={ox - 4} y={oy + rectH / 2} textAnchor="end" dominantBaseline="central"
                fill={labelFill} fontSize={7} fontFamily="monospace"
                transform={`rotate(-90, ${ox - 4}, ${oy + rectH / 2})`}>
                2&quot;
              </text>

              {/* Right jamb — 2" */}
              <rect x={ox + rectW - frameW} y={oy} width={frameW} height={rectH}
                fill={frameFill} stroke={frameStroke} strokeWidth={0.5} />

              {/* Head (top) — 2" */}
              <rect x={ox} y={oy} width={rectW} height={frameW}
                fill={frameFill} stroke={frameStroke} strokeWidth={0.5} />

              {/* Sill (bottom) — 2-7/16" (2" frame + 7/16" sill deduction) */}
              {(() => {
                const sillTotal = (2 + DLO_SILL_DEDUCTION) * scale;
                return (
                  <>
                    <rect x={ox} y={oy + rectH - sillTotal} width={rectW} height={sillTotal}
                      fill={frameFill} stroke={frameStroke} strokeWidth={0.5} />
                    <text x={ox + rectW / 2} y={oy + rectH + 8} textAnchor="middle"
                      fill={labelFill} fontSize={7} fontFamily="monospace">
                      2-7/16&quot;
                    </text>
                  </>
                );
              })()}
            </>
          );
        })()}

        {/* Vertical mullions (between bays) — 2″ wide */}
        {vLinePositions.map((pos, i) => {
          const mullionW = DLO_INTERIOR_DEDUCTION * scale;
          const cx = ox + pos * scale;
          return (
            <g key={`v-${i}`}>
              <rect
                x={cx - mullionW / 2}
                y={oy}
                width={mullionW}
                height={rectH}
                fill="#3a3a4c"
                stroke="#50506a"
                strokeWidth={0.5}
              />
              {/* 2″ label at top, between adjacent DLO labels */}
              <text
                x={cx}
                y={oy - 18}
                textAnchor="middle"
                fill="#8888aa"
                fontSize={8}
                fontFamily="monospace"
              >
                2&quot;
              </text>
            </g>
          );
        })}

        {/* Horizontal mullions (between bays) — 2″ tall */}
        {hLinePositions.map((pos, i) => {
          const mullionH = DLO_INTERIOR_DEDUCTION * scale;
          const cy = oy + pos * scale;
          return (
            <g key={`h-${i}`}>
              <rect
                x={ox}
                y={cy - mullionH / 2}
                width={rectW}
                height={mullionH}
                fill="#3a3a4c"
                stroke="#50506a"
                strokeWidth={0.5}
              />
              {/* 2″ label on right, between adjacent DLO labels */}
              <text
                x={ox + rectW + 10}
                y={cy}
                textAnchor="start"
                dominantBaseline="central"
                fill="#8888aa"
                fontSize={8}
                fontFamily="monospace"
              >
                2&quot;
              </text>
            </g>
          );
        })}

        {/* (7/16″ sill deduction is now integrated into the combined 2-7/16" sill frame above) */}

        {/* Bay width labels (at top) — C/L and D.L.O. */}
        {(() => {
          let xAcc = 0;
          return verticalLines.map((v, i) => {
            const midX = ox + (xAcc + bayWidths[i] / 2) * scale;
            const dlo = calculateDloWidth(bayWidths[i], i, baysWide);
            xAcc += bayWidths[i];
            return (
              <g key={`wl-${i}`}>
                <text
                  x={midX}
                  y={oy - 26}
                  textAnchor="middle"
                  fill="#ffffff"
                  fontSize={9}
                  fontFamily="monospace"
                  opacity={0.5}
                >
                  C/L {v.label}
                </text>
                <text
                  x={midX}
                  y={oy - 12}
                  textAnchor="middle"
                  fill="#3b82f6"
                  fontSize={11}
                  fontWeight={600}
                  fontFamily="monospace"
                >
                  DLO {dlo.toFixed(1)}&quot;
                </text>
              </g>
            );
          });
        })()}

        {/* Bay height labels (on right) — C/L and D.L.O. */}
        {/* Note: heights are stored bottom-to-top (index 0 = bottom row) */}
        {(() => {
          // SVG draws top-to-bottom, but bayHeights[0] = bottom row
          // The diagram draws from top, so reverse the display order
          let yAcc = 0;
          return horizontalLines.map((h, i) => {
            const midY = oy + (yAcc + bayHeights[i] / 2) * scale;
            // In SVG, row 0 drawn at top — but in our data row 0 = bottom
            // The diagram reversal: SVG index i corresponds to data row (baysTall - 1 - i) for bottom=0
            const dataRow = baysTall - 1 - i;
            const dlo = calculateDloHeight(bayHeights[i], dataRow, baysTall);
            yAcc += bayHeights[i];
            return (
              <g key={`hl-${i}`}>
                <text
                  x={ox + rectW + 10}
                  y={midY - 4}
                  textAnchor="start"
                  fill="#ffffff"
                  fontSize={9}
                  fontFamily="monospace"
                  opacity={0.5}
                >
                  C/L {h.label}
                </text>
                <text
                  x={ox + rectW + 10}
                  y={midY + 10}
                  textAnchor="start"
                  fill="#3b82f6"
                  fontSize={11}
                  fontWeight={600}
                  fontFamily="monospace"
                >
                  DLO {dlo.toFixed(1)}&quot;
                </text>
              </g>
            );
          });
        })()}

        {/* Overall width label (bottom) */}
        <line
          x1={ox}
          y1={oy + rectH + 20}
          x2={ox + rectW}
          y2={oy + rectH + 20}
          stroke="#ffffff"
          strokeWidth={0.75}
          markerStart="url(#arrowLeft)"
          markerEnd="url(#arrowRight)"
        />
        <text
          x={ox + rectW / 2}
          y={oy + rectH + 36}
          textAnchor="middle"
          fill="#ffffff"
          fontSize={12}
          fontWeight={500}
          fontFamily="monospace"
        >
          {openingWidth.toFixed(1)}&quot; W
        </text>

        {/* Overall height label (right side, further out) */}
        <line
          x1={ox + rectW + 70}
          y1={oy}
          x2={ox + rectW + 70}
          y2={oy + rectH}
          stroke="#ffffff"
          strokeWidth={0.75}
          markerStart="url(#arrowUp)"
          markerEnd="url(#arrowDown)"
        />
        <text
          x={ox + rectW + 74}
          y={oy + rectH / 2}
          textAnchor="start"
          fill="#ffffff"
          fontSize={12}
          fontWeight={500}
          fontFamily="monospace"
          transform={`rotate(90, ${ox + rectW + 74}, ${oy + rectH / 2})`}
        >
          {openingHeight.toFixed(1)}&quot; H
        </text>

        {/* Door placement rectangles */}
        {doorRects.map((door, i) => (
          <g key={`door-${i}`}>
            <rect
              x={ox + door.x}
              y={oy + door.y}
              width={door.width}
              height={door.height}
              fill="#34d399"
              fillOpacity={0.15}
              stroke="#34d399"
              strokeWidth={1.5}
              rx={2}
            />
            <text
              x={ox + door.x + door.width / 2}
              y={oy + door.y + door.height / 2}
              textAnchor="middle"
              dominantBaseline="central"
              fill="#34d399"
              fontSize={10}
              fontWeight={600}
              fontFamily="monospace"
            >
              {door.label}
            </text>
            {/* Door icon indicator */}
            <text
              x={ox + door.x + door.width / 2}
              y={oy + door.y + door.height / 2 + 14}
              textAnchor="middle"
              dominantBaseline="central"
              fill="#34d399"
              fontSize={8}
              fontFamily="monospace"
              opacity={0.7}
            >
              DOOR
            </text>
          </g>
        ))}

        {/* Arrow markers */}
        <defs>
          <marker
            id="arrowLeft"
            viewBox="0 0 6 6"
            refX={0}
            refY={3}
            markerWidth={6}
            markerHeight={6}
            orient="auto"
          >
            <path d="M6,0 L0,3 L6,6" fill="none" stroke="#ffffff" strokeWidth={1} />
          </marker>
          <marker
            id="arrowRight"
            viewBox="0 0 6 6"
            refX={6}
            refY={3}
            markerWidth={6}
            markerHeight={6}
            orient="auto"
          >
            <path d="M0,0 L6,3 L0,6" fill="none" stroke="#ffffff" strokeWidth={1} />
          </marker>
          <marker
            id="arrowUp"
            viewBox="0 0 6 6"
            refX={3}
            refY={0}
            markerWidth={6}
            markerHeight={6}
            orient="auto"
          >
            <path d="M0,6 L3,0 L6,6" fill="none" stroke="#ffffff" strokeWidth={1} />
          </marker>
          <marker
            id="arrowDown"
            viewBox="0 0 6 6"
            refX={3}
            refY={6}
            markerWidth={6}
            markerHeight={6}
            orient="auto"
          >
            <path d="M0,0 L3,6 L6,0" fill="none" stroke="#ffffff" strokeWidth={1} />
          </marker>
        </defs>
      </svg>
    </div>
  );
}
