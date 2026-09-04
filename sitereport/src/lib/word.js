import {
  AlignmentType,
  BorderStyle,
  Document,
  ImageRun,
  Packer,
  Paragraph,
  ShadingType,
  Table,
  TableCell,
  TableRow,
  TextRun,
  VerticalAlign,
  WidthType,
  convertMillimetersToTwip
} from 'docx';
import { blobToDataUrl } from './image.js';
import { layoutPhotoCollage, normalizeEntryPhotos } from './photos.js';
import { pdfEntryBadgeText } from './pdf-entry.js';

const CONTENT_WIDTH_MM = 180;
const CONTENT_TWIP = convertMillimetersToTwip(CONTENT_WIDTH_MM);
const CONTENT_WIDTH_PX = 680;
const MIN_PHOTO_CELL = 180;
const MAX_PHOTO_CELL = 280;
const PHOTO_GAP = 4;
const PHOTO_FRAME = 2;

const thinBorder = { style: BorderStyle.SINGLE, size: 4, color: 'D1D5DB' };
const cellBorders = {
  top: thinBorder,
  bottom: thinBorder,
  left: thinBorder,
  right: thinBorder
};

export async function exportToDocxData({
  protocolTitle,
  projectName,
  protocolDate,
  protocolDescription,
  attendees,
  logoDataUrl,
  columns,
  entries
}) {
  const issues = [];
  const addIssue = (message) => {
    if (issues.length < 20) issues.push(message);
  };

  const children = [];
  children.push(
    ...(await buildHeader({
      protocolTitle,
      projectName,
      protocolDate,
      protocolDescription,
      attendees,
      logoDataUrl,
      addIssue
    }))
  );

  const tableColumns = (columns || []).filter((col) => !col.isPhoto);
  let exportedEntries = 0;
  for (const [idx, entry] of (entries || []).entries()) {
    try {
      children.push(
        ...(await buildEntry({
          entry,
          index: idx,
          tableColumns,
          addIssue
        }))
      );
      exportedEntries += 1;
    } catch (err) {
      addIssue(`Eintrag ${idx + 1}: Word-Block konnte nicht erstellt werden (${err?.message || 'Unbekannt'}).`);
    }
  }

  if (!exportedEntries) {
    children.push(
      new Paragraph({
        spacing: { before: 200 },
        children: [new TextRun({ text: 'Keine Einträge.', italics: true, color: '6B7280' })]
      })
    );
  }

  const doc = new Document({
    creator: 'SiteReport',
    title: protocolTitle || 'Protokoll',
    styles: {
      default: {
        document: {
          run: {
            font: 'Calibri',
            size: 22
          }
        }
      }
    },
    sections: [
      {
        properties: {
          page: {
            size: {
              width: convertMillimetersToTwip(210),
              height: convertMillimetersToTwip(297)
            },
            margin: {
              top: convertMillimetersToTwip(15),
              bottom: convertMillimetersToTwip(15),
              left: convertMillimetersToTwip(15),
              right: convertMillimetersToTwip(15)
            }
          }
        },
        children
      }
    ]
  });

  const base64 = await Packer.toBase64String(doc);
  return {
    filename: buildDocxFilename(projectName, protocolDate),
    base64,
    stats: {
      format: 'docx',
      requestedEntries: (entries || []).length,
      exportedEntries,
      issues
    }
  };
}

export function buildDocxFilename(projectName, protocolDate) {
  const namePart = sanitizeFilename(projectName);
  const datePart = sanitizeFilename(protocolDate || new Date().toISOString().slice(0, 10));
  return `${namePart}_${datePart}.docx`;
}

async function buildHeader({
  protocolTitle,
  projectName,
  protocolDate,
  protocolDescription,
  attendees,
  logoDataUrl,
  addIssue
}) {
  const title = String(protocolTitle || 'Protokoll');
  const meta = [
    { label: 'Projekt', value: projectName || '', stacked: false },
    { label: 'Datum', value: protocolDate || '', stacked: false },
    { label: 'Beschreibung', value: protocolDescription || '', stacked: true },
    { label: 'Anwesende Personen', value: attendees || '', stacked: true }
  ];

  const textChildren = [
    new Paragraph({
      spacing: { after: 80 },
      children: [new TextRun({ text: title, bold: true, size: 32, color: '2C3E59' })]
    }),
    ...meta.flatMap((row) => {
      const value = String(row.value || '').trim() || '—';
      if (row.stacked) {
        return [
          new Paragraph({
            spacing: { before: 40 },
            children: [new TextRun({ text: `${row.label}:`, bold: true })]
          }),
          ...value.split(/\r?\n/).map(
            (line) =>
              new Paragraph({
                children: [new TextRun(line || ' ')]
              })
          )
        ];
      }
      return [
        new Paragraph({
          children: [
            new TextRun({ text: `${row.label}: `, bold: true }),
            new TextRun(value)
          ]
        })
      ];
    })
  ];

  const logo = await tryEmbedImage(logoDataUrl, { maxWidth: 140, maxHeight: 70, addIssue, label: 'Logo' });
  if (!logo) {
    return [
      headerBox(textChildren),
      new Paragraph({ spacing: { after: 200 }, children: [] })
    ];
  }

  const headerTable = new Table({
    width: { size: CONTENT_TWIP, type: WidthType.DXA },
    columnWidths: [CONTENT_TWIP - convertMillimetersToTwip(40), convertMillimetersToTwip(40)],
    rows: [
      new TableRow({
        children: [
          new TableCell({
            borders: noBorderSet(),
            width: { size: CONTENT_TWIP - convertMillimetersToTwip(40), type: WidthType.DXA },
            children: textChildren
          }),
          new TableCell({
            borders: noBorderSet(),
            width: { size: convertMillimetersToTwip(40), type: WidthType.DXA },
            verticalAlign: VerticalAlign.TOP,
            children: [
              new Paragraph({
                alignment: AlignmentType.RIGHT,
                children: [logo]
              })
            ]
          })
        ]
      })
    ]
  });

  return [headerTable, new Paragraph({ spacing: { after: 200 }, children: [] })];
}

function headerBox(textChildren) {
  return new Table({
    width: { size: CONTENT_TWIP, type: WidthType.DXA },
    rows: [
      new TableRow({
        children: [
          new TableCell({
            borders: cellBorders,
            shading: { type: ShadingType.CLEAR, fill: 'F3F4F6' },
            margins: { top: 80, bottom: 80, left: 120, right: 120 },
            children: textChildren
          })
        ]
      })
    ]
  });
}

async function buildEntry({ entry, index, tableColumns, addIssue }) {
  const blocks = [
    new Paragraph({ spacing: { before: 240 }, children: [] }),
    new Table({
      width: { size: convertMillimetersToTwip(32), type: WidthType.DXA },
      rows: [
        new TableRow({
          children: [
            new TableCell({
              borders: {
                top: { style: BorderStyle.SINGLE, size: 4, color: '2C3E59' },
                bottom: { style: BorderStyle.SINGLE, size: 4, color: '2C3E59' },
                left: { style: BorderStyle.SINGLE, size: 4, color: '2C3E59' },
                right: { style: BorderStyle.SINGLE, size: 4, color: '2C3E59' }
              },
              shading: { type: ShadingType.CLEAR, fill: '2C3E59' },
              margins: { top: 40, bottom: 40, left: 80, right: 80 },
              children: [
                new Paragraph({
                  children: [
                    new TextRun({
                      text: pdfEntryBadgeText(index),
                      bold: true,
                      color: 'FFFFFF',
                      size: 20
                    })
                  ]
                })
              ]
            })
          ]
        })
      ]
    }),
    new Paragraph({ spacing: { after: 80 }, children: [] })
  ];

  const photoTable = await buildPhotoTable(entry, index, addIssue);
  if (photoTable) {
    blocks.push(photoTable);
    blocks.push(new Paragraph({ spacing: { after: 80 }, children: [] }));
  }

  const labelTwip = Math.round(CONTENT_TWIP * 0.35);
  const valueTwip = CONTENT_TWIP - labelTwip;
  const fieldRows = tableColumns.map(
    (col, rowIdx) =>
      new TableRow({
        children: [
          new TableCell({
            borders: cellBorders,
            width: { size: labelTwip, type: WidthType.DXA },
            shading: { type: ShadingType.CLEAR, fill: rowIdx % 2 === 1 ? 'F3F4F6' : 'FFFFFF' },
            margins: { top: 40, bottom: 40, left: 80, right: 80 },
            children: [
              new Paragraph({
                children: [new TextRun({ text: col.name || '', bold: true })]
              })
            ]
          }),
          new TableCell({
            borders: cellBorders,
            width: { size: valueTwip, type: WidthType.DXA },
            shading: { type: ShadingType.CLEAR, fill: rowIdx % 2 === 1 ? 'F3F4F6' : 'FFFFFF' },
            margins: { top: 40, bottom: 40, left: 80, right: 80 },
            children: splitMultiline(entry.fields?.[col.name] ?? '')
          })
        ]
      })
  );

  if (fieldRows.length) {
    blocks.push(
      new Table({
        width: { size: CONTENT_TWIP, type: WidthType.DXA },
        columnWidths: [labelTwip, valueTwip],
        rows: fieldRows
      })
    );
  }

  return blocks;
}

async function buildPhotoTable(entry, index, addIssue) {
  const blobs = normalizeEntryPhotos(entry);
  if (!blobs.length) return null;

  const prepared = [];
  for (const [photoIndex, blob] of blobs.entries()) {
    try {
      const dataUrl = await blobToDataUrl(blob);
      const parsed = dataUrlToImage(dataUrl);
      if (!parsed) throw new Error('Ungültiges Bildformat');
      const size = await getImageSize(dataUrl);
      prepared.push({
        ...parsed,
        width: size.width,
        height: size.height
      });
    } catch (err) {
      addIssue(
        `Eintrag ${index + 1}, Bild ${photoIndex + 1}: konnte nicht eingebettet werden (${err?.message || 'Unbekannt'}).`
      );
    }
  }
  if (!prepared.length) return null;

  const collage = layoutPhotoCollage(
    prepared.map((img) => ({ width: img.width, height: img.height })),
    CONTENT_WIDTH_PX,
    2000,
    { gap: PHOTO_GAP, frame: PHOTO_FRAME, minCell: MIN_PHOTO_CELL, maxCell: MAX_PHOTO_CELL }
  );

  const cols = Math.max(1, collage.cols || 1);
  const rows = Math.max(1, collage.rows || 1);
  const colTwip = Math.floor(CONTENT_TWIP / cols);
  const tableRows = [];

  for (let row = 0; row < rows; row += 1) {
    const cells = [];
    for (let col = 0; col < cols; col += 1) {
      const photoIndex = row * cols + col;
      const item = collage.items[photoIndex];
      const image = prepared[photoIndex];
      if (!item || !image) {
        cells.push(
          new TableCell({
            borders: noBorderSet(),
            width: { size: colTwip, type: WidthType.DXA },
            children: [new Paragraph({ children: [] })]
          })
        );
        continue;
      }
      cells.push(
        new TableCell({
          borders: {
            top: thinBorder,
            bottom: thinBorder,
            left: thinBorder,
            right: thinBorder
          },
          width: { size: colTwip, type: WidthType.DXA },
          margins: { top: 40, bottom: 40, left: 40, right: 40 },
          verticalAlign: VerticalAlign.CENTER,
          children: [
            new Paragraph({
              alignment: AlignmentType.CENTER,
              children: [
                new ImageRun({
                  type: image.type,
                  data: image.bytes,
                  transformation: {
                    width: Math.max(1, Math.round(item.width)),
                    height: Math.max(1, Math.round(item.height))
                  },
                  altText: { title: `Eintrag ${index + 1}`, description: `Foto ${photoIndex + 1}`, name: `foto-${index + 1}-${photoIndex + 1}` }
                })
              ]
            })
          ]
        })
      );
    }
    tableRows.push(new TableRow({ children: cells }));
  }

  return new Table({
    width: { size: CONTENT_TWIP, type: WidthType.DXA },
    columnWidths: Array.from({ length: cols }, () => colTwip),
    rows: tableRows
  });
}

async function tryEmbedImage(dataUrl, { maxWidth, maxHeight, addIssue, label }) {
  if (!dataUrl) return null;
  try {
    const parsed = dataUrlToImage(dataUrl);
    if (!parsed) return null;
    const size = await getImageSize(dataUrl);
    const scale = Math.min(maxWidth / (size.width || 1), maxHeight / (size.height || 1), 1);
    return new ImageRun({
      type: parsed.type,
      data: parsed.bytes,
      transformation: {
        width: Math.max(1, Math.round(size.width * scale)),
        height: Math.max(1, Math.round(size.height * scale))
      },
      altText: { title: label, description: label, name: label }
    });
  } catch (err) {
    addIssue(`${label} konnte nicht eingebettet werden (${err?.message || 'Unbekannt'}).`);
    return null;
  }
}

function splitMultiline(value) {
  const raw = String(value ?? '').trim();
  const lines = raw ? raw.split(/\r?\n/) : ['—'];
  return lines.map(
    (line) =>
      new Paragraph({
        children: [new TextRun(line || ' ')]
      })
  );
}

function noBorderSet() {
  const none = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
  return { top: none, bottom: none, left: none, right: none };
}

function dataUrlToImage(dataUrl) {
  const match = String(dataUrl).match(/^data:image\/(png|jpe?g);base64,(.+)$/i);
  if (!match) return null;
  const type = match[1].toLowerCase() === 'png' ? 'png' : 'jpg';
  const binary = atob(match[2]);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i += 1) {
    bytes[i] = binary.charCodeAt(i);
  }
  return { bytes, type };
}

function getImageSize(dataUrl) {
  return new Promise((resolve) => {
    if (typeof Image === 'undefined') {
      resolve({ width: 1600, height: 900 });
      return;
    }
    const img = new Image();
    img.onload = () => resolve({ width: img.width || 1, height: img.height || 1 });
    img.onerror = () => resolve({ width: 1600, height: 900 });
    img.src = dataUrl;
  });
}

function sanitizeFilename(name) {
  const clean = (name || 'protokoll').replace(/[^a-z0-9\-_. ]/gi, '_').trim();
  return clean.length ? clean : 'protokoll';
}
