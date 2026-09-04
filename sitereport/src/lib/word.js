import {
  AlignmentType,
  BorderStyle,
  Document,
  ImageRun,
  Packer,
  PageBreak,
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
import { normalizeEntryPhotos } from './photos.js';
import {
  planExportEntryFlow,
  fitPdfPhotoCollage,
  pdfEntryBadgeText,
  pdfPhotoAreaWidth
} from './pdf-entry.js';

const PAGE_MARGIN_MM = 12.7;
const CONTENT_WIDTH_MM = 210 - PAGE_MARGIN_MM * 2;
const CONTENT_TWIP = convertMillimetersToTwip(CONTENT_WIDTH_MM);
const PT_TO_PX = 96 / 72;
const CARD_PADDING_TWIP = 80;
const MAX_HEADER_STACKED_LINES = 5;

const thinBorder = { style: BorderStyle.SINGLE, size: 4, color: 'D1D5DB' };
const cardBorder = { style: BorderStyle.SINGLE, size: 18, color: '4B5563' };
const noneBorder = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
const cellBorders = {
  top: thinBorder,
  bottom: thinBorder,
  left: thinBorder,
  right: thinBorder
};
const cardCellBorders = {
  top: cardBorder,
  bottom: cardBorder,
  left: cardBorder,
  right: cardBorder
};
const noBorders = {
  top: noneBorder,
  bottom: noneBorder,
  left: noneBorder,
  right: noneBorder
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
  const list = entries || [];
  const prepared = [];
  for (const [idx, entry] of list.entries()) {
    prepared.push(await prepareEntryPhotos(entry, idx, addIssue));
  }

  const flow = planExportEntryFlow({
    entries: list,
    tableColumns,
    photoSizesForEntry: (_entry, idx) => prepared[idx]?.sizes || [],
    protocolTitle,
    protocolDescription,
    attendees,
    hasLogo: Boolean(logoDataUrl)
  });

  let exportedEntries = 0;
  for (const [idx, entry] of list.entries()) {
    try {
      const plan = flow[idx] || {};
      if (plan.pageBreakBefore) {
        children.push(
          new Paragraph({
            children: [new PageBreak()]
          })
        );
      }
      children.push(
        ...(await buildEntry({
          entry,
          index: idx,
          tableColumns,
          photos: prepared[idx],
          maxImageHeightPt: plan.maxImageHeight || 0
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
              top: convertMillimetersToTwip(PAGE_MARGIN_MM),
              bottom: convertMillimetersToTwip(PAGE_MARGIN_MM),
              left: convertMillimetersToTwip(PAGE_MARGIN_MM),
              right: convertMillimetersToTwip(PAGE_MARGIN_MM)
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
      requestedEntries: list.length,
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
      keepNext: true,
      spacing: { after: 60 },
      children: [new TextRun({ text: title, bold: true, size: 32, color: '2C3E59' })]
    }),
    ...meta.flatMap((row) => {
      const value = String(row.value || '').trim() || '—';
      if (row.stacked) {
        const lines = value.split(/\r?\n/).slice(0, MAX_HEADER_STACKED_LINES);
        return [
          new Paragraph({
            keepNext: true,
            spacing: { before: 20 },
            children: [new TextRun({ text: `${row.label}:`, bold: true })]
          }),
          ...lines.map(
            (line) =>
              new Paragraph({
                keepNext: true,
                children: [new TextRun(line || ' ')]
              })
          )
        ];
      }
      return [
        new Paragraph({
          keepNext: true,
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
      new Paragraph({ keepNext: true, spacing: { after: 120 }, children: [] })
    ];
  }

  const headerTable = new Table({
    width: { size: CONTENT_TWIP, type: WidthType.DXA },
    columnWidths: [CONTENT_TWIP - convertMillimetersToTwip(40), convertMillimetersToTwip(40)],
    rows: [
      new TableRow({
        children: [
          new TableCell({
            borders: noBorders,
            width: { size: CONTENT_TWIP - convertMillimetersToTwip(40), type: WidthType.DXA },
            children: textChildren
          }),
          new TableCell({
            borders: noBorders,
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

  return [headerTable, new Paragraph({ keepNext: true, spacing: { after: 120 }, children: [] })];
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
            margins: { top: 60, bottom: 60, left: 80, right: 80 },
            children: textChildren
          })
        ]
      })
    ]
  });
}

async function buildEntry({ entry, index, tableColumns, photos, maxImageHeightPt }) {
  const inner = [buildBadgeTable(index)];

  const photoTable = buildPhotoTable(photos, index, maxImageHeightPt);
  if (photoTable) {
    inner.push(new Paragraph({ spacing: { after: 80 }, children: [] }));
    inner.push(photoTable);
  }

  const fieldTable = buildFieldTable(entry, tableColumns);
  if (fieldTable) {
    inner.push(new Paragraph({ spacing: { after: 80 }, children: [] }));
    inner.push(fieldTable);
  }

  return [
    new Table({
      width: { size: CONTENT_TWIP, type: WidthType.DXA },
      borders: {
        top: cardBorder,
        bottom: cardBorder,
        left: cardBorder,
        right: cardBorder,
        insideHorizontal: noneBorder,
        insideVertical: noneBorder
      },
      rows: [
        new TableRow({
          cantSplit: true,
          children: [
            new TableCell({
              borders: cardCellBorders,
              shading: { type: ShadingType.CLEAR, fill: 'F3F4F6' },
              margins: {
                top: CARD_PADDING_TWIP,
                bottom: CARD_PADDING_TWIP,
                left: CARD_PADDING_TWIP,
                right: CARD_PADDING_TWIP
              },
              children: inner
            })
          ]
        })
      ]
    }),
    new Paragraph({ spacing: { after: 160 }, children: [] })
  ];
}

function buildBadgeTable(index) {
  return new Table({
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
  });
}

function buildFieldTable(entry, tableColumns) {
  if (!tableColumns.length) return null;
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

  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    columnWidths: [labelTwip, valueTwip],
    rows: fieldRows
  });
}

async function prepareEntryPhotos(entry, index, addIssue) {
  const blobs = normalizeEntryPhotos(entry);
  const images = [];
  for (const [photoIndex, blob] of blobs.entries()) {
    try {
      const dataUrl = await blobToDataUrl(blob);
      const parsed = dataUrlToImage(dataUrl);
      if (!parsed) throw new Error('Ungültiges Bildformat');
      const size = await getImageSize(dataUrl);
      images.push({
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
  return {
    images,
    sizes: images.map((img) => ({ width: img.width, height: img.height }))
  };
}

function buildPhotoTable(photos, index, maxImageHeightPt) {
  const images = photos?.images || [];
  if (!images.length) return null;

  const collage = fitPdfPhotoCollage(photos.sizes, pdfPhotoAreaWidth(), maxImageHeightPt);
  const cols = Math.max(1, collage.cols || 1);
  const rows = Math.max(1, collage.rows || 1);
  const colTwip = Math.floor(CONTENT_TWIP / cols);
  const tableRows = [];

  for (let row = 0; row < rows; row += 1) {
    const cells = [];
    for (let col = 0; col < cols; col += 1) {
      const photoIndex = row * cols + col;
      const item = collage.items[photoIndex];
      const image = images[photoIndex];
      if (!item || !image) {
        cells.push(
          new TableCell({
            borders: noBorders,
            width: { size: colTwip, type: WidthType.DXA },
            children: [new Paragraph({ children: [] })]
          })
        );
        continue;
      }
      cells.push(
        new TableCell({
          borders: cellBorders,
          width: { size: colTwip, type: WidthType.DXA },
          margins: { top: 20, bottom: 20, left: 20, right: 20 },
          verticalAlign: VerticalAlign.CENTER,
          children: [
            new Paragraph({
              alignment: AlignmentType.CENTER,
              children: [
                new ImageRun({
                  type: image.type,
                  data: image.bytes,
                  transformation: {
                    width: Math.max(1, Math.round(item.width * PT_TO_PX)),
                    height: Math.max(1, Math.round(item.height * PT_TO_PX))
                  },
                  altText: {
                    title: `Eintrag ${index + 1}`,
                    description: `Foto ${photoIndex + 1}`,
                    name: `foto-${index + 1}-${photoIndex + 1}`
                  }
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
    width: { size: 100, type: WidthType.PERCENTAGE },
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
