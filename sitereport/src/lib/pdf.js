import { PDFDocument, rgb, StandardFonts } from 'pdf-lib';
import { blobToDataUrl } from './image';
import { bufferToBase64 } from './native';
import { normalizeEntryPhotos } from './photos';
import {
  fitPdfPhotoCollage,
  naturalPdfPhotoCollage,
  pdfEntryBadgeText,
  planPdfEntryPlacement
} from './pdf-entry';

const A4_WIDTH = 595.28;
const A4_HEIGHT = 841.89;

export async function exportToPdfData({
  protocolTitle,
  projectName,
  protocolDate,
  protocolDescription,
  attendees,
  logoDataUrl,
  columns,
  entries
}) {
  const pdfDoc = await PDFDocument.create();
  const issues = [];
  const addIssue = (message) => {
    if (issues.length < 20) issues.push(message);
  };
  const font = await pdfDoc.embedFont(StandardFonts.Helvetica);
  const fontBold = await pdfDoc.embedFont(StandardFonts.HelveticaBold);

  const margin = 36;
  const lineHeight = 14;
  const headerGap = 18;
  const blockGap = 16;
  const tableGap = 12;
  const headerPadding = 14;
  const cardPadding = 12;
  const photoGap = 4;
  const photoFrame = 2;
  const badgeHeight = 18;
  const badgeGap = 8;
  const cornerColor = rgb(0.17, 0.24, 0.35);
  const softBorder = rgb(0.86, 0.88, 0.9);
  const softBg = rgb(0.97, 0.98, 0.99);
  const rowAlt = rgb(0.96, 0.97, 0.98);
  const frameColor = rgb(0.78, 0.8, 0.83);

  let page = null;
  let cursorY = 0;

  const wrapText = (text, maxWidth, size = 11, usedFont = font) => {
    const raw = pdfSafeText(text);
    if (!raw.trim()) return ['—'];
    const paragraphs = raw.replace(/\r\n/g, '\n').split('\n');
    const lines = [];
    for (const paragraph of paragraphs) {
      if (!paragraph) {
        lines.push(' ');
        continue;
      }
      const words = paragraph.split(/\s+/);
      let current = '';
      const pushLongWord = (word) => {
        let chunk = '';
        for (const ch of word) {
          const trial = chunk + ch;
          if (usedFont.widthOfTextAtSize(trial, size) <= maxWidth) {
            chunk = trial;
          } else {
            if (chunk) lines.push(chunk);
            chunk = ch;
          }
        }
        current = chunk;
      };
      for (const word of words) {
        const next = current ? `${current} ${word}` : word;
        if (usedFont.widthOfTextAtSize(next, size) <= maxWidth) {
          current = next;
        } else {
          if (current) lines.push(current);
          if (usedFont.widthOfTextAtSize(word, size) > maxWidth) {
            pushLongWord(word);
          } else {
            current = word;
          }
        }
      }
      if (current) lines.push(current);
    }
    return lines.length ? lines : ['—'];
  };

  const drawHeader = async () => {
    const pageWidth = A4_WIDTH;
    const pageHeight = A4_HEIGHT;
    const headerTop = pageHeight - margin;
    const headerWidth = pageWidth - margin * 2;

    let logoWidth = 0;
    let logoHeight = 0;
    let logoImage = null;
    if (logoDataUrl) {
      try {
        const { bytes, extension } = dataUrlToBytes(logoDataUrl);
        if (extension === 'png') {
          logoImage = await pdfDoc.embedPng(bytes);
        } else {
          logoImage = await pdfDoc.embedJpg(bytes);
        }
        const maxLogoW = 120;
        const maxLogoH = 60;
        const scale = Math.min(maxLogoW / logoImage.width, maxLogoH / logoImage.height, 1);
        logoWidth = logoImage.width * scale;
        logoHeight = logoImage.height * scale;
      } catch (err) {
        addIssue(`Logo konnte nicht eingebettet werden (${err?.message || 'Unbekannt'}).`);
      }
    }

    const title = pdfSafeText(protocolTitle || 'Protokoll');
    const titleSize = 16;
    const textMaxWidth = headerWidth - headerPadding * 2 - (logoImage ? logoWidth + 16 : 0);
    const stackedLabels = new Set(['Beschreibung', 'Anwesende Personen']);
    const metaRows = [
      { label: 'Projekt', value: projectName || '' },
      { label: 'Datum', value: protocolDate || '' },
      { label: 'Beschreibung', value: protocolDescription || '' },
      { label: 'Anwesende Personen', value: attendees || '' }
    ].map((row) => {
      const stacked = stackedLabels.has(row.label);
      const label = stacked ? `${row.label}:` : `${row.label}: `;
      const labelWidth = fontBold.widthOfTextAtSize(label, 11);
      const valueWidth = stacked ? textMaxWidth : Math.max(40, textMaxWidth - labelWidth);
      const valueLines = wrapText(row.value, valueWidth, 11, font).slice(0, 10);
      return { label, labelWidth, stacked, valueLines };
    });

    const titleLines = wrapText(title, textMaxWidth, titleSize, fontBold);
    const metaHeight = metaRows.reduce((sum, row) => {
      const valueH = row.valueLines.length * lineHeight;
      return sum + (row.stacked ? lineHeight + valueH : valueH);
    }, 0);
    const headerTextHeight = titleLines.length * (titleSize + 2) + 8 + metaHeight;
    const headerHeight = Math.max(headerTextHeight, logoHeight);
    const headerBoxHeight = headerHeight + headerPadding * 2;

    const headerBottom = headerTop - headerBoxHeight;
    page.drawRectangle({
      x: margin,
      y: headerBottom,
      width: headerWidth,
      height: headerBoxHeight,
      color: softBg,
      borderColor: softBorder,
      borderWidth: 1
    });

    page.drawRectangle({
      x: margin,
      y: headerTop - 4,
      width: headerWidth,
      height: 4,
      color: cornerColor
    });

    let textY = headerTop - headerPadding - titleSize;
    titleLines.forEach((line, index) => {
      page.drawText(line, {
        x: margin + headerPadding,
        y: textY - index * (titleSize + 2),
        size: titleSize,
        font: fontBold,
        color: cornerColor
      });
    });
    textY -= titleLines.length * (titleSize + 2) + 6;

    for (const row of metaRows) {
      page.drawText(row.label, {
        x: margin + headerPadding,
        y: textY,
        size: 11,
        font: fontBold,
        color: rgb(0.1, 0.1, 0.1)
      });
      if (row.stacked) {
        row.valueLines.forEach((line, i) => {
          page.drawText(line, {
            x: margin + headerPadding,
            y: textY - (i + 1) * lineHeight,
            size: 11,
            font,
            color: rgb(0.1, 0.1, 0.1)
          });
        });
        textY -= (1 + row.valueLines.length) * lineHeight;
      } else {
        row.valueLines.forEach((line, i) => {
          page.drawText(line, {
            x: margin + headerPadding + row.labelWidth,
            y: textY - i * lineHeight,
            size: 11,
            font,
            color: rgb(0.1, 0.1, 0.1)
          });
        });
        textY -= row.valueLines.length * lineHeight;
      }
    }

    if (logoImage) {
      const logoX = pageWidth - margin - headerPadding - logoWidth;
      const logoY = headerTop - headerPadding - logoHeight;
      page.drawImage(logoImage, {
        x: logoX,
        y: logoY,
        width: logoWidth,
        height: logoHeight
      });
    }

    return headerBottom - headerGap;
  };

  let pageCount = 0;
  let entriesOnCurrentPage = 0;
  const startPage = async () => {
    page = pdfDoc.addPage([A4_WIDTH, A4_HEIGHT]);
    pageCount += 1;
    entriesOnCurrentPage = 0;
    if (pageCount === 1) {
      cursorY = await drawHeader();
    } else {
      cursorY = A4_HEIGHT - margin;
    }
  };
  const remainingSpace = () => cursorY - margin;

  const tableColumns = columns.filter((c) => !c.isPhoto);
  const blockWidth = A4_WIDTH - margin * 2;
  const labelWidth = blockWidth * 0.35;
  const valueWidth = blockWidth - labelWidth - 8;

  const embedPhoto = async (blob) => {
    const dataUrl = await blobToDataUrl(blob);
    const { bytes, extension } = dataUrlToBytes(dataUrl);
    return extension === 'png' ? pdfDoc.embedPng(bytes) : pdfDoc.embedJpg(bytes);
  };

  const fitCollage = (sizes, maxImageWidth, maxImageHeight) =>
    fitPdfPhotoCollage(sizes, maxImageWidth, maxImageHeight);

  const prepareEntry = async (entry, index) => {
    const photoBlobs = normalizeEntryPhotos(entry);
    const images = [];
    for (const [photoIndex, blob] of photoBlobs.entries()) {
      try {
        images.push(await embedPhoto(blob));
      } catch (err) {
        addIssue(
          `Eintrag ${index + 1}, Bild ${photoIndex + 1}: konnte nicht eingebettet werden (${err?.message || 'Unbekannt'}).`
        );
      }
    }

    const rows = tableColumns.map((col) => ({
      label: col.name,
      value: entry.fields?.[col.name] ?? ''
    }));

    let tableHeight = 0;
    const rowHeights = rows.map((row) => {
      const labelLines = wrapText(row.label, labelWidth - 6, 11, fontBold);
      const valueLines = wrapText(row.value, valueWidth - 6, 11, font);
      const lines = Math.max(labelLines.length, valueLines.length, 1);
      const height = lines * lineHeight + 6;
      tableHeight += height;
      return { height, labelLines, valueLines };
    });

    const hasImages = images.length > 0;
    const imageChrome = hasImages ? tableGap : 0;
    const chromeWithoutImage = cardPadding * 2 + badgeHeight + badgeGap + imageChrome + tableHeight;
    const sizes = images.map((image) => ({ width: image.width, height: image.height }));
    const naturalCollage = hasImages
      ? naturalPdfPhotoCollage(sizes)
      : { items: [], width: 0, height: 0, cols: 0, rows: 0 };

    return {
      images,
      sizes,
      hasImages,
      rowHeights,
      tableHeight,
      chromeWithoutImage,
      naturalCardHeight: chromeWithoutImage + (hasImages ? naturalCollage.height : 0)
    };
  };

  const drawPreparedEntry = async (prepared, index, maxCardHeight) => {
    const maxImageWidth = blockWidth - cardPadding * 2;
    let collage = { items: [], width: 0, height: 0, cols: 0, rows: 0 };
    let imageHeight = 0;
    if (prepared.hasImages) {
      const maxImageHeight = Math.max(56, Math.min(maxCardHeight, remainingSpace()) - prepared.chromeWithoutImage);
      collage = fitCollage(prepared.sizes, maxImageWidth, maxImageHeight);
      imageHeight = collage.height;
    }

    const cardHeight = prepared.chromeWithoutImage + imageHeight;
    const cardTop = cursorY;
    const cardBottom = cursorY - cardHeight;

    page.drawRectangle({
      x: margin,
      y: cardBottom,
      width: blockWidth,
      height: cardHeight,
      color: softBg,
      borderColor: softBorder,
      borderWidth: 1
    });

    const badgeText = pdfEntryBadgeText(index);
    const badgePaddingX = 8;
    const badgeWidth = Math.min(
      blockWidth - cardPadding * 2,
      fontBold.widthOfTextAtSize(badgeText, 10) + badgePaddingX * 2
    );
    const badgeX = margin + cardPadding;
    const badgeY = cardTop - cardPadding - badgeHeight;
    page.drawRectangle({
      x: badgeX,
      y: badgeY,
      width: badgeWidth,
      height: badgeHeight,
      color: cornerColor
    });
    page.drawText(badgeText, {
      x: badgeX + badgePaddingX,
      y: badgeY + 4,
      size: 10,
      font: fontBold,
      color: rgb(1, 1, 1)
    });

    const collageTop = badgeY - badgeGap;
    const collageLeft = margin + cardPadding;
    const imageY = collageTop - imageHeight;
    const tableTopStart = prepared.hasImages ? imageY - tableGap : collageTop;

    if (prepared.hasImages) {
      collage.items.forEach((item, photoIndex) => {
        const frameX = collageLeft + item.frameX;
        const frameY = collageTop - item.frameY - item.frameH;
        page.drawRectangle({
          x: frameX,
          y: frameY,
          width: item.frameW,
          height: item.frameH,
          borderColor: frameColor,
          borderWidth: 1,
          color: rgb(1, 1, 1)
        });
        const image = prepared.images[photoIndex];
        if (!image) return;
        page.drawImage(image, {
          x: collageLeft + item.x,
          y: collageTop - item.y - item.height,
          width: item.width,
          height: item.height
        });
      });
    }

    let tableTop = tableTopStart;
    const tableLeft = margin + cardPadding;
    const tableRight = margin + blockWidth - cardPadding;
    const dividerX = tableLeft + labelWidth + 4;
    const textPadding = 4;

    page.drawLine({
      start: { x: tableLeft, y: tableTop },
      end: { x: tableRight, y: tableTop },
      thickness: 1,
      color: softBorder
    });

    prepared.rowHeights.forEach((rowData, idx) => {
      const rowBottom = tableTop - rowData.height;
      if (idx % 2 === 1) {
        page.drawRectangle({
          x: tableLeft,
          y: rowBottom,
          width: tableRight - tableLeft,
          height: rowData.height,
          color: rowAlt
        });
      }

      const labelY = tableTop - lineHeight;
      rowData.labelLines.forEach((line, i) => {
        page.drawText(line, {
          x: tableLeft + textPadding,
          y: labelY - i * lineHeight,
          size: 11,
          font: fontBold,
          color: rgb(0.1, 0.1, 0.1)
        });
      });

      const valueY = tableTop - lineHeight;
      rowData.valueLines.forEach((line, i) => {
        page.drawText(line, {
          x: dividerX + textPadding,
          y: valueY - i * lineHeight,
          size: 11,
          font,
          color: rgb(0.1, 0.1, 0.1)
        });
      });

      page.drawLine({
        start: { x: tableLeft, y: rowBottom },
        end: { x: tableRight, y: rowBottom },
        thickness: 1,
        color: softBorder
      });

      tableTop = rowBottom;
    });

    page.drawLine({
      start: { x: dividerX, y: tableTopStart },
      end: { x: dividerX, y: tableTopStart - prepared.tableHeight },
      thickness: 1,
      color: softBorder
    });

    cursorY = tableTop - blockGap;
    entriesOnCurrentPage += 1;
  };

  await startPage();
  let exportedEntries = 0;
  for (const [idx, entry] of entries.entries()) {
    try {
      const prepared = await prepareEntry(entry, idx);
      let plan = planPdfEntryPlacement({
        remaining: remainingSpace(),
        isFirstDocumentPage: pageCount === 1,
        entriesOnPage: entriesOnCurrentPage,
        photoCount: prepared.images.length,
        naturalCardHeight: prepared.naturalCardHeight,
        chromeHeight: prepared.chromeWithoutImage
      });
      if (!plan.stayOnPage) {
        await startPage();
        plan = planPdfEntryPlacement({
          remaining: remainingSpace(),
          isFirstDocumentPage: false,
          entriesOnPage: 0,
          photoCount: prepared.images.length,
          naturalCardHeight: prepared.naturalCardHeight,
          chromeHeight: prepared.chromeWithoutImage
        });
      }
      const maxCardHeight = Math.min(plan.maxCardHeight, Math.max(prepared.chromeWithoutImage, remainingSpace()));
      await drawPreparedEntry(prepared, idx, maxCardHeight);
      exportedEntries += 1;
    } catch (err) {
      addIssue(`Eintrag ${idx + 1}: PDF-Block konnte nicht erstellt werden (${err?.message || 'Unbekannt'}).`);
    }
  }

  const pdfBytes = await pdfDoc.save();
  const filename = buildPdfFilename(projectName, protocolDate);
  const base64 = await bufferToBase64(pdfBytes);
  return {
    filename,
    base64,
    stats: {
      format: 'pdf',
      requestedEntries: entries.length,
      exportedEntries,
      issues
    }
  };
}

function pdfSafeText(text) {
  const map = {
    '\u2018': "'",
    '\u2019': "'",
    '\u201c': '"',
    '\u201d': '"',
    '\u2013': '-',
    '\u2014': '-',
    '\u2026': '...',
    '\u00a0': ' ',
    '\u2022': '-',
    '\u20ac': 'EUR'
  };
  return Array.from(String(text ?? ''), (ch) => {
    if (ch === '\n' || ch === '\r') return ch;
    if (map[ch]) return map[ch];
    const code = ch.charCodeAt(0);
    if (code === 9) return ' ';
    if (code >= 0x20 && code <= 0x7e) return ch;
    if (code >= 0xa0 && code <= 0xff) return ch;
    return '?';
  }).join('');
}

function dataUrlToBytes(dataUrl) {
  const match = String(dataUrl).match(/^data:image\/(png|jpe?g);base64,(.+)$/i);
  if (!match) {
    return { bytes: new Uint8Array(), extension: 'jpg' };
  }
  const extension = match[1].toLowerCase() === 'png' ? 'png' : 'jpg';
  const binary = atob(match[2]);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i += 1) {
    bytes[i] = binary.charCodeAt(i);
  }
  return { bytes, extension };
}

function sanitizeFilename(name) {
  const clean = (name || 'protokoll').replace(/[^a-z0-9\-_. ]/gi, '_').trim();
  return clean.length ? clean : 'protokoll';
}

function buildPdfFilename(projectName, protocolDate) {
  const namePart = sanitizeFilename(projectName);
  const datePart = sanitizeFilename(protocolDate || new Date().toISOString().slice(0, 10));
  return `${namePart}_${datePart}.pdf`;
}
