import { createClientFromRequest } from 'npm:@base44/sdk@0.8.54';
import * as XLSX from 'npm:xlsx@0.18.5';
import { unzipSync } from 'npm:fflate@0.8.2';

const MAX_URL_LENGTH = 600;
const MAX_ROWS = 5000;
const MAX_BYTES = 20 * 1024 * 1024;
const MAX_IMAGES = 200;
const MAX_IMAGE_BYTES = 8 * 1024 * 1024;
const MAX_IMAGE_PAYLOAD = 6 * 1024 * 1024;

const IMAGE_TYPES = {
  png: 'image/png',
  jpg: 'image/jpeg',
  jpeg: 'image/jpeg',
  gif: 'image/gif',
  webp: 'image/webp',
};

function resolvePath(baseDir, target) {
  const raw = String(target || '');
  const parts = (raw.startsWith('/') ? raw.slice(1) : `${baseDir}/${raw}`).split('/');
  const out = [];
  for (const part of parts) {
    if (!part || part === '.') continue;
    if (part === '..') out.pop();
    else out.push(part);
  }
  return out.join('/');
}

// Relationships of a drawing file: rId -> media path inside the workbook zip.
function drawingRels(zip, drawingPath) {
  const dir = drawingPath.slice(0, drawingPath.lastIndexOf('/'));
  const name = drawingPath.slice(drawingPath.lastIndexOf('/') + 1);
  const relsPath = `${dir}/_rels/${name}.rels`;
  const raw = zip[relsPath] ? new TextDecoder().decode(zip[relsPath]) : '';
  const map = {};
  for (const tag of raw.match(/<Relationship\b[^>]*>/g) || []) {
    const id = tag.match(/Id="([^"]+)"/);
    const target = tag.match(/Target="([^"]+)"/);
    if (id && target) map[id[1]] = resolvePath(dir, target[1]);
  }
  return map;
}

// Each anchored picture in a drawing, with the sheet row it sits on (1-based).
function imageAnchors(xml) {
  const starts = [
    ...xml.matchAll(/<(?:[A-Za-z0-9]+:)?(?:oneCellAnchor|twoCellAnchor|absoluteAnchor)[\s>]/gi),
  ].map((match) => match.index);
  const out = [];
  starts.forEach((start, index) => {
    const chunk = xml.slice(start, index + 1 < starts.length ? starts[index + 1] : xml.length);
    const row = chunk.match(/<(?:[A-Za-z0-9]+:)?row>\s*(\d+)\s*<\//i);
    const embed = chunk.match(/<[A-Za-z0-9]*:?blip\b[^>]*?embed="([^"]+)"/i);
    if (row && embed) out.push({ row: Number(row[1]) + 1, rel: embed[1] });
  });
  return out;
}

const decode = (data) => new TextDecoder().decode(data);

// Pictures placed inside a cell (Excel's "place in cell") carry a value-metadata
// pointer on the cell itself; their images live in richValueRel.xml.
function cellImageRows(zip, sheetPath) {
  const path = zip[sheetPath]
    ? sheetPath
    : Object.keys(zip)
        .filter((key) => /^xl\/worksheets\/sheet\d+\.xml$/i.test(key))
        .sort()[0];
  const xml = path && zip[path] ? decode(zip[path]) : '';
  const rows = [];
  for (const tag of xml.match(/<c\b[^>]*>/g) || []) {
    if (!/\bvm="/.test(tag)) continue;
    const ref = tag.match(/\br="[A-Z]+(\d+)"/);
    if (ref) rows.push(Number(ref[1]));
  }
  return rows;
}

function richMediaPaths(zip) {
  const relsXml = zip['xl/richData/richValueRel.xml'];
  if (!relsXml) return [];
  const relsPath = 'xl/richData/_rels/richValueRel.xml.rels';
  const relsMap = {};
  const relTags = (zip[relsPath] ? decode(zip[relsPath]) : '').match(/<Relationship\b[^>]*>/g) || [];
  for (const tag of relTags) {
    const id = tag.match(/Id="([^"]+)"/);
    const target = tag.match(/Target="([^"]+)"/);
    if (id && target) relsMap[id[1]] = resolvePath('xl/richData', target[1]);
  }
  const out = [];
  for (const tag of decode(relsXml).match(/<rel\b[^>]*>/g) || []) {
    const id = tag.match(/r:id="([^"]+)"/);
    if (id && relsMap[id[1]]) out.push(relsMap[id[1]]);
  }
  return out;
}

function toBase64(bytes) {
  let out = '';
  for (let i = 0; i < bytes.length; i += 0x8000) {
    out += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
  }
  return btoa(out);
}

// Pictures embedded in the workbook, matched to the rows they are anchored on,
// returned as base64 so the page can upload them.
async function extractImages(bytes, sheetPath) {
  let zip;
  try {
    zip = unzipSync(new Uint8Array(bytes));
  } catch (error) {
    return { images: [], media: 0 };
  }
  if (!zip) return { images: [], media: 0 };

  const mediaFiles = Object.keys(zip).filter((path) => /^xl\/media\/[^/]+$/i.test(path));
  const found = [];

  const drawings = Object.keys(zip)
    .filter((path) => /^xl\/drawings\/drawing\d+\.xml$/i.test(path))
    .sort();

  for (const path of drawings) {
    const xml = new TextDecoder().decode(zip[path]);
    const rels = drawingRels(zip, path);
    for (const anchor of imageAnchors(xml)) {
      const mediaPath = rels[anchor.rel];
      const data = mediaPath ? zip[mediaPath] : null;
      const ext = String(mediaPath || '').split('.').pop().toLowerCase();
      const type = IMAGE_TYPES[ext];
      if (!data || !type || data.length > MAX_IMAGE_BYTES) continue;
      found.push({
        row: anchor.row,
        data,
        type,
        name: mediaPath.split('/').pop() || `sheet-image.${ext}`,
      });
    }
  }

  // Pictures stored in the cells themselves: pair each image-bearing cell with
  // its picture in the order Excel wrote them.
  const cellRows = cellImageRows(zip, sheetPath);
  const richMedia = richMediaPaths(zip);
  cellRows.forEach((row, index) => {
    const mediaPath = richMedia[index];
    const data = mediaPath ? zip[mediaPath] : null;
    const ext = String(mediaPath || '').split('.').pop().toLowerCase();
    const type = IMAGE_TYPES[ext];
    if (!data || !type || data.length > MAX_IMAGE_BYTES) return;
    found.push({
      row,
      data,
      type,
      name: mediaPath.split('/').pop() || `sheet-image.${ext}`,
    });
  });

  const seen = new Set();
  const unique = found.filter((image) => {
    if (seen.has(image.row) || seen.size >= MAX_IMAGES) return false;
    seen.add(image.row);
    return true;
  });

  const images = [];
  let total = 0;
  for (const image of unique) {
    const data = toBase64(image.data);
    if (total + data.length > MAX_IMAGE_PAYLOAD) break;
    total += data.length;
    images.push({ row: image.row, type: image.type, name: image.name, data });
  }

  return { images, media: mediaFiles.length };
}

// Row numbers and cells of the first sheet, keeping the true sheet row of each
// line so pictures can be matched to the item they belong to.
function sheetRows(sheet) {
  const range = XLSX.utils.decode_range(sheet['!ref'] || 'A1:A1');
  const rows = [];
  const rowNumbers = [];
  for (let R = range.s.r; R <= range.e.r; R += 1) {
    const cells = [];
    for (let C = range.s.c; C <= range.e.c; C += 1) {
      const cell = sheet[XLSX.utils.encode_cell({ r: R, c: C })];
      cells.push(
        String(cell?.w ?? cell?.v ?? '')
          .replace(/[\t\r\n]+/g, ' ')
          .trim()
      );
    }
    if (!cells.some(Boolean)) continue;
    rows.push(cells);
    rowNumbers.push(R + 1);
  }
  return { rows, rowNumbers };
}

// Reads an uploaded Excel/CSV file and returns its first sheet as tab separated
// text, plus any pictures embedded in the sheet keyed to their row.
export default async function (req) {
  try {
    const base44 = createClientFromRequest(req);
    const user = await base44.auth
      .me()
      .catch((e) => (e.status === 401 ? null : Promise.reject(e)));
    if (!user) return Response.json({ error: 'Unauthorized' }, { status: 401 });

    const body = await req.json().catch(() => ({}));
    const fileUrl = typeof body.file_url === 'string' ? body.file_url.trim() : '';
    if (!fileUrl || fileUrl.length > MAX_URL_LENGTH || !/^https?:\/\//i.test(fileUrl)) {
      return Response.json({ error: 'A spreadsheet file is required.' }, { status: 400 });
    }

    const download = await fetch(fileUrl, { signal: AbortSignal.timeout(20000) });
    if (!download.ok) {
      return Response.json(
        { error: 'That file could not be downloaded. Try uploading it again.' },
        { status: 422 }
      );
    }

    const bytes = new Uint8Array(await download.arrayBuffer());
    if (!bytes.length || bytes.length > MAX_BYTES) {
      return Response.json(
        { error: 'That file is empty or too large to read.' },
        { status: 422 }
      );
    }

    let workbook;
    try {
      workbook = XLSX.read(bytes, { type: 'array' });
    } catch (readError) {
      return Response.json(
        {
          error:
            'That file is not a spreadsheet we can read. Save it as .xlsx or .csv and upload it again.',
        },
        { status: 422 }
      );
    }

    const sheetName = workbook?.SheetNames?.[0];
    const sheet = sheetName ? workbook.Sheets[sheetName] : null;
    if (!sheet) {
      return Response.json({ error: 'That workbook has no sheets to read.' }, { status: 422 });
    }

    const { rows, rowNumbers } = sheetRows(sheet);

    if (!rows.length) {
      return Response.json({ error: 'No rows found in that sheet.' }, { status: 422 });
    }
    if (rows.length > MAX_ROWS) {
      return Response.json(
        { error: 'That sheet has too many rows — split it into smaller files.' },
        { status: 422 }
      );
    }

    const sheetIndex = Math.max(0, workbook.SheetNames.indexOf(sheetName));
    const sheetPath = `xl/worksheets/sheet${sheetIndex + 1}.xml`;

    const { images, media } = await extractImages(bytes, sheetPath);

    return Response.json({
      text: rows.map((row) => row.join('\t')).join('\n'),
      rows: rows.length,
      sheet: sheetName,
      row_numbers: rowNumbers,
      media,
      images,
    });
  } catch (error) {
    return Response.json({ error: error.message }, { status: 500 });
  }
}