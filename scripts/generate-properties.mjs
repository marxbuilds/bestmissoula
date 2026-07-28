// Reads the Properties tab from the lodging workbook and generates one
// .md file per verified, non-closed property into src/content/properties/.
//
// Usage: node scripts/generate-properties.mjs path/to/workbook.xlsx

import XLSX from 'xlsx';
import fs from 'fs';
import path from 'path';

const inputPath = process.argv[2];
if (!inputPath) {
  console.error('Usage: node scripts/generate-properties.mjs <path-to-xlsx>');
  process.exit(1);
}

const outDir = 'src/content/properties';

// --- Column positions on the Properties tab (0-indexed, row 0 = header) ---
// If you ever add/reorder columns in Excel, these numbers need updating too.
const COL = {
  name: 0,
  town: 1,
  region: 2,
  propertyType: 3,
  priceTier: 4,
  priceDisplay: 5,
  amenities: 6,
  roomTypes: 7,
  phone: 8,
  website: 9,
  onBooking: 10,
  description: 11,
  verified: 12,
  distance: 13,
};

function slugify(name) {
  return name
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/^-+|-+$/g, '');
}

function cleanValue(v) {
  if (v === undefined || v === null) return undefined;
  const s = String(v).trim();
  if (s === '' || s.toLowerCase() === 'none') return undefined;
  return s;
}

function splitList(v) {
  const s = cleanValue(v);
  if (!s) return [];
  return s.split(',').map((x) => x.trim()).filter(Boolean);
}

// Escapes a string for safe use inside a double-quoted YAML value.
function yamlString(s) {
  return `"${s.replace(/\\/g, '\\\\').replace(/"/g, '\\"')}"`;
}

function yamlStringList(arr) {
  if (arr.length === 0) return '[]';
  return '\n' + arr.map((item) => `  - ${yamlString(item)}`).join('\n');
}

// --- Read the workbook ---
const workbook = XLSX.readFile(inputPath);
const sheet = workbook.Sheets['Properties'];
const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false });

// Row 0 = header, data starts at Row 1
const dataRows = rows.slice(1);

fs.mkdirSync(outDir, { recursive: true });

let written = 0;
let skipped = 0;

for (const row of dataRows) {
  const name = cleanValue(row[COL.name]);
  if (!name) continue; // blank row

  const verified = cleanValue(row[COL.verified]) === 'Yes';
  if (!verified) {
    skipped++;
    continue; // skips Closed, blank, or anything not confirmed open
  }

  const description = cleanValue(row[COL.description]);
  if (!description) {
    console.warn(`Skipping "${name}" - no Description in spreadsheet yet.`);
    skipped++;
    continue;
  }

  const frontmatter = {
    name,
    town: cleanValue(row[COL.town]) || '',
    region: cleanValue(row[COL.region]) || '',
    propertyType: cleanValue(row[COL.propertyType]) || '',
    priceTier: cleanValue(row[COL.priceTier]) || '$',
    priceDisplay: cleanValue(row[COL.priceDisplay]) || 'Contact for pricing',
    amenities: splitList(row[COL.amenities]),
    roomTypes: cleanValue(row[COL.roomTypes]),
    phone: cleanValue(row[COL.phone]),
    website: cleanValue(row[COL.website]),
    distance: cleanValue(row[COL.distance]),
    description,
    verified: true,
  };

  const lines = ['---'];
  for (const [key, value] of Object.entries(frontmatter)) {
    if (value === undefined) continue; // omit optional fields that are blank
    if (Array.isArray(value)) {
      lines.push(`${key}:${yamlStringList(value)}`);
    } else if (typeof value === 'boolean') {
      lines.push(`${key}: ${value}`);
    } else {
      lines.push(`${key}: ${yamlString(value)}`);
    }
  }
  lines.push('---', '');

  const slug = slugify(name);
  fs.writeFileSync(path.join(outDir, `${slug}.md`), lines.join('\n'));
  written++;
}

console.log(`Wrote ${written} property file(s), skipped ${skipped}.`);