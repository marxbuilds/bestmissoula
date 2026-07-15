// Reads the Outfitters tab from the research workbook and generates one
// .md file per verified, non-closed outfitter into src/content/outfitters/.
//
// Usage: node scripts/generate-outfitters.mjs path/to/workbook.xlsx

import XLSX from 'xlsx';
import fs from 'fs';
import path from 'path';

const inputPath = process.argv[2];
if (!inputPath) {
  console.error('Usage: node scripts/generate-outfitters.mjs <path-to-xlsx>');
  process.exit(1);
}

const outDir = 'src/content/outfitters';

// --- Column positions on the Outfitters tab (0-indexed, row 2 = header) ---
// If you ever add/reorder columns in Excel, these numbers need updating too.
const COL = {
  name: 0,
  baseTown: 1,
  region: 2,
  website: 3,
  fishingType: 4,
  tripTypes: 5,
  species: 6,
  waterFished: 7, // raw text, not used — we use waterTags instead
  season: 8,
  priceRange: 9,
  webPresence: 10,
  phone: 11,
  notes: 12,
  verified: 13,
  lastUpdated: 14,
  leadStatus: 15,
  approachPriority: 16,
  license: 17,
  waterTags: 18,
  description: 19,
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

function derivePriceTier(rawPrice) {
  const s = cleanValue(rawPrice);
  if (!s) return { priceTier: 'Contact for pricing', priceDisplay: 'Contact for pricing' };
  const match = s.replace(/,/g, '').match(/(\d+)/);
  if (!match) return { priceTier: 'Contact for pricing', priceDisplay: 'Contact for pricing' };
  const low = parseInt(match[1], 10);
  let priceTier;
  if (low < 500) priceTier = '$';
  else if (low <= 700) priceTier = '$$';
  else priceTier = '$$$';
  return { priceTier, priceDisplay: `${s}/day` };
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
const sheet = workbook.Sheets['Outfitters'];
const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false });

// Row 0 = title, Row 1 = header, data starts at Row 2
const dataRows = rows.slice(2);

fs.mkdirSync(outDir, { recursive: true });

let written = 0;
let skipped = 0;

for (const row of dataRows) {
  const name = cleanValue(row[COL.name]);
  if (!name) continue; // blank row

  const verified = cleanValue(row[COL.verified]) === 'Yes';
  const leadStatus = cleanValue(row[COL.leadStatus]) || 'Not yet contacted';

  if (!verified || leadStatus === 'Closed/Defunct') {
    skipped++;
    continue;
  }

  const description = cleanValue(row[COL.description]);
  if (!description) {
    console.warn(`⚠️  Skipping "${name}" — no Description in spreadsheet yet.`);
    skipped++;
    continue;
  }

  const { priceTier, priceDisplay } = derivePriceTier(row[COL.priceRange]);

  const frontmatter = {
    name,
    baseTown: cleanValue(row[COL.baseTown]) || '',
    region: cleanValue(row[COL.region]) || '',
    waters: splitList(row[COL.waterTags]),
    fishingType: splitList(row[COL.fishingType]),
    priceTier,
    priceDisplay,
    phone: cleanValue(row[COL.phone]),
    website: cleanValue(row[COL.website]),
    tripTypes: cleanValue(row[COL.tripTypes]),
    speciesTargeted: cleanValue(row[COL.species]),
    seasonActive: cleanValue(row[COL.season]),
    description,
    verified: true,
    leadStatus,
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

console.log(`✅ Wrote ${written} outfitter file(s), skipped ${skipped}.`);