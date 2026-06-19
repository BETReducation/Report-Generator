// Spreadsheet import and Excel template generation.
// Data validation + cell styling is injected via JSZip because SheetJS
// community edition does not support these features natively.

// ── Template download ─────────────────────────────────────────────────────────

async function downloadTemplate(){
  const wb = XLSX.utils.book_new();

  // ── Student Data sheet ─────────────────────────────────────────────────────
  const headers = ['Student Name', 'Nickname', 'Gender', 'Grade', '%', 'Progress', 'Effort', 'Behaviour'];
  // Example rows are prefixed with [EXAMPLE] so handleSheet() will skip them automatically
  const ex1 = ['[EXAMPLE] Le Thanh Hung (Nancy)', 'Nancy', 'F', 'A', '87', 'Very Good', 'Good', 'Very Good'];
  const ex2 = ['[EXAMPLE] Araki Shoei', '', 'M', 'B', '72', 'Good', 'Satisfactory', 'Good'];

  const ws = XLSX.utils.aoa_to_sheet([headers, ex1, ex2]);
  ws['!cols'] = [
    { wch: 36 }, { wch: 14 }, { wch: 9 }, { wch: 8 },
    { wch: 7  }, { wch: 20 }, { wch: 20 }, { wch: 20 },
  ];
  XLSX.utils.book_append_sheet(wb, ws, 'Student Data');

  // ── Instructions sheet ─────────────────────────────────────────────────────
  const instrRows = [
    ['BETR Report Generator — Class Data Template'],
    [''],
    ['HOW TO USE THIS TEMPLATE'],
    ['1. Delete the two [EXAMPLE] rows in the "Student Data" tab.'],
    ['2. Fill in your class — use the dropdown menus for Gender, Grade, Progress, Effort & Behaviour.'],
    ['3. Excel: File → Save As → Excel Workbook (.xlsx)'],
    ['   Google Sheets: File → Download → Microsoft Excel (.xlsx)  or  → CSV'],
    ['4. Upload the completed file to the Report Generator website.'],
    [''],
    ['COLUMN GUIDE'],
    ['Column', 'Description', 'Valid values'],
    ['Student Name', 'Full legal name. Put nickname in brackets: Le Thanh Hung (Nancy)', 'Any text'],
    ['Nickname', 'Optional — leave blank for automatic short-name rules.', 'Any text'],
    ['Gender', 'Required to generate comments.', 'M  or  F'],
    ['Grade', 'Achieved grade for this reporting period.', 'A*, A, B, C, D, E, F, G'],
    ['%', 'Percentage score — no % symbol needed.', '0 – 100'],
    ['Progress', 'Progress rating for the term.', 'Needs Improvement / Satisfactory / Good / Very Good'],
    ['Effort', 'Effort rating for the term.', 'Needs Improvement / Satisfactory / Good / Very Good'],
    ['Behaviour', 'Behaviour rating for the term.', 'Needs Improvement / Satisfactory / Good / Very Good'],
    [''],
    ['TIPS'],
    ['• Gender, Grade, Progress, Effort and Behaviour columns have dropdown menus in Excel.'],
    ['• Names with nicknames in brackets — Le Thanh Hung (Nancy) — are parsed automatically.'],
    ['• P/E/B are optional but drive auto sentence selection on the website.'],
    ['• Text matching is case-insensitive — "very good" and "Very Good" both work.'],
  ];
  const wsI = XLSX.utils.aoa_to_sheet(instrRows);
  wsI['!cols'] = [{ wch: 18 }, { wch: 60 }, { wch: 42 }];
  XLSX.utils.book_append_sheet(wb, wsI, 'Instructions');

  // ── Write base workbook ────────────────────────────────────────────────────
  const arr = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });

  // ── Inject data validations via JSZip ─────────────────────────────────────
  try {
    const zip = await JSZip.loadAsync(arr);

    // Read sheet1.xml as binary string (JSZip uses Latin-1; for ASCII-only
    // modifications this is identical to UTF-8 so no encoding issues arise)
    let s1 = await zip.file('xl/worksheets/sheet1.xml').async('string');

    // Append dataValidations immediately before </worksheet>
    if (!s1.includes('<dataValidations')) {
      const pebList = '"Needs Improvement,Satisfactory,Good,Very Good"';
      const dvBlock =
        '<dataValidations count="5">' +
        dv('C2:C3000', '"M,F"') +
        dv('D2:D3000', '"A*,A,B,C,D,E,F,G"') +
        dv('F2:F3000', pebList) +
        dv('G2:G3000', pebList) +
        dv('H2:H3000', pebList) +
        '</dataValidations>';
      // Use lastIndexOf so we replace only the final </worksheet> tag
      const idx = s1.lastIndexOf('</worksheet>');
      if (idx !== -1) s1 = s1.slice(0, idx) + dvBlock + '</worksheet>';
    }

    zip.file('xl/worksheets/sheet1.xml', s1);

    const blob = await zip.generateAsync({
      type: 'blob',
      mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      compression: 'DEFLATE',
    });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'BETR_Class_Template.xlsx';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 10000);

  } catch (err) {
    console.error('[Template]', err);
    // Fallback: plain xlsx without dropdowns
    XLSX.writeFile(wb, 'BETR_Class_Template.xlsx');
  }
}

// Build a dataValidation XML element for a list dropdown
function dv(sqref, formula1){
  return `<dataValidation type="list" allowBlank="1" showDropDown="0" showErrorMessage="1" sqref="${sqref}">` +
         `<formula1>${formula1}</formula1></dataValidation>`;
}

// ── Google Sheets import ──────────────────────────────────────────────────────

async function importFromGoogleSheets(){
  const input = document.getElementById('gs-url');
  const url   = (input ? input.value : '').trim();
  if(!url){ alert('Please paste a Google Sheets link first.'); return; }

  // Extract the spreadsheet ID from any Google Sheets URL format
  const match = url.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  if(!match){ alert('That doesn\'t look like a Google Sheets link.\n\nPaste the full URL from your browser address bar, e.g.\nhttps://docs.google.com/spreadsheets/d/abc123/edit'); return; }

  const id     = match[1];
  const csvUrl = `https://docs.google.com/spreadsheets/d/${id}/export?format=csv`;

  const btn = document.getElementById('gs-btn');
  if(btn){ btn.textContent = 'Importing…'; btn.disabled = true; }

  try {
    const res = await fetch(csvUrl);
    if(!res.ok) throw new Error(`HTTP ${res.status}`);
    const text = await res.text();
    // Parse CSV via SheetJS then hand off to the shared import logic
    const wb   = XLSX.read(text, { type: 'string' });
    const fake = { target: { _wb: wb, files: null } };
    _importWorkbook(wb, 'Google Sheets');
    if(input) input.value = '';
  } catch(err){
    alert('Could not fetch the sheet.\n\nMake sure:\n• The sheet is shared as "Anyone with the link can view"\n• You copied the full URL from the browser address bar');
  } finally {
    if(btn){ btn.textContent = 'Import'; btn.disabled = false; }
  }
}

// ── Spreadsheet upload ────────────────────────────────────────────────────────

function handleSheet(e){
  const file = e.target.files[0]; if(!file) return;
  const reader = new FileReader();
  reader.onload = ev => {
    try {
      _importWorkbook(XLSX.read(ev.target.result, { type: 'binary' }), file.name);
    } catch(err){
      alert('Could not read file.\n\nMake sure you are uploading the BETR template or a CSV/Excel file with a "Student Name" column.');
    }
  };
  reader.readAsBinaryString(file);
  e.target.value = '';
}

function _importWorkbook(wb, sourceName){
  const sheetName = wb.SheetNames.find(n => !/instructions/i.test(n)) || wb.SheetNames[0];
  const rows = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { defval: '' });

  if(students.length){
    const replace = confirm(
      `You have ${students.length} student${students.length !== 1 ? 's' : ''} in the current roster.\n\n` +
      `OK = Replace roster\nCancel = Add to current roster`
    );
    if(replace){ students = []; selections = {}; }
  }

  let added = 0, skipped = 0;
  rows.forEach(row => {
    const get = (...names) => {
      for(const n of names){
        const k = Object.keys(row).find(k => k.toLowerCase().replace(/[\s_%]/g,'').includes(n));
        if(k !== undefined && row[k] !== '') return String(row[k]).trim();
      }
      return '';
    };
    const rawName = get('studentname','name','student');
    if(!rawName || /^\[?example/i.test(rawName)){ skipped++; return; }
    const parsed = parseBracketed(rawName);
    const nick   = get('nickname','englishname','nick') || parsed.nickname;
    const grade  = get('grade','symbol','mark');
    const pct    = get('percent','percentage','score');
    const prog   = normalise(get('progress'));
    const eff    = normalise(get('effort'));
    const beh    = normalise(get('behaviour','behavior'));
    const rawG   = get('gender','sex');
    const gender = rawG ? (/^f/i.test(rawG) ? 'F' : 'M') : '';
    if(parsed.fullName){ pushStudent(parsed.fullName, nick, grade||'C', pct, prog, eff, beh, gender); added++; }
    else skipped++;
  });

  const msg = `Imported ${added} student${added !== 1 ? 's' : ''} from "${sourceName || sheetName}".`
    + (skipped > 0 ? ` (${skipped} rows skipped)` : '');
  alert(msg);
  saveState(); renderAll();
}

// Normalise P/E/B text — case-insensitive, forgiving of minor typos
function normalise(val){
  const v = (val || '').trim().toLowerCase();
  if(!v) return '';
  if(v.includes('very') && v.includes('good'))   return 'Very Good';
  if(v.includes('good'))                         return 'Good';
  if(v.includes('satisf'))                       return 'Satisfactory';
  if(v.includes('need') || v.includes('improv')) return 'Needs Improvement';
  return '';
}
