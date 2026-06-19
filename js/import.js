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

  // ── Enhance with JSZip: styles + data validations ─────────────────────────
  // Strategy: replace styles.xml wholesale (avoids fragile regex patching),
  // then inject row/cell styles and dataValidations into sheet1.xml.
  try {
    const zip = await JSZip.loadAsync(arr);

    // 1. Replace styles.xml with a complete, known-good version ───────────────
    // Styles: 0=default, 1=header (bold white on blue), 2=example (light-blue tint)
    zip.file('xl/styles.xml', [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
      '<fonts count="2">',
      '<font><sz val="12"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>',
      '<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/><family val="2"/></font>',
      '</fonts>',
      '<fills count="4">',
      '<fill><patternFill patternType="none"/></fill>',
      '<fill><patternFill patternType="gray125"/></fill>',
      '<fill><patternFill patternType="solid"><fgColor rgb="FF1D4ED8"/><bgColor indexed="64"/></patternFill></fill>',
      '<fill><patternFill patternType="solid"><fgColor rgb="FFEFF6FF"/><bgColor indexed="64"/></patternFill></fill>',
      '</fills>',
      '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>',
      '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>',
      '<cellXfs count="3">',
      '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>',
      '<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>',
      '<xf numFmtId="0" fontId="0" fillId="3" borderId="0" xfId="0" applyFill="1"/>',
      '</cellXfs>',
      '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>',
      '</styleSheet>',
    ].join('\n'));

    // 2. Patch sheet1.xml ────────────────────────────────────────────────────
    let s1 = await zip.file('xl/worksheets/sheet1.xml').async('string');

    // Header row: taller, style 1 (blue bg + white bold)
    s1 = s1.replace('<row r="1"', '<row r="1" ht="20" customHeight="1" s="1" customFormat="1"');
    ['A','B','C','D','E','F','G','H'].forEach(col => {
      s1 = s1.replace(new RegExp(`<c r="${col}1"`, 'g'), `<c r="${col}1" s="1"`);
    });

    // Example rows: style 2 (light-blue tint)
    s1 = s1.replace('<row r="2"', '<row r="2" s="2" customFormat="1"');
    s1 = s1.replace('<row r="3"', '<row r="3" s="2" customFormat="1"');
    ['A','B','C','D','E','F','G','H'].forEach(col => {
      s1 = s1.replace(new RegExp(`<c r="${col}2"`, 'g'), `<c r="${col}2" s="2"`);
      s1 = s1.replace(new RegExp(`<c r="${col}3"`, 'g'), `<c r="${col}3" s="2"`);
    });

    // Data validation dropdowns
    if (!s1.includes('<dataValidations')) {
      const pebList = '"Needs Improvement,Satisfactory,Good,Very Good"';
      s1 = s1.replace('</worksheet>',
        '<dataValidations count="5">' +
        dv('C2:C3000', '"M,F"') +
        dv('D2:D3000', '"A*,A,B,C,D,E,F,G"') +
        dv('F2:F3000', pebList) +
        dv('G2:G3000', pebList) +
        dv('H2:H3000', pebList) +
        '</dataValidations></worksheet>');
    }

    zip.file('xl/worksheets/sheet1.xml', s1);

    // 3. Download ──────────────────────────────────────────────────────────────
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
    console.warn('Enhanced template failed, falling back:', err);
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
