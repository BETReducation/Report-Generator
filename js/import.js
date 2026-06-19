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

  // ── Enhance with JSZip: styles + freeze pane + data validations ────────────
  try {
    const zip = await JSZip.loadAsync(arr);

    // 1. Patch styles.xml ──────────────────────────────────────────────────────
    let stylesXml = await zip.file('xl/styles.xml').async('string');

    // fontId=1: bold white — for header row
    stylesXml = stylesXml.replace(/<fonts count="(\d+)"/, (_, n) => `<fonts count="${+n + 1}"`);
    stylesXml = stylesXml.replace('</fonts>',
      '<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/><family val="2"/></font></fonts>');

    // fillId=2: solid blue (#1D4ED8) for header
    // fillId=3: light blue tint (#EFF6FF) for example rows
    stylesXml = stylesXml.replace(/<fills count="(\d+)"/, (_, n) => `<fills count="${+n + 2}"`);
    stylesXml = stylesXml.replace('</fills>',
      '<fill><patternFill patternType="solid"><fgColor rgb="FF1D4ED8"/><bgColor indexed="64"/></patternFill></fill>' +
      '<fill><patternFill patternType="solid"><fgColor rgb="FFEFF6FF"/><bgColor indexed="64"/></patternFill></fill>' +
      '</fills>');

    // xfId=1: header style (bold white on blue)
    // xfId=2: example row style (light blue tint, normal text)
    stylesXml = stylesXml.replace(/<cellXfs count="(\d+)"/, (_, n) => `<cellXfs count="${+n + 2}"`);
    stylesXml = stylesXml.replace('</cellXfs>',
      '<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1">' +
        '<alignment vertical="center"/></xf>' +
      '<xf numFmtId="0" fontId="0" fillId="3" borderId="0" xfId="0" applyFill="1">' +
        '<alignment vertical="center"/></xf>' +
      '</cellXfs>');

    zip.file('xl/styles.xml', stylesXml);

    // 2. Patch xl/worksheets/sheet1.xml ───────────────────────────────────────
    let s1 = await zip.file('xl/worksheets/sheet1.xml').async('string');

    // Freeze top row (replace entire sheetViews block to avoid duplicate pane elements)
    const frozenView =
      '<sheetViews><sheetView tabSelected="1" workbookViewId="0">' +
      '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>' +
      '<selection pane="bottomLeft" activeCell="A2" sqref="A2"/>' +
      '</sheetView></sheetViews>';
    if (s1.includes('<sheetViews>')) {
      s1 = s1.replace(/<sheetViews>[\s\S]*?<\/sheetViews>/, frozenView);
    } else {
      s1 = s1.replace('<sheetData>', frozenView + '<sheetData>');
    }

    // Style header row: taller + xfId=1 on row and each cell
    s1 = s1.replace('<row r="1"', '<row r="1" ht="20" customHeight="1" s="1" customFormat="1"');
    ['A','B','C','D','E','F','G','H'].forEach(col => {
      s1 = s1.replace(new RegExp(`<c r="${col}1"`, 'g'), `<c r="${col}1" s="1"`);
    });

    // Style example rows: xfId=2 on row and cells
    s1 = s1.replace('<row r="2"', '<row r="2" s="2" customFormat="1"');
    s1 = s1.replace('<row r="3"', '<row r="3" s="2" customFormat="1"');
    ['A','B','C','D','E','F','G','H'].forEach(col => {
      s1 = s1.replace(new RegExp(`<c r="${col}2"`, 'g'), `<c r="${col}2" s="2"`);
      s1 = s1.replace(new RegExp(`<c r="${col}3"`, 'g'), `<c r="${col}3" s="2"`);
    });

    // Inject data validation dropdowns
    const pebList = '"Needs Improvement,Satisfactory,Good,Very Good"';
    s1 = s1.replace('</worksheet>',
      '<dataValidations count="5">' +
      dv('C2:C3000', '"M,F"') +
      dv('D2:D3000', '"A*,A,B,C,D,E,F,G"') +
      dv('F2:F3000', pebList) +
      dv('G2:G3000', pebList) +
      dv('H2:H3000', pebList) +
      '</dataValidations></worksheet>');

    zip.file('xl/worksheets/sheet1.xml', s1);

    // 3. Download enhanced file ─────────────────────────────────────────────────
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
    console.warn('Enhanced template failed, falling back to plain version:', err);
    XLSX.writeFile(wb, 'BETR_Class_Template.xlsx');
  }
}

// Build a dataValidation XML element for a list dropdown
function dv(sqref, formula1){
  return `<dataValidation type="list" allowBlank="1" showDropDown="0" showErrorMessage="1" sqref="${sqref}">` +
         `<formula1>${formula1}</formula1></dataValidation>`;
}

// ── Spreadsheet upload ────────────────────────────────────────────────────────

function handleSheet(e){
  const file = e.target.files[0]; if(!file) return;
  const reader = new FileReader();
  reader.onload = ev => {
    try {
      const wb = XLSX.read(ev.target.result, { type: 'binary' });
      // Skip the Instructions sheet — use the first non-Instructions sheet
      const sheetName = wb.SheetNames.find(n => !/instructions/i.test(n)) || wb.SheetNames[0];
      const ws   = wb.Sheets[sheetName];
      const rows = XLSX.utils.sheet_to_json(ws, { defval: '' });

      if(students.length){
        const replace = confirm(
          `You have ${students.length} student${students.length !== 1 ? 's' : ''} in the current roster.\n\n` +
          `OK = Replace roster with this file\nCancel = Add to current roster`
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
        // Skip blank rows and [EXAMPLE] rows generated by the template
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

      const msg = `Imported ${added} student${added !== 1 ? 's' : ''} from "${sheetName}".`
        + (skipped > 0 ? ` (${skipped} rows skipped)` : '');
      alert(msg);
      saveState(); renderAll();
    } catch(err){
      alert('Could not read file.\n\nMake sure you are uploading the BETR template or a CSV/Excel file with a "Student Name" column.');
    }
  };
  reader.readAsBinaryString(file);
  e.target.value = '';
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
