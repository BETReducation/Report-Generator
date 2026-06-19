// Spreadsheet import and Excel template generation.

const PEB_VALUES   = ['Needs Improvement', 'Satisfactory', 'Good', 'Very Good'];
const GRADE_VALUES = ['A*', 'A', 'B', 'C', 'D', 'E', 'F', 'G'];

// ── Template download ─────────────────────────────────────────────────────────

function downloadTemplate(){
  const wb = XLSX.utils.book_new();

  // ── Student Data sheet ──────────────────────────────────────────────────────
  const headers = ['Student Name', 'Nickname', 'Gender', 'Grade', '%', 'Progress', 'Effort', 'Behaviour'];
  const examples = [
    ['Le Thanh Hung (Nancy)', 'Nancy', 'F', 'A', '87', 'Very Good', 'Good', 'Very Good'],
    ['Araki Shoei', '', 'M', 'B', '72', 'Good', 'Satisfactory', 'Good'],
  ];
  const ws = XLSX.utils.aoa_to_sheet([headers, ...examples]);

  ws['!cols'] = [
    { wch: 30 }, // Student Name
    { wch: 15 }, // Nickname
    { wch: 10 }, // Gender
    { wch: 8  }, // Grade
    { wch: 6  }, // %
    { wch: 22 }, // Progress
    { wch: 22 }, // Effort
    { wch: 22 }, // Behaviour
  ];

  // Data validation dropdowns
  const pebFormula   = `"${PEB_VALUES.join(',')}"`;
  const gradeFormula = `"${GRADE_VALUES.join(',')}"`;
  ws['!dataValidations'] = [
    { sqref: 'C2:C2000', type: 'list', formula1: '"M,F"',    allowBlank: true, showDropDown: false },
    { sqref: 'D2:D2000', type: 'list', formula1: gradeFormula, allowBlank: true, showDropDown: false },
    { sqref: 'F2:F2000', type: 'list', formula1: pebFormula,  allowBlank: true, showDropDown: false },
    { sqref: 'G2:G2000', type: 'list', formula1: pebFormula,  allowBlank: true, showDropDown: false },
    { sqref: 'H2:H2000', type: 'list', formula1: pebFormula,  allowBlank: true, showDropDown: false },
  ];

  XLSX.utils.book_append_sheet(wb, ws, 'Student Data');

  // ── Instructions sheet ──────────────────────────────────────────────────────
  const instrRows = [
    ['BETR Report Generator — Class Data Template'],
    [''],
    ['HOW TO USE THIS TEMPLATE'],
    ['1. Fill in your class data in the "Student Data" tab — delete the two example rows first.'],
    ['2. Save the file as .xlsx  (File → Save As → Excel Workbook).'],
    ['   Google Sheets users: File → Download → Microsoft Excel (.xlsx)  or  → CSV (.csv)'],
    ['3. Upload the completed file to the Report Generator website.'],
    [''],
    ['COLUMN GUIDE'],
    ['Column', 'Description', 'Valid values'],
    ['Student Name', 'Full legal name. Put nickname in brackets: Le Thanh Hung (Nancy)', 'Any text'],
    ['Nickname', 'Optional — leave blank to use automatic name rules.', 'Any text'],
    ['Gender', 'Required to generate comments.', 'M  or  F'],
    ['Grade', 'Achieved grade for this reporting period.', 'A*, A, B, C, D, E, F, G'],
    ['%', 'Percentage score. No % symbol needed.', '0 – 100'],
    ['Progress', 'Progress rating for the term.', 'Needs Improvement / Satisfactory / Good / Very Good'],
    ['Effort', 'Effort rating for the term.', 'Needs Improvement / Satisfactory / Good / Very Good'],
    ['Behaviour', 'Behaviour rating for the term.', 'Needs Improvement / Satisfactory / Good / Very Good'],
    [''],
    ['TIPS'],
    ['• The Gender, Grade, Progress, Effort and Behaviour columns have dropdown menus in Excel.'],
    ['• Names with nicknames in brackets — Le Thanh Hung (Nancy) — are parsed automatically.'],
    ['• Progress, Effort and Behaviour are optional but are used to auto-select comment sentences.'],
    ['• All text matching is case-insensitive, so "very good" and "Very Good" both work.'],
  ];
  const wsInstr = XLSX.utils.aoa_to_sheet(instrRows);
  wsInstr['!cols'] = [{ wch: 18 }, { wch: 58 }, { wch: 45 }];
  XLSX.utils.book_append_sheet(wb, wsInstr, 'Instructions');

  XLSX.writeFile(wb, 'BETR_Class_Template.xlsx');
}

// ── Spreadsheet upload ────────────────────────────────────────────────────────

function handleSheet(e){
  const file = e.target.files[0]; if(!file) return;
  const reader = new FileReader();
  reader.onload = ev => {
    try {
      const wb   = XLSX.read(ev.target.result, { type: 'binary' });
      // Skip the Instructions sheet — always use the first non-Instructions sheet
      const sheetName = wb.SheetNames.find(n => !/instructions/i.test(n)) || wb.SheetNames[0];
      const ws   = wb.Sheets[sheetName];
      const rows = XLSX.utils.sheet_to_json(ws, { defval: '' });

      if(students.length){
        const replace = confirm(`You have ${students.length} student${students.length !== 1 ? 's' : ''} in the current roster.\n\nOK = Replace roster with this file\nCancel = Add to current roster`);
        if(replace){ students = []; selections = {}; }
      }

      let added = 0, skipped = 0;
      rows.forEach(row => {
        const get = (...names) => {
          for(const n of names){
            const k = Object.keys(row).find(k => k.toLowerCase().replace(/[\s_%]/g, '').includes(n));
            if(k !== undefined && row[k] !== '') return String(row[k]).trim();
          }
          return '';
        };
        const rawName = get('studentname', 'name', 'student');
        if(!rawName || /^example/i.test(rawName)){ skipped++; return; } // skip example rows
        const parsed = parseBracketed(rawName);
        const nick   = get('nickname', 'englishname', 'nick') || parsed.nickname;
        const grade  = get('grade', 'symbol', 'mark');
        const pct    = get('percent', 'percentage', 'score');
        const prog   = normalise(get('progress'));
        const eff    = normalise(get('effort'));
        const beh    = normalise(get('behaviour', 'behavior'));
        const rawG   = get('gender', 'sex');
        const gender = rawG ? (/^f/i.test(rawG) ? 'F' : 'M') : '';
        if(parsed.fullName){ pushStudent(parsed.fullName, nick, grade || 'C', pct, prog, eff, beh, gender); added++; }
        else skipped++;
      });

      const msg = `Imported ${added} student${added !== 1 ? 's' : ''} from "${sheetName}".`
        + (skipped > 0 ? ` (${skipped} rows skipped)` : '');
      alert(msg);
    } catch(err){
      alert('Could not read file.\n\nMake sure you are uploading the BETR template or a file with a "Student Name" column.');
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
