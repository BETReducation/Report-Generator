// Spreadsheet (Excel/CSV) and image OCR import.

const REPORT_NOISE = /\b(satisfactory|good|very|excellent|outstanding|poor|needs|improvement|student|code|name|progress|effort|behaviour|behavior|percent|grade|symbol|comment|subject|teacher|stu\w*)\b/gi;

function cleanOCRName(raw){
  return raw
    .replace(/STU\w+/gi, '')
    .replace(new RegExp(REPORT_NOISE.source, 'gi'), '')
    .replace(/\d+(\.\d+)?%?/g, '')
    .replace(/[^a-zA-ZÀ-ÿ\s()\-']/g, '')
    .replace(/\s+/g, ' ').trim();
}

// Extract up to 3 Progress/Effort/Behaviour values from an OCR line segment.
// Each pass picks whichever value starts EARLIEST in the remaining text,
// so "Satisfactory Good Very Good" correctly yields all three in order.
function extractPEB(text){
  const vals = ['Very Good', 'Needs Improvement', 'Satisfactory', 'Good'];
  const found = [];
  let t = text;
  for(let i = 0; i < 3; i++){
    let bestIdx = Infinity, bestVal = null;
    for(const v of vals){
      const idx = t.toLowerCase().indexOf(v.toLowerCase());
      if(idx !== -1 && idx < bestIdx){ bestIdx = idx; bestVal = v; }
    }
    if(!bestVal) break;
    found.push(bestVal);
    t = t.substring(bestIdx + bestVal.length);
  }
  return { progress: found[0] || '', effort: found[1] || '', behaviour: found[2] || '' };
}

// Returns true if a line contains anchors that mark it as a data row
// (grade letter, percentage, STU code, or P/E/B values).
function lineHasAnchors(line){
  return /\b(A\+|A\*|A|B|C)\b/.test(line)
      || /\b\d{2,3}(?:\.\d+)?\b/.test(line)
      || /\bSTU\w+/i.test(line)
      || /(Very Good|Needs Improvement|Satisfactory|Good)/i.test(line);
}

// When a name is too long to fit on one line in the screenshot, the scanner
// produces a name-only first line followed by a data line with no name.
// Merge such orphan lines with the next line when the next line has a grade.
function mergeWrappedLines(lines){
  const out = [];
  let i = 0;
  while(i < lines.length){
    const curr = lines[i];
    const currIsOrphan = !lineHasAnchors(curr) && curr.trim().length > 1;
    if(currIsOrphan && i + 1 < lines.length){
      const next = lines[i + 1];
      if(/\b(A\+|A\*|A|B|C)\b/.test(next)){
        out.push(curr.trim() + ' ' + next.trim());
        i += 2;
        continue;
      }
    }
    out.push(curr);
    i++;
  }
  return out;
}

function handleSheet(e){
  const file = e.target.files[0]; if(!file) return;
  const reader = new FileReader();
  reader.onload = ev => {
    try {
      const wb   = XLSX.read(ev.target.result, { type: 'binary' });
      const ws   = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws, { defval: '' });
      if(students.length){
        const replace = confirm(`You have ${students.length} student${students.length !== 1 ? 's' : ''} in the current roster.\n\nOK = Replace roster with this file\nCancel = Add to current roster`);
        if(replace){ students = []; selections = {}; }
      }
      let added = 0;
      rows.forEach(row => {
        const get = (...names) => {
          for(const n of names){
            const k = Object.keys(row).find(k => k.toLowerCase().replace(/[\s_]/g, '').includes(n));
            if(k && row[k]) return String(row[k]).trim();
          }
          return '';
        };
        const rawName = get('name', 'student'); if(!rawName) return;
        const parsed  = parseBracketed(rawName);
        const nick    = get('nickname', 'englishname', 'nick') || parsed.nickname;
        const grade   = get('grade', 'symbol', 'mark');
        const pct     = get('%', 'percent', 'percentage', 'score');
        const prog    = get('progress');
        const eff     = get('effort');
        const beh     = get('behaviour', 'behavior');
        const rawG    = get('gender', 'sex', 'm/f');
        const gender  = rawG ? (/^f/i.test(rawG) ? 'F' : 'M') : '';
        if(parsed.fullName){ pushStudent(parsed.fullName, nick, grade || 'C', pct, prog, eff, beh, gender); added++; }
      });
      alert(`Imported ${added} student${added !== 1 ? 's' : ''} from spreadsheet.`);
    } catch(err){
      alert('Could not read file. Ensure it has a "Name" or "Student Name" column.');
    }
  };
  reader.readAsBinaryString(file);
  e.target.value = '';
}

async function handleImg(e){
  const file   = e.target.files[0]; if(!file) return;
  const status = document.getElementById('ocr-status');
  status.style.display = 'block';
  status.innerHTML     = '⏳ Loading OCR engine...';
  try {
    if(typeof Tesseract === 'undefined'){
      await new Promise((res, rej) => {
        const sc  = document.createElement('script');
        sc.src    = 'https://cdn.jsdelivr.net/npm/tesseract.js@5/dist/tesseract.min.js';
        sc.onload = res; sc.onerror = rej;
        document.head.appendChild(sc);
      });
    }
    const { data: { text } } = await Tesseract.recognize(file, 'eng', {
      logger: m => { if(m.status === 'recognizing text') status.innerHTML = `⏳ OCR: ${Math.round(m.progress * 100)}%`; }
    });

    status.innerHTML = '✅ Done. Parsing rows...';

    if(students.length){
      const replace = confirm(`You have ${students.length} student${students.length !== 1 ? 's' : ''} in the current roster.\n\nOK = Replace roster with this image\nCancel = Add to current roster`);
      if(replace){ students = []; selections = {}; }
    }

    // Merge lines where a name has wrapped onto the next line
    const rawLines   = text.split('\n').map(l => l.trim()).filter(Boolean);
    const lines      = mergeWrappedLines(rawLines);

    let added = 0, skippedLines = [];

    lines.forEach(line => {
      const gm = line.match(/\b(A\+|A\*|A|B|C)\b/);
      const pm = line.match(/\b(\d{2,3}(?:\.\d+)?)\b/);
      if(!gm){ skippedLines.push(line); return; }
      const gradeIdx = line.search(/\b(A\+|A\*|A|B|C)\b/);
      const pctIdx   = pm ? line.indexOf(pm[0]) : gradeIdx;
      const stuIdx   = line.search(/\bSTU\w+/i);
      const anchors  = [gradeIdx, pctIdx, stuIdx].filter(i => i > 0);
      const cutAt    = anchors.length ? Math.min(...anchors) : -1;
      let namePart   = cutAt > 0 ? line.substring(0, cutAt) : line.replace(/\b(A\+|A\*|A|B|C)\b.*/, '');
      namePart = cleanOCRName(namePart);
      if(namePart.length < 4 || !namePart.includes(' ')){ skippedLines.push(line); return; }
      const p = parseBracketed(namePart);
      if(p.fullName.split(/\s+/).length < 2){ skippedLines.push(line); return; }
      const remainder = cutAt > 0 ? line.substring(cutAt) : line;
      const peb = extractPEB(remainder);
      pushStudent(p.fullName, p.nickname, gm[0].toUpperCase(), pm ? pm[1] : '', peb.progress, peb.effort, peb.behaviour, '');
      added++;
    });

    // Build result message + optional raw-text review panel
    const skCount = skippedLines.length;
    let html = added
      ? `✅ Imported <strong>${added}</strong> student${added !== 1 ? 's' : ''}${skCount ? ` · <strong>${skCount}</strong> lines not recognised` : ''}. Review roster below.`
      : `⚠️ Could not detect any rows. Check the raw OCR text below and add students manually.`;

    // Always show the raw OCR panel so teacher can spot missed names
    const escapedRaw = text.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
    html += `
      <div style="margin-top:.6rem">
        <button class="btn btn-ghost btn-sm" onclick="toggleOcrRaw(this)" style="margin-bottom:.4rem">
          🔍 Review raw OCR text (${rawLines.length} lines)
        </button>
        <div id="ocr-raw" style="display:none">
          <p style="font-size:.72rem;color:var(--muted);margin-bottom:.35rem">
            Names split across two lines are merged automatically. If any are still missing,
            copy the name from below and add it manually using the form above.
          </p>
          <textarea readonly style="width:100%;height:180px;font-size:.72rem;font-family:monospace;border:1px solid var(--border);border-radius:6px;padding:.5rem;resize:vertical;background:#f8fafc">${escapedRaw}</textarea>
        </div>
      </div>`;

    status.innerHTML = html;

  } catch(err){
    status.innerHTML = '❌ OCR failed. Please add students manually.';
  }
  e.target.value = '';
}

function toggleOcrRaw(btn){
  const el = document.getElementById('ocr-raw');
  const open = el.style.display === 'none';
  el.style.display = open ? 'block' : 'none';
  btn.textContent = open
    ? '▲ Hide raw OCR text'
    : `🔍 Review raw OCR text`;
}
