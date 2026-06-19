// Export validation, copy, Excel/Word download, and save-report modal.

function checkGenders(single){
  const missing = students.filter(s => !s.gender);
  if(single){
    const s = students.find(x => x.id === single);
    if(s && !s.gender){ alert(`Please set gender (M/F) for ${s.fullName} before copying.`); return false; }
    return true;
  }
  if(missing.length){
    alert(`Please set gender (M/F) for the following students before exporting:\n\n${missing.map(s => '• ' + s.fullName).join('\n')}`);
    return false;
  }
  return true;
}

function copyOne(id){
  if(!checkGenders(id)) return;
  const s = students.find(x => x.id === id);
  const sel = selections[id] || { mode: 'auto', s1: 0, s2cat: 's2_academic', s2: 0, s3: 0, s4: -1 };
  navigator.clipboard.writeText(assembleFull(s, sel)).then(() => {
    document.querySelectorAll(`button[onclick="copyOne(${id})"]`).forEach(btn => {
      const o = btn.textContent; btn.textContent = '✅ Copied!';
      setTimeout(() => btn.textContent = o, 1500);
    });
  });
}

function copyAll(){
  if(!checkGenders()) return;
  const txt = students.map(s => {
    const sel = selections[s.id] || { mode: 'auto', s1: 0, s2cat: 's2_academic', s2: 0, s3: 0, s4: -1 };
    return `${displayName(s)} (${s.grade})\n${assembleFull(s, sel)}`;
  }).join('\n\n');
  navigator.clipboard.writeText(txt).then(() => alert('All comments copied!'));
}

// ── Download Excel ────────────────────────────────────────────────────────────

function _buildExportRows(){
  const rows = [['Student Name','Nickname','Gender','Grade','%','Progress','Effort','Behaviour','Comment']];
  students.forEach(s => {
    const sel = selections[s.id] || { mode: 'auto', s1: 0, s2cat: 's2_academic', s2: 0, s3: 0, s4: -1 };
    rows.push([
      s.fullName, s.nickname,
      s.gender === 'M' ? 'Male' : s.gender === 'F' ? 'Female' : '',
      s.grade, s.percent || '', s.progress, s.effort, s.behaviour,
      assembleFull(s, sel)
    ]);
  });
  return rows;
}

function downloadExcel(){
  if(!checkGenders()) return;
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(_buildExportRows());
  ws['!cols'] = [
    {wch:28},{wch:14},{wch:8},{wch:7},{wch:5},
    {wch:18},{wch:18},{wch:18},{wch:80}
  ];
  XLSX.utils.book_append_sheet(wb, ws, 'Report Comments');
  XLSX.writeFile(wb, 'Report Comments.xlsx');
}

// ── Download Word ─────────────────────────────────────────────────────────────

async function downloadWord(){
  if(!checkGenders()) return;
  if(typeof docx === 'undefined'){
    alert('Word export library is still loading — please try again in a moment.');
    return;
  }
  const { Document, Paragraph, TextRun, Packer, HeadingLevel, AlignmentType } = docx;

  const children = [];
  students.forEach((s, i) => {
    const sel = selections[s.id] || { mode: 'auto', s1: 0, s2cat: 's2_academic', s2: 0, s3: 0, s4: -1 };
    const dn      = displayName(s);
    const comment = assembleFull(s, sel);
    if(i > 0) children.push(new Paragraph({ text: '' }));
    children.push(new Paragraph({
      children: [
        new TextRun({ text: dn, bold: true }),
        new TextRun({ text: `  (${s.grade}${s.percent ? ' · ' + s.percent + '%' : ''})`, color: '888888' })
      ]
    }));
    children.push(new Paragraph({ children: [new TextRun({ text: comment })] }));
  });

  const doc = new Document({ sections: [{ children }] });
  const blob = await Packer.toBlob(doc);
  const a = Object.assign(document.createElement('a'), {
    href: URL.createObjectURL(blob),
    download: 'Report Comments.docx'
  });
  a.click();
  setTimeout(() => URL.revokeObjectURL(a.href), 10000);
}

// ── Cloud save helpers ────────────────────────────────────────────────────────

function openCloudModal(){
  if(!students.length){ alert('No students to export.'); return; }
  document.getElementById('cloud-modal').style.display = 'flex';
}

function closeCloudModal(){
  document.getElementById('cloud-modal').style.display = 'none';
}

async function cloudDownloadAndOpen(format, service){
  if(format === 'word') await downloadWord();
  else downloadExcel();

  const urls = {
    onedrive: 'https://onedrive.live.com',
    gdrive:   'https://drive.google.com/drive/my-drive'
  };
  setTimeout(() => window.open(urls[service], '_blank'), 800);
}

// ── Save Report Modal ─────────────────────────────────────────────────────────

function openSaveModal(){
  if(!students.length){ alert('No students to save.'); return; }

  const ySelect = document.getElementById('m-year');
  const curr    = currentAcademicYear();
  ySelect.innerHTML = ACADEMIC_YEARS.map(y =>
    `<option value="${y}"${y === curr ? ' selected' : ''}>${y}</option>`
  ).join('');

  document.getElementById('m-level').innerHTML = YEAR_LEVELS.map(l =>
    `<option value="${l}">${l}</option>`
  ).join('');

  document.getElementById('m-subject').innerHTML = SUBJECTS.map(s =>
    `<option value="${s}">${s}</option>`
  ).join('');

  document.getElementById('m-term').value = 'Term 1';
  document.getElementById('save-modal').style.display = 'flex';
  document.getElementById('m-year').focus();
}

function closeSaveModal(){
  document.getElementById('save-modal').style.display = 'none';
}

function confirmSave(){
  const meta = {
    academicYear: document.getElementById('m-year').value,
    yearLevel:    document.getElementById('m-level').value,
    subject:      document.getElementById('m-subject').value,
    term:         document.getElementById('m-term').value.trim()
  };
  if(!meta.academicYear || !meta.term || !meta.yearLevel || !meta.subject){
    alert('Please fill in all fields.');
    return;
  }
  saveReport(meta);
  closeSaveModal();

  const btn = document.getElementById('btn-save-report');
  if(btn){ const o = btn.textContent; btn.textContent = '✅ Saved!'; setTimeout(() => btn.textContent = o, 2000); }
}
