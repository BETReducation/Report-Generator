// Export validation, copy, CSV download, and save-report modal.

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

function downloadCSV(){
  if(!checkGenders()) return;
  const rows = [['Student Name', 'Nickname', 'Gender', 'Grade', '%', 'Progress', 'Effort', 'Behaviour', 'Comment']];
  students.forEach(s => {
    const sel = selections[s.id] || { mode: 'auto', s1: 0, s2cat: 's2_academic', s2: 0, s3: 0, s4: -1 };
    rows.push([
      s.fullName, s.nickname,
      s.gender === 'M' ? 'Male' : s.gender === 'F' ? 'Female' : '',
      s.grade, s.percent, s.progress, s.effort, s.behaviour,
      assembleFull(s, sel)
    ]);
  });
  const csv = rows.map(r => r.map(c => `"${String(c).replace(/"/g, '""')}"`).join(',')).join('\n');
  const a = Object.assign(document.createElement('a'), {
    href:     URL.createObjectURL(new Blob([csv], { type: 'text/csv;charset=utf-8;' })),
    download: 'igcse_report_comments.csv'
  });
  a.click();
}

// ── Save Report Modal ─────────────────────────────────────────────────────────

function openSaveModal(){
  if(!students.length){ alert('No students to save.'); return; }

  // Populate academic year dropdown
  const ySelect = document.getElementById('m-year');
  const curr    = currentAcademicYear();
  ySelect.innerHTML = ACADEMIC_YEARS.map(y =>
    `<option value="${y}"${y === curr ? ' selected' : ''}>${y}</option>`
  ).join('');

  // Populate year level dropdown
  document.getElementById('m-level').innerHTML = YEAR_LEVELS.map(l =>
    `<option value="${l}">${l}</option>`
  ).join('');

  // Populate subject dropdown
  document.getElementById('m-subject').innerHTML = SUBJECTS.map(s =>
    `<option value="${s}">${s}</option>`
  ).join('');

  // Clear term field
  document.getElementById('m-term').value = '';

  document.getElementById('save-modal').style.display = 'flex';
  document.getElementById('m-term').focus();
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
  if(!meta.academicYear || !meta.yearLevel || !meta.subject){
    alert('Please fill in Academic Year, Year Level, and Subject.');
    return;
  }
  saveReport(meta);
  closeSaveModal();

  // Brief success flash on the save button
  const btn = document.getElementById('btn-save-report');
  if(btn){ const o = btn.textContent; btn.textContent = '✅ Saved!'; setTimeout(() => btn.textContent = o, 2000); }
}
