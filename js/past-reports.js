// Past Reports tab — drill-down navigation:
// All Years → Academic Year → Term → Level → Subject → Reports

let prNav      = { year: null, term: null, level: null, subject: null };
let prExpanded = new Set();

// Encode a string safely for use inside a single-quoted JS string in an HTML attribute.
function prEsc(v){ return v === null ? 'null' : `'${v}'`; }

function renderPastReports(){
  const w       = document.getElementById('past-reports-wrap');
  const reports = getSavedReports();

  if(!reports.length){
    w.innerHTML = `<div class="empty"><div class="ei">📁</div><p>No saved reports yet — use the <strong>💾 Save Report</strong> button on the Export tab.</p></div>`;
    return;
  }

  let h = buildBreadcrumb() + '<div class="pr-content">';

  if(!prNav.year){
    h += buildYearGrid(reports);
  } else if(!prNav.term){
    h += buildTermGrid(reports);
  } else if(!prNav.level){
    h += buildLevelGrid(reports);
  } else if(!prNav.subject){
    h += buildSubjectGrid(reports);
  } else {
    h += buildReportList(reports);
  }

  h += '</div>';
  w.innerHTML = h;
}

function buildBreadcrumb(){
  if(!prNav.year) return '';
  const y = prNav.year, t = prNav.term, l = prNav.level;
  const crumbs = [
    { label: 'All Years', fn: `prSetNav(null,null,null,null)` },
    { label: y,           fn: t ? `prSetNav(${prEsc(y)},null,null,null)` : null }
  ];
  if(t) crumbs.push({ label: t, fn: l ? `prSetNav(${prEsc(y)},${prEsc(t)},null,null)` : null });
  if(l) crumbs.push({ label: l, fn: prNav.subject ? `prSetNav(${prEsc(y)},${prEsc(t)},${prEsc(l)},null)` : null });
  if(prNav.subject) crumbs.push({ label: prNav.subject, fn: null });

  return `<div class="pr-breadcrumb">
    ${crumbs.map(c =>
      c.fn
        ? `<button class="pr-crumb" onclick="${c.fn}">${c.label}</button>`
        : `<span class="pr-crumb-active">${c.label}</span>`
    ).join('<span class="pr-sep">›</span>')}
  </div>`;
}

function buildYearGrid(reports){
  const counts = {};
  reports.forEach(r => { counts[r.academicYear] = (counts[r.academicYear] || 0) + 1; });
  const years = Object.keys(counts).sort().reverse();
  return `<div class="pr-grid">
    ${years.map(y => `
      <div class="pr-card" onclick="prSetNav(${prEsc(y)},null,null,null)">
        <div class="pr-card-icon">📅</div>
        <div class="pr-card-title">${y}</div>
        <div class="pr-card-sub">${counts[y]} report${counts[y] !== 1 ? 's' : ''}</div>
      </div>`).join('')}
  </div>`;
}

function buildTermGrid(reports){
  const filtered = reports.filter(r => r.academicYear === prNav.year);
  const counts   = {};
  filtered.forEach(r => {
    const t = r.term || 'Unlabelled';
    counts[t] = (counts[t] || 0) + 1;
  });
  // Sort: Term 1 → Term 2 → Term 3 → Term 4 → anything else → Unlabelled last
  const termOrder = ['Term 1', 'Term 2', 'Term 3', 'Term 4'];
  const sorted = [
    ...termOrder.filter(t => counts[t]),
    ...Object.keys(counts).filter(t => !termOrder.includes(t) && t !== 'Unlabelled').sort(),
    ...(counts['Unlabelled'] ? ['Unlabelled'] : [])
  ];
  const icons = { 'Term 1': '1️⃣', 'Term 2': '2️⃣', 'Term 3': '3️⃣', 'Term 4': '4️⃣' };
  return `<div class="pr-grid">
    ${sorted.map(t => `
      <div class="pr-card" onclick="prSetNav(${prEsc(prNav.year)},${prEsc(t)},null,null)">
        <button class="pr-rename-btn" onclick="event.stopPropagation();prRenameTerm(${prEsc(prNav.year)},${prEsc(t)})" title="Rename term">✏️</button>
        <div class="pr-card-icon">${icons[t] || '📋'}</div>
        <div class="pr-card-title">${t}</div>
        <div class="pr-card-sub">${counts[t]} report${counts[t] !== 1 ? 's' : ''}</div>
      </div>`).join('')}
  </div>`;
}

function buildLevelGrid(reports){
  const filtered = reports.filter(r =>
    r.academicYear === prNav.year &&
    (r.term || 'Unlabelled') === prNav.term
  );
  const counts = {};
  filtered.forEach(r => { counts[r.yearLevel] = (counts[r.yearLevel] || 0) + 1; });
  const levels = YEAR_LEVELS.filter(l => counts[l]);
  return `<div class="pr-grid">
    ${levels.map(l => `
      <div class="pr-card" onclick="prSetNav(${prEsc(prNav.year)},${prEsc(prNav.term)},${prEsc(l)},null)">
        <div class="pr-card-icon">🎓</div>
        <div class="pr-card-title">${l}</div>
        <div class="pr-card-sub">${counts[l]} report${counts[l] !== 1 ? 's' : ''}</div>
      </div>`).join('')}
  </div>`;
}

function buildSubjectGrid(reports){
  const filtered = reports.filter(r =>
    r.academicYear === prNav.year &&
    (r.term || 'Unlabelled') === prNav.term &&
    r.yearLevel === prNav.level
  );
  const counts = {};
  filtered.forEach(r => { counts[r.subject] = (counts[r.subject] || 0) + 1; });
  const subs = Object.keys(counts).sort();
  return `<div class="pr-grid">
    ${subs.map(s => `
      <div class="pr-card" onclick="prSetNav(${prEsc(prNav.year)},${prEsc(prNav.term)},${prEsc(prNav.level)},${prEsc(s)})">
        <div class="pr-card-icon">📚</div>
        <div class="pr-card-title">${s}</div>
        <div class="pr-card-sub">${counts[s]} report${counts[s] !== 1 ? 's' : ''}</div>
      </div>`).join('')}
  </div>`;
}

function buildReportList(reports){
  const filtered = reports.filter(r =>
    r.academicYear === prNav.year &&
    (r.term || 'Unlabelled') === prNav.term &&
    r.yearLevel    === prNav.level &&
    r.subject      === prNav.subject
  );
  if(!filtered.length) return `<div class="empty"><p>No reports found.</p></div>`;

  return filtered.map(r => {
    const date    = new Date(r.savedAt);
    const dateStr = date.toLocaleDateString('en-GB', { day: 'numeric', month: 'short', year: 'numeric' });
    const timeStr = date.toLocaleTimeString('en-GB', { hour: '2-digit', minute: '2-digit' });
    const isOpen  = prExpanded.has(r.id);

    let commentsHtml = '';
    if(isOpen){
      const commentCards = r.students.map(s => {
        const dn      = displayName(s);
        const sel     = r.selections[s.id] || { mode: 'auto', s1: 0, s2cat: 's2_academic', s2: 0, s3: 0, s4: -1 };
        const comment = assembleFull(s, sel);
        const t       = tier(s.grade);
        const len     = comment.length;
        return `<div class="exp-card">
          <div class="exp-hdr">
            <span class="exp-name">${dn} <span class="badge b${t}">${s.grade}</span></span>
            <span class="cc ${len > 350 ? 'over' : len >= 330 ? 'good' : 'near'}" style="font-size:.7rem">${len} / 350</span>
          </div>
          <div class="exp-txt">${comment}</div>
        </div>`;
      }).join('');
      commentsHtml = `<div class="pr-comments">${commentCards}</div>`;
    }

    return `<div class="pr-report-card">
      <div class="pr-report-hdr">
        <div>
          <span class="pr-report-title">${r.subject} · ${r.yearLevel}</span>
          <span class="pr-report-meta">${dateStr} at ${timeStr} · ${r.studentCount} student${r.studentCount !== 1 ? 's' : ''}</span>
        </div>
        <div style="display:flex;gap:.4rem;flex-wrap:wrap">
          <button class="btn btn-ghost btn-sm" onclick="prToggle('${r.id}')">${isOpen ? '▲ Hide' : '▼ View'}</button>
          <button class="btn btn-primary btn-sm" onclick="loadSavedReportToWorkspace('${r.id}')">⬆ Load to Workspace</button>
          <button class="btn btn-red btn-sm" onclick="deleteSavedReport('${r.id}')">🗑 Delete</button>
        </div>
      </div>
      ${commentsHtml}
    </div>`;
  }).join('');
}

function prSetNav(year, term, level, subject){
  prNav = { year, term, level, subject };
  renderPastReports();
}

function prRenameTerm(year, term){
  const newName = prompt(`Rename "${term}" to:`, term);
  if(!newName || newName.trim() === '' || newName.trim() === term) return;
  renameTerm(year, term, newName.trim());
  // If we're currently navigated into this term, update the nav state
  if(prNav.term === term) prNav.term = newName.trim();
  renderPastReports();
}

function prToggle(id){
  if(prExpanded.has(id)) prExpanded.delete(id);
  else prExpanded.add(id);
  renderPastReports();
}
