document.addEventListener('DOMContentLoaded', () => {
  'use strict';

  const fileInput = document.getElementById('fileInput');
  const dropZone = document.getElementById('dropZone');
  const loadBtn = document.getElementById('loadBtn');
  const fileDisplayName = document.getElementById('fileDisplayName');
  const fileStatus = document.getElementById('fileStatus');
  const progressRow = document.getElementById('progressRow');
  const progressBar = document.getElementById('progressBar');
  const filterEl = document.getElementById('filter');
  const resultsEl = document.getElementById('results');
  const moreBtn = document.getElementById('moreBtn');

  const dim = document.getElementById('dim');
  const overlay = document.getElementById('overlay');
  const closeModal = document.getElementById('closeModal');
  const modalTitle = document.getElementById('modalTitle');
  const modalBody = document.getElementById('modalBody');

  let selectedFile = null;
  let items = [];
  let rendered = 0;
  const PER_PAGE = 50;

  const INCLUDE_ALL_BY_DEFAULT = true;
  const BLOCKED_SUBJECT_RE = /\b(retirada|reagend|tentativa|^re$)\b/i;

  const log = (...a) => console.log('[Athon]', ...a);

  function setStatus(txt, type = 'info') {
    if (fileStatus) fileStatus.innerHTML = `<span style="color:${type === 'error' ? 'var(--danger)' : 'var(--muted)'}">${txt}</span>`;
    log(txt);
  }

  function showProgress(p = 0) {
    if (!progressRow || !progressBar) return;
    progressRow.classList.toggle('hidden', p <= 0.01 || p >= 0.99);
    progressBar.style.width = Math.round(p * 100) + '%';
  }

  function updateFileDisplay(file) {
    const t = file ? `${file.name} — ${Math.round(file.size / 1024)} KB` : 'Clique ou arraste um arquivo CSV/Excel';
    if (fileDisplayName) fileDisplayName.textContent = t;
  }

  function escapeHtml(s) {
    return s == null ? '' : String(s).replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;').replaceAll('"', '&quot;');
  }

  function syncInputFiles(f) {
    if (!fileInput || !f) return;
    try {
      const dt = new DataTransfer();
      dt.items.add(f);
      fileInput.files = dt.files;
    } catch (e) {
      log('DataTransfer failed (ok fallback)', e);
    }
  }

  function showMessage(title, message, isError = false) {
    if (!overlay || !dim || !modalBody || !modalTitle) {
      console.error(`${title}: ${message}`);
      return;
    }
    modalTitle.textContent = title;
    modalTitle.style.color = isError ? 'var(--danger)' : 'var(--text)';
    modalBody.innerHTML = `<p>${escapeHtml(message)}</p>`;
    overlay.classList.remove('hidden');
    dim.classList.remove('hidden');
    closeModal.onclick = closeMessage;
    dim.onclick = closeMessage;
  }

  function closeMessage() {
    if (overlay) overlay.classList.add('hidden');
    if (dim) dim.classList.add('hidden');
    dim.onclick = null; 
  }

  async function readAsTextDetectEncoding(file) {
    const ab = await file.arrayBuffer();
    const bytes = new Uint8Array(ab);
    if (bytes[0] === 0xFF && bytes[1] === 0xFE) return new TextDecoder('utf-16le').decode(ab);
    if (bytes[0] === 0xFE && bytes[1] === 0xFF) return new TextDecoder('utf-16be').decode(ab);
    if (bytes[0] === 0xEF && bytes[1] === 0xBB && bytes[2] === 0xBF) return new TextDecoder('utf-8').decode(ab);
    try { return new TextDecoder('utf-8', { fatal: true }).decode(ab); } catch (e) { return new TextDecoder('iso-8859-1').decode(ab); }
  }

  function detectDelimiter(sample) {
    const lines = sample.split(/\r\n|\n/).slice(0, 8);
    const counts = { ',': 0, ';': 0, '\t': 0 };
    lines.forEach(l => { counts[','] += (l.match(/,/g) || []).length; counts[';'] += (l.match(/;/g) || []).length; counts['\t'] += (l.match(/\t/g) || []).length; });
    const order = [',', ';', '\t'];
    order.sort((a, b) => counts[b] - counts[a] || order.indexOf(a) - order.indexOf(b));
    return order[0];
  }

  function subjectAllowed(assuntoRaw) {
    if (!assuntoRaw) return false;
    if (BLOCKED_SUBJECT_RE.test(assuntoRaw)) return false;
    if (INCLUDE_ALL_BY_DEFAULT) return true;
    return true;
  }

  function parseCSVText(csvText, delimiter = ',', onProgress = () => { }) {
    const lines = csvText.replace(/\r\n/g, '\n').split('\n').filter(l => l.trim().length > 0);
    if (lines.length === 0) return [];

    function parseLine(line) {
      const arr = []; let cur = ''; let inQuotes = false;
      for (let i = 0; i < line.length; i++) {
        const ch = line[i];
        if (ch === '"') {
          if (inQuotes && line[i + 1] === '"') { cur += '"'; i++; continue; }
          inQuotes = !inQuotes; continue;
        }
        if (ch === delimiter && !inQuotes) { arr.push(cur); cur = ''; }
        else cur += ch;
      }
      arr.push(cur);
      return arr.map(s => s.replace(/\u00A0/g, ' ').trim().replace(/^"|"$/g, ''));
    }

    function normalize(h) { return (h || '').toString().trim().toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/\s+/g, ' '); }

    const header = parseLine(lines[0]).map(normalize);
    const expected = { filial: ['filial'], cliente: ['cliente', 'nome', 'titular'], assunto: ['assunto', 'motivo', 'tipo'], setor: ['setor'], cidade: ['cidade', 'city'], endereco: ['endereco', 'logradouro', 'rua', 'address'], complemento: ['complemento'], condominio: ['condominio'], bloco: ['bloco'], apartamento: ['apartamento', 'apto'], bairro: ['bairro', 'neighborhood'], referencia: ['referencia', 'observacao', 'obs'] };
    const idx = {};

    for (const k in expected) {
      idx[k] = -1;
      for (let i = 0; i < header.length; i++) {
        if (expected[k].some(v => header[i].includes(v))) {
          idx[k] = i;
          break;
        }
      }
    }

    const fallbackOrder = ['filial', 'cliente', 'assunto', 'setor', 'cidade', 'endereco', 'complemento', 'condominio', 'bloco', 'apartamento', 'bairro', 'referencia'];
    const hasHeader = Object.values(idx).some(v => v !== -1);
    const total = lines.length;
    const SMALL = 5000;

    if (total <= SMALL) {
      const out = [];
      for (let li = 0; li < total; li++) {
        const cols = parseLine(lines[li]);
        if (li === 0 && hasHeader) continue;

        const it = { filial: '', cliente: '', assunto: '', setor: '', cidade: '', endereco: '', complemento: '', condominio: '', bloco: '', apartamento: '', bairro: '', referencia: '' };

        if (hasHeader) {
          for (const k in idx) if (idx[k] >= 0 && idx[k] < cols.length) it[k] = cols[idx[k]].trim();
          if (cols.length > header.length && idx.endereco >= 0) { const start = idx.endereco; const last = cols.length - 1; const joinEnd = Math.max(start, last - 1); if (joinEnd > start) it.endereco = cols.slice(start, joinEnd + 1).join(', ').trim(); }
        } else {
          for (let i = 0; i < fallbackOrder.length && i < cols.length; i++) it[fallbackOrder[i]] = (cols[i] || '').trim();
          if (cols.length > fallbackOrder.length) { it.bairro = cols[cols.length - 1].trim(); it.endereco = cols.slice(5, cols.length - 1).join(', ').trim(); }
        }

        for (const k in it) if (typeof it[k] === 'string') it[k] = it[k].replace(/^\uFEFF/, '').replace(/\u00A0/g, ' ').trim();

        const assuntoRaw = (it.assunto || '').toString();
        if (!subjectAllowed(assuntoRaw)) continue;

        out.push(it);
        if (li % 100 === 0) onProgress(Math.min(li / total, 1));
      }
      onProgress(1); return out;
    }

    return new Promise((resolve) => {
      const out = []; let li = 0;
      const batch = Math.min(5000, Math.max(800, Math.floor(total / 25))); // Tamanho do lote

      function work() {
        const end = Math.min(li + batch, total);
        for (; li < end; li++) {
          const cols = parseLine(lines[li]);
          if (li === 0 && hasHeader) continue;
          const it = { filial: '', cliente: '', assunto: '', setor: '', cidade: '', endereco: '', complemento: '', condominio: '', bloco: '', apartamento: '', bairro: '', referencia: '' };
          if (hasHeader) { for (const k in idx) if (idx[k] >= 0 && idx[k] < cols.length) it[k] = cols[idx[k]].trim(); if (cols.length > header.length && idx.endereco >= 0) { const start = idx.endereco; const last = cols.length - 1; const joinEnd = Math.max(start, last - 1); if (joinEnd > start) it.endereco = cols.slice(start, joinEnd + 1).join(', ').trim(); } }
          else { for (let i = 0; i < fallbackOrder.length && i < cols.length; i++) it[fallbackOrder[i]] = (cols[i] || '').trim(); if (cols.length > fallbackOrder.length) { it.bairro = cols[cols.length - 1].trim(); it.endereco = cols.slice(5, cols.length - 1).join(', ').trim(); } }
          for (const k in it) if (typeof it[k] === 'string') it[k] = it[k].replace(/^\uFEFF/, '').replace(/\u00A0/g, ' ').trim();
          const assuntoRaw = (it.assunto || '').toString();
          if (!subjectAllowed(assuntoRaw)) continue;
          out.push(it);
        }
        showProgress(Math.min(li / total, 1));
        if (li < total) setTimeout(work, 0); else resolve(out);
      }
      work();
    });
  }

  function loadSheetJs() {
    return new Promise((res, rej) => {
      if (window.XLSX) return res(window.XLSX);
      const cdns = ['https://cdn.sheetjs.com/xlsx-latest/package/dist/xlsx.full.min.js', 'https://cdn.jsdelivr.net/npm/xlsx/dist/xlsx.full.min.js'];
      let i = 0;
      function next() {
        if (i >= cdns.length) return rej(new Error('SheetJS não disponível'));
        const s = document.createElement('script');
        s.src = cdns[i++]; s.async = true;
        s.onload = () => window.XLSX ? res(window.XLSX) : next();
        s.onerror = () => setTimeout(next, 200);
        document.head.appendChild(s);
      }
      next();
    });
  }

  function mapRows(rows) {
    if (!Array.isArray(rows)) return [];
    const expected = { filial: ['filial'], cliente: ['cliente', 'nome', 'titular'], assunto: ['assunto', 'motivo', 'tipo'], setor: ['setor'], cidade: ['cidade', 'city'], endereco: ['endereco', 'logradouro', 'rua', 'address'], complemento: ['complemento'], condominio: ['condominio'], bloco: ['bloco'], apartamento: ['apartamento', 'apto'], bairro: ['bairro', 'neighborhood'], referencia: ['referencia', 'observacao', 'obs'] };
    const norm = k => (k || '').toString().trim().toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/\s+/g, ' ');
    const headers = Object.keys(rows[0] || {});
    const map = {};

    headers.forEach(h => {
      const nh = norm(h);
      for (const key in expected) if (expected[key].some(v => nh.includes(v))) { map[h] = key; break; }
    });

    const out = [];
    for (const r of rows) {
      const it = { filial: '', cliente: '', assunto: '', setor: '', cidade: '', endereco: '', complemento: '', condominio: '', bloco: '', apartamento: '', bairro: '', referencia: '' };
      let filled = 0;

      for (const h of headers) {
        const v = r[h] == null ? '' : String(r[h]).trim();
        if (map[h]) { it[map[h]] = v; filled++; }
      }

      if (filled === 0) {
        const vals = headers.map(h => r[h] == null ? '' : String(r[h]).trim());
        const order = ['filial', 'cliente', 'assunto', 'setor', 'cidade', 'endereco', 'complemento', 'condominio', 'bloco', 'apartamento', 'bairro', 'referencia'];
        for (let i = 0; i < vals.length && i < order.length; i++) it[order[i]] = vals[i] || '';
      }

      for (const k in it) if (typeof it[k] === 'string') it[k] = it[k].replace(/^\uFEFF/, '').replace(/\u00A0/g, ' ').trim();

      const assuntoRaw = (it.assunto || '').toString();
      if (!subjectAllowed(assuntoRaw)) continue;

      out.push(it);
    }
    return out;
  }

  function buildAddress(it) {
    const p = [];
    if (it.endereco) p.push(it.endereco);
    if (it.complemento) p.push(it.complemento);
    if (it.condominio) p.push(it.condominio);
    if (it.bloco) p.push('Bl.' + it.bloco);
    if (it.apartamento) p.push('Apt.' + it.apartamento);
    if (it.bairro) p.push(it.bairro);
    if (it.cidade) p.push(it.cidade);
    return p.filter(Boolean).join(', ');
  }

  function hasAddress(it) {
    const a = (it.endereco || '').toLowerCase().trim();
    if (!a) return false;
    const neg = ['não informado', 'nao informado', 'não localizado', 'nao localizado', 'endereço não encontrado', 'endereco nao encontrado', 'endereco nao informado', 'endereco nao localizado'];
    for (const n of neg) if (a.includes(n)) return false;
    return true;
  }

  function renderSkeleton(count = 6) {
    resultsEl.innerHTML = '';
    for (let i = 0; i < count; i++) {
      const s = document.createElement('div'); s.className = 'card result';
      s.innerHTML = `
        <div class="card-header">
          <div class="card-title loader" style="width: 70%; margin-top: 4px;">
            <div class="skeleton-line" style="height: 16px; width: 80%;"></div>
            <div class="skeleton-line" style="height: 12px; width: 60%;"></div>
          </div>
          <div class="chip skeleton" style="width: 50px; height: 20px;"></div>
        </div>
        <div class="address skeleton" style="height: 12px; width: 95%;"></div>
        <div class="actions skeleton" style="height: 38px; width: 100%; margin-top: 8px;"></div>
      `;
      resultsEl.appendChild(s);
    }
  }

  function createCardHtml(it, idx) {
    const addr = buildAddress(it);
    const displayAddr = addr || 'Endereço não informado';
    const query = addr ? displayAddr : it.cidade || 'Brasil';
    const maps = 'https://www.google.com/maps/search/?api=1&query=' + encodeURIComponent(query);

    const noAddress = !hasAddress(it);
    const cardClasses = `result ${noAddress ? 'no-address' : ''}`;

    return `
      <div class="${cardClasses} card" data-idx="${idx}" tabindex="0" data-tilt="true">
        <div class="card-header">
          <div class="card-title">
            <h3>${escapeHtml(it.cliente || '—')}</h3>
            <div class="sub"><strong>Assunto:</strong> ${escapeHtml(it.assunto || '—')}</div>
          </div>
          <div class="card-meta">
            <div class="chip">${escapeHtml(it.filial || '—')}</div>
            ${noAddress ? `<div style="margin-top:8px"><span class="badge">❗ Sem endereço</span></div>` : `<div class="muted text-sm" style="margin-top:8px">${escapeHtml(it.cidade || '—')}</div>`}
          </div>
        </div>

        <div class="address"><strong>Endereço:</strong> ${escapeHtml(displayAddr)}</div>

        <div class="actions">
          <a class="btn btn-primary" href="${escapeHtml(maps)}" target="_blank" onclick="event.stopPropagation();" style="text-decoration:none">📍 Abrir mapa</a>
          <button class="btn btn-ghost copy-btn" data-url="${escapeHtml(maps)}" data-copy-index="${idx}">🔗 Copiar link</button>
          <span class="copy-feedback" data-copy-feedback="${idx}" style="display:none;margin-left:auto">Copiado!</span>
        </div>
      </div>
    `;
  }

  function renderInitial() {
    rendered = 0; resultsEl.innerHTML = ''; renderMore();
  }

  function renderMore() {
    const filter = filterEl && filterEl.value ? filterEl.value.toLowerCase().trim() : '';
    const list = filter ? items.filter(it => JSON.stringify(it).toLowerCase().includes(filter)) : items;

    list.sort((a, b) => hasAddress(b) - hasAddress(a));

    const slice = list.slice(rendered, rendered + PER_PAGE);
    if (slice.length === 0 && rendered === 0) {
      resultsEl.innerHTML = '<div class="card text-center muted mt-4">Nenhuma O.S. encontrada com o filtro atual.</div>';
      if (moreBtn) moreBtn.classList.add('hidden');
      setStatus('Nenhuma O.S. encontrada');
      return;
    }

    slice.forEach((it, i) => resultsEl.insertAdjacentHTML('beforeend', createCardHtml(it, rendered + i)));
    rendered += slice.length;

    if (moreBtn) {
      if (rendered >= list.length) moreBtn.classList.add('hidden');
      else moreBtn.classList.remove('hidden');
    }

    setStatus(`Mostrando ${Math.min(rendered, list.length)} de ${list.length} O.S.`);
    staggerEntrance();
    enableTiltOnResults();
  }

  function staggerEntrance() {
    const nodes = Array.from(resultsEl.querySelectorAll('.result')).slice(-PER_PAGE);
    nodes.forEach((el, i) => {
      el.style.opacity = 0;
      el.style.transform = 'translateY(12px) scale(.995)';
      el.style.transition = 'opacity 420ms var(--transition-smooth), transform 420ms var(--transition-smooth)';
      setTimeout(() => {
        el.style.opacity = 1;
        el.style.transform = 'none';
      }, i * 40);
    });
  }

  function enableTiltOnResults(selector = '.result[data-tilt="true"]') {
    const els = document.querySelectorAll(selector);
    els.forEach(el => {
      if (el._tiltInitialized) return;
      el._tiltInitialized = true;
      el.style.transition = 'transform var(--transition-fast)';

      el.addEventListener('pointermove', (ev) => {
        const r = el.getBoundingClientRect();
        const px = (ev.clientX - r.left) / r.width;
        const py = (ev.clientY - r.top) / r.height;
        const rotateY = (px - 0.5) * -4;
        const rotateX = (py - 0.5) * 4;
        el.style.transform = `rotateX(${rotateX}deg) rotateY(${rotateY}deg) scale(1.005)`;
        el.style.boxShadow = 'var(--shadow-lg)';
      });

      el.addEventListener('pointerleave', () => {
        el.style.transform = 'none';
        el.style.boxShadow = 'var(--shadow-md)';
      });
    });
  }

  if (dropZone) {
    ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(ev => dropZone.addEventListener(ev, e => e.preventDefault()));
    dropZone.addEventListener('dragover', () => dropZone.classList.add('dragover'));
    dropZone.addEventListener('dragleave', () => dropZone.classList.remove('dragover'));
    dropZone.addEventListener('drop', (e) => {
      dropZone.classList.remove('dragover');
      const dt = e.dataTransfer;
      if (!dt || !dt.files || dt.files.length === 0) { setStatus('Nenhum arquivo detectado no arraste.', 'error'); return; }
      selectedFile = dt.files[0];
      syncInputFiles(selectedFile);
      updateFileDisplay(selectedFile);
      if (loadBtn) loadBtn.disabled = false;
      setStatus('Arquivo solto. Clique em CARREGAR.');
      log('drop ->', selectedFile.name);
    });
    dropZone.addEventListener('click', () => fileInput && fileInput.click());
  }

  if (fileInput) fileInput.addEventListener('change', () => {
    const f = fileInput.files && fileInput.files[0] ? fileInput.files[0] : null;
    selectedFile = f;
    updateFileDisplay(f);
    if (loadBtn) loadBtn.disabled = !f;
    setStatus(f ? 'Arquivo pronto. Clique em CARREGAR.' : 'Nenhum arquivo selecionado.', f ? 'info' : 'error');
    log('input.change ->', f ? f.name : null);
  });

  resultsEl.addEventListener('click', async (e) => {
    const copyBtn = e.target.closest && e.target.closest('.copy-btn');
    if (copyBtn) {
      e.stopPropagation();
      const url = copyBtn.getAttribute('data-url');
      const idx = copyBtn.getAttribute('data-copy-index');
      try {
        if (navigator.clipboard && navigator.clipboard.writeText) await navigator.clipboard.writeText(url);
        else { const ta = document.createElement('textarea'); ta.value = url; document.body.appendChild(ta); ta.select(); document.execCommand('copy'); document.body.removeChild(ta); }
        const card = copyBtn.closest('.result');
        const feed = card && card.querySelector(`[data-copy-feedback="${idx}"]`);
        if (feed) {
          feed.style.display = 'inline-block';
          feed.classList.add('show');
          setTimeout(() => { feed.classList.remove('show'); feed.style.display = 'none'; }, 1400);
        }
      } catch (err) {
        console.error(err);
        setStatus('Erro ao copiar link', 'error');
      }
      return;
    }

    const card = e.target.closest && e.target.closest('.result');
    if (card) {
      openExpandedClone(card);
    }
  });

  resultsEl.addEventListener('keydown', (e) => {
    if (e.key === 'Enter' || e.key === ' ') {
      const card = e.target.closest && e.target.closest('.result');
      if (card) { e.preventDefault(); openExpandedClone(card); }
    }
  });

  function openExpandedClone(card) {
    if (!card) return;
    closeExistingClone();

    const clone = card.cloneNode(true);
    clone.classList.add('result--expanded');
    clone.removeAttribute('data-tilt'); 

    const inner = document.createElement('div');
    inner.className = 'expanded-content';
    inner.style.position = 'relative';
    inner.style.paddingRight = '8px'; 
    while (clone.firstChild) inner.appendChild(clone.firstChild);
    clone.appendChild(inner);

    const closeBtn = document.createElement('button');
    closeBtn.className = 'expanded-close';
    closeBtn.innerHTML = '✕';
    closeBtn.setAttribute('aria-label', 'Fechar detalhe');
    clone.appendChild(closeBtn);

    document.body.appendChild(clone);
    if (dim) dim.classList.remove('hidden');
    card.style.visibility = 'hidden'; 

    clone.querySelectorAll('.copy-btn').forEach(btn => {
      btn.addEventListener('click', async (ev) => {
        ev.stopPropagation();
        const url = btn.getAttribute('data-url');
        try {
          if (navigator.clipboard && navigator.clipboard.writeText) await navigator.clipboard.writeText(url);
          else { const ta = document.createElement('textarea'); ta.value = url; document.body.appendChild(ta); ta.select(); document.execCommand('copy'); document.body.removeChild(ta); }
          const feed = clone.querySelector('.copy-feedback');
          if (feed) { feed.style.display = 'inline-block'; feed.classList.add('show'); setTimeout(() => { feed.classList.remove('show'); feed.style.display = 'none'; }, 1400); }
        } catch (err) { console.error(err); setStatus('Erro ao copiar', 'error'); }
      });
    });

    const removeClone = () => {
      clone.remove();
      if (dim) dim.classList.add('hidden');
      card.style.visibility = '';
      document.removeEventListener('keydown', onKey);
    };

    closeBtn.addEventListener('click', removeClone);
    if (dim) dim.addEventListener('click', removeClone, { once: true });
    function onKey(ev) { if (ev.key === 'Escape') removeClone(); }
    document.addEventListener('keydown', onKey);
    closeBtn.focus();
  }

  function closeExistingClone() {
    const existing = document.querySelector('.result--expanded');
    if (existing) existing.remove();
    if (dim) dim.classList.add('hidden');
    document.querySelectorAll('.result[style*="visibility: hidden"]').forEach(el => el.style.visibility = '');
  }

  if (loadBtn) {
    loadBtn.addEventListener('click', async () => {
      if (!selectedFile && fileInput && fileInput.files && fileInput.files[0]) selectedFile = fileInput.files[0];
      if (!selectedFile) {
        setStatus('Nenhum arquivo selecionado. Selecione ou arraste um arquivo antes de carregar.', 'error');
        return;
      }

      loadBtn.disabled = true; if (fileInput) fileInput.disabled = true;
      renderSkeleton(6); showProgress(0.03); setStatus('Processando arquivo...');

      try {
        const name = (selectedFile.name || '').toLowerCase();
        let parsed = [];

        if (name.endsWith('.xls') || name.endsWith('.xlsx')) {
          if (!window.XLSX) await loadSheetJs();
          if (!window.XLSX) throw new Error('SheetJS indisponível.');
          const ab = await selectedFile.arrayBuffer();
          const wb = window.XLSX.read(ab, { type: 'array' });
          const sheet = wb.SheetNames && wb.SheetNames[0];
          if (!sheet) throw new Error('Planilha sem abas.');
          const rows = window.XLSX.utils.sheet_to_json(wb.Sheets[sheet], { defval: '' });
          parsed = mapRows(rows);
        } else {
          const text = await readAsTextDetectEncoding(selectedFile);
          const delim = detectDelimiter(text);
          const maybe = parseCSVText(text, delim, (p) => showProgress(p));
          parsed = Array.isArray(maybe) ? maybe : await maybe;
        }

        items = parsed;
        rendered = 0;
        if (filterEl) filterEl.value = '';
        showProgress(1);
        setStatus(`Total carregado: ${items.length} O.S.`);
        const missing = items.filter(it => !hasAddress(it)).length;
        if (missing > 0) setStatus(`${items.length} O.S. carregadas — ${missing} sem endereço.`, missing > 0 ? 'info' : 'ok');
        renderInitial();

      } catch (err) {
        console.error(err);
        setStatus('Erro ao processar arquivo — veja console.', 'error');
        showMessage('Erro de Processamento', 'Não foi possível ler o arquivo. Motivo: ' + (err && err.message ? err.message : 'Erro desconhecido. Verifique o formato do arquivo.'), true);
        resultsEl.innerHTML = '';
      } finally {
        loadBtn.disabled = false; if (fileInput) fileInput.disabled = false;
        setTimeout(() => showProgress(0), 250);
      }
    });
  }

  function debounce(fn, t = 180) { let h; return (...a) => { clearTimeout(h); h = setTimeout(() => fn(...a), t); }; }

  if (filterEl) filterEl.addEventListener('input', debounce(() => {
    rendered = 0; resultsEl.innerHTML = ''; renderMore();
  }, 200));

  if (moreBtn) moreBtn.addEventListener('click', () => renderMore());

  if (loadBtn) loadBtn.disabled = true;
  updateFileDisplay(null);
  window._athon = { itemsRef: () => items };

  log('log.js v2.0 updated (Modal/Tilt) — ready');
});


