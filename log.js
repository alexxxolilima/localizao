document.addEventListener('DOMContentLoaded', () => {
  'use strict';

  const ICONS = {
    map: `<svg class="icon" viewBox="0 0 24 24"><path d="M12 2C8.13 2 5 5.13 5 9c0 5.25 7 13 7 13s7-7.75 7-13c0-3.87-3.13-7-7-7zm0 9.5c-1.38 0-2.5-1.12-2.5-2.5s1.12-2.5 2.5-2.5 2.5 1.12 2.5 2.5-1.12 2.5-2.5 2.5z"/></svg>`,
    copy: `<svg class="icon" viewBox="0 0 24 24"><path d="M16 1H4c-1.1 0-2 .9-2 2v14h2V3h12V1zm3 4H8c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h11c1.1 0 2-.9 2-2V7c0-1.1-.9-2-2-2zm0 16H8V7h11v14z"/></svg>`,
    search: `<svg class="icon" viewBox="0 0 24 24"><path d="M15.5 14h-.79l-.28-.27C15.41 12.59 16 11.11 16 9.5 16 5.91 13.09 3 9.5 3S3 5.91 3 9.5 5.91 16 9.5 16c1.61 0 3.09-.59 4.23-1.57l.27.28v.79l5 4.99L20.49 19l-4.99-5zm-6 0C7.01 14 5 11.99 5 9.5S7.01 5 9.5 5 14 7.01 14 9.5 11.99 14 9.5 14z"/></svg>`,
    boat: `<svg class="icon" viewBox="0 0 24 24"><path d="M20 21c-1.39 0-2.78-.47-4-1.32-2.44 1.71-5.56 1.71-8 0C6.78 20.53 5.39 21 4 21H2v2h2c1.38 0 2.74-.35 4-.99 2.52 1.29 5.48 1.29 8 0 1.26.65 2.62.99 4 .99h2v-2h-2zM3.95 19H4c1.6 0 3.02-.88 4-2 .98 1.12 2.4 2 4 2s3.02-.88 4-2c.98 1.12 2.4 2 4 2h.05l.02-1.91C20.03 17.04 20.01 17 20 17c-1.66 0-3-1.34-3-3 0-2 3-5.4 3-5.4V7h-6V6h-4v1H4v1.6s3 3.4 3 5.4c0 1.66-1.34 3-3 3 0 0-.03.04-.07.09L3.95 19z"/></svg>`,
    warning: `<svg class="icon" viewBox="0 0 24 24"><path d="M1 21h22L12 2 1 21zm12-3h-2v-2h2v2zm0-4h-2v-4h2v4z"/></svg>`,
    bing: `<svg class="icon" viewBox="0 0 24 24"><path d="M3 6l14-4 4 2-6 16L3 20V6z"/></svg>`
  };

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
  const modalBody = document.getElementById('modalBody');
  const toggleTheme = document.getElementById('toggleTheme');
  const cardTemplate = document.getElementById('cardTemplate');
  const modalTemplate = document.getElementById('modalTemplate');

  let selectedFile = null;
  let items = [];
  let rendered = 0;
  const PER_PAGE = 50;
  const BLOCKED_SUBJECT_RE = /\b(retirada|reagend|tentativa|^re$)\b/i;

  function setStatus(txt) {
    if (fileStatus) fileStatus.textContent = txt;
  }

  function showProgress(p = 0) {
    if (!progressRow || !progressBar) return;
    const v = Math.max(0, Math.min(1, p));
    const visible = v > 0.01 && v < 0.99;
    progressRow.classList.toggle('hidden', !visible);
    progressBar.style.width = Math.round(v * 100) + '%';
  }

  function updateFileDisplay(file) {
    if (!file) {
      fileDisplayName.textContent = 'Escolha o arquivo ';
      fileStatus.textContent = '';
      loadBtn.disabled = true;
      return;
    }
    const sizeKB = Math.round(file.size / 1024);
    fileDisplayName.textContent = `${file.name}`;
    fileStatus.textContent = '';
    loadBtn.disabled = false;
  }

  function readAsTextDetectEncoding(file) {
    return file.arrayBuffer().then(ab => {
      const bytes = new Uint8Array(ab);
      if (bytes[0] === 0xFF && bytes[1] === 0xFE) return new TextDecoder('utf-16le').decode(ab);
      if (bytes[0] === 0xFE && bytes[1] === 0xFF) return new TextDecoder('utf-16be').decode(ab);
      if (bytes[0] === 0xEF && bytes[1] === 0xBB && bytes[2] === 0xBF) return new TextDecoder('utf-8').decode(ab);
      try {
        return new TextDecoder('utf-8', { fatal: true }).decode(ab);
      } catch (e) {
        return new TextDecoder('iso-8859-1').decode(ab);
      }
    });
  }

  function detectDelimiter(text) {
    const firstLine = text.split('\n')[0];
    const commas = (firstLine.match(/,/g) || []).length;
    const semicolons = (firstLine.match(/;/g) || []).length;
    const tabs = (firstLine.match(/\t/g) || []).length;
    if (tabs > commas && tabs > semicolons) return '\t';
    if (semicolons > commas) return ';';
    return ',';
  }

  function parseCSV(text, delimiter) {
    const lines = text.replace(/\r\n/g, '\n').split('\n').filter(l => l.trim().length > 0);
    if (lines.length === 0) return [];

    const parseLine = (line) => {
      const res = [];
      let cur = '';
      let inQuotes = false;
      for (let i = 0; i < line.length; i++) {
        const char = line[i];
        if (char === '"') {
          if (inQuotes && line[i + 1] === '"') {
            cur += '"';
            i++;
          } else {
            inQuotes = !inQuotes;
          }
        } else if (char === delimiter && !inQuotes) {
          res.push(cur);
          cur = '';
        } else {
          cur += char;
        }
      }
      res.push(cur);
      return res.map(c => c.trim().replace(/^"|"$/g, ''));
    };

    const headers = parseLine(lines[0]);
    const result = [];
    
    for (let i = 1; i < lines.length; i++) {
      const row = parseLine(lines[i]);
      const obj = {};
      headers.forEach((h, idx) => {
        obj[h] = row[idx] || '';
      });
      result.push(obj);
    }
    return result;
  }

  async function loadSheetJs() {
    if (window.XLSX) return window.XLSX;
    return new Promise((resolve, reject) => {
      const script = document.createElement('script');
      script.src = "https://cdn.sheetjs.com/xlsx-latest/package/dist/xlsx.full.min.js";
      script.onload = () => resolve(window.XLSX);
      script.onerror = () => reject(new Error("Falha ao baixar SheetJS"));
      document.head.appendChild(script);
    });
  }

  function mapData(rows) {
    const mapping = {
      filial: ['filial'],
      cliente: ['cliente', 'nome', 'titular', 'assinante'],
      assunto: ['assunto', 'motivo', 'tipo', 'serviço'],
      setor: ['setor', 'area'],
      cidade: ['cidade', 'city', 'municipio'],
      endereco: ['endereco', 'logradouro', 'rua', 'avenida'],
      complemento: ['complemento'],
      condominio: ['condominio'],
      bloco: ['bloco'],
      apartamento: ['apartamento', 'apto', 'apt'],
      bairro: ['bairro', 'neighborhood'],
      referencia: ['referencia', 'obs', 'observacao']
    };

    const normalizeKey = k => k.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, "").trim();
    const fileKeys = Object.keys(rows[0] || {});
    const keyMap = {}; 

    fileKeys.forEach(fKey => {
      const normFKey = normalizeKey(fKey);
      for (const [target, patterns] of Object.entries(mapping)) {
        if (patterns.some(p => normFKey.includes(p))) {
          keyMap[fKey] = target;
          break;
        }
      }
    });

    return rows.map((row, idx) => {
      const newItem = { __idx: idx };
      for (const [target] of Object.entries(mapping)) {
        newItem[target] = ''; 
      }
      for (const [fKey, val] of Object.entries(row)) {
        const target = keyMap[fKey];
        if (target) {
          newItem[target] = String(val || '').trim();
        }
      }
      if (BLOCKED_SUBJECT_RE.test(newItem.assunto)) return null;
      return newItem;
    }).filter(Boolean);
  }

  function buildAddress(it) {
    const parts = [
      it.endereco,
      it.complemento,
      it.condominio ? `Cond. ${it.condominio}` : null,
      it.bloco ? `Bl. ${it.bloco}` : null,
      it.apartamento ? `Apt. ${it.apartamento}` : null,
      it.bairro,
      it.cidade
    ];
    return parts.filter(p => p && p.trim().length > 0).join(', ');
  }

  function checkBalsaRegion(it) {
    const text = (it.endereco + ' ' + it.bairro + ' ' + it.referencia).toLowerCase();
    const keywords = [
      'balsa', 'riacho grande', 'tatetos', 'taquacetuba', 'pos balsa', 
      'pós balsa', 'curucutu', 'capivari', 'finco', 'santa cruz', 'estrada do rio acima'
    ];
    return keywords.some(k => text.includes(k));
  }

  function createCard(it) {
    const frag = cardTemplate.content.cloneNode(true);
    const el = frag.querySelector('.result-card');
    
    const isBalsa = checkBalsaRegion(it);
    const fullAddr = buildAddress(it);
    const hasAddress = fullAddr && fullAddr.length > 8 && !/não informado|não localizado/i.test(fullAddr);

    if (isBalsa) el.classList.add('is-balsa');
    if (!hasAddress) el.classList.add('is-error');

    const badgeContainer = el.querySelector('.card-badges') || el.insertBefore(document.createElement('div'), el.firstChild);
    if (isBalsa) {
      const b = document.createElement('div');
      b.className = 'special-badge badge-balsa';
      b.innerHTML = ICONS.boat + ' Região de Balsa';
      badgeContainer.appendChild(b);
    }
    if (!hasAddress) {
      const b = document.createElement('div');
      b.className = 'special-badge badge-error';
      b.innerHTML = ICONS.warning + ' Endereço Inválido';
      badgeContainer.appendChild(b);
    }

    el.querySelector('.card-client').textContent = it.cliente || 'Cliente Desconhecido';
    el.querySelector('.card-subject').textContent = it.assunto || '—';
    el.querySelector('.card-branch').textContent = it.filial || '—';
    el.querySelector('.card-city').textContent = it.cidade || '—';
    el.querySelector('.card-address').innerHTML = hasAddress ? fullAddr : 'Endereço não identificado';

    const q = encodeURIComponent(fullAddr || it.cidade || 'Brasil');
    
    const btnMaps = el.querySelector('.open-maps');
    btnMaps.href = `https://www.google.com/maps/search/?api=1&query=${q}`;
    btnMaps.innerHTML = ICONS.map + ' Maps';

    const btnBing = el.querySelector('.open-bing');
    btnBing.href = `https://www.bing.com/maps?q=${q}`;
    btnBing.innerHTML = ICONS.bing + ' Bing';

    const btnCopy = el.querySelector('.copy-btn');
    btnCopy.innerHTML = ICONS.copy + ' Copiar';
    btnCopy.onclick = (e) => {
      e.stopPropagation();
      navigator.clipboard.writeText(fullAddr).then(() => {
        btnCopy.innerHTML = '✅ Copiado';
        setTimeout(() => btnCopy.innerHTML = ICONS.copy + ' Copiar', 2000);
      });
    };

    const btnDetails = el.querySelector('.details-btn');
    btnDetails.innerHTML = ICONS.search + ' Detalhes';
    btnDetails.onclick = (e) => {
      e.stopPropagation();
      openModal(it);
    };

    el.onclick = (e) => {
      if (e.target.tagName !== 'A' && e.target.tagName !== 'BUTTON') openModal(it);
    };

    return el;
  }

  function renderBatch() {
    const filterVal = filterEl.value.toLowerCase();
    const filtered = filterVal 
      ? items.filter(it => JSON.stringify(it).toLowerCase().includes(filterVal)) 
      : items;

    if (rendered >= filtered.length) {
      moreBtn.classList.add('hidden');
      if (rendered === 0) resultsEl.innerHTML = '<div style="grid-column:1/-1;text-align:center;padding:20px;color:#888">Nenhum resultado encontrado.</div>';
      return;
    }

    moreBtn.classList.remove('hidden');
    const batch = filtered.slice(rendered, rendered + PER_PAGE);
    const frag = document.createDocumentFragment();
    batch.forEach(it => frag.appendChild(createCard(it)));
    resultsEl.appendChild(frag);
    rendered += batch.length;
    
    if (rendered >= filtered.length) moreBtn.classList.add('hidden');
  }

  function openModal(it) {
    const content = modalTemplate.content.cloneNode(true);
    const address = buildAddress(it);
    const isBalsa = checkBalsaRegion(it);
    const q = encodeURIComponent(address || it.cidade || '');

    content.querySelector('.modal-client').textContent = it.cliente;
    content.querySelector('.modal-branch').textContent = it.filial;
    content.querySelector('.modal-setor').textContent = it.setor;
    content.querySelector('.modal-subject').textContent = it.assunto;
    content.querySelector('.modal-address').textContent = address;
    content.querySelector('.modal-complement').textContent = it.complemento || '—';
    content.querySelector('.modal-condo').textContent = it.condominio || '—';
    content.querySelector('.modal-blockapt').textContent = (it.bloco ? 'Bl '+it.bloco : '') + (it.apartamento ? ' Ap '+it.apartamento : '') || '—';
    content.querySelector('.modal-district').textContent = it.bairro || '—';

    const btns = content.querySelectorAll('.provider-btn');
    const mapContainer = content.getElementById('mapContainer');
    
    const loadMap = (provider) => {
      btns.forEach(b => b.classList.toggle('active', b.dataset.provider === provider));
      let src = '';
      if (provider === 'google') {
        const type = isBalsa ? 'k' : 'm'; 
        src = `https://maps.google.com/maps?q=${q}&t=${type}&z=16&ie=UTF8&iwloc=&output=embed`;
      } else if (provider === 'bing') {
        src = `https://www.bing.com/maps/embed?h=350&w=500&cp=&lvl=16&typ=d&sty=r&src=SHELL&FORM=MBEDV8&q=${q}`;
      }
      
      if (provider === 'waze') return; 

      mapContainer.innerHTML = `<iframe class="map-iframe" src="${src}" loading="lazy"></iframe>`;
    };

    btns.forEach(btn => {
      btn.onclick = () => {
        if (btn.dataset.openExternal) {
          window.open(`https://waze.com/ul?q=${q}`, '_blank');
        } else {
          loadMap(btn.dataset.provider);
        }
      };
    });

    content.querySelector('.copy-link').onclick = (e) => {
      navigator.clipboard.writeText(`https://www.google.com/maps/search/?api=1&query=${q}`);
      e.target.textContent = 'Link copiado!';
      setTimeout(() => e.target.textContent = 'Copiar Link', 2000);
    };

    loadMap('google');

    modalBody.innerHTML = '';
    modalBody.appendChild(content);
    overlay.classList.remove('hidden');
    dim.classList.remove('hidden');
    document.body.style.overflow = 'hidden';
  }

  function closeModalFunc() {
    overlay.classList.add('hidden');
    dim.classList.add('hidden');
    document.body.style.overflow = '';
  }

  const handleFileSelect = (file) => {
    if (!file) return;
    selectedFile = file;
    updateFileDisplay(file);
    fileInput.value = ''; 
  };

  fileInput.addEventListener('change', (e) => {
    if (e.target.files.length > 0) handleFileSelect(e.target.files[0]);
  });

  dropZone.addEventListener('dragover', (e) => {
    e.preventDefault();
    dropZone.classList.add('dragover');
  });

  dropZone.addEventListener('dragleave', () => dropZone.classList.remove('dragover'));

  dropZone.addEventListener('drop', (e) => {
    e.preventDefault();
    dropZone.classList.remove('dragover');
    if (e.dataTransfer.files.length > 0) {
      handleFileSelect(e.dataTransfer.files[0]);
    }
  });

  loadBtn.addEventListener('click', async (e) => {
    e.preventDefault(); 
    e.stopPropagation(); 

    if (!selectedFile) return;
    loadBtn.disabled = true;
    loadBtn.textContent = 'Processando...';
    showProgress(0.2);
    setStatus('Lendo arquivo...');
    resultsEl.innerHTML = '';
    rendered = 0;

    try {
      let data = [];
      const name = selectedFile.name.toLowerCase();

      if (name.endsWith('.csv') || name.endsWith('.txt')) {
        const text = await readAsTextDetectEncoding(selectedFile);
        const delim = detectDelimiter(text);
        const raw = parseCSV(text, delim);
        data = mapData(raw);
      } else {
        setStatus('Carregando planilha...');
        const XLSX = await loadSheetJs();
        const ab = await selectedFile.arrayBuffer();
        const wb = XLSX.read(ab, { type: 'array' });
        const sheetName = wb.SheetNames[0];
        const raw = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { defval: '' });
        data = mapData(raw);
      }

      items = data;
      showProgress(1);
      setStatus(`Sucesso! ${items.length} registros.`);
      renderBatch();

    } catch (err) {
      console.error(err);
      setStatus('Erro: ' + err.message);
      alert('Erro ao ler arquivo.');
    } finally {
      loadBtn.disabled = false;
      loadBtn.textContent = 'CARREGAR';
      setTimeout(() => showProgress(0), 1000);
    }
  });

  let debounce;
  filterEl.addEventListener('input', () => {
    clearTimeout(debounce);
    debounce = setTimeout(() => {
      rendered = 0;
      resultsEl.innerHTML = '';
      renderBatch();
    }, 300);
  });

  moreBtn.addEventListener('click', renderBatch);
  closeModal.addEventListener('click', closeModalFunc);
  dim.addEventListener('click', closeModalFunc);
  document.addEventListener('keydown', (e) => { if (e.key === 'Escape') closeModalFunc(); });

  const userPref = localStorage.getItem('theme');
  if (userPref === 'light') {
    document.body.classList.remove('dark-mode');
    toggleTheme.checked = true;
  }
  
  toggleTheme.addEventListener('change', (e) => {
    if (e.target.checked) {
      document.body.classList.remove('dark-mode');
      localStorage.setItem('theme', 'light');
    } else {
      document.body.classList.add('dark-mode');
      localStorage.setItem('theme', 'dark');
    }
  });
  
  updateFileDisplay(null);
});


