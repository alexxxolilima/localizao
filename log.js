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
  const techFilterEl = document.getElementById('techFilter');
  const resultsEl = document.getElementById('results');
  const moreBtn = document.getElementById('moreBtn');
  const exportBtn = document.getElementById('exportBtn');
  const copyLinkListBtn = document.getElementById('copyLinkListBtn');
  const openRouteBtn = document.getElementById('openRouteBtn');
  const filterChips = document.querySelectorAll('.filter-chip');
  const overlay = document.getElementById('overlay');
  const closeModal = document.getElementById('closeModal');
  const modalBody = document.getElementById('modalBody');
  const cardTemplate = document.getElementById('cardTemplate');
  const modalTemplate = document.getElementById('modalTemplate');
  const themeBtn = document.getElementById('themeBtn');

  let selectedFile = null;
  let items = [];
  let filteredItems = [];
  let rendered = 0;
  let activeCategory = 'all';
  let PER_PAGE = window.innerWidth < 600 ? 6 : window.innerWidth < 900 ? 12 : 50;
  const BLOCKED_SUBJECT_RE = /\b(retirada|reagend|tentativa|^re$)\b/i;

  function setStatus(txt) { if(fileStatus) fileStatus.textContent = txt; }
  function showProgress(p) {
    const v = Math.max(0, Math.min(1, p));
    progressRow.classList.toggle('hidden', v === 0);
    progressBar.style.width = (v * 100) + '%';
  }
  function updateFileDisplay(file) {
    if (!file) {
      fileDisplayName.textContent = 'Selecione o arquivo';
      loadBtn.disabled = true;
      return;
    }
    fileDisplayName.textContent = file.name;
    loadBtn.disabled = false;
  }

  function getGoogleMapsLink(address) {
    return `https://www.google.com/maps/search/?api=1&query=${encodeURIComponent(address)}`;
  }

  function getWhatsAppLink(phone, text) {
    if (!phone) return '#';
    const onlyDigits = String(phone).replace(/\D/g, '');
    if (!onlyDigits) return '#';
    let clean = onlyDigits;
    if (clean.length < 10) return '#';
    if (!clean.startsWith('55') && (clean.length === 10 || clean.length === 11)) clean = '55' + clean;
    const base = `https://wa.me/${clean}`;
    return text ? `${base}?text=${encodeURIComponent(text)}` : base;
  }

  function formatCredentials(it) {
    const login = it.login || '';
    const pass = it.senha || '';
    return `${login}\n \n${pass}\n \nVlan: 500`;
  }

  function getRouteLink(addresses) {
    const destinations = addresses.map(a => encodeURIComponent(a)).join('/');
    return `https://www.google.com/maps/dir/${destinations}`;
  }

  async function loadSheetJs() {
    if(window.XLSX) return window.XLSX;
    return new Promise((resolve, reject) => {
      const s = document.createElement('script');
      s.src = "https://cdn.sheetjs.com/xlsx-latest/package/dist/xlsx.full.min.js";
      s.onload = () => resolve(window.XLSX);
      s.onerror = () => reject(new Error("Erro SheetJS"));
      document.head.appendChild(s);
    });
  }

  function parseCSVManual(text) {
    const firstLine = (text || '').split('\n')[0] || '';
    const delimiter = (firstLine.match(/;/g) || []).length > (firstLine.match(/,/g) || []).length ? ';' : ',';
    const rows = [];
    let curRow = [];
    let curVal = '';
    let insideQuote = false;
    for (let i = 0; i < text.length; i++) {
        const char = text[i];
        const nextChar = text[i + 1];
        if (char === '"') {
            if (insideQuote && nextChar === '"') { curVal += '"'; i++; } 
            else { insideQuote = !insideQuote; }
        } else if (char === delimiter && !insideQuote) {
            curRow.push(curVal.trim()); curVal = '';
        } else if ((char === '\n' || char === '\r') && !insideQuote) {
            if (curVal || curRow.length > 0) { curRow.push(curVal.trim()); rows.push(curRow); }
            curRow = []; curVal = '';
            if (char === '\r' && nextChar === '\n') i++;
        } else { curVal += char; }
    }
    if (curRow.length > 0 || curVal) { curRow.push(curVal.trim()); rows.push(curRow); }
    return rows;
  }

  function readAsTextDetectEncoding(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (e) => resolve(e.target.result);
        reader.onerror = reject;
        reader.readAsText(file, 'ISO-8859-1');
    });
  }

  function mapData(rawRows) {
    if(!rawRows || rawRows.length < 2) return [];
    const headers = rawRows[0].map(h => String(h).toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, "").replace(/^"|"$/g, '').trim());
    const schema = {
      filial: ['filial'],
      cliente: ['cliente', 'nome', 'titular'],
      tecnico: ['colaborador', 'tecnico', 'responsavel'],
      assunto: ['assunto', 'motivo', 'tipo'],
      setor: ['setor'],
      cidade: ['cidade', 'municipio'],
      endereco: ['endereco', 'logradouro', 'rua'],
      complemento: ['complemento'],
      condominio: ['condominio'],
      bloco: ['bloco'],
      apartamento: ['apartamento', 'apto'],
      bairro: ['bairro'],
      referencia: ['referencia', 'obs'],
      telefone: ['telefone', 'celular', 'contato', 'reside'],
      login: ['login', 'usuario', 'pppoe'],
      senha: ['senha', 'password', 'md5']
    };
    const colMap = {};
    headers.forEach((h, idx) => {
      for (const key in schema) {
        if (schema[key].some(pat => h.includes(pat))) { 
            if (!colMap[key]) colMap[key] = [];
            colMap[key].push(idx);
        }
      }
    });
    const data = [];
    for (let i = 1; i < rawRows.length; i++) {
      const row = rawRows[i];
      if (!row || row.length < 2) continue;
      const it = {};
      for (const key in schema) {
        it[key] = '';
        if (colMap[key]) {
            for (const idx of colMap[key]) {
                let val = (row[idx] !== undefined) ? String(row[idx]) : '';
                val = val.replace(/^"|"$/g, '').trim();
                if (val) {
                    if (key === 'telefone' && it[key]) {
                        it[key] += ' / ' + val;
                    } else {
                        it[key] = val;
                        if (key !== 'telefone') break;
                    }
                }
            }
        }
      }
      if (BLOCKED_SUBJECT_RE.test(it.assunto)) continue;
      data.push(it);
    }
    return data;
  }

  function buildAddress(it) {
    const parts = [];
    if(it.endereco) parts.push(it.endereco);
    if(it.complemento) parts.push(it.complemento);
    if(it.condominio) parts.push("Cond. " + it.condominio);
    if(it.bloco) parts.push("Bl " + it.bloco);
    if(it.apartamento) parts.push("Ap " + it.apartamento);
    if(it.bairro) parts.push(it.bairro);
    if(it.cidade) parts.push(it.cidade);
    return parts.filter(p => p && p.length > 1).join(', ').replace(/, ,/g, ',');
  }

  function checkBalsa(it) {
    const txt = (it.endereco + ' ' + it.bairro + ' ' + it.referencia).toLowerCase();
    return /balsa|itaquacetuba|Bororé|Borore|curucutu|Curucutu|Tatetos|Agua limpa|agua limpa|Jardim Santa Tereza| Jardim Borba Gato|tatetos|taquacetuba|pos balsa|pós balsa|curucutu|capivari|santa cruz/i.test(txt);
  }

  function populateTechFilter() {
    const techs = [...new Set(items.map(i => i.tecnico).filter(t => t && t.length > 2))].sort();
    techFilterEl.innerHTML = '<option value="">Todos os Técnicos</option>';
    techs.forEach(t => {
        const opt = document.createElement('option');
        opt.value = t; opt.textContent = t; techFilterEl.appendChild(opt);
    });
  }

  function createCard(it) {
    const clone = cardTemplate.content.cloneNode(true);
    const card = clone.querySelector('.result-card');
    const fullAddr = buildAddress(it);
    const isBalsa = checkBalsa(it);
    const hasAddr = fullAddr.length > 8 && !/não informado|não localizado/i.test(fullAddr);
    const badgeBox = card.querySelector('.card-badges');
    if(isBalsa) {
        card.classList.add('is-balsa');
        badgeBox.innerHTML += `<span class="special-badge badge-balsa">⚠️ Região Balsa</span>`;
    }
    if(!hasAddr) {
        card.classList.add('is-warning');
        badgeBox.innerHTML += `<span class="special-badge badge-warning">🟧 Sem Endereço</span>`;
    }
    card.querySelector('.card-client').textContent = it.cliente || 'Cliente Desconhecido';
    card.querySelector('.card-subject').textContent = it.assunto || '-';
    card.querySelector('.card-branch').textContent = it.filial || 'Matriz';
    card.querySelector('.card-tech').textContent = it.tecnico || 'Sem Técnico';
    card.querySelector('.card-address').textContent = hasAddr ? fullAddr : 'Endereço não identificado';
    const mapsLink = hasAddr ? getGoogleMapsLink(fullAddr || it.cidade) : '#';
    const btnMaps = card.querySelector('.open-maps');
    btnMaps.href = mapsLink;
    if(!hasAddr) btnMaps.style.opacity = 0.5;
    const btnWhatsapp = card.querySelector('.open-zap');
    const whatsappLink = getWhatsAppLink(it.telefone);
    if (whatsappLink !== '#') {
        btnWhatsapp.href = whatsappLink;
    } else {
        btnWhatsapp.style.opacity = 0.4;
        btnWhatsapp.style.pointerEvents = 'none';
    }
    const btnCreds = card.querySelector('.copy-creds');
    btnCreds.onclick = (e) => {
        e.stopPropagation();
        const creds = formatCredentials(it);
        navigator.clipboard.writeText(creds).then(() => {
            const original = btnCreds.innerHTML;
            btnCreds.innerHTML = '<span style="font-size:10px">OK</span>';
            setTimeout(() => btnCreds.innerHTML = original, 1500);
        });
    };
    const btnCopy = card.querySelector('.copy-btn');
    btnCopy.onclick = (e) => {
        e.stopPropagation();
        if(!hasAddr) return alert('Sem endereço válido.');
        navigator.clipboard.writeText(mapsLink).then(() => {
            btnCopy.textContent = 'Copiado!';
            setTimeout(() => btnCopy.textContent = 'Link', 2000);
        });
    };
    card.querySelector('.details-btn').onclick = (e) => {
        e.stopPropagation();
        openModal(it, fullAddr, mapsLink, isBalsa);
    };
    card.onclick = (e) => {
        if(!['A','BUTTON','svg','path'].includes(e.target.tagName))
            openModal(it, fullAddr, mapsLink, isBalsa);
    };
    btnWhatsapp.onclick = (e) => {
        e.preventDefault();
        e.stopPropagation();
        if (!it.telefone) return;
        const hour = new Date().getHours();
        const greeting = hour < 12 ? 'Bom dia' : 'Boa tarde';
        let genderHint = 'Sr./Sra.';
        const name = (it.cliente || '').trim();
        if (/\b(sra|senhora|srt[aã])\b/i.test(name)) genderHint = 'Sra.';
        else if (/\b(sr|senhor)\b/i.test(name)) genderHint = 'Sr.';
        const templates = [
          `${greeting}, tudo bem? Falo com ${genderHint} ${name}? Temos uma ordem de serviço agendada para a data de hoje! Podemos confirmar?`,
          `${greeting}, o técnico está tentando localizar a residência, poderia por gentileza encaminhar a localização em tempo real?`
        ];
        const choice = prompt(`1 - Confirmar agendamento\n2 - Pedir localização\n\nDigite 1 ou 2:`,'1');
        const idx = choice === '2' ? 1 : 0;
        const text = templates[idx];
        const url = getWhatsAppLink(it.telefone, text);
        if (url === '#') return alert('Número inválido.');
        window.open(url, '_blank');
    };
    return clone;
  }

  function render() {
    const term = (filterEl.value || '').toLowerCase();
    const tech = techFilterEl.value;
    let list = items.filter(it => {
        const txt = ((it.cliente || '') + ' ' + (it.endereco || '') + ' ' + (it.assunto || '') + ' ' + (it.login || '')).toLowerCase();
        return txt.includes(term);
    });
    if(tech) list = list.filter(it => it.tecnico === tech);
    if(activeCategory === 'balsa') list = list.filter(it => checkBalsa(it));
    else if(activeCategory === 'noaddr') list = list.filter(it => !buildAddress(it) || buildAddress(it).length < 10);
    filteredItems = list;
    if(rendered === 0) resultsEl.innerHTML = '';
    if(list.length === 0) {
        resultsEl.innerHTML = '<div style="grid-column:1/-1;text-align:center;padding:40px;opacity:0.6">Nenhum resultado encontrado.</div>';
        moreBtn.classList.add('hidden');
        return;
    }
    const batch = list.slice(rendered, rendered + PER_PAGE);
    const frag = document.createDocumentFragment();
    batch.forEach(it => frag.appendChild(createCard(it)));
    resultsEl.appendChild(frag);
    rendered += batch.length;
    moreBtn.classList.toggle('hidden', rendered >= list.length);
  }

  function openModal(it, addr, mapsLink, isBalsa) {
    modalBody.innerHTML = '';
    const clone = modalTemplate.content.cloneNode(true);
    clone.querySelector('.modal-client').textContent = it.cliente;
    clone.querySelector('.modal-branch').textContent = it.filial;
    clone.querySelector('.modal-tech').textContent = it.tecnico;
    clone.querySelector('.modal-login').textContent = `${it.login || '-'} / ${it.senha || '-'}`;
    clone.querySelector('.modal-phone').textContent = it.telefone || '-';
    clone.querySelector('.modal-subject').textContent = it.assunto;
    clone.querySelector('.modal-address').textContent = addr || 'Não informado';
    clone.querySelector('.modal-complement').textContent = it.complemento || '-';
    const districtEl = clone.querySelector('.modal-district');
    if (districtEl) districtEl.textContent = it.bairro || '-';
    clone.querySelector('.modal-ref').textContent = it.referencia || '-';
    const mapContainer = clone.querySelector('#mapContainer');
    const q = encodeURIComponent(addr || it.cidade || '');
    const gType = isBalsa ? 'k' : 'm';
    const googleUrl = `https://maps.google.com/maps?q=${q}&t=${gType}&z=15&ie=UTF8&iwloc=&output=embed`;
    const bingUrl = `https://www.bing.com/maps/embed?h=350&w=500&lvl=15&typ=d&sty=r&src=SHELL&FORM=MBEDV8&q=${q}&mkt=pt-BR`;
    if (mapContainer) mapContainer.innerHTML = `<iframe class="map-iframe" src="${googleUrl}" loading="lazy"></iframe>`;
    const btns = clone.querySelectorAll('.provider-btn');
    btns.forEach(b => {
        b.onclick = () => {
            btns.forEach(x => x.classList.remove('active'));
            b.classList.add('active');
            const prov = b.dataset.provider;
            if(prov === 'waze') window.open(`https://waze.com/ul?q=${q}`, '_blank');
            else if(prov === 'bing') if (mapContainer) mapContainer.innerHTML = `<iframe class="map-iframe" src="${bingUrl}"></iframe>`;
            else if (mapContainer) mapContainer.innerHTML = `<iframe class="map-iframe" src="${googleUrl}"></iframe>`;
        };
    });
    const copyCredsBtn = clone.querySelector('.copy-creds-btn');
    if (copyCredsBtn) {
      copyCredsBtn.onclick = (e) => {
        navigator.clipboard.writeText(formatCredentials(it));
        e.target.textContent = 'Copiado!'; setTimeout(()=>e.target.textContent='Copiar Dados', 2000);
      };
    }
    const copyLinkBtn = clone.querySelector('.copy-link-btn');
    if (copyLinkBtn) {
      copyLinkBtn.onclick = (e) => {
        navigator.clipboard.writeText(mapsLink);
        e.target.textContent = 'Copiado!'; setTimeout(()=>e.target.textContent='Copiar Link', 2000);
      };
    }
    modalBody.appendChild(clone);
    overlay.classList.remove('hidden');
    document.body.style.overflow = 'hidden';
  }

  copyLinkListBtn.onclick = () => {
    if(!filteredItems.length) return alert('Nada listado.');
    const text = filteredItems.map((it, i) => {
        const addr = buildAddress(it);
        const link = (addr && addr.length > 8) ? getGoogleMapsLink(addr) : 'Sem endereço válido';
        return `*${i+1}. ${it.cliente}*\n🔗 ${link}`;
    }).join('\n\n');
    navigator.clipboard.writeText(text).then(() => alert('Lista copiada!'));
  };

  openRouteBtn.onclick = () => {
    if(!filteredItems.length) return alert('Nada listado.');
    const validItems = filteredItems.filter(it => buildAddress(it).length > 10);
    if(!validItems.length) return alert('Sem endereços válidos.');
    if(validItems.length > 10 && !confirm(`Abrir os primeiros 10 de ${validItems.length}?`)) return;
    const dests = validItems.slice(0, 10).map(it => encodeURIComponent(buildAddress(it))).join('/');
    window.open(`https://www.google.com/maps/dir/${dests}`, '_blank');
  };

  exportBtn.onclick = async () => {
    if(!filteredItems.length) return alert('Nada para exportar.');
    if(!window.XLSX) {
      try { await loadSheetJs(); } catch(err) { alert('Erro ao carregar biblioteca de exportação.'); return; }
    }
    const rows = filteredItems.map(it => ({
        Filial: it.filial, Tecnico: it.tecnico, Cliente: it.cliente, Assunto: it.assunto,
        Endereco: buildAddress(it), Bairro: it.bairro, Login: it.login, Senha: it.senha,
        Telefone: it.telefone, Link: getGoogleMapsLink(buildAddress(it))
    }));
    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Relatorio");
    XLSX.writeFile(wb, "Relatorio_Athon.xlsx");
  };

  loadBtn.onclick = async (e) => {
    e.preventDefault(); e.stopPropagation();
    if(!selectedFile) return;
    loadBtn.disabled = true; loadBtn.textContent = 'Lendo...';
    showProgress(0.3); setStatus('Lendo arquivo...');
    resultsEl.innerHTML = ''; rendered = 0;
    try {
        let rawRows = [];
        if(selectedFile.name.toLowerCase().endsWith('.csv')) {
            const text = await readAsTextDetectEncoding(selectedFile);
            rawRows = parseCSVManual(text);
        } else {
            const XLSX = await loadSheetJs();
            const buffer = await selectedFile.arrayBuffer();
            const wb = XLSX.read(buffer, {type:'array'});
            const ws = wb.Sheets[wb.SheetNames[0]];
            rawRows = XLSX.utils.sheet_to_json(ws, {header:1});
        }
        items = mapData(rawRows);
        if(items.length === 0) throw new Error("Nenhum dado encontrado.");
        populateTechFilter();
        showProgress(1); setStatus(`${items.length} registros.`);
        render();
    } catch(err) {
        console.error(err);
        alert(`Erro ao ler: ${err.message}`);
        setStatus('Erro na leitura.');
    } finally {
        loadBtn.disabled = false; loadBtn.textContent = 'CARREGAR';
        setTimeout(() => showProgress(0), 1000);
    }
  };

  fileInput.onchange = (e) => { if(e.target.files.length) { selectedFile = e.target.files[0]; updateFileDisplay(selectedFile); }};
  dropZone.ondragover = (e) => { e.preventDefault(); dropZone.classList.add('dragover'); };
  dropZone.ondragleave = () => dropZone.classList.remove('dragover');
  dropZone.ondrop = (e) => { e.preventDefault(); dropZone.classList.remove('dragover'); if(e.dataTransfer.files.length) { selectedFile=e.dataTransfer.files[0]; updateFileDisplay(selectedFile); }};

  let deb;
  filterEl.oninput = () => { clearTimeout(deb); deb = setTimeout(() => { rendered=0; render(); }, 300); };
  techFilterEl.onchange = () => { rendered=0; render(); };

  filterChips.forEach(btn => btn.onclick = () => {
    filterChips.forEach(b => b.classList.remove('active')); btn.classList.add('active');
    activeCategory = btn.dataset.filter; rendered=0; render();
  });

  moreBtn.onclick = render;

  closeModal.onclick = () => {
    overlay.classList.add('hidden');
    document.body.style.overflow = '';
  };
  overlay.onclick = (e) => { if(e.target===overlay) { overlay.classList.add('hidden'); document.body.style.overflow = ''; } };

  function updateThemeIcon() {
    const isDark = document.body.classList.contains('dark-mode');
    themeBtn.innerHTML = isDark 
      ? `<svg class="icon-theme" viewBox="0 0 24 24"><path d="M12 3c-4.97 0-9 4.03-9 9s4.03 9 9 9 9-4.03 9-9c0-.46-.04-.92-.1-1.36-.98 1.37-2.58 2.26-4.4 2.26-2.98 0-5.4-2.42-5.4-5.4 0-1.81.89-3.42 2.26-4.4-.44-.06-.9-.1-1.36-.1z"/></svg>`
      : `<svg class="icon-theme" viewBox="0 0 24 24"><path d="M6.76 4.84l-1.8-1.79-1.41 1.41 1.79 1.79 1.42-1.41zM4 10.5H1v2h3v-2zm9-9.95h-2V3.5h2V.55zm7.45 3.91l-1.41-1.41-1.79 1.79 1.41 1.41 1.79-1.79zm-3.21 13.7l1.79 1.8 1.41-1.41-1.8-1.79-1.4 1.4zM20 10.5v2h3v-2h-3zm-8-5c-3.31 0-6 2.69-6 6s2.69 6 6 6 6-2.69 6-6-2.69-6-6-6zm-1 16.95h2V19.5h-2v2.95zm-7.45-3.91l1.41 1.41 1.79-1.8-1.41-1.41-1.79 1.8z"/></svg>`;
  }

  if(localStorage.getItem('theme') === 'light') { document.body.classList.remove('dark-mode'); }
  updateThemeIcon();

  themeBtn.onclick = () => {
    document.body.classList.toggle('dark-mode');
    localStorage.setItem('theme', document.body.classList.contains('dark-mode') ? 'dark' : 'light');
    updateThemeIcon();
  };

  updateFileDisplay(null);

  window.addEventListener('resize', () => {
    const prev = PER_PAGE;
    PER_PAGE = window.innerWidth < 600 ? 6 : window.innerWidth < 900 ? 12 : 50;
    if (PER_PAGE !== prev) { rendered = 0; resultsEl.innerHTML = ''; render(); }
  });
});




