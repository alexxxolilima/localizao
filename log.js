document.addEventListener('DOMContentLoaded', () => {
    'use strict';

    const els = {
        fileInput: document.getElementById('fileInput'),
        dropZone: document.getElementById('dropZone'),
        loadBtn: document.getElementById('loadBtn'),
        fileDisplay: document.getElementById('fileDisplayName'),
        fileStatus: document.getElementById('fileStatus'),
        results: document.getElementById('results'),
        filter: document.getElementById('filter'),
        techFilter: document.getElementById('techFilter'),
        regionFilter: document.getElementById('regionFilter'),
        moreBtn: document.getElementById('moreBtn'),
        toolbar: document.getElementById('toolbarSection'),
        loader: document.getElementById('globalLoader'),
        loaderBar: document.getElementById('loaderBar'),
        toastContainer: document.getElementById('toastContainer'),
        overlay: document.getElementById('overlay'),
        modalBody: document.getElementById('modalBody'),
        closeModal: document.getElementById('closeModal'),
        templatesModal: document.getElementById('templatesModal'),
        numberEntryModal: document.getElementById('numberEntryModal'),
        inlinePhoneInput: document.getElementById('inlinePhoneInput'),
        savePhoneBtn: document.getElementById('savePhoneBtn'),
        manualPhoneTriggerBtn: document.getElementById('manualPhoneTriggerBtn'),
        themeBtn: document.getElementById('themeBtn'),
        copyListBtn: document.getElementById('copyLinkListBtn'),
        openRouteBtn: document.getElementById('openRouteBtn'),
        statsDashboard: document.getElementById('statsDashboard'),
        statTotal: document.getElementById('totalCount'),
        statBalsa: document.getElementById('balsaCount'),
        statDone: document.getElementById('doneCount')
    };

    let state = {
        items: [],
        filtered: [],
        renderedCount: 0,
        currentFile: null,
        activeFilter: 'all',
        itemForTemplate: null,
        PAGE_SIZE: 50
    };

    const REGEX = {
        BALSA: /balsa|itaquacetuba|boror[ée]|curucutu|tatetos|agua limpa|jardim santa tereza|jardim borba gato|p[oó]s balsa|capivari|santa cruz/i,
        RIACHO: /fincos|tup[ãa]|rio grande|boa vista|capelinha|cocaia|zanzala|vila pele|riacho grande|areião|jussara|balne[áa]ria/i,
        BLOCKED: /\b(retirada|reagend|tentativa|^re$)\b/i
    };

    function showToast(msg, type = 'info') {
        const div = document.createElement('div');
        div.className = 'toast';
        div.innerHTML = `<span>${msg}</span>`;
        div.style.borderLeftColor = type === 'error' ? 'var(--danger)' : type === 'success' ? 'var(--success)' : 'var(--primary)';
        els.toastContainer.appendChild(div);
        setTimeout(() => {
            div.style.opacity = '0';
            div.style.transform = 'translateX(100%)';
            setTimeout(() => div.remove(), 300);
        }, 3000);
    }

    function setLoader(percent, text) {
        if (percent === 0) {
            els.loader.classList.add('hidden');
            return;
        }
        els.loader.classList.remove('hidden');
        els.loaderBar.style.width = `${percent}%`;
        if (text) document.querySelector('.loader-text').textContent = text;
    }

    function normalizeKey(key) {
        return String(key).toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, "").replace(/[^a-z0-9]/g, "");
    }

    function mapData(rawJson) {
        if (!rawJson || !rawJson.length) return [];

        const headers = Object.keys(rawJson[0]);
        const mapKeys = {};

        headers.forEach(h => {
            const norm = normalizeKey(h);
            if (norm.includes('cliente') || norm.includes('nome')) mapKeys.cliente = h;
            else if (norm.includes('filial')) mapKeys.filial = h;
            else if (norm.includes('assunto') || norm.includes('motivo')) mapKeys.assunto = h;
            else if (norm.includes('tecnico') || norm.includes('colaborador')) mapKeys.tecnico = h;
            else if (norm.includes('endereco') || norm.includes('logradouro')) mapKeys.endereco = h;
            else if (norm.includes('numero')) mapKeys.numero = h;
            else if (norm.includes('complemento')) mapKeys.complemento = h;
            else if (norm.includes('bairro')) mapKeys.bairro = h;
            else if (norm.includes('cidade')) mapKeys.cidade = h;
            else if (norm.includes('referencia') || norm.includes('obs')) mapKeys.referencia = h;
            else if (norm.includes('login') || norm.includes('usuario')) mapKeys.login = h;
            else if (norm.includes('senha') || norm.includes('password') || norm.includes('md5')) mapKeys.senha = h;
            else if (norm.includes('telefone') || norm.includes('celular') || norm.includes('whatsapp')) {
                if (!mapKeys.telefone) mapKeys.telefone = h;
            }
        });

        return rawJson.map(row => {
            let tels = [];
            headers.forEach(h => {
                if (normalizeKey(h).includes('telef') || normalizeKey(h).includes('whats') || normalizeKey(h).includes('contato')) {
                    if (row[h]) tels.push(row[h]);
                }
            });

            const item = {
                cliente: row[mapKeys.cliente] || 'Cliente Desconhecido',
                filial: row[mapKeys.filial] || '',
                assunto: row[mapKeys.assunto] || '',
                tecnico: row[mapKeys.tecnico] || '',
                endereco: row[mapKeys.endereco] || '',
                numero: row[mapKeys.numero] || '',
                complemento: row[mapKeys.complemento] || '',
                bairro: row[mapKeys.bairro] || '',
                cidade: row[mapKeys.cidade] || '',
                referencia: row[mapKeys.referencia] || '',
                login: row[mapKeys.login] || '',
                senha: row[mapKeys.senha] || '',
                telefone: tels.join(' / '),
                raw: row,
                done: false
            };
            
            const fullAddr = (item.endereco + ' ' + item.bairro + ' ' + item.cidade).toLowerCase();
            const cep = (String(row['CEP'] || row['cep'] || '')).replace(/\D/g, '').substring(0,3);

            if (cep === '078' || fullAddr.includes('franco')) item.region = '078';
            else if (cep === '048' || fullAddr.includes('graja')) item.region = '048';
            else if (cep === '041' || fullAddr.includes('moraes') || fullAddr.includes('mv') || fullAddr.includes('mundo virtua')) item.region = '041';
            else if (cep === '098' || fullAddr.includes('bernardo')) {
                if (REGEX.BALSA.test(fullAddr)) item.region = '098-balsa';
                else if (REGEX.RIACHO.test(fullAddr)) item.region = '098-riacho';
                else item.region = '098-sbc';
            } else {
                item.region = 'other';
            }
            
            return item;
        }).filter(i => !REGEX.BLOCKED.test(i.assunto));
    }

    function buildAddress(it) {
        let parts = [];
        if (it.endereco) parts.push(it.endereco);
        if (it.numero) parts.push(it.numero);
        if (it.complemento) parts.push(it.complemento);
        if (it.bairro) parts.push(it.bairro);
        if (it.cidade) parts.push(it.cidade);
        return parts.filter(p => p && String(p).trim().length > 0).join(', ');
    }

    function createCard(it) {
        const template = document.getElementById('cardTemplate');
        const clone = template.content.cloneNode(true);
        const card = clone.querySelector('.result-card');
        
        const fullAddr = buildAddress(it);
        const hasAddr = fullAddr.length > 10 && !/n[ãa]o (informado|consta)/i.test(fullAddr);
        
        card.classList.add(`region-${it.region}`);
        if (!hasAddr) card.classList.add('no-addr');
        if (it.done) card.classList.add('is-done');

        card.querySelector('.card-client').textContent = it.cliente;
        card.querySelector('.card-subject').textContent = it.assunto;
        card.querySelector('.card-tech').textContent = it.tecnico || 'Sem Técnico';
        card.querySelector('.card-address').textContent = hasAddr ? fullAddr : 'Endereço não identificado';
        
        const badgeBox = card.querySelector('.card-badges');
        if (it.region === '098-balsa') badgeBox.innerHTML += `<span class="badge-count" style="background:var(--warning);color:#000">⚠️ Balsa</span>`;
        if (!hasAddr) badgeBox.innerHTML += `<span class="badge-count" style="background:var(--danger);color:#fff">Sem Endereço</span>`;

        const mapsLink = hasAddr ? `https://www.google.com/maps/search/?api=1&query=${encodeURIComponent(fullAddr)}` : '#';
        
        const btnMaps = card.querySelector('.open-maps');
        btnMaps.href = mapsLink;
        if (!hasAddr) btnMaps.style.opacity = '0.5';

        const btnZap = card.querySelector('.open-zap');
        btnZap.onclick = (e) => {
            e.stopPropagation();
            state.itemForTemplate = it;
            els.templatesModal.classList.remove('hidden');
        };

        const btnCreds = card.querySelector('.copy-creds');
        btnCreds.onclick = (e) => {
            e.stopPropagation();
            const txt = `${it.login}\n \n${it.senha}\n \nVlan 500`;
            navigator.clipboard.writeText(txt);
            showToast('Login copiado!', 'success');
        };
        
        const btnDetails = card.querySelector('.details-btn');
        btnDetails.onclick = (e) => { e.stopPropagation(); openDetails(it); };
        card.onclick = (e) => { if(e.target === card || e.target.closest('.card-main-info')) openDetails(it); };

        const btnCheck = document.createElement('button');
        btnCheck.className = 'action-btn btn-check';
        btnCheck.title = "Marcar como Concluído";
        btnCheck.innerHTML = '<span style="color: var(--success); font-weight: 800;">✔</span>';
        btnCheck.onclick = (e) => {
            e.stopPropagation();
            it.done = !it.done;
            card.classList.toggle('is-done');
            updateCounters();
        };
        card.querySelector('.card-actions').prepend(btnCheck);

        return clone;
    }

    function render() {
        const { items, filtered, renderedCount, PAGE_SIZE } = state;
        const container = els.results;
        
        if (renderedCount === 0) container.innerHTML = '';

        if (filtered.length === 0) {
            container.innerHTML = '<div style="grid-column:1/-1;text-align:center;padding:40px;color:var(--text-secondary)">Nenhum resultado encontrado.</div>';
            els.moreBtn.classList.add('hidden');
            return;
        }

        const nextBatch = filtered.slice(renderedCount, renderedCount + PAGE_SIZE);
        const frag = document.createDocumentFragment();
        
        nextBatch.forEach(it => frag.appendChild(createCard(it)));
        container.appendChild(frag);

        state.renderedCount += nextBatch.length;
        
        if (state.renderedCount >= filtered.length) els.moreBtn.classList.add('hidden');
        else els.moreBtn.classList.remove('hidden');

        updateCounters();
    }

    function filterData() {
        const term = els.filter.value.toLowerCase();
        const tech = els.techFilter.value;
        const region = els.regionFilter.value;
        const quick = state.activeFilter;

        state.filtered = state.items.filter(it => {
            const matchText = (it.cliente + ' ' + it.endereco + ' ' + it.login + ' ' + it.assunto).toLowerCase().includes(term);
            const matchTech = !tech || it.tecnico === tech;
            const matchRegion = region === 'all' || it.region === region;
            let matchQuick = true;
            if (quick === 'balsa') matchQuick = it.region === '098-balsa';
            if (quick === 'noaddr') matchQuick = buildAddress(it).length < 10;

            return matchText && matchTech && matchRegion && matchQuick;
        });

        state.renderedCount = 0;
        render();
    }

    function updateCounters() {
        const balsaCount = state.items.filter(it => it.region === '098-balsa').length;
        const doneCount = state.items.filter(it => it.done).length;

        const badge = document.querySelector('.badge-count[data-region="098-balsa"]');
        if (badge) {
            badge.textContent = balsaCount;
            badge.classList.toggle('hidden', balsaCount === 0);
        }

        if (els.statTotal) els.statTotal.textContent = state.items.length;
        if (els.statBalsa) els.statBalsa.textContent = balsaCount;
        if (els.statDone) els.statDone.textContent = doneCount;

        if (els.statsDashboard && state.items.length > 0) {
            els.statsDashboard.classList.remove('hidden');
        }
    }

    function populateSelects() {
        const techs = [...new Set(state.items.map(i => i.tecnico).filter(Boolean))].sort();
        els.techFilter.innerHTML = '<option value="">Todos os Técnicos</option>';
        techs.forEach(t => {
            const opt = document.createElement('option');
            opt.value = t;
            opt.textContent = t;
            els.techFilter.appendChild(opt);
        });
    }

    function openDetails(it) {
        const tpl = document.getElementById('modalTemplate').content.cloneNode(true);
        const addr = buildAddress(it);
        const mapQ = encodeURIComponent(addr || it.cidade);

        tpl.querySelector('.modal-client').textContent = it.cliente;
        tpl.querySelector('.modal-branch').textContent = it.filial;
        tpl.querySelector('.modal-tech').textContent = it.tecnico;
        tpl.querySelector('.modal-login').textContent = `${it.login || '-'} / ${it.senha || '-'}`;
        tpl.querySelector('.modal-phone').textContent = it.telefone || '-';
        tpl.querySelector('.modal-subject').textContent = it.assunto;
        tpl.querySelector('.modal-address').textContent = addr || 'Endereço não cadastrado';
        tpl.querySelector('.modal-ref').textContent = it.referencia || '-';

        const mapContainer = tpl.getElementById('mapContainer');
        const gUrl = `https://maps.google.com/maps?q=${mapQ}&t=m&z=15&output=embed&iwloc=near`;
        mapContainer.innerHTML = `<iframe class="map-iframe" src="${gUrl}" loading="lazy"></iframe>`;

        const btns = tpl.querySelectorAll('.provider-btn');
        btns.forEach(b => {
            b.onclick = () => {
                btns.forEach(x => x.classList.remove('active'));
                b.classList.add('active');
                if(b.dataset.provider === 'waze') {
                    window.open(`https://waze.com/ul?q=${mapQ}`, '_blank');
                } else {
                    mapContainer.innerHTML = `<iframe class="map-iframe" src="${gUrl}"></iframe>`;
                }
            };
        });

        tpl.querySelector('.copy-creds-btn').onclick = (e) => {
            const txt = `${it.login}\n \n${it.senha}\n \nVlan 500`;
            navigator.clipboard.writeText(txt);
            e.target.textContent = 'Copiado!';
            setTimeout(() => e.target.textContent = 'Copiar', 1500);
        };

        els.modalBody.innerHTML = '';
        els.modalBody.appendChild(tpl);
        els.overlay.classList.remove('hidden');
    }

    els.loadBtn.onclick = () => {
        if (!state.currentFile) return;
        setLoader(30, 'Lendo Arquivo...');
        
        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                const jsonData = XLSX.utils.sheet_to_json(firstSheet);

                if (jsonData.length === 0) throw new Error("Planilha vazia");

                state.items = mapData(jsonData);
                
                populateSelects();
                setLoader(100, 'Finalizando...');
                setTimeout(() => {
                    setLoader(0);
                    els.toolbar.classList.remove('hidden');
                    filterData();
                    showToast(`${state.items.length} O.S. carregadas!`, 'success');
                }, 500);

            } catch (err) {
                console.error(err);
                setLoader(0);
                showToast("Erro ao ler arquivo. Verifique o formato.", 'error');
            }
        };
        reader.readAsArrayBuffer(state.currentFile);
    };

    els.fileInput.onchange = (e) => {
        if (e.target.files.length) {
            state.currentFile = e.target.files[0];
            els.fileDisplay.textContent = state.currentFile.name;
            els.loadBtn.disabled = false;
            els.fileStatus.textContent = 'Pronto para carregar';
        }
    };

    els.dropZone.ondragover = (e) => { e.preventDefault(); els.dropZone.classList.add('dragover'); };
    els.dropZone.ondragleave = () => els.dropZone.classList.remove('dragover');
    els.dropZone.ondrop = (e) => {
        e.preventDefault();
        els.dropZone.classList.remove('dragover');
        if (e.dataTransfer.files.length) {
            state.currentFile = e.dataTransfer.files[0];
            els.fileDisplay.textContent = state.currentFile.name;
            els.loadBtn.disabled = false;
        }
    };

    els.filter.oninput = () => { clearTimeout(window.deb); window.deb = setTimeout(filterData, 300); };
    els.techFilter.onchange = filterData;
    els.regionFilter.onchange = filterData;
    
    document.querySelectorAll('.chip').forEach(chip => {
        chip.onclick = () => {
            document.querySelectorAll('.chip').forEach(c => c.classList.remove('active'));
            chip.classList.add('active');
            state.activeFilter = chip.dataset.filter;
            filterData();
        };
    });

    els.moreBtn.onclick = render;

    els.manualPhoneTriggerBtn.onclick = () => {
        els.templatesModal.classList.add('hidden');
        els.inlinePhoneInput.value = '';
        els.numberEntryModal.classList.remove('hidden');
    };

    document.querySelectorAll('.template-option').forEach(btn => {
        btn.onclick = () => {
            const it = state.itemForTemplate;
            let text = btn.dataset.template;
            const saudacao = new Date().getHours() < 12 ? 'Bom dia' : 'Boa tarde';
            const nome = (it.cliente || '').split(' ')[0];
            
            text = text.replace('{{saudacao}}', saudacao)
                       .replace('{{nome}}', nome)
                       .replace('{{tratamento}}', 'Sr(a).');

            const rawPhone = it.telefone ? it.telefone.split('/')[0] : '';
            const phone = rawPhone.replace(/\D/g, '');

            if (phone.length >= 10) {
                window.open(`https://wa.me/55${phone}?text=${encodeURIComponent(text)}`, '_blank');
                els.templatesModal.classList.add('hidden');
            } else {
                els.templatesModal.classList.add('hidden');
                els.inlinePhoneInput.value = '';
                els.numberEntryModal.classList.remove('hidden');
            }
        };
    });

    els.savePhoneBtn.onclick = () => {
        const num = els.inlinePhoneInput.value.replace(/\D/g, '');
        if (num.length < 10) return showToast('Número inválido', 'error');
        state.itemForTemplate.telefone = num;
        els.numberEntryModal.classList.add('hidden');
        els.templatesModal.classList.remove('hidden');
    };

    els.closeModal.onclick = () => els.overlay.classList.add('hidden');
    els.overlay.onclick = (e) => { if(e.target === els.overlay) els.overlay.classList.add('hidden'); };

    els.themeBtn.onclick = () => {
        document.body.classList.toggle('dark-mode');
    };
    
    els.copyListBtn.onclick = () => {
        if (!state.filtered.length) return;
        const txt = state.filtered.slice(0, 50).map(i => `*${i.cliente}*\n${buildAddress(i)}`).join('\n\n');
        navigator.clipboard.writeText(txt);
        showToast('Lista copiada!', 'success');
    };

    els.openRouteBtn.onclick = () => {
        if (!state.filtered.length) return;
        const validItems = state.filtered.filter(it => buildAddress(it).length > 10).slice(0, 10);
        if (validItems.length === 0) return showToast('Sem endereços válidos', 'error');
        
        const dests = validItems.map(it => encodeURIComponent(buildAddress(it))).join('/');
        window.open(`https://www.google.com/maps/dir//${dests}`, '_blank');
    };
});




