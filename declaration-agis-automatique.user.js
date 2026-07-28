// ==UserScript==
// @name         LTOA - Agent Déclaration AGIS (Lot visible)
// @namespace    https://ltoa-assurances.fr/
// @version      1.1.3
// @description  Traite les contrats AGIS, contrôle chaque preuve de paiement dans toute la GED, puis génère une déclaration Excel financièrement sécurisée et un JSON auditable.
// @author       LTOA Assurances
// @match        https://courtage.modulr.fr/*
// @require      https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js
// @require      https://cdn.jsdelivr.net/npm/tesseract.js@5.1.1/dist/tesseract.min.js
// @require      https://cdn.jsdelivr.net/npm/exceljs@4.4.0/dist/exceljs.min.js
// @grant        GM_xmlhttpRequest
// @connect      courtage.modulr.fr
// @run-at       document-idle
// @updateURL    https://raw.githubusercontent.com/BiggerThanTheMall/declaration-agis/main/declaration-agis-automatique.user.js
// @downloadURL  https://raw.githubusercontent.com/BiggerThanTheMall/declaration-agis/main/declaration-agis-automatique.user.js
// ==/UserScript==

(function () {
    'use strict';

    const APP_ID = 'ltoa-agent-declaration-agis';
    const CURRENT_VERSION = '1.1.2';
    const TESSERACT_OCR_OPTIONS = Object.freeze({
        workerPath: 'https://cdn.jsdelivr.net/npm/tesseract.js@5.1.1/dist/worker.min.js',
        corePath: 'https://cdn.jsdelivr.net/npm/tesseract.js-core@5.1.1',
        langPath: 'https://cdn.jsdelivr.net/npm/@tesseract.js-data/fra@1.0.0/4.0.0_best_int',
    });
    const STORAGE_KEY = 'ltoa.agentDeclarationAgis.v0.1';
    const ACTIVE_MONTH_KEY = 'ltoa.agentDeclarationAgis.activeMonth.v1';
    const MONTH_STATE_PREFIX = 'ltoa.agentDeclarationAgis.month.v1.';
    const START_REQUEST_KEY = 'ltoa.agentDeclarationAgis.startRequest';
    const LIST_PATH = '/fr/scripts/policies/policies_list.php';
    const CLIENT_PATH = '/fr/scripts/clients/clients_card.php';
    const AGIS_COMPANY_PATH = '/fr/scripts/companies/companies_card.php';
    const AGIS_COMPANY_ID = '50';
    const AGIS_COMPANY_URL = `${location.origin}${AGIS_COMPANY_PATH}?company_id=${AGIS_COMPANY_ID}#entity_menu_company=0`;
    const GED_PREFIX = '/fr/intranet/edm/display/Client/';
    const PAYMENT_WORDS = Object.freeze([
        'justificatif', 'justif', 'paiement', 'reglement', 'cheque', 'virement',
        'stripe', 'paypal', 'ppal', 'recu', 'facture', 'acquittee', 'encaissement',
        'encaisse', 'especes', 'carte', 'cb',
    ]);
    const PAYMENT_TITLE_NEGATIVE = /\b(perdu(?:e)?|perte|refus[eé]?|rejet[eé]?|annul[eé]?|rembours[eé]?|impay[eé]?|opposition|fraude|erreur|sans\s+suite|non[\s-]+encaiss[eé]?|non[\s-]+re[cç]u|en\s+attente|[àa]\s+encaisser)\b/i;
    const PAYMENT_CONTENT_NEGATIVE = /\b(?:(?:[ée]tat|statut|status)(?:\s+de\s+la\s+transaction)?\s*:?\s*(?:en\s+attente|pending|processing|programm[eé]?|planifi[eé]?|brouillon|draft|initiated|refus[eé]?|failed|declined|rejet[eé]?|annul[eé]?|cancelled|canceled|rembours[eé]?|refunded|impay[eé]?|chargeback)|(?:paiement|transaction|virement|ch[eè]que)\s+(?:est\s+)?(?:en\s+attente|refus[eé]?|rejet[eé]?|annul[eé]?|rembours[eé]?|impay[eé]?|non[\s-]+ex[eé]cut[eé]?|non[\s-]+encaiss[eé]?))\b/i;
    const PAYMENT_FINAL_POSITIVE = /\b(termin[eé]e?|completed|r[eé]ussi(?:e)?|succeeded|pay[eé]e?|paid|re[cç]u|received|effectu[eé]e?|ex[eé]cut[eé]e?|cr[eé]dit[eé]e?|encaiss[eé]e?|valid[eé]e?|confirm[eé]e?|settled|captur[eé]e?)\b/i;
    const QUOTE_WORDS = /\b(devis|proposition|adh[eé]sion)\b/i;
    const COVERAGE_WORDS = /\b(devis|proposition|adh[eé]sion|contrat|conditions? particuli[eè]res?|bulletin|attestation|renouvellement|sign[eé])\b/i;
    const AGIS_CP_TEMPLATE = Object.freeze({
        subscriberHeading: 'INFORMATIONS DU SOUSCRIPTEUR',
        coveredPersonsHeading: 'INFORMATION DES ASSURES / BENEFICIAIRES',
        contractTypeLabel: 'Type de contrat',
        destinationCountryLabel: 'Pays désigné en cas de rapatriement',
        guaranteePeriodLabel: 'Période de garantie',
    });

    if (document.getElementById(APP_ID)) return;

    function clean(value, max = 1000) {
        return String(value || '').replace(/\s+/g, ' ').trim().slice(0, max);
    }

    function normalize(value) {
        return clean(value)
            .normalize('NFD')
            .replace(/[\u0300-\u036f]/g, '')
            .toLowerCase()
            .replace(/[^a-z0-9]+/g, ' ')
            .replace(/\s+/g, ' ')
            .trim();
    }

    function editDistance(left, right) {
        const a = String(left || '');
        const b = String(right || '');
        const row = Array.from({ length: b.length + 1 }, (_, index) => index);
        for (let i = 1; i <= a.length; i += 1) {
            let previous = row[0];
            row[0] = i;
            for (let j = 1; j <= b.length; j += 1) {
                const saved = row[j];
                row[j] = Math.min(
                    row[j] + 1,
                    row[j - 1] + 1,
                    previous + (a[i - 1] === b[j - 1] ? 0 : 1)
                );
                previous = saved;
            }
        }
        return row[b.length];
    }

    function titleHasPaymentSignal(value) {
        const normalized = normalize(value);
        const tokens = normalized.split(' ').filter(Boolean);
        if (/\b(facture\s+acquittee|preuve\s+de\s+paiement|recu\s+de\s+paiement)\b/.test(normalized)) return true;
        return tokens.some(token => PAYMENT_WORDS.some(keyword => {
            if (token === keyword) return true;
            if (token.length < 5 || keyword.length < 5) return false;
            return Math.abs(token.length - keyword.length) <= 2 && editDistance(token, keyword) <= 2;
        }));
    }

    function normalizePersonName(value) {
        const latinized = String(value || '').replace(/[аеорсхуікмтвн]/gi, character => ({
            а: 'a', е: 'e', о: 'o', р: 'p', с: 'c', х: 'x', у: 'y', і: 'i', к: 'k', м: 'm', т: 't', в: 'b', н: 'h',
        })[character.toLowerCase()] || character);
        return normalize(latinized);
    }

    function monthStateKey(monthValue) {
        return `${MONTH_STATE_PREFIX}${monthValue}`;
    }

    function normalizeStoredState(parsed) {
        if (!parsed || typeof parsed !== 'object') return null;
        parsed.version = CURRENT_VERSION;
        parsed.review = {
            overrides: {},
            contractOverrides: {},
            validatedRowIds: [],
            overviewConfirmed: false,
            completedAt: null,
            fingerprint: null,
            ...(parsed.review || {}),
        };
        return parsed;
    }

    function loadState() {
        try {
            const activeMonth = localStorage.getItem(ACTIVE_MONTH_KEY);
            if (activeMonth) {
                const active = normalizeStoredState(JSON.parse(localStorage.getItem(monthStateKey(activeMonth)) || 'null'));
                if (active) return active;
            }
            const legacy = normalizeStoredState(JSON.parse(localStorage.getItem(STORAGE_KEY) || 'null'));
            if (legacy?.month) {
                localStorage.setItem(monthStateKey(legacy.month), JSON.stringify(legacy));
                localStorage.setItem(ACTIVE_MONTH_KEY, legacy.month);
                localStorage.removeItem(STORAGE_KEY);
            }
            return legacy;
        } catch (_) {
            return null;
        }
    }

    function loadMonthState(monthValue) {
        try {
            const saved = normalizeStoredState(JSON.parse(localStorage.getItem(monthStateKey(monthValue)) || 'null'));
            if (saved) return saved;
        } catch (_) { /* mois absent ou corrompu */ }
        const legacy = loadState();
        return legacy?.month === monthValue ? legacy : null;
    }

    function activateMonthState(monthValue) {
        const state = loadMonthState(monthValue);
        if (!state) return null;
        localStorage.setItem(ACTIVE_MONTH_KEY, monthValue);
        return state;
    }

    function saveState(state) {
        state.updatedAt = new Date().toISOString();
        if (state.month) {
            localStorage.setItem(monthStateKey(state.month), JSON.stringify(state));
            localStorage.setItem(ACTIVE_MONTH_KEY, state.month);
        }
        localStorage.removeItem(STORAGE_KEY);
        updatePanel();
    }

    function clearState(monthValue) {
        const activeMonth = localStorage.getItem(ACTIVE_MONTH_KEY);
        if (monthValue) localStorage.removeItem(monthStateKey(monthValue));
        else {
            Object.keys(localStorage)
                .filter(key => key.startsWith(MONTH_STATE_PREFIX))
                .forEach(key => localStorage.removeItem(key));
        }
        if (!monthValue || activeMonth === monthValue) localStorage.removeItem(ACTIVE_MONTH_KEY);
        localStorage.removeItem(STORAGE_KEY);
        const request = loadStartRequest();
        if (!monthValue || request?.month === monthValue) clearStartRequest();
        updatePanel();
    }

    function formatMonthLabel(monthValue) {
        const match = /^(\d{4})-(\d{2})$/.exec(monthValue || '');
        if (!match) return monthValue || 'Mois non défini';
        const date = new Date(Number(match[1]), Number(match[2]) - 1, 1);
        const label = new Intl.DateTimeFormat('fr-FR', { month: 'long', year: 'numeric' }).format(date);
        return label.charAt(0).toUpperCase() + label.slice(1);
    }

    function formatDateTime(value) {
        if (!value) return '';
        const date = new Date(value);
        if (Number.isNaN(date.getTime())) return '';
        return new Intl.DateTimeFormat('fr-FR', {
            day: '2-digit',
            month: '2-digit',
            year: 'numeric',
            hour: '2-digit',
            minute: '2-digit',
        }).format(date);
    }

    function previousMonthValue() {
        const date = new Date();
        date.setDate(1);
        date.setMonth(date.getMonth() - 1);
        return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}`;
    }

    function isAgisCompanyCard() {
        return location.pathname === AGIS_COMPANY_PATH
            && new URLSearchParams(location.search).get('company_id') === AGIS_COMPANY_ID;
    }

    function queueDeclarationStart(monthValue) {
        monthBounds(monthValue);
        localStorage.setItem(START_REQUEST_KEY, JSON.stringify({
            month: monthValue,
            stage: 'open-list',
            requestedAt: new Date().toISOString(),
        }));
        location.href = `${location.origin}${LIST_PATH}`;
    }

    function loadStartRequest() {
        try {
            const request = JSON.parse(localStorage.getItem(START_REQUEST_KEY) || 'null');
            return request && /^\d{4}-\d{2}$/.test(request.month || '') ? request : null;
        } catch (_) {
            return null;
        }
    }

    function saveStartRequest(request) {
        localStorage.setItem(START_REQUEST_KEY, JSON.stringify(request));
    }

    function clearStartRequest() {
        localStorage.removeItem(START_REQUEST_KEY);
    }

    function monthBounds(monthValue) {
        const match = /^(\d{4})-(\d{2})$/.exec(monthValue || '');
        if (!match) throw new Error('Mois invalide.');
        const year = Number(match[1]);
        const monthIndex = Number(match[2]) - 1;
        const start = new Date(year, monthIndex, 1);
        const end = new Date(year, monthIndex + 1, 0);
        const format = date => `${String(date.getDate()).padStart(2, '0')}/${String(date.getMonth() + 1).padStart(2, '0')}/${date.getFullYear()}`;
        return { start: format(start), end: format(end), label: monthValue };
    }

    function selectOptionByText(select, targetText) {
        if (!select) return false;
        const wanted = normalize(targetText);
        let found = false;
        for (const option of select.options) {
            const matches = normalize(option.textContent) === wanted;
            option.selected = matches;
            if (matches) found = true;
        }
        select.dispatchEvent(new Event('change', { bubbles: true }));
        if (window.jQuery) {
            try { window.jQuery(select).multipleSelect('refresh'); } catch (_) { /* plugin absent */ }
        }
        return found;
    }

    function setInput(name, value) {
        const input = document.querySelector(`[name="${CSS.escape(name)}"]`);
        if (!input) return false;
        input.value = value;
        input.dispatchEvent(new Event('input', { bubbles: true }));
        input.dispatchEvent(new Event('change', { bubbles: true }));
        return true;
    }

    function findSearchButton() {
        const candidates = Array.from(document.querySelectorAll('button, input[type="submit"], a.button_link'))
            .filter(element => {
                if (element.closest(`#${APP_ID}`)) return false;
                if (element.classList.contains('icon_global_search')) return false;
                const label = normalize(element.textContent || element.value || element.getAttribute('aria-label'));
                if (label !== 'rechercher') return false;
                const rect = element.getBoundingClientRect();
                return rect.width > 0 && rect.height > 0;
            });

        return candidates.find(element => element.closest('aside'))
            || candidates.find(element => element.closest('form.search_filter_form'))
            || candidates[0]
            || null;
    }

    function applyListFilters(monthValue) {
        if (location.pathname !== LIST_PATH) throw new Error('Ouvrez la liste générale des contrats.');
        const bounds = monthBounds(monthValue);
        const company = document.querySelector('select[name="policiesFilter[company_id|in][]"]');
        const status = document.querySelector('select[name="policiesFilter[status|in][]"]');
        const results = {
            company: selectOptionByText(company, 'AGIS'),
            status: selectOptionByText(status, 'En cours'),
            start: setInput('policiesFilter[first_effect_date|greater_equal]', bounds.start),
            end: setInput('policiesFilter[first_effect_date|lower_equal]', bounds.end),
        };
        if (Object.values(results).some(value => !value)) {
            throw new Error(`Filtres incomplets : ${JSON.stringify(results)}`);
        }

        const state = {
            version: CURRENT_VERSION,
            mode: 'visible-policy-batch',
            status: 'filters-applied',
            month: bounds.label,
            period: { start: bounds.start, end: bounds.end },
            company: 'AGIS',
            listUrl: location.href,
            selectedPolicy: null,
            policies: [],
            currentIndex: 0,
            results: [],
            client: null,
            ged: null,
            warnings: [],
            review: { overrides: {}, validatedRowIds: [], completedAt: null },
            analysisStartedAt: loadStartRequest()?.requestedAt || new Date().toISOString(),
            createdAt: new Date().toISOString(),
        };
        saveState(state);
        setMessage(`Filtres AGIS / En cours / effet du ${bounds.start} au ${bounds.end} appliqués. Cliquez sur « Rechercher Modulr » puis lancez le test.`);
    }

    function submitFilters() {
        const form = document.querySelector('form.search_filter_form');
        if (!form) throw new Error('Formulaire des filtres contrats introuvable. Ouvrez d’abord le panneau des filtres.');

        const submitButton = Array.from(form.querySelectorAll('button[type="submit"], input[type="submit"]'))
            .find(element => normalize(element.textContent || element.value) === 'rechercher');

        if (typeof form.requestSubmit === 'function') {
            form.requestSubmit(submitButton || undefined);
            return;
        }

        const submitEvent = new Event('submit', { bubbles: true, cancelable: true });
        const allowed = form.dispatchEvent(submitEvent);
        if (allowed) HTMLFormElement.prototype.submit.call(form);
    }

    function getPolicyRows() {
        const table = Array.from(document.querySelectorAll('table.table_list'))
            .find(candidate => {
                const text = normalize(candidate.querySelector('thead')?.textContent || candidate.querySelector('tr')?.textContent);
                return text.includes('compagnie') && text.includes('date d effet') && text.includes('etat');
            });
        if (!table) return [];
        return Array.from(table.querySelectorAll('tbody tr')).filter(row => row.querySelector('input[id^="forSelection_"]'));
    }

    function parsePolicyRow(row) {
        const cells = Array.from(row.querySelectorAll(':scope > td'));
        const checkbox = row.querySelector('input[id^="forSelection_"]');
        const policyId = checkbox?.id.match(/(\d+)$/)?.[1] || '';
        return {
            policyId,
            referenceAndBusinessType: clean(cells[1]?.textContent),
            company: clean(cells[2]?.textContent),
            product: clean(cells[3]?.textContent),
            identity: clean(cells[4]?.textContent),
            effectDate: clean(cells[5]?.textContent),
            state: clean(cells[6]?.textContent),
            rowText: clean(row.textContent),
        };
    }

    function openPolicyRow(row, policy) {
        const clickable = row.querySelector('td:nth-child(5), td:nth-child(2)') || row;
        setMessage(`Ouverture ${policy.identity || policy.policyId}…`);
        setTimeout(() => clickable.click(), 300);
    }

    function startVisibleBatch() {
        if (location.pathname !== LIST_PATH) throw new Error('Revenez sur la liste générale des contrats.');
        const rows = getPolicyRows();
        if (!rows.length) throw new Error('Aucun contrat visible après filtrage.');

        const policies = rows.map(parsePolicyRow);
        const policy = policies[0];
        const state = loadState() || {
            version: CURRENT_VERSION,
            mode: 'visible-policy-batch',
            month: previousMonthValue(),
            company: 'AGIS',
            warnings: [],
            createdAt: new Date().toISOString(),
        };
        state.version = CURRENT_VERSION;
        state.mode = 'visible-policy-batch';
        state.status = 'opening-client-card';
        state.listUrl = location.href;
        state.policies = policies;
        state.currentIndex = 0;
        state.results = [];
        state.warnings = [];
        state.review = { overrides: {}, validatedRowIds: [], completedAt: null };
        state.selectedPolicy = policy;
        state.client = null;
        state.ged = null;
        state.gedFilterClearAttempts = 0;
        state.gedFilterWasActive = false;
        state.gedFilterCleared = false;
        if (!state.period && state.month) {
            const bounds = monthBounds(state.month);
            state.period = { start: bounds.start, end: bounds.end };
        }
        saveState(state);
        setMessage(`Traitement de 1/${policies.length} : ${policy.identity || policy.policyId}.`);
        openPolicyRow(rows[0], policy);
    }

    function resumeCompanyStart() {
        if (location.pathname !== LIST_PATH) return false;
        const request = loadStartRequest();
        if (!request) return false;

        if (request.stage === 'open-list') {
            applyListFilters(request.month);
            request.stage = 'awaiting-results';
            request.filteredAt = new Date().toISOString();
            saveStartRequest(request);
            setTimeout(() => {
                try {
                    submitFilters();
                } catch (error) {
                    clearStartRequest();
                    setMessage(`Erreur : ${error.message}`);
                }
            }, 350);
            return true;
        }

        let attempts = 0;
        const startWhenReady = () => {
            const rows = getPolicyRows();
            if (rows.length) {
                clearStartRequest();
                startVisibleBatch();
                return;
            }
            attempts += 1;
            if (attempts < 20) {
                setTimeout(startWhenReady, 500);
                return;
            }
            clearStartRequest();
            setMessage('Aucun contrat AGIS visible après la recherche. Vérifiez les filtres.');
        };
        setTimeout(startWhenReady, 700);
        return true;
    }

    function resumeVisibleBatch() {
        const state = loadState();
        if (!state || state.status !== 'returning-list') return;
        const policy = state.policies?.[state.currentIndex];
        if (!policy) return;
        const row = getPolicyRows().find(candidate => parsePolicyRow(candidate).policyId === policy.policyId);
        if (!row) {
            state.status = 'error';
            state.warnings.push(`Contrat ${policy.policyId} introuvable au retour sur la liste.`);
            saveState(state);
            setMessage(`Contrat ${policy.policyId} introuvable. Exportez le résultat partiel.`);
            return;
        }
        state.status = 'opening-client-card';
        state.selectedPolicy = policy;
        state.client = null;
        state.ged = null;
        state.gedFilterClearAttempts = 0;
        state.gedFilterWasActive = false;
        state.gedFilterCleared = false;
        saveState(state);
        setMessage(`Traitement de ${state.currentIndex + 1}/${state.policies.length} : ${policy.identity || policy.policyId}.`);
        openPolicyRow(row, policy);
    }

    function extractPhone() {
        const scope = document.querySelector('td.valign_top.medium_padding_right > div.card_template_section')
            || document.querySelector('.card_template_section')
            || document;
        const candidates = Array.from(scope.querySelectorAll('a, li, span, p'));
        for (const element of candidates) {
            const text = clean(element.textContent, 80);
            const match = text.match(/(?:\+33\s?|0)[1-9](?:[ .-]?\d{2}){4}/);
            if (match) return match[0];
        }
        return '';
    }

    function extractEmail() {
        const scope = document.querySelector('td.valign_top.medium_padding_right > div.card_template_section')
            || document.querySelector('.card_template_section')
            || document;
        const candidates = Array.from(scope.querySelectorAll('a, li, span, p'));
        for (const element of candidates) {
            const text = clean(element.textContent, 160);
            const match = text.match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/i);
            if (match) return match[0];
        }
        return '';
    }

    function normalizeCivility(value) {
        const normalized = normalize(value);
        if (/\b(mme|madame)\b/.test(normalized)) return 'Mme';
        if (/\b(mlle|mademoiselle)\b/.test(normalized)) return 'Mlle';
        if (/\b(monsieur|m)\b/.test(normalized)) return 'M.';
        return '';
    }

    function extractClientPostalDetails(clientId, identity) {
        const section = document.querySelector('td.valign_top.medium_padding_right > div.card_template_section')
            || document.querySelector('.card_template_section');
        const rawText = section?.innerText || section?.textContent || '';
        const lines = rawText.split(/\r?\n/).map(line => clean(line, 250)).filter(Boolean);
        const identityTokens = normalize(identity).split(' ').filter(token => token.length >= 2);
        const civility = normalizeCivility(lines.find(line => {
            const normalizedLine = normalize(line);
            return identityTokens.length >= 2 && identityTokens.every(token => normalizedLine.includes(token));
        }) || rawText);

        let postalCode = '';
        let city = '';
        let postalLineIndex = -1;
        for (let index = 0; index < lines.length; index += 1) {
            const match = lines[index].match(/\b(\d{5})\b\s*[-–]?\s*(.*)$/);
            if (!match) continue;
            postalCode = match[1];
            city = clean(match[2], 120);
            postalLineIndex = index;
            break;
        }

        let addressLines = [];
        if (postalLineIndex >= 0) {
            let startIndex = 0;
            for (let index = postalLineIndex - 1; index >= 0; index -= 1) {
                const line = lines[index];
                if ((clientId && normalize(line).includes(normalize(`n°${clientId}`)))
                    || (identityTokens.length >= 2 && identityTokens.every(token => normalize(line).includes(token)))) {
                    startIndex = index + 1;
                    break;
                }
            }
            addressLines = lines.slice(startIndex, postalLineIndex).filter(line =>
                !/^\d{1,3}$/.test(line)
                && !/^(coordonn[eé]es|uuid|cr[eé]ation|derni[eè]re modification)/i.test(line)
                && !/(?:\+33\s?|0)[1-9](?:[ .-]?\d{2}){4}/.test(line)
                && !/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/i.test(line)
            );
        }

        if (!postalCode || !addressLines.length) {
            const compact = clean(rawText, 4000);
            const match = compact.match(/n[°o]\s*\d+\s+(?:\d\s+)?(.+?)\s+(\d{5})\s+(.+?)(?=(?:\+33\s?|0)[1-9](?:[ .-]?\d{2}){4}|[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}|$)/i);
            if (match) {
                addressLines = [clean(match[1], 250)];
                postalCode = match[2];
                city = clean(match[3], 120);
            }
        }

        return {
            civility,
            address: addressLines[0] || '',
            address2: addressLines.slice(1).join(' ') || '',
            postalCode,
            city,
        };
    }

    function findPolicyRowOnClientCard(policyId) {
        const marker = document.querySelector(`#expected_document_policy_${CSS.escape(policyId)}`);
        if (marker) return marker.closest('tr');
        const tables = Array.from(document.querySelectorAll('table.table_list'));
        return tables.flatMap(table => Array.from(table.querySelectorAll('tbody tr')))
            .find(row => normalize(row.textContent).includes('agis')) || null;
    }

    function extractBeneficiaries() {
        const tables = Array.from(document.querySelectorAll('table'));
        const table = tables.find(candidate => {
            const header = normalize(candidate.querySelector('thead')?.textContent || candidate.querySelector('tr')?.textContent);
            return header.includes('nom') && header.includes('date d adhesion') && header.includes('date radiation');
        });
        if (!table) return [];
        return Array.from(table.querySelectorAll('tbody tr'))
            .map(row => Array.from(row.querySelectorAll(':scope > td')).map(cell => clean(cell.textContent)))
            .filter(cells => cells.some(Boolean))
            .map(cells => ({ rawCells: cells }));
    }

    function extractLinkedClients() {
        const linkedClients = [];
        const seen = new Set();
        document.querySelectorAll('a[href*="clients_card.php"]').forEach(link => {
            let url;
            try {
                url = new URL(link.href, location.origin);
            } catch (_) {
                return;
            }
            const clientId = url.searchParams.get('id') || '';
            if (!clientId || seen.has(clientId)) return;
            const container = link.closest('tr, li, [class*="linked_contact"]') || link.parentElement;
            const label = clean(link.textContent || container?.textContent, 300);
            if (!label) return;
            seen.add(clientId);
            linkedClients.push({
                clientId,
                label,
                clientUrl: `${location.origin}${CLIENT_PATH}?id=${clientId}`,
                gedUrl: `${location.origin}${GED_PREFIX}${clientId}`,
            });
        });
        return linkedClients;
    }

    function captureClientCardAndOpenGed() {
        const state = loadState();
        if (!state || state.status !== 'opening-client-card') return;
        const policyId = state.selectedPolicy?.policyId || '';
        const title = clean(document.querySelector('h1.page_title')?.textContent || document.title);
        const policyRow = findPolicyRowOnClientCard(policyId);
        const clientId = document.querySelector('[id*="entity_id"]')?.id.match(/entity_id[:_](\d+)/)?.[1]
            || document.querySelector('[class*="linked_contacts_container_"]')?.className.match(/linked_contacts_container_(\d+)/)?.[1]
            || '';
        const pageTextSample = clean(document.querySelector('.entity_info_block')?.textContent, 2500);
        const clientBirthDate = pageTextSample.match(/Date de naissance(\d{2}\/\d{2}\/\d{4})/i)?.[1] || '';
        const postalDetails = extractClientPostalDetails(clientId, state.selectedPolicy?.identity || '');

        state.client = {
            url: location.href,
            title,
            clientId,
            phone: extractPhone(),
            email: extractEmail(),
            birthDate: clientBirthDate,
            ...postalDetails,
            policyRowText: clean(policyRow?.textContent),
            beneficiaries: extractBeneficiaries(),
            linkedClients: extractLinkedClients(),
            primaryDetailsText: clean((document.querySelector('td.valign_top.medium_padding_right > div.card_template_section')
                || document.querySelector('.card_template_section'))?.textContent, 2500),
            pageTextSample,
        };

        const gedLink = Array.from(document.querySelectorAll('a[href]')).find(link =>
            normalize(link.textContent).includes('ouvrir la ged') ||
            (link.classList.contains('ged_launcher') && link.closest('tr') === policyRow)
        );
        const directGedUrl = clientId ? `${location.origin}${GED_PREFIX}${clientId}` : '';
        if (!directGedUrl && !gedLink) {
            state.status = 'error';
            state.warnings.push('Lien GED introuvable sur la fiche client.');
            saveState(state);
            setMessage('Fiche client lue, mais lien GED introuvable. Exportez le diagnostic de test.');
            return;
        }

        state.status = 'opening-ged';
        state.gedUrl = directGedUrl || gedLink.href;
        saveState(state);
        setMessage('Coordonnées et personnes assurées récupérées. Ouverture de la GED…');
        setTimeout(() => {
            if (directGedUrl) {
                location.href = directGedUrl;
            } else {
                gedLink.click();
            }
        }, 500);
    }

    function findGedClearFilterLink() {
        return document.querySelector('p.client_subentity_container a.fa.fa-ban')
            || Array.from(document.querySelectorAll('a')).find(link => normalize(link.textContent || link.title) === 'effacer le filtre')
            || null;
    }

    async function prepareFullGedAndCapture() {
        const state = loadState();
        if (!state || !['opening-ged', 'clearing-ged-filter'].includes(state.status)) return;
        const clearLink = findGedClearFilterLink();
        const attempts = Number(state.gedFilterClearAttempts || 0);
        if (clearLink && attempts < 2) {
            state.status = 'clearing-ged-filter';
            state.gedFilterClearAttempts = attempts + 1;
            state.gedFilterWasActive = true;
            saveState(state);
            setMessage('Suppression du filtre contrat pour analyser toute la GED client…');
            clearLink.click();
            setTimeout(() => prepareFullGedAndCapture().catch(error => setMessage(`Erreur GED complète : ${error.message}`)), 1800);
            return;
        }
        state.status = 'opening-ged';
        state.gedFilterCleared = !clearLink;
        state.gedDocumentsScope = clearLink ? 'filtered-after-clear-failure' : 'all-client-documents';
        if (clearLink) state.warnings.push(`Impossible de supprimer le filtre GED pour le contrat ${state.selectedPolicy?.policyId || ''}.`);
        saveState(state);
        await captureGedAndContinue();
    }

    function parseGedDocuments(scope = document, pageNumber = 1) {
        const rows = Array.from(scope.querySelectorAll('table.table_ged tbody tr.no_hover_background'));
        const documents = rows.map(row => {
            const link = row.querySelector('a[id*="_ged_link_"][href*="/edm/download/"]');
            if (!link) return null;
            const text = clean(row.textContent, 1800);
            const dateMatch = text.match(/Date d'ajout\s*:\s*(\d{2}\/\d{2}\/\d{4})\s+(\d{2}:\d{2})/i);
            const idMatch = link.href.match(/\/download\/(\d+)/);
            const name = clean(link.textContent);
            const isPaymentCandidate = titleHasPaymentSignal(name);
            const isQuoteCandidate = QUOTE_WORDS.test(name);
            const isCoverageCandidate = COVERAGE_WORDS.test(name)
                || (!isPaymentCandidate && /\b(adobe scan|scan|document)\b/i.test(name));
            const isAttestation = /\battestation\b/i.test(name);
            return {
                documentId: idMatch?.[1] || '',
                name,
                uploadDate: dateMatch ? `${dateMatch[1]} ${dateMatch[2]}` : '',
                downloadUrl: link.href,
                gedPage: pageNumber,
                isPaymentCandidate,
                isQuoteCandidate,
                isCoverageCandidate,
                isAttestation,
                documentRole: isAttestation
                    ? 'final-attestation'
                    : (isPaymentCandidate && isCoverageCandidate ? 'contract-and-payment' : (isPaymentCandidate ? 'payment' : (isCoverageCandidate ? 'contractual' : 'other'))),
                rawText: text,
            };
        }).filter(Boolean);
        return documents.filter((documentInfo, index, all) =>
            all.findIndex(candidate => candidate.documentId === documentInfo.documentId) === index
        );
    }

    function findGedNextPageLink() {
        const table = document.querySelector('table.table_ged');
        const scope = table?.closest('.ged_container') || table?.parentElement || document;
        return Array.from(scope.querySelectorAll('a, button')).find(element => {
            const label = normalize(`${element.textContent || ''} ${element.title || ''} ${element.getAttribute('aria-label') || ''}`);
            const disabled = element.disabled
                || element.getAttribute('aria-disabled') === 'true'
                || element.classList.contains('disabled');
            return !disabled && (label === 'page suivante' || label === 'suivant' || label.includes('page suivante'));
        }) || null;
    }

    function waitForGedPageChange(previousSignature, timeout = 12000) {
        return new Promise((resolve, reject) => {
            const startedAt = Date.now();
            const check = () => {
                const signature = parseGedDocuments().map(item => item.documentId).join('|');
                if (signature && signature !== previousSignature) {
                    resolve();
                    return true;
                }
                if (Date.now() - startedAt >= timeout) {
                    reject(new Error('La page suivante de la GED ne s’est pas chargée.'));
                    return true;
                }
                return false;
            };
            if (check()) return;
            const observer = new MutationObserver(() => {
                if (check()) observer.disconnect();
            });
            observer.observe(document.body, { childList: true, subtree: true });
            const timer = setInterval(() => {
                if (check()) {
                    clearInterval(timer);
                    observer.disconnect();
                }
            }, 350);
        });
    }

    async function collectAllGedDocuments(state) {
        const collected = [];
        const seenIds = new Set();
        const seenPages = new Set();
        let pageNumber = 1;
        let complete = true;

        while (pageNumber <= 50) {
            const pageDocuments = parseGedDocuments(document, pageNumber);
            const signature = pageDocuments.map(item => item.documentId).join('|');
            if (!signature || seenPages.has(signature)) break;
            seenPages.add(signature);
            pageDocuments.forEach(item => {
                if (seenIds.has(item.documentId)) return;
                seenIds.add(item.documentId);
                collected.push(item);
            });
            const nextLink = findGedNextPageLink();
            if (!nextLink) break;
            setMessage(`GED : ${collected.length} pièce(s) trouvée(s), ouverture de la page ${pageNumber + 1}…`);
            nextLink.click();
            try {
                await waitForGedPageChange(signature);
            } catch (error) {
                complete = false;
                state.warnings.push(`Pagination GED interrompue après ${collected.length} pièce(s) : ${error.message}`);
                break;
            }
            pageNumber += 1;
        }
        if (pageNumber > 50 && findGedNextPageLink()) {
            complete = false;
            state.warnings.push('Plus de 50 pages GED détectées : arrêt de sécurité et contrôle humain requis.');
        }
        state.gedDocumentsScope = complete ? 'all-client-documents-all-pages' : 'all-client-documents-partial-pages';
        state.gedPagesScanned = pageNumber;
        state.gedDocumentsScanned = collected.length;
        saveState(state);
        return collected;
    }

    function extractCoveredPersons(text) {
        const persons = [];
        const rejectedLabels = /\b(p[eé]riode|garantie|date|[eé]mission|effet|contrat|devis|facture|r[eè]glement|paiement|signature|cotisation|montant|total|page|paris)\b/i;
        for (const rawLine of String(text || '').split(/\r?\n/)) {
            const line = clean(rawLine, 180).replace(/^[^\p{L}]+/u, '');
            const match = line.match(/^([\p{L}][\p{L}'’!()., -]{1,120}?)\s*(?:[-–:]\s*)?([0-3]?\d[\/.-][01]?\d[\/.-](?:18|19|20)\d{2})\b/u);
            if (!match) continue;
            const name = clean(match[1]).replace(/^[,.;!\s]+|[,.;!\s]+$/g, '').replace(/!+/g, 'I');
            const nameWords = normalize(name).split(' ').filter(Boolean);
            if (!name || nameWords.length < 2 || rejectedLabels.test(name)) continue;
            const birthDate = match[2].replace(/[.-]/g, '/').split('/').map((part, index) =>
                index < 2 ? part.padStart(2, '0') : part
            ).join('/');
            persons.push({ name, birthDate });
        }
        return persons.filter((person, index, all) =>
            all.findIndex(candidate => normalize(candidate.name) === normalize(person.name) && candidate.birthDate === person.birthDate) === index
        );
    }

    function extractStructuredPersons(text) {
        const persons = [];
        const pattern = /(?:^|\n)\s*Nom\s*:\s*\n\s*([^\n]{2,100})\s*\n\s*Pr[eé]nom(?:\(s\))?\s*:\s*\n\s*([^\n]{2,100})\s*\n\s*Date de naissance\s*:\s*\n\s*(\d{1,2}[\/.-]\d{1,2}[\/.-](?:18|19|20)\d{2})/giu;
        for (const match of String(text || '').matchAll(pattern)) {
            const lastName = clean(match[1], 100).replace(/^[,.;\s]+|[,.;\s]+$/g, '');
            const firstName = clean(match[2], 100).replace(/^[,.;\s]+|[,.;\s]+$/g, '');
            const birthDate = match[3].replace(/[.-]/g, '/').split('/').map((part, index) =>
                index < 2 ? part.padStart(2, '0') : part
            ).join('/');
            if (lastName && firstName) persons.push({
                name: `${lastName} ${firstName}`,
                lastName,
                firstName,
                birthDate,
                structuredFields: true,
            });
        }
        return persons;
    }

    function extractAllCoveredPersons(text) {
        const persons = [...extractCoveredPersons(text), ...extractStructuredPersons(text)];
        return persons.filter((person, index, all) =>
            all.findIndex(candidate => normalize(candidate.name) === normalize(person.name) && candidate.birthDate === person.birthDate) === index
        );
    }

    function reconcileCoveredPersons(persons, state) {
        const identityTokens = normalizePersonName(state.selectedPolicy?.identity).split(' ').filter(token => token.length >= 3);
        const clientBirthDate = state.client?.birthDate || '';
        return persons.map(person => {
            const nameNormalized = normalizePersonName(person.name);
            const matchesSubscriber = identityTokens.length >= 2 && identityTokens.every(token => nameNormalized.includes(token));
            if (matchesSubscriber && clientBirthDate) {
                const subscriberFields = {
                    documentName: person.name,
                    name: clean(state.selectedPolicy?.identity),
                    crossCheckedSubscriber: true,
                };
                if (person.birthDate !== clientBirthDate) {
                    return {
                        ...person,
                        ...subscriberFields,
                        documentBirthDate: person.birthDate,
                        birthDate: clientBirthDate,
                        birthDateSource: 'fiche-client-apres-controle-ocr',
                    };
                }
                return { ...person, ...subscriberFields, birthDateSource: 'pdf-et-fiche-client' };
            }
            return {
                ...person,
                birthDateSource: 'pdf',
                needsReview: Number(person.birthDate.slice(-4)) < 1900,
            };
        });
    }

    function editDistance(first, second) {
        const left = String(first || '');
        const right = String(second || '');
        const row = Array.from({ length: right.length + 1 }, (_, index) => index);
        for (let leftIndex = 1; leftIndex <= left.length; leftIndex += 1) {
            let previous = row[0];
            row[0] = leftIndex;
            for (let rightIndex = 1; rightIndex <= right.length; rightIndex += 1) {
                const saved = row[rightIndex];
                row[rightIndex] = Math.min(
                    row[rightIndex] + 1,
                    row[rightIndex - 1] + 1,
                    previous + (left[leftIndex - 1] === right[rightIndex - 1] ? 0 : 1)
                );
                previous = saved;
            }
        }
        return row[right.length];
    }

    function personNameTokens(person) {
        return normalizePersonName(person?.name).split(' ').filter(Boolean);
    }

    function personNamesProbablyMatch(first, second) {
        const firstTokens = personNameTokens(first);
        const secondTokens = personNameTokens(second);
        if (firstTokens.length < 2 || secondTokens.length < 2) return false;
        const firstKey = [...firstTokens].sort().join('|');
        const secondKey = [...secondTokens].sort().join('|');
        if (firstKey === secondKey) return true;
        if (!first.birthDate || first.birthDate !== second.birthDate) return false;

        const firstRemaining = [...firstTokens];
        const secondRemaining = [...secondTokens];
        for (let index = firstRemaining.length - 1; index >= 0; index -= 1) {
            const matchIndex = secondRemaining.indexOf(firstRemaining[index]);
            if (matchIndex < 0) continue;
            firstRemaining.splice(index, 1);
            secondRemaining.splice(matchIndex, 1);
        }
        return firstRemaining.length === 1
            && secondRemaining.length === 1
            && editDistance(firstRemaining[0], secondRemaining[0]) <= 2;
    }

    function personEvidenceScore(person) {
        const source = normalize(person.sourceDocumentName);
        return (person.crossCheckedSubscriber ? 100 : 0)
            + (/attestation|certificat/.test(source) ? 30 : 0)
            + (person.structuredFields ? 10 : 0)
            + (/signe|contrat|bulletin/.test(source) ? 5 : 0)
            + (person.lastName && person.firstName ? 2 : 0);
    }

    function deduplicateCoveredPersons(persons) {
        const groups = [];
        for (const person of persons) {
            const group = groups.find(candidate => personNamesProbablyMatch(candidate.selected, person));
            if (!group) {
                groups.push({ selected: person, members: [person] });
                continue;
            }
            group.members.push(person);
            if (personEvidenceScore(person) > personEvidenceScore(group.selected)) group.selected = person;
        }
        return groups.map(group => {
            const aliases = group.members.map(person => clean(person.name)).filter((value, index, all) => all.indexOf(value) === index);
            const birthDates = group.members.map(person => person.birthDate).filter(Boolean)
                .filter((value, index, all) => all.indexOf(value) === index);
            return {
                ...group.selected,
                aliases,
                conflictingBirthDates: birthDates.length > 1 ? birthDates : [],
                evidenceDocuments: group.members.map(person => person.sourceDocumentName).filter(Boolean)
                    .filter((value, index, all) => all.indexOf(value) === index),
            };
        });
    }

    function addSubscriberFromDocumentEvidence(persons, documents, state) {
        const identity = clean(state.selectedPolicy?.identity);
        const identityTokens = normalize(identity).split(' ').filter(token => token.length >= 3);
        const birthDate = state.client?.birthDate || '';
        if (identityTokens.length < 2 || !birthDate) return persons;
        const alreadyPresent = persons.some(person => {
            const personName = normalize(person.name);
            return identityTokens.every(token => personName.includes(token));
        });
        if (alreadyPresent) return persons;

        const source = documents.find(documentInfo => {
            const content = `${documentInfo.contentRead?.extractedText || ''}\n${documentInfo.contentRead?.ocrText || ''}`;
            const normalizedContent = normalize(content);
            return content.includes(birthDate) && identityTokens.every(token => normalizedContent.includes(token));
        });
        if (!source) return persons;
        return persons.concat({
            name: identity,
            birthDate,
            sourceDocumentId: source.documentId,
            sourceDocumentName: source.name,
            birthDateSource: 'pdf-et-fiche-client',
            crossCheckedSubscriber: true,
            extractedFromStructuredEvidence: true,
        });
    }

    async function ocrPdfPages(pdf, documentName) {
        const tesseract = typeof Tesseract !== 'undefined' ? Tesseract : window.Tesseract;
        if (!tesseract?.createWorker) {
            return { status: 'ocr-unavailable', ocrText: '', coveredPersons: [] };
        }
        let worker;
        try {
            setMessage(`Initialisation OCR pour « ${documentName} »…`);
            worker = await tesseract.createWorker(
                'fra',
                tesseract.OEM?.LSTM_ONLY ?? 1,
                TESSERACT_OCR_OPTIONS
            );
            const pageTexts = [];
            const pageLimit = Math.min(pdf.numPages, 30);
            for (let pageNumber = 1; pageNumber <= pageLimit; pageNumber += 1) {
                setMessage(`OCR « ${documentName} » — page ${pageNumber}/${pageLimit}…`);
                const page = await pdf.getPage(pageNumber);
                const viewport = page.getViewport({ scale: 1.7 });
                const canvas = document.createElement('canvas');
                canvas.width = Math.ceil(viewport.width);
                canvas.height = Math.ceil(viewport.height);
                const context = canvas.getContext('2d', { alpha: false });
                await page.render({ canvasContext: context, viewport }).promise;
                const recognition = await worker.recognize(canvas);
                pageTexts.push(recognition?.data?.text || '');
                canvas.width = 1;
                canvas.height = 1;
            }
            const ocrText = pageTexts.join('\n').slice(0, 30000);
            return {
                status: ocrText.trim()
                    ? (pageLimit < pdf.numPages ? 'ocr-partial' : 'ocr-read')
                    : 'ocr-empty',
                ocrPageCount: pageLimit,
                ocrTotalPages: pdf.numPages,
                ocrText,
                coveredPersons: extractAllCoveredPersons(ocrText),
            };
        } catch (error) {
            return { status: 'ocr-error', error: clean(error.message, 300), ocrText: '', coveredPersons: [] };
        } finally {
            if (worker) {
                try { await worker.terminate(); } catch (_) { /* worker already stopped */ }
            }
        }
    }

    function detectImageMime(bytes) {
        if (bytes[0] === 0xFF && bytes[1] === 0xD8 && bytes[2] === 0xFF) return 'image/jpeg';
        if (bytes[0] === 0x89 && bytes[1] === 0x50 && bytes[2] === 0x4E && bytes[3] === 0x47) return 'image/png';
        if (bytes[0] === 0x47 && bytes[1] === 0x49 && bytes[2] === 0x46) return 'image/gif';
        if (new TextDecoder('ascii').decode(bytes.slice(0, 4)) === 'RIFF'
            && new TextDecoder('ascii').decode(bytes.slice(8, 12)) === 'WEBP') return 'image/webp';
        return '';
    }

    async function ocrImageBytes(bytes, documentName) {
        const tesseract = typeof Tesseract !== 'undefined' ? Tesseract : window.Tesseract;
        const mimeType = detectImageMime(bytes);
        if (!mimeType) return { status: 'unsupported-format', ocrText: '', coveredPersons: [] };
        if (!tesseract?.createWorker) return { status: 'ocr-unavailable', ocrText: '', coveredPersons: [] };
        let worker;
        try {
            setMessage(`OCR de l’image « ${documentName} »…`);
            worker = await tesseract.createWorker(
                'fra',
                tesseract.OEM?.LSTM_ONLY ?? 1,
                TESSERACT_OCR_OPTIONS
            );
            const recognition = await worker.recognize(new Blob([bytes], { type: mimeType }));
            const ocrText = String(recognition?.data?.text || '').slice(0, 30000);
            return {
                status: ocrText.trim() ? 'ocr-read' : 'ocr-empty',
                imageMimeType: mimeType,
                ocrPageCount: 1,
                ocrText,
                coveredPersons: extractAllCoveredPersons(ocrText),
            };
        } catch (error) {
            return { status: 'ocr-error', error: clean(error.message, 300), ocrText: '', coveredPersons: [] };
        } finally {
            if (worker) {
                try { await worker.terminate(); } catch (_) { /* worker already stopped */ }
            }
        }
    }

    function downloadBinaryOnce(url) {
        return new Promise((resolve, reject) => {
            GM_xmlhttpRequest({
                method: 'GET',
                url,
                responseType: 'arraybuffer',
                anonymous: false,
                timeout: 45000,
                onload: response => {
                    if (response.status < 200 || response.status >= 300 || !response.response) {
                        reject(new Error(`Téléchargement HTTP ${response.status || 'inconnu'}`));
                        return;
                    }
                    resolve(new Uint8Array(response.response));
                },
                onerror: () => reject(new Error('Téléchargement Modulr impossible')),
                ontimeout: () => reject(new Error('Délai de téléchargement dépassé')),
            });
        });
    }

    async function downloadBinary(url) {
        let lastError;
        for (let attempt = 1; attempt <= 2; attempt += 1) {
            try {
                return await downloadBinaryOnce(url);
            } catch (error) {
                lastError = error;
                if (attempt < 2) await new Promise(resolve => setTimeout(resolve, 700));
            }
        }
        throw lastError || new Error('Téléchargement Modulr impossible');
    }

    async function readPdfDocument(documentInfo) {
        const pdfjs = typeof pdfjsLib !== 'undefined' ? pdfjsLib : window.pdfjsLib;
        if (!pdfjs?.getDocument) {
            return { status: 'reader-unavailable', extractedText: '', coveredPersons: [] };
        }
        try {
            if (pdfjs.GlobalWorkerOptions && !pdfjs.GlobalWorkerOptions.workerSrc) {
                pdfjs.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
            }
            const bytes = await downloadBinary(documentInfo.downloadUrl);
            const isPdf = new TextDecoder('ascii').decode(bytes.slice(0, 4)) === '%PDF';
            if (!isPdf) return { ...(await ocrImageBytes(bytes, documentInfo.name)), extractedText: '' };

            const pdf = await pdfjs.getDocument({ data: bytes, disableWorker: true }).promise;
            const lines = [];
            for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber += 1) {
                const page = await pdf.getPage(pageNumber);
                const content = await page.getTextContent();
                for (const item of content.items || []) {
                    const value = clean(item.str, 500);
                    if (value) lines.push(value);
                }
            }
            const extractedText = lines.join('\n').slice(0, 20000);
            let coveredPersons = extractAllCoveredPersons(extractedText);
            if (!coveredPersons.length) {
                const ocrResult = await ocrPdfPages(pdf, documentInfo.name);
                if (['ocr-read', 'ocr-partial'].includes(ocrResult.status)) coveredPersons = ocrResult.coveredPersons;
                return {
                    status: ocrResult.status,
                    pageCount: pdf.numPages,
                    extractedText,
                    ocrPageCount: ocrResult.ocrPageCount,
                    ocrTotalPages: ocrResult.ocrTotalPages,
                    ocrText: ocrResult.ocrText,
                    coveredPersons,
                    error: ocrResult.error,
                };
            }
            return {
                status: 'text-read',
                pageCount: pdf.numPages,
                extractedText,
                coveredPersons,
            };
        } catch (error) {
            return { status: 'read-error', error: clean(error.message, 300), extractedText: '', coveredPersons: [] };
        }
    }

    function extractPaymentFacts(documentInfo) {
        const fileContent = `${documentInfo.contentRead?.extractedText || ''}\n${documentInfo.contentRead?.ocrText || ''}`;
        const content = `${documentInfo.name || ''}\n${fileContent}`;
        const methods = [];
        const normalizedContent = normalize(content);
        if (/\b(ch[eè]que|cheque)\b/i.test(content) || titleHasPaymentSignal(documentInfo.name) && /\bcheq\w*\b/.test(normalizedContent)) methods.push('cheque');
        if (/\bvirement\b|ordre de virement/i.test(content) || /\bvire?ment\b/.test(normalizedContent)) methods.push('virement');
        if (/\bstripe\b/i.test(content)) methods.push('stripe');
        if (/\bpaypal\b/i.test(content) || /\bppal\b/.test(normalizedContent)) methods.push('paypal');
        if (/\b(?:carte bancaire|paiement par carte|cb)\b/i.test(content)) methods.push('carte-bancaire');
        if (/\b(?:esp[eè]ces?|liquide|billet)\b/i.test(content)) methods.push('especes');
        if (/\bpr[eé]l[eè]vement\b/i.test(content)) methods.push('prelevement');
        if (/\bfacture\b.*\bacquitt[eé]e?\b|\bacquitt[eé]e?\b/i.test(content)) methods.push('facture-acquittee');
        const amounts = Array.from(content.matchAll(/\b(\d{1,5}(?:[ .]\d{3})*(?:[,.]\d{2})?)\s*(?:€|EUR)/gi))
            .map(match => match[1].replace(/\s/g, '').replace(',', '.'))
            .map(Number)
            .filter(value => Number.isFinite(value) && value > 0)
            .filter((value, index, all) => all.indexOf(value) === index);
        const dates = Array.from(content.matchAll(/\b([0-3]\d\/[01]\d\/(?:19|20)\d{2})\b/g))
            .map(match => match[1])
            .filter((value, index, all) => all.indexOf(value) === index)
            .slice(0, 20);
        return {
            readStatus: documentInfo.contentRead?.status || 'not-read',
            methods: methods.filter((value, index, all) => all.indexOf(value) === index),
            amounts,
            dates,
            hasReadableContent: Boolean(normalize(fileContent)),
            titleNegative: PAYMENT_TITLE_NEGATIVE.test(documentInfo.name || ''),
            contentNegative: PAYMENT_CONTENT_NEGATIVE.test(fileContent),
            hasFinalStatus: PAYMENT_FINAL_POSITIVE.test(fileContent),
            titlePaymentSignal: titleHasPaymentSignal(documentInfo.name),
            isImage: Boolean(documentInfo.contentRead?.imageMimeType),
        };
    }

    async function readRelevantDocuments(documents) {
        for (const documentInfo of documents.filter(item => item.isCoverageCandidate || item.isPaymentCandidate)) {
            setMessage(`Lecture du document « ${documentInfo.name} »…`);
            documentInfo.contentRead = await readPdfDocument(documentInfo);
            if (documentInfo.isPaymentCandidate) documentInfo.paymentFacts = extractPaymentFacts(documentInfo);
        }
    }

    function annotateDocumentRelevance(documentInfo, state) {
        const targetYear = Number(String(state.month || '').slice(0, 4));
        const policyReference = clean(state.selectedPolicy?.referenceAndBusinessType).split(' ')[0] || '';
        const effectDate = state.selectedPolicy?.effectDate || '';
        const identityTokens = normalize(state.selectedPolicy?.identity).split(' ').filter(token => token.length >= 3);
        const title = documentInfo.name || '';
        const fileContent = `${documentInfo.contentRead?.extractedText || ''}\n${documentInfo.contentRead?.ocrText || ''}`;
        const content = `${title}\n${fileContent}`;
        const normalizedContent = normalize(content);
        const extractYears = value => Array.from(String(value || '').matchAll(/(20\d{2})/g)).map(match => Number(match[1]))
            .filter((value, index, all) => all.indexOf(value) === index);
        const titleYears = extractYears(title);
        const contentYears = extractYears(fileContent);
        const years = [...titleYears, ...contentYears].filter((value, index, all) => all.indexOf(value) === index);
        const uploadYear = Number(documentInfo.uploadDate.match(/\d{2}\/\d{2}\/(20\d{2})/)?.[1] || 0);
        const referenceMatch = Boolean(policyReference && normalize(content).includes(normalize(policyReference)));
        const effectDateMatch = Boolean(effectDate && content.includes(effectDate));
        const identityMatch = identityTokens.length >= 2 && identityTokens.every(token => normalizedContent.includes(token));
        const titleTargetYear = Boolean(targetYear && titleYears.includes(targetYear));
        const contentTargetYear = Boolean(targetYear && contentYears.includes(targetYear));
        const titleHasOtherYears = titleYears.length > 0 && !titleTargetYear;
        const contentHasOtherYears = contentYears.length > 0 && !contentTargetYear;
        const parseDate = value => {
            const match = String(value || '').match(/(\d{1,2})\/(\d{1,2})\/(20\d{2})/);
            return match ? new Date(Number(match[3]), Number(match[2]) - 1, Number(match[1])) : null;
        };
        const uploadDate = parseDate(documentInfo.uploadDate);
        const selectedEffectDate = parseDate(effectDate);
        const uploadDistanceDays = uploadDate && selectedEffectDate
            ? Math.round((uploadDate.getTime() - selectedEffectDate.getTime()) / 86400000)
            : null;
        // Une campagne peut commencer quelques mois avant l'effet. Une pièce beaucoup plus ancienne
        // n'est retenue que si son titre, la référence ou la date d'effet la rattache explicitement.
        const uploadNearPeriod = uploadDistanceDays !== null
            ? uploadDistanceDays >= -240 && uploadDistanceDays <= 550
            : Boolean(targetYear && uploadYear >= targetYear - 1 && uploadYear <= targetYear + 1);

        let level = 'possible';
        const reasons = [];
        if (referenceMatch) reasons.push('reference-contrat');
        if (effectDateMatch) reasons.push('date-effet');
        if (identityMatch) reasons.push('identite');
        if (titleTargetYear) reasons.push('annee-declaration-titre');
        else if (contentTargetYear) reasons.push('annee-declaration-contenu');
        if (uploadNearPeriod) reasons.push('date-ajout-proche');
        if (referenceMatch || effectDateMatch) level = 'confirmed';
        else if (titleTargetYear && identityMatch) level = 'confirmed';
        else if (titleTargetYear) level = 'likely';
        // Un millésime explicite dans le titre prime sur une date d'expiration 2026 trouvée dans un ancien PDF.
        else if (titleHasOtherYears) level = 'unlikely';
        else if (contentTargetYear && identityMatch && uploadNearPeriod) level = 'likely';
        else if (contentHasOtherYears || !uploadNearPeriod) level = 'unlikely';
        else if (identityMatch) level = 'likely';

        documentInfo.declarationRelevance = {
            level,
            reasons,
            years,
            titleYears,
            contentYears,
            uploadYear,
            uploadDistanceDays,
        };
        return documentInfo;
    }

    function documentMatchesSelectedPolicy(documentInfo, state) {
        const reasons = documentInfo.declarationRelevance?.reasons || [];
        if (reasons.includes('reference-contrat') || reasons.includes('identite')) return true;
        const identityTokens = normalizePersonName(state.selectedPolicy?.identity).split(' ').filter(token => token.length >= 2);
        const content = normalizePersonName(`${documentInfo.name || ''}\n${documentContent(documentInfo)}`);
        return identityTokens.length >= 2 && identityTokens.every(token => content.includes(token));
    }

    function selectContractCoverageDocuments(relevantDocuments, state) {
        const coverageDocuments = relevantDocuments.filter(documentInfo => documentInfo.isCoverageCandidate);
        const anchors = coverageDocuments.filter(documentInfo => documentMatchesSelectedPolicy(documentInfo, state));
        const anchorPersons = anchors.flatMap(documentInfo => documentInfo.contentRead?.coveredPersons || []);

        return coverageDocuments.filter(documentInfo => {
            if (anchors.includes(documentInfo)) {
                documentInfo.contractAttribution = { level: 'anchor', reason: 'identite-ou-reference-du-souscripteur' };
                return true;
            }
            const persons = documentInfo.contentRead?.coveredPersons || [];
            const supported = persons.length > 0 && persons.some(person =>
                anchorPersons.some(anchorPerson => personNamesProbablyMatch(anchorPerson, person))
            );
            documentInfo.contractAttribution = supported
                ? { level: 'supported-person', reason: 'personne-retrouvee-dans-un-document-ancre' }
                : { level: 'excluded', reason: 'aucun-lien-avec-le-contrat-selectionne' };
            return supported;
        });
    }

    function paymentDocumentContradictsPolicy(documentInfo, state) {
        const persons = documentInfo.contentRead?.coveredPersons || [];
        if (!persons.length || documentMatchesSelectedPolicy(documentInfo, state)) return false;
        const subscriber = { name: state.selectedPolicy?.identity || '', birthDate: state.client?.birthDate || '' };
        return !persons.some(person => personNamesProbablyMatch(person, subscriber));
    }

    function paymentContactMatch(first, second) {
        const firstPhone = String(first.client?.phone || '').replace(/\D/g, '');
        const secondPhone = String(second.client?.phone || '').replace(/\D/g, '');
        if (firstPhone.length >= 10 && firstPhone === secondPhone) return true;
        const firstEmail = clean(first.client?.email).toLowerCase();
        const secondEmail = clean(second.client?.email).toLowerCase();
        const ignoredEmails = /@(ltoa-assurances\.fr|ciel-assurances\.fr)$/i;
        return Boolean(firstEmail && firstEmail === secondEmail && !ignoredEmails.test(firstEmail));
    }

    function relatedPaymentGroup(origin, results) {
        const group = new Set([origin]);
        const queue = [origin];
        while (queue.length) {
            const current = queue.shift();
            for (const candidate of results) {
                if (group.has(candidate) || !paymentContactMatch(current, candidate)) continue;
                group.add(candidate);
                queue.push(candidate);
            }
        }
        return Array.from(group);
    }

    function extractExpectedPaymentAmounts(result) {
        const amounts = [];
        const amountPattern = /(\d{1,5}(?:[ .]\d{3})*(?:[,.]\d{2})?)\s*(?:€|EUR)/gi;
        const paymentLabel = /\b(cotisation|prime|tarif|total\s+[àa]\s+payer|montant\s+[àa]\s+payer|net\s+[àa]\s+payer|annuel(?:le)?)\b/i;
        for (const documentInfo of contractDocumentsForResult(result)) {
            const content = documentContent(documentInfo);
            for (const match of content.matchAll(amountPattern)) {
                const context = content.slice(Math.max(0, match.index - 90), match.index + match[0].length + 90);
                if (!paymentLabel.test(context)) continue;
                const value = Number(match[1].replace(/\s/g, '').replace(',', '.'));
                if (Number.isFinite(value) && value > 0) amounts.push(value);
            }
        }
        return amounts.filter((value, index, all) =>
            all.findIndex(candidate => Math.abs(candidate - value) < 0.01) === index
        );
    }

    function paymentAmountMatches(paymentAmounts, expectedAmounts) {
        if (!paymentAmounts.length || !expectedAmounts.length) return false;
        return paymentAmounts.some(paid => expectedAmounts.some(expected => Math.abs(paid - expected) < 0.02));
    }

    function assessPaymentDocument(documentInfo, result) {
        const facts = documentInfo.paymentFacts || extractPaymentFacts(documentInfo);
        const expectedAmounts = extractExpectedPaymentAmounts(result);
        const amountMatches = paymentAmountMatches(facts.amounts || [], expectedAmounts);
        const methods = facts.methods || [];
        const isPhysical = methods.some(method => ['cheque', 'especes'].includes(method));
        const isElectronic = methods.some(method =>
            ['paypal', 'stripe', 'virement', 'carte-bancaire', 'prelevement'].includes(method)
        );
        const isScannedEvidence = facts.isImage || facts.readStatus === 'ocr-read';
        const reasons = [];

        if (facts.titleNegative || facts.contentNegative) {
            reasons.push(facts.titleNegative
                ? 'Le titre de la pièce indique un paiement problématique.'
                : 'Le document contient une mention d’attente, de rejet, d’annulation ou de remboursement.');
            return { status: 'non-justified', reasons, expectedAmounts, amountMatches };
        }
        if (!facts.hasReadableContent && !facts.isImage) {
            reasons.push('La pièce n’a pas pu être lue.');
            return { status: 'needs-review', reasons, expectedAmounts, amountMatches };
        }
        if (facts.readStatus === 'ocr-partial') {
            reasons.push('Le PDF scanné dépasse 30 pages : la lecture OCR est partielle.');
            return { status: 'needs-review', reasons, expectedAmounts, amountMatches };
        }
        if (!expectedAmounts.length) reasons.push('Montant contractuel attendu non extrait automatiquement.');
        else if (!facts.amounts?.length) reasons.push('Aucun montant lisible dans le justificatif.');
        else if (!amountMatches) {
            reasons.push(`Montant du justificatif (${facts.amounts.join(' / ')} €) différent du montant attendu (${expectedAmounts.join(' / ')} €).`);
        }
        if (isElectronic && !facts.hasFinalStatus) {
            reasons.push('Aucun statut définitif de paiement n’est lisible dans la pièce électronique.');
        }
        if (!methods.length) reasons.push('Moyen de paiement non reconnu dans le titre ou le contenu.');

        const physicalConfirmed = (isPhysical || (isScannedEvidence && facts.titlePaymentSignal && !isElectronic))
            && facts.titlePaymentSignal
            && amountMatches
            && !facts.titleNegative
            && !facts.contentNegative;
        const electronicConfirmed = isElectronic
            && facts.hasFinalStatus
            && amountMatches
            && !facts.titleNegative
            && !facts.contentNegative;
        const acquittedConfirmed = methods.includes('facture-acquittee')
            && facts.hasFinalStatus
            && amountMatches;

        return {
            status: physicalConfirmed || electronicConfirmed || acquittedConfirmed ? 'confirmed' : 'needs-review',
            reasons,
            expectedAmounts,
            amountMatches,
            method: methods.join(', '),
        };
    }

    function assessBatchPayments(results) {
        return results.map(result => {
            const directDocuments = result.ged?.paymentCandidates || [];
            if (directDocuments.length) {
                const documentAssessments = directDocuments.map(documentInfo => ({
                    documentId: documentInfo.documentId,
                    name: documentInfo.name,
                    uploadDate: documentInfo.uploadDate,
                    paymentFacts: documentInfo.paymentFacts || null,
                    decision: assessPaymentDocument(documentInfo, result),
                }));
                const confirmed = documentAssessments.find(item => item.decision.status === 'confirmed');
                const uncertain = documentAssessments.find(item => item.decision.status === 'needs-review');
                const rejected = documentAssessments.every(item => item.decision.status === 'non-justified');
                result.paymentAssessment = {
                    status: confirmed ? 'confirmed' : (uncertain ? 'needs-review' : (rejected ? 'non-justified' : 'needs-review')),
                    directDocuments: documentAssessments,
                    selectedDocumentId: confirmed?.documentId || uncertain?.documentId || documentAssessments[0]?.documentId || '',
                    note: confirmed
                        ? 'Paiement confirmé automatiquement par le contenu de la pièce.'
                        : documentAssessments.flatMap(item => item.decision.reasons || []).filter(Boolean).join(' '),
                };
                return result;
            }
            const paymentGroup = relatedPaymentGroup(result, results);
            const relatedEvidence = paymentGroup.filter(candidate => candidate !== result)
                .flatMap(candidate => (candidate.ged?.paymentCandidates || []).map(documentInfo => ({
                    sourcePolicyId: candidate.policy?.policyId,
                    sourceIdentity: candidate.policy?.identity,
                    documentId: documentInfo.documentId,
                    name: documentInfo.name,
                    uploadDate: documentInfo.uploadDate,
                    paymentFacts: documentInfo.paymentFacts || null,
                })));
            if (relatedEvidence.length) {
                result.paymentAssessment = {
                    status: 'needs-review',
                    relatedPolicyIds: paymentGroup.map(item => item.policy?.policyId).filter(Boolean),
                    relatedEvidence,
                    note: 'Paiement familial potentiel : le contenu et le montant doivent couvrir les contrats regroupés.',
                };
                return result;
            }
            result.paymentAssessment = {
                status: 'non-justified',
                relatedEvidence: [],
                note: normalize(result.policy?.state).includes('en cours')
                    ? 'Contrat en cours, mais aucun justificatif de paiement n’a été retrouvé dans la GED.'
                    : 'Aucun justificatif de paiement confirmé.',
            };
            return result;
        });
    }

    async function captureGedAndContinue() {
        const state = loadState();
        if (!state || state.status !== 'opening-ged') return;
        const documents = await collectAllGedDocuments(state);
        documents.forEach(documentInfo => annotateDocumentRelevance(documentInfo, state));
        await readRelevantDocuments(documents.filter(documentInfo => documentInfo.declarationRelevance?.level !== 'unlikely'));
        documents.forEach(documentInfo => annotateDocumentRelevance(documentInfo, state));
        const relevantDocuments = documents.filter(documentInfo => documentInfo.declarationRelevance?.level !== 'unlikely');
        const contractDocuments = selectContractCoverageDocuments(relevantDocuments, state);
        const excludedCoverageDocuments = relevantDocuments.filter(documentInfo =>
            documentInfo.isCoverageCandidate && documentInfo.contractAttribution?.level === 'excluded'
        );
        const extractedPersons = contractDocuments.flatMap(documentInfo =>
            (documentInfo.contentRead?.coveredPersons || []).map(person => ({
                ...person,
                sourceDocumentId: documentInfo.documentId,
                sourceDocumentName: documentInfo.name,
            }))
        ).filter((person, index, all) =>
            all.findIndex(candidate => normalize(candidate.name) === normalize(person.name) && candidate.birthDate === person.birthDate) === index
        );
        const personsWithSubscriberEvidence = addSubscriberFromDocumentEvidence(
            extractedPersons,
            contractDocuments,
            state
        );
        const coveredPersons = deduplicateCoveredPersons(reconcileCoveredPersons(personsWithSubscriberEvidence, state));
        const paymentCandidates = relevantDocuments.filter(documentInfo =>
            documentInfo.isPaymentCandidate && !paymentDocumentContradictsPolicy(documentInfo, state)
        );
        state.ged = {
            url: location.href,
            clientId: location.pathname.match(/\/Client\/(\d+)/)?.[1] || state.client?.clientId || '',
            contractOptions: Array.from(document.querySelectorAll('select[name="ged_filters[client_subentity]"] option')).map(option => ({
                label: clean(option.textContent),
                value: option.value,
            })),
            documents,
            relevantDocuments,
            contractDocuments,
            excludedCoverageDocuments: excludedCoverageDocuments.map(documentInfo => ({
                documentId: documentInfo.documentId,
                name: documentInfo.name,
                reason: documentInfo.contractAttribution?.reason || '',
            })),
            documentsScope: state.gedDocumentsScope || 'unknown',
            pagesScanned: state.gedPagesScanned || 1,
            documentsScanned: state.gedDocumentsScanned || documents.length,
            paymentCandidates,
            quoteCandidates: contractDocuments.filter(document => document.isQuoteCandidate),
            coveredPersons,
        };
        if (!contractDocuments.length) {
            state.warnings.push(`Aucun document contractuel détecté pour le contrat ${state.selectedPolicy?.policyId || ''}.`);
        }
        if (!state.ged.coveredPersons.length) {
            state.warnings.push(`Aucune personne assurée extraite des PDF pour le contrat ${state.selectedPolicy?.policyId || ''}.`);
        }
        for (const person of coveredPersons.filter(item => item.conflictingBirthDates?.length > 1)) {
            state.warnings.push(`Lectures OCR divergentes de la date de naissance pour ${person.name} (${person.conflictingBirthDates.join(' / ')}) : ${person.birthDate} retenue d’après la pièce la plus fiable.`);
        }
        for (const documentInfo of excludedCoverageDocuments.filter(item => item.contentRead?.coveredPersons?.length)) {
            state.warnings.push(`Pièce « ${documentInfo.name} » écartée du contrat ${state.selectedPolicy?.policyId || ''} : les personnes lues ne correspondent pas au souscripteur.`);
        }
        for (const documentInfo of relevantDocuments.filter(item => ['ocr-empty', 'ocr-unavailable', 'unsupported-format'].includes(item.contentRead?.status))) {
            state.warnings.push(`OCR sans résultat pour « ${documentInfo.name} » (contrat ${state.selectedPolicy?.policyId || ''}).`);
        }
        for (const documentInfo of relevantDocuments.filter(item => item.contentRead?.status === 'ocr-partial')) {
            state.warnings.push(`OCR partiel pour « ${documentInfo.name} » : ${documentInfo.contentRead?.ocrPageCount || 30} page(s) lue(s) sur ${documentInfo.contentRead?.ocrTotalPages || 'davantage'}.`);
        }
        for (const documentInfo of relevantDocuments.filter(item => ['read-error', 'reader-unavailable', 'ocr-error'].includes(item.contentRead?.status))) {
            state.warnings.push(`Lecture impossible pour « ${documentInfo.name} » : ${documentInfo.contentRead.error || documentInfo.contentRead.status}.`);
        }
        state.results = Array.isArray(state.results) ? state.results : [];
        const currentResult = {
            policy: state.selectedPolicy,
            client: state.client,
            ged: state.ged,
        };
        currentResult.declarationFacts = extractAgisContractFacts(currentResult);
        if (!currentResult.declarationFacts.country) {
            state.warnings.push(`Pays d'inhumation introuvable dans les pièces rattachées au contrat ${state.selectedPolicy?.policyId || ''}.`);
        }
        state.results.push(currentResult);

        const total = state.policies?.length || 1;
        const nextIndex = (state.currentIndex || 0) + 1;
        if (nextIndex < total) {
            state.currentIndex = nextIndex;
            state.status = 'returning-list';
            saveState(state);
            setMessage(`Contrat ${nextIndex}/${total} terminé. Retour à la liste…`);
            setTimeout(() => { location.href = state.listUrl || `${location.origin}${LIST_PATH}`; }, 500);
            return;
        }

        state.results = assessBatchPayments(state.results);
        for (const result of state.results.filter(item => item.paymentAssessment?.status === 'needs-review')) {
            state.warnings.push(`Paiement à contrôler humainement pour le contrat ${result.policy?.policyId || ''} : ${result.paymentAssessment?.note || 'preuve incertaine'}`);
        }
        for (const result of state.results.filter(item => item.paymentAssessment?.status === 'non-justified')) {
            state.warnings.push(`Contrat exclu par défaut faute de paiement justifié : ${result.policy?.policyId || ''}.`);
        }
        state.status = 'batch-completed';
        state.completedAt = new Date().toISOString();
        state.analysisCompletedAt = state.completedAt;
        saveState(state);
        setMessage(`Lot terminé : ${state.results.length} contrat(s) traité(s). Revenez sur la fiche AGIS pour vérifier puis exporter.`);
        setTimeout(() => { location.href = AGIS_COMPANY_URL; }, 700);
    }

    function downloadResult() {
        const state = loadState();
        if (!state) throw new Error('Aucun test à exporter.');
        const blob = new Blob([JSON.stringify(state, null, 2)], { type: 'application/json;charset=utf-8' });
        const url = URL.createObjectURL(blob);
        const anchor = document.createElement('a');
        anchor.href = url;
        anchor.download = `test-agent-agis-${state.month || 'periode'}-${new Date().toISOString().replace(/[:.]/g, '-')}.json`;
        document.body.appendChild(anchor);
        anchor.click();
        anchor.remove();
        state.jsonExportedAt = new Date().toISOString();
        saveState(state);
        setTimeout(() => URL.revokeObjectURL(url), 2000);
    }

    const AGIS_HEADERS = [
        'Code partenaire',
        'Type adhésion',
        'N Contrat',
        "Type d'enregistrement",
        'Statut du contrat',
        'Option',
        'Zone',
        "Pays d'inhumation",
        "Date d'effet",
        "Date d'expiration prévue",
        'Formule',
        'Prime CAD OK',
        'Civilité',
        'Nom',
        'Prénom',
        'Date de naissance',
        'Age assureprincipal',
        'Tranche age',
        'Adresse',
        'Adresse (suite)',
        'Code postal',
        'Ville',
    ];

    function parseFrenchDateValue(value) {
        const match = String(value || '').match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
        if (!match) return null;
        return new Date(Number(match[3]), Number(match[2]) - 1, Number(match[1]), 12, 0, 0);
    }

    function splitPolicyIdentity(identity) {
        const value = clean(identity, 180);
        if (value.includes(',')) {
            const [lastName, ...rest] = value.split(',');
            return { lastName: clean(lastName), firstName: clean(rest.join(' ')) };
        }
        const parts = value.split(' ').filter(Boolean);
        return { lastName: parts.shift() || '', firstName: parts.join(' ') };
    }

    function personMatchesSubscriber(person, result) {
        const subscriberTokens = normalizePersonName(result.policy?.identity).split(' ').filter(token => token.length >= 2);
        const personTokens = normalizePersonName(person.name).split(' ').filter(token => token.length >= 2);
        // Le rôle contractuel dépend de l'identité, jamais d'une date OCR.
        // Une lecture erronée de la naissance du souscripteur ne doit pas créer
        // une seconde ligne « bénéficiaire » pour la même personne.
        return subscriberTokens.length >= 2 && subscriberTokens.every(token => personTokens.includes(token));
    }

    function splitCoveredPerson(person, subscriberLastName) {
        if (person.lastName && person.firstName) {
            return { lastName: clean(person.lastName), firstName: clean(person.firstName) };
        }
        const value = clean(person.name, 180);
        if (value.includes(',')) return splitPolicyIdentity(value);
        const parts = value.split(' ').filter(Boolean);
        const subscriberLastTokens = normalizePersonName(subscriberLastName).split(' ').filter(Boolean);
        const normalizedParts = parts.map(normalizePersonName);
        const matchingLastName = subscriberLastTokens.length
            && subscriberLastTokens.every(token => normalizedParts.includes(token));
        if (matchingLastName) {
            const firstNameParts = parts.filter(part => !subscriberLastTokens.includes(normalizePersonName(part)));
            return { lastName: subscriberLastName, firstName: clean(firstNameParts.join(' ')) };
        }
        return { lastName: parts.shift() || '', firstName: parts.join(' ') };
    }

    function documentContent(documentInfo) {
        return `${documentInfo.contentRead?.extractedText || ''}\n${documentInfo.contentRead?.ocrText || ''}`;
    }

    function contractDocumentsForResult(result) {
        if (result.ged?.contractDocuments?.length) return result.ged.contractDocuments;
        const stateLike = { selectedPolicy: result.policy, client: result.client };
        return selectContractCoverageDocuments(result.ged?.relevantDocuments || [], stateLike);
    }

    function extractExplicitContractSubscriber(result) {
        const heading = normalize(AGIS_CP_TEMPLATE.subscriberHeading);
        for (const documentInfo of contractDocumentsForResult(result)) {
            const rawContent = documentContent(documentInfo);
            if (!normalize(rawContent).includes(heading)) continue;
            const match = rawContent.match(
                /informations?\s+du\s+souscripteur[\s\S]{0,500}?\bnom\s*:\s*([^\n\r]{2,100}?)\s+pr[eé]nom\s*:\s*([^\n\r]{2,100}?)(?=\s+(?:adresse|code\s+postal|ville|pays|mobile|fixe)\b|\r?\n|$)/i
            );
            if (!match) continue;
            const lastName = clean(match[1], 100).replace(/^[,.;:\s]+|[,.;:\s]+$/g, '');
            const firstName = clean(match[2], 100).replace(/^[,.;:\s]+|[,.;:\s]+$/g, '');
            if (!lastName || !firstName) continue;
            return {
                lastName,
                firstName,
                name: `${lastName} ${firstName}`,
                documentId: documentInfo.documentId || '',
                documentName: documentInfo.name || '',
            };
        }
        return null;
    }

    function namesMatch(left, right) {
        const leftTokens = normalizePersonName(left).split(' ').filter(token => token.length >= 2);
        const rightTokens = normalizePersonName(right).split(' ').filter(token => token.length >= 2);
        return leftTokens.length >= 2
            && rightTokens.length >= 2
            && leftTokens.every(token => rightTokens.includes(token))
            && rightTokens.every(token => leftTokens.includes(token));
    }

    function sanitizeAddressLine(value) {
        const cleaned = clean(value, 250);
        return /^\d{1,3}$/.test(cleaned) ? '' : cleaned;
    }

    function coveredPersonsForResult(result) {
        const stateLike = { selectedPolicy: result.policy, client: result.client };
        const persons = contractDocumentsForResult(result).flatMap(documentInfo =>
            (documentInfo.contentRead?.coveredPersons || []).map(person => ({
                ...person,
                sourceDocumentId: documentInfo.documentId,
                sourceDocumentName: documentInfo.name,
            }))
        );
        return deduplicateCoveredPersons(reconcileCoveredPersons(persons, stateLike));
    }

    function extractDestinationCountry(text) {
        const normalized = normalize(text);
        const labels = [
            'pays designe en cas de rapatriement',
            'pays d inhumation',
            'pays de destination',
            'destination du rapatriement',
        ];
        const countries = [
            ['algerie', 'Algérie'], ['maroc', 'Maroc'], ['tunisie', 'Tunisie'],
            ['senegal', 'Sénégal'], ['mali', 'Mali'], ['mauritanie', 'Mauritanie'],
            ['turquie', 'Turquie'], ['france', 'France'],
        ];
        for (const label of labels) {
            const index = normalized.indexOf(label);
            if (index < 0) continue;
            const context = normalized.slice(index + label.length, index + label.length + 100);
            const country = countries.find(([needle]) => new RegExp(`\\b${needle}\\b`).test(context));
            if (country) return country[1];
        }
        return '';
    }

    function extractAgisContractFacts(result) {
        const effectDate = result.policy?.effectDate || '';
        const documents = contractDocumentsForResult(result);
        const facts = {
            type: '', typeScore: 0,
            country: '', countryScore: 0,
            personCount: null, personCountScore: 0,
            expiration: result.policy?.state?.match(/expire le\s+(\d{1,2}\/\d{1,2}\/\d{4})/i)?.[1] || '',
            expirationScore: 1,
            sources: [],
        };

        for (const documentInfo of documents) {
            const rawContent = documentContent(documentInfo);
            const normalizedContent = normalize(rawContent);
            const lines = rawContent.split(/\r?\n/).map(line => clean(line, 160)).filter(Boolean);
            let documentUsed = false;

            const explicitType = normalizedContent.match(/\btype de contrat\s+(individuel|famille)\b/)?.[1] || '';
            const profileType = normalizedContent.match(/\bprofile\s+(individuel|famille)\b/)?.[1] || '';
            const typeValue = explicitType || profileType;
            const typeScore = explicitType ? 3 : (profileType ? 1 : 0);
            if (typeScore > facts.typeScore) {
                facts.type = typeValue.charAt(0).toUpperCase() + typeValue.slice(1);
                facts.typeScore = typeScore;
                documentUsed = true;
            }

            const explicitCountry = extractDestinationCountry(rawContent);
            if (explicitCountry && facts.countryScore < 3) {
                facts.country = explicitCountry;
                facts.countryScore = 3;
                documentUsed = true;
            }

            const personCountMatch = normalizedContent.match(
                /\bnombre de personnes designees au contrat d assistance\s+(\d{1,2})\b/
            );
            if (personCountMatch && facts.personCountScore < 3) {
                facts.personCount = Number(personCountMatch[1]);
                facts.personCountScore = 3;
                documentUsed = true;
            }

            const periodMatch = rawContent.match(/p[eé]riode\s+de\s+garantie[^\d]{0,30}(\d{1,2}\/\d{1,2}\/\d{4})[^\d]{0,30}(\d{1,2}\/\d{1,2}\/\d{4})/i);
            if (periodMatch && periodMatch[1] === effectDate && facts.expirationScore < 3) {
                facts.expiration = periodMatch[2];
                facts.expirationScore = 3;
                documentUsed = true;
            }

            const effectIndex = lines.findIndex(line => line === effectDate);
            if (effectIndex >= 0) {
                let typeIndex = -1;
                for (let index = effectIndex - 1; index >= Math.max(0, effectIndex - 8); index -= 1) {
                    if (/^(individuel|famille)$/i.test(lines[index])) {
                        typeIndex = index;
                        break;
                    }
                }
                if (typeIndex >= 0 && facts.typeScore < 2) {
                    facts.type = clean(lines[typeIndex]);
                    facts.typeScore = 2;
                    documentUsed = true;
                }
                const structuredCountry = typeIndex >= 0 && typeIndex + 1 < effectIndex ? lines[typeIndex + 1] : '';
                if (structuredCountry && !/^\d+$/.test(structuredCountry) && facts.countryScore < 2) {
                    facts.country = clean(structuredCountry);
                    facts.countryScore = 2;
                    documentUsed = true;
                }
                const structuredExpiration = lines.slice(effectIndex + 1, effectIndex + 5)
                    .find(line => /^\d{1,2}\/\d{1,2}\/\d{4}$/.test(line)) || '';
                if (structuredExpiration && facts.expirationScore < 2) {
                    facts.expiration = structuredExpiration;
                    facts.expirationScore = 2;
                    documentUsed = true;
                }
            }
            if (documentUsed) facts.sources.push({ documentId: documentInfo.documentId, name: documentInfo.name });
        }

        return {
            type: facts.type,
            country: facts.country,
            personCount: facts.personCount,
            expiration: facts.expiration,
            sources: facts.sources.filter((source, index, all) =>
                all.findIndex(candidate => candidate.documentId === source.documentId) === index
            ),
        };
    }

    function detectPersonCivility(person, result) {
        const nameTokens = normalizePersonName(person.name).split(' ').filter(token => token.length >= 2);
        if (nameTokens.length < 2) return '';
        const documents = contractDocumentsForResult(result);
        for (const documentInfo of documents) {
            const content = normalizePersonName(documentContent(documentInfo));
            let position = -1;
            for (const token of nameTokens) {
                const found = content.indexOf(token);
                if (found >= 0 && (position < 0 || found < position)) position = found;
            }
            if (position < 0) continue;
            const context = content.slice(Math.max(0, position - 35), position + 160);
            if (!nameTokens.every(token => context.includes(token))) continue;
            const civility = normalizeCivility(context);
            if (civility) return civility;
        }
        return '';
    }

    function personSourceUrls(state, result, lastName, firstName) {
        const wantedTokens = normalizePersonName(`${lastName} ${firstName}`).split(' ').filter(token => token.length >= 2);
        const matchesLabel = value => {
            const candidate = normalizePersonName(value);
            return wantedTokens.length > 0 && wantedTokens.every(token => candidate.includes(token));
        };

        const currentIdentity = result.policy?.identity || result.client?.title || '';
        if (matchesLabel(currentIdentity)) {
            const clientId = result.client?.clientId || result.ged?.clientId || '';
            return {
                clientId,
                clientUrl: result.client?.url || (clientId ? `${location.origin}${CLIENT_PATH}?id=${clientId}` : ''),
                gedUrl: result.ged?.url || (clientId ? `${location.origin}${GED_PREFIX}${clientId}` : ''),
            };
        }

        const linked = (result.client?.linkedClients || []).find(item => matchesLabel(item.label));
        if (linked) return linked;

        const relatedResult = (state.results || []).find(candidate =>
            matchesLabel(candidate.policy?.identity || candidate.client?.title || '')
        );
        if (relatedResult) {
            const clientId = relatedResult.client?.clientId || relatedResult.ged?.clientId || '';
            return {
                clientId,
                clientUrl: relatedResult.client?.url || (clientId ? `${location.origin}${CLIENT_PATH}?id=${clientId}` : ''),
                gedUrl: relatedResult.ged?.url || (clientId ? `${location.origin}${GED_PREFIX}${clientId}` : ''),
            };
        }

        const pageText = clean(result.client?.pageTextSample || '', 12000);
        const normalizedPageText = normalizePersonName(pageText);
        const clientNumberPattern = /n[°ºo]\s*(\d{2,})/gi;
        let match;
        const candidates = [];
        while ((match = clientNumberPattern.exec(pageText))) {
            const context = pageText.slice(Math.max(0, match.index - 180), match.index + match[0].length + 30);
            if (!matchesLabel(context)) continue;
            const normalizedPrefix = normalizePersonName(pageText.slice(0, match.index));
            const distances = wantedTokens.map(token => {
                const position = normalizedPrefix.lastIndexOf(token);
                return position < 0 ? Number.POSITIVE_INFINITY : normalizedPageText.length - position;
            });
            candidates.push({ clientId: match[1], score: Math.max(...distances) });
        }
        candidates.sort((left, right) => left.score - right.score);
        if (candidates.length) {
            const clientId = candidates[0].clientId;
            return {
                clientId,
                clientUrl: `${location.origin}${CLIENT_PATH}?id=${clientId}`,
                gedUrl: `${location.origin}${GED_PREFIX}${clientId}`,
            };
        }
        return { clientId: '', clientUrl: '', gedUrl: '' };
    }

    function buildAgisRowModels(state) {
        const review = state.review || {};
        const rowOverrides = review.overrides || {};
        const contractOverrides = review.contractOverrides || {};
        const models = (state.results || []).flatMap((result, contractIndex) => {
            const facts = extractAgisContractFacts(result);
            const subscriber = splitPolicyIdentity(result.policy?.identity || '');
            const explicitSubscriber = extractExplicitContractSubscriber(result);
            const extractedPersons = coveredPersonsForResult(result);
            const beneficiaries = extractedPersons.filter(person => !personMatchesSubscriber(person, result));
            const policyId = result.policy?.policyId || String(contractIndex + 1);
            const contractOverride = contractOverrides[policyId] || {};
            const derivedAdhesionType = /^famille$/i.test(facts.type) || beneficiaries.length ? 'Famille' : 'Individuel';
            const adhesionType = contractOverride.adhesionType || derivedAdhesionType;
            const country = contractOverride.country !== undefined ? contractOverride.country : (facts.country || '');
            const detectedPaymentStatus = ['confirmed', 'needs-review', 'non-justified'].includes(result.paymentAssessment?.status)
                ? result.paymentAssessment.status
                : 'needs-review';
            const paymentStatus = ['confirmed', 'needs-review', 'non-justified'].includes(contractOverride.paymentStatus)
                ? contractOverride.paymentStatus
                : detectedPaymentStatus;
            const excludedDocuments = (result.ged?.relevantDocuments || []).filter(documentInfo =>
                documentInfo.isCoverageCandidate
                && documentInfo.contractAttribution?.level === 'excluded'
                && documentInfo.contentRead?.coveredPersons?.length
            );
            const people = [{
                rowId: `${policyId}:subscriber`,
                role: 'Souscripteur',
                lastName: subscriber.lastName,
                firstName: subscriber.firstName,
                birthDate: result.client?.birthDate || '',
                civility: result.client?.civility || '',
                sourcePerson: null,
            }, ...beneficiaries.map((person, index) => ({
                rowId: `${policyId}:beneficiary:${index + 1}`,
                role: 'Bénéficiaire',
                ...splitCoveredPerson(person, subscriber.lastName),
                birthDate: person.birthDate || '',
                civility: detectPersonCivility(person, result),
                sourcePerson: person,
            }))];

            return people.map(person => {
                const override = rowOverrides[person.rowId] || {};
                const personUrls = personSourceUrls(state, result, person.lastName, person.firstName);
                const model = {
                    rowId: person.rowId,
                    policyId,
                    contractNumber: contractIndex + 1,
                    role: override.role === 'Bénéficiaire' ? 'Bénéficiaire'
                        : (override.role === 'Souscripteur' ? 'Souscripteur' : person.role),
                    adhesionType,
                    country,
                    paymentStatus,
                    detectedPaymentStatus,
                    paymentNote: result.paymentAssessment?.note || '',
                    declaredPersonCount: facts.personCount,
                    effectDate: result.policy?.effectDate || '',
                    expiration: facts.expiration || '',
                    civility: normalizeCivility(override.civility !== undefined ? override.civility : person.civility),
                    lastName: override.lastName !== undefined ? override.lastName : person.lastName,
                    firstName: override.firstName !== undefined ? override.firstName : person.firstName,
                    birthDate: override.birthDate !== undefined ? override.birthDate : person.birthDate,
                    address: result.client?.address || '',
                    address2: sanitizeAddressLine(result.client?.address2 || ''),
                    postalCode: result.client?.postalCode || '',
                    city: clean(result.client?.city || '').replace(/\s+FRANCE$/i, ''),
                    clientId: personUrls.clientId,
                    clientUrl: personUrls.clientUrl,
                    gedUrl: personUrls.gedUrl,
                    issues: [],
                    evidenceDocuments: [
                        ...(person.sourcePerson?.evidenceDocuments || []),
                        ...facts.sources.map(source => source.name),
                        ...(person.role === 'Souscripteur'
                            ? (result.paymentAssessment?.directDocuments || []).map(source => source.name)
                            : []),
                    ].filter(Boolean).filter((value, index, all) => all.indexOf(value) === index),
                };
                if (model.role === 'Souscripteur' && !model.country) {
                    model.issues.push("Pays d'inhumation introuvable");
                }
                if (!model.civility) {
                    model.issues.push('Civilité introuvable');
                }
                if (model.role === 'Souscripteur' && excludedDocuments.length) {
                    model.issues.push(`Pièce(s) d'un autre contrat écartée(s) : ${excludedDocuments.map(item => item.name).join(', ')}`);
                }
                if (model.role === 'Souscripteur' && result.ged?.documentsScope === 'all-client-documents-partial-pages') {
                    model.issues.push(`Parcours GED incomplet : ${result.ged?.pagesScanned || 1} page(s) et ${result.ged?.documentsScanned || 0} pièce(s) lues avant interruption`);
                }
                if (model.role === 'Souscripteur' && model.paymentStatus === 'needs-review') {
                    model.issues.push(`Paiement à vérifier : ${model.paymentNote || 'la preuve ne permet pas une confirmation automatique'}`);
                }
                if (model.role === 'Souscripteur' && model.paymentStatus === 'non-justified') {
                    model.issues.push(`Paiement non justifié — contrat exclu de l’Excel : ${model.paymentNote || 'aucune preuve recevable'}`);
                }
                if (person.role === 'Souscripteur' && explicitSubscriber
                    && !namesMatch(`${subscriber.lastName} ${subscriber.firstName}`, explicitSubscriber.name)) {
                    model.issues.push(
                        `Souscripteur du document différent : ${explicitSubscriber.firstName} ${explicitSubscriber.lastName}`
                    );
                    if (explicitSubscriber.documentName) model.evidenceDocuments.push(explicitSubscriber.documentName);
                }
                if (person.sourcePerson?.aliases?.length > 1) {
                    model.issues.push(`Variantes OCR regroupées : ${person.sourcePerson.aliases.join(' / ')}`);
                }
                if (person.sourcePerson?.conflictingBirthDates?.length > 1) {
                    model.issues.push(`Lectures OCR de naissance à vérifier : ${person.sourcePerson.conflictingBirthDates.join(' / ')}`);
                }
                return model;
            });
        });

        const contracts = new Map();
        for (const model of models) {
            if (!contracts.has(model.contractNumber)) contracts.set(model.contractNumber, []);
            contracts.get(model.contractNumber).push(model);
        }
        for (const [contractNumber, contractModels] of contracts) {
            const subscribers = contractModels.filter(model => model.role === 'Souscripteur');
            if (subscribers.length !== 1) {
                contractModels.forEach(model => model.issues.push(
                    `Structure invalide : contrat ${contractNumber} avec ${subscribers.length} souscripteur(s)`
                ));
            }
            if (contractModels[0]?.adhesionType === 'Famille' && contractModels.length < 2) {
                contractModels.forEach(model => model.issues.push('Contrat Famille sans bénéficiaire extrait'));
            }
            if (contractModels[0]?.adhesionType === 'Individuel' && contractModels.length > 1) {
                contractModels.forEach(model => model.issues.push('Plusieurs personnes trouvées sur un contrat Individuel'));
            }
            const declaredPersonCount = contractModels[0]?.declaredPersonCount;
            if (Number.isInteger(declaredPersonCount) && declaredPersonCount !== contractModels.length) {
                contractModels.forEach(model => model.issues.push(
                    `Le document indique ${declaredPersonCount} personne(s), mais ${contractModels.length} identité(s) sont listées`
                ));
            }
            const identityGroups = new Map();
            for (const model of contractModels) {
                const key = [...normalizePersonName(`${model.lastName} ${model.firstName}`).split(' ').filter(Boolean)]
                    .sort().join('|');
                if (!key) continue;
                if (!identityGroups.has(key)) identityGroups.set(key, []);
                identityGroups.get(key).push(model);
            }
            for (const duplicateModels of identityGroups.values()) {
                if (duplicateModels.length < 2) continue;
                duplicateModels.forEach(model => model.issues.push(
                    `Personne dupliquée ${duplicateModels.length} fois dans le contrat ${contractNumber}`
                ));
            }
        }

        const batchIdentities = new Map();
        for (const model of models) {
            const key = [
                ...normalizePersonName(`${model.lastName} ${model.firstName}`).split(' ').filter(Boolean),
            ].sort().join('|') + `|${model.birthDate || ''}`;
            if (!batchIdentities.has(key)) batchIdentities.set(key, []);
            batchIdentities.get(key).push(model);
        }
        for (const duplicateModels of batchIdentities.values()) {
            const contractNumbers = [...new Set(duplicateModels.map(model => model.contractNumber))];
            if (contractNumbers.length < 2) continue;
            duplicateModels.forEach(model => model.issues.push(
                `Même personne présente sur plusieurs contrats : ${contractNumbers.join(', ')}`
            ));
        }

        for (const model of models) {
            model.issues = model.issues.filter((value, index, all) => all.indexOf(value) === index);
            model.evidenceDocuments = model.evidenceDocuments.filter((value, index, all) => all.indexOf(value) === index);
            model.requiresReview = model.issues.length > 0;
            model.validated = (review.validatedRowIds || []).includes(model.rowId);
        }
        return models;
    }

    function buildAgisRows(state) {
        return buildAgisRowModels(state).filter(model => model.paymentStatus === 'confirmed').map(model => [
            'Avenir',
            model.adhesionType,
            model.contractNumber,
            model.role,
            'Renouvellement',
            1,
            'Zone1',
            model.country,
            parseFrenchDateValue(model.effectDate),
            parseFrenchDateValue(model.expiration),
            null,
            null,
            model.civility,
            model.lastName,
            model.firstName,
            parseFrenchDateValue(model.birthDate),
            null,
            null,
            model.address,
            model.address2,
            model.postalCode,
            model.city,
        ]);
    }

    function declarationFingerprint(state) {
        return JSON.stringify(buildAgisRowModels(state).map(model => ({
            rowId: model.rowId,
            policyId: model.policyId,
            contractNumber: model.contractNumber,
            role: model.role,
            adhesionType: model.adhesionType,
            country: model.country,
            paymentStatus: model.paymentStatus,
            civility: model.civility,
            lastName: model.lastName,
            firstName: model.firstName,
            birthDate: model.birthDate,
            effectDate: model.effectDate,
            expiration: model.expiration,
            address: model.address,
            address2: model.address2,
            postalCode: model.postalCode,
            city: model.city,
        })));
    }

    function unresolvedReviewRows(state) {
        return buildAgisRowModels(state).filter(model => model.requiresReview && !model.validated);
    }

    function unresolvedPaymentContracts(state) {
        return buildAgisRowModels(state).filter(model =>
            model.role === 'Souscripteur' && model.paymentStatus === 'needs-review'
        );
    }

    function escapeHtml(value) {
        return String(value ?? '')
            .replaceAll('&', '&amp;')
            .replaceAll('<', '&lt;')
            .replaceAll('>', '&gt;')
            .replaceAll('"', '&quot;')
            .replaceAll("'", '&#039;');
    }

    function ensureReviewState(state) {
        state.review = state.review || {};
        state.review.overrides = state.review.overrides || {};
        state.review.contractOverrides = state.review.contractOverrides || {};
        state.review.validatedRowIds = state.review.validatedRowIds || [];
        state.review.overviewConfirmed = Boolean(state.review.overviewConfirmed);
        return state.review;
    }

    function reviewLink(url, label) {
        if (!url) return `<span class="ltoa-link-missing" title="Relancez le traitement pour enregistrer ce lien">${escapeHtml(label)} indisponible</span>`;
        const icon = label === 'GED'
            ? '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M4.75 5.75A1.75 1.75 0 0 1 6.5 4h4.086c.464 0 .91.184 1.238.513l1.163 1.162c.328.329.773.513 1.237.513H17.5a1.75 1.75 0 0 1 1.75 1.75V17.5a1.75 1.75 0 0 1-1.75 1.75h-11A1.75 1.75 0 0 1 4.75 17.5V5.75Z"/></svg>'
            : '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M12 12a3.25 3.25 0 1 0 0-6.5 3.25 3.25 0 0 0 0 6.5Zm-6.25 7c.46-3.1 2.68-5 6.25-5s5.79 1.9 6.25 5"/></svg>';
        return `<a class="ltoa-source-link" href="${escapeHtml(url)}" target="_blank" rel="noopener noreferrer">${icon}<span>${escapeHtml(label)}</span><b>↗</b></a>`;
    }

    function segmentedChoice(fieldAttribute, fieldName, value, options, label, compact = false) {
        const buttons = options.map(option => {
            const active = option.value === value;
            return `<button type="button" data-choice-value="${escapeHtml(option.value)}" class="${active ? 'active' : ''}" aria-pressed="${active}">${escapeHtml(option.label)}</button>`;
        }).join('');
        return `<label class="ltoa-field ${compact ? 'ltoa-field-compact' : ''}">
            <span>${escapeHtml(label)}</span>
            <div class="ltoa-segmented" data-choice-group>
                <input type="hidden" ${fieldAttribute}="${escapeHtml(fieldName)}" value="${escapeHtml(value)}">
                ${buttons}
            </div>
        </label>`;
    }

    let activeReviewFilter = 'issues';

    function applyReviewFilter(filter) {
        const panel = root.querySelector('.ltoa-review-panel');
        if (!panel) return;
        activeReviewFilter = filter === 'all' ? 'all' : 'issues';
        panel.dataset.filter = activeReviewFilter;
        panel.querySelectorAll('[data-review-filter]').forEach(button => {
            const active = button.dataset.reviewFilter === activeReviewFilter;
            button.classList.toggle('active', active);
            button.setAttribute('aria-pressed', String(active));
        });
        let visibleCards = 0;
        panel.querySelectorAll('.ltoa-contract-card').forEach(card => {
            const visible = activeReviewFilter === 'all' || card.dataset.needsReview === 'true';
            card.hidden = !visible;
            card.style.display = visible ? '' : 'none';
            if (visible) visibleCards += 1;
        });
        const emptyState = panel.querySelector('.ltoa-review-empty-filter');
        if (emptyState) emptyState.hidden = activeReviewFilter !== 'issues' || visibleCards > 0;
    }

    function renderReviewPanel() {
        const state = loadState();
        if (!state?.results?.length) throw new Error('Aucun contrat traité à vérifier.');
        ensureReviewState(state);
        const models = buildAgisRowModels(state).sort((left, right) =>
            Number(right.requiresReview) - Number(left.requiresReview)
            || left.contractNumber - right.contractNumber
            || (left.role === 'Souscripteur' ? -1 : 1)
        );
        const reviewPanel = root.querySelector('.ltoa-review-panel');
        const content = reviewPanel.querySelector('.ltoa-review-content');
        const unresolved = models.filter(model => model.requiresReview && !model.validated);
        const issueRows = models.filter(model => model.requiresReview);
        const contractGroups = new Map();
        models.forEach(model => {
            if (!contractGroups.has(model.contractNumber)) contractGroups.set(model.contractNumber, []);
            contractGroups.get(model.contractNumber).push(model);
        });
        const contractsToReview = [...contractGroups.values()].filter(group =>
            group.some(model => model.requiresReview && !model.validated)
        ).length;
        const progress = issueRows.length
            ? Math.round(((issueRows.length - unresolved.length) / issueRows.length) * 100)
            : 100;

        reviewPanel.querySelector('.ltoa-review-summary').textContent = unresolved.length
            ? `${unresolved.length} contrôle${unresolved.length > 1 ? 's' : ''} restant${unresolved.length > 1 ? 's' : ''} avant l’export.`
            : `Tout est prêt. Effectuez un dernier coup d’œil puis validez la déclaration.`;
        reviewPanel.querySelector('[data-stat="contracts"]').textContent = String(contractGroups.size);
        reviewPanel.querySelector('[data-stat="lines"]').textContent = String(models.length);
        reviewPanel.querySelector('[data-stat="issues"]').textContent = String(unresolved.length);
        reviewPanel.querySelector('.ltoa-review-progress-bar').style.width = `${progress}%`;
        reviewPanel.querySelector('.ltoa-review-progress-label').textContent = issueRows.length
            ? `${issueRows.length - unresolved.length}/${issueRows.length} alertes vérifiées`
            : 'Aucune anomalie détectée';
        reviewPanel.querySelector('[data-review-filter="issues"] span').textContent = String(contractsToReview);
        reviewPanel.querySelector('[data-review-filter="all"] span').textContent = String(contractGroups.size);
        reviewPanel.querySelector('[data-review-overview]').checked = Boolean(state.review?.overviewConfirmed);
        const validated = Boolean(state.review?.completedAt)
            && state.review?.fingerprint === declarationFingerprint(state);
        reviewPanel.classList.toggle('is-validated', validated);
        reviewPanel.querySelectorAll('[data-review-step]').forEach(step => {
            const stepNumber = Number(step.dataset.reviewStep);
            step.classList.toggle('done', stepNumber < (validated ? 4 : 3));
            step.classList.toggle('current', stepNumber === (validated ? 4 : 3));
        });
        const excelButton = reviewPanel.querySelector('[data-review-action="excel"]');
        excelButton.disabled = !validated;
        excelButton.title = validated ? 'Télécharger le fichier Excel validé' : 'Validez d’abord la déclaration';
        const confirmButton = reviewPanel.querySelector('[data-review-action="confirm"]');
        confirmButton.hidden = validated;
        reviewPanel.querySelector('.ltoa-review-validation-state').textContent = validated
            ? 'Déclaration validée — Excel prêt'
            : unresolved.length
                ? `${unresolved.length} contrôle${unresolved.length > 1 ? 's' : ''} à terminer`
                : 'Contrôle final à confirmer';
        reviewPanel.dataset.dirty = 'false';

        const cleanBanner = issueRows.length === 0
            ? `<section class="ltoa-clean-banner">
                <span>✓</span>
                <div><strong>Aucune anomalie détectée</strong><p>Le contrôle automatique est terminé. Parcourez le récapitulatif, confirmez le contrôle global, puis validez la déclaration.</p></div>
            </section>`
            : '';
        const emptyFilter = `<section class="ltoa-review-empty-filter" hidden>
            <span>✓</span><div><strong>Aucun contrôle restant</strong><p>Passez sur « Tous » pour effectuer le dernier contrôle visuel du lot.</p></div>
        </section>`;

        content.innerHTML = cleanBanner + emptyFilter + [...contractGroups.entries()].map(([contractNumber, group]) => {
            const subscriberModel = group.find(model => model.role === 'Souscripteur') || group[0];
            const hasIssues = group.some(model => model.requiresReview);
            const unresolvedInContract = group.filter(model => model.requiresReview && !model.validated).length;
            const contractStatus = unresolvedInContract
                ? `<span class="ltoa-status ltoa-status-warning"><i></i>${unresolvedInContract} à vérifier</span>`
                : `<span class="ltoa-status ltoa-status-ready"><i></i>${hasIssues ? 'Vérifié' : 'Prêt'}</span>`;
            const contractFields = `
                ${segmentedChoice('data-contract-field', 'adhesionType', subscriberModel.adhesionType, [
                    { value: 'Individuel', label: 'Individuel' },
                    { value: 'Famille', label: 'Famille' },
                ], 'Adhésion', true)}
                ${segmentedChoice('data-contract-field', 'paymentStatus', subscriberModel.paymentStatus, [
                    { value: 'confirmed', label: 'Confirmé' },
                    { value: 'needs-review', label: 'À vérifier' },
                    { value: 'non-justified', label: 'Non justifié' },
                ], 'Paiement')}
                <label class="ltoa-field ltoa-field-country"><span>Pays de rapatriement</span><input data-contract-field="country" value="${escapeHtml(subscriberModel.country)}" placeholder="À renseigner" aria-label="Pays d'inhumation"></label>`;
            const people = group.map((model, personIndex) => {
                const evidence = model.evidenceDocuments?.length
                    ? `<details class="ltoa-evidence"><summary>Voir les pièces utilisées</summary><div>${escapeHtml(model.evidenceDocuments.map(item => item.name || item).join(' · '))}</div></details>`
                    : '';
                const issues = model.issues.length
                    ? `<div class="ltoa-issues">${model.issues.map(issue => `<div><span>!</span><p>${escapeHtml(issue)}</p></div>`).join('')}</div>`
                    : '';
                const validation = model.requiresReview
                    ? `<label class="ltoa-check"><input type="checkbox" data-review-validate ${model.validated ? 'checked' : ''}><span class="ltoa-check-box">✓</span><span>J’ai vérifié cette ligne</span></label>`
                    : '<span class="ltoa-auto-ok"><b>✓</b> Contrôle automatique réussi</span>';
                const personLinks = `<div class="ltoa-person-links">${reviewLink(model.clientUrl, 'Fiche client')}${reviewLink(model.gedUrl, 'GED')}</div>`;
                return `<article class="ltoa-person ${model.requiresReview ? 'has-issue' : ''}" data-row-id="${escapeHtml(model.rowId)}" data-policy-id="${escapeHtml(model.policyId)}">
                    <div class="ltoa-person-heading">
                        <div class="ltoa-person-number">${personIndex + 1}</div>
                        <div><strong>${escapeHtml(model.firstName || 'Prénom manquant')} ${escapeHtml(model.lastName || 'Nom manquant')}</strong><span>${escapeHtml(model.role)}</span></div>
                        ${personLinks}
                        ${validation}
                    </div>
                    ${issues}
                    <div class="ltoa-fields-grid">
                        ${segmentedChoice('data-row-field', 'role', model.role, [
                            { value: 'Souscripteur', label: 'Souscripteur' },
                            { value: 'Bénéficiaire', label: 'Bénéficiaire' },
                        ], 'Rôle')}
                        ${segmentedChoice('data-row-field', 'civility', model.civility, [
                            { value: '', label: '—' },
                            { value: 'M.', label: 'M.' },
                            { value: 'Mme', label: 'Mme' },
                            { value: 'Mlle', label: 'Mlle' },
                        ], 'Civilité')}
                        <label class="ltoa-field"><span>Nom</span><input data-row-field="lastName" value="${escapeHtml(model.lastName)}" placeholder="Nom"></label>
                        <label class="ltoa-field"><span>Prénom</span><input data-row-field="firstName" value="${escapeHtml(model.firstName)}" placeholder="Prénom"></label>
                        <label class="ltoa-field"><span>Date de naissance</span><input data-row-field="birthDate" value="${escapeHtml(model.birthDate)}" placeholder="jj/mm/aaaa" inputmode="numeric"></label>
                    </div>
                    ${evidence}
                </article>`;
            }).join('');
            return `<section class="ltoa-contract-card" data-has-issues="${hasIssues}" data-needs-review="${unresolvedInContract > 0}" data-contract-number="${contractNumber}">
                <div class="ltoa-contract-header">
                    <div class="ltoa-contract-title">
                        <span class="ltoa-contract-index">${contractNumber}</span>
                        <div><small>CONTRAT</small><h3>${escapeHtml(subscriberModel.firstName)} ${escapeHtml(subscriberModel.lastName)}</h3></div>
                        ${contractStatus}
                    </div>
                </div>
                <div class="ltoa-contract-meta">${contractFields}<div class="ltoa-contract-dates"><span>Effet <b>${escapeHtml(subscriberModel.effectDate || '—')}</b></span><span>Expiration <b>${escapeHtml(subscriberModel.expiration || '—')}</b></span></div></div>
                <div class="ltoa-people">${people}</div>
            </section>`;
        }).join('');
        reviewPanel.hidden = false;
        document.documentElement.classList.add('ltoa-review-open');
        if (!unresolved.length && activeReviewFilter === 'issues') activeReviewFilter = 'all';
        applyReviewFilter(activeReviewFilter);
    }

    function saveReviewPanel(confirmDeclaration = false) {
        const state = loadState();
        if (!state?.results?.length) throw new Error('Aucun contrat traité à vérifier.');
        const review = ensureReviewState(state);
        const validated = new Set(review.validatedRowIds);
        const reviewPanel = root.querySelector('.ltoa-review-panel');
        root.querySelectorAll('.ltoa-review-panel .ltoa-contract-card').forEach(card => {
            const subscriberRow = card.querySelector('[data-row-id$=":subscriber"]') || card.querySelector('[data-row-id]');
            const policyId = subscriberRow?.dataset.policyId;
            if (!policyId) return;
            card.querySelectorAll('[data-contract-field]').forEach(input => {
                review.contractOverrides[policyId] = review.contractOverrides[policyId] || {};
                review.contractOverrides[policyId][input.dataset.contractField] = clean(input.value);
            });
        });
        root.querySelectorAll('.ltoa-review-panel [data-row-id]').forEach(row => {
            const rowId = row.dataset.rowId;
            review.overrides[rowId] = review.overrides[rowId] || {};
            row.querySelectorAll('[data-row-field]').forEach(input => {
                review.overrides[rowId][input.dataset.rowField] = clean(input.value);
            });
            const checkbox = row.querySelector('[data-review-validate]');
            if (checkbox?.checked) validated.add(rowId);
            else if (checkbox) validated.delete(rowId);
        });
        review.validatedRowIds = [...validated];
        review.overviewConfirmed = Boolean(reviewPanel.querySelector('[data-review-overview]')?.checked);
        review.completedAt = null;
        review.fingerprint = null;
        saveState(state);
        const remaining = unresolvedReviewRows(state);
        const undecidedPayments = unresolvedPaymentContracts(state);
        if (confirmDeclaration && undecidedPayments.length) {
            renderReviewPanel();
            throw new Error(`${undecidedPayments.length} paiement(s) sont encore « À vérifier ». Passez-les en « Confirmé » ou « Non justifié » avant validation.`);
        }
        if (confirmDeclaration && remaining.length) {
            renderReviewPanel();
            throw new Error(`${remaining.length} alerte(s) doivent encore être cochées comme vérifiées.`);
        }
        if (confirmDeclaration && !review.overviewConfirmed) {
            renderReviewPanel();
            throw new Error('Confirmez le contrôle global du récapitulatif avant de valider.');
        }
        if (confirmDeclaration) {
            review.completedAt = new Date().toISOString();
            state.reviewCompletedAt = review.completedAt;
            review.fingerprint = declarationFingerprint(state);
            saveState(state);
            renderReviewPanel();
            setMessage('Déclaration vérifiée et validée. Téléchargez maintenant le fichier Excel.');
            return;
        }
        saveState(state);
        renderReviewPanel();
        setMessage('Corrections enregistrées. La validation finale reste obligatoire.');
    }

    async function downloadAgisWorkbook() {
        const state = loadState();
        if (!state?.results?.length) throw new Error('Aucun contrat traité à exporter.');
        const unresolved = unresolvedReviewRows(state);
        const undecidedPayments = unresolvedPaymentContracts(state);
        if (unresolved.length || undecidedPayments.length || !state.review?.completedAt) {
            renderReviewPanel();
            throw new Error(undecidedPayments.length
                ? `${undecidedPayments.length} paiement(s) restent à trancher avant l’export.`
                : unresolved.length
                ? `${unresolved.length} alerte(s) doivent être vérifiées avant l’export.`
                : 'Validez la déclaration dans l’écran de vérification avant l’export.');
        }
        if (state.review.fingerprint !== declarationFingerprint(state)) {
            state.review.completedAt = null;
            state.review.fingerprint = null;
            saveState(state);
            renderReviewPanel();
            throw new Error('Les données ont changé depuis la validation. Vérifiez puis validez à nouveau la déclaration.');
        }
        if (!globalThis.ExcelJS) throw new Error('Le générateur Excel ne s’est pas chargé. Rechargez la page puis réessayez.');

        setMessage('Création de la déclaration Excel AGIS…');
        const workbook = new globalThis.ExcelJS.Workbook();
        workbook.creator = 'LTOA Assurances';
        workbook.created = new Date();
        const worksheet = workbook.addWorksheet('Export_Declaration_mensuelle_de');
        const rows = buildAgisRows(state);
        worksheet.addTable({
            name: 'DeclarationMensuelleAGIS',
            ref: 'A1',
            headerRow: true,
            totalsRow: false,
            style: { theme: 'TableStyleMedium2', showRowStripes: true },
            columns: AGIS_HEADERS.map(name => ({ name })),
            rows,
        });
        worksheet.views = [{ state: 'frozen', ySplit: 1 }];
        const widths = [16, 14, 10, 21, 17, 10, 10, 18, 13, 23, 12, 14, 11, 18, 18, 17, 19, 13, 41, 22, 13, 24];
        widths.forEach((width, index) => { worksheet.getColumn(index + 1).width = width; });
        [9, 10, 16].forEach(columnNumber => {
            worksheet.getColumn(columnNumber).numFmt = 'dd/mm/yyyy';
        });
        worksheet.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
        worksheet.getRow(1).alignment = { vertical: 'middle' };
        worksheet.getRow(1).height = 20;

        const buffer = await workbook.xlsx.writeBuffer();
        const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        const url = URL.createObjectURL(blob);
        const anchor = document.createElement('a');
        anchor.href = url;
        anchor.download = `DECLARATION-AGIS-${state.month || 'periode'}.xlsx`;
        document.body.appendChild(anchor);
        anchor.click();
        anchor.remove();
        state.excelExportedAt = new Date().toISOString();
        saveState(state);
        setTimeout(() => URL.revokeObjectURL(url), 3000);
        const excludedContracts = buildAgisRowModels(state).filter(model =>
            model.role === 'Souscripteur' && model.paymentStatus === 'non-justified'
        ).length;
        setMessage(`Déclaration Excel créée : ${rows.length} ligne(s). ${excludedContracts} contrat(s) sans paiement justifié ont été exclus.`);
    }

    const root = document.createElement('div');
    root.id = APP_ID;
    root.innerHTML = `
        <section class="ltoa-agent-panel" hidden>
            <header>
                <strong>Déclaration mensuelle AGIS</strong>
                <button type="button" class="ltoa-agent-close" aria-label="Réduire le panneau" title="Réduire">−</button>
            </header>
            <label class="ltoa-agent-label">Mois à déclarer
                <input type="month" class="ltoa-agent-month">
            </label>
            <section class="ltoa-month-status" aria-live="polite"></section>
            <p class="ltoa-agent-message" hidden>Prêt.</p>
            <div class="ltoa-agent-actions">
                <button type="button" class="primary" data-action="launch">Lancer la déclaration</button>
                <button type="button" class="primary" data-action="resume" hidden>Reprendre le traitement</button>
                <button type="button" class="review" data-action="review" hidden>Consulter / contrôler le lot</button>
                <button type="button" class="excel" data-action="excel" hidden>Télécharger l’Excel validé</button>
                <button type="button" class="secondary" data-action="export" hidden>Télécharger le JSON</button>
                <button type="button" class="danger" data-action="reset" hidden>Réinitialiser ce mois</button>
            </div>
            <p class="ltoa-agent-note">Choisissez un mois. Le bot mémorise séparément l’avancement, le contrôle et les téléchargements de chaque déclaration.</p>
        </section>
        <section class="ltoa-review-panel" hidden>
            <header>
                <div class="ltoa-review-title">
                    <span class="ltoa-review-app-icon">A</span>
                    <div><strong>Contrôle de la déclaration AGIS</strong><p class="ltoa-review-summary"></p></div>
                </div>
                <button type="button" class="ltoa-review-close" aria-label="Fermer">×</button>
            </header>
            <div class="ltoa-review-steps" aria-label="Étapes de la déclaration">
                <div data-review-step="1"><b>1</b><span><strong>Analyse</strong><small>Contrats traités</small></span></div>
                <i></i>
                <div data-review-step="2"><b>2</b><span><strong>Anomalies</strong><small>Corrections ciblées</small></span></div>
                <i></i>
                <div data-review-step="3"><b>3</b><span><strong>Contrôle final</strong><small>Validation humaine</small></span></div>
                <i></i>
                <div data-review-step="4"><b>4</b><span><strong>Export Excel</strong><small>Fichier AGIS final</small></span></div>
            </div>
            <div class="ltoa-review-toolbar">
                <div class="ltoa-review-stats">
                    <div><b data-stat="contracts">0</b><span>Contrats</span></div>
                    <div><b data-stat="lines">0</b><span>Lignes AGIS</span></div>
                    <div class="warning"><b data-stat="issues">0</b><span>À vérifier</span></div>
                </div>
                <div class="ltoa-review-progress">
                    <div><span class="ltoa-review-progress-label"></span><small>Export verrouillé jusqu’à validation</small></div>
                    <div class="ltoa-review-progress-track"><i class="ltoa-review-progress-bar"></i></div>
                </div>
                <div class="ltoa-review-filters" role="group" aria-label="Filtrer les contrats">
                    <button type="button" data-review-filter="issues">À vérifier <span>0</span></button>
                    <button type="button" data-review-filter="all">Tous <span>0</span></button>
                </div>
            </div>
            <main class="ltoa-review-content"></main>
            <footer>
                <div class="ltoa-review-final-check">
                    <label>
                        <input type="checkbox" data-review-overview>
                        <span class="ltoa-check-box">✓</span>
                        <p><strong>J’ai contrôlé le récapitulatif du lot</strong><small>Obligatoire, même si aucune anomalie n’a été détectée.</small></p>
                    </label>
                    <em class="ltoa-review-validation-state">Contrôle final à confirmer</em>
                </div>
                <nav>
                    <button type="button" class="secondary" data-review-action="json">Télécharger le JSON</button>
                    <button type="button" class="secondary" data-review-action="save">Enregistrer</button>
                    <button type="button" class="confirm" data-review-action="confirm"><span>✓</span> Valider les contrôles</button>
                    <button type="button" class="excel" data-review-action="excel" disabled>Télécharger l’Excel</button>
                </nav>
            </footer>
        </section>
    `;
    const style = document.createElement('style');
    style.textContent = `
        #${APP_ID} { position: fixed; right: 18px; bottom: 18px; z-index: 2147483647; font-family: -apple-system, BlinkMacSystemFont, "SF Pro Display", "Segoe UI", sans-serif; pointer-events: none; }
        #${APP_ID} > section { pointer-events: auto; }
        #${APP_ID} button { cursor: pointer; border: 0; border-radius: 10px; font: inherit; font-weight: 650; }
        #${APP_ID} .ltoa-agent-panel { width: 390px; margin-top: 10px; padding: 0 16px 14px; color: #1f2933; background: #fff; border: 1px solid #d8dee4; border-radius: 14px; box-shadow: 0 14px 38px rgba(15,23,42,.22); }
        #${APP_ID} .ltoa-agent-panel[hidden] { display: none !important; visibility: hidden !important; pointer-events: none !important; }
        #${APP_ID}-company-entry { cursor: pointer; }
        #${APP_ID} header { display: flex; align-items: center; justify-content: space-between; padding: 14px 0 10px; border-bottom: 1px solid #e7ebef; }
        #${APP_ID} .ltoa-agent-close { width: 30px; height: 30px; font-size: 22px; color: #52616b; background: #eef1f4; }
        #${APP_ID} .ltoa-agent-label { display: block; margin: 13px 0 8px; font-weight: 700; }
        #${APP_ID} .ltoa-agent-month { display: block; box-sizing: border-box; width: 100%; margin-top: 6px; padding: 9px; border: 1px solid #b8c2cc; border-radius: 7px; }
        #${APP_ID} .ltoa-agent-message { min-height: 42px; padding: 9px; line-height: 1.4; background: #edf7f5; border-radius: 7px; }
        #${APP_ID} .ltoa-agent-message[hidden] { display: none !important; }
        #${APP_ID} .ltoa-month-status { margin: 11px 0; padding: 13px; background: #f5f7fa; border: 1px solid #e3e7ec; border-radius: 11px; }
        #${APP_ID} .ltoa-month-status header { padding: 0 0 9px; border: 0; }
        #${APP_ID} .ltoa-month-status h3 { margin: 0; font-size: 14px; }
        #${APP_ID} .ltoa-month-status .ltoa-state-badge { padding: 4px 8px; color: #344054; background: #e5e9ef; border-radius: 99px; font-size: 10px; font-weight: 750; }
        #${APP_ID} .ltoa-month-status .ltoa-state-badge.progress { color: #8a5107; background: #fff0d5; }
        #${APP_ID} .ltoa-month-status .ltoa-state-badge.review { color: #1769c2; background: #e7f2ff; }
        #${APP_ID} .ltoa-month-status .ltoa-state-badge.done { color: #087742; background: #e4f7ec; }
        #${APP_ID} .ltoa-month-status p { margin: 0 0 8px; color: #586474; font-size: 12px; line-height: 1.4; }
        #${APP_ID} .ltoa-month-status ul { display: grid; gap: 5px; margin: 0; padding: 0; list-style: none; }
        #${APP_ID} .ltoa-month-status li { display: flex; align-items: center; gap: 7px; color: #657083; font-size: 11px; }
        #${APP_ID} .ltoa-month-status li::before { content: "•"; color: #9aa3af; font-weight: 800; }
        #${APP_ID} .ltoa-month-status li.ok::before { content: "✓"; color: #159455; }
        #${APP_ID} .ltoa-month-status li.missing::before { content: "!"; display: grid; width: 14px; height: 14px; place-items: center; color: #fff; background: #df8d15; border-radius: 50%; font-size: 9px; }
        #${APP_ID} .ltoa-agent-actions { display: grid; gap: 8px; }
        #${APP_ID} .ltoa-agent-actions button { padding: 10px; color: #fff; background: #18324b; }
        #${APP_ID} .ltoa-agent-actions button[hidden] { display: none !important; }
        #${APP_ID} .ltoa-agent-actions button.primary { background: #0878f9; }
        #${APP_ID} .ltoa-agent-actions button.secondary { color: #344054; background: #e9edf2; }
        #${APP_ID} .ltoa-agent-actions button.review { background: #9a6700; }
        #${APP_ID} .ltoa-agent-actions button.excel { background: #107c41; }
        #${APP_ID} .ltoa-agent-actions button.danger { background: #b42318; }
        #${APP_ID} .ltoa-agent-note { margin-bottom: 0; font-size: 12px; line-height: 1.4; color: #667784; }
        html.ltoa-review-open, html.ltoa-review-open body { overflow: hidden !important; }
        #${APP_ID} .ltoa-review-panel { --ink: #18212f; --muted: #667085; --line: rgba(15,23,42,.09); --blue: #0878f9; --blue-dark: #0068dc; --green: #159455; --amber: #bf6a02; position: fixed; inset: 0; width: 100vw; height: 100vh; display: flex; flex-direction: column; color: var(--ink); background: #f5f7fa; border: 0; border-radius: 0; box-shadow: none; overflow: hidden; }
        #${APP_ID} .ltoa-review-panel[hidden] { display: none; }
        #${APP_ID} .ltoa-review-panel > header { flex: 0 0 auto; min-height: 58px; padding: 12px 22px; background: rgba(255,255,255,.96); border-bottom: 1px solid var(--line); }
        #${APP_ID} .ltoa-review-title { display: flex; align-items: center; gap: 12px; }
        #${APP_ID} .ltoa-review-title strong { display: block; font-size: 17px; letter-spacing: -.02em; }
        #${APP_ID} .ltoa-review-title p { margin: 3px 0 0; color: var(--muted); font-size: 12.5px; font-weight: 450; }
        #${APP_ID} .ltoa-review-app-icon { display: grid; width: 36px; height: 36px; place-items: center; color: #fff; background: linear-gradient(145deg,#1d8cff,#0067df); border-radius: 10px; box-shadow: inset 0 1px rgba(255,255,255,.35), 0 5px 14px rgba(8,120,249,.25); font-size: 18px; font-weight: 750; }
        #${APP_ID} .ltoa-review-close { width: 32px; height: 32px; color: #556070; background: #e9edf2; border-radius: 50%; font-size: 20px; line-height: 1; }
        #${APP_ID} .ltoa-review-close:hover { background: #dfe4ea; }
        #${APP_ID} .ltoa-review-steps { display: flex; flex: 0 0 auto; align-items: center; justify-content: center; gap: 14px; padding: 10px 22px; background: #fff; border-bottom: 1px solid var(--line); }
        #${APP_ID} .ltoa-review-steps > div { display: flex; align-items: center; gap: 8px; min-width: 130px; color: #8b95a3; }
        #${APP_ID} .ltoa-review-steps > div > b { display: grid; flex: 0 0 auto; width: 26px; height: 26px; place-items: center; background: #e9edf2; border-radius: 50%; font-size: 11px; }
        #${APP_ID} .ltoa-review-steps span { display: block; }
        #${APP_ID} .ltoa-review-steps strong, #${APP_ID} .ltoa-review-steps small { display: block; white-space: nowrap; }
        #${APP_ID} .ltoa-review-steps strong { font-size: 11px; }
        #${APP_ID} .ltoa-review-steps small { margin-top: 1px; font-size: 9px; font-weight: 500; }
        #${APP_ID} .ltoa-review-steps > i { width: 44px; height: 2px; background: #e0e5eb; border-radius: 99px; }
        #${APP_ID} .ltoa-review-steps > div.done { color: #168153; }
        #${APP_ID} .ltoa-review-steps > div.done > b { color: #fff; background: #20a36a; }
        #${APP_ID} .ltoa-review-steps > div.done > b::after { content: "✓"; }
        #${APP_ID} .ltoa-review-steps > div.done > b { font-size: 0; }
        #${APP_ID} .ltoa-review-steps > div.done > b::after { font-size: 11px; }
        #${APP_ID} .ltoa-review-steps > div.current { color: #096fd6; }
        #${APP_ID} .ltoa-review-steps > div.current > b { color: #fff; background: #0878f9; box-shadow: 0 0 0 4px rgba(8,120,249,.12); }
        #${APP_ID} .ltoa-review-toolbar { display: grid; grid-template-columns: auto minmax(240px,1fr) auto; align-items: center; gap: 26px; padding: 14px 22px; background: rgba(255,255,255,.72); border-bottom: 1px solid var(--line); }
        #${APP_ID} .ltoa-review-stats { display: flex; gap: 8px; }
        #${APP_ID} .ltoa-review-stats > div { min-width: 78px; padding: 8px 12px; background: #f2f4f7; border-radius: 12px; }
        #${APP_ID} .ltoa-review-stats b { display: block; font-size: 17px; line-height: 1; }
        #${APP_ID} .ltoa-review-stats span { display: block; margin-top: 5px; color: var(--muted); font-size: 10px; font-weight: 650; text-transform: uppercase; letter-spacing: .04em; }
        #${APP_ID} .ltoa-review-stats .warning { background: #fff2dc; color: #9a5500; }
        #${APP_ID} .ltoa-review-progress > div:first-child { display: flex; justify-content: space-between; gap: 10px; margin-bottom: 7px; }
        #${APP_ID} .ltoa-review-progress-label { font-size: 12px; font-weight: 650; }
        #${APP_ID} .ltoa-review-progress small { color: var(--muted); font-size: 11px; }
        #${APP_ID} .ltoa-review-progress-track { height: 6px; overflow: hidden; background: #e5e9ef; border-radius: 99px; }
        #${APP_ID} .ltoa-review-progress-bar { display: block; width: 0; height: 100%; background: linear-gradient(90deg,#1688ff,#24b47e); border-radius: inherit; transition: width .25s ease; }
        #${APP_ID} .ltoa-review-filters { display: flex; padding: 3px; background: #e9edf2; border-radius: 11px; }
        #${APP_ID} .ltoa-review-filters button { padding: 7px 11px; color: #5b6573; background: transparent; border-radius: 8px; font-size: 12px; }
        #${APP_ID} .ltoa-review-filters button.active { color: #18212f; background: #fff; box-shadow: 0 1px 4px rgba(15,23,42,.12); }
        #${APP_ID} .ltoa-review-filters span { display: inline-grid; min-width: 17px; height: 17px; margin-left: 4px; padding: 0 3px; place-items: center; color: #606a78; background: rgba(102,112,133,.12); border-radius: 99px; font-size: 10px; }
        #${APP_ID} .ltoa-review-content { flex: 1 1 auto; min-height: 0; overflow: auto; padding: 18px 22px 26px; scroll-behavior: smooth; }
        #${APP_ID} .ltoa-contract-card { margin: 0 auto 14px; overflow: visible; background: #fff; border: 1px solid rgba(15,23,42,.08); border-radius: 17px; box-shadow: 0 5px 18px rgba(15,23,42,.055); }
        #${APP_ID} .ltoa-contract-card[hidden] { display: none; }
        #${APP_ID} .ltoa-clean-banner, #${APP_ID} .ltoa-review-empty-filter { display: flex; align-items: center; gap: 12px; margin: 0 0 14px; padding: 14px 16px; color: #176f49; background: #ecf9f2; border: 1px solid #c9eedb; border-radius: 14px; }
        #${APP_ID} .ltoa-clean-banner > span, #${APP_ID} .ltoa-review-empty-filter > span { display: grid; flex: 0 0 auto; width: 30px; height: 30px; place-items: center; color: #fff; background: #20a36a; border-radius: 50%; font-weight: 800; }
        #${APP_ID} .ltoa-clean-banner strong, #${APP_ID} .ltoa-review-empty-filter strong { display: block; font-size: 13px; }
        #${APP_ID} .ltoa-clean-banner p, #${APP_ID} .ltoa-review-empty-filter p { margin: 3px 0 0; color: #4f7463; font-size: 11px; }
        #${APP_ID} .ltoa-review-empty-filter[hidden] { display: none; }
        #${APP_ID} .ltoa-contract-header { display: flex; align-items: center; justify-content: space-between; gap: 20px; padding: 14px 16px 12px; border-bottom: 1px solid var(--line); }
        #${APP_ID} .ltoa-contract-title { display: flex; align-items: center; gap: 11px; min-width: 0; }
        #${APP_ID} .ltoa-contract-index { display: grid; flex: 0 0 auto; width: 34px; height: 34px; place-items: center; color: #fff; background: #1b2736; border-radius: 10px; font-size: 14px; font-weight: 750; }
        #${APP_ID} .ltoa-contract-title small { display: block; margin-bottom: 2px; color: #8992a0; font-size: 9px; font-weight: 750; letter-spacing: .09em; }
        #${APP_ID} .ltoa-contract-title h3 { max-width: 340px; margin: 0; overflow: hidden; font-size: 15px; text-overflow: ellipsis; white-space: nowrap; }
        #${APP_ID} .ltoa-status { display: inline-flex; align-items: center; gap: 6px; margin-left: 7px; padding: 5px 9px; border-radius: 99px; font-size: 10.5px; font-weight: 700; }
        #${APP_ID} .ltoa-status i { width: 6px; height: 6px; border-radius: 50%; }
        #${APP_ID} .ltoa-status-warning { color: #935000; background: #fff0d6; }
        #${APP_ID} .ltoa-status-warning i { background: #f3a11b; }
        #${APP_ID} .ltoa-status-ready { color: #087742; background: #e5f7ed; }
        #${APP_ID} .ltoa-status-ready i { background: #18a663; }
        #${APP_ID} .ltoa-source-links { display: flex; flex: 0 0 auto; gap: 7px; }
        #${APP_ID} .ltoa-person-links { display: flex; flex: 0 0 auto; gap: 6px; margin-left: auto; }
        #${APP_ID} .ltoa-person-links + .ltoa-check, #${APP_ID} .ltoa-person-links + .ltoa-auto-ok { margin-left: 4px; }
        #${APP_ID} .ltoa-source-link { display: inline-flex; align-items: center; gap: 6px; padding: 7px 9px; color: #1769c2; background: #eef6ff; border-radius: 9px; font-size: 11px; font-weight: 700; text-decoration: none; }
        #${APP_ID} .ltoa-source-link:hover { background: #e2f0ff; }
        #${APP_ID} .ltoa-source-link svg { width: 14px; height: 14px; fill: none; stroke: currentColor; stroke-linecap: round; stroke-linejoin: round; stroke-width: 1.8; }
        #${APP_ID} .ltoa-source-link b { font-size: 10px; }
        #${APP_ID} .ltoa-link-missing { padding: 7px 9px; color: #9b5c18; background: #fff5e8; border-radius: 9px; font-size: 11px; font-weight: 650; }
        #${APP_ID} .ltoa-contract-meta { display: grid; grid-template-columns: 210px minmax(250px,1.2fr) minmax(190px,1fr) auto; align-items: end; gap: 10px; padding: 11px 16px; background: #fbfcfd; border-bottom: 1px solid var(--line); }
        #${APP_ID} .ltoa-contract-dates { display: flex; align-items: center; gap: 14px; padding: 8px 0; color: var(--muted); font-size: 11px; }
        #${APP_ID} .ltoa-contract-dates span { white-space: nowrap; }
        #${APP_ID} .ltoa-contract-dates b { margin-left: 3px; color: #354052; }
        #${APP_ID} .ltoa-people { padding: 3px 16px 10px; }
        #${APP_ID} .ltoa-person { padding: 11px 0 9px; border-bottom: 1px solid var(--line); }
        #${APP_ID} .ltoa-person:last-child { border-bottom: 0; }
        #${APP_ID} .ltoa-person-heading { display: flex; align-items: center; gap: 9px; min-height: 32px; }
        #${APP_ID} .ltoa-person-number { display: grid; width: 24px; height: 24px; place-items: center; color: #657083; background: #edf0f4; border-radius: 7px; font-size: 10px; font-weight: 750; }
        #${APP_ID} .ltoa-person-heading > div:nth-child(2) { display: flex; align-items: baseline; gap: 7px; min-width: 0; }
        #${APP_ID} .ltoa-person-heading strong { font-size: 13px; }
        #${APP_ID} .ltoa-person-heading > div:nth-child(2) span { color: var(--muted); font-size: 10.5px; }
        #${APP_ID} .ltoa-fields-grid { display: grid; grid-template-columns: minmax(220px,1.15fr) minmax(175px,.9fr) minmax(150px,1.2fr) minmax(150px,1.2fr) 145px; gap: 9px; margin-top: 9px; }
        #${APP_ID} .ltoa-field { display: block; min-width: 0; }
        #${APP_ID} .ltoa-field > span { display: block; margin: 0 0 4px 2px; color: #717b8b; font-size: 9.5px; font-weight: 700; letter-spacing: .025em; text-transform: uppercase; }
        #${APP_ID} .ltoa-review-panel input:not([type="hidden"]):not([type="checkbox"]) { box-sizing: border-box; width: 100%; height: 34px; padding: 0 9px; color: #202a38; background: #fff; border: 1px solid #d7dce3; border-radius: 9px; outline: none; font: inherit; font-size: 12px; transition: border-color .15s,box-shadow .15s; }
        #${APP_ID} .ltoa-review-panel input:not([type="hidden"]):not([type="checkbox"]):focus { border-color: #4c9dff; box-shadow: 0 0 0 3px rgba(8,120,249,.12); }
        #${APP_ID} .ltoa-segmented { display: flex; width: 100%; height: 34px; padding: 3px; box-sizing: border-box; overflow: visible; background: #edf0f4; border: 1px solid #dfe3e8; border-radius: 9px; }
        #${APP_ID} .ltoa-segmented button { flex: 1 1 auto; min-width: 0; padding: 0 6px; color: #657083; background: transparent; border-radius: 6px; font-size: 10.5px; white-space: nowrap; }
        #${APP_ID} .ltoa-segmented button.active { color: #172033; background: #fff; box-shadow: 0 1px 4px rgba(15,23,42,.16); }
        #${APP_ID} .ltoa-segmented button:focus-visible { outline: 2px solid #4c9dff; outline-offset: 1px; }
        #${APP_ID} .ltoa-issues { display: grid; gap: 5px; margin: 8px 0 4px 33px; }
        #${APP_ID} .ltoa-issues > div { display: flex; align-items: flex-start; gap: 7px; padding: 7px 9px; color: #7f4700; background: #fff5e4; border: 1px solid #ffe3b6; border-radius: 9px; }
        #${APP_ID} .ltoa-issues span { display: grid; flex: 0 0 auto; width: 16px; height: 16px; place-items: center; color: #fff; background: #e99517; border-radius: 50%; font-size: 10px; font-weight: 800; }
        #${APP_ID} .ltoa-issues p { margin: 0; font-size: 11px; line-height: 1.35; }
        #${APP_ID} .ltoa-check { display: inline-flex; align-items: center; gap: 7px; margin-left: auto; color: #8a5107; font-size: 11px; font-weight: 700; cursor: pointer; }
        #${APP_ID} .ltoa-check input { position: absolute; width: 1px; height: 1px; opacity: 0; pointer-events: none; }
        #${APP_ID} .ltoa-check-box { display: grid; width: 20px; height: 20px; place-items: center; color: transparent; background: #fff; border: 1.5px solid #d4a04c; border-radius: 6px; transition: .15s; }
        #${APP_ID} .ltoa-check input:checked + .ltoa-check-box { color: #fff; background: #19975a; border-color: #19975a; }
        #${APP_ID} .ltoa-auto-ok { display: inline-flex; align-items: center; gap: 6px; margin-left: auto; color: #15804c; font-size: 10.5px; font-weight: 650; }
        #${APP_ID} .ltoa-auto-ok b { display: grid; width: 17px; height: 17px; place-items: center; color: #fff; background: #21a765; border-radius: 50%; font-size: 10px; }
        #${APP_ID} .ltoa-evidence { margin: 7px 0 0 33px; color: var(--muted); font-size: 10.5px; }
        #${APP_ID} .ltoa-evidence summary { cursor: pointer; font-weight: 650; }
        #${APP_ID} .ltoa-evidence div { margin-top: 5px; line-height: 1.4; }
        #${APP_ID} .ltoa-review-panel footer { display: flex; flex: 0 0 auto; align-items: center; justify-content: space-between; gap: 14px; min-height: 62px; padding: 10px 18px; background: rgba(255,255,255,.97); border-top: 1px solid var(--line); box-shadow: 0 -8px 24px rgba(15,23,42,.035); }
        #${APP_ID} .ltoa-review-final-check { display: flex; align-items: center; gap: 14px; color: #566172; }
        #${APP_ID} .ltoa-review-final-check label { display: flex; align-items: center; gap: 9px; cursor: pointer; }
        #${APP_ID} .ltoa-review-final-check input { position: absolute; width: 1px; height: 1px; opacity: 0; pointer-events: none; }
        #${APP_ID} .ltoa-review-final-check input:checked + .ltoa-check-box { color: #fff; background: #19975a; border-color: #19975a; }
        #${APP_ID} .ltoa-review-validation-state { padding: 5px 8px; color: #8a5107; background: #fff2dc; border-radius: 99px; font-size: 9.5px; font-style: normal; font-weight: 700; white-space: nowrap; }
        #${APP_ID} .ltoa-review-panel.is-validated .ltoa-review-validation-state { color: #087742; background: #e5f7ed; }
        #${APP_ID} .ltoa-review-panel footer p { margin: 0; }
        #${APP_ID} .ltoa-review-panel footer strong { display: block; font-size: 10.5px; }
        #${APP_ID} .ltoa-review-panel footer small { display: block; margin-top: 2px; color: #818a98; font-size: 9.5px; }
        #${APP_ID} .ltoa-review-panel footer nav { display: flex; gap: 8px; }
        #${APP_ID} .ltoa-review-panel footer button { min-height: 36px; padding: 0 14px; color: #344054; background: #edf0f4; font-size: 12px; }
        #${APP_ID} .ltoa-review-panel footer button:hover { filter: brightness(.98); }
        #${APP_ID} .ltoa-review-panel footer button.confirm { color: #fff; background: linear-gradient(#1688ff,#0878f9); box-shadow: 0 4px 12px rgba(8,120,249,.22); }
        #${APP_ID} .ltoa-review-panel footer button.confirm span { margin-right: 4px; }
        #${APP_ID} .ltoa-review-panel footer button.excel { color: #fff; background: linear-gradient(#1a9850,#107c41); box-shadow: 0 4px 12px rgba(16,124,65,.2); }
        #${APP_ID} .ltoa-review-panel footer button:disabled { cursor: not-allowed; color: #99a1ad; background: #e8ebef; box-shadow: none; opacity: .75; }
        @media (max-width: 980px) {
            #${APP_ID} .ltoa-review-toolbar { grid-template-columns: 1fr auto; gap: 12px; }
            #${APP_ID} .ltoa-review-progress { grid-column: 1 / -1; grid-row: 2; }
            #${APP_ID} .ltoa-fields-grid { grid-template-columns: repeat(3,minmax(120px,1fr)); }
            #${APP_ID} .ltoa-contract-meta { grid-template-columns: 130px 1fr; }
            #${APP_ID} .ltoa-contract-dates { grid-column: 1 / -1; }
            #${APP_ID} .ltoa-review-steps > div { min-width: 0; }
            #${APP_ID} .ltoa-review-steps small { display: none; }
            #${APP_ID} .ltoa-review-panel footer { align-items: stretch; flex-direction: column; }
            #${APP_ID} .ltoa-review-panel footer nav { justify-content: flex-end; }
        }
    `;
    document.head.appendChild(style);
    document.body.appendChild(root);

    const panel = root.querySelector('.ltoa-agent-panel');
    const monthInput = root.querySelector('.ltoa-agent-month');
    const message = root.querySelector('.ltoa-agent-message');
    monthInput.value = loadState()?.month || previousMonthValue();

    function setLauncherOpen(open) {
        panel.hidden = !open;
        panel.style.setProperty('display', open ? 'block' : 'none', 'important');
        panel.style.setProperty('visibility', open ? 'visible' : 'hidden', 'important');
        panel.style.setProperty('pointer-events', open ? 'auto' : 'none', 'important');
    }

    // Modulr surcharge parfois l'attribut HTML `hidden`. On ferme donc aussi
    // explicitement le lanceur au chargement pour qu'il ne reste jamais affiché.
    setLauncherOpen(false);
    monthInput.addEventListener('change', () => {
        message.hidden = true;
        renderMonthStatus();
    });

    function setMessage(value) {
        message.textContent = value;
    }

    function isCompletedState(state) {
        return state?.status === 'batch-completed' || state?.status === 'test-completed';
    }

    function isValidatedState(state) {
        return Boolean(state?.review?.completedAt)
            && state.review?.fingerprint === declarationFingerprint(state);
    }

    function selectedMonthState() {
        return loadMonthState(monthInput.value);
    }

    function setActionVisibility(action, visible) {
        const button = root.querySelector(`[data-action="${action}"]`);
        if (button) button.hidden = !visible;
    }

    function renderMonthStatus() {
        const month = monthInput.value || previousMonthValue();
        const state = loadMonthState(month);
        const statusBox = root.querySelector('.ltoa-month-status');
        const completed = isCompletedState(state);
        const validated = completed && isValidatedState(state);
        const exported = Boolean(state?.excelExportedAt);
        const inProgress = Boolean(state) && !completed;
        let badge = 'À lancer';
        let badgeClass = '';
        let description = 'Aucune déclaration enregistrée pour ce mois.';
        const details = [];

        if (inProgress) {
            badge = state.status === 'error' ? 'À reprendre' : 'En cours';
            badgeClass = 'progress';
            const total = state.policies?.length || 0;
            const current = Math.min((state.currentIndex || 0) + 1, total || 1);
            description = total
                ? `Le traitement a été interrompu au contrat ${current} sur ${total}.`
                : 'Le traitement a été lancé mais n’est pas encore terminé.';
            details.push(`<li class="ok">Export initié le ${escapeHtml(formatDateTime(state.analysisStartedAt || state.createdAt || state.updatedAt))}</li>`);
            details.push('<li class="missing">Analyse à terminer</li>');
        } else if (completed && !validated) {
            badge = 'Contrôle manquant';
            badgeClass = 'review';
            description = `${state.results?.length || 0} contrat(s) analysé(s). Le fichier Excel reste verrouillé jusqu’au contrôle final.`;
            details.push(`<li class="ok">Analyse terminée le ${escapeHtml(formatDateTime(state.analysisCompletedAt || state.completedAt))}</li>`);
            details.push('<li class="missing">Contrôle humain à effectuer</li>');
        } else if (validated && !exported) {
            badge = 'Prêt à exporter';
            badgeClass = 'review';
            description = `${state.results?.length || 0} contrat(s) contrôlé(s). L’Excel validé peut maintenant être téléchargé.`;
            details.push(`<li class="ok">Analyse terminée le ${escapeHtml(formatDateTime(state.analysisCompletedAt || state.completedAt))}</li>`);
            details.push(`<li class="ok">Contrôle validé le ${escapeHtml(formatDateTime(state.reviewCompletedAt || state.review?.completedAt))}</li>`);
            details.push('<li class="missing">Excel non téléchargé</li>');
        } else if (exported) {
            badge = 'Terminé';
            badgeClass = 'done';
            description = `${state.results?.length || 0} contrat(s) traités. Cette déclaration reste consultable et réexportable.`;
            details.push(`<li class="ok">Analyse terminée le ${escapeHtml(formatDateTime(state.analysisCompletedAt || state.completedAt))}</li>`);
            details.push(`<li class="ok">Contrôle validé le ${escapeHtml(formatDateTime(state.reviewCompletedAt || state.review?.completedAt))}</li>`);
            details.push(`<li class="ok">Excel téléchargé le ${escapeHtml(formatDateTime(state.excelExportedAt))}</li>`);
        }
        if (state?.jsonExportedAt) details.push(`<li class="ok">JSON téléchargé le ${escapeHtml(formatDateTime(state.jsonExportedAt))}</li>`);

        statusBox.innerHTML = `
            <header><h3>${escapeHtml(formatMonthLabel(month))}</h3><span class="ltoa-state-badge ${badgeClass}">${escapeHtml(badge)}</span></header>
            <p>${escapeHtml(description)}</p>
            ${details.length ? `<ul>${details.join('')}</ul>` : ''}
        `;
        setActionVisibility('launch', !state);
        setActionVisibility('resume', inProgress);
        setActionVisibility('review', completed);
        setActionVisibility('excel', validated);
        setActionVisibility('export', Boolean(state));
        setActionVisibility('reset', Boolean(state));
    }

    function updatePanel() {
        renderMonthStatus();
    }

    function resumeSelectedMonth() {
        const month = monthInput.value;
        const state = activateMonthState(month);
        if (!state) throw new Error('Aucun traitement à reprendre pour ce mois.');
        if (isCompletedState(state)) {
            renderReviewPanel();
            return;
        }
        if (state.status === 'filters-applied' || !state.policies?.length) {
            queueDeclarationStart(month);
            return;
        }
        state.status = 'returning-list';
        state.resumedAt = new Date().toISOString();
        saveState(state);
        location.href = state.listUrl || `${location.origin}${LIST_PATH}`;
    }

    root.querySelector('.ltoa-agent-close').addEventListener('click', () => setLauncherOpen(false));
    root.querySelector('.ltoa-review-close').addEventListener('click', () => {
        root.querySelector('.ltoa-review-panel').hidden = true;
        document.documentElement.classList.remove('ltoa-review-open');
    });
    root.querySelectorAll('[data-review-filter]').forEach(button => {
        button.addEventListener('click', () => applyReviewFilter(button.dataset.reviewFilter));
    });
    const reviewPanelElement = root.querySelector('.ltoa-review-panel');
    reviewPanelElement.addEventListener('click', event => {
        const choice = event.target.closest('[data-choice-value]');
        if (!choice) return;
        const group = choice.closest('[data-choice-group]');
        const input = group?.querySelector('input[type="hidden"][data-row-field], input[type="hidden"][data-contract-field]');
        if (!group || !input) return;
        group.querySelectorAll('[data-choice-value]').forEach(button => {
            const active = button === choice;
            button.classList.toggle('active', active);
            button.setAttribute('aria-pressed', String(active));
        });
        input.value = choice.dataset.choiceValue || '';
        input.dispatchEvent(new Event('input', { bubbles: true }));
    });
    reviewPanelElement.addEventListener('input', event => {
        if (!event.target.matches('[data-row-field], [data-contract-field], [data-review-validate]')) return;
        reviewPanelElement.dataset.dirty = 'true';
        reviewPanelElement.classList.remove('is-validated');
        reviewPanelElement.querySelector('[data-review-overview]').checked = false;
        reviewPanelElement.querySelector('[data-review-action="excel"]').disabled = true;
        reviewPanelElement.querySelector('[data-review-action="confirm"]').hidden = false;
        reviewPanelElement.querySelector('.ltoa-review-validation-state').textContent = 'Modifications à valider';
    });
    reviewPanelElement.addEventListener('click', event => {
        if (event.target.closest('.ltoa-source-link')) event.stopPropagation();
    });
    root.querySelectorAll('[data-review-action]').forEach(button => {
        button.addEventListener('click', async () => {
            try {
                const action = button.dataset.reviewAction;
                if (action === 'json') downloadResult();
                if (action === 'save') saveReviewPanel(false);
                if (action === 'confirm') saveReviewPanel(true);
                if (action === 'excel') {
                    if (reviewPanelElement.dataset.dirty === 'true') {
                        throw new Error('Des modifications ne sont pas encore validées.');
                    }
                    await downloadAgisWorkbook();
                }
            } catch (error) {
                setMessage(`Erreur : ${error.message}`);
                reviewPanelElement.querySelector('.ltoa-review-summary').textContent = `Action impossible : ${error.message}`;
            }
        });
    });
    root.querySelectorAll('[data-action]').forEach(button => {
        button.addEventListener('click', async () => {
            try {
                const action = button.dataset.action;
                if (action === 'launch') {
                    if (loadMonthState(monthInput.value)) throw new Error('Ce mois possède déjà un suivi. Reprenez-le ou réinitialisez-le.');
                    queueDeclarationStart(monthInput.value);
                }
                if (action === 'resume') resumeSelectedMonth();
                if (action === 'review') {
                    activateMonthState(monthInput.value);
                    renderReviewPanel();
                }
                if (action === 'excel') {
                    activateMonthState(monthInput.value);
                    await downloadAgisWorkbook();
                }
                if (action === 'export') {
                    activateMonthState(monthInput.value);
                    downloadResult();
                }
                if (action === 'reset') {
                    const month = monthInput.value;
                    if (!window.confirm(`Réinitialiser uniquement la déclaration de ${formatMonthLabel(month)} ?\n\nL’analyse, les corrections et l’historique de téléchargement de ce mois seront supprimés.`)) return;
                    clearState(month);
                    root.querySelector('.ltoa-review-panel').hidden = true;
                    document.documentElement.classList.remove('ltoa-review-open');
                    renderMonthStatus();
                }
            } catch (error) {
                setMessage(`Erreur : ${error.message}`);
                message.hidden = false;
            }
        });
    });

    updatePanel();

    function mountAgisCompanyEntry() {
        if (!isAgisCompanyCard() || document.getElementById(`${APP_ID}-company-entry`)) return;
        const companyHeading = Array.from(document.querySelectorAll('h1')).find(heading => normalize(heading.textContent) === 'agis');
        const headingRow = companyHeading?.closest('.left')?.parentElement;
        const actions = headingRow
            ? Array.from(headingRow.children).find(element => element.classList?.contains('right'))
                || headingRow.querySelector('.right')
            : null;
        if (!actions) return;

        const entry = document.createElement('a');
        entry.id = `${APP_ID}-company-entry`;
        entry.href = '#';
        entry.className = 'square_icon_text low_margin_right tooltip';
        entry.title = 'Choisir un mois, lancer, reprendre ou consulter une déclaration AGIS';
        entry.innerHTML = `
            <span class="fa fa-file-excel fa-inverse low_padding_right"></span>
            Déclaration AGIS
        `;
        const gedEntry = actions.querySelector('a.edm_opening');
        actions.insertBefore(entry, gedEntry || actions.firstChild);
        entry.addEventListener('click', event => {
            event.preventDefault();
            event.stopPropagation();
            updatePanel();
            setLauncherOpen(panel.hidden);
        });
    }

    mountAgisCompanyEntry();
    document.addEventListener('keydown', event => {
        if (event.key !== 'Escape') return;
        const reviewPanel = root.querySelector('.ltoa-review-panel');
        if (!reviewPanel.hidden) {
            reviewPanel.hidden = true;
            document.documentElement.classList.remove('ltoa-review-open');
            return;
        }
        setLauncherOpen(false);
    });
    if (location.pathname === LIST_PATH) {
        setTimeout(() => {
            if (!resumeCompanyStart()) resumeVisibleBatch();
        }, 700);
    }
    if (location.pathname === CLIENT_PATH) setTimeout(captureClientCardAndOpenGed, 700);
    if (location.pathname.startsWith(GED_PREFIX)) {
        setTimeout(() => prepareFullGedAndCapture().catch(error => setMessage(`Erreur de lecture GED : ${error.message}`)), 700);
    }
})();
