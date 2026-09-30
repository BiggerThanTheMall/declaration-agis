// ==UserScript==
// @name         LTOA - Diagnostic Viewer Documents Modulr
// @namespace    https://ltoa-assurances.fr/
// @version      1.0.0
// @description  Diagnostic non intrusif du viewer de documents Modulr : iframe, canvas, blob, images et requêtes réseau. N'envoie aucune donnée à un service externe.
// @author       LTOA Assurances
// @match        https://courtage.modulr.fr/*
// @grant        none
// @run-at       document-start
// @updateURL    https://raw.githubusercontent.com/BiggerThanTheMall/declaration-agis/main/diagnostic-viewer-modulr.user.js
// @downloadURL  https://raw.githubusercontent.com/BiggerThanTheMall/declaration-agis/main/diagnostic-viewer-modulr.user.js
// ==/UserScript==

(function () {
    'use strict';

    const VERSION = '1.0.0';
    const CAPTURE_MS = 20000;
    const MAX_EVENTS = 600;
    const state = {
        active: false,
        startedAt: null,
        stoppedAt: null,
        events: [],
        initialUrl: location.href,
        userAgent: navigator.userAgent,
        version: VERSION,
    };

    function now() {
        return new Date().toISOString();
    }

    function cleanText(value, max = 600) {
        return String(value || '').replace(/\s+/g, ' ').trim().slice(0, max);
    }

    function safeUrl(value) {
        if (!value) return '';
        const raw = String(value);
        if (raw.startsWith('blob:')) return 'blob:<redacted>';
        if (raw.startsWith('data:')) return 'data:<redacted>';
        try {
            const u = new URL(raw, location.href);
            const keys = Array.from(u.searchParams.keys());
            const safeQuery = keys.length ? '?' + keys.map(key => encodeURIComponent(key) + '=<redacted>').join('&') : '';
            return u.origin + u.pathname + safeQuery + (u.hash ? '#<redacted>' : '');
        } catch (_) {
            return cleanText(raw, 300);
        }
    }

    function add(type, data = {}) {
        if (!state.active && !['diagnostic-ready', 'capture-start', 'capture-stop'].includes(type)) return;
        state.events.push({ ts: now(), type, ...data });
        if (state.events.length > MAX_EVENTS) state.events.splice(0, state.events.length - MAX_EVENTS);
        renderStatus();
    }

    function describeElement(el) {
        if (!el || el.nodeType !== 1) return null;
        const tag = el.tagName?.toLowerCase() || '';
        const rect = el.getBoundingClientRect?.();
        const out = {
            tag,
            id: cleanText(el.id, 120),
            className: cleanText(typeof el.className === 'string' ? el.className : '', 180),
            text: cleanText(el.innerText || el.textContent, 220),
            role: cleanText(el.getAttribute?.('role'), 80),
            title: cleanText(el.getAttribute?.('title'), 160),
            ariaLabel: cleanText(el.getAttribute?.('aria-label'), 160),
            width: rect ? Math.round(rect.width) : null,
            height: rect ? Math.round(rect.height) : null,
        };
        for (const attr of ['src', 'href', 'data', 'action']) {
            const value = el.getAttribute?.(attr);
            if (value) out[attr] = safeUrl(value);
        }
        if (tag === 'canvas') {
            out.canvasWidth = el.width || 0;
            out.canvasHeight = el.height || 0;
            try {
                const ctx = el.getContext('2d');
                if (ctx && el.width > 0 && el.height > 0) {
                    ctx.getImageData(0, 0, 1, 1);
                    out.canvasReadable = true;
                } else {
                    out.canvasReadable = null;
                }
            } catch (error) {
                out.canvasReadable = false;
                out.canvasReadError = cleanText(error?.message, 180);
            }
        }
        return out;
    }

    function inspectDocument(doc, label) {
        if (!doc) return;
        const selectors = [
            'iframe', 'embed', 'object', 'canvas', 'img',
            '[role="dialog"]', '.modal', '.modal-dialog', '.modal-content',
            'a[href^="blob:"]', 'a[href*="/edm/"]', 'a[href*="download"]'
        ];
        const found = [];
        for (const selector of selectors) {
            for (const el of Array.from(doc.querySelectorAll(selector)).slice(0, 80)) {
                const d = describeElement(el);
                if (d) found.push(d);
            }
        }
        add('dom-snapshot', {
            label,
            title: cleanText(doc.title, 200),
            url: safeUrl(doc.location?.href || ''),
            elements: found.slice(0, 180),
        });

        for (const iframe of Array.from(doc.querySelectorAll('iframe')).slice(0, 30)) {
            const src = iframe.getAttribute('src') || '';
            const item = { src: safeUrl(src), descriptor: describeElement(iframe) };
            try {
                const child = iframe.contentDocument;
                item.sameOriginAccessible = Boolean(child);
                if (child) {
                    item.childTitle = cleanText(child.title, 160);
                    item.childUrl = safeUrl(child.location?.href || '');
                    item.childCanvasCount = child.querySelectorAll('canvas').length;
                    item.childImageCount = child.querySelectorAll('img').length;
                    item.childEmbedCount = child.querySelectorAll('embed,object').length;
                    item.childTextSample = cleanText(child.body?.innerText, 400);
                }
            } catch (error) {
                item.sameOriginAccessible = false;
                item.accessError = cleanText(error?.message, 180);
            }
            add('iframe-inspect', item);
        }
    }

    function snapshot(label) {
        inspectDocument(document, label);
        const resources = performance.getEntriesByType('resource').slice(-160).map(entry => ({
            name: safeUrl(entry.name),
            initiatorType: entry.initiatorType,
            duration: Math.round(entry.duration),
            transferSize: entry.transferSize || 0,
            encodedBodySize: entry.encodedBodySize || 0,
            decodedBodySize: entry.decodedBodySize || 0,
        }));
        add('performance-snapshot', { label, resources });
    }

    // Capture les fetch lancés par le viewer sans lire le contenu.
    const originalFetch = window.fetch;
    if (typeof originalFetch === 'function') {
        window.fetch = async function (...args) {
            const request = args[0];
            const options = args[1] || {};
            const url = typeof request === 'string' ? request : request?.url;
            const method = options.method || request?.method || 'GET';
            if (state.active) add('fetch-start', { method, url: safeUrl(url) });
            try {
                const response = await originalFetch.apply(this, args);
                if (state.active) {
                    add('fetch-end', {
                        method,
                        url: safeUrl(url),
                        status: response.status,
                        redirected: response.redirected,
                        finalUrl: safeUrl(response.url),
                        contentType: cleanText(response.headers.get('content-type'), 120),
                        contentDisposition: cleanText(response.headers.get('content-disposition'), 180),
                    });
                }
                return response;
            } catch (error) {
                if (state.active) add('fetch-error', { method, url: safeUrl(url), error: cleanText(error?.message, 240) });
                throw error;
            }
        };
    }

    // Capture les XHR lancés par le viewer.
    const xhrOpen = XMLHttpRequest.prototype.open;
    const xhrSend = XMLHttpRequest.prototype.send;
    XMLHttpRequest.prototype.open = function (method, url, ...rest) {
        this.__ltoaDiag = { method: method || 'GET', url: safeUrl(url) };
        return xhrOpen.call(this, method, url, ...rest);
    };
    XMLHttpRequest.prototype.send = function (...args) {
        const meta = this.__ltoaDiag || { method: 'GET', url: '' };
        if (state.active) add('xhr-start', meta);
        this.addEventListener('loadend', () => {
            if (!state.active) return;
            add('xhr-end', {
                ...meta,
                status: this.status,
                finalUrl: safeUrl(this.responseURL),
                responseType: this.responseType || 'text',
                contentType: cleanText(this.getResponseHeader('content-type'), 120),
                contentDisposition: cleanText(this.getResponseHeader('content-disposition'), 180),
            });
        }, { once: true });
        return xhrSend.apply(this, args);
    };

    // Capture window.open si le viewer ouvre une nouvelle fenêtre.
    const originalOpen = window.open;
    window.open = function (url, target, features) {
        if (state.active) add('window-open', {
            url: safeUrl(url),
            target: cleanText(target, 80),
            features: cleanText(features, 220),
        });
        return originalOpen.apply(this, arguments);
    };

    // Capture l'ajout des iframes/canvas/blob après clic.
    const observer = new MutationObserver(mutations => {
        if (!state.active) return;
        for (const mutation of mutations) {
            for (const node of mutation.addedNodes || []) {
                if (!node || node.nodeType !== 1) continue;
                const elements = [node, ...Array.from(node.querySelectorAll?.('iframe,embed,object,canvas,img,a[href^="blob:"],a[href*="/edm/"]') || [])];
                for (const el of elements.slice(0, 80)) {
                    const d = describeElement(el);
                    if (d) add('dom-added', d);
                }
            }
        }
    });

    function startCapture() {
        state.active = true;
        state.startedAt = now();
        state.stoppedAt = null;
        state.events = [];
        add('capture-start', {
            page: safeUrl(location.href),
            instruction: 'Cliquer maintenant sur un document Modulr pour ouvrir sa prévisualisation.',
        });
        observer.observe(document.documentElement, { childList: true, subtree: true, attributes: false });
        snapshot('before-click');
        renderStatus();
        setTimeout(() => {
            if (state.active) stopCapture('automatic-timeout');
        }, CAPTURE_MS);
    }

    function stopCapture(reason = 'manual') {
        if (!state.active) return;
        snapshot('capture-end');
        state.active = false;
        state.stoppedAt = now();
        observer.disconnect();
        state.events.push({ ts: now(), type: 'capture-stop', reason });
        renderStatus();
    }

    function exportJson() {
        if (state.active) stopCapture('export');
        snapshotAfterStop();
        const payload = {
            tool: 'LTOA Modulr Viewer Diagnostic',
            version: VERSION,
            capturedAt: now(),
            startedAt: state.startedAt,
            stoppedAt: state.stoppedAt,
            page: safeUrl(location.href),
            userAgent: state.userAgent,
            privacy: {
                queryValuesRedacted: true,
                blobContentsExcluded: true,
                documentContentsExcluded: true,
                networkBodiesExcluded: true,
            },
            events: state.events,
        };
        const blob = new Blob([JSON.stringify(payload, null, 2)], { type: 'application/json' });
        const a = document.createElement('a');
        a.href = URL.createObjectURL(blob);
        a.download = 'diagnostic-viewer-modulr-' + new Date().toISOString().replace(/[:.]/g, '-') + '.json';
        document.body.appendChild(a);
        a.click();
        setTimeout(() => {
            URL.revokeObjectURL(a.href);
            a.remove();
        }, 1000);
    }

    function snapshotAfterStop() {
        const wasActive = state.active;
        state.active = true;
        snapshot('export-final');
        state.active = wasActive;
    }

    let panel;
    let status;
    function renderStatus() {
        if (!status) return;
        status.textContent = state.active
            ? 'CAPTURE EN COURS - clique sur un document (' + state.events.length + ' événements)'
            : state.startedAt
                ? 'Capture terminée - ' + state.events.length + ' événements'
                : 'Prêt - clique sur SCAN VIEWER puis ouvre un document';
        if (panel) panel.style.borderColor = state.active ? '#c62828' : '#1f5f99';
    }

    function makeButton(label, onClick, primary = false) {
        const b = document.createElement('button');
        b.type = 'button';
        b.textContent = label;
        Object.assign(b.style, {
            border: '1px solid #1f5f99',
            background: primary ? '#1f5f99' : '#fff',
            color: primary ? '#fff' : '#1f5f99',
            borderRadius: '6px',
            padding: '7px 10px',
            cursor: 'pointer',
            fontSize: '12px',
            fontWeight: '700',
        });
        b.addEventListener('click', onClick);
        return b;
    }

    function mount() {
        if (!document.body || document.getElementById('ltoa-viewer-diag-panel')) return;
        panel = document.createElement('div');
        panel.id = 'ltoa-viewer-diag-panel';
        Object.assign(panel.style, {
            position: 'fixed',
            right: '14px',
            bottom: '14px',
            zIndex: '2147483647',
            width: '330px',
            background: '#fff',
            border: '2px solid #1f5f99',
            borderRadius: '10px',
            padding: '10px',
            boxShadow: '0 4px 18px rgba(0,0,0,.2)',
            fontFamily: 'Arial,sans-serif',
            color: '#222',
        });

        const title = document.createElement('div');
        title.textContent = 'Diagnostic Viewer Modulr v' + VERSION;
        title.style.fontWeight = '700';
        title.style.marginBottom = '7px';

        status = document.createElement('div');
        Object.assign(status.style, { fontSize: '12px', lineHeight: '1.35', marginBottom: '8px' });

        const note = document.createElement('div');
        note.textContent = '1. Clique SCAN VIEWER  2. Ouvre normalement un document  3. Attends la prévisualisation  4. Exporte le JSON.';
        Object.assign(note.style, { fontSize: '11px', lineHeight: '1.35', marginBottom: '8px', color: '#555' });

        const actions = document.createElement('div');
        Object.assign(actions.style, { display: 'flex', gap: '6px', flexWrap: 'wrap' });
        actions.append(
            makeButton('SCAN VIEWER', startCapture, true),
            makeButton('STOP', () => stopCapture('manual')),
            makeButton('EXPORT JSON', exportJson)
        );

        panel.append(title, status, note, actions);
        document.body.appendChild(panel);
        renderStatus();
        add('diagnostic-ready', { page: safeUrl(location.href) });
    }

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', mount, { once: true });
    } else {
        mount();
    }
})();
