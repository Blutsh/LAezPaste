(function () {
    'use strict';

    const BUTTON_ID_MARKDOWN = 'laezpaste-copy-btn-md';
    const BUTTON_ID_HTML = 'laezpaste-copy-btn-html';
    const TOOLBAR_CHECK_INTERVAL = 2000;
    const COPIED_FEEDBACK_MS = 2000;

    function isLogAnalyticsPage() {
        return !!(document.querySelector('la-results-bar') || document.querySelector('ag-grid-angular'));
    }

    function getColumnHeaders() {
        const headerCells = document.querySelectorAll('div.ag-header-cell[role="columnheader"]');
        const headers = [];
        const seen = new Set();
        for (const cell of headerCells) {
            const colId = cell.getAttribute('col-id');
            if (!colId || seen.has(colId)) continue;
            seen.add(colId);
            const textEl = cell.querySelector('span.ag-header-cell-text');
            if (textEl) {
                headers.push({ colId, name: textEl.textContent.trim() });
            }
        }
        return headers;
    }

    function getCellValue(row, colId) {
        const cell = row.querySelector(`div.ag-cell[col-id="${colId}"]`);
        if (!cell) return '';

        // title attribute is the most reliable source (full text, no icons)
        const title = cell.getAttribute('title');
        if (title) return title;

        // Fallback: for first column with group expand/contract structure
        const groupValue = cell.querySelector('span.ag-group-value span');
        if (groupValue) return groupValue.textContent.trim();

        // Fallback: standard cell value
        const cellValue = cell.querySelector('span.ag-cell-value');
        if (cellValue) return cellValue.textContent.trim();

        return cell.textContent.trim();
    }

    function getSelectedRows() {
        return document.querySelectorAll('div.ag-center-cols-container div.ag-row-selected[role="row"]');
    }

    function escapeForTeams(value) {
        return value
            .replace(/\|/g, '\\|')
            .replace(/\n/g, ' ')
            .replace(/\r/g, '');
    }

    function buildMarkdownTable(headers, rows) {
        if (headers.length === 0 || rows.length === 0) return '';

        const headerLine = '| ' + headers.map(h => escapeForTeams(h.name)).join(' | ') + ' |';
        const separatorLine = '| ' + headers.map(() => '---').join(' | ') + ' |';
        const dataLines = [];

        for (const row of rows) {
            const cells = headers.map(h => escapeForTeams(getCellValue(row, h.colId)));
            dataLines.push('| ' + cells.join(' | ') + ' |');
        }

        // Teams needs blank lines before and after the table
        return '\n' + headerLine + '\n' + separatorLine + '\n' + dataLines.join('\n') + '\n';
    }

    function escapeHTML(str) {
        return str
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&#039;');
    }

    function buildHTMLTable(headers, rows) {
        if (headers.length === 0 || rows.length === 0) return '';

        let html = '<table border="1" cellpadding="4" cellspacing="0" style="border-collapse: collapse; font-family: Segoe UI, sans-serif; font-size: 14px;">\n';

        // Header row
        html += '  <thead style="background-color: #0078d4; color: white;">\n    <tr>\n';
        for (const header of headers) {
            html += `      <th style="padding: 8px; text-align: left; border: 1px solid #ddd;">${escapeHTML(header.name)}</th>\n`;
        }
        html += '    </tr>\n  </thead>\n';

        // Data rows
        html += '  <tbody>\n';
        for (let i = 0; i < rows.length; i++) {
            const row = rows[i];
            const bgColor = i % 2 === 0 ? '#f9f9f9' : '#ffffff';
            html += `    <tr style="background-color: ${bgColor};">\n`;
            for (const header of headers) {
                const value = getCellValue(row, header.colId);
                html += `      <td style="padding: 8px; border: 1px solid #ddd;">${escapeHTML(value)}</td>\n`;
            }
            html += '    </tr>\n';
        }
        html += '  </tbody>\n</table>';

        return html;
    }

    function showFeedback(button, message, isError) {
        const originalHTML = button.innerHTML;
        const originalTitle = button.title;
        const originalBgColor = button.style.backgroundColor;
        const originalBorderColor = button.style.borderColor;

        button.innerHTML = `<span>${message}</span>`;
        button.title = message;
        button.style.color = 'white';

        if (isError) {
            button.style.backgroundColor = '#d13438';
            button.style.borderColor = '#d13438';
        } else {
            button.style.backgroundColor = '#107c10';
            button.style.borderColor = '#107c10';
        }

        setTimeout(() => {
            button.innerHTML = originalHTML;
            button.title = originalTitle;
            button.style.backgroundColor = originalBgColor;
            button.style.borderColor = originalBorderColor;
            button.style.color = 'white';
        }, COPIED_FEEDBACK_MS);
    }

    function handleCopyMarkdown() {
        const headers = getColumnHeaders();
        const selectedRows = getSelectedRows();

        console.log('[LAezPaste] Copy Markdown clicked - Headers:', headers.length, 'Selected rows:', selectedRows.length);

        if (selectedRows.length === 0) {
            showFeedback(this, 'Select rows first!', true);
            return;
        }

        if (headers.length === 0) {
            showFeedback(this, 'No columns found!', true);
            return;
        }

        const table = buildMarkdownTable(headers, selectedRows);
        console.log('[LAezPaste] Generated markdown table:\n', table);

        // Copy to clipboard
        navigator.clipboard.writeText(table).then(() => {
            console.log('[LAezPaste] Copied to clipboard successfully!');
            showFeedback(this, `Copied ${selectedRows.length} row(s)!`, false);
        }).catch(err => {
            console.error('[LAezPaste] Clipboard write failed:', err);
            // Fallback method
            const ta = document.createElement('textarea');
            ta.value = table;
            ta.style.position = 'fixed';
            ta.style.opacity = '0';
            document.body.appendChild(ta);
            ta.select();
            document.execCommand('copy');
            document.body.removeChild(ta);
            showFeedback(this, `Copied ${selectedRows.length} row(s)!`, false);
        });
    }

    function handleCopyHTML() {
        const headers = getColumnHeaders();
        const selectedRows = getSelectedRows();

        console.log('[LAezPaste] Copy HTML clicked - Headers:', headers.length, 'Selected rows:', selectedRows.length);

        if (selectedRows.length === 0) {
            showFeedback(this, 'Select rows first!', true);
            return;
        }

        if (headers.length === 0) {
            showFeedback(this, 'No columns found!', true);
            return;
        }

        const htmlTable = buildHTMLTable(headers, selectedRows);
        const plainText = buildMarkdownTable(headers, selectedRows);
        console.log('[LAezPaste] Generated HTML table');

        // Copy both HTML and plain text to clipboard
        const clipboardItem = new ClipboardItem({
            'text/html': new Blob([htmlTable], { type: 'text/html' }),
            'text/plain': new Blob([plainText], { type: 'text/plain' })
        });

        navigator.clipboard.write([clipboardItem]).then(() => {
            console.log('[LAezPaste] Copied HTML to clipboard successfully!');
            showFeedback(this, `Copied ${selectedRows.length} row(s) as HTML!`, false);
        }).catch(err => {
            console.error('[LAezPaste] Clipboard write failed:', err);
            // Fallback: just copy HTML as text
            navigator.clipboard.writeText(htmlTable).then(() => {
                showFeedback(this, `Copied ${selectedRows.length} row(s) as HTML!`, false);
            });
        });
    }

    function createMarkdownButton() {
        const btn = document.createElement('button');
        btn.id = BUTTON_ID_MARKDOWN;
        btn.className = 'query-command-button la-button';
        btn.title = 'Copy selected rows as Markdown table';
        btn.style.backgroundColor = '#0078d4';
        btn.style.color = 'white';
        btn.style.border = '1px solid #0078d4';
        btn.style.padding = '6px 12px';
        btn.style.borderRadius = '2px';
        btn.style.cursor = 'pointer';
        btn.style.display = 'inline-flex';
        btn.style.alignItems = 'center';
        btn.innerHTML = `
            <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 16 16" fill="white" style="margin-right: 4px; vertical-align: middle;">
                <path d="M4 4V1.5C4 .67 4.67 0 5.5 0h9C15.33 0 16 .67 16 1.5v9c0 .83-.67 1.5-1.5 1.5H12v2.5c0 .83-.67 1.5-1.5 1.5h-9C.67 16 0 15.33 0 14.5v-9C0 4.67.67 4 1.5 4H4zm1.5-3a.5.5 0 0 0-.5.5V4h5.5c.83 0 1.5.67 1.5 1.5V12h2.5a.5.5 0 0 0 .5-.5v-9a.5.5 0 0 0-.5-.5h-9zM1 5.5a.5.5 0 0 1 .5-.5h9a.5.5 0 0 1 .5.5v9a.5.5 0 0 1-.5.5h-9a.5.5 0 0 1-.5-.5v-9z"/>
            </svg>
            <span>Markdown</span>`;
        btn.addEventListener('click', handleCopyMarkdown);
        btn.addEventListener('mouseenter', () => {
            btn.style.backgroundColor = '#005a9e';
            btn.style.borderColor = '#005a9e';
        });
        btn.addEventListener('mouseleave', () => {
            btn.style.backgroundColor = '#0078d4';
            btn.style.borderColor = '#0078d4';
        });
        return btn;
    }

    function createHTMLButton() {
        const btn = document.createElement('button');
        btn.id = BUTTON_ID_HTML;
        btn.className = 'query-command-button la-button';
        btn.title = 'Copy selected rows as HTML table (paste into Teams/Outlook/Word)';
        btn.style.backgroundColor = '#0078d4';
        btn.style.color = 'white';
        btn.style.border = '1px solid #0078d4';
        btn.style.padding = '6px 12px';
        btn.style.borderRadius = '2px';
        btn.style.cursor = 'pointer';
        btn.style.display = 'inline-flex';
        btn.style.alignItems = 'center';
        btn.innerHTML = `
            <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 16 16" fill="white" style="margin-right: 4px; vertical-align: middle;">
                <path d="M0 1.5A1.5 1.5 0 0 1 1.5 0h13A1.5 1.5 0 0 1 16 1.5v13a1.5 1.5 0 0 1-1.5 1.5h-13A1.5 1.5 0 0 1 0 14.5v-13zM1.5 1a.5.5 0 0 0-.5.5v13a.5.5 0 0 0 .5.5h13a.5.5 0 0 0 .5-.5v-13a.5.5 0 0 0-.5-.5h-13z"/>
                <path d="M2 3h12v2H2V3zm0 3h12v2H2V6zm0 3h12v2H2V9zm0 3h12v2H2v-2z"/>
            </svg>
            <span>HTML</span>`;
        btn.addEventListener('click', handleCopyHTML);
        btn.addEventListener('mouseenter', () => {
            btn.style.backgroundColor = '#005a9e';
            btn.style.borderColor = '#005a9e';
        });
        btn.addEventListener('mouseleave', () => {
            btn.style.backgroundColor = '#0078d4';
            btn.style.borderColor = '#0078d4';
        });
        return btn;
    }

    function injectButton() {
        // Already injected?
        if (document.getElementById(BUTTON_ID_MARKDOWN) && document.getElementById(BUTTON_ID_HTML)) {
            return true;
        }

        // Find the Add bookmark toolbar
        const bookmarkToolbar = document.querySelector('kendo-toolbar.add-bookmark-toolbar');
        if (!bookmarkToolbar) {
            return false;
        }

        console.log('[LAezPaste] Found bookmark toolbar, injecting buttons...');

        // Create both buttons
        const btnMarkdown = createMarkdownButton();
        const btnHTML = createHTMLButton();

        // Create a wrapper matching the toolbar style
        const wrapper = document.createElement('div');
        wrapper.style.display = 'inline-flex';
        wrapper.style.alignItems = 'center';
        wrapper.style.marginLeft = '4px';
        wrapper.style.gap = '4px';
        wrapper.appendChild(btnMarkdown);
        wrapper.appendChild(btnHTML);

        // Insert after the bookmark toolbar
        bookmarkToolbar.parentElement.insertBefore(wrapper, bookmarkToolbar.nextSibling);

        console.log('[LAezPaste] Buttons injected successfully!');
        return true;
    }

    function init() {
        console.log('[LAezPaste] Content script loaded on:', window.location.href);
        console.log('[LAezPaste] Frame type:', window === window.top ? 'TOP' : 'IFRAME');

        // For the parent page, just exit
        if (window === window.top) {
            console.log('[LAezPaste] Running on parent page, exiting');
            return;
        }

        // We're in an iframe - check if it's the right one (reactblade)
        if (!window.location.hostname.includes('reactblade')) {
            console.log('[LAezPaste] Not in reactblade iframe, exiting');
            return;
        }

        console.log('[LAezPaste] Inside reactblade iframe! Setting up observers...');

        // CONTINUOUSLY observe for the toolbar to appear
        // (it won't exist until the user runs a query)
        const observer = new MutationObserver(() => {
            // Check if Log Analytics elements appeared
            if (isLogAnalyticsPage()) {
                console.log('[LAezPaste] Log Analytics elements detected!');
                injectButton();
            }
        });

        // Start observing immediately
        observer.observe(document.documentElement, {
            childList: true,
            subtree: true
        });

        // Also do periodic checks (in case MutationObserver misses something)
        setInterval(() => {
            if (isLogAnalyticsPage()) {
                injectButton();
            }
        }, TOOLBAR_CHECK_INTERVAL);

        console.log('[LAezPaste] Observers active - waiting for query results...');
    }

    // Wait for DOM ready then init
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', init);
    } else {
        init();
    }
})();
