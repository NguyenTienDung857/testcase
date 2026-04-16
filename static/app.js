document.addEventListener('DOMContentLoaded', () => {
    const btnSelect = document.getElementById('btn-select');
    const btnClear = document.getElementById('btn-clear');
    const btnRun = document.getElementById('btn-run');
    const fileList = document.getElementById('file-list');
    const terminal = document.getElementById('terminal');
    const backupCheckbox = document.getElementById('backup-mode');

    let currentFiles = [];

    // Columns generator K to Z
    const getColumnOptions = (selectedCol) => {
        let options = '';
        for (let i = 1; i <= 26; i++) {
            const colName = String.fromCharCode(64 + i);
            const isSelected = colName === selectedCol ? 'selected' : '';
            options += `<option value="${colName}" ${isSelected}>${colName}</option>`;
        }
        return options;
    };

    const renderFileList = () => {
        if (currentFiles.length === 0) {
            fileList.innerHTML = `<div class="empty-state">No payloads detected. Click "INITIALIZE FILES [ + ]" to append target nodes.</div>`;
            btnRun.disabled = true;
            return;
        }

        btnRun.disabled = false;
        fileList.innerHTML = '';
        currentFiles.forEach((fileObj, index) => {
            const div = document.createElement('div');
            div.className = 'file-row';
            
            // Extract filename from full path
            const filename = fileObj.path.split('\\').pop().split('/').pop();

            div.innerHTML = `
                <div class="col-id">[0${index + 1}]</div>
                <div class="col-path" title="${fileObj.path}">// ${filename}</div>
                <div class="col-col">
                    <select data-idx="${index}">
                        ${getColumnOptions(fileObj.column)}
                    </select>
                </div>
                <div class="col-action">
                    <!-- Open Button -->
                    <button class="btn-icon btn-open" data-idx="${index}" title="Open this file">
                        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" width="16" height="16"><path d="M2 3h6a4 4 0 0 1 4 4v14a3 3 0 0 0-3-3H2z"/><path d="M22 3h-6a4 4 0 0 0-4 4v14a3 3 0 0 1 3-3h7z"/></svg>
                    </button>
                    <!-- Single Run Button -->
                    <button class="btn-icon btn-run-single" data-idx="${index}" title="Execute this file">
                        <svg viewBox="0 0 24 24" fill="currentColor" width="16" height="16"><path d="M8 5v14l11-7z"/></svg>
                    </button>
                    <!-- Remove Button -->
                    <button class="btn-icon btn-remove" data-idx="${index}" title="Remove from queue">
                        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" width="16" height="16"><path d="M18 6L6 18M6 6l12 12"/></svg>
                    </button>
                </div>
            `;
            fileList.appendChild(div);
        });

        // Add event listeners for selects and buttons
        document.querySelectorAll('.file-row select').forEach(select => {
            select.addEventListener('change', (e) => {
                const idx = parseInt(e.target.getAttribute('data-idx'));
                currentFiles[idx].column = e.target.value;
            });
        });

        document.querySelectorAll('.btn-remove').forEach(btn => {
            btn.addEventListener('click', (e) => {
                const idx = parseInt(e.currentTarget.getAttribute('data-idx'));
                currentFiles.splice(idx, 1);
                renderFileList();
            });
        });

        document.querySelectorAll('.btn-run-single').forEach(btn => {
            btn.addEventListener('click', (e) => {
                const idx = parseInt(e.currentTarget.getAttribute('data-idx'));
                runTasks([currentFiles[idx]]);
            });
        });

        document.querySelectorAll('.btn-open').forEach(btn => {
            btn.addEventListener('click', async (e) => {
                const idx = parseInt(e.currentTarget.getAttribute('data-idx'));
                const fileObj = currentFiles[idx];
                appendLog(`Opening payload [0${idx + 1}]...`, 'log-sys');
                // Poll logs to show saving/closing actions dynamically
                if (!logInterval) {
                     pollLogs();
                     logInterval = setInterval(pollLogs, 500);
                }
                
                try {
                    await fetch('/api/open_file', {
                        method: 'POST',
                        headers: { 'Content-Type': 'application/json' },
                        body: JSON.stringify({ path: fileObj.path })
                    });
                } catch(e) {
                    appendLog('[ERROR] Open request failed.', 'log-error');
                }
                setTimeout(() => { if(logInterval && !btnRun.disabled) { clearInterval(logInterval); logInterval = null; } }, 2000); // Stop polling after a few seconds if no other task is running
            });
        });
    };

    const appendLog = (text, type = 'log-sys') => {
        const line = document.createElement('div');
        line.className = type;
        
        const now = new Date();
        const timeStr = now.toTimeString().split(' ')[0];
        
        line.innerHTML = `[${timeStr}] >> ${text}`;
        terminal.appendChild(line);
        terminal.scrollTop = terminal.scrollHeight;
    };

    let logInterval = null;
    const pollLogs = async () => {
        try {
            const res = await fetch('/api/logs');
            const data = await res.json();
            data.logs.forEach(msg => appendLog(msg, 'log-success'));
        } catch (err) {
            console.error('Terminal polling interrupted', err);
        }
    };

    const runTasks = async (tasksToRun) => {
        if (!tasksToRun || tasksToRun.length === 0) return;

        btnRun.disabled = true;
        btnSelect.disabled = true;
        document.querySelectorAll('.btn-icon').forEach(b => b.disabled = true);
        
        appendLog(`Executing directives for ${tasksToRun.length} payload(s)...`, 'log-sys');

        logInterval = setInterval(pollLogs, 500);

        try {
            const res = await fetch('/api/run', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    make_backup: backupCheckbox.checked,
                    tasks: tasksToRun
                })
            });
            const data = await res.json();
            await pollLogs(); // Final poll
            
            clearInterval(logInterval);
            if (data.status === 'success') {
                appendLog(`[OPERATION COMPLETE] Resized: ${data.total_resized} | Pictures: ${data.total_pictures} | Exceptions: ${data.total_errors}`, 'log-success');
            } else {
                appendLog(`[ERROR] ${data.message}`, 'log-error');
            }

        } catch (err) {
            clearInterval(logInterval);
            appendLog('[CRITICAL] Server / Network Failure.', 'log-error');
        } finally {
            btnRun.disabled = false;
            btnSelect.disabled = false;
            document.querySelectorAll('.btn-icon').forEach(b => b.disabled = false);
            appendLog('System Ready.', 'log-sys');
            renderFileList();
        }
    };

    btnSelect.addEventListener('click', async () => {
        try {
            const res = await fetch('/api/select_files', { method: 'POST' });
            const data = await res.json();
            if (data.files && data.files.length > 0) {
                data.files.forEach(path => {
                    if (!currentFiles.find(f => f.path === path)) {
                        currentFiles.push({ path: path, column: 'K' });
                    }
                });
                renderFileList();
                appendLog(`Acquired ${data.files.length} payload(s).`, 'log-sys');
            }
        } catch (err) {
            appendLog('API Error: File selection failed.', 'log-error');
        }
    });

    btnClear.addEventListener('click', () => {
        currentFiles = [];
        renderFileList();
        appendLog('Queue purged.', 'log-sys');
    });

    btnRun.addEventListener('click', () => {
        runTasks(currentFiles);
    });

    // Start UI
    renderFileList();
});
