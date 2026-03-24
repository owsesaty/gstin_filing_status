document.addEventListener('DOMContentLoaded', () => {
    const form = document.getElementById('search-form');
    const submitBtn = document.getElementById('submit-btn');
    const btnText = document.querySelector('.btn-text');
    const loader = document.querySelector('.loader');
    
    const resultsContainer = document.getElementById('results-container');
    const resultsThead = document.getElementById('results-thead');
    const resultsTbody = document.getElementById('results-tbody');
    const progressText = document.getElementById('progress-text');
    const errorMsg = document.getElementById('error-message');

    const filterRtnType = document.getElementById('filter-rtntype');
    const filterMonth = document.getElementById('filter-month');
    const monthFilterGroup = document.getElementById('month-filter-group');
    const viewToggle = document.getElementById('view-toggle');
    const toggleLabelText = document.getElementById('toggle-label-text');
    const exportBtn = document.getElementById('export-btn');
    
    const excelUpload = document.getElementById('excel-upload');
    const gstinTextarea = document.getElementById('gstin-list');

    // Pagination elements
    const paginationControls = document.getElementById('pagination-controls');
    const prevPageBtn = document.getElementById('prev-page');
    const nextPageBtn = document.getElementById('next-page');
    const pageInfo = document.getElementById('page-info');

    // Global store
    let allRecords = [];
    let seenReturnTypes = new Set();
    let seenMonths = new Set();

    let currentPage = 1;
    const ROWS_PER_PAGE = 200;

    // Standard Indian Financial Year Months
    const FY_MONTHS = ['April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December', 'January', 'February', 'March'];
    const monthRank = (m) => {
        let idx = FY_MONTHS.indexOf(m);
        return idx === -1 ? 999 : idx;
    };

    // Handle Excel parsing
    excelUpload.addEventListener('change', (e) => {
        const file = e.target.files[0];
        if (!file) return;

        const reader = new FileReader();
        reader.onload = function(evt) {
            try {
                const data = evt.target.result;
                const workbook = XLSX.read(data, { type: 'binary' });
                
                let extractedGstins = new Set();
                const gstinRegex = /[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z]{1}[1-9A-Z]{1}Z[0-9A-Z]{1}/i;

                workbook.SheetNames.forEach(sheetName => {
                    const ws = workbook.Sheets[sheetName];
                    const json = XLSX.utils.sheet_to_json(ws, { header: 1 });
                    
                    json.forEach(row => {
                        row.forEach(cell => {
                            if (cell && typeof cell === 'string') {
                                const match = cell.match(gstinRegex);
                                if (match) {
                                    extractedGstins.add(match[0].toUpperCase());
                                }
                            }
                        });
                    });
                });

                if (extractedGstins.size > 0) {
                    const existing = gstinTextarea.value ? gstinTextarea.value.trim() : '';
                    const combined = new Set([...existing.split(/[\n,]+/).map(i => i.trim()).filter(Boolean), ...Array.from(extractedGstins)]);
                    gstinTextarea.value = Array.from(combined).join('\n');
                    alert(`Successfully extracted ${extractedGstins.size} GSTIN(s) from the file!`);
                } else {
                    alert('No valid GSTINs found in the uploaded file.');
                }
            } catch (err) {
                alert('Error parsing file: ' + err.message);
            }
        };
        reader.readAsBinaryString(file);
    });

    form.addEventListener('submit', async (e) => {
        e.preventDefault();
        const gstinInputText = gstinTextarea.value;
        const fy = document.getElementById('fy').value.trim();
        const gstins = gstinInputText.split(/[\n,]+/).map(id => id.trim()).filter(id => id.length > 0);

        if (gstins.length === 0) {
            showError("Please enter at least one valid GSTIN.");
            return;
        }

        const MAX_GSTINS = 1000;
        let finalGstins = gstins;

        if (finalGstins.length > MAX_GSTINS) {
            alert(`For system stability, the bulk processor is limited to ${MAX_GSTINS} GSTINs at a time. The list will be automatically truncated. Please split larger datasets into multiple smaller files.`);
            finalGstins = finalGstins.slice(0, MAX_GSTINS);
            gstinTextarea.value = finalGstins.join('\n'); // update textarea to reflect what's actually being processed
        }

        allRecords = [];
        seenReturnTypes.clear();
        seenMonths.clear();
        updateFilterDropdowns();
        
        resultsTbody.innerHTML = '<tr><td colspan="6" class="empty-state">Sending bulk payload to server... Please wait. (This may take a minute)</td></tr>';
        resultsContainer.classList.remove('hidden');
        paginationControls.classList.add('hidden');
        startSearch(finalGstins.length);

        try {
            const chunkSize = 50; // Process 50 GSTINs per bulk request
            let totalProcessed = 0;

            for (let i = 0; i < finalGstins.length; i += chunkSize) {
                const chunk = finalGstins.slice(i, i + chunkSize);

                const response = await fetch('/api/search_bulk', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ gstins: chunk, fy: fy })
                });

                const data = await response.json();
                if (!response.ok) throw new Error(data.error || `Bulk Error ${response.status}`);
                
                if (data.results) {
                    for (const gstin of Object.keys(data.results)) {
                        const rData = data.results[gstin];
                        
                        if (rData.error) {
                            allRecords.push({
                                gstin: gstin, fy: fy, month: '-', rtntype: '-', dof: '-', status: `Error: ${rData.error}`, mof: '-', arn: '-'
                            });
                            continue;
                        }
                        
                        if (rData.filingStatus && Array.isArray(rData.filingStatus) && rData.filingStatus.length > 0 && Array.isArray(rData.filingStatus[0])) {
                            rData.filingStatus[0].forEach(r => {
                                const isNotFiled = r.status && r.status.toLowerCase().includes('not filed');
                                allRecords.push({
                                    gstin: gstin,
                                    fy: r.fy || fy,
                                    month: r.taxp || 'Unknown',
                                    rtntype: r.rtntype || 'Unknown',
                                    dof: isNotFiled ? '-' : (r.dof || '-'),
                                    status: r.status || 'Unknown',
                                    mof: r.mof || '-',
                                    arn: r.arn || '-'
                                });
                                if (r.rtntype) seenReturnTypes.add(r.rtntype);
                                if (r.taxp && r.taxp !== 'Unknown') seenMonths.add(r.taxp);
                            });
                        } else {
                            allRecords.push({
                                gstin: gstin, fy: fy, month: '-', rtntype: '-', dof: '-', status: rData.status || 'No Records Found', mof: '-', arn: '-'
                            });
                        }
                    }
                } else {
                     showError("Received malformed bulk response from server.");
                }

                totalProcessed += chunk.length;
                progressText.textContent = `${totalProcessed} / ${finalGstins.length} Completed`;
            }
        } catch (err) {
            showError("Failed to fetch bulk data: " + err.message);
        }

        finishSearch();
        updateFilterDropdowns();
        currentPage = 1;
        renderTable();
    });

    filterRtnType.addEventListener('change', () => { currentPage = 1; renderTable(); });
    filterMonth.addEventListener('change', () => { currentPage = 1; renderTable(); });
    
    viewToggle.addEventListener('change', () => {
        currentPage = 1;
        const isPivot = viewToggle.checked;
        toggleLabelText.textContent = isPivot ? 'Pivot View' : 'List View';
        monthFilterGroup.style.display = isPivot ? 'none' : 'flex';
        renderTable();
    });

    exportBtn.addEventListener('click', () => {
        exportToCsv();
    });

    prevPageBtn.addEventListener('click', () => {
        if (currentPage > 1) {
            currentPage--;
            renderTable();
        }
    });

    nextPageBtn.addEventListener('click', () => {
        currentPage++;
        renderTable();
    });

    function updateFilterDropdowns() {
        const currRtn = filterRtnType.value;
        const currMonth = filterMonth.value;

        filterRtnType.innerHTML = '<option value="ALL">All Types</option>';
        [...seenReturnTypes].sort().forEach(rt => {
            const opt = document.createElement('option');
            opt.value = rt; opt.textContent = rt;
            filterRtnType.appendChild(opt);
        });
        if (seenReturnTypes.has(currRtn)) filterRtnType.value = currRtn;

        filterMonth.innerHTML = '<option value="ALL">All Months</option>';
        [...seenMonths].sort((a,b) => monthRank(a) - monthRank(b)).forEach(m => {
            const opt = document.createElement('option');
            opt.value = m; opt.textContent = m;
            filterMonth.appendChild(opt);
        });
        if (seenMonths.has(currMonth)) filterMonth.value = currMonth;
    }

    function renderTable() {
        const isPivot = viewToggle.checked;
        const fRtn = filterRtnType.value;
        const fMnth = filterMonth.value;

        let records = allRecords;
        if (fRtn !== 'ALL') records = records.filter(r => r.rtntype === fRtn);
        if (!isPivot && fMnth !== 'ALL') records = records.filter(r => r.month === fMnth);

        if (records.length === 0) {
            resultsThead.innerHTML = `<tr><th>Results</th></tr>`;
            resultsTbody.innerHTML = '<tr><td class="empty-state">No records to display.</td></tr>';
            paginationControls.classList.add('hidden');
            return;
        }

        if (isPivot) renderPivotedTable(records);
        else renderListTable(records);
    }

    function renderListTable(records) {
        // Handle pagination
        const totalPages = Math.ceil(records.length / ROWS_PER_PAGE);
        if (currentPage > totalPages) currentPage = totalPages;
        
        const startIndex = (currentPage - 1) * ROWS_PER_PAGE;
        const endIndex = startIndex + ROWS_PER_PAGE;
        const pagedRecords = records.slice(startIndex, endIndex);

        resultsThead.innerHTML = `
            <tr>
                <th>GSTIN</th>
                <th>Month</th>
                <th>Return Type</th>
                <th>Date of Filing</th>
                <th>Status</th>
                <th>ARN</th>
            </tr>
        `;
        let html = '';
        pagedRecords.forEach(item => {
            let badgeClass = 'status-success';
            const sL = item.status.toLowerCase();
            if (sL.includes('pending') || sL.includes('not')) badgeClass = 'status-pending';
            else if (sL.includes('error') || sL.includes('no records')) badgeClass = 'status-error';

            html += `
                <tr>
                    <td style="font-family: monospace;">${item.gstin}</td>
                    <td><strong>${item.month}</strong></td>
                    <td>${item.rtntype}</td>
                    <td>${item.dof !== '-' ? item.dof : '-'}</td>
                    <td><span class="status-badge ${badgeClass}">${item.status}</span></td>
                    <td><span style="font-size: 0.85rem; color: var(--text-muted); font-family: monospace;">${item.arn}</span></td>
                </tr>
            `;
        });
        resultsTbody.innerHTML = html;

        updatePaginationUI(totalPages);
    }

    function renderPivotedTable(records) {
        const grouped = {};
        const availableMonths = new Set();
        
        records.forEach(r => {
            if (r.month === '-' || r.rtntype === '-') return; 
            availableMonths.add(r.month);
            const key = `${r.gstin}___${r.rtntype}`;
            if (!grouped[key]) grouped[key] = { gstin: r.gstin, rtntype: r.rtntype, months: {} };
            grouped[key].months[r.month] = r.dof;
        });

        const pivotedArray = Object.values(grouped);
        
        // Handle pagination
        const totalPages = Math.ceil(pivotedArray.length / ROWS_PER_PAGE);
        if (currentPage > totalPages && totalPages > 0) currentPage = totalPages;
        
        const startIndex = (currentPage - 1) * ROWS_PER_PAGE;
        const endIndex = startIndex + ROWS_PER_PAGE;
        const pagedRecords = pivotedArray.slice(startIndex, endIndex);

        const cols = [...availableMonths].sort((a,b) => monthRank(a) - monthRank(b));
        let theadHtml = `<tr><th>GSTIN</th><th>Return Type</th>`;
        cols.forEach(c => theadHtml += `<th>${c}</th>`);
        theadHtml += `</tr>`;
        resultsThead.innerHTML = theadHtml;

        let bodyHtml = '';
        pagedRecords.forEach(g => {
            bodyHtml += `<tr>
                <td style="font-family: monospace;">${g.gstin}</td>
                <td><strong>${g.rtntype}</strong></td>`;
            cols.forEach(c => {
                const val = g.months[c];
                if (val && val !== '-') {
                    bodyHtml += `<td><span class="status-badge status-success" style="background:transparent; border:1px solid var(--success);">${val}</span></td>`;
                } else {
                    bodyHtml += `<td style="color:var(--text-muted); text-align:center;">-</td>`;
                }
            });
            bodyHtml += `</tr>`;
        });

        if (pivotedArray.length === 0) {
            resultsTbody.innerHTML = `<tr><td colspan="${cols.length + 2}" class="empty-state">No valid monthly data to pivot.</td></tr>`;
            paginationControls.classList.add('hidden');
        } else {
            resultsTbody.innerHTML = bodyHtml;
            updatePaginationUI(totalPages);
        }
    }

    function updatePaginationUI(totalPages) {
        if (totalPages <= 1) {
            paginationControls.classList.add('hidden');
            return;
        }
        paginationControls.classList.remove('hidden');
        pageInfo.textContent = `Page ${currentPage} of ${totalPages}`;
        
        prevPageBtn.disabled = currentPage === 1;
        nextPageBtn.disabled = currentPage === totalPages;
    }

    function exportToCsv() {
        const isPivot = viewToggle.checked;
        const fRtn = filterRtnType.value;
        const fMnth = filterMonth.value;

        let records = allRecords;
        if (fRtn !== 'ALL') records = records.filter(r => r.rtntype === fRtn);
        if (!isPivot && fMnth !== 'ALL') records = records.filter(r => r.month === fMnth);

        if (records.length === 0) {
            alert("No data to export!");
            return;
        }

        const escapeCSV = (str) => {
            let text = String(str).replace(/"/g, '""');
            if (text.search(/("|,|\n)/g) >= 0) text = `"${text}"`;
            return text;
        };

        let csvContent = '';

        if (!isPivot) {
            const headers = ['GSTIN', 'Financial Year', 'Month', 'Return Type', 'Date of Filing', 'Status', 'Mode of Filing', 'ARN'];
            csvContent += headers.join(',') + '\n';
            records.forEach(r => {
                const row = [r.gstin, r.fy, r.month, r.rtntype, r.dof, r.status, r.mof, r.arn].map(escapeCSV);
                csvContent += row.join(',') + '\n';
            });
        } else {
            const grouped = {};
            const availableMonths = new Set();
            records.forEach(r => {
                if (r.month === '-' || r.rtntype === '-') return; 
                availableMonths.add(r.month);
                const key = `${r.gstin}___${r.rtntype}`;
                if (!grouped[key]) grouped[key] = { gstin: r.gstin, rtntype: r.rtntype, months: {} };
                grouped[key].months[r.month] = r.dof;
            });

            const cols = [...availableMonths].sort((a,b) => monthRank(a) - monthRank(b));
            const headers = ['GSTIN', 'Return Type', ...cols];
            csvContent += headers.join(',') + '\n';

            Object.values(grouped).forEach(g => {
                const row = [g.gstin, g.rtntype];
                cols.forEach(c => row.push(g.months[c] || 'Null'));
                csvContent += row.map(escapeCSV).join(',') + '\n';
            });
        }

        const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
        const url = URL.createObjectURL(blob);
        const link = document.createElement("a");
        link.setAttribute("href", url);
        
        const pad = (n) => n.toString().padStart(2, '0');
        const d = new Date();
        const dateTimeStr = `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}_${pad(d.getHours())}-${pad(d.getMinutes())}`;
        
        link.setAttribute("download", `GSTIN_Filing_${isPivot ? 'Pivoted' : 'Details'}_${dateTimeStr}.csv`);
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
    }

    function startSearch(total) {
        submitBtn.disabled = true;
        btnText.style.display = 'none';
        loader.style.display = 'block';
        hideError();
        progressText.textContent = `0 / ${total} Completed`;
    }

    function finishSearch() {
        submitBtn.disabled = false;
        btnText.style.display = 'block';
        loader.style.display = 'none';
    }

    function showError(message) {
        errorMsg.textContent = message;
        errorMsg.classList.remove('hidden');
    }

    function hideError() {
        errorMsg.classList.add('hidden');
    }
});
