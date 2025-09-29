let excelData = [];
let headerNames = [];

// Load Excel
async function loadExcel() {
    const res = await fetch('source/data.xlsx');
    const data = await res.arrayBuffer();
    const workbook = XLSX.read(data, { type: 'array' });
    const sheet = workbook.Sheets['EU Doc Archive'];
    const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });
    headerNames = jsonData[1]; // row 2 = headers
    excelData = jsonData.slice(2); // row 3+ = data
    populateTable(excelData);
    updateCounter(excelData.length);
}

// Populate table
function populateTable(data) {
    const tbody = document.querySelector('.table tbody');
    tbody.innerHTML = '';
    data.slice(0, 100).forEach(r => {
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td>${r[0]||''}</td>
            <td>${r[1]||''}</td>
            <td>${r[2]||''}</td>
            <td>${r[4]||''}</td>
            <td>${r[5]||''}</td>
            <td>${excelDateToJS(r[16])}</td>
            <td>${excelDateToJS(r[17])}</td>
            <td>${r[26]||''}</td>
        `;
        tr.addEventListener('click', () => openForm(r));
        tbody.appendChild(tr);
    });
}

// Counter
function updateCounter(count) {
    const counter = document.getElementById('itemCounter');
    if (counter) counter.textContent = `${count} items`;
}

// Filter table
function filterTable() {
    const searchTerm = document.querySelector('.search').value.toLowerCase();
    const countryVal = document.getElementById('countryFilter').value.toLowerCase();
    const ownerVal = document.getElementById('ownerFilter').value.toLowerCase();
    const completedSort = document.getElementById('completedFilter').value;

    let filtered = excelData.filter(r => {
        let keep = true;
        if (countryVal) keep = keep && r[0].toString().toLowerCase() === countryVal;
        if (ownerVal) keep = keep && r[26].toString().toLowerCase() === ownerVal;
        if (searchTerm) {
            keep = keep && r.some(cell => cell.toString().toLowerCase().includes(searchTerm));
        }
        return keep;
    });

    if (completedSort === 'asc') filtered.sort((a,b) => (''+a[17]).localeCompare(b[17]));
    if (completedSort === 'desc') filtered.sort((a,b) => (''+b[17]).localeCompare(a[17]));

    populateTable(filtered);
    updateCounter(filtered.length);
}

// Dropdown setup
function setupDropdown(id, colIdx) {
    const dropdown = document.getElementById(id);
    const unique = [...new Set(excelData.map(r => r[colIdx]))].sort();
    dropdown.innerHTML = '<option value="">All</option>';
    unique.forEach(v => {
        const opt = document.createElement('option');
        opt.value = v; 
        opt.textContent = v;
        dropdown.appendChild(opt);
    });
}

// Open form overlay
function openForm(row) {
    const overlay = document.getElementById('overlay');
    const leftSection = document.querySelector('.left-section');
    const rightSection = document.querySelector('.right-section');
    leftSection.innerHTML = ''; 
    rightSection.innerHTML = ''; // keep empty for left-right split

    // Right section columns in specific order
    const rightCols = [0,24,25,26,27,28,29,11,17,31,32,30,33]; // 0-based indices
    const leftCols = row.map((_, i) => i).filter(i => !rightCols.includes(i));

    const collapsibleCols = [8,20,21,22,23]; // left section collapsibles

    // LEFT section
    leftCols.forEach(i => {
        const val = excelDateToJS(row[i]) || row[i];
        const target = leftSection;

        if (collapsibleCols.includes(i)) {
            const div = document.createElement('div'); 
            div.classList.add('collapsible');

            const header = document.createElement('div'); 
            header.classList.add('field-header'); 
            header.textContent = headerNames[i];

            const hint = document.createElement('div');
            hint.classList.add('hint');
            hint.textContent = "Click on the icon to show content";

            header.addEventListener('click', ()=> {
                div.classList.toggle('open');
                hint.style.display = div.classList.contains('open') ? 'none' : 'block';
            });

            const content = document.createElement('div'); 
            content.classList.add('field-content'); 
            content.textContent = val || '';

            div.appendChild(header); 
            div.appendChild(hint);
            div.appendChild(content); 
            target.appendChild(div);
        } else {
            const div = document.createElement('div'); 
            div.classList.add('form-field');

            const label = document.createElement('label'); 
            label.textContent = headerNames[i];

            const value = document.createElement('div'); 
            value.classList.add('value'); 
            value.textContent = val || '';

            div.appendChild(label); 
            div.appendChild(value); 
            target.appendChild(div);
        }
    });

    // RIGHT section in exact order
    rightCols.forEach(i => {
        const val = excelDateToJS(row[i]) || row[i];
        const div = document.createElement('div'); 
        div.classList.add('form-field');

        const label = document.createElement('label'); 
        label.textContent = headerNames[i];

        const value = document.createElement('div'); 
        value.classList.add('value'); 
        value.textContent = val || '';

        div.appendChild(label); 
        div.appendChild(value); 
        rightSection.appendChild(div);
    });

    overlay.classList.remove('hidden');
}

// Excel serial date to JS date
function excelDateToJS(serial) {
    if (!serial || isNaN(serial)) return '';
    const utc_days = Math.floor(serial - 25569);
    const utc_value = utc_days * 86400; 
    const date_info = new Date(utc_value * 1000);
    const year = date_info.getFullYear();
    const month = ('0' + (date_info.getMonth() + 1)).slice(-2);
    const day = ('0' + date_info.getDate()).slice(-2);
    return `${year}-${month}-${day}`;
}

// Init
window.onload = async () => {
    await loadExcel();
    setupDropdown('countryFilter', 0);
    setupDropdown('ownerFilter', 26);

    document.querySelector('.search').addEventListener('input', filterTable);
    document.getElementById('countryFilter').addEventListener('change', filterTable);
    document.getElementById('ownerFilter').addEventListener('change', filterTable);
    document.getElementById('completedFilter').addEventListener('change', filterTable);

    document.getElementById('overlay-close').addEventListener('click', () => 
        document.getElementById('overlay').classList.add('hidden')
    );

    const themeToggle = document.getElementById('themeToggle');
    themeToggle.addEventListener('click', () => {
        document.body.classList.toggle('light-theme');
        themeToggle.textContent = document.body.classList.contains('light-theme') ? '🌙' : '☀️';
    });
};