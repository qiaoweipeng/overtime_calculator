let currentFile = null;
let workbookData = null;
let selectedSheets = [];

const dropZone = document.getElementById('dropZone');
const fileInput = document.getElementById('fileInput');
const settings = document.getElementById('settings');
const sheetCheckboxes = document.getElementById('sheetCheckboxes');
const processBtn = document.getElementById('processBtn');

dropZone.addEventListener('click', () => fileInput.click());

fileInput.addEventListener('change', (e) => {
    if (e.target.files.length > 0) {
        handleFile(e.target.files[0]);
    }
});

dropZone.addEventListener('dragover', (e) => {
    e.preventDefault();
    dropZone.classList.add('dragover');
});

dropZone.addEventListener('dragleave', () => {
    dropZone.classList.remove('dragover');
});

dropZone.addEventListener('drop', (e) => {
    e.preventDefault();
    dropZone.classList.remove('dragover');
    const files = e.dataTransfer.files;
    if (files.length > 0 && files[0].name.endsWith('.xlsx')) {
        handleFile(files[0]);
    } else {
        showAlert('请上传 .xlsx 格式的 Excel 文件', 'error');
    }
});

function handleFile(file) {
    currentFile = file;
    dropZone.innerHTML = `
        <div class="icon">✅</div>
        <h3>已选择文件：${file.name}</h3>
        <p>大小：${(file.size / 1024).toFixed(1)} KB</p>
    `;

    loadWorkbook(file);
}

function loadWorkbook(file) {
    document.getElementById('loading').classList.add('show');

    const reader = new FileReader();
    reader.onload = (e) => {
        try {
            const data = new Uint8Array(e.target.result);
            workbookData = XLSX.read(data, { type: 'array' });

            renderSheetCheckboxes(workbookData.SheetNames);
            document.getElementById('loading').classList.remove('show');
            settings.style.display = 'block';
            document.getElementById('actionArea').style.display = 'block';
            settings.scrollIntoView({ behavior: 'smooth' });
        } catch (err) {
            showAlert('读取文件时出错：' + err.message, 'error');
            document.getElementById('loading').classList.remove('show');
        }
    };
    reader.readAsArrayBuffer(file);
}

function renderSheetCheckboxes(sheetNames) {
    sheetCheckboxes.innerHTML = '';
    selectedSheets = [];
    
    const allowedCities = ['深圳', '梅州', '长沙', '武汉', '广西'];
    
    allowedCities.forEach(city => {
        if (sheetNames.includes(city)) {
            const isDefaultSelected = city === '深圳';
            
            const checkbox = document.createElement('div');
            checkbox.className = 'sheet-checkbox' + (isDefaultSelected ? ' selected' : '');
            checkbox.innerHTML = `
                <input type="checkbox" ${isDefaultSelected ? 'checked' : ''} value="${city}">
                <span>${city}</span>
            `;
            
            checkbox.addEventListener('click', (e) => {
                const input = checkbox.querySelector('input');
                input.checked = !input.checked;
                checkbox.classList.toggle('selected', input.checked);
                updateSelectedSheets();
            });
            
            sheetCheckboxes.appendChild(checkbox);
            
            if (isDefaultSelected) {
                selectedSheets.push(city);
            }
        }
    });
}

function updateSelectedSheets() {
    const checkboxes = sheetCheckboxes.querySelectorAll('input[type="checkbox"]');
    selectedSheets = [];
    checkboxes.forEach(cb => {
        if (cb.checked) {
            selectedSheets.push(cb.value);
        }
    });
}

processBtn.addEventListener('click', () => {
    if (selectedSheets.length === 0) {
        showAlert('请至少选择一个工作表', 'warning');
        return;
    }
    processSelectedSheets();
});

function processSelectedSheets() {
    document.getElementById('loading').classList.add('show');
    
    setTimeout(() => {
        try {
            const allResults = [];
            
            selectedSheets.forEach(sheetName => {
                const worksheet = workbookData.Sheets[sheetName];
                const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
                const results = calculateOvertime(jsonData, sheetName);
                
                if (results) {
                    allResults.push({
                        sheetName: sheetName,
                        data: results
                    });
                }
            });
            
            if (allResults.length > 0) {
                displayResults(allResults, currentFile.name);
            }
            
            document.getElementById('loading').classList.remove('show');
        } catch (err) {
            showAlert('处理文件时出错：' + err.message, 'error');
            document.getElementById('loading').classList.remove('show');
        }
    }, 100);
}

function calculateOvertime(data, sheetName) {
    if (!data || data.length < 2) {
        showAlert(`工作表 "${sheetName}" 数据为空或格式不正确`, 'error');
        return null;
    }

    const headers = data[0];

    let targetMonth = null;
    for (const header of headers) {
        if (!header) continue;

        let headerStr = String(header);

        let match = headerStr.match(/^(\d{1,2})-\d{1,2}/);
        if (!match) match = headerStr.match(/^(\d{1,2})\/\d{1,2}/);
        if (!match) match = headerStr.match(/^\d{4}-(\d{1,2})-\d{2}/);
        if (!match) match = headerStr.match(/^(\d{1,2})月/);

        if (match) {
            targetMonth = match[1];
            break;
        }
    }

    if (!targetMonth) {
        showAlert(`工作表 "${sheetName}" 无法识别月份，请检查日期列格式`, 'error');
        return null;
    }

    const monthNum = targetMonth.replace(/^0+/, '');
    const datePatterns = [
        new RegExp(`^0?${monthNum}-\\d{1,2}(\\s?)星期[一二三四五六日]?$`),
        new RegExp(`^0?${monthNum}\\/\\d{1,2}(\\s?)星期[一二三四五六日]?$`),
        new RegExp(`^\\d{4}-0?${monthNum}-\\d{2}(\\s?)星期[一二三四五六日]?$`),
        new RegExp(`^0?${monthNum}月\\d{1,2}日?$`)
    ];

    const dateCols = [];
    for (let i = 0; i < headers.length; i++) {
        if (!headers[i]) continue;

        const headerStr = String(headers[i]);

        let matched = false;
        for (const pattern of datePatterns) {
            if (pattern.test(headerStr)) {
                dateCols.push({ col: i, header: headers[i] });
                matched = true;
                break;
            }
        }

        if (!matched && typeof headers[i] === 'number' && headers[i] > 40000 && headers[i] < 60000) {
            const date = XLSX.SSF.parse_date_code(headers[i]);
            if (date && date.m === parseInt(monthNum)) {
                const dateStr = `${date.m}-${date.d}`;
                dateCols.push({ col: i, header: dateStr });
            }
        }
    }

    if (dateCols.length === 0) {
        showAlert(`工作表 "${sheetName}" 未找到日期列，请检查日期列格式`, 'error');
        return null;
    }

    const firstDate = dateCols[0].header;
    const lastDate = dateCols[dateCols.length - 1].header;

    function extractDay(dateStr) {
        let dayMatch = String(dateStr).match(/-(\d{1,2})/);
        if (!dayMatch) dayMatch = String(dateStr).match(/\/(\d{1,2})/);
        if (!dayMatch) dayMatch = String(dateStr).match(/月(\d{1,2})/);
        return dayMatch ? parseInt(dayMatch[1]) : null;
    }

    const firstDay = extractDay(firstDate);
    const lastDay = extractDay(lastDate);
    const checkDateRange = firstDay && lastDay
        ? `${targetMonth}月${firstDay}日 - ${targetMonth}月${lastDay}日`
        : '';

    let overtimeTotalCol = -1;
    for (let i = 0; i < headers.length; i++) {
        if (headers[i] && String(headers[i]).includes('加班总时长')) {
            overtimeTotalCol = i;
            break;
        }
    }

    const overtimeStartMinutes = 19 * 60;

    const results = [];

    for (let rowIdx = 1; rowIdx < data.length; rowIdx++) {
        const row = data[rowIdx];
        if (!row || !row[0]) continue;

        const name = row[0];
        const dept = row[1] || '';

        let existingTotal = null;
        if (overtimeTotalCol >= 0 && row[overtimeTotalCol] !== undefined) {
            existingTotal = row[overtimeTotalCol];
        }

        let totalMinutes = 0;

        for (const dateCol of dateCols) {
            const cellVal = row[dateCol.col];
            if (cellVal === undefined || cellVal === null) continue;

            const cellStr = String(cellVal);
            const endTime = extractLastTime(cellStr);

            if (endTime) {
                const [rawOt, minutes] = calcOvertimeMinutes(endTime, overtimeStartMinutes);
                totalMinutes += minutes;
            }
        }

        const totalHours = totalMinutes / 60;

        let diff = null;
        let existingFloat = null;
        if (existingTotal !== null) {
            const parsed = parseFloat(existingTotal);
            if (!isNaN(parsed)) {
                existingFloat = parsed;
                diff = totalHours - parsed;
            }
        }

        results.push({
            name: name,
            dept: dept,
            totalMinutes: totalMinutes,
            totalHours: totalHours,
            existingTotal: existingTotal,
            existingFloat: existingFloat,
            diff: diff
        });
    }

    const diffItems = results.filter(r => r.diff !== null && Math.abs(r.diff) > 0.001);

    return {
        results: results,
        diffItems: diffItems,
        checkDateRange: checkDateRange
    };
}

function extractLastTime(cellStr) {
    if (!cellStr) return null;
    const str = String(cellStr).trim();
    const matches = str.match(/\((\d{1,2}:\d{2})\)/g);
    if (matches && matches.length > 0) {
        const last = matches[matches.length - 1];
        const timeMatch = last.match(/\((\d{1,2}:\d{2})\)/);
        return timeMatch ? timeMatch[1] : null;
    }
    return null;
}

function calcOvertimeMinutes(endTimeStr, overtimeStartMinutes) {
    if (!endTimeStr) return [0, 0];

    try {
        const [h, m] = endTimeStr.split(':').map(Number);

        let endMinutes;
        if (h < 7) {
            endMinutes = 24 * 60 + h * 60 + m;
        } else {
            endMinutes = h * 60 + m;
        }

        if (endMinutes <= overtimeStartMinutes) {
            return [0, 0];
        }

        const rawOvertime = endMinutes - overtimeStartMinutes;
        const units = Math.floor(rawOvertime / 30);
        const roundedMinutes = units > 0 ? units * 30 : 0;

        return [rawOvertime, roundedMinutes];
    } catch (e) {
        return [0, 0];
    }
}

function displayResults(allResults, filename) {
    document.getElementById('statsArea').innerHTML = '';
    document.getElementById('alertArea').innerHTML = '';

    const reportHTML = generateReportHTML(allResults, filename);
    document.getElementById('report-content').innerHTML = reportHTML;
    document.getElementById('report-container').style.display = 'block';
    setTimeout(() => {
        const reportSection = document.querySelector('.report-section');
        if (reportSection) {
            reportSection.scrollIntoView({ behavior: 'smooth', block: 'start' });
        } else {
            const reportContent = document.querySelector('.report-content');
            if (reportContent) {
                reportContent.scrollIntoView({ behavior: 'smooth', block: 'start' });
            }
        }
    }, 50);
}

function generateReportHTML(allResults, filename) {
    let html = `<div class="report-content">`;

    html += `
    <div class="report-nav">
        <div class="report-nav-links">
            <a href="#" onclick="jumpTo('conclusion'); return false;">核对结论</a>
            <a href="#" onclick="jumpTo('summary'); return false;">加班明细</a>
            <a href="#" onclick="jumpTo('top10'); return false;">TOP 10</a>
            <a href="#" onclick="jumpTo('rules'); return false;">核对规则</a>
        </div>
    </div>`;

    if (allResults.length > 1) {
        html += `<div class="report-section">
            <div class="report-section-title">工作表切换</div>
            <div class="sheet-tabs">`;
        
        allResults.forEach((item, index) => {
            html += `<div class="sheet-tab ${index === 0 ? 'active' : ''}" onclick="switchSheet(this)">${item.sheetName}</div>`;
        });
        
        html += `</div>`;
    }

    allResults.forEach((item, index) => {
        const { results, diffItems, checkDateRange } = item.data;
        const isFirst = index === 0;
        
        if (allResults.length > 1) {
            html += `<div class="sheet-content ${isFirst ? 'active' : ''}" id="sheet-content-${index}">`;
        }

        html += `
        <div class="report-section conclusion-section" id="${item.sheetName}-conclusion">
            <div class="report-section-title">${item.sheetName} - 核对结论</div>
            <div class="report-meta" style="justify-content: flex-start; gap: 16px; margin-bottom: 8px;">
                <div class="report-meta-item"><span>核对范围：</span>${new Date().getFullYear()}年${checkDateRange}</div>
                <div class="report-meta-item"><span>人数：</span>${results.length}人</div>
            </div>
            ${diffItems.length === 0
                ? `<div class="conclusion success">核对结果：全部${results.length}人加班时长均与表格记录一致，无差异项目。</div>`
                : `<div class="conclusion warning">核对结果：发现 ${diffItems.length} 处差异。</div>`
            }`;

        if (diffItems.length > 0) {
            html += `
            <div id="${item.sheetName}-diff">
                <table class="data-table">
                    <thead><tr><th>序号</th><th>姓名</th><th>部门</th><th>计算加班(小时)</th><th>表格已有(小时)</th><th>差异(小时)</th><th>差异说明</th></tr></thead>
                    <tbody>`;
            for (let i = 0; i < diffItems.length; i++) {
                const r = diffItems[i];
                html += `<tr class="diff-row"><td>${i+1}</td><td>${r.name}</td><td>${r.dept}</td><td>${r.totalHours.toFixed(2)}</td><td>${r.existingFloat.toFixed(2)}</td><td>${r.diff.toFixed(2)}</td><td>计算值${r.totalHours.toFixed(2)}h，表格值${r.existingFloat.toFixed(2)}h，差异${r.diff.toFixed(2)}h</td></tr>`;
            }
            html += `</tbody></table></div>`;
        }

        html += `
        </div>
        <div class="report-section summary-section" id="${item.sheetName}-summary">
            <div class="report-section-title">加班明细</div>
            <table class="data-table">
                <thead><tr><th>序号</th><th>姓名</th><th>部门</th><th>计算加班(小时)</th><th>表格已有(小时)</th><th>差异(小时)</th><th>核对结果</th></tr></thead>
                <tbody>`;

        for (let i = 0; i < results.length; i++) {
            const r = results[i];
            const isDiff = r.diff !== null && Math.abs(r.diff) > 0.001;
            const diffStr = r.diff !== null ? r.diff.toFixed(2) : 'N/A';
            const existingVal = r.existingFloat !== null ? r.existingFloat : (r.existingTotal !== null ? r.existingTotal : 'N/A');
            const rowClass = isDiff ? 'diff-row' : '';
            const statusClass = isDiff ? 'status-fail' : 'status-pass';
            const statusIcon = isDiff ? '❌' : '✅';
            html += `<tr class="${rowClass}"><td>${i+1}</td><td>${r.name}</td><td>${r.dept}</td><td>${r.totalHours.toFixed(2)}</td><td>${existingVal}</td><td>${diffStr}</td><td class="${statusClass}">${statusIcon}</td></tr>`;
        }

        html += `</tbody></table></div>`;

        const sortedResults = [...results].sort((a, b) => b.totalHours - a.totalHours);
        html += `
        <div class="report-section top10-section" id="${item.sheetName}-top10">
            <div class="report-section-title">加班时长<span class="top10-badge">TOP 10</span></div>
            <table class="data-table">
                <thead><tr><th>排名</th><th>姓名</th><th>部门</th><th>加班总时长(小时)</th></tr></thead>
                <tbody>`;
        for (let i = 0; i < Math.min(10, sortedResults.length); i++) {
            const r = sortedResults[i];
            html += `<tr><td>${i+1}</td><td>${r.name}</td><td>${r.dept}</td><td>${r.totalHours.toFixed(2)}</td></tr>`;
        }
        html += `</tbody></table></div>`;

        if (allResults.length > 1) {
            html += `</div>`;
        }
    });

    html += `
    <div class="report-section" id="rules">
        <div class="report-section-title">核对规则</div>
        <table class="rules-table">
            <thead><tr><th>规则项目</th><th>说明</th></tr></thead>
            <tbody>
                <tr><td>加班起算时间</td><td>19:00之后才算加班</td></tr>
                <tr><td>不足30分钟</td><td>不计入加班</td></tr>
                <tr><td>超过30分钟不足1小时</td><td>计算为30分钟</td></tr>
                <tr><td>超过1小时不足1.5小时</td><td>计算为1小时</td></tr>
                <tr><td>计算单位</td><td>以30分钟为单位向下取整</td></tr>
                <tr><td>跨午夜情况</td><td>凌晨时间（如00:07）视为当天加班延续</td></tr>
            </tbody>
        </table>
    </div>
    <div class="report-footer">
        <p>报告由自动化脚本生成，核对逻辑：提取每日最后一次打卡时间，计算19:00后加班时长，以30分钟为单位向下取整（不足30分钟不计）。</p>
    </div>
</div>`;

    return html;
}

let currentSheetIndex = 0;

function switchSheet(tab) {
    const tabs = tab.parentElement.querySelectorAll('.sheet-tab');
    tabs.forEach(t => t.classList.remove('active'));
    tab.classList.add('active');

    currentSheetIndex = Array.from(tabs).indexOf(tab);
    const contents = document.querySelectorAll('.sheet-content');
    contents.forEach(c => c.classList.remove('active'));
    contents[currentSheetIndex].classList.add('active');
}

function jumpTo(target) {
    if (target === 'rules') {
        const element = document.getElementById('rules');
        if (element) {
            element.scrollIntoView({ behavior: 'smooth' });
        }
        return;
    }
    
    const sheetTabs = document.querySelectorAll('.sheet-tab');
    const currentSheetName = sheetTabs.length > 0 ? sheetTabs[currentSheetIndex].textContent : '深圳';
    
    const element = document.getElementById(currentSheetName + '-' + target);
    if (element) {
        element.scrollIntoView({ behavior: 'smooth' });
    }
}

function showAlert(message, type) {
    const alertArea = document.getElementById('alertArea');
    const className = type === 'error' ? 'alert-error' : (type === 'warning' ? 'alert-warning' : 'alert-success');
    alertArea.innerHTML = `<div class="alert ${className}">${message}</div>`;
}