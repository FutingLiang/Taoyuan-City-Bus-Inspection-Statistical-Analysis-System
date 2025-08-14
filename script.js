let analysisData = null;
let charts = {};

// 註冊 DataLabels 插件
Chart.register(ChartDataLabels);

// 設定全域圖表預設值
Chart.defaults.font.family = "'Microsoft JhengHei', 'Segoe UI', Arial, sans-serif";
Chart.defaults.font.size = 12;
Chart.defaults.color = '#2c3e50';

// 增強的顏色調色盤
const colorPalettes = {
    primary: [
        'rgba(102, 126, 234, 0.8)',
        'rgba(118, 75, 162, 0.8)',
        'rgba(240, 147, 251, 0.8)',
        'rgba(245, 87, 108, 0.8)',
        'rgba(79, 172, 254, 0.8)',
        'rgba(0, 242, 254, 0.8)',
        'rgba(255, 107, 107, 0.8)',
        'rgba(78, 205, 196, 0.8)',
        'rgba(69, 183, 209, 0.8)',
        'rgba(150, 206, 180, 0.8)',
        'rgba(254, 202, 87, 0.8)',
        'rgba(255, 159, 243, 0.8)'
    ],
    borders: [
        'rgba(102, 126, 234, 1)',
        'rgba(118, 75, 162, 1)',
        'rgba(240, 147, 251, 1)',
        'rgba(245, 87, 108, 1)',
        'rgba(79, 172, 254, 1)',
        'rgba(0, 242, 254, 1)',
        'rgba(255, 107, 107, 1)',
        'rgba(78, 205, 196, 1)',
        'rgba(69, 183, 209, 1)',
        'rgba(150, 206, 180, 1)',
        'rgba(254, 202, 87, 1)',
        'rgba(255, 159, 243, 1)'
    ]
};

document.getElementById('fileInput').addEventListener('change', handleFileUpload);

function handleFileUpload(e) {
    const file = e.target.files[0];
    if (!file) return;
    
    const reader = new FileReader();
    const fileInfo = document.getElementById('fileInfo');
    fileInfo.innerHTML = `<i class="fas fa-spinner fa-spin"></i> 正在處理: ${file.name}`;
    
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { 
                type: 'array',
                cellDates: true,
                dateNF: 'yyyy/mm/dd'
            });
            
            const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
            const jsonData = XLSX.utils.sheet_to_json(firstSheet, {
                raw: false,
                dateNF: 'yyyy/mm/dd'
            });
            
            const realData = jsonData.filter(row => row['編號']);
            
            if (realData.length === 0) {
                alert('找不到有效的稽查資料！');
                return;
            }
            
            analyzeData(realData);
            
            document.getElementById('results').classList.remove('hidden');
            document.getElementById('results').classList.add('fade-in');
            
            fileInfo.innerHTML = `
                <i class="fas fa-check-circle" style="color: #00b894;"></i> 
                成功載入 ${file.name} - 共 ${realData.length} 筆稽查資料
            `;
            
        } catch (error) {
            alert('檔案處理錯誤: ' + error.message);
            console.error(error);
        }
    };
    
    reader.readAsArrayBuffer(file);
}

function analyzeData(data) {
    analysisData = data;
    
    const stats = {
        total: data.length,
        companies: {},
        areas: {},
        violations: {},
        safety: {
            normal: 0, warning: 0, breaker: 0, sticker: 0, font: 0,
            extinguisher: 0, pressure: 0, placement: 0, instruction: 0, seat: 0
        }
    };
    
    const violationFields = {
        equipment: {
            'A2 路線圖': '未張貼路線圖',
            'A3 意見箱': '無意見箱',
            'A4 電話': '無服務電話'
        },
        behavior: {
            'B4 吸菸': '駕駛吸菸',
            'C1 早發': '早發',
            'C2 誤點': '誤點',
            'C3 脫班': '脫班',
            'C4 漏班': '漏班',
            'C5 紀錄器': '行車紀錄器異常',
            'C6 過站': '過站不停',
            'C7 依路線': '未依路線行駛',
            'C8 拒載': '拒載',
            'C9 未停妥': '未停妥',
            'C10 關車門': '未等乘客坐穩關門',
            'C11 闖紅燈': '闖紅燈',
            'C12 佔用': '佔用車道',
            'C13 延滯': '延滯發車'
        }
    };
    
    data.forEach(row => {
        const company = row['公司'] || row['客運公司'] || '未知';
        stats.companies[company] = (stats.companies[company] || 0) + 1;
        
        const area = row['區域'] || '未知';
        stats.areas[area] = (stats.areas[area] || 0) + 1;
        
        for (let field in violationFields.equipment) {
            if (row[field] === '0' || row[field] === 0) {
                const fieldName = violationFields.equipment[field];
                stats.violations[fieldName] = (stats.violations[fieldName] || 0) + 1;
            }
        }
        
        for (let field in violationFields.behavior) {
            if (row[field] === '1' || row[field] === 1) {
                const fieldName = violationFields.behavior[field];
                stats.violations[fieldName] = (stats.violations[fieldName] || 0) + 1;
            }
        }
        
        if (row['D1 正常'] === '1' || row['D1 正常'] === 1) stats.safety.normal++;
        if (row['D2 警告'] === '1' || row['D2 警告'] === 1) stats.safety.warning++;
        if (row['D3 擊破器'] === '1' || row['D3 擊破器'] === 1) stats.safety.breaker++;
        if (row['D4 貼紙'] === '1' || row['D4 貼紙'] === 1) stats.safety.sticker++;
        if (row['D5 字體'] === '1' || row['D5 字體'] === 1) stats.safety.font++;
        if (row['E1 有效期'] === '1' || row['E1 有效期'] === 1) stats.safety.extinguisher++;
        if (row['E2 壓力針'] === '1' || row['E2 壓力針'] === 1) stats.safety.pressure++;
        if (row['E3 放置'] === '1' || row['E3 放置'] === 1) stats.safety.placement++;
        if (row['E4 說明'] === '1' || row['E4 說明'] === 1) stats.safety.instruction++;
        if (row['F1 座椅'] === '1' || row['F1 座椅'] === 1) stats.safety.seat++;
    });
    
    updateStatsCards(stats);
    updateCharts(stats);
    updateViolationTable(data, violationFields);
}

function updateStatsCards(stats) {
    const statsGrid = document.getElementById('statsGrid');
    
    const totalViolations = Object.values(stats.violations).reduce((a, b) => a + b, 0);
    const violationRate = ((totalViolations / (stats.total * 18)) * 100).toFixed(1);
    
    statsGrid.innerHTML = `
        <div class="stat-card fade-in">
            <h3><i class="fas fa-clipboard-list"></i> 總稽查次數</h3>
            <div class="stat-value">${stats.total}</div>
            <div class="stat-subtitle">筆資料</div>
            <div class="progress-bar">
                <div class="progress-fill" style="width: 100%"></div>
            </div>
        </div>
        <div class="stat-card fade-in">
            <h3><i class="fas fa-building"></i> 涵蓋客運公司</h3>
            <div class="stat-value">${Object.keys(stats.companies).length}</div>
            <div class="stat-subtitle">家業者</div>
            <div class="progress-bar">
                <div class="progress-fill" style="width: ${Math.min(Object.keys(stats.companies).length * 10, 100)}%"></div>
            </div>
        </div>
        <div class="stat-card fade-in">
            <h3><i class="fas fa-map-marked-alt"></i> 稽查區域</h3>
            <div class="stat-value">${Object.keys(stats.areas).length}</div>
            <div class="stat-subtitle">個行政區</div>
            <div class="progress-bar">
                <div class="progress-fill" style="width: ${Math.min(Object.keys(stats.areas).length * 8, 100)}%"></div>
            </div>
        </div>
        <div class="stat-card fade-in">
            <h3><i class="fas fa-exclamation-triangle"></i> 違規發現率</h3>
            <div class="stat-value">${violationRate}%</div>
            <div class="stat-subtitle">需改善項目</div>
            <div class="progress-bar">
                <div class="progress-fill" style="width: ${violationRate}%"></div>
            </div>
        </div>
    `;
}

function updateCharts(stats) {
    Object.values(charts).forEach(chart => chart.destroy());
    charts = {};
    
    // 客運公司圖表 - 增強版長條圖
    const companyCtx = document.getElementById('companyChart').getContext('2d');
    const companyData = Object.entries(stats.companies).sort((a, b) => b[1] - a[1]);
    
    charts.company = new Chart(companyCtx, {
        type: 'bar',
        data: {
            labels: companyData.map(item => item[0]),
            datasets: [{
                label: '稽查次數',
                data: companyData.map(item => item[1]),
                backgroundColor: colorPalettes.primary.slice(0, companyData.length),
                borderColor: colorPalettes.borders.slice(0, companyData.length),
                borderWidth: 2,
                borderRadius: 8,
                borderSkipped: false,
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            layout: {
                padding: { top: 40, right: 20, bottom: 10, left: 10 }
            },
            plugins: {
                legend: { display: false },
                tooltip: {
                    backgroundColor: 'rgba(0, 0, 0, 0.8)',
                    titleColor: '#fff',
                    bodyColor: '#fff',
                    borderColor: '#667eea',
                    borderWidth: 1,
                    cornerRadius: 8,
                    callbacks: {
                        label: function(context) {
                            const total = context.dataset.data.reduce((a, b) => a + b, 0);
                            const percentage = ((context.parsed.y / total) * 100).toFixed(1);
                            return `稽查次數: ${context.parsed.y} (${percentage}%)`;
                        }
                    }
                },
                datalabels: {
                    anchor: 'end',
                    align: 'top',
                    offset: 5,
                    color: '#2c3e50',
                    font: { weight: 'bold', size: 13 },
                    formatter: (value) => value,
                    clip: false
                }
            },
            scales: {
                y: {
                    beginAtZero: true,
                    ticks: { 
                        stepSize: 1,
                        color: '#7f8c8d',
                        font: { size: 11 }
                    },
                    grid: {
                        color: 'rgba(0, 0, 0, 0.05)',
                        drawBorder: false
                    },
                    max: function(context) {
                        const values = context.chart.data.datasets[0].data;
                        const max = Math.max(...values);
                        return Math.ceil(max * 1.3);
                    }
                },
                x: {
                    ticks: {
                        color: '#0f0d0d',
                        font: { size: 14, weight: 'bold' },
                        maxRotation: 45
                    },
                    grid: {
                        display: false
                    }
                }
            }
        }
    });
    
    // 區域分布圖表 - 增強版甜甜圈圖
    const areaCtx = document.getElementById('areaChart').getContext('2d');
    const areaValues = Object.values(stats.areas);
    const areaLabels = Object.keys(stats.areas);
    const total = areaValues.reduce((a, b) => a + b, 0);
    
    charts.area = new Chart(areaCtx, {
        type: 'doughnut',
        data: {
            labels: areaLabels,
            datasets: [{
                data: areaValues,
                backgroundColor: colorPalettes.primary.slice(0, areaLabels.length),
                borderColor: '#fff',
                borderWidth: 3,
                hoverBorderWidth: 5,
                hoverOffset: 10
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            cutout: '60%',
            plugins: {
                legend: {
                    position: 'right',
                    labels: {
                        generateLabels: function(chart) {
                            const data = chart.data;
                            return data.labels.map((label, i) => {
                                const percentage = ((data.datasets[0].data[i] / total) * 100).toFixed(1);
                                return {
                                    text: `${label} (${data.datasets[0].data[i]}次, ${percentage}%)`,
                                    fillStyle: data.datasets[0].backgroundColor[i],
                                    hidden: false,
                                    index: i,
                                    fontColor: '#2c3e50',
                                    fontSize: 11
                                };
                            });
                        },
                        padding: 15,
                        usePointStyle: true,
                        pointStyle: 'circle'
                    }
                },
                tooltip: {
                    backgroundColor: 'rgba(0, 0, 0, 0.8)',
                    titleColor: '#fff',
                    bodyColor: '#fff',
                    borderColor: '#667eea',
                    borderWidth: 1,
                    cornerRadius: 8,
                    callbacks: {
                        label: function(context) {
                            const label = context.label || '';
                            const value = context.parsed;
                            const percentage = ((value / total) * 100).toFixed(1);
                            return `${label}: ${value}次 (${percentage}%)`;
                        }
                    }
                },
                datalabels: {
                    color: '#fff',
                    font: { weight: 'bold', size: 11 },
                    formatter: (value, ctx) => {
                        const percentage = ((value / total) * 100).toFixed(0);
                        return percentage > 8 ? `${percentage}%` : '';
                    }
                }
            }
        }
    });
    
    // 違規項目圖表 - 增強版橫向長條圖
    const sortedViolations = Object.entries(stats.violations)
        .sort((a, b) => b[1] - a[1])
        .slice(0, 10);
    
    const violationCtx = document.getElementById('violationChart').getContext('2d');
    charts.violation = new Chart(violationCtx, {
        type: 'bar',
        data: {
            labels: sortedViolations.map(v => v[0]),
            datasets: [{
                label: '違規次數',
                data: sortedViolations.map(v => v[1]),
                backgroundColor: 'rgba(255, 99, 132, 0.8)',
                borderColor: 'rgba(255, 99, 132, 1)',
                borderWidth: 2,
                borderRadius: 6,
                borderSkipped: false
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            indexAxis: 'y',
            layout: {
                padding: { top: 10, right: 60, bottom: 10, left: 10 }
            },
            plugins: {
                legend: { display: false },
                tooltip: {
                    backgroundColor: 'rgba(0, 0, 0, 0.8)',
                    titleColor: '#fff',
                    bodyColor: '#fff',
                    borderColor: '#ff6b6b',
                    borderWidth: 1,
                    cornerRadius: 8
                },
                datalabels: {
                    anchor: 'end',
                    align: 'end',
                    offset: 8,
                    color: '#2c3e50',
                    font: { weight: 'bold', size: 14 },
                    formatter: (value) => value,
                    clip: false
                }
            },
            scales: {
                x: {
                    beginAtZero: true,
                    ticks: { 
                        stepSize: 1,
                        color: '#7f8c8d',
                        font: { size: 13, weight: 'bold' }
                    },
                    grid: {
                        color: 'rgba(0, 0, 0, 0.05)',
                        drawBorder: false
                    },
                    max: function(context) {
                        const values = context.chart.data.datasets[0].data;
                        const max = Math.max(...values);
                        return Math.ceil(max * 1.4);
                    }
                },
                y: {
                    ticks: {
                        color: '#0f0d0d',
                        font: { size: 13, weight: 'bold' }
                    },
                    grid: { display: false }
                }
            }
        }
    });
    
    // 安全設備圖表 - 增強版雷達圖
    const safetyCtx = document.getElementById('safetyChart').getContext('2d');
    const safetyLabels = [
        '逃生門正常', '警告標示', '擊破器', '安全貼紙', '字體清晰',
        '滅火器有效', '壓力正常', '放置適當', '使用說明', '座椅正常'
    ];
    const safetyData = [
        stats.safety.normal, stats.safety.warning, stats.safety.breaker,
        stats.safety.sticker, stats.safety.font, stats.safety.extinguisher,
        stats.safety.pressure, stats.safety.placement, stats.safety.instruction, stats.safety.seat
    ];
    
    charts.safety = new Chart(safetyCtx, {
        type: 'radar',
        data: {
            labels: safetyLabels,
            datasets: [{
                label: '符合數量',
                data: safetyData,
                backgroundColor: 'rgba(75, 192, 192, 0.2)',
                borderColor: 'rgba(75, 192, 192, 1)',
                borderWidth: 3,
                pointBackgroundColor: 'rgba(75, 192, 192, 1)',
                pointBorderColor: '#fff',
                pointBorderWidth: 2,
                pointHoverBackgroundColor: '#fff',
                pointHoverBorderColor: 'rgba(75, 192, 192, 1)',
                pointRadius: 6,
                pointHoverRadius: 8
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            layout: { padding: 50 },
            plugins: {
                legend: {
                    labels: {
                        color: '#2c3e50',
                        font: { size: 14, weight: 'bold' }
                    }
                },
                tooltip: {
                    backgroundColor: 'rgba(0, 0, 0, 0.8)',
                    titleColor: '#fff',
                    bodyColor: '#fff',
                    borderColor: '#4bc0c0',
                    borderWidth: 1,
                    cornerRadius: 8,
                    callbacks: {
                        label: function(context) {
                            const label = context.dataset.label || '';
                            const value = context.parsed.r;
                            const percentage = ((value / stats.total) * 100).toFixed(1);
                            return `${label}: ${value}/${stats.total} (${percentage}%)`;
                        }
                    }
                },
                datalabels: {
                    backgroundColor: 'rgba(255, 255, 255, 0.9)',
                    borderColor: 'rgba(75, 192, 192, 1)',
                    borderRadius: 4,
                    borderWidth: 1,
                    color: '#2c3e50',
                    font: { weight: 'bold', size: 11 },
                    formatter: (value) => value,
                    padding: 3,
                    clip: false,
                    anchor: 'center',
                    align: 'center',
                    offset: 0,
                    display: function(context) {
                        // 安全檢查並只在數值大於0時顯示
                        try {
                            return context && context.parsed && typeof context.parsed.r === 'number' && context.parsed.r > 0;
                        } catch (e) {
                            return true; // 如果出錯，預設顯示標籤
                        }
                    }
                }
            },
            scales: {
                r: {
                    beginAtZero: true,
                    max: Math.ceil(stats.total * 1.1),
                    ticks: {
                        stepSize: Math.ceil(stats.total / 5),
                        color: '#7f8c8d',
                        font: { size: 12 },
                        backdropColor: 'rgba(255, 255, 255, 0.8)',
                        backdropPadding: 3
                    },
                    grid: {
                        color: 'rgba(0, 0, 0, 0.1)'
                    },
                    angleLines: {
                        color: 'rgba(0, 0, 0, 0.1)'
                    },
                    pointLabels: {
                        color: '#2c3e50',
                        font: { size: 13, weight: '600' },
                        padding: 10,
                        centerPointLabels: false
                    }
                }
            }
        }
    });
}

function updateViolationTable(data, violationFields) {
    const tbody = document.getElementById('violationTableBody');
    tbody.innerHTML = '';
    
    data.forEach(row => {
        const violations = [];
        
        for (let field in violationFields.equipment) {
            if (row[field] === '0' || row[field] === 0) {
                violations.push(violationFields.equipment[field]);
            }
        }
        
        for (let field in violationFields.behavior) {
            if (row[field] === '1' || row[field] === 1) {
                violations.push(violationFields.behavior[field]);
            }
        }
        
        if (violations.length > 0) {
            const tr = document.createElement('tr');
            let dateStr = '-';
            if (row['日期']) {
                let date;
                if (typeof row['日期'] === 'object' && row['日期'] instanceof Date) {
                    date = row['日期'];
                } else if (typeof row['日期'] === 'string') {
                    if (row['日期'].indexOf('/') > 0 && row['日期'].split('/').length === 2) {
                        date = new Date(`2024/${row['日期']}`);
                    } else {
                        date = new Date(row['日期']);
                    }
                } else if (typeof row['日期'] === 'number') {
                    date = new Date((row['日期'] - 25569) * 86400 * 1000);
                }
                
                if (date && !isNaN(date.getTime()) && date.getFullYear() >= 2000) {
                    dateStr = date.toLocaleDateString('zh-TW');
                }
            }
            
            tr.innerHTML = `
                <td>${row['編號'] || '-'}</td>
                <td>${dateStr}</td>
                <td>${row['路線'] || '-'}</td>
                <td>${row['公司'] || row['客運公司'] || '-'}</td>
                <td>${row['車牌'] || '-'}</td>
                <td>${row['駕駛'] || '-'}</td>
                <td>${violations.map(v => `<span class="violation">${v}</span>`).join(' ')}</td>
            `;
            tbody.appendChild(tr);
        }
    });
    
    if (tbody.children.length === 0) {
        tbody.innerHTML = '<tr><td colspan="7" style="text-align: center; color: #7f8c8d; font-style: italic;">無違規記錄</td></tr>';
    }
}

function exportReport() {
    if (!analysisData) return;
    
    const exportData = [];
    const header = ['編號', '日期', '路線', '客運公司', '車牌', '駕駛', '違規項目數', '主要違規'];
    
    analysisData.forEach(row => {
        const violations = [];
        const violationFields = {
            equipment: {
                'A2 路線圖': '未張貼路線圖',
                'A3 意見箱': '無意見箱',
                'A4 電話': '無服務電話'
            },
            behavior: {
                'B4 吸菸': '駕駛吸菸',
                'C1 早發': '早發',
                'C2 誤點': '誤點',
                'C3 脫班': '脫班',
                'C4 漏班': '漏班',
                'C5 紀錄器': '行車紀錄器異常',
                'C6 過站': '過站不停',
                'C7 依路線': '未依路線行駛',
                'C8 拒載': '拒載',
                'C9 未停妥': '未停妥',
                'C10 關車門': '未等乘客坐穩關門',
                'C11 闖紅燈': '闖紅燈',
                'C12 佔用': '佔用車道',
                'C13 延滯': '延滯發車'
            }
        };
        
        for (let field in violationFields.equipment) {
            if (row[field] === '0' || row[field] === 0) {
                violations.push(violationFields.equipment[field]);
            }
        }
        
        for (let field in violationFields.behavior) {
            if (row[field] === '1' || row[field] === 1) {
                violations.push(violationFields.behavior[field]);
            }
        }
        
        exportData.push([
            row['編號'] || '',
            row['日期'] ? (new Date(row['日期']).getFullYear() > 1970 ? new Date(row['日期']).toLocaleDateString('zh-TW') : '') : '',
            row['路線'] || '',
            row['公司'] || row['客運公司'] || '',
            row['車牌'] || '',
            row['駕駛'] || '',
            violations.length,
            violations.join(', ')
        ]);
    });
    
    const wb = XLSX.utils.book_new();
    
    const summaryData = [
        ['桃園公車稽查統計報告'],
        [''],
        ['報告日期', new Date().toLocaleDateString('zh-TW')],
        ['總稽查次數', analysisData.length],
        [''],
        ['客運公司統計'],
    ];
    
    const companies = {};
    analysisData.forEach(row => {
        const company = row['客運公司'] || '未知';
        companies[company] = (companies[company] || 0) + 1;
    });
    
    for (let company in companies) {
        summaryData.push([company, companies[company]]);
    }
    
    const ws1 = XLSX.utils.aoa_to_sheet(summaryData);
    const ws2 = XLSX.utils.aoa_to_sheet([header, ...exportData]);
    
    XLSX.utils.book_append_sheet(wb, ws1, '統計摘要');
    XLSX.utils.book_append_sheet(wb, ws2, '違規明細');
    
    const fileName = `稽查統計報告_${new Date().toLocaleDateString('zh-TW').replace(/\//g, '')}.xlsx`;
    XLSX.writeFile(wb, fileName);
    
    alert(`報告已匯出: ${fileName}`);
}
