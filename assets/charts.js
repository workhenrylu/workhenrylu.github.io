// assets/charts.js - Salary Dashboard
(function() {
  'use strict';

  // === Color Palette ===
  var C = {
    accent: '#00e5ff', accent2: '#ff6e40', accent3: '#7c4dff',
    accent4: '#69f0ae', accent5: '#ffd740', accent6: '#ff4081',
    ink: '#e8eaf6', muted: '#7986cb', rule: '#2a2f5e',
    bg2: '#111638', bg3: '#1a1f4e'
  };
  var palette = [C.accent, C.accent3, C.accent2, C.accent4, C.accent5, C.accent6];

  // === Global State ===
  var rawData = [];
  var charts = {};

  // === Utility Functions ===
  function parseMoney(v) {
    if (v === null || v === undefined) return 0;
    if (typeof v === 'number') return v;
    return parseFloat(String(v).replace(/[¥,\s]/g, '')) || 0;
  }

  function fmt(n) {
    return n.toLocaleString('zh-CN', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  }

  function fmtW(n) {
    if (n >= 10000) return (n / 10000).toFixed(2) + '万';
    return fmt(n);
  }

  function sum(arr, key) {
    return arr.reduce(function(s, r) { return s + parseMoney(r[key]); }, 0);
  }

  function avg(arr, key) {
    var vals = arr.filter(function(r) { return parseMoney(r[key]) > 0; });
    return vals.length ? sum(vals, key) / vals.length : 0;
  }

  function groupBy(data, key) {
    var m = {};
    data.forEach(function(r) {
      var k = r[key] || '未知';
      if (!m[k]) m[k] = [];
      m[k].push(r);
    });
    return m;
  }

  // === Data Loading ===
  function loadDefaultData() {
    var xhr = new XMLHttpRequest();
    xhr.open('GET', 'assets/default-data.json', true);
    xhr.onload = function() {
      if (xhr.status === 200) {
        rawData = JSON.parse(xhr.responseText);
        initDashboard();
      }
    };
    xhr.send();
  }

  function parseExcelData(workbook) {
    var sheet = workbook.Sheets[workbook.SheetNames[0]];
    var json = XLSX.utils.sheet_to_json(sheet);
    var data = json.map(function(row) {
      var obj = {};
      Object.keys(row).forEach(function(k) {
        obj[k] = row[k];
      });
      return obj;
    });
    return data;
  }

  // === KPI Cards ===
  function renderKPIs() {
    var totalIncome = sum(rawData, '广义工资');
    var totalNet = sum(rawData, '实发金额');
    var totalDeduction = sum(rawData, '社保小计') + sum(rawData, '个人所得税');
    var avgMonthly = rawData.length ? totalNet / rawData.length : 0;
    var maxMonth = rawData.reduce(function(m, r) {
      var v = parseMoney(r['实发金额']);
      return v > m.v ? { v: v, label: r['年月'] } : m;
    }, { v: 0, label: '-' });
    var totalTax = sum(rawData, '个人所得税');

    // Year-over-year growth
    var years = Object.keys(groupBy(rawData, '年度')).sort();
    var yoyGrowth = 0;
    if (years.length >= 2) {
      var lastYear = years[years.length - 1];
      var prevYear = years[years.length - 2];
      var lastTotal = sum(rawData.filter(function(r) { return r['年度'] === lastYear; }), '实发金额');
      var prevTotal = sum(rawData.filter(function(r) { return r['年度'] === prevYear; }), '实发金额');
      yoyGrowth = prevTotal ? ((lastTotal - prevTotal) / prevTotal * 100) : 0;
    }

    var kpis = [
      { label: '累计广义收入', value: fmtW(totalIncome), unit: '元', trend: yoyGrowth >= 0 ? '+' + yoyGrowth.toFixed(1) + '% YoY' : yoyGrowth.toFixed(1) + '% YoY', up: yoyGrowth >= 0 },
      { label: '累计实发金额', value: fmtW(totalNet), unit: '元', trend: null, up: true },
      { label: '月均实发工资', value: fmtW(avgMonthly), unit: '元', trend: null, up: true },
      { label: '累计扣除总额', value: fmtW(totalDeduction), unit: '元', trend: null, up: false },
      { label: '最高月实发', value: fmtW(maxMonth.v), unit: maxMonth.label, trend: null, up: true },
      { label: '累计个人所得税', value: fmtW(totalTax), unit: '元', trend: null, up: false }
    ];

    var html = '';
    kpis.forEach(function(k) {
      html += '<div class="kpi-card">' +
        '<div class="kpi-label">' + k.label + '</div>' +
        '<div class="kpi-value">' + k.value + '<span class="kpi-unit">' + k.unit + '</span></div>' +
        (k.trend ? '<div class="kpi-trend ' + (k.up ? 'up' : 'down') + '">' + (k.up ? '&#9650; ' : '&#9660; ') + k.trend + '</div>' : '') +
        '</div>';
    });
    document.getElementById('kpi-row').innerHTML = html;
  }

  // === Chart 1: Monthly Trend ===
  function renderTrend() {
    var sorted = rawData.slice().sort(function(a, b) { return a['年月'] < b['年月'] ? -1 : 1; });
    var xData = sorted.map(function(r) { return r['年月']; });
    var grossData = sorted.map(function(r) { return parseMoney(r['应发工资']); });
    var netData = sorted.map(function(r) { return parseMoney(r['实发金额']); });
    var deductionData = sorted.map(function(r) { return parseMoney(r['社保小计']) + parseMoney(r['个人所得税']); });

    var option = {
      tooltip: { trigger: 'axis', appendToBody: true, backgroundColor: 'rgba(17,22,56,0.95)', borderColor: C.rule, textStyle: { color: C.ink, fontSize: 12 },
        formatter: function(p) {
          var s = '<b>' + p[0].axisValue + '</b><br/>';
          p.forEach(function(item) { s += item.marker + item.seriesName + ': ¥' + fmt(item.value) + '<br/>'; });
          return s;
        }
      },
      legend: { data: ['应发工资', '实发金额', '扣除合计'], textStyle: { color: C.muted, fontSize: 11 }, top: 0, right: 10 },
      grid: { left: 60, right: 20, top: 40, bottom: 30 },
      xAxis: { type: 'category', data: xData, axisLine: { lineStyle: { color: C.rule } }, axisLabel: { color: C.muted, fontSize: 10, rotate: 45 } },
      yAxis: { type: 'value', axisLine: { show: false }, splitLine: { lineStyle: { color: C.rule, type: 'dashed' } }, axisLabel: { color: C.muted, fontSize: 10, formatter: function(v) { return v >= 10000 ? (v/10000)+'万' : v; } } },
      series: [
        { name: '应发工资', type: 'line', data: grossData, smooth: true, symbol: 'circle', symbolSize: 4,
          lineStyle: { color: C.accent, width: 2 },
          itemStyle: { color: C.accent },
          areaStyle: { color: { type: 'linear', x: 0, y: 0, x2: 0, y2: 1, colorStops: [{ offset: 0, color: 'rgba(0,229,255,0.25)' }, { offset: 1, color: 'rgba(0,229,255,0)' }] } }
        },
        { name: '实发金额', type: 'line', data: netData, smooth: true, symbol: 'circle', symbolSize: 4,
          lineStyle: { color: C.accent4, width: 2 },
          itemStyle: { color: C.accent4 },
          areaStyle: { color: { type: 'linear', x: 0, y: 0, x2: 0, y2: 1, colorStops: [{ offset: 0, color: 'rgba(105,240,174,0.2)' }, { offset: 1, color: 'rgba(105,240,174,0)' }] } }
        },
        { name: '扣除合计', type: 'bar', data: deductionData, barWidth: 8,
          itemStyle: { color: { type: 'linear', x: 0, y: 0, x2: 0, y2: 1, colorStops: [{ offset: 0, color: 'rgba(255,110,64,0.7)' }, { offset: 1, color: 'rgba(255,110,64,0.1)' }] }, borderRadius: [3, 3, 0, 0] }
        }
      ],
      animation: true, animationDuration: 1500, animationEasing: 'cubicOut'
    };
    charts.trend = echarts.init(document.getElementById('chart-trend'));
    charts.trend.setOption(option);
  }

  // === Chart 2: Yearly Income Composition (Stacked Bar) ===
  function renderIncomeCompose() {
    var years = Object.keys(groupBy(rawData, '年度')).sort();
    var categories = ['基本工资', '月度效益', '综合福利', '预发效益奖', '各类补贴'];
    var series = categories.map(function(cat, i) {
      return {
        name: cat, type: 'bar', stack: 'total', barWidth: 24,
        itemStyle: { color: palette[i], borderRadius: i === categories.length - 1 ? [4, 4, 0, 0] : 0 },
        data: years.map(function(y) {
          var rows = rawData.filter(function(r) { return r['年度'] === y; });
          if (cat === '各类补贴') {
            return sum(rows, '车改补贴') + sum(rows, '防暑降温') + sum(rows, '住房补贴') + sum(rows, '通讯补贴') + sum(rows, '生活补贴') + sum(rows, '野外补贴');
          }
          return sum(rows, cat);
        })
      };
    });

    var option = {
      tooltip: { trigger: 'axis', appendToBody: true, backgroundColor: 'rgba(17,22,56,0.95)', borderColor: C.rule, textStyle: { color: C.ink, fontSize: 12 },
        formatter: function(p) {
          var s = '<b>' + p[0].axisValue + '</b><br/>';
          var total = 0;
          p.forEach(function(item) { s += item.marker + item.seriesName + ': ¥' + fmt(item.value) + '<br/>'; total += item.value; });
          s += '<b>合计: ¥' + fmt(total) + '</b>';
          return s;
        }
      },
      legend: { data: categories, textStyle: { color: C.muted, fontSize: 10 }, top: 0, type: 'scroll' },
      grid: { left: 55, right: 15, top: 40, bottom: 25 },
      xAxis: { type: 'category', data: years, axisLine: { lineStyle: { color: C.rule } }, axisLabel: { color: C.muted, fontSize: 11 } },
      yAxis: { type: 'value', axisLine: { show: false }, splitLine: { lineStyle: { color: C.rule, type: 'dashed' } }, axisLabel: { color: C.muted, fontSize: 10, formatter: function(v) { return v >= 10000 ? (v/10000)+'万' : v; } } },
      series: series,
      animation: true, animationDuration: 1200
    };
    charts.incomeCompose = echarts.init(document.getElementById('chart-income-compose'));
    charts.incomeCompose.setOption(option);
  }

  // === Chart 3: Latest Month Pie ===
  function renderMonthlyPie() {
    var sorted = rawData.slice().sort(function(a, b) { return a['年月'] > b['年月'] ? -1 : 1; });
    var latest = sorted[0];
    if (!latest) return;

    var items = [
      { name: '基本工资', value: parseMoney(latest['基本工资']) },
      { name: '月度效益', value: parseMoney(latest['月度效益']) },
      { name: '综合福利', value: parseMoney(latest['综合福利']) },
      { name: '预发效益奖', value: parseMoney(latest['预发效益奖']) },
      { name: '各类补贴', value: parseMoney(latest['车改补贴']) + parseMoney(latest['防暑降温']) + parseMoney(latest['住房补贴']) + parseMoney(latest['通讯补贴']) + parseMoney(latest['生活补贴']) + parseMoney(latest['野外补贴']) }
    ].filter(function(d) { return d.value > 0; });

    var option = {
      tooltip: { trigger: 'item', appendToBody: true, backgroundColor: 'rgba(17,22,56,0.95)', borderColor: C.rule, textStyle: { color: C.ink },
        formatter: function(p) { return p.marker + p.name + '<br/>¥' + fmt(p.value) + ' (' + p.percent + '%)'; }
      },
      legend: { orient: 'vertical', right: 10, top: 'center', textStyle: { color: C.muted, fontSize: 10 } },
      series: [{
        type: 'pie', radius: ['40%', '70%'], center: ['40%', '50%'],
        label: { show: false },
        emphasis: { label: { show: true, fontSize: 13, fontWeight: 'bold', color: C.ink },
          itemStyle: { shadowBlur: 20, shadowColor: 'rgba(0,229,255,0.4)' } },
        itemStyle: { borderColor: C.bg2, borderWidth: 2, borderRadius: 6 },
        data: items,
        color: palette
      }],
      graphic: [{
        type: 'text', left: '36%', top: '46%',
        style: { text: latest['年月'] || '', fill: C.muted, fontSize: 12, textAlign: 'center' }
      }],
      animation: true, animationDuration: 1200
    };
    charts.monthlyPie = echarts.init(document.getElementById('chart-monthly-pie'));
    charts.monthlyPie.setOption(option);
  }

  // === Chart 4: Yearly Income vs Deduction Bar ===
  function renderYearlyBar() {
    var years = Object.keys(groupBy(rawData, '年度')).sort();
    var incomeData = years.map(function(y) { return sum(rawData.filter(function(r) { return r['年度'] === y; }), '应发工资'); });
    var deductionData = years.map(function(y) {
      var rows = rawData.filter(function(r) { return r['年度'] === y; });
      return sum(rows, '社保小计') + sum(rows, '个人所得税') + sum(rows, '企业年金');
    });
    var netData = years.map(function(y) { return sum(rawData.filter(function(r) { return r['年度'] === y; }), '实发金额'); });

    var option = {
      tooltip: { trigger: 'axis', appendToBody: true, backgroundColor: 'rgba(17,22,56,0.95)', borderColor: C.rule, textStyle: { color: C.ink, fontSize: 12 },
        formatter: function(p) {
          var s = '<b>' + p[0].axisValue + '</b><br/>';
          p.forEach(function(item) { s += item.marker + item.seriesName + ': ¥' + fmt(item.value) + '<br/>'; });
          return s;
        }
      },
      legend: { data: ['应发工资', '扣除总额', '实发金额'], textStyle: { color: C.muted, fontSize: 11 }, top: 0 },
      grid: { left: 55, right: 15, top: 40, bottom: 25 },
      xAxis: { type: 'category', data: years, axisLine: { lineStyle: { color: C.rule } }, axisLabel: { color: C.muted, fontSize: 11 } },
      yAxis: { type: 'value', axisLine: { show: false }, splitLine: { lineStyle: { color: C.rule, type: 'dashed' } }, axisLabel: { color: C.muted, fontSize: 10, formatter: function(v) { return v >= 10000 ? (v/10000)+'万' : v; } } },
      series: [
        { name: '应发工资', type: 'bar', data: incomeData, barWidth: 20,
          itemStyle: { color: { type: 'linear', x: 0, y: 0, x2: 0, y2: 1, colorStops: [{ offset: 0, color: C.accent }, { offset: 1, color: C.accent3 }] }, borderRadius: [4, 4, 0, 0] }
        },
        { name: '扣除总额', type: 'bar', data: deductionData, barWidth: 20,
          itemStyle: { color: { type: 'linear', x: 0, y: 0, x2: 0, y2: 1, colorStops: [{ offset: 0, color: C.accent2 }, { offset: 1, color: C.accent6 }] }, borderRadius: [4, 4, 0, 0] }
        },
        { name: '实发金额', type: 'line', data: netData, smooth: true, symbol: 'diamond', symbolSize: 10,
          lineStyle: { color: C.accent4, width: 3 },
          itemStyle: { color: C.accent4, borderColor: C.bg2, borderWidth: 2 }
        }
      ],
      animation: true, animationDuration: 1200
    };
    charts.yearlyBar = echarts.init(document.getElementById('chart-yearly-bar'));
    charts.yearlyBar.setOption(option);
  }

  // === Chart 5: Radar - Subsidy Analysis ===
  function renderRadar() {
    var years = Object.keys(groupBy(rawData, '年度')).sort();
    var indicators = [
      { name: '车改补贴', max: 500 },
      { name: '住房补贴', max: 1500 },
      { name: '通讯补贴', max: 300 },
      { name: '生活补贴', max: 500 },
      { name: '野外补贴', max: 3000 },
      { name: '防暑降温', max: 800 }
    ];
    var keys = indicators.map(function(i) { return i.name; });

    var series = years.slice(-3).map(function(y, idx) {
      var rows = rawData.filter(function(r) { return r['年度'] === y; });
      var vals = keys.map(function(k) { return avg(rows, k); });
      return {
        value: vals,
        name: y,
        areaStyle: { opacity: 0.15 },
        lineStyle: { width: 2 }
      };
    });

    var option = {
      tooltip: { appendToBody: true, backgroundColor: 'rgba(17,22,56,0.95)', borderColor: C.rule, textStyle: { color: C.ink, fontSize: 12 } },
      legend: { data: years.slice(-3), textStyle: { color: C.muted, fontSize: 11 }, top: 0 },
      radar: {
        indicator: indicators,
        shape: 'polygon',
        splitNumber: 4,
        axisName: { color: C.muted, fontSize: 10 },
        splitLine: { lineStyle: { color: C.rule } },
        splitArea: { show: true, areaStyle: { color: ['rgba(0,229,255,0.02)', 'rgba(124,77,255,0.02)'] } },
        axisLine: { lineStyle: { color: C.rule } }
      },
      series: [{
        type: 'radar',
        data: series,
        color: palette
      }],
      animation: true, animationDuration: 1200
    };
    charts.radar = echarts.init(document.getElementById('chart-radar'));
    charts.radar.setOption(option);
  }

  // === Chart 6: Deduction Stacked Bar ===
  function renderDeduction() {
    var years = Object.keys(groupBy(rawData, '年度')).sort();
    var cats = ['养老保险', '医疗保险', '失业保险', '住房公积金', '企业年金', '个人所得税'];
    var series = cats.map(function(cat, i) {
      return {
        name: cat, type: 'bar', stack: 'deduction', barWidth: 20,
        itemStyle: { color: palette[i], borderRadius: i === cats.length - 1 ? [4, 4, 0, 0] : 0 },
        data: years.map(function(y) { return sum(rawData.filter(function(r) { return r['年度'] === y; }), cat); })
      };
    });

    var option = {
      tooltip: { trigger: 'axis', appendToBody: true, backgroundColor: 'rgba(17,22,56,0.95)', borderColor: C.rule, textStyle: { color: C.ink, fontSize: 12 },
        formatter: function(p) {
          var s = '<b>' + p[0].axisValue + '</b><br/>';
          var total = 0;
          p.forEach(function(item) { s += item.marker + item.seriesName + ': ¥' + fmt(item.value) + '<br/>'; total += item.value; });
          s += '<b>合计: ¥' + fmt(total) + '</b>';
          return s;
        }
      },
      legend: { data: cats, textStyle: { color: C.muted, fontSize: 10 }, top: 0, type: 'scroll' },
      grid: { left: 55, right: 15, top: 40, bottom: 25 },
      xAxis: { type: 'category', data: years, axisLine: { lineStyle: { color: C.rule } }, axisLabel: { color: C.muted, fontSize: 11 } },
      yAxis: { type: 'value', axisLine: { show: false }, splitLine: { lineStyle: { color: C.rule, type: 'dashed' } }, axisLabel: { color: C.muted, fontSize: 10, formatter: function(v) { return v >= 10000 ? (v/10000)+'万' : v; } } },
      series: series,
      animation: true, animationDuration: 1200
    };
    charts.deduction = echarts.init(document.getElementById('chart-deduction'));
    charts.deduction.setOption(option);
  }

  // === Chart 7: Heatmap ===
  function renderHeatmap() {
    var years = Object.keys(groupBy(rawData, '年度')).sort();
    var months = ['1月','2月','3月','4月','5月','6月','7月','8月','9月','10月','11月','12月'];
    var heatData = [];
    var maxVal = 0;

    rawData.forEach(function(r) {
      var yi = years.indexOf(r['年度']);
      var mi = -1;
      var ym = String(r['年月']);
      var m = ym.split('-');
      if (m.length >= 2) {
        mi = parseInt(m[1]) - 1;
      }
      if (yi >= 0 && mi >= 0) {
        var val = parseMoney(r['实发金额']);
        heatData.push([mi, yi, val]);
        if (val > maxVal) maxVal = val;
      }
    });

    var option = {
      tooltip: { appendToBody: true, backgroundColor: 'rgba(17,22,56,0.95)', borderColor: C.rule, textStyle: { color: C.ink, fontSize: 12 },
        formatter: function(p) { return p.data[1] !== undefined ? '<b>' + years[p.data[1]] + ' ' + months[p.data[0]] + '</b><br/>实发: ¥' + fmt(p.data[2]) : 'N/A'; }
      },
      grid: { left: 60, right: 40, top: 10, bottom: 30 },
      xAxis: { type: 'category', data: months, splitArea: { show: false }, axisLine: { lineStyle: { color: C.rule } }, axisLabel: { color: C.muted, fontSize: 11 } },
      yAxis: { type: 'category', data: years, splitArea: { show: false }, axisLine: { lineStyle: { color: C.rule } }, axisLabel: { color: C.muted, fontSize: 11 } },
      visualMap: {
        min: 0, max: maxVal,
        calculable: true, orient: 'horizontal', left: 'center', bottom: 0,
        inRange: { color: ['#0d1137', '#1a237e', '#00bcd4', '#00e5ff', '#69f0ae', '#ffd740'] },
        textStyle: { color: C.muted, fontSize: 10 },
        formatter: function(v) { return v >= 10000 ? (v/10000).toFixed(1) + '万' : Math.round(v); }
      },
      series: [{
        type: 'heatmap', data: heatData,
        label: { show: true, fontSize: 9, color: C.ink, formatter: function(p) { return p.value[2] > 0 ? (p.value[2] >= 10000 ? (p.value[2]/10000).toFixed(1)+'万' : Math.round(p.value[2])) : ''; } },
        emphasis: { itemStyle: { shadowBlur: 10, shadowColor: 'rgba(0,229,255,0.5)' } },
        itemStyle: { borderColor: C.bg, borderWidth: 3, borderRadius: 4 }
      }],
      animation: true, animationDuration: 1200
    };
    charts.heatmap = echarts.init(document.getElementById('chart-heatmap'));
    charts.heatmap.setOption(option);
  }

  // === Data Table ===
  function renderTable() {
    var cols = ['年月', '基本工资', '月度效益', '综合福利', '预发效益奖', '应发工资', '社保小计', '个人所得税', '实发金额', '广义工资'];
    var sorted = rawData.slice().sort(function(a, b) { return a['年月'] > b['年月'] ? -1 : 1; });

    var html = '<table class="data-table"><thead><tr>';
    cols.forEach(function(c) { html += '<th>' + c + '</th>'; });
    html += '</tr></thead><tbody>';
    sorted.forEach(function(r) {
      html += '<tr>';
      cols.forEach(function(c) {
        var v = r[c];
        if (c === '年月') {
          html += '<td>' + v + '</td>';
        } else {
          var n = parseMoney(v);
          html += '<td>' + (n > 0 ? fmt(n) : '-') + '</td>';
        }
      });
      html += '</tr>';
    });
    html += '</tbody></table>';
    document.getElementById('data-table-wrap').innerHTML = html;
  }

  // === Initialize All Charts ===
  function initDashboard() {
    renderKPIs();
    renderTrend();
    renderIncomeCompose();
    renderMonthlyPie();
    renderYearlyBar();
    renderRadar();
    renderDeduction();
    renderHeatmap();
    renderTable();
  }

  // === Resize Handler ===
  window.addEventListener('resize', function() {
    Object.keys(charts).forEach(function(k) { charts[k] && charts[k].resize(); });
  });

  // === Upload Functions ===
  window.openUpload = function() {
    document.getElementById('upload-overlay').classList.add('active');
    document.getElementById('upload-success').style.display = 'none';
  };

  window.closeUpload = function() {
    document.getElementById('upload-overlay').classList.remove('active');
  };

  window.handleFile = function(file) {
    if (!file) return;
    var reader = new FileReader();
    reader.onload = function(e) {
      try {
        var data = new Uint8Array(e.target.result);
        var workbook = XLSX.read(data, { type: 'array' });
        var parsed = parseExcelData(workbook);
        if (parsed.length === 0) {
          alert('文件中没有找到有效数据，请检查文件格式。');
          return;
        }
        rawData = parsed;
        // Dispose old charts
        Object.keys(charts).forEach(function(k) { charts[k] && charts[k].dispose(); });
        charts = {};
        initDashboard();
        document.getElementById('upload-success').style.display = 'block';
        document.getElementById('file-input').value = '';
      } catch (err) {
        alert('文件解析失败：' + err.message);
      }
    };
    reader.readAsArrayBuffer(file);
  };

  // === Drag & Drop ===
  (function() {
    var dropZone = document.getElementById('drop-zone');
    if (!dropZone) return;
    ['dragenter', 'dragover'].forEach(function(evt) {
      dropZone.addEventListener(evt, function(e) { e.preventDefault(); dropZone.classList.add('dragover'); });
    });
    ['dragleave', 'drop'].forEach(function(evt) {
      dropZone.addEventListener(evt, function(e) { e.preventDefault(); dropZone.classList.remove('dragover'); });
    });
    dropZone.addEventListener('drop', function(e) {
      var files = e.dataTransfer.files;
      if (files.length > 0) handleFile(files[0]);
    });
  })();

  // === Clock ===
  function updateClock() {
    var now = new Date();
    var str = now.getFullYear() + '-' +
      String(now.getMonth() + 1).padStart(2, '0') + '-' +
      String(now.getDate()).padStart(2, '0') + ' ' +
      String(now.getHours()).padStart(2, '0') + ':' +
      String(now.getMinutes()).padStart(2, '0') + ':' +
      String(now.getSeconds()).padStart(2, '0');
    document.getElementById('datetime').textContent = str;
  }
  setInterval(updateClock, 1000);
  updateClock();

  // === Start ===
  loadDefaultData();

})();
