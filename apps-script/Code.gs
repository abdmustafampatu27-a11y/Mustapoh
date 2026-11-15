/**
 * KPI ARCHIVE V3 + DASHBOARD OVERVIEW
 * Google Apps Script - satu file (Code.gs)
 *
 * Pemasangan:
 * 1. Buka Extensions > Apps Script dari Google Spreadsheet tracker.
 * 2. Hapus kode lama, tempel seluruh isi file ini ke Code.gs.
 * 3. Save, jalankan setupKpiSystem() satu kali, lalu izinkan akses.
 * 4. Deploy > New deployment > Web app untuk membuka Dashboard Overview.
 *
 * Arsip menyimpan snapshot nilai tampilan (bukan formula) per tanggal laporan.
 * Arsip tanggal + sheet yang sama akan diperbarui, bukan digandakan.
 */

const KPI_V3 = Object.freeze({
  ARCHIVE_SHEET: 'KPI Archive V3',
  LOG_SHEET: 'KPI System Log',
  TIMEZONE: 'Asia/Makassar',
  SNAPSHOT_HOUR: 23,
  CHUNK_SIZE: 40000,
  DEFAULT_DAYS: 30,
  SOURCES: [
    'Sameday STD',
    'Sameday Balanded',
    'Sameday Balanced',
    'FIFO Delivering',
    'FIFO Assigned',
    'View FIFO',
    'Region 5 Projection and propose'
  ],
  EXCLUDED: [
    'KPI Archive V3',
    'KPI System Log',
    'KPI Dashboard',
    'Dashboard Overview',
    'Looker Data',
    'link',
    'Form Responses 1'
  ]
});

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('KPI SYSTEM V3')
    .addItem('Setup / Aktifkan Sistem', 'setupKpiSystem')
    .addItem('Arsipkan Data Sekarang', 'archiveKpiNow')
    .addItem('Buka Dashboard', 'showDashboardUrl')
    .addSeparator()
    .addItem('Periksa Status Sistem', 'checkKpiSystem')
    .addToUi();
}

/**
 * Jalankan satu kali untuk membuat sheet arsip, log, dan trigger harian.
 */
function setupKpiSystem() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Spreadsheet aktif tidak ditemukan.');

  ensureArchiveSheet_(ss);
  ensureLogSheet_(ss);
  installDailyTrigger_();

  const result = archiveKpiSnapshot_(getReportDate_());
  logKpi_('SETUP', 'Setup selesai. ' + result.message);

  SpreadsheetApp.getUi().alert(
    'KPI SYSTEM V3 berhasil dipasang.\n\n' +
    'Arsip: ' + KPI_V3.ARCHIVE_SHEET + '\n' +
    'Trigger: setiap hari sekitar pukul ' + KPI_V3.SNAPSHOT_HOUR + ':00 WITA\n' +
    'Snapshot awal: ' + result.message + '\n\n' +
    'Untuk membuka dashboard, deploy project ini sebagai Web App.'
  );
}

/**
 * Handler trigger harian. Laporan default memakai tanggal kemarin,
 * mengikuti pola data tracker harian yang biasanya selesai setelah hari berjalan.
 */
function runDailyKpiArchive() {
  const result = archiveKpiSnapshot_(getReportDate_());
  logKpi_('AUTO ARCHIVE', result.message);
  return result;
}

/**
 * Menu manual. Tetap menggunakan tanggal laporan kemarin.
 */
function archiveKpiNow() {
  const result = archiveKpiSnapshot_(getReportDate_());
  SpreadsheetApp.getUi().alert(result.message);
  return result;
}

function archiveKpiSnapshot_(reportDate) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Spreadsheet aktif tidak ditemukan.');

  const archive = ensureArchiveSheet_(ss);
  const sources = findSourceSheets_(ss);

  if (!sources.length) {
    const message = 'Tidak ada sheet sumber KPI yang cocok. Periksa nama tab sumber.';
    logKpi_('WARNING', message);
    return { success: false, message: message, count: 0 };
  }

  let archived = 0;
  const errors = [];

  sources.forEach(function(source) {
    try {
      const values = source.range.getDisplayValues();
      if (!values || !values.length || !values.some(function(row) {
        return row.some(function(cell) { return String(cell).trim() !== ''; });
      })) {
        return;
      }

      const payload = JSON.stringify({
        source: source.sheet.getName(),
        range: source.range.getA1Notation(),
        capturedAt: new Date().toISOString(),
        values: values
      });

      const chunks = splitText_(payload, KPI_V3.CHUNK_SIZE);
      removeExistingSnapshot_(archive, reportDate, source.sheet.getName());

      const capturedAt = new Date();
      const rows = chunks.map(function(chunk, index) {
        return [
          reportDate,
          capturedAt,
          source.sheet.getName(),
          index + 1,
          chunks.length,
          chunk
        ];
      });

      archive.getRange(archive.getLastRow() + 1, 1, rows.length, 6).setValues(rows);
      archived++;
    } catch (err) {
      errors.push(source.sheet.getName() + ': ' + err.message);
    }
  });

  archive.getRange('A:A').setNumberFormat('@');
  archive.setFrozenRows(1);

  const message = 'Tanggal laporan ' + reportDate + ': ' + archived +
    ' sheet berhasil diarsipkan.' +
    (errors.length ? ' Gagal: ' + errors.join(' | ') : '');

  logKpi_(errors.length ? 'PARTIAL' : 'ARCHIVE', message);
  return { success: errors.length === 0, message: message, count: archived, errors: errors };
}

function findSourceSheets_(ss) {
  const sheets = ss.getSheets();
  const found = [];
  const used = {};

  KPI_V3.SOURCES.forEach(function(name) {
    const exact = sheets.find(function(sh) {
      return sh.getName().toLowerCase() === name.toLowerCase();
    });
    if (exact && !used[exact.getSheetId()]) {
      found.push(makeSourceDefinition_(exact));
      used[exact.getSheetId()] = true;
    }
  });

  // Penemuan otomatis untuk tab KPI tambahan dengan penamaan yang relevan.
  sheets.forEach(function(sh) {
    const name = sh.getName();
    const lower = name.toLowerCase();
    if (used[sh.getSheetId()] || KPI_V3.EXCLUDED.some(function(x) {
      return x.toLowerCase() === lower;
    })) return;

    const relevant = ['sameday', 'fifo', 'projection', 'propose', 'kpi'].some(function(key) {
      return lower.indexOf(key) !== -1;
    });

    if (relevant) {
      found.push(makeSourceDefinition_(sh));
      used[sh.getSheetId()] = true;
    }
  });

  return found;
}

function makeSourceDefinition_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 1) return { sheet: sheet, range: sheet.getRange(1, 1, 1, 1) };

  if (sheet.getName().toLowerCase() === 'view fifo') {
    const startCol = 36; // AJ
    const width = 14;    // AJ:AW
    return {
      sheet: sheet,
      range: sheet.getRange(1, startCol, lastRow, width)
    };
  }

  const lastCol = Math.max(sheet.getLastColumn(), 1);
  return {
    sheet: sheet,
    range: sheet.getRange(1, 1, lastRow, lastCol)
  };
}

function ensureArchiveSheet_(ss) {
  let sheet = ss.getSheetByName(KPI_V3.ARCHIVE_SHEET);
  if (!sheet) sheet = ss.insertSheet(KPI_V3.ARCHIVE_SHEET);

  const headers = [
    'Report Date',
    'Captured At',
    'Source Sheet',
    'Chunk No',
    'Chunk Total',
    'Payload'
  ];

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  } else {
    const existing = sheet.getRange(1, 1, 1, headers.length).getDisplayValues()[0];
    if (existing.join('|') !== headers.join('|')) {
      throw new Error(
        'Struktur sheet "' + KPI_V3.ARCHIVE_SHEET +
        '" tidak sesuai. Ganti nama atau kosongkan sheet arsip lama terlebih dahulu.'
      );
    }
  }

  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#252525')
    .setFontColor('#FFFFFF');
  sheet.setColumnWidth(1, 115);
  sheet.setColumnWidth(2, 170);
  sheet.setColumnWidth(3, 190);
  sheet.setColumnWidth(4, 80);
  sheet.setColumnWidth(5, 90);
  sheet.setColumnWidth(6, 420);
  return sheet;
}

function ensureLogSheet_(ss) {
  let sheet = ss.getSheetByName(KPI_V3.LOG_SHEET);
  if (!sheet) sheet = ss.insertSheet(KPI_V3.LOG_SHEET);

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, 3).setValues([['Timestamp', 'Event', 'Message']]);
    sheet.getRange(1, 1, 1, 3)
      .setFontWeight('bold')
      .setBackground('#252525')
      .setFontColor('#FFFFFF');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function removeExistingSnapshot_(archive, reportDate, sourceName) {
  const lastRow = archive.getLastRow();
  if (lastRow < 2) return;

  const keys = archive.getRange(2, 1, lastRow - 1, 3).getDisplayValues();
  const deleteRows = [];

  keys.forEach(function(row, index) {
    if (String(row[0]) === reportDate && String(row[2]) === sourceName) {
      deleteRows.push(index + 2);
    }
  });

  // Hapus dari bawah agar nomor baris tidak bergeser.
  deleteRows.reverse().forEach(function(rowNumber) {
    archive.deleteRow(rowNumber);
  });
}

function installDailyTrigger_() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'runDailyKpiArchive') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger('runDailyKpiArchive')
    .timeBased()
    .everyDays(1)
    .atHour(KPI_V3.SNAPSHOT_HOUR)
    .inTimezone(getTimezone_())
    .create();
}

function getReportDate_() {
  const date = new Date(Date.now() - 24 * 60 * 60 * 1000);
  return Utilities.formatDate(date, getTimezone_(), 'yyyy-MM-dd');
}

function getTimezone_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return (ss && ss.getSpreadsheetTimeZone()) || KPI_V3.TIMEZONE;
}

function splitText_(text, size) {
  const chunks = [];
  for (let i = 0; i < text.length; i += size) {
    chunks.push(text.substring(i, i + size));
  }
  return chunks.length ? chunks : [''];
}

function logKpi_(event, message) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ensureLogSheet_(ss);
    sheet.appendRow([new Date(), event, message]);
    sheet.getRange(sheet.getLastRow(), 1).setNumberFormat('yyyy-mm-dd hh:mm:ss');
  } catch (err) {
    console.error('KPI log error: ' + err.message);
  }
}

function checkKpiSystem() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const archive = ss.getSheetByName(KPI_V3.ARCHIVE_SHEET);
  const triggers = ScriptApp.getProjectTriggers().filter(function(trigger) {
    return trigger.getHandlerFunction() === 'runDailyKpiArchive';
  });
  const sources = findSourceSheets_(ss);

  const message = [
    'Sheet arsip: ' + (archive ? 'tersedia' : 'belum dibuat'),
    'Trigger harian: ' + triggers.length,
    'Sheet sumber terdeteksi: ' + sources.map(function(x) { return x.sheet.getName(); }).join(', ')
  ].join('\n');

  SpreadsheetApp.getUi().alert('Status KPI SYSTEM V3\n\n' + message);
}

function showDashboardUrl() {
  const url = ScriptApp.getService().getUrl();
  if (!url) {
    SpreadsheetApp.getUi().alert(
      'Web App belum di-deploy.\nBuka Deploy > New deployment > Web app, lalu pilih akses yang sesuai.'
    );
    return;
  }

  const html = HtmlService.createHtmlOutput(
    '<div style="font:14px Arial;padding:18px">' +
    '<p>Dashboard Overview siap dibuka:</p>' +
    '<p><a href="' + url + '" target="_blank">' + url + '</a></p>' +
    '</div>'
  ).setWidth(420).setHeight(150);

  SpreadsheetApp.getUi().showModalDialog(html, 'KPI Dashboard');
}

/**
 * Entry point Web App.
 * Deploy > New deployment > Web app.
 */
function doGet() {
  return HtmlService.createHtmlOutput(dashboardHtml_())
    .setTitle('KPI Dashboard Overview')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Dipanggil oleh halaman dashboard menggunakan google.script.run.
 */
function getDashboardData(days) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const archive = ss.getSheetByName(KPI_V3.ARCHIVE_SHEET);
  if (!archive || archive.getLastRow() < 2) {
    return emptyDashboardData_('Belum ada data arsip. Jalankan setupKpiSystem() terlebih dahulu.');
  }

  const rangeDays = Number(days || KPI_V3.DEFAULT_DAYS);
  const cutoff = rangeDays > 0
    ? Utilities.formatDate(
        new Date(Date.now() - rangeDays * 24 * 60 * 60 * 1000),
        getTimezone_(),
        'yyyy-MM-dd'
      )
    : '0000-00-00';

  const rows = archive.getRange(2, 1, archive.getLastRow() - 1, 6).getValues();
  const groups = {};

  rows.forEach(function(row) {
    const reportDate = normalizeDateKey_(row[0]);
    const sourceName = String(row[2] || '');
    if (!reportDate || reportDate < cutoff || !sourceName) return;

    const key = reportDate + '||' + sourceName;
    if (!groups[key]) {
      groups[key] = {
        date: reportDate,
        source: sourceName,
        chunks: [],
        total: Number(row[4]) || 0,
        capturedAt: row[1]
      };
    }
    groups[key].chunks[Number(row[3]) - 1] = String(row[5] || '');
  });

  const snapshots = [];
  Object.keys(groups).forEach(function(key) {
    const group = groups[key];
    try {
      const payload = JSON.parse(group.chunks.join(''));
      if (payload && payload.values) {
        snapshots.push({
          date: group.date,
          source: group.source,
          capturedAt: group.capturedAt,
          range: payload.range || '',
          values: payload.values
        });
      }
    } catch (err) {
      // Abaikan satu snapshot rusak dan lanjutkan membaca snapshot lain.
    }
  });

  if (!snapshots.length) {
    return emptyDashboardData_('Belum ada arsip pada rentang tanggal yang dipilih.');
  }

  snapshots.sort(function(a, b) {
    return a.date.localeCompare(b.date) || a.source.localeCompare(b.source);
  });

  return buildDashboardModel_(snapshots);
}

function normalizeDateKey_(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, getTimezone_(), 'yyyy-MM-dd');
  }
  const text = String(value || '').trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(text)) return text;
  const parsed = new Date(value);
  return isNaN(parsed.getTime())
    ? text
    : Utilities.formatDate(parsed, getTimezone_(), 'yyyy-MM-dd');
}

function emptyDashboardData_(message) {
  return {
    empty: true,
    message: message,
    latestDate: '-',
    hubCount: 0,
    sourceCount: 0,
    rowCount: 0,
    metricCount: 0,
    metrics: [],
    trend: { labels: [], series: [] },
    sources: []
  };
}

function buildDashboardModel_(snapshots) {
  const latestDate = snapshots[snapshots.length - 1].date;
  const latestBySource = {};
  const trendMap = {};
  const hubSet = {};
  let totalRows = 0;
  let metricNames = {};

  snapshots.forEach(function(snapshot) {
    const analyzed = analyzeSnapshot_(snapshot);
    if (!latestBySource[snapshot.source] ||
        snapshot.date > latestBySource[snapshot.source].date) {
      latestBySource[snapshot.source] = analyzed;
    }

    totalRows += analyzed.rowCount;

    analyzed.metrics.forEach(function(metric) {
      metricNames[metric.name] = true;
      if (!trendMap[metric.name]) trendMap[metric.name] = {};
      trendMap[metric.name][snapshot.date] =
        (trendMap[metric.name][snapshot.date] || 0) + metric.value;
    });
  });

  const latestSources = Object.keys(latestBySource).map(function(name) {
    const source = latestBySource[name];
    source.hubs.forEach(function(hub) {
      hubSet[hub.toLowerCase()] = true;
    });
    return source;
  }).sort(function(a, b) {
    return a.name.localeCompare(b.name);
  });

  const dates = Array.from(new Set(snapshots.map(function(x) { return x.date; }))).sort();
  const metricTotals = {};

  latestSources.forEach(function(source) {
    source.metrics.forEach(function(metric) {
      metricTotals[metric.name] = (metricTotals[metric.name] || 0) + metric.value;
    });
  });

  const metricCards = Object.keys(metricTotals).map(function(name) {
    return { name: name, value: metricTotals[name] };
  }).sort(function(a, b) {
    return Math.abs(b.value) - Math.abs(a.value);
  }).slice(0, 8);

  const trendSeries = Object.keys(trendMap).map(function(name) {
    return {
      name: name,
      values: dates.map(function(date) {
        return trendMap[name][date] || 0;
      })
    };
  }).sort(function(a, b) {
    const totalA = a.values.reduce(function(sum, value) { return sum + Math.abs(value); }, 0);
    const totalB = b.values.reduce(function(sum, value) { return sum + Math.abs(value); }, 0);
    return totalB - totalA;
  }).slice(0, 5);

  return {
    empty: false,
    latestDate: latestDate,
    hubCount: Object.keys(hubSet).length,
    sourceCount: latestSources.length,
    rowCount: totalRows,
    metricCount: Object.keys(metricNames).length,
    metrics: metricCards,
    trend: {
      labels: dates,
      series: trendSeries
    },
    sources: latestSources.map(function(source) {
      return {
        name: source.name,
        date: source.date,
        rowCount: source.rowCount,
        hubCount: source.hubs.length,
        metrics: source.metrics.slice(0, 8),
        headers: source.headers,
        rows: source.rows.slice(0, 80)
      };
    })
  };
}

function analyzeSnapshot_(snapshot) {
  const matrix = snapshot.values || [];
  const headerIndex = findHeaderRow_(matrix);
  const headers = (matrix[headerIndex] || []).map(function(value, index) {
    const name = String(value || '').trim();
    return name || ('Kolom ' + (index + 1));
  });

  const dataRows = matrix.slice(headerIndex + 1).filter(function(row) {
    return row.some(function(cell) { return String(cell).trim() !== ''; });
  });

  const hubIndex = findHubColumn_(headers);
  const metricIndexes = findMetricColumns_(headers, dataRows);
  const hubs = [];
  const hubSeen = {};
  if (hubIndex >= 0) {
    dataRows.forEach(function(row) {
      const hub = String(row[hubIndex] || '').trim();
      if (hub && !hubSeen[hub.toLowerCase()]) {
        hubSeen[hub.toLowerCase()] = true;
        hubs.push(hub);
      }
    });
  }

  const metrics = metricIndexes.map(function(index) {
    let sum = 0;
    dataRows.forEach(function(row) {
      const number = parseMetricNumber_(row[index]);
      if (number !== null) sum += number;
    });
    return { name: headers[index], value: sum };
  }).sort(function(a, b) {
    return Math.abs(b.value) - Math.abs(a.value);
  });

  return {
    name: snapshot.source,
    date: snapshot.date,
    rowCount: dataRows.length,
    hubs: hubs,
    metrics: metrics,
    headers: headers,
    rows: dataRows
  };
}

function findHeaderRow_(matrix) {
  const cues = [
    'hub', 'station', 'site', 'facility', 'branch', 'location', 'area',
    'date', 'tanggal', 'delivered', 'assigned', 'assign', 'delivery',
    'productivity', 'target', 'actual', 'kpi', 'rider', 'driver',
    'operator', 'achievement', 'pencapaian', 'fifo', 'sameday',
    'status', 'volume', 'sla', 'region'
  ];
  let bestIndex = 0;
  let bestScore = -1;
  const limit = Math.min(matrix.length, 15);

  for (let r = 0; r < limit; r++) {
    const row = matrix[r] || [];
    const filled = row.map(function(cell) {
      return String(cell || '').trim();
    }).filter(function(cell) { return cell !== ''; });
    if (filled.length < 2) continue;

    const cueHits = filled.filter(function(cell) {
      const value = cell.toLowerCase().replace(/[_-]/g, ' ');
      return cues.some(function(cue) { return value.indexOf(cue) !== -1; });
    }).length;

    const textCount = filled.filter(function(cell) {
      return isNaN(Number(cell.replace(/,/g, '')));
    }).length;
    const textRatio = textCount / filled.length;
    const score = cueHits * 8 + Math.min(filled.length, 40) * 0.25 +
      textRatio * 2 - r * 0.05;

    if (score > bestScore) {
      bestScore = score;
      bestIndex = r;
    }
  }
  return bestIndex;
}

function findHubColumn_(headers) {
  const aliases = [
    'hub', 'hub name', 'nama hub', 'station', 'station name',
    'site', 'facility', 'branch', 'location', 'lokasi', 'area'
  ];

  for (let i = 0; i < headers.length; i++) {
    const header = headers[i].toLowerCase().replace(/[_-]/g, ' ').trim();
    if (aliases.some(function(alias) {
      return header === alias || header.indexOf(alias) !== -1;
    })) return i;
  }
  return -1;
}

function findMetricColumns_(headers, rows) {
  const excluded = [
    'id', 'nik', 'phone', 'telephone', 'latitude', 'longitude',
    'postal', 'zipcode', 'kode', 'code', 'tanggal', 'date',
    'time', 'waktu', 'no', 'nomor', 'index'
  ];
  const result = [];

  headers.forEach(function(header, index) {
    const name = String(header || '').trim();
    const lower = name.toLowerCase();
    if (!name || excluded.some(function(word) {
      return lower === word || lower.indexOf(word + ' ') === 0;
    })) return;

    let inspected = 0;
    let numeric = 0;
    rows.slice(0, 60).forEach(function(row) {
      const value = row[index];
      if (String(value).trim() === '') return;
      inspected++;
      if (parseMetricNumber_(value) !== null) numeric++;
    });

    if (inspected > 0 && numeric / inspected >= 0.6) {
      result.push(index);
    }
  });

  return result;
}

function parseMetricNumber_(value) {
  if (typeof value === 'number' && isFinite(value)) return value;
  let text = String(value || '').trim();
  if (!text) return null;
  text = text.replace(/\s/g, '').replace(/%$/, '');
  if (!/^[+-]?[0-9.,]+$/.test(text)) return null;

  const comma = text.lastIndexOf(',');
  const dot = text.lastIndexOf('.');
  if (comma !== -1 && dot !== -1) {
    if (comma > dot) {
      text = text.replace(/\./g, '').replace(',', '.');
    } else {
      text = text.replace(/,/g, '');
    }
  } else if (comma !== -1) {
    const decimals = text.length - comma - 1;
    text = decimals === 3 ? text.replace(/,/g, '') : text.replace(',', '.');
  } else if (dot !== -1) {
    const decimals = text.length - dot - 1;
    if (decimals === 3) text = text.replace(/\./g, '');
  }

  const number = Number(text);
  return isFinite(number) ? number : null;
}

/**
 * HTML Dashboard satu halaman panjang / scrollable.
 * Tema terang: hitam arang + aksen orange, tidak terlalu gelap.
 */
function dashboardHtml_() {
  return [
'<!doctype html>',
'<html lang="id">',
'<head>',
'<meta charset="utf-8">',
'<meta name="viewport" content="width=device-width, initial-scale=1">',
'<title>KPI Dashboard Overview</title>',
'<style>',
':root{--orange:#ef7b24;--ink:#272727;--muted:#737373;--line:#e9e9e9;--bg:#f5f5f4;--white:#fff}',
'*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--ink);font:14px/1.5 Arial,Helvetica,sans-serif}',
'.wrap{max-width:1440px;margin:auto;padding:28px 30px 56px}',
'.top{display:flex;justify-content:space-between;align-items:center;gap:18px;margin-bottom:24px}',
'.brand{display:flex;align-items:center;gap:13px}.mark{width:43px;height:43px;border-radius:13px;background:var(--orange);color:#fff;display:grid;place-items:center;font-size:19px;font-weight:800}',
'h1{font-size:25px;margin:0 0 3px;letter-spacing:-.5px}.sub{color:var(--muted);font-size:12px}',
'.actions{display:flex;gap:10px;align-items:center;flex-wrap:wrap}select,button{border:1px solid #dedede;background:#fff;border-radius:9px;padding:10px 13px;color:var(--ink);font:inherit}',
'button{background:var(--ink);color:#fff;border-color:var(--ink);cursor:pointer}button:hover{background:#444}',
'.stamp{color:var(--muted);font-size:12px;margin:0 0 17px}',
'.grid{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:14px;margin-bottom:20px}',
'.card,.panel,.source{background:var(--white);border:1px solid var(--line);border-radius:15px;box-shadow:0 3px 12px rgba(0,0,0,.025)}',
'.card{padding:18px 19px;position:relative;overflow:hidden}.card:after{content:"";position:absolute;right:-17px;top:-17px;width:64px;height:64px;border-radius:50%;background:#fff4eb}',
'.klabel{color:var(--muted);font-size:12px;margin-bottom:8px}.kvalue{font-size:28px;font-weight:750;letter-spacing:-.7px}.kfoot{font-size:11px;color:#999;margin-top:5px}',
'.panel{padding:20px;margin-bottom:19px}.panelhead{display:flex;justify-content:space-between;align-items:flex-start;gap:12px;margin-bottom:17px}',
'h2{margin:0;font-size:16px;letter-spacing:-.2px}.hint{color:var(--muted);font-size:11px;margin-top:3px}',
'.metricgrid{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:12px}',
'.metric{border:1px solid var(--line);border-radius:12px;padding:14px;background:#fff}.metricname{font-size:11px;color:var(--muted);min-height:31px}.metricvalue{font-size:20px;font-weight:700;margin-top:5px}',
'.chartbox{height:270px;position:relative}canvas{width:100%;height:100%;display:block}',
'.legend{display:flex;flex-wrap:wrap;gap:14px;margin-top:10px}.legend span{font-size:11px;color:var(--muted);display:flex;align-items:center;gap:6px}.dot{width:8px;height:8px;border-radius:50%;display:inline-block}',
'.sources{display:grid;gap:15px}.source{overflow:hidden}.sourcehead{padding:16px 18px;display:flex;justify-content:space-between;gap:15px;align-items:center;border-bottom:1px solid var(--line)}',
'.sourcehead strong{font-size:14px}.sourcemeta{font-size:11px;color:var(--muted);margin-top:3px}.sourcebody{padding:16px 18px}',
'.mini{display:flex;gap:9px;flex-wrap:wrap;margin-bottom:13px}.pill{padding:7px 10px;border-radius:8px;background:#fff5ed;color:#9b4e13;font-size:11px}',
'.tablewrap{overflow:auto;border:1px solid var(--line);border-radius:10px}table{border-collapse:collapse;width:100%;min-width:650px;font-size:11px}th,td{padding:9px 11px;border-bottom:1px solid #eee;text-align:left;white-space:nowrap;max-width:240px;overflow:hidden;text-overflow:ellipsis}',
'th{background:#f8f8f7;color:#555;font-weight:700;position:sticky;top:0}tr:last-child td{border-bottom:0}',
'.state{padding:28px;text-align:center;color:var(--muted);background:#fff;border:1px dashed #ddd;border-radius:14px}',
'.error{color:#a43d20;background:#fff6f2;border:1px solid #f2d3c7;padding:13px 15px;border-radius:10px;margin:15px 0}',
'.footer{text-align:center;color:#999;font-size:11px;padding-top:13px}',
'@media(max-width:900px){.grid{grid-template-columns:repeat(2,minmax(0,1fr))}.metricgrid{grid-template-columns:repeat(2,minmax(0,1fr))}}',
'@media(max-width:600px){.wrap{padding:18px 13px 38px}.top{align-items:flex-start;flex-direction:column}h1{font-size:22px}.grid{gap:9px}.card{padding:14px}.kvalue{font-size:24px}.panel{padding:15px}.metricgrid{gap:8px}.chartbox{height:220px}}',
'</style>',
'</head>',
'<body>',
'<main class="wrap">',
'<header class="top">',
'<div class="brand"><div class="mark">K</div><div><h1>Dashboard Overview</h1><div class="sub">KPI performance · Hub operations · Historical trends</div></div></div>',
'<div class="actions"><select id="period" onchange="loadData()"><option value="7">7 hari</option><option value="30" selected>30 hari</option><option value="90">90 hari</option><option value="365">365 hari</option><option value="0">Semua data</option></select><button onclick="loadData()">Refresh Data</button></div>',
'</header>',
'<div class="stamp" id="stamp">Memuat data KPI...</div>',
'<div id="notice"></div>',
'<section class="grid">',
'<div class="card"><div class="klabel">Total Hub Terdeteksi</div><div class="kvalue" id="hubCount">-</div><div class="kfoot">Berdasarkan data terbaru</div></div>',
'<div class="card"><div class="klabel">Sumber KPI</div><div class="kvalue" id="sourceCount">-</div><div class="kfoot">Sheet tracker yang terarsip</div></div>',
'<div class="card"><div class="klabel">Baris Data</div><div class="kvalue" id="rowCount">-</div><div class="kfoot">Snapshot dalam periode terpilih</div></div>',
'<div class="card"><div class="klabel">Jenis Metrik</div><div class="kvalue" id="metricCount">-</div><div class="kfoot">Kolom numerik yang terdeteksi otomatis</div></div>',
'</section>',
'<section class="panel"><div class="panelhead"><div><h2>Ringkasan KPI</h2><div class="hint">Akumulasi metrik dari snapshot terbaru setiap sumber</div></div></div><div class="metricgrid" id="metricCards"><div class="state">Belum ada metrik.</div></div></section>',
'<section class="panel"><div class="panelhead"><div><h2>Historical Trends</h2><div class="hint">Pergerakan hingga lima metrik numerik teratas sepanjang periode</div></div></div><div class="chartbox"><canvas id="trendChart"></canvas></div><div class="legend" id="legend"></div></section>',
'<section class="panel"><div class="panelhead"><div><h2>Detail Performa per Sumber</h2><div class="hint">Data terbaru yang tersedia untuk setiap sheet KPI</div></div></div><div class="sources" id="sourceList"><div class="state">Memuat sumber data...</div></div></section>',
'<div class="footer">KPI SYSTEM V3 · Data bersumber dari Google Spreadsheet tracker</div>',
'</main>',
'<script>',
'var chartData=null;',
'function fmt(v){return Number(v||0).toLocaleString("id-ID",{maximumFractionDigits:2});}',
'function esc(v){return String(v==null?"":v).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;");}',
'function loadData(){',
'  document.getElementById("stamp").textContent="Memuat data KPI...";',
'  document.getElementById("notice").innerHTML="";',
'  google.script.run.withSuccessHandler(renderDashboard).withFailureHandler(function(e){',
'    document.getElementById("notice").innerHTML="<div class=\"error\">Gagal memuat dashboard: "+esc(e.message||e)+"</div>";',
'    document.getElementById("stamp").textContent="Pembaruan gagal";',
'  }).getDashboardData(Number(document.getElementById("period").value));',
'}',
'function renderDashboard(data){',
'  chartData=data;',
'  document.getElementById("stamp").textContent=data.latestDate==="-"?"Belum ada arsip":"Tanggal laporan terbaru: "+data.latestDate;',
'  document.getElementById("hubCount").textContent=fmt(data.hubCount);',
'  document.getElementById("sourceCount").textContent=fmt(data.sourceCount);',
'  document.getElementById("rowCount").textContent=fmt(data.rowCount);',
'  document.getElementById("metricCount").textContent=fmt(data.metricCount);',
'  if(data.empty){',
'    document.getElementById("notice").innerHTML="<div class=\"state\">"+esc(data.message)+"</div>";',
'    document.getElementById("metricCards").innerHTML="<div class=\"state\">Belum ada data KPI.</div>";',
'    document.getElementById("sourceList").innerHTML="";',
'    drawChart({labels:[],series:[]});return;',
'  }',
'  document.getElementById("metricCards").innerHTML=data.metrics.length?data.metrics.map(function(m){',
'    return "<div class=\"metric\"><div class=\"metricname\">"+esc(m.name)+"</div><div class=\"metricvalue\">"+fmt(m.value)+"</div></div>";',
'  }).join(""):"<div class=\"state\">Kolom metrik numerik belum terdeteksi.</div>";',
'  document.getElementById("sourceList").innerHTML=data.sources.map(renderSource).join("");',
'  drawChart(data.trend);',
'}',
'function renderSource(source){',
'  var pills=source.metrics.map(function(m){return "<span class=\"pill\">"+esc(m.name)+": <b>"+fmt(m.value)+"</b></span>";}).join("");',
'  var headers=source.headers||[];',
'  var rows=source.rows||[];',
'  var table="<div class=\"tablewrap\"><table><thead><tr>"+headers.map(function(h){return "<th>"+esc(h)+"</th>";}).join("")+"</tr></thead><tbody>";',
'  table+=rows.slice(0,30).map(function(row){return "<tr>"+headers.map(function(h,i){return "<td>"+esc(row[i])+"</td>";}).join("")+"</tr>";}).join("");',
'  table+="</tbody></table></div>";',
'  return "<article class=\"source\"><div class=\"sourcehead\"><div><strong>"+esc(source.name)+"</strong><div class=\"sourcemeta\">Tanggal: "+esc(source.date)+" · "+fmt(source.rowCount)+" baris · "+fmt(source.hubCount)+" hub</div></div></div><div class=\"sourcebody\"><div class=\"mini\">"+(pills||"<span class=\"sourcemeta\">Belum ada kolom metrik numerik yang dapat dibaca.</span>")+"</div>"+(rows.length?table:"<div class=\"state\">Tidak ada baris data.</div>")+"</div></article>";',
'}',
'function drawChart(data){',
'  var canvas=document.getElementById("trendChart"),ctx=canvas.getContext("2d");',
'  var w=canvas.clientWidth||700,h=canvas.clientHeight||260,dpr=window.devicePixelRatio||1;',
'  canvas.width=w*dpr;canvas.height=h*dpr;ctx.setTransform(dpr,0,0,dpr,0,0);ctx.clearRect(0,0,w,h);',
'  document.getElementById("legend").innerHTML="";',
'  if(!data||!data.labels||data.labels.length<1||!data.series||data.series.length<1){',
'    ctx.fillStyle="#888";ctx.font="12px Arial";ctx.fillText("Tren akan tampil setelah data KPI berhasil diarsipkan.",12,28);return;',
'  }',
'  var colors=["#ef7b24","#343434","#8c9a9e","#d8a15d","#718c78"];',
'  var left=42,right=12,top=12,bottom=34,plotW=w-left-right,plotH=h-top-bottom;',
'  var max=0;data.series.forEach(function(s){s.values.forEach(function(v){max=Math.max(max,Math.abs(Number(v)||0));});});if(max===0)max=1;',
'  ctx.strokeStyle="#ededed";ctx.lineWidth=1;ctx.fillStyle="#888";ctx.font="10px Arial";ctx.textAlign="right";',
'  for(var g=0;g<=4;g++){var y=top+plotH*g/4;ctx.beginPath();ctx.moveTo(left,y);ctx.lineTo(w-right,y);ctx.stroke();ctx.fillText(fmt(max*(4-g)/4),left-8,y+3);}',
'  var count=data.labels.length;',
'  data.series.forEach(function(s,si){',
'    ctx.beginPath();ctx.lineWidth=2.3;ctx.strokeStyle=colors[si%colors.length];',
'    s.values.forEach(function(value,i){var x=left+(count===1?plotW/2:i*plotW/(count-1));var y=top+plotH-(Number(value)||0)/max*plotH;if(i===0)ctx.moveTo(x,y);else ctx.lineTo(x,y);});ctx.stroke();',
'  });',
'  ctx.fillStyle="#888";ctx.textAlign="center";ctx.font="10px Arial";',
'  var step=Math.max(1,Math.ceil(count/6));data.labels.forEach(function(label,i){if(i%step===0||i===count-1){var x=left+(count===1?plotW/2:i*plotW/(count-1));ctx.fillText(label.slice(5),x,h-10);}});',
'  document.getElementById("legend").innerHTML=data.series.map(function(s,i){return "<span><i class=\"dot\" style=\"background:"+colors[i%colors.length]+"\"></i>"+esc(s.name)+"</span>";}).join("");',
'}',
'window.addEventListener("resize",function(){if(chartData&&chartData.trend)drawChart(chartData.trend);});',
'loadData();',
'</script>',
'</body>',
'</html>'
  ].join('\n');
}
