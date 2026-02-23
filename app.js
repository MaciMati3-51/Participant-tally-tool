'use strict';

/* ============================================================
   参加者 CSV 集計ツール — メインロジック
   ============================================================ */

// ---- 列アルファベット → 0始まりインデックス ----
function colLetterToIndex(col) {
  col = col.toUpperCase().trim().replace(/[^A-Z]/g, '');
  if (!col) return -1;
  let idx = 0;
  for (let i = 0; i < col.length; i++) {
    idx = idx * 26 + (col.charCodeAt(i) - 64);
  }
  return idx - 1;
}

// ---- CSV フィールドエスケープ (RFC 4180) ----
function escField(val) {
  const s = String(val != null ? val : '');
  return /[,"\n\r]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
}

// ---- CSV 文字列構築 ----
function buildCSV(headers, rows) {
  const lines = [headers.map(escField).join(',')];
  for (const row of rows) lines.push(row.map(escField).join(','));
  return lines.join('\n');
}

// ---- CSV ダウンロード（UTF-8 BOM 付き） ----
function downloadCSV(filename, content) {
  const blob = new Blob(['\uFEFF' + content], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 200);
}

// ---- ISO 8601 タイムスタンプ ----
function isoNow() {
  const d = new Date();
  const p = n => String(n).padStart(2, '0');
  const off = -d.getTimezoneOffset();
  const sign = off >= 0 ? '+' : '-';
  const abs = Math.abs(off);
  const tz = `${sign}${p(Math.floor(abs / 60))}:${p(abs % 60)}`;
  return `${d.getFullYear()}-${p(d.getMonth()+1)}-${p(d.getDate())}` +
         `T${p(d.getHours())}:${p(d.getMinutes())}:${p(d.getSeconds())}${tz}`;
}

// ---- HTML エスケープ ----
function esc(s) {
  return String(s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;')
    .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

// ---- ファイル読み込み (UTF-8) ----
function readText(file) {
  return new Promise((resolve, reject) => {
    const r = new FileReader();
    r.onload  = e => resolve(e.target.result);
    r.onerror = () => reject(new Error('読み込み失敗: ' + file.name));
    r.readAsText(file, 'UTF-8');
  });
}

// ---- CSV パース（Papa Parse 使用）+ ヘッダー行スキップ ----
async function parseCSV(file, headerRow) {
  const text = await readText(file);
  return new Promise((resolve, reject) => {
    Papa.parse(text, {
      header: false,
      skipEmptyLines: true,
      complete: result => {
        // headerRow は 1始まり → 0始まりに変換してスライス
        resolve(result.data.slice(headerRow));
      },
      error: err => reject(new Error(err.message)),
    });
  });
}

/* ============================================================
   集計コア
   ============================================================ */
async function aggregate(files, cfg) {
  const { pkIdx, nameIdx, statusIdx, condition, headerRow } = cfg;

  // ① ファイル名重複の検出
  const nameCnt = {};
  files.forEach(f => { nameCnt[f.name] = (nameCnt[f.name] || 0) + 1; });
  const dupNames = new Set(
    Object.entries(nameCnt).filter(([, c]) => c > 1).map(([n]) => n)
  );

  const toProcess = files.filter(f => !dupNames.has(f.name));
  const skipped   = [...dupNames].map(n => ({ name: n, reason: 'duplicate_filename' }));

  // ② 集計変数
  const master   = new Map(); // primaryKey → { name, count }
  let grossTotal = 0;
  let attendTotal = 0;
  const attendUniq = new Set();
  let ignoredNoKey  = 0;
  let ignoredNoName = 0;
  let processed = 0;

  for (const file of toProcess) {
    let rows;
    try {
      rows = await parseCSV(file, headerRow);
    } catch {
      skipped.push({ name: file.name, reason: 'parse_error' });
      continue;
    }

    const eventApplicants = new Set(); // イベント内ユニーク申込者
    const eventAttendees  = new Set(); // イベント内ユニーク参加者

    for (const row of rows) {
      const pk     = row[pkIdx]     != null ? String(row[pkIdx]).trim()     : '';
      const name   = row[nameIdx]   != null ? String(row[nameIdx]).trim()   : '';
      const status = row[statusIdx] != null ? String(row[statusIdx]).trim() : '';

      // バリデーション
      if (!pk)   { ignoredNoKey++;  continue; }
      if (!name) { ignoredNoName++; continue; }

      // 初回登場の名前を採用（表記揺れ防止）
      if (!master.has(pk)) master.set(pk, { name, count: 0 });

      // 申込者（主キー+名前が有効ならカウント）
      eventApplicants.add(pk);

      // 参加者（条件一致、trim + 大文字小文字無視）
      if (status.toLowerCase() === condition.toLowerCase()) {
        eventAttendees.add(pk);
      }
    }

    // イベント単位で集計に加算
    grossTotal  += eventApplicants.size;
    attendTotal += eventAttendees.size;

    for (const pk of eventAttendees) {
      master.get(pk).count++;
      attendUniq.add(pk);
    }

    processed++;
  }

  // ③ master_counts 行を生成（count 降順 → primaryKey 昇順）
  const masterRows = [...master.entries()]
    .map(([pk, { name, count }]) => [pk, name, count])
    .sort((a, b) => b[2] - a[2] || String(a[0]).localeCompare(String(b[0]), 'ja'));

  const rate = grossTotal > 0
    ? (attendTotal / grossTotal).toFixed(6)
    : 0;

  return {
    processedCount: processed,
    skipped,
    masterRows,
    summary: {
      grossTotal,
      attendTotal,
      attendUnique: attendUniq.size,
      rate,
      ignoredNoKey,
      ignoredNoName,
    },
  };
}

/* ============================================================
   CSV ビルダー
   ============================================================ */
function buildMasterCSV(rows) {
  return buildCSV(['主キー', '名前', '参加回数'], rows);
}

function buildSummaryCSV(result) {
  const { processedCount, skipped, summary } = result;
  return buildCSV(
    [
      '集計日時', '処理ファイル数', 'スキップファイル数',
      '総申込人数', '延べ人数', 'ユニーク人数',
      '実質参加割合', '無視行_主キー欠損', '無視行_名前欠損',
    ],
    [[
      isoNow(),
      processedCount,
      skipped.length,
      summary.grossTotal,
      summary.attendTotal,
      summary.attendUnique,
      summary.rate,
      summary.ignoredNoKey,
      summary.ignoredNoName,
    ]]
  );
}

/* ============================================================
   UI
   ============================================================ */
document.addEventListener('DOMContentLoaded', () => {
  let uploadedFiles = [];
  let result = null;

  // 要素取得
  const dropZone       = document.getElementById('drop-zone');
  const fileInput      = document.getElementById('file-input');
  const fileSelectBtn  = document.getElementById('file-select-btn');
  const fileListEl     = document.getElementById('file-list');
  const runBtn         = document.getElementById('run-btn');
  const resultsSection = document.getElementById('results-section');

  const headerRowInput = document.getElementById('header-row');
  const pkColInput     = document.getElementById('primary-key-col');
  const nameColInput   = document.getElementById('name-col');
  const statusColInput = document.getElementById('status-col');
  const condInput      = document.getElementById('condition-str');

  const DEFAULT_SETTINGS = Object.freeze({
    headerRow: '1',
    pkCol: 'E',
    nameCol: 'B',
    statusCol: 'H',
    condition: 'approved',
  });

  // HTML の value 属性だけだとブラウザの復元値に上書きされることがあるため、
  // 初期表示時に JS から規定値を明示的に適用する。
  function applyDefaultSettings() {
    headerRowInput.value = DEFAULT_SETTINGS.headerRow;
    headerRowInput.defaultValue = DEFAULT_SETTINGS.headerRow;
    pkColInput.value = DEFAULT_SETTINGS.pkCol;
    pkColInput.defaultValue = DEFAULT_SETTINGS.pkCol;
    nameColInput.value = DEFAULT_SETTINGS.nameCol;
    nameColInput.defaultValue = DEFAULT_SETTINGS.nameCol;
    statusColInput.value = DEFAULT_SETTINGS.statusCol;
    statusColInput.defaultValue = DEFAULT_SETTINGS.statusCol;
    condInput.value = DEFAULT_SETTINGS.condition;
    condInput.defaultValue = DEFAULT_SETTINGS.condition;
  }

  applyDefaultSettings();

  // ---- ファイル投入 ----
  dropZone.addEventListener('dragover', e => {
    e.preventDefault();
    dropZone.classList.add('drag-over');
  });

  dropZone.addEventListener('dragleave', e => {
    if (!dropZone.contains(e.relatedTarget)) dropZone.classList.remove('drag-over');
  });

  dropZone.addEventListener('drop', e => {
    e.preventDefault();
    dropZone.classList.remove('drag-over');
    addFiles([...e.dataTransfer.files].filter(f => f.name.toLowerCase().endsWith('.csv')));
  });

  fileSelectBtn.addEventListener('click', () => fileInput.click());

  fileInput.addEventListener('change', () => {
    addFiles([...fileInput.files]);
    fileInput.value = '';
  });

  function addFiles(newFiles) {
    uploadedFiles.push(...newFiles);
    renderFileList();
    updateRunBtn();
  }

  function removeFile(i) {
    uploadedFiles.splice(i, 1);
    renderFileList();
    updateRunBtn();
  }

  function renderFileList() {
    if (!uploadedFiles.length) { fileListEl.innerHTML = ''; return; }

    const cnt = {};
    uploadedFiles.forEach(f => { cnt[f.name] = (cnt[f.name] || 0) + 1; });

    fileListEl.innerHTML = uploadedFiles.map((f, i) => {
      const dup = cnt[f.name] > 1;
      return `<div class="file-item${dup ? ' file-dup' : ''}">
        <span class="file-name" title="${esc(f.name)}">${esc(f.name)}</span>
        ${dup ? '<span class="badge-dup">重複</span>' : ''}
        <button class="file-remove-btn" data-i="${i}" title="削除">&#10005;</button>
      </div>`;
    }).join('');

    fileListEl.querySelectorAll('.file-remove-btn').forEach(btn => {
      btn.addEventListener('click', () => removeFile(+btn.dataset.i));
    });
  }

  // ---- 必須入力チェック ----
  const required = [pkColInput, nameColInput, statusColInput, condInput];
  required.forEach(el => el.addEventListener('input', updateRunBtn));

  // 列入力は自動大文字化
  [pkColInput, nameColInput, statusColInput].forEach(el => {
    el.addEventListener('input', () => {
      const pos = el.selectionStart;
      el.value = el.value.toUpperCase();
      el.setSelectionRange(pos, pos);
    });
  });

  function updateRunBtn() {
    const ok = uploadedFiles.length > 0 &&
               required.every(el => el.value.trim() !== '');
    runBtn.disabled = !ok;
  }

  updateRunBtn();

  // ---- 集計実行 ----
  runBtn.addEventListener('click', async () => {
    const headerRow = Math.max(1, parseInt(headerRowInput.value) || 1);

    const pkIdx     = colLetterToIndex(pkColInput.value);
    const nameIdx   = colLetterToIndex(nameColInput.value);
    const statusIdx = colLetterToIndex(statusColInput.value);

    if (pkIdx < 0 || nameIdx < 0 || statusIdx < 0) {
      alert('列指定が無効です。A〜Z または AA〜AZ などの形式で入力してください。');
      return;
    }

    runBtn.disabled = true;
    runBtn.textContent = '集計中…';
    result = null;

    try {
      result = await aggregate(uploadedFiles, {
        pkIdx, nameIdx, statusIdx,
        condition: condInput.value.trim(),
        headerRow,
      });
      renderResults(result);
    } catch (e) {
      console.error(e);
      alert('集計中にエラーが発生しました:\n' + e.message);
    } finally {
      runBtn.textContent = '集計する';
      updateRunBtn();
    }
  });

  // ---- 結果表示 ----
  function renderResults(r) {
    const { summary, processedCount, skipped } = r;

    setText('res-processed-files',   processedCount);
    setText('res-skipped-files',     skipped.length);
    setText('res-gross-applications',summary.grossTotal.toLocaleString('ja-JP'));
    setText('res-attended-total',    summary.attendTotal.toLocaleString('ja-JP'));
    setText('res-attended-unique',   summary.attendUnique.toLocaleString('ja-JP'));

    const rateDisp = summary.grossTotal > 0
      ? (parseFloat(summary.rate) * 100).toFixed(1) + '%'
      : '—';
    setText('res-effective-rate', rateDisp);

    // スキップ警告
    const warn = document.getElementById('skipped-warning');
    const list = document.getElementById('skipped-file-list');
    const dups = skipped.filter(s => s.reason === 'duplicate_filename');
    if (dups.length > 0) {
      list.innerHTML = dups.map(s => `<li>${esc(s.name)}</li>`).join('');
      warn.hidden = false;
    } else {
      warn.hidden = true;
    }

    resultsSection.hidden = false;
    resultsSection.scrollIntoView({ behavior: 'smooth', block: 'start' });
  }

  function setText(id, val) {
    document.getElementById(id).textContent = val;
  }

  // ---- ダウンロード ----
  document.getElementById('dl-master').addEventListener('click', () => {
    if (!result) return;
    downloadCSV('master_counts.csv', buildMasterCSV(result.masterRows));
  });

  document.getElementById('dl-summary').addEventListener('click', () => {
    if (!result) return;
    downloadCSV('summary.csv', buildSummaryCSV(result));
  });
});
