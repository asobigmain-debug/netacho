// ===================================================
// 根多帳 GAS版
// スプレッドシート保存 / JSON・CSVインポート・エクスポート対応
// ===================================================

function doGet() {
  return HtmlService.createHtmlOutput(getHtml())
    .setTitle('根多帳')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ===== シート操作 =====

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('根多帳データ');
  if (!sheet) {
    sheet = ss.insertSheet('根多帳データ');
    sheet.appendRow(['会名', '回数', '日付', 'タイトル', '演目JSON']);
    sheet.getRange(1, 1, 1, 5).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// ===== 全データ取得（DB形式で返す） =====

function getAllData() {
  try {
    var sheet = getSheet_();
    var data = sheet.getDataRange().getValues();
    var db = {};
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[0]) continue;
      var kai = String(row[0]);
      var programs = [];
      try { programs = JSON.parse(row[4] || '[]'); } catch(e) {}
      var entry = {
        kaisu: String(row[1] || ''),
        date:  String(row[2] || ''),
        title: String(row[3] || ''),
        programs: programs
      };
      if (!db[kai]) db[kai] = [];
      db[kai].push(entry);
    }
    return JSON.stringify({ status: 'ok', data: db });
  } catch(e) { return JSON.stringify({ error: e.toString() }); }
}

// ===== 1件保存（会名・回数をキーに上書き or 追加） =====

function saveEntry(kai, kaisuStr, entryJson) {
  try {
    var p = JSON.parse(entryJson);
    var sheet = getSheet_();
    var data = sheet.getDataRange().getValues();
    var rowIdx = -1;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === kai && String(data[i][1]) === kaisuStr) {
        rowIdx = i + 1; break;
      }
    }
    var row = [kai, kaisuStr, p.date || '', p.title || '', JSON.stringify(p.programs || [])];
    if (rowIdx > 0) {
      sheet.getRange(rowIdx, 1, 1, 5).setValues([row]);
    } else {
      sheet.appendRow(row);
    }
    return 'ok';
  } catch(e) { return 'error: ' + e.toString(); }
}

// ===== 1件削除 =====

function deleteEntry(kai, kaisuStr) {
  try {
    var sheet = getSheet_();
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === kai && String(data[i][1]) === kaisuStr) {
        sheet.deleteRow(i + 1);
        return 'ok';
      }
    }
    return 'notfound';
  } catch(e) { return 'error: ' + e.toString(); }
}

// ===== JSONインポート（PWAの書き出しJSON対応） =====

function importJSON(jsonText) {
  try {
    var parsed = JSON.parse(jsonText);
    // PWA版: { "2024": [{kaisu,date,title,programs},...], "千鳥橋寄席": [...] }
    // または { version:"1.0", data: {...} }
    var incoming = parsed.data || parsed;
    var sheet = getSheet_();
    var data = sheet.getDataRange().getValues();

    // 既存インデックス作成
    var existMap = {};
    for (var i = 1; i < data.length; i++) {
      var k = data[i][0] + '|||' + data[i][1];
      existMap[k] = i + 1;
    }

    var created = 0, updated = 0;
    var keys = Object.keys(incoming);
    for (var ki = 0; ki < keys.length; ki++) {
      var kai = keys[ki];
      var entries = incoming[kai];
      if (!Array.isArray(entries)) continue;
      for (var ei = 0; ei < entries.length; ei++) {
        var entry = entries[ei];
        var kaisuStr = String(entry.kaisu || '');
        var mapKey = kai + '|||' + kaisuStr;
        var row = [kai, kaisuStr, entry.date || '', entry.title || '', JSON.stringify(entry.programs || [])];
        if (existMap[mapKey]) {
          sheet.getRange(existMap[mapKey], 1, 1, 5).setValues([row]);
          updated++;
        } else {
          sheet.appendRow(row);
          existMap[mapKey] = sheet.getLastRow();
          created++;
        }
      }
    }
    return JSON.stringify({ status: 'ok', created: created, updated: updated });
  } catch(e) { return JSON.stringify({ error: e.toString() }); }
}

// ===== JSONエクスポート =====

function exportJSON() {
  try {
    var result = JSON.parse(getAllData());
    if (result.error) return JSON.stringify({ error: result.error });
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var json = JSON.stringify({ version: '1.0', exported: new Date().toISOString(), data: result.data }, null, 2);
    var name = '根多帳_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    return JSON.stringify({ status: 'ok', json: json, name: name });
  } catch(e) { return JSON.stringify({ error: e.toString() }); }
}

// ===== CSVエクスポート =====

function exportCSV() {
  try {
    var result = JSON.parse(getAllData());
    if (result.error) return JSON.stringify({ error: result.error });
    var db = result.data;
    var lines = ['\uFEFF会名,回数,日付,タイトル,ネタ,演者']; // BOM付きUTF-8
    var keys = Object.keys(db);
    for (var ki = 0; ki < keys.length; ki++) {
      var kai = keys[ki];
      var entries = db[kai];
      for (var ei = 0; ei < entries.length; ei++) {
        var e = entries[ei];
        var programs = e.programs || [];
        if (programs.length === 0) {
          lines.push([csvEsc_(kai), csvEsc_(e.kaisu), csvEsc_(e.date), csvEsc_(e.title), '', ''].join(','));
        } else {
          for (var pi = 0; pi < programs.length; pi++) {
            var p = programs[pi];
            lines.push([
              csvEsc_(kai), csvEsc_(e.kaisu), csvEsc_(e.date), csvEsc_(e.title),
              csvEsc_(p.neta || ''), csvEsc_(p.enjya || '')
            ].join(','));
          }
        }
      }
    }
    var csv = lines.join('\n');
    var name = '根多帳_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    return JSON.stringify({ status: 'ok', csv: csv, name: name });
  } catch(e) { return JSON.stringify({ error: e.toString() }); }
}

function csvEsc_(s) {
  s = String(s || '');
  if (s.indexOf(',') >= 0 || s.indexOf('"') >= 0 || s.indexOf('\n') >= 0) {
    return '"' + s.replace(/"/g, '""') + '"';
  }
  return s;
}

// ===== CSVインポート =====

function importCSV(csvText) {
  try {
    // BOM除去
    var text = csvText.replace(/^\uFEFF/, '');
    var lines = text.split('\n');
    if (lines.length < 2) return JSON.stringify({ error: 'データが空です' });

    // ヘッダー確認
    var header = parseCsvLine_(lines[0]);
    // 期待: 会名,回数,日付,タイトル,ネタ,演者

    var entriesMap = {}; // kai|||kaisu → {kaisu,date,title,programs:[]}
    for (var i = 1; i < lines.length; i++) {
      var line = lines[i].trim();
      if (!line) continue;
      var cols = parseCsvLine_(line);
      var kai     = cols[0] || '';
      var kaisu   = cols[1] || '';
      var date    = cols[2] || '';
      var title   = cols[3] || '';
      var neta    = cols[4] || '';
      var enjya   = cols[5] || '';
      if (!kai) continue;
      var mapKey = kai + '|||' + kaisu;
      if (!entriesMap[mapKey]) {
        entriesMap[mapKey] = { kai: kai, kaisu: kaisu, date: date, title: title, programs: [] };
      }
      if (neta || enjya) {
        entriesMap[mapKey].programs.push({ neta: neta, enjya: enjya });
      }
    }

    // DB形式に変換してインポート
    var db = {};
    var mapKeys = Object.keys(entriesMap);
    for (var mk = 0; mk < mapKeys.length; mk++) {
      var item = entriesMap[mapKeys[mk]];
      if (!db[item.kai]) db[item.kai] = [];
      db[item.kai].push({ kaisu: item.kaisu, date: item.date, title: item.title, programs: item.programs });
    }
    return importJSON(JSON.stringify(db));
  } catch(e) { return JSON.stringify({ error: e.toString() }); }
}

function parseCsvLine_(line) {
  var cols = [], cur = '', inQ = false;
  for (var i = 0; i < line.length; i++) {
    var c = line[i];
    if (c === '"') {
      if (inQ && line[i+1] === '"') { cur += '"'; i++; }
      else inQ = !inQ;
    } else if (c === ',' && !inQ) {
      cols.push(cur); cur = '';
    } else {
      cur += c;
    }
  }
  cols.push(cur);
  return cols;
}

// ===== Driveバックアップ =====

function backupToDrive() {
  try {
    var result = JSON.parse(getAllData());
    if (result.error) return JSON.stringify({ error: result.error });
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var json = JSON.stringify({ version: '1.0', exported: new Date().toISOString(), data: result.data }, null, 2);
    var folderName = ss.getName() + '_バックアップ';
    var folders = DriveApp.getFoldersByName(folderName);
    var folder = folders.hasNext() ? folders.next() : DriveApp.getFileById(ss.getId()).getParents().next().createFolder(folderName);
    var date = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HH-mm');
    folder.createFile(Utilities.newBlob(json, 'application/json', '根多帳_' + date + '.json'));
    return JSON.stringify({ status: 'ok', folder: folderName });
  } catch(e) { return JSON.stringify({ error: e.toString() }); }
}

function setupAutoBackup() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'autoBackup') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('autoBackup').timeBased().everyDays(1).atHour(3).create();
  return 'ok';
}

function autoBackup() { backupToDrive(); }

// ===== HTML =====

function getHtml() {
  var h = '';
  h += '<!DOCTYPE html>\n<html lang="ja">\n<head>\n';
  h += '<meta charset="UTF-8">\n';
  h += '<meta name="viewport" content="width=device-width,initial-scale=1.0">\n';
  h += '<title>根多帳</title>\n';
  h += '<style>\n';
  h += ':root{--main:#3d4f34;--mainD:#2a3824;--mainL:#7a9a68;--mainLL:#e8f0e4;--washi:#faf8f2;--washiD:#f0ede4;--text:#1e2b1a;--textS:#7a8a74;}\n';
  h += '*{box-sizing:border-box;margin:0;padding:0;}\n';
  h += 'body{font-family:"Hiragino Mincho ProN","Yu Mincho",serif;background:var(--washi);color:var(--text);min-height:100vh;}\n';
  h += '.hd{background:var(--mainD);color:white;padding:10px 16px;display:flex;align-items:center;gap:10px;position:sticky;top:0;z-index:100;box-shadow:0 2px 10px rgba(42,56,36,.5);}\n';
  h += '.hd-bar{width:4px;height:26px;background:rgba(255,255,255,.45);border-radius:2px;}\n';
  h += '.hd-title{font-size:1.05rem;letter-spacing:.12em;font-weight:600;flex:1;}\n';
  h += '.hbtn{border:1px solid rgba(255,255,255,.3);border-radius:5px;padding:5px 10px;font-size:11px;cursor:pointer;font-family:inherit;}\n';
  h += '.hbtn-p{background:var(--main);color:white;border-color:transparent;}\n';
  h += '.hbtn-g{background:rgba(255,255,255,.12);color:white;}\n';
  h += '.hbtn-back{background:none;border:none;color:white;font-size:12px;padding:4px 8px;cursor:pointer;}\n';

  // カバー
  h += '#view-cover{background:linear-gradient(150deg,var(--mainD) 0%,var(--main) 55%,var(--mainLL) 100%);min-height:calc(100vh - 46px);padding:20px 16px 50px;}\n';
  h += '.tab-bar{display:flex;flex-wrap:wrap;gap:6px;margin-bottom:14px;}\n';
  h += '.tab-btn{background:rgba(255,255,255,.18);color:white;border:1px solid rgba(255,255,255,.3);border-radius:20px;padding:5px 14px;font-size:12px;cursor:pointer;font-family:inherit;transition:.15s;}\n';
  h += '.tab-btn.active,.tab-btn:hover{background:rgba(255,255,255,.9);color:var(--mainD);}\n';
  h += '.entry-list{display:flex;flex-direction:column;gap:7px;}\n';
  h += '.entry-card{background:rgba(255,255,255,.9);border-left:4px solid var(--mainL);border-radius:6px;padding:10px 12px;cursor:pointer;transition:.12s;}\n';
  h += '.entry-card:hover{transform:translateX(3px);box-shadow:2px 2px 8px rgba(42,56,36,.15);}\n';
  h += '.ec-top{display:flex;align-items:baseline;gap:8px;}\n';
  h += '.ec-kaisu{font-size:11px;color:var(--textS);min-width:30px;}\n';
  h += '.ec-title{font-size:14px;font-weight:600;flex:1;}\n';
  h += '.ec-date{font-size:11px;color:var(--textS);}\n';
  h += '.ec-programs{font-size:11px;color:var(--textS);margin-top:3px;line-height:1.6;}\n';
  h += '.add-btn{background:rgba(255,255,255,.2);color:white;border:1.5px dashed rgba(255,255,255,.5);border-radius:8px;padding:12px;width:100%;font-size:13px;cursor:pointer;font-family:inherit;margin-top:10px;transition:.15s;}\n';
  h += '.add-btn:hover{background:rgba(255,255,255,.35);}\n';

  // 詳細
  h += '.detail-wrap{padding:12px 12px 90px;}\n';
  h += '.detail-head{background:var(--mainLL);border-radius:8px;padding:12px;margin-bottom:10px;}\n';
  h += '.fi{width:100%;border:1px solid #ccc;border-radius:4px;padding:7px 9px;font-size:13px;font-family:inherit;color:var(--text);background:white;}\n';
  h += '.fi:focus{outline:none;border-color:var(--mainL);background:#f6faf4;}\n';
  h += '.lbl{font-size:11px;color:var(--textS);margin-bottom:3px;margin-top:8px;}\n';
  h += '.prog-sec{background:white;border:1px solid var(--mainLL);border-radius:8px;overflow:hidden;margin-bottom:8px;}\n';
  h += '.prog-header{background:var(--main);color:white;padding:7px 12px;display:flex;align-items:center;gap:8px;font-size:12px;}\n';
  h += '.prog-row{display:flex;gap:6px;padding:8px 10px;border-bottom:1px solid var(--mainLL);align-items:center;}\n';
  h += '.prog-row:last-child{border-bottom:none;}\n';
  h += '.prog-num{font-size:11px;color:var(--textS);min-width:20px;}\n';
  h += '.prog-del{background:none;border:1px solid #ddd;border-radius:4px;padding:4px 7px;font-size:11px;cursor:pointer;color:#c00;}\n';
  h += '.add-prog-btn{background:var(--mainLL);border:1.5px dashed var(--mainL);border-radius:6px;padding:8px;width:100%;font-size:12px;cursor:pointer;font-family:inherit;color:var(--mainD);margin:6px 0;}\n';
  h += '.save-btn{position:fixed;bottom:22px;right:18px;background:var(--main);color:white;border:none;border-radius:50px;padding:13px 28px;font-size:14px;font-family:inherit;box-shadow:0 4px 18px rgba(61,79,52,.5);z-index:200;display:none;cursor:pointer;}\n';
  h += '.del-btn-sm{background:none;border:1px solid #c00;border-radius:4px;padding:4px 9px;font-size:11px;cursor:pointer;color:#c00;font-family:inherit;}\n';

  // データ管理
  h += '.main{padding:14px 14px;max-width:760px;margin:0 auto;}\n';
  h += '.bk-box{background:rgba(255,255,255,.6);border:1px solid var(--mainLL);border-radius:8px;padding:14px;margin-bottom:12px;}\n';
  h += '.bk-ttl{font-size:.9rem;font-weight:600;margin-bottom:5px;}\n';
  h += '.bk-txt{font-size:.78rem;color:var(--textS);line-height:1.8;margin-bottom:10px;}\n';
  h += '.btn{padding:9px 18px;border-radius:5px;font-family:inherit;font-size:.82rem;cursor:pointer;border:none;font-weight:500;}\n';
  h += '.btn-p{background:var(--main);color:white;}\n';
  h += '.btn-s{background:transparent;color:var(--text);border:1.5px solid var(--mainLL);}\n';
  h += '.btn-sm{padding:6px 12px;font-size:.74rem;}\n';
  h += '.btn-row{display:flex;gap:8px;flex-wrap:wrap;}\n';
  h += '.upload-area{border:2px dashed var(--mainL);border-radius:7px;padding:14px;text-align:center;cursor:pointer;position:relative;margin-bottom:6px;}\n';
  h += '.upload-area input{position:absolute;inset:0;opacity:0;cursor:pointer;width:100%;height:100%;}\n';
  h += '.prog-msg{font-size:.75rem;color:var(--main);margin-top:6px;min-height:1.2em;}\n';

  h += '.view{display:none;}.view.on{display:block;}\n';
  h += '.toast{position:fixed;bottom:24px;left:50%;transform:translateX(-50%) translateY(65px);background:var(--mainD);color:white;padding:10px 20px;border-radius:20px;font-size:.82rem;z-index:999;transition:transform .3s;border-left:4px solid rgba(255,255,255,.4);white-space:nowrap;}\n';
  h += '.toast.on{transform:translateX(-50%) translateY(0);}\n';
  h += '</style>\n</head>\n<body>\n';

  // ヘッダー
  h += '<div class="hd">';
  h += '<div class="hd-bar"></div>';
  h += '<button id="btn-back" class="hbtn-back" style="display:none">◀ 戻る</button>';
  h += '<div class="hd-title" id="hd-title">根多帳</div>';
  h += '<button class="hbtn hbtn-g btn-sm" onclick="go(\'data\')">💾 データ</button>';
  h += '</div>\n';

  // カバー（会別リスト）
  h += '<div id="view-cover" class="view on">';
  h += '<div class="tab-bar" id="tab-bar"></div>';
  h += '<div class="entry-list" id="entry-list"></div>';
  h += '<button class="add-btn" id="add-entry-btn">＋ この会に新しい回を追加</button>';
  h += '</div>\n';

  // 詳細編集
  h += '<div id="view-detail" class="view">';
  h += '<div class="detail-wrap">';
  h += '<div class="detail-head">';
  h += '<div class="lbl">会名</div><input class="fi" id="d-kai">';
  h += '<div class="lbl">回数</div><input class="fi" id="d-kaisu" placeholder="例: 第10回">';
  h += '<div class="lbl">日付</div><input class="fi" id="d-date" placeholder="例: 2024年3月1日">';
  h += '<div class="lbl">タイトル</div><input class="fi" id="d-title" placeholder="タイトル（任意）">';
  h += '</div>';
  h += '<div class="prog-sec">';
  h += '<div class="prog-header"><span style="flex:1">📋 演目一覧</span><button class="hbtn hbtn-g btn-sm" onclick="addProgram()">＋ 演目追加</button></div>';
  h += '<div id="prog-list"></div>';
  h += '</div>';
  h += '<div style="text-align:right;margin-top:6px"><button class="del-btn-sm" onclick="deleteThisEntry()">🗑 この回を削除</button></div>';
  h += '</div>';
  h += '<button id="save-btn" class="save-btn" onclick="saveDetail()">保存する ✓</button>';
  h += '</div>\n';

  // データ管理
  h += '<div id="view-data" class="view main">';
  h += '<div class="bk-box"><div class="bk-ttl">📤 JSONエクスポート</div><div class="bk-txt">全データをJSONファイルとして書き出します。PWA版への取り込みに使えます。</div><button class="btn btn-p btn-sm" onclick="doExportJSON()">📤 JSON書き出し</button></div>';
  h += '<div class="bk-box"><div class="bk-ttl">📊 CSVエクスポート</div><div class="bk-txt">全データをCSVとして書き出します。Excelなどで開けます。</div><button class="btn btn-p btn-sm" onclick="doExportCSV()">📊 CSV書き出し</button></div>';
  h += '<div class="bk-box"><div class="bk-ttl">📥 インポート（JSON / CSV）</div><div class="bk-txt">PWA版から書き出したJSONまたはCSVを読み込みます。同じ会名＋回数は上書き、新しいものは追加されます。</div>';
  h += '<div class="upload-area"><input type="file" id="import-file" accept=".json,.csv" onchange="doImport(event)"><div style="font-size:.8rem;color:var(--textS)">📂 JSONまたはCSVファイルを選択</div></div>';
  h += '<div class="prog-msg" id="import-prog"></div></div>';
  h += '<div class="bk-box"><div class="bk-ttl">☁️ Googleドライブへバックアップ</div><div class="bk-txt">今すぐドライブにJSONを保存します。</div><button class="btn btn-s btn-sm" onclick="doBkDrive()">☁️ 今すぐバックアップ</button></div>';
  h += '<div class="bk-box"><div class="bk-ttl">⏰ 自動バックアップ設定</div><div class="bk-txt">毎日午前3時に自動でバックアップします。</div><button class="btn btn-s btn-sm" onclick="doSetAuto()">⏰ 自動バックアップを設定する</button></div>';
  h += '</div>\n';

  h += '<div class="toast" id="toast"></div>\n';

  // JavaScript
  h += '<script>\n';
  h += 'var DB={};\n';
  h += 'var curKai=null,curEntry=null,curIsNew=false;\n';

  // トースト
  h += 'function toast(m){var t=document.getElementById("toast");t.textContent=m;t.classList.add("on");setTimeout(function(){t.classList.remove("on");},2500);}\n';

  // 画面切替
  h += 'function go(id){\n';
  h += '  ["cover","detail","data"].forEach(function(v){document.getElementById("view-"+v).classList.remove("on");});\n';
  h += '  document.getElementById("view-"+id).classList.add("on");\n';
  h += '  document.getElementById("btn-back").style.display=(id==="cover")?"none":"inline-block";\n';
  h += '  if(id==="cover"){document.getElementById("hd-title").textContent="根多帳";buildCover();}\n';
  h += '}\n';
  h += 'document.getElementById("btn-back").onclick=function(){go("cover");};\n';

  // 会タブ一覧
  h += 'var KAI_ORDER=[];\n';
  h += 'function buildCover(){\n';
  h += '  // タブ生成（年＋定期会）\n';
  h += '  var keys=Object.keys(DB);\n';
  h += '  var years=[],others=[];\n';
  h += '  keys.forEach(function(k){if(/^[0-9]{4}$/.test(k))years.push(k);else others.push(k);});\n';
  h += '  years.sort().reverse();\n';
  h += '  others.sort();\n';
  h += '  KAI_ORDER=years.concat(others);\n';
  h += '  // 新しい会名追加用\n';
  h += '  var bar=document.getElementById("tab-bar");bar.innerHTML="";\n';
  h += '  KAI_ORDER.forEach(function(k){\n';
  h += '    var b=document.createElement("button");b.className="tab-btn"+(curKai===k?" active":"");\n';
  h += '    b.textContent=k;b.onclick=function(){curKai=k;buildCover();};\n';
  h += '    bar.appendChild(b);\n';
  h += '  });\n';
  h += '  // 新会追加ボタン\n';
  h += '  var nb=document.createElement("button");nb.className="tab-btn";nb.textContent="＋ 新しい会";\n';
  h += '  nb.onclick=function(){var name=prompt("新しい会名を入力してください（例: 令和7年 または 千鳥橋寄席）");if(!name)return;if(!DB[name])DB[name]=[];curKai=name;buildCover();};\n';
  h += '  bar.appendChild(nb);\n';
  h += '  // 現在の会を選択\n';
  h += '  if(!curKai&&KAI_ORDER.length>0)curKai=KAI_ORDER[0];\n';
  h += '  // タブアクティブ更新\n';
  h += '  bar.querySelectorAll(".tab-btn").forEach(function(b){b.classList.toggle("active",b.textContent===curKai);});\n';
  h += '  // エントリ一覧\n';
  h += '  var list=document.getElementById("entry-list");list.innerHTML="";\n';
  h += '  var entries=(DB[curKai]||[]).slice().sort(function(a,b){return String(b.kaisu).localeCompare(String(a.kaisu),"ja",{numeric:true});});\n';
  h += '  entries.forEach(function(e){\n';
  h += '    var card=document.createElement("div");card.className="entry-card";\n';
  h += '    var progsText=(e.programs||[]).map(function(p){return (p.neta||"")+(p.enjya?"（"+p.enjya+"）":"");}).join(" / ");\n';
  h += '    card.innerHTML=\'<div class="ec-top"><span class="ec-kaisu">\'+e.kaisu+\'</span><span class="ec-title">\'+esc(e.title||"（タイトルなし）")+\'</span><span class="ec-date">\'+esc(e.date)+\'</span></div>\'+(progsText?\'<div class="ec-programs">\'+esc(progsText)+\'</div>\':"");\n';
  h += '    card.onclick=function(){openEntry(curKai,e.kaisu);};\n';
  h += '    list.appendChild(card);\n';
  h += '  });\n';
  h += '  // 追加ボタン\n';
  h += '  document.getElementById("add-entry-btn").textContent="＋ "+(curKai||"この会")+"に新しい回を追加";\n';
  h += '  document.getElementById("add-entry-btn").onclick=function(){newEntry();};\n';
  h += '}\n';

  // エントリ開く
  h += 'function openEntry(kai,kaisuStr){\n';
  h += '  curKai=kai;\n';
  h += '  curIsNew=false;\n';
  h += '  var entries=DB[kai]||[];\n';
  h += '  curEntry=null;\n';
  h += '  for(var i=0;i<entries.length;i++){if(String(entries[i].kaisu)===String(kaisuStr)){curEntry=JSON.parse(JSON.stringify(entries[i]));break;}}\n';
  h += '  if(!curEntry)curEntry={kaisu:kaisuStr,date:"",title:"",programs:[]};\n';
  h += '  go("detail");\n';
  h += '  document.getElementById("hd-title").textContent=kai+" "+kaisuStr;\n';
  h += '  buildDetail();\n';
  h += '}\n';

  // 新規エントリ
  h += 'function newEntry(){\n';
  h += '  curIsNew=true;\n';
  h += '  curEntry={kaisu:"",date:"",title:"",programs:[]};\n';
  h += '  go("detail");\n';
  h += '  document.getElementById("hd-title").textContent=(curKai||"")+" 新規";\n';
  h += '  buildDetail();\n';
  h += '}\n';

  // 詳細フォーム構築
  h += 'function buildDetail(){\n';
  h += '  var e=curEntry;\n';
  h += '  document.getElementById("d-kai").value=curKai||"";\n';
  h += '  document.getElementById("d-kaisu").value=e.kaisu||"";\n';
  h += '  document.getElementById("d-date").value=e.date||"";\n';
  h += '  document.getElementById("d-title").value=e.title||"";\n';
  h += '  renderPrograms();\n';
  h += '  ["d-kai","d-kaisu","d-date","d-title"].forEach(function(id){\n';
  h += '    document.getElementById(id).oninput=function(){markDirty();};\n';
  h += '  });\n';
  h += '  document.getElementById("save-btn").style.display="block";\n';
  h += '}\n';

  // 演目レンダリング
  h += 'function renderPrograms(){\n';
  h += '  var list=document.getElementById("prog-list");\n';
  h += '  list.innerHTML="";\n';
  h += '  var progs=curEntry.programs||[];\n';
  h += '  progs.forEach(function(p,i){\n';
  h += '    var row=document.createElement("div");row.className="prog-row";\n';
  h += '    row.innerHTML=\'<span class="prog-num">\'+(i+1)+\'</span>\'\n';
  h += '      +\'<input class="fi" style="flex:2" placeholder="ネタ名" value="\'+esc(p.neta||"")+\'" data-pi="\'+i+\'" data-pk="neta">\'\n';
  h += '      +\'<input class="fi" style="flex:1" placeholder="演者" value="\'+esc(p.enjya||"")+\'" data-pi="\'+i+\'" data-pk="enjya">\'\n';
  h += '      +\'<button class="prog-del" onclick="removeProgram(\'+i+\')">×</button>\';\n';
  h += '    row.querySelectorAll("input").forEach(function(el){\n';
  h += '      el.addEventListener("input",function(){curEntry.programs[+el.dataset.pi][el.dataset.pk]=el.value;markDirty();});\n';
  h += '    });\n';
  h += '    list.appendChild(row);\n';
  h += '  });\n';
  h += '}\n';

  h += 'function addProgram(){curEntry.programs.push({neta:"",enjya:""});renderPrograms();markDirty();}\n';
  h += 'function removeProgram(i){curEntry.programs.splice(i,1);renderPrograms();markDirty();}\n';
  h += 'function markDirty(){document.getElementById("save-btn").style.display="block";}\n';

  // 保存
  h += 'function saveDetail(){\n';
  h += '  curEntry.kaisu=document.getElementById("d-kaisu").value;\n';
  h += '  curEntry.date=document.getElementById("d-date").value;\n';
  h += '  curEntry.title=document.getElementById("d-title").value;\n';
  h += '  var newKai=document.getElementById("d-kai").value;\n';
  h += '  if(!newKai){toast("会名を入力してください");return;}\n';
  h += '  if(!curEntry.kaisu){toast("回数を入力してください");return;}\n';
  h += '  // ローカルDB更新\n';
  h += '  if(newKai!==curKai){\n';
  h += '    // 会名変更 → 旧会から削除\n';
  h += '    if(DB[curKai])DB[curKai]=DB[curKai].filter(function(e){return String(e.kaisu)!==String(curEntry.kaisu);});\n';
  h += '    if(!DB[newKai])DB[newKai]=[];\n';
  h += '    curKai=newKai;\n';
  h += '  }\n';
  h += '  if(!DB[curKai])DB[curKai]=[];\n';
  h += '  var found=false;\n';
  h += '  for(var i=0;i<DB[curKai].length;i++){if(String(DB[curKai][i].kaisu)===String(curEntry.kaisu)){DB[curKai][i]=JSON.parse(JSON.stringify(curEntry));found=true;break;}}\n';
  h += '  if(!found)DB[curKai].push(JSON.parse(JSON.stringify(curEntry)));\n';
  h += '  document.getElementById("save-btn").style.display="none";\n';
  h += '  // GASに保存\n';
  h += '  google.script.run\n';
  h += '    .withSuccessHandler(function(r){if(r==="ok")toast("保存しました ✓");else toast("エラー: "+r);})\n';
  h += '    .withFailureHandler(function(e){toast("エラー: "+e.message);})\n';
  h += '    .saveEntry(curKai,String(curEntry.kaisu),JSON.stringify(curEntry));\n';
  h += '}\n';

  // 削除
  h += 'function deleteThisEntry(){\n';
  h += '  if(!confirm("この回を削除しますか？"))return;\n';
  h += '  var kai=document.getElementById("d-kai").value;\n';
  h += '  var kaisu=document.getElementById("d-kaisu").value;\n';
  h += '  if(DB[kai])DB[kai]=DB[kai].filter(function(e){return String(e.kaisu)!==String(kaisu);});\n';
  h += '  google.script.run\n';
  h += '    .withSuccessHandler(function(){toast("削除しました");go("cover");})\n';
  h += '    .withFailureHandler(function(e){toast("エラー: "+e.message);})\n';
  h += '    .deleteEntry(kai,kaisu);\n';
  h += '}\n';

  // エスケープ
  h += 'function esc(s){return String(s).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;");}\n';

  // データ管理
  h += 'function doExportJSON(){google.script.run.withSuccessHandler(function(res){var r=JSON.parse(res);if(r.error){toast("エラー: "+r.error);return;}var a=document.createElement("a");a.href="data:application/json;charset=utf-8,"+encodeURIComponent(r.json);a.download=r.name+".json";a.click();toast("✓ JSONを書き出しました");}).exportJSON();}\n';
  h += 'function doExportCSV(){google.script.run.withSuccessHandler(function(res){var r=JSON.parse(res);if(r.error){toast("エラー: "+r.error);return;}var a=document.createElement("a");a.href="data:text/csv;charset=utf-8,"+encodeURIComponent(r.csv);a.download=r.name+".csv";a.click();toast("✓ CSVを書き出しました");}).exportCSV();}\n';

  h += 'function doImport(ev){\n';
  h += '  var f=ev.target.files[0];if(!f)return;\n';
  h += '  var prog=document.getElementById("import-prog");prog.textContent="読み込み中…";\n';
  h += '  var r=new FileReader();\n';
  h += '  r.onload=function(e){\n';
  h += '    var isCSV=f.name.toLowerCase().endsWith(".csv");\n';
  h += '    var fn=isCSV?"importCSV":"importJSON";\n';
  h += '    google.script.run\n';
  h += '      .withSuccessHandler(function(res){var d=JSON.parse(res);if(d.error){prog.textContent="エラー: "+d.error;return;}prog.textContent="✓ 完了：新規"+d.created+"件・更新"+d.updated+"件";toast("✓ インポート完了");\n';
  h += '        google.script.run.withSuccessHandler(function(res2){var r2=JSON.parse(res2);if(!r2.error){DB=r2.data;buildCover();}}).getAllData();})\n';
  h += '      .withFailureHandler(function(e){prog.textContent="エラー: "+e.message;})\n';
  h += '      [fn](e.target.result);\n';
  h += '  };\n';
  h += '  r.readAsText(f,"UTF-8");\n';
  h += '  ev.target.value="";\n';
  h += '}\n';

  h += 'function doBkDrive(){toast("バックアップ中…");google.script.run.withSuccessHandler(function(res){var r=JSON.parse(res);if(r.error){toast("エラー: "+r.error);return;}toast("✓ "+r.folder+"に保存しました");}).backupToDrive();}\n';
  h += 'function doSetAuto(){google.script.run.withSuccessHandler(function(){toast("✓ 自動バックアップを設定しました（毎日午前3時）");}).setupAutoBackup();}\n';

  // 初期ロード
  h += 'google.script.run.withSuccessHandler(function(res){\n';
  h += '  var r=JSON.parse(res);\n';
  h += '  if(!r.error)DB=r.data;\n';
  h += '  buildCover();\n';
  h += '}).getAllData();\n';

  h += '</script>\n</body>\n</html>';
  return h;
}
