/**
 * 西村OS 受付窓口（Google Apps Script）
 *
 * 「西村家ファイナンスDB_v2」スプシに貼り付けて「ウェブアプリ」として公開する。
 * これ1つで以下を担う:
 *   - 読み出し(GET): 西村OSが全シートを1回で取得（14個のCSV公開URLが不要になる）
 *   - 書き込み(POST): 体重・メモ・写真をスプシ/ドライブに直接記録（Claudeとの会話は不要）
 *   - 写真の読み取り: 食事・体重計の写真を1枚ずつ使い捨てでClaude APIに読ませる
 *
 * 設定（プロジェクトの設定 → スクリプト プロパティ）:
 *   KEY               … 合言葉。西村OSとiPhoneショートカットに同じものを入れる（必須）
 *   ANTHROPIC_API_KEY … 任意。入れるとその場でClaude APIが写真を読む（従量課金）。
 *                       入れない場合は「未読取」で受信箱に溜まり、クラウドの定期実行（Maxプランの範囲）が
 *                       ?pending=1 で受け取り、type:'photoDone' で結果を書き戻す
 *
 * 手順は gas/README.md を参照。
 */

/* 西村OSのデータ名 → スプシのシート名。シート名が違う場合はここだけ直す */
const SHEETS = {
  snap: '資産スナップショット',
  cf: '月次収支',
  port: 'ポートフォリオ',
  loan: '負債',
  meisai: '購入明細',
  report: 'レポート',
  hist: '配当履歴',
  fixed: '固定費',
  income: '収入明細',
  watch: 'ウォッチリスト',
  wish: '欲しいものリスト',
  meal: '食事ログ',
  weight: '体重ログ',
  kabuhist: '株価履歴',
  inbox: '受信箱',
  presets: 'いつもの',
  trades: '取引履歴'
};

/* 無ければ自動で作るシートと見出し行 */
const HEADERS = {
  '体重ログ': ['日付', '体重kg', '体脂肪率%', 'メモ', '筋肉量kg', '体脂肪量kg', '基礎代謝kcal', '水分量%', 'BMI', '心拍bpm', '測定時刻'],
  '食事ログ': ['日付', '時刻', '区分', '内容', '推定カロリー', 'P', 'F', 'C', '根拠', 'メモ'],
  '受信箱': ['受信日時', '種類', '内容', '写真', '状態'],
  'いつもの': ['名前', 'kcal', 'P', 'F', 'C']
};

const PHOTO_FOLDER = '西村OS写真';
const TZ = 'Asia/Tokyo';
const CLAUDE_MODEL = 'claude-opus-5';

/* ================= 入口 ================= */

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (!checkKey_(p.key)) return json_({ ok: false, error: '合言葉が違います' });
  if (p.ping) return json_({ ok: true, sheets: listSheets_() });
  if (p.pending) return json_(pendingPhotos_(parseInt(p.limit, 10) || 5));
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const data = {};
  Object.keys(SHEETS).forEach(function (k) {
    const sh = ss.getSheetByName(SHEETS[k]);
    data[k] = sh ? sh.getDataRange().getDisplayValues() : null;
  });
  return json_({ ok: true, updated: now_(), data: data });
}

function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'JSONが読めません' }); }
  if (!checkKey_(body.key)) return json_({ ok: false, error: '合言葉が違います' });
  try {
    switch (body.type) {
      case 'weight': return json_(addWeight_(body));
      case 'memo': return json_(addInbox_('メモ', body.text || '', ''));
      case 'meal': return json_(addMeal_(body));
      case 'preset': return json_(addPreset_(body));
      case 'trade': return json_(addInbox_('売買', body.text || '', '', '未読取'));
      case 'sheetOps': return json_(sheetOps_(body));
      case 'radar': return json_(addInbox_('レーダー', (body.title ? '【' + body.title + '】' : '') + (body.text || ''), '', '処理済'));
      case 'photo': return json_(addPhoto_(body));
      case 'photoDone': return json_(photoDone_(body));
      default: return json_({ ok: false, error: '不明な種類: ' + body.type });
    }
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

/* ================= 体重 ================= */

function addWeight_(b) {
  const kg = parseFloat(b.kg);
  if (!(kg > 20 && kg < 300)) return { ok: false, error: '体重の数字が読めません: ' + b.kg };
  const date = normDate_(b.date) || today_();
  const sh = sheet_('体重ログ');
  // 同じ日付・同じ体重がすでにあれば二重登録しない（ショートカットの重複実行対策）
  const rows = sh.getDataRange().getDisplayValues();
  for (let i = 1; i < rows.length; i++) {
    if (normDate_(rows[i][0]) === date && Math.abs(parseFloat(rows[i][1]) - kg) < 0.05) {
      return { ok: true, dup: true, message: date + ' ' + kg.toFixed(1) + 'kg は記録済みです' };
    }
  }
  const fat = num_(b.fat);
  sh.appendRow([date, kg, fat === null ? '' : fat, b.memo || b.source || '', '', '', '', '', num_(b.bmi) === null ? '' : num_(b.bmi), '', b.time || '']);
  return { ok: true, message: date + ' ' + kg.toFixed(1) + 'kg を記録しました' };
}

/* ================= 食事（テキスト・いつもの） ================= */

/* {items, meal_type?, date?, kcal?, p?, f?, c?}
   カロリーが分かっている（いつもの）→ 食事ログに即記録。分からない → 受信箱に「未読取」で置き、定期実行が見積もる */
function addMeal_(b) {
  const items = String(b.items || '').trim();
  if (!items) return { ok: false, error: '食べたものが空です' };
  const d = new Date();
  const date = normDate_(b.date) || today_();
  const type = b.meal_type || mealType_(d);
  // 過去に同じ内容を記録していれば、その値でその場で記録（写真から読み取った分も含む）
  if (num_(b.kcal) === null) {
    const past = findPastMeal_(items);
    if (past) { b.kcal = past.kcal; b.p = past.p; b.f = past.f; b.c = past.c; b.basis = '過去の記録から'; }
  }
  if (num_(b.kcal) !== null) {
    sheet_('食事ログ').appendRow([date, Utilities.formatDate(d, TZ, 'HH:mm'), type, items, num_(b.kcal), num_(b.p), num_(b.f), num_(b.c), b.basis || 'いつもの', '']);
    return { ok: true, message: items + ' ' + Math.round(num_(b.kcal)) + 'kcal を記録しました' };
  }
  sheet_('受信箱').appendRow([stamp_(b.date), '食事(テキスト)', '[' + type + '] ' + items, '', '未読取']);
  return { ok: true, message: '記録しました。カロリーは次の自動読み取り（朝・昼・夜）で入ります' };
}

function addPreset_(b) {
  const name = String(b.name || '').trim();
  if (!name || num_(b.kcal) === null) return { ok: false, error: '名前とkcalを入れてください' };
  const sh = sheet_('いつもの');
  const rows = sh.getDataRange().getDisplayValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === name) { sh.getRange(i + 1, 2, 1, 4).setValues([[num_(b.kcal), num_(b.p), num_(b.f), num_(b.c)]]); return { ok: true, message: '「' + name + '」を更新しました' }; }
  }
  sh.appendRow([name, num_(b.kcal), num_(b.p), num_(b.f), num_(b.c)]);
  return { ok: true, message: '「' + name + '」をいつものに登録しました' };
}

/* ================= メモ・受信箱 ================= */

function addInbox_(kind, text, photoUrl, status) {
  if (!text && !photoUrl) return { ok: false, error: '中身が空です' };
  sheet_('受信箱').appendRow([now_(), kind, text, photoUrl, status || '未処理']);
  if (status === '未読取') return { ok: true, message: '記録しました。次の自動読み取り（朝・昼・夜）で反映されます' };
  return { ok: true, message: '受信箱に入れました' };
}

/* ================= 写真 ================= */

function addPhoto_(b) {
  if (!b.data) return { ok: false, error: '写真データがありません' };
  const mime = b.mime || 'image/jpeg';
  const kind = b.kind || 'other'; // meal / scale / other
  const blob = Utilities.newBlob(Utilities.base64Decode(b.data), mime,
    Utilities.formatDate(new Date(), TZ, 'yyyyMMdd-HHmmss') + '-' + kind + '.jpg');
  const file = photoFolder_().createFile(blob);
  const url = file.getUrl();

  const apiKey = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  if (kind === 'other') {
    addInbox_('写真(その他)', b.note || '', url);
    return { ok: true, message: '写真を受信箱に保存しました' };
  }
  if (!apiKey || kind === 'asset') {
    // APIキー無し（資産スクショは常にこちら） → 定期実行（Maxプラン）が後で読む
    sheet_('受信箱').appendRow([stamp_(b.date), '写真(' + kindLabel_(kind) + ')', b.note || '', url, '未読取']);
    return { ok: true, message: '写真を保存しました。次の自動読み取り（朝・昼・夜）で反映されます' };
  }

  const r = readPhoto_(apiKey, kind, b.data, mime, b.note || '');
  if (!r) {
    addInbox_('写真(' + kindLabel_(kind) + ')・読み取り失敗', b.note || '', url);
    return { ok: true, message: '写真は保存しました（読み取れなかったので受信箱へ）' };
  }
  if (kind === 'scale') {
    return addWeight_({ kg: r.weight_kg, fat: r.body_fat_pct, bmi: r.bmi, date: b.date, memo: '写真から' });
  }
  // meal
  const d = new Date();
  sheet_('食事ログ').appendRow([
    normDate_(b.date) || today_(), Utilities.formatDate(d, TZ, 'HH:mm'), r.meal_type || mealType_(d),
    r.items, r.kcal, r.protein_g, r.fat_g, r.carb_g, r.basis, (b.note ? b.note + ' ' : '') + url
  ]);
  return { ok: true, message: r.items + ' 約' + Math.round(r.kcal) + 'kcal を記録しました' };
}

/* 写真1枚を使い捨ての1回のリクエストで読む（会話履歴を持たないので、何日続けても1回の費用は同じ） */
function readPhoto_(apiKey, kind, b64, mime, note) {
  const schema = kind === 'scale' ? {
    type: 'object', additionalProperties: false,
    required: ['weight_kg', 'body_fat_pct', 'bmi'],
    properties: {
      weight_kg: { type: 'number' },
      body_fat_pct: { anyOf: [{ type: 'number' }, { type: 'null' }] },
      bmi: { anyOf: [{ type: 'number' }, { type: 'null' }] }
    }
  } : {
    type: 'object', additionalProperties: false,
    required: ['meal_type', 'items', 'kcal', 'protein_g', 'fat_g', 'carb_g', 'basis'],
    properties: {
      meal_type: { type: 'string', enum: ['朝食', '昼食', '夕食', '間食'] },
      items: { type: 'string' },
      kcal: { type: 'number' },
      protein_g: { type: 'number' },
      fat_g: { type: 'number' },
      carb_g: { type: 'number' },
      basis: { type: 'string' }
    }
  };
  const prompt = kind === 'scale'
    ? '体重計またはヘルスケアアプリの画面写真です。表示されている体重(kg)、体脂肪率(%)、BMIを読み取ってください。表示が無い項目はnull。'
    : '食事の写真です。写っている料理名（複数なら「、」区切り）と、全体の推定カロリー・たんぱく質・脂質・炭水化物(g)を出してください。栄養成分表示が写っていればそれを優先。basisには推定の根拠を一文で。' +
      (note ? '\n補足: ' + note : '');

  const res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post',
    contentType: 'application/json',
    muteHttpExceptions: true,
    headers: {
      'x-api-key': apiKey,
      'anthropic-version': '2023-06-01',
      'anthropic-beta': 'server-side-fallback-2026-07-01'
    },
    payload: JSON.stringify({
      model: CLAUDE_MODEL,
      fallbacks: 'default',
      max_tokens: 2000,
      output_config: { effort: 'low', format: { type: 'json_schema', schema: schema } },
      messages: [{
        role: 'user',
        content: [
          { type: 'image', source: { type: 'base64', media_type: mime, data: b64 } },
          { type: 'text', text: prompt }
        ]
      }]
    })
  });
  if (res.getResponseCode() !== 200) { console.error(res.getContentText()); return null; }
  const msg = JSON.parse(res.getContentText());
  if (msg.stop_reason === 'refusal') return null;
  const text = (msg.content || []).filter(function (c) { return c.type === 'text'; }).map(function (c) { return c.text; }).join('');
  try { return JSON.parse(text); } catch (err) { console.error(text); return null; }
}

/* ================= 小道具 ================= */

function checkKey_(k) {
  const key = PropertiesService.getScriptProperties().getProperty('KEY');
  return !!key && k === key;
}
function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function sheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    if (HEADERS[name]) sh.appendRow(HEADERS[name]);
  }
  return sh;
}
function photoFolder_() {
  const it = DriveApp.getFoldersByName(PHOTO_FOLDER);
  return it.hasNext() ? it.next() : DriveApp.createFolder(PHOTO_FOLDER);
}
function listSheets_() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheets().map(function (s) { return s.getName(); });
}
/* 受信日時。食べた日を「昨日」にした場合はその日付＋今の時刻（読み取り時もこの日付で記録される） */
function stamp_(date) {
  const d = normDate_(date);
  return d ? d + ' ' + Utilities.formatDate(new Date(), TZ, 'HH:mm') : now_();
}
function findPastMeal_(items) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('食事ログ');
  if (!sh) return null;
  const rows = sh.getDataRange().getDisplayValues();
  const key = String(items).replace(/\s/g, '');
  for (let i = rows.length - 1; i >= 1; i--) {
    if (String(rows[i][3]).replace(/\s/g, '') === key && num_(rows[i][4]) !== null) {
      return { kcal: num_(rows[i][4]), p: num_(rows[i][5]), f: num_(rows[i][6]), c: num_(rows[i][7]) };
    }
  }
  return null;
}
function now_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm'); }
function today_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }
function normDate_(s) {
  if (!s) return '';
  const m = String(s).match(/(\d{4})[-\/年.](\d{1,2})[-\/月.](\d{1,2})/);
  return m ? m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2) : '';
}
function num_(v) {
  if (v === undefined || v === null || v === '') return null;
  const n = parseFloat(String(v).replace(/[^\d.\-]/g, ''));
  return isNaN(n) ? null : n;
}
function kindLabel_(k) { return { meal: '食事', scale: '体重計', asset: '資産', other: 'その他' }[k] || k; }
function labelKind_(label) {
  const m = String(label).match(/写真\((.+?)\)/);
  const map = { '食事': 'meal', '体重計': 'scale', '資産': 'asset' };
  return m ? (map[m[1]] || 'other') : 'other';
}

/* ================= 定期実行（Maxプラン）向け：読み取り待ちの写真 ================= */

function pendingPhotos_(limit) {
  const sh = sheet_('受信箱');
  const rows = sh.getDataRange().getDisplayValues();
  const items = [];
  for (let i = 1; i < rows.length && items.length < limit; i++) {
    if (rows[i][4] !== '未読取') continue;
    if (rows[i][1] === '売買') {
      items.push({ row: i + 1, kind: 'trade', note: rows[i][2], received: rows[i][0] });
      continue;
    }
    if (rows[i][1] === '食事(テキスト)') {
      items.push({ row: i + 1, kind: 'mealText', note: rows[i][2], received: rows[i][0] });
      continue;
    }
    const id = (String(rows[i][3]).match(/[-\w]{25,}/) || [])[0];
    if (!id) continue;
    const blob = DriveApp.getFileById(id).getBlob();
    items.push({
      row: i + 1, kind: labelKind_(rows[i][1]), note: rows[i][2], received: rows[i][0],
      mime: blob.getContentType(), data: Utilities.base64Encode(blob.getBytes())
    });
  }
  let left = 0;
  for (let i = 1; i < rows.length; i++) if (rows[i][4] === '未読取') left++;
  const out = { ok: true, items: items, left: left, snapHeader: headerOf_('snap') };
  if (items.some(function (x) { return x.kind === 'trade' || x.kind === 'asset'; })) {
    out.port = sheetValues_('port');      // 保有銘柄（見出し行つき）
    out.tradesHeader = headerOf_('trades');
  }
  return out;
}

/* 読み取り結果の書き戻し。{row, meal:{...}} / {row, weight:{kg,fat,bmi}} / {row, snapRow:[...]} / {row, text:'...'} */
function photoDone_(b) {
  const sh = sheet_('受信箱');
  const row = parseInt(b.row, 10);
  if (!(row > 1) || sh.getRange(row, 5).getValue() !== '未読取') return { ok: false, error: 'その行は読み取り待ちではありません: ' + b.row };
  const received = String(sh.getRange(row, 1).getDisplayValue());
  const date = normDate_(received) || today_();
  let msg = '';
  if (b.meal) {
    const m = b.meal;
    sheet_('食事ログ').appendRow([date, (received.match(/\d{1,2}:\d{2}/) || [''])[0], m.meal_type || '', m.items || '', num_(m.kcal), num_(m.protein_g), num_(m.fat_g), num_(m.carb_g), m.basis || '', sh.getRange(row, 4).getDisplayValue() || 'テキストから']);
    msg = '食事ログに記録';
  } else if (b.weight) {
    const r = addWeight_({ kg: b.weight.kg, fat: b.weight.fat, bmi: b.weight.bmi, date: b.weight.date || date, memo: '写真から' });
    if (!r.ok) return r;
    msg = r.message;
  } else if (b.snapRow) {
    const name = SHEETS.snap;
    const target = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
    if (!target) return { ok: false, error: name + ' シートがありません' };
    target.appendRow(b.snapRow);
    msg = name + 'に1行追加';
  } else if (!b.text) {
    return { ok: false, error: '結果がありません' };
  }
  sh.getRange(row, 5).setValue('処理済');
  if (b.text) sh.getRange(row, 3).setValue((sh.getRange(row, 3).getValue() ? sh.getRange(row, 3).getValue() + ' / ' : '') + b.text);
  return { ok: true, message: msg || '受信箱にメモしました' };
}

function sheetValues_(key) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS[key]);
  return sh ? sh.getDataRange().getDisplayValues() : null;
}

/* 定期実行（売買・配当レーダー）からのスプシ更新。書けるのはポートフォリオ・取引履歴・資産スナップショットだけ。
   {row?, ops:[{op:'append', sheet:'trades'|'port'|'snap', values:[...]},
               {op:'set', sheet:'port', matchCol:0, matchValue:'7203', set:{"3":"200","4":"2500"}}], text?}
   行の削除はしない（売り切ったら株数を0にする） */
function sheetOps_(b) {
  const allowed = { trades: 1, port: 1, snap: 1 };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const done = [];
  (b.ops || []).forEach(function (o) {
    if (!allowed[o.sheet]) throw new Error('書き込めないシート: ' + o.sheet);
    const sh = ss.getSheetByName(SHEETS[o.sheet]);
    if (!sh) throw new Error(SHEETS[o.sheet] + ' シートがありません');
    if (o.op === 'append') {
      sh.appendRow(o.values);
      done.push(SHEETS[o.sheet] + 'に1行追加');
    } else if (o.op === 'set') {
      const v = sh.getDataRange().getDisplayValues();
      let r = -1;
      for (let i = 1; i < v.length; i++) if (String(v[i][o.matchCol]).trim() === String(o.matchValue).trim()) { r = i; break; }
      if (r < 0) throw new Error(SHEETS[o.sheet] + 'に ' + o.matchValue + ' が見つかりません');
      Object.keys(o.set || {}).forEach(function (c) { sh.getRange(r + 1, parseInt(c, 10) + 1).setValue(o.set[c]); });
      done.push(SHEETS[o.sheet] + ' ' + o.matchValue + ' を更新');
    } else {
      throw new Error('不明な操作: ' + o.op);
    }
  });
  if (b.row) {
    const ib = sheet_('受信箱');
    ib.getRange(parseInt(b.row, 10), 5).setValue('処理済');
    if (b.text) ib.getRange(parseInt(b.row, 10), 3).setValue(ib.getRange(parseInt(b.row, 10), 3).getValue() + ' / ' + b.text);
  }
  return { ok: true, message: done.join('、') || '更新なし' };
}

function headerOf_(key) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS[key]);
  if (!sh) return null;
  const v = sh.getDataRange().getDisplayValues();
  return { header: v[0], last: v.length > 1 ? v[v.length - 1] : null };
}
function mealType_(d) {
  const h = parseInt(Utilities.formatDate(d, TZ, 'H'), 10);
  return h < 10 ? '朝食' : h < 15 ? '昼食' : h < 21 ? '夕食' : '間食';
}

/* 動作確認用: エディタで実行すると、合言葉とシートの見つかり具合をログに出す */
function selfCheck() {
  const props = PropertiesService.getScriptProperties();
  console.log('KEY 設定: ' + (props.getProperty('KEY') ? 'OK' : '未設定（スクリプトプロパティに KEY を追加）'));
  console.log('ANTHROPIC_API_KEY: ' + (props.getProperty('ANTHROPIC_API_KEY') ? 'OK' : '未設定（写真の自動読み取りは無効）'));
  const names = listSheets_();
  Object.keys(SHEETS).forEach(function (k) {
    console.log((names.indexOf(SHEETS[k]) >= 0 ? '✅ ' : '— ') + SHEETS[k]);
  });
}
