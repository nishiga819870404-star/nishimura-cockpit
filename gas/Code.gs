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
 *   ANTHROPIC_API_KEY … 写真の自動読み取りを使う場合のみ（任意。無ければ写真は保存だけ）
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
  inbox: '受信箱'
};

/* 無ければ自動で作るシートと見出し行 */
const HEADERS = {
  '体重ログ': ['日付', '体重kg', '体脂肪率%', 'メモ', '筋肉量kg', '体脂肪量kg', '基礎代謝kcal', '水分量%', 'BMI', '心拍bpm', '測定時刻'],
  '食事ログ': ['日付', '時刻', '区分', '内容', '推定カロリー', 'P', 'F', 'C', '根拠', 'メモ'],
  '受信箱': ['受信日時', '種類', '内容', '写真', '状態']
};

const PHOTO_FOLDER = '西村OS写真';
const TZ = 'Asia/Tokyo';
const CLAUDE_MODEL = 'claude-opus-5';

/* ================= 入口 ================= */

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (!checkKey_(p.key)) return json_({ ok: false, error: '合言葉が違います' });
  if (p.ping) return json_({ ok: true, sheets: listSheets_() });
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
      case 'photo': return json_(addPhoto_(body));
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

/* ================= メモ・受信箱 ================= */

function addInbox_(kind, text, photoUrl) {
  if (!text && !photoUrl) return { ok: false, error: '中身が空です' };
  sheet_('受信箱').appendRow([now_(), kind, text, photoUrl, '未処理']);
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
  if (!apiKey || kind === 'other') {
    addInbox_('写真(' + kindLabel_(kind) + ')', b.note || '', url);
    return { ok: true, message: apiKey ? '写真を受信箱に保存しました' : '写真を保存しました（自動読み取りは未設定）' };
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
function kindLabel_(k) { return { meal: '食事', scale: '体重計', other: 'その他' }[k] || k; }
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
