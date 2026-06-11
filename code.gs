/********************************************************
 * みんあそくん バックエンド (Google Apps Script)
 *
 * 【セットアップ手順】
 * 1. https://sheets.google.com で新しいスプレッドシートを作成
 *    （名前は「みんあそくん会員DB」など自由）
 * 2. メニュー「拡張機能」→「Apps Script」を開く
 * 3. このコードを全て貼り付けて保存（Ctrl+S）
 * 4. 右上「デプロイ」→「新しいデプロイ」
 *    - 種類：「ウェブアプリ」
 *    - 次のユーザーとして実行：「自分」
 *    - アクセスできるユーザー：「全員」 ←重要
 * 5. 「デプロイ」→ 表示された「ウェブアプリのURL」をコピー
 *    （https://script.google.com/macros/s/XXXX/exec の形式）
 * 6. index.html の GAS_API_URL に貼り付ける
 *
 * シート（Members/Ideas/Events/Help）は初回アクセス時に
 * 自動作成されます。手動で作る必要はありません。
 ********************************************************/

/* ─── セキュリティ設定 ───────────────────────────
   管理者番号は「プロジェクトの設定」→「スクリプト プロパティ」で
   ADMIN_NO という名前で登録すると、コードを変えずに変更できます。
   未登録の場合は下の初期値が使われます。
   ※ index.html には管理者番号を一切書かないこと。 */
function adminNo(){
  return PropertiesService.getScriptProperties().getProperty('ADMIN_NO') || '19626291';
}
function isAdminNo(no){ return String(no) === adminNo(); }

// 入力値の長さ制限（荒らし・シート破壊対策）
function clip(s, n){ return String(s == null ? '' : s).slice(0, n); }
const LIMITS = {nick:20, bio:200, title:80, desc:500, place:80, date:40,
                era:30, comment:500, cat:20, type:10};
function clipReq(req){
  Object.keys(LIMITS).forEach(k=>{ if(k in req) req[k] = clip(req[k], LIMITS[k]); });
  if(Array.isArray(req.interests)) req.interests = req.interests.slice(0,15).map(x=>clip(x,20));
  if(Array.isArray(req.tags))      req.tags      = req.tags.slice(0,10).map(x=>clip(x,20));
  return req;
}

// ─── シート定義 ───────────────────────────────
const SHEETS = {
  Members: ['memberNo','nick','interests','bio','created'],
  Ideas:   ['id','memberNo','nick','title','cat','desc','place','interested','created'],
  Events:  ['id','memberNo','nick','title','date','place','desc','rsvp','created'],
  Help:    ['id','memberNo','nick','type','title','tags','created'],
  Memories:['id','memberNo','nick','place','era','comment','created']
};

function getSheet(name){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if(!sh){
    sh = ss.insertSheet(name);
    sh.appendRow(SHEETS[name]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function readAll(name){
  const sh = getSheet(name);
  const values = sh.getDataRange().getValues();
  const headers = values.shift();
  return values.map(row=>{
    const obj = {};
    headers.forEach((h,i)=>obj[h]=row[i]);
    return obj;
  });
}

function json(o){
  return ContentService.createTextOutput(JSON.stringify(o))
    .setMimeType(ContentService.MimeType.JSON);
}

// ─── エントリポイント ─────────────────────────
function doGet(){ return json({ok:true, service:'minna-aso'}); }

function doPost(e){
  let req;
  try{ req = JSON.parse(e.postData.contents); }
  catch(err){ return json({ok:false, error:'bad request'}); }

  clipReq(req);
  try{
    switch(req.action){
      case 'register':      return register(req);
      case 'login':         return login(req);
      case 'updateProfile': return updateProfile(req);
      case 'listAll':       return listAll(req);
      case 'addIdea':       return addItem('Ideas', req, ['title','cat','desc','place'], {interested:0});
      case 'addEvent':      return addItem('Events', req, ['title','date','place','desc'], {rsvp:0});
      case 'addHelp':       return addItem('Help', req, ['type','title','tags'], {});
      case 'addMemory':     return addItem('Memories', req, ['place','era','comment'], {});
      case 'bump':          return bump(req);
      case 'deleteItem':    return deleteItem(req);
      case 'editItem':      return editItem(req);
      case 'adminDelete':   return adminDelete(req);
      default: return json({ok:false, error:'unknown action'});
    }
  }catch(err){
    return json({ok:false, error:String(err)});
  }
}

// ─── 会員登録：8桁番号を発行 ───────────────────
function register(req){
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try{
    const sh = getSheet('Members');
    const existing = readAll('Members').map(m=>String(m.memberNo));
    let no;
    do{
      no = String(Math.floor(10000000 + Math.random()*90000000)); // 8桁
    }while(existing.includes(no) || isAdminNo(no));

    sh.appendRow([no, req.nick||'', JSON.stringify(req.interests||[]),
                  req.bio||'', new Date().toISOString()]);
    return json({ok:true, memberNo:no,
      profile:{memberNo:no, nick:req.nick||'', interests:req.interests||[], bio:req.bio||''}});
  }finally{ lock.releaseLock(); }
}

// ─── ログイン ──────────────────────────────────
function login(req){
  const no = String(req.memberNo||'').trim();
  if(!/^\d{8}$/.test(no)) return json({ok:false, error:'会員番号は8桁の数字です'});

  const members = readAll('Members');
  let m = members.find(x=>String(x.memberNo)===no);

  // 管理者番号は初回ログイン時に自動作成
  if(!m && isAdminNo(no)){
    getSheet('Members').appendRow([adminNo(),'管理者','[]','',new Date().toISOString()]);
    m = {memberNo:adminNo(), nick:'管理者', interests:'[]', bio:''};
  }
  if(!m) return json({ok:false, error:'会員番号が見つかりません'});

  return json({ok:true, profile:{
    memberNo:String(m.memberNo), nick:m.nick,
    interests: safeParse(m.interests), bio:m.bio,
    isAdmin: isAdminNo(m.memberNo)
  }});
}

function safeParse(s){ try{ return JSON.parse(s)||[]; }catch(e){ return []; } }

// ─── プロフィール更新 ─────────────────────────
function updateProfile(req){
  const sh = getSheet('Members');
  const data = sh.getDataRange().getValues();
  for(let i=1;i<data.length;i++){
    if(String(data[i][0])===String(req.memberNo)){
      sh.getRange(i+1,2).setValue(req.nick||'');
      sh.getRange(i+1,3).setValue(JSON.stringify(req.interests||[]));
      sh.getRange(i+1,4).setValue(req.bio||'');
      return json({ok:true});
    }
  }
  return json({ok:false, error:'member not found'});
}

// ─── 全データ取得 ─────────────────────────────
// セキュリティ：会員番号はログインの鍵なので、他人の番号は
// 絶対にそのまま返さない（管理者にのみ全番号を返す）。
// 本人判定は isMe / mine フラグで返す。
function listAll(req){
  const auth = verifyMember(req.memberNo);
  if(!auth.ok) return json(auth);
  const me = String(req.memberNo);
  const admin = isAdminNo(me);
  const mask = no => admin ? String(no) : '*****' + String(no).slice(-3);

  const members = readAll('Members').map(m=>({
    memberNo: mask(m.memberNo),
    isMe: String(m.memberNo)===me,
    nick:m.nick, interests:safeParse(m.interests), bio:m.bio, created:m.created
  }));
  const strip = x => { const o={...x, mine:String(x.memberNo)===me}; delete o.memberNo; return o; };
  const ideas  = readAll('Ideas').map(strip);
  const events = readAll('Events').map(strip);
  const help   = readAll('Help').map(x=>{ const o=strip(x); o.tags=safeParse(x.tags); return o; });
  const memories = readAll('Memories').map(strip);
  return json({ok:true, members, ideas, events, help, memories, isAdmin:admin});
}

function verifyMember(no){
  const m = readAll('Members').find(x=>String(x.memberNo)===String(no));
  return m ? {ok:true} : {ok:false, error:'ログインしてください'};
}

// ─── 投稿追加（汎用） ─────────────────────────
function addItem(sheetName, req, fields, extra){
  const auth = verifyMember(req.memberNo);
  if(!auth.ok) return json(auth);
  const sh = getSheet(sheetName);
  const id = Utilities.getUuid();
  const headers = SHEETS[sheetName];
  const row = headers.map(h=>{
    if(h==='id') return id;
    if(h==='memberNo') return String(req.memberNo);
    if(h==='nick') return req.nick||'';
    if(h==='created') return new Date().toISOString();
    if(h==='tags') return JSON.stringify(req.tags||[]);
    if(h in extra) return extra[h];
    return req[h]||'';
  });
  sh.appendRow(row);
  return json({ok:true, id});
}

// ─── 興味あり・参加カウント +1 ─────────────────
function bump(req){
  const auth = verifyMember(req.memberNo);
  if(!auth.ok) return json(auth);
  const map = {ideas:{sheet:'Ideas',col:'interested'}, events:{sheet:'Events',col:'rsvp'}};
  const t = map[req.type]; if(!t) return json({ok:false,error:'bad type'});
  const sh = getSheet(t.sheet);
  const data = sh.getDataRange().getValues();
  const headers = data[0];
  const idCol = headers.indexOf('id'), cntCol = headers.indexOf(t.col);
  for(let i=1;i<data.length;i++){
    if(data[i][idCol]===req.id){
      sh.getRange(i+1, cntCol+1).setValue(Number(data[i][cntCol]||0)+1);
      return json({ok:true});
    }
  }
  return json({ok:false, error:'not found'});
}

// ─── 削除（本人 or 管理者） ───────────────────
function deleteItem(req){
  const isAdmin = isAdminNo(req.memberNo);
  const sheetName = {ideas:'Ideas',events:'Events',help:'Help',memories:'Memories'}[req.type];
  if(!sheetName) return json({ok:false,error:'bad type'});
  const sh = getSheet(sheetName);
  const data = sh.getDataRange().getValues();
  const headers = data[0];
  const idCol = headers.indexOf('id'), noCol = headers.indexOf('memberNo');
  for(let i=1;i<data.length;i++){
    if(data[i][idCol]===req.id){
      if(!isAdmin && String(data[i][noCol])!==String(req.memberNo))
        return json({ok:false, error:'自分の投稿のみ削除できます'});
      sh.deleteRow(i+1);
      return json({ok:true});
    }
  }
  return json({ok:false, error:'not found'});
}

// ─── 編集（本人 or 管理者） ───────────────────
function editItem(req){
  const isAdmin = isAdminNo(req.memberNo);
  const sheetName = {ideas:'Ideas',events:'Events',help:'Help',memories:'Memories'}[req.type];
  if(!sheetName) return json({ok:false,error:'bad type'});
  const sh = getSheet(sheetName);
  const data = sh.getDataRange().getValues();
  const headers = data[0];
  const idCol = headers.indexOf('id'), noCol = headers.indexOf('memberNo');
  const editable = ['title','desc','place','date','cat','type','era','comment'];
  for(let i=1;i<data.length;i++){
    if(data[i][idCol]===req.id){
      if(!isAdmin && String(data[i][noCol])!==String(req.memberNo))
        return json({ok:false, error:'自分の投稿のみ編集できます'});
      editable.forEach(f=>{
        if(f in (req.fields||{})){
          const c = headers.indexOf(f);
          if(c>=0) sh.getRange(i+1,c+1).setValue(req.fields[f]);
        }
      });
      return json({ok:true});
    }
  }
  return json({ok:false, error:'not found'});
}

// ─── 管理者：複数選択一括削除 ─────────────────
function adminDelete(req){
  if(!isAdminNo(req.memberNo))
    return json({ok:false, error:'管理者権限がありません'});
  const sheetName = {members:'Members',ideas:'Ideas',events:'Events',help:'Help',memories:'Memories'}[req.type];
  if(!sheetName) return json({ok:false,error:'bad type'});

  const sh = getSheet(sheetName);
  const data = sh.getDataRange().getValues();
  const headers = data[0];
  const keyCol = (req.type==='members') ? headers.indexOf('memberNo') : headers.indexOf('id');
  const ids = (req.ids||[]).map(String);

  // 管理者自身は削除不可
  let deleted = 0;
  for(let i=data.length-1;i>=1;i--){  // 下から削除
    const key = String(data[i][keyCol]);
    if(ids.includes(key) && !isAdminNo(key)){
      sh.deleteRow(i+1); deleted++;
    }
  }
  return json({ok:true, deleted});
}
