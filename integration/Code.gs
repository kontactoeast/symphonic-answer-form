// このスクリプトを保存先スプレッドシートの「拡張機能 > Apps Script」に設置。
function doGet(){return HtmlService.createHtmlOutputFromFile('Index').setTitle('Red Bull Symphonic 2026｜客演登録').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);}
function registerGuest(d){
 if(!d || !d.id || !d.artist || !d.contact || !d.email || !Array.isArray(d.events) || !d.events.length) throw new Error('必須情報を確認してください。');
 if(!/^[^\s@,;]+@[^\s@,;]+\.[^\s@,;]+$/.test(d.email))throw new Error('メールアドレスを確認してください。');
 if(!d.emailConfirm || d.email.trim()!==d.emailConfirm.trim())throw new Error('メールアドレスが一致していません。');
 d.email=d.email.trim();
 const requestId=d.id;
 const allowed=['osaka','fukuoka','sendai','yokohama'];
 if(d.events.some(e=>!allowed.includes(e.venue)||!Number.isInteger(e.mealCount)||e.mealCount<0||!['1食希望','不要'].includes(e.mealType)))throw new Error('会場・食数を確認してください。');
 const lock=LockService.getScriptLock();lock.waitLock(30000);
 try{
 const ss=SpreadsheetApp.getActiveSpreadsheet();
 const sheet=(name,headers)=>{let s=ss.getSheetByName(name);if(!s){s=ss.insertSheet(name);s.appendRow(headers);s.setFrozenRows(1);}return s;};
 const basic=sheet('基本情報',['登録ID','登録日時','アーティスト','本名','担当者','電話','メール','食事制限','水常温','水冷水','タオル','その他','状態','持ち込み機材','当日の連絡先','送信識別子','メール送信状態']);
 basic.getRange(1,16,1,2).setValues([['送信識別子','メール送信状態']]);
 const rows=basic.getDataRange().getValues();const prior=rows.findIndex(r=>r[15]===requestId || r[0]===requestId);
 let seq=rows.reduce((n,r)=>Math.max(n,/^RBS-([0-9]+)$/.test(String(r[0]))?Number(String(r[0]).slice(4)):0),0);
 d.id=prior>=0?rows[prior][0]:'RBS-'+String(seq+1).padStart(4,'0');
 if(prior>=0&&rows[prior][12]==='完了')return finishRegistration_(basic,prior+1,d);
 const safe=v=>typeof v==='string'&&/^[=+@\-]/.test(v)?"'"+v:v;
 const add=(s,rs)=>{if(rs.length)s.getRange(s.getLastRow()+1,1,rs.length,rs[0].length).setValues(rs.map(r=>r.map(safe)));};
 let idx=prior+1;if(prior<0){add(basic,[[d.id,new Date(),d.artist,'',d.contact,d.phone,d.email,d.diet,d.waterRoom,d.waterCold,d.towels,d.notes,'処理中',d.equipment||'',d.dayContact||'',requestId,'未送信']]);idx=basic.getLastRow();}
 const people=sheet('入館名簿',['登録ID','会場','アーティスト','氏名カタカナ','役割']);
 const food=sheet('当日手配',['登録ID','会場','アーティスト','会場入り予定時刻','食事種別','食数']);
 // 再試行時は同じ登録IDの途中保存を除去して再作成。
 [people,food].forEach(s=>{const values=s.getDataRange().getValues();for(let i=values.length-1;i>0;i--)if(values[i][0]===d.id)s.deleteRow(i+1);});
 const pr=[],fr=[];d.events.forEach(e=>{pr.push([d.id,e.venueName,d.artist,d.artist,'出演者']);(d.members||[]).filter(m=>m.venues.includes(e.venue)).forEach(m=>pr.push([d.id,e.venueName,d.artist,m.name,m.role]));fr.push([d.id,e.venueName,d.artist,e.arrival,e.mealType,e.mealCount]);});
 add(people,pr);add(food,fr);basic.getRange(idx,13).setValue('完了');SpreadsheetApp.flush();return finishRegistration_(basic,idx,d);
 }finally{lock.releaseLock();}
}

function finishRegistration_(sheet,row,d){
 const state=sheet.getRange(row,17).getValue();
 if(state==='送信済み')return {ok:true,id:d.id,emailSent:true};
 if(state==='送信中')return {ok:true,id:d.id,emailSent:false};
 try{
  if(MailApp.getRemainingDailyQuota()<2)throw new Error('メール送信上限');
  sheet.getRange(row,17).setValue('送信中');SpreadsheetApp.flush();
  MailApp.sendEmail({to:d.email,cc:'symphonic@kontacto.jp',replyTo:'symphonic@kontacto.jp',name:'Red Bull Symphonic 制作',subject:'【Red Bull Symphonic】ご登録内容の確認 '+d.id,body:registrationBody_(d)});
  sheet.getRange(row,17).setValue('送信済み');return {ok:true,id:d.id,emailSent:true};
 }catch(err){sheet.getRange(row,17).setValue('送信未確認：制作確認要');return {ok:true,id:d.id,emailSent:false};}
}
function registrationBody_(d){
 const lines=[d.contact+' 様','','ご登録ありがとうございました。以下の内容で承りました。','登録ID：'+d.id,'','【基本情報】','アーティスト：'+d.artist,'ご担当者：'+d.contact,'携帯電話：'+d.phone,'メール：'+d.email,'当日の連絡先：'+d.dayContact,'','【楽屋口入館者】'];
 (d.members||[]).forEach(m=>lines.push(m.name+'／'+m.role+'／'+m.venues.map(v=>({osaka:'大阪',fukuoka:'福岡',sendai:'仙台',yokohama:'横浜'}[v]||v)).join('・')));
 if(!(d.members||[]).length)lines.push('帯同者なし');
 (d.events||[]).forEach(e=>{lines.push('','【'+e.venueName+'】','会場入り予定時刻：'+e.arrival,'お食事：'+e.mealType+'／'+e.mealCount+'食');});
 lines.push('','【食事・備品・その他のご要望】','食物アレルギー・食事制限：'+(d.diet||'なし'),'水500ml・常温：'+d.waterRoom+'本','水500ml・冷水：'+d.waterCold+'本','フェイスタオル：'+d.towels+'枚','持ち込み機材：'+(d.equipment||'なし'),'特別リクエスト・その他：'+(d.notes||'なし'),'','ご要望は確認のうえ、対応可否をご連絡いたします。','登録後の追加・変更は、制作窓口へお問い合わせください。','symphonic@kontacto.jp','','アンサーシート：','https://kontactoeast.github.io/symphonic-answer-form/');
 return lines.join('\n');
}
// エディターで一度実行し、メール送信権限を承認（メールは送信しません）。
function authorizeMail(){return MailApp.getRemainingDailyQuota();}
