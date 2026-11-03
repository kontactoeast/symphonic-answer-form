// このスクリプトを保存先スプレッドシートの「拡張機能 > Apps Script」に設置。
function doGet(){return HtmlService.createHtmlOutputFromFile('Index').setTitle('Red Bull Symphonic 2026｜客演登録').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);}
function registerGuest(d){
 if(!d || !d.id || !d.artist || !d.contact || !d.email || !Array.isArray(d.events) || !d.events.length) throw new Error('必須情報を確認してください。');
 const allowed=['osaka','fukuoka','sendai','yokohama'];
 if(d.events.some(e=>!allowed.includes(e.venue)||!Number.isInteger(e.mealCount)||e.mealCount<0||!['1食希望','不要'].includes(e.mealType)))throw new Error('会場・食数を確認してください。');
 const lock=LockService.getScriptLock();lock.waitLock(30000);
 try{
 const ss=SpreadsheetApp.getActiveSpreadsheet();
 const sheet=(name,headers)=>{let s=ss.getSheetByName(name);if(!s){s=ss.insertSheet(name);s.appendRow(headers);s.setFrozenRows(1);}return s;};
 const basic=sheet('基本情報',['登録ID','登録日時','アーティスト','本名','担当者','電話','メール','食事制限','水常温','水冷水','タオル','その他','状態','持ち込み機材','当日の連絡先']);
 const rows=basic.getDataRange().getValues();const prior=rows.findIndex(r=>r[0]===d.id);if(prior>=0&&rows[prior][12]==='完了')return {ok:true,id:d.id};
 const safe=v=>typeof v==='string'&&/^[=+@\-]/.test(v)?"'"+v:v;
 const add=(s,rs)=>{if(rs.length)s.getRange(s.getLastRow()+1,1,rs.length,rs[0].length).setValues(rs.map(r=>r.map(safe)));};
 let idx=prior+1;if(prior<0){add(basic,[[d.id,new Date(),d.artist,(d.artists||[]).map(a=>a.name).join('、'),d.contact,d.phone,d.email,d.diet,d.waterRoom,d.waterCold,d.towels,d.notes,'処理中',d.equipment||'',d.dayContact||'']]);idx=basic.getLastRow();}
 const people=sheet('入館名簿',['登録ID','会場','アーティスト','氏名カタカナ','役割']);
 const food=sheet('当日手配',['登録ID','会場','アーティスト','会場入り予定時刻','食事種別','食数']);
 const cars=sheet('車両',['登録ID','会場','アーティスト','車種','ナンバー','利用方法','到着出発']);
 // 再試行時は同じ登録IDの途中保存を除去して再作成。
 [people,food,cars].forEach(s=>{const values=s.getDataRange().getValues();for(let i=values.length-1;i>0;i--)if(values[i][0]===d.id)s.deleteRow(i+1);});
 const pr=[],fr=[],cr=[];d.events.forEach(e=>{(d.artists||[]).forEach(a=>pr.push([d.id,e.venueName,d.artist,a.name,'出演者']));(d.members||[]).filter(m=>m.venues.includes(e.venue)).forEach(m=>pr.push([d.id,e.venueName,d.artist,m.name,m.role]));fr.push([d.id,e.venueName,d.artist,e.arrival,e.mealType,e.mealCount]);(e.vehicles||[]).forEach(c=>cr.push([d.id,e.venueName,d.artist,c.type,c.plate,c.parking,c.time]));});
 add(people,pr);add(food,fr);add(cars,cr);basic.getRange(idx,13).setValue('完了');SpreadsheetApp.flush();return {ok:true,id:d.id};
 }finally{lock.releaseLock();}
}
