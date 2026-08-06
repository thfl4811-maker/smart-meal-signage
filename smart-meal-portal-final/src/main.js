import * as XLSX from 'xlsx';import QRCode from 'qrcode';import './style.css';
const $=s=>document.querySelector(s),$$=s=>[...document.querySelectorAll(s)],esc=s=>String(s??'').replace(/[&<>"']/g,m=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#039;'}[m]));
const A={1:'난류',2:'우유',3:'메밀',4:'땅콩',5:'대두',6:'밀',7:'고등어',8:'게',9:'새우',10:'돼지고기',11:'복숭아',12:'토마토',13:'아황산염',14:'호두',15:'닭고기',16:'소고기',17:'오징어',18:'조개류',19:'잣'};

/* ── 분류 & 급별 배식량 ── */
const STEPS=[['rice','주 식','🍚'],['soup','국·찌개','🥣'],['main','메인요리','🍖'],['side','곁들임찬','🥗'],['kimchi','김치류','🌶️'],['dessert','후 식','🍎']];
const KW={
  rice:['밥','라이스','죽','덮밥','비빔','볶음밥','국수','면','파스타','스파게티','떡국','만두국'],
  soup:['국','탕','찌개','전골','스프','수제비'],
  kimchi:['김치','깍두기','총각','석박지','동치미'],
  dessert:['과일','주스','음료','요구르트','요거트','아이스크림','케이크','쿠키','빵','푸딩','젤리','우유','수박','메론','샤인','사과','귤','바나나','파인애플','약과','카스테라','비타','식혜'],
  main:['갈비','불고기','제육','돈까스','돈가스','닭','치킨','오리','스테이크','장조림','찜','구이','탕수','강정','생선','고등어','삼치','조기','미트볼','떡갈비','함박','오징어볶음','쭈꾸미','낙지','새우','너비아니','산적','커틀릿','까스','수육','보쌈','편육'],
};
function stepOf(name){const n=name.replace(/\s/g,'');
  for(const k of ['soup','kimchi','dessert'])if(KW[k].some(w=>n.includes(w)))return k;
  if(KW.rice.some(w=>n.includes(w)))return 'rice';
  if(KW.main.some(w=>n.includes(w)))return 'main';
  return 'side';}
function schoolLevel(){const n=school?.schoolName||'';if(/초등/.test(n))return'초등';if(/중학/.test(n))return'중등';return'고등';}
const PORTION={
  rice:{초등:[70,'어린이 주먹 1.5개 분량'],중등:[85,'야구공 1개 수북이 분량'],고등:[100,'테니스공 1개 분량 (1인 식판 밥칸 가득)']},
  soup:{초등:[50,'국그릇 절반, 건더기 위주'],중등:[50,'건더기 위주 수북이 + 국물 절반'],고등:[50,'건더기 위주 수북이 + 국물 3분의 1']},
  main:{초등:[80,'어린이 손바닥 크기 분량'],중등:[90,'종이컵 1컵 분량'],고등:[100,'종이컵 1컵 수북이 분량']},
  side:{초등:[60,'종이컵 3분의 1 분량'],중등:[70,'종이컵 절반 분량'],고등:[70,'종이컵 1/2컵 한 젓가락 분량']},
  kimchi:{초등:[30,'집게 기준 1회 가볍게 배식'],중등:[40,'집게 기준 약 2회 가볍게 배식'],고등:[40,'집게 기준 약 2회 가볍게 배식']},
  dessert:{초등:[100,'개별 포장 1개 정량 섭취'],중등:[100,'개별 포장 1개 정량 섭취'],고등:[100,'개별 포장 1개 정량 섭취']},
};
const BAR={rice:'bg-sky-500',soup:'bg-amber-400',main:'bg-emerald-500',side:'bg-teal-400',kimchi:'bg-orange-500',dessert:'bg-rose-500'};

const p=new URLSearchParams(location.search);
let school=JSON.parse(localStorage.getItem('meal_school')||'null'),month=p.get('month')||new Date().toISOString().slice(0,7),data={},date='',tab='day',allergies=new Set(JSON.parse(localStorage.getItem('allergies')||'[]')),allergyOn=localStorage.getItem('allergyOn')==='1',source='api';
const signage=p.get('mode')==='signage',student=p.get('mode')==='student';
if(p.get('office'))school={officeCode:p.get('office'),schoolCode:p.get('school'),schoolName:p.get('name')||'학교 급식'};
window.addEventListener('DOMContentLoaded',()=>student?bootStudent():signage?bootSignage():renderHome());

/* ═══════ 관리자 홈 ═══════ */
function renderHome(){app.innerHTML=`
<div class="max-w-7xl mx-auto px-4 py-5">
  <header class="bg-white border-b-4 border-brand-500/30 rounded-3xl shadow-sm px-6 py-4 flex items-center justify-between sticky top-3 z-40">
    <div class="flex items-center gap-3">
      <div class="w-13 h-13 w-14 h-14 rounded-2xl bg-brand-50 flex items-center justify-center text-3xl shadow-inner border border-brand-100">🥗</div>
      <div>
        <h1 class="text-xl md:text-2xl font-black text-gray-800 flex items-center gap-2 flex-wrap">스마트 월간식단 뷰어
          <span class="text-xs bg-brand-500 text-white px-3 py-1 rounded-full font-bold">${school?esc(schoolLevel())+' 배식 가이드':'식단 사이니지 2.0'}</span>
        </h1>
        <p class="text-xs md:text-sm text-gray-400 font-bold">API 또는 엑셀로 불러오고 수정본을 사이니지·학생 화면에 표시합니다.</p>
      </div>
    </div>
    ${school?`<button id="change" class="text-sm bg-gray-100 hover:bg-gray-200 text-gray-600 font-bold px-4 py-2.5 rounded-xl border-2 border-gray-200 transition-all">학교 변경</button>`:''}
  </header>
  ${school?workspace():landing()}
  <div id="modal" class="fixed inset-0 bg-slate-900/70 z-50 hidden items-center justify-center p-5"></div>
</div>`;bind()}

function landing(){return `
<section class="max-w-2xl mx-auto mt-16 text-center">
  <h2 class="text-2xl font-black mb-6">식단 불러오기</h2>
  <div class="grid md:grid-cols-2 gap-4">
    <button id="apiMode" class="bg-white border-2 border-gray-200 hover:border-brand-300 rounded-3xl p-8 card-hover text-left">
      <span class="text-3xl block mb-3">🔎</span><b class="block text-lg">나이스 API</b><span class="text-gray-400 text-sm font-bold">학교 검색으로 자동 불러오기</span>
    </button>
    <label class="bg-white border-2 border-gray-200 hover:border-brand-300 rounded-3xl p-8 card-hover text-left cursor-pointer">
      <span class="text-3xl block mb-3">📄</span><b class="block text-lg">월간식단 엑셀 업로드</b><span class="text-gray-400 text-sm font-bold">학교 파일을 직접 분석</span>
      <input id="excel" type="file" accept=".xlsx,.xls" class="hidden">
    </label>
  </div>
  <div class="flex mt-5 bg-white border-2 border-gray-200 rounded-2xl p-2 shadow-sm">
    <input id="q" placeholder="학교명 입력" class="flex-1 border-0 outline-none px-4 text-sm">
    <button id="find" class="bg-brand-600 hover:bg-brand-700 text-white font-black rounded-xl px-6 py-3 text-sm transition-all">학교 검색</button>
  </div>
  <div id="results" class="mt-3 space-y-2 text-left"></div>
</section>`}

function workspace(){return `
<section class="bg-white border border-gray-200 rounded-3xl shadow-sm p-5 mt-4 flex flex-wrap items-center gap-3">
  <div class="flex-1 min-w-40">
    <small class="text-brand-700 font-black text-xs">${source==='api'?'나이스 API':'엑셀 업로드'} · ${schoolLevel()} 기준</small>
    <h2 class="text-xl font-black">${esc(school.schoolName)}</h2>
  </div>
  <input id="month" type="month" value="${month}" class="border-2 border-gray-200 rounded-xl px-3 py-2.5 text-sm font-bold">
  <button id="reload" class="bg-brand-500 hover:bg-brand-600 text-white font-black px-4 py-2.5 rounded-xl text-sm transition-all">식단 불러오기</button>
  <button id="filter" class="bg-red-50 hover:bg-red-100 text-red-700 font-black px-4 py-2.5 rounded-xl text-sm border-2 border-red-200 transition-all">🚨 알레르기</button>
  <button id="signage" class="bg-gray-800 hover:bg-gray-900 text-white font-black px-4 py-2.5 rounded-xl text-sm transition-all">📺 사이니지·공유</button>
</section>
<nav class="flex gap-2 mt-4 bg-white border border-gray-200 rounded-2xl p-2 shadow-sm">
  ${[['day','fa-calendar-day','일간'],['week','fa-calendar-week','주간'],['month','fa-calendar-days','월간'],['edit','fa-pen-to-square','월 전체 편집'],['detail','fa-file-pen','상세 편집']].map(x=>`
  <button data-tab="${x[0]}" class="flex-1 py-3 rounded-xl text-sm font-black transition-all flex items-center justify-center gap-2 ${tab===x[0]?'bg-brand-500 text-white shadow-md':'text-gray-400 hover:text-gray-700 hover:bg-gray-50'}"><i class="fa-solid ${x[1]}"></i> ${x[2]}</button>`).join('')}
</nav>
<main id="content" class="mt-4"><div class="bg-white rounded-3xl p-16 text-center text-gray-400">식단을 불러오는 중입니다.</div></main>`}

function bind(){
  if($('#change'))$('#change').onclick=()=>{localStorage.removeItem('meal_school');school=null;data={};renderHome()};
  if($('#find')){$('#find').onclick=findSchools;$('#q').onkeydown=e=>e.key==='Enter'&&findSchools()}
  if($('#excel'))$('#excel').onchange=loadExcel;
  if($('#reload'))$('#reload').onclick=loadApi;
  if($('#month'))$('#month').onchange=e=>{month=e.target.value;loadApi()};
  if($('#filter'))$('#filter').onclick=openFilter;
  if($('#signage'))$('#signage').onclick=openSignage;
  $$('[data-tab]').forEach(b=>b.onclick=()=>{tab=b.dataset.tab;renderHome();});
  if(school&&source==='api')loadApi()}

async function findSchools(){const q=$('#q').value.trim();if(q.length<2)return alert('두 글자 이상 입력하세요.');$('#results').innerHTML='<p class="text-gray-400 text-sm">검색 중…</p>';const r=await fetch('/api/schools?q='+encodeURIComponent(q)),j=await r.json();
$('#results').innerHTML=j.map((s,i)=>`<button class="school w-full bg-white border-2 border-gray-200 hover:border-brand-300 rounded-2xl p-4 text-left card-hover" data-i="${i}"><b class="block">${esc(s.schoolName)}</b><span class="text-gray-400 text-xs font-bold">${esc(s.officeName)} · ${esc(s.address)}</span></button>`).join('');
$$('.school').forEach(b=>b.onclick=()=>{school=j[+b.dataset.i];source='api';localStorage.setItem('meal_school',JSON.stringify(school));renderHome()})}

async function loadApi(){try{$('#content').innerHTML='<div class="bg-white rounded-3xl p-16 text-center text-gray-400">나이스 식단을 불러오는 중입니다.</div>';const r=await fetch(`/api/meals?office=${school.officeCode}&school=${school.schoolCode}&month=${month.replace('-','')}`),rows=await r.json();if(!r.ok)throw Error(rows.error);data=parseRows(rows);await applySaved();setDefaultDate();renderTab()}catch(e){$('#content').innerHTML=`<div class="bg-white rounded-3xl p-16 text-center text-gray-400">${esc(e.message)}</div>`}}
function parseRows(rows){const out={};rows.filter(r=>r.mealCode==='2').forEach(r=>{const d=`${r.date.slice(0,4)}-${r.date.slice(4,6)}-${r.date.slice(6,8)}`;const items=r.dishes.split('<br/>').map(parseDish).filter(Boolean);const nutrient={};r.nutrients.split('<br/>').forEach(x=>{const m=x.match(/^(.+?)\s*:\s*([\d.]+)\s*(.*)$/);if(m)nutrient[m[1].trim()]={value:m[2],unit:m[3]}});out[d]={date:d,items,calories:r.calories,nutrient,note:'',edited:false}});return out}
function parseDish(t){t=t.replace(/<[^>]+>/g,'').trim();if(!t)return null;const al=[];for(const m of t.matchAll(/\(([\d.]+)\)/g))m[1].split('.').map(Number).forEach(n=>n>=1&&n<=19&&al.push(n));let name=t.replace(/\(([\d.]+)\)/g,'').trim(),type=/자율/.test(name)?'self':'normal';return{name,allergy:[...new Set(al)],type,category:'',order:0}}
async function applySaved(){try{const r=await fetch(`/api/project?office=${school.officeCode}&school=${school.schoolCode}&month=${month}`),j=await r.json();if(j?.payload?.days)data=j.payload.days}catch{}}
function setDefaultDate(){const today=new Date().toISOString().slice(0,10);date=data[today]?today:Object.keys(data).sort()[0]||''}

async function loadExcel(e){const f=e.target.files[0];if(!f)return;
try{const book=XLSX.read(await f.arrayBuffer(),{type:'array'});const ws=book.Sheets[book.SheetNames[0]];const rows=XLSX.utils.sheet_to_json(ws,{header:1,defval:'',raw:false});
let year='',mon='',schoolName='';
for(const row of rows){for(const cell of row){const text=String(cell||'').trim();const ym=text.match(/조회년월\s*:\s*(\d{4})년\s*(\d{1,2})월/);if(ym){year=ym[1];mon=String(ym[2]).padStart(2,'0')}if(!schoolName&&/학교$/.test(text)&&!/조회|월간|식단/.test(text))schoolName=text;}}
if(!year||!mon)throw Error('엑셀에서 조회 연월을 찾지 못했습니다.');
const parsed={};
const parseExcelMeal=(cell,dateKey)=>{const raw=String(cell||'').replace(/\r/g,'').trim();if(!raw||raw==='0')return null;
const lines=raw.split('\n').map(x=>x.trim()).filter(Boolean);const energyIndex=lines.findIndex(x=>/^\*\s*에너지\/단백질\/칼슘\/철/.test(x));
let menuLines=energyIndex>=0?lines.slice(0,energyIndex):lines.slice();menuLines=menuLines.filter(x=>x!=='[식단]'&&x!=='중식'&&x!=='0');
let calories='',nutrient={};
if(energyIndex>=0&&lines[energyIndex+1]){const nums=lines[energyIndex+1].replace(/,/g,'').split('/').map(x=>x.trim());if(nums[0])calories=`${nums[0]} Kcal`;if(nums[1])nutrient['단백질']={value:nums[1],unit:'g'};if(nums[2])nutrient['칼슘']={value:nums[2],unit:'mg'};if(nums[3])nutrient['철분']={value:nums[3],unit:'mg'};}
const items=menuLines.map(parseDish).filter(Boolean);if(!items.length)return null;
return{date:dateKey,items,calories,nutrient,note:'',edited:false};};
for(let r=0;r<rows.length-1;r++){const dateRow=rows[r]||[],mealRow=rows[r+1]||[];const numericDates=[];
for(let c=0;c<dateRow.length;c++){const value=String(dateRow[c]??'').trim();if(/^\d{1,2}$/.test(value)){const day=Number(value);if(day>=1&&day<=31)numericDates.push({c,day})}}
if(!numericDates.length)continue;
const looksLikeMealRow=mealRow.some(cell=>{const t=String(cell||'');return /\n/.test(t)||/\*\s*에너지\/단백질\/칼슘\/철/.test(t)});
if(!looksLikeMealRow)continue;
for(const{c,day}of numericDates){const dateKey=`${year}-${mon}-${String(day).padStart(2,'0')}`;const meal=parseExcelMeal(mealRow[c],dateKey);if(meal)parsed[dateKey]=meal}}
if(!Object.keys(parsed).length)throw Error('날짜별 식단을 찾지 못했습니다. 나이스 월간식단 엑셀 파일인지 확인하세요.');
data=parsed;school={officeCode:'EXCEL',schoolCode:`LOCAL-${year}${mon}`,schoolName:schoolName||f.name.replace(/\.[^.]+$/,'')};source='excel';month=`${year}-${mon}`;localStorage.setItem('meal_school',JSON.stringify(school));setDefaultDate();renderHome();renderTab();
alert(`${school.schoolName} ${year}년 ${Number(mon)}월 식단 ${Object.keys(data).length}일을 불러왔습니다.`);
}catch(err){alert(`엑셀 업로드 실패: ${err.message}`);e.target.value='';}}

function renderTab(){if(!$('#content'))return;({day:renderDay,week:renderWeek,month:renderMonth,edit:renderEdit,detail:renderDetail}[tab]||renderDay)()}

/* ═══════ 일간 (STEP 카드 + AI팁) ═══════ */
function grouped(x){const g={};STEPS.forEach(([k])=>g[k]=[]);x.items.forEach(i=>g[stepOf(i.name)].push(i));return g}
function stepCards(x){const lv=schoolLevel(),g=grouped(x);let n=0;
return STEPS.flatMap(([key,label,emo])=>g[key].map(i=>{n++;const hit=allergyOn&&i.allergy.some(a=>allergies.has(a));const[pct,tip]=PORTION[key][lv];
return `<article class="bg-white border-2 ${hit?'border-red-400 bg-red-50/40':'border-gray-200'} rounded-3xl p-5 shadow-sm relative card-hover">
  ${hit?'<span class="absolute top-3 right-3 bg-red-600 text-white text-[10px] font-black px-2.5 py-1 rounded-full">주의</span>':''}
  <span class="text-[11px] font-black text-gray-400 bg-gray-100 px-2.5 py-1 rounded-lg">STEP ${String(n).padStart(2,'0')} · ${label}</span>
  <div class="flex items-center gap-3 mt-4 mb-1">
    <span class="text-4xl">${emo}</span>
    <div>
      <h3 class="font-black text-lg leading-tight">${esc(i.name)}</h3>
      ${i.allergy.length?`<span class="text-[11px] font-black text-red-500 bg-red-50 border border-red-100 px-2 py-0.5 rounded-full">알레르기: ${i.allergy.join(',')}</span>`:'<span class="text-[11px] font-bold text-gray-300 bg-gray-50 px-2 py-0.5 rounded-full">안심 메뉴</span>'}
    </div>
  </div>
  <div class="mt-4 pt-4 border-t border-gray-100">
    <div class="flex justify-between text-xs font-black mb-1.5"><span class="text-gray-400">⚖️ ${lv}학생 1인 권장 정량</span><span>${pct}%</span></div>
    <div class="h-2 bg-gray-100 rounded-full overflow-hidden"><div class="h-full ${BAR[key]} rounded-full" style="width:${pct}%"></div></div>
    <div class="mt-3 text-xs font-bold text-gray-500 bg-gray-50 border border-gray-100 rounded-xl px-3 py-2.5">✏️ ${tip}</div>
  </div>
</article>`}))
.join('')}
function aiTips(x){const names=x.items.map(i=>i.name).join(' ');const tips=[];
if(/국|찌개|탕|전골|라면|짬뽕/.test(names))tips.push(['염','bg-amber-100 text-amber-800','나트륨 저감 가이드',`오늘 국물 메뉴가 있어요. ${schoolLevel()}학생 하루 나트륨 권장량은 2,000mg 이하! 국물은 3분의 1만, 건더기 위주로 먹으면 나트륨을 절반으로 줄일 수 있어요.`]);
if(/주스|아이스크림|케이크|빵|초콜릿|쿠키|약과|카스테라|식혜|요구르트/.test(names))tips.push(['당','bg-red-100 text-red-800','당류 제한 식생활','달콤한 후식은 정량 1개만! 남은 갈증은 물이나 우유로 채우면 혈당 스파이크를 막고 오후 졸음도 예방할 수 있어요.']);
if(/샐러드|나물|채소|쌈|무침|숙주|시금치|브로콜리/.test(names))tips.push(['섬','bg-emerald-100 text-emerald-800','식이섬유 청소기 효과','채소 반찬을 먼저 먹으면 식이섬유가 혈당 상승을 늦추고 장 건강까지 지켜줘요. 오늘 채소 반찬부터 시작해보세요!']);
if(/고기|불고기|제육|닭|돈까스|갈비|생선|고등어|삼치|두부|계란|달걀/.test(names))tips.push(['단','bg-sky-100 text-sky-800','단백질 성장 파워',`성장기 ${schoolLevel()}학생은 단백질이 필수! 오늘 메인요리를 남기지 말고 권장량만큼 먹으면 근육과 키 성장에 도움이 돼요.`]);
if(/우유|치즈|요구르트|멸치|뼈/.test(names))tips.push(['칼','bg-violet-100 text-violet-800','칼슘 뼈 튼튼 교실','칼슘이 풍부한 메뉴가 있어요. 비타민D(햇볕)와 함께하면 흡수율이 올라가요. 점심 후 10분 산책 어때요?']);
if(!tips.length)tips.push(['균','bg-brand-100 text-brand-800','균형 잡힌 식판','오늘은 모든 반찬을 골고루! 편식 없이 다양한 색깔의 음식을 먹는 것이 최고의 영양 습관이에요.']);
return tips.map(t=>`<div class="flex gap-3 items-start"><div class="w-7 h-7 rounded-full ${t[1]} flex items-center justify-center font-black text-xs shrink-0">${t[0]}</div><div><span class="font-black text-sm block mb-0.5">${t[2]}</span><p class="text-xs text-gray-500 leading-relaxed">${t[3]}</p></div></div>`).join('')}

function renderDay(){const x=data[date];if(!x)return $('#content').innerHTML='<div class="bg-white rounded-3xl p-16 text-center text-gray-400">식단이 없습니다.</div>';
const dt=new Date(date+'T00:00:00');
$('#content').innerHTML=`
<section class="bg-brand-700 text-white rounded-3xl p-5 flex flex-wrap items-center gap-4 shadow-md">
  <button id="prev" class="w-11 h-11 rounded-2xl bg-white/20 hover:bg-white/30 font-black text-xl transition-all">‹</button>
  <div class="flex-1 text-center">
    <h2 class="text-2xl font-black">📅 ${dt.getFullYear()}년 ${dt.getMonth()+1}월 ${dt.getDate()}일 (${'일월화수목금토'[dt.getDay()]}요일)</h2>
    <p class="text-brand-100 text-sm font-bold mt-0.5">${esc(school.schoolName)} · ${schoolLevel()} 배식 기준 중식</p>
  </div>
  <button id="next" class="w-11 h-11 rounded-2xl bg-white/20 hover:bg-white/30 font-black text-xl transition-all">›</button>
  <button id="today" class="px-4 h-11 rounded-2xl bg-white text-brand-700 font-black text-sm">오늘 급식</button>
</section>
${filterBanner()}
<section class="grid lg:grid-cols-12 gap-5 mt-4">
  <div class="lg:col-span-8"><div class="grid md:grid-cols-2 xl:grid-cols-3 gap-4">${stepCards(x)}</div>
    <div class="mt-5 bg-white border border-gray-200 rounded-3xl p-5 flex flex-wrap justify-around items-center gap-4 text-center shadow-sm">
      <div><span class="text-xs text-gray-400 block font-black">🔥 에너지</span><span class="text-2xl font-black text-sky-600">${esc(x.calories||'-')}</span></div>
      ${['단백질','칼슘','철분'].map((k,i)=>`<div class="hidden md:block w-px h-10 bg-gray-200"></div><div><span class="text-xs text-gray-400 block font-black">${['🥩','🥛','🔩'][i]} ${k}</span><span class="text-2xl font-black text-gray-700">${x.nutrient[k]?x.nutrient[k].value+' '+x.nutrient[k].unit:'-'}</span></div>`).join('')}
    </div>
  </div>
  <aside class="lg:col-span-4 space-y-4">
    <div class="bg-white border border-gray-200 rounded-3xl p-5 shadow-sm">
      <h3 class="font-black text-base mb-1 flex items-center gap-2"><i class="fa-solid fa-list-check text-brand-600"></i> 오늘 급식 전체 리스트</h3>
      <p class="text-xs text-gray-400 font-bold mb-4">알레르기 유무를 큰 글씨로 점검하세요.</p>
      <ul class="space-y-2.5">${x.items.map(i=>{const hit=allergyOn&&i.allergy.some(a=>allergies.has(a));
        return `<li class="flex items-center gap-2 text-sm font-bold ${hit?'bg-red-50 border border-red-100 rounded-xl px-3 py-2 text-red-800':''}">${icon(i.name)} ${esc(i.name)}${i.allergy.length?`<small class="text-red-500 font-black">${i.allergy.join('·')}</small>`:''}${hit?'<b class="ml-auto bg-red-600 text-white text-[10px] px-2 py-0.5 rounded-full">주의</b>':''}</li>`}).join('')}</ul>
    </div>
    <div class="bg-slate-50 border border-gray-100 rounded-3xl p-5">
      <h3 class="font-black text-base mb-4 flex items-center gap-1.5"><span class="text-xl">💡</span> AI 영양교사의 오늘 식단 팁</h3>
      <div class="space-y-4">${aiTips(x)}</div>
    </div>
  </aside>
</section>`;
$('#prev').onclick=()=>move(-1);$('#next').onclick=()=>move(1);
$('#today').onclick=()=>{const t=new Date().toISOString().slice(0,10);if(data[t]){date=t;renderDay()}else alert('오늘은 등록된 급식이 없습니다.')};}

/* ═══════ 주간 ═══════ */
function renderWeek(){if(!date)setDefaultDate();const base=new Date(date+'T00:00:00');const mon=new Date(base);mon.setDate(base.getDate()-((base.getDay()+6)%7));
const days=[...Array(5)].map((_,i)=>{const d=new Date(mon);d.setDate(mon.getDate()+i);return d.toISOString().slice(0,10)});
$('#content').innerHTML=`
<div class="flex items-center justify-between mb-4">
  <h2 class="text-xl font-black">🗓️ 주간 식단 (${days[0].slice(5)} ~ ${days[4].slice(5)})</h2>
  <div class="flex gap-2">
    <button id="wprev" class="bg-white border-2 border-gray-200 rounded-xl px-4 py-2 font-black text-sm">‹ 이전 주</button>
    <button id="wnext" class="bg-white border-2 border-gray-200 rounded-xl px-4 py-2 font-black text-sm">다음 주 ›</button>
  </div>
</div>
${filterBanner()}
<div class="grid md:grid-cols-5 gap-3">
  ${days.map(d=>{const x=data[d],dt=new Date(d+'T00:00:00');const dayHit=x&&allergyOn&&x.items.some(i=>i.allergy.some(n=>allergies.has(n)));
  return `<article class="bg-white border-2 ${dayHit?'border-red-400':'border-gray-200'} rounded-3xl p-4 min-h-56 ${x?'cursor-pointer card-hover':''} shadow-sm" ${x?`data-date="${d}"`:''}>
    <div class="flex justify-between items-center mb-3 pb-2 border-b border-gray-100">
      <b class="text-sm font-black">${dt.getMonth()+1}/${dt.getDate()} ${'일월화수목금토'[dt.getDay()]}</b>
      ${dayHit?'<span class="text-[10px] bg-red-100 text-red-700 font-black px-2 py-0.5 rounded-full">주의</span>':''}
    </div>
    ${x?x.items.map(i=>{const hit=allergyOn&&i.allergy.some(n=>allergies.has(n));return `<div class="text-xs font-bold my-1.5 leading-snug ${hit?'bg-red-50 rounded-lg px-2 py-1 text-red-800':''}">${icon(i.name)} ${esc(i.name)}</div>`}).join('')+`<footer class="border-t border-dashed mt-2 pt-2 text-[11px] text-gray-400 font-bold">${esc(x.calories||'')}</footer>`
    :'<p class="text-xs text-gray-300 text-center mt-10 font-bold">급식 없음</p>'}
  </article>`}).join('')}
</div>`;
$('#wprev').onclick=()=>{const d=new Date(date+'T00:00:00');d.setDate(d.getDate()-7);date=nearest(d);renderWeek()};
$('#wnext').onclick=()=>{const d=new Date(date+'T00:00:00');d.setDate(d.getDate()+7);date=nearest(d);renderWeek()};
$$('.card-hover[data-date]').forEach(c=>c.onclick=()=>{date=c.dataset.date;tab='day';renderHome()});}
function nearest(d){const t=d.toISOString().slice(0,10);if(data[t])return t;const ds=Object.keys(data).sort();return ds.find(x=>x>=t)||ds[ds.length-1]||t}

/* ═══════ 월간 (왼쪽 알레르기 패널) ═══════ */
function renderMonth(){const cards=Object.keys(data).sort().map(d=>{const x=data[d],dt=new Date(d+'T00:00:00');const dayHit=allergyOn&&x.items.some(i=>i.allergy.some(n=>allergies.has(n)));
return `<article class="bg-white border-2 ${dayHit?'border-red-400 shadow-red-100':'border-gray-200'} rounded-3xl p-4 min-h-56 cursor-pointer card-hover shadow-sm" data-date="${d}">
  <div class="flex justify-between items-center mb-2 pb-2 border-b border-gray-100">
    <b class="text-sm font-black">${dt.getMonth()+1}월 ${dt.getDate()}일 ${'일월화수목금토'[dt.getDay()]}요일</b>
    <div class="flex gap-1">${dayHit?'<span class="text-[10px] bg-red-100 text-red-700 font-black px-2 py-0.5 rounded-full">알레르기 주의</span>':''}${x.edited?'<span class="text-[10px] bg-brand-50 text-brand-700 font-black px-2 py-0.5 rounded-full">수정됨</span>':''}</div>
  </div>
  ${x.items.map(i=>{const hit=allergyOn&&i.allergy.some(n=>allergies.has(n));return `<div class="text-xs font-bold my-1.5 leading-snug flex items-center gap-1 flex-wrap ${hit?'bg-red-50 border border-red-100 rounded-lg px-2 py-1 text-red-800':''}">${icon(i.name)} ${esc(i.name)}${i.allergy.length?`<small class="text-red-400 font-black">${i.allergy.join('·')}</small>`:''}${hit?'<b class="ml-auto bg-red-600 text-white text-[9px] px-1.5 py-0.5 rounded-full">주의</b>':''}</div>`}).join('')}
  <footer class="border-t border-dashed mt-2 pt-2 text-[11px] text-gray-400 font-bold">${esc(x.calories||'영양정보 없음')}</footer>
</article>`}).join('');
$('#content').innerHTML=`
<div class="grid lg:grid-cols-12 gap-5">
  <aside class="lg:col-span-3">
    <div class="bg-white border border-gray-200 rounded-3xl p-5 shadow-sm lg:sticky lg:top-24">
      <h3 class="font-black text-base mb-1 flex items-center gap-2">🚨 알레르기 필터</h3>
      <p class="text-xs text-gray-400 font-bold mb-4">번호를 누르면 해당 메뉴가 오른쪽에서 강조돼요.</p>
      <div class="grid grid-cols-3 gap-1.5" id="sideAllergy">
        ${Object.entries(A).map(([n,v])=>`<button data-sa="${n}" class="border-2 rounded-xl py-2 px-1 text-center transition-all ${allergies.has(+n)?'border-red-400 bg-red-50 text-red-700':'border-gray-200 bg-gray-50 text-gray-500 hover:border-gray-300'}"><b class="block text-sm font-black">${n}</b><span class="text-[10px] font-bold">${v}</span></button>`).join('')}
      </div>
      <button id="sideClear" class="w-full mt-3 bg-gray-100 hover:bg-gray-200 text-gray-500 rounded-xl py-2.5 text-xs font-black transition-all">전체 해제</button>
      <div class="mt-3 text-center text-xs font-black ${allergies.size?'text-red-600':'text-gray-300'}">${allergies.size?`${allergies.size}개 선택 · 필터 켜짐`:'선택 없음'}</div>
    </div>
  </aside>
  <div class="lg:col-span-9">
    <div class="flex items-center justify-between mb-3"><h2 class="text-xl font-black">${month} 전체 식단</h2></div>
    ${filterBanner()}
    <section class="grid md:grid-cols-2 xl:grid-cols-3 gap-3 mt-2">${cards||'<div class="bg-white rounded-3xl p-16 text-center text-gray-400 col-span-full">식단이 없습니다.</div>'}</section>
  </div>
</div>`;
$$('[data-sa]').forEach(b=>b.onclick=()=>{const n=+b.dataset.sa;allergies.has(n)?allergies.delete(n):allergies.add(n);allergyOn=allergies.size>0;localStorage.setItem('allergies',JSON.stringify([...allergies]));localStorage.setItem('allergyOn',allergyOn?'1':'0');renderMonth()});
$('#sideClear').onclick=()=>{allergies.clear();allergyOn=false;localStorage.setItem('allergies','[]');localStorage.setItem('allergyOn','0');renderMonth()};
$$('[data-date]').forEach(c=>c.onclick=()=>{date=c.dataset.date;tab='day';renderHome()});}

function filterBanner(){if(!allergyOn||!allergies.size)return'';return `<div class="flex gap-2 flex-wrap items-center bg-white border-2 border-red-200 rounded-2xl px-4 py-3 my-3"><b class="text-sm font-black">🚨 알레르기 필터 적용 중</b>${[...allergies].map(n=>`<span class="bg-red-100 text-red-700 rounded-full px-2.5 py-1 text-[11px] font-black">${n} ${A[n]}</span>`).join('')}</div>`}
function move(n){const ds=Object.keys(data).sort(),i=ds.indexOf(date);if(ds[i+n]){date=ds[i+n];renderDay()}}

/* ═══════ 편집 (기존 유지) ═══════ */
function renderEdit(){const rows=Object.keys(data).sort().flatMap(d=>data[d].items.map((i,k)=>`<tr data-d="${d}" data-i="${k}" class="border-b border-gray-100"><td class="p-2 text-xs font-bold">${d.slice(5)}</td><td class="p-2"><input class="name w-full border-2 border-gray-200 rounded-lg px-2 py-1.5 text-sm" value="${esc(i.name)}"></td><td class="p-2"><input class="alg w-24 border-2 border-gray-200 rounded-lg px-2 py-1.5 text-sm" value="${i.allergy.join(',')}"></td><td class="p-2"><select class="type border-2 border-gray-200 rounded-lg px-2 py-1.5 text-sm"><option value="normal">일반</option><option value="choiceA" ${i.type==='choiceA'?'selected':''}>선택 A</option><option value="choiceB" ${i.type==='choiceB'?'selected':''}>선택 B</option><option value="self" ${i.type==='self'?'selected':''}>자율배식</option></select></td><td class="p-2"><button class="saveRow bg-gray-100 hover:bg-gray-200 rounded-lg px-3 py-1.5 text-xs font-black">적용</button></td></tr>`)).join('');
$('#content').innerHTML=`<section class="bg-white border border-gray-200 rounded-3xl p-6 shadow-sm">
<div class="flex items-center justify-between mb-4"><h2 class="text-xl font-black">월 전체 편집</h2><div class="flex gap-2"><button id="replace" class="bg-gray-100 hover:bg-gray-200 rounded-xl px-4 py-2.5 text-sm font-black">찾아 바꾸기</button><button id="saveAll" class="bg-brand-500 hover:bg-brand-600 text-white rounded-xl px-4 py-2.5 text-sm font-black">전체 저장</button></div></div>
<div class="overflow-auto"><table class="w-full min-w-[850px] border-collapse"><thead><tr class="border-b-2 border-gray-200 text-left text-xs font-black text-gray-400"><th class="p-2">날짜</th><th class="p-2">메뉴명</th><th class="p-2">알레르기</th><th class="p-2">표시</th><th></th></tr></thead><tbody>${rows}</tbody></table></div></section>`;
$$('.saveRow').forEach(b=>b.onclick=()=>applyRow(b.closest('tr')));$('#replace').onclick=replaceAll;$('#saveAll').onclick=saveProject}
function applyRow(tr){const x=data[tr.dataset.d].items[+tr.dataset.i];x.name=tr.querySelector('.name').value;x.allergy=tr.querySelector('.alg').value.split(',').map(Number).filter(n=>n>=1&&n<=19);x.type=tr.querySelector('.type').value;data[tr.dataset.d].edited=true}
function replaceAll(){const a=prompt('찾을 내용'),b=a!==null?prompt('바꿀 내용',''):null;if(a===null||b===null)return;Object.values(data).forEach(d=>d.items.forEach(i=>{i.name=i.name.split(a).join(b)}));renderEdit()}
function renderDetail(){if(!date)setDefaultDate();const x=data[date];$('#content').innerHTML=`<section class="bg-white border border-gray-200 rounded-3xl p-6 shadow-sm"><div class="flex items-center justify-between mb-4"><h2 class="text-xl font-black">${date} 상세 편집</h2><button id="detailSave" class="bg-brand-500 hover:bg-brand-600 text-white rounded-xl px-4 py-2.5 text-sm font-black">이 날짜 저장</button></div><textarea id="note" class="w-full min-h-32 border-2 border-gray-200 rounded-2xl p-4 text-sm" placeholder="오늘의 안내문">${esc(x?.note||'')}</textarea><p class="text-xs text-gray-400 font-bold mt-3">메뉴 추가·삭제와 세부 편집은 월 전체 편집에서 진행합니다.</p></section>`;$('#detailSave').onclick=()=>{x.note=$('#note').value;x.edited=true;saveProject()}}
async function saveProject(){$$('tbody tr').forEach(applyRow);const password=prompt('이 학교의 편집 비밀번호를 입력하세요. 처음 저장할 때 설정됩니다.');if(!password)return;const r=await fetch('/api/project',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({office:school.officeCode,school:school.schoolCode,schoolName:school.schoolName,month,password,payload:{days:data}})}),j=await r.json();alert(r.ok?'저장되었습니다.':j.error)}

/* ═══════ 알레르기 모달 (기존 유지) ═══════ */
function openFilter(){const m=$('#modal');m.classList.remove('hidden');m.classList.add('flex');
m.innerHTML=`<div class="bg-white rounded-3xl p-7 w-full max-w-2xl max-h-[92vh] overflow-auto relative">
<button id="filterClose" class="absolute right-4 top-3 text-gray-400 text-3xl leading-none px-2">×</button>
<h2 class="text-xl font-black text-center mb-1">알레르기 필터</h2>
<p class="text-center text-gray-400 text-sm font-bold mb-3">피해야 하는 식품을 누르면 즉시 선택됩니다.</p>
<div id="selectedCount" class="text-center bg-slate-100 rounded-xl py-2.5 mb-4 text-sm font-black text-gray-500"></div>
<div class="grid grid-cols-3 sm:grid-cols-5 gap-2">
${Object.entries(A).map(([n,v])=>`<button type="button" data-a="${n}" class="flex items-center justify-center gap-1.5 min-h-12 border-2 rounded-xl py-2 transition-all ${allergies.has(+n)?'border-red-400 bg-red-50 text-red-700':'border-gray-200 bg-slate-50'}"><b class="inline-grid place-items-center w-6 h-6 rounded-full text-xs ${allergies.has(+n)?'bg-red-600 text-white':'bg-gray-200'}">${n}</b><span class="text-xs font-bold">${v}</span></button>`).join('')}
</div>
<div class="grid grid-cols-3 gap-2 mt-5">
  <button id="clearAll" class="bg-gray-100 hover:bg-gray-200 rounded-xl py-3 font-black text-sm">전체 해제</button>
  <button id="done" class="col-span-2 bg-brand-500 hover:bg-brand-600 text-white rounded-xl py-3 font-black text-sm">적용하고 닫기</button>
</div></div>`;
const refresh=()=>{allergyOn=allergies.size>0;$('#selectedCount').textContent=allergies.size?`선택 ${allergies.size}개 · 필터 켜짐`:'선택 없음 · 필터 꺼짐'};refresh();
$$('[data-a]').forEach(b=>b.onclick=()=>{const n=+b.dataset.a;allergies.has(n)?allergies.delete(n):allergies.add(n);openFilter()});
$('#clearAll').onclick=()=>{allergies.clear();openFilter()};
const closeAndApply=()=>{allergyOn=allergies.size>0;localStorage.setItem('allergies',JSON.stringify([...allergies]));localStorage.setItem('allergyOn',allergyOn?'1':'0');m.classList.add('hidden');m.classList.remove('flex');renderTab()};
$('#done').onclick=closeAndApply;$('#filterClose').onclick=closeAndApply}

/* ═══════ 사이니지·공유 모달 (학생/학부모 탭 추가) ═══════ */
function openSignage(){const m=$('#modal');m.classList.remove('hidden');m.classList.add('flex');
m.innerHTML=`<div class="bg-white rounded-3xl p-7 w-full max-w-2xl max-h-[92vh] overflow-auto text-center">
<h2 class="text-xl font-black mb-5">사이니지·공유 설정</h2>
<div class="grid grid-cols-2 sm:grid-cols-4 gap-2.5 mb-3">
  ${[['tv','📺','TV 16:9'],['standby','🖥️','스탠바이미'],['kiosk','📱','키오스크 9:16'],['student','👨‍👩‍👧','학생/학부모 전송']].map(x=>`
  <button data-device="${x[0]}" class="border-2 border-gray-200 hover:border-brand-300 rounded-2xl py-4 px-2 font-black text-sm transition-all ${x[0]==='student'?'bg-brand-50 border-brand-200 text-brand-700':''}"><span class="text-2xl block mb-1">${x[1]}</span>${x[2]}</button>`).join('')}
</div>
<div id="signTools"></div>
<button id="close" class="w-full mt-3 bg-gray-100 hover:bg-gray-200 rounded-xl py-3 font-black text-sm">닫기</button></div>`;
$$('[data-device]').forEach(b=>b.onclick=()=>showSignTools(b.dataset.device));
$('#close').onclick=()=>{m.classList.add('hidden');m.classList.remove('flex')}}
async function showSignTools(type){
const isStudent=type==='student';
const url=isStudent
  ?`${location.origin}${location.pathname}?mode=student&office=${school.officeCode}&school=${school.schoolCode}&name=${encodeURIComponent(school.schoolName)}&month=${month}`
  :`${location.origin}${location.pathname}?mode=signage&type=${type}&office=${school.officeCode}&school=${school.schoolCode}&name=${encodeURIComponent(school.schoolName)}&month=${month}`;
$('#signTools').innerHTML=`
${isStudent?`<div class="bg-brand-50 border border-brand-100 rounded-2xl p-4 mb-3 text-left"><b class="text-sm font-black text-brand-800 block mb-1">📱 학생/학부모용 모바일 화면</b><p class="text-xs text-brand-700 font-bold">QR을 급식실·가정통신문에 넣거나, 링크를 학급 단체방에 공유하세요. 학생이 스스로 알레르기를 선택해 안전하게 확인할 수 있어요.</p></div>`:''}
<canvas id="qr" class="mx-auto my-3"></canvas>
<input id="url" value="${esc(url)}" readonly class="w-full border-2 border-gray-200 rounded-xl px-3 py-2.5 text-xs mb-2">
<a href="${url}" target="_blank" class="block bg-brand-500 hover:bg-brand-600 text-white rounded-xl py-3 font-black text-sm no-underline transition-all">${isStudent?'모바일 화면 미리보기':'전체화면 미리보기'}</a>
<button id="copy" class="w-full mt-2 bg-gray-100 hover:bg-gray-200 rounded-xl py-2.5 font-black text-sm">URL 복사</button>`;
await QRCode.toCanvas($('#qr'),url,{width:180});
$('#copy').onclick=()=>{navigator.clipboard.writeText(url);alert('링크가 복사되었습니다!')}}

/* ═══════ 사이니지 모드 ═══════ */
async function bootSignage(){month=p.get('month')||month;document.body.className=`bg-slate-50 signage-${p.get('type')||'tv'}`;
app.innerHTML=`<div class="pub-wrap max-w-7xl mx-auto px-4 py-5">
<header class="bg-white border-b-4 border-brand-500/30 rounded-3xl shadow-sm px-6 py-4 flex items-center gap-3">
  <div class="w-14 h-14 rounded-2xl bg-brand-50 flex items-center justify-center text-3xl border border-brand-100">🥗</div>
  <div class="flex-1"><h1 class="text-xl font-black">${esc(school.schoolName)} 스마트 급식실</h1><p class="text-xs text-gray-400 font-bold">건강한 선택으로 즐거운 점심시간을 만들어요!</p></div>
  <button id="publicFilter" class="bg-red-50 text-red-700 font-black px-4 py-2.5 rounded-xl text-sm border-2 border-red-200">🚨 알레르기 필터</button>
</header>
<nav class="flex gap-2 mt-4 bg-white border border-gray-200 rounded-2xl p-2">
  <button data-public="day" class="flex-1 py-3 rounded-xl font-black text-sm bg-brand-500 text-white">오늘의 식단</button>
  <button data-public="week" class="flex-1 py-3 rounded-xl font-black text-sm text-gray-400">주간 식단</button>
  <button data-public="month" class="flex-1 py-3 rounded-xl font-black text-sm text-gray-400">월간 식단</button>
</nav>
<main id="content" class="mt-4"><div class="bg-white rounded-3xl p-16 text-center text-gray-400">식단을 불러오는 중입니다.</div></main>
<div id="modal" class="fixed inset-0 bg-slate-900/70 z-50 hidden items-center justify-center p-5"></div></div>`;
tab='day';$('#publicFilter').onclick=openFilter;
$$('[data-public]').forEach(b=>b.onclick=()=>{tab=b.dataset.public;$$('[data-public]').forEach(x=>{x.classList.toggle('bg-brand-500',x===b);x.classList.toggle('text-white',x===b);x.classList.toggle('text-gray-400',x!==b)});renderTab()});
await loadApi()}

/* ═══════ 학생/학부모 모바일 모드 ═══════ */
async function bootStudent(){month=p.get('month')||month;document.body.className='bg-slate-50';
app.innerHTML=`<div class="max-w-md mx-auto px-4 py-5 pb-24">
<header class="bg-white border-b-4 border-brand-500/30 rounded-3xl shadow-sm px-5 py-4 flex items-center gap-3">
  <div class="w-12 h-12 rounded-2xl bg-brand-50 flex items-center justify-center text-2xl border border-brand-100">🥗</div>
  <div class="flex-1 min-w-0"><h1 class="text-base font-black leading-tight">${esc(school.schoolName)}</h1><p class="text-[11px] text-gray-400 font-bold">오늘의 급식 · 학생/학부모용</p></div>
</header>
<div class="mt-3 flex gap-2">
  <button data-st="day" class="flex-1 py-3 rounded-2xl font-black text-sm bg-brand-500 text-white shadow-md">오늘 급식</button>
  <button data-st="month" class="flex-1 py-3 rounded-2xl font-black text-sm bg-white border-2 border-gray-200 text-gray-400">월간 보기</button>
</div>
<button id="stFilter" class="w-full mt-2 bg-red-50 border-2 border-red-200 text-red-700 rounded-2xl py-3 font-black text-sm">🚨 내 알레르기 선택 ${allergies.size?`(${allergies.size}개 적용 중)`:''}</button>
<main id="content" class="mt-3"><div class="bg-white rounded-3xl p-12 text-center text-gray-400 text-sm">불러오는 중…</div></main>
<div id="modal" class="fixed inset-0 bg-slate-900/70 z-50 hidden items-center justify-center p-5"></div></div>`;
tab='day';
$('#stFilter').onclick=openFilter;
$$('[data-st]').forEach(b=>b.onclick=()=>{tab=b.dataset.st;$$('[data-st]').forEach(x=>{const on=x===b;x.className=on?'flex-1 py-3 rounded-2xl font-black text-sm bg-brand-500 text-white shadow-md':'flex-1 py-3 rounded-2xl font-black text-sm bg-white border-2 border-gray-200 text-gray-400'});renderTab()});
await loadApi()}

function icon(n){if(/밥|덮밥/.test(n))return'🍚';if(/국|탕|찌개|스프|면/.test(n))return'🥣';if(/김치|깍두기/.test(n))return'🌶️';if(/샐러드|나물|채소|쌈/.test(n))return'🥗';if(/과일|주스|요구르트|아이스크림|빵/.test(n))return'🍎';return'🍽️'}
