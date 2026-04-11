/* ============================================================
   NMA GRADE Tool - Core Logic v2.2
   js/core.js: tool.html の共通ロジック
   ============================================================ */

// ===== DATA =====
let projectData={settings:{},comparisons:{},rawStats:{},transitivity:"",directAssessments:{},indirectAssessments:{},nmaAssessments:{},finalAssessments:{},treatmentNames:{},previousAssessments:null};
let network=null;

// ===== UTILITIES =====
const resolveName=k=>{if(!k)return k;const p=k.split(' vs ');return p.length===2?`${projectData.treatmentNames[p[0]]||p[0]} vs ${projectData.treatmentNames[p[1]]||p[1]}`:k;};
const getKey=(t1,t2)=>{if(!t1||!t2)return null;const a=`${t1} vs ${t2}`,b=`${t2} vs ${t1}`;return projectData.comparisons[a]?a:projectData.comparisons[b]?b:null;};

function showTab(id){
    document.querySelectorAll('.tab-content').forEach(t=>{t.classList.add('hidden');t.classList.remove('active');});
    document.querySelectorAll('.tab-btn').forEach(b=>b.classList.remove('active'));
    document.getElementById(id).classList.remove('hidden');
    document.getElementById(id).classList.add('active');
    document.querySelectorAll('.tab-btn').forEach(b=>{if(b.getAttribute('onclick').includes(id))b.classList.add('active');});
    if(id==='direct')renderDirect();
    if(id==='indirect'){renderIndirect();setTimeout(drawGraph,300);}
    if(id==='network')renderNMA();
    if(id==='imprecision')renderImprecision();
    if(id==='summary')generateSummary();
}

function toggleCollapsible(el){
    const c=el.nextElementSibling;
    c.classList.toggle('show');
    el.querySelector('span:last-child').textContent=c.classList.contains('show')?'▲':'▼';
}

// ===== CERTAINTY =====
function calcCert(d){d=Math.abs(d);if(d<0.25)return{level:'高',cls:'grade-high',sym:'⊕⊕⊕⊕'};if(d<=1)return{level:'中',cls:'grade-moderate',sym:'⊕⊕⊕◯'};if(d<=2)return{level:'低',cls:'grade-low',sym:'⊕⊕◯◯'};return{level:'非常に低',cls:'grade-very-low',sym:'⊕◯◯◯'};}
function gradeNum(l){return l==='高'?4:l==='中'?3:l==='低'?2:1;}
function numToGrade(n){if(n>=4)return{level:'高',cls:'grade-high',sym:'⊕⊕⊕⊕'};if(n>=3)return{level:'中',cls:'grade-moderate',sym:'⊕⊕⊕◯'};if(n>=2)return{level:'低',cls:'grade-low',sym:'⊕⊕◯◯'};return{level:'非常に低',cls:'grade-very-low',sym:'⊕◯◯◯'};}
function gradeColor(v){return v==0?'':v<=1?'grade-moderate':'grade-low';}
function badge(id,c){const el=document.getElementById(id);if(el){el.textContent=c.level;el.className=`certainty-badge ${c.cls}`;}}

// Defaults
const DA=n=>projectData.directAssessments[n]||{rob:0,inconsistency:0,indirectness:0,publicationBias:0,totalDowngrade:0,certainty:{level:'高',cls:'grade-high',sym:'⊕⊕⊕⊕'}};
const IA=n=>projectData.indirectAssessments[n]||{intransitivity:0,baseDowngrade:0,totalDowngrade:0,certainty:{level:'高',cls:'grade-high',sym:'⊕⊕⊕⊕'}};
const NA=n=>projectData.nmaAssessments[n]||{startPoint:'direct',incoherence:0,totalDowngrade:0,certainty:{level:'高',cls:'grade-high',sym:'⊕⊕⊕⊕'}};

// ===== 15-SCENARIO STARTING POINT (Supp file 3) =====
function autoStartingPoint(name){
    const d=projectData.comparisons[name];
    const hasDirect=!!(d.direct||d.raw), hasIndirect=!!(d.indirect||(d.nma&&!d.direct));
    const da=DA(name), ia=IA(name);
    const dCert=gradeNum(da.certainty.level), iCert=gradeNum(ia.certainty.level);

    if(hasDirect&&!hasIndirect) return {start:'direct',scenario:'S1',reason:'直接のみ'};
    if(!hasDirect&&hasIndirect) return {start:'indirect',scenario:'S15',reason:'間接のみ'};

    let dirDominant=false, indDominant=false, similar=false;
    if(d.direct&&d.indirect){
        const dW=d.direct.ciUpper-d.direct.ciLower;
        const iW=d.indirect.ciUpper-d.indirect.ciLower;
        const ratio=dW>0?iW/dW:1;
        if(ratio>1.5) dirDominant=true;
        else if(ratio<0.67) indDominant=true;
        else similar=true;
    } else if(d.direct&&d.nma){
        dirDominant=true;
    }

    if(dCert===iCert) return {start:'direct',scenario:'S14',reason:`直接=間接(${da.certainty.level})`};

    if(dCert>iCert){
        if(dCert===4){
            if(dirDominant||similar) return {start:'direct',scenario:'S2-S3',reason:'直接=高&支配的/同程度'};
            if(indDominant) return {start:'indirect',scenario:'S4',reason:'直接=高だが間接が支配的'};
            return {start:'direct',scenario:'S5',reason:'直接=高&どちらも非支配的'};
        }
        if(dirDominant||similar) return {start:'direct',scenario:'S6-S7',reason:'直接が高い(非最高)&支配的/同程度'};
        if(indDominant) return {start:'indirect',scenario:'S8',reason:'直接が高いが間接が支配的'};
        return {start:'direct',scenario:'S9',reason:'直接が高い&どちらも非支配的'};
    }

    if(indDominant||similar) return {start:'indirect',scenario:'S10-S12',reason:'間接が高い&支配的/同程度'};
    if(dirDominant) return {start:'direct',scenario:'S11',reason:'間接が高いが直接が支配的'};
    return {start:'direct',scenario:'S13',reason:'間接が高い&どちらも非支配的→直接'};
}

// ===== EFFICIENCY RULE: Skip indirect =====
function canSkipIndirect(name){
    const d=projectData.comparisons[name], da=DA(name);
    if(!d.direct) return false;
    if(gradeNum(da.certainty.level)<4) return false;
    if(d.direct&&d.indirect){
        const dW=d.direct.ciUpper-d.direct.ciLower;
        const iW=d.indirect.ciUpper-d.indirect.ciLower;
        if(dW<=iW) return true;
    }
    if(d.direct&&!d.indirect) return true;
    return false;
}

// ===== AUTO IMPRECISION =====
function autoImprecision(name){
    const fa=projectData.finalAssessments[name]||{selection:'nma'};
    const d=projectData.comparisons[name];
    let est=fa.selection==='direct'?d.direct:fa.selection==='indirect'?d.indirect:d.nma;
    if(!est) return {level:0,reason:'データなし',label:'N/A'};
    const ot=document.getElementById('outcomeType').value;
    const sm=parseFloat(document.getElementById('midThreshold').value)||20;
    const lg=parseFloat(document.getElementById('largeThreshold').value)||100;
    const br=parseFloat(document.getElementById('baselineRisk').value)||0.13;

    if(ot==='continuous'){
        const w=Math.abs(est.ciUpper-est.ciLower);
        if(w>=sm*4) return {level:3,reason:'CI極広(MID×4以上)',label:'-3'};
        if(w>=sm*2) return {level:2,reason:'CI広(MID×2以上)',label:'-2'};
        if(est.ciLower<=0&&est.ciUpper>=0) return {level:1,reason:'CIがnull越え',label:'-1'};
        return {level:0,reason:'精確',label:'0'};
    }
    const rd=or=>(or*br/(1-br+or*br)-br)*1000;
    const rdL=rd(est.ciLower),rdH=rd(est.ciUpper);
    const mn=Math.min(rdL,rdH),mx=Math.max(rdL,rdH);
    const ciR=(est.ciLower>0&&isFinite(est.ciUpper))?est.ciUpper/est.ciLower:Infinity;
    if(ciR>=30||!isFinite(ciR)) return {level:3,reason:`CI比${isFinite(ciR)?ciR.toFixed(0):'∞'}≥30`,label:'-3'};
    const crossNull=mn<0&&mx>0;
    const crossLg=Math.abs(mn)<lg&&Math.abs(mx)>=lg;
    if(crossLg&&crossNull) return {level:2,reason:'null+大効果閾値越え',label:'-2'};
    if(ciR>=3){return crossNull?{level:2,reason:`null越え+CI比${ciR.toFixed(1)}≥3`,label:'-2'}:{level:1,reason:`CI比${ciR.toFixed(1)}≥3(OIS不足)`,label:'-1'};}
    if(crossNull) return {level:1,reason:`null越え(RD:${Math.round(mn)}~${Math.round(mx)})`,label:'-1'};
    if(Math.abs(mn)<sm&&Math.abs(mx)>=sm) return {level:1,reason:'MID閾値越え',label:'-1'};
    return {level:0,reason:'閾値不越',label:'0'};
}

// ===== AUTO BEST ESTIMATE (Step D) =====
function autoBest(name){
    const na=NA(name);
    if(na.incoherence>=1){
        const d=projectData.comparisons[name];
        let best='nma',bestN=Math.max(1,gradeNum(na.certainty.level)-autoImprecision(name).level);
        if(d.direct){const dn=Math.max(1,gradeNum(DA(name).certainty.level)-autoImpForEst(d.direct).level);if(dn>bestN){best='direct';bestN=dn;}}
        if(d.indirect){const iN=Math.max(1,gradeNum(IA(name).certainty.level)-autoImpForEst(d.indirect).level);if(iN>bestN){best='indirect';bestN=iN;}}
        return {sel:best,reason:`非整合性あり→最高(${best})`};
    }
    return {sel:'nma',reason:'非整合性なし→NMA'};
}

function autoImpForEst(est){
    if(!est) return {level:0};
    const br=parseFloat(document.getElementById('baselineRisk').value)||0.13;
    const ciR=(est.ciLower>0&&isFinite(est.ciUpper))?est.ciUpper/est.ciLower:Infinity;
    if(ciR>=30) return {level:3}; if(ciR>=3) return {level:2};
    const rd=or=>(or*br/(1-br+or*br)-br)*1000;
    const mn=Math.min(rd(est.ciLower),rd(est.ciUpper)),mx=Math.max(rd(est.ciLower),rd(est.ciUpper));
    if(mn<0&&mx>0) return {level:1}; return {level:0};
}

// ===== DOWNGRADE REASONS =====
function getReasons(n){
    const r=[], da=DA(n), ia=IA(n), na=NA(n), fa=projectData.finalAssessments[n];
    if(da.rob>=1) r.push(`RoB-${da.rob}`);
    if(da.inconsistency>=1) r.push(`非一貫-${da.inconsistency}`);
    if(da.indirectness>=1) r.push(`非直接-${da.indirectness}`);
    if(da.publicationBias>=1) r.push(`出版Bias-${da.publicationBias}`);
    if(ia.intransitivity>=1) r.push(`非推移-${ia.intransitivity}`);
    if(na.incoherence>=1) r.push(`非整合-${na.incoherence}`);
    if(fa&&fa.imprecision>=1) r.push(`不精確-${fa.imprecision}`);
    return r;
}

// ===== LIVING SR: STATUS =====
function getStatus(name){
    if(!projectData.previousAssessments) return 'new';
    if(projectData.previousAssessments[name]) return 'updated';
    return 'new';
}

// ===== PARSING =====
function ensureAll(){const ts=Object.keys(projectData.treatmentNames);ts.sort();for(let i=0;i<ts.length;i++)for(let j=i+1;j<ts.length;j++){if(!getKey(ts[i],ts[j]))projectData.comparisons[`${ts[i]} vs ${ts[j]}`]={treat1:ts[i],treat2:ts[j]};}}
function parseEff(s){if(!s||s.trim()==='-'||s.trim()==='--'||!s.trim())return null;const m=s.replace(/∞/g,'Infinity').replace(/>999/g,'Infinity').match(/([<>]?-?[\d\.]+|Infinity)\s*\(\s*([<>]?-?[\d\.]+|Infinity)\s*,\s*([<>]?-?[\d\.]+|Infinity)\s*\)/);if(m){const p=v=>v==='Infinity'||v.includes('Inf')?Infinity:parseFloat(v.replace(/[<>]/g,''));return{effect:p(m[1]),ciLower:p(m[2]),ciUpper:p(m[3])};}return null;}

function parseIntegratedTable(){
    const input=document.getElementById('integratedTableInput').value;if(!input.trim()){alert('入力してください');return;}
    let count=0;const ep=/(\s+([<>]?-?[\d\.]+|Infinity)\s*\([^\)]+\)|\s+-{1,2}|\s+NA)$/;
    input.split('\n').forEach(line=>{
        line=line.trim();if(!line||line.toLowerCase().startsWith('comp'))return;
        const nm=line.match(ep);if(!nm)return;let rem=line.substring(0,nm.index);
        const im=rem.match(ep);if(!im)return;rem=rem.substring(0,im.index);
        const dm=rem.match(ep);if(!dm)return;rem=rem.substring(0,dm.index);
        const nn=rem.match(/\s+([≥<>]?\d+)$/);if(nn)rem=rem.substring(0,nn.index);
        const parts=rem.trim().split(/\s+vs\.?\s+/i);if(parts.length!==2)return;
        const t1=parts[0].trim(),t2=parts[1].trim();if(t1===t2)return;
        projectData.treatmentNames[t1]=t1;projectData.treatmentNames[t2]=t2;
        let cn=getKey(t1,t2)||`${t1} vs ${t2}`;
        if(!projectData.comparisons[cn])projectData.comparisons[cn]={treat1:t1,treat2:t2};
        const net=parseEff(nm[1]),ind=parseEff(im[1]),dir=parseEff(dm[1]);
        if(net)projectData.comparisons[cn].nma={...net,treat1:t1,treat2:t2};
        if(ind)projectData.comparisons[cn].indirect=ind;
        if(dir)projectData.comparisons[cn].direct=dir;
        count++;
    });
    updateStats();showResult(`✅ ${count}件読込`);
}

function parseNMAInput(){
    const input=document.getElementById('sumPmaInput').value;if(!input.trim()){alert('入力');return;}let count=0;
    input.split('\n').forEach(line=>{line=line.trim();if(!line)return;const p=line.replace(/"/g,'').split(',');if(p.length<4)return;const e=parseFloat(p[1]);if(isNaN(e))return;
        const m=p[0].trim().split(/\s+vs\.?\s+/i);if(m.length!==2)return;const t1=m[0].trim(),t2=m[1].trim();if(t1===t2)return;
        projectData.treatmentNames[t1]=t1;projectData.treatmentNames[t2]=t2;
        let cn=getKey(t1,t2)||`${t1} vs ${t2}`;if(!projectData.comparisons[cn])projectData.comparisons[cn]={treat1:t1,treat2:t2};
        projectData.comparisons[cn].nma={effect:e,ciLower:parseFloat(p[2]),ciUpper:parseFloat(p[3]),treat1:t1,treat2:t2};count++;});
    updateStats();showResult(`✅ ${count}件`);
}

function parseSidesplitOutput(){
    const output=document.getElementById('sidesplitOutput').value;if(!output.trim()){alert('入力');return;}
    let sec='',count=0;const temp={},coding={};
    output.split('\n').forEach(line=>{
        line=line.trim();
        if(line.match(/^Coding:/i)){sec='coding';return;}
        if(line.match(/^Direct\s+evidence:/i)){sec='direct';return;}
        if(line.match(/^Indirect\s+evidence:/i)){sec='indirect';return;}
        if(line.match(/^Difference/i)){sec='diff';return;}
        if(!line||line.startsWith('Ref')||line.startsWith('Call')||line.match(/Est\./))return;
        if(sec==='coding'){const m=line.match(/^(\d+)[:.]?\s+(.+)$/);if(m){coding[m[1]]=m[2].trim();projectData.treatmentNames[m[1]]=m[2].trim();}return;}
        if(['direct','indirect','diff'].includes(sec)){
            const m=line.match(/(\w+)\s+vs\.\s+(\w+)\s+([-\d\.]+)\s+([-\d\.]+)\s+([-\d\.]+)\s+([-\d\.]+)\s+([-\d\.]+)/);
            if(m){const ik=`${m[1]}-${m[2]}`;if(!temp[ik])temp[ik]={id1:m[1],id2:m[2]};
                const dd={effect:parseFloat(m[3]),ciLower:parseFloat(m[5]),ciUpper:parseFloat(m[6]),pvalue:parseFloat(m[7])};
                if(sec==='direct')temp[ik].direct=dd;if(sec==='indirect')temp[ik].indirect=dd;if(sec==='diff')temp[ik].diff=dd;}
        }
    });
    Object.values(temp).forEach(item=>{
        const t1=coding[item.id1]||item.id1,t2=coding[item.id2]||item.id2;
        let cn=getKey(t1,t2)||`${t1} vs ${t2}`;
        if(!projectData.comparisons[cn])projectData.comparisons[cn]={treat1:t1,treat2:t2};
        projectData.comparisons[cn].hasSidesplit=true;
        if(item.direct)projectData.comparisons[cn].direct=item.direct;
        if(item.indirect)projectData.comparisons[cn].indirect=item.indirect;
        if(item.diff)projectData.comparisons[cn].pvalue=item.diff.pvalue;
        count++;
    });
    updateStats();showResult(`✅ SideSplit ${count}件`);
}

function saveTransitivity(){projectData.transitivity=document.getElementById('transitivityInput').value;alert('保存完了');}
function showResult(msg){document.getElementById('parsedResults').classList.remove('hidden');document.getElementById('parsedResultsContent').innerHTML=`<div class="text-green-700 text-sm">${msg}</div>`;}

// ===== RENDER HELPERS =====
function sel(name,field,value,type){
    const m=type==='direct'?'updD':type==='indirect'?'updI':type==='nma'?'updN':'updImp';
    return `<select onchange="${m}('${name}','${field}',this.value)" class="w-full p-0.5 border rounded ${gradeColor(value)} text-xs">
        <option value="0" ${value==0?'selected':''}>0</option><option value="1" ${value==1?'selected':''}>-1</option><option value="2" ${value==2?'selected':''}>-2</option>
        ${type==='imprecision'?'<option value="3" '+(value==3?'selected':'')+'>-3</option>':''}
    </select>`;
}
function fmtE(d){return d?`<span class="estimate-val">${d.effect.toFixed(2)}</span><br><span class="ci-val">[${d.ciLower.toFixed(2)},${d.ciUpper.toFixed(2)}]</span>`:'<span class="text-gray-400 text-xs">—</span>';}
function fmtB(a){return `<span class="certainty-badge ${a.certainty.cls}">${a.certainty.level}</span>`;}
function absEffect(name,src){
    const d=projectData.comparisons[name];let e=src==='direct'?d.direct:src==='indirect'?d.indirect:d.nma;if(!e)return'—';
    if(document.getElementById('outcomeType').value==='continuous')return`MD:${e.effect.toFixed(2)}`;
    const br=parseFloat(document.getElementById('baselineRisk').value)||0.13;
    const r=or=>(or*br/(1-br+or*br)-br)*1000;
    const a=r(e.effect),al=r(e.ciLower),ah=r(e.ciUpper);
    return`${Math.abs(Math.round(a))}${a>0?'↑':'↓'}/1000<br><span class="text-xs">(${Math.round(Math.min(al,ah))}~${Math.round(Math.max(al,ah))})</span>`;
}

// ===== RENDER: A1 =====
function renderDirect(){
    const c=document.getElementById('directAssessmentTableContainer');
    const comps=Object.entries(projectData.comparisons).filter(([,d])=>d.treat1!==d.treat2);
    if(!comps.length){c.innerHTML='<p class="text-gray-500 p-4 text-center">データなし</p>';return;}
    let h=`<table class="assessment-table w-full text-left border-collapse bg-white shadow"><thead><tr>
        <th>比較</th><th>RoB</th><th>非一貫</th><th>非直接</th><th>出版B</th><th>A1仮確実性</th></tr></thead><tbody>`;
    comps.forEach(([n,d])=>{const dn=resolveName(n),a=DA(n);
        const eff=d.direct?`${d.direct.effect.toFixed(2)}[${d.direct.ciLower.toFixed(2)},${d.direct.ciUpper.toFixed(2)}]`:'—';
        h+=`<tr class="border-b"><td><strong class="text-xs">${dn}</strong><br><span class="ci-val">${eff}</span></td>
            <td>${sel(n,'rob',a.rob,'direct')}</td><td>${sel(n,'inconsistency',a.inconsistency,'direct')}</td>
            <td>${sel(n,'indirectness',a.indirectness,'direct')}</td><td>${sel(n,'publicationBias',a.publicationBias,'direct')}</td>
            <td><span id="bd-${n.replace(/\s+/g,'')}" class="certainty-badge ${a.certainty.cls}">${a.certainty.level}</span></td></tr>`;
    });h+='</tbody></table>';c.innerHTML=h;
}

// ===== RENDER: B =====
function findLoops(t1,t2){const r=[];const has=k=>{const d=projectData.comparisons[k];return d&&(d.direct||d.raw||d.nma);};
    Object.keys(projectData.treatmentNames).forEach(c=>{if(c===t1||c===t2)return;const k1=getKey(t1,c),k2=getKey(t2,c);if(k1&&has(k1)&&k2&&has(k2))r.push(c);});return r;}

function renderIndirect(){
    const c=document.getElementById('indirectAssessmentTableContainer');
    const tr=document.getElementById('transitivityReference'),tc=document.getElementById('transitivityContent');
    if(projectData.transitivity&&projectData.transitivity.trim()){tr.classList.remove('hidden');tc.textContent=projectData.transitivity;}else tr.classList.add('hidden');

    let notices='';
    Object.entries(projectData.comparisons).forEach(([n,d])=>{
        if(d.treat1===d.treat2) return;
        if(canSkipIndirect(n)) notices+=`<div class="efficiency-notice"><span class="auto-badge skip">省略可</span> <strong>${resolveName(n)}</strong>: 直接COE=高 & 直接が支配的 → 間接評価不要 <span class="ref">Brignardello-Petersen JCE 2018 Sec 3.2</span></div>`;
    });
    document.getElementById('efficiencyNotices').innerHTML=notices;

    const comps=Object.entries(projectData.comparisons).filter(([,d])=>d.treat1!==d.treat2);
    if(!comps.length){c.innerHTML='<p class="text-gray-500 p-4 text-center">データなし</p>';return;}
    let h=`<table class="assessment-table w-full text-left border-collapse bg-white shadow"><thead><tr>
        <th>比較</th><th>ループ</th><th>非推移性</th><th>B2仮確実性</th></tr></thead><tbody>`;
    comps.forEach(([n,d])=>{const dn=resolveName(n),a=IA(n);
        let eff=d.indirect?`${d.indirect.effect.toFixed(2)}[${d.indirect.ciLower.toFixed(2)},${d.indirect.ciUpper.toFixed(2)}]`:(!d.direct&&d.nma?`${d.nma.effect.toFixed(2)}[NMA]`:'—');
        const loops=findLoops(d.treat1,d.treat2);
        let opts='<option value="">--</option>';loops.forEach(cc=>{opts+=`<option value="${cc}" ${a.commonComparator===cc?'selected':''}>via ${projectData.treatmentNames[cc]||cc}</option>`;});
        if(!loops.length) opts+='<option disabled>なし</option>';
        let cert=a.certainty;if(!a.commonComparator&&!a.baseDowngrade) cert={level:'--',cls:'bg-gray-200 text-gray-500'};
        const skip=canSkipIndirect(n)?'<span class="auto-badge skip ml-1">省略可</span>':'';
        h+=`<tr class="border-b"><td><strong class="text-xs">${dn}</strong>${skip}<br><span class="ci-val">${eff}</span></td>
            <td><select class="w-full p-0.5 border rounded text-xs" onchange="updLoop('${n}',this.value)">${opts}</select>
                <div class="text-xs text-blue-600 mt-0.5">${a.commonComparator?`min(${DA(getKey(d.treat1,a.commonComparator)).certainty.level},${DA(getKey(d.treat2,a.commonComparator)).certainty.level})`:''}</div></td>
            <td>${sel(n,'intransitivity',a.intransitivity,'indirect')}</td>
            <td><span id="bi-${n.replace(/\s+/g,'')}" class="certainty-badge ${cert.cls}">${cert.level}</span></td></tr>`;
    });h+='</tbody></table>';c.innerHTML=h;
}

// ===== RENDER: C1-C2 =====
function renderNMA(){
    const c=document.getElementById('nmaAssessmentTableContainer');
    const comps=Object.entries(projectData.comparisons).filter(([,d])=>d.treat1!==d.treat2);
    if(!comps.length){c.innerHTML='<p class="text-gray-500 p-4 text-center">データなし</p>';return;}
    let h=`<table class="assessment-table w-full text-left border-collapse bg-white shadow"><thead><tr>
        <th>比較</th><th>NMA</th><th>直接</th><th>間接</th><th>C1出発点</th><th>C1シナリオ</th><th>C2非整合性</th><th>NMA確実性</th></tr></thead><tbody>`;
    comps.forEach(([n,d])=>{
        const dn=resolveName(n),na=NA(n),da=DA(n),ia=IA(n);
        const auto=autoStartingPoint(n);
        if(!projectData.nmaAssessments[n]||!projectData.nmaAssessments[n]._manual){
            if(!projectData.nmaAssessments[n]) projectData.nmaAssessments[n]={incoherence:0};
            projectData.nmaAssessments[n].startPoint=auto.start;
            let bd=auto.start==='direct'?DA(n).totalDowngrade:IA(n).totalDowngrade;
            projectData.nmaAssessments[n].totalDowngrade=bd+(projectData.nmaAssessments[n].incoherence||0);
            projectData.nmaAssessments[n].certainty=calcCert(projectData.nmaAssessments[n].totalDowngrade);
        }
        const naR=NA(n);
        h+=`<tr class="border-b"><td class="text-xs font-bold">${dn}</td>
            <td>${d.nma?fmtE(d.nma):'—'}</td>
            <td class="bg-gray-50">${d.direct?`${fmtE(d.direct)}<br>${fmtB(da)}`:'—'}</td>
            <td class="bg-gray-50">${d.indirect?`${fmtE(d.indirect)}<br>${fmtB(ia)}`:'—'}</td>
            <td><select onchange="updNManual('${n}','startPoint',this.value)" class="w-full p-0.5 border rounded text-xs">
                <option value="direct" ${naR.startPoint=='direct'?'selected':''}>直接</option>
                <option value="indirect" ${naR.startPoint=='indirect'?'selected':''}>間接</option></select></td>
            <td><span class="auto-badge auto">${auto.scenario}</span><br><span class="text-xs text-gray-500">${auto.reason}</span></td>
            <td>${sel(n,'incoherence',naR.incoherence,'nma')}${d.pvalue?`<span class="text-xs text-gray-400">P=${d.pvalue}</span>`:''}</td>
            <td><span id="bn-${n.replace(/\s+/g,'')}" class="certainty-badge ${naR.certainty.cls}">${naR.certainty.level}</span></td></tr>`;
    });h+='</tbody></table>';c.innerHTML=h;
}

// ===== RENDER: C3 & D =====
function renderImprecision(){
    const c=document.getElementById('imprecisionAssessmentTableContainer');
    const comps=Object.entries(projectData.comparisons).filter(([,d])=>d.treat1!==d.treat2&&d.nma);
    if(!comps.length){c.innerHTML='<p class="text-gray-500 p-4 text-center">データなし</p>';return;}
    let h=`<table class="assessment-table w-full text-left border-collapse bg-white shadow"><thead><tr>
        <th>比較</th><th>直接</th><th>間接</th><th>NMA</th><th>D:採用</th><th>C3:精確性</th><th>自動</th><th>最終</th><th>絶対効果</th></tr></thead><tbody>`;
    comps.forEach(([n,d])=>{
        const dn=resolveName(n),da=DA(n),ia=IA(n),na=NA(n);
        if(!projectData.finalAssessments[n]){const ab=autoBest(n);projectData.finalAssessments[n]={selection:ab.sel,imprecision:0,autoReason:ab.reason};}
        const fa=projectData.finalAssessments[n];
        const ai=autoImprecision(n);
        if(!fa._manual) fa.imprecision=ai.level;
        const s=fa.selection;
        let bg=s==='direct'?gradeNum(da.certainty.level):s==='indirect'?gradeNum(ia.certainty.level):gradeNum(na.certainty.level);
        const fn=Math.max(1,bg-fa.imprecision), fc=numToGrade(fn);
        h+=`<tr class="border-b">
            <td class="text-xs font-bold">${dn}</td>
            <td class="${s=='direct'?'selected-source':''}">${fmtE(d.direct)}${d.direct?'<br>'+fmtB(da):''}</td>
            <td class="${s=='indirect'?'selected-source':''}">${fmtE(d.indirect)}${d.indirect?'<br>'+fmtB(ia):''}</td>
            <td class="${s=='nma'?'selected-source':''}">${fmtE(d.nma)}<br>${fmtB(na)}</td>
            <td><select onchange="updFinal('${n}','selection',this.value)" class="w-full p-0.5 border rounded text-xs font-bold bg-blue-50">
                <option value="nma" ${s=='nma'?'selected':''}>NMA</option><option value="direct" ${s=='direct'?'selected':''}>直接</option><option value="indirect" ${s=='indirect'?'selected':''}>間接</option>
            </select><div class="text-xs text-gray-400 mt-0.5">${fa.autoReason||''}</div></td>
            <td>${sel(n,'imprecision',fa.imprecision,'imprecision')}</td>
            <td><span class="auto-badge auto">${ai.label}</span><div class="text-xs text-gray-500 mt-0.5">${ai.reason}</div></td>
            <td><span class="certainty-badge ${fc.cls}">${fc.level}</span></td>
            <td class="text-xs font-mono">${absEffect(n,s)}</td></tr>`;
    });h+='</tbody></table>';c.innerHTML=h;
}
// Alias for backward compatibility with HTML onclick
const renderImprecisionAssessmentTable = renderImprecision;

// ===== UPDATES =====
function updD(n,f,v){if(!projectData.directAssessments[n])projectData.directAssessments[n]={rob:0,inconsistency:0,indirectness:0,publicationBias:0};projectData.directAssessments[n][f]=parseFloat(v);const a=projectData.directAssessments[n];a.totalDowngrade=(a.rob||0)+(a.inconsistency||0)+(a.indirectness||0)+(a.publicationBias||0);a.certainty=calcCert(a.totalDowngrade);badge(`bd-${n.replace(/\s+/g,'')}`,a.certainty);}
function updLoop(n,cc){if(!cc)return;if(!projectData.indirectAssessments[n])projectData.indirectAssessments[n]={};projectData.indirectAssessments[n].commonComparator=cc;const d=projectData.comparisons[n],k1=getKey(d.treat1,cc),k2=getKey(d.treat2,cc);let g1=4,g2=4;if(k1)g1=gradeNum(DA(k1).certainty.level);if(k2)g2=gradeNum(DA(k2).certainty.level);projectData.indirectAssessments[n].baseDowngrade=4-Math.min(g1,g2);updI(n,'intransitivity',projectData.indirectAssessments[n].intransitivity||0);}
function updI(n,f,v){if(!projectData.indirectAssessments[n])projectData.indirectAssessments[n]={};projectData.indirectAssessments[n][f]=parseFloat(v);const a=projectData.indirectAssessments[n];a.totalDowngrade=(a.baseDowngrade||0)+(a.intransitivity||0);a.certainty=calcCert(a.totalDowngrade);badge(`bi-${n.replace(/\s+/g,'')}`,a.certainty);}
function updN(n,f,v){if(!projectData.nmaAssessments[n])projectData.nmaAssessments[n]={};projectData.nmaAssessments[n][f]=parseFloat(v);const a=projectData.nmaAssessments[n];let bd=a.startPoint==='direct'?DA(n).totalDowngrade:IA(n).totalDowngrade;a.totalDowngrade=bd+(a.incoherence||0);a.certainty=calcCert(a.totalDowngrade);badge(`bn-${n.replace(/\s+/g,'')}`,a.certainty);}
function updNManual(n,f,v){if(!projectData.nmaAssessments[n])projectData.nmaAssessments[n]={};projectData.nmaAssessments[n][f]=v;projectData.nmaAssessments[n]._manual=true;updN(n,'incoherence',projectData.nmaAssessments[n].incoherence||0);}
function updFinal(n,f,v){if(!projectData.finalAssessments[n])projectData.finalAssessments[n]={selection:'nma',imprecision:0};projectData.finalAssessments[n][f]=v;delete projectData.finalAssessments[n]._manual;renderImprecision();}
function updImp(n,f,v){if(!projectData.finalAssessments[n])projectData.finalAssessments[n]={selection:'nma',imprecision:0};projectData.finalAssessments[n][f]=parseFloat(v);projectData.finalAssessments[n]._manual=true;renderImprecision();}

// ===== GRAPH =====
function drawGraph(){
    const c=document.getElementById('network-graph-container');if(!c||c.offsetParent===null)return;
    const nodes=new vis.DataSet(),edges=new vis.DataSet(),ts=new Set();
    const valid=Object.entries(projectData.comparisons).filter(([,d])=>(d.direct||d.raw||d.nma)&&d.treat1!==d.treat2);
    if(!valid.length){c.innerHTML='<p class="text-gray-400 p-8 text-center">データなし</p>';return;}c.innerHTML='';
    valid.forEach(([,d])=>{ts.add(d.treat1);ts.add(d.treat2);let col='#ccc',dash=true,w=2;if(d.direct||d.raw){col='#667eea';dash=false;w=d.raw?Math.max(1,Math.log(d.raw.studies)*2):2;}else{col='#f59e0b';}edges.add({from:d.treat1,to:d.treat2,width:w,color:{color:col},dashes:dash});});
    const arr=[...ts],nc=arr.length,r=140;arr.forEach((t,i)=>{const a=2*Math.PI*i/nc;nodes.add({id:t,label:projectData.treatmentNames[t]||t,x:r*Math.cos(a),y:r*Math.sin(a),shape:'dot',size:14,color:{background:'#ff9f43',border:'#e67e22'},font:{size:12}});});
    if(network)network.destroy();network=new vis.Network(c,{nodes,edges},{physics:false,interaction:{dragNodes:true,zoomView:true},edges:{smooth:false}});
}

// ===== STATS =====
function updateStats(){
    ensureAll();const f=k=>projectData.comparisons[k].treat1!==projectData.comparisons[k].treat2;
    document.getElementById('totalComparisons').textContent=Object.keys(projectData.comparisons).filter(f).length;
    document.getElementById('directCount').textContent=Object.values(projectData.comparisons).filter(d=>d.direct&&d.treat1!==d.treat2).length;
    document.getElementById('indirectCount').textContent=Object.values(projectData.comparisons).filter(d=>d.indirect&&d.treat1!==d.treat2).length;
    document.getElementById('assessedCount').textContent=Object.keys(projectData.finalAssessments).length;
    document.getElementById('skipCount').textContent=Object.keys(projectData.comparisons).filter(k=>f(k)&&canSkipIndirect(k)).length;
}

// ===== SUMMARY =====
function generateSummary(){
    const tbody=document.getElementById('summaryTableBody');
    const comps=Object.keys(projectData.comparisons).filter(k=>projectData.comparisons[k].treat1!==projectData.comparisons[k].treat2);
    if(!comps.length){tbody.innerHTML='<tr><td colspan="10" class="p-4 text-center">データなし</td></tr>';return;}
    let html='';
    comps.forEach(n=>{
        const dn=resolveName(n),d=projectData.comparisons[n];
        const da=DA(n),ia=IA(n),na=NA(n),fa=projectData.finalAssessments[n];
        const nmaE=d.nma?`${d.nma.effect.toFixed(2)}(${d.nma.ciLower.toFixed(2)},${d.nma.ciUpper.toFixed(2)})`:'-';
        const absE=fa?absEffect(n,fa.selection):'—';
        let fc='-',sl='-',reasonsHtml='-';
        const status=getStatus(n);
        const statusBadge=status==='new'?'<span class="status-new">NEW</span>':status==='updated'?'<span class="status-updated">更新</span>':'<span class="status-unchanged">変更なし</span>';
        if(fa){
            let base=fa.selection==='direct'?gradeNum(da.certainty.level):fa.selection==='indirect'?gradeNum(ia.certainty.level):gradeNum(na.certainty.level);
            const fn=Math.max(1,base-(fa.imprecision||0)),fcc=numToGrade(fn);
            fc=`<span class="certainty-badge ${fcc.cls}">${fcc.sym}<br>${fcc.level}</span>`;
            sl=fa.selection==='direct'?'直接':fa.selection==='indirect'?'間接':'NMA';
            const reasons=getReasons(n);
            reasonsHtml=reasons.length?'<div class="reason-tags">'+reasons.map(r=>`<span class="${r.includes('-2')||r.includes('-3')?'bg-red-100 text-red-800':'bg-yellow-100 text-yellow-800'}">${r}</span>`).join('')+'</div>':'<span class="text-green-600 text-xs">なし</span>';
        }
        html+=`<tr class="border-b hover:bg-gray-50"><td class="text-xs">${dn}</td><td>${statusBadge}</td>
            <td class="text-xs font-mono">${nmaE}</td><td class="text-xs">${absE}</td>
            <td>${fmtB(da)}</td><td>${fmtB(ia)}</td><td>${fmtB(na)}</td>
            <td>${fc}</td><td class="text-xs">${sl}</td><td>${reasonsHtml}</td></tr>`;
    });
    tbody.innerHTML=html;
}

// ===== EXPORT / IMPORT =====
function exportJSON(){const d="data:text/json;charset=utf-8,"+encodeURIComponent(JSON.stringify(projectData));const a=document.createElement('a');a.href=d;a.download="nma_grade_v22.json";document.body.appendChild(a);a.click();a.remove();}
function importJSON(e){
    const f=e.target.files[0];if(!f)return;const r=new FileReader();
    r.onload=ev=>{
        try{
            const prev=JSON.parse(JSON.stringify(projectData));
            projectData=JSON.parse(ev.target.result);
            projectData.previousAssessments=prev.finalAssessments;
            updateStats();alert("読込完了（前回評価を引き継ぎ）");
        }catch(err){alert("失敗");}
    };r.readAsText(f);
}
function exportCSV(){
    const comps=Object.keys(projectData.comparisons).filter(k=>projectData.comparisons[k].treat1!==projectData.comparisons[k].treat2);
    let csv='比較,NMA推定値,絶対効果,A1直接,B2間接,C_NMA,C3精確性,最終確実性,採用,ダウングレード理由\n';
    comps.forEach(n=>{const dn=resolveName(n),d=projectData.comparisons[n],da=DA(n),ia=IA(n),na=NA(n),fa=projectData.finalAssessments[n];
        const nE=d.nma?`${d.nma.effect.toFixed(2)}(${d.nma.ciLower.toFixed(2)}-${d.nma.ciUpper.toFixed(2)})`:'-';
        let fc='-',sl='-';if(fa){const base=fa.selection==='direct'?gradeNum(da.certainty.level):fa.selection==='indirect'?gradeNum(ia.certainty.level):gradeNum(na.certainty.level);fc=numToGrade(Math.max(1,base-(fa.imprecision||0))).level;sl=fa.selection;}
        csv+=`"${dn}","${nE}","","${da.certainty.level}","${ia.certainty.level}","${na.certainty.level}","${fa?-fa.imprecision:'-'}","${fc}","${sl}","${getReasons(n).join('; ')}"\n`;
    });
    const blob=new Blob(['\ufeff'+csv],{type:'text/csv;charset=utf-8;'});const a=document.createElement('a');a.href=URL.createObjectURL(blob);a.download='nma_grade_sof.csv';a.click();
}

// ===== TUTORIAL DATA INTEGRATION =====
function checkTutorialData(){
    const raw = localStorage.getItem('nmaGradeTutorialData');
    if(!raw) return;
    const banner = document.getElementById('tutorialImportBanner');
    if(banner) banner.style.display='flex';
}

function importTutorialData(){
    const raw = localStorage.getItem('nmaGradeTutorialData');
    if(!raw) return;
    try{
        const data = JSON.parse(raw);
        projectData = data;
        updateStats();
        document.getElementById('tutorialImportBanner').style.display='none';
        alert('✅ チュートリアルデータを読み込みました。各タブで評価を確認できます。');
    }catch(e){ alert('読込エラー'); }
}

function dismissTutorialBanner(){
    const banner = document.getElementById('tutorialImportBanner');
    if(banner) banner.style.display='none';
    localStorage.removeItem('nmaGradeTutorialData');
}

// ===== INIT =====
document.addEventListener('DOMContentLoaded', () => {
    updateStats();
    checkTutorialData();
});
