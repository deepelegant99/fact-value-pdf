const pptxgen = require('pptxgenjs');
const fs = require('fs');
const path = require('path');

function warnIfSlideHasOverlaps(slide,pptx,options={}){
  const objs = slide._slideObjects || [];
  const get = (o)=> o.options || o.data || {};
  const boxes = objs.map((o,i)=>({i,type:o.type||'object',...get(o)})).filter(b=>Number.isFinite(b.x)&&Number.isFinite(b.y)&&Number.isFinite(b.w)&&Number.isFinite(b.h));
  for(let a=0;a<boxes.length;a++) for(let b=a+1;b<boxes.length;b++){
    const A=boxes[a], B=boxes[b];
    const ix=Math.max(0, Math.min(A.x+A.w,B.x+B.w)-Math.max(A.x,B.x));
    const iy=Math.max(0, Math.min(A.y+A.h,B.y+B.h)-Math.max(A.y,B.y));
    if(ix*iy>0.2 && String(A.type).includes('text') && String(B.type).includes('text'))
      console.warn(`Potential text overlap on slide ${pptx._slides.indexOf(slide)+1}: elements ${A.i} and ${B.i}`);
  }
}
function warnIfSlideElementsOutOfBounds(slide,pptx){
  const objs = slide._slideObjects || [];
  const get = (o)=> o.options || o.data || {};
  objs.forEach((o,i)=>{ const b=get(o); if(Number.isFinite(b.x)&&Number.isFinite(b.y)&&Number.isFinite(b.w)&&Number.isFinite(b.h)){
    if(b.x<-0.02 || b.y<-0.02 || b.x+b.w>13.353 || b.y+b.h>7.52) console.warn(`Out of bounds on slide ${pptx._slides.indexOf(slide)+1}: element ${i}`);
  }});
}

const MD = 'Content/putnam-fact-value-slides.md';
const OUT = 'Presentation/putnam-fact-value-slides.pptx';
const pptx = new pptxgen();
pptx.layout = 'LAYOUT_WIDE';
pptx.author = 'OpenAI';
pptx.subject = 'Putnam fact-value presentation';
pptx.title = 'The Entanglement of Fact and Value';
pptx.company = 'OpenAI';
pptx.lang = 'en-US';
pptx.theme = {
  headFontFace: 'Arial',
  bodyFontFace: 'Arial',
  lang: 'en-US'
};
pptx.defineLayout({ name:'CUSTOM_WIDE', width:13.333, height:7.5 });
pptx.layout = 'CUSTOM_WIDE';

const W=13.333, H=7.5;
const C = {
  ink:'1F2A2B', text:'344246', mut:'69787A', line:'D9E4E2', paper:'FFFFFF', wash:'F4F8F7', wash2:'EEF7F6',
  teal:'087C7E', teal2:'0AA1A2', blue:'4CB6D4', green:'86C440', yellow:'F0C845', orange:'EF8B39', red:'D9674E'
};
const accent = [C.teal,C.blue,C.green,C.orange,C.yellow];
function mdclean(s){ return (s||'').replace(/\[([^\]]+)\]\(([^)]+)\)/g,'$1').replace(/[*`]/g,'').trim(); }
function urls(s){ return [...s.matchAll(/\[([^\]]+)\]\(([^)]+)\)/g)].map(m=>m[2]); }
function host(u){ try { return new URL(u).hostname.replace(/^www\./,''); } catch(e){return '';} }
function parse(){
  const raw=fs.readFileSync(MD,'utf8');
  return raw.split(/\n---\s*\n/).map(chunk=>{
    const m=chunk.match(/^##\s+Slide\s+(\d+)\s+[—-]\s+(.+)$/m);
    if(!m) return null;
    const d={n:+m[1], title:mdclean(m[2]), bullets:[], paras:[], callouts:[], table:[], urls:urls(chunk)};
    const body=chunk.slice(m.index+m[0].length);
    for(const line of body.split('\n')){
      let z=line.trim();
      if(!z) continue;
      if(z.startsWith('|') && z.endsWith('|')){
        const cells=z.slice(1,-1).split('|').map(mdclean);
        if(!cells.every(x=>/^:?-{3,}:?$/.test(x))) d.table.push(cells);
      } else if(z.startsWith('- ')) d.bullets.push(mdclean(z.slice(2)));
      else if(z.startsWith('**') && z.endsWith('**')) d.callouts.push(mdclean(z));
      else if(z.startsWith('*') && z.endsWith('*')) { /* skip image-style notes */ }
      else {
        const q=mdclean(z);
        if(q.endsWith('?') && q.length<150) d.callouts.push(q); else d.paras.push(q);
      }
    }
    return d;
  }).filter(Boolean).sort((a,b)=>a.n-b.n);
}
function section(d){
  if(d.n<=10) return ['ECONOMICS + MEASUREMENT', C.green];
  if(d.n<=19) return ['POLITICS + HISTORY', C.blue];
  if(d.n<=28) return ["PUTNAM'S ARGUMENT", C.teal];
  return ['STAKES + READING', C.orange];
}
function addHeader(slide,d){
  slide.background = { color: C.paper };
  const [s,col]=section(d);
  slide.addShape(pptx.ShapeType.rect,{x:0,y:0,w:W,h:0.18,fill:{color:col},line:{color:col}});
  slide.addText(s,{x:0.55,y:0.36,w:5.5,h:0.23,fontFace:'Arial',fontSize:9,bold:true,color:col,margin:0,fit:'shrink'});
  slide.addText(String(d.n).padStart(2,'0'),{x:12.05,y:0.34,w:0.65,h:0.25,fontSize:9,bold:true,color:C.mut,align:'right',margin:0});
  slide.addText(d.title,{x:0.55,y:0.72,w:11.8,h:0.55,fontSize:d.title.length>62?22:26,bold:true,color:C.ink,margin:0,breakLine:false,fit:'shrink'});
  slide.addShape(pptx.ShapeType.line,{x:0.55,y:1.35,w:1.45,h:0,line:{color:col,width:2.5}});
}
function footer(slide,d){
  const hs=[...new Set(d.urls.map(host).filter(Boolean))];
  if(hs.length) slide.addText('sources: '+hs.join(' · '),{x:0.55,y:7.13,w:9.9,h:0.16,fontSize:6.8,color:C.mut,margin:0,fit:'shrink'});
}
function addCallout(slide,text,col=C.teal,y=6.3){
  if(!text) return;
  slide.addText(text,{x:0.75,y,w:11.85,h:0.52,margin:0.08,fontSize:text.length>110?14.5:17,bold:true,color:'FFFFFF',align:'center',valign:'mid',fill:{color:col},line:{color:col},radius:0.11,fit:'shrink'});
}
function bulletLines(slide,bullets,x,y,w,h,opts={}){
  const fs=opts.fontSize||18;
  const arr=bullets.slice(0,opts.max||4).map(b=>({text:'• '+b, options:{bullet:false,breakLine:true}}));
  slide.addText(arr,{x,y,w,h,fontSize:fs,color:C.text,breakLine:false,fit:'shrink',margin:0.02,paraSpaceAfterPt:9,bold:false});
}
function pill(slide,text,x,y,w,h,col,fs=18){
  slide.addText(text,{x,y,w,h,margin:0.08,fontSize:fs,bold:true,color:'FFFFFF',align:'center',valign:'mid',fill:{color:col},line:{color:col},radius:0.2,fit:'shrink'});
}
function bigBlock(slide,label,body,x,y,w,h,col,fs=20){
  slide.addText(label.toUpperCase(),{x,y,w,h:0.28,fontSize:9,bold:true,color:col,margin:0,fit:'shrink'});
  slide.addText(body,{x,y:y+0.42,w,h:h-0.42,fontSize:fs,bold:false,color:C.text,margin:0.02,breakLine:false,fit:'shrink',valign:'mid'});
}
function twoColumn(slide,d,leftLabel='ONE',rightLabel='TWO'){
  addHeader(slide,d);
  const items = d.bullets.length>=2 ? d.bullets : d.paras;
  bigBlock(slide,leftLabel,items[0]||'',0.85,1.75,5.55,3.45,C.blue,22);
  bigBlock(slide,rightLabel,items[1]||'',6.95,1.75,5.55,3.45,C.orange,22);
  const extra = (items.length>2?items.slice(2):d.paras.slice(2)).join(' ');
  if(extra) slide.addText(extra,{x:1.1,y:5.42,w:11.1,h:0.55,fontSize:15,color:C.mut,align:'center',margin:0,fit:'shrink'});
  addCallout(slide,d.callouts[0],section(d)[1]); footer(slide,d);
}
function listSlide(slide,d){
  addHeader(slide,d);
  if(d.paras[0]) slide.addText(d.paras[0],{x:0.95,y:1.55,w:11.4,h:0.5,fontSize:17,color:C.mut,align:'center',margin:0.02,fit:'shrink'});
  bulletLines(slide,d.bullets,1.35,2.25,10.7,2.9,{fontSize:20,max:4});
  const rest=d.paras.slice(1).join(' ');
  if(rest) slide.addText(rest,{x:1.25,y:5.25,w:10.9,h:0.5,fontSize:15.5,color:C.mut,align:'center',margin:0,fit:'shrink'});
  addCallout(slide,d.callouts[0],section(d)[1]); footer(slide,d);
}
function stepsSlide(slide,d){
  addHeader(slide,d);
  if(d.paras[0]) slide.addText(d.paras[0],{x:0.95,y:1.55,w:11.4,h:0.4,fontSize:16,color:C.mut,align:'center',margin:0,fit:'shrink'});
  const n=Math.min(4,d.bullets.length);
  const gap=0.25, startX=0.78, cardW=(11.8-(n-1)*gap)/n;
  for(let i=0;i<n;i++){
    pill(slide,String(i+1),startX+i*(cardW+gap)+cardW/2-0.27,2.32,0.54,0.54,accent[i],20);
    slide.addText(d.bullets[i],{x:startX+i*(cardW+gap),y:3.1,w:cardW,h:1.45,fontSize:16.5,color:C.text,align:'center',margin:0.05,fit:'shrink'});
  }
  addCallout(slide,d.callouts[0],section(d)[1]); footer(slide,d);
}
function titleSlide(slide,d){
  slide.background={color:C.paper};
  slide.addShape(pptx.ShapeType.rect,{x:0,y:0,w:0.18,h:H,fill:{color:C.teal},line:{color:C.teal}});
  slide.addText('FACT / VALUE',{x:0.75,y:0.65,w:3.2,h:0.25,fontSize:11,bold:true,color:C.teal,margin:0});
  slide.addText(d.title,{x:0.75,y:1.25,w:6.1,h:1.05,fontSize:34,bold:true,color:C.ink,margin:0,fit:'shrink'});
  slide.addText(d.paras.slice(0,2).join(' '),{x:0.78,y:2.7,w:5.8,h:0.88,fontSize:18,color:C.mut,margin:0,fit:'shrink'});
  pill(slide,'FACT',7.35,1.55,4.55,0.9,C.teal,24);
  slide.addText('+',{x:8.95,y:2.7,w:1.35,h:0.58,fontSize:38,bold:true,color:C.ink,align:'center',margin:0});
  pill(slide,'VALUE',7.35,3.58,4.55,0.9,C.green,24);
  slide.addText('not two sealed boxes', {x:7.35,y:4.85,w:4.55,h:0.35,fontSize:15,color:C.mut,align:'center',margin:0});
  slide.addText('Hilary Putnam - The Entanglement of Fact and Value',{x:0.78,y:5.35,w:5.9,h:0.38,fontSize:15,bold:true,color:C.ink,margin:0});
  addCallout(slide,d.callouts[0],C.teal,6.22); footer(slide,d);
}
function progressSlide(slide,d){
  addHeader(slide,d);
  slide.addText('5%',{x:0.9,y:1.72,w:2.8,h:0.8,fontSize:54,bold:true,color:C.green,align:'center',margin:0});
  slide.addText('GDP growth',{x:1.0,y:2.6,w:2.6,h:0.25,fontSize:13,bold:true,color:C.ink,align:'center',margin:0});
  const items=d.bullets.slice(1,4);
  items.forEach((b,i)=>{
    slide.addText(b,{x:4.45,y:1.72+i*0.95,w:7.3,h:0.55,fontSize:18.5,color:C.text,margin:0.06,fill:{color:C.wash},line:{color:C.line},fit:'shrink'});
  });
  addCallout(slide,d.callouts[0],C.teal); footer(slide,d);
}
function allocationSlide(slide,d){
  addHeader(slide,d);
  slide.addText('99',{x:1.0,y:1.72,w:1.6,h:0.7,fontSize:42,bold:true,color:C.teal,margin:0});
  slide.addText('1',{x:11.3,y:1.72,w:0.5,h:0.7,fontSize:42,bold:true,color:C.orange,align:'right',margin:0});
  slide.addShape(pptx.ShapeType.rect,{x:1.0,y:2.67,w:10.65,h:0.5,fill:{color:C.teal},line:{color:C.teal}});
  slide.addShape(pptx.ShapeType.rect,{x:11.65,y:2.67,w:0.18,h:0.5,fill:{color:C.orange},line:{color:C.orange}});
  bulletLines(slide,d.bullets,1.05,3.75,11.1,1.3,{fontSize:18,max:3});
  addCallout(slide,d.callouts[0],C.teal); footer(slide,d);
}
function efficiencySlide(slide,d){
  addHeader(slide,d);
  ['99 / 1','50 / 50'].forEach((lab,j)=>{
    const x=1.0+j*6.0;
    slide.addText(lab,{x,y:1.75,w:4.9,h:0.45,fontSize:28,bold:true,color:j?C.green:C.orange,align:'center',margin:0});
    slide.addText('Pareto efficient',{x,y:2.35,w:4.9,h:0.35,fontSize:17,bold:true,color:C.ink,align:'center',margin:0});
    const a=j?2.25:4.45, b=j?2.25:0.12;
    slide.addShape(pptx.ShapeType.rect,{x:x+0.15,y:3.22,w:a,h:0.42,fill:{color:C.teal},line:{color:C.teal}});
    slide.addShape(pptx.ShapeType.rect,{x:x+0.15+a,y:3.22,w:b,h:0.42,fill:{color:C.orange},line:{color:C.orange}});
  });
  slide.addText(d.paras.join(' '),{x:1.2,y:4.52,w:10.85,h:0.6,fontSize:16,color:C.mut,align:'center',margin:0,fit:'shrink'});
  addCallout(slide,d.callouts[0],C.teal); footer(slide,d);
}
function tableSlide(slide,d){
  addHeader(slide,d);
  const rows=d.table.slice(1,4);
  rows.forEach((r,i)=>{
    const y=1.7+i*1.28;
    slide.addText(['POWER','BELIEF','JUSTIFICATION'][i],{x:0.85,y,w:2.15,h:0.8,fontSize:16,bold:true,color:'FFFFFF',align:'center',valign:'mid',margin:0.05,fill:{color:accent[i]},line:{color:accent[i]},fit:'shrink'});
    slide.addText(r[0],{x:3.25,y,w:4.5,h:0.8,fontSize:17.5,bold:true,color:C.text,margin:0.05,fill:{color:C.wash},line:{color:C.line},fit:'shrink'});
    slide.addText(r[1],{x:8.0,y,w:4.3,h:0.8,fontSize:16,color:C.mut,margin:0.05,fill:{color:C.wash},line:{color:C.line},fit:'shrink'});
  });
  addCallout(slide,d.callouts[0],C.teal); footer(slide,d);
}
function conceptsSlide(slide,d){
  addHeader(slide,d);
  const words=d.bullets.slice(0,3).map(x=>x.replace(/\.$/,''));
  words.forEach((w,i)=> pill(slide,w.toUpperCase(),1.1+i*4.05,2.25,2.85,1.15,accent[[3,1,2][i]],22));
  slide.addText((d.paras[1]||'')+' '+(d.paras[2]||''),{x:1.6,y:4.35,w:10.1,h:0.65,fontSize:18,color:C.text,align:'center',margin:0,fit:'shrink'});
  addCallout(slide,d.callouts[0],C.teal); footer(slide,d);
}
function readingSlide(slide,d){
  addHeader(slide,d);
  slide.addText('HILARY\nPUTNAM',{x:1.05,y:1.65,w:3.6,h:1.15,fontSize:27,bold:true,color:'FFFFFF',align:'center',valign:'mid',margin:0.06,fill:{color:C.ink},line:{color:C.ink},fit:'shrink'});
  slide.addText('The Entanglement\nof Fact and Value',{x:1.05,y:3.1,w:3.6,h:0.85,fontSize:18,bold:true,color:C.green,align:'center',valign:'mid',margin:0.04,fill:{color:C.ink},line:{color:C.ink},fit:'shrink'});
  slide.addText('2002 - Chapter 2 - pp. 28-45',{x:1.05,y:4.22,w:3.6,h:0.28,fontSize:12,color:'DDE6E3',align:'center',margin:0});
  bulletLines(slide,d.bullets,5.35,1.7,6.8,3.45,{fontSize:18.5,max:3});
  addCallout(slide,d.callouts[0],C.teal); footer(slide,d);
}
function generic(slide,d){
  if(d.bullets.length>=4) return listSlide(slide,d);
  if(d.bullets.length>=1) return stepsSlide(slide,d);
  addHeader(slide,d);
  const text=d.paras.join('\n\n');
  slide.addText(text,{x:1.05,y:1.75,w:11.0,h:3.6,fontSize:22,color:C.text,margin:0.02,fit:'shrink',breakLine:false,align:'center',valign:'mid'});
  addCallout(slide,d.callouts[0],section(d)[1]); footer(slide,d);
}
function makeSlide(d){
  const slide=pptx.addSlide();
  if(d.n===1) titleSlide(slide,d);
  else if(d.n===3) progressSlide(slide,d);
  else if(d.n===5 || d.n===6 || d.n===20 || d.n===23 || d.n===26) twoColumn(slide,d,d.n===5?'DESCRIPTIVE':'SIDE A',d.n===5?'EVALUATIVE':'SIDE B');
  else if(d.n===9) allocationSlide(slide,d);
  else if(d.n===10) efficiencySlide(slide,d);
  else if(d.n===14) tableSlide(slide,d);
  else if(d.n===21) conceptsSlide(slide,d);
  else if(d.n===30) readingSlide(slide,d);
  else generic(slide,d);
  warnIfSlideHasOverlaps(slide,pptx,{muteContainment:true,ignoreLines:true,ignoreDecorativeShapes:true});
  warnIfSlideElementsOutOfBounds(slide,pptx);
}
function credits(){
  const s=pptx.addSlide();
  s.background={color:C.paper};
  s.addText('Credits',{x:0.85,y:0.9,w:6,h:0.6,fontSize:32,bold:true,color:C.ink,margin:0});
  s.addText('Original visual direction: Ecology Infographics by Slidesgo.\nDeck rebuilt for readability from the Markdown source in the repository.',{x:0.9,y:2.0,w:10.8,h:1.2,fontSize:19,color:C.text,margin:0.02,fit:'shrink'});
}
const data=parse();
if(data.length!==30) throw new Error('Expected 30 slides, got '+data.length);
data.forEach(makeSlide);
credits();
fs.mkdirSync(path.dirname(OUT),{recursive:true});
pptx.writeFile({ fileName: OUT });
