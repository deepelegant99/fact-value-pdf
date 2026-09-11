const pptxgen = require('pptxgenjs');
const fs = require('fs');

const pptx = new pptxgen();
pptx.defineLayout({ name: 'CUSTOM_WIDE', width: 13.333, height: 7.5 });
pptx.layout = 'CUSTOM_WIDE';
pptx.author = 'Phillip Wong';
pptx.subject = 'Putnam fact-value entanglement YouTube presentation';
pptx.title = 'Can Facts Really Be Separated from Values?';
pptx.company = 'Philosophers and Gamblers';
pptx.lang = 'en-US';
pptx.theme = { headFontFace: 'Aptos Display', bodyFontFace: 'Aptos', lang: 'en-US' };

const W = 13.333, H = 7.5;
const C = {
  bg: 'F7F8F6', paper: 'FFFFFF', ink: '142326', text: '304245', muted: '6B7A7C',
  line: 'D8E2E0', teal: '087C7E', teal2: '10A3A5', blue: '4EA7C4', green: '6DAA45',
  orange: 'E58A3B', red: 'C65C4B', yellow: 'E3BF4A', paleTeal: 'E8F4F3', paleBlue: 'EAF3F7',
  paleGreen: 'EFF6E9', paleOrange: 'FBF0E7', paleRed: 'F8ECE9', dark: '0D1A1C'
};

function addBg(slide, color=C.bg) { slide.background = { color }; }
function addTop(slide, n, title, kicker='PUTNAM · FACT / VALUE') {
  addBg(slide);
  slide.addShape(pptx.ShapeType.rect,{x:0,y:0,w:W,h:0.12,fill:{color:C.teal},line:{color:C.teal}});
  slide.addText(kicker,{x:0.55,y:0.34,w:4.6,h:0.22,fontFace:'Aptos',fontSize:9,bold:true,color:C.teal,margin:0,charSpacing:0.4});
  slide.addText(String(n).padStart(2,'0'),{x:12.15,y:0.32,w:0.6,h:0.22,fontSize:9,bold:true,color:C.muted,align:'right',margin:0});
  slide.addText(title,{x:0.55,y:0.72,w:12.0,h:0.6,fontFace:'Aptos Display',fontSize:27,bold:true,color:C.ink,margin:0,fit:'shrink'});
  slide.addShape(pptx.ShapeType.line,{x:0.55,y:1.40,w:1.55,h:0,line:{color:C.teal,width:2.2}});
}
function addSub(slide, text, y=1.52) {
  slide.addText(text,{x:0.75,y,w:11.8,h:0.38,fontSize:14.5,color:C.muted,align:'center',margin:0,fit:'shrink'});
}
function pill(slide, text, x,y,w,h, color, fs=18) {
  slide.addText(text,{x,y,w,h,fontSize:fs,bold:true,color:'FFFFFF',align:'center',valign:'mid',margin:0.06,
    fill:{color},line:{color},radius:0.14,fit:'shrink'});
}
function card(slide, x,y,w,h, title, body, color=C.teal, fill=C.paper, fs=16.5) {
  slide.addShape(pptx.ShapeType.roundRect,{x,y,w,h,rectRadius:0.08,fill:{color:fill},line:{color:C.line,width:1.1}});
  slide.addShape(pptx.ShapeType.rect,{x,y,w:0.08,h,fill:{color},line:{color}});
  if(title) slide.addText(title,{x:x+0.28,y:y+0.22,w:w-0.5,h:0.36,fontSize:15,bold:true,color,margin:0,fit:'shrink'});
  slide.addText(body,{x:x+0.28,y:y+0.72,w:w-0.55,h:h-0.9,fontSize:fs,color:C.text,margin:0.02,breakLine:false,fit:'shrink',valign:'mid'});
}
function arrow(slide, x1,y1,x2,y2,color=C.teal,width=2.2,end='triangle') {
  slide.addShape(pptx.ShapeType.line,{x:x1,y:y1,w:x2-x1,h:y2-y1,line:{color,width,beginArrowType:'none',endArrowType:end}});
}
function returnArrow(slide, xLeft,y,xRight,color=C.teal,width=2.2) {
  slide.addShape(pptx.ShapeType.line,{x:xLeft,y,w:xRight-xLeft,h:0,line:{color,width,beginArrowType:'triangle',endArrowType:'none'}});
}
function label(slide, text, x,y,w,h, fs=18, color=C.ink, bold=false, align='center') {
  slide.addText(text,{x,y,w,h,fontSize:fs,color,bold,align,valign:'mid',margin:0.02,fit:'shrink'});
}
function callout(slide, text, color=C.teal, y=6.45, fs=16.2) {
  slide.addText(text,{x:0.8,y,w:11.73,h:0.58,fontSize:fs,bold:true,color:'FFFFFF',align:'center',valign:'mid',margin:0.05,
    fill:{color},line:{color},radius:0.12,fit:'shrink'});
}
function axis(slide, ox, oy, w, h) {
  arrow(slide,ox,oy,ox+w,oy,C.muted,1.3);
  arrow(slide,ox,oy,ox,oy-h,C.muted,1.3);
  label(slide,'descriptive content  f',ox+w-1.75,oy+0.16,1.9,0.28,12,C.muted,true);
  label(slide,'evaluative content  v',ox-0.45,oy-h-0.24,2.0,0.28,12,C.muted,true,'left');
}
function dot(slide,x,y,color=C.teal,r=0.14) { slide.addShape(pptx.ShapeType.ellipse,{x:x-r,y:y-r,w:2*r,h:2*r,fill:{color},line:{color}}); }
function node(slide,text,x,y,w,h,color=C.teal,fill=C.paper,fs=17){
  slide.addText(text,{x,y,w,h,fontSize:fs,bold:true,color,align:'center',valign:'mid',margin:0.05,fill:{color:fill},line:{color,width:1.6},radius:0.12,fit:'shrink'});
}

// 1
{
  const s=pptx.addSlide(); addBg(s,C.dark);
  s.addShape(pptx.ShapeType.rect,{x:0,y:0,w:0.18,h:H,fill:{color:C.teal2},line:{color:C.teal2}});
  s.addText('CAN FACTS REALLY BE\nSEPARATED FROM VALUES?',{x:0.8,y:0.8,w:7.0,h:1.6,fontFace:'Aptos Display',fontSize:32,bold:true,color:'FFFFFF',margin:0,breakLine:false,fit:'shrink'});
  s.addText('A visual introduction to Hilary Putnam’s fact–value entanglement',{x:0.82,y:2.6,w:6.7,h:0.5,fontSize:17,color:'C9D8D8',margin:0,fit:'shrink'});
  pill(s,'FACTS',8.15,1.0,3.65,0.76,C.teal,22);
  label(s,'“The temperature is 90°F”',8.05,1.9,4.0,0.38,14,'D5E2E2');
  s.addText('↕',{x:9.35,y:2.35,w:1.3,h:0.48,fontSize:28,bold:true,color:'FFFFFF',align:'center',margin:0});
  pill(s,'VALUES',8.15,2.95,3.65,0.76,C.green,22);
  label(s,'“Cruelty is bad”',8.05,3.85,4.0,0.38,14,'D5E2E2');
  s.addText('The ordinary picture looks clean.\nMaybe too clean.',{x:0.82,y:5.15,w:6.7,h:0.9,fontSize:19.5,bold:true,color:'FFFFFF',margin:0,fit:'shrink'});
  s.addText('01',{x:12.1,y:6.95,w:0.6,h:0.24,fontSize:9,bold:true,color:'8CA0A1',align:'right',margin:0});
}
// 2
{
  const s=pptx.addSlide(); addTop(s,2,'The traditional dichotomy'); addSub(s,'The picture we are about to test.');
  card(s,0.9,2.05,5.45,3.35,'F = FACTS','observation\nmeasurement\nempirical claims\ntheory',C.blue,C.paleBlue,17);
  card(s,6.95,2.05,5.45,3.35,'V = VALUES','good / bad\nright / wrong\nrational / irrational\nought',C.orange,C.paleOrange,17);
  label(s,'F ∩ V = ∅',4.75,5.48,3.85,0.72,32,C.ink,true);
  callout(s,'The alleged boundary: description on one side, evaluation on the other.',C.teal);
}
// 3
{
  const s=pptx.addSlide(); addTop(s,3,'The first crack: mutual influence'); addSub(s,'Even before Putnam, the clean wall starts leaking.');
  node(s,'V\nvalues / priorities',0.95,2.35,3.4,1.32,C.green,C.paleGreen,19);
  node(s,'F\nfacts / institutions / theories',8.95,2.35,3.4,1.32,C.blue,C.paleBlue,18);
  arrow(s,4.55,2.78,8.75,2.78,C.green,3); arrow(s,8.75,3.28,4.55,3.28,C.blue,3);
  label(s,'Values select what gets\nbuilt, funded, measured.',1.05,4.18,4.6,0.72,16,C.text,true);
  label(s,'Facts and consequences\nreshape values next.',7.7,4.18,4.6,0.72,16,C.text,true);
  ['law','research','design'].forEach((t,i)=>pill(s,t,1.0+i*1.52,5.22,1.3,0.44,[C.green,C.teal,C.orange][i],11.5));
  callout(s,'Interaction is real. But interaction alone is not yet Putnam.',C.teal);
}
// 4
{
  const s=pptx.addSlide(); addTop(s,4,'Interaction is not yet entanglement'); addSub(s,'A useful objection keeps us honest.');
  node(s,'Thermostat',1.15,2.35,3.25,0.95,C.blue,C.paleBlue,20); node(s,'Furnace',8.95,2.35,3.25,0.95,C.orange,C.paleOrange,20);
  arrow(s,4.55,2.65,8.75,2.65,C.blue,2.4); arrow(s,8.75,3.0,4.55,3.0,C.orange,2.4);
  label(s,'They interact.',1.15,3.9,3.25,0.45,18,C.ink,true); label(s,'They remain distinct.',8.7,3.9,3.75,0.45,18,C.ink,true);
  label(s,'V → F and F → V',2.1,5.05,2.7,0.45,18,C.teal,true);
  label(s,'≠',6.1,4.9,1.1,0.6,30,C.red,true);
  label(s,'conceptual inseparability',8.0,5.05,3.3,0.45,17,C.red,true);
  callout(s,'This slide blocks the sloppy version of the argument.',C.red);
}
// 5
{
  const s=pptx.addSlide(); addTop(s,5,'Enter the thick concept'); addSub(s,'Putnam’s pressure point is conceptual, not merely causal.');
  s.addShape(pptx.ShapeType.ellipse,{x:2.85,y:2.0,w:3.85,h:3.2,fill:{color:C.paleBlue,transparency:6},line:{color:C.blue,width:2}});
  s.addShape(pptx.ShapeType.ellipse,{x:6.25,y:2.0,w:3.85,h:3.2,fill:{color:C.paleOrange,transparency:6},line:{color:C.orange,width:2}});
  label(s,'description',3.45,2.25,2.0,0.34,14,C.blue,true); label(s,'evaluation',7.65,2.25,1.8,0.34,14,C.orange,true);
  label(s,'CRUEL',5.25,3.0,2.8,0.85,37,C.ink,true);
  label(s,'“The guard behaved cruelly toward the prisoner.”',2.0,5.4,9.3,0.58,18,C.text,true);
  callout(s,'A thick concept says what happened and how to assess it.',C.teal);
}
// 6
{
  const s=pptx.addSlide(); addTop(s,6,'The tempting decomposition'); addSub(s,'The clean split tries to survive.');
  node(s,'CRUEL',0.85,2.55,2.1,0.85,C.ink,C.paper,23);
  arrow(s,3.05,2.9,4.35,2.25,C.muted,2.1); arrow(s,3.05,2.95,4.35,4.55,C.muted,2.1);
  card(s,4.45,1.72,7.75,1.85,'D = NEUTRAL DESCRIPTION','hit the prisoner · humiliated him\nwithheld food · caused pain',C.blue,C.paleBlue,16.5);
  card(s,4.45,4.0,7.75,1.85,'E = EVALUATION','bad · wrong · condemnable\nnot to be done',C.orange,C.paleOrange,16.5);
  callout(s,'Cruel = D + E ?  That is the tempting rescue operation.',C.red);
}
// 7
{
  const s=pptx.addSlide(); addTop(s,7,'Putnam: not so fast'); addSub(s,'The value is not merely stapled on afterward.');
  card(s,0.72,2.0,3.0,3.38,'RAW BEHAVIORS','caused pain\ntaught discipline\nhumiliated for amusement\nfollowed policy',C.blue,C.paleBlue,15.8);
  node(s,'Which facts\nare relevant?',4.65,2.55,3.35,1.25,C.teal,C.paleTeal,20);
  node(s,'Cruelty\njudgment',9.25,2.55,3.0,1.25,C.orange,C.paleOrange,19);
  arrow(s,3.85,3.18,4.5,3.18,C.blue,2.4); arrow(s,8.15,3.18,9.08,3.18,C.teal,2.4);
  // return arrow routed through open space below nodes
  arrow(s,10.75,4.12,6.55,5.0,C.orange,2.2);
  label(s,'evaluative understanding guides\nselection and interpretation',4.7,4.55,3.55,0.65,14.5,C.text,true);
  callout(s,'To recognize the relevant facts as cruelty, you already need the evaluative concept.',C.teal,6.45,15.6);
}
// 8
{
  const s=pptx.addSlide(); addTop(s,8,'A visual model of judgment'); addSub(s,'This is our map, not Putnam’s math.');
  axis(s,1.55,5.8,9.55,3.65); dot(s,6.65,3.52,C.teal,0.18);
  node(s,'J = (f, v)',9.0,2.1,2.55,0.78,C.teal,C.paleTeal,21);
  label(s,'each judgment has a descriptive\nand evaluative coordinate',5.1,4.1,3.25,0.72,14,C.text,true);
  callout(s,'The model lets us avoid the dumb claim that facts and values are simply identical.',C.teal);
}
// 9
{
  const s=pptx.addSlide(); addTop(s,9,'Some judgments lean descriptive'); addSub(s,'The model allows degrees.');
  axis(s,1.4,5.82,10.0,3.72); dot(s,9.9,5.0,C.blue,0.18);
  node(s,'J₁ ≈ (0.95, 0.05)',4.4,2.12,3.3,0.72,C.blue,C.paleBlue,20);
  card(s,8.0,3.25,3.45,1.2,'EXAMPLE','“The thermometer reads 22°C.”',C.blue,C.paleBlue,15);
  callout(s,'Close to the descriptive axis does not mean metaphysically value-free.',C.blue);
}
// 10
{
  const s=pptx.addSlide(); addTop(s,10,'Some judgments lean evaluative'); addSub(s,'The other extreme exists too.');
  axis(s,1.4,5.82,10.0,3.72); dot(s,2.55,2.62,C.orange,0.18);
  node(s,'J₂ ≈ (0.10, 0.90)',8.6,2.12,3.3,0.72,C.orange,C.paleOrange,20);
  card(s,2.65,3.25,3.55,1.3,'EXAMPLES','“That was admirable.”\n“That was disgusting.”',C.orange,C.paleOrange,15);
  callout(s,'Some claims are overwhelmingly evaluative without being empty of description.',C.orange,6.45,15.7);
}
// 11
{
  const s=pptx.addSlide(); addTop(s,11,'The thick region'); addSub(s,'Much of human life sits in the middle.');
  axis(s,1.45,5.78,9.95,3.68);
  s.addShape(pptx.ShapeType.ellipse,{x:3.45,y:2.55,w:5.0,h:2.45,fill:{color:C.paleTeal,transparency:14},line:{color:C.teal,width:1.6}});
  [['cruel',4.45,3.08,C.red],['brave',5.55,3.52,C.green],['unjust',6.85,2.95,C.orange],['dishonest',4.35,4.15,C.blue],['reasonable',6.5,4.18,C.teal]].forEach(([t,x,y,c])=>{dot(s,x,y,c,0.1);label(s,t,x+0.12,y-0.16,1.55,0.3,13,c,true,'left');});
  callout(s,'Thick concepts occupy the zone where description and evaluation travel together.',C.teal);
}
// 12
{
  const s=pptx.addSlide(); addTop(s,12,'Static entanglement is only the beginning'); addSub(s,'Now we let the picture move through time.');
  const xs=[1.35,5.1,8.85], labs=['J₀','J₁','J₂'], cols=[C.blue,C.teal,C.orange], fills=[C.paleBlue,C.paleTeal,C.paleOrange];
  xs.forEach((x,i)=>{node(s,labs[i],x,2.35,2.35,1.2,cols[i],fills[i],24);label(s,['t₀','t₁','t₂'][i],x+0.78,3.8,0.8,0.32,14,C.muted,true);if(i<2)arrow(s,x+2.45,2.95,xs[i+1]-0.15,2.95,C.muted,2.2);});
  label(s,'Jₜ = (fₜ, vₜ)',4.3,4.85,4.7,0.72,31,C.ink,true);
  callout(s,'The judgment changes because both factual situation and evaluative orientation can change.',C.teal,6.45,15.3);
}
// 13
{
  const s=pptx.addSlide(); addTop(s,13,'Values guide inquiry'); addSub(s,'Inquiry starts with standards, interests, and priorities.');
  node(s,'V₀\nstarting values',0.95,2.25,3.45,1.35,C.green,C.paleGreen,19);
  node(s,'T₁\ntheory / practice',8.95,2.25,3.45,1.35,C.blue,C.paleBlue,19); arrow(s,4.58,2.92,8.75,2.92,C.teal,3);
  card(s,0.98,4.1,5.2,1.45,'VALUES ASK','What counts as important?\nWhat counts as elegant?\nWhat counts as explanation?',C.green,C.paleGreen,14.7);
  [['simplicity',7.25,4.3,C.green],['evidence',9.2,4.3,C.blue],['scope',7.25,4.95,C.teal],['fairness',9.2,4.95,C.orange]].forEach(([t,x,y,c])=>pill(s,t,x,y,1.65,0.48,c,11.5));
  callout(s,'Values do not dictate the answer. They help determine the question and the standard of success.',C.teal,6.45,15.0);
}
// 14
{
  const s=pptx.addSlide(); addTop(s,14,'Inquiry reshapes value'); addSub(s,'The return loop is where the machine gets interesting.');
  const xs=[0.78,3.18,5.58,7.98,10.38], labs=['V₀','T₁','V₁','T₂','V₂'], cols=[C.green,C.blue,C.green,C.blue,C.orange], fills=[C.paleGreen,C.paleBlue,C.paleGreen,C.paleBlue,C.paleOrange];
  xs.forEach((x,i)=>{node(s,labs[i],x,2.35,1.8,0.98,cols[i],fills[i],22);if(i<4)arrow(s,x+1.87,2.84,xs[i+1]-0.08,2.84,C.muted,2.1);});
  const subs=['standards','theory','revised standards','new theory','new values']; xs.forEach((x,i)=>label(s,subs[i],x-0.05,3.65,1.9,0.28,12,C.muted));
  label(s,'V₀ → T₁ → V₁ → T₂ → V₂ → …',3.05,4.75,7.2,0.68,28,C.ink,true);
  callout(s,'Standards shape inquiry; successful inquiry can reshape standards.',C.teal);
}
// 15
{
  const s=pptx.addSlide(); addTop(s,15,'Artifact loop'); addSub(s,'Values become things; things train future values.');
  const nodes=[['Value\nstandards',0.72,C.green,C.paleGreen],['Design / UI\nArtifact',3.55,C.blue,C.paleBlue],['Use /\nExperience',6.38,C.orange,C.paleOrange],['New\nstandards',9.21,C.teal,C.paleTeal]];
  nodes.forEach(([t,x,c,f],i)=>{node(s,t,x,2.35,2.35,1.18,c,f,18);if(i<3)arrow(s,x+2.43,2.94,x+2.72,2.94,C.muted,2.1);});
  // deliberate feedback arrow below nodes
  returnArrow(s,1.95,4.25,10.35,C.teal,2.2);
  label(s,'feedback: artifact changes expectation',4.7,4.45,4.0,0.34,13,C.teal,true);
  label(s,'Vₜ → Aₜ → Vₜ₊₁',4.45,5.15,4.45,0.58,25,C.ink,true);
  callout(s,'Aesthetic and practical values shape artifacts; living with artifacts reshapes taste.',C.teal,6.45,15.3);
}
// 16
{
  const s=pptx.addSlide(); addTop(s,16,'World, belief, value, action'); addSub(s,'The agent-level feedback loop.');
  const ns=[['Wₜ\nworld',0.55,C.blue,C.paleBlue],['Bₜ\nbelief',3.03,C.teal,C.paleTeal],['Vₜ\nvalue',5.51,C.green,C.paleGreen],['Aₜ\naction',7.99,C.orange,C.paleOrange],['Wₜ₊₁\nnew world',10.47,C.red,C.paleRed]];
  ns.forEach(([t,x,c,f],i)=>{node(s,t,x,2.4,2.0,1.12,c,f,18);if(i<4)arrow(s,x+2.08,2.96,x+2.36,2.96,C.muted,2.1);});
  // clean return arrow directly beneath chain
  returnArrow(s,1.55,4.32,11.45,C.teal,2.1);
  label(s,'new world → new evidence → new evaluation',3.75,4.53,5.85,0.38,13.5,C.teal,true);
  callout(s,'We do not merely observe state space. Evaluation helps choose the next state.',C.teal,6.45,15.0);
}
// 17
{
  const s=pptx.addSlide(); addTop(s,17,'A coupled dynamical system'); addSub(s,'Not identical. Not independent.');
  node(s,'Fₜ\nfactual state',1.0,2.15,3.0,1.1,C.blue,C.paleBlue,19); node(s,'Vₜ\nvalue state',1.0,4.2,3.0,1.1,C.green,C.paleGreen,19);
  node(s,'Fₜ₊₁',9.35,2.15,3.0,1.1,C.blue,C.paleBlue,21); node(s,'Vₜ₊₁',9.35,4.2,3.0,1.1,C.green,C.paleGreen,21);
  node(s,'coupling',5.45,3.08,2.45,0.72,C.teal,C.paleTeal,16);
  // two inputs enter the coupling node; two outputs leave it — no crossing lines
  arrow(s,4.15,2.7,5.28,3.24,C.blue,1.8);
  arrow(s,4.15,4.75,5.28,3.66,C.green,1.8);
  arrow(s,7.9,3.24,9.18,2.7,C.blue,1.8);
  arrow(s,7.9,3.66,9.18,4.75,C.green,1.8);
  callout(s,'(Fₜ, Vₜ) ↦ (Fₜ₊₁, Vₜ₊₁)',C.ink);
}
// 18
{
  const s=pptx.addSlide(); addTop(s,18,'Rationality is already normative'); addSub(s,'The strongest “pure fact” refuge is not pure.');
  node(s,'RATIONALITY',4.65,1.95,4.0,0.88,C.teal,C.paleTeal,24);
  const items=[['follow evidence',0.8,3.42,C.blue,C.paleBlue],['avoid contradiction',3.75,3.42,C.red,C.paleRed],['revise belief',6.7,3.42,C.green,C.paleGreen],['prefer better\nexplanations',9.65,3.42,C.orange,C.paleOrange]];
  items.forEach(([t,x,y,c,f])=>node(s,t,x,y,2.55,1.0,c,f,14.8));
  items.forEach(([t,x,y])=>arrow(s,6.65,2.84,x+1.27,y-0.12,C.muted,1.35));
  label(s,'These are standards for how one ought to reason — not merely descriptions of what brains happen to do.',1.15,5.05,11.0,0.62,16,C.text,true);
  callout(s,'Strip out every norm and rationality collapses into causal psychology.',C.red);
}
// 19
{
  const s=pptx.addSlide(); addTop(s,19,'Conclusion'); addSub(s,'Distinguishable does not mean cleanly separable.');
  node(s,'F ≠ V',0.95,2.0,3.15,1.1,C.blue,C.paleBlue,28);
  node(s,'F and V\nare not independent',5.05,2.0,3.25,1.1,C.orange,C.paleOrange,18.5);
  node(s,'(Fₜ,Vₜ) →\n(Fₜ₊₁,Vₜ₊₁)',9.25,2.0,3.15,1.1,C.teal,C.paleTeal,20);
  [['1. Not identical',C.blue],['2. Not sealed off',C.orange],['3. Often dynamically entangled',C.teal]].forEach(([t,c],i)=>{
    s.addText(t,{x:2.0,y:4.0+i*0.58,w:9.3,h:0.42,fontSize:20,bold:true,color:c,align:'center',margin:0,fit:'shrink'});
  });
  callout(s,'Putnam’s point is not that facts are values. It is that our best descriptions and evaluations often cannot be cleanly pulled apart.',C.ink,6.45,14.2);
}

fs.mkdirSync('Presentation',{recursive:true});
(async()=>{ await pptx.writeFile({ fileName: 'Presentation/Putnam_Fact_Value_Entanglement_YouTube_Deck.pptx' }); console.log('wrote deck'); })().catch(e=>{console.error(e);process.exit(1)});
