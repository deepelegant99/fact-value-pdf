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
pptx.margin = 0;

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
function card(slide, x,y,w,h, title, body, color=C.teal, fill=C.paper) {
  slide.addShape(pptx.ShapeType.roundRect,{x,y,w,h,rectRadius:0.08,fill:{color:fill},line:{color:C.line,width:1.1}});
  slide.addShape(pptx.ShapeType.rect,{x,y,w:0.08,h,fill:{color},line:{color}});
  if(title) slide.addText(title,{x:x+0.28,y:y+0.22,w:w-0.5,h:0.36,fontSize:15,bold:true,color,margin:0,fit:'shrink'});
  slide.addText(body,{x:x+0.28,y:y+0.72,w:w-0.55,h:h-0.9,fontSize:16.5,color:C.text,margin:0.02,breakLine:false,fit:'shrink',valign:'mid'});
}
function arrow(slide, x1,y1,x2,y2,color=C.teal,width=2.2,end='triangle') {
  slide.addShape(pptx.ShapeType.line,{x:x1,y:y1,w:x2-x1,h:y2-y1,line:{color,width,beginArrowType:'none',endArrowType:end}});
}
function label(slide, text, x,y,w,h, fs=18, color=C.ink, bold=false, align='center') {
  slide.addText(text,{x,y,w,h,fontSize:fs,color,bold,align,valign:'mid',margin:0.02,fit:'shrink'});
}
function callout(slide, text, color=C.teal, y=6.45) {
  slide.addText(text,{x:0.8,y,w:11.73,h:0.58,fontSize:16.2,bold:true,color:'FFFFFF',align:'center',valign:'mid',margin:0.05,
    fill:{color},line:{color},radius:0.12,fit:'shrink'});
}
function footer(slide, text='') {
  if(text) slide.addText(text,{x:0.58,y:7.17,w:11.4,h:0.13,fontSize:6.8,color:C.muted,margin:0,fit:'shrink'});
}
function axis(slide, ox, oy, w, h) {
  arrow(slide,ox,oy,ox+w,oy,C.muted,1.3);
  arrow(slide,ox,oy,ox,oy-h,C.muted,1.3);
  label(slide,'f',ox+w-0.15,oy+0.12,0.35,0.25,15,C.muted,true);
  label(slide,'v',ox-0.35,oy-h-0.18,0.3,0.25,15,C.muted,true);
}
function dot(slide,x,y,color=C.teal,r=0.14) { slide.addShape(pptx.ShapeType.ellipse,{x:x-r,y:y-r,w:2*r,h:2*r,fill:{color},line:{color}}); }
function node(slide,text,x,y,w,h,color=C.teal,fill=C.paper,fs=17){
  slide.addText(text,{x,y,w,h,fontSize:fs,bold:true,color,align:'center',valign:'mid',margin:0.05,fill:{color:fill},line:{color,width:1.6},radius:0.12,fit:'shrink'});
}

{
  const s=pptx.addSlide(); addBg(s,C.dark);
  s.addShape(pptx.ShapeType.rect,{x:0,y:0,w:0.18,h:H,fill:{color:C.teal2},line:{color:C.teal2}});
  s.addText('CAN FACTS REALLY BE\nSEPARATED FROM VALUES?',{x:0.8,y:0.8,w:7.0,h:1.6,fontFace:'Aptos Display',fontSize:32,bold:true,color:'FFFFFF',margin:0,breakLine:false,fit:'shrink'});
  s.addText('A visual introduction to Hilary Putnam’s fact–value entanglement',{x:0.82,y:2.6,w:6.7,h:0.5,fontSize:17,color:'C9D8D8',margin:0,fit:'shrink'});
  pill(s,'FACTS',8.1,1.15,3.8,0.86,C.teal,23);
  label(s,'“The temperature is 90°F”',8.05,2.15,4.0,0.45,14,'D5E2E2');
  s.addText('↕',{x:9.4,y:2.72,w:1.3,h:0.55,fontSize:30,bold:true,color:'FFFFFF',align:'center',margin:0});
  pill(s,'VALUES',8.1,3.35,3.8,0.86,C.green,23);
  label(s,'“Cruelty is bad”',8.05,4.35,4.0,0.45,14,'D5E2E2');
  s.addText('The ordinary picture looks clean. Too clean.',{x:0.82,y:5.4,w:6.9,h:0.7,fontSize:20,bold:true,color:'FFFFFF',margin:0,fit:'shrink'});
  s.addText('01',{x:12.1,y:6.95,w:0.6,h:0.24,fontSize:9,bold:true,color:'8CA0A1',align:'right',margin:0});
}
{
  const s=pptx.addSlide(); addTop(s,2,'The traditional dichotomy'); addSub(s,'The picture we are about to test.');
  card(s,0.9,2.15,5.4,3.25,'F = FACTS','observation\nmeasurement\nempirical claims\ntheory',C.blue,C.paleBlue);
  card(s,7.0,2.15,5.4,3.25,'V = VALUES','good / bad\nright / wrong\nrational / irrational\nought',C.orange,C.paleOrange);
  label(s,'F ∩ V = ∅',5.1,5.68,3.15,0.55,28,C.ink,true);
  callout(s,'The alleged boundary: description on one side, evaluation on the other.',C.teal);
}
{
  const s=pptx.addSlide(); addTop(s,3,'The first crack: mutual influence'); addSub(s,'Even before Putnam, the clean wall starts leaking.');
  node(s,'V\nvalues / priorities',1.15,2.5,3.15,1.25,C.green,C.paleGreen,19);
  node(s,'F\nfacts / institutions / theories',9.0,2.5,3.15,1.25,C.blue,C.paleBlue,18);
  arrow(s,4.5,2.88,8.85,2.88,C.green,3); arrow(s,8.85,3.42,4.5,3.42,C.blue,3);
  label(s,'Values select what gets built, funded, measured, defended.',1.0,4.45,5.1,0.62,16,C.text,false,'left');
  label(s,'Facts and consequences reshape what people value next.',7.15,4.45,5.1,0.62,16,C.text,false,'left');
  ['law','research agenda','design choice'].forEach((t,i)=>pill(s,t,1.2+i*1.55,5.35,1.35,0.44,[C.green,C.teal,C.orange][i],11.5));
  callout(s,'V → F and F → V: interaction is real. But interaction alone is not yet Putnam.',C.teal);
}
{
  const s=pptx.addSlide(); addTop(s,4,'Interaction is not yet entanglement'); addSub(s,'A useful objection keeps us honest.');
  node(s,'Thermostat',1.25,2.45,3.2,1.0,C.blue,C.paleBlue,20); node(s,'Furnace',8.9,2.45,3.2,1.0,C.orange,C.paleOrange,20);
  arrow(s,4.65,2.78,8.7,2.78,C.blue,2.5); arrow(s,8.7,3.28,4.65,3.28,C.orange,2.5);
  label(s,'They interact.',1.25,4.1,3.2,0.45,19,C.ink,true); label(s,'They are not the same thing.',8.45,4.1,4.1,0.45,19,C.ink,true);
  callout(s,'V → F and F → V  ≠  conceptual inseparability',C.red);
}
{
  const s=pptx.addSlide(); addTop(s,5,'Enter the thick concept'); addSub(s,'Putnam’s pressure point is not causal. It is conceptual.');
  s.addShape(pptx.ShapeType.ellipse,{x:3.2,y:2.08,w:3.45,h:3.0,fill:{color:C.paleBlue,transparency:5},line:{color:C.blue,width:2}});
  s.addShape(pptx.ShapeType.ellipse,{x:6.45,y:2.08,w:3.45,h:3.0,fill:{color:C.paleOrange,transparency:5},line:{color:C.orange,width:2}});
  label(s,'description',3.5,2.35,2.3,0.35,15,C.blue,true); label(s,'evaluation',7.3,2.35,2.0,0.35,15,C.orange,true);
  label(s,'CRUEL',5.35,3.05,2.65,0.78,34,C.ink,true);
  label(s,'“The guard behaved cruelly toward the prisoner.”',2.2,5.35,8.9,0.62,18,C.text,true);
  callout(s,'A thick concept seems to say what happened and how to assess it.',C.teal);
}
{
  const s=pptx.addSlide(); addTop(s,6,'The tempting decomposition'); addSub(s,'The clean split tries to survive.');
  node(s,'CRUEL',0.95,2.55,2.2,0.9,C.ink,C.paper,24);
  arrow(s,3.25,3.0,4.45,2.35,C.muted,2.2); arrow(s,3.25,3.0,4.45,4.65,C.muted,2.2);
  card(s,4.6,1.8,7.55,1.65,'D = NEUTRAL DESCRIPTION','hit the prisoner · humiliated him · withheld food · caused pain',C.blue,C.paleBlue);
  card(s,4.6,4.15,7.55,1.65,'E = EVALUATION','bad · wrong · condemnable · not to be done',C.orange,C.paleOrange);
  callout(s,'Cruel = D + E ?',C.red);
}
{
  const s=pptx.addSlide(); addTop(s,7,'Putnam: not so fast'); addSub(s,'The value is not merely stapled on afterward.');
  card(s,0.8,2.05,3.15,3.2,'RAW BEHAVIORS','“caused pain”\n“taught discipline”\n“humiliated for amusement”\n“followed policy”',C.blue,C.paleBlue);
  node(s,'Which details\nmatter?',5.0,2.72,3.2,1.2,C.teal,C.paleTeal,19);
  node(s,'Cruelty\njudgment',9.35,2.72,3.0,1.2,C.orange,C.paleOrange,19);
  arrow(s,4.0,3.3,4.85,3.3,C.blue,2.4); arrow(s,8.25,3.3,9.2,3.3,C.teal,2.4); arrow(s,10.85,4.1,7.55,5.05,C.orange,2.4);
  label(s,'evaluative understanding guides\nselection and interpretation',5.0,4.55,3.3,0.72,15,C.text,true);
  callout(s,'To recognize the relevant facts as cruelty, you already need the evaluative concept.',C.teal);
}
{
  const s=pptx.addSlide(); addTop(s,8,'A visual model of judgment'); addSub(s,'This is our map, not Putnam’s math.');
  axis(s,2.05,5.75,8.9,3.45); dot(s,6.7,3.55,C.teal,0.18);
  label(s,'f = descriptive content',9.4,5.8,2.5,0.35,14,C.muted); label(s,'v = evaluative content',0.75,2.05,2.5,0.35,14,C.muted);
  node(s,'J = (f, v)',8.45,2.55,2.9,0.85,C.teal,C.paleTeal,22);
  callout(s,'A judgment can have both descriptive and evaluative coordinates.',C.teal);
}
{
  const s=pptx.addSlide(); addTop(s,9,'Some judgments lean descriptive'); addSub(s,'The model allows degrees.');
  axis(s,2.05,5.75,8.9,3.45); dot(s,9.8,5.18,C.blue,0.18);
  label(s,'“thermometer reads 22°C”',7.1,4.3,3.8,0.5,17,C.text,true); label(s,'“water boils at sea level”',7.1,3.72,3.8,0.45,15,C.muted);
  node(s,'J₁ ≈ (0.95, 0.05)',4.35,2.55,3.2,0.82,C.blue,C.paleBlue,20);
  callout(s,'Close to the descriptive axis does not mean literally value-free.',C.blue);
}
{
  const s=pptx.addSlide(); addTop(s,10,'Some judgments lean evaluative'); addSub(s,'The other extreme also exists.');
  axis(s,2.05,5.75,8.9,3.45); dot(s,2.95,2.75,C.orange,0.18);
  label(s,'“admirable”',3.35,2.42,2.5,0.45,18,C.text,true); label(s,'“disgusting”',3.35,3.0,2.5,0.45,16,C.muted);
  node(s,'J₂ ≈ (0.10, 0.90)',7.0,2.55,3.2,0.82,C.orange,C.paleOrange,20);
  callout(s,'Some claims are overwhelmingly evaluative without being empty of description.',C.orange);
}
{
  const s=pptx.addSlide(); addTop(s,11,'The thick region'); addSub(s,'Much of human life sits in the middle.');
  axis(s,2.05,5.75,8.9,3.45);
  s.addShape(pptx.ShapeType.ellipse,{x:4.5,y:2.55,w:4.25,h:2.3,fill:{color:C.paleTeal,transparency:18},line:{color:C.teal,width:1.5,dash:'dash'}});
  [['cruel',5.25,3.08,C.red],['brave',6.45,3.7,C.green],['unjust',7.15,2.9,C.orange],['dishonest',5.15,4.1,C.blue],['reasonable',7.0,4.2,C.teal]].forEach(([t,x,y,c])=>{dot(s,x,y,c,0.11);label(s,t,x+0.15,y-0.17,1.5,0.32,13,c,true,'left');});
  callout(s,'Thick concepts occupy the zone where description and evaluation travel together.',C.teal);
}
{
  const s=pptx.addSlide(); addTop(s,12,'Static entanglement is only the beginning'); addSub(s,'Now we let the picture move through time.');
  const xs=[1.3,5.15,9.0]; const labs=['J₀','J₁','J₂']; const ts=['t₀','t₁','t₂'];
  xs.forEach((x,i)=>{node(s,labs[i],x,2.4,2.3,1.2,[C.blue,C.teal,C.orange][i],[C.paleBlue,C.paleTeal,C.paleOrange][i],24);label(s,ts[i],x+0.75,3.9,0.8,0.35,15,C.muted,true);if(i<2) arrow(s,x+2.45,3.0,xs[i+1]-0.15,3.0,C.muted,2.2);});
  label(s,'Jₜ = (fₜ, vₜ)',4.65,5.1,4.0,0.6,26,C.ink,true);
  callout(s,'The judgment changes because both our factual situation and our evaluative orientation can change.',C.teal);
}
{
  const s=pptx.addSlide(); addTop(s,13,'Values guide inquiry'); addSub(s,'Inquiry starts with standards, interests, and priorities.');
  node(s,'V₀\nstarting values',1.2,2.45,3.15,1.3,C.green,C.paleGreen,19);
  node(s,'T₁\ntheory / practice',8.95,2.45,3.15,1.3,C.blue,C.paleBlue,19); arrow(s,4.5,3.1,8.75,3.1,C.teal,3);
  ['What counts as important?','What counts as elegant?','What counts as explanation?'].forEach((t,i)=>label(s,t,2.0,4.45+i*0.47,5.2,0.33,15,C.text,false,'left'));
  ['simplicity','evidence','scope','fairness'].forEach((t,i)=>pill(s,t,7.9+(i%2)*1.75,4.45+Math.floor(i/2)*0.62,1.55,0.44,[C.green,C.blue,C.teal,C.orange][i],11.5));
  callout(s,'Values do not dictate the answer. They help determine what counts as a question and a good answer.',C.teal);
}
{
  const s=pptx.addSlide(); addTop(s,14,'Inquiry reshapes value'); addSub(s,'The return loop is where the machine gets interesting.');
  const xs=[0.85,3.25,5.65,8.05,10.45]; const labs=['V₀','T₁','V₁','T₂','V₂']; const cols=[C.green,C.blue,C.green,C.blue,C.orange]; const fills=[C.paleGreen,C.paleBlue,C.paleGreen,C.paleBlue,C.paleOrange];
  xs.forEach((x,i)=>{node(s,labs[i],x,2.55,1.75,1.0,cols[i],fills[i],22);if(i<4)arrow(s,x+1.82,3.05,xs[i+1]-0.08,3.05,C.muted,2.2);});
  label(s,'standards',0.9,4.15,1.65,0.3,13,C.muted);label(s,'theory',3.3,4.15,1.65,0.3,13,C.muted);label(s,'revised standards',5.42,4.15,2.2,0.3,13,C.muted);label(s,'new theory',8.0,4.15,1.9,0.3,13,C.muted);label(s,'new values',10.4,4.15,1.9,0.3,13,C.muted);
  label(s,'V₀ → T₁ → V₁ → T₂ → V₂ → …',3.05,5.0,7.2,0.6,25,C.ink,true);
  callout(s,'Inquiry is recursive: standards shape inquiry; successful inquiry can reshape standards.',C.teal);
}
{
  const s=pptx.addSlide(); addTop(s,15,'Artifact loop'); addSub(s,'Values become things; things train future values.');
  const nodes=[['Value\nstandards',0.65,C.green,C.paleGreen],['Design / UI\nArtifact',3.55,C.blue,C.paleBlue],['Use /\nExperience',6.45,C.orange,C.paleOrange],['New\nstandards',9.35,C.teal,C.paleTeal]];
  nodes.forEach(([t,x,c,f],i)=>{node(s,t,x,2.45,2.45,1.25,c,f,18);if(i<3)arrow(s,x+2.55,3.08,x+2.8,3.08,C.muted,2.2);});
  arrow(s,10.55,4.0,1.65,5.15,C.teal,2.2); label(s,'feedback',5.95,4.55,1.4,0.35,14,C.teal,true);
  label(s,'Vₜ → Aₜ → Vₜ₊₁',4.5,5.55,4.3,0.55,25,C.ink,true);
  callout(s,'Aesthetic and practical values shape artifacts; living with artifacts reshapes taste and expectation.',C.teal);
}
{
  const s=pptx.addSlide(); addTop(s,16,'World, belief, value, action'); addSub(s,'The agent-level feedback loop.');
  const ns=[['Wₜ\nworld',0.65,C.blue,C.paleBlue],['Bₜ\nbelief',3.15,C.teal,C.paleTeal],['Vₜ\nvalue',5.65,C.green,C.paleGreen],['Aₜ\naction',8.15,C.orange,C.paleOrange],['Wₜ₊₁\nnew world',10.65,C.red,C.paleRed]];
  ns.forEach(([t,x,c,f],i)=>{node(s,t,x,2.55,1.95,1.15,c,f,18);if(i<4)arrow(s,x+2.02,3.13,x+2.35,3.13,C.muted,2.2);});
  arrow(s,11.7,4.1,1.55,5.3,C.teal,2.0); label(s,'new state → new evidence → new evaluation',3.7,4.85,6.0,0.45,15,C.teal,true);
  callout(s,'We do not merely observe state space. Evaluation helps determine which state we try to move into.',C.teal);
}
{
  const s=pptx.addSlide(); addTop(s,17,'A coupled dynamical system'); addSub(s,'Not identical. Not independent.');
  node(s,'Fₜ\nfactual state',1.05,2.25,3.1,1.2,C.blue,C.paleBlue,20); node(s,'Vₜ\nvalue state',1.05,4.25,3.1,1.2,C.green,C.paleGreen,20);
  node(s,'Fₜ₊₁',9.25,2.25,3.1,1.2,C.blue,C.paleBlue,22); node(s,'Vₜ₊₁',9.25,4.25,3.1,1.2,C.green,C.paleGreen,22);
  arrow(s,4.3,2.85,9.05,2.85,C.blue,2.5); arrow(s,4.3,4.85,9.05,4.85,C.green,2.5); arrow(s,4.1,2.95,9.1,4.55,C.teal,2.0); arrow(s,4.1,4.55,9.1,3.15,C.orange,2.0);
  label(s,'coupling',5.95,3.52,1.5,0.4,14,C.teal,true);
  callout(s,'(Fₜ, Vₜ) ↦ (Fₜ₊₁, Vₜ₊₁)',C.ink);
}
{
  const s=pptx.addSlide(); addTop(s,18,'Rationality is already normative'); addSub(s,'The strongest “pure fact” refuge is not pure.');
  node(s,'RATIONALITY',4.55,2.0,4.2,0.9,C.teal,C.paleTeal,25);
  const items=[['follow evidence',1.0,3.55,C.blue,C.paleBlue],['avoid contradiction',3.95,3.55,C.red,C.paleRed],['revise belief',6.9,3.55,C.green,C.paleGreen],['prefer better explanations',9.85,3.55,C.orange,C.paleOrange]];
  items.forEach(([t,x,y,c,f])=>node(s,t,x,y,2.5,1.0,c,f,15.5));
  items.forEach(([t,x,y])=>arrow(s,6.65,2.95,x+1.25,y-0.13,C.muted,1.5));
  label(s,'These are standards for how one ought to reason—not merely descriptions of what brains happen to do.',1.15,5.2,11.0,0.7,18,C.text,true);
  callout(s,'Remove every normative standard and “rationality” collapses into causal psychology.',C.red);
}
{
  const s=pptx.addSlide(); addTop(s,19,'Conclusion'); addSub(s,'Distinguishable does not mean cleanly separable.');
  node(s,'F ≠ V',1.0,2.0,3.1,1.15,C.blue,C.paleBlue,28);
  node(s,'F ⟂ V ?',5.1,2.0,3.1,1.15,C.orange,C.paleOrange,28);
  node(s,'(Fₜ,Vₜ) →\n(Fₜ₊₁,Vₜ₊₁)',9.2,2.0,3.1,1.15,C.teal,C.paleTeal,21);
  const lines=[['1. Not identical',C.blue],['2. Not always separable',C.orange],['3. Often dynamically entangled',C.teal]];
  lines.forEach(([t,c],i)=>{s.addText(t,{x:2.0,y:4.0+i*0.58,w:9.3,h:0.42,fontSize:20,bold:true,color:c,align:'center',margin:0,fit:'shrink'});});
  callout(s,'Putnam’s point is not that facts are values. It is that our best descriptions and our best evaluations often cannot be cleanly pulled apart.',C.ink);
}

fs.mkdirSync('Presentation',{recursive:true});
pptx.writeFile({ fileName: 'Presentation/Putnam_Fact_Value_Entanglement_YouTube_Deck.pptx' });
