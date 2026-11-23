const docx = require('docx');
const fs = require('fs');
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, HeadingLevel, AlignmentType, ShadingType, LevelFormat } = docx;
const FONT = "Calibri"; const BLUE = "1F3864";
function h(text){ return new Paragraph({heading:HeadingLevel.HEADING_1, spacing:{before:200,after:90}, children:[new TextRun({text, bold:true, color:BLUE, font:FONT})]}); }
function p(text, opts={}){ return new Paragraph({spacing:{after:80}, children:[new TextRun({text, font:FONT, size:21, ...opts})]}); }
function pr(runs){ return new Paragraph({spacing:{after:80}, children: runs.map(r=>new TextRun({font:FONT, size:21, ...r}))}); }
function bullet(runs){ return new Paragraph({numbering:{reference:"bul", level:0}, spacing:{after:60}, children: runs.map(r=>new TextRun({font:FONT, size:21, ...r}))}); }
function cell(text, {bold=false, shade=null, w=1800}={}){
  return new TableCell({width:{size:w,type:WidthType.DXA}, shading: shade?{type:ShadingType.CLEAR, fill:shade}:undefined,
    margins:{top:60,bottom:60,left:100,right:100},
    children:[new Paragraph({children:[new TextRun({text, bold, font:FONT, size:20})]})]});
}
const CW=[1900,1900,2400,3160]; const TW=CW.reduce((a,b)=>a+b,0);
const hdr = ["Location","Core Target","Flex Band (your call)","Over This = Ask Ryan First"];
const rows = [
 ["Studio","30%","30–34%","34%"],
 ["Downtown","32%","32–36%","36%"],
 ["1899 Penn","35%","35–40%","40%"],
 ["NMWA","1 person/day","—","Any 2nd person / event staffing"],
];
const table = new Table({columnWidths:CW, width:{size:TW,type:WidthType.DXA}, rows:[
  new TableRow({tableHeader:true, children: hdr.map((t,i)=>cell(t,{bold:true, shade:"DCE6F1", w:CW[i]}))}),
  ...rows.map(r=> new TableRow({children: r.map((t,i)=>cell(t,{w:CW[i]}))}))
]});
const doc = new Document({
 numbering:{config:[{reference:"bul", levels:[{level:0, format:LevelFormat.BULLET, text:"•", alignment:AlignmentType.LEFT, style:{paragraph:{indent:{left:360,hanging:200}}}}]}]},
 sections:[{properties:{page:{size:{width:12240,height:15840}, margin:{top:1080,bottom:1080,left:1260,right:1260}}}, children:[
  new Paragraph({spacing:{after:40}, children:[new TextRun({text:"SLIPSTREAM COFFEE", bold:true, color:BLUE, size:30, font:FONT})]}),
  new Paragraph({spacing:{after:160}, children:[new TextRun({text:"Labor Targets & Scheduling Process — effective with September schedules", bold:true, size:24, font:FONT})]}),
  pr([{text:"To: Mo + ALLs      From: Ryan      Re: ", italics:true},{text:"New labor targets, admin/training split, and the schedule approval rule", italics:true}]),
  h("What's changing, and why"),
  p("July's prime cost hit 68.5% — past our 65% problem line — mostly because labor didn't flex down with slower sales. At the same time, admin shifts (ALL coverage/admin blocks) have become a regular part of every week. So we're splitting labor into three buckets and setting targets only on the part you control when you build a schedule."),
  h("Definitions"),
  bullet([{text:"Core labor: ", bold:true},{text:"every scheduled shift EXCEPT Manager and Training shifts. This is what the targets measure."}]),
  bullet([{text:"Admin: ", bold:true},{text:"shifts scheduled under the “Manager” job in Square. Tracked separately — no target or cap for now, just keep them visible and reasonable."}]),
  bullet([{text:"Training: ", bold:true},{text:"shifts under the “Training” role. Plan ~1% of sales; over 1.5% in a week, check with Ryan. Cert classes will create spikes — that's fine, just expected ones."}]),
  bullet([{text:"Sales basis: ", bold:true},{text:"walk-in sales (register business). Invoiced/house-account items don't count in the denominator."}]),
  h("The numbers"),
  table,
  new Paragraph({spacing:{before:80,after:80}, children:[new TextRun({text:"Between target and the ceiling is your flex band — slower weeks, newer team members, coverage judgment calls. No approval needed, just aim back at target. Over the ceiling: message Ryan BEFORE publishing, with the reason. August and December will test the ceilings — that's expected; the ask is the point.", italics:true, font:FONT, size:20})]}),
  h("Rule of thumb — building the week without a calculator"),
  bullet([{text:"One full 8-hour shift per $500–600 of projected sales. ", bold:true},{text:"$1,000 day? 2 shifts. $1,500? 3 shifts. $1,700–1,800? 3.5. Get close with this, then run the 30-second check before publishing."}]),
  bullet([{text:"Aim for the high end at Studio and Downtown (~$600/shift) and the low end at 1899 (~$500/shift) — that difference is exactly why their targets differ."}]),
  bullet([{text:"It's a WEEKLY number, not a daily one. ", bold:true},{text:"If Mon–Thu projects ~$1,250/day and you staffed 3 shifts each day instead of 2.5, you owe about 2 shifts back across Fri–Sun. Balance the week, not each day."}]),
  h("The 30-second check before you publish"),
  bullet([{text:"1.  ", bold:true},{text:"In the Square draft, take the week's scheduled $ for the location."}]),
  bullet([{text:"2.  ", bold:true},{text:"Subtract Manager and Training shifts (tracked separately, outside the target)."}]),
  bullet([{text:"3.  ", bold:true},{text:"Divide by that week's projected sales — the weekly report gives you the next 4 weeks of projected sales and target labor $ per location, so this is just a comparison, not math."}]),
  bullet([{text:"4.  ", bold:true},{text:"Under target: publish. In the band: publish. Over the ceiling: send to Ryan first."}]),
  h("Standing rules"),
  bullet([{text:"Federal-holiday Mondays: ", bold:true},{text:"Downtown and 1899 are CLOSED — no staffing."}]),
  bullet([{text:"1899: ", bold:true},{text:"the 40% ceiling is where the cafe stops making money — staying under it is what funds our incentive there."}]),
  bullet([{text:"NMWA: ", bold:true},{text:"1 person per day plus the shared admin block. Focus is building sales, not cutting cost."}]),
  h("Working notes (from Mo's questions)"),
  bullet([{text:"Weekly targets: ", bold:true},{text:"you won't need to calculate them — the weekly report will include the next 4 weeks of projected sales and target labor $ for each location."}]),
  bullet([{text:"Typical training shift cost: ", bold:true},{text:"guidance is already in the training sheet."}]),
  bullet([{text:"Typical admin shift cost: ", bold:true},{text:"Mo to propose what “reasonable” looks like; Ryan approves. No cap until then — just keep it visible."}]),
  bullet([{text:"Flexibility mindset: ", bold:true},{text:"admin shifts are a coverage tool. If a day only needs 3.5 shifts, staff 3 and the ALL covers rushes and breaks — accepting that some project/admin work slides that day. In slow stretches, that's exactly the trade we want."}]),
  h("Worked example — the 10/11 template week"),
  bullet([{text:"Studio: ", bold:true},{text:"ships as-is (30.8% core). This is the reference template — full admin coverage AND on target."}]),
  bullet([{text:"NMWA: ", bold:true},{text:"ships as-is."}]),
  bullet([{text:"Downtown: ", bold:true},{text:"holiday week: Monday is closed (unschedule it). Standard weeks: trim ~$300 — Tuesday runs 4-deep (cut the overlapping Production Barista) and end the 3:30+ afternoon tails."}]),
  bullet([{text:"1899: ", bold:true},{text:"holiday Monday is closed (unschedule it); second person goes half-day on slow days → lands ~$1,500/wk scheduled labor."}]),
  h("Why this matters"),
  p("This is the path back under 65% prime cost — and when we consistently beat these targets, it creates the room to share busy-week upside with the team, which is where I want this to go."),
 ]}]
});
Packer.toBuffer(doc).then(b=>{fs.writeFileSync("/sessions/vigilant-beautiful-cray/mnt/Weekly and Monthly Reports/Slipstream Labor Targets Memo - Sept 2026.docx", b); console.log("written");});
