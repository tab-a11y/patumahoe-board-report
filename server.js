const express = require('express');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Footer, AlignmentType, BorderStyle, WidthType, ShadingType,
  VerticalAlign, PageNumber, LevelFormat
} = require('docx');

const app  = express();
const PORT = process.env.PORT || 3000;
app.use(express.json({ limit: '4mb' }));

app.get('/', (req, res) => res.send('Patumahoe Board Report Service running OK'));

app.post('/generate', async (req, res) => {
  try {
    const { reportText, month, year, term } = req.body;
    if (!reportText) return res.status(400).json({ error: 'reportText is required' });
    const buffer = await buildDocx(reportText, month || '', year || '', term || '');
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', `attachment; filename="Board_Report_${month}_${year}.docx"`);
    res.send(buffer);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

async function buildDocx(reportText, month, year, term) {
  const NAVY = "1F3864", BLUE = "2E75B6", GREY = "F2F2F2", WHITE = "FFFFFF", BLACK = "000000", DGREY = "595959";
  const CONTENT = 9026;
  const bdr = (c="AAAAAA") => ({ style: BorderStyle.SINGLE, size: 1, color: c });
  const borders = { top: bdr(), bottom: bdr(), left: bdr(), right: bdr() };
  const noBorder = { style: BorderStyle.NONE, size: 0, color: WHITE };
  const noBorders = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder };

  const hdrCell = (text, w) => new TableCell({
    borders, width: { size: w, type: WidthType.DXA },
    shading: { fill: NAVY, type: ShadingType.CLEAR },
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    verticalAlign: VerticalAlign.CENTER,
    children: [new Paragraph({ alignment: AlignmentType.CENTER,
      children: [new TextRun({ text, bold: true, color: WHITE, size: 18, font: "Arial" })] })]
  });

  const dataCell = (text, w, shade=WHITE, center=false) => new TableCell({
    borders, width: { size: w, type: WidthType.DXA },
    shading: { fill: shade, type: ShadingType.CLEAR },
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    verticalAlign: VerticalAlign.CENTER,
    children: [new Paragraph({
      alignment: center ? AlignmentType.CENTER : AlignmentType.LEFT,
      children: [new TextRun({ text, color: BLACK, size: 18, font: "Arial" })]
    })]
  });

  const spacer = () => new Paragraph({ spacing: { before: 60, after: 60 }, children: [new TextRun("")] });

  const sectionHead = (text) => new Paragraph({
    spacing: { before: 280, after: 100 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: BLUE, space: 4 } },
    children: [new TextRun({ text: text.replace(/^#+\s*/, '').toUpperCase(), bold: true, size: 22, color: NAVY, font: "Arial" })]
  });

  const subHead = (text) => new Paragraph({
    spacing: { before: 160, after: 60 },
    children: [new TextRun({ text: text.replace(/^#+\s*/, '').replace(/^\*\*|\*\*$/g, ''), bold: true, size: 20, color: BLUE, font: "Arial" })]
  });

  const bodyPara = (text, italic=false, color=BLACK) => new Paragraph({
    spacing: { before: 60, after: 60 },
    children: [new TextRun({ text, italic, color, size: 18, font: "Arial" })]
  });

  const bulletItem = (text) => new Paragraph({
    spacing: { before: 40, after: 40 },
    numbering: { reference: "bullets", level: 0 },
    children: [new TextRun({ text, size: 18, color: BLACK, font: "Arial" })]
  });

  const numberedItem = (text) => new Paragraph({
    spacing: { before: 40, after: 40 },
    numbering: { reference: "numbers", level: 0 },
    children: [new TextRun({ text, size: 18, color: BLACK, font: "Arial" })]
  });

  const children = [];

  // Title block
  children.push(
    new Table({
      width: { size: CONTENT, type: WidthType.DXA }, columnWidths: [6300, 2726],
      rows: [new TableRow({ children: [
        new TableCell({ borders, width: { size: 6300, type: WidthType.DXA },
          shading: { fill: NAVY, type: ShadingType.CLEAR }, margins: { top: 160, bottom: 160, left: 240, right: 240 },
          children: [
            new Paragraph({ children: [new TextRun({ text: "PATUMAHOE SCHOOL", bold: true, size: 28, color: WHITE, font: "Arial" })] }),
            new Paragraph({ children: [new TextRun({ text: "Principal Report to the Board of Trustees", size: 20, color: "CCDDEE", font: "Arial" })] }),
          ] }),
        new TableCell({ borders, width: { size: 2726, type: WidthType.DXA },
          shading: { fill: BLUE, type: ShadingType.CLEAR }, margins: { top: 160, bottom: 160, left: 200, right: 200 },
          verticalAlign: VerticalAlign.CENTER,
          children: [
            new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: `${(month||'').toUpperCase()} ${year}`, bold: true, size: 26, color: WHITE, font: "Arial" })] }),
            new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: term, size: 18, color: "CCDDEE", font: "Arial" })] }),
          ] }),
      ]})]
    }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 80, after: 200 },
      children: [new TextRun({ text: "Honouring our past, growing our future, learning together", italic: true, size: 18, color: "666666", font: "Arial" })] })
  );

  // Parse lines
  const lines = reportText.split('\n');
  let tableLines = [], inTable = false;

  function flushTable() {
    if (tableLines.length < 2) { tableLines.forEach(l => children.push(bodyPara(l, false, DGREY))); tableLines=[]; inTable=false; return; }
    const rows = tableLines
      .map(l => l.split('|').map(c=>c.trim()).filter((_,i,a)=>i>0&&i<a.length-1))
      .filter(r => r.length>0 && !r.every(c=>c.match(/^[-:]+$/)));
    if (!rows.length) { tableLines=[]; inTable=false; return; }
    const colCount = Math.max(...rows.map(r=>r.length));
    const colW = Math.floor(CONTENT/colCount);
    const colWidths = Array(colCount).fill(colW);
    colWidths[colCount-1] = CONTENT - colW*(colCount-1);
    children.push(new Table({
      width: { size: CONTENT, type: WidthType.DXA }, columnWidths: colWidths,
      rows: rows.map((row,ri) => new TableRow({ children:
        Array(colCount).fill(0).map((_,ci) => {
          const text = row[ci] || '';
          return ri===0 ? hdrCell(text,colWidths[ci]) : dataCell(text,colWidths[ci],ri%2===0?WHITE:GREY);
        })
      }))
    }));
    children.push(spacer());
    tableLines=[]; inTable=false;
  }

  for (const line of lines) {
    const t = line.trim();
    if (t==='') { if(inTable) flushTable(); children.push(spacer()); continue; }
    if (t.match(/^[═─=\-*]{4,}$/)) { if(inTable) flushTable(); continue; }
    if (t.startsWith('|')&&t.endsWith('|')) { inTable=true; tableLines.push(t); continue; }
    else if (inTable) flushTable();
    if (t.match(/^#{1,2} /)) { children.push(sectionHead(t)); continue; }
    if (t.match(/^### /)||t.match(/^\*\*[^*]+\*\*$/)) { children.push(subHead(t)); continue; }
    if (t.match(/^[A-Z][A-Z\s&\/–—\(\)]{5,}$/)&&!t.includes('[')) { children.push(sectionHead(t)); continue; }
    if (t.match(/^[-•]\s/)) { children.push(bulletItem(t.replace(/^[-•]\s/,''))); continue; }
    if (t.match(/^\d+\.\s/)) { children.push(numberedItem(t.replace(/^\d+\.\s/,''))); continue; }
    if (t.startsWith('[')&&t.endsWith(']')) { children.push(bodyPara(t,true,"999999")); continue; }
    children.push(bodyPara(t));
  }
  if (inTable) flushTable();

  // Sign-off
  children.push(
    new Paragraph({ spacing:{before:400,after:80}, border:{top:{style:BorderStyle.SINGLE,size:4,color:BLUE,space:8}}, children:[new TextRun("")] }),
    new Paragraph({ spacing:{before:80,after:40}, children:[
      new TextRun({ text:"Report prepared by: ", bold:true, size:18, font:"Arial" }),
      new TextRun({ text:"_________________________________     ", size:18, font:"Arial" }),
      new TextRun({ text:"Date: ", bold:true, size:18, font:"Arial" }),
      new TextRun({ text:"______________", size:18, font:"Arial" }),
    ]}),
    new Paragraph({ spacing:{before:60,after:0}, children:[new TextRun({
      text:`Generated: ${new Date().toLocaleDateString('en-NZ',{day:'numeric',month:'long',year:'numeric'})}`,
      size:16, color:"AAAAAA", italic:true, font:"Arial"
    })]})
  );

  const doc = new Document({
    numbering: { config: [
      { reference:"bullets", levels:[{level:0,format:LevelFormat.BULLET,text:"•",alignment:AlignmentType.LEFT,style:{paragraph:{indent:{left:720,hanging:360}}}}] },
      { reference:"numbers", levels:[{level:0,format:LevelFormat.DECIMAL,text:"%1.",alignment:AlignmentType.LEFT,style:{paragraph:{indent:{left:720,hanging:360}}}}] },
    ]},
    styles: { default: { document: { run: { font:"Arial", size:18 } } } },
    sections: [{
      properties: { page: { size:{width:11906,height:16838}, margin:{top:1080,right:1080,bottom:1080,left:1080} } },
      footers: { default: new Footer({ children:[new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [
          new TextRun({ text:"Patumahoe School  |  Principal Board Report — Confidential  |  Page ", size:16, color:"888888", font:"Arial" }),
          new TextRun({ children:[PageNumber.CURRENT], size:16, color:"888888", font:"Arial" }),
        ]
      })]}) },
      children
    }]
  });

  return await Packer.toBuffer(doc);
}

app.listen(PORT, () => console.log(`Patumahoe Board Report Service running on port ${PORT}`));
