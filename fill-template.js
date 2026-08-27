// AI Teacher Copilot — Alpha Cambridge template filler (production)
// Calls Claude to generate structured lesson content, then surgically fills the school template.

import fs from 'fs';
import path from 'path';
import JSZip from 'jszip';

export const config = { api: { bodyParser: { sizeLimit: '8mb' } } };

const RESOURCE_OPTIONS = ['Worksheet','Smart Board','Data Show','Presentations','Manipulative','Photo and cards','Video','Other'];
const STRATEGY_OPTIONS = ['Discussion','Problem Solving','Cooperative Learning','Direct Teaching','Photo','Hands on Activity','Modelling','learning Station','Other','Role Play','Brainstorming','Software'];

// Word splits some labels across multiple <w:t> runs; use the unique first-run fragment.
const FRAG = {
  'Worksheet':'Worksheet','Smart Board':'Smart Board','Data Show':'Data Show','Presentations':'Presentations',
  'Manipulative':'Manipulative','Photo and cards':'Photo','Video':'Video','Other':'Other',
  'Discussion':'Discussion','Problem Solving':'Problem Solving','Cooperative Learning':'Cooperative Learning',
  'Direct Teaching':'Direct','Photo':'Photo','Hands on Activity':'Hands on',
  'Modelling':'Modelling','learning Station':'learning Station','Role Play':'Role Play',
  'Brainstorming':'Brainstorming','Software':'Software'
};

export default async function handler(req, res) {
  if (req.method !== 'POST') return res.status(405).json({ error: 'Method not allowed' });

  const {
    topic, year, subject, term, week, lessonOfUnit, teacherName,
    framework, stage, pedagogy, prior, resources,
    resourceFile
  } = req.body || {};
  if (!topic) return res.status(400).json({ error: 'Missing topic' });

  // --- Build AI prompt ---
  const system = `You are an experienced Head of Department writing a structured lesson plan to be inserted into a school's existing Word template.
Output ONLY valid JSON — no preamble, no markdown, no code fences.
Be specific to the topic. Each field is sized for the cell it fills. Use ${framework === 'myp' ? 'IB MYP' : 'Cambridge'} language. Stage: ${stage || 'Lower Secondary'}.
${pedagogy && pedagogy !== 'standard' ? `Pedagogy: ${pedagogy.toUpperCase()}.` : ''}

JSON schema (all fields required):
{
  "lessonTopic": "string, <=80 chars",
  "learningOutcomes": "3-4 short outcome bullets joined by newlines, each starts with a verb",
  "starter": "2-3 sentences describing the opening activity",
  "mainActivities": [ { "objectives": "...", "strategy": "...", "studentActivity": "...", "assessment": "...", "time": "e.g. 15 MIN" } ],
  "closure": "2-3 sentences for the plenary",
  "assignment": "one-line homework",
  "selfReflection": "one short prompt for teacher reflection",
  "integration": "subjects this links to, e.g. 'Mathematics and English'",
  "nationalValue": "one short value-of-the-week statement in quotes",
  "resourcesToTick": ["items from: ${RESOURCE_OPTIONS.join(', ')}"],
  "strategiesToTick": ["items from: ${STRATEGY_OPTIONS.join(', ')}"]
}
Provide 2-3 main activities. Tick only resources/strategies you actually use.`;

  const userText = `Generate JSON for this lesson:

Topic: ${topic}
Year/Grade: ${year}
Subject: ${subject}
${term ? `Term: ${term}` : ''}${week ? `, Week: ${week}` : ''}
${lessonOfUnit ? `Lesson of unit: ${lessonOfUnit}` : ''}
${prior ? `Prior knowledge: ${prior}` : ''}
${resources ? `Resources available: ${resources}` : ''}
${resourceFile ? `A supporting resource is attached — use it to inform the lesson.` : ''}`;

  const userContent = [];
  if (resourceFile && resourceFile.base64 && resourceFile.mediaType) {
    if (resourceFile.mediaType.startsWith('image/')) {
      userContent.push({ type:'image', source:{ type:'base64', media_type:resourceFile.mediaType, data:resourceFile.base64 } });
    } else if (resourceFile.mediaType === 'application/pdf') {
      userContent.push({ type:'document', source:{ type:'base64', media_type:'application/pdf', data:resourceFile.base64 } });
    }
  }
  userContent.push({ type:'text', text:userText });

  let plan;
  try {
    const apiRes = await fetch('https://api.anthropic.com/v1/messages', {
      method:'POST',
      headers:{ 'Content-Type':'application/json', 'x-api-key':process.env.ANTHROPIC_API_KEY, 'anthropic-version':'2023-06-01' },
      body: JSON.stringify({
        model:'claude-haiku-4-5', max_tokens:2500, system,
        messages:[{ role:'user', content:userContent }]
      })
    });
    if (!apiRes.ok) {
      console.error('Anthropic error:', apiRes.status, await apiRes.text());
      return res.status(502).json({ error:'AI service error', detail:apiRes.status });
    }
    const data = await apiRes.json();
    const text = (data.content||[]).filter(b=>b.type==='text').map(b=>b.text).join('').trim();
    const cleaned = text.replace(/^```json\s*/i,'').replace(/^```\s*/,'').replace(/```\s*$/,'').trim();
    plan = JSON.parse(cleaned);
  } catch (err) {
    console.error('AI/JSON error:', err);
    return res.status(500).json({ error:'AI response could not be parsed' });
  }

  // --- Load template & fill ---
  let templateBytes;
  try {
    templateBytes = fs.readFileSync(path.join(process.cwd(),'templates','alpha-cambridge.docx'));
  } catch (err) {
    console.error('Template missing:', err);
    return res.status(500).json({ error:'Template file not found on server' });
  }

  let docxBuffer;
  try {
    docxBuffer = await fillTemplate(templateBytes, plan, {
      teacherName: teacherName || 'Mr. Dickens',
      year, term, week, subject, lessonOfUnit
    });
  } catch (err) {
    console.error('Fill error:', err);
    return res.status(500).json({ error:'Could not fill the template', detail:String(err.message||err) });
  }

  res.setHeader('Content-Type','application/vnd.openxmlformats-officedocument.wordprocessingml.document');
  res.setHeader('Content-Disposition',`attachment; filename="LessonPlan_${(topic||'lesson').replace(/[^a-z0-9]/gi,'_').slice(0,40)}.docx"`);
  res.setHeader('Content-Length', docxBuffer.length);
  return res.status(200).send(docxBuffer);
}

// ===== Template-fill engine (verified against Alpha Cambridge template) =====

function escapeXml(s){return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&apos;');}

function rowRangeAt(xml, pos) {
  const opens = [...xml.matchAll(/<w:tr\b[^>]*>/g)];
  const closes = [...xml.matchAll(/<\/w:tr>/g)];
  const events = [];
  for (const m of opens) events.push({ pos:m.index, end:m.index+m[0].length, type:'open' });
  for (const m of closes) events.push({ pos:m.index, end:m.index+m[0].length, type:'close' });
  events.sort((a,b)=>a.pos-b.pos);
  const stack = [];
  for (const e of events) {
    if (e.type==='open') stack.push(e.pos);
    else { const start=stack.pop(), end=e.end; if (start<=pos&&pos<end) return {start,end}; }
  }
  return null;
}
function rowRange(xml, anchor){ const i=xml.indexOf(anchor); return i===-1?null:rowRangeAt(xml,i); }

function topLevelCellsInRow(xml, rowStart, rowEnd) {
  const row = xml.slice(rowStart, rowEnd);
  const opens = [...row.matchAll(/<w:tc\b[^>]*>/g)];
  const closes = [...row.matchAll(/<\/w:tc>/g)];
  const events = [];
  for (const m of opens) events.push({ pos:m.index, end:m.index+m[0].length, type:'open' });
  for (const m of closes) events.push({ pos:m.index, end:m.index+m[0].length, type:'close' });
  events.sort((a,b)=>a.pos-b.pos);
  const stack = [], result = [];
  for (const e of events) {
    if (e.type==='open') stack.push(e.pos);
    else { const start=stack.pop(), end=e.end; if (stack.length===0) result.push({start:rowStart+start,end:rowStart+end}); }
  }
  return result;
}

function setCellText(xml, cellStart, cellEnd, newText) {
  const cell = xml.slice(cellStart, cellEnd);
  const tcPrMatch = cell.match(/<w:tcPr>[\s\S]*?<\/w:tcPr>/);
  const tcPr = tcPrMatch ? tcPrMatch[0] : '';
  const paras = String(newText).split(/\r?\n/).filter(l=>l.length).map(line =>
    `<w:p><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="20"/></w:rPr><w:t xml:space="preserve">${escapeXml(line)}</w:t></w:r></w:p>`
  ).join('') || `<w:p><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/></w:rPr><w:t></w:t></w:r></w:p>`;
  return xml.slice(0,cellStart) + `<w:tc>${tcPr}${paras}</w:tc>` + xml.slice(cellEnd);
}

function tickCheckboxAtPos(xml, pos) {
  const rr = rowRangeAt(xml, pos);
  if (!rr) return xml;
  const row = xml.slice(rr.start, rr.end);
  const localPos = pos - rr.start;
  const before = row.slice(0, localPos);
  const lastBox = before.lastIndexOf('w:char="F072"');
  if (lastBox === -1) return xml;
  const updated = row.slice(0, lastBox) + 'w:char="F0FE"' + row.slice(lastBox + 'w:char="F072"'.length);
  return xml.slice(0, rr.start) + updated + xml.slice(rr.end);
}

function findVMergeRestartCell(xml, anchorInSameRow) {
  const ai = xml.indexOf(anchorInSameRow);
  if (ai === -1) return null;
  const rr = rowRangeAt(xml, ai);
  if (!rr) return null;
  const cells = topLevelCellsInRow(xml, rr.start, rr.end);
  for (const c of cells) {
    if (/<w:vMerge\s+w:val="restart"\s*\/>/.test(xml.slice(c.start, c.end))) return c;
  }
  return null;
}

async function fillTemplate(templateBytes, plan, meta) {
  const zip = await JSZip.loadAsync(templateBytes);
  let xml = await zip.file('word/document.xml').async('string');

  // 1) Metadata row
  {
    const rr = rowRange(xml, 'Mr. Dickens');
    if (rr) {
      const cells = topLevelCellsInRow(xml, rr.start, rr.end);
      const values = [meta.teacherName, meta.year, `${meta.term||'TERM 1'}   WEEK ${meta.week||''}`.trim(), (meta.subject||'SCIENCE').toUpperCase()];
      for (let i = cells.length - 1; i >= 0 && i < values.length; i--) {
        xml = setCellText(xml, cells[i].start, cells[i].end, values[i]);
      }
    }
  }

  // 2) Lesson topic + unit
  {
    const rr = rowRange(xml, 'LESSON 1/4');
    if (rr) {
      const cells = topLevelCellsInRow(xml, rr.start, rr.end);
      if (cells.length >= 2) {
        xml = setCellText(xml, cells[1].start, cells[1].end, meta.lessonOfUnit || 'LESSON 1/4');
        const rr2 = rowRange(xml, meta.lessonOfUnit || 'LESSON 1/4');
        const c2 = rr2 ? topLevelCellsInRow(xml, rr2.start, rr2.end) : null;
        if (c2 && c2.length >= 1) xml = setCellText(xml, c2[0].start, c2[0].end, plan.lessonTopic || '');
      }
    }
  }

  // 3) Learning Outcomes — vMerge restart cell in row with "Supporting"
  {
    const cell = findVMergeRestartCell(xml, 'Supporting');
    if (cell) {
      const cellXml = xml.slice(cell.start, cell.end);
      const tcPrMatch = cellXml.match(/<w:tcPr>[\s\S]*?<\/w:tcPr>/);
      const tcPr = tcPrMatch ? tcPrMatch[0] : '';
      const paras = String(plan.learningOutcomes||'').split(/\r?\n/).filter(l=>l.length).map(line =>
        `<w:p><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="20"/></w:rPr><w:t xml:space="preserve">${escapeXml(line)}</w:t></w:r></w:p>`
      ).join('');
      xml = xml.slice(0,cell.start) + `<w:tc>${tcPr}${paras}</w:tc>` + xml.slice(cell.end);
    }
  }

  // 4 & 5) Resources & Strategies checkboxes, scoped by region to avoid cross-grid collisions
  const resStart = xml.indexOf('>Resources<');
  const stratStart = xml.indexOf('Teaching &amp; Learning');
  const stratEnd = xml.indexOf('Starter', stratStart);
  function tickInRegion(regionStart, regionEnd, optionsList, requested) {
    for (const opt of requested) {
      if (!optionsList.includes(opt)) continue;
      const frag = FRAG[opt];
      let i = regionStart, pos = -1;
      while (true) {
        const found = xml.indexOf(frag, i);
        if (found === -1 || found >= regionEnd) break;
        const charBefore = xml[found - 1];
        if (charBefore === '>' || /\s/.test(charBefore)) { pos = found; break; }
        i = found + 1;
      }
      if (pos !== -1) xml = tickCheckboxAtPos(xml, pos);
    }
  }
  if (resStart !== -1 && stratStart !== -1) tickInRegion(resStart, stratStart, RESOURCE_OPTIONS, plan.resourcesToTick||[]);
  if (stratStart !== -1 && stratEnd !== -1) tickInRegion(stratStart, stratEnd, STRATEGY_OPTIONS, plan.strategiesToTick||[]);

  // 6) Starter
  {
    const rr = rowRange(xml, '3 Part-lesson');
    if (rr) {
      const cells = topLevelCellsInRow(xml, rr.start, rr.end);
      if (cells.length >= 2) xml = setCellText(xml, cells[1].start, cells[1].end, plan.starter || '');
    }
  }

  // 7) Main activities
  {
    const sampleAnchor = 'Students will take turns sharing their knowledge';
    const rr = rowRange(xml, sampleAnchor);
    if (rr) {
      const sampleRow = xml.slice(rr.start, rr.end);
      const activities = (plan.mainActivities && plan.mainActivities.length ? plan.mainActivities : [{}]).slice(0,4);
      const newRows = activities.map(a => buildActivityRow(sampleRow, a)).join('');
      xml = xml.slice(0, rr.start) + newRows + xml.slice(rr.end);
    }
  }

  // 8) Closure
  {
    const rr = rowRange(xml, 'Summary of the lesson');
    if (rr) {
      const cells = topLevelCellsInRow(xml, rr.start, rr.end);
      if (cells.length >= 1) xml = setCellText(xml, cells[0].start, cells[0].end, plan.closure || '');
    }
  }

  // 9) Assignment
  {
    const rr = rowRange(xml, 'Portal Homework');
    if (rr) {
      const cells = topLevelCellsInRow(xml, rr.start, rr.end);
      if (cells.length >= 1) xml = setCellText(xml, cells[0].start, cells[0].end, plan.assignment || 'Portal Homework');
    }
  }

  // 10) Self-Reflection / Integration / National Value
  {
    const rr = rowRange(xml, 'Social studies and English');
    if (rr) {
      const cells = topLevelCellsInRow(xml, rr.start, rr.end);
      if (cells.length >= 3) {
        xml = setCellText(xml, cells[2].start, cells[2].end, plan.nationalValue || '');
        const valAnchor = plan.nationalValue ? plan.nationalValue.slice(0, 12).replace(/^"/,'') : 'name, style';
        const rrA = rowRange(xml, valAnchor);
        if (rrA) {
          const c = topLevelCellsInRow(xml, rrA.start, rrA.end);
          if (c.length >= 3) {
            xml = setCellText(xml, c[1].start, c[1].end, plan.integration || '');
            const integAnchor = plan.integration ? plan.integration.slice(0, 10) : 'Social';
            const rrB = rowRange(xml, integAnchor);
            if (rrB) {
              const c2 = topLevelCellsInRow(xml, rrB.start, rrB.end);
              if (c2.length >= 3) xml = setCellText(xml, c2[0].start, c2[0].end, plan.selfReflection || '');
            }
          }
        }
      }
    }
  }

  zip.file('word/document.xml', xml);
  return zip.generateAsync({ type:'nodebuffer', compression:'DEFLATE' });
}

function buildActivityRow(sampleRow, a) {
  const trOpenMatch = sampleRow.match(/^<w:tr\b[^>]*>(<w:trPr>[\s\S]*?<\/w:trPr>)?/);
  const trOpen = trOpenMatch ? sampleRow.slice(0, trOpenMatch[0].length) : '<w:tr>';
  const tcPrs = [];
  const opens = [...sampleRow.matchAll(/<w:tc\b[^>]*>/g)];
  const closes = [...sampleRow.matchAll(/<\/w:tc>/g)];
  const events = [];
  for (const m of opens) events.push({ pos:m.index, end:m.index+m[0].length, type:'open' });
  for (const m of closes) events.push({ pos:m.index, end:m.index+m[0].length, type:'close' });
  events.sort((a,b)=>a.pos-b.pos);
  const stack = [];
  for (const e of events) {
    if (e.type==='open') stack.push(e.pos);
    else {
      const start = stack.pop(), end = e.end;
      if (stack.length === 0) {
        const inner = sampleRow.slice(start, end);
        const tcPrMatch = inner.match(/<w:tcPr>[\s\S]*?<\/w:tcPr>/);
        tcPrs.push(tcPrMatch ? tcPrMatch[0] : '');
      }
    }
  }
  const texts = [a.objectives, a.strategy, a.studentActivity, a.assessment, a.time].map(t => String(t||''));
  return `${trOpen}${tcPrs.slice(0,5).map((tcPr,i) => {
    const paras = texts[i].split(/\r?\n/).filter(l=>l.length).map(line =>
      `<w:p><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="20"/></w:rPr><w:t xml:space="preserve">${escapeXml(line)}</w:t></w:r></w:p>`
    ).join('') || `<w:p><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/></w:rPr><w:t></w:t></w:r></w:p>`;
    return `<w:tc>${tcPr}${paras}</w:tc>`;
  }).join('')}</w:tr>`;
}
