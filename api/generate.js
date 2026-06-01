// AI Teacher Copilot — serverless backend
// Receives a teacher's structured request, calls Claude Haiku 4.5, returns the generated content.
// The API key lives in Vercel environment variables, never in the frontend.

export default async function handler(req, res) {
  // Only accept POST requests
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  // Read the teacher's request from the frontend
  const { framework, mode, topic, year, subject, stage, keyConcept, globalContext,
          prior, resources, qnum, qstyle, marks, len, learners, sname, tone, pedagogy } = req.body || {};

  if (!topic || !mode) {
    return res.status(400).json({ error: 'Missing topic or mode' });
  }

  // Build the system prompt — this is where the "Head of Department" thinking lives
  const systemPrompt = `You are an experienced Head of Department and expert teacher.
You write practical, classroom-ready materials with genuine pedagogical rigour — not generic AI filler.
You know the ${framework === 'myp' ? 'IB MYP' : 'Cambridge'} framework deeply and use its specific structures, vocabulary, and assessment language.
You are concise, specific, and useful. You avoid platitudes. You write as a colleague helping another teacher, not as a chatbot.
Output clean HTML only — use <h3>, <p>, <ul>, <ol>, <li>, <table>, <tr>, <th>, <td>, <b>, <i>. No markdown, no code fences, no preamble.`;

  // Build the user prompt — structured context the model can actually reason over
  const ctx = [];
  if (prior) ctx.push(`Prior knowledge: ${prior}`);
  if (resources) ctx.push(`Resources available: ${resources}`);
  if (framework === 'cambridge' && stage) ctx.push(`Cambridge stage: ${stage}`);
  if (framework === 'myp' && keyConcept) ctx.push(`MYP Key Concept: ${keyConcept}`);
  if (framework === 'myp' && globalContext) ctx.push(`MYP Global Context: ${globalContext}`);
  if (pedagogy && pedagogy !== 'standard') {
    const pedMap = {
      pbl: 'Use Project-Based Learning (PBL) structure with a driving question, sustained inquiry, authentic product, and reflection.',
      ssi: 'Frame around a Socio-Scientific Issue (SSI): include an ethical dilemma, multiple stakeholder perspectives, and evidence-based reasoning.',
      ngss: 'Align tightly to NGSS three-dimensional learning (DCI, SEP, CCC). Name the relevant performance expectation if you can.',
      edp: 'Use the Engineering Design Process: Ask, Imagine, Plan, Create, Improve. Make each stage concrete and time-boxed.'
    };
    if (pedMap[pedagogy]) ctx.push(`Pedagogical approach: ${pedMap[pedagogy]}`);
  }
  const contextBlock = ctx.length ? `\n\nContext:\n${ctx.map(c => '- ' + c).join('\n')}` : '';

  // Mode-specific instructions
  const modeInstructions = {
    lesson: `Produce a complete, classroom-ready lesson plan with these sections (use <h3> for each):
Learning Objectives, Success Criteria, Starter (5 min), Main Activity (${len || '40 min'}), Plenary (5 min), Key Vocabulary, Assessment & Next Steps, Resources Needed.
${framework === 'cambridge' ? 'Use Cambridge command words (state, describe, explain, suggest, evaluate) and reference appropriate Cambridge assessment objectives.' : 'Open with a Statement of Inquiry. Include factual/conceptual/debatable inquiry questions. Map to MYP criteria A–D. Reference relevant ATL skills.'}`,
    quiz: `Produce a ${qnum || 6}-question ${qstyle || 'mixed'} quiz on this topic.
${framework === 'cambridge' ? 'Ramp by Cambridge command-word demand (state → describe → explain → suggest → apply → evaluate). Allocate marks per question and end with a mark scheme.' : 'Map questions explicitly to MYP Criteria A (Knowing), B/C (Inquiring/Processing), and D (Reflecting). Allocate marks and end with 1–8 level descriptors.'}`,
    rubric: `Produce an assessment rubric as a clean <table>. ${marks ? `Total marks: ${marks}.` : ''}
${framework === 'cambridge' ? 'Use four performance levels (Beginning, Developing, Secure, Excelling) across Cambridge assessment objectives (AO1 Knowledge, AO2 Application, AO3 Analysis, plus Communication).' : 'Use MYP 1–8 level bands across one or more criteria (A–D). State the criterion clearly.'}
End with a short "how to use" note.`,
    differentiation: `Produce a concrete differentiation plan with three sections: Support (EAL), Support (SEN / scaffolded), Stretch (greater depth / extended inquiry), plus a brief "For all" section.
${learners ? `Specific focus learners: ${learners}.` : ''}
Give classroom-ready, named strategies — not generic advice. Reference the framework's assessment language.`,
    parent: `Produce a warm, professional parent message of 4–6 sentences as <p> tags.
Student: ${sname || 'the student'}. Tone: ${tone || 'warm and positive'}.
Sign off as "Mr Okoth, ${subject} Teacher". ${framework === 'cambridge' ? 'Briefly mention the Cambridge framework context.' : 'Briefly mention MYP learning and ATL skills.'}
Offer one concrete next step for home support.`
  };

  const userPrompt = `Create a ${mode} for a ${year} ${subject} class.
Topic: ${topic}${contextBlock}

${modeInstructions[mode] || modeInstructions.lesson}

Be specific to this topic — not generic. Make it something I could actually use in class tomorrow.`;

  try {
    const apiRes = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-api-key': process.env.ANTHROPIC_API_KEY,
        'anthropic-version': '2023-06-01'
      },
      body: JSON.stringify({
        model: 'claude-haiku-4-5',
        max_tokens: 1500,
        system: systemPrompt,
        messages: [{ role: 'user', content: userPrompt }]
      })
    });

    if (!apiRes.ok) {
      const errText = await apiRes.text();
      console.error('Anthropic API error:', apiRes.status, errText);
      return res.status(502).json({ error: 'AI service error', detail: apiRes.status });
    }

    const data = await apiRes.json();
    const html = (data.content || []).filter(b => b.type === 'text').map(b => b.text).join('\n');
    return res.status(200).json({ html });
  } catch (err) {
    console.error('Function error:', err);
    return res.status(500).json({ error: 'Server error' });
  }
}
