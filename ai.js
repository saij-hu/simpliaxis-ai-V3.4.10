// functions/api/ai.js
// Cloudflare Pages Function for Askly AI Assistant.
//
// GET  /api/ai -> health check
// POST /api/ai -> mode=parse | chat | explain
//
// The dashboard remains the source of truth for numerical answers. In parse mode this
// function converts natural language into a structured semantic query. In chat mode it
// handles general conversation. In explain mode it narrates an already-computed dashboard
// result. No raw 20,000+ row dataset is required by this endpoint.

const DEFAULT_MODEL = 'gpt-5.6-luna';
const MAX_HISTORY = 10;
const MAX_CONTEXT_CHARS = 12000;

export async function onRequestGet(context) {
  const key = String(context.env?.OPENAI_API_KEY || '').trim();
  const model = String(context.env?.ASKLY_AI_MODEL || DEFAULT_MODEL).trim();
  return json({
    function: 'working',
    openaiKeyConfigured: key.length > 0,
    model,
  });
}

export async function onRequestPost(context) {
  const env = context.env || {};
  const key = String(env.OPENAI_API_KEY || '').trim();
  if (!key) {
    return json({ error: 'openai_key_not_configured', openaiKeyConfigured: false }, 503);
  }

  let body;
  try {
    body = await context.request.json();
  } catch {
    return json({ error: 'invalid_json_body' }, 400);
  }

  const mode = String(body?.mode || 'chat').toLowerCase();
  const question = typeof body?.question === 'string' ? body.question.trim() : '';
  if (!question) return json({ error: 'missing_question' }, 400);
  if (question.length > 600) return json({ error: 'question_too_long' }, 400);
  if (!['parse', 'chat', 'explain'].includes(mode)) return json({ error: 'unsupported_mode' }, 400);

  const model = String(env.ASKLY_AI_MODEL || DEFAULT_MODEL).trim();

  if (mode === 'parse') {
    try {
      const result = await parseIntent({ key, model, question, context: body?.context, conversation: body?.conversation });
      return json({ mode, result });
    } catch (error) {
      console.error('Askly parse error:', error?.message || error);
      return json({ error: 'ai_parse_failed', message: 'AI interpretation failed' }, 502);
    }
  }

  if (mode === 'chat') {
    const instructions = buildChatInstructions();
    const input = buildConversationInput(question, body?.context, body?.conversation);
    return streamResponse({ key, model, instructions, input, maxOutputTokens: 500 });
  }

  // explain: the exact dashboard answer/result is supplied by the local engine.
  // The model is only allowed to explain that existing result and not introduce new figures.
  const answer = typeof body?.answer === 'string' ? body.answer.trim() : '';
  const result = body?.result && typeof body.result === 'object' ? body.result : null;
  if (!answer && !result) return json({ error: 'missing_computed_result' }, 400);

  const bounded = JSON.stringify({ answer, result, context: body?.context || {} }).slice(0, 6000);
  const instructions = [
    'You are Askly AI Assistant for the Simpliaxis Live Dashboard.',
    'The local dashboard engine already calculated the exact result.',
    'Explain the supplied result in 1-3 concise business sentences.',
    'Never invent, change, recalculate, round differently, or contradict a number.',
    'Never introduce a new metric or fact that is not supported by the supplied computed result.',
    'If the result does not support additional interpretation, return one concise sentence.',
  ].join('\n');
  const input = `User question:\n${question}\n\nComputed dashboard result:\n${bounded}`;
  return streamResponse({ key, model, instructions, input, maxOutputTokens: 260 });
}

async function parseIntent({ key, model, question, context, conversation }) {
  const contextText = JSON.stringify({
    context: context || {},
    conversation: sanitizeConversation(conversation),
  }).slice(0, MAX_CONTEXT_CHARS);

  const instructions = [
    'You are the semantic router for Askly AI Assistant, a live Simpliaxis dashboard.',
    'Understand the user naturally, including capitalization variants, shorthand, typos, follow-up references, date phrases, and multiple filters.',
    'For dashboard questions, return a structured query so the browser can execute it against the existing live dashboard engine.',
    'Never return a dashboard number. Never calculate a dashboard metric yourself.',
    'Never invent a field value. Values must be chosen only when supported by the supplied availableValues/context.',
    'Treat capitalization as irrelevant: Direct, direct and DIRECT are the same value; CPC/cpc/Cpc are the same value; Akhilesh/akhilesh/AKHILESH are the same value.',
    'Primary mappings: specific executive/name -> CRE; CPC/Organic/Direct lead questions -> CHANNELS; paid/payment count -> TOTAL PAYMENTS; revenue/total this month -> REVENUE.',
    'Lead questions use Assign DATE. Payment and revenue questions use Payment Date.',
    'Relative periods like last 3 months are interpreted from today by the existing client date engine; do not fabricate exact dates when a normalized natural-language question is enough.',
    'If the user asks a genuine general/conversational question that does not require dashboard data, return kind=chat.',
    'If the user asks a dashboard question that cannot be mapped safely, return kind=clarify with one minimal clarification question.',
    'Return ONLY valid JSON with this shape:',
    '{"kind":"dashboard|chat|clarify","reply":"optional string","query":{"metric":"optional metric key","metric2":"optional metric key","filters":{"field":["value"]},"groupBy":"optional field","sort":"asc|desc","limit":null,"isTarget":false,"wantsNarrative":false,"normalizedQuestion":"optional normalized question"}}',
    'Allowed metric keys: leadCount, paymentCount, paymentCountIncludingOld, oldPayments, revenue, avgTicketSize, avgPaymentDuration, lostLeads, unpaidLeads, junkLeads, validLeads, conversionRate.',
    'Allowed filter fields: assignDate, paymentDate, cre, course, country, region, channel, leadSource, paymentSource, paymentStatus, leadStatus, leadValidity, team.',
    'Allowed groupBy fields: cre, course, country, region, channel, leadSource, paymentSource, paymentStatus, leadStatus, leadValidity, team.',
  ].join('\n');

  const input = `Dashboard semantic context:\n${contextText}\n\nUser question:\n${question}`;
  const data = await responsesRequest({ key, model, instructions, input, maxOutputTokens: 700, stream: false });
  const raw = getOutputText(data);
  const parsed = safeJsonFromText(raw);
  if (!parsed || typeof parsed !== 'object') {
    throw new Error('Model returned invalid semantic JSON');
  }
  return parsed;
}

function buildChatInstructions() {
  return [
    'You are Askly AI Assistant for the Simpliaxis Live Dashboard.',
    'Talk naturally like a helpful ChatGPT-style assistant.',
    'Be concise, clear, and useful.',
    'You may explain how the dashboard works, answer general questions about Askly, or help the user phrase a dashboard query.',
    'Do not claim to know a live dashboard number unless the browser has supplied a computed result.',
    'Do not invent live metrics, targets, users, payments, revenue, or other dashboard data.',
    'Respect the user permissions and the dashboard context supplied to you.',
  ].join('\n');
}

function buildConversationInput(question, context, conversation) {
  const payload = {
    context: context || {},
    conversation: sanitizeConversation(conversation),
    question,
  };
  return JSON.stringify(payload).slice(0, MAX_CONTEXT_CHARS);
}

function sanitizeConversation(conversation) {
  if (!Array.isArray(conversation)) return [];
  return conversation.slice(-MAX_HISTORY).map(item => ({
    role: item?.role === 'assistant' ? 'assistant' : 'user',
    text: String(item?.text || '').slice(0, 800),
  }));
}

async function streamResponse({ key, model, instructions, input, maxOutputTokens }) {
  let upstream;
  try {
    upstream = await responsesRequestRaw({ key, model, instructions, input, maxOutputTokens, stream: true });
  } catch (error) {
    console.error('Askly streaming request failed:', error?.message || error);
    return json({ error: 'server_error' }, 500);
  }

  if (!upstream.ok) {
    const detail = await upstream.text().catch(() => '');
    console.error('OpenAI upstream error:', upstream.status, detail.slice(0, 500));
    return json({ error: 'upstream_error' }, 502);
  }
  if (!upstream.body) return json({ error: 'empty_upstream_response' }, 502);

  const encoder = new TextEncoder();
  const decoder = new TextDecoder();
  let buffered = '';

  const stream = new ReadableStream({
    async start(controller) {
      const reader = upstream.body.getReader();
      try {
        while (true) {
          const { done, value } = await reader.read();
          if (done) break;
          buffered += decoder.decode(value, { stream: true });
          const lines = buffered.split('\n');
          buffered = lines.pop() || '';
          for (const line of lines) {
            const trimmed = line.trim();
            if (!trimmed.startsWith('data:')) continue;
            const payload = trimmed.slice(5).trim();
            if (!payload || payload === '[DONE]') continue;
            try {
              const event = JSON.parse(payload);
              if (event?.type === 'response.output_text.delta' && typeof event.delta === 'string') {
                controller.enqueue(encoder.encode(event.delta));
              }
            } catch {
              // Ignore malformed/partial SSE fragments; the next read will continue the stream.
            }
          }
        }
        controller.close();
      } catch (error) {
        console.error('Askly stream relay failed:', error?.message || error);
        controller.error(error);
      }
    },
  });

  return new Response(stream, {
    status: 200,
    headers: {
      'Content-Type': 'text/plain; charset=utf-8',
      'Cache-Control': 'no-store',
      'X-Askly-Model': model,
      'X-Askly-Source': 'OpenAI Responses API',
    },
  });
}

async function responsesRequest({ key, model, instructions, input, maxOutputTokens, stream }) {
  const response = await responsesRequestRaw({ key, model, instructions, input, maxOutputTokens, stream });
  if (!response.ok) {
    const detail = await response.text().catch(() => '');
    throw new Error(`OpenAI ${response.status}: ${detail.slice(0, 500)}`);
  }
  return response.json();
}

async function responsesRequestRaw({ key, model, instructions, input, maxOutputTokens, stream }) {
  return fetch('https://api.openai.com/v1/responses', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${key}`,
    },
    body: JSON.stringify({
      model,
      instructions,
      input,
      max_output_tokens: maxOutputTokens,
      stream: !!stream,
    }),
  });
}

function getOutputText(data) {
  if (typeof data?.output_text === 'string') return data.output_text;
  const chunks = [];
  for (const item of data?.output || []) {
    for (const content of item?.content || []) {
      if (typeof content?.text === 'string') chunks.push(content.text);
    }
  }
  return chunks.join('\n');
}

function safeJsonFromText(text) {
  const raw = String(text || '').trim();
  if (!raw) return null;
  try { return JSON.parse(raw); } catch {}
  const fenced = raw.replace(/^```json\s*/i, '').replace(/^```\s*/i, '').replace(/```$/,'').trim();
  try { return JSON.parse(fenced); } catch {}
  const first = fenced.indexOf('{');
  const last = fenced.lastIndexOf('}');
  if (first >= 0 && last > first) {
    try { return JSON.parse(fenced.slice(first, last + 1)); } catch {}
  }
  return null;
}

function json(obj, status = 200) {
  return new Response(JSON.stringify(obj), {
    status,
    headers: {
      'Content-Type': 'application/json; charset=utf-8',
      'Cache-Control': 'no-store',
    },
  });
}
