// /api/extract-quote.js
//
// Vercel serverless function. Runs server-side only - this is the one
// place the Anthropic API key is used, since it must never be shipped
// to the browser. The frontend uploads the scanned quote here as
// base64; this function asks Claude to read it and return the line
// items as JSON.
//
// Requires an ANTHROPIC_API_KEY environment variable set in the Vercel
// project (Project Settings -> Environment Variables). Get a key from
// https://console.anthropic.com - this is separate from a Claude.ai
// subscription and bills per request (extracting one quote is a small
// fraction of a cent to a few cents, depending on the model and image
// size).

const EXTRACTION_PROMPT = `You are reading a scanned supplier quote for mining equipment parts. The quote may be in any language, or a mix of languages (e.g. Arabic and English side by side).

First, identify the primary language the quote is written in.

Then extract every part/price line item. Return ONLY a JSON object (no other text, no markdown code fences) with exactly this shape:
{
  "detected_language": "the primary language of the document, in English, e.g. \\"Arabic\\", \\"English\\", \\"Arabic and English\\"",
  "items": [
    {
      "supplier": "the supplier/vendor company name as printed on the quote - keep company/brand names as printed, do not translate them",
      "part_no": "the part or item number/code for that line, as printed",
      "part_description": "ONLY the part name/type itself (e.g. \\"Alternator\\", \\"Hydraulic Filter\\"), translated into English if the original is in another language",
      "currency": "the currency of the price, as a 3-letter ISO code if it can be determined (e.g. \\"ZAR\\", \\"USD\\", \\"SAR\\", \\"EUR\\") - infer it from a currency symbol, currency code, or country/company context printed on the quote; if genuinely undeterminable, use an empty string",
      "price": "the unit price as a plain number (no currency symbols, no commas) - if the line shows quantity and a line total, calculate and return the unit price, not the total"
    }
  ]
}

Rules:
- Only include real part/price line items. Skip headers, page totals, subtotals, tax lines, delivery/discount lines, and terms & conditions.
- part_description must contain ONLY the part's name/type - nothing else. Specifically exclude, even if printed directly next to or below the part name on the same line or cell: availability/lead-time notes (e.g. "one week from receiving date"), delivery or shipping terms, warranty text, packaging/quantity notes, remarks in parentheses or brackets, and any other annotation. If such text is mixed in with the part name, strip it out and keep only the part name.
- Translate part_description into clear English regardless of the source language. Do not translate the supplier's company name or the part number/code - keep those exactly as printed.
- Every line item on the same quote is normally the same currency - if you can determine it from any line (a symbol, code, or heading like "Prices in USD"), apply it consistently across all items unless a specific line clearly states otherwise.
- If a text field can't be determined for a line, use an empty string. If price can't be determined, use null.
- If you cannot find any part/price line items at all, return {"detected_language": "...", "items": []}.
- Return ONLY the JSON object described above. Nothing before it, nothing after it.`;

export default async function handler(req, res) {
  if (req.method !== "POST") {
    res.status(405).json({ error: "Method not allowed" });
    return;
  }

  const { imageBase64, mediaType } = req.body || {};
  if (!imageBase64 || !mediaType) {
    res.status(400).json({ error: "Missing imageBase64 or mediaType" });
    return;
  }

  const apiKey = process.env.ANTHROPIC_API_KEY;
  if (!apiKey) {
    res.status(500).json({ error: "Server is missing ANTHROPIC_API_KEY - add it in Vercel Project Settings > Environment Variables, then redeploy." });
    return;
  }

  const isPdf = mediaType === "application/pdf";

  try {
    const response = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "x-api-key": apiKey,
        "anthropic-version": "2023-06-01",
      },
      body: JSON.stringify({
        model: "claude-sonnet-5",
        max_tokens: 2000,
        messages: [
          {
            role: "user",
            content: [
              {
                type: isPdf ? "document" : "image",
                source: { type: "base64", media_type: mediaType, data: imageBase64 },
              },
              { type: "text", text: EXTRACTION_PROMPT },
            ],
          },
        ],
      }),
    });

    const data = await response.json();
    if (!response.ok) {
      res.status(response.status).json({ error: data?.error?.message || "Anthropic API request failed" });
      return;
    }

    const textBlock = (data.content || []).find((c) => c.type === "text");
    if (!textBlock) {
      res.status(502).json({ error: "AI response didn't contain any text" });
      return;
    }

    const cleaned = textBlock.text.trim().replace(/^```(?:json)?\s*/i, "").replace(/```\s*$/i, "").trim();
    let parsed;
    try {
      parsed = JSON.parse(cleaned);
    } catch (parseErr) {
      res.status(502).json({ error: "Couldn't parse the AI's response as JSON", raw: textBlock.text });
      return;
    }

    const items = Array.isArray(parsed) ? parsed : parsed?.items; // tolerate either shape
    if (!Array.isArray(items)) {
      res.status(502).json({ error: "AI response wasn't in the expected shape", raw: textBlock.text });
      return;
    }

    res.status(200).json({ items, detectedLanguage: parsed?.detected_language || null });
  } catch (err) {
    res.status(500).json({ error: err.message || String(err) });
  }
}
