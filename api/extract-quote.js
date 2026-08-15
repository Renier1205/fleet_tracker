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

const EXTRACTION_PROMPT = `You are reading a scanned supplier quote for mining equipment parts.

Extract every part/price line item from this quote. Return ONLY a JSON array (no other text, no markdown code fences) where each item is an object with exactly these keys:
- "supplier": the supplier/vendor company name as printed on the quote (same value for every line unless the document genuinely lists different suppliers per line)
- "part_no": the part or item number/code for that line
- "part_description": a short description of the part
- "price": the unit price as a plain number (no currency symbols, no commas) - if the line shows quantity and a line total, calculate and return the unit price, not the total

Rules:
- Only include real part/price line items. Skip headers, page totals, subtotals, tax lines, delivery charges, and terms & conditions.
- If a text field can't be determined for a line, use an empty string. If price can't be determined, use null.
- If you cannot find any part/price line items at all, return an empty array: []
- Return ONLY the JSON array. Nothing before it, nothing after it.`;

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
    let items;
    try {
      items = JSON.parse(cleaned);
    } catch (parseErr) {
      res.status(502).json({ error: "Couldn't parse the AI's response as JSON", raw: textBlock.text });
      return;
    }

    if (!Array.isArray(items)) {
      res.status(502).json({ error: "AI response wasn't a JSON array as expected", raw: textBlock.text });
      return;
    }

    res.status(200).json({ items });
  } catch (err) {
    res.status(500).json({ error: err.message || String(err) });
  }
}
