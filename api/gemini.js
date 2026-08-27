// Vercel Serverless Function — secure proxy to the Google Gemini API.
// The API key stays here on the server (as an Environment Variable),
// never exposed to the browser. The frontend calls this endpoint
// (same-origin, no CORS issues) instead of Google's API directly.

const GEMINI_MODEL = "gemini-3.6-flash";

export default async function handler(req, res) {
  if (req.method !== "POST") {
    return res.status(405).json({ error: "Method not allowed" });
  }

  const apiKey = process.env.GEMINI_API_KEY;
  if (!apiKey) {
    return res.status(500).json({
      error: "GEMINI_API_KEY is not set. Add it in Vercel → Project → Settings → Environment Variables, then redeploy.",
    });
  }

  const { systemInstruction, contents, maxOutputTokens } = req.body || {};
  if (!contents) {
    return res.status(400).json({ error: "Missing 'contents' in request body." });
  }

  try {
    const upstream = await fetch(
      `https://generativelanguage.googleapis.com/v1beta/models/${GEMINI_MODEL}:generateContent`,
      {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "x-goog-api-key": apiKey,
        },
        body: JSON.stringify({
          contents,
          ...(systemInstruction ? { systemInstruction: { parts: [{ text: systemInstruction }] } } : {}),
          generationConfig: { maxOutputTokens: maxOutputTokens || 1000 },
        }),
      }
    );
    const data = await upstream.json();
    return res.status(upstream.status).json(data);
  } catch (err) {
    return res.status(500).json({ error: String(err) });
  }
}
