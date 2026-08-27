# CEO Command Center — standalone project

This is the same CEO Command Center app, packaged as a normal React project
so it can run on real hosting (Vercel, Netlify, etc.) instead of inside the
chat sandbox. Once deployed here, the automatic silent archive/notify calls
(no tap needed) will actually work, because the sandbox restriction that
required one-tap buttons doesn't exist on real hosting.

## AI Agent — now powered by Google Gemini

The AI Assistant chat and every Team Agent's auto-answer are powered by
Google's Gemini API, kept safely server-side (never in browser code) via
the included `api/gemini.js` serverless function.

**Setup (2 minutes):**
1. Go to [aistudio.google.com/apikey](https://aistudio.google.com/apikey), sign in with your Google account, click **Create API key**. Gemini has a generous free tier, which fits this app's usage well.
2. In your Vercel project → **Settings → Environment Variables**, add:
   - Key: `GEMINI_API_KEY`
   - Value: (paste your key)
   - Environment: Production (and Preview if you want)
3. Redeploy (Deployments tab → latest deployment → ⋯ → Redeploy).

### Two things that make the agent genuinely smarter over time
- **Skill Profile** — in the Chat tab, click "Skill & Memori" to freely rewrite the assistant's persona/skills in your own words. Saved permanently.
- **Agent Memory** — after each chat exchange, the assistant quietly extracts one short fact worth remembering about your business/preferences and saves it. Future conversations automatically include everything it has learned so far, so context genuinely accumulates the more you use it. You can view/delete individual memories or clear them anytime in the same panel.

## Settings now sync across devices via your Google Sheet

Integrations, teams, CEO Policy, Skill Profile, and Agent Memory are all
saved permanently in the browser (localStorage) **and** mirrored to a
"Settings" tab in the same Google Sheet used for archiving — as long as a
Google Apps Script bridge URL is set. Open the app on a different device,
paste the same bridge URL once, and everything else pulls in automatically.
This requires the updated `GAS-CEO-Archive-Bot.gs` (redeploy your Apps
Script with the new version — same URL, no need to change it in the app).

## Local development (optional, needs Node.js installed)
```
npm install
npm run dev
```

## Deploy to Vercel (recommended, no local setup needed)
See the deployment guide provided in chat — short version:
1. Push this folder to a new GitHub repository.
2. Go to vercel.com → New Project → Import that repository.
3. Vercel auto-detects Vite and deploys. Done — you get a live URL.
4. Add your `GEMINI_API_KEY` as described above, then redeploy.

## After deploying
Everything works exactly the same as in the chat preview — Integrations
panel, one-tap buttons, AI-generated answers, file exports, Chat Assistant
with Skill Profile & Memory. Background archive/notify calls no longer need
a click; they'll fire automatically when `Auto-notify CEO on every
completed task` is on, since real hosting doesn't block outbound requests
the way the chat sandbox does. Settings, teams, policy, skill profile, and
agent memory are all saved permanently in the browser via localStorage.
