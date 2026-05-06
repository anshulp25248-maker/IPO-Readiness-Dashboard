# LinkedIn Macro Agent

A production-ready Python agent that:

- fetches global macro, finance, and geopolitical news
- scores the most important stories
- generates analyst-level LinkedIn posts with Gemini or OpenAI
- posts through the LinkedIn API when available
- falls back to Selenium browser automation when needed
- runs on a 3-day cadence with a major-event override
- stores memory so it avoids repeating the same topic
- logs every decision and posting attempt

## Simple Architecture

The system works in one pipeline:

1. `news_fetcher.py` pulls recent macro-relevant news from NewsData.io, NewsAPI, and/or GDELT.
2. `analyzer.py` scores every story for macro importance, India relevance, recency, and event urgency.
3. `content_generator.py` asks Gemini or OpenAI to think like a macro analyst and write the LinkedIn post.
4. `linkedin_poster.py` publishes the post using:
   - LinkedIn API first
   - Selenium fallback if API is unavailable or fails
5. `memory.py` stores the last posted topics and run history.
6. `scheduler.py` checks every hour and decides whether to post now or wait.

## Project Structure

```text
linkedin_macro_agent/
├── .env.example
├── .gitignore
├── README.md
├── requirements.txt
├── config.py
├── memory.py
├── news_fetcher.py
├── analyzer.py
├── content_generator.py
├── linkedin_poster.py
├── main.py
└── scheduler.py
```

## Step-by-Step Setup

### 1. Open the project folder

```powershell
cd C:\Users\anshu\Desktop\IPO-READINESS-SYSTEM\linkedin_macro_agent
```

### 2. Create a virtual environment

```powershell
python -m venv .venv
.venv\Scripts\activate
```

### 3. Install dependencies

```powershell
pip install -r requirements.txt
```

### 4. Create your `.env` file

```powershell
Copy-Item .env.example .env
```

Then open `.env` and fill in your real values.

At minimum, set:

- `AI_PROVIDER=gemini`
- `GEMINI_API_KEY`
- `NEWS_PROVIDER=newsdata`
- `NEWSDATA_API_KEY`
- `LINKEDIN_EMAIL` and `LINKEDIN_PASSWORD` if you want Selenium posting

For LinkedIn API posting, also set:

- `LINKEDIN_CLIENT_ID`
- `LINKEDIN_CLIENT_SECRET`
- `LINKEDIN_REDIRECT_URI`
- `LINKEDIN_ACCESS_TOKEN`
- `LINKEDIN_PERSON_ID` or `LINKEDIN_PERSON_URN`

## AI Provider Setup

This project now supports two content-generation backends:

- `gemini` through the official Google `google-genai` SDK
- `openai` through the OpenAI Responses API

### Gemini setup

Recommended `.env` values:

```env
AI_PROVIDER=gemini
GEMINI_API_KEY=your_gemini_api_key
GEMINI_MODEL=gemini-2.5-flash
```

### OpenAI setup

If you want to keep OpenAI as an option:

```env
AI_PROVIDER=openai
OPENAI_API_KEY=your_openai_api_key
OPENAI_MODEL=gpt-4.1
```

The active model is controlled by `AI_PROVIDER` plus the matching model variable in `.env`.

## News Provider Setup

### Option A: NewsData.io

If this is the news provider you already use, set:

```env
NEWS_PROVIDER=newsdata
NEWSDATA_API_KEY=your_newsdata_api_key
NEWSDATA_COUNTRY=
NEWSDATA_CATEGORIES=business,politics
NEWSDATA_QUERY=
NEWSDATA_QUERY_IN_TITLE=
NEWS_LANGUAGE=en
```

Notes:

- the code uses NewsData.io's `latest` endpoint
- it pulls business and politics categories by default
- if you build a custom filter in NewsData.io's query builder, you can paste it into `NEWSDATA_QUERY` or `NEWSDATA_QUERY_IN_TITLE`
- `FETCH_WINDOW_HOURS` is automatically capped at 48 hours for this provider
- if you are on the free plan, NewsData.io returns delayed news

### Option B: NewsAPI

1. Create a NewsAPI account.
2. Copy your API key into `NEWS_API_KEY`.
3. Keep `NEWS_PROVIDER=newsapi`.

### Option C: GDELT

If you do not want NewsAPI:

1. Set `NEWS_PROVIDER=gdelt`
2. Leave `NEWS_API_KEY` blank

The code can also fall back to GDELT automatically if NewsAPI is not configured.

## LinkedIn Posting Setup

You have two ways to publish.

### Option A: LinkedIn API

This is the cleaner and more stable option when your app has the right LinkedIn products/scopes.

You need a LinkedIn developer app with:

- `w_member_social` to create posts
- OpenID/profile access so you can identify the authenticated member

Typical flow:

1. Fill `LINKEDIN_CLIENT_ID`, `LINKEDIN_CLIENT_SECRET`, and `LINKEDIN_REDIRECT_URI` in `.env`
2. Generate the OAuth URL:

```powershell
python main.py --print-linkedin-auth-url
```

3. Open the printed URL in your browser and complete the authorization flow.
4. Copy the returned authorization `code`.
5. Exchange it for an access token:

```powershell
python main.py --exchange-linkedin-code "PASTE_CODE_HERE"
```

6. Put the access token into `LINKEDIN_ACCESS_TOKEN`.
7. Fetch your profile to discover the member id / URN:

```powershell
python main.py --fetch-linkedin-profile
```

8. Put either:
   - `LINKEDIN_PERSON_ID`
   - or `LINKEDIN_PERSON_URN`

9. Set `POSTING_MODE=api` or `POSTING_MODE=auto`

### Option B: Selenium Fallback

Use this if:

- LinkedIn API access is not approved for your app
- you want a quick local workaround

Set these values in `.env`:

- `LINKEDIN_EMAIL`
- `LINKEDIN_PASSWORD`
- optionally `LINKEDIN_HEADLESS=false`

Important:

- Selenium is more fragile than the API because LinkedIn can change its UI.
- CAPTCHAs and 2FA can interrupt the login flow.
- A visible browser session is usually more reliable than headless mode.

## How to Run the Project

### Manual dry run

This is the safest first test because it generates the post but does not publish it:

```powershell
python main.py --dry-run
```

### Generate without publishing

This saves the post locally and skips LinkedIn:

```powershell
python main.py --skip-posting
```

### Force a live run now

This ignores the 3-day delay and posts immediately:

```powershell
python main.py --force
```

### Choose the posting method explicitly

```powershell
python main.py --force --post-method api
python main.py --force --post-method selenium
```

## Scheduler Behavior

The agent checks every `CHECK_INTERVAL_MINUTES` minutes. By default that is every 60 minutes.

Why hourly checks instead of exactly every 3 days?

- the memory logic enforces the normal 3-day cadence
- the hourly check allows a fast override when a major event appears

Start the scheduler:

```powershell
python scheduler.py
```

The scheduler will post when:

- 3 or more days have passed since the last successful post
- or a new major event is detected, such as war, crisis, collapse, a surprise rate move, or an inflation shock

## Files Created at Runtime

The code automatically creates:

- `data/state.json`
- `data/run_history.jsonl`
- `data/generated_posts/`
- `logs/agent.log`

## Common Errors and Fixes

### `OPENAI_API_KEY` missing

Fix:

- add your OpenAI key to `.env`

What happens:

- the code will fall back to a template post
- the post will still be generated, but quality will be lower than the OpenAI path

### `GEMINI_API_KEY` missing

Fix:

- add your Gemini key to `.env`
- confirm `AI_PROVIDER=gemini` only when that key is present

What happens:

- the code tries the next available provider if configured
- otherwise it falls back to the template post

### `NEWSDATA_API_KEY` missing

Fix:

- add your NewsData.io key
- keep `NEWS_PROVIDER=newsdata`

### `NEWS_API_KEY` missing

Fix:

- add a NewsAPI key
- or switch to `NEWS_PROVIDER=gdelt`

### LinkedIn API returns permission errors

Fix:

- check that your LinkedIn app has `w_member_social`
- confirm the access token belongs to the same member you are trying to post for
- set `LINKEDIN_PERSON_ID` or `LINKEDIN_PERSON_URN` correctly

### Selenium fails after login

Fix:

- complete CAPTCHA or 2FA manually
- keep `LINKEDIN_HEADLESS=false`
- update Chrome to the latest version

### Repeated topics are being skipped

Fix:

- this is expected behavior from the memory system
- use `python main.py --force` if you intentionally want to publish again
- or delete `data/state.json` to reset memory during testing

## How to Scale This Later

### 1. Multiple posts per cycle

Add a loop in `main.py` so each selected article becomes its own post and store multiple topic signatures in memory.

### 2. Better content quality

Improve the prompt with:

- house style examples
- your own writing samples
- fixed positioning such as “global macro strategist for India-first professionals”

### 3. Better selection quality

Upgrade scoring with:

- source quality weights
- entity recognition
- market data overlays
- central bank calendar awareness

### 4. Personalization

You can make the content more “you” by adding:

- preferred tone
- signature phrases
- favored sectors
- India-first or global-first emphasis

### 5. Stronger production deployment

For a more serious deployment later, move this from a local machine to:

- a VPS
- GitHub Actions
- a cloud cron job
- a small database instead of JSON files

## Official Docs Used For This Build

- Gemini API Python quickstart: https://ai.google.dev/gemini-api/docs/quickstart
- Gemini structured output: https://ai.google.dev/gemini-api/docs/structured-output
- NewsData.io latest endpoint: https://newsdata.io/blog/newsdata-latest-news-api/
- OpenAI Responses API: https://platform.openai.com/docs/api-reference/responses/create
- LinkedIn Share on LinkedIn: https://learn.microsoft.com/en-us/linkedin/consumer/integrations/self-serve/share-on-linkedin
- LinkedIn Sign In with OpenID Connect: https://learn.microsoft.com/mt-mt/linkedin/consumer/integrations/self-serve/sign-in-with-linkedin-v2

## Recommended First Test

1. Fill `.env`
2. Run `python main.py --dry-run`
3. Read the generated post in the terminal and in `data/generated_posts/`
4. Test live posting with `--post-method api` or `--post-method selenium`
5. When happy, start `python scheduler.py`
