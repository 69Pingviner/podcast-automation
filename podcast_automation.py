import os
import json
import requests
import msal
import feedparser
from datetime import datetime, timezone
from youtube_transcript_api import YouTubeTranscriptApi
from msal import PublicClientApplication

# ========== CONFIGURATION ==========
ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY")
NOTIFICATION_EMAIL = "jacobmichaelsen@gmail.com"
ONENOTE_NOTEBOOK = "AI integration"
MIN_DURATION_SECONDS = 2700  # 45 minutes

YOUTUBE_CHANNELS = [
    {
        "name": "Peter Diamandis",
        "channel_id": "UCvxm0qTrGN_1LMYgUaftWyQ"
    }
]

RSS_FEEDS = []

# ========== MICROSOFT AUTH ==========
MS_APP_ID = "14d82eec-204b-4c2f-b7e8-296a70dab67e"
MS_SCOPES = ["Notes.ReadWrite", "Notes.Create"]
TOKEN_CACHE_FILE = "ms_token_cache.json"

def get_ms_token():
    cache = msal.SerializableTokenCache()
    if os.path.exists(TOKEN_CACHE_FILE):
        cache.deserialize(open(TOKEN_CACHE_FILE, "r").read())
    
    app = PublicClientApplication(MS_APP_ID, token_cache=cache)
    accounts = app.get_accounts()
    
    if accounts:
        result = app.acquire_token_silent(MS_SCOPES, account=accounts[0])
        if result and "access_token" in result:
            return result["access_token"]
    
    flow = app.initiate_device_flow(scopes=MS_SCOPES)
    print("\n" + flow["message"])
    print("\nWaiting for you to authenticate...")
    result = app.acquire_token_by_device_flow(flow)
    
    if "access_token" in result:
        with open(TOKEN_CACHE_FILE, "w") as f:
            f.write(cache.serialize())
        return result["access_token"]
    else:
        raise Exception("Authentication failed: " + str(result.get("error_description")))

# ========== YOUTUBE ==========
def get_youtube_videos(channel_id):
    url = f"https://www.youtube.com/feeds/videos.xml?channel_id={channel_id}"
    feed = feedparser.parse(url)
    videos = []
    for entry in feed.entries[:5]:
        video_id = entry.yt_videoid
        published = entry.published
        videos.append({
            "id": video_id,
            "title": entry.title,
            "published": published,
            "url": f"https://youtube.com/watch?v={video_id}"
        })
    return videos

def get_transcript(video_id):
    try:
        transcript_list = YouTubeTranscriptApi.get_transcript(video_id)
        full_text = " ".join([t["text"] for t in transcript_list])
        duration = transcript_list[-1]["start"] + transcript_list[-1]["duration"]
        return full_text, duration
    except Exception as e:
        print(f"Could not get transcript: {e}")
        return None, 0

# ========== CLAUDE ==========
def summarise_transcript(transcript, title):
    headers = {
        "x-api-key": ANTHROPIC_API_KEY,
        "anthropic-version": "2023-06-01",
        "content-type": "application/json"
    }
    
    prompt = f"""Here is a podcast transcript titled '{title}'. Structure your response exactly as follows:

EPISODE SUMMARY: 2-3 sentences overview

KEY INSIGHTS: 5-7 bullet points, each 2-3 sentences

BEST QUOTES: 3-5 verbatim quotes worth saving

SURPRISING IDEAS: Anything contrarian or unexpected

ACTION ITEMS: Things to explore or act on

VERDICT: One line - Essential / Useful / Optional

Transcript: {transcript[:50000]}"""

    body = {
        "model": "claude-sonnet-4-6",
        "max_tokens": 1500,
        "messages": [{"role": "user", "content": prompt}]
    }
    
    response = requests.post(
        "https://api.anthropic.com/v1/messages",
        headers=headers,
        json=body
    )
    
    result = response.json()
    return result["content"][0]["text"]

# ========== ONENOTE ==========
def get_or_create_section(token, notebook_name, section_name):
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }
    
    r = requests.get(
        "https://graph.microsoft.com/v1.0/me/onenote/notebooks",
        headers=headers
    )
    notebooks = r.json().get("value", [])
    notebook = next((n for n in notebooks if n["displayName"] == notebook_name), None)
    
    if not notebook:
        raise Exception(f"Notebook '{notebook_name}' not found")
    
    notebook_id = notebook["id"]
    
    r = requests.get(
        f"https://graph.microsoft.com/v1.0/me/onenote/notebooks/{notebook_id}/sections",
        headers=headers
    )
    sections = r.json().get("value", [])
    section = next((s for s in sections if s["displayName"] == section_name), None)
    
    if not section:
        r = requests.post(
            f"https://graph.microsoft.com/v1.0/me/onenote/notebooks/{notebook_id}/sections",
            headers=headers,
            json={"displayName": section_name}
        )
        section = r.json()
    
    return section["id"]

def create_onenote_page(token, section_id, title, summary, video_url, date):
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/xhtml+xml"
    }
    
    summary_html = summary.replace("\n", "<br/>")
    
    content = f"""<?xml version="1.0" encoding="utf-8" ?>
<html xmlns="http://www.w3.org/1999/xhtml" xmlns:o="urn:schemas-microsoft-com:office:office">
<head>
    <title>{title}</title>
    <meta name="created" content="{date}" />
</head>
<body>
    <h1>{title}</h1>
    <p><a href="{video_url}">Watch Episode</a></p>
    <p>{summary_html}</p>
</body>
</html>"""
    
    r = requests.post(
        f"https://graph.microsoft.com/v1.0/me/onenote/sections/{section_id}/pages",
        headers=headers,
        data=content.encode("utf-8")
    )
    
    if r.status_code == 201:
        print(f"Created OneNote page: {title}")
    else:
        print(f"Failed to create page: {r.text}")

# ========== TRACKING ==========
PROCESSED_FILE = "processed_episodes.json"

def load_processed():
    if os.path.exists(PROCESSED_FILE):
        return json.load(open(PROCESSED_FILE))
    return []

def save_processed(processed):
    json.dump(processed, open(PROCESSED_FILE, "w"))

# ========== MAIN ==========
def main():
    print(f"Starting podcast automation - {datetime.now()}")
    
    token = get_ms_token()
    processed = load_processed()
    new_episodes = 0
    
    for channel in YOUTUBE_CHANNELS:
        print(f"Checking {channel['name']}...")
        videos = get_youtube_videos(channel["channel_id"])
        section_id = get_or_create_section(token, ONENOTE_NOTEBOOK, channel["name"])
        
        for video in videos:
            if video["id"] in processed:
                continue
            
            print(f"Processing: {video['title']}")
            transcript, duration = get_transcript(video["id"])
            
            if not transcript:
                continue
                
            if duration < MIN_DURATION_SECONDS:
                print(f"Skipping - too short ({int(duration/60)} mins)")
                processed.append(video["id"])
                continue
            
            summary = summarise_transcript(transcript, video["title"])
            date = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
            
            create_onenote_page(
                token,
                section_id,
                video["title"],
                summary,
                video["url"],
                date
            )
            
            processed.append(video["id"])
            new_episodes += 1
            print(f"Done: {video['title']}")
    
    save_processed(processed)
    print(f"Finished. Processed {new_episodes} new episodes.")

if __name__ == "__main__":
    main()
