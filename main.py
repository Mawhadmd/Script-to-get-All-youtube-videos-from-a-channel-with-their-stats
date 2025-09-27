from googleapiclient.discovery import build
import openpyxl

# Replace with your API key
API_KEY = "YOUR_API_KEY"
CHANNEL_ID = "CHANNEL_ID_HERE"

youtube = build("youtube", "v3", developerKey=API_KEY)

def get_channel_uploads_id(channel_id):
    req = youtube.channels().list(
        part="contentDetails",
        id=channel_id
    )
    res = req.execute()
    return res["items"][0]["contentDetails"]["relatedPlaylists"]["uploads"]

def get_videos(playlist_id):
    videos = []
    next_page_token = None

    while True:
        req = youtube.playlistItems().list(
            part="snippet",
            playlistId=playlist_id,
            maxResults=50,
            pageToken=next_page_token
        )
        res = req.execute()

        for item in res["items"]:
            videos.append({
                "video_id": item["snippet"]["resourceId"]["videoId"],
                "title": item["snippet"]["title"],
                "published_at": item["snippet"]["publishedAt"]
            })

        next_page_token = res.get("nextPageToken")
        if not next_page_token:
            break
    return videos

def get_video_stats(video_ids):
    req = youtube.videos().list(
        part="statistics",
        id=",".join(video_ids)
    )
    res = req.execute()
    stats = {}
    for item in res["items"]:
        stats[item["id"]] = item["statistics"]
    return stats

def save_to_excel(videos, filename="youtube_videos.xlsx"):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Videos"

    # Header row
    ws.append(["Title", "Video ID", "Upload Date", "Views", "Likes", "Comments"])

    for v in videos:
        stats = v.get("stats", {})
        ws.append([
            v["title"],
            v["video_id"],
            v["published_at"],
            stats.get("viewCount", "0"),
            stats.get("likeCount", "0"),
            stats.get("commentCount", "0")
        ])

    wb.save(filename)
    print(f"Saved to {filename}")

if __name__ == "__main__":
    uploads_id = get_channel_uploads_id(CHANNEL_ID)
    videos = get_videos(uploads_id)

    # batch stats in chunks of 50
    for i in range(0, len(videos), 50):
        ids = [v["video_id"] for v in videos[i:i+50]]
        stats = get_video_stats(ids)
        for v in videos[i:i+50]:
            v["stats"] = stats.get(v["video_id"], {})

    save_to_excel(videos)
