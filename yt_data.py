
import os
import re
import statistics
from datetime import datetime, timezone, timedelta
from urllib.parse import urlparse

import pandas as pd
from dotenv import load_dotenv
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError


# Country ranking: lower number means higher priority.
COUNTRY_ORDER = {
    "US": 1,
    "GB": 2,
    "CA": 3,
    "AU": 4,
    "DE": 5,
    "IN": 6,
}

COUNTRY_NAMES = {
    "US": "USA",
    "GB": "UK",
    "CA": "Canada",
    "AU": "Australia",
    "DE": "Germany",
    "IN": "India",
}

SOCIAL_DOMAINS = {
    "instagram": ("instagram.com",),
    "facebook": ("facebook.com", "fb.com"),
    "x": ("x.com", "twitter.com"),
}

EMAIL_PATTERN = re.compile(
    r"[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}",
    re.IGNORECASE,
)

URL_PATTERN = re.compile(
    r"""(?:https?://|www\.)[^\s<>"']+""",
    re.IGNORECASE,
)


def _extract_contacts(texts):
    """Extract only contact details publicly present in descriptions."""
    result = {
        "Email": None,
        "Instagram": None,
        "Facebook": None,
        "X": None,
        "Website": None,
    }

    for text in texts:
        text = text or ""

        if result["Email"] is None:
            match = EMAIL_PATTERN.search(text)
            if match:
                result["Email"] = match.group(0)

        for raw_url in URL_PATTERN.findall(text):
            url = raw_url.rstrip(".,;:!?)\u005d")
            if url.lower().startswith("www."):
                url = "https://" + url

            try:
                parsed = urlparse(url)
                host = (parsed.hostname or "").lower()
            except ValueError:
                continue

            if not host:
                continue

            if host.startswith("www."):
                host = host[4:]

            full_url = url

            # Do not mistake YouTube links for an external website.
            if host.endswith("youtube.com") or host == "youtu.be":
                continue

            matched_social = False

            for field, domains in SOCIAL_DOMAINS.items():
                if any(
                    host == domain or host.endswith("." + domain)
                    for domain in domains
                ):
                    if result[field] is None:
                        result[field] = full_url
                    matched_social = True
                    break

            if matched_social:
                continue

            # Other public URLs count as website/contact links.
            if result["Website"] is None:
                result["Website"] = full_url

    return result


def _country_score(code):
    # Maximum 20 points. Unknown country is not assumed to be USA.
    scores = {
        "US": 20,
        "GB": 18,
        "CA": 16,
        "AU": 15,
        "DE": 13,
        "IN": 10,
    }
    return scores.get(code, 5)


def _subscriber_score(subscribers):
    if subscribers is None:
        return 5

    if 10_000 <= subscribers <= 100_000:
        return 15
    if 100_001 <= subscribers <= 500_000:
        return 13
    if 1_000 <= subscribers < 10_000:
        return 12
    if 500_001 <= subscribers <= 1_000_000:
        return 10
    if subscribers > 1_000_000:
        return 7
    return 4


def _activity_score(count_60):
    if count_60 >= 12:
        return 25
    if count_60 >= 6:
        return 22
    if count_60 >= 4:
        return 18
    if count_60 >= 2:
        return 13
    if count_60 == 1:
        return 6
    return 0


def _views_score(median_views, subscribers):
    if median_views is None:
        return 4

    score = 5

    if median_views >= 1_000:
        score += 3
    if median_views >= 5_000:
        score += 3
    if median_views >= 20_000:
        score += 2

    # Engagement proxy, not an actual engagement rate.
    if subscribers and subscribers > 0:
        ratio = median_views / subscribers
        if ratio >= 0.10:
            score += 2
        elif ratio >= 0.03:
            score += 1

    return min(score, 15)


def _contact_score(contacts):
    # Maximum 15 points. Contact availability increases approachability.
    score = 0

    if contacts["Email"]:
        score += 8
    if contacts["Website"]:
        score += 3
    if any(
        contacts[field]
        for field in ("Instagram", "Facebook", "X")
    ):
        score += 4

    return score


def _date_or_none(value):
    if not value:
        return None

    try:
        return datetime.fromisoformat(
            value.replace("Z", "+00:00")
        )
    except (ValueError, TypeError):
        return None


def find_youtube_editing_leads(
    keywords,
    api_key=None,
    max_channels_per_keyword=10,
    recent_videos=15,
    output_dir=".",
):
    """
    Discover YouTube channels and rank potential editing clients.

    Args:
        keywords: List of niches/search phrases.
        api_key: Optional YouTube Data API key.
        max_channels_per_keyword: 1-50 search results per keyword.
        recent_videos: 1-50 recent uploads to analyze per channel.
        output_dir: Folder for the dated Excel workbook.

    Returns:
        pandas.DataFrame sorted by lead priority.
    """
    load_dotenv()

    api_key = api_key or os.getenv("YOUTUBE_API_KEY")

    if not api_key:
        raise ValueError(
            "YOUTUBE_API_KEY missing. Add it to your .env file."
        )

    if isinstance(keywords, str):
        keywords = [keywords]

    keywords = list(dict.fromkeys(
        str(k).strip() for k in keywords if str(k).strip()
    ))

    if not keywords:
        raise ValueError("Provide at least one keyword.")

    if not 1 <= max_channels_per_keyword <= 50:
        raise ValueError("max_channels_per_keyword must be 1-50.")

    if not 1 <= recent_videos <= 50:
        raise ValueError("recent_videos must be 1-50.")

    youtube = build(
        "youtube",
        "v3",
        developerKey=api_key,
        cache_discovery=False,
    )

    # Step 1: Discover channels from keyword searches.
    discovered = {}

    for keyword in keywords:
        response = youtube.search().list(
            part="snippet",
            q=keyword,
            type="channel",
            maxResults=max_channels_per_keyword,
        ).execute()

        for item in response.get("items", []):
            channel_id = item.get("id", {}).get("channelId")

            if channel_id:
                discovered.setdefault(channel_id, set()).add(keyword)

    if not discovered:
        print("No channels found for the supplied keywords.")
        return pd.DataFrame()

    # Step 2: Fetch channel statistics in batches.
    channel_ids = list(discovered.keys())
    channel_data = {}

    for start in range(0, len(channel_ids), 50):
        batch = channel_ids[start:start + 50]

        response = youtube.channels().list(
            part="snippet,statistics,contentDetails",
            id=",".join(batch),
            maxResults=50,
        ).execute()

        for channel in response.get("items", []):
            channel_data[channel["id"]] = channel

    now = datetime.now(timezone.utc)
    cutoff_30 = now - timedelta(days=30)
    cutoff_60 = now - timedelta(days=60)
    rows = []

    # Step 3: Inspect recent uploads and calculate scores.
    for index, (channel_id, channel) in enumerate(
        channel_data.items(), start=1
    ):
        print(f"[{index}/{len(channel_data)}] "
              f"Analyzing {channel_id}")

        snippet = channel.get("snippet", {})
        stats = channel.get("statistics", {})
        content = channel.get("contentDetails", {})

        country_code = snippet.get("country")
        country = COUNTRY_NAMES.get(
            country_code, country_code or "Unknown"
        )

        hidden_subscribers = stats.get(
            "hiddenSubscriberCount", False
        )

        subscribers = (
            int(stats["subscriberCount"])
            if stats.get("subscriberCount") is not None
            and not hidden_subscribers
            else None
        )

        total_videos = int(stats.get("videoCount", 0))
        total_views = int(stats.get("viewCount", 0))

        uploads_playlist = content.get(
            "relatedPlaylists", {}
        ).get("uploads")

        video_items = []

        if uploads_playlist:
            playlist_response = youtube.playlistItems().list(
                part="snippet,contentDetails",
                playlistId=uploads_playlist,
                maxResults=recent_videos,
            ).execute()

            video_items = playlist_response.get("items", [])

        video_ids = [
            item.get("contentDetails", {}).get("videoId")
            for item in video_items
            if item.get("contentDetails", {}).get("videoId")
        ]

        videos = []

        for start in range(0, len(video_ids), 50):
            batch = video_ids[start:start + 50]

            if not batch:
                continue

            response = youtube.videos().list(
                part="snippet,statistics",
                id=",".join(batch),
                maxResults=50,
            ).execute()

            videos.extend(response.get("items", []))

        published_dates = []
        recent_views = []
        video_descriptions = []

        for video in videos:
            video_snippet = video.get("snippet", {})
            video_stats = video.get("statistics", {})

            published = _date_or_none(
                video_snippet.get("publishedAt")
            )

            if published:
                published_dates.append(published)

                if published >= cutoff_30:
                    pass

            if video_stats.get("viewCount") is not None:
                recent_views.append(
                    int(video_stats["viewCount"])
                )

            video_descriptions.append(
                video_snippet.get("description", "")
            )

        uploads_30 = sum(
            1 for d in published_dates if d >= cutoff_30
        )
        uploads_60 = sum(
            1 for d in published_dates if d >= cutoff_60
        )

        latest_upload = (
            max(published_dates) if published_dates else None
        )

        days_since_upload = (
            max(0, (now - latest_upload).days)
            if latest_upload else None
        )

        median_views = (
            int(statistics.median(recent_views))
            if recent_views else None
        )

        # Channel description + recent video descriptions only.
        contacts = _extract_contacts(
            [snippet.get("description", "")]
            + video_descriptions
        )

        country_points = _country_score(country_code)
        subscriber_points = _subscriber_score(subscribers)
        activity_points = _activity_score(uploads_60)
        views_points = _views_score(median_views, subscribers)
        contact_points = _contact_score(contacts)

        # Recency is worth 10 points.
        if days_since_upload is None:
            recency_points = 0
        elif days_since_upload <= 7:
            recency_points = 10
        elif days_since_upload <= 14:
            recency_points = 8
        elif days_since_upload <= 30:
            recency_points = 5
        elif days_since_upload <= 60:
            recency_points = 2
        else:
            recency_points = 0

        score = (
            country_points
            + subscriber_points
            + activity_points
            + views_points
            + contact_points
            + recency_points
        )

        # Business potential and editing quality cannot be
        # reliably inferred from subscriber counts alone.
        if score >= 75:
            priority = "A - Contact first"
        elif score >= 60:
            priority = "B - Good prospect"
        elif score >= 45:
            priority = "C - Review manually"
        else:
            priority = "D - Low priority"

        country_rank = COUNTRY_ORDER.get(country_code, 99)

        channel_url = (
            f"https://www.youtube.com/channel/{channel_id}"
        )

        rows.append({
            "Channel ID": channel_id,
            "Channel Name": snippet.get("title"),
            "Channel URL": channel_url,
            "Country": country,
            "Country Priority": country_rank,
            "Search Keywords": ", ".join(
                sorted(discovered[channel_id])
            ),
            "Description": snippet.get("description"),
            "Subscribers": subscribers,
            "Total Videos": total_videos,
            "Total Views": total_views,
            "Uploads Last 30 Days": uploads_30,
            "Uploads Last 60 Days": uploads_60,
            "Last Upload": (
                latest_upload.isoformat()
                if latest_upload else None
            ),
            "Days Since Upload": days_since_upload,
            "Median Recent Views": median_views,
            "Lead Score": score,
            "Priority": priority,
            "Country Score": country_points,
            "Subscriber Score": subscriber_points,
            "Activity Score": activity_points,
            "Views Score": views_points,
            "Contact Score": contact_points,
            "Recency Score": recency_points,
            **contacts,
        })

    df = pd.DataFrame(rows)

    if df.empty:
        return df

    # Rank by score first, then preferred country, then views.
    df = df.sort_values(
        by=[
            "Lead Score",
            "Country Priority",
            "Median Recent Views",
        ],
        ascending=[False, True, False],
        na_position="last",
    ).reset_index(drop=True)

    df.insert(0, "Rank", range(1, len(df) + 1))

    # Excel should contain actual None values for missing data.
    df = df.astype(object).where(pd.notna(df), None)

    os.makedirs(output_dir, exist_ok=True)

    save_date = datetime.now().strftime("%Y-%m-%d")
    filename = os.path.join(
        output_dir,
        f"youtube_leads_{save_date}.xlsx",
    )

    with pd.ExcelWriter(
        filename, engine="openpyxl"
    ) as writer:
        df.to_excel(
            writer,
            sheet_name="Ranked Leads",
            index=False,
        )

        notes = pd.DataFrame({
            "Field": [
                "Lead Score",
                "Country Priority",
                "Uploads Last 60 Days",
                "Median Recent Views",
                "Contacts",
                "Limitations",
            ],
            "Explanation": [
                "Weighted heuristic; maximum 100 points.",
                "USA, UK, Canada, Australia, Germany, India, then unknown/other.",
                "Count among the most recent videos retrieved, not necessarily every upload.",
                "Median views among the recent videos whose view counts are available.",
                "Extracted from public channel and recent video descriptions only.",
                "API search results are not exhaustive. Country may be missing. "
                "No guarantee of editing need or willingness to pay.",
            ],
        })

        notes.to_excel(
            writer,
            sheet_name="Scoring Guide",
            index=False,
        )

        ws = writer.sheets["Ranked Leads"]
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

        for column_cells in ws.columns:
            letter = column_cells[0].column_letter
            max_length = max(
                len(str(cell.value or ""))
                for cell in column_cells[:100]
            )
            ws.column_dimensions[letter].width = min(
                max(max_length + 2, 12), 45
            )

    print(f"\nFound {len(df)} unique channels.")
    print(f"Excel saved: {os.path.abspath(filename)}")

    return df


if __name__ == "__main__":
    search_keywords = [
        "business podcast",
        "entrepreneurship YouTuber",
        "AI tools tutorials",
        "personal branding coach",
    ]

    try:
        leads = find_youtube_editing_leads(
            keywords=search_keywords,
            max_channels_per_keyword=10,
            recent_videos=15,
            output_dir="output",
        )

        if not leads.empty:
            print(
                leads[
                    [
                        "Rank",
                        "Channel Name",
                        "Country",
                        "Subscribers",
                        "Uploads Last 60 Days",
                        "Lead Score",
                        "Priority",
                    ]
                ].head(15).to_string(index=False)
            )

    except HttpError as error:
        print("YouTube API error:", error)
        print("Check your API key, enabled API and quota.")
