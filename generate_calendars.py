import requests
from bs4 import BeautifulSoup
from ics import Calendar, Event
from datetime import datetime, timedelta, timezone
import json
import os
import time
from msal import ConfidentialClientApplication

OUTPUT_DIR = "calendars"
os.makedirs(OUTPUT_DIR, exist_ok=True)

# ---------------------------------------------------------
#   MAIL-FUNKTION (Microsoft Graph API)
# ---------------------------------------------------------

def send_mail(subject, body):
    try:
        tenant_id = os.environ["SMTP_TENANT_ID"]
        client_id = os.environ["SMTP_CLIENT_ID"]
        client_secret = os.environ["SMTP_CLIENT_SECRET"]
        sender = "automation@sg-oftersheim.de"

        app = ConfidentialClientApplication(
            client_id,
            authority=f"https://login.microsoftonline.com/{tenant_id}",
            client_credential=client_secret
        )

        result = app.acquire_token_for_client(
            scopes=["https://graph.microsoft.com/.default"]
        )

        if "access_token" not in result:
            print("⚠ Mail: Kein Token erhalten")
            return False

        access_token = result["access_token"]

        mail = {
            "message": {
                "subject": subject,
                "body": {
                    "contentType": "Text",
                    "content": body
                },
                "toRecipients": [
                    {"emailAddress": {"address": "it@sg-oftersheim.de"}}
                ]
            }
        }

        response = requests.post(
            f"https://graph.microsoft.com/v1.0/users/{sender}/sendMail",
            headers={"Authorization": f"Bearer {access_token}"},
            json=mail
        )

        print("Mail Status:", response.status_code)
        return response.status_code == 202

    except Exception as e:
        print("⚠ Mail Fehler:", e)
        return False


# ---------------------------------------------------------
#   MAIL-STEUERUNG
# ---------------------------------------------------------

send_success = os.getenv("SEND_EMAIL_ON_SUCCESS", "false").lower() == "true"
send_error = os.getenv("SEND_EMAIL_ON_ERROR", "false").lower() == "true"

GITHUB_PAGES_BASE = "https://sgo-it.github.io/TeamCalendars/calendars/"


# ---------------------------------------------------------
# Geocoding: Adresse → GPS Koordinaten (OpenStreetMap)
# ---------------------------------------------------------

def geocode(address):
    if not address:
        return None, None

    url = "https://nominatim.openstreetmap.org/search"
    params = {"q": address, "format": "json", "limit": 1}

    try:
        r = requests.get(url, params=params, headers={"User-Agent": "SGO-CalendarBot"}).json()
        if r:
            return r[0]["lat"], r[0]["lon"]
    except Exception:
        pass

    return None, None


# ---------------------------------------------------------
# Spielort extrahieren
# ---------------------------------------------------------

def fetch_venue(match_url):
    if not match_url:
        return ""

    try:
        html = requests.get(match_url).text
        soup = BeautifulSoup(html, "html.parser")

        loc = soup.select_one("a.location")
        if loc:
            return loc.get_text(strip=True)

        fallback = [
            ".venue-name", ".match-location", ".location-name",
            "div.venue", "span.venue", ".match-info-location"
        ]

        for sel in fallback:
            el = soup.select_one(sel)
            if el:
                return el.get_text(strip=True)

    except Exception:
        pass

    return ""


# ---------------------------------------------------------
# Platzname + Adresse automatisch trennen
# ---------------------------------------------------------

def split_venue(venue):
    if not venue:
        return "", ""

    parts = [p.strip() for p in venue.split(",")]

    for i, p in enumerate(parts):
        if any(char.isdigit() for char in p):
            name = ", ".join(parts[:i])
            address = ", ".join(parts[i:])
            return name, address

    return venue, ""


# ---------------------------------------------------------
# Spiele laden
# ---------------------------------------------------------

def fetch_matches(team_url):
    if "#!" in team_url:
        team_url = team_url.split("#!")[0]

    html = requests.get(team_url).text
    soup = BeautifulSoup(html, "html.parser")

    matches = []
    current_date = None
    current_competition = None

    for row in soup.select("tr"):

        if "row-headline" in row.get("class", []):
            headline = row.get_text(" ", strip=True)
            parts = headline.split("|")

            raw = parts[0].strip().replace("Uhr", "").strip()

            if "," in raw:
                raw = raw.split(",", 1)[1].strip()

            if "-" in raw:
                dt = datetime.strptime(raw, "%d.%m.%Y - %H:%M")
            else:
                dt = datetime.strptime(raw, "%d.%m.%Y")

            dt = dt.replace(tzinfo=timezone(timedelta(hours=2)))

            current_date = dt
            current_competition = parts[1].strip() if len(parts) > 1 else ""
            continue

        if row.select_one(".column-club"):
            clubs = row.select(".column-club .club-name")
            if len(clubs) < 2:
                continue

            home = clubs[0].get_text(strip=True)
            away = clubs[1].get_text(strip=True)

            score_link = row.select_one(".column-score a")
            match_url = score_link["href"] if score_link else ""

            venue = fetch_venue(match_url)

            matches.append({
                "title": f"{home} - {away}",
                "start": current_date,
                "end": current_date + timedelta(minutes=90),
                "league": current_competition,
                "location": venue
            })

    return matches


# ---------------------------------------------------------
# ICS erzeugen
# ---------------------------------------------------------

def build_calendar(matches, team_short):
    cal = Calendar()

    for m in matches:
        e = Event()

        # SUMMARY mit Team-Kürzel
        e.name = f"{team_short}{m['title']}"

        e.begin = m["start"]
        e.end = m["end"]

        venue_name, venue_address = split_venue(m["location"])

        lat, lon = geocode(venue_address)
        time.sleep(1)

        e.location = venue_address

        desc = m["league"]

        if venue_name:
            desc += f" – {venue_name}"

        desc += f"\nAdresse: {venue_address}"

        if lat and lon:
            desc += f"\nGPS: {lat}, {lon}"
            desc += f"\nGoogle Maps: https://maps.google.com/?q={lat},{lon}"
            desc += f"\nApple Maps: https://maps.apple.com/?q={lat},{lon}"

        e.description = desc

        cal.events.add(e)

    return cal


# ---------------------------------------------------------
# Hauptprogramm
# ---------------------------------------------------------

try:
    teams = json.load(open("teams.json", "r"))

    for team in teams:
        matches = fetch_matches(team["url"])
        print("Matches gefunden:", len(matches))

        cal = build_calendar(matches, team.get("short", ""))

        filename = f"{OUTPUT_DIR}/{team['name']}.ics"
        with open(filename, "w", encoding="utf-8") as f:
            f.write(cal.serialize())

        print(f"Erzeugt: {filename} (Events: {len(matches)})")

        if send_success:
            send_mail(
                subject=f"Kalender aktualisiert: {team['name']}",
                body=(
                    f"Der Kalender für '{team['name']}' wurde erfolgreich erzeugt.\n"
                    f"Anzahl Spiele: {len(matches)}\n"
                    f"ICS-Datei: {GITHUB_PAGES_BASE}{team['name']}.ics\n"
                )
            )

    print("✔ Alle Kalender erfolgreich erzeugt")

except Exception as e:
    print("⚠ Fehler beim Kalender-Update:", e)

    if send_error:
        send_mail(
            subject="Fehler beim Kalender-Update",
            body=f"Beim Erzeugen der Kalender ist ein Fehler aufgetreten:\n\n{e}"
        )

    raise
