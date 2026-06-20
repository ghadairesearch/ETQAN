import argparse
import json
import os
import re
import ssl
import sys
import time
import urllib.parse
import urllib.request
from html.parser import HTMLParser


UNIVERSITIES = {
    "Umm Al-Qura University": "https://uqu.edu.sa",
    "Islamic University of Madinah": "https://iu.edu.sa",
    "Imam Mohammad Ibn Saud Islamic University": "https://imamu.edu.sa",
    "King Saud University": "https://ksu.edu.sa",
    "King Abdulaziz University": "https://www.kau.edu.sa",
    "King Faisal University": "https://www.kfu.edu.sa",
    "King Khalid University": "https://www.kku.edu.sa",
    "Qassim University": "https://www.qu.edu.sa",
    "Taibah University": "https://www.taibahu.edu.sa",
    "Taif University": "https://edugate.tu.edu.sa",
    "University of Hail": "https://www.uoh.edu.sa",
    "Jazan University": "https://www.jazanu.edu.sa",
    "Al Jouf University": "https://www.ju.edu.sa",
    "Al Baha University": "https://bu.edu.sa",
    "Tabuk University": "https://www.ut.edu.sa",
    "Najran University": "https://www.nu.edu.sa",
    "Northern Border University": "https://www.nbu.edu.sa",
    "Princess Nourah bint Abdulrahman University": "https://www.pnu.edu.sa",
    "King Saud bin Abdulaziz University for Health Sciences": "https://www.ksau-hs.edu.sa",
    "Imam Abdulrahman Bin Faisal University": "https://www.iau.edu.sa",
    "Prince Sattam bin Abdulaziz University": "https://www.psau.edu.sa",
    "Shaqra University": "https://www.su.edu.sa",
    "Majmaah University": "https://www.mu.edu.sa",
    "Saudi Electronic University": "https://seu.edu.sa",
    "University of Jeddah": "https://www.uj.edu.sa",
    "University of Bisha": "https://www.ub.edu.sa",
    "King Abdullah University of Science and Technology": "https://www.kaust.edu.sa",
    "King Fahd University of Petroleum & Minerals": "https://www.kfupm.edu.sa",
    "Prince Sultan University": "https://www.psu.edu.sa",
    "Effat University": "https://www.effatuniversity.edu.sa",
    "Dar Al-Hekma University": "https://www.dah.edu.sa",
    "Alfaisal University": "https://www.alfaisal.edu",
    "Arab Open University": "https://www.arabou.edu.sa",
}

COLOR_FALLBACKS = {
    "Umm Al-Qura University": "#1A656B",
    "Islamic University of Madinah": "#0F6B4F",
    "Imam Mohammad Ibn Saud Islamic University": "#1F4E79",
    "King Saud University": "#00558C",
    "King Abdulaziz University": "#006B54",
    "King Faisal University": "#0B6F4A",
    "King Khalid University": "#005B7F",
    "Qassim University": "#006C67",
    "Taibah University": "#006A71",
    "Taif University": "#174E7C",
    "University of Hail": "#005A70",
    "Jazan University": "#006B54",
    "Al Jouf University": "#214E78",
    "Al Baha University": "#006B54",
    "Tabuk University": "#005C7A",
    "Najran University": "#007A5E",
    "Northern Border University": "#234E70",
    "Princess Nourah bint Abdulrahman University": "#5B2C83",
    "King Saud bin Abdulaziz University for Health Sciences": "#007A53",
    "Imam Abdulrahman Bin Faisal University": "#008C95",
    "Prince Sattam bin Abdulaziz University": "#006B54",
    "Shaqra University": "#006B54",
    "Majmaah University": "#006B54",
    "Saudi Electronic University": "#4B287F",
    "University of Jeddah": "#005F86",
    "University of Bisha": "#006B54",
    "King Abdullah University of Science and Technology": "#00A5B5",
    "King Fahd University of Petroleum & Minerals": "#006747",
    "Prince Sultan University": "#1E3A8A",
    "Effat University": "#7A1E5C",
    "Dar Al-Hekma University": "#7A1E5C",
    "Alfaisal University": "#005A8B",
    "Arab Open University": "#1E4D8C",
}

USER_AGENT = "ETQAN-University-Identity-Collector/1.0"
IMAGE_EXTS = (".png", ".jpg", ".jpeg", ".webp", ".svg")
REPORT_FRIENDLY_EXTS = (".png", ".jpg", ".jpeg")

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass


class HomepageParser(HTMLParser):
    def __init__(self):
        super().__init__()
        self.images = []
        self.icons = []
        self.stylesheets = []
        self.inline_styles = []

    def handle_starttag(self, tag, attrs):
        attrs = {k.lower(): (v or "") for k, v in attrs}
        if tag.lower() == "img":
            src = attrs.get("src") or attrs.get("data-src") or attrs.get("data-lazy-src")
            if src:
                self.images.append({"src": src, "attrs": attrs})
        if tag.lower() == "link":
            rel = attrs.get("rel", "").lower()
            href = attrs.get("href")
            if not href:
                return
            if "stylesheet" in rel:
                self.stylesheets.append(href)
            if "icon" in rel or "apple-touch-icon" in rel:
                self.icons.append({"src": href, "attrs": attrs})
        style = attrs.get("style")
        if style:
            self.inline_styles.append(style)


def slugify(value):
    value = re.sub(r"[^a-zA-Z0-9]+", "-", value.lower()).strip("-")
    return value or "university"


def fetch_bytes(url, timeout=20):
    request = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
    context = ssl.create_default_context()
    try:
        with urllib.request.urlopen(request, timeout=timeout, context=context) as response:
            final_url = response.geturl()
            content_type = response.headers.get("content-type", "")
            return response.read(), final_url, content_type
    except Exception as exc:
        if "CERTIFICATE_VERIFY_FAILED" not in str(exc):
            raise
        unverified_context = ssl._create_unverified_context()
        with urllib.request.urlopen(request, timeout=timeout, context=unverified_context) as response:
            final_url = response.geturl()
            content_type = response.headers.get("content-type", "")
            return response.read(), final_url, content_type


def fetch_text(url, timeout=20):
    data, final_url, content_type = fetch_bytes(url, timeout)
    encoding = "utf-8"
    match = re.search(r"charset=([\w-]+)", content_type, flags=re.I)
    if match:
        encoding = match.group(1)
    return data.decode(encoding, errors="replace"), final_url


def absolutize(base_url, value):
    if not value or value.startswith("data:"):
        return ""
    absolute = urllib.parse.urljoin(base_url, value)
    parts = urllib.parse.urlsplit(absolute)
    return urllib.parse.urlunsplit((
        parts.scheme,
        parts.netloc,
        urllib.parse.quote(parts.path, safe="/%:@"),
        urllib.parse.quote(parts.query, safe="=&%:@/?"),
        urllib.parse.quote(parts.fragment, safe="=&%:@/?"),
    ))


def url_extension(url):
    path = urllib.parse.urlparse(url).path.lower()
    _, ext = os.path.splitext(path)
    return ext if ext in IMAGE_EXTS else ""


def score_logo_candidate(candidate):
    src = candidate.get("src", "")
    attrs = candidate.get("attrs", {})
    haystack = " ".join([src, attrs.get("alt", ""), attrs.get("title", ""), attrs.get("class", ""), attrs.get("id", "")]).lower()
    score = 0
    for word, points in {
        "logo": 80,
        "brand": 35,
        "identity": 25,
        "header": 18,
        "navbar": 15,
        "site": 8,
        "university": 8,
        "favicon": 5,
    }.items():
        if word in haystack:
            score += points
    ext = url_extension(src)
    if ext in REPORT_FRIENDLY_EXTS:
        score += 16
    elif ext == ".svg":
        score += 8
    elif ext == ".webp":
        score += 4
    if "sprite" in haystack or "placeholder" in haystack or "loader" in haystack:
        score -= 50
    return score


def choose_logo(parser, final_url):
    candidates = []
    for image in parser.images:
        src = absolutize(final_url, image["src"])
        if url_extension(src):
            candidates.append({"src": src, "attrs": image["attrs"]})
    for icon in parser.icons:
        src = absolutize(final_url, icon["src"])
        if url_extension(src):
            candidates.append({"src": src, "attrs": icon["attrs"]})

    report_friendly = [candidate for candidate in candidates if url_extension(candidate["src"]) in REPORT_FRIENDLY_EXTS]
    pool = report_friendly or candidates
    if not pool:
        return ""
    return max(pool, key=score_logo_candidate)["src"]


def hex_to_rgb(value):
    value = value.lstrip("#")
    if len(value) == 3:
        value = "".join(ch * 2 for ch in value)
    if len(value) != 6:
        return None
    try:
        return tuple(int(value[index:index + 2], 16) for index in (0, 2, 4))
    except ValueError:
        return None


def color_score(value):
    rgb = hex_to_rgb(value)
    if not rgb:
        return -1
    red, green, blue = rgb
    brightness = (red + green + blue) / 3
    spread = max(rgb) - min(rgb)
    if brightness < 28 or brightness > 235 or spread < 20:
        return -1
    blue_green_bias = max(green, blue) - red * 0.25
    return spread + blue_green_bias - abs(brightness - 95) * 0.25


def extract_colors(text):
    colors = []
    for match in re.finditer(r"#[0-9A-Fa-f]{3}(?:[0-9A-Fa-f]{3})?\b", text or ""):
        raw = match.group(0).upper()
        rgb = hex_to_rgb(raw)
        if not rgb:
            continue
        if len(raw) == 4:
            raw = "#" + "".join(ch * 2 for ch in raw[1:])
        colors.append(raw)
    return colors


def choose_color(home_html, stylesheet_texts, fallback):
    colors = extract_colors(home_html)
    for css_text in stylesheet_texts:
        colors.extend(extract_colors(css_text))
    if not colors:
        return fallback
    counts = {}
    for color in colors:
        score = color_score(color)
        if score < 0:
            continue
        counts[color] = counts.get(color, 0) + 1 + (score / 100.0)
    if not counts:
        return fallback
    return max(counts.items(), key=lambda item: item[1])[0]


def download_logo(logo_url, output_dir, slug):
    if not logo_url:
        return ""
    ext = url_extension(logo_url)
    if ext == ".webp":
        return ""
    if not ext:
        ext = ".png"
    filename = f"{slug}{ext}"
    path = os.path.join(output_dir, filename)
    try:
        data, _, _ = fetch_bytes(logo_url, timeout=25)
    except Exception:
        return ""
    if len(data) < 100:
        return ""
    with open(path, "wb") as f:
        f.write(data)
    return filename


def collect_one(name, website, output_dir, css_limit=3):
    html, final_url = fetch_text(website)
    parser = HomepageParser()
    parser.feed(html)
    logo_url = choose_logo(parser, final_url)
    stylesheet_texts = []
    for stylesheet in parser.stylesheets[:css_limit]:
        css_url = absolutize(final_url, stylesheet)
        if not css_url:
            continue
        try:
            css_text, _ = fetch_text(css_url, timeout=12)
            stylesheet_texts.append(css_text)
        except Exception:
            pass
    primary_color = choose_color(html + "\n".join(parser.inline_styles), stylesheet_texts, COLOR_FALLBACKS.get(name, "#26365f"))
    logo_filename = ""
    if logo_url:
        logo_filename = download_logo(logo_url, output_dir, slugify(name))
    return {
        "website": website,
        "resolved_website": final_url,
        "logo_url": logo_url,
        "logo_filename": logo_filename,
        "primary_color": primary_color,
    }


def main():
    parser = argparse.ArgumentParser(description="Collect university website identity defaults for ETQAN.")
    parser.add_argument("--output", default="university_identity.json")
    parser.add_argument("--logo-dir", default=os.path.join("public", "university_logos"))
    parser.add_argument("--sleep", type=float, default=0.5)
    parser.add_argument("--only", default="", help="Collect one university by exact name.")
    args = parser.parse_args()

    os.makedirs(args.logo_dir, exist_ok=True)
    existing = {}
    if os.path.exists(args.output):
        with open(args.output, "r", encoding="utf-8") as f:
            existing = json.load(f)

    result = dict(existing)
    items = UNIVERSITIES.items()
    if args.only:
        items = [(args.only, UNIVERSITIES[args.only])]

    for name, website in items:
        print(f"Collecting {name} ...")
        try:
            result[name] = collect_one(name, website, args.logo_dir)
            print(f"  logo={result[name].get('logo_filename') or '-'} color={result[name].get('primary_color')}")
        except Exception as exc:
            print(f"  failed: {exc}")
            result[name] = {
                "website": website,
                "resolved_website": "",
                "logo_url": "",
                "logo_filename": "",
                "primary_color": COLOR_FALLBACKS.get(name, "#26365f"),
                "error": str(exc),
            }
        with open(args.output, "w", encoding="utf-8") as f:
            json.dump(result, f, ensure_ascii=False, indent=2, sort_keys=True)
        time.sleep(args.sleep)


if __name__ == "__main__":
    main()
