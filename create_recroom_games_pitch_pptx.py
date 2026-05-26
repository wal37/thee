from __future__ import annotations

from datetime import datetime, timezone
from pathlib import Path
from xml.sax.saxutils import escape
import zipfile


BASE = Path("/Users/wale/Desktop/game/gweb")
OUT = BASE / "RECROOM_GAMES_PITCH_DECK.pptx"
TEMPLATE = Path("/Users/wale/Desktop/scifi/sweb/LOOPNODE_PITCH_DECK.pptx")

EMU = 914400


def e(value: str) -> str:
    return escape(value)


def pt(value: float) -> int:
    return int(value * 100)


def inch(value: float) -> int:
    return int(value * EMU)


class SlideBuilder:
    def __init__(self) -> None:
        self._shape_id = 1
        self.parts: list[str] = []
        self.images: list[Path] = []

    def _next_id(self) -> int:
        self._shape_id += 1
        return self._shape_id

    def rect(
        self,
        x: float,
        y: float,
        w: float,
        h: float,
        fill: str,
        *,
        line: str | None = None,
        rounded: bool = False,
        alpha: int | None = None,
    ) -> None:
        sid = self._next_id()
        line_xml = (
            f"<a:ln w='12700'><a:solidFill><a:srgbClr val='{line}'/></a:solidFill></a:ln>"
            if line
            else "<a:ln><a:noFill/></a:ln>"
        )
        fill_inner = (
            f"<a:srgbClr val='{fill}'><a:alpha val='{alpha}'/></a:srgbClr>"
            if alpha is not None
            else f"<a:srgbClr val='{fill}'/>"
        )
        geom = "roundRect" if rounded else "rect"
        self.parts.append(
            f"""
<p:sp>
  <p:nvSpPr>
    <p:cNvPr id="{sid}" name="Shape {sid}"/>
    <p:cNvSpPr/>
    <p:nvPr/>
  </p:nvSpPr>
  <p:spPr>
    <a:xfrm>
      <a:off x="{inch(x)}" y="{inch(y)}"/>
      <a:ext cx="{inch(w)}" cy="{inch(h)}"/>
    </a:xfrm>
    <a:prstGeom prst="{geom}"><a:avLst/></a:prstGeom>
    <a:solidFill>{fill_inner}</a:solidFill>
    {line_xml}
  </p:spPr>
  <p:txBody>
    <a:bodyPr/>
    <a:lstStyle/>
    <a:p/>
  </p:txBody>
</p:sp>"""
        )

    def text(
        self,
        x: float,
        y: float,
        w: float,
        h: float,
        text: str,
        *,
        size: float = 20,
        color: str = "111111",
        bold: bool = False,
        font: str = "Aptos",
        align: str = "l",
        valign: str = "t",
        all_caps: bool = False,
    ) -> None:
        sid = self._next_id()
        paragraphs = []
        for line in text.split("\n"):
            cap = " cap='all'" if all_caps else ""
            paragraphs.append(
                f"""
    <a:p>
      <a:pPr algn="{align}"/>
      <a:r>
        <a:rPr lang="en-US" sz="{pt(size)}" b="{'1' if bold else '0'}"{cap}>
          <a:solidFill><a:srgbClr val="{color}"/></a:solidFill>
          <a:latin typeface="{font}"/>
        </a:rPr>
        <a:t>{e(line)}</a:t>
      </a:r>
      <a:endParaRPr lang="en-US" sz="{pt(size)}">
        <a:solidFill><a:srgbClr val="{color}"/></a:solidFill>
        <a:latin typeface="{font}"/>
      </a:endParaRPr>
    </a:p>"""
            )
        self.parts.append(
            f"""
<p:sp>
  <p:nvSpPr>
    <p:cNvPr id="{sid}" name="Text {sid}"/>
    <p:cNvSpPr txBox="1"/>
    <p:nvPr/>
  </p:nvSpPr>
  <p:spPr>
    <a:xfrm>
      <a:off x="{inch(x)}" y="{inch(y)}"/>
      <a:ext cx="{inch(w)}" cy="{inch(h)}"/>
    </a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
    <a:noFill/>
    <a:ln><a:noFill/></a:ln>
  </p:spPr>
  <p:txBody>
    <a:bodyPr anchor="{valign}" wrap="square"/>
    <a:lstStyle/>
    {''.join(paragraphs)}
  </p:txBody>
</p:sp>"""
        )

    def image(self, x: float, y: float, w: float, h: float, path: str | Path, *, rounded: bool = False) -> None:
        sid = self._next_id()
        image_path = Path(path)
        self.images.append(image_path)
        rid = f"rId{len(self.images) + 1}"
        geom = "roundRect" if rounded else "rect"
        self.parts.append(
            f"""
<p:pic>
  <p:nvPicPr>
    <p:cNvPr id="{sid}" name="Picture {sid}" descr="{e(image_path.name)}"/>
    <p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr>
    <p:nvPr/>
  </p:nvPicPr>
  <p:blipFill>
    <a:blip r:embed="{rid}"/>
    <a:stretch><a:fillRect/></a:stretch>
  </p:blipFill>
  <p:spPr>
    <a:xfrm>
      <a:off x="{inch(x)}" y="{inch(y)}"/>
      <a:ext cx="{inch(w)}" cy="{inch(h)}"/>
    </a:xfrm>
    <a:prstGeom prst="{geom}"><a:avLst/></a:prstGeom>
  </p:spPr>
</p:pic>"""
        )

    def slide_xml(self) -> str:
        return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
       xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr>
        <a:xfrm>
          <a:off x="0" y="0"/>
          <a:ext cx="0" cy="0"/>
          <a:chOff x="0" y="0"/>
          <a:chExt cx="0" cy="0"/>
        </a:xfrm>
      </p:grpSpPr>
      {''.join(self.parts)}
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>"""


CREAM = "F6F0E5"
IVORY = "FFF9F0"
CHAR = "0B0D10"
COAL = "16181D"
YELLOW = "F5C242"
GOLD = "E0A833"
RED = "C94B38"
BLUE = "3C7DFF"
SLATE = "7E8793"
WHITE = "FFFFFF"
INK = "111111"
SOFT = "E9DFC9"
MUTED = "5A616C"


def small_label(s: SlideBuilder, x: float, y: float, text: str, color: str = YELLOW) -> None:
    s.rect(x, y, 1.1, 0.34, color, rounded=True)
    s.text(x, y + 0.02, 1.1, 0.2, text, size=10, color=CHAR, bold=True, align="c", all_caps=True)


def build_slides() -> list[tuple[str, list[Path]]]:
    cineplex = BASE / "cineplex.png"
    cntower = BASE / "cntower.jpeg"
    ubisoft = BASE / "ubisoft.jpeg"

    slides: list[tuple[str, list[Path]]] = []

    # 1 cover
    s = SlideBuilder()
    s.rect(0, 0, 13.333, 7.5, CHAR)
    s.rect(0.55, 0.55, 6.0, 6.4, CREAM, rounded=True)
    s.rect(6.75, 0.55, 6.05, 2.0, YELLOW, rounded=True)
    s.rect(6.75, 2.78, 2.9, 4.17, COAL, rounded=True)
    s.rect(9.9, 2.78, 2.9, 4.17, RED, rounded=True)
    s.image(6.95, 0.77, 5.65, 1.56, cineplex, rounded=True)
    s.image(6.98, 3.0, 2.5, 3.73, cntower, rounded=True)
    s.image(10.12, 3.0, 2.46, 3.73, ubisoft, rounded=True)
    small_label(s, 1.0, 1.0, "Deck")
    s.text(1.0, 1.55, 4.8, 1.8, "Recroom\nGames", size=28, color=INK, bold=True, font="Georgia")
    s.text(1.0, 3.65, 4.9, 1.25, "Digital infrastructure for online games, in-house development, sports programming, recreation venues, and immersive entertainment.", size=13, color=MUTED)
    s.text(1.0, 5.65, 4.9, 0.45, "13-slide company deck | May 2026", size=11.5, color="6B5940", bold=True)
    slides.append((s.slide_xml(), s.images))

    # 2 thesis
    s = SlideBuilder()
    s.rect(0, 0, 13.333, 7.5, CREAM)
    s.rect(0.65, 0.7, 12.0, 6.1, IVORY, rounded=True)
    s.rect(0.95, 1.05, 11.4, 0.14, YELLOW)
    s.text(1.0, 1.45, 5.8, 0.6, "The rec room is no longer only a venue.", size=27, color=INK, bold=True, font="Georgia")
    s.text(1.0, 2.25, 5.8, 0.95, "The customer journey now begins online, continues through identity, content, and booking, and loops back through loyalty, live events, and digital play.", size=13, color=MUTED)
    for i, (x, title, body, fill) in enumerate([
        (1.0, "Discover", "Games, events, leagues, and experiences need to feel alive before arrival.", "EFE5D0"),
        (4.15, "Transact", "Bookings, passes, registrations, and account access need one operating layer.", "111111"),
        (7.3, "Engage", "Online games, community participation, and seasonal drops keep users returning.", "E4C056"),
        (10.45, "Retain", "Membership, credits, team identity, and venue loyalty drive repeat demand.", "D55A43"),
    ]):
        s.rect(x, 4.0, 2.7, 2.0, fill, rounded=True)
        color = WHITE if fill in {"111111", "D55A43"} else INK
        s.text(x + 0.2, 4.2, 2.2, 0.35, title, size=16, color=color, bold=True)
        s.text(x + 0.2, 4.7, 2.25, 0.9, body, size=11.3, color=color)
        s.text(x + 2.15, 5.6, 0.3, 0.3, f"0{i+1}", size=18, color=color, bold=True, align="c")
    slides.append((s.slide_xml(), s.images))

    # 3 what platform controls
    s = SlideBuilder()
    s.rect(0, 0, 13.333, 7.5, COAL)
    s.rect(0.7, 0.7, 4.45, 6.1, YELLOW, rounded=True)
    s.text(1.05, 1.0, 3.5, 0.45, "Platform thesis", size=12, color=CHAR, bold=True, all_caps=True)
    s.text(1.05, 1.55, 3.6, 1.6, "One system for play, programming, access, and return.", size=27, color=CHAR, bold=True, font="Georgia")
    s.text(1.05, 3.55, 3.55, 1.35, "Recroom Games connects digital products with physical entertainment operations instead of treating them as separate businesses.", size=12.5, color="2A2620")
    for x, y, title, body in [
        (5.5, 1.0, "Online games", "Adaptive game products, browser play, and recurring digital engagement."),
        (8.55, 1.0, "Venue operations", "Sessions, parties, events, and guest scheduling from one surface."),
        (5.5, 3.2, "Sports and leagues", "Registrations, fixtures, court inventory, and recurring pass structures."),
        (8.55, 3.2, "AR / VR", "Immersive rooms, timed experiences, activations, and mixed-reality programs."),
        (5.5, 5.4, "Membership + identity", "Accounts, credits, loyalty, squads, return behavior, and pricing logic."),
        (8.55, 5.4, "Creator pipeline", "In-house builds, playtests, launches, and showcase programming."),
    ]:
        s.rect(x, y, 2.55, 1.55, "151A22", line="283142", rounded=True)
        s.text(x + 0.18, y + 0.18, 2.0, 0.32, title, size=14.5, color=WHITE, bold=True)
        s.text(x + 0.18, y + 0.56, 2.1, 0.7, body, size=10.5, color="D5DBE3")
    slides.append((s.slide_xml(), s.images))

    # 4 product surface
    s = SlideBuilder()
    s.rect(0, 0, 13.333, 7.5, IVORY)
    s.text(0.9, 0.95, 5.5, 0.55, "A digital entertainment stack with multiple revenue layers.", size=26, color=INK, bold=True, font="Georgia")
    s.text(0.9, 1.75, 5.5, 0.8, "The opportunity is not one app. It is a system that monetizes gameplay, access, events, memberships, and destination programming together.", size=12.7, color=MUTED)
    s.image(7.3, 0.8, 5.0, 2.65, cntower, rounded=True)
    cols = [
        ("Game products", "Online titles, digital tournaments, launch nights, AI-style game concepts."),
        ("Venue commerce", "Bookings, parties, premium sessions, ticketing, and timed attractions."),
        ("Community loops", "Leagues, squads, loyalty, digital rewards, and recurring player identity."),
        ("Brand extensions", "Partnership activations, creator programs, and live seasonal content."),
    ]
    x = 0.95
    for title, body in cols:
        s.rect(x, 4.05, 2.88, 2.05, CREAM if x < 3 else ("F8E08A" if x < 6 else "FFFFFF"), line="D6CAB3", rounded=True)
        s.text(x + 0.18, 4.28, 2.3, 0.35, title, size=15, color=INK, bold=True)
        s.text(x + 0.18, 4.75, 2.4, 0.95, body, size=10.7, color=MUTED)
        x += 3.07
    slides.append((s.slide_xml(), s.images))

    # 5 online games
    s = SlideBuilder()
    s.rect(0, 0, 13.333, 7.5, CHAR)
    s.image(0.72, 0.72, 3.85, 6.0, ubisoft, rounded=True)
    s.rect(4.85, 0.72, 7.76, 6.0, COAL, rounded=True)
    s.text(5.25, 1.05, 6.2, 0.48, "Online games are not a side module.", size=25, color=WHITE, bold=True, font="Georgia")
    s.text(5.25, 1.82, 6.2, 0.8, "The platform can support adaptive browser games, social competition, digital leagues, and venue-tied drops that keep the ecosystem alive between visits.", size=12.3, color="D5DBE3")
    items = [
        "Adaptive browser play",
        "Seasonal digital releases",
        "Game-linked memberships",
        "Play Game auth funnel",
    ]
    y = 3.0
    for item in items:
        s.rect(5.25, y, 6.4, 0.7, "151A22", line="2E3745", rounded=True)
        s.text(5.5, y + 0.16, 5.8, 0.3, item, size=13, color=WHITE, bold=True)
        y += 0.9
    s.text(5.25, 6.1, 2.2, 0.35, "Example titles", size=11, color=YELLOW, bold=True, all_caps=True)
    s.text(7.1, 6.08, 4.2, 0.35, "Neon Arena • Signal Chase • Pulse Circuit • Echo Room", size=11, color="D5DBE3")
    slides.append((s.slide_xml(), s.images))

    # 6 sports and rec
    s = SlideBuilder()
    s.rect(0, 0, 13.333, 7.5, YELLOW)
    s.rect(0.85, 0.75, 4.2, 6.0, COAL, rounded=True)
    s.text(1.2, 1.1, 3.3, 0.4, "Sports + recreation", size=12, color=YELLOW, bold=True, all_caps=True)
    s.text(1.2, 1.62, 3.3, 1.45, "A real system for leagues, courts, sessions, and repeat access.", size=25, color=WHITE, bold=True, font="Georgia")
    s.text(1.2, 3.45, 3.2, 1.1, "The same platform can handle registrations, passes, schedules, tournaments, and live programming across sports and recreation models.", size=12.2, color="D5DBE3")
    s.image(5.35, 0.75, 7.1, 3.05, cntower, rounded=True)
    for x, title, body, fill in [
        (5.35, "League scheduling", "Fixtures, brackets, recurring sessions, and event timing.", "FFF7E0"),
        (8.0, "Membership bundles", "Family plans, drop-ins, season passes, and premium access.", "F2E7BF"),
        (10.65, "Venue utilization", "Fill off-peak time with structured blocks and social play.", "F8E08A"),
    ]:
        s.rect(x, 4.22, 2.45, 1.92, fill, rounded=True)
        s.text(x + 0.16, 4.44, 2.0, 0.32, title, size=14, color=INK, bold=True)
        s.text(x + 0.16, 4.86, 2.0, 0.85, body, size=10.6, color="4E4A42")
    slides.append((s.slide_xml(), s.images))

    # 7 AR/VR
    s = SlideBuilder()
    s.rect(0, 0, 13.333, 7.5, IVORY)
    s.rect(0.7, 0.75, 12.0, 6.0, CREAM, rounded=True)
    s.text(0.98, 1.05, 5.6, 0.55, "AR / VR can sit inside the same operating stack.", size=25, color=INK, bold=True, font="Georgia")
    s.text(0.98, 1.85, 5.5, 0.75, "Immersive rooms, activations, branded experiments, and mixed-reality experiences still need the same foundations: scheduling, identity, pricing, and content control.", size=12.2, color=MUTED)
    s.image(7.15, 1.0, 5.0, 2.5, cineplex, rounded=True)
    s.rect(1.0, 3.15, 11.1, 0.12, BLUE)
    for x, title in [(1.0, "Session timing"), (3.7, "Queue logic"), (6.4, "Event packaging"), (9.1, "Guest identity")]:
        s.text(x, 3.45, 2.0, 0.35, title, size=13.5, color=INK, bold=True)
    s.text(1.0, 4.15, 10.8, 1.2, "The value is not only the attraction itself. The value is the ability to package it with the rest of the business and make it visible across the digital front door.", size=18, color="3A4350", bold=True)
    slides.append((s.slide_xml(), s.images))

    # 8 development
    s = SlideBuilder()
    s.rect(0, 0, 13.333, 7.5, COAL)
    s.rect(0.8, 0.8, 12.0, 5.9, "11151B", rounded=True)
    s.text(1.1, 1.1, 5.4, 0.48, "In-house game development belongs inside the brand system.", size=25, color=WHITE, bold=True, font="Georgia")
    s.text(1.1, 1.85, 5.4, 0.8, "Recroom Games is not only operating venues. It can also originate digital products, test them with real audiences, and turn the venue into a launch channel.", size=12.1, color="D5DBE3")
    s.image(7.55, 1.05, 4.7, 4.85, ubisoft, rounded=True)
    for y, title, body in [
        (3.1, "In-house development", "Build original game concepts, attractions, and digital entertainment products."),
        (4.2, "Playtests and release loops", "Use closed tests, launch nights, and audience feedback before wider release."),
        (5.3, "Creator programming", "Support student work, indie collaborations, and experimental drops tied to the venue."),
    ]:
        s.rect(1.1, y, 5.8, 0.82, "191F28", line="2F3947", rounded=True)
        s.text(1.3, y + 0.13, 2.35, 0.24, title, size=12.8, color=YELLOW, bold=True)
        s.text(3.95, y + 0.12, 2.65, 0.36, body, size=10.2, color="E0E6EE")
    slides.append((s.slide_xml(), s.images))

    # 9 operator page thesis
    s = SlideBuilder()
    s.rect(0, 0, 13.333, 7.5, CREAM)
    s.text(0.9, 1.0, 5.6, 0.55, "Built for operators across multiple entertainment models.", size=26, color=INK, bold=True, font="Georgia")
    s.text(0.9, 1.8, 5.8, 0.85, "The platform is designed for businesses that mix digital products, destination programming, scheduled access, and recurring membership demand.", size=12.3, color=MUTED)
    blocks = [
        ("Arcade + esports venues", "Tournaments, gamer identity, passes, and digital competition."),
        ("Sports + recreation centers", "Leagues, drop-ins, programs, and recurring access logic."),
        ("Immersive attractions", "AR / VR rooms, premium sessions, and packageable experiences."),
        ("Hybrid entertainment brands", "Gaming, venue programming, parties, and content drops together."),
    ]
    x = 0.95
    y = 3.2
    for i, (title, body) in enumerate(blocks):
        s.rect(x, y, 2.85, 2.2, "FFFFFF" if i % 2 == 0 else "F3E7BD", line="DACBAE", rounded=True)
        s.text(x + 0.16, y + 0.2, 2.2, 0.34, title, size=14, color=INK, bold=True)
        s.text(x + 0.16, y + 0.72, 2.25, 0.95, body, size=10.6, color=MUTED)
        x += 3.05
    slides.append((s.slide_xml(), s.images))

    # 10 memberships/community
    s = SlideBuilder()
    s.rect(0, 0, 13.333, 7.5, CHAR)
    s.image(0.75, 0.75, 4.0, 6.0, cineplex, rounded=True)
    s.rect(5.0, 0.75, 7.55, 6.0, YELLOW, rounded=True)
    s.text(5.35, 1.1, 6.0, 0.48, "Community is a revenue loop, not a marketing extra.", size=24.5, color=CHAR, bold=True, font="Georgia")
    s.text(5.35, 1.84, 6.0, 0.78, "Player profiles, clubs, ladders, seasonal drops, and loyalty systems convert one-time traffic into recurring participation.", size=12.1, color="3E372F")
    for i, (title, body) in enumerate([
        ("Player profiles", "Persistent identity across games, venue sessions, and rewards."),
        ("Clubs + squads", "Team structures for social participation and recurring competition."),
        ("Seasonal loops", "Events, milestones, challenges, and return-driven progression."),
        ("Digital loyalty", "Credits, premium access, and membership-grade retention pathways."),
    ]):
        x = 5.35 + (i % 2) * 3.2
        y = 3.05 + (i // 2) * 1.55
        s.rect(x, y, 2.8, 1.2, "F8E8AE", line="C9A645", rounded=True)
        s.text(x + 0.16, y + 0.14, 2.1, 0.28, title, size=13.2, color=CHAR, bold=True)
        s.text(x + 0.16, y + 0.48, 2.3, 0.5, body, size=10.2, color="4C453D")
    slides.append((s.slide_xml(), s.images))

    # 11 partnerships
    s = SlideBuilder()
    s.rect(0, 0, 13.333, 7.5, IVORY)
    s.text(0.9, 1.0, 5.8, 0.55, "The brand can live across entertainment, destination, and content partnerships.", size=25, color=INK, bold=True, font="Georgia")
    s.text(0.9, 1.78, 5.9, 0.8, "Partnerships are not cosmetic. They can drive discovery, live programming, co-branded launches, and venue traffic.", size=12.3, color=MUTED)
    for x, img, label, fill in [
        (0.95, cineplex, "Entertainment partner", "FFFFFF"),
        (4.43, cntower, "Destination signal", "F7EBC6"),
        (7.91, ubisoft, "Game creation partner", "FFFFFF"),
    ]:
        s.rect(x, 3.0, 3.0, 2.7, fill, line="D5C7B3", rounded=True)
        s.image(x + 0.2, 3.18, 2.6, 1.75, img, rounded=True)
        s.text(x + 0.2, 5.15, 2.6, 0.35, label, size=11.8, color=INK, bold=True, align="c")
    slides.append((s.slide_xml(), s.images))

    # 12 rollout
    s = SlideBuilder()
    s.rect(0, 0, 13.333, 7.5, RED)
    s.rect(0.78, 0.75, 11.8, 6.0, CHAR, rounded=True)
    s.text(1.1, 1.1, 5.4, 0.5, "How rollout can happen.", size=25, color=WHITE, bold=True, font="Georgia")
    s.text(1.1, 1.85, 5.4, 0.75, "The model can start narrow and expand into a broader rec room operating layer.", size=12, color="E7D8D4")
    phases = [
        ("Phase 1", "Launch the digital front door", "Bookings, account creation, support, and visible venue programming."),
        ("Phase 2", "Add online game loops", "Play Game flows, recurring titles, tournaments, and community identity."),
        ("Phase 3", "Expand into deeper infrastructure", "Membership, leagues, creator releases, and immersive programming."),
    ]
    y = 3.0
    for phase, heading, body in phases:
        s.rect(1.1, y, 11.0, 0.95, "171A20", line="303745", rounded=True)
        s.text(1.35, y + 0.16, 1.35, 0.25, phase, size=12.5, color=YELLOW, bold=True)
        s.text(2.6, y + 0.14, 2.8, 0.28, heading, size=13.5, color=WHITE, bold=True)
        s.text(5.5, y + 0.14, 5.8, 0.36, body, size=10.4, color="D8DEE6")
        y += 1.25
    slides.append((s.slide_xml(), s.images))

    # 13 close
    s = SlideBuilder()
    s.rect(0, 0, 13.333, 7.5, CREAM)
    s.rect(0.85, 0.85, 5.6, 5.8, CHAR, rounded=True)
    s.rect(6.7, 0.85, 5.8, 5.8, WHITE, line="D6CAB4", rounded=True)
    s.text(1.2, 1.25, 4.4, 1.1, "Recroom\nGames", size=29, color=WHITE, bold=True, font="Georgia")
    s.text(1.2, 2.95, 4.3, 1.0, "A digital entertainment platform for online games, venue operations, sports programming, creator launches, and immersive play.", size=12.5, color="E6EAF0")
    s.text(7.05, 1.3, 3.5, 0.4, "Contact", size=18, color=INK, bold=True)
    s.text(7.05, 2.0, 4.4, 0.45, "support@recroomgames.org", size=15, color=RED, bold=True)
    s.text(7.05, 2.75, 4.7, 0.8, "255 Bremner Blvd\nToronto, ON M5V 3L9, Canada", size=12.5, color=MUTED)
    s.text(7.05, 4.55, 4.7, 0.7, "Build the online layer that makes the venue feel active before people arrive.", size=17, color=INK, bold=True)
    s.text(7.05, 5.45, 4.4, 0.35, "13-slide deck generated for Recroom Games", size=10.8, color="7B705F")
    slides.append((s.slide_xml(), s.images))

    return slides


def slide_rels(image_targets: list[str]) -> str:
    image_rels = "".join(
        f'\n  <Relationship Id="rId{i + 2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/{target}"/>'
        for i, target in enumerate(image_targets)
    )
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout7.xml"/>
{image_rels}
</Relationships>"""


def core_xml() -> str:
    now = datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
                   xmlns:dc="http://purl.org/dc/elements/1.1/"
                   xmlns:dcterms="http://purl.org/dc/terms/"
                   xmlns:dcmitype="http://purl.org/dc/dcmitype/"
                   xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>Recroom Games Pitch Deck</dc:title>
  <dc:creator>OpenAI Codex</dc:creator>
  <cp:lastModifiedBy>OpenAI Codex</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">{now}</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">{now}</dcterms:modified>
</cp:coreProperties>"""


def app_xml(slide_count: int) -> str:
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><TotalTime>1</TotalTime><Words>0</Words><Application>Microsoft Macintosh PowerPoint</Application><PresentationFormat>On-screen Show (16:9)</PresentationFormat><Paragraphs>0</Paragraphs><Slides>{slide_count}</Slides><Notes>0</Notes><HiddenSlides>0</HiddenSlides><MMClips>0</MMClips><ScaleCrop>false</ScaleCrop><HeadingPairs><vt:vector size="4" baseType="variant"><vt:variant><vt:lpstr>Theme</vt:lpstr></vt:variant><vt:variant><vt:i4>1</vt:i4></vt:variant><vt:variant><vt:lpstr>Slide Titles</vt:lpstr></vt:variant><vt:variant><vt:i4>0</vt:i4></vt:variant></vt:vector></HeadingPairs><TitlesOfParts><vt:vector size="1" baseType="lpstr"><vt:lpstr>Office Theme</vt:lpstr></vt:vector></TitlesOfParts><Manager></Manager><Company></Company><LinksUpToDate>false</LinksUpToDate><SharedDoc>false</SharedDoc><HyperlinkBase></HyperlinkBase><HyperlinksChanged>false</HyperlinksChanged><AppVersion>14.0000</AppVersion></Properties>"""


def presentation_xml_from_template(slide_count: int) -> str:
    slide_ids = "".join(f'<p:sldId id="{255 + i}" r:id="rId{i + 6}"/>' for i in range(1, slide_count + 1))
    return f"""<?xml version='1.0' encoding='UTF-8' standalone='yes'?>
<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" saveSubsetFonts="1" autoCompressPictures="0"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst><p:sldIdLst>{slide_ids}</p:sldIdLst><p:sldSz cx="12191695" cy="6858000" type="screen16x9"/><p:notesSz cx="6858000" cy="9144000"/><p:defaultTextStyle><a:defPPr><a:defRPr lang="en-US"/></a:defPPr><a:lvl1pPr marL="0" algn="l" defTabSz="457200" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl1pPr><a:lvl2pPr marL="457200" algn="l" defTabSz="457200" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl2pPr><a:lvl3pPr marL="914400" algn="l" defTabSz="457200" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl3pPr><a:lvl4pPr marL="1371600" algn="l" defTabSz="457200" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl4pPr><a:lvl5pPr marL="1828800" algn="l" defTabSz="457200" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl5pPr><a:lvl6pPr marL="2286000" algn="l" defTabSz="457200" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl6pPr><a:lvl7pPr marL="2743200" algn="l" defTabSz="457200" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl7pPr><a:lvl8pPr marL="3200400" algn="l" defTabSz="457200" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl8pPr><a:lvl9pPr marL="3657600" algn="l" defTabSz="457200" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl9pPr></p:defaultTextStyle></p:presentation>"""


def presentation_rels_from_template(slide_count: int) -> str:
    fixed = [
        ("rId1", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster", "slideMasters/slideMaster1.xml"),
        ("rId2", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings", "printerSettings/printerSettings1.bin"),
        ("rId3", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps", "presProps.xml"),
        ("rId4", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/viewProps", "viewProps.xml"),
        ("rId5", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme", "theme/theme1.xml"),
        ("rId6", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles", "tableStyles.xml"),
    ]
    slides = [
        (f"rId{i + 6}", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide", f"slides/slide{i}.xml")
        for i in range(1, slide_count + 1)
    ]
    rels = "".join(f'<Relationship Id="{rid}" Type="{typ}" Target="{target}"/>' for rid, typ, target in fixed + slides)
    return f"""<?xml version='1.0' encoding='UTF-8' standalone='yes'?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{rels}</Relationships>"""


def content_types_from_template(slide_count: int) -> str:
    overrides = "".join(
        f'<Override PartName="/ppt/slides/slide{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>'
        for i in range(1, slide_count + 1)
    )
    return f"""<?xml version='1.0' encoding='UTF-8' standalone='yes'?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="bin" ContentType="application/vnd.openxmlformats-officedocument.presentationml.printerSettings"/><Default Extension="jpeg" ContentType="image/jpeg"/><Default Extension="jpg" ContentType="image/jpeg"/><Default Extension="png" ContentType="image/png"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/ppt/presProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presProps+xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/><Override PartName="/ppt/slideLayouts/slideLayout10.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/><Override PartName="/ppt/slideLayouts/slideLayout11.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/><Override PartName="/ppt/slideLayouts/slideLayout2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/><Override PartName="/ppt/slideLayouts/slideLayout3.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/><Override PartName="/ppt/slideLayouts/slideLayout4.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/><Override PartName="/ppt/slideLayouts/slideLayout5.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/><Override PartName="/ppt/slideLayouts/slideLayout6.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/><Override PartName="/ppt/slideLayouts/slideLayout7.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/><Override PartName="/ppt/slideLayouts/slideLayout8.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/><Override PartName="/ppt/slideLayouts/slideLayout9.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/><Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>{overrides}<Override PartName="/ppt/tableStyles.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml"/><Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/><Override PartName="/ppt/viewProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml"/></Types>"""


def write_pptx() -> None:
    slides = build_slides()
    media_map: dict[Path, str] = {}
    media_index = 1
    for _, image_paths in slides:
        for image_path in image_paths:
            if image_path not in media_map:
                suffix = image_path.suffix.lower()
                ext = ".jpg" if suffix == ".jpg" else ".jpeg" if suffix == ".jpeg" else ".png"
                media_map[image_path] = f"image{media_index}{ext}"
                media_index += 1
    with zipfile.ZipFile(TEMPLATE) as src, zipfile.ZipFile(OUT, "w", compression=zipfile.ZIP_DEFLATED) as z:
        keep = []
        for name in src.namelist():
            if name.startswith("ppt/slides/slide") or name.startswith("ppt/slides/_rels/slide"):
                continue
            if name.startswith("ppt/media/"):
                continue
            if name in {"ppt/presentation.xml", "ppt/_rels/presentation.xml.rels", "[Content_Types].xml", "docProps/core.xml", "docProps/app.xml"}:
                continue
            keep.append(name)
        for name in keep:
            z.writestr(name, src.read(name))
        for image_path, media_name in media_map.items():
            z.writestr(f"ppt/media/{media_name}", image_path.read_bytes())
        z.writestr("[Content_Types].xml", content_types_from_template(len(slides)))
        z.writestr("docProps/core.xml", core_xml())
        z.writestr("docProps/app.xml", app_xml(len(slides)))
        z.writestr("ppt/presentation.xml", presentation_xml_from_template(len(slides)))
        z.writestr("ppt/_rels/presentation.xml.rels", presentation_rels_from_template(len(slides)))
        for i, (xml, image_paths) in enumerate(slides, start=1):
            z.writestr(f"ppt/slides/slide{i}.xml", xml)
            targets = [media_map[p] for p in image_paths]
            z.writestr(f"ppt/slides/_rels/slide{i}.xml.rels", slide_rels(targets))


if __name__ == "__main__":
    write_pptx()
    print(OUT)
