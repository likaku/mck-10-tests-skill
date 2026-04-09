# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/).

---

## [2.0.0] - 2026-04-09

### Changed
- **License**: MIT → Apache 2.0 (author: Kaku Li)
- **Report template**: Full-length → half-length concise format
  - Executive summary now declares "本文为初始想法的梳理，而非最终结论"
  - Added From→To key shifts with BECAUSE rationale (natural language, not mechanical)
  - 10-dimension findings condensed into a single table
  - Action plan simplified to 3 items: 最紧急 / 最重要 / Quick Win
  - Recommendations must include quantified targets or deadlines (no vague "加强关注")
- **Report card renderer** (`scripts/render_report_card.py`): Complete rewrite (v4)
  - Chinese font: SongTi → **KaiTi** (楷体)
  - English font: unchanged (Arial)
  - Font sizes significantly increased: title 38pt, section 30pt, body 22pt, advice 18pt
  - Text color deepened: body `(30,30,30)` near-black, meta `(70,70,70)`
  - From/To mechanical labels → natural language "核心判断" narrative with numbered circles
  - Radar chart now rendered via matplotlib (better quality than PIL-only version)
  - Layout tightened: reduced whitespace, increased information density
- **Dual output**: Skill now generates both Markdown report + PNG long-image report card

### Added
- `CHANGELOG.md` (this file)
- Author attribution in LICENSE (Kaku Li)
- `NOTICE` section in README

### Removed
- Verbose per-dimension detailed report sections (replaced by compact table)
- Time-layered action plan (48h/1w/1m → 3 priority items)

## [1.1.0] - 2026-04-08

### Added
- Report card renderer (`scripts/render_report_card.py`) — first version
- 10 scenario-specific persona roles with ASCII art avatars
- Emotion system (greet / probe / challenge / affirm / wrap)

### Changed
- Renamed all "McKinsey" references to "MCK" in README

## [1.0.0] - 2026-04-08

### Added
- Initial release: SKILL.md with full 10-test coaching workflow
- `references/methodology.md` — detailed methodology reference
- README with usage instructions, persona showcase, and design philosophy
- MIT License
