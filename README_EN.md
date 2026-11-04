# Rainbow Six Siege Operator Rankings · Y11S3

[中文版](README.md) | [English](README_EN.md)

This community project combines Athieno's operator ratings, weapon fire rates, operator speed, rare weapons, and secondary gadgets into five Chinese ranking reports. The current published examples cover 78 operators and were last updated on 2026-09-22 (UTC+8). The rating source is [Athieno's complete Y11S3 tier list](https://youtu.be/AWT8ckw7wMo).

## Download the reports

Each report has an XLSX workbook, a PDF, and three page images. The images show attackers, defenders, and patch notes. Open the [report directory](docs/outputs/) or use the links below.

| Ranking | XLSX | PDF | Page images |
|---|---|---|---|
| Video rating | [Download](docs/outputs/视频评分榜.xlsx) | [Open](docs/outputs/视频评分榜.pdf) | [View](docs/outputs/图片版/视频评分榜/) |
| Primary weapon fire rate | [Download](docs/outputs/主武器射速榜.xlsx) | [Open](docs/outputs/主武器射速榜.pdf) | [View](docs/outputs/图片版/主武器射速榜/) |
| Operator speed | [Download](docs/outputs/速度榜.xlsx) | [Open](docs/outputs/速度榜.pdf) | [View](docs/outputs/图片版/速度榜/) |
| Rare weapons | [Download](docs/outputs/稀有枪械榜.xlsx) | [Open](docs/outputs/稀有枪械榜.pdf) | [View](docs/outputs/图片版/稀有枪械榜/) |
| Secondary gadgets | [Download](docs/outputs/次要装备榜.xlsx) | [Open](docs/outputs/次要装备榜.pdf) | [View](docs/outputs/图片版/次要装备榜/) |

Within a ranking group, ties are resolved using video score, highest primary weapon fire rate, speed, rare weapons, and secondary gadgets, in that order. Reports are written in Chinese.

## Requirements for rebuilding

Rebuilding is for development or data updates. To read the published reports, download the files above; no local setup is needed.

- Windows
- Python 3.9 or newer
- `curl.exe`
- Dependencies from [requirements.txt](requirements.txt)

Create an environment in PowerShell:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
```

To refresh Athieno's rating source, use the project skill in [skills/build-r6-operator-report/SKILL.md](skills/build-r6-operator-report/SKILL.md). It verifies the latest complete tier list and updates `data/athieno/latest.json`.

To collect the remaining sources and rebuild all reports, run:

```powershell
.\src\rebuild_leaderboards.bat
```

The batch file runs `collector`, `operator_stats`, and `leaderboards` in order, and stops if a stage fails. New reports go to `~outputs/` for review. Only reviewed examples used by this README belong in `docs/outputs/`.

To rebuild from the existing `data/` snapshot without collecting new data:

```powershell
$env:PYTHONPATH='src'
$env:PYTHONUTF8='1'
python -m r6_report.operator_stats --data-dir data --output '~temp/r6_operator_stats.xlsx'
python -m r6_report.leaderboards --data-dir data --input '~temp/r6_operator_stats.xlsx' --output-dir '~outputs'
```

Run the test suite with:

```powershell
$env:PYTHONPATH='src'
$env:PYTHONUTF8='1'
python -m unittest discover -s tests -v
```

## Sources

- [Athieno's Y11S3 tier list](https://youtu.be/AWT8ckw7wMo)
- [Rainbow Six Siege Huiji Wiki](https://r6s.huijiwiki.com/)
- [Official Rainbow Six Siege website](https://www.ubisoft.com/en-us/game/rainbow-six/siege)

This is a community report project and is not affiliated with Ubisoft, Athieno, or Huiji Wiki. Names, trademarks, and source materials belong to their respective owners.
