# eMedi 허가정보 수집기

A desktop app that collects medical device permit details from the MFDS medical device portal ([emedi.mfds.go.kr](https://emedi.mfds.go.kr)) into Excel, then merges in the permit list you download from 전자민원.

Built with PyQt6 and Selenium. Works on Windows and macOS.

## Features

- **① 크롤링:** open Chrome from the app, search on the site as usual, then press start. The app walks every result page, opens each item's detail popup, and saves everything to `.xlsx` and `.json`.
  - By default each popup is read after a fixed 10-second wait, the proven original method.
  - Optional **빠른 모드 (실험적)** reads a popup as soon as it has loaded. If that fails for an item, the item is retried with the original method.
  - When an item fails, the log shows what the popup looked like, which helps find the cause.
  - Failed items are retried, and stopping or a crash still saves what was collected.
  - **누락분만 수집:** after reopening Chrome and repeating the same site search, choose an existing `_partial.json`, final `.json`, or crawl `.xlsx` and click **누락분만 수집**. Completed detail records are kept; failed, list-only, and not-yet-collected rows are requested again. An incomplete popup is not counted as success.
  - The existing file is never overwritten. Recovery produces a new numbered Excel/JSON pair; while recovering, a separate `_partial.json` checkpoint is updated after each item. You can select that checkpoint again if the connection drops. If the old rows are missing from the current search list, the app stops rather than silently losing data.
  - Live progress: total, done, failed, waiting, time left. The counters are refreshed again after stopping and saving; waiting rows are not mislabeled as failures.
- The main screen shows a compact results table (`허가번호 / 항목 / 상태`) with the newest work at the bottom. The current item is inverted, failed items are red, and checkpoint saves appear as blue rows. **상세 로그 펼치기** reveals a bounded 250-line diagnostic view. **로그 지우기**, starting a new run, or pressing Stop clears the previous on-screen log; the full diagnostic history remains in `~/emedi_crawler.log`.
- Chrome opens maximized. The desktop window starts near 1264×1180 where the screen allows it, shrinks to the available screen, and leaves the results table to fill the remaining height. Its black title strip supports dragging, minimize, maximize, and close.
- IBM Plex Sans KR and IBM Plex Mono are included under the SIL Open Font License (`fonts/LICENSE.txt`), so the UI does not depend on fonts installed on the computer.
- **② 허가사항 병합:** adds the columns from 전자민원's 품목 허가(신고)사항 export to the crawl result, matched by permit number, as a formatted Excel file.
- You choose the save folder and file name (a dated default is filled in). If the name already exists, `_2`, `_3`, … is added.
- **파일 열기 / 폴더 열기** buttons open the result when a run finishes.
- UDI columns separate multiple values with `; `, because values can contain commas.

For example, if a 125-row partial file has 120 complete rows and 5 failures, and the current search shows 227 results, recovery skips the 120 complete rows and collects the remaining 107. Keep the Chrome search filters the same as the original run. Once it finishes, use the new Excel/JSON output for subsequent work.

## Requirements

- Python 3.10+
- Google Chrome. The matching ChromeDriver is downloaded automatically on first run, so the first run needs internet.

## Run from source

```bash
pip install -r requirements.txt
python app.py
```

## Build a double-clickable app

Build on the OS you want to run it on.

| OS | How | Result |
|---|---|---|
| Windows | double-click `build_windows.bat` | `dist\emedi_crawler\emedi_crawler.exe`. Keep the whole folder together and make a desktop shortcut to the exe. |
| macOS | `bash build_mac.sh` | `dist/emedi_crawler.app`. Drag it to Applications. The first time, right-click it and choose Open. |

The built app doesn't need Python, only Chrome. Rebuild after changing `app.py` or `recovery.py`.

If something goes wrong in the built app, details are written to `~/emedi_crawler.log`.

## Tuning

Wait times are constants at the top of `app.py`:

| Constant | Default | Meaning |
|---|---|---|
| `MODAL_TIMEOUT` | 20 s | Max wait for a detail popup |
| `MODAL_SETTLE` | 1.0 s | Popup content must stay unchanged this long to count as loaded |
| `POLITE_DELAY` | 0.3 s | Pause between items, to be gentle on the site |
| `MODAL_WAIT` | 10 s | Fixed wait per item (default mode) |

## Notes

The crawler relies on the site's current page structure (the `getList` / `itemDetail` functions and the popup's table captions). If the site changes, `PARSE_JS` and the selectors near the top of `app.py` are where to look.

---

Dev. Yuri Kim · 2026-10
