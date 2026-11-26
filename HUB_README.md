# NPI HUB

One page, one tab per dashboard. Each tab is the original dashboard file, copied byte-for-byte and shown in its own frame, so its data, filters and logic are untouched.

## Daily use

| Tab | When | Steps |
|---|---|---|
| Daily NPI EO Orders | daily | Save the new `Masterfile_<date>.xlsx` in `...\playground\NPI Tracker\`, close it in Excel, then ask Bob to refresh (`.\REFRESH_HUB.ps1 -Tab npi [-Master "<file>"]`) |
| ESS Installs | on request | Ask Bob to refresh (`.\REFRESH_HUB.ps1 -Tab ess`). Needs VPN/office network and an up-to-date P11 (2025).xlsx in `...\.bob\ESS` |
| both | | `.\REFRESH_HUB.ps1` (or `-Tab all`) |

Every refresh rebuilds `hub\`, runs `check-hub.ps1`, and stops. **Nothing is pushed.** Review `hub\index.html`, then:

```
.\PUBLISH_HUB.ps1 -Approved
```

`PUBLISH_HUB.ps1` refuses without `-Approved`, and refuses unless `check-hub.ps1` passed for the current build.
`-RebuildOnly` re-copies the dashboards as they are now (no data refresh) and re-checks.

## Add a new dashboard tab

1. Make sure the dashboard is a single self-contained `.html` file.
2. Add an entry to `hub-tabs.json`:
   ```json
   { "id": "my-dashboard", "label": "My Dashboard", "source": "C:\\path\\to\\My_Dashboard.html", "file": "my-dashboard.html" }
   ```
3. Run `.\REFRESH_HUB.ps1 -RebuildOnly`. The tab appears, is verified identical to its source, and is click-tested.
4. Review, then `.\PUBLISH_HUB.ps1 -Approved`.

If the new dashboard has its own refresh script, add a branch for it in `REFRESH_HUB.ps1`.

## Files

- `hub.template.html` - hub shell (header, tab bar, lazy-loaded frames)
- `hub-tabs.json` - the tab list
- `Build-Hub.ps1` - builds `hub\` (verifies SHA-256 of each copy)
- `check-hub.ps1` - byte identity, headless browser load, every tab and inner tab clicked, 0 JS errors
- `REFRESH_HUB.ps1`, `PUBLISH_HUB.ps1` - refresh / publish
- `PUBLISH_ESS.ps1` - retired (it would overwrite the hub's `index.html`)

## Revert to the ESS-only dashboard

```
git checkout main        # before publishing: nothing changed
git revert <hub commit>  # after publishing
```
Tag `pre-npi-hub` and the folder `backup-pre-npi-hub\` hold the ESS-only version.
