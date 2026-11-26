# NPI HUB

One page, one tab per dashboard. Each tab is the original dashboard file, copied byte-for-byte and shown in its own frame, so its data, filters and logic are untouched.

## Refresh and publish process (agreed)

Each tab is refreshed on its own, shown for review, and published only after approval.

| Tab | You do | Bob does (refresh, then stops for your review) | After you approve |
|---|---|---|---|
| ESS Installs | ask for a refresh (needs VPN/office network and an up-to-date P11 in `...\.bob\ESS`) | `.\REFRESH_HUB.ps1 -Tab ess` | `.\PUBLISH_HUB.ps1 -Tab ess -Approved` |
| Daily NPI EO Orders | update the Masterfile daily (`...\playground\NPI Tracker\Masterfile_<date>.xlsx`, closed in Excel), then ask for a refresh | `.\REFRESH_HUB.ps1 -Tab npi [-Master "<file>"]` | `.\PUBLISH_HUB.ps1 -Tab npi -Approved` |

- A refresh rebuilds `hub\`, runs `check-hub.ps1`, and stops. **Nothing is pushed.** Review `hub\index.html`.
- `PUBLISH_HUB.ps1 -Tab <tab> -Approved` publishes only that tab's dashboard; the other tab on the live site is not touched.
- `PUBLISH_HUB.ps1` refuses without `-Approved`, and refuses unless `check-hub.ps1` passed for the current build.
- `-Tab all` (refresh or publish) handles both. `-RebuildOnly` re-copies the dashboards as they are now and re-checks.
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

