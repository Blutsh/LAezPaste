# LAezPaste - Copy Log Analytics to Teams

Firefox extension that adds copy buttons to Azure Log Analytics results for easy sharing in Teams, Outlook, and other applications.


https://github.com/user-attachments/assets/8ea2f747-4a30-4348-b5db-4863876c250e



## Features

- Adds **two buttons** next to "Add bookmark" in Log Analytics results toolbar
  - **Markdown** - Copy as markdown table (for GitHub, Slack, etc.)
  - **HTML** - Copy as rich HTML table (for Teams, Outlook, Word)
- Works with all visible columns
- Handles special characters and escapes properly

## Installation (Firefox)

1. Open Firefox and navigate to `about:debugging`
2. Click **"This Firefox"** in the left sidebar
3. Click **"Load Temporary Add-on..."**
4. Navigate to the extension folder
5. Select the `manifest.json` file
6. ✅ Extension loaded!

**Note:** This is a temporary installation - you'll need to reload it each time you restart Firefox.

## Usage

1. Navigate to Azure Log Analytics (Sentinel or Log Analytics workspace)
2. Run a query to get results
3. **Select one or more rows** by clicking the checkboxes
4. Click either:
   - **Markdown** button → paste anywhere that supports markdown
   - **HTML** button → paste into Teams/Outlook/Word for a formatted table
5. ✅ Table copied to clipboard!

## Files

```
LAezPaste/
├── manifest.json      # Extension configuration
├── content.js         # Main script
├── icon48.png         # Extension icon (48x48)
├── icon96.png         # Extension icon (96x96)
├── example.mp4        # Example video
└── README.md          # This file
```

## Troubleshooting

| Issue | Solution |
|-------|----------|
| Buttons don't appear | Open DevTools (F12) → Console, check for `[LAezPaste]` messages |
| Copy doesn't work | Check browser clipboard permissions |
| Wrong data copied | Ensure rows are selected via checkboxes (not just clicked) |
| Extension disappeared | Temporary add-ons are removed on browser restart - reload it |

## Permanent Installation Options

### Option 1: Firefox Developer Edition (Recommended)
1. Install [Firefox Developer Edition](https://www.mozilla.org/firefox/developer/)
2. Go to `about:config`
3. Set `xpinstall.signatures.required` to `false`
4. Install the extension permanently

