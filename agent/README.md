# EHEQ Monthly Presentation Agent

This local Windows agent uses Microsoft PowerPoint to turn a monthly `.pptx` file into a standalone EHEQ presentation webpage that keeps the manager's fullscreen layout and navigation style.

## One-time setup

1. Ensure Microsoft PowerPoint is installed on the Windows machine.
2. Put monthly PowerPoint files in `monthly-input`.
3. Run the watcher:

```powershell
.\agent\watch-monthly-folder.ps1
```

Generated reviews are written to `monthly-output\<Month>-<Year>\index.html`, with slide images in the matching `slides` folder and extracted text in `slide-manifest.json`.

## Process one presentation

```powershell
.\agent\run-monthly-agent.ps1 `
  -PowerPointPath ".\monthly-input\EHEQ Monthly Review October 2026.pptx" `
  -Month "October" `
  -Year "2026"
```

## Automation boundary

The current agent automates the reliable presentation operations locally:

- detects and processes monthly PowerPoint files;
- exports every slide at presentation quality;
- extracts slide text and section metadata;
- generates a standalone webpage using the cloned manager layout.

The generated `slide-manifest.json` is the handoff point for a future AI layer that can classify slide content, detect changed sections, validate required monthly sections, or populate the data-driven application.
