# TeGe Outlook Add-in PoC

Internal Outlook add-in proof of concept for New Outlook and Outlook Web.

## What It Does

- Runs when a user starts composing a new message.
- Detects the current sender/mailbox address where Outlook exposes it.
- Inserts the matching TeGe HTML signature with `setSignatureAsync`.
- Provides a task pane to manually choose and insert a signature.

## Hosting

The manifest expects the add-in files to be hosted at:

```text
https://jorgecastelan-svg.github.io/tege-signature-assets/addin/
```

Upload the contents of this `addin/` folder to:

```text
github-pages/addin/
```

Then upload `manifest.xml` in Microsoft 365 Admin Center as a custom integrated app.

## Test Flow

1. Upload the add-in files to GitHub Pages.
2. Verify these URLs open:
   - `https://jorgecastelan-svg.github.io/tege-signature-assets/addin/runtime.html`
   - `https://jorgecastelan-svg.github.io/tege-signature-assets/addin/taskpane.html`
   - `https://jorgecastelan-svg.github.io/tege-signature-assets/addin/signatures.json`
3. In Microsoft 365 Admin Center, upload `manifest.xml`.
4. Assign the add-in to a small pilot group first.
5. Open New Outlook or Outlook Web and compose a new message.

## Admin Notes

Microsoft documents `OnNewMessageCompose` for event-based activation and `setSignatureAsync` for setting signatures in compose mode. This PoC intentionally uses hosted HTTPS image URLs because New Outlook does not reliably preserve local image files.
