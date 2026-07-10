# TeGe Signature Assets

Static GitHub Pages package for TeGe Outlook signature images.

## Files

```text
index.html
assets/
  logo/
    tege-logo.png
  icons/
    linkedin.png
    booking.png
  seals/
    business-innovator.png
```

## AssetBaseUrl

If the repository is named `tege-signature-assets`, the V3 asset base URL is:

```text
https://jorgecastelan-svg.github.io/tege-signature-assets/assets/
```

Generate production signatures with:

```powershell
pwsh ./scripts/generate-signatures.ps1 -AssetBaseUrl "https://jorgecastelan-svg.github.io/tege-signature-assets/assets/"
```

Direct image URLs:

```text
https://jorgecastelan-svg.github.io/tege-signature-assets/assets/logo/tege-logo.png
https://jorgecastelan-svg.github.io/tege-signature-assets/assets/icons/linkedin.png
https://jorgecastelan-svg.github.io/tege-signature-assets/assets/icons/booking.png
https://jorgecastelan-svg.github.io/tege-signature-assets/assets/seals/business-innovator.png
```

## GitHub Setup

1. Create or open the public GitHub repository `tege-signature-assets`.
2. Upload everything from this folder to the repository root.
3. In GitHub, open `Settings > Pages`.
4. Set source to `Deploy from a branch`.
5. Select branch `main` and folder `/root`.
6. Save and wait until GitHub shows the Pages URL.
