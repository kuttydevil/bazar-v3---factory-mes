<div align="center">
<img width="1200" height="475" alt="GHBanner" src="https://ai.google.dev/static/site-assets/images/share-ais-513315318.png" />
</div>

# Run and deploy your AI Studio app

This contains everything you need to run your app locally.

View your app in AI Studio: https://ai.studio/apps/1f7d1490-5d56-451c-b2e6-e4fe0b82fd0b

## Run Locally

**Prerequisites:**  Node.js


1. Install dependencies:
   `npm install`
2. Set the `GEMINI_API_KEY` in [.env.local](.env.local) to your Gemini API key
3. Run the app:
   `npm run dev`

## Villa Set Readiness

The Villa Set Readiness tab estimates complete sets from parts labeled `STANDARD 40`, excluding packing and skirting parts. A complete set requires all four groups: PDF 2 (Bedroom 03), PDF 3 (Bedroom 02), PDF 5 (Bedroom 01 & Drawers), and PDF 1 + PDF 4 (Side Cabinet + Master Bedroom). The ready-set estimate is limited by the least-complete required group. Quantities are divided into 40 sets only when they divide evenly. Non-even quantities block the verified ready count rather than being rounded or silently excluded; update the source cutlist with their exact villa allocations before relying on a verified count. The cutlist does not map shared sets to individual Villa IDs or a single responsible team, so the view reports set readiness and assigned-team cut contributions, not individual villa or worker ownership.
