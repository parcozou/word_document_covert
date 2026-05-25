# Markdown Report to DOCX Plugin for Coze

This plugin turns the combined report Markdown received by the Coze document node into a downloadable `.docx`. Unlike a plain text exporter, it parses Markdown table elements and writes editable native Word tables. It can also embed permitted ECharts PNG links into the report.

## Recommended Architecture

Coze calls a public HTTPS API imported as a custom plugin through `openapi.yaml`:

1. The LLM report node passes `formatted_markdown`, `title`, and `to_format`.
2. `POST /generate-docx` converts headings, paragraphs, lists, hyperlinks, chart images, and Markdown tables into DOCX content.
3. The service stores the file in object storage for production, or exposes a local file URL for initial testing.
4. The plugin returns `download_url`, which the Coze final output node can display as the Word report download link.

A file link must resolve to a persisted binary document. For this reason, production deployment should use S3-compatible storage such as Cloudflare R2 or AWS S3, rather than relying on a temporary code-node filesystem.

## Input And Output

The input matches the supplied `Untitled-1.md` payload:

```json
{
  "formatted_markdown": "# CDMO Industry and Company Analysis Report: ...",
  "title": "Corporate Valuation Report for Asymchem 002821",
  "to_format": "docx",
  "include_images": true
}
```

The API returns:

```json
{
  "download_url": "https://files.example.com/generated-docx/Corporate_Valuation_Report_for_Asymchem_002821_xxxxx.docx",
  "file_name": "Corporate_Valuation_Report_for_Asymchem_002821_xxxxx.docx",
  "table_count": 14,
  "image_count": 12,
  "skipped_image_count": 0,
  "generated_at_utc": "2026-05-25T00:00:00+00:00",
  "message": "DOCX report generated with editable Word tables."
}
```

`table_count` and `skipped_image_count` should be kept in the workflow output while testing. They show whether tables were converted and whether any chart failed to embed, which avoids silently publishing an incomplete report.

## Formatting Behaviour

- Markdown pipe tables become editable Word tables with a shaded header row and fixed column geometry.
- LLM output is normalised before conversion so a heading or paragraph immediately after a table is not incorrectly inserted as an additional table row.
- Section dividers written as `---` are separated from preceding narrative paragraphs so Markdown does not interpret analytical prose as a level-two heading.
- Headings and body copy use a restrained business-report style suitable for equity research output.
- URLs in Markdown links remain clickable in Word.
- ECharts images are embedded only from HTTPS hosts listed in `IMAGE_HOST_ALLOWLIST`; the default accepts the Coze chart-resource host shown in the sample payload.
- Accepted chart images are decoded and re-encoded as clean PNG files before insertion, reducing image-encoding compatibility problems.
- Missing or rejected images are labelled as unavailable in the generated file rather than being treated as evidence by the report generator.

## Deploy The API

### Option A: Initial Test With Local File Serving

Install and run:

```bash
pip install -r requirements.txt
set PUBLIC_BASE_URL=https://your-public-api-domain.example.com
set DOCX_API_KEY=replace-with-a-long-random-token
uvicorn app:app --host 0.0.0.0 --port 8000
```

Use `STORAGE_MODE=local` only where the deployed service retains its `generated_files` directory and where its public URL serves `/files/...`.

### Option B: Production With Object Storage

Set `STORAGE_MODE=s3` and configure the S3-compatible values shown in `.env.example`. With a private bucket, omit `S3_PUBLIC_BASE_URL` and the plugin returns an expiring signed download link. With a public/custom-domain bucket, set `S3_PUBLIC_BASE_URL` for stable URLs.

The included `Dockerfile` can be deployed on a container host such as Render, Railway, or Cloud Run. Store API and object-storage credentials as environment secrets.

## Configure Coze

1. Deploy this API to a public HTTPS domain.
2. In `openapi.yaml`, replace `https://YOUR-DOCX-PLUGIN-DOMAIN.example.com` with that domain.
3. In Coze, open **Library** > **+ Resource** > **Plugin** > **Import**, upload `openapi.yaml`, and set service API-key authentication in header `X-API-Key` to the same secret as `DOCX_API_KEY`.
4. Connect the final report Markdown output to `formatted_markdown`, the report title to `title`, and set `to_format` to `docx`.
5. Enable and debug `generateDocxReport`, then publish the plugin before adding it to a Coze workflow.
6. Return `download_url` to the user. During testing, also expose `table_count`, `image_count`, and `skipped_image_count`.

Full setup instructions are in [Coze Deployment Guide.md](./Coze%20Deployment%20Guide.md). The applicable official Coze documentation is [Create a plugin by importing a JSON or YAML file](https://www.coze.com/open/docs/guides/plugin_import) and [Plugin node](https://www.coze.com/open/docs/guides/plugin_node). The local-plugin callback guide is a different API integration route and is not required for this hosted DOCX conversion plugin.

## Local Verification With The Supplied Payload

After dependencies are installed, generate a quick table-only audit:

```bash
python test_from_payload.py --skip-images
```

Then run without `--skip-images` to check chart retrieval:

```bash
python test_from_payload.py
```

Open the generated file under `generated_files/` and verify the report tables are editable. The test fails if no native Word tables are created.

## Operational Safeguards

- Use `DOCX_API_KEY` and do not expose the endpoint without authentication.
- Limit allowed image hosts to trusted sources to prevent arbitrary remote fetches.
- Use signed URLs or bucket lifecycle rules so generated reports are not retained longer than necessary.
- The conversion plugin formats submitted content; it does not validate the investment analysis or the reliability of underlying source claims.
