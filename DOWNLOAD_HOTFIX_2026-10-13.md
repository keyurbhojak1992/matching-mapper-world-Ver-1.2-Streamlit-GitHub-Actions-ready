# Download delivery hotfix — 2026-10-13

## Diagnosis

The uploaded server log repeatedly shows Uvicorn rejecting an HTTP body before
response headers (`Expected ASGI message 'http.response.start', but got
'http.response.body'`). The trace is entirely in Starlette's static-file
response, Streamlit's GZip middleware, and Uvicorn; it does not show a failure
in either Demand Filter matching algorithm. The regular Demand Filter page
reported a completed ZIP of about 0.03 MB in the screenshot, so a broken
HTTP response can explain why Chrome says the site was unavailable.

The stack installed Streamlit 1.59.2 and Starlette 1.8.0. Streamlit 1.62.0
incorporates the native Starlette GZip/media response refactor (release notes
refer to streamlit/streamlit PRs 16344 and 16462). This update chooses
Streamlit 1.62.0 plus Starlette 1.5.0 for that implementation. This is a
compatibility mitigation; the supplied log does not prove the exact middleware
bug triggering the invalid ASGI response order.

## Changed files

- `requirements.txt`: move from Streamlit 1.59.2 to 1.62.0, and explicitly use
  Starlette 1.5.0 to avoid the previous dependency combination.
- `app.py`: retain normal Streamlit download buttons. Offer a small-result HTML
  data-URI download as a bypass if the Streamlit `/media` HTTP endpoint is
  unavailable (2 MiB cap). Enable `on_click="ignore"` on batch-result downloads.
- `ui/download_fallback.py`: pure helper for safely encoding small downloads.
- `tests/test_download_fallback.py`: validates byte integrity, escaping, and cap.

All matching algorithms, workbook writers, and batch checkpoint processors are
unchanged. The server must REDEPLOY / rebuild its environment for the newer
Streamlit dependency to take effect. Browser reload alone does not upgrade it.

## Verification

Run `python -m pytest -q` and verify a small regular Demand Filter ZIP via both
the default button and alternative link. Verify the large Advanced Demand batch
ZIP with the standard download button. Check the Streamlit Cloud log to confirm
it installed Streamlit 1.62.0 and Starlette 1.5.0, and that the ASGI exception
is gone. If it persists, capture the request path/response headers from Chrome
DevTools Network tab and the new server trace. If the replacement Cloud container
was provisioned, old advanced batch job checkpoint IDs may no longer be valid.

This patch was unit-tested without a live Streamlit Community Cloud server;
remote/browser download success must be verified after deploying.
