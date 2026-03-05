# ASP4 — Classic ASP/VBScript Runtime for Python

**Run your legacy Classic ASP pages on modern infrastructure — no IIS required.**

ASP4 is a Python-based runtime that executes Classic ASP (VBScript) pages on Windows, Linux, and macOS. It implements the full Classic ASP object model (`Request`, `Response`, `Session`, `Application`, `Server`) alongside broad VBScript built-in coverage, so most legacy ASP applications just work.

---

## Why ASP4?

Classic ASP applications represent decades of business logic. Rewriting them is expensive and risky. ASP4 lets you **keep your existing `.asp` files** and serve them through a lightweight Python HTTP server — no COM, no Windows dependency, no IIS license. Linux typically runs Python 10–30% faster than Windows, increasing the performance advantage of modern frameworks like ASP4 over Classic ASP/VBScript on IIS.

---

## Quick Start

```bash
pip install fpdf2 bcrypt pillow pyodbc
python -m ASP4.server 0.0.0.0 8080 www
```

Point your browser at `http://localhost:8080` and your `.asp` pages are live.

---

## What's Supported

| Area | Coverage |
|------|----------|
| VBScript built-ins (strings, dates, math, arrays) | Near-complete |
| `Request` / `Response` / `Session` / `Application` / `Server` | Near-complete |
| `Scripting.Dictionary` / `Scripting.FileSystemObject` | Supported |
| `ADODB.Connection` / `Recordset` / `Command` / `Stream` | Partial |
| Database backends | SQLite, Access, Excel (read-only), ODBC, PostgreSQL |
| `VBScript.RegExp` | Supported |
| MSXML HTTP & DOM | Partial (security-sandboxed) |
| `CDO.Message` (SMTP) | Partial |
| POP3 / IMAP | Partial |
| `Global.asa` events | Supported |

ASP4 also ships extended helpers beyond classic ASP — JSON encode/decode, ZIP, PDF generation, image processing, and bcrypt password hashing — all accessible from VBScript via the global `ASP4` object.

---

## Requirements

- Python 3.8+
- `fpdf2`, `bcrypt`, `pillow`, `pyodbc` (install as needed for the features you use)

---

## Platform Support

Windows · Linux · macOS

---

## Compatibility Notes

ASP4 targets practical app-level compatibility, not byte-for-byte IIS parity. Locale-specific formatting, edge-case type coercion, and COM-level quirks may differ. SQL is executed as-is by the underlying driver — no dialect translation is performed.

If you're migrating a critical application, run your own regression tests against ASP4 alongside IIS before cutting over.

---

## License

See [LICENSE](LICENSE) for details.
