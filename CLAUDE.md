# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

**Dunorte Truck — Central de Relatórios** is a self-contained, zero-build reporting portal for Dunorte Truck. It consists of two standalone HTML files with no dependencies to install and no build step.

- `index.html` — original version (includes a "Boletos" report section)
- `index_2.html` — updated version with minor UI refinements (spacing, font sizes, colors)
- `relatorio` — a file (no extension) containing report data used by the app

## How to run

Open either HTML file directly in a browser, or serve via any static file server:

```bash
# Python
python -m http.server 8080

# Node
npx serve .
```

There is no build, compile, or install step.

## Architecture

Both files are **single-file SPAs** — all CSS, JS, and report content live inside one HTML file.

### Authentication
Supabase Auth (email + password) is used. The Supabase client is initialized at the top of the `<script>` block with hardcoded `SUPABASE_URL` and `SUPABASE_KEY` (anon/public key). `sb.auth.onAuthStateChange` controls showing the login screen vs. the app.

### Report rendering pattern
All sub-reports are stored as **base64-encoded HTML strings** inside a JS object `const R = { cobranca: "...", atendimento: "...", ... }` embedded in the file. When a user navigates to a section, `showPage()` decodes the relevant entry with `b64d()` / `b64decode()` and renders it into a sandboxed `<iframe>` via `URL.createObjectURL(new Blob([...]))`. This means the entire app — including all sub-reports — ships as a single HTML file with no external report fetching.

### Navigation
`showPage(page, btn)` is the single routing function. It toggles between the home dashboard grid (`#home-content`) and the report iframe (`#report-frame`). An `iframe` → parent message channel (`postMessage` with type `'dn-update'`) lets embedded reports trigger URL changes in the frame.

### Pages
| Key | Label |
|---|---|
| `home` | Painel Inicial (dashboard grid) |
| `cobranca` | Cobrança (billing) |
| `atendimento` | Atendimento (customer service) |
| `resultados` | Resultados (monthly results) |
| `meta` | Meta e Bonificação (goals/bonuses) |
| `pmo` | Painel PMO (strategic management) |
| `boletos` | Relatório de Boletos (index.html only) |

## Report update workflow

When the user provides a new HTML report file, do the following automatically without asking for confirmation:

1. Base64-encode the new HTML content
2. Replace the corresponding value in `const R = {}` in `index.html` (and `index_2.html` if the key exists there too)
3. Commit with a message like `feat: atualiza relatório de cobrança`
4. Push to GitHub (`git push`)

To identify which key to replace, match the report content/title to one of: `cobranca`, `atendimento`, `resultados`, `meta`, `pmo`, `boletos`.

To base64-encode in PowerShell:
```powershell
$bytes = [System.IO.File]::ReadAllBytes("C:\path\to\relatorio.html")
$b64 = [Convert]::ToBase64String($bytes)
```

## Key development patterns

- To **add or update a report**: base64-encode the report HTML and replace the corresponding value in `const R = {}`.
- To **add a new section**: add a nav button calling `showPage('key', this)`, a home-grid card, and a `key` entry in `R`.
- CSS uses CSS custom properties defined in `:root` — always use those variables (`--navy`, `--yellow`, `--mist`, etc.) rather than hardcoded colors.
- `index_2.html` is the canonical version for new changes; `index.html` is kept for reference/backward compatibility.
