# Publishing `sheet-lite` to npm

How to ship a new version of the package to [npmjs.com/package/sheet-lite](https://www.npmjs.com/package/sheet-lite).

Two supported paths:

| Path | When | How |
|---|---|---|
| **CI (recommended)** | always | `.github/workflows/publish.yml` (`NPM_TOKEN` secret) — runs typecheck/test/build, publishes **with provenance** |
| **Local** | fallback only | `npm login` / `_authToken`, then `npm publish --access public` |

> Rule of thumb: **never publish what does not pass typecheck + tests.** Both paths below run the full gate.

---

## 0. Pre-publish checklist

Run from the repo root (PowerShell):

```powershell
pnpm install            # dependencies in sync
pnpm typecheck          # tsc --noEmit — must be clean
pnpm test               # must be green (35/35 as of 0.3.0)
pnpm build              # tsdown -> dist/ (mjs, cjs, .d.mts, .d.cts + sourcemaps)
node scripts/check-publish.mjs   # tarball allowlist + registry availability + auth
npm pack --dry-run      # eyeball the contents one last time
```

`scripts/check-publish.mjs` verifies — without uploading anything — that:

- the tarball contains **exactly** `dist/` (6 files), `README.md`, `LICENSE`, `package.json` (9 entries, ~78 kB packed / ~290 kB unpacked for 0.3.0), and
- **nothing else**: no `src/`, `docs/`, `docs-site/`, `playground/`, `tests/`, `scripts/`, `reports/` (this is enforced by the `files` allowlist in `package.json` and re-checked by the script), and
- the exact `name@version` is still free on the registry (`npm view sheet-lite@0.3.0 version` → `E404` = free), and
- you are authenticated (`npm whoami`).

Also:

- [ ] Bump `version` in `package.json` **before** publishing (see §4) — npm refuses to overwrite an existing version.
- [ ] Commit and tag the release (`vX.Y.Z`) before/with the publish so the tree matches the tarball.
- [ ] `pnpm test` count matches the milestone in `docs/STATUS.md`.

---

## 1. Recommended: publish from GitHub Actions

`​.github/workflows/publish.yml` triggers on:

- a GitHub **Release published**, or
- a manual **workflow_dispatch**: *Actions → “Publish to npm” → Run workflow*.

It runs `pnpm install --frozen-lockfile` → `pnpm typecheck` → `pnpm test` → `pnpm build`, then
`npm publish --provenance --access public` using the `NPM_TOKEN` secret (`NODE_AUTH_TOKEN`).
Provenance works there because the job has `id-token: write` (Actions OIDC). **Provenance cannot work from a local machine** — that is expected; don't force it locally (§2).

### One-time setup — store the token

Create an npm **automation** token on npmjs.com (*Access Tokens → Generate New Token*; classic “Automation” or a granular token with publish rights to `sheet-lite`), then store it as a repository secret named **exactly** `NPM_TOKEN`:

```powershell
gh secret set NPM_TOKEN
# paste the token when prompted, then press Enter
```

(or via the web UI: *Settings → Secrets and variables → Actions → New repository secret*).

Token hygiene:

- Never paste a token into files, commits, issues, or chat.
- Never commit `.npmrc` — tokens must only live in `gh secret set` / your user-level `~/.npmrc`.
- Revoke and rotate the token if it is ever exposed.

### Publish

1. Bump the version and push a tag `vX.Y.Z` (or just run the workflow — it publishes whatever `version` is on the default branch).
2. Create the GitHub Release (or run the workflow manually).
3. Watch the workflow run turn green — typecheck/test/build must pass first, otherwise nothing is published.

---

## 2. Fallback: publish from a local machine

Authenticate first — either:

```powershell
npm login
```

or set a token in your **user-level** npm config (default location for `npm config set` is `~/.npmrc`, *not* the repo):

```powershell
npm config set //registry.npmjs.org/:_authToken <YOUR_TOKEN>
```

> ⚠️ **Never commit `.npmrc`.** Verify with `git status --porcelain .npmrc` (should be empty) and never pass `--location=project` for auth tokens. Remove a stored token with `npm config delete //registry.npmjs.org/:_authToken`.

Then:

```powershell
npm publish --access public
```

Notes:

- `package.json` sets `publishConfig.provenance: true` for CI. A local publish cannot produce provenance (no OIDC); if npm errors mentioning provenance/OIDC, rerun with `npm publish --access public --no-provenance`. Do **not** remove the `publishConfig` entry — the CI workflow relies on it.
- If your account enforces 2FA, npm will prompt for an OTP. An *automation*-type token skips the OTP prompt (that's the point of the type).
- A local publish is not blocked by `prepublishOnly` — it is the gate: it runs `pnpm build && pnpm test` and aborts the publish if either fails.

---

## 3. Verify a published release

```powershell
npm view sheet-lite version          # should print the new version
npm view sheet-lite dist.tarball     # tarball URL on registry.npmjs.org
npm view sheet-lite name version description repository.url
```

Inspect the actual published tarball:

```powershell
npm pack sheet-lite@0.3.0            # downloads sheet-lite-0.3.0.tgz
tar -tzf sheet-lite-0.3.0.tgz        # must list only package/{LICENSE,README.md,package.json,dist/…}
```

Smoke-test the install in a scratch directory:

```powershell
npm i sheet-lite
node -e "import('sheet-lite').then(m => console.log(Object.keys(m)))"   # ESM import works
node -e "console.log(Object.keys(require('sheet-lite')))"               # CJS require works
```

For CI-published versions, npmjs.com shows a **Provenance** badge; you can also verify attestations locally with `npm audit signatures` inside a project that installs the package.

---

## 4. Name conflicts / version bumps

- `npm view sheet-lite version` → `E404` = the name is free; proceed.
- Publish fails with “cannot publish over the previously published version” / `EPUBLISHCONFLICT` → that **version** exists; bump.
- `npm view sheet-lite` shows **someone else's package** → **stop**. Do not silently bump and retry; report it and decide the package name/path with the maintainer.
- Bump style (semver): fixes → `npm version patch`, new features → `npm version minor` (0.x: minor = feature release, e.g. `0.4.0` runtime, `0.5.0` stream per `docs/STATUS.md`).

---

## 5. Troubleshooting

| Error | Meaning | Fix |
|---|---|---|
| `ENEEDAUTH` | not logged in on this machine | §2 (`npm login` / token) or publish via CI (§1) |
| `E409` / “cannot publish over previously published version” | version already on the registry | bump the version (§4) |
| provenance/OIDC error locally | provenance needs CI OIDC | `npm publish --access public --no-provenance` locally; keep provenance for CI |
| `E403` forbidden | token lacks publish rights (or package blocked) | regenerate the token with publish access to `sheet-lite` |
| tarball contains unexpected files | something outside the `files` allowlist | fix `files` in `package.json`, rerun `node scripts/check-publish.mjs` |
| tests/typecheck fail in the workflow | gate did its job | fix, merge, rerun the workflow — nothing was published |

---

## Status

- **2026-10-14** — `sheet-lite@0.3.0` verified locally: `pnpm typecheck` clean, `pnpm test` **35/35**, `pnpm build` ok, tarball exact (9 files, ~78 kB), name free (`E404`). Publish blocked on this machine: **not authenticated** (`npm whoami` → `ENEEDAUTH`). Next step: `gh secret set NPM_TOKEN`, then publish via the **Publish to npm** workflow (§1) — or `npm login` and `npm publish --access public` locally (§2).
