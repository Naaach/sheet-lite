# Publishing & deployment

## Release to npm

Publishing is **manual and opt-in**: nothing publishes on push, so a green CI run can never release by accident.

### 1. Pre-publish checklist
- [ ] CI green on `main` (typecheck, tests, build, size gate, pack)
- [ ] Version bumped in `package.json` (and updated here/README if needed)
- [ ] `pnpm build && pnpm size && npm pack --dry-run` locally looks right
- [ ] `node scripts/check-publish.mjs` passes (identity, tarball allowlist, name availability, auth)

### 2. Dry run first (recommended)
From GitHub → **Actions → "Release to npm" → Run workflow**, leave `dry_run: true`.
It runs the same gates as CI (typecheck, test, build, size, pack) and then
`npm publish --dry-run` — it packages and validates everything without touching npm.

CLI equivalent:

```sh
gh workflow run release.yml -f dry_run=true
```

### 3. Publish for real
Run the same workflow with `dry_run: false` (uncheck the box, or `gh workflow run release.yml -f dry_run=false`).
It gates again, publishes with **provenance**, then creates the `v<version>` git tag and a GitHub Release.

### 4. Verify
```sh
npm view sheet-lite version dist.tarball
```
and open https://www.npmjs.com/package/sheet-lite.

## The npm token (read this before debugging a 403)

The repo secret `NPM_TOKEN` must be able to **create new packages** and to
**bypass 2FA** (npm requires 2FA for publishing since 2026). A token without
those gets:

```
npm error 403 Forbidden - PUT https://registry.npmjs.org/sheet-lite
You may not perform that action with these credentials.
```

Two valid options:

- **Classic token of type "Automation"** (simplest), or
- **Granular token** with *"Bypass two-factor authentication"* checked,
  Permissions = **"Read and write (publish and stage)"** (NOT stage only) and
  Select Packages = **"All Packages"** (specific packages cannot be created).

Auth is injected by `actions/setup-node` (`registry-url` + `NODE_AUTH_TOKEN`)
into a user-level `.npmrc` outside the repo — pnpm 12 ignores `${ENV}`
placeholders in the repository's own `.npmrc`.

### Local alternative
```sh
npm login            # or: npm config set //registry.npmjs.org/:_authToken <token>
npm publish --access public
```
Never commit `.npmrc`. Note: local publishes cannot produce provenance;
provenance only works from GitHub Actions.

Once the package exists on npm, you can migrate it to a **Trusted Publisher**
(Repo: `Naaach/sheet-lite`, workflow: `release.yml`) and drop the token entirely.

## Docs & playground on Cloudflare Pages

`.github/workflows/deploy-docs.yml` deploys on push to `main` (or manually):

| Project | Source | Output |
|---|---|---|
| `sheet-lite-docs` | `docs-site/` (prebuilt static site) | the documentation web |
| `sheet-lite-playground` | `playground/` → Vite build → `playground/dist` | the interactive demo |

Required repo secrets:

- `CLOUDFLARE_API_TOKEN` — Account → Cloudflare Pages → **Edit**
- `CLOUDFLARE_ACCOUNT_ID`

Until those secrets exist the deploy job **skips gracefully** (CI stays green).
`wrangler` runs through `pnpm dlx` (not the official action) because the action
installs wrangler with a bare `pnpm add`, which fails in pnpm-managed repos with
`ERR_PNPM_ADDING_TO_ROOT`. Project creation in the workflow is best-effort and
idempotent; the projects and any custom domain are managed in the Cloudflare dashboard.