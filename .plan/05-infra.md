# 05 — Infrastructure (Root Config, CI/CD, Repo Rename)

## Root Monorepo Setup

### pnpm-workspace.yaml

```yaml
packages:
  - "packages/*"
```

### Root package.json

```json
{
  "name": "open-office",
  "private": true,
  "scripts": {
    "build": "pnpm -r build",
    "typecheck": "pnpm -r typecheck",
    "lint": "biome check packages/",
    "lint:fix": "biome check --write packages/",
    "format": "biome format --write packages/",
    "test": "pnpm -r test",
    "check": "pnpm typecheck && pnpm lint",
    "dev:excel": "pnpm --filter @open-office/excel dev-server",
    "dev:ppt": "pnpm --filter @open-office/powerpoint dev-server",
    "start:excel": "pnpm --filter @open-office/excel start",
    "start:ppt": "pnpm --filter @open-office/powerpoint start",
    "deploy:excel": "pnpm --filter @open-office/excel deploy",
    "deploy:ppt": "pnpm --filter @open-office/powerpoint deploy"
  },
  "devDependencies": {
    "@biomejs/biome": "^2.3.13",
    "@tailwindcss/postcss": "^4.1.18",
    "autoprefixer": "^10.4.24",
    "postcss": "^8.5.6",
    "tailwindcss": "^4.1.18",
    "typescript": "^5.4.2"
  },
  "pnpm": {
    "onlyBuiltDependencies": ["esbuild"]
  }
}
```

### tsconfig.base.json

```json
{
  "compilerOptions": {
    "target": "ES2020",
    "module": "ESNext",
    "moduleResolution": "bundler",
    "jsx": "react-jsx",
    "strict": true,
    "esModuleInterop": true,
    "skipLibCheck": true,
    "forceConsistentCasingInFileNames": true,
    "resolveJsonModule": true,
    "isolatedModules": true,
    "noEmit": true,
    "declaration": true,
    "declarationMap": true,
    "sourceMap": true,
    "lib": ["ES2020", "DOM", "DOM.Iterable"]
  }
}
```

Each package's tsconfig.json extends this:
```json
{
  "extends": "../../tsconfig.base.json",
  "compilerOptions": {
    "rootDir": "src",
    "outDir": "dist"
  },
  "include": ["src"],
  "references": [
    { "path": "../shared" }
  ]
}
```

### biome.json

Move current `biome.json` to root. Applies to all packages.

## CI Workflows

### ci.yml (all pushes + PRs)

```yaml
name: CI

on:
  push:
    branches: [main]
  pull_request:
    branches: [main]

jobs:
  quality:
    name: Quality Checks
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v6
      - uses: pnpm/action-setup@v4
        with:
          version: 9
      - uses: actions/setup-node@v6
        with:
          node-version: 20
          cache: pnpm
      - run: pnpm install --frozen-lockfile
      - run: pnpm typecheck
      - run: pnpm lint
      - run: pnpm build
```

### release-excel.yml (tag: excel-v*)

```yaml
name: Release Excel

on:
  push:
    tags:
      - "excel-v*"

permissions:
  contents: write

jobs:
  release:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v6
      - uses: pnpm/action-setup@v4
        with:
          version: 9
      - uses: actions/setup-node@v6
        with:
          node-version: 20
          cache: pnpm
      - run: pnpm install --frozen-lockfile
      - run: pnpm typecheck
      - run: pnpm lint

      - name: Build Excel
        run: pnpm --filter @open-office/excel build

      - name: Extract changelog
        id: changelog
        run: |
          VERSION="${GITHUB_REF_NAME#excel-v}"
          NOTES=$(awk "/^## \\[${VERSION}\\]/{found=1; next} /^## \\[/{if(found) exit} found{print}" packages/excel/CHANGELOG.md)
          echo "$NOTES" > /tmp/release-notes.md

      - name: Deploy to Cloudflare Pages
        uses: cloudflare/wrangler-action@v3
        with:
          apiToken: ${{ secrets.CLOUDFLARE_API_TOKEN }}
          accountId: ${{ secrets.CLOUDFLARE_ACCOUNT_ID }}
          command: pages deploy packages/excel/dist --project-name=openexcel --branch=main --commit-dirty=true

      - name: Create GitHub Release
        uses: softprops/action-gh-release@v2
        with:
          body_path: /tmp/release-notes.md
```

### release-ppt.yml (tag: ppt-v*)

Same structure, but:
- Tag pattern: `ppt-v*`
- Build filter: `@open-office/powerpoint`
- Changelog: `packages/powerpoint/CHANGELOG.md`
- Cloudflare project: `openppt`
- Deploy path: `packages/powerpoint/dist`

## Repo Rename

### Steps

1. **GitHub**: Settings → General → Repository name → `open-office`
2. **Git remote** (GitHub auto-redirects, but update anyway):
   ```bash
   git remote set-url origin git@github.com:hewliyang/open-office.git
   ```
3. **README.md**: Update title, description, URLs
4. **AGENTS.md**: Update project overview, structure, references
5. **manifest.prod.xml** (Excel): Update `SupportUrl` and `LearnMoreUrl` to new repo URL
6. **package.json** names:
   - Root: `"name": "open-office"`
   - Excel: `"name": "@open-office/excel"`
   - PPT: `"name": "@open-office/powerpoint"`
   - Shared: `"name": "@open-office/shared"`

### What DOESN'T change

- Cloudflare Pages project `openexcel` — stays (it's a deployment target, not user-facing)
- localStorage keys `openexcel-*` — stays (changing would log out existing users)
- IndexedDB name `OpenExcelDB_v3` — stays (changing would lose existing sessions)
- Manifest display name `OpenExcel` — stays (it's the add-in brand name)

## Cloudflare Pages

### Existing: `openexcel`
- Already deployed at `openexcel.pages.dev`
- Deploy command changes to: `pnpm --filter @open-office/excel build && wrangler pages deploy packages/excel/dist --project-name=openexcel`

### New: `openppt`
- Create new Pages project on Cloudflare dashboard
- Will be at `openppt.pages.dev`
- Deploy: `pnpm --filter @open-office/powerpoint build && wrangler pages deploy packages/powerpoint/dist --project-name=openppt`

## Tag Strategy

Instead of `v0.2.1` (ambiguous in monorepo), use prefixed tags:
- `excel-v0.2.2` → triggers `release-excel.yml`
- `ppt-v0.1.0` → triggers `release-ppt.yml`

Each package has its own CHANGELOG.md and independent versioning.

## Migration of Existing Git History

No need to rewrite history. The monorepo restructure is just a commit that moves files. Old tags (`v0.1.0` through `v0.2.1`) remain valid — they point to commits before the restructure.

## Checklist

- [ ] Create root `package.json` with workspace scripts
- [ ] Create `pnpm-workspace.yaml`
- [ ] Create `tsconfig.base.json`
- [ ] Move `biome.json` to root (adjust paths if needed)
- [ ] Create `.github/workflows/ci.yml` (replace current)
- [ ] Create `.github/workflows/release-excel.yml`
- [ ] Create `.github/workflows/release-ppt.yml`
- [ ] Remove old `.github/workflows/release.yml`
- [ ] Rename repo on GitHub
- [ ] Update git remote
- [ ] Update README.md
- [ ] Update AGENTS.md
- [ ] Update manifest URLs (repo links only)
- [ ] Create Cloudflare Pages project `openppt`
- [ ] Verify `pnpm install` works from root
- [ ] Verify `pnpm build` builds all packages
- [ ] Verify `pnpm dev:excel` starts Excel dev server
