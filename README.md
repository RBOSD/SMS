# SMS v2（Render 部署版）

這個 repo 是 SMS 系統的 v2 重寫版本（Next.js + NestJS + PostgreSQL + Prisma）。
已放棄 v1 資料結構，直接採用 v2 正規化後 schema（`Issue` + `IssueRound`）。

## Render 部署（Blueprint）
本 repo 已包含 `render.yaml`，可直接用 Render Blueprint 建立：
- `sms-v2-db`：Postgres
- `sms-v2-api`：NestJS API（會在啟動時自動 `prisma migrate deploy`）
- `sms-v2-web`：Next.js Web（透過 Render private network 反向代理到 API）

### 1) 先把 `D:\\sms-v2` 推到 GitHub
Render 只能從 Git repo 部署，所以請先把此資料夾初始化 git、推到 GitHub。

### 2) Render Dashboard
- `New` → `Blueprint`
- 選你的 GitHub repo
- Render 會讀取 `render.yaml` 自動建立 2 個 web service + 1 個 Postgres

### 3) 需要你在 Render 上填的變數
`sms-v2-api`：
- **GEMINI_API_KEY**：要啟用 AI 審查才需要（可先留空）

> 其他像 `JWT_SECRET`、`DEFAULT_ADMIN_PASSWORD` 會由 Blueprint 自動產生。

## 資料庫（v2 正規化）
Prisma schema：`apps/api/prisma/schema.prisma`

已建立初始 migration：`apps/api/prisma/migrations/20260202000000_init/migration.sql`

Render 上的 API 會在啟動時執行：
- `prisma migrate deploy`

## 本機（可選）
你有 PostgreSQL 的話：
1. 設定 `apps/api/.env` 的 `DATABASE_URL`
2. 安裝：`npm install`
3. 建表：`npm run db:migrate -w @sms/api`
4. 啟動：`npm run dev:api`、`npm run dev:web`

