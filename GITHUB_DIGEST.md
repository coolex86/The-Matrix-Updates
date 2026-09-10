# GitHub Actions 每日 PDF 日报（方案 B）

仓库：https://github.com/coolex86/The-Matrix-Updates （公开仓只放脚本，密码在 Secrets）

每天 **UTC 01:00 = 马来西亚 09:00** 跑 `tools/send_usage_digest.py`（上一完整窗口）。  
发件人仍是 Gmail，附件仍是 PDF。电脑不用开机。

## 1. 把文件推进更新仓

需要在 `The-Matrix-Updates` 的 `main` 上有：

- `tools/send_usage_digest.py`
- `.github/workflows/daily-usage-digest.yml`

不要提交 `Json/usage_mail_config.json`、`sessions`、Gmail 密码。

## 2. 在 GitHub 填 Secrets

仓库 → **Settings** → **Secrets and variables** → **Actions** → **New repository secret**

| 名称 | 填什么 |
|------|--------|
| `DIGEST_SUPABASE_URL` | `Json/supabase_config.json` 的 `url` |
| `DIGEST_SUPABASE_ANON_KEY` | 同上 `anon_key` |
| `DIGEST_ADMIN_EMAIL` | 能登录 Matrix 的 **admin** 邮箱 |
| `DIGEST_ADMIN_PASSWORD` | 该账号 Matrix 登录密码 |
| `DIGEST_SMTP_USER` | `aerisfuryx@gmail.com` |
| `DIGEST_SMTP_PASSWORD` | 该 Gmail 的 16 位应用专用密码 |
| `DIGEST_FROM_EMAIL` | `aerisfuryx@gmail.com` |
| `DIGEST_TO_EMAILS` | 可选。逗号分隔；空则用 Supabase `admin_emails` |

收件人平时改 Supabase `usage_report_settings.admin_emails` 即可，不必改 Secrets。

## 3. 试跑

**Actions** → **daily-usage-digest** → **Run workflow**  
勾选 force，Day 可填 `2026-09-10`。成功后邮箱应收到 PDF。

定时任务不要勾 force：已发过的窗口会跳过。
