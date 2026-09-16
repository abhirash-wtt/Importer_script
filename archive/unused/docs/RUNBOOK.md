# DDO++ attendance sidecar — one-page runbook

## Start / stop
```
sudo systemctl start ddo-attendance-api
sudo systemctl stop ddo-attendance-api
sudo systemctl status ddo-attendance-api
```
Docker: `docker compose up -d` / `docker compose down` from the install directory.

## Logs
```
journalctl -u ddo-attendance-api -f
```
Office importer: `output\importer.log` and `output\last_run.json`.

## Health
`GET /health` must return 200 with `"store":"postgres"`. No token required.

## Rotate token
1. `node scripts/generate_token.js`
2. Set `DDO_API_TOKEN` on the server `.env` and every office PC `.env`
3. Restart the API
4. Confirm old token returns 401

## Import URL
Offices POST to `https://ddoplusnodeapi.walkingtree.tech/api/attendance/import`
Never to `/admin/vb0faior0eg`.

## Database
```
npm run check-db
npm run apply-schema
npm run backup-att
```
Restore: `psql "$DATABASE_URL" -f backups/<file>.sql` onto a copy database first.

## Rollback
1. Disable Task Scheduler `DDO-Attendance-Importer` on office PCs
2. `git checkout <previous-tag>` on the API host
3. Restart `ddo-attendance-api`

## HR failure path
Failed Excel files land in `failed\`. Empty inbox is success, not a failure.
