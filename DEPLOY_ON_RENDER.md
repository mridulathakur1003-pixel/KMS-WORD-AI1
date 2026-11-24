# Deploy the KMS WORD AI activation server on Render

This package is prepared for a Render Docker web service with a persistent disk for the licence database.

## Deploy

1. Create a private Git repository and upload the contents of this package to it.
2. Sign in to Render and choose **New → Blueprint**.
3. Connect the repository and select the `render.yaml` file at the repository root.
4. Review the service and persistent-disk plan/cost before confirming deployment.
5. Render will generate `KMS_ADMIN_KEY` and `KMS_SIGNING_SECRET` as secret environment variables. Keep them private.
6. After deployment, copy the service's HTTPS URL. Check that opening `<your-service-url>/health` returns `{"status":"ok","service":"KMS WORD AI licensing"}`.

## Connect the desktop app

Edit `kms_config.json` in the project root:

```json
{
  "api_base_url": "https://YOUR-REAL-SERVICE.onrender.com"
}
```

Use the exact HTTPS URL Render assigns to your service. Then rebuild the Windows app and installer. Do not distribute an app build that still contains the placeholder URL.

## Generate product keys

In Render, open the service's **Environment** page and copy `KMS_ADMIN_KEY` privately. From your own computer, run:

```bash
python tools/generate_keys.py --url https://YOUR-REAL-SERVICE.onrender.com --admin-key "YOUR_PRIVATE_ADMIN_KEY" --count 10 --devices 1
```

Save the returned keys securely and give one key to each customer. Never put the admin key or signing secret into the desktop application or customer installer.

## Important

- This is an online-only activation design. The desktop app contacts the server at each launch to check that the licence is still active.
- Keep the persistent disk attached. Deleting it deletes the licence database.
- Review current hosting charges and availability before deploying; this template uses a paid persistent-disk-capable service configuration.
- Before commercial release, add a custom domain, backups, monitoring, rate limiting, and a support process for legitimate device changes.
- I cannot create a Render account, accept hosting terms, connect your Git provider, or deploy into your account from this chat. Those steps require your own account authorization.
