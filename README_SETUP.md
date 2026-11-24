# KMS WORD AI V20 PRO — Online activation + Windows installer

## What is included
- Updated desktop app with an online activation gate.
- FastAPI licensing server with activation, validation, key generation and revocation.
- Admin key-generation utility.
- PyInstaller build script and Inno Setup installer definition.

## 1. Deploy the licensing server
Use a Linux VPS or a managed Python host with HTTPS. Do not expose the server over plain HTTP in production.

```bash
cd server
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
export KMS_ADMIN_KEY="$(python -c 'import secrets; print(secrets.token_urlsafe(32))')"
export KMS_SIGNING_SECRET="$(python -c 'import secrets; print(secrets.token_urlsafe(48))')"
uvicorn main:app --host 0.0.0.0 --port 8000
```

Keep both secrets in your hosting provider's secret/environment-variable settings. Do not put them in the desktop app or share them with customers. Put a trusted HTTPS reverse proxy/domain in front of port 8000. Persist `licenses.sqlite3` on durable storage. For production sales, configure backups, monitoring, rate limiting and a managed database.

## 2. Configure the desktop client
Edit the root `kms_config.json` and replace the example URL with your actual HTTPS API base URL, for example:
`https://licenses.yourdomain.com`

The app checks the server each time it starts. It will not open the editor if the server is unavailable. This is deliberate for strict online-only activation.

## 3. Generate customer product keys
Run from the project folder:

```bash
python tools/generate_keys.py --url https://licenses.yourdomain.com --admin-key YOUR_ADMIN_KEY --count 20 --devices 1
```

Optional expiry:
`--days 365`

Keep the returned keys private. The server stores only SHA-256 hashes of keys.

## 4. Build on Windows
On a Windows 10/11 machine with Python 3.11 or newer:

1. Install Python and enable "Add Python to PATH".
2. Open this folder and double-click `installer\build_windows.bat`.
3. Install Inno Setup 6.
4. Open `installer\KMS_Word_AI_Setup.iss` in Inno Setup Compiler and click Compile.
5. The installer will be created in the `release` folder.

The installer output is not prebuilt here because PyInstaller and Inno Setup need to run in a Windows build environment. The project includes the scripts to create it there.

## Important security notes
- Use HTTPS only for the live service.
- Keep the admin key and signing secret server-side.
- The current device binding uses the Windows MachineGuid, hashed before transmission. Reinstalling Windows or changing the motherboard may require a licence reset.
- The current service supports one or more activations per key and key revocation. It does not yet include a customer self-service portal or automated payment integration.
- Have the server reviewed and hardened before commercial distribution.


## Easier hosted deployment
See `DEPLOY_ON_RENDER.md` and the root `render.yaml` for a Render Blueprint deployment using a persistent disk.
