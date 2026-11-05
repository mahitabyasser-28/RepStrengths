# =====================================================================
# Runs rep_strengths.py (your Colab code, UNCHANGED) on GitHub Actions.
# Colab signs in with your Google account; GitHub can't, so this file makes
# the Colab sign-in calls use the Google service account instead
# (secret GOOGLE_SERVICE_ACCOUNT_JSON). Then it runs your script as-is.
# =====================================================================
import json, os, sys, types, runpy

key = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON", "").strip()
if not key:
    sys.exit("❌ GOOGLE_SERVICE_ACCOUNT_JSON secret is empty — add it under Settings → Secrets and variables → Actions")
info = json.loads(key)
print(f"🔑 Google service account: {info.get('client_email')}")
print("   → it needs VIEWER on the master sheet and EDITOR on the working (target) sheet")

from google.oauth2.service_account import Credentials
creds = Credentials.from_service_account_info(info, scopes=[
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.readonly",
])

import google.auth
google.auth.default = lambda *a, **k: (creds, info.get("project_id"))     # what the Colab code calls

# stand-ins for google.colab (only exists inside Colab)
colab = types.ModuleType("google.colab")
auth = types.ModuleType("google.colab.auth"); auth.authenticate_user = lambda *a, **k: None
files = types.ModuleType("google.colab.files"); files.download = lambda *a, **k: None
colab.auth, colab.files = auth, files
sys.modules.update({"google.colab": colab, "google.colab.auth": auth, "google.colab.files": files})
import google; google.colab = colab

runpy.run_path("rep_strengths.py", run_name="__main__")
