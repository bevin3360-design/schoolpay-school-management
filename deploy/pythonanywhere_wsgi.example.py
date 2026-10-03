import os
import sys


username = "REPLACE_WITH_PYTHONANYWHERE_USERNAME"
project_dir = f"/home/{username}/schoolpay-school-management"
data_dir = f"/home/{username}/.local/share/schoolpay-demo"
secret_key = "REPLACE_WITH_A_RANDOM_SECRET_KEY"
admin_username = "REPLACE_WITH_DEMO_ADMIN_USERNAME"
initial_admin_password = "REPLACE_WITH_A_UNIQUE_PASSWORD_OF_AT_LEAST_12_CHARACTERS"

if any("REPLACE_WITH" in value for value in (
    username,
    secret_key,
    admin_username,
)):
    raise RuntimeError("Replace all deployment placeholders before reloading.")
if initial_admin_password.startswith("REPLACE_WITH"):
    raise RuntimeError("Set an initial admin password before the first reload.")
if initial_admin_password and len(initial_admin_password) < 12:
    raise RuntimeError("The initial admin password must be at least 12 characters.")

os.makedirs(os.path.join(data_dir, "uploads"), exist_ok=True)
os.environ.update(
    {
        "APP_ENV": "production",
        "SECRET_KEY": secret_key,
        "DATABASE_PATH": os.path.join(data_dir, "school.db"),
        "UPLOAD_FOLDER": os.path.join(data_dir, "uploads"),
        "INITIAL_SCHOOL_NAME": "SchoolPay Demo",
        "INITIAL_ADMIN_USERNAME": admin_username,
        "INITIAL_ADMIN_FULL_NAME": "Demo Administrator",
    }
)
if initial_admin_password:
    os.environ["INITIAL_ADMIN_PASSWORD"] = initial_admin_password
else:
    os.environ.pop("INITIAL_ADMIN_PASSWORD", None)
    os.environ.pop("INITIAL_ADMIN_USERNAME", None)

if project_dir not in sys.path:
    sys.path.insert(0, project_dir)

from app import app as application
