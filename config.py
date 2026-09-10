import os
from datetime import timedelta


class Config:
    SECRET_KEY = os.environ.get(
        "SECRET_KEY",
        "telephone_directory_secret_key"
    )

    SQLALCHEMY_DATABASE_URI = os.environ.get(
        "DATABASE_URL",
        "postgresql+psycopg://postgres:Wrldc%40123@localhost:5432/telephone_directory"
    )

    SQLALCHEMY_TRACK_MODIFICATIONS = False

    UPLOAD_FOLDER = os.path.join(os.getcwd(), "uploads")

    DIRECTORY_VERSIONS_STORAGE = os.path.join(os.getcwd(), "storage", "directory_versions")

    # Session expires after 20 minutes of inactivity
    PERMANENT_SESSION_LIFETIME = timedelta(minutes=20)

    # Filter string for WRLDC employee directory
    WRLDC_ORG_FILTER = "%WRLDC%"
