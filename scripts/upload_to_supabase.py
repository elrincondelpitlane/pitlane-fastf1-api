import argparse
import mimetypes
import os
from pathlib import Path

from dotenv import load_dotenv
from supabase import create_client


PROCESSED_DIR = Path("processed")


def required_env(name):
    value = os.environ.get(name)
    if not value:
        raise RuntimeError(f"Missing required environment variable: {name}")
    return value


def iter_files(base_dir):
    for path in sorted(base_dir.rglob("*")):
        if path.is_file():
            yield path


def upload_file(bucket, local_path, storage_path):
    content_type, _ = mimetypes.guess_type(local_path.name)
    file_options = {
        "content-type": content_type or "application/octet-stream",
        "upsert": "true",
    }

    with local_path.open("rb") as file:
        bucket.upload(storage_path, file, file_options)


def upload_processed_files(year, gp, session_code):
    load_dotenv()

    session_dir = PROCESSED_DIR / "sessions" / str(year) / gp / session_code
    telemetry_dir = PROCESSED_DIR / "telemetry" / str(year) / gp / session_code

    if not session_dir.exists():
        raise FileNotFoundError(f"Processed session directory not found: {session_dir}")
    if not telemetry_dir.exists():
        raise FileNotFoundError(f"Processed telemetry directory not found: {telemetry_dir}")

    supabase_url = required_env("SUPABASE_URL")
    supabase_key = required_env("SUPABASE_SERVICE_ROLE_KEY")
    bucket_name = required_env("SUPABASE_FASTF1_BUCKET")

    client = create_client(supabase_url, supabase_key)
    bucket = client.storage.from_(bucket_name)

    uploaded = []
    storage_base = Path(str(year)) / gp / session_code

    for local_path in iter_files(session_dir):
        storage_path = (storage_base / local_path.name).as_posix()
        upload_file(bucket, local_path, storage_path)
        uploaded.append(storage_path)
        print(f"Uploaded {storage_path}")

    for local_path in iter_files(telemetry_dir):
        storage_path = (storage_base / "telemetry" / local_path.name).as_posix()
        upload_file(bucket, local_path, storage_path)
        uploaded.append(storage_path)
        print(f"Uploaded {storage_path}")

    return uploaded


def parse_args():
    parser = argparse.ArgumentParser(description="Upload processed FastF1 JSON files to Supabase Storage.")
    parser.add_argument("--year", type=int, required=True)
    parser.add_argument("--gp", required=True)
    parser.add_argument("--session", required=True)
    return parser.parse_args()


def main():
    args = parse_args()
    uploaded = upload_processed_files(args.year, args.gp, args.session)
    print(f"Uploaded {len(uploaded)} files to Supabase Storage")


if __name__ == "__main__":
    main()
