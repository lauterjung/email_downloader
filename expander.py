import os
import re
import shutil
from email import policy
from email.parser import BytesParser
from email.utils import parseaddr, parsedate_to_datetime

def move_to_processed(eml_path):
    source_dir = os.path.dirname(eml_path)
    processed_dir = os.path.join(source_dir, "processed")

    os.makedirs(processed_dir, exist_ok=True)

    filename = os.path.basename(eml_path)
    destination = os.path.join(processed_dir, filename)

    counter = 1
    base_name, ext = os.path.splitext(filename)

    while os.path.exists(destination):
        destination = os.path.join(
            processed_dir,
            f"{base_name} ({counter}){ext}"
        )
        counter += 1

    shutil.move(eml_path, destination)
    return destination


def sanitize_name(name: str) -> str:
    if not name:
        return "no_name"

    if "@" in name:
        name = name.split("@")[0]

    name = re.sub(r'[<>:"/\\|?*\n\r\t]', '', name)
    name = name.strip().strip(".")

    return name if name else "no_name"


def create_unique_folder(base_path):
    if not os.path.exists(base_path):
        return base_path

    counter = 1
    while True:
        new_path = f"{base_path} ({counter})"
        if not os.path.exists(new_path):
            return new_path
        counter += 1


def extract_metadata(msg):
    sender_raw = msg.get('From', 'unknown')
    sender_email = parseaddr(sender_raw)[1]
    sender_name = sender_email.split("@")[0] if sender_email else "unknown"

    return {
        "sender_raw": sender_raw,
        "sender_name": sender_name,
        "subject": msg.get('Subject', 'no_subject'),
        "to": msg.get('To', ''),
        "cc": msg.get('Cc', ''),
        "date_raw": msg.get('Date', '')
    }


def format_date(date_raw):
    try:
        parsed_date = parsedate_to_datetime(date_raw)
        return parsed_date.strftime("%Y-%m-%d")
    except Exception:
        return "unknown_date"


def build_folder_path(output_root, metadata):
    sender_folder = sanitize_name(metadata["sender_name"])

    date_prefix = format_date(metadata["date_raw"])
    subject_clean = sanitize_name(metadata["subject"])

    subject_folder = f"{date_prefix} {subject_clean}"

    sender_path = os.path.join(output_root, sender_folder)
    os.makedirs(sender_path, exist_ok=True)

    subject_base = os.path.join(sender_path, subject_folder)
    subject_path = create_unique_folder(subject_base)
    os.makedirs(subject_path)

    return subject_path


def extract_body(msg):
    if msg.is_multipart():
        for part in msg.walk():
            if (
                part.get_content_type() == "text/plain"
                and part.get_content_disposition() != "attachment"
            ):
                return part.get_content()
    else:
        return msg.get_content()

    return "No plain text body found."


def save_body(folder_path, body):
    with open(os.path.join(folder_path, "body.txt"), "w", encoding="utf-8") as f:
        f.write(body)


def save_metadata(folder_path, metadata):
    content = f"""From: {metadata['sender_raw']}
To: {metadata['to']}
Cc: {metadata['cc']}
Date: {metadata['date_raw']}
Subject: {metadata['subject']}
"""

    with open(os.path.join(folder_path, "metadata.txt"), "w", encoding="utf-8") as f:
        f.write(content)


def save_attachments(msg, folder_path):
    for part in msg.walk():
        if part.get_content_disposition() == "attachment":
            filename = part.get_filename()
            if filename:
                filename = sanitize_name(filename)
                filepath = os.path.join(folder_path, filename)

                with open(filepath, "wb") as f:
                    f.write(part.get_payload(decode=True))


def process_eml(eml_path, output_root):
    with open(eml_path, "rb") as f:
        msg = BytesParser(policy=policy.default).parse(f)

    metadata = extract_metadata(msg)
    folder_path = build_folder_path(output_root, metadata)

    body = extract_body(msg)

    save_body(folder_path, body)
    save_metadata(folder_path, metadata)
    save_attachments(msg, folder_path)

    moved_path = move_to_processed(eml_path)

    print(f"Processed: {moved_path}")
    print(f"Saved to: {folder_path}")
    print("-" * 50)


def process_eml_folder(eml_folder, output_root):
    if not os.path.isdir(eml_folder):
        print(f"Folder does not exist: {eml_folder}")
        return

    for filename in os.listdir(eml_folder):
        if filename.lower() == "processed":
            continue

        if filename.lower().endswith(".eml"):
            eml_path = os.path.join(eml_folder, filename)

            try:
                process_eml(eml_path, output_root)
            except Exception as e:
                print(f"Error processing {filename}: {e}")

source_folder = r"D:\emails"
output_folder = r"D:\organized"

process_eml_folder(source_folder, output_folder)