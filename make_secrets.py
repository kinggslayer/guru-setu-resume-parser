"""
Turn the Google service-account JSON key into the TOML block that
Streamlit Community Cloud expects.

Run it locally, then copy the output into:
    your app -> Settings -> Secrets

    python make_secrets.py path/to/service_account.json

Hand-copying the 11 fields is where this usually goes wrong — private_key
contains escaped newlines that must survive intact, and TOML needs them
inside a triple-quoted string. This does that part for you.

Nothing is uploaded anywhere; it only reads the file and prints to your
terminal. Don't commit or paste the output anywhere public.
"""

import json
import os
import sys


REQUIRED_FIELDS = [
    "type",
    "project_id",
    "private_key_id",
    "private_key",
    "client_email",
    "client_id",
    "auth_uri",
    "token_uri",
    "auth_provider_x509_cert_url",
    "client_x509_cert_url",
]


def toml_value(value):
    """
    Render one JSON value as TOML.

    private_key is the reason this function exists: it holds literal "\\n"
    escape sequences that must reach Streamlit unchanged. A TOML basic
    string would re-interpret the backslashes, so it goes into a
    triple-quoted literal string instead.
    """

    text = str(value)

    if "\\n" in text or "\n" in text:
        # Normalise real newlines back to the escaped form the key uses.
        text = text.replace("\n", "\\n")
        return f'"""{text}"""'

    return json.dumps(text)


def main():

    if len(sys.argv) != 2:
        print(__doc__)
        print("Error: pass the path to your service-account JSON file.")
        return 1

    path = sys.argv[1]

    if not os.path.exists(path):
        print(f"Error: no such file: {path}")
        return 1

    with open(path, "r", encoding="utf-8") as handle:

        try:
            data = json.load(handle)

        except json.JSONDecodeError as error:
            print(f"Error: that file isn't valid JSON ({error}).")
            return 1

    if data.get("type") != "service_account":
        print(
            "Warning: this doesn't look like a service-account key "
            f"(type is {data.get('type')!r}). Check you downloaded the "
            "right file from Google Cloud."
        )

    missing = [f for f in REQUIRED_FIELDS if f not in data]

    if missing:
        print(f"Warning: missing expected fields: {', '.join(missing)}")

    print()
    print("=" * 68)
    print("  Copy EVERYTHING below into Settings -> Secrets")
    print("  Replace sk-... with your real OpenAI key first.")
    print("=" * 68)
    print()

    print('OPENAI_API_KEY = "sk-REPLACE-WITH-YOUR-KEY"')
    print("OPENAI_WORKERS = 2")
    print()
    print("[gcp_service_account]")

    # Keep a stable, readable field order; anything unexpected in the file
    # still gets emitted rather than silently dropped.
    ordered = REQUIRED_FIELDS + [
        k for k in data if k not in REQUIRED_FIELDS
    ]

    for field in ordered:

        if field in data:
            print(f"{field} = {toml_value(data[field])}")

    print()
    print("=" * 68)
    print(f"  Share your Drive folder and the Resume_Master_DB sheet with:")
    print(f"    {data.get('client_email', '(client_email missing)')}")
    print("  Otherwise the app signs in fine but sees no files.")
    print("=" * 68)

    return 0


if __name__ == "__main__":
    sys.exit(main())
