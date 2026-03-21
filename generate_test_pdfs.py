import urllib.request
import os

base_url = "http://localhost:8000"
endpoints = [
    ("/quotations/1/pdf", "test_quotation.pdf"),
    ("/orders/1/delivery_note", "test_delivery_note.pdf"),
    ("/invoices/1/pdf", "test_invoice.pdf")
]

artifact_dir = r"C:\Users\nakatake\.gemini\antigravity\brain\b184d9ba-1b21-4c57-936f-3c33dc28d7e4"

for endpoint, filename in endpoints:
    url = f"{base_url}{endpoint}"
    print(f"Fetching {url}...")
    try:
        with urllib.request.urlopen(url) as response:
            if response.status == 200:
                filepath = os.path.join(artifact_dir, filename)
                with open(filepath, "wb") as f:
                    f.write(response.read())
                print(f"Saved to {filepath}")
            else:
                print(f"Failed to fetch {url}: {response.status}")
    except Exception as e:
        print(f"Error fetching {url}: {e}")
