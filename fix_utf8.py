import os

files = [
    os.path.abspath(r"templates\invoices\bulk_print.html"),
    os.path.abspath(r"templates\print_layout.html")
]

for f in files:
    if os.path.exists(f):
        with open(f, "rb") as fp:
            data = fp.read()
        
        # remove BOM if present
        if data.startswith(b'\xef\xbb\xbf'):
            data = data[3:]
        elif data.startswith(b'\xff\xfe') or data.startswith(b'\xfe\xff'):
            # If accidentally saved as utf-16, decode and encode
            data = data.decode('utf-16').encode('utf-8')
            
        try:
            text = data.decode('utf-8')
        except UnicodeDecodeError:
            try:
                text = data.decode('cp932')
            except:
                text = data.decode('utf-8', errors='ignore')
                
        text = text.replace('\r\n', '\n')
            
        with open(f, "w", encoding="utf-8", newline="\n") as fp:
            fp.write(text)
        print(f"Fixed encoding for {f}")
