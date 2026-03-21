import os
import re

templates_dir = r"c:\Users\nakatake\Desktop\Antigravityフォルダ\kumanomori\templates"

# Regex to match onsubmit="return confirm('...')" or onclick="return confirm('...')"
pattern1 = re.compile(r'\s*onsubmit="return confirm\([^)]+\)"')
pattern2 = re.compile(r"\s*onsubmit='return confirm\([^)]+\)'")
pattern3 = re.compile(r'\s*onclick="return confirm\([^)]+\)"')
pattern4 = re.compile(r"\s*onclick='return confirm\([^)]+\)'")

count = 0
for root, _, files in os.walk(templates_dir):
    for filename in files:
        if filename.endswith(".html"):
            filepath = os.path.join(root, filename)
            with open(filepath, "r", encoding="utf-8") as f:
                content = f.read()
            
            original_content = content
            content = pattern1.sub("", content)
            content = pattern2.sub("", content)
            content = pattern3.sub("", content)
            content = pattern4.sub("", content)
            
            if content != original_content:
                with open(filepath, "w", encoding="utf-8") as f:
                    f.write(content)
                print(f"Fixed: {filepath}")
                count += 1

print(f"Total files fixed: {count}")
