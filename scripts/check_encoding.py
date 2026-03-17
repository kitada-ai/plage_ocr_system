import os

root = r"c:\Users\北田航太郎\OneDrive\デスクトップ\VEXUM\プラージュ様\-OCR-main\-OCR-main"
keywords = ["AZURE", "GEMINI", "AIza", "ENDPOINT", "KEY="]

for filename in os.listdir(root):
    if filename.startswith("lint") or filename.endswith(".txt") or filename.endswith(".json"):
        path = os.path.join(root, filename)
        if os.path.isfile(path):
            try:
                # Try reading as UTF-16LE
                with open(path, "rb") as f:
                    content = f.read()
                    text = None
                    if content.startswith(b"\xff\xfe"): # BOM for UTF-16LE
                        text = content.decode("utf-16-le")
                    else:
                        # try UTF-8
                        try:
                            text = content.decode("utf-8")
                        except:
                            pass
                    
                    if text:
                        for kw in keywords:
                            if kw in text:
                                print(f"Found {kw} in {filename}")
                                lines = text.splitlines()
                                for i, line in enumerate(lines):
                                    if kw in line:
                                        print(f"  Line {i+1}: {line.strip()}")
            except Exception as e:
                pass
