import sys

def cleanup_main_py():
    with open("main.py", "r", encoding="utf-8") as f:
        lines = f.readlines()
    
    # 2380行目あたりから「    except Exception as e:」が2回出てくるはずなので、2回目以降をカット
    # または、より確実に、2377行目の「return pdf.output()」の後の「except Exception as e:」ブロックを探す
    
    new_lines = []
    in_garbage = False
    
    # 2382行目の「return None」の後の「# --- 明細テーブル ---」以降を削除
    # 2487行目の「@app.post("/invoices/{id}/send_email")」までをスキップ
    
    start_skip = -1
    end_skip = -1
    
    for i, line in enumerate(lines):
        if i >= 2380 and "pdf.cell(35, 8, \"単価（円）\"" in line:
            start_skip = i
            break
            
    if start_skip != -1:
        for i in range(start_skip, len(lines)):
            if "@app.post(\"/invoices/{id}/send_email\")" in lines[i]:
                end_skip = i
                break
                
    if start_skip != -1 and end_skip != -1:
        new_lines = lines[:start_skip] + lines[end_skip:]
        with open("main.py", "w", encoding="utf-8") as f:
            f.writelines(new_lines)
        print(f"Cleaned up main.py: Deleted lines {start_skip+1} to {end_skip}")
    else:
        print(f"Could not find skip range. Start: {start_skip}, End: {end_skip}")

if __name__ == "__main__":
    cleanup_main_py()
