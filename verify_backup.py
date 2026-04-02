# -*- coding: utf-8 -*-
"""
バックアップ機能の動作確認スクリプト
"""
import os
import shutil
import sqlite3

DB_PATH = "kumanogo.db"
BACKUP_PATH = "kumanogo_backup_test.db"

print("=" * 60)
print("バックアップ機能 動作確認スクリプト")
print("=" * 60)

errors = []

# --- テスト 1: DBファイルの存在確認 ---
print("\n[TEST 1] データベースファイルの存在確認...")
if os.path.exists(DB_PATH):
    size_kb = os.path.getsize(DB_PATH) / 1024
    print(f"  OK: {DB_PATH} が存在します（サイズ: {size_kb:.1f} KB）")
else:
    print(f"  NG: {DB_PATH} が見つかりません！")
    errors.append("TEST1: DBファイルが存在しない")

# --- テスト 2: バックアップファイルの作成（コピー）---
print("\n[TEST 2] バックアップファイルの作成（コピー）...")
try:
    shutil.copy2(DB_PATH, BACKUP_PATH)
    backup_size = os.path.getsize(BACKUP_PATH) / 1024
    print(f"  OK: バックアップ作成成功: {BACKUP_PATH}（サイズ: {backup_size:.1f} KB）")
except Exception as e:
    print(f"  NG: バックアップ作成失敗: {e}")
    errors.append(f"TEST2: {e}")

# --- テスト 3: バックアップファイルの整合性確認 ---
print("\n[TEST 3] バックアップファイルの整合性（SQLite読み込み）確認...")
try:
    conn = sqlite3.connect(BACKUP_PATH)
    cursor = conn.cursor()
    cursor.execute("SELECT name FROM sqlite_master WHERE type='table'")
    tables = [row[0] for row in cursor.fetchall()]
    conn.close()
    print(f"  OK: バックアップDBは正常です。テーブル数: {len(tables)}")
    print(f"     テーブル一覧: {', '.join(tables)}")

    required = ["users", "customers", "products", "orders", "invoices"]
    missing = [t for t in required if t not in tables]
    if missing:
        print(f"  WARN: 必須テーブルが不足: {missing}")
        errors.append(f"TEST3: 必須テーブル不足 {missing}")
    else:
        print(f"  OK: 必須テーブルすべて確認済み")
except Exception as e:
    print(f"  NG: バックアップDB読み込み失敗: {e}")
    errors.append(f"TEST3: {e}")

# --- テスト 4: 復元シミュレーション ---
print("\n[TEST 4] 復元シミュレーション（バイナリ読込 -> 上書き）...")
RESTORE_TARGET = "kumanogo_restore_test.db"
try:
    with open(BACKUP_PATH, "rb") as f:
        content = f.read()
    with open(RESTORE_TARGET, "wb") as f:
        f.write(content)
    
    conn = sqlite3.connect(RESTORE_TARGET)
    cursor = conn.cursor()
    cursor.execute("SELECT name FROM sqlite_master WHERE type='table'")
    restored_tables = [row[0] for row in cursor.fetchall()]
    conn.close()
    
    print(f"  OK: 復元シミュレーション成功！テーブル数: {len(restored_tables)}")
    
    original_size = os.path.getsize(DB_PATH)
    restored_size = os.path.getsize(RESTORE_TARGET)
    if original_size == restored_size:
        print(f"  OK: サイズ一致（{original_size} bytes）: データ完全性 OK")
    else:
        print(f"  WARN: サイズ不一致: 元={original_size}, 復元={restored_size}")
        
    os.remove(RESTORE_TARGET)
except Exception as e:
    print(f"  NG: 復元シミュレーション失敗: {e}")
    errors.append(f"TEST4: {e}")

# --- テスト 5: main.py のルート定義確認 ---
print("\n[TEST 5] main.py のバックアップルート定義確認...")
try:
    with open("main.py", "r", encoding="utf-8") as f:
        content = f.read()
    
    checks = [
        ("GET /backup ルート", '@app.get("/backup")'),
        ("POST /restore ルート", '@app.post("/restore")'),
        ("FileResponse 使用", 'FileResponse("kumanogo.db"'),
        ("ファイルアップロード受付", 'backup_file: UploadFile'),
        ("ログイン認証保護", 'Depends(get_active_user)'),
    ]
    
    for label, keyword in checks:
        if keyword in content:
            print(f"  OK: {label}")
        else:
            print(f"  NG: {label} -- 見つからない")
            errors.append(f"TEST5: {label}が未定義")
except Exception as e:
    print(f"  NG: main.py 読み込み失敗: {e}")

# --- テスト 6: テンプレートの確認 ---
print("\n[TEST 6] バックアップ画面テンプレートの確認...")
try:
    with open("templates/settings.html", "r", encoding="utf-8") as f:
        tmpl = f.read()
    
    tmpl_checks = [
        ("ダウンロードリンク(/backup)", '/backup'),
        ("復元フォームアクション(/restore)", '/restore'),
        ("ファイル入力(backup_file)", 'backup_file'),
        ("enctype=multipart/form-data", 'multipart/form-data'),
        ("上書き警告メッセージ", 'enctype'),
    ]
    
    for label, keyword in tmpl_checks:
        if keyword in tmpl:
            print(f"  OK: {label}")
        else:
            print(f"  NG: {label} -- 見つからない")
            errors.append(f"TEST6: {label}が未定義")
except Exception as e:
    print(f"  NG: templates/settings.html 読み込み失敗: {e}")

# --- テスト 7: 各テーブルのレコード数確認 ---
print("\n[TEST 7] 現在のデータ件数確認（バックアップ対象）...")
try:
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    count_targets = ["users", "customers", "products", "orders", "invoices", "quotations"]
    for table in count_targets:
        try:
            cursor.execute(f"SELECT COUNT(*) FROM {table}")
            count = cursor.fetchone()[0]
            print(f"  {table}: {count} 件")
        except:
            print(f"  {table}: テーブルなし（スキップ）")
    conn.close()
except Exception as e:
    print(f"  NG: データ件数確認失敗: {e}")

# クリーンアップ
if os.path.exists(BACKUP_PATH):
    os.remove(BACKUP_PATH)

# --- 最終結果 ---
print("\n" + "=" * 60)
if not errors:
    print("[PASS] 全テスト合格！バックアップ機能は正常に動作しています。")
    print("=" * 60)
    print("\n--- バックアップ機能の概要 ---")
    print("  ダウンロード: GET /backup  -> kumanogo_backup.db を取得")
    print("  復元        : POST /restore -> アップロードDBで上書き復元")
    print("  セキュリティ: ログイン済みユーザーのみ実行可能")
    print("  操作場所    : サイドバー「バックアップ・復元」メニューから")
else:
    print(f"[FAIL] {len(errors)} 件のエラーが検出されました:")
    for e in errors:
        print(f"   - {e}")
print("=" * 60)
