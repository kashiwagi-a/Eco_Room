import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import sqlite3
import openpyxl
from openpyxl.styles import PatternFill, Font
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont
from datetime import datetime, timedelta
import os
import subprocess
import platform
import csv
import shutil
import glob

# IME関連警告を抑制
if platform.system() == "Darwin":
    os.environ['TK_SILENCE_DEPRECATION'] = '1'


class HotelCleaningSystem:
    def __init__(self):
        self.db_file = "hotel_cleaning.db"
        self.excel_file = "hotel_cleaning_now.xlsx"
        self.backup_prefix = "hotel_cleaning_backup_"
        self.records = []
        self.existing_rooms = set()
        # 直前のCSV読込で使った部屋状態CSVのパス（エコ票作成→アーニング表出力の再利用用）
        self.last_room_status_csv = None

        # GUI設定
        self.root = tk.Tk()
        self.root.title("客室清掃管理システム")
        # ボタン4つ（エコ票作成／アーニング表出力／エコ部屋リスト／起動メニュー）
        # が横並びで見切れない幅を確保する
        self.root.geometry("560x200")

        self.init_database()

        # 起動時にチェックアウト削除を実行（バックアップは手動）
        self.startup_cleanup()

        self.setup_gui()
        self.load_data()
        self.update_room_count_display()

    def startup_cleanup(self):
        """起動時のチェックアウト削除（バックアップは手動で行う方針のため自動作成しない）"""
        # チェックアウト削除ダイアログを表示
        checkout_date = self.show_checkout_cleanup_dialog()

        if checkout_date:
            # C/Oの部屋と空白の部屋を削除
            deleted_info = self.cleanup_checkout_rooms(checkout_date)

            if deleted_info['total'] > 0:
                message = f"データ整理が完了しました。\n\n"
                message += f"✓ 削除された部屋: {deleted_info['total']}件\n"
                if deleted_info['checkout'] > 0:
                    message += f"  - チェックアウト完了: {deleted_info['checkout']}件\n"
                if deleted_info['empty'] > 0:
                    message += f"  - 空白データ: {deleted_info['empty']}件\n"
                message += f"✓ チェックアウト日: {checkout_date.strftime('%Y年%m月%d日')}"

                messagebox.showinfo("起動時整理完了", message)

    def show_checkout_cleanup_dialog(self):
        """チェックアウト削除用の日付選択ダイアログ"""
        dialog = tk.Toplevel(self.root)
        dialog.title("起動時データ整理")
        # 並び順: CSV読込（メイン）→ 出力 → 削除機能（手動）→ 管理機能
        dialog.geometry("560x660")
        dialog.resizable(False, False)
        dialog.transient(self.root)
        dialog.grab_set()

        frame = ttk.Frame(dialog, padding="25")
        frame.pack(fill="both", expand=True)

        # タイトル
        title_label = ttk.Label(frame, text="起動時データ整理", font=("", 16, "bold"))
        title_label.pack(pady=(0, 10))

        # 結果を格納する変数
        result = {'date': None}

        # --- メイン：CSV読込（独立エリア） ---
        csv_frame = ttk.LabelFrame(frame, text="エコ追加", padding="15")
        csv_frame.pack(fill="x", pady=(0, 15))

        ttk.Label(csv_frame,
                  text="部屋状態CSVを読み込んで部屋を登録します。\n"
                       "読込時にC/O部屋の自動削除（CSV日付の翌日基準）も行われます。",
                  font=("", 10), justify=tk.CENTER).pack(pady=(0, 8))

        def open_csv_import():
            dialog.destroy()
            result['date'] = None
            self.root.after(100, self.import_csv)

        ttk.Button(csv_frame, text="CSV読込", command=open_csv_import,
                   width=20).pack()

        # --- 出力エリア ---
        output_frame = ttk.LabelFrame(frame, text="出力", padding="15")
        output_frame.pack(fill="x", pady=(0, 15))

        output_btn_frame = ttk.Frame(output_frame)
        output_btn_frame.pack()

        def open_earning_output():
            dialog.destroy()
            result['date'] = None  # スキップ扱い
            self.root.after(100, self.output_earning_table)

        def open_eco_room_list():
            dialog.destroy()
            result['date'] = None  # スキップ扱い
            self.root.after(100, self.show_eco_room_list_dialog)

        ttk.Button(output_btn_frame, text="アーニング表出力", command=open_earning_output, width=14).pack(side="left", padx=8, pady=5)
        ttk.Button(output_btn_frame, text="エコ部屋リスト", command=open_eco_room_list, width=14).pack(side="left", padx=8, pady=5)

        # --- 削除機能エリア（手動） ---
        date_frame = ttk.LabelFrame(frame, text="C/O部屋の削除（手動）", padding="15")
        date_frame.pack(fill="x", pady=(0, 15))

        ttk.Label(date_frame,
                  text="CSVを読まない日の整理用。指定日以前にC/Oの部屋を削除します。",
                  font=("", 10), justify=tk.CENTER).pack(pady=(0, 8))

        # デフォルトは今日の日付
        today = datetime.now()
        checkout_vars = {
            'year': tk.StringVar(value=str(today.year)),
            'month': tk.StringVar(value=str(today.month)),
            'day': tk.StringVar(value=str(today.day))
        }

        date_input_frame = ttk.Frame(date_frame)
        date_input_frame.pack()

        ttk.Entry(date_input_frame, textvariable=checkout_vars['year'], width=6).pack(side="left")
        ttk.Label(date_input_frame, text="年").pack(side="left", padx=(0, 10))

        ttk.Entry(date_input_frame, textvariable=checkout_vars['month'], width=4).pack(side="left")
        ttk.Label(date_input_frame, text="月").pack(side="left", padx=(0, 10))

        ttk.Entry(date_input_frame, textvariable=checkout_vars['day'], width=4).pack(side="left")
        ttk.Label(date_input_frame, text="日").pack(side="left")

        def on_ok():
            try:
                checkout_date = datetime(
                    int(checkout_vars['year'].get()),
                    int(checkout_vars['month'].get()),
                    int(checkout_vars['day'].get())
                )
                result['date'] = checkout_date
                dialog.destroy()
            except ValueError:
                messagebox.showerror("エラー", "正しい日付を入力してください")

        delete_frame = ttk.Frame(date_frame)
        delete_frame.pack(pady=(10, 0))
        ttk.Button(delete_frame, text="削除実行", command=on_ok, width=14).pack()

        # --- 管理機能エリア ---
        manage_frame = ttk.LabelFrame(frame, text="管理機能", padding="15")
        manage_frame.pack(fill="x", pady=(0, 10))

        manage_btn_frame = ttk.Frame(manage_frame)
        manage_btn_frame.pack()

        def open_edit_room():
            dialog.destroy()
            result['date'] = None  # スキップ扱い
            self.root.after(100, self.edit_room)

        def open_backup_management():
            dialog.destroy()
            result['date'] = None  # スキップ扱い
            self.root.after(100, self.show_backup_management)

        ttk.Button(manage_btn_frame, text="部屋編集", command=open_edit_room, width=14).pack(side="left", padx=8, pady=5)
        ttk.Button(manage_btn_frame, text="バックアップ管理", command=open_backup_management, width=14).pack(side="left", padx=8, pady=5)

        # ダイアログが閉じられるまで待機
        dialog.wait_window()

        return result['date']

    def cleanup_checkout_rooms(self, checkout_date):
        """指定日より前にC/Oステータスの部屋と空白の部屋を削除"""
        cursor = self.conn.cursor()

        # 指定日以前にC/Oステータスの部屋を検索
        checkout_date_str = checkout_date.strftime('%Y-%m-%d')
        cursor.execute("""
                       SELECT DISTINCT room_number
                       FROM cleaning_schedule
                       WHERE cleaning_date <= ?
                         AND cleaning_status = 'C/O'
                       """, (checkout_date_str,))

        checkout_rooms = [row[0] for row in cursor.fetchall()]

        # デバッグ：C/O部屋を表示
        print(f"C/O部屋: {checkout_rooms}")

        # 清掃スケジュールが空白（全く登録されていない）部屋を検索
        cursor.execute("""
                       SELECT room_number
                       FROM rooms
                       WHERE room_number NOT IN (SELECT DISTINCT room_number
                                                 FROM cleaning_schedule)
                       """)

        empty_rooms = [row[0] for row in cursor.fetchall()]

        # デバッグ：空白部屋を表示
        print(f"空白部屋: {empty_rooms}")

        # さらに、cleaning_statusが空文字列やNULLの部屋も検索
        cursor.execute("""
                       SELECT DISTINCT room_number
                       FROM cleaning_schedule
                       WHERE cleaning_status = ''
                          OR cleaning_status IS NULL
                       """)

        null_or_empty_rooms = [row[0] for row in cursor.fetchall()]
        print(f"空文字列/NULL部屋: {null_or_empty_rooms}")

        # 全ての空白パターンを統合
        rooms_to_delete = list(set(checkout_rooms + empty_rooms + null_or_empty_rooms))

        # デバッグ：削除対象の部屋を表示
        print(f"削除対象の部屋: {rooms_to_delete}")

        if rooms_to_delete:
            # rooms テーブルから削除
            placeholders = ','.join(['?' for _ in rooms_to_delete])
            cursor.execute(f"DELETE FROM rooms WHERE room_number IN ({placeholders})", rooms_to_delete)

            # cleaning_schedule テーブルからも削除
            cursor.execute(f"DELETE FROM cleaning_schedule WHERE room_number IN ({placeholders})", rooms_to_delete)

            self.conn.commit()

            # 削除の内訳を返す
            return {
                'total': len(rooms_to_delete),
                'checkout': len(checkout_rooms),
                'empty': len(empty_rooms),
                'null_or_empty': len(null_or_empty_rooms)
            }

        return {'total': 0, 'checkout': 0, 'empty': 0, 'null_or_empty': 0}

    def create_database_backup_silent(self):
        """サイレントバックアップ作成（メッセージなし）"""
        try:
            backup_name = f"{self.backup_prefix}{datetime.now().strftime('%Y%m%d_%H%M%S')}.db"
            shutil.copy2(self.db_file, backup_name)
            return backup_name
        except Exception as e:
            print(f"バックアップエラー: {e}")
            return "作成失敗"

    def get_backup_list(self):
        """バックアップファイルの一覧を取得"""
        backup_files = glob.glob(f"{self.backup_prefix}*.db")
        # 手動で付けた任意名のバックアップ(.db)も拾えるよう、カレントの.dbも対象に含める
        # （ただし運用中のDB本体は除外する）
        for f in glob.glob("*.db"):
            if f not in backup_files and os.path.basename(f) != os.path.basename(self.db_file):
                backup_files.append(f)
        backup_info = []

        for backup_file in backup_files:
            try:
                # ファイル名から日時を抽出（旧形式 YYYYMMDD_HHMMSS のときのみ成功）
                filename = os.path.basename(backup_file)
                date_str = filename.replace(self.backup_prefix, "").replace(".db", "")
                try:
                    backup_date = datetime.strptime(date_str, '%Y%m%d_%H%M%S')
                except ValueError:
                    # 手動命名など解析できない場合はファイルの更新日時で代替
                    backup_date = datetime.fromtimestamp(os.path.getmtime(backup_file))

                # ファイルサイズを取得
                file_size = os.path.getsize(backup_file)
                size_str = self.format_file_size(file_size)

                # バックアップ内の部屋数を取得
                room_count = self.get_backup_room_count(backup_file)

                backup_info.append({
                    'filename': backup_file,
                    'date': backup_date,
                    'date_str': backup_date.strftime('%Y年%m月%d日 %H:%M:%S'),
                    'size': size_str,
                    'room_count': room_count
                })
            except Exception as e:
                print(f"バックアップ情報取得エラー ({backup_file}): {e}")
                continue

        # 日付の新しい順にソート
        backup_info.sort(key=lambda x: x['date'], reverse=True)
        return backup_info

    def format_file_size(self, size_bytes):
        """ファイルサイズを読みやすい形式に変換"""
        if size_bytes < 1024:
            return f"{size_bytes} B"
        elif size_bytes < 1024 * 1024:
            return f"{size_bytes / 1024:.1f} KB"
        else:
            return f"{size_bytes / (1024 * 1024):.1f} MB"

    def get_backup_room_count(self, backup_file):
        """バックアップファイル内の部屋数を取得"""
        try:
            conn = sqlite3.connect(backup_file)
            cursor = conn.cursor()
            cursor.execute("SELECT COUNT(*) FROM rooms")
            count = cursor.fetchone()[0]
            conn.close()
            return count
        except Exception:
            return "不明"

    def show_restore_dialog(self):
        """バックアップ復元ダイアログを表示"""
        backup_list = self.get_backup_list()

        if not backup_list:
            messagebox.showinfo("情報", "復元可能なバックアップファイルがありません。")
            return

        dialog = tk.Toplevel(self.root)
        dialog.title("バックアップ復元")
        dialog.geometry("650x650")
        dialog.minsize(600, 600)
        dialog.resizable(True, True)
        dialog.transient(self.root)
        dialog.grab_set()

        frame = ttk.Frame(dialog, padding="15")
        frame.pack(fill="both", expand=True)

        # タイトル
        title_label = ttk.Label(frame, text="バックアップからの復元", font=("", 14, "bold"))
        title_label.pack(pady=(0, 10))

        # 説明
        info_label = ttk.Label(frame,
                               text="復元したいバックアップを選択してください。\n"
                                    "⚠️ 復元を実行すると、現在のデータは上書きされます。",
                               font=("", 10), justify=tk.CENTER, foreground="red")
        info_label.pack(pady=(0, 15))

        # バックアップ一覧フレーム
        list_frame = ttk.LabelFrame(frame, text=f"バックアップ一覧（{len(backup_list)}件）", padding="10")
        list_frame.pack(fill="both", expand=True, pady=(0, 15))

        # Treeviewでバックアップ一覧を表示
        columns = ("date", "rooms", "size")
        tree = ttk.Treeview(list_frame, columns=columns, show="headings", height=15)

        tree.heading("date", text="作成日時")
        tree.heading("rooms", text="部屋数")
        tree.heading("size", text="サイズ")

        tree.column("date", width=200, anchor="w")
        tree.column("rooms", width=80, anchor="center")
        tree.column("size", width=80, anchor="center")

        # スクロールバー
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)

        tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # バックアップデータを挿入
        for backup in backup_list:
            tree.insert("", "end", values=(
                backup['date_str'],
                backup['room_count'],
                backup['size']
            ), tags=(backup['filename'],))

        # 選択されたバックアップのファイル名を格納
        selected_backup = {'filename': None}

        def on_select(event):
            selection = tree.selection()
            if selection:
                item = tree.item(selection[0])
                selected_backup['filename'] = item['tags'][0]

        tree.bind('<<TreeviewSelect>>', on_select)

        # プレビュー情報フレーム
        preview_frame = ttk.LabelFrame(frame, text="選択中のバックアップ情報", padding="10")
        preview_frame.pack(fill="x", pady=(0, 10))

        preview_label = ttk.Label(preview_frame, text="バックアップを選択してください", font=("", 9))
        preview_label.pack()

        def update_preview(event):
            selection = tree.selection()
            if selection:
                item = tree.item(selection[0])
                filename = item['tags'][0]
                for backup in backup_list:
                    if backup['filename'] == filename:
                        preview_text = f"ファイル: {backup['filename']}\n"
                        preview_text += f"作成日時: {backup['date_str']}\n"
                        preview_text += f"部屋数: {backup['room_count']}件 | サイズ: {backup['size']}"
                        preview_label.config(text=preview_text)
                        break

        tree.bind('<<TreeviewSelect>>', lambda e: (on_select(e), update_preview(e)))

        # ボタンフレーム
        button_frame = ttk.Frame(frame)
        button_frame.pack(fill="x", pady=(10, 5))

        def restore_backup():
            if not selected_backup['filename']:
                messagebox.showwarning("警告", "復元するバックアップを選択してください。")
                return

            # 確認ダイアログ
            confirm = messagebox.askyesno(
                "復元の確認",
                f"選択したバックアップから復元しますか？\n\n"
                f"ファイル: {selected_backup['filename']}\n\n"
                f"⚠️ 現在のデータは完全に上書きされます。\n"
                f"この操作は取り消せません。",
                icon='warning'
            )

            if confirm:
                success = self.restore_from_backup(selected_backup['filename'])
                dialog.destroy()

                if success:
                    messagebox.showinfo("完了", "バックアップからの復元が完了しました。\n\nデータを再読み込みしました。")
                else:
                    messagebox.showerror("エラー", "バックアップの復元に失敗しました。")

        def delete_backup():
            if not selected_backup['filename']:
                messagebox.showwarning("警告", "削除するバックアップを選択してください。")
                return

            confirm = messagebox.askyesno(
                "削除の確認",
                f"選択したバックアップを削除しますか？\n\n"
                f"ファイル: {selected_backup['filename']}\n\n"
                f"この操作は取り消せません。",
                icon='warning'
            )

            if confirm:
                try:
                    os.remove(selected_backup['filename'])
                    # 一覧を更新
                    for item in tree.get_children():
                        if tree.item(item)['tags'][0] == selected_backup['filename']:
                            tree.delete(item)
                            break
                    selected_backup['filename'] = None
                    preview_label.config(text="バックアップを選択してください")
                    messagebox.showinfo("完了", "バックアップを削除しました。")
                except Exception as e:
                    messagebox.showerror("エラー", f"バックアップの削除に失敗しました: {e}")

        def browse_backup():
            """外部のバックアップファイルを選択"""
            file_path = filedialog.askopenfilename(
                title="バックアップファイルを選択"
            )

            if file_path:
                # 確認ダイアログ
                confirm = messagebox.askyesno(
                    "復元の確認",
                    f"選択したファイルから復元しますか？\n\n"
                    f"ファイル: {file_path}\n\n"
                    f"⚠️ 現在のデータは完全に上書きされます。",
                    icon='warning'
                )

                if confirm:
                    success = self.restore_from_backup(file_path)
                    dialog.destroy()

                    if success:
                        messagebox.showinfo("完了", "バックアップからの復元が完了しました。")
                    else:
                        messagebox.showerror("エラー",
                                             "バックアップの復元に失敗しました。\nファイル形式が正しいか確認してください。")

        ttk.Button(button_frame, text="復元", command=restore_backup).pack(side="left", padx=5, pady=5)
        ttk.Button(button_frame, text="削除", command=delete_backup).pack(side="left", padx=5, pady=5)
        ttk.Button(button_frame, text="ファイルから選択...", command=browse_backup).pack(side="left", padx=5, pady=5)
        ttk.Button(button_frame, text="閉じる", command=dialog.destroy).pack(side="right", padx=5, pady=5)

        # 下部に余白を追加
        ttk.Label(frame, text="").pack(pady=5)

    def restore_from_backup(self, backup_file):
        """バックアップファイルからデータベースを復元"""
        try:
            # バックアップファイルの検証
            test_conn = sqlite3.connect(backup_file)
            test_cursor = test_conn.cursor()

            # 必要なテーブルが存在するか確認
            test_cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='rooms'")
            if not test_cursor.fetchone():
                test_conn.close()
                return False

            test_cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='cleaning_schedule'")
            if not test_cursor.fetchone():
                test_conn.close()
                return False

            test_conn.close()

            # 現在の接続を閉じる
            self.conn.close()

            # 復元前に現在のデータベースのバックアップを作成
            pre_restore_backup = f"{self.backup_prefix}pre_restore_{datetime.now().strftime('%Y%m%d_%H%M%S')}.db"
            shutil.copy2(self.db_file, pre_restore_backup)
            print(f"復元前バックアップ作成: {pre_restore_backup}")

            # バックアップファイルで上書き
            shutil.copy2(backup_file, self.db_file)

            # データベースに再接続
            self.conn = sqlite3.connect(self.db_file)

            # データを再読み込み
            self.records.clear()
            self.existing_rooms.clear()
            self.load_data()

            return True

        except Exception as e:
            print(f"復元エラー: {e}")
            # エラー時はデータベースに再接続を試みる
            try:
                self.conn = sqlite3.connect(self.db_file)
            except:
                pass
            return False

    def show_backup_management(self):
        """バックアップ管理メニューを表示"""
        menu = tk.Menu(self.root, tearoff=0)
        menu.add_command(label="バックアップから復元...", command=self.show_restore_dialog)
        menu.add_command(label="今すぐバックアップ作成", command=self.create_manual_backup)
        menu.add_separator()
        menu.add_command(label="古いバックアップを削除...", command=self.cleanup_old_backups_dialog)

        # ボタンの位置に表示
        try:
            menu.tk_popup(self.root.winfo_pointerx(), self.root.winfo_pointery())
        finally:
            menu.grab_release()

    def create_manual_backup(self):
        """手動でバックアップを作成（ファイル名はユーザーが入力）"""
        # ファイル名入力ダイアログ（初期値なし）
        dialog = tk.Toplevel(self.root)
        dialog.title("バックアップ作成")
        dialog.geometry("420x200")
        dialog.resizable(False, False)
        dialog.transient(self.root)
        dialog.grab_set()

        frame = ttk.Frame(dialog, padding="20")
        frame.pack(fill="both", expand=True)

        ttk.Label(frame, text="バックアップのファイル名を入力してください。",
                  font=("", 11)).pack(pady=(0, 5))
        ttk.Label(frame, text="（拡張子 .db は自動で付きます）",
                  font=("", 9), foreground="gray").pack(pady=(0, 12))

        name_var = tk.StringVar(value="")
        name_entry = ttk.Entry(frame, textvariable=name_var, width=36)
        name_entry.pack(pady=(0, 15))
        name_entry.focus_set()

        def do_save():
            raw = name_var.get().strip()
            if not raw:
                messagebox.showwarning("警告", "ファイル名を入力してください。", parent=dialog)
                return

            # 拡張子 .db を補完
            filename = raw if raw.lower().endswith(".db") else raw + ".db"

            # 同名ファイルの存在チェック → 上書き確認
            if os.path.exists(filename):
                overwrite = messagebox.askyesno(
                    "上書きの確認",
                    f"同じ名前のファイルがあります。上書きしますか？\n\nファイル: {filename}",
                    parent=dialog
                )
                if not overwrite:
                    return

            try:
                shutil.copy2(self.db_file, filename)
                dialog.destroy()
                messagebox.showinfo("完了", f"バックアップを作成しました。\n\nファイル名: {filename}")
            except Exception as e:
                messagebox.showerror("エラー", f"バックアップの作成に失敗しました: {e}", parent=dialog)

        button_frame = ttk.Frame(frame)
        button_frame.pack()

        ttk.Button(button_frame, text="作成", command=do_save).pack(side="left", padx=5)
        ttk.Button(button_frame, text="キャンセル", command=dialog.destroy).pack(side="left", padx=5)

        # Enterキーで作成
        name_entry.bind("<Return>", lambda e: do_save())

    def cleanup_old_backups_dialog(self):
        """古いバックアップを削除するダイアログ"""
        backup_list = self.get_backup_list()

        if len(backup_list) <= 5:
            messagebox.showinfo("情報",
                                f"現在のバックアップ数: {len(backup_list)}件\n\n5件以下のため、削除の必要はありません。")
            return

        dialog = tk.Toplevel(self.root)
        dialog.title("古いバックアップの削除")
        dialog.geometry("400x200")
        dialog.resizable(False, False)
        dialog.transient(self.root)
        dialog.grab_set()

        frame = ttk.Frame(dialog, padding="20")
        frame.pack(fill="both", expand=True)

        ttk.Label(frame, text=f"現在のバックアップ数: {len(backup_list)}件", font=("", 11)).pack(pady=(0, 15))

        ttk.Label(frame, text="保持するバックアップ数を選択してください：").pack()

        keep_var = tk.IntVar(value=10)
        keep_frame = ttk.Frame(frame)
        keep_frame.pack(pady=10)

        ttk.Spinbox(keep_frame, from_=1, to=len(backup_list), textvariable=keep_var, width=10).pack(side="left")
        ttk.Label(keep_frame, text="件（新しい順に保持）").pack(side="left", padx=5)

        def do_cleanup():
            keep_count = keep_var.get()
            if keep_count >= len(backup_list):
                messagebox.showinfo("情報", "削除対象のバックアップがありません。")
                return

            to_delete = backup_list[keep_count:]

            confirm = messagebox.askyesno(
                "確認",
                f"{len(to_delete)}件のバックアップを削除しますか？\n\n"
                f"最も古いバックアップ: {to_delete[-1]['date_str']}"
            )

            if confirm:
                deleted = 0
                for backup in to_delete:
                    try:
                        os.remove(backup['filename'])
                        deleted += 1
                    except Exception as e:
                        print(f"削除エラー: {e}")

                dialog.destroy()
                messagebox.showinfo("完了", f"{deleted}件のバックアップを削除しました。")

        button_frame = ttk.Frame(frame)
        button_frame.pack(pady=15)

        ttk.Button(button_frame, text="削除実行", command=do_cleanup).pack(side="left", padx=5)
        ttk.Button(button_frame, text="キャンセル", command=dialog.destroy).pack(side="left", padx=5)

    def setup_gui(self):
        frame = ttk.Frame(self.root, padding="10")
        frame.pack(fill="both", expand=True)

        # 登録済み部屋数の表示
        self.room_count_var = tk.StringVar(value="登録部屋数: 0件")
        ttk.Label(frame, textvariable=self.room_count_var, font=("", 12)).pack(pady=(10, 5))

        # ステータス表示
        self.status_var = tk.StringVar(value="CSVを読み込んでください")
        ttk.Label(frame, textvariable=self.status_var, font=("", 10)).pack(pady=(0, 20))

        # ボタン
        btn_frame = ttk.Frame(frame)
        btn_frame.pack(pady=10)

        ttk.Button(btn_frame, text="エコ票作成", command=self.create_schedule).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="アーニング表出力", command=self.output_earning_table).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="エコ部屋リスト", command=self.show_eco_room_list_dialog).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="起動メニュー", command=self.show_startup_menu).pack(side="left", padx=5)

    def update_room_count_display(self):
        """登録部屋数の表示を更新"""
        count = len(self.records)
        self.room_count_var.set(f"登録部屋数: {count}件")
        if count > 0:
            self.status_var.set("「エコ票作成」でExcelを生成できます")
        else:
            self.status_var.set("CSVを読み込んでください")

    def show_startup_menu(self):
        """起動メニューを再表示（データ整理・CSV読込・部屋編集・バックアップ）"""
        checkout_date = self.show_checkout_cleanup_dialog()

        if checkout_date:
            deleted_info = self.cleanup_checkout_rooms(checkout_date)

            if deleted_info['total'] > 0:
                message = f"データ整理が完了しました。\n\n"
                message += f"✓ 削除された部屋: {deleted_info['total']}件\n"
                if deleted_info['checkout'] > 0:
                    message += f"  - チェックアウト完了: {deleted_info['checkout']}件\n"
                if deleted_info['empty'] > 0:
                    message += f"  - 空白データ: {deleted_info['empty']}件\n"

                messagebox.showinfo("整理完了", message)

            # データ再読み込み
            self.records.clear()
            self.existing_rooms.clear()
            self.load_data()

        self.update_room_count_display()

    def init_database(self):
        self.conn = sqlite3.connect(self.db_file)
        cursor = self.conn.cursor()

        cursor.execute('''CREATE TABLE IF NOT EXISTS rooms
                          (
                              room_number
                              TEXT
                              PRIMARY
                              KEY,
                              guest_name
                              TEXT,
                              check_in_date
                              DATE,
                              cleaning_days
                              INTEGER,
                              is_ecodoor
                              BOOLEAN,
                              is_ecoplan
                              BOOLEAN,
                              middle_status
                              TEXT
                          )''')

        # 既存DBに middle_status 列が無ければ追加（清掃不要・アメ不要 対応）
        cursor.execute("PRAGMA table_info(rooms)")
        existing_columns = [row[1] for row in cursor.fetchall()]
        if 'middle_status' not in existing_columns:
            cursor.execute(
                "ALTER TABLE rooms ADD COLUMN middle_status TEXT DEFAULT ''")

        cursor.execute('''CREATE TABLE IF NOT EXISTS cleaning_schedule
                          (
                              room_number
                              TEXT,
                              cleaning_date
                              DATE,
                              cleaning_status
                              TEXT
                          )''')

        self.conn.commit()

    def load_data(self):
        cursor = self.conn.cursor()
        cursor.execute(
            "SELECT room_number, guest_name, check_in_date, cleaning_days, "
            "is_ecodoor, is_ecoplan, middle_status "
            "FROM rooms ORDER BY CAST(room_number AS INTEGER)")

        for row in cursor.fetchall():
            room, guest, date_str, days, ecodoor, ecoplan, middle_status = row
            self.existing_rooms.add(room)

            # 旧データ（middle_status未設定）は is_ecodoor から補完
            if not middle_status:
                middle_status = 'エコドア' if ecodoor else '×'

            record = {
                'room': room,
                'guest': guest,
                'date': datetime.strptime(date_str, '%Y-%m-%d'),
                'days': days,
                'ecodoor': bool(ecodoor),
                'ecoplan': bool(ecoplan),
                'middle_status': middle_status,
                'schedule': {},
                'is_new': False
            }

            # スケジュール読み込み
            cursor.execute(
                "SELECT cleaning_date, cleaning_status FROM cleaning_schedule WHERE room_number = ? ORDER BY cleaning_date",
                (room,))
            for date_str, status in cursor.fetchall():
                date = datetime.strptime(date_str, '%Y-%m-%d')
                record['schedule'][f"{date.month}/{date.day}"] = status

            self.records.append(record)

    def detect_csv_type(self, file_path):
        """CSVファイルの種別を中身から自動判定する。

        Returns:
            'room_status': 部屋状態CSV (utf-8/ASCII, 8桁日付+部屋番号で始まる)
            'yoyaku'    : 予約CSV (cp932/shift_jis, 12列以上, 202605292形式)
            'yoyaku_co' : 予約詳細CSV (cp932/shift_jis, 202605291形式,
                          5列目=C/I日・6列目=C/O日が8桁日付・7列目=泊数)
            'unknown'   : 判定不能
        """

        def classify_yoyaku(row):
            """12列以上のCSV行を 'yoyaku_co' / 'yoyaku' に振り分ける"""
            # 202605291形式: 列4(5列目)=チェックイン日, 列5(6列目)=チェックアウト日
            # がともに8桁日付で、列6(7列目)が泊数（数字）
            if (len(row) >= 7
                    and row[4].strip().isdigit() and len(row[4].strip()) == 8
                    and row[5].strip().isdigit() and len(row[5].strip()) == 8
                    and row[6].strip().isdigit()):
                return 'yoyaku_co'
            return 'yoyaku'
        # ① まず utf-8 で開けるか試す（部屋状態CSVは数字とASCIIのみ）
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                reader = csv.reader(f)
                first_row = next(reader, None)
                if first_row is None:
                    return 'unknown'

                # 部屋状態CSVのシグネチャ:
                #   列数 7〜10、1列目が8桁数字(日付)、2列目が数字(部屋番号)
                if (7 <= len(first_row) <= 10
                        and first_row[0].isdigit() and len(first_row[0]) == 8
                        and first_row[1].isdigit()):
                    return 'room_status'

                # utf-8で読めて12列以上なら予約CSV(レアケース)
                if len(first_row) >= 12:
                    return classify_yoyaku(first_row)

                return 'unknown'

        except UnicodeDecodeError:
            pass  # utf-8で読めない → 日本語入りのファイル
        except Exception:
            return 'unknown'

        # ② cp932 / shift_jis で予約CSVとして判定
        for encoding in ('cp932', 'shift_jis'):
            try:
                with open(file_path, 'r', encoding=encoding) as f:
                    reader = csv.reader(f)
                    first_row = next(reader, None)
                    if first_row is None:
                        continue
                    if len(first_row) >= 12:
                        return classify_yoyaku(first_row)
            except (UnicodeDecodeError, Exception):
                continue

        return 'unknown'

    def import_csv(self):
        """CSVファイルを読み込んでエコ清掃対象部屋を選択するダイアログを表示。
        部屋状態CSV・予約CSV・予約詳細CSV(C/O日入り) はどの順番で選択しても
        自動で振り分ける。
        """
        # 1ファイル目（必須）
        file1 = filedialog.askopenfilename(
            title="CSVファイルを選択（部屋状態CSV / 予約CSV / 予約詳細CSV）"
        )
        if not file1:
            return

        # 2ファイル目（スキップ可）
        file2 = filedialog.askopenfilename(
            title="2つ目のCSVを選択（スキップ可）"
        )

        # 3ファイル目（スキップ可）※2つ目をスキップした場合は聞かない
        file3 = None
        if file2:
            file3 = filedialog.askopenfilename(
                title="3つ目のCSVを選択（スキップ可）"
            )

        # 自動判定して振り分け
        room_status_path = None
        yoyaku_path = None
        yoyaku_co_path = None  # 202605291形式（C/O日・泊数入り）

        kind_names = {
            'room_status': '部屋状態CSV',
            'yoyaku': '予約CSV',
            'yoyaku_co': '予約詳細CSV（C/O日入り）',
        }

        for path in filter(None, [file1, file2, file3]):
            kind = self.detect_csv_type(path)
            if kind == 'room_status':
                if room_status_path:
                    messagebox.showerror(
                        "エラー",
                        "部屋状態CSVが2つ選択されています。"
                    )
                    return
                room_status_path = path
            elif kind == 'yoyaku':
                if yoyaku_path:
                    messagebox.showerror(
                        "エラー",
                        "予約CSVが2つ選択されています。"
                    )
                    return
                yoyaku_path = path
            elif kind == 'yoyaku_co':
                if yoyaku_co_path:
                    messagebox.showerror(
                        "エラー",
                        "予約詳細CSV（C/O日入り）が2つ選択されています。"
                    )
                    return
                yoyaku_co_path = path
            else:
                messagebox.showerror(
                    "エラー",
                    f"ファイル形式を判定できませんでした:\n"
                    f"{os.path.basename(path)}\n\n"
                    "部屋状態CSV / 予約CSV / 予約詳細CSV を選択してください。"
                )
                return

        if not room_status_path:
            selected = [kind_names[self.detect_csv_type(p)]
                        for p in filter(None, [file1, file2, file3])]
            messagebox.showerror(
                "エラー",
                "部屋状態CSVが必要です。\n"
                f"選択されたファイル: {', '.join(selected) if selected else 'なし'}"
            )
            return

        # 後でエコ票作成→アーニング表出力に再利用するため、部屋状態CSVのパスを保持
        self.last_room_status_csv = room_status_path

        try:
            # 部屋状態CSVを読み込み
            eco_rooms = []
            room_status_map = {}  # 部屋番号 -> 状態コード（全部屋。DB登録との不一致確認用）
            csv_date = None  # CSV 1行目1列目から取得するチェックイン日
            with open(room_status_path, 'r', encoding='utf-8') as f:
                reader = csv.reader(f)
                for i, row in enumerate(reader):
                    # 1行目1列目から YYYYMMDD 形式の日付を抽出
                    if i == 0 and len(row) >= 1 and row[0].isdigit() and len(row[0]) == 8:
                        try:
                            csv_date = datetime.strptime(row[0], '%Y%m%d')
                        except ValueError:
                            csv_date = None

                    if len(row) >= 7:
                        room_number = row[1]  # 2列目：部屋番号
                        room_status = row[6]  # 7列目：部屋の状態
                        room_status_map[room_number] = room_status

                        # 状態が'3'（在室・連泊中＝エコ清掃対象）の部屋のみ抽出
                        # ※ '1'=未チェックイン、'0'=空室、'2'=チェックアウト は対象外
                        if room_status == '3':
                            eco_rooms.append({
                                'room': room_number,
                                'status': room_status
                            })

            if not eco_rooms:
                messagebox.showinfo("情報", "エコ清掃対象の部屋が見つかりませんでした。")
                return

            # 予約詳細CSV(202605291形式)から 予約番号→(C/O日, 泊数) のマップ作成（あれば）
            checkout_map = {}
            if yoyaku_co_path:
                checkout_map = self.load_checkout_map(yoyaku_co_path)

            # 予約CSVから宿泊者名取得（あれば）。checkout_mapを渡すと
            # 予約番号をキーに部屋ごとのC/O日・泊数も紐づける
            guest_name_map = {}
            if yoyaku_path:
                guest_name_map = self.load_guest_names_from_yoyaku(
                    yoyaku_path, checkout_map)

            # CSVの日付（=清掃日）の「翌日」を基準に、C/O部屋等を自動削除する。
            # 従来は翌朝の「削除実行」（起動時データ整理）で行っていた削除を
            # 読み込み時点で済ませる（翌朝C/Oの部屋もここで消える。
            # そのため当日のアーニング表ではC/O=翌日の部屋は連泊扱いになる）。
            # 誤ったCSVを読み込んだ場合はバックアップ管理から復元する。
            # 起動時に既にバックアップ作成済みのため、ここでは作成しない
            if csv_date:
                cleanup_date = csv_date + timedelta(days=1)
                deleted_info = self.cleanup_checkout_rooms(cleanup_date)
                if deleted_info['total'] > 0:
                    # データ再読み込み（削除を反映 → 同番号の再来があれば「未登録」扱いになる）
                    self.records.clear()
                    self.existing_rooms.clear()
                    self.load_data()
                    self.update_room_count_display()

                    message = f"CSVの日付に基づき、データ整理を実施しました。\n\n"
                    message += f"✓ 削除された部屋: {deleted_info['total']}件\n"
                    if deleted_info['checkout'] > 0:
                        message += f"  - チェックアウト完了: {deleted_info['checkout']}件\n"
                    if deleted_info['empty'] > 0:
                        message += f"  - 空白データ: {deleted_info['empty']}件\n"
                    message += f"✓ 清掃日: {csv_date.strftime('%Y年%m月%d日')}\n"
                    message += (f"✓ 削除基準日: {cleanup_date.strftime('%Y年%m月%d日')}"
                                "（清掃日の翌日。C/Oがこの日以前の部屋を削除）")

                    messagebox.showinfo("CSV読込時データ整理", message)

            # 部屋移動・アーリーC/Oなどで、DBに登録が残っているのに
            # CSVでは在室(3)でなくなった部屋を確認し、部屋移動／削除／そのままを選んでもらう
            # （C/O日基準の自動削除では消えないため）
            mismatched = self._find_status_mismatch_rooms(room_status_map, guest_name_map)
            if mismatched:
                self.show_status_mismatch_dialog(mismatched, room_status_map,
                                                 guest_name_map, csv_date)

            # 部屋選択ダイアログを表示
            self.show_csv_room_selection_dialog(eco_rooms, guest_name_map, csv_date)

        except Exception as e:
            messagebox.showerror("エラー", f"CSVファイルの読み込みに失敗しました: {e}")

    # 部屋状態CSVの状態コード表示名（不一致確認ダイアログ用）
    CSV_STATUS_LABELS = {'0': '空室', '1': '未C/I', '2': 'C/O', '3': '在室'}

    # 不一致確認ダイアログの処理選択肢
    MISMATCH_ACTION_MOVE = '部屋移動'
    MISMATCH_ACTION_DELETE = '削除'
    MISMATCH_ACTION_KEEP = 'そのまま'

    # 不一致確認の実行結果を追記するログファイル（実行ディレクトリ直下）
    STATUS_CHECK_LOG_FILE = '部屋状態確認ログ.txt'

    @staticmethod
    def _normalize_guest_name(name):
        """宿泊者名の比較用正規化（NFKC＋空白除去）。空なら空文字"""
        if not name:
            return ''
        import unicodedata
        name = unicodedata.normalize('NFKC', name)
        return ''.join(name.split())

    def _find_status_mismatch_rooms(self, room_status_map, guest_name_map):
        """DBに登録が残っているのに、部屋状態CSVでは在室(3)でない部屋を探す。

        状態 0(空室)／1(未C/I)／2(C/O) の登録部屋が対象。CSVに載っていない
        部屋は判断できないので対象外。C/O日基準の自動削除の後に呼ぶ前提
        （残っているのは登録C/O日より前に部屋が空いたもの＝部屋移動や
        アーリーC/O）。

        予約CSVがあれば、登録者名と同じ宿泊者がいる在室部屋を移動先候補として
        自動で探す。その部屋がすでにDB登録済みなら候補ではなく
        registered_matches（同名の宿泊者がすでに別部屋で登録済み）に入れる。
        """
        # 在室部屋の宿泊者名（予約CSV）を正規化して部屋ごとに持っておく
        in_house_names = {}
        for room, status in room_status_map.items():
            if status != '3':
                continue
            info = guest_name_map.get(room, {})
            name = info.get('name', '') if isinstance(info, dict) else (info or '')
            in_house_names[room] = self._normalize_guest_name(name)

        mismatched = []
        for record in sorted(self.records, key=lambda r: r['room']):
            room = record['room']
            status = room_status_map.get(room)
            if status is None or status == '3':
                continue

            guest_key = self._normalize_guest_name(record.get('guest'))
            candidates = []
            registered_matches = []
            if guest_key:
                for other_room, other_key in sorted(in_house_names.items()):
                    if other_room == room or other_key != guest_key:
                        continue
                    if other_room in self.existing_rooms:
                        registered_matches.append(other_room)
                    else:
                        candidates.append(other_room)

            mismatched.append({
                'room': room,
                'guest': record.get('guest') or '',
                'middle_status': record.get('middle_status') or '',
                'csv_status': status,
                'candidates': candidates,
                'registered_matches': registered_matches,
            })
        return mismatched

    def _move_room_in_db(self, cursor, old_room, new_room):
        """部屋移動：登録内容・清掃予定をそのまま新しい部屋番号へ付け替える
        （commitは呼び出し側）"""
        cursor.execute("UPDATE rooms SET room_number = ? WHERE room_number = ?",
                       (new_room, old_room))
        cursor.execute("UPDATE cleaning_schedule SET room_number = ? WHERE room_number = ?",
                       (new_room, old_room))

    def _delete_room_in_db(self, cursor, room):
        """部屋の登録と清掃予定を削除する（commitは呼び出し側）"""
        cursor.execute("DELETE FROM rooms WHERE room_number = ?", (room,))
        cursor.execute("DELETE FROM cleaning_schedule WHERE room_number = ?", (room,))

    def _write_status_check_log(self, lines, csv_date):
        """不一致確認の実行結果をログファイルへ追記する。失敗しても処理は続ける"""
        try:
            with open(self.STATUS_CHECK_LOG_FILE, 'a', encoding='utf-8') as f:
                stamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                csv_text = csv_date.strftime('%Y-%m-%d') if csv_date else '不明'
                f.write(f"[{stamp}] 部屋状態CSV日付: {csv_text}\n")
                for line in lines:
                    f.write(f"  {line}\n")
                f.write("\n")
        except Exception as e:
            print(f"部屋状態確認ログの書き込みに失敗: {e}")

    def show_status_mismatch_dialog(self, mismatched, room_status_map,
                                    guest_name_map, csv_date=None):
        """DB登録があるのにCSVで在室でない部屋の確認ダイアログ。

        部屋ごとに「部屋移動（移動先の部屋番号へ登録を付け替え）」
        「削除」「そのまま」を選んで実行する。閉じるまで呼び出し元を待たせる。

        初期選択：
          ・予約CSVで同名の宿泊者の在室部屋（未登録）が見つかった → 部屋移動（移動先自動入力）
          ・それ以外 → 削除
        移動先がすでにDB登録済みの部屋には移動できない（二重登録防止）。
        """
        dialog = tk.Toplevel(self.root)
        dialog.title("部屋状態の確認")
        dialog.transient(self.root)
        dialog.grab_set()
        self._fit_toplevel_to_screen(dialog, 1050, 560, min_w=640, min_h=360)

        row_font = ("", 12)
        header_font = ("", 12, "bold")

        frame = ttk.Frame(dialog, padding="15")
        frame.pack(fill="both", expand=True)

        ttk.Label(frame, text="登録済みの部屋の状態確認", font=("", 16, "bold")).pack(pady=(0, 8))
        ttk.Label(frame,
                  text=f"DBに登録があるのに、部屋状態CSVで在室になっていない部屋が"
                       f"{len(mismatched)}件あります（部屋移動・アーリーC/Oなど）。\n"
                       "部屋ごとに処理を選んで「実行」を押してください。"
                       "「部屋移動」は登録内容と清掃予定をそのまま移動先へ付け替えます。",
                  font=("", 11), justify=tk.CENTER).pack(pady=(0, 10))

        # 移動先の選択肢：CSVで在室かつ未登録の部屋（手入力も可）
        free_in_house = sorted(room for room, status in room_status_map.items()
                               if status == '3' and room not in self.existing_rooms)

        def guest_name_of(room):
            info = guest_name_map.get(room, {})
            return info.get('name', '') if isinstance(info, dict) else (info or '')

        # --- 一覧（件数が多い場合に備えてスクロール可能に） ---
        list_outer = ttk.Frame(frame)
        list_outer.pack(fill="both", expand=True)
        canvas = tk.Canvas(list_outer, highlightthickness=0)
        scrollbar = ttk.Scrollbar(list_outer, orient="vertical", command=canvas.yview)
        table = ttk.Frame(canvas)
        table.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=table, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        headers = ("部屋番号", "登録者名", "登録区分", "CSV状態", "処理", "移動先", "備考")
        for col, text in enumerate(headers):
            ttk.Label(table, text=text, font=header_font).grid(
                row=0, column=col, padx=6, pady=(0, 4), sticky="w")

        row_ctrls = []
        for i, item in enumerate(mismatched, start=1):
            candidates = item['candidates']
            registered = item['registered_matches']

            notes = []
            if candidates:
                notes.append(f"同名の宿泊者を {'、'.join(candidates)} で検出")
            if registered:
                notes.append(f"同名の宿泊者が {'、'.join(registered)} に登録済み")
            if not candidates and not registered:
                notes.append("移動先の候補なし")

            default_action = (self.MISMATCH_ACTION_MOVE if candidates
                              else self.MISMATCH_ACTION_DELETE)
            action_var = tk.StringVar(value=default_action)
            dest_var = tk.StringVar(value=candidates[0] if candidates else '')

            ttk.Label(table, text=item['room'], font=row_font).grid(row=i, column=0, padx=6, pady=3, sticky="w")
            ttk.Label(table, text=item['guest'], font=row_font).grid(row=i, column=1, padx=6, pady=3, sticky="w")
            ttk.Label(table, text=item['middle_status'], font=row_font).grid(row=i, column=2, padx=6, pady=3, sticky="w")
            ttk.Label(table, text=self.CSV_STATUS_LABELS.get(item['csv_status'], item['csv_status']),
                      font=row_font).grid(row=i, column=3, padx=6, pady=3, sticky="w")

            action_combo = ttk.Combobox(
                table, textvariable=action_var, state='readonly', width=9, font=row_font,
                values=(self.MISMATCH_ACTION_MOVE, self.MISMATCH_ACTION_DELETE,
                        self.MISMATCH_ACTION_KEEP))
            action_combo.grid(row=i, column=4, padx=6, pady=3, sticky="w")

            # 候補（同名の宿泊者）を先頭に、残りの在室・未登録部屋を続ける
            dest_values = candidates + [r for r in free_in_house if r not in candidates]
            dest_combo = ttk.Combobox(table, textvariable=dest_var, width=8,
                                      font=row_font, values=dest_values)
            dest_combo.grid(row=i, column=5, padx=6, pady=3, sticky="w")

            ttk.Label(table, text=" ／ ".join(notes), font=("", 11),
                      foreground="red" if registered else "").grid(
                row=i, column=6, padx=6, pady=3, sticky="w")

            # ホイールで値が変わる事故防止（部屋選択ダイアログと同じ対策）
            for combo in (action_combo, dest_combo):
                combo.bind("<MouseWheel>", lambda e: "break")
                combo.bind("<Button-4>", lambda e: "break")
                combo.bind("<Button-5>", lambda e: "break")

            # 部屋移動以外では移動先を入力不可にする
            def sync_dest_state(*_, av=action_var, dc=dest_combo):
                dc.config(state='normal' if av.get() == self.MISMATCH_ACTION_MOVE
                          else 'disabled')
            action_var.trace('w', sync_dest_state)
            sync_dest_state()

            row_ctrls.append((item, action_var, dest_var))

        def close():
            dialog.grab_release()
            dialog.destroy()

        def execute():
            moves = []    # (item, 移動先)
            deletes = []  # item
            keeps = []    # item
            errors = []
            warnings = []
            used_dests = {}

            for item, action_var, dest_var in row_ctrls:
                room = item['room']
                action = action_var.get()
                if action == self.MISMATCH_ACTION_KEEP:
                    keeps.append(item)
                    continue
                if action == self.MISMATCH_ACTION_DELETE:
                    deletes.append(item)
                    continue

                dest = dest_var.get().strip()
                if not dest:
                    errors.append(f"{room}：移動先の部屋番号を入力してください。")
                    continue
                # 「305」のように0を省いて入力されても登録済みの表記（0305）に合わせる
                if dest not in room_status_map and dest.isdigit():
                    for known in room_status_map:
                        if known.isdigit() and int(known) == int(dest):
                            dest = known
                            break
                if dest == room:
                    errors.append(f"{room}：移動先が同じ部屋番号です。")
                    continue
                if dest in self.existing_rooms:
                    reg = next((r for r in self.records if r['room'] == dest), None)
                    reg_name = f"（{reg['guest']}）" if reg and reg.get('guest') else ''
                    errors.append(f"{room}：移動先 {dest} はすでにデータベースに登録済みです{reg_name}。"
                                  "先にそちらを確認してください。")
                    continue
                if dest in used_dests:
                    errors.append(f"{room}：移動先 {dest} は {used_dests[dest]} の移動先にも指定されています。")
                    continue
                used_dests[dest] = room

                dest_status = room_status_map.get(dest)
                if dest_status is None:
                    warnings.append(f"{room}→{dest}：移動先が部屋状態CSVにありません。")
                elif dest_status != '3':
                    warnings.append(f"{room}→{dest}：移動先がCSVで在室ではありません"
                                    f"（{self.CSV_STATUS_LABELS.get(dest_status, dest_status)}）。")
                dest_name = guest_name_of(dest)
                if (dest_name and item['guest']
                        and self._normalize_guest_name(dest_name)
                        != self._normalize_guest_name(item['guest'])):
                    warnings.append(f"{room}→{dest}：登録者名（{item['guest']}）と"
                                    f"移動先の宿泊者名（{dest_name}）が違います。")
                moves.append((item, dest))

            if errors:
                messagebox.showerror("入力エラー", "\n".join(errors), parent=dialog)
                return

            if not moves and not deletes:
                close()
                return

            # 確認ログ（実行内容の一覧）を表示して最終確認
            log_lines = []
            for item, dest in moves:
                log_lines.append(f"部屋移動: {item['room']} → {dest}（{item['guest']}／{item['middle_status']}）")
            for item in deletes:
                log_lines.append(f"削除: {item['room']}（{item['guest']}／{item['middle_status']}／"
                                 f"CSV状態: {self.CSV_STATUS_LABELS.get(item['csv_status'], item['csv_status'])}）")
            for item in keeps:
                log_lines.append(f"そのまま: {item['room']}（{item['guest']}）")

            msg = "以下の内容で実行します。\n\n" + "\n".join(log_lines)
            if warnings:
                msg += "\n\n【注意】\n" + "\n".join(warnings)
            msg += "\n\nよろしいですか？"
            if not messagebox.askyesno("実行内容の確認", msg, parent=dialog):
                return

            # 削除は元に戻せないため、実行前にDBをバックアップしておく
            # （バックアップ管理から復元できる）
            self.conn.commit()
            backup_name = self.create_database_backup_silent()
            log_lines.append(f"実行前バックアップ: {backup_name}")

            cursor = self.conn.cursor()
            try:
                for item, dest in moves:
                    self._move_room_in_db(cursor, item['room'], dest)
                for item in deletes:
                    self._delete_room_in_db(cursor, item['room'])
                self.conn.commit()
            except Exception as e:
                self.conn.rollback()
                messagebox.showerror("エラー", f"データベースの更新に失敗しました: {e}\n"
                                               "変更は取り消されました。", parent=dialog)
                return

            self._write_status_check_log(log_lines + [f"注意: {w}" for w in warnings], csv_date)

            # 付け替え・削除をメモリ上の一覧にも反映
            self.records.clear()
            self.existing_rooms.clear()
            self.load_data()
            self.update_room_count_display()

            close()
            messagebox.showinfo("完了",
                                f"部屋移動: {len(moves)}件\n削除: {len(deletes)}件\n"
                                f"そのまま: {len(keeps)}件\n\n"
                                f"実行内容は「{self.STATUS_CHECK_LOG_FILE}」に記録しました。\n"
                                f"実行前バックアップ: {backup_name}")

        button_frame = ttk.Frame(frame)
        button_frame.pack(pady=(10, 0))
        ttk.Button(button_frame, text="実行", command=execute).pack(side="left", padx=5)
        ttk.Button(button_frame, text="すべてそのまま（閉じる）",
                   command=close).pack(side="left", padx=5)
        dialog.protocol("WM_DELETE_WINDOW", close)

        dialog.wait_window()

    # エコプラン判定キーワード（NFKCで正規化してから部分一致で判定）
    ECO_PLAN_KEYWORDS = ['長期', 'ウィークリー', 'マンスリー', 'エコ']

    # 中日ステータスの選択肢（プルダウン用）
    NO_CLEAN_LABEL = '清掃不要・アメ不要'
    # 通常清掃（エコ対象に選ばなかった在室部屋）を表す中日ステータス。
    # ・チェックなしで自動登録された部屋 → 実C/I日起点の3日周期で〇
    # ・全日設定でこの値を明示選択した部屋 → 毎日清掃（全中日に〇）
    # どちらもアーニング表では従来どおり「連泊(stay)」扱いにする
    # （_load_db_registered_rooms で登録済み集合から除外している）。
    NORMAL_CLEAN_LABEL = '通常清掃'
    MIDDLE_STATUS_CHOICES = ('×', 'エコドア', NO_CLEAN_LABEL, NORMAL_CLEAN_LABEL)

    # エコ清掃扱いの中日ステータス（判定用）
    ECO_MIDDLE_STATUSES = ('×', 'エコドア', NO_CLEAN_LABEL)

    # 通常清掃の周期（日）と、清掃日に立てる記号。
    # 実C/I日を起点に NORMAL_CLEAN_CYCLE 日ごと（+3, +6, +9 …）に〇を立て、
    # それ以外の中日は cleaning_schedule に行を作らない（＝エコ票では空欄）。
    # 空文字ステータスの行を作ると cleanup_checkout_rooms の削除対象に
    # なってしまうため、非清掃日は「行なし」で表現すること。
    NORMAL_CLEAN_CYCLE = 3
    NORMAL_CLEAN_MARK = '〇'

    # アーニング表【指示書用】への出力区分と表示文字
    #   vacant  : 状態0  空室
    #   pre_ci  : 状態1  未チェックイン
    #   checkout: 状態2  チェックアウト → 空欄（記号なし）
    #   ecodoor : 状態3 かつ DBで is_ecodoor=True（ドアにエコ札）
    #   eco     : 状態3 かつ DBに登録あり・エコドア以外 ＝ エコ清掃
    #   noclean : 状態3 かつ DBで middle_status='清掃不要・アメ不要'
    #             （エコ清掃扱い：緑色塗り・eco丸数字にカウント）
    #   stay    : 状態3 かつ DBに登録なし（未登録含む）＝ 連泊
    EARNING_MARKS = {
        'vacant':   '×',
        'pre_ci':   '未C/I',
        'checkout': '',        # C/O は空欄表示にする
        'ecodoor':  'エコドア',
        'eco':      'エコ清掃',
        'noclean':  '清掃不要',
        'stay':     '',        # 連泊は空欄表示（部屋番号セルの緑色塗りは維持）
    }

    # エコ清掃・連泊の部屋番号セルを塗る背景色（medium sea green #3cb371）
    EARNING_GREEN_FILL = PatternFill(
        start_color="FFC6EFCE", end_color="FFC6EFCE", fill_type="solid"
    )

    # アーニング表テンプレートのファイル名（EcoRoomClean.py と同じフォルダに置く）
    # この名前で見つからない場合は、同フォルダ内の .xlsx を走査して
    # 【指示書用】シートを持つものを自動検出する。
    EARNING_TEMPLATE_NAME = "アーニング表.xlsx"

    @staticmethod
    def _is_ecoplan(name_field):
        """12列目（名前＋プラン情報）からエコプラン該当かを判定"""
        if not name_field:
            return False

        # 半角カナ・全角英数などをNFKCで正規化してから判定する。
        # これにより 'ﾏﾝｽﾘｰ'→'マンスリー'、'ｳｨｰｸﾘｰ'→'ウィークリー' となり、
        # 半角／全角どちらの表記でもヒットする。
        import unicodedata
        normalized = unicodedata.normalize('NFKC', name_field)

        # 日本語・カタカナキーワード（部分一致）
        for kw in HotelCleaningSystem.ECO_PLAN_KEYWORDS:
            if kw in normalized:
                return True

        # Eco / ECO / eco / Ｅｃｏ を大文字小文字を無視して判定
        if 'ECO' in normalized.upper():
            return True

        return False

    def load_checkout_map(self, file_path):
        """予約詳細CSV(202605291形式)から
        {予約番号: (チェックイン日datetime, チェックアウト日datetime, 泊数int)}
        のマップを作成。

        列構成: 列0=予約番号, 列4=チェックイン日(5列目), 列5=チェックアウト日(6列目),
                列6=泊数(7列目)。同一予約番号は複数行あるがC/O日は同一
        （前回検証済み: 食い違い0件）なので単純な辞書でよい。

        C/I日は通常清掃の3日周期（実C/I日起点）の計算に使う。
        読めなかった場合は None を返し、呼び出し側で登録起点日にフォールバックする。
        """
        co_map = {}
        encodings = ['cp932', 'shift_jis', 'utf-8']

        for encoding in encodings:
            try:
                with open(file_path, 'r', encoding=encoding) as f:
                    for row in csv.reader(f):
                        if len(row) < 7 or not row[0].strip().isdigit():
                            continue
                        co_raw = row[5].strip()
                        if not (co_raw.isdigit() and len(co_raw) == 8):
                            continue
                        try:
                            co_date = datetime.strptime(co_raw, '%Y%m%d')
                        except ValueError:
                            continue
                        # C/I日（列4）は取れないこともあるので個別に判定する
                        ci_raw = row[4].strip()
                        ci_date = None
                        if ci_raw.isdigit() and len(ci_raw) == 8:
                            try:
                                ci_date = datetime.strptime(ci_raw, '%Y%m%d')
                            except ValueError:
                                ci_date = None
                        nights_raw = row[6].strip()
                        nights = int(nights_raw) if nights_raw.isdigit() else None
                        co_map[row[0].strip()] = (ci_date, co_date, nights)
                return co_map
            except (UnicodeDecodeError, Exception):
                continue

        messagebox.showwarning(
            "警告",
            "予約詳細CSV（C/O日入り）の読み込みに失敗しました。\n"
            "チェックアウト日なしで続行します。"
        )
        return {}

    def load_guest_names_from_yoyaku(self, file_path, checkout_map=None):
        """予約CSVから部屋番号→{'name', 'is_ecoplan', 'checkin', 'checkout', 'nights'}
        のマッピングを作成。

        checkout_map（予約番号→(C/I日, C/O日, 泊数)）が渡された場合は、
        予約番号（列0）をキーに各部屋へC/I日・C/O日・泊数を紐づける。
        """
        if checkout_map is None:
            checkout_map = {}
        guest_map = {}
        # 複数のエンコーディングを試行
        encodings = ['cp932', 'shift_jis', 'utf-8']

        for encoding in encodings:
            try:
                with open(file_path, 'r', encoding=encoding) as f:
                    reader = csv.reader(f)
                    for row in reader:
                        if len(row) >= 12:
                            yoyaku_no = row[0].strip()  # 1列目：予約番号
                            room_number = row[10]  # 11列目：部屋番号
                            name_field = row[11]   # 12列目：名前＋プラン情報

                            # エコプラン判定（_is_ecoplan内部でNFKC正規化するため、
                            # 半角カナの元フィールドをそのまま渡してよい）
                            is_ecoplan = self._is_ecoplan(name_field)

                            # '_'の手前までが名前
                            if '_' in name_field:
                                guest_name = name_field.split('_')[0]
                            else:
                                guest_name = name_field

                            # 半角カナを全角カナに変換（可能であれば）
                            try:
                                import unicodedata
                                guest_name = unicodedata.normalize('NFKC', guest_name)
                            except Exception:
                                pass

                            # 予約番号からC/I日・C/O日・泊数を引く（無ければNone）
                            ci_date, co_date, nights = checkout_map.get(
                                yoyaku_no, (None, None, None))

                            guest_map[room_number] = {
                                'name': guest_name,
                                'is_ecoplan': is_ecoplan,
                                'checkin': ci_date,    # datetime or None
                                'checkout': co_date,   # datetime or None
                                'nights': nights,      # int or None
                            }
                return guest_map
            except (UnicodeDecodeError, Exception):
                continue

        messagebox.showwarning("警告", "予約CSVの読み込みに失敗しました。宿泊者名なしで続行します。")
        return {}

    def _fit_toplevel_to_screen(self, win, want_w, want_h,
                                min_w=480, min_h=360, margin=70):
        """Toplevelウィンドウの初期サイズを画面に収まる範囲に調整して中央寄せする。
        希望サイズ(want_w × want_h)が画面より大きい場合は自動で縮小するため、
        画面が小さい端末でも下部のボタンが見切れない。あわせてリサイズ可能にし、
        最小サイズも設定する。"""
        try:
            win.update_idletasks()
            screen_w = win.winfo_screenwidth()
            screen_h = win.winfo_screenheight()

            # 画面から余白(margin)を引いた範囲を上限とする
            max_w = max(min_w, screen_w - margin)
            max_h = max(min_h, screen_h - margin)

            w = min(want_w, max_w)
            h = min(want_h, max_h)

            # 中央（やや上寄せ＝タスクバー等を避ける）に配置
            x = max(0, (screen_w - w) // 2)
            y = max(0, (screen_h - h) // 3)

            win.geometry(f"{w}x{h}+{x}+{y}")
            win.resizable(True, True)
            win.minsize(min(min_w, w), min(min_h, h))
        except Exception:
            # 何かあっても最低限サイズだけは設定する
            win.geometry(f"{want_w}x{want_h}")
            win.resizable(True, True)

    def _resolve_stay_period(self, guest_info, base_date):
        """予約CSVから取れた情報と登録起点日（＝清掃日）から
        (C/I日, 泊数) を決めて返す。

        ・C/I日は予約詳細CSVの実C/I日を優先し、取れなければ登録起点日
        ・C/O日が取れていればそこまでの実日数、取れなければ「登録起点日+2日」
          までとする。実C/I日が過去の部屋でもC/Oが必ず登録起点日より後に
          来るようにするため（登録起点日以前がC/Oになると
          cleanup_checkout_rooms が次回CSV読込で即削除してしまう）。
        """
        if isinstance(guest_info, dict):
            ci_date = guest_info.get('checkin')
            co_date = guest_info.get('checkout')
        else:
            ci_date = None
            co_date = None

        checkin = ci_date or base_date
        if co_date and co_date > checkin:
            days = (co_date - checkin).days
        else:
            days = (base_date - checkin).days + 2

        return checkin, max(days, 1)

    def _build_room_schedule(self, checkin_date, days, middle_status,
                             cycle_anchor=None, daily_clean=False):
        """C/I日から days 泊分のスケジュール辞書 {'M/D': ステータス} を作る。

        ・初日 = 'C/I'、最終日（checkin_date + days）= 'C/O'
        ・3日周期の〇：実C/I日（cycle_anchor。無ければ checkin_date）を起点に
          NORMAL_CLEAN_CYCLE 日ごと（+3, +6, +9 …）の中日を清掃日とする。
          - 通常清掃      : 清掃日は〇、それ以外の中日はキーを作らない。
                            キーが無い日はエコ票で空欄になる。空文字ステータスの
                            行を作ると cleanup_checkout_rooms の削除対象に
                            なってしまうため、非清掃日は必ず「キー無し」で表現する
          - ×／エコドア   : 清掃日は〇、それ以外の中日は middle_status
          - 清掃不要      : 〇は立てず全中日 middle_status（清掃しない部屋のため）
        ・daily_clean=True（全日設定で「通常清掃」を明示選択した部屋）は
          毎日清掃：全中日に〇を立てる（通常清掃とのみ組み合わせて使う）

        cycle_anchor はエコ部屋用。エコ部屋の checkin_date は登録起点日
        （＝清掃日）のままなので、実C/I日を別途渡して周期を合わせる。
        """
        schedule = {}
        is_normal = (middle_status == self.NORMAL_CLEAN_LABEL)
        cycles = (middle_status in ('×', 'エコドア')) or is_normal
        anchor = cycle_anchor or checkin_date
        current = checkin_date

        for day in range(days + 1):
            date_str = f"{current.month}/{current.day}"
            offset = (current - anchor).days
            is_clean_day = (cycles and offset > 0
                            and offset % self.NORMAL_CLEAN_CYCLE == 0)
            if day == 0:
                schedule[date_str] = "C/I"
            elif day == days:
                schedule[date_str] = "C/O"
            elif is_normal:
                # 通常清掃：清掃日にだけ〇（毎日清掃なら全中日〇）
                if daily_clean or is_clean_day:
                    schedule[date_str] = self.NORMAL_CLEAN_MARK
            elif is_clean_day:
                # エコ部屋（×／エコドア）も清掃日は〇
                schedule[date_str] = self.NORMAL_CLEAN_MARK
            else:
                schedule[date_str] = middle_status
            current += timedelta(days=1)

        return schedule

    def show_csv_room_selection_dialog(self, eco_rooms, guest_name_map=None, csv_date=None):
        """CSVから読み込んだ部屋の選択ダイアログを表示"""
        if guest_name_map is None:
            guest_name_map = {}

        dialog = tk.Toplevel(self.root)
        dialog.title("エコ清掃部屋選択")
        dialog.transient(self.root)
        dialog.grab_set()
        # 画面に収まるサイズに自動調整（小さい画面での見切れ防止）＋リサイズ可能に
        self._fit_toplevel_to_screen(dialog, 900, 860, min_w=560, min_h=420)

        # このダイアログ内の文字を大きくする
        row_font = ("", 13)            # 部屋一覧の各行テキスト用
        header_font = ("", 12, "bold") # 一覧ヘッダー用
        dlg_style = ttk.Style(dialog)
        dlg_style.configure("Eco.TCheckbutton", font=row_font)
        dlg_style.configure("Eco.TButton", font=row_font)

        frame = ttk.Frame(dialog, padding="15")
        frame.pack(fill="both", expand=True)

        # タイトル
        title_label = ttk.Label(frame, text="エコ清掃対象部屋の選択", font=("", 18, "bold"))
        title_label.pack(pady=(0, 10))

        # 説明
        info_label = ttk.Label(frame,
                               text=f"CSVから{len(eco_rooms)}件のエコ清掃対象部屋が見つかりました。\n"
                                    "チェックを入れた部屋を登録します"
                                    "（C/O日あり→実泊数／なし→2泊）。",
                               font=("", 13), justify=tk.CENTER)
        info_label.pack(pady=(0, 10))

        # チェックイン日選択
        checkin_frame = ttk.LabelFrame(frame, text="チェックイン日", padding="10")
        checkin_frame.pack(fill="x", pady=(0, 10))

        # デフォルトはCSV 1行目1列目の日付を優先、なければ今日
        default_date = csv_date if csv_date else datetime.now()
        checkin_vars = {
            'year': tk.StringVar(value=str(default_date.year)),
            'month': tk.StringVar(value=str(default_date.month)),
            'day': tk.StringVar(value=str(default_date.day))
        }

        date_input_frame = ttk.Frame(checkin_frame)
        date_input_frame.pack()

        ttk.Entry(date_input_frame, textvariable=checkin_vars['year'], width=6).pack(side="left")
        ttk.Label(date_input_frame, text="年").pack(side="left", padx=(0, 5))

        ttk.Entry(date_input_frame, textvariable=checkin_vars['month'], width=4).pack(side="left")
        ttk.Label(date_input_frame, text="月").pack(side="left", padx=(0, 5))

        ttk.Entry(date_input_frame, textvariable=checkin_vars['day'], width=4).pack(side="left")
        ttk.Label(date_input_frame, text="日").pack(side="left")

        # 全選択/全解除ボタン
        select_frame = ttk.Frame(frame)
        select_frame.pack(fill="x", pady=(0, 10))

        # チェックボックス用の変数を格納
        check_vars = {}
        # ステータス選択用の変数を格納
        status_vars = {}
        # 強制エコ（チェックを外せない）部屋の集合
        forced_eco_rooms = set()
        # 行ごとのウィジェット参照（編集保存後の行表示更新用）
        row_ctrls = {}

        def select_all():
            for var in check_vars.values():
                var.set(True)

        def deselect_all():
            # 強制エコの部屋は全解除でも外さない
            for room, var in check_vars.items():
                if room not in forced_eco_rooms:
                    var.set(False)

        ttk.Button(select_frame, text="全選択", command=select_all, style="Eco.TButton").pack(side="left", padx=3)
        ttk.Button(select_frame, text="全解除", command=deselect_all, style="Eco.TButton").pack(side="left", padx=3)

        # 選択数表示ラベル
        count_label = ttk.Label(select_frame, text="選択: 0件", font=row_font)
        count_label.pack(side="right", padx=5)

        def update_count(*args):
            selected = sum(1 for var in check_vars.values() if var.get())
            count_label.config(text=f"選択: {selected}件")

        # 登録／キャンセルボタンの枠を先に最下部へ固定しておく。
        # （list_frame より前に pack することで、ウィンドウが低くても
        #   ボタンが見切れず、代わりに一覧側がスクロールで縮む）
        button_frame = ttk.Frame(frame)
        button_frame.pack(side="bottom", fill="x", pady=(10, 0))

        # スクロール可能なフレーム
        list_frame = ttk.LabelFrame(frame, text="部屋一覧", padding="10")
        list_frame.pack(fill="both", expand=True, pady=(0, 15))

        canvas = tk.Canvas(list_frame, width=800, height=440, bg="white")
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)

        scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        # bind_all はアプリ全体にかかるため、このダイアログ以外
        # （部屋編集ダイアログ等の別ウィンドウ）上のイベントには反応させない
        def _is_own_event(event):
            try:
                return event.widget.winfo_toplevel() is dialog
            except (AttributeError, tk.TclError):
                return False

        # マウスホイールスクロール（ダイアログのどこにカーソルがあっても効く）
        def on_mousewheel(event):
            if not _is_own_event(event):
                return
            # canvasがまだ存在するか確認
            if not canvas.winfo_exists():
                return
            if platform.system() == "Darwin":
                canvas.yview_scroll(int(-1 * event.delta), "units")
            else:
                canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
            return "break"

        # Linux (X11) ではホイールが Button-4 / Button-5 として届く
        def on_button4(event):
            if not _is_own_event(event):
                return
            if not canvas.winfo_exists():
                return
            canvas.yview_scroll(-1, "units")
            return "break"

        def on_button5(event):
            if not _is_own_event(event):
                return
            if not canvas.winfo_exists():
                return
            canvas.yview_scroll(1, "units")
            return "break"

        # ダイアログ全体でホイールイベントを拾う
        dialog.bind_all("<MouseWheel>", on_mousewheel)
        dialog.bind_all("<Button-4>", on_button4)
        dialog.bind_all("<Button-5>", on_button5)

        # ダイアログ終了時にバインドを解除（メイン画面に副作用を残さない）
        def cleanup_bindings():
            try:
                dialog.unbind_all("<MouseWheel>")
                dialog.unbind_all("<Button-4>")
                dialog.unbind_all("<Button-5>")
            except Exception:
                pass

        dialog.protocol("WM_DELETE_WINDOW",
                        lambda: (cleanup_bindings(), dialog.destroy()))

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # 列幅定数（pixel単位）— ヘッダーと各行で共通使用して列を揃える
        # 順: 選択, 登録状況, 部屋番号, 宿泊者名, C/O, 状態, 全日設定, 編集
        COL_MINSIZE = (55, 100, 100, 170, 80, 110, 160, 80)

        def _configure_row_columns(frame):
            """全ての行フレームで同じ列幅を共有させる"""
            for i, w in enumerate(COL_MINSIZE):
                frame.columnconfigure(i, minsize=w, weight=0)

        # ヘッダー
        header_frame = ttk.Frame(scrollable_frame)
        header_frame.pack(fill="x", pady=(0, 5))
        _configure_row_columns(header_frame)

        ttk.Label(header_frame, text="選択", font=header_font, anchor="w").grid(row=0, column=0, padx=5, sticky="w")
        ttk.Label(header_frame, text="登録状況", font=header_font, anchor="w").grid(row=0, column=1, padx=5, sticky="w")
        ttk.Label(header_frame, text="部屋番号", font=header_font, anchor="w").grid(row=0, column=2, padx=5, sticky="w")
        ttk.Label(header_frame, text="宿泊者名", font=header_font, anchor="w").grid(row=0, column=3, padx=5, sticky="w")
        ttk.Label(header_frame, text="C/O", font=header_font, anchor="w").grid(row=0, column=4, padx=5, sticky="w")
        ttk.Label(header_frame, text="状態", font=header_font, anchor="w").grid(row=0, column=5, padx=5, sticky="w")
        ttk.Label(header_frame, text="全日設定", font=header_font, anchor="w").grid(row=0, column=6, padx=5, sticky="w")
        ttk.Label(header_frame, text="編集", font=header_font, anchor="w").grid(row=0, column=7, padx=5, sticky="w")

        def refresh_row(room_number):
            """編集画面での保存後に、その部屋の行表示（チェック・登録状況・
            全日設定プルダウン）を最新の登録状態に合わせて描き直す。
            赤チェック（登録済みかつ清掃日翌日が×／エコドア／清掃不要）の
            付け外しもここで反映される。"""
            ctrl = row_ctrls.get(room_number)
            if not ctrl or not ctrl['row_frame'].winfo_exists():
                return

            is_existing = (room_number in self.existing_rooms
                           or any(r['room'] == room_number
                                  for r in self.records))
            try:
                base_date = datetime(int(checkin_vars['year'].get()),
                                     int(checkin_vars['month'].get()),
                                     int(checkin_vars['day'].get()))
            except ValueError:
                base_date = default_date
            is_existing_eco = is_existing and self._is_eco_on_date(
                room_number, base_date + timedelta(days=1))

            if is_existing_eco:
                reg_status_text = "登録済み(エコ)"
            elif is_existing:
                reg_status_text = "登録済み"
            elif room_number in forced_eco_rooms:
                reg_status_text = "自動エコ"
            else:
                reg_status_text = "未登録"
            ctrl['reg_status_lbl'].config(
                text=reg_status_text,
                foreground="red" if is_existing_eco else "")

            # 列0のチェック表示を作り直す（赤☑ラベル⇔チェックボックス）
            check_var = ctrl['check_var']
            if is_existing:
                check_var.set(False)
            ctrl['check_widget'].destroy()
            if is_existing_eco:
                widget = ttk.Label(ctrl['row_frame'], text="☑",
                                   foreground="red", font=("", 16))
            else:
                widget = ttk.Checkbutton(ctrl['row_frame'],
                                         variable=check_var)
                if is_existing or room_number in forced_eco_rooms:
                    widget.config(state='disabled')
            widget.grid(row=0, column=0, padx=5, sticky="w")
            ctrl['check_widget'] = widget

            # 登録済みになったらプルダウンを保存値に合わせてロックする
            record = next((r for r in self.records
                           if r['room'] == room_number), None)
            if is_existing and record:
                ctrl['status_var'].set(record.get('middle_status') or '×')
            ctrl['status_combo'].config(
                state='disabled' if is_existing else 'readonly')

        def open_room_editor(room_number, guest_info, status_var):
            """行右端の「編集」ボタン：その部屋の編集画面を開く。

            まだDB登録前の部屋でも開けるよう、行の内容から仮レコードを作って
            渡す。編集画面で保存するとその時点でDBに登録される
            （その後この選択ダイアログで登録ボタンを押しても二重登録はしない）。
            保存後は refresh_row でこの一覧の該当行を描き直す。
            """
            try:
                base_date = datetime(int(checkin_vars['year'].get()),
                                     int(checkin_vars['month'].get()),
                                     int(checkin_vars['day'].get()))
            except ValueError:
                base_date = default_date

            checkin, days = self._resolve_stay_period(guest_info, base_date)
            middle_status = status_var.get() or '×'

            if isinstance(guest_info, dict):
                guest_name = guest_info.get('name', '')
                is_ecoplan_guest = guest_info.get('is_ecoplan', False)
            else:
                guest_name = guest_info or ''
                is_ecoplan_guest = False

            self.open_edit_dialog(room_number, default_record={
                'room': room_number,
                'guest': guest_name,
                'date': checkin,
                'days': days,
                'ecodoor': (middle_status == 'エコドア'),
                'ecoplan': is_ecoplan_guest,
                'middle_status': middle_status,
            }, on_saved=lambda: refresh_row(room_number))

        # 部屋ごとのチェックボックス
        for idx, room_info in enumerate(eco_rooms):
            room_number = room_info['room']
            status = room_info['status']
            guest_info = guest_name_map.get(room_number, {})
            # 後方互換：万一文字列が入っていてもクラッシュしないように
            if isinstance(guest_info, dict):
                guest_name = guest_info.get('name', '')
                is_ecoplan_guest = guest_info.get('is_ecoplan', False)
                checkout_date = guest_info.get('checkout')
            else:
                guest_name = guest_info or ''
                is_ecoplan_guest = False
                checkout_date = None

            # C/O列の表示（例: 6/1）。C/O日が取れていない部屋は空欄
            checkout_text = (f"{checkout_date.month}/{checkout_date.day}"
                             if checkout_date else '')

            # 「状態」列の表示：エコプラン該当者は「エコプラン」、非該当者は空欄
            display_status = 'エコプラン' if is_ecoplan_guest else ''

            # 強制エコ判定：エコプラン該当（プラン名に 長期／ウィークリー／
            # マンスリー／エコ）だけを自動チェックし、チェックを外せなくする。
            # それ以外のエコ認識はC/I翌日（部屋が初めて一覧に出る日）に
            # 現場がチェックを付けて行う。チェック＝DBに ×／エコドア／
            # 清掃不要・アメ不要 として登録され、以降は「登録済み」のまま
            # 再登録不要（変更は行右端の「編集」から）。
            is_forced_eco = is_ecoplan_guest

            row_frame = ttk.Frame(scrollable_frame)
            row_frame.pack(fill="x", pady=2)
            _configure_row_columns(row_frame)

            # 既存チェック
            is_existing = room_number in self.existing_rooms or any(r['room'] == room_number for r in self.records)
            # 登録済みかつ、ダイアログのチェックイン日（CSVの清掃日）の
            # 翌日のセルが ×／エコドア／清掃不要 の部屋は赤チェックで
            # 見分けられるようにする（翌日が〇＝3日周期の清掃日の部屋は黒）
            is_existing_eco = is_existing and self._is_eco_on_date(
                room_number, default_date + timedelta(days=1))
            if is_existing_eco:
                reg_status_text = "登録済み(エコ)"
            elif is_existing:
                reg_status_text = "登録済み"
            elif is_forced_eco:
                reg_status_text = "自動エコ"
            else:
                reg_status_text = "未登録"

            # 初期状態：強制エコの部屋だけチェック済み、それ以外は解除
            check_var = tk.BooleanVar(value=(is_forced_eco and not is_existing))
            check_vars[room_number] = check_var
            check_var.trace('w', update_count)
            if is_forced_eco and not is_existing:
                forced_eco_rooms.add(room_number)

            if is_existing_eco:
                # macOSではチェックボックスの色変更が効かないため、
                # 赤い「☑」ラベルで代用する（登録済み行はもともと操作不可）
                check_widget = ttk.Label(row_frame, text="☑", foreground="red",
                                         font=("", 16))
                check_widget.grid(row=0, column=0, padx=5, sticky="w")
            else:
                check_widget = ttk.Checkbutton(row_frame, variable=check_var)
                check_widget.grid(row=0, column=0, padx=5, sticky="w")
                # 登録済み＝操作不可、強制エコ＝チェックを外せない
                if is_existing or is_forced_eco:
                    check_widget.config(state='disabled')

            # 登録状況はチェックボックスの隣（列1）に表示する
            reg_status_lbl = ttk.Label(
                row_frame, text=reg_status_text, anchor="w", font=row_font,
                foreground="red" if is_existing_eco else "")
            reg_status_lbl.grid(row=0, column=1, padx=5, sticky="w")

            ttk.Label(row_frame, text=room_number, anchor="w", font=row_font).grid(row=0, column=2, padx=5, sticky="w")
            ttk.Label(row_frame, text=guest_name, anchor="w", font=row_font).grid(row=0, column=3, padx=5, sticky="w")
            ttk.Label(row_frame, text=checkout_text, anchor="w", font=row_font).grid(row=0, column=4, padx=5, sticky="w")
            ttk.Label(row_frame, text=display_status, anchor="w", font=row_font).grid(row=0, column=5, padx=5, sticky="w")

            # 中日ステータス選択プルダウン（デフォルトは「×」）
            # ※以前コンボボックスでホイールにより値が変わる事故があったため、
            #   readonly＋ホイールイベント無効化で対策している
            status_var = tk.StringVar(value='×')
            status_vars[room_number] = status_var
            status_combo = ttk.Combobox(
                row_frame, textvariable=status_var,
                values=self.MIDDLE_STATUS_CHOICES,
                state='readonly', width=16, font=row_font)
            status_combo.grid(row=0, column=6, padx=5, sticky="w")

            # コンボボックス上でのホイール操作を無効化（値の誤変更防止）
            status_combo.bind("<MouseWheel>", lambda e: "break")
            status_combo.bind("<Button-4>", lambda e: "break")
            status_combo.bind("<Button-5>", lambda e: "break")

            # プルダウンの選択＝エコ認識としてチェックを連動させる。
            # ×／エコドア／清掃不要 を選んだらチェックON（エコ登録の意思表示）、
            # 通常清掃を選んだらチェックOFF（未チェック＝通常清掃として登録）。
            # 強制エコ（エコプラン）の行はチェック固定のため連動させない
            def on_status_selected(event, var=check_var, sv=status_var,
                                   rn=room_number):
                if rn in forced_eco_rooms:
                    return
                var.set(sv.get() != self.NORMAL_CLEAN_LABEL)
            status_combo.bind('<<ComboboxSelected>>', on_status_selected)

            if is_existing:
                status_combo.config(state='disabled')

            # 編集保存後の行表示更新（refresh_row）用に参照を保持する
            row_ctrls[room_number] = {
                'row_frame': row_frame,
                'check_widget': check_widget,
                'check_var': check_var,
                'reg_status_lbl': reg_status_lbl,
                'status_var': status_var,
                'status_combo': status_combo,
            }

            # 行右端の「編集」ボタン：部屋編集画面を開く。
            # 強制エコを通常清掃に落としたい場合もここから変更する。
            ttk.Button(
                row_frame, text="編集", width=6, style="Eco.TButton",
                command=lambda rn=room_number, gi=guest_info,
                               sv=status_var: open_room_editor(rn, gi, sv)
            ).grid(row=0, column=7, padx=5, sticky="w")

            # 2部屋ごとに区切り線を引いて見やすくする
            # （最後の行の後ろには線を入れない）
            if (idx + 1) % 2 == 0 and (idx + 1) < len(eco_rooms):
                ttk.Separator(scrollable_frame, orient="horizontal").pack(
                    fill="x", padx=2, pady=3
                )

        # 初期カウント更新
        update_count()

        # ボタン（枠は上部で最下部固定済み。ここでは中身のボタンだけ追加する）
        def register_rooms():
            """チェックした部屋をエコ清掃として、チェックしなかった部屋を
            通常清掃（実C/I日起点の3日周期）として登録する"""
            selected_rooms = [room for room, var in check_vars.items() if var.get()]

            # 選択されたチェックイン日（＝登録起点日・清掃日）を取得
            try:
                checkin_date = datetime(
                    int(checkin_vars['year'].get()),
                    int(checkin_vars['month'].get()),
                    int(checkin_vars['day'].get())
                )
            except ValueError:
                messagebox.showerror("エラー", "正しいチェックイン日を入力してください。")
                return

            def is_already_registered(room_number):
                """DB登録済み、またはこの実行中に既に登録した部屋か"""
                return (room_number in self.existing_rooms
                        or any(r['room'] == room_number for r in self.records))

            def add_record(room_number, guest_info, middle_status,
                           room_checkin, days, daily_clean=False):
                """レコードを組み立てて self.records に追加する。

                3日周期〇の起点（cycle_anchor）には実C/I日を使う。
                エコ部屋は room_checkin が登録起点日のままなので、
                予約CSVから実C/I日が取れていればそちらを優先する。
                daily_clean=True は毎日清掃（全中日〇）。"""
                if isinstance(guest_info, dict):
                    guest_name = guest_info.get('name', '')
                    is_ecoplan_guest = guest_info.get('is_ecoplan', False)
                    real_checkin = guest_info.get('checkin')
                else:
                    guest_name = guest_info or ''
                    is_ecoplan_guest = False
                    real_checkin = None

                record = {
                    'room': room_number,
                    'guest': guest_name,
                    'date': room_checkin,
                    'days': days,
                    'ecodoor': (middle_status == 'エコドア'),
                    'ecoplan': is_ecoplan_guest,
                    'middle_status': middle_status,
                    'schedule': self._build_room_schedule(
                        room_checkin, days, middle_status,
                        cycle_anchor=real_checkin,
                        daily_clean=daily_clean),
                    'is_new': True
                }
                self.records.append(record)
                self.existing_rooms.add(room_number)
                newly_registered.append(record)

            eco_count = 0
            normal_count = 0
            daily_count = 0
            newly_registered = []   # この操作で登録した部屋（即時DB保存用）

            # --- チェックした部屋：エコ清掃として登録 ---
            for room_number in selected_rooms:
                if is_already_registered(room_number):
                    continue

                # 選択されたステータスを取得
                # （プルダウン: ×／エコドア／清掃不要・アメ不要／通常清掃）
                middle_status = status_vars[room_number].get() or '×'
                guest_info = guest_name_map.get(room_number, {})

                if middle_status == self.NORMAL_CLEAN_LABEL:
                    # 全日設定で通常清掃を明示選択した部屋は毎日清掃（全中日〇）
                    room_checkin, days = self._resolve_stay_period(
                        guest_info, checkin_date)
                    add_record(room_number, guest_info, middle_status,
                               room_checkin, days, daily_clean=True)
                    daily_count += 1
                    continue
                else:
                    # 泊数の決定：
                    #   予約詳細CSVからC/O日が取れていれば
                    #   「登録起点日（=清掃日）〜C/O日」で実際の残り日数を使う。
                    #   取れていない場合は従来どおり2泊固定。
                    checkout_date = (guest_info.get('checkout')
                                     if isinstance(guest_info, dict) else None)
                    room_checkin = checkin_date
                    if checkout_date and checkout_date > checkin_date:
                        days = (checkout_date - checkin_date).days
                    else:
                        days = 2

                add_record(room_number, guest_info, middle_status,
                           room_checkin, days)
                eco_count += 1

            # --- チェックしなかった部屋：通常清掃として自動登録 ---
            # エコ対象に選ばなかった在室部屋は、実C/I日を起点に
            # NORMAL_CLEAN_CYCLE 日ごとの通常清掃（〇）としてDBに登録する。
            # 全日設定で「通常清掃」を明示選択した部屋（選択するとチェックが
            # 外れる）は毎日清掃として全中日に〇を立てる。
            # アーニング表ではどちらも従来どおり連泊(stay)扱いのまま
            # （_load_db_registered_rooms が通常清掃を除外している）。
            for room_info in eco_rooms:
                room_number = room_info['room']
                if room_number in selected_rooms or is_already_registered(room_number):
                    continue

                guest_info = guest_name_map.get(room_number, {})
                room_checkin, days = self._resolve_stay_period(
                    guest_info, checkin_date)

                daily = (status_vars[room_number].get()
                         == self.NORMAL_CLEAN_LABEL)
                add_record(room_number, guest_info, self.NORMAL_CLEAN_LABEL,
                           room_checkin, days, daily_clean=daily)
                if daily:
                    daily_count += 1
                else:
                    normal_count += 1

            # この時点で即DBへ保存する（「エコ票作成」の押し忘れ対策）。
            # 失敗した場合は is_new を残し、「エコ票作成」時に再保存を試みる
            save_error = None
            if newly_registered:
                try:
                    self._save_records_to_db(newly_registered)
                    for r in newly_registered:
                        r['is_new'] = False
                except Exception as e:
                    save_error = e

            cleanup_bindings()
            dialog.destroy()

            if save_error is not None:
                messagebox.showerror(
                    "エラー",
                    f"データベースへの保存に失敗しました: {save_error}\n\n"
                    "登録内容はメモリ上に残っています。\n"
                    "「エコ票作成」を押すと再保存を試みます。")
            elif eco_count > 0 or normal_count > 0 or daily_count > 0:
                messagebox.showinfo("完了",
                                    f"データベースに登録しました。\n\n"
                                    f"エコ清掃: {eco_count}件\n"
                                    f"通常清掃: {normal_count}件"
                                    f"（実C/I日から{self.NORMAL_CLEAN_CYCLE}日ごとに〇）\n"
                                    f"毎日清掃: {daily_count}件（全中日に〇）\n"
                                    f"登録起点日: {checkin_date.strftime('%Y年%m月%d日')}\n\n"
                                    "エコ票（Excel）が必要な場合は「エコ票作成」を押してください。")
            else:
                messagebox.showinfo("情報", "新規登録された部屋はありませんでした。")

            self.update_room_count_display()

        ttk.Button(button_frame, text="選択した部屋を登録", command=register_rooms, style="Eco.TButton").pack(side="left", padx=5)
        ttk.Button(button_frame, text="キャンセル", style="Eco.TButton",
                   command=lambda: (cleanup_bindings(), dialog.destroy())).pack(side="left", padx=5)

        # 一覧の必要幅を実測してウィンドウ幅を再調整（右端「編集」列の見切れ防止）
        dialog.update_idletasks()
        need_w = (scrollable_frame.winfo_reqwidth()
                  + scrollbar.winfo_reqwidth() + 40)  # 外枠padding15×2 + 余裕
        self._fit_toplevel_to_screen(dialog, max(900, need_w), 860,
                                     min_w=560, min_h=420)

    def _save_records_to_db(self, records_to_save):
        """レコードを rooms / cleaning_schedule テーブルへ保存して commit する。

        「選択した部屋を登録」時と「エコ票作成」時の共通処理。
        空欄ステータスの日は行を作らない（空文字を保存すると
        cleanup_checkout_rooms が部屋ごと削除してしまうため）。"""
        cursor = self.conn.cursor()
        for record in records_to_save:
            cursor.execute(
                '''INSERT OR REPLACE INTO rooms
                   (room_number, guest_name, check_in_date, cleaning_days,
                    is_ecodoor, is_ecoplan, middle_status)
                   VALUES (?, ?, ?, ?, ?, ?, ?)''',
                (record['room'], record['guest'],
                 record['date'].strftime('%Y-%m-%d'),
                 record['days'], record['ecodoor'], record['ecoplan'],
                 record.get('middle_status',
                            'エコドア' if record['ecodoor'] else '×')))

            cursor.execute("DELETE FROM cleaning_schedule WHERE room_number = ?", (record['room'],))
            for date_str, status in record['schedule'].items():
                if not status:
                    continue
                month, day = map(int, date_str.split('/'))
                year = record['date'].year
                if month < record['date'].month:
                    year += 1
                cleaning_date = datetime(year, month, day)
                cursor.execute("INSERT INTO cleaning_schedule VALUES (?, ?, ?)",
                               (record['room'], cleaning_date.strftime('%Y-%m-%d'), status))

        self.conn.commit()

    def create_schedule(self):
        """エコ票（Excel）を作成する。
        DB登録は「選択した部屋を登録」時に済んでいるが、
        保存に失敗したレコード（is_new のまま）があればここで再保存する。"""
        if not self.records:
            messagebox.showerror("エラー", "記録する部屋がありません")
            return

        try:
            # 未保存レコードがあれば保存（通常は登録時に保存済み）
            new_records = [r for r in self.records if r.get('is_new', True)]
            if new_records:
                self._save_records_to_db(new_records)

            # Excel生成
            self.generate_excel()

            # データ再読み込み
            self.records.clear()
            self.existing_rooms.clear()
            self.load_data()

            # 新しいレコードのフラグをクリア
            for record in self.records:
                record['is_new'] = False

            self.update_room_count_display()

            messagebox.showinfo("完了",
                                f"エコ票を作成しました。\n"
                                f"対象部屋: {len(self.records)}件")

            # Excel ファイルを開く
            self.open_excel()

            # 続けてアーニング表も自動出力する（直前のCSV読込で使った部屋状態CSVを再利用）。
            # 部屋状態CSVのパスが無い場合は、output_earning_table 内で選択ダイアログが出る。
            self.output_earning_table(self.last_room_status_csv)

        except Exception as e:
            messagebox.showerror("エラー", f"処理中にエラーが発生しました: {e}")

    def generate_excel(self):
        wb = openpyxl.Workbook()
        wb.remove(wb.active)

        # ソート
        self.records.sort(key=lambda x: int(x['room']) if x['room'].isdigit() else float('inf'))

        # 全レコードから使用されている月を収集
        months_used = set()
        for record in self.records:
            for date_str in record['schedule'].keys():
                month = int(date_str.split('/')[0])
                months_used.add(month)

        # 月ごとにシートを作成
        for month in sorted(months_used):
            month_name = f"{month}月"
            ws = wb.create_sheet(title=month_name)

            # ヘッダー
            ws.cell(3, 1, "氏名")
            ws.cell(3, 3, "部屋番号")

            # その月の日数を取得
            import calendar
            year = self.records[0]['date'].year if self.records else datetime.now().year
            days_in_month = calendar.monthrange(year, month)[1]

            # 日付ヘッダー
            for day in range(1, days_in_month + 1):
                ws.cell(3, 3 + day, str(day))

            ws.cell(3, 3 + days_in_month + 1, "エコプラン")

            # この月にスケジュールがある部屋のみをフィルタ
            records_this_month = []
            for record in self.records:
                has_schedule_this_month = any(
                    int(date_str.split('/')[0]) == month
                    for date_str in record['schedule'].keys()
                )
                if has_schedule_this_month:
                    records_this_month.append(record)

            # データ
            for i, record in enumerate(records_this_month):
                row = 4 + i
                ws.cell(row, 1, record['guest'])
                ws.cell(row, 3, record['room'])
                ws.cell(row, 3 + days_in_month + 1, "エコプラン" if record['ecoplan'] else "")

                # この月のスケジュールのみ
                for date_str, status in record['schedule'].items():
                    if int(date_str.split('/')[0]) == month:
                        day = int(date_str.split('/')[1])
                        if day <= days_in_month:
                            ws.cell(row, 3 + day, status)

            # 列幅調整
            ws.column_dimensions['A'].width = 12
            ws.column_dimensions['C'].width = 8
            for day in range(1, days_in_month + 1):
                col_letter = openpyxl.utils.get_column_letter(3 + day)
                ws.column_dimensions[col_letter].width = 6

        wb.save(self.excel_file)
        wb.close()

    def edit_room(self):
        """部屋の編集"""
        self.show_room_edit_dialog()

    def show_room_edit_dialog(self, select_room=None):
        """部屋選択ダイアログを表示（編集用）。

        編集画面を閉じる（保存・キャンセル）とこの一覧に戻り、続けて別の部屋を
        編集できる。select_room を渡すとその部屋を選択・表示した状態で開く。
        """
        all_rooms = list(self.existing_rooms) + [r['room'] for r in self.records]
        all_rooms = sorted(set(all_rooms), key=lambda x: int(x) if x.isdigit() else float('inf'))

        if not all_rooms:
            messagebox.showinfo("情報", "編集できる部屋がありません")
            return

        dialog = tk.Toplevel(self.root)
        dialog.title("部屋編集")
        dialog.geometry("250x400")
        dialog.resizable(False, False)
        dialog.transient(self.root)
        dialog.grab_set()

        ttk.Label(dialog, text="編集する部屋を選択してください：").pack(pady=10)

        listbox = tk.Listbox(dialog, selectmode=tk.SINGLE)
        for room in all_rooms:
            record = self.find_room_record(room)
            if record:
                display_text = f"{room} - {record['guest']} ({record['days']}日)"
            else:
                display_text = f"{room} - (詳細不明)"
            listbox.insert(tk.END, display_text)
        listbox.pack(fill="both", expand=True, padx=10, pady=5)

        # 直前に編集した部屋を選択状態にして見える位置までスクロール
        if select_room in all_rooms:
            idx = all_rooms.index(select_room)
            listbox.selection_set(idx)
            listbox.activate(idx)
            listbox.see(idx)

        button_frame = ttk.Frame(dialog)
        button_frame.pack(pady=10)

        def on_edit():
            selection = listbox.curselection()
            if selection:
                selected_room = all_rooms[selection[0]]
                dialog.destroy()
                # 編集画面を閉じたらこの一覧に戻る（続けて別の部屋を編集できるように）
                self.open_edit_dialog(
                    selected_room,
                    on_closed=lambda: self.show_room_edit_dialog(
                        select_room=selected_room))
            else:
                messagebox.showwarning("警告", "部屋を選択してください。")

        ttk.Button(button_frame, text="編集", command=on_edit).pack(side="left", padx=5)
        ttk.Button(button_frame, text="キャンセル", command=dialog.destroy).pack(side="left", padx=5)

        # ダブルクリック・Enterキーでも編集画面を開く
        listbox.bind("<Double-Button-1>", lambda e: on_edit())
        listbox.bind("<Return>", lambda e: on_edit())
        listbox.focus_set()

    def find_room_record(self, room_number):
        """部屋番号からレコードを検索"""
        for record in self.records:
            if record['room'] == room_number:
                return record

        # データベースから検索
        cursor = self.conn.cursor()
        cursor.execute(
            "SELECT room_number, guest_name, check_in_date, cleaning_days, "
            "is_ecodoor, is_ecoplan, middle_status "
            "FROM rooms WHERE room_number = ?", (room_number,))
        row = cursor.fetchone()
        if row:
            room, guest, date_str, days, ecodoor, ecoplan, middle_status = row
            if not middle_status:
                middle_status = 'エコドア' if ecodoor else '×'
            return {
                'room': room,
                'guest': guest,
                'date': datetime.strptime(date_str, '%Y-%m-%d'),
                'days': days,
                'ecodoor': bool(ecodoor),
                'ecoplan': bool(ecoplan),
                'middle_status': middle_status
            }
        return None

    def open_edit_dialog(self, room_number, default_record=None,
                         on_saved=None, on_closed=None):
        """編集ダイアログを開く。

        default_record を渡すと、DB未登録の部屋でもその内容を初期値として
        編集画面を開ける（CSV選択ダイアログの行右端「編集」ボタン用）。
        保存した時点でDBに登録される。
        on_saved を渡すと、保存成功後（DBコミット後）に呼び出す。
        呼び出し元の一覧表示を保存内容に合わせて更新する用途。
        on_closed を渡すと、ダイアログを閉じた後（保存なら完了メッセージの後、
        またはキャンセル・×ボタン）に呼び出す。部屋編集の一覧へ戻る用途。
        """
        record = self.find_room_record(room_number) or default_record
        if not record:
            messagebox.showerror("エラー", f"部屋番号 {room_number} の情報が見つかりません")
            return

        dialog = tk.Toplevel(self.root)
        dialog.title(f"部屋 {room_number} の編集")
        dialog.geometry("560x740")
        dialog.resizable(False, False)
        dialog.transient(self.root)
        dialog.grab_set()

        frame = ttk.Frame(dialog, padding="10")
        frame.pack(fill="both", expand=True)

        # 編集用変数
        edit_vars = {}

        # 部屋番号（編集不可）
        ttk.Label(frame, text="部屋番号：").grid(row=0, column=0, sticky="w", pady=5)
        ttk.Label(frame, text=room_number, font=("", 10, "bold")).grid(row=0, column=1, sticky="w", pady=5)

        # お客様名
        ttk.Label(frame, text="お客様名：").grid(row=1, column=0, sticky="w", pady=5)
        edit_vars['guest'] = tk.StringVar(value=record['guest'])
        ttk.Entry(frame, textvariable=edit_vars['guest'], width=30).grid(row=1, column=1, sticky="w", pady=5)

        # チェックイン日
        ttk.Label(frame, text="チェックイン日：").grid(row=2, column=0, sticky="w", pady=5)
        checkin_frame = ttk.Frame(frame)
        checkin_frame.grid(row=2, column=1, sticky="w", pady=5)

        edit_vars['checkin_year'] = tk.StringVar(value=str(record['date'].year))
        edit_vars['checkin_month'] = tk.StringVar(value=str(record['date'].month))
        edit_vars['checkin_day'] = tk.StringVar(value=str(record['date'].day))

        ttk.Entry(checkin_frame, textvariable=edit_vars['checkin_year'], width=6).pack(side="left")
        ttk.Label(checkin_frame, text="年").pack(side="left")
        ttk.Entry(checkin_frame, textvariable=edit_vars['checkin_month'], width=4).pack(side="left")
        ttk.Label(checkin_frame, text="月").pack(side="left")
        ttk.Entry(checkin_frame, textvariable=edit_vars['checkin_day'], width=4).pack(side="left")
        ttk.Label(checkin_frame, text="日").pack(side="left")

        # チェックアウト日
        ttk.Label(frame, text="チェックアウト日：").grid(row=3, column=0, sticky="w", pady=5)
        checkout_frame = ttk.Frame(frame)
        checkout_frame.grid(row=3, column=1, sticky="w", pady=5)

        checkout_date = record['date'] + timedelta(days=record['days'])
        edit_vars['checkout_year'] = tk.StringVar(value=str(checkout_date.year))
        edit_vars['checkout_month'] = tk.StringVar(value=str(checkout_date.month))
        edit_vars['checkout_day'] = tk.StringVar(value=str(checkout_date.day))

        ttk.Entry(checkout_frame, textvariable=edit_vars['checkout_year'], width=6).pack(side="left")
        ttk.Label(checkout_frame, text="年").pack(side="left")
        ttk.Entry(checkout_frame, textvariable=edit_vars['checkout_month'], width=4).pack(side="left")
        ttk.Label(checkout_frame, text="月").pack(side="left")
        ttk.Entry(checkout_frame, textvariable=edit_vars['checkout_day'], width=4).pack(side="left")
        ttk.Label(checkout_frame, text="日").pack(side="left")

        # 宿泊日数（自動計算・表示のみ）
        ttk.Label(frame, text="宿泊日数：").grid(row=4, column=0, sticky="w", pady=5)
        days_label = ttk.Label(frame, text=f"{record['days']}日", font=("", 10, "bold"))
        days_label.grid(row=4, column=1, sticky="w", pady=5)

        # 中日ステータス（×／エコドア／清掃不要・アメ不要）
        ttk.Label(frame, text="全日設定：").grid(row=5, column=0, sticky="w", pady=5)
        current_middle = record.get('middle_status') or (
            'エコドア' if record['ecodoor'] else '×')
        edit_vars['middle_status'] = tk.StringVar(value=current_middle)
        middle_combo = ttk.Combobox(
            frame, textvariable=edit_vars['middle_status'],
            values=self.MIDDLE_STATUS_CHOICES, state='readonly', width=16)
        middle_combo.grid(row=5, column=1, sticky="w", pady=5)
        # ホイールによる値の誤変更を防止
        middle_combo.bind("<MouseWheel>", lambda e: "break")
        middle_combo.bind("<Button-4>", lambda e: "break")
        middle_combo.bind("<Button-5>", lambda e: "break")

        # エコプラン
        ttk.Label(frame, text="プラン：").grid(row=6, column=0, sticky="w", pady=5)
        edit_vars['ecoplan'] = tk.BooleanVar(value=record['ecoplan'])
        ttk.Checkbutton(frame, text="エコプラン", variable=edit_vars['ecoplan']).grid(row=6, column=1, sticky="w",
                                                                                      pady=5)

        # エコ再開日（この日から3日周期を引き直す）
        # 指定した日がエコ（中日ステータス）の1日目になり、〇はその2日後・5日後 …
        # に立つ。前日を「直前の清掃日」とみなすので×は2日連続に収まる。
        # 指定日より前の日と C/I・C/O 行は一切触らない。
        ttk.Label(frame, text="エコ再開日：").grid(row=7, column=0, sticky="w", pady=5)
        reassign_frame = ttk.Frame(frame)
        reassign_frame.grid(row=7, column=1, sticky="w", pady=5)

        # 初期値：今日が中日ならその日（＝昨日まで清掃済みで今日からエコ、の想定）、
        # 範囲外ならC/I日の翌日（＝標準の周期と同じ結果になる日）
        today = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
        if record['date'] < today < checkout_date:
            default_anchor = today
        else:
            default_anchor = record['date'] + timedelta(days=1)

        edit_vars['reassign_year'] = tk.StringVar(value=str(default_anchor.year))
        edit_vars['reassign_month'] = tk.StringVar(value=str(default_anchor.month))
        edit_vars['reassign_day'] = tk.StringVar(value=str(default_anchor.day))

        ttk.Entry(reassign_frame, textvariable=edit_vars['reassign_year'], width=6).pack(side="left")
        ttk.Label(reassign_frame, text="年").pack(side="left")
        ttk.Entry(reassign_frame, textvariable=edit_vars['reassign_month'], width=4).pack(side="left")
        ttk.Label(reassign_frame, text="月").pack(side="left")
        ttk.Entry(reassign_frame, textvariable=edit_vars['reassign_day'], width=4).pack(side="left")
        ttk.Label(reassign_frame, text="日").pack(side="left")
        # 実行／解除ボタンは do_reassign・clear_reassign の定義後に
        # この reassign_frame へ追加する

        # 個別スケジュール編集
        ttk.Label(frame, text="スケジュール編集：").grid(row=8, column=0, sticky="nw", pady=5)
        schedule_frame = ttk.Frame(frame)
        schedule_frame.grid(row=8, column=1, sticky="ew", pady=5)

        # スケジュール編集用のScrollable Frame
        canvas = tk.Canvas(schedule_frame, width=400, height=250, bg="white")
        scrollbar = ttk.Scrollbar(schedule_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)

        # スクロール設定
        def configure_scroll_region(event):
            canvas.configure(scrollregion=canvas.bbox("all"))

        def on_mousewheel(event):
            if not canvas.winfo_exists():
                return
            if platform.system() == "Darwin":
                canvas.yview_scroll(int(-1 * event.delta), "units")
            else:  # Windows/Linux
                canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
            # "break" で伝播を止め、背後のダイアログの bind_all まで届かせない
            return "break"

        def on_button4(event):
            if not canvas.winfo_exists():
                return
            canvas.yview_scroll(-1, "units")
            return "break"

        def on_button5(event):
            if not canvas.winfo_exists():
                return
            canvas.yview_scroll(1, "units")
            return "break"

        # マウスホイールバインド
        def bind_mousewheel(widget):
            widget.bind("<MouseWheel>", on_mousewheel)
            if platform.system() != "Darwin":  # Linux (X11)
                widget.bind("<Button-4>", on_button4)
                widget.bind("<Button-5>", on_button5)

        bind_mousewheel(canvas)
        bind_mousewheel(scrollable_frame)

        def on_canvas_click(event):
            canvas.focus_set()

        canvas.bind("<Button-1>", on_canvas_click)
        canvas.bind("<Configure>", configure_scroll_region)
        scrollable_frame.bind("<Configure>", configure_scroll_region)

        canvas_window = canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        def configure_canvas(event):
            canvas.itemconfig(canvas_window, width=event.width)

        canvas.bind('<Configure>', configure_canvas)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # スケジュール項目の辞書
        schedule_vars = {}

        # 振り直しのエコ再開日。日付や全日設定を変えて
        # スケジュールを描き直すときも同じ日で引き直せるよう覚えておく
        reassign_state = {'anchor': None}

        # {'M/D': datetime}。'M/D' キーは年を持たないため、
        # エコ再開日との前後比較用に実日付を別途保持する
        schedule_dates = {}

        def apply_reassign(anchor, checkin, checkout, middle_now):
            """anchor（エコ再開日）以降の中日を3日周期で引き直す。

            anchor はエコ（中日ステータス）を数え直す最初の日。その前日を
            「直前の清掃日」とみなして cycle_anchor に渡すので、〇は
            anchor+2, anchor+5, anchor+8 … に立ち、×は2日連続に収まる
            （anchor をそのまま cycle_anchor にすると anchor 自身も×になり
            ×が3日連続してしまう）。
            anchor より前の日と C/I・C/O 行は触らない。
            「現在」列のラベルは元の値を示すためあえて更新しない
            （手動でコンボボックスを変えたときの挙動と揃える）。
            """
            pattern = self._build_room_schedule(
                checkin, (checkout - checkin).days, middle_now,
                cycle_anchor=anchor - timedelta(days=1))

            for date_str, var in schedule_vars.items():
                target = schedule_dates.get(date_str)
                if target is None or target < anchor:
                    continue  # エコ再開日より前は据え置き
                if target == checkin or target == checkout:
                    continue  # C/I・C/O は固定
                # パターンにキーが無い日は空欄（＝その日は行を作らない）
                var.set(pattern.get(date_str, ''))

        # スケジュール表示を更新する関数
        def update_schedule_display():
            # 既存のウィジェットを削除
            for widget in scrollable_frame.winfo_children():
                widget.destroy()

            schedule_vars.clear()

            try:
                if not all([edit_vars['checkin_year'].get(), edit_vars['checkin_month'].get(),
                            edit_vars['checkin_day'].get(), edit_vars['checkout_year'].get(),
                            edit_vars['checkout_month'].get(), edit_vars['checkout_day'].get()]):
                    error_label = tk.Label(scrollable_frame, text="日付を入力してください", fg="red")
                    error_label.pack()
                    return

                checkin = datetime(
                    int(edit_vars['checkin_year'].get()),
                    int(edit_vars['checkin_month'].get()),
                    int(edit_vars['checkin_day'].get())
                )
                checkout = datetime(
                    int(edit_vars['checkout_year'].get()),
                    int(edit_vars['checkout_month'].get()),
                    int(edit_vars['checkout_day'].get())
                )

                if checkout <= checkin:
                    error_label = tk.Label(scrollable_frame, text="無効な日付範囲", fg="red")
                    error_label.pack()
                    return

                days = (checkout - checkin).days
                days_label.config(text=f"{days}日")

                # 現在のスケジュールを取得
                current_schedule = self.get_room_schedule(room_number)

                # メインコンテナフレーム
                main_container = ttk.Frame(scrollable_frame)
                main_container.pack(fill="both", expand=True, padx=5, pady=5)

                # ヘッダー
                header_frame = ttk.Frame(main_container)
                header_frame.pack(fill="x", pady=(0, 5))

                ttk.Label(header_frame, text="日付", font=("", 9, "bold")).grid(row=0, column=0, padx=5, pady=2,
                                                                                sticky="w")
                ttk.Label(header_frame, text="現在", font=("", 9, "bold")).grid(row=0, column=1, padx=5, pady=2,
                                                                                sticky="w")
                ttk.Label(header_frame, text="変更", font=("", 9, "bold")).grid(row=0, column=2, padx=5, pady=2,
                                                                                sticky="w")

                # データフレーム
                data_frame = ttk.Frame(main_container)
                data_frame.pack(fill="both", expand=True)

                # 3日周期を反映した既定パターンを生成する。
                # ・DBに行が無い日（未登録の部屋・C/O延長で増えた日）は
                #   このパターンから〇／×等を補完する
                # ・ダイアログで中日ステータスを変更した場合は、保存済みの
                #   内容よりパターンを優先して全中日を引き直す
                #   （変更していなければ手動編集済みの保存内容を優先）
                middle_now = edit_vars['middle_status'].get() or '×'
                original_middle = record.get('middle_status') or (
                    'エコドア' if record['ecodoor'] else '×')
                pattern = self._build_room_schedule(
                    checkin, (checkout - checkin).days, middle_now)
                repattern = (middle_now != original_middle)

                # 日付範囲でスケジュール表示
                current_date = checkin
                row = 0
                while current_date <= checkout:
                    date_str = f"{current_date.month}/{current_date.day}"

                    # 行フレームを作成
                    row_frame = ttk.Frame(data_frame)
                    row_frame.pack(fill="x", pady=1)

                    # 日付表示
                    date_label = ttk.Label(row_frame, text=date_str, width=8)
                    date_label.grid(row=0, column=0, padx=5, pady=2, sticky="w")

                    # 現在の状態を取得（パターンにキーが無い日は空欄＝何も付けない。
                    # 空欄をDBに保存すると cleanup_checkout_rooms の削除対象に
                    # なるため、保存時に空欄の日は行を作らない）
                    if current_date == checkin:
                        current_status = "C/I"
                    elif current_date == checkout:
                        current_status = "C/O"
                    elif repattern:
                        current_status = pattern.get(date_str, '')
                    else:
                        current_status = current_schedule.get(
                            date_str, pattern.get(date_str, ''))

                    # 現在の状態表示
                    status_label = ttk.Label(row_frame, text=current_status, width=8)
                    status_label.grid(row=0, column=1, padx=5, pady=2, sticky="w")

                    # 変更用コンボボックス（空欄＝その日は何も付けない）
                    schedule_vars[date_str] = tk.StringVar(value=current_status)
                    schedule_dates[date_str] = current_date
                    status_combo = ttk.Combobox(row_frame, textvariable=schedule_vars[date_str], width=16)
                    status_combo['values'] = ('', 'C/I', 'C/O', '〇', '×', 'エコドア',
                                              self.NO_CLEAN_LABEL)
                    status_combo.grid(row=0, column=2, padx=5, pady=2, sticky="w")
                    # ホイールによる値の誤変更を防止
                    status_combo.bind("<MouseWheel>", lambda e: "break")
                    status_combo.bind("<Button-4>", lambda e: "break")
                    status_combo.bind("<Button-5>", lambda e: "break")

                    # C/IとC/Oは固定（変更不可）
                    if current_date == checkin or current_date == checkout:
                        status_combo.config(state='disabled')

                    # マウスホイールをバインド
                    bind_mousewheel(row_frame)
                    bind_mousewheel(date_label)
                    bind_mousewheel(status_label)

                    current_date += timedelta(days=1)
                    row += 1

                # 全てのウィジェットにマウスホイールをバインド
                bind_mousewheel(main_container)
                bind_mousewheel(header_frame)
                bind_mousewheel(data_frame)

                # スクロール範囲を更新
                def update_scroll_region():
                    scrollable_frame.update_idletasks()
                    canvas.update_idletasks()
                    canvas.configure(scrollregion=canvas.bbox("all"))
                    canvas.yview_moveto(0)

                canvas.after(100, update_scroll_region)

                # 長期間の場合はスクロールバーを目立たせる
                if days > 10:
                    info_label = tk.Label(main_container,
                                          text=f"※{days}日間の長期滞在 - マウスホイールまたはスクロールバーで全日程を確認してください",
                                          fg="blue", font=("", 8), wraplength=350)
                    info_label.pack(pady=5)
                    bind_mousewheel(info_label)

                # エコ再開日が設定済みなら、描き直した内容に再適用する
                # （C/O日を延ばした場合も同じ起点で周期が続くようにするため）
                anchor = reassign_state['anchor']
                if anchor and checkin < anchor < checkout:
                    apply_reassign(anchor, checkin, checkout, middle_now)

            except ValueError:
                days_label.config(text="日付エラー")
                error_label = tk.Label(scrollable_frame, text="正しい日付を入力してください", fg="red")
                error_label.pack()

        # 日数自動計算機能
        def update_days():
            update_schedule_display()

        # 日付変更時の自動更新
        for var in ['checkin_year', 'checkin_month', 'checkin_day', 'checkout_year', 'checkout_month', 'checkout_day']:
            edit_vars[var].trace('w', lambda *args: update_days())

        # 中日ステータス変更時もスケジュール表示を更新
        # （3日周期の〇パターンを引き直すため）
        edit_vars['middle_status'].trace('w', lambda *args: update_days())

        def do_reassign():
            """「この日から振り直し」ボタン：エコ再開日以降の中日を引き直す"""
            try:
                checkin = datetime(
                    int(edit_vars['checkin_year'].get()),
                    int(edit_vars['checkin_month'].get()),
                    int(edit_vars['checkin_day'].get())
                )
                checkout = datetime(
                    int(edit_vars['checkout_year'].get()),
                    int(edit_vars['checkout_month'].get()),
                    int(edit_vars['checkout_day'].get())
                )
            except ValueError:
                messagebox.showerror(
                    "エラー", "正しいチェックイン日／チェックアウト日を入力してください",
                    parent=dialog)
                return

            try:
                anchor = datetime(
                    int(edit_vars['reassign_year'].get()),
                    int(edit_vars['reassign_month'].get()),
                    int(edit_vars['reassign_day'].get())
                )
            except ValueError:
                messagebox.showerror(
                    "エラー", "正しいエコ再開日を入力してください", parent=dialog)
                return

            # C/O日以降は振り直す中日が無いので不可。C/I日は指定可
            # （C/I日をエコ1日目とみなし、〇がC/I日の2日後から立つ。
            #   C/I行自体は apply_reassign で固定のまま）
            if not (checkin <= anchor < checkout):
                messagebox.showerror(
                    "エラー",
                    "エコ再開日はチェックイン日〜チェックアウト日の"
                    "前日で指定してください",
                    parent=dialog)
                return

            reassign_state['anchor'] = anchor
            apply_reassign(anchor, checkin, checkout,
                           edit_vars['middle_status'].get() or '×')

        def clear_reassign():
            """「解除」ボタン：振り直しを取り消して既定の表示に戻す"""
            reassign_state['anchor'] = None
            update_schedule_display()

        ttk.Button(reassign_frame, text="この日から振り直し",
                   command=do_reassign).pack(side="left", padx=(8, 2))
        ttk.Button(reassign_frame, text="解除",
                   command=clear_reassign).pack(side="left")

        # 初期表示
        update_schedule_display()

        # ボタン
        button_frame = ttk.Frame(frame)
        button_frame.grid(row=9, column=0, columnspan=2, pady=20)

        def save_changes():
            try:
                checkin_date = datetime(
                    int(edit_vars['checkin_year'].get()),
                    int(edit_vars['checkin_month'].get()),
                    int(edit_vars['checkin_day'].get())
                )
                checkout_date = datetime(
                    int(edit_vars['checkout_year'].get()),
                    int(edit_vars['checkout_month'].get()),
                    int(edit_vars['checkout_day'].get())
                )

                if checkout_date <= checkin_date:
                    messagebox.showerror("エラー", "チェックアウト日はチェックイン日より後である必要があります")
                    return

                calculated_days = (checkout_date - checkin_date).days

            except ValueError:
                messagebox.showerror("エラー", "正しい日付を入力してください")
                return

            # 更新されたレコードを作成
            new_middle_status = edit_vars['middle_status'].get() or '×'
            updated_record = {
                'room': room_number,
                'guest': edit_vars['guest'].get().strip(),
                'date': checkin_date,
                'days': calculated_days,
                'ecodoor': (new_middle_status == 'エコドア'),
                'ecoplan': edit_vars['ecoplan'].get(),
                'middle_status': new_middle_status,
                'schedule': {},
                'is_new': True
            }

            # 個別に編集されたスケジュールを使用
            # 空欄の日は登録しない（＝その日は何も付けない）。
            # 空文字で保存すると cleanup_checkout_rooms の削除対象になる
            for date_str, var in schedule_vars.items():
                status = var.get()
                if status:
                    updated_record['schedule'][date_str] = status

            # メモリ上のレコードを更新
            for i, rec in enumerate(self.records):
                if rec['room'] == room_number:
                    self.records[i] = updated_record
                    break
            else:
                self.records.append(updated_record)

            # データベースに保存
            cursor = self.conn.cursor()
            cursor.execute(
                '''INSERT OR REPLACE INTO rooms
                   (room_number, guest_name, check_in_date, cleaning_days,
                    is_ecodoor, is_ecoplan, middle_status)
                   VALUES (?, ?, ?, ?, ?, ?, ?)''',
                (updated_record['room'], updated_record['guest'],
                 updated_record['date'].strftime('%Y-%m-%d'),
                 updated_record['days'], updated_record['ecodoor'],
                 updated_record['ecoplan'], updated_record['middle_status']))

            cursor.execute("DELETE FROM cleaning_schedule WHERE room_number = ?", (room_number,))
            for date_str, status in updated_record['schedule'].items():
                month, day = map(int, date_str.split('/'))
                year = updated_record['date'].year
                if month < updated_record['date'].month:
                    year += 1
                cleaning_date = datetime(year, month, day)
                cursor.execute("INSERT INTO cleaning_schedule VALUES (?, ?, ?)",
                               (room_number, cleaning_date.strftime('%Y-%m-%d'), status))

            self.conn.commit()

            dialog.destroy()
            if on_saved:
                on_saved()
            messagebox.showinfo("成功", f"部屋 {room_number} の情報を更新しました")
            if on_closed:
                on_closed()

        def cancel_changes():
            dialog.destroy()
            if on_closed:
                on_closed()

        dialog.protocol("WM_DELETE_WINDOW", cancel_changes)

        ttk.Button(button_frame, text="保存", command=save_changes).pack(side="left", padx=5)
        ttk.Button(button_frame, text="キャンセル", command=cancel_changes).pack(side="left", padx=5)

    def get_room_schedule(self, room_number):
        """部屋のスケジュールを取得"""
        # メモリから検索
        for record in self.records:
            if record['room'] == room_number:
                return record.get('schedule', {})

        # データベースから検索
        cursor = self.conn.cursor()
        cursor.execute(
            "SELECT cleaning_date, cleaning_status FROM cleaning_schedule WHERE room_number = ? ORDER BY cleaning_date",
            (room_number,))
        schedule = {}
        for date_str, status in cursor.fetchall():
            date = datetime.strptime(date_str, '%Y-%m-%d')
            schedule[f"{date.month}/{date.day}"] = status

        return schedule

    def _is_eco_on_date(self, room_number, target_date):
        """登録済み部屋の target_date のセルがエコ清掃
        （×/エコドア/清掃不要）になっているかを判定。
        〇（3日周期の清掃日）やC/I・C/Oは該当しない。"""
        schedule = self.get_room_schedule(room_number)
        return schedule.get(f"{target_date.month}/{target_date.day}") in self.ECO_MIDDLE_STATUSES

    def _find_earning_template(self):
        """アーニング表テンプレートを自動的に探して、そのパスを返す。
        見つからなければ None。
        探索順:
          1) 既定ファイル名（EARNING_TEMPLATE_NAME）をカレント／スクリプトの
             フォルダから探す
          2) 同フォルダ内の .xlsx を走査し、【指示書用】シート（シート名に
             「指示書」を含む）を2つ以上持つものを自動検出
        ※ .xlsm（指示書最新版などのマクロ付きファイル）は誤検出を避けるため
          自動検出の対象から除外する。
        """
        # 探索対象フォルダ（カレント と スクリプトと同じフォルダ）
        search_dirs = []
        cwd = os.getcwd()
        search_dirs.append(cwd)
        try:
            script_dir = os.path.dirname(os.path.abspath(__file__))
            if script_dir and script_dir not in search_dirs:
                search_dirs.append(script_dir)
        except Exception:
            pass

        # 1) 既定ファイル名で探す
        for d in search_dirs:
            cand = os.path.join(d, self.EARNING_TEMPLATE_NAME)
            if os.path.exists(cand):
                return cand

        # 2) .xlsx を走査して【指示書用】シートを持つものを自動検出
        for d in search_dirs:
            try:
                for fname in sorted(os.listdir(d)):
                    if not fname.lower().endswith('.xlsx'):
                        continue
                    path = os.path.join(d, fname)
                    try:
                        wb = openpyxl.load_workbook(path, read_only=True)
                        shidousho = [s for s in wb.sheetnames if '指示書' in s]
                        wb.close()
                        if len(shidousho) >= 2:
                            return path
                    except Exception:
                        continue
            except Exception:
                continue

        return None

    def _load_db_registered_rooms(self):
        """DBの rooms テーブルに登録されている部屋番号(int)について
        (登録済み集合, エコドア集合, 清掃不要集合) のタプルを返す。

        ※ 判定方針：状態3の部屋が登録済み集合に含まれていれば
          「エコドア」「エコ清掃」「清掃不要」のいずれか、
          含まれていなければ「○（連泊）」。
          middle_status='清掃不要・アメ不要' → 清掃不要、
          is_ecodoor=True → エコドア、それ以外の登録部屋 → エコ清掃。

        ※ middle_status='通常清掃' の部屋は rooms に登録されていても
          エコ扱いにはしない（アーニング表では従来どおり連泊(stay)として
          空欄・緑塗りで出す）。そのため登録済み集合には加えない。"""
        registered = set()
        ecodoor = set()
        noclean = set()
        cursor = self.conn.cursor()
        cursor.execute(
            "SELECT room_number, is_ecodoor, middle_status FROM rooms")
        for room_number, is_ecodoor, middle_status in cursor.fetchall():
            try:
                rn = int(room_number)
            except (TypeError, ValueError):
                continue
            # 通常清掃の部屋はエコ判定の対象外（連泊扱いのまま）
            if middle_status == self.NORMAL_CLEAN_LABEL:
                continue
            registered.add(rn)
            if middle_status == self.NO_CLEAN_LABEL:
                noclean.add(rn)
            elif is_ecodoor:
                ecodoor.add(rn)
        return registered, ecodoor, noclean

    def _build_instruction_cell_map(self, wb):
        """【指示書用】シートを走査し {部屋番号(int): (worksheet, 行, 列)} を返す。
        シート名に「指示書」を含むシートが対象。列レイアウト(A/D/G)に依存せず、
        3桁以上の数字セルを部屋番号とみなして検出する。"""
        cell_map = {}
        for sheet_name in wb.sheetnames:
            if '指示書' not in sheet_name:
                continue
            ws = wb[sheet_name]
            for row in ws.iter_rows():
                for cell in row:
                    v = cell.value
                    rn = None
                    if isinstance(v, int) and v >= 200:
                        rn = v
                    elif isinstance(v, str) and v.strip().isdigit() and len(v.strip()) >= 3:
                        rn = int(v.strip())
                    if rn is not None:
                        cell_map[rn] = (ws, cell.row, cell.column)
        return cell_map

    # 階見出しサマリーのフォントサイズ（結合セル幅に収めるため少し小さめ）
    FLOOR_SUMMARY_FONT_SIZE = 12
    # 階数（○F）部分のフォント設定（見やすく大きめ・別フォント）
    FLOOR_LABEL_FONT_NAME = 'メイリオ'
    FLOOR_LABEL_FONT_SIZE = 14
    # 階数の前後に入れる空き（全角スペース2つ分）
    FLOOR_LABEL_GAP = '　　'

    def _write_floor_summaries(self, wb, room_class):
        """【指示書用】シートの各階見出しセル（「2階」など）を、その階の
        部屋情報「S:n,Tn,丸数字　○F　空室:n」の表記に書き換える。
        例：エコが3室なら「S:16,T1,③　3F　空室:10」

        ・S/T … 在室中（○連泊）・チェックアウト(C/O)・未C/Iの部屋を
          カウントする。空室・エコ清掃・エコドアは含めない。
          部屋タイプ（S, NS, T, NT, D, ND, ANS, ANT, AND など）の
          末尾文字で判定し、D（ダブル）はシングル(S)に含める。
        ・丸数字 … eco部屋数＝「エコ清掃」＋「エコドア」の合計を
          ①②③…の丸数字で表示（0のときは⓪）
        ・空室 … 区分「×（状態0）」になった部屋数
        ・○F  … 元の見出し「2階」を「2F」表記に変換して中央に配置

        room_class: {部屋番号(int): 区分キー('eco','vacant','stay'など)}
        ※ 部屋は「同じ列で直上にある階見出し」に割り当てるため、
          列レイアウト(A/D/G)が変わっても動作する。
        """
        import re
        heading_pat = re.compile(r'^\d+階$')

        for sheet_name in wb.sheetnames:
            if '指示書' not in sheet_name:
                continue
            ws = wb[sheet_name]

            headings = []   # (列, 行)
            rooms = []      # (列, 行, 部屋番号int)
            for row in ws.iter_rows():
                for cell in row:
                    v = cell.value
                    if isinstance(v, str) and heading_pat.match(v.strip()):
                        headings.append((cell.column, cell.row))
                        continue
                    rn = None
                    if isinstance(v, int) and v >= 200:
                        rn = v
                    elif isinstance(v, str) and v.strip().isdigit() and len(v.strip()) >= 3:
                        rn = int(v.strip())
                    if rn is not None:
                        rooms.append((cell.column, cell.row, rn))

            # 各部屋を「同じ列で直上にある見出し」に割り当てて集計
            stats = {}  # (列, 行) -> カウンタ
            for col, r, rn in rooms:
                cands = [h for h in headings if h[0] == col and h[1] < r]
                if not cands:
                    continue
                hk = max(cands, key=lambda x: x[1])
                st = stats.setdefault(hk, {'S': 0, 'T': 0,
                                           'eco': 0, 'vacant': 0})
                # 当日の区分を取得
                key = room_class.get(rn)

                # S/T … 在室中（○連泊）・チェックアウト・未C/Iの部屋を
                # カウントする。空室・エコ清掃・エコドアは含めない。
                # 部屋タイプは部屋番号の右隣セルの末尾文字で判定し、
                # D（ダブル）はシングル(S)に含める。
                if key in ('stay', 'checkout', 'pre_ci'):
                    t = ws.cell(r, col + 1).value
                    if isinstance(t, str) and t.strip():
                        last = t.strip().upper()[-1]
                        if last in ('S', 'D'):
                            st['S'] += 1
                        elif last == 'T':
                            st['T'] += 1

                # eco（丸数字）＝エコ清掃＋エコドア＋清掃不要（エコ清掃扱い）、
                # 空室＝×（状態0）
                if key in ('eco', 'ecodoor', 'noclean'):
                    st['eco'] += 1
                elif key == 'vacant':
                    st['vacant'] += 1

            # 見出しセルに書き込み（例：「S:16,T1,③　　3F　　空室:10」）
            # eco数は丸数字（①②③…）で表示し、
            # 階数（○F）だけをリッチテキストで大きく・別フォントにする
            for (col, r), st in stats.items():
                cell = ws.cell(r, col)
                base = str(cell.value).strip() if cell.value else ''
                # 「2階」→「2F」表記に変換
                m = heading_pat.match(base)
                floor_label = f"{base[:-1]}F" if m else base
                eco_mark = self._circled_number(st['eco'])
                left = f"S:{st['S']},T{st['T']},{eco_mark}{self.FLOOR_LABEL_GAP}"
                right = f"{self.FLOOR_LABEL_GAP}空室:{st['vacant']}"
                small_font = InlineFont(sz=self.FLOOR_SUMMARY_FONT_SIZE, b=True)
                big_font = InlineFont(rFont=self.FLOOR_LABEL_FONT_NAME,
                                      sz=self.FLOOR_LABEL_FONT_SIZE, b=True)
                cell.value = CellRichText(
                    TextBlock(small_font, left),
                    TextBlock(big_font, floor_label),
                    TextBlock(small_font, right),
                )
                # ベースのセルフォント（リッチテキストが効かない環境向けの保険）
                f = cell.font
                cell.font = Font(name=f.name, bold=True,
                                 size=self.FLOOR_SUMMARY_FONT_SIZE)
                # 大きくした階数フォントが縦に切れないよう行の高さを確保
                need_h = self.FLOOR_LABEL_FONT_SIZE + 8
                cur_h = ws.row_dimensions[r].height
                if cur_h is None or cur_h < need_h:
                    ws.row_dimensions[r].height = need_h

    @staticmethod
    def _circled_number(n):
        """0〜50を丸数字（⓪①②…⑳㉑…㊿）に変換する。範囲外は数字のまま返す"""
        if n == 0:
            return '⓪'
        if 1 <= n <= 20:
            return chr(0x2460 + n - 1)   # ①〜⑳
        if 21 <= n <= 35:
            return chr(0x3251 + n - 21)  # ㉑〜㉟
        if 36 <= n <= 50:
            return chr(0x32B1 + n - 36)  # ㊱〜㊿
        return str(n)

    def output_earning_table(self, csv_path=None):
        """部屋状態CSVとアーニング表テンプレートから、【指示書用】シートの
        各部屋番号セルの2つ隣に区分（×／未C/I／C/O／エコ清掃／○）を書き込み、
        日付入りの別ファイルとして出力して開く。

        csv_path を渡すとそのCSVを使い、CSV選択ダイアログを出さない
        （エコ票作成からの自動実行で、直前に読み込んだCSVを再利用するため）。
        渡されない／ファイルが無い場合は、部屋状態CSVの選択ダイアログを表示する。"""
        # 1) 部屋状態CSVを決定（指定が無ければダイアログ。macOSのグレーアウト回避でフィルタ無し）
        if not csv_path or not os.path.exists(csv_path):
            csv_path = filedialog.askopenfilename(
                title="部屋状態CSVを選択"
            )
            if not csv_path:
                return

        # 部屋状態CSVかどうかを判定
        if self.detect_csv_type(csv_path) != 'room_status':
            messagebox.showerror(
                "エラー",
                "選択されたファイルは部屋状態CSVではないようです。\n"
                "状態列を含む部屋状態CSVを選択してください。"
            )
            return

        # 2) アーニング表テンプレートを取得
        #    まず EcoRoomClean.py と同じフォルダから自動検出する。
        #    見つからない場合のみ、フィルタなしの選択ダイアログを出す。
        template_path = self._find_earning_template()
        if not template_path:
            messagebox.showinfo(
                "テンプレート選択",
                f"同じフォルダにアーニング表テンプレート"
                f"（{self.EARNING_TEMPLATE_NAME}）が見つかりませんでした。\n"
                "テンプレートのExcelファイルを手動で選択してください。"
            )
            template_path = filedialog.askopenfilename(
                title="アーニング表テンプレートを選択"
            )
            if not template_path:
                return

        try:
            # CSV読み込み: 部屋番号(int) -> 状態コード、ファイル日付
            room_status = {}
            file_date = None
            with open(csv_path, 'r', encoding='utf-8') as f:
                for row in csv.reader(f):
                    if len(row) < 7:
                        continue
                    # 1列目の8桁日付をファイル名用に取得
                    if file_date is None and row[0].isdigit() and len(row[0]) == 8:
                        file_date = row[0]
                    raw_room = row[1].strip()
                    if not raw_room.isdigit():
                        continue
                    room_status[int(raw_room)] = row[6].strip()

            if not room_status:
                messagebox.showerror("エラー", "CSVから部屋データを読み込めませんでした。")
                return

            # DBに登録されている部屋
            # （状態3はこれに含まれればエコドア／エコ清掃／清掃不要のいずれか）
            db_registered, db_ecodoor, db_noclean = \
                self._load_db_registered_rooms()

            # テンプレート読み込み & 指示書用セルマップ作成
            # .xlsm（マクロ有効ブック）の場合は keep_vba=True でマクロを保持する
            template_ext = os.path.splitext(template_path)[1].lower()
            is_macro = (template_ext == '.xlsm')
            try:
                wb = openpyxl.load_workbook(template_path, keep_vba=is_macro)
            except Exception as load_err:
                messagebox.showerror(
                    "エラー",
                    "アーニング表テンプレートを開けませんでした。\n\n"
                    f"ファイル: {os.path.basename(template_path)}\n"
                    f"原因: {load_err}\n\n"
                    "次の点を確認してください。\n"
                    "・拡張子が .xlsx または .xlsm のExcelファイルか\n"
                    "・そのファイルをExcelで開いたままにしていないか\n"
                    "・.xls（旧形式）や .xlsb（バイナリ形式）ではないか"
                )
                return
            cell_map = self._build_instruction_cell_map(wb)

            if not cell_map:
                messagebox.showerror(
                    "エラー",
                    "テンプレートから【指示書用】シートの部屋番号を認識できませんでした。\n"
                    "シート名に「指示書」を含むシートがあるか確認してください。"
                )
                wb.close()
                return

            # 区分ごとのカウンタ
            counts = {'vacant': 0, 'pre_ci': 0, 'checkout': 0,
                      'ecodoor': 0, 'eco': 0, 'noclean': 0, 'stay': 0}
            room_class = {}  # 部屋番号(int) -> 区分キー（階別サマリー用）
            unmatched = []   # CSVにあるがテンプレートに無い部屋
            written = 0

            for room_int, status in room_status.items():
                if room_int not in cell_map:
                    unmatched.append(room_int)
                    continue

                # 状態コード -> 区分判定
                if status == '0':
                    key = 'vacant'
                elif status == '1':
                    key = 'pre_ci'
                elif status == '2':
                    key = 'checkout'
                elif status == '3':
                    # 状態3：DB登録あり → 清掃不要 / エコドア / エコ清掃、
                    #        登録なし → ○（連泊）
                    if room_int in db_registered:
                        if room_int in db_noclean:
                            key = 'noclean'
                        elif room_int in db_ecodoor:
                            key = 'ecodoor'
                        else:
                            key = 'eco'
                    else:
                        key = 'stay'
                else:
                    # 想定外の状態コードはスキップ
                    continue

                ws, r, c = cell_map[room_int]
                mark_cell = ws.cell(r, c + 2, self.EARNING_MARKS[key])  # 番号セルの2つ隣に書き込み

                # エコ清掃・エコドア・清掃不要・連泊は部屋番号セルを色付け
                if key in ('eco', 'ecodoor', 'noclean', 'stay'):
                    ws.cell(r, c).fill = self.EARNING_GREEN_FILL
                # エコ清掃・エコドア・清掃不要は記号欄も色付け（エコ清掃扱い）
                if key in ('eco', 'ecodoor', 'noclean'):
                    mark_cell.fill = self.EARNING_GREEN_FILL

                counts[key] += 1
                room_class[room_int] = key
                written += 1

            # 各階見出しセルに部屋情報サマリー（S/T/D・エコ・空室）を追記
            self._write_floor_summaries(wb, room_class)

            # 出力ファイル名（日付入り）
            # 出力ファイル名（日付入り）。テンプレートが.xlsmならマクロ保持のため.xlsmで保存
            date_part = file_date if file_date else datetime.now().strftime('%Y%m%d')
            out_ext = '.xlsm' if is_macro else '.xlsx'
            output_path = f"アーニング表_出力_{date_part}{out_ext}"
            wb.save(output_path)
            wb.close()

            # 結果メッセージ
            msg = "アーニング表を出力しました。\n\n"
            msg += f"出力ファイル: {output_path}\n"
            msg += f"書き込み: {written}室\n\n"
            msg += f"  × 空室: {counts['vacant']}室\n"
            msg += f"  未C/I 未チェックイン: {counts['pre_ci']}室\n"
            msg += f"  （空欄）チェックアウト: {counts['checkout']}室\n"
            msg += f"  エコドア: {counts['ecodoor']}室\n"
            msg += f"  エコ清掃: {counts['eco']}室\n"
            msg += f"  清掃不要: {counts['noclean']}室\n"
            msg += f"  （空欄・緑色）連泊: {counts['stay']}室"
            msg += "\n\n各階見出しに部屋情報（在室S/T・eco丸数字・空室）を書き込みました。"
            if unmatched:
                msg += f"\n\n⚠ テンプレートに無い部屋{len(unmatched)}室はスキップしました:\n  {sorted(unmatched)}"

            messagebox.showinfo("出力完了", msg)

            # 出力ファイルをそのまま開く
            self.open_file(output_path)

        except Exception as e:
            messagebox.showerror("エラー", f"アーニング表の出力に失敗しました: {e}")

    def show_eco_room_list_dialog(self):
        """日付を指定して、その日エコ対応（×／エコドア／清掃不要・アメ不要）の
        部屋をExcelに出力するダイアログを表示する"""
        dialog = tk.Toplevel(self.root)
        dialog.title("エコ部屋リスト出力")
        dialog.resizable(False, False)
        dialog.transient(self.root)
        dialog.grab_set()

        frame = ttk.Frame(dialog, padding="15")
        frame.pack(fill="both", expand=True)

        ttk.Label(frame, text="リストを出力する日付",
                  font=("", 13, "bold")).pack(pady=(0, 10))

        today = datetime.now()
        date_vars = {
            'year': tk.StringVar(value=str(today.year)),
            'month': tk.StringVar(value=str(today.month)),
            'day': tk.StringVar(value=str(today.day)),
        }

        date_frame = ttk.Frame(frame)
        date_frame.pack()
        ttk.Entry(date_frame, textvariable=date_vars['year'], width=6).pack(side="left")
        ttk.Label(date_frame, text="年").pack(side="left", padx=(0, 5))
        ttk.Entry(date_frame, textvariable=date_vars['month'], width=4).pack(side="left")
        ttk.Label(date_frame, text="月").pack(side="left", padx=(0, 5))
        ttk.Entry(date_frame, textvariable=date_vars['day'], width=4).pack(side="left")
        ttk.Label(date_frame, text="日").pack(side="left")

        ttk.Label(frame,
                  text="その日の欄が ×／エコドア／清掃不要・アメ不要 の\n"
                       "部屋を一覧にします（〇の日の部屋は含みません）",
                  font=("", 10), justify=tk.CENTER).pack(pady=10)

        def do_export():
            try:
                target_date = datetime(
                    int(date_vars['year'].get()),
                    int(date_vars['month'].get()),
                    int(date_vars['day'].get()))
            except ValueError:
                messagebox.showerror("エラー", "正しい日付を入力してください。")
                return
            dialog.destroy()
            self.export_eco_room_list(target_date)

        btn_frame = ttk.Frame(frame)
        btn_frame.pack(pady=(5, 0))
        ttk.Button(btn_frame, text="出力", command=do_export).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="キャンセル",
                   command=dialog.destroy).pack(side="left", padx=5)

    def export_eco_room_list(self, target_date):
        """target_date の欄が ×／エコドア／清掃不要・アメ不要 になっている
        部屋（＝その日エコ対応する部屋）を部屋番号順で
        「エコ部屋リスト_YYYYMMDD.xlsx」に出力して開く。

        〇（3日周期の清掃日）はエコ・通常共通の清掃作業なので含めない。
        戻り値は出力ファイルパス（該当なし・失敗時は None）。"""
        eco_statuses = self.ECO_MIDDLE_STATUSES
        cursor = self.conn.cursor()
        cursor.execute(
            """SELECT s.room_number, s.cleaning_status,
                      r.guest_name, r.check_in_date, r.cleaning_days
               FROM cleaning_schedule s
               LEFT JOIN rooms r ON r.room_number = s.room_number
               WHERE s.cleaning_date = ?
                 AND s.cleaning_status IN (?, ?, ?)""",
            (target_date.strftime('%Y-%m-%d'), *eco_statuses))
        rows = cursor.fetchall()

        if not rows:
            messagebox.showinfo(
                "情報",
                f"{target_date.strftime('%Y年%m月%d日')} に"
                "エコ対応の部屋はありません。")
            return None

        # 部屋番号順（数字でない部屋番号は末尾）
        def room_key(row):
            rn = row[0]
            return (0, int(rn)) if str(rn).isdigit() else (1, str(rn))
        rows.sort(key=room_key)

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "エコ部屋リスト"

        # タイトル・ヘッダー
        title_cell = ws.cell(1, 1,
                             f"エコ部屋リスト {target_date.strftime('%Y/%m/%d')}"
                             f"（{len(rows)}室）")
        title_cell.font = Font(bold=True, size=14)

        headers = ("部屋番号", "宿泊者名", "当日状態", "C/I", "C/O")
        for col, text in enumerate(headers, start=1):
            cell = ws.cell(3, col, text)
            cell.font = Font(bold=True)
            cell.fill = self.EARNING_GREEN_FILL

        # データ行
        for i, (room, status, guest, ci_str, days) in enumerate(rows):
            row_no = 4 + i
            ws.cell(row_no, 1, room)
            ws.cell(row_no, 2, guest or '')
            ws.cell(row_no, 3, status)
            # C/I・C/O は rooms テーブルから復元（無ければ空欄）
            if ci_str and days is not None:
                try:
                    ci = datetime.strptime(ci_str, '%Y-%m-%d')
                    co = ci + timedelta(days=days)
                    ws.cell(row_no, 4, f"{ci.month}/{ci.day}")
                    ws.cell(row_no, 5, f"{co.month}/{co.day}")
                except ValueError:
                    pass

        # 列幅
        for col_letter, width in (('A', 10), ('B', 22), ('C', 16),
                                  ('D', 8), ('E', 8)):
            ws.column_dimensions[col_letter].width = width

        output_path = f"エコ部屋リスト_{target_date.strftime('%Y%m%d')}.xlsx"
        try:
            wb.save(output_path)
        except Exception as e:
            messagebox.showerror(
                "エラー",
                f"エコ部屋リストの保存に失敗しました: {e}\n\n"
                "同名ファイルをExcelで開いたままにしていないか確認してください。")
            return None
        finally:
            wb.close()

        messagebox.showinfo(
            "出力完了",
            f"{target_date.strftime('%Y年%m月%d日')} のエコ部屋 {len(rows)}室を\n"
            f"{output_path} に出力しました。")
        self.open_file(output_path)
        return output_path

    def open_file(self, path):
        """指定したファイルを既定のアプリで開く"""
        try:
            if platform.system() == "Darwin":
                subprocess.call(("open", path))
            elif platform.system() == "Windows":
                os.startfile(path)
            else:
                subprocess.call(("xdg-open", path))
        except Exception as e:
            messagebox.showerror("エラー", f"ファイルを開けませんでした: {e}")

    def open_excel(self):
        """エクセルファイルを開く"""
        try:
            if platform.system() == "Darwin":
                subprocess.call(("open", self.excel_file))
            elif platform.system() == "Windows":
                os.startfile(self.excel_file)
            else:
                subprocess.call(("xdg-open", self.excel_file))
        except Exception as e:
            messagebox.showerror("エラー", f"ファイルを開けませんでした: {e}")

    def run(self):
        """アプリケーションの実行"""
        self.root.protocol("WM_DELETE_WINDOW", lambda: (self.conn.close(), self.root.destroy()))
        self.root.mainloop()


if __name__ == "__main__":
    app = HotelCleaningSystem()
    app.run()