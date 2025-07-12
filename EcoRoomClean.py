import tkinter as tk
from tkinter import ttk, messagebox
import sqlite3
import openpyxl
from datetime import datetime, timedelta
import os
import subprocess
import platform
import re

# IME関連警告を抑制
if platform.system() == "Darwin":
    os.environ['TK_SILENCE_DEPRECATION'] = '1'


class HotelCleaningSystem:
    def __init__(self):
        self.db_file = "hotel_cleaning.db"
        self.excel_file = "hotel_cleaning_now.xlsx"
        self.records = []
        self.existing_rooms = set()

        # GUI設定
        self.root = tk.Tk()
        self.root.title("客室清掃管理システム")
        self.root.geometry("400x350")

        self.setup_gui()
        self.init_database()
        self.load_data()

    def setup_gui(self):
        frame = ttk.Frame(self.root, padding="10")
        frame.pack(fill="both", expand=True)

        # 入力フィールド
        fields = [
            ("部屋番号", "room"),
            ("お客様名", "guest"),
            ("チェックイン日", "date"),
            ("清掃期間（日）", "days")
        ]

        self.vars = {}
        for i, (label, key) in enumerate(fields):
            ttk.Label(frame, text=f"{label}：").grid(row=i, column=0, sticky="w", pady=5)

            if key == "date":
                date_frame = ttk.Frame(frame)
                date_frame.grid(row=i, column=1, sticky="w")
                today = datetime.now()
                self.vars['year'] = tk.StringVar(value=str(today.year))
                self.vars['month'] = tk.StringVar(value=str(today.month))
                self.vars['day'] = tk.StringVar(value=str(today.day))

                for var, label_text in [('year', '年'), ('month', '月'), ('day', '日')]:
                    ttk.Entry(date_frame, textvariable=self.vars[var], width=5).pack(side="left")
                    ttk.Label(date_frame, text=label_text).pack(side="left")
            elif key == "days":
                self.vars[key] = tk.IntVar(value=2)
                ttk.Spinbox(frame, from_=2, to=999, textvariable=self.vars[key], width=20).grid(row=i, column=1,
                                                                                                sticky="w")
            else:
                self.vars[key] = tk.StringVar()
                entry = ttk.Entry(frame, textvariable=self.vars[key], width=20)
                entry.grid(row=i, column=1, sticky="w")
                if key == "room":
                    entry.bind('<KeyRelease>', self.filter_numbers)

        # チェックボックス
        ttk.Label(frame, text="オプション：").grid(row=4, column=0, sticky="w", pady=5)
        self.vars['ecodoor'] = tk.BooleanVar()
        ttk.Checkbutton(frame, text="エコドア", variable=self.vars['ecodoor']).grid(row=4, column=1, sticky="w")

        ttk.Label(frame, text="プラン：").grid(row=5, column=0, sticky="w", pady=5)
        self.vars['ecoplan'] = tk.BooleanVar()
        ttk.Checkbutton(frame, text="エコプラン", variable=self.vars['ecoplan']).grid(row=5, column=1, sticky="w")

        # ボタン
        btn_frame = ttk.Frame(frame)
        btn_frame.grid(row=6, column=0, columnspan=2, pady=20)

        ttk.Button(btn_frame, text="次の部屋", command=self.add_room).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="エコ票作成", command=self.create_schedule_with_cleanup).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="部屋編集", command=self.edit_room).pack(side="left", padx=5)

    def filter_numbers(self, event):
        text = self.vars['room'].get()
        filtered = re.sub(r'[^0-9]', '', text)
        if text != filtered:
            self.vars['room'].set(filtered)

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
                              BOOLEAN
                          )''')

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
        self.cleanup_expired()
        cursor = self.conn.cursor()
        cursor.execute("SELECT * FROM rooms ORDER BY CAST(room_number AS INTEGER)")

        for row in cursor.fetchall():
            room, guest, date_str, days, ecodoor, ecoplan = row
            self.existing_rooms.add(room)

            record = {
                'room': room,
                'guest': guest,
                'date': datetime.strptime(date_str, '%Y-%m-%d'),
                'days': days,
                'ecodoor': bool(ecodoor),
                'ecoplan': bool(ecoplan),
                'schedule': {},
                'is_new': False  # 既存レコード
            }

            # スケジュール読み込み
            cursor.execute(
                "SELECT cleaning_date, cleaning_status FROM cleaning_schedule WHERE room_number = ? ORDER BY cleaning_date",
                (room,))
            for date_str, status in cursor.fetchall():
                date = datetime.strptime(date_str, '%Y-%m-%d')
                record['schedule'][f"{date.month}/{date.day}"] = status

            self.records.append(record)

    def cleanup_expired(self):
        # Excelファイル基準で期限切れデータを削除
        if not os.path.exists(self.excel_file):
            return

        try:
            wb = openpyxl.load_workbook(self.excel_file)
            today_str = str(datetime.now().day)
            rooms_to_delete = set()

            for sheet in wb.worksheets:
                # 今日の日付の列を探す
                today_col = None
                for col in range(4, 35):
                    if str(sheet.cell(3, col).value) == today_str:
                        today_col = col
                        break

                if today_col:
                    # 今日の列が空白の部屋を削除対象に
                    for row in range(4, sheet.max_row + 1):
                        room = sheet.cell(row, 3).value
                        if room and not sheet.cell(row, today_col).value:
                            rooms_to_delete.add(str(room))

            wb.close()

            if rooms_to_delete:
                cursor = self.conn.cursor()
                cursor.execute(f"DELETE FROM rooms WHERE room_number IN ({','.join(['?'] * len(rooms_to_delete))})",
                               list(rooms_to_delete))
                self.conn.commit()
                print(f"{len(rooms_to_delete)}件の部屋を削除しました")
        except Exception as e:
            print(f"削除処理エラー: {e}")

    def add_room(self):
        room = self.vars['room'].get().strip()
        if not room:
            messagebox.showerror("エラー", "部屋番号を入力してください")
            return

        if room in self.existing_rooms or any(r['room'] == room for r in self.records):
            messagebox.showerror("エラー", "この部屋番号は既に登録されています")
            return

        try:
            date = datetime(int(self.vars['year'].get()), int(self.vars['month'].get()), int(self.vars['day'].get()))
        except ValueError:
            messagebox.showerror("エラー", "正しい日付を入力してください")
            return

        record = {
            'room': room,
            'guest': self.vars['guest'].get().strip(),
            'date': date,
            'days': self.vars['days'].get(),
            'ecodoor': self.vars['ecodoor'].get(),
            'ecoplan': self.vars['ecoplan'].get(),
            'schedule': {}
        }

        # スケジュール生成
        current = date
        for day in range(record['days'] + 1):
            date_str = f"{current.month}/{current.day}"
            if day == 0:
                status = "C/I"
            elif day == record['days']:
                status = "C/O"
            else:
                status = "〇" if day % 3 == 0 else ("エコドア" if record['ecodoor'] else "×")

            record['schedule'][date_str] = status
            current += timedelta(days=1)

        self.records.append(record)
        self.existing_rooms.add(room)

        # 新しいレコードとしてマーク
        record['is_new'] = True

        # フォームクリア
        for key in ['room', 'guest']:
            self.vars[key].set("")
        for key in ['ecodoor', 'ecoplan']:
            self.vars[key].set(False)

        messagebox.showinfo("情報", "部屋が追加されました")

    def create_schedule_with_cleanup(self):
        """エコ票作成と同時に清掃日決定・データベース整理・バックアップを実行"""
        if not self.records:
            messagebox.showerror("エラー", "記録する部屋がありません")
            return

        # 入力中の部屋があれば追加
        if self.vars['room'].get().strip():
            self.add_room()

        # 清掃日選択ダイアログを表示
        cleanup_date = self.show_cleanup_date_dialog()
        if not cleanup_date:
            return  # キャンセルされた場合

        # エコ票作成前の確認（清掃日とバックアップを含む）
        result = messagebox.askyesno(
            "エコ票作成・データベース整理確認",
            f"以下の処理を実行しますか？\n\n"
            f"1. エコ票の作成\n"
            f"2. 清掃日({cleanup_date.strftime('%Y年%m月%d日')})基準でのデータベース整理\n"
            f"3. データベースのバックアップ作成\n\n"
            f"※データベース整理により、{cleanup_date.strftime('%m月%d日')}の列が空白の部屋が削除されます",
            icon='question'
        )

        if not result:
            return

        try:
            # 1. バックアップ作成
            backup_file = self.create_database_backup()
            if not backup_file:
                # バックアップ失敗時は処理を中止
                return

            # 2. 新しいレコードをデータベース保存
            new_records = [r for r in self.records if r.get('is_new', True)]

            cursor = self.conn.cursor()
            for record in new_records:
                cursor.execute('''INSERT OR REPLACE INTO rooms VALUES (?, ?, ?, ?, ?, ?)''',
                               (record['room'], record['guest'], record['date'].strftime('%Y-%m-%d'),
                                record['days'], record['ecodoor'], record['ecoplan']))

                cursor.execute("DELETE FROM cleaning_schedule WHERE room_number = ?", (record['room'],))
                for date_str, status in record['schedule'].items():
                    month, day = map(int, date_str.split('/'))
                    # 年をまたぐ場合の処理
                    year = record['date'].year
                    if month < record['date'].month:
                        year += 1
                    cleaning_date = datetime(year, month, day)
                    cursor.execute("INSERT INTO cleaning_schedule VALUES (?, ?, ?)",
                                   (record['room'], cleaning_date.strftime('%Y-%m-%d'), status))

            self.conn.commit()

            # 3. Excel生成（全レコード使用）
            self.generate_excel()

            # 4. 清掃日基準でデータベース整理
            deleted_count = self.cleanup_based_on_exact_date(cleanup_date)

            # 5. データ再読み込み
            self.records.clear()
            self.existing_rooms.clear()
            self.load_data()

            # 6. 新しいレコードのフラグをクリア（残っているレコード用）
            for record in self.records:
                record['is_new'] = False

            # 完了メッセージ
            messagebox.showinfo("処理完了",
                                f"エコ票作成とデータベース整理が完了しました。\n\n"
                                f"✓ 新規登録: {len(new_records)}件\n"
                                f"✓ 削除された部屋: {deleted_count}件\n"
                                f"✓ 清掃基準日: {cleanup_date.strftime('%Y年%m月%d日')}\n"
                                f"✓ バックアップ: {backup_file}")

            # Excel ファイルを開く
            self.open_excel()

        except Exception as e:
            messagebox.showerror("エラー", f"処理中にエラーが発生しました: {e}")

    def show_cleanup_date_dialog(self):
        """清掃日選択ダイアログを表示"""
        dialog = tk.Toplevel(self.root)
        dialog.title("清掃日選択")
        dialog.geometry("400x300")
        dialog.resizable(False, False)
        dialog.transient(self.root)
        dialog.grab_set()

        frame = ttk.Frame(dialog, padding="20")
        frame.pack(fill="both", expand=True)

        # タイトル
        title_label = ttk.Label(frame, text="清掃日選択", font=("", 14, "bold"))
        title_label.pack(pady=(0, 20))

        # 説明
        info_label = ttk.Label(frame,
                               text="データベース整理の基準となる清掃日を選択してください。\n"
                                    "この日付の列が空白の部屋がデータベースから削除されます。",
                               font=("", 10), justify=tk.CENTER)
        info_label.pack(pady=(0, 20))

        # 日付選択
        date_frame = ttk.LabelFrame(frame, text="清掃日", padding="15")
        date_frame.pack(fill="x", pady=(0, 20))

        # デフォルトは今日の日付
        today = datetime.now()
        cleanup_vars = {
            'year': tk.StringVar(value=str(today.year)),
            'month': tk.StringVar(value=str(today.month)),
            'day': tk.StringVar(value=str(today.day))
        }

        date_input_frame = ttk.Frame(date_frame)
        date_input_frame.pack()

        ttk.Entry(date_input_frame, textvariable=cleanup_vars['year'], width=6).pack(side="left")
        ttk.Label(date_input_frame, text="年").pack(side="left", padx=(0, 10))

        ttk.Entry(date_input_frame, textvariable=cleanup_vars['month'], width=4).pack(side="left")
        ttk.Label(date_input_frame, text="月").pack(side="left", padx=(0, 10))

        ttk.Entry(date_input_frame, textvariable=cleanup_vars['day'], width=4).pack(side="left")
        ttk.Label(date_input_frame, text="日").pack(side="left")

        # 今日の日付ボタン
        today_frame = ttk.Frame(date_frame)
        today_frame.pack(pady=(10, 0))

        def set_today():
            today = datetime.now()
            cleanup_vars['year'].set(str(today.year))
            cleanup_vars['month'].set(str(today.month))
            cleanup_vars['day'].set(str(today.day))

        ttk.Button(today_frame, text="今日の日付", command=set_today).pack()

        # 結果を格納する変数
        result = {'date': None}

        # ボタン
        button_frame = ttk.Frame(frame)
        button_frame.pack(fill="x")

        def on_ok():
            try:
                cleanup_date = datetime(
                    int(cleanup_vars['year'].get()),
                    int(cleanup_vars['month'].get()),
                    int(cleanup_vars['day'].get())
                )
                result['date'] = cleanup_date
                dialog.destroy()
            except ValueError:
                messagebox.showerror("エラー", "正しい日付を入力してください")

        def on_cancel():
            dialog.destroy()

        ttk.Button(button_frame, text="OK", command=on_ok).pack(side="left", padx=5)
        ttk.Button(button_frame, text="キャンセル", command=on_cancel).pack(side="left", padx=5)

        # ダイアログが閉じられるまで待機
        dialog.wait_window()

        return result['date']

    def create_schedule(self):
        """従来のエコ票作成機能（バックアップ機能付きに変更）"""
        # 新しい統合機能を呼び出す
        self.create_schedule_with_cleanup()

    def generate_excel(self):
        wb = openpyxl.Workbook()
        wb.remove(wb.active)  # デフォルトシートを削除

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

            # 日付ヘッダー（その月の日数分のみ）
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
            ws.column_dimensions['A'].width = 12  # 氏名
            ws.column_dimensions['C'].width = 8  # 部屋番号
            for day in range(1, days_in_month + 1):
                col_letter = openpyxl.utils.get_column_letter(3 + day)
                ws.column_dimensions[col_letter].width = 6

        wb.save(self.excel_file)
        wb.close()

    def delete_room(self):
        room = self.vars['room'].get().strip()

        if not room:
            # 選択ダイアログ
            all_rooms = list(self.existing_rooms) + [r['room'] for r in self.records]
            all_rooms = sorted(set(all_rooms), key=lambda x: int(x) if x.isdigit() else float('inf'))

            if not all_rooms:
                messagebox.showinfo("情報", "削除できる部屋がありません")
                return

            dialog = tk.Toplevel(self.root)
            dialog.title("部屋削除")
            dialog.geometry("200x300")

            listbox = tk.Listbox(dialog)
            for r in all_rooms:
                listbox.insert(tk.END, r)
            listbox.pack(fill="both", expand=True, padx=10, pady=10)

            def delete_selected():
                sel = listbox.curselection()
                if sel:
                    room = all_rooms[sel[0]]
                    dialog.destroy()
                    self.confirm_delete(room)

            ttk.Button(dialog, text="削除", command=delete_selected).pack(pady=5)
        else:
            self.confirm_delete(room)

    def confirm_delete(self, room):
        if messagebox.askyesno("削除確認", f"部屋番号 {room} を削除しますか？"):
            cursor = self.conn.cursor()
            cursor.execute("DELETE FROM rooms WHERE room_number = ?", (room,))
            self.conn.commit()

            self.records = [r for r in self.records if r['room'] != room]
            self.existing_rooms.discard(room)

            messagebox.showinfo("削除完了", f"部屋番号 {room} を削除しました")

    def edit_room(self):
        """部屋の編集"""
        room = self.vars['room'].get().strip()

        if not room:
            self.show_room_edit_dialog()
            return

        # 入力された部屋番号の編集
        self.open_edit_dialog(room)

    def show_room_edit_dialog(self):
        """部屋選択ダイアログを表示（編集用）"""
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
            # 部屋の詳細情報を表示
            record = self.find_room_record(room)
            if record:
                display_text = f"{room} - {record['guest']} ({record['days']}日)"
            else:
                display_text = f"{room} - (詳細不明)"
            listbox.insert(tk.END, display_text)
        listbox.pack(fill="both", expand=True, padx=10, pady=5)

        button_frame = ttk.Frame(dialog)
        button_frame.pack(pady=10)

        def on_edit():
            selection = listbox.curselection()
            if selection:
                selected_room = all_rooms[selection[0]]
                dialog.destroy()
                self.open_edit_dialog(selected_room)
            else:
                messagebox.showwarning("警告", "部屋を選択してください。")

        ttk.Button(button_frame, text="編集", command=on_edit).pack(side="left", padx=5)
        ttk.Button(button_frame, text="キャンセル", command=dialog.destroy).pack(side="left", padx=5)

    def find_room_record(self, room_number):
        """部屋番号からレコードを検索"""
        for record in self.records:
            if record['room'] == room_number:
                return record

        # データベースから検索
        cursor = self.conn.cursor()
        cursor.execute("SELECT * FROM rooms WHERE room_number = ?", (room_number,))
        row = cursor.fetchone()
        if row:
            room, guest, date_str, days, ecodoor, ecoplan = row
            return {
                'room': room,
                'guest': guest,
                'date': datetime.strptime(date_str, '%Y-%m-%d'),
                'days': days,
                'ecodoor': bool(ecodoor),
                'ecoplan': bool(ecoplan)
            }
        return None

    def open_edit_dialog(self, room_number):
        """編集ダイアログを開く"""
        record = self.find_room_record(room_number)
        if not record:
            messagebox.showerror("エラー", f"部屋番号 {room_number} の情報が見つかりません")
            return

        dialog = tk.Toplevel(self.root)
        dialog.title(f"部屋 {room_number} の編集")
        dialog.geometry("500x650")
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

        # チェックアウト日を計算
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

        # エコドア
        ttk.Label(frame, text="オプション：").grid(row=5, column=0, sticky="w", pady=5)
        edit_vars['ecodoor'] = tk.BooleanVar(value=record['ecodoor'])
        ttk.Checkbutton(frame, text="エコドア", variable=edit_vars['ecodoor']).grid(row=5, column=1, sticky="w", pady=5)

        # エコプラン
        ttk.Label(frame, text="プラン：").grid(row=6, column=0, sticky="w", pady=5)
        edit_vars['ecoplan'] = tk.BooleanVar(value=record['ecoplan'])
        ttk.Checkbutton(frame, text="エコプラン", variable=edit_vars['ecoplan']).grid(row=6, column=1, sticky="w",
                                                                                      pady=5)

        # 個別スケジュール編集
        ttk.Label(frame, text="スケジュール編集：").grid(row=7, column=0, sticky="nw", pady=5)
        schedule_frame = ttk.Frame(frame)
        schedule_frame.grid(row=7, column=1, sticky="ew", pady=5)

        # スケジュール編集用のScrollable Frame
        canvas = tk.Canvas(schedule_frame, width=400, height=250, bg="white")
        scrollbar = ttk.Scrollbar(schedule_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)

        # スクロール設定
        def configure_scroll_region(event):
            canvas.configure(scrollregion=canvas.bbox("all"))

        def on_mousewheel(event):
            # macOS/Windows/Linux対応のマウスホイール
            if platform.system() == "Darwin":  # macOS
                # macOSでは event.delta の値が異なる
                canvas.yview_scroll(int(-1 * event.delta), "units")
            else:  # Windows/Linux
                canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        # マウスホイールバインド（複数のイベントに対応）
        def bind_mousewheel(widget):
            if platform.system() == "Darwin":  # macOS
                widget.bind("<MouseWheel>", on_mousewheel)
            else:  # Windows/Linux
                widget.bind("<MouseWheel>", on_mousewheel)
                widget.bind("<Button-4>", lambda e: canvas.yview_scroll(-1, "units"))
                widget.bind("<Button-5>", lambda e: canvas.yview_scroll(1, "units"))

        # キャンバスとスクロール可能フレームにマウスホイールをバインド
        bind_mousewheel(canvas)
        bind_mousewheel(scrollable_frame)

        # フォーカスを設定してスクロールを有効にする
        def on_canvas_click(event):
            canvas.focus_set()

        canvas.bind("<Button-1>", on_canvas_click)
        canvas.bind("<Configure>", configure_scroll_region)
        scrollable_frame.bind("<Configure>", configure_scroll_region)

        canvas_window = canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        # キャンバスのサイズ調整
        def configure_canvas(event):
            canvas.itemconfig(canvas_window, width=event.width)

        canvas.bind('<Configure>', configure_canvas)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # スケジュール項目の辞書
        schedule_vars = {}

        # スケジュール表示を更新する関数
        def update_schedule_display():
            # 既存のウィジェットを削除
            for widget in scrollable_frame.winfo_children():
                widget.destroy()

            schedule_vars.clear()

            try:
                # 空文字列チェック
                if not all([edit_vars['checkin_year'].get(), edit_vars['checkin_month'].get(),
                            edit_vars['checkin_day'].get(), edit_vars['checkout_year'].get(),
                            edit_vars['checkout_month'].get(), edit_vars['checkout_day'].get()]):
                    # 通常のLabel（ttk.Labelではない）を使用
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

                    # 現在の状態を取得
                    if current_date == checkin:
                        current_status = "C/I"
                    elif current_date == checkout:
                        current_status = "C/O"
                    else:
                        current_status = current_schedule.get(date_str, "×")

                    # 現在の状態表示
                    status_label = ttk.Label(row_frame, text=current_status, width=8)
                    status_label.grid(row=0, column=1, padx=5, pady=2, sticky="w")

                    # 変更用コンボボックス
                    schedule_vars[date_str] = tk.StringVar(value=current_status)
                    status_combo = ttk.Combobox(row_frame, textvariable=schedule_vars[date_str], width=8)
                    status_combo['values'] = ('C/I', 'C/O', '〇', '×', 'エコドア')
                    status_combo.grid(row=0, column=2, padx=5, pady=2, sticky="w")

                    # C/IとC/Oは固定（変更不可）
                    if current_date == checkin or current_date == checkout:
                        status_combo.config(state='disabled')

                    # 各行フレームにもマウスホイールをバインド
                    bind_mousewheel(row_frame)
                    bind_mousewheel(date_label)
                    bind_mousewheel(status_label)

                    current_date += timedelta(days=1)
                    row += 1

                # 全てのウィジェットにマウスホイールをバインド
                bind_mousewheel(main_container)
                bind_mousewheel(header_frame)
                bind_mousewheel(data_frame)

                # スクロール範囲を更新（少し遅延させて確実に更新）
                def update_scroll_region():
                    scrollable_frame.update_idletasks()
                    canvas.update_idletasks()
                    canvas.configure(scrollregion=canvas.bbox("all"))
                    # 初期位置を一番上に設定
                    canvas.yview_moveto(0)

                # 少し遅延してスクロール範囲を更新
                canvas.after(100, update_scroll_region)

                # 長期間の場合はスクロールバーを目立たせる
                if days > 10:
                    info_label = tk.Label(main_container,
                                          text=f"※{days}日間の長期滞在 - マウスホイールまたはスクロールバーで全日程を確認してください",
                                          fg="blue", font=("", 8), wraplength=350)
                    info_label.pack(pady=5)
                    bind_mousewheel(info_label)

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

        # 初期表示
        update_schedule_display()

        # ボタン
        button_frame = ttk.Frame(frame)
        button_frame.grid(row=8, column=0, columnspan=2, pady=20)

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
            updated_record = {
                'room': room_number,
                'guest': edit_vars['guest'].get().strip(),
                'date': checkin_date,
                'days': calculated_days,
                'ecodoor': edit_vars['ecodoor'].get(),
                'ecoplan': edit_vars['ecoplan'].get(),
                'schedule': {},
                'is_new': True
            }

            # 個別に編集されたスケジュールを使用
            for date_str, var in schedule_vars.items():
                updated_record['schedule'][date_str] = var.get()

            # メモリ上のレコードを更新
            for i, rec in enumerate(self.records):
                if rec['room'] == room_number:
                    self.records[i] = updated_record
                    break
            else:
                self.records.append(updated_record)

            # データベースに保存
            cursor = self.conn.cursor()
            cursor.execute('''INSERT OR REPLACE INTO rooms VALUES (?, ?, ?, ?, ?, ?)''',
                           (updated_record['room'], updated_record['guest'],
                            updated_record['date'].strftime('%Y-%m-%d'),
                            updated_record['days'], updated_record['ecodoor'], updated_record['ecoplan']))

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
            messagebox.showinfo("成功", f"部屋 {room_number} の情報を更新しました")

        ttk.Button(button_frame, text="保存", command=save_changes).pack(side="left", padx=5)
        ttk.Button(button_frame, text="キャンセル", command=dialog.destroy).pack(side="left", padx=5)

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

    def clear_form(self):
        """フォームのクリア"""
        self.vars['room'].set("")
        self.vars['guest'].set("")
        self.vars['ecodoor'].set(False)
        self.vars['ecoplan'].set(False)

    def execute_database_cleanup(self):
        """メインフォームの清掃日を使用してデータベース整理を実行"""
        try:
            cleanup_date = datetime(
                int(self.vars['cleanup_year'].get()),
                int(self.vars['cleanup_month'].get()),
                int(self.vars['cleanup_day'].get())
            )
        except ValueError:
            messagebox.showerror("エラー", "正しい清掃日を入力してください")
            return

        # 確認ダイアログ
        result = messagebox.askyesno(
            "データベース整理確認",
            f"清掃日 {cleanup_date.strftime('%Y年%m月%d日')} を基準にデータベース整理を実行しますか？\n\n"
            "この操作により、該当日の列が空白の部屋がデータベースから削除されます。\n"
            "実行前にバックアップを作成することを推奨します。",
            icon='warning'
        )

        if not result:
            return

        # バックアップ作成の確認
        create_backup = messagebox.askyesno(
            "バックアップ作成",
            "実行前にデータベースのバックアップを作成しますか？\n"
            "（推奨：はい）"
        )

        backup_file = None
        if create_backup:
            backup_file = self.create_database_backup()
            if not backup_file:
                return

        try:
            # データベース整理実行
            deleted_count = self.cleanup_based_on_exact_date(cleanup_date)

            # データ再読み込み
            self.records.clear()
            self.existing_rooms.clear()
            self.load_data()

            if deleted_count > 0:
                messagebox.showinfo("整理完了",
                                    f"データベース整理が完了しました。\n"
                                    f"削除された部屋数: {deleted_count}件\n"
                                    f"削除基準日: {cleanup_date.strftime('%Y年%m月%d日')}\n"
                                    f"バックアップ: {backup_file if backup_file else '作成されませんでした'}")
            else:
                messagebox.showinfo("整理完了", "削除対象の部屋はありませんでした。")

        except Exception as e:
            messagebox.showerror("エラー", f"データベース整理中にエラーが発生しました: {e}")

    def cleanup_based_on_exact_date(self, cleanup_date):
        """指定された正確な日付を基準にしたデータベース整理"""
        if not os.path.exists(self.excel_file):
            print("Excelファイルが存在しないため、日付での削除をスキップします。")
            return 0

        try:
            wb = openpyxl.load_workbook(self.excel_file)
            cleanup_day_str = str(cleanup_date.day)
            cleanup_month = cleanup_date.month
            rooms_to_delete = set()

            print(f"削除基準: {cleanup_date.strftime('%Y年%m月%d日')}")

            for sheet in wb.worksheets:
                sheet_name = sheet.title
                print(f"シート '{sheet_name}' を処理中")

                # シート名から月を判定（例：「7月」「8月」）
                sheet_month = None
                if '月' in sheet_name:
                    try:
                        sheet_month = int(sheet_name.replace('月', ''))
                    except ValueError:
                        print(f"  - シート名から月を判定できません: {sheet_name}")
                        continue
                else:
                    print(f"  - 月の情報がないシートをスキップ: {sheet_name}")
                    continue

                # 清掃日の月と一致するシートのみ処理
                if sheet_month != cleanup_month:
                    print(f"  - 月が一致しないためスキップ（シート: {sheet_month}月, 清掃日: {cleanup_month}月）")
                    continue

                print(f"  - 月が一致するため処理実行（{cleanup_month}月）")

                # 指定された日付の列を探す
                cleanup_col = None
                for col in range(4, 35):
                    cell_value = sheet.cell(3, col).value
                    if cell_value and str(cell_value).strip() == cleanup_day_str:
                        cleanup_col = col
                        print(f"    - {cleanup_date.day}日の列を発見（列{cleanup_col}）")
                        break

                if cleanup_col:
                    # 指定日の列が空白の部屋を削除対象に
                    for row in range(4, sheet.max_row + 1):
                        room_cell = sheet.cell(row, 3)
                        if room_cell.value:
                            room_number = str(room_cell.value).strip()

                            cleanup_cell = sheet.cell(row, cleanup_col)
                            cleanup_value = str(cleanup_cell.value).strip() if cleanup_cell.value else ""

                            if not cleanup_value:
                                rooms_to_delete.add(room_number)
                                print(f"      削除対象: 部屋{room_number}")
                            else:
                                print(f"      保持: 部屋{room_number} ({cleanup_value})")
                else:
                    print(f"    - {cleanup_date.day}日の列が見つかりません")

            wb.close()

            # 削除対象の部屋をデータベースから削除
            if rooms_to_delete:
                self.delete_rooms_from_database(rooms_to_delete)
                print(
                    f"正確な日付({cleanup_date.strftime('%Y-%m-%d')})基準で {len(rooms_to_delete)} 件の部屋を削除しました。")
                return len(rooms_to_delete)
            else:
                print("正確な日付基準での削除対象はありませんでした。")
                return 0

        except Exception as e:
            print(f"正確な日付での削除中にエラーが発生しました: {e}")
            raise e

    def create_database_backup(self):
        """データベースのバックアップを作成"""
        try:
            import shutil
            from datetime import datetime

            backup_name = f"hotel_cleaning_backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.db"
            shutil.copy2(self.db_file, backup_name)

            messagebox.showinfo("バックアップ完了", f"バックアップを作成しました: {backup_name}")
            return backup_name
        except Exception as e:
            messagebox.showerror("バックアップエラー", f"バックアップの作成に失敗しました: {e}")
            return None

    def delete_rooms_from_database(self, room_numbers):
        """指定された部屋番号のリストをデータベースから削除"""
        if not room_numbers:
            return

        cursor = self.conn.cursor()
        placeholders = ','.join(['?' for _ in room_numbers])
        query = f"DELETE FROM rooms WHERE room_number IN ({placeholders})"

        cursor.execute(query, list(room_numbers))
        deleted_count = cursor.rowcount
        self.conn.commit()

        print(f"データベースから {deleted_count} 件の部屋を削除しました。")

    def open_excel(self):
        """エクセルファイルを開く"""
        try:
            if platform.system() == "Darwin":  # macOS
                subprocess.call(("open", self.excel_file))
            elif platform.system() == "Windows":  # Windows
                os.startfile(self.excel_file)
            else:  # Linux
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