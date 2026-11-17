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

# IME関連警告を抑制
if platform.system() == "Darwin":
    os.environ['TK_SILENCE_DEPRECATION'] = '1'


class HotelCleaningSystem:
    def __init__(self):
        self.db_file = "hotel_cleaning.db"
        # 直前のCSV読込で使った部屋状態CSVのパス（登録後のアーニング表出力の再利用用）
        self.last_room_status_csv = None

        # GUI設定
        self.root = tk.Tk()
        self.root.title("客室清掃管理システム")
        # ボタン3つ（CSV読込／アーニング表出力／エコ部屋リスト）
        # が横並びで見切れない幅を確保する
        self.root.geometry("560x200")

        # Combobox のドロップダウン一覧（tk.Listbox）の文字を大きくする
        self.root.option_add("*TCombobox*Listbox.font", ("", self.BIG_FONT_SIZE))

        self.init_database()

        self.setup_gui()
        self.update_room_count_display()
        self._show_main_window()

    def _show_main_window(self):
        """メイン画面を画面中央に出し、他のウィンドウより手前に表示する。

        起動時に開くダイアログが無くなったため、何もしないとメイン画面が
        画面左上の隅に、他のアプリの後ろに隠れた状態で開いてしまう
        （macOSでは起動しても何も出ていないように見える）。"""
        root = self.root
        root.update_idletasks()
        w, h = 560, 200
        x = max(0, (root.winfo_screenwidth() - w) // 2)
        y = max(0, (root.winfo_screenheight() - h) // 3)
        root.geometry(f"{w}x{h}+{x}+{y}")
        root.deiconify()
        root.lift()
        # 一時的に最前面に固定して手前に出し、すぐ解除する
        # （固定したままだと他のアプリの上に居座ってしまうため）
        root.attributes('-topmost', True)
        root.after(500, lambda: root.attributes('-topmost', False))
        root.focus_force()

    # プルダウン一覧の文字サイズ（pt）
    BIG_FONT_SIZE = 15

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

        ttk.Button(btn_frame, text="CSV読込", command=self.import_csv).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="アーニング表出力", command=self.output_earning_table).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="エコ部屋リスト", command=self.export_eco_room_list).pack(side="left", padx=5)

    def update_room_count_display(self):
        """登録内容（当日分）の件数表示を更新"""
        cursor = self.conn.cursor()
        cursor.execute("SELECT middle_status FROM rooms")
        statuses = [row[0] for row in cursor.fetchall()]
        eco = sum(1 for s in statuses if s in self.ECO_MIDDLE_STATUSES)
        duvet = sum(1 for s in statuses if s == self.DAILY_CLEAN_LABEL)
        target_date = self._get_registered_date()
        if statuses and target_date:
            self.room_count_var.set(
                f"{target_date.month}/{target_date.day} 分　"
                f"エコ: {eco}件　デュベ: {duvet}件　在室: {len(statuses)}件")
            self.status_var.set("アーニング表・エコ部屋リストを出力できます")
        else:
            self.room_count_var.set("登録なし")
            self.status_var.set("CSVを読み込んでください")

    def _get_registered_date(self):
        """いまDBに入っているデータの対象日（清掃日）を返す。無ければ None。

        DBは当日分しか持たないので、対象日は app_state に1件だけ保存している。"""
        cursor = self.conn.cursor()
        cursor.execute("SELECT value FROM app_state WHERE key = 'target_date'")
        row = cursor.fetchone()
        if not row or not row[0]:
            return None
        try:
            return datetime.strptime(row[0], '%Y-%m-%d')
        except ValueError:
            return None

    def clear_daily_data(self):
        """登録内容（rooms / cleaning_schedule / 対象日）を全消去する。

        エコ（×）・デュベ（〇）は翌日に引き継がない方針のため、
        CSVを読み込むたびに呼び出して白紙に戻す（commitは呼び出し側）。"""
        cursor = self.conn.cursor()
        cursor.execute("DELETE FROM rooms")
        cursor.execute("DELETE FROM cleaning_schedule")
        cursor.execute("DELETE FROM app_state WHERE key = 'target_date'")

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

        # CSV共有: 清掃管理システム（Java）側が同じCSVを選び直さずに使えるよう、
        # 直前に読み込んだCSVの中身を種別ごとに保存する
        cursor.execute('''CREATE TABLE IF NOT EXISTS shared_csv
                          (
                              kind
                              TEXT
                              PRIMARY
                              KEY,
                              file_name
                              TEXT,
                              csv_date
                              TEXT,
                              imported_at
                              TEXT,
                              content
                              TEXT
                          )''')

        # 当日分データの対象日（清掃日）などを持つ小さな設定テーブル
        cursor.execute('''CREATE TABLE IF NOT EXISTS app_state
                          (
                              key
                              TEXT
                              PRIMARY
                              KEY,
                              value
                              TEXT
                          )''')

        self.conn.commit()

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

    def read_csv_text(self, file_path):
        """CSVをテキストとして読む（utf-8 → cp932 → shift_jis の順に試す）"""
        for encoding in ('utf-8', 'cp932', 'shift_jis'):
            try:
                with open(file_path, 'r', encoding=encoding, newline='') as f:
                    return f.read()
            except UnicodeDecodeError:
                continue
        return None

    def save_shared_csvs(self, paths, csv_date):
        """読み込んだCSVの中身を shared_csv テーブルへ保存する。

        前回分は消してから保存し、今回選ばなかった種別（予約CSVなど）が古いまま残らないようにする。
        保存に失敗してもエコ側のCSV読込は続行する。
        """
        imported_at = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        date_text = csv_date.strftime('%Y-%m-%d') if csv_date else None
        try:
            cursor = self.conn.cursor()
            cursor.execute("DELETE FROM shared_csv")
            for kind, path in paths.items():
                if not path:
                    continue
                content = self.read_csv_text(path)
                if content is None:
                    continue
                # csv_date は部屋状態CSVの日付（清掃日）。他の種別は日付を持たない
                cursor.execute(
                    "INSERT INTO shared_csv (kind, file_name, csv_date, imported_at, content) "
                    "VALUES (?, ?, ?, ?, ?)",
                    (kind, os.path.basename(path),
                     date_text if kind == 'room_status' else None,
                     imported_at, content))
            self.conn.commit()
        except Exception as e:
            print(f"共有CSVの保存に失敗しました: {e}")

    def import_csv(self):
        """CSVファイルを読み込んで、その日のエコ・デュベ部屋を選択するダイアログを表示。
        部屋状態CSV・予約CSV・予約詳細CSV(C/O日入り) はどの順番で選択しても
        自動で振り分ける。

        エコ（×）・デュベ（〇）は翌日に引き継がないため、読み込んだ時点で
        DBの登録内容を全消去する。
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

        # 登録後のアーニング表出力に再利用するため、部屋状態CSVのパスを保持
        self.last_room_status_csv = room_status_path

        try:
            # 部屋状態CSVを読み込み
            eco_rooms = []
            csv_date = None  # CSV 1行目1列目から取得する日付（＝清掃日の前日）
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

                        # 状態が'3'（在室・連泊中＝エコ／デュベの選択対象）の部屋のみ抽出
                        # ※ '1'=未チェックイン、'0'=空室、'2'=チェックアウト は対象外
                        if room_status == '3':
                            eco_rooms.append({
                                'room': room_number,
                                'status': room_status
                            })

            # 前回の登録内容を全消去する（エコ・デュベは翌日に引き継がない）。
            # ただし同じ日付のCSVを読み直した場合（選び間違いの修正）に限り、
            # 直前の選択を選択画面の初期表示に使えるよう、消す前に控えておく
            prev_selection = {}
            if csv_date and (self._get_registered_date()
                             == csv_date + timedelta(days=1)):
                cursor = self.conn.cursor()
                cursor.execute("SELECT room_number, middle_status FROM rooms")
                prev_selection = dict(cursor.fetchall())
            self.clear_daily_data()
            self.conn.commit()
            self.update_room_count_display()

            # 清掃管理システム側へCSVを共有（在室部屋が無い日でも部屋状態は使えるため先に保存）
            self.save_shared_csvs({
                'room_status': room_status_path,
                'yoyaku': yoyaku_path,
                'yoyaku_co': yoyaku_co_path,
            }, csv_date)

            if not eco_rooms:
                messagebox.showinfo("情報", "在室（連泊中）の部屋が見つかりませんでした。")
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

            # 部屋選択ダイアログを表示
            self.show_csv_room_selection_dialog(eco_rooms, guest_name_map,
                                                csv_date, prev_selection)

        except Exception as e:
            messagebox.showerror("エラー", f"CSVファイルの読み込みに失敗しました: {e}")

    # エコプラン判定キーワード（NFKCで正規化してから部分一致で判定）
    ECO_PLAN_KEYWORDS = ['長期', 'ウィークリー', 'マンスリー', 'エコ']

    # rooms.middle_status に入れる区分（その日の扱い）。
    #   ×／エコドア／清掃不要・アメ不要 : エコとして選んだ部屋
    #   毎日デュベ清掃 : デュベ（〇）を選んだ部屋
    #   通常清掃       : どちらも選ばなかった在室部屋
    #                    （実C/I日から3日ごとの日だけ自動で〇）
    # 値の文字列は清掃管理システム側も読むため、以前のまま変えていない
    # （「毎日デュベ清掃」は今は「その日デュベにする部屋」の意味）。
    # 通常清掃・毎日デュベ清掃はアーニング表では「連泊(stay)」扱いにする
    # （_load_db_registered_rooms で登録済み集合から除外している）。
    NO_CLEAN_LABEL = '清掃不要・アメ不要'
    NORMAL_CLEAN_LABEL = '通常清掃'
    DAILY_CLEAN_LABEL = '毎日デュベ清掃'
    NON_ECO_MIDDLE_STATUSES = (NORMAL_CLEAN_LABEL, DAILY_CLEAN_LABEL)

    # エコ清掃扱いの区分（判定用）＝選択画面のエコ区分プルダウンの選択肢
    ECO_MIDDLE_STATUSES = ('×', 'エコドア', NO_CLEAN_LABEL)
    ECO_KIND_CHOICES = ECO_MIDDLE_STATUSES

    # 通常清掃の周期（日）と、清掃日に立てる記号。
    # 実C/I日を起点に NORMAL_CLEAN_CYCLE 日ごと（+3, +6, +9 …）が〇の日。
    # 〇でない日は cleaning_schedule に行を作らない。
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
        (チェックイン日datetime, チェックアウト日datetime, 泊数int) のマップを作成。

        列構成: 列0=予約番号, 列1=枝番(2列目), 列4=チェックイン日(5列目),
                列5=チェックアウト日(6列目), 列6=泊数(7列目), 列12=「取消」(13列目)。

        キーは2種類持つ:
          ・(予約番号, 枝番) … 基本のキー。同一予約番号でも枝番（部屋）ごとに
            C/O日が違うことがある（例: 団体予約の一部だけ延泊）ため。
          ・予約番号のみ    … 枝番が一致しないとき用のフォールバック。
            その予約番号の有効行の日付がすべて同じ場合にだけ登録する
            （食い違う場合はどれが正しいか判断できないので登録しない）。

        「取消」の行は読み飛ばす（有効な行の日付を上書きしないように）。

        C/I日は通常清掃の3日周期（実C/I日起点）の計算に使う。
        C/I日が読めなかった行は None を入れ、呼び出し側で登録起点日にフォールバックする。
        """
        encodings = ['cp932', 'shift_jis', 'utf-8']

        for encoding in encodings:
            try:
                # エンコーディングを変えて読み直すときに前回の途中結果を残さない
                co_map = {}
                values_by_no = {}  # 予約番号 -> その予約の有効行の日付の集合
                with open(file_path, 'r', encoding=encoding) as f:
                    for row in csv.reader(f):
                        if len(row) < 7 or not row[0].strip().isdigit():
                            continue
                        # 取消行は使わない
                        if len(row) > 12 and row[12].strip() == '取消':
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
                        yoyaku_no = row[0].strip()
                        value = (ci_date, co_date, nights)
                        co_map[(yoyaku_no, row[1].strip())] = value
                        values_by_no.setdefault(yoyaku_no, set()).add(value)
                # 日付が1通りに決まる予約番号だけ、予約番号のみのキーも登録する
                for yoyaku_no, values in values_by_no.items():
                    if len(values) == 1:
                        co_map[yoyaku_no] = next(iter(values))
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

        checkout_map（load_checkout_map の戻り値）が渡された場合は、
        (予約番号(列0), 枝番(列1)) をキーに各部屋へC/I日・C/O日・泊数を紐づける。
        枝番が一致しないときは予約番号のみで引く（日付が1通りの予約に限る）。
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
                            eda_no = row[1].strip()     # 2列目：枝番
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

                            # 予約番号＋枝番でC/I日・C/O日・泊数を引く。
                            # 無ければ予約番号のみで引き、それも無ければNone
                            ci_date, co_date, nights = checkout_map.get(
                                (yoyaku_no, eda_no),
                                checkout_map.get(yoyaku_no, (None, None, None)))

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
        """予約CSVから取れた情報と清掃日から (C/I日, 泊数) を決めて返す。

        ・C/I日は予約詳細CSVの実C/I日を優先し、取れなければ清掃日
          （この場合は3日周期の起点が分からないので自動の〇は付かない）
        ・C/O日が取れていればそこまでの実日数、取れなければ「清掃日+2日」まで
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

    def _is_cycle_clean_day(self, checkin_date, target_date):
        """target_date が、実C/I日から NORMAL_CLEAN_CYCLE 日ごと
        （+3, +6, +9 …）の自動〇の日かどうか"""
        offset = (target_date - checkin_date).days
        return offset > 0 and offset % self.NORMAL_CLEAN_CYCLE == 0

    def _build_day_schedule(self, checkin_date, days, middle_status, target_date):
        """対象日（清掃日）1日分のスケジュール {datetime: ステータス} を作る。

        エコ（×）・デュベ（〇）は対象日の分だけを人が選ぶ方式のため、
        先の日程は埋めない。対象日のセルに入る値は次のとおり。
          - ×／エコドア／清掃不要・アメ不要 : その値（エコとして選んだ部屋）
          - 毎日デュベ清掃（デュベを選んだ部屋）: 〇
          - 通常清掃（どちらも選ばなかった部屋）: 実C/I日から3日ごとの日だけ〇。
            それ以外の日は行を作らない
        あわせて C/I 日・C/O 日の行も入れる（清掃管理システム側が滞在期間を
        参照できるように）。清掃日がC/O日の部屋は×も〇も付けない。
        """
        schedule = {}
        checkout_date = checkin_date + timedelta(days=days)
        if checkin_date < target_date:
            schedule[checkin_date] = "C/I"
        schedule[checkout_date] = "C/O"
        if checkout_date <= target_date:
            # 清掃日にC/Oする部屋はエコ・デュベの対象外（C/Oの行だけ）
            return schedule

        if middle_status in self.ECO_MIDDLE_STATUSES:
            schedule[target_date] = middle_status
        elif middle_status == self.DAILY_CLEAN_LABEL:
            schedule[target_date] = self.NORMAL_CLEAN_MARK
        elif self._is_cycle_clean_day(checkin_date, target_date):
            schedule[target_date] = self.NORMAL_CLEAN_MARK
        return schedule

    def show_csv_room_selection_dialog(self, eco_rooms, guest_name_map=None,
                                       csv_date=None, prev_selection=None):
        """在室部屋の一覧を表示し、その日のエコ（×）とデュベ（〇）を
        部屋ごとに選んで登録するダイアログ。

        prev_selection は {部屋番号: middle_status}。同じ日付のCSVを
        読み直したとき（選び間違いの修正）に、直前の選択を初期表示するために使う。
        日付が変わった場合は渡されないので、毎日白紙から選ぶことになる。
        """
        if guest_name_map is None:
            guest_name_map = {}
        if prev_selection is None:
            prev_selection = {}

        dialog = tk.Toplevel(self.root)
        dialog.title("エコ・デュベ部屋選択")
        dialog.transient(self.root)
        dialog.grab_set()
        # 画面に収まるサイズに自動調整（小さい画面での見切れ防止）＋リサイズ可能に
        self._fit_toplevel_to_screen(dialog, 960, 1000, min_w=560, min_h=420)

        # このダイアログ内の文字を大きくする
        row_font = ("", 15)            # 部屋一覧の各行テキスト用
        header_font = ("", 14, "bold") # 一覧ヘッダー用
        dlg_style = ttk.Style(dialog)
        dlg_style.configure("Eco.TCheckbutton", font=row_font)
        dlg_style.configure("Eco.TButton", font=row_font)

        frame = ttk.Frame(dialog, padding="15")
        frame.pack(fill="both", expand=True)

        # タイトル
        title_label = ttk.Label(frame, text="デュベ（〇）・エコ（×）の選択", font=("", 20, "bold"))
        title_label.pack(pady=(0, 10))

        # 説明
        info_text = (f"在室の部屋が{len(eco_rooms)}件あります。"
                     "清掃日にエコ／デュベにする部屋へチェックを入れてください。\n"
                     f"どちらも選ばない部屋は、C/I日から{self.NORMAL_CLEAN_CYCLE}日ごとに"
                     "自動で〇が付きます。\n"
                     "エコプランの部屋は最初からチェック済みです"
                     f"（{self.NORMAL_CLEAN_CYCLE}日ごとの日はデュベ、それ以外はエコ）。")
        if prev_selection:
            info_text += "\n（同じ日付のCSVのため、前回の選択を表示しています）"
        info_label = ttk.Label(frame, text=info_text, font=("", 15), justify=tk.CENTER)
        info_label.pack(pady=(0, 10))

        # 清掃日（×・〇を付ける日）
        date_frame = ttk.LabelFrame(frame, text="清掃日（×・〇を付ける日＝CSV日付の翌日）", padding="10")
        date_frame.pack(fill="x", pady=(0, 10))

        # 清掃日はCSV 1行目1列目の日付の翌日（CSVは清掃日の前日時点の部屋状態）。
        # CSVから日付が取れなければ明日
        base = csv_date if csv_date else datetime.now().replace(
            hour=0, minute=0, second=0, microsecond=0)
        default_date = base + timedelta(days=1)
        date_vars = {
            'year': tk.StringVar(value=str(default_date.year)),
            'month': tk.StringVar(value=str(default_date.month)),
            'day': tk.StringVar(value=str(default_date.day))
        }

        date_input_frame = ttk.Frame(date_frame)
        date_input_frame.pack()

        ttk.Entry(date_input_frame, textvariable=date_vars['year'], width=6, font=row_font).pack(side="left")
        ttk.Label(date_input_frame, text="年", font=row_font).pack(side="left", padx=(0, 5))

        ttk.Entry(date_input_frame, textvariable=date_vars['month'], width=4, font=row_font).pack(side="left")
        ttk.Label(date_input_frame, text="月", font=row_font).pack(side="left", padx=(0, 5))

        ttk.Entry(date_input_frame, textvariable=date_vars['day'], width=4, font=row_font).pack(side="left")
        ttk.Label(date_input_frame, text="日", font=row_font).pack(side="left")

        def get_target_date():
            """画面上部の清掃日。入力不正なら None"""
            try:
                return datetime(int(date_vars['year'].get()),
                                int(date_vars['month'].get()),
                                int(date_vars['day'].get()))
            except ValueError:
                return None

        # 全解除ボタンと選択数
        select_frame = ttk.Frame(frame)
        select_frame.pack(fill="x", pady=(0, 10))

        eco_vars = {}     # 部屋番号 -> エコのチェック
        duvet_vars = {}   # 部屋番号 -> デュベのチェック
        kind_vars = {}    # 部屋番号 -> エコ区分（×／エコドア／清掃不要・アメ不要）
        auto_labels = {}  # 部屋番号 -> 自動〇の表示ラベル
        check_widgets = {}  # 部屋番号 -> (エコ, デュベ) のチェックボックス
        kind_combos = {}    # 部屋番号 -> エコ区分プルダウン

        def is_checkout_day(room_number, target_date):
            """清掃日がその部屋のC/O日（以降）か。C/Oする部屋は選択対象外"""
            checkout = guest_fields(room_number)[3]
            return bool(target_date and checkout and checkout <= target_date)

        def deselect_all():
            for var in eco_vars.values():
                var.set(False)
            for var in duvet_vars.values():
                var.set(False)

        ttk.Button(select_frame, text="全解除", command=deselect_all, style="Eco.TButton").pack(side="left", padx=3)

        # 選択数表示ラベル
        count_label = ttk.Label(select_frame, text="", font=row_font)
        count_label.pack(side="right", padx=5)

        def update_count(*args):
            eco = sum(1 for var in eco_vars.values() if var.get())
            duvet = sum(1 for var in duvet_vars.values() if var.get())
            count_label.config(text=f"デュベ: {duvet}件　エコ: {eco}件")

        # 登録／キャンセルボタンの枠を先に最下部へ固定しておく。
        # （list_frame より前に pack することで、ウィンドウが低くても
        #   ボタンが見切れず、代わりに一覧側がスクロールで縮む）
        button_frame = ttk.Frame(frame)
        button_frame.pack(side="bottom", fill="x", pady=(10, 0))

        # スクロール可能なフレーム
        list_frame = ttk.LabelFrame(frame, text="部屋一覧", padding="10")
        list_frame.pack(fill="both", expand=True, pady=(0, 15))

        canvas = tk.Canvas(list_frame, width=860, height=600, bg="white")
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)

        scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        # bind_all はアプリ全体にかかるため、このダイアログ以外の
        # ウィンドウ上のイベントには反応させない
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

        def close_dialog():
            cleanup_bindings()
            dialog.destroy()
            self.update_room_count_display()

        dialog.protocol("WM_DELETE_WINDOW", close_dialog)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # 列幅定数（pixel単位）— ヘッダーと各行で共通使用して列を揃える
        # 順: デュベ, エコ, 部屋番号, 宿泊者名, C/I, C/O, プラン, エコ区分, 自動
        COL_MINSIZE = (70, 60, 100, 170, 70, 70, 110, 200, 80)
        HEADERS = ("デュベ", "エコ", "部屋番号", "宿泊者名", "C/I", "C/O",
                   "プラン", "エコ区分", "自動")

        def _configure_row_columns(frame):
            """全ての行フレームで同じ列幅を共有させる"""
            for i, w in enumerate(COL_MINSIZE):
                frame.columnconfigure(i, minsize=w, weight=0)

        # ヘッダー
        header_frame = ttk.Frame(scrollable_frame)
        header_frame.pack(fill="x", pady=(0, 5))
        _configure_row_columns(header_frame)
        for col, text in enumerate(HEADERS):
            ttk.Label(header_frame, text=text, font=header_font, anchor="w").grid(
                row=0, column=col, padx=5, sticky="w")

        def guest_fields(room_number):
            """予約CSVの情報を (宿泊者名, エコプランか, C/I日, C/O日) で返す"""
            info = guest_name_map.get(room_number, {})
            # 後方互換：万一文字列が入っていてもクラッシュしないように
            if isinstance(info, dict):
                return (info.get('name', ''), info.get('is_ecoplan', False),
                        info.get('checkin'), info.get('checkout'))
            return (info or '', False, None, None)

        def update_auto_mark(room_number):
            """どちらも選んでいない部屋について、清掃日が3日周期の
            自動〇の日なら「〇」を表示する（選んでいる部屋は空欄）。
            清掃日にC/Oする部屋は「C/O」と表示してチェック不可にする"""
            label = auto_labels.get(room_number)
            if not label or not label.winfo_exists():
                return
            text = ""
            target_date = get_target_date()
            checkin = guest_fields(room_number)[2]
            checkout_day = is_checkout_day(room_number, target_date)
            for widget in check_widgets[room_number]:
                widget.config(state='disabled' if checkout_day else 'normal')
            if checkout_day:
                # 清掃日にC/Oする部屋：チェックを外して選べなくする
                # （値が変わるときだけ set する。trace 経由でここへ戻ってくるため）
                for var in (eco_vars[room_number], duvet_vars[room_number]):
                    if var.get():
                        var.set(False)
                text = "C/O"
            elif (target_date and checkin
                    and not eco_vars[room_number].get()
                    and not duvet_vars[room_number].get()
                    and self._is_cycle_clean_day(checkin, target_date)):
                text = self.NORMAL_CLEAN_MARK
            label.config(text=text)
            # エコ区分はエコ用の項目なので、デュベを選んだ部屋と
            # C/Oする部屋では選べなくする
            kind_combos[room_number].config(
                state='disabled'
                if checkout_day or duvet_vars[room_number].get()
                else 'readonly')

        def update_all_auto_marks(*_):
            for rn in list(auto_labels):
                update_auto_mark(rn)

        # 部屋ごとの行
        for idx, room_info in enumerate(eco_rooms):
            room_number = room_info['room']
            guest_name, is_ecoplan_guest, checkin_date, checkout_date = \
                guest_fields(room_number)

            # C/I・C/O列の表示（例: 6/1）。取れていない部屋は空欄
            checkin_text = (f"{checkin_date.month}/{checkin_date.day}"
                            if checkin_date else '')
            checkout_text = (f"{checkout_date.month}/{checkout_date.day}"
                             if checkout_date else '')

            # 初期状態：同じ日付の読み直しなら前回の選択。
            # それ以外はエコプラン該当者（長期／ウィークリー／マンスリー／エコ）
            # だけチェック済みにする。清掃日が実C/I日から3日ごとの日なら
            # デュベ、それ以外の日はエコ（どちらも手で変更できる）
            prev_status = prev_selection.get(room_number)
            if prev_selection:
                init_eco = prev_status in self.ECO_MIDDLE_STATUSES
                init_duvet = (prev_status == self.DAILY_CLEAN_LABEL)
            else:
                init_duvet = bool(
                    is_ecoplan_guest and checkin_date
                    and self._is_cycle_clean_day(checkin_date, default_date))
                init_eco = is_ecoplan_guest and not init_duvet
            init_kind = prev_status if init_eco and prev_status else '×'

            row_frame = ttk.Frame(scrollable_frame)
            row_frame.pack(fill="x", pady=4)
            _configure_row_columns(row_frame)

            eco_var = tk.BooleanVar(value=init_eco)
            duvet_var = tk.BooleanVar(value=init_duvet)
            kind_var = tk.StringVar(value=init_kind)
            eco_vars[room_number] = eco_var
            duvet_vars[room_number] = duvet_var
            kind_vars[room_number] = kind_var

            # エコとデュベは同じ日に両立しないので、片方を入れたらもう片方を外す
            def on_eco_changed(*_, ev=eco_var, dv=duvet_var, rn=room_number):
                if ev.get() and dv.get():
                    dv.set(False)
                update_count()
                update_auto_mark(rn)

            def on_duvet_changed(*_, ev=eco_var, dv=duvet_var, rn=room_number):
                if dv.get() and ev.get():
                    ev.set(False)
                update_count()
                update_auto_mark(rn)

            eco_var.trace('w', on_eco_changed)
            duvet_var.trace('w', on_duvet_changed)

            eco_check = ttk.Checkbutton(row_frame, variable=eco_var)
            eco_check.grid(row=0, column=1, padx=5, sticky="w")
            duvet_check = ttk.Checkbutton(row_frame, variable=duvet_var)
            duvet_check.grid(row=0, column=0, padx=5, sticky="w")
            check_widgets[room_number] = (eco_check, duvet_check)
            ttk.Label(row_frame, text=room_number, anchor="w", font=row_font).grid(row=0, column=2, padx=5, sticky="w")
            ttk.Label(row_frame, text=guest_name, anchor="w", font=row_font).grid(row=0, column=3, padx=5, sticky="w")
            ttk.Label(row_frame, text=checkin_text, anchor="w", font=row_font).grid(row=0, column=4, padx=5, sticky="w")
            ttk.Label(row_frame, text=checkout_text, anchor="w", font=row_font).grid(row=0, column=5, padx=5, sticky="w")
            ttk.Label(row_frame, text='エコプラン' if is_ecoplan_guest else '',
                      anchor="w", font=row_font).grid(row=0, column=6, padx=5, sticky="w")

            # エコ区分プルダウン（デフォルトは「×」）
            # ※以前コンボボックスでホイールにより値が変わる事故があったため、
            #   readonly＋ホイールイベント無効化で対策している
            kind_combo = ttk.Combobox(
                row_frame, textvariable=kind_var,
                values=self.ECO_KIND_CHOICES,
                state='readonly', width=16, font=row_font)
            kind_combo.grid(row=0, column=7, padx=5, sticky="w")
            kind_combo.bind("<MouseWheel>", lambda e: "break")
            kind_combo.bind("<Button-4>", lambda e: "break")
            kind_combo.bind("<Button-5>", lambda e: "break")

            # エコ区分を選んだ＝エコにする意思表示としてエコにチェックを入れる
            kind_combo.bind('<<ComboboxSelected>>',
                            lambda e, ev=eco_var: ev.set(True))

            kind_combos[room_number] = kind_combo

            auto_label = ttk.Label(row_frame, text="", anchor="w", font=row_font)
            auto_label.grid(row=0, column=8, padx=5, sticky="w")
            auto_labels[room_number] = auto_label
            update_auto_mark(room_number)

            # 2部屋ごとに区切り線を引いて見やすくする
            # （最後の行の後ろには線を入れない）
            if (idx + 1) % 2 == 0 and (idx + 1) < len(eco_rooms):
                ttk.Separator(scrollable_frame, orient="horizontal").pack(
                    fill="x", padx=2, pady=3
                )

        # 初期カウント更新
        update_count()

        # 清掃日を変えたら自動〇の表示を判定し直す
        for var in date_vars.values():
            var.trace('w', update_all_auto_marks)

        def register_rooms():
            """在室の全部屋を、選択内容に応じて登録する。

            エコにチェック → エコ区分（×／エコドア／清掃不要・アメ不要）、
            デュベにチェック → 〇、
            どちらもなし → 通常清掃（実C/I日から3日ごとの日だけ自動で〇）。
            DBは当日分だけを持つので、登録のたびに全消去してから入れ直す。"""
            target_date = get_target_date()
            if target_date is None:
                messagebox.showerror("エラー", "正しい清掃日を入力してください。",
                                     parent=dialog)
                return

            records = []
            eco_count = duvet_count = auto_count = checkout_count = 0
            for room_info in eco_rooms:
                room_number = room_info['room']
                guest_info = guest_name_map.get(room_number, {})
                guest_name, is_ecoplan_guest, _, _ = guest_fields(room_number)

                if is_checkout_day(room_number, target_date):
                    # 清掃日にC/Oする部屋は×も〇も付けない
                    middle_status = self.NORMAL_CLEAN_LABEL
                    checkout_count += 1
                elif eco_vars[room_number].get():
                    middle_status = kind_vars[room_number].get() or '×'
                    eco_count += 1
                elif duvet_vars[room_number].get():
                    middle_status = self.DAILY_CLEAN_LABEL
                    duvet_count += 1
                else:
                    middle_status = self.NORMAL_CLEAN_LABEL

                # C/I日・泊数は部屋ごとの実際のC/I日／C/O日を使う
                # （予約CSV・予約詳細CSVから取得。取れなければ清掃日基準）
                room_checkin, days = self._resolve_stay_period(
                    guest_info, target_date)
                schedule = self._build_day_schedule(
                    room_checkin, days, middle_status, target_date)
                if schedule.get(target_date) == self.NORMAL_CLEAN_MARK and (
                        middle_status == self.NORMAL_CLEAN_LABEL):
                    auto_count += 1

                records.append({
                    'room': room_number,
                    'guest': guest_name,
                    'date': room_checkin,
                    'days': days,
                    'ecodoor': (middle_status == 'エコドア'),
                    'ecoplan': is_ecoplan_guest,
                    'middle_status': middle_status,
                    'schedule': schedule,
                })

            try:
                self._save_records_to_db(records, target_date)
            except Exception as e:
                self.conn.rollback()
                messagebox.showerror(
                    "エラー",
                    f"データベースへの保存に失敗しました: {e}\n"
                    "もう一度「登録」を押してください。", parent=dialog)
                return

            close_dialog()

            output_now = messagebox.askyesno(
                "登録完了",
                f"{target_date.strftime('%Y年%m月%d日')} 分を登録しました。\n\n"
                f"エコ: {eco_count}件\n"
                f"デュベ（〇）: {duvet_count}件\n"
                f"自動の〇（C/I日から{self.NORMAL_CLEAN_CYCLE}日ごと）: {auto_count}件\n"
                f"清掃日にC/O: {checkout_count}件\n"
                f"在室: {len(records)}件\n\n"
                "続けてアーニング表とエコ部屋リストを出力しますか？")
            if output_now:
                # 直前に読み込んだ部屋状態CSVを再利用する
                self.output_earning_table(self.last_room_status_csv)
                self.export_eco_room_list()

        ttk.Button(button_frame, text="登録", command=register_rooms, style="Eco.TButton").pack(side="left", padx=5)
        ttk.Button(button_frame, text="キャンセル", style="Eco.TButton",
                   command=close_dialog).pack(side="left", padx=5)

        # 一覧の必要幅を実測してウィンドウ幅を再調整（右端の列の見切れ防止）
        dialog.update_idletasks()
        need_w = (scrollable_frame.winfo_reqwidth()
                  + scrollbar.winfo_reqwidth() + 40)  # 外枠padding15×2 + 余裕
        self._fit_toplevel_to_screen(dialog, max(960, need_w), 1000,
                                     min_w=560, min_h=420)

    def _save_records_to_db(self, records, target_date):
        """当日分のレコードを rooms / cleaning_schedule へ保存して commit する。

        DBは当日分だけを持つので、既存の登録を全消去してから入れ直す。
        失敗時の rollback は呼び出し側で行う。"""
        self.clear_daily_data()
        cursor = self.conn.cursor()
        for record in records:
            cursor.execute(
                '''INSERT OR REPLACE INTO rooms
                   (room_number, guest_name, check_in_date, cleaning_days,
                    is_ecodoor, is_ecoplan, middle_status)
                   VALUES (?, ?, ?, ?, ?, ?, ?)''',
                (record['room'], record['guest'],
                 record['date'].strftime('%Y-%m-%d'),
                 record['days'], record['ecodoor'], record['ecoplan'],
                 record['middle_status']))

            for cleaning_date, status in record['schedule'].items():
                cursor.execute("INSERT INTO cleaning_schedule VALUES (?, ?, ?)",
                               (record['room'],
                                cleaning_date.strftime('%Y-%m-%d'), status))

        cursor.execute(
            "INSERT OR REPLACE INTO app_state (key, value) VALUES ('target_date', ?)",
            (target_date.strftime('%Y-%m-%d'),))
        self.conn.commit()

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

        ※ middle_status='通常清掃'／'毎日デュベ清掃' の部屋は rooms に登録されていても
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
            if middle_status in self.NON_ECO_MIDDLE_STATUSES:
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
        （登録直後の自動実行で、直前に読み込んだCSVを再利用するため）。
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

    def export_eco_room_list(self, target_date=None):
        """登録済みの清掃日について、その日エコ対応する部屋
        （×／エコドア／清掃不要・アメ不要）を部屋番号順で
        「エコ部屋リスト_YYYYMMDD.xlsx」に出力して開く。

        DBは当日分だけを持つので、日付は登録時の清掃日を使う
        （target_date を渡した場合はその日付）。
        〇（デュベ・3日周期の清掃日）の部屋は含めない。
        戻り値は出力ファイルパス（該当なし・失敗時は None）。"""
        if target_date is None:
            target_date = self._get_registered_date()
        if target_date is None:
            messagebox.showinfo(
                "情報",
                "登録データがありません。\n"
                "先に「CSV読込」でエコ部屋を登録してください。")
            return None

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

    def run(self):
        """アプリケーションの実行"""
        self.root.protocol("WM_DELETE_WINDOW", lambda: (self.conn.close(), self.root.destroy()))
        self.root.mainloop()


if __name__ == "__main__":
    app = HotelCleaningSystem()
    app.run()