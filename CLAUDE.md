# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## 概要

ホテルの客室清掃管理システム（tkinter デスクトップアプリ）。CSVから宿泊予約データを取り込み、その日のエコ（×）・デュベ（〇）部屋を選んでSQLiteに登録し、Excel（アーニング表・エコ部屋リスト）を出力する。UI・コミットメッセージ・コード内コメントはすべて日本語。

## 実行・ビルド

```bash
# 実行（venvは .venv、外部依存は openpyxl のみ）
source .venv/bin/activate
python EcoRoomClean.py
```

- テスト・リンターは存在しない。requirements.txt もない（READEME.md に記載はあるが未作成）。
- Windows用EXEは GitHub Actions（`.github/workflows/build-windows.yml`）が main への push 時に PyInstaller でビルドし、Artifact として出力する。
- README はファイル名が `READEME.md`（typo のままのファイル名）。実行ファイルは `EcoRoomClean.py`。

## アーキテクチャ

コードは単一ファイル `EcoRoomClean.py`（約1,700行）の `HotelCleaningSystem` クラス1つに集約されている。エントリポイントは末尾の `if __name__ == "__main__"`。

### 処理の流れ

1. **起動**: `__init__` → `init_database` → `setup_gui`（メイン画面のボタンは CSV読込／アーニング表出力／エコ部屋リスト）
2. **CSV読込** (`import_csv`): `detect_csv_type` で「部屋状態CSV」（2列目=部屋番号、7列目=状態、3＝在室が選択対象）「予約CSV」（11列目=部屋番号、12列目=宿泊者名_プラン）「予約詳細CSV」（C/I・C/O日入り）を判別。エンコーディングは cp932 / shift_jis / utf-8 を自動判定。**読み込んだ時点で `clear_daily_data` によりDBの登録内容を全消去する**
3. **エコ・デュベ選択** (`show_csv_room_selection_dialog`): 在室部屋ごとにエコ（×／エコドア／清掃不要・アメ不要）とデュベ（〇）を排他でチェック。どちらも選ばない部屋は通常清掃として、実C/I日から3日ごとの日だけ自動で〇（`_is_cycle_clean_day`）。清掃日（×・〇を付ける日）は**部屋状態CSVの日付の翌日**（CSVは清掃日の前日時点の部屋状態）。登録内容は `_build_day_schedule` で清掃日1日分だけ作る
4. **アーニング表出力** (`output_earning_table`): テンプレート `アーニング表.xlsx` を `_find_earning_template` で探索(既定名 → 「指示書」を含むシートを2つ以上持つ .xlsx を自動検出。.xlsm は除外)し、`アーニング表_出力_YYYYMMDD.xlsx` として出力
5. **エコ部屋リスト出力** (`export_eco_room_list`): 登録済みの清掃日のエコ部屋を `エコ部屋リスト_YYYYMMDD.xlsx` に出力

### データベース（SQLite: `hotel_cleaning.db`）

DBは**当日分だけ**を持つ一時置き場。エコ・デュベは翌日に引き継がず、履歴も残さない（バックアップ・部屋編集・エコ票・起動時のC/O削除は廃止済み）。同じ日付のCSVを読み直したときだけ、直前の選択を選択画面の初期表示に使う。

- `rooms`: room_number (PK), guest_name, check_in_date, cleaning_days, is_ecodoor, is_ecoplan, middle_status（×／エコドア／清掃不要・アメ不要／毎日デュベ清掃＝デュベを選んだ部屋／通常清掃＝未選択）
- `cleaning_schedule`: room_number, cleaning_date, cleaning_status（清掃日の ×・エコドア・清掃不要・アメ不要・〇 と、C/I日・C/O日の行のみ）
- `app_state`: key, value（`target_date`＝登録済みデータの清掃日）
- `shared_csv`: 直前に読み込んだCSVの中身。清掃管理システム（Java）側が同じDBを読むため、テーブル構成や middle_status の文字列は変えないこと
- スキーマ変更は `init_database` 内で `PRAGMA table_info` により列の有無を確認して `ALTER TABLE` で追加する方式（既存DBとの互換維持）。この慣例に従うこと

### 注意点

- `hotel_cleaning.db` は生成物だが git 管理されており、実行のたびに差分が出る（`hotel_cleaning_now.xlsx` は廃止したエコ票の残りで、もう生成されない）
- データファイル（DB・Excel）はすべて実行ディレクトリ直下に読み書きされる
- macOS では tkinter のIME警告抑制のため `TK_SILENCE_DEPRECATION=1` を設定している