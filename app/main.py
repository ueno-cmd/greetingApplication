import tkinter as tk
from tkinter import messagebox
import csv
import os
import sys

# csvファイルのパスを取得
csv_file_path = "data/employee_list.csv"

class App:
    def __init__(self, root):
        """GUIの設定"""
        self.root = root
        self.root.title("出席アプリ")
        self.root.geometry("500x300")

        # メニューバーの作成
        menubar = tk.Menu(root)
        root.config(menu=menubar)

        # メニューアイテムの追加
        operation_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="設定", menu=operation_menu)
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="csvファイル操作", menu=file_menu)
        blank_b = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Coming soon", menu=blank_b)
        blank_c = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Coming soon", menu=blank_c)

        # 設定（operation_menu）
        operation_menu.add_command(label="終了", command=root.quit)
        operation_menu.add_command(label="リロード", command=self.reload_data)

        # ファイル操作（file_menu）
        file_menu.add_command(label="csv入力(coming soon)")
        file_menu.add_command(label="csv出力(coming soon)")
        file_menu.add_command(label="csv初期化(coming soon)")

        # 社員番号ラベル
        self.employee_number_label = tk.Label(self.root, text="社員番号:", font=("Arial", 14))
        self.employee_number_label.pack()

        # 社員番号入力欄
        self.employee_number_entry = tk.Entry(self.root, font=("Arial", 14))
        self.employee_number_entry.pack(pady=10)

        # 参加人数の表示用ラベル
        self.participant_count_label = tk.Label(self.root, text="", font=("Arial", 14))
        self.participant_count_label.pack(pady=10)

        # 初期データの読み込み
        self.count_participants()

        # csv登録ボタン
        self.search_button = tk.Button(self.root, text="登録", font=("Arial", 14), command=self.update_participant_status)
        self.root.bind("<Return>", lambda event: self.update_participant_status())
        self.root.bind("<Return>", lambda event: self.clear_entry_and_register(), "+")
        self.search_button.pack(pady=10)

    def count_participants(self):
        """参加人数をカウントする関数"""
        try:
            participants_count = 0
            with open(csv_file_path, mode='r+', newline='', encoding='utf-8-sig') as file:
                reader = csv.DictReader(file)
                for row in reader:
                    if row['participated'] == "1":
                        participants_count += 1
            self.participant_count_label.config(text=f"参加人数: {participants_count} 人")
        except Exception as e:
            print(f"エラーが発生しました:{e}")

    def update_participant_status(self):
        """CSVファイルに参加者情報を更新する"""
        employee_number_str = self.employee_number_entry.get()
        try:
            employee_number_int = int(employee_number_str)
        except ValueError:
            messagebox.showerror("エラー", "無効な社員番号です")
            return

        try:
            with open(csv_file_path, mode='r+', newline='', encoding='utf-8-sig') as file:
                reader = csv.DictReader(file)
                rows = list(reader)

                for row in rows:
                    if int(row['number']) == employee_number_int:
                        # 名前確認のメッセージを表示
                        confirm = messagebox.askyesno("確認",
                                                      f"あなたは {row['name']} さんで正しいですか？")

                        if confirm:  # ユーザーが「はい」を選択した場合
                            row['participated'] = "1"
                            messagebox.showinfo("登録完了", f"ようこそ {row['name']} さん")

                            # CSVファイルに変更を書き戻す
                            fieldnames = ['number', 'name', 'participated']
                            with open(csv_file_path, mode='w', newline='',
                                      encoding='utf-8-sig') as file:
                                writer = csv.DictWriter(file, fieldnames=fieldnames)
                                writer.writeheader()
                                writer.writerows(rows)

                            # 参加人数を再カウント
                            self.count_participants()
                            return
                        else:
                            messagebox.showinfo("キャンセル", "操作がキャンセルされました")
                            return

                messagebox.showerror("エラー", "該当の社員番号は見つかりません")
        except Exception as e:
            print(f"エラーが発生しました:{e}")

    def reload_data(self):
        """データの再読み込み"""
        self.count_participants()

    def clear_entry_and_register(self):
        """入力フィールドをクリア"""
        self.employee_number_entry.delete(0, tk.END)

class MainWindow(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("出席アプリ")
        self.geometry("500x300")

        # アプリケーションのインスタンス作成
        self.app_instance = App(self)

    def run_app(self):
        self.mainloop()

if __name__ == "__main__":
    app = MainWindow()
    app.run_app()
