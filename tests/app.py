import tkinter as tk
from tkinter import messagebox
import csv
import os
import sys

# リソースパス取得関数
def resource_path(relative_path):
    """概要
    詳細説明
    ・PyInstallerでexe化した際に、リソースファイルへのパスを取得する。
    ・:param relative_path:
    ・:return:
    """
    try:
        # PyInstallerでパッケージ化された場合の一時ディレクトリ
        base_path = sys.MEIPASS
    except Exception:
        # 通常のPython実行時
        base_path = os.path.abspath("../app")
    return os.path.join(base_path, relative_path)

# csvファイルのパスを取得
csv_file_path = resource_path("../app/data/employee_list.csv")

class App:
    def __init__(self, root):
        """GUIの設定

        変数
        :param root:引数
        :param self.root:
        :param menubar:メニューバー
        :param operation_menu:設定用のメニュー
        :param file_menu:CSV入出力操作用のメニュー
        :param blank_b:CSV内部を編集するメニュー（未実装）
        :param blank_c:未定
        """
        # メンバー変数の定義
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

        # Fileメニューにメニューアイテムを追加
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

        # 参加済みのフラグを表示するラベル
        self.participant_status_label = tk.Label(self.root, text="", font=("Arial", 24))
        self.participant_status_label.pack()

        # csv内検索ボタン
        self.search_button = tk.Button(self.root, text="登録", font=("Arial", 14))
        self.root.bind("<Return>", lambda event: self.update_participant_status())
        self.root.bind("<Return>", lambda event: self.clear_entry_and_register(), "+")
        self.search_button.pack(pady=10)

        # 参加人数の表示用ラベル
        self.participant_count_label = tk.Label(self.root, text="", font=("Arial", 14))
        self.count_participants()
        self.participant_count_label.pack(pady=10)

        # 初期データの読み込み
        self.count_participants()

        # 定期的にデータを更新する
        self.root.after(10000, self.count_participants)

    # csv操作用関数
    def count_participants(self):
        """参加人数をカウントする関数
        詳細説明
        csvファイル内のparticipatedで、１の数を数えて表示する関数

        変数
        :param participants_count: 参加人数
        :param reader: csvファイルの代入

        例外処理
        ファイルが見つからない場合、例外エラーが発生する
        """
        try:
            participants_count = 0
            while True:
                with open(csv_file_path, mode='r+', newline='', encoding='utf-8-sig') as file:
                    reader = csv.DictReader(file)
                    for row in reader:
                        if row['participated'] == "1":
                            participants_count += 1
                # ファイル更新後にファイルを閉じる処理
                break
        except Exception as e:
            print(f"エラーが発生しました:{e}")
        finally:
            self.participant_count_label.config(text=f"ただ今の参加人数は... {participants_count} 人です")

    def update_participant_status(self):
        """CSVファイルに書き込む関数
        csvファイルを開き、入力された社員番号に基づいてparticipatedに
        設定されている０の数値を1に変更する処理
        処理で変更を加えたcsvファイルを閉じて保存する

        変数
        :param employee_number_str:ファイル内の社員番号を文字列にして受け取る
        :param employee_number_int:文字列の社員番号を数値型に変える
        :param reader:ファイルを格納する変数
        :param rows:ファイルを格納した変数を格納する変数
        :param number_int:社員番号を格納する変数
        :param fieldnames:csvのヘッダーに記入するための要素を格納する変数
        :param writer:csvファイルを格納する変数

        例外
        :ValueError:
        :Exception:

        :return:入力した社員番号に対し、メッセージボックスに社員名を返す
        """

        employee_number_str = self.employee_number_entry.get()
        employee_number_int = int(employee_number_str)

        try:
            # CSVファイルからデータを読み込む
            with open(csv_file_path, mode='r+', newline='', encoding='utf-8-sig') as file:
                reader = csv.DictReader(file)
                rows = list(reader)  # ファイルの全行をリストに格納

                # ユーザーが入力した社員番号に対応する行を検索
                for row in rows:
                    number_int = int(row['number'])
                    if number_int == employee_number_int:
                        # 参加済みのフラグを「0→1」に更新
                        row['participated'] = "1"
                        messagebox.showinfo("登録完了", f"ようこそ{row['name']}さん")

                        # 更新した行をファイルに書き戻す
                        fieldnames = ['number', 'name', 'participated']
                        with open(csv_file_path, mode='w', newline='', encoding='utf-8-sig') as file:
                            writer = csv.DictWriter(file, fieldnames=fieldnames)
                            writer.writeheader()
                            for r in rows:
                                writer.writerow(r)

                        print("参加済みのフラグが正常に更新されました。")
                        return  # 一致した社員名を見つけ次のループを終了

                print("該当の社員番号は見つかりませんでした。")
        except ValueError:
            print("入力された社員番号が有効な数字ではありません。")
        except Exception as e:
            print(f"エラーが発生しました：{e}")

        # 該当の社員番号が見つからない場合
        messagebox.showinfo("エラー", "登録外番号です")

    # この機能を呼び出すことを禁ずる。
    def reset_participants(self):
        """参加者のリセット
        csvファイルのparticipated列を０に戻す処理
        file_menuのCSV入出力操作用のメニューに実装することでリセット
        リセットすることで次の懇親会に繰り返し使うことができるようにする

        変数
        :param reader:ファイルを格納する変数
        :param fieldnames:csvのヘッダーに記入するための要素を格納する変数

        :return:削除の処理を返す

        例外
        Exception
        """
        try:
            with open(csv_file_path, mode='r+', newline='', encoding='utf-8-sig') as file:
                reader = csv.DictReader(file)

                file.seek(0)
                file.truncate()

                fieldnames = ['number', 'name', 'participated']

                writer = csv.DictWriter(file, fieldnames=fieldnames)
                writer.writeheader()

                for row in reader:
                    row['participated'] = "0"
                    writer.writerow(row)

            print("初期化処理が成功しました")
        except Exception as e:
            print(f"初期化処理に失敗しました:{e}")

    # ボタン操作用関数
    def reload_data(self):
        """概要
        詳細説明
        人数カウント表示用の関数
        """
        self.count_participants()

    def clear_entry_and_register(self):
        """概要
        詳細説明
        登録ボタン押下のイベントの１つ
        Entryウィジェットの内容を削除
        """
        self.employee_number_entry.delete(0, tk.END)
        print("登録処理を行いました。")
