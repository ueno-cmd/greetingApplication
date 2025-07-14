import tkinter as tk
from tkinter import messagebox
from tkinter import filedialog
import csv
import datetime

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

        # 設定（operation_menu）
        operation_menu.add_command(label="終了", command=root.quit)
        operation_menu.add_command(label="リロード", command=self.reload_data)
        operation_menu.add_command(label="出席状況リセット", command=self.reset_participants)

        # ファイル操作（file_menu）
        file_menu.add_command(label="新規社員追加", command=self.add_employee_dialog)
        file_menu.add_command(label="出席結果エクスポート", command=self.export_attendance_report)
        file_menu.add_command(label="社員情報編集", command=self.edit_employee_dialog)

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
        self.search_button = tk.Button(self.root, text="登録", font=("Arial", 14), command=self.register_and_clear)
        self.root.bind("<Return>", 
                       lambda event: self.register_and_clear())
        self.search_button.pack(pady=10)

    def count_participants(self):
        """参加人数をカウントする関数"""
        try:
            participants_count = 0
            with open(csv_file_path, mode='r', newline='', encoding='utf-8-sig') as file:
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
            with open(csv_file_path, mode='r', newline='', encoding='utf-8-sig') as file:
                reader = csv.DictReader(file)
                rows = list(reader)

                for row in rows:
                    if int(row['number']) == employee_number_int:
                        # 名前確認のメッセージを表示
                        confirm = messagebox.askyesno(
                            "確認", f"あなたは {row['name']} さんで正しいですか？")

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

    def register_and_clear(self):
        """出席登録と入力フィールドクリアを統合"""
        self.update_participant_status()
        self.clear_entry_field()

    def clear_entry_field(self):
        """入力フィールドをクリア"""
        self.employee_number_entry.delete(0, tk.END)

    def export_attendance_report(self):
        """出席結果をCSVでエクスポート"""
        try:
            # ファイル保存ダイアログを表示
            file_path = filedialog.asksaveasfilename(
                defaultextension=".csv",
                filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
            )
            
            # 拡張子がない場合は.csvを追加
            if file_path and not file_path.endswith('.csv'):
                file_path += '.csv'
            
            if not file_path:
                return
            
            # CSVファイルからデータを読み込み
            with open(csv_file_path, mode='r', newline='', encoding='utf-8-sig') as file:
                reader = csv.DictReader(file)
                all_employees = list(reader)
            
            # 出席者と未出席者に分ける
            attended = [emp for emp in all_employees if emp['participated'] == '1']
            not_attended = [emp for emp in all_employees if emp['participated'] == '0']
            
            # レポートファイルを作成
            with open(file_path, mode='w', newline='', encoding='utf-8-sig') as file:
                writer = csv.writer(file)
                
                # ヘッダー情報
                writer.writerow(["出席管理レポート"])
                writer.writerow([f"作成日時: {datetime.datetime.now().strftime('%Y年%m月%d日 %H:%M:%S')}"])
                writer.writerow([f"総社員数: {len(all_employees)}人"])
                writer.writerow([f"出席者数: {len(attended)}人"])
                writer.writerow([f"未出席者数: {len(not_attended)}人"])
                writer.writerow([])  # 空行
                
                # 出席者リスト
                writer.writerow(["【出席者一覧】"])
                writer.writerow(["社員番号", "氏名"])
                for emp in attended:
                    writer.writerow([emp['number'], emp['name']])
                
                writer.writerow([])  # 空行
                
                # 未出席者リスト
                writer.writerow(["【未出席者一覧】"])
                writer.writerow(["社員番号", "氏名"])
                for emp in not_attended:
                    writer.writerow([emp['number'], emp['name']])
            
            messagebox.showinfo("エクスポート完了", f"出席結果をエクスポートしました:\n{file_path}")
            
        except Exception as e:
            messagebox.showerror("エラー", f"エクスポートに失敗しました: {str(e)}")

    def reset_participants(self):
        """全員の出席状況をリセット"""
        # 確認ダイアログを表示
        confirm = messagebox.askyesno("確認", "全員の出席状況をリセットしますか？\nこの操作は取り消せません。")
        
        if not confirm:
            return
        
        try:
            # CSVファイルを読み込み
            with open(csv_file_path, mode='r', newline='', encoding='utf-8-sig') as file:
                reader = csv.DictReader(file)
                rows = list(reader)
            
            # 全員の出席状況を0にリセット
            for row in rows:
                row['participated'] = '0'
            
            # CSVファイルに書き戻し
            fieldnames = ['number', 'name', 'participated']
            with open(csv_file_path, mode='w', newline='', encoding='utf-8-sig') as file:
                writer = csv.DictWriter(file, fieldnames=fieldnames)
                writer.writeheader()
                writer.writerows(rows)
            
            # 参加人数を再カウント
            self.count_participants()
            
            messagebox.showinfo("リセット完了", "全員の出席状況をリセットしました。")
            
        except Exception as e:
            messagebox.showerror("エラー", f"リセットに失敗しました: {str(e)}")

    def add_employee_dialog(self):
        """新規社員追加ダイアログ"""
        # 新しいウィンドウを作成
        dialog = tk.Toplevel(self.root)
        dialog.title("新規社員追加")
        dialog.geometry("300x200")
        dialog.resizable(False, False)
        
        # 入力フィールド
        tk.Label(dialog, text="社員番号:", font=("Arial", 12)).pack(pady=5)
        number_entry = tk.Entry(dialog, font=("Arial", 12))
        number_entry.pack(pady=5)
        
        tk.Label(dialog, text="氏名:", font=("Arial", 12)).pack(pady=5)
        name_entry = tk.Entry(dialog, font=("Arial", 12))
        name_entry.pack(pady=5)
        
        def add_employee():
            """社員を追加する処理"""
            try:
                # 入力値を取得
                number = number_entry.get().strip()
                name = name_entry.get().strip()
                
                if not number or not name:
                    messagebox.showerror("エラー", "社員番号と氏名を入力してください")
                    return
                
                # 社員番号が数値かチェック
                try:
                    int(number)
                except ValueError:
                    messagebox.showerror("エラー", "社員番号は数値で入力してください")
                    return
                
                # 既存データを読み込み
                with open(csv_file_path, mode='r', newline='', encoding='utf-8-sig') as file:
                    reader = csv.DictReader(file)
                    rows = list(reader)
                
                # 重複チェック
                for row in rows:
                    if row['number'] == number:
                        messagebox.showerror("エラー", f"社員番号 {number} は既に存在します")
                        return
                
                # 新規社員を追加
                new_employee = {'number': number, 'name': name, 'participated': '0'}
                rows.append(new_employee)
                
                # CSVファイルに書き戻し
                fieldnames = ['number', 'name', 'participated']
                with open(csv_file_path, mode='w', newline='', encoding='utf-8-sig') as file:
                    writer = csv.DictWriter(file, fieldnames=fieldnames)
                    writer.writeheader()
                    writer.writerows(rows)
                
                messagebox.showinfo("完了", f"社員 {name} さん（番号: {number}）を追加しました")
                dialog.destroy()
                
                # 参加人数を再カウント
                self.count_participants()
                
            except Exception as e:
                messagebox.showerror("エラー", f"社員追加に失敗しました: {str(e)}")
        
        # ボタン
        button_frame = tk.Frame(dialog)
        button_frame.pack(pady=20)
        
        tk.Button(button_frame, text="追加", command=add_employee, 
                 font=("Arial", 12)).pack(side=tk.LEFT, padx=5)
        tk.Button(button_frame, text="キャンセル", command=dialog.destroy, 
                 font=("Arial", 12)).pack(side=tk.LEFT, padx=5)
        
        # エンターキーでの実行
        dialog.bind('<Return>', lambda event: add_employee())
        
        # 最初のフィールドにフォーカス
        number_entry.focus()

    def edit_employee_dialog(self):
        """既存社員情報編集ダイアログ"""
        # 新しいウィンドウを作成
        dialog = tk.Toplevel(self.root)
        dialog.title("社員情報編集")
        dialog.geometry("300x250")
        dialog.resizable(False, False)
        
        # 検索フィールド
        tk.Label(dialog, text="検索する社員番号:", font=("Arial", 12)).pack(pady=5)
        search_entry = tk.Entry(dialog, font=("Arial", 12))
        search_entry.pack(pady=5)
        
        # 編集フィールド（初期は無効化）
        tk.Label(dialog, text="氏名:", font=("Arial", 12)).pack(pady=(20, 5))
        name_entry = tk.Entry(dialog, font=("Arial", 12), state='disabled')
        name_entry.pack(pady=5)
        
        current_employee = {'number': '', 'name': '', 'participated': '0'}
        
        def search_employee():
            """社員を検索する処理"""
            try:
                search_number = search_entry.get().strip()
                
                if not search_number:
                    messagebox.showerror("エラー", "社員番号を入力してください")
                    return
                
                # 社員番号が数値かチェック
                try:
                    int(search_number)
                except ValueError:
                    messagebox.showerror("エラー", "社員番号は数値で入力してください")
                    return
                
                # 既存データを読み込み
                with open(csv_file_path, mode='r', newline='', encoding='utf-8-sig') as file:
                    reader = csv.DictReader(file)
                    rows = list(reader)
                
                # 社員を検索
                for row in rows:
                    if row['number'] == search_number:
                        current_employee['number'] = row['number']
                        current_employee['name'] = row['name']
                        current_employee['participated'] = row['participated']
                        
                        # 編集フィールドを有効化して値を設定
                        name_entry.config(state='normal')
                        name_entry.delete(0, tk.END)
                        name_entry.insert(0, row['name'])
                        
                        messagebox.showinfo("検索結果", f"社員 {row['name']} さんが見つかりました")
                        return
                
                messagebox.showerror("エラー", f"社員番号 {search_number} は見つかりません")
                
            except Exception as e:
                messagebox.showerror("エラー", f"検索に失敗しました: {str(e)}")
        
        def update_employee():
            """社員情報を更新する処理"""
            try:
                if not current_employee['number']:
                    messagebox.showerror("エラー", "まず社員を検索してください")
                    return
                
                new_name = name_entry.get().strip()
                
                if not new_name:
                    messagebox.showerror("エラー", "氏名を入力してください")
                    return
                
                # 既存データを読み込み
                with open(csv_file_path, mode='r', newline='', encoding='utf-8-sig') as file:
                    reader = csv.DictReader(file)
                    rows = list(reader)
                
                # 該当社員の情報を更新
                updated = False
                for row in rows:
                    if row['number'] == current_employee['number']:
                        row['name'] = new_name
                        updated = True
                        break
                
                if not updated:
                    messagebox.showerror("エラー", "社員が見つかりません")
                    return
                
                # CSVファイルに書き戻し
                fieldnames = ['number', 'name', 'participated']
                with open(csv_file_path, mode='w', newline='', encoding='utf-8-sig') as file:
                    writer = csv.DictWriter(file, fieldnames=fieldnames)
                    writer.writeheader()
                    writer.writerows(rows)
                
                messagebox.showinfo("完了", f"社員番号 {current_employee['number']} の情報を更新しました")
                dialog.destroy()
                
                # 参加人数を再カウント
                self.count_participants()
                
            except Exception as e:
                messagebox.showerror("エラー", f"更新に失敗しました: {str(e)}")
        
        # ボタン
        button_frame1 = tk.Frame(dialog)
        button_frame1.pack(pady=10)
        tk.Button(button_frame1, text="検索", command=search_employee, font=("Arial", 12)).pack()
        
        button_frame2 = tk.Frame(dialog)
        button_frame2.pack(pady=10)
        tk.Button(button_frame2, text="更新", command=update_employee, 
                 font=("Arial", 12)).pack(side=tk.LEFT, padx=5)
        tk.Button(button_frame2, text="キャンセル", command=dialog.destroy, 
                 font=("Arial", 12)).pack(side=tk.LEFT, padx=5)
        
        # エンターキーでの実行（検索フィールドにフォーカスがある場合）
        search_entry.bind('<Return>', lambda event: search_employee())
        
        # 最初のフィールドにフォーカス
        search_entry.focus()


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
