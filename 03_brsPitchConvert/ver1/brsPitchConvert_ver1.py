import chardet
import os
import tkinter as tk
from tkinter import filedialog
import re


class Application(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("App_BRS_FileConvert_ver-1.0")

        # 　０　ベース画面　サイズ設定
        self.BASE_WINDOW_WIDTH = 1200  # << アプリサイズ（横幅）
        self.BASE_WINDOW_HEIGHT = 600  # << アプリサイズ（縦幅）
        self.BASE_WINDOW_POS_X = 25  # << アプリの左上位置（液晶画面の左端からの距離）
        self.BASE_WINDOW_POS_Y = 25  # << アプリの左上位置（液晶画面の上側からの距離）
        self.geometry(
            f"{self.BASE_WINDOW_WIDTH}x{self.BASE_WINDOW_HEIGHT}+{self.BASE_WINDOW_POS_X}+{self.BASE_WINDOW_POS_Y}"
        )

        # 　０　メインフレーム設置（ベース画面と同一サイズ）
        self.base_frame = tk.Frame(
            self,
            width=self.BASE_WINDOW_WIDTH,
            height=self.BASE_WINDOW_HEIGHT,
            bd=5,
            relief="ridge",
        )
        self.base_frame.propagate(False)
        self.base_frame.pack()

        self.initial_dir = os.getcwd()
        self.working_dir = self.initial_dir

        # 　０　アプリの基本色
        self.tk_setPalette(background="#EBF4FA")

        self.button_color = "#E6DCDC"

        self.error_window_color = "#E3E9EC"
        self.error_window_button_color = "#E6DCDC"

        # 　１　タイトル設置【Title】
        self.clsApp01_Title_label01 = tk.Label(
            self.base_frame,
            text="BRSファイルピッチ一括変換用",
            font=("BIZ UDPゴシック", 20, "bold"),
            height=2,
        )  # << タイトル高さ
        self.clsApp01_Title_label01.place(
            relx=0.5,  # << タイトル設置の横位置（中間地点）
            y=20,  # << タイトル設置の縦位置
            anchor=tk.N,
        )

        # 　２－１　説明文設置【Exp】
        self.clsApp02_Exp_POS_X = 50  # << 説明文タイトルの横位置
        self.clsApp02_Exp_POS_Y = 100  # << 説明文タイトルの縦位置
        self.clsApp02_Exp_SIDE_SPACE = 25  # << 説明文本文の左詰めスペース
        self.clsApp02_Exp_LINE_SPACE = 35  # << 説明文の行間
        self.clsApp02_Exp_FONT = ("BIZ UDPゴシック", 14)

        self.clsApp02_exp_list = [
            "○ アプリ概要",
            "BRSファイルの項目であるピッチを一括で変換します。",
            "読込フォルダを指定後、ピッチを指定して、変換実行ボタンをクリックしてください。",
            "指定した読込フォルダ内のファイル全てに対し、変換用フォルダを作成して保存します。",
        ]

        for num, ind_exp in enumerate(self.clsApp02_exp_list):
            if num == 0 or num == 4:
                ind_Exp_side_space = 0
            else:
                ind_Exp_side_space = self.clsApp02_Exp_SIDE_SPACE
            ind_Exp_line_space = self.clsApp02_Exp_LINE_SPACE * num

            self.ind_Exp_label = tk.Label(
                self.base_frame,
                text=ind_exp,
                font=self.clsApp02_Exp_FONT,
                justify=tk.LEFT,
            )
            self.ind_Exp_label.place(
                x=self.clsApp02_Exp_POS_X + ind_Exp_side_space,
                y=self.clsApp02_Exp_POS_Y + ind_Exp_line_space,
            )

        # 　３－ＢＧ　図形挿入（四角）
        canvas = tk.Canvas(self.base_frame, width=1100, height=200)
        canvas.create_rectangle(5, 5, 1095, 130)
        canvas.place(x=50, y=265)

        # 　３　読込フォルダ設定【FldRef】
        self.clsApp03_FldRef_POS_X = 100  # << 説明の横位置
        self.clsApp03_FldRef_POS_Y = 285  # << 説明の縦位置
        self.clsApp03_FldRef_SIDE_SPACE = (
            100  # << 参照フォルダ表示設の、説明の横位置からの左詰めスペース
        )
        self.clsApp03_FldRef_LINE_SPACE = 65  # << 説明文と参照ボタンの行間

        self.clsApp03_FldRef_entry_str = tk.StringVar(self, self.initial_dir)

        # 　３－１　ラベル１（説明文）
        self.clsApp03_FldRef_label01 = tk.Label(
            self.base_frame,
            text="１　BRSファイルを読み込むフォルダを指定してください。",
            font=("BIZ UDPゴシック", 14, "bold", "underline"),
        )
        self.clsApp03_FldRef_label01.place(
            x=self.clsApp03_FldRef_POS_X, y=self.clsApp03_FldRef_POS_Y
        )

        # 　３－２　参照ボタン
        self.clsApp03_FldRef_button = tk.Button(
            self.base_frame,
            text="参 照",
            font=("BIZ UDPゴシック", 12),
            relief="raised",
            width=5,
            bd=5,
            bg=self.button_color,
            command=lambda: self.button_click_FldRef(),
        )
        self.clsApp03_FldRef_button.place(
            x=self.clsApp03_FldRef_POS_X,
            y=self.clsApp03_FldRef_POS_Y + self.clsApp03_FldRef_LINE_SPACE,
            anchor=tk.W,
        )

        # 　３－３　ラベル２（読込フォルダ表示）
        self.clsApp03_FldRef_label02 = tk.Label(
            self.base_frame,
            textvariable=self.clsApp03_FldRef_entry_str,
            font=("BIZ UDPゴシック", 11),
            relief="sunken",
            width=80,
            bd=5,
            pady=3,
            bg="#D4E7F3",
            fg="#333333",
            anchor=tk.W,
        )
        self.clsApp03_FldRef_label02.place(
            x=self.clsApp03_FldRef_POS_X + self.clsApp03_FldRef_SIDE_SPACE,
            y=self.clsApp03_FldRef_POS_Y + self.clsApp03_FldRef_LINE_SPACE,
            anchor=tk.W,
        )

        # 　４　変換実行【RunConv】
        self.clsApp04_RunConv_POS_X = 50  # << 説明の横位置
        self.clsApp04_RunConv_POS_Y = 420  # << 説明の縦位置
        self.clsApp04_RunConv_SIDE_SPACE_for_Exp = (
            25  # << 変換実行ボタンの、説明の横位置からの左詰めスペース
        )
        self.clsApp04_RunConv_SIDE_SPACE_for_PitchData = (
            220  # << 変換実行ボタンの、説明の横位置からの左詰めスペース
        )
        self.clsApp04_RunConv_LINE_SPACE = 70  # << 行間
        self.clsApp04_RunConv_SIDE_SPACE_for_Run = (
            450  # << 変換実行ボタンの、説明の横位置からの左詰めスペース
        )

        # 　４－１　ラベル（説明文）
        self.clsApp04_RunConv_label01 = tk.Label(
            self.base_frame,
            text="２　変換したいピッチ数を入力して、実行ボタンを押してしてください。　（ピッチ変換が実行されます）",
            font=("BIZ UDPゴシック", 14, "bold", "underline"),
        )
        self.clsApp04_RunConv_label01.place(
            x=self.clsApp04_RunConv_POS_X, y=self.clsApp04_RunConv_POS_Y
        )

        # 　４－２　ピッチデータ取得
        # ------------------------------------------------------------------------------
        self.clsApp04_RunConv_explain_label = tk.Label(
            self.base_frame,
            text="ピッチ（sec）\n入力範囲：0.01～1.00",
            font=("BIZ UDPゴシック", 12, "bold"),
            justify=tk.LEFT,
        )
        self.clsApp04_RunConv_explain_label.place(
            x=self.clsApp04_RunConv_POS_X + self.clsApp04_RunConv_SIDE_SPACE_for_Exp,
            y=self.clsApp04_RunConv_POS_Y + self.clsApp04_RunConv_LINE_SPACE,
            anchor=tk.W,
        )
        # ------------------------------------------------------------------------------

        # ------------------------------------------------------------------------------
        self.last_valid_value = 0.10
        self.new_pitch_data = self.last_valid_value

        font_style = ("BIZ UDPゴシック", 16)
        self.clsApp04_RunConv_entry = tk.Entry(
            self.base_frame,
            bg="#EFF5F7",
            bd=1,
            width=8,
            relief="solid",
            font=font_style,
            justify=tk.CENTER,
        )
        self.clsApp04_RunConv_entry.place(
            x=self.clsApp04_RunConv_POS_X
            + self.clsApp04_RunConv_SIDE_SPACE_for_PitchData,
            y=self.clsApp04_RunConv_POS_Y + self.clsApp04_RunConv_LINE_SPACE,
            anchor=tk.W,
            height=40,
        )
        self.clsApp04_RunConv_entry.insert(0, "0.10")
        self.clsApp04_RunConv_entry.bind("<Button-1>", self.clear_entry)
        self.clsApp04_RunConv_entry.bind("<Return>", self.format_decimal)
        # self.clsApp04_RunConv_entry.bind("<FocusOut>", self.format_decimal)
        # ------------------------------------------------------------------------------

        # 　４－３　変換実行ボタン
        self.clsApp04_RunConv_button = tk.Button(
            self.base_frame,
            text="変 換 実 行",
            font=("BIZ UDPゴシック", 16),
            relief="raised",
            width=16,
            height=2,
            bd=5,
            bg=self.button_color,
            command=lambda: self.button_click_RunConv(),
        )
        self.clsApp04_RunConv_button.place(
            x=self.clsApp04_RunConv_POS_X + self.clsApp04_RunConv_SIDE_SPACE_for_Run,
            y=self.clsApp04_RunConv_POS_Y + self.clsApp04_RunConv_LINE_SPACE,
            anchor=tk.W,
        )

        # ５　終了ボタン【Exit】
        self.clsApp05_Exit_POS_X = self.BASE_WINDOW_WIDTH - 150  # 右下に配置（横位置）
        self.clsApp05_Exit_POS_Y = self.BASE_WINDOW_HEIGHT - 50  # 右下に配置（縦位置）

        self.clsApp05_Exit_button = tk.Button(
            self.base_frame,
            text="終 了",
            font=("BIZ UDPゴシック", 12),
            relief="raised",
            width=10,
            height=2,
            bd=5,
            bg="#E0E0E0",
            command=self.quit_app,
        )
        self.clsApp05_Exit_button.place(
            x=self.clsApp05_Exit_POS_X, y=self.clsApp05_Exit_POS_Y, anchor=tk.CENTER
        )

        # 　０　アプリを前面に
        self.lift()
        self.mainloop()

    # 　アプリ終了－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－▼▼▼▼
    def quit_app(self):
        self.destroy()

    # 　－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－▲▲▲▲

    # 　ボタンクリック（参照ボタン）－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－▼▼▼▼
    def button_click_FldRef(self):
        dfBC_log = filedialog.askdirectory(initialdir=self.initial_dir)
        # ※ ファイル選択をキャンセルした時にパスが非表示になるのを防止
        if dfBC_log:
            pass
        else:
            dfBC_log = self.clsApp03_FldRef_entry_str.get()
        self.clsApp03_FldRef_entry_str.set(dfBC_log)

        self.working_dir = self.clsApp03_FldRef_entry_str.get()

    # 　－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－▲▲▲▲

    # 　変換実行　－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－▼▼▼▼
    def button_click_RunConv(self):
        self.format_decimal()

        self.new_pitch_data = self.last_valid_value

        os.chdir(self.working_dir)

        self.file_list = [
            f for f in os.listdir() if ".lnk" not in f if ".brs" in f
        ]  # ファイル一覧（ショートカットファイルを除く）
        self.file_name_list = [
            f.split(".")[0] for f in self.file_list
        ]  # ファイル名（拡張子なし）一覧
        self.file_count = len(self.file_name_list)  # ファイル数

        # 　ファイルがない場合はエラー表示して終了　－－－－－－－－－－－－－－－－－－－
        if self.file_count == 0:
            self.show_error_window("指定フォルダにBRSファイルがありません。")
            return
        # 　－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－

        # 　エクセルファイル保存フォルダを作成　－－－－－－－－－－－－－－－－－－－－－
        # 　※ 保存フォルダがある場合は連番（(1)、(2)・・・）を作成
        self.save_folder_name = "ピッチ変換後ファイル保存"  # 　保存フォルダ名

        if not os.path.isdir(self.save_folder_name):
            os.makedirs(self.save_folder_name)
            self.save_folder_fullpath = os.path.join(
                self.working_dir, self.save_folder_name
            )  # 　保存フォルダフルパス
        # 　保存フォルダがある場合の連番作成
        else:
            self.std_save_folder_name = self.save_folder_name
            for num in range(len(os.listdir())):
                try:
                    self.save_folder_name = (
                        self.std_save_folder_name + "(" + str(num + 1) + ")"
                    )
                    os.makedirs(self.save_folder_name)
                    self.save_folder_fullpath = os.path.join(
                        self.working_dir, self.save_folder_name
                    )
                    break
                except:
                    pass
        # 　－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－

        # 　エクセルファイル保存フォルダを作成　－－－－－－－－－－－－－－－－－－－－－
        for each_file in self.file_list:
            try:
                with open(each_file, "r", encoding="UTF-8") as file_data:
                    read_file = file_data.readlines()
            except:
                try:
                    with open(each_file, "r", encoding="SHIFT-JIS") as file_data:
                        read_file = file_data.readlines()
                except:
                    # エンコーディング確認
                    with open(each_file, "rb") as f:
                        char = f.read()
                        result = chardet.detect(char)
                        encoding = result["encoding"]
                    with open(each_file, "r", encoding=encoding) as file_data:
                        read_file = file_data.readlines()

            self.save_file = os.path.join(
                self.save_folder_fullpath, "（ピッチ変換）" + each_file
            )  # 　保存フォルダフルパス

            self.original_pitch_data = float(read_file[60].strip())
            self.original_data_length = int(read_file[174].strip())

            self.pitch_convert_facor = int(
                self.new_pitch_data / self.original_pitch_data
            )
            self.new_data_length = int(
                self.original_data_length / self.pitch_convert_facor
            )

            read_file[60] = f"{self.new_pitch_data}\n"
            read_file[174] = f"{self.new_data_length}\n"

            self.new_data = read_file[175 :: self.pitch_convert_facor]

            self.new_file_data = read_file[:175] + self.new_data

            with open(self.save_file, mode="w", encoding="SHIFT-JIS") as new_file:
                new_file.writelines(self.new_file_data)

        self.show_completion_window(self.new_pitch_data, self.file_count)

    # 　－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－▲▲▲▲

    # 　ピッチデータ取得用　－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－▼▼▼▼
    def clear_entry(self, event=None):
        self.clsApp04_RunConv_entry.delete(0, tk.END)  # 中身を削除して空白に

    def format_decimal(self, event=None):
        try:
            new_val = float(self.clsApp04_RunConv_entry.get())
            if 0.01 <= new_val <= 1:
                self.clsApp04_RunConv_entry.delete(0, tk.END)
                self.clsApp04_RunConv_entry.insert(
                    0, f"{new_val:.2f}"
                )  # 小数点以下2桁に整形
                self.last_valid_value = new_val
                self.new_pitch_data = self.last_valid_value
            else:
                self.clsApp04_RunConv_entry.delete(0, tk.END)
                self.clsApp04_RunConv_entry.insert(0, f"{self.last_valid_value:.2f}")
        except ValueError:
            self.clsApp04_RunConv_entry.delete(0, tk.END)
            self.clsApp04_RunConv_entry.insert(0, f"{self.last_valid_value:.2f}")

    # 　－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－▲▲▲▲

    # 　ファイルがないときのエラー表示　－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－▼▼▼▼
    def show_error_window(self, message):
        win = tk.Toplevel(self)
        win.title("エラー")
        win.geometry("500x200+400+250")  # 終了画面と同じサイズ感
        win.configure(bg="#E3E9EC")

        label = tk.Label(
            win,
            text=message,
            font=("BIZ UDPゴシック", 14),
            bg="#E3E9EC",
            justify=tk.CENTER,
        )
        label.pack(pady=40)

        # OKボタン → エラー画面だけ閉じる
        ok_button = tk.Button(
            win,
            text="OK",
            font=("BIZ UDPゴシック", 12),
            width=10,
            height=2,
            command=win.destroy,
        )
        ok_button.pack(pady=10)

        # --- モーダル化（閉じるまでメイン画面操作不可） ---
        win.grab_set()  # このウィンドウに操作を限定
        win.transient(self)  # 親ウィンドウの上に表示
        win.focus_set()  # フォーカスを当てる
        win.wait_window(win)  # このウィンドウが閉じられるまで処理を待機

    # 　－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－▲▲▲▲

    # 　終了時の表示　－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－▼▼▼▼
    def show_completion_window(self, pitch, file_count):
        # 新しいトップレベルウィンドウ
        win = tk.Toplevel(self)
        win.title("処理完了")
        win.geometry("500x280+400+250")  # 少し大きめ
        win.configure(bg="#E3E9EC")

        # メッセージ
        msg = (
            f"処理が完了しました。\n\n"
            f"ピッチ：{pitch:.2f} sec\n"
            f"処理ファイル数：{file_count} 件\n\n"
            f"10秒後に自動で終了します。"
        )
        label = tk.Label(
            win, text=msg, font=("BIZ UDPゴシック", 14), bg="#E3E9EC", justify=tk.CENTER
        )
        label.pack(pady=40)

        # OKボタン（縦幅を少し大きく）
        ok_button = tk.Button(
            win,
            text="OK",
            font=("BIZ UDPゴシック", 12),
            width=10,
            height=2,  # ← ここで縦幅を大きく
            command=self.quit_app,
        )
        ok_button.pack(pady=10)

        # 10秒後に自動終了
        win.after(10000, self.quit_app)

    # 　－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－▲▲▲▲


if __name__ == "__main__":
    Application()
