import os
import tkinter as tk
from tkinter import filedialog, ttk
import chardet
import pandas as pd
from PIL import Image, ImageTk
import math
from io import BytesIO

os.environ['TCL_LIBRARY'] = r'C:\Users\User\AppData\Local\Programs\Python\Python313\tcl\tcl8.6'
os.environ['TK_LIBRARY'] = r'C:\Users\User\AppData\Local\Programs\Python\Python313\tcl\tk8.6'

#　ファイルをpandasで読込
def open_csv(f_name):
    # 文字コードを判別
    with open(f_name, 'rb') as f:
        char = f.read()
        result = chardet.detect(char)
        encoding = result["encoding"]
    # pandasで読込
    read_file = pd.read_csv(f_name, encoding=encoding, header=None)
    return read_file

#　パスのスラッシュ/バックスラッシュをスラッシュに統一
def path_check(text_data):
    try:
        if os.path.exists(text_data):
            pass
    except:
        text_data = os.getcwd()
    if "\n" or "/n" in text_data:
        text_data = text_data.rstrip("\n")
        text_data = text_data.rstrip("/n") 
    check01 = "\\"
    check02 = r"\\"
    check03 = "//"
    if check01 in text_data:
        text_data = text_data.replace(check01,"/")
    if check02 in text_data:
        text_data = text_data.replace(check02,"/")
    if check03 in text_data:
        text_data = text_data.replace(check03,"/")
    return text_data

#　アプリケーション起動ディレクトリ内にある［docsフォルダ］から［RefereceData.csv］を探して読込
for RDcsv_curDir, _, RDcsv_files in os.walk(os.getcwd()):
    if "_PyApplication" and "01_PhotoImageCompression" and "_docs" in RDcsv_curDir:
        if "RefereceData.csv" in RDcsv_files:
            global_init_state_fname_fullpath = os.path.join(RDcsv_curDir,"RefereceData.csv")
            global_init_state_file = open_csv(global_init_state_fname_fullpath)

global_current_path = os.getcwd()

#　グローバル変数
#　global_current_path               ／ 現在の絶対パス 〔Str〕
#　global_init_state_fname_fullpath  ／ ［RefereceData.csv］ファイルの絶対パス 〔Str〕
#　global_init_state_file            ／ ［RefereceData.csv］ファイルに記録されている初期状態データの表 〔pandas.DataFrame〕
#　（global_init_state_file[1][0]  ／ 初期状態１　読込フォルダの絶対パス 〔Str〕）
#　（global_init_state_file[1][1]  ／ 初期状態２　保存フォルダの絶対パス 〔Str〕）
#　（global_init_state_file[1][2]  ／ 初期状態３　ファイル圧縮率 〔str〕）

#　▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼　クラス（１）アプリケーション【clsApp】　▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
class Application(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("App_PhotoCompress_ver-1.0")

        #　画面サイズ設定
        self.BASE_WINDOW_WIDTH = 1200    # << アプリサイズ（横幅）
        self.BASE_WINDOW_HEIGHT = 750    # << アプリサイズ（縦幅）
        self.BASE_WINDOW_POS_X = 25   # << アプリの左上位置（液晶画面の左端からの距離）
        self.BASE_WINDOW_POS_Y = 25    # << アプリの左上位置（液晶画面の上側からの距離）
        self.geometry(f"{self.BASE_WINDOW_WIDTH}x{self.BASE_WINDOW_HEIGHT}+{self.BASE_WINDOW_POS_X}+{self.BASE_WINDOW_POS_Y}")
                
        #　１　タイトル設置【Title】
        clsApp_Title_LABEL01 = tk.Label(self,
                                        text = "画像ファイル　サイズ一括圧縮変換ソフト",
                                        font =  ("BIZ UDPゴシック", 32),
                                        height = 2,    # << タイトル高さ
                                        width = 34,    # << タイトル幅
                                        bg = "#D1D9EE")    # << タイトル背景色
        clsApp_Title_LABEL01.place(relx = 0.5,    # << タイトル設置の横位置（中間地点）
                                   y = 20,    # << タイトル設置の縦位置
                                   anchor = tk.N)

        #　２　説明文設置【Exp】
        clsApp_Exp_POS_X = 100    # << 説明文設置の横位置
        clsApp_Exp_POS_Y = 125    # << 説明文設置の縦位置
        clsApp_Exp_SIDE_SPACE = 25    # << 説明文設置の左詰めスペース
        clsApp_Exp_LINE_SPACE = 30    # << 説明文設置の行間
        clsApp_Exp_FONT = ("BIZ UDPゴシック", 14)

        clsApp_Exp_LABEL01 = tk.Label(self, 
                                      text = "○ 概要",
                                      font = clsApp_Exp_FONT,
                                      justify = tk.LEFT)
        clsApp_Exp_LABEL02 = tk.Label(self,
                                      text = "画像ファイルを、指定した圧縮率で、一括圧縮できます。　変換前の画像は、変換しても残ります。",
                                      font = clsApp_Exp_FONT,
                                      justify=tk.LEFT)
        clsApp_Exp_LABEL03 = tk.Label(self,
                                      text = "複数のフォルダや、サブフォルダに画像ファイルが入っていても変換可能です。",
                                      font = clsApp_Exp_FONT,
                                      justify=tk.LEFT)
        clsApp_Exp_LABEL01.place(x=clsApp_Exp_POS_X,
                                 y=clsApp_Exp_POS_Y)
        clsApp_Exp_LABEL02.place(x=clsApp_Exp_POS_X+clsApp_Exp_SIDE_SPACE,
                                 y=clsApp_Exp_POS_Y+clsApp_Exp_LINE_SPACE)
        clsApp_Exp_LABEL03.place(x=clsApp_Exp_POS_X+clsApp_Exp_SIDE_SPACE,
                                 y=clsApp_Exp_POS_Y+clsApp_Exp_LINE_SPACE*2)

        #　３　読込先フォルダの指定【RdFld：⇒クラス（２）】
        clsApp_RdFld_POS_X = 100
        clsApp_RdFld_POS_Y = 240
        clsApp_RdFld_TEXT = "≪ １ ≫　画像を読み込むフォルダを指定してください。"
       
        clsApp_set_RdFld = Set_Folder_Reference_Frame(self,
                                                      clsApp_RdFld_POS_X,
                                                      clsApp_RdFld_POS_Y,
                                                      global_init_state_file[1][0])    # << 初期状態データの１列１行目（読込フォルダパス）の読込
        clsApp_set_RdFld.label_01(clsApp_RdFld_TEXT)
        clsApp_set_RdFld.button_01(0)

        #　４　保存先フォルダの指定【SvFld：⇒クラス（２）】
        clsApp_SvFld_POS_X = 100
        clsApp_SvFld_POS_Y = 350
        clsApp_SvFld_TEXT = "≪ ２ ≫　画像を保存するフォルダを指定してください。"

        clsApp_set_SvFld = Set_Folder_Reference_Frame(self,
                                                      clsApp_SvFld_POS_X,
                                                      clsApp_SvFld_POS_Y,
                                                      global_init_state_file[1][1])    # << 初期状態データの１列２行目（保存フォルダパス）の読込
        clsApp_set_SvFld.label_01(clsApp_SvFld_TEXT)
        clsApp_set_SvFld.button_01(1)

        #　５　圧縮率フォルダの指定【CompRate：⇒クラス（３）】
        clsApp_CompRate_POS_X = 100
        clsApp_CompRate_POS_Y = 460
        clsApp_CompRate_TEXT01 = "≪ ３ ≫　圧縮率を指定してください。"
        clsApp_CompRate_TEXT02 = "　　　　　　圧縮率 "
        clsApp_CompRate_TEXT03 = "％　"
        clsApp_CompRate_TEXT04 = "（1～100％の範囲で指定してください。）"
        
        clsApp_set_CompRate = Set_Compression_Rate_Frame(self,
                                                         clsApp_CompRate_POS_X,
                                                         clsApp_CompRate_POS_Y,
                                                         global_init_state_file[1][2])
        clsApp_set_CompRate.label_01(clsApp_CompRate_TEXT01)
        clsApp_set_CompRate.label_02(clsApp_CompRate_TEXT02)
        clsApp_set_CompRate.label_03(clsApp_CompRate_TEXT03)
        clsApp_set_CompRate.label_04(clsApp_CompRate_TEXT04)
        
        #　６　ファイル変換実行（説明文）【ConvExp：⇒クラス（４）】
        clsApp_ConvExp_POS_X = 100
        clsApp_ConvExp_POS_Y = 570
        clsApp_ConvExp_TEXT01 = "≪ ４ ≫　変換を実行してください。"
        clsApp_ConvExp_TEXT02 = "　　　　　　　　［変換する前に一覧を確認］"
        clsApp_ConvExp_TEXT03 = "　　　　　　　　［一覧を確認せずに変換を実行］"
        
        clsApp_set_ConvExp = Set_Convert_Explain_Frame(self, clsApp_ConvExp_POS_X, clsApp_ConvExp_POS_Y)
        clsApp_set_ConvExp.label_01(clsApp_ConvExp_TEXT01)
        clsApp_set_ConvExp.label_02(clsApp_ConvExp_TEXT02)
        clsApp_set_ConvExp.label_03(clsApp_ConvExp_TEXT03)
        
        #　７　ファイル変換実行（変換前プレビューボタン）【clsPrevBut：⇒クラス（５）】
        run_convert_button_01 = Set_Preview_Button_Frame(self, 225, 655)

        #　８　ファイル変換実行【clsRunConv：⇒クラス（７）】
        run_convert_button_02 = Run_convert(self, 510, 655)

        #　９　終了ボタン
        e_style001 = ttk.Style()
        e_style001.configure("St00.TButton",
                           font = ("BIZ UDPゴシック",12,"bold"),
                           padding = [0, 10])            
        e_button001 = ttk.Button(self,
                               text = "終了",
                               default = "active",
                               style = "St00.TButton",
                               width = 6,
                               command = lambda:self.destroy()) 
        e_button001.place(x=1000, y=675)
                        
        #　アプリの基本色
        self.tk_setPalette(background="#F0F0F0")
        
        self.lift()

        # アプリケーションの実行
        self.mainloop()
#　▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲　クラス（１）アプリケーション【clsApp】　▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲


#　▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼　クラス（２）＞ ３，４：フォルダ参照用【clsFldRef】　▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
class Set_Folder_Reference_Frame(tk.Frame):
    def __init__(self,
                 clsFldRef_root,
                 clsFldRef_pos_x,
                 clsFldRef_pos_y,
                 clsFldRef_init_state_value):
        super().__init__(clsFldRef_root)
 
        self.clsFldRef_entry_str = tk.StringVar(self, clsFldRef_init_state_value)
        if os.path.isdir(self.clsFldRef_entry_str.get()):
            pass
        else:
            self.clsFldRef_entry_str.set(os.getcwd())

        self.label_01()
        self.button_01()
        self.label_02()

        self.place(x=clsFldRef_pos_x, y=clsFldRef_pos_y)

    #　ラベル（説明文）
    def label_01(self,
                 dfLb01_txt = ""):
        label001 = tk.Label(self,
                            text = dfLb01_txt,
                            font = ("BIZ UDPゴシック", 16),
                            fg = "blue")
        label001.grid(row=0, column=0, columnspan=2, sticky=tk.W)

    #　参照ボタン
    def button_01(self,
                  dfBt01_num = 0):
        button001 = tk.Button(self,
                              text = "参照",
                              font = ("BIZ UDPゴシック", 14),
                              relief = "raised",
                              width = 5,
                              bd = 5,
                              bg = "#E0E0E0",
                              command = lambda:self.button_click(dfBt01_num))
        button001.grid(row=1, column=0, sticky=tk.W, padx=10, pady=10)

    #　ボタンクリック時動作【⇒参照ボタン用】
    #　→　読込/保存フォルダを選択して絶対パスを取得
    def button_click(self, dfBC_num):
        dfBC_initialdir = path_check(global_init_state_file[1][dfBC_num])
        dfBC_log = filedialog.askdirectory(initialdir = dfBC_initialdir)
        # ※ ファイル選択をキャンセルした時にパスが非表示になるのを防止
        if dfBC_log:    
            pass
        else:
            dfBC_log = dfBC_initialdir
        self.clsFldRef_entry_str.set(dfBC_log)
        self.label_02()
        #　初期状態１/２（読込/保存フォルダの絶対パス）＜global_init_state_file[1][0]/[1][1]＞に数値を代入、［RefereceData.csv］ファイルに保存
        global_init_state_file.loc[dfBC_num, 1] = dfBC_log
        global_init_state_file.to_csv(global_init_state_fname_fullpath, encoding="utf-8", index=False, header=False)
        
    #　フォルダパス表示
    def label_02(self):
        labe002 = tk.Label(self,
                           text = self.clsFldRef_entry_str.get(),
                           font = ("BIZ UDPゴシック", 12),
                           relief = "sunken",
                           width = 80,
                           bd = 5,
                           fg = "#333333",
                           anchor=tk.W)
        labe002.grid(row=1, column=1, sticky=tk.W, padx=10)
#　▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲　クラス（２）＞ ３，４：フォルダ参照用【clsFldRef】　▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲


#　▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼　クラス（３）＞ ５：圧縮率指定用【clsComRat】　▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
class Set_Compression_Rate_Frame(tk.Frame):
    def __init__(self,
                 clsComRat_root,
                 clsComRat_pos_x,
                 clsComRat_pos_y,
                 clsComRat_init_state_value):
        super().__init__(clsComRat_root)

        self.clsComRat_entry_int = tk.IntVar(self, clsComRat_init_state_value)
        self.clsComRat_previous_num = self.clsComRat_entry_int.get()
        
        self.label_01()
        self.label_02()
        self.spinbox_01()
        self.label_03()
        self.label_04()
        
        self.place(x=clsComRat_pos_x, y=clsComRat_pos_y)

    #　ラベル（説明文）
    def label_01(self,
                 dfLb01_txt = ""):
        label001 = tk.Label(self,
                            text = dfLb01_txt,
                            font = ("BIZ UDPゴシック", 16),
                            fg = "blue")
        label001.grid(row=0, column=0, columnspan=4, sticky=tk.W)

    #　ラベル
    def label_02(self,
                 dfLb02_txt = ""):
        label002 = tk.Label(self,
                            text = dfLb02_txt,
                            font = ("BIZ UDPゴシック", 16, "bold"),
                            pady = 15)
        label002.grid(row=1, column=0, sticky=tk.W)

    #　圧縮率表示スピンボックス
    def spinbox_01(self):
        spinbox001 = tk.Spinbox(self,
                                textvariable = self.clsComRat_entry_int,
                                font = ("BIZ UDPゴシック", 16, "bold"),
                                justify = tk.RIGHT, 
                                width = 3,
                                bg = "#FFFFEE",
                                from_= 1,    # << 圧縮率の最小値
                                to = 100,    # << 圧縮率の最大値 
                                increment = 5,    # << 数値を５の倍数で上下 
                                command = self.spinbox_increment_adjust)
        spinbox001.bind("<Return>", self.spinbox_bind_enter)
        spinbox001.bind("<KeyRelease>", self.spinbox_bind_input_num)
        spinbox001.grid(row=1, column=1, sticky=tk.W)
    
    #　「Enter」キークリック時の動作【バインド：⇒圧縮率表示スピンボックス用】
    # 　→　初期状態３（ファイル圧縮率）＜global_init_state_file[1][2]＞に数値を代入、［RefereceData.csv］ファイルに保存
    def spinbox_bind_enter(self, event):
        global_init_state_file.loc[2,1] = self.clsComRat_entry_int.get()
        global_init_state_file.to_csv(global_init_state_fname_fullpath, encoding="utf-8", index=False, header=False)

    #　キー入力操作時の動作（※キーを離したタイミング）【バインド：⇒圧縮率表示スピンボックス用】
    #　→　数値が０以下になったときに１、１００を超えたときに１００にする
    #　→　初期状態３（ファイル圧縮率）＜global_init_state_file[1][2]＞に数値を代入
    def spinbox_bind_input_num(self, event):
        #　数値以外を入力したとき、元の数値に戻る
        try:    
            self.clsComRat_entry_int.get()
        except:
            self.clsComRat_entry_int.set(self.clsComRat_previous_num)
        #　数値が０以下になったときに１、１００を超えたときに１００にする
        if self.clsComRat_entry_int.get() > 100:
            self.clsComRat_entry_int.set(100)
        elif self.clsComRat_entry_int.get()<=0:
            self.clsComRat_entry_int.set(1)
        self.clsComRat_previous_num = self.clsComRat_entry_int.get()
        global_init_state_file.loc[2,1] = self.clsComRat_entry_int.get()

    #　インクリメント調整【コマンド：⇒圧縮率表示スピンボックス用】
    #　→　５の倍数以外でも５の倍数ごとに移動
    #　→　初期状態３（ファイル圧縮率）＜global_init_state_file[1][2]＞に数値を代入       
    def spinbox_increment_adjust(self):
        self.clsComRat_dfBCincrement_input_num = self.clsComRat_entry_int.get()
        if self.clsComRat_dfBCincrement_input_num%5==0:    #　５で割り切れるときはそのまま
            self.clsComRat_entry_int.set(self.clsComRat_dfBCincrement_input_num)
        elif self.clsComRat_dfBCincrement_input_num<=1:    #　１以下になったときは１にする 
            self.clsComRat_entry_int.set(1)
        elif (self.clsComRat_dfBCincrement_input_num-self.clsComRat_previous_num)>=0:    #　数値が増加したとき５の倍数で切り捨て
            self.clsComRat_entry_int.set(self.clsComRat_dfBCincrement_input_num//5*5)
        elif (self.clsComRat_dfBCincrement_input_num-self.clsComRat_previous_num)<0:    #　数値が減少したとき５の倍数で切り捨てた後プラス５
            self.clsComRat_entry_int.set(self.clsComRat_dfBCincrement_input_num//5*5+5)  
        self.clsComRat_previous_num = self.clsComRat_entry_int.get()
        global_init_state_file.loc[2,1] = self.clsComRat_entry_int.get()
    
    #　ラベル
    def label_03(self,
                 dfLb03_txt = ""):
        label003 = tk.Label(self,
                            text = dfLb03_txt,
                            font = ("BIZ UDPゴシック", 16, "bold"))
        label003.grid(row=1, column=2, sticky=tk.W)

    #　ラベル
    def label_04(self,
                 dfLb04_txt = ""):
        label004 = tk.Label(self,
                            text = dfLb04_txt,
                            font = ("BIZ UDPゴシック", 12))
        label004.grid(row=1, column=3, sticky=tk.W)
#　▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲　クラス（３）＞ ５：圧縮率指定用【clsComRat】　▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲


#　▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼　クラス（４）＞ ６　ファイル変換実行（説明文）【clsConvExp】　▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
class Set_Convert_Explain_Frame(tk.Frame):
    def __init__(self,
                 clsConvExp_root,
                 clsConvExp_pos_x,
                 clsConvExp_pos_y):
        super().__init__(clsConvExp_root)

        self.label_01()
        self.label_02()
        self.label_03()
        
        self.place(x=clsConvExp_pos_x, y=clsConvExp_pos_y)

    #　ラベル（説明文）
    def label_01(self,
                 dfLb01_txt = ""):
        label001 = tk.Label(self,
                            text = dfLb01_txt,
                            font = ("BIZ UDPゴシック", 16),
                            fg = "blue")
        label001.grid(row=0, column=0, columnspan=2, sticky=tk.W)
        
    #　ラベル（説明文）
    def label_02(self,
                 dfLb02_txt = ""):
        label002 = tk.Label(self,
                            text = dfLb02_txt,
                            font = ("BIZ UDPゴシック", 12, "bold"),
                            pady = 20)
        label002.grid(row=1, column=0, sticky=tk.W)

    #　ラベル（説明文）
    def label_03(self,
                 dfLb03_txt = ""):
        label003 = tk.Label(self,
                            text = dfLb03_txt,
                            font = ("BIZ UDPゴシック", 12, "bold"))
        label003.grid(row=1, column=1, sticky=tk.W)
#　▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲　クラス（４）＞ ６　ファイル変換実行（説明文）【clsConvExp】　▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲


#　▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼　クラス（５）＞ ７－１　ファイル変換実行（一覧ボタン）【clsPrevBut】　▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
class Set_Preview_Button_Frame(tk.Frame):
    def __init__(self,
                 clsPrevBut_root,
                 clsPrevBut_pos_x,
                 clsPrevBut_pos_y):
        super().__init__(clsPrevBut_root)
        
        style001 = ttk.Style()
        style001.configure("St01.TButton",
                           font = ("BIZ UDPゴシック",12,"bold"),
                           padding = [0, 20],
                           foreground = "#666666")            
        button001 = ttk.Button(self,
                               text = "プレビュー",
                               default = "active",
                               style = "St01.TButton",
                               width = 9,
                               command = self.preview_button) 
        button001.pack()
        
        self.place(x=clsPrevBut_pos_x, y=clsPrevBut_pos_y)
            
    def capacity_check(self, dfCC_folder_path):
        dfCC_capacity_total = 0
        for dfCC_curDir, _, dfCC_files in os.walk(dfCC_folder_path):
            for dfCC_ind_file in dfCC_files:
                dfCC_fname_fullpath = os.path.join(dfCC_curDir, dfCC_ind_file)
                check_file_capacity = os.path.getsize(dfCC_fname_fullpath)/(1024*1024)
                dfCC_capacity_total += check_file_capacity
        return dfCC_capacity_total
    
    def preview_button(self):
        if self.capacity_check(global_init_state_file[1][0]) > 1000:
            
            dfPB_sub_window01 = tk.Toplevel(self)
            dfPB_sub_window01.title("一覧確認用")
            dfPB_sub_window01.geometry("1000x300")
                        
            label001 = tk.Label(dfPB_sub_window01,
                                font = ("BIZ UDPゴシック", 18, "bold"),
                                fg = "red",
                                text = "\n\n読込フォルダ内にある、画像以外も含めた総ファイルサイズが大きすぎます。")
            label002 = tk.Label(dfPB_sub_window01,
                                font = ("BIZ UDPゴシック", 18, "bold"),
                                fg = "red",
                                text = "読込フォルダを変更するか、ファイルを減らしてください。")
            label003 = tk.Label(dfPB_sub_window01,
                                font = ("BIZ UDPゴシック", 16, "bold"),
                                text = f"\n総ファイルサイズ：{self.capacity_check(global_init_state_file[1][0]):.2f}MB　")
            label004 = tk.Label(dfPB_sub_window01,
                                font=("BIZ UDPゴシック", 14), 
                                text = "\n（※総ファイルサイズを1000MB以下にしてください。）")
            
            label001.grid(row=0, column=0 ,columnspan=2, sticky="nsew")
            label002.grid(row=1, column=0 ,columnspan=2, sticky="nsew", pady=10)
            label003.grid(row=2, column=0, sticky=tk.E, pady=20)             
            label004.grid(row=2, column=1, sticky=tk.W)
            
            dfPB_sub_window01.grid_columnconfigure(0, weight=1)
            dfPB_sub_window01.grid_columnconfigure(1, weight=1)
        else:
            check_count = 0
            for dfPB01_curDir, _, dfPB01_files in os.walk(global_init_state_file[1][0]):
                for dfPB01_ind_fname in dfPB01_files:
                    try:
                        temp = Image.open(os.path.join(dfPB01_curDir, dfPB01_ind_fname))
                        check_count = 1
                        break
                    except:
                        pass
            
            if check_count == 1:
                global_init_state_file.to_csv(global_init_state_fname_fullpath, encoding="utf-8", index=False, header=False)
                Perview_Sub_Window(self)
            else:
                dfPB_sub_window02 = tk.Toplevel(self)
                dfPB_sub_window02.title("一覧確認用")
                dfPB_sub_window02.geometry("1000x300")
                            
                label101 = tk.Label(dfPB_sub_window02,
                                    font = ("BIZ UDPゴシック", 18, "bold"),
                                    fg = "red",
                                    text = "\n\n")
                label102 = tk.Label(dfPB_sub_window02,
                                    font = ("BIZ UDPゴシック", 18, "bold"),
                                    fg = "red",
                                    text = "読込フォルダに画像がありません。")
                label103 = tk.Label(dfPB_sub_window02,
                                    font = ("BIZ UDPゴシック", 16, "bold"),
                                    text = f"\n")
                label104 = tk.Label(dfPB_sub_window02,
                                    font=("BIZ UDPゴシック", 14), 
                                    text = "\n（※画像のある読込フォルダに設定してください。）")
                
                label101.grid(row=0, column=0 ,columnspan=2, sticky="nsew")
                label102.grid(row=1, column=0 ,columnspan=2, sticky="nsew", pady=10)
                label103.grid(row=2, column=0, sticky=tk.E, pady=20)             
                label104.grid(row=2, column=1, sticky=tk.W)

                dfPB_sub_window02.grid_columnconfigure(0, weight=1)
                dfPB_sub_window02.grid_columnconfigure(1, weight=1)

#　▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲　クラス（５）＞ ７－１　ファイル変換実行（一覧ボタン）【clsPrevBut】　▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲


#　▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼　クラス（６）＞ ７－２　ファイル変換実行（プレビューサブウインドウ）【clsPreSub】　▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
class Perview_Sub_Window(tk.Toplevel):
    def __init__(self, clsPreSub_root):
        super().__init__(clsPreSub_root)    
        
        self.geometry("1250x740+50+30")
        self.title("一覧確認用")
        
        self.make_figlist()

        self.clsPreSub_entry_int = tk.IntVar(self, 1)    #　ファイル番号
        self.clsPreSub_log = self.clsPreSub_entry_int.get()
        self.clsPreSub_entry_str = tk.StringVar(self, self.clsPreSub_fig_list[self.clsPreSub_log-1][1])    #　ファイル名（フィグリスト２列目）

        self.clsPreSub_entry_comp_int = tk.IntVar(self, global_init_state_file[1][2])
        self.clsPreSub_previous_comp_num = self.clsPreSub_entry_comp_int.get()
                
        #　ファイル選択（矢印、ファイル番号表示、ファイル名表示）
        self.frame_01(self, 50, 10)
        self.frame_02(self, 50, 60)
        self.frame_03(self, 50, 100)     
        
        #　ファイル一覧表示       
        self.frame102 = tk.Frame(self)
        self.frame102.place(x=50, y=190)        
        self.table_01(self.frame102)
        
        self.frame103 = tk.Frame(self)
        self.frame103.place(x=800, y=10)
        self.image_01(self.frame103, self.clsPreSub_fig_list[0][4])
        
        self.frame104 = tk.Frame(self)
        self.frame104.place(x=400, y=100) 
        self.spinbox_01(self.frame104)
        
        clsPreSub_label001 = tk.Label(self,
                                      text = "【このまま変換を実行】",
                                      font = ("BIZ UDPゴシック", 12, "bold")) 
        clsPreSub_label001.place(x=250, y=600, anchor=tk.N)   

        self.frame105 = tk.Frame(self)
        self.frame105.place(x=250, y=640, anchor=tk.N) 
        self.convert_button_in_preview(self.frame105) 
        
        e_style002 = ttk.Style()
        e_style002.configure("St00.TButton",
                           font = ("BIZ UDPゴシック",12,"bold"),
                           padding = [0, 10])            
        e_button001 = ttk.Button(self,
                               text = "メイン画面に戻る",
                               default = "active",
                               style = "St00.TButton",
                               width = 12,
                               command = lambda:self.destroy()) 
        e_button001.place(x=500, y=650)

    def make_figlist(self):
        self.clsPreSub_fig_list = []
        clsPreSub_fig_num = 1
        
        for clsPreSub_curDir, _, clsPreSub_files in os.walk(global_init_state_file[1][0]):
            for clsPreSub_ind_fname in clsPreSub_files:
                try:            
                    clsPreSub_ind_read_fullpath = path_check(os.path.join(clsPreSub_curDir, clsPreSub_ind_fname))
                    clsPreSub_ind_file_in_folder = clsPreSub_curDir.replace(global_init_state_file[1][0],"")[1:]
                    clsPreSub_ind_save_fullpath = os.path.join(global_init_state_file[1][1], clsPreSub_ind_file_in_folder, clsPreSub_ind_fname)
                    clsPreSub_buffer = BytesIO()
        
                    clsPreSub_ind_pre_img = Image.open(clsPreSub_ind_read_fullpath)

                    #　ファイル名を「ファイル名、拡張子（ドットを含む）」とに分割し、ドットを除いた拡張子を取得
                    clsPreSub_file_ind_file_extension = os.path.splitext(clsPreSub_ind_fname)[1][1:]
                    if clsPreSub_file_ind_file_extension == "jpg" or clsPreSub_file_ind_file_extension == "JPG":
                        clsPreSub_file_ind_file_extension = "jpeg"
                    if clsPreSub_file_ind_file_extension == "tif" or clsPreSub_file_ind_file_extension == "TIF":
                        clsPreSub_file_ind_file_extension = "tiff"
                        
                    clsPreSub_ind_resize_img = clsPreSub_ind_pre_img.resize((int(clsPreSub_ind_pre_img.width * int(global_init_state_file[1][2])/100), 
                                                         int(clsPreSub_ind_pre_img.height * int(global_init_state_file[1][2])/100)))
        
                    clsPreSub_ind_resize_img.save(clsPreSub_buffer, format=clsPreSub_file_ind_file_extension) 
                    clsPreSub_ind_pre_capacity = math.ceil(os.path.getsize(clsPreSub_ind_read_fullpath)/1024)
                    clsPreSub_ind_resize_capacity = math.ceil(clsPreSub_buffer.tell()/1024)
                    self.clsPreSub_fig_list.append([clsPreSub_fig_num,                #　（リスト１列目）ファイル番号
                                                    clsPreSub_ind_fname,              #　（リスト２列目）ファイル名
                                                    clsPreSub_ind_pre_capacity,       #　（リスト３列目）読込ファイルサイズ
                                                    clsPreSub_ind_resize_capacity,    #　（リスト４列目）保存ファイルサイズ
                                                    clsPreSub_ind_read_fullpath,      #　（リスト５列目）読込ファイルのフルパス
                                                    clsPreSub_ind_file_in_folder,     #　（リスト６列目）読込ファイル内のフォルダ（読込フォルダ直下ではなく、フォルダ内に画像がある場合）
                                                    clsPreSub_ind_save_fullpath])     #　（リスト７列目）保存ファイルのフルパス
                    clsPreSub_fig_num += 1
                except:
                    pass
        self.clsPreSub_max_num = clsPreSub_fig_num - 1
            
    #　タイトル   
    def frame_01(self, dfFr01_root, dfFr01_pos_x , dfFr01_pos_y):
        frame001 = tk.Frame(dfFr01_root)
        label001 = tk.Label(frame001,
                            text = "個別画像ファイルの確認",
                            font = ("BIZ UDPゴシック", 16, "bold", "underline"))
        label001.pack() 
        frame001.place(x=dfFr01_pos_x, y=dfFr01_pos_y)
    
    #　矢印ボタン       
    def frame_02(self, dfFr02_root, dfFr02_pos_x , dfFr02_pos_y):
        dfFr02_background = "#E0E0E0"
        frame002 = tk.Frame(dfFr02_root)
        button201 = tk.Button(frame002,
                              text = "◀◀",
                              bg = dfFr02_background,
                              command = self.button_click_01)
        button202 = tk.Button(frame002,
                              text = " ◀ ",
                              bg = dfFr02_background,
                              command = self.button_click_02)
        button203 = tk.Button(frame002,
                              text = " ▶ ",
                              bg = dfFr02_background,
                              command = self.button_click_03)
        button204 = tk.Button(frame002,
                              text = "▶▶",
                              bg = dfFr02_background,
                              command = self.button_click_04)
        button201.grid(row=0, column=0) 
        button202.grid(row=0, column=1)        
        button203.grid(row=0, column=2) 
        button204.grid(row=0, column=3) 
        frame002.place(x=dfFr02_pos_x, y=dfFr02_pos_y)
    
    def button_click_01(self):
        self.clsPreSub_entry_int.set(self.clsPreSub_entry_int.get()-5)
        if self.clsPreSub_entry_int.get() < 1:
            self.clsPreSub_entry_int.set(1)
        self.clsPreSub_log = self.clsPreSub_entry_int.get()
        self.clsPreSub_entry_str.set(self.clsPreSub_fig_list[self.clsPreSub_log-1][1])
        self.dfIm01_label001.pack_forget()
        self.dfIm01_label002.pack_forget()
        self.dfIm01_label003.pack_forget()
        self.dfIm01_label004.pack_forget()
        self.dfIm01_label005.pack_forget()
        self.image_01(self.frame103, self.clsPreSub_fig_list[self.clsPreSub_log-1][4])
    def button_click_02(self):
        self.clsPreSub_entry_int.set(self.clsPreSub_entry_int.get()-1)
        if self.clsPreSub_entry_int.get() < 1:
            self.clsPreSub_entry_int.set(1)
        self.clsPreSub_log = self.clsPreSub_entry_int.get()
        self.clsPreSub_entry_str.set(self.clsPreSub_fig_list[self.clsPreSub_log-1][1])
        self.dfIm01_label001.pack_forget()
        self.dfIm01_label002.pack_forget()
        self.dfIm01_label003.pack_forget()
        self.dfIm01_label004.pack_forget()
        self.dfIm01_label005.pack_forget()
        self.image_01(self.frame103, self.clsPreSub_fig_list[self.clsPreSub_log-1][4])
    def button_click_03(self):
        self.clsPreSub_entry_int.set(self.clsPreSub_entry_int.get()+1)
        if self.clsPreSub_entry_int.get() > self.clsPreSub_max_num:
            self.clsPreSub_entry_int.set(self.clsPreSub_max_num)
        self.clsPreSub_log = self.clsPreSub_entry_int.get()
        self.clsPreSub_entry_str.set(self.clsPreSub_fig_list[self.clsPreSub_log-1][1]) 
        self.dfIm01_label001.pack_forget()
        self.dfIm01_label002.pack_forget()
        self.dfIm01_label003.pack_forget()
        self.dfIm01_label004.pack_forget()
        self.dfIm01_label005.pack_forget()
        self.image_01(self.frame103, self.clsPreSub_fig_list[self.clsPreSub_log-1][4])
    def button_click_04(self):
        self.clsPreSub_entry_int.set(self.clsPreSub_entry_int.get()+5)
        if self.clsPreSub_entry_int.get() > self.clsPreSub_max_num:
            self.clsPreSub_entry_int.set(self.clsPreSub_max_num)
        self.clsPreSub_log = self.clsPreSub_entry_int.get()
        self.clsPreSub_entry_str.set(self.clsPreSub_fig_list[self.clsPreSub_log-1][1])
        self.dfIm01_label001.pack_forget()
        self.dfIm01_label002.pack_forget()
        self.dfIm01_label003.pack_forget()
        self.dfIm01_label004.pack_forget()
        self.dfIm01_label005.pack_forget()
        self.image_01(self.frame103, self.clsPreSub_fig_list[self.clsPreSub_log-1][4])
                           
    #　ファイル番号、ファイル名（Entry）表示
    def frame_03(self, dfFr03_root, dfFr03_pos_x, dfFr03_pos_y):
        frame003 = tk.Frame(dfFr03_root)
        label301 = tk.Label(frame003,
                            text = "ファイルNo.",
                            font = ("BIZ UDPゴシック", 14))
        entry302 = tk.Entry(frame003,
                            textvariable = self.clsPreSub_entry_int,
                            font = ("BIZ UDPゴシック", 16, "bold"), 
                            width = 3,
                            bg = "#FFFFEE",
                            justify = tk.RIGHT)
        label303 = tk.Label(frame003,
                            text = "/" + str(self.clsPreSub_max_num),
                            font = ("BIZ UDPゴシック", 16, "bold"))
        label314 = tk.Label(frame003,
                            text = "ファイル名",
                            font = ("BIZ UDPゴシック", 14),
                            anchor = tk.W)
        label315 = tk.Label(frame003,
                            textvariable = self.clsPreSub_entry_str,
                            font = ("BIZ UDPゴシック", 16, "bold"),
                            anchor = tk.W)
        label301.grid(row=0, column=0)
        entry302.grid(row=0, column=1)
        label303.grid(row=0, column=2)
        label314.grid(row=1, column=0, sticky=tk.W)
        label315.grid(row=1, column=1, columnspan=2, sticky=tk.W, pady=10)                 
        frame003.place(x=dfFr03_pos_x, y=dfFr03_pos_y)
        
        entry302.bind("<KeyRelease>", self.entry_bind_input_num)
    
    #　ファイル番号を入力したとき（キーを離したタイミング）の操作
    def entry_bind_input_num(self, event):
        try:    
            self.clsPreSub_entry_int.get()
        except:
            self.clsPreSub_entry_int.set(self.clsPreSub_log)
        #　数値が０以下になったときに１、ファイル最大数を超えたときにファイル最大数にする
        if self.clsPreSub_entry_int.get() > self.clsPreSub_max_num:
            self.clsPreSub_entry_int.set(self.clsPreSub_max_num)
        elif self.clsPreSub_entry_int.get()<=0:
            self.clsPreSub_entry_int.set(1)
        self.clsPreSub_log = self.clsPreSub_entry_int.get()
        self.clsPreSub_entry_str.set(self.clsPreSub_fig_list[self.clsPreSub_log-1][1]) 
        self.dfIm01_label001.pack_forget()
        self.dfIm01_label002.pack_forget()
        self.dfIm01_label003.pack_forget()
        self.dfIm01_label004.pack_forget()
        self.dfIm01_label005.pack_forget()
        self.image_01(self.frame103, self.clsPreSub_fig_list[self.clsPreSub_log-1][4])
    
    #　圧縮変換一覧のリスト表示    
    def table_01(self, dfTb01_root):
        style002 = ttk.Style()
        style002.configure("Treeview.Heading", font=("BIZ UDPゴシック", 12, "bold"))  # ヘッダーのフォントを設定
        style002.configure("Treeview", font=("BIZ UDPゴシック", 12))  # フォントを設定
        column = ("No.", 'ファイル名', '元サイズ', '圧縮後サイズ')
        self.tree = ttk.Treeview(dfTb01_root, columns=column, height=18)
        # 列の設定
        self.tree.column('#0', width=0, stretch='no')
        self.tree.column('No.', anchor='center', width=60)
        self.tree.column('ファイル名', anchor='w', width=250)
        self.tree.column('元サイズ', anchor='e', width=170)
        self.tree.column('圧縮後サイズ', anchor='e', width=170)
        # 列の見出し設定
        self.tree.heading('#0', text='')
        self.tree.heading('No.', text='No.', anchor='center')
        self.tree.heading('ファイル名', text='ファイル名', anchor='center')
        self.tree.heading('元サイズ', text='元サイズ', anchor='center')
        self.tree.heading('圧縮後サイズ',text='圧縮後サイズ', anchor='center')

        #　読込ファイルと保存ファイルのサイズ容量の合計値を取得
        dfTb01_pre_capacity_total = 0
        dfTb01_resize_capacity_total = 0
        for dfTb01_fig_list_line_data in self.clsPreSub_fig_list:
            dfTb01_pre_capacity_total += dfTb01_fig_list_line_data[2]    #　読込ファイルのサイズ容量の合計値
            dfTb01_resize_capacity_total += dfTb01_fig_list_line_data[3]    #　保存ファイルのサイズ容量の合計値
        
        # レコードの追加
        self.tree.insert(parent = '',
                    index = 'end',
                    iid = 0,
                    values = ("", "", "●合計 "+str(dfTb01_pre_capacity_total)+" kb ", "●合計 "+str(dfTb01_resize_capacity_total)+" kb "))
        for dfTb01_fig_list_line_data in self.clsPreSub_fig_list:
            value = (dfTb01_fig_list_line_data[0],
                     dfTb01_fig_list_line_data[1],
                     str(dfTb01_fig_list_line_data[2])+" kb ",
                     str(dfTb01_fig_list_line_data[3])+" kb ")
            self.tree.insert(parent = '',
                        index = 'end',
                        iid = int(dfTb01_fig_list_line_data[0]),
                        values = value)

        # ウィジェットの配置
        self.tree.pack()
    
    def image_01(self, dfIm01_root, dfIm01_img_fname_fullpath):
        self.dfIm01_label001 = tk.Label(dfIm01_root,
                                   text = "－－－－－－　元画像　－－－－－－",
                                   font = ("BIZ UDPゴシック", 16, "bold"))
        self.dfIm01_label001.pack()

        #　オリジナル画像を表示
        self.clsPreSub_dfIm01_original_img = Image.open(dfIm01_img_fname_fullpath)
        self.clsPreSub_dfIm01_preview_image001 = self.clsPreSub_dfIm01_original_img.resize((360, 270), Image.Resampling.LANCZOS)  # 高品質な縮小
        self.photo001 = ImageTk.PhotoImage(self.clsPreSub_dfIm01_preview_image001)
        # ラベルに画像を設定
        self.dfIm01_label002 = tk.Label(dfIm01_root, image=self.photo001)
        self.dfIm01_label002.pack()

        self.dfIm01_label003 = tk.Label(dfIm01_root,
                                        text = "↓",
                                        font = ("BIZ UDPゴシック", 30, "bold"),
                                        pady = 10)
        self.dfIm01_label003.pack()

        self.dfIm01_label004 = tk.Label(dfIm01_root,
                                        text = "－－－－－　圧縮後画像　－－－－－",
                                        font = ("BIZ UDPゴシック", 16, "bold"))
        self.dfIm01_label004.pack()

        #　圧縮画像を表示
        self.clsPreSub_dfIm01_resize_img = self.clsPreSub_dfIm01_original_img.resize((int(self.clsPreSub_dfIm01_original_img.width * int(global_init_state_file[1][2])/100),
                                                                                      int(self.clsPreSub_dfIm01_original_img.height * int(global_init_state_file[1][2])/100)))
        self.clsPreSub_dfIm01_preview_image002 = self.clsPreSub_dfIm01_resize_img.resize((360, 270), Image.Resampling.LANCZOS)  # 高品質な縮小
        self.photo002 = ImageTk.PhotoImage(self.clsPreSub_dfIm01_preview_image002)
        # ラベルに画像を設定
        self.dfIm01_label005 = tk.Label(dfIm01_root, image=self.photo002)
        self.dfIm01_label005.pack()

    #　圧縮率表示
    def spinbox_01(self, dfSp01_root):
        label101 = tk.Label(dfSp01_root,
                            text = "圧縮率 ",
                            font = ("BIZ UDPゴシック", 16, "bold"))
        spinbox101 = tk.Spinbox(dfSp01_root,
                                textvariable = self.clsPreSub_entry_comp_int,
                                font = ("BIZ UDPゴシック", 16, "bold"),
                                justify = tk.RIGHT, 
                                width = 3,
                                bg = "#FFFFEE",
                                from_= 1,    # << 圧縮率の最小値
                                to = 100,    # << 圧縮率の最大値 
                                increment = 5,    # << 数値を５の倍数で上下 
                                command = self.spinbox_comp_increment_adjust)
        spinbox101.bind("<KeyRelease>", self.spinbox_comp_bind_input_num)
        label102 = tk.Label(dfSp01_root,
                            text = "　　　　　　　　　",
                            font = ("BIZ UDPゴシック", 16, "bold"))
        label103 = tk.Label(dfSp01_root,
                            text = "（1～100％の範囲で指定してください。）",
                            font = ("BIZ UDPゴシック", 12))
        label101.grid(row=0, column=0, sticky=tk.W)
        spinbox101.grid(row=0, column=1, sticky=tk.W)
        label102.grid(row=0, column=2, sticky=tk.W)
        label103.grid(row=1, column=0, columnspan=3, sticky=tk.W, pady=5)
    
    def spinbox_comp_increment_adjust(self):
        self.clsPreSub_dfSCincrement_input_num = self.clsPreSub_entry_comp_int.get()
        if self.clsPreSub_dfSCincrement_input_num%5==0:    #　５で割り切れるときはそのまま
            self.clsPreSub_entry_comp_int.set(self.clsPreSub_dfSCincrement_input_num)
        elif self.clsPreSub_dfSCincrement_input_num<=1:    #　１以下になったときは１にする 
            self.clsPreSub_entry_comp_int.set(1)
        elif (self.clsPreSub_dfSCincrement_input_num-self.clsPreSub_previous_comp_num)>=0:
            self.clsPreSub_entry_comp_int.set(self.clsPreSub_dfSCincrement_input_num//5*5)
        elif (self.clsPreSub_dfSCincrement_input_num-self.clsPreSub_previous_comp_num)<0:
            self.clsPreSub_entry_comp_int.set(self.clsPreSub_dfSCincrement_input_num//5*5+5)
        self.clsPreSub_previous_comp_num = self.clsPreSub_entry_comp_int.get()
        global_init_state_file.loc[2,1] = self.clsPreSub_entry_comp_int.get()
        self.dfIm01_label001.pack_forget()
        self.dfIm01_label002.pack_forget()
        self.dfIm01_label003.pack_forget()
        self.dfIm01_label004.pack_forget()
        self.dfIm01_label005.pack_forget()
        self.image_01(self.frame103, self.clsPreSub_fig_list[self.clsPreSub_log-1][4])
        self.make_figlist()
        self.tree.pack_forget()
        self.table_01(self.frame102)
    
    def spinbox_comp_bind_input_num(self, event):
        #　数値以外を入力したとき、元の数値に戻る
        try:    
            self.clsPreSub_entry_comp_int.get()
        except:
            self.clsPreSub_entry_comp_int.set(self.clsPreSub_previous_comp_num)
        #　数値が０以下になったときに１、１００を超えたときに１００にする
        if self.clsPreSub_entry_comp_int.get() > 100:
            self.clsPreSub_entry_comp_int.set(100)
        elif self.clsPreSub_entry_comp_int.get()<=0:
            self.clsPreSub_entry_comp_int.set(1)
        self.clsPreSub_previous_comp_num = self.clsPreSub_entry_comp_int.get()
        global_init_state_file.loc[2,1] = self.clsPreSub_entry_comp_int.get()
        self.dfIm01_label001.pack_forget()
        self.dfIm01_label002.pack_forget()
        self.dfIm01_label003.pack_forget()
        self.dfIm01_label004.pack_forget()
        self.dfIm01_label005.pack_forget()
        self.image_01(self.frame103, self.clsPreSub_fig_list[self.clsPreSub_log-1][4])
        self.make_figlist()
        self.tree.pack_forget()
        self.table_01(self.frame102)

    def convert_button_in_preview(self, dfPB01_root):        
        style003 = ttk.Style()
        style003.configure("St02.TButton",
                        font = ("BIZ UDPゴシック",12,"bold"),
                        padding = [20, 20],
                        foreground = "#666666")            
        button001 = ttk.Button(dfPB01_root,
                               text = "変換実行",
                               default = "active",
                               style = "St02.TButton",
                               width = 9,
                               command = self.run_convert_in_preview) 
        button001.pack()
    
    def run_convert_in_preview(self):
        if not os.path.exists(global_init_state_file[1][1]):
            dfRC_sub_window01 = tk.Toplevel(self)
            dfRC_sub_window01.title("一覧確認用")
            dfRC_sub_window01.geometry("1000x300")
                        
            dfRC_label101 = tk.Label(dfRC_sub_window01,
                                font = ("BIZ UDPゴシック", 18, "bold"),
                                fg = "red",
                                text = "\n\n")
            dfRC_label102 = tk.Label(dfRC_sub_window01,
                                font = ("BIZ UDPゴシック", 18, "bold"),
                                fg = "red",
                                text = "保存フォルダがありません。")
            dfRC_label103 = tk.Label(dfRC_sub_window01,
                                font = ("BIZ UDPゴシック", 16, "bold"),
                                text = f"\n")
            dfRC_label104 = tk.Label(dfRC_sub_window01,
                                font=("BIZ UDPゴシック", 14), 
                                text = "\n（※保存フォルダを設定してください。）")
            
            dfRC_label101.grid(row=0, column=0 ,columnspan=2, sticky="nsew")
            dfRC_label102.grid(row=1, column=0 ,columnspan=2, sticky="nsew", pady=10)
            dfRC_label103.grid(row=2, column=0, sticky=tk.E, pady=20)             
            dfRC_label104.grid(row=2, column=1, sticky=tk.W)

            dfRC_sub_window01.grid_columnconfigure(0, weight=1)
            dfRC_sub_window01.grid_columnconfigure(1, weight=1) 
        else:  
            bar_window = tk.Toplevel(self)
            bar_window.title("")
            bar_label001 = tk.Label(bar_window, text="作業進捗状況", font=("BIZ UDPゴシック", 20))
            bar_label001.grid(row=0, column=0, columnspan=5)
            bar_label002 = tk.Label(bar_window, text="　　", font=("BIZ UDPゴシック", 12))
            bar_label002.grid(row=1, column=1)   
            bar_label003 = tk.Label(bar_window, text="0％", font=("BIZ UDPゴシック", 12))
            bar_label003.grid(row=1, column=2)   
            bar = ttk.Progressbar(bar_window, length=600, mode="determinate", maximum=self.clsPreSub_max_num)
            bar.grid(row=1, column=3, padx=10, pady=20)
            bar_label004 = tk.Label(bar_window, text="100％", font=("BIZ UDPゴシック", 12))
            bar_label004.grid(row=1, column=4)        
            bar_label005 = tk.Label(bar_window, text="　　", font=("BIZ UDPゴシック", 12))
            bar_label005.grid(row=1, column=5)        
            bar_count = 1
            
            for ind_img_data in self.clsPreSub_fig_list:
                if ind_img_data[5] == "":
                    pass
                elif not os.path.isdir(os.path.join(global_init_state_file[1][1],ind_img_data[5])):
                    os.makedirs(os.path.join(global_init_state_file[1][1],ind_img_data[5]))
                dfRC01_original_img = Image.open(ind_img_data[4])
                dfRC01_resize_img = dfRC01_original_img.resize((int(dfRC01_original_img.width * int(global_init_state_file[1][2])/100),
                                                                int(dfRC01_original_img.height * int(global_init_state_file[1][2])/100)))
                dfRC01_resize_img.save(ind_img_data[6])
                
                bar.configure(value=bar_count)
                bar.update()
                bar_count += 1

            global_init_state_file.to_csv(global_init_state_fname_fullpath, encoding="utf-8", index=False, header=False)
        
            bar_window.destroy()
            self.destroy()

#　▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲　クラス（６）＞ ７－２　ファイル変換実行（プレビューサブウインドウ）【clsPreSub】　▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲


#　▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼　クラス（７）＞ ８　ファイル変換実行【clsRunConv】　▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
class Run_convert(tk.Frame):
    def __init__(self,
                 clsRunConv_root,
                 clsRunConv_pos_x,
                 clsRunConv_pos_y):
        super().__init__(clsRunConv_root)
        
        style004 = ttk.Style()
        style004.configure("St03.TButton",
                           font = ("BIZ UDPゴシック",12,"bold"),
                           padding = [20, 20],
                           foreground = "#666666")            
        button001 = ttk.Button(self,
                               text = "変換実行",
                               default = "active",
                               style = "St03.TButton",
                               width = 9,
                               command = self.run_convert_start) 
        button001.pack()
        
        self.place(x=clsRunConv_pos_x, y=clsRunConv_pos_y)

    def conv_make_figlist(self):
        self.clsPreSub_fig_list = []
        clsPreSub_fig_num = 1
        
        for clsPreSub_curDir, _, clsPreSub_files in os.walk(global_init_state_file[1][0]):
            for clsPreSub_ind_fname in clsPreSub_files:
                try:            
                    clsPreSub_ind_read_fullpath = path_check(os.path.join(clsPreSub_curDir, clsPreSub_ind_fname))
                    clsPreSub_ind_file_in_folder = clsPreSub_curDir.replace(global_init_state_file[1][0],"")[1:]
                    clsPreSub_ind_save_fullpath = os.path.join(global_init_state_file[1][1], clsPreSub_ind_file_in_folder, clsPreSub_ind_fname)
                    clsPreSub_buffer = BytesIO()
        
                    clsPreSub_ind_pre_img = Image.open(clsPreSub_ind_read_fullpath)

                    #　ファイル名を「ファイル名、拡張子（ドットを含む）」とに分割し、ドットを除いた拡張子を取得
                    clsPreSub_file_ind_file_extension = os.path.splitext(clsPreSub_ind_fname)[1][1:]
                    if clsPreSub_file_ind_file_extension == "jpg" or clsPreSub_file_ind_file_extension == "JPG":
                        clsPreSub_file_ind_file_extension = "jpeg"
                    if clsPreSub_file_ind_file_extension == "tif" or clsPreSub_file_ind_file_extension == "TIF":
                        clsPreSub_file_ind_file_extension = "tiff"
                        
                    clsPreSub_ind_resize_img = clsPreSub_ind_pre_img.resize((int(clsPreSub_ind_pre_img.width * int(global_init_state_file[1][2])/100), 
                                                         int(clsPreSub_ind_pre_img.height * int(global_init_state_file[1][2])/100)))
        
                    clsPreSub_ind_resize_img.save(clsPreSub_buffer, format=clsPreSub_file_ind_file_extension) 
                    clsPreSub_ind_pre_capacity = math.ceil(os.path.getsize(clsPreSub_ind_read_fullpath)/1024)
                    clsPreSub_ind_resize_capacity = math.ceil(clsPreSub_buffer.tell()/1024)
                    self.clsPreSub_fig_list.append([clsPreSub_fig_num,                #　（リスト１列目）ファイル番号
                                                    clsPreSub_ind_fname,              #　（リスト２列目）ファイル名
                                                    clsPreSub_ind_pre_capacity,       #　（リスト３列目）読込ファイルサイズ
                                                    clsPreSub_ind_resize_capacity,    #　（リスト４列目）保存ファイルサイズ
                                                    clsPreSub_ind_read_fullpath,      #　（リスト５列目）読込ファイルのフルパス
                                                    clsPreSub_ind_file_in_folder,     #　（リスト６列目）読込ファイル内のフォルダ（読込フォルダ直下ではなく、フォルダ内に画像がある場合）
                                                    clsPreSub_ind_save_fullpath])     #　（リスト７列目）保存ファイルのフルパス
                    clsPreSub_fig_num += 1
                except:
                    pass
        self.clsPreSub_max_num = clsPreSub_fig_num - 1
    
    def run_convert_start(self):
        self.conv_make_figlist()

        if not os.path.exists(global_init_state_file[1][1]):
            dfRC_sub_window01 = tk.Toplevel(self)
            dfRC_sub_window01.title("一覧確認用")
            dfRC_sub_window01.geometry("1000x300")
                        
            dfRC_label101 = tk.Label(dfRC_sub_window01,
                                font = ("BIZ UDPゴシック", 18, "bold"),
                                fg = "red",
                                text = "\n\n")
            dfRC_label102 = tk.Label(dfRC_sub_window01,
                                font = ("BIZ UDPゴシック", 18, "bold"),
                                fg = "red",
                                text = "保存フォルダがありません。")
            dfRC_label103 = tk.Label(dfRC_sub_window01,
                                font = ("BIZ UDPゴシック", 16, "bold"),
                                text = f"\n")
            dfRC_label104 = tk.Label(dfRC_sub_window01,
                                font=("BIZ UDPゴシック", 14), 
                                text = "\n（※保存フォルダを設定してください。）")
            
            dfRC_label101.grid(row=0, column=0 ,columnspan=2, sticky="nsew")
            dfRC_label102.grid(row=1, column=0 ,columnspan=2, sticky="nsew", pady=10)
            dfRC_label103.grid(row=2, column=0, sticky=tk.E, pady=20)             
            dfRC_label104.grid(row=2, column=1, sticky=tk.W)

            dfRC_sub_window01.grid_columnconfigure(0, weight=1)
            dfRC_sub_window01.grid_columnconfigure(1, weight=1)        
        elif self.clsPreSub_max_num == 0:
            dfRC_sub_window02 = tk.Toplevel(self)
            dfRC_sub_window02.title("一覧確認用")
            dfRC_sub_window02.geometry("1000x300")
                        
            dfRC_label201 = tk.Label(dfRC_sub_window02,
                                font = ("BIZ UDPゴシック", 18, "bold"),
                                fg = "red",
                                text = "\n\n")
            dfRC_label202 = tk.Label(dfRC_sub_window02,
                                font = ("BIZ UDPゴシック", 18, "bold"),
                                fg = "red",
                                text = "読込フォルダに画像がありません。")
            dfRC_label203 = tk.Label(dfRC_sub_window02,
                                font = ("BIZ UDPゴシック", 16, "bold"),
                                text = f"\n")
            dfRC_label204 = tk.Label(dfRC_sub_window02,
                                font=("BIZ UDPゴシック", 14), 
                                text = "\n（※画像のある読込フォルダに設定してください。）")
            
            dfRC_label201.grid(row=0, column=0 ,columnspan=2, sticky="nsew")
            dfRC_label202.grid(row=1, column=0 ,columnspan=2, sticky="nsew", pady=10)
            dfRC_label203.grid(row=2, column=0, sticky=tk.E, pady=20)             
            dfRC_label204.grid(row=2, column=1, sticky=tk.W)

            dfRC_sub_window02.grid_columnconfigure(0, weight=1)
            dfRC_sub_window02.grid_columnconfigure(1, weight=1)
        else:                    
            bar_window = tk.Toplevel(self)
            bar_window.title("")
            bar_label001 = tk.Label(bar_window, text="作業進捗状況", font=("BIZ UDPゴシック", 20))
            bar_label001.grid(row=0, column=0, columnspan=5)
            bar_label002 = tk.Label(bar_window, text="　　", font=("BIZ UDPゴシック", 12))
            bar_label002.grid(row=1, column=1)   
            bar_label003 = tk.Label(bar_window, text="0％", font=("BIZ UDPゴシック", 12))
            bar_label003.grid(row=1, column=2)   
            bar = ttk.Progressbar(bar_window, length=600, mode="determinate", maximum=self.clsPreSub_max_num)
            bar.grid(row=1, column=3, padx=10, pady=20)
            bar_label004 = tk.Label(bar_window, text="100％", font=("BIZ UDPゴシック", 12))
            bar_label004.grid(row=1, column=4)        
            bar_label005 = tk.Label(bar_window, text="　　", font=("BIZ UDPゴシック", 12))
            bar_label005.grid(row=1, column=5)        
            bar_count = 1
            
            for ind_img_data in self.clsPreSub_fig_list:
                if ind_img_data[5] == "":
                    pass
                elif not os.path.isdir(os.path.join(global_init_state_file[1][1],ind_img_data[5])):
                    os.makedirs(os.path.join(global_init_state_file[1][1],ind_img_data[5]))
                dfRC01_original_img = Image.open(ind_img_data[4])
                dfRC01_resize_img = dfRC01_original_img.resize((int(dfRC01_original_img.width * int(global_init_state_file[1][2])/100),
                                                                int(dfRC01_original_img.height * int(global_init_state_file[1][2])/100)))
                dfRC01_resize_img.save(ind_img_data[6])
                                
                bar.configure(value=bar_count)
                bar.update()
                bar_count += 1
            
            bar_window.destroy()
#　▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲　クラス（７）＞ ８　ファイル変換実行【clsRunConv】　▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
       
if __name__ == "__main__":
    Application()