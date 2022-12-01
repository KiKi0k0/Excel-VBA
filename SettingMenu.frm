VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SettingMenu 
   Caption         =   "文字列検索/文字列置換　設定画面"
   ClientHeight    =   2640
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   8780
   OleObjectBlob   =   "SettingMenu.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "SettingMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub OpenFolder_Button_Click()

    Call Setting.OpenFolder

End Sub

Private Sub Return_Button_Click()

    SettingMenu.Hide
    MainMenu.Show

    Call Setting.AddExcel_Information


End Sub

