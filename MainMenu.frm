VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MainMenu 
   Caption         =   "文字列検索/文字列置換"
   ClientHeight    =   6550
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   8780
   OleObjectBlob   =   "MainMenu.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "MainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Setting_Button_Click()

    MainMenu.Hide
    SettingMenu.Show

End Sub
