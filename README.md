# sample-repo

Git practice

Sub file_download(Driver As WebDriver, URL As String, KeyWord As String)

    'see https://www.browserstack.com/guide/download-file-using-selenium-python
    Dim filepath As String
    Dim caps As Capabilities
    
    Call WebDriverManager4TinySelenium.CreateFolderEx("D:\program\Download")
    filepath = ".\Download" 'download to same directory as this excel file
    
    
    Driver.StartEdge (".\Driver\msedgedriver.exe")
    
    Set caps = Driver.CreateCapabilities

    caps.AddPref "download.default_directory", filepath
    caps.AddPref "download.prompt_for_download", False
    
    'caps.SetDownloadPrefs filepath 'this does the above in one line

    Driver.OpenBrowser caps
    
    Driver.NavigateTo URL

    Driver.Wait 500
    
    Driver.FindElement(by.XPath, KeyWord).Click
    Driver.Wait 5000
        

End Sub

Sub InstallDriver()

    Call WebDriverManager4TinySelenium.InstallWebDriver(Edge, "D:\program\")
    MsgBox ("Driverインストール完了")

End Sub


'// zipから中身を取り出して指定の場所に実行ファイルを展開する
'// chromedriver.exe(デフォルトの名前)があるところにchromedriver_94.exeとかで展開できるよう、
'// 元の実行ファイルを上書きしないように一度tempフォルダを作ってから実行ファイルを目的のパスへ移す
'// 普通zipを展開するときは展開先のフォルダを指定するが、
'// この関数はWebDriverの実行ファイルのパスで指定するので注意！(展開するのもexeだけ)
'// 使用例
'//     Extract "C:\Users\yamato\Downloads\chromedriver_win32.zip", "C:\Users\yamato\Downloads\chromedriver_94.exe"
Sub Extract(path_zip As String, path_save_to As String)
    Debug.Print "zipを展開します"
    
    Dim file_Driver As String
    Dim file_check As Boolean
    
    If Right(path_save_to, 1) = "\" Then
        file_Driver = path_save_to & "Driver"
    Else
        file_Driver = path_save_to & "\Driver"
    End If
    
    
    CreateFolderEx file_Driver
        
    'Dim folder_temp As String
    'folder_temp = fso.BuildPath(fso.GetParentFolderName(path_save_to), fso.GetTempName)
    'fso.CreateFolder folder_temp
    'Debug.Print "    一時フォルダ : " & folder_temp

    'PowerShellを使って展開するとマルウェア判定されたので，
    'MS非推奨だがShell.Applicationを使ってzipを解凍する
    'On Error GoTo Catch
    Dim sh As Object
    Set sh = CreateObject("Shell.Application")
    'zipファイルに入っているファイルを指定したフォルダーにコピーする
    '文字列を一度()で評価してからNamespaceに渡さないとエラーが出る
    'sh.Namespace((folder_temp)).CopyHere sh.Namespace((path_zip)).Items
    

    file_check = fso.FileExists(file_Driver & "\msedgedriver.exe")
    Debug.Print file_Driver
    
    If file_check = False Then
        sh.Namespace((file_Driver)).CopyHere sh.Namespace((path_zip)).Items
    Else
        fso.DeleteFolder (file_Driver)
        CreateFolderEx file_Driver
        sh.Namespace((file_Driver)).CopyHere sh.Namespace((path_zip)).Items
    End If



    'Dim path_exe As String
    'path_exe = fso.BuildPath(folder_temp, Dir(folder_temp & "\*.exe"))
    
    'If fso.FileExists(folder_temp) Then
    '    fso.CopyFile folder_temp, path_save_to, True
    'End If

       
    file_check = fso.FileExists(path_zip)
    
    If file_check = True Then fso.DeleteFile (path_zip)
    
    Debug.Print "    展開 : " & path_save_to
    Debug.Print "WebDriverを配置しました"
    Exit Sub
    
'Catch:
    'fso.DeleteFolder folder_temp
    'Err.raise 4002, , "Zipの展開に失敗しました。原因：" & Err.Description
    'Exit Sub
End Sub
