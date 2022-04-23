
'***************************************************************************************************
'参考文献:
'***************************************************************************************************
'https://itbyari.wordpress.com/2015/12/08/excel-vba-%E6%9A%97%E5%8F%B7%E5%8C%96%E3%81%A8%E5%BE%A9%E5%8F%B7%E5%8C%96%EF%BC%88tripledes%EF%BC%89/

Option Explicit
'***************************************************************************************************
'定数宣言
'***************************************************************************************************
'初期ベクトル
Public Const INITIALIZATION_VECTOR = "78945612" '必ず8文字分
'暗号用共通鍵
Public Const TRIPLE_DES_KEY = "webcontrolver1.0" '必ず16文字分

'UserID 座標
Const UserID_Column As Single = 1
Const UserID_Row As Single = 2

'UserPassword 座標
Const UserPassword_Column As Single = 2
Const UserPassword_Row As Single = 2

'***************************************************************************************************
'関数呼び出し
'***************************************************************************************************
Sub Encryption()
    ThisWorkbook.ActiveSheet.Cells(UserID_Column, UserID_Row).Value = EncryptStringTripleDES(ThisWorkbook.ActiveSheet.Cells(UserID_Column, UserID_Row).Value)
    ThisWorkbook.ActiveSheet.Cells(UserPassword_Column, UserPassword_Row).Value = EncryptStringTripleDES(ThisWorkbook.ActiveSheet.Cells(UserPassword_Column, UserPassword_Row).Value)
End Sub

Sub Decode()
    ThisWorkbook.ActiveSheet.Cells(UserID_Column, UserID_Row).Value = DecryptStringTripleDES(ThisWorkbook.ActiveSheet.Cells(UserID_Column, UserID_Row).Value)
    ThisWorkbook.ActiveSheet.Cells(UserPassword_Column, UserPassword_Row).Value = DecryptStringTripleDES(ThisWorkbook.ActiveSheet.Cells(UserPassword_Column, UserPassword_Row).Value)
End Sub

'***************************************************************************************************
'機　能：TripleDESによる暗号化（TripleDES暗号化⇒BASE64符号化）
'引　数：暗号化対象平文
'戻り値：暗号文（正常終了） or Null（異常終了)
'備　考：
'***************************************************************************************************
Function EncryptStringTripleDES(plain_string As String) As Variant
    'P0_変数
    'P0-1_変数宣言
    Dim encryption_object As Object
    Dim plain_byte_data() As Byte
    Dim encrypted_byte_data() As Byte
    Dim encrypted_base64_string As String
    
    'P0-2_変数設定
    
    'P0-3_戻り値設定
    EncryptStringTripleDES = Null
    
    'P1_事前処理
    On Error GoTo FunctionError
    
    'P2_主処理
    'P2-1_平文文字列⇒平文バイトデータ
    plain_byte_data = CreateObject("System.Text.UTF8Encoding").GetBytes_4(plain_string)
    
    'P2-2_平文バイトデータ⇒暗号バイトデータ
    Set encryption_object = CreateObject("System.Security.Cryptography.TripleDESCryptoServiceProvider")
    encryption_object.key = CreateObject("System.Text.UTF8Encoding").GetBytes_4(TRIPLE_DES_KEY)
    encryption_object.IV = CreateObject("System.Text.UTF8Encoding").GetBytes_4(INITIALIZATION_VECTOR)
    encrypted_byte_data = _
            encryption_object.CreateEncryptor().TransformFinalBlock(plain_byte_data, 0, UBound(plain_byte_data) + 1)
    
    'P2-3_暗号バイトデータ⇒BASE64符号文字列
    encrypted_base64_string = BytesToBase64(encrypted_byte_data)

    'P3_事後処理
    
    'P4_結果表示 or 戻り値設定
    EncryptStringTripleDES = encrypted_base64_string
    
    'P5_エラーハンドリング
    Exit Function
FunctionError:
    MsgBox "TripleDESによる暗号化に失敗しました。"
End Function

'***************************************************************************************************
'機　能：TripleDESによる復号化（BASE64復号化⇒DES復号化）
'引　数：暗号文
'戻り値：平文（正常終了） or Null（異常終了)
'備　考：
'***************************************************************************************************
Function DecryptStringTripleDES(encrypted_string As String) As Variant
    'P0_変数
    'P0-1_変数宣言
    Dim encryption_object As Object
    Dim encrypted_byte_data() As Byte
    Dim plain_byte_data() As Byte
    Dim plain_string As String
     
    'P0-2_変数設定
    
    'P0-3_戻り値設定
    DecryptStringTripleDES = Null
    
    'P1_事前処理
    On Error GoTo FunctionError
    
    'P2_主処理
    'P2-1_BASE64符号文字列⇒DES暗号バイトデータ
    encrypted_byte_data = Base64toBytes(encrypted_string)
    
    'P2-2_DES暗号バイトデータ⇒平文バイトデータ
    Set encryption_object = CreateObject("System.Security.Cryptography.TripleDESCryptoServiceProvider")
    encryption_object.key = CreateObject("System.Text.UTF8Encoding").GetBytes_4(TRIPLE_DES_KEY)
    encryption_object.IV = CreateObject("System.Text.UTF8Encoding").GetBytes_4(INITIALIZATION_VECTOR)
    plain_byte_data = encryption_object.CreateDecryptor().TransformFinalBlock(encrypted_byte_data, 0, UBound(encrypted_byte_data) + 1)
            
    'P2-3_平文バイトデータ⇒平文文字列化
    plain_string = CreateObject("System.Text.UTF8Encoding").GetString(plain_byte_data)
    
    'P3_事後処理
    
    'P4_結果表示 or 戻り値設定
    DecryptStringTripleDES = plain_string
    
    'P5_エラーハンドリング
    Exit Function
FunctionError:
    MsgBox "TripleDESによる復号化に失敗しました。"
End Function

'***************************************************************************************************
'関数名：BytesToBase64
'機　能：Byte配列→base64文字列への変換
'引　数：Byte配列
'戻り値：base64文字列
'備　考：
'***************************************************************************************************
Function BytesToBase64(varBytes() As Byte) As String
    With CreateObject("MSXML2.DomDocument").createElement("b64")
        .DataType = "bin.base64"
        .nodeTypedValue = varBytes
        BytesToBase64 = Replace(.text, vbLf, "") '無意味に改行が含まれてしまうので除去
    End With
End Function

'***************************************************************************************************
'関数名：Base64toBytes
'機　能：base64文字列→Byte配列への変換
'引　数：base64文字列
'戻り値：Byte配列
'備　考：
'***************************************************************************************************
 Function Base64toBytes(varStr As String) As Byte()
    With CreateObject("MSXML2.DOMDocument").createElement("b64")
         .DataType = "bin.base64"
         .text = varStr
         Base64toBytes = .nodeTypedValue
    End With
 End Function
