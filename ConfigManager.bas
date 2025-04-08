Option Explicit

Private Const CONFIG_FOLDER As String = "C:\ProgramData\DandoriVBA\config\"
Private Const CONFIG_FILE As String = "kintone_config.dat"

' 設定情報の構造体
Public Type KintoneConfig
    Subdomain As String
    APIToken As String
    LastUser As String
End Type

' 設定の保存
Public Sub SaveConfig(config As KintoneConfig, password As String)
    Dim fso As Object
    Dim configFile As Object
    Dim encryptedData As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' 設定フォルダの作成
    If Not fso.FolderExists(CONFIG_FOLDER) Then
        fso.CreateFolder CONFIG_FOLDER
    End If
    
    ' データの暗号化
    encryptedData = EncryptData(config, password)
    
    ' 設定ファイルの保存
    Set configFile = fso.CreateTextFile(CONFIG_FOLDER & CONFIG_FILE, True)
    configFile.Write encryptedData
    configFile.Close
End Sub

' 設定の読み込み
Public Function LoadConfig(password As String) As KintoneConfig
    Dim fso As Object
    Dim configFile As Object
    Dim encryptedData As String
    Dim config As KintoneConfig
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' 設定ファイルの存在確認
    If Not fso.FileExists(CONFIG_FOLDER & CONFIG_FILE) Then
        Err.Raise vbObjectError + 1, "ConfigManager", "設定ファイルが存在しません"
    End If
    
    ' 設定ファイルの読み込み
    Set configFile = fso.OpenTextFile(CONFIG_FOLDER & CONFIG_FILE, 1)
    encryptedData = configFile.ReadAll
    configFile.Close
    
    ' データの復号化
    config = DecryptData(encryptedData, password)
    
    Set LoadConfig = config
End Function

' データの暗号化
Private Function EncryptData(config As KintoneConfig, password As String) As String
    Dim crypto As Object
    Dim data As String
    Dim key As String
    Dim iv As String
    
    ' 暗号化オブジェクトの作成
    Set crypto = CreateObject("System.Security.Cryptography.RijndaelManaged")
    
    ' キーとIVの生成
    key = GenerateKey(password)
    iv = GenerateIV(password)
    
    ' データの準備
    data = config.Subdomain & "|" & config.APIToken & "|" & config.LastUser
    
    ' 暗号化
    With crypto
        .Key = key
        .IV = iv
        .Mode = 1 ' CBC
        .Padding = 2 ' PKCS7
    End With
    
    ' 暗号化データの生成
    Dim encryptedBytes() As Byte
    encryptedBytes = crypto.CreateEncryptor().TransformFinalBlock(StrConv(data, vbFromUnicode), 0, Len(data))
    
    ' Base64エンコード
    EncryptData = Base64Encode(encryptedBytes)
End Function

' データの復号化
Private Function DecryptData(encryptedData As String, password As String) As KintoneConfig
    Dim crypto As Object
    Dim data As String
    Dim key As String
    Dim iv As String
    Dim config As KintoneConfig
    
    ' 暗号化オブジェクトの作成
    Set crypto = CreateObject("System.Security.Cryptography.RijndaelManaged")
    
    ' キーとIVの生成
    key = GenerateKey(password)
    iv = GenerateIV(password)
    
    ' Base64デコード
    Dim encryptedBytes() As Byte
    encryptedBytes = Base64Decode(encryptedData)
    
    ' 復号化
    With crypto
        .Key = key
        .IV = iv
        .Mode = 1 ' CBC
        .Padding = 2 ' PKCS7
    End With
    
    ' 復号化データの生成
    Dim decryptedBytes() As Byte
    decryptedBytes = crypto.CreateDecryptor().TransformFinalBlock(encryptedBytes, 0, UBound(encryptedBytes) + 1)
    
    ' 文字列に変換
    data = StrConv(decryptedBytes, vbUnicode)
    
    ' データの分割
    Dim parts() As String
    parts = Split(data, "|")
    config.Subdomain = parts(0)
    config.APIToken = parts(1)
    config.LastUser = parts(2)
    
    Set DecryptData = config
End Function

' キーの生成
Private Function GenerateKey(password As String) As String
    ' パスワードから32バイトのキーを生成
    Dim sha256 As Object
    Set sha256 = CreateObject("System.Security.Cryptography.SHA256Managed")
    
    Dim hash() As Byte
    hash = sha256.ComputeHash(StrConv(password, vbFromUnicode))
    
    GenerateKey = StrConv(hash, vbUnicode)
End Function

' IVの生成
Private Function GenerateIV(password As String) As String
    ' パスワードから16バイトのIVを生成
    Dim md5 As Object
    Set md5 = CreateObject("System.Security.Cryptography.MD5CryptoServiceProvider")
    
    Dim hash() As Byte
    hash = md5.ComputeHash(StrConv(password, vbFromUnicode))
    
    GenerateIV = StrConv(hash, vbUnicode)
End Function

' Base64エンコード
Private Function Base64Encode(bytes() As Byte) As String
    Dim xml As Object
    Set xml = CreateObject("MSXML2.DOMDocument")
    
    Dim element As Object
    Set element = xml.createElement("b64")
    element.DataType = "bin.base64"
    element.Text = StrConv(bytes, vbUnicode)
    
    Base64Encode = element.Text
End Function

' Base64デコード
Private Function Base64Decode(base64String As String) As Byte()
    Dim xml As Object
    Set xml = CreateObject("MSXML2.DOMDocument")
    
    Dim element As Object
    Set element = xml.createElement("b64")
    element.DataType = "bin.base64"
    element.Text = base64String
    
    Base64Decode = StrConv(element.Text, vbFromUnicode)
End Function 