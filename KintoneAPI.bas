Option Explicit

' API設定
Private m_config As KintoneConfig

' レコード取得用の構造体
Public Type KintoneRecord
    RecordID As String
    Fields As Object
End Type

' 初期化
Public Sub Initialize()
    ' 設定の読み込み
    On Error Resume Next
    Set m_config = LoadConfig(GetWindowsUser())
    If Err.Number <> 0 Then
        ' 設定が存在しない場合は新規作成
        m_config.Subdomain = "iyell"
        m_config.APIToken = ""
        m_config.LastUser = GetWindowsUser()
    End If
    On Error GoTo 0
End Sub

' APIトークンを設定
Public Sub SetAPIToken(token As String)
    m_config.APIToken = token
    SaveConfig m_config, GetWindowsUser()
End Sub

' レコードを取得
Public Function GetRecords(appID As String) As Collection
    Dim http As Object
    Dim url As String
    Dim response As String
    Dim records As Collection
    Dim jsonParser As Object
    
    Set records = New Collection
    Set jsonParser = New JsonParser
    
    ' HTTPリクエストの設定
    Set http = CreateObject("MSXML2.XMLHTTP")
    url = "https://" & m_config.Subdomain & ".cybozu.com/k/v1/records.json?app=" & appID
    
    With http
        .Open "GET", url, False
        .setRequestHeader "X-Cybozu-API-Token", m_config.APIToken
        .send
    End With
    
    ' レスポンスの処理
    If http.Status = 200 Then
        response = http.responseText
        Dim jsonData As Object
        Set jsonData = jsonParser.ParseJson(response)
        
        ' レコードの変換
        Dim record As KintoneRecord
        Set record = jsonParser.ConvertToRecord(jsonData)
        records.Add record
    Else
        ' エラー処理
        HandleAPIError http.Status, http.responseText
    End If
    
    Set GetRecords = records
End Function

' Windowsユーザー名の取得
Private Function GetWindowsUser() As String
    GetWindowsUser = Environ("USERNAME")
End Function

' エラーハンドリング
Private Sub HandleAPIError(status As Long, responseText As String)
    Dim errMsg As String
    errMsg = "APIリクエストに失敗しました。" & vbCrLf & _
             "ステータス: " & status & vbCrLf & _
             "レスポンス: " & responseText
    
    ' エラーログの出力
    LogError errMsg
    
    ' ユーザーへの通知
    MsgBox errMsg, vbCritical, "Kintone API エラー"
End Sub 