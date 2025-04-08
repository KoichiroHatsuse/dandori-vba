Option Explicit

' テスト実行
Public Sub RunTests()
    ' 初期化テスト
    TestInitialize
    
    ' APIトークン設定テスト
    TestSetAPIToken
    
    ' レコード取得テスト
    TestGetRecords
    
    ' 設定保存テスト
    TestSaveConfig
    
    ' 設定読み込みテスト
    TestLoadConfig
    
    MsgBox "すべてのテストが完了しました。", vbInformation, "テスト完了"
End Sub

' 初期化テスト
Private Sub TestInitialize()
    On Error GoTo ErrorHandler
    
    Call KintoneAPI.Initialize
    
    Exit Sub
    
ErrorHandler:
    MsgBox "初期化テストに失敗しました。" & vbCrLf & _
           "エラー: " & Err.Description, vbCritical, "テスト失敗"
End Sub

' APIトークン設定テスト
Private Sub TestSetAPIToken()
    On Error GoTo ErrorHandler
    
    Call KintoneAPI.SetAPIToken("test_token")
    
    Exit Sub
    
ErrorHandler:
    MsgBox "APIトークン設定テストに失敗しました。" & vbCrLf & _
           "エラー: " & Err.Description, vbCritical, "テスト失敗"
End Sub

' レコード取得テスト
Private Sub TestGetRecords()
    On Error GoTo ErrorHandler
    
    Dim records As Collection
    Set records = KintoneAPI.GetRecords("test_app")
    
    If records.Count = 0 Then
        MsgBox "レコードが取得できませんでした。", vbExclamation, "テスト警告"
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "レコード取得テストに失敗しました。" & vbCrLf & _
           "エラー: " & Err.Description, vbCritical, "テスト失敗"
End Sub

' 設定保存テスト
Private Sub TestSaveConfig()
    On Error GoTo ErrorHandler
    
    Dim config As KintoneConfig
    config.Subdomain = "test_subdomain"
    config.APIToken = "test_token"
    config.LastUser = "test_user"
    
    Call SaveConfig(config, "test_password")
    
    Exit Sub
    
ErrorHandler:
    MsgBox "設定保存テストに失敗しました。" & vbCrLf & _
           "エラー: " & Err.Description, vbCritical, "テスト失敗"
End Sub

' 設定読み込みテスト
Private Sub TestLoadConfig()
    On Error GoTo ErrorHandler
    
    Dim config As KintoneConfig
    Set config = LoadConfig("test_password")
    
    If config.Subdomain <> "test_subdomain" Then
        MsgBox "設定の読み込みに失敗しました。", vbExclamation, "テスト警告"
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "設定読み込みテストに失敗しました。" & vbCrLf & _
           "エラー: " & Err.Description, vbCritical, "テスト失敗"
End Sub 