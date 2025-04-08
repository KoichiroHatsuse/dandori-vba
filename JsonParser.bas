Option Explicit

' JSONパース用の関数
Public Function ParseJson(jsonString As String) As Object
    Dim json As Object
    Set json = CreateObject("Scripting.Dictionary")
    
    ' JSON文字列の解析
    Dim jsonObj As Object
    Set jsonObj = JsonToDictionary(jsonString)
    
    ' レコードの取得
    If jsonObj.Exists("records") Then
        Set json = jsonObj("records")
    End If
    
    Set ParseJson = json
End Function

' レコードデータの変換
Public Function ConvertToRecord(jsonData As Object) As KintoneRecord
    Dim record As KintoneRecord
    Dim fields As Object
    Set fields = CreateObject("Scripting.Dictionary")
    
    ' フィールドデータの取得
    If jsonData.Exists("$id") Then
        record.RecordID = jsonData("$id")
    End If
    
    If jsonData.Exists("$revision") Then
        fields.Add "$revision", jsonData("$revision")
    End If
    
    If jsonData.Exists("$updated_by") Then
        fields.Add "$updated_by", jsonData("$updated_by")
    End If
    
    If jsonData.Exists("$created_by") Then
        fields.Add "$created_by", jsonData("$created_by")
    End If
    
    If jsonData.Exists("$updated_time") Then
        fields.Add "$updated_time", jsonData("$updated_time")
    End If
    
    If jsonData.Exists("$created_time") Then
        fields.Add "$created_time", jsonData("$created_time")
    End If
    
    ' その他のフィールドを取得
    Dim key As Variant
    For Each key In jsonData.Keys
        If Left(key, 1) <> "$" Then
            fields.Add key, jsonData(key)
        End If
    Next key
    
    Set record.Fields = fields
    Set ConvertToRecord = record
End Function

' JSON文字列をDictionaryに変換
Private Function JsonToDictionary(jsonString As String) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' 簡易的なJSONパース処理
    ' 実際の実装では、より堅牢なパース処理が必要
    Dim jsonObj As Object
    Set jsonObj = CreateObject("Scripting.Dictionary")
    
    ' ここでJSON文字列を解析してDictionaryに変換
    ' 仮の実装
    jsonObj.Add "records", dict
    
    Set JsonToDictionary = jsonObj
End Function 