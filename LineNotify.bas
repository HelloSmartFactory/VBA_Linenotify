Attribute VB_Name = "LineNotify"
Option Explicit

'tokenの設定
Const strToken As String = "yourtoken"


Sub doMsg()
    Dim msg As String
    msg = "2020/06/12"
    sendLineNotify (msg)
End Sub


'Line送信
'引数：strメッセージ
Private Sub sendLineNotify(msg As String)
    'オブジェクト生成 '参照設定なし
    Dim objHTTP As Object
    Set objHTTP = CreateObject("MSXML2.XMLHTTP")
    '参照設定ありの場合はこっち microsoft xml v6.0
    'Dim objHTTP As XMLHTTP60
    'Set objHTTP = New XMLHTTP60
    
On Error GoTo errHandler    'エラーは飛ばす
    objHTTP.Open "POST", "https://notify-api.line.me/api/notify", False 'オブジェクト初期化
    'ヘッダ設定
    objHTTP.setRequestHeader "Content-Type", "application/x-www-form-urlencoded" '解析方法指定
    objHTTP.setRequestHeader "Authorization", "Bearer " + strToken  'headers
    '送信
    objHTTP.send "message=" + msg '+ "&stickerPackageId=1" + "&stickerId=113" 'payload

    'ステータス確認
    If objHTTP.Status = 200 Then    '400リクエストが不正 401アクセストークンが無効　500サーバ内エラーにより失敗
        Debug.Print "うまくいったわ " + objHTTP.responseText
    Else
        Debug.Print "なんかおかしいで！　" + objHTTP.responseText
    End If
    
    Set objHTTP = Nothing   'オブジェクト破棄
Exit Sub

errHandler: 'エラーで飛んでくる
    Dim number As Long: number = Err.number 'エラーコード取得
    
    'エラー別に処理
    Select Case number
        Case -2146197211
            Debug.Print "エラーコード = " & number & vbCrLf & "     指定されたソースが見つかりません。ネットワーク接続を確認してください。"
        Case Else
            Debug.Print "エラーコード = " & number
    End Select
    
    Set objHTTP = Nothing   'オブジェクト破棄
End Sub
