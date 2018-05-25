Attribute VB_Name = "M_ExportModule"
Option Explicit


'--------------------------------------------------------------------------------------------------
' ☆[VBA プロジェクト オブジェクト モデル]の信頼設定
' １．Microsoft Office ボタンをクリックし、[Excel のオプション] をクリックします。
' ２．[セキュリティ センター] をクリックします｡
' ３．[セキュリティ センターの設定] をクリックします｡
' ４．[マクロの設定] をクリックします｡
' ５．[VBA プロジェクト オブジェクト モデルへのアクセスを信頼する] チェック ボックスをオンにします。
' ６．[OK] をクリックして [Excel のオプション] ダイアログ ボックスを閉じます。
'
' ☆[Microsoft Visual Basic for Applications Extensibility]の有効化
' １．VBEのツール(T)→参照設定(R)で参照設定ウィンドウを開く
' ２．リスト中の「Microsoft Visual Basic for Applications Extensibility」左のチェックボックスをチェック
'--------------------------------------------------------------------------------------------------


'--------------------------------------------------------------------------------------------------
' 全モジュール(VBAコード)のエクスポート
'--------------------------------------------------------------------------------------------------
Public Sub ExportAll()
    Dim module                  As VBComponent      '// モジュール
    Dim moduleList              As VBComponents     '// VBAプロジェクトの全モジュール
    Dim extension                                   '// モジュールの拡張子
    Dim sPath As String                             '// 処理対象ブックのパス
    Dim sFilePath                                   '// エクスポートファイルパス
    Dim TargetBook As Workbook                      '// 処理対象ブックオブジェクト
    Dim Count As Long
    
    If Workbooks.Count > 1 Then
        MsgBox "ワークブックが２つ以上開かれています。", vbOKOnly, "エラー"
        Exit Sub
    End If
    
    Dim targetPath As String
    
    '------------------------------------------------------
    ' フォルダの選択ダイアログを開く
    '------------------------------------------------------
    With Application.FileDialog(msoFileDialogFolderPicker)
        .AllowMultiSelect = True
        .title = "エクスポート先のフォルダを選択"
    
        If .Show = True Then
            targetPath = .SelectedItems(1)
        End If
    End With

    If targetPath = "" Then
        
        ' フォルダが選択されなかったとき
        Exit Sub
    
    End If
    
    Set TargetBook = ActiveWorkbook
    sPath = ActiveWorkbook.Path
    
    If Dir(targetPath, vbDirectory) = "" Then
        MsgBox targetPath & " が存在しません。", vbOKOnly, "エラー"
        Exit Sub
    End If
    
    '// 処理対象ブックのモジュール一覧を取得
    Set moduleList = TargetBook.VBProject.VBComponents
    
    '// VBAプロジェクトに含まれる全てのモジュールをループ
    For Each module In moduleList
        
        If (module.Type = vbext_ct_ClassModule) Then
            '// クラス
            extension = "cls"
        
        ElseIf (module.Type = vbext_ct_MSForm) Then
            '// フォーム
            '// .frxも一緒にエクスポートされる
            extension = "frm"
        
        ElseIf (module.Type = vbext_ct_StdModule) Then
            '// 標準モジュール
            extension = "bas"
        
        ElseIf (module.Type = vbext_ct_Document) Then
            '// ドキュメント（シート）
            extension = "cls"
        
        ElseIf (module.Type = vbext_ct_ActiveXDesigner) Then
            '// ActiveXデザイナ
            '// エクスポート対象外のため次ループへ
            GoTo CONTINUE
        
        Else
            '// その他
            '// エクスポート対象外のため次ループへ
            GoTo CONTINUE
        
        End If
        
        '// エクスポート実施
        sFilePath = targetPath & "\" & module.Name & "." & extension
        Application.StatusBar = sFilePath & " をエクスポート中..."
        
        Call module.Export(sFilePath)
        Count = Count + 1
        
        '// 出力先確認用ログ出力
        Debug.Print sFilePath

CONTINUE:
    Next
    
    Application.StatusBar = False
    
    MsgBox "全モジュールのエクスポートが終わりました" & vbCrLf _
        & vbCrLf _
        & "出力ファイル数 = " & Count _
        , vbOKOnly, "エクスポート完了"

End Sub

