Attribute VB_Name = "modFolderScan"
Option Explicit

'==============================================================
' FolderScan - フォルダ内ファイル一覧出力マクロ
'
' 指定したフォルダ以下（サブフォルダも含む）にあるすべての
' ファイルを調べ、ファイル名・フォルダパス・サイズ・更新日時・
' 作成日時を一覧シートに出力します。
'
' 作者: 天電輝合同会社 (TENDENKO)
' ライセンス: MIT License
'
' ---- 使い方 ----
' 1. リボンまたはAlt+F8から「Main_一覧出力実行」を実行
' 2. フォルダ選択ダイアログで、一覧化したいフォルダを選ぶ
' 3. 「一覧」シートに、該当フォルダ以下の全ファイルが出力される
'
' ---- カスタマイズポイント（下の設定セクションを書き換えるだけ） ----
' ・INCLUDE_SUBFOLDERS : サブフォルダの中まで対象にするか
' ・OUTPUT_SHEET_NAME  : 一覧を書き出すシート名
'
' ---- 注意事項 ----
' ファイルの並び順はOS（Windows）が返す順序に依存し、必ずしも
' 名前順・パス順にはなりません。並べ替えたい場合はExcel側の
' 「並べ替え」機能をご利用ください。
' アクセス権限がないフォルダ・ファイルはスキップされます。
'==============================================================

'---------------- 設定セクション（ここだけ書き換えればOK） ----------------

' サブフォルダの中にあるファイルも対象にするか
Private Const INCLUDE_SUBFOLDERS As Boolean = True

' 一覧を書き出すシート名
Private Const OUTPUT_SHEET_NAME As String = "一覧"

'---------------------------------------------------------------------

'==============================================================
' Main_一覧出力実行
'   エントリーポイント。ボタンやAlt+F8から呼び出す。
'==============================================================
Sub Main_一覧出力実行()

    Dim targetFolder As String

    On Error GoTo ErrHandler

    ' --- フォルダ選択 ---
    targetFolder = PickFolder()
    If targetFolder = "" Then
        MsgBox "フォルダが選択されませんでした。処理を中止します。", vbInformation
        Exit Sub
    End If

    ' --- ファイル収集（再帰的にサブフォルダも調べる） ---
    Dim results As New Collection
    Dim skippedCount As Long
    skippedCount = 0

    Application.StatusBar = "フォルダを調べています... " & targetFolder
    DoEvents

    CollectFiles targetFolder, results, skippedCount

    If results.Count = 0 Then
        Application.StatusBar = False
        MsgBox "対象フォルダにファイルが見つかりませんでした。" & vbCrLf & _
               "フォルダ: " & targetFolder, vbExclamation
        Exit Sub
    End If

    ' --- 「一覧」シートへ書き出し ---
    WriteResultSheet results, targetFolder

    Application.StatusBar = False

    ' --- 完了メッセージ ---
    Dim msg As String
    msg = "一覧出力が完了しました。" & vbCrLf & _
          "対象フォルダ: " & targetFolder & vbCrLf & _
          "出力件数: " & results.Count & " 件" & vbCrLf & vbCrLf & _
          "結果は「" & OUTPUT_SHEET_NAME & "」シートに出力しました。"

    If skippedCount > 0 Then
        msg = msg & vbCrLf & vbCrLf & _
              "【ご注意】アクセス権限がないなどの理由で、" & skippedCount & " 件のフォルダ・ファイルをスキップしました。"
    End If

    MsgBox msg, vbInformation, "完了"

    Exit Sub

ErrHandler:
    Application.StatusBar = False
    MsgBox "予期しないエラーが発生しました。" & vbCrLf & _
           "エラー番号: " & Err.Number & vbCrLf & _
           "内容: " & Err.Description, vbCritical, "エラー"
End Sub

'==============================================================
' Main_結果シート削除
'   「一覧」シートを削除する。
'   配布・再テスト前のクリーンアップ用のエントリーポイント。
'==============================================================
Sub Main_結果シート削除()

    If Not SheetExists(OUTPUT_SHEET_NAME) Then
        MsgBox "「" & OUTPUT_SHEET_NAME & "」シートは存在しません。", vbInformation, "削除不要"
        Exit Sub
    End If

    ' 削除すると全シートがなくなってしまう場合は中止（Excelは最低1シート必要）
    If ThisWorkbook.Worksheets.Count - 1 < 1 Then
        MsgBox "削除すると全てのシートがなくなってしまうため中止しました。" & vbCrLf & _
               "操作用のシートなどを別に1つ残してから実行してください。", vbCritical, "中止"
        Exit Sub
    End If

    If MsgBox("「" & OUTPUT_SHEET_NAME & "」シートを削除します。" & vbCrLf & _
              "元に戻せません。よろしいですか？", vbYesNo + vbExclamation, "確認") <> vbYes Then
        Exit Sub
    End If

    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets(OUTPUT_SHEET_NAME).Delete
    Application.DisplayAlerts = True

    MsgBox "削除が完了しました。", vbInformation, "完了"

End Sub

'==============================================================
' 指定シートが存在するかどうかを判定する
'==============================================================
Private Function SheetExists(ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    SheetExists = Not (ws Is Nothing)
End Function

'==============================================================
' フォルダ選択ダイアログを表示し、選択されたパスを返す
'==============================================================
Private Function PickFolder() As String

    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)

    fd.Title = "一覧化したいフォルダを選択してください"
    fd.AllowMultiSelect = False

    ' 初期表示フォルダ：このブックが置かれているフォルダを開く
    If ThisWorkbook.Path <> "" Then
        fd.InitialFileName = ThisWorkbook.Path & "\"
    End If

    If fd.Show = -1 Then
        PickFolder = fd.SelectedItems(1)
    Else
        PickFolder = ""
    End If

End Function

'==============================================================
' 指定フォルダ（必要に応じてサブフォルダも）にあるファイルを
' 再帰的に集め、resultsコレクションに追加する。
' 各要素は Array(ファイル名, フォルダパス, サイズ, 更新日時, 作成日時)
'==============================================================
Private Sub CollectFiles(ByVal folderPath As String, ByRef results As Collection, ByRef skippedCount As Long)

    On Error GoTo FolderErrHandler

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim fld As Object
    Set fld = fso.GetFolder(folderPath)

    Dim f As Object
    For Each f In fld.Files
        On Error Resume Next
        Dim rec(4) As Variant
        rec(0) = f.Name
        rec(1) = fld.Path
        rec(2) = f.Size
        rec(3) = f.DateLastModified
        rec(4) = f.DateCreated

        If Err.Number <> 0 Then
            skippedCount = skippedCount + 1
            Err.Clear
        Else
            results.Add rec
        End If
        On Error GoTo FolderErrHandler
    Next f

    If INCLUDE_SUBFOLDERS Then
        Dim subFld As Object
        For Each subFld In fld.SubFolders
            CollectFiles subFld.Path, results, skippedCount
        Next subFld
    End If

    Exit Sub

FolderErrHandler:
    ' アクセス権限がないフォルダなどは、スキップしてカウントだけ増やす
    skippedCount = skippedCount + 1

End Sub

'==============================================================
' 「一覧」シートに結果を書き出す
'==============================================================
Private Sub WriteResultSheet(ByRef results As Collection, ByVal targetFolder As String)

    Dim ws As Worksheet
    Set ws = GetOrCreateSheet(OUTPUT_SHEET_NAME)
    ws.Cells.Clear

    ws.Cells(1, 1).Value = "対象フォルダ"
    ws.Cells(1, 2).Value = targetFolder
    ws.Cells(1, 1).Font.Bold = True

    ws.Cells(2, 1).Value = "実行日時"
    ws.Cells(2, 2).Value = Format(Now, "yyyy/mm/dd hh:nn:ss")
    ws.Cells(2, 1).Font.Bold = True

    Dim headerRow As Long
    headerRow = 4

    ws.Cells(headerRow, 1).Value = "ファイル名"
    ws.Cells(headerRow, 2).Value = "フォルダパス"
    ws.Cells(headerRow, 3).Value = "サイズ（バイト）"
    ws.Cells(headerRow, 4).Value = "更新日時"
    ws.Cells(headerRow, 5).Value = "作成日時"
    ws.Rows(headerRow).Font.Bold = True
    ws.Rows(headerRow).Interior.Color = RGB(221, 235, 247)

    Dim outRow As Long
    outRow = headerRow + 1

    Dim i As Long
    Dim rec As Variant
    For i = 1 To results.Count
        rec = results(i)
        ws.Cells(outRow, 1).Value = rec(0)
        ws.Cells(outRow, 2).Value = rec(1)
        ws.Cells(outRow, 3).Value = rec(2)
        ws.Cells(outRow, 4).Value = rec(3)
        ws.Cells(outRow, 5).Value = rec(4)
        outRow = outRow + 1
    Next i

    Dim lastRow As Long
    lastRow = outRow - 1

    If lastRow >= headerRow + 1 Then
        ws.Range(ws.Cells(headerRow + 1, 3), ws.Cells(lastRow, 3)).NumberFormat = "#,##0"
        ws.Range(ws.Cells(headerRow + 1, 4), ws.Cells(lastRow, 5)).NumberFormat = "yyyy/mm/dd hh:mm:ss"
    End If

    ws.Columns.AutoFit
    ws.Activate
    ws.Range("A1").Select

End Sub

'==============================================================
' シートが存在すれば取得、なければ新規作成して返す
'==============================================================
Private Function GetOrCreateSheet(ByVal sheetName As String) As Worksheet

    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = sheetName
    End If

    Set GetOrCreateSheet = ws

End Function
