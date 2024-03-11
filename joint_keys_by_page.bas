Attribute VB_Name = "joint_keys_by_page"
    
Sub ConsolidateRowsUniqueValuesAndSaveAsCSV()
    Dim srcSheet As Worksheet
    Dim destSheet As Worksheet
    Dim lastRow As Long, i As Long
    Dim page_num As Variant
    Dim dict As Object, info As Object
    Dim outputPath As String
    Dim baseFileName As String
    Dim csvFileName As String
    
    Set dict = CreateObject("Scripting.Dictionary")
    Set srcSheet = ThisWorkbook.Sheets("original")
    Set destSheet = ThisWorkbook.Sheets.Add
    destSheet.Name = "converted"
    
    lastRow = srcSheet.Cells(srcSheet.Rows.Count, "A").End(xlUp).Row
    
    ' データ加工ロジック...
     For i = 2 To lastRow
        page_num = srcSheet.Cells(i, "AP").Value
        If Not dict.Exists(page_num) Then
            Set info = CreateObject("Scripting.Dictionary")
            ' 初期値として各列からのデータを設定
            info("作品名") = srcSheet.Cells(i, "AH").Value
            info("副題") = srcSheet.Cells(i, "AJ").Value
            info("作者名") = srcSheet.Cells(i, "AL").Value
            info("成立年代") = srcSheet.Cells(i, "AI").Value
            info("unidic") = "" ' 空の値
            info("古文") = srcSheet.Cells(i, "L").Value
            info("現代文") = "" ' 空の値
            dict.Add page_num, info
        Else
            ' "古文"の列のみ値を結合
            dict(page_num)("古文") = dict(page_num)("古文") & srcSheet.Cells(i, "L").Value
        End If
    Next i
    
    ' ヘッダー出力
    With destSheet
        .Cells(1, 1).Value = "作品名"
        .Cells(1, 2).Value = "副題"
        .Cells(1, 3).Value = "作者名"
        .Cells(1, 4).Value = "成立年代"
        .Cells(1, 5).Value = "unidic"
        .Cells(1, 6).Value = "古文"
        .Cells(1, 7).Value = "現代文"
        .Cells(1, 8).Value = "page_num"
    End With
    
    ' データ出力
    i = 2
    For Each page_num In dict.Keys
        With destSheet
            .Cells(i, 1).Value = dict(page_num)("作品名")
            .Cells(i, 2).Value = dict(page_num)("副題")
            .Cells(i, 3).Value = dict(page_num)("作者名")
            .Cells(i, 4).Value = dict(page_num)("成立年代")
            .Cells(i, 5).Value = dict(page_num)("unidic")
            .Cells(i, 6).Value = dict(page_num)("古文")
            .Cells(i, 7).Value = dict(page_num)("現代文")
            .Cells(i, 8).Value = page_num
        End With
        i = i + 1
    Next page_num
    
    ' 元のExcelファイル名（拡張子なし）を取得
    baseFileName = Left(ThisWorkbook.Name, InStrRev(ThisWorkbook.Name, ".") - 1)
    
    ' CSVファイル名の生成（元のファイル名に "_jointed.csv" を追加）
    csvFileName = baseFileName & ".csv"
    
    ' 出力パスを設定（Excelファイルと同じディレクトリ）
    outputPath = ThisWorkbook.Path & "\outputs\" & csvFileName
    
    ' 一時的に作成したシートをCSVファイルとして保存
    destSheet.SaveAs Filename:=outputPath, FileFormat:=xlCSV, Local:=True
    
    ' 一時シートを削除（ユーザーに確認なしで）
    Application.DisplayAlerts = False
    destSheet.Delete
    Application.DisplayAlerts = True
    
    MsgBox "CSVファイルが保存されました: " & outputPath
End Sub
