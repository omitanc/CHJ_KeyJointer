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
    
    ' �f�[�^���H���W�b�N...
     For i = 2 To lastRow
        page_num = srcSheet.Cells(i, "AP").Value
        If Not dict.Exists(page_num) Then
            Set info = CreateObject("Scripting.Dictionary")
            ' �����l�Ƃ��Ċe�񂩂�̃f�[�^��ݒ�
            info("��i��") = srcSheet.Cells(i, "AH").Value
            info("����") = srcSheet.Cells(i, "AJ").Value
            info("��Җ�") = srcSheet.Cells(i, "AL").Value
            info("�����N��") = srcSheet.Cells(i, "AI").Value
            info("unidic") = "" ' ��̒l
            info("�Õ�") = srcSheet.Cells(i, "L").Value
            info("���㕶") = "" ' ��̒l
            dict.Add page_num, info
        Else
            ' "�Õ�"�̗�̂ݒl������
            dict(page_num)("�Õ�") = dict(page_num)("�Õ�") & srcSheet.Cells(i, "L").Value
        End If
    Next i
    
    ' �w�b�_�[�o��
    With destSheet
        .Cells(1, 1).Value = "��i��"
        .Cells(1, 2).Value = "����"
        .Cells(1, 3).Value = "��Җ�"
        .Cells(1, 4).Value = "�����N��"
        .Cells(1, 5).Value = "unidic"
        .Cells(1, 6).Value = "�Õ�"
        .Cells(1, 7).Value = "���㕶"
        .Cells(1, 8).Value = "page_num"
    End With
    
    ' �f�[�^�o��
    i = 2
    For Each page_num In dict.Keys
        With destSheet
            .Cells(i, 1).Value = dict(page_num)("��i��")
            .Cells(i, 2).Value = dict(page_num)("����")
            .Cells(i, 3).Value = dict(page_num)("��Җ�")
            .Cells(i, 4).Value = dict(page_num)("�����N��")
            .Cells(i, 5).Value = dict(page_num)("unidic")
            .Cells(i, 6).Value = dict(page_num)("�Õ�")
            .Cells(i, 7).Value = dict(page_num)("���㕶")
            .Cells(i, 8).Value = page_num
        End With
        i = i + 1
    Next page_num
    
    ' ����Excel�t�@�C�����i�g���q�Ȃ��j���擾
    baseFileName = Left(ThisWorkbook.Name, InStrRev(ThisWorkbook.Name, ".") - 1)
    
    ' CSV�t�@�C�����̐����i���̃t�@�C������ "_jointed.csv" ��ǉ��j
    csvFileName = baseFileName & ".csv"
    
    ' �o�̓p�X��ݒ�iExcel�t�@�C���Ɠ����f�B���N�g���j
    outputPath = ThisWorkbook.Path & "\outputs\" & csvFileName
    
    ' �ꎞ�I�ɍ쐬�����V�[�g��CSV�t�@�C���Ƃ��ĕۑ�
    destSheet.SaveAs Filename:=outputPath, FileFormat:=xlCSV, Local:=True
    
    ' �ꎞ�V�[�g���폜�i���[�U�[�Ɋm�F�Ȃ��Łj
    Application.DisplayAlerts = False
    destSheet.Delete
    Application.DisplayAlerts = True
    
    MsgBox "CSV�t�@�C�����ۑ�����܂���: " & outputPath
End Sub
