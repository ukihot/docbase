Attribute VB_Name = "DocBase"
Option Explicit
'VBA���D�܂Ȃ�����Rust�̖����K����K�p����
Const MAIN_SHEET As String = "�ۊǌ���"
Const TARGET_SHEET As String = "���i�Ǘ��ꗗ"
Const PURPOSE_CELL As String = "B2"
Const MSG_CELL As String = "C3"

'�������[�X�P�[�X
Private Sub search_usecase()
    '����������
    Dim purpose As Range: Set purpose = Worksheets(MAIN_SHEET).Range(PURPOSE_CELL)
    '��������
    Dim results_dto() As Integer: ReDim results_dto(0)

    '���ʂ������_�����O
    presenter doc_dao(purpose, results_dto())
    '������
    Erase results_dto
End Sub

'�����ۊ�DAO
Private Function doc_dao(ByVal purpose As String, ByRef results_dto() As Integer) As Integer()
    '�����ΏۃV�[�g(�}�X�^)
    Dim tar As Worksheet: Set tar = Worksheets(TARGET_SHEET)
    Dim i As Integer
    Dim key As Variant
    '�������W�b�N
    For i = 1 To tar.Cells(Rows.count, "A").End(xlUp).Row
        'FIXME: �v���ǉ��̂��тɂ����������Ȃ�
        If tar.Cells(i, "G") Like "*" & purpose & "*" Or tar.Cells(i, "O") Like "*" & purpose Then
            Dim is_registered As Boolean: is_registered = False
            'FIXME: �����ɏ�����DAO���h���C���m�������Ă�݂����ɂȂ邽�߃T�[�r�X��������
            For Each key In results_dto
                If key = i Then
                    '�o�^���Ȃ�
                    is_registered = True
                End If
            Next
            If Not is_registered Then
                ReDim Preserve results_dto(UBound(results_dto) + 1)
                results_dto(UBound(results_dto)) = i
            End If
        End If
    Next

    doc_dao = results_dto()

End Function

'�\��
Private Function presenter(ByRef results_dto() As Integer)
    Dim main As Worksheet: Set main = Worksheets(MAIN_SHEET)
    Dim msg As Range: Set msg = main.Range(MSG_CELL)
    Dim key As Variant
    Dim tar As Worksheet: Set tar = Worksheets(TARGET_SHEET)
    Dim count As Integer: count = UBound(results_dto)
    Dim j As Integer: j = 1

    '�V�[�g�̕ی����
    main.Unprotect "tyco"

    '������
    Rows("6:6").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents

    '�ꗗ��
    For Each key In results_dto
        If key <> 0 Then
            main.Cells(4 + j * 2, "B") = tar.Cells(key, "B")
            main.Hyperlinks.Add Anchor:=main.Cells(4 + j * 2, "E"), Address:="", SubAddress:= _
        "���i�Ǘ��ꗗ!A" & key, TextToDisplay:=tar.Cells(key, "F").Value
            j = j + 1
        End If
    Next key

    If count > 0 Then
        msg = count & "��������܂����D"
    Else
        msg = "������܂���ł����D"
    End If

    '�Z���߂�
    Range("A1").Select

    '�V�[�g�̕ی�L����
    main.Protect "tyco"

End Function