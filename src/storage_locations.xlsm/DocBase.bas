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
    results_dto(0) = 0

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

    '�������W�b�N
    For i = 1 To tar.Cells(Rows.count, "A").End(xlUp).Row
        If tar.Cells(i, "O") Like "*" & purpose & "*" Then
            ReDim Preserve results_dto(UBound(results_dto) + 1)
            results_dto(UBound(results_dto)) = i
        End If
    Next

    doc_dao = results_dto()

End Function

'�\��
Private Function presenter(ByRef results_dto() As Integer)
    '�o�͕�����
    Dim msg As Range: Set msg = Worksheets(MAIN_SHEET).Range(MSG_CELL)
    Dim key As Variant
    Dim tar As Worksheet: Set tar = Worksheets(TARGET_SHEET)
    Dim count As Integer: count = UBound(results_dto)

    '�V�[�g�̕ی����
    Worksheets(MAIN_SHEET).UnProtect "tyco"

    For Each key In results_dto
        '
    Next key

    If count > 0 Then
        msg = count & "��������܂����D"
    Else
        msg = "������܂���ł����D"
    End If

    '�V�[�g�̕ی�L����
    Worksheets(MAIN_SHEET).Protect "tyco"

End Function


