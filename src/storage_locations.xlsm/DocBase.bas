Attribute VB_Name = "DocBase"
Option Explicit
'VBAを好まないためRustの命名規則を適用した
Const MAIN_SHEET As String = "保管検索"
Const TARGET_SHEET As String = "備品管理一覧"
Const PURPOSE_CELL As String = "B2"
Const MSG_CELL As String = "C3"

'検索ユースケース
Private Sub search_usecase()
    '検索文字列
    Dim purpose As Range: Set purpose = Worksheets(MAIN_SHEET).Range(PURPOSE_CELL)
    '検索結果
    Dim results_dto() As Integer: ReDim results_dto(0)
    results_dto(0) = 0

    '結果をレンダリング
    presenter doc_dao(purpose, results_dto())
    '初期化
    Erase results_dto
End Sub

'文書保管DAO
Private Function doc_dao(ByVal purpose As String, ByRef results_dto() As Integer) As Integer()
    '検索対象シート(マスタ)
    Dim tar As Worksheet: Set tar = Worksheets(TARGET_SHEET)
    Dim i As Integer

    '検索ロジック
    For i = 1 To tar.Cells(Rows.count, "A").End(xlUp).Row
        If tar.Cells(i, "O") Like "*" & purpose & "*" Then
            ReDim Preserve results_dto(UBound(results_dto) + 1)
            results_dto(UBound(results_dto)) = i
        End If
    Next

    doc_dao = results_dto()

End Function

'表示
Private Function presenter(ByRef results_dto() As Integer)
    '出力文字列
    Dim msg As Range: Set msg = Worksheets(MAIN_SHEET).Range(MSG_CELL)
    Dim key As Variant
    Dim tar As Worksheet: Set tar = Worksheets(TARGET_SHEET)
    Dim count As Integer: count = UBound(results_dto)

    'シートの保護解除
    Worksheets(MAIN_SHEET).UnProtect "tyco"

    For Each key In results_dto
        '
    Next key

    If count > 0 Then
        msg = count & "件見つかりました．"
    Else
        msg = "見つかりませんでした．"
    End If

    'シートの保護有効化
    Worksheets(MAIN_SHEET).Protect "tyco"

End Function


