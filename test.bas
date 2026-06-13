' @Description: シート存在チェック
' @Param: (String) SheetName - 第1引数
' @Return: (Boolean) 真偽値
' @Constraint: 空ではないシート名文字列を受ける
Function SheetExists(SheetName As String) As Boolean
' シート存在チェック
    Dim ws As Worksheet
    SheetExists = False
    For Each ws In Worksheets
        If ws.Name = SheetName Then
            SheetExists = True
            Exit For
        End If
    Next ws
End Function
