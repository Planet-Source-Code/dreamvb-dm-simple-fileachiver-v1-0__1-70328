Attribute VB_Name = "LVTools"
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Const LVM_FIRST = &H1000
Private Const LVM_SETCOLUMNWIDTH As Long = (LVM_FIRST + 30)

Public Sub LVAddFileInfo(LstView32 As ListView, lzFileName As String, lzFileSize As Long)
    With LstView32
        .ListItems.Add , , lzFileName, 1, 1
        .ListItems(.ListItems.Count).SubItems(1) = GetFileType(GetFilePart(lzFileName, fFileExt))
        .ListItems(.ListItems.Count).SubItems(2) = Format(lzFileSize, "#,###,###,##0")
    End With
End Sub

Public Sub LVResizeColumn(LstView32 As ListView, Index As Integer)
    'Used to auto size Column headers
    SendMessage LstView32.hwnd, LVM_SETCOLUMNWIDTH, Index, ByVal (-2)
End Sub

Public Sub LVSelectSelectAll(LstView32 As ListView)
Dim lItem As ListItem
    'Selects all Items in Listview control.
    For Each lItem In LstView32.ListItems
        lItem.Selected = True
    Next lItem
    
    Set lItem = Nothing
End Sub

Public Function LVHasSelectedItems(LstView32 As ListView) As Boolean
Dim lItem As ListItem
Dim Idx As Boolean
    
    'Checks to see if any items are selected.
    For Each lItem In LstView32.ListItems
        If (lItem.Selected) Then
            Idx = True
            Exit For
        End If
    Next lItem
    
    LVHasSelectedItems = Idx
End Function

