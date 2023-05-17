Attribute VB_Name = "Module1"
Sub test()
Call table_filtering_val("B2", 1)
End Sub
Sub table_filtering_val(ByVal str As String, ByVal fld As Long)
'Sub table_filtering_val(str As String, fld As Long)
    Dim filterValues As Variant
    filterValues = Array("ISO")
    Range(str).AutoFilter Field:=fld, Criteria1:=filterValues
End Sub
Sub table_filtering_val_button(ByVal str As String, ByVal fld As Long)
'Sub table_filtering_val(str As String, fld As Long)
    Dim filterValues As Variant
    filterValues = Array("ISO")
    Range(str).AutoFilter Field:=fld, Criteria1:=filterValues
    MsgBox "‘æ1ˆø”F" & str & ",@‘æ2ˆø”F" & fld
End Sub
Sub release_filtering()
With Worksheets("Sheet1")
    If .AutoFilterMode Then
        .AutoFilterMode = False
    End If
End With
    Range("B2").Select
End Sub
