Attribute VB_Name = "Module1"

Public Sub combine_sub()
Attribute combine_sub.VB_ProcData.VB_Invoke_Func = "t\n14"

Dim slctRange As Range
Dim shtName As String
Dim MyStr() As String
Dim MyDbl() As Double

Dim fs, f, f1, fc, s, x, rowss, columnss
    
'** 使用FileDialog对象来选择文件夹
Dim fd As FileDialog
Dim strPath As String

Set slctRange = Selection
shtName = ThisWorkbook.ActiveSheet.Name

slct_row = slctRange.Rows.Count
slect_col = slctRange.Columns.Count


vrf = slct_row * slect_col = slct_row Or slct_row * slect_col = slect_col

If vrf Then

    MsgBox vrf

    ReDim Preserve MyStr(slct_row * slect_col)
    ReDim Preserve MyDbl(slct_row * slect_col)
    i = 0
    For Each Rng In slctRange
            MyStr(i) = Replace(Rng.Address, "$", "")
            MyDbl(i) = 0
            i = i + 1
    Next
    
    Set ExcelApp = CreateObject("Excel.Application")
    
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    
            '** 显示选择文件夹对话框
    If fd.Show = -1 Then        '** 用户选择了文件夹
        f_path = fd.SelectedItems(1)
    Else
        f_path = ""
    End If
        Set fd = Nothing
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFolder(f_path) 'Directory of excel files will be merge
    Set fc = f.Files
    
    For Each f1 In fc
        Set ExcelBook = Workbooks.Open(Filename:=f1, ReadOnly:=True)
        ExcelBook.Sheets(shtName).Select
        For i = 0 To UBound(MyStr) - 1
            MyDbl(i) = MyDbl(i) + Val(ExcelBook.Sheets(shtName).Range(MyStr(i)).Value)
        Next i
    Next
    
    For i = 0 To UBound(MyDbl) - 1
        ThisWorkbook.ActiveSheet.Range(MyStr(i)).Value = MyDbl(i)
    Next i

Else

    ReDim Preserve MyStr(slect_col, slct_row)
    ReDim Preserve MyDbl(slect_col, slct_row)
    
    
    'For Each Rng In slctRange
        'MsgBox Rng.Address
      For i = 0 To UBound(MyStr, 2) - 1
          For j = 0 To UBound(MyStr, 1) - 1
              MyStr(i, j) = Replace(slctRange.Cells(i + 1, j + 1).Address, "$", "")
              MyDbl(i, j) = 0
          Next j
      Next i
    'Next
    
    Set ExcelApp = CreateObject("Excel.Application")
    
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    
          '** 显示选择文件夹对话框
    If fd.Show = -1 Then        '** 用户选择了文件夹
      f_path = fd.SelectedItems(1)
    Else
      f_path = ""
    End If
      Set fd = Nothing
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFolder(f_path) 'Directory of excel files will be merge
    Set fc = f.Files
    
    For Each f1 In fc
      Set ExcelBook = Workbooks.Open(Filename:=f1, ReadOnly:=True)
      ExcelBook.Sheets(shtName).Select
      For i = 0 To UBound(MyStr, 2) - 1
          For j = 0 To UBound(MyStr, 1) - 1
              MyDbl(i, j) = MyDbl(i, j) + ExcelBook.Sheets(shtName).Range(MyStr(i, j)).Value
          Next j
      Next i
    Next
    
    For i = 0 To UBound(MyStr, 2) - 1
      For j = 0 To UBound(MyStr, 1) - 1
          ThisWorkbook.ActiveSheet.Range(MyStr(i, j)).Value = MyDbl(i, j)
      Next j
    Next i
End If

End Sub

