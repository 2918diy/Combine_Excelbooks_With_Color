VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   5655
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   11058
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
   
Dim fs, f, f1, fc, s, x, rowss, columnss

  '** 使用FileDialog对象来选择文件夹
Dim fd As FileDialog
Dim strPath As String
       
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

ListBox1.AddItem (f1.Path)

Next

End Sub

Private Sub CommandButton2_Click()

Set excelObj = CreateObject("Excel.Application")

excelObj.Visible = True

Set myExcel = excelObj.Workbooks.Open(ListBox1.Text)

For Each sht In myExcel.Worksheets

    ListBox2.AddItem (sht.Name)

Next

myExcel.Close

excelObj.Quit

Set excelObj = Nothing

End Sub

Private Sub CommandButton3_Click()

Set myRange = Application.InputBox(prompt:="选择颜色区域", Type:=8)

Dim str As String

For Each Rng In myRange

    str = Rng.Value & ":" & Rng.Interior.ColorIndex
    
    ListBox3.AddItem (str)

Next

End Sub

Private Sub CommandButton4_Click()

Set excelObj = CreateObject("Excel.Application")

excelObj.Visible = True

Set myExcel = excelObj.Workbooks.Open(ListBox1.Text)

With Me.ListBox2
    
    For i = 0 To .ListCount - 1
        
        If .Selected(i) Then
            
            sheet_name = .List(i)
            
            myExcel.Worksheets(sheet_name).UsedRange.Select
            
            For Each Rng In myExcel.Worksheets(sheet_name).UsedRange
            
                    On Error Resume Next
            
                    If Rng.Interior.ColorIndex = CStr(Right(ListBox3.Text, Len(ListBox3.Text) - InStr(ListBox3.Text, ":"))) Then
                        
                        ThisWorkbook.Worksheets(sheet_name).Range(Replace(Rng.Address, "$", "")).Value = Rng.Value
                        
                    End If
            
            Next
            
            ListBox4.AddItem (sheet_name)
            
        End If
    
    Next

End With

End Sub
