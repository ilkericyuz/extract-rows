VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ExtractUserForm 
   Caption         =   "Please Enter the Criteria"
   ClientHeight    =   1515
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   2760
   OleObjectBlob   =   "ExtractUserForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ExtractUserForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()

' Macro code by
'===================================================
'~~~~~~~~~~~~~~~~~~~~ILKER ICYÜZ~~~~~~~~~~~~~~~~~~~~
'===================================================


    'Empty Customer Number
    CriteriaTextBox.Value = ""
    
End Sub


Private Sub CancelButton_Click()

' Macro code by
'===================================================
'~~~~~~~~~~~~~~~~~~~~ILKER ICYÜZ~~~~~~~~~~~~~~~~~~~~
'===================================================

Unload Me


End Sub


Private Sub OKButton_Click()

' Macro code by
'===================================================
'~~~~~~~~~~~~~~~~~~~~ILKER ICYÜZ~~~~~~~~~~~~~~~~~~~~
'===================================================

    On Error GoTo Errorhandler
    Unload Me

    Dim Wbname As String
    Dim OtherWorkbook As Workbook
    Dim ws1 As Worksheet
    Dim lngCalc As Long
    Dim lngrow As Long
    Dim pathstring As String
    Dim flag As Integer
    Dim ColumnCount As Long
'    Dim Password As String

'    Password = "CRM"

    
    Application.DisplayAlerts = False
    
    '===================== Get the folder path =====================

    ' Open the file dialog
    Set diaFolder = Application.FileDialog(msoFileDialogFolderPicker)
    diaFolder.AllowMultiSelect = False
    MsgBox ("Please select the folder")
    diaFolder.Title = "Please select the folder"
    diaFolder.Show
    FolderName = diaFolder.SelectedItems(1)
    Set diaFolder = Nothing
    '----------------------------------------------------------------
 

    Wbname = Dir(FolderName & "\" & "*.xls*")

    'ThisWorkbook.Sheets(1).UsedRange.ClearContents
    flag = 0
    
    Do While Len(Wbname) > 0
        Set OtherWorkbook = Workbooks.Open(FolderName & "\" & Wbname)
        On Error Resume Next
        
        Set ws = OtherWorkbook.Sheets(1)
        
'        OtherWorkbook.ActiveSheet.Unprotect Password
        
        'SEARCH FOR THE CRITERIA
        On Error GoTo NotFound
        OtherWorkbook.ActiveSheet.Cells.Find(What:=CriteriaTextBox.Value, After:=ActiveCell, LookIn:=xlFormulas, _
            LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
            MatchCase:=False, SearchFormat:=False).Activate
        
        
        j = ActiveCell.Row
        i = 0
        
        'CHECK IF SEARCH COMPLETED ONE TOUR
        Do While i <> j
            Cells.FindNext(After:=ActiveCell).Activate
            i = ActiveCell.Row
            
            'COPY A TO CC
            Range("A" & i & ":EE" & i).Copy

            'PASTE FROM COLUMN B
            ThisWorkbook.Sheets(1).Range("B" & ThisWorkbook.Sheets(1).Range("A65536").End(xlUp)(2).Row).PasteSpecial xlPasteValues
            'PASTE COLUMN A THE NAME OF THE FILE
            ThisWorkbook.Sheets(1).Range("A" & ThisWorkbook.Sheets(1).Range("A65536").End(xlUp)(2).Row) = Wbname
            
            'COPY COLUMN NAMES INTO THE FIRST ROW ONE TIME
            If flag = 0 Then
            
                Rows(1).SpecialCells(xlCellTypeConstants).Copy
                ThisWorkbook.Sheets(1).Range("B1").PasteSpecial xlPasteValues

                flag = 1
            End If
            
        Loop
            
        
            
NotFound:
Resume NotFound2

NotFound2:
        OtherWorkbook.Close False
        Wbname = Dir
           
    Loop
    
    
    'INSERT A BLANK ROW ON TOP
    Rows("1:1").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove

    'REPOSITION THE BUTTON
    ActiveSheet.Shapes("Button 1").IncrementTop -15
    Range("A2").Select



    Application.DisplayAlerts = True
     
    
    
    Exit Sub
    
Errorhandler:
    
    MsgBox ("ERROR! An error occured!")
    
    
End Sub



 


