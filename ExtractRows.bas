Attribute VB_Name = "Module1"
Sub Extract()

' Macro code by
'===================================================
'~~~~~~~~~~~~~~~~~~~~ILKER ICYÜZ~~~~~~~~~~~~~~~~~~~~
'===================================================

'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Application.Calculation = xlCalculationManual       '!!
Application.ScreenUpdating = False                  '!!
Application.EnableEvents = False                    '!!
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
On Error GoTo Errorhandler

ExtractUserForm.Show


Errorhandler:

'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Application.Calculation = xlCalculationAutomatic    '!!
Application.ScreenUpdating = True                   '!!
Application.EnableEvents = True                     '!!
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

End Sub


