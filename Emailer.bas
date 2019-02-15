Attribute VB_Name = "Emailer"
Option Explicit

Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)

Sub Emailer()
    Application.ScreenUpdating = False
    Dim t As Double: t = Timer()
    
    Dim WS As Worksheet: Set WS = ThisWorkbook.Sheets("Emailer")
    Dim r As Long: r = WS.Cells(Rows.Count, 1).End(xlUp).Row
    
    Dim outlookApp As Object: Set outlookApp = CreateObject("Outlook.Application")
    Dim outlookMail As Object
    
    Dim body_text As String
    body_text = "Hello," & vbNewLine & vbNewLine & _
    "Please find attached the tremendous report entitled 'Big Report - Grand rapport.pdf'." & vbNewLine & vbNewLine & _
    "Sincerely," & vbNewLine & vbNewLine & _
    "Sam Louden" & vbNewLine
    
    Dim i As Long
    For i = 2 To r
        '0 for olMailItem
        Set outlookMail = outlookApp.CreateItem(0)
        With outlookMail
            .To = WS.Range("A" & i).Value2
            .Subject = "Report"
            '1 for olFormatPlain
            .BodyFormat = 1
            .Body = body_text
            .Attachments.Add ThisWorkbook.Path & "\" & "Big Report - Grand rapport.pdf"
            '.Send to send immediately; .Save to save to Drafts
            .Save
        End With
        Sleep (100)
    Next
    
    Set outlookMail = Nothing
    Set outlookApp = Nothing
    
    Debug.Print Timer() - t
    Application.ScreenUpdating = True
    MsgBox "Done!"
End Sub
