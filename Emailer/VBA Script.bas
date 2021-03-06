Attribute VB_Name = "Emailer"
Option Explicit

'Import the Sleep function
'Used to add small pause between each email sent to avoid
'overwhelming the email service
Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)

Sub Emailer()
    'Disable ScreenUpdating for the duration of the script to
    'prevent screen flicker (also mildly increases speed)
    Application.ScreenUpdating = False
    
    'Start timer
    Dim t As Double: t = Timer()
    
    'Declare a variable to refer to the Emailer worksheet
    Dim WS As Worksheet: Set WS = ThisWorkbook.Sheets("Emailer")
    'Get the number of rows
    Dim r As Long: r = WS.Cells(Rows.Count, 1).End(xlUp).Row
    
    'Create an instance of Outlook
    Dim outlookApp As Object: Set outlookApp = CreateObject("Outlook.Application")
    Dim outlookMail As Object
    
    'Store email text in variable, using vbNewLine for Enter
    Dim body_text As String
    body_text = "Hello," & vbNewLine & vbNewLine & _
    "Please find attached the tremendous report entitled 'Big Report - Grand rapport.pdf'." & vbNewLine & vbNewLine & _
    "Note the code from today's presentation can be found at:" & vbNewLine & vbNewLine & _
    "https://github.com/deepklim/Data-Analytics-in-Financial-Management-GC" & vbNewLine & vbNewLine & _
    "Sincerely," & vbNewLine & vbNewLine & _
    "Sam Louden" & vbNewLine
    
    'Loop through each email address in column A and send it attachment in column B + message
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
            'Add attachments; place the Reports folder in same folder as this program
            .Attachments.Add ThisWorkbook.Path & "\Reports\" & WS.Range("B" & i).Value2
            '.Send to send immediately; .Save to save to Drafts
            .Send
        End With
        'Pause for 100 miliseconds after sending each email
        Sleep (100)
    Next
    
    'Close instance of Outlook
    Set outlookMail = Nothing
    Set outlookApp = Nothing
    
    'Print execution time in console and tell user the script has finished
    Debug.Print Timer() - t
    MsgBox "Done!"
    
    'Re-enable ScreenUpdating
    Application.ScreenUpdating = True
End Sub
