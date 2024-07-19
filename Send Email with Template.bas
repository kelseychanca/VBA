Attribute VB_Name = "SendEmail"
Option Explicit

Sub SendEmail()

    Dim OutApp As Object
    Dim OutMail As Object
    Dim sh As Worksheet
    Dim cell As Range
    Dim FileCell As Range
    Dim rng As Range
    Dim nRow As Integer
    Dim account, FilePath, Att1, Att2, Att3 As String
    Dim Answer As Long
    
    ' Disable events and screen updating for better performance
    With Application
        .EnableEvents = False
        .ScreenUpdating = False
    End With
    
    ' Create Outlook application object
    Set OutApp = CreateObject("Outlook.Application")
    nRow = 0
    
    ' Read account and file path from worksheet
    account = Range("B1").Value
    FilePath = "C:\Users\" & account & "\xxxxxxx\"
    
    ' Loop through selected sheets in the active window
    For Each sh In ActiveWindow.SelectedSheets
    
        ' Prompt the user to confirm if they want to run the code for the current sheet
        Answer = MsgBox("This is sheet : " & sh.Name & Chr(10) & Chr(10) & "Do you want to run?", vbExclamation + vbYesNo, "Reply")
        
        If Answer = vbYes Then
        
            ' Loop through cells in column B of the current sheet
            For Each cell In sh.Columns("B").Cells.SpecialCells(xlCellTypeConstants)
            
                ' Check if the cell value is a valid email address and if there are attachments specified in columns F:H
                Set rng = sh.Cells(cell.Row, 1).Range("A1:H1")
                
                If cell.Value Like "?*@?*.?*" And Application.WorksheetFunction.CountA(rng) > 0 Then
                    
                    ' Create a new email item
                    Set OutMail = OutApp.CreateItem(0)
                    
                    With OutMail
                        ' Set the recipient, CC, subject, and body of the email
                        .TO = sh.Cells(cell.Row, 2).Value
                        .CC = sh.Cells(cell.Row, 3).Value
                        .Subject = sh.Cells(cell.Row, 4).Value
                        .Body = sh.Cells(cell.Row, 5).Value
                        
                        ' Set the paths of the attachments specified in columns F:H
                        Att1 = FilePath & sh.Cells(cell.Row, 6)
                        Att2 = FilePath & sh.Cells(cell.Row, 7)
                        Att3 = FilePath & sh.Cells(cell.Row, 8)
                        
                        ' Attach the files if they exist
                        If sh.Cells(cell.Row, 6) <> "" Then
                            .Attachments.Add Att1
                            If sh.Cells(cell.Row, 7) <> "" Then
                                .Attachments.Add Att2
                                If sh.Cells(cell.Row, 8) <> "" Then
                                    .Attachments.Add Att3
                                End If
                            End If
                        End If
                        
                        sh.Cells(cell.Row, 1).Select
                        Selection.Interior.Color = RGB(204, 204, 255)
                        
                        nRow = nRow + 1
                        
                        '.Send ' Uncomment this line to send the email immediately
                        .Display ' Comment out this line if you uncomment the line above
                        
                    End With
                    
                    Set OutMail = Nothing
                    
                End If
                
            Next cell
            
            ' Release the Outlook application object
            Set OutApp = Nothing
            
            ' Enable events and screen updating
            With Application
                .EnableEvents = True
                .ScreenUpdating = True
            End With
            
            ' Display a message box with the total number of emails sent
            MsgBox ("Finish!" & Chr(10) & Chr(10) & "Total Number of Email : " & nRow)
            
        End If
        
    Next sh

End Sub