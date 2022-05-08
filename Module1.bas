Sub ExportToEmail()
  ' David Poole 2022/05/09
  ' https://github.com/davidlpoole/Send-Email-Excel-VBA

On Error GoTo errcatch

Application.ScreenUpdating = False
Application.DisplayAlerts = False

    Dim wbMain As Workbook, wbExport As Workbook
    Dim strDir As String, strFile As String, strFileExt As String, strExportFile As String
    
    Set wbMain = ActiveWorkbook
    
    ' get the 'latest sales date' from 'DateTo' in 'days' (Cell 019?)
    strLatestDate = Format(Range(ActiveWorkbook.Names("DateTo").RefersTo).Value, "yyyy mm dd")
    
    ' get the current path and create the file name to export to
    strDir = wbMain.Path
    strFile = "\2023POULTRYSALESBUDGET vs ACTUAL " & strLatestDate
    strFileExt = ".xlsx"
    
    strExportFile = strDir & strFile & strFileExt
  
    
    'set up the details to send email
    Dim strMailTo As String, strMailCC As String, strMailSubject As String, strMailBody As String
    Dim c As Range
  
    'loop through the MailTo table in 'days' sheet
    For Each c In Range("MailTo")
        If c <> "" Then strMailTo = strMailTo & c.Value & "; "
        Debug.Print (c)
    Next
    
    'loop through the MailCC table in 'days' sheet
    For Each c In Range("MailCC")
        If c <> "" Then strMailCC = strMailCC & c.Value & "; "
    Next
    
    'get the subject in the 'days' sheet
    strMailSubject = Range("MailSubject")
    
    'set up the mail body text (HTML)
    strMailBody = _
            "<font face=""Calibri"" size=""12px"">" & _
            "Hi All,<br>" & _
            "<br>" & _
            "Please see the sales update attached.<br>" & _
            "<br>" & _
            "Cheers" & _
            "</font>"

    'create the attachment
    Sheets(Array("Channels", "Islands", "Food Serv NI", "Food Serv SI", "Retail NI", _
        "Retail SI")).Select
    Sheets("Retail SI").Activate
    Sheets(Array("Channels", "Islands", "Food Serv NI", "Food Serv SI", "Retail NI", _
        "Retail SI")).Copy
    Sheets("Channels").Select
    Sheets("Channels").Activate
    
    Set wbExport = ActiveWorkbook

    ' break links to remove links to main sheet
    wbExport.BreakLink Name:= _
        wbMain.FullName _
        , Type:=xlExcelLinks
    
    'remove any named ranges that reference the main sheet
    wbExport.Names("DateTo").Delete
    wbExport.Names("MtdHead").Delete
    wbExport.Names("MtdPct").Delete
    wbExport.Names("SelMth").Delete
    wbExport.Names("YtdHead").Delete
    
    ' repeat the remove links command now that the named ranges are deleted
    wbExport.BreakLink Name:= _
        wbMain.FullName _
        , Type:=xlExcelLinks
        
        
    'save the attachment and close it
    ChDir strDir
    wbExport.SaveAs Filename:= _
        strExportFile _
        , FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    wbExport.Close
    
    'send the email
    Call SendMail(strMailTo, strMailCC, strMailSubject, strExportFile, strMailBody)
    
    'save mkain file
    wbMain.Save

    'restore normal application
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    'close Excel
    Application.Quit
    
    Exit Sub
    
errcatch:
  Application.ScreenUpdating = True
  Application.DisplayAlerts = True
  MsgBox ("Uh oh, the macro failed, sorry!")
     
End Sub


Public Function SendMail(strMailTo As String, strMailCC As String, strMailSubject As String, StrAttachment As String, strMailBody As String)

    Dim OutApp As Object
    Dim OutMail As Object

    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)

    With OutMail
        .To = strMailTo
        .CC = strMailCC
        .BCC = ""
        .Subject = strMailSubject
        .HTMLBody = strMailBody
        .Attachments.Add StrAttachment
        .Send   'or use .Display to view email before sending
    End With

    Set OutMail = Nothing
    Set OutApp = Nothing
  
End Function
