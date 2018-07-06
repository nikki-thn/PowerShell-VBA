Private Sub CommandButton1_Click() 'CLEAR SHEETS Button

    
    Dim sheetsName(0 To 2) As String

    sheetsName(0) = "Current"
    sheetsName(1) = "Previous_Month"
    sheetsName(2) = "Previous_Quarter"
    
    For x = 0 To 2
                
        With Sheets(sheetsName(x))
            .Rows("2:" & Rows.Count).ClearContents 'Clear Content
        End With
        
    Next x

End Sub


Private Sub CommandButton4_Click() 'FILTER DATA
   
    With Sheets("COUNT")
        .Rows("7:" & Rows.Count).ClearContents 'Clear Content
    End With

    Dim sheetsName(0 To 2) As String

    sheetsName(0) = "Current"
    sheetsName(1) = "Previous_Month"
    sheetsName(2) = "Previous_Quarter"
    
    Dim lRowPROD As Long
    Dim lRowPRD1D As Long
    Dim x As String
    
    For y = 0 To 2
     
        'Find the last non-blank cell in worksheet PROD_Current, column A(1)
        lRowPROD = Worksheets(sheetsName(y)).Cells(Rows.Count, 7).End(xlUp).row
        lRowPRD1D = Worksheets(sheetsName(y)).Cells(Rows.Count, 1).End(xlUp).row
            
        If (y = 0) Then
            x = "A7"
        ElseIf (y = 1) Then
            x = "D7"
        Else
            x = "G7"
        End If
          
        With Sheets(sheetsName(y))
            .AutoFilterMode = False
            .range("D1:D" & CStr(lRowPRD1D)).AutoFilter Field:=1, Criteria1:="FALSE"
            .range("A1:B" & CStr(lRowPRD1D)).SpecialCells(xlCellTypeVisible).Copy Destination:=Sheets("COUNT").range(x)
            .AutoFilterMode = False
        End With
        
    On Error Resume Next
        Sheet3.ShowAllData
    On Error GoTo 0

    Next y


End Sub


Private Sub CommandButton2_Click() 'RUN SCRIPT Button

On Error GoTo ErrHandling
    Application.ScreenUpdating = False
    
    Dim Current As String
    Dim PreviousMonth As String
    Dim PreviousQuarter As String
    
    Current = Format(DateSerial(Year(Date), Month(Date) - 1, 0), "yyyy-mm-dd") ' Current Month End
    PreviousMonth = Format(DateSerial(Year(Date), Month(Date) - 1, 0), "yyyy-mm-dd") ' Previous Month End
    PreviousQuarter = Format(DateSerial(Year(Date), (((Month(Date) - 3) \ 3) * 3) - 1, 0), "yyyy-mm-dd") ' Previous Quarter End
            
    'To test for the dates end
    'MsgBox (Current & vbCrLf & PreviousMonth & vbCrLf & PreviousQuarter)
    'Exit Sub

    Dim msgStr As String
    msgStr = "Running on: " & vbCrLf & "Month End: " & Current & vbCrLf
    msgStr = msgStr & "Previous Month: " & PreviousMonth & vbCrLf
    msgStr = msgStr & "Previous Quarter: " & PreviousQuarter & vbCrLf
    msgStr = msgStr & "Please wait until the complete message shown"
    'MsgBox (msgStr)

    Dim dateStr As String 'To store set @asOfDate statement
    Dim SQLQuery As String 'To store SQL Query

    Dim sheetsName(0 To 2) As String
    Dim datesEnd(0 To 2) As String

    sheetsName(0) = "Current"
    sheetsName(1) = "Previous_Month"
    sheetsName(2) = "Previous_Quarter"
    datesEnd(0) = Current
    datesEnd(1) = PreviousMonth
    datesEnd(2) = PreviousQuarter
    
    Dim progress As Double 'Counter for progress bar
    progress = 0
              
    For x = 0 To 2

        'dateStr = "set @asOfDate = '" & datesEnd(x) & "'" & vbCrLf
        SQLQuery = ConcatSQLQuery(datesEnd(x))
        'MsgBox (SQLQuery)

'        With Sheets(sheetsName(x))
'            Rows("2:" & Rows.Count).ClearContents 'Clear Content
'        End With

        Call executeUATQuery(sheetsName(x), SQLQuery)

        progress = progress + 35
        Application.StatusBar = "Progress: " & Format(progress / 300, "Percent")
        DoEvents

        Call executePRODQuery(sheetsName(x), SQLQuery)

        progress = progress + 50
        Application.StatusBar = "Progress: " & Format(progress / 300, "Percent")
        DoEvents

    Next x
    
    For y = 0 To 2
                    
        Call CompareData(sheetsName(y))
        progress = progress + 15
        Application.StatusBar = "Progress: " & Format(progress / 300, "Percent")
        DoEvents
        
    Next y
          
Done:
    
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox ("Done!")
    Exit Sub
    
ErrHandling:

    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox ("Error:" & Err.Description)
       
End Sub

Private Sub CompareData(nameSheet As String)

    Dim lRowPROD As Long
    Dim lRowPRD1D As Long

    'Find the last non-blank cell in worksheet PROD_Current, column A(1)
    lRowPROD = Worksheets(nameSheet).Cells(Rows.Count, 7).End(xlUp).row
    lRowPRD1D = Worksheets(nameSheet).Cells(Rows.Count, 1).End(xlUp).row
    'MsgBox (lRowPRD1D & "  " & lRowPROD)

    With Sheets(nameSheet)
        .range("C2").Formula = "=VLOOKUP(RC[-2],R2C7:R" & CStr(lRowPROD) & "C8,2,FALSE)"
        .range("C2:C" & CStr(lRowPRD1D)).FillDown
        .range("D2").Formula = "=EXACT(RC[-2],RC[-1])"
        .range("D2:D" & CStr(lRowPRD1D)).FillDown
    End With
    
End Sub

Private Function ConcatSQLQuery(dateStr As String) As String

    Dim SQLStr1 As String 'To store first half of the store procedure
    Dim SQLStr2 As String 'To store second half of the store precedure

SQLStr1 = "sql query"

          
SQLStr2 = "sql query"
       
    'MsgBox (SQLStr1 & dateStr & SQLStr2)
    'Read in sql script
'    Dim hnd As Integer
'    hnd = FreeFile
'
'    FilePath = ActiveWorkbook.Path & "\PROD.sql" ' & "Sub-DirectoryName\"
'    'MsgBox (FilePath)
'    Open FilePath For Input As hnd
'
'    Dim row As String
'    Do Until EOF(hnd)
'    Line Input #hnd, row
'    SQLStr2 = SQLStr2 & row & vbCrLf
'    Loop
'
'    Close #hnd
    
    ConcatSQLQuery = SQLStr1 & dateStr & SQLStr2
    Application.ScreenUpdating = True
    
End Function

Private Sub executePRODQuery(nameSheet As String, SQLQuery As String)

    Dim Cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim PROD_Server_Name As String
    Dim PROD_Database_Name As String
    Dim User_ID As String
    Dim Password As String
    
    UAT_Server_Name = "server name" 'UAT server name here
    UAT_Database_Name = "database name"
    'User_ID = "user name" ' enter your user ID here
    'Password = "pw" ' Enter your password here
    
    'Exit Sub
    Set Cn = New ADODB.Connection
    Cn.Open "Driver={SQL Server};Server=" & PROD_Server_Name & ";Database=" & PROD_Database_Name & _
        ";Uid=" & User_ID & ";Pwd=" & Password & ";"
     
    rs.Open SQLQuery, Cn, adOpenStatic
  
     ' Paste data to spreadsheet
    With Worksheets(nameSheet).range("G2")  ' Enter your worksheet name here
        .CopyFromRecordset rs
    End With
    
    ' Clean up
    
    rs.Close
    Set rs = Nothing
    Cn.Close
    Set Cn = Nothing
    
    Application.ScreenUpdating = True

End Sub

Private Sub executeUATQuery(nameSheet As String, SQLQuery As String)

    Dim Cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim UAT_Server_Name As String
    Dim UAT_Database_Name As String
    'Dim User_ID As String
    'Dim Password As String
    
    UAT_Server_Name = "server name" 'UAT server name here
    UAT_Database_Name = "database name"
    'User_ID = "user name" ' enter your user ID here
    'Password = "pw" ' Enter your password here
    
    'Exit Sub
    Set Cn = New ADODB.Connection
    Cn.Open "Driver={SQL Server};Database=" & UAT_Database_Name & ";Server=" & UAT_Server_Name & ";"
     
    rs.Open SQLQuery, Cn, adOpenStatic

     ' Paste data to spreadsheet
    With Worksheets(nameSheet).range("A2") ' Enter your sheet name and range here
        .CopyFromRecordset rs
    End With
    
    ' Clean up
    rs.Close
    Set rs = Nothing
    Cn.Close
    Set Cn = Nothing
       
    Application.ScreenUpdating = True

End Sub




'Credits:
'https://social.msdn.microsoft.com/Forums/en-US/f599f280-b786-43f4-9357-77f19a107c3f/vba-in-excel-connect-to-sql-server
'https://stackoverflow.com/questions/27043116/running-sql-file-from-vba
'https://msdn.microsoft.com/en-us/vba/excel-vba/articles/application-statusbar-property-excel
'https://stackoverflow.com/questions/5181164/progress-bar-in-vba-excel
'http://www.endprod.com/colors/
'https://stackoverflow.com/questions/13955708/excel-filtering-and-copying-in-vba
'https://stackoverflow.com/questions/30099311/how-to-paste-multiple-range-of-cells-in-email-body-using-excel-vba
'https://stackoverflow.com/questions/10849177/easiest-way-to-loop-through-a-filtered-list-with-vba


Private Sub CommandButton5_Click()

'SEND EMAIL Button

On Error GoTo ErrHandler

    Dim OutlookApp As Object
    Dim MItem As Object
    Dim recipients As String
    Dim ccRecipients As String
    Dim bodyMsg As String
    
    On Error GoTo ErrHandler:
 
    'Create Outlook object
    Set OutlookApp = CreateObject("Outlook.Application")
    
    'Create Mail Item and view before sending
    Set MItem = OutlookApp.CreateItem(olMailItem)
    
    'Set recipients
    recipients = "
    
    'Set CC recipients
    ccRecipients = ""


   'Set email bosy message
    bodyMsg = "<span style='font-size: small'>Hi Aru, <br><br>Below is the " _
        & "results of today's data. Thank you!<br><br><br>" & generateHTMLTable()
    
    'Set Outlook email variables
    With MItem
        .Display 'this line is crucial to insert outlook signature to the email
        .To = recipients 'add recipients
        .cc = ccRecipients 'add cc recipients
        'Set email subject and current date
        .Subject = "Portia ID Check - as of " & Date
        
        'put everything together
        .HTMLBody = bodyMsg & .HTMLBody 'the .HTMLBody is a must to insert default signature to the email
        
        'display the email
        .Display
              
    End With

    ' CLEAR
    Set objEmail = Nothing
    Set objOutlook = Nothing
        
ErrHandler:

    ' When error occurs, the code stops and displays the error.
    On Error GoTo 0

End Sub

Private Function RangeToHTML(rng As range, noOfColumn As Integer) As String

    Dim htmlStr As String
    Dim x As Integer
    x = 1

    For Each Cell In rng1

        If x = 1 Then
            htmlStr = htmlStr & "<tr><td>" & CStr(Cell.Value) & "</td>"
            x = x + 1
        ElseIf x = noOfColumn Then
            htmlStr = htmlStr & "<td>" & CStr(Cell.Value) & "</td></tr>"
            x = 1
        Else
            htmlStr = htmlStr & "<td>" & CStr(Cell.Value) & "</td>"
            x = x + 1
        End If

    Next Cell

    RangeToHTML = htmlStr

End Function

Private Function generateHTMLTable() As String

    Dim rng1 As range
    Dim rng2 As range
    Dim htmlStr As String
    htmlStr = ""
    
    Set rng1 = Sheets("COUNT").range("A2:D2", range("A2:D2").End(xlDown)).Cells.SpecialCells(xlCellTypeVisible)
    Set rng2 = Sheets("COUNT").range("A8:B8", range("A8:B8").End(xlDown))
    
    htmlStr = "<table style='width:200pt'><tr><th style='background-color:#004d00;color:white'>COUNT</th>" _
        & "<th style='background-color:#004d00;color:white'>Current Month</th>" _
        & "<th style='background-color:#004d00;color:white'>Previous Month</th>" _
        & "<th style='background-color:#004d00;color:white'>Previous Quarter</th>"

    htmlStr = htmlStr & RangeToHTML(rng1, 4) & "</table><br><br>"
    
    htmlStr = htmlStr & "<table style='width:200pt'><tr><th style='background-color:#004d00;color:white'>security_id</th>" _
        & "<th style='background-color:#004d00;color:white'>portia_id_bk</th></tr>"
        
    htmlStr = htmlStr & RangeToHTML(rng2, 2)
      
    generateHTMLTable = htmlStr & "</table>"
    
End Function
