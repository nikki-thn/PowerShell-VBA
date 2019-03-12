' Issue_Log_Template.xslm
' This macro captures defined areas of worksheet and attach those captured images to an outlook email
' Version 1.0
' Last Modified: May 25th, 2018


' createJpg takes in 3 parameters to perform capturing excel sheet
Sub createJpg(nameSheet As String, nameRange As String, nameFile As String)

    'Go to the sheet named Namesheet
    ThisWorkbook.Activate
    Worksheets(nameSheet).Activate
    
    'set captureRange to the range passed through parameyer name
    Set captureRange = ThisWorkbook.Worksheets(nameSheet).Range(nameRange)
    
    'Copies the selected range to the Clipboard as a picture
    captureRange.CopyPicture
    
    'Convert the capture to ChartObject for exporting
    With ThisWorkbook.Worksheets(nameSheet).ChartObjects.Add(captureRange.Left, captureRange.Top, captureRange.Width, captureRange.Height)
        .Activate
        .Chart.Paste
        'Export the capture to specified filepath
        .Chart.Export Environ$("temp") & "\" & nameFile & ".jpg", "JPG"
    End With
    
    'Delete the ChartObjects afterward
    Worksheets(nameSheet).ChartObjects(Worksheets(nameSheet).ChartObjects.Count).Delete
    
    
Set captureRange = Nothing

End Sub


' function return the active range (aka. cells with values) from a worksheet
' parameters required are sheet name, and the range of columns
Public Function activeRange(nameSheet As String, rng As String) As String

    ThisWorkbook.Activate
    Worksheets(nameSheet).Activate
    Worksheets(nameSheet).Range(rng).Select

    activeRange = ThisWorkbook.Worksheets(nameSheet).Range(Selection, Selection.End(xlDown)).Address
    
End Function


' Event handler for "SEND EMAIL" button

Private Sub CommandButton1_Click()

On Error GoTo ErrHandler

    Dim OutlookApp As Object
    Dim MItem As Object
    Dim recipients As String
    Dim ccRecipients As String
    Dim bodyMsg As String
    Dim PPMRange As String
    Dim TempFilePath As String
    
    'call activeRange to get a current available cells in PPM Request sheet
    PPMRange = activeRange("PPMRequest", "A1:H1")
    
    'filepath to the captured images, this path must match the filepath from createJpg subroutine
    TempFilePath = Environ$("temp")
    
    'Create Outlook object
    Set OutlookApp = CreateObject("Outlook.Application")
    
    'Create Mail Item and view before sending
    Set MItem = OutlookApp.CreateItem(olMailItem)
    
    'Set recipients
     recipients = "email@meial.com"
     recipients = recipients & "djdjd@jfjf.com " ;


    
    'Set CC recipients
     ccRecipients = "ejej@jejd.com "

    'Set email bosy message
    bodyMsg = "<span style='font-size:11pt'>Hi Team, <br><br> Below is a summary of the requests that are in progress. Please reach out to us if you have any issues/concerns.</span><br><br><br>"
    
    'Set Outlook email variables
    With MItem
        .Display 'this line is crucial to insert outlook signature to the email
        .to = recipients 'add recipients
        .cc = ccRecipients 'add cc recipients
        'Set email subject and current date
        .Subject = "Client Onboarding and Reporting - Issue Log as of " & Date
        
        'create the first capture for the charts
        Call createJpg("Charts", "E3:N41", "appImg")
        .Attachments.Add TempFilePath & "\appImg.jpg", olByValue, 0
        
        'create the second capture for PPMRequest
        Call createJpg("PPMRequest", PPMRange, "ppmImg")
        .Attachments.Add TempFilePath & "\ppmImg.jpg", olByValue, 0

        'create HTML code for those images to insert to email body
        imagesAttachment = "<img src='cid:appImg.jpg'><br><br>"
        imagesAttachment = imagesAttachment & "<img src='cid:ppmImg.jpg'><br><br>"
        
        'put everything together
        .HTMLBody = bodyMsg + imagesAttachment + "<span style='font-size:10pt'>This email was automatically generated</span>" + .HTMLBody 'the .HTMLBody is a must to insert default signature to the email"
        
        'display the email
        .Display
        
    End With
    
    ' Delete the images after email
    Kill TempFilePath & "\appImg.jpg"
    Kill TempFilePath & "\ppmImg.jpg"
    Kill "H:\requestSearch.xlsx"
    
    ' CLEAR
    Set objEmail = Nothing:    Set objOutlook = Nothing
        
ErrHandler:

    ' When error occurs, the code stops and displays the error.
    On Error GoTo 0

End Sub


'Credits to:
'http://www.encodedna.com/excel/send-email-from-excel-using-vba-and-outlook.htm
'https://stackoverflow.com/questions/30552624/how-to-copy-text-and-charts-in-an-excel-sheet-to-outlook-body
'https://www.mrexcel.com/forum/excel-questions/532048-vba-excel-outlook-create-one-email-list-email-addresses.html
'https://www.experts-exchange.com/questions/27979826/Excel-VBA-take-Screenshot-and-then-Embed-in-Email-Message-Body-using-Outlook.html


