'To load data into the template
Sub Button27_Click()

    Dim owbk As Workbook
    Dim sPath As String
    Dim lastRow As Integer
  
    Dim sheetsName(0 To 3) As String 'an array to store sheets name
    sheetsName(0) = "Weights currency"
    sheetsName(1) = "Weights sector"
    sheetsName(2) = "Weights industry"
    sheetsName(3) = "Weights country"
    
    '-----============ SET THE FILE PATH HERE ===============-------
    Set owbk = Workbooks.Open("H:\CR index\CR Index.xlsx")
    'filePath = InputBox("please enter the path to cr_index file")

    For x = 0 To 3

        lastRow = owbk.Worksheets(sheetsName(x)).Cells(Rows.Count, 1).End(xlUp).Row
        owbk.Sheets(sheetsName(x)).Range("a1:BZ" & CStr(lastRow)).Copy
        ThisWorkbook.Sheets(sheetsName(x)).Range("A1").PasteSpecial xlPasteValues

    Next x
    
    Application.DisplayAlerts = False
    owbk.Close
    
End Sub
