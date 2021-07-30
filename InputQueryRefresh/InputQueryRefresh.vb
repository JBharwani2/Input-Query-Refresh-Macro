Sub InputQueryRefresh()
    '
    ' InputQueryRefresh Macro
    ' Created by Jeremy Bharwani on 7/1/21
    ' (questions- email jcb926@gmail.com)
    '
    ' There are four csv formatted queries and an excel spreadsheet with data that is used within the 'input' worksheet for each month's CHS file.
    ' When this macro is run it asks for the user to input the date for the file using a simplified 4 digit format that includes the month and year
    ' of the CHS file. Once submitted, the macro opens the files, grabs the data, and closes the files without saving any changes. This all happens
    ' behind the scenes with no visible movements in order to save computing power and time.
    '
    ' An issue that may arise would be a change in any of the file names or the row numbers within the CHS file. These issues can be dealt with by
    ' editing the variables near the top of this macro under the seciton labeled: "DEFINE VARIABLES".
    '
    '   Time: > 10 seconds
    '   References: none
    '

    'VARIABLES ------------------------------------------------------------------------------------------------------------------------------------------
    Dim wsQueryA As Worksheet
    Dim wsQueryB As Worksheet
    Dim wsQueryC As Worksheet
    Dim wsQueryD As Worksheet
    Dim QueryAFileName As String
    Dim QueryBFileName As String
    Dim QueryCFileName As String
    Dim QueryDFileName As String
    Dim QueryEFileName As String
    Dim QueryCWorksheetName As String
    Dim wsDest As Worksheet
    Dim col As Integer
    Dim extRow As Integer
    Dim fileDate As String
    Dim month As String
    Dim year As String
    Dim inv As String
    Dim invNum As Integer
    Dim numStart As String
    Dim numEnd As String
    Dim complete As Boolean

    'Gets user input for the month and year of this batch of files
    fileDate = InputBox("Please input the month and year for the queries you want to access using this format: 0521")
    month = Left(fileDate, 2)
    year = Right(fileDate, 2)


    'DEFINE VARIABLES -----------------------------------------------------------------------------------------------------------------------------------
    '(if file names or row numbers are changed in the future they must be changes in this area below)
    '(the fileDate variable corresponds to the exact date format that is entered by the user and is appended to the text using the '+' symbol)
    QueryAFileName = "Region & Branch " + fileDate
    QueryBFileName = "Chargeoff " + fileDate
    QueryCFileName = "Extensions by Investor and Pool Report -" + fileDate + " EOM"
    QueryCWorksheetName = "Investor Summary"
    QueryDFileName = "Payoff " + fileDate
    QueryEFileName = "Proceeds " + fileDate

    RegExtensionsDollarsRow = 50
    RegExtensionsCountRow = 51
    PrepaymentsRow = 78
    LiquidationProceedsRow = 80
    ChargeOffBalanceRow = 83
    TotalNumActiveAcctsRow = 86
    NetBalanceRow = 87
    CurrentBalanceRow = 88
    MOSXBALRow = 89
    RATEXBALRow = 90
    '(the above code is the only location where these variables would need to be edited)


    'ACCESS DATA FILES ----------------------------------------------------------------------------------------------------------------------------------
    Application.ScreenUpdating = False 'improves time efficiency
    Set wsDest = ThisWorkbook.Worksheets("input") 'destination is the file that the macro is run from
    
    Workbooks.Open "{FILEPATH}" + year + "\" + year + "{FILEPATH}" + fileDate + "\" + QueryAFileName + ".csv"
    Set wsQueryA = Workbooks(QueryAFileName + ".csv").Worksheets(QueryAFileName)
    
    Workbooks.Open "{FILEPATH}" + year + "\" + year + "{FILEPATH}" + fileDate + "\" + QueryBFileName + ".csv"
    Set wsQueryB = Workbooks(QueryBFileName + ".csv").Worksheets(QueryBFileName)
    
    Workbooks.Open "{FILEPATH}" + year + "\" + year + "{FILEPATH}" + fileDate + "\" + QueryCFileName + ".xlsb"
    Set wsQueryC = Workbooks(QueryCFileName + ".xlsb").Worksheets(QueryCWorksheetName)
    
    Workbooks.Open "{FILEPATH}" + year + "\" + year + "{FILEPATH}" + fileDate + "\" + QueryDFileName + ".csv"
    Set wsQueryD = Workbooks(QueryDFileName + ".csv").Worksheets(QueryDFileName)
    
    Workbooks.Open "{FILEPATH}" + year + "\" + year + "{FILEPATH}" + fileDate + "\" + QueryEFileName + ".csv"
    Set wsQueryE = Workbooks(QueryEFileName + ".csv").Worksheets(QueryEFileName)
    
    complete = False
    col = 3
    extRow = 6

    'FILTER AND GRAB DATA -------------------------------------------------------------------------------------------------------------------------------
    While complete = False
        If (wsDest.Cells(4, col) = "CPS AR") Then
            'Splitting inventory number into searchable format
            inv = wsDest.Cells(3, col) 'String
            invNum = wsDest.Cells(3, col) 'Int
            numStart = Left(inv, 1)
            numEnd = Right(inv, 2)

            'Filters down to the single row that corresponds with the inventory number
            wsQueryA.Range("A1").AutoFilter()
            wsQueryA.Range("$A$1:$L$2000").AutoFilter Field:=1, Criteria1:="2"
            wsQueryA.Range("$A$1:$L$2000").AutoFilter Field:=3, Criteria1:=numStart
            wsQueryA.Range("$A$1:$L$2000").AutoFilter Field:=4, Criteria1:=numEnd
            wsQueryB.Range("A1").AutoFilter()
            wsQueryB.Range("$A$1:$L$2000").AutoFilter Field:=1, Criteria1:="2"
            wsQueryB.Range("$A$1:$L$2000").AutoFilter Field:=3, Criteria1:=numStart
            wsQueryB.Range("$A$1:$L$2000").AutoFilter Field:=4, Criteria1:=numEnd
            'Grabs the necessary data from the one row left after filtering
            wsDest.Cells(TotalNumActiveAcctsRow, col) = wsQueryA.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Cells(1, 8)
            wsDest.Cells(NetBalanceRow, col) = wsQueryA.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Cells(1, 6)
            wsDest.Cells(CurrentBalanceRow, col) = wsQueryA.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Cells(1, 7)
            wsDest.Cells(MOSXBALRow, col) = wsQueryA.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Cells(1, 11)
            wsDest.Cells(RATEXBALRow, col) = wsQueryA.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Cells(1, 12)
            wsDest.Cells(ChargeOffBalanceRow, col) = wsQueryB.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Cells(1, 7)

            'Uses SUMIFS equations to grab necessary data
            wsQueryD.Cells(1, 9).Formula = "=SUMIFS(G:G, A:A, 4, B:B, " + numStart + ", C:C, " + numEnd + ")"
            wsDest.Cells(PrepaymentsRow, col) = wsQueryD.Cells(1, 9).Value
            wsQueryE.Cells(1, 10).Formula = "=SUMIFS(I:I, B:B, " + numStart + ", C:C, " + numEnd + ")"
            wsDest.Cells(LiquidationProceedsRow, col) = wsQueryE.Cells(1, 10).Value

            'Grabs extension data by matching the inventory number to the corresponding row of data
            Do While extRow < 100
                If wsQueryC.Cells(extRow, 3) = invNum Then
                    wsDest.Cells(RegExtensionsDollarsRow, col) = wsQueryC.Cells(extRow, 5)
                    wsDest.Cells(RegExtensionsCountRow, col) = wsQueryC.Cells(extRow, 4)
                    Exit Do
                End If
                extRow = extRow + 1
            Loop

            'Moves to the next column
            col = col + 1
        Else
            complete = True
        End If
    Wend
    
'CLEANUP --------------------------------------------------------------------------------------------------------------------------------------------
    'Closes tabs and shows completion message
    Workbooks(QueryAFileName + ".csv").Close SaveChanges:=False
    Workbooks(QueryBFileName + ".csv").Close SaveChanges:=False
    Workbooks(QueryCFileName + ".xlsb").Close SaveChanges:=False
    Workbooks(QueryDFileName + ".csv").Close SaveChanges:=False
    Workbooks(QueryEFileName + ".csv").Close SaveChanges:=False

    wsDest.Cells(RegExtensionsDollarsRow, 1).Font.Color = RGB(0, 176, 80)
    wsDest.Cells(RegExtensionsCountRow, 1).Font.Color = RGB(0, 176, 80)
    wsDest.Cells(ChargeOffBalanceRow, 1).Font.Color = RGB(0, 176, 80)
    wsDest.Cells(TotalNumActiveAcctsRow, 1).Font.Color = RGB(0, 176, 80)
    wsDest.Cells(NetBalanceRow, 1).Font.Color = RGB(0, 176, 80)
    wsDest.Cells(CurrentBalanceRow, 1).Font.Color = RGB(0, 176, 80)
    wsDest.Cells(MOSXBALRow, 1).Font.Color = RGB(0, 176, 80)
    wsDest.Cells(RATEXBALRow, 1).Font.Color = RGB(0, 176, 80)
    wsDest.Cells(PrepaymentsRow, 1).Font.Color = RGB(0, 176, 80)
    wsDest.Cells(LiquidationProceedsRow, 1).Font.Color = RGB(0, 176, 80)

    Application.ScreenUpdating = True
    MsgBox "CHS input file rows " & RegExtensionsDollarsRow & ", " & RegExtensionsCountRow & ", " & PrepaymentsRow & ", " & LiquidationProceedsRow & ", " & ChargeOffBalanceRow & ", " & TotalNumActiveAcctsRow & ", " & NetBalanceRow & ", " & CurrentBalanceRow & ", " & MOSXBALRow & ", and " & RATEXBALRow & " have been successfully refreshed to specified query"

End Sub

