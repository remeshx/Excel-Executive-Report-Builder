''''''''''''''''''''''''''''''''''''''''''''''''''
' MAIN Modules
' P.O.G.C
' This Project is Desgin for Pars oil and Gas Company\
' Developer : Reza meshkat
' contact : Meshkat@ymail.com
'           0915 316 0277
''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Type Reference
    rType As String * 1
    rDescription As String
    rFileAddress As String
    rTitle As String
    rMainFileTitle As String
    rActiveSheet As String
    rCheckCells As String
    rCompatible As Boolean
    rsource As String
End Type
    'compatible is true if the mainFile CheckCells  with the RefFile is the same

Public RefArr(1 To 11) As Reference  ' refrence docs
Public DrArr(10, 31) As String 'daily reports
Public DrArrCompatible(10, 31) As Boolean ' is the file match the ref file
Public ReportType As String
Public currReport As String
Public FcurrReport As String
Public BaseOnActivesheet As Boolean
Public Fpath As String






Function doCreateReport(rpCat, rpType As String, Optional AllSheets As Boolean = False) As Boolean
    'rpCat is  a catagory lke production report
    'rpType is a Type of report like dialy production report
    Dim SearchRange As Range
    Dim FoundCells As Range
    Dim FoundCell As Range
    Dim CompareLikeString As String
    Dim SearchOrder As XlSearchOrder
    Dim MatchCase As Boolean
    Dim FoundRanges As Variant
    Dim n As Long
    Dim FoundRange As Range
    Dim s As String
    Dim Found As Boolean
    Dim filename As String
    
    BaseOnActivesheet = False
    
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Application.AskToUpdateLinks = False
    
    'On Error GoTo ErrF
    'fa = ThisWorkbook.Path & "\forms\" & frm_welcome.cmb_select_type.Text & "\" & frm_welcome.cmb_select_rp.Text & ".xls"
    OpenRefs (ThisWorkbook.Path & "\forms\" & rpCat & "\" & rpType & ".xls")
    
    CheckRefCompatibility

    currDate = Replace(Date, "/", "-")
    filename = Application.GetSaveAsFilename(rpType & " - " & currDate, "Excel WorkBook (*.xls),*.xls", , "Save Report As ...")
    If filename <> "False" Then
        currReport = filename
        FcurrReport = ExtractFileName(filename)
        Workbooks(rpType & ".xls").Activate
        
        
        CopySheets (findSheets(ActiveWorkbook))
        ActiveWorkbook.SaveAs filename:=filename, FileFormat:=xlNormal, _
             Password:="", WriteResPassword:="", ReadOnlyRecommended:=False, _
             CreateBackup:=False
        Windows(rpType & ".xls").Close False
        fname = ExtractFileName(CStr(filename))
    Else
       doCreateReport = False
       closeTemplates
       Windows(rpType & ".xls").Close
       Exit Function
    End If
    
    
    'serching for [[*]] in a newly created document
    Set SearchRange = Range("A1:AZ50")
    CompareLikeString = "*[[][[]*[]][]]*"
    SearchOrder = xlByRows
    MatchCase = True
    
    hasERR = True
    'On Error GoTo ErrF
    Application.ScreenUpdating = False
    
    'Debug.Print findSheets(ActiveWorkbook)
    If AllSheets Then
            FoundRanges = FindAllOnWorksheets(InWorkbook:=ActiveWorkbook, _
                InWorksheets:=findSheets(ActiveWorkbook), _
                SearchAddress:=SearchRange.Address, _
                FindWhat:=CompareLikeString, _
                LookIn:=xlValues, _
                LookAt:=xlWhole, _
                SearchOrder:=xlByRows, _
                MatchCase:=False)
             If UBound(FoundRanges) > 0 Then
                For n = LBound(FoundRanges) To UBound(FoundRanges)
                    If Not FoundRanges(n) Is Nothing Then
                        For Each FoundCell In FoundRanges(n).Cells
                            frm_daily_production_report.Repaint
                            s = FoundCell.Worksheet.Name
                            'Debug.Print FoundCell.Address
                            'On Error Resume Next
                            Workbooks(FcurrReport).Activate
                            Sheets(s).Select
                            X = CalcValue(FoundCell.Value)
                            x2 = X
                            If (FoundCell.NumberFormat <> "@") And (FoundCell.NumberFormat <> "General") _
                                And (X <> "") And (X <> "-") Then X = "=" & X
                            On Error Resume Next
                            FoundCell.Value = X
                            If (err.Number <> 0) Then FoundCell.Value = x2
                                    'Or (IsError(FoundCell.Value))
                            
                            '  Debug.Print FoundCell.Worksheet.Name & ": " & FoundCell.AddressFalse, False)
                        Next FoundCell
                    End If
                Next n
            End If
    Else
        Set FoundCells = WildCardMatchCells(SearchRange:=SearchRange, CompareLikeString:=CompareLikeString, _
            SearchOrder:=SearchOrder, MatchCase:=MatchCase)
        If FoundCells Is Nothing Then
             Debug.Print "No cells found."
             doCreateReport = False
             closeTemplates
             On Error Resume Next
             Windows(currReport).Close flase
             Exit Function
        End If
        
        For Each FoundCell In FoundCells
              frm_daily_production_report.Repaint
              s = FoundCell.Worksheet.Name
              'Debug.Print FoundCell.Address
              'On Error Resume Next
              Workbooks(FcurrReport).Activate
              Sheets(s).Select
              X = CalcValue(FoundCell.Value)
              x2 = X
              If (FoundCell.NumberFormat <> "@") And (FoundCell.NumberFormat <> "General") _
                                And (X <> "") And (X <> "-") Then X = "=" & X
              On Error Resume Next
              FoundCell.Value = X
              If (err.Number <> 0) Then FoundCell.Value = x2
              'Or ((IsError(FoundCell.Value)) And (FoundCell.Value <> CVErr(xlErrValue)))
        Next FoundCell
    End If
    
    hasERR = False
    
ErrF:
    If hasERR Then
        raiseError ("unable to create Report!")
        doCreateReport = False
    Else
        Application.ScreenUpdating = True
        closeTemplates
        doCreateReport = True
    End If
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.AskToUpdateLinks = True
    
End Function

Sub closeTemplates()
     For i = 1 To 10
        On Error Resume Next
         Windows(RefArr(i).rTitle & ".xls").Close False
        On Error Resume Next
         Windows(RefArr(i).rMainFileTitle).Close False
     Next i
     For j = 1 To 10
         For i = 1 To 31
            If DrArr(j, i) <> "" Then
                DrName = ExtractFileName(DrArr(j, i))
                On Error Resume Next
                Windows(DrName).Close False
            End If
        Next i
    Next j
End Sub

Function ReadRefs(rpRef As String)
    'Added in v 1.3
    On Error GoTo EndF
    Workbooks.Open filename:=rpRef
    p = InStrRev(rpRef, "\")
    p2 = Len(rpRef) - p
    
    ExPath = Left(rpRef, p) ' the rpRef path
    ExFile = Right(rpRef, p2) ' the rpRef filename
    
    Sheets("Refs").Select
    Range("B2").Select
    ReportType = ActiveCell.FormulaR1C1
    
    For i = 11 To 21
        Workbooks(ExFile).Activate
    
        Range("A" & i).Select
        If ActiveCell.FormulaR1C1 = "" Then
            GoTo endloop
        End If
              
        RefArr(i - 10).rType = ActiveCell.FormulaR1C1
        Range("B" & i).Select
        RefArr(i - 10).rDescription = ActiveCell.FormulaR1C1
        Range("C" & i).Select
        RefArr(i - 10).rFileAddress = ActiveCell.FormulaR1C1
        Range("D" & i).Select
        RefArr(i - 10).rTitle = ActiveCell.FormulaR1C1
        Range("E" & i).Select
        RefArr(i - 10).rMainFileTitle = ActiveCell.FormulaR1C1
        Range("F" & i).Select
        RefArr(i - 10).rActiveSheet = ActiveCell.FormulaR1C1
        Range("G" & i).Select
        RefArr(i - 10).rCheckCells = ActiveCell.FormulaR1C1
        RefArr(i - 10).rsource = ""

    Next i
    
endloop:

EndF:
    If hasERR Then raiseError ("Can not find References!")
   
    End Function
    

Function OpenRefs(rpRef As String)
 
    p = InStrRev(rpRef, "\")
    p2 = Len(rpRef) - p
    
    ExPath = Left(rpRef, p) ' the rpRef path
    ExFile = Right(rpRef, p2) ' the rpRef filename
    
    For i = 1 To 11
        If RefArr(i).rFileAddress <> "" Then
            Workbooks.Open filename:=ExPath & RefArr(i).rFileAddress
            On Error Resume Next
            If RefArr(i).rActiveSheet <> "" Then Sheets(RefArr(i).rActiveSheet).Select
            Windows(RefArr(i).rTitle & ".xls").Visible = False
            If RefArr(i).rMainFileTitle <> "DTMR" Then
                Workbooks.Open filename:=RefArr(i).rsource
                On Error Resume Next
                If RefArr(i).rActiveSheet <> "" Then Sheets(RefArr(i).rActiveSheet).Select
                fname = ExtractFileName(RefArr(i).rsource)
                RefArr(i).rMainFileTitle = fname
                Windows(fname).Visible = False
            End If
            Workbooks(ExFile).Activate
        End If
        'Windows(ExFile).Visible = False
endloop:
    Next i
    
    For j = 1 To 11
        For i = 1 To 31
            If DrArr(j, i) <> "" Then
                Workbooks.Open filename:=DrArr(j, i)
                DrName = ExtractFileName(DrArr(j, i))
                Windows(DrName).Visible = False
            End If
        Next i
    Next j
    
    hasERR = False
    
EndF:
    If hasERR Then raiseError ("Can not Open References!")
   
End Function
    
Function ReadAndOpenRefs(rpRef As String)
 
    hasERR = True
    On Error GoTo EndF
    Workbooks.Open filename:=rpRef
    p = InStrRev(rpRef, "\")
    p2 = Len(rpRef) - p
    
    ExPath = Left(rpRef, p) ' the rpRef path
    ExFile = Right(rpRef, p2) ' the rpRef filename
    
    Sheets("Refs").Select
    Range("B2").Select
    ReportType = ActiveCell.FormulaR1C1
    
    For i = 11 To 21
        Workbooks(ExFile).Activate
    
        Range("A" & i).Select
        If ActiveCell.FormulaR1C1 = "" Then
            GoTo endloop
        End If
              
        RefArr(i - 10).rType = ActiveCell.FormulaR1C1
        Range("B" & i).Select
        RefArr(i - 10).rDescription = ActiveCell.FormulaR1C1
        Range("C" & i).Select
        RefArr(i - 10).rFileAddress = ActiveCell.FormulaR1C1
        Range("D" & i).Select
        RefArr(i - 10).rTitle = ActiveCell.FormulaR1C1
        Range("E" & i).Select
        RefArr(i - 10).rMainFileTitle = ActiveCell.FormulaR1C1
        Range("F" & i).Select
        RefArr(i - 10).rActiveSheet = ActiveCell.FormulaR1C1
        Range("G" & i).Select
        RefArr(i - 10).rCheckCells = ActiveCell.FormulaR1C1
        If RefArr(i - 10).rFileAddress <> "" Then
            Workbooks.Open filename:=ExPath & RefArr(i - 10).rFileAddress
            On Error Resume Next
            If RefArr(i - 10).rActiveSheet <> "" Then Sheets(RefArr(i - 10).rActiveSheet).Select
            Windows(RefArr(i - 10).rTitle & ".xls").Visible = False
            Workbooks(ExFile).Activate
        End If
        'Windows(ExFile).Visible = False
endloop:
    Next i
    
    For i = 1 To 31
        If DrArr(i) <> "" Then
            Workbooks.Open filename:=DrArr(i)
            DrName = ExtractFileName(DrArr(i))
            Windows(DrName).Visible = False
        End If
    Next i
    
    hasERR = False
    
EndF:
    If hasERR Then raiseError ("Can not Open References!")
   
End Function

Function ReadDocumentType(fname As String) As String
   
    Dim n As String
    On Error GoTo err
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Application.AskToUpdateLinks = False
    
    Workbooks.Open filename:=fname
    n = ExtractFileName(fname)
    Windows(n).Visible = False
    Workbooks(n).Activate
    Sheets("Refs").Select
    Range("B2").Select
    ReadDocumentType = ActiveCell.FormulaR1C1
    'Windows(n).Visible = False
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.AskToUpdateLinks = True
    Exit Function
    
err:
    raiseError ("Cannot Read Document Type")
   
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.AskToUpdateLinks = True
End Function

Private Function Report_daily_Production_Report() As Boolean
   
    ' open the selected source files
    For i = 1 To 10
        If RefArr(i).rTitle = "ref-Daily Production Report" Then
            fname = ExtractFileName(frm_daily_production_report.ed_previous_day_report.Text)
            Workbooks.Open filename:=frm_daily_production_report.ed_previous_day_report.Text
            Windows(fname).Visible = False
            RefArr(i).rMainFileTitle = fname
        End If
        If RefArr(i).rTitle = "ref-Executive Report" Then
            fname = ExtractFileName(frm_daily_production_report.ed_today_executive_report.Text)
            Workbooks.Open filename:=frm_daily_production_report.ed_today_executive_report.Text
            Windows(fname).Visible = False
            RefArr(i).rMainFileTitle = fname
        End If
        
    Next i
    Report_daily_Production_Report = True
   
End Function

Sub raiseError(msg As String)
    closeTemplates
    MsgBox msg
    End
End Sub
