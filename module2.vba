''''''''''''''''''''''''''''''''''''''''''''''''''
' FUNCTIONS Modules
' P.O.G.C
' This Project is Desgin for Pars oil and Gas Company\
' Developer : Reza meshkat
' contact : Meshkat@ymail.com
'           0915 316 0277
''''''''''''''''''''''''''''''''''''''''''''''''''''''

Function BrowseForFolder(Optional OpenAt As Variant) As Variant
     'Function purpose:  To Browser for a user selected folder.
     'If the "OpenAt" path is provided, open the browser at that directory
     'NOTE:  If invalid, it will open at the Desktop level
     
    Dim ShellApp As Object
     
     'Create a file browser window at the default folder
    Set ShellApp = CreateObject("Shell.Application"). _
    BrowseForFolder(0, "Please choose a folder", 0, OpenAt)
     
     'Set the folder to that selected.  (On error in case cancelled)
    On Error Resume Next
    BrowseForFolder = ShellApp.self.Path
    On Error GoTo 0
     
     'Destroy the Shell Application
    Set ShellApp = Nothing
     
     'Check for invalid or non-entries and send to the Invalid error
     'handler if found
     'Valid selections can begin L: (where L is a letter) or
     '\\ (as in \\servername\sharename.  All others are invalid
    Select Case Mid(BrowseForFolder, 2, 1)
    Case Is = ":"
        If Left(BrowseForFolder, 1) = ":" Then GoTo Invalid
    Case Is = "\"
        If Not Left(BrowseForFolder, 1) = "\" Then GoTo Invalid
    Case Else
        GoTo Invalid
    End Select
     
    Exit Function
     
Invalid:
     'If it was determined that the selection was invalid, set to False
    BrowseForFolder = False
     
End Function


Function ExtractFilePath(filename As String) As String
   p = InStrRev(filename, "\")
   ExtractFilePath = Left(filename, p)
End Function

Function ExtractFileName(filename As String) As String
    p = Len(filename) - InStrRev(filename, "\")
    ExtractFileName = Right(filename, p)
End Function

Public Function ExtractFileExtention(filename As String) As String
    For i = Len(filename) To 2 Step -1
      C = Mid(filename, i, 1)
      If C = "." Then
      pos = i + 1
      End If
    Next
    ExtractFileExtention = Mid(filename, pos, (Len(filename) + 1 - pos))
End Function

Function findSheets(wb As Workbook) As String
    'return sheets name seperating by :
    Dim WS As Worksheet
    Dim s As String
    
    For Each WS In wb.Worksheets
        If s = "" Then
            s = WS.Name
        Else
            s = s & ":" & WS.Name
        End If
    Next WS
    findSheets = s
End Function


Function FindAllOnWorksheets(InWorkbook As Workbook, _
                InWorksheets As Variant, _
                SearchAddress As String, _
                FindWhat As String, _
                Optional LookIn As XlFindLookIn = xlValues, _
                Optional LookAt As XlLookAt = xlWhole, _
                Optional SearchOrder As XlSearchOrder, _
                Optional MatchCase As Boolean = False, _
                Optional BeginsWith As String = vbNullString, _
                Optional EndsWith As String = vbNullString, _
                Optional BeginEndCompare As VbCompareMethod = vbTextCompare) As Variant
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' FindAllOnWorksheets
' This function searches a range on one or more worksheets, in the range specified by
' SearchAddress.
'
' InWorkbook specifies the workbook in which to search. If this is Nothing, the active
'   workbook is used.
'
' InWorksheets specifies what worksheets to search. InWorksheets can be any of the
' following:
'   - Empty: This will search all worksheets of the workbook.
'   - String: The name of the worksheet to search.
'   - String: The names of the worksheets to search, separated by a ':' character.
'   - Array: A one dimensional array whose elements are any of the following:
'           - Object: A worksheet object to search. This must be in the same workbook
'               as InWorkbook.
'           - String: The name of the worksheet to search.
'           - Number: The index number of the worksheet to search.
' If any one of the specificed worksheets is not found in InWorkbook, no search is
' performed. The search takes place only after everything has been validated.
'
' The other parameters have the same meaning and effect on the search as they do
' in the Range.Find method.
'
' Most of the code in this procedure deals with the InWorksheets parameter to give
' the absolute maximum flexibility in specifying which sheet to search.
'
' This function requires the FindAll procedure, also in this module or avaialable
' at www.cpearson.com/Excel/FindAll.aspx.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim WSArray() As String
Dim WS As Worksheet
Dim wb As Workbook
Dim ResultRange() As Range
Dim WSNdx As Long
Dim R As Range
Dim SearchRange As Range
Dim FoundRange As Range
Dim wss As Variant
Dim n As Long


'''''''''''''''''''''''''''''''''''''''''''
' Determine what Workbook to search.
'''''''''''''''''''''''''''''''''''''''''''
If InWorkbook Is Nothing Then
    Set wb = ActiveWorkbook
Else
    Set wb = InWorkbook
End If

'''''''''''''''''''''''''''''''''''''''''''
' Determine what sheets to search
'''''''''''''''''''''''''''''''''''''''''''
If IsEmpty(InWorksheets) = True Then
    ''''''''''''''''''''''''''''''''''''''''''
    ' Empty. Search all sheets.
    ''''''''''''''''''''''''''''''''''''''''''
    With wb.Worksheets
        ReDim WSArray(1 To .Count)
        For WSNdx = 1 To .Count
            WSArray(WSNdx) = .item(WSNdx).Name
        Next WSNdx
    End With

Else
    '''''''''''''''''''''''''''''''''''''''
    ' If Object, ensure it is a Worksheet
    ' object.
    ''''''''''''''''''''''''''''''''''''''
    If IsObject(InWorksheets) = True Then
        If TypeOf InWorksheets Is Excel.Worksheet Then
            ''''''''''''''''''''''''''''''''''''''''''
            ' Ensure Worksheet is in the WB workbook.
            ''''''''''''''''''''''''''''''''''''''''''
            If StrComp(InWorksheets.Parent.Name, wb.Name, vbTextCompare) <> 0 Then
                ''''''''''''''''''''''''''''''
                ' Sheet is not in WB. Get out.
                ''''''''''''''''''''''''''''''
                Exit Function
            Else
                ''''''''''''''''''''''''''''''
                ' Same workbook. Set the array
                ' to the worksheet name.
                ''''''''''''''''''''''''''''''
                ReDim WSArray(1 To 1)
                WSArray(1) = InWorksheets.Name
            End If
        Else
            '''''''''''''''''''''''''''''''''''''
            ' Object is not a Worksheet. Get out.
            '''''''''''''''''''''''''''''''''''''
        End If
    Else
        '''''''''''''''''''''''''''''''''''''''''''
        ' Not empty, not an object. Test for array.
        '''''''''''''''''''''''''''''''''''''''''''
        If IsArray(InWorksheets) = True Then
            '''''''''''''''''''''''''''''''''''''''
            ' It is an array. Test if each element
            ' is an object. If it is a worksheet
            ' object, get its name. Any other object
            ' type, get out. Not an object, assume
            ' it is the name.
            ''''''''''''''''''''''''''''''''''''''''
            ReDim WSArray(LBound(InWorksheets) To UBound(InWorksheets))
            For WSNdx = LBound(InWorksheets) To UBound(InWorksheets)
                If IsObject(InWorksheets(WSNdx)) = True Then
                    If TypeOf InWorksheets(WSNdx) Is Excel.Worksheet Then
                        ''''''''''''''''''''''''''''''''''''''
                        ' It is a worksheet object, get name.
                        ''''''''''''''''''''''''''''''''''''''
                        WSArray(WSNdx) = InWorksheets(WSNdx).Name
                    Else
                        ''''''''''''''''''''''''''''''''
                        ' Other type of object, get out.
                        ''''''''''''''''''''''''''''''''
                        Exit Function
                    End If
                Else
                    '''''''''''''''''''''''''''''''''''''''''''
                    ' Not an object. If it is an integer or
                    ' long, assume it is the worksheet index
                    ' in workbook WB.
                    '''''''''''''''''''''''''''''''''''''''''''
                    Select Case UCase(TypeName(InWorksheets(WSNdx)))
                        Case "LONG", "INTEGER"
                            err.Clear
                            '''''''''''''''''''''''''''''''''''
                            ' Ensure integer if valid index.
                            '''''''''''''''''''''''''''''''''''
                            Set WS = wb.Worksheets(InWorksheets(WSNdx))
                            If err.Number <> 0 Then
                                '''''''''''''''''''''''''''''''
                                ' Invalid index.
                                '''''''''''''''''''''''''''''''
                                Exit Function
                            End If
                            ''''''''''''''''''''''''''''''''''''
                            ' Valid index. Get name.
                            ''''''''''''''''''''''''''''''''''''
                            WSArray(WSNdx) = wb.Worksheets(InWorksheets(WSNdx)).Name
                        Case "STRING"
                            err.Clear
                            '''''''''''''''''''''''''''''''''''''
                            ' Ensure valid name.
                            '''''''''''''''''''''''''''''''''''''
                            Set WS = wb.Worksheets(InWorksheets(WSNdx))
                            If err.Number <> 0 Then
                                '''''''''''''''''''''''''''''''''
                                ' Invalid name, get out.
                                '''''''''''''''''''''''''''''''''
                                Exit Function
                            End If
                            WSArray(WSNdx) = InWorksheets(WSNdx)
                    End Select
                End If
                'WSArray(WSNdx) = InWorksheets(WSNdx)
            Next WSNdx
        Else
            ''''''''''''''''''''''''''''''''''''''''''''
            ' InWorksheets is neither an object nor an
            ' array. It is either the name or index of
            ' the worksheet.
            ''''''''''''''''''''''''''''''''''''''''''''
            Select Case UCase(TypeName(InWorksheets))
                Case "INTEGER", "LONG"
                    '''''''''''''''''''''''''''''''''''''''
                    ' It is a number. Ensure sheet exists.
                    '''''''''''''''''''''''''''''''''''''''
                    err.Clear
                    Set WS = wb.Worksheets(InWorksheets)
                    If err.Number <> 0 Then
                        '''''''''''''''''''''''''''''''
                        ' Invalid index, get out.
                        '''''''''''''''''''''''''''''''
                        Exit Function
                    Else
                        WSArray = Array(wb.Worksheets(InWorksheets).Name)
                    End If
                Case "STRING"
                    '''''''''''''''''''''''''''''''''''''''''''''''''''
                    ' See if the string contains a ':' character. If
                    ' so, the InWorksheets contains a string of multiple
                    ' worksheets.
                    '''''''''''''''''''''''''''''''''''''''''''''''''''
                    If InStr(1, InWorksheets, ":", vbBinaryCompare) > 0 Then
                        ''''''''''''''''''''''''''''''''''''''''''
                        ' ":" character found. split apart sheet
                        ' names.
                        ''''''''''''''''''''''''''''''''''''''''''
                        wss = Split(InWorksheets, ":")
                        err.Clear
                        n = LBound(wss)
                        If err.Number <> 0 Then
                            '''''''''''''''''''''''''''''
                            ' Unallocated array. Get out.
                            '''''''''''''''''''''''''''''
                            Exit Function
                        End If
                        If LBound(wss) > UBound(wss) Then
                            '''''''''''''''''''''''''''''
                            ' Unallocated array. Get out.
                            '''''''''''''''''''''''''''''
                            Exit Function
                        End If
                            
                                                
                        ReDim WSArray(LBound(wss) To UBound(wss))
                        For n = LBound(wss) To UBound(wss)
                            err.Clear
                            Set WS = wb.Worksheets(wss(n))
                            If err.Number <> 0 Then
                                Exit Function
                            End If
                            WSArray(n) = wss(n)
                         Next n
                    Else
                        err.Clear
                        
                        'Set WS = WB.Worksheets(InWorksheets)
                        Set WS = wb.Worksheets.item(1)
                        If err.Number <> 0 Then
                            '''''''''''''''''''''''''''''''''
                            ' Invalid name, get out.
                            '''''''''''''''''''''''''''''''''
                            Exit Function
                        Else
                            ReDim WSArray(1 To 1)
                            WSArray(1) = InWorksheets
                        End If
                    End If
            End Select
        End If
    End If
End If
'''''''''''''''''''''''''''''''''''''''''''
' Ensure SearchAddress is valid
'''''''''''''''''''''''''''''''''''''''''''
On Error Resume Next
For WSNdx = LBound(WSArray) To UBound(WSArray)
    err.Clear
    Set WS = wb.Worksheets(WSArray(WSNdx))
    ''''''''''''''''''''''''''''''''''''''''
    ' Worksheet does not exist
    ''''''''''''''''''''''''''''''''''''''''
    If err.Number <> 0 Then
        Exit Function
    End If
    err.Clear
    Set R = wb.Worksheets(WSArray(WSNdx)).Range(SearchAddress)
    If err.Number <> 0 Then
        ''''''''''''''''''''''''''''''''''''
        ' Invalid Range. Get out.
        ''''''''''''''''''''''''''''''''''''
        Exit Function
    End If
Next WSNdx

''''''''''''''''''''''''''''''''''''''''
' SearchAddress is valid for all sheets.
' Call FindAll to search the range on
' each sheet.
''''''''''''''''''''''''''''''''''''''''
ReDim ResultRange(LBound(WSArray) To UBound(WSArray))
For WSNdx = LBound(WSArray) To UBound(WSArray)
    Set WS = wb.Worksheets(WSArray(WSNdx))
    Set SearchRange = WS.Range(SearchAddress)
    Set FoundRange = WildCardMatchCells(SearchRange, _
                        FindWhat, _
                        SearchOrder, _
                        MatchCase)
    'FindAll(SearchRange:=SearchRange, _
                    FindWhat:=FindWhat, _
                    LookIn:=LookIn, LookAt:=LookAt, _
                    SearchOrder:=SearchOrder, _
                    MatchCase:=MatchCase, _
                    BeginsWith:=BeginsWith, _
                    EndsWith:=EndsWith, _
                    BeginEndCompare:=BeginEndCompare)
                    
    
    If FoundRange Is Nothing Then
        Set ResultRange(WSNdx) = Nothing
    Else
        Set ResultRange(WSNdx) = FoundRange
    End If
Next WSNdx

 FindAllOnWorksheets = ResultRange

End Function
Function WildCardMatchCells(SearchRange As Range, CompareLikeString As String, _
    Optional SearchOrder As XlSearchOrder = xlByRows, _
    Optional MatchCase As Boolean = False) As Range
    
    Dim FoundCells As Range
    Dim FirstCell As Range
    Dim LastCell As Range
    Dim RowNdx As Long
    Dim ColNdx As Long
    Dim StartRow As Long
    Dim EndRow As Long
    Dim StartCol As Long
    Dim EndCol As Long
    Dim WS As Worksheet
    Dim Rng As Range
    
    
    If SearchRange Is Nothing Then
        Exit Function
    End If
    If SearchRange.Areas.Count > 1 Then
        Exit Function
    End If
    
    With SearchRange
      
        Set WS = .Worksheet
        Set FirstCell = .Cells(1)
        Set LastCell = .Cells(.Cells.Count)
    End With
    
    StartRow = FirstCell.Row
    StartCol = FirstCell.Column
    EndRow = LastCell.Row
    EndCol = LastCell.Column
    
    If SearchOrder = xlByRows Then
      
        With WS
    
            For RowNdx = StartRow To EndRow
     
                For ColNdx = StartCol To EndCol
                    Set Rng = .Cells(RowNdx, ColNdx)
                    If MatchCase = False Then
                        
                        If UCase(Rng.Text) Like UCase(CompareLikeString) Then
                            If FoundCells Is Nothing Then
                                Set FoundCells = Rng
                            Else
                                Set FoundCells = Application.Union(FoundCells, Rng)
                            End If
                        End If
                    Else
                      
                        If Rng.Text Like CompareLikeString Then
                            If FoundCells Is Nothing Then
                                Set FoundCells = Rng
                            Else
                                Set FoundCells = Application.Union(FoundCells, Rng)
                            End If
                        End If ' Like
                    End If ' MatchCase
                Next ColNdx
            Next RowNdx
        End With
    Else
    
        With WS
    
            For ColNdx = StartCol To EndCol
               
                For RowNdx = StartRow To EndRow
                    Set Rng = .Cells(RowNdx, ColNdx)
                    If MatchCase = False Then
                       
                        If UCase(Rng.Text) Like UCase(CompareLikeString) Then
                            If FoundCells Is Nothing Then
                                Set FoundCells = Rng
                            Else
                                Set FoundCells = Application.Union(FoundCells, Rng)
                            End If
                        End If
                    Else
                       
                        If Rng.Text Like CompareLikeString Then
                            If FoundCells Is Nothing Then
                                Set FoundCells = Rng
                            Else
                                Set FoundCells = Application.Union(FoundCells, Rng)
                            End If
                        End If ' Like
                    End If ' MatchCase
                Next RowNdx
            Next ColNdx
        End With
    End If ' SearchOrder
    
    
    If FoundCells Is Nothing Then
        Set WildCardMatchCells = Nothing
    Else
        Set WildCardMatchCells = FoundCells
    End If

End Function


Function FindAll(SearchRange As Range, FindWhat As Variant, _
    Optional LookIn As XlFindLookIn = xlValues, Optional LookAt As XlLookAt = xlWhole, _
    Optional SearchOrder As XlSearchOrder = xlByRows, _
    Optional MatchCase As Boolean = False) As Range

    Dim FoundCell As Range
    Dim FoundCells As Range
    Dim LastCell As Range
    Dim FirstAddr As String

    With SearchRange
        
        Set LastCell = .Cells(.Cells.Count)
    End With
    
    Set FoundCell = SearchRange.Find(what:=FindWhat, after:=LastCell, _
        LookIn:=LookIn, LookAt:=LookAt, SearchOrder:=SearchOrder, MatchCase:=MatchCase)
    If Not FoundCell Is Nothing Then
    
         Set FoundCells = FoundCell
        
         FirstAddr = FoundCell.Address
         Do
             Set FoundCells = Application.Union(FoundCells, FoundCell)
             Set FoundCell = SearchRange.FindNext(after:=FoundCell)
             s = (FoundCell Is Nothing)
             If Not s Then s = (FoundCell.Address = FirstAddr)
         Loop Until s
    End If

    
    If FoundCells Is Nothing Then
        Set FindAll = Nothing
    Else
        Set FindAll = FoundCells
    End If
End Function

Function SplitMultiDelims(Text As String, DelimChars As String, Optional includeOperators As Boolean = False) As String()

    Dim Pos1 As Long
    Dim n As Long
    Dim M As Long
    Dim Arr() As String
    Dim i As Long
    
    If Len(Text) = 0 Then
        Exit Function
    End If
    
    If DelimChars = vbNullString Then
        SplitMultiDelims = Array(Text)
        Exit Function
    End If
    
    ReDim Arr(1 To Len(Text))
    
    i = 0
    n = 0
    Pos1 = 1
    
    For n = 1 To Len(Text)
        For M = 1 To Len(DelimChars)
            If StrComp(Mid(Text, n, 1), Mid(DelimChars, M, 1), vbTextCompare) = 0 Then
                i = i + 1
                Arr(i) = Mid(Text, Pos1, n - Pos1)
                Pos1 = n + 1
                n = n + 1
                If includeOperators Then
                    i = i + 1
                    Arr(i) = Mid(Text, Pos1 - 1, 1)
                End If
            End If
        Next M
    Next n
    
    If Pos1 <= Len(Text) Then
        i = i + 1
        Arr(i) = Mid(Text, Pos1)
    End If
    
    ReDim Preserve Arr(1 To i)
    SplitMultiDelims = Arr
        
End Function
Function DoFunc(FcName, FcParam, Val As String) As String

    Dim SearchRange As Range
    Dim FoundCells As Range
    Dim FoundCell As Range
    Dim FindWhat As Variant
    Dim MatchCase As Boolean
    Dim LookIn As XlFindLookIn
    Dim LookAt As XlLookAt
    Dim SearchOrder As XlSearchOrder

   Dim FFCounter As Byte
   Dim X As String
   
   FFCounter = 0
   Select Case FcName
       Case "CIE"
            'Count If Equal (CIE)
            CIF = Split(Val, ":")
            X = CIF(0)
            RD = RefferedDocument(X)
            Workbooks(RefArr(RD).rTitle & ".xls").Activate
            'Windows(RefArr(RD).rTitle & ".xls").Visible = False
            
            
            Set SearchRange = Range("A1:CZ250")
            FindWhat = CIF(1)
            LookIn = xlValues
            LookAt = xlWhole
            SearchOrder = xlByRows
            MatchCase = False
                       
                       
            Set FoundCells = FindAll(SearchRange:=SearchRange, FindWhat:=FindWhat, _
                  LookIn:=LookIn, LookAt:=LookAt, SearchOrder:=SearchOrder, MatchCase:=MatchCase)
    
            Workbooks(RefArr(RD).rMainFileTitle).Activate
            'Windows(RefArr(RD).rMainFileTitle).Visible = False '========================
            
            If FoundCells Is Nothing Then
                Debug.Print "No cells found."
                DoFunc = ""
            Else
                For Each FoundCell In FoundCells.Cells
                    FF = Range(FoundCell.Address).Value
                    If UCase(FF) = UCase(FcParam) Then
                        FFCounter = FFCounter + 1
                    End If
                Next FoundCell
            End If
            DoFunc = Str(FFCounter)
   End Select
   
End Function

Function CalcValue(contents As String) As String
    Dim RCC() As String
    Dim ContArr() As String
    If (InStr(1, contents, "[[", 1) > 0 And InStr(1, contents, "]]", 1) > 0) Then
        contents = Replace(contents, "[[", "~")
        contents = Replace(contents, "]]", "~")
        ContArr = Split(contents, "~")
        contents = ContArr(1)
        
        RCC = Split(contents, "|")
        If UBound(RCC) = 0 Then
            CalcValue = ContArr(0) & GetValueOf(contents) & ContArr(2)
            Exit Function
        Else
            FC = Split(RCC(0), ":")
            FcN = FC(0)
            FcP = FC(1)
            CalcValue = ContArr(0) & DoFunc(FcN, FcP, RCC(1)) & ContArr(2)
            Exit Function
        End If
    Else
        CalcValue = contents
    End If
End Function

Function RefferedDocument(s As String) As Byte
    ' s would be something like DC,DA,DB,...
    '   A,B or c is refer to a type of document that defined
    '   in main report templates's Refs sheet
    '   they may include a number with them witch shows the
    '   daily source document(in daily to monthly reports)
    '                         s = DC23
    s = Right(s, Len(s) - 1) 's = C23
    s = Left(s, 1)           's = C
    For i = 1 To 10
        If RefArr(i).rType = s Then
            RefferedDocument = i
            Exit Function
        End If
    Next i
    RefferedDocument = 100 'error
End Function

Function Calc(num1, num2 As String, oprator As String) As String
   Select Case oprator
        Case "+"
            res = Val(num1) + Val(num2)
        Case "-"
            res = Val(num1) - Val(num2)
        Case "*"
            res = Val(num1) * Val(num2)
        Case "/"
            res = Val(num1) / Val(num2)
        Case "&"
            res = num1 & num2
   End Select
   Calc = Str(res)
End Function


Function GetValueOf(RCC As String) As String
    Dim SearchRange As Range
    Dim FoundCell As Range
    Dim LastCell As Range
    Dim X As String
    Dim item As Variant
    Dim TotalVal As String
    Dim ArrCode() As String
    Dim NewArrCode() As String
    Dim CellContent As String
    
    
    
    
    ArrCode = SplitMultiDelims(RCC, "+-*/&", True)
    If UBound(ArrCode) > 1 Then
        '############################################
        'this part will calculate * and / and create a simple array
        '   containing + and -
         j = 1
         For i = 1 To UBound(ArrCode)
            If (i) > UBound(ArrCode) Then
                Exit For
            End If
            x2 = ArrCode(i)
            If (x2 = "/") Or (x2 = "*") Then
              
              If x2 = "*" Then
                NewArrCode(j - 2) = Val(GetValueOf(NewArrCode(j - 2))) * Val(GetValueOf(ArrCode(i + 1)))
              Else
                NewArrCode(j - 2) = Val(GetValueOf(NewArrCode(j - 2))) / Val(GetValueOf(ArrCode(i + 1)))
              End If
              i = i + 1
            Else
              ReDim Preserve NewArrCode(j)
              NewArrCode(j - 1) = ArrCode(i)
              j = j + 1
            End If
        Next i
        '#############################################
        j = 0
        i = 0
        For Each item In NewArrCode
          X = item
          If i = 0 Then
            TotalVal = GetValueOf(X)
          Else
              If (i Mod 2) = 0 Then
                 TotalVal = Calc(TotalVal, GetValueOf(X), NewArrCode(i - 1))
              End If
          End If
          i = i + 1
        Next
    Else
        CIF = Split(RCC, ":")
        If (UBound(CIF) < 1) Then
           If IsNumeric(RCC) Then
                TotalVal = RCC
           Else
                Workbooks(FcurrReport).Activate
                CellContent = Range(RCC).Formula
                If (InStr(1, CellContent, "[[", 1) > 0 And InStr(1, CellContent, "]]", 1) > 0) Then
                    s = CalcValue(CellContent)
                    Workbooks(FcurrReport).Activate
                    Range(RCC).Select
                    Selection.Formula = s
                    CellContent = s
                End If
                TotalVal = CellContent
           End If
           GetValueOf = TotalVal
           Exit Function
        End If
        X = CIF(0)
        RD = RefferedDocument(X)
        
        mx = RefArr(RD).rTitle & ".xls"
        Workbooks(mx).Activate
        
        X = CIF(0)
        Set SearchRange = Range("A1:CZ250")
        With SearchRange
            Set LastCell = .Cells(.Cells.Count)
        End With

        
        Set FoundCell = SearchRange.Find(what:=CIF(1), after:=LastCell, LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, MatchCase:=False)
        'Windows(RefArr(RD).rTitle & ".xls").Visible = False '.Visible = False ========================
        If Not FoundCell Is Nothing Then
            GR = Right(X, Len(X) - 2) 'DA3:$e extract 3 here
            If GR = "" Then
                Workbooks(RefArr(RD).rMainFileTitle).Activate
                On Error Resume Next
                If RefArr(RD).rActiveSheet <> "" Then Sheets(RefArr(RD).rActiveSheet).Select
                If err.Number <> 0 Then
                    If Not BaseOnActivesheet Then
                        If MsgBox("Sheet names had modified on " & RefArr(RD).rMainFileTitle & " Do you want to continue based on an active sheet?", vbYesNo) = vbNo Then raiseError ("Report Aborted .")
                        MsgBox ("the result may not be correct. the operation will continue...")
                        BaseOnActivesheet = True
                    End If
                    Workbooks(RefArr(RD).rMainFileTitle).Activate
                End If
                TotalVal = Range(FoundCell.Address).Value
               ' Windows(RefArr(RD).rMainFileTitle).Visible = False '========================
            Else
                Num = frm_Reference.typeToNum(Mid(X, 2, 1))
                If DrArr(Num, GR) = "" Then
                    TotalVal = "" 'no more resource (day reports)
                    GetValueOf = TotalVal
                    Exit Function
                Else
                    ExName = ExtractFileName(DrArr(Num, GR))
                    'On Error Resume Next
                    Workbooks(ExName).Activate
                    On Error Resume Next
                    If RefArr(RD).rActiveSheet <> "" Then Sheets(RefArr(RD).rActiveSheet).Select
                    If err.Number <> 0 Then
                        If Not BaseOnActivesheet Then
                            If MsgBox("Sheet names had modified on " & ExName & " Do you want to continue based on an active sheet?", vbYesNo) = vbNo Then raiseError ("Report Aborted .")
                            MsgBox ("the result may not be correct. the operation will continue...")
                            BaseOnActivesheet = True
                        End If
                        Workbooks(ExName).Activate
                    End If
                    
                    TotalVal = Range(FoundCell.Address).Value
                    'Windows(ExName).Visible = False '========================
                End If
            End If
            
        End If
        
    End If
    GetValueOf = TotalVal
    
End Function

Sub CopySheets(wss)
    'copy the sheets minus last one
    Dim X() As String
  
    If (ActiveWorkbook.Worksheets.Count > 1) Then
        X = Split(wss, ":")
        X(UBound(X)) = ""
        j = 0
        ReDim newArr(LBound(X) To UBound(X) - 1)
        For i = LBound(X) To UBound(X) - 1
            If X(i) <> "" Then
                newArr(j) = X(i)
                j = j + 1
            End If
        Next i
        If UBound(newArr) < 1 Then
            Sheets(newArr(0)).Copy
        Else
            Sheets(newArr).Copy
        End If
        
    Else
        Sheets(wss).Copy
    End If
End Sub


Sub CheckRefCompatibility()
    Dim cellArr() As String
    Dim cell As Variant
    Dim resp
    
     For i = 1 To 11
       If (RefArr(i).rType <> "") And (RefArr(i).rType <> "A") And (RefArr(i).rType <> "R") Then
            If (RefArr(i).rMainFileTitle <> "") And (RefArr(i).rMainFileTitle <> "*") And (RefArr(i).rMainFileTitle <> "DTMR") Then
               cellArr = Split(RefArr(i).rCheckCells, ",")
               For Each cell In cellArr
                    Workbooks(RefArr(i).rMainFileTitle).Activate
                    Range(cell).Select
                    A1 = Range(cell).Value
                    Workbooks(RefArr(i).rTitle & ".xls").Activate
                    a2 = Range(cell).Text
                    Debug.Print A1 + " = " + a2
                    If (A1 <> a2) Then
                        resp = MsgBox("The File " + RefArr(i).rMainFileTitle + " Does Not Match The Reference File. Some Report Data Will Not Be Correct. Continue Anyway ?" + Chr(13) _
                                        + "(You Can Press Cancel To Ignore This Message For All Files)", vbYesNoCancel)
                        If resp = vbCancel Then
                            GoTo exitLoop
                        End If
                        If resp = vbNo Then
                            closeTemplates
                            End
                        End If
                        GoTo endLoop2
                    End If
               Next cell
            Else
                If (RefArr(i).rMainFileTitle = "DTMR") Then
                    cellArr = Split(RefArr(i).rCheckCells, ",")
                    Num = frm_Reference.typeToNum(RefArr(i).rType)
                    
                    For j = 1 To 31
                        If DrArr(Num, j) <> "" Then
                            ExName = ExtractFileName(DrArr(Num, j))
                            For Each cell In cellArr
                                 Workbooks(ExName).Activate
                                 A1 = Range(cell).Text
                                 Workbooks(RefArr(i).rTitle & ".xls").Activate
                                 a2 = Range(cell).Text
                                 Debug.Print A1 + " = " + a2
                                 If (A1 <> a2) Then
                                     resp = MsgBox("The File " + ExName + " Does Not Match The Reference File. Some Report Data Will Not Be Correct. Continue Anyway ?" + Chr(13) _
                                                     + "(You Can Press Cancel To Ignore This Message For All Files)", vbYesNoCancel)
                                     If resp = vbCancel Then
                                         GoTo exitLoop
                                     End If
                                     If resp = vbNo Then
                                         closeTemplates
                                         End
                                     End If
                                     GoTo endloop
                                 End If
                            Next cell
                        
                        End If
endloop:
                    Next j
                    
                End If
            End If
            
       End If
endLoop2:
     Next i
     
exitLoop:
End Sub




Sub macro()
     Dim FoundCell As Range
    Dim CompareLikeString As String
    Dim SearchOrder As XlSearchOrder
    Dim MatchCase As Boolean
    Dim FoundRanges As Variant
    Dim n As Long
    Dim FoundRange As Range
    Dim s As String
    Dim Found As Boolean
    Dim contents As String
    Dim ContArr() As String
    
    'serching for [[*]] in a newly created document
    Set SearchRange = Range("A1:Z50")
    CompareLikeString = "*[[][[]*[]][]]*"
    SearchOrder = xlByRows
    MatchCase = True
    
    
            FoundRanges = FindAllOnWorksheets(InWorkbook:=ThisWorkbook, _
                InWorksheets:=findSheets(ThisWorkbook), _
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
                                contents = FoundCell.Text
                                If (InStr(1, contents, "[[", 1) > 0 And InStr(1, contents, "]]", 1) > 0) Then
                                    contents = Replace(contents, "[[", "~")
                                    contents = Replace(contents, "]]", "~")
                                    ContArr = Split(contents, "~")
                                    'Debug.Print FoundCell.Worksheet.Name & ": " & FoundCell.Address
                                    For i = 0 To UBound(ContArr)
                                        'Debug.Print "     contArr(" & i & ") = " & ContArr(i)
                                    Next i
                                End If
                        Next FoundCell
                    End If
                Next n
            End If
End Sub


Sub TEMPMACRO()
    Debug.Print Range("A1").NumberFormat
End Sub
