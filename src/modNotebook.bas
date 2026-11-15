Attribute VB_Name = "modNotebook"
'==========================================================================
' Investigator Notebook
'
' Builds one tab per account number pasted on the HOME tab, with a
' structured ACCOUNT INFO block, an activity summary area and a HOME
' button, and keeps an index of all accounts on the HOME tab.
'
' Macros the user runs:
'   SetupHome              - one time: adds the PROCESS button and the
'                            HOME tab dropdowns
'   ProcessAccountNumbers  - creates tabs for the account numbers in E2:E26
'   ShowRemoveAccounts     - opens the form for deleting account tabs
'   GoHome                 - returns to the HOME tab (used by tab buttons)
'==========================================================================
Option Explicit

' ---- Layout of the HOME tab --------------------------------------------
Public Const HOME_SHEET As String = "HOME"
Private Const INPUT_RANGE As String = "E2:E26"     ' where account numbers are pasted
Private Const LIST_FIRST_ROW As Long = 2           ' index table rows (G2:J26)
Private Const LIST_LAST_ROW As Long = 26
Private Const COL_NAME As Long = 7                 ' G  account name
Private Const COL_NUMBER As Long = 8               ' H  account number
Private Const COL_LINK As Long = 9                 ' I  hyperlink to the tab
Private Const SUSPICIOUS_RANGE As String = "J2:J26"
Private Const DISPOSITION_CELL As String = "B6"
Private Const PROCESS_BUTTON_CELL As String = "E28"
Private Const REMOVE_BUTTON_CELL As String = "E31"
Private Const REMOVE_FORM As String = "frmRemoveAccounts"

' ---- Layout of each account tab ----------------------------------------
Public Const INFO_HEADER As String = "ACCOUNT INFO"
Public Const SUMMARY_RANGE As String = "A12:O25"
Public Const SUMMARY_PLACEHOLDER As String = "Type a summary of account activity here..."
Private Const HOME_BUTTON_CELL As String = "E2"


'==========================================================================
' ONE-TIME SETUP
'==========================================================================
Public Sub SetupHome()
    Dim ws As Worksheet
    Dim anchor As Range

    Set ws = ThisWorkbook.Worksheets(HOME_SHEET)

    AddDropdown ws.Range(DISPOSITION_CELL), "File SAR,No SAR"
    AddDropdown ws.Range(SUSPICIOUS_RANGE), "Yes,No"

    ' Store account numbers as text so Excel never reformats them
    ws.Range(ws.Cells(LIST_FIRST_ROW, COL_NUMBER), ws.Cells(LIST_LAST_ROW, COL_NUMBER)).NumberFormat = "@"

    Set anchor = ws.Range(PROCESS_BUTTON_CELL)
    AddButton ws, "btnProcess", "PROCESS ACCOUNT NUMBERS", _
              anchor.Left, anchor.Top, anchor.Width, 32, "ProcessAccountNumbers"

    Set anchor = ws.Range(REMOVE_BUTTON_CELL)
    AddButton ws, "btnRemove", "REMOVE ACCOUNTS", _
              anchor.Left, anchor.Top, anchor.Width, 32, "ShowRemoveAccounts", RGB(192, 0, 0)

    MsgBox "The HOME tab is set up." & vbCrLf & vbCrLf & _
           "Paste account numbers in E2:E26, then click PROCESS ACCOUNT NUMBERS." & vbCrLf & _
           "Use REMOVE ACCOUNTS to delete account tabs you no longer need.", _
           vbInformation, "Investigator Notebook"
End Sub


'==========================================================================
' MAIN PROCESS
'==========================================================================
Public Sub ProcessAccountNumbers()
    Dim wsHome As Worksheet
    Dim cell As Range
    Dim acct As String
    Dim r As Long
    Dim listed As Object          ' account number -> its row in the HOME index
    Dim seenThisRun As Object     ' account numbers already handled in this run
    Dim prevAutoFill As Boolean
    Dim nCreated As Long, nListed As Long
    Dim msgInvalid As String, msgDuplicate As String
    Dim msgAlready As String, msgNoRoom As String

    Set wsHome = ThisWorkbook.Worksheets(HOME_SHEET)
    Set listed = CreateObject("Scripting.Dictionary")
    Set seenThisRun = CreateObject("Scripting.Dictionary")

    ' Accounts already in the index from an earlier run
    For r = LIST_FIRST_ROW To LIST_LAST_ROW
        acct = Trim$(CStr(wsHome.Cells(r, COL_NUMBER).Value))
        If Len(acct) > 0 Then listed(acct) = r
    Next r

    prevAutoFill = Application.AutoCorrect.AutoFillFormulasInLists
    On Error GoTo Failed
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    ' The index is an Excel table. Without this, Excel copies the first
    ' account-name formula into every row of the column.
    Application.AutoCorrect.AutoFillFormulasInLists = False

    For Each cell In wsHome.Range(INPUT_RANGE).Cells
        acct = NormalizeAccount(cell.Value)

        If Len(acct) = 0 Then
            ' blank cell: nothing to do

        ElseIf Not (acct Like "##########") Then
            msgInvalid = msgInvalid & vbCrLf & "   " & cell.Address(False, False) & ":  " & CStr(cell.Value)

        ElseIf seenThisRun.Exists(acct) Then
            msgDuplicate = msgDuplicate & vbCrLf & "   " & acct

        Else
            seenThisRun(acct) = True

            If listed.Exists(acct) Then
                ' Already processed. Rebuild the tab only if it was deleted.
                If SheetExists(acct) Then
                    msgAlready = msgAlready & vbCrLf & "   " & acct
                Else
                    CreateAccountSheet acct
                    WriteIndexRow wsHome, CLng(listed(acct)), acct
                    nCreated = nCreated + 1
                End If
            Else
                r = NextFreeIndexRow(wsHome)
                If r = 0 Then
                    msgNoRoom = msgNoRoom & vbCrLf & "   " & acct
                Else
                    If Not SheetExists(acct) Then
                        CreateAccountSheet acct
                        nCreated = nCreated + 1
                    End If
                    WriteIndexRow wsHome, r, acct
                    listed(acct) = r
                    nListed = nListed + 1
                End If
            End If
        End If
    Next cell

    ' Clear the paste area. Anything skipped is listed in the report below.
    wsHome.Range(INPUT_RANGE).ClearContents

    RestoreSettings prevAutoFill
    wsHome.Activate

    MsgBox BuildReport(nCreated, nListed, msgInvalid, msgDuplicate, msgAlready, msgNoRoom), _
           vbInformation, "Investigator Notebook"
    Exit Sub

Failed:
    RestoreSettings prevAutoFill
    MsgBox "Processing stopped because of an error:" & vbCrLf & vbCrLf & Err.Description, _
           vbExclamation, "Investigator Notebook"
End Sub


Public Sub GoHome()
    ThisWorkbook.Worksheets(HOME_SHEET).Activate
End Sub


'==========================================================================
' REMOVING ACCOUNTS
'==========================================================================
Public Sub ShowRemoveAccounts()
    If Not IsArray(GetAccountList()) Then
        MsgBox "There are no account tabs to remove.", vbInformation, "Investigator Notebook"
        Exit Sub
    End If
    ' Opened by name so this module still compiles if the form is missing
    VBA.UserForms.Add(REMOVE_FORM).Show
End Sub

' All account numbers in the notebook: the HOME list first, then any
' account tab that is missing from the list. Returns Empty if there are none.
Public Function GetAccountList() As Variant
    Dim wsHome As Worksheet
    Dim ws As Worksheet
    Dim found As Object
    Dim r As Long
    Dim acct As String

    Set wsHome = ThisWorkbook.Worksheets(HOME_SHEET)
    Set found = CreateObject("Scripting.Dictionary")

    For r = LIST_FIRST_ROW To LIST_LAST_ROW
        acct = Trim$(CStr(wsHome.Cells(r, COL_NUMBER).Value))
        If Len(acct) > 0 And Not found.Exists(acct) Then found(acct) = True
    Next r

    For Each ws In ThisWorkbook.Worksheets
        If IsAccountSheet(ws) Then
            If Not found.Exists(ws.Name) Then found(ws.Name) = True
        End If
    Next ws

    If found.Count = 0 Then
        GetAccountList = Empty
    Else
        GetAccountList = found.Keys
    End If
End Function

' Deletes each account's tab, its row in the HOME list and its number in
' the paste area. Returns the number of accounts removed.
Public Function RemoveAccounts(ByVal accts As Variant) As Long
    Dim wsHome As Worksheet
    Dim tbl As ListObject
    Dim cell As Range
    Dim acct As String
    Dim i As Long, r As Long, nRemoved As Long
    Dim prevAutoFill As Boolean

    Set wsHome = ThisWorkbook.Worksheets(HOME_SHEET)
    Set tbl = GetIndexTable(wsHome)

    prevAutoFill = Application.AutoCorrect.AutoFillFormulasInLists
    On Error GoTo Failed
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False          ' no "are you sure" prompt per tab
    Application.AutoCorrect.AutoFillFormulasInLists = False

    For i = LBound(accts) To UBound(accts)
        acct = CStr(accts(i))

        ' 1. HOME list: remove the row and add an empty one at the bottom,
        '    so the list stays in one block and keeps 25 rows
        r = FindIndexRow(wsHome, acct)
        If r > 0 Then
            If tbl Is Nothing Then
                wsHome.Range(wsHome.Cells(r, COL_NAME), wsHome.Cells(r, COL_LINK + 1)).ClearContents
                wsHome.Cells(r, COL_LINK).Hyperlinks.Delete
            Else
                tbl.ListRows(r - tbl.HeaderRowRange.Row).Delete
                tbl.ListRows.Add
            End If
        End If

        ' 2. Paste area: clear the number so it is not processed again
        For Each cell In wsHome.Range(INPUT_RANGE).Cells
            If NormalizeAccount(cell.Value) = acct Then cell.ClearContents
        Next cell

        ' 3. The account tab itself
        If SheetExists(acct) Then ThisWorkbook.Worksheets(acct).Delete

        nRemoved = nRemoved + 1
    Next i

    Application.DisplayAlerts = True
    RestoreSettings prevAutoFill
    wsHome.Activate
    RemoveAccounts = nRemoved
    Exit Function

Failed:
    Application.DisplayAlerts = True
    RestoreSettings prevAutoFill
    MsgBox "Removing accounts stopped because of an error:" & vbCrLf & vbCrLf & Err.Description, _
           vbExclamation, "Investigator Notebook"
    RemoveAccounts = nRemoved
End Function

Private Function FindIndexRow(ByVal wsHome As Worksheet, ByVal acct As String) As Long
    Dim r As Long
    For r = LIST_FIRST_ROW To LIST_LAST_ROW
        If Trim$(CStr(wsHome.Cells(r, COL_NUMBER).Value)) = acct Then
            FindIndexRow = r
            Exit Function
        End If
    Next r
    FindIndexRow = 0
End Function

' The Excel table that holds the HOME list (the one containing H1)
Private Function GetIndexTable(ByVal wsHome As Worksheet) As ListObject
    Dim lo As ListObject
    For Each lo In wsHome.ListObjects
        If Not Intersect(lo.Range, wsHome.Cells(1, COL_NUMBER)) Is Nothing Then
            Set GetIndexTable = lo
            Exit Function
        End If
    Next lo
End Function


'==========================================================================
' BUILDING AN ACCOUNT TAB
'==========================================================================
Private Sub CreateAccountSheet(ByVal acct As String)
    Dim wb As Workbook
    Dim ws As Worksheet

    Set wb = ThisWorkbook
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = acct

    BuildAccountInfo ws
    BuildSummaryArea ws
    AddButton ws, "btnHome", "HOME", _
              ws.Range(HOME_BUTTON_CELL).Left, ws.Range(HOME_BUTTON_CELL).Top, 90, 28, "GoHome"
End Sub


Private Sub BuildAccountInfo(ByVal ws As Worksheet)
    Dim labels As Variant
    Dim i As Long
    Dim rowNum As Long

    labels = Array("ACCOUNT NAME", "SIGNER(S)", "SEGMENT", "ACCOUNT TYPE", _
                   "PRODUCT TYPE", "BALANCE", "OPEN DATE", "CLOSING DATE")

    ws.Columns("A").ColumnWidth = 18
    ws.Columns("B:C").ColumnWidth = 20

    ' Header: A1:C1
    With ws.Range("A1:C1")
        .Merge
        .Value = INFO_HEADER
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(217, 217, 217)       ' light gray
    End With

    ' Labels in A2:A9, entry cells B:C merged row by row
    For i = 0 To UBound(labels)
        rowNum = i + 2

        With ws.Cells(rowNum, 1)
            .Value = labels(i)
            .Font.Bold = True
            .Interior.Color = RGB(221, 235, 247)   ' light blue: labels
        End With

        With ws.Range(ws.Cells(rowNum, 2), ws.Cells(rowNum, 3))
            .Merge
            .HorizontalAlignment = xlLeft
            .Interior.Color = RGB(255, 242, 204)   ' light yellow: entry cells
        End With
    Next i

    AddDropdown ws.Range("B4:C4"), "RETAIL,COMMERCIAL,C&IB"     ' SEGMENT
    AddDropdown ws.Range("B5:C5"), "PERSONAL,BUSINESS"          ' ACCOUNT TYPE
    AddDropdown ws.Range("B6:C6"), "CHECKING,SAVINGS"           ' PRODUCT TYPE
    ws.Range("B7:C7").NumberFormat = "$#,##0.00"                ' BALANCE
    ws.Range("B8:C9").NumberFormat = "mm/dd/yyyy"               ' OPEN / CLOSING DATE

    With ws.Range("A1:C9").Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
End Sub


Private Sub BuildSummaryArea(ByVal ws As Worksheet)
    With ws.Range(SUMMARY_RANGE)
        .Merge
        .Interior.Color = RGB(255, 255, 0)         ' yellow
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop                 ' text starts top left
        .WrapText = True
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With
    ShowPlaceholder ws
End Sub


'==========================================================================
' SUMMARY PLACEHOLDER (called from the ThisWorkbook module)
'==========================================================================
Public Function IsAccountSheet(ByVal ws As Worksheet) As Boolean
    IsAccountSheet = (ws.Name Like "##########") And (CStr(ws.Range("A1").Value) = INFO_HEADER)
End Function

Public Sub ShowPlaceholder(ByVal ws As Worksheet)
    With ws.Range(SUMMARY_RANGE)
        .Cells(1, 1).Value = SUMMARY_PLACEHOLDER
        .Font.Italic = True
        .Font.Color = RGB(128, 128, 128)           ' gray "background" text
    End With
End Sub

Public Sub ClearPlaceholder(ByVal ws As Worksheet)
    With ws.Range(SUMMARY_RANGE)
        .Cells(1, 1).Value = ""
        .Font.Italic = False
        .Font.Color = RGB(0, 0, 0)
    End With
End Sub


'==========================================================================
' HOME TAB INDEX
'==========================================================================
Private Sub WriteIndexRow(ByVal wsHome As Worksheet, ByVal r As Long, ByVal acct As String)
    Dim ref As String
    ref = "'" & acct & "'!B2"

    With wsHome.Cells(r, COL_NUMBER)
        .NumberFormat = "@"
        .Value = acct
        .Errors(xlNumberAsText).Ignore = True     ' hide the "Number Stored as Text" warning
    End With

    ' Account name follows whatever is typed in B2 of the account tab
    wsHome.Cells(r, COL_NAME).Formula = "=IF(" & ref & "="""",""""," & ref & ")"

    wsHome.Cells(r, COL_LINK).Hyperlinks.Delete
    wsHome.Hyperlinks.Add Anchor:=wsHome.Cells(r, COL_LINK), Address:="", _
                          SubAddress:="'" & acct & "'!A1", TextToDisplay:="Go to tab"
End Sub

Private Function NextFreeIndexRow(ByVal wsHome As Worksheet) As Long
    Dim r As Long
    For r = LIST_FIRST_ROW To LIST_LAST_ROW
        If Len(Trim$(CStr(wsHome.Cells(r, COL_NUMBER).Value))) = 0 Then
            NextFreeIndexRow = r
            Exit Function
        End If
    Next r
    NextFreeIndexRow = 0      ' index is full
End Function


'==========================================================================
' HELPERS
'==========================================================================
' Turns a pasted value into a plain string of digits, without spaces.
Private Function NormalizeAccount(ByVal v As Variant) As String
    Dim s As String

    If IsError(v) Or IsEmpty(v) Then
        NormalizeAccount = ""
        Exit Function
    End If

    If VarType(v) = vbDouble Or VarType(v) = vbCurrency Or VarType(v) = vbLong Then
        s = Format$(v, "0")                        ' avoids 1.23E+09 style text
    Else
        s = CStr(v)
    End If

    s = Replace(s, Chr(160), "")                   ' non-breaking spaces from web pastes
    s = Replace(s, " ", "")
    NormalizeAccount = Trim$(s)
End Function

Private Function SheetExists(ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    SheetExists = Not ws Is Nothing
End Function

Private Sub AddDropdown(ByVal target As Range, ByVal items As String)
    With target.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, Formula1:=items
        .IgnoreBlank = True
        .InCellDropdown = True
    End With
End Sub

Private Sub AddButton(ByVal ws As Worksheet, ByVal shapeName As String, ByVal buttonText As String, _
                      ByVal leftPos As Double, ByVal topPos As Double, _
                      ByVal buttonWidth As Double, ByVal buttonHeight As Double, _
                      ByVal macroName As String, Optional ByVal fillColor As Long = -1)
    Dim shp As Shape

    If fillColor = -1 Then fillColor = RGB(0, 112, 192)   ' default: blue

    On Error Resume Next
    ws.Shapes(shapeName).Delete                    ' replace if it already exists
    On Error GoTo 0

    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, leftPos, topPos, buttonWidth, buttonHeight)
    shp.Name = shapeName
    shp.Fill.ForeColor.RGB = fillColor
    shp.Line.Visible = msoFalse
    shp.Placement = xlFreeFloating

    With shp.TextFrame2
        .VerticalAnchor = msoAnchorMiddle
        With .TextRange
            .Text = buttonText
            .Font.Bold = msoTrue
            .Font.Size = 11
            .Font.Fill.ForeColor.RGB = RGB(255, 255, 255)   ' white text
            .ParagraphFormat.Alignment = msoAlignCenter
        End With
    End With

    shp.OnAction = macroName
End Sub

Private Sub RestoreSettings(ByVal prevAutoFill As Boolean)
    Application.AutoCorrect.AutoFillFormulasInLists = prevAutoFill
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

Private Function BuildReport(ByVal nCreated As Long, ByVal nListed As Long, _
                             ByVal msgInvalid As String, ByVal msgDuplicate As String, _
                             ByVal msgAlready As String, ByVal msgNoRoom As String) As String
    Dim s As String

    s = nCreated & " new tab(s) created; " & nListed & " account(s) added to the list."

    If Len(msgAlready) > 0 Then
        s = s & vbCrLf & vbCrLf & "Already processed (no new tab):" & msgAlready
    End If
    If Len(msgDuplicate) > 0 Then
        s = s & vbCrLf & vbCrLf & "Pasted more than once (processed once):" & msgDuplicate
    End If
    If Len(msgInvalid) > 0 Then
        s = s & vbCrLf & vbCrLf & "Skipped - not exactly 10 digits:" & msgInvalid
    End If
    If Len(msgNoRoom) > 0 Then
        s = s & vbCrLf & vbCrLf & "Skipped - the list already holds 25 accounts:" & msgNoRoom
    End If
    If Len(msgInvalid) > 0 Or Len(msgNoRoom) > 0 Then
        s = s & vbCrLf & vbCrLf & "The paste area has been cleared. Re-paste any skipped" & _
                vbCrLf & "accounts once they are corrected."
    End If

    BuildReport = s
End Function
