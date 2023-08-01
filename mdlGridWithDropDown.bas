Attribute VB_Name = "mdlGridWithDropDown"
Option Explicit
Dim frmParent As VB.Form
Public mfgTarget As MSFlexGrid
Public txtSearchWord As VB.TextBox
Public mfgSuggestion As MSFlexGrid
Public lineSeperator As VB.Line
Public lblDetail As VB.Label

Public Sub InitComponent(ByRef frmMain As VB.Form, ByRef mfg As MSFlexGrid)
    Dim i As Integer
    Dim nTotalWidth As Integer
    
    Set frmParent = frmMain
    Set mfgTarget = mfg
    
    Set txtSearchWord = frmMain.Controls.Add("VB.TextBox", "txtSearchWord")
    With txtSearchWord
        .BorderStyle = 0
        .Visible = False
        .ZOrder 0
    End With
    
    Set mfgSuggestion = frmMain.Controls.Add("MSFlexGridLib.MSFlexGrid", "mfgSuggestion")
    With mfgSuggestion
        .Appearance = flexFlat
        .BorderStyle = flexBorderNone
        .BackColorBkg = vbWhite
        .FocusRect = flexFocusNone
        .GridLines = flexGridNone
        .GridLinesFixed = flexGridNone
        .HighLight = flexHighlightAlways
        .ScrollBars = flexScrollBarVertical
        .SelectionMode = flexSelectionByRow
                
        .Font.Name = "Consolas"
        .Font.Size = 13
        .Font.Bold = True
        .FixedCols = 0
        .FixedRows = 0
        .Rows = 1
        .Cols = 6
        .ColWidth(0) = 1500
        .ColWidth(1) = 5000
        .ColWidth(2) = 3000
        .ColWidth(3) = 500
        .ColWidth(4) = 0
        .ColWidth(5) = 0
        .ColAlignment(2) = flexAlignRightCenter
        
        .RowHeight(0) = 0
        .RowHeightMin = 300
        
        For i = 0 To .Cols - 1
            nTotalWidth = nTotalWidth + .ColWidth(i)
        Next i
        .Width = nTotalWidth
        .Height = 3000
        .Visible = False
        .ZOrder 0
    End With
    
    Set lineSeperator = frmMain.Controls.Add("VB.Line", "lineSeperator")
    With lineSeperator
        .BorderWidth = 2
        .BorderColor = &H646464
        .Visible = False
        .ZOrder 0
    End With
    
    Set lblDetail = frmMain.Controls.Add("VB.Label", "lblDetail")
    With lblDetail
        .BackColor = vbWhite
        .ForeColor = vbGrayed
        .Font.Name = "Consolas"
        .Font.Size = 10
        .Visible = False
        .ZOrder 0
    End With
    
End Sub

Public Sub DecideInput()
    If mfgSuggestion.Row = -1 Then
        Exit Sub
    End If
    bSelectText = True
    If (txtSearchWord.Text = mfgSuggestion.TextMatrix(mfgSuggestion.Row, 1)) Then
        txtSearchWord_Change
    Else
        txtSearchWord.Text = mfgSuggestion.TextMatrix(mfgSuggestion.Row, 1)
    End If
    txtSearchWord.SelStart = Len(txtSearchWord.Text)
    txtSearchWord.Visible = False
    
    With mfgMain
        .TextMatrix(.Row, .Col) = txtSearchWord.Text
        If .Col < .Cols - 1 Then
            .Col = .Col + 1
        Else
            If .Row < .Rows - 1 Then
                .Row = .Row + 1
                .Col = 1
            End If
        End If
    End With
End Sub

Public Sub EnterEdit(ByVal selRow As Integer, ByVal selCol As Integer, ByVal strAdd As String)
    With mfgTarget
        .Col = selCol
        .Row = selRow
        If .Row > 0 And .Col > 0 Then
            txtSearchWord.Visible = True
            txtSearchWord.Top = .CellTop + .Top
            txtSearchWord.Left = .CellLeft + .Left
            txtSearchWord.Width = .CellWidth - 20
            txtSearchWord.Height = .CellHeight - 20
            txtSearchWord.Text = .TextMatrix(.Row, .Col) + strAdd
            txtSearchWord.SetFocus
        End If
    End With
End Sub

Public Sub UpdateSuggestion()
    Dim strWhere As String
    If strSearchText = txtSearchWord Then
        'Exit Sub
    End If
    
    strSearchText = txtSearchWord
    strSearchText = Replace(strSearchText, "'", "''")
    strWhere = " AND CompanyName LIKE '*" & strSearchText & "*'"
    
    dataSuggestion.RecordSource = "SELECT * FROM Customers WHERE 1 = 1" & IIf(Len(strSearchText) = 0, "", strWhere)
    dataSuggestion.Refresh
    With mfgSuggestion
        .Rows = dataSuggestion.Recordset.RecordCount
        If .Rows > 0 Then
            .Row = 0
            .Col = 0
            .ColSel = .Cols - 1
        End If
    End With
End Sub

