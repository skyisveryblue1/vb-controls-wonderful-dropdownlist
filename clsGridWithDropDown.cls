VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGridWithDropDown"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_frmParent As VB.Form
Private m_mfgTarget As MSFlexGrid
Attribute m_mfgTarget.VB_VarHelpID = -1
Private WithEvents m_txtSearchWord As VB.TextBox
Attribute m_txtSearchWord.VB_VarHelpID = -1
Private WithEvents m_mfgSuggestion As MSFlexGrid
Attribute m_mfgSuggestion.VB_VarHelpID = -1
Private m_lineSeperator As VB.Line
Private m_lblDetail As VB.Label
Private WithEvents m_dataMain As VB.Data
Attribute m_dataMain.VB_VarHelpID = -1
Private m_strEnterSelField As String, m_nEnterSelCol As Integer
Private m_strDetailFields() As String, m_nDetailFields() As Integer
Private m_strSearchText As String
Private m_bSelectText As Boolean


Private Sub m_dataMain_ObjectEvent(Info As EventInfo)
    Select Case Info.Name
    Case "DataChanged"
        m_mfgSuggestion.Refresh
    Case Else
    End Select
End Sub
Private Sub m_txtSearchWord_ObjectEvent(Info As EventInfo)
    With m_mfgSuggestion
        Select Case Info.Name
        Case "Change"
            m_txtSearchWord_Change
        Case "Click"
            m_txtSearchWord_Click
        Case "GotFocus"
            m_txtSearchWord_GotFocus
        Case "LostFocus"
            m_txtSearchWord_LostFocus
        Case Else ' Unknown Event
        End Select
    End With
End Sub

Private Sub m_txtSearchWord_Change()
    If m_bSelectText = True Then
        m_bSelectText = False
        m_txtSearchWord_LostFocus
        Exit Sub
    End If
    UpdateSuggestion
    m_mfgSuggestion_RowColChange
End Sub

Private Sub m_txtSearchWord_Click()
    m_txtSearchWord_GotFocus
End Sub

Private Sub m_txtSearchWord_GotFocus()
    With m_mfgSuggestion
        .Top = m_txtSearchWord.Top + m_txtSearchWord.Height
        .Left = m_txtSearchWord.Left
        .Visible = True
    End With
    
    With m_lineSeperator
        .X1 = m_mfgSuggestion.Left
        .Y1 = m_mfgSuggestion.Top + m_mfgSuggestion.Height + 20
        .X2 = .X1 + m_mfgSuggestion.Width
        .Y2 = .Y1
        .Visible = True
    End With
    
    With m_lblDetail
        .Left = m_mfgSuggestion.Left
        .Width = m_mfgSuggestion.Width
        .Top = m_lineSeperator.Y1 + 20
        .Visible = True
    End With
    
    UpdateSuggestion
    m_mfgSuggestion_RowColChange
    
End Sub
Private Sub m_txtSearchWord_LostFocus()
    m_mfgSuggestion.Visible = False
    m_lineSeperator.Visible = False
    m_lblDetail.Visible = False
End Sub
Private Sub m_txtSearchWord_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then Exit Sub
    
    If KeyAscii <= 31 Then
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Sub m_txtSearchWord_KeyDown(KeyCode As Integer, Shift As Integer)
    If (KeyCode = vbKeyDown) Or (KeyCode = vbKeyUp) Then
        With m_mfgSuggestion
            Select Case KeyCode
                Case vbKeyDown:
                    If .Row < .Rows - 1 Then
                        .Row = .Row + 1
                        If .TopRow < .Rows - 1 Then
                            .TopRow = .TopRow + 1
                        End If
                    End If
                Case vbKeyUp:
                    If .Row > 0 Then
                        .Row = .Row - 1
                        If .TopRow > 0 Then
                            .TopRow = .TopRow - 1
                        End If
                    End If
            End Select
            
            .ColSel = .Cols - 1
        End With
    End If
    If (KeyCode = &HD Or KeyCode = &HA) Then
        DecideInput
    End If
    If KeyCode = vbKeyUp Then
        KeyCode = 0 'set KeyCode to 0 to prevent the "Up arrow" key from resetting the cursor position
    End If
End Sub

Private Sub m_mfgSuggestion_ObjectEvent(Info As EventInfo)
    MsgBox Info.Name

    With m_mfgSuggestion
        Select Case Info.Name
        Case "Click"
            m_mfgSuggestion_Click
        Case "RowColChange"
            m_mfgSuggestion_RowColChange
        Case Else ' Unknown Event
        End Select
    End With
End Sub
Private Sub m_mfgSuggestion_Click()
    With m_mfgSuggestion
        .Row = .MouseRow
        .ColSel = .Cols - 1
        DecideInput
    End With
End Sub

Private Sub m_mfgSuggestion_RowColChange()
    Dim i As Integer
    With m_mfgSuggestion
        If .Row < 0 Or .Rows = 0 Then
            m_lblDetail.Caption = ""
            Exit Sub
        End If
    
        m_lblDetail.Caption = ""
        
        For i = 0 To UBound(m_strDetailFields) - 1
            m_lblDetail.Caption = m_lblDetail.Caption & PadString(m_strDetailFields(i), 15) & _
                ": " & .TextMatrix(.Row, m_nDetailFields(i)) & vbCrLf
        Next
    End With
End Sub

Public Sub InitComponent(ByRef frmParent As VB.Form, ByRef mfgTarget As MSFlexGrid, ByVal strSelField As String, _
        strDetailFields() As String, ByVal strDatabase As String, ByVal strRecordSource As String)
    Dim i As Integer, j As Integer
    Dim nTotalWidth As Integer
    Dim strFieldName As String
    ReDim m_nDetailFields(UBound(strDetailFields)) As Integer
    
    Set m_frmParent = frmParent
    Set m_mfgTarget = mfgTarget
    
    Set m_dataMain = m_frmParent.Controls.Add("VB.Data", "m_dataMain")
    With m_dataMain
        .Connect = "Access"
        .DefaultType = 2
        .DatabaseName = strDatabase
        .RecordSource = strRecordSource
        .Refresh
    End With
    
    m_strDetailFields = strDetailFields
        
    m_strEnterSelField = strSelField
    For i = 0 To m_dataMain.Recordset.Fields.Count - 1
        strFieldName = m_dataMain.Recordset.Fields(i).Name
        If strFieldName = strSelField Then
            m_nEnterSelCol = i
        End If
        
        For j = 0 To UBound(m_strDetailFields) - 1
            If strFieldName = m_strDetailFields(j) Then
                m_nDetailFields(j) = i
            End If
        Next
    Next
    
    Set m_txtSearchWord = m_frmParent.Controls.Add("VB.TextBox", "m_txtSearchWord")
    With m_txtSearchWord
        .BorderStyle = 0
        .Visible = False
        .ZOrder 0
    End With
    
    Set m_mfgSuggestion = m_frmParent.Controls.Add("MSFlexGridLib.MSFlexGrid", "m_mfgSuggestion")
    
    With m_mfgSuggestion
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
        .Rows = 2
        .Cols = 6
        .ColWidth(0) = 1500
        .ColWidth(1) = 5000
        .ColWidth(2) = 3000
        .ColWidth(3) = 500
        .ColWidth(4) = 0
        .ColWidth(5) = 0
        .ColAlignment(2) = flexAlignRightCenter
        
        .RowHeightMin = 300
        
        For i = 0 To .Cols - 1
            nTotalWidth = nTotalWidth + .ColWidth(i)
        Next i
        .Width = nTotalWidth
        .Height = 3000
        .Visible = False
        .ZOrder 0
        .Refresh
    End With
    
    Set m_lineSeperator = m_frmParent.Controls.Add("VB.Line", "m_lineSeperator")
    With m_lineSeperator
        .BorderWidth = 2
        .BorderColor = &H646464
        .Visible = False
        .ZOrder 0
    End With
    
    Set m_lblDetail = m_frmParent.Controls.Add("VB.Label", "m_lblDetail")
    With m_lblDetail
        .BackColor = vbWhite
        .ForeColor = vbGrayed
        .Font.Name = "Consolas"
        .Font.Size = 10
        .Height = 1000
        .Visible = False
        .ZOrder 0
    End With

    m_bSelectText = False
End Sub

Public Sub DecideInput()
    With m_mfgSuggestion
        If .Row = -1 Then
            Exit Sub
        End If
        m_bSelectText = True
        If (m_txtSearchWord.Text = .TextMatrix(.Row, m_nEnterSelCol)) Then
            m_txtSearchWord_Change
        Else
            m_txtSearchWord.Text = .TextMatrix(.Row, m_nEnterSelCol)
        End If
        m_txtSearchWord.SelStart = Len(m_txtSearchWord.Text)
        m_txtSearchWord.Visible = False
    End With
    
    With m_mfgTarget
        .TextMatrix(.Row, .Col) = m_txtSearchWord.Text
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
    With m_mfgTarget
        .Col = selCol
        .Row = selRow
        If .Row > 0 And .Col > 0 Then
            m_txtSearchWord.Visible = True
            m_txtSearchWord.Top = .CellTop + .Top
            m_txtSearchWord.Left = .CellLeft + .Left
            m_txtSearchWord.Width = .CellWidth - 20
            m_txtSearchWord.Height = .CellHeight - 20
            m_txtSearchWord.Text = .TextMatrix(.Row, .Col) + strAdd
            m_txtSearchWord.SetFocus
        End If
        m_txtSearchWord.SelStart = Len(m_txtSearchWord.Text)
    End With
End Sub

Public Sub UpdateSuggestion()
    Dim i As Integer, j As Integer
    Dim strWhere As String
    If m_strSearchText = m_txtSearchWord Then
        'Exit Sub
    End If
    
    m_strSearchText = m_txtSearchWord
    m_strSearchText = Replace(m_strSearchText, "'", "''")
    strWhere = " AND " & m_strEnterSelField & " LIKE '*" & m_strSearchText & "*'"
    
    m_dataMain.RecordSource = "SELECT * FROM Customers WHERE 1 = 1" & IIf(Len(m_strSearchText) = 0, "", strWhere)
    m_dataMain.Refresh
    
    With m_mfgSuggestion
         If m_dataMain.Recordset.RecordCount = 0 Then
            .Rows = 0
            Exit Sub
        End If
        m_dataMain.Recordset.MoveLast
        .Rows = m_dataMain.Recordset.RecordCount
        .Cols = m_dataMain.Recordset.Fields.Count
        
        m_dataMain.Recordset.MoveFirst
        For i = 0 To .Rows - 1
            For j = 0 To .Cols - 1
                If IsNull(m_dataMain.Recordset.Fields(j).Value) = False Then
                    .TextMatrix(i, j) = m_dataMain.Recordset.Fields(j).Value
                End If
            Next
            m_dataMain.Recordset.MoveNext
            If m_dataMain.Recordset.EOF() Then
                Exit For
            End If
        Next

        .Row = 0
        .TopRow = 0
        .Col = 0
        .ColSel = .Cols - 1
    End With
End Sub

Public Function PadString(strSource As String, lPadLen As Long, Optional PadChar As String = " ") As String
    PadString = String(lPadLen, PadChar)
    Mid(PadString, 1, lPadLen) = strSource
End Function

