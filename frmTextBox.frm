VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmTextBox 
   Caption         =   "Test Textbox with dropdown"
   ClientHeight    =   7530
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10905
   DrawStyle       =   5  'Transparent
   LinkTopic       =   "Form1"
   ScaleHeight     =   7530
   ScaleWidth      =   10905
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtOther 
      Height          =   285
      Left            =   7920
      TabIndex        =   1
      Top             =   480
      Width           =   2535
   End
   Begin VB.TextBox txtSearchWord 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   6375
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmTextBox.frx":0000
      Height          =   1575
      Left            =   0
      OleObjectBlob   =   "frmTextBox.frx":0014
      TabIndex        =   2
      Top             =   5160
      Visible         =   0   'False
      Width           =   7335
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   7560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6720
      Visible         =   0   'False
      Width           =   2775
   End
   Begin MSFlexGridLib.MSFlexGrid mfgSuggestion 
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   3201
      _Version        =   393216
      Rows            =   10
      FixedRows       =   0
      FixedCols       =   0
      BackColorBkg    =   16777215
      AllowBigSelection=   -1  'True
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      FillStyle       =   1
      GridLines       =   0
      GridLinesFixed  =   0
      ScrollBars      =   2
      SelectionMode   =   1
      Appearance      =   0
   End
   Begin VB.Label Label1 
      Caption         =   "Company Name:"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   2535
   End
   Begin VB.Line lineSeperator 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      X1              =   120
      X2              =   7320
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Label lblDetail 
      Height          =   1215
      Left            =   120
      TabIndex        =   4
      Top             =   2880
      Width           =   7095
   End
End
Attribute VB_Name = "frmTextBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private strSearchText As String, temp As String
Private i As Integer, j As Integer
Private bSelectText As Boolean

Private Sub txtSearchWord_Change()
    If bSelectText = True Then
        bSelectText = False
        txtSearchWord_LostFocus
        Exit Sub
    End If
    UpdateSuggestion
    mfgSuggestion_RowColChange
End Sub

Private Sub UpdateSuggestion()
    Dim strWhere As String
    If strSearchText = txtSearchWord Then
        'Exit Sub
    End If
    
    strSearchText = txtSearchWord
    strSearchText = Replace(strSearchText, "'", "''")
    strWhere = " AND CompanyName LIKE '*" & strSearchText & "*'"
    
    Data1.RecordSource = "SELECT * FROM Customers WHERE 1 = 1" & IIf(Len(strSearchText) = 0, "", strWhere)
    Data1.Refresh
    With mfgSuggestion
        .Rows = Data1.Recordset.RecordCount + 1
        For i = 1 To Data1.Recordset.RecordCount
            .TextMatrix(i, 0) = Data1.Recordset.Fields("CustomerID").Value
            .TextMatrix(i, 1) = Data1.Recordset.Fields("CompanyName").Value
            Data1.Recordset.MoveNext
            If Data1.Recordset.EOF = True Then
               Exit For
            End If
        Next
        If .Rows > 0 Then
            .Row = 0
            .Col = 0
            .ColSel = .Cols - 1
        End If
    End With
End Sub
Private Sub Form_Load()
    Dim nTotalWidth As Integer
    Data1.DatabaseName = "C:\Program Files\Microsoft Visual Studio\VB98\NWIND.mdb"
    
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
        
    End With
    
    With lineSeperator
        .BorderWidth = 2
        .BorderColor = &H646464
        .Visible = False
    End With
    
    With lblDetail
        .BackColor = vbWhite
        .ForeColor = vbGrayed
        .Font.Name = "Consolas"
        .Font.Size = 10
        .Visible = False
    End With
    
    bSelectText = False
End Sub

Private Sub Data1_DataChanged()
    mfgSuggestion.Refresh
End Sub

Private Sub mfgSuggestion_RowColChange()
    With mfgSuggestion
        If .Row < 0 Or .Rows = 0 Then
            lblDetail.Caption = ""
            Exit Sub
        End If
    
        lblDetail.Caption = _
            "Contact Name: " & .TextMatrix(.Row, 2) & vbCrLf & _
            "Contact Title: " & .TextMatrix(.Row, 3) & vbCrLf & _
            "Contact Address: " & .TextMatrix(.Row, 4) & vbCrLf
    End With
End Sub

Private Sub txtSearchWord_Click()
    txtSearchWord_GotFocus
End Sub

Private Sub txtSearchWord_GotFocus()
    With mfgSuggestion
        .Top = txtSearchWord.Top + txtSearchWord.Height
        .Left = txtSearchWord.Left
        .Visible = True
    End With
    
    With lineSeperator
        .X1 = mfgSuggestion.Left
        .Y1 = mfgSuggestion.Top + mfgSuggestion.Height + 20
        .X2 = .X1 + mfgSuggestion.Width
        .Y2 = .Y1
        .Visible = True
    End With
    
    With lblDetail
        .Left = mfgSuggestion.Left
        .Width = mfgSuggestion.Width
        .Top = lineSeperator.Y1 + 20
        .Visible = True
    End With
    
    UpdateSuggestion
    mfgSuggestion_RowColChange
    
End Sub
Private Sub txtSearchWord_LostFocus()
    mfgSuggestion.Visible = False
    lineSeperator.Visible = False
    lblDetail.Visible = False
End Sub
Private Sub txtSearchWord_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then Exit Sub
    
    If KeyAscii <= 31 Then
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Sub txtSearchWord_KeyDown(KeyCode As Integer, Shift As Integer)
    If (KeyCode = vbKeyDown) Or (KeyCode = vbKeyUp) Then
        With mfgSuggestion
            Select Case KeyCode
                Case vbKeyDown:
                    If Not (.Row = .Rows - 1) Then
                        .Row = .Row + 1
                    End If
                Case vbKeyUp:
                    If Not (.Row < 1) Then
                        .Row = .Row - 1
                    End If
            End Select
            
            .ColSel = .Cols - 1
            '.TopRow = .Row
            If KeyCode = vbKeyUp Then
                ' Scroll up
                If .TopRow > 0 Then
                    .TopRow = .TopRow - 1
                End If
            ElseIf KeyCode = vbKeyDown Then
                ' Scroll down
                If .Row < .Rows - 1 Then
                    .TopRow = .TopRow + 1
                End If
            End If
        End With
    End If
    If (KeyCode = &HD Or KeyCode = &HA) Then
        bSelectText = True
        If (txtSearchWord.Text = mfgSuggestion.TextMatrix(mfgSuggestion.Row, 1)) Then
            txtSearchWord_Change
        Else
            txtSearchWord.Text = mfgSuggestion.TextMatrix(mfgSuggestion.Row, 1)
        End If
        txtSearchWord.SelStart = Len(txtSearchWord.Text)
    End If
    If KeyCode = vbKeyUp Then
        KeyCode = 0 'set KeyCode to 0 to prevent the "Up arrow" key from resetting the cursor position
    End If
End Sub


