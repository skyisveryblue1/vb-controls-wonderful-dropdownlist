VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmFlexGrid 
   Caption         =   "Test Grid with dropdown"
   ClientHeight    =   7020
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11415
   LinkTopic       =   "Form1"
   ScaleHeight     =   7020
   ScaleWidth      =   11415
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   360
      TabIndex        =   4
      Top             =   360
      Width           =   2895
   End
   Begin VB.CommandButton cmdGetCheckedData 
      Caption         =   "Get Checked Data"
      Height          =   375
      Left            =   8760
      TabIndex        =   3
      Top             =   1920
      Width           =   1815
   End
   Begin VB.PictureBox picUnchecked 
      Height          =   255
      Left            =   8640
      Picture         =   "frmMFG.frx":0000
      ScaleHeight     =   169
      ScaleMode       =   0  'User
      ScaleWidth      =   169
      TabIndex        =   2
      Top             =   840
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picChecked 
      Height          =   255
      Left            =   8640
      Picture         =   "frmMFG.frx":00EA
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   255
   End
   Begin MSFlexGridLib.MSFlexGrid mfgMain 
      Height          =   3855
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   6800
      _Version        =   393216
   End
End
Attribute VB_Name = "frmFlexGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim gridSuggestion As clsGridWithDropDown
Dim copyPaste As clsCopyPasteExelFlexGrid

Private Sub cmdGetCheckedData_Click()
    Dim strCheckedData As String
    Dim i As Integer
    With mfgMain
        For i = 1 To .Rows - 1
            .row = i
            .col = 0
            If .CellPicture = picChecked Then
                strCheckedData = strCheckedData + .TextMatrix(.row, 1) + vbCrLf
            End If
        Next
    End With
    
    MsgBox strCheckedData
End Sub

Private Sub mfgMain_KeyDown(KeyCode As Integer, Shift As Integer)
    If (KeyCode = &HD Or KeyCode = &HA) Then
        gridSuggestion.EnterEdit mfgMain.row, mfgMain.col, ""
    End If
     
    If Shift And vbCtrlMask And KeyCode = vbKeyC Then
        copyPaste.CopyToClipboard mfgMain
    End If
    
    If Shift And vbCtrlMask And KeyCode = vbKeyV Then
        copyPaste.PasteFromClipboard mfgMain
    End If
    
    If KeyCode = vbKeyDelete Then
        copyPaste.ClearSelection mfgMain
    End If
End Sub
Private Sub mfgMain_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If mfgMain.row <> mfgMain.RowSel Or mfgMain.col <> mfgMain.ColSel Then
        Exit Sub
    End If
    
    gridSuggestion.EnterEdit mfgMain.MouseRow, mfgMain.MouseCol, ""
End Sub
Private Sub mfgMain_KeyPress(KeyAscii As Integer)
    If KeyAscii > 31 Then
        gridSuggestion.EnterEdit mfgMain.row, mfgMain.col, Chr(KeyAscii)
    End If
End Sub
Private Sub Form_Activate()
    Dim strDetailFields(3) As String
    Dim strClipboardText As String
    strDetailFields(0) = "ContactName"
    strDetailFields(1) = "ContactTitle"
    strDetailFields(2) = "Address"
    
    gridSuggestion.InitComponent Me, mfgMain, "CompanyName", strDetailFields, _
      "C:\Program Files\Microsoft Visual Studio\VB98\NWIND.mdb", _
      "SELECT * FROM Customers"

End Sub

Private Sub mfgMain_Click()
    Dim nCurRow As Integer
    With mfgMain
        If .MouseCol = 0 Then
            .col = .MouseCol
            .row = .MouseRow
          
            If .CellPicture = picChecked Then
                Set .CellPicture = picUnchecked
            Else
                Set .CellPicture = picChecked
            End If
           
        End If
    End With
End Sub

Private Sub Form_Load()
    Dim i As Integer
    picChecked.Visible = False
    picUnchecked.Visible = False
    
    With mfgMain
        .Cols = 5
        .Rows = 10
        .ColWidth(0) = 250 ' CheckBox column
        .ColWidth(1) = 2500
        .ColWidth(2) = 2000
        .ZOrder 1
        For i = 1 To .Rows - 1
            .row = i
            .col = 0
            'Align the checkbox
            .CellPictureAlignment = 4
            ' Set the default checkbox picture to the empty box
            Set .CellPicture = picUnchecked.Picture
        Next
    End With
    Set gridSuggestion = New clsGridWithDropDown
    
    Set copyPaste = New clsCopyPasteExelFlexGrid
End Sub




