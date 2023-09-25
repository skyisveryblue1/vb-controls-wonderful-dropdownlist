VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCopyPasteExelFlexGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function CloseClipboard Lib "user32" () As Long
Private Declare Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As Long
Private Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function EmptyClipboard Lib "user32" () As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Const CF_TEXT As Long = 1
Private Const GMEM_SHARE = &H2000&
Private Const GMEM_MOVEABLE = &H2
Private Const GMEM_ZEROINIT = &H40
Private Const FOR_CLIPBOARD = GMEM_MOVEABLE Or GMEM_SHARE Or GMEM_ZEROINIT

Public Sub CopyToClipboard(ByRef mfg As MSFlexGrid)
    Dim strSelectedText As String, strLine As String, strCell As String
    Dim nStartRow As Integer, nStartCol As Integer
    Dim nEndRow As Integer, nEndCol As Integer
    Dim nRow As Integer, nCol As Integer
    
    nStartRow = mfg.RowSel
    nStartCol = mfg.ColSel
    nEndRow = mfg.row
    nEndCol = mfg.col

    If nStartRow > nEndRow Then
        Swap nStartRow, nEndRow
    End If

    If nStartCol > nEndCol Then
        Swap nStartCol, nEndCol
    End If
    
    For nRow = nStartRow To nEndRow
        strLine = ""
        For nCol = nStartCol To nEndCol
            If nCol <> nStartCol Then
                strLine = strLine & vbTab
            End If
            strCell = mfg.TextMatrix(nRow, nCol)
            strLine = strLine & IIf(strCell = "", " ", strCell)
        Next nCol
        If strSelectedText <> "" Then
            strSelectedText = strSelectedText & vbCrLf
        End If
        strSelectedText = strSelectedText & strLine
    Next nRow

    CopyTextToClipboard strSelectedText
End Sub

Public Sub PasteFromClipboard(ByRef mfg As MSFlexGrid)
    
    Dim sText As String
    Dim lines() As String
    Dim columns() As String
    
    Dim nStartRow As Integer, nStartCol As Integer
    Dim nEndRow As Integer, nEndCol As Integer
    Dim nRow As Integer, nCol As Integer
    
    nStartRow = mfg.RowSel
    nStartCol = mfg.ColSel
    nEndRow = mfg.row
    nEndCol = mfg.col

    If nStartRow > nEndRow Then
        Swap nStartRow, nEndRow
    End If

    If nStartCol > nEndCol Then
        Swap nStartCol, nEndCol
    End If

    sText = GetTextFromClipboard
    sText = Replace(sText, vbNewLine, vbCr)
    lines = Split(sText, vbCr)
    For nRow = 0 To UBound(lines)
        columns = Split(lines(nRow), vbTab)
        If nRow + nStartRow > nEndRow Then Exit For
        For nCol = 0 To UBound(columns)
            If nCol + nStartCol > nEndCol Then Exit For
            mfg.TextMatrix(nRow + nStartRow, nCol + nStartCol) = columns(nCol)
        Next nCol
    Next nRow
End Sub

Private Sub CopyTextToClipboard(ByVal sText As String)
    Dim hMem As Long, pMem As Long

    hMem = GlobalAlloc(FOR_CLIPBOARD, LenB(sText) + 1)
    pMem = GlobalLock(hMem)
    CopyMemory ByVal pMem, ByVal sText, LenB(sText)
    GlobalUnlock hMem
    
    If OpenClipboard(0&) = 0 Then
       MsgBox "Clipboard opened by another application."
    Else
       EmptyClipboard
       SetClipboardData CF_TEXT, hMem
       CloseClipboard
    End If
End Sub

Private Function GetTextFromClipboard() As String
    Dim hClipMemory As Long
    Dim lpClipMemory As Long
    Dim sText As String
    Dim lMemSize As Long

    ' Open the clipboard
    If OpenClipboard(0) = 0 Then
        MsgBox "Clipboard open by another application.", vbExclamation
        GetTextFromClipboard = ""
    End If

    ' Get a handle to the clipboard data
    hClipMemory = GetClipboardData(CF_TEXT)

    If hClipMemory <> 0 Then
        lMemSize = GlobalSize(hClipMemory)
        ' Lock the memory and retrieve the clipboard text
        lpClipMemory = GlobalLock(hClipMemory)
        sText = String$(lMemSize, 0)
        CopyMemory ByVal sText, ByVal lpClipMemory, lMemSize

        ' Unlock the memory
        GlobalUnlock hClipMemory
    End If

    ' Close the clipboard
    CloseClipboard

    GetTextFromClipboard = sText
End Function
Public Sub ClearSelection(ByRef mfg As MSFlexGrid)
    Dim nStartRow As Integer, nStartCol As Integer
    Dim nEndRow As Integer, nEndCol As Integer
    Dim nRow As Integer, nCol As Integer
    
    nStartRow = mfg.RowSel
    nStartCol = mfg.ColSel
    nEndRow = mfg.row
    nEndCol = mfg.col

    If nStartRow > nEndRow Then
        Swap nStartRow, nEndRow
    End If

    If nStartCol > nEndCol Then
        Swap nStartCol, nEndCol
    End If
    For nRow = nStartRow To nEndRow
        For nCol = nStartCol To nEndCol
            mfg.TextMatrix(nRow, nCol) = ""
        Next nCol
    Next nRow
End Sub

Private Sub Swap(ByRef a As Integer, ByRef b As Integer)
    Dim temp As Integer
    temp = a
    a = b
    b = temp
End Sub

