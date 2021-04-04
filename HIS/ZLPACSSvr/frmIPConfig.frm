VERSION 5.00
Begin VB.Form frmIPConfig 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "允许的接入设备"
   ClientHeight    =   4290
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   7335
   Icon            =   "frmIPConfig.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   7335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin ZlPacsSrv.VsfGrid Vsf 
      Height          =   2625
      Left            =   330
      TabIndex        =   0
      Top             =   750
      Width           =   6675
      _extentx        =   11774
      _extenty        =   4630
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   4710
      TabIndex        =   1
      Top             =   3660
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   5940
      TabIndex        =   2
      Top             =   3660
      Width           =   1100
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "当前授权可接入[]DICOM设备，请在列表中设置允许接入的设备信息"
      Height          =   180
      Left            =   900
      TabIndex        =   3
      Top             =   300
      Width           =   5310
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   270
      Picture         =   "frmIPConfig.frx":000C
      Top             =   180
      Width           =   240
   End
End
Attribute VB_Name = "frmIPConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnOK As Boolean
Private mblnStartUp As Boolean
Private mfrmMain As Form
Private mlngLoop As Long
Private mRs As New ADODB.Recordset
Private mstrSQL As String
Private mintMaxDevs As Integer

Private Function CheckHave(ByVal strIP As String) As Boolean
    '-----------------------------------------------------------------------------------------
    '功能:
    '参数:
    '-----------------------------------------------------------------------------------------
    For mlngLoop = 1 To Vsf.Rows - 1
        If UCase(Vsf.TextMatrix(mlngLoop, 1)) = UCase(strIP) And Vsf.Row <> mlngLoop Then
            CheckHave = True
            Exit Function
        End If
    Next
End Function

Private Function FillGrid(ByRef objMsf As Object, ByVal rsData As ADODB.Recordset, Optional ByVal MaskArray As Variant, Optional ByVal blnClear As Boolean = True) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------
    '功能:填充数据到网格
    '参数:
    '返回:
    '---------------------------------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long
    Dim strMask As String
    Dim lngRow As Long
    
    If blnClear Then
        objMsf.Rows = 2
        objMsf.RowData(1) = 0
        For lngLoop = 0 To objMsf.Cols - 1
            objMsf.TextMatrix(1, lngLoop) = ""
        Next
    End If
    
    lngRow = 0
    Do While Not rsData.EOF
        
        lngRow = lngRow + 1
        If objMsf.Rows < lngRow + 1 Then objMsf.Rows = lngRow + 1
        
        On Error GoTo errHand
        For lngLoop = 0 To objMsf.Cols - 1
        
            On Error Resume Next
            strMask = ""
            strMask = MaskArray(lngLoop)
                                    
            On Error GoTo errHand
            If strMask <> "" Then
                objMsf.TextMatrix(lngRow, lngLoop) = Format(zlCommFun.Nvl(rsData(objMsf.TextMatrix(0, lngLoop))), strMask)
            Else
                objMsf.TextMatrix(lngRow, lngLoop) = zlCommFun.Nvl(rsData(objMsf.TextMatrix(0, lngLoop)))
            End If
                        
        Next
        
        rsData.MoveNext
    Loop
    
    FillGrid = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function ShowEdit(ByVal frmMain As Form, Optional ByVal iMaxDevs As Integer = 2) As Boolean
    '------------------------------------------------------------------------------------------------------
    '功能：
    '------------------------------------------------------------------------------------------------------
    Dim strTmp As String
    
    mblnStartUp = True
    mblnOK = False
    
    mintMaxDevs = iMaxDevs
    strTmp = Switch(mintMaxDevs = -1, "任意", mintMaxDevs > 0, mintMaxDevs) & "台"
    lblNote.Caption = Replace(lblNote.Caption, "[]", strTmp)
    Set mfrmMain = frmMain
    
    If InitData = False Then
        cmdOK.Tag = ""
        Exit Function
    End If
    
    If ReadData = False Then
        cmdOK.Tag = ""
        Exit Function
    End If
    
    cmdOK.Tag = ""
    
    Me.Show 1, frmMain
    
    ShowEdit = mblnOK
    
End Function

Private Function InitData() As Boolean
    '------------------------------------------------------------------------------------------------------
    '功能：
    '------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    
    On Error GoTo errHand
    
    With Vsf
        .Cols = 0
        .NewColumn "ID", 0
        .NewColumn "IP地址", 2500, 1, , 1, 20
        .NewColumn "设备名称", 2500, 1, , 1, 100
        .NewColumn "影像类别", 1000, 1, , 1, 20
        .FixedCols = 1
        .ColHidden(0) = True
    End With
        
    InitData = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Private Function ReadData() As Boolean
    '------------------------------------------------------------------------------------------------------
    '功能：
    '------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    On Error GoTo errHand
    
    mstrSQL = "SELECT 接入ID As ID,IP地址,设备名称,影像类别 FROM 影像接入设备"
                
    Call zlDatabase.OpenRecordset(rs, mstrSQL, Me.Caption)
    If rs.BOF = False Then
        Call FillGrid(Vsf, rs)
    End If
    
    ReadData = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Private Function SaveData() As Boolean
    '------------------------------------------------------------------------------------------------------
    '功能：
    '------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    
    On Error GoTo errHand
    strSQL = ""
    For mlngLoop = 1 To Vsf.Rows - 1
        If Len(Trim(Vsf.TextMatrix(mlngLoop, 1))) > 0 Then
            strSQL = strSQL & "|" & Trim(Vsf.TextMatrix(mlngLoop, 1)) & "^" & Trim(Vsf.TextMatrix(mlngLoop, 2)) & "^" & Trim(Vsf.TextMatrix(mlngLoop, 3))
        End If
    Next
    If Len(strSQL) > 0 Then strSQL = Mid(strSQL, 2)
    
    gstrSQL = "zl_影像接入设备_SAVE('" & strSQL & "')"
    ExecuteProcedure Me.Caption
    
    SaveData = True
    
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdOK_Click()
    Dim strTmp As String
    
    If cmdOK.Tag <> "" Then
        
        If SaveData = False Then Exit Sub
        
        mblnOK = True
        
    End If
    
    cmdOK.Tag = ""
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If cmdOK.Tag <> "" Then
        Cancel = (MsgBox("新增或修改的设置必须保存后才生效，是否不保存就退出？", vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) = vbNo)
    End If
End Sub

Private Sub vsf_AfterDeleteRow(ByVal Row As Long, ByVal Col As Long)
    cmdOK.Tag = "Changed"
End Sub

Private Sub vsf_BeforeNewRow(ByVal Row As Long, Col As Long, Cancel As Boolean)
    If Len(Trim(Vsf.TextMatrix(Row, 1))) = 0 Then Cancel = True
    If Row >= mintMaxDevs And mintMaxDevs > 0 Then Cancel = True
End Sub

Private Sub vsf_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, ByVal ComboList As String, KeyCode As Integer, ByVal Shift As Integer, Cancel As Boolean)
    If KeyCode <> vbKeyReturn Then cmdOK.Tag = "Changed"
End Sub

Private Sub vsf_KeyPress(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer, Cancel As Boolean)
    If KeyAscii = vbKeyReturn Then
        If (Col = 1 And Trim(Vsf.TextMatrix(Row, Col)) = "") Or (Col = 2 And Row = mintMaxDevs And mintMaxDevs > 0) Then
            zlCommFun.PressKey vbKeyTab
            Cancel = True
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub Vsf_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = Asc("|") Or KeyAscii = Asc("^") Then KeyAscii = 0
End Sub

Private Sub vsf_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If CheckHave(Vsf.EditText) And Col = 1 Then Cancel = True
End Sub
