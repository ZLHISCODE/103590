VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTendBodyArrage 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "������Ŀ����"
   ClientHeight    =   4830
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   8430
   Icon            =   "frmTendBodyArrage.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   8430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdDown 
      Caption         =   "����(&D)"
      Height          =   350
      Left            =   7260
      TabIndex        =   4
      Top             =   1380
      Width           =   1100
   End
   Begin VB.CommandButton cmdUp 
      Caption         =   "����(&U)"
      Height          =   350
      Left            =   7260
      TabIndex        =   3
      Top             =   960
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   7260
      TabIndex        =   2
      Top             =   510
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   7260
      TabIndex        =   1
      Top             =   90
      Width           =   1100
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   7440
      Top             =   2940
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTendBodyArrage.frx":000C
            Key             =   "User"
         EndProperty
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid vsf 
      Height          =   4725
      Left            =   45
      TabIndex        =   0
      Top             =   60
      Width           =   7065
      _cx             =   12462
      _cy             =   8334
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   270
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
End
Attribute VB_Name = "frmTendBodyArrage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'######################################################################################################################
'�ֲ�������������


Private mblnStartUp As Boolean
Private mblnOK As Boolean
Private mblnDataChanged As Boolean

Private mfrmMain As Form
Private mstrSQL As String

Private Enum mCol
    ��¼��
    ��¼��
    ��¼��
    ��¼ɫ
    ��Сֵ
    ���ֵ
    ��λֵ
    ��λ
    �����
    ��ɫ
End Enum

'######################################################################################################################
'�Զ��庯������������
Private Property Let DataChanged(ByVal vData As Boolean)
    mblnDataChanged = vData
End Property

Private Property Get DataChanged() As Boolean
    DataChanged = mblnDataChanged
End Property

Public Function ShowEdit(ByVal frmMain As Form) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:��/��ʾ�༭����,���������������(��ں���)
    '------------------------------------------------------------------------------------------------------------------
    mblnStartUp = True
    mblnOK = False
    
    Set mfrmMain = frmMain

    If InitData = False Then GoTo errHand
    If ReadData = False Then GoTo errHand
    
    vsf.Row = 1
    
    Call SetCmdButtonEnable
    
    DataChanged = False
    
    Me.Show 1, frmMain
    
    ShowEdit = mblnOK
    
    Exit Function
    
errHand:
    On Error Resume Next
    DataChanged = False
    Unload Me
End Function

Private Function ReadData() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:��ȡ�������ϣ��Թ���ʾ
    '------------------------------------------------------------------------------------------------------------------
    Dim RS As New ADODB.Recordset
    Dim lngLoop As Long
    Dim objItem As Object
    
    On Error GoTo errHand

        
    mstrSQL = "Select ��Ŀ��� As ID,��¼��,Decode(��¼��,2,'���','����') As ��¼��,��¼��,��¼ɫ,��Сֵ,���ֵ,��λֵ,��λ,����� From ���¼�¼��Ŀ A Order By �������"
    Set RS = zlDatabase.OpenSQLRecord(mstrSQL, Me.Caption)
    If RS.BOF = False Then
        
        With vsf
            Do While Not RS.EOF
                
                If Val(.RowData(.Rows - 1)) > 0 Then .Rows = .Rows + 1
                
                .RowData(.Rows - 1) = zlCommFun.NVL(RS("ID"), 0)
                
                .TextMatrix(.Rows - 1, mCol.��¼��) = zlCommFun.NVL(RS("��¼��"))
                .TextMatrix(.Rows - 1, mCol.��¼��) = zlCommFun.NVL(RS("��¼��"))
                .TextMatrix(.Rows - 1, mCol.��¼��) = zlCommFun.NVL(RS("��¼��"))
                
                '������ɫ
                On Error Resume Next
                Set objItem = Nothing
                Set objItem = ils16.ListImages("K" & NVL(RS("��¼ɫ"), 0))
                On Error GoTo 0
                
                If objItem Is Nothing Then Call SetColorIcon(Me, "K" & NVL(RS("��¼ɫ"), 0), NVL(RS("��¼ɫ"), 0), ils16)
                Set .Cell(flexcpPicture, .Rows - 1, mCol.��¼ɫ) = ils16.ListImages("K" & NVL(RS("��¼ɫ"), 0)).Picture
                .Cell(flexcpPictureAlignment, .Rows - 1, mCol.��¼ɫ) = flexAlignCenterCenter
                
                .TextMatrix(.Rows - 1, mCol.��Сֵ) = Zero(zlCommFun.NVL(RS("��Сֵ")))
                .TextMatrix(.Rows - 1, mCol.���ֵ) = Zero(zlCommFun.NVL(RS("���ֵ")))
                .TextMatrix(.Rows - 1, mCol.��λֵ) = Zero(zlCommFun.NVL(RS("��λֵ")))
                .TextMatrix(.Rows - 1, mCol.��λ) = zlCommFun.NVL(RS("��λ"))
                .TextMatrix(.Rows - 1, mCol.�����) = Zero(zlCommFun.NVL(RS("�����")))
                .TextMatrix(.Rows - 1, mCol.��ɫ) = zlCommFun.NVL(RS("��¼ɫ"), 0)
                
                RS.MoveNext
            Loop
        End With
        
    End If
    
    ReadData = True
    
    Exit Function
    
errHand:
    
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function SetCmdButtonEnable() As Boolean

    
    cmdUp.Enabled = (vsf.Row > 1)
    cmdDown.Enabled = (vsf.Row < vsf.Rows - 1)
    
    
    SetCmdButtonEnable = True
    
End Function


Private Function SaveData(ByRef lngKey As Long) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '���ܣ������޸Ļ�����������
    '���أ��ɹ����淵��True�����򷵻�False
    '------------------------------------------------------------------------------------------------------------------
    Dim blnTran As Boolean
    Dim lngLoop As Long
    Dim strSQL() As String
        
    On Error GoTo errHand
    
    ReDim Preserve strSQL(1 To 1)
    
    For lngLoop = 1 To vsf.Rows - 1
        
        If Val(vsf.RowData(lngLoop)) > 0 Then
            strSQL(ReDimArray(strSQL)) = "ZL_���¼�¼��Ŀ_ARRAGE(" & Val(vsf.RowData(lngLoop)) & "," & lngLoop & ")"
        End If
    Next
    
    'ִ��
    blnTran = True
    gcnOracle.BeginTrans
    For lngLoop = 1 To UBound(strSQL)
        If strSQL(lngLoop) <> "" Then Call zlDatabase.ExecuteProcedure(strSQL(lngLoop), Me.Caption)
    Next
    gcnOracle.CommitTrans
    blnTran = False
    
    SaveData = True
    
    Exit Function
    
errHand:
    '������
    
    If blnTran Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
    
    
End Function

Private Function InitData() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '���ܣ���ʼ�����ݣ�һ��ָ�ؼ������ݳ�ʼ��
    '------------------------------------------------------------------------------------------------------------------
    Dim strVsf As String
    
    On Error GoTo errHand
    
    strVsf = "��¼��,1350,1,1,1,;��¼��,720,1,1,1,;��¼��,720,1,1,1,;��¼ɫ,720,1,1,1,;��Сֵ,720,7,1,1,;���ֵ,720,7,1,1,;��λֵ,720,7,1,1,;��λ,600,1,1,1,;�����,720,7,1,1,;��ɫ,0,1,1,0,"
    Call CreateVsf(vsf, strVsf)
    With vsf
        .Cols = .Cols + 1
        .ColWidth(vsf.Cols - 1) = 15
    End With
    vsf.Rows = 2
    vsf.ColFormat(mCol.��λֵ) = "0.0"
    
    InitData = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    
End Function


'######################################################################################################################
'�ؼ�������ȶ�������ԡ����̡��¼�����������


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100))
End Sub

Private Sub cmdDown_Click()
    
    Dim strTmp As String
    Dim lngLoop As Long
    
    If vsf.Row = vsf.Rows - 1 Then Exit Sub
    '
    strTmp = vsf.RowData(vsf.Row)
    vsf.RowData(vsf.Row) = Val(vsf.RowData(vsf.Row + 1))
    vsf.RowData(vsf.Row + 1) = Val(strTmp)
    
    For lngLoop = 0 To vsf.Cols - 1
        
        If lngLoop <> mCol.��¼ɫ Then
            strTmp = vsf.TextMatrix(vsf.Row, lngLoop)
                
            vsf.TextMatrix(vsf.Row, lngLoop) = vsf.TextMatrix(vsf.Row + 1, lngLoop)
            vsf.TextMatrix(vsf.Row + 1, lngLoop) = strTmp
        End If
    Next
    
    Set vsf.Cell(flexcpPicture, vsf.Row, mCol.��¼ɫ) = ils16.ListImages("K" & Val(vsf.TextMatrix(vsf.Row, mCol.��ɫ))).Picture
    Set vsf.Cell(flexcpPicture, vsf.Row + 1, mCol.��¼ɫ) = ils16.ListImages("K" & Val(vsf.TextMatrix(vsf.Row + 1, mCol.��ɫ))).Picture
    
    vsf.Row = vsf.Row + 1
    vsf.ShowCell vsf.Row, vsf.Col
    vsf.SetFocus
    
    DataChanged = True
    
    Call SetCmdButtonEnable
    
End Sub

Private Sub cmdOK_Click()
    Dim lngKey As Long
    
    If DataChanged Then
        If SaveData(lngKey) = False Then Exit Sub
                
        mblnOK = True
        
        DataChanged = False
    End If
    
    Unload Me
End Sub

Private Sub cmdUp_Click()
    
    Dim strTmp As String
    Dim lngLoop As Long
    
    If vsf.Row = 1 Then Exit Sub
    '
    strTmp = vsf.RowData(vsf.Row)
    vsf.RowData(vsf.Row) = Val(vsf.RowData(vsf.Row - 1))
    vsf.RowData(vsf.Row - 1) = Val(strTmp)
    
    For lngLoop = 0 To vsf.Cols - 1
        
        If lngLoop <> mCol.��¼ɫ Then
            strTmp = vsf.TextMatrix(vsf.Row, lngLoop)
                
            vsf.TextMatrix(vsf.Row, lngLoop) = vsf.TextMatrix(vsf.Row - 1, lngLoop)
            vsf.TextMatrix(vsf.Row - 1, lngLoop) = strTmp
        End If
    Next
    
    Set vsf.Cell(flexcpPicture, vsf.Row, mCol.��¼ɫ) = ils16.ListImages("K" & Val(vsf.TextMatrix(vsf.Row, mCol.��ɫ))).Picture
    Set vsf.Cell(flexcpPicture, vsf.Row - 1, mCol.��¼ɫ) = ils16.ListImages("K" & Val(vsf.TextMatrix(vsf.Row - 1, mCol.��ɫ))).Picture
    
    vsf.Row = vsf.Row - 1
    vsf.ShowCell vsf.Row, vsf.Col
    vsf.SetFocus
    
    DataChanged = True
    
    Call SetCmdButtonEnable
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If DataChanged Then
        Cancel = (MsgBox("����/�޸ĵ����ݱ��뱣������Ч���Ƿ񲻱�����˳���", vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) = vbNo)
    End If
End Sub

Private Sub vsf_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call SetCmdButtonEnable
End Sub

