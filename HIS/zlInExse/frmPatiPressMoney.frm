VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPatiPressMoney 
   Caption         =   "���˴߿����"
   ClientHeight    =   7305
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11775
   Icon            =   "frmPatiPressMoney.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7305
   ScaleWidth      =   11775
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picDown 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   0
      ScaleHeight     =   705
      ScaleWidth      =   11775
      TabIndex        =   9
      Top             =   6600
      Width           =   11775
      Begin VB.CommandButton cmdPrint 
         Caption         =   "��ӡ�ϱ�(&U)"
         Height          =   375
         Left            =   4140
         TabIndex        =   16
         Top             =   165
         Width           =   1455
      End
      Begin VB.CommandButton cmdExcel 
         Caption         =   "�ϱ����&Excel"
         Height          =   375
         Left            =   2640
         TabIndex        =   15
         Top             =   165
         Width           =   1455
      End
      Begin VB.CommandButton cmdSetup 
         Caption         =   "����(&Z)"
         Height          =   380
         Left            =   7530
         TabIndex        =   14
         Top             =   165
         Width           =   1250
      End
      Begin VB.CommandButton cmdAllSel 
         Caption         =   "ȫѡ(&A)"
         Height          =   380
         Left            =   105
         TabIndex        =   12
         ToolTipText     =   "���:CTRL+A"
         Top             =   165
         Width           =   1250
      End
      Begin VB.CommandButton cmdALLCls 
         Caption         =   "ȫ��(&S)"
         Height          =   380
         Left            =   1365
         TabIndex        =   11
         ToolTipText     =   "���:CTRL+C"
         Top             =   165
         Width           =   1250
      End
      Begin VB.Frame fraBottomSplit 
         Height          =   30
         Left            =   -210
         TabIndex        =   10
         Top             =   0
         Width           =   12405
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "��ӡ(&P)"
         Height          =   380
         Left            =   8865
         TabIndex        =   5
         Top             =   165
         Width           =   1250
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "ȡ��(&C)"
         Height          =   380
         Left            =   10125
         TabIndex        =   6
         Top             =   165
         Width           =   1250
      End
   End
   Begin VB.PictureBox picSeach 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1095
      ScaleWidth      =   11775
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   0
      Width           =   11775
      Begin VB.Frame fraSearch 
         Caption         =   "����:һ����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   840
         Left            =   90
         TabIndex        =   8
         Top             =   120
         Width           =   11235
         Begin VB.CommandButton cmdˢ�� 
            Caption         =   "ˢ��(&R)"
            Height          =   375
            Left            =   8310
            TabIndex        =   3
            Top             =   285
            Width           =   1125
         End
         Begin MSComCtl2.DTPicker dtpEnd 
            Height          =   300
            Left            =   1035
            TabIndex        =   1
            Top             =   330
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy��MM��dd��"
            Format          =   183959555
            CurrentDate     =   36576
         End
         Begin VB.Label lbl��ֹ���� 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��ֹ����"
            Height          =   180
            Left            =   225
            TabIndex        =   0
            Top             =   390
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "֪ͨ������ӡ������ָ����ֹ���������ڼ��ڵķ���Ƿ�������"
            ForeColor       =   &H00800000&
            Height          =   180
            Left            =   2775
            TabIndex        =   2
            Top             =   390
            Width           =   5040
         End
      End
      Begin VB.Image img16 
         Height          =   240
         Left            =   0
         Picture         =   "frmPatiPressMoney.frx":06EA
         Top             =   0
         Visible         =   0   'False
         Width           =   240
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsPressMoney 
      Height          =   5430
      Left            =   75
      TabIndex        =   4
      Top             =   1080
      Width           =   11565
      _cx             =   20399
      _cy             =   9578
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   10.5
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
      GridColor       =   -2147483632
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   9
      GridLineWidth   =   1
      Rows            =   10
      Cols            =   1
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmPatiPressMoney.frx":0C74
      ScrollTrack     =   -1  'True
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
      ExplorerBar     =   3
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   2
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
      Begin VB.PictureBox picImg 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   75
         ScaleHeight     =   225
         ScaleWidth      =   210
         TabIndex        =   13
         Top             =   60
         Width           =   210
         Begin VB.Image imgCol 
            Height          =   195
            Left            =   0
            Picture         =   "frmPatiPressMoney.frx":0CA1
            ToolTipText     =   "ѡ����Ҫ��ʾ����(ALT+C)"
            Top             =   0
            Width           =   195
         End
      End
   End
End
Attribute VB_Name = "frmPatiPressMoney"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngModule As Long, mstrPrivs As String
Private mlng���� As Long, mblnOk As Boolean, mlng����ID As Long
Private mlng��ҳID As Long
Private mblnFirst As Boolean, mbytPrintModule As Byte
Private mlngPrintRow As Long    '��ǰ���ڴ�ӡ����
Private mstrPrintDate As Date
Private WithEvents mobjReport As clsReport
Attribute mobjReport.VB_VarHelpID = -1

Public Function zlPatiPressMoney(ByVal frmMain As Object, ByVal lngMoudle As Long, _
    ByVal strPrivs As String, ByVal lng���� As Long, ByVal str�������� As String, _
    Optional lng����ID As Long = 0, Optional bytPrintModule As Byte = 2) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���벡�˴߿�������
    '���:frmMain-���õĴ���
    '       bytPrintModule-2.��ӡ;1-Ԥ��
    '����:
    '����:�����ӡ�ɹ�1�����ϵĲ���,����true,���򷵻�False
    '����:���˺�
    '����:2010-12-16 10:28:25
    '����:35386
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    On Error GoTo errHandle
    mbytPrintModule = bytPrintModule
    mlngModule = lngMoudle: mstrPrivs = strPrivs: mlng���� = lng����: mblnOk = False: mlng����ID = lng����ID
    If lng���� = 0 And lng����ID = 0 Then
        MsgBox "ע��:" & vbCrLf & "    ��֧�ֶ����в������д�ӡ!", vbInformation + vbDefaultButton1 + vbOKOnly
        Exit Function
    End If
    If lng����ID <> 0 Then
        '76451,Ƚ����,2014-8-19
        gstrSQL = "Select ����,�Ա�,����,��ҳID From ������Ϣ where ����id=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ������Ϣ", lng����ID)
        If rsTemp.EOF Then '
            MsgBox "ע��:" & vbCrLf & " δ�ҵ���صĲ���,���ܼ���!", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
        mlng��ҳID = Nvl(rsTemp!��ҳID)   '42626
        fraSearch.Caption = "����:" & rsTemp!���� & String(4, " ") & "�Ա�:" & Nvl(rsTemp!�Ա�) & String(4, " ") & "����:" & Nvl(rsTemp!����)
        If lng���� <> 0 Then
            gstrSQL = "Select  Max(��ҳID) as ��ҳID From ������ҳ where ����id=[1] And ��ǰ����ID=[2] "
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ������Ϣ", lng����ID, lng����)
            If Val(Nvl(rsTemp!��ҳID)) <> 0 Then mlng��ҳID = Val(Nvl(rsTemp!��ҳID))
        End If
    Else
        fraSearch.Caption = "��" & str�������� & "������Ժ����"
    End If
    
    mblnFirst = True
    Me.Show 1, frmMain
    zlPatiPressMoney = mblnOk
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub cmdALLCls_Click()
    Dim i As Long
    With vsPressMoney
        For i = 1 To .Rows - 1
            .TextMatrix(i, .ColIndex("ѡ��")) = 0
        Next
    End With
End Sub
Private Sub cmdAllSel_Click()
    Dim i As Long
    With vsPressMoney
        For i = 1 To .Rows - 1
            .TextMatrix(i, .ColIndex("ѡ��")) = 1
            If Val(.TextMatrix(i, .ColIndex("�߿���"))) = 0 Then
                .TextMatrix(i, .ColIndex("�߿���")) = .Cell(flexcpData, i, .ColIndex("�߿���"))
            End If
        Next
    End With
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdExcel_Click()
    Call PrintGrid(3)
End Sub

Private Sub cmdOK_Click()
    
    If zlPrintPatiPressMoney = False Then Exit Sub
    mblnOk = True
End Sub

Private Sub cmdPrint_Click()
    Call PrintGrid(1)
End Sub

Private Sub cmdˢ��_Click()
    Call FillData
End Sub

Private Sub dtpEnd_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    If mlng����ID <> 0 Then cmdOK.SetFocus: Exit Sub
    vsPressMoney.SetFocus
    With vsPressMoney
         .Col = .ColIndex("ѡ��")
        .Row = 1
    End With
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

        Select Case KeyCode
        Case vbKeyA
            If Shift = vbCtrlMask Then cmdAllSel_Click
        Case vbKeyC
            If Shift = vbCtrlMask Then cmdALLCls_Click
        End Select
        If Not Me.ActiveControl Is vsPressMoney Then
            If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
        End If
End Sub

Private Sub Form_Load()
    RestoreWinState Me, Me.Name
    dtpEnd.MaxDate = Format(zlDatabase.Currentdate, "yyyy-mm-dd")
    dtpEnd.Value = DateAdd("d", -1, dtpEnd.MaxDate)
    Set mobjReport = New clsReport
    Call FillData
End Sub

Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    With vsPressMoney
        .Left = Me.ScaleLeft + 50
        .Width = Me.ScaleWidth - 100
        .Top = picSeach.Top + picSeach.Height + 20
        .Height = Me.ScaleHeight - .Top - picDown.Height - 50
    End With
End Sub
 
Private Sub mobjReport_AfterPrint(ByVal ReportNum As String)
    Dim i As Long
    i = mlngPrintRow
    With vsPressMoney
        Screen.MousePointer = 11
        If i < 1 Or i > .Rows - 1 Then Exit Sub
        '���²��˵Ľɿ�����
         gstrSQL = "Zl_������ҳ�ӱ�_��ҳ����("
        '    ����id_In ������ҳ�ӱ�.����id%Type,
        gstrSQL = gstrSQL & "" & Val(.TextMatrix(i, .ColIndex("����ID"))) & ","
        '    ��ҳid_In ������ҳ�ӱ�.��ҳid%Type,
        gstrSQL = gstrSQL & "" & Val(.TextMatrix(i, .ColIndex("��ҳID"))) & ","
        '    ��Ϣ��_In ������ҳ�ӱ�.��Ϣ��%Type,
        gstrSQL = gstrSQL & "'�ϴδ߿���',"
        '    ��Ϣֵ_In ������ҳ�ӱ�.��Ϣֵ%Type
        gstrSQL = gstrSQL & "" & Val(.TextMatrix(i, .ColIndex("�߿���"))) & ")"
        zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
        'ZL_���˴߿��¼_INSERT(
        gstrSQL = "ZL_���˴߿��¼_INSERT("
        '    ����ID_IN IN ���˴߿��¼.����ID%TYPE,
        gstrSQL = gstrSQL & "" & Val(.TextMatrix(i, .ColIndex("����ID"))) & ","
        '    ��ҳID_IN IN ���˴߿��¼.��ҳID%TYPE,
        gstrSQL = gstrSQL & "" & Val(.TextMatrix(i, .ColIndex("��ҳID"))) & ","
        '    Ԥ�����_IN IN ���˴߿��¼.Ԥ�����%TYPE,
        gstrSQL = gstrSQL & "" & Val(.TextMatrix(i, .ColIndex("Ԥ�����"))) & ","
        '    δ�����_IN IN ���˴߿��¼.δ�����%TYPE,
        gstrSQL = gstrSQL & "" & Val(.TextMatrix(i, .ColIndex("δ�����"))) & ","
        '    �Էѽ��_IN IN ���˴߿��¼.�Էѽ��,
        gstrSQL = gstrSQL & "" & Val(.TextMatrix(i, .ColIndex("δ�����"))) - Val(.TextMatrix(i, .ColIndex("ҽ��Ԥ��"))) & ","
        '    ҽ��Ԥ��_IN IN ���˴߿��¼.ҽ��Ԥ��%TYPE,
        gstrSQL = gstrSQL & "" & Val(.TextMatrix(i, .ColIndex("ҽ��Ԥ��"))) & ","
        '    ��ǰ���_IN IN ���˴߿��¼.��ǰ���%TYPE,
        gstrSQL = gstrSQL & "" & Val(.TextMatrix(i, .ColIndex("Ԥ�����"))) + Val(.TextMatrix(i, .ColIndex("ҽ��Ԥ��"))) - Val(.TextMatrix(i, .ColIndex("δ�����"))) & ","
        '    �߿�����_IN IN ���˴߿��¼.�߿�����%TYPE,
        gstrSQL = gstrSQL & "" & Val(.TextMatrix(i, .ColIndex("�߿�����"))) & ","
        '    �߿��׼_IN IN ���˴߿��¼.�߿��׼%TYPE,
        gstrSQL = gstrSQL & "" & Val(.TextMatrix(i, .ColIndex("�߿��׼"))) & ","
        '    �߿���_IN IN ���˴߿��¼.�߿���%TYPE,
        gstrSQL = gstrSQL & "" & Val(.TextMatrix(i, .ColIndex("�߿���"))) & ","
        '    ��ӡ����_IN IN ���˴߿��¼.��ӡ����%TYPE,
        gstrSQL = gstrSQL & "to_date('" & mstrPrintDate & "','yyyy-mm-dd hh24:mi:ss'),"
        '    ��ӡ��_IN IN ���˴߿��¼.��ӡ��%TYPE
        gstrSQL = gstrSQL & "'" & UserInfo.���� & "')"
        zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
        .Cell(flexcpData, i, .ColIndex("����")) = "1"
        .Cell(flexcpPicture, i, .ColIndex("����")) = img16.Picture
        .Cell(flexcpPictureAlignment, i, .ColIndex("����")) = 1
    End With
End Sub
 
Private Sub picSeach_Resize()
    Err = 0: On Error Resume Next
    With picSeach
        fraSearch.Left = .ScaleLeft + 50
        fraSearch.Top = .ScaleTop + 100
       ' fraSearch.Height = .ScaleHeight - 100
        fraSearch.Width = .ScaleWidth - 100
        cmdˢ��.Left = .ScaleWidth - fraSearch.Left - cmdˢ��.Width * 2
    End With
End Sub

Private Sub picDown_Resize()
    Err = 0: On Error Resume Next
    With picDown
        fraBottomSplit.Left = .ScaleLeft
        fraBottomSplit.Top = .ScaleTop
        fraBottomSplit.Width = .ScaleWidth
        cmdCancel.Left = .ScaleWidth - cmdCancel.Width - cmdCancel.Width / 2
        cmdOK.Left = cmdCancel.Left - cmdOK.Width - 50
        cmdSetup.Left = cmdOK.Left - cmdSetup.Width - 50
    End With
End Sub
Private Function FillData(Optional blnDefault As Boolean = True) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������
    '���:
    '����:
    '����:���سɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2010-12-21 15:06:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String, strWhere As String, rsTemp As ADODB.Recordset, i As Long
    
    On Error GoTo errHandle
    strWhere = ""
    '��ǰ��Ժ�Ĳ���
    '42626
    If mlng����ID = 0 Then strWhere = "   And A.����ID=J1.����ID And J1.����ID=[1] And Nvl(B.״̬,0)<>3 And A.��ҳID=B.��ҳID  "
     If mlng���� > 0 And mlng����ID <> 0 Then strWhere = strWhere & " And B.��ǰ����ID=[1]"
     If mlng����ID > 0 Then strWhere = strWhere & " And B.����ID=[2] And B.��ҳID=[3]"
     'ʣ���:Format((!Ԥ����� - !������� + !Ԥ�����+ Nvl(A.������, 0) ):Max(Nvl(A.������, 0)):37785
    '��ʽ:Ԥ�����+������+��ҽ�����˱����ܶ-δ��������-�߿�����<0 ��Ϊ����Ƿ��Ĳ���
   strSql = "" & _
    "   Select A.����ID, B.��ҳID, B.״̬, B.��������,B.��Ժ����id As ��ǰ����id, B.����, B.��ǰ����ID,  " & _
    "            '1' as ѡ��,A.����, B.סԺ��, B.��Ժ���� As ����, B.�ѱ�, A.�Ա�, A.����, C.���� As ��ǰ����, A.���￨��, E.����,  " & _
    "           to_char(B.��Ժ����,'yyyy-mm-dd hh24:mi:ss') as ��Ժ����,  to_char(B.��Ժ����,'yyyy-mm-dd hh24:mi:ss') as ��Ժ����,  to_char(A.�Ǽ�ʱ��,'yyyy-mm-dd hh24:mi:ss') as �Ǽ�ʱ��, " & _
    "           ltrim(to_char(nvl(Max(M.�߿�����),0),'9999999990.00')) As �߿�����, ltrim(to_char(nvl(Max(M.�߿��׼),0),'9999999990.00')) As �߿��׼," & vbNewLine & _
    "           ltrim(to_char(Max(Nvl(X.Ԥ�����, 0)),'9999999990.00')) As Ԥ�����,ltrim(to_char(Sum(Nvl(X1.���, 0)),'9999999990.00')) As ҽ��Ԥ��, " & vbNewLine & _
    "           ltrim(to_char(Max(Nvl(A.������, 0)),'9999999990.00')) As ������,ltrim(to_char(Max(Nvl(X.�������, 0)),'9999999990.00')) As δ�����," & vbNewLine & _
    "           ltrim(to_char(Max(Nvl(X.Ԥ�����, 0))-Max(Nvl(X.�������, 0))+Sum(Nvl(X1.���, 0)) ,'9999999990.00')) As ʣ���," & vbNewLine & _
    "           ltrim(to_char(case when to_number(Max(nvl(D1.��Ϣֵ,'0')))> 0 then  to_number(Max(nvl(D1.��Ϣֵ,'0')))  " & _
    "                               When  Max(nvl(X.Ԥ�����,0))+Max(Nvl(A.������, 0))+Sum(nvl(X1.���,0))-Max(nvl(x.�������,0))<0 then round(abs(Max(nvl(X.Ԥ�����,0))+Max(Nvl(A.������, 0))+Sum(nvl(X1.���,0))-Max(nvl(x.�������,0)))/100,0)*100+Max(nvl(M.�߿��׼,0)) " & _
    "                               Else   Max(M.�߿��׼) end,'9999999990.00')) As �߿���," & vbNewLine & _
    "           Nvl(E.ҽ����, D.��Ϣֵ) ҽ����, A.��ͥ�绰, B.ҽ�Ƹ��ʽ, B.�����, B.��������, H.���� ��ǰ����" & vbNewLine & _
    "   From ������Ϣ A, ������ҳ B, ������ҳ�ӱ� D, ������ҳ�ӱ� D1, ҽ�����˵��� E, ҽ�����˹����� F, ������� X,����ģ����� X1, " & vbNewLine & _
    "         ���ʱ����� M,���ű� C, ���ű� H" & IIf(mlng����ID = 0, ",��Ժ���� J1", "") & vbNewLine & _
    "   Where A.����ID = B.����ID And B.��Ժ����ID = C.ID And Nvl(B.��ҳID, 0) <> 0 " & vbNewLine & _
    "          And B.����ID = D.����ID(+) And B.��ҳID = D.��ҳID(+) And  D.��Ϣ��(+) = 'ҽ����'  " & vbNewLine & _
    "          And B.����ID = D1.����ID(+) And B.��ҳID = D1.��ҳID(+) And  D1.��Ϣ��(+) = '�ϴδ߿���'  " & vbNewLine & _
    "          And A.����ID = X.����ID(+) And X.����(+) = 1 And X.����(+)=2 And B.����ID=X1.����ID(+) and B.��ҳid=X1.��ҳID(+)  " & vbNewLine & _
    "          And A.����ID = F.����ID(+) And F.��־(+) = 1 And F.ҽ���� = E.ҽ����(+) And F.���� = E.����(+) And F.���� = E.����(+)  " & vbNewLine & _
    "          And B.��ǰ����ID = H.ID And (H.վ��='" & gstrNodeNo & "' Or H.վ�� is Null)" & vbNewLine & strWhere & vbNewLine & _
    "          And B.��ǰ����ID=M.����ID(+) And zl_PatiWarnScheme(b.����id,b.��ҳID) =M.���ò���(+) " & vbNewLine & _
    "   Group by A.����ID, B.��ҳID, A.�Ǽ�ʱ��, B.״̬, B.��������, B.����ת��, B.��Ժ����id, A.���￨��, B.����, E.����, B.��ǰ����ID, A.����, B.סԺ��, B.��Ժ����,B.�ѱ�, A.�Ա�, A.����, B.��Ժ����, B.��Ժ����, C.����, Decode(Nvl(X.�������, 0), 0, '��', ''), " & vbNewLine & _
    "          Nvl(E.ҽ����, D.��Ϣֵ), A.��ͥ�绰, B.ҽ�Ƹ��ʽ, B.�����, B.��������, H.���� " & _
         IIf(mlng����ID <> 0, "", "   having (Max(nvl(X.Ԥ�����,0))+Max(Nvl(A.������, 0))+Sum(nvl(X1.���,0))-Max(nvl(x.�������,0))-Max(nvl(M.�߿�����,0)))<0 " & vbNewLine) & _
         IIf(mlng���� = 0, " Order by סԺ�� Desc", " Order by LPAD(����,10,' ')")
 
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng����, mlng����ID, mlng��ҳID)
    With vsPressMoney
        .Clear 0: .Cols = 1
        .FixedCols = 1
       Set .DataSource = rsTemp
        If .Rows <= 1 Then .Rows = 2
        For i = 1 To .Cols - 1
            .ColKey(i) = UCase(.TextMatrix(0, i))
            .ColData(i) = "0||1"
            .ColAlignment(i) = flexAlignLeftCenter
            .FixedAlignment(i) = flexAlignCenterCenter
            If .ColKey(i) Like "*ID" Or .ColKey(i) = "״̬" Or .ColKey(i) = "����" Or .ColKey(i) = "��������" Or .ColKey(i) = "����" Then
                .ColHidden(i) = True
                ' ColData(i):����������(1-�̶�,-1-����ѡ,0-��ѡ)||������(0-��������,1-��ֹ����,2-��������,�����س���������)
                .ColData(i) = "-1||1"
            End If
            If .ColKey(i) Like "*��׼*" Or .ColKey(i) Like "*����*" Or .ColKey(i) Like "*��*" Then
                 .ColAlignment(i) = flexAlignRightCenter
            End If
            '���������п�
            Select Case .ColKey(i)
            Case "סԺ��", "����", "�Ա�", "����", "��ǰ����", "��������"
            Case "Ԥ�����", "δ�����", "�߿�����", "�߿��׼", "ʣ���", "ҽ��Ԥ��"
                 .ColAlignment(i) = flexAlignRightCenter
            Case "����", "�߿���", "ѡ��"
                   .ColData(i) = "1||0"
                   If .ColKey(i) = "ѡ��" Then .ColDataType(i) = flexDTBoolean
                   If .ColKey(i) = "ѡ��" Then .ColAlignment(i) = flexAlignCenterCenter
            Case Else
                .ColHidden(i) = True
            End Select
        Next
        '������ɫ
        .Redraw = flexRDNone
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, .ColIndex("����"))) <> 0 Then
                .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = vbRed
            End If
            If Val(.TextMatrix(i, .ColIndex("Ԥ�����"))) = 0 Then .TextMatrix(i, .ColIndex("Ԥ�����")) = ""
            If Val(.TextMatrix(i, .ColIndex("δ�����"))) = 0 Then .TextMatrix(i, .ColIndex("δ�����")) = ""
            If Val(.TextMatrix(i, .ColIndex("�߿�����"))) = 0 Then .TextMatrix(i, .ColIndex("�߿�����")) = ""
            If Val(.TextMatrix(i, .ColIndex("�߿��׼"))) = 0 Then .TextMatrix(i, .ColIndex("�߿��׼")) = ""
            If Val(.TextMatrix(i, .ColIndex("ʣ���"))) = 0 Then .TextMatrix(i, .ColIndex("ʣ���")) = ""
            If Val(.TextMatrix(i, .ColIndex("ҽ��Ԥ��"))) = 0 Then .TextMatrix(i, .ColIndex("ҽ��Ԥ��")) = ""
            
        Next
        '�Զ��п�
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSize 0, .Cols - 1
        '�ָ�������
        zl_vsGrid_Para_Restore mlngModule, vsPressMoney, Me.Caption, "�߿��б�", False
        If .ColIndex("��־") >= 0 Then .ColWidth(.ColIndex("��־")) = 300
        .Cell(flexcpBackColor, 1, .ColIndex("ѡ��"), .Rows - 1, .ColIndex("ѡ��")) = &HE7CFBA
        .Cell(flexcpBackColor, 1, .ColIndex("�߿���"), .Rows - 1, .ColIndex("�߿���")) = &HE7CFBA
        .Redraw = flexRDBuffered
    End With
    
    FillData = True
    Exit Function
errHandle:
    vsPressMoney.Redraw = flexRDBuffered
    If ErrCenter() = 1 Then
        Resume
    End If
    
End Function
Private Sub Form_Unload(Cancel As Integer)
    Set mobjReport = Nothing
    SaveWinState Me, Me.Name
    zl_vsGrid_Para_Save mlngModule, vsPressMoney, Me.Caption, "�߿��б�", False, , InStr(1, mstrPrivs, ";��������;") > 0
End Sub
 
Private Sub vsPressMoney_AfterEdit(ByVal Row As Long, ByVal Col As Long)
        With vsPressMoney
            Select Case Col
            Case .ColIndex("�߿���")
                .TextMatrix(Row, Col) = Format(Val(.TextMatrix(Row, .Col)), "0.00")
                If Val(.TextMatrix(Row, Col)) <> 0 And GetVsGridBoolColVal(vsPressMoney, Row, .ColIndex("ѡ��")) = False Then
                    vsPressMoney.TextMatrix(Row, .ColIndex("ѡ��")) = 1
                ElseIf Val(.TextMatrix(Row, Col)) = 0 Then
                    vsPressMoney.TextMatrix(Row, .ColIndex("ѡ��")) = 0
                End If
            Case .ColIndex("ѡ��")
                If GetVsGridBoolColVal(vsPressMoney, Row, Col) Then
                    If Val(.TextMatrix(Row, .ColIndex("�߿���"))) = 0 Then
                        .TextMatrix(Row, .ColIndex("�߿���")) = Format(Val(.Cell(flexcpData, Row, .ColIndex("�߿���"))), "0.00")
                    End If
                End If
            Case Else
            End Select
        End With
End Sub
Private Sub vsPressMoney_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModule, vsPressMoney, Me.Caption, "�߿��б�", False, , InStr(1, mstrPrivs, ";��������;") > 0
End Sub

Private Sub vsPressMoney_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModule, vsPressMoney, Me.Caption, "�߿��б�", False, , InStr(1, mstrPrivs, ";��������;") > 0
End Sub

Private Sub imgCol_Click()
    Dim lngLeft As Long, lngTop As Long
    Dim vRect  As RECT
    vRect = zlcontrol.GetControlRect(picImg.hWnd)
    lngLeft = vRect.Left
    lngTop = vRect.Top + picImg.Height
    Call frmVsColSel.ShowColSet(Me, Me.Caption, vsPressMoney, lngLeft, lngTop, imgCol.Height)
    zl_vsGrid_Para_Save mlngModule, vsPressMoney, Me.Caption, "�߿��б�", False, , InStr(1, mstrPrivs, ";��������;") > 0
End Sub
Private Sub picImg_Click()
    Call imgCol_Click
End Sub
Private Function zlPrintPatiPressMoney() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ӡ���˽ɿ�֪ͨ��
    '����:���˺�
    '����:2010-12-16 15:15:01
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim i As Long, blnData As Boolean
    Dim str��ֹ���� As String
    Dim lngCount As Long
    '�ȼ��ɼ�������
    With vsPressMoney
        blnData = False
        mstrPrintDate = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
        For i = 1 To .Rows - 1
            If GetVsGridBoolColVal(vsPressMoney, i, .ColIndex("ѡ��")) Then
                If Val(.TextMatrix(i, .ColIndex("�߿���"))) <= 0 Then
                    MsgBox "ע��:" & "    �ڵ�" & i & "���еĴ߿�����������,����!", vbOKOnly + vbInformation, gstrSysName
                    .Row = i: .Col = .ColIndex("�߿���")
                    If .RowIsVisible(.Row) = False Or .ColIsVisible(.Col) = False Then
                        Call .ShowCell(.Row, .Col)
                    End If
                    Exit Function
                End If
                lngCount = lngCount + 1
                blnData = True
            End If
        Next
    End With
    
    If blnData = False Then
        MsgBox "ע��:" & "    δѡ��ָ���Ĵ�ӡ����,����!", vbOKOnly + vbInformation, gstrSysName
        vsPressMoney.SetFocus
        Exit Function
    End If
    If MsgBox("���Ƿ���Ҫ��ӡ��" & lngCount & "�����˵Ĵ߿��?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    str��ֹ���� = Format(dtpEnd.Value, "yyyy-mm-dd")
    
    With vsPressMoney
        Screen.MousePointer = 11
        For i = 1 To .Rows - 1
            If GetVsGridBoolColVal(vsPressMoney, i, .ColIndex("ѡ��")) And Val(.TextMatrix(i, .ColIndex("����ID"))) > 0 Then
                mlngPrintRow = i
                Call mobjReport.ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_INSIDE_1139_3", Me, "����ID=" & Val(.TextMatrix(i, .ColIndex("����ID"))), _
                    "����=" & str��ֹ����, "�߿���=" & Val(.TextMatrix(i, .ColIndex("�߿���"))), mbytPrintModule)
            Else
               .Cell(flexcpData, i, .ColIndex("����")) = "0"
            End If
        Next
    End With
    Screen.MousePointer = 0
    MsgBox "���в��˴�ӡ��ɣ�", vbInformation, gstrSysName
    zlPrintPatiPressMoney = True
End Function
Private Sub vsPressMoney_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsPressMoney
        Select Case Col
        Case .ColIndex("�߿���"), .ColIndex("ѡ��")
            Exit Sub
        Case Else
            Cancel = True
        End Select
    End With
End Sub

Private Sub vsPressMoney_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsPressMoney
        Select Case Col
        Case .ColIndex("��־")
            Cancel = True
        Case Else
        End Select
    End With
End Sub

Private Sub vsPressMoney_EnterCell()
    '��δ����
    With vsPressMoney
    End With
End Sub

Private Sub vsPressMoney_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngCol As Long, blnCancel As Boolean, lngRow As Long
    If KeyCode <> vbKeyReturn Then Exit Sub
    With vsPressMoney
        If Val(.TextMatrix(.Row, .ColIndex("����ID"))) = 0 Or (.Col >= .ColIndex("�߿���") And .Row = .Rows - 1) Then
            zlCommFun.PressKey vbKeyTab
            Exit Sub
        End If
    With vsPressMoney
        Select Case .Col
        Case .ColIndex("�߿���")
                If .Row < .Rows - 1 Then
                    .Col = .Col: .Row = .Row + 1
                End If
        Case .ColIndex("ѡ��")
                If .ColIndex("ѡ��") > .ColIndex("�߿���") Then
                   .Col = .ColIndex("�߿���"): .Row = .Row + 1
                Else
                    .Col = .ColIndex("�߿���")
                End If
        Case Else
        End Select
    End With
        
    End With
End Sub

Private Sub vsPressMoney_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    '�༭����
    Dim intCol As Integer, strKey As String, lngRow As Long
    If KeyCode <> vbKeyReturn Then Exit Sub
    With vsPressMoney
        Select Case Col
        Case .ColIndex("�߿���")
                If Row < .Rows - 1 Then
                    .Col = Col: .Row = .Row + 1
                End If
        Case .ColIndex("ѡ��")
                If .ColIndex("ѡ��") > .ColIndex("�߿���") Then
                   .Col = .ColIndex("�߿���"): .Row = .Row + 1
                Else
                    .Col = .ColIndex("�߿���")
                End If
        Case Else
        End Select
    End With
End Sub

Private Sub vsPressMoney_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
    End If
End Sub

Private Sub vsPressMoney_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    With vsPressMoney
        Select Case .Col
            Case .ColIndex("�߿���")
                If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
                    If KeyAscii = vbKeyBack Then Exit Sub
                    If KeyAscii = vbKeyReturn Then Exit Sub
                    If KeyAscii = Asc(".") Then
                        If InStr(1, .EditText, ".") = 0 Then
                            Exit Sub
                        End If
                    End If
                    KeyAscii = 0
                End If
            Case Else
            
        End Select
    End With
End Sub

Private Sub vsPressMoney_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim strKey As String
    '������֤
    With vsPressMoney
        strKey = Trim(.EditText): strKey = Replace(strKey, Chr(vbKeyReturn), ""): strKey = Replace(strKey, Chr(10), "")
        Select Case Col
            Case .ColIndex("�߿���")
                If Val(strKey) > 999999999 Then
                    MsgBox "ע��:" & vbCrLf & "    �߿���ֻ����0-999999999��Χ��!"
                    Cancel = True
                End If
                strKey = Format(Val(strKey), "0.00")
                .EditText = strKey
                .TextMatrix(Row, .Col) = strKey
        End Select
    End With
End Sub
  
Private Sub cmdSetup_Click()
    ReportPrintSet gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_INSIDE_1139_3", Me
End Sub

Private Sub PrintGrid(bytMode As Byte)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ӡ�������
    '���:bytMode:1-��ӡ,2-Ԥ��,3-�����Excel
    '����:���˺�
    '����:2011-05-13 10:10:23
    '����:37934
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPrintObject  As New zlPrint1Grd
    Dim objRow As New zlTabAppRow
    Dim lngCol As Long, lngRow As Long, i As Long
    Dim cllCol As New Collection

    Err = 0: On Error GoTo errHandle
    '��¼��״̬
    With vsPressMoney
        .Redraw = flexRDNone
        lngRow = .Row: lngCol = .Col
        For i = 0 To .Cols - 1
             cllCol.Add Array(CStr(.ColData(i)), .ColWidth(i), IIf(.ColHidden(i), 1, 0)), "K" & i
             If i = .ColIndex("��־") Then .ColWidth(i) = 0
             If .ColHidden(i) Then .ColWidth(i) = 0
        Next
    End With
        
    '��ͷ
    objPrintObject.Title.Text = "���˴߿��"
    objPrintObject.Title.Font.Name = "����_GB2312"
    objPrintObject.Title.Font.Size = 18
    objPrintObject.Title.Font.Bold = True
    '����
    objRow.Add fraSearch.Caption
    objRow.Add "��ֹ���ڣ�" & Format(dtpEnd.Value, "yyyy��mm��DD��")
    objPrintObject.UnderAppRows.Add objRow
    
    Set objRow = New zlTabAppRow
    objRow.Add "��ӡ�ˣ�" & UserInfo.����
    objRow.Add "��ӡ���ڣ�" & Format(zlDatabase.Currentdate(), "yyyy��MM��dd�� HH:MM:SS")
    objPrintObject.BelowAppRows.Add objRow
    '����
    Set objPrintObject.Body = vsPressMoney
    
    '���
    If bytMode = 1 Then
        bytMode = zlPrintAsk(objPrintObject)
        Me.Refresh
        If bytMode <> 0 Then zlPrintOrView1Grd objPrintObject, bytMode
    Else
        zlPrintOrView1Grd objPrintObject, bytMode
    End If
    '�ָ�ԭʼ״̬
     With vsPressMoney
         .Row = lngRow: .Col = lngCol
        For i = 1 To cllCol.Count
             .ColData(i - 1) = cllCol(i)(0)
             .ColWidth(i - 1) = cllCol(i)(1)
             .ColHidden(i - 1) = IIf(cllCol(i)(2) = 1, True, False)
        Next
        .Redraw = flexRDBuffered
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    '�ָ�ԭʼ״̬
     With vsPressMoney
         .Row = lngRow: .Col = lngCol
        For i = 1 To cllCol.Count
             .ColData(i - 1) = cllCol(i)(0)
             .ColWidth(i - 1) = cllCol(i)(1)
             .ColHidden(i - 1) = IIf(cllCol(i)(2) = 1, True, False)
        Next
        .Redraw = flexRDBuffered
    End With
End Sub
