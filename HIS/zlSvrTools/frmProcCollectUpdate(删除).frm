VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProcCollectUpdate 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�Ѽ�����/���̲�����"
   ClientHeight    =   6720
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10455
   Icon            =   "frmProcCollectUpdate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6720
   ScaleWidth      =   10455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin MSComctlLib.ImageList imgTrueFalse 
      Left            =   9360
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProcCollectUpdate.frx":6852
            Key             =   "T"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProcCollectUpdate.frx":6DEC
            Key             =   "F"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picBottom 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   0
      ScaleHeight     =   705
      ScaleWidth      =   10455
      TabIndex        =   8
      Top             =   6015
      Width           =   10455
      Begin VB.CommandButton cmdDel 
         Caption         =   "ɾ��(&D)"
         Height          =   350
         Left            =   5880
         TabIndex        =   13
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdExit 
         Cancel          =   -1  'True
         Caption         =   "�˳�(&X)"
         Default         =   -1  'True
         Height          =   350
         Left            =   9000
         TabIndex        =   10
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdStart 
         Caption         =   "��ʼ(&S)"
         Height          =   350
         Left            =   7800
         TabIndex        =   9
         Top             =   120
         Width           =   1100
      End
      Begin MSComctlLib.ProgressBar pbrCollect 
         Height          =   105
         Left            =   120
         TabIndex        =   11
         Top             =   510
         Visible         =   0   'False
         Width           =   9960
         _ExtentX        =   17568
         _ExtentY        =   185
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "��ʼ�ռ�"
         Height          =   180
         Left            =   135
         TabIndex        =   12
         Top             =   285
         Visible         =   0   'False
         Width           =   720
      End
   End
   Begin VB.PictureBox picTop 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   0
      ScaleHeight     =   1455
      ScaleWidth      =   10455
      TabIndex        =   1
      Top             =   0
      Width           =   10455
      Begin VB.CommandButton cmdConnet 
         Caption         =   "��������(&L)"
         Enabled         =   0   'False
         Height          =   350
         Left            =   3090
         TabIndex        =   5
         Top             =   1125
         Width           =   1290
      End
      Begin VB.OptionButton optDB 
         Caption         =   "�������ݿ�"
         Height          =   255
         Index           =   1
         Left            =   1635
         TabIndex        =   4
         Top             =   1170
         Width           =   1380
      End
      Begin VB.OptionButton optDB 
         Caption         =   "��ǰ���ݿ�"
         Height          =   255
         Index           =   0
         Left            =   270
         TabIndex        =   3
         Top             =   1170
         Value           =   -1  'True
         Width           =   1305
      End
      Begin VB.PictureBox picFunCap 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   720
         Left            =   120
         Picture         =   "frmProcCollectUpdate.frx":7386
         ScaleHeight     =   720
         ScaleWidth      =   720
         TabIndex        =   2
         Top             =   120
         Width           =   720
      End
      Begin VB.Label lblFunNote 
         Caption         =   $"frmProcCollectUpdate.frx":8250
         Height          =   450
         Left            =   1020
         TabIndex        =   7
         Top             =   630
         Width           =   9180
      End
      Begin VB.Label lblFunCap 
         AutoSize        =   -1  'True
         Caption         =   "�Ѽ��Ǽǹ���/����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   990
         TabIndex        =   6
         Top             =   150
         Width           =   2820
      End
   End
   Begin MSComDlg.CommonDialog dlg 
      Left            =   8760
      Top             =   3600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfMain 
      Height          =   4155
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   9975
      _cx             =   17595
      _cy             =   7329
      Appearance      =   1
      BorderStyle     =   0
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
      BackColorSel    =   16772055
      ForeColorSel    =   0
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483632
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483638
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   330
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmProcCollectUpdate.frx":82B2
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
      ExplorerBar     =   1
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
   End
End
Attribute VB_Name = "frmProcCollectUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'==============================================================
'==ģ�����
'==============================================================
Private mobjMain As Object
Private mblnOk As Boolean
Private mcnOracle As ADODB.Connection
Private mrsUpgradeFiles As ADODB.Recordset
Private mintType As Integer
Private WithEvents mfrmPageConfigure As frmProcConfigure
Attribute mfrmPageConfigure.VB_VarHelpID = -1
Private Enum SysInfoCol
    SC_��� = 0
    SC_�汾�� = 1
    SC_ϵͳ���� = 2
    SC_��װ�ű� = 3
    SC_���ð汾 = 4
End Enum

Private Enum DBType
    CurDB = 0
    OtherDB = 1
End Enum
'==============================================================
'==�����ӿ�
'==============================================================
Public Function ShowMe(ByVal objMain As Object, Optional ByVal intType As Integer) As Boolean
'������intType�� 0-�洢�����ռ�,1-����У��
    Dim strSQL As String, rsData As ADODB.Recordset
        
    On Error GoTo errHand
    mblnOk = False
    Set mobjMain = objMain
    mintType = intType
    '����Աȱ������ݿ�����Զ���洢������հ״洢����
    If mintType <> 0 Then
        strSQL = "Select ID,����,������ From zlprocedure Where ���� In (1,2)"
        Set rsData = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "��ȡ�����б�")
        If rsData.EOF Then
            MsgBox "��ǰ������û�б�׼���̺Ϳհ׹��̣�", vbInformation + vbOKOnly, "�������"
            Exit Function
        End If
    Else
        strSQL = "Select 1 From Zlsystems Where Upper(������) = User"
        Set rsData = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "�ж��Ƿ�ϵͳ������")
        If rsData.RecordCount = 0 Then
            MsgBox "����ϵͳ�����ߵ�¼���д洢�����ռ���", vbInformation + vbOKOnly, "�������"
            Exit Function
        End If
    End If
    Me.Show 1, mobjMain
    ShowMe = mblnOk
    Exit Function
errHand:
    MsgBox err.Description, vbInformation, gstrSysName
End Function

Private Sub cmdDel_Click()
    Call vsfMain_KeyDown(vbKeyDelete, 0)
End Sub

'==============================================================
'==�ؼ��¼�
'==============================================================
Private Sub cmdExit_Click()
    '�����ռ����߼�����
    If Not cmdStart.Enabled Then
        If MsgBox("���ڽ���" & IIf(mintType = 0, "�����ռ�", "���̼��") & ",ȷ��Ҫ�˳���", vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Sub
        End If
    End If
    Unload Me
End Sub

Private Sub cmdConnet_Click()
    If mfrmPageConfigure Is Nothing Then Set mfrmPageConfigure = New frmProcConfigure
    Call mfrmPageConfigure.ShowConfigure(Me)
End Sub

Private Sub cmdStart_Click()
    Dim arrTmp As Variant, lngLoop As Long, lngEnd As Long
    Dim strUpgrade As String, rsFile As ADODB.Recordset
    Dim arrSQL() As Variant
    
    mblnOk = False
    Call ShowState("(1/8)���ڼ���Ҫ����.")
    DoEvents
    If Not ValidData Then GoTo errEnd
    Call ShowState("(2/8)���ڽ�����ʱĿ¼..")
    Call DealWithTmpFolder(True)
    Call ShowState(IIf(mintType = 0, "(3/8)����׼�����ݿ����..", "(3/8)����׼���ϴεı�׼����.."))
    '�����ݿ�������ɵ����ű��ļ�
    If Not LoadBaseProcs(App.Path & "\BaseProcedure") Then GoTo errEnd
    Call ShowState("(4/8)����׼����׼�ű�����..")
    If Not LoadComProcs(App.Path & "\ComProcedure") Then GoTo errEnd
    '�����ݿ��еĹ�����ű����бȶԣ�����html����
    Call ShowState("(5/8)���ڱȽ�..")
    If Not CompareFolder(App.Path & "\BaseProcedure", App.Path & "\ComProcedure", App.Path & "\Reports") Then
        GoTo errEnd
    End If
    Call ShowState("(6/8)������������...")
    Call CreateResultSQL(App.Path & "\BaseProcedure", App.Path & "\ComProcedure", App.Path & "\Reports", arrSQL)
    On Error GoTo errHand
    Call ShowState("(7/8)�����ύ����...")
    If Not gclsBase.ExecuteProcedureBeach(gcnOracle, arrSQL, "������") Then GoTo errHand
    Call ShowState("(8/8)���������ʱ����...")
    Call DealWithTmpFolder
    Call ShowState("", True)
    mblnOk = True
    If mintType = 0 Then
        MsgBox "�洢���̺ͺ������ռ�������ɣ�", vbInformation, gstrSysName
    Else
        MsgBox "�洢���̺ͺ����ļ�������ɣ�", vbInformation, gstrSysName
    End If
    Unload Me
    Exit Sub
    '------------------------------------------------------------------------------------------------------------------
errEnd:
    Call ShowState("", True)
    Exit Sub
errHand:
    Call ShowState("", True)
    MsgBox IIf(mintType = 0, "�ռ�����ʧ�ܣ�", "������ʧ�ܣ�") & vbCrLf & err.Description, vbInformation, gstrSysName
End Sub

Private Sub Form_Load()
    Call InitFace
    Call optDB_Click(CurDB)
    Call Form_Resize
End Sub

Private Sub Form_Resize()
    If mintType = 0 Then
        picTop.Height = cmdConnet.Top + cmdConnet.Height + 60
    Else
        picTop.Height = picFunCap.Top + picFunCap.Height + 30
    End If
    vsfMain.Move 120, picTop.Top + picTop.Height + 30
    vsfMain.Height = picBottom.Top - 30 - vsfMain.Top
    vsfMain.Width = Me.ScaleWidth - 3 * vsfMain.Left
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not (mfrmPageConfigure Is Nothing) Then Unload mfrmPageConfigure
End Sub

Private Sub mfrmPageConfigure_AfterConn(ByVal cnOracle As ADODB.Connection)
    Set mcnOracle = cnOracle
    Call LoadData
End Sub

Private Sub optDB_Click(Index As Integer)
    cmdConnet.Enabled = optDB(OtherDB).value
    Select Case Index
        Case 0
            Set mcnOracle = gcnOracle
            Call LoadData
        Case 1
            vsfMain.Rows = vsfMain.FixedRows
    End Select
End Sub

Private Sub vsfMain_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    With vsfMain
        .Redraw = False
        If .Rows - 1 > 0 Then
            .Cell(flexcpForeColor, .FixedRows, SC_���, .Rows - 1, SC_���) = Color.���ɫ
            .Cell(flexcpFontBold, .FixedRows, SC_���, .Rows - 1, SC_���) = False
        End If
        .Cell(flexcpFontBold, .Row, SC_���, .Row, SC_���) = True
        .Cell(flexcpForeColor, .Row, SC_���, .Row, SC_���) = Color.��ɫ
        .Redraw = True
    End With
End Sub

Private Sub vsfMain_AfterSort(ByVal Col As Long, Order As Integer)
    Call SetSerial
End Sub

Private Sub vsfMain_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <= SC_ϵͳ���� Then
        Cancel = True
    End If
End Sub

Private Sub vsfMain_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    If Col = SC_��װ�ű� Then
        With dlg
            .DialogTitle = "ѡ��Ӧ�ð�װ�����ļ�"
            .Filter = "(Ӧ�ð�װ�����ļ�)|zlSetup.ini"
            .ShowOpen
            If .FileName = "" Then
                Exit Sub
            Else
                vsfMain.TextMatrix(Row, Col) = .FileName
            End If
        End With
    End If
End Sub

Private Sub vsfMain_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        If vsfMain.Row >= vsfMain.FixedRows Then
            If Shift <> vbCtrlMask Then
                vsfMain.RemoveItem vsfMain.Row
                Call SetSerial
            ElseIf vsfMain.Col = SC_��װ�ű� Then
                vsfMain.TextMatrix(vsfMain.Row, SC_��װ�ű�) = ""
            End If
        End If
    End If
End Sub

Private Sub vsfMain_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = Col <> SC_��װ�ű�
End Sub
'==============================================================
'==˽�з���
'==============================================================
Private Sub InitFace()
'���ܣ���ʼ������
    If mintType = 0 Then
        lblFunCap.Caption = "�ռ��Ǽǹ���/����"
        lblFunNote.Caption = "���ݰ�װ�汾��ϵͳ�汾�Ƚϣ��Ӱ�װ�ű��������л�ȡ���µı�׼���̡����±�׼���������ݿ���̽��жԱȣ���þ����û������ı�׼�������û��������Զ�����̡�"
        Me.Caption = "�ռ��Ǽ�"
    Else
         lblFunCap.Caption = "������/��������"
         lblFunNote.Caption = "���ݰ�װ�汾��ϵͳ�汾�Ƚϣ��Ӱ�װ�ű��������л�ȡ���µı�׼���̡����±�׼����������ǰ���ռ����ı�׼���̽��жԱȣ������Ҫ���׼�ű������ı����Ҫ�������û����̡�"
         Me.Caption = "������"
    End If
End Sub

Private Sub LoadData()
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim lngRow As Long
    
    On Error GoTo errH
    strSQL = "Select a.���, a.�汾��, a.���� As ϵͳ����, b.�ļ���" & vbNewLine & _
                    "From Zlsystems a, Zlsysfiles b" & vbNewLine & _
                    "Where a.��� = b.ϵͳ(+) And b.���� = 1" & vbNewLine & _
                    "Order By Nvl(a.�����, 0), a.���"
                    
    Set rsTmp = gclsBase.OpenSQLRecord(mcnOracle, strSQL, "��ȡ��װ�����ļ�")
    With vsfMain
        .Redraw = flexRDNone
        .Rows = vsfMain.FixedRows
        Do While Not rsTmp.EOF
            .Rows = .Rows + 1: lngRow = .Rows - 1
            .TextMatrix(lngRow, SC_���) = lngRow
            .TextMatrix(lngRow, SC_�汾��) = rsTmp!�汾�� & ""
            .TextMatrix(lngRow, SC_ϵͳ����) = rsTmp!ϵͳ���� & ""
            .TextMatrix(lngRow, SC_��װ�ű�) = rsTmp!�ļ��� & ""
            .RowData(lngRow) = Val(rsTmp!��� & "")
            rsTmp.MoveNext
        Loop
        If .Rows <> vsfMain.FixedRows Then
            vsfMain.Row = vsfMain.FixedRows
        End If
        Call vsfMain_AfterRowColChange(-1, -1, 1, 1)
        .Redraw = flexRDDirect
    End With
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, gstrSysName
    If 0 = 1 Then
        Resume
    End If
End Sub

Private Sub SetSerial()
'���ܣ����������
    Dim i As Long
    With vsfMain
        .Redraw = flexRDNone
        For i = .FixedRows To .Rows - 1
            .TextMatrix(i, SC_���) = i
        Next
        If .Rows - 1 > 0 Then
            .Cell(flexcpForeColor, .FixedRows, SC_���, .Rows - 1, SC_���) = Color.���ɫ
            .Cell(flexcpFontBold, .FixedRows, SC_���, .Rows - 1, SC_���) = False
        End If
        If .Row > 0 Then
            .Cell(flexcpFontBold, .Row, SC_���, .Row, SC_���) = True
            .Cell(flexcpForeColor, .Row, SC_���, .Row, SC_���) = Color.��ɫ
        End If
        .Redraw = flexRDDirect
    End With
End Sub

Private Function ValidData() As Boolean
    Dim i As Long, rsInit As ADODB.Recordset
    Dim strPath As String, strCurMax As String
    
    On Error GoTo errH
    Set mrsUpgradeFiles = Nothing
    If mcnOracle Is Nothing Then
        MsgBox "���Ƚ����������ã���ȷ���ռ���Դ��", vbInformation + vbOKOnly, "�������"
        Exit Function
    End If
    pbrCollect.value = 5
    With vsfMain
        If .Rows = .FixedRows Then
            MsgBox "��ǰ���ݿ�û�а�װ�κ�ϵͳ��", vbInformation + vbOKOnly, "�������"
            Exit Function
        End If
        pbrCollect.value = 10
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, SC_��װ�ű�) = "" Then
                MsgBox "��ѡ��" & .TextMatrix(i, SC_ϵͳ����) & "û��ѡ��װ�����ļ�", vbInformation, gstrSysName
                .Row = i: .Col = SC_��װ�ű�: Exit Function
            End If
            If Not CheckInitFile(.RowData(i), .TextMatrix(i, SC_��װ�ű�), False, rsInit, False) Then
                .Row = i: .Col = SC_��װ�ű�: Exit Function
            End If
            rsInit.Filter = "��Ŀ='�汾��'"
            .TextMatrix(i, SC_���ð汾) = rsInit!���� & ""
        Next
        pbrCollect.value = 20
        '��֤��Ǩ�ű����ܷ�֧��ϵͳ��Ǩ
        For i = .FixedRows To .Rows - 1
            '�����ļ��汾��Ӧ��ϵͳ�ߣ����޷�
            If VerFull(GetPrimaryVer(.TextMatrix(i, SC_���ð汾))) > VerFull(.TextMatrix(i, SC_�汾��)) Then
                If MsgBox(.TextMatrix(i, SC_ϵͳ����) & "�İ�װ�ű��汾Ϊ" & GetPrimaryVer(.TextMatrix(i, SC_���ð汾)) & "������ϵͳ��ǰ�汾" & .TextMatrix(i, SC_�汾��) & "���޷���ȡ��׼�ű����Ƿ������", vbInformation + vbYesNo + vbDefaultButton2, "�������") = vbNo Then
                    Exit Function
                End If
            Else
                '��װ�ű���Ӧ�Ĵ�汾�����ʹ��GetPrimaryVer(.TextMatrix(i, SC_���ð汾))�����뵱ǰϵͳ�汾.TextMatrix(i, SC_�汾��)����ȡ�Ӱ�װ�ű��汾����������ǰϵͳ�汾�����������ű����Լ���Щ�ű�֧�ֵ����汾������ȱʧ�ű���Ӧ�뵱ǰϵͳ�汾��ͬ����
                Set mrsUpgradeFiles = GetUpgradeFiles(mrsUpgradeFiles, .RowData(i), GetPrimaryVer(.TextMatrix(i, SC_���ð汾)), .TextMatrix(i, SC_��װ�ű�), , , .TextMatrix(i, SC_�汾��), strCurMax, , True)
                'û�л�ȡ���ű����Ұ�װ�ű��汾��ϵͳ��ǰ�ű�����ͬ������ʾ�ű�ȱʧ
                If strCurMax = "" And GetPrimaryVer(.TextMatrix(i, SC_���ð汾)) <> .TextMatrix(i, SC_�汾��) Then
                    MsgBox .TextMatrix(i, SC_ϵͳ����) & "ȱʧ��" & GetPrimaryVer(.TextMatrix(i, SC_���ð汾)) & "��Ǩ��" & .TextMatrix(i, SC_�汾��) & "�������ű���", vbInformation + vbOKOnly, "�������"
                    Exit Function
                '��ȡ���ű������ǲ�С�ڵ�ǰϵͳ�汾�������ű��޷�֧��ϵͳ��������ǰϵͳ�汾������ʾ�ű�ȱʧ
                ElseIf strCurMax <> "" And strCurMax <> .TextMatrix(i, SC_�汾��) Then
                    MsgBox .TextMatrix(i, SC_ϵͳ����) & "ȱʧ��" & GetPrimaryVer(.TextMatrix(i, SC_���ð汾)) & "��Ǩ��" & .TextMatrix(i, SC_�汾��) & "�������ű���", vbInformation + vbOKOnly, "�������"
                    Exit Function
                End If
            End If
            pbrCollect.value = 20 + (i / (.Rows - 1)) * 70
        Next
        If mrsUpgradeFiles Is Nothing Then
            MsgBox "δ�ռ�����Ҫ�ű����������ȷ�İ�װ�ű��������ű���", vbInformation + vbOKOnly, "�������"
            Exit Function
        End If
        'ֻ����Ӧ��ϵͳ�ı�׼�����ű�
        Call RecDelete(mrsUpgradeFiles, "(SysType<>" & ST_App & ") OR (SysType = " & ST_App & " And FileType<>" & FT_Standard & ")")
    End With
    pbrCollect.value = 100
    ValidData = True
    Exit Function
errH:
    MsgBox err.Description, vbInformation, gstrSysName
    If 0 = 1 Then
        Resume
    End If
End Function

Private Sub DealWithTmpFolder(Optional ByVal blnCreate As Boolean)
'���ܣ�������ʱĿ¼
    'ת��Ϊ��д�Ľű�
    If gobjFile.FolderExists(App.Path & "\BaseProcedure") Then Call gobjFile.DeleteFolder(App.Path & "\BaseProcedure", True)
    pbrCollect.value = 16 * IIf(blnCreate, 1, 2)
    If gobjFile.FolderExists(App.Path & "\ComProcedure") Then Call gobjFile.DeleteFolder(App.Path & "\ComProcedure")
    pbrCollect.value = 16 * IIf(blnCreate, 1, 2) * 2
    '�Աȱ���
    If gobjFile.FolderExists(App.Path & "\Reports") And blnCreate Then gobjFile.DeleteFolder (App.Path & "\Reports")
    pbrCollect.value = 50 * IIf(blnCreate, 1, 2)
    If blnCreate Then
        Call gobjFile.CreateFolder(App.Path & "\BaseProcedure")
        pbrCollect.value = 16.5 * 4
        Call gobjFile.CreateFolder(App.Path & "\ComProcedure")
        pbrCollect.value = 16.5 * 5
        Call gobjFile.CreateFolder(App.Path & "\Reports")
        pbrCollect.value = 100
    End If
End Sub

Private Sub ShowState(Optional ByVal strInfo As String, Optional ByVal blnEnd As Boolean)
    lblTitle.Caption = strInfo
    lblTitle.Tag = strInfo
    lblTitle.Visible = strInfo <> ""
    pbrCollect.Visible = strInfo <> ""
    pbrCollect.value = 0
    cmdStart.Enabled = blnEnd
    cmdDel.Enabled = blnEnd
End Sub

Private Function LoadBaseProcs(ByVal strPath As String) As Boolean
    '���ܣ��������ݿ�洢����
    Dim rsSource As ADODB.Recordset, strSQL As String
    Dim objText As TextStream, strProcName As String, strProcText As String
    Dim objPercent As New clsPercent
    
    On Error GoTo errH
    '�洢�����ռ����ռ����ݿ���Ϊ�����洢����
    If mintType = 0 Then
        strSQL = "Select Name, Type, Text, Line ��� From User_Source Where Type In ('PROCEDURE', 'FUNCTION') Order By Name, Line"
        Set rsSource = gclsBase.OpenSQLRecord(mcnOracle, strSQL, "��ȡ���ݿ����Դ��")
    '����Աȣ����ϴα�׼�洢����Ϊ�����洢����
    Else
        strSQL = "Select a.Id, Upper(a.����) Name, b.���, b.���� Text" & vbNewLine & _
                        "From Zlprocedure a, Zlproceduretext b" & vbNewLine & _
                        "Where a.Id = b.����id And b.���� = " & ProcTextType.���α�׼���� & " And a.���� In (1, 2)" & vbNewLine & _
                        "Order By a.Id, b.���"
        Set rsSource = gclsBase.OpenSQLRecord(mcnOracle, strSQL, "��ȡ���ݿ����Դ��")
    End If
    If Not rsSource.EOF Then
        pbrCollect.Visible = True
        Call objPercent.InitPercent(pbrCollect, rsSource.RecordCount)
        Do While Not rsSource.EOF
            If strProcName <> rsSource!name & "" Then
                If strProcName <> "" Then
                    '���ݿ�Դ��û��CREATE OR REPLACE
                    If mintType = 0 Then
                        strProcText = "CREATE OR REPLACE " & strProcText
                    End If
                    '�����������̽ű��ļ�
                    Set objText = gobjFile.CreateTextFile(strPath & "\" & strProcName & ".sql", True)
                    objText.Write strProcText
                End If
                strProcName = rsSource!name & ""
                strProcText = ""
            End If
            If rsSource!��� = 1 Then
                '���ƴ�˫���ţ���ȥ��
                If UCase(rsSource!Text) Like "*" & """" & UCase(strProcName) & """" & "*" Then
                    strProcText = strProcText & Replace(UCase(rsSource!Text), """" & UCase(strProcName) & """", strProcName)
                Else
                    strProcText = strProcText & rsSource!Text
                End If
            Else
                strProcText = strProcText & rsSource!Text
            End If
            rsSource.MoveNext
            Call objPercent.LoopPercent
        Loop
        If strProcName <> "" Then
            '�����������̽ű��ļ�
            Set objText = gobjFile.CreateTextFile(strPath & "\" & strProcName & ".sql", True)
            objText.Write strProcText
        End If
        objText.Close
        pbrCollect.Visible = False
    End If
    LoadBaseProcs = True
    Exit Function
errH:
    MsgBox err.Description, vbInformation, gstrSysName
    If 0 = 1 Then
        Resume
    End If
End Function

Private Function LoadComProcs(ByVal strPath As String) As Boolean
'���ܣ��������ݿ�洢����
    Dim i As Long, strFile As String
    Dim objPercent As New clsPercent
    On Error GoTo errH
    With vsfMain
        mrsUpgradeFiles.Filter = ""
        mrsUpgradeFiles.Sort = "ϵͳ���,FullSPVer"
        Call objPercent.InitPercent(pbrCollect, mrsUpgradeFiles.RecordCount + .Rows - 1)
        For i = .FixedRows To .Rows - 1
            lblTitle.Caption = lblTitle.Tag & "    ������ȡ��" & .TextMatrix(i, SC_ϵͳ����) & "����װ�ű�.."
            strFile = gobjFile.GetParentFolderName(.TextMatrix(i, SC_��װ�ű�)) & "\zlProgram.sql"
            Call LoadProcFile(strFile, strPath)
            objPercent.LoopPercent
            mrsUpgradeFiles.Filter = "ϵͳ���=" & .RowData(i)
            mrsUpgradeFiles.Sort = "FullSPVer"
            lblTitle.Caption = lblTitle.Tag & "    ������ȡ" & .TextMatrix(i, SC_ϵͳ����) & "�����ű�.."
            Do While Not mrsUpgradeFiles.EOF
                Call LoadProcFile(mrsUpgradeFiles!FilePath, strPath)
                objPercent.LoopPercent
                mrsUpgradeFiles.MoveNext
            Loop
        Next
    End With
    LoadComProcs = True
    Exit Function
errH:
    MsgBox err.Description, vbInformation, gstrSysName
    If 0 = 1 Then
        Resume
    End If
End Function

Private Function LoadProcFile(ByVal strFile As String, ByVal strFilePath As String) As Boolean
'���ܣ����ű��д��ڵĴ洢�����뺯���ֱ���Ϊ�ļ��洢��
    Dim objScript As New clsRunScript
    Dim objText As TextStream
    Dim objPercent As New clsPercent
    
    With objScript
         If .OpenFile(strFile) And Not .EOF Then
            Do While Not .EOF
                If .SQLInfo.Block = True Then
                    If .SQLInfo.BlockType = "PROCEDURE" Or .SQLInfo.BlockType = "FUNCTION" Then
                        Set objText = gobjFile.CreateTextFile(strFilePath & "\" & .SQLInfo.BlockName & ".sql", True)
                        objText.Write .SQLInfo.SQL
                        objText.Close
                    End If
                End If
                Call .ReadNextSQL
            Loop
         End If
    End With
    LoadProcFile = True
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    MsgBox err.Description, vbInformation, gstrSysName
End Function

Private Sub CreateResultSQL(ByVal strBasePath As String, strComPath As String, ByVal strRportPth As String, ByRef arrSQL As Variant)
    Dim objFolder As Folder, objFile As File
    Dim strFileName As String
    Dim rsObjectInfo As ADODB.Recordset, strSQL As String
    Dim lngKey As Long, pt As ProcType, strNote As String, strOwner As String
    Dim objPercent As New clsPercent
    Dim rsSouce As ADODB.Recordset
    Dim strTMp As String
    Dim objLog As TextStream
    
    On Error GoTo errH
    If mintType = 0 Then
        '�����д��ڵļ�Ϊ��Ҫ�����Ĺ���
        lblTitle.Caption = lblTitle.Tag & "    ���ڲ����䶯����.."
        Call objPercent.InitPercent(pbrCollect, gobjFile.GetFolder(strRportPth).Files.Count + gobjFile.GetFolder(strBasePath).Files.Count)
        Set objFolder = gobjFile.GetFolder(strRportPth)
        strSQL = "Select b.Owner, b.Object_Name, c.Id, c.����, c.����, c.״̬, c.˵��, c.�޸���Ա, c.�޸�ʱ��, c.�ϴ��޸���Ա, c.�ϴ��޸�ʱ��" & vbNewLine & _
                        "From (Select a.Owner, a.Object_Name" & vbNewLine & _
                        "       From All_Objects a" & vbNewLine & _
                        "       Where a.Object_Type In ('PROCEDURE', 'FUNCTION') And a.Owner In (Select Distinct ������ From Zlsystems)) b," & vbNewLine & _
                        "     Zlprocedure c" & vbNewLine & _
                        "Where b.Object_Name = Upper(c.����(+)) " & vbNewLine & _
                        "Order by Object_Name,Owner,Id"
        Set rsObjectInfo = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "��ȡ������Ϣ")
        
        Call gclsBase.addItem(arrSQL, "Zl_Zlprocedure_Manage(0)")
        For Each objFile In objFolder.Files
            strFileName = Split(objFile.name, ".")(0): strOwner = "": lngKey = 0: pt = ProcType.�䶯����
            rsObjectInfo.Filter = "Object_Name='" & UCase(strFileName) & "'"
            If Not rsObjectInfo.EOF Then
                lblTitle.Caption = lblTitle.Tag & "    ���ڲ����䶯���̣�" & strFileName
                strOwner = rsObjectInfo!Owner & "": lngKey = Val(rsObjectInfo!Id & "")
                '���̲����ڣ��Զ����Ϊ�䶯���̻�հ׹���
                If lngKey = 0 Then
                    lngKey = gclsBase.GetNextId("zlProcedure")
                ElseIf Val(rsObjectInfo!���� & "") = 2 Then
                    pt = ProcType.�հ׹���
                End If
                Call gclsBase.addItem(arrSQL, "Zl_Zlprocedure_Update(" & lngKey & "," & pt & ",'" & strFileName & "'," & ProcState.����� & ",'" & rsObjectInfo!˵�� & "','" & strOwner & "')")
                '���汾���Զ������
                Call gclsBase.GetProcSQL(lngKey, ProcTextType.�����Զ�����, strBasePath & "\" & strFileName & ".sql", arrSQL, True)
                '���汾�α�׼����
                Call gclsBase.GetProcSQL(lngKey, ProcTextType.���α�׼����, strComPath & "\" & strFileName & ".sql", arrSQL, True)
            End If
            objPercent.LoopPercent
        Next
        
        lblTitle.Caption = lblTitle.Tag & "     ���ڲ����û�����.."
        Set objFolder = gobjFile.GetFolder(strBasePath)
        For Each objFile In objFolder.Files
            '���ݿ��еĹ����ڽű���û�У�˵�����û�����
            If Not gobjFile.FileExists(strComPath & "\" & objFile.name) Then
                strFileName = Split(objFile.name, ".")(0)
                If Not UCase(strFileName) Like "ZL*_UPGRADECHECK" Then '��Ǩ��麯���ų�
                    strOwner = "": lngKey = 0: pt = ProcType.�û�����
                    rsObjectInfo.Filter = "Object_Name='" & UCase(strFileName) & "'"
                    If Not rsObjectInfo.EOF Then
                        lblTitle.Caption = lblTitle.Tag & "    ���ڲ����û����̣�" & strFileName
                        strOwner = rsObjectInfo!Owner & "": lngKey = Val(rsObjectInfo!Id & "")
                        If lngKey = 0 Then '���̲����ڣ��Զ����Ϊ�û�����
                            lngKey = gclsBase.GetNextId("zlProcedure")
                        End If
                        Call gclsBase.addItem(arrSQL, "Zl_Zlprocedure_Update(" & lngKey & "," & pt & ",'" & strFileName & "'," & ProcState.����� & ",'" & rsObjectInfo!˵�� & "','" & strOwner & "')")
                        '���汾���Զ������
                        Call gclsBase.GetProcSQL(lngKey, ProcTextType.�����Զ�����, strBasePath & "\" & objFile.name, arrSQL, True)
                    End If
                End If
            End If
            Call objPercent.LoopPercent
        Next
        '���Ի������ռ�ȱʧ�ĺ���
        If gblnInIDE Then
            If Not gobjFile.FolderExists("C:\AppSoft\Log\���̹���\") Then
                gobjFile.CreateFolder ("C:\AppSoft\Log\���̹���\")
            End If
            Set objLog = gobjFile.CreateTextFile("C:\AppSoft\Log\���̹���\Proc_ȱʧ����.ini", True)
            Set objFolder = gobjFile.GetFolder(strComPath)
            For Each objFile In objFolder.Files
                If Not gobjFile.FileExists(strBasePath & "\" & objFile.name) Then
                    objLog.WriteLine Split(objFile.name, ".")(0)
                End If
            Next
        End If
    Else
        lblTitle.Caption = lblTitle.Tag & "     ���ڵ�������״̬.."
        strSQL = "Select ID,����,����,˵��,������ From zlprocedure"
        Set rsObjectInfo = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "��ȡ�����б�")
        Call objPercent.InitPercent(pbrCollect, rsObjectInfo.RecordCount)
        Do While Not rsObjectInfo.EOF
            strFileName = rsObjectInfo!���� & ""
            If Val(rsObjectInfo!���� & "") <> ProcType.�û����� Then
                pt = Val(rsObjectInfo!���� & "")
                If gobjFile.FileExists(strComPath & "\" & strFileName & ".sql") Then
                    '��׼����������ǰ���б仯
                    If gobjFile.FileExists(strRportPth & "\" & strFileName & ".sql.htm") Then
                        Call gclsBase.addItem(arrSQL, "Zl_Zlprocedure_Update(" & rsObjectInfo!Id & "," & pt & ",'" & strFileName & "'," & ProcState.������ & ",'" & rsObjectInfo!˵�� & "','" & rsObjectInfo!������ & "')")
                    '��׼����������ǰ���ޱ仯
                    Else
                        Call gclsBase.addItem(arrSQL, "Zl_Zlprocedure_Update(" & rsObjectInfo!Id & "," & pt & ",'" & strFileName & "'," & ProcState.�ޱ仯 & ",'" & rsObjectInfo!˵�� & "','" & rsObjectInfo!������ & "')")
                    End If
                    '���汾�α�׼����
                    Call gclsBase.GetProcSQL(Val(rsObjectInfo!Id & ""), ProcTextType.���α�׼����, strComPath & "\" & strFileName & ".sql", arrSQL, True)
                End If
            Else '�û����̣�����Ϊ�䶯����
                If gobjFile.FileExists(strComPath & "\" & strFileName & ".sql") Then '�û����̴����ڱ�׼�ű�����Ӧ�Զ�����Ϊ�䶯����
                    pt = ProcType.�䶯����: strOwner = ""
                    strTMp = gclsBase.GetProgram(strFileName, strOwner)
                    If strTMp = "" Then
                        strOwner = rsObjectInfo!������ & ""
                    End If
                    Call gclsBase.addItem(arrSQL, "Zl_Zlprocedure_Update(" & Val(rsObjectInfo!Id & "") & "," & pt & ",'" & strFileName & "'," & ProcState.������ & ",'" & rsObjectInfo!˵�� & "','" & strOwner & "')")
                    If strTMp <> "" Then
                        Call gclsBase.GetProcSQL(lngKey, ProcTextType.�����Զ�����, strTMp, arrSQL)
                    End If
                    '���汾�α�׼����
                    Call gclsBase.GetProcSQL(Val(rsObjectInfo!Id & ""), ProcTextType.���α�׼����, strComPath & "\" & strFileName & ".sql", arrSQL, True)
                End If
            End If
            Call objPercent.LoopPercent
            rsObjectInfo.MoveNext
        Loop
        Call gclsBase.addItem(arrSQL, "Zl_Zlprocedure_Manage(1)")
    End If
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, gstrSysName
    If 0 = 1 Then
        Resume
    End If
End Sub

