VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmHealthArchives 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���񽡿�������Ϣ"
   ClientHeight    =   10485
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12000
   Icon            =   "frmHealthArchives.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10485
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdExit 
      Caption         =   "�˳�(&X)"
      Height          =   350
      Left            =   10485
      TabIndex        =   4
      Top             =   360
      Width           =   1100
   End
   Begin VB.PictureBox picPatiInfor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   9495
      Left            =   210
      ScaleHeight     =   9465
      ScaleWidth      =   9675
      TabIndex        =   0
      Top             =   210
      Width           =   9705
      Begin VSFlex8Ctl.VSFlexGrid vsGrid 
         Height          =   8850
         Left            =   -15
         TabIndex        =   1
         Top             =   -15
         Width           =   9765
         _cx             =   17224
         _cy             =   15610
         Appearance      =   0
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   16777215
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483639
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   23
         Cols            =   9
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   350
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmHealthArchives.frx":0442
         ScrollTrack     =   0   'False
         ScrollBars      =   0
         ScrollTips      =   0   'False
         MergeCells      =   1
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
         Begin VB.PictureBox picPhoto 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   1740
            Left            =   6645
            ScaleHeight     =   1710
            ScaleWidth      =   3030
            TabIndex        =   2
            Top             =   -15
            Width           =   3060
            Begin VB.Image imgPhoto 
               Height          =   435
               Left            =   1185
               Stretch         =   -1  'True
               Top             =   765
               Width           =   315
            End
         End
      End
   End
   Begin XtremeSuiteControls.TaskPanel wndTaskPanel 
      Height          =   10290
      Left            =   120
      TabIndex        =   3
      Top             =   60
      Width           =   10125
      _Version        =   589884
      _ExtentX        =   17859
      _ExtentY        =   18150
      _StockProps     =   64
      VisualTheme     =   6
      ItemLayout      =   2
      HotTrackStyle   =   1
   End
End
Attribute VB_Name = "frmHealthArchives"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const M_IDX_TP_BASE = 100
Private mlng����ID As Long
Private mrsInfor As ADODB.Recordset '���˻�����Ϣ
Private mrsOtherCertificate As ADODB.Recordset
Private mrsDrug As ADODB.Recordset
Private mrsBacterin As ADODB.Recordset
Private mblnUnLoad As Boolean
Private mcnOracle As ADODB.Connection
Private mobjDataBase As clsDataBase
Public Sub zlShowHealthArchives(ByVal frmMain As Object, ByVal lng����ID As Long, _
    Optional cnOracle As ADODB.Connection)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʾ������Ϣ
    '���:lng����ID-����ID
    '     cnOracle-���ݿ�����
    '����:���˺�
    '����:2012-12-14 13:34:29
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mlng����ID = lng����ID
    Set mcnOracle = cnOracle
    If zlGetOneDataBase(cnOracle, mobjDataBase) = False Then Exit Sub
    
    If frmMain Is Nothing Then
        Me.Show 1
    Else
        Me.Show 1, frmMain
    End If
End Sub
Private Function LoadPatiInfor() As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ز�����Ϣ
    '����:���ڲ�����Ϣ,����true,���򷵻�False
    '����:���˺�
    '����:2012-12-14 13:38:09
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String
    Err = 0: On Error GoTo errHandle
    
    
    '82072:���ϴ�,2015/1/23,Ѫ�ͺ�RHȡ����ID Ϊnull�ļ�¼
    strSql = "" & _
    "  Select a.����id,a.����,to_char(a.��������,'yyyy-mm-dd') as ��������,a.�Ա�,a.����,a.����,a.����״��,a.ѧ��,a.��ͥ�绰,a.ְҵ,max(a.���֤��) as ���֤��, " & _
    "          max(a.��ϵ������) as ��ϵ������1,max(a.��ϵ�˹�ϵ) as ��ϵ�˹�ϵ1,max(a.��ϵ�˵绰) as ��ϵ�˵绰1, " & _
    "          max(a.���ڵ�ַ) as ���ڵ�ַ,max(a.��ͥ��ַ) as ��ͥ��ַ, " & _
    "          max(decode(b.��Ϣ��,'��ϵ������1',b.��Ϣֵ,'')) as ��ϵ������2,max(decode(b.��Ϣ��,'��ϵ�˹�ϵ1',b.��Ϣֵ,'')) as  ��ϵ�˹�ϵ2,max(decode(b.��Ϣ��,'��ϵ�˵绰1',b.��Ϣֵ,''))   as ��ϵ�˵绰2, " & _
    "          max(decode(b.��Ϣ��,'��ϵ������2',b.��Ϣֵ,'')) as ��ϵ������3,max(decode(b.��Ϣ��,'��ϵ�˹�ϵ2',b.��Ϣֵ,'')) as  ��ϵ�˹�ϵ3,max(decode(b.��Ϣ��,'��ϵ�˵绰2',b.��Ϣֵ,''))   as ��ϵ�˵绰3, " & _
    "          max(decode(b.��Ϣ��,'��ũ��(��)��',b.��Ϣֵ,'')) as  ��ũ�Ϻ�, max(decode(b.��Ϣ��,'�����������',b.��Ϣֵ,'')) as  �����������, " & _
    "          max(decode(b.��Ϣ��,'ҽ�Ʒ���֧����ʽ',b.��Ϣֵ,'')) as  ����֧����ʽ, " & _
    "          max(decode(b.����id,null,decode(b.��Ϣ��,'ABO',b.��Ϣֵ,'Ѫ��',b.��Ϣֵ,''),'')) as  ABOѪ��, " & _
    "          max(decode(b.����id,null,decode(b.��Ϣ��,'RH',b.��Ϣֵ,''),'')) as  RH," & _
    "          max(decode(b.��Ϣ��,'ҽѧ��ʾ',b.��Ϣֵ,'')) as  ҽѧ��ʾ," & _
    "          max(decode(b.��Ϣ��,'����ҽѧ��ʾ',b.��Ϣֵ,'')) as  ����ҽѧ��ʾ" & _
    "   From ������Ϣ A,������Ϣ�ӱ� B  " & _
    "   Where  a.����ID=b.����ID(+) and a.����ID=[1] " & _
    "   Group by a.����ID,a.����,a.��������,a.�Ա�,a.����,a.����,a.����״��,a.ѧ��,a.��ͥ�绰,a.ְҵ"
    
    Set mrsInfor = mobjDataBase.OpenSQLRecord(strSql, Me.Caption, mlng����ID)
    If mrsInfor.RecordCount = 0 Then
        MsgBox "��ǰ������Ϣ�����ڣ����飡", vbInformation, gstrSysName
        Exit Function
    End If
    
    '��ȡ����֤��
    strSql = "Select a.��Ϣ�� as ֤������, a.��Ϣֵ as ֤������ From ������Ϣ�ӱ� A, ֤������ B Where a.����id = [1] And a.��Ϣ�� = b.����  Order by ��Ϣ��"
    Set mrsOtherCertificate = mobjDataBase.OpenSQLRecord(strSql, Me.Caption, mlng����ID)
    
    '��ȡ����ҩ��
    strSql = "  Select ����ҩ��,������Ӧ from ���˹���ҩ�� where ����ID=[1] order by ����ҩ��"
    Set mrsDrug = mobjDataBase.OpenSQLRecord(strSql, Me.Caption, mlng����ID)
    
    '��ȡ����������
    strSql = " " & _
    "   Select ����id, To_char(����ʱ��, 'yyyy-mm-dd hh24:mi') As ����ʱ��, �������� " & _
    "   From �������߼�¼ " & _
    "   Where ����id = [1] " & _
    "   Order By ����ʱ��"
    Set mrsBacterin = mobjDataBase.OpenSQLRecord(strSql, Me.Caption, mlng����ID)
    LoadPatiInfor = True
    Exit Function
errHandle:
    If mobjDataBase.ErrCenter() = 1 Then
        Resume
    End If
    
End Function

Private Sub InitTaskPancel()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ��InitTaskPancel
    '����:���˺�
    '����:2012-12-13 17:48:22
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim tkpGroup As TaskPanelGroup
    Dim Item As TaskPanelGroupItem
    'Call wndTaskPanel.SetGroupInnerMargins(2, 0, 2, 0)
    
    Call wndTaskPanel.SetGroupOuterMargins(2, -10, 2, -10)
      Call wndTaskPanel.SetMargins(2, 16, 2, 10, 30)
    wndTaskPanel.HotTrackStyle = xtpTaskPanelHighlightItem
    Set tkpGroup = wndTaskPanel.Groups.Add(M_IDX_TP_BASE, "���˽���������Ϣ")
    Set Item = tkpGroup.Items.Add(M_IDX_TP_BASE, "", xtpTaskItemTypeControl)
    Set Item.Control = picPatiInfor
    picPatiInfor.BackColor = Item.BackColor
    tkpGroup.Expandable = False
    wndTaskPanel.Reposition
    wndTaskPanel.DrawFocusRect = True
End Sub

Private Sub InitGrid()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ������
    '����:���˺�
    '����:2012-12-13 18:35:13
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, r As Long
    
    On Error GoTo errHandle
    
    With vsGrid
        .Clear
        .MergeCells = flexMergeFree
        .Rows = 27: .Cols = 9
        .RowHeightMin = 350
        .FixedRows = 0: .FixedCols = 0
        For i = 0 To .Cols - 1
            .ColAlignment(i) = flexAlignCenterCenter
            .FixedAlignment(i) = flexAlignCenterCenter
        Next
        .RowHidden(0) = True
        .TextMatrix(1, 1) = "����"
        .TextMatrix(1, 2) = "����"
        For i = 3 To 6
            .TextMatrix(1, i) = NVL(mrsInfor!����, " ")
        Next
        .Cell(flexcpAlignment, 1, 3, 1, 6) = 1
        .Cell(flexcpForeColor, 1, 3, 1, 6) = vbBlue
        .TextMatrix(2, 1) = "��������"
        .TextMatrix(2, 2) = "��������"
        For i = 3 To 6
            .TextMatrix(2, i) = NVL(mrsInfor!��������, "  ")
        Next
        .Cell(flexcpAlignment, 2, 3, 2, 6) = 1
        .Cell(flexcpForeColor, 2, 3, 2, 6) = vbBlue
        
        .TextMatrix(3, 1) = "�Ա�"
        .TextMatrix(3, 2) = "�Ա�"
        .TextMatrix(3, 3) = NVL(mrsInfor!�Ա�, "   ")
        .Cell(flexcpAlignment, 3, 3, 3, 3) = 1
        .Cell(flexcpForeColor, 3, 3, 3, 3) = vbBlue
        
        .TextMatrix(3, 4) = "����"
        .TextMatrix(3, 5) = NVL(mrsInfor!����, "   ")
        .Cell(flexcpForeColor, 3, 5, 3, 5) = vbBlue
        .TextMatrix(3, 6) = .TextMatrix(3, 5)
        .Cell(flexcpForeColor, 3, 5, 3, 6) = vbBlue
        .Cell(flexcpAlignment, 3, 5, 3, 6) = 1
        
                
        .TextMatrix(4, 1) = "����״��"
        .TextMatrix(4, 2) = "����״��"
        .TextMatrix(4, 3) = NVL(mrsInfor!����״��)
        .Cell(flexcpForeColor, 4, 3, 4, 3) = vbBlue
        .Cell(flexcpAlignment, 4, 3, 4, 3) = 1
        .TextMatrix(4, 4) = "�Ļ��̶�"
        .TextMatrix(4, 5) = NVL(mrsInfor!ѧ��, "  ")
        .TextMatrix(4, 6) = .TextMatrix(4, 5)
        .Cell(flexcpAlignment, 4, 5, 4, 6) = 1
        .Cell(flexcpForeColor, 4, 5, 4, 6) = vbBlue
        .Cell(flexcpAlignment, 5, 3, 8, 3) = 1

        
        .TextMatrix(5, 1) = "���˵绰"
        .TextMatrix(5, 2) = "���˵绰"
        .TextMatrix(5, 3) = NVL(mrsInfor!��ͥ�绰, " ")
        .Cell(flexcpForeColor, 5, 3, 5, 3) = vbBlue
        .TextMatrix(5, 4) = "ְҵ"
        .TextMatrix(5, 5) = NVL(mrsInfor!ְҵ, "   ")
        .TextMatrix(5, 6) = .TextMatrix(5, 5)
        .Cell(flexcpAlignment, 5, 5, 5, 6) = 1
        .Cell(flexcpForeColor, 5, 5, 5, 6) = vbBlue
                
        .TextMatrix(6, 1) = "��ϵ��"
        .TextMatrix(7, 1) = "��ϵ��"
        .TextMatrix(8, 1) = "��ϵ��"
        
        .TextMatrix(6, 2) = "����1"
        .TextMatrix(6, 3) = NVL(mrsInfor!��ϵ������1, "")
        .Cell(flexcpAlignment, 6, 3, 6, 3) = 1
        .Cell(flexcpForeColor, 6, 3, 6, 3) = vbBlue
        
        .TextMatrix(7, 2) = "����2"
        .TextMatrix(7, 3) = NVL(mrsInfor!��ϵ������2, "  ")
        .Cell(flexcpAlignment, 7, 3, 7, 3) = 1
        .Cell(flexcpForeColor, 7, 3, 7, 3) = vbBlue
        
        .TextMatrix(8, 2) = "����3"
        .TextMatrix(8, 3) = NVL(mrsInfor!��ϵ������3, "")
        .Cell(flexcpForeColor, 8, 3, 8, 3) = vbBlue
        .Cell(flexcpAlignment, 8, 3, 8, 3) = 1
        .Cell(flexcpAlignment, 6, 5, 8, 5) = 1
        .Cell(flexcpForeColor, 6, 5, 8, 5) = vbBlue
        
        
        .TextMatrix(6, 4) = "��ϵ1"
        .TextMatrix(6, 5) = NVL(mrsInfor!��ϵ�˹�ϵ1, "")
        .TextMatrix(7, 4) = "��ϵ2"
        .TextMatrix(7, 5) = NVL(mrsInfor!��ϵ�˹�ϵ2, " ")
        .TextMatrix(8, 4) = "��ϵ3"
        .TextMatrix(8, 5) = NVL(mrsInfor!��ϵ�˹�ϵ3, "")
        
        .TextMatrix(6, 6) = "�绰1"
        .TextMatrix(7, 6) = "�绰2"
        .TextMatrix(8, 6) = "�绰3"
        
        .Cell(flexcpAlignment, 6, 7, 8, 8) = 1
        .Cell(flexcpForeColor, 6, 7, 8, 8) = vbBlue
        For i = 7 To 8
            .TextMatrix(6, i) = NVL(mrsInfor!��ϵ�˵绰1, " ")
        Next
        For i = 7 To 8
            .TextMatrix(7, i) = NVL(mrsInfor!��ϵ�˵绰2, "  ")
        Next
        For i = 7 To 8
            .TextMatrix(8, i) = NVL(mrsInfor!��ϵ�˵绰3, "   ")
        Next
        
        For i = 9 To 12
            .TextMatrix(i, 1) = "��ݱ�ʶ"
            .TextMatrix(i, 2) = "��ݱ�ʶ"
        Next
                        
        .TextMatrix(9, 3) = "���֤"
        .Cell(flexcpAlignment, 9, 4, 9, .Cols - 1) = 1
        .Cell(flexcpForeColor, 9, 4, 9, .Cols - 1) = vbBlue
        For i = 4 To .Cols - 1
            .TextMatrix(9, i) = NVL(mrsInfor!���֤��, " ")
        Next
        
        .TextMatrix(10, 3) = "����֤��"
        .TextMatrix(10, 6) = "֤������"
        .Cell(flexcpAlignment, 10, 4, 10, 5) = 1
        .Cell(flexcpAlignment, 10, 7, 10, 8) = 1
        .Cell(flexcpForeColor, 10, 4, 10, 5) = vbBlue
        .Cell(flexcpForeColor, 10, 7, 10, 8) = vbBlue
        If mrsOtherCertificate.RecordCount > 0 Then
            .TextMatrix(10, 4) = NVL(mrsOtherCertificate!֤������)
            .TextMatrix(10, 5) = NVL(mrsOtherCertificate!֤������)
            
            .TextMatrix(10, 7) = NVL(mrsOtherCertificate!֤������)
            .TextMatrix(10, 8) = .TextMatrix(10, 7)
        Else
            .TextMatrix(10, 4) = "     "
            .TextMatrix(10, 5) = "     "
            For i = 7 To .Cols - 1
                .TextMatrix(10, i) = "  "
            Next
        End If
                
        .TextMatrix(11, 3) = "��ũ��֤(��)��"
        .Cell(flexcpAlignment, 11, 4, 11, 5) = 1
        .Cell(flexcpForeColor, 11, 4, 11, 5) = vbBlue
        For i = 4 To 5
            .TextMatrix(11, i) = NVL(mrsInfor!��ũ�Ϻ�, "  ")
        Next
        
        .TextMatrix(11, 6) = "�����������"
        .Cell(flexcpAlignment, 11, 7, 11, 8) = 1
        .Cell(flexcpForeColor, 11, 7, 11, 8) = vbBlue
        For i = 7 To .Cols - 1
            .TextMatrix(11, i) = NVL(mrsInfor!�����������, "   ")
        Next
        
        .TextMatrix(12, 1) = "������ַ"
        .TextMatrix(12, 2) = "������ַ"
        .TextMatrix(12, 3) = NVL(mrsInfor!���ڵ�ַ, " ")
        .Cell(flexcpForeColor, 12, 3, 12, .Cols - 1) = vbBlue
        .Cell(flexcpAlignment, 12, 3, 12, .Cols - 1) = 1
        For i = 4 To .Cols - 1
            .TextMatrix(12, i) = .TextMatrix(12, 3)
        Next
        
        .TextMatrix(13, 1) = "��ס��ַ"
        .TextMatrix(13, 2) = "��ס��ַ"
        .TextMatrix(13, 3) = NVL(mrsInfor!��ͥ��ַ, "  ")
        .Cell(flexcpAlignment, 13, 3, 13, .Cols - 1) = 1
        .Cell(flexcpForeColor, 13, 3, 13, .Cols - 1) = vbBlue
        For i = 4 To .Cols - 1
            .TextMatrix(13, i) = .TextMatrix(13, 3)
        Next
        
        .TextMatrix(14, 1) = "ҽ�Ʒ���֧����ʽ"
        .TextMatrix(14, 2) = "ҽ�Ʒ���֧����ʽ"
        .TextMatrix(14, 3) = NVL(mrsInfor!����֧����ʽ, " ")
        .RowHeight(14) = 600
        .Cell(flexcpAlignment, 14, 3, 14, .Cols - 1) = 1
        .Cell(flexcpForeColor, 14, 3, 14, .Cols - 1) = vbBlue
        For i = 4 To .Cols - 1
            .TextMatrix(14, i) = .TextMatrix(14, 3)
        Next
                
        For i = 1 To 14
            .TextMatrix(i, 0) = "���ʶ������"
        Next
        For i = 15 To .Rows - 1
            .TextMatrix(i, 0) = "������������"
        Next
        
        .TextMatrix(15, 1) = "�����ʶ"
        .TextMatrix(15, 2) = "�����ʶ"
        .TextMatrix(15, 3) = "ABOѪ��"
        .TextMatrix(15, 4) = NVL(mrsInfor!ABOѪ��, "  ")
        .TextMatrix(15, 5) = .TextMatrix(15, 4)
        .Cell(flexcpAlignment, 15, 4, 15, 5) = 1
        .Cell(flexcpForeColor, 15, 4, 15, 5) = vbBlue
        
        .TextMatrix(15, 6) = "RH"
        .TextMatrix(15, 7) = NVL(mrsInfor!RH, "    ")
        .TextMatrix(15, 8) = .TextMatrix(15, 7)
        .Cell(flexcpAlignment, 15, 7, 15, 8) = 1
        .Cell(flexcpForeColor, 15, 7, 15, 8) = vbBlue
        
       For r = 16 To 21
            .TextMatrix(r, 1) = "ҽԺ��ʾ"
            .TextMatrix(r, 2) = "ҽԺ��ʾ"
        Next
        .TextMatrix(16, 3) = NVL(mrsInfor!ҽѧ��ʾ, " ")
        .TextMatrix(16, 3) = IIf(Trim(.TextMatrix(16, 3)) = "", "", Trim(.TextMatrix(16, 3)) & ";") & NVL(mrsInfor!����ҽѧ��ʾ, " ")
        .Cell(flexcpAlignment, 16, 3, 16, .Cols - 1) = 1
        .Cell(flexcpForeColor, 16, 3, 16, .Cols - 1) = vbBlue
        
        For i = 4 To .Cols - 1
            .TextMatrix(16, i) = .TextMatrix(16, 3)
        Next
        .RowHeight(16) = 600
        
        r = 17
        .Cell(flexcpBackColor, r, 3, r, .Cols - 1) = &HFFC0C0
        .Cell(flexcpBackColor, r + 1, 3, r + 1, .Cols - 1) = &H8000000F
        For i = 3 To .Cols - 1
            .TextMatrix(r, i) = "����ҩ��"
        Next
        .TextMatrix(18, 3) = "ҩ������"
        .TextMatrix(18, 4) = "ҩ������"
        .TextMatrix(18, 5) = "ҩ�ﷴӦ"
        .TextMatrix(18, 6) = "ҩ�ﷴӦ"
        .TextMatrix(18, 7) = "ҩ�ﷴӦ"
        .TextMatrix(18, 8) = "ҩ�ﷴӦ"
        .Cell(flexcpAlignment, 19, 3, 21, 8) = 1
        .Cell(flexcpForeColor, 19, 3, 21, .Cols - 1) = vbBlue
        
        r = 19
        Do While Not mrsDrug.EOF
            If r > 21 Then Exit Do
            .TextMatrix(r, 3) = NVL(mrsDrug!����ҩ��) & Space(r - 19 + 1)
            .TextMatrix(r, 4) = .TextMatrix(r, 3)
            .TextMatrix(r, 5) = NVL(mrsDrug!������Ӧ) & Space(r - 19 + 1)
            .TextMatrix(r, 6) = .TextMatrix(r, 5)
            .TextMatrix(r, 7) = .TextMatrix(r, 5)
            .TextMatrix(r, 8) = .TextMatrix(r, 5)
            r = r + 1
            mrsDrug.MoveNext
        Loop
        If r <= 21 Then
            For i = r To 21
                .TextMatrix(i, 3) = Space(i - 19 + 1)
                .TextMatrix(i, 4) = Space(i - 19 + 1)
                
                .TextMatrix(i, 5) = Space(i - 19 + 2)
                .TextMatrix(i, 6) = Space(i - 19 + 2)
                .TextMatrix(i, 7) = Space(i - 19 + 2)
                .TextMatrix(i, 8) = Space(i - 19 + 2)
            Next
        End If
        For r = 22 To 26
            .TextMatrix(r, 1) = "���߽���"
            .TextMatrix(r, 2) = "���߽���"
        Next
        .Cell(flexcpAlignment, 23, 3, .Rows - 1, 4) = 1
        .Cell(flexcpAlignment, 23, 6, .Rows - 1, 7) = 1
        .Cell(flexcpForeColor, 23, 3, .Rows - 1, .Cols - 1) = vbBlue
        r = 22
        .Cell(flexcpBackColor, r, 3, r, .Cols - 1) = &H8000000F
        .TextMatrix(22, 3) = "��������"
        .TextMatrix(22, 4) = "��������"
        .TextMatrix(22, 5) = "��������"
        
        .TextMatrix(22, 6) = "��������"
        .TextMatrix(22, 7) = "��������"
        .TextMatrix(22, 8) = "��������"
        r = 23
        i = 0
        Do While Not mrsBacterin.EOF
            If r > .Rows - 1 Then Exit Do
            If i = 0 Then
                .TextMatrix(r, 3) = NVL(mrsBacterin!��������) & Space(r - 19 + 1)
                .TextMatrix(r, 4) = .TextMatrix(r, 3)
                .TextMatrix(r, 5) = NVL(mrsBacterin!����ʱ��) & Space(r - 19 + 1)
            Else
                .TextMatrix(r, 6) = NVL(mrsBacterin!��������) & Space(r - 19 + 1)
                .TextMatrix(r, 7) = .TextMatrix(r, 6)
                .TextMatrix(r, 8) = NVL(mrsBacterin!����ʱ��) & Space(r - 19 + 1)
            End If
            If i Mod 2 <> 0 Then
                r = r + 1
                i = 0
            Else
                i = 1
            End If
            mrsBacterin.MoveNext
        Loop
        
        For i = r To .Rows - 1
            If Trim(.TextMatrix(i, 3)) = "" Then
                .TextMatrix(i, 3) = Space(i - 19 + 1)
                .TextMatrix(i, 4) = Space(i - 19 + 1)
            End If
            If Trim(.TextMatrix(i, 6)) = "" Then
                .TextMatrix(i, 6) = Space(i - 19 + 1)
                .TextMatrix(i, 7) = Space(i - 19 + 1)
            End If
        Next
        For i = 0 To .Rows - 1
            .MergeRow(i) = True
        Next
        For i = 0 To .Cols - 1
            .MergeCol(i) = True
        Next
        .WordWrap = True
    End With
    

    Exit Sub
errHandle:
    If mobjDataBase.ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If mblnUnLoad Then Unload Me: Exit Sub
End Sub

Private Sub Form_Load()
    mblnUnLoad = Not LoadPatiInfor
    If mblnUnLoad Then Exit Sub
    
    Call InitGrid
    Call LoadPhoto
    Call InitTaskPancel
    Call picPatiInfor_Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Err = 0: On Error Resume Next
    If Not mrsInfor Is Nothing Then Set mrsInfor = Nothing
    If Not mrsOtherCertificate Is Nothing Then Set mrsOtherCertificate = Nothing
    If Not mrsDrug Is Nothing Then Set mrsDrug = Nothing
    If Not mrsBacterin Is Nothing Then Set mrsBacterin = Nothing
    If Not mobjDataBase Is Nothing Then Set mobjDataBase = Nothing
    If Not mcnOracle Is Nothing Then Set mcnOracle = Nothing
End Sub

Private Sub picPatiInfor_Resize()
    Err = 0: On Error Resume Next
    With picPatiInfor
        vsGrid.Left = .ScaleLeft
        vsGrid.Top = .ScaleTop
        vsGrid.Width = .ScaleWidth + 15
        vsGrid.Height = .ScaleHeight
    End With
End Sub
Private Sub LoadPhoto()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������Ƭ
    '����:���˺�
    '����:2012-12-14 16:01:43
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strTempFile As String
    Dim objTemp As clsDataBase
    Dim objDatabase As Object
    
     
    '��ʾ��Ƭ
    picPhoto.Cls
    strTempFile = mobjDataBase.ReadLob(glngSys, 27, mlng����ID)
    imgPhoto.Picture = LoadPicture(strTempFile)
    'ɾ������ʱ�ļ�
    Kill strTempFile
    imgPhoto.Left = picPhoto.ScaleLeft
    imgPhoto.Top = picPhoto.ScaleTop
    imgPhoto.Width = picPhoto.ScaleWidth
    imgPhoto.Height = picPhoto.ScaleHeight
    Set objDatabase = Nothing
End Sub
