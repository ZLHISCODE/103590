VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frm���ķ��Ź��� 
   Caption         =   "���ķ��Ź���"
   ClientHeight    =   7560
   ClientLeft      =   60
   ClientTop       =   525
   ClientWidth     =   11400
   Icon            =   "frm���ķ��Ź���.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7560
   ScaleWidth      =   11400
   ShowInTaskbar   =   0   'False
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshBody 
      Height          =   1860
      Left            =   165
      TabIndex        =   8
      Top             =   5220
      Width           =   9675
      _ExtentX        =   17066
      _ExtentY        =   3281
      _Version        =   393216
      FixedCols       =   0
      RowHeightMin    =   250
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      MouseIcon       =   "frm���ķ��Ź���.frx":014A
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.PictureBox PicLine_S 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   45
      Left            =   45
      MousePointer    =   7  'Size N S
      ScaleHeight     =   45
      ScaleWidth      =   4815
      TabIndex        =   6
      Top             =   3000
      Width           =   4815
   End
   Begin TabDlg.SSTab tabShow 
      Height          =   360
      Left            =   15
      TabIndex        =   2
      Top             =   615
      Width           =   10740
      _ExtentX        =   18944
      _ExtentY        =   635
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "�������嵥(&1)"
      TabPicture(0)   =   "frm���ķ��Ź���.frx":0464
      Tab(0).ControlEnabled=   0   'False
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "�������嵥(&2)"
      TabPicture(1)   =   "frm���ķ��Ź���.frx":0480
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Chk�嵥"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.CheckBox Chk�嵥 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Caption         =   "��ʾ���й��̵���"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4710
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   75
         Width           =   3450
      End
   End
   Begin VB.Timer TimeRefresh 
      Enabled         =   0   'False
      Left            =   4200
      Top             =   240
   End
   Begin VB.Timer TimePrintCancelBill 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   4905
      Top             =   285
   End
   Begin MSComctlLib.ImageList ImgTbarBlack 
      Left            =   8460
      Top             =   30
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImgTbarColor 
      Left            =   7890
      Top             =   30
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin ComCtl3.CoolBar cbrThis 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11400
      _ExtentX        =   20108
      _ExtentY        =   1164
      BandCount       =   2
      _CBWidth        =   11400
      _CBHeight       =   660
      _Version        =   "6.7.9782"
      Child1          =   "tlbThis"
      MinHeight1      =   600
      Width1          =   4995
      NewRow1         =   0   'False
      Caption2        =   "���ϲ���"
      Child2          =   "cboStock"
      MinHeight2      =   300
      Width2          =   3000
      NewRow2         =   0   'False
      Begin VB.ComboBox cboStock 
         Height          =   300
         Left            =   5970
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   180
         Width           =   5340
      End
      Begin MSComctlLib.Toolbar tlbThis 
         Height          =   600
         Left            =   165
         TabIndex        =   1
         Top             =   30
         Width           =   4800
         _ExtentX        =   8467
         _ExtentY        =   1058
         ButtonWidth     =   820
         ButtonHeight    =   1058
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   10
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��ӡ"
               Key             =   "��ӡ"
               Object.ToolTipText     =   "��ӡ"
               Object.Tag             =   "��ӡ"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Ԥ��"
               Key             =   "Ԥ��"
               Object.ToolTipText     =   "Ԥ��"
               Object.Tag             =   "Ԥ��"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "����"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "����"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "EditSp"
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "����"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "����"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�˳�"
               Key             =   "�˳�"
               Object.ToolTipText     =   "�˳�"
               Object.Tag             =   "�˳�"
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   7185
      Width           =   11400
      _ExtentX        =   20108
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frm���ķ��Ź���.frx":049C
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15028
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "��д"
            TextSave        =   "��д"
            Key             =   "STACAPS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshHead 
      Height          =   3405
      Left            =   45
      TabIndex        =   7
      Top             =   1050
      Width           =   9675
      _ExtentX        =   17066
      _ExtentY        =   6006
      _Version        =   393216
      FixedCols       =   0
      RowHeightMin    =   250
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      MouseIcon       =   "frm���ķ��Ź���.frx":0D2E
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSComctlLib.TabStrip tbsSel 
      Height          =   300
      Left            =   225
      TabIndex        =   4
      Top             =   4770
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   529
      Style           =   2
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "�ļ�(&F)"
      Begin VB.Menu mnuFileSet 
         Caption         =   "��ӡ����(&S)"
      End
      Begin VB.Menu mnuFilePreview 
         Caption         =   "��ӡԤ��(&V)"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "��ӡ(&P)"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileExcel 
         Caption         =   "�����&Excel"
      End
      Begin VB.Menu MnuFile1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileBillprint 
         Caption         =   "���ݴ�ӡ(&B)"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuFileRestore 
         Caption         =   "��ӡ����֪ͨ��(&R)"
      End
      Begin VB.Menu mnuFile2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePara 
         Caption         =   "��������(&A)"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuFile3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "�˳�(&X)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "�༭(&E)"
      Begin VB.Menu mnuEditPay 
         Caption         =   "����(&P)"
      End
      Begin VB.Menu mnuEditOutPay 
         Caption         =   "����(&O)"
      End
      Begin VB.Menu mnuEditSplit0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditPayCf 
         Caption         =   "����������(&F)"
      End
      Begin VB.Menu mnuEditFpPay 
         Caption         =   "��Ʊ�ݺŷ���(&R)"
      End
      Begin VB.Menu mnuEditStrict 
         Caption         =   "����������(&B)"
      End
      Begin VB.Menu mnuEditSp 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditStop 
         Caption         =   "ֹͣ���ϱ��(&S)"
      End
      Begin VB.Menu mnuEditStopSp 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditPayType 
         Caption         =   "δ����ģʽ(&W)"
         Checked         =   -1  'True
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditBackType 
         Caption         =   "�ѷ�������ģʽ(&H)"
         Shortcut        =   ^H
      End
      Begin VB.Menu mnuEditSelSplit 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditSelAll 
         Caption         =   "ȫѡ(&Q)"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEditClear 
         Caption         =   "ȫ��(&C)"
         Shortcut        =   ^R
      End
   End
   Begin VB.Menu mnuReport 
      Caption         =   "����(&R)"
      Visible         =   0   'False
      Begin VB.Menu mnuReportItem 
         Caption         =   "-"
         Index           =   0
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "�鿴(&V)"
      Begin VB.Menu mnuViewTool 
         Caption         =   "������(&T)"
         Begin VB.Menu mnuViewToolButton 
            Caption         =   "��׼��ť(&S)"
            Checked         =   -1  'True
         End
         Begin VB.Menu sdfsdfsd 
            Caption         =   "-"
         End
         Begin VB.Menu mnuViewToolText 
            Caption         =   "�ı���ǩ(&T)"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuViewStatus 
         Caption         =   "״̬��(&S)"
         Checked         =   -1  'True
      End
      Begin VB.Menu MnuView1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewFont 
         Caption         =   "����(&O)"
         Begin VB.Menu mnuViewFontSize 
            Caption         =   "С����(&S)"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu mnuViewFontSize 
            Caption         =   "������(&M)"
            Index           =   1
         End
         Begin VB.Menu mnuViewFontSize 
            Caption         =   "������(&B)"
            Index           =   2
         End
      End
      Begin VB.Menu MnuView2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewFind 
         Caption         =   "����(&F)"
         Shortcut        =   ^F
      End
      Begin VB.Menu MnuView3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewRefresh 
         Caption         =   "ˢ��(&R)"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu MnuHelp 
      Caption         =   "����(&H)"
      Begin VB.Menu MnuHelpTitle 
         Caption         =   "��������(&H)"
         Shortcut        =   {F1}
      End
      Begin VB.Menu MnuHelpWeb 
         Caption         =   "Web�ϵ�����(&W)"
         Begin VB.Menu mnuHelpWebHome 
            Caption         =   "������ҳ(&H)"
         End
         Begin VB.Menu mnuHelpWebForum 
            Caption         =   "������̳(&F)"
         End
         Begin VB.Menu mnuHelpWebMail 
            Caption         =   "���ͷ���(&E)..."
         End
      End
      Begin VB.Menu MnuHelp1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuHelpAbout 
         Caption         =   "����(&A)..."
      End
   End
End
Attribute VB_Name = "frm���ķ��Ź���"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private msngOldY As Single          '����Y������
Private mblnFirst As Boolean        '��һ�μ���ϵͳ
Dim mintFont As Integer

Private Type HeadCol
    ��־    As Byte
    ����    As Byte
    ����    As Byte
    ����    As Byte
    �շ�    As Byte
    ������  As Byte
    NO      As Byte
    ����    As Byte
    ����    As Byte
    סԺ��  As Byte
    ���    As Byte
    ����    As Byte
    �ɲ���  As Byte
    ˵��    As Byte
    Cols As Byte
End Type
Private mHeadCol As HeadCol

Private Type BodyCol
    ����ID      As Byte
    ����ID      As Byte
    ״̬        As Byte
    ����        As Byte
    ���÷���    As Byte
    ����        As Byte
    ����ҽ��    As Byte
    ����        As Byte
    NO          As Byte
    ����        As Byte
    ����        As Byte
    סԺ��      As Byte
    ��������    As Byte
    ���        As Byte
    ����        As Byte
    ��λ        As Byte
    ����ϵ��    As Byte
    ����        As Byte
    ����        As Byte
    ԭʼ����    As Byte
    ������      As Byte
    ׼����      As Byte
    ������      As Byte
    ����        As Byte
    ���        As Byte
    �����      As Byte
    ����Ա      As Byte
    ������      As Byte
    ����ʱ��    As Byte
    �ɲ���      As Byte
    ��¼״̬    As Byte
    Cols        As Byte
End Type

Private mBodyCol As BodyCol
Private mstrSelCon     As String   'ѡ��ĵ���:��ʽ�� No:����:��¼״̬||No:����:��¼״̬
Private mintCheckStock  As Integer  '����鷽ʽ

Private mstrStartDate As String
Private mstrEndDate As String
Private mlng����id As Long
Private mlng����id As Long
Private mstrסԺ�� As String
Private mstr�������� As String
Private mstrStartNo As String
Private mstrEndNo As String
Private mint���� As Long        '0-���ＰסԺ���е���,1-���ﻮ�ۼ��������,2-סԺ����
Private mintҵ������ As Long
Private mstr���� As String  '��In��ʽ
Private mlngCountSel As Long

'��ҩƷ������ҩ�������Ĳ���
Private mblnTrans As Boolean            'True��ʾ��ҩƷ������ҩ���ڵ���
Private mstrNo  As String               '���ݺţ������ڶ�λ
Private mlng�ⷿid As Long              '��ҩ�ⷿID��һ��ͷ��ϲ���һ��
Private mstrDrugStartDate As String     'ҩƷ���ݿ�ʼʱ��
Private mstrDrugEndDate As String       'ҩƷ���ݽ���ʱ��
Private mlngModule As Long
Private mintUnit As Integer        '0-ɢװ��λ,1-��װ��λ

'----------------------------------------------------------------------------------------------------------
'���˺�:����С��λ���ĸ�ʽ��
'�޸�:2007/03/06
Private mFMT As g_FmtString
Private mOraFMT As g_FmtString
'----------------------------------------------------------------------------------------------------------
Private mstrPreSelKey As String '�ϴ�ѡ���Key��

Private mintPrintPar As Integer    '0-��ʾ��ӡ,1-�Զ���ӡ,2-����ӡ
Private mblnExit As Boolean
Private mstrPrintCon As String '��ӡ����
Private mstrPrivs As String 'Ȩ�޴�

Public Sub ShowList(ByVal frmMain As Form, ByVal lng����id As Long, ByVal strNo As String, ByVal lng�ⷿID As Long, ByVal strStartDate As String, ByVal strEndDate As String)
    mlng����id = lng����id
    mstrNo = strNo
    mlng�ⷿid = lng�ⷿID
    mstrDrugStartDate = strStartDate
    mstrDrugEndDate = strEndDate
    mblnTrans = True
    
    Me.Show , frmMain
    Me.ZOrder 0

End Sub
Private Function CheckDepend() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:�������������
    '--�����:
    '--������:
    '--��  ��:
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New Recordset
    Dim lng���ϲ���ID As Long
    
    CheckDepend = False
    
    gstrSQL = "" & _
        "   SELECT DISTINCT a.id, a.���� " & _
        "   FROM ��������˵�� c, �������ʷ��� b, ���ű� a " & _
        "   Where c.�������� = b.���� " & _
        "           AND b.���� ='W' " & _
        "           AND a.id = c.����id " & _
        "           AND TO_CHAR (a.����ʱ��, 'yyyy-MM-dd') = '3000-01-01'" & _
        IIf(InStr(gstrPrivs, "���в���") <> 0, "", " And a.ID IN (Select ����ID From ������Ա Where ��ԱID=[1])")
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��Ӧ�Ŀⷿ", UserInfo.Id)
    
    If rsTemp.EOF Then
        ShowMsgBox "����Ӧ������һ�����з��ϲ������ʻ�����" & vbCrLf & "���Ƿ��ϲ��ŵĹ�����Ա,��鿴���Ź���"
        rsTemp.Close
        Exit Function
    End If
    
    '�����ҩƷ���ڴ��룬���÷��ϲ�����ҩƷ��ҩ����һ��
    If mblnTrans Then
        If mlng�ⷿid <> UserInfo.����ID Then
            lng���ϲ���ID = mlng�ⷿid
        Else
            lng���ϲ���ID = UserInfo.����ID
        End If
    End If
    
    'װ�뷢�ϲ�������
    With cboStock
        .Clear
        Do While Not rsTemp.EOF
            .AddItem rsTemp!����
            .ItemData(.NewIndex) = rsTemp!Id
            If rsTemp!Id = lng���ϲ���ID Then
                .ListIndex = .NewIndex
            End If
            rsTemp.MoveNext
        Loop
        If .ListIndex = -1 Then .ListIndex = 0
        rsTemp.Close
    End With
    CheckDepend = True
End Function

Private Function InitSet()
    Dim i As Long
    Dim rsTemp As New ADODB.Recordset
    
    '����ʱ�䣬����Ǵ�ҩƷ���ڴ��룬����ҩƷ��ҩʱ��һ��
    If mblnTrans Then
        mstrEndDate = mstrDrugEndDate
        mstrStartDate = mstrDrugStartDate
    Else
        mstrEndDate = Format(zlDatabase.Currentdate, "yyyy-mm-dd") & " 23:59:59"
        mstrStartDate = Format(DateAdd("d", -7, zlDatabase.Currentdate), "yyyy-mm-dd") & " 00:00:00"
    End If
    
    Call LoadInIcon
    
    Call Ȩ�޿���
       
    '�ָ�����
    Dim strReg As String
    strReg = zlDatabase.GetPara("�����ֺ�", glngSys, mlngModule, "0")
    mnuViewFontSize_Click Val(strReg)
    
    strReg = Val(zlDatabase.GetPara("���ĵ�λ", glngSys, mlngModule, "0"))
    mintUnit = Val(strReg)
  
    '���˺�:����С����ʽ����
    With mFMT
        .FM_�ɱ��� = GetFmtString(mintUnit, g_�ɱ���)
        .FM_��� = GetFmtString(mintUnit, g_���)
        .FM_���ۼ� = GetFmtString(mintUnit, g_�ۼ�)
        .FM_���� = GetFmtString(mintUnit, g_����)
    End With
    With mOraFMT
        .FM_�ɱ��� = GetFmtString(mintUnit, g_�ɱ���, True)
        .FM_��� = GetFmtString(mintUnit, g_���, True)
        .FM_���ۼ� = GetFmtString(mintUnit, g_�ۼ�, True)
        .FM_���� = GetFmtString(mintUnit, g_����, True)
    End With
    
    mintPrintPar = Val(zlDatabase.GetPara("���ϴ�ӡ���ѷ�ʽ", glngSys, mlngModule, "0"))
    strReg = Trim(zlDatabase.GetPara("��ѯҵ������", glngSys, mlngModule, ""))
    If strReg = "" Then strReg = "24,25,26"
    mstr���� = strReg
    
    With tbsSel
        .Tabs.Clear
        .Tabs.Add , "K1", "������ϸ"
        .Tabs.Add , "K2", "��ϸ���"
        .Tabs.Add , "K3", "�������"
    End With
        
    '������ȷ��
    With mHeadCol
        .��־ = 0
        .���� = 1
        .���� = 2
        .���� = 3
        .�շ� = 4
        .������ = 5
        .NO = 6
        .���� = 7
        .���� = 8
        .סԺ�� = 9
        .��� = 10
        .���� = 11
        .�ɲ��� = 12
        .˵�� = 13
        
        .Cols = 14
    End With
    With mBodyCol
       i = 0: .����ID = i
       i = i + 1: .����ID = i
       i = i + 1: .���� = i
       i = i + 1: .���÷��� = i
       i = i + 1: .״̬ = i
       i = i + 1: .���� = i
       i = i + 1: .����ҽ�� = i
       i = i + 1: .���� = i
       i = i + 1: .NO = i
       i = i + 1: .���� = i
       i = i + 1: .���� = i
       i = i + 1: .סԺ�� = i
       i = i + 1: .�������� = i
       i = i + 1: .��� = i
       i = i + 1: .���� = i
       i = i + 1: .��λ = i
       i = i + 1: .����ϵ�� = i
       i = i + 1: .���� = i
       i = i + 1: .���� = i
       i = i + 1: .ԭʼ���� = i
       i = i + 1: .������ = i
       i = i + 1: .׼���� = i
       i = i + 1: .������ = i
       i = i + 1: .���� = i
       i = i + 1: .��� = i
       i = i + 1: .����� = i
       i = i + 1: .����Ա = i
       i = i + 1: .������ = i
       i = i + 1: .�ɲ��� = i
       i = i + 1: .��¼״̬ = i
       i = i + 1: .����ʱ�� = i
       
       i = i + 1: .Cols = i
    End With
    
End Function

Private Function ReadSystemPara()
    Dim rsTemp As New ADODB.Recordset
    gstrSQL = " Select Nvl(��鷽ʽ,0) ����� From ���ϳ����� Where �ⷿID=[1]"

    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, cboStock.ItemData(cboStock.ListIndex))
    
    With rsTemp
        If Not .EOF Then
            mintCheckStock = NVL(!�����, 0)
        End If
    End With
    
End Function

Private Sub cboStock_Click()
        Call ReadSystemPara
        tabShow_Click -1
        SetMnuEnable
End Sub

Private Sub cbrThis_Resize()
    Form_Resize
End Sub

Private Sub Chk�嵥_Click()
    tabShow_Click 0
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    '��ʼ��ͷ
    Call SetGrdColHead(1)
    Call SetGrdColHead(2)
    
    '�������������ϵ
    If CheckDepend = False Then Unload Me: Exit Sub
    
    '����ؼ���ʼ
    ' 1.Ȩ�޿���
    Call Ȩ�޿���
    ' 2.�����ͷ��ʼ
    
    
    'װ���������
    Call tabShow_Click(tabShow.Tab + 1)
 
        
End Sub

Private Sub Form_Load()
    mblnFirst = True
    mlngModule = glngModul
        
    mstrPrivs = gstrPrivs
    '��ʼ��ؿؼ�
    Call InitSet
    
    '�ָ����Ի�����
    Call RestoreWinState(Me)
    '2006-04-25:���˺�,ͳһ���ӱ�������ģ��Ĺ���
    Call zlDatabase.ShowReportMenu(Me, glngSys, glngModul, gstrPrivs)
     
End Sub

Private Sub Form_Resize()
    '����λ������
    Dim sngCbrHeight As Single
    Dim sngStbHeight As Single
    
    
    On Error Resume Next
    sngCbrHeight = IIf(cbrThis.Visible, cbrThis.Height, 0)
    sngStbHeight = IIf(stbThis.Visible, stbThis.Height, 0)
    
    If Me.WindowState = vbMinimized Then Exit Sub
    cbrThis.Bands(1).MinHeight = tlbThis.Height
    If Me.Height < PicLine_S.Height + PicLine_S.Top + tbsSel.Height + 650 + sngStbHeight Then
        Me.Height = PicLine_S.Height + PicLine_S.Top + tbsSel.Height + 650 + sngStbHeight
    End If
    
    With tabShow
        .Left = 0
        .Width = ScaleWidth
        .Top = sngCbrHeight + 10
    End With
    
    With mshHead
        .Left = 0
        .Top = tabShow.Height + tabShow.Top + 10
        .Height = PicLine_S.Top - .Top
        .Width = ScaleWidth - 10
    End With
    
    With PicLine_S
        .Left = 0
        .Width = mshHead.Width
    End With
    
    With tbsSel
        .Left = 0
        .Top = PicLine_S.Top + PicLine_S.Height + 10
        .Width = mshHead.Width
    End With
    
    With mshBody
        .Left = 0
        If tabShow.Tab = 1 Then
            .Top = mshHead.Top
        Else
            .Top = tbsSel.Top + tbsSel.Height + 10
        End If
        .Height = ScaleHeight - .Top - sngStbHeight
        .Width = mshHead.Width
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)

    '������Ի�����
    Call SaveWinState(Me)
End Sub
Private Function Save����(ByVal strDate As String) As Boolean
    '������ϵ��������
    Dim strNo As String
    Dim lngID As Long
    Dim lngRow As Long
    Dim ����ID As Long
    Dim int�Զ����� As Integer
    Dim strReg As String
    int�Զ����� = IIf(Val(zlDatabase.GetPara("�Զ�����", glngSys, mlngModule)) = 1, 1, 0)
    
    Save���� = False
    err = 0
    On Error GoTo ErrHand:
    
    gcnOracle.BeginTrans
    
    With mshBody
        For lngRow = 1 To .Rows - 1
        
                strNo = Trim(.TextMatrix(lngRow, mBodyCol.NO))
                '���̲���:ID_IN,�����_IN,�������_IN,����_IN,Ч��_IN,����_IN,��������_IN,�Զ�����_IN(1-�Զ�����,0-���Զ�����)
                If strNo <> "" And Trim(.TextMatrix(lngRow, mBodyCol.״̬)) = "��" Then
                   gstrSQL = "zl_�����շ���¼_��������("
                   gstrSQL = gstrSQL & .RowData(lngRow) & ","
                   gstrSQL = gstrSQL & "'" & gstrUserName & "',"
                   gstrSQL = gstrSQL & "to_date('" & strDate & "','yyyy-mm-dd HH24:mi:ss'),"
                   gstrSQL = gstrSQL & "'" & Replace(.TextMatrix(lngRow, mBodyCol.����), "(" & .TextMatrix(lngRow, mBodyCol.����) & ")", "") & "',"
                   gstrSQL = gstrSQL & "NULL" & ","
                   gstrSQL = gstrSQL & "NULL" & ","
                  ' If mintUnit = 0 Then
                        gstrSQL = gstrSQL & Val(.TextMatrix(lngRow, mBodyCol.ԭʼ����))
                   'Else
                 '       gstrSQL = gstrSQL & Val(.TextMatrix(lngRow, mBodyCol.����)) * Val(.TextMatrix(lngRow, mBodyCol.����ϵ��))
                  ' End If
                   gstrSQL = gstrSQL & "," & int�Զ����� & ")"
                   Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
                End If
        Next
    End With
    gcnOracle.CommitTrans
    Save���� = True
    Exit Function
ErrHand:
    gcnOracle.RollbackTrans
    Call ErrCenter
    Call SaveErrLog
End Function

Private Sub mnuEditBackType_Click()
    Dim blnYes As Boolean
    If mstrSelCon <> "" Then
        ShowMsgBox "�Ѿ��б�ѡ�����Ŀ,���Ƿ�ϣ���ı�ѡ��!", True, blnYes
        If Not blnYes Then
            mblnExit = True
            mnuEditPayType.Checked = True
            mnuEditBackType.Checked = False
            Exit Sub
        End If
        mstrSelCon = ""
    End If
    
    tabShow.Tab = 1
    mnuEditPayType.Checked = False
    mnuEditBackType.Checked = True
    tabShow_Click 0
    
End Sub

Private Sub mnuEditClear_Click()
     If tabShow.Tab = 0 Then
        Call SelAndClearAllPlay(False)
     Else
        Call SelAndClearAllOutPlay(False)
     End If
     Call SetMnuEnable
End Sub

Private Sub mnuEditFpPay_Click()
    '��Ʊ�ݺŷ���
    With Frm�����ŷ���
        .In_���� = mintҵ������
        .In_����IN = mstr����
        .In_���ϲ���id = cboStock.ItemData(cboStock.ListIndex)
        .In_����� = mintCheckStock
        .In_����δ���Ϸ��� = 1
        .��Ʊ�ݺŷ��� = True
        .In_Ȩ�� = mstrPrivs
        .mstr������ = gstrUserName
        .Show 1, Me
    End With
    mnuViewRefresh_Click
    
End Sub

Private Sub mnuEditOutPay_Click()
    Dim strDate As String
    Dim blnYe As Boolean
    
    ShowMsgBox "���Ƿ����Ҫ����Щ��¼����������?", True, blnYe
    If blnYe = False Then
        Exit Sub
    End If
    
    strDate = Format(zlDatabase.Currentdate(), "yyyy-MM-dd HH:MM:SS")

    If Save����(strDate) = False Then Exit Sub
    BillListPrint 0, strDate, 2
    mstrSelCon = ""
    mlngCountSel = 0
    mnuViewRefresh_Click
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuEditPay_Click()
    If CheckStock = False Then Exit Sub
    If SendBill = False Then Exit Sub
    mstrSelCon = ""
    mstrPrintCon = ""
    '����ˢ��
    tabShow_Click 1
End Sub

Private Sub mnuEditPayCf_Click()

    With Frm�����ŷ���
        .In_���� = mintҵ������
        .In_����IN = mstr����
        '.In_���ϴ��� = str����
        .In_���ϲ���id = cboStock.ItemData(cboStock.ListIndex)
        .In_����� = mintCheckStock
        '.In_У�鴦�� = intVerify
        .In_����δ���Ϸ��� = 1
        '.IN_����δ��˷�ҩ = Int����δ��˴�����ҩ
        .In_Ȩ�� = mstrPrivs
        .mstr������ = gstrUserName
        .��Ʊ�ݺŷ��� = False
        .Show 1, Me
    End With
    mnuViewRefresh_Click
End Sub

Private Sub mnuEditPayType_Click()
'    Dim blnYes As Boolean
'    ShowMsgbox "�Ѿ��б�ѡ�����Ŀ,���Ƿ�ϣ���ı�ѡ��!", True, blnYes
'    If Not blnYes Then
'        mnuEditPayType.Checked = False
'        mnuEditBackType.Checked = True
'        Exit Sub
'    End If
    mlngCountSel = 0
    tabShow.Tab = 0
    mnuEditPayType.Checked = True
    mnuEditBackType.Checked = False
    tabShow_Click 1
End Sub

Private Sub mnuEditSelAll_Click()
     If tabShow.Tab = 0 Then
        Call SelAndClearAllPlay(True)
     Else
        Call SelAndClearAllOutPlay(True)
     End If
     Call SetMnuEnable
End Sub

Private Sub mnuEditStop_Click()
    'ֹͣ����
    '��ҩ��ʽ=-1
        
    Dim frmFlag As New Frm���ٷ�ҩ������־
    frmFlag.gstrParentName = Me.Name
    frmFlag.Show vbModal
    mnuViewRefresh_Click
        
End Sub

Private Sub mnuEditStrict_Click()
    '
    If Frm����������.ShowCard(Me, cboStock.ItemData(cboStock.ListIndex), mstrPrivs) = False Then Exit Sub
    Call mnuViewRefresh_Click
End Sub

Private Sub mnuFileBillPrint_Click()
        '���ݴ�ӡ:
        Dim lng���� As Long
        Dim strNo As String
        Dim strDate As String
        Dim int���Ϸ�ʽ As Integer
        Dim rsTemp As New ADODB.Recordset
        If tabShow.Tab = 0 Then Exit Sub
        
        With mshBody
            
            strDate = .TextMatrix(.Row, mBodyCol.����ʱ��)
            lng���� = Decode(.TextMatrix(.Row, mBodyCol.����), "�շ�", 24, "���ʵ�", 25, "���ʱ�", 26, 0)
            strNo = Trim(.TextMatrix(.Row, mBodyCol.NO))
            mstrPrintCon = strNo & "||" & lng���� & "||" & cboStock.ItemData(cboStock.ListIndex)
            If strNo = "" Then Exit Sub
            
            gstrSQL = "Select ��ҩ��ʽ from ҩƷ�շ���¼ where id=[1]"
            gstrSQL = AnalyseHistorySQL(gstrSQL)
            
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, .RowData(.Row))

            If rsTemp.EOF Then
                Exit Sub
            End If
            int���Ϸ�ʽ = NVL(rsTemp!��ҩ��ʽ, 1)
            If int���Ϸ�ʽ <> 1 Then
                mstrPrintCon = ""
            End If
        End With
        BillListPrint int���Ϸ�ʽ, strDate, 0
End Sub

Private Sub mnuFileExcel_Click()
    '�����EXCEL
    subPrint 3
End Sub

Private Sub mnufileexit_Click()
    Unload Me
End Sub



Private Sub mnuFilePara_Click()
    Dim strReg As String
    
    If frmPayExitParaSet.ShowSetPara(Me, mlngModule, mstrPrivs) = False Then Exit Sub
    strReg = Val(zlDatabase.GetPara("���ĵ�λ", glngSys, mlngModule, "0"))
    mintUnit = Val(strReg)
    
    mintPrintPar = Val(zlDatabase.GetPara("���ϴ�ӡ���ѷ�ʽ", glngSys, mlngModule, "0"))
    strReg = Trim(zlDatabase.GetPara("��ѯҵ������", glngSys, mlngModule, ""))
    If strReg = "" Then strReg = "24,25,26"
    mstr���� = strReg
    '���»�ȡ����
    mnuViewRefresh_Click
End Sub
Private Sub mnuFilePreView_Click()
    '��ӡԤ��
    subPrint 2
End Sub

Private Sub mnuFilePrint_Click()
    '��ӡ
    subPrint 1
End Sub
Private Sub mnuFileRestore_Click()
        '
        If tabShow.Tab <> 1 Then Exit Sub
        
        Dim strDate As String
        strDate = mshBody.TextMatrix(mshBody.Row, mBodyCol.����ʱ��)
        BillListPrint , strDate, 2
End Sub

Private Sub mnuFileSet_Click()
'��ӡ����
    zlPrintSet
End Sub

Private Sub mnuViewFind_Click()
    Dim strStartDate As String, strEndDate As String
    Dim strStartNo As String, strEndNo As String
    Dim str���� As String, intҵ������ As Integer
    Dim strסԺ�� As String, lng����id As Long, lng����id As Long
    Dim str���� As String
    Dim blnreturn As Boolean
    
    blnreturn = frm���ķ��Ź���Search.ShowEdit( _
        Me, strStartDate, strEndDate, strStartNo, strEndNo, str����, _
        intҵ������, strסԺ��, lng����id, str����, lng����id)
    If blnreturn = False Then Exit Sub
    
    mstrStartDate = strStartDate
    mstrEndDate = strEndDate
    mstrStartNo = strStartNo
    mstrEndNo = strEndNo
    mint���� = intҵ������
    mintҵ������ = intҵ������
    mstrסԺ�� = strסԺ��
    mlng����id = lng����id
    mstr�������� = str����
    mlng����id = lng����id
    
    mnuViewRefresh_Click
End Sub
 

Private Sub mnuViewFontSize_Click(Index As Integer)
    Dim i As Integer
    For i = 0 To 2
        Me.mnuViewFontSize(i).Checked = False
    Next
    Me.mnuViewFontSize(Index).Checked = True

    Select Case Index
    Case 0
        Me.mshHead.Font.Size = 9
        Me.tabShow.Font.Size = 9
        mshBody.Font.Size = 9
        tbsSel.Font.Size = 9
     Case 1
        Me.mshHead.Font.Size = 11
        Me.tabShow.Font.Size = 11
        mshBody.Font.Size = 11
        tbsSel.Font.Size = 11
    Case 2
        Me.mshHead.Font.Size = 15
        Me.tabShow.Font.Size = 15
        mshBody.Font.Size = 15
        tbsSel.Font.Size = 15
    End Select
    mintFont = Index
    Call zlDatabase.SetPara("�����ֺ�", mintFont, glngSys, mlngModule)
    Form_Resize
    Me.Refresh
    
End Sub



Private Sub mnuViewRefresh_Click()
    mstrSelCon = ""
    Select Case tabShow.Tab
        Case 0 '--δ�����嵥
            SetMnuEnable
            Form_Resize
            Call GetHeadData(0)
            If Me.mshHead.Enabled Then mshHead.SetFocus
            mshHead_EnterCell
        Case 1  '--�ѷ����嵥
            SetMnuEnable
            Form_Resize
            Call ReadBillData
            If Me.mshBody.Enabled Then mshBody.SetFocus
    End Select
End Sub


Private Sub mnuViewStatus_Click()
    With mnuViewStatus
        .Checked = Not .Checked
        stbThis.Visible = .Checked
    End With
    
    Form_Resize
End Sub
Private Sub mnuReportItem_Click(Index As Integer)
    Dim strNo As String
    Dim intRecodeSta As Integer
    Dim lng���ϲ���ID As Long
    Dim lngCol As Long
    Dim lng����ID As Long
    Dim lng����ID As Long
    Dim strסԺ�� As String
    Dim lng���� As Long
    
    If cboStock.ListIndex < 0 Then
        lng���ϲ���ID = 0
    Else
        lng���ϲ���ID = cboStock.ItemData(cboStock.ListIndex)
    End If
        
    
    Select Case tabShow.Tab
    Case 0
        With mshHead
                lng���� = Decode(.TextMatrix(.Row, mHeadCol.����), "�շ�", 24, "���ʵ�", 25, "���ʱ�", 26, 0)
                strNo = Trim(.TextMatrix(.Row, mHeadCol.NO))
                lng����ID = Val(mshBody.TextMatrix(mshBody.Row, mBodyCol.����ID))
                lng����ID = Val(mshBody.TextMatrix(mshBody.Row, mBodyCol.����ID))
                intRecodeSta = Val(mshBody.TextMatrix(mshBody.Row, mBodyCol.��¼״̬))
                strסԺ�� = Trim(mshBody.TextMatrix(mshBody.Row, mBodyCol.סԺ��))
        End With
    Case Else
        With mshBody
                strNo = Trim(.TextMatrix(.Row, mBodyCol.NO))
                lng���� = Decode(.TextMatrix(.Row, mBodyCol.����), "�շ�", 24, "���ʵ�", 25, "���ʱ�", 26, 0)
                lng����ID = Val(.TextMatrix(.Row, mBodyCol.����ID))
                lng����ID = Val(.TextMatrix(.Row, mBodyCol.����ID))
                intRecodeSta = Val(.TextMatrix(.Row, mBodyCol.��¼״̬))
                strסԺ�� = Trim(.TextMatrix(.Row, mBodyCol.סԺ��))
        End With
    End Select
   
    '2006-04-25:���˺�:�����Զ��屨������ģ��Ĺ���
    Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, "NO=" & strNo, "��¼״̬=" & intRecodeSta, "���ϲ���=" & lng���ϲ���ID, "��������=" & lng����, "����=" & lng����ID, "����=" & lng����ID, "סԺ��=" & strסԺ��)
End Sub
Private Sub mnuViewToolButton_Click()
    With mnuViewToolButton
        .Checked = Not .Checked
        cbrThis.Bands(1).Visible = .Checked
        mnuViewToolText.Enabled = .Checked
    End With
    Form_Resize
End Sub

Private Sub mnuViewToolText_Click()
    Dim intCount As Integer      '����������
    
    mnuViewToolText.Checked = Not mnuViewToolText.Checked
    With tlbThis.Buttons
        If mnuViewToolText.Checked = False Then
            'ȡ�����е��ı���ǩ��ʾ
            For intCount = 1 To .Count
                .Item(intCount).Caption = ""
            Next
        Else
            '�����е��ı���ǩ��ʾ��˵����Tag�зŵ��ı���ǩ
            For intCount = 1 To .Count
                .Item(intCount).Caption = .Item(intCount).Tag
            Next
        End If
    End With
    
    cbrThis.Bands(1).MinHeight = tlbThis.Height
    
    Form_Resize
End Sub

Private Sub mshBody_DblClick()
    '
    Dim strNo As String
    Dim str���� As String
    Dim int���� As Integer
    
    If mnuEditOutPay.Visible = False Then Exit Sub
      
    If tabShow.Tab <> 1 Then
        mlngCountSel = 0
        Exit Sub
    End If
    With mshBody
        strNo = Trim(.TextMatrix(.Row, mBodyCol.NO))
        
        If strNo = "" Then Exit Sub
        int���� = Val(.TextMatrix(.Row, mBodyCol.�ɲ���))
        If int���� <> 1 Then
            If int���� = -99 Then
                ShowMsgBox "�ü�¼�Ѿ���ת����ʷ����,���ܽ������ϴ���!"
            End If
            Exit Sub
        End If
        
        str���� = Decode(Trim(.TextMatrix(.Row, mHeadCol.����)), "�շ�", 24, "���ʵ�", 25, 26)
        
        If Trim(.TextMatrix(.Row, mBodyCol.״̬)) <> "��" Then
            .TextMatrix(.Row, mBodyCol.״̬) = "��"
            mlngCountSel = mlngCountSel + 1

        Else
            .TextMatrix(.Row, mBodyCol.״̬) = ""
            mlngCountSel = mlngCountSel - 1
        End If
        If mlngCountSel < 0 Then mlngCountSel = 0
    End With
    SetMnuEnable
End Sub

Private Sub mshBody_EnterCell()
    
    If tabShow.Tab <> 1 Then
        Exit Sub
    End If
    With mshBody
        .ForeColorSel = .CellForeColor
    End With
    SetMnuEnable
    
End Sub

Private Sub mshBody_GotFocus()
    SetGrdSelBackColor mshBody
End Sub

Private Sub mshBody_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 32 Then  '�ո�
        mshBody_DblClick
    End If
End Sub

Private Sub mshBody_LostFocus()
    With mshBody
        .ForeColorSel = .CellForeColor
    End With
End Sub

Private Sub mshBody_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 2 Then Exit Sub
    If tabShow.Tab <> 1 Then Exit Sub
    PopupMenu mnuEdit
    
End Sub

Private Sub mshHead_DblClick()
    Dim strNo As String
    Dim str���� As String
    If mnuEditPay.Visible = False Then Exit Sub
    
    With mshHead
        strNo = Trim(.TextMatrix(.Row, mHeadCol.NO))
        If strNo = "" Then Exit Sub
        str���� = Decode(Trim(.TextMatrix(.Row, mHeadCol.����)), "�շ�", 24, "���ʵ�", 25, 26)
        
        If Trim(.TextMatrix(.Row, mHeadCol.��־)) <> "��" Then
            .TextMatrix(.Row, mHeadCol.��־) = "��"
            '�����ݴ��м���ֵ
            mstrSelCon = mstrSelCon & "||" & strNo & ":" & str����
        Else
            mstrSelCon = Replace(mstrSelCon, "||" & strNo & ":" & str����, "")
            .TextMatrix(.Row, mHeadCol.��־) = ""
        End If
    End With
    If tbsSel.SelectedItem.Key <> "K1" Then
        '�����ݽ���ˢ��
        tbsSel_Click
    End If
    SetMnuEnable
End Sub
Private Function SelAndClearAllPlay(ByVal blnSel As Boolean) As Boolean
    'ѡ���������ϵ������Ϣ
    Dim strNo As String
    Dim str���� As String
    Dim i As Long
    If mnuEditPay.Visible = False Then Exit Function
    err = 0: On Error GoTo ErrHand:
    With mshHead
        mstrSelCon = ""
        For i = 1 To .Rows - 1
            strNo = Trim(.TextMatrix(i, mHeadCol.NO))
            If strNo = "" Then Exit For
            If blnSel Then
                str���� = Decode(Trim(.TextMatrix(i, mHeadCol.����)), "�շ�", 24, "���ʵ�", 25, 26)
                .TextMatrix(i, mHeadCol.��־) = "��"
                '�����ݴ��м���ֵ
                mstrSelCon = mstrSelCon & "||" & strNo & ":" & str����
            Else
                .TextMatrix(i, mHeadCol.��־) = ""
            End If
       Next
    End With
    SelAndClearAllPlay = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function
Private Function SelAndClearAllOutPlay(ByVal blnSel As Boolean) As Boolean
    '����:ѡ���������е��ѷ��ϲ���
    Dim strNo As String
    Dim str���� As String
    Dim int���� As Integer
    Dim i As Long
    If mnuEditOutPay.Visible = False Then Exit Function
    err = 0: On Error GoTo ErrHand:
    With mshBody
        mlngCountSel = 0
        For i = 1 To .Rows - 1
            strNo = Trim(.TextMatrix(i, mBodyCol.NO))
            If strNo = "" Then Exit For
            int���� = Val(.TextMatrix(i, mBodyCol.�ɲ���))
            If int���� = 1 And blnSel Then
                .TextMatrix(i, mBodyCol.״̬) = "��"
                mlngCountSel = mlngCountSel + 1
            Else
                .TextMatrix(i, mBodyCol.״̬) = ""
            End If
        Next
    End With
    SelAndClearAllOutPlay = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Private Sub SetMnuEnable()
    Dim blnHead As Boolean
    Dim blnBack As Boolean  '����ģʽ
    Dim blnData As Boolean '�Ƿ��������
    Dim blnSelData As Boolean '���ڱ�ѡ�������
    Dim blnPrint As Boolean
    Dim blnSelAndClear As Boolean
    blnBack = tabShow.Tab <> 0
           
    If blnBack Then
        blnSelData = mlngCountSel > 0
        blnHead = Trim(mshBody.TextMatrix(mshBody.Row, mBodyCol.NO)) <> ""
        blnData = Trim(mshBody.TextMatrix(1, mBodyCol.NO)) <> ""
        blnPrint = Val(mshBody.TextMatrix(mshBody.Row, mBodyCol.��¼״̬)) Mod 3 = 2
        blnSelAndClear = mnuEditOutPay.Visible And blnData
    Else
        blnSelData = mstrSelCon <> ""
        blnHead = Trim(mshHead.TextMatrix(mshHead.Row, mHeadCol.NO)) <> ""
        blnData = Trim(mshHead.TextMatrix(1, mHeadCol.NO)) <> ""
        blnSelAndClear = mnuEditPay.Visible And blnData
    End If
    
    Chk�嵥.Enabled = blnBack
    mnuEditClear.Enabled = blnSelAndClear
    mnuEditSelAll.Enabled = blnSelAndClear
    mnuEditPay.Enabled = blnHead And blnSelData And Not blnBack
    mnuEditOutPay.Enabled = blnHead And blnSelData And blnBack
    
    mnuFilePreview.Enabled = blnData
    mnuFilePrint.Enabled = blnData
    mnuFileExcel.Enabled = blnData
    
    mnuFileRestore.Enabled = blnBack And blnData And blnPrint
    mnuFileBillprint.Enabled = blnBack And blnData And Not blnPrint

    
    tlbThis.Buttons("��ӡ").Enabled = blnData
    tlbThis.Buttons("Ԥ��").Enabled = blnData
    
    tlbThis.Buttons("����").Enabled = mnuEditPay.Enabled
    tlbThis.Buttons("����").Enabled = mnuEditOutPay.Enabled
    mshHead.Visible = Not blnBack
    tbsSel.Visible = Not blnBack
End Sub

Private Sub mshHead_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 32 Then  '�ո�
        mshHead_DblClick
    End If
End Sub

Private Sub mshHead_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
        If Button <> 2 Then Exit Sub
        PopupMenu mnuEdit
End Sub

Private Sub tlbThis_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Ԥ��"
            mnuFilePreView_Click
        Case "��ӡ"
            mnuFilePrint_Click
        Case "����"
            mnuEditPay_Click
        Case "����"
            mnuEditOutPay_Click
        Case "����"
            mnuViewFind_Click
        Case "����"
            mnuHelpTitle_Click
        Case "�˳�"
            mnufileexit_Click
    End Select
End Sub

Private Sub mshHead_EnterCell()
    With mshHead
        If .TextMatrix(.Row, mHeadCol.NO) = "" Then SetMnuEnable: Exit Sub
        .ForeColorSel = .CellForeColor
        tbsSel_Click
        SetMnuEnable
    End With
End Sub

Private Sub mshHead_GotFocus()
    SetGrdSelBackColor mshHead
End Sub

Private Sub mshHead_LostFocus()
  With mshHead
        If .TextMatrix(.Row, mHeadCol.NO) = "" Then Exit Sub
        .ForeColorSel = .CellForeColor
    End With
End Sub

Private Sub PicLine_S_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 1 Then Exit Sub
    msngOldY = y
End Sub

Private Sub PicLine_S_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    '�ָ�������
    
    If Button <> 1 Then Exit Sub
    
    With PicLine_S
        If .Top + y < 2000 Then Exit Sub
        If .Top + y > ScaleHeight - 2000 Then Exit Sub
        .Move .Left, .Top + y - msngOldY
    End With
    
    With mshHead
        .Height = PicLine_S.Top - .Top
    End With
    With tbsSel
        .Top = PicLine_S.Top + PicLine_S.Height + 10
    End With
    With mshBody
        .Top = tbsSel.Top + tbsSel.Height + 10
        .Height = ScaleHeight - .Top - IIf(stbThis.Visible, stbThis.Height, 0)
        .Width = mshHead.Width
    End With
End Sub
Private Sub PicLine_S_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    msngOldY = 0
End Sub

Private Sub SetGrdColHead(ByVal IntStyle As Integer, Optional ByVal blnInitColHead As Boolean = True)
    Dim intCol As Integer
    '--���ø��б�ؼ��ĸ�ʽ--
    
    Select Case IntStyle
    Case 1
        With mshHead
             If blnInitColHead Then
                .Clear
                .Rows = 2
                .Cols = mHeadCol.Cols
                .TextMatrix(0, mHeadCol.��־) = "��־"
                .TextMatrix(0, mHeadCol.����) = "����"
                .TextMatrix(0, mHeadCol.����) = "����"
                .TextMatrix(0, mHeadCol.����) = "����"
                .TextMatrix(0, mHeadCol.�շ�) = "�շ�"
                .TextMatrix(0, mHeadCol.������) = "������"
                .TextMatrix(0, mHeadCol.NO) = "NO"
                .TextMatrix(0, mHeadCol.����) = "����"
                .TextMatrix(0, mHeadCol.����) = "����"
                .TextMatrix(0, mHeadCol.סԺ��) = "סԺ��"
                .TextMatrix(0, mHeadCol.���) = "���"
                .TextMatrix(0, mHeadCol.����) = "����"
                .TextMatrix(0, mHeadCol.�ɲ���) = "�ɲ���"
                .TextMatrix(0, mHeadCol.˵��) = "˵��"
            End If
            
            For intCol = 0 To .Cols - 1
                .ColAlignmentFixed(intCol) = 4
            Next
            .ColWidth(mHeadCol.����) = 0
            .ColWidth(mHeadCol.�շ�) = 0
            .ColWidth(mHeadCol.������) = 0
            
            If RestoreFlexState(Me.mshHead, Me.Caption) = False Then
                .ColWidth(mHeadCol.����) = 600
                .ColWidth(mHeadCol.��־) = 500
                .ColWidth(mHeadCol.����) = 1500
                .ColWidth(mHeadCol.NO) = 800
                .ColWidth(mHeadCol.����) = 800
                .ColWidth(mHeadCol.���) = 1000
                .ColWidth(mHeadCol.����) = 1500
                .ColWidth(mHeadCol.�ɲ���) = 0
                .ColWidth(mHeadCol.˵��) = 1500
            End If
            .ColAlignment(mHeadCol.����) = 4
            .ColAlignment(mHeadCol.��־) = 4
            .ColAlignment(mHeadCol.����) = 1
            .ColAlignment(mHeadCol.����) = 0
            .ColAlignment(mHeadCol.�շ�) = 0
            .ColAlignment(mHeadCol.������) = 0
            .ColAlignment(mHeadCol.NO) = 4
            .ColAlignment(mHeadCol.���) = 7
            .ColAlignment(mHeadCol.����) = 4
            .ColAlignment(mHeadCol.����) = 4
            .ColAlignment(mHeadCol.�ɲ���) = 0
            .ColAlignment(mHeadCol.˵��) = 1
        End With
        '�ָ�������
       ' Call RestoreFlexState(mshHead, Me.Name & "\" & TabShow.Tab)
    Case 2
        '��ϸ����ʽ
        With mshBody
               If blnInitColHead Then
                    .Clear
                    .Rows = 2
                    .Cols = mBodyCol.Cols
                    .TextMatrix(0, mBodyCol.����ID) = "����id"
                    .TextMatrix(0, mBodyCol.����ID) = "����id"
                    .TextMatrix(0, mBodyCol.���÷���) = "���÷���"
                    .TextMatrix(0, mBodyCol.����) = "����"
                    .TextMatrix(0, mBodyCol.����) = "����"
                    .TextMatrix(0, mBodyCol.����ҽ��) = "����ҽ��"
                    .TextMatrix(0, mBodyCol.״̬) = "��־"
                    .TextMatrix(0, mBodyCol.����) = "����"
                    .TextMatrix(0, mBodyCol.NO) = "NO"
                    .TextMatrix(0, mBodyCol.����) = "����"
                    .TextMatrix(0, mBodyCol.����) = "����"
                    .TextMatrix(0, mBodyCol.סԺ��) = "סԺ��"
                    .TextMatrix(0, mBodyCol.��������) = "��������"
                    .TextMatrix(0, mBodyCol.���) = "���"
                    .TextMatrix(0, mBodyCol.����) = "����"
                    .TextMatrix(0, mBodyCol.��λ) = "��λ"
                    .TextMatrix(0, mBodyCol.����ϵ��) = "����ϵ��"
                    .TextMatrix(0, mBodyCol.����) = "����"
                    .TextMatrix(0, mBodyCol.����) = "����"
                    .TextMatrix(0, mBodyCol.ԭʼ����) = "ԭʼ����"
                    .TextMatrix(0, mBodyCol.������) = "������"
                    .TextMatrix(0, mBodyCol.׼����) = "׼����"
                    .TextMatrix(0, mBodyCol.������) = "������"
                    .TextMatrix(0, mBodyCol.����) = "����"
                    .TextMatrix(0, mBodyCol.���) = "���"
                    .TextMatrix(0, mBodyCol.�����) = "�����"
                    .TextMatrix(0, mBodyCol.����Ա) = "����Ա"
                    .TextMatrix(0, mBodyCol.������) = "������"
                    .TextMatrix(0, mBodyCol.�ɲ���) = "�ɲ��� "
                    .TextMatrix(0, mBodyCol.��¼״̬) = "��¼״̬"
                    
                    .TextMatrix(0, mBodyCol.������) = "������"
                    .TextMatrix(0, mBodyCol.����ʱ��) = "����ʱ��"
            End If
            For intCol = 0 To .Cols - 1
                .ColAlignmentFixed(intCol) = 4
            Next
            
            .ColWidth(mBodyCol.����ID) = 0
            .ColWidth(mBodyCol.����ID) = 0
            .ColWidth(mBodyCol.����) = 0
            .ColWidth(mBodyCol.����ϵ��) = 0
            .ColWidth(mBodyCol.���÷���) = 0
            .ColWidth(mBodyCol.�ɲ���) = 0
            .ColWidth(mBodyCol.��¼״̬) = 0
            .ColWidth(mBodyCol.ԭʼ����) = 0
            
            Dim bytK3 As Byte   '��ǰΪ����
            Dim byt���ʱ� As Byte '���ڼ��ʱ�
            Dim byt����ģʽ As Byte
            
            byt����ģʽ = IIf(tabShow.Tab = 0, 0, 1)
            If byt����ģʽ = 1 Then
                bytK3 = 1
                byt���ʱ� = 1
            Else
                bytK3 = IIf(tbsSel.SelectedItem.Key = "K3", 0, 1)
                byt���ʱ� = IIf(InStr(1, mstrSelCon, ":26") <> 0, 1, 0)
                If Trim(mshHead.TextMatrix(mshHead.Row, mHeadCol.����)) <> "" And byt���ʱ� = 0 Then
                    byt���ʱ� = IIf(Val(mshHead.TextMatrix(mshHead.Row, mHeadCol.����)) = 26, 1, 0)
                End If
            End If
            
              If tabShow.Tab = 0 Then
                    'δ��
                    If tbsSel.SelectedItem.Key = "K1" Then '��ϸ
                        .Tag = "_��ϸ"
                    ElseIf tbsSel.SelectedItem.Key = "K2" Then 'ѡ�񵥾���ϸ
                        .Tag = "_ѡ��"
                    Else
                        .Tag = "_����"
                    End If
              Else
                    '�ѷ�
                    .Tag = "_�ѷ�"
              End If
              If RestoreFlexState(mshBody, Me.Caption) = False Then
                    .ColWidth(mBodyCol.����) = 1400 * bytK3
                    .ColWidth(mBodyCol.����ҽ��) = 1000 * bytK3
                    .ColWidth(mBodyCol.״̬) = 800 * bytK3 * byt���ʱ�
                    .ColWidth(mBodyCol.����) = 800 * bytK3 * byt���ʱ�
                    .ColWidth(mBodyCol.NO) = 1000 * bytK3
                    .ColWidth(mBodyCol.����) = 1000 * bytK3 * byt���ʱ�
                    .ColWidth(mBodyCol.����) = 800 * bytK3 * byt���ʱ�
                    .ColWidth(mBodyCol.סԺ��) = 800 * bytK3 * byt���ʱ�
                    .ColWidth(mBodyCol.��������) = 1600
                    .ColWidth(mBodyCol.���) = 1400
                    .ColWidth(mBodyCol.����) = 1000
                    .ColWidth(mBodyCol.��λ) = 800
                    .ColWidth(mBodyCol.����) = 800
                    .ColWidth(mBodyCol.����) = 800
                    .ColWidth(mBodyCol.������) = 800 * bytK3 * byt����ģʽ
                    .ColWidth(mBodyCol.׼����) = 800 * bytK3 * byt����ģʽ
                    .ColWidth(mBodyCol.������) = 800 * bytK3 * byt����ģʽ
                    .ColWidth(mBodyCol.����) = 1000
                    .ColWidth(mBodyCol.���) = 1000
                    .ColWidth(mBodyCol.�����) = 1000
                    .ColWidth(mBodyCol.����Ա) = 800 * bytK3
                    .ColWidth(mBodyCol.������) = 800 * bytK3
                    .ColWidth(mBodyCol.����ʱ��) = 1400 * bytK3
             End If
                    
            .ColAlignment(mBodyCol.����) = 1
            .ColAlignment(mBodyCol.����ҽ��) = 4
            .ColAlignment(mBodyCol.״̬) = 4
            .ColAlignment(mBodyCol.����) = 4
            .ColAlignment(mBodyCol.NO) = 4
            .ColAlignment(mBodyCol.����) = 4
            .ColAlignment(mBodyCol.����) = 4
            .ColAlignment(mBodyCol.סԺ��) = 4
            .ColAlignment(mBodyCol.��������) = 1
            .ColAlignment(mBodyCol.���) = 1
            .ColAlignment(mBodyCol.����) = 4
            .ColAlignment(mBodyCol.��λ) = 1
            .ColAlignment(mBodyCol.����) = 7
            .ColAlignment(mBodyCol.����) = 7
            .ColAlignment(mBodyCol.������) = 7
            .ColAlignment(mBodyCol.׼����) = 7
            .ColAlignment(mBodyCol.������) = 7
            .ColAlignment(mBodyCol.����) = 7
            .ColAlignment(mBodyCol.���) = 7
            .ColAlignment(mBodyCol.�����) = 7
            .ColAlignment(mBodyCol.����Ա) = 4
            .ColAlignment(mBodyCol.������) = 4
            .ColAlignment(mBodyCol.����ʱ��) = 4
        End With
    End Select
End Sub
Private Sub GetHeadData(ByVal int���� As Integer)
    Dim rsTemp As New ADODB.Recordset
    Dim strCon As String        '������
    Dim lng���ϲ���ID As Long
    Dim n As Integer
    
    If tabShow.Tab = 1 Then Exit Sub
    
    lng���ϲ���ID = cboStock.ItemData(cboStock.ListIndex)
    
    '24-�շѴ������ϣ�25-���ʵ��������ϣ�26-���ʱ������ϣ�
    
    If mint���� = 0 Then
        strCon = " And A.���� In (24,25,26)" '���ＰסԺ���е���
    Else
        If mint���� = 1 Then
            strCon = " And A.���� In (24,25) And A.��ҳID Is NULL " '���ﻮ�ۼ��������
        Else
            strCon = " And A.���� IN (25,26) And A.��ҳID Is Not NULL " 'סԺ����
        End If
    End If
    
    '��־,    ����,    ����,    ����,    �շ�,    ������,    NO  ,    ����,����,סԺ��,����,    �ɲ���,    ˵��,
    'δ������
    gstrSQL = " Select '' ��־, ����, ����,����,���շ�,��ҩ�� ������,NO,����,����,סԺ��,ltrim(rtrim(to_char(���۽��," & mOraFMT.FM_��� & ")))  AS ���,����,�ɲ���,˵�� " & _
              " From (" & _
              "     Select A.���ȼ�,A.����,D.���� as ����,A.����,A.���շ�,A.��ҩ��,A.NO ,decode(a.����,26,'',A.����) ����,Max(decode(a.����,26,'' ,decode(H.�����־,2,H.����,''))) as ����,max(decode(a.����,26,'' ,decode(H.�����־,2,H.��ʶ��,''))) as סԺ��,Sum(C.���۽��) ���۽��,A.����,A.�ɲ���,A.˵��" & _
              "     From ( " & _
              "             Select B.�����,B.סԺ��,A.���ȼ�,A.��������,Decode(A.����,24,'�շ�',25,'���ʵ�',26,'���ʱ�') ����,A.����,A.���շ�,'' ��ҩ��,A.No,A.����,To_Char(A.��������,'yyyy-MM-dd hh24:mi:ss') ����,1 �ɲ���,' ' ˵�� " & _
              "             From δ��ҩƷ��¼ A,������Ϣ B" & _
              "             Where A.����ID=B.����ID(+) and  nvl(a.���շ�,0)=1 and  (A.�ⷿID=[9] Or A.�ⷿID Is NULL) " & _
              "                     AND A.�������� between [7] and [8]" & _
                                  IIf(mlng����id = 0, "", " And A.����id=[1]") & _
                                  IIf(mstrStartNo = "", "", " And A.NO >= [2]") & _
                                  IIf(mstrEndNo = "", "", " And A.NO <= [3]") & _
                                  IIf(mstr�������� = "", "", " And A.���� like [4]") & _
                                  IIf(mstr���� = "", "", " And A.���� in (" & mstr���� & ")") & _
              "                     " & strCon & _
              "             ) A,ҩƷ�շ���¼ C,���˷��ü�¼ H,���ű� D" & _
              "     Where A.����=C.���� and nvl(c.��ҩ��ʽ,0)<>-1 and C.����id=H.id   and H.��������ID=D.id(+) And A.NO=C.NO And C.����� Is NULL And MOD(C.��¼״̬,3)=1" & _
              "         " & IIf(mlng����id = 0, "", " And H.��������id+0=[5]") & _
              "         " & IIf(Val(mstrסԺ��) = 0, "", " And H.��ʶ��=[6]") & _
              "     GROUP BY A.���ȼ�,A.����,D.����,A.����,A.���շ�,A.��ҩ��,A.No,A.����,A.����,A.�ɲ���,A.˵��)"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng����id, mstrStartNo, mstrEndNo, mstr��������, mlng����id, mstrסԺ��, CDate(mstrStartDate), CDate(mstrEndDate), lng���ϲ���ID)
        
    If rsTemp.RecordCount <> 0 Then
        Set mshHead.DataSource = rsTemp
        Call SetGrdColHead(1, False)
    Else
        Call SetGrdColHead(1)
        Call SetGrdColHead(2)
    End If
    
    '�����ҩƷ���ڴ��룬��λ������ĵ��ݺ�
    If mblnTrans Then
        With mshHead
            For n = 1 To .Rows - 1
                If Trim(.TextMatrix(n, mHeadCol.NO)) = mstrNo Then
                    .Row = n
                    Call mshHead_EnterCell
                    .TopRow = n
                    Exit For
                End If
            Next
        End With
    End If
End Sub
Private Function GetSelCon(Optional strAliaName As String = "A") As String
    '��ȡ��ѡ�������
    Dim strArr(0 To 1)
    Dim strTemp As String
    Dim strCon(0 To 2) As String    '�ֱ�������
    Dim i As Integer
    
    If mstrSelCon = "" Then Exit Function
    
    strArr(0) = Split(Mid(mstrSelCon, 3), "||")
    For i = 0 To UBound(strArr(0))
        strArr(1) = Split(strArr(0)(i), ":")
        Select Case strArr(1)(1)
        Case 24
            strCon(0) = strCon(0) & ",'" & strArr(1)(0) & "'"
        Case 25
            strCon(1) = strCon(1) & ",'" & strArr(1)(0) & "'"
        Case Else
            strCon(2) = strCon(2) & ",'" & strArr(1)(0) & "'"
        End Select
    Next
    
    strTemp = ""
    If strCon(0) <> "" Then
        strTemp = " Or (" & IIf(strAliaName = "", "", strAliaName & ".") & "NO in (" & Mid(strCon(0), 2) & " ) And " & IIf(strAliaName = "", "", strAliaName & ".") & "����=24) "
    End If
    If strCon(1) <> "" Then
        strTemp = strTemp & " Or (" & IIf(strAliaName = "", "", strAliaName & ".") & "NO in (" & Mid(strCon(1), 2) & " ) And " & IIf(strAliaName = "", "", strAliaName & ".") & "����=25) "
    End If
    If strCon(2) <> "" Then
        strTemp = strTemp & " Or (" & IIf(strAliaName = "", "", strAliaName & ".") & "NO in (" & Mid(strCon(2), 2) & " ) And " & IIf(strAliaName = "", "", strAliaName & ".") & "����=26) "
    End If
    If strTemp <> "" Then
        strTemp = " (" & Mid(strTemp, 4) & ") "
    End If
    GetSelCon = strTemp
End Function

Private Function GetBodyQurysSQL(Optional ByVal strTbsSelKey As String = "") As ADODB.Recordset
    '����:��ȡ��ѯ��ϸ����
    '����:
    Dim lng���ϲ���ID As Long
    Dim strFields  As String
    Dim strCon As String
    Dim strKey As String
    Dim strTableName As String
    Dim rsTemp As New ADODB.Recordset
    Dim int���� As Long
    Dim strNo As String
    
    lng���ϲ���ID = cboStock.ItemData(cboStock.ListIndex)
    
    
    Select Case mintUnit
    Case 0  'ɢװ��λ
         strFields = "C.���㵥λ ��λ,D.����ϵ��,ltrim(to_char(B.����,'9999999999')) ����,B.ʵ������ as ԭʼ����,ltrim(to_char(B.ʵ������ ," & mOraFMT.FM_���� & "  )) ����,ltrim(to_char(B.���ۼ�," & mOraFMT.FM_���ۼ� & ")) ����,trim(to_char(K.ʵ������," & mOraFMT.FM_���� & "))  �����, "
    Case Else
         strFields = "D.��װ��λ ��λ,D.����ϵ��,ltrim(to_char(B.����,'9999999999')) ����,B.ʵ������ as ԭʼ����,ltrim(to_char(B.ʵ������/D.����ϵ��," & mOraFMT.FM_���� & ")) ����,ltrim(to_char(B.���ۼ�*D.����ϵ��," & mOraFMT.FM_���ۼ� & ")) ����,trim(to_char(K.ʵ������/D.����ϵ��," & mOraFMT.FM_���� & "))  �����, "
    End Select
    
    
    If tabShow.Tab = 0 Then
        'δ����ģʽ
        If strTbsSelKey = "" Then
            strKey = tbsSel.SelectedItem.Key
        Else
            strKey = strTbsSelKey
        End If
        Select Case strKey
            Case "K1"           '�鵥����ϸ
                    strTableName = " ҩƷ�շ���¼ B,���˷��ü�¼ H"
                    int���� = Val(Decode(mshHead.TextMatrix(mshHead.Row, mHeadCol.����), "�շ�", "24", "���ʵ�", "25", 26))
                    strNo = mshHead.TextMatrix(mshHead.Row, mHeadCol.NO)
                    If strNo <> "" Then
                        If zlDatabase.NOMoved("ҩƷ�շ���¼", strNo, "����=", int����) Then
                            strTableName = " HҩƷ�շ���¼ B,H���˷��ü�¼ H"
                        End If
                    End If
                    gstrSQL = "" & _
                        "  SELECT DISTINCT 0 as �ɲ���,B.��¼״̬,B.id,B.����ID as ����id, B.ҩƷID as ����id,NVL(B.����,0) ����,NVL(D.���÷���,0) ���÷���," & _
                        "       T.���� ����,H.������ ����ҽ��,'' ��־ ,'' ����,B.NO,H.����,H.����,H.��ʶ�� סԺ��," & _
                        "      '['||C.����||']'||C.����  ��������,H.���,DECODE(C.���,NULL,C.����,DECODE(C.����,NULL,C.���,C.���||'|'||C.����)) ���," & _
                        "      DECODE(B.����,NULL,'',B.����)||DECODE(B.����,NULL,'',0,'','('||B.����||')') ����, " & strFields & _
                        "      0 as ������,0 as ׼����,0 as ������,B.���۽�� as ���," & _
                        "      H.������ as ����Ա ,B.��������,H.����Ա���� as ����Ա,B.����� ������ " & _
                        " FROM  " & strTableName & ",�������� D,�շ���ĿĿ¼ C, " & _
                        "      ���ű� S,���ű� T,(Select �ⷿid,ҩƷid,����,ʵ������ From ҩƷ��� where ����=1) K" & _
                        " WHERE D.����ID=C.ID and nvl(b.��ҩ��ʽ,0)<>-1 " & _
                        "      AND H.��������ID=T.ID(+) AND B.ҩƷID=D.����ID AND MOD(B.��¼״̬,3)=1  and B.NO=[1] and B.����=[3] " & _
                        "      AND S.ID=NVL(B.�ⷿID,[2]) AND B.����ID=H.ID " & _
                        "      AND NVL(B.�ⷿID,[2])+0=[2] AND LTRIM(RTRIM(NVL(B.ժҪ,'�ܷ���')))<>'�ܷ�'" & _
                        "      AND B.ҩƷID=K.ҩƷID(+) AND NVL(B.�ⷿID,[2])=K.�ⷿID(+) AND NVL(B.����,0)=NVL(K.����(+),0)  " & _
                        "      AND B.����� IS NULL "
                    gstrSQL = gstrSQL & " Order by H.���,B.ҩƷID,Nvl(B.����,0)"
                    
                    
            Case "K2"           '�鱻ѡ��ĵ�����ϸ
                strCon = GetSelCon("B")
                strCon = IIf(strCon = "", "1=2", strCon)
                
                gstrSQL = "" & _
                    "  SELECT DISTINCT 0 as �ɲ���,B.��¼״̬,B.id,B.����ID as ����id, B.ҩƷID as ����id,NVL(B.����,0) ����,NVL(D.���÷���,0) ���÷���," & _
                    "       T.���� ����,H.������ ����ҽ��,'' ��־ ,'' ����,B.NO,H.����,H.����,H.��ʶ�� סԺ��," & _
                    "      '['||C.����||']'||C.����  ��������,H.���,DECODE(C.���,NULL,C.����,DECODE(C.����,NULL,C.���,C.���||'|'||C.����)) ���," & _
                    "      DECODE(B.����,NULL,'',B.����)||DECODE(B.����,NULL,'',0,'','('||B.����||')') ����, " & strFields & _
                    "      0 as ������,0 as ׼����,0 as ������,B.���۽�� as ���," & _
                    "      H.������ as ����Ա ,B.��������,H.����Ա���� as ����Ա,B.����� ������ " & _
                    " FROM ҩƷ�շ���¼ B,���˷��ü�¼ H,�������� D,�շ���ĿĿ¼ C, " & _
                    "      ���ű� S,���ű� T,(Select �ⷿid,ҩƷid,����,ʵ������ From ҩƷ��� where ����=1) K" & _
                    " WHERE D.����ID=C.ID  and nvl(b.��ҩ��ʽ,0)<>-1 " & _
                    "      AND H.��������ID=T.ID(+) AND B.ҩƷID=D.����ID AND MOD(B.��¼״̬,3)=1 " & _
                    "      AND S.ID=NVL(B.�ⷿID,[2]) AND B.����ID=H.ID " & _
                    "      AND NVL(B.�ⷿID,[2])+0=[2] AND LTRIM(RTRIM(NVL(B.ժҪ,'�ܷ���')))<>'�ܷ�'" & _
                    "      AND B.ҩƷID=K.ҩƷID(+) AND NVL(B.�ⷿID,[2])=K.�ⷿID(+) AND NVL(B.����,0)=NVL(K.����(+),0)  " & _
                    "      AND B.����� IS NULL And " & strCon
                gstrSQL = gstrSQL & " Order by B.��������,B.NO,H.���,B.ҩƷID,Nvl(B.����,0)"
            Case "K3"           '�鱻ѡ��ĵ��ݵĻ�����ϸ
                strCon = GetSelCon("B")
                strCon = IIf(strCon = "", "1=2", strCon)
                
'                If strTbsSelKey <> "" Then
'                    'ֻ��ȡ����id,��Ʒ�����,����,����,����
'                    gstrSQL = "" & _
'                        "  SELECT DISTINCT 0 as �ɲ���,0 ��¼״̬,0 as id,0 as ����ID,B.ҩƷID as ����id,NVL(B.����,0) ����," & _
'                        "       '['||C.����||']'||C.����  ��������,DECODE(C.���,NULL,C.����,DECODE(C.����,NULL,C.���,C.���||'|'||C.����)) ���," & _
'                        "       DECODE(B.����,NULL,'',B.����)||DECODE(B.����,NULL,'',0,'','('||B.����||')') ����," & _
'                        "       SUM(nvl(B.ʵ������,0)) as ʵ������,SUM(nvl(C.ʵ������,0)) as �������" & _
'                        "      0 as ������,0 as ׼����,0 as ������,sum(nvl(B.���۽��,0)) as ���" & _
'                        " FROM ҩƷ�շ���¼ B,�շ���ĿĿ¼ C,(Select * From ҩƷ��� where ����=1) K" & _
'                        " WHERE B.ҩƷID=C.ID and nvl(b.��ҩ��ʽ,0)<>-1 AND MOD(B.��¼״̬,3)=1 AND NVL(B.�ⷿID,[2])+0=[2] " & _
'                        "       AND LTRIM(RTRIM(NVL(B.ժҪ,'�ܷ���')))<>'�ܷ�'" & _
'                        "      AND B.ҩƷID=K.ҩƷID(+) AND NVL(B.�ⷿID,[2])=K.�ⷿID(+) AND NVL(B.����,0)=NVL(K.����(+),0)  " & _
'                        "      AND B.����� IS NULL And " & strCon & _
'                        " Group by B.ҩƷid,B.����,b.����,C.����,C.����,DECODE(C.���,NULL,C.����,DECODE(C.����,NULL,C.���,C.���||'|'||C.����))" & _
'                        "       "
'
'                Else
                    
                    Select Case mintUnit
                    Case 0  'ɢװ��λ
                         strFields = "C.���㵥λ ��λ,max(D.����ϵ��) ����ϵ��,ltrim(to_char(sum(nvl(B.����,1)),'9999999999')) ����,sum(nvl(B.ʵ������,0)) as ԭʼ����,ltrim(to_char(sum(nvl(B.ʵ������,0))," & mOraFMT.FM_���� & ")) ����,ltrim(to_char(sum(nvl(B.���ۼ�,0))," & mOraFMT.FM_���ۼ� & ")) ����,trim(to_char(sum(nvl(K.ʵ������,0))," & mOraFMT.FM_���� & "))  �����, "
                    Case Else
                         strFields = "D.��װ��λ ��λ,max(D.����ϵ��) ����ϵ��,ltrim(to_char(sum(nvl(B.����,0)),'9999999999')) ����,sum(nvl(B.ʵ������,0)) as ԭʼ����,ltrim(to_char(sum(nvl(B.ʵ������/D.����ϵ��,0))," & mOraFMT.FM_���� & ")) ����,ltrim(to_char(sum(nvl(B.���ۼ�*D.����ϵ��,0))," & mOraFMT.FM_���ۼ� & ")) ����,trim(to_char(sum(nvl(K.ʵ������/D.����ϵ��,0))," & mOraFMT.FM_���� & "))  �����, "
                    End Select
                                  
                                  
                    gstrSQL = "" & _
                        "  SELECT DISTINCT 0 as �ɲ���,0 ��¼״̬,0 id,0 as ����id, B.ҩƷID as ����id,NVL(B.����,0) ����,'' ���÷���," & _
                        "       '' ����,'' ����ҽ��,'' ��־ ,'' ����,'' NO,'' ����,'' ����,''  סԺ��," & _
                        "      '['||C.����||']'||C.����  ��������,0 ���,DECODE(C.���,NULL,C.����,DECODE(C.����,NULL,C.���,C.���||'|'||C.����)) ���," & _
                        "      DECODE(B.����,NULL,'',B.����)||DECODE(B.����,NULL,'',0,'','('||B.����||')') ����, " & strFields & _
                        "      0 as ������,0 as ׼����,0 as ������,sum(nvl(B.���۽��,0)) as ���," & _
                        "      '' as ����Ա ,'' ��������,''  ����Ա,'' ������ " & _
                        " FROM ҩƷ�շ���¼ B,�������� D,�շ���ĿĿ¼ C, " & _
                        "      ���˷��ü�¼ H,���ű� S,���ű� T,(Select �ⷿid,ҩƷid,����,ʵ������ From ҩƷ��� where ����=1) K" & _
                        " WHERE D.����ID=C.ID  and nvl(b.��ҩ��ʽ,0)<>-1" & _
                        "      AND H.��������ID=T.ID(+) AND B.ҩƷID=D.����ID AND MOD(B.��¼״̬,3)=1 " & _
                        "      AND S.ID=NVL(B.�ⷿID,[2]) AND B.����ID=H.ID " & _
                        "      AND NVL(B.�ⷿID,[2])+0=[2] AND LTRIM(RTRIM(NVL(B.ժҪ,'�ܷ���')))<>'�ܷ�'" & _
                        "      AND B.ҩƷID=K.ҩƷID(+) AND NVL(B.�ⷿID,[2])=K.�ⷿID(+) AND NVL(B.����,0)=NVL(K.����(+),0)  " & _
                        "      AND B.����� IS NULL And " & strCon & _
                        " Group by B.ҩƷid,b.����,C.����,C.����,DECODE(C.���,NULL,C.����,DECODE(C.����,NULL,C.���,C.���||'|'||C.����))," & _
                        "       B.����,B.����," & IIf(mintUnit = 0, "C.���㵥λ", "D.��װ��λ")
               ' End If
        End Select
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strNo, lng���ϲ���ID, int����)
        Set GetBodyQurysSQL = rsTemp
        Exit Function
    End If
    
     
    Dim strCon1 As String
    strCon1 = ""
    '24-�շѴ������ϣ�25-���ʵ��������ϣ�26-���ʱ������ϣ�
    If mint���� = 0 Then
        strCon = " And S.���� In (24,25,26)" '���ＰסԺ���е���
    Else
        If mint���� = 1 Then
            strCon = " And S.���� In (24,25)" '���ﻮ�ۼ��������
            strCon1 = " and M.��ҳID is null"
        Else
            strCon = " And S.���� IN (24,25,26)  " 'סԺ����
            strCon1 = " and M.��ҳID is not null"
        End If
    End If
    If mstr���� <> "" Then
        strCon = strCon & " and S.���� in (" & mstr���� & ") "
    End If
    
    Select Case mintUnit
    Case 0  'ɢװ��λ
         strFields = "S.���㵥λ ��λ,D.����ϵ��,ltrim(to_char(S.����,'9999999999')) ����,s.ʵ������ as ԭʼ����,ltrim(to_char(S.ʵ������," & mOraFMT.FM_���� & ")) ����,ltrim(to_char(S.��������," & mOraFMT.FM_���� & ")) as ������,ltrim(to_char(S.�ѷ�����," & mOraFMT.FM_���� & ")) as ׼����,'' ������,ltrim(to_char(S.���ۼ�," & mOraFMT.FM_���ۼ� & ")) ����,''  �����, "
    Case Else
         strFields = "D.��װ��λ ��λ,D.����ϵ��,ltrim(to_char(S.����,'9999999999')) ����,s.ʵ������ as ԭʼ����,ltrim(to_char(S.ʵ������/D.����ϵ��," & mOraFMT.FM_���� & ")) ����,ltrim(to_char(S.��������/D.����ϵ��," & mOraFMT.FM_���� & ")) as ������,ltrim(to_char(S.�ѷ�����/D.����ϵ��," & mOraFMT.FM_���� & ")) as ׼����,'' ������,ltrim(to_char(S.���ۼ�*D.����ϵ��," & mOraFMT.FM_���ۼ� & ")) ����,'' �����, "
    End Select
    
    Dim strTemp As String
    Dim blnHistory As Boolean
    blnHistory = zlDatabase.DateMoved(mstrStartDate, , , Me.Caption)
    
    If Chk�嵥.Value = 0 Then
        '��ȡ�ѷ��ϻ����ϵĽ��
        
        gstrSQL = " SELECT DISTINCT S.id,S.��¼״̬ ,S.����ID,s.����ҽ��,'' ��־,decode(S.����,24,'�շ�',25,'���ʵ�',26,'���ʱ�' ) as ����,s.סԺ��,s.����Ա,s.����Ա, S.ID,S.����,S.ҩƷID as ����id,S.NO,S.����,P.���� ����,s.�����־,s.����,s.����," & _
            " '['||X.����||']'||X.����  ��������,NVL(D.���÷���,0) ���÷���,DECODE(x.���,NULL,x.����,DECODE(x.����,NULL,x.���,x.���||'|'||x.����)) ���," & strFields & _
            "  DECODE(S.����,NULL,'',S.����)||DECODE(S.����,NULL,'',0,'','('||S.����||')') ����,NVL(S.����,0) ����,S.Ч��," & _
            "  S.���۽�� ���,S.ժҪ ˵��,S.�����,TO_CHAR(S.�������,'YYYY-MM-DD HH24:MI:SS') ����ʱ��,s.�ɲ���" & _
            " FROM (    SELECT A.ID,A.NO,A.����,A.ҩƷID,A.����ID,A.����,A.����,A.Ч��,NVL(A.����,0) ����," & _
            "                   NVL(A.����,1) ����,A.ʵ������ ʵ������,NVL(A.����,1)*A.ʵ������-B.�ѷ����� ��������,B.�ѷ�����,A.��¼״̬," & _
            "                   A.���ۼ�,A.���۽��,A.����,A.Ƶ��,A.�÷�,A.ժҪ,A.�����,A.�������,A.�Է�����ID,A.�ⷿID,A.����ҽ��,A.���㵥λ,A.סԺ��,A.����Ա,A.����Ա,A.�����־,A.����,A.����,A.�ɲ���" & _
            "           FROM(  "
            
             strTemp = "" & _
                "   SELECT A.ID,A.NO,A.ҩƷid,A.���,A.����,A.����ID,A.����,A.����,A.Ч��,nvl(A.����,0) ����,nvl(A.����,0) ����,A.ʵ������,A.��¼״̬," & _
                "        A.���ۼ�,A.���۽��,A.����,A.Ƶ��,A.�÷�,A.ժҪ,A.�����,A.�������,A.�Է�����id,A.�ⷿID," & _
                "        m.������ as ����ҽ��,M.���㵥λ,m.��ʶ�� as סԺ��,m.����Ա���� as ����Ա,m.������ ����Ա,m.�����־,m.����,m.����,1 �ɲ��� " & _
                "   FROM ҩƷ�շ���¼ A,���˷��ü�¼ M" & _
                "   WHERE A.����� IS NOT NULL and A.����id=M.ID  and nvl(a.��ҩ��ʽ,0)<>-1 AND (A.��¼״̬=1 OR MOD(A.��¼״̬,3)=0)" & _
                "       AND A.�ⷿID+0=[9]" & _
                "       AND A.������� BETWEEN [7] AND [8] " & _
                        IIf(mlng����id = 0, "", " AND M.����ID+0=[1]") & _
                        IIf(Val(mstrסԺ��) = 0, "", " AND M.��ʶ��+0=[2]") & _
                        IIf(mstr�������� = "", "", " AND M.���� LIKE [3]") & _
                        IIf(mstrStartNo = "", "", " AND A.NO>=[4]") & _
                        IIf(mstrEndNo = "", "", " AND A.NO<=[5]") & _
                        IIf(mlng����id = 0, "", " And M.ִ�в���id+0=[6]")
            
            If blnHistory Then
                strTemp = AnalyseHistorySQL(strTemp, "1 �ɲ���", "-99 �ɲ���")
            End If
            gstrSQL = gstrSQL & strTemp & " ) A,(    "
            
            
            strTemp = "" & _
                "               SELECT A.NO,A.����,A.ҩƷID,A.���,SUM(NVL(A.����,1)*A.ʵ������) �ѷ�����" & _
                "               FROM ҩƷ�շ���¼ A" & _
                "               WHERE A.����� IS NOT NULL and nvl(a.��ҩ��ʽ,0)<>-1 AND A.�ⷿID+0=[9]" & _
                "                        and (A.NO,����) in (Select NO,���� From ҩƷ�շ���¼ " & _
                "                                            where ����� is not null and nvl(��ҩ��ʽ,0)<>-1  AND " & _
                "                                                  (��¼״̬=1 OR MOD(��¼״̬,3)=0) and �ⷿid+0 =[9]" & _
                                                                IIf(mstrStartNo = "", "", " AND NO>=[4]") & _
                                                                IIf(mstrEndNo = "", "", " AND  NO<=[5]") & _
                "                                                   AND ������� BETWEEN [7] AND [8]) " & _
                "               GROUP BY A.NO,A.����,A.ҩƷID,A.���    "
            
            If blnHistory Then
                    strTemp = AnalyseHistorySQL(strTemp)
            End If
            gstrSQL = gstrSQL & strTemp & " ) B" & _
                "           WHERE A.NO = B.NO AND A.���� = B.���� AND A.ҩƷID+0 = B.ҩƷID AND A.��� = B.��� AND B.�ѷ�����<>0) S,"
    
            gstrSQL = gstrSQL & "" & _
                "      ���ű� P,�������� D,�շ���ĿĿ¼ X" & _
                " WHERE S.ҩƷID=D.����ID AND S.�Է�����ID+0=P.ID  AND d.����ID=X.ID" & _
                "       AND (S.��¼״̬=1 OR MOD(S.��¼״̬,3)=0) AND S.ʵ������*S.����>S.�������� " & _
                "       AND S.����� IS NOT NULL AND S.�ⷿID+0=[9]" & strCon
            gstrSQL = gstrSQL & " Order By S.No,S.����"
        
    Else
        '�嵥��ʾÿ�ʲ�������
        
        gstrSQL = " SELECT DISTINCT  S.id,S.��¼״̬,S.����ID,S.����ҽ��,'' ��־,decode(S.����,24,'�շ�',25,'���ʵ�',26,'���ʱ�' ) as ����,s.סԺ��,s.����Ա,s.����Ա,S.ID,S.����,S.ҩƷID ����id,S.NO,S.����,P.���� ����,s.�����־,s.����,s.����,'['||X.����||']'||X.����  ��������," & _
                 "          NVL(D.���÷���,0)  ���÷���,DECODE(x.���,NULL,x.����,DECODE(x.����,NULL,x.���,x.���||'|'||X.����)) ���," & strFields & _
                 "          DECODE(S.����,NULL,'',S.����)||DECODE(S.����,NULL,'',0,'','('||S.����||')') ����,NVL(S.����,0) ����,S.Ч��," & _
                 "          S.���ۼ� ����,S.���۽�� ���,S.ժҪ ˵��,TO_CHAR(S.�������,'YYYY-MM-DD HH24:MI:SS') ����ʱ��,S.�����,S.�������,decode(S.�ѷ�����,0,0,�ɲ���) as �ɲ���" & _
                 " FROM "
                 
        gstrSQL = gstrSQL & _
                 "      (   SELECT * FROM" & _
                 "              (   SELECT A.ID,A.NO,A.����,A.ҩƷID,A.����ID,A.����,A.����,A.Ч��,NVL(A.����,0) ����," & _
                 "                          NVL(A.����,1) ����,A.ʵ������,NVL(A.����,1)*A.ʵ������-B.�ѷ����� ��������,B.�ѷ�����,A.��¼״̬," & _
                 "                          A.���ۼ� , A.���۽��, A.����, A.Ƶ��, A.�÷�, A.ժҪ, A.�����, A.�������, A.�Է�����ID, A.�ⷿID,A.�ɲ���," & _
                 "                          A.����ҽ��,A.סԺ��,A.����Ա,A.���㵥λ,A.����Ա,A.�����־,A.����,A.���� " & _
                 "                  FROM (  "
                 
        strTemp = "" & _
                "   SELECT A.ID,A.NO,A.����,A.ҩƷid,A.���,A.����ID,A.����,A.����,A.Ч��,nvl(A.����,0) ����,nvl(A.����,0) ����,A.ʵ������,A.��¼״̬," & _
                "        A.���ۼ�,A.���۽��,A.����,A.Ƶ��,A.�÷�,A.ժҪ,A.�����,A.�������,A.�Է�����id,A.�ⷿID," & _
                "        m.������ as ����ҽ��,m.��ʶ�� as סԺ��,m.����Ա���� as ����Ա,m.���㵥λ,m.������ ����Ա,m.�����־,m.����,m.����,1 �ɲ��� " & _
                "   FROM ҩƷ�շ���¼ A,���˷��ü�¼ M" & _
                "   WHERE A.����� IS NOT NULL and A.����id=M.ID  and nvl(a.��ҩ��ʽ,0)<>-1 AND (A.��¼״̬=1 OR MOD(A.��¼״̬,3)=0)" & _
                "       AND A.�ⷿID+0=[9]" & _
                "       AND A.������� BETWEEN [7] AND [8] " & _
                        IIf(mlng����id = 0, "", " AND M.����ID+0=[1]") & _
                        IIf(Val(mstrסԺ��) = 0, "", " AND M.��ʶ��+0=[2]") & _
                        IIf(mstr�������� = "", "", " AND M.���� LIKE [3]") & _
                        IIf(mstrStartNo = "", "", " AND A.NO>=[4]") & _
                        IIf(mstrEndNo = "", "", " AND A.NO<=[5]") & _
                        IIf(mlng����id = 0, "", " And M.ִ�в���id+0=[6]") & strCon1
                 
            If blnHistory Then
                strTemp = AnalyseHistorySQL(strTemp, "1 �ɲ���", "-99 �ɲ���")
            End If
            
            gstrSQL = gstrSQL & strTemp & " ) A,( "
            strTemp = "" & _
                      "               SELECT A.NO,A.����,A.ҩƷID,A.���,SUM(NVL(A.����,1)*A.ʵ������) �ѷ�����" & _
                      "               FROM ҩƷ�շ���¼ A" & _
                      "               WHERE A.����� IS NOT NULL and nvl(a.��ҩ��ʽ,0)<>-1 AND A.�ⷿID+0=[9]" & _
                      "                        and (A.NO,����) in (Select NO,���� From ҩƷ�շ���¼ " & _
                      "                                            where ����� is not null and nvl(��ҩ��ʽ,0)<>-1  AND " & _
                      "                                                  (��¼״̬=1 OR MOD(��¼״̬,3)=0) and �ⷿid+0 =[9]" & _
                                                                      IIf(mstrStartNo = "", "", " AND NO>=[4]") & _
                                                                      IIf(mstrEndNo = "", "", " AND  NO<=[5]") & _
                      "                                                   AND ������� BETWEEN [7] AND [8]) " & _
                      "               GROUP BY A.NO,A.����,A.ҩƷID,A.���    "
                     
            If blnHistory Then
                strTemp = AnalyseHistorySQL(strTemp)
            End If
        
            gstrSQL = gstrSQL & strTemp & ") B" & _
                     "              WHERE A.NO = B.NO AND A.���� = B.���� AND A.ҩƷID+0 = B.ҩƷID AND A.��� = B.���)" & _
                     "              UNION"
            strTemp = "" & _
                     "              SELECT A.ID,A.NO,A.����,A.ҩƷID,A.����ID,A.����,A.����,A.Ч��,NVL(A.����,0)," & _
                     "                      NVL(A.����,1) ����,A.ʵ������,0 ������,0 �ѷ�����,A.��¼״̬," & _
                     "                      A.���ۼ� , A.���۽��, A.����, A.Ƶ��, A.�÷�, A.ժҪ, A.�����, A.�������, A.�Է�����ID, A.�ⷿID," & _
                     "                      DECODE(A.��¼״̬,1,1,DECODE(MOD(A.��¼״̬,3),0,1,MOD(A.��¼״̬,3)+1)) �ɲ���," & _
                     "                       m.������ as ����ҽ��,m.��ʶ�� as סԺ��,m.����Ա���� as ����Ա,M.���㵥λ,m.������ ����Ա,m.�����־,m.����,m.���� " & _
                     "              FROM ҩƷ�շ���¼ A, ���˷��ü�¼ M" & _
                     "              WHERE A.����� IS NOT NULL and a.����id=M.id   and nvl(a.��ҩ��ʽ,0)<>-1 AND NOT (a.��¼״̬=1 OR MOD(a.��¼״̬,3)=0)" & _
                     "                      AND A.�ⷿID+0=[9]" & _
                     "                      AND A.������� BETWEEN [7] AND [8]" & _
                                            IIf(mlng����id = 0, "", " AND M.����ID=[1]") & _
                                            IIf(Val(mstrסԺ��) = 0, "", " AND M.��ʶ��=[2]") & _
                                            IIf(mstr�������� = "", "", " AND M.���� LIKE  [3] ") & _
                                            IIf(mstrStartNo = "", "", " AND A.NO>=[4]") & _
                                            IIf(mstrEndNo = "", "", " AND A.NO<=[5]") & _
                                            IIf(mlng����id = 0, "", " And M.ִ�в���id+0=[6]") & strCon1
            If blnHistory Then
                strTemp = AnalyseHistorySQL(strTemp, "DECODE(A.��¼״̬,1,1,DECODE(MOD(A.��¼״̬,3),0,1,MOD(A.��¼״̬,3)+1)) �ɲ���", " -99 �ɲ���")
            End If
            
            
        gstrSQL = gstrSQL & strTemp & " ) S," & _
                 "      ���ű� P,�������� D,�շ���ĿĿ¼ X " & _
                 " WHERE S.ҩƷID=D.����ID AND d.����ID=x.ID AND S.�Է�����ID+0=P.ID" & _
                 " " & strCon & _
                 "    and  S.����� IS NOT NULL"
        gstrSQL = gstrSQL & " Order By S.No,S.����,S.�������"
    End If
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng����id, mstrסԺ��, mstr��������, mstrStartNo, mstrEndNo, mlng����id, CDate(mstrStartDate), CDate(mstrEndDate), lng���ϲ���ID)
    Set GetBodyQurysSQL = rsTemp
End Function
Private Function ReadBillData() As Boolean
    Dim RsBody As New ADODB.Recordset
    Dim IntStyle As Integer
        
    '--��ȡ��������--
    On Error GoTo ErrHand:
    err = 0
    ReadBillData = False
    
    Set RsBody = GetBodyQurysSQL

    'zlDatabase.OpenRecordset RsBody, gstrSQL, Me.Caption
    
    '���������
    If RsBody.RecordCount <> 0 Then
        '   ��������
        LoadBodyData RsBody
        Call SetGrdColHead(2, False)
    Else
        Call SetGrdColHead(2)
    End If
    ReadBillData = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SetGrdColHead(2)
    ReadBillData = False
End Function
Private Function LoadBodyData(ByVal RsBody As ADODB.Recordset) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:���ر�������
    '--�����:
    '--������:
    '--��  ��:
    '-----------------------------------------------------------------------------------------------------------
    Dim lngRow As Long
    err = 0
    On Error GoTo ErrHand:
    LoadBodyData = False
    With mshBody
        .Redraw = False
        .Rows = RsBody.RecordCount + 1
        lngRow = 1
        Do While Not RsBody.EOF
            
            .TextMatrix(lngRow, mBodyCol.����ID) = NVL(RsBody!����ID, 0)
            .TextMatrix(lngRow, mBodyCol.����ID) = NVL(RsBody!����ID, 0)
            .TextMatrix(lngRow, mBodyCol.���÷���) = NVL(RsBody!���÷���, 0)
            .TextMatrix(lngRow, mBodyCol.����) = NVL(RsBody!����, 0)
            .TextMatrix(lngRow, mBodyCol.����) = NVL(RsBody!����)
            .TextMatrix(lngRow, mBodyCol.����ҽ��) = NVL(RsBody!����ҽ��)
            .TextMatrix(lngRow, mBodyCol.״̬) = NVL(RsBody!��־)
            .TextMatrix(lngRow, mBodyCol.����) = NVL(RsBody!����)
            .TextMatrix(lngRow, mBodyCol.NO) = NVL(RsBody!NO)
            .TextMatrix(lngRow, mBodyCol.����) = NVL(RsBody!����)
            .TextMatrix(lngRow, mBodyCol.����) = NVL(RsBody!����)
            .TextMatrix(lngRow, mBodyCol.סԺ��) = NVL(RsBody!סԺ��)
            .TextMatrix(lngRow, mBodyCol.��������) = NVL(RsBody!��������)
            .TextMatrix(lngRow, mBodyCol.���) = NVL(RsBody!���)
            .TextMatrix(lngRow, mBodyCol.����) = NVL(RsBody!����)
            .TextMatrix(lngRow, mBodyCol.��λ) = NVL(RsBody!��λ)
            .TextMatrix(lngRow, mBodyCol.����ϵ��) = NVL(RsBody!����ϵ��)
            .TextMatrix(lngRow, mBodyCol.����) = NVL(RsBody!����)
            .TextMatrix(lngRow, mBodyCol.����) = NVL(RsBody!����)
            
            .TextMatrix(lngRow, mBodyCol.ԭʼ����) = NVL(RsBody!ԭʼ����)
            .TextMatrix(lngRow, mBodyCol.������) = NVL(RsBody!������)
            .TextMatrix(lngRow, mBodyCol.׼����) = NVL(RsBody!׼����)
            .TextMatrix(lngRow, mBodyCol.������) = NVL(RsBody!������)
            .TextMatrix(lngRow, mBodyCol.����) = NVL(RsBody!����)
            .TextMatrix(lngRow, mBodyCol.���) = Format(Val(NVL(RsBody!���)), mFMT.FM_���)
            .TextMatrix(lngRow, mBodyCol.�����) = Format(Val(NVL(RsBody!�����)), mFMT.FM_����)
            .TextMatrix(lngRow, mBodyCol.����Ա) = NVL(RsBody!����Ա)
            If tabShow.Tab = 0 Then
                .TextMatrix(lngRow, mBodyCol.������) = NVL(RsBody!������)
            Else
                .TextMatrix(lngRow, mBodyCol.������) = NVL(RsBody!�����)
            End If
            .TextMatrix(lngRow, mBodyCol.�ɲ���) = NVL(RsBody!�ɲ���)
            .TextMatrix(lngRow, mBodyCol.��¼״̬) = NVL(RsBody!��¼״̬)
            
            If tabShow.Tab <> 0 Then
                .TextMatrix(lngRow, mBodyCol.����ʱ��) = NVL(RsBody!����ʱ��)
                .RowData(lngRow) = NVL(RsBody!Id, 0)
                SetGRDCOLOR mshBody, lngRow, IIf(NVL(RsBody!�ɲ���) = 1, 1, NVL(RsBody!��¼״̬, 0))
            Else
                .TextMatrix(lngRow, mBodyCol.����ʱ��) = NVL(RsBody!��������)
                SetGRDCOLOR mshBody, lngRow, 1
            End If
            lngRow = lngRow + 1
            RsBody.MoveNext
        Loop
        .Redraw = True
    End With
    LoadBodyData = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub tabShow_Click(PreviousTab As Integer)
    Dim blnYes As Boolean
    If PreviousTab = tabShow.Tab Then Exit Sub
    
    If mblnExit = True Then
        mblnExit = False
        Exit Sub
    End If
    
    If mstrSelCon <> "" And PreviousTab = 0 Then
        ShowMsgBox "�Ѿ��б�ѡ�����Ŀ,���Ƿ�ϣ���ı�ѡ��!", True, blnYes
        If Not blnYes Then
            mblnExit = True
            tabShow.Tab = PreviousTab
            mnuEditPayType.Checked = True
            mnuEditBackType.Checked = False
            Exit Sub
        End If
        mstrSelCon = ""
    
    End If
    If mlngCountSel > 0 And PreviousTab = 1 Then
        ShowMsgBox "�Ѿ��б�ѡ�����Ŀ,���Ƿ�ϣ���ı�ѡ��!", True, blnYes
        If Not blnYes Then
            mblnExit = True
            tabShow.Tab = PreviousTab
            mnuEditPayType.Checked = False
            mnuEditBackType.Checked = True
            Exit Sub
        End If
        mlngCountSel = 0
    End If
    
    mblnExit = False
    
    mlngCountSel = 0
    '�޸�:���˺�   Bug:    ����:2008-05-14 11:17:53
    If PreviousTab = 0 Then
        'δ��
        Select Case tbsSel.SelectedItem.Key
        Case "K1"   '��ϸ
            mshBody.Tag = "_��ϸ"
        Case "K2"  'ѡ����ϸ
            mshBody.Tag = "_ѡ��"
        Case "K3"  '���ܷ���
            mshBody.Tag = "_����"
        End Select
        SaveFlexState mshHead, Me.Caption
    Else
        '�ѷ�
        '�ȱ���
        mshBody.Tag = "_�ѷ�"
    End If
    If PreviousTab <> -1 Then
        SaveFlexState mshBody, Me.Caption
    End If
    
    Select Case tabShow.Tab
        Case 0 '--δ�����嵥
            SetMnuEnable
            Form_Resize
            Call GetHeadData(0)
            If Me.mshHead.Enabled Then mshHead.SetFocus
            mshHead_EnterCell
            mnuEditPayType.Checked = True
            mnuEditBackType.Checked = False
        
        Case 1  '--�ѷ����嵥
            SetMnuEnable
            Form_Resize
            Call ReadBillData
            If Me.mshBody.Enabled Then mshBody.SetFocus
            mnuEditPayType.Checked = False
            mnuEditBackType.Checked = True
                
    End Select
    SetMnuEnable
    mstrPreSelKey = ""
End Sub

Private Sub tbsSel_Click()
    'δ��
    If mstrPreSelKey <> "" Then
        Select Case mstrPreSelKey
        Case "K1"   '��ϸ
            mshBody.Tag = "_��ϸ"
        Case "K2"  'ѡ����ϸ
            mshBody.Tag = "_ѡ��"
        Case "K3"  '���ܷ���
            mshBody.Tag = "_����"
        End Select
        SaveFlexState mshBody, Me.Caption
    End If
    
    mstrPreSelKey = tbsSel.SelectedItem.Key
    ReadBillData
End Sub
Private Sub SetGrdSelBackColor(objGrid As Object)
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:����ѡ��ɫ�û�
    '--�����:
    '--������:
    '--��  ��:
    '-----------------------------------------------------------------------------------------------------------
    If objGrid Is mshHead Then
        mshHead.BackColorSel = &H8000000C     ' &HC0C0C0
        mshBody.BackColorSel = &HE0E0E0
    ElseIf objGrid Is mshBody Then
        mshHead.BackColorSel = &HE0E0E0
        mshBody.BackColorSel = &H8000000C     '&HC0C0C0
    End If
End Sub

Private Function CheckStock() As Boolean
    Dim dblStock As Double
    Dim rsTemp As New ADODB.Recordset
    Dim lngRow As Long
    Dim lng����ID As Long
    Dim lng���� As Long
    Dim RsBody As New ADODB.Recordset
    
    
    '�����
    If mintCheckStock = 0 Then CheckStock = True: Exit Function
    
    Set RsBody = GetBodyQurysSQL("K3")
    'zlDatabase.OpenRecordset RsBody, gstrSQL, "�����"
    
    CheckStock = False
    With RsBody
            Do While Not .EOF
                lng����ID = NVL(!����ID, 0)
                lng���� = NVL(!����, 0)
                
               If lng����ID <> 0 Then
                        dblStock = NVL(!�����, 0)
                        
                        If dblStock < NVL(!����, 0) Then
                            If lng���� <> 0 Then
                                MsgBox NVL(!��������) & "�����ο�������������ܼ������ϣ�", vbInformation, gstrSysName: Exit Function
                            Else
                                Select Case mintCheckStock
                                Case 1
                                    If MsgBox(NVL(!��������) & "�Ŀ�����������Ƿ�������ϣ�", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                                Case 2
                                    MsgBox NVL(!��������) & "�Ŀ�������������ܼ������ϣ�", vbInformation, gstrSysName: Exit Function
                                End Select
                            End If
                        End If
               End If
               .MoveNext
            Loop
    End With
    CheckStock = True
End Function

Private Function SendBill() As Boolean
    Dim intRow As Integer
    Dim strDate As String
    Dim mlng���ϲ���ID As Long
    Dim strNo As String
    Dim int���Ϸ�ʽ As Integer     '1-��������;2-��������;3-���ŷ���
    Dim lng���� As Long
    
    On Error GoTo ErrHand
    err = 0
    SendBill = False
    
    strDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    mlng���ϲ���ID = cboStock.ItemData(cboStock.ListIndex)
    gcnOracle.BeginTrans
    With mshHead
        int���Ϸ�ʽ = UBound(Split(mstrSelCon, "||"))
        
        int���Ϸ�ʽ = IIf(int���Ϸ�ʽ = 0 Or int���Ϸ�ʽ = 1, 1, 2)
        
        If InStr(1, mstrSelCon, ":26") <> 0 Then
            '���ŷ���
            int���Ϸ�ʽ = 3
        End If
            
        For intRow = 1 To .Rows - 1
            lng���� = Decode(.TextMatrix(intRow, mHeadCol.����), "�շ�", 24, "���ʵ�", 25, "���ʱ�", 26, 0)
            strNo = Trim(.TextMatrix(intRow, mHeadCol.NO))
             If lng���� <> 0 And strNo <> "" And Trim(.TextMatrix(intRow, mHeadCol.��־)) = "��" Then
                mstrPrintCon = strNo & "||" & lng���� & "||" & mlng���ϲ���ID
                '���̲���:�ⷿID_IN,����_IN,NO_IN,�����_IN,������_IN,У����_IN,���Ϸ�ʽ_IN,�������_IN
                gstrSQL = "zl_�����շ���¼_��������(" & _
                    mlng���ϲ���ID & "," & _
                    lng���� & ",'" & _
                    strNo & "','" & _
                    gstrUserName & "','" & _
                    gstrUserName & "','NULL'," & _
                    int���Ϸ�ʽ & ",to_date('" & _
                    strDate & "','yyyy-MM-dd hh24:mi:ss'))"
                Call zlDatabase.ExecuteProcedure(gstrSQL, (Me.Caption & "-�������Ϸ���"))
            End If
        Next
    End With
    
    gcnOracle.CommitTrans
    BillListPrint int���Ϸ�ʽ, strDate
  
    SendBill = True
    Exit Function
ErrHand:
    gcnOracle.RollbackTrans
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function
Private Sub SetGRDCOLOR(ByVal objGrd As Object, ByVal lngRow As Long, ByVal int��¼״̬ As Integer)
    Dim lngColor As Long
    Dim i As Long
    If int��¼״̬ = 1 Then
        lngColor = &H80000008
    ElseIf zlCommFun.ZyMod(int��¼״̬, 3) = 2 Then
         lngColor = vbRed
    Else
        lngColor = vbBlue
    End If
    With mshBody
        For i = 0 To .Cols - 1
            .Row = lngRow
            .Col = i
            .CellForeColor = lngColor
        Next
    End With
End Sub

Private Function LoadInIcon() As Boolean
    '--Ϊ���ؼ�װ��ͼ��--
    On Error Resume Next
    err = 0
    LoadInIcon = False
    
    '������
    With ImgTbarBlack
        .ImageHeight = 24
        .ImageWidth = 24
        .ListImages.Add , , LoadResPicture("BPREVIEW", vbResIcon)
        .ListImages.Add , , LoadResPicture("BPRINT", vbResIcon)
        .ListImages.Add , , LoadResPicture("BSTOP", vbResIcon)
        .ListImages.Add , , LoadResPicture("BSTART", vbResIcon)
       ' .ListImages.Add , , LoadResPicture("BSEND", vbResIcon)
       ' .ListImages.Add , , LoadResPicture("BSEND", vbResIcon)
        .ListImages.Add , , LoadResPicture("BFILTER", vbResIcon)
        .ListImages.Add , , LoadResPicture("BHELP", vbResIcon)
        .ListImages.Add , , LoadResPicture("BEXIT", vbResIcon)
    End With
    
    With ImgTbarColor
        .ImageHeight = 24
        .ImageWidth = 24
        .ListImages.Add , , LoadResPicture("CPREVIEW", vbResIcon)
        .ListImages.Add , , LoadResPicture("CPRINT", vbResIcon)
        .ListImages.Add , , LoadResPicture("CSTOP", vbResIcon)
        .ListImages.Add , , LoadResPicture("CSTART", vbResIcon)
'        .ListImages.Add , , LoadResPicture("CSEND", vbResIcon)
'        .ListImages.Add , , LoadResPicture("CSEND", vbResIcon)
        .ListImages.Add , , LoadResPicture("CFILTER", vbResIcon)
        .ListImages.Add , , LoadResPicture("CHELP", vbResIcon)
        .ListImages.Add , , LoadResPicture("CEXIT", vbResIcon)
    End With
    
    With tlbThis
        Set .ImageList = ImgTbarBlack
        Set .HotImageList = ImgTbarColor
        
        .Buttons("Ԥ��").Image = 1
        .Buttons("��ӡ").Image = 2
        .Buttons("����").Image = 3
        .Buttons("����").Image = 4
        .Buttons("����").Image = 5
        
        .Buttons("����").Image = 6
        .Buttons("�˳�").Image = 7
    End With
    
    cbrThis.Bands(1).MinHeight = tlbThis.Height
    If err <> 0 Then
        MsgBox "�����Դ�ļ���ʧ�����������������ϵ��", vbInformation, gstrSysName
        Exit Function
    End If
    LoadInIcon = True
End Function


Private Sub mnuHelpAbout_Click()
    '����
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuHelpTitle_Click()
    '��������
    Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int(glngSys / 100))
End Sub

Private Sub mnuHelpWebHome_Click()
    '������ҳ
    Call zlHomePage(Me.hwnd)
End Sub

Private Sub mnuHelpWebMail_Click()
    '���ͷ���
    Call zlMailTo(Me.hwnd)
End Sub
Private Sub subPrint(bytMode As Byte)
    '����:���д�ӡ,Ԥ���������EXCEL
    '����:bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    
    Dim objPrint As Object
    Dim objRow As New zlTabAppRow
    Dim strRange As String
    
    If Me.tabShow.Tab = 1 Then
        strRange = "������� " & Format(mstrStartDate, "yyyy��MM��dd��") & "��" & Format(mstrEndDate, "yyyy��MM��dd��")
    Else
        strRange = "�������� " & Format(mstrStartDate, "yyyy��MM��dd��") & "��" & Format(mstrEndDate, "yyyy��MM��dd��")
    End If
    
    Set objPrint = New zlPrint1Grd
    objPrint.Title.Font.Name = "����_GB2312"
    objPrint.Title.Font.Size = 18
    objPrint.Title.Font.Bold = True
    
    objPrint.Title.Text = IIf(Me.tabShow.Tab = 0, "δ�������", "�ѷ������")
        
    objRow.Add "ʱ�䣺" & strRange
    objRow.Add "���ϲ��ţ�" & cboStock.Text
    objPrint.UnderAppRows.Add objRow
    Set objRow = New zlTabAppRow
        
    objRow.Add "��ӡ��:" & UserInfo.�û���
    objRow.Add "��ӡ����:" & Format(zlDatabase.Currentdate, "yyyy��MM��dd��")
    objPrint.BelowAppRows.Add objRow
    
    If Me.ActiveControl Is mshBody Then
        Set objPrint.Body = mshBody
    Else
        Set objPrint.Body = IIf(Me.tabShow.Tab = 1, mshBody, mshHead)
    End If
    
    If bytMode = 1 Then
      Select Case zlPrintAsk(objPrint)
          Case 1
               zlPrintOrView1Grd objPrint, 1
          Case 2
              zlPrintOrView1Grd objPrint, 2
          Case 3
              zlPrintOrView1Grd objPrint, 3
      End Select
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
End Sub
Private Sub BillListPrint(Optional int���Ϸ�ʽ As Integer = 1, Optional strDate As String = "", Optional IntStyle As Integer = 0)
    '���ݻ�����ӡ
    '���Ϸ�ʽ:1-��������;2-��������;3-���ŷ���
    ' intStyle:0-�����Ϸ�ʽ��ӡ,1-���ݴ�ӡ
    Dim bln���ϵ� As Boolean
    Dim bln�ѷ����嵥 As Boolean
    Dim bln���ݴ�ӡ As Boolean
    
    bln���ϵ� = InStr(1, mstrPrivs, "����֪ͨ��") <> 0
    bln�ѷ����嵥 = InStr(1, gstrPrivs, "��ӡ�ѷ����嵥") <> 0
    bln���ݴ�ӡ = InStr(1, gstrPrivs, "���ݴ�ӡ") <> 0
    
    Select Case IntStyle
        Case 0
            If mintPrintPar = 0 Then
                '��ʾ��ӡ
                If mstrPrintCon <> "" And int���Ϸ�ʽ = 1 Then
                    If bln���ݴ�ӡ = False Then Exit Sub
                Else
                    If bln�ѷ����嵥 = False Then Exit Sub
                End If
                If MsgBox("����Ҫ��ӡ��ص�����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) <> vbYes Then Exit Sub
            ElseIf mintPrintPar = 1 Then
                '�Զ���ӡ
            Else
                Exit Sub
            End If
            Select Case int���Ϸ�ʽ
            Case 1  '������ӡ
                'mstrPrintCon
                'mstrPrintCon = strNo & "||" & lng���� & "||" & mlng���ϲ���id
                If mstrPrintCon <> "" Then
                    Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1723", Me, "�ⷿ==" & cboStock.ItemData(cboStock.ListIndex), "NO=" & Split(mstrPrintCon, "||")(0), "����=" & Split(mstrPrintCon, "||")(1), "�����=����� is not null", 2)
                Else
                    Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1723_1", Me, "�ⷿ=" & cboStock.ItemData(cboStock.ListIndex), "���Ϸ�ʽ=���ݷ���|1", "����ʱ��=" & strDate, 2)
                End If
            Case 2
                Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1723_1", Me, "�ⷿ=" & cboStock.ItemData(cboStock.ListIndex), "���Ϸ�ʽ=��������|2", "����ʱ��=" & strDate, 2)
            Case 3
                Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1723_1", Me, "�ⷿ=" & cboStock.ItemData(cboStock.ListIndex), "���Ϸ�ʽ=���ŷ���|3", "����ʱ��=" & strDate, 2)
            End Select
       Case 1
            '���ݴ�ӡ
            Dim strNo As String
            Dim int���� As Integer
            If bln���ݴ�ӡ = False Then Exit Sub
            
            strNo = Trim(mshHead.TextMatrix(mshHead.Row, mHeadCol.NO))
            int���� = Val(mshHead.TextMatrix(mshHead.Row, mHeadCol.����))
            
            Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1723", Me, "�ⷿ==" & cboStock.ItemData(cboStock.ListIndex), "NO=" & strNo, "����=" & int����, "�����=" & "����� is not null ", 2)
            
       Case 2
            '���ϵ���
            If bln���ϵ� = False Then Exit Sub
             Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1723_2", Me, "����ʱ��=" & strDate, "��λ=" & mintUnit + 1, 2)
    End Select
End Sub
Private Sub Ȩ�޿���()
    '
    Dim bln���� As Boolean
    Dim bln���� As Boolean
    Dim bln���� As Boolean
    Dim bln���ϵ� As Boolean
    Dim bln�ѷ����嵥 As Boolean
    Dim bln���ݴ�ӡ As Boolean
    
    bln���� = InStr(1, gstrPrivs, "�������Ϸ���") <> 0
    bln���� = InStr(1, gstrPrivs, "������������") <> 0
    bln���� = True ' InStr(1, gstrPrivs, "��������") <> 0
    
    bln���ϵ� = InStr(1, gstrPrivs, "����֪ͨ��") <> 0
    bln�ѷ����嵥 = InStr(1, gstrPrivs, "��ӡ�ѷ����嵥") <> 0
    bln���ݴ�ӡ = InStr(1, gstrPrivs, "���ݴ�ӡ") <> 0
    
    mnuFilePara.Visible = bln����
    mnuFile3.Visible = bln����
    mnuEditPay.Visible = bln����
    mnuEditPayCf.Visible = bln����
    mnuEditFpPay.Visible = bln����
    
    mnuEditOutPay.Visible = bln����
    mnuEditSplit0.Visible = bln���� Or bln����
    
    mnuFileBillprint.Visible = bln�ѷ����嵥 Or bln���ݴ�ӡ
    mnuFileRestore.Visible = bln���ϵ�
    mnuFile2.Visible = bln�ѷ����嵥 Or bln���ݴ�ӡ Or bln���ϵ�
    mnuEditStop.Visible = bln���� Or bln����
    mnuEditStopSp.Visible = mnuEditStop.Visible
    mnuEditStrict.Visible = bln����
    tlbThis.Buttons("����").Visible = bln����
    tlbThis.Buttons("����").Visible = bln����
    tlbThis.Buttons("EditSp").Visible = bln���� Or bln����
    
End Sub
Private Function AnalyseHistorySQL(ByVal strSQL As String, Optional strԭ�� As String = "", Optional str�ִ� As String = "") As String
    '������ʷ���ݵ�SQL���
    Dim strTemp As String
    strTemp = Replace(strSQL, "ҩƷ�շ���¼", "HҩƷ�շ���¼")
    strTemp = Replace(strTemp, "���˷��ü�¼", "H���˷��ü�¼")
    If strԭ�� <> "" Then
        strTemp = Replace(strTemp, strԭ��, str�ִ�)
    End If
    strTemp = strSQL & " Union ALL " & strTemp
    AnalyseHistorySQL = strTemp
End Function



Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '����:���ӵ�������̳
    '�޸���:���˺�
    '�޸�����:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hwnd)
End Sub

