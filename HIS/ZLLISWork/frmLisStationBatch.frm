VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmLisStationBatch 
   Caption         =   "������������"
   ClientHeight    =   5865
   ClientLeft      =   2775
   ClientTop       =   3765
   ClientWidth     =   8970
   Icon            =   "frmLisStationBatch.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   5865
   ScaleWidth      =   8970
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fra 
      Height          =   4845
      Left            =   0
      TabIndex        =   17
      Top             =   585
      Width           =   3120
      Begin VB.CheckBox chkQc 
         Caption         =   "����ʱ�����ʿر걾(&5)"
         Height          =   240
         Left            =   300
         TabIndex        =   10
         Top             =   2790
         Visible         =   0   'False
         Width           =   2865
      End
      Begin VB.CommandButton cmdReset 
         Caption         =   "��������(&J)"
         Height          =   350
         Left            =   1590
         TabIndex        =   12
         Top             =   4365
         Width           =   1185
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   0
         Left            =   300
         TabIndex        =   5
         Top             =   1020
         Width           =   2715
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   2
         Left            =   300
         TabIndex        =   7
         Top             =   1635
         Width           =   2715
      End
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   0
         Left            =   300
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   2235
         Width           =   2715
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "��������(&S)"
         Height          =   350
         Left            =   315
         TabIndex        =   11
         Top             =   4365
         Width           =   1185
      End
      Begin MSComCtl2.DTPicker dtp 
         Height          =   300
         Index           =   0
         Left            =   300
         TabIndex        =   1
         Top             =   420
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   68681731
         CurrentDate     =   38229
      End
      Begin MSComCtl2.DTPicker dtp 
         Height          =   300
         Index           =   1
         Left            =   1785
         TabIndex        =   3
         Top             =   420
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   68681731
         CurrentDate     =   38229
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "&4.��������"
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   2010
         Width           =   900
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "&1.�걾ʱ��"
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   0
         Top             =   195
         Width           =   900
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "&2.�걾����"
         Height          =   180
         Index           =   2
         Left            =   120
         TabIndex        =   4
         Top             =   795
         Width           =   900
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "&3.�� �� ��"
         Height          =   180
         Index           =   3
         Left            =   120
         TabIndex        =   6
         Top             =   1395
         Width           =   900
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "��"
         Height          =   180
         Index           =   4
         Left            =   1590
         TabIndex        =   2
         Top             =   480
         Width           =   180
      End
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&C"
      Height          =   350
      Index           =   5
      Left            =   405
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   3555
      Width           =   1100
   End
   Begin MSComctlLib.ImageList ilsMenu 
      Left            =   6570
      Top             =   675
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLisStationBatch.frx":0E42
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLisStationBatch.frx":15BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLisStationBatch.frx":1D36
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLisStationBatch.frx":1F50
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLisStationBatch.frx":2170
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLisStationBatch.frx":2390
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLisStationBatch.frx":25AA
            Key             =   "Qc"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsHotMenu 
      Left            =   7245
      Top             =   690
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLisStationBatch.frx":27C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLisStationBatch.frx":2F3E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLisStationBatch.frx":36B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLisStationBatch.frx":38D2
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLisStationBatch.frx":3AF2
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLisStationBatch.frx":3D12
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLisStationBatch.frx":3F2C
            Key             =   "Qc"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   16
      Top             =   5505
      Width           =   8970
      _ExtentX        =   15822
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmLisStationBatch.frx":4146
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10742
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
   Begin VSFlex8Ctl.VSFlexGrid vsf 
      Height          =   2685
      Left            =   3420
      TabIndex        =   13
      Top             =   1065
      Width           =   3600
      _cx             =   6350
      _cy             =   4736
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
      BackColorSel    =   16768667
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483632
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483639
      FocusRect       =   2
      HighLight       =   0
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   240
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
      AutoSizeMode    =   1
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
      Editable        =   2
      ShowComboButton =   1
      WordWrap        =   -1  'True
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
      WallPaperAlignment=   8
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin ComCtl3.CoolBar cbrThis 
      Align           =   1  'Align Top
      Height          =   705
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   8970
      _ExtentX        =   15822
      _ExtentY        =   1244
      BandCount       =   1
      _CBWidth        =   8970
      _CBHeight       =   705
      _Version        =   "6.7.9782"
      Child1          =   "tbrThis"
      MinHeight1      =   645
      Width1          =   8370
      Key1            =   "only"
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tbrThis 
         Height          =   645
         Left            =   30
         TabIndex        =   15
         Top             =   30
         Width           =   8850
         _ExtentX        =   15610
         _ExtentY        =   1138
         ButtonWidth     =   1296
         ButtonHeight    =   1138
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "ilsMenu"
         HotImageList    =   "ilsHotMenu"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   8
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&A.ȫѡ"
               Key             =   "ȫѡ"
               Object.ToolTipText     =   "ȫѡ"
               Object.Tag             =   "&A.ȫѡ"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&R.ȫ��"
               Key             =   "ȫ��"
               Object.ToolTipText     =   "ȫ��"
               Object.Tag             =   "&R.ȫ��"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split_1"
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&D.ɾ��"
               Key             =   "ɾ��"
               Object.ToolTipText     =   "ɾ��ѡ���������걾"
               Object.Tag             =   "&D.ɾ��"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Caption         =   "&C.�ʿ�"
               Key             =   "�ʿ�"
               Object.ToolTipText     =   "��ѡ���������걾תΪ�ʿ�"
               Object.Tag             =   "&C.�ʿ�"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split_5"
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&H.����"
               Key             =   "����"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "&H.����"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&E.�˳�"
               Key             =   "�˳�"
               Object.ToolTipText     =   "�˳�"
               Object.Tag             =   "&E.�˳�"
               ImageIndex      =   5
            EndProperty
         EndProperty
      End
   End
   Begin VB.CommandButton cmd 
      Cancel          =   -1  'True
      Caption         =   "&E"
      Height          =   350
      Index           =   4
      Left            =   405
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   3210
      Width           =   1100
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&H"
      Height          =   350
      Index           =   3
      Left            =   405
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   2850
      Width           =   1100
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&D"
      Height          =   350
      Index           =   2
      Left            =   405
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   2505
      Width           =   1100
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&R"
      Height          =   350
      Index           =   1
      Left            =   405
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   2160
      Width           =   1100
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&A"
      Height          =   350
      Index           =   0
      Left            =   405
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   1800
      Width           =   1100
   End
End
Attribute VB_Name = "frmLisStationBatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������

Private mblnOK As Boolean
Private mblnStartUp As Boolean
Private mfrmMain As Form
Private mlngLoop As Long
Private mRs As New ADODB.Recordset
Private mstrSQL As String
Private mblnChangeEdit As Boolean
Private lngExecDept As Long '��ǰִ�в��ţ����ϼ�����

Private Enum mCol
    ѡ�� = 0
    �ʿ�Ʒ
    �걾��
    �걾����
    ����ʱ��
    ������
    ��������
    ת��
End Enum

Private Sub AdjustEnableState()
    '-----------------------------------------------------------------------------------------
    '����:�����޸�״̬���ð�ť���˵��ȵĿ���״̬
    '-----------------------------------------------------------------------------------------
    cmd(2).Enabled = True
        
    If mblnChangeEdit = False Then cmd(2).Enabled = False
        
    tbrThis.Buttons("ɾ��").Enabled = cmd(2).Enabled
    tbrThis.Buttons("�ʿ�").Enabled = cmd(2).Enabled
End Sub

Private Sub RefreshStatus()
    '-----------------------------------------------------------------------------------------
    '����:
    '-----------------------------------------------------------------------------------------
    If Vsf.Rows = 2 And Trim(Vsf.TextMatrix(1, 1)) = "" Then
        stbThis.Panels(2).Text = "û�б걾��Ϣ��"
    Else
        stbThis.Panels(2).Text = "���ҵ� " & Vsf.Rows - 1 & " ���걾��Ϣ��"
    End If
    
End Sub

Public Function ShowEdit(ByVal frmMain As Form, Optional ByVal ExecDeptID As Long = 0) As Boolean
    '------------------------------------------------------------------------------------------------------
    '���ܣ���ʾ���༭����
    '������
    '���أ�
    '------------------------------------------------------------------------------------------------------
    mblnStartUp = True
    mblnOK = False
        
    Set mfrmMain = frmMain
    lngExecDept = ExecDeptID
    
    If InitData = False Then Exit Function
                    
    mblnChangeEdit = False
    Call AdjustEnableState
    
    Me.Show 1, frmMain
    
    ShowEdit = mblnOK
    
End Function

Private Function InitData() As Boolean
    '------------------------------------------------------------------------------------------------------
    '���ܣ�
    '------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    
    On Error GoTo ErrHand
    
    Vsf.Cols = 0
    Call NewColumn(Vsf, "ѡ��", 510, 4)
    Call NewColumn(Vsf, "�ʿ�Ʒ", 510, 1)
    Call NewColumn(Vsf, "�걾��", 750, 1)
    Call NewColumn(Vsf, "�걾����", 900, 1)
    Call NewColumn(Vsf, "����ʱ��", 1080, 1)
    Call NewColumn(Vsf, "������", 750, 1)
    Call NewColumn(Vsf, "��������", 1200, 1)
    Call NewColumn(Vsf, "ת��", 0, 1)
    Vsf.ColDataType(mCol.ѡ��) = flexDTBoolean
    
        
    InitData = True
    
    Exit Function
    
ErrHand:
    
    If ErrCenter = 1 Then Resume
    
End Function

Private Function ReadData() As Boolean
    '------------------------------------------------------------------------------------------------------
    '���ܣ�
    '------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    Dim strWhere As String
    Dim varTmp2 As Variant
    Dim strTmp As String
    Dim lngLoop As Long
    Dim blnMoved As Boolean, strSQLBak As String, strOrderBy As String
    Dim varItem As Variant                          '�ֽ�","��
    Dim varBetween As Variant                       '�ֽ�"~"
    
    On Error GoTo ErrHand
    
    Vsf.Rows = 2
    Vsf.RowData(1) = 0
    Vsf.Cell(flexcpText, 1, 0, 1, Vsf.Cols - 1) = ""
        
    strWhere = " AND A.����ʱ�� BETWEEN TO_DATE('" & Format(dtp(0).Value, dtp(0).CustomFormat) & " 00:00:00', 'yyyy-mm-dd hh24:mi:ss') AND TO_DATE('" & Format(dtp(1).Value, dtp(1).CustomFormat) & " 23:59:59', 'yyyy-mm-dd hh24:mi:ss') "
    If chkQc.Value <> 1 Then
'        strWhere = strWhere & " And Not (A.�Ƿ��ʿ�Ʒ=1 Or Instr(','||D.�ʿر걾��||',',','||A.�걾���||',')>0)"
    End If
    If lngExecDept > 0 Then strWhere = strWhere & " And A.ִ�п���ID+0=" & lngExecDept
    blnMoved = MovedByDate(dtp(0).Value)
    
    If Trim(txt(2).Text) <> "" Then strWhere = strWhere & " AND A.������ = '" & Trim(txt(2).Text) & "'"
'    If cbo(1).ListIndex > 0 Then strWhere = strWhere & " AND B.��������ID + 0 = " & cbo(1).ItemData(cbo(1).ListIndex)
    If cbo(0).ListIndex > 0 Then strWhere = strWhere & _
        IIf(cbo(0).ListIndex = 1, " AND A.����id IS Null", " AND A.����id=" & cbo(0).ItemData(cbo(0).ListIndex))
        
    If Trim(txt(0).Text) <> "" Then
        
'        varTmp2 = Split(Trim(txt(0).Text), ",")
'        strTmp = " 1=2 "
'        For mlngLoop = 0 To UBound(varTmp2)
'            If InStr(varTmp2(mlngLoop), "-") = 0 Then
'                strTmp = strTmp & "  OR A.�걾���='" & varTmp2(mlngLoop) & "'"
'            Else
''                strTmp = strTmp & "  OR A.�걾��� BETWEEN '" & Mid(varTmp2(mlngLoop), 1, InStr(varTmp2(mlngLoop), "-") - 1) & "' AND '" & Mid(varTmp2(mlngLoop), InStr(varTmp2(mlngLoop), "-") + 1) & "'"
'                strTmp = strTmp & "  OR A.�걾��� IN (''"
'                For lngLoop = Val(Mid(varTmp2(mlngLoop), 1, InStr(varTmp2(mlngLoop), "-") - 1)) To Val(Mid(varTmp2(mlngLoop), InStr(varTmp2(mlngLoop), "-") + 1))
'                    strTmp = strTmp & ",'" & lngLoop & "'"
'                Next
'                strTmp = strTmp & ")"
'            End If
'        Next
'        If strTmp <> " 1=2 " Then strWhere = strWhere & " AND (" & strTmp & ")"
        
        'strWhere = strWhere & " AND A.�걾��� BETWEEN '" & txt(0).Text & "' AND '" & txt(0).Text & "'"
        varItem = Split(Trim(txt(0).Text), ",")
        For mlngLoop = 0 To UBound(varItem)
            varBetween = Split(varItem(mlngLoop), "~")
            If UBound(varBetween) > 0 Then
                strTmp = strTmp & "  OR A.�걾��� BETWEEN " & TransSampleNO(varBetween(0)) & " AND " & TransSampleNO(varBetween(1))
            Else
                strTmp = strTmp & " OR A.�걾���=" & TransSampleNO(varItem(mlngLoop))
            End If
        Next
        If strTmp <> "" Then strWhere = strWhere & " AND (1=2 " & strTmp & ")"
    End If
    
    strOrderBy = " ORDER BY ��������,����ʱ��,�걾��"
            
    mstrSQL = "select DISTINCT A.ID,0 AS ѡ��,Decode(�Ƿ��ʿ�Ʒ,1,'��',Decode(Instr(','||D.�ʿر걾��||',',','||A.�걾���||','),0,'','��')) As �ʿ�Ʒ, " & _
                      " Decode(A.����id, Null, " & vbCrLf & _
                        " to_Char(Trunc(A.�걾���/10000)+1,'0000')|| '-'||to_Char(MOD(A.�걾���,10000),'0000'), A.�걾���) As �걾��, " & _
                      "A.�걾����," & _
                      "TO_CHAR(A.����ʱ��,'MM-DD HH24:MI') AS ����ʱ��," & _
                      "A.������," & _
                      "A.������," & _
                      "D.���� AS ��������,0 As ת�� " & _
                 "from ����걾��¼ A, �������� D " & _
                "WHERE A.����ID = D.ID(+) And A.ҽ��ID Is Null " & strWhere
    If blnMoved Then
        strSQLBak = mstrSQL
        strSQLBak = Replace(strSQLBak, "0 As ת��", "1 As ת��")
        strSQLBak = Replace(strSQLBak, "����걾��¼", "H����걾��¼")
        mstrSQL = mstrSQL & " Union ALL " & strSQLBak
    End If
    mstrSQL = mstrSQL & strOrderBy

    Call OpenRecord(rs, mstrSQL, Me.Caption)
    If rs.BOF = False Then
        Call FillGrid(Vsf, rs)
    End If
    
    ReadData = True
    
    Exit Function
    
ErrHand:
    
    If ErrCenter = 1 Then Resume
    
End Function

Private Function SaveData(ByVal SaveMode As Integer) As Boolean
    '------------------------------------------------------------------------------------------------------
    '���ܣ�
    '------------------------------------------------------------------------------------------------------
    Dim strReportCode As String
    Dim strReportParaNo As String
    Dim bytReportParaMode As Byte
    
    On Error GoTo ErrHand
    
    gcnOracle.BeginTrans
    For mlngLoop = 1 To Vsf.Rows - 1
        If Abs(Val(Vsf.TextMatrix(mlngLoop, mCol.ѡ��))) = 1 And Val(Vsf.RowData(mlngLoop)) > 0 Then
            If SaveMode = 0 Then 'ɾ���걾
                ExecuteProc "ZL_����걾��¼_�걾ɾ��(" & Val(Vsf.RowData(mlngLoop)) & ")", Me.Caption
            Else
                ExecuteProc "ZL_����걾��¼_�걾�ʿ�(" & Val(Vsf.RowData(mlngLoop)) & ")", Me.Caption
            End If
        End If
    Next
    gcnOracle.CommitTrans

    SaveData = True
    ReadData
    Exit Function
ErrHand:
    gcnOracle.RollbackTrans
    SaveData = False
    Call ErrCenter
End Function


Private Sub cbo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlcommfun.PressKey vbKeyTab
End Sub

Private Sub cmd_Click(Index As Integer)
    
    Select Case Index
    Case 0
        For mlngLoop = 1 To Vsf.Rows - 1
            If Val(Vsf.RowData(mlngLoop)) > 0 Then
                Vsf.TextMatrix(mlngLoop, mCol.ѡ��) = 1
            End If
        Next
        
        mblnChangeEdit = True
        Call AdjustEnableState
    Case 1
        For mlngLoop = 1 To Vsf.Rows - 1
            Vsf.TextMatrix(mlngLoop, mCol.ѡ��) = 0
        Next
        mblnChangeEdit = False
        Call AdjustEnableState
    Case 2
        If mblnChangeEdit Then
            If MsgBox("ȷ��Ҫɾ��ѡ���������걾��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub

            If SaveData(0) = False Then Exit Sub
            
            mblnOK = True
            
            mblnChangeEdit = False
            Call AdjustEnableState

'            Unload Me
            Exit Sub
        End If
        
    Case 3
        ShowHelp App.ProductName, Me.Hwnd, Me.Name, Int((glngSys) / 100)
    Case 4
        Unload Me
    Case 5
        If mblnChangeEdit Then
            If MsgBox("ȷ��Ҫ��ѡ���������걾תΪ�ʿ���", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub

            If SaveData(1) = False Then Exit Sub
            
            mblnOK = True
            
            mblnChangeEdit = False
            Call AdjustEnableState

'            Unload Me
            Exit Sub
        End If
    End Select
End Sub

Private Sub cmdRefresh_Click()
    
    Call ReadData
    
    mblnChangeEdit = False
    Call AdjustEnableState
    Call RefreshStatus
    
    Vsf.Col = 1
    Vsf.SetFocus
    Vsf.Col = 0
End Sub

Private Sub cmdReset_Click()
    
    dtp(0).Value = Format(zldatabase.Currentdate, "yyyy-mm-dd")
    dtp(1).Value = Format(zldatabase.Currentdate, "yyyy-mm-dd")
    
    txt(0).Text = ""
    txt(2).Text = ""
    
    dtp(0).SetFocus
End Sub

Private Sub dtp_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlcommfun.PressKey vbKeyTab
End Sub

Private Sub Form_Activate()
    Dim rs As New ADODB.Recordset
    Dim lngDefaultDev As Long
    
    If mblnStartUp = False Then Exit Sub
    mblnStartUp = False
    
    '��������
    mstrSQL = "SELECT A.����||'-'||A.����,ID FROM �������� A ORDER BY A.����||'-'||A.����"
    Call OpenRecord(rs, mstrSQL, Me.Caption)
    cbo(0).AddItem "��������"
    cbo(0).AddItem "�ֹ�"
    If rs.BOF = False Then Call AddComboData(cbo(0), rs, False)
    lngDefaultDev = Val(Split(GetConnectDevs & ";1", ";")(0))
    cbo(0).ListIndex = FindComboItem(cbo(0), lngDefaultDev)
    If cbo(0).ListCount > 0 And cbo(0).ListIndex = -1 Then cbo(0).ListIndex = 0
    
    Call cmdReset_Click
    
End Sub

Private Sub Form_Load()
    
    Call RestoreWinState(Me, App.ProductName)
        
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    With fra
        .Left = 0
        .Top = cbrthis.Height - 90
        .Height = Me.ScaleHeight - .Top - stbThis.Height
    End With
    
    With Vsf
        .Left = fra.Left + fra.Width
        .Top = cbrthis.Height
        .Width = Me.ScaleWidth - .Left
        .Height = Me.ScaleHeight - stbThis.Height - .Top
    End With
    
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
    If mblnChangeEdit Then
        Cancel = (MsgBox("�������޸ĵ����ݱ��뱣������Ч���Ƿ񲻱�����˳���", vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) = vbNo)
        If Cancel Then Exit Sub
    End If
    
    Call SaveWinState(Me, App.ProductName)
End Sub



Private Sub tbrThis_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
    Case "ȫѡ"
        Call cmd_Click(0)
    Case "ȫ��"
        Call cmd_Click(1)
    Case "ɾ��"
        Call cmd_Click(2)
    Case "�ʿ�"
        Call cmd_Click(5)
    Case "����"
        Call cmd_Click(3)
    Case "�˳�"
        Call cmd_Click(4)
    End Select
End Sub

Private Sub txt_GotFocus(Index As Integer)
    If Index = 2 Then zlcommfun.OpenIme True
    
    zlControl.TxtSelAll txt(Index)
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    
    If Chr(KeyAscii) = "'" Then KeyAscii = 0
    
    If KeyAscii = vbKeyReturn Then
        zlcommfun.PressKey vbKeyTab
    Else
    
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Select Case Index
        Case 0
            KeyAscii = FilterKeyAscii(KeyAscii, 99, "ZXCVBNMASDFGHJKLQWERTYUIOP01234567890,-~")
        End Select
    End If
        
End Sub

Private Sub txt_LostFocus(Index As Integer)
    If Index = 2 Then zlcommfun.OpenIme False
End Sub

Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
    Cancel = Not StrIsValid(txt(Index).Text, txt(Index).MaxLength)
End Sub

Private Sub vsf_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    mblnChangeEdit = True
    Call AdjustEnableState
End Sub

Private Sub vsf_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    On Error Resume Next
    
    If NewRow + 1 > Vsf.FixedRows And OldRow + 1 > Vsf.FixedRows Then
        Vsf.Cell(flexcpBackColor, OldRow, 0, OldRow, Vsf.Cols - 1) = Vsf.BackColor
        Vsf.Cell(flexcpBackColor, NewRow, 0, NewRow, Vsf.Cols - 1) = Vsf.BackColorSel
    End If
    
End Sub

Private Sub vsf_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = 0 Then Cancel = True
End Sub

Private Sub vsf_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Val(Vsf.RowData(Row)) = 0 Then Cancel = True
    If Col <> 0 Then Cancel = True
    
End Sub
