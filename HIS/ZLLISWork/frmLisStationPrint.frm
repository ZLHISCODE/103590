VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmLisStationPrint 
   Caption         =   "�����ӡ"
   ClientHeight    =   5865
   ClientLeft      =   2775
   ClientTop       =   3765
   ClientWidth     =   8970
   Icon            =   "frmLisStationPrint.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5865
   ScaleWidth      =   8970
   ShowInTaskbar   =   0   'False
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
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLisStationPrint.frx":0E42
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLisStationPrint.frx":15BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLisStationPrint.frx":1D36
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLisStationPrint.frx":1F50
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLisStationPrint.frx":2170
            Key             =   "Quit"
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
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLisStationPrint.frx":2390
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLisStationPrint.frx":2B0A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLisStationPrint.frx":3284
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLisStationPrint.frx":349E
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLisStationPrint.frx":36BE
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   21
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
            Picture         =   "frmLisStationPrint.frx":38DE
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
      TabIndex        =   18
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
      TabIndex        =   19
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
         TabIndex        =   20
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
            NumButtons      =   7
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
               Caption         =   "&P.��ӡ"
               Key             =   "��ӡ"
               Object.ToolTipText     =   "��ӡ"
               Object.Tag             =   "&P.��ӡ"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split_5"
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&H.����"
               Key             =   "����"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "&H.����"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&E.�˳�"
               Key             =   "�˳�"
               Object.ToolTipText     =   "�˳�"
               Object.Tag             =   "&E.�˳�"
               ImageIndex      =   5
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame fra 
      Height          =   4845
      Left            =   105
      TabIndex        =   22
      Top             =   570
      Width           =   3120
      Begin VB.CommandButton cmdReset 
         Caption         =   "��������(&J)"
         Height          =   350
         Left            =   1590
         TabIndex        =   17
         Top             =   4365
         Width           =   1185
      End
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   3
         Left            =   300
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   4020
         Width           =   2715
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   0
         Left            =   300
         TabIndex        =   5
         ToolTipText     =   "�걾����,�ָ����ԡ�ָ����Χ"
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
         TabIndex        =   11
         Top             =   2820
         Width           =   2715
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "��������(&S)"
         Height          =   350
         Left            =   300
         TabIndex        =   16
         Top             =   4365
         Width           =   1185
      End
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   1
         Left            =   300
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   2220
         Width           =   2715
      End
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   2
         Left            =   300
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   3420
         Width           =   2715
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
         Caption         =   "&7.��ӡ˳��"
         Height          =   180
         Index           =   5
         Left            =   120
         TabIndex        =   14
         Top             =   3795
         Width           =   900
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "&5.��������"
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   2595
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
         Caption         =   "&4.�������"
         Height          =   180
         Index           =   6
         Left            =   120
         TabIndex        =   8
         Top             =   2010
         Width           =   900
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "&6.ִ�п���"
         Height          =   180
         Index           =   7
         Left            =   120
         TabIndex        =   12
         Top             =   3195
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
      Cancel          =   -1  'True
      Caption         =   "&E"
      Height          =   350
      Index           =   4
      Left            =   405
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   3210
      Width           =   1100
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&H"
      Height          =   350
      Index           =   3
      Left            =   405
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   2850
      Width           =   1100
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&P"
      Height          =   350
      Index           =   2
      Left            =   405
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   2505
      Width           =   1100
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&R"
      Height          =   350
      Index           =   1
      Left            =   405
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   2160
      Width           =   1100
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&A"
      Height          =   350
      Index           =   0
      Left            =   405
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   1800
      Width           =   1100
   End
End
Attribute VB_Name = "frmLisStationPrint"
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
    ����
    �걾��
    �걾����
    ����ʱ��
    ������
    ������
    ����ʱ��
    ������
    �������
    ��������
    ִ�п���
    ҽ��id
    ���ͺ�
    ת��
    �걾ID
    ����ID
    �Ƿ����
End Enum

Private Sub AdjustEnableState()
    '-----------------------------------------------------------------------------------------
    '����:�����޸�״̬���ð�ť���˵��ȵĿ���״̬
    '-----------------------------------------------------------------------------------------
    cmd(2).Enabled = True
        
    If mblnChangeEdit = False Then cmd(2).Enabled = False
        
    tbrThis.Buttons("��ӡ").Enabled = cmd(2).Enabled
        
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
    Call NewColumn(Vsf, "����", 500, 1)
    Call NewColumn(Vsf, "�걾��", 750, 1)
    Call NewColumn(Vsf, "�걾����", 900, 1)
    Call NewColumn(Vsf, "����ʱ��", 1080, 1)
    Call NewColumn(Vsf, "������", 750, 1)
    Call NewColumn(Vsf, "������", 750, 1)
    Call NewColumn(Vsf, "����ʱ��", 1080, 1)
    Call NewColumn(Vsf, "�������", 1200, 1)
    Call NewColumn(Vsf, "������", 750, 1)
    Call NewColumn(Vsf, "��������", 1200, 1)
    Call NewColumn(Vsf, "ִ�п���", 1200, 1)
    Call NewColumn(Vsf, "ҽ��id", 0, 1)
    Call NewColumn(Vsf, "���ͺ�", 0, 1)
    Call NewColumn(Vsf, "ת��", 0, 1)
    Call NewColumn(Vsf, "�걾ID", 0, 1)
    Call NewColumn(Vsf, "����ID", 0, 1)
    Call NewColumn(Vsf, "�Ƿ����", 0, 1)
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
'    If lngExecDept > 0 Then strWhere = strWhere & " And A.ִ�п���ID+0=" & lngExecDept
    blnMoved = MovedByDate(dtp(0).Value)
    
    If Trim(txt(2).Text) <> "" Then strWhere = strWhere & " AND A.������ = '" & Trim(txt(2).Text) & "'"
    If cbo(1).ListIndex > 0 Then strWhere = strWhere & " AND B.��������ID + 0 = " & cbo(1).ItemData(cbo(1).ListIndex)
    If cbo(0).ListIndex >= 0 Then strWhere = strWhere & _
        IIf(cbo(0).ListIndex = 0, " AND A.����id IS Null", " AND A.����id=" & cbo(0).ItemData(cbo(0).ListIndex))
    If cbo(2).ListIndex >= 0 Then strWhere = strWhere & " AND A.ִ�п���ID + 0=" & cbo(2).ItemData(cbo(2).ListIndex)
        
    If Trim(txt(0).Text) <> "" Then
'        varTmp2 = Split(Trim(txt(0).Text), ",")
'        strTmp = ""
'        For mlngLoop = 0 To UBound(varTmp2)
'            If InStr(varTmp2(mlngLoop), "-") = 0 Then
'                strTmp = strTmp & "  OR A.�걾���=" & TransSampleNO(varTmp2(mlngLoop))
'            Else
'                strTmp = strTmp & "  OR A.�걾��� BETWEEN " & TransSampleNO(Mid(varTmp2(mlngLoop), 1, InStr(varTmp2(mlngLoop), "~") - 1)) & " AND " & TransSampleNO(Mid(varTmp2(mlngLoop), InStr(varTmp2(mlngLoop), "~") + 1))
'            End If
'        Next
'        If strTmp <> "" Then strWhere = strWhere & " AND (1=2 " & strTmp & ")"
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
    
    Select Case cbo(3).ListIndex
    Case 0
        strOrderBy = " ORDER BY ��������,����ʱ��,�걾��"
    Case 1
        strOrderBy = " ORDER BY ��������,�������"
    Case 2
        strOrderBy = " ORDER BY �������,����ʱ��,�걾��"
    End Select
                
    mstrSQL = "select DISTINCT B.���ID AS ID,A.ҽ��id,F.���ͺ�,0 AS ѡ��," & _
                      " Decode(A.����id, Null, " & vbCrLf & _
                        " to_Char(Trunc(A.�걾���/10000)+1,'0000')|| '-'||to_Char(MOD(A.�걾���,10000),'0000'), A.�걾���) As �걾��, " & _
                      "A.�걾����," & _
                      "TO_CHAR(A.����ʱ��,'MM-DD HH24:MI') AS ����ʱ��," & _
                      "A.������," & _
                      "A.������," & _
                      "TO_CHAR(B.����ʱ��,'MM-DD HH24:MI') AS ����ʱ��," & _
                      "B.����ҽ�� AS ������," & _
                      "C.���� AS �������," & _
                      "E.���� AS ִ�п���," & _
                      "A.id as �걾ID, " & _
                      "B.����id, " & _
                      "D.���� AS ��������,0 As ת��,Decode(A.�걾���,1,'��','') As ����, " & _
                      "decode(a.���ʱ��,Null,'��','��') as �Ƿ���� " & _
                 "from ����걾��¼ A, ����ҽ����¼ B, ���ű� C, �������� D,���ű� E,����ҽ������ F " & _
                "WHERE A.ҽ��ID = B.���ID AND C.ID = B.��������ID AND B.ID=F.ҽ��id AND " & _
                      "A.����ID = D.ID(+) AND E.ID=B.ִ�п���id AND A.����״̬ IN (1,2) " & strWhere
    If blnMoved Then
        strSQLBak = mstrSQL
        strSQLBak = Replace(strSQLBak, "0 As ת��", "1 As ת��")
        strSQLBak = Replace(strSQLBak, "����ҽ����¼", "H����ҽ����¼")
        strSQLBak = Replace(strSQLBak, "����ҽ������", "H����ҽ������")
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
    
    If ErrCenter = 1 Then
        Resume
    End If
    
    Call SaveErrLog
End Function

Private Function SaveData() As Boolean
    '------------------------------------------------------------------------------------------------------
    '���ܣ�
    '------------------------------------------------------------------------------------------------------
    Dim strReportCode As String
    Dim strReportParaNo As String
    Dim bytReportParaMode As Byte
    Dim blPrint As Boolean
    
    On Error GoTo ErrHand
    
    For mlngLoop = 1 To Vsf.Rows - 1
        If Abs(Val(Vsf.TextMatrix(mlngLoop, mCol.ѡ��))) = 1 And Val(Vsf.RowData(mlngLoop)) > 0 Then
            If GetReportCode(Val(Vsf.TextMatrix(mlngLoop, mCol.ҽ��id)), Val(Vsf.TextMatrix(mlngLoop, mCol.���ͺ�)), strReportCode, strReportParaNo, bytReportParaMode, _
                Val(Vsf.TextMatrix(mlngLoop, mCol.ת��)) = 1) Then
                
                If Vsf.TextMatrix(mlngLoop, mCol.�Ƿ����) = "��" And blPrint = False Then
                    If MsgBox("��ǰ���鵥��δ��ˣ��Ƿ�ȷ����Ҫ��ӡ��", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
                        blPrint = True
                        Call ReportOpen(gcnOracle, glngSys, strReportCode, Me, "NO=" & strReportParaNo, "����=" & bytReportParaMode, _
                        "ҽ��ID=" & Val(Vsf.TextMatrix(mlngLoop, mCol.ҽ��id)), _
                        "����ID=" & Val(Vsf.TextMatrix(mlngLoop, mCol.����ID)), _
                        "�걾ID=" & Val(Vsf.TextMatrix(mlngLoop, mCol.�걾ID)), 2)
                    Else
                        Exit Function
                    End If
                Else
                    Call ReportOpen(gcnOracle, glngSys, strReportCode, Me, "NO=" & strReportParaNo, "����=" & bytReportParaMode, _
                        "ҽ��ID=" & Val(Vsf.TextMatrix(mlngLoop, mCol.ҽ��id)), _
                        "����ID=" & Val(Vsf.TextMatrix(mlngLoop, mCol.����ID)), _
                        "�걾ID=" & Val(Vsf.TextMatrix(mlngLoop, mCol.�걾ID)), 2)
                End If
            End If
        End If
    Next
    
    SaveData = True
    
    Exit Function
    
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
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

            If SaveData() = False Then Exit Sub
            
            mblnOK = True
            
            mblnChangeEdit = False
            Call AdjustEnableState

            Unload Me
            Exit Sub
        End If
        
    Case 3
        ShowHelp App.ProductName, Me.Hwnd, Me.Name, Int((glngSys) / 100)
    Case 4
        Unload Me
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
    
'    cbo(0).ListIndex = 0
    cbo(2).ListIndex = 0
    
    zlControl.CboLocate cbo(1), UserInfo.����ID, True
    If cbo(1).ListCount > 0 And cbo(1).ListIndex = -1 Then cbo(1).ListIndex = 0
    
    txt(0).Text = ""
    txt(2).Text = ""
    
    dtp(0).SetFocus
End Sub

Private Sub dtp_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlcommfun.PressKey vbKeyTab
End Sub

Private Sub Form_Activate()
    Dim rs As New ADODB.Recordset
    Dim lngDefaultDev As Long, mlngLoop As Long
    Dim ControlcboDept As CommandBarComboBox
    
    If mblnStartUp = False Then Exit Sub
    mblnStartUp = False
    
    '���벿��
    mstrSQL = "SELECT A.����||'-'||A.����,ID FROM ���ű� A,��������˵�� B WHERE A.ID=B.����id AND B.��������='�ٴ�' ORDER BY A.����||'-'||A.����"
    Call OpenRecord(rs, mstrSQL, Me.Caption)
    cbo(1).AddItem "���п���"
    If rs.BOF = False Then Call AddComboData(cbo(1), rs, False)
    If cbo(1).ListCount > 0 And cbo(1).ListIndex = -1 Then cbo(1).ListIndex = 0
    
'    '���˿���
'    mstrSQL = "SELECT A.����||'-'||A.����,ID FROM ���ű� A,��������˵�� B WHERE A.ID=B.����id AND B.��������='�ٴ�' ORDER BY A.����||'-'||A.����"
'    Call OpenRecord(rs, mstrSQL, Me.Caption)
'    cbo(2).AddItem "���п���"
'    If rs.BOF = False Then Call AddComboData(cbo(2), rs, False)
'    zlControl.CboLocate cbo(2), UserInfo.����ID, True
'    If cbo(2).ListCount > 0 And cbo(2).ListIndex = -1 Then cbo(2).ListIndex = 0
    cbo(2).Clear
'    cbo(2).AddItem "���в���"
'    For mlngLoop = 0 To ControlcboDept.ListCount - 1
'        cbo(2).AddItem ControlcboDept.List(mlngLoop)
'        cbo(2).ItemData(cbo(2).NewIndex) = ControlcboDept.ItemData(mlngLoop)
'    Next
    mstrSQL = " SELECT A.����||'-'||A.����,ID FROM ���ű� A,��������˵�� B " & _
              " WHERE A.ID=B.����id AND a.ID = [1] ORDER BY A.����||'-'||A.����  "
              
    Set rs = zldatabase.OpenSQLRecord(mstrSQL, Me.Caption, lngExecDept)
    If rs.BOF = False Then
        Call AddComboData(cbo(2), rs, False)
        cbo(2).ListIndex = cbo(2).ListIndex = 1
    End If
    
    '��������
    mstrSQL = "SELECT A.����||'-'||A.����,ID FROM �������� A where a.ʹ��С��ID =[1] ORDER BY A.����||'-'||A.����"
    Set rs = zldatabase.OpenSQLRecord(mstrSQL, Me.Caption, lngExecDept)
    cbo(0).AddItem "�ֹ�"
    If rs.BOF = False Then Call AddComboData(cbo(0), rs, False)
    lngDefaultDev = Val(Split(GetConnectDevs & ";1", ";")(0))
    cbo(0).ListIndex = FindComboItem(cbo(0), lngDefaultDev)
    If cbo(0).ListCount > 0 And cbo(0).ListIndex = -1 Then cbo(0).ListIndex = 0
    
    cbo(3).AddItem "���� + �걾���"
    cbo(3).AddItem "���� + �������"
    cbo(3).AddItem "������� + �걾���"
    cbo(3).ListIndex = 0
    
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
'    If mblnChangeEdit Then
'        Cancel = (MsgBox("�������޸ĵ����ݱ��뱣������Ч���Ƿ񲻱�����˳���", vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) = vbNo)
'        If Cancel Then Exit Sub
'    End If
    
    Call SaveWinState(Me, App.ProductName)
End Sub



Private Sub tbrThis_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
    Case "ȫѡ"
        Call cmd_Click(0)
    Case "ȫ��"
        Call cmd_Click(1)
    Case "��ӡ"
        Call cmd_Click(2)
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
            KeyAscii = FilterKeyAscii(KeyAscii, 99, "0123456789,~-")
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





