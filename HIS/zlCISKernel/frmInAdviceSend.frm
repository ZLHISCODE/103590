VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmInAdviceSend 
   AutoRedraw      =   -1  'True
   Caption         =   "סԺ��������"
   ClientHeight    =   6510
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9615
   Icon            =   "frmInAdviceSend.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6510
   ScaleWidth      =   9615
   Begin VB.Frame fraSetup 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   120
      TabIndex        =   9
      Top             =   840
      Visible         =   0   'False
      Width           =   9315
      Begin VB.ComboBox cboDrugType 
         Height          =   300
         Left            =   3045
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   0
         Visible         =   0   'False
         Width           =   2715
      End
      Begin VB.Frame fraBaby 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   6120
         TabIndex        =   10
         Top             =   50
         Width           =   3195
         Begin VB.OptionButton optBaby 
            Caption         =   "����ҽ��"
            Height          =   180
            Index           =   1
            Left            =   1080
            TabIndex        =   13
            Top             =   0
            Width           =   1020
         End
         Begin VB.OptionButton optBaby 
            Caption         =   "����ҽ��"
            Height          =   180
            Index           =   0
            Left            =   0
            TabIndex        =   12
            Top             =   0
            Value           =   -1  'True
            Width           =   1020
         End
         Begin VB.OptionButton optBaby 
            Caption         =   "Ӥ��ҽ��"
            Height          =   180
            Index           =   2
            Left            =   2175
            TabIndex        =   11
            Top             =   0
            Width           =   1020
         End
      End
      Begin VB.Label lblDrugType 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ҩ������"
         Height          =   180
         Left            =   2280
         TabIndex        =   15
         Top             =   45
         Visible         =   0   'False
         Width           =   720
      End
   End
   Begin VB.TextBox txtPer 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   180
      Left            =   6615
      TabIndex        =   2
      TabStop         =   0   'False
      Text            =   "100%"
      Top             =   6255
      Visible         =   0   'False
      Width           =   405
   End
   Begin MSComctlLib.ProgressBar psb 
      Height          =   270
      Left            =   2115
      TabIndex        =   3
      Top             =   6210
      Visible         =   0   'False
      Width           =   4515
      _ExtentX        =   7964
      _ExtentY        =   476
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   4
      Top             =   6150
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmInAdviceSend.frx":058A
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12091
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   318
            MinWidth        =   25
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   318
            MinWidth        =   25
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmInAdviceSend.frx":0E1E
            Key             =   "PY"
            Object.ToolTipText     =   "ƴ��(F7)"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmInAdviceSend.frx":1458
            Key             =   "WB"
            Object.ToolTipText     =   "���(F7)"
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
   Begin VB.Frame fraUD 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   0
      MousePointer    =   7  'Size N S
      TabIndex        =   5
      Top             =   4605
      Width           =   9495
   End
   Begin VB.Frame fraInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   60
      TabIndex        =   6
      Top             =   525
      Width           =   9435
      Begin VB.Label lblPati 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00C0FFFF&
         Height          =   180
         Left            =   120
         TabIndex        =   7
         Top             =   60
         Width           =   90
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsAdvice 
      Height          =   3405
      Left            =   0
      TabIndex        =   0
      Top             =   1185
      Width           =   9540
      _cx             =   16828
      _cy             =   6006
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
      BackColorSel    =   16771802
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   0
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   2000
      ColWidthMin     =   0
      ColWidthMax     =   5000
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmInAdviceSend.frx":1A92
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   0   'False
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
      OwnerDraw       =   1
      Editable        =   0
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
      AllowUserFreezing=   1
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VSFlex8Ctl.VSFlexGrid vsPrice 
      Height          =   1470
      Left            =   0
      TabIndex        =   1
      Top             =   4665
      Width           =   9525
      _cx             =   16801
      _cy             =   2593
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
      BackColorSel    =   4210752
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   3
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   5
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
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
   Begin MSComCtl2.DTPicker dkpExecTime 
      Height          =   300
      Left            =   2880
      TabIndex        =   8
      Top             =   120
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd HH:mm"
      Format          =   216727555
      CurrentDate     =   40976
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   105
      Top             =   60
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmInAdviceSend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlng����ID As Long 'IN
Private mlng��ҳID As Long 'IN
Private mstrǰ��IDs As String 'IN
Private mlng���˲���ID As Long 'IN
Private mlng���˿���id As Long 'IN
Private mlngҽ������ID As Long 'IN
Private mblnSend As Boolean 'OUT:�Ƿ�ɹ����͹���
Private mblnRefresh As Boolean 'OUT'�Ƿ���Ҫˢ��������
Private mstr���� As String
Private mstrסԺ�� As String
Private mstr���� As String
Private mstr�Ա� As String

Private mlngNOSequence As Long
Private mcolStock1 As Collection '��Ÿ���ҩƷ�ⷿ�ĳ����鷽ʽ
Private mcolStock2 As Collection '��Ÿ������Ŀⷿ�ĳ����鷽ʽ
Private mrsPati As ADODB.Recordset '����������Ϣ
Private mrsPrice As ADODB.Recordset '�����Ƽ۹�ϵ
Private mrsBill As ADODB.Recordset
Private mrsRXKey As ADODB.Recordset

Private mstrLike As String
Private mblnFirst As Boolean
Private mint���� As Integer
Private mint���� As Integer
Private mbln��ҩ�� As Boolean
Private mstr��ҩ�� As String
Private mlng��ҩ����ID As Long
Private mblnAutoExe As Boolean
Private mblnһ����ҩ����Ϊһ�� As Boolean 'һ����ҩ��ҩƷ��Ӧ�Ĵ����㲻ͬʱ���Ƿ��Է���Ϊһ�ŵ���
Private mlngRefModld As Long        '0����ҽ����1=����ҽ��
Private mobjCustom As CommandBarControlCustom
Private mobjlblMsg As CommandBarControl
Private mstr���s As String
Private mstrҩƷ As String
Private mbln���͵��������� As Boolean  '�����˷��͵��������ĵĿ��ң��Ų�����ҩ��¼
Private mblnʹ��Ԥ�� As Boolean '���������֧������ʹ��Ԥ����
Private mclsMipModule As zl9ComLib.clsMipModule
Attribute mclsMipModule.VB_VarHelpID = -1
Private mbln���鵥���������� As Boolean  '����ҽ��������������
Private mlng��ҩ�� As Long
Private mlng��ҩ�� As Long
Private mlng��ҩ�� As Long
Private mlng���ϲ��� As Long
Private mstrҩƷ�۸�ȼ� As String '���˵�ҩƷ�۸�ȼ�
Private mstr���ļ۸�ȼ� As String '���˵����ļ۸�ȼ�
Private mstr��ͨ��Ŀ�۸�ȼ� As String '���˵���ͨ��Ŀ�۸�ȼ�
Private mbln������ҩ As Boolean  'Ƥ��������ҩ �����������ô˲��������ж�Ƥ�Խ��������Ҫ��дƤ��������ҩ˵��
Private mstrAdDrugIDs As String '���һ���������˵����ҩƷ��ҽ��ID����
Private mblnҽ������Χ As Boolean '�Ƿ���Էֿ�����Ӥ���Ͳ���ҽ��

'----------------------------------------------
Private Const COL_ѡ�� = 0
Private Const COL_Ӥ�� = 1
Private Const col_ҽ������ = 2
Private Const COL_���� = 3
Private Const COL_������λ = 4
Private Const COL_���� = 5
Private Const COL_������λ = 6
Private Const COL_��� = 7
Private Const COL_Ƶ�� = 8
Private Const COL_�÷� = 9
Private Const COL_ҽ������ = 10 'Data���ڴ��ժҪ(ҽ��)
Private Const COL_ִ��ʱ�� = 11
Private Const COL_ִ�п��� = 12
Private Const COL_ִ������ = 13
Private Const COL_ID = 14 '������
Private Const COL_���ID = 15
Private Const COL_ҽ��״̬ = 16
Private Const COL_���˿���ID = 17
Private Const COL_��������ID = 18
Private Const COL_����ҽ�� = 19
Private Const COL_����ʱ�� = 20
Private Const COL_������� = 21
Private Const COL_������ĿID = 22
Private Const COL_�걾��λ = 23
Private Const COL_��鷽�� = 24
Private Const COL_ִ�б�� = 25
Private Const COL_�Ƽ����� = 26
Private Const COL_ִ������ID = 27
Private Const COL_ִ�п���ID = 28
Private Const COL_�������� = 29
Private Const COL_�Թܱ��� = 30
Private Const COL_�շ�ϸĿID = 31
Private Const COL_����ϵ�� = 32
Private Const COL_סԺ��װ = 33
Private Const COL_סԺ��λ = 34
Private Const COL_�ɷ���� = 35
Private Const COL_��� = 36
Private Const COL_���� = 37
Private Const COL_�ֽ�ʱ�� = 38
Private Const COL_�״�ʱ�� = 39
Private Const COL_ĩ��ʱ�� = 40
Private Const COL_ǩ��ID = 41
Private Const COL_������־ = 42
Private Const COL_���㷽ʽ = 43
Private Const COL_ִ�а��� = 44
Private Const COL_��ʼʱ�� = 45
Private Const COL_������� = 46
Private Const COL_ִ�з��� = 47
Private Const COL_�������� = 48
Private Const COL_������� = 49
Private Const COL_��ҩ���� = 50


'-------------------------------------------------
Private Const COLP_�к� = 0
Private Const COLP_�շ�ϸĿID = 1
Private Const COLP_�̶� = 2
Private Const COLP_��� = 3
Private Const COLP_�Ƽ�ҽ�� = 4 '�ɼ���
Private Const COLP_��� = 5
Private Const COLP_�շ���Ŀ = 6
Private Const COLP_�Ƽ����� = 7
Private Const COLP_���� = 8
Private Const COLP_��λ = 9
Private Const COLP_���� = 10
Private Const COLP_Ӧ�ս�� = 11
Private Const COLP_ʵ�ս�� = 12
Private Const COLP_ִ�п��� = 13
Private Const COLP_�������� = 14
Private Const COLP_���� = 15
Private Const COLP_�շѷ�ʽ = 16
Private Const COLP_�շ���� = 17 '������
Private Const COLP_ִ�п���ID = 18
Private Const COLP_�������� = 19
Private Const COLP_�������� = 20

Private Enum ESend
    EInBilling = 0  'סԺ���ʵ�
    EOutCharge = 1  '�����շѵ�
    EOutBilling = 2 '������ʵ�
End Enum
Private mbytSendKind As ESend 'IN:0-����סԺ���ʣ�1=���������շ�,2=�����������
Private mlng�������� As Long  'IN:0-��ͨסԺ����,1-�������۲���,2-סԺ���۲���
Private mintҽ������Χ As Integer    'ҽ������Χ   0-����ҽ��,1-����ҽ��,2-Ӥ��ҽ��
Private mlngҽ������ID As Long

Private Property Let Progress(ByVal vNewValue As Single)
'vNewValue=0-100
    If vNewValue = 0 Then
        psb.value = 0: txtPer.Text = ""
        psb.Visible = False: txtPer.Visible = False
    Else
        psb.value = vNewValue
        txtPer.Text = CInt(psb.value) & "%"
        psb.Visible = True: txtPer.Visible = True
        txtPer.Refresh
    End If
End Property

Public Function ShowMe(frmParent As Object, ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal strǰ��IDs As String, _
    ByVal lng���˲���ID As Long, ByVal lng���˿���ID As Long, ByVal lngҽ������ID As Long, blnRefresh As Boolean, _
    ByVal bytSendKind As Byte, ByVal lng�������� As Long, Optional ByVal lngҽ������ID As Long, Optional ByRef objMip As Object) As Boolean
'���ܣ�����ҽ��
'������blnRefresh=�Ƿ�ˢ������������
'      bytSendKind=0-����סԺ���ʣ�1=���������շ�,2=�����������
'      strǰ��IDs ҽ��վ�´�ҽ����ǰ��ID
    mlng����ID = lng����ID
    mlng��ҳID = lng��ҳID
    mstrǰ��IDs = strǰ��IDs
    mlng���˲���ID = lng���˲���ID
    mlng���˿���id = lng���˿���ID
    mlngҽ������ID = lngҽ������ID
    mlngҽ������ID = lngҽ������ID
    mbytSendKind = bytSendKind
    mlng�������� = lng��������
    If InitObjPublicExpense Then Call gobjPublicExpense.zlGetPriceGrade(gstrNodeNo, mlng����ID, mlng��ҳID, "", mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�)
    If Not (objMip Is Nothing) Then Set mclsMipModule = objMip
    On Error Resume Next
    Me.Show 1, frmParent
    err.Clear: On Error GoTo 0
    blnRefresh = mblnRefresh
    ShowMe = mblnSend
End Function

Private Sub cboDrugType_Click()
'��ȡ����
    If Val(cboDrugType.Tag) <> cboDrugType.ListIndex Then
        Call LoadAdviceSend(mstr���s, mstrҩƷ, 0)
        cboDrugType.Tag = cboDrugType.ListIndex
    End If
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim i As Long
    
    If Not Control.Visible Then Exit Sub
    
    Select Case Control.ID
    Case conMenu_Edit_Send
        Call FuncSendAdvice(Control)
    
    Case conMenu_View_Refresh
        mobjCustom.Visible = False
        mobjlblMsg.Visible = False
        Call LoadAdviceSend(mstr���s, mstrҩƷ, 0)
    Case conMenu_View_RefreshSpare
        mobjCustom.Visible = True
        mobjlblMsg.Visible = True
        Call LoadAdviceSend(mstr���s, mstrҩƷ, 1)
    Case conMenu_Tool_Option
        With frmInAdviceSendCond
            .Show 1, Me
            If .mblnOK Then
                Call LoadAdviceSend(.mstr���s, .mstrҩƷ)
                If mbln��ҩ�� Then Call Refresh��ҩ��
            End If
        End With
    Case conMenu_Edit_SelAll
        With vsAdvice
            For i = .FixedRows To .Rows - 1
                If .Cell(flexcpData, i, COL_ѡ��) = 0 Then
                    Set .Cell(flexcpPicture, i, COL_ѡ��) = frmIcons.imgTrueFalse.ListImages("T").Picture
                End If
            Next
        End With
        Call ShowSendTotal
    Case conMenu_Edit_ClsAll
        With vsAdvice
            For i = .FixedRows To .Rows - 1
                If .Cell(flexcpData, i, COL_ѡ��) = 0 Then
                    Set .Cell(flexcpPicture, i, COL_ѡ��) = Nothing
                End If
            Next
        End With
        Call ShowSendTotal
    Case conMenu_Help_Help
        ShowHelp App.ProductName, Me.hwnd, Me.Name
    Case conMenu_File_Exit
        Unload Me
    End Select
End Sub

Private Sub cbsMain_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long

    Call Me.cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    
    On Error Resume Next
    
    fraInfo.Top = lngTop
    fraInfo.Left = lngLeft
    fraInfo.Width = lngRight - lngLeft
    
    fraSetup.Top = fraInfo.Top + fraInfo.Height
    fraSetup.Left = lngLeft
    fraSetup.Width = lngRight - lngLeft
    
    If fraSetup.Visible Then
        'mblnҽ������Χ Or gbln����ҩƷ�ֿ�����
        If gbln����ҩƷ�ֿ����� Then
            fraSetup.Height = 400
            fraBaby.Top = 100
            lblDrugType.Top = 100
            cboDrugType.Top = 50
            cboDrugType.Left = fraInfo.Width - cboDrugType.Width - 50 - IIF(mblnҽ������Χ, fraBaby.Width, 0)
            lblDrugType.Left = cboDrugType.Left - lblDrugType.Width - 50
        End If
    End If
    
    fraBaby.Left = fraSetup.Width - fraBaby.Width
    
    vsAdvice.Left = lngLeft
    vsAdvice.Top = IIF(fraSetup.Visible, fraSetup.Top + fraSetup.Height, fraInfo.Top + fraInfo.Height)
    vsAdvice.Width = lngRight - lngLeft
    vsAdvice.Height = lngBottom - lngTop - fraInfo.Height - vsPrice.Height - fraUD.Height - stbThis.Height
    
    fraUD.Top = vsAdvice.Top + vsAdvice.Height
    fraUD.Left = lngLeft
    fraUD.Width = vsAdvice.Width
    
    vsPrice.Left = lngLeft
    vsPrice.Top = fraUD.Top + fraUD.Height
    vsPrice.Width = vsAdvice.Width
    
    psb.Top = stbThis.Top + Screen.TwipsPerPixelY * 4
    psb.Width = stbThis.Panels(2).Width - txtPer.Width - Screen.TwipsPerPixelX * 6
    psb.Left = stbThis.Panels(2).Left + Screen.TwipsPerPixelX * 2
    
    txtPer.Left = psb.Left + psb.Width
    txtPer.Top = psb.Top + (psb.Height - txtPer.Height) / 2
 
    Me.Refresh
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case conMenu_View_Find
        Control.Visible = mbln��ҩ��
    End Select
End Sub

Private Sub dkpExecTime_Change()
    Call LoadAdviceSend(frmInAdviceSendCond.mstr���s, frmInAdviceSendCond.mstrҩƷ, 1)
End Sub

Private Sub Form_Activate()
    If mblnFirst Then
        mblnFirst = False
        
        '��ȡ�����嵥
        Me.Refresh
        mstr���s = zlDatabase.GetPara("סԺ�����������", glngSys, pסԺҽ���´�)
        mstrҩƷ = zlDatabase.GetPara("סԺ����ҩƷ�������", glngSys, pסԺҽ���´�)
        If LoadAdviceSend(mstr���s, mstrҩƷ) Then
            If mbln��ҩ�� Then Call Refresh��ҩ��
        Else
            Unload Me: Exit Sub
        End If
    End If
End Sub

Private Function GetPatiInfo() As Boolean
'���ܣ���ȡ������Ϣ
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = _
        " Select ����ID,Ԥ�����,�������,0 as Ԥ����� From ������� Where ����=1 And ����ID=[1] And ���� = " & IIF(mlng�������� = 1, 1, 2) & _
        " Union ALL" & _
        " Select A.����ID,0,0,Sum(���) From ����ģ����� A,������ҳ B" & _
        " Where A.����ID=B.����ID And A.��ҳID=B.��ҳID And B.���� Is Not Null And A.����ID=[1] And A.��ҳID=[2] Group by A.����ID"
    strSQL = "Select ����ID,Nvl(Sum(Ԥ�����),0)-Nvl(Sum(�������),0)+Nvl(Sum(Ԥ�����),0) as ʣ��� From (" & strSQL & ") Group by ����ID"
    
    strSQL = "Select A.�����,B.סԺ��,Nvl(B.����,A.����) ����,Nvl(B.�Ա�,A.�Ա�) �Ա�,Nvl(B.����,A.����) ����,B.��Ժ���� as ����," & _
        " B.�ѱ�,B.ҽ�Ƹ��ʽ,B.����,C.ʣ���," & _
        " B.״̬,B.��������,zl_PatiWarnScheme(A.����ID,B.��ҳID) as ���ò���," & _
        " Decode(A.������,Null,Null,zl_PatientSurety(A.����ID,B.��ҳID)) as ������,a.��ͥ�绰 as PhoneNO," & _
        "To_Char(A.��������,'YYYY-MM-DD HH24:MI:SS') as Birthdate,a.��ͥ��ַ as Address" & _
        " From ������Ϣ A,������ҳ B,(" & strSQL & ") C" & _
        " Where A.����ID=B.����ID And A.����ID=C.����ID(+)" & _
        " And B.����ID=[1] And B.��ҳID=[2]"
    Set mrsPati = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng��ҳID)
    
    lblPati.Caption = _
        IIF(mlng�������� = 1, "�����:" & NVL(mrsPati!�����), "סԺ��:" & NVL(mrsPati!סԺ��)) & "������:" & mrsPati!���� & "���Ա�:" & NVL(mrsPati!�Ա�) & "������:" & NVL(mrsPati!����) & _
        "������:" & NVL(mrsPati!����) & "���ѱ�:" & NVL(mrsPati!�ѱ�) & "������:" & NVL(mrsPati!��������) & "�����ʽ:" & NVL(mrsPati!ҽ�Ƹ��ʽ) & _
        "��ʣ���:" & Format(NVL(mrsPati!ʣ���, 0), "0.00")
    mint���� = NVL(mrsPati!����, 0)
    mstr���� = mrsPati!���� & ""
    mstrסԺ�� = NVL(mrsPati!סԺ��)
    mstr���� = NVL(mrsPati!����)
    mstr�Ա� = mrsPati!���� & ""
    '���ղ����ú�ɫ��ʾ
    If Not IsNull(mrsPati!����) Then lblPati.ForeColor = zlDatabase.GetPatiColor(NVL(mrsPati!��������))
    GetPatiInfo = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF7 Then '�л����뷨
        If stbThis.Panels("WB").Visible And stbThis.Panels("PY").Visible Then
            If stbThis.Panels("WB").Bevel = sbrRaised Then
                Call stbThis_PanelClick(stbThis.Panels("WB"))
            Else
                Call stbThis_PanelClick(stbThis.Panels("PY"))
            End If
        End If
    End If
End Sub

Private Sub Form_Load()
    Dim objBar As CommandBar
    Dim objControl As CommandBarControl
    Dim objCbo As CommandBarComboBox
    Dim strSQL As String
    Dim strPar As String
    Dim blnDo As Boolean
    
    If Not PatiFeeUsable(mlng����ID, mlng��ҳID) Then Unload Me: Exit Sub
    '������----------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsMain.VisualTheme = xtpThemeOffice2003
    With Me.cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
    End With
    cbsMain.EnableCustomization False
    cbsMain.ActiveMenuBar.Visible = False
    Set cbsMain.Icons = zlcommfun.GetPubIcons
    
    '���ɹ�����
    Set objBar = cbsMain.Add("������", xtpBarTop)
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Send, "����")
        If mbytSendKind = EInBilling Then
            objControl.Caption = "����סԺ����"
        ElseIf mbytSendKind = EOutCharge Then
            objControl.Caption = "���������շ�"
        Else
            objControl.Caption = "�����������"
        End If
        
        Set objControl = .Add(xtpControlButton, conMenu_View_Refresh, "��ȡ����ҽ��")
            objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_View_RefreshSpare, "��ȡ����ҽ��")
            
        Set mobjlblMsg = .Add(xtpControlLabel, conMenu_View_RefreshSpare * 100# + 1, "��ִ��ʱ�䡿:")
        mobjlblMsg.Visible = False
        Set mobjCustom = .Add(xtpControlCustom, conMenu_View_RefreshSpare * 100# + 2, "")
        mobjCustom.ToolTipText = "������ñ���ҽ��ִ�е�ʱ�䡣"
        dkpExecTime.value = zlDatabase.Currentdate
        mobjCustom.Handle = dkpExecTime.hwnd
        mobjCustom.Visible = False
        
        Set objControl = .Add(xtpControlButton, conMenu_Tool_Option, "ѡ��")
            objControl.BeginGroup = True
            objControl.IconId = conMenu_File_Parameter
        Set objControl = .Add(xtpControlButton, conMenu_Edit_SelAll, "ȫѡ")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_ClsAll, "ȫ��")
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "����")
            objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�")
    End With
    objBar.EnableDocking xtpFlagHideWrap
    objBar.ContextMenuPresent = False
    For Each objControl In objBar.Controls
        If objControl.Type <> xtpControlCustom And objControl.Type <> xtpControlLabel Then
            objControl.Style = xtpButtonIconAndCaption
        End If
    Next
    
    With cbsMain.KeyBindings
        .Add FCONTROL, vbKey1, conMenu_Edit_Send
        .Add 0, vbKeyF2, conMenu_Edit_Send
        .Add 0, vbKeyF12, conMenu_Tool_Option
        .Add FCONTROL, vbKeyA, conMenu_Edit_SelAll
        .Add FCONTROL, vbKeyR, conMenu_Edit_ClsAll
        .Add 0, vbKeyF1, conMenu_Help_Help
        .Add FALT, vbKeyX, conMenu_File_Exit
        .Add 0, vbKeyF5, conMenu_View_Refresh
    End With
        
    '���˵��Ҳ����ҩ��
    objBar.EnableDocking xtpFlagStretched
    With objBar.Controls
        Set objCbo = .Add(xtpControlComboBox, conMenu_View_Find, "��ҩ��")
            objCbo.BeginGroup = True
            objCbo.Flags = xtpFlagRightAlign
            objCbo.Style = xtpComboLabel
            objCbo.Width = 200
    End With
    '-----------------------------------------------------
    Call InitAdviceTable
    Call InitPriceTable
    Call RestoreWinState(Me, App.ProductName)
    
    mbln��ҩ�� = Val(zlDatabase.GetPara(27, glngSys)) <> 0
    If mstrǰ��IDs <> "" Then
        mlng��ҩ����ID = mlngҽ������ID
    Else
        mlng��ҩ����ID = GetICUDeptID
        If mlng��ҩ����ID = 0 Then mlng��ҩ����ID = IIF(mlng���˲���ID <> 0, mlng���˲���ID, mlng���˿���id)
    End If
    
    mblnһ����ҩ����Ϊһ�� = Val(zlDatabase.GetPara("һ����ҩ����Ϊһ��", glngSys, p����ҽ���´�, 1)) = 1
    mblnAutoExe = Val(Mid(zlDatabase.GetPara("����ִ���Զ����", glngSys, pסԺҽ������), 2, 1)) <> 0
    mstrLike = IIF(Val(zlDatabase.GetPara("����ƥ��")) = 0, "%", "")
    mint���� = Val(zlDatabase.GetPara("���뷽ʽ")) '����ƥ�䷽ʽ��0-ƴ��,1-���
    mblnʹ��Ԥ�� = Val(zlDatabase.GetPara("���֧������ʹ��Ԥ����", glngSys, p����ҽ���´�, "1"))
    
    mbln���͵��������� = gstr��Һ�������� <> ""
    If mbln���͵��������� Then
        strPar = zlDatabase.GetPara("��Դ����", glngSys, p��Һ��������, "")
        If strPar <> "" Then
            If InStr("," & strPar & ",", "," & mlng���˿���id & ",") = 0 Then mbln���͵��������� = False
        End If
        strPar = Val(zlDatabase.GetPara("ҽ������", glngSys, p��Һ��������, "1"))
        If Val(strPar) = 1 Then
            mbln���͵��������� = False
        End If
    End If
    
    mbln���鵥���������� = Val(zlDatabase.GetPara("����ҽ��������������", glngSys, pסԺҽ������, "0")) = 1
    mlng��ҩ�� = Val(zlDatabase.GetPara("סԺȱʡ��ҩ��", glngSys, pסԺҽ���´�, , , , , mlng���˿���id))
    mlng��ҩ�� = Val(zlDatabase.GetPara("סԺȱʡ��ҩ��", glngSys, pסԺҽ���´�, , , , , mlng���˿���id))
    mlng��ҩ�� = Val(zlDatabase.GetPara("סԺȱʡ��ҩ��", glngSys, pסԺҽ���´�, , , , , mlng���˿���id))
    mlng���ϲ��� = Val(zlDatabase.GetPara("סԺȱʡ���ϲ���", glngSys, pסԺҽ���´�, , , , , mlng���˿���id))
    
    mbln������ҩ = Val(zlDatabase.GetPara("Ƥ��������ҩ", glngSys, pסԺҽ���´�)) <> 0
    
    cboDrugType.AddItem "0-ȫ��"
    cboDrugType.AddItem "1-��Ʒ��"
    cboDrugType.AddItem "2-����;���I��"
    cboDrugType.AddItem "3-����(��1��2��)"
    Call Cbo.SetIndex(cboDrugType.hwnd, 0)
    cboDrugType.Tag = "0"
  
    cboDrugType.Visible = gbln����ҩƷ�ֿ�����
    lblDrugType.Visible = gbln����ҩƷ�ֿ�����
    
    mblnҽ������Χ = DeptIsWoman(0, Get����IDs(mlng���˲���ID))
    If mblnҽ������Χ Then
        fraSetup.Visible = True
        'ҽ������Χ
        mintҽ������Χ = Val(zlDatabase.GetPara("ҽ������Χ", glngSys, pסԺҽ������, "0"))
        optBaby(mintҽ������Χ).value = True
    End If
 
    fraSetup.Visible = (mblnҽ������Χ Or gbln����ҩƷ�ֿ�����)
    fraBaby.Visible = mblnҽ������Χ
    Select Case mint����
        Case 0
            stbThis.Panels("PY").Bevel = sbrInset
            stbThis.Panels("WB").Bevel = sbrRaised
        Case 1
            stbThis.Panels("PY").Bevel = sbrRaised
            stbThis.Panels("WB").Bevel = sbrInset
        Case Else
            stbThis.Panels("PY").Bevel = sbrInset
            stbThis.Panels("WB").Bevel = sbrInset
    End Select
    If Not gbln����ƥ�䷽ʽ�л� Then
        stbThis.Panels("PY").Visible = False
        stbThis.Panels("WB").Visible = False
    End If
   
    mblnSend = False
    mblnRefresh = False
    mblnFirst = True
    
    '�����ⷿҩƷ�����鷽ʽ
    Set mcolStock1 = GetStockCheck(0)
    Set mcolStock2 = GetStockCheck(1)
    
    '��ʾ������Ϣ
    If Not GetPatiInfo Then Unload Me: Exit Sub
    
    'һ��ͨ���㲿��
    If mblnAutoExe And gblnִ��ǰ�Ƚ��� And Not mbytSendKind = EInBilling Then
        If gobjSquareCard Is Nothing Then
            On Error Resume Next
            Set gobjSquareCard = CreateObject("zl9CardSquare.clsCardSquare")
            err.Clear: On Error GoTo 0
            
            If Not gobjSquareCard Is Nothing Then
                If gobjSquareCard.zlInitComponents(Me, pסԺҽ���´�, glngSys, gstrDBUser, gcnOracle, False) = False Then
                    Set gobjSquareCard = Nothing
                    MsgBox "�����㲿����zl9CardSquare����ʼ��ʧ��!�����Զ�ִ�е�ҽ�����ᱻ�Զ�ִ�С�", vbInformation, gstrSysName
                End If
            End If
        End If
    End If
End Sub

Private Function TheStockCheck(ByVal lng�ⷿID As Long, ByVal str��� As String) As Integer
'���ܣ���ȡָ���ⷿ�ĳ������鷽ʽ
    Dim intStyle As Integer
    On Error Resume Next
    If InStr(",5,6,7,", str���) > 0 Then
        intStyle = mcolStock1("_" & lng�ⷿID)
    ElseIf str��� = "4" Then
        intStyle = mcolStock2("_" & lng�ⷿID)
    End If
    err.Clear: On Error GoTo 0
    TheStockCheck = intStyle
End Function

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub
    Call cbsMain_Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
    
    '�ͷ�˽�м�IN����
    mlng��ҳID = 0
    mlng����ID = 0
    Set mrsPati = Nothing
    Set mrsPrice = Nothing
    Set mrsBill = Nothing
    Set mcolStock1 = Nothing
    Set mcolStock2 = Nothing
    Set mrsRXKey = Nothing
    Set mobjCustom = Nothing
    Set mobjlblMsg = Nothing
    Set mclsMipModule = Nothing
    gbln�Ӱ�Ӽ� = False
End Sub

Private Sub fraUD_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If vsAdvice.Height + Y < 1000 Or vsPrice.Height - Y < 500 Then Exit Sub
        fraUD.Top = fraUD.Top + Y
        vsAdvice.Height = vsAdvice.Height + Y
        vsPrice.Top = vsPrice.Top + Y
        vsPrice.Height = vsPrice.Height - Y
        Me.Refresh
    End If
End Sub

Private Sub optBaby_Click(Index As Integer)
    mintҽ������Χ = Index
    Call LoadAdviceSend(mstr���s, mstrҩƷ, 0)
End Sub

Private Sub stbThis_PanelClick(ByVal Panel As MSComctlLib.Panel)
    If Panel.Bevel = sbrRaised And (Panel.Key = "PY" Or Panel.Key = "WB") Then
        '�л����������ƥ�䷽ʽ
        Panel.Bevel = IIF(Panel.Bevel = sbrInset, sbrRaised, sbrInset)
        If Panel.Key = "PY" Then
            stbThis.Panels("WB").Bevel = IIF(Panel.Bevel = sbrInset, sbrRaised, sbrInset)
        Else
            stbThis.Panels("PY").Bevel = IIF(Panel.Bevel = sbrInset, sbrRaised, sbrInset)
        End If
        Call zlDatabase.SetPara("���뷽ʽ", IIF(stbThis.Panels("PY").Bevel = sbrInset And stbThis.Panels("WB").Bevel = sbrInset, 2, IIF(stbThis.Panels("WB").Bevel = sbrInset, 1, 0)))
        mint���� = Val(zlDatabase.GetPara("���뷽ʽ")) '����ƥ�䷽ʽ��0-ƴ��,1-���
    End If
End Sub

Private Sub FuncSendAdvice(ByVal Control As CommandBarControl)
    Dim lng���ͺ� As Long, strMsg As String, i As Long
    Dim objCbo As CommandBarComboBox
    
    With vsAdvice
        For i = .FixedRows To .Rows - 1
            If Val(.TextMatrix(i, COL_ID)) <> 0 And .Cell(flexcpData, i, COL_ѡ��) = 0 And Not .Cell(flexcpPicture, i, COL_ѡ��) Is Nothing Then
                Exit For
            End If
        Next
        If i > .Rows - 1 Then
            MsgBox "��ǰû�п��Է��͵�ҽ����", vbInformation, gstrSysName
            Exit Sub
        End If
    End With
    
    If mbytSendKind = EInBilling Then
        strMsg = "����ҽ�����͵ķ��ý�����ΪסԺ���ʵ��ݣ���ȷ����"
    ElseIf mbytSendKind = EOutCharge Then
        strMsg = "����ҽ�����͵ķ��ý�����Ϊ�����շѵ��ݣ���ȷ����"
    Else
        strMsg = "����ҽ�����͵ķ��ý�����Ϊ������ʵ��ݣ���ȷ����"
    End If
    If MsgBox(strMsg, vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then Exit Sub
    
    lng���ͺ� = SendAdvice()
    If lng���ͺ� <> 0 Then
        mblnSend = True
        
        'ʹ��������ҩ�ŵĴ���
        If mstr��ҩ�� <> "" Then
            Set objCbo = cbsMain.FindControl(, conMenu_View_Find)
            i = objCbo.FindItem(mstr��ҩ��)
            If i = 0 Then
                objCbo.AddItem mstr��ҩ��, 2
                objCbo.ListIndex = 2
            End If
        End If
            
        '��ӡ���Ƶ���
        Call frmSendBillPrint.ShowMe(lng���ͺ�, 2, Me, mstrǰ��IDs)
        
        '���ȫ���������,���˳�
        If vsAdvice.Rows = 2 Then
            If Val(vsAdvice.TextMatrix(1, COL_ID)) = 0 Then
                Unload Me: Exit Sub
            End If
        End If
        Call GetPatiInfo
    End If
End Sub

Private Sub RowSelectSame(ByVal lngRow As Long, ByVal lngCol As Long, _
    Optional rsSQL As ADODB.Recordset, Optional rsTotal As ADODB.Recordset, _
    Optional rsUpload As ADODB.Recordset, Optional strҽ��IDs As String)
'���ܣ����ݿɼ��е�ѡ��״̬,�����ҽ��һ��ѡ��
    Dim i As Long
    
    With vsAdvice
        If lngCol = COL_ѡ�� Then
            For i = lngRow + 1 To .Rows - 1
                If IIF(Val(.TextMatrix(i, COL_���ID)) <> 0, Val(.TextMatrix(i, COL_���ID)), Val(.TextMatrix(i, COL_ID))) _
                    = IIF(Val(.TextMatrix(lngRow, COL_���ID)) <> 0, Val(.TextMatrix(lngRow, COL_���ID)), Val(.TextMatrix(lngRow, COL_ID))) Then
                    .Cell(flexcpData, i, lngCol) = .Cell(flexcpData, lngRow, lngCol)
                    Set .Cell(flexcpPicture, i, lngCol) = .Cell(flexcpPicture, lngRow, lngCol)
                Else
                    Exit For
                End If
            Next
            
            For i = lngRow - 1 To .FixedRows Step -1
                If IIF(Val(.TextMatrix(i, COL_���ID)) <> 0, Val(.TextMatrix(i, COL_���ID)), Val(.TextMatrix(i, COL_ID))) _
                    = IIF(Val(.TextMatrix(lngRow, COL_���ID)) <> 0, Val(.TextMatrix(lngRow, COL_���ID)), Val(.TextMatrix(lngRow, COL_ID))) Then
                    .Cell(flexcpData, i, lngCol) = .Cell(flexcpData, lngRow, lngCol)
                    Set .Cell(flexcpPicture, i, lngCol) = .Cell(flexcpPicture, lngRow, lngCol)
                Else
                    Exit For
                End If
            Next
            
            'һ��������ŵĻ���ҽ��
            If Val(.TextMatrix(lngRow, COL_�������)) <> 0 And .TextMatrix(lngRow, COL_�������) = "Z" Then
                For i = lngRow - 1 To .FixedRows Step -1
                    If Val(.TextMatrix(i, COL_�������)) = Val(.TextMatrix(lngRow, COL_�������)) Then
                        .Cell(flexcpData, i, lngCol) = .Cell(flexcpData, lngRow, lngCol)
                        Set .Cell(flexcpPicture, i, lngCol) = .Cell(flexcpPicture, lngRow, lngCol)
                    Else
                        Exit For
                    End If
                Next
                
                For i = lngRow + 1 To .Rows - 1
                    If Val(.TextMatrix(i, COL_�������)) = Val(.TextMatrix(lngRow, COL_�������)) Then
                        .Cell(flexcpData, i, lngCol) = .Cell(flexcpData, lngRow, lngCol)
                        Set .Cell(flexcpPicture, i, lngCol) = .Cell(flexcpPicture, lngRow, lngCol)
                    Else
                        Exit For
                    End If
                Next
            End If
            
            'ȡ��ѡ��ʱ
            If Not (.Cell(flexcpData, lngRow, COL_ѡ��) = 0 And Not .Cell(flexcpPicture, lngRow, COL_ѡ��) Is Nothing) Then
                i = IIF(Val(.TextMatrix(lngRow, COL_���ID)) = 0, Val(.TextMatrix(lngRow, COL_ID)), Val(.TextMatrix(lngRow, COL_���ID)))
                '1.�����Ӧ�ķ��ü����ͼ�¼��д
                If Not rsSQL Is Nothing Then
                    rsSQL.Filter = "ҽ��ID=" & i
                    Do While Not rsSQL.EOF
                        rsSQL.Delete
                        rsSQL.Update
                        rsSQL.MoveNext
                    Loop
                    rsSQL.Filter = 0 '��ΪҪʹ��BookMark����˻ָ�
                End If
                '2.�����Ӧ�ķ��ͼƼ������ۼ�
                If Not rsTotal Is Nothing Then
                    rsTotal.Filter = "ҽ��ID=" & i
                    Do While Not rsTotal.EOF
                        rsTotal.Delete
                        rsTotal.Update
                        rsTotal.MoveNext
                    Loop
                End If
                '3.�����Ӧ��ҽ���ϴ����ݺ�
                If Not rsUpload Is Nothing Then
                    rsUpload.Filter = "ҽ��ID=" & i
                    Do While Not rsUpload.EOF
                        rsUpload.Delete
                        rsUpload.Update
                        rsUpload.MoveNext
                    Loop
                End If
                '4.��������͵�ǩ��ҽ����ID
                If strҽ��IDs <> "" Then
                    strҽ��IDs = strҽ��IDs & ","
                    strҽ��IDs = Replace(strҽ��IDs, "," & i & ",", ",")
                    If strҽ��IDs <> "" Then
                        strҽ��IDs = Left(strҽ��IDs, Len(strҽ��IDs) - 1)
                    End If
                End If
            End If
        End If
    End With
End Sub

Private Function GetVisibleRow(ByVal lngRow As Long, Optional ByVal blnFirst As Boolean) As Long
'���ܣ�����ָ��ҽ���У����ظ�ҽ���пɼ�����
    Dim lng��ID As Long, i As Long
    
    GetVisibleRow = lngRow
    
    With vsAdvice
        If Not .RowHidden(lngRow) Then Exit Function
        
        'һ����ҩ�Ķ�λ����һҩƷ��
        If blnFirst Then
            If .TextMatrix(lngRow, COL_�������) = "E" And InStr(",5,6,", .TextMatrix(lngRow - 1, COL_�������)) > 0 _
                And Val(.TextMatrix(lngRow, COL_���ID)) = 0 And Val(.TextMatrix(lngRow, COL_ID)) = Val(.TextMatrix(lngRow - 1, COL_���ID)) Then
                i = .FindRow(.TextMatrix(lngRow, COL_ID), , COL_���ID)
                If i <> -1 Then GetVisibleRow = i: Exit Function
            End If
        End If
        
        lng��ID = IIF(Val(.TextMatrix(lngRow, COL_���ID)) <> 0, Val(.TextMatrix(lngRow, COL_���ID)), Val(.TextMatrix(lngRow, COL_ID)))
        For i = lngRow - 1 To .FixedRows Step -1
            If lng��ID = IIF(Val(.TextMatrix(i, COL_���ID)) <> 0, Val(.TextMatrix(i, COL_���ID)), Val(.TextMatrix(i, COL_ID))) Then
                If Not .RowHidden(i) Then GetVisibleRow = i: Exit Function
            Else
                Exit For
            End If
        Next
        For i = lngRow + 1 To .Rows - 1
            If lng��ID = IIF(Val(.TextMatrix(i, COL_���ID)) <> 0, Val(.TextMatrix(i, COL_���ID)), Val(.TextMatrix(i, COL_ID))) Then
                If Not .RowHidden(i) Then GetVisibleRow = i: Exit Function
            Else
                Exit For
            End If
        Next
    End With
End Function

Private Sub vsAdvice_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    With vsAdvice
        If OldRow <> NewRow And .Redraw <> flexRDNone And Not .RowHidden(NewRow) Then
            If Val(.TextMatrix(NewRow, COL_ID)) <> 0 Then
                Call ShowAdvicePrice(NewRow)
                
                'ȱʡѡ��Ƽ�ҽ��(�������)
                Call ShowDefaultRow
            End If
        End If
    End With
End Sub

Private Sub vsAdvice_AfterUserFreeze()
    With vsAdvice
        If .FrozenCols < COL_ѡ�� + 1 - .FixedCols Then
            .FrozenCols = COL_ѡ�� + 1 - .FixedCols
        End If
    End With
End Sub

Private Sub vsAdvice_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Dim lngW As Long
    
    With vsAdvice
        If Col = col_ҽ������ Then
            .AutoSize col_ҽ������
            .RowHeight(0) = 320
        ElseIf Row = -1 Then
            lngW = Me.TextWidth(.TextMatrix(.FixedRows - 1, Col) & "A")
            If .ColWidth(Col) < lngW Then
                .ColWidth(Col) = lngW
            ElseIf .ColWidth(Col) > .Width * 0.5 Then
                .ColWidth(Col) = .Width * 0.5
            End If
        End If
    End With
End Sub

Private Sub vsAdvice_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = COL_ѡ�� Then Cancel = True
End Sub

Private Sub vsAdvice_DblClick()
    With vsAdvice
        If .MouseCol = COL_ѡ�� And .MouseRow >= .FixedRows And .MouseRow <= .Rows - 1 Then
            Call vsAdvice_KeyPress(32)
        End If
    End With
End Sub

Private Sub vsAdvice_DrawCell(ByVal hDC As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
'˵����1.OwnerDrawҪ����ΪOver(������Ԫ��������)
'      2.Cell��GridLine�������������ڶ��Ǵӵ�1���߿�ʼ
'      3.Cell��Border�������Ǵӵ�2���߿�ʼ,�����Ǵӵ�1���߿�ʼ
    Dim lngLeft As Long, lngRight As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim vRect As RECT
    
    With vsAdvice
        lngLeft = COL_Ƶ��: lngRight = COL_�÷�
        If Not Between(Col, lngLeft, lngRight) Then Exit Sub
        
        If Not RowInһ����ҩ(Row, lngBegin, lngEnd) Then Exit Sub
        
        vRect.Left = Left '������߱����
        vRect.Right = Right - 1 '�����ұ߱����
        If Row = lngBegin Then
            vRect.Top = Bottom - 1 '���б�����������
            vRect.Bottom = Bottom
        Else
            If Row = lngEnd Then
                vRect.Top = Top
                vRect.Bottom = Bottom - 1 '���б����±���
            Else
                vRect.Top = Top
                vRect.Bottom = Bottom
            End If
        End If
        
        If Between(Row, .Row, .RowSel) Then
            SetBkColor hDC, OS.SysColor2RGB(.BackColorSel)
        Else
            SetBkColor hDC, OS.SysColor2RGB(.BackColor)
        End If
        ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0
        Done = True
    End With
End Sub

Private Sub vsAdvice_KeyPress(KeyAscii As Integer)
    Dim i As Long
    With vsAdvice
        If KeyAscii = 13 Then
            KeyAscii = 0
            For i = .Row + 1 To .Rows - 1
                If Not .RowHidden(i) Then
                    .Row = i: Exit For
                End If
            Next
            If i > .Rows - 1 Then .Row = .FixedRows
            Call .ShowCell(.Row, .Col)
        ElseIf KeyAscii = 32 And .Col = COL_ѡ�� Then
            KeyAscii = 0
            If .Cell(flexcpData, .Row, COL_ѡ��) = 0 Then
                If .Cell(flexcpPicture, .Row, COL_ѡ��) Is Nothing Then
                    Set .Cell(flexcpPicture, .Row, COL_ѡ��) = frmIcons.imgTrueFalse.ListImages("T").Picture
                Else
                    Set .Cell(flexcpPicture, .Row, COL_ѡ��) = Nothing
                End If
                Call RowSelectSame(.Row, .Col)
                Call ShowSendTotal
            End If
        End If
    End With
End Sub

Private Sub ShowDefaultRow()
'���ܣ����ڿ��ԼƼ۵�ҽ��,ȱʡ����һ�в�����ȱʡ�Ƽ�ҽ��
'˵����ComboList="#ҽ��ID1;�Ƽ�ҽ��1|#ҽ��ID2;�Ƽ�ҽ��2|..."
'      ���ڵ�һ����ʾ�Ƽ۱�ͻس�������ʱ����
    Dim arrCombo As Variant, lngRow As Long, i As Long
    Dim lngҽ��ID As Long, lng�к� As Long, str�Ƽ�ҽ�� As String
    Dim blnFirst As Boolean, blnHave As Boolean
    
    With vsPrice
        If .ColData(COLP_�Ƽ�ҽ��) <> "" And .Editable <> flexEDNone Then
            arrCombo = Split(.ColData(COLP_�Ƽ�ҽ��), "|")
            
            If Val(.TextMatrix(.Rows - 1, COLP_�к�)) <> 0 _
                And Val(.TextMatrix(.Rows - 1, COLP_�շ�ϸĿID)) <> 0 Then
                '��һ����ʾʱȱʡ����һ��
                blnFirst = True
                .AddItem "", .Rows
                .Row = .Rows - 1
            End If
            lngRow = .Rows - 1
            
            '���ǵ�һ����ʾʱȱʡ�Ƽ�ҽ������һ����ͬ
            If lngRow > 1 And Not blnFirst Then
                If Val(.TextMatrix(lngRow - 1, COLP_�̶�)) = 0 _
                    And Val(.TextMatrix(lngRow - 1, COLP_�к�)) <> 0 Then
                    blnHave = True
                End If
            End If
            For i = 0 To UBound(arrCombo)
                lngҽ��ID = Val(Mid(Mid(arrCombo(i), 1, InStr(arrCombo(i), ";") - 1), 2))
                str�Ƽ�ҽ�� = Replace(arrCombo(i), "#" & lngҽ��ID & ";", "")
                lng�к� = vsAdvice.FindRow(CStr(lngҽ��ID), , COL_ID)
                If blnHave Then
                    If lng�к� = Val(.TextMatrix(lngRow - 1, COLP_�к�)) Then
                        Exit For
                    End If
                Else
                    Exit For
                End If
            Next
                        
            'ģ��ѡ������Ƽ�ҽ��
            .TextMatrix(lngRow, COLP_�к�) = lng�к�
            .TextMatrix(lngRow, COLP_�Ƽ�ҽ��) = str�Ƽ�ҽ��
            .Cell(flexcpData, lngRow, COLP_�Ƽ�ҽ��) = .TextMatrix(lngRow, COLP_�Ƽ�ҽ��)
            
            'ֻ��һ���Ƽ�ҽ��ʱ����ͣ��
            If UBound(arrCombo) = 0 Then
                .Col = COLP_�շ���Ŀ
            Else
                .Col = COLP_�Ƽ�ҽ��
            End If
        End If
        Call .ShowCell(.Row, .Col)
    End With
End Sub

Private Sub vsPrice_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim lngԭ��ID As Long, lngҽ��ID As Long
    Dim int�������� As Integer, intԭ�������� As Integer
    Dim lng�շ�ϸĿID As Long, i As Long
    Dim blnHaveSub As Boolean
    
    On Error GoTo errH
    
    With vsPrice
        If Col = COLP_�Ƽ�ҽ�� Then
            '�������ComboData,TextMatrixȡֵ��ΪComboData
            If .Cell(flexcpTextDisplay, Row, Col) <> .Cell(flexcpData, Row, Col) Then
                lngҽ��ID = .ComboData
                If lngҽ��ID < 0 Then
                    int�������� = Val(Left(Abs(lngҽ��ID), 1))
                    lngҽ��ID = Val(Mid(Abs(lngҽ��ID), 2))
                End If
                lngԭ��ID = Val(vsAdvice.TextMatrix(Val(.TextMatrix(Row, COLP_�к�)), COL_ID))
                intԭ�������� = Val(.TextMatrix(Row, COLP_��������))
                lng�շ�ϸĿID = Val(.TextMatrix(Row, COLP_�շ�ϸĿID))
                                
                '���üƼ�ҽ���Ƿ�������ͬ�շ�ϸĿ
                If lng�շ�ϸĿID <> 0 Then
                    mrsPrice.Filter = "ҽ��ID=" & lngҽ��ID & " And ��������=" & int�������� & " And �շ�ϸĿID=" & lng�շ�ϸĿID
                    If Not mrsPrice.EOF Then
                        MsgBox """" & .Cell(flexcpTextDisplay, Row, Col) & """�Ѿ��������շ���Ŀ""" & .TextMatrix(Row, COLP_�շ���Ŀ) & """��", vbInformation, gstrSysName
                        .TextMatrix(Row, Col) = .Cell(flexcpData, Row, Col): Exit Sub
                    End If
                End If
                
                'ԭ����ҽ������д�������Ҫ����һ��(�����ǹ̶����ɶ���)
                If lngԭ��ID <> 0 Then
                    mrsPrice.Filter = "ҽ��ID=" & lngԭ��ID & " And ��������=" & intԭ�������� & " And ����=1"
                    If mrsPrice.RecordCount = 1 And .TextMatrix(Row, COLP_����) <> "" Then
                        MsgBox """" & .Cell(flexcpData, Row, Col) & """����Ҫ����һ�������Ƽ���Ŀ��", vbInformation, gstrSysName
                        .TextMatrix(Row, Col) = .Cell(flexcpData, Row, Col): Exit Sub
                    End If
                End If
                
                '���������˵ļƼ�ҽ������
                i = vsAdvice.FindRow(CStr(lngҽ��ID), , COL_ID)
                .TextMatrix(Row, COLP_�к�) = i
                .TextMatrix(Row, COLP_��������) = int��������
                .TextMatrix(Row, Col) = .Cell(flexcpTextDisplay, Row, Col)
                .Cell(flexcpData, Row, Col) = .TextMatrix(Row, Col)
                
                If lng�շ�ϸĿID <> 0 Then
                    '��ѡ���ҽ���Ƿ��д�������޸ĺ����Ŀ�Ƿ����
                    mrsPrice.Filter = "ҽ��ID=" & lngҽ��ID & " ��������=" & int�������� & " And ����=1"
                    If Not mrsPrice.EOF Then blnHaveSub = True
                    .TextMatrix(Row, COLP_����) = IIF(blnHaveSub, "��", "")
                
                    '���»����Ӽ�¼������
                    If lngԭ��ID = 0 Then
                        mrsPrice.AddNew '����
                    Else '����
                        mrsPrice.Filter = "ҽ��ID=" & lngԭ��ID & " And ��������=" & intԭ�������� & " And �շ�ϸĿID=" & lng�շ�ϸĿID
                    End If
                    mrsPrice!ҽ��ID = lngҽ��ID
                    If Val(vsAdvice.TextMatrix(i, COL_���ID)) <> 0 Then
                        mrsPrice!���ID = vsAdvice.TextMatrix(i, COL_���ID)
                    Else
                        mrsPrice!���ID = Null
                    End If
                    mrsPrice!�������� = int��������
                    mrsPrice!�շѷ�ʽ = 0
                    If lngԭ��ID = 0 Then
                        mrsPrice!�շ�ϸĿID = lng�շ�ϸĿID
                        mrsPrice!���� = Val(.TextMatrix(Row, COLP_�Ƽ�����))
                        mrsPrice!���� = Val(.TextMatrix(Row, COLP_����))
                        mrsPrice!���� = Val(.TextMatrix(Row, COLP_��������))
                        mrsPrice!��� = Val(.TextMatrix(Row, COLP_���))
                        mrsPrice!�̶� = 0
                    End If
                    mrsPrice!���� = IIF(blnHaveSub, 1, 0)
                    mrsPrice.Update
                    
                    Call ShowAdvicePrice(vsAdvice.Row) '���¼�����ʾ
                End If
            End If
        ElseIf Col = COLP_�շ���Ŀ Or Col = COLP_ִ�п��� Then
            .TextMatrix(Row, Col) = .Cell(flexcpData, Row, Col)
            Call vsPrice_AfterRowColChange(Row, Col, Row, Col) '������ʾ��ť
        ElseIf Col = COLP_�Ƽ����� Then
            If Not IsNumeric(.TextMatrix(Row, Col)) _
                Or Val(.TextMatrix(Row, Col)) <= 0 _
                Or Val(.TextMatrix(Row, Col)) > LONG_MAX Then
                .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col))
                Exit Sub
            End If
            .Cell(flexcpData, Row, Col) = .TextMatrix(Row, Col)
            
            '���¼�¼��
            lngҽ��ID = Val(vsAdvice.TextMatrix(Val(.TextMatrix(Row, COLP_�к�)), COL_ID))
            int�������� = Val(.TextMatrix(Row, COLP_��������))
            lng�շ�ϸĿID = Val(.TextMatrix(Row, COLP_�շ�ϸĿID))
            If lngҽ��ID <> 0 And lng�շ�ϸĿID <> 0 Then
                mrsPrice.Filter = "ҽ��ID=" & lngҽ��ID & " And ��������=" & int�������� & " And �շ�ϸĿID=" & lng�շ�ϸĿID
                mrsPrice!���� = Val(.TextMatrix(Row, Col))
                mrsPrice.Update
                
                Call ShowAdvicePrice(vsAdvice.Row) '���¼�����ʾ
            End If
        ElseIf Col = COLP_���� Then
            If Not IsNumeric(.TextMatrix(Row, Col)) _
                Or Val(.TextMatrix(Row, Col)) <= 0 _
                Or Val(.TextMatrix(Row, Col)) > LONG_MAX Then
                .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col))
                Exit Sub
            End If
            If CheckScope(.Cell(flexcpData, Row, COLP_Ӧ�ս��), .Cell(flexcpData, Row, COLP_ʵ�ս��), .TextMatrix(Row, Col)) <> "" Then
                .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col))
                Exit Sub
            End If
            .TextMatrix(Row, Col) = Format(.TextMatrix(Row, Col), gstrDecPrice)
            .Cell(flexcpData, Row, Col) = .TextMatrix(Row, Col)
            
            '���¼�¼��
            lngҽ��ID = Val(vsAdvice.TextMatrix(Val(.TextMatrix(Row, COLP_�к�)), COL_ID))
            int�������� = Val(.TextMatrix(Row, COLP_��������))
            lng�շ�ϸĿID = Val(.TextMatrix(Row, COLP_�շ�ϸĿID))
            If lngҽ��ID <> 0 And lng�շ�ϸĿID <> 0 Then
                mrsPrice.Filter = "ҽ��ID=" & lngҽ��ID & " And ��������=" & int�������� & " And �շ�ϸĿID=" & lng�շ�ϸĿID
                mrsPrice!���� = Val(.TextMatrix(Row, Col))
                mrsPrice.Update
                
                Call ShowAdvicePrice(vsAdvice.Row) '���¼�����ʾ
            End If
        End If
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vsPrice_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim lngRow As Long
    
    '���ݿɷ�༭����
    If Not CellEditable(NewRow, NewCol) Then
        vsPrice.ComboList = ""
        vsPrice.FocusRect = flexFocusLight
    Else
        vsPrice.FocusRect = flexFocusSolid
        If NewCol = COLP_�Ƽ�ҽ�� Then
            vsPrice.ComboList = vsPrice.ColData(NewCol)
        ElseIf NewCol = COLP_�շ���Ŀ Or NewCol = COLP_ִ�п��� Then
            vsPrice.ComboList = "..."
        Else
            vsPrice.ComboList = ""
        End If
    End If
        
    If NewRow <> OldRow Then
        '��ʾҩƷ���
        With vsPrice
            stbThis.Panels(2).Text = ""
            lngRow = Val(.TextMatrix(NewRow, COLP_�к�))
            If lngRow <> 0 And .TextMatrix(NewRow, COLP_�շ����) <> "" Then
                If InStr(",5,6,7,", .TextMatrix(NewRow, COLP_�շ����)) > 0 _
                    Or .TextMatrix(NewRow, COLP_�շ����) = "4" And Val(.TextMatrix(NewRow, COLP_��������)) = 1 Then
                    '��ʾҩƷ���������ĵĿ��
                    If InStr(",5,6,7,", vsAdvice.TextMatrix(lngRow, COL_�������)) > 0 Then
                        If InStr(GetInsidePrivs(pסԺҽ���´�), "��ʾҩƷ���") = 0 Then
                            stbThis.Panels(2).Text = vsAdvice.TextMatrix(lngRow, col_ҽ������) & "," & vsAdvice.TextMatrix(lngRow, COL_ִ�п���) & IIF(Val(vsAdvice.TextMatrix(lngRow, COL_���)) > 0, "�п��", "�޿��")
                        Else
                            stbThis.Panels(2).Text = vsAdvice.TextMatrix(lngRow, col_ҽ������) & "," & vsAdvice.TextMatrix(lngRow, COL_ִ�п���) & "���ÿ��:" & FormatEx(Val(vsAdvice.TextMatrix(lngRow, COL_���)), 5) & vsAdvice.TextMatrix(lngRow, COL_סԺ��λ)
                        End If
                    Else
                        'ͬһ������ȡ:ҩƷ��סԺ��λ,���İ��ۼ۵�λ
                        If InStr(GetInsidePrivs(pסԺҽ���´�), "��ʾҩƷ���") = 0 Then
                            If GetStock(Val(.TextMatrix(NewRow, COLP_�շ�ϸĿID)), Val(.TextMatrix(NewRow, COLP_ִ�п���ID))) > 0 Then
                                stbThis.Panels(2).Text = .TextMatrix(NewRow, COLP_�շ���Ŀ) & "," & .TextMatrix(NewRow, COLP_ִ�п���) & "�п��"
                            Else
                                stbThis.Panels(2).Text = .TextMatrix(NewRow, COLP_�շ���Ŀ) & "," & .TextMatrix(NewRow, COLP_ִ�п���) & "�޿��"
                            End If
                        Else
                            stbThis.Panels(2).Text = .TextMatrix(NewRow, COLP_�շ���Ŀ) & "," & .TextMatrix(NewRow, COLP_ִ�п���) & "���ÿ��:" & _
                                FormatEx(GetStock(Val(.TextMatrix(NewRow, COLP_�շ�ϸĿID)), Val(.TextMatrix(NewRow, COLP_ִ�п���ID))), 5) & .TextMatrix(NewRow, COLP_��λ)
                        End If
                    End If
                End If
            End If
        End With
        
        '��ʾҽ������
        stbThis.Panels(3).Text = Getҽ������(NewRow)
    End If
End Sub

Private Function Getҽ������(ByVal lngRow As Long) As String
'���ܣ���ȡָ���еķ�������
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, str���� As String
    
    With vsPrice
        If Val(.TextMatrix(lngRow, COLP_�շ�ϸĿID)) <> 0 Then
            strSQL = "Select N.���� From ����֧����Ŀ M,����֧������ N Where M.�շ�ϸĿID=[1] And M.����ID=N.ID And M.����=[2]"
            On Error GoTo errH
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.TextMatrix(lngRow, COLP_�շ�ϸĿID)), mint����)
            If Not rsTmp.EOF Then str���� = NVL(rsTmp!����)
        End If
    End With
    Getҽ������ = IIF(str���� <> "", "ҽ������:" & str����, "")
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function RowInһ����ҩ(ByVal lngRow As Long, lngBegin As Long, lngEnd As Long) As Boolean
'���ܣ��ж�ָ�����Ƿ���һ����ҩ�ķ�Χ��,�����,ͬʱ�����кŷ�Χ
'˵�������ص��кŷ�Χ��������ҩ;�����к�
    Dim i As Long, blnTmp As Boolean
    With vsAdvice
        If .TextMatrix(lngRow, COL_�������) = "" Then Exit Function
        If InStr(",5,6,", .TextMatrix(lngRow, COL_�������)) = 0 Then Exit Function
        If Val(.TextMatrix(lngRow - 1, COL_���ID)) = Val(.TextMatrix(lngRow, COL_���ID)) Then
            blnTmp = True
        ElseIf lngRow + 1 <= .Rows - 1 Then
            If Val(.TextMatrix(lngRow + 1, COL_���ID)) = Val(.TextMatrix(lngRow, COL_���ID)) Then
                blnTmp = True
            End If
        End If
        If blnTmp Then
            lngBegin = lngRow
            For i = lngRow - 1 To .FixedRows Step -1
                If Val(.TextMatrix(i, COL_���ID)) = Val(.TextMatrix(lngRow, COL_���ID)) Then
                    lngBegin = i
                Else
                    Exit For
                End If
            Next
            lngEnd = lngRow
            For i = lngRow + 1 To .Rows - 1
                If Val(.TextMatrix(i, COL_���ID)) = Val(.TextMatrix(lngRow, COL_���ID)) Then
                    lngEnd = i
                Else
                    Exit For
                End If
            Next
        End If
        RowInһ����ҩ = blnTmp
    End With
End Function

Private Sub InitAdviceTable()
'���ܣ���ʼ���嵥��ʽ
    Dim arrHead As Variant, strHead As String, i As Long
    
    strHead = ",300,4;" & _
        "Ӥ��,550,1;ҽ������,3000,1;����,600,7;��λ,450,1;����,600,7;��λ,450,1;���,850,7;" & _
        "Ƶ��,1000,1;�÷�,1000,1;ҽ������,1500,1;ִ��ʱ��,1000,1;ִ�п���,850,1;ִ������,850,1;" & _
        "ID;���ID;ҽ��״̬;���˿���ID;��������ID;����ҽ��;����ʱ��;�������;������ĿID;�걾��λ;��鷽��;ִ�б��;�Ƽ�����;" & _
        "ִ������ID;ִ�п���ID;��������;�Թܱ���;�շ�ϸĿID;����ϵ��;סԺ��װ;סԺ��λ;�ɷ����;���;����;�ֽ�ʱ��;�״�ʱ��;" & _
        "ĩ��ʱ��;ǩ����;������־;���㷽ʽ;ִ�а���;��ʼʱ��;�������;ִ�з���;��������;�������;��ҩ����"
    arrHead = Split(strHead, ";")
    With vsAdvice
        .Clear
        .FixedRows = 1
        .FixedCols = 0
        .Cols = .FixedCols + UBound(arrHead) + 1
        .Rows = .FixedRows + 1
        
        For i = 0 To UBound(arrHead)
            .TextMatrix(.FixedRows - 1, .FixedCols + i) = Split(arrHead(i), ",")(0)
            If UBound(Split(arrHead(i), ",")) > 0 Then
                .ColWidth(.FixedCols + i) = Val(Split(arrHead(i), ",")(1))
                .ColAlignment(.FixedCols + i) = Val(Split(arrHead(i), ",")(2))
            Else
                .ColHidden(.FixedCols + i) = True
            End If
        Next
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = 4
        .FrozenCols = COL_ѡ�� + 1 - .FixedCols
        .RowHeight(0) = 320
    End With
End Sub

Private Sub InitPriceTable()
'���ܣ���ʼ���Ƽ��嵥��ʽ
    Dim arrHead As Variant, strHead As String, i As Long
    
    strHead = "�к�;�շ�ϸĿID;�̶�;���;�Ƽ�ҽ��,2000,1;���,650,1;�շ���Ŀ,2000,1;�Ƽ�����,900,7;" & _
        "����,800,7;��λ,500,1;����,1000,7;Ӧ�ս��,1050,7;ʵ�ս��,1050,7;ִ�п���,1000,1;��������,850,1;" & _
        "����,450,4;�շѷ�ʽ,1500,1;�շ����;ִ�п���ID;��������;��������"
    arrHead = Split(strHead, ";")
    With vsPrice
        .Clear
        .FixedRows = 1
        .FixedCols = 0
        .Cols = .FixedCols + UBound(arrHead) + 1
        .Rows = .FixedRows + 1
        
        For i = 0 To UBound(arrHead)
            .TextMatrix(.FixedRows - 1, .FixedCols + i) = Split(arrHead(i), ",")(0)
            If UBound(Split(arrHead(i), ",")) > 0 Then
                .ColWidth(.FixedCols + i) = Val(Split(arrHead(i), ",")(1))
                .ColAlignment(.FixedCols + i) = Val(Split(arrHead(i), ",")(2))
            Else
                .ColHidden(.FixedCols + i) = True
            End If
        Next
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = 4
    End With
End Sub

Private Sub DeleteCurRow(ByVal lngRow As Long, Optional ByVal blnDelCur As Boolean = True)
'���ܣ��ڴ���������嵥�Ĺ�����ɾ������������(��ҩ�ƻ��ҩ)
'������blnDelCur=�Ƿ�ɾ����ǰ��
    Dim lngҽ��ID As Long, lng���ID As Long, i As Long
    
    With vsAdvice
        lngҽ��ID = Val(.TextMatrix(lngRow, COL_ID))
        lng���ID = Val(.TextMatrix(lngRow, COL_���ID))
                
        'ɾ����ǰ��
        If blnDelCur Then .RemoveItem lngRow
        
        'ɾ�������
        If lng���ID <> 0 Then
            For i = .Rows - 1 To .FixedRows Step -1
                If Val(.TextMatrix(i, COL_���ID)) = lng���ID _
                    Or Val(.TextMatrix(i, COL_ID)) = lng���ID Then
                    .RemoveItem i
                End If
            Next
        Else
            For i = .Rows - 1 To .FixedRows Step -1
                If Val(.TextMatrix(i, COL_���ID)) = lngҽ��ID Then
                    .RemoveItem i
                End If
            Next
        End If
    End With
End Sub

Private Sub InitSeekSet(rsSeek As ADODB.Recordset)
'���ܣ���ʼ�����ڻ��ܼ����ۿ۵���ʱ��¼��
    Set rsSeek = New ADODB.Recordset
    rsSeek.Fields.Append "��������", adInteger
    rsSeek.Fields.Append "�����ǩ", adVariant
    rsSeek.Fields.Append "������ID", adBigInt
    rsSeek.Fields.Append "�ϼ�", adCurrency, , adFldIsNullable
    rsSeek.CursorLocation = adUseClient
    rsSeek.LockType = adLockOptimistic
    rsSeek.CursorType = adOpenStatic
    rsSeek.Open
End Sub

Private Sub InitPriceRecordset()
'���ܣ���ʼ��ҽ���Ƽۼ�¼��
    Set mrsPrice = New ADODB.Recordset
    
    mrsPrice.Fields.Append "ҽ��ID", adBigInt
    mrsPrice.Fields.Append "���ID", adBigInt, , adFldIsNullable
    mrsPrice.Fields.Append "��������", adInteger, , adFldIsNullable
    mrsPrice.Fields.Append "�շѷ�ʽ", adInteger, , adFldIsNullable
    mrsPrice.Fields.Append "�շ����", adVarChar, 1
    mrsPrice.Fields.Append "�շ�ϸĿID", adBigInt
    mrsPrice.Fields.Append "ִ�п���ID", adBigInt, , adFldIsNullable
    mrsPrice.Fields.Append "����", adDouble
    mrsPrice.Fields.Append "����", adDouble, , adFldIsNullable '��ۼ۸�
    mrsPrice.Fields.Append "����", adInteger '�����Ƿ��������
    mrsPrice.Fields.Append "���", adInteger
    mrsPrice.Fields.Append "����", adInteger
    mrsPrice.Fields.Append "�̶�", adInteger
    
    mrsPrice.CursorLocation = adUseClient
    mrsPrice.LockType = adLockOptimistic
    mrsPrice.CursorType = adOpenStatic
    mrsPrice.Open
End Sub

Private Sub InitRecordSet(rsSQL As ADODB.Recordset, rsTotal As ADODB.Recordset, rsUpload As ADODB.Recordset, _
    rsNumber As ADODB.Recordset, rsMoneyNow As ADODB.Recordset, rsItems As ADODB.Recordset)
'��ʼ����¼��
    'SQL��¼��
    Set rsSQL = New ADODB.Recordset
    rsSQL.Fields.Append "����", adInteger '1-�Ƽ�,2-ǩ��,3-У��,4-����,5-����,6-����
    rsSQL.Fields.Append "ҽ��ID", adBigInt 'һ��ҽ����ID
    rsSQL.Fields.Append "��ĿID", adBigInt '�շ�ϸĿID
    rsSQL.Fields.Append "���", adBigInt '��������
    rsSQL.Fields.Append "SQL", adVarChar, 5000 'SQL
    rsSQL.Fields.Append "NO", adVarChar, 30, adFldIsNullable '����NO�滻����ʱ����
    rsSQL.CursorLocation = adUseClient
    rsSQL.LockType = adLockOptimistic
    rsSQL.CursorType = adOpenStatic
    rsSQL.Open
    
    '�Ƽ������ۼƼ�¼��
    Set rsTotal = New ADODB.Recordset
    rsTotal.Fields.Append "ҽ��ID", adBigInt 'һ��ҽ����ID
    rsTotal.Fields.Append "��ĿID", adBigInt
    rsTotal.Fields.Append "�ⷿID", adBigInt
    rsTotal.Fields.Append "����", adDouble
    rsTotal.CursorLocation = adUseClient
    rsTotal.LockType = adLockOptimistic
    rsTotal.CursorType = adOpenStatic
    rsTotal.Open
    
    'ҽ���ϴ����ʵ�
    Set rsUpload = New ADODB.Recordset
    rsUpload.Fields.Append "ҽ��ID", adBigInt 'һ��ҽ����ID
    rsUpload.Fields.Append "NO", adVarChar, 30
    rsUpload.CursorLocation = adUseClient
    rsUpload.LockType = adLockOptimistic
    rsUpload.CursorType = adOpenStatic
    rsUpload.Open
    
    '��¼�Թܱ���
    Set rsNumber = New ADODB.Recordset
    rsNumber.Fields.Append "����", adVarChar, 18
    rsNumber.Fields.Append "���ID", adBigInt
    rsNumber.Fields.Append "��������", adVarChar, 18
    rsNumber.Fields.Append "ִ�п���ID", adVarChar, 18
    rsNumber.Fields.Append "������ĿID", adVarChar, 18
    rsNumber.Fields.Append "Ӥ��", adBigInt
    rsNumber.Fields.Append "������־", adBigInt
    rsNumber.Fields.Append "�걾", adVarChar, 18
    rsNumber.Fields.Append "�ɼ�����ID", adBigInt
    rsNumber.CursorLocation = adUseClient
    rsNumber.LockType = adLockOptimistic
    rsNumber.CursorType = adOpenStatic
    rsNumber.Open
    
    '��ǰ���˱���Ҫ���͵ķ���
    Set rsMoneyNow = New ADODB.Recordset
    rsMoneyNow.Fields.Append "ҽ��ID", adBigInt 'һ��ҽ����ID
    rsMoneyNow.Fields.Append "������ĿID", adBigInt
    rsMoneyNow.Fields.Append "�շ���ĿID", adBigInt
    rsMoneyNow.Fields.Append "�Թܱ���", adVarChar, 18, adFldIsNullable
    rsMoneyNow.Fields.Append "��������", adVarChar, 50, adFldIsNullable
    rsMoneyNow.Fields.Append "�շѷ�ʽ", adInteger
    rsMoneyNow.Fields.Append "�շ�ʱ��", adVarChar, 10
    rsMoneyNow.Fields.Append "ִ�в���ID", adBigInt
    rsMoneyNow.Fields.Append "��ҽ��ID", adBigInt '���ID��Ϊ�յ�ҽ���е�ҽ��ID
    rsMoneyNow.Fields.Append "��鲿λ", adVarChar, 100
    rsMoneyNow.Fields.Append "��鷽��", adVarChar, 100
    rsMoneyNow.Fields.Append "����", adDouble '�շ�����
    
    rsMoneyNow.CursorLocation = adUseClient
    rsMoneyNow.LockType = adLockOptimistic
    rsMoneyNow.CursorType = adOpenStatic
    rsMoneyNow.Open
    
    '��ǰ���˱��η��͵ķ�����Ŀ����
    Set rsItems = New ADODB.Recordset
    rsItems.Fields.Append "����ID", adBigInt
    rsItems.Fields.Append "��ҳID", adBigInt, , adFldIsNullable
    rsItems.Fields.Append "ҽ��ID", adBigInt
    rsItems.Fields.Append "�շ����", adVarChar, 1
    rsItems.Fields.Append "�շ�ϸĿID", adBigInt
    rsItems.Fields.Append "����", adDouble
    rsItems.Fields.Append "����", adDouble
    rsItems.Fields.Append "ʵ�ս��", adDouble
    rsItems.Fields.Append "������", adVarChar, 100, adFldIsNullable
    rsItems.Fields.Append "��������", adVarChar, 100, adFldIsNullable
    rsItems.CursorLocation = adUseClient
    rsItems.LockType = adLockOptimistic
    rsItems.CursorType = adOpenStatic
    rsItems.Open
End Sub

Private Function LoadAdvicePrice(ByVal lngRow As Long, rsSend As ADODB.Recordset, cur�ϼ� As Currency) As Boolean
'���ܣ���ȡָ��ҽ��(����ǰ��)�ļƼ۹�ϵ����ʱ��¼��,������ȱʡ���ͽ��(���ѱ����)
'���أ�cur�ϼ�=�������ҽ�����ͽ��(��ҩ���δ��,��Ҫ����۸�����)
    Dim rsTmp As New ADODB.Recordset
    Dim rsCur As New ADODB.Recordset
    Dim strSQL As String, strPrice As String
    Dim str�������� As String, arr�������� As Variant
    Dim blnDo As Boolean, i As Long, k As Long
    Dim dbl���� As Double, dbl���� As Double, dblӦ�� As Double
    Dim curӦ�� As Currency, curʵ�� As Currency
    Dim bln�������� As Boolean, lng��ĿID As Long
    Dim lng������ID As Long, blnHaveSub As Boolean
    Dim lngִ�п���ID As Long, cur��� As Currency
    Dim lng����ID As Long
    
    On Error GoTo errH
    
    cur��� = 0
    With vsAdvice
        If InStr(",4,5,6,7,", rsSend!�������) > 0 Then
            '��ΪԺ��ִ��(�Ա�ҩ),ҩƷ������Ϊ����,�ҹ̶������Ƽ�
            If NVL(rsSend!ִ������, 0) <> 5 Then
                mrsPrice.AddNew
                mrsPrice!ҽ��ID = rsSend!ID
                mrsPrice!���ID = rsSend!���ID
                mrsPrice!�������� = 0
                mrsPrice!�շѷ�ʽ = 0
                mrsPrice!�շ���� = rsSend!�������
                mrsPrice!�շ�ϸĿID = rsSend!�շ�ϸĿID
                mrsPrice!ִ�п���ID = rsSend!ִ�п���ID
                mrsPrice!���� = 1
                mrsPrice!���� = NVL(rsSend!��������, 0)
                mrsPrice!��� = NVL(rsSend!�Ƿ���, 0)
                mrsPrice!�̶� = 1
                mrsPrice!���� = 0
                                
                '���͵���������
                If rsSend!������� = "7" Then
                    '��ҩҩ����λ�����ɷ��㴦��:ÿ��
                    If NVL(rsSend!�ɷ����, 0) = 0 Then
                        dbl���� = Val(.TextMatrix(lngRow, COL_����)) * Val(.TextMatrix(lngRow, COL_����)) / NVL(rsSend!����ϵ��, 1)
                    Else
                        dbl���� = Val(.TextMatrix(lngRow, COL_����)) _
                            * IntEx(Val(.TextMatrix(lngRow, COL_����)) / NVL(rsSend!����ϵ��, 1) / NVL(rsSend!סԺ��װ, 1)) * NVL(rsSend!סԺ��װ, 1)
                    End If
                Else
                    dbl���� = Val(.TextMatrix(lngRow, COL_����)) * NVL(rsSend!סԺ��װ, 1)
                End If
                dbl���� = Format(dbl����, "0.00000")
                                
                '��¼�ۼ۵���
                If NVL(rsSend!�Ƿ���, 0) = 0 Or rsSend!������� = "4" And NVL(rsSend!��������, 0) = 0 Then
                    mrsPrice!���� = Format(CalcPrice(rsSend!�շ�ϸĿID, , , True, , , mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�), gstrDecPrice)
                Else '���ۼۼ���ҩƷʱ��,�Ա�ҩʱ�޶�Ӧҩ��
                    mrsPrice!���� = Format(CalcDrugPrice(rsSend!�շ�ϸĿID, NVL(rsSend!ִ�п���ID, 0), dbl����, , True, , mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�), gstrDecPrice)
                End If
                mrsPrice.Update
                                
                '����ҽ�����ͽ��(���ѱ���۵�ʵ�ս��)
                If Not IsNull(mrsPati!�ѱ�) Then
                    If NVL(rsSend!�Ƿ���, 0) = 0 Or rsSend!������� = "4" And NVL(rsSend!��������, 0) = 0 Then
                        cur��� = Format(CalcPrice(rsSend!�շ�ϸĿID, mrsPati!�ѱ�, dbl����, , NVL(rsSend!ִ�п���ID, 0), , mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�), gstrDec)
                    Else
                        cur��� = Format(CalcDrugPrice(rsSend!�շ�ϸĿID, NVL(rsSend!ִ�п���ID, 0), dbl����, mrsPati!�ѱ�, , , mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�), "0.00000")
                    End If
                Else
                    If gbln�Ӱ�Ӽ� Then
                        '����Ӱ�Ӽ�
                        If NVL(rsSend!�Ƿ���, 0) = 0 Or rsSend!������� = "4" And NVL(rsSend!��������, 0) = 0 Then
                            dbl���� = Format(CalcPrice(rsSend!�շ�ϸĿID, , , , , , mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�), gstrDecPrice)
                        Else '���ۼۼ���ҩƷʱ��,�Ա�ҩʱ�޶�Ӧҩ��
                            dbl���� = Format(CalcDrugPrice(rsSend!�շ�ϸĿID, NVL(rsSend!ִ�п���ID, 0), dbl����, , , , mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�), gstrDecPrice)
                        End If
                        cur��� = Format(mrsPrice!���� * dbl���� * dbl����, gstrDec)
                    Else
                        cur��� = Format(mrsPrice!���� * dbl���� * mrsPrice!����, gstrDec)
                    End If
                End If
            End If
            
            cur�ϼ� = cur���
        Else
            'ȡ�����շ� ��ϵ�еĶ���(����ʱ�Ŷ��Ƽ�):�����Ƽ�,��Ϊ������Ժ��ִ��
            If NVL(rsSend!�Ƽ�����, 0) = 0 And InStr(",0,5,", NVL(rsSend!ִ������, 0)) = 0 Then
                dbl���� = Format(Val(.TextMatrix(lngRow, COL_����)), "0.00000")
                bln�������� = (rsSend!������� = "F" And Not IsNull(rsSend!���ID))
                
                '���ֶ�Ӧ�ļƼ����
                If Not IsNull(rsSend!�걾��λ) And Not IsNull(rsSend!��鷽��) Then
                    strPrice = " And ��鲿λ=[4] And ��鷽��=[5] And Nvl(��������,0)=0"
                ElseIf NVL(rsSend!ִ�б��, 0) = 0 Then
                    strPrice = " And ��鲿λ Is Null And ��鷽�� is Null And Nvl(��������,0)=0"
                Else 'Ŀǰ�������Ի����м��յ����
                    strPrice = " And ��鲿λ Is Null And ��鷽�� is Null And Nvl(��������,0) IN(0,1)"
                End If
                                
                strPrice = "Select �շ���ĿID,���ж��� From (" & _
                    " Select c.�շ���ĿID, c.���ж���, c.���ÿ���id" & _
                    "   ,Max(Nvl(c.���ÿ���id, 0)) Over(Partition By c.������Ŀid, c.��鲿λ, c.��鷽��, c.��������) As Top" & _
                    " From �����շѹ�ϵ C Where C.������ĿID=[2]" & strPrice & _
                    "       And (C.���ÿ���ID is Null And C.������Դ = 0 or C.���ÿ���ID = Decode([3],0,[6],[3]) And C.������Դ = " & IIF(mbytSendKind = EInBilling, 2, 1) & ")" & _
                    " ) Where Nvl(���ÿ���id, 0) = Top"
                    
                '�ȶ�ȡ���еļƼ�
                strSQL = _
                    " Select C.���,A.�շ�ϸĿID as �շ���ĿID,A.���� as �շ�����,Nvl(E.���ж���,0) as ���ж���," & _
                    " B.������ĿID,C.�Ӱ�Ӽ�,B.�Ӱ�Ӽ���,Decode(C.�Ƿ���,1,A.����,B.�ּ�)" & IIF(bln��������, "*Nvl(B.�����շ���,100)/100", "") & " as ����," & _
                    " C.�Ƿ���,Nvl(A.����,0) as ����,D.��������,Nvl(A.ִ�п���ID,[3]) as ִ�п���ID,C.���ηѱ�," & _
                    " Nvl(A.��������,0) as ��������,Nvl(A.�շѷ�ʽ,0) as �շѷ�ʽ" & _
                    " From ����ҽ���Ƽ� A,�շѼ�Ŀ B,�շ���ĿĿ¼ C,�������� D,(" & strPrice & ")  E" & _
                    " Where A.ҽ��ID=[1] And A.�շ�ϸĿID=0+E.�շ���ĿID(+)" & _
                    " And A.�շ�ϸĿID=B.�շ�ϸĿID And A.�շ�ϸĿID=C.ID And A.�շ�ϸĿID=D.����ID(+)" & _
                    GetPriceGradeSQL(mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�, "C", "B", "7", "8", "9") & _
                    " And C.������� IN(" & IIF(mlng�������� = 1, 1, 2) & ",3) And (C.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
                    " And ((Sysdate Between B.ִ������ and B.��ֹ����) or (Sysdate>=B.ִ������ And B.��ֹ���� is NULL))" & _
                    " Order by ��������,����,A.�շ�ϸĿID"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(rsSend!ID), Val(rsSend!������ĿID), _
                    Val(NVL(rsSend!ִ�п���ID, 0)), CStr(NVL(rsSend!�걾��λ)), CStr(NVL(rsSend!��鷽��)), mlng���˲���ID, mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�)
                
                'û�����ȡĬ�ϵļƼۣ�ֻ�����δУ�Ե�ҽ��
                If rsTmp.EOF And rsSend!ҽ��״̬ = 1 Then
                    lng����ID = 0 '�����Թܷ���,ֻ��ȡ�Թܶ�Ӧ�����ķ���
                    If .TextMatrix(lngRow, COL_�Թܱ���) <> "" Then
                        lng����ID = GetTubeMaterial(.TextMatrix(lngRow, COL_�Թܱ���))
                    End If
                
                    '���ֶ�Ӧ�ļƼ����
                    If Not IsNull(rsSend!�걾��λ) And Not IsNull(rsSend!��鷽��) Then
                        strPrice = " And c.��鲿λ=[3] And c.��鷽��=[4] And Nvl(c.��������,0)=0"
                    ElseIf NVL(rsSend!ִ�б��, 0) = 0 Then
                        strPrice = " And c.��鲿λ Is Null And c.��鷽�� is Null And Nvl(c.��������,0)=0"
                    Else 'Ŀǰ�������Ի����м��յ����
                        strPrice = " And c.��鲿λ Is Null And c.��鷽�� is Null And Nvl(c.��������,0) IN(0,1)"
                    End If
                    
                    strPrice = "Select * From (" & _
                        "Select C.������ĿID,C.�շ���ĿID,C.��鲿λ,C.��鷽��,C.��������,C.�շ�����,C.���ж���,C.������Ŀ,C.�շѷ�ʽ,c.���ÿ���id" & _
                        " ,Max(Nvl(c.���ÿ���id, 0)) Over(Partition By c.������Ŀid, c.��鲿λ, c.��鷽��, c.��������) As Top" & _
                        " From �����շѹ�ϵ C Where C.������ĿID=[1]" & strPrice & _
                        "      And (C.���ÿ���ID is Null And C.������Դ = 0 or C.���ÿ���ID = Decode([2],0,[6],[2]) And C.������Դ = " & IIF(mbytSendKind = EInBilling, 2, 1) & ")" & _
                        " ) Where Nvl(���ÿ���id, 0) = Top"
                        
                    strSQL = _
                        " Select C.���,A.�շ���ĿID,A.�շ�����,A.���ж���,B.������ĿID," & _
                        " C.�Ӱ�Ӽ�,B.�Ӱ�Ӽ���,Decode(C.�Ƿ���,1,B.ȱʡ�۸�,B.�ּ�)" & IIF(bln��������, "*Nvl(B.�����շ���,100)/100", "") & " as ����," & _
                        " C.�Ƿ���,Nvl(A.������Ŀ,0) as ����,D.��������,[2] as ִ�п���ID,C.���ηѱ�," & _
                        " Nvl(A.��������,0) as ��������,Nvl(A.�շѷ�ʽ,0) as �շѷ�ʽ" & _
                        " From (" & strPrice & ") A,�շѼ�Ŀ B,�շ���ĿĿ¼ C,�������� D" & _
                        " Where A.�շ���ĿID=B.�շ�ϸĿID And A.�շ���ĿID=C.ID And A.�շ���ĿID=D.����ID(+)" & _
                        GetPriceGradeSQL(mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�, "C", "B", "7", "8", "9") & _
                        " And C.������� IN(" & IIF(mlng�������� = 1, 1, 2) & ",3) And (C.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
                        " And (C.վ��='" & gstrNodeNo & "' Or C.վ�� is Null)" & _
                        " And ((Sysdate Between B.ִ������ and B.��ֹ����) or (Sysdate>=B.ִ������ And B.��ֹ���� is NULL))" & _
                        " And (Nvl(A.�շѷ�ʽ,0)=1 And C.���='4' And A.�շ���ĿID=[5] Or Not(Nvl(A.�շѷ�ʽ,0)=1 And C.���='4' And [5]<>0))" & _
                        " Order by ��������,����,A.�շ���ĿID"
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(rsSend!������ĿID), _
                        Val(NVL(rsSend!ִ�п���ID, 0)), CStr(NVL(rsSend!�걾��λ)), CStr(NVL(rsSend!��鷽��)), lng����ID, mlng���˲���ID, mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�)
                End If
                
                'ȷ���Ƽ�֮���Ƿ���������Լ���������ID
                arr�������� = Array()
                If Not rsTmp.EOF Then
                    Do While Not rsTmp.EOF
                        If InStr(str�������� & ",", "," & rsTmp!�������� & ",") = 0 Then
                            str�������� = str�������� & "," & rsTmp!��������
                        End If
                        rsTmp.MoveNext
                    Loop
                    arr�������� = Split(Mid(str��������, 2), ",")
                End If
                                
                For k = 0 To UBound(arr��������)
                    rsTmp.Filter = "��������=" & arr��������(k)
                    
                    lng��ĿID = 0: cur��� = 0
                    lng������ID = 0: blnHaveSub = False
                    If Not rsTmp.EOF And gbln��������ۿ� Then
                        Do While Not rsTmp.EOF
                            If NVL(rsTmp!����, 0) = 0 Then
                                'SQL����������ǰ��,ֻȡ����Ŀ�ĵ�һ������
                                If lng������ID = 0 Then lng������ID = rsTmp!������ĿID
                            ElseIf NVL(rsTmp!����, 0) = 1 Then
                                blnHaveSub = True: Exit Do
                            End If
                            rsTmp.MoveNext
                        Loop
                        rsTmp.MoveFirst
                    End If
                    
                    Do While True
                        blnDo = False
                        If rsTmp.EOF Then
                            If lng��ĿID <> 0 Then blnDo = True
                        Else
                            If rsTmp!�շ���ĿID <> lng��ĿID And lng��ĿID <> 0 Then blnDo = True
                        End If
                        If blnDo Then
                            If Not IsNull(mrsPrice!����) Then
                                mrsPrice!���� = Format(mrsPrice!����, gstrDecPrice)
                            End If
                            mrsPrice.Update
                            
                            'ҽ�����ͽ��
                            cur��� = cur��� + Format(curʵ��, gstrDec)
                        End If
                        If rsTmp.EOF Then Exit Do
                        
                        '------------------------------------
                        If rsTmp!�շ���ĿID <> lng��ĿID Then
                            curʵ�� = 0
                            mrsPrice.AddNew
                            mrsPrice!ҽ��ID = rsSend!ID
                            mrsPrice!���ID = rsSend!���ID
                            mrsPrice!�������� = NVL(rsTmp!��������, 0)
                            mrsPrice!�շѷ�ʽ = NVL(rsTmp!�շѷ�ʽ, 0)
                            mrsPrice!�շ���� = rsTmp!���
                            mrsPrice!�շ�ϸĿID = rsTmp!�շ���ĿID
                            mrsPrice!���� = NVL(rsTmp!�շ�����, 0)
                            mrsPrice!���� = NVL(rsTmp!��������, 0)
                            mrsPrice!��� = NVL(rsTmp!�Ƿ���, 0)
                            mrsPrice!�̶� = NVL(rsTmp!���ж���, 0)
                            mrsPrice!���� = NVL(rsTmp!����, 0)
                            
                            If rsSend!������� = "E" And rsSend!�������� = "1" And rsSend!ִ�з��� = 5 And InStr(",5,6,", rsTmp!���) > 0 Then
                                'ԭҺƤ�����⡣�󶨵�ҩƷ�������û��ָ��������ԭ���߼�
                                If Val(rsSend!��ҩ���� & "") <> 0 Then
                                    lngִ�п���ID = Val(rsSend!��ҩ���� & "")
                                Else
                                    lngִ�п���ID = NVL(rsTmp!ִ�п���ID, 0)
                                End If
                                lngִ�п���ID = Get�շ�ִ�п���ID(mlng����ID, 0, rsTmp!���, rsTmp!�շ���ĿID, 4, NVL(rsSend!���˿���id, 0), 0, 2, lngִ�п���ID, , , 2)
                            Else
                                'ִ�п���:��ҩ��ҩƷ���������ĵ�ר��ȡ
                                lngִ�п���ID = NVL(rsTmp!ִ�п���ID, 0)
                                If rsTmp!��� = "4" And NVL(rsTmp!��������, 0) = 1 Or InStr(",5,6,7,", rsTmp!���) > 0 Then
                                    lngִ�п���ID = Get�շ�ִ�п���ID(mlng����ID, 0, rsTmp!���, rsTmp!�շ���ĿID, 4, NVL(rsSend!���˿���id, 0), 0, 2, lngִ�п���ID, , , 2)
                                End If
                            End If
                            If lngִ�п���ID <> 0 Then
                                mrsPrice!ִ�п���ID = lngִ�п���ID
                            Else
                                mrsPrice!ִ�п���ID = Null
                            End If
                        End If
                        lng��ĿID = rsTmp!�շ���ĿID
                        
                        '���㵥�ۺ�ʵ��
                        If NVL(rsTmp!�Ƿ���, 0) = 1 And InStr(",5,6,7,", rsTmp!���) > 0 Then
                            '��ҩ��ҩƷ�Ƽ۰�ʱ�ۼ���(��һ������),���������Ҫ��ҽ������
                            mrsPrice!���� = CalcDrugPrice(rsTmp!�շ���ĿID, NVL(mrsPrice!ִ�п���ID, 0), dbl���� * NVL(rsTmp!�շ�����, 0), , True, , mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�)
                            
                            dblӦ�� = Format(mrsPrice!���� * dbl����, "0.00000") * Format(mrsPrice!����, gstrDecPrice)
                            
                            '����Ӱ�Ӽ�
                            If gbln�Ӱ�Ӽ� And NVL(rsTmp!�Ӱ�Ӽ�, 0) = 1 Then
                                dblӦ�� = dblӦ�� * (1 + NVL(rsTmp!�Ӱ�Ӽ���, 0) / 100)
                            End If
    
                            curӦ�� = Format(dblӦ��, gstrDec)
                            
                            If Not IsNull(mrsPati!�ѱ�) And Not (gbln��������ۿ� And blnHaveSub) And NVL(rsTmp!���ηѱ�, 0) = 0 Then
                                curʵ�� = curʵ�� + Format(ActualMoney(mrsPati!�ѱ�, rsTmp!������ĿID, curӦ��, rsTmp!�շ���ĿID, lngִ�п���ID, _
                                    mrsPrice!���� * dbl����, IIF(gbln�Ӱ�Ӽ� And NVL(rsTmp!�Ӱ�Ӽ�, 0) = 1, NVL(rsTmp!�Ӱ�Ӽ���, 0) / 100, 0)), gstrDec)
                            Else
                                curʵ�� = curʵ�� + curӦ��
                            End If
                        ElseIf NVL(rsTmp!�Ƿ���, 0) = 1 And rsTmp!��� = "4" And NVL(rsTmp!��������, 0) = 1 Then
                            '�������õ�ʱ�����ĺ�ҩƷһ������
                            mrsPrice!���� = CalcDrugPrice(rsTmp!�շ���ĿID, NVL(mrsPrice!ִ�п���ID, 0), dbl���� * NVL(rsTmp!�շ�����, 0), , True, , mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�)
                            
                            dblӦ�� = Format(mrsPrice!���� * dbl����, "0.00000") * Format(mrsPrice!����, gstrDecPrice)
                            
                            '����Ӱ�Ӽ�
                            If gbln�Ӱ�Ӽ� And NVL(rsTmp!�Ӱ�Ӽ�, 0) = 1 Then
                                dblӦ�� = dblӦ�� * (1 + NVL(rsTmp!�Ӱ�Ӽ���, 0) / 100)
                            End If
    
                            curӦ�� = Format(dblӦ��, gstrDec)
                            
                            If Not IsNull(mrsPati!�ѱ�) And Not (gbln��������ۿ� And blnHaveSub) And NVL(rsTmp!���ηѱ�, 0) = 0 Then
                                curʵ�� = curʵ�� + Format(ActualMoney(mrsPati!�ѱ�, rsTmp!������ĿID, curӦ��, rsTmp!�շ���ĿID, lngִ�п���ID, _
                                    mrsPrice!���� * dbl����, IIF(gbln�Ӱ�Ӽ� And NVL(rsTmp!�Ӱ�Ӽ�, 0) = 1, NVL(rsTmp!�Ӱ�Ӽ���, 0) / 100, 0)), gstrDec)
                            Else
                                curʵ�� = curʵ�� + curӦ��
                            End If
                        Else '�̶��۸����ͨ���(ֻ��һ��������Ŀ)
                            mrsPrice!���� = NVL(mrsPrice!����, 0) + NVL(rsTmp!����, 0)
                            
                            dblӦ�� = Format(mrsPrice!���� * dbl����, "0.00000") * Format(NVL(rsTmp!����, 0), gstrDecPrice)
                            
                            '����Ӱ�Ӽ�
                            If gbln�Ӱ�Ӽ� And NVL(rsTmp!�Ӱ�Ӽ�, 0) = 1 Then
                                dblӦ�� = dblӦ�� * (1 + NVL(rsTmp!�Ӱ�Ӽ���, 0) / 100)
                            End If
                            
                            curӦ�� = Format(dblӦ��, gstrDec)
                            
                            If Not IsNull(mrsPati!�ѱ�) And Not (gbln��������ۿ� And blnHaveSub) And NVL(rsTmp!���ηѱ�, 0) = 0 Then
                                curʵ�� = curʵ�� + Format(ActualMoney(mrsPati!�ѱ�, rsTmp!������ĿID, curӦ��, rsTmp!�շ���ĿID, lngִ�п���ID, _
                                    mrsPrice!���� * dbl����, IIF(gbln�Ӱ�Ӽ� And NVL(rsTmp!�Ӱ�Ӽ�, 0) = 1, NVL(rsTmp!�Ӱ�Ӽ���, 0) / 100, 0)), gstrDec)
                            Else
                                curʵ�� = curʵ�� + curӦ��
                            End If
                        End If
                        
                        rsTmp.MoveNext
                    Loop
                    
                    '������Ŀ���ܼ����ۿ�
                    If gbln��������ۿ� And blnHaveSub And lng������ID <> 0 Then
                        cur��� = Format(ActualMoney(NVL(mrsPati!�ѱ�), lng������ID, cur���), gstrDec)
                    End If
                    
                    cur�ϼ� = cur�ϼ� + cur���
                Next
            End If
        End If
    End With
    LoadAdvicePrice = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetComboList(ByVal lngRow As Long) As String
'���ܣ����ݵ�ǰҽ���л�ȡ��ѡ��ļƼ�ҽ������
'������lngRow=�ɼ���(ҩ�ƻ��ҩ)
'˵����ע�������Ǹ��ݾ���ҽ����ȡ
    Dim strCombo As String
    Dim strTmp As String, lngTmp As Long
    Dim i As Long, j As Long
    
    With vsAdvice
        If .Cell(flexcpData, lngRow, COL_ID) = 3 Then
            '��ҩ�÷�����ҩ�÷�,��ҩ�巨
            lngTmp = .FindRow(CStr(Val(.TextMatrix(lngRow, COL_ID))), , COL_���ID)
            For i = lngTmp To lngRow
                If InStr(",2,3,", CLng(.Cell(flexcpData, i, COL_ID))) > 0 Then
                    If Val(.TextMatrix(i, COL_�Ƽ�����)) = 0 And InStr(",0,5,", Val(.TextMatrix(i, COL_ִ������ID))) = 0 Then
                        mrsPrice.Filter = "ҽ��ID=" & Val(.TextMatrix(i, COL_ID))
                        If Not mrsPrice.EOF Then
                            For j = 1 To mrsPrice.RecordCount
                                If NVL(mrsPrice!�̶�, 0) = 0 Then
                                    If .Cell(flexcpData, i, COL_ID) = 2 Then
                                        strTmp = Val(.TextMatrix(i, COL_ID)) & ";��ҩ�巨-" & .Cell(flexcpData, i, col_ҽ������)
                                    ElseIf .Cell(flexcpData, i, COL_ID) = 3 Then
                                        strTmp = Val(.TextMatrix(i, COL_ID)) & ";��ҩ�÷�-" & .Cell(flexcpData, i, col_ҽ������)
                                    End If
                                    If InStr(strCombo & "|", "|#" & strTmp & "|") = 0 Then
                                        strCombo = strCombo & "|#" & strTmp
                                    End If
                                End If
                                mrsPrice.MoveNext
                            Next
                        Else
                            If .Cell(flexcpData, i, COL_ID) = 2 Then
                                strTmp = Val(.TextMatrix(i, COL_ID)) & ";��ҩ�巨-" & .Cell(flexcpData, i, col_ҽ������)
                            ElseIf .Cell(flexcpData, i, COL_ID) = 3 Then
                                strTmp = Val(.TextMatrix(i, COL_ID)) & ";��ҩ�÷�-" & .Cell(flexcpData, i, col_ҽ������)
                            End If
                            If InStr(strCombo & "|", "|#" & strTmp & "|") = 0 Then
                                strCombo = strCombo & "|#" & strTmp
                            End If
                        End If
                    End If
                End If
            Next
        ElseIf .Cell(flexcpData, lngRow, COL_ID) = 4 Then
            '�ɼ�������
            lngTmp = .FindRow(CStr(Val(.TextMatrix(lngRow, COL_ID))), , COL_���ID)
            For i = lngTmp To lngRow
                If Val(.TextMatrix(i, COL_�Ƽ�����)) = 0 And InStr(",0,5,", Val(.TextMatrix(i, COL_ִ������ID))) = 0 Then
                    mrsPrice.Filter = "ҽ��ID=" & Val(.TextMatrix(i, COL_ID))
                    If Not mrsPrice.EOF Then
                        For j = 1 To mrsPrice.RecordCount
                            If NVL(mrsPrice!�̶�, 0) = 0 Then
                                If .TextMatrix(i, COL_�������) = "C" Then
                                    strTmp = Val(.TextMatrix(i, COL_ID)) & ";������Ŀ-" & .Cell(flexcpData, i, col_ҽ������)
                                ElseIf .TextMatrix(i, COL_�������) = "E" Then
                                    strTmp = Val(.TextMatrix(i, COL_ID)) & ";�ɼ�����-" & .Cell(flexcpData, i, col_ҽ������)
                                End If
                                If InStr(strCombo & "|", "|#" & strTmp & "|") = 0 Then
                                    strCombo = strCombo & "|#" & strTmp
                                End If
                            End If
                            mrsPrice.MoveNext
                        Next
                    Else
                        If .TextMatrix(i, COL_�������) = "C" Then
                            strTmp = Val(.TextMatrix(i, COL_ID)) & ";������Ŀ-" & .Cell(flexcpData, i, col_ҽ������)
                        ElseIf .TextMatrix(i, COL_�������) = "E" Then
                            strTmp = Val(.TextMatrix(i, COL_ID)) & ";�ɼ�����-" & .Cell(flexcpData, i, col_ҽ������)
                        End If
                        If InStr(strCombo & "|", "|#" & strTmp & "|") = 0 Then
                            strCombo = strCombo & "|#" & strTmp
                        End If
                    End If
                End If
            Next
        ElseIf InStr(",5,6,", .TextMatrix(lngRow, COL_�������)) > 0 Then
            '���г�ҩ����ҩ;��
            If Val(.TextMatrix(lngRow - 1, COL_���ID)) <> Val(.TextMatrix(lngRow, COL_���ID)) Then
                lngTmp = .FindRow(CStr(Val(.TextMatrix(lngRow, COL_���ID))), lngRow + 1, COL_ID)
                If Val(.TextMatrix(lngTmp, COL_�Ƽ�����)) = 0 And InStr(",0,5,", Val(.TextMatrix(lngTmp, COL_ִ������ID))) = 0 Then
                    mrsPrice.Filter = "ҽ��ID=" & Val(.TextMatrix(lngTmp, COL_ID))
                    If Not mrsPrice.EOF Then
                        For j = 1 To mrsPrice.RecordCount
                            If NVL(mrsPrice!�̶�, 0) = 0 Then
                                strCombo = "|#" & Val(.TextMatrix(lngTmp, COL_ID)) & ";��ҩ;��-" & .Cell(flexcpData, lngTmp, col_ҽ������)
                                Exit For
                            End If
                            mrsPrice.MoveNext
                        Next
                    Else
                        strCombo = "|#" & Val(.TextMatrix(lngTmp, COL_ID)) & ";��ҩ;��-" & .Cell(flexcpData, lngTmp, col_ҽ������)
                    End If
                End If
            End If
        Else
            'һ���������飬����Ѫҽ���������ҽ��
            For i = lngRow To .Rows - 1
                If i = lngRow Or Val(.TextMatrix(i, COL_���ID)) = Val(.TextMatrix(lngRow, COL_ID)) Then
                    If Val(.TextMatrix(i, COL_�Ƽ�����)) = 0 And InStr(",0,5,", Val(.TextMatrix(i, COL_ִ������ID))) = 0 Then
                        mrsPrice.Filter = "ҽ��ID=" & Val(.TextMatrix(i, COL_ID))
                        If Not mrsPrice.EOF Then
                            For j = 1 To mrsPrice.RecordCount
                                If NVL(mrsPrice!�̶�, 0) = 0 Then
                                    If .TextMatrix(i, COL_�������) = "F" And Val(.TextMatrix(i, COL_���ID)) <> 0 Then
                                        strTmp = Val(.TextMatrix(i, COL_ID)) & ";��������-" & .Cell(flexcpData, i, col_ҽ������)
                                    ElseIf .TextMatrix(i, COL_�������) = "G" And Val(.TextMatrix(i, COL_���ID)) <> 0 Then
                                        strTmp = Val(.TextMatrix(i, COL_ID)) & ";��������-" & .Cell(flexcpData, i, col_ҽ������)
                                    ElseIf .TextMatrix(i, COL_�������) = "D" And Val(.TextMatrix(i, COL_���ID)) <> 0 Then
                                        strTmp = Val(.TextMatrix(i, COL_ID)) & ";��鲿λ-" & .TextMatrix(i, COL_�걾��λ) & "(" & .TextMatrix(i, COL_��鷽��) & ")"
                                    ElseIf .TextMatrix(i, COL_�������) = "E" And .TextMatrix(lngRow, COL_�������) = "K" Then
                                        strTmp = Val(.TextMatrix(i, COL_ID)) & ";��Ѫ;��-" & .Cell(flexcpData, i, col_ҽ������)
                                    Else
                                        If mrsPrice!�������� <> 0 Then
                                            '���շ��ã�Ŀǰ�������Ĵ��Ժ����м���
                                            lngTmp = -1 * Val(mrsPrice!�������� & Val(.TextMatrix(i, COL_ID)))
                                            strTmp = lngTmp & ";" & .Cell(flexcpData, i, COL_�������) & "ҽ��-" & .Cell(flexcpData, i, col_ҽ������) & _
                                                "(" & decode(Val(.TextMatrix(i, COL_ִ�б��)), 1, "����", 2, "����", "") & "����)"
                                        Else
                                            strTmp = Val(.TextMatrix(i, COL_ID)) & ";" & .Cell(flexcpData, i, COL_�������) & "ҽ��-" & .Cell(flexcpData, i, col_ҽ������)
                                        End If
                                    End If
                                    If InStr(strCombo & "|", "|#" & strTmp & "|") = 0 Then
                                        strCombo = strCombo & "|#" & strTmp
                                    End If
                                End If
                                mrsPrice.MoveNext
                            Next
                        Else
                            'δ���üƼ۵ģ�����ѡ����ӼƼ���Ŀ
                            If .TextMatrix(i, COL_�������) = "F" And Val(.TextMatrix(i, COL_���ID)) <> 0 Then
                                strTmp = Val(.TextMatrix(i, COL_ID)) & ";��������-" & .Cell(flexcpData, i, col_ҽ������)
                            ElseIf .TextMatrix(i, COL_�������) = "G" And Val(.TextMatrix(i, COL_���ID)) <> 0 Then
                                strTmp = Val(.TextMatrix(i, COL_ID)) & ";��������-" & .Cell(flexcpData, i, col_ҽ������)
                            ElseIf .TextMatrix(i, COL_�������) = "D" And Val(.TextMatrix(i, COL_���ID)) <> 0 Then
                                strTmp = Val(.TextMatrix(i, COL_ID)) & ";��鲿λ-" & .TextMatrix(i, COL_�걾��λ) & "(" & .TextMatrix(i, COL_��鷽��) & ")"
                            ElseIf .TextMatrix(i, COL_�������) = "E" And .TextMatrix(lngRow, COL_�������) = "K" Then
                                strTmp = Val(.TextMatrix(i, COL_ID)) & ";��Ѫ;��-" & .Cell(flexcpData, i, col_ҽ������)
                            Else
                                strTmp = Val(.TextMatrix(i, COL_ID)) & ";" & .Cell(flexcpData, i, COL_�������) & "ҽ��-" & .Cell(flexcpData, i, col_ҽ������)
                            End If
                            If InStr(strCombo & "|", "|#" & strTmp & "|") = 0 Then
                                strCombo = strCombo & "|#" & strTmp
                            End If
                            
                            '���շ��ã�Ŀǰ�������Ĵ��Ի����м���
                            If .TextMatrix(i, COL_�������) = "D" And Val(.TextMatrix(i, COL_���ID)) = 0 _
                                And (Val(.TextMatrix(i, COL_ִ�б��)) = 1 Or Val(.TextMatrix(i, COL_ִ�б��)) = 2) Then
                                lngTmp = -1 * Val(1 & Val(.TextMatrix(i, COL_ID)))
                                strTmp = lngTmp & ";" & .Cell(flexcpData, i, COL_�������) & "ҽ��-" & .Cell(flexcpData, i, col_ҽ������) & _
                                    "(" & decode(Val(.TextMatrix(i, COL_ִ�б��)), 1, "����", 2, "����", "") & "����)"
                                If InStr(strCombo & "|", "|#" & strTmp & "|") = 0 Then
                                    strCombo = strCombo & "|#" & strTmp
                                End If
                            End If
                        End If
                    End If
                Else
                    Exit For
                End If
            Next
        End If
    End With
    
    GetComboList = Mid(strCombo, 2)
End Function

Private Function ShowAdvicePrice(ByVal lngRow As Long) As Boolean
'���ܣ�����ҽ���Ƽ۹�ϵ�����㲢��ʾָ��ҽ���ķ���(����ҽ�������ܶ���)
'������lngRow=�ɼ���(ҩ�ƻ��ҩ)
    Dim rsTmp As New ADODB.Recordset
    Dim rsExeDays As ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim lngTopRow As Long, lngLeftCol As Long
    Dim lngPreRow As Long, lngPreCol As Long
    Dim blnFirst As Boolean, str�Ƽ�ҽ�� As String
    Dim str��λ As String, dbl���� As Double
    Dim bln�������� As Boolean, strCombo As String, str�к� As String, str�ֽ�ʱ�� As String
    Dim dbl���� As Double, curӦ�� As Currency, curʵ�� As Currency
    Dim dbl��ǰ���� As Double, dbl��ǰӦ�� As Double, cur��ǰӦ�� As Currency, cur��ǰʵ�� As Currency
    Dim lng�к� As Long, cur�ϼ� As Currency
    
    Dim rsMain As New ADODB.Recordset
    Dim rsClone As New ADODB.Recordset
    Dim strHaveSub As String, strNoneSub As String
    Dim strPriceType As String
        
    On Error GoTo errH
    
    '���ڻ��ܼ����ۿ۵���ʱ��¼��
    rsMain.Fields.Append "ҽ���к�", adBigInt
    rsMain.Fields.Append "��������", adInteger
    rsMain.Fields.Append "�����к�", adBigInt
    rsMain.Fields.Append "������ID", adBigInt
    rsMain.Fields.Append "ҽ���ϼ�", adCurrency, , adFldIsNullable
    rsMain.CursorLocation = adUseClient
    rsMain.LockType = adLockOptimistic
    rsMain.CursorType = adOpenStatic
    rsMain.Open
    
    With vsAdvice
        blnFirst = True
        If InStr(",5,6,", .TextMatrix(lngRow, COL_�������)) > 0 Then
            If Val(.TextMatrix(lngRow - 1, COL_���ID)) = Val(.TextMatrix(lngRow, COL_���ID)) Then
                blnFirst = False 'һ����ҩ���Ƿ��һҩƷ��
            End If
        End If
        
        If Val(.TextMatrix(lngRow, COL_���ID)) <> 0 Then
            If blnFirst Then
                mrsPrice.Filter = "ҽ��ID=" & Val(.TextMatrix(lngRow, COL_ID)) & _
                    " Or ҽ��ID=" & Val(.TextMatrix(lngRow, COL_���ID))
            Else
                mrsPrice.Filter = "ҽ��ID=" & Val(.TextMatrix(lngRow, COL_ID))
            End If
        Else
            mrsPrice.Filter = "ҽ��ID=" & Val(.TextMatrix(lngRow, COL_ID)) & _
                " Or ���ID=" & Val(.TextMatrix(lngRow, COL_ID))
        End If
        
        For i = 1 To mrsPrice.RecordCount
            '�Ƽ�ҽ��
            bln�������� = False
            lng�к� = .FindRow(CStr(mrsPrice!ҽ��ID), , COL_ID)
            If .TextMatrix(lng�к�, COL_�������) = "4" Then
                str�Ƽ�ҽ�� = "��������-" & .Cell(flexcpData, lng�к�, col_ҽ������)
            ElseIf InStr(",5,6,7", .TextMatrix(lng�к�, COL_�������)) > 0 Then
                str�Ƽ�ҽ�� = "ҩƷҽ��-" & .Cell(flexcpData, lng�к�, col_ҽ������)
            ElseIf CLng(.Cell(flexcpData, lng�к�, COL_ID)) = 1 Then
                str�Ƽ�ҽ�� = "��ҩ;��-" & .Cell(flexcpData, lng�к�, col_ҽ������)
            ElseIf CLng(.Cell(flexcpData, lng�к�, COL_ID)) = 2 Then
                str�Ƽ�ҽ�� = "��ҩ�巨-" & .Cell(flexcpData, lng�к�, col_ҽ������)
            ElseIf CLng(.Cell(flexcpData, lng�к�, COL_ID)) = 3 Then
                str�Ƽ�ҽ�� = "��ҩ�÷�-" & .Cell(flexcpData, lng�к�, col_ҽ������)
            ElseIf CLng(.Cell(flexcpData, lng�к�, COL_ID)) = 4 Then
                str�Ƽ�ҽ�� = "�ɼ�����-" & .Cell(flexcpData, lng�к�, col_ҽ������)
            ElseIf CLng(.Cell(flexcpData, lng�к�, COL_ID)) = 5 Then
                str�Ƽ�ҽ�� = "��Ѫ;��-" & .Cell(flexcpData, lng�к�, col_ҽ������)
            ElseIf .TextMatrix(lng�к�, COL_�������) = "C" And Val(.TextMatrix(lng�к�, COL_���ID)) <> 0 Then
                str�Ƽ�ҽ�� = "������Ŀ-" & .Cell(flexcpData, lng�к�, col_ҽ������)
            ElseIf .TextMatrix(lng�к�, COL_�������) = "F" And Val(.TextMatrix(lng�к�, COL_���ID)) <> 0 Then
                bln�������� = True
                str�Ƽ�ҽ�� = "��������-" & .Cell(flexcpData, lng�к�, col_ҽ������)
            ElseIf .TextMatrix(lng�к�, COL_�������) = "G" And Val(.TextMatrix(lng�к�, COL_���ID)) <> 0 Then
                str�Ƽ�ҽ�� = "��������-" & .Cell(flexcpData, lng�к�, col_ҽ������)
            ElseIf .TextMatrix(lng�к�, COL_�������) = "D" And Val(.TextMatrix(lng�к�, COL_���ID)) <> 0 Then
                str�Ƽ�ҽ�� = "��鲿λ-" & .TextMatrix(lng�к�, COL_�걾��λ) & "(" & .TextMatrix(lng�к�, COL_��鷽��) & ")"
            Else
                If NVL(mrsPrice!��������, 0) = 1 Then
                    '���Ի����м��շ���
                    str�Ƽ�ҽ�� = .Cell(flexcpData, lng�к�, COL_�������) & "ҽ��-" & .Cell(flexcpData, lng�к�, col_ҽ������) & _
                        "(" & decode(Val(.TextMatrix(lng�к�, COL_ִ�б��)), 1, "����", 2, "����", "") & "����)"
                Else
                    str�Ƽ�ҽ�� = .Cell(flexcpData, lng�к�, COL_�������) & "ҽ��-" & .Cell(flexcpData, lng�к�, col_ҽ������)
                End If
            End If
            str�Ƽ�ҽ�� = Replace(str�Ƽ�ҽ��, "'", "''")
            
            '����:ҩƷ��סԺ��λ������,��������������
            If InStr(",5,6,", .TextMatrix(lng�к�, COL_�������)) > 0 Then
                dbl���� = Val(.TextMatrix(lng�к�, COL_����))
            ElseIf .TextMatrix(lng�к�, COL_�������) = "7" Then
                '��ҩҩ����λ�����ɷ��㴦��:ÿ��
                If Val(.TextMatrix(lng�к�, COL_�ɷ����)) = 0 Then
                    dbl���� = Val(.TextMatrix(lng�к�, COL_����)) * Val(.TextMatrix(lng�к�, COL_����)) _
                        / Val(.TextMatrix(lng�к�, COL_����ϵ��)) / Val(.TextMatrix(lng�к�, COL_סԺ��װ))
                Else
                    dbl���� = Val(.TextMatrix(lng�к�, COL_����)) _
                        * IntEx(Val(.TextMatrix(lng�к�, COL_����)) / Val(.TextMatrix(lng�к�, COL_����ϵ��)) / Val(.TextMatrix(lng�к�, COL_סԺ��װ)))
                End If
            Else
                If InStr(",3,4,5,6,", Val("" & mrsPrice!�շѷ�ʽ)) > 0 Then 'һ��ֻ��һ�ε�
                     '�ֽ�ʱ��
                    If .TextMatrix(lng�к�, COL_�ֽ�ʱ��) <> "" Then
                        str�ֽ�ʱ�� = .TextMatrix(lng�к�, COL_�ֽ�ʱ��)
                    Else
                        str�ֽ�ʱ�� = .Cell(flexcpData, lng�к�, COL_�ֽ�ʱ��)    '��ʼִ��ʱ��
                    End If
                    
                    Set rsExeDays = GetExecDays(str�ֽ�ʱ��)
                    dbl���� = rsExeDays.RecordCount
                ElseIf InStr(",1,2,", Val("" & mrsPrice!�շѷ�ʽ)) > 0 Then 'һ�η���ֻ��һ��
                    dbl���� = 1
                Else
                    dbl���� = Val(.TextMatrix(lng�к�, COL_����))
                End If
            End If
            dbl���� = Format(dbl���� * NVL(mrsPrice!����, 0), "0.00000")
                        
            '���SQL
            strSQL = strSQL & IIF(strSQL = "", "", " Union ALL ") & _
                " Select " & i & " as ���," & mrsPrice!ҽ��ID & " as ҽ��ID,ID," & _
                NVL(mrsPrice!�̶�, 0) & " as �̶�,'" & str�Ƽ�ҽ�� & "' as �Ƽ�ҽ��,���,����,����,���," & _
                "���㵥λ as ��λ," & NVL(mrsPrice!����, 0) & " as �Ƽ�����," & dbl���� & " as ����," & _
                Format(NVL(mrsPrice!����, 0), gstrDecPrice) & " as ����,��������," & lng�к� & " as �к�," & _
                " �Ƿ���,�Ӱ�Ӽ�," & IIF(bln��������, 1, 0) & " as ��������," & mrsPrice!���� & " as ����," & _
                NVL(mrsPrice!ִ�п���ID, 0) & " as ִ�п���ID,���ηѱ�," & mrsPrice!�������� & " as ��������," & _
                mrsPrice!�շѷ�ʽ & " as �շѷ�ʽ From �շ���ĿĿ¼ Where ID=" & mrsPrice!�շ�ϸĿID
            mrsPrice.MoveNext
        Next
    End With
    
    With vsPrice
        lngPreRow = .Row: lngPreCol = .Col
        lngTopRow = .TopRow: lngLeftCol = .LeftCol
        .Editable = flexEDNone
        .Redraw = flexRDNone
        .Rows = .FixedRows
        
        '��Ҫ�Ƽ۵�ҽ��ѡ��
        '���ݴ�����ҽ��ȡ�ɼƼ�ҽ��(���ܴ�mrsPriceȡ,��Ϊ�������շѹ�ϵ����ɾ��,����Ҳ�����ڼƼ���ȫ��ɾ��)
        strCombo = GetComboList(lngRow)
        If strCombo <> "" Then
            .ColData(COLP_�Ƽ�ҽ��) = strCombo
            .Editable = flexEDKbdMouse '����ѡ������Ա༭
        Else
            .ColData(COLP_�Ƽ�ҽ��) = ""
        End If
        
        '��ʾ���еļƼ���Ŀ
        If strSQL <> "" Then
            strSQL = "Select A.�к�,A.ID AS �շ�ϸĿID,A.�̶�,A.����,A.�Ƽ�ҽ��,A.���,C.���� as �������,A.ִ�п���ID,G.���� as ִ�п���," & _
                " Nvl(E.����,A.����)||Decode(A.����,NULL,NULL,'('||A.����||')')||Decode(A.���,NULL,NULL,' '||A.���) as ����," & _
                " A.��λ,A.�Ƽ�����,A.����,D.סԺ��װ,D.סԺ��λ,Decode(A.�Ƿ���,1,A.����,B.�ּ�) as ����,F.��������," & _
                " A.��������,A.�շѷ�ʽ,A.��������,A.���ηѱ�,A.�Ƿ���,A.�Ӱ�Ӽ�,B.�Ӱ�Ӽ���,B.ԭ��,B.�ּ�,A.��������,B.�����շ���,B.������ĿID" & _
                " From (" & strSQL & ") A,�շѼ�Ŀ B,�շ���Ŀ��� C,ҩƷ��� D,�շ���Ŀ���� E,�������� F,���ű� G" & _
                " Where A.ID=B.�շ�ϸĿID And A.���=C.���� And A.ID=D.ҩƷID(+)" & _
                " And A.ID=F.����ID(+) And A.ִ�п���ID=G.ID(+)" & _
                GetPriceGradeSQL(mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�, "A", "B", "1", "2", "3") & _
                " And ((Sysdate Between B.ִ������ and B.��ֹ����) or (Sysdate>=B.ִ������ And B.��ֹ���� is NULL))" & _
                " And A.ID=E.�շ�ϸĿID(+) And E.����(+)=1 And E.����(+)=" & IIF(gbytҩƷ������ʾ = 0, 1, 3) & _
                " Order by A.���"
                '��Ϊ������ǵ��ñ�����ˢ��,Ҫ���ֶ�̬��¼���м�¼˳��
                'Ҫ��֤��������ǰ��,LoadAdvicePriceʱ������������ǰ�棬���ұ༭��ֻ���ܼ��˴���
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�) 'û��
            
            If Not rsTmp.EOF And gbln��������ۿ� Then
                Set rsClone = rsTmp.Clone
            End If
            
            For i = 1 To rsTmp.RecordCount
                If str�к� <> rsTmp!�к� & "_" & rsTmp!�������� & "_" & rsTmp!�շ�ϸĿID Then
                    If str�к� <> "" Then
                        If Not (Val(.TextMatrix(.Rows - 1, COLP_���)) = 1 And dbl���� = 0) Then
                            .TextMatrix(.Rows - 1, COLP_����) = Format(dbl����, gstrDecPrice)
                            .Cell(flexcpData, .Rows - 1, COLP_����) = .TextMatrix(.Rows - 1, COLP_����) '��¼���ڻָ�����
                            .TextMatrix(.Rows - 1, COLP_Ӧ�ս��) = Format(curӦ��, gstrDec)
                            .TextMatrix(.Rows - 1, COLP_ʵ�ս��) = Format(curʵ��, gstrDec)
                        End If
                        cur�ϼ� = cur�ϼ� + Format(curʵ��, gstrDec)
                    End If
                    str�к� = rsTmp!�к� & "_" & rsTmp!�������� & "_" & rsTmp!�շ�ϸĿID
                    dbl���� = 0: curӦ�� = 0: curʵ�� = 0
                    .Rows = .Rows + 1
                    
                    '��ʶ�̶�����Ϊ��ɫ
                    If rsTmp!�̶� <> 0 Then
                        .Cell(flexcpBackColor, .Rows - 1, .FixedCols, .Rows - 1, .Cols - 1) = &HE0E0E0
                    End If

                    .TextMatrix(.Rows - 1, COLP_�к�) = rsTmp!�к�
                    .TextMatrix(.Rows - 1, COLP_�շ�ϸĿID) = rsTmp!�շ�ϸĿID
                    .TextMatrix(.Rows - 1, COLP_�̶�) = rsTmp!�̶�
                    .TextMatrix(.Rows - 1, COLP_�Ƽ�ҽ��) = rsTmp!�Ƽ�ҽ��
                    .TextMatrix(.Rows - 1, COLP_��������) = rsTmp!��������
                    .TextMatrix(.Rows - 1, COLP_�շѷ�ʽ) = getChargeMode(Val(NVL(rsTmp!�շѷ�ʽ, 0)))
                        .Cell(flexcpData, .Rows - 1, COLP_�շѷ�ʽ) = Val(NVL(rsTmp!�շѷ�ʽ, 0))
                    .TextMatrix(.Rows - 1, COLP_���) = rsTmp!�������
                    .TextMatrix(.Rows - 1, COLP_�շ����) = rsTmp!���
                    .TextMatrix(.Rows - 1, COLP_�շ���Ŀ) = rsTmp!����
                    .TextMatrix(.Rows - 1, COLP_�Ƽ�����) = NVL(rsTmp!�Ƽ�����, 0) '�������
                    
                    dbl���� = NVL(rsTmp!����, 0) '�ۼ��������ں��水�ɱ����ۼ���
                    If InStr(",5,6,7,", rsTmp!���) > 0 Then 'סԺ��װ
                        .TextMatrix(.Rows - 1, COLP_��λ) = NVL(rsTmp!סԺ��λ)
                        If InStr(",5,6,7,", vsAdvice.TextMatrix(rsTmp!�к�, COL_�������)) > 0 Then
                            .TextMatrix(.Rows - 1, COLP_����) = FormatEx(NVL(rsTmp!����, 0), 5)
                            dbl���� = dbl���� * NVL(rsTmp!סԺ��װ, 1)
                        Else
                            '��ҩҩ����λ�����ɷ��㴦��:ÿ��
                            '��ҩ��ҩƷ�Ƽ�:��Ϊ����Ԥ�����ۼ�����,���ת��Ϊҩ����λ��ʾʱ���������㴦��
                            .TextMatrix(.Rows - 1, COLP_����) = FormatEx(NVL(rsTmp!����, 0) / NVL(rsTmp!סԺ��װ, 1), 5)
                        End If
                    Else
                        .TextMatrix(.Rows - 1, COLP_��λ) = NVL(rsTmp!��λ)
                        .TextMatrix(.Rows - 1, COLP_����) = FormatEx(NVL(rsTmp!����, 0), 5)
                    End If
                    
                    .TextMatrix(.Rows - 1, COLP_ִ�п���) = NVL(rsTmp!ִ�п���)
                    .TextMatrix(.Rows - 1, COLP_ִ�п���ID) = NVL(rsTmp!ִ�п���ID, 0)
                    
                    '��ʾҽ���������ͣ�ҽ��վ����ֻ�ܷ�������
                    If Val(rsTmp!�շ�ϸĿID) <> 0 Then
                        strPriceType = GetPriceType(Val(mlng����ID), Val(rsTmp!�շ�ϸĿID & ""), Val(mint����), mlng�������� = 1)
                    End If
                    '��������
                    If strPriceType = "" Then
                        .TextMatrix(.Rows - 1, COLP_��������) = NVL(rsTmp!��������)
                    Else
                        .TextMatrix(.Rows - 1, COLP_��������) = strPriceType
                    End If
                    
                    .TextMatrix(.Rows - 1, COLP_����) = IIF(NVL(rsTmp!����, 0) = 0, "", "��")
                    .TextMatrix(.Rows - 1, COLP_��������) = NVL(rsTmp!��������, 0)
                    
                    '��¼��������ָ�
                    .Cell(flexcpData, .Rows - 1, COLP_�Ƽ�ҽ��) = .TextMatrix(.Rows - 1, COLP_�Ƽ�ҽ��)
                    .Cell(flexcpData, .Rows - 1, COLP_�շ���Ŀ) = .TextMatrix(.Rows - 1, COLP_�շ���Ŀ)
                    .Cell(flexcpData, .Rows - 1, COLP_�Ƽ�����) = .TextMatrix(.Rows - 1, COLP_�Ƽ�����)
                    .Cell(flexcpData, .Rows - 1, COLP_ִ�п���) = .TextMatrix(.Rows - 1, COLP_ִ�п���)
                    
                    '��¼�����������Ϣ���Ա����
                    If gbln��������ۿ� And rsTmp!���� = 0 Then
                        If InStr(strHaveSub & ",", "," & rsTmp!�к� & "_" & rsTmp!�������� & ",") = 0 _
                            And InStr(strNoneSub & ",", "," & rsTmp!�к� & "_" & rsTmp!�������� & ",") = 0 Then
                            rsClone.Filter = "�к�=" & rsTmp!�к� & " And ��������=" & rsTmp!�������� & " And ����=1"
                            If Not rsClone.EOF Then
                                rsMain.AddNew
                                rsMain!ҽ���к� = rsTmp!�к�
                                rsMain!�������� = rsTmp!��������
                                rsMain!�����к� = .Rows - 1
                                rsMain!������ID = rsTmp!������ĿID
                                rsMain.Update
                                strHaveSub = strHaveSub & "," & rsTmp!�к� & "_" & rsTmp!��������
                            Else
                                strNoneSub = strNoneSub & "," & rsTmp!�к� & "_" & rsTmp!��������
                            End If
                        End If
                    End If
                    
                    '��ҩƷ������ҽ����ҩƷ�͸������ļƼۣ���ʹ�̶�Ҳ�����޸�ִ�п���
                    If InStr(",5,6,7,", rsTmp!���) > 0 _
                        Or rsTmp!��� = "4" And NVL(rsTmp!��������, 0) = 1 Then
                        .Editable = flexEDKbdMouse
                    End If
                End If
                
                '���ۼ��㴦��
                If InStr(",5,6,7,", rsTmp!���) > 0 Then
                    If NVL(rsTmp!�Ƿ���, 0) = 0 Then
                        dbl��ǰ���� = NVL(rsTmp!����, 0)
                    Else
                        If InStr(",5,6,7,", vsAdvice.TextMatrix(rsTmp!�к�, COL_�������)) > 0 Then
                            dbl��ǰ���� = CalcDrugPrice(rsTmp!�շ�ϸĿID, NVL(rsTmp!ִ�п���ID, 0), Format(NVL(rsTmp!����, 0) * NVL(rsTmp!סԺ��װ, 1), "0.00000"), , True, , mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�)
                        Else
                            dbl��ǰ���� = CalcDrugPrice(rsTmp!�շ�ϸĿID, NVL(rsTmp!ִ�п���ID, 0), Format(NVL(rsTmp!����, 0), "0.00000"), , True, , mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�)
                        End If
                    End If
                    If InStr(",5,6,7,", vsAdvice.TextMatrix(rsTmp!�к�, COL_�������)) > 0 Then
                        dbl��ǰ���� = dbl��ǰ���� * NVL(rsTmp!סԺ��װ, 1)
                        dbl��ǰӦ�� = Format(NVL(rsTmp!����, 0), "0.00000") * dbl��ǰ����
                    Else
                        dbl��ǰӦ�� = Format(NVL(rsTmp!����, 0), "0.00000") * dbl��ǰ����
                        dbl��ǰ���� = dbl��ǰ���� * NVL(rsTmp!סԺ��װ, 1)
                    End If
                ElseIf rsTmp!��� = "4" And NVL(rsTmp!��������, 0) = 1 And NVL(rsTmp!�Ƿ���, 0) = 1 Then
                    '�������õ�ʱ�����ĺ�ҩƷһ������
                    dbl��ǰ���� = CalcDrugPrice(rsTmp!�շ�ϸĿID, NVL(rsTmp!ִ�п���ID, 0), Format(NVL(rsTmp!����, 0), "0.00000"), , True, , mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�)
                    dbl��ǰӦ�� = Format(NVL(rsTmp!����, 0), "0.00000") * dbl��ǰ����
                Else
                    dbl��ǰ���� = NVL(rsTmp!����, 0) '�������Ϊ��������û������
                    dbl��ǰӦ�� = Format(NVL(rsTmp!����, 0), "0.00000") * dbl��ǰ����
                    If NVL(rsTmp!�Ƿ���, 0) = 1 Then '��¼��ҩ��۷�Χ
                        .TextMatrix(.Rows - 1, COLP_���) = 1
                        .Cell(flexcpData, .Rows - 1, COLP_Ӧ�ս��) = CCur(NVL(rsTmp!ԭ��, 0))
                        .Cell(flexcpData, .Rows - 1, COLP_ʵ�ս��) = CCur(NVL(rsTmp!�ּ�, 0))
                        .Editable = flexEDKbdMouse '��ҩƷ���,��ʹ�̶�Ҳ���Զ���
                    End If
                End If
                'Ӧ��
                If rsTmp!�������� = 1 Then
                    dbl��ǰӦ�� = dbl��ǰӦ�� * NVL(rsTmp!�����շ���, 100) / 100
                End If
                '����Ӱ�Ӽ�
                If gbln�Ӱ�Ӽ� And NVL(rsTmp!�Ӱ�Ӽ�, 0) = 1 Then
                    dbl��ǰӦ�� = dbl��ǰӦ�� * (1 + NVL(rsTmp!�Ӱ�Ӽ���, 0) / 100)
                End If
                cur��ǰӦ�� = Format(dbl��ǰӦ��, gstrDec)
                
                'ʵ��
                If gbln��������ۿ� And (rsTmp!���� = 1 Or InStr(strHaveSub & ",", "," & rsTmp!�к� & "_" & rsTmp!�������� & ",") > 0) Then
                    cur��ǰʵ�� = Format(cur��ǰӦ��, gstrDec)
                    '�ۼ�ҽ���ϼ��������ۿ�
                    rsMain.Filter = "ҽ���к�=" & rsTmp!�к� & " And ��������=" & rsTmp!��������
                    rsMain!ҽ���ϼ� = NVL(rsMain!ҽ���ϼ�, 0) + cur��ǰʵ��
                    rsMain.Update
                ElseIf NVL(rsTmp!���ηѱ�, 0) = 0 And Not IsNull(mrsPati!�ѱ�) Then
                    cur��ǰʵ�� = Format(ActualMoney(mrsPati!�ѱ�, rsTmp!������ĿID, cur��ǰӦ��, rsTmp!�շ�ϸĿID, NVL(rsTmp!ִ�п���ID, 0), _
                        dbl����, IIF(gbln�Ӱ�Ӽ� And NVL(rsTmp!�Ӱ�Ӽ�, 0) = 1, NVL(rsTmp!�Ӱ�Ӽ���, 0) / 100, 0)), gstrDec)
                Else
                    cur��ǰʵ�� = Format(cur��ǰӦ��, gstrDec)
                End If
                
                dbl���� = dbl���� + dbl��ǰ����
                curӦ�� = curӦ�� + cur��ǰӦ��
                curʵ�� = curʵ�� + cur��ǰʵ��
                
                rsTmp.MoveNext
            Next
            If str�к� <> "" Then
                If Not (Val(.TextMatrix(.Rows - 1, COLP_���)) = 1 And dbl���� = 0) Then
                    .TextMatrix(.Rows - 1, COLP_����) = Format(dbl����, gstrDecPrice)
                    .Cell(flexcpData, .Rows - 1, COLP_����) = .TextMatrix(.Rows - 1, COLP_����) '��¼���ڻָ�����
                    .TextMatrix(.Rows - 1, COLP_Ӧ�ս��) = Format(curӦ��, gstrDec)
                    .TextMatrix(.Rows - 1, COLP_ʵ�ս��) = Format(curʵ��, gstrDec)
                End If
                cur�ϼ� = cur�ϼ� + Format(curʵ��, gstrDec)
            End If
        End If
        
        '���ܼ����ۿ�
        If gbln��������ۿ� And strHaveSub <> "" Then
            rsMain.Filter = 0
            Do While Not rsMain.EOF
                cur��ǰʵ�� = Format(ActualMoney(NVL(mrsPati!�ѱ�), rsMain!������ID, rsMain!ҽ���ϼ�), gstrDec)
                cur�ϼ� = cur�ϼ� - Val(.TextMatrix(rsMain!�����к�, COLP_ʵ�ս��))
                .TextMatrix(rsMain!�����к�, COLP_ʵ�ս��) = Format(Val(.TextMatrix(rsMain!�����к�, COLP_ʵ�ս��)) + (cur��ǰʵ�� - rsMain!ҽ���ϼ�), gstrDec)
                cur�ϼ� = cur�ϼ� + Val(.TextMatrix(rsMain!�����к�, COLP_ʵ�ս��))
                rsMain.MoveNext
            Loop
        End If
        
        '------------------------------------------------
        If .Rows = .FixedRows Then .Rows = .FixedRows + 1
        '��λȱʡ��Ԫ
        If lngPreRow >= .FixedRows And lngPreRow <= .Rows - 1 Then
            .Row = lngPreRow
        Else
            .Row = .FixedRows
        End If
        If lngPreCol >= COLP_�Ƽ�ҽ�� And lngPreCol <= .Cols - 1 Then
            .Col = lngPreCol
        Else
            .Col = COLP_�Ƽ�ҽ��
        End If
        '��λ�������λ��
        If lngTopRow >= .FixedRows And lngTopRow <= .Rows - 1 Then
            .TopRow = lngTopRow
        End If
        If lngLeftCol >= COLP_�Ƽ�ҽ�� And lngLeftCol <= .Cols - 1 Then
            .LeftCol = lngLeftCol
        End If
        .Redraw = flexRDDirect
    End With
    
    '���»�����ʾ�ɼ��еķ���ҽ�����
    vsAdvice.TextMatrix(lngRow, COL_���) = Format(cur�ϼ�, gstrDec)
    ShowAdvicePrice = True
    
    Call ShowSendTotal
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CellEditable(ByVal lngRow As Long, ByVal lngCol As Long, Optional bln�Ǳ��� As Boolean) As Boolean
'���ܣ��жϼ۱��е�Ԫ���Ƿ���Ա༭
    Dim lng�к� As Long
    
    With vsPrice
        bln�Ǳ��� = False
        CellEditable = .Editable
        lng�к� = Val(.TextMatrix(lngRow, COLP_�к�))
        If lngCol = COLP_ִ�п��� Then
            '�������õ�����,��ҩ��ҩƷ�Ƽ۵�ִ�п��ҿ����޸�
            If Not ((.TextMatrix(lngRow, COLP_�շ����) = "4" And Val(.TextMatrix(lngRow, COLP_��������)) = 1 _
                Or InStr(",5,6,7,", .TextMatrix(lngRow, COLP_�շ����)) > 0) And InStr(",4,5,6,7,", vsAdvice.TextMatrix(lng�к�, COL_�������)) = 0) Then
                CellEditable = False
            End If
            If .TextMatrix(lngRow, COLP_�շ���Ŀ) = "" Or .TextMatrix(lngRow, COLP_�к�) = "" Then
                CellEditable = False
            End If
        ElseIf Val(.TextMatrix(lngRow, COLP_�̶�)) <> 0 Then
            '�̶������н������޸ı��
            If Not (Val(.TextMatrix(lngRow, COLP_���)) = 1 And lngCol = COLP_����) Then
                CellEditable = False
            End If
        Else
            If lngCol = COLP_���� Then
                If Val(.TextMatrix(lngRow, COLP_���)) <> 1 Then
                    CellEditable = False
                Else
                    '�Ǳ���ִ�еı����Ŀ�������۸�
                    If lng�к� <> 0 Then
                        If Not Check����ִ��(Val(vsAdvice.TextMatrix(lng�к�, COL_ִ�п���ID))) Then
                            bln�Ǳ��� = True: CellEditable = False
                        End If
                    End If
                End If
            ElseIf lngCol <> COLP_�Ƽ�ҽ�� And lngCol <> COLP_�Ƽ����� And lngCol <> COLP_�շ���Ŀ Then
                CellEditable = False
            End If
        End If
    End With
End Function

Private Sub Refresh��ҩ��()
    Dim objCbo As CommandBarComboBox
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim strPre As String
    
    On Error GoTo errH
    
    Set objCbo = cbsMain.FindControl(, conMenu_View_Find)
    
    If objCbo.ListIndex > 0 Then strPre = objCbo.List(objCbo.ListIndex)
    
    objCbo.Clear
    objCbo.AddItem "<ʹ���µ���ҩ��>"
    objCbo.ListIndex = 1
    
    strSQL = "Select Distinct ��ҩ�� From δ��ҩƷ��¼ Where ��������>=Trunc(Sysdate) And ����=9 And �Է�����ID=[1] And ��ҩ�� is Not NULL Order by ��ҩ�� Desc"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng��ҩ����ID)
    Do While Not rsTmp.EOF
        objCbo.AddItem rsTmp!��ҩ��
        If rsTmp!��ҩ�� = strPre Then
            objCbo.ListIndex = objCbo.ListCount
        End If
        rsTmp.MoveNext
    Loop

    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function Get��ҩ��() As String
    Dim objCbo As CommandBarComboBox
    
    Set objCbo = cbsMain.FindControl(, conMenu_View_Find)
    If objCbo.ListIndex = 1 Then
        Get��ҩ�� = zlDatabase.GetNextNo(122, mlng��ҩ����ID)
    ElseIf objCbo.ListIndex > 1 Then
        Get��ҩ�� = objCbo.List(objCbo.ListIndex)
    End If
End Function

Private Function LoadAdviceSend(ByVal str���s As String, ByVal strҩƷ As String, Optional ByVal lngModle As Long) As Boolean
'���ܣ�����������ȡ����ʾҪ���͵�ҩƷҽ���嵥
'������str���s,strҩƷ����Ӧ����ѡ���е�ҽ���������ҩƷ�������,lngModle=0 ��ȡ����ҽ����=1��ȡ����ҽ��
'˵����ע��CellData�д�ŵ��и�������
'   RowData��0-δ���͵�,-1-�ѳɹ����͵�
'   COL_ѡ��0-������ѡ���,1-��ֹ�ı�ѡ��״̬��
'   COL_ID��1-��ҩ;����2-��ҩ�巨��3-��ҩ�÷���4-�ɼ�������5-��Ѫ;��
'   COL_Ӥ�������Ӥ�����
'   COL_������𣺴������������ƣ�������ʾ�Ƽ�ҽ��
'   COL_ҽ�����ݣ����������Ŀ���ƻ�걾��λ��������ʾ�Ƽ�ҽ��
'   COL_�ֽ�ʱ�䣺��ŷ��õķ���ʱ��(�޷ֽ�ʱ��ʱ)
'   COL_Ƶ�ʣ�1-"һ����"����
'   COL_��ԭʼ�Ľ��������ۼ���ʾ��
    Dim rsSend As New ADODB.Recordset
    Dim strSQL As String, lngTmp As Long, strTmp As String
    Dim lngRow As Long, lngDel��ID As Long
    Dim bln����ʱ�� As Boolean, lng���� As Long, lng��С���� As Long
    Dim str�ֽ�ʱ�� As String, dbl���� As Double, cur��� As Currency
    
    Dim vMsg As VbMsgBoxResult
    Dim blnҩƷʱ����ʾ As Boolean, blnҩƷ�����ʾ As Boolean, blnҩƷĬ�Ϸ��� As Boolean
    Dim bln����ʱ����ʾ As Boolean, bln���Ŀ����ʾ As Boolean, bln����Ĭ�Ϸ��� As Boolean
    Dim str�÷� As String, i As Long, j As Long
    Dim strͣ�� As String
    Dim str����ҽ�� As String
    Dim strNoneIDs As String
    Dim str��ҺҩƷ�ų� As String '�Ƿ���Է�����ҺҩƷ
    Dim blnҩƷ������ʾ As Boolean
    Dim str����ҽ���ų� As String
    
    Screen.MousePointer = 11
    mlngRefModld = lngModle
    stbThis.Panels(3).Text = "": stbThis.Panels(4).Text = "": Call Form_Resize
    
    vsPrice.Rows = vsPrice.FixedRows
    vsPrice.Rows = vsPrice.FixedRows + 1
    vsAdvice.Rows = vsAdvice.FixedRows '��ɾ���й���
    
    vsAdvice.ColHidden(COL_Ӥ��) = True
    Me.Refresh
    
    Call InitPriceRecordset '�Ƽ۹�ϵ��
 
    mstrAdDrugIDs = ""
    If mstrǰ��IDs = "" Then
        strNoneIDs = GetNoneSendID(mlng����ID, mlng��ҳID, 2, , , mstrAdDrugIDs)
    End If
    '��ȡ�����嵥:�¿�����У��ÿ��ҽ����¼(ҩƷ�ͷ�ҩƷ),����ҽ��Ϊ����
    '----------------------------------------------------------------------------------------------------------
    '����������ȼ�����ǰ����ҽ��������,�������ȶ�ȡ����(��ҩ;��,�÷�,�巨,�ɼ�����,��Ѫ;��)
    '����ҽ��
    If lngModle = 1 Then
        str����ҽ�� = " And NVL(a.ִ��Ƶ��,'��')='��Ҫʱ' And to_date([5],'yyyy-mm-dd hh24:mi') - a.��ʼִ��ʱ��<0.5 "
    Else
        str����ҽ�� = " And NVL(a.ִ��Ƶ��,'��')<>'��Ҫʱ' "
    End If

    '�ų���ҺҩƷ,���һ��ҩƷ����һ�����Ƿ���Ҫ��Һ�������ĵģ��Ϳ��������﷢��
    If lngModle = 0 And mbln���͵��������� Then
        str����ҽ���ų� = Get��Һ��ҽ��(mlng����ID, mlng��ҳID, 0)
        str��ҺҩƷ�ų� = " and instr(','||[8]|| ',',','||Nvl(A.���ID,A.ID)||',')=0"
    Else
        If gstr��Һ�������� <> "" And mbln���͵��������� = False Then
            '��������������������ģ����ų����еľ������̵�ҩƷ���������û�����ã���ȫԺ�����ˣ���ֻ�ų�����Ӫ��ҽ��
            str��ҺҩƷ�ų� = " And NVL(B.ִ�б��,0)<>2 And (Not Exists(Select 1 From ������ĿĿ¼ Y Where X.������Ŀid = y.Id And NVL(Y.ִ�б��,0)=2) OR x.������Ŀid is null)"
        End If
    End If
    strSQL = _
        " Select A.ID,A.���ID,Nvl(A.���ID,A.ID) as ��ID,Nvl(X.���,A.���) as ���,A.ҽ��״̬," & _
        " A.�������,F.���� as �������,A.������ĿID,B.���� as ������Ŀ,A.�շ�ϸĿID,A.Ӥ��," & _
        " A.ҽ������,A.�걾��λ,A.��鷽��,A.ִ�б��,A.����,A.�ܸ�����,D.סԺ��λ,A.��������," & _
        " Decode(A.�������,'4',C.���㵥λ,B.���㵥λ) as ���㵥λ,D.����ϵ��,D.סԺ��װ," & _
        " A.��ʼִ��ʱ��,A.ִ��Ƶ��,A.Ƶ�ʴ���,A.Ƶ�ʼ��,A.�����λ,A.ҽ������,A.ִ��ʱ�䷽��,a.�������,b.ִ�з���," & _
        " A.���˿���ID,A.��������ID,A.����ҽ��,A.����ʱ��,A.�Ƽ�����,A.ִ������,A.ִ�п���ID,Nvl(E.����,Decode(Nvl(A.ִ������,0),5,'-')) as ִ�п���," & _
        " B.��������,B.�Թܱ���,Nvl(A.�ɷ����,D.סԺ�ɷ����) as �ɷ����,Decode(A.�������,'4',G.���÷���,D.ҩ������) as ����," & _
        " G.��������,C.�Ƿ���,C.����ʱ��,C.�������,A.�¿�ǩ��ID as ǩ��ID,A.ժҪ,a.������־,c.����ʱ��,B.���㷽ʽ,b.ִ�а���,h.�������,a.��ҩ����" & _
        " From ����ҽ����¼ A,������ĿĿ¼ B,�շ���ĿĿ¼ C,ҩƷ��� D,������ҳ H,���ű� E,������Ŀ��� F,�������� G,ҩƷ���� H,����ҽ����¼ X" & _
        " Where A.����ID=[1] And A.��ҳID=[2] And h.����ID=A.����ID And h.��ҳID = a.��ҳID And Nvl(A.ǰ��ID,0) in (Select /*+cardinality(x,10)*/ x.Column_Value From Table(Cast(f_Num2list([3]) As zlTools.t_Numlist)) X)" & _
        " And A.ҽ��״̬ IN(1,3,5) And A.ҽ����Ч=1 And A.���ID=X.ID(+) And B.���=F.����" & _
        " And A.������ĿID=B.ID And A.�շ�ϸĿID=C.ID(+) And A.�շ�ϸĿID=D.ҩƷID(+) And A.�շ�ϸĿID=G.����ID(+)" & _
        " And A.ִ�п���ID=E.ID(+) And Nvl(A.ִ�б��,0)<>-1 And A.������Դ<>3" & _
        " And A.��ʼִ��ʱ�� is Not NULL And Nvl(A.ҽ��״̬,0)<>-1 and b.id=h.ҩ��id(+)" & _
        IIF(gblnKSSStrict Or gbln�����ּ����� Or gbln��Ѫ�ּ����� Or gblnѪ��ϵͳ, " And Nvl(A.���״̬,0) Not in " & IIF(gblnѪ��ϵͳ = True, " (1,3,7)", " (1,3,4,5,7)"), "") & _
        IIF(strNoneIDs <> "" And Not mbln������ҩ, " And Instr([6],','||A.ID||',')=0", "") & _
        IIF(InStr(UserInfo.����, "��ʿ") > 0, "", " And Decode(A.��˱��,1,Substr(A.����ҽ��,1,Instr(A.����ҽ��,'/')-1),Substr(A.����ҽ��,Instr(A.����ҽ��,'/')+1))=[4]") & _
        " And Exists(Select M.���� From ��Ա�� M,ִҵ��� N" & _
        " Where M.����=Decode(A.��˱��,1,Substr(A.����ҽ��,1,Instr(A.����ҽ��,'/')-1),Substr(A.����ҽ��,Instr(A.����ҽ��,'/')+1))" & _
        " And M.ִҵ���=N.���� And N.���� IN('ִҵҽʦ','ִҵ����ҽʦ'))" & _
        " And Not(A.�������='H' And B.��������='1' And B.ִ��Ƶ��=2) And Not(A.�������='Z' And B.�������� In('4','14'))" & _
         str����ҽ�� & str��ҺҩƷ�ų� & " And (h.Ӥ������ID is null or h.Ӥ������ID is not null and (h.Ӥ������ID=[7] or h.Ӥ������ID=[7]) and NVL(A.Ӥ��,0)<>0 or h.Ӥ������ID is not null and (h.Ӥ������ID<>[7] and h.Ӥ������ID<>[7]) and NVL(A.Ӥ��,0)=0) " & _
        decode(mintҽ������Χ, 1, " And nvl(a.Ӥ��,0) = 0 ", 2, " And nvl(a.Ӥ��,0) <> 0 ", "") & " Order by A.Ӥ��,���,��ID,A.���"

    On Error GoTo errH
    Set rsSend = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng��ҳID, IIF(mstrǰ��IDs = "", "0", mstrǰ��IDs), UserInfo.����, Format(dkpExecTime.value, "YYYY-MM-DD HH:mm"), _
        "," & strNoneIDs & ",", mlngҽ������ID, gstr��Һ��������, str����ҽ���ų�)
    
    '���㲢��ʾ�����嵥
    '----------------------------------------------------------------------------------------------------------
    If Not rsSend.EOF Then
        With vsAdvice
            blnҩƷʱ����ʾ = True: blnҩƷ�����ʾ = True: blnҩƷĬ�Ϸ��� = True
            bln����ʱ����ʾ = True: bln���Ŀ����ʾ = True: bln����Ĭ�Ϸ��� = True
            blnҩƷ������ʾ = True
            If strҩƷ = "" Then strҩƷ = "111" 'ҩƷ���Ժ����ҩ����Ժ��ҩ����ȡҩ
            .Redraw = flexRDNone
            For i = 1 To rsSend.RecordCount
                'һ��ҽ���е�һ�����ܷ���,�����鲻�ܷ���
                If lngDel��ID <> 0 Then
                    If NVL(rsSend!���ID, rsSend!ID) = lngDel��ID Then
                        GoTo NextLoop
                    Else
                        lngDel��ID = 0
                    End If
                End If
                
                
                '��鲻�����͵��������
                'һ��ҽ������һ��ҽ����,��������в����
                If str���s <> "" And lngDel��ID = 0 Then
                    If rsSend!������� = "7" Then
                        '��ҩ�䷽
                        If InStr(str���s, "'8'") = 0 Then lngDel��ID = rsSend!���ID
                    ElseIf InStr(",5,6,", rsSend!�������) > 0 Then
                        '������ҩ(��������ҩ���г�ҩ���һ����ҩ�����)
                        If InStr(str���s, "'" & rsSend!������� & "'") = 0 Then
                            lngDel��ID = rsSend!���ID: lng��С���� = 0
                        End If
                    ElseIf rsSend!������� = "D" Then
                        '������(������ļ��)
                        If InStr(str���s, "'D'") = 0 Then lngDel��ID = rsSend!ID
                    ElseIf rsSend!������� = "F" Then
                        '�������(�����������)
                        If InStr(str���s, "'F'") = 0 Then lngDel��ID = rsSend!ID
                    ElseIf rsSend!������� = "C" Then
                        '�������(������ļ���)
                        If InStr(str���s, "'C'") = 0 Then lngDel��ID = NVL(rsSend!���ID, rsSend!ID)
                    ElseIf rsSend!������� = "K" Then
                        '��Ѫ��Ŀ��;��(���������Ѫ)
                        If InStr(str���s, "'K'") = 0 Then lngDel��ID = rsSend!ID
                    ElseIf IsNull(rsSend!���ID) And rsSend!ID <> Val(.TextMatrix(.Rows - 1, COL_���ID)) Then
                        '����������Ŀ
                        If InStr(str���s, "'" & rsSend!������� & "'") = 0 Then lngDel��ID = rsSend!ID
                    End If
                    If lngDel��ID <> 0 Then GoTo NextLoop
                End If
                                                
                '���뵱ǰ��
                .Rows = .Rows + 1: lngRow = .Rows - 1
                .Cell(flexcpPictureAlignment, lngRow, COL_ѡ��) = 4
                Set .Cell(flexcpPicture, lngRow, COL_ѡ��) = frmIcons.imgTrueFalse.ListImages("T").Picture
                
                '�����ͣ�õģ�����ʾ���ܷ���
                If Format(NVL(rsSend!����ʱ��, "3000-1-1"), "YYYY-MM-DD") <> Format("3000-1-1", "YYYY-MM-DD") Then
                    .Cell(flexcpData, lngRow, COL_ѡ��) = 1 '��ǰ����ֹѡ��
                    Set .Cell(flexcpPicture, lngRow, COL_ѡ��) = frmIcons.imgTrueFalse.ListImages("F").Picture
                    If InStr(strͣ�� & ",", "," & rsSend!ҽ������ & ",") = 0 Then strͣ�� = strͣ�� & "," & rsSend!ҽ������
                End If
                
                '���������
                If rsSend!������� = "7" Then
                    .RowHidden(lngRow) = True '�в�ҩ
                ElseIf rsSend!������� = "E" Then
                    If Not IsNull(rsSend!���ID) Then
                        .RowHidden(lngRow) = True
                        If .TextMatrix(lngRow - 1, COL_�������) = "K" Then
                            .Cell(flexcpData, lngRow, COL_ID) = 5 '��Ѫ;��
                        Else
                            .Cell(flexcpData, lngRow, COL_ID) = 2 '��ҩ�巨
                        End If
                    ElseIf Val(.TextMatrix(lngRow - 1, COL_���ID)) = rsSend!ID Then
                        If InStr(",5,6,", .TextMatrix(lngRow - 1, COL_�������)) > 0 Then
                            .RowHidden(lngRow) = True
                            .Cell(flexcpData, lngRow, COL_ID) = 1 '��ҩ;��
                        ElseIf .TextMatrix(lngRow - 1, COL_�������) = "C" Then
                            .Cell(flexcpData, lngRow, COL_ID) = 4 '�ɼ�����
                        Else
                            .Cell(flexcpData, lngRow, COL_ID) = 3 '��ҩ�÷�
                        End If
                    End If
                ElseIf InStr(",5,6,", rsSend!�������) = 0 And Not IsNull(rsSend!���ID) Then
                    '��������,��������,��鲿λ,һ���ɼ��ļ�����Ŀ
                    .RowHidden(lngRow) = True
                End If
                
                '�ſ�һ��Ķ���(������ҩ;��,��ҩ�巨,�÷�,�ɼ�����,��Ѫ;��)
                If NVL(rsSend!ִ������, 0) = 0 Then
                    If InStr(",1,2,3,4,5,", CLng(.Cell(flexcpData, lngRow, COL_ID))) = 0 _
                        And InStr(",5,6,7,", rsSend!�������) = 0 Then
                        Call .RemoveItem(lngRow): GoTo NextLoop
                    End If
                End If
                
                'һ���и�ֵ
                '---------------------------------------------------------------
                .Cell(flexcpData, lngRow, COL_Ӥ��) = CLng(NVL(rsSend!Ӥ��, 0))
                If NVL(rsSend!Ӥ��, 0) = 0 Then
                    .TextMatrix(lngRow, COL_Ӥ��) = "����"
                Else
                    .TextMatrix(lngRow, COL_Ӥ��) = "Ӥ��" & rsSend!Ӥ��
                    .ColHidden(COL_Ӥ��) = False '��Ӥ��ҽ��ʱ����ʾ
                End If
                
                .TextMatrix(lngRow, COL_ID) = rsSend!ID
                .TextMatrix(lngRow, COL_���ID) = NVL(rsSend!���ID)
                .TextMatrix(lngRow, COL_ҽ��״̬) = rsSend!ҽ��״̬
                .TextMatrix(lngRow, COL_�������) = rsSend!�������
                .TextMatrix(lngRow, COL_������ĿID) = rsSend!������ĿID
                .TextMatrix(lngRow, col_ҽ������) = NVL(rsSend!ҽ������)
                
                .TextMatrix(lngRow, COL_�걾��λ) = NVL(rsSend!�걾��λ)
                .TextMatrix(lngRow, COL_��鷽��) = NVL(rsSend!��鷽��)
                .TextMatrix(lngRow, COL_ִ�б��) = NVL(rsSend!ִ�б��, 0)
                .TextMatrix(lngRow, COL_������־) = NVL(rsSend!������־, 0)
                If InStr(",4,5,6,7,", "," & rsSend!������� & ",") = 0 Then .TextMatrix(lngRow, COL_���㷽ʽ) = NVL(rsSend!���㷽ʽ, 0)
                .TextMatrix(lngRow, COL_ִ�а���) = NVL(rsSend!ִ�а���, 0)
                .TextMatrix(lngRow, COL_��ʼʱ��) = Format(NVL(rsSend!��ʼִ��ʱ��), "yyyy-MM-dd HH:mm:ss")
                .TextMatrix(lngRow, COL_�������) = NVL(rsSend!�������, 0)
                .TextMatrix(lngRow, COL_ִ�з���) = NVL(rsSend!ִ�з���, 0)
                .TextMatrix(lngRow, COL_��������) = NVL(rsSend!��������, 0)
                .TextMatrix(lngRow, COL_��ҩ����) = NVL(rsSend!��ҩ����)
                '����ǩ����ʶ
                .TextMatrix(lngRow, COL_ǩ��ID) = NVL(rsSend!ǩ��ID)
                If Val(.TextMatrix(lngRow, COL_ǩ��ID)) <> 0 Then
                    Set .Cell(flexcpPicture, lngRow, col_ҽ������) = frmIcons.imgSign.ListImages("ǩ��").Picture
                End If
                
                '������ʾ�Ƽ�ҽ��
                .Cell(flexcpData, lngRow, COL_�������) = CStr(NVL(rsSend!�������))
                .Cell(flexcpData, lngRow, col_ҽ������) = CStr(NVL(rsSend!������Ŀ))
                
                .TextMatrix(lngRow, COL_ҽ������) = NVL(rsSend!ҽ������)
                .Cell(flexcpData, lngRow, COL_ҽ������) = CStr(NVL(rsSend!ժҪ))
                
                .TextMatrix(lngRow, COL_ִ��ʱ��) = NVL(rsSend!ִ��ʱ�䷽��)
                .TextMatrix(lngRow, COL_Ƶ��) = NVL(rsSend!ִ��Ƶ��)
                
                .TextMatrix(lngRow, COL_���˿���ID) = NVL(rsSend!���˿���id)
                .TextMatrix(lngRow, COL_��������ID) = NVL(rsSend!��������id)
                .TextMatrix(lngRow, COL_����ҽ��) = NVL(rsSend!����ҽ��)
                .TextMatrix(lngRow, COL_����ʱ��) = Format(NVL(rsSend!����ʱ��), "yyyy-MM-dd HH:mm:ss")
                                
                '�ɼ���������ʾ������Ŀ��ִ�п���
                If Val(.Cell(flexcpData, lngRow, COL_ID)) = 4 Then
                    .TextMatrix(lngRow, COL_ִ�п���) = .TextMatrix(lngRow - 1, COL_ִ�п���)
                Else
                    .TextMatrix(lngRow, COL_ִ�п���) = NVL(rsSend!ִ�п���)
                End If
                .TextMatrix(lngRow, COL_ִ�п���ID) = NVL(rsSend!ִ�п���ID)
                
                .TextMatrix(lngRow, COL_�Ƽ�����) = NVL(rsSend!�Ƽ�����, 0)
                .TextMatrix(lngRow, COL_ִ������ID) = NVL(rsSend!ִ������, 0)
                .TextMatrix(lngRow, COL_��������) = NVL(rsSend!��������)
                                
                '�ɼ���ʽ�Ĺ�����һ���ĵ�һ��������ͬ
                If Val(.Cell(flexcpData, lngRow, COL_ID)) = 4 Then
                    j = .FindRow(CStr(rsSend!ID), .FixedRows, COL_���ID)
                    If j <> -1 Then
                        .TextMatrix(lngRow, COL_�Թܱ���) = .TextMatrix(j, COL_�Թܱ���)
                    End If
                Else
                    .TextMatrix(lngRow, COL_�Թܱ���) = NVL(rsSend!�Թܱ���)
                End If
                                
                'ҩƷ�����Ϣ
                If InStr(",5,6,7", rsSend!�������) > 0 Then
                    'ҩƷ��Ӧ�Ĺ���ѳ�����������(������Ŀ����Ҳ������ͬ����,Ŀǰ��δ����)
                    If Format(NVL(rsSend!����ʱ��, "3000-01-01"), "yyyy-MM-dd") <> "3000-01-01" Or ��InStr(",2,3,", NVL(rsSend!�������, 0)) = 0 And mlng�������� <> 1) Then
                        If rsSend!������� = "7" Then
                            strTmp = "���в�ҩ��Ӧ����ҩ�䷽�޷����ͣ�" & vbCrLf & vbCrLf & "����" & NVL(rsSend!ҽ������)
                        Else
                            strTmp = "��ҩƷ(��һ����ҩ������ҩƷ)�޷����ͣ�" & vbCrLf & vbCrLf & "����" & NVL(rsSend!ҽ������)
                        End If
                        strTmp = strTmp & vbCrLf & vbCrLf & "û�з�����Ч��ҩƷ�����Ϣ����ҩƷ�����Ѿ���ͣ�û�������סԺ���ˡ�"
                        strTmp = strTmp & vbCrLf & "���ȵ�ҩƷĿ¼�����д�����[ȷ��]������������ҽ����"
                        
                        .Redraw = flexRDDirect
                        Call .ShowCell(lngRow, COL_ѡ��)
                        Screen.MousePointer = 0
                        MsgBox strTmp, vbInformation, gstrSysName
                        
                        'ɾ����ǰ��(�������),��������һҽ��
                        Screen.MousePointer = 11
                        lngDel��ID = NVL(rsSend!���ID, rsSend!ID)
                        Call DeleteCurRow(lngRow)
                        .Refresh: .Redraw = flexRDNone
                        lng��С���� = 0: GoTo NextLoop
                    End If
                    
                    '��������ж�
                    If gbln����ҩƷ�ֿ����� Then
                        strTmp = ""
                        Select Case cboDrugType.ListIndex
                        Case 1
                            If rsSend!������� & "" <> "����ҩ" Then strTmp = "1"
                        Case 2
                            If InStr(",����ҩ,����I��,", "," & rsSend!������� & ",") = 0 Then strTmp = "1"
                        Case 3
                            If InStr(",����ҩ,����ҩ,����I��,", "," & rsSend!������� & ",") > 0 Then strTmp = "1"
                        End Select
                        
                        If strTmp <> "" Then
                            lngDel��ID = NVL(rsSend!���ID, 0)
                            Call DeleteCurRow(lngRow, rsSend!���ID)
                            lng��С���� = 0: GoTo NextLoop
                        End If
                        .TextMatrix(lngRow, COL_�������) = NVL(rsSend!�������, "��")
                    End If
                
                    .TextMatrix(lngRow, COL_�շ�ϸĿID) = rsSend!�շ�ϸĿID
                    .TextMatrix(lngRow, COL_����ϵ��) = NVL(rsSend!����ϵ��, 1)
                    .TextMatrix(lngRow, COL_סԺ��װ) = NVL(rsSend!סԺ��װ, 1)
                    .TextMatrix(lngRow, COL_סԺ��λ) = NVL(rsSend!סԺ��λ)
                    .TextMatrix(lngRow, COL_�ɷ����) = NVL(rsSend!�ɷ����, 0)
                    .TextMatrix(lngRow, COL_���) = GetStock(rsSend!�շ�ϸĿID, NVL(rsSend!ִ�п���ID, 0), 2) '��סԺ��װ
                ElseIf rsSend!������� = "4" Then
                    .TextMatrix(lngRow, COL_�շ�ϸĿID) = rsSend!�շ�ϸĿID
                    .TextMatrix(lngRow, COL_����ϵ��) = 1
                    .TextMatrix(lngRow, COL_סԺ��װ) = 1
                    .TextMatrix(lngRow, COL_סԺ��λ) = NVL(rsSend!���㵥λ)
                    .TextMatrix(lngRow, COL_���) = GetStock(rsSend!�շ�ϸĿID, NVL(rsSend!ִ�п���ID, 0), 2)
                End If
                                                                        
                '���㷢�ʹ�����ִ�еķֽ�ʱ���
                '---------------------------------------------------------------
                If rsSend!������� = "7" Then
                    .TextMatrix(lngRow, COL_����) = NVL(rsSend!�ܸ�����, 0)
                    If Not IsNull(rsSend!ִ��ʱ�䷽��) Or NVL(rsSend!�����λ) = "����" Then
                        .TextMatrix(lngRow, COL_�ֽ�ʱ��) = Calc�����ֽ�ʱ��(rsSend!�ܸ�����, rsSend!��ʼִ��ʱ��, CDate("3000-01-01"), "", NVL(rsSend!ִ��ʱ�䷽��), rsSend!Ƶ�ʴ���, rsSend!Ƶ�ʼ��, rsSend!�����λ)
                        .TextMatrix(lngRow, COL_�״�ʱ��) = Format(Split(.TextMatrix(lngRow, COL_�ֽ�ʱ��), ",")(0), "yyyy-MM-dd HH:mm")
                        .TextMatrix(lngRow, COL_ĩ��ʱ��) = Format(Split(.TextMatrix(lngRow, COL_�ֽ�ʱ��), ",")(rsSend!�ܸ����� - 1), "yyyy-MM-dd HH:mm")
                    Else
                        '�޷ֽ�ʱ��(һ��������δ����ִ��ʱ����޷��ֽ�)
                        '��¼���÷���ʱ��(��ҽ����ʼִ��ʱ��)
                        .Cell(flexcpData, lngRow, COL_�ֽ�ʱ��) = Format(rsSend!��ʼִ��ʱ��, "yyyy-MM-dd HH:mm:ss")
                    End If
                    
                    .TextMatrix(lngRow, COL_����) = NVL(rsSend!��������) '����
                    .TextMatrix(lngRow, COL_������λ) = NVL(rsSend!���㵥λ)
                    .TextMatrix(lngRow, COL_����) = NVL(rsSend!�ܸ�����, 0) '����
                    .TextMatrix(lngRow, COL_������λ) = "��"
                ElseIf InStr(",5,6,", rsSend!�������) > 0 Then
                    '����������ҩ����
                    If NVL(rsSend!Ƶ�ʴ���, 0) = 0 Or NVL(rsSend!Ƶ�ʼ��, 0) = 0 Then
                        lng���� = 1 '����Ϊһ���Ե�����ҩƷ
                    ElseIf NVL(rsSend!����, 0) <> 0 And Not IsNull(rsSend!ִ��Ƶ��) Then
                        'һ��Ƶ�����ڵĴ���
                        If rsSend!�����λ = "��" Then
                            lng���� = IntEx(rsSend!���� * (rsSend!Ƶ�ʴ��� / 7))
                        ElseIf rsSend!�����λ = "��" Then
                            lng���� = IntEx(rsSend!���� * (rsSend!Ƶ�ʴ��� / rsSend!Ƶ�ʼ��))
                        ElseIf rsSend!�����λ = "Сʱ" Then
                            lng���� = IntEx(rsSend!���� * (rsSend!Ƶ�ʴ��� / rsSend!Ƶ�ʼ��) * 24)
                        ElseIf rsSend!�����λ = "����" Then
                            lng���� = IntEx(rsSend!���� * (rsSend!Ƶ�ʴ��� / rsSend!Ƶ�ʼ��) * (24 * 60))
                        End If
                    Else
                         '�ɷ���ҩƷʱ,�������Ե����ı��������ҩ;���Ĵ���,���ɷ�����һ����ʹ��ҩƷʱ���������ԣ����������ϵ����ֵȡ�����ı��������ҩ;���Ĵ�����
                         '����һ��Ƶ�����ڵĴ�������
                        If NVL(rsSend!�ɷ����, 0) = 0 And NVL(rsSend!��������, 0) <> 0 Then
                            lng���� = IntEx(rsSend!�ܸ����� * rsSend!����ϵ�� / rsSend!��������)
                        ElseIf (NVL(rsSend!�ɷ����, 0) = 1 Or NVL(rsSend!�ɷ����, 0) = 2) And NVL(rsSend!��������, 0) <> 0 Then
                            lng���� = IntEx(rsSend!�ܸ����� / IntEx(rsSend!�������� / rsSend!����ϵ��))
                        Else
                            lng���� = NVL(rsSend!Ƶ�ʴ���, 0)
                        End If
                    End If
                    If Not IsNull(rsSend!ִ��ʱ�䷽��) Or NVL(rsSend!�����λ) = "����" Then
                        str�ֽ�ʱ�� = Calc�����ֽ�ʱ��(lng����, rsSend!��ʼִ��ʱ��, CDate("3000-01-01"), "", NVL(rsSend!ִ��ʱ�䷽��), rsSend!Ƶ�ʴ���, rsSend!Ƶ�ʼ��, rsSend!�����λ)
                        If str�ֽ�ʱ�� <> "" Then
                            .TextMatrix(lngRow, COL_�ֽ�ʱ��) = str�ֽ�ʱ��
                            .TextMatrix(lngRow, COL_�״�ʱ��) = Format(Split(str�ֽ�ʱ��, ",")(0), "yyyy-MM-dd HH:mm")
                            .TextMatrix(lngRow, COL_ĩ��ʱ��) = Format(Split(str�ֽ�ʱ��, ",")(lng���� - 1), "yyyy-MM-dd HH:mm")
                        End If
                    Else
                        '�޷ֽ�ʱ��(һ��������δ����ִ��ʱ����޷��ֽ�)
                        '��¼���÷���ʱ��(��ҽ����ʼִ��ʱ��)
                        .Cell(flexcpData, lngRow, COL_�ֽ�ʱ��) = Format(rsSend!��ʼִ��ʱ��, "yyyy-MM-dd HH:mm:ss")
                    End If
                    .TextMatrix(lngRow, COL_����) = lng����
                    .TextMatrix(lngRow, COL_����) = FormatEx(NVL(rsSend!��������), 5)
                    .TextMatrix(lngRow, COL_������λ) = NVL(rsSend!���㵥λ)
                    .TextMatrix(lngRow, COL_����) = FormatEx(rsSend!�ܸ����� / rsSend!סԺ��װ, 5) '��סԺ��λ��ʾ
                    .TextMatrix(lngRow, COL_������λ) = NVL(rsSend!סԺ��λ)
                    
                    If lng���� < lng��С���� Or lng��С���� = 0 Then lng��С���� = lng����
                ElseIf rsSend!������� = "E" And CLng(.Cell(flexcpData, lngRow, COL_ID)) <> 0 Then
                    '��ҩ;��,��ҩ�巨,��ҩ�÷�,�ɼ�����,��Ѫ;��
                    'һ����ҩ�İ���С��������(Ӱ���ҩ;���Ʒ�)
                    If .Cell(flexcpData, lngRow, COL_ID) = 1 Then '��ҩ;��
                        For j = lngRow - 1 To .FixedRows Step -1
                            If Val(.TextMatrix(j, COL_���ID)) = rsSend!ID Then
                                If Val(.TextMatrix(j, COL_����)) > lng��С���� Then
                                    .TextMatrix(j, COL_����) = lng��С����
                                    If .TextMatrix(j, COL_�ֽ�ʱ��) <> "" Then
                                        .TextMatrix(j, COL_�ֽ�ʱ��) = Trim�ֽ�ʱ��(lng��С����, .TextMatrix(j, COL_�ֽ�ʱ��))
                                        .TextMatrix(j, COL_�״�ʱ��) = Format(Split(.TextMatrix(j, COL_�ֽ�ʱ��), ",")(0), "yyyy-MM-dd HH:mm")
                                        .TextMatrix(j, COL_ĩ��ʱ��) = Format(Split(.TextMatrix(j, COL_�ֽ�ʱ��), ",")(lng��С���� - 1), "yyyy-MM-dd HH:mm")
                                    End If
                                End If
                            Else
                                Exit For
                            End If
                        Next
                        lng��С���� = 0
                    End If
                    
                    .TextMatrix(lngRow, COL_����) = .TextMatrix(lngRow - 1, COL_����) '���������
                    .TextMatrix(lngRow, COL_����) = .TextMatrix(lngRow - 1, COL_����)
                    If .Cell(flexcpData, lngRow, COL_ID) = 3 Then '��ҩ�÷�
                        .TextMatrix(lngRow, COL_������λ) = "��"
                    Else
                        .TextMatrix(lngRow, COL_������λ) = NVL(rsSend!���㵥λ)
                    End If
                    
                    .TextMatrix(lngRow, COL_�ֽ�ʱ��) = .TextMatrix(lngRow - 1, COL_�ֽ�ʱ��)
                    .Cell(flexcpData, lngRow, COL_�ֽ�ʱ��) = .Cell(flexcpData, lngRow - 1, COL_�ֽ�ʱ��)
                    If .TextMatrix(lngRow, COL_�������) = "E" And .TextMatrix(lngRow, COL_��������) = "6" And .Cell(flexcpData, lngRow - 1, COL_�ֽ�ʱ��) <> .TextMatrix(lngRow, COL_��ʼʱ��) Then
                        .Cell(flexcpData, lngRow, COL_�ֽ�ʱ��) = .TextMatrix(lngRow, COL_��ʼʱ��)
                    End If
                    .TextMatrix(lngRow, COL_�״�ʱ��) = .TextMatrix(lngRow - 1, COL_�״�ʱ��)
                    .TextMatrix(lngRow, COL_ĩ��ʱ��) = .TextMatrix(lngRow - 1, COL_ĩ��ʱ��)
                Else
                    '������ҩ����:�ɼ�����������ķ�֧����������
                    If IsNull(rsSend!���ID) Or (Not IsNull(rsSend!���ID) And rsSend!������� = "C") Then '��Ҫҽ��,�����������
                        If rsSend!������� = "K" Then
                            '��Ѫ;����ִ�д���
                            dbl���� = NVL(rsSend!�ܸ�����, 0)
                            If IsNull(rsSend!ִ��ʱ�䷽��) And (NVL(rsSend!Ƶ�ʴ���, 0) = 0 Or NVL(rsSend!Ƶ�ʼ��, 0) = 0 Or IsNull(rsSend!�����λ)) Then
                                lng���� = 1
                            Else
                                lng���� = NVL(rsSend!Ƶ�ʴ���, 1)
                            End If
                        Else
                            dbl���� = NVL(rsSend!�ܸ�����, 1)
                            lng���� = IntEx(dbl���� / NVL(rsSend!��������, 1))
                        End If
                        
                        If IsNull(rsSend!ִ��ʱ�䷽��) And (NVL(rsSend!Ƶ�ʴ���, 0) = 0 Or NVL(rsSend!Ƶ�ʼ��, 0) = 0 Or IsNull(rsSend!�����λ)) Then
                            'ִ��Ƶ��Ϊ"һ����"����Ŀ
                            str�ֽ�ʱ�� = ""
                            .Cell(flexcpData, lngRow, COL_Ƶ��) = 1
                        Else
                            'ִ��Ƶ��Ϊ"��ѡƵ��"����Ŀ:��ҽ��ʱӦ����������
                            If Not IsNull(rsSend!ִ��ʱ�䷽��) Or NVL(rsSend!�����λ) = "����" Then
                                str�ֽ�ʱ�� = Calc�����ֽ�ʱ��(lng����, rsSend!��ʼִ��ʱ��, CDate("3000-01-01"), "", NVL(rsSend!ִ��ʱ�䷽��), rsSend!Ƶ�ʴ���, rsSend!Ƶ�ʼ��, rsSend!�����λ)
                            Else
                                str�ֽ�ʱ�� = "" '����Ҳ��δ����ִ��ʱ��,�޷��ֽ�
                            End If
                        End If
                        .TextMatrix(lngRow, COL_����) = lng����
                        .TextMatrix(lngRow, COL_�ֽ�ʱ��) = str�ֽ�ʱ��
                        If str�ֽ�ʱ�� <> "" Then
                            .TextMatrix(lngRow, COL_�״�ʱ��) = Format(Split(str�ֽ�ʱ��, ",")(0), "yyyy-MM-dd HH:mm")
                            .TextMatrix(lngRow, COL_ĩ��ʱ��) = Format(Split(str�ֽ�ʱ��, ",")(lng���� - 1), "yyyy-MM-dd HH:mm")
                        Else
                            '��¼���÷���ʱ��(���޷ֽ�ʱ��ʱ),��ҽ���Ŀ�ʼִ��ʱ��
                            .Cell(flexcpData, lngRow, COL_�ֽ�ʱ��) = CStr(Format(rsSend!��ʼִ��ʱ��, "yyyy-MM-dd HH:mm:ss"))
                        End If
                        
                        .TextMatrix(lngRow, COL_����) = FormatEx(NVL(rsSend!��������), 5)
                        If Not IsNull(rsSend!��������) Then
                            .TextMatrix(lngRow, COL_������λ) = NVL(rsSend!���㵥λ)
                        End If
                        .TextMatrix(lngRow, COL_����) = IIF(dbl���� = 0, "", FormatEx(dbl����, 5))
                        .TextMatrix(lngRow, COL_������λ) = NVL(rsSend!���㵥λ)
                    Else
                        .TextMatrix(lngRow, COL_����) = .TextMatrix(lngRow - 1, COL_����)
                        .TextMatrix(lngRow, COL_����) = .TextMatrix(lngRow - 1, COL_����)
                        .TextMatrix(lngRow, COL_�ֽ�ʱ��) = .TextMatrix(lngRow - 1, COL_�ֽ�ʱ��)
                        .Cell(flexcpData, lngRow, COL_�ֽ�ʱ��) = .Cell(flexcpData, lngRow - 1, COL_�ֽ�ʱ��)
                        .TextMatrix(lngRow, COL_�״�ʱ��) = .TextMatrix(lngRow - 1, COL_�״�ʱ��)
                        .TextMatrix(lngRow, COL_ĩ��ʱ��) = .TextMatrix(lngRow - 1, COL_ĩ��ʱ��)
                    End If
                End If
                '������Ŀ���ͽ��
                cur��� = 0
                If Not LoadAdvicePrice(lngRow, rsSend, cur���) Then
                    lngDel��ID = NVL(rsSend!���ID, rsSend!ID)
                    Call DeleteCurRow(lngRow)
                    lng��С���� = 0: GoTo NextLoop
                End If
                .TextMatrix(lngRow, COL_���) = Format(cur���, gstrDec)
                .Cell(flexcpData, lngRow, COL_���) = CCur(.TextMatrix(lngRow, COL_���))
                
                '�����ʱ��һЩ�����ۼ���ʾ���,��ҩ;��,�÷�,ִ�п���,ִ������
                '---------------------------------------------------------------
                If rsSend!������� = "E" And InStr(",1,3,", Val(.Cell(flexcpData, lngRow, COL_ID))) > 0 Then '��ҩ;������ҩ�÷�
                    cur��� = 0
                    lngTmp = .FindRow(CStr(rsSend!ID), , COL_���ID)
                    
                    If .Cell(flexcpData, lngRow, COL_ID) = 1 Then '��ҩ;��
                        'һ����ҩʱ,��ҩ;���Ľ���ۼ���ʾ�ڵ�һ����ҩ��
                        .TextMatrix(lngTmp, COL_���) = Format(Val(.TextMatrix(lngTmp, COL_���)) + Val(.TextMatrix(lngRow, COL_���)), gstrDec)
                        '��ʾ��ҩ;��,ִ������
                        For j = lngTmp To lngRow - 1
                            strTmp = ""
                            If Val(.TextMatrix(j, COL_ִ������ID)) = 5 And Val(.TextMatrix(lngRow, COL_ִ������ID)) <> 5 Then
                                If Val(.TextMatrix(j, COL_ִ�б��)) = 2 Then
                                    strTmp = "��ȡҩ"
                                Else
                                    strTmp = "�Ա�ҩ"
                                End If
                            ElseIf Val(.TextMatrix(j, COL_ִ������ID)) <> 5 And Val(.TextMatrix(lngRow, COL_ִ������ID)) = 5 Then
                                strTmp = "��Ժ��ҩ"
                            Else
                                strTmp = IIF(Val(.TextMatrix(j, COL_ִ�б��)) = 1, "��ȡҩ", "")
                            End If
                            .TextMatrix(j, COL_ִ������) = strTmp
                            .TextMatrix(j, COL_�÷�) = rsSend!������Ŀ
                            
                            'ҩƷ�����������
                            If strҩƷ <> "111" Then
                                If Val(Mid(strҩƷ, 2, 1)) = 0 And strTmp = "��Ժ��ҩ" _
                                    Or Val(Mid(strҩƷ, 3, 1)) = 0 And strTmp = "��ȡҩ" _
                                    Or Val(Mid(strҩƷ, 1, 1)) = 0 And strTmp <> "��Ժ��ҩ" And strTmp <> "��ȡҩ" Then
                                    lngDel��ID = NVL(rsSend!���ID, rsSend!ID)
                                    Call DeleteCurRow(lngRow)
                                    lng��С���� = 0: GoTo NextLoop
                                End If
                            End If
                        Next
                    Else
                        'ҩƷ��ִ������
                        strTmp = ""
                        If Val(.TextMatrix(lngTmp, COL_ִ������ID)) = 5 And Val(.TextMatrix(lngRow, COL_ִ������ID)) <> 5 Then
                            If Val(.TextMatrix(lngTmp, COL_ִ�б��)) = 2 Then
                                strTmp = "��ȡҩ"
                            Else
                                strTmp = "�Ա�ҩ"
                            End If
                        ElseIf Val(.TextMatrix(lngTmp, COL_ִ������ID)) <> 5 And Val(.TextMatrix(lngRow, COL_ִ������ID)) = 5 Then
                            strTmp = "��Ժ��ҩ"
                        Else
                            strTmp = IIF(Val(.TextMatrix(lngTmp, COL_ִ�б��)) = 1, "��ȡҩ", "")
                        End If
                    
                        'ҩƷ�����������
                        If strҩƷ <> "111" Then
                            If Val(Mid(strҩƷ, 2, 1)) = 0 And strTmp = "��Ժ��ҩ" _
                                Or Val(Mid(strҩƷ, 3, 1)) = 0 And strTmp = "��ȡҩ" _
                                Or Val(Mid(strҩƷ, 1, 1)) = 0 And strTmp <> "��Ժ��ҩ" And strTmp <> "��ȡҩ" Then
                                lngDel��ID = NVL(rsSend!���ID, rsSend!ID)
                                Call DeleteCurRow(lngRow)
                                lng��С���� = 0: GoTo NextLoop
                            End If
                        End If
                    
                        '��ҩ�÷�,�巨
                        str�÷� = rsSend!������Ŀ
                        If Val(.Cell(flexcpData, lngRow - 1, COL_ID)) = 2 Then
                            str�÷� = str�÷� & "|" & sys.RowValue("������ĿĿ¼", Val(.TextMatrix(lngRow - 1, COL_������ĿID)), "����")
                        End If
                        For j = lngTmp To lngRow
                            .TextMatrix(j, COL_�÷�) = str�÷� '������д�շ���¼
                            cur��� = cur��� + Val(.TextMatrix(j, COL_���))
                        Next
                        .TextMatrix(lngRow, COL_���) = Format(cur���, gstrDec)
                        '��ʾִ������
                        .TextMatrix(lngRow, COL_ִ������) = strTmp
                        '��ʾ�䷽ִ�п���
                        .TextMatrix(lngRow, COL_ִ�п���) = .TextMatrix(lngTmp, COL_ִ�п���)
                    End If
                    
                    'ʹ���ҽ��ѡ��״̬��ͬ(��Ϊ����ԭ�򣻷�ҩҽ������)
                    For j = lngTmp To lngRow
                        If .Cell(flexcpData, j, COL_ѡ��) <> 0 Then
                            Call RowSelectSame(j, COL_ѡ��)
                            Exit For 'һ����ֹ,ȫ����ֹ
                        End If
                    Next
                    If j > lngRow Then
                        For j = lngRow To lngTmp Step -1
                            If InStr(",5,6,7,", .TextMatrix(j, COL_�������)) > 0 Then
                                If .Cell(flexcpPicture, j, COL_ѡ��) Is Nothing Then
                                    Call RowSelectSame(j, COL_ѡ��)
                                    Exit For '���ѡ,ȫ����ѡ
                                End If
                            End If
                        Next
                    End If
                ElseIf InStr(",5,6,7,", rsSend!�������) = 0 Then
                    If Not IsNull(rsSend!���ID) And rsSend!������� <> "C" Then
                        '������ҩҽ��
                        For j = lngRow - 1 To .FixedRows Step -1
                            If Val(.TextMatrix(j, COL_ID)) = rsSend!���ID Then
                                .TextMatrix(j, COL_���) = Format(Val(.TextMatrix(j, COL_���)) + Val(.TextMatrix(lngRow, COL_���)), gstrDec)
                                Exit For
                            End If
                        Next
                        
                        '��Ѫ;��
                        If rsSend!������� = "E" And Val(.Cell(flexcpData, lngRow, COL_ID)) = 5 Then
                            .TextMatrix(lngRow - 1, COL_�÷�) = rsSend!������Ŀ
                        End If
                    ElseIf Val(.Cell(flexcpData, lngRow, COL_ID)) = 4 Then
                        '����걾�ɼ�����Ϊ��ʾ��
                        .TextMatrix(lngRow, COL_�÷�) = rsSend!������Ŀ
                        For j = lngRow - 1 To .FixedRows Step -1
                            If Val(.TextMatrix(j, COL_���ID)) = rsSend!ID Then
                                .TextMatrix(lngRow, COL_���) = Format(Val(.TextMatrix(lngRow, COL_���)) + Val(.TextMatrix(j, COL_���)), gstrDec)
                            Else
                                Exit For
                            End If
                        Next
                    End If
                End If

                'ҩƷ�����Ŀ����(0-�����;1-���,��������;2-��飬�����ֹ),�Ա�ҩ�����
                '---------------------------------------------------------------
                If InStr(",5,6,7,", rsSend!�������) > 0 And NVL(rsSend!ִ������, 0) <> 5 Then
                    Call CheckStock(lngRow, rsSend, blnҩƷ�����ʾ, blnҩƷʱ����ʾ, blnҩƷĬ�Ϸ���)
                    Call CheckDrug����(lngRow, blnҩƷ������ʾ)
                ElseIf rsSend!������� = "4" And NVL(rsSend!��������, 0) = 1 Then
                    Call CheckStock(lngRow, rsSend, bln���Ŀ����ʾ, bln����ʱ����ʾ, bln����Ĭ�Ϸ���)
                End If
                
NextLoop:       '---------------------------------------------------------------
                Progress = i / rsSend.RecordCount * 100
                txtPer.Text = CInt(psb.value) & "%"
                txtPer.Refresh
                rsSend.MoveNext
            Next
        End With
    End If
    With vsAdvice
        .AutoSize col_ҽ������
        .RowHeight(0) = 320
        If .Rows = .FixedRows Then .Rows = .FixedRows + 1
        
        '����ǩ��ͼ�����
        .Cell(flexcpPictureAlignment, .FixedRows, col_ҽ������, .Rows - 1, col_ҽ������) = 0
        
        .Col = .FixedCols
        For i = .FixedRows To .Rows - 1
            If Not .RowHidden(i) Then
                .Row = i: Exit For
            End If
        Next
        
        Call .ShowCell(.Row, .Col)
        .Redraw = flexRDDirect
        
        '�����ͣ�õ���Ŀ������ʾ
        If strͣ�� <> "" Then
            Call MsgBox("������Ŀ��" & Mid(strͣ��, 2) & " �Ѿ�ͣ�ã����ܷ��͡�", vbInformation, Me.Caption)
        End If
        
        Call vsAdvice_AfterRowColChange(-1, -1, .Row, .Col)
    End With
    
    Call ShowSendTotal
    Progress = 0: Screen.MousePointer = 0
    LoadAdviceSend = True
    Exit Function
errH:
    vsAdvice.Redraw = flexRDDirect
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        vsAdvice.Redraw = flexRDNone
        Resume
    End If
    Call SaveErrLog
    Progress = 0
End Function

Private Sub CheckDrug����(ByVal lngRow As Long, ByRef bln��ʾ As Boolean)
'���ܣ����͹����ж�����ҩƷ���м���ֹ
    Dim strTmp As String
    Dim blnTmp As Boolean
    Dim vMsg As VbMsgBoxResult
    
    With vsAdvice
        If 0 <> Val(.TextMatrix(lngRow, COL_�շ�ϸĿID)) And 0 <> Val(.TextMatrix(lngRow, COL_ִ�п���ID)) And .Cell(flexcpData, lngRow, COL_ѡ��) <> 1 Then
            If InitObjPublicDrug Then
                blnTmp = gobjPublicDrug.zlCheckPriceAdjustBySell(Val(.TextMatrix(lngRow, COL_�շ�ϸĿID)), Val(.TextMatrix(lngRow, COL_ִ�п���ID)), False)
                If Not blnTmp Then
                    strTmp = "��(" & .TextMatrix(lngRow, COL_ִ�п���) & ")��ҩƷ""" & .TextMatrix(lngRow, col_ҽ������) & """" & vbCrLf & vbCrLf & _
                        "���������۹����Ҫ�󣺳ɱ��ۺ��ۼ۲�һ�£��������۳��⡣" & vbCrLf & vbCrLf & _
                        "����ϵҩ����ҩ���ƽ��е��۴���"
                    
                    If bln��ʾ Then
                        .Redraw = flexRDDirect:
                        Call .ShowCell(lngRow, COL_ѡ��)
                        Screen.MousePointer = 0
                        vMsg = frmMsgBox.ShowMsgBox(strTmp, Me, True)
                        If vMsg = vbIgnore Then bln��ʾ = False
                        .Cell(flexcpData, lngRow, COL_ѡ��) = 1 '��ǰ����ֹѡ��
                        Set .Cell(flexcpPicture, lngRow, COL_ѡ��) = frmIcons.imgTrueFalse.ListImages("F").Picture
                        Screen.MousePointer = 11
                        .Refresh: .Redraw = flexRDNone
                    Else
                        .Cell(flexcpData, lngRow, COL_ѡ��) = 1 '��ǰ����ֹѡ��
                        Set .Cell(flexcpPicture, lngRow, COL_ѡ��) = frmIcons.imgTrueFalse.ListImages("F").Picture
                    End If
                End If
            End If
        End If
    End With
End Sub

Private Sub CheckStock(ByVal lngRow As Long, rsSend As ADODB.Recordset, Optional bln�����ʾ As Boolean, Optional blnʱ����ʾ As Boolean, Optional blnĬ�Ϸ��� As Boolean)
'���ܣ����ݿ���������鷢��ҩƷ���������ĵĿ��
'������lngRow=ҽ���к�,rsSend=��ǰ����ҽ����Ϣ
'      bln�����ʾ,blnʱ����ʾ,blnĬ�Ϸ���=������ʾ�������ʾ����
'���أ�������ʾ���Ƿ��ѡ��״̬�����˴���
    Dim int����� As Integer, dbl���� As Double
    Dim dbl���ÿ�� As Double, dbl�ѷ���� As Double
    Dim bln����ʱ�� As Boolean, bln���� As Boolean, blnʱ�� As Boolean
    Dim vMsg As VbMsgBoxResult, strTmp As String
    Dim blnDo As Boolean, i As Long
    
    With vsAdvice
        'ҩƷ�����(0-�����;1-���,��������;2-��飬�����ֹ)
        int����� = TheStockCheck(Val(.TextMatrix(lngRow, COL_ִ�п���ID)), .TextMatrix(lngRow, COL_�������))
        bln���� = NVL(rsSend!����, 0) = 1
        blnʱ�� = NVL(rsSend!�Ƿ���, 0) = 1
        
        '������ʱ��ҩƷ����Ҫ���㹻�Ŀ��,�������ݿ�����������
        If int����� <> 0 Or bln���� Or blnʱ�� Then
            strTmp = .TextMatrix(lngRow, COL_סԺ��λ) '������ɢװ��λ
            
            '������Ͳ����ֹʱ,����ʱ��Ͳ��ص�������
            bln����ʱ�� = int����� <> 2 And (bln���� Or blnʱ��)
            
            '��ǰҩƷ����:סԺ��װ
            If .TextMatrix(lngRow, COL_�������) = "7" Then
                '��ҩҩ����λ�����ɷ��㴦��:ÿ��
                If Val(.TextMatrix(lngRow, COL_�ɷ����)) = 0 Then
                    dbl���� = Val(.TextMatrix(lngRow, COL_����)) * Val(.TextMatrix(lngRow, COL_����))
                    dbl���� = dbl���� / Val(.TextMatrix(lngRow, COL_����ϵ��)) / Val(.TextMatrix(lngRow, COL_סԺ��װ))
                Else
                    dbl���� = IntEx(Val(.TextMatrix(lngRow, COL_����)) / Val(.TextMatrix(lngRow, COL_����ϵ��)) / Val(.TextMatrix(lngRow, COL_סԺ��װ)))
                    dbl���� = dbl���� * Val(.TextMatrix(lngRow, COL_����))
                End If
            Else
                dbl���� = Val(.TextMatrix(lngRow, COL_����))
            End If
            
            '��ǰ���ÿ��:סԺ��װ,��ȥǰ����ͬҩƷҪ���͵Ŀ��
            For i = lngRow - 1 To .FixedRows Step -1
                If rsSend!������� = "4" Then
                    blnDo = .TextMatrix(i, COL_�������) = "4"
                Else
                    blnDo = InStr(",5,6,7,", .TextMatrix(i, COL_�������)) > 0
                End If
                If blnDo Then
                    blnDo = Val(.TextMatrix(i, COL_�շ�ϸĿID)) = Val(.TextMatrix(lngRow, COL_�շ�ϸĿID)) _
                        And Val(.TextMatrix(i, COL_ִ�п���ID)) = Val(.TextMatrix(lngRow, COL_ִ�п���ID))
                End If
                If blnDo Then
                    blnDo = .Cell(flexcpData, i, COL_ѡ��) = 0 And Not .Cell(flexcpPicture, i, COL_ѡ��) Is Nothing
                End If
                If blnDo Then
                    If .TextMatrix(i, COL_�������) = "7" Then
                        '��ҩҩ����λ�����ɷ��㴦��:ÿ��
                        If Val(.TextMatrix(i, COL_�ɷ����)) = 0 Then
                            dbl�ѷ���� = dbl�ѷ���� + _
                                Val(.TextMatrix(i, COL_����)) * Val(.TextMatrix(i, COL_����)) _
                                / Val(.TextMatrix(i, COL_����ϵ��)) / Val(.TextMatrix(i, COL_סԺ��װ))
                        Else
                            dbl�ѷ���� = dbl�ѷ���� + Val(.TextMatrix(i, COL_����)) _
                                * IntEx(Val(.TextMatrix(i, COL_����)) / Val(.TextMatrix(i, COL_����ϵ��)) / Val(.TextMatrix(i, COL_סԺ��װ)))
                        End If
                    Else
                        dbl�ѷ���� = dbl�ѷ���� + Val(.TextMatrix(i, COL_����))
                    End If
                End If
            Next
            dbl���ÿ�� = Val(.TextMatrix(lngRow, COL_���))
            dbl���ÿ�� = dbl���ÿ�� - dbl�ѷ����
            
            If dbl���� > dbl���ÿ�� Then
                If (Not bln����ʱ�� And int����� <> 0 And bln�����ʾ) Or (bln����ʱ�� And blnʱ����ʾ) Then
                    '��һ��û��ѡ������ʾ,����ʾ
                    If bln����ʱ�� Then
                        If InStr(GetInsidePrivs(pסԺҽ���´�), "��ʾҩƷ���") = 0 Then
                            strTmp = "������ʱ��ҩƷ""" & .TextMatrix(lngRow, col_ҽ������) & """��" & vbCrLf & vbCrLf & _
                                "��" & .TextMatrix(lngRow, COL_ִ�п���) & "��治��" & _
                                IIF(dbl�ѷ���� <> 0, "(�ſ�ǰ����ͬҩƷ������)", "") & "��" & _
                                "���η�������" & FormatEx(dbl����, 5) & strTmp & "��"
                        Else
                            strTmp = "������ʱ��ҩƷ""" & .TextMatrix(lngRow, col_ҽ������) & """��治�㣺" & vbCrLf & vbCrLf & _
                                .TextMatrix(lngRow, COL_ִ�п���) & "���ÿ�棺" & FormatEx(dbl���ÿ��, 5) & strTmp & _
                                IIF(dbl�ѷ���� <> 0, "(�ſ�ǰ����ͬҩƷ������)", "") & "��" & _
                                "���η�������" & FormatEx(dbl����, 5) & strTmp & "��"
                        End If
                    Else
                        If InStr(GetInsidePrivs(pסԺҽ���´�), "��ʾҩƷ���") = 0 Then
                            strTmp = "ҩƷ""" & .TextMatrix(lngRow, col_ҽ������) & """��" & vbCrLf & vbCrLf & _
                                "��" & .TextMatrix(lngRow, COL_ִ�п���) & "��治��" & _
                                IIF(dbl�ѷ���� <> 0, "(�ſ�ǰ����ͬҩƷ������)", "") & "��" & _
                                "���η�������" & FormatEx(dbl����, 5) & strTmp & "��"
                        Else
                            strTmp = "ҩƷ""" & .TextMatrix(lngRow, col_ҽ������) & """��治�㣺" & vbCrLf & vbCrLf & _
                                .TextMatrix(lngRow, COL_ִ�п���) & "���ÿ�棺" & FormatEx(dbl���ÿ��, 5) & strTmp & _
                                IIF(dbl�ѷ���� <> 0, "(�ſ�ǰ����ͬҩƷ������)", "") & "��" & _
                                "���η�������" & FormatEx(dbl����, 5) & strTmp & "��"
                        End If
                    End If
                    If int����� = 1 And Not bln����ʱ�� Then
                        strTmp = strTmp & vbCrLf & vbCrLf & "Ҫ���͸�ҩƷ��"
                    End If
                    If rsSend!������� = "4" Then
                        strTmp = Replace(strTmp, "ҩƷ", "����")
                    End If
                    
                    .Redraw = flexRDDirect:
                    Call .ShowCell(lngRow, COL_ѡ��)
                    Screen.MousePointer = 0
                    vMsg = frmMsgBox.ShowMsgBox(strTmp, Me, int����� = 2 Or bln����ʱ��)
                    
                    If bln����ʱ�� Then
                        If vMsg = vbIgnore Then blnʱ����ʾ = False
                        .Cell(flexcpData, lngRow, COL_ѡ��) = 1 '��ǰ����ֹѡ��
                        Set .Cell(flexcpPicture, lngRow, COL_ѡ��) = frmIcons.imgTrueFalse.ListImages("F").Picture
                    ElseIf int����� = 2 Then '����ֹ
                        If vMsg = vbIgnore Then bln�����ʾ = False
                        .Cell(flexcpData, lngRow, COL_ѡ��) = 1 '��ǰ����ֹѡ��
                        Set .Cell(flexcpPicture, lngRow, COL_ѡ��) = frmIcons.imgTrueFalse.ListImages("F").Picture
                    ElseIf int����� = 1 Then '�������
                        If vMsg = vbYes Or vMsg = vbIgnore Then
                            If vMsg = vbIgnore Then bln�����ʾ = False
                            blnĬ�Ϸ��� = True
                        ElseIf vMsg = vbNo Or vMsg = vbCancel Then
                            If vMsg = vbCancel Then bln�����ʾ = False
                            blnĬ�Ϸ��� = False
                            Set .Cell(flexcpPicture, lngRow, COL_ѡ��) = Nothing 'ȱʡ������
                        End If
                    End If
                    
                    Screen.MousePointer = 11
                    .Refresh: .Redraw = flexRDNone
                Else
                    '��һ��ѡ���˲�����ʾ
                    If int����� = 2 Or bln���� Or blnʱ�� Then
                        .Cell(flexcpData, lngRow, COL_ѡ��) = 1 '��ǰ����ֹѡ��
                        Set .Cell(flexcpPicture, lngRow, COL_ѡ��) = frmIcons.imgTrueFalse.ListImages("F").Picture
                    ElseIf int����� = 1 Then
                        '������һ�εĽ������
                        If Not blnĬ�Ϸ��� Then
                            Set .Cell(flexcpPicture, lngRow, COL_ѡ��) = Nothing 'ȱʡ������
                        End If
                    End If
                End If
            End If
        End If
    End With
End Sub

Private Function CheckPriceStock(ByVal lngRow As Long, rsPrice As ADODB.Recordset, ByVal lng�ⷿID As Long, ByVal dbl���� As Double, _
    rsTotal As ADODB.Recordset, Optional bln�����ʾ As Boolean, Optional blnʱ����ʾ As Boolean, Optional blnĬ�Ϸ��� As Boolean) As Boolean
'���ܣ����͹�����ʱ���Է�ҩ��ҩƷ���������õ����ļƼ۽��п����(�ۼƼ��)
'������lngRow=ҽ���к�
'      dbl����=�Ѽ���õļƼ�����(�ۼ۵�λ)
'      rsTotal=��ǰ����ǰ�����ۼƷ��͵ļƼ�ҩƷ����������(�ۼ۵�λ)
'      bln�����ʾ,blnʱ����ʾ,blnĬ�Ϸ���=������ʾ�������ʾ����
'���أ�������ʾ���Ƿ��ѡ��״̬�����˴���
    Dim int����� As Integer, dbl���� As Double
    Dim dbl���ÿ�� As Double, dbl�ѷ���� As Double
    Dim bln����ʱ�� As Boolean, bln���� As Boolean, blnʱ�� As Boolean
    Dim vMsg As VbMsgBoxResult, strTmp As String
    Dim blnDo As Boolean, i As Long
    
    With vsAdvice
        'ҩƷ�����(0-�����;1-���,��������;2-��飬�����ֹ)
        int����� = TheStockCheck(lng�ⷿID, rsPrice!���)
        bln���� = NVL(rsPrice!����, 0) = 1
        blnʱ�� = NVL(rsPrice!�Ƿ���, 0) = 1
        
        '������ʱ��ҩƷ����Ҫ���㹻�Ŀ��,�������ݿ�����������
        If int����� <> 0 Or bln���� Or blnʱ�� Then
            strTmp = NVL(rsPrice!סԺ��λ, NVL(rsPrice!���㵥λ)) '������ʾ
            
            '������Ͳ����ֹʱ,����ʱ��Ͳ��ص�������
            bln����ʱ�� = int����� <> 2 And (bln���� Or blnʱ��)
            
            '��ǰҩƷ����������:סԺ��װ
            dbl���� = Format(dbl���� / NVL(rsPrice!סԺ��װ, 1), "0.00000")
            
            '��ǰ���ÿ��:סԺ��װ,��ȥǰ����ͬҩƷҽ��Ҫ���͵Ŀ��
            If InStr(",5,6,7,", rsPrice!���) > 0 Then
                For i = lngRow - 1 To .FixedRows Step -1
                    blnDo = InStr(",5,6,7,", .TextMatrix(i, COL_�������)) > 0
                    If blnDo Then
                        blnDo = Val(.TextMatrix(i, COL_�շ�ϸĿID)) = rsPrice!ID And Val(.TextMatrix(i, COL_ִ�п���ID)) = lng�ⷿID
                    End If
                    If blnDo Then
                        blnDo = .Cell(flexcpData, i, COL_ѡ��) = 0 And Not .Cell(flexcpPicture, i, COL_ѡ��) Is Nothing
                    End If
                    If blnDo Then
                        If .TextMatrix(i, COL_�������) = "7" Then
                            '��ҩҩ����λ�����ɷ��㴦��:ÿ��
                            If Val(.TextMatrix(i, COL_�ɷ����)) = 0 Then
                                dbl�ѷ���� = dbl�ѷ���� + _
                                    Val(.TextMatrix(i, COL_����)) * Val(.TextMatrix(i, COL_����)) _
                                    / Val(.TextMatrix(i, COL_����ϵ��)) / Val(.TextMatrix(i, COL_סԺ��װ))
                            Else
                                dbl�ѷ���� = dbl�ѷ���� + Val(.TextMatrix(i, COL_����)) _
                                    * IntEx(Val(.TextMatrix(i, COL_����)) / Val(.TextMatrix(i, COL_����ϵ��)) / Val(.TextMatrix(i, COL_סԺ��װ)))
                            End If
                        Else
                            dbl�ѷ���� = dbl�ѷ���� + Val(.TextMatrix(i, COL_����))
                        End If
                    End If
                Next
            End If
            '�Ƽ۲���Ҫ���͵��ۼ�����
            rsTotal.Filter = "��ĿID=" & rsPrice!ID & " And �ⷿID=" & lng�ⷿID
            Do While Not rsTotal.EOF
                dbl�ѷ���� = dbl�ѷ���� + Format(rsTotal!���� / NVL(rsPrice!סԺ��װ, 1), "0.00000")
                rsTotal.MoveNext
            Loop
            
            dbl���ÿ�� = Format(GetStock(rsPrice!ID, lng�ⷿID, 2), "0.00000")
            dbl���ÿ�� = dbl���ÿ�� - dbl�ѷ����
            
            If dbl���� > dbl���ÿ�� Then
                If (Not bln����ʱ�� And int����� <> 0 And bln�����ʾ) Or (bln����ʱ�� And blnʱ����ʾ) Then
                    '��һ��û��ѡ������ʾ,����ʾ
                    If bln����ʱ�� Then
                        If InStr(GetInsidePrivs(pסԺҽ���´�), "��ʾҩƷ���") = 0 Then
                            strTmp = "ҽ��""" & .TextMatrix(lngRow, col_ҽ������) & """�ķ�����ʱ�ۼƼ���Ŀ��" & vbCrLf & vbCrLf & _
                                """" & rsPrice!���� & """��" & sys.RowValue("���ű�", lng�ⷿID, "����") & "��治��" & _
                                IIF(dbl�ѷ���� <> 0, "(�ſ�ǰ����ͬ��Ŀ������)", "") & "�����η���������" & FormatEx(dbl����, 5) & strTmp & "��"
                        Else
                            strTmp = "ҽ��""" & .TextMatrix(lngRow, col_ҽ������) & """�ķ�����ʱ�ۼƼ���Ŀ""" & rsPrice!���� & """��治�㣺" & _
                                vbCrLf & vbCrLf & sys.RowValue("���ű�", lng�ⷿID, "����") & "���ÿ�棺" & FormatEx(dbl���ÿ��, 5) & strTmp & _
                                IIF(dbl�ѷ���� <> 0, "(�ſ�ǰ����ͬ��Ŀ������)", "") & "�����η���������" & FormatEx(dbl����, 5) & strTmp & "��"
                        End If
                    Else
                        If InStr(GetInsidePrivs(pסԺҽ���´�), "��ʾҩƷ���") = 0 Then
                            strTmp = "ҽ��""" & .TextMatrix(lngRow, col_ҽ������) & """�ļƼ���Ŀ��" & vbCrLf & vbCrLf & _
                                """" & rsPrice!���� & """��" & sys.RowValue("���ű�", lng�ⷿID, "����") & "��治��" & _
                                IIF(dbl�ѷ���� <> 0, "(�ſ�ǰ����ͬ��Ŀ������)", "") & "�����η���������" & FormatEx(dbl����, 5) & strTmp & "��"
                        Else
                            strTmp = "ҽ��""" & .TextMatrix(lngRow, col_ҽ������) & """�ļƼ���Ŀ""" & rsPrice!���� & """��治�㣺" & _
                                vbCrLf & vbCrLf & sys.RowValue("���ű�", lng�ⷿID, "����") & "���ÿ�棺" & FormatEx(dbl���ÿ��, 5) & strTmp & _
                                IIF(dbl�ѷ���� <> 0, "(�ſ�ǰ����ͬ��Ŀ������)", "") & "�����η���������" & FormatEx(dbl����, 5) & strTmp & "��"
                        End If
                    End If
                    If int����� = 1 And Not bln����ʱ�� Then
                        strTmp = strTmp & vbCrLf & vbCrLf & "Ҫ���͸�ҽ����"
                    End If
                    
                    .Redraw = flexRDDirect
                    .Row = GetVisibleRow(lngRow, True)
                    Call .ShowCell(.Row, COL_ѡ��)
                    Screen.MousePointer = 0
                    vMsg = frmMsgBox.ShowMsgBox(strTmp, Me, int����� = 2 Or bln����ʱ��)
                    
                    If bln����ʱ�� Then
                        If vMsg = vbIgnore Then blnʱ����ʾ = False
                        .Cell(flexcpData, lngRow, COL_ѡ��) = 1 '��ǰ����ֹѡ��
                        Set .Cell(flexcpPicture, lngRow, COL_ѡ��) = frmIcons.imgTrueFalse.ListImages("F").Picture
                        CheckPriceStock = True
                    ElseIf int����� = 2 Then '����ֹ
                        If vMsg = vbIgnore Then bln�����ʾ = False
                        .Cell(flexcpData, lngRow, COL_ѡ��) = 1 '��ǰ����ֹѡ��
                        Set .Cell(flexcpPicture, lngRow, COL_ѡ��) = frmIcons.imgTrueFalse.ListImages("F").Picture
                        CheckPriceStock = True
                    ElseIf int����� = 1 Then '�������
                        If vMsg = vbYes Or vMsg = vbIgnore Then
                            If vMsg = vbIgnore Then bln�����ʾ = False
                            blnĬ�Ϸ��� = True
                        ElseIf vMsg = vbNo Or vMsg = vbCancel Then
                            If vMsg = vbCancel Then bln�����ʾ = False
                            blnĬ�Ϸ��� = False
                            Set .Cell(flexcpPicture, lngRow, COL_ѡ��) = Nothing 'ȱʡ������
                            CheckPriceStock = True
                        End If
                    End If
                    Screen.MousePointer = 11
                    .Refresh: .Redraw = flexRDNone
                Else
                    '��һ��ѡ���˲�����ʾ
                    If int����� = 2 Or bln���� Or blnʱ�� Then
                        .Cell(flexcpData, lngRow, COL_ѡ��) = 1 '��ǰ����ֹѡ��
                        Set .Cell(flexcpPicture, lngRow, COL_ѡ��) = frmIcons.imgTrueFalse.ListImages("F").Picture
                        CheckPriceStock = True
                    ElseIf int����� = 1 Then
                        '������һ�εĽ������
                        If Not blnĬ�Ϸ��� Then
                            Set .Cell(flexcpPicture, lngRow, COL_ѡ��) = Nothing 'ȱʡ������
                            CheckPriceStock = True
                        End If
                    End If
                End If
            End If
        End If
        
        '���δ��ʾ��Ҫ����,�����ۼƷ�������
        If Not CheckPriceStock Then
            rsTotal.AddNew
            If Val(.TextMatrix(lngRow, COL_���ID)) <> 0 Then
                rsTotal!ҽ��ID = Val(.TextMatrix(lngRow, COL_���ID))
            Else
                rsTotal!ҽ��ID = Val(.TextMatrix(lngRow, COL_ID))
            End If
            rsTotal!��ĿID = rsPrice!ID
            rsTotal!�ⷿID = lng�ⷿID
            rsTotal!���� = dbl����
            rsTotal.Update
        End If
    End With
End Function

Private Sub vsPrice_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, lng�к� As Long, i As Long
    Dim str��ĿIDs As String, blnCancel As Boolean
    Dim lngҽ��ID As Long, lngԭ��ĿID As Long
    Dim int�������� As Integer, vPoint As PointAPI
    Dim strSQL2 As String
    
    With vsPrice
        lng�к� = Val(.TextMatrix(Row, COLP_�к�))
        If Col = COLP_�շ���Ŀ Then
            '����ѡ�����е���Ŀ
            For i = .FixedRows To .Rows - 1
                If Val(.TextMatrix(i, COLP_�к�)) = lng�к� And lng�к� <> 0 And i <> Row Then
                    str��ĿIDs = str��ĿIDs & "," & Val(.TextMatrix(i, COLP_�շ�ϸĿID))
                End If
            Next
            str��ĿIDs = Mid(str��ĿIDs, 2)
            
            strSQL = _
                " Select Distinct 0 as ĩ��,To_Number('999999999'||����) as ID,-NULL as �ϼ�ID," & _
                " CHR(13)||���� as ����,Decode(����,1,'����ҩ',2,'�г�ҩ',3,'�в�ҩ',7,'��������') as ����," & _
                " NULL as ��λ,NULL as ���,NULL as ����,NULL as ���,NULL as ��������,NULL as ҽ������,NULL as ˵��,NULL as �۸�," & _
                " -NULL as ԭ��ID,-NULL as �ּ�ID,-NULL as ȱʡ�۸�ID,-NULL as �Ƿ���ID,Null as ���ID,-Null as ��������ID" & _
                " From ���Ʒ���Ŀ¼ Where ���� in (1,2,3,7) And (����ʱ�� Is Null Or ����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))"
            strSQL = strSQL & " Union ALL " & _
                " Select 0 as ĩ��,-ID as ID,Nvl(-�ϼ�ID,To_Number('999999999'||����)) as �ϼ�ID,����,����," & _
                " NULL as ��λ,NULL as ���,NULL as ����,NULL as ���,NULL as ��������,NULL as ҽ������,NULL as ˵��,NULL as �۸�," & _
                " -NULL as ԭ��ID,-NULL as �ּ�ID,-NULL as ȱʡ�۸�ID,-NULL as �Ƿ���ID,Null as ���ID,-Null as ��������ID" & _
                " From ���Ʒ���Ŀ¼ Where ���� in (1,2,3,7) And (����ʱ�� Is Null Or ����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                " Start With �ϼ�ID is NULL Connect by Prior ID=�ϼ�ID"
            strSQL = strSQL & " Union ALL " & _
                " Select 0 as ĩ��,ID,�ϼ�ID,����,����,NULL as ��λ,NULL as ���,NULL as ����,NULL as ���,NULL as ��������,NULL as ҽ������," & _
                " NULL as ˵��,NULL as �۸�,-NULL as ԭ��ID,-NULL as �ּ�ID,-NULL as ȱʡ�۸�ID,-NULL as �Ƿ���ID,Null as ���ID,-Null as ��������ID" & _
                " From �շѷ���Ŀ¼ Start With �ϼ�ID is NULL Connect by Prior ID=�ϼ�ID"
            strSQL2 = _
                " Select ĩ��,ID,�ϼ�ID,����,����,��λ,���,����,���,��������,ҽ������,˵��," & _
                " Decode(Nvl(�Ƿ���,0),1,Decode(Instr('567',���ID),0,Sum(Nvl(ԭ��,0))||'-'||Sum(Nvl(�ּ�,0)),'ʱ��'),Sum(�ּ�)) as �۸�," & _
                " Sum(ԭ��) as ԭ��ID,Sum(�ּ�) as �ּ�ID,Sum(ȱʡ�۸�) as ȱʡ�۸�ID,�Ƿ��� as �Ƿ���ID,���ID,��������ID" & _
                " From (" & _
                " Select Distinct 1 as ĩ��,A.ID,Decode(Instr('567',A.���),0,A.����ID,-E.����ID) as �ϼ�ID,A.����,A.����," & _
                " A.���㵥λ as ��λ,A.���,A.����,C.���� as ���,A.��������,N.���� as ҽ������,A.˵��,B.ԭ��,B.�ּ�,B.ȱʡ�۸�,A.�Ƿ���," & _
                " A.��� as ���ID,-Null as ��������ID" & _
                " From �շ���ĿĿ¼ A,�շѼ�Ŀ B,�շ���Ŀ��� C,ҩƷ��� D,������ĿĿ¼ E,����֧����Ŀ M,����֧������ N" & _
                " Where A.ID=B.�շ�ϸĿID  [ѡ���滻�Ĺ�����1] And (A.����ʱ�� is NULL Or A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                GetPriceGradeSQL(mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�, "A", "B", "4", "5", "6") & _
                " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
                " And ((Sysdate Between B.ִ������ and B.��ֹ����) or (Sysdate>=B.ִ������ And B.��ֹ���� is NULL))" & _
                " And A.������� IN(" & IIF(mlng�������� = 1, 1, 2) & ",3)" & IIF(str��ĿIDs <> "", " And Instr([1],','||A.ID||',')=0", "") & _
                " And A.��� Not IN('4','J','1') And A.���=C.���� And A.ID=D.ҩƷID(+) And D.ҩ��ID=E.ID(+)" & _
                " And A.ID=M.�շ�ϸĿID(+) And M.����ID=N.ID(+) And M.����(+)=[2]" & _
                " And (Nvl(a.ִ�п���,0) <> 4 Or Exists (Select 1 From �շ�ִ�п��� W Where w.�շ�ϸĿid = a.Id And (w.������Դ=2 or (w.������Դ is Null And Nvl(w.��������id,[3]) = [3]))))" & _
                " And (a.��� Not in ('5','6','7') Or Exists(Select 1 From �շ�ִ�п��� W Where w.�շ�ϸĿid=a.Id And Nvl(w.��������id,[3])=[3]))"
            If DeptExist("���ϲ���", 2) Then
                strSQL2 = strSQL2 & " Union ALL " & _
                    " Select Distinct 1 as ĩ��,A.ID,-E.����ID as �ϼ�ID,A.����,A.����," & _
                    " A.���㵥λ as ��λ,A.���,A.����,C.���� as ���,A.��������,N.���� as ҽ������,A.˵��," & _
                    " B.ԭ��,B.�ּ�,B.ȱʡ�۸�,A.�Ƿ���,A.��� as ���ID,D.�������� as ��������ID" & _
                    " From �շ���ĿĿ¼ A,�շѼ�Ŀ B,�շ���Ŀ��� C,�������� D,������ĿĿ¼ E,����֧����Ŀ M,����֧������ N" & _
                    " Where A.ID=B.�շ�ϸĿID [ѡ���滻�Ĺ�����2]  And (A.����ʱ�� is NULL Or A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                    GetPriceGradeSQL(mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�, "A", "B", "4", "5", "6") & _
                    " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
                    " And ((Sysdate Between B.ִ������ and B.��ֹ����) or (Sysdate>=B.ִ������ And B.��ֹ���� is NULL))" & _
                    " And A.������� IN(" & IIF(mlng�������� = 1, 1, 2) & ",3)" & IIF(str��ĿIDs <> "", " And Instr([1],','||A.ID||',')=0", "") & _
                    " And A.���='4' And A.���=C.���� And A.ID=D.����ID And D.����ID=E.ID" & _
                    " And A.ID=M.�շ�ϸĿID(+) And M.����ID=N.ID(+) And M.����(+)=[2]" & _
                    " And Exists(Select 1 From �շ�ִ�п��� W Where w.�շ�ϸĿid=a.Id And Nvl(w.��������id,[3])=[3])"
            End If
            strSQL2 = strSQL2 & " ) Group by ĩ��,ID,�ϼ�ID,���,����,����,��λ,���,����,��������,ҽ������,˵��,�Ƿ���,���ID,��������ID"
            '[ѡ���滻�Ĺ�����1],[ѡ���滻�Ĺ�����2],����������ѡ���д����
            'Ҫȷ�� "ռλ����" �����һλ���ò�����ѡ������ƴ�ӣ�Ҫ���4000���ȵ�����
            Set rsTmp = ShowSQLSelectCIS(Me, strSQL, strSQL2, 2, "�շ���Ŀ", False, "", "", False, False, False, 0, 0, 0, blnCancel, False, False, "," & str��ĿIDs & ",", mint����, mlng���˿���id, mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�, "ռλ����")
            If Not rsTmp Is Nothing Then
                '�Ǳ���ִ�е�ҽ����������������Ŀ
                If lng�к� <> 0 Then
                    If NVL(rsTmp!�Ƿ���ID, 0) = 1 And Not (InStr(",5,6,7,", rsTmp!���ID) > 0 Or rsTmp!���ID = "4" And NVL(rsTmp!��������ID, 0) = 1) Then
                        If Not Check����ִ��(Val(vsAdvice.TextMatrix(lng�к�, COL_ִ�п���ID))) Then
                            MsgBox "��ҽ���Ǳ���ִ�У�������Ա����Ŀ""" & rsTmp!���� & """���ۡ��üƼ���Ŀ��Ҫ�ֹ��Ƽۡ�", vbInformation, gstrSysName
                            .SetFocus: Exit Sub
                        End If
                    End If
                End If
                
                'ҽ��������
                If CheckItemInsure(rsTmp) Then
                    .SetFocus: Exit Sub
                End If
                
                lngҽ��ID = Val(vsAdvice.TextMatrix(lng�к�, COL_ID))
                int�������� = Val(.TextMatrix(Row, COLP_��������))
                lngԭ��ĿID = Val(.TextMatrix(Row, COLP_�շ�ϸĿID))
                Call SetItemInput(Row, rsTmp, lngҽ��ID, int��������, lngԭ��ĿID)
                If lng�к� <> 0 Then
                    Call ShowAdvicePrice(vsAdvice.Row) '���¼�����ʾ
                End If
                Call EnterNextCell(Row, Col)
            Else
                If Not blnCancel Then
                    MsgBox "û�п��õ��շ���Ŀ�����ȵ��շ���Ŀ���������ã�", vbInformation, gstrSysName
                End If
                .SetFocus
            End If
        ElseIf Col = COLP_ִ�п��� Then
            vPoint = zlControl.GetCoordPos(.hwnd, .CellLeft, .CellTop)
            If .TextMatrix(Row, COLP_�շ����) = "4" Then
                '�������õ�����
                strSQL = _
                    " Select Distinct C.ID,C.����,C.����,C.����,B.������� as ��ΧID" & _
                    " From �շ�ִ�п��� A,��������˵�� B,���ű� C" & _
                    " Where A.ִ�п���ID+0=B.����ID And B.��������='���ϲ���'" & _
                    " And B.������� IN(" & IIF(mlng�������� = 1, 1, 2) & ",3) And B.����ID=C.ID" & _
                    " And (C.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
                    " And (A.������Դ is NULL Or A.������Դ=" & IIF(mbytSendKind = EInBilling, 2, 1) & ")" & _
                    " And (A.��������ID is NULL Or A.��������ID=[2])" & _
                    " And (C.վ��='" & gstrNodeNo & "' Or C.վ�� is Null)" & _
                    " And A.�շ�ϸĿID=[1]" & _
                    " Order by B.�������,C.����"
                Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "���ϲ���", False, "", "", False, False, True, vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, _
                    Val(.TextMatrix(Row, COLP_�շ�ϸĿID)), Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_���˿���ID)))
            ElseIf InStr(",5,6,7,", .TextMatrix(Row, COLP_�շ����)) > 0 Then
                'ҩƷ
                'ҩƷ��ϵͳָ���Ĵ���ҩ������
                If Not Check�ϰల��(True) Then
                    strSQL = _
                        " Select Distinct C.ID,C.����,C.����,C.����,B.������� as ��ΧID" & _
                        " From �շ�ִ�п��� A,��������˵�� B,���ű� C" & _
                        " Where A.ִ�п���ID+0=B.����ID And B.��������=[3]" & _
                        " And B.������� IN(" & IIF(mlng�������� = 1, 1, 2) & ",3) And B.����ID=C.ID" & _
                        " And (C.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
                        " And (A.������Դ is NULL Or A.������Դ=" & IIF(mbytSendKind = EInBilling, 2, 1) & ")" & _
                        " And (A.��������ID is NULL Or A.��������ID=[2])" & _
                        " And (C.վ��='" & gstrNodeNo & "' Or C.վ�� is Null)" & _
                        " And A.�շ�ϸĿID=[1]" & _
                        " Order by B.�������,C.����"
                Else
                    strSQL = _
                        " Select Distinct C.ID,C.����,C.����,C.����,B.������� as ��ΧID" & _
                        " From �շ�ִ�п��� A,��������˵�� B,���ű� C,���Ű��� D" & _
                        " Where A.ִ�п���ID+0=B.����ID And B.��������=[3]" & _
                        " And B.������� IN(" & IIF(mlng�������� = 1, 1, 2) & ",3) And B.����ID=C.ID" & _
                        " And (C.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
                        " And D.����ID=C.ID And D.����=To_Number(To_Char(Sysdate,'D'))-1" & _
                        " And To_Char(Sysdate,'HH24:MI:SS') Between To_Char(D.��ʼʱ��,'HH24:MI:SS') and To_Char(D.��ֹʱ��,'HH24:MI:SS') " & _
                        " And (A.������Դ is NULL Or A.������Դ=" & IIF(mbytSendKind = EInBilling, 2, 1) & ")" & _
                        " And (A.��������ID is NULL Or A.��������ID=[2])" & _
                        " And (C.վ��='" & gstrNodeNo & "' Or C.վ�� is Null)" & _
                        " And A.�շ�ϸĿID=[1]" & _
                        " Order by B.�������,C.����"
                End If
                Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "ҩ��", False, "", "", False, False, True, vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, _
                    Val(.TextMatrix(Row, COLP_�շ�ϸĿID)), Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_���˿���ID)), _
                    decode(.TextMatrix(Row, COLP_�շ����), "5", "��ҩ��", "6", "��ҩ��", "7", "��ҩ��"))
            End If
            If Not rsTmp Is Nothing Then
                .TextMatrix(Row, COLP_ִ�п���ID) = rsTmp!ID
                .TextMatrix(Row, Col) = rsTmp!����
                .Cell(flexcpData, Row, Col) = .TextMatrix(Row, Col)
                
                '���¼�¼��
                lngҽ��ID = Val(vsAdvice.TextMatrix(lng�к�, COL_ID))
                int�������� = Val(.TextMatrix(Row, COLP_��������))
                lngԭ��ĿID = Val(.TextMatrix(Row, COLP_�շ�ϸĿID))
                If lngҽ��ID <> 0 And lngԭ��ĿID <> 0 Then
                    mrsPrice.Filter = "ҽ��ID=" & lngҽ��ID & " And ��������=" & int�������� & " And �շ�ϸĿID=" & lngԭ��ĿID
                    mrsPrice!ִ�п���ID = rsTmp!ID
                    mrsPrice.Update
                    Call ShowAdvicePrice(vsAdvice.Row) '���¼�����ʾ
                End If
                Call EnterNextCell(Row, Col)
            Else
                If Not blnCancel Then
                    MsgBox "û���ҵ����õĿ��ҡ�", vbInformation, gstrSysName
                End If
                .SetFocus
            End If
        End If
    End With
End Sub

Private Function CheckItemInsure(rsInput As ADODB.Recordset) As Boolean
'���ܣ��������(ѡ��)�Ƽ���Ŀ�Ƿ�ҽ������
'���أ����δ���룬������ʾѡ�񲻼������򷵻��档
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, int���� As Integer
    
    If gintҽ������ = 0 Then Exit Function
    
    On Error GoTo errH
    
    strSQL = "Select ���� From ������ҳ Where ����ID=[1] And ��ҳID=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "CheckItemInsure", mlng����ID, mlng��ҳID)
    If Not rsTmp.EOF Then int���� = NVL(rsTmp!����, 0)
    If int���� <> 0 Then
        If Not ItemExistInsure(mlng����ID, rsInput!ID, int����) Then
            If gintҽ������ = 1 Then
                If MsgBox("��Ŀ""" & rsInput!���� & """û�����ö�Ӧ�ı�����Ŀ��Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    CheckItemInsure = True
                End If
            ElseIf gintҽ������ = 2 Then
                MsgBox "��Ŀ""" & rsInput!���� & """û�����ö�Ӧ�ı�����Ŀ��", vbInformation, gstrSysName
                CheckItemInsure = True
            End If
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub vsPrice_DblClick()
    Call vsPrice_KeyPress(32)
End Sub

Private Sub vsPrice_KeyDown(KeyCode As Integer, Shift As Integer)
    With vsPrice
        If KeyCode = vbKeyF4 Then
            If CellEditable(.Row, .Col) And .Col = COLP_�Ƽ�ҽ�� Then
                Call zlcommfun.PressKey(vbKeySpace)
            End If
        ElseIf KeyCode = vbKeyDelete Then
            If .Editable And Val(.TextMatrix(.Row, COLP_�̶�)) = 0 Then
                If Val(.TextMatrix(.Row, COLP_�к�)) <> 0 And Val(.TextMatrix(.Row, COLP_�շ�ϸĿID)) <> 0 Then
                    'ҽ������д�������Ҫ����һ��(�����ǹ̶����ɶ���)
                    mrsPrice.Filter = "ҽ��ID=" & Val(vsAdvice.TextMatrix(Val(.TextMatrix(.Row, COLP_�к�)), COL_ID)) & _
                        " And ��������=" & Val(.TextMatrix(.Row, COLP_��������)) & " And ����=1"
                    If mrsPrice.RecordCount = 1 And .TextMatrix(.Row, COLP_����) <> "" Then
                        MsgBox """" & .Cell(flexcpData, .Row, COLP_�Ƽ�ҽ��) & """����Ҫ����һ�������Ƽ���Ŀ��", vbInformation, gstrSysName
                        Exit Sub
                    End If
                
                    If MsgBox("ȷʵҪɾ����ǰ�Ƽ�����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
                    mrsPrice.Filter = "ҽ��ID=" & Val(vsAdvice.TextMatrix(Val(.TextMatrix(.Row, COLP_�к�)), COL_ID)) & _
                        " And ��������=" & Val(.TextMatrix(.Row, COLP_��������)) & " And �շ�ϸĿID=" & Val(.TextMatrix(.Row, COLP_�շ�ϸĿID))
                    mrsPrice.Delete
                End If
                
                .RemoveItem .Row
                If .Rows = .FixedRows Then
                    .Rows = .FixedRows + 1
                    .Row = .FixedRows: .Col = COLP_�Ƽ�ҽ��
                End If
                
                Call ShowAdvicePrice(vsAdvice.Row) '���¼�����ʾ
            End If
        ElseIf KeyCode > 127 Then
            '���ֱ�����뺺�ֵ�����
            Call vsPrice_KeyPress(KeyCode)
        End If
    End With
End Sub

Private Sub vsPrice_KeyPress(KeyAscii As Integer)
    With vsPrice
        If KeyAscii = 13 Then
            KeyAscii = 0
            Call EnterNextCell(.Row, .Col)
        Else
            If CellEditable(.Row, .Col) And (.Col = COLP_�շ���Ŀ Or .Col = COLP_ִ�п���) Then
                If KeyAscii = Asc("*") Then
                    KeyAscii = 0
                    Call vsPrice_CellButtonClick(.Row, .Col)
                Else
                    .ComboList = "" 'ʹ��ť״̬��������״̬
                End If
            End If
        End If
    End With
End Sub

Private Sub EnterNextCell(ByVal lngRow As Long, ByVal lngCol As Long)
'���ܣ���λ���۱�����һ����������ĵ�Ԫ��
    Dim i As Long, j As Long
    
    With vsPrice
        '��ǰ��Ԫ�����δ��������,���˳�
        If CellEditable(lngRow, lngCol) Then
            If lngCol = COLP_���� And Val(.TextMatrix(lngRow, lngCol)) = 0 Then
                Exit Sub
            ElseIf .TextMatrix(lngRow, lngCol) = "" Then
                Exit Sub
            End If
        End If
        
        '����һ��Ԫ��ʼѭ������
        For i = lngRow To .Rows - 1
            For j = IIF(i = lngRow, lngCol + 1, COLP_�Ƽ�ҽ��) To .Cols - 1
                If CellEditable(i, j) Then Exit For
            Next
            If j <= .Cols - 1 Then Exit For
        Next
        If i <= .Rows - 1 Then
            .Row = i: .Col = j
        Else
            '��ǰ�����û���ҵ���һ���ɱ༭��Ԫ,�������Ƽ�ҽ��,������һ����
            If CStr(.ColData(COLP_�Ƽ�ҽ��)) <> "" Then
                '��ǰ��δ��������,��λ����������Ԫ
                If .TextMatrix(lngRow, COLP_�Ƽ�ҽ��) = "" Then
                    .Col = COLP_�Ƽ�ҽ��
                ElseIf .TextMatrix(lngRow, COLP_�Ƽ�����) = "" Then
                    .Col = COLP_�Ƽ�����
                ElseIf .TextMatrix(lngRow, COLP_�շ���Ŀ) = "" Then
                    .Col = COLP_�շ���Ŀ
                ElseIf Val(.TextMatrix(lngRow, COLP_���)) = 1 _
                    And Val(.TextMatrix(lngRow, COLP_����)) = 0 _
                    And CellEditable(lngRow, COLP_����) Then
                    .Col = COLP_����
                Else
                    .AddItem "", .Rows
                    .Row = .Rows - 1: .Col = COLP_�Ƽ�ҽ��
                    
                    'ȱʡѡ��Ƽ�ҽ��(�������)
                    Call ShowDefaultRow
                End If
            Else
                If .Col + 1 <= .Cols - 1 Then .Col = .Col + 1 '���ɱ༭ʱ���ⶨһ��
            End If
        End If
        .ShowCell .Row, .Col
    End With
End Sub

Private Sub vsPrice_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, lng�к� As Long, i As Long
    Dim str��ĿIDs As String, int�������� As Integer
    Dim lngҽ��ID As Long, lngԭ��ĿID As Long
    Dim strTmp As String, blnCancel As Boolean
    Dim strInput As String, strMatch As String
    Dim vPoint As PointAPI
    Dim strStock As String
    
    With vsPrice
        If KeyAscii = 13 Then
            KeyAscii = 0
            lng�к� = Val(.TextMatrix(Row, COLP_�к�))
            If Col = COLP_�Ƽ�ҽ�� Then
                '����ʱ�س�
                If .ComboIndex <> -1 Then
                    .TextMatrix(.Row, .Col) = .ComboItem(.ComboIndex) '��ȻEnterNextCell����Ҫ�˳�
                    Call EnterNextCell(Row, Col)
                End If
            ElseIf Col = COLP_�Ƽ����� Then
                If Not IsNumeric(.EditText) Or Val(.EditText) <= 0 Or Val(.EditText) > LONG_MAX Then
                    MsgBox "�Ƽ�����������󣬲��Ǵ���������ֻ�������ֵ����", vbInformation, gstrSysName
                    .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col))
                    Exit Sub
                End If
                .TextMatrix(Row, Col) = .EditText
                .Cell(flexcpData, Row, Col) = .TextMatrix(Row, Col)
                
                '���¼�¼��
                lngҽ��ID = Val(vsAdvice.TextMatrix(lng�к�, COL_ID))
                int�������� = Val(.TextMatrix(Row, COLP_��������))
                lngԭ��ĿID = Val(.TextMatrix(Row, COLP_�շ�ϸĿID))
                If lngҽ��ID <> 0 And lngԭ��ĿID <> 0 Then
                    mrsPrice.Filter = "ҽ��ID=" & lngҽ��ID & " And ��������=" & int�������� & " And �շ�ϸĿID=" & lngԭ��ĿID
                    mrsPrice!���� = Val(.TextMatrix(Row, Col))
                    mrsPrice.Update
                    Call ShowAdvicePrice(vsAdvice.Row) '���¼�����ʾ
                End If
                Call EnterNextCell(Row, Col)
            ElseIf Col = COLP_���� Then
                If Not IsNumeric(.EditText) Or Val(.EditText) <= 0 Or Val(.EditText) > LONG_MAX Then
                    MsgBox "�շѵ���������󣬲��Ǵ���������ֻ�������ֵ����", vbInformation, gstrSysName
                    .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col))
                    Exit Sub
                End If
                '��������뷶Χ
                strTmp = CheckScope(.Cell(flexcpData, Row, COLP_Ӧ�ս��), .Cell(flexcpData, Row, COLP_ʵ�ս��), .EditText)
                If strTmp <> "" Then
                    MsgBox strTmp, vbInformation, gstrSysName
                    .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col))
                    Exit Sub
                End If
                .EditText = Format(.EditText, gstrDecPrice)
                .TextMatrix(Row, Col) = .EditText
                .Cell(flexcpData, Row, Col) = .TextMatrix(Row, Col)
                
                '���¼�¼��
                lngҽ��ID = Val(vsAdvice.TextMatrix(lng�к�, COL_ID))
                int�������� = Val(.TextMatrix(Row, COLP_��������))
                lngԭ��ĿID = Val(.TextMatrix(Row, COLP_�շ�ϸĿID))
                If lngҽ��ID <> 0 And lngԭ��ĿID <> 0 Then
                    mrsPrice.Filter = "ҽ��ID=" & lngҽ��ID & " And ��������=" & int�������� & " And �շ�ϸĿID=" & lngԭ��ĿID
                    mrsPrice!���� = Val(.TextMatrix(Row, Col))
                    mrsPrice.Update
                    Call ShowAdvicePrice(vsAdvice.Row) '���¼�����ʾ
                End If
                Call EnterNextCell(Row, Col)
            ElseIf Col = COLP_�շ���Ŀ And .EditText <> "" Then
                '����ѡ�����е���Ŀ
                For i = .FixedRows To .Rows - 1
                    If Val(vsAdvice.TextMatrix(Val(.TextMatrix(i, COLP_�к�)), COL_ID)) = Val(vsAdvice.TextMatrix(lng�к�, COL_ID)) _
                        And Val(vsAdvice.TextMatrix(lng�к�, COL_ID)) <> 0 And i <> Row Then
                        str��ĿIDs = str��ĿIDs & "," & Val(.TextMatrix(i, COLP_�շ�ϸĿID))
                    End If
                Next
                str��ĿIDs = Mid(str��ĿIDs, 2)
                
                If mlng��ҩ�� <> 0 Or mlng��ҩ�� <> 0 Or mlng��ҩ�� <> 0 Or mlng���ϲ��� <> 0 Then
                    strStock = _
                        "Select A.ҩƷID,Sum(Nvl(A.��������,0)) as ���" & _
                        " From ҩƷ��� A,�շ���ĿĿ¼ B" & _
                        " Where A.���� = 1 And (Nvl(A.����,0)=0 Or A.Ч�� Is Null Or A.Ч��>Trunc(Sysdate))" & _
                        " And A.�ⷿID=Decode(B.���,'5',[7],'6',[8],'7',[9],'4',[10],Null)" & _
                        " And A.ҩƷID=B.ID And B.��� IN('4','5','6','7')" & _
                        " Group by A.ҩƷID Having Sum(Nvl(A.��������,0))<>0"
                Else
                    strStock = "Select Null as ҩƷID,Null as ��� From Dual"
                End If
                
                '��ͬ������ƥ�䷽ʽ
                strInput = UCase(.EditText)
                strMatch = " And (A.���� Like [1] And C.����=[3] Or C.���� Like [2] And C.����=[3] Or C.���� Like [2] And C.���� IN([3],3))"
                If IsNumeric(strInput) Then                         '10,11.����ȫ������ʱֻƥ�����'����ҩƷ,��Ҫƥ�����(����Ϊ3��������)
                    If Mid(gstrMatchMode, 1, 1) = "1" Then strMatch = " And (A.���� Like [1] And C.����=[3] Or C.���� Like [2] And C.����=3)"
                ElseIf zlcommfun.IsCharAlpha(strInput) Then         '01,11.����ȫ����ĸʱֻƥ�����
                    If Mid(gstrMatchMode, 2, 1) = "1" Then strMatch = " And C.���� Like [2] And C.����=[3]"
                ElseIf zlcommfun.IsCharChinese(strInput) Then
                    strMatch = " And C.���� Like [2] And C.����=[3]"
                End If
                
                strSQL = ""
                If Not DeptExist("���ϲ���", 2) Then strSQL = " And A.���<>'4'"
                strSQL = _
                    " Select A.ĩ��,A.ID,A.���,A.����,A.����,A.��λ,A.���,A.����," & _
                    " Decode(Instr('4567',A.���ID),0,NULL,1," & _
                    "   Decode(S.���,NULL,NULL,LTrim(To_Char(S.���,'999990.0000'))||A.��λ)," & _
                    "   Decode(S.���,NULL,NULL,LTrim(To_Char(S.���/Nvl(C.סԺ��װ,1),'999990.0000'))||C.סԺ��λ)) as ���," & _
                    "   A.��������,N.���� as ҽ������,A.˵��," & _
                    " Decode(Nvl(A.�Ƿ���,0),1,Decode(Instr('567',A.���ID),0,Sum(Nvl(A.ԭ��,0))||'-'||Sum(Nvl(A.�ּ�,0)),'ʱ��'),Sum(A.�ּ�)) as �۸�," & _
                    " Sum(A.ԭ��) as ԭ��ID,Sum(A.�ּ�) as �ּ�ID,Sum(A.ȱʡ�۸�) as ȱʡ�۸�ID,A.�Ƿ��� as �Ƿ���ID,A.���ID,B.�������� as ��������ID" & _
                    " From (" & _
                    " Select Distinct 1 as ĩ��,A.ID,a.ִ�п���,A.��� as ���ID,D.���� as ���,A.����,A.����,A.���㵥λ as ��λ," & _
                    " A.���,A.����,A.��������,A.˵��,B.ԭ��,B.�ּ�,B.ȱʡ�۸�,A.�Ƿ���" & _
                    " From �շ���ĿĿ¼ A,�շѼ�Ŀ B,�շ���Ŀ���� C,�շ���Ŀ��� D" & _
                    " Where A.ID=B.�շ�ϸĿID And (A.����ʱ�� is NULL Or A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                    GetPriceGradeSQL(mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�, "A", "B", "11", "12", "13") & _
                    " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
                    " And ((Sysdate Between B.ִ������ and B.��ֹ����) or (Sysdate>=B.ִ������ And B.��ֹ���� is NULL))" & _
                    " And A.������� IN(" & IIF(mlng�������� = 1, 1, 2) & ",3)" & IIF(str��ĿIDs <> "", " And Instr([4],','||A.ID||',')=0", "") & _
                    " And A.ID=C.�շ�ϸĿID And A.���=D.���� And A.��� Not IN('J','1')" & strSQL & strMatch & _
                    " ) A,�������� B,ҩƷ��� C,����֧����Ŀ M,����֧������ N,(" & strStock & ") S" & _
                    " Where A.ID=B.����ID(+) And A.ID=M.�շ�ϸĿID(+) And M.����ID=N.ID(+) And M.����(+)=[5]  And A.ID=C.ҩƷID(+) And A.ID=S.ҩƷID(+)" & _
                    " And (Nvl(a.ִ�п���,0) <> 4 Or Exists (Select 1 From �շ�ִ�п��� W Where w.�շ�ϸĿid = a.Id And (w.������Դ=2 or (w.������Դ is Null And Nvl(w.��������id,[6]) = [6]))))" & _
                    " And (a.���id not in ('4','5','6','7') Or Exists(Select 1 From �շ�ִ�п��� W Where w.�շ�ϸĿid=a.Id And Nvl(w.��������id,[6])=[6]))" & _
                    " Group by A.ĩ��,A.ID,A.���,A.����,A.����,A.��λ,A.���,A.����,A.��������,C.סԺ��λ,C.סԺ��װ,S.���,N.����,A.˵��,A.�Ƿ���,A.���ID,B.��������" & _
                    " Order by A.���,A.����"
                vPoint = zlControl.GetCoordPos(.hwnd, .CellLeft, .CellTop)
                Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "�շ���Ŀ", False, "", "", False, False, True, vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, _
                    strInput & "%", mstrLike & strInput & "%", mint���� + 1, "," & str��ĿIDs & ",", mint����, mlng���˿���id, mlng��ҩ��, mlng��ҩ��, mlng��ҩ��, mlng���ϲ���, mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�)
                If Not rsTmp Is Nothing Then
                    '�Ǳ���ִ�е�ҽ����������������Ŀ
                    If lng�к� <> 0 Then
                        If NVL(rsTmp!�Ƿ���ID, 0) = 1 And Not (InStr(",5,6,7,", rsTmp!���ID) > 0 Or rsTmp!���ID = "4" And NVL(rsTmp!��������ID, 0) = 1) Then
                            If Not Check����ִ��(Val(vsAdvice.TextMatrix(lng�к�, COL_ִ�п���ID))) Then
                                MsgBox "��ҽ���Ǳ���ִ�У�������Ա����Ŀ""" & rsTmp!���� & """���ۡ��üƼ���Ŀ��Ҫ�ֹ��Ƽۡ�", vbInformation, gstrSysName
                                .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col))
                                Call vsPrice_AfterRowColChange(Row, Col, Row, Col) '������ʾ��ť
                                .SetFocus: Exit Sub
                            End If
                        End If
                    End If
                
                    'ҽ��������
                    If CheckItemInsure(rsTmp) Then
                        .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col))
                        Call vsPrice_AfterRowColChange(Row, Col, Row, Col) '������ʾ��ť
                        .SetFocus: Exit Sub
                    End If
                
                    lngҽ��ID = Val(vsAdvice.TextMatrix(lng�к�, COL_ID))
                    int�������� = Val(.TextMatrix(Row, COLP_��������))
                    lngԭ��ĿID = Val(.TextMatrix(Row, COLP_�շ�ϸĿID))
                    Call SetItemInput(Row, rsTmp, lngҽ��ID, int��������, lngԭ��ĿID)
                    .EditText = .TextMatrix(Row, Col) 'ֱ������ƥ����Ҫ
                    If lng�к� <> 0 Then
                        Call ShowAdvicePrice(vsAdvice.Row) '���¼�����ʾ
                    End If
                    Call EnterNextCell(Row, Col)
                Else
                    If Not blnCancel Then
                        MsgBox "û���ҵ����õ��շ���Ŀ��", vbInformation, gstrSysName
                    End If
                    .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col))
                    Call vsPrice_AfterRowColChange(Row, Col, Row, Col) '������ʾ��ť
                    .SetFocus
                End If
            ElseIf Col = COLP_ִ�п��� And .EditText <> "" Then 'ִ�п���
                vPoint = zlControl.GetCoordPos(.hwnd, .CellLeft, .CellTop)
                If .TextMatrix(Row, COLP_�շ����) = "4" Then
                    '�������õ�����
                    strSQL = _
                        " Select Distinct C.ID,C.����,C.����,C.����,B.������� as ��ΧID" & _
                        " From �շ�ִ�п��� A,��������˵�� B,���ű� C" & _
                        " Where A.ִ�п���ID+0=B.����ID And B.��������='���ϲ���'" & _
                        " And B.������� IN(" & IIF(mlng�������� = 1, 1, 2) & ",3) And B.����ID=C.ID" & _
                        " And (C.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
                        " And (A.������Դ is NULL Or A.������Դ=" & IIF(mbytSendKind = EInBilling, 2, 1) & ")" & _
                        " And (A.��������ID is NULL Or A.��������ID=[2])" & _
                        " And (C.վ��='" & gstrNodeNo & "' Or C.վ�� is Null)" & _
                        " And A.�շ�ϸĿID=[1] And (C.���� Like [3] Or C.���� Like [4] Or C.���� Like [4])" & _
                        " Order by B.�������,C.����"
                    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "���ϲ���", False, "", "", False, False, True, vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, _
                        Val(.TextMatrix(Row, COLP_�շ�ϸĿID)), Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_���˿���ID)), UCase(.EditText) & "%", mstrLike & UCase(.EditText) & "%")
                ElseIf InStr(",5,6,7,", .TextMatrix(Row, COLP_�շ����)) > 0 Then
                    'ҩƷ��ϵͳָ���Ĵ���ҩ������
                    If Not Check�ϰల��(True) Then
                        strSQL = _
                            " Select Distinct C.ID,C.����,C.����,C.����,B.������� as ��ΧID" & _
                            " From �շ�ִ�п��� A,��������˵�� B,���ű� C" & _
                            " Where A.ִ�п���ID+0=B.����ID And B.��������=[3]" & _
                            " And B.������� IN(" & IIF(mlng�������� = 1, 1, 2) & ",3) And B.����ID=C.ID" & _
                            " And (C.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
                            " And (A.������Դ is NULL Or A.������Դ=" & IIF(mbytSendKind = EInBilling, 2, 1) & ")" & _
                            " And (A.��������ID is NULL Or A.��������ID=[2])" & _
                            " And (C.վ��='" & gstrNodeNo & "' Or C.վ�� is Null)" & _
                            " And A.�շ�ϸĿID=[1] And (C.���� Like [4] Or C.���� Like [5] Or C.���� Like [5])" & _
                            " Order by B.�������,C.����"
                    Else
                        strSQL = _
                            " Select Distinct C.ID,C.����,C.����,C.����,B.������� as ��ΧID" & _
                            " From �շ�ִ�п��� A,��������˵�� B,���ű� C,���Ű��� D" & _
                            " Where A.ִ�п���ID+0=B.����ID And B.��������=[3]" & _
                            " And B.������� IN(" & IIF(mlng�������� = 1, 1, 2) & ",3) And B.����ID=C.ID" & _
                            " And (C.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
                            " And D.����ID=C.ID And D.����=To_Number(To_Char(Sysdate,'D'))-1" & _
                            " And To_Char(Sysdate,'HH24:MI:SS') Between To_Char(D.��ʼʱ��,'HH24:MI:SS') and To_Char(D.��ֹʱ��,'HH24:MI:SS') " & _
                            " And (A.������Դ is NULL Or A.������Դ=" & IIF(mbytSendKind = EInBilling, 2, 1) & ")" & _
                            " And (A.��������ID is NULL Or A.��������ID=[2])" & _
                            " And (C.վ��='" & gstrNodeNo & "' Or C.վ�� is Null)" & _
                            " And A.�շ�ϸĿID=[1] And (C.���� Like [4] Or C.���� Like [5] Or C.���� Like [5])" & _
                            " Order by B.�������,C.����"
                    End If
                    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "ҩ��", False, "", "", False, False, True, vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, _
                        Val(.TextMatrix(Row, COLP_�շ�ϸĿID)), Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_���˿���ID)), _
                        decode(.TextMatrix(Row, COLP_�շ����), "5", "��ҩ��", "6", "��ҩ��", "7", "��ҩ��"), _
                        UCase(.EditText) & "%", mstrLike & UCase(.EditText) & "%")
                End If
                If Not rsTmp Is Nothing Then
                    .TextMatrix(Row, COLP_ִ�п���ID) = rsTmp!ID
                    .TextMatrix(Row, Col) = rsTmp!����
                    .Cell(flexcpData, Row, Col) = .TextMatrix(Row, Col)
                    .EditText = .TextMatrix(Row, Col) 'ֱ������ƥ����Ҫ
                    
                    '���¼�¼��
                    lngҽ��ID = Val(vsAdvice.TextMatrix(lng�к�, COL_ID))
                    int�������� = Val(.TextMatrix(Row, COLP_��������))
                    lngԭ��ĿID = Val(.TextMatrix(Row, COLP_�շ�ϸĿID))
                    If lngҽ��ID <> 0 And lngԭ��ĿID <> 0 Then
                        mrsPrice.Filter = "ҽ��ID=" & lngҽ��ID & " And ��������=" & int�������� & " And �շ�ϸĿID=" & lngԭ��ĿID
                        mrsPrice!ִ�п���ID = rsTmp!ID
                        mrsPrice.Update
                        Call ShowAdvicePrice(vsAdvice.Row) '���¼�����ʾ
                    End If
                    Call EnterNextCell(Row, Col)
                Else
                    If Not blnCancel Then
                        MsgBox "û���ҵ����õĿ��ҡ�", vbInformation, gstrSysName
                    End If
                    .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col))
                    Call vsPrice_AfterRowColChange(Row, Col, Row, Col) '������ʾ��ť
                    .SetFocus
                End If
            End If
        Else
            If Col = COLP_�Ƽ����� Or Col = COLP_���� Then
                If InStr("0123456789." & Chr(8) & Chr(27), Chr(KeyAscii)) = 0 Then
                    KeyAscii = 0: Exit Sub
                End If
            End If
        End If
    End With
End Sub

Private Sub SetItemInput(lngRow As Long, rsInput As ADODB.Recordset, ByVal lngҽ��ID As Long, ByVal int�������� As Integer, ByVal lngԭ��ĿID As Long)
    Dim lngִ�п���ID As Long, lng���˿���ID As Long
    Dim lng�к� As Long, dbl���� As Double
    Dim blnHaveSub As Boolean
    
    With vsPrice
        '��¼������
        '�������:����ʱ��ʾ�����������Ŀ,Ҳ���Դ���Ϊδ���Ƽ�ҽ��������������Ŀ
        .TextMatrix(lngRow, COLP_���) = rsInput!���
        .TextMatrix(lngRow, COLP_�շ����) = rsInput!���ID
        .TextMatrix(lngRow, COLP_�շ�ϸĿID) = rsInput!ID
        .TextMatrix(lngRow, COLP_�շ���Ŀ) = rsInput!����
        If Not IsNull(rsInput!����) Then
            .TextMatrix(lngRow, COLP_�շ���Ŀ) = .TextMatrix(lngRow, COLP_�շ���Ŀ) & "(" & rsInput!���� & ")"
        End If
        If Not IsNull(rsInput!���) Then
            .TextMatrix(lngRow, COLP_�շ���Ŀ) = .TextMatrix(lngRow, COLP_�շ���Ŀ) & " " & rsInput!���
        End If
        .TextMatrix(lngRow, COLP_��λ) = NVL(rsInput!��λ) '�������۵�λ(������ҩ��ҩƷ�Ƽ�)
        .TextMatrix(lngRow, COLP_�Ƽ�����) = 1 'ȱʡ��ԼƼ�1,ҩƷΪ��1�����۵�λ
        
        'ִ�п���
        lng�к� = Val(.TextMatrix(lngRow, COLP_�к�))
        If lng�к� <> 0 Then
            lngִ�п���ID = Val(vsAdvice.TextMatrix(lng�к�, COL_ִ�п���ID))
            '��ҩ��ҩƷ�͸������õ�����ר����ִ�п���
            If rsInput!���ID = "4" And NVL(rsInput!��������ID, 0) = 1 Or InStr(",5,6,7,", rsInput!���ID) > 0 Then
                lng���˿���ID = Val(vsAdvice.TextMatrix(lng�к�, COL_���˿���ID))
                lngִ�п���ID = Get�շ�ִ�п���ID(mlng����ID, 0, rsInput!���ID, rsInput!ID, 4, lng���˿���ID, 0, 2, lngִ�п���ID, , , 2)
            End If
        End If
        .TextMatrix(lngRow, COLP_ִ�п���) = sys.RowValue("���ű�", lngִ�п���ID, "����")
        .TextMatrix(lngRow, COLP_ִ�п���ID) = lngִ�п���ID
        
        '���ۼ��㴦��:ҩ����ҩƷ�Ƽ۲����������ﴦ��
        If InStr(",5,6,7,", rsInput!���ID) > 0 Then
            If NVL(rsInput!�Ƿ���ID, 0) = 0 Then
                dbl���� = NVL(rsInput!�ּ�ID, 0)
            ElseIf lng�к� <> 0 Then
                '��ÿ��ȱʡһ�����۵�λ,��ǰ�������μ���
                dbl���� = CalcDrugPrice(rsInput!ID, lngִ�п���ID, Val(vsAdvice.TextMatrix(lng�к�, COL_����)), , True, , mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�)
            End If
            .TextMatrix(lngRow, COLP_����) = Format(dbl����, gstrDecPrice)
                        
            'ʱ��ҩƷ������۸�
            .TextMatrix(lngRow, COLP_���) = 0
            .Cell(flexcpData, lngRow, COLP_Ӧ�ս��) = 0
            .Cell(flexcpData, lngRow, COLP_ʵ�ս��) = 0
        ElseIf rsInput!���ID = "4" And NVL(rsInput!��������ID, 0) = 1 And NVL(rsInput!�Ƿ���ID, 0) = 1 Then
            '�������õ�ʱ�����ĺ�ҩƷһ������
            dbl���� = 0
            If lng�к� <> 0 Then
                dbl���� = CalcDrugPrice(rsInput!ID, lngִ�п���ID, Val(vsAdvice.TextMatrix(lng�к�, COL_����)), , True, , mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�)
            End If
            .TextMatrix(lngRow, COLP_���) = 0
            .TextMatrix(lngRow, COLP_����) = Format(dbl����, gstrDecPrice)
            .Cell(flexcpData, lngRow, COLP_Ӧ�ս��) = 0
            .Cell(flexcpData, lngRow, COLP_ʵ�ս��) = 0
        Else
            If NVL(rsInput!�Ƿ���ID, 0) = 0 Then
                .TextMatrix(lngRow, COLP_���) = 0
                .TextMatrix(lngRow, COLP_����) = Format(NVL(rsInput!�ּ�ID, 0), gstrDecPrice)
                .Cell(flexcpData, lngRow, COLP_Ӧ�ս��) = 0
                .Cell(flexcpData, lngRow, COLP_ʵ�ս��) = 0
            Else
                .TextMatrix(lngRow, COLP_���) = 1
                .TextMatrix(lngRow, COLP_����) = Format(NVL(rsInput!ȱʡ�۸�ID), gstrDecPrice)
                .Cell(flexcpData, lngRow, COLP_Ӧ�ս��) = NVL(rsInput!ԭ��ID, 0)
                .Cell(flexcpData, lngRow, COLP_ʵ�ս��) = NVL(rsInput!�ּ�ID, 0)
            End If
        End If
        
        .TextMatrix(lngRow, COLP_��������) = NVL(rsInput!��������)
        .TextMatrix(lngRow, COLP_�̶�) = 0
        
        '��������ָ�
        .Cell(flexcpData, lngRow, COLP_�շ���Ŀ) = .TextMatrix(lngRow, COLP_�շ���Ŀ)
        .Cell(flexcpData, lngRow, COLP_�Ƽ�����) = .TextMatrix(lngRow, COLP_�Ƽ�����)
        .Cell(flexcpData, lngRow, COLP_����) = .TextMatrix(lngRow, COLP_����)
        .Cell(flexcpData, lngRow, COLP_ִ�п���) = .TextMatrix(lngRow, COLP_ִ�п���)
        
        '��¼������
        If lngҽ��ID <> 0 Then
            If lngԭ��ĿID = 0 Then
                '��ǰҽ���Ƿ��д��������������Ŀ�Ƿ����
                mrsPrice.Filter = "ҽ��ID=" & lngҽ��ID & " And ��������=" & int�������� & " And ����=1"
                If Not mrsPrice.EOF Then blnHaveSub = True
                .TextMatrix(lngRow, COLP_����) = IIF(blnHaveSub, "��", "")
            
                mrsPrice.AddNew '����
            Else '����
                mrsPrice.Filter = "ҽ��ID=" & lngҽ��ID & " And ��������=" & int�������� & " And �շ�ϸĿID=" & lngԭ��ĿID
            End If
            If lngԭ��ĿID = 0 Then
                mrsPrice!ҽ��ID = lngҽ��ID
                lng�к� = Val(.TextMatrix(lngRow, COLP_�к�))
                If Val(vsAdvice.TextMatrix(lng�к�, COL_���ID)) <> 0 Then
                    mrsPrice!���ID = Val(vsAdvice.TextMatrix(lng�к�, COL_���ID))
                Else
                    mrsPrice!���ID = Null
                End If
                mrsPrice!�������� = int��������
                mrsPrice!���� = IIF(blnHaveSub, 1, 0)
            End If
            mrsPrice!�շѷ�ʽ = 0
            mrsPrice!�շ���� = rsInput!���ID
            mrsPrice!�շ�ϸĿID = rsInput!ID
            If lngִ�п���ID <> 0 Then
                mrsPrice!ִ�п���ID = lngִ�п���ID
            Else
                mrsPrice!ִ�п���ID = Null
            End If
            mrsPrice!���� = NVL(rsInput!��������ID, 0)
            mrsPrice!��� = NVL(rsInput!�Ƿ���ID, 0)
            mrsPrice!���� = Val(.TextMatrix(lngRow, COLP_����))
            mrsPrice!���� = 1
            mrsPrice!�̶� = 0
            mrsPrice.Update
        End If
    End With
End Sub

Private Sub vsPrice_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    vsPrice.EditSelStart = 0
    vsPrice.EditSelLength = zlcommfun.ActualLen(vsPrice.EditText)
End Sub

Private Sub vsPrice_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim bln�Ǳ��� As Boolean
    
    If Not CellEditable(Row, Col, bln�Ǳ���) Then
        '�Ǳ���ִ�еı����Ŀ�������۸�
        If bln�Ǳ��� Then
            MsgBox "��ҽ���Ǳ���ִ�У�������Ա����Ŀ���ۡ��üƼ���Ŀ��Ҫ�ֹ��Ƽۡ�", vbInformation, gstrSysName
        End If
        Cancel = True
    Else
        If Col = COLP_�Ƽ����� Or Col = COLP_���� Or Col = COLP_ִ�п��� Then
            '������ȷ���շ���Ŀ
            If vsPrice.TextMatrix(Row, COLP_�շ���Ŀ) = "" Then Cancel = True
        End If
        If Col = COLP_���� Then
            '������ǰ������ȷ���Ƽ�ҽ��,�Ծ����Ƿ��������(����ִ��)
            If vsPrice.TextMatrix(Row, COLP_�Ƽ�ҽ��) = "" Then Cancel = True
        End If
    End If
    
    If Col = COLP_�Ƽ����� Or Col = COLP_���� Then
        vsPrice.EditMaxLength = 10
    Else
        vsPrice.EditMaxLength = 0
    End If
End Sub

Private Sub InitBillSet()
'���ܣ���ʼ��ҽ�����ʵ������ɼ�¼��
    Set mrsBill = New ADODB.Recordset
    
    mrsBill.Fields.Append "Key", adVarChar, 100
    mrsBill.Fields.Append "NO", adVarChar, 30
    mrsBill.Fields.Append "�������", adBigInt
    mrsBill.Fields.Append "�������", adBigInt
    mrsBill.CursorLocation = adUseClient
    mrsBill.LockType = adLockOptimistic
    mrsBill.CursorType = adOpenStatic
    mrsBill.Open
        
    Set mrsRXKey = New ADODB.Recordset
    mrsRXKey.Fields.Append "Key", adVarChar, 200
    mrsRXKey.Fields.Append "ҽ��ID", adVarChar, 200
    mrsRXKey.Fields.Append "����", adBigInt
    mrsRXKey.Fields.Append "����", adBigInt
    mrsRXKey.CursorLocation = adUseClient
    mrsRXKey.LockType = adLockOptimistic
    mrsRXKey.CursorType = adOpenStatic
    mrsRXKey.Open
End Sub

Private Sub GetCurBillSet(ByVal strKey As String, strNO As String, lng������� As Long, lng������� As Long, bln�շѵ� As Boolean)
'���ܣ���ȡ��ǰ���õ��ݵ�NO�����
'������lng�������=���ü�¼�е����,Ϊ-1ʱ��ʾ��ȡ�������
'      lng�������=���ͼ�¼�е����,Ϊ-1ʱ��ʾ��ȡ�������
'˵����strKey=���ݼ��ʵ������ɹ��򶨵�Ψһ�ؼ���
'1.������ҩ��"����(����ID,�Һŵ�)_���˿���ID_��������ID_����ҽ��_ִ�п���ID"�ֺš�
'2.һ���䷽�е����в�ҩ����һ���������ݺ�
'3.����ҽ�����ҩ�ֺŹ�����ͬ��
'4.������ҩҽ��ÿ��ҽ��һ���������ݺ�(������ҩ;�����䷽�巨���÷�)
'5.��鲿λ�͸�����������Ҫҽ��������ͬ���ݺţ�����������䵥���ĵ��ݺš�
'6.һ���ɼ��ļ�����Ϸ�����ͬ�ĵ��ݺţ��걾�ɼ��������䵥���ĵ��ݺ�
    mrsBill.Filter = "Key='" & strKey & "'"
    If mrsBill.EOF Then
        mrsBill.AddNew
        mrsBill!Key = strKey
        
        'ȡ���ݺ�
        'mrsBill!NO = zlDatabase.GetNextNo(IIF(bln�շѵ�, 13, 14)),������ʵ�Ҳ��14
        mlngNOSequence = mlngNOSequence + 1
        mrsBill!NO = "TemporaryNO=" & IIF(bln�շѵ�, 13, 14) & Format(mlngNOSequence, "00000")
        
        mrsBill!������� = IIF(lng������� = -1, 0, 1)
        mrsBill!������� = IIF(lng������� = -1, 0, 1)
        mrsBill.Update
    Else
        If lng������� <> -1 Then
            mrsBill!������� = mrsBill!������� + 1
        End If
        If lng������� <> -1 Then
            mrsBill!������� = mrsBill!������� + 1
        End If
        mrsBill.Update
    End If
    strNO = mrsBill!NO
    If lng������� <> -1 Then lng������� = mrsBill!�������
    If lng������� <> -1 Then lng������� = mrsBill!�������
End Sub

Private Sub ReplaceTrueNO(rsSQL As ADODB.Recordset, rsUpload As ADODB.Recordset)
'���ܣ�����ʱ������NO�滻�����ձ������ʵNO
    Dim strNO As String, strCur As String, strPre As String
    
    rsSQL.Filter = 0
    rsSQL.Sort = "NO"
    Do While Not rsSQL.EOF
        If Not IsNull(rsSQL!NO) Then
            strCur = Split(rsSQL!NO, "=")(1)
            If strCur <> strPre Then
                strPre = strCur
                strNO = zlDatabase.GetNextNo(Val(Left(strCur, 2)))
                            
                'rsUpload��һ��NOֻ��һ����¼
                rsUpload.Filter = "NO='" & rsSQL!NO & "'"
                If Not rsUpload.EOF Then
                    rsUpload!NO = strNO
                    rsUpload.Update
                End If
            End If
            
            rsSQL!Sql = Replace(rsSQL!Sql, rsSQL!NO, strNO)
            'rsSQL!NO = strNO '��������£����⵼��Sort��˳������
            rsSQL.Update
        End If
        rsSQL.MoveNext
    Loop
End Sub

Private Sub DeleteSendRow()
'���ܣ���������ҽ���嵥���ѷ��ͳɹ��ĵ���ɾ��
    Dim i As Long, blnDel As Boolean
    
    With vsAdvice
        .Redraw = flexRDNone
        For i = .Rows - 1 To .FixedRows Step -1
            If .RowData(i) = -1 Then .RemoveItem i: blnDel = True
        Next
        .Redraw = flexRDDirect
        
        If blnDel Then
            If .Rows = .FixedRows Then .Rows = .FixedRows + 1
            For i = .FixedRows To .Rows - 1
                If Not .RowHidden(i) Then
                    .Row = i: .Col = COL_ѡ��
                    Call .ShowCell(.Row, .Col)
                    Exit For
                End If
            Next
            
            vsPrice.Rows = vsPrice.FixedRows
            vsPrice.Rows = vsPrice.FixedRows + 1
            Call vsAdvice_AfterRowColChange(-1, -1, .Row, .Col)
        End If
    End With
End Sub

Private Function Getʵ�ս��(ByVal strSQL As String) As Currency
    Dim lngPos As Long, strMatch As String
    
    strMatch = Chr(0) & Chr(1) & "Begin"
    strSQL = Mid(strSQL, InStr(strSQL, strMatch) + Len(strMatch))
    strMatch = "End" & Chr(0) & Chr(1)
    strSQL = Left(strSQL, InStr(strSQL, strMatch) - 1)
    Getʵ�ս�� = CCur(strSQL)
End Function

Private Function Setʵ�ս��(ByVal strSQL As String, ByVal cur��� As Currency) As String
    Dim strLeft As String, strRight As String
    Dim strMatch As String, strVal As String
    
    strMatch = Chr(0) & Chr(1) & "Begin"
    strLeft = Mid(strSQL, 1, InStr(strSQL, strMatch) - 1)
    strMatch = "End" & Chr(0) & Chr(1)
    strRight = Mid(strSQL, InStr(strSQL, strMatch) + Len(strMatch))
    
    Setʵ�ս�� = strLeft & cur��� & strRight
End Function

Private Function CheckSignSend() As Boolean
'���ܣ����һ��ǩ����ҽ���Ƿ�һ���͵�
'˵��������ֻ����¿���ҽ������У�Ե�ҽ�����Ͳ�����Ӱ��(��ͬ������û��У��)
    Dim colǩ��ID As New Collection, strǩ��ID As String
    Dim lngǩ��id As Long, strTmp As String
    Dim i As Long, j As Long
    
    With vsAdvice
        For i = .FixedRows To .Rows - 1
            If Val(.TextMatrix(i, COL_ҽ��״̬)) = 1 Then
                '�ռ���ǩ��ҽ���ķ���״̬
                lngǩ��id = Val(.TextMatrix(i, COL_ǩ��ID))
                If lngǩ��id <> 0 Then
                    If InStr(strǩ��ID & ",", "," & lngǩ��id & ",") > 0 Then
                        strTmp = Split(colǩ��ID("_" & lngǩ��id), "=")(1)
                        If .Cell(flexcpData, i, COL_ѡ��) = 0 And Not .Cell(flexcpPicture, i, COL_ѡ��) Is Nothing Then
                            If InStr(strTmp, "1") = 0 Then
                                colǩ��ID.Remove "_" & lngǩ��id
                                colǩ��ID.Add lngǩ��id & "=" & strTmp & "1", "_" & lngǩ��id
                            End If
                        Else
                            If InStr(strTmp, "0") = 0 Then
                                colǩ��ID.Remove "_" & lngǩ��id
                                colǩ��ID.Add lngǩ��id & "=" & strTmp & "0", "_" & lngǩ��id
                            End If
                        End If
                    Else
                        strǩ��ID = strǩ��ID & "," & lngǩ��id
                        If .Cell(flexcpData, i, COL_ѡ��) = 0 And Not .Cell(flexcpPicture, i, COL_ѡ��) Is Nothing Then
                            colǩ��ID.Add lngǩ��id & "=1", "_" & lngǩ��id
                        Else
                            colǩ��ID.Add lngǩ��id & "=0", "_" & lngǩ��id
                        End If
                    End If
                End If
            End If
        Next
            
        '���ǩ�����(һ��ǩ����ҽ������һ����)
        strTmp = ""
        For i = 1 To colǩ��ID.Count
            lngǩ��id = Split(colǩ��ID(i), "=")(0)
            strǩ��ID = Split(colǩ��ID(i), "=")(1)
            If Not (strǩ��ID = "1" Or strǩ��ID = "0") Then
                '���ǩ�������ݲ���"��Ҫ���ͻ򶼲�����"�����
                j = .FindRow(CStr(lngǩ��id), , COL_ǩ��ID)
                Do While j <> -1
                    If Not .RowHidden(j) Then
                        If .Cell(flexcpData, j, COL_ѡ��) = 1 Or .Cell(flexcpPicture, j, COL_ѡ��) Is Nothing Then
                            strTmp = strTmp & vbCrLf & "��" & .TextMatrix(j, col_ҽ������)
                        End If
                    End If
                    j = .FindRow(CStr(lngǩ��id), j + 1, COL_ǩ��ID)
                Loop
                Exit For '��ֻ��ʾ��һ��
            End If
        Next
    End With
    
    If strTmp <> "" Then
        MsgBox "����ҽ������������Ҫ���͵�ҽ��һ��ǩ��������ǰ����Ϊ�����ͣ�" & vbCrLf & strTmp & _
            vbCrLf & vbCrLf & "һ��ǩ����ҽ������һ���ͣ���������ҽ���ķ���״̬��", vbInformation, gstrSysName
        Exit Function
    End If
    CheckSignSend = True
End Function

Private Sub SeekPriceRow(ByVal lngRow As Long, ByVal lng��ĿID As Long, ByVal int�������� As Integer, ByVal lngCol As Long)
'���ܣ���λ������ʾָ��ҽ����ָ���Ƽ���
'������lngRow=ҽ���к�
'      lng��ĿID=�Ƽ���ĿID
'      lngCol=�Ƽ۱����ʾ��
    Dim k As Long
    
    With vsAdvice
        .Col = col_ҽ������ '�������Զ�ShowPrice,mrsPrice�����仯
        If Not .RowHidden(lngRow) Then
            .Row = lngRow
        Else
            If InStr(",F,D,G,C,", .TextMatrix(lngRow, COL_�������)) > 0 And Val(.TextMatrix(lngRow, COL_���ID)) <> 0 Then
                '��������,��������,��鲿λ,���������Ŀ
                .Row = .FindRow(CStr(Val(.TextMatrix(lngRow, COL_���ID))), , COL_ID)
            ElseIf CLng(.Cell(flexcpData, lngRow, COL_ID)) = 1 Then
                '��ҩ;��
                .Row = .FindRow(CStr(Val(.TextMatrix(lngRow, COL_ID))), , COL_���ID)
            ElseIf CLng(.Cell(flexcpData, lngRow, COL_ID)) = 2 Then
                '��ҩ�巨
                .Row = .FindRow(CStr(Val(.TextMatrix(lngRow, COL_���ID))), lngRow + 1, COL_ID)
            End If
        End If
        For k = vsPrice.FixedRows To vsPrice.Rows - 1
            If Val(vsPrice.TextMatrix(k, COLP_�к�)) = lngRow _
                And Val(vsPrice.TextMatrix(k, COLP_��������)) = int�������� _
                And Val(vsPrice.TextMatrix(k, COLP_�շ�ϸĿID)) = lng��ĿID Then
                vsPrice.Row = k: vsPrice.Col = lngCol: Exit For
            End If
        Next
        Call .ShowCell(.Row, .Col)
        Call vsPrice.ShowCell(vsPrice.Row, vsPrice.Col)
    End With
End Sub

Private Function GetMergeDrugStore(ByVal lngRow As Long) As Long
'���ܣ���ȡһ����ҩ�Ļ�׼ҩ�����������ɷ���NO��Keyֵ
'˵����һ����ҩ��ҩƷ���͵�һ�𣬰����Ա�ҩ�Ͳ�ͬҩ�������
    Dim lngҩ��ID As Long, lngBegin As Long, i As Long
    
    With vsAdvice
        If Val(.TextMatrix(lngRow, COL_���ID)) <> Val(.TextMatrix(lngRow - 1, COL_���ID)) And Val(.TextMatrix(lngRow, COL_ִ�п���ID)) <> 0 Then
            lngҩ��ID = Val(.TextMatrix(lngRow, COL_ִ�п���ID))
        Else
            lngBegin = .FindRow(.TextMatrix(lngRow, COL_���ID), , COL_���ID)
            For i = lngBegin To .Rows - 1
                If Val(.TextMatrix(i, COL_���ID)) = Val(.TextMatrix(lngRow, COL_���ID)) Then
                    If Val(.TextMatrix(i, COL_ִ�п���ID)) <> 0 Then
                        lngҩ��ID = Val(.TextMatrix(i, COL_ִ�п���ID)): Exit For
                    End If
                Else
                    Exit For
                End If
            Next
        End If
    End With
    
    GetMergeDrugStore = lngҩ��ID
End Function

Private Function CheckWaitExecute(ByVal lngRow As Long, ByVal byt��Ŀ��鷽ʽ As Byte, ByVal bytҩƷ��鷽ʽ As Byte) As Boolean
'���ܣ�����ָ���ļ�鷽ʽ���Բ���δִ�е���Ȧ��δ��ҩƷ���м��
'������byt��鷽ʽ=0-�����,1-��鲢��ʾ,2-��鲢��ֹ
'���أ��Ƿ����
    Dim strTmp As String
        
    With vsAdvice
        If byt��Ŀ��鷽ʽ <> 0 Then
            strTmp = ExistWaitExe(mlng����ID, mlng��ҳID, -1)
            If strTmp <> "" Then
                Call .ShowCell(lngRow, col_ҽ������): .Refresh
                If byt��Ŀ��鷽ʽ = 1 Then
                    If MsgBox("���ֲ���""" & mrsPati!���� & """������δִ����ɵ����ݣ�" & _
                        vbCrLf & vbCrLf & strTmp & vbCrLf & vbCrLf & "ȷʵҪ����""" & .TextMatrix(lngRow, col_ҽ������) & """��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        Exit Function
                    End If
                Else
                    MsgBox "���ֲ���""" & mrsPati!���� & """������δִ����ɵ����ݣ�" & _
                        vbCrLf & vbCrLf & strTmp & vbCrLf & vbCrLf & "ҽ��""" & .TextMatrix(lngRow, col_ҽ������) & """���������͡�", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        End If
        
        If bytҩƷ��鷽ʽ <> 0 Then
            strTmp = ExistWaitDrug(mlng����ID, mlng��ҳID, -1)
            If strTmp <> "" Then
                Call .ShowCell(lngRow, col_ҽ������): .Refresh
                If bytҩƷ��鷽ʽ = 1 Then
                    If MsgBox("���ֲ���""" & mrsPati!���� & """" & _
                        strTmp & vbCrLf & vbCrLf & "ȷʵҪ����""" & .TextMatrix(lngRow, col_ҽ������) & """��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        Exit Function
                    End If
                Else
                    MsgBox "���ֲ���""" & mrsPati!���� & """" & _
                        strTmp & vbCrLf & vbCrLf & "ҽ��""" & .TextMatrix(lngRow, col_ҽ������) & """����������", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        End If
    End With
    
    CheckWaitExecute = True
End Function

Private Function SendAdvice() As Long
'���ܣ�����ҽ������(��������м��ʱ���)
'˵����������˷����ύ
'���أ�����ɹ��򷵻ط��ͺ�
    Dim rsSQL As ADODB.Recordset
    Dim rsTotal As ADODB.Recordset
    Dim rsUpload As ADODB.Recordset
    Dim rsNumber As ADODB.Recordset
    Dim rsItems As ADODB.Recordset '����ҽ���ܿصķ��ü�¼��,��̬��¼��
    Dim rsMoneyNow As ADODB.Recordset '��ǰ���˱���Ҫ���͵ķ���,��̬��¼��
    Dim rsMoneyDay As ADODB.Recordset '��ǰ���˵����ѷ��͵ķ���,��̬��¼��
    Dim rsTmp As ADODB.Recordset
    Dim rsMoney As New ADODB.Recordset

    Dim i As Long, j As Long
    Dim strSQL As String, curDate As Date, strCurDate As String
    Dim blnTran As Boolean, blnBool As Boolean, strTmp As String
    Dim str��� As String, str���� As String
    
    Dim lng���ͺ� As Long, int�Ʒ�״̬ As Integer, bln���� As Boolean, int���� As Integer, strNO As String, intִ��״̬ As Integer
    Dim str�շ���Ŀ As String, lng������� As Long, lng���ø��� As Long, lng������� As Long
    Dim int���� As Integer, dbl���� As Double, cur�ϼ� As Currency, cur���ʺϼ� As Currency
    Dim dbl���� As Double, dblӦ�� As Double, curӦ�� As Currency, curʵ�� As Currency
    Dim str�ֽ�ʱ�� As String, str�״�ʱ�� As String, strĩ��ʱ�� As String, strPre���Ƶ���ID As String
    Dim int�䷽�� As Integer, strNOKey As String, str�Զ����� As String
    Dim str����ʱ�� As String, str�Ǽ�ʱ�� As String
    Dim dbl�������� As Double, blnFirst As Boolean '�䷽�����ֺŹؼ���
    Dim lngҩƷ���ID As Long, lng�������ID As Long
    Dim lngִ�п���ID As Long, lng���˿���ID As Long
    Dim intҩƷ���� As Integer, bln�������� As Boolean
    
    Dim rsClone As ADODB.Recordset
    Dim rsSeek As ADODB.Recordset
    Dim rsExec As ADODB.Recordset  'ҽ��ִ�мƼ�
    Dim strNoneSub As String, strHaveSub As String
    Dim int����� As Integer, lng����ĿID As Long, strʵ�� As String
    
    Dim blnҩƷʱ����ʾ As Boolean, blnҩƷ�����ʾ As Boolean, blnҩƷĬ�Ϸ��� As Boolean
    Dim bln����ʱ����ʾ As Boolean, bln���Ŀ����ʾ As Boolean, bln����Ĭ�Ϸ��� As Boolean
    Dim bln������Ŀ�� As Boolean, lng���մ���ID As Long, curͳ���� As Currency, str���ձ��� As String, str�������� As String
    Dim str����ҽ��IDs As String
    Dim rsAudit As ADODB.Recordset, strAudit As String
    '����ǩ��
    Dim lng��ID As Long, strҽ��IDs As String, strSource As String
    Dim intRule As Integer, strSign As String, strTimeStamp As String, strTimeStampCode As String
    Dim lng֤��ID As Long, lngǩ��id As Long
    
    Dim str��ҩ�� As String, blnʵʱ��� As Boolean, strCuvetteNumber As String '��������
    Dim lng���ô��� As Long 'һ��ֻ��һ��ʱ�����η���Ӧ��ȡ�ķ��ô���
    Dim blnCheckAdvice As Boolean, strMsg As String, lngSpecialAdviceID, lngBabyNum As Long
    Dim str����ҽ�� As String
    Dim lng��ҽ���� As Long
    Dim lng�ɼ�����ID As Long
    Dim str��ҩIDs As String, str������ҩIDs As String, strҽ������ As String
    Dim bln�Զ�ִ�� As Boolean
    Dim str��λ���� As String '�����Ŀ�Ĳ�λ�������̶���ʽ����鲿λ<sTab>��鷽�����磺"ͷ��<sTab>ƽɨ"
    Dim dblOther���� As Double '������Ŀ�շѴ���
    Dim str����ҩ��  As String '������ҩƷ��ҽ�� ,"Ƥ��ҽ��ID,ҩƷ��ҽ��ID"
    Dim rsƤ�� As ADODB.Recordset
    Dim strMinDate As String
    
    On Error GoTo errH
    
    '��������������ҩ�󷽽���ж�
    Call Check�������
    
    '���һ��ǩ����ҽ���Ƿ�һ����
    If Not CheckSignSend Then Exit Function
    
    'RISԤԼ����ж���ʾ
    Call CheckRISScheduling
    
    blnCheckAdvice = Val(zlDatabase.GetPara("����ҽ������ǰ���δ��Чҽ��", glngSys, pסԺҽ������, 0)) = 1
    
    With vsAdvice
        If mbytSendKind <> EOutCharge And mbytSendKind <> EOutBilling Then
        '�ݲ�֧����������
            If InitObjRecipeAudit(pסԺҽ���´�) Then
                '�������ϵͳ������������
                For i = .FixedRows To .Rows - 1
                    If .Cell(flexcpData, i, COL_ѡ��) = 0 And Not .Cell(flexcpPicture, i, COL_ѡ��) Is Nothing Then
                        If .TextMatrix(i, COL_�������) = "E" And .TextMatrix(i, COL_��������) = "2" Then
                            str��ҩIDs = str��ҩIDs & "," & .TextMatrix(i, COL_ID)
                        End If
                    End If
                Next
                If Mid(str��ҩIDs, 2) <> "" Then
                    Call gobjRecipeAudit.BuildData(Mid(str��ҩIDs, 2), mlng���˿���id, 1, mlng����ID, mlng��ҳID, str������ҩIDs)
                End If
            End If
        End If
        
        '�ȼ�鲢��ʾ����ҽ��:3-ת��;4-����;5-��Ժ;6-תԺ,11-����,14-��ǰ
        strTmp = ""
        strMinDate = "3000-01-01 00:00"
        Call InitExecRecordset(rsExec)   'ҽ��ִ�мƼ�
        For i = .FixedRows To .Rows - 1
            If .Cell(flexcpData, i, COL_ѡ��) = 0 And Not .Cell(flexcpPicture, i, COL_ѡ��) Is Nothing Then
                If .TextMatrix(i, COL_�������) = "Z" And InStr(",3,5,6,11,", Val(.TextMatrix(i, COL_��������))) > 0 Then
                    lngSpecialAdviceID = .TextMatrix(i, COL_ID) 'ֻȡ����һ��������ҽ��������һ��
                    lngBabyNum = Val(.Cell(flexcpData, i, COL_Ӥ��))
                    strTmp = strTmp & vbCrLf & mrsPati!���� & IIF(.Cell(flexcpData, i, COL_Ӥ��) <> 0, "(Ӥ��" & .Cell(flexcpData, i, COL_Ӥ��) & ")", "") & "��" & .TextMatrix(i, col_ҽ������)
                    
                    'ת��ҽ������ʱ�жϳ������Լ�������
                    If Val(.TextMatrix(i, COL_��������)) = 3 Then
                        If CheckCanSendAdvice(mlng����ID, mlng��ҳID, lngSpecialAdviceID, lngBabyNum) Then
                            Call MsgBox("����ת��ҽ����" & vbCrLf & mrsPati!���� & IIF(.Cell(flexcpData, i, COL_Ӥ��) <> 0, "(Ӥ��" & .Cell(flexcpData, i, COL_Ӥ��) & ")", "") & "��" & .TextMatrix(i, col_ҽ������) & vbCrLf & vbCrLf & "���뽫���Է��͵ĳ���ҽ���������ܷ��͡�", vbInformation, gstrSysName)
                            Exit Function
                        End If
                    End If

                    'ת��ʱδ������ʵ��ݼ��
                    If Val(.TextMatrix(i, COL_��������)) = 3 Then
                        If CheckWaitQuittance(mlng����ID, mlng��ҳID) Then Exit Function
                    End If
                    
                End If
                
                If Mid(gstrESign, 2, 1) = "1" Then 'סԺҽ��վ�����˵���ǩ���ż��
                    If .TextMatrix(i, COL_�������) = "Z" And InStr(",3,5,6,11,4,14,", Val(.TextMatrix(i, COL_��������))) > 0 Then
                        str����ҽ�� = str����ҽ�� & "," & .TextMatrix(i, col_ҽ������)
                    End If
                End If
                '��������ж���Ϣ�ռ�
                If gbln����ҩƷ�ֿ����� Then
                    If cboDrugType.ListIndex = 0 Then
                        If InStr("," & str���� & ",", "," & .TextMatrix(i, COL_�������) & ",") = 0 Then
                            str���� = str���� & "," & .TextMatrix(i, COL_�������)
                        End If
                    ElseIf cboDrugType.ListIndex = 3 Then
                        str���� = ""
                    Else
                        str���� = ",����ҩ"
                    End If
                End If
                If .TextMatrix(i, COL_�״�ʱ��) < strMinDate Then
                    strMinDate = .TextMatrix(i, COL_�״�ʱ��)
                End If
            End If
        Next
        If strMinDate = "3000-01-01 00:00" Then strMinDate = ""
        
        If str���� <> "" And cboDrugType.ListIndex = 0 Then
            If Not (str���� = ",����ҩ" Or str���� = ",����I��" Or str���� = ",����ҩ" Or str���� = ",����ҩ,����I��" Or str���� = ",����I��,����ҩ") Then
                If Not (InStr(str���� & ",", ",����ҩ,") = 0 And InStr(str���� & ",", ",����ҩ,") = 0 And InStr(str���� & ",", ",����I��,") = 0) Then
                    MsgBox "���η��͵�ҽ���п��ܰ������龫��ҩƷ����ֱ��ͣ����޸Ĺ����������¶�ȡҽ�����ٷ��͡�", vbInformation, gstrSysName
                    Exit Function
                Else
                    str���� = ""
                End If
            End If
        End If
        
        If strTmp <> "" Then
            If blnCheckAdvice Then
                strMsg = CheckUnExecutedAdvice(mlng����ID, mlng��ҳID, lngSpecialAdviceID, lngBabyNum)
            End If
            
            If strMsg <> "" Then
                Call MsgBox("������������ҽ����" & vbCrLf & strTmp & vbCrLf & vbCrLf & "���뽫" & strMsg & "�������ܷ��͡�", vbInformation, gstrSysName)
                Exit Function
                
            ElseIf MsgBox("Ҫ���͵�ҽ���а�����������ҽ����" & vbCrLf & strTmp & vbCrLf & vbCrLf & "ȷʵҪ���͵�ǰѡ���ҽ����", _
                vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Function
            End If
        End If
        
        '��������˵���ǩ����������"��ֹͣ��δȷ��ֹͣ"��ҽ������ʾ��ʿ�Ƚ���ȷ��ֹͣ
        '��Ϊ����ҽ��У��ʱ�Ὣ"��ֹͣ��δȷ��ֹͣ"��ҽ����"ִ����ֹʱ��"����Ϊ����ҽ���Ŀ�ʼִ��ʱ�䣬ҽ��ֹͣ��ǩ��Դ�İ�����"ִ����ֹʱ��"����ᵼ��ǩ����֤�޷�ͨ��
        If str����ҽ�� <> "" Then
            str����ҽ�� = Mid(str����ҽ��, 2)
            If CheckStopedUnAffirm(mlng����ID & ":" & mlng��ҳID, "") Then
                MsgBox "Ҫ���͵�ҽ���а�������ҽ����" & vbCrLf & str����ҽ�� & _
                    vbCrLf & vbCrLf & "���ͺ�Ὣδȷ��ֹͣ��ҽ������ֹͣ��Ϊ�˲�Ӱ��ǩ����֤�����Ƚ���ȷ��ֹͣ������", vbInformation, gstrSysName
                Exit Function
            End If
        End If
        
        '������ҩ
        If mbln������ҩ Then
            blnBool = Set������ҩ()
            If Not blnBool Then
                GoTo FuncEnd
            End If
        End If
        
        If Not zlPluginAdviceBeforeSend Then
            Exit Function
        End If
    End With
    
    '��ȡ��ǰ���˵�������Ŀ�嵥
    strAudit = ""
    If Not IsNull(mrsPati!����) Then
        blnʵʱ��� = gclsInsure.GetCapability(supportʵʱ���, mlng����ID, mrsPati!����)
        If Val(zlDatabase.GetPara("���ҽ������", glngSys, pסԺҽ������, "1")) = 1 Then
            Set rsAudit = GetAuditRecord(mlng����ID, mlng��ҳID)
        Else
            Set rsAudit = Nothing
        End If
    Else
        blnʵʱ��� = False
        Set rsAudit = Nothing '��NothingΪ��־�ò��˲���Ҫ�ж�
    End If
    
    '��ȡҩƷ/����������
    lngҩƷ���ID = ExistIOClass(IIF(mbytSendKind = EOutCharge, 8, 9))    '8-�շѴ�����ҩ��9-���ʵ�������ҩ
    lng�������ID = ExistIOClass(IIF(mbytSendKind = EOutCharge, 40, 41))        '40-�շѴ�������,41-���ʵ���������
    
    Screen.MousePointer = 11
    
    mstr��ҩ�� = ""
    blnҩƷʱ����ʾ = True: blnҩƷ�����ʾ = True: blnҩƷĬ�Ϸ��� = True
    bln����ʱ����ʾ = True: bln���Ŀ����ʾ = True: bln����Ĭ�Ϸ��� = True
    
    Call InitBillSet
    Call InitRecordSet(rsSQL, rsTotal, rsUpload, rsNumber, rsMoneyNow, rsItems)
    lng���ͺ� = zlDatabase.GetNextNo(10)
    mlngNOSequence = 0 '���ݺ��������³�ʼ
    
    '���ʱ�䷢�͹�����δ����ֹͣʱ��,Ϊ������У��ʱ���ظ�(ȡ��Sysdate)
    curDate = zlDatabase.Currentdate
    strCurDate = "To_Date('" & Format(curDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
    bln���� = True '��ʼȫ���ǻ���
    int�䷽�� = 1 '��ʾ���͵ĵڼ����䷽,���ڷֵ��ݺ�
    
    With vsAdvice
        For i = .FixedRows To .Rows - 1
            If .Cell(flexcpData, i, COL_ѡ��) = 0 And Not .Cell(flexcpPicture, i, COL_ѡ��) Is Nothing Then
                '����ҽ����3-ת��;5-��Ժ;6-תԺ,11-����
                If .TextMatrix(i, COL_�������) = "Z" Then
                    'ת��,��Ժ,תԺ,����ҽ������ʱ������Ҫ��������״̬
                    If .Cell(flexcpData, i, COL_Ӥ��) = 0 Then
                        If InStr(",3,5,6,11,", .TextMatrix(i, COL_��������)) > 0 And NVL(mrsPati!״̬, 0) <> 0 Then
                            MsgBox "����""" & mrsPati!���� & """��ǰ����""" & decode(NVL(mrsPati!״̬, 0), 1, "�ȴ����", 2, "����ת��", 3, "��Ԥ��Ժ") & """״̬��" & _
                                "���ܷ���""" & .TextMatrix(i, col_ҽ������) & """ҽ����", vbInformation, gstrSysName
                            Set .Cell(flexcpPicture, i, COL_ѡ��) = Nothing
                            Call DeleteRsExec(rsExec, Val(.TextMatrix(i, COL_ID)))
                            GoTo NextAdvice
                        End If
                    End If
                    
                    '�����ת�ơ���Ժ��תԺҽ��,��鲡���Ƿ���δִ�е�ҽ����Ŀ��δ��ҩƷ
                    If InStr(",3,", .TextMatrix(i, COL_��������)) > 0 Then
                        If Not CheckWaitExecute(i, gbytת�Ƽ��δִ��, gbytת�Ƽ��δ��ҩ) Then
                            Set .Cell(flexcpPicture, i, COL_ѡ��) = Nothing
                            Call DeleteRsExec(rsExec, Val(.TextMatrix(i, COL_ID)))
                            GoTo NextAdvice
                        End If
                    End If
                    If InStr(",5,6,", .TextMatrix(i, COL_��������)) > 0 Then
                        If Not CheckWaitExecute(i, gbyt��Ժ���δִ��, gbyt��Ժ���δ��ҩ) Then
                            Set .Cell(flexcpPicture, i, COL_ѡ��) = Nothing
                            Call DeleteRsExec(rsExec, Val(.TextMatrix(i, COL_ID)))
                            GoTo NextAdvice
                        End If
                    End If
                End If
            
                '�������ݺŷ���ؼ���
                '-----------------------------------------------------------------------------------------
                If InStr(",5,6,", .TextMatrix(i, COL_�������)) > 0 Then
                    '���ò���������ҩƷ�ֿ����� ʱ������ҩƷҽ����ҩƷ�е������ɵ��ݺţ�һ��ҽ������һ����
                    If str���� <> "" Then
                        strNOKey = "������ҩ_" & .TextMatrix(i, COL_���ID)
                    Else
                        '������ҩ��"����(����ID,�Һŵ�)_���˿���ID_��������ID_����ҽ��_ִ�п���ID"�ֺš�
                        'һ����ҩ�ģ����͵�һ�𣺰����Ա�ҩ�Ͳ�ͬҩ�������
                        strNOKey = "������ҩ_" & mlng����ID & "_" & mlng��ҳID & "_" & _
                            Val(.TextMatrix(i, COL_���˿���ID)) & "_" & Val(.TextMatrix(i, COL_��������ID)) & "_" & _
                            .TextMatrix(i, COL_����ҽ��) & "_" & GetMergeDrugStore(i)
                        
                        If mblnһ����ҩ����Ϊһ�� Then
                            If Val(.TextMatrix(i, COL_���ID)) <> Val(.TextMatrix(i - 1, COL_���ID)) Then
                                '�ٰ�Ҫ��ӡ�����Ƶ��ݷֺ�(һ����ҩ�ģ�ֻȡ��һ��ҩƷ�����Ƶ���ID)
                                strPre���Ƶ���ID = GetClinicBillID(Val(.TextMatrix(i, COL_������ĿID)), IIF(mlng�������� = 1, 1, 2))
                            End If
                            strNOKey = strNOKey & "_" & strPre���Ƶ���ID
                        Else
                            strNOKey = strNOKey & "_" & GetClinicBillID(Val(.TextMatrix(i, COL_������ĿID)), IIF(mlng�������� = 1, 1, 2))
                        End If
                        
                        '�ٰ������������ƽ��зֺ�
                        If gintRXCount > 0 And mlng�������� = 1 Then
                            strTmp = ""
                            If Val(.TextMatrix(i, COL_���ID)) <> Val(.TextMatrix(i - 1, COL_���ID)) Then
                                strTmp = GetMergeIDs(vsAdvice, i, COL_���ID, COL_ID) 'һ����ҩ��ʼ�л����ҩƷ�в�ȡֵ
                            End If
                            strNOKey = strNOKey & "_" & GetRXKey(mrsRXKey, strNOKey, strTmp)
                        End If
                        '��ҩִ�п��Ҳ���ͬ������䲻ͬ��NO��
                        j = .FindRow(CStr(.TextMatrix(i, COL_���ID)), i + 1, COL_ID)
                        If j > 0 Then strNOKey = strNOKey & "_" & Val(.TextMatrix(j, COL_ִ�п���ID))
                    End If
                ElseIf InStr(",4,M,", .TextMatrix(i, COL_�������)) > 0 Then
                    '���ϰ�"����(����ID,�Һŵ�)_���˿���ID_��������ID_����ҽ��_ִ�п���ID"�ֺš�
                    strNOKey = "����ҽ��_" & mlng����ID & "_" & mlng��ҳID & "_" & _
                        Val(.TextMatrix(i, COL_���˿���ID)) & "_" & Val(.TextMatrix(i, COL_��������ID)) & "_" & _
                        .TextMatrix(i, COL_����ҽ��) & "_" & Val(.TextMatrix(i, COL_ִ�п���ID))
                    '�ٰ�Ҫ��ӡ�����Ƶ��ݷֺ�
                    strNOKey = strNOKey & "_" & GetClinicBillID(Val(.TextMatrix(i, COL_������ĿID)), 2)
                ElseIf .TextMatrix(i, COL_�������) = "7" Then
                    'һ���䷽�е����в�ҩ����һ���������ݺ�
                    strNOKey = "��ҩ�䷽_" & mlng����ID & "_" & mlng��ҳID & "_" & int�䷽��
                ElseIf Val(.TextMatrix(i, COL_���ID)) <> 0 And .TextMatrix(i, COL_�������) = "C" Then
                    'һ���ɼ��ļ�����Ϸ�����ͬ�ĵ��ݺţ��걾�ɼ��������䵥���ĵ��ݺ�
                    'ͬһ��������ͣ�ͬһ������ִ�п��ң�ͬһ�ɼ��ܣ�ͬһ���ɼ���ʽ��ͬһ���ɼ�ִ�п��ҵļ��������ͬ�ĵ��ݺ�
                    If mbln���鵥���������� Then
                        strNOKey = "һ���ɼ�_" & Val(.TextMatrix(i, COL_���ID))
                    Else
                        lng��ҽ���� = .FindRow(CStr(.TextMatrix(i, COL_���ID)), i + 1, COL_ID)
                        strNOKey = "һ���ɼ�_" & mlng����ID & "_" & mlng��ҳID & "_" & .TextMatrix(i, COL_�걾��λ) & "_" & _
                            .TextMatrix(i, COL_ִ�п���ID) & "_" & .TextMatrix(i, COL_��������) & "_" & .TextMatrix(i, COL_�Թܱ���) & "_" & _
                            .TextMatrix(lng��ҽ����, COL_������ĿID) & "_" & .TextMatrix(lng��ҽ����, COL_ִ�п���ID)
                    End If
                ElseIf Val(.TextMatrix(i, COL_���ID)) <> 0 And InStr(",F,D,", .TextMatrix(i, COL_�������)) > 0 Then
                    '��鲿λ�͸�����������Ҫҽ��������ͬ���ݺţ�����������䵥���ĵ��ݺš�
                    strNOKey = "��ҩҽ��_" & Val(.TextMatrix(i, COL_���ID))
                Else
                    '������ҩҽ��ÿ��ҽ��һ���������ݺ�(������ҩ;�����䷽�巨���÷����ɼ���ʽ������ʽ����Ѫҽ��/��Ѫ;��)
                    strNOKey = "��ҩҽ��_" & Val(.TextMatrix(i, COL_ID))
                End If
                
                '�Ƿ���Ժ��ҩ
                intҩƷ���� = 0
                If InStr(",5,6,", .TextMatrix(i, COL_�������)) > 0 Then
                    intҩƷ���� = decode(.TextMatrix(i, COL_ִ������), "��Ժ��ҩ", 3, "��ȡҩ", 4, intҩƷ����)
                ElseIf .TextMatrix(i, COL_�������) = "7" Then
                    j = .FindRow(CStr(.TextMatrix(i, COL_���ID)), i + 1, COL_ID)
                    If j <> -1 Then
                        intҩƷ���� = decode(.TextMatrix(j, COL_ִ������), "��Ժ��ҩ", 3, "��ȡҩ", 4, intҩƷ����)
                    End If
                End If
                
                '����ҽ�����ʷ���:�����¼۸����
                '-----------------------------------------------------------------------------------------
                strSQL = "": str�շ���Ŀ = ""
                If InStr(",5,6,7,", .TextMatrix(i, COL_�������)) > 0 Then
                    'ҩƷȱʡ�̶�Ϊ�����Ƽ�,����ҽ��ʱָ����Ϊ�Ա�ҩ(Ժ��ִ��)�Ĳ���ȡ;ҩƷ������Ϊ����
                    If Val(.TextMatrix(i, COL_ִ������ID)) <> 5 Then
                        strSQL = _
                            " Select A.ID,A.���,D.���� as �������,RTrim(A.����||' '||A.���) as ����," & _
                            " A.���㵥λ,A.�Ƿ���,A.���ηѱ�,A.����ȷ��,A.�Ӱ�Ӽ�,B.�Ӱ�Ӽ���,100 as �����շ���," & _
                            " Y.סԺ��λ,Y.סԺ��װ,Y.����ϵ��,Y.ҩ������ as ����,0 as ��������,B.������ĿID," & _
                            " C.�վݷ�Ŀ,1 as ����,B.�ּ� as ����,[2] as ִ�п���ID,0 as ����,0 as ��������,0 as �շѷ�ʽ,I.Ҫ������" & _
                            " From �շ���ĿĿ¼ A,�շѼ�Ŀ B,������Ŀ C,�շ���Ŀ��� D,ҩƷ��� Y,����֧����Ŀ I" & _
                            " Where A.ID=B.�շ�ϸĿID And B.������ĿID=C.ID And A.���=D.����" & _
                            GetPriceGradeSQL(mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�, "A", "B", "4", "5", "6") & _
                            " And A.ID=Y.ҩƷID(+) And A.ID=[1] And A.ID=I.�շ�ϸĿID(+) And I.����(+)=[3]" & _
                            " And ((Sysdate Between B.ִ������ and B.��ֹ����) Or (Sysdate>=B.ִ������ And B.��ֹ���� is NULL))" & _
                            " Order by A.����"
                    End If
                Else
                    '��ɾ��ԭ��ҩҽ���ļƼ�(Ӧ��û��)
                    rsSQL.AddNew
                    rsSQL!���� = 1: rsSQL!��ĿID = 0: rsSQL!��� = i
                    rsSQL!ҽ��ID = IIF(Val(.TextMatrix(i, COL_���ID)) = 0, Val(.TextMatrix(i, COL_ID)), Val(.TextMatrix(i, COL_���ID)))
                    rsSQL!Sql = "ZL_����ҽ���Ƽ�_Delete(" & Val(.TextMatrix(i, COL_ID)) & ",1)"
                    rsSQL.Update
                    
                    '���Ƽ�,�ֹ��Ƽۣ�����,Ժ��ִ�е�ҽ������ȡ
                    If Val(.TextMatrix(i, COL_�Ƽ�����)) = 0 And InStr(",0,5,", Val(.TextMatrix(i, COL_ִ������ID))) = 0 Then
                        mrsPrice.Filter = "ҽ��ID=" & Val(.TextMatrix(i, COL_ID))
                        If Not mrsPrice.EOF Then
                            For j = 1 To mrsPrice.RecordCount
                                If NVL(mrsPrice!�շ�ϸĿID, 0) <> 0 And NVL(mrsPrice!����, 0) <> 0 Then '��������Ϊ0���Զ����˵�
                                    '��ͨ��Ŀ�ı�۵���Ҫ�����룬�����Ǹ������õ�ʱ������ҽ��
                                    If NVL(mrsPrice!����, 0) = 0 And NVL(mrsPrice!���, 0) = 1 _
                                        And Not (InStr(",5,6,7,", mrsPrice!�շ����) > 0 Or mrsPrice!�շ���� = "4" And NVL(mrsPrice!����, 0) = 1) Then
                                        Call SeekPriceRow(i, mrsPrice!�շ�ϸĿID, mrsPrice!��������, COLP_����)
                                        Screen.MousePointer = 0
                                        MsgBox "����Ϊ��۵��շ���Ŀȷ��һ���շѼ۸�", vbInformation, gstrSysName
                                        vsPrice.SetFocus: GoTo FuncEnd
                                    End If
                                    
                                    '�Ƽ�ִ�п���:ֻ�����ҩƷ������ҽ���ģ�ҩƷ�����ļƼ۵�ִ�п���
                                    If InStr(",4,5,6,7,", .TextMatrix(i, COL_�������)) = 0 _
                                        And (InStr(",5,6,7,", mrsPrice!�շ����) > 0 Or mrsPrice!�շ���� = "4" And NVL(mrsPrice!����, 0) = 1) Then
                                        lngִ�п���ID = NVL(mrsPrice!ִ�п���ID, 0)
                                        
                                        '���ı�������ִ�п���
                                        If lngִ�п���ID = 0 And mrsPrice!�շ���� = "4" Then
                                            Call SeekPriceRow(i, mrsPrice!�շ�ϸĿID, mrsPrice!��������, COLP_ִ�п���)
                                            Screen.MousePointer = 0
                                            MsgBox "����""" & vsPrice.TextMatrix(vsPrice.Row, COLP_�շ���Ŀ) & """û��ȷ��ִ�п��ң����ֹ�������ȷ��ִ�п��ҡ�" & vbCrLf & _
                                                "�������ȷ����ȷ��ִ�п��ң��뵽""����Ŀ¼����""�м��洢�ⷿ�����Ƿ���ȷ��", vbInformation, gstrSysName
                                            vsPrice.SetFocus: GoTo FuncEnd
                                        End If
                                    Else
                                        lngִ�п���ID = 0
                                    End If
                                    
                                    'ҩƷ������ҽ���ļƼ۹̶���Ӧ�����棻�Ǹ������õ�ʱ�����ĵı����Ҫ���룬���Ҫ���浽�Ƽ۱���
                                    If InStr(",4,5,6,7,", .TextMatrix(i, COL_�������)) = 0 _
                                        Or .TextMatrix(i, COL_�������) = "4" And NVL(mrsPrice!����, 0) = 0 And NVL(mrsPrice!���, 0) = 1 Then
                                        rsSQL.AddNew
                                        rsSQL!���� = 1: rsSQL!��ĿID = mrsPrice!�շ�ϸĿID: rsSQL!��� = i
                                        rsSQL!ҽ��ID = IIF(Val(.TextMatrix(i, COL_���ID)) = 0, Val(.TextMatrix(i, COL_ID)), Val(.TextMatrix(i, COL_���ID)))
                                        rsSQL!Sql = "ZL_����ҽ���Ƽ�_INSERT(" & _
                                            mrsPrice!ҽ��ID & "," & mrsPrice!�շ�ϸĿID & "," & _
                                            NVL(mrsPrice!����, 0) & "," & NVL(mrsPrice!����, 0) & "," & _
                                            NVL(mrsPrice!����, 0) & "," & ZVal(lngִ�п���ID) & "," & _
                                            NVL(mrsPrice!��������, 0) & "," & NVL(mrsPrice!�շѷ�ʽ, 0) & ")"
                                        rsSQL.Update
                                    End If
                                    
                                    '��ʱ����ҽ���Ƽ۱�
                                    If Val(.TextMatrix(i, COL_����)) <> 0 Then
                                        strSQL = strSQL & IIF(strSQL = "", "", " Union ALL ") & _
                                            "Select " & mrsPrice!�շ�ϸĿID & " as �շ�ϸĿID," & _
                                            NVL(mrsPrice!ִ�п���ID, 0) & " as ִ�п���ID," & _
                                            NVL(mrsPrice!����, 0) & " as ����," & Format(NVL(mrsPrice!����, 0), gstrDecPrice) & " as ����," & _
                                            NVL(mrsPrice!����, 0) & " as ����," & NVL(mrsPrice!��������, 0) & " as ��������," & _
                                            NVL(mrsPrice!�շѷ�ʽ, 0) & " as �շѷ�ʽ From Dual"
                                    End If
                                End If
                                
                                mrsPrice.MoveNext
                            Next
                        End If
                    End If
                    
                    If strSQL <> "" Then
                        strSQL = _
                            " Select A.ID,A.���,D.���� as �������,A.����,A.���㵥λ,A.�Ƿ���," & _
                            " A.���ηѱ�,A.����ȷ��,A.�Ӱ�Ӽ�,B.�Ӱ�Ӽ���,B.�����շ���,Y.סԺ��λ,Y.סԺ��װ,Y.����ϵ��," & _
                            " Decode(A.���,'4',E.���÷���,Y.ҩ������) as ����,E.��������,B.������ĿID," & _
                            " C.�վݷ�Ŀ,X.����,Decode(A.�Ƿ���,1,X.����,B.�ּ�) as ����,X.ִ�п���ID," & _
                            " X.����,X.��������,X.�շѷ�ʽ,I.Ҫ������" & _
                            " From �շ���ĿĿ¼ A,�շѼ�Ŀ B,������Ŀ C,�շ���Ŀ��� D,�������� E,(" & strSQL & ") X,ҩƷ��� Y,����֧����Ŀ I" & _
                            " Where A.ID=B.�շ�ϸĿID And B.������ĿID=C.ID And A.ID=E.����ID(+)" & _
                            GetPriceGradeSQL(mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�, "A", "B", "4", "5", "6") & _
                            " And A.���=D.���� And X.�շ�ϸĿID=A.ID And A.ID=Y.ҩƷID(+)" & _
                            " And ((Sysdate Between B.ִ������ and B.��ֹ����) or (Sysdate>=B.ִ������ And B.��ֹ���� is NULL))" & _
                            " And A.ID=I.�շ�ϸĿID(+) And I.����(+)=[3]" & _
                            " Order by X.��������,X.����,X.�շѷ�ʽ Desc,A.ID"
                            'һ��Ҫ����������ǰ��,�Ա��ڼ�����ڷ��ü�¼�б������ӹ�ϵ
                    End If
                End If
                                
                '�����ۿ۱�����ʼ
                int����� = 0: lng����ĿID = 0
                strHaveSub = "": strNoneSub = ""
                Call InitSeekSet(rsSeek)
                
                '��ǰ������������(����"ҽ����������������"û������ʱҲ����һ����������룬�����ж��Ƿ��ղ�Ѫ�ܷ���)
                strCuvetteNumber = ""
                If Val(.TextMatrix(i, COL_ִ������ID)) <> 0 Then
                    j = .FindRow(CStr(.TextMatrix(i, COL_���ID)), i + 1, COL_ID)
                    If j > 0 Then lng�ɼ�����ID = Val(.TextMatrix(j, COL_ִ�п���ID))
                    strCuvetteNumber = GetCuvetteNumber(rsNumber, .TextMatrix(i, COL_�Թܱ���), _
                        Val(.TextMatrix(i, COL_ID)), Val(.TextMatrix(i, COL_���ID)), .TextMatrix(i, COL_�������), Val(.TextMatrix(i, COL_��������)), _
                        Val(.TextMatrix(i, COL_ִ�п���ID)), Val(.TextMatrix(i, COL_Ӥ��)), Val(.TextMatrix(i, COL_������ĿID)), _
                        Val(.TextMatrix(i, COL_������־)), .TextMatrix(i, COL_�걾��λ), lng�ɼ�����ID)
                End If
                
                '����ִ�е��Զ�ִ��(����ҽ�����ô���)
                intִ��״̬ = 0
                bln�Զ�ִ�� = False
                If mblnAutoExe Then
                    If Not (.TextMatrix(i, COL_�������) = "Z" And Val(.TextMatrix(i, COL_��������)) <> 0) _
                        And (mstrǰ��IDs = "" And (Val(.TextMatrix(i, COL_ִ�п���ID)) = Val(.TextMatrix(i, COL_���˿���ID)) Or _
                        Val(.TextMatrix(i, COL_ִ�п���ID)) = mlng���˲���ID) Or _
                            mstrǰ��IDs <> "" And Val(.TextMatrix(i, COL_ִ�п���ID)) = mlngҽ������ID) Then
                        bln�Զ�ִ�� = True
                    End If
                    If bln�Զ�ִ�� Then
                        bln�Զ�ִ�� = CanAutoExeItem(Val(.TextMatrix(i, COL_ִ�п���ID)), .TextMatrix(i, COL_�������), .TextMatrix(i, COL_��������), Val(.TextMatrix(i, COL_ִ�з���)))
                    End If
                    If bln�Զ�ִ�� Then
                        'ִ��ǰ�Ƚ���ʱ�������ڡ�ִ�к��Զ���˼��ʻ��۵���
                        If gblnִ��ǰ�Ƚ��� And (mbytSendKind = EOutBilling Or mbytSendKind = EOutCharge) Then
                            If Not gobjSquareCard Is Nothing Then
                                str����ҽ��IDs = str����ҽ��IDs & "," & .TextMatrix(i, COL_ID)
                            End If
                        Else
                            intִ��״̬ = 1
                            'Ѫ��������⴦��
                            If gblnѪ��ϵͳ Then
                                strTmp = .TextMatrix(i, COL_�������) & .TextMatrix(i, COL_��������)
                                    If strTmp = "E8" Or strTmp = "E9" Then
                                        strTmp = "Select 1 From ������ĿĿ¼ a where a.id=[1] and nvl(a.ִ�з���,0)  in (0,1)"
                                        Set rsTmp = zlDatabase.OpenSQLRecord(strTmp, Me.Caption, Val(.TextMatrix(i, COL_������ĿID)))
                                        If Not rsTmp.EOF Then
                                            intִ��״̬ = 0
                                        End If
                                    End If
                            End If
                        End If
                    End If
                End If
                
                If Val(.TextMatrix(i, COL_���ID)) <> 0 And .TextMatrix(i, COL_�������) = "D" Then
                    str��λ���� = .TextMatrix(i, COL_�걾��λ) & "<sTab>" & .TextMatrix(i, COL_��鷽��)
                Else
                    str��λ���� = ""
                End If
                                
                int�Ʒ�״̬ = IIF(Val(.TextMatrix(i, COL_�Ƽ�����)) = 1, -1, 0) '����Ʒѻ�δ�Ʒ�
                        
                '�ֽ�ʱ��
                If .TextMatrix(i, COL_�ֽ�ʱ��) <> "" Then
                    str�ֽ�ʱ�� = .TextMatrix(i, COL_�ֽ�ʱ��)
                Else
                    str�ֽ�ʱ�� = .Cell(flexcpData, i, COL_�ֽ�ʱ��)    '��ʼִ��ʱ��
                End If
                
                If Len(str�ֽ�ʱ��) > 4000 Then
                    Screen.MousePointer = 0
                    MsgBox "��ǰ���͵�ҽ��ʱ�䷶Χ̫��,����ִ��" & CStr(UBound(Split(str�ֽ�ʱ��, ",")) + 1) & "�Ρ�" & vbCrLf & _
                        "������֧�ֵ�������" & CStr(UBound(Split(Mid(str�ֽ�ʱ��, 1, 4000), ",")) + 1) & "��,��������������������·��ͣ�", vbInformation, gstrSysName
                    GoTo FuncEnd
                End If
                
                If strSQL <> "" Then
                    Set rsMoney = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.TextMatrix(i, COL_�շ�ϸĿID)), Val(.TextMatrix(i, COL_ִ�п���ID)), Val(NVL(mrsPati!����, 0)), mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�)
                    If Not rsMoney.EOF Then
                        int�Ʒ�״̬ = 1 '�ѼƷ�
                        Set rsClone = rsMoney.Clone
                    End If
                    
                    '����������Ŀ���ķ�����ϸ
                    bln�������� = .TextMatrix(i, COL_�������) = "F" And Val(.TextMatrix(i, COL_���ID)) <> 0
                    Do While Not rsMoney.EOF
MoneyItemBegin:
                        'ִ�п���ID
                        lngִ�п���ID = NVL(rsMoney!ִ�п���ID, 0)
                        '��ԭֵ������ȡ��Ч�ķ�ҩ��ҩƷ���������ĵ�ִ�п���
                        If InStr(",4,5,6,7", .TextMatrix(i, COL_�������)) = 0 _
                            And (rsMoney!��� = "4" And NVL(rsMoney!��������, 0) = 1 Or InStr(",5,6,7", rsMoney!���) > 0) Then
                            lng���˿���ID = Val(.TextMatrix(i, COL_���˿���ID))
                            lngִ�п���ID = Get�շ�ִ�п���ID(mlng����ID, 0, rsMoney!���, rsMoney!ID, 4, lng���˿���ID, 0, 2, lngִ�п���ID, , , 2)
                        End If
                        
                        '----------------------------------------
                        '�����շѷ�ʽ��ȷ����ǰ�շ���Ŀ�Ƿ�Ӧ�շ�
                        If rsMoney!�������� & "_" & rsMoney!ID <> str�շ���Ŀ Then
                            If Not AdviceMoneyMake(mlng����ID, mlng��ҳID, rsMoneyNow, rsMoneyDay, _
                                IIF(Val(.TextMatrix(i, COL_���ID)) <> 0, Val(.TextMatrix(i, COL_���ID)), Val(.TextMatrix(i, COL_ID))), _
                                Val(.TextMatrix(i, COL_������ĿID)), rsMoney!ID, lngִ�п���ID, .TextMatrix(i, COL_�Թܱ���), _
                                rsMoney!���, NVL(rsMoney!�շѷ�ʽ, 0), str�ֽ�ʱ��, IIF(mbytSendKind = EInBilling, 2, 1), lng���ô���, Val(.TextMatrix(i, COL_����)), _
                                Val(.TextMatrix(i, COL_ID)), lng���ͺ�, Val(rsMoney!���� & ""), rsExec, Val(.TextMatrix(i, COL_���㷽ʽ)), _
                                        .TextMatrix(i, COL_Ƶ��), Val(.TextMatrix(i, COL_����)), 1, rsMoney!��������, .TextMatrix(i, COL_�������), strCuvetteNumber, str��λ����, dblOther����, strMinDate) Then
                                '������ǰ�շ���Ŀ(���������Ŀ)
                                str�շ���Ŀ = rsMoney!�������� & "_" & rsMoney!ID
                                Do While rsMoney!�������� & "_" & rsMoney!ID = str�շ���Ŀ
                                    rsMoney.MoveNext
                                    If rsMoney.EOF Then Exit Do
                                Loop
                                If rsMoney.EOF Then Exit Do
                                GoTo MoneyItemBegin
                            End If
                        End If
                        '----------------------------------------
                        
                        '����Ƿ���Ҫ���Ѿ�����
                        If NVL(rsMoney!Ҫ������, 0) = 1 And Not rsAudit Is Nothing Then
                            rsAudit.Filter = "��ĿID=" & rsMoney!ID
                            If rsAudit.EOF Then
                                If UBound(Split(strAudit, vbCrLf)) < 10 Then
                                    If InStr(strAudit, "��" & rsMoney!����) = 0 Then
                                        strAudit = strAudit & vbCrLf & "��" & rsMoney!����
                                    End If
                                ElseIf UBound(Split(strAudit, vbCrLf)) = 10 Then
                                    strAudit = strAudit & vbCrLf & "�� ��"
                                End If
                            End If
                        End If
                    
                        If InStr(",5,6,7", rsMoney!���) > 0 Then
                            If lngҩƷ���ID = 0 Then
                                MsgBox "����ȷ��ҩƷ�������ݵ�������,���ȵ���������������ã�", vbInformation, gstrSysName
                                GoTo FuncEnd
                            End If
                        
                            If InStr(",5,6,7", .TextMatrix(i, COL_�������)) > 0 Then
                                If .TextMatrix(i, COL_�������) = "7" Then
                                    int���� = Val(.TextMatrix(i, COL_����))
                                    '��ҩҩ����λ�����ɷ��㴦��:ÿ��
                                    If Val(.TextMatrix(i, COL_�ɷ����)) = 0 Then
                                        dbl���� = Val(.TextMatrix(i, COL_����)) / NVL(rsMoney!����ϵ��, 1)
                                    Else
                                        dbl���� = IntEx(Val(.TextMatrix(i, COL_����)) / NVL(rsMoney!����ϵ��, 1) / NVL(rsMoney!סԺ��װ, 1)) * NVL(rsMoney!סԺ��װ, 1)
                                    End If
                                Else
                                    int���� = 1
                                    dbl���� = Val(.TextMatrix(i, COL_����)) * NVL(rsMoney!סԺ��װ, 1)
                                    If rsƤ�� Is Nothing Then
                                        Set rsƤ�� = GetԭҺƤ��(mlng����ID, mlng��ҳID, "")
                                    End If
                                    rsƤ��.Filter = "ҩƷID=" & Val(rsMoney!ID & "")
                                    If Not rsƤ��.EOF Then
                                        If Val(rsƤ��!��� & "") = 0 Then
                                            '���м���������
                                            dbl���� = (Val(.TextMatrix(i, COL_����)) - 1) * NVL(rsMoney!סԺ��װ, 1)
                                            rsƤ��!��� = Val(.TextMatrix(i, COL_ID))
                                            str����ҩ�� = "'" & rsƤ��!Ƥ��ҽ��ID & "," & rsƤ��!��� & "'"
                                            rsƤ��.Update
                                            If dbl���� <= 0 Then
                                                rsMoney.MoveNext
                                                If rsMoney.EOF Then Exit Do
                                                GoTo MoneyItemBegin
                                            End If
                                        End If
                                    End If
                                    
                                End If
                            Else
                                int���� = 1
                                '��ҩҩ����λ�����ɷ��㴦��:ÿ��
                                '��ҩ��ҩƷ�Ƽ�:��Ϊ����Ԥ�����ۼ�����,��˲��������㴦��
                                '�����շѶ����е�ҩƷ����Ϊ����ֻ��ȡһ�Σ�����Ϊ���ô���*��������
                                If InStr(",2,3,4,5,6,7,9,", Val("" & rsMoney!�շѷ�ʽ)) > 0 Then
                                    If dblOther���� > 0 Then
                                        dbl���� = Format(dblOther����, "0.00000")
                                    Else
                                        dbl���� = Format(lng���ô��� * NVL(rsMoney!����, 0), "0.00000")
                                    End If
                                Else
                                    dbl���� = Val(.TextMatrix(i, COL_����)) * NVL(rsMoney!����, 0)
                                End If
                            End If
                            dbl���� = Format(dbl����, "0.00000")
                            
                            If NVL(rsMoney!�Ƿ���, 0) = 1 Then
                                dbl���� = Format(CalcDrugPrice(rsMoney!ID, lngִ�п���ID, int���� * dbl����, , True, , mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�), gstrDecPrice)
                            Else
                                dbl���� = Format(NVL(rsMoney!����, 0), gstrDecPrice)
                            End If
                        ElseIf rsMoney!��� = "4" And NVL(rsMoney!��������, 0) = 1 Then
                            '�����������������
                            If lng�������ID = 0 Then
                                Screen.MousePointer = 0
                                MsgBox "����ȷ���������ϵ��ݵ�������,���ȵ���������������ã�", vbInformation, gstrSysName
                                GoTo FuncEnd
                            End If
                            
                            int���� = 1
                            If InStr(",1,2,3,4,5,6,7,9,", Val("" & rsMoney!�շѷ�ʽ)) > 0 Then
                                If dblOther���� > 0 Then
                                    dbl���� = Format(dblOther����, "0.00000")
                                Else
                                    dbl���� = Format(lng���ô��� * NVL(rsMoney!����, 0), "0.00000")
                                End If
                            Else
                                dbl���� = Format(Val(.TextMatrix(i, COL_����)) * NVL(rsMoney!����, 0), "0.00000")
                            End If
                            
                            'ȷ��ʱ�����ļ۸�
                            If NVL(rsMoney!�Ƿ���, 0) = 1 Then
                                dbl���� = Format(CalcDrugPrice(rsMoney!ID, lngִ�п���ID, dbl����, , True, , mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�), gstrDecPrice)
                            Else
                                dbl���� = Format(NVL(rsMoney!����, 0), gstrDecPrice)
                            End If
                        Else
                            '�������ڵ������������Ρ�һ��ֻ��һ��ʱ���ж�����Ҫִ�У����ն��ٴΣ����ܵ������������磺ÿ�����Σ�,��Ҫ���շѶ��յĴ���
                            int���� = 1
                            If InStr(",1,2,3,4,5,6,7,9,", Val("" & rsMoney!�շѷ�ʽ)) > 0 Then
                                If dblOther���� > 0 Then
                                    dbl���� = Format(dblOther����, "0.00000")
                                Else
                                    dbl���� = Format(lng���ô��� * NVL(rsMoney!����, 0), "0.00000")
                                End If
                            Else
                                dbl���� = Format(Val(.TextMatrix(i, COL_����)) * NVL(rsMoney!����, 0), "0.00000")
                            End If
                            dbl���� = Format(NVL(rsMoney!����, 0), gstrDecPrice)
                        End If
                        
                        '��ҩ��ҩƷ���������ĵĿ����
                        If InStr(",4,5,6,7", .TextMatrix(i, COL_�������)) = 0 _
                            And (rsMoney!��� = "4" And NVL(rsMoney!��������, 0) = 1 Or InStr(",5,6,7", rsMoney!���) > 0) Then
                            If TheStockCheck(lngִ�п���ID, rsMoney!���) <> 0 Or NVL(rsMoney!�Ƿ���, 0) = 1 Or NVL(rsMoney!����, 0) = 1 Then
                                If rsMoney!��� = "4" Then
                                    blnBool = CheckPriceStock(i, rsMoney, lngִ�п���ID, int���� * dbl����, rsTotal, bln���Ŀ����ʾ, bln����ʱ����ʾ, bln����Ĭ�Ϸ���)
                                Else
                                    blnBool = CheckPriceStock(i, rsMoney, lngִ�п���ID, int���� * dbl����, rsTotal, blnҩƷ�����ʾ, blnҩƷʱ����ʾ, blnҩƷĬ�Ϸ���)
                                End If
                                If blnBool Then
                                    Call RowSelectSame(i, COL_ѡ��, rsSQL, rsTotal, rsUpload, strҽ��IDs)
                                    '�����ǩ��ҽ��������Ƿ�һͬǩ����ҽ������һ����
                                    If Val(.TextMatrix(i, COL_ǩ��ID)) <> 0 Then
                                        If Not CheckSignSend Then
                                            GoTo FuncEnd
                                        Else
                                            Call DeleteRsExec(rsExec, Val(.TextMatrix(i, COL_ID)))
                                            GoTo NextAdvice
                                        End If
                                    Else
                                        Call DeleteRsExec(rsExec, Val(.TextMatrix(i, COL_ID)))
                                        GoTo NextAdvice
                                    End If
                                End If
                            End If
                        End If
                            
                        '���ͽ��
                        dblӦ�� = int���� * dbl���� * dbl����
                        If bln�������� Then
                            dblӦ�� = dblӦ�� * NVL(rsMoney!�����շ���, 100) / 100
                        End If
                        
                        '����Ӱ�Ӽ�
                        If gbln�Ӱ�Ӽ� And NVL(rsMoney!�Ӱ�Ӽ�, 0) = 1 Then
                            dblӦ�� = dblӦ�� * (1 + NVL(rsMoney!�Ӱ�Ӽ���, 0) / 100)
                        End If
                        
                        curӦ�� = Format(dblӦ��, gstrDec)

                        'NO,���
                        Call GetCurBillSet(strNOKey, strNO, lng�������, -1, mbytSendKind = EOutCharge)
                        rsSQL.AddNew: blnBool = False
                        If rsMoney!�������� & "_" & rsMoney!ID <> str�շ���Ŀ Then
                            lng���ø��� = lng�������
                            If rsMoney!���� = 0 Then
                                '��¼������Ϣ������϶��ڴ���ǰ
                                '��ʹ�������ۿۣ�ҲҪ��¼�������ϵ
                                If InStr(strHaveSub & ",", "," & rsMoney!�������� & ",") = 0 _
                                    And InStr(strNoneSub & ",", "," & rsMoney!�������� & ",") = 0 Then
                                    rsClone.Filter = "��������=" & rsMoney!�������� & " And ����=1"
                                    If Not rsClone.EOF Then
                                        int����� = lng�������
                                        lng����ĿID = rsMoney!ID
                                        
                                        rsSeek.AddNew
                                        rsSeek!�������� = rsMoney!��������
                                        rsSeek!�����ǩ = rsSQL.Bookmark 'Variant(Double)
                                        rsSeek!������ID = rsMoney!������ĿID
                                        rsSeek.Update
                                        strHaveSub = strHaveSub & "," & rsMoney!��������
                                        
                                        blnBool = True
                                    Else
                                        strNoneSub = strNoneSub & "," & rsMoney!��������
                                    End If
                                End If
                            End If
                        End If
                        
                        '��������ۿۺϼ�
                        If gbln��������ۿ� And (rsMoney!���� = 1 Or InStr(strHaveSub & ",", "," & rsMoney!�������� & ",") > 0) Then
                            curʵ�� = curӦ��
                            
                            '�ۼ�ҽ���ϼ��������ۿ�
                            rsSeek.Filter = "��������=" & rsMoney!��������
                            rsSeek!�ϼ� = NVL(rsSeek!�ϼ�, 0) + curʵ��
                            rsSeek.Update
                        ElseIf NVL(rsMoney!���ηѱ�, 0) = 0 Then
                            curʵ�� = Format(ActualMoney(NVL(mrsPati!�ѱ�), rsMoney!������ĿID, curӦ��, rsMoney!ID, lngִ�п���ID, int���� * dbl����, _
                                IIF(gbln�Ӱ�Ӽ� And NVL(rsMoney!�Ӱ�Ӽ�, 0) = 1, NVL(rsMoney!�Ӱ�Ӽ���, 0) / 100, 0)), gstrDec)
                        Else
                            curʵ�� = curӦ��
                        End If
                        '�����ۿ�ʱ���������ʵ�ս�������⴦��
                        If gbln��������ۿ� And blnBool Then
                            strʵ�� = Chr(0) & Chr(1) & "Begin" & curʵ�� & "End" & Chr(0) & Chr(1)
                        Else
                            strʵ�� = curʵ��
                        End If
                        
                        'ҽ������ֶ�
                        bln������Ŀ�� = False: lng���մ���ID = 0: curͳ���� = 0: str���ձ��� = "": str�������� = ""
                        If Not IsNull(mrsPati!����) Then
                            strTmp = gclsInsure.GetItemInsure(mlng����ID, rsMoney!ID, curʵ��, False, mrsPati!����, .Cell(flexcpData, i, COL_ҽ������) & "||" & int���� * dbl����)
                            If strTmp <> "" Then
                                bln������Ŀ�� = Val(Split(strTmp, ";")(0)) <> 0
                                lng���մ���ID = Val(Split(strTmp, ";")(1))
                                curͳ���� = Format(Val(Split(strTmp, ";")(2)), gstrDec)
                                str���ձ��� = CStr(Split(strTmp, ";")(3))
                                If UBound(Split(strTmp, ";")) >= 5 Then
                                    If Split(strTmp, ";")(5) <> "" Then
                                        str�������� = Split(strTmp, ";")(5)
                                    End If
                                End If
                            End If
                        End If
                        
                        '�ռ����ʱ������
                        cur�ϼ� = cur�ϼ� + curʵ��
                        If InStr(str���, rsMoney!���) = 0 Then
                            str��� = str��� & rsMoney!���
                        End If
                        
                        '����ʱ��
                        If intҩƷ���� = 3 Then
                            str����ʱ�� = strCurDate
                        ElseIf .TextMatrix(i, COL_�ֽ�ʱ��) <> "" Then
                            str����ʱ�� = "To_Date('" & Split(.TextMatrix(i, COL_�ֽ�ʱ��), ",")(0) & "','YYYY-MM-DD HH24:MI:SS')"
                        Else
                            str����ʱ�� = "To_Date('" & .Cell(flexcpData, i, COL_�ֽ�ʱ��) & "','YYYY-MM-DD HH24:MI:SS')"
                        End If
                                                
                        '��Ϊ���ڲ��Ƽ۵�ҽ������������,���Դ���ļƼ����Զ�Ϊ(0-�����Ƽ�)
                        rsSQL!���� = 4: rsSQL!��ĿID = rsMoney!ID: rsSQL!��� = i
                        rsSQL!ҽ��ID = IIF(Val(.TextMatrix(i, COL_���ID)) = 0, Val(.TextMatrix(i, COL_ID)), Val(.TextMatrix(i, COL_���ID)))
                        rsSQL!NO = strNO
                        If mbytSendKind = EOutCharge Then
                            '��δȡ��ҩ����
                            rsSQL!Sql = "ZL_���ﻮ�ۼ�¼_INSERT(" & _
                                "'" & strNO & "'," & lng������� & "," & mlng����ID & "," & ZVal(mlng��ҳID) & ",'" & _
                                IIF(mlng�������� = 1, NVL(mrsPati!�����), NVL(mrsPati!סԺ��)) & "','" & IIF(mlng�������� = 1, "", NVL(mrsPati!����)) & "','" & NVL(mrsPati!����) & "'," & _
                                "'" & NVL(mrsPati!�Ա�) & "','" & NVL(mrsPati!����) & "'," & _
                                "'" & NVL(mrsPati!�ѱ�) & "',NULL," & _
                                ZVal(.TextMatrix(i, COL_���˿���ID)) & "," & ZVal(.TextMatrix(i, COL_��������ID)) & "," & _
                                "'" & .TextMatrix(i, COL_����ҽ��) & "'," & IIF(rsMoney!���� = 1, ZVal(int�����), "NULL") & "," & _
                                rsMoney!ID & ",'" & rsMoney!��� & "','" & NVL(rsMoney!���㵥λ) & "',NULL," & _
                                int���� & "," & dbl���� & "," & IIF(bln��������, 1, 0) & "," & ZVal(lngִ�п���ID) & "," & _
                                IIF(lng���ø��� = lng�������, "NULL", lng���ø���) & "," & rsMoney!������ĿID & "," & _
                                "'" & NVL(rsMoney!�վݷ�Ŀ) & "'," & dbl���� & "," & curӦ�� & "," & strʵ�� & "," & _
                                str����ʱ�� & "," & strCurDate & "," & _
                                "'ҽ������','" & UserInfo.���� & "'," & _
                                "'" & .TextMatrix(i, col_ҽ������) & "'," & Val(.TextMatrix(i, COL_ID)) & ",'" & .TextMatrix(i, COL_Ƶ��) & "'," & _
                                ZVal(.TextMatrix(i, COL_����)) & ",'" & .TextMatrix(i, COL_�÷�) & "',1," & _
                                IIF(intҩƷ���� <> 0, intҩƷ����, Val(.TextMatrix(i, COL_�Ƽ�����))) & "," & IIF(mlng�������� = 1, "1", "2") & "," & _
                                "'" & str���ձ��� & "','" & str�������� & "'," & IIF(bln������Ŀ��, 1, 0) & "," & ZVal(lng���մ���ID) & ",NULL,0," & ZVal(Val(.TextMatrix(i, COL_��鷽��))) & "," & ZVal(mlng��ҳID) & "," & ZVal(mlng���˲���ID) & ")"
                                
                        ElseIf mbytSendKind = EOutBilling Then
                            '�Ƿ񻮼۷���
                            If InStr(",5,6,7,", .TextMatrix(i, COL_�������)) > 0 Then
                                int���� = IIF(InStr(gstr���﷢�ͻ��۵�, "5") > 0, 1, 0)
                            Else
                                int���� = IIF(InStr(gstr���﷢�ͻ��۵�, .TextMatrix(i, COL_�������)) > 0, 1, 0)
                            End If
                            If int���� = 0 Then int���� = IIF(NVL(rsMoney!����ȷ��, 0) = 1, 1, 0)
                            
                            If int���� = 0 Or intִ��״̬ = 1 Then
                                bln���� = False
                                If gdblԤ��������鿨 <> 0 Then cur���ʺϼ� = cur���ʺϼ� + curʵ��
                            End If
                            
                            '�Ǽ�ʱ��
                            If int���� = 1 Then '��ǻ��۵�ʱ�������ֿ�
                                str�Ǽ�ʱ�� = "To_Date('" & Format(DateAdd("s", 1, curDate), "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                            Else
                                str�Ǽ�ʱ�� = strCurDate
                            End If
                        
                            rsSQL!Sql = "zl_������ʼ�¼_INSERT(" & _
                                "'" & strNO & "'," & lng������� & "," & mlng����ID & ",'" & _
                                IIF(mlng�������� = 1, NVL(mrsPati!�����), NVL(mrsPati!סԺ��)) & "','" & mrsPati!���� & "'," & _
                                "'" & mrsPati!�Ա� & "','" & mrsPati!���� & "'," & "'" & mrsPati!�ѱ� & "',0," & Val(.Cell(flexcpData, i, COL_Ӥ��)) & "," & _
                                ZVal(.TextMatrix(i, COL_���˿���ID)) & "," & ZVal(.TextMatrix(i, COL_��������ID)) & "," & _
                                "'" & .TextMatrix(i, COL_����ҽ��) & "'," & IIF(rsMoney!���� = 1, ZVal(int�����), "NULL") & "," & _
                                rsMoney!ID & ",'" & rsMoney!��� & "','" & rsMoney!���㵥λ & "'," & _
                                int���� & "," & dbl���� & "," & IIF(bln��������, 1, 0) & "," & ZVal(lngִ�п���ID) & "," & _
                                IIF(lng���ø��� = lng�������, "NULL", lng���ø���) & "," & rsMoney!������ĿID & "," & _
                                "'" & rsMoney!�վݷ�Ŀ & "'," & dbl���� & "," & curӦ�� & "," & strʵ�� & "," & _
                                str����ʱ�� & "," & str�Ǽ�ʱ�� & "," & _
                                "'ҽ������'," & int���� & ",'" & UserInfo.��� & "','" & UserInfo.���� & "'," & _
                                "Null,'" & .TextMatrix(i, col_ҽ������) & "'," & Val(.TextMatrix(i, COL_ID)) & ",'" & .TextMatrix(i, COL_Ƶ��) & "'," & _
                                ZVal(.TextMatrix(i, COL_����)) & ",'" & .TextMatrix(i, COL_�÷�) & "',1," & _
                                IIF(intҩƷ���� <> 0, intҩƷ����, Val(.TextMatrix(i, COL_�Ƽ�����))) & ",1,Null,0," & ZVal(Val(.TextMatrix(i, COL_��鷽��))) & "," & ZVal(mlng��ҳID) & "," & ZVal(mlng���˲���ID) & ")"
                            '�����־Ҫ��1����������ʷ��ã���2����סԺ���ʷ�����
                        Else
                            '�Ƿ񻮼۷���
                            strTmp = mlng��ҩ����ID
                            If InStr(",5,6,7,", .TextMatrix(i, COL_�������)) > 0 Then
                                int���� = IIF(InStr(gstrסԺ���ͻ��۵�, "5") > 0, 1, 0)
                                
                                j = .FindRow(CStr(.TextMatrix(i, COL_���ID)), i + 1, COL_ID)
                                If Val(.TextMatrix(j, COL_ִ�п���ID)) <> 0 Then strTmp = Val(.TextMatrix(j, COL_ִ�п���ID))
                            Else
                                int���� = IIF(InStr(gstrסԺ���ͻ��۵�, .TextMatrix(i, COL_�������)) > 0, 1, 0)
                            End If
                            If int���� = 0 Then int���� = IIF(NVL(rsMoney!����ȷ��, 0) = 1, 1, 0)
                            
                            If int���� = 0 Or intִ��״̬ = 1 Then
                                bln���� = False
                                cur���ʺϼ� = cur���ʺϼ� + curʵ��
                            End If
                            
                            '�Ǽ�ʱ��
                            If int���� = 1 Then '��ǻ��۵�ʱ�������ֿ�
                                str�Ǽ�ʱ�� = "To_Date('" & Format(DateAdd("s", 1, curDate), "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                            Else
                                str�Ǽ�ʱ�� = strCurDate
                            End If
                            
                            '�ռ�ҽ���ϴ����ݺ�:mrsBill�еĲ�һ�������˷���
                            If int���� = 0 Then
                                rsUpload.Filter = "NO='" & strNO & "'"
                                If rsUpload.EOF Then
                                    rsUpload.AddNew
                                    rsUpload!ҽ��ID = IIF(Val(.TextMatrix(i, COL_���ID)) = 0, Val(.TextMatrix(i, COL_ID)), Val(.TextMatrix(i, COL_���ID)))
                                    rsUpload!NO = strNO
                                    rsUpload.Update
                                End If
                            End If
                            
                            rsSQL!Sql = "ZL_סԺ���ʼ�¼_Insert(" & _
                                "'" & strNO & "'," & lng������� & "," & mlng����ID & "," & ZVal(mlng��ҳID) & "," & _
                                IIF(IsNull(mrsPati!סԺ��), "NULL", "'" & mrsPati!סԺ�� & "'") & ",'" & NVL(mrsPati!����) & "'," & _
                                "'" & NVL(mrsPati!�Ա�) & "','" & NVL(mrsPati!����) & "'," & _
                                "'" & NVL(mrsPati!����) & "','" & NVL(mrsPati!�ѱ�) & "'," & _
                                ZVal(mlng���˲���ID) & "," & ZVal(.TextMatrix(i, COL_���˿���ID)) & ",0," & _
                                Val(.Cell(flexcpData, i, COL_Ӥ��)) & "," & _
                                ZVal(.TextMatrix(i, COL_��������ID)) & ",'" & .TextMatrix(i, COL_����ҽ��) & "'," & _
                                IIF(rsMoney!���� = 1, ZVal(int�����), "NULL") & "," & rsMoney!ID & "," & _
                                "'" & rsMoney!��� & "','" & NVL(rsMoney!���㵥λ) & "'," & _
                                IIF(bln������Ŀ��, 1, 0) & "," & ZVal(lng���մ���ID) & ",'" & str���ձ��� & "'," & _
                                int���� & "," & dbl���� & ",NULL," & ZVal(lngִ�п���ID) & "," & _
                                IIF(lng���ø��� = lng�������, "NULL", lng���ø���) & "," & rsMoney!������ĿID & "," & _
                                "'" & NVL(rsMoney!�վݷ�Ŀ) & "'," & dbl���� & "," & curӦ�� & "," & strʵ�� & "," & _
                                curͳ���� & "," & str����ʱ�� & "," & str�Ǽ�ʱ�� & "," & _
                                "'ҽ������'," & int���� & ",'" & UserInfo.��� & "','" & UserInfo.���� & "',0," & _
                                IIF(rsMoney!��� = "4", lng�������ID, lngҩƷ���ID) & "," & _
                                "NULL,'" & .TextMatrix(i, col_ҽ������) & "',NULL," & Val(.TextMatrix(i, COL_ID)) & "," & _
                                "'" & .TextMatrix(i, COL_Ƶ��) & "'," & ZVal(.TextMatrix(i, COL_����)) & "," & _
                                "'" & .TextMatrix(i, COL_�÷�) & "',1," & _
                                IIF(intҩƷ���� <> 0, intҩƷ����, Val(.TextMatrix(i, COL_�Ƽ�����))) & "," & _
                                "Null,'" & str�������� & "',Null," & strTmp & ",NULL,-1,0," & ZVal(Val(.TextMatrix(i, COL_��鷽��))) & ")"
                        End If
                        rsSQL.Update
                        
                        '��¼�Զ����ϵ�SQL
                        If ((gbytסԺ�Զ����� = 1 Or gbytסԺ�Զ����� = 2 And lngִ�п���ID = Val(.TextMatrix(i, COL_��������ID))) And mbytSendKind = EInBilling Or gbln�����Զ����� And mbytSendKind = EOutBilling) _
                            And int���� = 0 And lngִ�п���ID <> 0 And rsMoney!��� = "4" And NVL(rsMoney!��������, 0) = 1 Then
                            If InStr(str�Զ����� & ";", ";" & strNO & "," & lngִ�п���ID & ";") = 0 Then
                                rsSQL.AddNew
                                rsSQL!���� = 6
                                rsSQL!ҽ��ID = IIF(Val(.TextMatrix(i, COL_���ID)) = 0, Val(.TextMatrix(i, COL_ID)), Val(.TextMatrix(i, COL_���ID)))
                                rsSQL!��ĿID = 0
                                rsSQL!��� = i
                                rsSQL!NO = strNO
                                rsSQL!Sql = "zl_�����շ���¼_��������(" & lngִ�п���ID & ",25,'" & strNO & "','" & UserInfo.���� & "','" & UserInfo.���� & "','" & UserInfo.���� & "',1,Sysdate)"
                                rsSQL.Update
                                str�Զ����� = str�Զ����� & ";" & strNO & "," & lngִ�п���ID
                            End If
                        End If
                        
                        'ҽ���ܿ�ʵʱ��⣺���ɷ�����Ŀ��¼��,���շ�ϸĿ����
                        If Not IsNull(mrsPati!����) And blnʵʱ��� Then
                            rsItems.Filter = "�շ�ϸĿID=" & rsMoney!ID
                            If rsItems.EOF Then
                                '�����շ���Ŀ��Ӧ��ԭʼ��Ϣ
                                rsItems.AddNew
                                rsItems!����ID = mlng����ID
                                rsItems!��ҳID = mlng��ҳID
                                rsItems!ҽ��ID = Val(.TextMatrix(i, COL_ID))
                                rsItems!�շ���� = rsMoney!���
                                rsItems!�շ�ϸĿID = rsMoney!ID
                                rsItems!������ = .TextMatrix(i, COL_����ҽ��)
                                rsItems!�������� = CStr(sys.RowValue("���ű�", Val(.TextMatrix(i, COL_��������ID)), "����"))
                                
                                rsItems!���� = int���� * dbl����
                                rsItems!���� = dbl����
                            Else
                                '����һ��ҽ��(������Ŀ)���շѶ��ղ������ظ����շ�ϸĿ
                                '������ͬһ�շ���Ŀ�Ĳ�ͬ������Ŀ��¼��ͬ
                                If rsMoney!�������� & "_" & rsMoney!ID <> str�շ���Ŀ Then
                                    rsItems!���� = NVL(rsItems!����, 0) + int���� * dbl����
                                End If
                                '���ۣ�ͬһ�շ���Ŀ�Ĳ�ͬ������Ŀ�ۼ�
                                If Val(.TextMatrix(i, COL_ID)) = rsItems!ҽ��ID Then
                                    rsItems!���� = NVL(rsItems!����, 0) + dbl����
                                End If
                            End If
                            rsItems!ʵ�ս�� = NVL(rsItems!ʵ�ս��, 0) + curʵ��
                            rsItems.Update
                        End If
                            
                        str�շ���Ŀ = rsMoney!�������� & "_" & rsMoney!ID
                        rsMoney.MoveNext
                    Loop
                End If
                
                '��ҽ�������л����ۿ۴���
                If gbln��������ۿ� And strHaveSub <> "" Then
                    rsSeek.Filter = 0
                    Do While Not rsSeek.EOF
                        rsSQL.Bookmark = rsSeek!�����ǩ
                        curʵ�� = Format(ActualMoney(NVL(mrsPati!�ѱ�), rsSeek!������ID, rsSeek!�ϼ�), gstrDec)
                        curʵ�� = curʵ�� - rsSeek!�ϼ� '���۲��
                        
                        'ҽ���ܿ�ʵʱ��⣺������Ŀ����滻
                        If Not IsNull(mrsPati!����) And blnʵʱ��� Then
                            rsItems.Filter = "�շ�ϸĿID=" & lng����ĿID
                            If Not rsItems.EOF Then
                                rsItems!ʵ�ս�� = NVL(rsItems!ʵ�ս��, 0) + curʵ��
                                rsItems.Update
                            End If
                        End If
                        
                        '����SQL�����滻
                        curʵ�� = Getʵ�ս��(rsSQL!Sql) + curʵ��
                        rsSQL!Sql = Setʵ�ս��(rsSQL!Sql, curʵ��)
                        rsSQL.Update
                    
                        rsSeek.MoveNext
                    Loop
                End If
                
                'ѡ��Ҫ���͵�ҽ���Զ�����У��(ʵ�ʿ�����Ϊ����������)
                If Val(.TextMatrix(i, COL_ҽ��״̬)) = 1 And Val(.TextMatrix(i, COL_���ID)) = 0 Then
                    rsSQL.AddNew
                    rsSQL!���� = 3: rsSQL!��ĿID = 0: rsSQL!��� = i
                    rsSQL!ҽ��ID = Val(.TextMatrix(i, COL_ID))
                    rsSQL!Sql = "ZL_����ҽ����¼_У��(" & Val(.TextMatrix(i, COL_ID)) & ",3," & _
                        "To_Date('" & Format(.TextMatrix(i, COL_����ʱ��), "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI')," & _
                        "NULL,NULL,'" & UserInfo.��� & "','" & UserInfo.���� & "')"
                End If
                
                
                '����ҽ�����ͼ�¼
                '-----------------------------------------------------------------------------------------
                If Val(.TextMatrix(i, COL_ִ������ID)) <> 0 Then '����������(��ҩ;�����䷽�巨���÷�,�ɼ���������Ѫ;������Ϊ)
                    '�����˳�Ժ,תԺ,����ҽ��
                    If .TextMatrix(i, COL_�������) = "Z" _
                        And InStr(",5,6,11,", Val(.TextMatrix(i, COL_��������))) > 0 Then
                        mblnRefresh = True
                    End If
                    
                    'һ��Ҫ��������NO
                    Call GetCurBillSet(strNOKey, strNO, -1, lng�������, mbytSendKind = EOutCharge)
                                                            
                    '�Ƿ�һ��ҽ���ĵ�һҽ����:ҩ�Ƶĵ�һҩƷ��Ϊ��һҽ����
                    blnFirst = False
                    If InStr(",5,6,7,", .TextMatrix(i, COL_�������)) > 0 Then
                        If Val(.TextMatrix(i, COL_���ID)) <> Val(.TextMatrix(i - 1, COL_���ID)) Then
                            blnFirst = True
                        End If
                    ElseIf .TextMatrix(i, COL_�������) = "C" And Val(.TextMatrix(i, COL_���ID)) <> 0 Then
                        If Val(.TextMatrix(i, COL_���ID)) <> Val(.TextMatrix(i - 1, COL_���ID)) Then
                            blnFirst = True '��������еĵ�һ������
                        End If
                    ElseIf InStr(",1,2,3,4,5,", CLng(.Cell(flexcpData, i, COL_ID))) = 0 Then '�ſ���ҩ;������ҩ�巨����ҩ�÷����ɼ���������Ѫ;��
                        If Val(.TextMatrix(i, COL_���ID)) = 0 Then
                            blnFirst = True
                        End If
                    End If
                                        
                    '��������:ҩƷΪ������λ������,����Ϊ����
                    If .TextMatrix(i, COL_�������) = "7" Then
                        dbl�������� = Val(.TextMatrix(i, COL_����)) * Val(.TextMatrix(i, COL_����))
                    ElseIf InStr(",5,6,", .TextMatrix(i, COL_�������)) > 0 Then
                        dbl�������� = Val(.TextMatrix(i, COL_����)) * Val(.TextMatrix(i, COL_סԺ��װ)) * Val(.TextMatrix(i, COL_����ϵ��))
                    Else
                        dbl�������� = Val(.TextMatrix(i, COL_����))
                    End If
                    dbl�������� = Format(dbl��������, "0.00000")
                                                            
                    '��ĩʱ��(�����á�str�ֽ�ʱ�䡱�жϣ���Ϊһ����������¼�����״�ʱ��)
                    If .TextMatrix(i, COL_�ֽ�ʱ��) <> "" Or mlngRefModld = 1 Then
                        str�״�ʱ�� = "To_Date('" & Split(str�ֽ�ʱ��, ",")(0) & "','YYYY-MM-DD HH24:MI:SS')"
                        strĩ��ʱ�� = "To_Date('" & Split(str�ֽ�ʱ��, ",")(UBound(Split(str�ֽ�ʱ��, ","))) & "','YYYY-MM-DD HH24:MI:SS')"
                    Else
                        '�޷��ֽ��Ϊ"һ����"��������Ϊ��ʼִ��ʱ�䣨74366��
                        str�״�ʱ�� = "To_Date('" & .TextMatrix(i, COL_��ʼʱ��) & "','YYYY-MM-DD HH24:MI:SS')"
                        strĩ��ʱ�� = "To_Date('" & .TextMatrix(i, COL_��ʼʱ��) & "','YYYY-MM-DD HH24:MI:SS')"
                    End If
                    If str�ֽ�ʱ�� = "" Then str�ֽ�ʱ�� = .TextMatrix(i, COL_��ʼʱ��)
                        
                    '��ҩ��
                    str��ҩ�� = ""
                    If mbln��ҩ�� And InStr(",5,6,7,", .TextMatrix(i, COL_�������)) > 0 Then
                        If mstr��ҩ�� = "" Then mstr��ҩ�� = Get��ҩ��
                        str��ҩ�� = mstr��ҩ��
                    End If
                                       
                    If Not gbln�������������� Then strCuvetteNumber = ""
                    rsSQL.AddNew
                    rsSQL!���� = 5: rsSQL!��ĿID = 0: rsSQL!��� = i
                    rsSQL!ҽ��ID = IIF(Val(.TextMatrix(i, COL_���ID)) = 0, Val(.TextMatrix(i, COL_ID)), Val(.TextMatrix(i, COL_���ID)))
                    rsSQL!NO = strNO
                    rsSQL!Sql = "ZL_����ҽ������_Insert(" & _
                        Val(.TextMatrix(i, COL_ID)) & "," & lng���ͺ� & "," & IIF(mbytSendKind = EOutCharge, 1, 2) & ",'" & strNO & "'," & _
                        lng������� & "," & ZVal(dbl��������) & "," & str�״�ʱ�� & "," & strĩ��ʱ�� & "," & strCurDate & "," & _
                        intִ��״̬ & "," & ZVal(.TextMatrix(i, COL_ִ�п���ID)) & "," & int�Ʒ�״̬ & "," & _
                        IIF(blnFirst, 1, 0) & ",'" & strCuvetteNumber & "','" & UserInfo.��� & "'," & _
                        "'" & UserInfo.���� & "','" & str��ҩ�� & "'," & IIF(mbytSendKind = EOutBilling, 1, "Null") & ",'" & str�ֽ�ʱ�� & "'," & IIF(InStr(str����ҩ��, "," & Val(.TextMatrix(i, COL_ID)) & "'") > 0, str����ҩ��, "Null") & ")"
                    rsSQL.Update
                    
                    str����ҩ�� = "''"
                    
                    If gblnѪ��ϵͳ And .TextMatrix(i, COL_�������) = "K" Then
                        rsSQL.AddNew
                        rsSQL!���� = 9 'Ѫ����Ѫ����
                        rsSQL!��ĿID = 0
                        rsSQL!��� = 0
                        rsSQL!Sql = "Zl_ѪҺ��Ѫ����_Insert(" & Val(.TextMatrix(i, COL_ID)) & ")"
                        rsSQL.Update
                    End If
                    
                    'ҽ��ִ�мƼ�
                    If rsExec.RecordCount > 0 Then
                        rsExec.Filter = "ҽ��ID=" & Val(.TextMatrix(i, COL_ID)) & " And ���ͺ�=" & lng���ͺ�
                        If rsExec.RecordCount > 0 Then rsExec.MoveFirst
                        Do While Not rsExec.EOF
                            rsSQL.AddNew
                            rsSQL!���� = 8
                            rsSQL!��ĿID = 0
                            rsSQL!��� = 0
                            rsSQL!ҽ��ID = IIF(Val(.TextMatrix(i, COL_���ID)) = 0, Val(.TextMatrix(i, COL_ID)), Val(.TextMatrix(i, COL_���ID)))
                            rsSQL!Sql = "Zl_ҽ��ִ�мƼ�_Insert(" & rsExec!ҽ��ID & "," & rsExec!���ͺ� & ",To_date('" & _
                            rsExec!Ҫ��ʱ�� & "','yyyy-MM-dd HH24:mi:ss')," & ZVal(Val(rsExec!�շ�ϸĿID & "")) & "," & rsExec!���� & "," & rsExec!�������� & ")"
                            rsSQL.Update
                            rsExec.MoveNext
                        Loop
                        rsExec.Filter = 0
                    End If
                    
                    'Ҫ���͵���δǩ�����¿�ҽ��ID(��ID,һ���еĶ���Ҳ�ᱻǩ��)
                    If Val(.TextMatrix(i, COL_ǩ��ID)) = 0 And Val(.TextMatrix(i, COL_ҽ��״̬)) = 1 Then
                        If Val(.TextMatrix(i, COL_���ID)) <> 0 Then
                            lng��ID = Val(.TextMatrix(i, COL_���ID))
                        Else
                            lng��ID = Val(.TextMatrix(i, COL_ID))
                        End If
                        If InStr(strҽ��IDs & ",", "," & lng��ID & ",") = 0 Then
                            strҽ��IDs = strҽ��IDs & "," & lng��ID
                        End If
                    End If
                End If
                
                '������ҩ�䷽��
                If .Cell(flexcpData, i, COL_ID) = 3 Then '��ҩ�÷�
                    int�䷽�� = int�䷽�� + 1
                End If
            End If
NextAdvice:
            '----------------------------------------
            Progress = (i - .FixedRows + 1) / (.Rows - .FixedRows) * 100
            txtPer.Text = CInt(psb.value) & "%"
            txtPer.Refresh
        Next
                
        '��ʾδ�����Ŀ
        If strAudit <> "" Then
            MsgBox "����""" & mrsPati!���� & """���·�����Ŀ��û�о�����������Ӧ��ҽ�����ܷ��ͣ�" & vbCrLf & strAudit, vbInformation, gstrSysName
            GoTo errH
        End If
        
        '�Զ����е���ǩ��(δǩ������)
        '-----------------------------------------------------------------------------------------
        If Not gobjESign Is Nothing And CheckSign(IIF(mlngҽ������ID <> 0, 3, 1), 0, mlngҽ������ID, mlng���˿���id, 2, , gobjESign) And strҽ��IDs <> "" Then
            strҽ��IDs = Mid(strҽ��IDs, 2) '��������ID,����Ϊ��ϸ��ID
            intRule = ReadAdviceSignSource(1, mlng����ID, mlng��ҳID, strҽ��IDs, 0, False, strSource, mstrǰ��IDs)
            If intRule = 0 Then GoTo FuncEnd
            If strSource = "" Then
                Screen.MousePointer = 0
                MsgBox "���ܶ�ȡҪǩ����ҽ��Դ�ġ�", vbInformation, gstrSysName
                GoTo FuncEnd
            End If
            
            strSign = gobjESign.Signature(strSource, gstrDBUser, lng֤��ID, strTimeStamp, Nothing, strTimeStampCode)
            If strSign = "" Then GoTo FuncEnd
            If strTimeStamp <> "" Then
                strTimeStamp = "To_Date('" & strTimeStamp & "','YYYY-MM-DD HH24:MI:SS')"
            Else
                strTimeStamp = "NULL"
            End If
            lngǩ��id = zlDatabase.GetNextID("ҽ��ǩ����¼")
            rsSQL.AddNew
            rsSQL!���� = 2: rsSQL!ҽ��ID = 0: rsSQL!��ĿID = 0: rsSQL!��� = 0
            rsSQL!Sql = "zl_ҽ��ǩ����¼_Insert(" & lngǩ��id & ",1," & intRule & ",'" & Replace(strSign, "'", "''") & "'," & lng֤��ID & ",'" & strҽ��IDs & "'," & strTimeStamp & ",'" & UserInfo.���� & "','" & strTimeStampCode & "')"
            rsSQL.Update
        End If
        
        'ҽ���ܿ�ʵʱ���
        If Not IsNull(mrsPati!����) And blnʵʱ��� Then
            rsItems.Filter = 0
            If Not rsItems.EOF Then
                If Not gclsInsure.CheckItem(mrsPati!����, 1, 2, rsItems) Then GoTo FuncEnd
            End If
        End If
        
        '��ͨ����ҳ�涼��������ҩ��¼��64615��
        
        str����ҽ��IDs = Mid(str����ҽ��IDs, 2)
        
        '�ύ��������
        '-----------------------------------------------------------------------------------------
        If Not CompletePatiSend(rsSQL, rsUpload, cur�ϼ�, cur���ʺϼ�, str���, bln����, blnTran, lng���ͺ�, str����ҽ��IDs) Then GoTo errH
    End With
    SendAdvice = lng���ͺ�
    '������ҽӿ�
    If CreatePlugInOK(pסԺҽ������) Then
        On Error Resume Next
        Call gobjPlugIn.AdviceSendEnd(glngSys, pסԺҽ������, lng���ͺ� & "")
        Call zlPlugInErrH(err, "AdviceSendEnd")
        On Error GoTo 0
    End If
    Call Make��ִ����Ϣ(Format(curDate, "yyyy-MM-dd HH:mm:ss"))
FuncEnd:
    'ɾ�������ѳɹ����͵���
    Call DeleteSendRow: Call ShowSendTotal
    Progress = 0: Screen.MousePointer = 0
    Exit Function
errH:
    Screen.MousePointer = 0
    If blnTran Then
        gcnOracle.RollbackTrans
    End If
    If err.Number <> 0 Then
        If ErrCenter() = 1 Then
            Resume
        End If
        Call SaveErrLog
    End If
    Call DeleteSendRow: Call ShowSendTotal
    Progress = 0
End Function

Private Function CompletePatiSend(rsSQL As ADODB.Recordset, _
    rsUpload As ADODB.Recordset, ByVal cur�ϼ� As Currency, ByVal cur���ʺϼ� As Currency, ByVal str��� As String, _
    ByVal bln���� As Boolean, blnTran As Boolean, ByVal lng���ͺ� As Long, ByVal str����ҽ��IDs As String) As Boolean
'���ܣ��ύһ�����˵�ҽ����������,����֮ǰ������ʱ���
'������
'      bln����=�Ƿ�ȫ�����ö��ǻ���ģʽ�����ڱ��������⴦��
'      rsSQL=��������Ҫִ�е�SQL
'      rsUpload=����ҽ���ϴ��ļ��ʵ��ݺ�
'      cur�ϼ�=���˱���Ҫ����ҽ���ļ��ʽ��ϼ�,���ڼ��ʱ���
'      cur���ʺϼ�=���˱���Ҫ����ҽ���ļ��ʽ��ϼƣ���������ִ�к��Զ���˵Ļ��۷��ã������������۷���
'      str���=���˱��η��ͼ��ʷ��õ��շ����,���ڼ��ʱ���
'      lng���ͺ�=���η��͵����ؼ���
'      str����ҽ��IDs=һ��ͨ�����ҽ��ID��
'˵�����������,���ڵ��ú����д���,blnTran�����Ƿ�����������
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, intR As Integer, lng��ID As Long, strҽ��IDs As String
    Dim cur���� As Currency, i As Long, j As Long
    Dim arrNOs() As String, strMsg As String, cur��� As Currency
    Dim strAllmsg As String, strDiag As String, strAdviceInfo As String
    Dim arrSQL As Variant, arrAdviceID As Variant
    Dim strErr As String
    Dim str��ǰ���� As String
    Dim blnClearPatiCache As Boolean
    Dim blnPlugIn As Boolean, int���� As Integer
    Dim rsAdviceRis As ADODB.Recordset
    Dim strAdvices��Ѫ As String
    Dim var��Ѫ As Variant
    
    int���� = IIF("" = mstrǰ��IDs, 0, 2)
'    ������ҽӿڷ���ǰ���ҽ������
    If CreatePlugInOK(pסԺҽ������, int����) Then
        blnPlugIn = True
        On Error Resume Next
        blnPlugIn = gobjPlugIn.AdviceCheckSendFee(glngSys, pסԺҽ������, mlng����ID, mlng��ҳID, cur�ϼ�, int����)
        If Not blnPlugIn And err.Number <> 0 Then blnPlugIn = True
        Call zlPlugInErrH(err, "AdviceCheckSendFee")
        err.Clear: On Error GoTo 0
        If Not blnPlugIn Then
            Exit Function
        End If
    End If
    
    '���˷��ñ���
    blnClearPatiCache = True
    If Not (mbytSendKind = EOutCharge) And cur�ϼ� > 0 Then
        If InitObjPublicExpense Then
            For i = 1 To Len(str���)
                Call gobjPublicExpense.zlBillingWarn.zlBillingWarnCheck(Me, IIF(mbytSendKind = EInBilling, 1, 0), IIF(bln����, 1, 0), mlng����ID, IIF(mlng�������� = 1, 0, mlng��ҳID), mlng���˲���ID, Mid(str���, i, 1), IIF(gbln�����������۷���, cur�ϼ�, cur���ʺϼ�), InStr(";" & GetInsidePrivs(pסԺҽ���´�) & ";", ";Ƿ��ǿ�Ƽ���;") > 0, False, blnClearPatiCache, intR, , , , True)
                blnClearPatiCache = False
                If InStr(",2,3,", intR) > 0 Then Exit Function
            Next
        End If
    End If
    
    If mbytSendKind = EOutBilling And gdblԤ��������鿨 <> 0 And cur���ʺϼ� > 0 Then
        If Not zlDatabase.PatiIdentify(Me, glngSys, mlng����ID, cur���ʺϼ�, , , , IIF(-1 * gdblԤ��������鿨 >= Val(cur���ʺϼ�), False, True), , , (gdblԤ��������鿨 <> 0), (2 = gdblԤ��������鿨)) Then Exit Function
    End If
    
    Call InitObjLis(pסԺҽ��վ)
    '����LIS����ӿ�
    If Not gobjLIS Is Nothing Then
        strAdviceInfo = Get����ҽ����Ϣ
        If strAdviceInfo <> "" Then
            Set rsTmp = Get������ϼ�¼(mlng����ID, mlng��ҳID, "2")
            If rsTmp.RecordCount > 0 Then strDiag = rsTmp!�������
        End If
    End If
    
    If gblnѪ��ϵͳ Then
        If InitObjBlood(True) Then
            strAdvices��Ѫ = Get��Ѫҽ����Ϣ
            If strAdvices��Ѫ <> "" Then
                var��Ѫ = Split(strAdvices��Ѫ, ",")
            End If
        End If
    End If
    
    Call ReplaceTrueNO(rsSQL, rsUpload)
    
    'ִ��˳��:1-�Ƽ�,2-ǩ��,3-У��,4-����,5-����,6-����
    '1.����д����,��Ϊ����ʱ���ܴ������
    '2.�Է��ü�¼���շ�ϸĿID�������
    rsSQL.Filter = 0 '�ϲ㺯������ʹ�ù�,��ʹû�ù�ҲMoveFirst
    rsSQL.Sort = "����,��ĿID,���"
    rsUpload.Filter = 0 '�ϲ㺯������ʹ�ù�,��ʹû�ù�ҲMoveFirst

    With vsAdvice
        For i = .FixedRows To .Rows - 1
            If .Cell(flexcpData, i, COL_ѡ��) = 0 And Not .Cell(flexcpPicture, i, COL_ѡ��) Is Nothing Then
                If .TextMatrix(i, COL_�������) = "Z" And (Val(.TextMatrix(i, COL_��������)) = 9 Or Val(.TextMatrix(i, COL_��������)) = 10) Then
                    str��ǰ���� = Get���˵�ǰ����(mlng����ID, mlng��ҳID)
                    Exit For
                End If
            End If
        Next
    End With
 
    gcnOracle.BeginTrans: blnTran = True
    Do While Not rsSQL.EOF
        Call zlDatabase.ExecuteProcedure(rsSQL!Sql, Me.Caption)
        rsSQL.MoveNext
    Loop
    
    '����LIS����ӿ�
    If strAdviceInfo <> "" Then
        If gobjLIS.SendLisApplicationForm(strAdviceInfo, strDiag) = False Then
            gcnOracle.RollbackTrans: blnTran = False
            Screen.MousePointer = 0
            Call Del��������
            MsgBox "����ӿڵ���ʧ�ܣ����ܷ��ͼ���ҽ����", vbInformation, gstrSysName
            Exit Function
        End If
    End If
            
    'ҽ�������ϴ�
    If Not IsNull(mrsPati!����) Then
        If gclsInsure.GetCapability(supportҽ���ϴ�, mlng����ID, mrsPati!����) And Not gclsInsure.GetCapability(support������ɺ��ϴ�, mlng����ID, mrsPati!����) Then
            Do While Not rsUpload.EOF
                strMsg = "" '��Ϊ����һ��NO�ڿ϶�Ϊһ�����˵�,��������˲������Բ���
                If Not gclsInsure.TranChargeDetail(2, rsUpload!NO, 2, 1, strMsg, , mrsPati!����, rsUpload.RecordCount & "|" & rsUpload.AbsolutePosition) Then
                    'δ�ύǰ�ϴ�ʧ����ع�����ֹ����
                    gcnOracle.RollbackTrans: blnTran = False
                    Screen.MousePointer = 0
                    If strMsg <> "" Then
                        MsgBox strMsg, vbInformation, gstrSysName 'ÿ����ʾ
                    Else
                        MsgBox mrsPati!���� & "�ķ����ϴ�ʧ�ܣ����Ͳ���������ֹ��", vbExclamation, gstrSysName
                    End If
                    Exit Function
                Else
                    If strMsg <> "" Then strAllmsg = strAllmsg & rsUpload!NO & ":" & strMsg & vbCrLf
                End If
                rsUpload.MoveNext
            Loop
        End If
        
        
        'ҽ�������ϴ��ӿ�(������������)
        If gclsInsure.GetCapability(support�ϴ�סԺ����, mlng����ID, mrsPati!����) Then
            If Not gclsInsure.TranElecDossier(2, mlng����ID, mlng��ҳID, mrsPati!����) Then Exit Function
        End If
    End If
    If strAdvices��Ѫ <> "" Then
        For i = 0 To UBound(var��Ѫ)
            If gobjPublicBlood.AdviceOperation(pסԺҽ���´�, Val(var��Ѫ(i)), 5, False, strErr) = False Then
                gcnOracle.RollbackTrans: blnTran = False
                Screen.MousePointer = 0
                MsgBox "Ѫ��ϵͳ�ӿڵ���ʧ�ܣ�" & strErr, vbInformation, gstrSysName
                Exit Function
            End If
        Next
    End If
    gcnOracle.CommitTrans: blnTran = False
    Screen.MousePointer = 0
    If strAllmsg <> "" Then
        MsgBox strAllmsg, vbInformation, gstrSysName
    End If
    
    'һ��ͨ����(������ɺ���ý��㣬����ɹ����ٵ���ִ�У�ȡ����������ʧ�ܣ�����ִ��)
    If str����ҽ��IDs <> "" Then
        If gobjSquareCard.zlSquareAffirm(Me, pסԺҽ���´�, GetInsidePrivs(pסԺҽ���´�), mlng����ID, 0, False, , , str����ҽ��IDs, , , mblnʹ��Ԥ��) Then
            arrSQL = Array()
            arrAdviceID = Split(str����ҽ��IDs, ",")
            
            For i = 0 To UBound(arrAdviceID)
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "Zl_����ҽ��ִ��_Finish(" & arrAdviceID(i) & "," & lng���ͺ� & ",Null,0,'" & UserInfo.��� & "','" & UserInfo.���� & "'," & IIF(mlngҽ������ID <> 0, mlngҽ������ID, mlng���˿���id) & ")"
            Next
                            
            gcnOracle.BeginTrans: blnTran = True
            For i = 0 To UBound(arrSQL)
                Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
            Next
            gcnOracle.CommitTrans: blnTran = False
        End If
    End If
    
    'ҽ�������ϴ�
    If Not IsNull(mrsPati!����) Then
        If gclsInsure.GetCapability(supportҽ���ϴ�, mlng����ID, mrsPati!����) And gclsInsure.GetCapability(support������ɺ��ϴ�, mlng����ID, mrsPati!����) Then
            Do While Not rsUpload.EOF
                strMsg = ""
                If Not gclsInsure.TranChargeDetail(2, rsUpload!NO, 2, 1, strMsg, , mrsPati!����, rsUpload.RecordCount & "|" & rsUpload.AbsolutePosition) Then
                    '�ύ���ϴ�ʧ��,����ʾ
                    If strMsg <> "" Then
                        MsgBox strMsg, vbInformation, gstrSysName
                    Else
                        MsgBox mrsPati!���� & "�ļ��ʵ�""" & rsUpload!NO & """�ϴ�ʧ�ܣ�HIS���������ύ����ȷ���������͡�", vbExclamation, gstrSysName
                    End If
                Else
                    If strMsg <> "" Then MsgBox strMsg, vbInformation, gstrSysName
                End If
                rsUpload.MoveNext
            Loop
        End If
    End If
 
    Call SendMsg����(lng���ͺ�, IIF(mlng�������� = 1, 1, 2), IIF(bln����, 1, 2), str��ǰ����)
    
    'RIS�ӿ�
    If HaveRIS Then
        If GetAdviceRis(rsAdviceRis) Then
            On Error Resume Next
            If gobjRis.HISSendAdvice(rsAdviceRis, 2, mlng����ID, mlng��ҳID, "", lng���ͺ�) <> 1 Then
                MsgBox "��ǰ������Ӱ����Ϣϵͳ�ӿڣ� ������Ӱ����Ϣϵͳ�ӿ�(HISSendAdvice)δ���óɹ�������ϵͳ����Ա��ϵ��", vbInformation, gstrSysName
            End If
            err.Clear: On Error GoTo 0
        End If
    ElseIf gbln����Ӱ����Ϣϵͳ�ӿ� = True Then
        MsgBox "��ǰ������Ӱ����Ϣϵͳ�ӿڣ� ������RIS�ӿڴ���ʧ��δ����(HISSendAdvice)�ӿڣ�����ϵͳ����Ա��ϵ��", vbInformation, gstrSysName
    End If
    
    '�ύ�ɹ�,������ҽ���б��Ϊ��ɾ��
    With vsAdvice
        For i = .FixedRows To .Rows - 1
            If .Cell(flexcpData, i, COL_ѡ��) = 0 And Not .Cell(flexcpPicture, i, COL_ѡ��) Is Nothing Then
                .RowData(i) = -1
            End If
        Next
        
        '������ҽӿ�
        If CreatePlugInOK(pסԺҽ������) Then
            On Error Resume Next
            Call gobjPlugIn.AdviceSend(glngSys, pסԺҽ������, mlng����ID, mlng��ҳID, lng���ͺ�)
            Call zlPlugInErrH(err, "AdviceSend")
            On Error GoTo 0
        End If
        If gobjExchange Is Nothing Then
            On Error Resume Next
            Set gobjExchange = CreateObject("zlExchange.clsExchange")
            If Not gobjExchange Is Nothing Then Call gobjExchange.Init(gcnOracle)
            err.Clear: On Error GoTo 0
        End If
        '�������ݽ���ƽ̨����LIS,PACS�������뵥
        If Not gobjExchange Is Nothing Then
            For i = .FixedRows To .Rows - 1
                If .Cell(flexcpData, i, COL_ѡ��) = 0 And Not .Cell(flexcpPicture, i, COL_ѡ��) Is Nothing Then
                    'c-����,d-���
                    If .TextMatrix(i, COL_�������) = "C" Or .TextMatrix(i, COL_�������) = "D" Then
                        If Val(.TextMatrix(i, COL_���ID)) <> 0 Then
                            lng��ID = Val(.TextMatrix(i, COL_���ID))
                        Else
                            lng��ID = Val(.TextMatrix(i, COL_ID))
                        End If
                        If InStr(strҽ��IDs & ",", "," & lng��ID & ",") = 0 Then
                            strҽ��IDs = strҽ��IDs & "," & lng��ID
                            Call gobjExchange.SendMsg(IIF(.TextMatrix(i, COL_�������) = "C", 1, 2), "����ID::" & mlng����ID & "||��ҳID::" & mlng��ҳID & "||ҽ��ID::" & lng��ID & "||��������::1")
                        End If
                    End If
                End If
            Next
        End If
    End With
        
    CompletePatiSend = True
End Function

Private Sub SendMsg����(ByVal lng���ͺ� As Long, ByVal int�������� As Integer, ByVal int�������� As Integer, ByVal str��ǰ���� As String)
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim strIDs As String
    Dim lngTmp As Long
    Dim strTmp1 As String
    Dim strTmp2 As String
    Dim str�������� As String
    Dim i As Long
    Dim j As Long
    On Error GoTo errH
    strSQL = "select ���� from ���ű� where id=[1]"
    With vsAdvice
        For i = .FixedRows To .Rows - 1
            If .Cell(flexcpData, i, COL_ѡ��) = 0 And Not .Cell(flexcpPicture, i, COL_ѡ��) Is Nothing Then
                '���밲��
                If Val(.TextMatrix(i, COL_ִ�а���)) = 1 Then
                    Call ZLHIS_CIS_004(mclsMipModule, mlng����ID, mstr����, mstrסԺ��, , int��������, _
                        mlng��ҳID, .TextMatrix(i, COL_���˿���ID), "", , mstr����, Val(.TextMatrix(i, COL_ID)), 1, .TextMatrix(i, COL_�������), .TextMatrix(i, COL_��������), _
                        lng���ͺ�, .TextMatrix(i, COL_ִ�п���ID))
                End If
                '����ҽ��
                If .TextMatrix(i, COL_�������) = "E" And Val(.TextMatrix(i, COL_��������)) = 6 Then
                    strIDs = ""
                    For j = i - 1 To .FixedRows Step -1
                        If Val(.TextMatrix(j, COL_���ID)) <> Val(.TextMatrix(i, COL_ID)) Then
                            Exit For
                        Else
                            If .TextMatrix(j, COL_�������) = "C" Then
                                strIDs = strIDs & "," & Val(.TextMatrix(j, COL_ID))
                                lngTmp = Val(.TextMatrix(j, COL_ִ�п���ID))
                            End If
                        End If
                    Next
                    strIDs = Mid(strIDs, 2)
                    If strIDs <> "" Then
                        Call ZLHIS_CIS_016(mclsMipModule, mlng����ID, mstr����, mstrסԺ��, , int��������, mlng��ҳID, mlng���˿���id, , Val(.TextMatrix(i, COL_ID)), _
                            .TextMatrix(i, COL_�걾��λ), .TextMatrix(i, COL_������ĿID), , .TextMatrix(i, COL_ִ�п���ID), , strIDs, , lngTmp, , lng���ͺ�, "", _
                            int��������, .TextMatrix(i, COL_����ҽ��), .TextMatrix(i, COL_����ʱ��), .TextMatrix(i, COL_��������ID), , "")
                    End If
                '�������
                ElseIf .TextMatrix(i, COL_�������) = "D" And Val(.TextMatrix(i, COL_���ID)) = 0 Then
                    strTmp1 = ""
                    For j = i + 1 To .Rows - 1
                        If Val(.TextMatrix(j, COL_���ID)) <> Val(.TextMatrix(i, COL_ID)) Then
                            Exit For
                        Else
                            If .TextMatrix(j, COL_�������) = "D" Then
                                strTmp1 = strTmp1 & "," & .TextMatrix(j, COL_�걾��λ)
                            End If
                        End If
                    Next
                    strTmp1 = Mid(strTmp1, 2)
                    Call ZLHIS_CIS_017(mclsMipModule, mlng����ID, mstr����, mstrסԺ��, , int��������, mlng��ҳID, Val(.TextMatrix(i, COL_���˿���ID)), "", Val(.TextMatrix(i, COL_ID)), _
                        .TextMatrix(i, COL_������ĿID), .TextMatrix(i, col_ҽ������), strTmp1, .TextMatrix(i, COL_ִ�п���ID), , lng���ͺ�, _
                        "", int��������, .TextMatrix(i, COL_����ҽ��), .TextMatrix(i, COL_����ʱ��), .TextMatrix(i, COL_��������ID), , "")
                '��������
                ElseIf .TextMatrix(i, COL_�������) = "F" And Val(.TextMatrix(i, COL_���ID)) = 0 Then
                    strTmp1 = Getҽ����������(Val(.TextMatrix(i, COL_ID)), "����ҽ��")
                    strTmp2 = Getҽ����������(Val(.TextMatrix(i, COL_ID)), "����ҽ��")
                    For j = i + 1 To .Rows - 1
                        If Val(.TextMatrix(j, COL_���ID)) <> Val(.TextMatrix(i, COL_ID)) Then
                            Exit For
                        Else
                            If .TextMatrix(j, COL_�������) = "F" Then
                                strIDs = strIDs & "," & .TextMatrix(j, COL_ID)
                            ElseIf .TextMatrix(j, COL_�������) = "G" Then
                                lngTmp = Val(.TextMatrix(j, COL_ID))
                            End If
                        End If
                    Next
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.TextMatrix(i, COL_��������ID)))
                    If Not rsTmp.EOF Then str�������� = rsTmp!���� & ""
                    strIDs = Mid(strIDs, 2)
                    Call ZLHIS_CIS_018(mclsMipModule, mlng����ID, mstr����, mstrסԺ��, , int��������, _
                        mlng��ҳID, mlng���˿���id, "", Val(.TextMatrix(i, COL_ID)), strIDs, , lngTmp, , strTmp1, strTmp2, .TextMatrix(i, COL_ִ�п���ID), .TextMatrix(i, COL_ִ�п���), lng���ͺ�, _
                        "", int��������, .TextMatrix(i, COL_����ҽ��), .TextMatrix(i, COL_����ʱ��), .TextMatrix(i, COL_��������ID), str��������)
                '��Ѫ����
                ElseIf .TextMatrix(i, COL_�������) = "K" Then
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.TextMatrix(i, COL_��������ID)))
                    If Not rsTmp.EOF Then str�������� = rsTmp!���� & ""
                    Call ZLHIS_CIS_019(mclsMipModule, mlng����ID, mstr����, mstrסԺ��, , int��������, _
                        mlng��ҳID, mlng���˿���id, "", Val(.TextMatrix(i, COL_ID)), .TextMatrix(i, COL_ִ�п���ID), .TextMatrix(i, COL_ִ�п���), lng���ͺ�, _
                        "", int��������, .TextMatrix(i, COL_����ҽ��), .TextMatrix(i, COL_����ʱ��), .TextMatrix(i, COL_��������ID), str��������)
                ElseIf .TextMatrix(i, COL_�������) = "Z" And InStr(",7,8,11,", "," & .TextMatrix(i, COL_��������) & ",") > 0 _
                    Or .TextMatrix(i, COL_�������) = "E" And Val(.TextMatrix(i, COL_��������)) = 5 Then
                    If .TextMatrix(i, COL_��������) = "7" Then
                        strTmp1 = "ZLHIS_CIS_020" '��������
                    ElseIf .TextMatrix(i, COL_��������) = "8" Then
                        strTmp1 = "ZLHIS_CIS_021"  '��������ҽ��
                    ElseIf .TextMatrix(i, COL_��������) = "11" Then
                        strTmp1 = "ZLHIS_CIS_022" '��������ҽ��
                    ElseIf .TextMatrix(i, COL_�������) = "E" And Val(.TextMatrix(i, COL_��������)) = 5 Then
                        strTmp1 = "ZLHIS_CIS_023"  '������������ҽ��
                    End If
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.TextMatrix(i, COL_��������ID)))
                    If Not rsTmp.EOF Then str�������� = rsTmp!���� & ""
                    Call SendMsg(strTmp1, mclsMipModule, mlng����ID, mstr����, mstrסԺ��, , int��������, _
                        mlng��ҳID, Val(.TextMatrix(i, COL_���˿���ID)), "", Val(.TextMatrix(i, COL_ID)), .TextMatrix(i, COL_ִ�п���ID), .TextMatrix(i, COL_ִ�п���), lng���ͺ�, _
                        "", int��������, .TextMatrix(i, COL_����ҽ��), .TextMatrix(i, COL_����ʱ��), .TextMatrix(i, COL_��������ID), str��������)
                
                'סԺ����Ԥ��Ժ
                ElseIf .TextMatrix(i, COL_�������) = "Z" And Val(.TextMatrix(i, COL_��������)) = 5 Then
                    Call GetPatChange(Val(.TextMatrix(i, COL_ID)), 10, lngTmp, strTmp1)
                    Call ZLHIS_PATIENT_009(mclsMipModule, mlng����ID, mlng��ҳID, mstr����, mstr�Ա�, mstrסԺ��, _
                        lngTmp, .TextMatrix(i, COL_����ʱ��), mlng���˲���ID, , mlng���˿���id, "", , mstr����, Val(.TextMatrix(i, COL_ID)))
               
                'סԺ���߲�����
                ElseIf .TextMatrix(i, COL_�������) = "Z" And (Val(.TextMatrix(i, COL_��������)) = 9 Or Val(.TextMatrix(i, COL_��������)) = 10) Then
                    Call GetPatChange(Val(.TextMatrix(i, COL_ID)), 13, lngTmp, strTmp1)
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng���˿���id)
                    strTmp2 = "": If Not rsTmp.EOF Then strTmp2 = rsTmp!���� & ""
                    Call ZLHIS_PATIENT_005(mclsMipModule, mlng����ID, mlng��ҳID, mstr����, mstr�Ա�, mstrסԺ��, _
                        mlng���˲���ID, , mlng���˿���id, strTmp2, str��ǰ����, lngTmp, .TextMatrix(i, COL_����ʱ��), strTmp1, .TextMatrix(i, COL_����ҽ��), Val(.TextMatrix(i, COL_ID)))
                 
                'סԺ����ת������
                ElseIf .TextMatrix(i, COL_�������) = "Z" And Val(.TextMatrix(i, COL_��������)) = 3 Then
                    Call GetPatChange(Val(.TextMatrix(i, COL_ID)), 3, lngTmp, strTmp1)
                    Call ZLHIS_PATIENT_003(mclsMipModule, mlng����ID, mlng��ҳID, mstr����, mstr�Ա�, mstrסԺ��, _
                        mlng���˲���ID, , mlng���˿���id, "", , mstrסԺ��, _
                        lngTmp, .TextMatrix(i, COL_����ʱ��), , , Val(.TextMatrix(i, COL_ִ�п���ID)), , Val(.TextMatrix(i, COL_ID)))
                End If
            End If
        Next
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub ShowSendTotal()
'���ܣ����ݵ�ǰѡ��Ҫ���͵�ҽ�������㲢��ʾ���͵�ҽ���ϼ�
    Dim cur��� As Currency, curҩƷ��� As Currency, i As Long
    
    With vsAdvice
        For i = 1 To .Rows - 1
            If .Cell(flexcpData, i, COL_ѡ��) = 0 And Not .Cell(flexcpPicture, i, COL_ѡ��) Is Nothing Then
                '�ɼ��еĽ��:��һ��Ļ��ܽ��
                If Not .RowHidden(i) Then
                    cur��� = cur��� + Val(.TextMatrix(i, COL_���))
                End If
                'ҩƷ�Ľ��,ȡԭʼ���
                If InStr(",5,6,7,", .TextMatrix(i, COL_�������)) > 0 Then
                    curҩƷ��� = curҩƷ��� + Val(.Cell(flexcpData, i, COL_���))
                End If
            End If
        Next
    End With
    stbThis.Panels(4).Text = "���:" & FormatEx(cur���, gbytDec) & "(ҩ" & FormatEx(curҩƷ���, gbytDec) & ")"
    Call Form_Resize
End Sub


Private Function GetICUDeptID() As Long
'���ܣ���ȡ��ǰҽ�����ڵ�ICU����ID�������ٴ�����ʱ��
    Dim rsTmp As ADODB.Recordset, strSQL As String, i As Long, arrTmp As Variant
    
    strSQL = "Select a.����id, b.��������" & vbNewLine & _
            "From ������Ա A, ��������˵�� B" & vbNewLine & _
            "Where a.��Աid = [1] And a.����id = b.����id And b.�������� In ('ICU', '�ٴ�')"

    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.ID)
    If rsTmp.RecordCount > 0 Then
        rsTmp.Filter = "��������='�ٴ�'"
        If rsTmp.RecordCount = 0 Then
            rsTmp.Filter = "��������='ICU'"
                If rsTmp.RecordCount > 0 Then
                    GetICUDeptID = Val(rsTmp!����ID)
                End If
        Else
            strSQL = ""
            For i = 1 To rsTmp.RecordCount
                strSQL = strSQL & "," & rsTmp!����ID
                rsTmp.MoveNext
            Next
            arrTmp = Split(Mid(strSQL, 2), ",")
            For i = 0 To UBound(arrTmp)
                rsTmp.Filter = "��������='ICU' And ����ID=" & arrTmp(i)
                If rsTmp.RecordCount = 0 Then
                    GetICUDeptID = 0: Exit For
                Else
                    If i = 0 Then GetICUDeptID = arrTmp(i)
                End If
            Next
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Sub Del��������()
'���ܣ�ҽ������ʧ�ܣ�������˺󣬵��ü�������ɾ���ӿ�
    Dim i As Long, strҽ��IDs As String, strErr As String
        
    '�ռ��ɼ�����
    With vsAdvice
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, COL_��������)) = 6 And .TextMatrix(i, COL_�������) = "E" Then
                If .Cell(flexcpData, i, COL_ѡ��) = 0 And Not .Cell(flexcpPicture, i, COL_ѡ��) Is Nothing Then
                    strҽ��IDs = strҽ��IDs & "," & .TextMatrix(i, COL_ID)
                End If
            End If
        Next
    End With
    
    If strҽ��IDs <> "" Then
        strҽ��IDs = Mid(strҽ��IDs, 2)
        Call InitObjLis(pסԺҽ��վ)
        If Not gobjLIS Is Nothing Then
            If gobjLIS.DelLisApplicationForm(strҽ��IDs, strErr) = False Then
                MsgBox "ɾ����������ʧ�ܣ�" & strErr, vbInformation, gstrSysName
            End If
        End If
    End If
End Sub

Private Function Get����ҽ����Ϣ() As String
'���ܣ���ȡ����ҽ����Ϣ�����ݸ�����ӿڳ���
    Dim i As Long, strInfo As String
    
    With vsAdvice
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, COL_��������)) = 6 And .TextMatrix(i, COL_�������) = "E" Then
                '����ҽ��ID1,�ɼ�ҽ��ID1,ִ�п���ID1,�걾1;.....
                'LIS�ӿڲ����ļ��飬һ���ɼ���ʽֻ��һ������ҽ����û��һ���ɼ��������
                If .Cell(flexcpData, i, COL_ѡ��) = 0 And Not .Cell(flexcpPicture, i, COL_ѡ��) Is Nothing Then
                    strInfo = strInfo & ";" & .TextMatrix(i - 1, COL_ID) & "," & .TextMatrix(i, COL_ID) & "," & .TextMatrix(i - 1, COL_ִ�п���ID) & "," & .TextMatrix(i - 1, COL_�걾��λ)
                End If
            End If
        Next
    End With
    Get����ҽ����Ϣ = Mid(strInfo, 2)
End Function

Private Function Get��Ѫҽ����Ϣ() As String
'���ܣ���ȡ��Ѫҽ����Ϣ�����ݸ��ӿڳ��򣬽�ȡ��ҽ��ID
    Dim i As Long, strInfo As String
    
    With vsAdvice
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, COL_�������) = "K" Then
                If .Cell(flexcpData, i, COL_ѡ��) = 0 And Not .Cell(flexcpPicture, i, COL_ѡ��) Is Nothing Then
                    strInfo = strInfo & "," & .TextMatrix(i, COL_ID)
                End If
            End If
        Next
    End With
    Get��Ѫҽ����Ϣ = Mid(strInfo, 2)
End Function


Private Sub InitExecRecordset(rsExec As Recordset)
'���ܣ���ʼ��ҽ���Ƽۼ�¼��
    Set rsExec = New ADODB.Recordset
    
    rsExec.Fields.Append "ҽ��ID", adBigInt
    rsExec.Fields.Append "���ͺ�", adBigInt, , adFldIsNullable
    rsExec.Fields.Append "�շ�ϸĿID", adBigInt, , adFldIsNullable
    rsExec.Fields.Append "Ҫ��ʱ��", adDate, , adFldIsNullable
    rsExec.Fields.Append "����", adDouble, , adFldIsNullable
    rsExec.Fields.Append "��������", adInteger, , adFldIsNullable
    
    rsExec.CursorLocation = adUseClient
    rsExec.LockType = adLockOptimistic
    rsExec.CursorType = adOpenStatic
    rsExec.Open
End Sub

Private Function zlPluginAdviceBeforeSend() As Boolean
'���ܣ�ҽ������ǰ������Һ�
    Dim i As Long, j As Long
    Dim strAdviceIDs As String, strMsg  As String
    Dim rsDataPlugIn As ADODB.Recordset
    Dim lng���� As Long
    Dim str�ֽ�ʱ�� As String, strTmp As String
    Dim int���� As Integer
    
    zlPluginAdviceBeforeSend = True
    
    '������ҽӿڣ�ҽ������ǰ�ļ��
    If CreatePlugInOK(pסԺҽ������) Then
        Call InitPlugInRs(rsDataPlugIn)
        int���� = IIF("" = mstrǰ��IDs, 0, 2)
        With vsAdvice
            For i = .FixedRows To .Rows - 1
                If .Cell(flexcpData, i, COL_ѡ��) = 0 And Not .Cell(flexcpPicture, i, COL_ѡ��) Is Nothing Then
                    If .TextMatrix(i, COL_�ֽ�ʱ��) <> "" Then
                        str�ֽ�ʱ�� = .TextMatrix(i, COL_�ֽ�ʱ��)
                    Else
                        str�ֽ�ʱ�� = .Cell(flexcpData, i, COL_�ֽ�ʱ��)    '��ʼִ��ʱ��
                    End If
                    rsDataPlugIn.AddNew
                    rsDataPlugIn!����ID = mlng����ID
                    rsDataPlugIn!����ID = mlng��ҳID
                    rsDataPlugIn!ҽ��ID = Val(.TextMatrix(i, COL_ID))
                    rsDataPlugIn!���ID = Val(.TextMatrix(i, COL_���ID))
                    rsDataPlugIn!�շ�ϸĿID = Val(.TextMatrix(i, COL_�շ�ϸĿID))
                    rsDataPlugIn!�ֽ�ʱ�� = str�ֽ�ʱ��
                    rsDataPlugIn!���� = Val(.TextMatrix(i, COL_����))
                    rsDataPlugIn!���� = Val(.TextMatrix(i, COL_����))
                    rsDataPlugIn!������λ = .TextMatrix(i, COL_������λ)
                    rsDataPlugIn!���� = Val(.TextMatrix(i, COL_����))
                    rsDataPlugIn!������λ = .TextMatrix(i, COL_������λ)
                    rsDataPlugIn!���� = int����
                    rsDataPlugIn.Update
                End If
            Next
            If rsDataPlugIn.RecordCount > 0 Then rsDataPlugIn.MoveFirst
            strAdviceIDs = "": strMsg = ""
            On Error Resume Next
            Call gobjPlugIn.AdviceBeforeSend("", rsDataPlugIn, strAdviceIDs, strMsg)
            Call zlPlugInErrH(err, "AdviceBeforeSend")
            err.Clear
            On Error GoTo 0
             
            If strAdviceIDs <> "" Then
                strTmp = ""
                For i = .FixedRows To .Rows - 1
                    If .Cell(flexcpData, i, COL_ѡ��) = 0 And Not .Cell(flexcpPicture, i, COL_ѡ��) Is Nothing Then
                        If InStr("," & strAdviceIDs & ",", "," & Val(.TextMatrix(i, COL_ID)) & ",") > 0 Then
                            If Val(.TextMatrix(i, COL_���ID)) = 0 Then
                                j = Val(.TextMatrix(i, COL_ID))
                            Else
                                j = Val(.TextMatrix(i, COL_���ID))
                            End If
                            
                            If InStr("," & strTmp & ",", "," & j & ",") = 0 Then
                                strTmp = strTmp & "," & j
                            End If
                        End If
                    End If
                Next
                strAdviceIDs = Mid(strTmp, 2)
                lng���� = 0
                For i = .FixedRows To .Rows - 1
                    If .Cell(flexcpData, i, COL_ѡ��) = 0 And Not .Cell(flexcpPicture, i, COL_ѡ��) Is Nothing Then
                        If Val(.TextMatrix(i, COL_���ID)) = 0 Then
                            j = Val(.TextMatrix(i, COL_ID))
                        Else
                            j = Val(.TextMatrix(i, COL_���ID))
                        End If
                        lng���� = lng���� + 1
                        If InStr("," & strAdviceIDs & ",", "," & j & ",") > 0 Then
                            .Cell(flexcpData, i, COL_ѡ��) = 1
                            Set .Cell(flexcpPicture, i, COL_ѡ��) = frmIcons.imgTrueFalse.ListImages("F").Picture
                            lng���� = lng���� - 1
                        End If
                    End If
                Next
                
                If strMsg <> "" Then MsgBox strMsg, vbInformation, gstrSysName
                If lng���� = 0 Then
                    MsgBox "��ǰû�п��Է��͵�ҽ����", vbInformation, gstrSysName
                    zlPluginAdviceBeforeSend = False
                End If
            End If
        End With
    End If
End Function
 
Private Function GetAdviceRis(ByRef rsData As ADODB.Recordset) As Boolean
'���ܣ���ȡ���͵�RIS��ҽ����Ϣ
    Dim i As Long
    
    On Error GoTo errH
    
    Set rsData = New ADODB.Recordset
    
    rsData.Fields.Append "ҽ��ID", adBigInt
    rsData.Fields.Append "��������ID", adBigInt
    rsData.Fields.Append "ִ�п���ID", adBigInt
    rsData.Fields.Append "������ĿID", adBigInt
    rsData.Fields.Append "������Դ", adInteger '1-����;2-סԺ;
    rsData.Fields.Append "���", adVarChar, 10
    rsData.CursorLocation = adUseClient
    rsData.LockType = adLockOptimistic
    rsData.CursorType = adOpenStatic
    rsData.Open
    
    With vsAdvice
        For i = 1 To .Rows - 1
            If .Cell(flexcpData, i, COL_ѡ��) = 0 And Not .Cell(flexcpPicture, i, COL_ѡ��) Is Nothing Then
                If InStr(",D,F,", .TextMatrix(i, COL_�������)) > 0 Or InStr(",0,5,", Val(.TextMatrix(i, COL_��������))) > 0 And .TextMatrix(i, COL_�������) = "E" Then
                    If Val(.TextMatrix(i, COL_���ID)) = 0 Then
                        rsData.AddNew
                        rsData!ҽ��ID = Val(.TextMatrix(i, COL_ID))
                        rsData!��������id = Val(.TextMatrix(i, COL_��������ID))
                        rsData!ִ�п���ID = Val(.TextMatrix(i, COL_ִ�п���ID))
                        rsData!������ĿID = Val(.TextMatrix(i, COL_������ĿID))
                        rsData!������Դ = 2
                        rsData!��� = .TextMatrix(i, COL_�������)
                        rsData.Update
                    End If
                End If
            End If
        Next
    End With
    If rsData.RecordCount > 0 Then
        rsData.MoveFirst
        GetAdviceRis = True
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CheckRISScheduling() As Boolean
'���ܣ������Ŀ�Ƿ��Ǳ���ԤԼ
    Dim i As Long
    Dim blnDo As Boolean
    Dim lngҽ��ID As Long
    Dim lng������ĿID As Long
    Dim lngRst As Long
    Dim strMsg As String
    
    CheckRISScheduling = True
    
    If HaveRIS Then
        If gbln����Ӱ����ϢϵͳԤԼ Then
            blnDo = True
        End If
    End If
    
    If Not blnDo Then Exit Function
    
    With vsAdvice
        For i = 1 To .Rows - 1
            If .Cell(flexcpData, i, COL_ѡ��) = 0 And Not .Cell(flexcpPicture, i, COL_ѡ��) Is Nothing Then
                If Val(.TextMatrix(i, COL_������־)) <> 1 Then
                    If InStr(",D,F,", .TextMatrix(i, COL_�������)) > 0 Or InStr(",0,5,", Val(.TextMatrix(i, COL_��������))) > 0 And .TextMatrix(i, COL_�������) = "E" Then
                        If Val(.TextMatrix(i, COL_���ID)) = 0 Then
                            lngҽ��ID = Val(.TextMatrix(i, COL_ID))
                            lng������ĿID = Val(.TextMatrix(i, COL_������ĿID))
                            lngRst = -1
                            lngRst = gobjRis.HISScheduling(2, lngҽ��ID, lng������ĿID, False)
                            If lngRst <> 0 Then
                            '�ӿڷ���ʧ�ܸ�����ʾ
                                .Cell(flexcpData, i, COL_ѡ��) = 1 '��ǰ��ֹѡ��
                                Set .Cell(flexcpPicture, i, COL_ѡ��) = frmIcons.imgTrueFalse.ListImages("F").Picture
                                Call RowSelectSame(i, COL_ѡ��)
                                strMsg = IIF("" = strMsg, "", strMsg & "��") & .TextMatrix(i, col_ҽ������)
                            End If
                        End If
                    End If
                End If
            End If
        Next
    End With
    
    If strMsg <> "" Then
        MsgBox "����������RISϵͳԤԼ���̣�" & vbCrLf & "��" & strMsg & "��" & _
                vbCrLf & "ҽ��û��ԤԼ��ԤԼ�ɹ�����ܷ��͡�", vbInformation, gstrSysName
        CheckRISScheduling = False
    End If
End Function

Private Function Set������ҩ() As Boolean
'���ܣ�����ҩƷҽ���е�������ҩ˵��
    Dim i As Long
    Dim strMsg As String
    Dim str������ҩ As String
    Dim strSQL As String
    Dim strҽ��IDs As String
    
    On Error GoTo errH
    If mstrAdDrugIDs = "" Then
        Set������ҩ = True
        Exit Function
    End If
    With vsAdvice
        For i = .FixedRows To .Rows - 1
            If .Cell(flexcpData, i, COL_ѡ��) = 0 And Not .Cell(flexcpPicture, i, COL_ѡ��) Is Nothing Then
                If InStr("," & mstrAdDrugIDs & ",", "," & .TextMatrix(i, COL_ID) & ",") > 0 Then
                    strMsg = strMsg & "," & .TextMatrix(i, col_ҽ������)
                    strҽ��IDs = strҽ��IDs & "," & .TextMatrix(i, COL_ID)
                End If
            End If
        Next
    End With
    If strMsg = "" Then
        Set������ҩ = True
        Exit Function
    End If
    Call frmMsgDruExcess.ShowMe(Me, 1, Mid(strMsg, 2), str������ҩ)
    If str������ҩ = "*NULL*" Then
        Exit Function
    End If
    strSQL = "Zl_����ҽ����¼_������ҩ('" & Mid(strҽ��IDs, 2) & "','" & str������ҩ & "')"
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    Set������ҩ = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub Check�������()
'���ܣ������󷽽ӿ��жϵ�ǰҽ���ǲ���������
    Dim i As Long
    Dim str��ҩIDs As String '���뵽�ӿ��еĲ���
    Dim strOutҽ��IDs As String '���ܹ����͵���ҽ��ID
    Dim strҩ��ҽ��IDs As String
    Dim strErr As String
    Dim lngҽ��ID As Long
    Dim strҽ������ As String
    Dim rsTmp As ADODB.Recordset
    Dim blnTmp As Boolean
    
    On Error GoTo errH
    
    If Not gbln��ϵͳ Then Exit Sub
    
    With vsAdvice
        '���ú����ò������ṩ�Ľӿ�
        blnTmp = gobjPass.ZLPharmReviewResultView(mlng����ID, mlng��ҳID, rsTmp, strErr)
        '�ӿ�û����
        If blnTmp Then
            If strErr = "" Then
                If Not rsTmp Is Nothing Then
                    If Not rsTmp.EOF Then
                        For i = 1 To rsTmp.RecordCount
                            If InStr("," & strOutҽ��IDs & ",", "," & rsTmp!���ID & ",") = 0 Then
                                strOutҽ��IDs = strOutҽ��IDs & "," & rsTmp!���ID
                            End If
                            strҩ��ҽ��IDs = strҩ��ҽ��IDs & "," & rsTmp!ҽ��ID
                            rsTmp.MoveNext
                        Next
                    End If
                End If
            End If
        End If
            
        If strOutҽ��IDs <> "" Then
            'ȡ��ѡ��
            For i = .FixedRows To .Rows - 1
                If .Cell(flexcpData, i, COL_ѡ��) = 0 And Not .Cell(flexcpPicture, i, COL_ѡ��) Is Nothing Then
                    If Val(.TextMatrix(i, COL_ҽ��״̬)) = 1 Then
                        lngҽ��ID = IIF(0 = Val(.TextMatrix(i, COL_���ID)), Val(.TextMatrix(i, COL_ID)), Val(.TextMatrix(i, COL_���ID)))
                        If InStr("," & strOutҽ��IDs & ",", "," & lngҽ��ID & ",") > 0 Then
                            Set .Cell(flexcpPicture, i, COL_ѡ��) = Nothing 'ȱʡ������
                            If Val(.TextMatrix(i, COL_���ID)) <> 0 Then
                                If InStr("," & strҩ��ҽ��IDs & ",", "," & Val(.TextMatrix(i, COL_ID)) & ",") > 0 Then
                                    strҽ������ = strҽ������ & vbCrLf & .TextMatrix(i, col_ҽ������)
                                End If
                            End If
                        End If
                    End If
                End If
            Next
        End If
        If strҽ������ <> "" Then
            Call MsgBox("����ҽ��δͨ��������飬���ܷ��ͣ�" & strҽ������, vbInformation, Me.Caption)
        End If
            
       
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
