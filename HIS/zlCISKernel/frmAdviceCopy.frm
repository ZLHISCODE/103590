VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAdviceCopy 
   AutoRedraw      =   -1  'True
   Caption         =   "����ҽ��"
   ClientHeight    =   6540
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9960
   Icon            =   "frmAdviceCopy.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   9960
   Begin VB.ComboBox cboQX 
      Height          =   300
      Left            =   3945
      Style           =   2  'Dropdown List
      TabIndex        =   23
      Top             =   5820
      Width           =   1245
   End
   Begin VB.Frame fraLR 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5325
      Left            =   3270
      MousePointer    =   9  'Size W E
      TabIndex        =   16
      Top             =   870
      Width           =   45
   End
   Begin VSFlex8Ctl.VSFlexGrid vsPati 
      Height          =   5340
      Left            =   30
      TabIndex        =   7
      Top             =   840
      Width           =   3255
      _cx             =   5741
      _cy             =   9419
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
      BackColorSel    =   12632256
      ForeColorSel    =   0
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   2000
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmAdviceCopy.frx":058A
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   0   'False
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
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   465
      TabIndex        =   13
      ToolTipText     =   "F1"
      Top             =   6045
      Width           =   1100
   End
   Begin VSFlex8Ctl.VSFlexGrid vsAdvice 
      Height          =   5340
      Left            =   3390
      TabIndex        =   8
      Top             =   840
      Width           =   6525
      _cx             =   11509
      _cy             =   9419
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
      BackColorSel    =   12632256
      ForeColorSel    =   0
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   0
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   24
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   2000
      ColWidthMin     =   0
      ColWidthMax     =   5000
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmAdviceCopy.frx":0652
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
      FrozenCols      =   2
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   7935
      TabIndex        =   10
      Top             =   6045
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   6825
      TabIndex        =   9
      Top             =   6045
      Width           =   1100
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "ȫ��(&R)"
      Height          =   350
      Left            =   2745
      TabIndex        =   12
      ToolTipText     =   "Ctrl+R"
      Top             =   6045
      Width           =   1100
   End
   Begin VB.CommandButton cmdAll 
      Caption         =   "ȫѡ(&A)"
      Height          =   350
      Left            =   1635
      TabIndex        =   11
      ToolTipText     =   "Ctrl+A"
      Top             =   6045
      Width           =   1100
   End
   Begin VB.Frame fraPati 
      Height          =   900
      Left            =   45
      TabIndex        =   15
      Top             =   -75
      Width           =   9900
      Begin VB.ComboBox cboFinTim 
         Enabled         =   0   'False
         Height          =   300
         Left            =   3105
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   540
         Visible         =   0   'False
         Width           =   1665
      End
      Begin VB.OptionButton optType 
         Caption         =   "��ɾ���"
         Height          =   195
         Index           =   1
         Left            =   1230
         TabIndex        =   20
         Top             =   600
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.OptionButton optType 
         Caption         =   "���ھ���"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   19
         Top             =   600
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.PictureBox picDiag 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3255
         Picture         =   "frmAdviceCopy.frx":08EF
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   18
         Top             =   210
         Width           =   255
      End
      Begin VB.ComboBox cboBaby 
         Height          =   300
         Left            =   8370
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   195
         Width           =   1395
      End
      Begin VB.ComboBox cboTime 
         Height          =   300
         Left            =   3105
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   195
         Width           =   3870
      End
      Begin VB.CommandButton cmdPati 
         Height          =   240
         Left            =   1950
         Picture         =   "frmAdviceCopy.frx":7141
         Style           =   1  'Graphical
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "ѡ����(F4)"
         Top             =   225
         Width           =   255
      End
      Begin VB.TextBox txtPati 
         Height          =   300
         Left            =   780
         TabIndex        =   1
         Top             =   195
         Width           =   1455
      End
      Begin VB.Label lblTim 
         AutoSize        =   -1  'True
         Caption         =   "����ʱ��"
         Height          =   180
         Left            =   2340
         TabIndex        =   22
         Top             =   600
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Label lblDiag 
         Caption         =   "������ϣ�   "
         Height          =   180
         Left            =   2400
         TabIndex        =   17
         Top             =   225
         Width           =   7335
      End
      Begin VB.Label lblBaby 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ӥ��(&B)"
         Height          =   180
         Left            =   7695
         TabIndex        =   5
         Top             =   255
         Width           =   630
      End
      Begin VB.Label lblTime 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ʱ��(&T)"
         Height          =   180
         Left            =   2430
         TabIndex        =   3
         Top             =   255
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����(&P)"
         Height          =   180
         Left            =   135
         TabIndex        =   0
         Top             =   255
         Width           =   630
      End
   End
   Begin MSComctlLib.ListView lvwPati 
      Height          =   3975
      Left            =   795
      TabIndex        =   14
      Top             =   450
      Visible         =   0   'False
      Width           =   5700
      _ExtentX        =   10054
      _ExtentY        =   7011
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "img32"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   12
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "����"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "סԺ��"
         Object.Width           =   1499
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "����"
         Object.Width           =   1111
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "סԺҽʦ"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "�Ա�"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "����"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "�ѱ�"
         Object.Width           =   1499
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "����ȼ�"
         Object.Width           =   2028
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   $"frmAdviceCopy.frx":7237
         Object.Width           =   2857
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "��Ժ����"
         Object.Width           =   2857
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "���ʽ"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "��������"
         Object.Width           =   2646
      EndProperty
   End
End
Attribute VB_Name = "frmAdviceCopy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mfrmParent As Object
Private mMainPrivs As String
Private mbln��ʿվ As Boolean
Private mlngǰ��ID As Long
Private mlng����ID As Long
Private mlng��ҳID As Long
Private mlngPatiID As Long '���洫��Ĳ���ID
Private mstrNoIn As String '���洫��ĹҺŵ�������ǰ���˵ĹҺŵ�
Private mstr�Һŵ� As String
Private mblnMoved As Boolean
Private mblnItem As Boolean
Private mstrIDs As String
Private mstrAlter As String
Private mlngӤ�� As String
Private mstr�Ա� As String

Private mtmp��ҳID As Long
Private mtmp�Һŵ� As String
Private mlng���˿���id As Long
Private mlngӤ������ID As Long
Private mblnҽ������ As Boolean
Private mint��Դ As Integer

Private Enum COLҽ��
    colѡ�� = 0
    col��Ч = 1
    colʱ�� = 2
    col���� = 3
    col���� = 4
    col������λ = 5
    col���� = 6
    col������λ = 7
    colƵ�� = 8
    col�÷� = 9
    col���� = 10
    colִ��ʱ�� = 11
    colִ�п��� = 12
    colID = 13
    col���ID = 14
    col������� = 15
    col������ĿID = 16
    col�շ�ϸĿID = 17
    col�Ƿ����� = 18
    col������� = 19
    col��ֵ���� = 20
    col�Ա� = 21
    col����Ӧ�� = 22
    col�������� = 23
    col��Ŀ���� = 24
    col�շ����� = 25
End Enum

Private Enum mCtlID
    opt���ھ��� = 0
    opt��ɾ��� = 1
End Enum

Private Const con��Ŀ���� = -1
Private Const con��Ŀ���� = -2
Private Const con��Ŀ��� = -3
Private Const con�շѳ��� = -4
Private Const con�շѷ��� = -5

Private mbln������Ȩ�� As Boolean 'û���´�������Ȩ��ʱΪTrue,���´�������Ȩ��ʱΪFalse
Private mbln������Ȩ�� As Boolean 'û���´ﶾ����Ȩ��ʱΪTrue,���´ﶾ����Ȩ��ʱΪFalse
Private mbln������Ȩ�� As Boolean 'û���´ﾫ����Ȩ��ʱΪTrue,���´ﾫ����Ȩ��ʱΪFalse
Private mbln������Ȩ�� As Boolean 'û���´������Ȩ��ʱΪTrue,���´������Ȩ��ʱΪFalse
Private mintOutPreTime As Integer
Private mdtOutBegin As Date, mdtOutEnd As Date
Private mlngPre��Ч As Long

Public Function ShowMe(ByVal frmParent As Object, ByVal strPrivs As String, lng����ID As Long, varTime As Variant, blnMoved As Boolean, _
    Optional ByVal bln��ʿվ As Boolean, Optional ByVal lngǰ��ID As Long, Optional strAlter As String, Optional lng���˿���ID As Long, _
    Optional lngӤ�� As Long, Optional lngӤ������ID As Long, Optional str�Ա� As String) As String
'���أ�lng����ID,varTime=Ҫ����ҽ���Ĳ���ID����ҳID(�Һŵ�NO)
'      blnMoved=Ҫ���Ʋ��˵�ҽ���Ƿ�ת��
'      strAlter=���θ��Ƶ�ҽ����Ҫ�л���Ч��ҽ��ID(��ID):123,456,...
'      ShowMe=Ҫ���Ƶ�ҽ������ID��
    Set mfrmParent = frmParent
    mMainPrivs = strPrivs
    mbln��ʿվ = bln��ʿվ
    mlngǰ��ID = lngǰ��ID
    mlng����ID = lng����ID
    mlngӤ�� = lngӤ��
    mlng���˿���id = lng���˿���ID
    mlngӤ������ID = lngӤ������ID
    mstr�Ա� = str�Ա�
    If TypeName(varTime) = "String" Then
        mstr�Һŵ� = varTime
        mlng��ҳID = 0
    Else
        mlng��ҳID = varTime
        mstr�Һŵ� = ""
    End If
    mstrNoIn = mstr�Һŵ�: mlngPatiID = mlng����ID
    mblnMoved = blnMoved
    strAlter = "": mstrAlter = strAlter
    
    Me.Show 1, frmParent
    
    lng����ID = mlng����ID
    If TypeName(varTime) = "String" Then
        varTime = mstr�Һŵ�
    Else
        varTime = mlng��ҳID
    End If
    blnMoved = mblnMoved
    strAlter = mstrAlter
    ShowMe = mstrIDs
End Function

Private Function LoadPatients() As Boolean
'���ܣ���ȡ����ý�����ͬ��Χ�Ĳ����б�
    Dim rsTmp As ADODB.Recordset
    Dim objItem As ListItem, strSQL As String
    Dim i As Integer, j As Integer
    Dim lng����ID As Long, intBedLen As Long
    Dim curDate As Date, dtOutEnd As Date, dtOutBegin As Date
    Dim intTmp As Integer
    
    On Error GoTo errH
    
    '��ȡ��Ժ���˵�ʱ�䷶Χ
    curDate = zlDatabase.Currentdate
    intTmp = Val(zlDatabase.GetPara("��Ժ���˽������", glngSys, IIF(mbln��ʿվ, pסԺ��ʿվ, pסԺҽ��վ), 0))
    dtOutEnd = Format(curDate + intTmp, "yyyy-MM-dd 23:59:59")
    intTmp = Val(zlDatabase.GetPara("��Ժ���˿�ʼ���", glngSys, IIF(mbln��ʿվ, pסԺ��ʿվ, pסԺҽ��վ), 1))
    dtOutBegin = Format(curDate - intTmp, "yyyy-MM-dd 00:00:00")
    
    If mlngǰ��ID <> 0 Then
        cmdPati.Visible = False
        If mstr�Һŵ� <> "" Then
            strSQL = "Select B.NO,B.����ID,B.�����,B.����,B.�Ա�,B.����,A.����," & _
                " C.���� as ����,B.ִ��ʱ�� as ����ʱ��,Decode(B.ִ��״̬,0,'�ȴ�����',1,'�������',2,'���ھ���') as ����״̬" & _
                " From ������Ϣ A,���˹Һż�¼ B,���ű� C" & _
                " Where A.����ID=B.����ID And B.ִ�в���ID=C.ID And A.����ID=[1] And B.NO=[2] And B.��¼����=1 And B.��¼״̬=1"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mstr�Һŵ�)
        Else
            strSQL = _
                "Select A.����ID,B.��ҳID,B.סԺ��,NVL(B.����,A.����) ����,NVL(B.�Ա�,A.�Ա�) �Ա� ,NVL(B.����,A.����) ���� ," & _
                " B.��Ժ����,B.��Ժ����,B.סԺҽʦ,B.��Ժ���� as ����,B.�ѱ�," & _
                " B.����,B.��Ժ����ID as ����ID,B.��ǰ����ID as ����ID,D.���� as ����,C.���� as ����ȼ�," & _
                " B.״̬,B.����ת��,Nvl(B.ҽ�Ƹ��ʽ,A.ҽ�Ƹ��ʽ) as ҽ�Ƹ��ʽ,B.��������" & _
                " From ������Ϣ A,������ҳ B,�շ���ĿĿ¼ C,���ű� D" & _
                " Where A.����ID=B.����ID And B.����ȼ�ID=C.ID(+) And B.��Ժ����ID=D.ID" & _
                " And A.����ID=[1] And B.��ҳID=[2]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng��ҳID)
        End If
    Else
        If mstrNoIn <> "" Then
            If optType(opt��ɾ���).value Then
                strSQL = "Select B.NO,B.����ID,B.�����,B.����,B.�Ա�,B.����,A.����,C.���� as ����,B.ִ��ʱ�� as ����ʱ��,'�������' as ����״̬" & _
                    " From ������Ϣ A,���˹Һż�¼ B,���ű� C" & _
                    " Where A.����ID=B.����ID And B.ִ�в���ID=C.ID And B.ִ��״̬+0=1 And B.ִ����||''=[1] And B.��¼����=1 And B.��¼״̬=1" & _
                    " and B.ִ��ʱ�� between [2] and [3] Order By NO"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.����, CDate(Format(mdtOutBegin, "yyyy-MM-dd 00:00:00")), CDate(Format(mdtOutEnd, "yyyy-MM-dd 23:59:59")))
                '���˳����˺�ȡ��һ����
                If rsTmp.EOF Then
                    mstr�Һŵ� = ""
                    mlng����ID = 0
                Else
                    mstr�Һŵ� = rsTmp!NO & ""
                    mlng����ID = Val(rsTmp!����ID & "")
                End If
            Else
                mstr�Һŵ� = mstrNoIn: mlng����ID = mlngPatiID
                '�ṩ��ǰҽ�����ھ�����������Ĳ����嵥��ѡ��:�������ݲ��漰�жϺͶ�ȡ"H���˹Һż�¼"
                strSQL = "Select B.NO,B.����ID,B.�����,B.����,B.�Ա�,B.����,A.����,Decode(B.ִ��״̬,0,0,1,2,2,1) as ����," & _
                    " C.���� as ����,B.ִ��ʱ�� as ����ʱ��,Decode(B.ִ��״̬,0,'�ȴ�����',1,'�������',2,'���ھ���') as ����״̬" & _
                    " From ������Ϣ A,���˹Һż�¼ B,���ű� C Where A.����ID=B.����ID And B.ִ�в���ID=C.ID And A.����ID=[1] And B.NO=[2] And B.��¼����=1 And B.��¼״̬=1" & _
                    " Union " & _
                    " Select B.NO,B.����ID,B.�����,B.����,B.�Ա�,B.����,A.����,1 as ����,C.���� as ����,B.ִ��ʱ�� as ����ʱ��,'���ھ���' as ����״̬" & _
                    " From ������Ϣ A,���˹Һż�¼ B,���ű� C Where A.����ID=B.����ID And B.ִ�в���ID=C.ID And B.ִ��״̬=2 And B.ִ����||''=[3] And B.��¼����=1 And B.��¼״̬=1 Order By ����,NO"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mstr�Һŵ�, UserInfo.����)
            End If
        Else
            strSQL = "Select ��Ժ����ID as ����ID,��ǰ����ID  as ����ID,Ӥ������ID,Ӥ������ID From ������ҳ Where ����ID=[1] And ��ҳID=[2]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng��ҳID)
                
            '�ṩ��ǰ����/��������Ժ�����嵥��ѡ��
            lng����ID = IIF(mbln��ʿվ, IIF(mlngӤ������ID <> 0, NVL(rsTmp!Ӥ������ID, 0), NVL(rsTmp!����ID, 0)), IIF(mlngӤ������ID <> 0, NVL(rsTmp!Ӥ������ID, 0), NVL(rsTmp!����ID, 0)))
            intBedLen = GetMaxBedLen(lng����ID, Not mbln��ʿվ)
            strSQL = _
                "Select decode(b.סԺҽʦ,[4],1,2) as ����,A.����ID,B.��ҳID,B.סԺ��,NVL(B.����,A.����) ����,NVL(B.�Ա�,A.�Ա�) �Ա� ,NVL(B.����,A.����) ����,B.��Ժ����,B.��Ժ����," & _
                " B.סԺҽʦ,LPAD(B.��Ժ����," & intBedLen & ",' ') as ����,B.�ѱ�,B.����," & _
                " B.��Ժ����ID as ����ID,D.���� as ����,C.���� as ����ȼ�,B.״̬,B.����ת��," & _
                " Nvl(B.ҽ�Ƹ��ʽ,A.ҽ�Ƹ��ʽ) as ҽ�Ƹ��ʽ,B.��������" & _
                " From ������Ϣ A,������ҳ B,�շ���ĿĿ¼ C,���ű� D" & _
                " Where A.����ID=B.����ID And B.����ȼ�ID=C.ID(+) And B.��Ժ����ID=D.ID And A.����ID=[1] And B.��ҳID=[2]"
            strSQL = strSQL & " Union " & _
                "Select decode(b.סԺҽʦ,[4],1,2) as ����,A.����ID,B.��ҳID,B.סԺ��,NVL(B.����,A.����) ����,NVL(B.�Ա�,A.�Ա�) �Ա� ,NVL(B.����,A.����) ����,B.��Ժ����,B.��Ժ����," & _
                " B.סԺҽʦ,LPAD(B.��Ժ����," & intBedLen & ",' ') as ����,B.�ѱ�,B.����," & _
                " B.��Ժ����ID as ����ID,D.���� as ����,C.���� as ����ȼ�,B.״̬,B.����ת��," & _
                " Nvl(B.ҽ�Ƹ��ʽ,A.ҽ�Ƹ��ʽ) as ҽ�Ƹ��ʽ,B.��������" & _
                " From ������Ϣ A,������ҳ B,�շ���ĿĿ¼ C,���ű� D,��Ժ���� R" & _
                " Where A.����ID=B.����ID And A.��ҳID=B.��ҳID And B.����ȼ�ID=C.ID(+)" & _
                " And B.��Ժ����ID=D.ID And a.����ID=R.����ID " & IIF(mbln��ʿվ, " And A.��ǰ����ID=R.����ID", "  And A.��ǰ����ID=R.����ID") & _
                IIF(mbln��ʿվ, " And (R.����ID=[3] Or b.Ӥ������ID=[3])", " And (R.����ID=[3] Or b.Ӥ������ID=[3])") & _
                IIF(Not mbln��ʿվ And InStr(mMainPrivs, "���Ʋ���") = 0, " And B.סԺҽʦ=[4]", "")
            strSQL = strSQL & " Union " & _
                "Select decode(b.סԺҽʦ,[4],1,2) as ����,A.����ID,B.��ҳID,B.סԺ��,NVL(B.����,A.����) ����,NVL(B.�Ա�,A.�Ա�) �Ա� ,NVL(B.����,A.����) ����,B.��Ժ����,B.��Ժ����," & _
                " B.סԺҽʦ,LPAD(B.��Ժ����," & intBedLen & ",' ') as ����,B.�ѱ�,B.����," & _
                " B.��Ժ����ID as ����ID,D.���� as ����,C.���� as ����ȼ�,B.״̬,B.����ת��," & _
                " Nvl(B.ҽ�Ƹ��ʽ,A.ҽ�Ƹ��ʽ) as ҽ�Ƹ��ʽ,B.��������" & _
                " From ������Ϣ A,������ҳ B,�շ���ĿĿ¼ C,���ű� D" & _
                " Where A.����ID=B.����ID And Nvl(B.��ҳID,0)<>0 And B.����ȼ�ID=C.ID(+)" & _
                " And B.��Ժ����ID=D.ID And B.��Ժ���� between [5] and [6]" & _
                IIF(mbln��ʿվ, " And B.��ǰ����ID+0=[3]", " And B.��Ժ����ID+0=[3]") & _
                IIF(Not mbln��ʿվ And InStr(mMainPrivs, "���Ʋ���") = 0, " And B.סԺҽʦ=[4]", "") & _
                " Order by ����,����"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng��ҳID, lng����ID, UserInfo.����, dtOutBegin, dtOutEnd)
        End If
    End If
    
    lvwPati.ListItems.Clear
    For i = 1 To rsTmp.RecordCount
        If mstr�Һŵ� <> "" Then
            Set objItem = lvwPati.ListItems.Add(, "_" & rsTmp!����ID & "_" & rsTmp!NO, rsTmp!����, , "Pati")
            objItem.SubItems(1) = NVL(rsTmp!�����)
            objItem.SubItems(2) = NVL(rsTmp!�Ա�)
            objItem.SubItems(3) = NVL(rsTmp!����)
            objItem.SubItems(4) = NVL(rsTmp!����)
            objItem.SubItems(5) = Format(NVL(rsTmp!����ʱ��), "yyyy-MM-dd HH:mm")
            objItem.SubItems(6) = NVL(rsTmp!����״̬)
            objItem.SubItems(7) = NVL(rsTmp!NO)
            
            '���ղ����ú�ɫ��ʾ
            If Not IsNull(rsTmp!����) Then
                Call SetItemColor(objItem, vbRed)
            End If
            
            '��ʾ��ʼ���˵���Ϣ
            If rsTmp!����ID = mlng����ID And rsTmp!NO = mstr�Һŵ� Then
                With objItem
                    txtPati.ForeColor = .ForeColor
                    txtPati.Text = .Text
                    .Selected = True 'һ��Ҫѡ�е�ǰ����
                End With
            End If
        Else
            Set objItem = lvwPati.ListItems.Add(, "_" & rsTmp!����ID & "_" & rsTmp!��ҳID, rsTmp!����, , "Pati")
            objItem.SubItems(1) = NVL(rsTmp!סԺ��)
            objItem.SubItems(2) = NVL(rsTmp!����)
            objItem.SubItems(3) = NVL(rsTmp!סԺҽʦ)
            objItem.SubItems(4) = NVL(rsTmp!�Ա�)
            objItem.SubItems(5) = NVL(rsTmp!����)
            objItem.SubItems(6) = NVL(rsTmp!����)
            objItem.SubItems(7) = NVL(rsTmp!�ѱ�)
            objItem.SubItems(8) = NVL(rsTmp!����ȼ�)
            objItem.SubItems(9) = Format(rsTmp!��Ժ����, "yyyy-MM-dd HH:mm")
            objItem.SubItems(10) = Format(NVL(rsTmp!��Ժ����), "yyyy-MM-dd HH:mm")
            objItem.SubItems(11) = NVL(rsTmp!ҽ�Ƹ��ʽ)
            objItem.SubItems(12) = NVL(rsTmp!��������)
            objItem.Tag = NVL(rsTmp!����ת��, 0)
            
            '������ɫ
            Call SetItemColor(objItem, zlDatabase.GetPatiColor(NVL(rsTmp!��������)))
            
            '��ʾ��ʼ���˵���Ϣ
            If rsTmp!����ID = mlng����ID And rsTmp!��ҳID = mlng��ҳID Then
                With objItem
                    txtPati.ForeColor = .ForeColor
                    txtPati.Text = .Text
                    .Selected = True 'һ��Ҫѡ�е�ǰ����
                End With
            End If
        End If
        rsTmp.MoveNext
    Next
    
    LoadPatients = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub LoadPatiTime()
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim i As Long
    
    On Error GoTo errH
    
    cboTime.Clear
    cboBaby.Clear
    vsPati.Rows = vsPati.FixedRows
    vsPati.Rows = vsPati.FixedRows + 1
    vsPati.Row = 1
    vsAdvice.Rows = vsAdvice.FixedRows
    vsAdvice.Rows = vsAdvice.FixedRows + 1
    vsAdvice.Row = 1
    
    If mstr�Һŵ� = "" And optType(opt���ھ���).Visible Then
        txtPati.Text = ""
        txtPati.Enabled = False
        cmdPati.Enabled = False
        MsgBox "δ�ҵ��κβ��˵ľ�����Ϣ��", vbInformation, gstrSysName
        Exit Sub
    Else
        txtPati.Enabled = True
        cmdPati.Enabled = True
    End If
        
    If InStr(GetInsidePrivs(IIF(mstr�Һŵ� = "", pסԺҽ���´�, p����ҽ���´�)), ";��������ҽ��;") = 0 Then
        txtPati.Locked = True
        cmdPati.Enabled = False
        If mstr�Һŵ� <> "" Then
            optType(opt���ھ���).Enabled = False
            optType(opt��ɾ���).Enabled = False
        End If
    End If
    
    If mstr�Һŵ� <> "" Then
        strSQL = "Select A.ID,A.NO,A.����ʱ��,B.���� as ����,A.ִ���� as ҽ��,A.����,A.����,A.����" & _
            " From ���˹Һż�¼ A,���ű� B Where A.ִ�в���ID=B.ID And A.����ID=[1] And a.��¼����=1 And a.��¼״̬=1 Order by A.����ʱ�� Desc"
    Else
        strSQL = "Select A.��ҳID,A.��Ժ����,B.���� as ���� From ������ҳ A,���ű� B" & _
            " Where A.��Ժ����ID=B.ID And A.����ID=[1] And Nvl(A.��ҳID,0)<>0 Order by A.��ҳID Desc"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID)
    vsPati.Redraw = flexRDNone
    Do While Not rsTmp.EOF
        If mstr�Һŵ� <> "" Then
            cboTime.AddItem "[" & Format(rsTmp!����ʱ��, "yyyy-MM-dd HH:mm") & "]" & rsTmp!NO & "," & rsTmp!����
            cboTime.ItemData(cboTime.NewIndex) = rsTmp!ID
            
            With vsPati
                If .RowData(.Rows - 1) <> 0 Then .AddItem ""
                .RowData(.Rows - 1) = Val(rsTmp!ID)
                .TextMatrix(.Rows - 1, 0) = Format(rsTmp!����ʱ��, "yyyy-MM-dd HH:mm")
                .TextMatrix(.Rows - 1, 1) = NVL(rsTmp!����)
                .TextMatrix(.Rows - 1, 2) = NVL(rsTmp!ҽ��)
                .TextMatrix(.Rows - 1, 3) = NVL(rsTmp!����)
                .TextMatrix(.Rows - 1, 4) = IIF(NVL(rsTmp!����) = 1, "", "")
                .TextMatrix(.Rows - 1, 5) = IIF(NVL(rsTmp!����) = 1, "", "")
                .TextMatrix(.Rows - 1, 6) = rsTmp!NO
            End With
            
            If rsTmp!NO = mstr�Һŵ� Then
                cboTime.ListIndex = cboTime.NewIndex
                vsPati.Row = vsPati.Rows - 1
            End If
        Else
            cboTime.AddItem "[" & Format(rsTmp!��Ժ����, "yyyy-MM-dd HH:mm") & "]��" & rsTmp!��ҳID & "��סԺ," & rsTmp!����
            cboTime.ItemData(cboTime.NewIndex) = rsTmp!��ҳID
            If rsTmp!��ҳID = mlng��ҳID Then cboTime.ListIndex = cboTime.NewIndex
        End If
        rsTmp.MoveNext
    Loop
    
    With vsPati
        If vsPati.Tag = "���ز���" And .Rows > 4 Then
            For i = 4 To .Rows - 1
                .RowHidden(i) = True
            Next
            .AddItem ""
            .RowData(.Rows - 1) = -1
            .TextMatrix(.Rows - 1, 0) = "��ʾȫ��"
        End If
        .Row = .FixedRows
    End With
    Call vsPati.ShowCell(vsPati.Row, 0)
    vsPati.Redraw = flexRDDirect
    If cboTime.ListIndex = -1 Then cboTime.ListIndex = 0
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub SetItemColor(ByVal objItem As ListItem, ByVal lngColor As Long)
    Dim i As Long
    
    objItem.ForeColor = lngColor
    For i = 1 To objItem.ListSubItems.Count
        objItem.ListSubItems(i).ForeColor = lngColor
    Next
End Sub

Private Sub cboBaby_Click()
    Call LoadAdvice
End Sub

Private Sub cboFinTim_Click()
'���ܣ���ʱ�䷶Χ��ָ���ǣ�����ʱ��ѡ����
    Dim intDateCount As Integer
    Dim datCurr As Date
    
    With cboFinTim
        intDateCount = .ItemData(.ListIndex)
        If .ListIndex = mintOutPreTime And intDateCount <> -1 Then Exit Sub
        If intDateCount = -1 Then
            If Not frmSelectTime.ShowMe(Me, mdtOutBegin, mdtOutEnd, cboFinTim) Then
                'ȡ��ʱ�ָ�ԭ����ѡ��
                Call cbo.SetIndex(.hwnd, mintOutPreTime)
                Exit Sub
            End If
        Else
            mdtOutEnd = CDate(.Tag)
            mdtOutBegin = mdtOutEnd - intDateCount
        End If

        .ToolTipText = "��Χ��" & Format(mdtOutBegin, "yyyy-MM-dd 00:00") & " �� " & Format(mdtOutEnd, "yyyy-MM-dd 23:59")
        lblTim.ToolTipText = .ToolTipText
        mintOutPreTime = .ListIndex
    End With
    datCurr = CDate(cboFinTim.Tag)
    
    Call LoadPatients
    Call LoadPatiTime
End Sub

Private Sub cboTime_Click()
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, blnVisible As Boolean
    
    If cboTime.ListIndex = -1 Then Exit Sub
    
    On Error GoTo errH
    
    cboBaby.Clear
    vsAdvice.Rows = vsAdvice.FixedRows
    vsAdvice.Rows = vsAdvice.FixedRows + 1
    vsAdvice.Row = 1
    
    If mstr�Һŵ� <> "" Then
        strSQL = "Select Distinct A.Ӥ�� From ����ҽ����¼ A,���˹Һż�¼ B Where A.�Һŵ�=B.NO And B.ID=[2] Order by Nvl(A.Ӥ��,0)"
    Else
        strSQL = "Select Distinct Ӥ�� From ����ҽ����¼ Where ����ID=[1] And ��ҳID=[2] Order by Nvl(Ӥ��,0)"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, cboTime.ItemData(cboTime.ListIndex))
    Do While Not rsTmp.EOF
        If NVL(rsTmp!Ӥ��, 0) = 0 Then
            cboBaby.AddItem "����ҽ��"
        Else
            cboBaby.AddItem "Ӥ�� " & rsTmp!Ӥ�� & " ҽ��"
        End If
        cboBaby.ItemData(cboBaby.NewIndex) = NVL(rsTmp!Ӥ��, 0)
        If NVL(rsTmp!Ӥ��, 0) = mlngӤ�� Then cboBaby.ListIndex = cboBaby.NewIndex
        rsTmp.MoveNext
    Loop
    If cboBaby.ListIndex = -1 And cboBaby.ListCount > 0 Then cboBaby.ListIndex = 0
    Call LoadDiag
    
    blnVisible = cboBaby.ListCount > 0
    If cboBaby.ListCount = 1 Then
        If cboBaby.ItemData(cboBaby.ListIndex) = 0 Then
            blnVisible = False
        End If
    End If
    cboBaby.Visible = blnVisible
    lblBaby.Visible = blnVisible
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub LoadDiag()
'���ܣ����ظôξ���������Ϣ
    Dim lng����ID As Long
    Dim strSQL As String, rsTmp As Recordset
    Dim i As Integer
    Dim strDiag As String
    On Error GoTo errH
    
    lng����ID = cboTime.ItemData(cboTime.ListIndex)
    strSQL = "Select �������,������� From ������ϼ�¼ Where ��ҳid = [2] And ����id = [1] And ������� in (11,1) Order By ��ϴ���"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, lng����ID)
    picDiag.Tag = ""
    lblDiag.Caption = "������ϣ�   "
    If rsTmp.RecordCount > 0 Then
        For i = 1 To rsTmp.RecordCount
            If 1 = Val(rsTmp!�������) Then
                strDiag = strDiag & "," & rsTmp!�������
                picDiag.Tag = picDiag.Tag & "������" & rsTmp!������� & vbCrLf
            Else
                picDiag.Tag = picDiag.Tag & "���С�" & rsTmp!������� & vbCrLf
            End If
            rsTmp.MoveNext
        Next
        lblDiag.Caption = lblDiag.Caption & Mid(strDiag, 2)
    End If
    If picDiag.Tag = "" Then
        picDiag.Tag = "û���κ���ϣ�"
    Else 'ȥ��ĩβ�Ļس���
        picDiag.Tag = Mid(picDiag.Tag, 1, Len(picDiag.Tag) - 2)
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdALL_Click()
    Dim i As Long, lngEnd As Long
    
    With vsAdvice
        For i = .FixedRows To .Rows - 1
            If CheckCanSelGroup(i, False) Then
                Call SelGroup(i, 1, lngEnd)
            End If
            If i < lngEnd Then i = lngEnd
        Next
    End With
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdClear_Click()
    Dim i As Long
    
    With vsAdvice
        For i = .FixedRows To .Rows - 1
            .TextMatrix(i, colѡ��) = 0
        Next
    End With
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name
End Sub

Private Sub cmdOK_Click()
    Dim lngID As Long, i As Long
    Dim strIDs As String, strAlter As String
    
    With vsAdvice
        'ȡһ��ҽ����ID
        For i = .FixedRows To .Rows - 1
            If Val(.TextMatrix(i, colѡ��)) <> 0 Then
                lngID = Val(.TextMatrix(i, colID))
                If lngID <> 0 Then

                    'ѡ���Ʋ���
                    If InStr(strIDs & ",", "," & lngID & ",") = 0 Then
                        strIDs = strIDs & "," & lngID
                    End If
                    
                    '�л���Ч����
                    If .TextMatrix(i, col��Ч) <> .Cell(flexcpData, i, col��Ч) Then
                        If InStr(strAlter & ",", "," & lngID & ",") = 0 Then
                            strAlter = strAlter & "," & lngID
                        End If
                    End If
                End If
            End If
        Next
        strAlter = Mid(strAlter, 2)
        strIDs = Mid(strIDs, 2)
        If strIDs = "" Then
            MsgBox "��ѡ��Ҫ���Ƶ�ҽ����", vbInformation, gstrSysName
            Exit Sub
        End If
    End With
    
    mstrAlter = strAlter
    mstrIDs = strIDs
    mstr�Һŵ� = mtmp�Һŵ�
    mlng��ҳID = mtmp��ҳID
    
    Unload Me
End Sub

Private Sub cmdPati_Click()
    If mstr�Һŵ� <> "" Then
        lvwPati.ListItems("_" & mlng����ID & "_" & mstr�Һŵ�).Selected = True
    Else
        lvwPati.ListItems("_" & mlng����ID & "_" & mlng��ҳID).Selected = True
    End If
    lvwPati.SelectedItem.EnsureVisible
    lvwPati.Left = txtPati.Left + fraPati.Left
    lvwPati.Top = txtPati.Top + txtPati.Height + fraPati.Top
    lvwPati.Height = vsAdvice.Height - 300
    lvwPati.ZOrder
    lvwPati.Visible = True
    lvwPati.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        Call cmdHelp_Click
    ElseIf KeyCode = vbKeyEscape Then
        If lvwPati.Visible Then
            lvwPati.Visible = False
        Else
            Unload Me
        End If
    ElseIf KeyCode = vbKeyF4 Or KeyCode = vbKeyDown Then
        If Not (KeyCode = vbKeyDown And Shift <> vbAltMask) Then
            If Me.ActiveControl Is txtPati Then
                If cmdPati.Visible And cmdPati.Enabled Then cmdPati_Click
            End If
        End If
    ElseIf KeyCode = vbKeyA And Shift = vbCtrlMask Then
        Call cmdALL_Click
    ElseIf KeyCode = vbKeyR And Shift = vbCtrlMask Then
        Call cmdClear_Click
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub Form_Load()
    Dim strLvw As String
    
    mblnҽ������ = Val(zlDatabase.GetPara("ҽ��ҽ����������", glngSys, pסԺҽ������)) <> 0
    lvwPati.SmallIcons = frmIcons.imgPati
    If mstr�Һŵ� <> "" Then
        strLvw = "����,1000,0,1;�����,1000,0,1;�Ա�,600,0,1;����,600,0,1;����,1000,0,1;����ʱ��,1620,0,1;����״̬,1000,2,1;�Һŵ�,1000,0,1"
    Else
        strLvw = "����,1000,0,1;סԺ��,1000,0,1;����,630,0,1;סԺҽʦ,1000,0,1;�Ա�,600,0,1;����,600,0,1;����,1000,1,0;�ѱ�,850,0,1;����ȼ�,1150,0,1;��Ժ����,1620,0,1;��Ժ����,1620,0,1;���ʽ,1500,0,1;��������,1500,0,1"
    End If
    Call InitAdviceTable
    Call zlControl.LvwSelectColumns(lvwPati, strLvw, True)
    Call RestoreWinState(Me, App.ProductName, IIF(mstr�Һŵ� <> "", 1, 2))
    If mlng��ҳID <> 0 Then
        vsAdvice.FrozenCols = col��Ч + 1
    Else
        vsAdvice.FrozenCols = colѡ�� + 1
    End If
    
    If mstr�Һŵ� <> "" And mlngǰ��ID = 0 Then
        Call InitSelectTime
    End If
    
    
    '�б�ѡ������¼��ʽ��ֻ���������
    If mlngǰ��ID <> 0 Or mstr�Һŵ� = "" And mlng��ҳID <> 0 Then
        vsPati.Visible = False
        fraLR.Visible = False
        lblDiag.Visible = False
        picDiag.Visible = False
        mbln������Ȩ�� = InStr(GetTsPrivs(pסԺҽ���´�), ";�´�����ҩ��;") = 0
        mbln������Ȩ�� = InStr(GetTsPrivs(pסԺҽ���´�), ";�´ﶾ��ҩ��;") = 0
        mbln������Ȩ�� = InStr(GetTsPrivs(pסԺҽ���´�), ";�´ﾫ��ҩ��;") = 0
        mbln������Ȩ�� = InStr(GetTsPrivs(pסԺҽ���´�), ";�´����ҩ��;") = 0
        mint��Դ = 2
        cboQX.Visible = True
        cboQX.Clear
        cboQX.AddItem "����"
        cboQX.AddItem "����"
        cboQX.AddItem "����"
        cboQX.ListIndex = 0
        mlngPre��Ч = -1
    Else
        lblTime.Visible = False
        cboTime.Visible = False
        mbln������Ȩ�� = InStr(GetTsPrivs(p����ҽ���´�), ";�´�����ҩ��;") = 0
        mbln������Ȩ�� = InStr(GetTsPrivs(p����ҽ���´�), ";�´ﶾ��ҩ��;") = 0
        mbln������Ȩ�� = InStr(GetTsPrivs(p����ҽ���´�), ";�´ﾫ��ҩ��;") = 0
        mbln������Ȩ�� = InStr(GetTsPrivs(p����ҽ���´�), ";�´����ҩ��;") = 0
        mint��Դ = 1
        cboQX.Visible = False
    End If
    Call LoadPatients
    vsPati.Tag = "���ز���"
    Call LoadPatiTime
    mstrIDs = ""
End Sub

Private Sub cboQX_Click()
'���ܣ�����ҽ��
    If Not Me.Visible Then Exit Sub
    If cboQX.ListIndex <> mlngPre��Ч Then
        mlngPre��Ч = cboQX.ListIndex
        Call LoadAdvice
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    fraPati.Top = -75
    fraPati.Left = 0
    fraPati.Width = Me.ScaleWidth
    If mstr�Һŵ� <> "" And mlngǰ��ID = 0 Then
        fraPati.Height = 900
        optType(opt���ھ���).Visible = True
        optType(opt��ɾ���).Visible = True
    Else
        fraPati.Height = 600
    End If
    
    If fraPati.Width - cboBaby.Width - 200 > 7500 Then
        cboBaby.Left = fraPati.Width - cboBaby.Width - 150
        lblBaby.Left = cboBaby.Left - lblBaby.Width - 30
    End If
    
    picDiag.Top = lblDiag.Top - 40
    picDiag.Left = lblDiag.Left + 900
    
    vsPati.Left = 0
    vsPati.Top = fraPati.Top + fraPati.Height
    vsPati.Height = Me.ScaleHeight - vsAdvice.Top - cmdOK.Height * 1.6
    
    fraLR.Left = vsPati.Width
    fraLR.Top = vsPati.Top
    fraLR.Height = vsPati.Height
    
    vsAdvice.Left = IIF(vsPati.Visible, vsPati.Width + fraLR.Width, 0)
    vsAdvice.Top = fraPati.Top + fraPati.Height
    vsAdvice.Width = Me.ScaleWidth - IIF(vsPati.Visible, vsPati.Width + fraLR.Width, 0)
    vsAdvice.Height = vsPati.Height
        
    cmdHelp.Top = Me.ScaleHeight - cmdAll.Height * 1.3
    cmdAll.Top = cmdHelp.Top
    cmdClear.Top = cmdAll.Top
    cmdOK.Top = cmdAll.Top
    cmdCancel.Top = cmdAll.Top
    
    If Me.ScaleWidth - cmdCancel.Width - (cmdHelp.Left + cmdHelp.Width / 3) < 5000 Then
        cmdCancel.Left = 5000
    Else
        cmdCancel.Left = Me.ScaleWidth - cmdCancel.Width - (cmdHelp.Left + cmdHelp.Width / 3)
    End If
    cmdOK.Left = cmdCancel.Left - cmdOK.Width
    
    cboQX.Left = cmdHelp.Left
    cboQX.Top = cmdHelp.Top - 350
    cboQX.Width = cmdHelp.Width
    Me.Refresh
End Sub
 
Private Sub optType_Click(Index As Integer)

    cboFinTim.Enabled = optType(opt��ɾ���).value
    If Not (InStr(";" & mMainPrivs & ";", ";��������;") > 0 And cboFinTim.Enabled) Then cboFinTim.Enabled = False
    
    lblTim.Visible = optType(opt��ɾ���).value
    cboFinTim.Visible = lblTim.Visible
    
    Call LoadPatients
    Call LoadPatiTime
End Sub

Private Sub InitSelectTime()
    Dim datCurr As Date, intStart As Integer, intDay As Integer
    Dim blnSetPar As Boolean
    
    With cboFinTim
        .Clear
        .AddItem "һ����"
        .ItemData(.NewIndex) = 7
        .AddItem "15����"
        .ItemData(.NewIndex) = 15
        .AddItem "30����"
        .ItemData(.NewIndex) = 30
        .AddItem "60����"
        .ItemData(.NewIndex) = 60
        .AddItem "[ָ��...]"
        .ItemData(.NewIndex) = -1
        
        .Tag = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
    End With
    
    datCurr = CDate(cboFinTim.Tag)

    blnSetPar = InStr(";" & mMainPrivs & ";", ";��������;") > 0
    intStart = Val(zlDatabase.GetPara("���ﲡ�˽������", glngSys, p����ҽ��վ, "0", Array(lblTim, cboFinTim), blnSetPar))
    intDay = Val(zlDatabase.GetPara("���ﲡ�˿�ʼ���", glngSys, p����ҽ��վ, "7", Array(lblTim, cboFinTim), blnSetPar))
    
    mdtOutEnd = Format(datCurr + intStart, "yyyy-MM-dd 23:59:59")
    mdtOutBegin = Format(mdtOutEnd - intDay, "yyyy-MM-dd 00:00:00")
     
    cboFinTim.ToolTipText = Format(mdtOutBegin, "yyyy-MM-dd  00:00") & " - " & Format(mdtOutEnd, "yyyy-MM-dd 23:59")
    lblTim.ToolTipText = cboFinTim.ToolTipText
    
    If intStart = 0 Then
        Select Case intDay
        Case 7
            mintOutPreTime = 0
        Case 15
            mintOutPreTime = 1
        Case 30
            mintOutPreTime = 2
        Case 60
            mintOutPreTime = 3
        Case Else
            mintOutPreTime = 4
        End Select
    Else
        mintOutPreTime = 4
    End If
    
    Call cbo.SetIndex(cboFinTim.hwnd, mintOutPreTime)
End Sub

Private Sub picDiag_DblClick()
    MsgBox picDiag.Tag, vbInformation, Me.Caption
End Sub

Private Sub picDiag_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    zlCommFun.ShowTipInfo picDiag.hwnd, picDiag.Tag, True
    If X >= 0 And X <= picDiag.Width And Y >= 0 And Y <= picDiag.Height Then
        SetCapture picDiag.hwnd
    Else
        ReleaseCapture
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName, IIF(mstr�Һŵ� <> "", 1, 2))
End Sub

Private Sub fraLR_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If vsPati.Width + X < 1000 Or vsAdvice.Width - X < 1000 Then Exit Sub
        fraLR.Left = fraLR.Left + X
        vsPati.Width = vsPati.Width + X
        vsAdvice.Left = vsAdvice.Left + X
        vsAdvice.Width = vsAdvice.Width - X
        
        Me.Refresh
    End If
End Sub

Private Sub lvwPati_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Call zlControl.LvwSortColumn(lvwPati, ColumnHeader.Index)
End Sub

Private Sub lvwPati_DblClick()
    If mblnItem Then Call lvwPati_KeyPress(13)
End Sub

Private Sub lvwPati_ItemClick(ByVal Item As MSComctlLib.ListItem)
    mblnItem = True
End Sub

Private Sub lvwPati_KeyPress(KeyAscii As Integer)
    Dim lng����ID As Long, lng��ҳID As Long, strNO As String
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Not lvwPati.SelectedItem Is Nothing Then
            lng����ID = Val(Split(Mid(lvwPati.SelectedItem.Key, 2), "_")(0))
            If mstr�Һŵ� <> "" Then
                strNO = Split(Mid(lvwPati.SelectedItem.Key, 2), "_")(1)
                If lng����ID = mlng����ID And strNO = mstr�Һŵ� Then
                    lvwPati.Visible = False
                    vsAdvice.SetFocus: Exit Sub
                End If
                With lvwPati.SelectedItem
                    mlng����ID = lng����ID
                    mstr�Һŵ� = strNO
                    mblnMoved = zlDatabase.NOMoved("���˹Һż�¼", strNO)
                    
                    txtPati.Text = .Text
                    txtPati.ForeColor = .ForeColor
                End With
            Else
                lng��ҳID = Val(Split(Mid(lvwPati.SelectedItem.Key, 2), "_")(1))
                If lng����ID = mlng����ID And lng��ҳID = mlng��ҳID Then
                    lvwPati.Visible = False
                    vsAdvice.SetFocus: Exit Sub
                End If
                With lvwPati.SelectedItem
                    mlng����ID = lng����ID
                    mlng��ҳID = lng��ҳID
                    mblnMoved = Val(.Tag) = 1
                    
                    txtPati.Text = .Text
                    txtPati.ForeColor = .ForeColor
                End With
            End If
            lvwPati.Visible = False
            
            vsPati.Tag = "���ز���"
            
            Call LoadPatiTime
            
            If vsPati.Visible Then
                vsPati.SetFocus
            Else
                vsAdvice.SetFocus
            End If
        End If
    End If
End Sub

Private Sub lvwPati_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mblnItem = False
End Sub

Private Sub lvwPati_Validate(Cancel As Boolean)
    lvwPati.Visible = False
End Sub

Private Sub txtPati_GotFocus()
    Call zlControl.TxtSelAll(txtPati)
End Sub

Private Function RowInһ����ҩ(ByVal lngRow As Long, lngBegin As Long, lngEnd As Long) As Boolean
'���ܣ��ж�ָ�����Ƿ���һ����ҩ�ķ�Χ��,�����,ͬʱ�����кŷ�Χ
    Dim i As Long, blnTmp As Boolean
    With vsAdvice
        If .TextMatrix(lngRow, col�������) = "" Then Exit Function
        If InStr(",5,6,", .TextMatrix(lngRow, col�������)) = 0 Then Exit Function
        If Val(.TextMatrix(lngRow - 1, col���ID)) = Val(.TextMatrix(lngRow, col���ID)) Then
            blnTmp = True
        ElseIf lngRow + 1 <= .Rows - 1 Then
            If Val(.TextMatrix(lngRow + 1, col���ID)) = Val(.TextMatrix(lngRow, col���ID)) Then
                blnTmp = True
            End If
        End If
        If blnTmp Then
            lngBegin = lngRow
            For i = lngRow - 1 To .FixedRows Step -1
                If Val(.TextMatrix(i, col���ID)) = Val(.TextMatrix(lngRow, col���ID)) Then
                    lngBegin = i
                Else
                    Exit For
                End If
            Next
            lngEnd = lngRow
            For i = lngRow + 1 To .Rows - 1
                If Val(.TextMatrix(i, col���ID)) = Val(.TextMatrix(lngRow, col���ID)) Then
                    lngEnd = i
                Else
                    Exit For
                End If
            Next
        End If
        RowInһ����ҩ = blnTmp
    End With
End Function

Private Function LoadAdvice() As Boolean
'���ܣ���ȡ��ǰ����ָ����ҽ��
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long, j As Long
    Dim lng����ID As Long, intӤ�� As Integer
    Dim strDepartments As String
    Dim strWhere��Ч As String
    
    If cboTime.ListIndex = -1 Then Exit Function
    If cboBaby.ListIndex = -1 Then Exit Function
    lng����ID = cboTime.ItemData(cboTime.ListIndex)
    intӤ�� = cboBaby.ItemData(cboBaby.ListIndex)
    
    If cboQX.Visible Then
        Select Case cboQX.ListIndex
        Case 1
            strWhere��Ч = " and a.ҽ����Ч=0"
        Case 2
            strWhere��Ч = " and a.ҽ����Ч=1"
        End Select
    End If
    
    On Error GoTo errH
    
    '�ſ������Ͳ������ڵ�����
    strSQL = "Select Distinct (Select Count(1) From �������ÿ��� Where ��ĿID=b.ID) as ���ÿ�����,g.����id as ���ÿ���ID,A.ID,A.���,A.���ID,A.ҽ����Ч,A.��ʼִ��ʱ��,A.������ĿID," & _
        " A.ҽ������,A.ִ������,A.��������,A.ִ��Ƶ��,A.ҽ������,B.����Ӧ��,B.��������,Nvl(C.����,Decode(Nvl(A.ִ������,0),5,'-')) as ִ�п��� ,A.ִ��ʱ�䷽��,A.�շ�ϸĿID," & _
        " A.�걾��λ,A.��鷽��,B.���,a.�������,B.����,B.���㵥λ,A.�ܸ����� as ����,E.�����װ,E.���ﵥλ,E.סԺ��װ,E.סԺ��λ," & _
        " D.���㵥λ as ɢװ��λ,B.����ʱ��,B.�������,D.����ʱ�� as �շѳ���,D.������� as �շѷ���,h.�������,h.��ֵ����,Nvl(b.�����Ա�, 0) As �Ա�,b.���� as ��Ŀ����,d.���� as �շ�����"
    If mstr�Һŵ� <> "" Then
        strSQL = strSQL & ",A.�Һŵ�" & _
            " From ����ҽ����¼ A,������ĿĿ¼ B,���ű� C,�շ���ĿĿ¼ D,ҩƷ��� E,���˹Һż�¼ R,�������ÿ��� G,ҩƷ���� H" & _
            " Where A.������ĿID=B.ID(+) And A.ִ�п���ID=C.ID(+) And A.�շ�ϸĿID=D.ID(+) And b.id=g.��Ŀid(+) And h.ҩ��ID(+)=b.ID And g.����ID(+)=[4] And A.�շ�ϸĿID=E.ҩƷID(+)" & _
            " And Nvl(A.ִ�б��,0)<>-1 And A.ҽ��״̬ Not IN(2,4) And A.��ʼִ��ʱ�� is Not Null And Nvl(A.ҽ��״̬,0)<>-1 And A.������Դ=1" & _
            " And A.����ID+0=[1] And A.�Һŵ�=R.NO And R.ID=[2] And Nvl(A.Ӥ��,0)=[3]" & strWhere��Ч & _
            " Order by A.���"
    Else
        strSQL = strSQL & _
            " From ����ҽ����¼ A,������ĿĿ¼ B,���ű� C,�շ���ĿĿ¼ D,ҩƷ��� E,�������ÿ��� G,ҩƷ���� H" & _
            " Where A.������ĿID=B.ID(+) And A.ִ�п���ID=C.ID(+) And A.�շ�ϸĿID=D.ID(+)  And b.id=g.��Ŀid(+) And h.ҩ��ID(+)=b.ID And g.����ID(+)=[4] And A.�շ�ϸĿID=E.ҩƷID(+)" & _
            " And Nvl(A.ִ�б��,0)<>-1 And A.ҽ��״̬ Not IN(2,4) And A.��ʼִ��ʱ�� is Not Null And Nvl(A.ҽ��״̬,0)<>-1 And A.������Դ=2" & _
            " And A.����ID=[1] And A.��ҳID=[2] And Nvl(A.Ӥ��,0)=[3]" & strWhere��Ч & _
            " Order by A.���"
    End If
    If mblnMoved Then
        strSQL = Replace(strSQL, "����ҽ����¼", "H����ҽ����¼")
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, lng����ID, intӤ��, mlng���˿���id)
    With vsAdvice
        .Redraw = flexRDNone
        .Rows = .FixedRows '����������
        If rsTmp.EOF Then
            .Rows = .FixedRows + 1
        Else
            .Rows = .FixedRows + rsTmp.RecordCount
            For i = 1 To rsTmp.RecordCount
                If i = 1 Then
                    If mstr�Һŵ� <> "" Then
                        mtmp�Һŵ� = rsTmp!�Һŵ�
                    Else
                        mtmp��ҳID = lng����ID
                    End If
                End If
            
                .TextMatrix(i, colѡ��) = 0
                .TextMatrix(i, colID) = rsTmp!ID
                .TextMatrix(i, col���ID) = NVL(rsTmp!���ID)
                .TextMatrix(i, col�������) = NVL(rsTmp!���, "*")
                .TextMatrix(i, col������ĿID) = NVL(rsTmp!������ĿID)
                .TextMatrix(i, col�շ�ϸĿID) = NVL(rsTmp!�շ�ϸĿID)
                .TextMatrix(i, col��Ч) = IIF(NVL(rsTmp!ҽ����Ч, 0) = 0, "����", "����")
                .Cell(flexcpData, i, col��Ч) = .TextMatrix(i, col��Ч)
                .TextMatrix(i, colʱ��) = Format(rsTmp!��ʼִ��ʱ��, "yyyy-MM-dd HH:mm")
                .TextMatrix(i, col����) = rsTmp!ҽ������
                If .TextMatrix(i, col�������) = "D" And NVL(rsTmp!�걾��λ) <> "" Then
                    .Cell(flexcpData, i, col����) = "[" & NVL(rsTmp!�걾��λ) & "]" & NVL(rsTmp!��鷽��) '����걾
                Else
                    .Cell(flexcpData, i, col����) = NVL(rsTmp!�걾��λ)
                End If
                .TextMatrix(i, col����) = FormatEx(NVL(rsTmp!��������), 4)
                .Cell(flexcpData, i, col����) = NVL(rsTmp!����)
                If Val(rsTmp!���ÿ����� & "") > 0 And rsTmp!���ÿ���ID & "" = "" Then
                    .TextMatrix(i, col�Ƿ�����) = "1"
                End If
                .TextMatrix(i, col�������) = NVL(rsTmp!�������)
                .TextMatrix(i, col��ֵ����) = NVL(rsTmp!��ֵ����)
                .TextMatrix(i, col�Ա�) = Decode(Val(rsTmp!�Ա�), 0, "δ֪", 1, "��", 2, "Ů")
                If Not IsNull(rsTmp!��������) Then
                    If NVL(rsTmp!���) = "4" Then
                        .TextMatrix(i, col������λ) = NVL(rsTmp!ɢװ��λ)
                    Else
                        .TextMatrix(i, col������λ) = NVL(rsTmp!���㵥λ)
                    End If
                End If
                If InStr(",5,6,", NVL(rsTmp!���, "*")) > 0 Then
                    If mstr�Һŵ� <> "" Then
                        If Not IsNull(rsTmp!����) And Not IsNull(rsTmp!�����װ) Then
                            .TextMatrix(i, col����) = FormatEx(rsTmp!���� / rsTmp!�����װ, 5)
                        End If
                        If NVL(rsTmp!ҽ����Ч, 0) = 1 Then
                            .TextMatrix(i, col������λ) = NVL(rsTmp!���ﵥλ)
                        End If
                    Else
                        If Not IsNull(rsTmp!����) And Not IsNull(rsTmp!סԺ��װ) Then
                            .TextMatrix(i, col����) = FormatEx(rsTmp!���� / rsTmp!סԺ��װ, 5)
                        End If
                        If NVL(rsTmp!ҽ����Ч, 0) = 1 Then
                            .TextMatrix(i, col������λ) = NVL(rsTmp!סԺ��λ)
                        End If
                    End If
                Else
                    If Not IsNull(rsTmp!����) Then
                        .TextMatrix(i, col����) = FormatEx(rsTmp!����, 5)
                    End If
                    If NVL(rsTmp!ҽ����Ч, 0) = 1 Then
                        If NVL(rsTmp!���) = "4" Then
                            .TextMatrix(i, col������λ) = NVL(rsTmp!ɢװ��λ)
                        Else
                            .TextMatrix(i, col������λ) = NVL(rsTmp!���㵥λ)
                        End If
                    End If
                End If
                
                .TextMatrix(i, colƵ��) = NVL(rsTmp!ִ��Ƶ��)
                .TextMatrix(i, col����) = NVL(rsTmp!ҽ������)
                .TextMatrix(i, colִ��ʱ��) = NVL(rsTmp!ִ��ʱ�䷽��)
                .TextMatrix(i, colִ�п���) = NVL(rsTmp!ִ�п���)
                .TextMatrix(i, col����Ӧ��) = NVL(rsTmp!����Ӧ��)
                .TextMatrix(i, col��������) = NVL(rsTmp!��������)
                .TextMatrix(i, col��Ŀ����) = NVL(rsTmp!��Ŀ����)
                .TextMatrix(i, col�շ�����) = NVL(rsTmp!�շ�����)
                
                '���������ؼ��÷���ʾ
                If InStr(",C,D,F,G,E,", NVL(rsTmp!���, "*")) > 0 And Not IsNull(rsTmp!���ID) Then
                    .RowHidden(i) = True
                    
                    '��Ѫ;��
                    If rsTmp!��� = "E" And .TextMatrix(i - 1, col�������) = "K" _
                        And rsTmp!���ID = Val(.TextMatrix(i - 1, colID)) Then
                        .TextMatrix(i - 1, col�÷�) = rsTmp!����
                    End If
                ElseIf NVL(rsTmp!���) = "7" Then
                    .RowHidden(i) = True
                ElseIf NVL(rsTmp!���) = "E" And IsNull(rsTmp!���ID) _
                    And Val(.TextMatrix(i - 1, col���ID)) = rsTmp!ID _
                    And InStr(",5,6,", .TextMatrix(i - 1, col�������)) > 0 Then
                    '��ҩ;��
                    .RowHidden(i) = True
                    '��ʾ��ҩ;��
                    For j = i - 1 To .FixedRows Step -1
                        If Val(.TextMatrix(j, col���ID)) = rsTmp!ID Then
                            .TextMatrix(j, col�÷�) = rsTmp!���� & rsTmp!ҽ������
                        Else
                            Exit For
                        End If
                    Next
                ElseIf NVL(rsTmp!���) = "E" And IsNull(rsTmp!���ID) _
                    And Val(.TextMatrix(i - 1, col���ID)) = rsTmp!ID _
                    And InStr(",7,E,C,", .TextMatrix(i - 1, col�������)) > 0 Then
                    '��ҩ�÷������ɼ�����
                    .TextMatrix(i, col�÷�) = rsTmp!����
                    
                    '��ҩ������ִ�п���
                    For j = i - 1 To .FixedRows Step -1
                        If Val(.TextMatrix(j, col���ID)) = rsTmp!ID Then
                            If InStr(",7,C,", .TextMatrix(j, col�������)) > 0 Then
                                .TextMatrix(i, colִ�п���) = .TextMatrix(j, colִ�п���)
                                Exit For
                            End If
                        Else
                            Exit For
                        End If
                    Next
                    
                    '��ҩ����
                    If .TextMatrix(i - 1, col�������) <> "C" Then
                        .TextMatrix(i, col������λ) = "��"
                    End If
                End If
                
                '��ǰ������г����򲻷������Ŀ
                If Not IsNull(rsTmp!������ĿID) Then
                    If Not (IsNull(rsTmp!����ʱ��) Or Format(NVL(rsTmp!����ʱ��), "yyyy-MM-dd") = "3000-01-01") Then
                        .RowData(i) = con��Ŀ����
                    ElseIf Not (NVL(rsTmp!�������, 0) = 3 Or NVL(rsTmp!�������, 0) = IIF(mstr�Һŵ� <> "", 1, 2)) Then
                        .RowData(i) = con��Ŀ����
                    ElseIf Not IsNull(rsTmp!�շ�ϸĿID) Then
                        '��ҩƷ,ͬʱҪ�жϵ��շ���ĿĿ¼
                        If Not (IsNull(rsTmp!�շѳ���) Or Format(NVL(rsTmp!�շѳ���), "yyyy-MM-dd") = "3000-01-01") Then
                            .RowData(i) = con�շѳ���
                        ElseIf Not (NVL(rsTmp!�շѷ���, 0) = 3 Or NVL(rsTmp!�շѷ���, 0) = IIF(mstr�Һŵ� <> "", 1, 2)) Then
                            .RowData(i) = con�շѷ���
                        End If
                    ElseIf rsTmp!��� <> rsTmp!������� Then
                        .RowData(i) = con��Ŀ���
                    End If
                End If
                
                If gblnStock Then
                    '�жϷ�Ժ��ִ��ҩƷ�Ƿ��п��
                    strDepartments = ""
                    If mlng���˿���id <> 0 And Val(rsTmp!ִ������ & "") <> 5 And InStr(",5,6,7,", rsTmp!��� & "") > 0 And Val(rsTmp!�շ�ϸĿID & "") <> 0 Then
                        strDepartments = Get����ҩ��IDs(rsTmp!��� & "", NVL(rsTmp!������ĿID), Val(rsTmp!�շ�ϸĿID & ""), mlng���˿���id, mint��Դ)
                        '�жϿ���Ƿ��������
                        If strDepartments <> "" Then
                            If GetStock(Val(rsTmp!�շ�ϸĿID & ""), , mint��Դ, strDepartments, CDbl(Val(.TextMatrix(i, col����)))) = 0 Then
                                .Cell(flexcpData, i, col�Ƿ�����) = 1
                            End If
                        End If
                    End If
                End If

                rsTmp.MoveNext
            Next
        End If
        If mlng��ҳID <> 0 Then
            .Cell(flexcpBackColor, .FixedRows, colѡ��, .Rows - 1, col��Ч) = COLEditBackColor      'ǳ��
        End If
        
        .Row = .FixedRows: .Col = .FixedCols
        .AutoSize col����
        .ColHidden(col��Ч) = mstr�Һŵ� <> ""
        .Redraw = flexRDDirect
    End With
    LoadAdvice = True
    Exit Function
errH:
    vsAdvice.Redraw = flexRDDirect
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub txtPati_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lvwList As ListItem, strTmp As String
    strTmp = lvwPati.SelectedItem.Text
    If KeyCode = 13 Then
        KeyCode = 0
        If IsNumeric(Trim(txtPati.Text)) And InStr(Trim(txtPati.Text), ".") <= 0 _
                And InStr(Trim(txtPati.Text), "-") <= 0 And InStr(Trim(txtPati.Text), "+") <= 0 Then
            Set lvwList = lvwPati.FindItem(Trim(txtPati.Text), 1)
        Else
            Set lvwList = lvwPati.FindItem(Trim(txtPati.Text))
        End If
        If lvwList Is Nothing Then
            MsgBox "û���ҵ��ò��ˡ�", vbInformation, Me.Caption
            txtPati.Text = strTmp
            txtPati.SelStart = 0
            txtPati.SelLength = Len(txtPati.Text)
        Else
            lvwList.Selected = True
            Call lvwPati_KeyPress(13)
        End If
    End If
End Sub

Private Sub vsAdvice_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Dim lngW As Long
    
    If Col = col���� Then
        vsAdvice.AutoSize Col
    ElseIf Row = -1 Then
        lngW = Me.TextWidth(vsAdvice.TextMatrix(vsAdvice.FixedRows - 1, Col) & "A")
        If vsAdvice.ColWidth(Col) < lngW Then
            vsAdvice.ColWidth(Col) = lngW
        ElseIf vsAdvice.ColWidth(Col) > vsAdvice.Width * 0.5 Then
            vsAdvice.ColWidth(Col) = vsAdvice.Width * 0.5
        End If
    End If
End Sub

Private Sub vsAdvice_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = colѡ�� Then Cancel = True
End Sub

Private Sub vsAdvice_DblClick()
    Call vsAdvice_KeyPress(32)
End Sub

Private Sub vsAdvice_DrawCell(ByVal hDC As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
    Dim lngLeft As Long, lngRight As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim vRect As RECT
    
    With vsAdvice
        '����һ����ҩ������еı��߼�����
        lngLeft = col��Ч: lngRight = colʱ��
        If Not Between(Col, lngLeft, lngRight) Then
            lngLeft = colƵ��: lngRight = col�÷�
            If Not Between(Col, lngLeft, lngRight) Then Exit Sub
        End If
        
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
            If Col = lngLeft And lngLeft = col��Ч Then
                SetBkColor hDC, OS.SysColor2RGB(.Cell(flexcpBackColor, Row, lngLeft))
            Else
                SetBkColor hDC, OS.SysColor2RGB(.BackColor)
            End If
        End If
        ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0
        Done = True
    End With
End Sub

Private Sub vsAdvice_GotFocus()
    vsAdvice.BackColorSel = &HFFCC99
End Sub

Private Sub vsAdvice_KeyPress(KeyAscii As Integer)
    Dim i As Long
    
    With vsAdvice
        If KeyAscii = 32 Then
            If .Col <> colѡ�� Then
                KeyAscii = 0
                If .TextMatrix(.Row, colѡ��) = 0 Then
                    If CheckCanSelGroup(.Row, True) Then
                        If .Col = col��Ч And mlng��ҳID <> 0 And CanAlterType(.Row) Then
                            Call AlterGroupType(.Row)
                        End If
                        Call SelGroup(.Row, -1)
                    End If
                Else
                    If .Cell(flexcpFontBold, .Row, .Col) Then
                        Call AlterGroupType(.Row)
                    End If
                    Call SelGroup(.Row, 0)
                End If
            End If
        ElseIf KeyAscii = 13 Then
            KeyAscii = 0
            For i = .Row + 1 To .Rows - 1
                If Not .RowHidden(i) Then
                    .Row = i
                    Call .ShowCell(.Row, .Col)
                    Exit For
                End If
            Next
            If i > .Rows - 1 Then
                Call zlCommFun.PressKey(vbKeyTab)
            End If
        End If
    End With
End Sub

Private Sub vsAdvice_LostFocus()
    vsAdvice.BackColorSel = &HC0C0C0
End Sub

Private Sub vsAdvice_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsAdvice
        If Col <> colѡ�� Then
            Cancel = True
        ElseIf Val(.TextMatrix(vsAdvice.Row, colID)) = 0 Then
            Cancel = True
        Else
            If .TextMatrix(Row, colѡ��) <> 0 Then
                Call SelGroup(Row, 0)
            Else
                If CheckCanSelGroup(Row, True) Then
                    Call SelGroup(Row, -1)
                End If
            End If
            '�Ѿ������жϺ�ѡ�񣬲��败��AfterEdit�¼�
            Cancel = True
        End If
    End With
End Sub

Private Function CanAlterType(ByVal lngRow As Long) As Boolean
'���ܣ��ж�ָ����ҽ���Ƿ�����л���Ч
'������lngRow=�ɼ���ҽ����
'˵���������л���Ч��������
'   1.�ɳ�����ִ��Ƶ��=0(��ѡƵ��),2(������)
'   2.��������ִ��Ƶ��=0(��ѡƵ��),1(һ����);ҩƷ����ָ���˹��
    Dim rsMore As New ADODB.Recordset
    Dim strSQL As String, strType As String, i As Long
    Dim lngBegin As Long, lngEnd As Long
    
    With vsAdvice
        If Val(.TextMatrix(lngRow, colID)) = 0 Then
            CanAlterType = True: Exit Function
        ElseIf Val(.TextMatrix(lngRow, col������ĿID)) = 0 Then
            '��������Ŀ����л�
            CanAlterType = True: Exit Function
        ElseIf RowIn�䷽��(lngRow) Then
            '��ҩ�䷽�̶������л�
            CanAlterType = True: Exit Function
        ElseIf RowIn������(lngRow) Then
            '�����Լ�����Ϊ׼�ж�
            lngRow = .FindRow(.TextMatrix(lngRow, colID), , col���ID)
            If lngRow = -1 Then Exit Function
        End If
    
        strType = IIF(.TextMatrix(lngRow, col��Ч) = "����", "����", "����")
        
        '��ԭʼƵ��Ϊ׼�ж�:��Ϊ��ѡ��Ƶ�ʵĿ�����ȱ��һ����
        strSQL = "Select ִ��Ƶ�� From ������ĿĿ¼ Where ID=[1]"
        Set rsMore = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.TextMatrix(lngRow, col������ĿID)))
        
        If strType = "����" Then
            If InStr(",0,2,", NVL(rsMore!ִ��Ƶ��, 0)) = 0 Then Exit Function
        Else
            If InStr(",0,1,", NVL(rsMore!ִ��Ƶ��, 0)) = 0 Then Exit Function
            If InStr(",5,6,", .TextMatrix(lngRow, col�������)) > 0 Then
                Call GetRowScope(lngRow, lngBegin, lngEnd)
                For i = lngBegin To lngEnd
                    If InStr(",5,6,", .TextMatrix(i, col�������)) > 0 Then
                        If Val(.TextMatrix(i, col�շ�ϸĿID)) = 0 Then Exit Function
                    End If
                Next
            End If
        End If
    End With
    CanAlterType = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function RowIn������(ByVal lngRow As Long) As Boolean
'���ܣ��ж�ָ�����Ƿ����ڼ�������е�һ��
'˵���������е�ǰ�Ƿ�����
    If lngRow = -1 Then Exit Function
    If vsAdvice.TextMatrix(lngRow, colID) = 0 Then Exit Function
    
    With vsAdvice
        If .TextMatrix(lngRow, col�������) = "E" And Val(.TextMatrix(lngRow, col���ID)) = 0 Then
            '�ɼ�������
            If .TextMatrix(lngRow - 1, col�������) = "C" _
                And Val(.TextMatrix(lngRow - 1, col���ID)) = .TextMatrix(lngRow, colID) Then
                RowIn������ = True: Exit Function
            End If
        ElseIf .TextMatrix(lngRow, col�������) = "C" And Val(.TextMatrix(lngRow, col���ID)) <> 0 Then
            '������Ŀ��
            RowIn������ = True: Exit Function
        End If
    End With
End Function

Private Function RowIn�䷽��(ByVal lngRow As Long) As Boolean
'���ܣ��ж�ָ�����Ƿ�������ҩ�䷽�е�һ��
'˵���������е�ǰ�Ƿ�����
    If lngRow = -1 Then Exit Function
    If vsAdvice.TextMatrix(lngRow, colID) = 0 Then Exit Function
    
    With vsAdvice
        If .TextMatrix(lngRow, col�������) = "E" Then
            If Val(.TextMatrix(lngRow, col���ID)) = 0 Then
                '�÷���
                If Val(.TextMatrix(lngRow - 1, col���ID)) = .TextMatrix(lngRow, colID) _
                    And .TextMatrix(lngRow - 1, col�������) = "E" Then
                    RowIn�䷽�� = True: Exit Function
                End If
            Else
                '�巨��
                If .TextMatrix(lngRow - 1, col�������) = "7" _
                    And Val(.TextMatrix(lngRow - 1, col���ID)) = Val(.TextMatrix(lngRow, col���ID)) Then
                    RowIn�䷽�� = True: Exit Function
                End If
            End If
        ElseIf .TextMatrix(lngRow, col�������) = "7" And Val(.TextMatrix(lngRow, col���ID)) <> 0 Then
            '��ҩ��
            RowIn�䷽�� = True: Exit Function
        End If
    End With
End Function

Private Sub vsPati_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim i As Integer
    
    If vsPati.Redraw = flexRDNone Then Exit Sub
    If NewRow = -1 Then Exit Sub
    If NewRow = OldRow Then Exit Sub
    
    For i = 0 To cboTime.ListCount - 1
        If cboTime.ItemData(i) = vsPati.RowData(NewRow) Then
            cboTime.ListIndex = i: Exit For
        End If
    Next
    
    With vsPati
        If .RowData(NewRow) = -1 Then
            .Tag = "���ز���"
            For i = 4 To .Rows - 2
                .RowHidden(i) = False
            Next
            .RowHidden(.Rows - 1) = True
        End If
    End With
End Sub

Private Sub vsPati_GotFocus()
    vsPati.BackColorSel = &HFFCC99
End Sub

Private Sub vsPati_LostFocus()
    vsPati.BackColorSel = &HC0C0C0
End Sub

Private Function CheckCanSelRow(ByVal lngRow As Long) As String
'����:��ָ֤�����Ƿ����ѡ��
    Dim lngCol As Long
    Dim strContent As String

    With vsAdvice
        If .TextMatrix(lngRow, col�������) = "D" And Trim(.Cell(flexcpData, lngRow, col����) & "") <> "" Then
            If strContent <> "[]" Then
                strContent = Chr(34) & strContent & Chr(34)
            Else
                strContent = Chr(34) & .Cell(flexcpData, lngRow, col����) & Chr(34)
            End If
        Else
            strContent = Chr(34) & .Cell(flexcpData, lngRow, col����) & Chr(34)
        End If
        If Val(.RowData(lngRow)) < 0 Then
            Select Case Val(.RowData(lngRow))
            Case con��Ŀ����
                strContent = "������Ŀ��" & .TextMatrix(lngRow, col��Ŀ����) & "���ѳ�����"
            Case con��Ŀ����
                strContent = "������Ŀ��" & .TextMatrix(lngRow, col��Ŀ����) & "������������仯��"
            Case con��Ŀ���
                strContent = "������Ŀ��" & .TextMatrix(lngRow, col��Ŀ����) & "��������仯��"
            Case con�շѳ���
                strContent = "�շ���Ŀ��" & .TextMatrix(lngRow, col�շ�����) & "���ѳ�����"
            Case con�շѷ���
                strContent = "�շ���Ŀ��" & .TextMatrix(lngRow, col�շ�����) & "������������仯��"
            End Select
            CheckCanSelRow = strContent: Exit Function
        End If

        If InStr("δ֪" & mstr�Ա�, .TextMatrix(lngRow, col�Ա�)) = 0 Then
            CheckCanSelRow = strContent & "(�������ڵ�ǰ�����Ա�)": Exit Function
        End If
        
        If .TextMatrix(lngRow, col�Ƿ�����) = "1" Then
            CheckCanSelRow = strContent & "(�������ڵ�ǰ����)": Exit Function
        End If

        If mbln������Ȩ�� And .TextMatrix(lngRow, col�������) = "����ҩ" Then
            CheckCanSelRow = strContent & "(��������ҩƷȨ��)": Exit Function

        End If

        If mbln������Ȩ�� And .TextMatrix(lngRow, col�������) = "����ҩ" Then
            CheckCanSelRow = strContent & "(�޶���ҩƷȨ��)": Exit Function

        End If

        If mbln������Ȩ�� And (.TextMatrix(lngRow, col�������) = "����I��") Then
            CheckCanSelRow = strContent & "(�޾�����ҩƷȨ��)": Exit Function

        End If

        If mbln������Ȩ�� And (.TextMatrix(lngRow, col��ֵ����) = "����" Or .TextMatrix(lngRow, col��ֵ����) = "����") Then
            CheckCanSelRow = strContent & "(�޹�����ҩƷȨ��)": Exit Function

        End If
        
        '��Ѫҽ����飬�����м�������רҵ����ְ���ҽʦ�������´�
        If .TextMatrix(lngRow, col�������) = "K" And gbln��Ѫ�����м����� Then
            If UserInfo.רҵ����ְ�� <> "����ҽʦ" And UserInfo.רҵ����ְ�� <> "����ҽʦ" And UserInfo.רҵ����ְ�� <> "������ҽʦ" Then
                CheckCanSelRow = Trim(.Cell(flexcpText, lngRow, col����) & "") & "(���м�������רҵ����ְ��)": Exit Function
            End If
        End If

    End With

End Function

Private Sub GetRowScope(ByVal lngRow As Long, lngBegin As Long, lngEnd As Long, Optional lngRow��� As Long)
'����:��ȡһ��ҽ������ֹλ�ã�ͬʱ��ȡ��ҽ���к�
'����:
'   lngRow ��ǰ��
'����:
'   lngBegin ��ʼ��
'   lngEnd ��ֹ��
'   lngRow��� ��ҽ����

    Dim i As Long, lng��� As Long

    With vsAdvice
        If .TextMatrix(lngRow, col�������) = "" Then  '����¼��
            lngRow��� = lngRow: lngBegin = lngRow: lngEnd = lngRow
            Exit Sub
        End If
        '��ȡ������
        If Val(.TextMatrix(lngRow, col���ID)) <> 0 Then
            lng��� = Val(.TextMatrix(lngRow, col���ID))
            lngRow��� = .FindRow(lng���, , colID, , True)
            If lngRow��� = -1 Then
                lngRow��� = lngRow
            End If
        Else
            lng��� = Val(.TextMatrix(lngRow, colID)): lngRow��� = lngRow
        End If

        lngBegin = lngRow���: lngEnd = lngRow���

        For i = lngRow��� - 1 To .FixedRows Step -1
            If Not (Val(.TextMatrix(i, colID)) = 0 And i >= .FixedRows) Then '������һ����ҩ�У�����
                If Val(.TextMatrix(i, col���ID)) = lng��� Then
                    lngBegin = i
                Else
                    Exit For
                End If
            End If
        Next

        For i = lngRow��� + 1 To .Rows - 1
            If Not (Val(.TextMatrix(i, colID)) = 0 And i >= .FixedRows) Then '������һ����ҩ�У�����
                If Val(.TextMatrix(i, col���ID)) = lng��� Then
                    lngEnd = i
                Else
                    Exit For
                End If
            End If
        Next
    End With
End Sub

Private Function CheckCanSelGroup(ByVal lngRow As Long, Optional ByVal blnAsk As Boolean = True) As Boolean
'���ܣ��жϱ���ҽ���Ƿ����ѡ��
'������
'   lngRow ��ǰ��
'   blnAsk �Ƿ������ʾ��ѯ�ʣ�ȫѡʱ��������ѯ��)
    Dim i As Long, strResult As String
    Dim lngBegin As Long, lngEnd As Long, lngRow��� As Long
    Dim bln�䷽ As Boolean, blnCanSel As Boolean
    Dim strMsg As String
    Dim blnMedicineAdvice As Boolean

    With vsAdvice
        '��ȡ����ҽ����Ϣ
        Call GetRowScope(lngRow, lngBegin, lngEnd, lngRow���)
        
        If Not mblnҽ������ And mlngǰ��ID <> 0 Then
            If .TextMatrix(lngRow, col��Ч) = "����" Then
                If blnAsk Then
                    MsgBox "ϵͳ�������ҽ��ҽ�����к��������������Ƴ�����", vbInformation, gstrSysName
                End If
                Exit Function
            End If
        End If
        
         '����Ƿ���ҽ����������Ŀδ��ѡ�����Ե���Ӧ�á���δ��ѡ�Ĳ������ơ�
        If lngBegin = lngEnd Then
            If Val(.TextMatrix(lngRow, col����Ӧ��)) = 0 And Val(.TextMatrix(lngRow, col������ĿID)) <> 0 Then
                If blnAsk Then
                    MsgBox "ҽ����" & .TextMatrix(lngRow, col����) & "����Ӧ��������Ŀ���ܵ���Ӧ�ã������Ա����ơ��������ʣ�����ϵ����Ա��", vbInformation, gstrSysName
                End If
                Exit Function
            End If
        Else
            For i = lngBegin To lngEnd
                If InStr(",5,6,7,", .TextMatrix(i, col�������)) > 0 Then
                    blnMedicineAdvice = True
                End If
            Next
            If Not blnMedicineAdvice Then
                For i = lngBegin To lngEnd
                    If Not (.TextMatrix(i, col�������) = "G" Or (.TextMatrix(i, col�������) = "E" And InStr(",2,3,4,6,7,8,", .TextMatrix(i, col��������)) > 0)) Then
                        If Val(.TextMatrix(i, col����Ӧ��)) = 0 And Val(.TextMatrix(i, col������ĿID)) <> 0 Then
                            If blnAsk Then
                                MsgBox "ҽ����" & .TextMatrix(i, col����) & "����Ӧ��������Ŀ���ܵ���Ӧ�ã������Ա����ơ��������ʣ�����ϵ����Ա��", vbInformation, gstrSysName
                            End If
                            Exit Function
                        End If
                    End If
                Next
            End If
        End If

        
        strMsg = ""
        '�����ò��� ָ��ҩ��ʱ���ƿ�� ������£��������ƿ�治���ҽ��
        If gblnStock Then
            For i = lngBegin To lngEnd
                If Val(.Cell(flexcpData, i, col�Ƿ�����)) = 1 Then
                    If Val(.TextMatrix(lngBegin, col�������)) = 7 Then
                        strMsg = strMsg & "," & .TextMatrix(i, col����)
                    Else
                        If blnAsk Then
                            MsgBox "��ҩƷ��治��,ϵͳ�����˲������´��治���ҩƷ�����ܱ�ѡ��", vbInformation, gstrSysName
                        End If
                        Exit Function
                    End If
                End If
            Next
        End If
        If strMsg <> "" Then
            MsgBox "���䷽�д��ڿ�治���ҩƷ(" & Mid(strMsg, 2) & ")��", vbInformation, gstrSysName
        End If
        
        strMsg = CheckCanSelRow(lngRow���)
        If strMsg <> "" Then '��ҽ������ҽ�����
            If blnAsk Then
                MsgBox "��ҽ����:" & vbNewLine & strMsg & vbNewLine & "��Ч,���ܱ�ѡ��", vbInformation, gstrSysName
            End If
            Exit Function
        Else
            If lngBegin <> lngEnd Then
                If .TextMatrix(lngRow���, col�������) = "E" Then
                    If lngRow��� - 2 >= lngBegin Then
                        If .TextMatrix(lngRow��� - 2, col�������) = "7" And .TextMatrix(lngRow��� - 1, col�������) = "E" Then '��ҩ�䷽�ļ������
                            strMsg = CheckCanSelRow(lngRow��� - 1)
                            If strMsg <> "" Then
                                If blnAsk Then
                                    MsgBox "����ҩ�䷽�м巨:" & vbNewLine & strMsg & vbNewLine & "��Ч,���ܱ�ѡ��", vbInformation, gstrSysName
                                End If
                                Exit Function
                            Else
                                bln�䷽ = True
                            End If
                        End If
                    End If
                End If
            End If
        End If
        strMsg = ""
        '��ҽ��ȫ�����
        If lngBegin <> lngEnd Then
            For i = lngBegin To lngEnd
                If .TextMatrix(lngRow���, col�������) = "F" Then blnCanSel = True '����ҽ����ҽ�����þͿ�ѡ
                If Not (i = lngRow��� Or bln�䷽ And i = lngRow��� - 1) Then
                    strResult = CheckCanSelRow(i)
                    If strResult <> "" Then
                        strMsg = IIF(strMsg = "", "", strMsg & "��" & vbNewLine) & strResult
                    Else
                        If bln�䷽ Then  '��ҩ�䷽��һζ��ҩ���þͿ�ѡ���巨�Լ��÷�ǰ���Ѿ��жϣ�����������ֻҪһ����ҽ�����þͿ�ѡ
                            If .TextMatrix(i, col�������) = "7" Then
                                blnCanSel = True
                            End If
                        Else
                            blnCanSel = True
                        End If
                    End If
                End If
            Next
        Else
            blnCanSel = True '����ҽ������ڸ�ҽ��ʱ�Ѿ����
        End If
        
        If Not blnCanSel Then
        '��ҩ�䷽δ��ȡ��ҩƷ��Ϣ
            If bln�䷽ Then
                strMsg = "����ҩ�䷽��������ҩ�Ѿ���ͣ�û�û�п��ù��,���ܱ�ѡ��"
            Else
                If strMsg = "" Then strMsg = "��ҽ���в�������Ч��Ŀ,���ܱ�ѡ��"
            End If
        End If

        If blnCanSel Then
            If strMsg <> "" Then
                If blnAsk Then
                    If MsgBox(IIF(InStr(1, strMsg, "��") > 0, "��ҽ����:" & vbNewLine & strMsg & vbNewLine & "��Ч,��Щ��Ŀ", "��ҽ����:" & vbNewLine & strMsg & vbNewLine & "��Ч,����Ŀ") & "���ᱻѡ��,�Ƿ�ѡ���ҽ����", _
                        vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                        CheckCanSelGroup = True
                    End If
                End If
            Else
                CheckCanSelGroup = True
            End If
        Else
            If blnAsk Then
                MsgBox strMsg, vbInformation, gstrSysName
            End If
        End If
    End With
End Function

Private Sub SelGroup(ByVal lngRow As Long, ByVal intѡ�� As Integer, Optional ByRef lngEnd As Long)
'����:�������ѡ�����ҽ��
'������
'   lngRow ��ǰ��
'   lngEnd ����ҽ�����һ��
'   intѡ�� ѡ���� -1,���ѡ�񣨿�ѡ��ѡ�񣬲���ѡ��ѡ��),0��ѡ��,1��ȫѡ�����
    Dim lngBegin As Long
    Dim i As Long

    With vsAdvice
        
        '��ȡ����ҽ����Ϣ
        Call GetRowScope(lngRow, lngBegin, lngEnd)
        'ѡ���ȡ��ѡ��
        If intѡ�� = -1 Then 'checkCanSelGroup(i,true)=true�����
            For i = lngBegin To lngEnd
                If CheckCanSelRow(i) = "" Then
                    .TextMatrix(i, colѡ��) = intѡ��
                Else
                    .TextMatrix(i, colѡ��) = 0
                End If
            Next
        Else 'checkCanSelGroup(i,false)=true ����û�ȡ��ѡ���ʹ��
            intѡ�� = intѡ�� * -1
            For i = lngBegin To lngEnd
                .TextMatrix(i, colѡ��) = intѡ��
            Next
        End If
    End With
End Sub

Private Sub AlterGroupType(ByVal lngRow As Long)
'���ܣ��ı�ָ��������ҽ�����ҽ����Ч
    Dim lngBegin As Long, lngEnd As Long, i As Long
    
    With vsAdvice
        .Redraw = False
        '��ȡ����ҽ����Ϣ
        Call GetRowScope(lngRow, lngBegin, lngEnd)
        '�ı䱾������ҽ����Ч
        For i = lngBegin To lngEnd
            If .TextMatrix(i, .Col) <> "" Then
                .TextMatrix(i, .Col) = IIF(.TextMatrix(i, .Col) = "����", "����", "����")
            End If
            .Cell(flexcpFontBold, i, .Col) = .TextMatrix(i, .Col) <> .Cell(flexcpData, i, .Col)
        Next
        .Redraw = True
    End With
End Sub

Private Sub InitAdviceTable()
'���ܣ���ʼ���������
    Dim strHead As String, i As Integer
    Dim arrHead As Variant, arrCol As Variant
    
    strHead = ",300,4;��Ч,480,1;��ʼʱ��,1530,1;����,2910,1;����,630,7;��λ,465,1;����,630,7;��λ,465,1;Ƶ��,1140,1;�÷�,1140,1;ҽ������,1650,1;ִ��ʱ��,1095,1;ִ�п���,1110,1;" & _
        "ID;���ID;�������;������ĿID;�շ�ϸĿID;�Ƿ��ʺ�;�������;��ֵ����;�Ա�;����Ӧ��;��������;��Ŀ����;�շ�����"
    arrHead = Split(strHead, ";")
    With vsAdvice
        .Clear
        .FixedRows = 1: .FixedCols = 0
        .Rows = 2: .Cols = .FixedCols + UBound(arrHead) + 1
        For i = 0 To UBound(arrHead)
            .FixedAlignment(.FixedCols + i) = 4
            arrCol = Split(arrHead(i), ",")
            .TextMatrix(0, .FixedCols + i) = arrCol(0)
            If UBound(arrCol) > 0 Then
                .ColWidth(.FixedCols + i) = Val(arrCol(1))
                .ColAlignment(.FixedCols + i) = Val(arrCol(2))
                .ColHidden(.FixedCols + i) = False
            Else
                .ColHidden(.FixedCols + i) = True
            End If
        Next
    End With
End Sub

