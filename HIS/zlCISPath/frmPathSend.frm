VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{FBAFE9A8-8B26-4559-9D12-D70E36A97BE3}#2.1#0"; "zlRichEditor.ocx"
Begin VB.Form frmPathSend 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "����·����Ŀ"
   ClientHeight    =   8580
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11775
   Icon            =   "frmPathSend.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8580
   ScaleWidth      =   11775
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdPati 
      Height          =   240
      Left            =   7560
      Picture         =   "frmPathSend.frx":6852
      Style           =   1  'Graphical
      TabIndex        =   17
      TabStop         =   0   'False
      ToolTipText     =   "ѡ��Ӥ��"
      Top             =   4080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.ListBox lstPati 
      Appearance      =   0  'Flat
      Height          =   1080
      ItemData        =   "frmPathSend.frx":6948
      Left            =   5160
      List            =   "frmPathSend.frx":6955
      Sorted          =   -1  'True
      Style           =   1  'Checkbox
      TabIndex        =   16
      Top             =   4080
      Visible         =   0   'False
      Width           =   2295
   End
   Begin zlRichEditor.Editor edtEditor 
      Height          =   375
      Left            =   7320
      TabIndex        =   1
      Top             =   1320
      Visible         =   0   'False
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
   End
   Begin VB.PictureBox picBottom 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00F0F4E4&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   11775
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   8025
      Width           =   11775
      Begin VB.CommandButton cmdMergeStep 
         Caption         =   "�ϲ�·���׶�ѡ��(&S)"
         Height          =   350
         Left            =   120
         TabIndex        =   15
         Top             =   120
         Width           =   1935
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "ȷ��(&O)"
         Height          =   350
         Left            =   9360
         TabIndex        =   14
         Top             =   120
         Width           =   1100
      End
      Begin MSComCtl2.DTPicker dtpAdviceTime 
         Height          =   300
         Left            =   7320
         TabIndex        =   11
         Top             =   145
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   190513155
         CurrentDate     =   41129.5916666667
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "ȡ��(&C)"
         Height          =   350
         Left            =   10560
         TabIndex        =   8
         Top             =   120
         Width           =   1100
      End
      Begin VB.Label lblAdviceTime 
         Caption         =   "ҽ��ȱʡ��ʼʱ��"
         Height          =   180
         Left            =   5760
         TabIndex        =   10
         Top             =   205
         Width           =   1575
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   1
         X1              =   0
         X2              =   11760
         Y1              =   45
         Y2              =   45
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   0
         X1              =   0
         X2              =   11760
         Y1              =   30
         Y2              =   30
      End
   End
   Begin VB.PictureBox picInfo 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00F0F4E4&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   11775
      TabIndex        =   3
      Top             =   0
      Width           =   11775
      Begin VB.Label lblFont 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Height          =   180
         Left            =   3960
         TabIndex        =   18
         Top             =   0
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.Label lblNote 
         BackStyle       =   0  'Transparent
         Caption         =   "��ǰʱ��׶Σ�"
         Height          =   255
         Left            =   1200
         TabIndex        =   6
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label lblDate 
         BackStyle       =   0  'Transparent
         Caption         =   "��ǰʱ�䣺"
         Height          =   255
         Left            =   1200
         TabIndex        =   5
         Top             =   120
         Width           =   4455
      End
      Begin VB.Label lblPhase 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmPathSend.frx":6971
         Height          =   615
         Left            =   2520
         TabIndex        =   4
         Top             =   360
         Width           =   6855
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   2
         X1              =   0
         X2              =   10000
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   3
         X1              =   0
         X2              =   11760
         Y1              =   945
         Y2              =   945
      End
      Begin VB.Image imgInfo 
         Height          =   720
         Left            =   195
         Picture         =   "frmPathSend.frx":6A0B
         Top             =   45
         Width           =   720
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsItem 
      Height          =   3405
      Index           =   0
      Left            =   30
      TabIndex        =   2
      Top             =   1950
      Width           =   11655
      _cx             =   20558
      _cy             =   6006
      Appearance      =   0
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
      BackColor       =   16777215
      ForeColor       =   -2147483640
      BackColorFixed  =   15597549
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   16777215
      BackColorBkg    =   -2147483643
      BackColorAlternate=   16777215
      GridColor       =   32768
      GridColorFixed  =   32768
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   16777215
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   7
      Cols            =   9
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   250
      RowHeightMax    =   320
      ColWidthMin     =   0
      ColWidthMax     =   8000
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmPathSend.frx":7293
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   1
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
      OwnerDraw       =   0
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
      BackColorFrozen =   14811105
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VSFlex8Ctl.VSFlexGrid vsPhase 
      Height          =   705
      Left            =   30
      TabIndex        =   0
      Top             =   1200
      Width           =   11655
      _cx             =   20558
      _cy             =   1244
      Appearance      =   0
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
      BackColor       =   15597549
      ForeColor       =   -2147483640
      BackColorFixed  =   15597549
      ForeColorFixed  =   15724768
      BackColorSel    =   45056
      ForeColorSel    =   16777215
      BackColorBkg    =   15597549
      BackColorAlternate=   15597549
      GridColor       =   32768
      GridColorFixed  =   32768
      TreeColor       =   32768
      FloodColor      =   192
      SheetBorder     =   15724768
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   2
      GridLines       =   2
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   3
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   450
      ColWidthMin     =   0
      ColWidthMax     =   5000
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmPathSend.frx":7438
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   1
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
      OwnerDraw       =   0
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
      BackColorFrozen =   14811105
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin MSComctlLib.TabStrip tabBranch 
      Height          =   375
      Left            =   0
      TabIndex        =   9
      Top             =   870
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "��·��"
            Key             =   "_0"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid vsItem 
      Height          =   2355
      Index           =   1
      Left            =   30
      TabIndex        =   12
      Top             =   5640
      Width           =   11655
      _cx             =   20558
      _cy             =   4154
      Appearance      =   0
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
      BackColor       =   16777215
      ForeColor       =   -2147483640
      BackColorFixed  =   15597549
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   16777215
      BackColorBkg    =   -2147483643
      BackColorAlternate=   16777215
      GridColor       =   32768
      GridColorFixed  =   32768
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   16777215
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   7
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   250
      RowHeightMax    =   320
      ColWidthMin     =   0
      ColWidthMax     =   8000
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmPathSend.frx":74CD
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   1
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
      OwnerDraw       =   0
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
      BackColorFrozen =   14811105
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.Label lblMerge 
      Caption         =   "�ϲ�·��:"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   5400
      Width           =   11415
   End
End
Attribute VB_Name = "frmPathSend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngFun '0-����·����1-��������(����ѡ��׶�),2-�鿴·���׶ζ���,3-��������ҽ��

Private mPP As TYPE_PATH_Pati
Private mPati As TYPE_Pati

Private mint���� As Integer  '0-ҽ��վ����,1-��ʿվ����
Private mlngʱ����� As Integer 'mlngFun=0ʱ���룬1=��һ�׶���ǰ������,2=��һ�׶���ǰ������,-1=��һ�׶��Ӻ󣨼�����ǰ�׶Σ�,0=����

Private mlng��ĿID As Long       '�������ɵ���ĿID
Private mlngִ��ID As Long       '�������ɵ�·��ִ��ID

Private mlng���˽׶�ID As Long   '��ǰѡ��Ľ׶�(�鿴ʱ)�����˵�ǰ�׶Σ�����ʱ��
Private mlng���� As Long         '��ǰӦ�����ɵ�����(ʵ������)
Private mdatʱ�� As Date         '��ǰӦ�ý��������(����·����Ŀ��ҽ�������ں�ʱ��)
Private mlng��¼��� As Long
Private mlng��ǰ���� As Long         '����ʱӦ�����ɵ�����(������һ�׶���ǰʱ)
Private mlng·��ҽ������ As Long   '·��ҽ�����ɳ�ǰ����
Private mblnIsHaveBranch As Boolean  '�Ƿ���ڷ�֧·��
Private mdatDur As Date          '·������ʱ��
Private mrsMerge As ADODB.Recordset
Private mstrMerge As String      '�Ѿ�ѡ��ĺϲ�·���׶Σ���֧1:�׶�ID1,��֧2:�׶�ID2........
Private mstrMergeStep As String  '�Ѿ�ѡ��ϲ�·���׶�,�������ɣ��ϲ�·����¼ID1:�׶�ID1,�ϲ�·����¼ID2:�׶�ID2........
Private mclsMipModule As zl9ComLib.clsMipModule ' ��Ϣƽ̨����

Private mstrBaby As String 'Ӥ��������,��������,���˼,...
Private mfrmParent As Object
Private mrsPhase As ADODB.Recordset
Private mcol As Collection
Private mEditType As Collection
Private mlngMergeCount As Long   '�ϲ�·����

Private Enum ִ�з�ʽ
    T0����ִ�� = 0
    T1ÿ����� = 1
    T2����һ�� = 2
    T3��Ҫʱ = 3
    T4�����ҽ�һ�� = 4
End Enum

Private Enum TYPE_Func
    Func����·�� = 0
    Func�������� = 1
    Func�鿴·�� = 2
    Func�������� = 3
End Enum

Private mblnOK As Boolean

Public Function ShowMe(frmParent As Object, ByVal lngFun As Long, ByVal int���� As Integer, t_pati As TYPE_Pati, t_pp As TYPE_PATH_Pati, _
    ByVal lng���˽׶�ID As Long, ByVal lng���� As Long, Optional ByVal lng��ĿID As Long, Optional ByVal lngִ��ID As Long, _
    Optional ByVal lngʱ����� As Long, Optional ByRef objMip As Object, Optional ByVal bln��ǰ As Boolean = False, Optional ByVal strSQLPhase As String) As Boolean
'������lng��ĿID,lngִ��ID=��������ҽ��ʱ���贫��
'      lngʱ�����=mlngFun=0ʱ���룬1=��һ�׶���ǰ,2-��һ�׶���ǰ������,-1=��һ�׶��Ӻ󣨼�����ǰ�׶Σ�,0=����
'      bln��ǰ=true :��ǰ����·��,False-����ǰ����
'     strSQLPhase-����·������ʱ���� SQL���,�ֶ��У��׶�ID,����,����
    Set mfrmParent = frmParent
    mlngFun = lngFun
    mint���� = int����
    mlng��ĿID = lng��ĿID
    mlngִ��ID = lngִ��ID
    
    mPati = t_pati
    mPP = t_pp
    mlng���˽׶�ID = lng���˽׶�ID  'ȱʡѡ�е�ǰ�׶�
    mlng���� = lng����
    mlng��ǰ���� = lng����
    mlngʱ����� = lngʱ�����
    If bln��ǰ Then
        '��ǰ����
        mdatDur = DateAdd("d", 1, CDate(Format(mPP.��ǰ����, "YYYY-MM-DD 00:00:00")))
    Else
        mdatDur = zlDatabase.Currentdate
    End If
    If Not (objMip Is Nothing) Then Set mclsMipModule = objMip
    Set mrsPhase = GetPhase(mPP.·��ID, mPP.�汾��, mlng���˽׶�ID, mPP.��ǰ�׶η�֧ID, mlng����, , strSQLPhase)
    If mrsPhase.RecordCount = 0 Then
        MsgBox "��ǰʱ��(��" & lng���� & "��)û�����õ�·���׶Σ���������·����Ŀ��" & vbCrLf & "�����ǲ�����Ժ���������˱�׼סԺ�գ�����û�к�����ʱ��׶Ρ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    mblnOK = False
    Me.Show 1, frmParent
    ShowMe = mblnOK
End Function

Private Function GetPhase(ByVal lng·��ID As Long, ByVal lng�汾�� As Long, ByVal lng��ǰ�׶�ID As Long, ByVal lng��ǰ�׶η�֧ID As Long, _
        ByVal lng���� As Long, Optional ByVal lng�ϲ�·����¼ID As Long, Optional ByVal strPhase As String) As ADODB.Recordset
'���ܣ���ȡ��ǰʱ����õĽ׶�
'������lng�ϲ�·����¼ID =�ϲ�·����¼ID
'     strPhase -����·������ʱ����
    Dim strSql As String, strIF As String, str�׶η��� As String
    Dim rsTmp As ADODB.Recordset, datPathIn As Date, lngʱ����� As Long
    Dim lng�������� As Long, lng��� As Long
    Dim strMainIF As String
    Dim strSubSQL As String
    
    If mlngFun = 2 Then '�鿴�׶ζ������Ŀ
        strSql = "Select a.Id, Nvl(a.��id,0) as ��id, a.���, a.����, a.˵��,a.��ʼ����, a.��������, a.����,NVL(a.��֧ID,0) AS ��֧ID" & vbNewLine & _
                "From �ٴ�·���׶� A" & vbNewLine & _
                "Where a.·��id = [1] And a.�汾�� = [2] And a.id = [4]" & vbNewLine & _
                "Order by ���"
    Else
        datPathIn = GetPatiInPath(mPati, mPP.����·��ID)
        If lng·��ID = mPP.·��ID Then
            mdatʱ�� = DateAdd("d", lng���� - 1, datPathIn)  '����Ӧ�����ɵ�����
        
            strSql = "Select To_number(Trunc(Sysdate)-Trunc([1])) ��¼��� From Dual"
            On Error GoTo errH
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "��ȡ��¼���", mdatʱ��)
            mlng��¼��� = Val("" & rsTmp!��¼���)
        End If
        If mlngFun = 0 Then
            If mint���� = 1 Then
                '��ʿվ���ɿ�ѡ�׶Σ���ҽ��վ��ͬ�׶���Ϣ�ǴӲ���·��ִ����ȡ��
                strSql = "Select b.����,b.����,a.Id, Nvl(a.��id,0) as ��id, a.���, a.����, a.˵��,a.��ʼ����, a.��������, a.����,NVL(a.��֧ID,0) AS ��֧ID From (" & strPhase & ") B, �ٴ�·���׶� A Where a.Id = b.�׶�ID order by b.���� "
                On Error GoTo errH
                Set GetPhase = zlDatabase.OpenSQLRecord(strSql, "��ȡ���ý׶�", strPhase)
                Exit Function
            ElseIf mint���� = 0 Then 'ҽ��վ
                If mlngʱ����� = -1 Then    '�Ӻ�ʱ������ǰ�׶�
                    strIF = " And a.id = [4]"
                Else
                    If mPP.��ǰ�׶�ID <> 0 Then
                        If mPP.ԭ·��ID = lng·��ID Or lng�ϲ�·����¼ID <> 0 Then
                            lng��� = GetPhaseNO(IIf(lng�ϲ�·����¼ID <> 0, lng��ǰ�׶�ID, mPP.��ǰ�׶�ID))
                        Else
                            '�����ת��·��������·�����ܸò�����ǰ�ù������õĽ׶����Ӧ���ڵ����ϴ��ù������׶���š�
                            lng��� = GetLastPhaseNO(mPP.����·��ID, lng·��ID)
                        End If
                    End If
                    
                    If mlngʱ����� = 1 Or mlngʱ����� = 2 Then
                        'ʱ�����=2,��ǰ������,
                        '��ǰʱ��ʾ��һ��Ŀ��ý׶�(ֻ���ǵ�ǰ�׶εĺ����׶Ρ����컹�к����׶β���ʹ�ã����ֲ�Ӧ������ʱ��Ϊ��ǰ)
                        lng�������� = GetMustDay(mPP.����·��ID, lng����, , lng�ϲ�·����¼ID)
                        strIF = " And Decode(a.��֧ID,Null,NVL(d.���,a.���),NVL(d.���,a.���)+NVL(E.���,c.���))>[6] "
                    Else
                        '��������(��������������ѡ�׶�)
                        If mPP.��ǰ�׶�ID <> 0 Then
                            lng�������� = GetMustDay(mPP.����·��ID, lng����, , lng�ϲ�·����¼ID)
                            
                            '֮ǰ��������ǰִ�й��Ľ׶ε�ʱ�䷶Χ�ڵ�ǰ�����ڣ�Ҫ�ų���Щ�׶Ρ�·����תʱ�����
                            strIF = " And Decode(a.��֧ID,Null,NVL(d.���,a.���),NVL(d.���,a.���)+NVL(E.���,c.���))>=[6] "
                        Else
                            lng�������� = lng����
                        End If
                        
                         'ͬһ���ж���׶�ʱ����ǰ�׶μ���֧��������,����ǽ�����һ���ˣ���˵��û����ͬ�����Ľ׶�
                        If lng���� = mPP.��ǰ���� Then
                            strIF = strIF & " And Nvl(a.��id,a.id) <> " & IIf(mPP.�׶θ�ID <> 0, "[8]", "[4]")
                        End If
                    End If
                    
                    '����Ƿ�֧·���������ǰһ�׶ε����
                    If lng��ǰ�׶η�֧ID <> 0 Then
                        strSql = "Select ǰһ�׶�ID From �ٴ�·����֧ Where ID=[1]"
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "��ȡ��֧·��ǰһ�׶�ID", lng��ǰ�׶η�֧ID)
                        If rsTmp.RecordCount > 0 Then lng��� = lng��� + GetPhaseNO(Val(rsTmp!ǰһ�׶�ID & ""))
                    End If
                    
                    str�׶η��� = Get�׶η���(mPP.����·��ID)
                    If str�׶η��� <> "" Then
                        strIF = strIF & " And (a.��id is Null Or a.��id is Not Null And a.���� = [5])"
                    End If
                    '�����ǰ�׶��Ѿ��Ƿ�֧·������ֻ�ܼ����ߵ�ǰ��֧,�����жϵ�ǰ�׶��Ƿ���ڷ�֧
                    If lng��ǰ�׶η�֧ID <> 0 Then
                        strIF = strIF & " And a.��֧ID=[7]"
                    Else
                        strIF = strIF & " And (a.��֧ID is Null or a.��֧ID In(Select ID From �ٴ�·����֧ B Where a.·��id=b.·��id and a.�汾��=b.�汾�� And b.ǰһ�׶�ID=[4]))"
                    End If
                    
                    strMainIF = strIF
                    
                    strIF = strIF & " And (a.��ʼ���� Is Null Or [3] Between a.��ʼ���� And Nvl(a.��������,a.��ʼ����) "
                    '�ϲ�·����ǰֻ����ǰһ���׶�
                    If (mlngʱ����� = 1 Or mlngʱ����� = 2) And lng�ϲ�·����¼ID = 0 Then
                        strIF = strIF & " Or a.��ʼ���� >= [3])"
                    Else
                        strIF = strIF & ")"
                    End If
                End If
            End If
        Else
            strIF = " And a.id = [4]"
        End If
      
        strSql = "Select a.Id, Nvl(a.��id,0) as ��id, a.���, a.����, a.˵��,a.��ʼ����, a.��������, a.����,NVL(a.��֧ID,0) AS ��֧ID" & vbNewLine & _
                "From �ٴ�·���׶� A,�ٴ�·����֧ B,�ٴ�·���׶� C,�ٴ�·���׶� D,�ٴ�·���׶� E " & strSubSQL & vbNewLine & _
                "Where a.��֧id=b.id(+) and b.ǰһ�׶�id=c.id(+) And a.��ID=d.id(+) And c.��id=e.id(+) and a.·��id = [1] And a.�汾�� = [2]" & _
                strIF & vbNewLine & " Order by NVL(d.���,a.���)"
 
    End If
    On Error GoTo errH
    Set GetPhase = zlDatabase.OpenSQLRecord(strSql, "��ȡ���ý׶�", lng·��ID, lng�汾��, lng��������, lng��ǰ�׶�ID, str�׶η���, lng���, lng��ǰ�׶η�֧ID, mPP.�׶θ�ID, mPP.����·��ID)
    
    If (mlngʱ����� = 1 Or mlngʱ����� = 2) And GetPhase.RecordCount = 0 Then
        '�׶���ǰʱ�������ǰ�׶��ж��죬�򰴵�ǰ����ȡ������һ�׶Σ�ֱ��ȡ��Ŵ��ڵ�ǰ�׶ε���һ�׶�
        strSql = "Select * From (Select a.Id, Nvl(a.��id,0) as ��id, a.���, a.����, a.˵��,a.��ʼ����, a.��������, a.����,NVL(a.��֧ID,0) AS ��֧ID" & vbNewLine & _
                "From �ٴ�·���׶� A,�ٴ�·����֧ B,�ٴ�·���׶� C,�ٴ�·���׶� D,�ٴ�·���׶� E" & vbNewLine & _
                "Where a.��֧id=b.id(+) and b.ǰһ�׶�id=c.id(+) And a.��ID=d.id(+)  And c.��id=e.id(+) and a.·��id = [1] And a.�汾�� = [2]" & _
                strMainIF & vbNewLine & " Order by NVL(d.���,a.���)) Where Rownum=1"
        Set GetPhase = zlDatabase.OpenSQLRecord(strSql, "��ȡ���ý׶�", lng·��ID, lng�汾��, lng��������, lng��ǰ�׶�ID, str�׶η���, lng���, lng��ǰ�׶η�֧ID)
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function CheckMergeSend(ByVal lng�ϲ�·����¼ID As Long) As Boolean
'���ܣ��жϺϲ�·���Ƿ��ǵ�һ�����ɣ���������Ϣ����飬��Ϊִ���������δ���ɺϲ�·������Ŀ��
    Dim strSql As String, rsTmp As Recordset
    
    On Error GoTo errH
    strSql = "Select Count(1) as ���� From ���˺ϲ�·������ Where �ϲ�·����¼ID=[1] And ·����¼ID=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng�ϲ�·����¼ID, mPP.����·��ID)
    CheckMergeSend = rsTmp!���� = 0
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub cmdMergeStep_Click()
'���ܣ��ϲ�·���׶�ѡ��
    Dim lngDay As Long, lngEOF As Long
    
    If frmPathMergeStep.ShowMe(mfrmParent, mrsMerge, mlngMergeCount, mstrMerge) = True Then
        If mrsMerge.RecordCount > 0 Then
            mrsMerge.MoveFirst
            vsItem(1).Rows = vsItem(1).FixedRows
            mstrMergeStep = ""
            '�����Ҫ·���ĵ�ǰ������������������������(��Ҫ·���Ӻ����ǰ���ϲ�·��Ҳһ��)
            lngDay = mlng���� - mPP.��ǰ����
            Do While Not mrsMerge.EOF
                lngEOF = mrsMerge.AbsolutePosition
                Call LoadItem(Val(mrsMerge!ID & ""), vsItem(1), Val(mrsMerge!·��ID & ""), Val(mrsMerge!�汾�� & ""), Val(mrsMerge!��ǰ���� & "") + lngDay, Val(mrsMerge!�ϲ�·����¼ID & ""))
                mrsMerge.AbsolutePosition = lngEOF
                mrsMerge.MoveNext
            Loop
        End If
    End If
End Sub

Private Sub cmdPati_Click()
    Dim i As Long
    Dim arrtmp As Variant
    Dim lngW As Long
    Dim lngH As Long
    Dim strTmp As String
    Dim strSelect As String
            
    lstPati.Visible = True
    lblFont.FontSize = lblPhase.FontSize
    With lstPati
        .Clear
        strTmp = "���˱���|" & mstrBaby
        arrtmp = Split(strTmp, "|")
        For i = LBound(arrtmp) To UBound(arrtmp)
            lblFont.Caption = arrtmp(i)
            If lngW < lblFont.Width Then lngW = lblFont.Width
           .AddItem arrtmp(i)
        Next
        lngH = (i - 1) * 210 + 240
        If lngH > 1080 Then lngH = 1080
        lngW = lngW + 700
        If lngW > 2500 Then lngW = 2500
    End With

    With vsItem(Val(cmdPati.Tag))
        strSelect = .TextMatrix(.Row, .Col)
        For i = 0 To lstPati.ListCount - 1
            If InStr("|" & strSelect & "|", "|" & lstPati.List(i) & "|") > 0 Then
                lstPati.Selected(i) = True
            End If
        Next
        If lngW < .ColWidth(mcol("Ӥ��")) Then lngW = .ColWidth(mcol("Ӥ��"))
        lstPati.Move .Left + .ColPos(.Col), .Top + .RowPos(.Row) + .RowHeight(.Row) + 30, lngW, lngH
    End With
    Call lstPati.SetFocus
End Sub

Private Sub Form_Load()
    
    If mlngFun <> 2 Then vsItem(0).Editable = flexEDKbdMouse: vsItem(1).Editable = flexEDKbdMouse
    
    Call LoadBranch
    Call LoadPhase
    
    mlng·��ҽ������ = Val(zlDatabase.GetPara("·��ҽ�����ɳ�ǰ����", glngSys, p�ٴ�·��Ӧ��, "1"))
    'ҽ��ȱʡʱ��Ĭ��ȡ��ǰʱ��
    dtpAdviceTime.Value = mdatDur
    
    If vsPhase.Cols = 1 And tabBranch.Tabs.count = 1 Then
        vsPhase.Visible = False
        lblPhase.Caption = vsPhase.TextMatrix(0, 0) & vbCrLf & vsPhase.Cell(flexcpData, 0, 0)
                
        vsItem(0).Top = vsPhase.Top
        vsItem(0).Height = IIf(lblMerge.Visible, lblMerge.Top - 50, picBottom.Top) - vsItem(0).Top
    Else
        lblNote.Visible = False
        lblPhase.Left = lblNote.Left
    
        If Grid.HScrollVisible(vsPhase) Then
            '���������
            vsPhase.Height = 1000
            vsItem(0).Height = vsItem(0).Height - (vsPhase.Top + vsPhase.Height - vsItem(0).Top + 120)
            vsItem(0).Top = vsPhase.Top + vsPhase.Height + 60
        Else
            If vsPhase.Rows = 1 Then
                vsPhase.RowHeightMax = vsPhase.Height
                vsPhase.RowHeight(0) = vsPhase.Height
            End If
        End If
    End If
        
    If mlngFun = 2 Then
        Me.Caption = "�鿴�׶ζ������Ŀ"
        lblDate.Visible = False
        
        cmdOK.Visible = False
        cmdCancel.Caption = "�˳�(&X)"
        mstrBaby = ""
    Else
        lblDate.Caption = "����·����Ŀ���ڣ�" & Format(mdatʱ��, "yyyy-MM-dd") & ",��" & mlng���� & "��"
        
        If mlng��¼��� > 0 Then
            lblDate.Caption = lblDate.Caption & "(" & mlng��¼��� & "��ǰ)"
            lblDate.ForeColor = vbRed
        End If
        mstrBaby = GetBabyRegList
    End If
    If mlngFun <> 0 Then cmdMergeStep.Visible = False
    
    Call InitItem
    
    If mlngFun = 2 Then '�鿴ʱֻ��ʾ���� , ��Ŀ����
        Me.Width = vsItem(0).Width + 360
        cmdCancel.Left = vsItem(0).Left + vsItem(0).Width - 1200
        cmdOK.Left = cmdCancel.Left - cmdOK.Width - 150
    End If
    
    Set mEditType = New Collection
    Call LoadMerge
    Call LoadItem(Val(vsPhase.ColData(vsPhase.Col)), vsItem(0), mPP.·��ID, mPP.�汾��, mlng����)
    
   
    If vsItem(0).Rows = 1 And vsItem(1).Rows = 1 Then
        vsItem(0).Rows = 2
        vsItem(1).Rows = 2
        vsItem(0).TextMatrix(1, mcol("��Ŀ����")) = "û������ִ��һ�λ��ѡ�Ե�·����Ŀ"
        cmdOK.Enabled = False
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mrsPhase = Nothing
    Set mcol = Nothing
    Set mEditType = Nothing
    mstrMerge = ""
    Set mclsMipModule = Nothing
End Sub

Private Sub lstPati_ItemCheck(Item As Integer)
    Dim i As Long
    Dim strList As String
    strList = ""
    For i = 0 To lstPati.ListCount - 1
        If lstPati.Selected(i) Then
            strList = strList & "|" & lstPati.List(i)
        End If
    Next
    If strList <> "" Then
        strList = Mid(strList, 2)
    Else
        strList = lstPati.List(0)  'ȱʡѡ�в��˱���,������Ϊ��
    End If
    
    vsItem(Val(cmdPati.Tag)).TextMatrix(vsItem(Val(cmdPati.Tag)).Row, vsItem(Val(cmdPati.Tag)).Col) = strList
End Sub

Private Sub lstPati_KeyPress(KeyAscii As Integer)
    If lstPati.Visible Then
        If KeyAscii = vbKeyEscape Or KeyAscii = vbKeyReturn Then
            lstPati.Visible = False
        End If
    End If
End Sub

Private Sub lstPati_LostFocus()
    lstPati.Visible = False
End Sub

Private Sub tabBranch_Click()
    Call LoadPhase
    If vsPhase.Cols = 1 And tabBranch.Tabs.count = 1 Then
        vsPhase.Visible = False
        lblPhase.Caption = vsPhase.TextMatrix(0, 0) & vbCrLf & vsPhase.Cell(flexcpData, 0, 0)
        vsItem(0).Top = vsPhase.Top
        vsItem(0).Height = IIf(lblMerge.Visible, lblMerge.Top - 50, picBottom.Top) - vsItem(0).Top
    Else
        vsPhase.Visible = True
        vsItem(0).Top = vsPhase.Top + vsPhase.Height + 45
        vsItem(0).Height = IIf(lblMerge.Visible, lblMerge.Top - 50, picBottom.Top) - vsItem(0).Top
    End If
End Sub

Private Sub vsItem_AfterScroll(Index As Integer, ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    With vsItem(Index)
        If cmdPati.Visible Then
            cmdPati.Move .Left + .ColPos(.Col) + .ColWidth(.Col) - 255, .Top + .RowPos(.Row) + 15, 255, 240
            lstPati.Visible = False
        End If
    End With
End Sub

Private Sub vsItem_AfterUserResize(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    With vsItem(Index)
        If .Col = mcol("Ӥ��") And cmdPati.Visible Then
            cmdPati.Move .Left + .ColPos(.Col) + .ColWidth(.Col) - 255, .Top + .RowPos(.Row) + 15, 255, 240
            lstPati.Visible = False
        End If
    End With
End Sub

Private Sub vsItem_Click(Index As Integer)
    With vsItem(Index)
        If lstPati.Visible Then lstPati.Visible = False
    End With
End Sub

Private Sub vsItem_DblClick(Index As Integer)
    Dim lng��ĿID As Long
    
    If vsItem(Index).Col = mcol("��Ŀ����") Then
        lng��ĿID = Val(vsItem(Index).TextMatrix(vsItem(Index).Row, mcol("ID")))
        If lng��ĿID <> 0 Then
            Call frmPathItemEdit.ShowView(mfrmParent, lng��ĿID)
        End If
    End If
    
End Sub

Private Sub vsItem_GotFocus(Index As Integer)
    vsItem(Index).ForeColorSel = vbWhite
    vsItem(Index).BackColorSel = &H8000000D
End Sub

Private Sub vsItem_LostFocus(Index As Integer)
    vsItem(Index).ForeColorSel = vbBlack
    vsItem(Index).BackColorSel = vbWhite
End Sub

Private Sub vsItem_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call ResultEnterNextCell(vsItem(Index))
    End If
End Sub

Private Sub ResultEnterNextCell(vsthis As VSFlexGrid)
    With vsthis
        If .Col <= .Cols - 1 Then
            If .Row < .Rows - 1 Then
                .Row = .Row + 1
                .ShowCell .Row, .Col
            Else
                Call zlCommFun.PressKey(vbKeyTab)
            End If
        End If
    End With
End Sub

Private Sub vsItem_AfterEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    Dim i As Long
    
    With vsItem(Index)
        If Col = mcol("��ѡ") Then
            If .Cell(flexcpChecked, Row, mcol("��ѡ")) = 1 Then
                If .Cell(flexcpChecked, Row, mcol("ѡ��")) <> 1 Then .Cell(flexcpChecked, Row, mcol("ѡ��")) = 1
            End If
        ElseIf Col = mcol("ѡ��") Then
            If .Cell(flexcpChecked, Row, mcol("ѡ��")) = 2 Then
                'δѡ��ʱȡ����ѡ
                If .Cell(flexcpChecked, Row, mcol("��ѡ")) = 1 Then .Cell(flexcpChecked, Row, mcol("��ѡ")) = 2
                
                'δѡ��ʱ��������ԭ��ѡ��
                If mlngFun = 0 Then
                    If .RowData(Row) = ִ�з�ʽ.T1ÿ����� Then
                        Call vsItem_CellButtonClick(Index, Row, mcol("����ԭ��"))
                    End If
                End If
                
            ElseIf .Cell(flexcpChecked, Row, mcol("ѡ��")) = 1 Then
                If .RowData(Row) = ִ�з�ʽ.T1ÿ����� Then
                'ѡ��ʱ���������ԭ��
                    .TextMatrix(Row, mcol("����ԭ��")) = ""
                    .Cell(flexcpData, Row, mcol("����ԭ��")) = ""
                End If
            End If
        ElseIf Col = mcol("ȫѡ") Then
            If .Cell(flexcpChecked, Row, mcol("ȫѡ")) = 2 Then
                For i = Row To .Rows - 1
                    'ÿ�����ɵ�ȡ����Ҫ��дԭ�����Բ�ȡ����ѡ��Ҫȡ��ֻ��һ��һ��ȡ��
                    If .TextMatrix(i, mcol("����")) <> .TextMatrix(Row, mcol("����")) Then Exit For
                    If Not (.RowData(i) = ִ�з�ʽ.T0����ִ�� Or .RowData(i) = ִ�з�ʽ.T1ÿ�����) Then
                        If .Cell(flexcpChecked, i, mcol("ѡ��")) = 1 Then .Cell(flexcpChecked, i, mcol("ѡ��")) = 2
                    End If
                Next
                For i = Row - 1 To .FixedRows Step -1
                    If .TextMatrix(i, mcol("����")) <> .TextMatrix(Row, mcol("����")) Then Exit For
                    If Not (.RowData(i) = ִ�з�ʽ.T0����ִ�� Or .RowData(i) = ִ�з�ʽ.T1ÿ�����) Then
                        If .Cell(flexcpChecked, i, mcol("ѡ��")) = 1 Then .Cell(flexcpChecked, i, mcol("ѡ��")) = 2
                    End If
                Next
                
            ElseIf .Cell(flexcpChecked, Row, mcol("ȫѡ")) = 1 Then
                For i = Row To .Rows - 1
                    If .TextMatrix(i, mcol("����")) <> .TextMatrix(Row, mcol("����")) Then Exit For
                    If .Cell(flexcpChecked, i, mcol("ѡ��")) = 2 Then
                        .Cell(flexcpChecked, i, mcol("ѡ��")) = 1
                        If .RowData(i) = ִ�з�ʽ.T1ÿ����� Then
                        'ѡ��ʱ���������ԭ��
                            .TextMatrix(i, mcol("����ԭ��")) = ""
                            .Cell(flexcpData, i, mcol("����ԭ��")) = ""
                        End If
                    End If
                Next
                For i = Row - 1 To .FixedRows Step -1
                    If .TextMatrix(i, mcol("����")) <> .TextMatrix(Row, mcol("����")) Then Exit For
                    If .Cell(flexcpChecked, i, mcol("ѡ��")) = 2 Then
                        .Cell(flexcpChecked, i, mcol("ѡ��")) = 1
                        If .RowData(i) = ִ�з�ʽ.T1ÿ����� Then
                        'ѡ��ʱ���������ԭ��
                            .TextMatrix(i, mcol("����ԭ��")) = ""
                            .Cell(flexcpData, i, mcol("����ԭ��")) = ""
                        End If
                    End If
                Next
            End If
        End If
    End With
End Sub


Private Sub vsItem_AfterRowColChange(Index As Integer, ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    With vsItem(Index)
        If NewRow >= .FixedRows And Me.Visible Then
            If mlngFun = 0 Then
                If NewCol = mcol("����ԭ��") Then
                    'δѡ��ʱ�������û�ѡ�����ԭ��
                    If .RowData(NewRow) = ִ�з�ʽ.T1ÿ����� And .Cell(flexcpChecked, NewRow, mcol("ѡ��")) = 2 Then
                        .ColComboList(mcol("����ԭ��")) = "..."
                    Else
                        .ColComboList(mcol("����ԭ��")) = ""
                    End If
                End If
            End If
            cmdPati.Visible = False: lstPati.Visible = False
            If NewCol = mcol("Ӥ��") And mlngFun <> 2 Then
                If mstrBaby <> "" Then
                    cmdPati.Visible = True
                    cmdPati.Enabled = True
                    cmdPati.Move .Left + .ColPos(NewCol) + .ColWidth(NewCol) - 255, .Top + .RowPos(NewRow) + 15, 255, 240
                    cmdPati.Tag = Index
                    If .RowData(NewRow) = ִ�з�ʽ.T0����ִ�� Then cmdPati.Enabled = False
                End If
            End If
        End If
    End With
End Sub

Private Sub vsItem_StartEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsItem(Index)
        If Col = mcol("ѡ��") Then
            '����·��ʱ��ÿ�����ɵģ����Բ�ѡ����Ҫ�������ԭ��
            If .RowData(Row) = ִ�з�ʽ.T0����ִ�� Or mlngFun <> 0 And .RowData(Row) = ִ�з�ʽ.T1ÿ����� Then
                Cancel = True
            End If
        ElseIf Col = mcol("��ѡ") Then
            If Val(.Cell(flexcpChecked, Row, Col)) = 0 Then Cancel = True
        ElseIf Col = mcol("ȫѡ") Then
            If Val(.Cell(flexcpChecked, Row, Col)) = 0 Then Cancel = True
        ElseIf Col = mcol("Ӥ��") Then
            Cancel = True
        ElseIf Col = mcol("����ԭ��") Then
            If .ColComboList(mcol("����ԭ��")) = "" Then Cancel = True
        Else
            Cancel = True
        End If
    End With
End Sub

'�ݲ�֧�ֱ���ԭ�������
'Private Sub vsItem_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
'    If Col = mcol("����ԭ��") Then
'        Dim rsTmp As ADODB.Recordset, strSQL As String, strInput As String
'        Dim vPoint As POINTAPI, blnCancel As Boolean
'
'        vPoint = GetCoordPos(vsItem(0).EditWindow, vsItem(0).CellTop, vsItem(0).CellLeft)
'        strInput = gstrLike & vsItem(0).EditText & "%"
'        strSQL = "Select b.���� As ����, a.����, a.����, a.����" & vbNewLine & _
'                "From ���쳣��ԭ�� A, ���쳣��ԭ�� B" & vbNewLine & _
'                "Where a.ĩ�� = 1 And a.�ϼ� = b.���� and a.����=1 And (a.���� like [1] or ���� like [1] or ���� like [1]" & vbNewLine & _
'                "order by b.����"
'        Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "���쳣��ԭ��", True, False, "��ѡ��", False, False, False, vPoint.X, vPoint.Y, vsItem(0).EditWindow, blnCancel, False, True, strInput)
'        If rsTmp Is Nothing Then
'            If Not blnCancel Then
'                Cancel = True
'                MsgBox "ϵͳû�г�ʼ���쳣��ԭ������ϵͳ����Ա��ϵ��", vbInformation, gstrSysName
'                Exit Sub
'            End If
'        Else
'            vsItem(0).TextMatrix(Row, Col) = rsTmp!����
'            vsItem(0).Cell(flexcpData, Row, Col) = CStr(rsTmp!����)
'        End If
'    End If
'End Sub

Private Sub vsItem_CellButtonClick(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    With vsItem(Index)
        If Col = mcol("����ԭ��") Then
            Dim strSql As String, blnCancel As Boolean
            Dim rsTmp As ADODB.Recordset
                    
            strSql = "Select b.���� as ����,a.���� as ID,a.����,a.����,a.���� From ���쳣��ԭ�� a,���쳣��ԭ�� b" & _
                    " Where a.����=1 And a.ĩ��=1 And a.�ϼ�=b.���� And b.ĩ��=0 " & _
                    " Order by ����,a.����"
            
            Set rsTmp = zlDatabase.ShowSelect(Me, strSql, 0, "���쳣��ԭ��", True, , , True, True, True, _
                     Me.Left + .Left + .ColPos(Col), Me.Top + .Top + .RowPos(Row) + .RowHeight(Row) * 2, .RowHeight(Row), blnCancel, False, True)
            If rsTmp Is Nothing Then
                If Not blnCancel Then
                    MsgBox "ϵͳû�г�ʼ���쳣��ԭ������ϵͳ����Ա��ϵ��", vbInformation, gstrSysName
                    Exit Sub
                End If
            Else
                .TextMatrix(Row, Col) = rsTmp!����
                .Cell(flexcpData, Row, Col) = CStr(rsTmp!����)
            End If
        End If
    End With
End Sub

Private Sub vsPhase_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim lng�׶�ID As Long
    Dim str���� As String

    If OldCol <> NewCol And Me.Visible = True And NewCol >= 0 And NewRow >= 0 And vsPhase.Redraw = flexRDDirect Then
        lng�׶�ID = vsPhase.ColData(NewCol)
        
        If mint���� = 1 And mlngFun = 0 Then
            str���� = vsPhase.TextMatrix(1, NewCol)  '��ʽ�����ڣ���n�죩
            str���� = Val(Mid(str����, InStr(str����, "(��") + 2))
            mrsPhase.Filter = "ID=" & lng�׶�ID & " and ����= " & str����
            mdatʱ�� = CDate(mrsPhase!���� & "")
            mlng���� = Val(mrsPhase!���� & "")
            lblDate.Caption = "����·����Ŀ���ڣ�" & Format(mdatʱ��, "yyyy-MM-dd") & ",��" & mlng���� & "��"
            mlng��ǰ���� = mlng����
        Else
            mrsPhase.Filter = "ID=" & lng�׶�ID
            mlng���� = (mrsPhase!��ʼ���� & "")
        End If
        Call LoadItem(lng�׶�ID, vsItem(0), mPP.·��ID, mPP.�汾��, mlng����)
                
    End If
End Sub

Private Sub LoadBranch()
'���ܣ����ط�֧·��
    Dim i As Long, j As Long, strSql As String
    Dim rsTmp As ADODB.Recordset
    Dim strTmp As String
    
    If Not (mint���� = 1 And mlngFun = 0) Then
        If mrsPhase.RecordCount > 0 Then
            Do While Not mrsPhase.EOF
                If Val(mrsPhase!��֧ID & "") <> 0 Then
                    strTmp = strTmp & "," & mrsPhase!��֧ID
                End If
                mrsPhase.MoveNext
            Loop
            strTmp = strTmp & ","
            mrsPhase.MoveFirst
        End If
    
        strSql = "Select ID,���� From �ٴ�·����֧ Where ǰһ�׶�ID=[3] And ·��ID=[1] And �汾��=[2]"
        On Error GoTo errH
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "��ȡ��֧��Ϣ", mPP.·��ID, mPP.�汾��, mPP.��ǰ�׶�ID)
        If rsTmp.RecordCount > 0 Then
            Do While Not rsTmp.EOF
                '���ܷ�֧·���Ľ׶���û���ʺϵ�ǰ�����Ľ׶�
                If InStr(strTmp, "," & rsTmp!ID & ",") > 0 Then
                    tabBranch.Tabs.Add , "_" & rsTmp!ID, "��֧:" & rsTmp!����
                End If
                rsTmp.MoveNext
            Loop
            mblnIsHaveBranch = True
        End If
    End If
    If tabBranch.Tabs.count = 1 Then
        tabBranch.Visible = False
        mblnIsHaveBranch = False
        vsPhase.Top = tabBranch.Top
        vsItem(0).Top = vsPhase.Top + vsPhase.Height + 45
        vsItem(0).Height = vsItem(0).Height + tabBranch.Height - 15
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub LoadPhase()
'���ܣ����ؿ�ѡ��Ľ׶�,������˵ĵ�ǰʱ��׶���Ȼ���ã���ѡ�У�����ȱʡΪ��һ��
    Dim i As Long, j As Long, str�׶η��� As String
    Dim rsNode As ADODB.Recordset

    With vsPhase
        .Clear
        .Redraw = flexRDNone
        .Col = -1
        If mint���� = 1 And mlngFun = 0 Then
            mrsPhase.Filter = ""  '��ʿ��������·��
            .Cols = mrsPhase.RecordCount
        Else
            mrsPhase.Filter = IIf(mblnIsHaveBranch, "��֧ID=" & Mid(tabBranch.SelectedItem.Key, 2), "")
            .Cols = mrsPhase.RecordCount
            str�׶η��� = Get�׶η���(0, mPP.��ǰ�׶�ID)
            If mlngFun = 0 And mlngʱ����� <> -1 Then '�������ɡ���������ʱ����һ�׶��Ӻ󣨼�����ǰ�׶Σ���ֻ�е�ǰ�׶εļ�¼
                mrsPhase.Filter = "��ID<>0 " & IIf(mblnIsHaveBranch, " And ��֧ID=" & Mid(tabBranch.SelectedItem.Key, 2), "")
                If mrsPhase.RecordCount > 0 Then    '�б��÷�֧
                    Set rsNode = mrsPhase.Clone
                    .Rows = 2
                    .MergeRow(0) = True
                Else
                    .Rows = 1
                End If
                mrsPhase.Filter = "��ID=0" & IIf(mblnIsHaveBranch, " And ��֧ID=" & Mid(tabBranch.SelectedItem.Key, 2), "")
            End If
        End If
        
        For i = 0 To .Cols - 1
            .ColWidth(i) = 2000
            .ColAlignment(i) = flexAlignCenterCenter
            .TextMatrix(0, i) = mrsPhase!����
            .Cell(flexcpData, 0, i) = CStr(IIf(IsNull(mrsPhase!����), "", "���ࣺ" & mrsPhase!���� & " ") & mrsPhase!˵��)
            .ColData(i) = Val(mrsPhase!ID)
            
            
            If mint���� = 1 And mlngFun = 0 Then
                If .ColData(i) & "_" & mrsPhase!���� = mlng���˽׶�ID & "_" & mlng���� Then .Col = i
                .MergeCol(i) = True
                .TextMatrix(1, i) = mrsPhase!���� & " " & "(��" & mrsPhase!���� & "��)"
            Else
                If .ColData(i) = mlng���˽׶�ID Then .Col = i
                If Not rsNode Is Nothing Then
                    rsNode.Filter = "��ID=" & mrsPhase!ID
                    If rsNode.RecordCount = 0 Then
                         .MergeCol(i) = True
                         .TextMatrix(1, i) = mrsPhase!����
                    Else
                         .TextMatrix(1, i) = "ȱʡ"
                         .ColWidth(i) = 1000
                        For j = 1 To rsNode.RecordCount
                            i = i + 1
                             .ColWidth(i) = 1000
                             .ColAlignment(i) = flexAlignCenterCenter
                            .TextMatrix(0, i) = mrsPhase!���� '��һ��������ͬ�������ںϲ�
                            .TextMatrix(1, i) = IIf(IsNull(rsNode!˵��), "��֧" & j, "" & rsNode!˵��)
                            .Cell(flexcpData, 1, i) = CStr(IIf(IsNull(rsNode!����), "", "���ࣺ" & rsNode!���� & " ") & rsNode!˵��)
                            
                            .ColData(i) = Val(rsNode!ID)
                            If .ColData(i) = mlng���˽׶�ID Then
                                .Col = i
                            ElseIf .Col = 0 And str�׶η��� <> "" Then
                                If str�׶η��� = "" & rsNode!���� Then .Col = i
                            End If
                            rsNode.MoveNext
                        Next
                    End If
                End If
            End If
            mrsPhase.MoveNext
        Next
        If .Col < 0 Then .Col = 0
        mrsPhase.Filter = "ID=" & Val(.ColData(.Col))
        .Redraw = True
        vsPhase_AfterRowColChange -1, -1, .Row, .Col
    End With
End Sub

Private Sub LoadItem(lng�׶�ID As Long, objVsg As VSFlexGrid, ByVal lng·��ID As Long, ByVal lng�汾�� As Long, ByVal lng���� As Long, Optional ByVal lng�ϲ�·����¼ID As Long)
'���ܣ����ص�ǰ�׶ε�·����Ŀ
'������objVsg����Ҫ���صı����Ҫ·����ϲ�·����,����Ǽ��غϲ�·�������ں�����ӣ�������б�
    Dim i As Long, j As Long, blnFocus As Boolean, bln���� As Boolean
    Dim rsTmp As ADODB.Recordset, strSql As String, strIDs As String, strTmp As String
    Dim str��ѡ���� As String, lngOld���ID As Long, bln���鳤�� As Boolean
    Dim rsAdvice As ADODB.Recordset, rsFile As ADODB.Recordset
    Dim lngRow As Long
    Dim strFilter As String, blnEnd As Boolean
    Dim str������ĿIDs As String
    Dim lng��Ҫ·���׶�ID As Long
    Dim strNewTmp As String
     
    If mlngFun = 1 Then '�������ɣ�����ִ�еĲ���ʾ������ִ�й��Ĳ����ظ�����,ִֻ��һ�εĵ�ǰ�׶���ִ������ʾ
        strSql = " And a.ִ�з�ʽ<>0 And Not Exists(Select 1 From ����·��ִ�� c " & _
                "Where c.·����¼id = [4] And c.�׶�id = [7] And c.��Ŀid = a.id And (c.���� = [5] and a.ִ�з�ʽ<>4 or a.ִ�з�ʽ=4))"
        lng��Ҫ·���׶�ID = mPP.��ǰ�׶�ID
    ElseIf mlngFun = 3 Then '��������
        strSql = " And a.ID = [6]"
    Else
        strSql = " And (a.ִ�з�ʽ<>4 or a.ִ�з�ʽ=4 And Not Exists(Select 1 From ����·��ִ�� c " & _
                "Where c.·����¼id = [4] And c.�׶�id = [7] And c.��Ŀid = a.id))"
        If objVsg.Index = 0 Then
            lng��Ҫ·���׶�ID = lng�׶�ID
        Else
            lng��Ҫ·���׶�ID = Val(vsPhase.ColData(vsPhase.Col))
        End If
    End If
    '���ӡ��ٴ�·�����ࡱ��ֻ��Ϊ�˰���������'����ʱ�ټ�飬�Ƿ�Ϊ����׶ε����һ�죬����ִ��һ�ε���Ŀ�Ƿ�ѡ��
    strSql = "Select a.����, a.ID, a.��Ŀ����, a.ִ�з�ʽ, a.ͼ��id, a.����Ҫ��" & vbNewLine & _
        "From �ٴ�·����Ŀ A, �ٴ�·������ B" & vbNewLine & _
        "Where a.���� = b.���� And a.·��id = b.·��id And a.�汾�� = b.�汾�� And a.·��id = [1] And a.�汾�� = [2] And a.�׶�id = [3] And NVL(a.��֧ID,0)=nvl(b.��֧id,0)" & vbNewLine & _
        Decode(mint����, 0, " And NVL(a.������,1) = 1 ", 1, " And a.������ = 2 ") & vbNewLine & _
        strSql & vbNewLine & _
        IIf(mlngFun = 3, "", "Order By b.���, a.��Ŀ���")
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng·��ID, lng�汾��, lng�׶�ID, mPP.����·��ID, mlng����, mlng��ĿID, lng��Ҫ·���׶�ID)
    
    With objVsg
        .Redraw = flexRDNone
        If objVsg.Index = 0 Then
            .Rows = .FixedRows
            .Rows = .FixedRows + rsTmp.RecordCount
            lngRow = 1
        Else
            lngRow = .Rows
            If lngRow = 2 Then
                If .TextMatrix(1, mcol("ID")) = "" Then lngRow = 1
            End If
            .Rows = lngRow + rsTmp.RecordCount
        End If
        '���ڹ̶��кϲ�����Ӱ�������У�����������һ���жϰ�����ϲ�ȫѡ��
        .MergeCells = flexMergeRestrictAll
        .MergeCol(mcol("����")) = True
        .MergeCol(mcol("����ֵ")) = True
        .MergeCol(mcol("ȫѡ")) = True
        '�ж��Ƿ������һ�죬����Ǻϲ�·�����ż���
        If objVsg.Index = 1 Then
            strFilter = mrsMerge.Filter
            If lng���� = 0 Then lng���� = 1
            mrsMerge.Filter = "ID=" & lng�׶�ID
            If mrsMerge.RecordCount > 0 Then
                With mrsMerge
                    If Not IsNull(!��ʼ����) Then
                        If IsNull(!��������) Then
                            blnEnd = (Val(!��ʼ����) = lng����)
                        Else
                            blnEnd = (Val(!��������) = lng����)
                        End If
                    End If
                End With
                mrsMerge.Filter = IIf(strFilter = "0", 0, strFilter)
                '��¼�¼��صĺϲ�·���׶�
                mstrMergeStep = mstrMergeStep & "," & lng�ϲ�·����¼ID & ":" & lng�׶�ID
            End If
        End If
        For i = lngRow To rsTmp.RecordCount + lngRow - 1
            .TextMatrix(i, mcol("ID")) = rsTmp!ID
            strIDs = strIDs & "," & rsTmp!ID
            .TextMatrix(i, mcol("����")) = rsTmp!����
            .TextMatrix(i, mcol("����ֵ")) = rsTmp!����
            .TextMatrix(i, mcol("��Ŀ����")) = rsTmp!��Ŀ����
            If mlngFun <> 2 Then .TextMatrix(i, mcol("����Ҫ��")) = Val("" & rsTmp!����Ҫ��)
            .TextMatrix(i, mcol("ִ�з�ʽ")) = Decode(rsTmp!ִ�з�ʽ, 0, "��", 1, "ÿ��", 2, "����һ��", 3, "��Ҫʱ", 4, "����һ��")
            .RowData(i) = Val(rsTmp!ִ�з�ʽ)
            If objVsg.Index = 1 And blnEnd And mlngFun <> 2 Then
                .TextMatrix(i, mcol("�Ƿ����һ��")) = "1"
                .TextMatrix(i, mcol("�׶�ID")) = lng�׶�ID
                .TextMatrix(i, mcol("�ϲ�·����¼ID")) = lng�ϲ�·����¼ID
            End If
            
            If mlngFun <> 2 Then
                Select Case rsTmp!ִ�з�ʽ
                    Case ִ�з�ʽ.T0����ִ��
                        .TextMatrix(i, mcol("ѡ��")) = " "
                        .Cell(flexcpBackColor, i, mcol("ѡ��")) = &H8000000F
                    Case ִ�з�ʽ.T1ÿ�����
                        .Cell(flexcpChecked, i, mcol("ѡ��")) = 1
                        .Cell(flexcpPictureAlignment, i, mcol("ѡ��")) = flexPicAlignCenterCenter
                        If mlngFun <> 0 Then .Cell(flexcpBackColor, i, mcol("ѡ��")) = &H8000000F
                        '����ʱ��������Ϊ��ɫ��������Ϊ���Բ�ѡ�����ɣ���¼�����ԭ��
                   Case Else
                        If mlngFun = 3 Then '��ѡʱ��ѡ����δ��ʾ���Զ�����
                            .Cell(flexcpChecked, i, mcol("ѡ��")) = 1
                        Else
                            .Cell(flexcpChecked, i, mcol("ѡ��")) = IIf(rsTmp.RecordCount = 1, 1, 2)
                            .Cell(flexcpPictureAlignment, i, mcol("ѡ��")) = flexPicAlignCenterCenter
                        End If
                End Select
            End If
            
                        
            If Not IsNull(rsTmp!ͼ��ID) Then
                Call .Select(i, mcol("��Ŀ����"))
                .CellPictureAlignment = flexPicAlignRightCenter 'flexPicAlignLeftCenter
                .CellPicture = GetPathIcon(rsTmp!ͼ��ID)
            End If
            
            If mstrBaby <> "" Then
                .TextMatrix(i, mcol("Ӥ��")) = "���˱���"
            End If
            
            If (rsTmp!ִ�з�ʽ = ִ�з�ʽ.T3��Ҫʱ) And blnFocus = False Then
                Call .Select(i, mcol("ѡ��"))
                blnFocus = True
            End If
            rsTmp.MoveNext
        Next
        
        strIDs = Mid(strIDs, 2)
        '������Ŀ��Ӧ��ҽ��
        str��ѡ���� = ""
        Set rsAdvice = GetAdvice(strIDs)
        If rsAdvice.RecordCount > 0 Then
            For i = .FixedRows To .Rows - 1
                rsAdvice.Filter = "·����ĿID=" & ZVal(Val(.TextMatrix(i, mcol("ID"))))
                strTmp = "": bln���� = False: bln���鳤�� = False: lngOld���ID = 0: str������ĿIDs = ""
                For j = 1 To rsAdvice.RecordCount
                    strTmp = strTmp & "," & rsAdvice!ҽ������ID
                    str������ĿIDs = str������ĿIDs & "," & rsAdvice!������ĿID
                    If rsAdvice!��Ч = 0 Then
                        bln���� = True    'һ�������ͬһ��Ŀ��ҽ����Ч��ͬ,����л��õ������ֻҪ�г�������
                        
                        If mlngFun <> 2 Then
                            If j > 1 And bln���鳤�� = False Then
                                If lngOld���ID <> Val(rsAdvice!���id) Then bln���鳤�� = True
                            End If
                            lngOld���ID = rsAdvice!���id
                        End If
                    End If
                    rsAdvice.MoveNext
                Next
                If strTmp <> "" Then
                    If bln���鳤�� Then
                        If .TextMatrix(i, mcol("����Ҫ��")) = "1" Then str��ѡ���� = str��ѡ���� & "," & .TextMatrix(i, mcol("ID"))
                    End If
                    .TextMatrix(i, mcol("ҽ������ID")) = Mid(strTmp, 2)
                    If mlngFun <> 2 Then .TextMatrix(i, mcol("������ĿID")) = Mid(str������ĿIDs, 2)
                    .TextMatrix(i, mcol("��Ŀ����")) = .TextMatrix(i, mcol("��Ŀ����")) & " ����"
                    .TextMatrix(i, mcol("����")) = IIf(bln����, 1, 0)
                    If bln���� Then .Cell(flexcpForeColor, i, .FixedCols, i, .Cols - 1) = &HC00000      '��ɫ
                    
                    'ǰ���Ѳ����ĳ����Զ���ѡ
                    If bln���� And Not (.RowData(i) = ִ�з�ʽ.T1ÿ����� Or .RowData(i) = ִ�з�ʽ.T0����ִ��) Then
                        Set rsTmp = GetLastAdvice(.TextMatrix(i, mcol("ID")))
                        If rsTmp.RecordCount > 0 Then
                            .Cell(flexcpChecked, i, mcol("ѡ��")) = 1
                        End If
                    End If
                End If
            Next
        End If
        
        '�����ϴ������ɵģ����ڶ����ѡҽ���ģ����ø���Ϊ��ѡ��״̬
        If mlngFun <> 2 Then
            If str��ѡ���� <> "" Then
                bln���鳤�� = False
                Set rsAdvice = GetLastAdvice(str��ѡ����)
                For i = .FixedRows To .Rows - 1
                    rsAdvice.Filter = "��Ŀid=" & .TextMatrix(i, mcol("ID"))
                    If rsAdvice.RecordCount = 0 Then
                        .TextMatrix(i, mcol("��ѡ")) = " "
                        .Cell(flexcpBackColor, i, mcol("��ѡ")) = &H8000000F
                    Else
                        .Cell(flexcpChecked, i, mcol("��ѡ")) = IIf(mlngFun = 3, 1, 2)  '��������ҽ������ѡ����Ŀ���Զ�����
                        .Cell(flexcpPictureAlignment, i, mcol("��ѡ")) = flexPicAlignCenterCenter
                        If mlngFun = 3 Then
                            .Cell(flexcpBackColor, i, mcol("��ѡ")) = &H8000000F
                        Else
                            .Editable = flexEDKbdMouse
                        End If
                        If bln���鳤�� = False Then bln���鳤�� = True
                    End If
                Next
                'û��һ�м�¼����ѡʱ�����ظ���
                If bln���鳤�� = False Then
                    .ColHidden(mcol("��ѡ")) = True
                Else
                    If .ColHidden(mcol("��ѡ")) Then .ColHidden(mcol("��ѡ")) = False
                End If
            Else
                .ColHidden(mcol("��ѡ")) = True
            End If
        End If
        
        '������Ŀ��Ӧ�Ĳ����ļ�
        If mlngFun <> 3 Then
            Set rsFile = GetFile(strIDs)
            If rsFile.RecordCount > 0 Then
                strIDs = ""
                For i = .FixedRows To .Rows - .FixedRows
                    rsFile.Filter = "·����ĿID=" & ZVal(Val(.TextMatrix(i, mcol("ID"))))
                    strTmp = "": strNewTmp = "" '��¼�°���Ӳ���ID
                    For j = 1 To rsFile.RecordCount
                        If rsFile!�ļ�ID & "" <> "" Then
                            strTmp = strTmp & "," & rsFile!�ļ�ID
                            If InStr(strIDs & ",", "," & rsFile!�ļ�ID & ",") = 0 Then
                                On Error Resume Next
                                mEditType.Add Val(rsFile!����), "C" & rsFile!�ļ�ID
                                On Error GoTo errH
                            End If
                        Else
                            strNewTmp = strNewTmp & "," & rsFile!ԭ��ID
                        End If
                        rsFile.MoveNext
                    Next
                    If strTmp <> "" Or strNewTmp <> "" Then
                        strIDs = strIDs & strTmp    'ͬһ��·����Ŀ���ļ�ID�����أ����Էŵ��ڶ���ѭ����
                        .TextMatrix(i, mcol("�ļ�ID")) = IIf(strTmp <> "", Mid(strTmp, 2), "") & "|" & IIf(strNewTmp <> "", Mid(strNewTmp, 2), "")
                        .TextMatrix(i, mcol("��Ŀ����")) = .TextMatrix(i, mcol("��Ŀ����")) & " ��"
                    End If
                    
                Next
            End If
        End If
        If .Rows = .FixedRows Then .Rows = .Rows + 1
        .Redraw = True
    End With
               
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub InitMergeRs()
    If Not mrsMerge Is Nothing Then
        If mrsMerge.State = 1 Then mrsMerge.Close
    End If
    Set mrsMerge = New ADODB.Recordset
    
    mrsMerge.Fields.Append "ID", adBigInt
    mrsMerge.Fields.Append "��id", adBigInt, , adFldIsNullable
    mrsMerge.Fields.Append "���", adBigInt, , adFldIsNullable
    mrsMerge.Fields.Append "����", adVarChar, 100, adFldIsNullable
    mrsMerge.Fields.Append "˵��", adVarChar, 200, adFldIsNullable
    mrsMerge.Fields.Append "��ʼ����", adBigInt, , adFldIsNullable
    mrsMerge.Fields.Append "��������", adBigInt, , adFldIsNullable
    mrsMerge.Fields.Append "����", adVarChar, 50, adFldIsNullable
    mrsMerge.Fields.Append "��֧ID", adBigInt, , adFldIsNullable
    mrsMerge.Fields.Append "·��ID", adBigInt, , adFldIsNullable
    mrsMerge.Fields.Append "�汾��", adBigInt, , adFldIsNullable
    mrsMerge.Fields.Append "·������", adVarChar, 200, adFldIsNullable
    mrsMerge.Fields.Append "��ǰ�׶�ID", adBigInt, , adFldIsNullable
    mrsMerge.Fields.Append "��ǰ����", adBigInt, , adFldIsNullable
    mrsMerge.Fields.Append "�ϲ�·����¼ID", adBigInt, , adFldIsNullable
    
    mrsMerge.CursorLocation = adUseClient
    mrsMerge.LockType = adLockOptimistic
    mrsMerge.CursorType = adOpenStatic
    mrsMerge.Open
End Sub

Private Sub LoadMerge()
'���ܣ����غϲ�·����Ŀ
    Dim i As Long
    Dim rsTmp As ADODB.Recordset, strSql As String
    Dim strTmp As String
    Dim rsMerge As ADODB.Recordset
    Dim lngDay As Long
    
    strSql = "Select a.id,b.����,a.�汾��,b.ID as ·��ID,NVL(a.��ǰ����,0) as ��ǰ����,a.��ǰ�׶�ID,c.��֧ID as ��ǰ�׶η�֧ID " & _
            " From ���˺ϲ�·�� A,�ٴ�·��Ŀ¼ B,�ٴ�·���׶� C " & _
            " Where a.·��ID=b.id And a.��ǰ�׶�ID = c.ID(+) And a.����ʱ�� is null And a.��Ҫ·����¼ID=[1] order by a.����ʱ��"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.����·��ID)
    mlngMergeCount = rsTmp.RecordCount
    Call InitMergeRs
    vsItem(1).Rows = vsItem(1).FixedRows
    mstrMergeStep = ""
    If rsTmp.RecordCount > 0 Then
        Do While Not rsTmp.EOF
            strTmp = strTmp & "," & rsTmp!����
            '�����Ҫ·���ĵ�ǰ������������������������(��Ҫ·���Ӻ����ǰ���ϲ�·��Ҳһ��)
            lngDay = mlng���� - mPP.��ǰ����
            '��ȡ�ϲ�·���׶�
            Set rsMerge = GetPhase(Val(rsTmp!·��ID & ""), Val(rsTmp!�汾�� & ""), Val(rsTmp!��ǰ�׶�ID & ""), Val(rsTmp!��ǰ�׶η�֧ID & ""), Val(rsTmp!��ǰ���� & "") + lngDay, Val(rsTmp!ID & ""))
            If rsMerge.RecordCount > 0 Then
                Do While Not rsMerge.EOF
                    mrsMerge.AddNew
                    mrsMerge!ID = rsMerge!ID
                    mrsMerge!��ID = rsMerge!��ID
                    mrsMerge!��� = rsMerge!���
                    mrsMerge!���� = rsMerge!����
                    mrsMerge!˵�� = rsMerge!˵��
                    mrsMerge!��ʼ���� = rsMerge!��ʼ����
                    mrsMerge!�������� = rsMerge!��������
                    mrsMerge!���� = rsMerge!����
                    mrsMerge!��֧ID = rsMerge!��֧ID
                    mrsMerge!·��ID = rsTmp!·��ID
                    mrsMerge!�汾�� = rsTmp!�汾��
                    mrsMerge!·������ = rsTmp!����
                    mrsMerge!��ǰ�׶�ID = rsTmp!��ǰ�׶�ID
                    mrsMerge!��ǰ���� = rsTmp!��ǰ����
                    mrsMerge!�ϲ�·����¼ID = rsTmp!ID
                    mrsMerge.Update
                    rsMerge.MoveNext
                Loop
                rsMerge.MoveFirst
                Call LoadItem(Val(rsMerge!ID & ""), vsItem(1), Val(rsTmp!·��ID & ""), Val(rsTmp!�汾�� & ""), Val(rsTmp!��ǰ���� & "") + lngDay, Val(rsTmp!ID & ""))
                mstrMerge = mstrMerge & "," & rsMerge!��֧ID & ":" & Val(rsMerge!ID & "")
            End If
            rsTmp.MoveNext
        Loop
        mstrMerge = Mid(mstrMerge, 2)
        lblMerge.Caption = lblMerge.Caption & Mid(strTmp, 2)
        '�����ǰû�п��úϲ�·���׶�,������
        If mrsMerge.RecordCount = 0 Then
            lblMerge.Visible = False
            vsItem(1).Visible = False
            cmdMergeStep.Visible = False
            vsItem(0).Height = vsItem(1).Top + vsItem(1).Height - vsItem(0).Top
        End If
    Else
        'û�кϲ�·���������ذ�ť�ͱ��
        lblMerge.Visible = False
        vsItem(1).Visible = False
        cmdMergeStep.Visible = False
        vsItem(0).Height = vsItem(1).Top + vsItem(1).Height - vsItem(0).Top
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub InitItem()
'����: ��ʼ��·����Ŀ��ͷ
    Dim strcol As String, arrHead As Variant
    Dim i As Long
    
    If mlngFun = 2 Then
        strcol = "����,1200,4;����ֵ;ȫѡ,450,4;��Ŀ����,5950,1;ִ�з�ʽ;ѡ��;Ӥ��;ID;ҽ������ID;����;�ļ�ID"
    Else
        strcol = "����,1200,4;����ֵ;ȫѡ" & IIf(mlngFun <> Func��������, ",450,4", "") & ";��Ŀ����," & IIf(mstrBaby = "", 5950, 4950) & ",1;ִ�з�ʽ,900,1" & _
                ";ѡ��" & IIf(mlngFun <> Func��������, ",500,4", "") & _
                ";��ѡ,500,4;Ӥ��" & IIf(mstrBaby = "", "", ",1100,1") & _
                ";ID;ҽ������ID;����;�ļ�ID;����Ҫ��;����ԭ��" & IIf(mlngFun = 0, ",1800,4", "") & ";�Ƿ����һ��;�׶�ID;�ϲ�·����¼ID;������ĿID;�ظ���Ŀ"
    End If
    arrHead = Split(strcol, ";")
    Set mcol = New Collection
   
    With vsItem(0)
        .Redraw = flexRDNone
        .Clear
        .FixedRows = 1: .FixedCols = 1
        .Cols = UBound(arrHead) + 1
        .Rows = .FixedRows
        
        For i = 0 To UBound(arrHead)
            mcol.Add i, Split(arrHead(i), ",")(0)
            .TextMatrix(.FixedRows - 1, i) = Split(arrHead(i), ",")(0)
            
            If UBound(Split(arrHead(i), ",")) > 0 Then
                .ColHidden(i) = False
                .ColWidth(i) = Val(Split(arrHead(i), ",")(1))
                .ColAlignment(i) = Val(Split(arrHead(i), ",")(2))
            Else
                .ColHidden(i) = True
                .ColWidth(i) = 0
            End If
        Next
        .Redraw = True
    End With
    With vsItem(1)
        .Redraw = flexRDNone
        .Clear
        .FixedRows = 1: .FixedCols = 1
        .Cols = UBound(arrHead) + 1
        .Rows = .FixedRows
        
        For i = 0 To UBound(arrHead)
            .TextMatrix(.FixedRows - 1, i) = Split(arrHead(i), ",")(0)
            
            If UBound(Split(arrHead(i), ",")) > 0 Then
                .ColHidden(i) = False
                .ColWidth(i) = Val(Split(arrHead(i), ",")(1))
                .ColAlignment(i) = Val(Split(arrHead(i), ",")(2))
            Else
                .ColHidden(i) = True
                .ColWidth(i) = 0
            End If
        Next
        .Redraw = True
    End With
End Sub

Private Sub vsPhase_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If vsPhase.MouseCol >= 0 And vsPhase.MouseRow >= 0 Then
        Dim strInfo As String
        strInfo = Trim(vsPhase.Cell(flexcpData, vsPhase.MouseRow, vsPhase.MouseCol))
        Call zlCommFun.ShowTipInfo(vsPhase.Hwnd, strInfo)
    End If
End Sub


Private Function GetBabyRegList() As String
'���ܣ���ȡ���˵�Ӥ�������б�
'������
'���أ�"����1,����2,����3��
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errH
    
    strSql = "Select ���,Ӥ������ From ������������¼ Where ����ID=[1] And ��ҳID=[2] Order by ���"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "GetBabyRegList", mPati.����ID, mPati.��ҳID)
    
    strSql = ""
    Do While Not rsTmp.EOF
        strSql = IIf(strSql = "", "", strSql & "|") & "Ӥ��:" & Nvl(Replace(rsTmp!Ӥ������, "|", "_"))
        rsTmp.MoveNext
    Loop
    GetBabyRegList = strSql
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetBabyIndex(strtxt As String) As String
'���ܣ����ݵ�ǰ�е����ݷ���Ӥ�����
    Dim i As Long, j As Long
    Dim arrtmp As Variant
    Dim arrBaby As Variant
    Dim strBaby As String
    
    If mstrBaby <> "" Then
        arrtmp = Split("���˱���|" & mstrBaby, "|")
        arrBaby = Split(strtxt, "|")
        For j = 0 To UBound(arrBaby)
            For i = 0 To UBound(arrtmp)
                If arrtmp(i) = arrBaby(j) Then
                    strBaby = strBaby & "|" & i
                    Exit For
                End If
            Next
        Next
        GetBabyIndex = Mid(strBaby, 2)
    Else
        GetBabyIndex = "0"  'û��Ӥ��ȱʡȡ���˱���
    End If
End Function

Private Function UnExecutedOfPhase(ByVal lng��ĿID As Long) As Boolean
'���ܣ����ָ������Ŀ�ڵ�ǰ�׶��Ƿ�ִ�й�
    Dim rsTmp As ADODB.Recordset, strSql As String
 
    strSql = "Select 1 From ����·��ִ�� Where ·����¼ID = [1] And �׶�ID = [2] And ��ĿID = [3]"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.����·��ID, Val(vsPhase.ColData(vsPhase.Col)), lng��ĿID)
    UnExecutedOfPhase = rsTmp.RecordCount = 0
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Sub cmdCancel_Click()
    mblnOK = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim i As Long, j As Long, blnEnd As Boolean, strIDs As String, strAdviceOfItem As String
    Dim arrSQL As Variant, arrBaby As Variant, DatCurr As Date
    Dim strTmp As String, strBaby As String, strBB As String, lng���� As Long
    Dim rsTmp As ADODB.Recordset, rsLastAdvice As ADODB.Recordset
    Dim rsUsed As ADODB.Recordset   '��������ʱУ�Ե�δ���ϵ�ҽ��
    Dim blnHave As Boolean, blnHaveDoc As Boolean
    Dim strLAdivceOfItem As String  '�ϴ�����·����Ŀ������ID
    Dim strLAdvices As String       '�ϴ����ɵĳ���ID
    Dim str��ĿIDs As String, str��ѡ��ĿIDs As String, strҽ��IDs As String
    Dim k As Long, n As Long, strAgain As String
    Dim colItem As New Collection
    Dim strAgaignTmp As String
    Dim str·����ĿIDs As String   '·������ʱ��ҽ�޸��˵��䷽�ģ��ҳ����������޸��䷽�ı�������Ŀ����Ӧ�ı���ԭ����ĿID1|�������1,��Ŀ2|�������2��������
    Dim colPathItems As New Collection
    
    arrSQL = Array()
    '1.������ִ��һ�ε���Ŀ
    For k = 0 To vsItem.count - 1
        With vsItem(k)
            If mlngFun = 0 Then
                If k = 0 Then
                    With mrsPhase
                        .Filter = "ID=" & vsPhase.ColData(vsPhase.Col)
                        If Not IsNull(!��ʼ����) Then
                            If IsNull(!��������) Then
                                blnEnd = (Val(!��ʼ����) = mlng����)
                            Else
                                blnEnd = (Val(!��������) = mlng����)
                            End If
                        End If
                    End With
                End If
                '�ϲ�·�����ڿ����Ƕ���׶���Ŀ�������Ƿ����һ�죬�����У��Ƿ����һ�����ж�
                If blnEnd Or k = 1 Then
                    For i = 1 To .Rows - 1
                        If k = 0 Or .TextMatrix(i, mcol("�Ƿ����һ��")) = "1" Then
                            If .RowData(i) = ִ�з�ʽ.T2����һ�� Or .RowData(i) = ִ�з�ʽ.T4�����ҽ�һ�� Then
                                If .Cell(flexcpChecked, i, mcol("ѡ��")) = 2 Then
                                    If UnExecutedOfPhase(Val(.TextMatrix(i, mcol("ID")))) Then
                                        .Cell(flexcpChecked, i, mcol("ѡ��")) = 1
                                        .Row = i
                                        blnHave = True
                                    End If
                                End If
                            End If
                        End If
                    Next
                End If
                If k = vsItem.count - 1 And blnHave Then
                    MsgBox "���׶����ٻ�������һ�ε���Ŀû��ѡ��ϵͳ���Զ�ѡ������ȷ�Ϻ������", vbInformation, gstrSysName
                    Exit Sub
                End If
                
                'ÿ�����ɵ���Ŀ�����û��ѡ��������������ԭ��
                For i = 1 To .Rows - 1
                    If .RowData(i) = ִ�з�ʽ.T1ÿ����� Then
                        If .Cell(flexcpChecked, i, mcol("ѡ��")) = 2 Then
                            If .TextMatrix(i, mcol("����ԭ��")) = "" Then
                                MsgBox "�������ɵ���Ŀ�����ѡ�����ɣ���Ҫ�����ѡ�����ԭ��", vbInformation, gstrSysName
                                If .Visible And .Enabled Then .SetFocus
                                .Select i, mcol("����ԭ��")
                                Exit Sub
                            End If
                        End If
                    End If
                Next
            End If
        End With
        blnEnd = False
    Next
    
    '2.��ȡ·����Ŀ��Ӧ��ҽ��
    If mPP.��ǰ���� > 0 Then
        For k = 0 To vsItem.count - 1
            With vsItem(k)
                For i = .FixedRows To .Rows - 1
                    If .Cell(flexcpChecked, i, mcol("ѡ��")) = 1 Then
                        If Val(.TextMatrix(i, mcol("����"))) = 1 Then
                            '���ѡ������ѡҽ��������Щ����Ŀ��Ӧ��ҽ��Ҫ��ʾ��������ҽ���༭���棬����ѡ�����ǰһ���ģ�
                            'ҽ���´���汣��ʱ���²���ҽ����¼�����ݣ�����Ҫ����ǰ��ҽ��ID�ռ��������·��ҽ����
                            If .Cell(flexcpChecked, i, mcol("��ѡ")) = 1 Then
                                str��ѡ��ĿIDs = str��ѡ��ĿIDs & "," & .TextMatrix(i, mcol("ID"))
                            Else
                                str��ĿIDs = str��ĿIDs & "," & .TextMatrix(i, mcol("ID"))
                            End If
                        End If
                    End If
                Next
            End With
        Next
        If str��ѡ��ĿIDs <> "" Then
            str��ѡ��ĿIDs = Mid(str��ѡ��ĿIDs, 2)
            Set rsLastAdvice = GetLastAdvice(str��ѡ��ĿIDs)    '�ϴ�������ҽ���ļ�¼�������ڴ��뵽ҽ���´ﴰ�壬����ʱ����Ƿ������µ�ҽ��
        End If
        If str��ĿIDs <> "" Then
            str��ĿIDs = Mid(str��ĿIDs, 2)
            Set rsTmp = GetLastAdvice(str��ĿIDs) '����׶β�ͬ�����ε���ĿID��ǰ�εĲ�һ��,ֻ��������ͬ
            
            str��ĿIDs = ""
            strLAdivceOfItem = ""
            strLAdvices = ""
            For i = 1 To rsTmp.RecordCount
                'ǰһ�������˵ľͲ����ظ�����ҽ������Ҫ����ǰ��ҽ��ID�ռ��������·��ҽ����
                strLAdivceOfItem = strLAdivceOfItem & "," & rsTmp!��ĿID & ":" & rsTmp!����ҽ��id
                strLAdvices = strLAdvices & "," & rsTmp!����ҽ��id
                '�ռ�ǰһ���������˳���ҽ������ĿID
                If InStr("," & str��ĿIDs & ",", "," & rsTmp!��ĿID & ",") = 0 Then
                    str��ĿIDs = str��ĿIDs & "," & rsTmp!��ĿID
                End If
                rsTmp.MoveNext
            Next
            strLAdivceOfItem = Mid(strLAdivceOfItem, 2)
            strLAdvices = Mid(strLAdvices, 2)
            str��ĿIDs = str��ĿIDs & ","   '���β���������ҽ����·����ĿID
        End If
    End If
    '91635 ��������ҽ����
    '1�����һ������������Ŀ��Ӧ��ҽ���д����Ѿ�У�Ե�δ���ϵ�ҽ����¼ʱ,�����û���������ҽ��,��У��δ���ϵ�ҽ�����ֲ��䡣
    '2�������������������Ŀ��Ӧ��ҽ������δУ�Ե�,��ɾ������Ŀ��Ӧ������ҽ��,���²���ҽ����¼���ݡ�
    '��������ʱ,�Ѿ�У�Ե�δ���ϵ�ҽ����¼��,���ڴ��˵�ҽ���´ﴰ��,����ʱ����Ƿ������µ�ҽ��
    If mlngFun = Func�������� Then
        Set rsUsed = GetUsedAdvice(mlngִ��ID, mlng��ĿID)
        If rsUsed.RecordCount > 0 Then
            If rsLastAdvice Is Nothing Then
                Set rsLastAdvice = rsUsed
            Else
                For i = 1 To rsUsed.RecordCount
                    rsLastAdvice.Filter = "��ĿID =" & rsUsed!��ĿID & " And ��ID = " & rsUsed!��ID
                    If rsLastAdvice.RecordCount = 0 Then
                        rsLastAdvice.AddNew
                        rsLastAdvice!��ĿID = rsUsed!��ĿID
                        rsLastAdvice!����ҽ��id = rsUsed!����ҽ��id
                        rsLastAdvice!��ID = rsUsed!��ID
                        rsLastAdvice!������ĿID = rsUsed!������ĿID
                        rsLastAdvice.Update
                    End If
                    rsUsed.MoveNext
                Next
            End If
        End If
    End If
    
    strIDs = ""
    For k = 0 To vsItem.count - 1
        With vsItem(k)
            '�������ɵģ����ѡ������ҽ����ѡ���˱���ԭ�򣩣���Ҫ����·����Ŀ
            If mlngFun = 0 Then
                For i = 1 To .Rows - 1
                    If .RowData(i) = ִ�з�ʽ.T1ÿ����� Then
                        If .Cell(flexcpChecked, i, mcol("ѡ��")) = 2 Then
                            If InStr("," & str��ĿIDs & ",", "," & .TextMatrix(i, mcol("ID")) & ",") = 0 Then
                                str��ĿIDs = str��ĿIDs & "," & .TextMatrix(i, mcol("ID"))
                            End If
                        End If
                    End If
                Next
            End If
            
            '����Ҫ����ҽ������ĿID����ǰ�������ɳ�������Ŀ����������,�������ɵĶ�ѡ���˲����ɵĲ�������
            For i = 1 To .Rows - 1
                .TextMatrix(i, mcol("�ظ���Ŀ")) = ""
                If .Cell(flexcpChecked, i, mcol("ѡ��")) = 1 Then
                    If InStr(str��ĿIDs, "," & .TextMatrix(i, mcol("ID")) & ",") = 0 Then
                        strTmp = Trim(.TextMatrix(i, mcol("ҽ������ID")))
                        If strTmp <> "" Then
                            '��Ŀ������ͬ�Ҷ�Ӧ��ҽ��Ҳ��ͬ�����ظ�����
                            strAgaignTmp = Trim(.TextMatrix(i, mcol("������ĿID")))
                            '��¼ҽ�����ж��ظ�
                            If InStr(strAgain & vbCrLf, vbCrLf & strAgaignTmp & vbCrLf) = 0 Or strAgaignTmp = "" Then
                                strAgain = strAgain & vbCrLf & strAgaignTmp
                                If strAgaignTmp <> "" Then
                                    colItem.Add .TextMatrix(i, mcol("ID")) & vbCrLf & .TextMatrix(i, mcol("��Ŀ����")), strAgaignTmp
                                End If
                                arrBaby = Split(GetBabyIndex(.TextMatrix(i, mcol("Ӥ��"))), "|")
                                For n = LBound(arrBaby) To UBound(arrBaby)
                                    strBB = arrBaby(n) & ":" & .TextMatrix(i, mcol("ID"))
                                    If InStr(strTmp, ",") = 0 Then
                                         strBaby = strTmp & ":" & strBB
                                     Else
                                         strBaby = Replace(strTmp, ",", ":" & strBB & ",") & ":" & strBB
                                     End If
                                     strIDs = strIDs & "," & strBaby
                                Next
                            Else
                                .TextMatrix(i, mcol("�ظ���Ŀ")) = colItem(strAgaignTmp)
                            End If
                        End If
                    End If
                    If blnHaveDoc = False Then
                        If Trim(.TextMatrix(i, mcol("�ļ�ID"))) <> "" Then blnHaveDoc = True
                    End If
                End If
            Next
        End With
    Next
    strIDs = Mid(strIDs, 2) 'ҽ������ID:Ӥ�����:·����ĿID,...������227:0:38,335:1:69
    
    If blnHaveDoc Then
        If InStr(GetInsidePrivs(pסԺ��������), ";������д;") = 0 Then
            MsgBox "��û�в�����д��Ȩ�ޣ��������ɰ���������·����Ŀ��", vbInformation + vbOKOnly, gstrSysName
            Exit Sub
        End If
    End If
    
    
    '����ҽ����ȱʡ��ʼִ��ʱ��
    DatCurr = mdatDur
        
    If strIDs <> "" Then    'ȫ������ִ�е���Ŀʱ������ҽ������Ҫ����·��ִ����Ŀ
        If InStr(GetInsidePrivs(pסԺҽ���´�), ";ҽ���´�;") = 0 Then
            MsgBox "��û��ҽ���´��Ȩ�ޣ��������ɰ���ҽ����·����Ŀ��", vbInformation + vbOKOnly, gstrSysName
            Exit Sub
        End If
        '���ʱ��
        If Format(DatCurr, "YYYY-MM-DD") > Format(dtpAdviceTime.Value, "YYYY-MM-DD") Or Format(dtpAdviceTime.Value, "YYYY-MM-DD") > Format(DatCurr + mlng·��ҽ������, "YYYY-MM-DD") Then
            MsgBox "�ٴ�·����ҽ�������ڵ�ǰ���ں���ǰ������֮�䣬��ǰ������ǰ" & mlng·��ҽ������ & "�졣", vbInformation, gstrSysName
            If dtpAdviceTime.Enabled And dtpAdviceTime.Visible Then dtpAdviceTime.SetFocus
            Exit Sub
        End If
        
        Me.Hide
        If gobjKernel.ShowAdviceEdit(mfrmParent, mint����, 1, mPati.����ID, mPati.��ҳID, strIDs, CDate(dtpAdviceTime.Value), arrSQL, strAdviceOfItem, rsLastAdvice, DatCurr, str·����ĿIDs, mclsMipModule) = False Then
            Unload Me
            Exit Sub
        End If
        '�����ҽ�䷽�޸ĵ�ζ���������õı�׼�����������˱���ԭ�����ϱ���ԭ��
        If str·����ĿIDs <> "" Then
            '���һ����Ŀ����������ԭ����ȡ��һ��
            On Error Resume Next
            For i = 0 To UBound(Split(str·����ĿIDs, ","))
                colPathItems.Add Split(Split(str·����ĿIDs, ",")(i), "|")(1), "_" & Split(Split(str·����ĿIDs, ",")(i), "|")(0)
            Next
            For i = 1 To vsItem(0).Rows - 1
                strTmp = ""
                If vsItem(0).TextMatrix(i, mcol("ID")) & "" <> "" Then
                    strTmp = colPathItems("_" & vsItem(0).TextMatrix(i, mcol("ID")))
                    If strTmp <> "" Then
                        vsItem(0).Cell(flexcpData, i, mcol("����ԭ��")) = strTmp
                    End If
                End If
            Next
            For i = 1 To vsItem(1).Rows - 1
                strTmp = ""
                If vsItem(1).TextMatrix(i, mcol("ID")) & "" <> "" Then
                    strTmp = colPathItems("_" & vsItem(1).TextMatrix(i, mcol("ID")))
                    If strTmp <> "" Then
                        vsItem(1).Cell(flexcpData, i, mcol("����ԭ��")) = strTmp
                    End If
                End If
            Next
            On Error GoTo 0
        End If
    End If
    
    strҽ��IDs = ""  '��¼��ֹͣ�ĳ���ID
    If (mlngFun = 0 Or mlngFun = 3) Then
        If InStr(GetInsidePrivs(pסԺҽ���´�), ";ҽ��ֹͣ;") > 0 Then
            Set rsTmp = GetLastAdvice(, "," & strLAdvices & ",")
            For i = 1 To rsTmp.RecordCount
                strҽ��IDs = strҽ��IDs & "," & rsTmp!����ҽ��id
                rsTmp.MoveNext
            Next
            strҽ��IDs = Mid(strҽ��IDs, 2)
        End If
    End If
    
    '��Ҫ���������ĳ����Ĳ���·��ҽ��
    If strLAdivceOfItem <> "" Then
        If strAdviceOfItem = "" Then
            strAdviceOfItem = strLAdivceOfItem
        Else
            strAdviceOfItem = strAdviceOfItem & "," & strLAdivceOfItem
        End If
    End If
    '��������ʱ�ռ�У��δͣ�õ�ҽ��Id����"����·��ҽ��"
    If mlngFun = Func�������� Then
        rsUsed.Filter = ""
        For i = 1 To rsUsed.RecordCount
            If i = 1 And strAdviceOfItem = "" Then
                strAdviceOfItem = mlng��ĿID & ":" & rsUsed!����ҽ��id
            Else
                If InStr("," & strAdviceOfItem & ",", "," & mlng��ĿID & ":" & rsUsed!����ҽ��id & ",") = 0 Then '�����ظ���ӣ��п���ҽ���´�����Ѿ������˸����ݣ�
                    strAdviceOfItem = strAdviceOfItem & "," & mlng��ĿID & ":" & rsUsed!����ҽ��id
                End If
            End If
            rsUsed.MoveNext
        Next
    End If
    Call SaveData(arrSQL, strAdviceOfItem, lng����)
    '����ҽ�����
    Call ModifyAdviceSerialNum
    
    '����·���󣬼���Ƿ�����Ҫֹͣ�ĳ���(�ϴ��У�������û�еĳ���)
    '-----------------------------------------------------------------------------
    If strҽ��IDs <> "" Then
        '�����������û����Щ����������Ҫֹͣ
         strIDs = GetShouldStopAdvice(strҽ��IDs, lng����)
         If strIDs <> "" Then
            Me.Hide
            Call gobjKernel.ShowAdviceOperate(mfrmParent, mint����, mPati.����ID, mPati.��ҳID, mPati.����ID, strIDs, DateAdd("s", 1, DatCurr), mclsMipModule)
            
            Call CheckStopAdvice(mPati.����ID, mPati.��ҳID, strIDs)
            'ҽ��û��ֹͣ�ĳ�������Ҫ����Ϊ·������Ŀ
            If strIDs <> "" Then
            Call AddOutPathItem(strIDs, 1, mPati.����ID, mPati.��ҳID)
         End If
         End If
         
    End If
    
    mblnOK = True
    Unload Me
End Sub

Private Sub ModifyAdviceSerialNum()
'���ܣ���������ҽ�����
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errH
    Screen.MousePointer = 11
    strSql = "Select Count(*) as Num From (Select ���,Count(ID) From ����ҽ����¼ Where ����ID=[1] And ��ҳID=[2] Having Count(ID)>1 Group by ���)"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "��ѯ����ҽ������", mPati.����ID, mPati.��ҳID)
    
    If rsTmp.EOF Then Screen.MousePointer = 0: Exit Sub
    
    If Nvl(rsTmp!Num, 0) = 0 Then Screen.MousePointer = 0: Exit Sub
    
    strSql = "ZL_����ҽ����¼_�������(NULL,NULL," & mPati.����ID & "," & mPati.��ҳID & ")"
    Call zlDatabase.ExecuteProcedure(strSql, "����ҽ�����")
    
    Screen.MousePointer = 0
    Exit Sub
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub GetPreSendData(ByRef lng�׶�ID As Long, ByRef dat���� As Date)
'���ܣ����ݵ�ǰ�׶κ����ڷ�����һ������·����Ŀ�Ľ׶κ�����
    Dim strSql As String, rsTmp As ADODB.Recordset

    On Error GoTo errH
    strSql = "Select �׶�id, ����, ����" & vbNewLine & _
             "From ����·��ִ��" & vbNewLine & _
             "Where ·����¼id = [1] And �Ǽ�ʱ�� = (Select Max(�Ǽ�ʱ��)" & vbNewLine & _
             "                             From ����·��ִ��" & vbNewLine & _
             "                             Where ·����¼id = [1] And �Ǽ�ʱ�� <  (Select Min(�Ǽ�ʱ��) �Ǽ�ʱ��" & vbNewLine & _
             "                                    From ����·��ִ�� A" & vbNewLine & _
             "                                    Where a.·����¼id = [1] And a.�׶�id = [2] And a.���� = [3])" & vbNewLine & _
             "                             ) And Rownum < 2"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.����·��ID, mPP.��ǰ�׶�ID, CDate(mPP.��ǰ����))
    If rsTmp.RecordCount > 0 Then    '���������ǵ�һ�����ɵ�����޼�¼
        lng�׶�ID = rsTmp!�׶�ID
        dat���� = rsTmp!����
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Function GetLastAdvice(Optional ByVal strIDsOfLA As String, Optional ByVal strLAdvices As String) As ADODB.Recordset
'���ܣ���ȡ·����Ŀ�����һ�����ɵĳ���ҽ������Ŀ��¼��
'������strIDsOfLA=��ǰѡ��ĺ��г���ҽ������ĿID��
'    :strLAdvices=ȡ�ϴε�������Ч�ĳ���(��У�ԣ�������������ͣ��������),�ų�����ҽ������ǰ�����󣬲�Σ�����أ���¼�������,����ȼ�
'���أ�1.���ص�·����ĿID�Ǳ������ɵ�·����Ŀ��ID;2-����ҽ��Id�Ǳ���������Ҫֹͣ�ĳ���ID
    Dim strSql As String
    Dim lng�׶�ID As Long, dat���� As Date, lng���� As Long
    
    '�ҵ�ǰ����ִ�еĳ���ҽ��
    If strIDsOfLA <> "" Then
        '�����뵱ǰ��Ŀͬ���ģ�����һ��ִ�У�ǰһ���ͬһ�죩�����˳���ҽ����(��δ���ϻ�ֹͣ��)���򱾴β��ظ�����
        '��Ŀid����������ȷ����·��id���汾��
        strSql = "Select /*+ rule*/ f.id as ��Ŀid, b.����ҽ��id,Nvl(d.���id,d.id) ��ID,d.������ĿID" & vbNewLine & _
            "From ����·��ִ�� A, ����·��ҽ�� B, ����ҽ����¼ D, �ٴ�·����Ŀ E, �ٴ�·����Ŀ F" & vbNewLine & _
            IIf(InStr(strIDsOfLA, ",") > 0, ",(Select Column_Value As ��Ŀid From Table(f_Num2list([1]))) C Where c.��Ŀid = f.Id ", " Where f.id = [1]") & vbNewLine & _
            "     And f.��Ŀ���� = e.��Ŀ���� And e.Id = a.��Ŀid And a.·����¼id = [2] And" & vbNewLine & _
            "      a.�׶�id = [3] And a.���� = [4] And a.Id = b.·��ִ��id And b.����ҽ��id = d.Id And d.ҽ����Ч = 0 And d.ҽ��״̬ Not In(4,8,9)" & vbNewLine & _
            " Group By f.Id, b.����ҽ��id, Nvl(d.���id, d.Id), d.������Ŀid, d.��� " & _
            " Order by d.���"
    Else
        '�ų�����ȼ�����ҽ��ֹͣ���汣��һ�£�
        strSql = "Select b.����ҽ��id" & vbNewLine & _
                "From ����·��ִ�� A, ����·��ҽ�� B, ����ҽ����¼ C,������ĿĿ¼ D" & vbNewLine & _
                "Where a.·����¼id = [2] And a.�׶�id = [3] And a.���� = [4] And a.Id = b.·��ִ��id And b.����ҽ��id = c.Id And c.ҽ����Ч = 0 And" & vbNewLine & _
                "      c.ҽ��״̬ In (3, 5, 6, 7) And C.������ĿID=D.ID And Not(D.���='H' and D.��������='1' And D.ִ��Ƶ��=2) " & _
                "   And Not(D.���='Z' And D.�������� IN('4','14', '9', '10', '12')) And instr( '" & strLAdvices & "',','|| b.����ҽ��ID||',')=0"
    End If
    On Error GoTo errH
    If mlngFun = 0 Then
        lng�׶�ID = mPP.��ǰ�׶�ID
        dat���� = CDate(mPP.��ǰ����)
    Else
    '�������ɣ���������ʱ��ȡǰ�����ɵĽ׶κ�����
        Call GetPreSendData(lng�׶�ID, dat����)
    End If
    Set GetLastAdvice = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strIDsOfLA, mPP.����·��ID, lng�׶�ID, dat����)
        
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetUsedAdvice(ByVal lngִ��ID As Long, ByVal lng��ĿID As Long) As ADODB.Recordset
'����:��������ʱ,���ص�ǰ��Ŀ��У�Ե�δ���ϵ�ҽ����¼
    Dim strSql As String
    
    strSql = "Select [1] As ��Ŀid, a.����ҽ��id, Nvl(b.���id, b.Id) As ��id, b.������Ŀid" & vbNewLine & _
            "From ����·��ҽ�� A, ����ҽ����¼ B" & vbNewLine & _
            "Where a.����ҽ��id = b.Id And a.·��ִ��id = [2] And b.ҽ��״̬ > 1 And b.ҽ��״̬ <> 4" & vbNewLine & _
            "Order By b.���"
    On Error GoTo errH
    
    Set GetUsedAdvice = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng��ĿID, lngִ��ID)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub GetLastEvaluate(strLastVariation As String, str����� As String, str������ As String)
'���ܣ�������һ����������Ϣ
    Dim strSql As String, rsTmp As Recordset
    
    On Error GoTo errH
    strSql = "Select ����ԭ��,���������,������ From ����·������ Where ·����¼ID=[1] And ����=[2] And �׶�ID=[3]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.����·��ID, mPP.��ǰ����, mPP.��ǰ�׶�ID)
    If rsTmp.RecordCount > 0 Then
        strLastVariation = rsTmp!����ԭ�� & ""
        str����� = rsTmp!��������� & ""
        str������ = rsTmp!������ & ""
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function GetFirstType()
'���ܣ����һ����Ŀ��û�У�������ݿ���ȡ��һ������
    Dim rsTmp As ADODB.Recordset, strSql As String
    
    On Error GoTo errH
    strSql = "Select ���� from �ٴ�·������ where ·��ID=[1] and �汾��=[2] and NVL(��֧ID,0)=[3] And ���=1"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "ȡ��һ������", mPP.·��ID, mPP.�汾��, Val(Mid(tabBranch.SelectedItem.Key, 2)))
    If rsTmp.RecordCount > 0 Then GetFirstType = rsTmp!���� & ""
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub SaveData(ByVal arrSQL As Variant, ByVal strAdviceOfItem As String, ByRef lng���� As Long)
'���ܣ�����·����Ŀ
'������strAdviceOfItem=·����Ŀ��ҽ��ID�Ķ�Ӧ,����38:1983,69:1978
    Dim colSQL As New Collection, colDoc As New Collection, blnTrans As Boolean, colNewDoc As New Collection
    Dim strSql As String, i As Long, j As Long, l As Long, k As Long, lngBaby As Long
    Dim strDate As String, strAddDate As String, strAdviceIDs As String, strFileIDs As String, strFileID As String
    Dim str���˲���IDs As String, strEMRID As String, lng����ID As Long, strVariation As String, strBaby As String
    Dim strFileIDsTmp As String, strFiles As String
    Dim arrItem As Variant, lng��� As Long
    Dim blnIsSend As Boolean   '�ж��û��Ƿ�ѡ����Ŀ
    Dim strLastVariation As String
    Dim str����� As String
    Dim str������ As String
    Dim varFilter As Variant
    Dim AddDate As Date
    Dim strFirstType As String
    Dim strMergeStep As String
    Dim strEPR As String
    Dim blnAgain As Boolean
    Dim strAgaignTmp As String
    Dim strAgain As String
    Dim colItemName As New Collection
    Dim blnDef As Boolean
    Dim strPara As String, strParaTmp As String
    Dim rsTmp As ADODB.Recordset
    Dim arrtmp As Variant
    
    AddDate = zlDatabase.Currentdate
    strAddDate = "To_Date('" & Format(AddDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
    arrItem = Split(strAdviceOfItem, ",")
    
    '���һ����Ŀ��û�У�������ݿ���ȡ��һ������
    If vsItem(0).TextMatrix(1, mcol("����")) = "" Then
        strFirstType = GetFirstType
    Else
        strFirstType = vsItem(0).TextMatrix(1, mcol("����"))
    End If

    strMergeStep = Mid(mstrMergeStep, 2)
    lng���� = mlng����
    
    strDate = "To_Date('" & Format(mdatʱ��, "yyyy-MM-dd") & "','YYYY-MM-DD')"
    'mrsPhase����ǰ��ִ��filter
    '��ǰ����
    If mlngʱ����� = 2 Then
        k = mlng��ǰ����
    Else
        k = mlng��ǰ���� + 1
    End If
     '�жϵ�ǰѡ��Ľ׶ο�ʼ�����Ƿ���ڴ��������������м��������δ�����κ���Ŀ������
    If lng���� > k Then
        varFilter = mrsPhase.Filter
        Call GetLastEvaluate(strLastVariation, str�����, str������)
        For i = k To lng���� - 1
            '����
            mrsPhase.Filter = "��ʼ����=" & i & " And ��ID = 0" & IIf(mblnIsHaveBranch, " And ��֧ID=" & Mid(tabBranch.SelectedItem.Key, 2), "")
            If Not mrsPhase.EOF Then
                strSql = "Zl_����·������_Insert(1," & mPati.����ID & "," & mPati.��ҳID & ",NULL," & mPati.����ID & "," & _
                        mPP.����·��ID & "," & mrsPhase!ID & _
                        "," & strDate & "," & mlng��ǰ���� & _
                        ",'" & strFirstType & "',Null" & _
                        ",Null,Null,Null,'" & UserInfo.���� & "',To_Date('" & Format(AddDate, "yyyy-MM-dd HH:mm:ss") & _
                        "','YYYY-MM-DD HH24:MI:SS'),'δ�����κ���Ŀ',Null,'�Ѿ�ִ��|1" & vbTab & "�Ѿ�ִ��',Null,Null,'',1" & _
                        ",Null,'" & strMergeStep & "')"
                colSQL.Add strSql, "C" & colSQL.count + 1
                
                '�Ǽ�ʱ���һ�룬��Ϊ��ȡ������ʱȡ��һ���׶ε�ID��
                AddDate = AddDate + 1 / 24 / 60 / 60
                '����
                strSql = "Zl_����·������_Insert(1," & mPP.����·��ID & "," & mrsPhase!ID & _
                        "," & strDate & "," & mlng��ǰ���� & ",'" & _
                        str������ & "',1,'','" & UserInfo.���� & "','" & str����� & "','" & strLastVariation & "',1,Null,Null" & ",Null,1" & ")"
                        
                colSQL.Add strSql, "C" & colSQL.count + 1
            End If
        Next
        mrsPhase.Filter = varFilter
    End If
        
    For k = 0 To vsItem.count - 1
        With vsItem(k)
            For i = 1 To .Rows - 1
                If .Cell(flexcpChecked, i, mcol("ѡ��")) = 1 Or .Cell(flexcpChecked, i, mcol("ѡ��")) = 2 And mlngFun = 0 And .RowData(i) = ִ�з�ʽ.T1ÿ����� Then
                    strBaby = GetBabyIndex(.TextMatrix(i, mcol("Ӥ��")))
                    
                    strAdviceIDs = ""
                    str���˲���IDs = ""
                    strFileIDs = ""
                    strVariation = ""
                    blnAgain = False
                    strFileIDsTmp = ""
                    blnDef = False
                    
                    If .Cell(flexcpChecked, i, mcol("ѡ��")) = 1 Then
                        If Val(.TextMatrix(i, mcol("ҽ������ID"))) <> 0 Then
                            strAgaignTmp = Trim(.TextMatrix(i, mcol("������ĿID")))
                            If InStr(strAgain & vbCrLf, vbCrLf & strAgaignTmp & vbCrLf) = 0 Then
                                For j = 0 To UBound(arrItem)
                                    If Split(arrItem(j), ":")(0) = .TextMatrix(i, mcol("ID")) Then  '·����ĿID
                                        strAdviceIDs = strAdviceIDs & "," & Split(arrItem(j), ":")(1)  'ҽ��ID
                                    ElseIf .TextMatrix(i, mcol("�ظ���Ŀ")) <> "" Then
                                        '������ظ���Ŀ�������Ŀ������ͬ�����ظ�������ͬ��Ŀ��������Ʋ�ͬ����������Ŀָ����ͬҽ��
                                        If Split(arrItem(j), ":")(0) = Split(.TextMatrix(i, mcol("�ظ���Ŀ")), vbCrLf)(0) Then
                                            If .TextMatrix(i, mcol("��Ŀ����")) = Split(.TextMatrix(i, mcol("�ظ���Ŀ")), vbCrLf)(1) Then
                                                blnAgain = True
                                            Else
                                                strAdviceIDs = strAdviceIDs & "," & Split(arrItem(j), ":")(1)  'ҽ��ID
                                            End If
                                        End If
                                    End If
                                Next
                            Else
                                '����̳е���Ŀ��������ظ�����ֻ����һ����Ŀ
                                If .TextMatrix(i, mcol("��Ŀ����")) = colItemName("C" & strAgaignTmp) Then
                                    blnAgain = True
                                Else
                                    For j = 0 To UBound(arrItem)
                                        If Split(arrItem(j), ":")(0) = .TextMatrix(i, mcol("ID")) Then  '·����ĿID
                                            strAdviceIDs = strAdviceIDs & "," & Split(arrItem(j), ":")(1)  'ҽ��ID
                                        ElseIf .TextMatrix(i, mcol("�ظ���Ŀ")) <> "" Then
                                            '������ظ���Ŀ�������Ŀ������ͬ�����ظ�������ͬ��Ŀ��������Ʋ�ͬ����������Ŀָ����ͬҽ��
                                            If Split(arrItem(j), ":")(0) = Split(.TextMatrix(i, mcol("�ظ���Ŀ")), vbCrLf)(0) Then
                                                If .TextMatrix(i, mcol("��Ŀ����")) = Split(.TextMatrix(i, mcol("�ظ���Ŀ")), vbCrLf)(1) Then
                                                    blnAgain = True
                                                Else
                                                    strAdviceIDs = strAdviceIDs & "," & Split(arrItem(j), ":")(1)  'ҽ��ID
                                                End If
                                            End If
                                        End If
                                    Next
                                End If
                                blnDef = True
                            End If
                            If Not blnAgain And Not blnDef Then
                                strAgain = strAgain & vbCrLf & strAgaignTmp
                                colItemName.Add .TextMatrix(i, mcol("��Ŀ����")), "C" & strAgaignTmp
                            End If
                            strAdviceIDs = Mid(strAdviceIDs, 2)
                            
                            '�����ҩ�䷽�޸Ĺ�����д�˱���ԭ���򱣴浽������Ŀ��
                            If .Cell(flexcpData, i, mcol("����ԭ��")) <> "" Then
                                strVariation = .Cell(flexcpData, i, mcol("����ԭ��"))
                            End If
                        End If
                        
                        strEPR = Trim(.TextMatrix(i, mcol("�ļ�ID")))     '�����ж��
                        If strEPR <> "" Then
                           strFiles = Split(strEPR, "|")(0)  '�ɰ�
                           strPara = Split(strEPR, "|")(1)  '�°�
                           strEPR = ""
                           If strFiles <> "" Then
                                arrtmp = Split(strBaby, "|")
                                For l = LBound(arrtmp) To UBound(arrtmp)
                                    strEMRID = "": strFileIDsTmp = ""
                                    For j = 0 To UBound(Split(strFiles, ","))
                                        strFileID = Split(strFiles, ",")(j)
                                         '����ʼ�ղ������ظ��ģ�һ�������ļ�ֻ����һ��
                                        lngBaby = CLng(arrtmp(l) & "")
                                        If InStr(strEPR & ",", "," & lngBaby & "_" & strFileID & ",") = 0 Then
                                            lng����ID = zlDatabase.GetNextId("���Ӳ�����¼")
                                            strEMRID = strEMRID & "," & lng����ID
                                            colDoc.Add lng����ID & ":" & lngBaby & ":" & mEditType("C" & strFileID), "C" & (colDoc.count + 1)
                                            strFileIDsTmp = strFileIDsTmp & "," & strFileID
                                            strEPR = strEPR & "," & lngBaby & "_" & strFileID
                                        End If
                                    Next
                                    str���˲���IDs = str���˲���IDs & "|" & Mid(strEMRID, 2)
                                    strFileIDs = strFileIDs & "|" & Mid(strFileIDsTmp, 2)
                                Next
                                str���˲���IDs = Mid(str���˲���IDs, 2)
                                strFileIDs = Mid(strFileIDs, 2)
                            End If
                            If strPara <> "" And Not gobjEmr Is Nothing Then '�°没��
                                If Not gobjEmr.IsInited Or gobjEmr.IsOffline Then Set gobjEmr = Nothing
                                If Not gobjEmr Is Nothing Then
                                    strParaTmp = "": strFileIDsTmp = ""
                                    For j = 0 To UBound(Split(strPara, ","))
                                        strParaTmp = "<parameter><antetypeid>" & Split(strPara, ",")(j) & "</antetypeid><patient>" & mPati.����ID & "</patient></parameter>"
                                        '��¼�������ֶΣ�ԭ��ID,����ID,����ʱ��,��ʼʱ��,��ֹʱ�䣻
                                        On Error Resume Next
                                        Set rsTmp = gobjEmr.MakeBeforTask(strParaTmp)
                                        Err.Clear: On Error GoTo 0
                                        If rsTmp.State <> adStateClosed Then
                                            If rsTmp.RecordCount = 1 Then
                                                strFileIDsTmp = strFileIDsTmp & "," & rsTmp!����ID
                                            End If
                                        End If
                                    Next
                                    strPara = Mid(strFileIDsTmp, 2)
                                    colNewDoc.Add strPara, "C" & (colNewDoc.count + 1) '��¼���ص�����ID,���������ύʧ��,ɾ�����ɳɹ����°没��
                                End If
                            End If
                            If str���˲���IDs & strPara = "" Then blnAgain = True
                        End If
                    Else
                        strVariation = .Cell(flexcpData, i, mcol("����ԭ��"))
                    End If
                    
                    If Not blnAgain Then
                        lng��� = lng��� + 1
                        strSql = "Zl_����·������_Insert(" & lng��� & "," & mPati.����ID & "," & mPati.��ҳID & ",'" & strBaby & "'," & mPati.����ID & "," & _
                            mPP.����·��ID & "," & mrsPhase!ID & _
                            "," & strDate & "," & mlng��ǰ���� & _
                            ",'" & .TextMatrix(i, mcol("����")) & "'," & .TextMatrix(i, mcol("ID")) & _
                            ",'" & strAdviceIDs & "','" & strFileIDs & "','" & str���˲���IDs & "'" & _
                            ",'" & UserInfo.���� & "',To_Date('" & Format(AddDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),Null,Null,Null,Null,Null,'" & _
                            strVariation & "',Null,Null,'" & strMergeStep & "'," & ZVal(Val(.TextMatrix(i, mcol("�ϲ�·����¼ID")))) & "," & ZVal(Val(.TextMatrix(i, mcol("�׶�ID")))) & ",0," & IIf(mint���� = 0, 1, 2) & ",'" & strPara & "')"
                        colSQL.Add strSql, "C" & colSQL.count + 1
                        blnIsSend = True
                    End If
                End If
            Next
        End With
    Next
    '���û�й�ѡ�κ���Ŀ��������һ���������Ŀ��δ�����κ���Ŀ
    If Not blnIsSend Then
        If mlngFun = 0 Then
            lng��� = lng��� + 1
            strSql = "Zl_����·������_Insert(" & lng��� & "," & mPati.����ID & "," & mPati.��ҳID & ",NULL," & mPati.����ID & "," & _
                    mPP.����·��ID & "," & mrsPhase!ID & _
                    "," & strDate & "," & mlng��ǰ���� & _
                    ",'" & strFirstType & "',Null" & _
                    ",Null,Null,Null,'" & UserInfo.���� & "',To_Date('" & Format(AddDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),'δ�����κ���Ŀ',Null,'�Ѿ�ִ��|1" & vbTab & "�Ѿ�ִ��',Null,Null,'',Null" & _
                    ",Null,'" & strMergeStep & "',NULL,NULL,0," & IIf(mint���� = 0, 1, 2) & ")"
            colSQL.Add strSql, "C" & colSQL.count + 1
        End If
    End If
    
    On Error GoTo errH
    gcnOracle.BeginTrans: blnTrans = True
        If mlngFun = 3 Then
            strSql = "Zl_����·������_Delete(" & mlngִ��ID & ",1)"
            zlDatabase.ExecuteProcedure strSql, Me.Caption
        End If
        
        '1.�Ȳ���ҽ��,��Ϊ����·��ҽ�������
        For i = 0 To UBound(arrSQL)
            zlDatabase.ExecuteProcedure CStr(arrSQL(i)), Me.Caption
        Next
        '2.��������·�����ݣ��Լ������ļ�����
        For i = 1 To colSQL.count
            Call zlDatabase.ExecuteProcedure(colSQL("C" & i), Me.Caption)
        Next
        '3.���������ļ�RTF����
        For i = 1 To colDoc.count
            arrItem = Split(colDoc("C" & i), ":")
            If arrItem(2) = 0 Or arrItem(2) = 1 Then     'ȫ�ı༭��ʽ�Ĳ���
                lng����ID = (arrItem(0))
                Call ReadRTFData(lng����ID, edtEditor)
                Call SaveRTFData(lng����ID, mPati.����ID, mPati.��ҳID, Val(arrItem(1)), edtEditor)
            End If
        Next
    gcnOracle.CommitTrans: blnTrans = False
    Call ZLHIS_CIS_001(mclsMipModule, mPati.����ID, mPati.��ҳID, mPati.����ID, mPati.����ID)
 
    Exit Sub
errH:
    If blnTrans Then
        gcnOracle.RollbackTrans
        '--ɾ���������°没��
        If Not gobjEmr Is Nothing Then
            If Not gobjEmr.IsInited Or gobjEmr.IsOffline Then Set gobjEmr = Nothing
            If Not gobjEmr Is Nothing Then
                For i = 1 To colNewDoc.count
                    strPara = "<parameter><taskid>" & colNewDoc("C" & i) & "</taskid></parameter>"
                    On Error Resume Next
                    Call gobjEmr.DeleteTask(strPara)
                    Err.Clear: On Error GoTo 0
                Next
            End If
        End If
    End If
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function GetShouldStopAdvice(ByVal strIDs As String, ByVal lng���� As Long) As String
'���ܣ���ȡ��ǰӦ��ֹͣ�ĳ���ҽ������һ��ִ���д��ڣ�������ִ���в����ڣ�
'������strIDs=���һ��ִ�еĳ���ҽ��ID
'      lng����=�������ɵ�����
'      ���أ�����ҽ��ID
    Dim rsTmp As ADODB.Recordset, strSql As String
    Dim i As Long
    
    On Error GoTo errH
    strSql = "Select /*+ rule*/ Column_Value As ����ҽ��id" & vbNewLine & _
            "From Table(f_Num2list([1])) " & vbNewLine & _
            "Minus" & vbNewLine & _
            "Select b.����ҽ��id" & vbNewLine & _
            "From ����·��ִ�� A, ����·��ҽ�� B" & vbNewLine & _
            "Where a.·����¼id = [2] And a.�׶�id = [3] And a.���� = [4] And a.Id = b.·��ִ��id"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strIDs, mPP.����·��ID, Val(mrsPhase!ID), lng����)
    For i = 1 To rsTmp.RecordCount
        GetShouldStopAdvice = GetShouldStopAdvice & "," & rsTmp!����ҽ��id
        rsTmp.MoveNext
    Next
    GetShouldStopAdvice = Mid(GetShouldStopAdvice, 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function MakePathAdivceRS() As ADODB.Recordset
    Set MakePathAdivceRS = New ADODB.Recordset
    MakePathAdivceRS.Fields.Append "·����ĿID", adBigInt
    MakePathAdivceRS.Fields.Append "ԭҽ��ID", adBigInt
    
    MakePathAdivceRS.Fields.Append "·����Ŀ����", adVarChar, 50, adFldIsNullable
    MakePathAdivceRS.Fields.Append "ҽ��IDS", adLongVarWChar, 4000, adFldIsNullable
    MakePathAdivceRS.CursorLocation = adUseClient
    MakePathAdivceRS.LockType = adLockOptimistic
    MakePathAdivceRS.CursorType = adOpenStatic
    MakePathAdivceRS.Open
End Function

Private Sub CheckStopAdvice(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByRef strUnStopIDs As String)
'����:
'����:
'strUnStopIDs-δֹͣ��ҽ��ID��һ��ҽ��������ID������Ҫ��ӵ�·������Ŀ
'lng��ǰ�׶�ID-��ǰ�׶�ID
    Dim rsUnStop As ADODB.Recordset
    Dim rsPath As ADODB.Recordset
    Dim rsPathAdvice As ADODB.Recordset
    Dim rsTmp As ADODB.Recordset

    Dim strSql As String
    Dim i As Long, j As Long
    Dim k As Long

    Dim lng����·��Id  As Long
    Dim lng�׶�ID As Long
    Dim lng���� As Long
    Dim lngPos As Long
    Dim strDate As String
    Dim strTag As String
    Dim str���ID As String
    Dim AddDate As Date
    Dim colSQL As New Collection
    Dim blnTrans As Boolean
    Dim strҽ��ID As String
    
    On Error GoTo errH
    strSql = "Select b.Id" & vbNewLine & _
    " From ����ҽ����¼ B" & vbNewLine & _
    " Where b.Id in (select Column_Value As ����ҽ��id From Table(f_Num2list([1]))) And b.ͣ��ʱ�� Is Null"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strUnStopIDs)

    '��ȡδֹͣ�ĳ���ID
    For i = 1 To rsTmp.RecordCount
        strҽ��ID = strҽ��ID & "," & rsTmp!ID
        rsTmp.MoveNext
    Next
    strUnStopIDs = Mid(strҽ��ID, 2)
    If strUnStopIDs = "" Then Exit Sub
    
    '��ȡ��ǰ·����·����¼ID,��ǰ�׶�Id,��ǰ���ڣ�����
    strSql = "Select a.·����¼id, a.��ǰ�׶�id, a.��ǰ����, b.���� " & vbNewLine & _
             "From (Select a.Id As ·����¼id, a.��ǰ�׶�id, a.��ǰ����, Max(b.Id) ִ��id" & vbNewLine & _
             "       From �����ٴ�·�� A, ����·��ִ�� B" & vbNewLine & _
             "       Where a.����id = [1] And a.��ҳid = [2] And a.Id = b.·����¼id And b.�׶�id = a.��ǰ�׶�id And b.���� = a.��ǰ����" & vbNewLine & _
             "       Group By a.Id, a.��ǰ�׶�id, a.��ǰ����) A, ����·��ִ�� B" & vbNewLine & _
             "Where a.ִ��id = b.Id"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng����ID, lng��ҳID)

    If rsTmp.RecordCount = 1 Then
        lng����·��Id = Val(rsTmp!·����¼ID)
        lng�׶�ID = Val(rsTmp!��ǰ�׶�ID)
        strDate = "To_Date('" & Format(rsTmp!����, "yyyy-MM-dd") & "','YYYY-MM-DD')"
        lng���� = Val(rsTmp!��ǰ����)
    Else
        Exit Sub
    End If

    strSql = "select a.ID, a.���ID, b.���, a.������ĿID, b.��������" & vbNewLine & _
            "  from ����ҽ����¼ a, ������ĿĿ¼ b" & vbNewLine & _
            " where a.������ĿID = b.id" & vbNewLine & _
            "   and a.id in (Select Column_Value As ����ҽ��id" & vbNewLine & _
            "                  From Table(f_Num2list([1])))"


    Set rsUnStop = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strUnStopIDs)

    strSql = "Select c.ID, c.���ID,c.������Ŀid,a.id as ·����ĿID,a.���� as ·����Ŀ���� " & vbNewLine & _
            "From �ٴ�·����Ŀ a, �ٴ�·��ҽ�� b, ·��ҽ������ c" & vbNewLine & _
            "where a.id = b.·����Ŀid" & vbNewLine & _
            "   and b.ҽ������id = c.id" & vbNewLine & _
            "   and a.�׶�id = [1]" & vbNewLine & _
            "   and c.��Ч = 0"

    Set rsPath = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng�׶�ID)

    strTag = ""
    Set rsPathAdvice = Nothing
    For i = 1 To rsUnStop.RecordCount
        lngPos = rsUnStop.AbsolutePosition
        If Val(rsUnStop!���id & "") = 0 And Not (rsUnStop!��� & "" = "E" And rsUnStop!�������� & "" = "2") Or InStr(",5,6,", "," & rsUnStop!��� & ",") > 0 Then
            '��һ����ҩ������һ��ʱ��ֻ�������õ�ǰ�У���Ϊ·������Ŀ���ܺ�·������Ŀһ����ҩ
            If InStr(",5,6,", "," & rsUnStop!��� & ",") > 0 Then
                'ҩƷ����ҩ����ƥ�� 65982
                rsUnStop.Filter = "ID=" & rsUnStop!ID
                str���ID = Val(rsUnStop!���id & "")
            Else
                rsUnStop.Filter = "ID=" & rsUnStop!ID & " Or ���ID=" & rsUnStop!ID
                str���ID = Val(rsUnStop!ID & "")
            End If
            'ҩƷ������ҩ;�����÷����巨����Ѫ����;��,���鲻���ɼ���ʽ��������������������������鲻����λ����
            If Not (rsUnStop!��� & "" = "E" And InStr(",2,3,4,6,", "," & rsUnStop!�������� & ",") > 0) _
                And Not (InStr(",G,F,D,", "," & rsUnStop!��� & ",") > 0 And Val(rsUnStop!���id & "") <> 0) Then
                rsPath.Filter = ""
                For j = 1 To rsPath.RecordCount
                    If Nvl(rsPath!������ĿID, 0) = Nvl(rsUnStop!������ĿID, 0) Then '����ҽ��ͳһ����
                        '·������Ŀ
                        If InStr("," & strTag & ",", "," & str���ID & ",") = 0 Then
                            rsUnStop.Filter = "���ID=" & str���ID & " OR ID =" & str���ID
                            If InStr(",5,6,", "," & rsUnStop!��� & ",") > 0 Then
                                strTag = strTag & "," & rsUnStop!���id
                            Else
                                strTag = strTag & "," & rsUnStop!ID
                            End If
                            
                            If rsPathAdvice Is Nothing Then Set rsPathAdvice = MakePathAdivceRS
                            rsPathAdvice.Filter = "·����ĿID = " & rsPath!·����ĿID
                            
                            For k = 1 To rsUnStop.RecordCount
                                rsPathAdvice.Filter = "·����ĿID = " & rsPath!·����ĿID
                                If rsPathAdvice.RecordCount = 0 Then
                                    rsPathAdvice.AddNew
                                    rsPathAdvice!·����ĿID = rsPath!·����ĿID & ""
                                    rsPathAdvice!·����Ŀ���� = rsPath!·����Ŀ���� & ""
                                    rsPathAdvice!ҽ��IDs = rsUnStop!ID & ""
                                Else
                                    rsPathAdvice!ҽ��IDs = rsPathAdvice!ҽ��IDs & "," & rsUnStop!ID
                                End If
                                rsPathAdvice.Update
                                '��δֹͣ�ĳ������Ƴ�
                                strUnStopIDs = Replace("," & strUnStopIDs & ",", "," & rsUnStop!ID & ",", ",")
                                If Left(strUnStopIDs, 1) = "," Then strUnStopIDs = Mid(strUnStopIDs, 2)
                                If Right(strUnStopIDs, 1) = "," Then strUnStopIDs = Mid(strUnStopIDs, 1, Len(strUnStopIDs) - 1)
                                rsUnStop.MoveNext
                            Next
                        End If
                        Exit For
                    End If
                    rsPath.MoveNext
                Next
            End If
        End If
        rsUnStop.Filter = ""
        rsUnStop.AbsolutePosition = lngPos
        rsUnStop.MoveNext
    Next
    
    If rsPathAdvice Is Nothing Then Exit Sub
    rsPathAdvice.Filter = ""
    AddDate = zlDatabase.Currentdate
    For j = 1 To rsPathAdvice.RecordCount
        strSql = "Zl_����·������_Insert(1," & lng����ID & "," & lng��ҳID & ",NULL,0," & lng����·��Id & "," & lng�׶�ID & _
            "," & strDate & "," & lng���� & ",'" & rsPathAdvice!·����Ŀ���� & "'," & rsPathAdvice!·����ĿID & ",'" & rsPathAdvice!ҽ��IDs & "',Null,Null" & _
            ",'" & UserInfo.���� & "'," & "To_Date('" & Format(AddDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & "'',NULL,NULL,NULL,NULL,'',1)"
            
        colSQL.Add strSql, "C" & colSQL.count + 1
        '�Ǽ�ʱ���һ�룬��Ϊ��ȡ������ʱȡ��һ���׶ε�ID��
        AddDate = AddDate + 1 / 24 / 60 / 60
        rsPathAdvice.MoveNext
    Next
  
    gcnOracle.BeginTrans: blnTrans = True
    For i = 1 To colSQL.count
        Call zlDatabase.ExecuteProcedure(colSQL("C" & i), "·������")
    Next
    gcnOracle.CommitTrans: blnTrans = False
    
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
