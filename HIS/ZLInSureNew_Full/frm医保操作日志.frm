VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmҽ��������־ 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ҽ��������־�鿴����"
   ClientHeight    =   9525
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14340
   Icon            =   "frmҽ��������־.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9525
   ScaleWidth      =   14340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmdCancle 
      Caption         =   "�ر�(&C)"
      Height          =   345
      Left            =   12915
      TabIndex        =   2
      Top             =   8970
      Width           =   1200
   End
   Begin VB.TextBox txtDetail 
      Height          =   8520
      Left            =   6375
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   30
      Width           =   7935
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfDetail 
      Height          =   8520
      Left            =   0
      TabIndex        =   0
      ToolTipText     =   "Shift+Deleteɾ����ǰ��"
      Top             =   30
      Width           =   6330
      _cx             =   11165
      _cy             =   15028
      Appearance      =   2
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
      BackColorSel    =   16772055
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   9
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmҽ��������־.frx":000C
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   1
      AutoSearch      =   1
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   1
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
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      X1              =   15
      X2              =   15520
      Y1              =   8775
      Y2              =   8775
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      X1              =   0
      X2              =   15520
      Y1              =   8745
      Y2              =   8745
   End
End
Attribute VB_Name = "frmҽ��������־"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrģ��                As String
Private mstr����                As String
Private mstr����1               As String
Private mstr����2               As String
Private mstr����3               As String
Private mstr����4               As String
Private mrsDetail               As ADODB.Recordset
Dim strSql                      As String

Const strDetail = "Select Decode(����, 0, '����', 1, '�޸�', 2, 'ɾ��', 3, '״̬') ����, Decode(��־����, 1, '˵��', 2, '��־') As ��־����, ��־����, ����, �û�, ����վ,������Դ,��־����" & vbCrLf & _
                  "From ҽ��������־" & vbCrLf & _
                  "Where ģ�� = [1] And ���� = [2] And ����1 = [3] And ����2 = [4] And ����3 = [5] And ����4 = [6]" & vbCrLf & _
                  "Order By ����"

Public Property Let strģ��(ByVal vstrģ�� As String)
    mstrģ�� = vstrģ��
End Property

Public Property Let str����(ByVal vstr���� As String)
    mstr���� = vstr����
End Property
 
Public Property Let str����1(ByVal vstr����1 As String)
    mstr����1 = vstr����1
End Property
 
Public Property Let str����2(ByVal vstr����2 As String)
    mstr����2 = vstr����2
End Property
 
Public Property Let str����3(ByVal vstr����3 As String)
    mstr����3 = vstr����3
End Property

Public Property Let str����4(ByVal vstr����4 As String)
    mstr����4 = vstr����4
End Property

Private Sub cmdCancle_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Call mDataload
End Sub

Private Sub mDataload()
On Error GoTo ErrH
    strSql = strDetail
    If mstr����2 = "" Then strSql = Replace(strSql, " And ����2 = [4]", "")
    If mstr����3 = "" Then strSql = Replace(strSql, " And ����3 = [5]", "")
    If mstr����4 = "" Then strSql = Replace(strSql, " And ����4 = [6]", "")
    Set mrsDetail = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mstrģ��, mstr����, mstr����1, mstr����2, mstr����3, mstr����4)
    Set vsfDetail.DataSource = mrsDetail
    If vsfDetail.Rows > 1 Then vsfDetail.Row = 1
    Call vsfDetail_RowColChange
    Exit Sub
ErrH:
    Err.Clear
    Exit Sub
End Sub

Private Sub vsfDetail_BeforeEdit(ByVal Row As Long, ByVal COL As Long, Cancel As Boolean)
    On Error GoTo ErrH
    Cancel = True
    Exit Sub
ErrH:
    Err.Clear
End Sub
 
Private Sub vsfDetail_CellChanged(ByVal Row As Long, ByVal COL As Long)
    Call vsfDetail_RowColChange
End Sub

Private Sub vsfDetail_Click()
    Call vsfDetail_RowColChange
End Sub

Private Sub vsfDetail_RowColChange()
On Error GoTo ErrH
    If vsfDetail.Row < 1 Or vsfDetail.COL < 1 Then
        txtDetail.Text = ""
    Else
        txtDetail.Text = vsfDetail.TextMatrix(vsfDetail.Row, vsfDetail.ColIndex("��־����"))
    End If
    Exit Sub
ErrH:
    Err.Clear
    Exit Sub
End Sub
