VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frm���� 
   BorderStyle     =   0  'None
   ClientHeight    =   735
   ClientLeft      =   6300
   ClientTop       =   0
   ClientWidth     =   5820
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   735
   ScaleWidth      =   5820
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txt��ϸ 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   3870
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   75
      Width           =   1965
   End
   Begin VB.TextBox txt��� 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   2835
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   75
      Width           =   750
   End
   Begin VB.TextBox txt�ܷ��� 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   75
      Width           =   750
   End
   Begin VB.TextBox txt��Ա��� 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   375
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   75
      Width           =   735
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfDetail 
      Height          =   285
      Left            =   30
      TabIndex        =   8
      Top             =   420
      Width           =   5745
      _cx             =   10134
      _cy             =   503
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
      Rows            =   1
      Cols            =   5
      FixedRows       =   0
      FixedCols       =   1
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frm����.frx":0000
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
      ExplorerBar     =   1
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
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.Line Line6 
      BorderColor     =   &H0080FFFF&
      X1              =   -120
      X2              =   8425
      Y1              =   375
      Y2              =   375
   End
   Begin VB.Line Line5 
      BorderColor     =   &H000000FF&
      X1              =   -120
      X2              =   8425
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "��"
      Height          =   180
      Left            =   3690
      TabIndex        =   7
      Top             =   75
      Width           =   180
   End
   Begin VB.Line Line4 
      DrawMode        =   1  'Blackness
      X1              =   3885
      X2              =   5855
      Y1              =   300
      Y2              =   300
   End
   Begin VB.Line Line3 
      DrawMode        =   1  'Blackness
      X1              =   2835
      X2              =   3645
      Y1              =   300
      Y2              =   300
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "���"
      Height          =   180
      Left            =   2415
      TabIndex        =   5
      Top             =   75
      Width           =   360
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "����"
      Height          =   180
      Left            =   1170
      TabIndex        =   3
      Top             =   75
      Width           =   360
   End
   Begin VB.Line Line2 
      DrawMode        =   1  'Blackness
      X1              =   1560
      X2              =   2370
      Y1              =   300
      Y2              =   300
   End
   Begin VB.Line Line1 
      DrawMode        =   1  'Blackness
      X1              =   375
      X2              =   1150
      Y1              =   300
      Y2              =   300
   End
   Begin VB.Label lab��Ա��� 
      AutoSize        =   -1  'True
      Caption         =   "��Ա"
      Height          =   180
      Left            =   0
      TabIndex        =   0
      Top             =   75
      Width           =   360
   End
End
Attribute VB_Name = "frm����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngPatiID          As Long
Private mvarRecId           As Variant
Private mvarKeyId           As Variant
Private mstrReserve         As String
Private mintRecord          As Long

Const col������ = &HFF&
Const col��ͨ�� = vbBlack
Const col���Բ� = &HFF0000
Const col���ֲ� = &HFF00FF

Private Type typ_������Ϣ
    str����                 As String
    str���                 As String
    str����                 As String
    str˵��                 As String
    color                   As Long
End Type
Private var����()           As typ_������Ϣ

Const con���ݿɱ�����       As Double = 8000
Dim rsTmp                   As ADODB.Recordset

Public Property Let PatiID(ByVal vNewValue As Long)
    mlngPatiID = vNewValue
End Property

Public Property Let RecId(ByVal vNewValue As Variant)
    mvarRecId = vNewValue
End Property

Public Property Let KeyId(ByVal vNewValue As Variant)
    mvarKeyId = vNewValue
End Property

Public Property Let Reserve(ByVal vNewValue As String)
    mstrReserve = vNewValue
End Property

Public Sub RefreshData()
    Dim rtn                 As Long
    Dim rsSum               As ADODB.Recordset
    Dim dbl�����ܷ���       As Double
    Dim dblסԺ�ܷ���       As Double
    
    DoEvents
    Me.Show
    rtn = SetWindowPos(Me.hWnd, -1, CurrentX, CurrentY, 0, 0, 3)
    '��ȡ��Ա���
    gstrSql = "select A.��ְ,B.���� from �����ʻ� A ,������Ⱥ B where A.��ְ=B.��� AND A.����=B.���� And A.����ID=[1]"
    Set rsTmp = gDatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngPatiID)
    If ChkRsState(rsTmp) Then
        txt��Ա���.Text = ""
        txt�ܷ���.Text = ""
        txt��ϸ.Text = ""
        txt���.Text = ""
        txt�ܷ���.Visible = False
        txt��ϸ.Visible = False
        txt���.Visible = False
        Label2.Visible = False
        Label3.Visible = False
        Label1.Visible = False
        Line2.Visible = False
        Line3.Visible = False
        Line4.Visible = False
        Me.Hide
        DoEvents
    Else
        txt��Ա���.Text = rsTmp!����
        txt�ܷ���.Visible = (rsTmp!��ְ = "3" Or rsTmp!��ְ = "5")
        txt��ϸ.Visible = (rsTmp!��ְ = "3" Or rsTmp!��ְ = "5")
        txt���.Visible = (rsTmp!��ְ = "3" Or rsTmp!��ְ = "5")
        Label2.Visible = (rsTmp!��ְ = "3" Or rsTmp!��ְ = "5")
        Label3.Visible = (rsTmp!��ְ = "3" Or rsTmp!��ְ = "5")
        Label1.Visible = (rsTmp!��ְ = "3" Or rsTmp!��ְ = "5")
        Line2.Visible = (rsTmp!��ְ = "3" Or rsTmp!��ְ = "5")
        Line3.Visible = (rsTmp!��ְ = "3" Or rsTmp!��ְ = "5")
        Line4.Visible = (rsTmp!��ְ = "3" Or rsTmp!��ְ = "5")
        If Not (rsTmp!��ְ = "3" Or rsTmp!��ְ = "5") Then
            txt��Ա���.Text = rsTmp!����
        Else
            '��ȡ�����ܷ���
            gstrSql = "select nvl(sum(�ۼ�ͳ�ﱨ��), 0) as ���" & vbCrLf & _
                      "  From ���ս����¼" & vbCrLf & _
                      " Where ���� = [1]" & vbCrLf & _
                      "   And ����ID in" & vbCrLf & _
                      "       (Select ����ID" & vbCrLf & _
                      "          From ҽ�����˹�����" & vbCrLf & _
                      "         where ҽ���� in" & vbCrLf & _
                      "               (Select ҽ���� from ҽ�����˹����� where ����ID = [2]))"
            Set rsSum = gDatabase.OpenSQLRecord(gstrSql, Me.Caption, 1, mlngPatiID)
            dbl�����ܷ��� = rsSum!���
            '��ȡסԺ�ܷ���
            Set rsSum = gDatabase.OpenSQLRecord(gstrSql, Me.Caption, 2, mlngPatiID)
            dblסԺ�ܷ��� = rsSum!���
            '�ܷ���
            txt�ܷ���.Text = Format(dbl�����ܷ��� + dblסԺ�ܷ���, "0.00")
            txt��ϸ.Text = "  �ţ�" & Format(dbl�����ܷ���, "0") & ";ס��" & Format(dblסԺ�ܷ���, "0")
            txt���.Text = Format(con���ݿɱ����� - dbl�����ܷ��� - dblסԺ�ܷ���, "0.00")
        End If
    End If
    '��ȡ������Ϣ
'    cmb����.Clear
    gstrSql = "SELECT B.����,DECODE(B.���,1,'���Բ�',2,'���ֲ�',3,'������','��ͨ��') AS ���,B.����,C.˵��  FROM ����_�ز���Ա A,���ղ��� B,��������Ŀ¼ C" & vbCrLf & _
              "WHERE A.ȡ���� is Null And A.����ID=B.ID AND A.���� = B.���� AND B.���� = C.����(+) AND A.ҽ����=(Select ҽ���� from ҽ�����˹����� where ����ID = [1])"
    Set rsTmp = gDatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngPatiID)
    Set vsfDetail.DataSource = rsTmp
    If ChkRsState(rsTmp) Then
        Me.Height = 370
    Else
        vsfDetail.Height = rsTmp.RecordCount * 265 + 20
        Me.Height = vsfDetail.Top + vsfDetail.Height + 10
        
'        ReDim var����(rsTmp.RecordCount - 1) As typ_������Ϣ
'        Do While Not rsTmp.EOF
'            var����(rsTmp.Bookmark - 1).color = Decode(rsTmp!���, "���Բ�", col���Բ�, "���ֲ�", col���ֲ�, "������", col������, col��ͨ��)
'            var����(rsTmp.Bookmark - 1).str���� = "" & rsTmp!����
'            var����(rsTmp.Bookmark - 1).str��� = "" & rsTmp!���
'            var����(rsTmp.Bookmark - 1).str���� = "" & rsTmp!����
'            var����(rsTmp.Bookmark - 1).str˵�� = "" & rsTmp!˵��
'            cmb����.AddItem var����(rsTmp.Bookmark - 1).str����
'            cmb����.ItemData((rsTmp.Bookmark - 1)) = rsTmp.Bookmark - 1
'            rsTmp.MoveNext
'        Loop
'        cmb����.ListIndex = 0
'        cmb����.Enabled = rsTmp.RecordCount > 1
    End If
End Sub

Private Sub Form_Resize()
    Me.Top = 0
    Me.Left = 6300
End Sub

Private Sub cmb����_Click()
'    txt���.ForeColor = var����(cmb����.ListIndex).color
'    txt���.Text = var����(cmb����.ListIndex).str���
'    txt����.Text = var����(cmb����.ListIndex).str����
'    txt˵��.Text = var����(cmb����.ListIndex).str˵��
End Sub

