VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAntiEdit 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7065
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10425
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7065
   ScaleWidth      =   10425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picName 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   1395
      Left            =   0
      ScaleHeight     =   1395
      ScaleWidth      =   5370
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   5370
      Begin VB.TextBox txtӢ�� 
         Height          =   300
         Left            =   615
         MaxLength       =   60
         TabIndex        =   8
         Top             =   960
         Width           =   4545
      End
      Begin VB.TextBox txt���� 
         Height          =   300
         Left            =   615
         MaxLength       =   60
         TabIndex        =   6
         Top             =   540
         Width           =   4545
      End
      Begin VB.TextBox txt���� 
         Height          =   300
         Left            =   615
         MaxLength       =   10
         TabIndex        =   2
         Top             =   120
         Width           =   1635
      End
      Begin VB.TextBox txt��д 
         Height          =   300
         Left            =   3525
         MaxLength       =   10
         TabIndex        =   4
         Top             =   120
         Width           =   1635
      End
      Begin VB.Label lblӢ�� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Ӣ��"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   165
         TabIndex        =   7
         Top             =   1020
         Width           =   360
      End
      Begin VB.Label lbl���� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   165
         TabIndex        =   5
         Top             =   615
         Width           =   360
      End
      Begin VB.Label lbl���� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   165
         TabIndex        =   1
         Top             =   195
         Width           =   360
      End
      Begin VB.Label lbl��д 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��д"
         Height          =   180
         Left            =   3075
         TabIndex        =   3
         Top             =   180
         Width           =   360
      End
   End
   Begin VB.PictureBox picSingle 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   4275
      Left            =   5000
      ScaleHeight     =   4275
      ScaleWidth      =   5370
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1395
      Visible         =   0   'False
      Width           =   5370
      Begin VB.ComboBox cboҩ������ 
         Height          =   300
         Left            =   1335
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   150
         Width           =   1335
      End
      Begin VB.TextBox txtWHONET�� 
         Height          =   300
         Left            =   3900
         MaxLength       =   10
         TabIndex        =   13
         Top             =   150
         Width           =   1260
      End
      Begin VB.TextBox txt˵�� 
         Height          =   300
         Left            =   1335
         MaxLength       =   10
         TabIndex        =   15
         Top             =   570
         Width           =   3825
      End
      Begin VB.Frame frmLine 
         Height          =   15
         Left            =   165
         TabIndex        =   40
         Top             =   0
         Width           =   5040
      End
      Begin VB.Frame fraLine 
         Height          =   15
         Index           =   1
         Left            =   165
         TabIndex        =   39
         Top             =   1005
         Width           =   5055
      End
      Begin VB.TextBox txt�÷����� 
         Height          =   300
         Index           =   0
         Left            =   1170
         MaxLength       =   10
         TabIndex        =   17
         Top             =   1155
         Width           =   3990
      End
      Begin VB.TextBox txtѪҩŨ�� 
         Height          =   300
         Index           =   0
         Left            =   1965
         MaxLength       =   10
         TabIndex        =   19
         Top             =   1560
         Width           =   3195
      End
      Begin VB.TextBox txt��ҩŨ�� 
         Height          =   300
         Index           =   0
         Left            =   1965
         MaxLength       =   10
         TabIndex        =   21
         Top             =   1965
         Width           =   3195
      End
      Begin VB.TextBox txt�÷����� 
         Height          =   300
         Index           =   1
         Left            =   1170
         MaxLength       =   10
         TabIndex        =   23
         Top             =   2400
         Width           =   3990
      End
      Begin VB.TextBox txtѪҩŨ�� 
         Height          =   300
         Index           =   1
         Left            =   1965
         MaxLength       =   10
         TabIndex        =   25
         Top             =   2805
         Width           =   3195
      End
      Begin VB.TextBox txt��ҩŨ�� 
         Height          =   300
         Index           =   1
         Left            =   1965
         MaxLength       =   10
         TabIndex        =   27
         Top             =   3210
         Width           =   3195
      End
      Begin VB.Frame fraLine 
         Height          =   15
         Index           =   0
         Left            =   165
         TabIndex        =   38
         Top             =   3660
         Width           =   5055
      End
      Begin VB.Label lblNote 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�÷�������Ҫ����΢����������鱨���й��ٴ��Ĳο���"
         Height          =   180
         Left            =   555
         TabIndex        =   28
         Top             =   3810
         Width           =   4500
      End
      Begin VB.Image imgNote 
         Height          =   240
         Left            =   180
         Picture         =   "frmAntiEdit.frx":0000
         Top             =   3780
         Width           =   240
      End
      Begin VB.Label lblҩ������ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ĭ��ҩ������"
         Height          =   180
         Left            =   165
         TabIndex        =   10
         Top             =   210
         Width           =   1080
      End
      Begin VB.Label lblWHONET�� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "WHONET��"
         Height          =   180
         Left            =   3135
         TabIndex        =   12
         Top             =   210
         Width           =   720
      End
      Begin VB.Label lbl˵�� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ҩƷ����˵��"
         Height          =   180
         Left            =   165
         TabIndex        =   14
         Top             =   630
         Width           =   1080
      End
      Begin VB.Label lbl�÷����� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�÷�������"
         Height          =   180
         Index           =   0
         Left            =   165
         TabIndex        =   16
         Top             =   1215
         Width           =   900
      End
      Begin VB.Label lblѪҩŨ�� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ѪҩŨ��"
         Height          =   180
         Index           =   0
         Left            =   1170
         TabIndex        =   18
         Top             =   1620
         Width           =   720
      End
      Begin VB.Label lbl��ҩŨ�� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ҩŨ��"
         Height          =   180
         Index           =   0
         Left            =   1170
         TabIndex        =   20
         Top             =   2025
         Width           =   720
      End
      Begin VB.Label lbl�÷����� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�÷�������"
         Height          =   180
         Index           =   1
         Left            =   165
         TabIndex        =   22
         Top             =   2460
         Width           =   900
      End
      Begin VB.Label lblѪҩŨ�� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ѪҩŨ��"
         Height          =   180
         Index           =   1
         Left            =   1170
         TabIndex        =   24
         Top             =   2865
         Width           =   720
      End
      Begin VB.Label lbl��ҩŨ�� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ҩŨ��"
         Height          =   180
         Index           =   1
         Left            =   1170
         TabIndex        =   26
         Top             =   3270
         Width           =   720
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vfgGroup 
      Height          =   2805
      Left            =   165
      TabIndex        =   30
      Top             =   1665
      Visible         =   0   'False
      Width           =   5040
      _cx             =   8890
      _cy             =   4948
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
      BackColorFixed  =   15790320
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
      Rows            =   3
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
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
   Begin VB.PictureBox picGroup 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   2505
      Left            =   165
      ScaleHeight     =   2505
      ScaleWidth      =   5040
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   4530
      Visible         =   0   'False
      Width           =   5040
      Begin VB.TextBox txtFind 
         Height          =   300
         Left            =   540
         TabIndex        =   33
         Top             =   70
         Width           =   1650
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "��"
         Height          =   315
         Left            =   2190
         TabIndex        =   34
         TabStop         =   0   'False
         ToolTipText     =   "���ҷ�����������Ŀ"
         Top             =   63
         Width           =   360
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "�� ���"
         Height          =   350
         Index           =   0
         Left            =   2925
         TabIndex        =   35
         Top             =   45
         Width           =   1080
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "�� ɾ��"
         Height          =   350
         Index           =   1
         Left            =   4005
         TabIndex        =   36
         Top             =   45
         Width           =   1080
      End
      Begin MSComctlLib.ListView lvwGroup 
         Height          =   2040
         Left            =   0
         TabIndex        =   37
         Top             =   450
         Width           =   5025
         _ExtentX        =   8864
         _ExtentY        =   3598
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label lblFind 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����:"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   75
         TabIndex        =   32
         Top             =   130
         Width           =   450
      End
   End
   Begin VB.Label lblGroup 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��ҩ��������Ŀ�����:"
      Height          =   180
      Left            =   165
      TabIndex        =   29
      Top             =   1440
      Width           =   1890
   End
End
Attribute VB_Name = "frmAntiEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngItemID As Long          '��ǰ��ʾ����Ŀid
Private mintGroup As Integer        '��ǰ��ʾ����Ŀid

Private Enum mcol
    ID = 0: ����: ������: ��д
End Enum

Dim objItem As ListItem
Dim lngCount As Long

'--------------------------------------------
'����Ϊ���幫������
'--------------------------------------------
Private Sub setListFormat(Optional blnKeepData As Boolean)
    '���ܣ���ʼ�������б�
    '������ blnKeepData-�Ƿ������ݣ���ֻ���������ø�ʽ
    With Me.vfgGroup
        .Redraw = flexRDNone
        If blnKeepData = False Then
            .Clear
            .Rows = 1: .FixedRows = 1: .Cols = 4: .FixedCols = 0
        End If
        .TextMatrix(0, mcol.ID) = "ID": .TextMatrix(0, mcol.����) = "����"
        .TextMatrix(0, mcol.������) = "������": .TextMatrix(0, mcol.��д) = "��д"
        
        .ColWidth(mcol.ID) = 0:  .ColWidth(mcol.����) = 900
        .ColWidth(mcol.������) = 3000: .ColWidth(mcol.����) = 1000
        For lngCount = 0 To .Cols - 1
            .FixedAlignment(lngCount) = flexAlignCenterCenter
            If .ColWidth(lngCount) = 0 Then .ColHidden(lngCount) = True
        Next
        .Redraw = flexRDDirect
    End With
End Sub

Public Function zlRefresh(lngItemID As Long, intGroup As Integer) As Boolean
    '���ܣ�������Ŀidˢ�µ�ǰ��ʾ����
    Dim rsTemp As New ADODB.Recordset
    
    mlngItemID = lngItemID: mintGroup = intGroup
    
    '�����ǰ��Ŀ����ʾ
    Me.txt����.Text = "": Me.txt����.Text = "": Me.txtӢ��.Text = "": Me.txt��д.Text = ""
    If intGroup = 0 Then
        Me.txtWHONET��.Text = "": Me.txt˵��.Text = ""
        Me.txt�÷�����(0).Text = "": Me.txtѪҩŨ��(0).Text = "": Me.txt��ҩŨ��(0).Text = ""
        Me.txt�÷�����(1).Text = "": Me.txtѪҩŨ��(1).Text = "": Me.txt��ҩŨ��(1).Text = ""
        
        Me.picSingle.Visible = True: Me.vfgGroup.Visible = False: Me.picGroup.Visible = False
    Else
        Me.txtFind.Text = "": Me.lvwGroup.ListItems.Clear: Call setListFormat
        Me.picSingle.Visible = False: Me.vfgGroup.Visible = True: Me.picGroup.Visible = True
    End If
    If lngItemID = 0 Then zlRefresh = True: Exit Function
    
    '��ȡָ����Ŀ����Ϣ
    Err = 0: On Error GoTo ErrHand
    
    If intGroup = 0 Then
        gstrSql = "Select ����, ������, Ӣ����, ����, ˵��, ҩ������, Whonet��, �÷�����1, ѪҩŨ��1, ��ҩŨ��1, �÷�����2, ѪҩŨ��2," & vbNewLine & _
                "       ��ҩŨ��2" & vbNewLine & _
                "From �����ÿ�����" & vbNewLine & _
                "Where ID = [1]"
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, lngItemID)
        With rsTemp
            Me.txt����.MaxLength = .Fields("����").DefinedSize: Me.txt����.MaxLength = .Fields("������").DefinedSize
            Me.txtӢ��.MaxLength = .Fields("Ӣ����").DefinedSize: Me.txt��д.MaxLength = .Fields("����").DefinedSize
            Me.txt˵��.MaxLength = .Fields("˵��").DefinedSize: Me.txtWHONET��.MaxLength = .Fields("WHONET��").DefinedSize
            Me.txt�÷�����(0).MaxLength = .Fields("�÷�����1").DefinedSize: Me.txt�÷�����(1).MaxLength = Me.txt�÷�����(0).MaxLength
            Me.txtѪҩŨ��(0).MaxLength = .Fields("ѪҩŨ��1").DefinedSize: Me.txtѪҩŨ��(1).MaxLength = Me.txtѪҩŨ��(0).MaxLength
            Me.txt��ҩŨ��(0).MaxLength = .Fields("��ҩŨ��1").DefinedSize: Me.txt��ҩŨ��(1).MaxLength = Me.txt��ҩŨ��(0).MaxLength
            If .RecordCount > 0 Then
                Me.txt����.Text = "" & !����: Me.txt����.Text = "" & !������
                Me.txtӢ��.Text = "" & !Ӣ����: Me.txt��д.Text = "" & !����
                Me.txt˵��.Text = "" & !˵��: Me.txtWHONET��.Text = "" & !WHONET��
                If Val("" & !ҩ������) = 3 Then
                    Me.cboҩ������.ListIndex = 2
                ElseIf Val("" & !ҩ������) = 2 Then
                    Me.cboҩ������.ListIndex = 1
                Else
                    Me.cboҩ������.ListIndex = 0
                End If
                Me.txt�÷�����(0).Text = "" & !�÷�����1: Me.txt�÷�����(1).Text = "" & !�÷�����2
                Me.txtѪҩŨ��(0).Text = "" & !ѪҩŨ��1: Me.txtѪҩŨ��(1).Text = "" & !ѪҩŨ��2
                Me.txt��ҩŨ��(0).Text = "" & !��ҩŨ��1: Me.txt��ҩŨ��(1).Text = "" & !��ҩŨ��2
            End If
        End With
    Else
        gstrSql = "Select ����, ����, Ӣ��, ���� From ���鿹������ Where ID = [1]"
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, lngItemID)
        With rsTemp
            Me.txt����.MaxLength = .Fields("����").DefinedSize: Me.txt����.MaxLength = .Fields("����").DefinedSize
            Me.txtӢ��.MaxLength = .Fields("Ӣ��").DefinedSize: Me.txt��д.MaxLength = .Fields("����").DefinedSize
            If .RecordCount > 0 Then
                Me.txt����.Text = "" & !����: Me.txt����.Text = "" & !����
                Me.txtӢ��.Text = "" & !Ӣ��: Me.txt��д.Text = "" & !����
            End If
        End With
        
        gstrSql = "Select I.ID, I.����, I.������, I.���� As ��д" & vbNewLine & _
                "From ���鿹������ҩ L, �����ÿ����� I" & vbNewLine & _
                "Where L.������id = I.ID And L.�����ط���id = [1]"
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, lngItemID)
        Set Me.vfgGroup.DataSource = rsTemp: Call setListFormat(True)
        If Me.vfgGroup.Rows > Me.vfgGroup.FixedRows Then Me.vfgGroup.Row = Me.vfgGroup.FixedRows
    
    End If
    zlRefresh = True: Exit Function

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlRefresh = False: Exit Function
End Function

Public Function zlEditStart(blnAdd As Boolean, lngItemID As Long, intGroup As Integer) As Boolean
    '���ܣ���ʼ��Ŀ�༭
    '������ blnAdd-�Ƿ����ӣ�����Ϊ�޸�
    '       lngItemId-���ӵĲ�����Ŀ������ָ���༭����Ŀ
    Dim rsTemp As New ADODB.Recordset
    
    mintGroup = intGroup
    
    If blnAdd Then
        Err = 0: On Error GoTo ErrHand
        If intGroup = 0 Then
            gstrSql = "Select Nvl(Max(To_Number(����)), 0) As ����, Nvl(Max(Length(����)), 0) As ���� From �����ÿ�����"
        Else
            gstrSql = "Select Nvl(Max(To_Number(����)), 0) As ����, Nvl(Max(Length(����)), 0) As ���� From ���鿹������"
        End If
        
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, "cmd����_Click")
        With rsTemp
'            If .State = adStateOpen Then .Close
'            Call SQLTest(App.ProductName, Me.Caption, gstrSql)
'            Call SQLTest
            If !���� <> 0 And !���� <= Me.txt����.MaxLength Then
                Me.txt����.Text = Format(Val(!����) + 1, String(!����, "0"))
            Else
                Me.txt����.Text = Format(Val(!����) + 1, String(Me.txt����.MaxLength, "0"))
            End If
        End With
        Me.txt����.Text = "": Me.txtӢ��.Text = "": Me.txt��д.Text = ""
        If intGroup = 0 Then
            Me.txtWHONET��.Text = "": Me.txt˵��.Text = ""
            Me.txt�÷�����(0).Text = "": Me.txtѪҩŨ��(0).Text = "": Me.txt��ҩŨ��(0).Text = ""
            Me.txt�÷�����(1).Text = "": Me.txtѪҩŨ��(1).Text = "": Me.txt��ҩŨ��(1).Text = ""
            Me.picSingle.Visible = True: Me.vfgGroup.Visible = False: Me.picGroup.Visible = False
        Else
            Me.txtFind.Text = "": Me.lvwGroup.ListItems.Clear: Call setListFormat
            Me.picSingle.Visible = False: Me.vfgGroup.Visible = True: Me.picGroup.Visible = True
        End If
    End If

    Me.Tag = IIf(blnAdd, "����", "�޸�")
    Me.BackColor = RGB(250, 250, 250)
    Me.picName.BackColor = Me.BackColor: Me.picSingle.BackColor = Me.BackColor: Me.picGroup.BackColor = Me.BackColor
    Me.picName.Enabled = True
    If intGroup = 0 Then
        Me.picSingle.Enabled = True
    Else
        Me.picGroup.Enabled = True
        Call Form_Resize
    End If
    
    Me.txt����.SetFocus
    zlEditStart = True: Exit Function

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlEditStart = False: Exit Function
End Function

Public Sub zlEditCancel()
    '���ܣ��������ڽ��еı༭
    Me.Tag = ""
    Me.BackColor = &H8000000F
    Me.picName.BackColor = Me.BackColor: Me.picSingle.BackColor = Me.BackColor: Me.picGroup.BackColor = Me.BackColor
    Me.picName.Enabled = True
    Me.picSingle.Enabled = True
    Me.picGroup.Enabled = True
    Call Form_Resize
    
'    Call Me.zlRefresh(mlngItemID, mintGroup)
End Sub

Public Function zlEditSave() As Long
    '���ܣ��������ڽ��еı༭,���������ڱ༭��Ŀid,����ʧ�ܷ���0
    Dim lngNewId As Long, strLists As String
    
    'һ�����Լ��
    If Trim(Me.txt����.Text) = "" Then
        MsgBox "��������룡", vbInformation, gstrSysName
        Me.txt����.SetFocus: zlEditSave = 0: Exit Function
    End If
    If Val(Me.txt����.Text) > Val(String(Me.txt����.MaxLength, "9")) Then
        MsgBox "����̫��", vbInformation, gstrSysName
        Me.txt����.SetFocus: zlEditSave = 0: Exit Function
    End If
    If Trim(Me.txt����.Text) = "" Then
        MsgBox "�������������ƣ�", vbInformation, gstrSysName
        Me.txt����.SetFocus: zlEditSave = 0: Exit Function
    End If
    If LenB(StrConv(Trim(Me.txt����.Text), vbFromUnicode)) > Me.txt����.MaxLength Then
        MsgBox "�������Ƴ��������" & Me.txt����.MaxLength & "���ַ���ȳ����֣���", vbInformation, gstrSysName
        Me.txt����.SetFocus: zlEditSave = 0: Exit Function
    End If
    If LenB(StrConv(Trim(Me.txtӢ��.Text), vbFromUnicode)) > Me.txtӢ��.MaxLength Then
        MsgBox "Ӣ�����Ƴ��������" & Me.txtӢ��.MaxLength & "���ַ�����", vbInformation, gstrSysName
        Me.txtӢ��.SetFocus: zlEditSave = 0: Exit Function
    End If
    If LenB(StrConv(Trim(Me.txt��д.Text), vbFromUnicode)) > Me.txt��д.MaxLength Then
        MsgBox "��д���������" & Me.txt��д.MaxLength & "���ַ�����", vbInformation, gstrSysName
        Me.txt��д.SetFocus: zlEditSave = 0: Exit Function
    End If
    
    gstrSql = "'" & Trim(Me.txt����.Text) & "','" & Trim(Me.txt����.Text) & "'"
    gstrSql = gstrSql & ",'" & Trim(Me.txtӢ��.Text) & "','" & Trim(Me.txt��д.Text) & "'"
    If mintGroup = 0 Then
        If Me.cboҩ������.ListIndex = -1 Then Me.cboҩ������.ListIndex = 0
        If LenB(StrConv(Trim(Me.txtWHONET��.Text), vbFromUnicode)) > Me.txtWHONET��.MaxLength Then
            MsgBox "WHONET�볬�������" & Me.txtWHONET��.MaxLength & "���ַ�����", vbInformation, gstrSysName
            Me.txtWHONET��.SetFocus: zlEditSave = 0: Exit Function
        End If
        If LenB(StrConv(Trim(Me.txt˵��.Text), vbFromUnicode)) > Me.txt˵��.MaxLength Then
            MsgBox "����˵�����������" & Me.txt˵��.MaxLength & "���ַ�����", vbInformation, gstrSysName
            Me.txt˵��.SetFocus: zlEditSave = 0: Exit Function
        End If
        For lngCount = 0 To 1
            If LenB(StrConv(Trim(Me.txt�÷�����(lngCount).Text), vbFromUnicode)) > Me.txt�÷�����(lngCount).MaxLength Then
                MsgBox "�÷�����" & "lngCount" & "���������" & Me.txt�÷�����(lngCount).MaxLength & "���ַ�����", vbInformation, gstrSysName
                Me.txt�÷�����(lngCount).SetFocus: zlEditSave = 0: Exit Function
            End If
            If LenB(StrConv(Trim(Me.txtѪҩŨ��(lngCount).Text), vbFromUnicode)) > Me.txtѪҩŨ��(lngCount).MaxLength Then
                MsgBox "ѪҩŨ��" & "lngCount" & "���������" & Me.txtѪҩŨ��(lngCount).MaxLength & "���ַ�����", vbInformation, gstrSysName
                Me.txtѪҩŨ��(lngCount).SetFocus: zlEditSave = 0: Exit Function
            End If
            If LenB(StrConv(Trim(Me.txt��ҩŨ��(lngCount).Text), vbFromUnicode)) > Me.txt��ҩŨ��(lngCount).MaxLength Then
                MsgBox "��ҩŨ��" & "lngCount" & "���������" & Me.txt��ҩŨ��(lngCount).MaxLength & "���ַ�����", vbInformation, gstrSysName
                Me.txt��ҩŨ��(lngCount).SetFocus: zlEditSave = 0: Exit Function
            End If
        Next
    
        gstrSql = gstrSql & ",'" & Trim(Me.txt˵��.Text) & "'," & Me.cboҩ������.ListIndex + 1 & ",'" & Trim(Me.txtWHONET��.Text) & "'"
        gstrSql = gstrSql & ",'" & Trim(Me.txt�÷�����(0).Text) & "','" & Trim(Me.txtѪҩŨ��(0).Text) & "','" & Trim(Me.txt��ҩŨ��(0).Text) & "'"
        gstrSql = gstrSql & ",'" & Trim(Me.txt�÷�����(1).Text) & "','" & Trim(Me.txtѪҩŨ��(1).Text) & "','" & Trim(Me.txt��ҩŨ��(1).Text) & "'"
    
    Else
        strLists = ""
        With Me.vfgGroup
            For lngCount = .FixedRows To .Rows - 1
                strLists = strLists & "," & .TextMatrix(lngCount, mcol.ID)
            Next
        End With
        If strLists <> "" Then strLists = Mid(strLists, 2)
        gstrSql = gstrSql & ",'" & strLists & "'"
    End If
    
    '���ݱ��������֯
    
    lngNewId = mlngItemID
    If mintGroup = 0 Then
        If Me.Tag = "����" Then
            lngNewId = zldatabase.GetNextId("�����ÿ�����")
            gstrSql = "Zl_�����ÿ�����_Edit(1," & lngNewId & "," & gstrSql & ")"
        Else
            gstrSql = "Zl_�����ÿ�����_Edit(2," & lngNewId & "," & gstrSql & ")"
        End If
    Else
        If Me.Tag = "����" Then
            lngNewId = zldatabase.GetNextId("���鿹������")
            gstrSql = "Zl_���鿹������_Edit(1," & lngNewId & "," & gstrSql & ")"
        Else
            gstrSql = "Zl_���鿹������_Edit(2," & lngNewId & "," & gstrSql & ")"
        End If
    End If
    
    Err = 0: On Error GoTo ErrHand
    Call SQLTest(App.ProductName, Me.Caption, gstrSql): gcnOracle.Execute gstrSql, , adCmdStoredProc: Call SQLTest
    
    If Me.Tag = "����" Then mlngItemID = lngNewId
    
    Me.Tag = ""
    Me.BackColor = &H8000000F
    Me.picName.BackColor = Me.BackColor: Me.picSingle.BackColor = Me.BackColor: Me.picGroup.BackColor = Me.BackColor
    Me.picName.Enabled = True
    Me.picSingle.Enabled = True
    Me.picGroup.Enabled = True
    Call Form_Resize
    
    zlEditSave = mlngItemID: Exit Function
    
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlEditSave = 0: Exit Function
End Function

'--------------------------------------------
'����Ϊ����ؼ���Ӧ�¼�
'--------------------------------------------

Private Sub cboҩ������_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cmdEdit_Click(Index As Integer)
    Dim lngCurRow As Long
    With Me.vfgGroup
        Select Case Index
        Case 0         '���
            If Me.lvwGroup.SelectedItem Is Nothing Then Exit Sub
            Set objItem = Me.lvwGroup.SelectedItem
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, mcol.ID) = Mid(objItem.Key, 2)
            .TextMatrix(.Rows - 1, mcol.����) = objItem.Text
            .TextMatrix(.Rows - 1, mcol.������) = objItem.SubItems(Me.lvwGroup.ColumnHeaders("_������").Index - 1)
            .TextMatrix(.Rows - 1, mcol.��д) = objItem.SubItems(Me.lvwGroup.ColumnHeaders("_��д").Index - 1)
            If .Rows > .FixedRows And .Row < .FixedRows Then .Row = .FixedRows
            Me.lvwGroup.ListItems.Remove objItem.Key: Me.lvwGroup.SetFocus
        Case 1          'ɾ��
            If .Row < .FixedRows Then Exit Sub
            Set objItem = Me.lvwGroup.ListItems.Add(, "_" & .TextMatrix(.Row, mcol.ID), .TextMatrix(.Row, mcol.����))
            objItem.SubItems(Me.lvwGroup.ColumnHeaders("_������").Index - 1) = .TextMatrix(.Row, mcol.������)
            objItem.SubItems(Me.lvwGroup.ColumnHeaders("_��д").Index - 1) = .TextMatrix(.Row, mcol.��д)
            objItem.Selected = True
            .RemoveItem .Row
        End Select
        .SetFocus
    End With
End Sub

Private Sub cmdFind_Click()
    Dim rsTemp As New ADODB.Recordset
    Dim strFind As String
    strFind = Trim(UCase(Me.txtFind.Text))
    gstrSql = "Select ID, ����, ������, ����" & vbNewLine & _
            "From �����ÿ�����" & vbNewLine & _
            "Where ���� Like '" & strFind & "%' Or Upper(������) Like '" & gstrMatch & strFind & "%' Or" & vbNewLine & _
            "      Upper(Ӣ����) Like '" & gstrMatch & strFind & "%' Or Upper(����) Like '" & gstrMatch & strFind & "%'"
    Err = 0: On Error GoTo ErrHand
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption)
    With rsTemp
        Me.lvwGroup.ListItems.Clear
        Do While Not .EOF
            Set objItem = Me.lvwGroup.ListItems.Add(, "_" & !ID, !����)
            objItem.SubItems(Me.lvwGroup.ColumnHeaders("_������").Index - 1) = "" & !������
            objItem.SubItems(Me.lvwGroup.ColumnHeaders("_��д").Index - 1) = "" & !����
            .MoveNext
        Loop
    End With
    
    Err = 0: On Error Resume Next
    With Me.vfgGroup
        For lngCount = .FixedRows To .Rows - 1
            Me.lvwGroup.ListItems.Remove "_" & .TextMatrix(lngCount, mcol.ID)
        Next
    End With
    If Me.lvwGroup.ListItems.count = 0 Then
        MsgBox "û��ƥ��Ŀ����أ�", vbInformation, gstrSysName
        Me.txtFind.SetFocus
    Else
        Me.vfgGroup.SetFocus
    End If
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    mlngItemID = 0: mintGroup = 0
    
    Me.picName.BackColor = Me.BackColor
    Me.picSingle.BackColor = Me.BackColor
    Me.picGroup.BackColor = Me.BackColor
    
    Me.picName.Left = 0: Me.picName.Top = 0
    Me.picSingle.Left = 0: Me.picSingle.Top = Me.picName.Height
    
    With Me.cboҩ������
        .AddItem "MIC": .AddItem "DISK": .AddItem "K-B"
    End With
    With Me.lvwGroup.ColumnHeaders
        .Clear
        .Add , "_����", "����", 900
        .Add , "_������", "������", 3000
        .Add , "_��д", "��д", 1000
    End With
    With Me.lvwGroup
        .SortKey = .ColumnHeaders("_����").Index - 1
        .SortOrder = lvwAscending
    End With
    Me.vfgGroup.ZOrder 0

End Sub

Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    Me.picSingle.Height = Me.ScaleHeight - Me.picSingle.Top
    Me.picGroup.Top = Me.ScaleHeight - Me.picGroup.Height - 105
    If Me.Tag <> "" Then
        Me.vfgGroup.Height = Me.picGroup.Top - Me.vfgGroup.Top
    Else
        Me.vfgGroup.Height = Me.ScaleHeight - Me.vfgGroup.Top - 105
    End If
End Sub

Private Sub lvwGroup_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With Me.lvwGroup
        If .SortKey = ColumnHeader.Index - 1 Then
            .SortOrder = IIf(.SortOrder = lvwAscending, lvwDescending, lvwAscending)
        Else
            .SortKey = ColumnHeader.Index - 1
            .SortOrder = lvwAscending
        End If
    End With
End Sub

Private Sub lvwGroup_DblClick()
    Call cmdEdit_Click(0)
End Sub

Private Sub picGroup_Resize()
    Err = 0: On Error Resume Next
    Me.lvwGroup.Height = Me.picGroup.ScaleHeight - Me.lvwGroup.Top
End Sub

Private Sub txtFind_GotFocus()
    Me.txtFind.SelStart = 0: Me.txtFind.SelLength = 1000
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call cmdFind_Click
End Sub

Private Sub txtWHONET��_GotFocus()
    Me.txtWHONET��.SelStart = 0: Me.txtWHONET��.SelLength = 1000
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txtWHONET��_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(Trim(GCST_INVALIDCHAR), Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt����_GotFocus()
    Me.txt����.SelStart = 0: Me.txt����.SelLength = 1000
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt��ҩŨ��_GotFocus(Index As Integer)
    Me.txt��ҩŨ��(Index).SelStart = 0: Me.txt��ҩŨ��(Index).SelLength = 1000
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txt��ҩŨ��_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(Trim(GCST_INVALIDCHAR), Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt˵��_GotFocus()
    Me.txt˵��.SelStart = 0: Me.txt˵��.SelLength = 1000
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt˵��_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(GCST_INVALIDCHAR, Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt��д_GotFocus()
    Me.txt��д.SelStart = 0: Me.txt��д.SelLength = 1000
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txt��д_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(Trim(GCST_INVALIDCHAR), Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txtѪҩŨ��_GotFocus(Index As Integer)
    Me.txtѪҩŨ��(Index).SelStart = 0: Me.txtѪҩŨ��(Index).SelLength = 1000
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txtѪҩŨ��_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(Trim(GCST_INVALIDCHAR), Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txtӢ��_GotFocus()
    Me.txtӢ��.SelStart = 0: Me.txtӢ��.SelLength = 1000
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txtӢ��_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(Trim(GCST_INVALIDCHAR), Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt�÷�����_GotFocus(Index As Integer)
    Me.txt�÷�����(Index).SelStart = 0: Me.txt�÷�����(Index).SelLength = 1000
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txt�÷�����_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(Trim(GCST_INVALIDCHAR), Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt����_GotFocus()
    Me.txt����.SelStart = 0: Me.txt����.SelLength = 1000
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(GCST_INVALIDCHAR, Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub vfgGroup_DblClick()
    If Me.vfgGroup.MouseRow < Me.vfgGroup.FixedRows Then Exit Sub
    If Me.Tag = "" Then Exit Sub
    Call cmdEdit_Click(1)
End Sub

