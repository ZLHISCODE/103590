VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmPurchaseVerifySelect 
   Caption         =   "������˲�ѯ"
   ClientHeight    =   6150
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8910
   Icon            =   "frmPurchaseVerifySelect.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6150
   ScaleWidth      =   8910
   StartUpPosition =   1  '����������
   Begin MSComctlLib.ImageList imgPicture 
      Left            =   4560
      Top             =   5520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPurchaseVerifySelect.frx":6852
            Key             =   "old"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picSplit 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   5895
      Left            =   3840
      MousePointer    =   9  'Size W E
      ScaleHeight     =   5895
      ScaleWidth      =   15
      TabIndex        =   20
      Top             =   120
      Width           =   10
   End
   Begin TabDlg.SSTab sstInfo 
      Height          =   4935
      Left            =   4200
      TabIndex        =   9
      Top             =   240
      Width           =   6525
      _ExtentX        =   11509
      _ExtentY        =   8705
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "������Ϣ"
      TabPicture(0)   =   "frmPurchaseVerifySelect.frx":D0B4
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "vsfAll"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "chkALLVisible1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "��ϸ��Ϣ"
      TabPicture(1)   =   "frmPurchaseVerifySelect.frx":D0D0
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "chkALLVisible2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "optMedi"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "optFloor"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "vsfList"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "lblGroup"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).ControlCount=   5
      Begin VB.CheckBox chkALLVisible2 
         Caption         =   "��ʾ����������˵���"
         Height          =   180
         Left            =   -74880
         TabIndex        =   19
         Top             =   510
         Width           =   2175
      End
      Begin VB.OptionButton optMedi 
         Caption         =   "���ķ���"
         Height          =   180
         Left            =   -69600
         TabIndex        =   18
         Top             =   510
         Width           =   1095
      End
      Begin VB.OptionButton optFloor 
         Caption         =   "���ݷ���"
         Height          =   180
         Left            =   -70800
         TabIndex        =   17
         Top             =   480
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.CheckBox chkALLVisible1 
         Caption         =   "��ʾ����������˵���"
         Height          =   180
         Left            =   120
         TabIndex        =   13
         Top             =   480
         Width           =   2175
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfAll 
         Height          =   1845
         Left            =   120
         TabIndex        =   10
         ToolTipText     =   "��ɫ�����ʾ������˵�����ԭʼ���ݻ�������һ������������"
         Top             =   1080
         Width           =   3255
         _cx             =   5741
         _cy             =   3254
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
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   7
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmPurchaseVerifySelect.frx":D0EC
         ScrollTrack     =   0   'False
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
         VirtualData     =   0   'False
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
      Begin VSFlex8Ctl.VSFlexGrid vsfList 
         Height          =   1845
         Left            =   -74880
         TabIndex        =   12
         ToolTipText     =   "��ɫ�����ʾ������˵�����ԭʼ���ݻ�������һ������������"
         Top             =   1080
         Width           =   3255
         _cx             =   5741
         _cy             =   3254
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
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   15
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmPurchaseVerifySelect.frx":D1E1
         ScrollTrack     =   0   'False
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
         VirtualData     =   0   'False
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
      Begin VB.Label lblGroup 
         AutoSize        =   -1  'True
         Caption         =   "���鷽ʽ"
         Height          =   180
         Left            =   -71640
         TabIndex        =   16
         Top             =   510
         Width           =   720
      End
   End
   Begin VB.PictureBox picLeft 
      BorderStyle     =   0  'None
      Height          =   4695
      Left            =   120
      ScaleHeight     =   4695
      ScaleWidth      =   3375
      TabIndex        =   0
      Top             =   120
      Width           =   3375
      Begin VB.ComboBox cbo�ⷿ 
         Height          =   300
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   60
         Width           =   2535
      End
      Begin VB.TextBox txtNO 
         Height          =   300
         Left            =   840
         TabIndex        =   14
         Top             =   1780
         Width           =   2535
      End
      Begin VB.PictureBox picDate 
         BorderStyle     =   0  'None
         Height          =   800
         Left            =   0
         ScaleHeight     =   795
         ScaleWidth      =   3735
         TabIndex        =   4
         Top             =   800
         Width           =   3735
         Begin MSComCtl2.DTPicker dtpEndDate 
            Height          =   300
            Left            =   840
            TabIndex        =   5
            Top             =   540
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   529
            _Version        =   393216
            Format          =   114491392
            CurrentDate     =   41775
         End
         Begin MSComCtl2.DTPicker dtpBeginDate 
            Height          =   300
            Left            =   840
            TabIndex        =   6
            Top             =   120
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   529
            _Version        =   393216
            Format          =   114491392
            CurrentDate     =   41775
         End
         Begin VB.Label lblBeginDate 
            AutoSize        =   -1  'True
            Caption         =   "��ʼ����"
            Height          =   180
            Left            =   0
            TabIndex        =   8
            Top             =   180
            Width           =   720
         End
         Begin VB.Label lblEndDate 
            AutoSize        =   -1  'True
            Caption         =   "��������"
            Height          =   180
            Left            =   0
            TabIndex        =   7
            Top             =   600
            Width           =   720
         End
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "����"
         Height          =   300
         Left            =   2835
         TabIndex        =   3
         Top             =   500
         Width           =   510
      End
      Begin VB.ComboBox cboDate 
         Height          =   300
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   500
         Width           =   2015
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfLeft 
         Height          =   2205
         Left            =   0
         TabIndex        =   11
         Top             =   2520
         Width           =   3375
         _cx             =   5953
         _cy             =   3889
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
         ForeColor       =   0
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   12632256
         ForeColorSel    =   0
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmPurchaseVerifySelect.frx":D400
         ScrollTrack     =   0   'False
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
         VirtualData     =   0   'False
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
      Begin VB.Label lbl�ⷿ 
         AutoSize        =   -1  'True
         Caption         =   "��    ��"
         Height          =   180
         Left            =   0
         TabIndex        =   21
         Top             =   120
         Width           =   720
      End
      Begin VB.Label lblNO 
         AutoSize        =   -1  'True
         Caption         =   "NO"
         Height          =   180
         Left            =   0
         TabIndex        =   15
         Top             =   1840
         Width           =   180
      End
      Begin VB.Label lblDate 
         AutoSize        =   -1  'True
         Caption         =   "��    ��"
         Height          =   180
         Left            =   0
         TabIndex        =   1
         Top             =   560
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmPurchaseVerifySelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mrsData As New ADODB.Recordset  '���ݼ�
Private mrsCloneDta As New ADODB.Recordset '��¡���ݼ�
Private mstr��ǰ�ⷿ As Long  '�������ĵ�ǰ�ⷿ
Private mStr�ⷿ As String  '�������Ŀⷿ����
Private mlng�ⷿid As Long '��ǰѡ�пⷿ
Private mintUnit As Integer                 '��ʾ��λ:0-ɢװ��λ,1-��װ��λ
Private mdatBeginDate As Date    '��ʼ��ѯʱ��
Private mdatEndDate As Date    '������ѯʱ��

'�Ӳ�������ȡҩƷ�۸����������С��λ��
Private mFMT As g_FmtString

Private Sub SetControlLocation()
    '���ÿؼ�λ��
    On Error Resume Next
    
    picLeft.Move 50, 50, txtNO.Left + txtNO.Width, Me.ScaleHeight - 50
    cmdFind.Move cboDate.Left + cboDate.Width, cboDate.Top
    picDate.Left = 0
    LblNo.Move lblDate.Left, txtNO.Top + 60
    vsfLeft.Move 0, txtNO.Top + txtNO.Height + 100, picLeft.Width, picLeft.ScaleHeight - (txtNO.Top + txtNO.Height + 150)
    picSplit.Move picLeft.Left + picLeft.Width, 0, 10, Me.ScaleHeight
    sstInfo.Move picLeft.Left + picLeft.Width, 50, Me.ScaleWidth - picSplit.Left + 30, Me.ScaleHeight - 50
    chkALLVisible1.Move 100, 480
    chkALLVisible2.Move 100, chkALLVisible1.Top
    lblGroup.Top = chkALLVisible1.Top
    optFloor.Top = chkALLVisible1.Top
    optMedi.Top = chkALLVisible1.Top
    vsfAll.Move 100, chkALLVisible1.Top + chkALLVisible1.Height + 50, sstInfo.Width - 100, sstInfo.Height - (chkALLVisible1.Top + chkALLVisible1.Height + 50)
    VSFList.Move 100, chkALLVisible1.Top + chkALLVisible1.Height + 50, sstInfo.Width - 100, sstInfo.Height - (chkALLVisible1.Top + chkALLVisible1.Height + 50)
End Sub

Private Sub cboDate_Click()
    With cboDate
        If .Text = "�Զ���" Then
            picDate.Visible = True
            txtNO.Top = picDate.Top + picDate.Height + 120
            LblNo.Top = txtNO.Top + 60
            vsfLeft.Top = txtNO.Top + txtNO.Height + 100
            vsfLeft.Height = picLeft.ScaleHeight - (txtNO.Top + txtNO.Height + 100)
        Else
            picDate.Visible = False
            txtNO.Top = picDate.Top + 100
            LblNo.Top = txtNO.Top + 60
            vsfLeft.Top = txtNO.Top + txtNO.Height + 100
            vsfLeft.Height = picLeft.ScaleHeight - (txtNO.Top + txtNO.Height + 100)
        End If
        
        Select Case .Text
            Case "һ������"
                mdatBeginDate = CDate(Format(DateAdd("M", -1, Date), "yyyy-mm-dd") & " 00:00:00")
                mdatEndDate = Sys.Currentdate
            Case "��������"
                mdatBeginDate = CDate(Format(DateAdd("M", -3, Date), "yyyy-mm-dd") & " 00:00:00")
                mdatEndDate = Sys.Currentdate
            Case "������"
                mdatBeginDate = CDate(Format(DateAdd("M", -6, Date), "yyyy-mm-dd") & " 00:00:00")
                mdatEndDate = Sys.Currentdate
            Case "һ����"
                mdatBeginDate = CDate(Format(DateAdd("yyyy", -1, Date), "yyyy-mm-dd") & " 00:00:00")
                mdatEndDate = Sys.Currentdate
            Case "�Զ���"
                mdatBeginDate = CDate(Format(dtpBeginDate, "yyyy-mm-dd") & " 00:00:00")
                mdatEndDate = CDate(Format(dtpEndDate, "yyyy-mm-dd") & " 23:59:59")
        End Select
    End With
End Sub

Private Sub cbo�ⷿ_Click()
    mlng�ⷿid = cbo�ⷿ.ItemData(cbo�ⷿ.ListIndex)
    If cbo�ⷿ.Text = "���пⷿ" Then
        vsfLeft.ColHidden(vsfLeft.ColIndex("�ⷿ")) = False
    Else
        vsfLeft.ColHidden(vsfLeft.ColIndex("�ⷿ")) = True
    End If
End Sub

Private Sub chkALLVisible1_Click()
    If vsfAll.Rows = 1 Then Exit Sub
    chkALLVisible2.Value = chkALLVisible1.Value
    Call SetVsfDta(1)
    Call SetDetailsData
End Sub

Private Sub chkALLVisible2_Click()
    chkALLVisible1.Value = chkALLVisible2.Value
    If vsfAll.Rows = 1 Then Exit Sub
    Call SetVsfDta(1)
    Call SetDetailsData
End Sub

Private Sub cmdFind_Click()
    '��ȡ���ݴ���
    Dim datBeginDate As Date
    Dim datEndDate As Date
    Dim rsTemp As ADODB.Recordset
    Dim lngRow As Long
    
    vsfAll.Rows = 1
    VSFList.Rows = 1
    If cboDate.Text = "�Զ���" Then
        mdatBeginDate = CDate(Format(dtpBeginDate, "yyyy-mm-dd") & " 00:00:00")
        mdatEndDate = CDate(Format(dtpEndDate, "yyyy-mm-dd") & " 23:59:59")
    End If
    If ActiveControl Is cmdFind Then
        txtNO.Text = ""
        If cbo�ⷿ.Text = "���пⷿ" Then
            gstrSQL = ""
        Else
            gstrSQL = "  And A.�ⷿid=[3]"
        End If
        gstrSQL = "Select b.����, a.ԭʼno, a.�ϴ�no, a.����no As NO, a.�����, a.�������" & vbNewLine & _
                "From ҩƷ������� A, ���ű� B" & vbNewLine & _
                "Where a.�ⷿid = b.Id And a.���� = 15 " & gstrSQL & " And a.������� Between [1] And [2]" & vbNewLine & _
                "Order By a.������� Desc"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "������˲�ѯ", mdatBeginDate, mdatEndDate, mlng�ⷿid)
    Else
        If cbo�ⷿ.Text = "���пⷿ" Then
            gstrSQL = ""
        Else
            gstrSQL = " And A.�ⷿid=[2]"
        End If
        gstrSQL = "Select b.����, a.ԭʼno, a.�ϴ�no, a.����no As NO, a.�����, a.�������" & vbNewLine & _
                "From ҩƷ������� A, ���ű� B" & vbNewLine & _
                "Where a.�ⷿid = b.Id And ���� = 15" & gstrSQL & " And ����no = [1]" & vbNewLine & _
                "Order By ������� Desc"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "������˲�ѯ", txtNO.Text, mlng�ⷿid)
    End If
    
    vsfLeft.Rows = 1
    If rsTemp.RecordCount > 0 Then
        rsTemp.Sort = " no asc"
        With vsfLeft
            .Rows = rsTemp.RecordCount + 1
            For lngRow = 1 To rsTemp.RecordCount
                .TextMatrix(lngRow, .ColIndex("�ⷿ")) = rsTemp!����
                .TextMatrix(lngRow, .ColIndex("ԭʼNO")) = rsTemp!ԭʼNO
                .TextMatrix(lngRow, .ColIndex("�ϴ�no")) = rsTemp!�ϴ�no
                .TextMatrix(lngRow, .ColIndex("no")) = rsTemp!NO
                .TextMatrix(lngRow, .ColIndex("�����")) = rsTemp!�����
                .TextMatrix(lngRow, .ColIndex("���ʱ��")) = Format(rsTemp!�������, "yyyy-mm-dd")
                rsTemp.MoveNext
            Next
        End With
    End If
End Sub

Private Sub GetALLData()
    '��ȡ������Ϣ
    Dim strSql As String
    Dim strԭʼNO As String
    
    On Error GoTo ErrHandle
    Set mrsData = Nothing
    If vsfLeft.Rows = 1 Then Exit Sub
    gstrSQL = "Select '��������' As ����, a.No, a.ҩƷid, c.����, c.����, c.���, a.����, a.�������, c.���㵥λ, d.��װ��λ, d.����ϵ��, a.����, a.ʵ������, a.�ɱ���, a.�ɱ����," & vbNewLine & _
        "       a.���ۼ�, a.���۽��, a.���, e.��Ʊ��, e.��Ʊ����, e.��Ʊ����, e.��Ʊ���, a.ժҪ" & vbNewLine & _
        "From ҩƷ�շ���¼ A, ҩƷ������� B, �շ���ĿĿ¼ C, �������� D, Ӧ����¼ E" & vbNewLine & _
        "Where a.No = b.����no And a.ҩƷid = c.Id And c.Id = d.����id And a.Id = e.�շ�id(+) And a.���� = 15 And b.ԭʼno = [1] And" & vbNewLine & _
        "      a.������� Is Not Null And (Mod(a.��¼״̬, 3) = 0 Or a.��¼״̬ = 1)" & vbNewLine & _
        "Union" & vbNewLine & _
        "Select 'ԭʼ����' As ����, a.No, a.ҩƷid, c.����, c.����, c.���, a.����, a.�������, c.���㵥λ, d.��װ��λ, d.����ϵ��, a.����, a.ʵ������, a.�ɱ���, a.�ɱ����," & vbNewLine & _
        "       a.���ۼ�, a.���۽��, a.���, e.��Ʊ��, e.��Ʊ����, e.��Ʊ����, e.��Ʊ���, a.ժҪ" & vbNewLine & _
        "From ҩƷ�շ���¼ A, �շ���ĿĿ¼ C, �������� D, Ӧ����¼ E" & vbNewLine & _
        "Where a.ҩƷid = c.Id And c.Id = d.����id And a.Id = e.�շ�id(+) And a.���� = 15 And a.No = [1] And a.������� Is Not Null And" & vbNewLine & _
        "      Mod(a.��¼״̬, 3) = 0"
    
    Set mrsData = zlDatabase.OpenSQLRecord(gstrSQL, "��ѯ��������", vsfLeft.TextMatrix(vsfLeft.Row, vsfLeft.ColIndex("ԭʼno")))
    Set mrsCloneDta = mrsData.Clone  '��¡���ݼ�
    Exit Sub
    
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub GetDetailsData()
    '��ȡ��ϸ����
    
End Sub

Private Sub Form_Load()
    Me.Height = 600 * 15
    Me.Width = 800 * 15
    Call SetControlLocation
    Call SetCBOValue
    dtpBeginDate.Value = DateAdd("d", -7, Sys.Currentdate)
    dtpEndDate.Value = Sys.Currentdate
    
    mintUnit = Val(zlDatabase.GetPara("���ĵ�λ", glngSys, 1712, "0"))
    With mFMT
        .FM_�ɱ��� = GetFmtString(mintUnit, g_�ɱ���)
        .FM_��� = GetFmtString(mintUnit, g_���)
        .FM_���ۼ� = GetFmtString(mintUnit, g_�ۼ�)
        .FM_���� = GetFmtString(mintUnit, g_����)
        .FM_ɢװ���ۼ� = GetFmtString(2, g_�ۼ�)
    End With
End Sub

Private Sub SetCBOValue()
    Dim arrtemp As Variant
    Dim i As Integer
    Dim strIndex As String
    Dim strTemp As String
    'Ϊ����������ֵ
    With cboDate
        .AddItem "һ������"
        .AddItem "��������"
        .AddItem "������"
        .AddItem "һ����"
        .AddItem "�Զ���"
        .ListIndex = 0
    End With
    
    ReDim arrtemp(UBound(Split(mStr�ⷿ, "|"))) As String
    
    With cbo�ⷿ
        .Clear
        .AddItem "���пⷿ"
        .ItemData(.NewIndex) = "0"
        For i = 0 To UBound(arrtemp) - 1
            strIndex = ""
            strTemp = ""
            arrtemp(i) = Split(mStr�ⷿ, "|")(i)
            strIndex = Mid(arrtemp(i), 1, InStr(1, arrtemp(i), ",") - 1)
            strTemp = Mid(arrtemp(i), InStr(1, arrtemp(i), ",") + 1)
            .AddItem strTemp
            .ItemData(.NewIndex) = strIndex
        Next
        
        .ListIndex = Val(mstr��ǰ�ⷿ) + 1
    End With
End Sub

Private Sub Form_Resize()
    Call SetControlLocation
    If sstInfo.Tab = 0 Then
        VSFList.Visible = False
        vsfAll.Visible = True
    Else
        VSFList.Visible = True
        vsfAll.Visible = False
    End If
End Sub

Private Sub optFloor_Click()
    VSFList.ColHidden(VSFList.ColIndex("no")) = True
'    vsfList.ColHidden(vsfList.ColIndex("ԭʼ")) = True
    If VSFList.Rows = 1 Then Exit Sub
    Call SetDetailsData
End Sub

Private Sub optMedi_Click()
    VSFList.ColHidden(VSFList.ColIndex("no")) = False
'    vsfList.ColHidden(vsfList.ColIndex("ԭʼ")) = False
    If VSFList.Rows = 1 Then Exit Sub
    Call SetDetailsData
End Sub


Private Sub picSplit_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    If Button = 1 Then
        If picLeft.Width + x < 1000 Then Exit Sub
        If sstInfo.Width - x < 2000 Then Exit Sub
        picLeft.Width = picLeft.Width + x
        picSplit.Left = picSplit.Left + x
        sstInfo.Width = sstInfo.Width - x
        sstInfo.Left = sstInfo.Left + x
        vsfLeft.Width = picLeft.ScaleWidth - 120
        vsfAll.Width = sstInfo.Width - 100
        VSFList.Width = sstInfo.Width - 100
    End If
End Sub

Private Sub sstInfo_Click(PreviousTab As Integer)
    If sstInfo.Tab = 0 Then
        VSFList.Visible = False
        vsfAll.Visible = True
    Else
        VSFList.Visible = True
        vsfAll.Visible = False
    End If
End Sub

Private Sub TxtNo_GotFocus()
    With txtNO
        .SelStart = 0
        .SelLength = Len(txtNO.Text)
    End With
End Sub

Private Sub TxtNo_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim intNO As Integer, strNo As String
    
    If KeyCode = vbKeyReturn Then
        '��ȡ���ݴ���
        intNO = 68
        If KeyCode = vbKeyReturn Then
            If Len(txtNO) < 8 And Len(txtNO) > 0 Then
                txtNO.Text = zlCommFun.GetFullNO(txtNO.Text, intNO, mlng�ⷿid)
            End If
            Call cmdFind_Click
        End If
    End If
End Sub

Private Sub vsfLeft_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If vsfLeft.Rows = 1 Then Exit Sub
    If OldRow <> NewRow Then
        Call GetALLData '��ѯ����
        If mrsData.RecordCount > 0 Then
            mrsData.Sort = " no asc"
            Call SetVsfDta(0) '��ֵ
            Call SetDetailsData
        End If
    End If
End Sub

Private Sub SetVsfDta(ByVal intModel As Integer)
    'Ϊ���ܺ���ϸ�ؼ���ֵ
    '���� intModel 0-����б��ѯ 1-�����ı��ѯ
    Dim lngRow As Long
    Dim lngCol As Long
    Dim str�ϴ�NO As String
    Dim strNewNO As String
    Dim strԭʼNO As String
    Dim strNo As String
    Dim strTemp As String
    Dim rsTemp As ADODB.Recordset
    Dim i As Integer
    Dim arrNo As Variant
    Dim dbl�ɹ����ϼ� As Double
    Dim dbl�ۼ۽��ϼ� As Double
    Dim dbl��۽��ϼ� As Double
    Dim dbl��Ʊ���ϼ� As Double
    Dim str��λ As String
    Dim str����ϵ�� As String
    Dim strNOType As String
    Dim strժҪ As String
    
    With vsfAll
        .Rows = 1
        str�ϴ�NO = vsfLeft.TextMatrix(vsfLeft.Row, vsfLeft.ColIndex("�ϴ�NO"))
        strNewNO = vsfLeft.TextMatrix(vsfLeft.Row, vsfLeft.ColIndex("NO"))
        strԭʼNO = vsfLeft.TextMatrix(vsfLeft.Row, vsfLeft.ColIndex("ԭʼNO"))
        If intModel = 1 Then
            Set mrsData = Nothing
            Set mrsData = mrsCloneDta.Clone
            mrsData.Sort = " no asc"
        End If
        
        If chkALLVisible1.Value = 1 Then '��ʾ��������
            '��ȡ��������
            mrsData.MoveFirst
            Do While Not mrsData.EOF
                strTemp = mrsData!NO
                If InStr(1, "," & strNo & ",", "," & strTemp & ",") = 0 Then
                    strNo = strNo & "," & strTemp
                End If
                mrsData.MoveNext
            Loop
            If strNo <> "" Then
                strNo = Mid(strNo, 2)
                arrNo = Split(strNo, ",")
                For i = 0 To UBound(arrNo)
                    strTemp = " no='" & arrNo(i) & "'"
                    mrsData.Filter = strTemp
                    dbl�ɹ����ϼ� = 0
                    dbl�ۼ۽��ϼ� = 0
                    dbl��۽��ϼ� = 0
                    dbl��Ʊ���ϼ� = 0
                    Do While Not mrsData.EOF
                        strNOType = mrsData!����
                        strժҪ = IIf(IsNull(mrsData!ժҪ), "", mrsData!ժҪ)
                        dbl�ɹ����ϼ� = dbl�ɹ����ϼ� + mrsData!�ɱ����
                        dbl�ۼ۽��ϼ� = dbl�ۼ۽��ϼ� + mrsData!���۽��
                        dbl��۽��ϼ� = dbl��۽��ϼ� + mrsData!���
                        dbl��Ʊ���ϼ� = dbl��Ʊ���ϼ� + IIf(IsNull(mrsData!��Ʊ���), 0, mrsData!��Ʊ���)
                        mrsData.MoveNext
                    Loop
                    .Rows = .Rows + 1
                    .Cell(flexcpPicture, .Rows - 1, .ColIndex("ԭʼ"), .Rows - 1, .ColIndex("ԭʼ")) = IIf(strNOType = "ԭʼ����", imgPicture.ListImages(1).Picture, "")
                    .TextMatrix(.Rows - 1, .ColIndex("ժҪ")) = strժҪ
                    .TextMatrix(.Rows - 1, .ColIndex("no")) = arrNo(i)
                    .TextMatrix(.Rows - 1, .ColIndex("�ɹ����")) = Format(dbl�ɹ����ϼ�, mFMT.FM_���)
                    .TextMatrix(.Rows - 1, .ColIndex("�ۼ۽��")) = Format(dbl�ۼ۽��ϼ�, mFMT.FM_���)
                    .TextMatrix(.Rows - 1, .ColIndex("���")) = Format(dbl��۽��ϼ�, mFMT.FM_���)
                    .TextMatrix(.Rows - 1, .ColIndex("��Ʊ���")) = Format(dbl��Ʊ���ϼ�, mFMT.FM_���)
                Next
            End If
        Else 'ֻ��ʾ��ǰ���ݺͲ�����ǰ���ݵĳ���ԭʼ����
            For i = 1 To 2
                If i = 1 Then
                    strTemp = " no='" & str�ϴ�NO & "'"
                Else
                    strTemp = " no='" & strNewNO & "'"
                End If
                dbl�ɹ����ϼ� = 0
                dbl�ۼ۽��ϼ� = 0
                dbl��۽��ϼ� = 0
                dbl��Ʊ���ϼ� = 0
                mrsData.Filter = strTemp
                Do While Not mrsData.EOF
                    strNOType = mrsData!����
                    strժҪ = IIf(IsNull(mrsData!ժҪ), "", mrsData!ժҪ)
                    dbl�ɹ����ϼ� = dbl�ɹ����ϼ� + mrsData!�ɱ����
                    dbl�ۼ۽��ϼ� = dbl�ۼ۽��ϼ� + mrsData!���۽��
                    dbl��۽��ϼ� = dbl��۽��ϼ� + mrsData!���
                    dbl��Ʊ���ϼ� = dbl��Ʊ���ϼ� + IIf(IsNull(mrsData!��Ʊ���), 0, mrsData!��Ʊ���)
                    mrsData.MoveNext
                Loop
                .Rows = .Rows + 1
                .Cell(flexcpPicture, .Rows - 1, .ColIndex("ԭʼ"), .Rows - 1, .ColIndex("ԭʼ")) = IIf(strNOType = "ԭʼ����", imgPicture.ListImages(1).Picture, "")
                .TextMatrix(.Rows - 1, .ColIndex("ժҪ")) = strժҪ
                .TextMatrix(.Rows - 1, .ColIndex("no")) = IIf(i = 1, str�ϴ�NO, strNewNO)
                .TextMatrix(.Rows - 1, .ColIndex("�ɹ����")) = Format(dbl�ɹ����ϼ�, mFMT.FM_���)
                .TextMatrix(.Rows - 1, .ColIndex("�ۼ۽��")) = Format(dbl�ۼ۽��ϼ�, mFMT.FM_���)
                .TextMatrix(.Rows - 1, .ColIndex("���")) = Format(dbl��۽��ϼ�, mFMT.FM_���)
                .TextMatrix(.Rows - 1, .ColIndex("��Ʊ���")) = Format(dbl��Ʊ���ϼ�, mFMT.FM_���)
            Next
        End If
        Call CheckValue
    End With
End Sub

Private Sub SetDetailsData()
    'Ϊ��ϸ���ֵ
    'Ϊ���ܺ���ϸ�ؼ���ֵ
    '���� intModel 0-����б��ѯ 1-�����ı��ѯ
    Dim lngRow As Long
    Dim lngCol As Long
    Dim str�ϴ�NO As String
    Dim strNewNO As String
    Dim strԭʼNO As String
    Dim strNo As String
    Dim strTemp As String
    Dim rsTemp As ADODB.Recordset
    Dim i As Integer
    Dim arrNo As Variant
    Dim dbl�ɹ����ϼ� As Double
    Dim dbl�ۼ۽��ϼ� As Double
    Dim dbl��۽��ϼ� As Double
    Dim dbl��Ʊ���ϼ� As Double
    Dim str��λ As String
    Dim str����ϵ�� As String
    Dim dbl��Ʊ��� As Double
    Dim strNOType As String
    
    With VSFList
        .Rows = 1
        str�ϴ�NO = vsfLeft.TextMatrix(vsfLeft.Row, vsfLeft.ColIndex("�ϴ�NO"))
        strNewNO = vsfLeft.TextMatrix(vsfLeft.Row, vsfLeft.ColIndex("NO"))
        strԭʼNO = vsfLeft.TextMatrix(vsfLeft.Row, vsfLeft.ColIndex("ԭʼNO"))
        
        Set mrsData = mrsCloneDta.Clone
        
        If chkALLVisible1.Value = 1 Then '��ʾ��������
        Else
            strTemp = " no='" & str�ϴ�NO & "' or no='" & strNewNO & " '"
            mrsData.Filter = strTemp
        End If
        '��ȡ��������
        If optFloor.Value = True Then '���յ��ݷ���
            mrsData.Sort = " no asc"
        Else
            mrsData.Sort = " ҩƷid,no asc"
        End If
        
        mrsData.MoveFirst
        Do While Not mrsData.EOF
            VSFList.Rows = VSFList.Rows + 1
            If optFloor.Value = True Then '���յ��ݷ���
                If VSFList.Rows > 2 Then
                    If mrsData!NO <> VSFList.TextMatrix(VSFList.Rows - 2, VSFList.ColIndex("no")) Then
                        VSFList.MergeCells = flexMergeFree
                        VSFList.MergeRow(VSFList.Rows - 1) = True
                        VSFList.Cell(flexcpText, VSFList.Rows - 1, 0, VSFList.Rows - 1, VSFList.Cols - 1) = "NO��" & VSFList.TextMatrix(VSFList.Rows - 2, VSFList.ColIndex("no"))
                        VSFList.Cell(flexcpBackColor, VSFList.Rows - 1, 0, VSFList.Rows - 1, VSFList.Cols - 1) = &HFFFFFF  ' &HFFC0C0
                        VSFList.Rows = VSFList.Rows + 1
                    End If
                End If
            End If
            Select Case mintUnit
                Case 0
                    str��λ = mrsData!���㵥λ
                    str����ϵ�� = 1
                Case 1
                    str��λ = mrsData!��װ��λ
                    str����ϵ�� = mrsData!����ϵ��
            End Select
            strNOType = mrsData!����
            VSFList.Cell(flexcpPicture, VSFList.Rows - 1, VSFList.ColIndex("ԭʼ"), VSFList.Rows - 1, VSFList.ColIndex("ԭʼ")) = IIf(strNOType = "ԭʼ����", imgPicture.ListImages(1).Picture, "")
            VSFList.TextMatrix(VSFList.Rows - 1, .ColIndex("no")) = mrsData!NO
            VSFList.TextMatrix(VSFList.Rows - 1, .ColIndex("ҩƷid")) = mrsData!ҩƷID
            VSFList.TextMatrix(VSFList.Rows - 1, .ColIndex("����ҩƷ���ƺ͹��")) = "[" & mrsData!���� & "]" & mrsData!���� & "(" & IIf(IsNull(mrsData!���), "", mrsData!���) & ")"
            VSFList.TextMatrix(VSFList.Rows - 1, .ColIndex("��������")) = IIf(IsNull(mrsData!����), "", mrsData!����) & "(" & IIf(IsNull(mrsData!����), "", mrsData!����) & ")"
            VSFList.TextMatrix(VSFList.Rows - 1, .ColIndex("����")) = Format(mrsData!ʵ������ / str����ϵ��, mFMT.FM_����) & "(" & str��λ & ")"
            VSFList.TextMatrix(VSFList.Rows - 1, .ColIndex("�ɹ���")) = Format(mrsData!�ɱ��� * str����ϵ��, mFMT.FM_�ɱ���)
            VSFList.TextMatrix(VSFList.Rows - 1, .ColIndex("�ɹ����")) = Format(mrsData!�ɱ����, mFMT.FM_���)
            VSFList.TextMatrix(VSFList.Rows - 1, .ColIndex("�ۼ�")) = Format(mrsData!���ۼ� * str����ϵ��, mFMT.FM_���ۼ�)
            VSFList.TextMatrix(VSFList.Rows - 1, .ColIndex("�ۼ۽��")) = Format(mrsData!���۽��, mFMT.FM_���)
            VSFList.TextMatrix(VSFList.Rows - 1, .ColIndex("���")) = Format(mrsData!���, mFMT.FM_���)
            VSFList.TextMatrix(VSFList.Rows - 1, .ColIndex("��Ʊ��")) = IIf(IsNull(mrsData!��Ʊ��), "", mrsData!��Ʊ��)
            VSFList.TextMatrix(VSFList.Rows - 1, .ColIndex("��Ʊ����")) = IIf(IsNull(mrsData!��Ʊ����), "", mrsData!��Ʊ����)
            VSFList.TextMatrix(VSFList.Rows - 1, .ColIndex("��Ʊ����")) = IIf(IsNull(mrsData!��Ʊ����), "", Format(mrsData!��Ʊ����, "yyyy-mm-dd"))
            dbl��Ʊ��� = IIf(IsNull(mrsData!��Ʊ���), 0, mrsData!��Ʊ���)
            VSFList.TextMatrix(VSFList.Rows - 1, .ColIndex("��Ʊ���")) = IIf(dbl��Ʊ��� = 0, "", Format(dbl��Ʊ���, mFMT.FM_���))
            
            mrsData.MoveNext
        Loop
        If optFloor.Value = True Then '���յ��ݷ���
            VSFList.Rows = VSFList.Rows + 1
            VSFList.MergeCells = flexMergeFree
            VSFList.MergeRow(VSFList.Rows - 1) = True
            VSFList.Cell(flexcpText, VSFList.Rows - 1, 0, VSFList.Rows - 1, VSFList.Cols - 1) = "NO��" & VSFList.TextMatrix(VSFList.Rows - 2, VSFList.ColIndex("no"))
            VSFList.Cell(flexcpBackColor, VSFList.Rows - 1, 0, VSFList.Rows - 1, VSFList.Cols - 1) = &HFFFFFF  ' &HFFC0C0
        End If
        Call CheckValue
    End With
End Sub

Private Sub vsfLeft_EnterCell()
    With vsfLeft
        .FocusRect = flexFocusSolid
    End With
End Sub

Private Sub CheckValue()
    Dim lngRow As Long
    Dim i As Long
    Dim lngCol As Long
    '���������Щ��Ϣ����ͬ��ͬ�в���ͬ�����ú�ɫ�����ע
    '���ܱ��
    With vsfAll
        For lngRow = 2 To .Rows - 1
            For lngCol = 2 To .Cols - 1
                If .TextMatrix(1, lngCol) <> .TextMatrix(lngRow, lngCol) Then
                    .Cell(flexcpForeColor, lngRow, lngCol, lngRow, lngCol) = vbRed
                End If
            Next
        Next
    End With
    '��ϸ���
    With VSFList
        If .Rows < 3 Then Exit Sub
        
        For lngRow = 1 To .Rows - 1
            For i = lngRow + 1 To .Rows - 1
                If i > .Rows - 1 Then Exit For
                If .TextMatrix(lngRow, .ColIndex("ҩƷid")) = .TextMatrix(i, .ColIndex("ҩƷid")) Then
                    For lngCol = 3 To .Cols - 1
                        If .TextMatrix(lngRow, lngCol) <> .TextMatrix(i, lngCol) Then
                            .Cell(flexcpForeColor, i, lngCol, i, lngCol) = vbRed
                        End If
                    Next
                End If
            Next
        Next
    End With
End Sub

Public Sub ShowMe(ByVal frmPar As Form, ByVal str�ⷿ As String, ByVal str��ǰ�ⷿ As Long)
    mStr�ⷿ = str�ⷿ
    mstr��ǰ�ⷿ = str��ǰ�ⷿ
    Me.Show vbModal, frmPar
End Sub

