VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "VSFLEX8.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLisRptMicrobiology 
   BorderStyle     =   0  'None
   Caption         =   "΢���ﱨ��"
   ClientHeight    =   8325
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10020
   Icon            =   "frmLisRptMicrobiology.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8325
   ScaleWidth      =   10020
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Frame fraNS 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   45
      Left            =   -150
      MousePointer    =   7  'Size N S
      TabIndex        =   10
      Top             =   3315
      Width           =   5055
   End
   Begin VB.PictureBox pic�ٴ����� 
      BorderStyle     =   0  'None
      Height          =   2280
      Left            =   4575
      ScaleHeight     =   2280
      ScaleWidth      =   3900
      TabIndex        =   8
      Top             =   4215
      Visible         =   0   'False
      Width           =   3900
      Begin VB.TextBox txt�ο� 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   1950
         Left            =   255
         Locked          =   -1  'True
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   300
         Width           =   3600
      End
   End
   Begin VB.PictureBox picTab 
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      Height          =   4110
      Left            =   75
      ScaleHeight     =   4110
      ScaleWidth      =   9300
      TabIndex        =   6
      Top             =   3555
      Width           =   9300
      Begin XtremeSuiteControls.TabControl TabThis 
         Height          =   3930
         Left            =   90
         TabIndex        =   7
         Top             =   60
         Width           =   8715
         _Version        =   589884
         _ExtentX        =   15372
         _ExtentY        =   6932
         _StockProps     =   64
      End
   End
   Begin VB.PictureBox picDetail 
      BorderStyle     =   0  'None
      Height          =   2715
      Left            =   4890
      ScaleHeight     =   2715
      ScaleWidth      =   6900
      TabIndex        =   4
      Top             =   5580
      Width           =   6900
      Begin VSFlex8Ctl.VSFlexGrid vsfDetail 
         Height          =   3270
         Left            =   195
         TabIndex        =   5
         Top             =   195
         Width           =   6750
         _cx             =   11906
         _cy             =   5768
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483634
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
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   270
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
   End
   Begin VB.PictureBox PicVsf 
      BorderStyle     =   0  'None
      Height          =   2865
      Left            =   -15
      ScaleHeight     =   2865
      ScaleWidth      =   9540
      TabIndex        =   1
      Top             =   45
      Width           =   9540
      Begin VB.Frame fraNS1 
         BorderStyle     =   0  'None
         Height          =   45
         Left            =   8715
         MousePointer    =   7  'Size N S
         TabIndex        =   17
         Top             =   1035
         Width           =   1875
      End
      Begin VB.Frame fraSW 
         BorderStyle     =   0  'None
         Height          =   2820
         Left            =   7275
         MousePointer    =   9  'Size W E
         TabIndex        =   16
         Top             =   90
         Width           =   45
      End
      Begin VB.PictureBox picResult 
         BorderStyle     =   0  'None
         Height          =   1500
         Left            =   7140
         ScaleHeight     =   1500
         ScaleWidth      =   2805
         TabIndex        =   13
         Top             =   105
         Width           =   2800
         Begin VB.TextBox txtResult 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   1005
            Left            =   90
            Locked          =   -1  'True
            MaxLength       =   200
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   14
            Top             =   255
            Width           =   4215
         End
         Begin VB.Label lblResult 
            Caption         =   "�������"
            Height          =   195
            Left            =   150
            TabIndex        =   15
            Top             =   60
            Width           =   1035
         End
      End
      Begin VB.PictureBox picComment 
         BorderStyle     =   0  'None
         Height          =   1500
         Left            =   5715
         ScaleHeight     =   1500
         ScaleWidth      =   3900
         TabIndex        =   11
         Top             =   1440
         Width           =   3900
         Begin VB.TextBox txtComment 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   540
            Left            =   120
            Locked          =   -1  'True
            MaxLength       =   200
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   12
            Top             =   420
            Width           =   7305
         End
         Begin VB.Label lblComment 
            Caption         =   "���鱸ע"
            Height          =   195
            Left            =   135
            TabIndex        =   18
            Top             =   45
            Width           =   1035
         End
      End
      Begin VB.CheckBox chkLast 
         Caption         =   "�ϴν��"
         Height          =   180
         Left            =   45
         TabIndex        =   3
         Top             =   15
         Width           =   1455
      End
      Begin VSFlex8Ctl.VSFlexGrid vsf 
         Height          =   2295
         Left            =   30
         TabIndex        =   2
         Top             =   255
         Width           =   5625
         _cx             =   9922
         _cy             =   4048
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483634
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
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   270
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
   End
   Begin MSComctlLib.StatusBar sbrInfo 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   7965
      Width           =   10020
      _ExtentX        =   17674
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2822
            MinWidth        =   2822
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4586
            MinWidth        =   4586
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2822
            MinWidth        =   2822
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4586
            MinWidth        =   4586
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   2699
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
End
Attribute VB_Name = "frmLisRptMicrobiology"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum mCol
    ������� = 0: ID: ϸ������: �������: ��������: �ϴξ������
    ���������� = 1: ҩ������: ������: �����־: �ϴν��: �ϴα�־
End Enum
Private mlngRedoNumber As Long '��������
Private mlng�걾ID As Long, mlng������� As Long, mlngϸ��id As Long
Public mlngMode As Long

Public Sub zlRefresh(ByVal lng�걾id As Long, lng������� As Long)
    '
    Dim rs As New ADODB.Recordset, mstrSQL As String
    
    On Error GoTo Errhand
    mlng�걾ID = lng�걾id
    mlng������� = lng�������
    mlngϸ��id = 0
    
    vsf.Rows = 2
    vsf.Cell(flexcpText, 1, 0, 1, vsf.Cols - 1) = ""
    
    mstrSQL = "SELECT A.������,A.������,A.����ʱ��,A.�����,A.���ʱ��,A.���鱸ע,A.��ע FROM ����걾��¼ A WHERE A.ID= [1] "
    Set rs = zlDatabase.OpenSQLRecord(mstrSQL, Me.Caption, lng�걾id)
    If Not rs.EOF Then
        mlngRedoNumber = Val("" & rs("������"))
        Me.txtComment = "" & rs("���鱸ע")
        Me.txtResult = "" & rs("��ע")
        
        With sbrInfo
            .Panels(1).Text = "�����ˣ�" & rs("������")
            .Panels(2).Text = "����ʱ�䣺" & IIF(IsNull(rs("����ʱ��")), "", Format(rs("����ʱ��"), "yyyy-MM-dd hh:mm"))
            .Panels(3).Text = "����ˣ�" & rs("�����")
            .Panels(4).Text = "���ʱ�䣺" & IIF(IsNull(rs("���ʱ��")), "", Format(rs("���ʱ��"), "yyyy-MM-dd hh:mm"))
        End With
    Else
        mlngRedoNumber = 0
        Me.txtComment = ""
        Me.txtResult = ""
        
        With sbrInfo
            .Panels(1).Text = "�����ˣ�"
            .Panels(2).Text = "����ʱ�䣺"
            .Panels(3).Text = "����ˣ�"
            .Panels(4).Text = "���ʱ�䣺"
        End With
    End If
    
'    mstrSQL = "SELECT C.������ĿID FROM ����걾��¼ A,����������Ŀ B,���鱨����Ŀ C " & _
'                    "WHERE A.ID=B.�걾ID And B.������ĿID=C.������ĿID " & _
'                        "AND A.ID= [1] "
'    Set rs = zlDatabase.OpenSQLRecord(mstrSQL, Me.Caption, lng�걾id)
'    If rs.BOF = False Then
'        mlngItemID = Nvl(rs("������ĿID"), 0)
'    Else
'        mlngItemID = 0
'    End If
'
    '��ʼ�����
    initVsf
    mstrSQL = "SELECT Distinct B.����, B.ID,D.������,B.������ AS ������Ŀ," & _
                    "A.������ AS ������,A.�������� as �������,'' as �ϴν�� " & _
                    "FROM ������ͨ��� A,����ϸ�� B,����걾��¼ D " & _
                    "WHERE A.ϸ��id = B.ID And D.����� is Not null " & _
                        "AND A.��¼���� = [1] " & _
                        "AND D.ID=A.����걾ID " & _
                        "AND D.ID= [2] Order by B.����"
    Set rs = zlDatabase.OpenSQLRecord(mstrSQL, Me.Caption, lng�������, lng�걾id)
    Do Until rs.EOF
        With vsf
            .TextMatrix(.Rows - 1, mCol.�������) = .Rows - 1
            .TextMatrix(.Rows - 1, mCol.ID) = Val("" & rs!ID)
            .TextMatrix(.Rows - 1, mCol.ϸ������) = Trim("" & rs!������Ŀ)
            .TextMatrix(.Rows - 1, mCol.�������) = Trim("" & rs!������)
            .TextMatrix(.Rows - 1, mCol.��������) = Trim("" & rs!�������)
            .TextMatrix(.Rows - 1, mCol.�ϴξ������) = Trim("" & rs!�ϴν��)
            
            If mlngϸ��id = 0 Then mlngϸ��id = Val("" & rs!ID)
            
            .Rows = .Rows + 1
        End With
        rs.MoveNext
    Loop
    vsf.Rows = vsf.Rows - 1
    
    vsf.Cell(flexcpBackColor, 0, 0, vsf.Rows - 1, 0) = &HFDD6C6
    If vsf.Rows > 1 Then Call vsf.Select(1, 1)
    
    Call vsf_RowColChange
    
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
    Resume
    End If
End Sub

Private Sub Refresh_vsfDetail(ByVal lng�걾id As Long, ByVal lng������� As Long, ByVal lngϸ��ID As Long)
    Dim rs As ADODB.Recordset
    Dim strSQL As String
    
    Call initVsfDetail
    strSQL = "SELECT C.ϸ��ID AS Key,B.ID,B.������ AS ����������, A.��� AS ������, " & _
            "DECODE(A.�������,'R','R-��ҩ','I','I-�н�','S','S-����','') AS �������, " & _
            "DECODE(A.ҩ������,1,'1-MIC',2,'2-DISK',3,'3-K-B','') As ҩ������ " & _
             "FROM ����ҩ����� A, �����ÿ����� B,������ͨ��� C " & _
            "Where A.������ID = B.ID And C.ID=A.ϸ�����ID AND C.��¼����=A.��¼���� AND C.����걾id= [1] AND C.��¼����= [2] And C.ϸ��ID=[3] Order By B.����"
    On Error GoTo errH
    Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng�걾id, lng�������, lngϸ��ID)
    Do Until rs.EOF
        With vsfDetail
            .TextMatrix(.Rows - 1, mCol.�������) = .Rows - 1
            .TextMatrix(.Rows - 1, mCol.����������) = Trim("" & rs!����������)
            .TextMatrix(.Rows - 1, mCol.ҩ������) = Trim("" & rs!ҩ������)
            .TextMatrix(.Rows - 1, mCol.������) = Trim("" & rs!������)
            .TextMatrix(.Rows - 1, mCol.�����־) = Trim("" & rs!�������)
            
            .TextMatrix(.Rows - 1, mCol.�ϴν��) = ""
            .TextMatrix(.Rows - 1, mCol.�ϴα�־) = ""
            If chkLast.value = 1 Then
            
            End If
            .Rows = .Rows + 1
        End With
        rs.MoveNext
    Loop
    Call Check_ColWidth
    
    vsfDetail.Rows = vsfDetail.Rows - 1
    vsfDetail.Cell(flexcpBackColor, 0, 0, vsfDetail.Rows - 1, 0) = &HFDD6C6
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
Private Sub IntiTab()

    On Error Resume Next

    With Me.TabThis
        Set .Icons = zlCommFun.GetPubIcons
        .PaintManager.Appearance = xtpTabAppearanceExcel
        .PaintManager.Color = xtpTabColorOffice2003
        .PaintManager.OneNoteColors = False
        .PaintManager.BoldSelected = True

        .PaintManager.ClientFrame = xtpTabFrameSingleLine
'        .PaintManager.Position = xtpTabPositionBottom
        .InsertItem(0, "������", Me.picDetail.Hwnd, conMenu_Tool_Monitor).Tag = "������"
        '.InsertItem(1, "�ٴ�����", Me.pic�ٴ�����.Hwnd, conMenu_View_ToolBar_Text).Tag = "΢�����ٴ�����"
        
        .PaintManager.Layout = xtpTabLayoutAutoSize
        .PaintManager.ShowIcons = True
        .Item(0).Selected = True
        
    End With
End Sub

Private Sub initVsf()
    With Me.vsf
        .BackColor = &H80000005
        .Appearance = flex3DLight
        .BorderStyle = flexBorderFlat
        .BackColorFixed = &HFDD6C6
        .GridLinesFixed = flexGridFlat
        .RowHeightMin = 300
        .Editable = flexEDNone
        .Clear
        .Rows = 2: .FixedRows = 1
        .Cols = 6: .FixedCols = 0
        
        .TextMatrix(0, mCol.�������) = "": .ColWidth(mCol.�������) = 300: .ColAlignment(mCol.�������) = flexAlignRightCenter
        .TextMatrix(0, mCol.ϸ������) = "ϸ������": .ColWidth(mCol.ϸ������) = 2500: .ColAlignment(mCol.ϸ������) = flexAlignLeftCenter
        .TextMatrix(0, mCol.�������) = "�������": .ColWidth(mCol.�������) = 1500: .ColAlignment(mCol.�������) = flexAlignLeftCenter
        .TextMatrix(0, mCol.��������) = "��������": .ColWidth(mCol.��������) = 2000: .ColAlignment(mCol.��������) = flexAlignLeftCenter
        .TextMatrix(0, mCol.�ϴξ������) = "��������": .ColWidth(mCol.�ϴξ������) = 1500: .ColAlignment(mCol.�ϴξ������) = flexAlignLeftCenter
    End With
        
    chkLast.value = Val(zlDatabase.GetPara("�ϴν��", glngSys, mlngMode, 0))
    
    Call initVsfDetail
End Sub

Private Sub initVsfDetail()
    With Me.vsfDetail
        .BackColor = &H80000005
        .Appearance = flex3DLight
        .BorderStyle = flexBorderFlat
        .BackColorFixed = &HFDD6C6
        .GridLinesFixed = flexGridFlat
        .RowHeightMin = 300
        .Editable = flexEDNone
        
        .Clear
        .Rows = 2: .FixedRows = 1
        .Cols = 7: .FixedCols = 0
        
        .TextMatrix(0, mCol.�������) = "": .ColWidth(mCol.�������) = 300: .ColAlignment(mCol.�������) = flexAlignRightCenter
        .TextMatrix(0, mCol.����������) = "����������": .ColWidth(mCol.����������) = 2500: .ColAlignment(mCol.����������) = flexAlignLeftCenter
        .TextMatrix(0, mCol.ҩ������) = "ҩ������": .ColWidth(mCol.ҩ������) = 850: .ColAlignment(mCol.ҩ������) = flexAlignLeftCenter
        .TextMatrix(0, mCol.������) = "������": .ColWidth(mCol.������) = 1300: .ColAlignment(mCol.������) = flexAlignLeftCenter
        .TextMatrix(0, mCol.�����־) = "�������": .ColWidth(mCol.�����־) = 1000: .ColAlignment(mCol.�����־) = flexAlignLeftCenter
        .TextMatrix(0, mCol.�ϴν��) = "�ϴν��": .ColWidth(mCol.�ϴν��) = 1300: .ColAlignment(mCol.�ϴν��) = flexAlignLeftCenter
        .TextMatrix(0, mCol.�ϴα�־) = "�ϴ�����": .ColWidth(mCol.�ϴα�־) = 1300: .ColAlignment(mCol.�ϴα�־) = flexAlignLeftCenter
    End With
End Sub

Private Sub chkLast_Click()
    Call Check_ColWidth
End Sub

Private Sub Form_Load()
    
    Call IntiTab
    Call initVsf
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    With Me.PicVsf
        .Left = 0
        .Top = 0
        .Height = Me.fraNS.Top
        .Width = Me.ScaleWidth
    
    End With
    Me.fraNS.Left = 0
    Me.fraNS.Width = Me.ScaleWidth
    
    With Me.picTab
        .Left = 0
        .Top = Me.fraNS.Top + fraNS.Height
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight - .Top - Me.sbrInfo.Height
    End With
    Call PicVsf_Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call zlDatabase.SetPara("�ϴν��", Me.chkLast.value, glngSys, mlngMode) '����ҽ���´�,סԺҽ���´�
End Sub

Private Sub fraNS_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'
    On Error Resume Next
    If Button = 1 Then
        If PicVsf.Height + Y < 1000 Or PicVsf.Height - Y < 1000 Then
            PicVsf.Height = 1100
            Exit Sub
        End If
        fraNS.Top = fraNS.Top + Y
        PicVsf.Height = PicVsf.Height + Y
        picTab.Top = picTab.Top + Y
        picTab.Height = picTab.Height - Y
    End If
End Sub

Private Sub fraNS1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If Button = 1 Then
        If picResult.Height + Y < 1000 Or picResult.Height - Y < 1000 Then
            picResult.Height = 1100
            Exit Sub
        End If
        fraNS1.Top = fraNS1.Top + Y
        picResult.Height = picResult.Height + Y
        picComment.Top = picComment.Top + Y
        picComment.Height = picComment.Height - Y
    End If
End Sub

Private Sub fraSW_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If Button = 1 Then
        If vsf.Width + X < 1000 Or vsf.Width - X < 1000 Then
            'vsf.Width = 1100
            Exit Sub
        End If
        
        If picResult.Width - X < 1000 Then
            'picResult.Width = 1100
            Exit Sub
        End If

        vsf.Width = vsf.Width + X
        
        fraSW.Left = fraSW.Left + X
        
        fraNS1.Left = fraNS1.Left + X
        fraNS1.Width = fraNS1.Width - X
        picResult.Left = picResult.Left + X
        picResult.Width = picResult.Width - X
        
        picComment.Left = picResult.Left
        picComment.Width = picResult.Width
    End If
End Sub

Private Sub picComment_Resize()
    On Error Resume Next
    With Me.lblComment
        .Left = 10
        .Top = 10
        .Width = Me.picComment.ScaleWidth - 10
    End With
    With Me.txtComment
        .Left = 10
        .Top = lblComment.Top + lblComment.Height + 20
        .Width = Me.picComment.ScaleWidth - 10
        .Height = Me.picComment.ScaleHeight - .Top
    End With
End Sub

Private Sub picDetail_Resize()
 On Error Resume Next
 With Me.vsfDetail
     .Left = 0
     
    .Width = Me.picDetail.ScaleWidth
    .Height = Me.picDetail.ScaleHeight - .Top
 End With
End Sub

Private Sub picResult_Resize()
    
    With lblResult
        .Left = 10
        .Top = 10
        .Width = Me.picResult.ScaleWidth - 10
    End With
    With Me.txtResult
        .Left = 10
        .Top = lblResult.Top + lblResult.Height + 20
        .Width = Me.picResult.ScaleWidth - 10
        .Height = Me.picResult.ScaleHeight - .Top
    End With
End Sub

Private Sub picTab_Resize()
    With Me.TabThis
        .Top = 0
        .Left = 0
        .Width = Me.picTab.ScaleWidth
        .Height = Me.picTab.ScaleHeight
    End With
End Sub

Private Sub PicVsf_Resize()
    On Error Resume Next
    With Me.vsf
        .Top = Me.chkLast.Top + Me.chkLast.Height + 30
        .Left = 0
        .Width = Me.fraSW.Left
        .Height = Me.PicVsf.ScaleHeight - .Top - 10
    End With
    
    With fraSW
        .Height = Me.PicVsf.ScaleHeight
        .Top = Me.PicVsf.ScaleTop
    End With
    
    With fraNS1
        .Left = Me.fraSW.Left + fraSW.Width
        .Width = Me.PicVsf.ScaleWidth - .Left
    End With
    
    With Me.picResult
        .Top = 0
        .Left = Me.fraSW.Left + fraSW.Width
        .Width = Me.PicVsf.ScaleWidth - .Left - 10
        .Height = Me.fraNS1.Top
    End With
    
    With Me.picComment
        .Top = Me.fraNS1.Top + Me.fraNS1.Height
        .Left = Me.picResult.Left
        .Width = Me.picResult.Width
        .Height = Me.PicVsf.ScaleHeight - .Top
    End With
End Sub

Private Sub pic�ٴ�����_Resize()
    With Me.txt�ο�
        .Left = 0
        .Top = 0
        .Width = Me.pic�ٴ�����.ScaleWidth
        .Height = Me.pic�ٴ�����.ScaleHeight
    End With
End Sub

Private Sub vsf_RowColChange()
    mlngϸ��id = Val(vsf.TextMatrix(vsf.Row, mCol.ID))
    Call Refresh_vsfDetail(mlng�걾ID, mlng�������, mlngϸ��id)
End Sub

Private Sub Check_ColWidth()
    '���ݿؼ�״̬�������п�
    
    vsf.ColWidth(mCol.�ϴξ������) = IIF(chkLast.value = 0, 0, 1000)
    
    With vsfDetail
        .ColWidth(mCol.�ϴν��) = IIF(chkLast.value = 0, 0, 1000)
        .ColWidth(mCol.�ϴα�־) = IIF(chkLast.value = 0, 0, 1000)
    End With

End Sub

