VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMediLimit 
   Caption         =   "ҩƷ��������"
   ClientHeight    =   6675
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10575
   Icon            =   "frmMediLimit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   10575
   StartUpPosition =   1  '����������
   Begin VSFlex8Ctl.VSFlexGrid vsfStore 
      Height          =   2445
      Left            =   6240
      TabIndex        =   20
      Top             =   1920
      Visible         =   0   'False
      Width           =   3495
      _cx             =   6165
      _cy             =   4313
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
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16769992
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
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   275
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmMediLimit.frx":058A
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
      ExplorerBar     =   1
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
   Begin MSComctlLib.ImageList imgStoreRoom 
      Left            =   9960
      Top             =   2280
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
            Picture         =   "frmMediLimit.frx":060D
            Key             =   "StroeRoomPic"
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txt���� 
      Height          =   300
      Left            =   6600
      TabIndex        =   17
      Top             =   600
      Visible         =   0   'False
      Width           =   1845
   End
   Begin VB.ComboBox cboDrugUnit 
      Height          =   300
      ItemData        =   "frmMediLimit.frx":6E6F
      Left            =   1800
      List            =   "frmMediLimit.frx":6E7F
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1065
      Width           =   2400
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   14
      Top             =   6300
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmMediLimit.frx":6EAB
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13573
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
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
   Begin VB.Frame fraFunc 
      BorderStyle     =   0  'None
      Height          =   645
      Left            =   0
      TabIndex        =   13
      Top             =   5040
      Width           =   9810
      Begin VB.CommandButton cmdFilter 
         Caption         =   "����(&T)"
         Height          =   350
         Left            =   5350
         TabIndex        =   16
         Top             =   165
         Width           =   1100
      End
      Begin VB.CommandButton cmdApply 
         Caption         =   "Ӧ���ڱ���(&O)"
         Height          =   350
         Left            =   3990
         TabIndex        =   11
         Top             =   165
         Width           =   1365
      End
      Begin VB.CommandButton cmdRestore 
         Caption         =   "�ָ�(&R)"
         Height          =   350
         Left            =   2685
         Picture         =   "frmMediLimit.frx":773D
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   165
         Width           =   1290
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "ȫ�����(&C)"
         Height          =   350
         Left            =   1380
         Picture         =   "frmMediLimit.frx":7887
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   165
         Width           =   1290
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "����(&S)"
         Height          =   350
         Left            =   6720
         TabIndex        =   7
         Top             =   165
         Width           =   1100
      End
      Begin VB.CommandButton cmdHelp 
         Caption         =   "����(&H)"
         Height          =   350
         Left            =   90
         Picture         =   "frmMediLimit.frx":79D1
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   165
         Width           =   1100
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "�ر�(&X)"
         Height          =   350
         Left            =   7920
         TabIndex        =   8
         Top             =   165
         Width           =   1100
      End
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Left            =   -195
      TabIndex        =   6
      Top             =   1440
      Width           =   9810
   End
   Begin VB.ComboBox cboRoom 
      Height          =   276
      Left            =   1800
      TabIndex        =   2
      Text            =   "cboRoom"
      Top             =   585
      Width           =   2400
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfLimit 
      Height          =   1335
      Left            =   120
      TabIndex        =   15
      Top             =   1920
      Width           =   5895
      _cx             =   10398
      _cy             =   2355
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
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483633
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483632
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   3
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   29
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmMediLimit.frx":7B1B
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
   Begin VB.Label lbl���� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "����(&F)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   5880
      TabIndex        =   19
      Top             =   660
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.Label lblComment1 
      AutoSize        =   -1  'True
      Caption         =   "��F3������������"
      Height          =   180
      Left            =   8550
      TabIndex        =   18
      Top             =   660
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.Label lblDrugUnit 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "��λ(&U)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   1080
      TabIndex        =   3
      Top             =   1125
      Width           =   630
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ҩƷ�ⷿ(&D)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   720
      TabIndex        =   1
      Top             =   645
      Width           =   990
   End
   Begin VB.Image imgNote 
      Height          =   480
      Left            =   75
      Picture         =   "frmMediLimit.frx":7EF7
      Top             =   90
      Width           =   480
   End
   Begin VB.Label lblnote 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "    ѡ��ҩƷ�ⷿ��ָ���ÿⷿҩƷ�Ĵ���������������ҩƷ�Ĺ���Ҫ�󣬿���ͬʱָ�����̵����ԺͿⷿ��λ��"
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   720
      TabIndex        =   0
      Top             =   150
      Width           =   7725
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblLimit 
      AutoSize        =   -1  'True
      Caption         =   "ҩƷ�ڸ��ⷿ���޶����̵�Ҫ��(&T)��"
      Height          =   180
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   2970
   End
End
Attribute VB_Name = "frmMediLimit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'---------------------------------------------------
'˵����
'   1����ǰ���ʣ���me.tag����,�ֱ�Ϊ"5","6","7"
'   2����ǰ״̬����me.cmdClose.tag���棬�ֱ�Ϊ"�޸�"��"����"�����ϼ�������
'   3��ָ��ҩƷ����me.lblMedi.tag���棬���ϼ���������Դ��ݣ�Ҳ���Բ�����
'---------------------------------------------------
Public strPrivs As String       '��ǰ�û����еı�����Ȩ��

Private mrsNormal As New ADODB.Recordset
Private mintCount As Integer
Private mlng�ⷿID As Long
Private mlngFind As Long
Private mlngFindFirst As Long
Private mrsFindName As ADODB.Recordset
Private mblnChanged As Boolean
Private mstr���� As String
Private mstr����ID  As String
Private mstr���� As String
Private mblnActive As Boolean
Private mintҩƷ������ʾ As Integer         '0-��ʾͨ������1-��ʾ��Ʒ����2-ͬʱ��ʾͨ��������Ʒ��
Private mlngRow As Long     '������¼�����λ��ťʱ����
Private Const mlngBorderColor As Long = &H8000000D     'ѡ���б߿���ɫ
Private Const mlngNoneBorderColor As Long = &HE0E0E0    ' ûѡ���б߿���ɫ
Private Sub FindGridRow(ByVal strInput As String)
    Dim lngStart As Long, lngRows As Long
    Dim str���� As String, str���� As String, str���� As String
    Dim str�������� As String
    Dim n As Integer
    Dim blnEnd As Boolean
    Dim lngFindRow As Long
    Dim strFindStyle As String
    Dim strTmp As String
    
    '����ҩƷ
    On Error GoTo errHandle
    If strInput = txt����.Tag Then
        '��ʾ������һ����¼
        If mlngFind >= vsfLimit.Rows - 1 Then
            lngStart = 0
        Else
            lngStart = mlngFind
        End If
    Else
        '��ʾ�µĲ���
        lngStart = 0
        mlngFindFirst = 0
        txt����.Tag = strInput
        
        strFindStyle = IIf(GetSetting("ZLSOFT", "����ģ��\����", "����ƥ��", 0) = "0", "%", "")
        
        Set mrsFindName = New ADODB.Recordset

        gstrSql = "Select Distinct A.Id,A.���� From �շ���ĿĿ¼ A,�շ���Ŀ���� B" & _
                 " Where A.Id =B.�շ�ϸĿid And A.���=[1] "

        If IsNumeric(Replace(strInput, "-", "")) Then       '����ȫ�����֣������һ��"-"��ʱֻƥ�����
            gstrSql = gstrSql & " And A.���� Like [2] Or B.���� Like [2] And B.����=3 "
        ElseIf zlStr.IsCharAlpha(strInput) Then          '����ȫ����ĸʱֻƥ�����
            gstrSql = gstrSql & " And B.���� Like [3] "
        ElseIf zlStr.IsCharChinese(strInput) Then        '����ȫ�Ǻ���ʱֻƥ������
            gstrSql = gstrSql & " And B.���� Like [3] "
        Else
            gstrSql = gstrSql & " And (A.���� Like [2] Or B.���� Like [3] Or B.���� Like [3] )"
        End If
        
        gstrSql = gstrSql & " Order By A.���� "
                 
        Set mrsFindName = zldatabase.OpenSQLRecord(gstrSql, "ȡƥ���ҩƷID", Me.Tag, strInput & "%", strFindStyle & strInput & "%")
        
        If mrsFindName.RecordCount = 0 Then Exit Sub
    End If
    
    '��ʼ����
    If mrsFindName.State <> adStateOpen Then Exit Sub
    If mrsFindName.RecordCount = 0 Then Exit Sub

    lngStart = lngStart + 1
    lngRows = vsfLimit.Rows - 1
    
    With mrsFindName
        If .EOF Then .MoveFirst
        
        Do While Not .EOF
            lngFindRow = vsfLimit.FindRow(!����, 0, vsfLimit.ColIndex("����"), True, True)
            If lngFindRow > 0 Then
                vsfLimit.Select lngFindRow, 1, lngFindRow, vsfLimit.Cols - 1
                vsfLimit.TopRow = lngFindRow
                mlngFind = lngFindRow
                
                '��¼�ҵ��ĵ�1����¼
                If mlngFindFirst = 0 Then mlngFindFirst = mlngFind
                
                mrsFindName.MoveNext
                Exit Do
            End If
            mrsFindName.MoveNext
    
            '��������ˣ��򷵻ص�1����¼
            If .EOF And lngFindRow = -1 Then
                vsfLimit.Select mlngFindFirst, 1, mlngFindFirst, vsfLimit.Cols - 1
                vsfLimit.TopRow = mlngFindFirst
                mlngFind = mlngFindFirst
            End If
        Loop
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub



Private Sub IniGrid()
    With vsfLimit
        .Redraw = flexRDNone
        .Rows = 1
        .SelectionMode = flexSelectionFree
        .ExplorerBar = flexExSortShowAndMove
        .Editable = flexEDNone
        
        .ColComboList(.ColIndex("��λ")) = "..."
        
        .ColWidth(.ColIndex("��Ʒ��")) = IIf(mintҩƷ������ʾ = 2, 2000, 0)
        
        .TextMatrix(0, .ColIndex("����")) = "������"
        .ColWidth(.ColIndex("���")) = 1500 'IIf(Me.Tag <> "7", 1500, 0)
        .ColWidth(.ColIndex("����")) = 1200
        .ColHidden(.ColIndex("ԭ����")) = IIf(Me.Tag = "7", False, True)
        
        If InStr(1, strPrivs, "�����޿���") > 0 Then
            .ColHidden(.ColIndex("����")) = False
            .ColHidden(.ColIndex("����")) = False
            
            If .ColWidth(.ColIndex("����")) = 0 Then .ColWidth(.ColIndex("����")) = 1050
            If .ColWidth(.ColIndex("����")) = 0 Then .ColWidth(.ColIndex("����")) = 1050
        Else
            .ColHidden(.ColIndex("����")) = True
            .ColHidden(.ColIndex("����")) = True
        End If
        
        If InStr(1, strPrivs, "�̵���������") > 0 Then
            .ColHidden(.ColIndex("����")) = False
            .ColHidden(.ColIndex("����")) = False
            .ColHidden(.ColIndex("����")) = False
            .ColHidden(.ColIndex("����")) = False
            
            If .ColWidth(.ColIndex("����")) = 0 Then .ColWidth(.ColIndex("����")) = 500
            If .ColWidth(.ColIndex("����")) = 0 Then .ColWidth(.ColIndex("����")) = 500
            If .ColWidth(.ColIndex("����")) = 0 Then .ColWidth(.ColIndex("����")) = 500
            If .ColWidth(.ColIndex("����")) = 0 Then .ColWidth(.ColIndex("����")) = 500
        Else
            .ColHidden(.ColIndex("����")) = True
            .ColHidden(.ColIndex("����")) = True
            .ColHidden(.ColIndex("����")) = True
            .ColHidden(.ColIndex("����")) = True
        End If
        
        .ColDataType(.ColIndex("����")) = flexDTDouble
        .ColDataType(.ColIndex("����")) = flexDTDouble
        
        .Cell(flexcpForeColor, 0, .ColIndex("��������")) = vbBlue
        .Cell(flexcpForeColor, 0, .ColIndex("����")) = vbBlue
        .Cell(flexcpForeColor, 0, .ColIndex("����")) = vbBlue
        .Cell(flexcpForeColor, 0, .ColIndex("����")) = vbBlue
        .Cell(flexcpForeColor, 0, .ColIndex("����")) = vbBlue
        .Cell(flexcpForeColor, 0, .ColIndex("����")) = vbBlue
        .Cell(flexcpForeColor, 0, .ColIndex("����")) = vbBlue
        .Cell(flexcpForeColor, 0, .ColIndex("��λ")) = vbBlue
        
        .Redraw = flexRDDirect
    End With

End Sub

Private Sub cboDrugUnit_Click()
'���ٶ����������ݣ�ֱ�ӽ��滻��ˢ��
    Dim i As Long
    
    If Val(cboDrugUnit.Tag) = cboDrugUnit.ListIndex Or cboDrugUnit.Tag = "-1" Then Exit Sub
    
    With Me.vsfLimit
        .Redraw = flexRDNone
        For i = 1 To .Rows - 1
            '��ԭ���ۼ۵�λ
            Select Case Val(cboDrugUnit.Tag)
                Case 1  'סԺ��λ
                    .TextMatrix(i, .ColIndex("����")) = Val(.TextMatrix(i, .ColIndex("����"))) * Val(.TextMatrix(i, .ColIndex("סԺ��װ")))
                    .TextMatrix(i, .ColIndex("����")) = Val(.TextMatrix(i, .ColIndex("����"))) * Val(.TextMatrix(i, .ColIndex("סԺ��װ")))
                Case 2  '���ﵥλ
                    .TextMatrix(i, .ColIndex("����")) = Val(.TextMatrix(i, .ColIndex("����"))) * Val(.TextMatrix(i, .ColIndex("�����װ")))
                    .TextMatrix(i, .ColIndex("����")) = Val(.TextMatrix(i, .ColIndex("����"))) * Val(.TextMatrix(i, .ColIndex("�����װ")))
                Case 3  'ҩ�ⵥλ
                    .TextMatrix(i, .ColIndex("����")) = Val(.TextMatrix(i, .ColIndex("����"))) * Val(.TextMatrix(i, .ColIndex("ҩ���װ")))
                    .TextMatrix(i, .ColIndex("����")) = Val(.TextMatrix(i, .ColIndex("����"))) * Val(.TextMatrix(i, .ColIndex("ҩ���װ")))
            End Select
            
            '��ʼ����
            Select Case cboDrugUnit.ListIndex
                Case 0  '�ۼ۵�λ
                    .TextMatrix(i, .ColIndex("��λ")) = .TextMatrix(i, .ColIndex("�ۼ۵�λ"))
                    .TextMatrix(i, .ColIndex("��װ")) = 1
                    .TextMatrix(i, .ColIndex("���ۼ�")) = Format(.TextMatrix(i, .ColIndex("�̶����ۼ�")), "0.000")
                    .TextMatrix(i, .ColIndex("�������")) = Format(Val(.TextMatrix(i, .ColIndex("ʵ������"))), "0.00")
                Case 1  'סԺ��λ
                    .TextMatrix(i, .ColIndex("��λ")) = .TextMatrix(i, .ColIndex("סԺ��λ"))
                    .TextMatrix(i, .ColIndex("��װ")) = .TextMatrix(i, .ColIndex("סԺ��װ"))
                    .TextMatrix(i, .ColIndex("���ۼ�")) = Format(Val(.TextMatrix(i, .ColIndex("�̶����ۼ�"))) * Val(.TextMatrix(i, .ColIndex("סԺ��װ"))), "0.000")
                    .TextMatrix(i, .ColIndex("�������")) = Format(Val(.TextMatrix(i, .ColIndex("ʵ������"))) / Val(.TextMatrix(i, .ColIndex("סԺ��װ"))), "0.00")
                    .TextMatrix(i, .ColIndex("����")) = Format(Val(.TextMatrix(i, .ColIndex("����"))) / Val(.TextMatrix(i, .ColIndex("סԺ��װ"))), "0.00000")
                    .TextMatrix(i, .ColIndex("����")) = Format(Val(.TextMatrix(i, .ColIndex("����"))) / Val(.TextMatrix(i, .ColIndex("סԺ��װ"))), "0.00000")
                Case 2  '���ﵥλ
                    .TextMatrix(i, .ColIndex("��λ")) = .TextMatrix(i, .ColIndex("���ﵥλ"))
                    .TextMatrix(i, .ColIndex("��װ")) = .TextMatrix(i, .ColIndex("�����װ"))
                    .TextMatrix(i, .ColIndex("���ۼ�")) = Format(Val(.TextMatrix(i, .ColIndex("�̶����ۼ�"))) * Val(.TextMatrix(i, .ColIndex("�����װ"))), "0.000")
                    .TextMatrix(i, .ColIndex("�������")) = Format(Val(.TextMatrix(i, .ColIndex("ʵ������"))) / Val(.TextMatrix(i, .ColIndex("�����װ"))), "0.00")
                    .TextMatrix(i, .ColIndex("����")) = Format(Val(.TextMatrix(i, .ColIndex("����"))) / Val(.TextMatrix(i, .ColIndex("�����װ"))), "0.00000")
                    .TextMatrix(i, .ColIndex("����")) = Format(Val(.TextMatrix(i, .ColIndex("����"))) / Val(.TextMatrix(i, .ColIndex("�����װ"))), "0.00000")
                Case 3  'ҩ�ⵥλ
                    .TextMatrix(i, .ColIndex("��λ")) = .TextMatrix(i, .ColIndex("ҩ�ⵥλ"))
                    .TextMatrix(i, .ColIndex("��װ")) = .TextMatrix(i, .ColIndex("ҩ���װ"))
                    .TextMatrix(i, .ColIndex("���ۼ�")) = Format(Val(.TextMatrix(i, .ColIndex("�̶����ۼ�"))) * Val(.TextMatrix(i, .ColIndex("ҩ���װ"))), "0.000")
                    .TextMatrix(i, .ColIndex("�������")) = Format(Val(.TextMatrix(i, .ColIndex("ʵ������"))) / Val(.TextMatrix(i, .ColIndex("ҩ���װ"))), "0.00")
                    .TextMatrix(i, .ColIndex("����")) = Format(Val(.TextMatrix(i, .ColIndex("����"))) / Val(.TextMatrix(i, .ColIndex("ҩ���װ"))), "0.00000")
                    .TextMatrix(i, .ColIndex("����")) = Format(Val(.TextMatrix(i, .ColIndex("����"))) / Val(.TextMatrix(i, .ColIndex("ҩ���װ"))), "0.00000")
            End Select
        Next
        .Redraw = flexRDBuffered
    End With
    cboDrugUnit.Tag = cboDrugUnit.ListIndex
End Sub

Private Sub cboRoom_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim str�������� As String
    
    If Me.Tag = "5" Then
        str�������� = "I,M,K"
    ElseIf Me.Tag = "6" Then
        str�������� = "N,J,K"
    Else
        str�������� = "L,H,K"
    End If

    If KeyCode <> vbKeyReturn Then Exit Sub
    If cboRoom.ListCount = 0 Then Call zlControl.ControlSetFocus(vsfLimit): Exit Sub
    
    If cboRoom.ListIndex >= 0 Then
        If Val(cboRoom.Tag) = cboRoom.ItemData(cboRoom.ListIndex) Then
            Call zlControl.ControlSetFocus(vsfLimit, True)
            Exit Sub
        End If
    End If
    
    If Select����ѡ����(Me, cboRoom, Trim(cboRoom.Text), str��������, IIf(InStr(1, strPrivs, "�����������пⷿ�޶��̵�") = 0, True, False)) = False Then
        Exit Sub
    End If
    If cboRoom.ListIndex >= 0 Then
        cboRoom.Tag = cboRoom.ItemData(cboRoom.ListIndex)
    End If
End Sub

Private Sub cboRoom_KeyPress(KeyAscii As Integer)
    '�������뵥����
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub


Private Sub cboRoom_Validate(Cancel As Boolean)
    If cboRoom.ListCount > 0 Then
        If cboRoom.ListIndex = -1 Then
            MsgBox "��ѡ��һ��ҩ�����ҩ����", vbInformation, gstrSysName
            Cancel = True
        End If
    End If
End Sub


Private Sub cmdClear_Click()
    If Me.vsfLimit.Rows = 1 Then Exit Sub
    
    If MsgBox("������������ã��Ƿ�ȷ����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    With Me.vsfLimit
        .Redraw = flexRDNone
        
        .Cell(flexcpText, 1, 0, .Rows - 1, 0) = ""
        
        If InStr(1, strPrivs, "�����޿���") > 0 Then
            .Cell(flexcpText, 1, .ColIndex("����"), .Rows - 1, .ColIndex("����")) = Format(0, "0.00000")
            .Cell(flexcpText, 1, .ColIndex("����"), .Rows - 1, .ColIndex("����")) = Format(0, "0.00000")
        End If
        
        If InStr(1, strPrivs, "�̵���������") > 0 Then
            .Cell(flexcpText, 1, .ColIndex("����"), .Rows - 1, .ColIndex("����")) = ""
            .Cell(flexcpText, 1, .ColIndex("����"), .Rows - 1, .ColIndex("����")) = ""
            .Cell(flexcpText, 1, .ColIndex("����"), .Rows - 1, .ColIndex("����")) = ""
            .Cell(flexcpText, 1, .ColIndex("����"), .Rows - 1, .ColIndex("����")) = ""
        End If
        
        .Cell(flexcpText, 1, .ColIndex("��λ"), .Rows - 1, .ColIndex("��λ")) = ""
        
        .Redraw = flexRDBuffered
    End With

End Sub
Private Sub cmdClose_Click()
    Unload Me
    Exit Sub
End Sub

Private Sub cmdFilter_Click()
    If Me.vsfLimit.Rows = 1 Then Exit Sub
    
    If frmMediLimitFilter.GetCondition(Me, mlng�ⷿID, Me.Tag, mstr����, mstr����ID, mstr����) = True Then
        If mblnChanged = True Then
            If MsgBox("��ǰ�޸�δ���棬�Ƿ񰴹���������ȡ���ݣ�", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        End If
        Call zlLimitRef
    End If
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
End Sub


Private Sub cmdRestore_Click()
    If Me.vsfLimit.Rows = 1 Then Exit Sub
    If MsgBox("���ָ��������ã��Ƿ�ȷ����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    Call zlLimitRef
End Sub


Private Sub cmdSave_Click()
    Dim strMsgBox As String, strErrors As String
    Dim intNewStocks As Integer
    Dim intMaxStocks As Integer
    
    strErrors = ""
    
    If mblnChanged = False Then Exit Sub
    
    With Me.vsfLimit
        For mintCount = 1 To .Rows - 1
            If Val(.TextMatrix(mintCount, .ColIndex("����"))) <> 0 _
                And Val(.TextMatrix(mintCount, .ColIndex("����"))) < Val(.TextMatrix(mintCount, .ColIndex("����"))) Then
                .TextMatrix(mintCount, 0) = "��"
                strErrors = strErrors & vbCrLf & .TextMatrix(mintCount, .ColIndex("����")) & "-" & .TextMatrix(mintCount, .ColIndex("����"))
                strMsgBox = "��" & .TextMatrix(mintCount, .ColIndex("����")) & "-" & .TextMatrix(mintCount, .ColIndex("����")) & "���Ĵ������޴��ڴ������ޣ�" & _
                        vbCrLf & vbCrLf & "������������ҩƷ��"
                If MsgBox(strMsgBox, vbQuestion + vbYesNo, gstrSysName) = vbNo Then
                    Me.stbThis.Panels(2).Text = ""
                    .TopRow = mintCount: .Row = mintCount: .SetFocus: Exit Sub
                End If
            ElseIf Val(.RowData(mintCount)) <> 0 Then
                gstrSql = "zl_ҩƷ�����޶�_Update(" & Me.cboRoom.ItemData(Me.cboRoom.ListIndex)
                gstrSql = gstrSql & "," & .RowData(mintCount)
                gstrSql = gstrSql & "," & Format(Val(.TextMatrix(mintCount, .ColIndex("����"))) * Val(.TextMatrix(mintCount, .ColIndex("��װ"))), "0.00000")
                gstrSql = gstrSql & "," & Format(Val(.TextMatrix(mintCount, .ColIndex("����"))) * Val(.TextMatrix(mintCount, .ColIndex("��װ"))), "0.00000")
                gstrSql = gstrSql & ",'" & IIf(Trim(.TextMatrix(mintCount, .ColIndex("����"))) = "", "0", "1")
                gstrSql = gstrSql & IIf(Trim(.TextMatrix(mintCount, .ColIndex("����"))) = "", "0", "1")
                gstrSql = gstrSql & IIf(Trim(.TextMatrix(mintCount, .ColIndex("����"))) = "", "0", "1")
                gstrSql = gstrSql & IIf(Trim(.TextMatrix(mintCount, .ColIndex("����"))) = "", "0", "1")
                gstrSql = gstrSql & "','" & Trim(.TextMatrix(mintCount, .ColIndex("��λ"))) & "'"
                gstrSql = gstrSql & "," & IIf(Trim(.TextMatrix(mintCount, .ColIndex("��������"))) = "", "0", "1")
                gstrSql = gstrSql & ")"
                err = 0: On Error Resume Next
                Call zldatabase.ExecuteProcedure(gstrSql, Me.Caption)
                If err <> 0 Then
                    Call SaveErrLog
                    err = 0
                    .TextMatrix(mintCount, 0) = "��"
                    strErrors = strErrors & vbCrLf & .TextMatrix(mintCount, .ColIndex("����")) & "-" & .TextMatrix(mintCount, .ColIndex("����"))
                    strMsgBox = "���桰" & .TextMatrix(mintCount, .ColIndex("����")) & .TextMatrix(mintCount, .ColIndex("����")) & "��ʱ��������" & _
                            vbCrLf & vbCrLf & "������������ҩƷ��"
                    If MsgBox(strMsgBox, vbQuestion + vbYesNo, gstrSysName) = vbNo Then
                        Me.stbThis.Panels(2).Text = ""
                        .TopRow = mintCount: .Row = mintCount: .SetFocus: Exit Sub
                    End If
                End If
                If mintCount Mod IIf(.Rows > 20, .Rows \ 20, 1) = 0 Then
                    Me.stbThis.Panels(2).Text = "���ڱ��棺" & String(mintCount \ IIf(.Rows > 20, .Rows \ 20, 1), "��")
                End If
            End If
        Next
    End With
    Me.stbThis.Panels(2).Text = ""
    strMsgBox = "��" & Me.cboRoom.Text & "���������Ա�����ϣ�"
    If strErrors <> "" Then
        strMsgBox = strMsgBox & vbCrLf & "������ҩƷ�����������飺" & strErrors
    End If
    MsgBox strMsgBox, vbExclamation, gstrSysName

End Sub

Private Sub cmdApply_Click()
    Dim strValue As String
    
    With vsfLimit
        If .Rows = 1 Then Exit Sub
        
        Select Case .Col
            Case .ColIndex("����"), .ColIndex("����"), .ColIndex("����"), .ColIndex("����")
            Case .ColIndex("����"), .ColIndex("����")
                If InStr(1, strPrivs, "�����޿���") = 0 Then Exit Sub
            Case .ColIndex("��λ")
                If InStr(1, strPrivs, "������λ") = 0 Then Exit Sub
            Case Else
                Exit Sub
        End Select
        
        If MsgBox("��[" & .TextMatrix(0, .Col) & "]�е�����Ӧ�õ�����ҩƷ���Ƿ�ȷ����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
        '����ǰ�е�����Ӧ�õ�����ҩƷ��ͬ��
        strValue = .TextMatrix(.Row, .Col)
        .Cell(flexcpText, 1, .Col, .Rows - 1, .Col) = strValue
    End With
End Sub
Private Sub Form_Activate()
    If mblnActive = True Then Exit Sub
    
    If Me.cmdClose.Tag = "����" Then
        Me.cmdSave.Visible = False
        Me.cmdClear.Visible = False
        Me.cmdRestore.Visible = False
    End If
    lbl����.Visible = True
    txt����.Visible = True
    lblComment1.Visible = True
    
    err = 0: On Error GoTo ErrHand
    gstrSql = "select ID,����,����" & _
            "  from ���ű� D"
    If Me.Tag = "5" Then
        gstrSql = gstrSql & " where ID in (select distinct ����id from ��������˵�� where �������� like '��ҩ%' or ��������='�Ƽ���') and (d.����ʱ�� is null or to_char(d.����ʱ��,'yyyy-mm-dd')='3000-01-01')"
    ElseIf Me.Tag = "6" Then
        gstrSql = gstrSql & " where ID in (select distinct ����id from ��������˵�� where �������� like '��ҩ%' or ��������='�Ƽ���') and (d.����ʱ�� is null or to_char(d.����ʱ��,'yyyy-mm-dd')='3000-01-01')"
    Else
        gstrSql = gstrSql & " where ID in (select distinct ����id from ��������˵�� where �������� like '��ҩ%' or ��������='�Ƽ���') and (d.����ʱ�� is null or to_char(d.����ʱ��,'yyyy-mm-dd')='3000-01-01')"
    End If
    If InStr(1, strPrivs, "�����������пⷿ�޶��̵�") = 0 Then
        gstrSql = gstrSql & "      and ID in (select ����ID from ������Ա R where R.��ԱID=[1])"
    End If
    
    Set mrsNormal = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, UserInfo.ID)
        
    With mrsNormal
        Me.cboRoom.Clear
        Do While Not .EOF
            Me.cboRoom.AddItem !���� & "-" & !����
            Me.cboRoom.ItemData(Me.cboRoom.NewIndex) = !ID
            .MoveNext
        Loop
    End With
    If Me.cboRoom.ListCount <= 0 Then
        MsgBox "δ����" & IIf(Me.Tag = "5", "����ҩ", IIf(Me.Tag = "6", "�г�ҩ", "�в�ҩ")) & "�ⷿ���޷����ô�������", vbExclamation, gstrSysName
        Unload Me: Exit Sub
    End If
    Me.cboRoom.ListIndex = 0
    
    Call RestoreWinState(Me, App.ProductName, Me.Caption)
    
    mblnActive = True
    
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cboRoom_Click()
    err = 0: On Error GoTo ErrHand
    
    If mlng�ⷿID = cboRoom.ItemData(cboRoom.ListIndex) Then Exit Sub
    mlng�ⷿID = cboRoom.ItemData(cboRoom.ListIndex)
    cboRoom.Tag = GetDrugUnit(mlng�ⷿID)
    cboDrugUnit.Text = cboRoom.Tag
    mlngFind = 0
    mstr���� = "����"
    mstr����ID = ""
    mstr���� = ""
    Call zlLimitRef
    cboDrugUnit.Tag = cboDrugUnit.ListIndex
    Exit Sub
    
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'ȡҩƷ��λ����
Public Function GetDrugUnit(ByVal lng�ⷿID As Long) As String
    Dim rsProperty As New Recordset
    Dim strobjTemp As String                    '�����������ַ���
    Dim strWorkTemp As String                   '���湤�������ַ���
    Dim intUnit As Integer, strUnit As String
    Dim blnȱʡ As Boolean
    Dim lngModul As Long
    
    On Error GoTo ErrHand
    
    gstrSql = "SELECT distinct �������,�������� From ��������˵�� Where ����ID =[1]"
    Set rsProperty = zldatabase.OpenSQLRecord(gstrSql, "��ȡҩƷ��λ", lng�ⷿID)

    'ȡ������󼰲�������
    With rsProperty
        Do While Not .EOF
            strobjTemp = strobjTemp & .Fields(0)
            strWorkTemp = strWorkTemp & .Fields(1)
            .MoveNext
        Loop
        .Close
    End With
    If InStr(strWorkTemp, "ҩ��") <> 0 Then
        'ҩ�ⵥλ
        intUnit = 1
        strUnit = 4
    ElseIf InStr(strobjTemp, "1") <> 0 Or InStr(strobjTemp, "3") <> 0 Then
        '���ﵥλ
        intUnit = 2
        strUnit = 2
    ElseIf InStr(strobjTemp, "2") <> 0 Then
        'סԺ��λ
        intUnit = 3
        strUnit = 3
    Else
        '�ۼ۵�λ����Ҫ���Ƽ���
        intUnit = 4
        strUnit = 1
    End If
    
    'ȡ��ҩ��ȱʡ��ʹ�õĵ�λ
    GetDrugUnit = GetSpecUnit(lng�ⷿID, intUnit)
        
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
    GetDrugUnit = "�ۼ۵�λ"
End Function

'����ָ���ָⷿ�����÷�Χ�ĵ�λ
Public Function GetSpecUnit(ByVal lng�ⷿID As Long, ByVal int��Χ As Integer) As String
    Dim strobjTemp As String                    '�����������ַ���
    Dim strWorkTemp As String                   '���湤�������ַ���
    Dim strUnit As String
    Dim rsProperty As New ADODB.Recordset
    
    On Error GoTo ErrHand
    
    gstrSql = "Select Nvl(����,1) AS ��λ From ҩƷ�ⷿ��λ Where �ⷿID=[1] And ���÷�Χ=[2]"
    Set rsProperty = zldatabase.OpenSQLRecord(gstrSql, "��ȡ��λ", lng�ⷿID, int��Χ)

    If rsProperty.RecordCount = 1 Then
        strUnit = rsProperty!��λ
    Else
'        MsgBox "�ÿⷿδ���ÿⷿ��λ�����ݲ��������Լ��������ȡȱʡ��λ��" & _
'            vbCrLf & "ȱʡ��λ�Ĺ���" & _
'            vbCrLf & "  ���������סԺ�������סԺ�ģ�ȡסԺ��λ" & _
'            vbCrLf & "  ������������ģ�ȡ���ﵥλ" & _
'            vbCrLf & "  ����ҩ�����Եģ�ȡҩ�ⵥλ" & _
'            vbCrLf & "  ����ȡ�ۼ۵�λ", vbInformation, gstrSysName
        
        gstrSql = "SELECT distinct �������,�������� From ��������˵�� Where ����ID =[1]"
        Set rsProperty = zldatabase.OpenSQLRecord(gstrSql, "��ȡҩƷ��λ", lng�ⷿID)

        'ȡ������󼰲�������
        With rsProperty
            Do While Not .EOF
                strobjTemp = strobjTemp & .Fields(0)
                strWorkTemp = strWorkTemp & .Fields(1)
                .MoveNext
            Loop
            .Close
        End With
        If InStr(strobjTemp, "2") <> 0 Or InStr(strobjTemp, "3") <> 0 Then
            'סԺ��λ
            strUnit = 3
        ElseIf InStr(strobjTemp, "1") <> 0 Then
            '���ﵥλ
            strUnit = 2
        ElseIf InStr(strWorkTemp, "ҩ��") <> 0 Then
            'ҩ�ⵥλ
            strUnit = 4
        Else
            '�ۼ۵�λ����Ҫ���Ƽ���
            strUnit = 1
        End If
    End If
    
    'ת��Ϊ��ʵ�ĵ�λ���ظ�������
    GetSpecUnit = Switch(strUnit = 1, "�ۼ۵�λ", strUnit = 2, "���ﵥλ", strUnit = 3, "סԺ��λ", strUnit = 4, "ҩ�ⵥλ")
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function
Private Sub zlLimitRef()
    '--------------------------------------------------------
    '���ܣ�ˢ�¿���޶�
    '--------------------------------------------------------
    Dim rsFind As ADODB.Recordset
    Dim lngRow As Long
    Dim intǰ׺ As Integer
    Dim int��׺ As Integer
    Dim str����ǰ׺ As String
    Dim str�����׺ As String
    Dim blnLine As Boolean
    Dim lngCount As Long
    Dim strRule As String
    Dim lngId As Long
    
    err = 0: On Error GoTo ErrHand
    
    If mstr���� = "" And mstr����ID = "" Then
        strRule = ""
    Else '������������ʹ���Զ����oracle����
        strRule = "/*+ RULE*/"
    End If
    
    gstrSql = "Select " & strRule & " I.ID,I.����,I.����, i.��Ʒ��,I.���,I.����,I.ԭ����," & _
         Switch(Me.cboRoom.Tag = "�ۼ۵�λ", "I.���㵥λ as ��λ,1 as ��װ,", _
                Me.cboRoom.Tag = "ҩ�ⵥλ", "I.ҩ�ⵥλ as ��λ,nvl(I.ҩ���װ,1) as ��װ,", _
                Me.cboRoom.Tag = "���ﵥλ", "I.���ﵥλ as ��λ,nvl(I.�����װ,1) as ��װ,", _
                Me.cboRoom.Tag = "סԺ��λ", "I.סԺ��λ as ��λ,nvl(I.סԺ��װ,1) as ��װ,") & _
            "I.���㵥λ as �ۼ۵�λ, 1 as �ۼ۰�װ," & _
            "I.סԺ��λ, nvl(I.סԺ��װ,1) as סԺ��װ," & _
            "I.���ﵥλ, nvl(I.�����װ,1) as �����װ," & _
            "I.ҩ�ⵥλ, nvl(I.ҩ���װ,1) as ҩ���װ," & _
            "   nvl(L.����,0) as ����,nvl(L.����,0) as ����,L.�̵�����,L.�ⷿ��λ,l.���ñ�־,K.ʵ������," & _
            "   Decode(I.�Ƿ���, 0, P.�ּ�, Decode(Sign(K.ʵ������ - 1), -1, 0, K.ʵ�ʽ�� / K.ʵ������)) As ���ۼ� " & _
            " From (Select  I.ID,I.����,I.����, b.���� As ��Ʒ��,I.���,I.����,S.ԭ����,I.���㵥λ,S.���ﵥλ,S.�����װ, " & _
            "           S.סԺ��λ,S.סԺ��װ,S.ҩ�ⵥλ,S.ҩ���װ, I.�Ƿ���, S.ҩ��id " & _
            "       From �շ���ĿĿ¼ I, �շ���Ŀ���� B,ҩƷ��� S," & _
            "            (Select Distinct ������Ŀid From ����ִ�п��� Where ִ�п���id=[1]) E,(select distinct �շ�ϸĿid from �շ�ִ�п��� where ִ�п���id=[1]) F " & _
            "       Where i.Id = b.�շ�ϸĿid(+) And b.����(+) = 3 And I.Id=S.ҩƷid And S.ҩ��id=E.������Ŀid and I.���=[2] And i.id=f.�շ�ϸĿid " & _
            "            and (I.����ʱ�� is null or I.����ʱ��=to_date('3000-01-01','YYYY-MM-DD'))) I," & _
            "      (Select ҩƷid,����,����,�̵�����,�ⷿ��λ,���ñ�־ From ҩƷ�����޶� L Where �ⷿid=[1]) L," & _
            " (Select ҩƷid, Sum(ʵ������) As ʵ������, Sum(ʵ�ʽ��) As ʵ�ʽ�� From ҩƷ��� " & _
            "  Where ���� = 1 And �ⷿid = [1] Group By ҩƷid) K, �շѼ�Ŀ P "
            
    If mstr����ID <> "" Then
        gstrSql = gstrSql & ", ������ĿĿ¼ Z, ���Ʒ���Ŀ¼ M, Table(Cast(f_Num2list([3]) As zlTools.t_Numlist)) G "
    End If
    
    If mstr���� <> "" Then
        gstrSql = gstrSql & ", ҩƷ���� T, Table(Cast(f_Str2list([4]) As zlTools.t_strlist)) H "
    End If
    
    gstrSql = gstrSql & " Where I.ID = P.�շ�ϸĿid And I.Id=L.ҩƷid(+) And I.ID = K.ҩƷid(+) And (p.��ֹ���� Is Null Or Sysdate Between p.ִ������ And Nvl(p.��ֹ����, To_Date('3000-01-01', 'yyyy-MM-dd'))) " & _
            GetPriceClassString("P")
    
    If mstr����ID <> "" Then
        gstrSql = gstrSql & " And I.ҩ��id = Z.ID And Z.����id = M.ID And M.ID = G.Column_Value "
    End If
    
    If mstr���� <> "" Then
        gstrSql = gstrSql & " And I.ҩ��id = T.ҩ��id And T.ҩƷ���� = H.Column_Value "
    End If
    
    gstrSql = gstrSql & " Order By I.����"
    
    Set mrsNormal = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, Me.cboRoom.ItemData(Me.cboRoom.ListIndex), Me.Tag, mstr����ID, mstr����)
    
    If Not mrsNormal.EOF Then lngCount = mrsNormal.RecordCount
    With mrsNormal
        Me.vsfLimit.Rows = 1
        Me.vsfLimit.Redraw = False
        Call IniGrid
        Do While Not .EOF
            If lngId <> mrsNormal!ID Then
                lngId = mrsNormal!ID
                Me.vsfLimit.Rows = vsfLimit.Rows + 1
                Me.vsfLimit.RowData(vsfLimit.Rows - 1) = Val(!ID)
                Me.vsfLimit.TextMatrix(vsfLimit.Rows - 1, Me.vsfLimit.ColIndex("����")) = !����
            
                If mintҩƷ������ʾ = 0 Then
                    Me.vsfLimit.TextMatrix(vsfLimit.Rows - 1, Me.vsfLimit.ColIndex("����")) = !����
                ElseIf mintҩƷ������ʾ = 1 Then
                    Me.vsfLimit.TextMatrix(vsfLimit.Rows - 1, Me.vsfLimit.ColIndex("����")) = IIf(IsNull(!��Ʒ��), !����, !��Ʒ��)
                Else
                    Me.vsfLimit.TextMatrix(vsfLimit.Rows - 1, Me.vsfLimit.ColIndex("����")) = !����
                    Me.vsfLimit.TextMatrix(vsfLimit.Rows - 1, Me.vsfLimit.ColIndex("��Ʒ��")) = IIf(IsNull(!��Ʒ��), "", !��Ʒ��)
                End If
                
                Me.vsfLimit.TextMatrix(vsfLimit.Rows - 1, Me.vsfLimit.ColIndex("���")) = IIf(IsNull(!���), "", !���)
                Me.vsfLimit.TextMatrix(vsfLimit.Rows - 1, Me.vsfLimit.ColIndex("����")) = IIf(IsNull(!����), "", !����)
                Me.vsfLimit.TextMatrix(vsfLimit.Rows - 1, Me.vsfLimit.ColIndex("ԭ����")) = IIf(IsNull(!ԭ����), "", !ԭ����)
                Me.vsfLimit.TextMatrix(vsfLimit.Rows - 1, Me.vsfLimit.ColIndex("��λ")) = IIf(IsNull(!��λ), "", !��λ)
                Me.vsfLimit.TextMatrix(vsfLimit.Rows - 1, Me.vsfLimit.ColIndex("��װ")) = !��װ
                Me.vsfLimit.TextMatrix(vsfLimit.Rows - 1, Me.vsfLimit.ColIndex("���ۼ�")) = IIf(!���ۼ� = 0, "", Format(!���ۼ� * !��װ, "0.000"))
                Me.vsfLimit.TextMatrix(vsfLimit.Rows - 1, Me.vsfLimit.ColIndex("�������")) = Format(!ʵ������ / !��װ, "0.00")
                Me.vsfLimit.TextMatrix(vsfLimit.Rows - 1, Me.vsfLimit.ColIndex("����")) = Format(!���� / !��װ, "0.00000")
                Me.vsfLimit.TextMatrix(vsfLimit.Rows - 1, Me.vsfLimit.ColIndex("����")) = Format(!���� / !��װ, "0.00000")
                Me.vsfLimit.TextMatrix(vsfLimit.Rows - 1, Me.vsfLimit.ColIndex("����")) = IIf(Mid(!�̵�����, 1, 1) = "1", "��", "")
                Me.vsfLimit.TextMatrix(vsfLimit.Rows - 1, Me.vsfLimit.ColIndex("����")) = IIf(Mid(!�̵�����, 2, 1) = "1", "��", "")
                Me.vsfLimit.TextMatrix(vsfLimit.Rows - 1, Me.vsfLimit.ColIndex("����")) = IIf(Mid(!�̵�����, 3, 1) = "1", "��", "")
                Me.vsfLimit.TextMatrix(vsfLimit.Rows - 1, Me.vsfLimit.ColIndex("����")) = IIf(Mid(!�̵�����, 4, 1) = "1", "��", "")
                Me.vsfLimit.TextMatrix(vsfLimit.Rows - 1, Me.vsfLimit.ColIndex("��λ")) = IIf(IsNull(!�ⷿ��λ), "", !�ⷿ��λ)
                Me.vsfLimit.TextMatrix(vsfLimit.Rows - 1, Me.vsfLimit.ColIndex("��������")) = IIf(IsNull(!���ñ�־), "��", IIf(!���ñ�־ = 0, "", "��"))
                
                Me.vsfLimit.TextMatrix(vsfLimit.Rows - 1, Me.vsfLimit.ColIndex("�ۼ۵�λ")) = IIf(IsNull(!�ۼ۵�λ), "", !�ۼ۵�λ)
                Me.vsfLimit.TextMatrix(vsfLimit.Rows - 1, Me.vsfLimit.ColIndex("סԺ��λ")) = IIf(IsNull(!סԺ��λ), "", !סԺ��λ)
                Me.vsfLimit.TextMatrix(vsfLimit.Rows - 1, Me.vsfLimit.ColIndex("סԺ��װ")) = IIf(IsNull(!סԺ��װ), "", !סԺ��װ)
                Me.vsfLimit.TextMatrix(vsfLimit.Rows - 1, Me.vsfLimit.ColIndex("���ﵥλ")) = IIf(IsNull(!���ﵥλ), "", !���ﵥλ)
                Me.vsfLimit.TextMatrix(vsfLimit.Rows - 1, Me.vsfLimit.ColIndex("�����װ")) = IIf(IsNull(!�����װ), "", !�����װ)
                Me.vsfLimit.TextMatrix(vsfLimit.Rows - 1, Me.vsfLimit.ColIndex("ҩ�ⵥλ")) = IIf(IsNull(!ҩ�ⵥλ), "", !ҩ�ⵥλ)
                Me.vsfLimit.TextMatrix(vsfLimit.Rows - 1, Me.vsfLimit.ColIndex("ҩ���װ")) = IIf(IsNull(!ҩ���װ), "", !ҩ���װ)
                Me.vsfLimit.TextMatrix(vsfLimit.Rows - 1, Me.vsfLimit.ColIndex("�̶����ۼ�")) = IIf(IsNull(!���ۼ�), 0, !���ۼ�)
                Me.vsfLimit.TextMatrix(vsfLimit.Rows - 1, Me.vsfLimit.ColIndex("ʵ������")) = IIf(IsNull(!ʵ������), 0, !ʵ������)
                
                If InStr(!����, "-") > 0 Then
                    blnLine = True
                    If Len(Mid(!����, 1, InStr(!����, "-") - 1)) > intǰ׺ Then
                        intǰ׺ = Len(Mid(!����, 1, InStr(!����, "-") - 1))
                    End If
                    
                    If Len(Mid(!����, InStr(!����, "-") + 1)) > int��׺ Then
                        int��׺ = Len(Mid(!����, InStr(!����, "-") + 1))
                    End If
                Else
                    If Len(!����) > intǰ׺ Then
                        intǰ׺ = Len(!����)
                    End If
                End If
                
                If vsfLimit.Rows - 1 Mod IIf(.RecordCount > 20, .RecordCount \ 20, 1) = 0 Then
                    Me.stbThis.Panels(2).Text = "������ȡ��" & String(vsfLimit.Rows - 1 \ IIf(.RecordCount > 20, .RecordCount \ 20, 1), "��")
                End If
            End If
            .MoveNext
        Loop
        
        For lngRow = 1 To Me.vsfLimit.Rows - 1
            If blnLine = False Then
                Me.vsfLimit.TextMatrix(lngRow, Me.vsfLimit.ColIndex("�������")) = Format(Me.vsfLimit.TextMatrix(lngRow, Me.vsfLimit.ColIndex("����")), String(intǰ׺, "0"))
            Else
                If InStr(Me.vsfLimit.TextMatrix(lngRow, Me.vsfLimit.ColIndex("����")), "-") > 0 Then
                    str����ǰ׺ = Mid(Me.vsfLimit.TextMatrix(lngRow, Me.vsfLimit.ColIndex("����")), 1, InStr(Me.vsfLimit.TextMatrix(lngRow, Me.vsfLimit.ColIndex("����")), "-") - 1)
                    str�����׺ = Mid(Me.vsfLimit.TextMatrix(lngRow, Me.vsfLimit.ColIndex("����")), InStr(Me.vsfLimit.TextMatrix(lngRow, Me.vsfLimit.ColIndex("����")), "-") + 1)
                    
                    str����ǰ׺ = Format(str����ǰ׺, String(intǰ׺, "0"))
                    str�����׺ = Format(str�����׺, String(int��׺, "0"))
                Else
                    str����ǰ׺ = Format(Me.vsfLimit.TextMatrix(lngRow, Me.vsfLimit.ColIndex("����")), String(intǰ׺, "0"))
                    str�����׺ = String(int��׺, "0")
                End If
                
                Me.vsfLimit.TextMatrix(lngRow, Me.vsfLimit.ColIndex("�������")) = str����ǰ׺ & "-" & str�����׺
            End If
        Next
        
        Me.vsfLimit.Col = Me.vsfLimit.ColIndex("�������")
        Me.vsfLimit.Sort = flexSortStringAscending
        
        If Me.vsfLimit.Rows > 1 Then
            Me.vsfLimit.Cell(flexcpBackColor, 1, Me.vsfLimit.ColIndex("����"), Me.vsfLimit.Rows - 1, Me.vsfLimit.ColIndex("�������")) = &HEFEFEF
        End If
        
        Me.vsfLimit.Redraw = True
    End With
    Me.stbThis.Panels(2).Text = "����" & lngCount & "��ҩƷ���" & " " & " ��ǰ���ࣺ" & IIf(mstr���� = "", "����", mstr����) & "  ��ǰ���ͣ�" & IIf(mstr���� = "", "����", mstr����)
    mblnChanged = False
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function IsHaveStock(ByVal strStockName As String) As Boolean
    Dim rs As ADODB.Recordset
    On Error GoTo errHandle
    gstrSql = "Select ���� From ҩƷ�ⷿ��λ where ����=[1]"
    Set rs = zldatabase.OpenSQLRecord(gstrSql, "�ж��Ƿ����ҩƷ�ⷿ��λ", strStockName)
        
    IsHaveStock = (rs.RecordCount > 0)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If txt����.Visible And KeyCode = vbKeyF3 Then
        Call txt����_KeyDown(vbKeyReturn, 0)
    End If
End Sub

Private Sub Form_Load()
    mintҩƷ������ʾ = Val(zldatabase.GetPara("ҩƷ������ʾ", , , 2))
    Call RestoreWinState(Me, App.ProductName)
    mlng�ⷿID = 0
    cboDrugUnit.Tag = "-1"
End Sub

Private Sub Form_Resize()
    err = 0: On Error Resume Next
    Me.fraLine.Left = 0: Me.fraLine.Width = Me.ScaleWidth + 100
    Me.vsfLimit.Left = 0: Me.vsfLimit.Width = Me.ScaleWidth
    Me.vsfLimit.Height = Me.ScaleHeight - Me.vsfLimit.Top - Me.fraFunc.Height - Me.stbThis.Height
    Me.fraFunc.Left = 0: Me.fraFunc.Width = Me.ScaleWidth: Me.fraFunc.Top = Me.vsfLimit.Top + Me.vsfLimit.Height
    Me.cmdClose.Left = Me.fraFunc.Width - Me.cmdClose.Width - 90
    Me.cmdSave.Left = Me.cmdClose.Left - Me.cmdSave.Width
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName, Me.Caption)
    
    mblnActive = False
End Sub




Private Function CheckNode(ByVal Node As Object, blnCheck As Boolean)
    Dim intIdx As Integer

    If Node.Children > 0 Then
        Set Node = Node.Child
        Do While Not Node Is Nothing
            Node.Checked = blnCheck
            If Node.Children > 0 Then
                CheckNode Node, blnCheck
            End If
            Set Node = Node.Next
        Loop
    Else
        Node.Checked = blnCheck
    End If
End Function

Private Sub SetParentNode(ByVal objMyTreeView As TreeView, ByVal Node As MSComctlLib.Node, blnCheck As Boolean)
    Dim intIdx As Integer
    
    If Not Node.Parent Is Nothing Then
        If blnCheck = True Then
            '���Ƿ������ֵܽӵ��Ƿ�Ҳȫ��TRUE�����ǣ������丸�ڵ�ҲΪTRUE�����򣬲���
            intIdx = Node.FirstSibling.Index
            Do While intIdx <> Node.LastSibling.Index
                If objMyTreeView.Nodes(intIdx).Checked = False Then
                    Node.Parent.Checked = False
                    Exit Do
                End If
                intIdx = objMyTreeView.Nodes(intIdx).Next.Index
            Loop
            If intIdx = Node.LastSibling.Index Then
                If objMyTreeView.Nodes(intIdx).Checked = True Then
                    Node.Parent.Checked = True
                End If
            End If
        Else
            Node.Parent.Checked = False
        End If
        
        Set Node = Node.Parent
        If Not Node Is Nothing Then
            SetParentNode objMyTreeView, Node, blnCheck
        End If
    End If
End Sub





Private Sub txt����_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strInput As String
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    strInput = Trim(UCase(txt����.Text))
    If strInput = "" Then Exit Sub
    
    Call FindGridRow(strInput)
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub vsfLimit_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With vsfLimit
        Select Case Col
            Case .ColIndex("��λ")
                .ColComboList(.ColIndex("��λ")) = "..."
        End Select
    End With
End Sub


Private Sub vsfLimit_AfterSort(ByVal Col As Long, Order As Integer)
    With vsfLimit
        If Col = .ColIndex("����") Then
            .Col = .ColIndex("�������")
            .Sort = Order
        End If
    End With
End Sub

Private Sub vsfLimit_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    With vsfLimit
        Select Case Col
            Case .ColIndex("��λ")
                mlngRow = Row
                If Select��λ("") = False Then
                    Exit Sub
                End If
            Case Else
        End Select
    End With
End Sub

Private Sub vsfLimit_CellChanged(ByVal Row As Long, ByVal Col As Long)
    mblnChanged = True
End Sub

Private Sub vsfLimit_DblClick()
    Dim blnNext As Boolean
    
    With vsfLimit
        If .Row < 1 Then Exit Sub
        If .MouseRow <> .Row Or .MouseCol <> .Col Then Exit Sub
        
        If .Col = .ColIndex("��������") Then
            If .TextMatrix(.Row, .Col) = "��" Then
                .TextMatrix(.Row, .Col) = ""
            Else
                .TextMatrix(.Row, .Col) = "��"
            End If
        End If
        If .Col = .ColIndex("����") Or .Col = .ColIndex("����") Or .Col = .ColIndex("����") Or .Col = .ColIndex("����") Then
            If InStr(1, strPrivs, "�̵���������") = 0 Then Exit Sub
            blnNext = True
        End If
        
        If blnNext = True Then
            If .TextMatrix(.Row, .Col) = "��" Then
                .TextMatrix(.Row, .Col) = ""
            Else
                .TextMatrix(.Row, .Col) = "��"
            End If
        End If
    End With
End Sub

Private Sub vsfLimit_EnterCell()
    Dim intRow As Integer
    
    With vsfLimit
        If .Row < 1 Then Exit Sub
        .FocusRect = flexFocusLight
        .Editable = flexEDNone
        Select Case .Col
            Case .ColIndex("����"), .ColIndex("����"), .ColIndex("����"), .ColIndex("����"), .ColIndex("��������")
                .FocusRect = flexFocusSolid
            Case .ColIndex("����"), .ColIndex("����")
                If InStr(1, strPrivs, "�����޿���") > 0 Then
                    .Editable = flexEDKbdMouse
                    .FocusRect = flexFocusSolid
                End If
            Case .ColIndex("��λ")
                If InStr(1, strPrivs, "������λ") > 0 Then
                    .Editable = flexEDKbdMouse
                    .FocusRect = flexFocusSolid
                End If
        End Select
        
        '������ѡ�б߿�
        If .Rows <> 1 Then
            For intRow = 0 To .Rows - 1
                .CellBorderRange intRow, 0, intRow, .Cols - 1, mlngNoneBorderColor, 0, 0, 0, 0, 0, 0
            Next
            
            .CellBorderRange .Row, .ColIndex("����"), .Row, .ColIndex("��λ"), mlngBorderColor, 0, 2, 0, 2, 0, 2
            .CellBorderRange .Row, .ColIndex("����"), .Row, .ColIndex("����"), mlngBorderColor, 2, 2, 0, 2, 0, 0
            .CellBorderRange .Row, .ColIndex("��λ"), .Row, .ColIndex("��λ"), mlngBorderColor, 0, 2, 2, 2, 0, 0
        End If
    End With
End Sub


Private Sub vsfLimit_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        vsfStore.Visible = False
    End If
    If vsfLimit.Col = vsfLimit.ColIndex("��λ") Then
        If KeyCode <> vbKeyReturn Then
            vsfLimit.ColComboList(vsfLimit.ColIndex("��λ")) = ""
        End If
        
        If KeyCode = vbKeyDelete Then
            vsfLimit.TextMatrix(vsfLimit.Row, vsfLimit.Col) = ""
        End If
    End If
    
    If txt����.Visible And KeyCode = vbKeyF3 Then
        Call txt����_KeyDown(vbKeyReturn, 0)
    End If
End Sub

Private Sub vsfLimit_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    Dim lvwItem As ListItem
    Dim rsTemp As New ADODB.Recordset
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    With vsfLimit
        If Trim(.EditText) = "" Then Exit Sub
        
        If Col = .ColIndex("��λ") Then
            If LenB(StrConv(.EditText, vbFromUnicode)) > 50 Then
'                MsgBox "��λ���������50����ĸ��25������", vbInformation, gstrSysName
'                vsfLimit.TextMatrix(Row, Col) = ""
'                vsfLimit.EditText = vsfLimit.TextMatrix(Row, Col)
                Exit Sub
            End If
        ElseIf Col = .ColIndex("����") Or Col = .ColIndex("����") Then
            If Not IsNumeric(.EditText) Then
                KeyCode = 0
                Exit Sub
            End If
            If Val(.EditText) < 0 Then
                KeyCode = 0
                Exit Sub
            End If
            If Val(.EditText) > 10000000000000# Then
                KeyCode = 0
                Exit Sub
            End If
        End If

       Select Case Col
            Case .ColIndex("����")
                .EditText = Format(.EditText, "0.00000"): .TextMatrix(Row, .ColIndex("����")) = .EditText
            Case .ColIndex("����")
                .EditText = Format(.EditText, "0.00000"): .TextMatrix(Row, .ColIndex("����")) = .EditText
            Case .ColIndex("��λ")
                If Select��λ(.EditText) = False Then
                    vsfLimit.TextMatrix(Row, Col) = vsfLimit.EditText
                    vsfLimit.Cell(flexcpForeColor, Row, Col) = vbRed
'                    If MsgBox("û���ҵ��û�λ���Ƿ����Ӹû�λ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
'                        vsfLimit.TextMatrix(Row, Col) = ""
'                    End If
                    Exit Sub
                End If
                vsfLimit.EditText = vsfLimit.TextMatrix(Row, Col)
        End Select
    End With
End Sub

Private Function Select��λ(ByVal strKey As String) As Boolean
    Dim rsTemp As ADODB.Recordset
    Dim blnCancel As Boolean
    Dim strID As String
    Dim str���� As String
    Dim objNode As Node
    Dim str��λ As String
    
    err = 0: On Error GoTo ErrHand:
    
    strKey = UCase(strKey)

    If strKey <> "" Then
        gstrSql = " Select id,����,����,���� From ҩƷ�ⷿ��λ " & _
                " Where �ⷿid=[1] And(���� Like [2] Or ���� Like [3] Or ���� Like [3]) Order By ����"
    Else
        gstrSql = " Select id,����,����,���� From ҩƷ�ⷿ��λ " & _
                " Where �ⷿid=[1]"
    End If
    
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, "��λ", mlng�ⷿID, strKey & "%", gstrMatch & strKey & "%")
    
    If rsTemp.EOF Then
        vsfLimit.EditText = strKey
        Exit Function
    End If
    
    str��λ = vsfLimit.TextMatrix(vsfLimit.Row, vsfLimit.ColIndex("��λ"))
    vsfStore.Rows = 1
    Do While Not rsTemp.EOF
        With vsfStore
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, .ColIndex("id")) = rsTemp!ID
            .TextMatrix(.Rows - 1, .ColIndex("����")) = rsTemp!����
            .TextMatrix(.Rows - 1, .ColIndex("��λ")) = rsTemp!����
            
            If str��λ <> "" Then
                If InStr(1, "," & str��λ & ",", "," & rsTemp!���� & ",") > 0 Then
                    .TextMatrix(.Rows - 1, .ColIndex("ѡ��")) = 1
                End If
            End If
        End With
        rsTemp.MoveNext
    Loop
    If rsTemp.RecordCount > 0 Then
        vsfStore.Move vsfLimit.CellLeft + 30, vsfLimit.CellTop + vsfStore.Height - 200
        vsfStore.Visible = True
        If strKey = "" Then
            vsfStore.SetFocus
        End If
        vsfStore.Row = 1
    End If
    
    Select��λ = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub vsfLimit_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If InStr(GCST_INVALIDCHAR, Chr(KeyAscii)) > 0 Then KeyAscii = 0
    
    Select Case Col
        Case vsfLimit.ColIndex("����"), vsfLimit.ColIndex("����")
            If InStr("1234567890" + Chr(46) + Chr(8) + Chr(13), Chr(KeyAscii)) = 0 Then
                KeyAscii = 0
            ElseIf KeyAscii = Asc(".") Then
                If InStr(vsfLimit.EditText, ".") <> 0 Then     'ֻ�ܴ���һ��С����
                    KeyAscii = 0
                End If
            End If
    End Select
    
End Sub

Private Sub vsfLimit_RowColChange()
    With vsfLimit()
        .Cell(flexcpText, 0, 0, .Rows - 1, 0) = ""
        If .Row > 0 Then
            .Cell(flexcpFontName, , 0) = "Marlett"
            .TextMatrix(.Row, 0) = 4
        End If
    End With
End Sub

Private Sub vsfLimit_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim lvwItem As ListItem
    Dim rsTemp As New ADODB.Recordset
    
    With vsfLimit
        If Trim(.EditText) = "" Then Exit Sub
        
        If Col = .ColIndex("��λ") Then
            If LenB(StrConv(.EditText, vbFromUnicode)) > 50 Then
                MsgBox "��λ���������50����ĸ��25������", vbInformation, gstrSysName
                vsfLimit.TextMatrix(Row, Col) = ""
                vsfLimit.EditText = vsfLimit.TextMatrix(Row, Col)
                Exit Sub
            End If
        End If

       Select Case Col
            Case .ColIndex("��λ")
                If Select��λ(.EditText) = False Then
                    vsfLimit.TextMatrix(Row, Col) = vsfLimit.EditText
                    vsfLimit.Cell(flexcpForeColor, Row, Col) = vbRed
                    If MsgBox("û���ҵ��û�λ���Ƿ����Ӹû�λ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        vsfLimit.TextMatrix(Row, Col) = ""
                        vsfLimit.EditText = vsfLimit.TextMatrix(Row, Col)
                    Else
                        gstrSql = "Zl_ҩƷ�ⷿ��λ_Insert(Null, '" & Trim(vsfLimit.TextMatrix(vsfLimit.Row, vsfLimit.ColIndex("��λ"))) & "', Null, " & Me.cboRoom.ItemData(Me.cboRoom.ListIndex) & ", Null)"
                        Call zldatabase.ExecuteProcedure(gstrSql, Me.Caption)
                    End If
                    Exit Sub
                End If
                vsfLimit.EditText = vsfLimit.TextMatrix(Row, Col)
        End Select
    End With
End Sub

Private Sub vsfStore_Click()
    With vsfStore
        If .Col = .ColIndex("ѡ��") Then
            If Val(.TextMatrix(.Row, .ColIndex("ѡ��"))) = 1 Then
                .TextMatrix(.Row, .ColIndex("ѡ��")) = ""
            Else
                .TextMatrix(.Row, .ColIndex("ѡ��")) = "1"
            End If
        End If
    End With
End Sub


Private Sub vsfStore_DblClick()
    Dim i As Integer
    Dim str��λ As String
    
    With vsfStore
        If .Rows <= 1 Then Exit Sub
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, .ColIndex("ѡ��"))) = 1 Then
                str��λ = str��λ & "," & .TextMatrix(i, .ColIndex("��λ"))
            End If
        Next
        vsfStore.Visible = False
    End With
    
    If str��λ <> "" Then
        str��λ = Mid(str��λ, 2)
    End If
    With vsfLimit
        .Redraw = flexRDNone
        .TextMatrix(.Row, .ColIndex("��λ")) = str��λ
        vsfLimit.Cell(flexcpForeColor, .Row, .ColIndex("��λ"), .Row, .ColIndex("��λ")) = vbBlack
        .Redraw = flexRDBuffered
    End With
End Sub


Private Sub vsfStore_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        vsfStore.Visible = False
    End If
End Sub


Private Sub vsfStore_LostFocus()
    If vsfStore.Visible = True Then
        vsfStore.Visible = False
    End If
End Sub


