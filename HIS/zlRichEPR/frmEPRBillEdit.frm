VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmEPRBillEdit 
   BorderStyle     =   0  'None
   Caption         =   "���Ƶ��ݱ༭"
   ClientHeight    =   4350
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6285
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   6285
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VSFlex8Ctl.VSFlexGrid vfg���� 
      Height          =   1230
      Left            =   150
      TabIndex        =   8
      Top             =   1440
      Width           =   6000
      _cx             =   10583
      _cy             =   2170
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
      BackColorSel    =   16635590
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
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
      Rows            =   4
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   ""
      ScrollTrack     =   -1  'True
      ScrollBars      =   2
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
   Begin VB.PictureBox picEdit 
      Align           =   1  'Align Top
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   3795
      Left            =   0
      ScaleHeight     =   3795
      ScaleWidth      =   6285
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   6285
      Begin VB.PictureBox picApply 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   1850
         Left            =   120
         ScaleHeight     =   1845
         ScaleWidth      =   6135
         TabIndex        =   27
         Top             =   1200
         Visible         =   0   'False
         Width           =   6135
         Begin VSFlex8Ctl.VSFlexGrid vsFile 
            Height          =   1470
            Left            =   35
            TabIndex        =   29
            Top             =   245
            Width           =   6000
            _cx             =   10583
            _cy             =   2593
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
            BackColorSel    =   16635590
            ForeColorSel    =   -2147483640
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483633
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
            Rows            =   4
            Cols            =   3
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   360
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   ""
            ScrollTrack     =   -1  'True
            ScrollBars      =   2
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
         Begin VB.Label lblApplysMark 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "�Զ������뵥: (��Ҫ���ö�Ӧ��xsl��ʽ�ļ���xml�����ļ���html��ʾ�ļ�)"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   30
            TabIndex        =   28
            Top             =   0
            Width           =   6120
         End
      End
      Begin VB.PictureBox picApplyType 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         ScaleHeight     =   285
         ScaleWidth      =   4035
         TabIndex        =   24
         Top             =   888
         Width           =   4035
         Begin VB.OptionButton optApply 
            Caption         =   "���븽��ģʽ"
            Height          =   210
            Index           =   0
            Left            =   45
            TabIndex        =   26
            Top             =   45
            Value           =   -1  'True
            Width           =   1530
         End
         Begin VB.OptionButton optApply 
            Caption         =   "�Զ������뵥"
            Height          =   210
            Index           =   1
            Left            =   1680
            TabIndex        =   25
            Top             =   45
            Width           =   1770
         End
      End
      Begin VB.ComboBox cbxSubClass 
         Height          =   300
         Left            =   4755
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   135
         Width           =   1395
      End
      Begin VB.Frame frasplit 
         BackColor       =   &H00FFC0C0&
         Height          =   30
         Left            =   105
         TabIndex        =   21
         Top             =   3405
         Width           =   6120
      End
      Begin VB.PictureBox picEditType 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2100
         ScaleHeight     =   285
         ScaleWidth      =   4035
         TabIndex        =   17
         Top             =   3090
         Width           =   4035
         Begin VB.OptionButton optEditType 
            Caption         =   "������༭��"
            Height          =   210
            Index           =   1
            Left            =   2160
            TabIndex        =   19
            Top             =   45
            Width           =   1770
         End
         Begin VB.OptionButton optEditType 
            Caption         =   "ȫ�Ĳ����༭��"
            Height          =   210
            Index           =   0
            Left            =   45
            TabIndex        =   18
            Top             =   45
            Width           =   1770
         End
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "����һ��(&N)"
         Height          =   350
         Index           =   3
         Left            =   4905
         TabIndex        =   12
         Top             =   2685
         Width           =   1245
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "ɾ����(&D)"
         Height          =   350
         Index           =   1
         Left            =   1395
         TabIndex        =   10
         Top             =   2685
         Width           =   1245
      End
      Begin VB.TextBox txt���� 
         Height          =   300
         Left            =   1995
         MaxLength       =   60
         TabIndex        =   4
         Top             =   135
         Width           =   2115
      End
      Begin VB.TextBox txt��� 
         Height          =   300
         Left            =   600
         MaxLength       =   13
         TabIndex        =   2
         Top             =   135
         Width           =   780
      End
      Begin VB.TextBox txt˵�� 
         Height          =   300
         Left            =   600
         MaxLength       =   60
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   540
         Width           =   5550
      End
      Begin VB.CheckBox chk���� 
         Caption         =   "ִ�к���ִ�б���:"
         Height          =   195
         Left            =   150
         TabIndex        =   13
         Top             =   3135
         Width           =   1845
      End
      Begin VB.OptionButton opt���� 
         Caption         =   "���༭��ʽ��ӡ(&1)"
         Height          =   180
         Index           =   0
         Left            =   2145
         TabIndex        =   14
         Top             =   3525
         Width           =   1890
      End
      Begin VB.OptionButton opt���� 
         Caption         =   "�Զ��屨���ӡ(&2)"
         Height          =   180
         Index           =   1
         Left            =   4260
         TabIndex        =   15
         Top             =   3525
         Width           =   1890
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "�����(&A)"
         Height          =   350
         Index           =   0
         Left            =   150
         TabIndex        =   9
         Top             =   2685
         Width           =   1245
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "����һ��(&U)"
         Height          =   350
         Index           =   2
         Left            =   3660
         TabIndex        =   11
         Top             =   2685
         Width           =   1245
      End
      Begin MSComDlg.CommonDialog cdgFile 
         Left            =   0
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   4320
         TabIndex        =   22
         Top             =   195
         Width           =   360
      End
      Begin VB.Label Label1 
         Caption         =   "ȫ�Ĳ����༭��ӡ��ʽ"
         Height          =   255
         Left            =   150
         TabIndex        =   20
         Top             =   3495
         Width           =   1845
      End
      Begin VB.Label lbl���� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   1545
         TabIndex        =   3
         Top             =   195
         Width           =   360
      End
      Begin VB.Label lbl��� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   150
         TabIndex        =   1
         Top             =   195
         Width           =   360
      End
      Begin VB.Label lbl˵�� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "˵��"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   150
         TabIndex        =   5
         Top             =   600
         Width           =   360
      End
      Begin VB.Label lbl���� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "���븽��: (���ٴ�ҽ������ʱ������д�ĸ�������)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   150
         TabIndex        =   7
         Top             =   1215
         Width           =   4140
      End
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��������Ͳ����Զ��屨������ı��棬��Ҫ����Ա���Զ��屨�����У��Ա���'ZLCISBILL00���-?'�����޸���Ƶ�����"
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   570
      TabIndex        =   16
      Top             =   3870
      Width           =   5610
      WordWrap        =   -1  'True
   End
   Begin VB.Image imgNote 
      Height          =   240
      Left            =   150
      Picture         =   "frmEPRBillEdit.frx":0000
      Top             =   3840
      Width           =   240
   End
End
Attribute VB_Name = "frmEPRBillEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum mCol
    ��Ŀ = 0: ����: ֻ��: Ҫ��id: ����
End Enum
Private mlngBillID As Long          '��ǰ��ʾ����Ŀid
Private mstrCombos As String        'Ҫ���б�
Private mrsItems As New Recordset
Private mobjFile As New FileSystemObject     '�ļ���������
Private arrSQL() As String
Private arrSQLFile() As String
Private mblndb As Boolean
'--------------------------------------------
'����Ϊ���幫������
'--------------------------------------------

Public Function zlRefresh(lngBillId As Long) As Boolean
    '���ܣ�������Ŀidˢ�µ�ǰ��ʾ����
Dim rsTemp As New ADODB.Recordset
Dim i As Long
Dim lngCount As Long
    
    mlngBillID = lngBillId
    ReDim arrSQL(0): ReDim arrSQLFile(0)
    
    '�����ǰ��Ŀ����ʾ
    Me.txt���.Text = "": Me.txt����.Text = "": Me.txt˵��.Text = ""
    Me.chk����.Value = vbUnchecked
    
    '��ȡָ����Ŀ����Ϣ
    Err = 0: On Error GoTo errHand
    gstrSQL = "Select ID, ���, ����, ˵��,����, ͨ��,����, ��ʽ From �����ļ��б� Where ���� = 7 And ID = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngBillId)
    With rsTemp
        Me.txt���.MaxLength = .Fields("���").DefinedSize
        Me.txt����.MaxLength = .Fields("����").DefinedSize
        Me.txt˵��.MaxLength = .Fields("˵��").DefinedSize
        
        Me.cbxSubClass.ListIndex = 0
        
        If .RecordCount > 0 Then
            Me.txt���.Text = "" & !���
            Me.txt����.Text = "" & !����
            Me.txt˵��.Text = "" & !˵��
            
            For i = 0 To cbxSubClass.ListCount - 1
                If Me.cbxSubClass.List(i) = NVL(!����) Then
                    Me.cbxSubClass.ListIndex = i
                    Exit For
                End If
            Next i
            
            Select Case NVL(!����, 0)
            Case 2
                chk����.Value = vbChecked: optEditType(1).Value = True
            Case 0
                Select Case NVL(!ͨ��, 0)
                Case 2
                    Me.chk����.Value = vbChecked: Me.opt����(1).Value = True
                Case 1
                    Me.chk����.Value = vbChecked: Me.opt����(0).Value = True
                Case Else
                    Me.chk����.Value = vbUnchecked
                End Select
            End Select
            optApply(Val(!��ʽ & "")).Value = True
        End If
    End With
    
    gstrSQL = "Select ��Ŀ, Nvl(����, 0) As ����,nvl(ֻ��,0) as ֻ��, Nvl(Ҫ��id, 0) As Ҫ��, ���� From �������ݸ��� Where �ļ�id = [1] Order By ����"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngBillId)
    
    Err = 0: On Error GoTo 0
    With Me.vfg����
        .Redraw = flexRDNone
        .Clear
        
        Set .DataSource = rsTemp
        
        .ColComboList(mCol.Ҫ��id) = mstrCombos
        .ColWidth(mCol.��Ŀ) = 1500: .ColWidth(mCol.����) = 450: .ColWidth(mCol.ֻ��) = 450: .ColWidth(mCol.Ҫ��id) = 1000
        
        For lngCount = 0 To .Cols - 1
            .FixedAlignment(lngCount) = flexAlignCenterCenter
        Next
        
        For lngCount = .FixedRows To .Rows - 1
            .Cell(flexcpChecked, lngCount, mCol.����) = IIf(.TextMatrix(lngCount, mCol.����) = "0", flexUnchecked, flexChecked)
            .Cell(flexcpChecked, lngCount, mCol.ֻ��) = IIf(.TextMatrix(lngCount, mCol.ֻ��) = "0", flexUnchecked, flexChecked)
            
            .TextMatrix(lngCount, mCol.����) = ""
            .TextMatrix(lngCount, mCol.ֻ��) = ""
        Next
        If .Rows > .FixedRows Then .Row = .FixedRows
        .Redraw = flexRDDirect
    End With
    LoadApplyFile
    zlRefresh = True: Exit Function

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlRefresh = False: Exit Function
End Function

Public Function zlEditStart(blnAdd As Boolean, lngBillId As Long) As Boolean
    '���ܣ���ʼ��Ŀ�༭
    '������ blnAdd-�Ƿ����ӣ�����Ϊ�޸�
    '       lngBillID-���ӵĲ�����Ŀ������ָ���༭����Ŀ
    Dim rsTemp As New ADODB.Recordset
    
    If blnAdd Then
        Err = 0: On Error GoTo errHand
        gstrSQL = "Select Nvl(Max(To_Number(���)), 0) As ���, Nvl(Max(Length(���)), 0) As ���� From �����ļ��б� Where ���� = 7"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���")
        
        If rsTemp!���� <> 0 And rsTemp!���� <= Me.txt���.MaxLength Then
            Me.txt���.Text = Format(Val(rsTemp!���) + 1, String(rsTemp!����, "0"))
        Else
            Me.txt���.Text = Format(Val(rsTemp!���) + 1, String(Me.txt���.MaxLength, "0"))
        End If
        Me.txt����.Text = "": Me.txt˵��.Text = ""
    Else
        If optEditType(1).Value Then
            optEditType(0).Enabled = False: opt����(0).Enabled = False: opt����(1).Enabled = False
        Else
            optEditType(1).Enabled = False
        End If
    End If

    Me.Tag = IIf(blnAdd, "����", "�޸�")
    Me.picEdit.Enabled = True: Call Form_Resize
    Me.txt���.SetFocus
    zlEditStart = True: Exit Function

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlEditStart = False: Exit Function
End Function

Public Sub zlEditCancel()
    '���ܣ��������ڽ��еı༭
    Me.Tag = ""
    Me.picEdit.Enabled = False: Call Form_Resize
    Call Me.zlRefresh(mlngBillID)
End Sub

Public Function zlEditSave() As Long
    '���ܣ��������ڽ��еı༭,���������ڱ༭��Ŀid,����ʧ�ܷ���0
Dim lngNewId As Long, strLists As String
Dim strSQL As String, rsTmp As ADODB.Recordset
Dim strRPTPass1 As String, strRPTPass2 As String
Dim lngCount As Long
    Dim i As Long, blnTrans As Boolean
    Static objRpt As clsReport
    
    'һ�����Լ��
    If Trim(Me.txt���.Text) = "" Then
        MsgBox "�������ţ�", vbInformation, gstrSysName
        Me.txt���.SetFocus: zlEditSave = 0: Exit Function
    End If
    
    If Val(Me.txt���.Text) > Val(String(Me.txt���.MaxLength, "9")) Then
        MsgBox "���̫��", vbInformation, gstrSysName
        Me.txt���.SetFocus: zlEditSave = 0: Exit Function
    End If
    
    If Me.Tag = "����" Then
        strSQL = "Select ID, ���, ����, ˵��, ͨ�� From �����ļ��б� Where ���� = 7 And ��� = [1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Format(Trim(Me.txt���.Text), String(Me.txt���.MaxLength, "0")))
        If rsTmp.RecordCount > 0 Then
            MsgBox "����ظ���", vbInformation, gstrSysName
            Me.txt���.SetFocus: zlEditSave = 0: Exit Function
        End If
    End If
    
    If Trim(Me.txt����.Text) = "" Then
        MsgBox "���������ƣ�", vbInformation, gstrSysName
        Me.txt����.SetFocus: zlEditSave = 0: Exit Function
    End If
    
    If LenB(StrConv(Trim(Me.txt����.Text), vbFromUnicode)) > Me.txt����.MaxLength Then
        MsgBox "���Ƴ��������" & Me.txt����.MaxLength & "���ַ���ȳ����֣���", vbInformation, gstrSysName
        Me.txt����.SetFocus: zlEditSave = 0: Exit Function
    End If
    
    If LenB(StrConv(Trim(Me.txt˵��.Text), vbFromUnicode)) > Me.txt˵��.MaxLength Then
        MsgBox "˵�����������" & Me.txt˵��.MaxLength & "���ַ���ȳ����֣���", vbInformation, gstrSysName
        Me.txt˵��.SetFocus: zlEditSave = 0: Exit Function
    End If
    
    With Me.vfg����
        strLists = ""
        If .Col = mCol.Ҫ��id Then .Select .Row, .Col + 1
        For lngCount = .FixedRows To .Rows - 1
            strLists = strLists & vbLf & .TextMatrix(lngCount, mCol.��Ŀ)
            
            If .Cell(flexcpChecked, lngCount, mCol.����) = flexChecked Then
                strLists = strLists & vbTab & "1"
            Else
                strLists = strLists & vbTab & "0"
            End If
            
            If .Cell(flexcpChecked, lngCount, mCol.ֻ��) = flexChecked Then
                strLists = strLists & vbTab & "1"
            Else
                strLists = strLists & vbTab & "0"
            End If
            
            
            If Val(.TextMatrix(lngCount, mCol.Ҫ��id)) = 0 Then
                strLists = strLists & vbTab
            Else
                strLists = strLists & vbTab & .TextMatrix(lngCount, mCol.Ҫ��id)
            End If
            
            strLists = strLists & vbTab & .TextMatrix(lngCount, mCol.����)
        Next
        If strLists <> "" Then strLists = Mid(strLists, 2)
    End With
    
    '���ݱ��������֯
    gstrSQL = "'" & Format(Trim(Me.txt���.Text), String(Me.txt���.MaxLength, "0")) & "','" & Trim(Me.txt����.Text) & "','" & Trim(Me.txt˵��.Text) & "'"
    If optEditType(1).Value Then   '����
        gstrSQL = gstrSQL & ",2"
    Else
        gstrSQL = gstrSQL & ",0"
    End If
    
    If Me.chk����.Value = vbUnchecked Then 'ͨ��
        gstrSQL = gstrSQL & ",0"
    ElseIf optEditType(0).Value Then
        If Me.opt����(0).Value Then
            gstrSQL = gstrSQL & ",1"
        Else
            gstrSQL = gstrSQL & ",2"
        End If
    Else
        gstrSQL = gstrSQL & ",1"
    End If
    
    'ZLCISBILL00' || ���_In || '-' || Form_In
    '11698 - ���������ı������ʱ����ԭ�򣬱�������δ���ɡ�
    If objRpt Is Nothing Then
        Set objRpt = New clsReport
        Call objRpt.InitOracle(gcnOracle)
    End If
    
    strRPTPass1 = objRpt.GenReportPass("ZLCISBILL00" & Format(Trim(Me.txt���.Text), String(Me.txt���.MaxLength, "0")) & "-1", Trim(Me.txt����.Text))
    strRPTPass2 = objRpt.GenReportPass("ZLCISBILL00" & Format(Trim(Me.txt���.Text), String(Me.txt���.MaxLength, "0")) & "-2", Trim(Me.txt����.Text))
    
    If Me.Tag = "����" Then
        lngNewId = zlDatabase.GetNextId("�����ļ��б�")
        gstrSQL = "Zl_���Ƶ���Ŀ¼_Edit(1," & lngNewId & "," & gstrSQL & ",'" & strLists & "','" & Replace(strRPTPass1, "'", "''") & "','" & Replace(strRPTPass2, "'", "''") & "','" & cbxSubClass.Text & "'," & IIf(optApply(1).Value, 1, 0) & ")"
    Else
        lngNewId = mlngBillID
        gstrSQL = "Zl_���Ƶ���Ŀ¼_Edit(2," & lngNewId & "," & gstrSQL & ",'" & strLists & "','" & Replace(strRPTPass1, "'", "''") & "','" & Replace(strRPTPass2, "'", "''") & "','" & cbxSubClass.Text & "'," & IIf(optApply(1).Value, 1, 0) & ")"
    End If
    Err = 0: On Error GoTo errHand
    Screen.MousePointer = 11
    gcnOracle.BeginTrans: blnTrans = True
    Call zlDatabase.ExecuteProcedure(gstrSQL, "�������Ƶ���")
    
    For i = LBound(arrSQL) To UBound(arrSQL)
        If arrSQL(i) <> "" Then Call zlDatabase.ExecuteProcedure(arrSQL(i), Me.Caption)
    Next
    For i = LBound(arrSQLFile) To UBound(arrSQLFile)
        If arrSQLFile(i) <> "" Then Call zlDatabase.ExecuteProcedure(arrSQLFile(i), Me.Caption)
    Next

    gcnOracle.CommitTrans: blnTrans = False
    Screen.MousePointer = 0
    
    If Me.Tag = "����" Then mlngBillID = lngNewId
    Me.Tag = ""
    Me.picEdit.Enabled = False: Call Form_Resize
    zlEditSave = mlngBillID: Exit Function
    
errHand:
    If blnTrans Then gcnOracle.RollbackTrans
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlEditSave = 0: Exit Function
End Function

'--------------------------------------------
'����Ϊ����ؼ���Ӧ�¼�
'--------------------------------------------
Private Sub chk����_Click()
    If Me.Tag <> "�޸�" Then
        If Me.chk����.Value = vbUnchecked Then
            optEditType(0).Enabled = False: optEditType(0).Value = False
            optEditType(1).Enabled = False: optEditType(1).Value = False
            Me.opt����(0).Enabled = False: Me.opt����(0).Value = False
            Me.opt����(1).Enabled = False: Me.opt����(1).Value = False
        Else
            optEditType(0).Value = True: optEditType(0).Enabled = True
            optEditType(1).Enabled = True
            Me.opt����(0).Value = True: Me.opt����(0).Enabled = True
            Me.opt����(1).Enabled = True
        End If
    Else
        If optEditType(0).Value Then
            optEditType(0).Enabled = chk����.Value = 1: opt����(0).Enabled = chk����.Value = 1: opt����(1).Enabled = chk����.Value = 1
        Else
            optEditType(0).Enabled = False: optEditType(1).Enabled = False
            opt����(0).Enabled = False: opt����(1).Enabled = False
        End If
    End If
End Sub

Private Sub cmdEdit_Click(Index As Integer)
Dim strCell As String
Dim lngCount As Long
    With Me.vfg����
        Select Case Index
        Case 0
            .Rows = .Rows + 1: .Row = .Rows - 1
            
            .Cell(flexcpChecked, .Row, mCol.����) = flexUnchecked
            .Cell(flexcpChecked, .Row, mCol.ֻ��) = flexUnchecked
            
            If .RowIsVisible(.Row) = False Then .TopRow = .Row
        Case 1
            If .Row < .FixedRows Then Exit Sub
            .RemoveItem .Row
        Case 2
            If .Row < .FixedRows + 1 Then Exit Sub
            For lngCount = 0 To .Cols - 1
                strCell = .TextMatrix(.Row - 1, lngCount)
                .TextMatrix(.Row - 1, lngCount) = .TextMatrix(.Row, lngCount)
                .TextMatrix(.Row, lngCount) = strCell
            Next
            .Row = .Row - 1
        Case 3
            If .Row < .FixedRows Then Exit Sub
            If .Row >= .Rows - 1 Then Exit Sub
            For lngCount = 0 To .Cols - 1
                strCell = .TextMatrix(.Row + 1, lngCount)
                .TextMatrix(.Row + 1, lngCount) = .TextMatrix(.Row, lngCount)
                .TextMatrix(.Row, lngCount) = strCell
            Next
            .Row = .Row + 1
        End Select
        If .Visible And .Editable Then .SetFocus
    End With
End Sub

Private Sub Form_Load()
    Dim strLists As String
    
    '���뵥����������
    Call LoadSubClass
    With vsFile
        .ColWidth(0) = 1500: .ColWidth(1) = 255
        .TextMatrix(0, 0) = "����"
        .TextMatrix(0, 2) = "�ļ���"
        .TextMatrix(1, 0) = "xsl��ʽ�ļ�"
        .TextMatrix(2, 0) = "xml�����ļ�"
        .TextMatrix(3, 0) = "html��ʾ�ļ�"
        .Cell(flexcpAlignment, 0, 0, 0, 2) = flexAlignCenterCenter
        .Cell(flexcpAlignment, 0, 0, 3, 0) = flexAlignCenterCenter
    End With
    
    
    mlngBillID = 0: Me.picEdit.BackColor = Me.BackColor: picEditType.BackColor = Me.BackColor
    Err = 0: On Error GoTo errHand
    gstrSQL = "Select I.ID, I.������,decode(I.��ʾ��,4,I.��ֵ��,NULL) as ��ֵ��" & vbNewLine & _
            "From ����������Ŀ I, ������������ K" & vbNewLine & _
            "Where I.����id = K.ID And (K.���� = 1 And K.���� = '06' And" & vbNewLine & _
            "      I.������ Not In ('�������', 'һ��סԺ���', '����סԺ���', '�ϴ�סԺ���')  Or k.���� = 6) " & vbNewLine & _
            "Order By I.����"
    strLists = ""
    
    Call zlDatabase.OpenRecordset(mrsItems, gstrSQL, "��ȡ����������Ŀ")
    
    Do While Not mrsItems.EOF
        strLists = strLists & "|#" & mrsItems!ID & ";" & mrsItems!������
        mrsItems.MoveNext
    Loop
    
    mstrCombos = "#0; " & strLists
    
    Exit Sub

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub LoadSubClass()
'����������Ŀ���
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    
    strSQL = "select ���� from ������Ŀ��� order by ����"
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    
    cbxSubClass.Clear
    
    If rsData.RecordCount <= 0 Then Exit Sub
    
    cbxSubClass.AddItem ""
    While Not rsData.EOF
        Call cbxSubClass.AddItem(NVL(rsData!����))
        
        Call rsData.MoveNext
    Wend
    
End Sub

Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    With Me.vfg����
        If Me.picEdit.Enabled = True Then
            Me.picEdit.BackColor = RGB(230, 230, 230)
            .FocusRect = flexFocusHeavy: .Editable = flexEDKbdMouse
            .Height = Me.cmdEdit(0).Top - .Top - Screen.TwipsPerPixelY
        Else
            Me.picEdit.BackColor = Me.BackColor
            .FocusRect = flexFocusNone: .Editable = flexEDNone
            .Height = Me.cmdEdit(0).Top + Me.cmdEdit(0).Height - .Top
        End If
    End With
    Me.chk����.BackColor = Me.picEdit.BackColor
    Me.opt����(0).BackColor = Me.picEdit.BackColor
    Me.opt����(1).BackColor = Me.picEdit.BackColor
    Me.picEditType.BackColor = Me.picEdit.BackColor
    Me.optEditType(0).BackColor = Me.picEdit.BackColor
    Me.optEditType(1).BackColor = Me.picEdit.BackColor
    Me.Label1.BackColor = Me.picEdit.BackColor
    Me.picApplyType.BackColor = Me.picEdit.BackColor
    Me.optApply(0).BackColor = Me.picEdit.BackColor
    Me.optApply(1).BackColor = Me.picEdit.BackColor
    Me.picApply.BackColor = Me.picEdit.BackColor
    lblApplysMark.BackColor = Me.picEdit.BackColor
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mrsItems = Nothing
End Sub

Private Sub optApply_Click(Index As Integer)
    If Index = 1 Then
        picApply.Visible = True
        vfg����.Visible = False
    Else
        picApply.Visible = False
        vfg����.Visible = True
    End If
End Sub

Private Sub optEditType_Click(Index As Integer)
    If Index = 0 Then
        opt����(0).Value = True: opt����(0).Enabled = True
        opt����(1).Enabled = True
    Else
        opt����(0).Value = False: opt����(1).Value = False
        opt����(0).Enabled = False: opt����(1).Enabled = False
    End If
End Sub

Private Sub txt���_GotFocus()
    Me.txt���.SelStart = 0: Me.txt���.SelLength = 1000
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txt���_KeyPress(KeyAscii As Integer)
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

Private Sub txt����_GotFocus()
    Me.txt����.SelStart = 0: Me.txt����.SelLength = 1000
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr("%_'""", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt˵��_GotFocus()
    Me.txt˵��.SelStart = 0: Me.txt˵��.SelLength = 1000
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt˵��_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr("%_'""", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub vfg����_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If Col <> mCol.��Ŀ Then Exit Sub
    Me.vfg����.TextMatrix(Row, Col) = Replace(Me.vfg����.TextMatrix(Row, Col), """", "")
    Me.vfg����.TextMatrix(Row, Col) = Replace(Me.vfg����.TextMatrix(Row, Col), "'", "")
    Me.vfg����.TextMatrix(Row, Col) = Replace(Me.vfg����.TextMatrix(Row, Col), vbTab, " ")
    Me.vfg����.TextMatrix(Row, Col) = Replace(Me.vfg����.TextMatrix(Row, Col), vbLf, " ")
End Sub

Private Sub vfg����_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If Val(vfg����.TextMatrix(NewRow, mCol.Ҫ��id)) <> 0 Then
        mrsItems.Filter = "ID=" & Val(vfg����.TextMatrix(NewRow, mCol.Ҫ��id))
        If mrsItems.RecordCount > 0 Then
            mrsItems.MoveFirst
            vfg����.ColComboList(mCol.����) = " |" & Replace(mrsItems!��ֵ�� & "", ";", "|")
        Else
           vfg����.ColComboList(mCol.����) = ""
        End If
    Else
        vfg����.ColComboList(mCol.����) = ""
    End If
End Sub

Private Sub vfg����_DblClick()
    With Me.vfg����
        If .Row < .FixedRows Then Exit Sub
        
        If .Col = mCol.���� Then
            If .Cell(flexcpChecked, .Row, mCol.����) = flexChecked Then
                .Cell(flexcpChecked, .Row, mCol.����) = flexUnchecked
            Else
                .Cell(flexcpChecked, .Row, mCol.����) = flexChecked
            End If
        End If
        
        If .Col = mCol.ֻ�� Then
            If .Cell(flexcpChecked, .Row, mCol.ֻ��) = flexChecked Then
                .Cell(flexcpChecked, .Row, mCol.ֻ��) = flexUnchecked
            Else
                .Cell(flexcpChecked, .Row, mCol.ֻ��) = flexChecked
            End If
        End If
        
    End With
End Sub

Private Sub vfg����_KeyPress(KeyAscii As Integer)
    Call vfg����_DblClick
End Sub

Private Sub vfg����_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If Chr(KeyAscii) = "'" Then KeyAscii = 0
    If Chr(KeyAscii) = """" Then KeyAscii = 0
End Sub

Private Sub vfg����_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = mCol.���� Or Col = mCol.ֻ�� Then Cancel = True
End Sub

Private Sub vsFile_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    vsFile.Editable = flexEDKbdMouse
    vsFile.ColComboList(2) = "..."
End Sub

Private Function LoadApplyFile() As Boolean
'���ܣ���ʾ�ٴ�·���ļ����ݺͻ����ٴ�·��
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Long

    On Error GoTo ErrH
    
    strSQL = "Select �ļ���, ��� From �Զ������뵥�ļ� Where �ļ�id = [1] Order By ���"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngBillID)
    With vsFile
        .TextMatrix(1, 2) = ""
        .TextMatrix(2, 2) = ""
        .TextMatrix(3, 2) = ""
        Set .Cell(flexcpPicture, 1, 1) = Nothing
        Set .Cell(flexcpPicture, 2, 1) = Nothing
        Set .Cell(flexcpPicture, 3, 1) = Nothing
        For i = 1 To rsTmp.RecordCount
            .TextMatrix(Val(rsTmp!��� & ""), 2) = rsTmp!�ļ���
            Set .Cell(flexcpPicture, Val(rsTmp!��� & ""), 1) = zlCommFun.GetFileIcon(rsTmp!�ļ���, True, App.hInstance)
            .Cell(flexcpPictureAlignment, Val(rsTmp!��� & ""), 1) = flexAlignCenterCenter
            rsTmp.MoveNext
        Next
    End With

    LoadApplyFile = True
    Exit Function
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub vsFile_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim arrTmp() As String
    Dim strFile As String
    Dim strFileName As String
    Dim strSQL As String
    Dim StrText As String
    Dim i As Long
    Dim stmTmp As Stream
    
	If mblndb = True Then mblndb = False: Exit Sub
    If Row = 1 Then
        cdgFile.DialogTitle = "ѡ��Ҫ��ӵ�xsl��ʽ�ļ�"
        cdgFile.Filter = "html��ʾ�ļ�(*.xsl)|*.xsl"
    ElseIf Row = 2 Then
        cdgFile.DialogTitle = "ѡ��Ҫ��ӵ�xml�����ļ�"
        cdgFile.Filter = "html��ʾ�ļ�(*.xml)|*.xml"
    ElseIf Row = 3 Then
        cdgFile.DialogTitle = "ѡ��Ҫ��ӵ�html��ʾ�ļ�"
        cdgFile.Filter = "html��ʾ�ļ�(*.html)|*.html"
    End If
    cdgFile.flags = &H80000 Or &H4 Or &H1000 Or &H200000 Or &H800
    cdgFile.InitDir = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, "�Զ������뵥ѡ��Ŀ¼")
    cdgFile.CancelError = True
    On Error Resume Next
    cdgFile.ShowOpen
    If Err.Number <> 0 Then
        Err.Clear: Exit Sub
    End If
    On Error GoTo ErrH
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName, "�Զ������뵥ѡ��Ŀ¼", mobjFile.GetFile(cdgFile.Filename).ParentFolder.Path
    strFile = cdgFile.Filename '����·��
    strFileName = mobjFile.GetFile(cdgFile.Filename).Name
    
    '����ļ���С������3M
    If mobjFile.GetFile(strFile).Size / 1024 / 1024 > 3 Then
        MsgBox "�ļ��ߴ�̫��(����3M)������ļ������ʵ������������ӡ�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    On Error Resume Next
    Set stmTmp = New Stream
    stmTmp.Open
    stmTmp.Charset = "UTF-8"
    stmTmp.Position = 0
    stmTmp.LoadFromFile strFile
    stmTmp.Position = 0
    
    StrText = stmTmp.ReadText
    stmTmp.Close

    If Err.Number <> 0 Then
        MsgBox "�ļ����ʧ��,�����ļ���ʽ���ļ������Ƿ���ȷ��", vbExclamation, gstrSysName
        Screen.MousePointer = 0: Err.Clear: Exit Sub
    End If
    On Error GoTo ErrH

    ReDim arrTmp(0)
    strSQL = "Zl_�Զ������뵥�ļ�_Edit(1," & mlngBillID & "," & Row & ",'" & strFileName & "')"
    If Not Sys.GetlobSql(glngSys, 24, mlngBillID & "," & Row, Replace(StrText, "'", "''"), arrTmp(), 1) Then
        MsgBox "�ļ����ʧ�ܣ�", vbExclamation, gstrSysName
        Screen.MousePointer = 0: Exit Sub
    End If
    If arrSQL(UBound(arrSQL)) <> "" Then ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = strSQL
    For i = LBound(arrTmp) To UBound(arrTmp)
        If arrSQLFile(UBound(arrSQLFile)) <> "" Then ReDim Preserve arrSQLFile(UBound(arrSQLFile) + 1)
        arrSQLFile(UBound(arrSQLFile)) = arrTmp(i)
    Next

    vsFile.TextMatrix(Row, 2) = strFileName
    Set vsFile.Cell(flexcpPicture, Row, 1) = zlCommFun.GetFileIcon(strFileName, True, App.hInstance)
    vsFile.Cell(flexcpPictureAlignment, Row, 1) = flexAlignCenterCenter
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vsFile_DblClick()
    Dim strFile As String
    Dim lngRetu As Long, strInfo As String
    Dim StrText As String
    Dim stmTmp As Stream
    
On Error Resume Next
    Screen.MousePointer = 11
    mblndb = True
    strFile = mobjFile.GetSpecialFolder(TemporaryFolder) & "\" & vsFile.TextMatrix(vsFile.Row, 2)
    If mobjFile.FileExists(strFile) Then mobjFile.DeleteFile strFile, True
    
    StrText = Sys.Readlob(glngSys, 24, mlngBillID & "," & vsFile.Row, strFile, 1)
    
    If Not mobjFile.FileExists(strFile) Then Call mobjFile.CreateTextFile(strFile)
    Set stmTmp = New Stream
    stmTmp.Open
    stmTmp.Charset = "UTF-8"
    stmTmp.WriteText StrText
    stmTmp.SaveToFile strFile, adSaveCreateOverWrite
    stmTmp.Close
    
    If Not mobjFile.FileExists(strFile) Then
        MsgBox "�ļ����ݶ�ȡʧ�ܣ�", vbInformation, gstrSysName:
        Screen.MousePointer = 0: Exit Sub
    End If
    
    lngRetu = ShellExecute(Me.hWnd, "open", strFile, "", "", SW_SHOWNORMAL)
    If lngRetu <= 32 Then
        Select Case lngRetu
        Case 2: strInfo = "����Ĺ���"
        Case 29: strInfo = "����ʧ��"
        Case 30: strInfo = "����Ӧ�ó�ʽæµ��..."
        Case 31: strInfo = "û�й����κ�Ӧ�ó�ʽ"
        Case Else: strInfo = "�޷�ʶ��Ĵ���"
        End Select
        MsgBox "�ļ���ʱ����" & vbCrLf & vbCrLf & vbTab & strInfo, vbExclamation, gstrSysName
    End If
    'If mobjFile.FileExists(strFile) Then mobjFile.DeleteFile strFile, True
    
    Screen.MousePointer = 0
End Sub
