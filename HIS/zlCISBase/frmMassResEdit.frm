VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmMassResEdit 
   BorderStyle     =   0  'None
   Caption         =   "�ʿ�Ʒ��Ϣ"
   ClientHeight    =   7890
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8790
   LinkTopic       =   "Form1"
   ScaleHeight     =   7890
   ScaleWidth      =   8790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CheckBox chkHide 
      Caption         =   "����������"
      Height          =   195
      Left            =   7335
      TabIndex        =   21
      Top             =   98
      Value           =   1  'Checked
      Width           =   1350
   End
   Begin VSFlex8Ctl.VSFlexGrid vfgList 
      Height          =   4605
      Left            =   3240
      TabIndex        =   17
      Top             =   330
      Width           =   5460
      _cx             =   9631
      _cy             =   8123
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
      FixedRows       =   2
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
      WallPaperAlignment=   8
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.PictureBox picEdit 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   7845
      Left            =   0
      ScaleHeight     =   7845
      ScaleMode       =   0  'User
      ScaleWidth      =   3225
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   3225
      Begin VB.ComboBox cboУ׼�� 
         Height          =   300
         Left            =   405
         TabIndex        =   28
         Top             =   4785
         Width           =   2745
      End
      Begin VB.ComboBox cbo�Լ� 
         Height          =   300
         Left            =   405
         TabIndex        =   26
         Top             =   4200
         Width           =   2745
      End
      Begin VB.TextBox txtȡֵ���� 
         Height          =   300
         Left            =   405
         TabIndex        =   23
         Top             =   3090
         Width           =   2745
      End
      Begin VB.TextBox txt����ֵ 
         Height          =   300
         Left            =   405
         TabIndex        =   22
         Top             =   3645
         Width           =   2745
      End
      Begin VB.CheckBox chk�Ƕ�ֵ 
         Caption         =   "�Ƕ�ֵ (�Ƕ�ֵ����Ԥ��ֵ)"
         Height          =   195
         Left            =   600
         TabIndex        =   10
         Top             =   1440
         Width           =   2925
      End
      Begin VB.ComboBox cboˮƽ 
         Height          =   300
         Left            =   2310
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1080
         Width           =   825
      End
      Begin VB.ComboBox cbo���� 
         Height          =   300
         Left            =   330
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   7665
         Visible         =   0   'False
         Width           =   2550
      End
      Begin VB.TextBox txtŨ�� 
         Height          =   300
         Left            =   600
         MaxLength       =   10
         TabIndex        =   6
         Top             =   1080
         Width           =   1710
      End
      Begin VB.TextBox txt���� 
         Height          =   300
         Left            =   600
         MaxLength       =   10
         TabIndex        =   2
         Top             =   405
         Width           =   1155
      End
      Begin VB.TextBox txt���� 
         Height          =   300
         Left            =   600
         MaxLength       =   40
         TabIndex        =   4
         Top             =   735
         Width           =   2520
      End
      Begin VB.TextBox txt�걾�� 
         Height          =   300
         Left            =   390
         MaxLength       =   40
         TabIndex        =   16
         ToolTipText     =   "����ָ�����ʿ���ʹ�õı걾�ţ��걾��֮����,�ָ�"
         Top             =   2460
         Width           =   2745
      End
      Begin MSComCtl2.DTPicker dtp��ʼ���� 
         Height          =   300
         Left            =   390
         TabIndex        =   12
         Top             =   1905
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   101711875
         CurrentDate     =   39064
      End
      Begin MSComCtl2.DTPicker dtp�������� 
         Height          =   300
         Left            =   1890
         TabIndex        =   14
         Top             =   1905
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   101711875
         CurrentDate     =   39429
      End
      Begin VB.Label lblУ׼�� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "У׼��:"
         Height          =   180
         Left            =   180
         TabIndex        =   29
         Top             =   4545
         Width           =   630
      End
      Begin VB.Label lbl�Լ���Դ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�Լ�:"
         Height          =   180
         Left            =   210
         TabIndex        =   27
         Top             =   3960
         Width           =   450
      End
      Begin VB.Label lblȡֵ���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ȡֵ����:"
         Height          =   180
         Left            =   210
         TabIndex        =   25
         Top             =   2880
         Width           =   810
      End
      Begin VB.Label lbl����ֵ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ȡֵ���ж�Ӧ����:"
         Height          =   180
         Left            =   210
         TabIndex        =   24
         Top             =   3435
         Width           =   1530
      End
      Begin VB.Label lbl������Ϣ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����Ʒ��Ϣ:"
         Height          =   180
         Left            =   150
         TabIndex        =   19
         Top             =   150
         Width           =   990
      End
      Begin VB.Label lblNote 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   $"frmMassResEdit.frx":0000
         ForeColor       =   &H00008000&
         Height          =   2340
         Left            =   360
         TabIndex        =   18
         Top             =   5235
         Width           =   2865
         WordWrap        =   -1  'True
      End
      Begin VB.Image imgNote 
         Height          =   240
         Left            =   45
         Picture         =   "frmMassResEdit.frx":0153
         Top             =   5205
         Width           =   240
      End
      Begin VB.Label lbl�������� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��"
         Height          =   180
         Left            =   1680
         TabIndex        =   13
         Top             =   1965
         Width           =   180
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   195
         TabIndex        =   3
         Top             =   795
         Width           =   360
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   195
         TabIndex        =   1
         Top             =   465
         Width           =   360
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   60
         TabIndex        =   8
         Top             =   7710
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label lblŨ�� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ũ��"
         Height          =   180
         Left            =   195
         TabIndex        =   5
         Top             =   1140
         Width           =   360
      End
      Begin VB.Label lbl��ʼ���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ʹ�����ڷ�Χ:"
         Height          =   180
         Left            =   195
         TabIndex        =   11
         Top             =   1680
         Width           =   1170
      End
      Begin VB.Label lbl�걾�� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Ӧ�걾��:"
         Height          =   180
         Left            =   195
         TabIndex        =   15
         Top             =   2250
         Width           =   990
      End
   End
   Begin VB.Label lbl�ʿ���Ŀ 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�ʿؼ����Ŀ: "
      Height          =   180
      Left            =   3240
      TabIndex        =   20
      Top             =   105
      Width           =   1260
   End
End
Attribute VB_Name = "frmMassResEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngResId As Long          '��ǰ��ʾ����Ʒid
Private mlngDevId As Long          '��ǰ��ʾ������id

Private Enum mCol
    ID = 0: ѡ��: ������: Ӣ����: ��λ: ȡֵ����: �ο���ֵ: �ο�SD: Ԥ���ֵ: Ԥ��SD: ��Ŀqc��: ����qc��: �������: �ʿ�ȡֵ: ����: ����ֵ
End Enum
Private mstr���� As String '���ڱ���е�ѡ����
Private mlng�������� As Long '0-��ͨ���� 1-΢���� 2-ø����
Dim lngCount As Long
Dim lngLastID As Long
Private mblnEditRow As Boolean  '�Ƿ��޸�������ֵ

'--------------------------------------------
'����Ϊ���幫������
'--------------------------------------------
Private Sub setListFormat(Optional blnKeepData As Boolean)
    '���ܣ���ʼ�����òο�ֵ�б�
    '������ blnKeepData-�Ƿ������ݣ���ֻ���������ø�ʽ
    Dim strLists As String, strValue As String
    
    With Me.vfgList
        .Redraw = flexRDNone
        If blnKeepData = False Then
            .Clear
            .Rows = 2: .FixedRows = 2: .Cols = 16: .FixedCols = 0
        End If
        If .Cols < 16 Then .Cols = 16
        .MergeCells = flexMergeFixedOnly: .MergeRow(0) = True: .MergeCol(mCol.����) = True: .MergeCol(mCol.�ʿ�ȡֵ) = True
       
        .TextMatrix(0, mCol.ѡ��) = "��Ŀ": .TextMatrix(0, mCol.������) = "��Ŀ": .TextMatrix(0, mCol.Ӣ����) = "��Ŀ"
        .TextMatrix(0, mCol.��λ) = "��Ŀ": .TextMatrix(1, mCol.ȡֵ����) = "��Ŀ"
        .TextMatrix(1, mCol.ѡ��) = "": .TextMatrix(1, mCol.������) = "������": .TextMatrix(1, mCol.Ӣ����) = "Ӣ����"
        .TextMatrix(1, mCol.��λ) = "��λ": .TextMatrix(1, mCol.ȡֵ����) = "ȡֵ����"
        .TextMatrix(0, mCol.�ο���ֵ) = "�ο�����ֵ": .TextMatrix(0, mCol.�ο�SD) = "�ο�����ֵ"
        .TextMatrix(1, mCol.�ο���ֵ) = "��ֵ": .TextMatrix(1, mCol.�ο�SD) = "SD"
        .TextMatrix(0, mCol.Ԥ���ֵ) = "Ԥ�����ֵ": .TextMatrix(0, mCol.Ԥ��SD) = "Ԥ�����ֵ"
        .TextMatrix(1, mCol.Ԥ���ֵ) = "��ֵ": .TextMatrix(1, mCol.Ԥ��SD) = "SD"
        .TextMatrix(0, mCol.��Ŀqc��) = "��ӦQC��": .TextMatrix(0, mCol.����qc��) = "��ӦQC��"
        .TextMatrix(1, mCol.��Ŀqc��) = "��Ŀ��": .TextMatrix(1, mCol.����qc��) = "������"
        .TextMatrix(0, mCol.����) = "����": .TextMatrix(1, mCol.����) = "����"
        .TextMatrix(0, mCol.�������) = "�������": .TextMatrix(1, mCol.�������) = "�������"
        .TextMatrix(0, mCol.����ֵ) = "����ֵ": .TextMatrix(1, mCol.����ֵ) = "����ֵ"
        
        .TextMatrix(0, mCol.�ʿ�ȡֵ) = "�ʿ�ȡֵ": .TextMatrix(1, mCol.�ʿ�ȡֵ) = "�ʿ�ȡֵ"
        
        
        .ColWidth(mCol.������) = IIf(Me.chkHide.Value = vbChecked, 0, 2000)
        .ColWidth(mCol.Ӣ����) = 800: .ColWidth(mCol.��λ) = 800: .ColWidth(mCol.ȡֵ����) = 0
        .ColWidth(mCol.�ο���ֵ) = 700: .ColWidth(mCol.�ο�SD) = 700
        .ColWidth(mCol.Ԥ���ֵ) = 700: .ColWidth(mCol.Ԥ��SD) = 700
        .ColWidth(mCol.��Ŀqc��) = 800: .ColWidth(mCol.����qc��) = 800
        .ColWidth(mCol.ID) = 0: .ColWidth(mCol.ѡ��) = 270
        .ColWidth(mCol.�������) = 0: .ColWidth(mCol.����ֵ) = 0
        
        .ColWidth(mCol.�ʿ�ȡֵ) = 0
        .ColHidden(mCol.�ʿ�ȡֵ) = False
        If mlng�������� = 2 Then .ColWidth(mCol.�ʿ�ȡֵ) = 900
        
        .ColComboList(mCol.����) = mstr����
        .ColComboList(mCol.�ʿ�ȡֵ) = "|[OD]|[SCO]"
        
        For lngCount = 0 To .Cols - 1
            .FixedAlignment(lngCount) = flexAlignCenterCenter
            If .ColWidth(lngCount) = 0 Then .ColHidden(lngCount) = True
        Next
        For lngCount = .FixedRows To .Rows - 1
            If Val(.TextMatrix(lngCount, mCol.ѡ��)) = 1 Then
                .Cell(flexcpChecked, lngCount, mCol.ѡ��) = flexChecked
            Else
                .Cell(flexcpChecked, lngCount, mCol.ѡ��) = flexUnchecked
            End If
            .TextMatrix(lngCount, mCol.ѡ��) = ""
            If Trim(.TextMatrix(lngCount, mCol.ȡֵ����)) <> "" And Trim(.TextMatrix(lngCount, mCol.�ʿ�ȡֵ)) = "" Then
                strLists = Trim(.TextMatrix(lngCount, mCol.ȡֵ����))
                strValue = Trim(.TextMatrix(lngCount, mCol.�ο���ֵ))
                If Val(strValue) = Int(Val(strValue)) And Val(strValue) > 0 And Val(strValue) <= UBound(Split(strLists, ";")) + 1 Then
                    .TextMatrix(lngCount, mCol.�ο���ֵ) = Split(strLists, ";")(strValue - 1)
                Else
                    .TextMatrix(lngCount, mCol.�ο���ֵ) = "": .TextMatrix(lngCount, mCol.�ο�SD) = ""
                End If
                strValue = Trim(.TextMatrix(lngCount, mCol.Ԥ���ֵ))
                If Val(strValue) = Int(Val(strValue)) And Val(strValue) > 0 And Val(strValue) <= UBound(Split(strLists, ";")) + 1 Then
                    .TextMatrix(lngCount, mCol.Ԥ���ֵ) = Split(strLists, ";")(strValue - 1)
                Else
                    .TextMatrix(lngCount, mCol.Ԥ���ֵ) = "": .TextMatrix(lngCount, mCol.Ԥ��SD) = ""
                End If
            End If
            If Left(.TextMatrix(lngCount, mCol.�ο���ֵ), 1) = "." Then .TextMatrix(lngCount, mCol.�ο���ֵ) = "0" & .TextMatrix(lngCount, mCol.�ο���ֵ)
            If Left(.TextMatrix(lngCount, mCol.�ο�SD), 1) = "." Then .TextMatrix(lngCount, mCol.�ο�SD) = "0" & .TextMatrix(lngCount, mCol.�ο�SD)
            If Left(.TextMatrix(lngCount, mCol.Ԥ���ֵ), 1) = "." Then .TextMatrix(lngCount, mCol.Ԥ���ֵ) = "0" & .TextMatrix(lngCount, mCol.Ԥ���ֵ)
            If Left(.TextMatrix(lngCount, mCol.Ԥ��SD), 1) = "." Then .TextMatrix(lngCount, mCol.Ԥ��SD) = "0" & .TextMatrix(lngCount, mCol.Ԥ��SD)
        Next
        .Redraw = flexRDDirect
    End With
End Sub

Public Function zlRefresh(lngResID As Long, lngDevId As Long) As Boolean
    '���ܣ�������Ŀidˢ�µ�ǰ��ʾ����
    Dim rsTemp As New ADODB.Recordset
    mlngResId = lngResID
    
    '�����ǰ��Ŀ����ʾ
    Me.txt����.Text = "": Me.txt����.Text = ""
    Me.txtŨ��.Text = "": Me.cboˮƽ.Clear: Me.cbo����.ListIndex = -1
    Me.dtp��ʼ����.Value = Now(): Me.dtp��������.Value = DateAdd("m", 13, Now()) - 1
    Me.txt�걾�� = "": Me.chk�Ƕ�ֵ.Value = vbUnchecked
    '--------------------------------------------------
    ' 2009-09-03����
    Me.cbo�Լ�.Text = "": Me.cboУ׼��.Text = ""
    '--------------------------------------------------
    Err = 0: On Error GoTo errHand
    If lngDevId <> 0 And mlngDevId <> lngDevId Then
        mlngDevId = lngDevId
        gstrSql = "Select ΢���� From �������� where id=[1] "
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngDevId)
        Do Until rsTemp.EOF
            mlng�������� = Val("" & rsTemp!΢����)
            rsTemp.MoveNext
        Loop
         
    End If
    
    If lngResID = 0 Then
        Call setListFormat:        zlRefresh = True: Exit Function
    End If
    '��ȡָ����Ŀ����Ϣ
    gstrSql = "Select D.�ʿ�ˮƽ��, R.����, R.����, R.�Ƕ�ֵ, R.Ũ��, R.ˮƽ, R.����, R.��ʼ����, R.��������, R.�걾��, D.΢����, R.�Լ�, R.У׼��" & vbNewLine & _
            "From �����ʿ�Ʒ R, �������� D" & vbNewLine & _
            "Where R.����id  = D.ID And R.ID  = [1]"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, lngResID)
    With rsTemp
        If .RecordCount > 0 Then
            Me.txt����.Text = "" & !����
            Me.txt����.Text = "" & !����
            Me.txtŨ��.Text = "" & !Ũ��
            For lngCount = 0 To Me.cbo����.ListCount - 1
                If Me.cbo����.List(lngCount) = "" & !���� Then Me.cbo����.ListIndex = lngCount: Exit For
            Next
            For lngCount = 1 To Val("" & !�ʿ�ˮƽ��)
                Me.cboˮƽ.AddItem "ˮƽ" & lngCount
                If lngCount = Val("" & !ˮƽ) Then Me.cboˮƽ.ListIndex = Me.cboˮƽ.NewIndex
            Next
            
            If Not IsNull(!��ʼ����) Then Me.dtp��ʼ����.Value = !��ʼ����
            If Not IsNull(!��������) Then Me.dtp��������.Value = !��������
            Me.txt�걾��.Text = "" & !�걾��
            Me.chk�Ƕ�ֵ.Value = IIf(Val("" & !�Ƕ�ֵ) = 1, vbChecked, vbUnchecked)
            '--------------------------------------------------
            ' 2009-09-03����
            Me.cbo�Լ�.Text = Trim("" & !�Լ�)
            Me.cboУ׼��.Text = Trim("" & !У׼��)
            '--------------------------------------------------
        End If
    End With
    
    gstrSql = "Select I.ID, 1 As ѡ��, I.������, I.Ӣ����, I.��λ, Decode(P.�������, 1, '',P.ȡֵ����) As ȡֵ����," & vbNewLine & _
            "       K.��ֵ As �ο���ֵ, K.Sd As �ο�sd, X.��ֵ As Ԥ���ֵ, X.Sd As Ԥ��sd, K.��Ŀqc��, K.����qc��,P.�������, K.�ʿ�ȡֵ, K.����, '' as ����ֵ" & vbNewLine & _
            "From ����������Ŀ I, ������Ŀ P, �����ʿ�Ʒ��Ŀ K," & vbNewLine & _
            "     (Select X.�ʿ�Ʒid, X.��Ŀid, X.��ֵ, X.Sd" & vbNewLine & _
            "       From �����ʿ�Ʒ R, �����ʿؾ�ֵ X" & vbNewLine & _
            "       Where R.ID = X.�ʿ�Ʒid(+) And R.��ʼ���� = X.��ʼ����(+) And R.ID = [1]) X" & vbNewLine & _
            "Where I.ID = P.������Ŀid And I.ID = K.��Ŀid And K.�ʿ�Ʒid = X.�ʿ�Ʒid(+) And K.��Ŀid = X.��Ŀid(+) And" & vbNewLine & _
            "      K.�ʿ�Ʒid = [1]" & vbNewLine & _
            "Order By I.����"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, lngResID)
    Set Me.vfgList.DataSource = rsTemp: Call setListFormat(True)
    Call Refresh����(lngResID)
    Call Show�ο�(vfgList.Row)
    zlRefresh = True: Exit Function

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlRefresh = False: Exit Function
End Function

Public Function zlEditStart(blnAdd As Boolean, lngResID As Long, lngDevId As Long) As Boolean
    '���ܣ���ʼ��Ŀ�༭
    '������ blnAdd-�Ƿ����ӣ�����Ϊ�޸�
    '       lngResId-���ӵĲ�����Ŀ������ָ���༭����Ŀ
    '       lngDevId-��ǰ������Ʒ�������豸id
    Dim rsTemp As New ADODB.Recordset
    Dim str�Լ� As String, strУ׼�� As String
    mlngDevId = lngDevId
    
    str�Լ� = Trim(Me.cbo�Լ�.Text)
    strУ׼�� = Trim(Me.cboУ׼��.Text)
    
    Err = 0: On Error GoTo errHand
    gstrSql = "Select ���� From �ʿ��Լ���Դ"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption)
    Me.cbo�Լ�.Clear: Me.cboУ׼��.Clear
    Me.cbo�Լ�.AddItem "": cboУ׼��.AddItem ""
    Do Until rsTemp.EOF
        Me.cbo�Լ�.AddItem Trim("" & rsTemp!����)
        Me.cboУ׼��.AddItem Trim("" & rsTemp!����)
        rsTemp.MoveNext
    Loop
    
    cbo�Լ�.Text = str�Լ�: cboУ׼��.Text = strУ׼��
    
    If blnAdd Then
        gstrSql = "Select �ʿ�ˮƽ�� From �������� Where ID = [1]"
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, lngDevId)
        Me.cboˮƽ.Clear
        With rsTemp
            If .RecordCount > 0 Then
                For lngCount = 1 To Val("" & !�ʿ�ˮƽ��)
                    Me.cboˮƽ.AddItem "ˮƽ" & lngCount
                Next
            End If
        End With
        
        '��������ñ�עֵ
        Me.txt����.Text = "": Me.txt����.Text = ""
        Me.txtŨ��.Text = ""
        Me.dtp��ʼ����.Value = Now(): Me.dtp��������.Value = DateAdd("m", 13, Now()) - 1
        Me.txt�걾�� = ""
        Me.chk�Ƕ�ֵ.Value = vbUnchecked
        
        Me.cbo�Լ�.Text = "": Me.cboУ׼��.Text = ""
        
    End If
    
'    gstrSql = "Select I.ID, Decode(K.��Ŀid, Null, 0, 1) As ѡ��, I.������, I.Ӣ����, I.��λ, I.ȡֵ����, K.��ֵ As �ο���ֵ," & vbNewLine & _
'            "       K.Sd As �ο�sd, X.��ֵ As Ԥ���ֵ, X.Sd As Ԥ��sd, K.��Ŀqc��, K.����qc��, K.����" & vbNewLine & _
'            "From (Select I.ID, I.����, I.������, I.Ӣ����, I.��λ, Decode(P.�������, 3, P.ȡֵ����, '') As ȡֵ����" & vbNewLine & _
'            "       From ����������Ŀ L, ����������Ŀ I, ������Ŀ P" & vbNewLine & _
'            "       Where L.��Ŀid = I.ID And I.ID = P.������Ŀid And L.����id = [2] And" & vbNewLine & _
'            "             (P.������� = 1 Or P.������� = 3 And P.ȡֵ���� Is Not Null)) I," & vbNewLine & _
'            "     (Select �ʿ�Ʒid, ��Ŀid, ��ֵ, Sd, ��Ŀqc��, ����qc��, ���� From �����ʿ�Ʒ��Ŀ Where �ʿ�Ʒid = [1]) K," & vbNewLine & _
'            "     (Select X.�ʿ�Ʒid, X.��Ŀid, X.��ֵ, X.Sd" & vbNewLine & _
'            "       From �����ʿ�Ʒ R, �����ʿؾ�ֵ X" & vbNewLine & _
'            "       Where R.ID = X.�ʿ�Ʒid(+) And R.��ʼ���� = X.��ʼ����(+) And R.ID = [1]) X" & vbNewLine & _
'            "Where I.ID = K.��Ŀid(+) And K.�ʿ�Ʒid = X.�ʿ�Ʒid(+) And K.��Ŀid = X.��Ŀid(+)" & vbNewLine & _
'            "Order By I.����"

    gstrSql = "Select I.ID, Decode(K.��Ŀid, Null, 0, 1) As ѡ��, I.������, I.Ӣ����, I.��λ, I.ȡֵ����, K.��ֵ As �ο���ֵ," & vbNewLine & _
            "       K.Sd As �ο�sd, X.��ֵ As Ԥ���ֵ, X.Sd As Ԥ��sd, K.��Ŀqc��, K.����qc�� ,I.�������, K.�ʿ�ȡֵ, K.����,'' as ����ֵ" & vbNewLine & _
            "From (Select I.ID, I.����, I.������, I.Ӣ����, I.��λ, Decode(P.�������, 1,'', P.ȡֵ����) As ȡֵ����, P.�������" & vbNewLine & _
            "       From ����������Ŀ L, ����������Ŀ I, ������Ŀ P" & vbNewLine & _
            "       Where L.��Ŀid = I.ID And I.ID = P.������Ŀid And L.����id = [2] And ( P.�������=1 Or P.ȡֵ���� is Not null)) I," & vbNewLine & _
            "     (Select �ʿ�Ʒid, ��Ŀid, ��ֵ, Sd, ��Ŀqc��, ����qc��, ����, �ʿ�ȡֵ From �����ʿ�Ʒ��Ŀ Where �ʿ�Ʒid = [1]) K," & vbNewLine & _
            "     (Select X.�ʿ�Ʒid, X.��Ŀid, X.��ֵ, X.Sd" & vbNewLine & _
            "       From �����ʿ�Ʒ R, �����ʿؾ�ֵ X" & vbNewLine & _
            "       Where R.ID = X.�ʿ�Ʒid(+) And R.��ʼ���� = X.��ʼ����(+) And R.ID = [1]) X" & vbNewLine & _
            "Where I.ID = K.��Ŀid(+) And K.�ʿ�Ʒid = X.�ʿ�Ʒid(+) And K.��Ŀid = X.��Ŀid(+)" & vbNewLine & _
            "Order By I.����"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, lngResID, lngDevId)
    Set Me.vfgList.DataSource = rsTemp: Call setListFormat(True)
    Call Refresh����(lngResID)
    Me.Tag = IIf(blnAdd, "����", "�޸�"): Call Form_Resize
    Me.txt����.SetFocus
    Call Show�ο�(vfgList.Row)
    zlEditStart = True: Exit Function

errHand:
    
    If ErrCenter() = 1 Then
    Resume
    End If
    Call SaveErrLog
    zlEditStart = False: Exit Function
End Function

Public Sub zlEditCancel()
    '���ܣ��������ڽ��еı༭
    Me.Tag = "": Call Form_Resize
    Call Me.zlRefresh(mlngResId, mlngDevId)
End Sub

Public Function zlEditSave() As Long
    '���ܣ��������ڽ��еı༭,���������ڱ༭��Ŀid,����ʧ�ܷ���0
    Dim lngNewId As Long
    Dim strLists As String, strItems As String
    Dim strValList As String, strCurValue As String, lngValCount As Long
    Dim dblValue As Double
         
    If ActiveControl = txt����ֵ And mblnEditRow = True Then
        If Chk����ֵ(txt����ֵ.Text) Then
            With vfgList
                If .TextMatrix(.Row, mCol.����ֵ) <> txt����ֵ Then
                    .TextMatrix(.Row, mCol.����ֵ) = txt����ֵ
                End If
                mblnEditRow = False
            End With
        Else
            Exit Function
        End If
    End If
    strLists = ""
    With Me.vfgList
        For lngCount = .FixedRows To .Rows - 1
            If .Cell(flexcpChecked, lngCount, mCol.ѡ��) = flexChecked Then
                dblValue = Val(.TextMatrix(lngCount, mCol.�ο���ֵ))
                If dblValue > 999999 Or Val(dblValue * 10000) - Int(Val(dblValue * 10000)) > 0 Then
                    MsgBox "��" & lngCount & "�С��ο���ֵ������(̫��򾫶�̫�ߣ���", vbInformation, gstrSysName
                    .SetFocus: zlEditSave = 0: Exit Function
                End If
                dblValue = Val(.TextMatrix(lngCount, mCol.�ο�SD))
                If dblValue > 999999 Or Val(dblValue * 10000) - Int(Val(dblValue * 10000)) > 0 Then
                    MsgBox "��" & lngCount & "�С��ο�SD������(̫��򾫶�̫�ߣ���", vbInformation, gstrSysName
                    .SetFocus: zlEditSave = 0: Exit Function
                End If
                If Me.chk�Ƕ�ֵ.Value = vbUnchecked Then
                    dblValue = Val(.TextMatrix(lngCount, mCol.Ԥ���ֵ))
                    If dblValue > 999999 Or Val(dblValue * 10000) - Int(Val(dblValue * 10000)) > 0 Then
                        MsgBox "��" & lngCount & "�С�Ԥ���ֵ������(̫��򾫶�̫�ߣ���", vbInformation, gstrSysName
                        .SetFocus: zlEditSave = 0: Exit Function
                    End If
                    dblValue = Val(.TextMatrix(lngCount, mCol.Ԥ��SD))
                    If dblValue > 999999 Or Val(dblValue * 10000) - Int(Val(dblValue * 10000)) > 0 Then
                        MsgBox "��" & lngCount & "�С�Ԥ��SD������(̫��򾫶�̫�ߣ���", vbInformation, gstrSysName
                        .SetFocus: zlEditSave = 0: Exit Function
                    End If
                End If
                
                strItems = .TextMatrix(lngCount, mCol.ID)
                If Trim(.TextMatrix(lngCount, mCol.ȡֵ����)) = "" Or Trim(.TextMatrix(lngCount, mCol.�ʿ�ȡֵ)) <> "" Then
                    strItems = strItems & ";" & Val(.TextMatrix(lngCount, mCol.�ο���ֵ))
                    strItems = strItems & ";" & Val(.TextMatrix(lngCount, mCol.�ο�SD))
                    If Me.chk�Ƕ�ֵ.Value = vbUnchecked Then
                        strItems = strItems & ";" & Val(.TextMatrix(lngCount, mCol.Ԥ���ֵ))
                        strItems = strItems & ";" & Val(.TextMatrix(lngCount, mCol.Ԥ��SD))
                    Else
                        strItems = strItems & ";;"
                    End If
                Else
                    strValList = Trim(Me.vfgList.TextMatrix(lngCount, mCol.ȡֵ����))
                    
                    strCurValue = Trim(Me.vfgList.TextMatrix(lngCount, mCol.�ο���ֵ))
                    strItems = strItems & ";"
                    For lngValCount = 0 To UBound(Split(strValList, ";"))
                        If strCurValue = Split(strValList, ";")(lngValCount) Then strItems = strItems & lngValCount + 1: Exit For
                    Next
                    strItems = strItems & ";" & Val(.TextMatrix(lngCount, mCol.�ο�SD))
                    If Me.chk�Ƕ�ֵ.Value = vbUnchecked Then
                        strCurValue = Trim(Me.vfgList.TextMatrix(lngCount, mCol.Ԥ���ֵ))
                        strItems = strItems & ";"
                        For lngValCount = 0 To UBound(Split(strValList, ";"))
                            If strCurValue = Split(strValList, ";")(lngValCount) Then strItems = strItems & lngValCount + 1: Exit For
                        Next
                        strItems = strItems & ";" & Val(.TextMatrix(lngCount, mCol.Ԥ��SD))
                    Else
                        strItems = strItems & ";;"
                    End If
                
                End If
                strItems = strItems & ";" & Left(Trim(.TextMatrix(lngCount, mCol.��Ŀqc��)), 8)
                strItems = strItems & ";" & Left(Trim(.TextMatrix(lngCount, mCol.����qc��)), 8)
                strItems = strItems & ";" & Trim(.TextMatrix(lngCount, mCol.����))
                strItems = strItems & ";" & Replace(Trim(.TextMatrix(lngCount, mCol.ȡֵ����)), ";", "��")
                strItems = strItems & ";" & Replace(Trim(.TextMatrix(lngCount, mCol.����ֵ)), ";", "��")
                strItems = strItems & ";" & IIf(mlng�������� = 2, Trim(.TextMatrix(lngCount, mCol.�ʿ�ȡֵ)), "")
                strLists = strLists & "|" & strItems
            End If
        Next
    End With
    If strLists <> "" Then strLists = Mid(strLists, 2)
    
    'һ�����Լ��
    If Trim(Me.txt����.Text) = "" Then
        MsgBox "���������ţ�", vbInformation, gstrSysName
        Me.txt����.SetFocus: zlEditSave = 0: Exit Function
    End If
    If LenB(StrConv(Trim(Me.txt����.Text), vbFromUnicode)) > Me.txt����.MaxLength Then
        MsgBox "���ų��������" & Me.txt����.MaxLength & "���ַ�����", vbInformation, gstrSysName
        Me.txt����.SetFocus: zlEditSave = 0: Exit Function
    End If
    If Trim(Me.txt����.Text) = "" Then
        MsgBox "���������ƣ�", vbInformation, gstrSysName
        Me.txt����.SetFocus: zlEditSave = 0: Exit Function
    End If
    If LenB(StrConv(Trim(Me.txt����.Text), vbFromUnicode)) > Me.txt����.MaxLength Then
        MsgBox "���Ƴ��������" & Me.txt����.MaxLength & "���ַ���ȳ����֣���", vbInformation, gstrSysName
        Me.txt����.SetFocus: zlEditSave = 0: Exit Function
    End If
    If LenB(StrConv(Trim(Me.txtŨ��.Text), vbFromUnicode)) > Me.txtŨ��.MaxLength Then
        MsgBox "Ũ�ȳ��������" & Me.txtŨ��.MaxLength & "���ַ���ȳ����֣���", vbInformation, gstrSysName
        Me.txtŨ��.SetFocus: zlEditSave = 0: Exit Function
    End If
    If Me.cboˮƽ.ListIndex = -1 Then
        MsgBox "����ָ��Ũ�ȵ�ˮƽ��ǣ�(���޷�ָ�����������δ�����������ʿ�ˮƽ��)", vbInformation, gstrSysName
        Me.cboˮƽ.SetFocus: zlEditSave = 0: Exit Function
    End If
    If LenB(StrConv(Trim(Me.txt�걾��.Text), vbFromUnicode)) > Me.txt�걾��.MaxLength Then
        MsgBox "�걾�ų��������" & Me.txt�걾��.MaxLength & "���ַ�����", vbInformation, gstrSysName
        Me.txt�걾��.SetFocus: zlEditSave = 0: Exit Function
    End If
    
    '���ݱ��������֯
    
    gstrSql = "'" & Trim(Me.txt����.Text) & "','" & Trim(Me.txt����.Text) & "'"
    gstrSql = gstrSql & ",To_Date('" & Format(Me.dtp��ʼ����.Value, "yyyy-MM-dd") & "','yyyy-mm-dd')"
    gstrSql = gstrSql & ",To_Date('" & Format(Me.dtp��������.Value, "yyyy-MM-dd") & "','yyyy-mm-dd')"
    gstrSql = gstrSql & ",'" & Trim(Me.txtŨ��.Text) & "'," & Me.cboˮƽ.ListIndex + 1 & ",'" & Trim(Me.cbo����.Text) & "'," & mlngDevId
    gstrSql = gstrSql & "," & IIf(Me.chk�Ƕ�ֵ.Value = vbChecked, 1, 0) & ",'" & Trim(Me.txt�걾��.Text) & "'"
    
    If Me.Tag = "����" Then
        lngNewId = zldatabase.GetNextId("�����ʿ�Ʒ")
        gstrSql = "Zl_�����ʿ�Ʒ_Edit(1," & lngNewId & "," & gstrSql & ",'" & strLists & "','" & Trim(cbo�Լ�.Text) & "','" & Trim(cboУ׼��.Text) & "')"
    Else
        gstrSql = "Zl_�����ʿ�Ʒ_Edit(2," & mlngResId & "," & gstrSql & ",'" & strLists & "','" & Trim(cbo�Լ�.Text) & "','" & Trim(cboУ׼��.Text) & "')"
    End If
    
    Err = 0: On Error GoTo errHand
    Call SQLTest(App.ProductName, Me.Caption, gstrSql): gcnOracle.Execute gstrSql, , adCmdStoredProc: Call SQLTest
    
    If Me.Tag = "����" Then mlngResId = lngNewId
    Me.Tag = "": Call Form_Resize
    zlEditSave = mlngResId: Exit Function
    
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlEditSave = 0: Exit Function
End Function

Private Sub Refresh����(ByVal lngResID As Long)
    '��������Ŀ�Ķ������գ��vfgList�Ķ�Ӧ��
    Dim rsQua As New ADODB.Recordset '������ֵ
    Dim lngItemID As Long
    Dim iRow As Integer
    '������Ŀ�Ͷ�����Ŀ��ȡ������ֵ
    On Error GoTo errHand
    With Me.vfgList
        If Me.vfgList.Rows > 2 Then
            For iRow = 2 To Me.vfgList.Rows - 1
                If Val(vfgList.TextMatrix(iRow, mCol.�������)) <> 1 Then
                    lngItemID = Val(vfgList.TextMatrix(iRow, mCol.ID))
                    gstrSql = "Select ȡֵ����,����ֵ From �����ʿ�Ʒ��Ŀ Where �ʿ�ƷID=[1] And ��ĿID=[2]"
                    Set rsQua = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, lngResID, lngItemID)
                    Do Until rsQua.EOF
                        If "" & rsQua!ȡֵ���� <> "" Then vfgList.TextMatrix(iRow, mCol.ȡֵ����) = "" & rsQua!ȡֵ����
                        vfgList.TextMatrix(iRow, mCol.����ֵ) = "" & rsQua!����ֵ
                        rsQua.MoveNext
                    Loop
                End If
            Next
        End If
    End With
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Show�ο�(ByVal intRow As Integer)

    mblnEditRow = False
    lblȡֵ����.Visible = False: txtȡֵ����.Visible = False
    lbl����ֵ.Visible = False: txt����ֵ.Visible = False
    txtȡֵ����.Text = "": txt����ֵ = ""
    With vfgList
        If intRow > 1 And intRow < .Rows And .Cols >= 15 Then
            
            If Val(.TextMatrix(intRow, mCol.�������)) <= 1 Then
                txtȡֵ����.Text = "": txt����ֵ = ""
            Else
                
                txtȡֵ����.Text = .TextMatrix(intRow, mCol.ȡֵ����)
                txt����ֵ = .TextMatrix(intRow, mCol.����ֵ)
                lblȡֵ����.Visible = True: txtȡֵ����.Visible = True
                lbl����ֵ.Visible = True: txt����ֵ.Visible = True
            End If
        End If
    End With
End Sub


'--------------------------------------------
'����Ϊ����ؼ���Ӧ�¼�
'--------------------------------------------
Private Sub cbo����_GotFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub cbo����_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
End Sub

Private Sub cboˮƽ_GotFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub cboˮƽ_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
End Sub

Private Sub chkHide_Click()
    If Me.chkHide.Value = vbChecked Then
        Me.vfgList.ColWidth(mCol.������) = 0: Me.vfgList.ColHidden(mCol.������) = True
    Else
        Me.vfgList.ColWidth(mCol.������) = 2000: Me.vfgList.ColHidden(mCol.������) = False
    End If
End Sub

Private Sub chk�Ƕ�ֵ_GotFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub chk�Ƕ�ֵ_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
End Sub

Private Sub dtp��������_GotFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub dtp��ʼ����_GotFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub Form_Load()
    Dim rsTemp As New ADODB.Recordset
    
    mlngResId = 0: mlngDevId = 0
    If Val(zldatabase.GetPara("����������", glngSys, 1062, 1)) = 0 Then
        Me.chkHide.Value = vbUnchecked
    Else
        Me.chkHide.Value = vbChecked
    End If
    
    Err = 0: On Error GoTo errHand
    '�ֶγ�������
    gstrSql = "Select ����, ����, Ũ��, ����, �걾�� From �����ʿ�Ʒ Where ID = [1]"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngResId)
    With rsTemp
        Me.txt����.MaxLength = .Fields("����").DefinedSize
        Me.txt����.MaxLength = .Fields("����").DefinedSize
        Me.txtŨ��.MaxLength = .Fields("Ũ��").DefinedSize
        Me.txt�걾��.MaxLength = .Fields("�걾��").DefinedSize
    End With
    
    '�ʿؼ��鷽��
    mstr���� = "|"
    gstrSql = "Select ���� From �ʿؼ��鷽�� Order By ����"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption)
    With rsTemp
        Me.cbo����.Clear
        Do While Not .EOF
            Me.cbo����.AddItem "" & Trim(!����)
            mstr���� = mstr���� & "|" & Trim(!����)
            .MoveNext
        Loop
        If Me.cbo����.ListCount > 0 Then Me.cbo����.ListIndex = 0
    End With
    
    Call setListFormat
    Exit Sub

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    Me.picEdit.Height = Me.ScaleHeight
    Me.chkHide.Left = Me.ScaleWidth - Me.chkHide.Width - 90
    With Me.vfgList
        .Width = Me.ScaleWidth - .Left - 90
        .Height = Me.ScaleHeight - .Top - 90
    End With
    If Me.Tag <> "" Then
        Me.picEdit.Enabled = True: Me.picEdit.BackColor = RGB(250, 250, 250)
        Me.vfgList.Editable = flexEDKbd: Me.vfgList.FocusRect = flexFocusHeavy
    Else
        Me.picEdit.Enabled = False: Me.picEdit.BackColor = Me.BackColor
        Me.vfgList.Editable = flexEDNone: Me.vfgList.FocusRect = flexFocusNone
    End If
    Me.chk�Ƕ�ֵ.BackColor = Me.picEdit.BackColor
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Me.chkHide.Value = vbChecked Then
        Call zldatabase.SetPara("����������", 1, glngSys, 1062)
    Else
        Call zldatabase.SetPara("����������", 0, glngSys, 1062)
    End If
End Sub

Private Sub txt�걾��_GotFocus()
    Me.txt�걾��.SelStart = 0: Me.txt�걾��.SelLength = 1000
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt�걾��_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then Exit Sub
        If KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then Exit Sub
        If InStr(1, ",-", Chr(KeyAscii)) > 0 Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt����ֵ_KeyPress(KeyAscii As Integer)
    mblnEditRow = True
End Sub

Private Function Chk����ֵ(ByVal str����ֵ As String) As Boolean

    Dim varֵ As Variant
    Dim var���� As Variant
    Dim i As Integer
    
    If str����ֵ <> "" Then
        var���� = Split(str����ֵ, ";")
        varֵ = Split(str����ֵ, ";")
    
        If UBound(var����) <> UBound(varֵ) Then
            MsgBox "����ֵ����Ŀ��ʽ��ȡֵ���еĸ�ʽ��һ�£����������ã�", vbQuestion, gstrSysName
            Exit Function
        End If
        For i = LBound(varֵ) To UBound(varֵ)
            If Not IsNumeric(varֵ(i)) Then
                MsgBox "����ֵ�У��ֺ��м������ӦΪ���֣����������ã�", vbQuestion, gstrSysName
                Exit Function
            End If
        Next
    End If
    
    Chk����ֵ = True
End Function
Private Sub txt����_GotFocus()
    Me.txt����.SelStart = 0: Me.txt����.SelLength = 1000
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(Trim(GCST_INVALIDCHAR), Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txtŨ��_GotFocus()
    Me.txtŨ��.SelStart = 0: Me.txtŨ��.SelLength = 1000
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txtŨ��_KeyPress(KeyAscii As Integer)
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
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt����ֵ_Validate(Cancel As Boolean)
    If Not Chk����ֵ(txt����ֵ.Text) Then
        Cancel = True
    Else
        With vfgList
            If .TextMatrix(.Row, mCol.����ֵ) <> txt����ֵ Then
                .TextMatrix(.Row, mCol.����ֵ) = txt����ֵ
            End If
            mblnEditRow = False
        End With
    End If
End Sub

Private Sub vfgList_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim strLists As String, strValue As String
    
    If Col <> mCol.�ο���ֵ And Col <> mCol.Ԥ���ֵ Then Exit Sub
    If Trim(Me.vfgList.TextMatrix(Row, Col)) = "" Then Exit Sub
      
    strLists = Trim(Me.vfgList.TextMatrix(Row, mCol.ȡֵ����))
    strValue = Trim(Me.vfgList.TextMatrix(Row, Col))
    If Trim(Me.vfgList.TextMatrix(Row, mCol.�ʿ�ȡֵ)) <> "" Then strLists = ""
    
    If strLists = "" Then Exit Sub
    For lngCount = 0 To UBound(Split(strLists, ";"))
        If strValue = Split(strLists, ";")(lngCount) Then Exit Sub
    Next
    Me.vfgList.TextMatrix(Row, Col) = ""
    
    strValue = "����ĿΪ�붨����Ŀ��" & IIf(Col = mCol.�ο���ֵ, "��ֵ", "��ֵ") & "���������ȡֵ����(" & strLists & ")Ҫ��"
    MsgBox strValue, vbInformation, gstrSysName
End Sub

Private Sub vfgList_DblClick()
    If Me.vfgList.MouseRow < Me.vfgList.FixedRows Then Exit Sub
    If Me.Tag = "" Then Exit Sub
    With Me.vfgList
        If .Row < .FixedRows And .Row > .Rows - 1 Then Exit Sub
        If .Cell(flexcpChecked, .Row, mCol.ѡ��) = flexChecked Then
            .Cell(flexcpChecked, .Row, mCol.ѡ��) = flexUnchecked
        Else
            .Cell(flexcpChecked, .Row, mCol.ѡ��) = flexChecked
        End If
    End With
End Sub

Private Sub vfgList_EnterCell()
     vfgList.Select vfgList.Row, vfgList.Col
End Sub

Private Sub vfgList_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If KeyAscii <> vbKeySpace Or Me.vfgList.Col > mCol.Ӣ���� Then Exit Sub
    Call vfgList_DblClick
End Sub

Private Sub vfgList_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If Col < mCol.�ο���ֵ Then Exit Sub
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22, vbKeyReturn: Exit Sub
    Case Else
        Select Case Col
        Case mCol.�ο���ֵ, mCol.Ԥ���ֵ
            With Me.vfgList
                If Trim(.TextMatrix(.Row, mCol.ȡֵ����)) <> "" Then Exit Sub
                If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = Asc(".") Then Exit Sub
            End With
        Case mCol.�ο�SD, mCol.Ԥ��SD
            If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = Asc(".") Then Exit Sub
        Case mCol.��Ŀqc��, mCol.����qc��
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
            If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then Exit Sub
'        Case mCol.����
'            Exit Sub
        End Select
    End Select
    KeyAscii = 0
End Sub

Private Sub vfgList_LeaveCell()
    'Call vfgList.Select(vfgList.Row, vfgList.Col)
End Sub

Private Sub vfgList_SelChange()
    If mblnEditRow Then
        
    End If
    Call Show�ο�(vfgList.Row)
End Sub

Private Sub vfgList_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col < mCol.�ο���ֵ Then Cancel = True: Exit Sub
    If Row < Me.vfgList.FixedRows Then Cancel = True: Exit Sub
    If Me.chk�Ƕ�ֵ.Value = vbChecked And (Col = mCol.Ԥ���ֵ Or Col = mCol.Ԥ��SD) Then Cancel = True: Exit Sub
End Sub
