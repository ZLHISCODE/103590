VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmQCRedefine 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���¶�ֵ"
   ClientHeight    =   5325
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   7215
   Icon            =   "frmQCRedefine.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5325
   ScaleWidth      =   7215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame fraEdit 
      Height          =   1095
      Left            =   180
      TabIndex        =   18
      Top             =   4140
      Width           =   6915
      Begin VB.TextBox txt�ڼ� 
         Height          =   300
         Left            =   3435
         TabIndex        =   21
         Top             =   705
         Width           =   1500
      End
      Begin VB.CommandButton cmdCusum 
         Cancel          =   -1  'True
         Caption         =   "��ȡ�ۻ�ֵ(&S)"
         Height          =   350
         Left            =   5190
         TabIndex        =   12
         Top             =   0
         Width           =   1710
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "�� ��ӵ�����ֵ�б���(&A)"
         Height          =   350
         Index           =   0
         Left            =   0
         TabIndex        =   13
         Top             =   0
         Width           =   2600
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "�� ɾ��������ֵ(&D)"
         Height          =   350
         Index           =   1
         Left            =   2595
         TabIndex        =   14
         Top             =   0
         Width           =   2600
      End
      Begin VB.TextBox txt��ֵ 
         Height          =   300
         Left            =   5040
         TabIndex        =   9
         Top             =   705
         Width           =   810
      End
      Begin VB.TextBox txtSD 
         Height          =   300
         Left            =   5865
         TabIndex        =   11
         Top             =   705
         Width           =   810
      End
      Begin MSComCtl2.DTPicker dtp���� 
         Height          =   300
         Left            =   1755
         TabIndex        =   6
         Top             =   705
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy��MM��dd��"
         Format          =   39714819
         CurrentDate     =   39110
      End
      Begin MSComCtl2.DTPicker dtp��ʼ���� 
         Height          =   300
         Left            =   90
         TabIndex        =   7
         Top             =   705
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy��MM��dd��"
         Format          =   39714819
         CurrentDate     =   39110
      End
      Begin VB.Label lbl�ڼ� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ڼ�"
         Height          =   180
         Left            =   3480
         TabIndex        =   20
         Top             =   480
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "��������"
         Height          =   180
         Left            =   1770
         TabIndex        =   19
         Top             =   480
         Width           =   720
      End
      Begin VB.Label lbl��ʼ���� 
         AutoSize        =   -1  'True
         Caption         =   "��ʼ����"
         Height          =   180
         Left            =   90
         TabIndex        =   5
         Top             =   495
         Width           =   720
      End
      Begin VB.Label lbl��ֵ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ֵ"
         Height          =   180
         Left            =   5055
         TabIndex        =   8
         Top             =   480
         Width           =   360
      End
      Begin VB.Label lblSD 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��׼��"
         Height          =   180
         Left            =   5880
         TabIndex        =   10
         Top             =   480
         Width           =   540
      End
   End
   Begin VB.ComboBox cbo�ʿ�Ʒ 
      Height          =   300
      Left            =   780
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1050
      Width           =   6240
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Left            =   -45
      TabIndex        =   17
      Top             =   345
      Width           =   7215
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "�ر�(&X)"
      Height          =   350
      Left            =   5970
      TabIndex        =   15
      Top             =   420
      Width           =   1100
   End
   Begin VSFlex8Ctl.VSFlexGrid vfgValue 
      Height          =   2655
      Left            =   180
      TabIndex        =   4
      Top             =   1425
      Width           =   6885
      _cx             =   12144
      _cy             =   4683
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
      BackColorSel    =   16635590
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
      Rows            =   6
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
   Begin VB.Image imgNote 
      Height          =   240
      Left            =   150
      Picture         =   "frmQCRedefine.frx":058A
      Top             =   75
      Width           =   240
   End
   Begin VB.Label lbl�ʿ�Ʒ 
      AutoSize        =   -1  'True
      Caption         =   "�ʿ�Ʒ"
      Height          =   180
      Left            =   180
      TabIndex        =   2
      Top             =   1110
      Width           =   540
   End
   Begin VB.Label lbl��Ŀ 
      AutoSize        =   -1  'True
      Caption         =   "��Ŀ: ####"
      Height          =   180
      Left            =   180
      TabIndex        =   1
      Top             =   787
      Width           =   900
   End
   Begin VB.Label lblNote 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "���û����ָ�������ʿ�Ʒ�ľ�ֵ�ͱ�׼�"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   450
      TabIndex        =   16
      Top             =   90
      Width           =   3600
   End
   Begin VB.Label lbl���� 
      AutoSize        =   -1  'True
      Caption         =   "����: ####"
      Height          =   180
      Left            =   180
      TabIndex        =   0
      Top             =   465
      Width           =   900
   End
End
Attribute VB_Name = "frmQCRedefine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Enum mCol
    �ڼ� = 0: ����: ��ֵ: SD: ��ע
End Enum

Private mlngDevID As Long       '����id
Private mlngItemId As Long      '��Ŀid
Private mdtSysdate As Date      '��ǰʱ��
Private mblnModify As Boolean   '�Ƿ�ִ��

Dim lngCount As Long

'--------------------------------------------
'����Ϊ���幫������
'--------------------------------------------
Public Function ShowMe(frmParent As Form, lngDevId As Long, lngItemID As Long, Optional dtDefault As Date, Optional lngResId As Long) As Boolean
    '���ܣ�����ָ����������Ŀ���ʿ�Ʒ������ʾ���¶�ֵ����
    Dim rsTemp As New ADODB.Recordset
    
    mlngDevID = lngDevId
    mlngItemId = lngItemID
    If Not IsNull(dtDefault) Then Me.dtp����.Value = dtDefault
    
    Err = 0: On Error GoTo ErrHand
    gstrSql = "Select ���� || ', ' || ���� As ������ From �������� Where ID = [1]"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngDevID)
    With rsTemp
        If .RecordCount <= 0 Then MsgBox "ָ�����������ڣ�", vbInformation, gstrSysName: Exit Function
        Me.lbl����.Caption = "����: " & !������
    End With
    
    gstrSql = "Select ���� || ', ' || ������ || ', ' || Ӣ���� As ��Ŀ�� From ����������Ŀ Where ID = [1]"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngItemId)
    With rsTemp
        If .RecordCount <= 0 Then MsgBox "ָ����Ŀ�����ڣ�", vbInformation, gstrSysName: Exit Function
        Me.lbl��Ŀ.Caption = "��Ŀ: " & !��Ŀ��
    End With
    
    gstrSql = "Select Sysdate As ����, M.ID," & vbNewLine & _
        "       M.���� || '-' || M.���� || ' ˮƽ:' || M.ˮƽ ||" & vbNewLine & _
        "        LPad(To_Char(M.��ʼ����, 'yyyy-MM-dd') || ' ' || To_Char(M.��������, 'yyyy-MM-dd'), 200, ' ') As �ʿ�Ʒ" & vbNewLine & _
        "From �����ʿ�Ʒ M, �����ʿ�Ʒ��Ŀ I" & vbNewLine & _
        "Where M.ID = I.�ʿ�Ʒid And M.����id = [1] And I.��Ŀid = [2] And Trunc(Sysdate) Between M.��ʼ���� And" & vbNewLine & _
        "      Trim(M.��������)" & vbNewLine & _
        "Order By M.��ʼ����, M.����"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngDevID, mlngItemId)
    With rsTemp
        Me.cbo�ʿ�Ʒ.Clear
        Do While Not .EOF
            mdtSysdate = !����
            Me.cbo�ʿ�Ʒ.AddItem "" & !�ʿ�Ʒ
            Me.cbo�ʿ�Ʒ.ItemData(Me.cbo�ʿ�Ʒ.NewIndex) = !ID
            .MoveNext
        Loop
        If Me.cbo�ʿ�Ʒ.ListCount = 0 Then MsgBox "��������Ŀ�޵�ǰ��Ч���ʿ�Ʒ��", vbInformation, gstrSysName: Exit Function
        For lngCount = 0 To Me.cbo�ʿ�Ʒ.ListCount - 1
            If Me.cbo�ʿ�Ʒ.ItemData(Me.cbo�ʿ�Ʒ.NewIndex) = lngResId Then Me.cbo�ʿ�Ʒ.ListIndex = Me.cbo�ʿ�Ʒ.NewIndex: Exit For
        Next
        If Me.cbo�ʿ�Ʒ.ListIndex = -1 Then Me.cbo�ʿ�Ʒ.ListIndex = 0
    End With
    
    mblnModify = False
    Me.Show vbModal, frmParent
    ShowMe = mblnModify
    Exit Function

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub cbo�ʿ�Ʒ_Click()
    Dim rsTemp As New ADODB.Recordset
    
    Dim strPeriod As String
    
    If Me.cbo�ʿ�Ʒ.ListIndex = -1 Then Exit Sub
    
    gstrSql = "Select ��ʼ���� From �����ʿ�Ʒ Where Id= [1]"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, CLng(Me.cbo�ʿ�Ʒ.ItemData(Me.cbo�ʿ�Ʒ.ListIndex)))
    Do Until rsTemp.EOF
        Me.dtp��ʼ����.MinDate = rsTemp!��ʼ����
        rsTemp.MoveNext
    Loop
    
    gstrSql = "Select �ڼ�, To_Char(��ʼ����, 'yyyy-MM-dd') As ��ʼ����, ��ֵ, Sd," & vbNewLine & _
            "       ������ || '��'|| To_Char(��������, 'yyyy-MM-dd') || '����' As ��ע" & vbNewLine & _
            "From �����ʿؾ�ֵ" & vbNewLine & _
            "Where �ʿ�Ʒid = [1] And ��Ŀid = [2]" & vbNewLine & _
            "Order By ��ʼ����"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, CLng(Me.cbo�ʿ�Ʒ.ItemData(Me.cbo�ʿ�Ʒ.ListIndex)), mlngItemId)
    With Me.vfgValue
        .Clear
        Set .DataSource = rsTemp
        .ColWidth(mCol.����) = 1100
        For lngCount = .FixedCols To .Cols - 1
            .FixedAlignment(lngCount) = flexAlignCenterCenter
        Next
        For lngCount = .FixedRows To .Rows - 1
            If Left(.TextMatrix(lngCount, mCol.��ֵ), 1) = "." Then .TextMatrix(lngCount, mCol.��ֵ) = "0" & .TextMatrix(lngCount, mCol.��ֵ)
            If Left(.TextMatrix(lngCount, mCol.SD), 1) = "." Then .TextMatrix(lngCount, mCol.SD) = "0" & .TextMatrix(lngCount, mCol.SD)
        Next
        If .Rows > .FixedRows Then .Row = .Rows - 1
    End With
    
    strPeriod = Right(Me.cbo�ʿ�Ʒ.Text, 21)
    Me.dtp����.MinDate = CDate(Left(strPeriod, 10))
    
    If CDate(Right(strPeriod, 10)) < mdtSysdate Then
        Me.dtp����.MaxDate = CDate(Right(strPeriod, 10))
    Else
        Me.dtp����.MaxDate = mdtSysdate
    End If
    Me.dtp��ʼ����.MaxDate = Me.dtp����.MaxDate
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdCusum_Click()
    Dim rsTemp As New ADODB.Recordset
    Dim strAsk As String
    
    gstrSql = "Select Round(Avg(���), 2) As ��ֵ, Round(Stddev(���), 2) As Sd, Count(*) As ����" & vbNewLine & _
            "From (Select Trunc(Q.����ʱ��) As ����," & vbNewLine & _
            "              Avg(zl_Lis_tonumber(Q.�ʿ�ƷID,R.������Ŀid,R.������,R.ID)) As ���" & vbNewLine & _
            "       From �����ʿؼ�¼ Q, ������ͨ��� R,�����ʿر��� T" & vbNewLine & _
            "       Where Q.�걾id = R.����걾id And Q.�ʿ�Ʒid = [1] And R.������Ŀid + 0 = [2] And" & vbNewLine & _
            "             Nvl(R.���ý��,0)=0 And R.ID=T.���ID(+) And Q.����ʱ�� + 0 between To_Date([3], 'yyyy-MM-dd') And  To_Date([4], 'yyyy-MM-dd')+ 1 And Nvl(T.���, 0) <> 2" & vbNewLine & _
            "       Group By Trunc(Q.����ʱ��))"

'    gstrSql = "Select Round(Avg(���), 2) As ��ֵ, Round(Stddev(���), 2) As Sd, Count(*) As ����" & vbNewLine & _
'            "From (Select Trunc(Q.����ʱ��) As ����," & vbNewLine & _
'            "              Avg(zl_Lis_tonumber(Q.�ʿ�ƷID,R.������Ŀid,R.������)) As ���" & vbNewLine & _
'            "       From �����ʿؼ�¼ Q, ������ͨ��� R" & vbNewLine & _
'            "       Where Q.�걾id = R.����걾id And Q.�ʿ�Ʒid = [1] And R.������Ŀid + 0 = [2] And" & vbNewLine & _
'            "             Q.����ʱ�� + 0 <To_Date([4], 'yyyy-MM-dd')+ 1 And Nvl(Q.���ü�¼, 0) = 0" & vbNewLine & _
'            "       Group By Trunc(Q.����ʱ��))"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, _
                    CLng(Me.cbo�ʿ�Ʒ.ItemData(Me.cbo�ʿ�Ʒ.ListIndex)), _
                    mlngItemId, Format(Me.dtp��ʼ����.Value, "yyyy-MM-dd"), Format(Me.dtp����.Value, "yyyy-MM-dd"))
    If rsTemp.RecordCount <= 0 Then MsgBox "�޷���ȡ�ۼƾ�ֵ�ͱ�׼��", vbInformation, gstrSysName: Exit Sub
    
    strAsk = "���ʿ�Ʒ��" & Format(Me.dtp��ʼ����.Value, "yyyy��mm��dd��") & "��" & Format(Me.dtp����.Value, "yyyy��mm��dd��") & "���ۼ�ֵ���£�"
    strAsk = strAsk & vbCrLf & "   ��ֵ: " & IIf(Left(rsTemp!��ֵ, 1) = ".", "0", "") & rsTemp!��ֵ
    strAsk = strAsk & vbCrLf & "   SDֵ: " & IIf(Left(rsTemp!SD, 1) = ".", "0", "") & rsTemp!SD
    If rsTemp!���� < 20 Then
        strAsk = strAsk & vbCrLf & "������Ч�ʿؼ�¼����20�Σ�������ֱ�ӽ��ۼ�ֵ��Ϊ����ֵ��"
    End If
    strAsk = strAsk & vbCrLf & vbCrLf & "Ҫʹ�ø��ۼ�ֵ��"
    If MsgBox(strAsk, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
        Me.txt��ֵ.Text = IIf(Left(rsTemp!��ֵ, 1) = ".", "0", "") & rsTemp!��ֵ
        Me.txtSD.Text = IIf(Left(rsTemp!SD, 1) = ".", "0", "") & rsTemp!SD
        Me.txt�ڼ�.Text = Format(Me.dtp��ʼ����.Value, "yyyyMM")
    End If
    
    Me.txt��ֵ.SetFocus
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdEdit_Click(Index As Integer)
    If Index = 0 Then
        With Me.vfgValue
            If .Rows > .FixedRows Then
                If .TextMatrix(.Rows - 1, mCol.����) >= Format(Me.dtp����.Value, "yyyy-MM-dd") Then
                    MsgBox "�¿���ֵ�Ľ������ڱ�������ϴ����ڣ�", vbInformation, gstrSysName
                    Me.dtp����.SetFocus: Exit Sub
                End If
            End If
            If Val(Trim(Me.txt��ֵ.Text)) = 0 And Val(Trim(Me.txtSD.Text)) = 0 Then
                MsgBox "����ͬʱ���þ�ֵ(x)�ͱ�׼��(SD)��", vbInformation, gstrSysName
                Me.txt��ֵ.SetFocus: Exit Sub
            End If
        End With
        gstrSql = "Zl_�����ʿؾ�ֵ_Edit(1," & Me.cbo�ʿ�Ʒ.ItemData(Me.cbo�ʿ�Ʒ.ListIndex) & "," & mlngItemId
        gstrSql = gstrSql & ",To_Date('" & Format(Me.dtp����.Value, "yyyy-MM-dd") & "','yyyy-MM-dd')"
        gstrSql = gstrSql & "," & Val(Trim(Me.txt��ֵ.Text)) & "," & Val(Trim(Me.txtSD.Text)) & ",'" & Replace(Trim(txt�ڼ�.Text), "'", "") & "')"
    Else
        With Me.vfgValue
            If .Rows <= .FixedRows Then
                MsgBox "�Ѿ�û�п���ֵ��", vbInformation, gstrSysName: Me.dtp����.SetFocus: Exit Sub
            ElseIf .Rows = .FixedRows + 1 Then
                If DateAdd("m", 1, CDate(.TextMatrix(.Rows - 1, mCol.����))) <= Me.dtp����.MaxDate And DateAdd("m", 1, CDate(.TextMatrix(.Rows - 1, mCol.����))) >= Me.dtp����.MinDate Then
                    Me.dtp����.Value = DateAdd("m", 1, CDate(.TextMatrix(.Rows - 1, mCol.����)))
                End If
            Else
                If CDate(.TextMatrix(.Rows - 1, mCol.����)) <= Me.dtp����.MaxDate And CDate(.TextMatrix(.Rows - 1, mCol.����)) >= Me.dtp����.MinDate Then
                    Me.dtp����.Value = CDate(.TextMatrix(.Rows - 1, mCol.����))
                End If
            End If
            Me.txt��ֵ.Text = Val(.TextMatrix(.Rows - 1, mCol.��ֵ))
            Me.txtSD.Text = Val(.TextMatrix(.Rows - 1, mCol.SD))
        End With
        gstrSql = "Zl_�����ʿؾ�ֵ_Edit(2," & Me.cbo�ʿ�Ʒ.ItemData(Me.cbo�ʿ�Ʒ.ListIndex) & "," & mlngItemId & ")"
    End If
    
    Err = 0: On Error GoTo ErrHand
    zldatabase.ExecuteProcedure gstrSql, Me.Caption
    mblnModify = True
    
    Call cbo�ʿ�Ʒ_Click
    If Index = 0 Then
        Me.vfgValue.SetFocus
    Else
        Me.dtp����.SetFocus
    End If
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

