VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmPatiFind 
   AutoRedraw      =   -1  'True
   Caption         =   "���Ҳ���"
   ClientHeight    =   5970
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7230
   Icon            =   "frmPatiFind.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   7230
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraInfo 
      Caption         =   " ������Ϣ "
      Height          =   2655
      Left            =   120
      TabIndex        =   17
      Top             =   30
      Width           =   5565
      Begin VB.TextBox txtIC�� 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   3975
         MaxLength       =   50
         TabIndex        =   8
         Top             =   1530
         Width           =   1455
      End
      Begin VB.TextBox txtҽ���� 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   3975
         MaxLength       =   20
         TabIndex        =   10
         Top             =   1900
         Width           =   1455
      End
      Begin VB.TextBox txtValue 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1065
         MaxLength       =   20
         TabIndex        =   9
         Top             =   1900
         Width           =   1830
      End
      Begin VB.ComboBox cbo���䵥λ 
         Height          =   300
         Left            =   2310
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1095
         Width           =   580
      End
      Begin MSComCtl2.DTPicker dtp����E 
         Height          =   300
         Left            =   3960
         TabIndex        =   12
         Top             =   2265
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   185925635
         CurrentDate     =   37401
      End
      Begin MSComCtl2.DTPicker dtp����B 
         Height          =   300
         Left            =   1065
         TabIndex        =   11
         Top             =   2265
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   529
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   185925635
         CurrentDate     =   37401
      End
      Begin VB.TextBox txt���֤ 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1065
         MaxLength       =   18
         TabIndex        =   7
         Top             =   1500
         Width           =   1830
      End
      Begin MSComCtl2.DTPicker dtp���� 
         Height          =   300
         Left            =   3975
         TabIndex        =   6
         Top             =   1095
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   185925635
         CurrentDate     =   37401
      End
      Begin VB.TextBox txtOld 
         Height          =   300
         Left            =   1065
         MaxLength       =   10
         TabIndex        =   4
         Top             =   1095
         Width           =   1215
      End
      Begin VB.ComboBox cbo�Ա� 
         Height          =   300
         Left            =   3975
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   690
         Width           =   1455
      End
      Begin VB.TextBox txt���� 
         Height          =   300
         Left            =   1065
         MaxLength       =   100
         TabIndex        =   2
         Top             =   690
         Width           =   1830
      End
      Begin VB.TextBox txt����� 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   3975
         MaxLength       =   18
         TabIndex        =   1
         Top             =   285
         Width           =   1455
      End
      Begin VB.TextBox txt����ID 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1065
         TabIndex        =   0
         Top             =   285
         Width           =   1830
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "IC����"
         Height          =   180
         Left            =   3360
         TabIndex        =   31
         Top             =   1590
         Width           =   540
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ҽ����"
         Height          =   180
         Left            =   3375
         TabIndex        =   30
         Top             =   1965
         Width           =   540
      End
      Begin VB.Label lblKind 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���￨��"
         Height          =   180
         Left            =   270
         TabIndex        =   29
         Top             =   1965
         Width           =   720
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��"
         Height          =   180
         Left            =   3285
         TabIndex        =   26
         Top             =   2325
         Width           =   180
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ϴξ���"
         Height          =   180
         Left            =   270
         TabIndex        =   25
         Top             =   2325
         Width           =   720
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���֤��"
         Height          =   180
         Left            =   270
         TabIndex        =   24
         Top             =   1560
         Width           =   720
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������"
         Height          =   180
         Left            =   3195
         TabIndex        =   23
         Top             =   1155
         Width           =   720
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   630
         TabIndex        =   22
         Top             =   1155
         Width           =   360
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�Ա�"
         Height          =   180
         Left            =   3555
         TabIndex        =   21
         Top             =   750
         Width           =   360
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   630
         TabIndex        =   20
         Top             =   750
         Width           =   360
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�����"
         Height          =   180
         Left            =   3375
         TabIndex        =   19
         Top             =   345
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����ID"
         Height          =   180
         Left            =   450
         TabIndex        =   18
         Top             =   345
         Width           =   540
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshPati 
      Height          =   2820
      Left            =   120
      TabIndex        =   14
      ToolTipText     =   "˫����س��鿴����ϸ����Ϣ"
      Top             =   3045
      Width           =   7005
      _ExtentX        =   12356
      _ExtentY        =   4974
      _Version        =   393216
      FixedCols       =   0
      RowHeightMin    =   250
      BackColorBkg    =   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLinesFixed  =   1
      AllowUserResizing=   1
      MousePointer    =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmPatiFind.frx":058A
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.PictureBox picCmd 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1980
      Left            =   5880
      ScaleHeight     =   1980
      ScaleWidth      =   1275
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   120
      Width           =   1275
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "ȡ��(&C)"
         Height          =   350
         Left            =   75
         TabIndex        =   16
         Top             =   1020
         Width           =   1100
      End
      Begin VB.CommandButton cmdSel 
         Caption         =   "ѡ��(&S)"
         Height          =   350
         Left            =   75
         TabIndex        =   15
         Top             =   600
         Width           =   1100
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "����(&F)"
         Height          =   350
         Left            =   75
         TabIndex        =   13
         Top             =   15
         Width           =   1100
      End
   End
   Begin VB.Label lblInfo 
      BackColor       =   &H00707070&
      Caption         =   " ���˲��ҽ��"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   120
      TabIndex        =   27
      Top             =   2835
      Width           =   6990
   End
   Begin VB.Menu mnuPop 
      Caption         =   "ҽ�ƿ�ѡ��"
      Visible         =   0   'False
      Begin VB.Menu mnuPopuItems 
         Caption         =   ""
         Index           =   0
      End
   End
End
Attribute VB_Name = "frmPatiFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������
Public mlng����ID As Long '���ڲ���:����ID
'-----------------------------------------------------
'���㿨���
Private mcllBrushCard As Collection
Private Type Tp_CardSquare
    blnȱʡ�������� As Boolean
    lngȱʡ�����ID As Long
    intȱʡ���ų��� As Integer
End Type
Private mTyCard As Tp_CardSquare
'-----------------------------------------------------
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
'���ܣ����ݵ�ǰ�����������Ҳ���
    Dim strSQL As String, i As Integer, rsTmp As ADODB.Recordset
    Dim DateB As Date, DateE As Date, strMCAccount As String, str����Ժ As String
    Dim lng����ID As Long, lng�����ID As Long, strErrMsg As String, strPassWord As String
    Dim strKind As String
    
    If Trim(txt����ID.Text) <> "" Then
        strSQL = strSQL & " And ����ID=[1]"
        lng����ID = Val(Trim(txt����ID.Text))
    End If
    If Trim(txt�����.Text) <> "" Then
        strSQL = strSQL & " And �����=[2]"
    End If
    If Trim(txt����.Text) <> "" Then
        If InStr(1, "ABCDEFGHIJKLMNOPQRSTUVWXYZ", UCase(Left(txt����.Text, 1))) > 0 Then
            strSQL = strSQL & " And Upper(����) Like [3]"
        Else
            strSQL = strSQL & " And ���� Like [4]"
        End If
    End If
    If cbo�Ա�.Text <> "" Then
        strSQL = strSQL & " And �Ա�=[5]"
    End If
    If Trim(txtOld.Text) <> "" Then
        strSQL = strSQL & " And ����=[6]"
    End If
    If Not IsNull(dtp����.Value) Then
        strSQL = strSQL & " And ��������=[7]"
    End If
    If Trim(txt���֤.Text) <> "" Then
        strSQL = strSQL & " And ���֤��=[8]"
    End If
    If Trim(txtValue.Text) <> "" Then
        strKind = mnuPopuItems(Val(lblKind.Tag)).Tag
        Select Case strKind
        Case "����"
        Case Else
            '��������,��ȡ��صĲ���ID
            '����|�����|ˢ����־|�����ID|���ų���|ȱʡ��־(1-��ǰȱʡ;0-��ȱʡ)|
            '�Ƿ�����ʻ�(1-�����ʻ�;0-�������ʻ�)|��������(�ڼ�λ���ڼ�λ����,��Ϊ������)
            '��7λ��,��ֻ��������,��Ȼȡ������
            lng�����ID = Val(mcllBrushCard(Val(lblKind.Tag) + 1)(3))
            If lng�����ID <> 0 Then
                If gobjSquare.objSquareCard.zlGetPatiID(lng�����ID, Trim(txtValue.Text), False, lng����ID, strPassWord, strErrMsg) = False Then lng����ID = 0
            Else
                If gobjSquare.objSquareCard.zlGetPatiID(strKind, Trim(txtValue.Text), False, lng����ID, _
                    strPassWord, strErrMsg) = False Then lng����ID = 0
            End If
            strSQL = strSQL & " And ����ID=[1]"
        End Select
    End If
    If Not IsNull(dtp����B.Value) And Not IsNull(dtp����E.Value) Then
        If dtp����E.Value <= dtp����B.Value Then
            MsgBox "�ϴξ���Ľ���ʱ�������ڿ�ʼʱ�䣡", vbInformation, gstrSysName
            dtp����E.SetFocus: Exit Sub
        End If
        strSQL = strSQL & " And ����ʱ�� Between [9] And [10]"
    ElseIf Not IsNull(dtp����B.Value) Then
        DateB = CDate(Format(dtp����B.Value, "yyyy-MM-dd 00:00:00"))
        DateE = CDate(Format(dtp����B.Value, "yyyy-MM-dd 23:59:59"))
        strSQL = strSQL & " And ����ʱ�� Between [11] And [12]"
    End If
    
    If Trim(txtIC��.Text) <> "" Then
        strSQL = strSQL & " And IC����=[13]"
    End If
    
    If Trim(txtҽ����.Text) <> "" Then
        strSQL = strSQL & " And ҽ����=[15]"
    End If
    
    If strSQL = "" Then
        MsgBox "����������һ������������", vbInformation, gstrSysName
        txt����.SetFocus: Exit Sub
    End If
    
    strMCAccount = Trim(txtҽ����.Text)
    
    On Error GoTo errH
    Screen.MousePointer = 11
    str����Ժ = " And Not Exists(Select 1 From ������ҳ Where ����ID=A.����ID And ��ҳID<>0 And ��ҳID=A.��ҳID And Nvl(��������,0)=0 And ��Ժ���� is Null)"
    strSQL = _
        " Select " & _
        " ����ID,�����,�ѱ�,ҽ�Ƹ��ʽ,����,�Ա�,����,To_Char(��������,'YYYY-MM-DD') as ��������," & _
        " ���֤��,�����ص�,��ͥ��ַ,������λ,���,ְҵ,ѧ��,To_Char(����ʱ��,'YYYY-MM-DD HH24:MI') as ����ʱ��" & _
        " From ������Ϣ A" & _
        " Where ͣ��ʱ�� is NULL " & str����Ժ & strSQL & _
        " Order by ����ID"
    Set rsTmp = zldatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, Trim(txt�����.Text), UCase(txt����.Text) & "%", _
                txt����.Text & "%", cbo�Ա�.Text, IIf(IsNumeric(txtOld.Text), txtOld.Text & cbo���䵥λ.Text, txtOld.Text), IIf(IsNull(dtp����.Value), "", dtp����.Value), txt���֤.Text, _
                IIf(IsNull(dtp����B.Value), "", dtp����B.Value), IIf(IsNull(dtp����E.Value), "", dtp����E.Value), _
                DateB, DateE, Trim(txtIC��.Text), "", strMCAccount)
    
    If Not rsTmp.EOF Then
        lblInfo.Caption = " ���˲��ҽ��:�� " & rsTmp.RecordCount & " �����������Ĳ���"
        Set mshPati.DataSource = rsTmp
        For i = 0 To mshPati.Cols - 1
            mshPati.ColAlignmentFixed(i) = 4
        Next
        mshPati.TextMatrix(0, 0) = "����ID": mshPati.ColWidth(0) = 750: mshPati.ColAlignment(0) = 1
        mshPati.TextMatrix(0, 1) = "�����": mshPati.ColWidth(1) = 750: mshPati.ColAlignment(1) = 1
        mshPati.TextMatrix(0, 2) = "�ѱ�": mshPati.ColWidth(2) = 850: mshPati.ColAlignment(2) = 1
        mshPati.TextMatrix(0, 3) = "���ʽ": mshPati.ColWidth(3) = 850: mshPati.ColAlignment(3) = 1
        mshPati.TextMatrix(0, 4) = "����": mshPati.ColWidth(4) = 700: mshPati.ColAlignment(4) = 1
        mshPati.TextMatrix(0, 5) = "�Ա�": mshPati.ColWidth(5) = 500: mshPati.ColAlignment(5) = 4
        mshPati.TextMatrix(0, 6) = "����": mshPati.ColWidth(6) = 500: mshPati.ColAlignment(6) = 1
        mshPati.TextMatrix(0, 7) = "��������": mshPati.ColWidth(7) = 1000: mshPati.ColAlignment(7) = 4
        mshPati.TextMatrix(0, 8) = "���֤��": mshPati.ColWidth(8) = 1600: mshPati.ColAlignment(8) = 1
        mshPati.TextMatrix(0, 9) = "�����ص�": mshPati.ColWidth(9) = 2000: mshPati.ColAlignment(9) = 1
        mshPati.TextMatrix(0, 10) = "��ͥ��ַ": mshPati.ColWidth(10) = 2000: mshPati.ColAlignment(10) = 1
        mshPati.TextMatrix(0, 11) = "������λ": mshPati.ColWidth(11) = 2000: mshPati.ColAlignment(11) = 1
        mshPati.TextMatrix(0, 12) = "���": mshPati.ColWidth(12) = 1000: mshPati.ColAlignment(12) = 1
        mshPati.TextMatrix(0, 13) = "ְҵ": mshPati.ColWidth(13) = 1000: mshPati.ColAlignment(13) = 1
        mshPati.TextMatrix(0, 14) = "ѧ��": mshPati.ColWidth(14) = 500: mshPati.ColAlignment(14) = 1
        mshPati.TextMatrix(0, 15) = "�ϴξ���ʱ��": mshPati.ColWidth(15) = 1600: mshPati.ColAlignment(15) = 4
    Else
        lblInfo.Caption = " ���˲��ҽ��"
        mshPati.Clear
        mshPati.ClearStructure
        mshPati.Cols = 2: mshPati.Rows = 2
        mshPati.FixedCols = 0: mshPati.FixedRows = 1
    End If
    mshPati.Row = 1: mshPati.Col = 0: mshPati.ColSel = mshPati.Cols - 1
    mshPati.TopRow = 1
    Call mshPati_EnterCell
    Screen.MousePointer = 0
    mshPati.SetFocus
    Exit Sub
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Screen.MousePointer = 11
        Resume
    End If
    Call SaveErrLog
End Sub



Private Sub cmdSel_Click()
    If Val(mshPati.TextMatrix(mshPati.Row, 0)) = 0 Then
        MsgBox "û�в�����Ϣ����ѡ��", vbInformation, gstrSysName
        Exit Sub
    End If
    mlng����ID = Val(mshPati.TextMatrix(mshPati.Row, 0))
    Unload Me
End Sub

Private Sub dtp����B_Change()
    If IsNull(dtp����B.Value) Then dtp����E.Value = Null
End Sub

Private Sub dtp����B_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub dtp����E_Change()
    If IsNull(dtp����B.Value) And Not IsNull(dtp����E.Value) Then dtp����E.Value = Null
End Sub

Private Sub dtp����E_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub dtp����_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Call zlCommFun.PressKey(vbKeyTab)
End Sub



Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("'&[]", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    If KeyAscii = 13 And TypeName(ActiveControl) <> "DTPicker" And Not ActiveControl Is mshPati Then KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Form_Load()
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim Datsys As Date
    
    txt����.MaxLength = zlGetPatiInforMaxLen.intPatiName
    
    Call InitMenus
    Call RestoreWinState(Me, App.ProductName)
    mlng����ID = 0
    Datsys = zldatabase.Currentdate
    
    dtp����E.MaxDate = Datsys
    dtp����B.MaxDate = dtp����E.MaxDate
    dtp����B.Value = DateAdd("m", -1, Datsys)
    dtp����E.Value = Datsys
    dtp����B.Value = Null
    dtp����E.Value = Null
    
    dtp����.MaxDate = Datsys
    dtp����.Value = DateAdd("yyyy", -25, Datsys)
    dtp����.Value = Null
    
    Call mshPati_EnterCell
    
    On Error GoTo errH
    strSQL = "Select ���� From �Ա�"
    Call zldatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
    cbo�Ա�.AddItem ""
    Do While Not rsTmp.EOF
        cbo�Ա�.AddItem rsTmp!����
        rsTmp.MoveNext
    Loop
    cbo�Ա�.ListIndex = 0
    
    '���䵥λ
    cbo���䵥λ.AddItem "��"
    cbo���䵥λ.AddItem "��"
    cbo���䵥λ.AddItem "��"
    cbo���䵥λ.ListIndex = 0
    
    txtҽ����.MaxLength = 20
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    If Me.ScaleWidth - picCmd.Width > fraInfo.Left + fraInfo.Width Then
        picCmd.Left = Me.ScaleWidth - picCmd.Width
    Else
        picCmd.Left = fraInfo.Left + fraInfo.Width
    End If
    lblInfo.Width = Me.ScaleWidth - lblInfo.Left * 2
    mshPati.Width = Me.ScaleWidth - mshPati.Left * 2
    
    If Me.ScaleHeight - mshPati.Top - mshPati.Left > 1000 Then
        mshPati.Height = Me.ScaleHeight - mshPati.Top - mshPati.Left
    Else
        mshPati.Height = 1000
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub lblKind_Click()
    PopupMenu mnuPop, 2
End Sub

Private Sub mshPati_DblClick()
    If mshPati.MouseRow > 0 Then Call mshPati_KeyPress(13)
End Sub

Private Sub mshPati_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 And Shift = 2 Then cmdSel_Click
End Sub

Private Sub mshPati_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Val(mshPati.TextMatrix(mshPati.Row, 0)) <> 0 Then
        frmDegreeCard.mlng����ID = Val(mshPati.TextMatrix(mshPati.Row, 0))
        frmDegreeCard.Show 1, Me
    End If
End Sub

Private Sub txtIC��_GotFocus()
    zlControl.TxtSelAll txtIC��
End Sub
Private Sub txtValue_GotFocus()
    zlControl.TxtSelAll txtValue
End Sub
Private Sub txtValue_KeyPress(KeyAscii As Integer)
    Dim blnCard As Boolean
    Dim strKind As String, intKind As Integer, int���ų��� As Long
    Dim bln���� As Boolean
    strKind = mnuPopuItems(Val(lblKind.Tag)).Tag
    intKind = Val(lblKind.Tag) + 1
    bln���� = mcllBrushCard(intKind)(7) <> ""
    txtValue.PasswordChar = IIf(bln����, "*", "")
    'ȡȱʡ��ˢ����ʽ
            '����|�����|ˢ����־|�����ID|���ų���|ȱʡ��־(1-��ǰȱʡ;0-��ȱʡ)|
            '�Ƿ�����ʻ�(1-�����ʻ�;0-�������ʻ�)|��������(�ڼ�λ���ڼ�λ����,��Ϊ������)
            '��7λ��,��ֻ��������,��Ȼȡ������
    Select Case strKind
    Case "����"
           blnCard = zlCommFun.InputIsCard(txtValue, KeyAscii, mTyCard.blnȱʡ��������)
           int���ų��� = mTyCard.intȱʡ���ų��� - 1
    Case "�����"
            If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
            int���ų��� = 0
    Case "ҽ����"
            int���ų��� = 0
    Case Else
        blnCard = zlCommFun.InputIsCard(txtValue, KeyAscii, bln����)
        int���ų��� = mcllBrushCard(intKind)(4)
    End Select
    If int���ų��� > 0 Then
         'ˢ����ϻ���������س�
         If blnCard And Len(txtValue.Text) = int���ų��� - 1 And KeyAscii <> 8 Or KeyAscii = 13 And Trim(txtValue.Text) <> "" Then
             If KeyAscii <> 13 Then
                 txtValue.Text = txtValue.Text & Chr(KeyAscii)
                 txtValue.SelStart = Len(txtValue.Text)
             End If
             KeyAscii = 0
             Call cmdFind_Click
             If mshPati.Rows > 1 Then
                If mshPati.TextMatrix(1, 0) = "" Then
                   txtValue.SetFocus
                   zlControl.TxtSelAll txtValue
                End If
            End If
        End If
    End If
End Sub

Private Sub txtҽ����_GotFocus()
    zlControl.TxtSelAll txtҽ����
End Sub

Private Sub txtIC��_Validate(Cancel As Boolean)
    txtIC��.Text = UCase(Trim(txtIC��.Text))
End Sub

Private Sub txtValue_Validate(Cancel As Boolean)
    txtValue.Text = UCase(Trim(txtValue.Text))
End Sub

Private Sub txtҽ����_Validate(Cancel As Boolean)
    txtҽ����.Text = UCase(Trim(txtҽ����.Text))
End Sub


Private Sub txt����ID_GotFocus()
    zlControl.TxtSelAll txt����ID
End Sub

Private Sub txt����ID_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txt�����_GotFocus()
    zlControl.TxtSelAll txt�����
End Sub

Private Sub txt�����_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtOld_GotFocus()
    zlControl.TxtSelAll txtOld
End Sub

Private Sub txtOld_KeyDown(KeyCode As Integer, Shift As Integer)
    '�����keydown����presskey,�����ٴ���һ��,�������䵥λ
    If KeyCode = vbKeyReturn And Not IsNumeric(txtOld.Text) Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub txtOld_Validate(Cancel As Boolean)
    Dim strTmp As String
    
    strTmp = cbo���䵥λ.Text
    Select Case strTmp
        Case "��"
            If Val(txtOld.Text) > 200 Then Cancel = True: Exit Sub
        Case "��"
            If Val(txtOld.Text) > 2400 Then Cancel = True: Exit Sub
        Case "��"
            If Val(txtOld.Text) > 73000 Then Cancel = True: Exit Sub
        Case Else
            Exit Sub
    End Select
End Sub

Private Sub txt���֤_GotFocus()
    zlControl.TxtSelAll txt���֤
    
End Sub

Private Sub txt���֤_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txt����_GotFocus()
    zlControl.TxtSelAll txt����
End Sub

Private Sub mshPati_EnterCell()
    Dim i As Integer, blnPre As Boolean
    Dim intRow As Integer, intCol As Integer
    
    blnPre = mshPati.Redraw
    intRow = mshPati.Row: intCol = mshPati.Col
    mshPati.Redraw = False
    
    For i = 0 To mshPati.Cols - 1
        mshPati.Col = i
        mshPati.CellBackColor = mshPati.BackColorSel
        mshPati.CellForeColor = mshPati.ForeColorSel
    Next
    
    mshPati.Row = intRow:  mshPati.Col = intCol
    mshPati.Redraw = blnPre
End Sub

Private Sub mshPati_LeaveCell()
    Dim i As Integer, blnPre As Boolean
    
    blnPre = mshPati.Redraw
    mshPati.Redraw = False
    
    For i = 0 To mshPati.Cols - 1
        mshPati.Col = i
        mshPati.CellBackColor = mshPati.BackColor
        mshPati.CellForeColor = mshPati.ForeColor
    Next
    mshPati.Redraw = blnPre
End Sub

Private Sub InitMenus()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��̬������ص�ҽ�ƿ����˵�
    '����:���˺�
    '����:2011-10-21 15:29:07
    '����:42315
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varData As Variant, varTemp As Variant, strKind As String
    Dim i As Integer, ObjItem As Menu, intDefaultKind As Integer
    Set mcllBrushCard = New Collection
    strKind = "��|���￨|0|0|18|0|0||"
    If Not gobjSquare.objSquareCard Is Nothing Then
        strKind = gobjSquare.objSquareCard.zlGetIDKindStr(strKind)
    End If
    intDefaultKind = 0
    varData = Split(strKind, ";")
    For i = 0 To UBound(varData)
        Set ObjItem = Me.mnuPopuItems(mnuPopuItems.UBound)
        If Not (ObjItem.Caption = "-" Or Trim(ObjItem.Caption) = "" Or Not ObjItem.Visible) Then
            Load mnuPopuItems(mnuPopuItems.UBound + 1)
            Set ObjItem = mnuPopuItems(mnuPopuItems.UBound)
        End If
        varTemp = Split(varData(i), "|")
        'ȡȱʡ��ˢ����ʽ
        '����|�����|ˢ����־|�����ID|���ų���|ȱʡ��־(1-��ǰȱʡ;0-��ȱʡ)|
        '�Ƿ�����ʻ�(1-�����ʻ�;0-�������ʻ�)|��������(�ڼ�λ���ڼ�λ����,��Ϊ������)
        '��7λ��,��ֻ��������,��Ȼȡ������
        mcllBrushCard.Add varTemp, varTemp(1)
        If Val(varTemp(5)) = 1 Then
            intDefaultKind = i
            mTyCard.blnȱʡ�������� = Trim(varTemp(7)) <> ""
            mTyCard.lngȱʡ�����ID = Val(varTemp(3))
            mTyCard.intȱʡ���ų��� = Val(varTemp(4))
        End If
        If i > 9 Then
            ObjItem.Caption = varTemp(1) & IIf(i - 9 > 24, "", "(&" & Chr(64 + i) & ")")
        Else
            ObjItem.Caption = varTemp(1) & "(&" & i & ")"
        End If
        ObjItem.Tag = CStr(varTemp(1))
    Next
    '����ȱʡ���Ҷ���
    mnuPopuItems_Click (intDefaultKind)
End Sub
Private Sub mnuPopuItems_Click(Index As Integer)
    Dim i As Long
    For i = 0 To mnuPopuItems.UBound
        mnuPopuItems(i).Checked = i = Index
    Next
    lblKind.Caption = mnuPopuItems(Index).Tag & "��"
    lblKind.Tag = Index
    lblKind.ToolTipText = mnuPopuItems(Index).Tag
    txtValue.ToolTipText = mnuPopuItems(Index).Tag
End Sub
