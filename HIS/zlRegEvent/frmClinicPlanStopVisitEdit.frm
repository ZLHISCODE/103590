VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmClinicPlanStopVisitEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ͣ������"
   ClientHeight    =   5205
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5595
   Icon            =   "frmClinicPlanStopVisitEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5438.276
   ScaleMode       =   0  'User
   ScaleWidth      =   5595
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.CheckBox chk����ͣ�� 
      Caption         =   "ͣ�ﲿ�ֺ�Դ"
      Height          =   180
      Left            =   3180
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   180
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfSignalSource 
      Height          =   1665
      Left            =   330
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   480
      Width           =   5025
      _cx             =   8864
      _cy             =   2937
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
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   0
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   260
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmClinicPlanStopVisitEdit.frx":000C
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
   Begin VB.CheckBox chkStopTime 
      Caption         =   "������ֹ"
      Height          =   210
      Left            =   3270
      TabIndex        =   10
      Top             =   3915
      Value           =   1  'Checked
      Width           =   1065
   End
   Begin VB.Frame fraButton 
      BorderStyle     =   0  'None
      Height          =   795
      Left            =   60
      TabIndex        =   20
      Top             =   4320
      Width           =   5475
      Begin VB.CommandButton cmdHelp 
         Caption         =   "����(&H)"
         Height          =   360
         Left            =   300
         TabIndex        =   14
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "ȡ��(&C)"
         Height          =   360
         Left            =   4140
         TabIndex        =   13
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "ȷ��(&O)"
         Height          =   360
         Left            =   3090
         TabIndex        =   12
         Top             =   240
         Width           =   975
      End
      Begin VB.Frame fraSplit 
         Height          =   25
         Left            =   0
         TabIndex        =   21
         Top             =   0
         Width           =   5745
      End
   End
   Begin VB.TextBox txtAuditTime 
      Enabled         =   0   'False
      Height          =   300
      Left            =   3270
      TabIndex        =   9
      Top             =   3480
      Width           =   2085
   End
   Begin VB.TextBox txtAuditName 
      Enabled         =   0   'False
      Height          =   300
      Left            =   900
      TabIndex        =   8
      Top             =   3480
      Width           =   1305
   End
   Begin VB.TextBox txtApplyTime 
      Enabled         =   0   'False
      Height          =   300
      Left            =   900
      TabIndex        =   7
      Top             =   3075
      Width           =   2085
   End
   Begin VB.ComboBox cboReason 
      Height          =   300
      Left            =   900
      TabIndex        =   6
      Top             =   2685
      Width           =   4455
   End
   Begin MSComCtl2.DTPicker dtpStartTime 
      Height          =   300
      Left            =   900
      TabIndex        =   4
      Top             =   2280
      Width           =   2085
      _ExtentX        =   3678
      _ExtentY        =   529
      _Version        =   393216
      CalendarTitleBackColor=   -2147483630
      CalendarTitleForeColor=   -2147483634
      CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
      Format          =   169869315
      CurrentDate     =   42362
   End
   Begin VB.ComboBox cboApplyName 
      Height          =   300
      Left            =   900
      TabIndex        =   1
      Top             =   120
      Width           =   2085
   End
   Begin MSComCtl2.DTPicker dtpEndTime 
      Height          =   300
      Left            =   3270
      TabIndex        =   5
      Top             =   2280
      Width           =   2085
      _ExtentX        =   3678
      _ExtentY        =   529
      _Version        =   393216
      CalendarTitleBackColor=   -2147483630
      CalendarTitleForeColor=   -2147483634
      CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
      Format          =   169869315
      CurrentDate     =   42367.9999884259
   End
   Begin MSComCtl2.DTPicker dtpStopTime 
      Height          =   300
      Left            =   900
      TabIndex        =   11
      Top             =   3870
      Width           =   2085
      _ExtentX        =   3678
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   0   'False
      CalendarTitleBackColor=   -2147483630
      CalendarTitleForeColor=   -2147483634
      CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
      Format          =   169869315
      CurrentDate     =   42362
   End
   Begin VB.Label lblStopTime 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��ֹʱ��"
      Height          =   180
      Left            =   150
      TabIndex        =   22
      Top             =   3930
      Width           =   720
   End
   Begin VB.Label lblTime 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ͣ��ʱ��                        ��"
      Height          =   180
      Left            =   150
      TabIndex        =   19
      Top             =   2340
      Width           =   3060
   End
   Begin VB.Label lblReason 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ͣ��ԭ��"
      Height          =   180
      Left            =   150
      TabIndex        =   18
      Top             =   2745
      Width           =   720
   End
   Begin VB.Label lblAuditName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "������"
      Height          =   180
      Left            =   330
      TabIndex        =   17
      Top             =   3540
      Width           =   540
   End
   Begin VB.Label lblAuditTime 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����ʱ��"
      Height          =   180
      Left            =   2520
      TabIndex        =   16
      Top             =   3540
      Width           =   720
   End
   Begin VB.Label lblApplyTime 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����ʱ��"
      Height          =   180
      Left            =   150
      TabIndex        =   15
      Top             =   3135
      Width           =   720
   End
   Begin VB.Label lblApplyName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "������"
      Height          =   180
      Left            =   330
      TabIndex        =   0
      Top             =   180
      Width           =   540
   End
End
Attribute VB_Name = "frmClinicPlanStopVisitEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mbytFun As m_BytFun '���ܣ�1-���룬2-ȡ�����룬3-������4-ȡ��������4-��ֹ����
Private mlngID As Long '�ٴ�����ͣ���¼ID
Private mblnOk As Boolean
Private mlngModule As Long
Private mstrPrivs As String

Private Enum m_BytFun
    Fun_Applay = 1
    Fun_UnApplay = 2
    Fun_Audit = 3
    Fun_UnAudit = 4
    Fun_StopPlan = 5
End Enum
Private mrsDoctor As ADODB.Recordset
Private mbytԤԼ�嵥���Ʒ�ʽ As Byte
Private mbytԤԼ�嵥��ӡ��ʽ As Byte
Private mstrDoctorName As String

Public Function ShowMe(frmParent As Form, ByVal lngModule As Long, _
    ByVal strPrivs As String, ByVal bytFun As Byte, _
    Optional ByVal lngID As Long, Optional ByRef strDoctorName As String) As Boolean
    '�������
    '��Σ�
    '   frmParent ������
    '   bytFun 1-���룬2-ȡ�����룬3-������4-ȡ������
    '   lngID �ٴ�����ͣ���¼ID
    mstrPrivs = strPrivs: mlngModule = lngModule
    mbytFun = bytFun: mlngID = lngID
    mstrDoctorName = ""
    
    Err = 0: On Error Resume Next
    If CheckDepend() = False Then Exit Function
    mblnOk = False
    Me.Show 1, frmParent
    ShowMe = mblnOk
    If mblnOk Then strDoctorName = mstrDoctorName
End Function

Private Sub cboApplyName_Click()
    Err = 0: On Error GoTo errHandle
    If cboApplyName.ListIndex = -1 Then Exit Sub
    Call LoadSignalSource(cboApplyName.ItemData(cboApplyName.ListIndex), vsfSignalSource.Tag)
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub cboApplyName_GotFocus()
    zlControl.TxtSelAll cboApplyName
End Sub

Private Sub cboApplyName_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long, lngҽ��ID As Long
    
    Err = 0: On Error GoTo errHandle
    If KeyAscii <> 13 Then Exit Sub
    If cboApplyName.ListIndex <> -1 Then
        zlCommFun.PressKey vbKeyTab: Exit Sub
    End If
    If mrsDoctor Is Nothing Then Exit Sub
    If Trim(cboApplyName.Text) = "" Then zlCommFun.PressKey vbKeyTab: Exit Sub
    
    If zlPersonSelect(Me, mlngModule, cboApplyName, mrsDoctor, Trim(cboApplyName.Text), True, "") = False Then
        KeyAscii = 0: Exit Sub
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub cboApplyName_Validate(Cancel As Boolean)
    If cboApplyName.ListIndex < 0 Then cboApplyName.Text = ""
End Sub

Private Sub cboReason_GotFocus()
    zlControl.TxtSelAll cboReason
End Sub

Private Sub cboReason_KeyPress(KeyAscii As Integer)
    Dim strReason As String
    
    If KeyAscii <> vbKeyReturn Then Exit Sub
    If Trim(cboReason.Text) = "" Then zlCommFun.PressKey vbKeyTab: Exit Sub
    strReason = SearchStopVisitReason(Me, cboReason, Trim(cboReason.Text))
    If strReason = "" Then Exit Sub
    zlControl.CboLocate cboReason, strReason
    If cboReason.ListIndex = -1 Then cboReason.Text = strReason
End Sub

Private Sub chkStopTime_Click()
    dtpStopTime.Enabled = (chkStopTime.Value = vbUnchecked)
    If dtpStopTime.Visible And dtpStopTime.Enabled Then dtpStopTime.SetFocus
End Sub

Private Sub chkStopTime_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab: Exit Sub
End Sub

Private Sub chk����ͣ��_Click()
    Err = 0: On Error GoTo errHandler
    With vsfSignalSource
        If chk����ͣ��.Value = vbChecked Then
            .Editable = flexEDKbdMouse
            .Cell(flexcpForeColor, 0, 0, .Rows - 1, .Cols - 1) = vbBlack
        Else
            .Editable = flexEDNone
            .Cell(flexcpForeColor, 0, 0, .Rows - 1, .Cols - 1) = vbGrayText
        End If
    End With
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.Hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub cmdOK_Click()
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim str��¼IDs As String, strTemp As String
    
    Err = 0: On Error GoTo errHandler
    If IsValied() = False Then Exit Sub
    If SaveData() = False Then Exit Sub
    mblnOk = True
    mstrDoctorName = NeedName(cboApplyName.Text)
    
    '����Ƿ�Ҫ���ԤԼ�嵥
    If mbytFun = Fun_Audit Then '���
        strSQL = "Select a.ID as ��¼ID" & vbNewLine & _
                " From �ٴ������¼ A, �ٴ�����ͣ���¼ B, ���˹Һż�¼ C,�ٴ������Դ D" & vbNewLine & _
                " Where ((a.����ҽ������ Is Null And a.ҽ��id Is Not Null And a.ҽ������ = b.������)" & vbNewLine & _
                "       Or (a.����ҽ������ Is Not Null And a.����ҽ��id Is Not Null And a.����ҽ������ = b.������))" & vbNewLine & _
                "       And b.Id = [1] And Not (a.��ʼʱ�� > b.��ֹʱ�� Or a.��ֹʱ�� < b.��ʼʱ��)" & vbNewLine & _
                "       And a.��ԴID = d.ID And (b.ͣ����� Is Null Or Instr(','||b.ͣ�����||',', ','||d.����||',') > 0)" & vbNewLine & _
                "       And Exists (Select 1" & vbNewLine & _
                "                   From �ٴ����ﰲ�� C, �ٴ������ D" & vbNewLine & _
                "                   Where c.����id = d.Id And c.Id = a.����id And d.����ʱ�� Is Not Null)" & vbNewLine & _
                "        And a.Id = c.�����¼id And c.��¼״̬ = 1" & vbNewLine & _
                "       And (c.��¼���� = 1 And c.����ʱ�� Between a.ͣ�￪ʼʱ�� And a.ͣ����ֹʱ��" & vbNewLine & _
                "           Or c.��¼���� = 2 And c.ԤԼʱ�� Between a.ͣ�￪ʼʱ�� And a.ͣ����ֹʱ��)"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngID)
        
        If rsTemp Is Nothing Then GoTo unloadForm:
        If rsTemp.EOF Then GoTo unloadForm:
    
        Do While Not rsTemp.EOF
            If InStr(strTemp & ",", "," & Nvl(rsTemp!��¼ID) & ",") = 0 Then
                str��¼IDs = str��¼IDs & "," & Nvl(rsTemp!��¼ID)
            End If
            rsTemp.MoveNext
        Loop
        If str��¼IDs <> "" Then str��¼IDs = Mid(str��¼IDs, 2)
        
        If mbytԤԼ�嵥���Ʒ�ʽ = 1 Then
            Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_INSIDE_1114_4", Me, "�����¼IDS=" & str��¼IDs, 3)
        ElseIf mbytԤԼ�嵥���Ʒ�ʽ = 2 Then
            If MsgBox("��ǰҽ��ͣ��ʱ�䷶Χ�ڴ���ԤԼ��ҺŲ��ˣ��Ƿ�ԤԼ�嵥�����Excel�У�", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_INSIDE_1114_4", Me, "�����¼IDS=" & str��¼IDs, 3)
            End If
        End If
        
        If mbytԤԼ�嵥��ӡ��ʽ = 1 Then
            Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_INSIDE_1114_4", Me, "�����¼IDS=" & str��¼IDs, 2)
        ElseIf mbytԤԼ�嵥��ӡ��ʽ = 2 Then
            If MsgBox("��ǰҽ��ͣ��ʱ�䷶Χ�ڴ���ԤԼ��ҺŲ��ˣ���ȷ��Ҫ��ӡԤԼ�嵥��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_INSIDE_1114_4", Me, "�����¼IDS=" & str��¼IDs, 2)
            End If
        End If
    End If
    
unloadForm:
    mblnOk = True
    Unload Me
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function SaveData() As Boolean
    Dim strSQL As String
    Dim strNOs As String, i As Integer
    
    Err = 0: On Error GoTo errHandler
    '1-���룬2-ȡ�����룬3-������4-ȡ������
    If mbytFun = Fun_Applay Then
        If chk����ͣ��.Value = vbChecked Then
            With vsfSignalSource
                For i = 1 To .Rows - 1
                    If .Cell(flexcpChecked, i, .ColIndex("ѡ��")) = 1 Then
                        strNOs = strNOs & "," & .TextMatrix(i, .ColIndex("����"))
                    End If
                Next
                If strNOs <> "" Then strNOs = Mid(strNOs, 2)
            End With
        End If
        
        'Zl_�ٴ�����ͣ��_Apply(
        strSQL = "Zl_�ٴ�����ͣ��_Apply("
        '��������_In Number,--0-���룬else-ȡ������
        strSQL = strSQL & "" & 0 & ","
        'Id_In       �ٴ�����ͣ���¼.Id%Type,
        strSQL = strSQL & "" & mlngID & ","
        'ͣ�����_In �ٴ�����ͣ���¼.ͣ�����%type := Null,
        strSQL = strSQL & "'" & strNOs & "',"
        '��ʼʱ��_In �ٴ�����ͣ���¼.��ʼʱ��%Type := Null,
        strSQL = strSQL & "" & ZDate(dtpStartTime.Value) & ","
        '��ֹʱ��_In �ٴ�����ͣ���¼.��ֹʱ��%Type := Null,
        strSQL = strSQL & "" & ZDate(dtpEndTime.Value) & ","
        'ͣ��ԭ��_In �ٴ�����ͣ���¼.ͣ��ԭ��%Type := Null,
        strSQL = strSQL & "'" & NeedName(cboReason.Text) & "',"
        '������_In   �ٴ�����ͣ���¼.������%Type := Null,
        strSQL = strSQL & "'" & NeedName(cboApplyName.Text) & "',"
        '����ʱ��_In �ٴ�����ͣ���¼.����ʱ��%Type := Null,
        strSQL = strSQL & "" & ZDate(txtApplyTime.Text) & ","
        '�Ǽ���_In   �ٴ�����ͣ���¼.�Ǽ���%Type := Null
        strSQL = strSQL & "'" & UserInfo.���� & "')"
    ElseIf mbytFun = Fun_UnApplay Then
        'Zl_�ٴ�����ͣ��_Apply(
        strSQL = "Zl_�ٴ�����ͣ��_Apply("
        '��������_In Number,--0-���룬else-ȡ������
        strSQL = strSQL & "" & 1 & ","
        'Id_In       �ٴ�����ͣ���¼.Id%Type,
        strSQL = strSQL & "" & mlngID & ")"
        'ͣ�����_In �ٴ�����ͣ���¼.ͣ�����%type := Null,
        '��ʼʱ��_In �ٴ�����ͣ���¼.��ʼʱ��%Type := Null,
        '��ֹʱ��_In �ٴ�����ͣ���¼.��ֹʱ��%Type := Null,
        'ͣ��ԭ��_In �ٴ�����ͣ���¼.ͣ��ԭ��%Type := Null,
        '������_In   �ٴ�����ͣ���¼.������%Type := Null,
        '����ʱ��_In �ٴ�����ͣ���¼.����ʱ��%Type := Null,
        '�Ǽ���_In   �ٴ�����ͣ���¼.�Ǽ���%Type := Null
    ElseIf mbytFun = Fun_Audit Then
        'Zl_�ٴ�����ͣ��_Audit(
        strSQL = "Zl_�ٴ�����ͣ��_Audit("
        '��������_In Number,--1-��ˣ�2-ȡ�����
        strSQL = strSQL & "" & 1 & ","
        'Id_In       �ٴ�����ͣ���¼.Id%Type,
        strSQL = strSQL & "" & mlngID & ","
        '������_In   �ٴ�����ͣ���¼.������%Type := Null,
        strSQL = strSQL & "'" & UserInfo.���� & "',"
        '����ʱ��_In �ٴ�����ͣ���¼.����ʱ��%Type := Null
        strSQL = strSQL & "" & ZDate(zlDatabase.Currentdate) & ")"
    ElseIf mbytFun = Fun_UnAudit Then
        'Zl_�ٴ�����ͣ��_Audit(
        strSQL = "Zl_�ٴ�����ͣ��_Audit("
        '��������_In Number,--1-��ˣ�2-ȡ�����
        strSQL = strSQL & "" & 2 & ","
        'Id_In       �ٴ�����ͣ���¼.Id%Type,
        strSQL = strSQL & "" & mlngID & ")"
        '������_In   �ٴ�����ͣ���¼.������%Type := Null,
        '����ʱ��_In �ٴ�����ͣ���¼.����ʱ��%Type := Null
    ElseIf mbytFun = Fun_StopPlan Then
        'Zl_�ٴ�����ͣ��_Stop
        strSQL = "Zl_�ٴ�����ͣ��_Stop("
        '  Id_In       �ٴ�����ͣ���¼.Id%Type,
        strSQL = strSQL & "" & mlngID & ","
        '  ��ֹ��_In   �ٴ�����ͣ���¼.ȡ����%Type,
        strSQL = strSQL & "'" & UserInfo.���� & "',"
        '  ��ֹʱ��_In �ٴ�����ͣ���¼.ʧЧʱ��%Type := Null--Null-������ֹ������-�������ֹʱ��
        If chkStopTime.Value = vbChecked Then
            strSQL = strSQL & "" & "NULL" & ")"
        Else
            strSQL = strSQL & "" & ZDate(dtpStopTime.Value) & ")"
        End If
    End If
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    SaveData = True
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub dtpEndTime_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab: Exit Sub
End Sub

Private Sub dtpStartTime_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab: Exit Sub
End Sub

Private Sub dtpStartTime_LostFocus()
    dtpEndTime.Value = Format(dtpStartTime.Value, "yyyy-mm-dd 23:59:59")
End Sub

Private Sub dtpStopTime_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab: Exit Sub
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("':��;��?��", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
End Sub

Private Sub Form_Load()
    Dim strSQL As String, rsTemp As ADODB.Recordset
    
    Err = 0: On Error GoTo errHandler
    Me.Caption = Choose(mbytFun, "ͣ������", "ȡ������", "ͣ������", "ȡ������", "��ֹ����")
    
    If mbytFun = Fun_Audit Then
        mbytԤԼ�嵥���Ʒ�ʽ = Val(zlDatabase.GetPara("ԤԼ�嵥���Ʒ�ʽ", glngSys, mlngModule, "0"))
        mbytԤԼ�嵥��ӡ��ʽ = Val(zlDatabase.GetPara("ԤԼ�嵥��ӡ��ʽ", glngSys, mlngModule, "0"))
    End If
    
    If mbytFun = Fun_Applay Then
        If Not zlStr.IsHavePrivs(mstrPrivs, "���������ͣ������") Then
            cboApplyName.Enabled = False
        End If
    Else
        cboApplyName.Enabled = False
        dtpStartTime.Enabled = False
        dtpEndTime.Enabled = False
        cboReason.Enabled = False
    End If
    
    lblAuditTime.Visible = (mbytFun = Fun_UnAudit Or mbytFun = Fun_StopPlan)
    txtAuditTime.Visible = (mbytFun = Fun_UnAudit Or mbytFun = Fun_StopPlan)
    lblAuditName.Visible = (mbytFun = Fun_UnAudit Or mbytFun = Fun_StopPlan)
    txtAuditName.Visible = (mbytFun = Fun_UnAudit Or mbytFun = Fun_StopPlan)
    
    lblStopTime.Visible = (mbytFun = Fun_StopPlan)
    dtpStopTime.Visible = (mbytFun = Fun_StopPlan)
    chkStopTime.Visible = (mbytFun = Fun_StopPlan)
    
    If mbytFun = Fun_UnAudit Then
        fraButton.Top = txtAuditTime.Top + txtAuditTime.Height + 150
    ElseIf mbytFun = Fun_StopPlan Then
        fraButton.Top = dtpStopTime.Top + dtpStopTime.Height + 150
    Else
        fraButton.Top = txtApplyTime.Top + txtApplyTime.Height + 150
    End If
    Me.Height = fraButton.Top + fraButton.Height + 280
    
    Call SetEnabledBackColor(Me.Controls)
    If InitData() = False Then Unload Me: Exit Sub
    If LoadData(mbytFun) = False Then Unload Me: Exit Sub
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function InitData() As Boolean
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim strPersons As String
    
    Err = 0: On Error GoTo errHandler
    If zlStr.IsHavePrivs(mstrPrivs, "���������ͣ������") Then
        Set mrsDoctor = GetDoctor(, "���")
        cboApplyName.Clear
        Do While Not mrsDoctor.EOF
            If InStr("," & strPersons & ",", "," & Nvl(mrsDoctor!ID) & ",") = 0 Then
                strPersons = strPersons & "," & Nvl(mrsDoctor!ID)
                cboApplyName.AddItem Nvl(mrsDoctor!����) & "-" & Nvl(mrsDoctor!����)
                cboApplyName.ItemData(cboApplyName.NewIndex) = Nvl(mrsDoctor!ID)
            End If
            mrsDoctor.MoveNext
        Loop
    Else
        cboApplyName.Clear
        cboApplyName.AddItem UserInfo.���� & "-" & UserInfo.����
        cboApplyName.ItemData(cboApplyName.NewIndex) = UserInfo.ID
    End If
    
    strSQL = "Select ����, ����, ����, Nvl(ȱʡ��־, 0) As ȱʡ��־" & vbNewLine & _
            " From ����ͣ��ԭ��"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    cboReason.Clear
    Do While Not rsTemp.EOF
        cboReason.AddItem Nvl(rsTemp!����) & "-" & Nvl(rsTemp!����)
        If Val(Nvl(rsTemp!ȱʡ��־)) = 1 Then cboReason.ListIndex = cboReason.NewIndex
        rsTemp.MoveNext
    Loop
    InitData = True
    Exit Function
errHandler:
    vsfSignalSource.Redraw = flexRDBuffered
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function LoadSignalSource(ByVal lngҽ��ID As Long, _
    Optional ByVal strͣ����� As String) As String
    '��ȡҽ����������Ч��Դ
    '��Σ�
    '   strͣ����� ������ĺ��룬����ö��ŷָ�
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim lngRow As Long, strWhere As String
    
    Err = 0: On Error GoTo errHandler
    If strͣ����� = "" Then
        strWhere = _
            " And Nvl(a.����ʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) > Sysdate" & vbNewLine & _
            " And Nvl(b.����ʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) > Sysdate" & vbNewLine & _
            " And Nvl(c.����ʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) > Sysdate"
    Else
        strWhere = " And a.���� In(Select /*+cardinality(j,10)*/Column_Value From Table(f_Str2list([2])) J)"
    End If
    strSQL = _
        " Select a.����, b.���� As ����, c.���� As �շ���Ŀ" & vbNewLine & _
        " From �ٴ������Դ A, ���ű� B, �շ���ĿĿ¼ C" & vbNewLine & _
        " Where a.����id = b.Id And a.��Ŀid = c.Id And a.ҽ��id = [1]" & vbNewLine & _
                strWhere
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngҽ��ID, strͣ�����)
    
    With vsfSignalSource
        .Redraw = flexRDNone
        .Clear 1
        .Rows = rsTemp.RecordCount + 1
        lngRow = 1
        Do While Not rsTemp.EOF
            .Cell(flexcpChecked, lngRow, .ColIndex("ѡ��")) = 1 '���Ϊѡ��
            .TextMatrix(lngRow, .ColIndex("����")) = Nvl(rsTemp!����)
            .TextMatrix(lngRow, .ColIndex("����")) = Nvl(rsTemp!����)
            .TextMatrix(lngRow, .ColIndex("�շ���Ŀ")) = Nvl(rsTemp!�շ���Ŀ)
            
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
        .Cell(flexcpForeColor, 0, 0, .Rows - 1, .Cols - 1) = IIf(chk����ͣ��.Value = vbChecked, vbBlack, vbGrayText)
        .Redraw = flexRDBuffered
    End With
    chk����ͣ��.Visible = mbytFun = Fun_Applay And rsTemp.RecordCount > 0
    LoadSignalSource = True
    Exit Function
errHandler:
    vsfSignalSource.Redraw = flexRDBuffered
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function LoadData(ByVal bytFun As Byte) As Boolean
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim dtNow As Date
    
    Err = 0: On Error GoTo errHandler
    dtNow = zlDatabase.Currentdate
    If bytFun = Fun_Applay Then
        zlControl.CboSetText cboApplyName, UserInfo.����
        dtpStartTime.MinDate = Format(dtNow, "yyyy-mm-dd hh:mm:ss")
        dtpStartTime.Value = Format(dtNow + 1, "yyyy-mm-dd 00:00:00")
        dtpEndTime.MinDate = Format(dtNow, "yyyy-mm-dd hh:mm:ss")
        dtpEndTime.Value = Format(dtNow + 1, "yyyy-mm-dd 23:59:59")
        txtApplyTime.Text = Format(dtNow, "yyyy-mm-dd hh:mm:ss")
        LoadData = True: Exit Function
    End If
    
    strSQL = "Select ͣ��ԭ��, ��ʼʱ��, ��ֹʱ��, ������, ����ʱ��, ������, ����ʱ��, �Ǽ���, ͣ�����" & vbNewLine & _
            " From �ٴ�����ͣ���¼" & vbNewLine & _
            " Where ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngID)
    If rsTemp.EOF Then
        MsgBox "��¼�����ڣ������ѱ�����" & IIf(bytFun = Fun_UnAudit, "ȡ������", "ȡ�����������") & "��", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    vsfSignalSource.Tag = Nvl(rsTemp!ͣ�����)
    
    zlControl.CboSetText cboApplyName, Nvl(rsTemp!������)
    If cboApplyName.ListIndex = -1 Then cboApplyName.Text = Nvl(rsTemp!������)
    cboApplyName.Tag = Nvl(rsTemp!�Ǽ���) '�洢�Ǽ��ˣ����ڼ��
    
    dtpStartTime.Value = Format(Nvl(rsTemp!��ʼʱ��), "yyyy-mm-dd hh:mm:ss")
    dtpEndTime.Value = Format(Nvl(rsTemp!��ֹʱ��), "yyyy-mm-dd hh:mm:ss")
    
    zlControl.CboSetText cboReason, Nvl(rsTemp!ͣ��ԭ��)
    If cboReason.ListIndex = -1 Then cboReason.Text = Nvl(rsTemp!ͣ��ԭ��)
    
    txtApplyTime.Text = Format(Nvl(rsTemp!����ʱ��), "yyyy-mm-dd hh:mm:ss")
    If bytFun = Fun_UnApplay Then LoadData = True: Exit Function
    
    txtAuditName.Text = Nvl(rsTemp!������)
    txtAuditTime.Text = Format(Nvl(rsTemp!����ʱ��), "yyyy-mm-dd hh:mm:ss")
    
    dtpStopTime.Value = Format(dtNow, "yyyy-mm-dd hh:mm:ss")
    dtpStopTime.MaxDate = Format(Nvl(rsTemp!��ֹʱ��), "yyyy-mm-dd hh:mm:ss")
    LoadData = True
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub Form_Unload(Cancel As Integer)
    Set mrsDoctor = Nothing
End Sub

Private Function IsValied() As Boolean
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim lngCount As Long, i As Integer, strStopNOs As String
    Dim str��ͣ���� As String, varData As Variant
    
    Err = 0: On Error GoTo errHandle
    If mbytFun = Fun_Applay Then
        If zlControl.FormCheckInput(Me) = False Then Exit Function
        If cboApplyName.ListIndex < 0 Or cboApplyName.Text = "" Then
            MsgBox "��ѡ�������ˣ�", vbInformation, gstrSysName
            If cboApplyName.Visible And cboApplyName.Enabled Then cboApplyName.SetFocus
            Exit Function
        End If
        
        If chk����ͣ��.Value = vbChecked Then
            With vsfSignalSource
                For i = 1 To .Rows - 1
                    If .Cell(flexcpChecked, i, .ColIndex("ѡ��")) = 1 Then
                        lngCount = lngCount + 1
                        strStopNOs = strStopNOs & "," & .TextMatrix(i, .ColIndex("����"))
                    End If
                Next
                If strStopNOs <> "" Then strStopNOs = Mid(strStopNOs, 2)
                If .Rows > 1 And lngCount = 0 Then
                    MsgBox "��ѡ��ͣ����룡", vbInformation, gstrSysName
                    If vsfSignalSource.Visible And vsfSignalSource.Enabled Then vsfSignalSource.SetFocus
                    Exit Function
                End If
                If lngCount > 100 Then
                    MsgBox "ÿһ�������ͣ����벻�ܳ���100������ֶ�����룡", vbInformation, gstrSysName
                    If vsfSignalSource.Visible And vsfSignalSource.Enabled Then vsfSignalSource.SetFocus
                    Exit Function
                End If
            End With
        End If
        
        If DateDiff("s", dtpStartTime.Value, zlDatabase.Currentdate) >= 0 Then
            MsgBox "ͣ��ʱ��Ŀ�ʼʱ�������ڵ�ǰʱ�䣡", vbInformation, gstrSysName
            If dtpEndTime.Visible And dtpEndTime.Enabled Then dtpEndTime.SetFocus
            Exit Function
        End If
        
        If DateDiff("s", dtpStartTime.Value, dtpEndTime.Value) <= 0 Then
            MsgBox "ͣ��ʱ��Ľ���ʱ�������ڿ�ʼʱ�䣡", vbInformation, gstrSysName
            If dtpEndTime.Visible And dtpEndTime.Enabled Then dtpEndTime.SetFocus
            Exit Function
        End If
        
        If zlControl.TxtCheckInput(cboReason, "ͣ��ԭ��", 50, False) = False Then Exit Function
        
        strSQL = "Select 1 From �ٴ�����ͣ���¼" & vbNewLine & _
                " Where ��¼id Is Null And Not (��ʼʱ�� > [2] Or Nvl(ʧЧʱ��, ��ֹʱ��) < [1])" & vbNewLine & _
                "       And ������ = [3] And ͣ����� Is Null And Rownum < 2"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, dtpStartTime.Value, dtpEndTime.Value, NeedName(cboApplyName.Text))
        If Not rsTemp.EOF Then
            MsgBox "��ǰͣ��ʱ����������ͣ��ʱ�䷶Χ�����ص������飡", vbInformation, gstrSysName
            Exit Function
        End If
        
        strSQL = "Select ͣ����� From �ٴ�����ͣ���¼" & vbNewLine & _
                " Where ��¼id Is Null And Not (��ʼʱ�� > [2] Or Nvl(ʧЧʱ��, ��ֹʱ��) < [1])" & vbNewLine & _
                "       And ������ = [3] And ͣ����� Is Not Null"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, dtpStartTime.Value, dtpEndTime.Value, NeedName(cboApplyName.Text))
        Do While Not rsTemp.EOF
            If strStopNOs = "" Then
                str��ͣ���� = str��ͣ���� & "," & Nvl(rsTemp!ͣ�����)
            Else
                varData = Split(strStopNOs, ",")
                For i = 0 To UBound(varData)
                    If InStr("," & Nvl(rsTemp!ͣ�����) & ",", "," & varData(i) & ",") > 0 Then
                        str��ͣ���� = str��ͣ���� & "," & varData(i)
                    End If
                Next
            End If
            rsTemp.MoveNext
        Loop
        If str��ͣ���� <> "" Then
            str��ͣ���� = Mid(str��ͣ����, 2)
            MsgBox "����(" & str��ͣ���� & ")��ǰͣ��ʱ����������ͣ��ʱ�䷶Χ�����ص������飡", vbInformation, gstrSysName
            Exit Function
        End If
    ElseIf mbytFun = Fun_UnApplay Then
        If Not zlStr.IsHavePrivs(mstrPrivs, "���������ͣ������") Then
            If Not (NeedName(cboApplyName.Text) = UserInfo.���� Or cboApplyName.Tag = UserInfo.����) Then
                MsgBox "��ֻ��ɾ���Լ������룬����ɾ�����˵����룡", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    ElseIf mbytFun = Fun_StopPlan Then
        If chkStopTime.Value = vbUnchecked Then
            If DateDiff("s", dtpStopTime.Value, zlDatabase.Currentdate) >= 0 Then
                MsgBox "��ֹʱ�������ڵ�ǰʱ�䣡", vbInformation, gstrSysName
                If dtpStopTime.Visible And dtpStopTime.Enabled Then dtpStopTime.SetFocus
                Exit Function
            End If
        End If
    End If
    If CheckDepend() = False Then Exit Function
    IsValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Function CheckDepend() As Boolean
    '����:�������
    Dim strSQL As String, rsTemp As ADODB.Recordset
    
    On Error GoTo errHandler
    If mbytFun = Fun_UnApplay Then
        strSQL = "Select 1 From �ٴ�����ͣ���¼ Where ID = [1] And ������ Is Not Null"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngID)
        If Not rsTemp.EOF Then
            MsgBox "�������ѱ�����������ȡ�����롣", vbInformation, gstrSysName
            Exit Function
        End If
    ElseIf mbytFun = Fun_Audit Then
        strSQL = "Select 1 From �ٴ�����ͣ���¼ Where ID = [1] And ������ Is Not Null"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngID)
        If Not rsTemp.EOF Then
            MsgBox "�������ѱ������������ٴ�������", vbInformation, gstrSysName
            Exit Function
        End If
    ElseIf mbytFun = Fun_UnAudit Then
        strSQL = "Select 1 From �ٴ�����ͣ���¼ Where ID = [1] And ��ֹʱ�� < Sysdate"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngID)
        If Not rsTemp.EOF Then
            MsgBox "��ͣ�ﰲ����ʧЧ������ȡ��������", vbInformation, gstrSysName
            Exit Function
        End If
        
        strSQL = "Select 1 From �ٴ�����ͣ���¼ Where ID = [1] And ʧЧʱ�� Is Not Null"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngID)
        If Not rsTemp.EOF Then
            MsgBox "��ͣ�ﰲ���ѱ���ֹ������ȡ��������", vbInformation, gstrSysName
            Exit Function
        End If
        
        strSQL = "Select 1 From �ٴ������¼ A, �ٴ�����ͣ���¼ B, ���˷�����Ϣ��¼ C" & vbNewLine & _
                " Where Nvl(a.����ҽ������, a.ҽ������) = b.������ And Nvl(a.����ҽ��id, a.ҽ��id) Is Not Null" & vbNewLine & _
                "       And a.Id = c.��¼id And (a.��ʼʱ�� Between b.��ʼʱ�� And b.��ֹʱ�� Or a.��ֹʱ�� Between b.��ʼʱ�� And b.��ֹʱ��)" & vbNewLine & _
                "       And c.������ Is Not Null And b.Id = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngID)
        If Not rsTemp.EOF Then
            MsgBox "��ͣ�ﰲ�ŵĲ���ͣ����Ϣ�ѱ���������ȡ��������", vbInformation, gstrSysName
            Exit Function
        End If
    ElseIf mbytFun = Fun_StopPlan Then
        strSQL = "Select 1 From �ٴ�����ͣ���¼ Where ID = [1] And ��ֹʱ�� < Sysdate"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngID)
        If Not rsTemp.EOF Then
            MsgBox "��ͣ�ﰲ����ʧЧ��������ֹ��", vbInformation, gstrSysName
            Exit Function
        End If
        
        strSQL = "Select 1 From �ٴ�����ͣ���¼ Where ID = [1] And ʧЧʱ�� Is Not Null"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngID)
        If Not rsTemp.EOF Then
            MsgBox "��ͣ�ﰲ���ѱ���ֹ����������ֹ��", vbInformation, gstrSysName
            Exit Function
        End If
        
        strSQL = "Select 1 From �ٴ�����ͣ���¼ Where ID = [1] And ������ Is Not Null"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngID)
        If rsTemp.EOF Then
            MsgBox "��ͣ�ﰲ�Ż�δ������������ֹ��", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    CheckDepend = True
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub vsfSignalSource_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = vsfSignalSource.ColIndex("ѡ��") Then Exit Sub
    Cancel = True
End Sub
