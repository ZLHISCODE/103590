VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmInsDiagnosis 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��ϱ༭"
   ClientHeight    =   4350
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6180
   Icon            =   "frmInsDiagnosis.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   6180
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VSFlex8Ctl.VSFlexGrid vfgSelect 
      Height          =   2175
      Left            =   -4080
      TabIndex        =   19
      Top             =   1395
      Visible         =   0   'False
      Width           =   4680
      _cx             =   8255
      _cy             =   3836
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
      Rows            =   2
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
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
   Begin VB.CommandButton cmdRef 
      Caption         =   "������ϲο�(&R)��"
      Height          =   350
      Left            =   135
      TabIndex        =   18
      Top             =   3825
      Width           =   1800
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4815
      TabIndex        =   11
      Top             =   3825
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   3630
      TabIndex        =   10
      Top             =   3825
      Width           =   1100
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Index           =   1
      Left            =   -15
      TabIndex        =   17
      Top             =   3630
      Width           =   6345
   End
   Begin VB.Frame fraHint 
      Height          =   1215
      Left            =   1350
      TabIndex        =   12
      Top             =   2235
      Width           =   4575
      Begin VB.OptionButton optHint 
         Caption         =   "���������Ŀ¼��������(F4)"
         Height          =   180
         Index           =   2
         Left            =   165
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   840
         Width           =   3360
      End
      Begin VB.OptionButton optHint 
         Caption         =   "����׼���������������(F3)"
         Height          =   180
         Index           =   1
         Left            =   165
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   562
         Width           =   3360
      End
      Begin VB.OptionButton optHint 
         Caption         =   "��������(F2)"
         Height          =   180
         Index           =   0
         Left            =   165
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   285
         Value           =   -1  'True
         Width           =   2190
      End
      Begin VB.Label lblHint 
         AutoSize        =   -1  'True
         Caption         =   "���뷽����ʾ:"
         Height          =   180
         Left            =   135
         TabIndex        =   13
         Top             =   15
         Width           =   1170
      End
   End
   Begin VB.CheckBox chkDoubt 
      Caption         =   "����(&U)"
      Height          =   195
      Left            =   1350
      TabIndex        =   7
      Top             =   1485
      Width           =   945
   End
   Begin VB.TextBox txtSymptom 
      Height          =   300
      Left            =   1350
      TabIndex        =   9
      Top             =   1800
      Width           =   4575
   End
   Begin VB.TextBox txtDisease 
      Height          =   300
      Left            =   1350
      TabIndex        =   6
      Top             =   1080
      Width           =   4575
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Index           =   0
      Left            =   -15
      TabIndex        =   4
      Top             =   900
      Width           =   6345
   End
   Begin VB.OptionButton optType 
      Caption         =   "��ҽ���(&H)"
      Height          =   180
      Index           =   1
      Left            =   4710
      TabIndex        =   3
      Top             =   510
      Width           =   1335
   End
   Begin VB.OptionButton optType 
      Caption         =   "��ҽ���(&W)"
      Height          =   180
      Index           =   0
      Left            =   3345
      TabIndex        =   2
      Top             =   510
      Value           =   -1  'True
      Width           =   1335
   End
   Begin VB.ComboBox cboKind 
      Height          =   300
      Left            =   1020
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   450
      Width           =   1725
   End
   Begin VB.ComboBox cboIn 
      Height          =   300
      Left            =   3210
      Style           =   2  'Dropdown List
      TabIndex        =   20
      Top             =   1432
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.ComboBox cboOut 
      Height          =   300
      Left            =   5055
      Style           =   2  'Dropdown List
      TabIndex        =   21
      Top             =   1432
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.Label lblSymptom 
      AutoSize        =   -1  'True
      Caption         =   "֤��(&S)"
      Height          =   180
      Left            =   690
      TabIndex        =   8
      Top             =   1860
      Width           =   630
   End
   Begin VB.Label lblDisease 
      AutoSize        =   -1  'True
      Caption         =   "����(&D)"
      Height          =   180
      Left            =   690
      TabIndex        =   5
      Top             =   1140
      Width           =   630
   End
   Begin VB.Label lblKind 
      AutoSize        =   -1  'True
      Caption         =   "��סԺ�����޶������У���ѡ�������������͵���ϣ�"
      Height          =   180
      Left            =   690
      TabIndex        =   0
      Top             =   135
      Width           =   4500
   End
   Begin VB.Image imgNote 
      Height          =   480
      Left            =   135
      Picture         =   "frmInsDiagnosis.frx":038A
      Top             =   195
      Width           =   480
   End
   Begin VB.Label LabOut 
      Caption         =   "��Ժ���"
      Height          =   210
      Left            =   4260
      TabIndex        =   23
      Top             =   1485
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Label LabIn 
      Caption         =   "��Ժ����"
      Height          =   210
      Left            =   2400
      TabIndex        =   22
      Top             =   1500
      Visible         =   0   'False
      Width           =   750
   End
End
Attribute VB_Name = "frmInsDiagnosis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private mblnOk As Boolean
Private mobjDoc As cEPRDocument
Private mblnSyncPage As Boolean

Public Function ShowMe(ByRef edtThis As Editor, ByRef frmParent As frmMain) As cEPRDiagnosis
    '���ܣ���ʾ��ϱ༭���壬�����ر༭�����
    '������ frmParent-������
    
Dim intFileKind As Integer  '�����ļ�����
Dim strFileName As String   '�����ļ�����
Dim lngFileID As Long       '�����ļ������Id
Dim lngDeptId As Long       '��д�����ĵ�ǰ����
Dim blnEmend As Boolean     '�Ƿ��޶�״̬
Dim strCurTime As String
Dim rsTemp As New ADODB.Recordset
    
    '------------------------------------
    Set mobjDoc = frmParent.Document
    Select Case mobjDoc.EditType
    Case cprET_�����ļ�����
        intFileKind = mobjDoc.EPRFileInfo.����
        strFileName = mobjDoc.EPRFileInfo.����
        lngFileID = mobjDoc.EPRFileInfo.ID
        lngDeptId = 0
        blnEmend = False
    Case cprET_ȫ��ʾ���༭
        intFileKind = 0
        strFileName = mobjDoc.EPRDemoInfo.����
        lngFileID = mobjDoc.EPRDemoInfo.�ļ�ID
        lngDeptId = 0
        blnEmend = False
    Case cprET_�������༭, cprET_���������
        intFileKind = mobjDoc.EPRPatiRecInfo.��������
        strFileName = mobjDoc.EPRPatiRecInfo.��������
        lngFileID = mobjDoc.EPRPatiRecInfo.�ļ�ID
        lngDeptId = mobjDoc.EPRPatiRecInfo.����ID
        blnEmend = (mobjDoc.EditType = cprET_���������)
    End Select
    
    Err = 0: On Error GoTo errHand
    gstrSQL = "Select l.����, q.�¼�, q.Ψһ, q.��дʱ��, h.��ҽ, Sysdate As ���� " & _
            " From �����ļ��б� l, ����ʱ��Ҫ�� q," & _
            "      (Select Sign(Nvl(Count(����ID), 0)) As ��ҽ From ��������˵�� Where ����id = [2] And �������� = '��ҽ��') h" & _
            " Where l.Id = q.�ļ�id(+) And l.Id = [1]"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngFileID, lngDeptId)
    If rsTemp.RecordCount <= 0 Then MsgBox "�ò����������ö�ʧ�����ܲ�����ϣ�", vbExclamation, gstrSysName: Exit Function
    intFileKind = rsTemp!����: strCurTime = Format(rsTemp!����, "yyyy-mm-dd hh:mm:ss")
    
    Me.cboKind.Clear
    If intFileKind = 1 Then
        Me.lblKind.Caption = "��" & strFileName & IIf(blnEmend, "�޶�", "�༭") & "�����У���ѡ�������������͵���ϣ�"
        Me.cboKind.AddItem "11-�������"
        Me.cboKind.ListIndex = 0
        Me.cboKind.Enabled = False
        Me.optType(1).Enabled = (rsTemp!��ҽ = 1)
    ElseIf intFileKind = 2 Then
        If (rsTemp!�¼� = "��Ժ" Or rsTemp!�¼� = "�״���Ժ" Or rsTemp!�¼� = "�ٴ���Ժ") And rsTemp!Ψһ = 1 Then
            Me.lblKind.Caption = "��" & strFileName & IIf(blnEmend, "�޶�", "�༭") & "�����У���ѡ�������������͵���ϣ�"
            If blnEmend = False Then
                Me.cboKind.AddItem "21-�������"
                Me.cboKind.ListIndex = 0
                Me.cboKind.Enabled = False
            Else
                Me.cboKind.AddItem "22-ȷ�����"
                Me.cboKind.AddItem "23-�������"
                Me.cboKind.AddItem "24-�������"
                Me.cboKind.ListIndex = 0
            End If
            Me.optType(1).Enabled = (rsTemp!��ҽ = 1)
        ElseIf rsTemp!�¼� = "24Сʱ��Ժ" Or rsTemp!�¼� = "24Сʱ����" Then
            Me.lblKind.Caption = "��" & strFileName & IIf(blnEmend, "�޶�", "�༭") & "�����У���ѡ�������������͵���ϣ�"
            If blnEmend = False Then
                Me.cboKind.AddItem "21-�������"
            Else
                Me.cboKind.AddItem "22-ȷ�����"
                Me.cboKind.AddItem "23-�������"
                Me.cboKind.AddItem "24-�������"
            End If
            Me.cboKind.AddItem "31-��Ժ���"
            Me.cboKind.ListIndex = 0
            Me.optType(1).Enabled = (rsTemp!��ҽ = 1)
        ElseIf rsTemp!�¼� = "��Ժ" Or rsTemp!�¼� = "����" Then
            Me.lblKind.Caption = "��" & strFileName & IIf(blnEmend, "�޶�", "�༭") & "�����У���ѡ�������������͵���ϣ�"
            Me.cboKind.AddItem "31-��Ժ���"
            Me.cboKind.ListIndex = 0
            Me.cboKind.Enabled = False
            Me.optType(1).Enabled = (rsTemp!��ҽ = 1)
            mblnSyncPage = (zldatabase.GetPara("SyncPage", glngSys, 1070, 0) = 1)
            If mblnSyncPage Then
                Call optType_Click(0)
            End If
        ElseIf rsTemp!�¼� = "����" And rsTemp!Ψһ = 1 Then
            Me.lblKind.Caption = "��" & strFileName & IIf(blnEmend, "�޶�", "�༭") & "�����У���ѡ�������������͵���ϣ�"
            Me.cboKind.AddItem "41-��ǰ���"
            Me.cboKind.AddItem "42-�������"
            Me.cboKind.ListIndex = 0
            Me.optType(1).Value = False: Me.optType(1).Enabled = False
        Else
            MsgBox "�ò������ܲ�����ϣ�", vbExclamation, gstrSysName: Exit Function
        End If
    ElseIf intFileKind = 7 Then     '���Ʊ���
        gstrSQL = "Select Nvl(Instr(i.��������, '����'), 0) As ����" & vbNewLine & _
                "From ����ҽ����¼ l, ������ĿĿ¼ i" & vbNewLine & _
                "Where l.������Ŀid = i.Id And l.������� = 'D' And l.Id = [1]"
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, mobjDoc.EPRPatiRecInfo.ҽ��id)
        If rsTemp.RecordCount <= 0 Then Exit Function   '�Ǽ�鱨�棬���ܲ������
        Me.lblKind.Caption = "��" & strFileName & IIf(blnEmend, "�޶�", "�༭") & "�����У���ѡ�������������͵���ϣ�"
        If rsTemp.Fields(0).Value > 0 Then
            Me.cboKind.AddItem "51-�������"
        Else
            Me.cboKind.AddItem "52-Ӱ�����"
        End If
        Me.cboKind.ListIndex = 0
        Me.optType(1).Value = False: Me.optType(1).Enabled = False
        Me.optType(0).Visible = False: Me.optType(1).Visible = False
    Else
        MsgBox "�ò������ܲ�����ϣ�", vbExclamation, gstrSysName: Exit Function
    End If
    
    '��¼��ǰ����ҽ��ҽ��־���Ա��жϸı�ʱ�����ϣ�
    If Me.optType(0).Value Then
        Me.lblKind.Tag = 0
    Else
        Me.lblKind.Tag = 1
    End If
    
    '------------------------------------
    If Me.optType(0).Value Then
        Me.lblSymptom.Enabled = False: Me.txtSymptom.Enabled = False
    Else
        Me.lblSymptom.Enabled = True: Me.txtSymptom.Enabled = True
    End If
    
    '���뷽ʽͨϵͳ��������
    '�Ƿ���������¼��
    If Mid(zldatabase.GetPara("������뷽ʽ", glngSys, , "11"), IIf(mobjDoc.EPRFileInfo.���� = cpr���ﲡ��, 1, 2), 1) = 1 Then
        optHint(0).Enabled = True
    Else
        optHint(0).Value = False
        optHint(0).Enabled = False
    End If
    
    Select Case zldatabase.GetPara("���������Դ", glngSys, , "1")
        Case 1 'ҽ������ѡ��
            optHint(1).Enabled = True
            optHint(2).Enabled = True
            If optHint(0).Enabled = False Then optHint(1).Value = True
        Case 2 '�����
            optHint(1).Enabled = False
            optHint(2).Enabled = True
            If optHint(0).Enabled = False Then optHint(2).Value = True
        Case 3 '��ICD10
            optHint(1).Enabled = True
            optHint(2).Enabled = False
            If optHint(0).Enabled = False Then optHint(1).Value = True
    End Select
    
    
    '��ʾ����
    Me.Show vbModal, frmParent
    If mblnOk = False Then Set ShowMe = Nothing: Unload Me: Exit Function
    
    '------------------------------------
    '���췵�ض���
    Dim rs As New ADODB.Recordset
    Dim oDiagnosis As cEPRDiagnosis
    Dim strTmp As String
    Dim aryDisease() As String
    
    Set oDiagnosis = New cEPRDiagnosis
    aryDisease = Split(Me.lblDisease.Tag, ",")
    
    
    '����Ӧ�ļ��������Ƿ���д
    
    If mobjDoc.EPRFileInfo.���� = cpr���ﲡ�� Or mobjDoc.EPRFileInfo.���� = cprסԺ���� Then
        Select Case mobjDoc.EditType
        Case cprET_�������༭, cprET_���������
            If UBound(aryDisease) >= 1 And mobjDoc.EPRPatiRecInfo.����ID > 0 Then
                If Val(aryDisease(1)) > 0 Then
    
                    gstrSQL = "Select Distinct b.����,c.����id From ��������ǰ�� a,�����ļ��б� b,���Ӳ�����¼ c Where a.���id=[1] And a.�ļ�id=b.Id And a.�ļ�id=c.�ļ�id(+) And c.����id(+)=1 And c.��ҳid(+)=1"
                    Set rs = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(aryDisease(1)), mobjDoc.EPRPatiRecInfo.����ID, mobjDoc.EPRPatiRecInfo.��ҳID)
                    If rs.BOF = False Then
                        strTmp = ""
                        Do While Not rs.EOF
                            If zlCommFun.NVL(rs("����id").Value, 0) = 0 Then
                                strTmp = strTmp & vbCrLf & Space(4) & rs("����").Value
                            End If
                            rs.MoveNext
                        Loop
                        If strTmp <> "" Then
                            MsgBox "���棺��ǰ���˵����¼���֤�����滹û����д��" & strTmp, vbInformation, gstrSysName
                        End If
                    End If
    
                ElseIf Val(aryDisease(0)) > 0 Then
                
                    gstrSQL = "Select Distinct b.����,c.����id From ��������ǰ�� a,�����ļ��б� b,���Ӳ�����¼ c Where a.����id=[1] And a.�ļ�id=b.Id And a.�ļ�id=c.�ļ�id(+) And c.����id(+)=1 And c.��ҳid(+)=1"
                    Set rs = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(aryDisease(0)), mobjDoc.EPRPatiRecInfo.����ID, mobjDoc.EPRPatiRecInfo.��ҳID)
                    If rs.BOF = False Then
                        strTmp = ""
                        Do While Not rs.EOF
                            If zlCommFun.NVL(rs("����id").Value, 0) = 0 Then
                                strTmp = strTmp & vbCrLf & Space(4) & rs("����").Value
                            End If
                            rs.MoveNext
                        Loop
                        If strTmp <> "" Then
                            MsgBox "���棺��ǰ���˵����¼���֤�����滹û����д��" & strTmp, vbInformation, gstrSysName
                        End If
                    End If
                End If
            End If
            
        End Select
    End If
    
    Err = 0: On Error GoTo 0
    With oDiagnosis
        .�ļ�ID = lngFileID
        .���� = Val(Me.cboKind.Text)
        If UBound(aryDisease) < 1 Then
            .����id = 0: .���id = 0
        Else
            .����id = Val(aryDisease(0)): .���id = Val(aryDisease(1))
        End If
        .֤��id = Val(Me.lblSymptom.Tag)
        If Me.optType(0).Value Then
            .���� = Trim(Me.txtDisease.Text)
        Else
            .���� = Trim(Me.txtDisease.Text) & "(" & Trim(Me.txtSymptom.Text) & ")"
        End If
        If Me.chkDoubt.Value = vbChecked Then
            .���� = .���� & "(?)"
            .���� = 1
        Else
            .���� = 0
        End If
        If optType(1).Value Then .��ҽ = 1
        .���� = strCurTime
        If mblnSyncPage Then
            .��Ժ���� = cboIn.Text
            .��Ժ��� = Mid(cboOut.Text, 3)
        End If
    End With
    Set ShowMe = oDiagnosis
    Unload Me: Exit Function

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Set ShowMe = Nothing
    Unload Me
End Function

Private Sub cboKind_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub chkDoubt_GotFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub chkDoubt_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cmdCancel_Click()
    mblnOk = False: Me.Hide: Exit Sub
End Sub

Private Sub cmdOK_Click()
    If (optHint(0).Value = False And txtDisease.Tag = "") Then MsgBox "�����뼲����ϲ��س���ȡ���룡", vbExclamation, gstrSysName: Me.txtDisease.SetFocus: Exit Sub
    If Trim(Me.txtDisease.Text) = "" Then MsgBox "û�����뼲����ϣ�", vbExclamation, gstrSysName: Me.txtDisease.SetFocus: Exit Sub
    If Me.optType(1).Value Then
        If (optHint(0).Value = False And txtSymptom.Tag = "") Then MsgBox "������֤�򲢻س���ȡ���룡", vbExclamation, gstrSysName: Me.txtDisease.SetFocus: Exit Sub
        If Trim(Me.txtSymptom.Text) = "" Then MsgBox "û������֤��", vbExclamation, gstrSysName: Me.txtSymptom.SetFocus:: Exit Sub
    End If
    mblnOk = True: Me.Hide: Exit Sub
End Sub

Private Sub cmdRef_Click()
    Dim aryDisease() As String, lngId As Long
    aryDisease = Split(Me.lblDisease.Tag, ",")
    If UBound(aryDisease) < 1 Then
        lngId = 0
    Else
        lngId = Val(aryDisease(1))
    End If
    Call mobjDoc.Event_ClickDiagRef(lngId, vbModal)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF2
        If Me.vfgSelect.Visible = False And optHint(0).Enabled Then Me.optHint(0).Value = True
    Case vbKeyF3
        If Me.vfgSelect.Visible = False And optHint(1).Enabled Then Me.optHint(1).Value = True
    Case vbKeyF4
        If Me.vfgSelect.Visible = False And optHint(2).Enabled Then Me.optHint(2).Value = True
    Case vbKeyEscape
        If Me.vfgSelect.Visible Then
            Me.vfgSelect.Visible = False
        Else
            Call cmdCancel_Click
        End If
    Case Else
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub optHint_Click(Index As Integer)
    If Me.txtDisease.Visible Then Me.txtDisease.SetFocus
End Sub

Private Sub optHint_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub optType_Click(Index As Integer)
Dim rsTemp As ADODB.Recordset, lCount As Long, i As Integer
    On Error GoTo errHand
    If Me.optType(0).Value Then
        Me.lblSymptom.Enabled = False: Me.txtSymptom.Enabled = False
    Else
        Me.lblSymptom.Enabled = True: Me.txtSymptom.Enabled = True
    End If
    
    If mblnSyncPage And InStr(cboKind.Text, "��Ժ���") > 0 Then '�������ͬ����ҳ��ϣ���ȡ��ҳ��ϡ���Ժ���顢��Ժ���
        For i = 1 To mobjDoc.Diagnosises.Count
            If mobjDoc.Diagnosises(i).��ҽ = Index And mobjDoc.Diagnosises(i).��ֹ�� = 0 Then
                lCount = lCount + 1
            End If
        Next
        
        With cboIn
            .Clear
            .AddItem "��"
            .AddItem "�ٴ�δȷ��"
            .AddItem "�������"
            .AddItem "��"
        End With
        gstrSQL = "Select ���� || '-' || ���� As ��Ժ��� From ���ƽ�� Order By ����"
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption)
        cboOut.Clear
        Do Until rsTemp.EOF
            cboOut.AddItem rsTemp!��Ժ���
            rsTemp.MoveNext
        Loop
            
        gstrSQL = "Select ����id, ���id, ֤��id, �������, ��Ժ����, ��Ժ���, �Ƿ�δ��, �Ƿ�����" & vbNewLine & _
                    "From ������ϼ�¼" & vbNewLine & _
                    "Where ����id = [1] And ��ҳid = [2] And ��¼��Դ=3 And �������=[3] And ������� = 1 And ��ϴ��� = [4]"
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, mobjDoc.EPRPatiRecInfo.����ID, mobjDoc.EPRPatiRecInfo.��ҳID, IIf(Index = 0, 3, 13), lCount + 1)
        If Not rsTemp.EOF Then
            chkDoubt.Value = NVL(rsTemp!�Ƿ�����, 0)
            Call zlControl.CboSetText(cboIn, NVL(rsTemp!��Ժ����))
            Call zlControl.CboSetText(cboOut, NVL(rsTemp!��Ժ���))
            
            Me.lblDisease.Tag = NVL(rsTemp!����id, 0) & "," & NVL(rsTemp!���id, 0)
            If Index = 0 Then '��ҽ���
                Me.lblKind.Tag = 0
                If InStr(NVL(rsTemp!�������), "(") > 0 Then '��ҳ�������������� (����)����
                    Me.txtDisease.Tag = Split(Split(NVL(rsTemp!�������), "(")(1), ")")(1): Me.txtDisease.Text = txtDisease.Tag
                Else '���������ֻ������
                    Me.txtDisease.Tag = NVL(rsTemp!�������): Me.txtDisease.Text = txtDisease.Tag
                End If
            Else '��ҽ���
                Me.lblKind.Tag = 1
                If UBound(Split(NVL(rsTemp!�������), "(")) > 1 Then
                    Me.txtDisease.Tag = Split(Split(NVL(rsTemp!�������), "(")(1), ")")(1)
                    Me.txtDisease.Text = Me.txtDisease.Tag

                    Me.lblSymptom.Tag = "" & NVL(rsTemp!֤��id, 0)
                    Me.txtSymptom.Tag = Split(Split(NVL(rsTemp!�������), "(")(2), ")")(0): Me.txtSymptom.Text = Me.txtSymptom.Tag
                Else
                    Me.txtDisease.Tag = Split(NVL(rsTemp!�������), "(")(0)
                    Me.txtDisease.Text = Me.txtDisease.Tag

                    Me.lblSymptom.Tag = "" & NVL(rsTemp!֤��id, 0)
                    Me.txtSymptom.Tag = Split(Split(NVL(rsTemp!�������), "(")(1), ")")(0): Me.txtSymptom.Text = Me.txtSymptom.Tag
                End If
            End If
        End If
        
        Me.LabIn.Visible = True: Me.cboIn.Visible = True
        Me.LabOut.Visible = True: Me.cboOut.Visible = True
    End If
    
    If Val(Me.lblKind.Tag) = 0 And Me.optType(0).Value = False Or Val(Me.lblKind.Tag) <> 0 And Me.optType(0).Value Then
        Me.lblDisease.Tag = "": Me.txtDisease.Tag = "": Me.txtDisease.Text = ""
        Me.lblSymptom.Tag = "": Me.txtSymptom.Tag = "": Me.txtSymptom.Text = ""
    End If
    Exit Sub

errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub optType_GotFocus(Index As Integer)
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub optType_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub txtDisease_Change()
    ValidControlText txtDisease
End Sub

Private Sub txtDisease_GotFocus()
    Me.txtDisease.SelStart = 0: Me.txtDisease.SelLength = 4000
    If Me.optHint(0).Value Then
        Call zlCommFun.OpenIme(True)
    Else
        Call zlCommFun.OpenIme(False)
    End If
End Sub

Private Sub txtDisease_KeyPress(KeyAscii As Integer)
Dim rsTemp As New ADODB.Recordset

    If KeyAscii <> vbKeyReturn Then Exit Sub
    
    If Me.optHint(0).Value Then
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    
    ElseIf Me.optHint(1).Value Then
        If Me.txtDisease.Tag = Trim(Me.txtDisease.Text) Or Trim(Me.txtDisease.Text) = "" Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
        gstrSQL = "Select Id As ����id, r.���id, l.����, l.����, l.����" & _
                " From ��������Ŀ¼ l, (Select ����id, Min(���id) As ���id From ������϶��� Group By ����id) r" & _
                " Where l.��� = [1] And l.Id = r.����id(+) And (l.���� Like [2] Or l.���� Like [3] Or l.���� Like [3])" & _
                " And (l.����ʱ�� is Null Or l.����ʱ��>=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                " Order By l.����"
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, _
            IIf(Me.optType(0).Value, "D", "B"), _
            UCase(Trim(Me.txtDisease.Text)) & "%", _
            gstrMatch & UCase(Trim(Me.txtDisease.Text)) & "%")
        If rsTemp.RecordCount <= 0 Then
            MsgBox "δ�ҵ�Ҫ��ı�׼�������룡", vbExclamation, gstrSysName: Exit Sub
        ElseIf rsTemp.RecordCount = 1 Then
            Me.lblDisease.Tag = rsTemp!����id & "," & rsTemp!���id
            Me.txtDisease.Tag = rsTemp!����: Me.txtDisease.Text = rsTemp!����
        Else
            With Me.vfgSelect
                .Tag = "D"
                .Clear
                Set .DataSource = rsTemp
                .ColHidden(0) = True: .ColHidden(1) = True
                .Row = .FixedRows
                .Move Me.txtDisease.Left, Me.txtDisease.Top + Me.txtDisease.Height, Me.txtDisease.Width
                .Visible = True
                .SetFocus
            End With
        End If
    
    ElseIf Me.optHint(2).Value Then
        If Me.txtDisease.Tag = Trim(Me.txtDisease.Text) Or Trim(Me.txtDisease.Text) = "" Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
        gstrSQL = "Select r.����id, n.���id, l.����, n.����, n.����" & _
                " From �������Ŀ¼ l, ������ϱ��� n, (Select ���id, Min(����id) As ����id From ������϶��� Group By ���id) r" & _
                " Where l.Id = n.���id And l.��� = [1] And l.Id = r.���id(+) And" & _
                "       (l.���� Like [2] Or n.���� Like [3] Or n.���� Like [3])" & _
                " And (l.����ʱ�� is Null Or l.����ʱ��>=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                " Order By l.����"
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, _
            IIf(Me.optType(0).Value, 1, 2), _
            UCase(Trim(Me.txtDisease.Text)) & "%", _
            gstrMatch & UCase(Trim(Me.txtDisease.Text)) & "%")
        If rsTemp.RecordCount <= 0 Then
            MsgBox "δ�ҵ�Ҫ��ļ��������Ŀ��", vbExclamation, gstrSysName: Exit Sub
        ElseIf rsTemp.RecordCount = 1 Then
            Me.lblDisease.Tag = rsTemp!����id & "," & rsTemp!���id
            Me.txtDisease.Tag = rsTemp!����: Me.txtDisease.Text = rsTemp!����
        Else
            With Me.vfgSelect
                .Tag = "D"
                .Clear
                Set .DataSource = rsTemp
                .ColHidden(0) = True: .ColHidden(1) = True
                .Row = .FixedRows
                .Move Me.txtDisease.Left, Me.txtDisease.Top + Me.txtDisease.Height, Me.txtDisease.Width
                .Visible = True
                .SetFocus
            End With
        End If
    End If
End Sub

Private Sub txtSymptom_Change()
    ValidControlText txtSymptom
End Sub

Private Sub txtSymptom_GotFocus()
    Me.txtSymptom.SelStart = 0: Me.txtSymptom.SelLength = 4000
    If Me.optHint(0).Value Then
        Call zlCommFun.OpenIme(True)
    Else
        Call zlCommFun.OpenIme(False)
    End If
End Sub

Private Sub txtSymptom_KeyPress(KeyAscii As Integer)
Dim rsTemp As New ADODB.Recordset
    If KeyAscii <> vbKeyReturn Then Exit Sub
    
    If Me.optHint(0).Value Then
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    
    ElseIf Me.optHint(1).Value Then
        If Me.txtSymptom.Tag = Trim(Me.txtSymptom.Text) Or Trim(Me.txtSymptom.Text) = "" Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
        gstrSQL = "Select Id As ֤��id, ����, ����, ����" & _
                " From ��������Ŀ¼" & _
                " Where ��� = 'Z' And (���� Like [1] Or ���� Like [2] Or ���� Like [2])" & _
                " And (����ʱ�� is Null Or ����ʱ��>=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                " Order By ����"
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, UCase(Trim(Me.txtSymptom.Text)) & "%", gstrMatch & UCase(Trim(Me.txtSymptom.Text)) & "%")
        If rsTemp.RecordCount <= 0 Then
            MsgBox "δ�ҵ�Ҫ��ı�׼��ҽ֤��", vbExclamation, gstrSysName: Exit Sub
        ElseIf rsTemp.RecordCount = 1 Then
            Me.lblSymptom.Tag = "" & rsTemp!֤��id
            Me.txtSymptom.Tag = rsTemp!����: Me.txtSymptom.Text = rsTemp!����
        Else
            With Me.vfgSelect
                .Tag = "S"
                .Clear
                Set .DataSource = rsTemp
                .ColHidden(0) = True
                .Row = .FixedRows
                .Move Me.txtSymptom.Left, Me.txtSymptom.Top + Me.txtSymptom.Height, Me.txtSymptom.Width
                .Visible = True
                .SetFocus
            End With
        End If
    Else
        If Me.txtSymptom.Tag = Trim(Me.txtSymptom.Text) And Trim(Me.txtSymptom.Text) <> "" Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
        
        Dim aryDisease() As String, lngDisease As Long
        aryDisease = Split(Me.lblDisease.Tag, ",")
        If UBound(aryDisease) < 1 Then
            lngDisease = 0
        Else
            lngDisease = Val(aryDisease(1))
        End If
        gstrSQL = "Select Distinct ֤��id, ֤����� As ���, ֤������ As ����, Zlspellcode(֤������) As ����" & _
                " From ������ϲο�" & _
                " Where ���id = [1] And ֤����� Is Not Null And (֤������ Like [2] Or Zlspellcode(֤������) Like [2])" & _
                " Order By ֤�����"
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngDisease, gstrMatch & UCase(Trim(Me.txtSymptom.Text)) & "%")
        If rsTemp.RecordCount <= 0 Then
            MsgBox "δ�ҵ�Ҫ��ĵ�ǰ��ҽ֤��", vbExclamation, gstrSysName: Exit Sub
        ElseIf rsTemp.RecordCount = 1 Then
            Me.lblSymptom.Tag = "" & rsTemp!֤��id
            Me.txtSymptom.Tag = rsTemp!����: Me.txtSymptom.Text = rsTemp!����
        Else
            With Me.vfgSelect
                .Tag = "S"
                .Clear
                Set .DataSource = rsTemp
                .ColHidden(0) = True
                .Row = .FixedRows
                .Move Me.txtSymptom.Left, Me.txtSymptom.Top + Me.txtSymptom.Height, Me.txtSymptom.Width
                .Visible = True
                .SetFocus
            End With
        End If
    End If
End Sub

Private Sub vfgSelect_DblClick()
    With Me.vfgSelect
        Select Case .Tag
        Case "D"
            Me.lblDisease.Tag = Val(.TextMatrix(.Row, 0)) & "," & Val(.TextMatrix(.Row, 1))
            Me.txtDisease.Tag = .TextMatrix(.Row, 3): Me.txtDisease.Text = .TextMatrix(.Row, 3)
            'Ϊ��֤��ҽ��֤�Ǻϣ����֤��ȴ���������
            Me.lblSymptom.Tag = "": Me.txtSymptom.Tag = "": Me.txtSymptom.Text = ""
        Case "S"
            Me.lblSymptom.Tag = Val(.TextMatrix(.Row, 0))
            Me.txtSymptom.Tag = .TextMatrix(.Row, 2): Me.txtSymptom.Text = .TextMatrix(.Row, 2)
        End Select
        .Visible = False
    End With
End Sub

Private Sub vfgSelect_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call vfgSelect_DblClick
    End If
End Sub

Private Sub vfgSelect_LostFocus()
    With Me.vfgSelect
        .Visible = False
        Select Case .Tag
        Case "D": Me.txtDisease.SetFocus
        Case "S": Me.txtSymptom.SetFocus
        End Select
    End With
End Sub
