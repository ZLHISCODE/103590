VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLisPrinterSetup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "LIS������ӡ����"
   ClientHeight    =   5505
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7710
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLisPrinterSetup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5505
   ScaleWidth      =   7710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame frame3 
      Height          =   1095
      Left            =   0
      TabIndex        =   18
      Top             =   3840
      Width           =   7695
      Begin VB.CheckBox chkBack 
         Caption         =   "��ӡ����󷵻���ҳ"
         Height          =   345
         Left            =   3180
         TabIndex        =   31
         Top             =   300
         Width           =   1935
      End
      Begin VB.TextBox txtClear 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   6540
         MaxLength       =   3
         TabIndex        =   28
         Text            =   "0"
         Top             =   330
         Width           =   405
      End
      Begin VB.TextBox txtPrintDelayed 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   6690
         MaxLength       =   3
         TabIndex        =   25
         Text            =   "0"
         Top             =   690
         Width           =   405
      End
      Begin VB.CheckBox chkGoBack 
         Caption         =   "������ӡ��ʾ[����]��ť"
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   720
         Width           =   2265
      End
      Begin VB.TextBox txtDays 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   3720
         MaxLength       =   3
         TabIndex        =   23
         Text            =   "0"
         Top             =   720
         Width           =   255
      End
      Begin VB.CheckBox chkTimeHorizon 
         Caption         =   "����ʾ       ���ڵı��浥"
         Height          =   195
         Left            =   2850
         TabIndex        =   22
         Top             =   750
         Width           =   2295
      End
      Begin VB.CommandButton cmdPage 
         Caption         =   "��"
         Height          =   240
         Left            =   2880
         TabIndex        =   21
         Top             =   330
         Width           =   255
      End
      Begin VB.TextBox txtPage 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   1770
         MaxLength       =   20
         TabIndex        =   19
         Top             =   307
         Width           =   1395
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�����"
         Height          =   195
         Left            =   6960
         TabIndex        =   30
         Top             =   360
         Width           =   540
      End
      Begin VB.Line Line2 
         X1              =   6510
         X2              =   6900
         Y1              =   540
         Y2              =   540
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ӡ�����"
         Height          =   195
         Left            =   5580
         TabIndex        =   29
         Top             =   360
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ӡ������ʱ"
         Height          =   195
         Left            =   5580
         TabIndex        =   27
         Top             =   720
         Width           =   1080
      End
      Begin VB.Line Line1 
         X1              =   6660
         X2              =   7050
         Y1              =   900
         Y2              =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��/��"
         Height          =   195
         Left            =   7080
         TabIndex        =   26
         Top             =   720
         Width           =   420
      End
      Begin VB.Line Line 
         X1              =   3435
         X2              =   3735
         Y1              =   915
         Y2              =   915
      End
      Begin VB.Label lblHelpPage 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������ӡ����ҳ��"
         Height          =   195
         Left            =   240
         TabIndex        =   20
         Top             =   360
         Width           =   1440
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   345
      Left            =   6090
      TabIndex        =   14
      Top             =   5100
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Caption         =   "���Ƶ���"
      Height          =   2535
      Left            =   0
      TabIndex        =   13
      Top             =   1290
      Width           =   7695
      Begin VB.CommandButton cmdClearAll 
         Caption         =   "ȫ��(&L)"
         Height          =   375
         Left            =   6300
         TabIndex        =   17
         Top             =   1050
         Width           =   1095
      End
      Begin VB.CommandButton cmdSelectAll 
         Caption         =   "ȫѡ(&A)"
         Height          =   375
         Left            =   6300
         TabIndex        =   16
         Top             =   450
         Width           =   1095
      End
      Begin MSComctlLib.ListView lveList 
         Height          =   2235
         Left            =   60
         TabIndex        =   15
         Top             =   210
         Width           =   5955
         _ExtentX        =   10504
         _ExtentY        =   3942
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "���"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "����"
            Object.Width           =   5292
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "������Դ"
      Height          =   1245
      Left            =   4380
      TabIndex        =   9
      Top             =   0
      Width           =   3315
      Begin VB.CheckBox chk 
         Caption         =   "���"
         Height          =   315
         Index           =   2
         Left            =   2280
         TabIndex        =   12
         Top             =   480
         Width           =   735
      End
      Begin VB.CheckBox chk 
         Caption         =   "סԺ"
         Height          =   315
         Index           =   1
         Left            =   1260
         TabIndex        =   11
         Top             =   480
         Width           =   735
      End
      Begin VB.CheckBox chk 
         Caption         =   "����"
         Height          =   315
         Index           =   0
         Left            =   300
         TabIndex        =   10
         Top             =   480
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   345
      Left            =   4530
      TabIndex        =   7
      Top             =   5100
      Width           =   1095
   End
   Begin VB.Frame fraMain 
      Caption         =   "����ѡ��"
      Height          =   1245
      Left            =   0
      MousePointer    =   1  'Arrow
      TabIndex        =   0
      Top             =   0
      Width           =   4335
      Begin VB.CommandButton cmdMachine 
         Caption         =   "�豸����(&M)"
         Height          =   345
         Left            =   2790
         MousePointer    =   1  'Arrow
         TabIndex        =   8
         Top             =   660
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.OptionButton OptItem 
         Caption         =   "����"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   6
         Top             =   300
         Value           =   -1  'True
         Width           =   885
      End
      Begin VB.OptionButton OptItem 
         Caption         =   "סԺ��"
         Height          =   255
         Index           =   1
         Left            =   1230
         MousePointer    =   1  'Arrow
         TabIndex        =   5
         Top             =   300
         Width           =   885
      End
      Begin VB.OptionButton OptItem 
         Caption         =   "�����"
         Height          =   255
         Index           =   2
         Left            =   2220
         MousePointer    =   1  'Arrow
         TabIndex        =   4
         Top             =   300
         Width           =   885
      End
      Begin VB.OptionButton OptItem 
         Caption         =   "���￨"
         Height          =   255
         Index           =   3
         Left            =   3210
         MousePointer    =   1  'Arrow
         TabIndex        =   3
         Top             =   300
         Width           =   885
      End
      Begin VB.OptionButton OptItem 
         Caption         =   "IC�����֤"
         Height          =   255
         Index           =   4
         Left            =   1230
         MousePointer    =   1  'Arrow
         TabIndex        =   2
         Top             =   750
         Width           =   1305
      End
      Begin VB.OptionButton OptItem 
         Caption         =   "����ID"
         Height          =   255
         Index           =   5
         Left            =   240
         MousePointer    =   1  'Arrow
         TabIndex        =   1
         Top             =   750
         Width           =   885
      End
   End
End
Attribute VB_Name = "frmLisPrinterSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mlngPageKey As Long

Private Sub chkTimeHorizon_Click()
    If txtDays.Enabled = False Then
        txtDays.Enabled = True
    Else
        txtDays.Enabled = False
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdClearAll_Click()
    SelectChk False
End Sub

Private Sub cmdMachine_Click()
    Call zlCommFun.DeviceSetup(Me, 100, 1101)
End Sub

Private Sub cmdOK_Click()
    Dim intLoop As Integer
    Dim str��� As String
    For intLoop = 0 To Me.OptItem.UBound
        If Me.OptItem(intLoop).Value = True Then
            Exit For
        End If
    Next
    Call SaveSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\frmPrinterSetup", "���ҷ�ʽ", intLoop)
    Call SaveSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\frmPrinterSetup", "������Դ", _
                        chk(0).Value & "," & chk(1).Value & "," & chk(2).Value)
    For intLoop = 1 To Me.lveList.ListItems.Count
        If Me.lveList.ListItems(intLoop).Checked = True Then
            str��� = str��� & "," & Mid(Me.lveList.ListItems(intLoop).Key, 2)
        End If
    Next
    str��� = str��� & ","
    Call SaveSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\frmPrinterSetup", "���Ƶ���", str���)
    Call SaveSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\frmPrinterSetup", "��ӡ��������", Val(txtClear.Text))
    Call SaveSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\frmPrinterSetup", "��ӡ����󷵻���ҳ", chkBack.Value)
    
    
    Call SetPara("������ӡ����ҳ��", mlngPageKey)
    Call SetPara("�������ڷ�Χ", chkTimeHorizon.Value & "-" & Val(Trim(txtDays.Text)))
    Call SetPara("������ӡ��ʾ���ذ�ť", chkGoBack.Value)
    Call SetPara("�����ӡ��ʱ", Val(txtPrintDelayed.Text))
    Unload Me
End Sub
Private Sub cmdPage_Click()
    Dim strID As String
    Dim blnRe As Boolean
    Dim str���� As String
    Dim str���� As String
        
    gstrSQL = "Select ҳ����� AS id,�ϼ���� AS �ϼ�id,ҳ������ AS ����,����,ĩ�� From ��ѯҳ��Ŀ¼ where ҳ�����>0 Start with �ϼ���� is null connect by prior ҳ����� =�ϼ����"
    str���� = txtPage.Text
                
'    strID = CStr(mlngPageKey)
'    strID = IIf(Val(strID) = 0, "", strID)
    
    blnRe = frm����ѡ��.ShowTree(gstrSQL, strID, str����, str����, "", Me.Caption, "����ҳ�����", , "", True)
    
    If blnRe Then       '�µı����Ŀ��
        txtPage.Text = str����
        mlngPageKey = Val(strID)
        txtPage.ForeColor = &HFF0000
        txtPage.BackColor = &HE0E0E0
        txtPage.Tag = ""
    End If
    
End Sub

Private Sub cmdSelectAll_Click()
    SelectChk True
End Sub

Private Sub Form_Load()
    Dim intLoop As Integer
    Dim strSource As String
    intLoop = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\frmPrinterSetup", "���ҷ�ʽ", 0)
    Me.OptItem(intLoop).Value = True
    strSource = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\frmPrinterSetup", "������Դ", "0,0,0")
    chk(0).Value = Split(strSource, ",")(0)
    chk(1).Value = Split(strSource, ",")(1)
    chk(2).Value = Split(strSource, ",")(2)
    
    mlngPageKey = GetPara("������ӡ����ҳ��", -5)
    If mlngPageKey > 0 Then
        gstrSQL = "Select ҳ������ From ��ѯҳ��Ŀ¼ Where ҳ�����=[1]"
        Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngPageKey)
        If gRs.RecordCount > 0 Then
            txtPage.Text = Trim(Nvl(gRs!ҳ������))
        End If
    Else
        txtPage.Text = ""
    End If
    
    txtPage.ForeColor = &HFF0000
    txtPage.BackColor = &HE0E0E0
    txtPage.Tag = ""
    
    chkTimeHorizon.Value = Split(GetPara("�������ڷ�Χ", 0), "-")(0)
    txtDays.Text = Split(GetPara("�������ڷ�Χ", 0), "-")(1)
    txtDays.Enabled = chkTimeHorizon.Value
    
    chkGoBack.Value = GetPara("������ӡ��ʾ���ذ�ť", 0)
    txtPrintDelayed.Text = Val(GetPara("�����ӡ��ʱ", 0))
    txtClear.Text = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\frmPrinterSetup", "��ӡ��������", 0)
    chkBack.Value = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\frmPrinterSetup", "��ӡ����󷵻���ҳ", 0)
    initVsf
End Sub

Private Sub OptItem_Click(Index As Integer)
    Me.cmdMachine.Visible = (Index = 3 Or Index = 4)
End Sub

Private Sub txtDays_GotFocus()
    Call SelAll(txtDays)
End Sub

Private Sub txtDays_KeyPress(KeyAscii As Integer)
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

Private Sub txtPage_Change()
    txtPage.Tag = "Changed"
    txtPage.ForeColor = &H0&
    txtPage.BackColor = &H80000005
End Sub

Private Sub txtPage_GotFocus()
    Call SelAll(txtPage)
End Sub

Private Sub txtPage_KeyPress(KeyAscii As Integer)

    Dim strInput As String

    Dim strColWidth As String

    Dim strColAlign As String

    Dim lngPostion  As Long

    Dim sglX As Single

    Dim sglY As Single
    
    If KeyAscii = vbKeyReturn Then
        If txtPage.Tag = "Changed" Then
            If InStr(txtPage.Text, "'") > 0 Then
                MsgBox "�������зǷ��ַ���", vbInformation, gstrSysName

                Exit Sub

            End If
            
            strInput = "'%" & txtPage.Text & "%'"
            
            gstrSQL = "Select ����,ҳ������ AS ����,����,ҳ����� From ��ѯҳ��Ŀ¼  where ҳ�����>0 and ĩ��=1 AND (���� Like [1] or ���� Like [1] or ҳ������ Like [1])"
            Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strInput)

            If gRs.BOF = False Then
                If gRs.RecordCount = 1 Then
                    lngPostion = 1
                Else
                    strColWidth = "900;2300;900;0"
                    strColAlign = "1;1;1;1"
                    Call CalcXY(Me, txtPage.Left + 30, txtPage.Top + txtPage.Height, sglX, sglY)
                    lngPostion = frmSelectList.ShowSelectList(Me, sglX, sglY, 4800, 2400, gRs, strColWidth, strColAlign)
                End If

                If lngPostion > 0 Then
                    gRs.MoveFirst
                    gRs.Move lngPostion - 1
                                    
                    txtPage.Text = IIf(IsNull(gRs("����")), "", gRs("����"))
                    mlngPageKey = IIf(IsNull(gRs("ҳ�����")), 0, gRs("ҳ�����"))
                Else
                    mlngPageKey = 0
                    txtPage.Text = ""
                End If
                
            Else
                mlngPageKey = 0
                txtPage.Text = ""
            End If

            txtPage.ForeColor = &HFF0000
            txtPage.BackColor = &HE0E0E0
            txtPage.Tag = ""
        Else
            SendKeys "{TAB}"
            SendKeys "{TAB}"
        End If

    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If

End Sub

Private Sub txtPage_Validate(Cancel As Boolean)
    If txtPage.Tag = "Changed" Then Cancel = True
End Sub

Private Sub initVsf()
    Dim rsTmp As New ADODB.Recordset
    Dim intRow As Integer
    Dim Item As ListItem
    Dim str��� As String
    
    '��ʹ���б�
    str��� = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\frmPrinterSetup", "���Ƶ���", "")
    
    On Error GoTo errH
    intRow = 0
    gstrSQL = "Select Distinct c.���, 'ZLCISBILL' || Trim(To_Char(C.���, '00000')) || '-2' As ������, C.����, C.˵��" & vbNewLine & _
            "From ������ĿĿ¼ A, ��������Ӧ�� B, �����ļ��б� C" & vbNewLine & _
            "Where ��� = 'C' And A.Id = B.������Ŀid And B.�����ļ�id = C.Id"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    Do Until rsTmp.EOF
        intRow = intRow + 1
        With Me.lveList
            Set Item = .ListItems.Add(intRow, "A" & rsTmp("���") & "", rsTmp("������") & "")
            Item.SubItems(1) = rsTmp("����") & ""
            If InStr(str���, "," & rsTmp("���") & ",") > 0 Then
                Item.Checked = True
            End If
        End With
        rsTmp.MoveNext
    Loop
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub SelectChk(blnSelect As Boolean)
    Dim intRow As Integer
    With Me.lveList
        For intRow = 1 To .ListItems.Count
            .ListItems(intRow).Checked = blnSelect
        Next
    End With
End Sub



Private Sub CalcXY(objFrm As Form, objX As Single, objY As Single, sglX As Single, sglY As Single)
    sglX = objFrm.Left + objX + Screen.TwipsPerPixelX
    sglY = objFrm.Top + objFrm.Height - objFrm.ScaleHeight + objY
    If sglX + 6030 > Screen.Width Then
        sglX = Screen.Width - 6030
    End If
    If sglY + 3195 > Screen.Height Then
        sglY = sglY - 3195
    End If
End Sub


