VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDefTreePage 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ѡ��ҳ��"
   ClientHeight    =   3420
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5250
   Icon            =   "frmDefTreePage.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   5250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin MSComctlLib.ImageList ils16 
      Left            =   4290
      Top             =   2160
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
            Picture         =   "frmDefTreePage.frx":000C
            Key             =   "default"
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   3345
      Left            =   60
      TabIndex        =   21
      Top             =   30
      Width           =   3600
      Begin VB.CommandButton cmdPage 
         Caption         =   "��"
         Height          =   240
         Left            =   3165
         TabIndex        =   4
         Top             =   570
         Width           =   255
      End
      Begin VB.TextBox txtPage 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   795
         MaxLength       =   20
         TabIndex        =   3
         Top             =   540
         Width           =   2655
      End
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   1
         Left            =   2655
         TabIndex        =   10
         Text            =   "cbo"
         Top             =   1260
         Width           =   795
      End
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   2
         Left            =   795
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1275
         Width           =   1125
      End
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   0
         Left            =   795
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   900
         Width           =   2655
      End
      Begin VB.CommandButton cmdOpen 
         Caption         =   "��"
         Height          =   240
         Left            =   1635
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   1665
         Width           =   255
      End
      Begin VB.PictureBox Picture1 
         Height          =   930
         Left            =   795
         ScaleHeight     =   870
         ScaleWidth      =   2595
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   2340
         Width           =   2655
         Begin VB.Label lblSample 
            AutoSize        =   -1  'True
            Caption         =   "ʾ��˵��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   210
            Left            =   15
            TabIndex        =   22
            Top             =   30
            Width           =   840
         End
      End
      Begin VB.TextBox txt 
         Height          =   300
         Left            =   795
         MaxLength       =   20
         TabIndex        =   1
         Top             =   180
         Width           =   2655
      End
      Begin VB.ComboBox cboIcon 
         Height          =   300
         Left            =   795
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   1980
         Width           =   1710
      End
      Begin zl9NewQuery.ctlPicture UsrPicture 
         Height          =   435
         Index           =   1
         Left            =   2865
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   1830
         Width           =   570
         _ExtentX        =   1005
         _ExtentY        =   767
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "����(&3)"
         Height          =   180
         Left            =   135
         TabIndex        =   5
         Top             =   975
         Width           =   630
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "����(&4)"
         Height          =   180
         Left            =   135
         TabIndex        =   7
         Top             =   1350
         Width           =   630
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "��С(&5)"
         Height          =   180
         Left            =   1995
         TabIndex        =   9
         Top             =   1320
         Width           =   630
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "ʾ��(&8)"
         Height          =   180
         Left            =   135
         TabIndex        =   17
         Top             =   2355
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "��ɫ(&6)"
         Height          =   180
         Left            =   135
         TabIndex        =   11
         Top             =   1710
         Width           =   630
      End
      Begin VB.Label lblColor 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "�~�~�~�~"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   780
         TabIndex        =   12
         Top             =   1635
         Width           =   1140
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "ҳ��(&2)"
         Height          =   180
         Left            =   120
         TabIndex        =   2
         Top             =   585
         Width           =   630
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "����(&1)"
         Height          =   180
         Left            =   120
         TabIndex        =   0
         Top             =   225
         Width           =   630
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "ͼ��(&7)"
         Height          =   180
         Left            =   135
         TabIndex        =   14
         Top             =   2025
         Width           =   630
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3885
      TabIndex        =   20
      Top             =   585
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   3885
      TabIndex        =   19
      Top             =   135
      Width           =   1100
   End
   Begin MSComDlg.CommonDialog dlg 
      Left            =   4065
      Top             =   1545
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmDefTreePage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnFirst As Boolean
Private blnOK As Boolean
Private mKey As Long
Private mUpKey As Long
Private mlngPageKey As Long

Private Sub ApplySample()
    '---------------------------------------------------------------------------------------
    '���ܣ�
    '������
    '���أ�
    '---------------------------------------------------------------------------------------
    
    On Error GoTo errHand
    
    With lblSample
        .FontName = cbo(0).Text
        .FontSize = IIf(Val(cbo(1).Text) <= 0, 12, Val(cbo(1).Text))
        .FontBold = (cbo(2).Text = "����" Or cbo(2).Text = "��б��")
        .FontItalic = (cbo(2).Text = "б��" Or cbo(2).Text = "��б��")
        .ForeColor = lblColor.ForeColor
        .Caption = IIf(txt.Text = "", "ʾ��˵��", txt.Text)
    End With
    
    Exit Sub
    
errHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Function CheckValid() As Boolean
    '---------------------------------------------------------------------------------------
    '���ܣ�����������ݵ���Ч��
    '���أ����Ϸ�����True;���򷵻�False
    '---------------------------------------------------------------------------------------
    Dim strTmp As String
    Dim blnWarn As Boolean
                
    If Trim(txt.Text) = "" Then
        strTmp = "���붨��һ����ʾ���ƣ�"
        GoTo Finally
    End If
    
    If mlngPageKey <= 0 Then
        strTmp = "����Ҫѡ��һ����Ӧҳ�棡"
        GoTo Finally
    End If
        
    '����8������ʱ����������ʾʱ���ܲ�����
    If LenB(StrConv(txt.Text, vbFromUnicode)) > 16 Then
        strTmp = "���ⳬ����8�����ֻ�16����ĸ����ʾʱ���ܲ�����ȫ��ʾ��"
        blnWarn = True
        GoTo Finally
    End If
    
    If Val(cbo(1).Text) <= 0 Or Val(cbo(1).Text) > 99 Then
        strTmp = "�����������0��С��99�����֣�"
        GoTo Finally
    End If
    
Finally:
    If strTmp <> "" Then MsgBox strTmp, vbInformation, gstrSysName
            
    '��������Ǿ���,����Ȼ�Ϸ�
    CheckValid = (strTmp = "" Or blnWarn)
    
End Function

Private Sub cbo_Change(Index As Integer)
    cmdOK.Tag = "1"
    Call ApplySample
End Sub

Private Sub cbo_Click(Index As Integer)
    cmdOK.Tag = "1"
    Call ApplySample
End Sub

Private Sub cbo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub cbocboPage_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
        Exit Sub
    End If
End Sub

Private Sub cboIcon_Click()
    If mblnFirst Then Exit Sub
    
    '����ͼƬ��ʾ
    Dim rs As New ADODB.Recordset
    
    UsrPicture(1).Cls
    
    If cboIcon.ItemData(cboIcon.ListIndex) > 0 Then
        gstrSQL = "select ���,���,�߶�,���� from ��ѯͼƬԪ�� where ���=[1]"
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(cboIcon.ItemData(cboIcon.ListIndex)))
        If rs.BOF = False Then
            Call UsrPicture(1).ShowPictureByFieldNew(rs!���, rs!��� * Screen.TwipsPerPixelX, rs!�߶� * Screen.TwipsPerPixelY, IIf(IsNull(rs!����), 0, rs!����))
        End If
        CloseRecord rs
    Else
        Call UsrPicture(1).ShowByIPictureDisp(ils16.ListImages(1).Picture, 16 * Screen.TwipsPerPixelX, 16 * Screen.TwipsPerPixelY)
    End If
    cmdOK.Tag = "1"
End Sub

Private Sub cboIcon_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If SaveData Then
        blnOK = True
        If mKey = 0 Then
            txt.Text = ""
            txt.SetFocus
            cmdOK.Tag = ""
        Else
            cmdOK.Tag = ""
            Unload Me
        End If
    End If
End Sub

Private Sub cmdOpen_Click()
    '����ָ����Ԫ���������ɫ,����һ��ָ�������Ԫ��
    Dim i As Long
    Dim j As Long
    
    On Error Resume Next
    dlg.CancelError = True
    dlg.flags = &H1
    dlg.Color = lblColor.ForeColor
    dlg.ShowColor
    If Err.Number = 0 Then
        lblColor.ForeColor = dlg.Color
        Call ApplySample
        cmdOK.Tag = "1"
    Else
        Err.Clear
    End If
End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmdPage_Click()
    Dim strID As String
    Dim blnRe As Boolean
    Dim str���� As String
    Dim str���� As String
        
    gstrSQL = "Select ҳ����� AS id,�ϼ���� AS �ϼ�id,ҳ������ AS ����,����,ĩ�� From ��ѯҳ��Ŀ¼ where ҳ�����>0 Start with �ϼ���� is null connect by prior ҳ����� =�ϼ����"
    str���� = txtPage.Text
                
    strID = CStr(mlngPageKey)
    strID = IIf(Val(strID) = 0, "", strID)
    
    blnRe = frm����ѡ��.ShowTree(gstrSQL, strID, str����, str����, "", Me.Caption, "����ҳ�����", , "", True)
    
    If blnRe Then       '�µı����Ŀ��
        txtPage.Text = str����
        mlngPageKey = Val(strID)
        txtPage.ForeColor = &HFF0000
        txtPage.BackColor = &HE0E0E0
        txtPage.Tag = ""
        
        cmdOK.Tag = "1"
    End If
    
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    
    Dim lngLoop As Long
    Dim rsFont As New ADODB.Recordset
    
    With rsFont
        .Fields.Append "��������", adVarChar, 100
        .Open
        .ActiveConnection = Nothing
    End With
    
    For lngLoop = 0 To Screen.FontCount - 1
        rsFont.AddNew
        rsFont.Fields(0).Value = Screen.Fonts(lngLoop)
    Next
        
    rsFont.Sort = "��������"
    
    If rsFont.RecordCount > 0 Then
        rsFont.MoveFirst
        Do While Not rsFont.EOF
            cbo(0).AddItem rsFont.Fields(0).Value
            rsFont.MoveNext
        Loop
    End If
    
    cbo(0).Text = "����"
    
    cbo(1).AddItem "8"
    cbo(1).AddItem "9"
    cbo(1).AddItem "10"
    cbo(1).AddItem "11"
    cbo(1).AddItem "12"
    cbo(1).AddItem "14"
    cbo(1).AddItem "16"
    cbo(1).AddItem "18"
    cbo(1).AddItem "20"
    cbo(1).AddItem "22"
    cbo(1).AddItem "24"
    cbo(1).AddItem "26"
    cbo(1).AddItem "28"
    cbo(1).AddItem "36"
    cbo(1).AddItem "48"
    cbo(1).AddItem "72"
    cbo(1).ListIndex = 4
    
    cbo(2).AddItem "����"
    cbo(2).AddItem "б��"
    cbo(2).AddItem "����"
    cbo(2).AddItem "��б��"
    cbo(2).ListIndex = 0
    
    '����ͼ�꼯��
    cboIcon.AddItem "ȱʡͼ��"
    gstrSQL = "select ����,��� from ��ѯͼƬԪ�� where ����=3"
    Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    If gRs.BOF = False Then
        While Not gRs.EOF
            cboIcon.AddItem IIf(IsNull(gRs!����), "", gRs!����)
            cboIcon.ItemData(cboIcon.NewIndex) = gRs!���
            gRs.MoveNext
        Wend
    End If
    cboIcon.ListIndex = 0
    
    If mKey <> 0 Then
        gstrSQL = "select A.����,A.ҳ��,A.ҳ��ͼ��,A.����,A.��С,A.����,A.��ɫ,A.���,A.�����,B.ҳ������,B.ҳ����� from ��ѯҳ������ A,��ѯҳ��Ŀ¼ B where A.ҳ��=B.ҳ����� and A.���=[1]"
        Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mKey)
        If gRs.BOF = False Then
            txt.Text = IIf(IsNull(gRs!����), "", gRs!����)
            'cboPage.Text = IIf(IsNull(gRs!ҳ������), "", gRs!ҳ������)
            
            txtPage.Text = IIf(IsNull(gRs!ҳ������), "", gRs!ҳ������)
            mlngPageKey = IIf(IsNull(gRs!ҳ�����), 0, gRs!ҳ�����)
            
            cboIcon.ListIndex = FindCboIndex(cboIcon, IIf(IsNull(gRs!ҳ��ͼ��), 0, gRs!ҳ��ͼ��))
            If cboIcon.ListIndex = -1 Then cboIcon.ListIndex = 0
            
            On Error Resume Next
            lblColor.ForeColor = IIf(IsNull(gRs!��ɫ), &HFF0000, gRs!��ɫ)
            cbo(0).Text = IIf(IsNull(gRs!����), "����", gRs!����)
            cbo(1).Text = IIf(IsNull(gRs!��С), "12", gRs!��С)
            
            Select Case IIf(IsNull(gRs!����), 1, gRs!����)
            Case 1
                cbo(2).ListIndex = 0
            Case 2
                cbo(2).ListIndex = 1
            Case 3
                cbo(2).ListIndex = 2
            Case 4
                cbo(2).ListIndex = 3
            End Select
            
            If cbo(0).ListCount > 0 And cbo(0).ListIndex = -1 Then cbo(0).Text = "����"
            If cbo(2).ListCount > 0 And cbo(2).ListIndex = -1 Then cbo(2).Text = "����"
        End If
    End If
    mblnFirst = False
    
    Call cboIcon_Click
    
    txtPage.ForeColor = &HFF0000
    txtPage.BackColor = &HE0E0E0
    txtPage.Tag = ""
    
    cmdOK.Tag = ""
    Call SelAll(txt)
End Sub

Private Sub Form_Load()
    blnOK = False
    mblnFirst = True
End Sub

Public Function ShowTreePageBox(frmMain As Form, ByVal Key As Long, ByVal UpKey As Long) As Boolean
    mKey = Key
    mUpKey = UpKey
    frmDefTreePage.Show 1, frmMain
    ShowTreePageBox = blnOK
    
End Function

Private Function SaveData() As Boolean
    Dim lng��� As Long
    Dim bytFontForm As Byte
    
    If cmdOK.Tag = "1" Then
        
        If CheckValid = False Then Exit Function
        
        Select Case cbo(2).Text
        Case "����"
            bytFontForm = 1
        Case "б��"
            bytFontForm = 2
        Case "����"
            bytFontForm = 3
        Case "��б��"
            bytFontForm = 4
        End Select
        
        If mKey = 0 Then

            lng��� = NextValue("��ѯҳ������", "���")
            gstrSQL = "zl_��ѯҳ������_insert(" & lng��� & "," & IIf(mUpKey = 0, "NULL", mUpKey) & ",'" & txt.Text & "'," & mlngPageKey & "," & IIf(cboIcon.ItemData(cboIcon.ListIndex) = 0, "NULL", cboIcon.ItemData(cboIcon.ListIndex)) & ",'" & _
                                            IIf(cbo(0).Text = "", "����", cbo(0).Text) & "'," & _
                                            IIf(Val(cbo(1).Text) <= 0, 12, Val(cbo(1).Text)) & "," & _
                                            bytFontForm & "," & _
                                            lblColor.ForeColor & ")"
        Else
            lng��� = mKey
            gstrSQL = "zl_��ѯҳ������_update(" & lng��� & "," & IIf(mUpKey = 0, "NULL", mUpKey) & ",'" & txt.Text & "'," & mlngPageKey & "," & IIf(cboIcon.ItemData(cboIcon.ListIndex) = 0, "NULL", cboIcon.ItemData(cboIcon.ListIndex)) & ",'" & _
                                            IIf(cbo(0).Text = "", "����", cbo(0).Text) & "'," & _
                                            IIf(Val(cbo(1).Text) <= 0, 12, Val(cbo(1).Text)) & "," & _
                                            bytFontForm & "," & _
                                            lblColor.ForeColor & ")"
        End If
        
        On Error GoTo errHand
        
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        
        Call frmDefTree.AddLvwItem(lng���)
    End If
    SaveData = True
    Exit Function
errHand:
    If ErrCenter() = -1 Then Resume
    
End Function

Private Sub Form_Unload(Cancel As Integer)
    If cmdOK.Tag = "1" Then
        If MsgBox("���ĺ����ʾҳ����Ϣ���뱣�������Ч" & vbCrLf & "ȷ�ϲ�������˳���", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Cancel = True
    End If
End Sub

Private Sub txt_Change()
    cmdOK.Tag = "1"
End Sub

Private Sub txt_GotFocus()
    SelAll txt
    zlCommFun.OpenIme True
End Sub

Private Sub txt_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
        Exit Sub
    End If
    If CheckIsInclude(UCase(Chr(KeyAscii)), "'") = True Then KeyAscii = 0
End Sub

Private Sub txt_LostFocus()
    zlCommFun.OpenIme
End Sub

Private Sub txt_Validate(Cancel As Boolean)
    Cancel = Not StrIsValid(txt.Text, txt.MaxLength)
End Sub

Private Sub txtPage_Change()
    txtPage.Tag = "Changed"
    mlngPageKey = 0
    txtPage.ForeColor = &H0&
    txtPage.BackColor = &H80000005
    cmdOK.Tag = "1"
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
