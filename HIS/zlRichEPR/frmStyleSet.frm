VERSION 5.00
Object = "*\A..\zlRichEditor\zlRichEdit.vbp"
Begin VB.Form frmStyleSet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "������ʽ����"
   ClientHeight    =   5625
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5520
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   5520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.Frame frmLine 
      Height          =   30
      Index           =   1
      Left            =   2730
      TabIndex        =   12
      Top             =   4950
      Width           =   2655
   End
   Begin VB.Frame frmLine 
      Height          =   30
      Index           =   0
      Left            =   165
      TabIndex        =   11
      Top             =   555
      Width           =   5220
   End
   Begin VB.TextBox txt��� 
      Height          =   300
      Left            =   855
      TabIndex        =   1
      Top             =   135
      Width           =   795
   End
   Begin VB.TextBox txt���� 
      Height          =   300
      Left            =   2655
      TabIndex        =   3
      Top             =   135
      Width           =   2685
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4260
      TabIndex        =   10
      Top             =   5115
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   3165
      TabIndex        =   9
      Top             =   5115
      Width           =   1100
   End
   Begin zlRichEditor.Editor edt���� 
      Height          =   1890
      Left            =   150
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   930
      Width           =   5205
      _ExtentX        =   9181
      _ExtentY        =   3334
      Border          =   -1  'True
      PaperWidth      =   1907
      WithViewButtonas=   0   'False
      PaperKind       =   256
      ShowRuler       =   0   'False
   End
   Begin VB.CommandButton cmdSet 
      Caption         =   "����(&F)��"
      Height          =   350
      Index           =   1
      Left            =   1260
      TabIndex        =   8
      Top             =   5115
      Width           =   1100
   End
   Begin VB.CommandButton cmdSet 
      Caption         =   "����(&P)��"
      Height          =   350
      Index           =   0
      Left            =   150
      TabIndex        =   7
      Top             =   5115
      Width           =   1100
   End
   Begin VB.Label lbl���� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "��ʽ����:"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   150
      TabIndex        =   4
      Top             =   705
      Width           =   810
   End
   Begin VB.Label lbl��� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "���(&D)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   150
      TabIndex        =   0
      Top             =   195
      Width           =   630
   End
   Begin VB.Label lbl���� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "����(&N)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   1935
      TabIndex        =   2
      Top             =   195
      Width           =   630
   End
   Begin VB.Label lbl���� 
      Caption         =   "###"
      Height          =   1935
      Left            =   150
      TabIndex        =   6
      Top             =   2895
      Width           =   5205
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmStyleSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngOldCode As Long, mlngNewCode As Long
Public Function ShowMe(ByVal frmParent As Object, ByVal blnAdd As Boolean, Optional ByVal lngCode As Long) As Long
    '---------------------------------------------------
    '���ܣ��ϼ�������ñ�����ģ����ݲ���������ʾ����
    '������ frmParent-������
    '       blnAdd-�Ƿ�����
    '       lngCode-�޸�ʱ��Ҫָ������ʽ�ı��
    '���أ�ȷ�������������޸ĵı�ţ�ȡ������0
    '---------------------------------------------------
Dim rsTemp As New ADODB.Recordset
Dim aryFormat() As String

    If blnAdd Then
        Me.Tag = "����": mlngOldCode = 0
    Else
        Me.Tag = "�޸�": mlngOldCode = lngCode
    End If
    Me.edt����.Text = "��ʽ���÷���"
    Me.edt����.PaperWidth = 3000
    Me.edt����.ResetWYSIWYG
    
    'ԭ��ʽ��Ϣ��ȡ
    Err = 0: On Error GoTo errHand
    gstrSQL = "Select ���, ����, ������ʽ, ������ʽ, ϵͳ From ����������ʽ Where ���=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngOldCode)
    With rsTemp
        Me.txt���.MaxLength = 3: Me.txt����.MaxLength = .Fields("����").DefinedSize
        If .RecordCount > 0 Then
            Me.txt���.Text = Format(!���, String(Me.txt���.MaxLength, "0")): Me.txt���� = "" & !����
            If Val("" & !ϵͳ) = 1 Then Me.txt���.Enabled = False: Me.txt����.Enabled = False
            
            If "" & !������ʽ <> "" Then
                aryFormat = Split("" & !������ʽ, ";")
                Me.edt����.ForceEdit = True
                With Me.edt����.Range(0, Len(Me.edt����.Text)).Font
                    If Trim(aryFormat(0)) <> "" Then .Name = aryFormat(0)
                    If Val(aryFormat(1)) > 0 Then .Size = Val(aryFormat(1))
                    .Bold = IIf(Mid(aryFormat(2), 1, 1) = 1, True, False)
                    .Italic = IIf(Mid(aryFormat(2), 2, 1) = 1, True, False)
                    .Superscript = IIf(Mid(aryFormat(2), 7, 1) = 1, True, False)
                    .Subscript = IIf(Mid(aryFormat(2), 8, 1) = 1, True, False)
                    .ForeColor = Val(aryFormat(5))
                End With
                Me.edt����.ForceEdit = False
            End If

            If "" & !������ʽ <> "" Then
                aryFormat = Split("" & !������ʽ, ";")
                Me.edt����.ForceEdit = True
                With Me.edt����.Range(0, Len(Me.edt����.Text)).Para
                    If Mid(aryFormat(0), 2, 1) < 9 Then .ListAlignment = Mid(aryFormat(0), 2, 1)                       '����ȡֵΪ��0��1��2
                    If Val(aryFormat(1)) <> -9999999 Then .Style = Val(aryFormat(1))        '����ȡֵ�� -1 ~ -10����Ҫ�����о����ã������о�ʧЧ
                    
                    If Val(aryFormat(2)) <> -9999999 Then
                        .ListType = Val(aryFormat(2))     '����ȡֵ��0 �� 6��65536��131072��196608
                        .ListStart = Val(aryFormat(3))
                    End If
                    If Val(aryFormat(4)) <> tomUndefined Then .FirstLineIndent = Val(aryFormat(4)) '��������һ��������
                    If Val(aryFormat(5)) <> tomUndefined Then .LeftIndent = Val(aryFormat(5))
                    If Val(aryFormat(6)) <> tomUndefined Then .RightIndent = Val(aryFormat(6))
                    If Val(aryFormat(8)) <> tomUndefined Then .ListTab = Val(aryFormat(8))
                    If Val(aryFormat(9)) <> tomUndefined Then .SpaceBefore = Val(aryFormat(9))
                    If Val(aryFormat(10)) <> tomUndefined Then .SpaceAfter = Val(aryFormat(10))
                    
                    If Mid(aryFormat(0), 3, 1) < 9 And Val(aryFormat(7)) <> tomUndefined Then .SetLineSpacing Mid(aryFormat(0), 3, 1), Val(aryFormat(7))
                    If Mid(aryFormat(0), 1, 1) < 9 Then .Alignment = Mid(aryFormat(0), 1, 1)                           '����ȡֵΪ��0��1��2
                End With
                Me.edt����.ForceEdit = False
            End If
            
        End If
    End With
    
    If Me.Tag = "����" Then
        gstrSQL = "Select nvl(max(���),'" & String(Me.txt���.MaxLength, "0") & "') as ��� From ����������ʽ"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
        Me.txt���.Text = Format(Val(rsTemp!���) + 1, String(Me.txt���.MaxLength, "0"))
    End If
    
    Me.lbl����.Caption = zlStyleDesc
    '��ʾ����
    Me.Show vbModal, frmParent
    ShowMe = mlngNewCode
    Unload Me
    Exit Function

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    ShowMe = 0
End Function

Private Sub cmdCancel_Click()
    mlngNewCode = 0: Me.Hide
End Sub

Private Sub cmdOK_Click()
    '�������
    If Trim(Me.txt���.Text) = "" Then MsgBox "�������ţ�", vbInformation, gstrSysName: Me.txt���.SetFocus: Exit Sub
    If Len(Me.txt���.Text) < Me.txt���.MaxLength Then MsgBox "��ų��Ȳ��㣡", vbInformation, gstrSysName: Me.txt���.SetFocus: Exit Sub
    If Trim(Me.txt����.Text) = "" Then MsgBox "���������ƣ�", vbInformation, gstrSysName: Me.txt����.SetFocus: Exit Sub
    If LenB(StrConv(Trim(Me.txt����.Text), vbFromUnicode)) > Me.txt����.MaxLength Then
        MsgBox "���Ƴ��������" & Me.txt����.MaxLength & "���ַ���ȳ��ĺ��֣���", vbInformation, gstrSysName: Me.txt����.SetFocus: Exit Sub
    End If
    
    '��ʽ��֯
    Dim strFont As String, strPara As String
    With Me.edt����.Range(0, Len(Me.edt����.Text)).Font
        strFont = .Name
        strFont = strFont & ";" & .Size
        strFont = strFont & ";" & Abs(CInt(.Bold)) & Abs(CInt(.Italic)) & Abs(CInt(.Hidden)) & Abs(CInt(.Protected)) _
                & Abs(CInt(.Link)) & Abs(CInt(.Strikethrough)) & Abs(CInt(.Superscript)) & Abs(CInt(.Subscript))
        strFont = strFont & ";" & .Underline
        strFont = strFont & ";" & .BackColor
        strFont = strFont & ";" & .ForeColor
    End With
    With Me.edt����.Range(0, Len(Me.edt����.Text)).Para
        strPara = .Alignment & .ListAlignment & .LineSpacingRule
        strPara = strPara & ";" & .Style
        strPara = strPara & ";" & .ListType
        strPara = strPara & ";" & .ListStart
        strPara = strPara & ";" & .FirstLineIndent
        strPara = strPara & ";" & .LeftIndent
        strPara = strPara & ";" & .RightIndent
        strPara = strPara & ";" & .LineSpacing
        strPara = strPara & ";" & .ListTab
        strPara = strPara & ";" & .SpaceBefore
        strPara = strPara & ";" & .SpaceAfter
    End With
    
    '�������
    If Me.Tag = "����" Then
        gstrSQL = "Zl_����������ʽ_Insert(" & Trim(Me.txt���.Text) & ",'" & Trim(Me.txt����.Text) & "','" & strPara & "','" & strFont & "')"
    Else
        gstrSQL = "Zl_����������ʽ_Update(" & mlngOldCode & "," & Trim(Me.txt���.Text) & ",'" & Trim(Me.txt����.Text) & "','" & strPara & "','" & strFont & "')"
    End If
    Err = 0: On Error GoTo errHand
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    mlngNewCode = Val(Trim(Me.txt���.Text)): Me.Hide
    Exit Sub

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdSet_Click(Index As Integer)
    Dim blnSet As Boolean
    With Me.edt����
        .SelStart = 0: .SelLength = Len(.Text)
        .ForceEdit = True
        If Index = 0 Then
            blnSet = .ShowParaDlg(False)
        Else
            blnSet = .ShowFontDlg(2 ^ 5 + 2 ^ 4 + 2 ^ 3 + 2 ^ 2 + 2 ^ 1 + 2 ^ 0)
        End If
        .ForceEdit = False
        .SelStart = 0
    End With
    Me.lbl����.Caption = zlStyleDesc
End Sub

Private Sub Form_Activate()
    If Me.txt���.Visible And Me.txt���.Enabled Then Me.txt���.SetFocus
End Sub

Private Sub txt���_Change()
    txt��� = Val(txt���)
End Sub

Private Sub txt���_GotFocus()
    Me.txt���.SelStart = 0: Me.txt���.SelLength = 100
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txt���_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt����_Change()
    ValidControlText txt����
End Sub

Private Sub txt����_GotFocus()
    Me.txt����.SelStart = 0: Me.txt����.SelLength = 4000
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(" ~!@#$%^&*()[]{}_+|=-`;'"":/.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Function zlStyleDesc() As String
    '���ݵ�ǰ��ʽ��д��ʽ˵���ı�
    Dim strStyle As String
    
    With Me.edt����.Range(0, Len(Me.edt����.Text)).Font
        strStyle = "������ʽ:" & .Name
        strStyle = strStyle & ", �ߴ�:" & .Size
        strStyle = strStyle & IIf(.Bold, ", ����", "")
        strStyle = strStyle & IIf(.Italic, ", б��", "")
        strStyle = strStyle & IIf(.Superscript, ", �ϱ�", "")
        strStyle = strStyle & IIf(.Subscript, ", �±�", "")
        If .ForeColor = tomAutoColor Then
            strStyle = strStyle & ", ǰ��ɫ:�Զ�"
        Else
            strStyle = strStyle & ", ǰ��ɫ:" & .ForeColor
        End If
    End With

    strStyle = strStyle & vbCrLf & vbCrLf & "������ʽ:"
    With Me.edt����.Range(0, Len(Me.edt����.Text)).Para
        Select Case .Alignment   '����ȡֵΪ��0��1��2
        Case 0: strStyle = strStyle & "�����"
        Case 1: strStyle = strStyle & "����"
        Case 2: strStyle = strStyle & "�Ҷ���"
        End Select
        If .Style = cprPSNormal Then
            strStyle = strStyle & ", ��ٲ��: ����"
        Else
            strStyle = strStyle & ", ��ٲ��: ����" & Abs(.Style) - 1
        End If
        strStyle = strStyle & ", ��������:" & .FirstLineIndent
        strStyle = strStyle & ", �������:" & .LeftIndent
        strStyle = strStyle & ", �Ҷ�����:" & .RightIndent
        
        Select Case .LineSpacingRule
        Case cprLSSignle:   strStyle = strStyle & ", �����о�"
        Case cprLS1pt5:     strStyle = strStyle & ", 1.5���о�"
        Case cprLSDouble:   strStyle = strStyle & ", �����о�"
        Case cprLSAtLeast:   strStyle = strStyle & ", ��С�о�:" & .LineSpacing
        Case cprLSExactly:   strStyle = strStyle & ", ��ȷ�о�:" & .LineSpacing
        Case cprLSMultiple:   strStyle = strStyle & ", �౶�о�:" & .LineSpacing
        End Select
        strStyle = strStyle & ", ��ǰ���:" & .SpaceBefore
        strStyle = strStyle & ", �κ���:" & .SpaceAfter
    End With
    zlStyleDesc = strStyle
End Function
