VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmFontSetup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����"
   ClientHeight    =   6000
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5895
   ControlBox      =   0   'False
   Icon            =   "frmFontSetup.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   5895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin zlRichEditor.Document docSample 
      Height          =   960
      Left            =   1980
      TabIndex        =   26
      Top             =   3735
      Width           =   3705
      _ExtentX        =   6535
      _ExtentY        =   1693
      BackColor       =   0
      WYSIWYG         =   0   'False
   End
   Begin VB.CommandButton cmdBackColor 
      Caption         =   "����ɫ(&B)..."
      Height          =   350
      Left            =   3495
      TabIndex        =   25
      Top             =   4755
      Width           =   1500
   End
   Begin VB.CommandButton cmdForeColor 
      Caption         =   "ǰ��ɫ(&F)..."
      Height          =   350
      Left            =   1980
      TabIndex        =   24
      Top             =   4755
      Width           =   1515
   End
   Begin MSComDlg.CommonDialog dlgThis 
      Left            =   1680
      Top             =   5580
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.CommandButton cmdDefault 
      Caption         =   "Ĭ������(&D)..."
      Height          =   350
      Left            =   165
      TabIndex        =   23
      Top             =   5460
      Width           =   1500
   End
   Begin VB.Frame fraLine3 
      Height          =   30
      Left            =   -555
      TabIndex        =   22
      Top             =   5280
      Width           =   6855
   End
   Begin VB.ComboBox cboUnderline 
      Height          =   300
      Left            =   2820
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   3000
      Width           =   2865
   End
   Begin VB.CheckBox chkAllCaps 
      Caption         =   "ȫ����д(&A)"
      Height          =   255
      Left            =   330
      TabIndex        =   11
      Top             =   3759
      Width           =   1365
   End
   Begin VB.CheckBox chkStrikethrough 
      Caption         =   "ɾ����(&K)"
      Height          =   255
      Left            =   330
      TabIndex        =   12
      Top             =   4131
      Width           =   1365
   End
   Begin VB.CheckBox chkHidden 
      Caption         =   "��������(&H)"
      Height          =   255
      Left            =   330
      TabIndex        =   14
      Top             =   4875
      Width           =   1365
   End
   Begin VB.CheckBox chkProtected 
      Caption         =   "����(&P)"
      Height          =   255
      Left            =   330
      TabIndex        =   13
      Top             =   4503
      Width           =   1365
   End
   Begin VB.CheckBox chkSubscript 
      Caption         =   "�±�(&N)"
      Height          =   255
      Left            =   330
      TabIndex        =   10
      Top             =   3387
      Width           =   1365
   End
   Begin VB.CheckBox chkSuperscript 
      Caption         =   "�ϱ�(&U)"
      Height          =   255
      Left            =   330
      TabIndex        =   9
      Top             =   3015
      Width           =   1365
   End
   Begin VB.Frame fraLine1 
      Height          =   30
      Left            =   585
      TabIndex        =   19
      Top             =   2835
      Width           =   5145
   End
   Begin VB.ListBox lstFontSize 
      Height          =   2040
      Left            =   4665
      TabIndex        =   8
      Top             =   645
      Width           =   1005
   End
   Begin VB.TextBox txtFontSize 
      Height          =   300
      Left            =   4665
      TabIndex        =   7
      Top             =   345
      Width           =   1005
   End
   Begin VB.ListBox lstFontStyle 
      Height          =   2040
      Left            =   3270
      TabIndex        =   5
      Top             =   645
      Width           =   1155
   End
   Begin VB.TextBox txtFontStyle 
      Height          =   300
      Left            =   3270
      TabIndex        =   4
      Top             =   345
      Width           =   1155
   End
   Begin VB.ListBox lstFontName 
      Height          =   2040
      Left            =   165
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   645
      Width           =   2850
   End
   Begin VB.TextBox txtFontName 
      Height          =   300
      Left            =   165
      TabIndex        =   1
      Top             =   345
      Width           =   2850
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4585
      TabIndex        =   18
      Top             =   5460
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   3375
      TabIndex        =   17
      Top             =   5460
      Width           =   1100
   End
   Begin VB.Label lblUnderline 
      AutoSize        =   -1  'True
      Caption         =   "�»���(&L)"
      Height          =   180
      Left            =   1980
      TabIndex        =   15
      Top             =   3060
      Width           =   810
   End
   Begin VB.Label lblSample 
      AutoSize        =   -1  'True
      Caption         =   "ʾ��Ԥ��"
      Height          =   180
      Left            =   1980
      TabIndex        =   21
      Top             =   3480
      Width           =   720
   End
   Begin VB.Label lblEffects 
      AutoSize        =   -1  'True
      Caption         =   "Ч��"
      Height          =   180
      Left            =   165
      TabIndex        =   20
      Top             =   2760
      Width           =   360
   End
   Begin VB.Label lblFontSize 
      AutoSize        =   -1  'True
      Caption         =   "�ֺ�(&S)"
      Height          =   180
      Left            =   4665
      TabIndex        =   6
      Top             =   105
      Width           =   630
   End
   Begin VB.Label lblFontStyle 
      AutoSize        =   -1  'True
      Caption         =   "����(&Y)"
      Height          =   180
      Left            =   3270
      TabIndex        =   3
      Top             =   105
      Width           =   630
   End
   Begin VB.Label lblFontName 
      AutoSize        =   -1  'True
      Caption         =   "����(&T)"
      Height          =   180
      Left            =   165
      TabIndex        =   0
      Top             =   105
      Width           =   630
   End
End
Attribute VB_Name = "frmFontSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Const conFontSizes As String = "����,42;С��,36;һ��,26;Сһ,24;����,22;С��,18;����,16;С��,15;�ĺ�,14;С��,12;���,10.5;С��,9;����,7.5;С��,6.5;�ߺ�,5.5;�˺�,5;5,5;5.5,5.5;6.5,6.5;7.5,7.5;8,8;9,9;10,10;10.5,10.5;11,11;12,12;14,14;16,16;18,18;20,20;22,22;24,24;26,26;28,28;36,36;48,48;72,72"
Const conUnderlines As String = "(��),0;�l�l�l�l ����,4;�ߣߣߣ� ����,5;��.��.�� �㻮��,6;..��..�� ˫�㻮��,7;�n�n�n�n ������,8;________ ϸ��,10;�x�x�x�x ����,9"

Dim blnOK As Boolean
Dim intCount As Integer

Private Sub cboUnderline_Click()
    If Me.cboUnderline.ListIndex = -1 Then Exit Sub
    Me.docSample.ReadOnly = False
    Me.docSample.Range(0, Len(Me.docSample.Text)).Font.Underline = Me.cboUnderline.itemData(Me.cboUnderline.ListIndex)
    Me.docSample.ReadOnly = True
End Sub

Private Sub cboUnderline_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call PressKey(vbKeyTab)
End Sub

Private Sub chkAllCaps_Click()
    If Me.chkAllCaps.Value = vbCold Then Exit Sub
    Me.docSample.ReadOnly = False
    Me.docSample.Range(0, Len(Me.docSample.Text)).Font.AllCaps = Me.chkAllCaps.Value
    Me.docSample.ReadOnly = True
End Sub

Private Sub chkAllCaps_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call PressKey(vbKeyTab)
End Sub

Private Sub chkHidden_Click()
    If Me.chkHidden.Value = vbCold Then Exit Sub
    Me.docSample.ReadOnly = False
    Me.docSample.Range(0, Len(Me.docSample.Text)).Font.Hidden = Me.chkHidden.Value
    Me.docSample.ReadOnly = True
End Sub

Private Sub chkHidden_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call PressKey(vbKeyTab)
End Sub

Private Sub chkProtected_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call PressKey(vbKeyTab)
End Sub

Private Sub chkStrikethrough_Click()
    If Me.chkStrikethrough.Value = vbCold Then Exit Sub
    Me.docSample.ReadOnly = False
    Me.docSample.Range(0, Len(Me.docSample.Text)).Font.Strikethrough = Me.chkStrikethrough.Value
    Me.docSample.ReadOnly = True
End Sub

Private Sub chkStrikethrough_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call PressKey(vbKeyTab)
End Sub

Private Sub chkSubscript_Click()
    If Me.chkSubscript.Value = vbCold Then Exit Sub
    If Me.chkSubscript.Value = vbChecked Then Me.chkSuperscript.Value = vbUnchecked
    Me.docSample.ReadOnly = False
    Me.docSample.Range(0, Len(Me.docSample.Text)).Font.Subscript = Me.chkSubscript.Value
    Me.docSample.ReadOnly = True
End Sub

Private Sub chkSubscript_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call PressKey(vbKeyTab)
End Sub

Private Sub chkSuperscript_Click()
    If Me.chkSuperscript.Value = vbCold Then Exit Sub
    If Me.chkSuperscript.Value = vbChecked Then Me.chkSubscript.Value = vbUnchecked
    Me.docSample.ReadOnly = False
    Me.docSample.Range(0, Len(Me.docSample.Text)).Font.Superscript = Me.chkSuperscript.Value
    Me.docSample.ReadOnly = True
End Sub

Private Sub chkSuperscript_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call PressKey(vbKeyTab)
End Sub

Private Sub cmdBackColor_Click()
    With Me.dlgThis
        If Me.docSample.Range(0, Len(Me.docSample.Text)).Font.BackColor <> tomAutoColor And Me.cmdBackColor.Tag <> CStr(tomAutoColor) Then
            .Color = Me.docSample.Range(0, Len(Me.docSample.Text)).Font.BackColor
        End If
        .DialogTitle = "����ɫ"
        Err = 0: On Error Resume Next
        .ShowColor
        If Err.Number <> 0 Then Exit Sub
        Me.cmdBackColor.Tag = ""
        Me.docSample.ReadOnly = False
        Me.docSample.Range(0, Len(Me.docSample.Text)).Font.BackColor = .Color
        Me.docSample.ReadOnly = True
    End With
End Sub

Private Sub cmdCancel_Click()
    Me.Hide
End Sub

Private Sub cmdDefault_Click()
    Dim strMsgInfo As String
    strMsgInfo = "�Ƿ�Ĭ���������Ϊ��" & Me.txtFontName.Text & "," & Me.txtFontStyle.Text & "," & Me.txtFontSize.Text & "����" & _
        vbCrLf & "�˸��Ľ�Ӱ���µ��ĵ���"
    If MsgBox(strMsgInfo, vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then Exit Sub
    
    With Me.docSample
        .ReadOnly = False
        SaveSetting UCase(App.ProductName), "FONT", UCase("Name"), .Selection.Font.Name
        SaveSetting UCase(App.ProductName), "FONT", UCase("Italic"), .Selection.Font.Italic
        SaveSetting UCase(App.ProductName), "FONT", UCase("Bold"), .Selection.Font.Bold
        SaveSetting UCase(App.ProductName), "FONT", UCase("Size"), .Selection.Font.SIZE
        .ReadOnly = True
    End With
End Sub

Private Sub cmdForeColor_Click()
    With Me.dlgThis
        If Me.docSample.Range(0, Len(Me.docSample.Text)).Font.ForeColor <> tomAutoColor And Me.cmdForeColor.Tag <> CStr(tomAutoColor) Then
            .Color = Me.docSample.Range(0, Len(Me.docSample.Text)).Font.ForeColor
        End If
        .DialogTitle = "ǰ��ɫ"
        Err = 0: On Error Resume Next
        .ShowColor
        If Err.Number <> 0 Then Exit Sub
        Me.cmdForeColor.Tag = ""
        Me.docSample.ReadOnly = False
        Me.docSample.Range(0, Len(Me.docSample.Text)).Font.ForeColor = .Color
        Me.docSample.ReadOnly = True
    End With
End Sub

Private Sub cmdOK_Click()
    blnOK = True: Me.Hide
End Sub

Private Sub Form_Activate()
    '�ʵ������ؼ�λ��
    If Me.cboUnderline.Visible = False Then
        Me.lblSample.Top = Me.lblUnderline.Top
        Me.docSample.Top = Me.cboUnderline.Top + Me.cboUnderline.Height
        Me.docSample.Height = Me.cmdForeColor.Top - Me.docSample.Top
    End If
    If Me.cmdForeColor.Visible = False Then Me.cmdBackColor.Left = Me.cmdForeColor.Left
    If Me.cmdForeColor.Visible = False And Me.cmdBackColor.Visible = False Then
        Me.docSample.Height = Me.cmdForeColor.Top + Me.cmdForeColor.Height - Me.docSample.Top
    End If
    Me.txtFontName.SetFocus
End Sub

Private Sub lstFontName_Click()
    Err = 0: On Error Resume Next
    If Me.ActiveControl.Name <> Me.txtFontName.Name Then
        Me.txtFontName.Text = Me.lstFontName.Text
    End If
    Me.docSample.ReadOnly = False
    Me.docSample.Range(0, Len(Me.docSample.Text)).Font.Name = Me.lstFontName.Text
    Me.docSample.ReadOnly = True
End Sub

Private Sub lstFontName_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call PressKey(vbKeyTab)
End Sub

Private Sub lstFontSize_Click()
    Me.txtFontSize.Tag = Me.lstFontSize.itemData(Me.lstFontSize.ListIndex) / 10
    Me.txtFontSize.Text = Me.lstFontSize.Text
    Me.docSample.ReadOnly = False
    Me.docSample.Range(0, Len(Me.docSample.Text)).Font.SIZE = Val(Me.txtFontSize.Tag)
    Me.docSample.ReadOnly = True
End Sub

Private Sub lstFontSize_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call PressKey(vbKeyTab)
End Sub

Private Sub lstFontStyle_Click()
    Err = 0: On Error Resume Next
    If Me.ActiveControl.Name <> Me.txtFontStyle.Name Then
        Me.txtFontStyle.Text = Me.lstFontStyle.Text
    End If
    Me.docSample.ReadOnly = False
    Select Case Me.lstFontStyle.Text
    Case "����"
        Me.docSample.Range(0, Len(Me.docSample.Text)).Font.Italic = False: Me.docSample.Range(0, Len(Me.docSample.Text)).Font.Bold = False
    Case "�Ӵ�"
        Me.docSample.Range(0, Len(Me.docSample.Text)).Font.Italic = False: Me.docSample.Range(0, Len(Me.docSample.Text)).Font.Bold = True
    Case "��б"
        Me.docSample.Range(0, Len(Me.docSample.Text)).Font.Italic = True: Me.docSample.Range(0, Len(Me.docSample.Text)).Font.Bold = False
    Case "�Ӵ� ��б"
        Me.docSample.Range(0, Len(Me.docSample.Text)).Font.Italic = True: Me.docSample.Range(0, Len(Me.docSample.Text)).Font.Bold = True
    End Select
    Me.docSample.ReadOnly = True
End Sub

Private Sub lstFontStyle_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call PressKey(vbKeyTab)
End Sub

Private Sub txtFontName_GotFocus()
    Me.txtFontName.SelStart = 0: Me.txtFontName.SelLength = Len(Me.txtFontName.Text) + 1
End Sub

Private Sub txtFontName_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call PressKey(vbKeyTab): Exit Sub
    If LenB(Me.txtFontName.Text) = 0 And KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Then KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtFontName_KeyUp(KeyCode As Integer, Shift As Integer)
    If InStr(1, Me.lstFontName.Text, Trim(Me.txtFontName.Text)) = 1 Then Exit Sub
    With Me.lstFontName
        For intCount = 0 To .ListCount - 1
            If InStr(1, .List(intCount), Trim(Me.txtFontName.Text)) = 1 Then
                .ListIndex = intCount: Exit Sub
            End If
        Next
    End With
End Sub

Private Sub txtFontName_LostFocus()
    Me.txtFontName.Text = Me.lstFontName.Text
End Sub

Private Sub txtFontSize_Change()
    If Val(Me.txtFontSize.Text) <> 0 Then
        Me.txtFontSize.Tag = Val(Me.txtFontSize.Text)
        Me.docSample.ReadOnly = False
        Me.docSample.Range(0, Len(Me.docSample.Text)).Font.SIZE = Val(Me.txtFontSize.Tag)
        Me.docSample.ReadOnly = True
    End If
End Sub

Private Sub txtFontSize_GotFocus()
    Me.txtFontSize.SelStart = 0: Me.txtFontSize.SelLength = Len(Me.txtFontSize.Text) + 1
    Call OpenIme(False)
End Sub

Private Sub txtFontSize_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call PressKey(vbKeyTab)
End Sub

Private Sub txtFontSize_KeyUp(KeyCode As Integer, Shift As Integer)
    With Me.lstFontSize
        For intCount = 0 To .ListCount - 1
            If .List(intCount) = Trim(Me.txtFontSize.Text) Then
                .ListIndex = intCount:  Exit Sub
            End If
        Next
    End With
End Sub

Private Sub txtFontStyle_GotFocus()
    Me.txtFontStyle.SelStart = 0: Me.txtFontStyle.SelLength = Len(Me.txtFontStyle.Text) + 1
    Call OpenIme(True)
End Sub

Private Sub txtFontStyle_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call PressKey(vbKeyTab)
End Sub

Private Sub txtFontStyle_KeyUp(KeyCode As Integer, Shift As Integer)
    If InStr(1, Me.lstFontStyle.Text, Trim(Me.txtFontStyle.Text)) = 1 Then Exit Sub
    With Me.lstFontStyle
        For intCount = 0 To .ListCount - 1
            If InStr(1, .List(intCount), Trim(Me.txtFontStyle.Text)) = 1 Then
                .ListIndex = intCount:  Exit Sub
            End If
        Next
    End With
End Sub

Private Sub txtFontStyle_LostFocus()
    Me.txtFontStyle.Text = Me.lstFontStyle.Text
End Sub

Public Function ShowMe(curTOM As cTextDocument, Optional intFlags As Integer, Optional strSample As String) As Boolean
    '���ܣ���ʾ������Ի���
    '������
    '   curTOM,��Ҫ����������ĵ�����
    '   intFlags,�Ƿ��ֹ��صĸ���Ч��ѡ�
    '       intFlags and (2^0) <> 0,��ֹ����ɾ��������
    '       intFlags and (2^1) <> 0,��ֹ���ı�������
    '       intFlags and (2^2) <> 0,��ֹ������������
    '       intFlags and (2^3) <> 0,��ֹ�����»�������
    '       intFlags and (2^4) <> 0,��ֹ����ǰ��ɫ����
    '       intFlags and (2^5) <> 0,��ֹ���ı���ɫ����
    '   strSample,��ʾ��������
    
    Dim objDevice As Object
    Dim aryRows() As String, aryItems() As String
    
    'ʾ����ʾ���ã�
    If Trim(strSample) = "" Or Trim(strSample) = "��" Then strSample = "���� Font"
    With Me.docSample
        .Text = strSample
        .Range(0, Len(.Text)).Para.Alignment = cprHACenter
        .ReadOnly = True
        .SelLength = 0
    End With
    
    '�����б��ʼ��������ڴ�ӡ�������Դ�ӡ������Ϊ�б���������Ļ����Ϊ�б�
    If Not ExistsPrinter Then
        Set objDevice = Screen
    Else
        Set objDevice = Printer
    End If
    With Me.lstFontName
        For intCount = 0 To objDevice.FontCount - 1
            .AddItem objDevice.Fonts(intCount)
            If curTOM.TextDocument.Selection.Font.Name = objDevice.Fonts(intCount) Then .ListIndex = .NewIndex
        Next
        If .ListIndex = -1 And .ListCount > 0 Then .ListIndex = 0
        .TopIndex = .ListIndex
    End With
    
    '�����б��ʼ
    With Me.lstFontStyle
        .AddItem "����"
        .AddItem "�Ӵ�"
'        .AddItem "��б"
'        .AddItem "�Ӵ� ��б"
    End With
    If curTOM.TextDocument.Selection.Font.Italic <> tomUndefined And curTOM.TextDocument.Selection.Font.Bold <> tomUndefined Then
        Me.lstFontStyle.ListIndex = Abs(curTOM.TextDocument.Selection.Font.Italic * 2 + curTOM.TextDocument.Selection.Font.Bold)
    ElseIf curTOM.TextDocument.Selection.Font.Italic <> tomUndefined And curTOM.TextDocument.Selection.Font.Bold = tomUndefined Then
        Me.lstFontStyle.ListIndex = Abs(curTOM.TextDocument.Selection.Font.Italic * 2)
    ElseIf curTOM.TextDocument.Selection.Font.Italic = tomUndefined And curTOM.TextDocument.Selection.Font.Bold <> tomUndefined Then
        Me.lstFontStyle.ListIndex = Abs(curTOM.TextDocument.Selection.Font.Bold)
    Else
        Me.lstFontStyle.ListIndex = 0
    End If
    
    '�ֺ��б��ʼ
    aryRows = Split(conFontSizes, ";")
    With Me.lstFontSize
        For intCount = 0 To UBound(aryRows)
            aryItems = Split(aryRows(intCount), ",")
            .AddItem aryItems(0)
            .itemData(.NewIndex) = aryItems(1) * 10
            If curTOM.TextDocument.Selection.Font.SIZE = aryItems(1) And .ListIndex = -1 Then .ListIndex = .NewIndex
        Next
        If .ListIndex = -1 Then
            If curTOM.TextDocument.Selection.Font.SIZE <> tomUndefined Then Me.txtFontSize.Text = curTOM.TextDocument.Selection.Font.SIZE
        Else
            .TopIndex = .ListIndex
        End If
    End With
    
    If curTOM.TextDocument.Selection.Font.Subscript = tomUndefined Then
        Me.chkSubscript.Value = vbCold: Me.chkSuperscript.Value = vbCold
    Else
        Me.chkSubscript.Value = Abs(curTOM.TextDocument.Selection.Font.Subscript): Call chkSubscript_Click
        If curTOM.TextDocument.Selection.Font.Subscript = False Then
            Me.chkSuperscript.Value = Abs(curTOM.TextDocument.Selection.Font.Superscript): Call chkSuperscript_Click
        End If
    End If
    If curTOM.TextDocument.Selection.Font.AllCaps = tomUndefined Then
        Me.chkAllCaps.Value = vbCold
    Else
        Me.chkAllCaps.Value = Abs(curTOM.TextDocument.Selection.Font.AllCaps): Call chkAllCaps_Click
    End If

    If (intFlags And (2 ^ 0)) <> 0 Then '��ֹɾ����
        Me.chkStrikethrough.Visible = False
    Else
        If curTOM.TextDocument.Selection.Font.Strikethrough = tomUndefined Then
            Me.chkStrikethrough.Value = vbCold
        Else
            Me.chkStrikethrough.Value = Abs(curTOM.TextDocument.Selection.Font.Strikethrough): Call chkStrikethrough_Click
        End If
    End If
    If (intFlags And (2 ^ 1)) <> 0 Then '��ֹ����
        Me.chkProtected.Visible = False
    Else
        If curTOM.TextDocument.Selection.Font.Protected = tomUndefined Then
            Me.chkProtected.Value = vbCold
        Else
            Me.chkProtected.Value = Abs(curTOM.TextDocument.Selection.Font.Protected)
        End If
    End If
    If (intFlags And (2 ^ 2)) <> 0 Then  '��ֹ����
        Me.chkHidden.Visible = False
    Else
        If curTOM.TextDocument.Selection.Font.Hidden = tomUndefined Then
            Me.chkHidden.Value = vbCold
        Else
            Me.chkHidden.Value = Abs(curTOM.TextDocument.Selection.Font.Hidden): Call chkHidden_Click
        End If
    End If
    
    If (intFlags And (2 ^ 3)) <> 0 Then '��ֹ�»���
        Me.lblUnderline.Visible = False: Me.cboUnderline.Visible = False
    Else
        '�»����б��ʼ
        aryRows = Split(conUnderlines, ";")
        With Me.cboUnderline
            For intCount = 0 To UBound(aryRows)
                aryItems = Split(aryRows(intCount), ",")
                .AddItem aryItems(0)
                .itemData(.NewIndex) = aryItems(1)
                If curTOM.TextDocument.Selection.Font.Underline = aryItems(1) Then .ListIndex = .NewIndex
            Next
'            If .ListIndex = -1 Then .ListIndex = 0
        End With
    End If
    
    Me.docSample.ReadOnly = False
    If (intFlags And (2 ^ 4)) <> 0 Then  '��ֹǰ��ɫ
        Me.cmdForeColor.Visible = False
    Else
        If curTOM.TextDocument.Selection.Font.ForeColor = tomAutoColor Then
            Me.cmdForeColor.Tag = CStr(tomAutoColor)
        Else
            Me.docSample.Range(0, Len(Me.docSample.Text)).Font.ForeColor = curTOM.TextDocument.Selection.Font.ForeColor
        End If
    End If
    If (intFlags And (2 ^ 5)) <> 0 Then  '��ֹ����ɫ?
        Me.cmdBackColor.Visible = False
    Else
        If curTOM.TextDocument.Selection.Font.BackColor = tomAutoColor Then
            Me.cmdBackColor.Tag = CStr(tomAutoColor)
        Else
            Me.docSample.Range(0, Len(Me.docSample.Text)).Font.BackColor = curTOM.TextDocument.Selection.Font.BackColor
        End If
    End If
    Me.docSample.ReadOnly = True
    
    blnOK = False
    Me.Show 1
    If blnOK = False Then Unload Me: ShowMe = False: Exit Function
    
    With Me.docSample
        .ReadOnly = False
        curTOM.TextDocument.Selection.Font.Name = .Range(0, 1).Font.Name
        curTOM.TextDocument.Selection.Font.Italic = .Range(0, 1).Font.Italic
        curTOM.TextDocument.Selection.Font.Bold = .Range(0, 1).Font.Bold
        If curTOM.TextDocument.Selection.Font.SIZE = tomUndefined And Val(Me.txtFontSize.Text) = 0 And Me.lstFontSize.ListIndex = -1 Then
            'û�������ֺ�
        Else
            curTOM.TextDocument.Selection.Font.SIZE = .Range(0, 1).Font.SIZE
        End If
        If Me.chkSubscript.Value <> vbCold And Me.chkSuperscript.Value <> vbCold Then
            curTOM.TextDocument.Selection.Font.Subscript = .Range(0, 1).Font.Subscript
            If curTOM.TextDocument.Selection.Font.Subscript = False Then
                curTOM.TextDocument.Selection.Font.Superscript = .Range(0, 1).Font.Superscript
            End If
        End If
        If Me.chkAllCaps.Value <> vbCold Then curTOM.TextDocument.Selection.Font.AllCaps = .Range(0, 1).Font.AllCaps
        If (intFlags And (2 ^ 0)) = 0 And Me.chkStrikethrough <> vbCold Then curTOM.TextDocument.Selection.Font.Strikethrough = .Range(0, 1).Font.Strikethrough 'ɾ����
        If (intFlags And (2 ^ 1)) = 0 And Me.chkProtected.Value <> vbCold Then curTOM.TextDocument.Selection.Font.Protected = Me.chkProtected.Value '����
        If (intFlags And (2 ^ 2)) = 0 And Me.chkHidden.Value <> vbCold Then curTOM.TextDocument.Selection.Font.Hidden = .Range(0, 1).Font.Hidden '����
        If (intFlags And (2 ^ 3)) = 0 And Me.cboUnderline.ListIndex <> -1 Then curTOM.TextDocument.Selection.Font.Underline = .Range(0, 1).Font.Underline '�»���
        If (intFlags And (2 ^ 4)) = 0 And Me.cmdForeColor.Tag <> CStr(tomAutoColor) Then curTOM.TextDocument.Selection.Font.ForeColor = .Range(0, 1).Font.ForeColor 'ǰ��ɫ
        If (intFlags And (2 ^ 5)) = 0 And Me.cmdBackColor.Tag <> CStr(tomAutoColor) Then curTOM.TextDocument.Selection.Font.BackColor = .Range(0, 1).Font.BackColor '����ɫ
    End With
    
    ShowMe = True: Unload Me
End Function
