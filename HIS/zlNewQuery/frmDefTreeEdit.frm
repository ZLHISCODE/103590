VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDefTreeEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ŀ¼�༭"
   ClientHeight    =   3000
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5055
   Icon            =   "frmDefTreeEdit.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   5055
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Height          =   3000
      Left            =   45
      TabIndex        =   17
      Top             =   -45
      Width           =   3570
      Begin VB.PictureBox Picture1 
         Height          =   930
         Left            =   780
         ScaleHeight     =   870
         ScaleWidth      =   2595
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   1995
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
            TabIndex        =   19
            Top             =   30
            Width           =   840
         End
      End
      Begin VB.CommandButton cmdOpen 
         Caption         =   "��"
         Height          =   240
         Left            =   1620
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   1320
         Width           =   255
      End
      Begin VB.TextBox txt 
         Height          =   300
         Left            =   780
         MaxLength       =   20
         TabIndex        =   1
         Top             =   195
         Width           =   2655
      End
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   0
         Left            =   780
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   555
         Width           =   2655
      End
      Begin VB.ComboBox cboIcon 
         Height          =   300
         Left            =   780
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   1650
         Width           =   1875
      End
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   2
         Left            =   780
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   930
         Width           =   1125
      End
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   1
         Left            =   2640
         TabIndex        =   7
         Text            =   "cbo"
         Top             =   915
         Width           =   795
      End
      Begin zl9NewQuery.ctlPicture UsrPicture 
         Height          =   435
         Index           =   1
         Left            =   2715
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   1515
         Width           =   570
         _ExtentX        =   1005
         _ExtentY        =   767
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
         Left            =   765
         TabIndex        =   9
         Top             =   1290
         Width           =   1140
      End
      Begin VB.Label Label2 
         Caption         =   "����(&1)"
         Height          =   225
         Left            =   120
         TabIndex        =   0
         Top             =   270
         Width           =   1245
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "ͼ��(&6)"
         Height          =   180
         Left            =   120
         TabIndex        =   11
         Top             =   1725
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "��ɫ(&5)"
         Height          =   180
         Left            =   120
         TabIndex        =   8
         Top             =   1365
         Width           =   630
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "ʾ��(&7)"
         Height          =   180
         Left            =   120
         TabIndex        =   13
         Top             =   2010
         Width           =   630
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "��С(&4)"
         Height          =   180
         Left            =   1980
         TabIndex        =   6
         Top             =   975
         Width           =   630
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "����(&3)"
         Height          =   180
         Left            =   120
         TabIndex        =   4
         Top             =   1005
         Width           =   630
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "����(&2)"
         Height          =   180
         Left            =   120
         TabIndex        =   2
         Top             =   630
         Width           =   630
      End
   End
   Begin MSComDlg.CommonDialog dlg 
      Left            =   3960
      Top             =   1845
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   3930
      Top             =   1050
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
            Picture         =   "frmDefTreeEdit.frx":000C
            Key             =   "default"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   3780
      TabIndex        =   15
      Top             =   105
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3780
      TabIndex        =   16
      Top             =   525
      Width           =   1100
   End
End
Attribute VB_Name = "frmDefTreeEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnFirst As Boolean
Private blnOK As Boolean
Private mKey As Long

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

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    
    On Error GoTo errHand
    
    '��ʼ������
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
        gstrSQL = "Select ���,�����,����,ҳ��,ҳ��ͼ��,����,��С,����,��ɫ From ��ѯҳ������ where ���=[1]"
        Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mKey)
        If gRs.BOF = False Then
            txt.Text = IIf(IsNull(gRs!����), "", gRs!����)
            cboIcon.ListIndex = FindCboIndex(cboIcon, IIf(IsNull(gRs!ҳ��ͼ��), 0, gRs!ҳ��ͼ��))
            If cboIcon.ListIndex = -1 Then cboIcon.ListIndex = 0
            
            On Error Resume Next
            
            cbo(0).Text = IIf(IsNull(gRs!����), "����", gRs!����)
            cbo(1).Text = IIf(IsNull(gRs!��С), "12", gRs!��С)
            lblColor.ForeColor = IIf(IsNull(gRs!��ɫ), &HFF0000, gRs!��ɫ)
            
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
            
            On Error GoTo errHand
            
            If cbo(0).ListCount > 0 And cbo(0).ListIndex = -1 Then cbo(0).Text = "����"
            If cbo(2).ListCount > 0 And cbo(2).ListIndex = -1 Then cbo(2).Text = "����"
        End If
    End If
    
    mblnFirst = False
    
    Call cboIcon_Click
    cmdOK.Tag = ""
    Call SelAll(txt)
    
    Exit Sub
    
    '���������
errHand:
    mblnFirst = False
    cmdOK.Tag = ""
    
    If ErrCenter = 1 Then Resume
    
End Sub

Private Sub Form_Load()
    blnOK = False
    mblnFirst = True
    
End Sub

Public Function ShowTreeBox(frmMain As Form, ByVal Key As Long) As Boolean
    mKey = Key
    
    frmDefTreeEdit.Show 1, frmMain
    ShowTreeBox = blnOK
    
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
            gstrSQL = "zl_��ѯҳ������_insert(" & lng��� & "," & _
                                            "NULL,'" & _
                                            txt.Text & "'," & _
                                            "NULL," & _
                                            IIf(cboIcon.ItemData(cboIcon.ListIndex) = 0, "NULL", cboIcon.ItemData(cboIcon.ListIndex)) & ",'" & _
                                            IIf(cbo(0).Text = "", "����", cbo(0).Text) & "'," & _
                                            IIf(Val(cbo(1).Text) <= 0, 12, Val(cbo(1).Text)) & "," & _
                                            bytFontForm & "," & _
                                            lblColor.ForeColor & ")"
        Else
            lng��� = mKey
            gstrSQL = "zl_��ѯҳ������_update(" & lng��� & ",NULL,'" & txt.Text & "',NULL," & IIf(cboIcon.ItemData(cboIcon.ListIndex) = 0, "NULL", cboIcon.ItemData(cboIcon.ListIndex)) & ",'" & _
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
        If MsgBox("���ĺ�Ĳ�ѯĿ¼���뱣�������Ч" & vbCrLf & "ȷ�ϲ�������˳���", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Cancel = True
    End If
End Sub

Private Sub txt_Change()
    cmdOK.Tag = "1"
    Call ApplySample
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
