VERSION 5.00
Begin VB.Form frm�����Ϣ_�Ĵ� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��¼�������Ϣ����׼��ICD-10����"
   ClientHeight    =   2880
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5370
   Icon            =   "frm�����Ϣ_�Ĵ�.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   5370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.TextBox Txt������Ϣ3 
      Height          =   300
      Left            =   1260
      TabIndex        =   2
      Top             =   1620
      Width           =   3675
   End
   Begin VB.TextBox Txt������Ϣ2 
      Height          =   300
      Left            =   1260
      TabIndex        =   1
      Top             =   1080
      Width           =   3675
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   2490
      TabIndex        =   3
      Top             =   2310
      Width           =   1150
   End
   Begin VB.CommandButton cmdCanc 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3870
      TabIndex        =   6
      Top             =   2310
      Width           =   1150
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   60
      TabIndex        =   5
      Top             =   2130
      Width           =   5475
   End
   Begin VB.TextBox txt������Ϣ 
      Height          =   300
      Left            =   1260
      TabIndex        =   0
      Top             =   540
      Width           =   3675
   End
   Begin VB.Label Lbl������Ϣ3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "������Ϣ3"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   360
      TabIndex        =   9
      Top             =   1680
      Width           =   810
   End
   Begin VB.Label Lbl������Ϣ2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "������Ϣ2"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   360
      TabIndex        =   8
      Top             =   1140
      Width           =   810
   End
   Begin VB.Label lblNote 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   510
      TabIndex        =   7
      Top             =   150
      Width           =   4425
   End
   Begin VB.Label lbl������Ϣ 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "������Ϣ1"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   360
      TabIndex        =   4
      Top             =   600
      Width           =   810
   End
End
Attribute VB_Name = "frm�����Ϣ_�Ĵ�"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type POINTAPI
     x As Long
     y As Long
End Type
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Private lngTXTProc As Long '����Ĭ�ϵ���Ϣ�����ĵ�ַ
Private Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

Private str�Ա� As String
Private mlng����ID As Long
Private mlng��ҳID As Long

Private mint������� As Integer
Private mint������� As Integer
Private mint���� As Integer

Private mbln����¼�� As Boolean


Private Sub cmdCanc_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If mbln����¼�� Then
        If Val(txt������Ϣ.Tag) = 0 Then
            MsgBox "ҽ������Ҫ�󣬱��밴ICD-10��׼¼�������Ϣ��", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    

    'HIS+

    If txt������Ϣ.Text <> "" Then
        gstrSQL = "ZL_������������Ϣ_INSERT(" & mint������� & "," & mlng����ID & "," & mlng��ҳID & "," & txt������Ϣ.Tag & "," & _
                  "'" & txt������Ϣ.Text & "',1 )"
    End If
    ExecuteProcedure_�ϳ����� "������ϼ�¼���м��"
    If Txt������Ϣ2.Text <> "" Then
        gstrSQL = "ZL_������������Ϣ_INSERT(" & mint������� & "," & mlng����ID & "," & mlng��ҳID & "," & Txt������Ϣ2.Tag & "," & _
                  "'" & Txt������Ϣ2.Text & "',2 )"
    End If
    ExecuteProcedure_�ϳ����� "������ϼ�¼���м��"
    If Txt������Ϣ3.Text <> "" Then
        gstrSQL = "ZL_������������Ϣ_INSERT(" & mint������� & "," & mlng����ID & "," & mlng��ҳID & "," & Txt������Ϣ3.Tag & "," & _
                  "'" & Txt������Ϣ3.Text & "',3 )"
    End If
    ExecuteProcedure_�ϳ����� "������ϼ�¼���м��"
    Unload Me
End Sub

Public Sub ShowME(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal int������� As Integer, ByVal int������� As Integer, ByVal int���� As Integer, Optional ByVal bln����¼�� As Boolean = True)
'mint�������:2��ʾ��Ժ��� 3��ʾ��Ժ���(��ҳ����)
'mint�������:���������������������1(�ڳ�Ժ�������¼��һ�����),���������֧������3��
    mlng����ID = lng����ID
    mint������� = int�������
    mint������� = int�������
    mint���� = int����
    mlng��ҳID = lng��ҳID
    
    mbln����¼�� = bln����¼��
    
    Me.Show 1

End Sub

Private Sub Form_Load()
    Dim rsTemp As New ADODB.Recordset
    
    If mint������� = 1 Then
        Me.Caption = "�벹��¼����Ժ�����Ϣ����׼��ICD-10����"
    Else
        Me.Caption = "�벹��¼���Ժ�����Ϣ����׼��ICD-10����"
    End If
    If mint������� = 1 Then
        lbl������Ϣ.Enabled = True
        txt������Ϣ.Enabled = True
        Lbl������Ϣ2.Enabled = False
        Txt������Ϣ2.Enabled = False
        Lbl������Ϣ3.Enabled = False
        Txt������Ϣ3.Enabled = False
    End If
    If mint������� = 2 Then
        lbl������Ϣ.Enabled = True
        txt������Ϣ.Enabled = True
        Lbl������Ϣ2.Enabled = True
        Txt������Ϣ2.Enabled = True
        Lbl������Ϣ3.Enabled = False
        Txt������Ϣ3.Enabled = False
    End If
    If mint������� = 3 Then
        lbl������Ϣ.Enabled = True
        txt������Ϣ.Enabled = True
        Lbl������Ϣ2.Enabled = True
        Txt������Ϣ2.Enabled = True
        Lbl������Ϣ3.Enabled = True
        Txt������Ϣ3.Enabled = True
    End If
    '��ȡ������Ϣ
    gstrSQL = " Select A.����,A.�Ա�,B.ҽ����,B.���� " & _
              " From ������Ϣ A,�����ʻ� B" & _
              " Where A.����ID=B.����ID And A.����ID=" & mlng����ID & " And B.����=" & mint����
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ������Ϣ")
    lblNote.Caption = "����:" & rsTemp!���� & "  �Ա�:" & rsTemp!�Ա� & "  ҽ����:" & Nvl(rsTemp!ҽ����) & "  ����:" & Nvl(rsTemp!����)
    str�Ա� = rsTemp!�Ա�
End Sub



Private Sub txt������Ϣ_GotFocus()
    zlControl.TxtSelAll txt������Ϣ
End Sub

Private Sub txt������Ϣ_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnCancel As Boolean
    Dim strLike As String, str�Ա� As String
    Dim vPoint As POINTAPI, StrInput As String
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txt������Ϣ.Text = lbl������Ϣ.Tag And txt������Ϣ.Text <> "" Then
            Call zlCommFun.PressKey(vbKeyTab)
        ElseIf txt������Ϣ.Text = "" Then
            txt������Ϣ.Tag = "": lbl������Ϣ.Tag = ""
            Call zlCommFun.PressKey(vbKeyTab) '��������
        Else
            strLike = IIf(GetSetting("ZLSOFT", "����ģ��\����", "����ƥ��", 0) = 0, "%", "")
            StrInput = UCase(txt������Ϣ.Text)
            If str�Ա� = "��" Then
                str�Ա� = " And (A.�Ա�����='��' Or A.�Ա����� is NULL)"
            ElseIf str�Ա� = "Ů" Then
                str�Ա� = " And (A.�Ա�����='Ů' Or A.�Ա����� is NULL)"
            Else
                str�Ա� = ""
            End If
            strSQL = "Select A.ID,A.����,A.����,A.����,A.����,A.˵��,A.�Ա�����,B.���" & _
                " From ��������Ŀ¼ A,����������� B" & _
                " Where A.���=B.���� And A.��� Not IN('B','Z')" & _
                " And (A.���� Like '" & StrInput & "%'" & _
                " Or Upper(A.����) Like '" & strLike & StrInput & "%'" & _
                " Or Upper(A.����) Like '" & strLike & StrInput & "%'" & _
                " Or Upper(A.����) Like '" & strLike & StrInput & "%')" & _
                " And Rownum<=100" & str�Ա� & _
                " Order by A.���,A.����"
            vPoint = GetCoordPos(Me.hwnd, txt������Ϣ.Left, txt������Ϣ.Top)
            Set rsTmp = zlDatabase.ShowSelect(Me, strSQL, 0, "��������Input", , , , , , True, vPoint.x, vPoint.y, txt������Ϣ.Height, blnCancel, , True)
            If Not rsTmp Is Nothing Then
                txt������Ϣ.Tag = rsTmp!ID
                txt������Ϣ.Text = rsTmp!����
                lbl������Ϣ.Tag = txt������Ϣ.Text '���ڻָ���ʾ
                Call zlCommFun.PressKey(vbKeyTab)
            Else
                If Not blnCancel Then
                    MsgBox "û���ҵ�ƥ��ļ������롣", vbInformation, gstrSysName
                End If
                If lbl������Ϣ.Tag <> "" Then txt������Ϣ.Text = lbl������Ϣ.Tag
                Call txt������Ϣ_GotFocus
                txt������Ϣ.SetFocus
            End If
        End If
    End If
End Sub

Private Function GetCoordPos(ByVal lngHwnd As Long, ByVal lngX As Long, ByVal lngY As Long) As POINTAPI
'���ܣ��ÿؼ���ָ����������Ļ�е�λ��(Twip)
    Dim vPoint As POINTAPI
    vPoint.x = lngX / Screen.TwipsPerPixelX: vPoint.y = lngY / Screen.TwipsPerPixelY
    Call ClientToScreen(lngHwnd, vPoint)
    vPoint.x = vPoint.x * Screen.TwipsPerPixelX: vPoint.y = vPoint.y * Screen.TwipsPerPixelY
    GetCoordPos = vPoint
End Function

Private Function GetControlRect(ByVal lngHwnd As Long) As RECT
'���ܣ���ȡָ���ؼ�����Ļ�е�λ��(Twip)
    Dim vRect As RECT
    Call GetWindowRect(lngHwnd, vRect)
    vRect.Left = vRect.Left * Screen.TwipsPerPixelX
    vRect.Right = vRect.Right * Screen.TwipsPerPixelX
    vRect.Top = vRect.Top * Screen.TwipsPerPixelY
    vRect.Bottom = vRect.Bottom * Screen.TwipsPerPixelY
    GetControlRect = vRect
End Function

Private Sub txt������Ϣ2_GotFocus()
    zlControl.TxtSelAll Txt������Ϣ2
End Sub

Private Sub txt������Ϣ3_GotFocus()
    zlControl.TxtSelAll Txt������Ϣ3
End Sub
Private Sub txt������Ϣ2_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnCancel As Boolean
    Dim strLike As String, str�Ա� As String
    Dim vPoint As POINTAPI, StrInput As String
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Txt������Ϣ2.Text = Lbl������Ϣ2.Tag And Txt������Ϣ2.Text <> "" Then
            Call zlCommFun.PressKey(vbKeyTab)
        ElseIf Txt������Ϣ2.Text = "" Then
            Txt������Ϣ2.Tag = "": lbl������Ϣ.Tag = ""
            Call zlCommFun.PressKey(vbKeyTab) '��������
        Else
            strLike = IIf(GetSetting("ZLSOFT", "����ģ��\����", "����ƥ��", 0) = 0, "%", "")
            StrInput = UCase(Txt������Ϣ2.Text)
            If str�Ա� = "��" Then
                str�Ա� = " And (A.�Ա�����='��' Or A.�Ա����� is NULL)"
            ElseIf str�Ա� = "Ů" Then
                str�Ա� = " And (A.�Ա�����='Ů' Or A.�Ա����� is NULL)"
            Else
                str�Ա� = ""
            End If
            strSQL = "Select A.ID,A.����,A.����,A.����,A.����,A.˵��,A.�Ա�����,B.���" & _
                " From ��������Ŀ¼ A,����������� B" & _
                " Where A.���=B.���� And A.��� Not IN('B','Z')" & _
                " And (A.���� Like '" & StrInput & "%'" & _
                " Or Upper(A.����) Like '" & strLike & StrInput & "%'" & _
                " Or Upper(A.����) Like '" & strLike & StrInput & "%'" & _
                " Or Upper(A.����) Like '" & strLike & StrInput & "%')" & _
                " And Rownum<=100" & str�Ա� & _
                " Order by A.���,A.����"
            vPoint = GetCoordPos(Me.hwnd, Txt������Ϣ2.Left, Txt������Ϣ2.Top)
            Set rsTmp = zlDatabase.ShowSelect(Me, strSQL, 0, "��������Input", , , , , , True, vPoint.x, vPoint.y, Txt������Ϣ2.Height, blnCancel, , True)
            If Not rsTmp Is Nothing Then
                Txt������Ϣ2.Tag = rsTmp!ID
                Txt������Ϣ2.Text = rsTmp!����
                Lbl������Ϣ2.Tag = Txt������Ϣ2.Text '���ڻָ���ʾ
                Call zlCommFun.PressKey(vbKeyTab)
            Else
                If Not blnCancel Then
                    MsgBox "û���ҵ�ƥ��ļ������롣", vbInformation, gstrSysName
                End If
                If Lbl������Ϣ2.Tag <> "" Then Txt������Ϣ2.Text = Lbl������Ϣ2.Tag
                Call txt������Ϣ2_GotFocus
                Txt������Ϣ2.SetFocus
            End If
        End If
    End If
End Sub
Private Sub txt������Ϣ3_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnCancel As Boolean
    Dim strLike As String, str�Ա� As String
    Dim vPoint As POINTAPI, StrInput As String
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Txt������Ϣ3.Text = Lbl������Ϣ3.Tag And Txt������Ϣ3.Text <> "" Then
            Call zlCommFun.PressKey(vbKeyTab)
        ElseIf Txt������Ϣ3.Text = "" Then
            Txt������Ϣ3.Tag = "": Lbl������Ϣ3.Tag = ""
            Call zlCommFun.PressKey(vbKeyTab) '��������
        Else
            strLike = IIf(GetSetting("ZLSOFT", "����ģ��\����", "����ƥ��", 0) = 0, "%", "")
            StrInput = UCase(Txt������Ϣ3.Text)
            If str�Ա� = "��" Then
                str�Ա� = " And (A.�Ա�����='��' Or A.�Ա����� is NULL)"
            ElseIf str�Ա� = "Ů" Then
                str�Ա� = " And (A.�Ա�����='Ů' Or A.�Ա����� is NULL)"
            Else
                str�Ա� = ""
            End If
            strSQL = "Select A.ID,A.����,A.����,A.����,A.����,A.˵��,A.�Ա�����,B.���" & _
                " From ��������Ŀ¼ A,����������� B" & _
                " Where A.���=B.���� And A.��� Not IN('B','Z')" & _
                " And (A.���� Like '" & StrInput & "%'" & _
                " Or Upper(A.����) Like '" & strLike & StrInput & "%'" & _
                " Or Upper(A.����) Like '" & strLike & StrInput & "%'" & _
                " Or Upper(A.����) Like '" & strLike & StrInput & "%')" & _
                " And Rownum<=100" & str�Ա� & _
                " Order by A.���,A.����"
            vPoint = GetCoordPos(Me.hwnd, Txt������Ϣ3.Left, Txt������Ϣ3.Top)
            Set rsTmp = zlDatabase.ShowSelect(Me, strSQL, 0, "��������Input", , , , , , True, vPoint.x, vPoint.y, Txt������Ϣ3.Height, blnCancel, , True)
            If Not rsTmp Is Nothing Then
                Txt������Ϣ3.Tag = rsTmp!ID
                Txt������Ϣ3.Text = rsTmp!����
                Lbl������Ϣ3.Tag = Txt������Ϣ3.Text '���ڻָ���ʾ
                Call zlCommFun.PressKey(vbKeyTab)
            Else
                If Not blnCancel Then
                    MsgBox "û���ҵ�ƥ��ļ������롣", vbInformation, gstrSysName
                End If
                If Lbl������Ϣ3.Tag <> "" Then Txt������Ϣ3.Text = Lbl������Ϣ3.Tag
                Call txt������Ϣ3_GotFocus
                Txt������Ϣ3.SetFocus
            End If
        End If
    End If
End Sub

