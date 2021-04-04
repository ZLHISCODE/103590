VERSION 5.00
Begin VB.Form frm�����Ϣ 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��¼�������Ϣ����׼��ICD-10����"
   ClientHeight    =   1965
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5370
   Icon            =   "frm�����Ϣ.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1965
   ScaleWidth      =   5370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   2490
      TabIndex        =   3
      Top             =   1500
      Width           =   1150
   End
   Begin VB.CommandButton cmdCanc 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3870
      TabIndex        =   4
      Top             =   1500
      Width           =   1150
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   60
      TabIndex        =   2
      Top             =   1320
      Width           =   5475
   End
   Begin VB.TextBox txt������Ϣ 
      Height          =   300
      Left            =   1260
      TabIndex        =   0
      Top             =   810
      Width           =   3675
   End
   Begin VB.Label lblOld 
      Height          =   225
      Left            =   240
      TabIndex        =   6
      Top             =   450
      Width           =   4725
   End
   Begin VB.Label lblNote 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   240
      TabIndex        =   5
      Top             =   150
      Width           =   4425
   End
   Begin VB.Label lbl������Ϣ 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "������Ϣ"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   360
      TabIndex        =   1
      Top             =   870
      Width           =   810
   End
End
Attribute VB_Name = "frm�����Ϣ"
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
Private mbln����¼�� As Boolean
Private mstr��ϱ��� As String
Private mstr������� As String

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
    
    If Trim(txt������Ϣ.Text) <> "" Then
        mstr��ϱ��� = Mid(txt������Ϣ.Text, 2, InStr(1, txt������Ϣ.Text, ")") - 2)
        mstr������� = Mid(txt������Ϣ.Text, InStr(1, txt������Ϣ.Text, ")") + 1)
    End If
    
    Unload Me
End Sub

Public Sub ShowME(ByVal lng����ID As Long, ByRef str��ϱ��� As String, ByRef str������� As String, Optional ByVal bln����¼�� As Boolean = True)
    mlng����ID = lng����ID
    mbln����¼�� = bln����¼��
    mstr��ϱ��� = str��ϱ���
    mstr������� = str�������
    Me.Show 1
    str��ϱ��� = mstr��ϱ���
    str������� = mstr�������
End Sub

Private Sub Form_Load()
    Dim rsTemp As New ADODB.Recordset
    lblOld = "ԭ���Ϊ��(" & mstr��ϱ��� & ")" & mstr�������
    mstr��ϱ��� = ""
    mstr������� = ""
    
    '��ȡ������Ϣ(һ�����˲��������ڶ��ҽ��)
    gstrSQL = " Select A.����,A.�Ա�,B.ҽ����,B.���� " & _
              " From ������Ϣ A,�����ʻ� B" & _
              " Where A.����ID=B.����ID And A.����ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ������Ϣ", mlng����ID)
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
                txt������Ϣ.Text = "(" & rsTmp!���� & ")" & rsTmp!����
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
