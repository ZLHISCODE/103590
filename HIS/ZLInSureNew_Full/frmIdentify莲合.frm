VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmIdentify���� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ҽ���������ʶ��"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6240
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "����"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmIdentify����.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   6240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   75
      Left            =   120
      TabIndex        =   14
      Top             =   2985
      Width           =   6030
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   405
      Left            =   3765
      TabIndex        =   13
      Top             =   3135
      Width           =   1305
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   405
      Left            =   2310
      TabIndex        =   12
      Top             =   3135
      Width           =   1305
   End
   Begin VB.TextBox txt���֤�� 
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   2100
      MaxLength       =   18
      TabIndex        =   9
      Top             =   1890
      Width           =   2715
   End
   Begin VB.ComboBox cbo�Ա� 
      Height          =   360
      Left            =   3975
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   870
      Width           =   840
   End
   Begin VB.TextBox txt���� 
      Height          =   360
      Left            =   2100
      TabIndex        =   3
      Top             =   870
      Width           =   1335
   End
   Begin VB.TextBox txtAccount 
      Height          =   360
      Left            =   2100
      TabIndex        =   1
      Top             =   360
      Visible         =   0   'False
      Width           =   2715
   End
   Begin VB.TextBox txtBanlance 
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   2100
      MaxLength       =   18
      TabIndex        =   11
      Top             =   2400
      Visible         =   0   'False
      Width           =   2715
   End
   Begin MSComCtl2.DTPicker dtpBirthday 
      Height          =   360
      Left            =   2100
      TabIndex        =   7
      Top             =   1380
      Width           =   2715
      _ExtentX        =   4789
      _ExtentY        =   635
      _Version        =   393216
      CustomFormat    =   "yyyy-mm-dd"
      Format          =   87031808
      CurrentDate     =   37243
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   555
      Picture         =   "frmIdentify����.frx":030A
      Top             =   390
      Width           =   480
   End
   Begin VB.Label lbl���֤�� 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "���֤��"
      Height          =   240
      Left            =   1065
      TabIndex        =   8
      Top             =   1950
      Width           =   960
   End
   Begin VB.Label lbl�������� 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��������"
      Height          =   240
      Left            =   1065
      TabIndex        =   6
      Top             =   1440
      Width           =   960
   End
   Begin VB.Label lbl�Ա� 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�Ա�"
      Height          =   240
      Left            =   3465
      TabIndex        =   4
      Top             =   930
      Width           =   600
   End
   Begin VB.Label lbl���� 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����"
      Height          =   240
      Left            =   1515
      TabIndex        =   2
      Top             =   930
      Width           =   510
   End
   Begin VB.Label lblCard 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�ʺ�"
      Height          =   240
      Left            =   1500
      TabIndex        =   0
      Top             =   420
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label lblbanlance 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�������"
      Height          =   240
      Left            =   1020
      TabIndex        =   10
      Top             =   2460
      Visible         =   0   'False
      Width           =   960
   End
End
Attribute VB_Name = "frmIdentify����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'���볣�����ܶ���ɹ����ģ�������ʹ�õ��ĵط��������壬�ڱ���ʱͳһ�޸�
#Const gverControl = 99  ' 0-��֧�ֶ�̬ҽ��(9.19��ǰ),1-֧�ֶ�̬ҽ���޸��Ӳ���(9.22��ǰ) , _
'    2-����������������ʽ��������һ��;����������ԭʼ��������һ��;�����շ�����������;3-������������GetNextNO();
'    99-���н������Ӹ��Ӳ���(���°�)

Public mlng����ID As Long
Public strPatiInfo As String

Private strCardMask As String
Private blnShowCard As Boolean
Private bytCardNOLen As Byte

Private rsTmp As New ADODB.Recordset
Private strSQL As String
Private mintHIS�շ� As Integer

Private Sub cbo�Ա�_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        zlCommFun.PressKey (vbKeyTab)
    End If
End Sub

Private Sub cmdCancel_Click()
    strPatiInfo = ""
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    If Trim(txt����.Text) = "" Then
        MsgBox "δ��ȷ����������,�����䣡", vbInformation, gstrSysName
        txt����.SetFocus
        
        Exit Sub
    End If
    If mintHIS�շ� = 1 Then
        If Trim(txtAccount.Text) = "" Then
            MsgBox "��������ҽ������,�����䣡", vbInformation, gstrSysName
            txtAccount.SetFocus
    
            Exit Sub
        End If
        
        If Trim(txtBanlance.Text) <> "" Then
            If Not IsNumeric(txtBanlance.Text) Then
                MsgBox "����������Ϊ������!", vbOKOnly, gstrSysName
                txtBanlance.SelStart = 0
                txtBanlance.SelLength = Len(txtBanlance.Text)
                txtBanlance.SetFocus
                Exit Sub
            End If
        Else
            MsgBox "��������Ϊ��!", vbOKOnly + vbExclamation, gstrSysName
            txtBanlance.SelStart = 0
            txtBanlance.SelLength = Len(txtBanlance.Text)
            txtBanlance.SetFocus
            Exit Sub
        End If
    
    End If
    
    If Trim(txt���֤��.Text) <> "" Then
        If Not IsNumeric(txt���֤��.Text) Then
            MsgBox "���֤�ű���Ϊ������1,2,3��", vbOKOnly, gstrSysName
            txt���֤��.SelStart = 0
            txt���֤��.SelLength = Len(txt���֤��.Text)
            txt���֤��.SetFocus
            Exit Sub
        End If
    End If
    
    Call SaveInfo
    Me.Hide
End Sub

Private Sub SaveInfo()
    'New:0����;1ҽ����;2����;3����;4�Ա�;5��������;6���֤;7��λ����(����)
    '0����;1ҽ����;2����;3����;4�Ա�;5��������;6���֤;7��λ����(����);8����ID
        '9����;10.˳���;11��Ա���;12�ʻ����;13��ǰ״̬;14����ID;15��ְ(0,1);16����֤��;17�����;18�Ҷȼ�
    
    Dim strKH As String
    Dim strSelfNo As String
    Dim strSelfPwd As String
    Dim STRNAME As String
    Dim strSex As String
    Dim strBirth As String
    Dim strSFZ As String
    Dim strDWMC As String
    Dim strdwbm As String
    Dim rsTemp As New ADODB.Recordset
    
    If mintHIS�շ� = 1 Then
        strKH = Trim(txtAccount.Text)
        strSelfNo = Trim(txtAccount.Text)
        gcurBanlance = Trim(txtBanlance.Text)
    Else
        strKH = Format(Now, "yyyymmddHHMMSS")
        strSelfNo = Format(Now, "yyyymmddHHMMSS")
        gcurBanlance = 0
    End If
    mlng����ID = Val(txt����.Tag)
    
    If mlng����ID <> 0 Then
        '�������ID��Ϊ�㣬����ȡ�ò������е�ҽ�����뿨�ţ������ٴβ���ҽ�����˵������������
        gstrSQL = "Select ����,ҽ���� From �����ʻ� Where ����=[1] And ����ID=[2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡҽ����Ϣ", TYPE_�ɶ�����, mlng����ID)
        If rsTemp.RecordCount <> 0 Then
            strKH = rsTemp!����
            strSelfNo = rsTemp!ҽ����
        End If
    End If
    
    strSelfPwd = ""
    STRNAME = Trim(txt����.Text)
    strSex = Mid(cbo�Ա�.List(cbo�Ա�.ListIndex), InStr(1, cbo�Ա�.List(cbo�Ա�.ListIndex), "-") + 1)
    strBirth = Format(dtpBirthday.Value, "yyyy-mm-dd")
    strSFZ = Trim(txt���֤��.Text)
    strDWMC = ""
    strdwbm = ""
    
    'New:0����;1ҽ����;2����;3����;4�Ա�;5��������;6���֤;7��λ����(����)
    strPatiInfo = strKH & ";" & strSelfNo & ";" & strSelfPwd & ";" & _
                    STRNAME & ";" & strSex & ";" & _
                    strBirth & ";" & strSFZ & ";" & _
                    strDWMC & "(" & strdwbm & ")"
End Sub

Private Sub dtpBirthday_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        zlCommFun.PressKey (vbKeyTab)
    End If
End Sub

Private Sub Form_Load()
    Dim rsTmp As New Recordset
    Dim i As Long
    
    mintHIS�շ� = GetSetting("ZLSOFT", "����ģ��\zl9Insure", UCase("HIS�շ�"), 0)
    If mintHIS�շ� = 1 Then
        lblCard.Visible = True
        txtAccount.Visible = True
        lblbanlance.Visible = True
        txtBanlance.Visible = True
    End If
    
    strSQL = "Select ����,����,����,ȱʡ��־ From �Ա� Order by ����"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    cbo�Ա�.Clear
    If rsTmp.RecordCount <> 0 Then
        For i = 1 To rsTmp.RecordCount
            cbo�Ա�.AddItem rsTmp!����
            If rsTmp!ȱʡ��־ Then
                cbo�Ա�.ListIndex = i - 1
                cbo�Ա�.ItemData(i - 1) = -1 '����־
            End If
            rsTmp.MoveNext
        Next
        If cbo�Ա�.ListIndex = -1 Then cbo�Ա�.ListIndex = 0
    End If
    dtpBirthday.Value = Now()
    dtpBirthday.MaxDate = Now()
    gcurBanlance = 0
    Me.txt����.Tag = mlng����ID
    
    'ȡϵͳ����
    bytCardNOLen = 7
    Dim strPar As String
    
    #If gverControl >= 4 Then
        blnShowCard = -Not Abs(Val(zlDatabase.GetPara(12, glngSys, , 0)))
        strCardMask = UCase(zlDatabase.GetPara(27, glngSys))
        strPar = zlDatabase.GetPara(20, glngSys, , "7|7|7|7|7")
    #Else
        blnShowCard = -Not Abs(Val(GetPara(12, glngSys, , , 0)))
        strCardMask = UCase(GetPara(27, glngSys))
        strPar = GetPara(20, glngSys, , , "7|7|7|7|7")
    #End If
    If InStr(1, strPar, "|") <> 0 Then
        bytCardNOLen = Val(Split(strPar, "|")(4))
    Else
        bytCardNOLen = Mid(strPar, 5, 1)
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    strPatiInfo = ""
    mlng����ID = 0
End Sub

Private Sub SeekCob(ByVal ConObj As ComboBox, ByVal strSeek As String)
    Dim intSeek As Integer
    
    ConObj.ListIndex = 0
    If strSeek = "" Then Exit Sub
    
    For intSeek = 0 To ConObj.ListCount
        If ConObj.List(intSeek) = strSeek Then
            ConObj.ListIndex = intSeek
            Exit For
        End If
    Next
End Sub

Private Function GetPatiRec(ByVal strAccount As String) As Recordset
    gstrSQL = "select a.����,a.ҽ����,a.����,b.����,b.�Ա�,b.��������,b.���֤��,b.������λ " _
        & " from �����ʻ� a,������Ϣ b " _
        & " where a.����id=b.����id " _
        & " and a.����=[1] and a.����=[2]"
        
        
    'New:0����;1ҽ����;2����;3����;4�Ա�;5��������;6���֤;7��λ����(����)
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strAccount, TYPE_�ɶ�����)
    Set GetPatiRec = rsTmp
End Function

Private Sub txtAccount_GotFocus()
    With txtAccount
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtAccount_KeyPress(KeyAscii As Integer)
    Dim rsPati As New Recordset
    
    If KeyAscii = 13 And Trim(txtAccount.Text) <> "" Then
        Set rsPati = GetPatiRec(txtAccount.Text)
        If Not rsPati.EOF Then
            txt����.Text = IIf(IsNull(rsPati!����), "", rsPati!����)
            Call SeekCob(cbo�Ա�, rsPati!�Ա�)
            dtpBirthday.Value = Format(IIf(IsNull(rsPati!��������), zlDatabase.Currentdate, rsPati!��������), "yyyy-mm-dd")
            txt���֤��.Text = IIf(IsNull(rsPati!���֤��), "", rsPati!���֤��)
       
            txtBanlance.SetFocus
        Else
            txt����.SetFocus
            txt����.SelStart = 0
            txt����.SelLength = Len(txt����.Text)
        End If
    End If
End Sub

Private Sub txtBanlance_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        zlCommFun.PressKey (vbKeyTab)
    Else
        If InStr("0123456789." & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Beep: Exit Sub
    End If
End Sub

Private Sub txt���֤��_GotFocus()
    With txt���֤��
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txt���֤��_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        zlCommFun.PressKey (vbKeyTab)
    Else
        If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Beep: Exit Sub
    End If
End Sub

Private Sub txt����_GotFocus()
    With txt����
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    Dim StrInput As String
    Dim blnCard As Boolean, blnRead As Boolean
    Dim rsTemp As New ADODB.Recordset
    
    blnCard = InputIsCard(txt����, KeyAscii)
    If blnCard And Len(txt����.Text) = bytCardNOLen - 1 And KeyAscii <> 8 _
        Or KeyAscii = 13 And Trim(txt����.Text) <> "" And Mid(txt����.Text, 1, 1) <> "*" Then
        'ͨ�����￨��ȡ�������
        blnRead = True
        StrInput = txt����.Text & Chr(KeyAscii)
        KeyAscii = 0
    End If
    
    If KeyAscii = 13 Then
        If Mid(txt����.Text, 1, 1) = "*" Then '��Ϊ������Ż�סԺ��
            blnRead = True
            StrInput = Val(Mid(txt����.Text, 2))
        End If
    End If
    If blnRead = False Then Exit Sub
    '���������ˢ����Ҳ��������Ż�סԺ�ţ��������SQL���䵱������Ŵ������Խ���ת��Ϊ������
    If Not blnCard And Mid(txt����.Text, 1, 1) <> "*" Then StrInput = Val(StrInput)
    
    '�����������ȡ���˻�����Ϣ
    gstrSQL = " Select ����ID,����,�Ա�,��������,���֤�� From ������Ϣ " & _
              " Where " & IIf(blnCard, "���￨��=", IIf(Me.Tag = "0", "�����=", "סԺ��=")) & StrInput
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���˻�����Ϣ")
    If rsTemp.RecordCount <> 0 Then
        Me.txt����.Tag = rsTemp!����ID
        Me.txt����.Text = rsTemp!����
        Me.dtpBirthday.Value = Format(rsTemp!��������, "yyyy-MM-dd")
        Me.txt���֤��.Text = Nvl(rsTemp!���֤��)
        
        Select Case rsTemp!�Ա�
        Case "��"
            Me.cbo�Ա�.ListIndex = 0
        Case "Ů"
            Me.cbo�Ա�.ListIndex = 1
        Case Else
            Me.cbo�Ա�.ListIndex = 2
        End Select
    End If
    zlCommFun.PressKey (vbKeyTab)
End Sub

Public Function InputIsCard(txtInput As Object, KeyAscii As Integer) As Boolean
'���ܣ��ж�ָ���ı����е�ǰ�����Ƿ���ˢ��,���ݴ���������ʾ
    Dim strText As String, blnCard As Boolean
    Dim arrMask As Variant, i As Long

    '��ǰ�������ʾ������(��δ��ʾ����)
    strText = txtInput.Text
    If txtInput.SelLength = Len(txtInput.Text) Then strText = ""
    If KeyAscii = 8 Then
        If strText <> "" Then strText = Mid(strText, 1, Len(strText) - 1)
    Else
        strText = UCase(strText & Chr(KeyAscii))
    End If
        
    '�ж��Ƿ���ˢ��
    blnCard = False
    If IsNumeric(strText) And IsNumeric(Left(strText, 1)) Then
        blnCard = True
    ElseIf strCardMask <> "" Then
        arrMask = Split(strCardMask, "|")
        For i = 0 To UBound(arrMask)
            If strText Like arrMask(i) & "*" Then
                If IsNumeric(Mid(strText, Len(arrMask(i)) + 1)) And IsNumeric(Mid(strText, Len(arrMask(i)) + 1, 1)) Then
                    blnCard = True
                End If
            End If
        Next
    End If
    
    'ˢ��ʱ�����Ƿ�������ʾ
    If blnCard Then
        txtInput.PasswordChar = IIf(blnShowCard, "", "*")
    Else
        txtInput.PasswordChar = ""
    End If
    
    InputIsCard = blnCard
End Function
