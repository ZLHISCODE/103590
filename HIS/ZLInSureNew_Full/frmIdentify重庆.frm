VERSION 5.00
Begin VB.Form frmIdentify���� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ҽ���������ʶ��"
   ClientHeight    =   3690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6105
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
   ScaleHeight     =   3690
   ScaleWidth      =   6105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.OptionButton opt��� 
      BackColor       =   &H8000000A&
      Caption         =   "������ժ����"
      Height          =   240
      Index           =   3
      Left            =   4110
      TabIndex        =   5
      Top             =   2670
      Width           =   1755
   End
   Begin VB.Frame Frame3 
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
      Left            =   0
      TabIndex        =   10
      Top             =   1350
      Width           =   6660
   End
   Begin VB.Frame Frame2 
      Height          =   1785
      Left            =   3570
      TabIndex        =   11
      Top             =   1260
      Width           =   30
   End
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
      Left            =   -210
      TabIndex        =   9
      Top             =   2985
      Width           =   6660
   End
   Begin VB.OptionButton opt��� 
      BackColor       =   &H8000000A&
      Caption         =   "��������"
      Height          =   240
      Index           =   2
      Left            =   4110
      TabIndex        =   4
      Top             =   2310
      Width           =   1275
   End
   Begin VB.OptionButton opt��� 
      BackColor       =   &H8000000A&
      Caption         =   "���ⲡ����"
      Height          =   240
      Index           =   1
      Left            =   4110
      TabIndex        =   3
      Top             =   1950
      Width           =   1515
   End
   Begin VB.OptionButton opt��� 
      Caption         =   "��ͨ����"
      Height          =   240
      Index           =   0
      Left            =   4110
      TabIndex        =   2
      Top             =   1590
      Value           =   -1  'True
      Width           =   1275
   End
   Begin VB.TextBox txtEdit 
      Height          =   420
      IMEMode         =   3  'DISABLE
      Index           =   0
      Left            =   1575
      MaxLength       =   20
      TabIndex        =   1
      Top             =   1860
      Width           =   1515
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   405
      Left            =   2355
      TabIndex        =   6
      Top             =   3210
      Width           =   1305
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   405
      Left            =   3990
      TabIndex        =   7
      Top             =   3210
      Width           =   1305
   End
   Begin VB.Image Image2 
      Height          =   1005
      Left            =   270
      Picture         =   "frmIdentify����.frx":030A
      Stretch         =   -1  'True
      Top             =   210
      Width           =   1440
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      Caption         =   "������ҽ�Ʊ���"
      BeginProperty Font 
         Name            =   "����"
         Size            =   24
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   480
      Left            =   2190
      TabIndex        =   8
      Top             =   495
      Width           =   3465
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "���˱��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   390
      TabIndex        =   0
      Top             =   1950
      Width           =   1020
   End
End
Attribute VB_Name = "frmIdentify����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mstr���˱�� As String
Private mint���  As Long   '��������Ǳ�ʾ0-���1-סԺ������ʱ��ʾ11-��ͨ���13-���ⲡ���14-�������ȣ�15-������ժ������22-��ͨסԺ
Private mblnOK As Boolean

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim lngIndex As Long
    
    For lngIndex = txtEdit.LBound To txtEdit.UBound
        If zlCommFun.StrIsValid(Trim(txtEdit(lngIndex).Text), txtEdit(lngIndex).MaxLength) = False Then
            txtEdit(lngIndex).SetFocus
            Exit Sub
        End If
    Next
    
    If Trim(txtEdit(0).Text) = "" Then
        MsgBox "δ��������ʻ�,����ͨ����֤��", vbInformation, gstrSysName
        Exit Sub
    End If
    
    mstr���˱�� = UCase(Trim(txtEdit(0).Text))
    If mint��� = 0 Then
        '����
        If opt���(1).Value = True Then
            mint��� = 13
        ElseIf opt���(2).Value = True Then
            mint��� = 14
        ElseIf opt���(3).Value = True Then
            mint��� = 15
        Else
            mint��� = 11
        End If
    Else
        'סԺ
        mint��� = 21
    End If
    
    '��������������ע�����ǰ��֤�Ĳ��˵�ҽ������Ϊ�´������֤��ȱʡҽ����
    If mint��� <> 21 Then
        Call SaveSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "�ϴ�ҽ����", mstr���˱�� & "|" & Format(zlDatabase.Currentdate, "yyyyMMdd"))
    End If
    
    mblnOK = True
    Unload Me
End Sub

Public Function GetIdentify(str���˱�� As String, int��� As Integer) As Boolean
    mblnOK = False
    mstr���˱�� = str���˱��
    mint��� = int���
    
    If int��� <> 0 Then
        '������Ǽ�
        opt���(0).Enabled = False
        opt���(1).Enabled = False
        opt���(2).Enabled = False
        opt���(3).Enabled = False
    End If
    frmIdentify����.Show vbModal
    
    GetIdentify = mblnOK
    If mblnOK = True Then
        str���˱�� = mstr���˱��
        int��� = mint���
    End If
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Form_Load()
    Dim int���� As Integer
    Dim strData As String
    Dim arrData
    Dim rsTemp As New ADODB.Recordset
    '��ȡ��һ��������֤���ҽ����������
    strData = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "�ϴ�ҽ����", "")
    
    int���� = 0
    gstrSQL = "Select ����ֵ From ���ղ��� Where ����=[1] And ������='����ҽ����'"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�ж��Ƿ񱣴��ϴ�ҽ����", TYPE_������)
    If rsTemp.RecordCount <> 0 Then
        int���� = Nvl(rsTemp!����ֵ, 0)
    End If
    
    '����ǽ��죬��������Ϊȱʡֵ
    If strData <> "" And int���� = 1 Then
        If InStr(1, strData, "|") <> 0 Then
            arrData = Split(strData, "|")
            If arrData(1) = Format(zlDatabase.Currentdate, "yyyyMMdd") Then Me.txtEdit(0).Text = UCase(arrData(0))
        End If
    End If
End Sub

Private Sub opt���_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call cmdOK_Click
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    zlControl.TxtSelAll txtEdit(Index)
End Sub

Private Sub txtEdit_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim lng����ID As Long
    Dim intҵ������ As Integer
    Dim rsTemp As New ADODB.Recordset
    
    If Index = 0 Then
        If KeyCode <> vbKeyReturn Then Exit Sub
        lng����ID = GetRegisted(UCase(txtEdit(0).Text))
        If lng����ID = 0 Then Exit Sub
        
        '���»ָ��ϴε�ҵ������
        gstrSQL = "Select ҵ������ From �����ʻ� Where ����=[1] ANd ����ID=[2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡҵ������", TYPE_������, lng����ID)
        If rsTemp.RecordCount <> 0 Then
            intҵ������ = Nvl(rsTemp!ҵ������, 11)
            If intҵ������ = 13 Then
                opt���(1).Value = True
            ElseIf intҵ������ = 14 Then
                opt���(2).Value = True
            ElseIf intҵ������ = 15 Then
                opt���(3).Value = True
            Else
                opt���(0).Value = True
            End If
        End If
    End If
End Sub

Private Function GetRegisted(ByVal strҽ���� As String) As Long
    Dim strDate As String, strStart As String, strEnd As String
    Dim rsTemp As New ADODB.Recordset
    
    strDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
    strStart = strDate & " 00:00:00"
    strEnd = strDate & " 23:59:59"
    '��������ڴ��ھ����¼(�ҺŻ��շ�)���򷵻ز���ID�����򷵻���
    gstrSQL = " Select A.����ID From ������ü�¼ A,���ս����¼ B " & _
              " Where A.��¼���� In (1,4) And A.����ID Is Not NULL" & _
              " And A.�Ǽ�ʱ�� Between to_date('" & strStart & "','yyyy-MM-dd hh24:mi:ss')" & _
              " And to_date('" & strEnd & "','yyyy-MM-dd hh24:mi:ss')" & _
              " And A.����ID=B.��¼ID And B.����=1" & _
              " And A.����ID+0 =(Select ����ID From �����ʻ� Where ����=[1] ANd ҽ����=[2])"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ����ID", TYPE_������, strҽ����)
    If rsTemp.RecordCount = 0 Then Exit Function
    GetRegisted = rsTemp!����ID
End Function
