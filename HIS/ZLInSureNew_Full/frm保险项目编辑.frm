VERSION 5.00
Begin VB.Form frm������Ŀ�༭ 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "������Ŀ�༭"
   ClientHeight    =   5835
   ClientLeft      =   2760
   ClientTop       =   3645
   ClientWidth     =   7500
   Icon            =   "frm������Ŀ�༭.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5835
   ScaleWidth      =   7500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.ComboBox cbo��� 
      Height          =   300
      Left            =   4980
      Style           =   2  'Dropdown List
      TabIndex        =   29
      Top             =   4590
      Width           =   2415
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   13
      Left            =   1335
      MaxLength       =   50
      TabIndex        =   27
      Tag             =   "����"
      Top             =   4590
      Width           =   2415
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   12
      Left            =   4995
      MaxLength       =   20
      TabIndex        =   3
      Tag             =   "ҽ������"
      Top             =   1185
      Width           =   2385
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   11
      Left            =   1335
      MaxLength       =   20
      TabIndex        =   23
      Tag             =   "ÿ���������"
      Top             =   4215
      Width           =   1950
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   10
      Left            =   4980
      MaxLength       =   20
      TabIndex        =   21
      Top             =   3840
      Width           =   2430
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   9
      Left            =   1335
      MaxLength       =   20
      TabIndex        =   19
      Tag             =   "��С��װ��λ"
      Top             =   3840
      Width           =   1950
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   8
      Left            =   1335
      MaxLength       =   50
      TabIndex        =   17
      Top             =   3450
      Width           =   6045
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   7
      Left            =   1335
      MaxLength       =   100
      TabIndex        =   15
      Tag             =   "����"
      Top             =   3090
      Width           =   6045
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   6
      Left            =   4980
      MaxLength       =   3
      TabIndex        =   25
      Tag             =   "Ŀ¼����"
      Top             =   4215
      Width           =   2415
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   5
      Left            =   4995
      MaxLength       =   20
      TabIndex        =   9
      Tag             =   "ƴ����"
      Top             =   1965
      Width           =   2385
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   4
      Left            =   1335
      MaxLength       =   20
      TabIndex        =   7
      Tag             =   "�����"
      Top             =   1965
      Width           =   1980
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   3
      Left            =   1335
      MaxLength       =   100
      TabIndex        =   13
      Tag             =   "ͨ��Ӣ����"
      Top             =   2715
      Width           =   6045
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   2
      Left            =   1335
      MaxLength       =   100
      TabIndex        =   11
      Tag             =   "ͨ��������"
      Top             =   2355
      Width           =   6045
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Index           =   1
      Left            =   0
      TabIndex        =   34
      Top             =   930
      Width           =   7620
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   90
      TabIndex        =   32
      Top             =   5265
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Index           =   0
      Left            =   -120
      TabIndex        =   33
      Top             =   5115
      Width           =   7680
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   0
      Left            =   1335
      MaxLength       =   20
      TabIndex        =   1
      Tag             =   "��Ŀ����"
      Top             =   1185
      Width           =   2010
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   1
      Left            =   1335
      MaxLength       =   100
      TabIndex        =   5
      Tag             =   "��Ŀ����"
      Top             =   1575
      Width           =   6045
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   6285
      TabIndex        =   31
      Top             =   5265
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   5040
      TabIndex        =   30
      Top             =   5265
      Width           =   1100
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "ҽ�����(&T)"
      Height          =   180
      Left            =   3960
      TabIndex        =   28
      Top             =   4650
      Width           =   990
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "����(&J)"
      Height          =   180
      Index           =   13
      Left            =   675
      TabIndex        =   26
      Top             =   4650
      Width           =   630
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "ҽ������(&Y)"
      Height          =   180
      Index           =   12
      Left            =   4020
      TabIndex        =   2
      Top             =   1245
      Width           =   990
   End
   Begin VB.Label lblInfor 
      Caption         =   "����ҽ�Ƶ�λ�������Ʒ��,Ŀǰֻ����ȫ�Է���Ŀ����"
      Height          =   240
      Left            =   1065
      TabIndex        =   35
      Top             =   510
      Width           =   5295
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "frm������Ŀ�༭.frx":000C
      Top             =   360
      Width           =   480
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "ÿ���������(&K)"
      Height          =   180
      Index           =   11
      Left            =   -15
      TabIndex        =   22
      Top             =   4275
      Width           =   1350
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "��С������λ(&F)"
      Height          =   180
      Index           =   10
      Left            =   3600
      TabIndex        =   20
      Tag             =   "��С������λ"
      Top             =   3900
      Width           =   1350
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "��С��װ��λ(&X)"
      Height          =   180
      Index           =   9
      Left            =   -15
      TabIndex        =   18
      Top             =   3900
      Width           =   1350
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "��װ���(&G)"
      Height          =   180
      Index           =   8
      Left            =   345
      TabIndex        =   16
      Top             =   3510
      Width           =   990
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "����(&A)"
      Height          =   180
      Index           =   7
      Left            =   705
      TabIndex        =   14
      Top             =   3195
      Width           =   630
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "Ŀ¼����(&M)"
      Height          =   180
      Index           =   6
      Left            =   3975
      TabIndex        =   24
      Top             =   4275
      Width           =   990
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "ƴ����(&P)"
      Height          =   180
      Index           =   5
      Left            =   4185
      TabIndex        =   8
      Top             =   2025
      Width           =   810
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "�����(&W)"
      Height          =   180
      Index           =   4
      Left            =   525
      TabIndex        =   6
      Top             =   2025
      Width           =   810
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "ͨ��Ӣ����(&E)"
      Height          =   180
      Index           =   3
      Left            =   165
      TabIndex        =   12
      Top             =   2775
      Width           =   1170
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "ͨ��������(&Z)"
      Height          =   180
      Index           =   2
      Left            =   165
      TabIndex        =   10
      Top             =   2415
      Width           =   1170
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "��Ŀ����(&U)"
      Height          =   180
      Index           =   0
      Left            =   345
      TabIndex        =   0
      Top             =   1245
      Width           =   990
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "��Ŀ����(&N)"
      Height          =   180
      Index           =   1
      Left            =   345
      TabIndex        =   4
      Top             =   1590
      Width           =   990
   End
End
Attribute VB_Name = "frm������Ŀ�༭"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mblnOK As Boolean
Dim mblnChange As Boolean     '�Ƿ�ı���
Dim mintSuccess As Integer
Dim mstr������� As String
Dim mstr��Ʒ���� As String
'
Dim mblnFirst  As Boolean


Private Sub cbo���_Change()
    mblnChange = True
    SetOk
End Sub

Private Sub cbo���_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    Call LoadCboData
    mblnChange = False
End Sub
Private Sub LoadCboData()
    Dim rsTemp As New ADODB.Recordset
    gstrSQL = "Select * From ҽ����Ŀ���� "
    With rsTemp
        If .State = 1 Then .Close
        .Open gstrSQL, gcnOracle_CQYB
        cbo���.Clear
        Do While Not .EOF
            cbo���.AddItem Nvl(rsTemp!����) & "-" & Nvl(rsTemp!����)
            If Nvl(rsTemp!����) = mstr������� Then
                cbo���.ListIndex = cbo���.NewIndex
            End If
            .MoveNext
        Loop
    End With
End Sub
Private Sub Form_Load()
    mblnFirst = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnChange = False Then Exit Sub
    If MsgBox("�����������˳��Ļ������е��޸Ķ�������Ч��" & vbCrLf & "�Ƿ�ȷ���˳���", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Cancel = 1
    End If
End Sub
'
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
     Dim intIndex As Integer
    
    If IsValid() = False Then Exit Sub
    If SaveData() = False Then Exit Sub
    
    mintSuccess = mintSuccess + 1
    mblnChange = False
    For intIndex = 0 To 12
        txtEdit(intIndex).Text = ""
    Next
    If txtEdit(0).Enabled Then
        txtEdit(0).SetFocus
    End If
    SetOk
End Sub



Public Function EditCard(ByVal frmMain As Object, ByVal str��Ʒ���� As String, ByVal str������� As String) As Boolean
    '����:��������õ�ҽ����Ŀ�����ڽ���ͨѶ�ĳ���
    '����:str���           ��ǰ�༭��ҽ�����ĵ����
    '����ֵ:�༭�ɹ�����True,����ΪFalse
    
    Dim rsTemp As New ADODB.Recordset
    Dim i As Integer

    mstr��Ʒ���� = str��Ʒ����
    mstr������� = str�������
   
    mintSuccess = 0

    
    If str��Ʒ���� <> "" Then
        gstrSQL = "Select ��Ʒ����, ��Ʒ��, ҩƷͨ��������, ҩƷͨ��Ӣ����, ���������1, ƴ��������1, Ŀ¼����, ����, ��װ���, ��С��װ��λ, ��С������λ, ÿ���������, ҽ������,���� From ҽ��������ĿĿ¼ where ��Ʒ����='" & mstr��Ʒ���� & "'"
        
        rsTemp.CursorLocation = adUseClient
        rsTemp.Open gstrSQL, gcnOracle_CQYB, adOpenStatic
        If rsTemp.RecordCount = 0 Then
            ShowMsgbox "����Ŀ�Ѿ�������ɾ�������ܽ����޸�."
            Exit Function
        End If
        For i = 0 To 13
            txtEdit(i).Text = Nvl(rsTemp.Fields(i))
        Next
    End If
    mblnChange = False
    Me.Show 1, frmMain
    EditCard = mintSuccess > 0
End Function

Private Sub txtEdit_Change(Index As Integer)
    mblnChange = True
    SetOk
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    zlControl.TxtSelAll txtEdit(Index)
    zlCommFun.OpenIme True
End Sub


Private Sub txtEdit_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txtEdit_LostFocus(Index As Integer)
    zlCommFun.OpenIme False
End Sub

Private Function IsValid() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��֤���ݵĺϷ���
    '--�����:
    '--������:
    '--��  ��:��֤�Ϸ�,����True,����=false
    '-----------------------------------------------------------------------------------------------------------


    Dim intIndex As Integer
    
    Dim strTemp As String
    
      For intIndex = 0 To 13
        strTemp = Trim(txtEdit(intIndex).Text)
        If intIndex = 0 Or intIndex = 1 Then
            If strTemp = "" Then
                ShowMsgbox txtEdit(intIndex).Tag & "��������!"
                If txtEdit(intIndex).Enabled Then txtEdit(intIndex).SetFocus
                Exit Function
            End If
        End If
        
        If strTemp <> "" Then
            If LenB(StrConv(strTemp, vbFromUnicode)) > txtEdit(intIndex).MaxLength Then
                ShowMsgbox txtEdit(intIndex).Tag & "����,���������" & txtEdit(intIndex).MaxLength / 2 & "�����ֻ�" & txtEdit(intIndex).MaxLength & "���ַ�!"
                If txtEdit(intIndex).Enabled Then txtEdit(intIndex).SetFocus
                Exit Function
            End If
            If InStr(1, strTemp, "'") <> 0 Then
                ShowMsgbox txtEdit(intIndex).Tag & "�������뵥����!"
                If txtEdit(intIndex).Enabled Then txtEdit(intIndex).SetFocus
                Exit Function
            End If
        End If
    Next
    IsValid = True
End Function
Private Function SaveData() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��������
    '--�����:
    '--������:
    '--��  ��:����ɹ�,����True,����=false
    '-----------------------------------------------------------------------------------------------------------
    Dim strTmp As String
    Dim intIndex As Integer
    Dim str������Ʒ���� As String
    Dim rsTemp As New ADODB.Recordset
    Dim strĿ¼���� As String
    If cbo���.Text = "" Then
    Else
        mstr������� = Split(cbo���.Text, "-")(0)
    End If
    gstrSQL = "Select ��Ʒ����,Ŀ¼���� From ҽ��������ĿĿ¼ Where ҽԺ�������='" & mstr������� & "' And  ҽ����ʶ like '__03' and rownum<=1"
    
    zlDataBase.OpenRecordset rsTemp, gstrSQL, Me.Caption
    If rsTemp.EOF Then
        ShowMsgbox "���Է���Ʒ!"
        Exit Function
    End If
    str������Ʒ���� = Nvl(rsTemp!��Ʒ����)
    strĿ¼���� = Nvl(rsTemp!Ŀ¼����)
    SaveData = False
    
    On Error GoTo errHandle
     
    gstrSQL = "ZL_ҽ��������ĿĿ¼_Insert( "

    '���̲�������:
    'ҽԺ�������_IN     IN ҽ��������ĿĿ¼.ҽԺ�������%TYPE,
    '    ҽ������_IN     IN ҽ��������ĿĿ¼.ҽ������%TYPE,
    '    ҩƷͨ��������_IN       IN ҽ��������ĿĿ¼.ҩƷͨ��������%TYPE,
    '    ҩƷͨ��Ӣ����_IN       IN ҽ��������ĿĿ¼.ҩƷͨ��Ӣ����%TYPE,
    '    ��Ʒ����_IN     IN ҽ��������ĿĿ¼.��Ʒ����%TYPE,
    '    ��Ʒ��_IN       IN ҽ��������ĿĿ¼.��Ʒ��%TYPE,
    '    ��Ʒ������_IN       IN ҽ��������ĿĿ¼.��Ʒ������%TYPE,
    '    ����_IN         IN ҽ��������ĿĿ¼.����%TYPE,
    '    ��װ���_IN     IN ҽ��������ĿĿ¼.��װ���%TYPE,
    '    ��С��װ��λ_IN     IN ҽ��������ĿĿ¼.��С��װ��λ%TYPE,
    '    ��С������λ_IN     IN ҽ��������ĿĿ¼.��С������λ%TYPE,
    '    ÿ���������_IN     IN ҽ��������ĿĿ¼.ÿ���������%TYPE,
    '    ���������1_IN      IN ҽ��������ĿĿ¼.���������1%TYPE,
    '    ƴ��������1_IN      IN ҽ��������ĿĿ¼.ƴ��������1%TYPE,
    '    ����_iN
    '    Ŀ¼����_IN     IN ҽ��������ĿĿ¼.Ŀ¼����%TYPE
    '    ��׼����

    gstrSQL = gstrSQL & "'" & _
            mstr������� & "'," & _
            IIf(Trim(txtEdit(12).Text) = "", "NULL", "'" & Trim(txtEdit(12).Text) & "'") & "," & _
            IIf(Trim(txtEdit(2).Text) = "", "NULL", "'" & Trim(txtEdit(2).Text) & "'") & "," & _
            IIf(Trim(txtEdit(3).Text) = "", "NULL", "'" & Trim(txtEdit(3).Text) & "'") & "," & _
            IIf(Trim(txtEdit(0).Text) = "", "NULL", "'" & Trim(txtEdit(0).Text) & "'") & "," & _
            IIf(Trim(txtEdit(1).Text) = "", "NULL", "'" & Trim(txtEdit(1).Text) & "'") & "," & _
            "NULL" & "," & _
            IIf(Trim(txtEdit(7).Text) = "", "NULL", "'" & Trim(txtEdit(7).Text) & "'") & "," & _
            IIf(Trim(txtEdit(8).Text) = "", "NULL", "'" & Trim(txtEdit(8).Text) & "'") & "," & _
            IIf(Trim(txtEdit(9).Text) = "", "NULL", "'" & Trim(txtEdit(9).Text) & "'") & "," & _
            IIf(Trim(txtEdit(10).Text) = "", "NULL", "'" & Trim(txtEdit(10).Text) & "'") & "," & _
            IIf(Trim(txtEdit(11).Text) = "", "NULL", "'" & Trim(txtEdit(11).Text) & "'") & "," & _
            IIf(Trim(txtEdit(4).Text) = "", "NULL", "'" & Trim(txtEdit(4).Text) & "'") & "," & _
            IIf(Trim(txtEdit(5).Text) = "", "NULL", "'" & Trim(txtEdit(5).Text) & "'") & "," & _
            IIf(Trim(txtEdit(13).Text) = "", "NULL", "'" & Trim(txtEdit(13).Text) & "'") & ",'" & _
            strĿ¼���� & "'," & _
            IIf(str������Ʒ���� = "", "NULL", "'" & str������Ʒ���� & "'") & "" & _
            ")"
    
    Call SQLTest(App.ProductName, "����������Ŀ", gstrSQL)
    gcnOracle_CQYB.Execute gstrSQL, , adCmdStoredProc
    Call SQLTest
    
    SaveData = True
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Sub SetOk()
    cmdOK.Enabled = mblnChange
End Sub
