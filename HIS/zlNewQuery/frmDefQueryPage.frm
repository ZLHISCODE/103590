VERSION 5.00
Begin VB.Form frmDefQueryPage 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ҳ��༭"
   ClientHeight    =   5700
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   9180
   Icon            =   "frmDefQueryPage.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   9180
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   1
      Left            =   1110
      MaxLength       =   100
      TabIndex        =   32
      Top             =   4800
      Width           =   7995
   End
   Begin VB.Frame fra 
      Caption         =   "������Ϣ"
      Height          =   4665
      Index           =   0
      Left            =   45
      TabIndex        =   0
      Top             =   60
      Width           =   4920
      Begin VB.CommandButton cmdOpen 
         Caption         =   "��"
         Height          =   240
         Index           =   2
         Left            =   4425
         TabIndex        =   10
         Top             =   1425
         Width           =   285
      End
      Begin VB.TextBox txtEdit 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Index           =   3
         Left            =   720
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   1395
         Width           =   4005
      End
      Begin VB.TextBox txt 
         Height          =   300
         Left            =   720
         MaxLength       =   30
         TabIndex        =   5
         Top             =   660
         Width           =   4005
      End
      Begin VB.ListBox lst 
         Height          =   1320
         Left            =   720
         Style           =   1  'Checkbox
         TabIndex        =   12
         Top             =   1755
         Width           =   4005
      End
      Begin VB.ComboBox cbo 
         Height          =   300
         Left            =   720
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   3225
         Width           =   4005
      End
      Begin VB.CommandButton cmdTest 
         Caption         =   "����(&H)"
         Height          =   350
         Left            =   3615
         TabIndex        =   15
         Top             =   3585
         Width           =   1100
      End
      Begin VB.TextBox txtEdit 
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   0
         Left            =   960
         MaxLength       =   10
         TabIndex        =   3
         Text            =   "11111"
         Top             =   345
         Width           =   900
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   2
         Left            =   720
         MaxLength       =   15
         TabIndex        =   7
         Top             =   1020
         Width           =   4005
      End
      Begin VB.TextBox txtTemp 
         Enabled         =   0   'False
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   720
         TabIndex        =   2
         TabStop         =   0   'False
         Tag             =   "����"
         Text            =   "1111111111"
         Top             =   300
         Width           =   1275
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "���(&S)"
         Height          =   180
         Left            =   75
         TabIndex        =   11
         Top             =   1815
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "����(&N)"
         Height          =   180
         Left            =   75
         TabIndex        =   4
         Top             =   720
         Width           =   630
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "����(&M)"
         Height          =   180
         Left            =   75
         TabIndex        =   13
         Top             =   3285
         Width           =   630
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "����(&B)"
         Height          =   180
         Index           =   0
         Left            =   75
         TabIndex        =   1
         Top             =   360
         Width           =   630
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "����(&Y)"
         Height          =   180
         Index           =   2
         Left            =   75
         TabIndex        =   6
         Top             =   1110
         Width           =   630
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "�ϼ�(&U)"
         Height          =   180
         Index           =   3
         Left            =   75
         TabIndex        =   8
         Top             =   1455
         Width           =   630
      End
   End
   Begin VB.Frame fra 
      Caption         =   "��������"
      Height          =   2085
      Index           =   2
      Left            =   4965
      TabIndex        =   21
      Top             =   2640
      Width           =   4110
      Begin VB.CommandButton cmdPos 
         Height          =   345
         Index           =   1
         Left            =   3630
         Picture         =   "frmDefQueryPage.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "��ʾ���������ڲ�ѯ�е�λ��"
         Top             =   1470
         Width           =   345
      End
      Begin VB.CommandButton cmdOpen 
         Height          =   345
         Index           =   1
         Left            =   3630
         Picture         =   "frmDefQueryPage.frx":0596
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "ѡ����������"
         Top             =   210
         Width           =   345
      End
      Begin VB.CommandButton cmdClear 
         Height          =   345
         Index           =   1
         Left            =   3630
         Picture         =   "frmDefQueryPage.frx":0643
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "�����������"
         Top             =   615
         Width           =   345
      End
      Begin zl9NewQuery.ctlPicture UsrPic 
         Height          =   1590
         Index           =   1
         Left            =   75
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   225
         Width           =   3450
         _ExtentX        =   6085
         _ExtentY        =   2805
      End
      Begin VB.Label lblSize 
         AutoSize        =   -1  'True
         Caption         =   "800 X 600"
         Height          =   180
         Index           =   1
         Left            =   90
         TabIndex        =   30
         Top             =   1845
         Width           =   810
      End
   End
   Begin VB.Frame fra 
      Caption         =   "����ͼƬ"
      Height          =   2535
      Index           =   1
      Left            =   4965
      TabIndex        =   16
      Top             =   60
      Width           =   4110
      Begin VB.CommandButton cmdPos 
         Height          =   345
         Index           =   0
         Left            =   3615
         Picture         =   "frmDefQueryPage.frx":06E9
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "��ʾ����ͼƬ�ڲ�ѯ�е�λ��"
         Top             =   1905
         Width           =   345
      End
      Begin VB.CommandButton cmdOpen 
         Height          =   345
         Index           =   0
         Left            =   3615
         Picture         =   "frmDefQueryPage.frx":0C73
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "ѡ�񱳾�ͼƬ"
         Top             =   210
         Width           =   345
      End
      Begin VB.CommandButton cmdClear 
         Height          =   345
         Index           =   0
         Left            =   3615
         Picture         =   "frmDefQueryPage.frx":0D20
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "�������ͼƬ"
         Top             =   615
         Width           =   345
      End
      Begin zl9NewQuery.ctlPicture UsrPic 
         Height          =   2010
         Index           =   0
         Left            =   75
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   225
         Width           =   3450
         _ExtentX        =   6085
         _ExtentY        =   3545
      End
      Begin VB.Label lblSize 
         AutoSize        =   -1  'True
         Caption         =   "800 X 600"
         Height          =   180
         Index           =   0
         Left            =   105
         TabIndex        =   29
         Top             =   2280
         Width           =   810
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   75
      TabIndex        =   28
      Top             =   5250
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   7995
      TabIndex        =   27
      Top             =   5250
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   6765
      TabIndex        =   26
      Top             =   5250
      Width           =   1100
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�������(&A)"
      Height          =   180
      Left            =   75
      TabIndex        =   31
      Top             =   4845
      Width           =   990
   End
End
Attribute VB_Name = "frmDefQueryPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnFirst As Boolean
Private mKey As Long
Private mOK As Boolean

Private mvarSvrPicRange As String           '��������ͼƬ�ķ�Χ
Private mvarSvrPicType As String            '��������ͼƬ������

Private mlngKey As Long
Private mlngUpKey As Long
Private mstr�ϼ�ID As String
Private mstr�ϼ����� As String
Private mstr���� As String
Const mlng���볤�� = 10

Private Sub GetTreeCode(ByVal lngUpKey As Long)
    '��ȡ���ͽṹ�ı������,�����ϼ�����,��������
    
    If lngUpKey = 0 Then
        '���û��ָ���ϼ�
        mstr�ϼ����� = ""
        txtTemp.Text = ""
        
        txtEdit(3).Text = "��"
        
        'ȡ���ϼ����룬�������볤�ȵ�ֵ
        txtTemp.MaxLength = GetLocalCodeLength("", "��ѯҳ��Ŀ¼")
        
    Else
        'ָ�����ϼ�
        gstrSQL = "select ���� as �ϼ�����,ҳ������ as �ϼ�����,ҳ����� as �ϼ�ID from ��ѯҳ��Ŀ¼ where ҳ�����=[1]"
        Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngUpKey)
                        
        mstr�ϼ�ID = IIf(IsNull(gRs("�ϼ�ID")), "", gRs("�ϼ�ID"))
        mstr�ϼ����� = IIf(IsNull(gRs("�ϼ�����")), "", gRs("�ϼ�����"))
        txtEdit(3).Text = IIf(IsNull(gRs("�ϼ�����")), "��", gRs("�ϼ�����"))
        txtEdit(3).Tag = lngUpKey
        txtTemp.MaxLength = 0
        txtTemp.Text = mstr�ϼ�����
        
        '�жϱ����Ƿ�����
        If Len(mstr�ϼ�����) >= mlng���볤�� Then
            MsgBox "�����������ӷ����ˣ����볤���Ѿ��þ���", vbExclamation, gstrSysName
            Exit Sub
        End If
        
        'ȡ���ϼ����룬�������볤�ȵ�ֵ
        txtTemp.MaxLength = GetLocalCodeLength(mstr�ϼ�ID, "��ѯҳ��Ŀ¼")
    End If
        
    txtEdit(0).MaxLength = IIf(txtTemp.MaxLength = 0, mlng���볤��, txtTemp.MaxLength) - Len(mstr�ϼ�����)
    txtEdit(0).Text = Mid(txtEdit(0).Text, Len(txtTemp.Text) + 1)
    
    If mKey = 0 Then txtEdit(0).Text = GetMaxLocalCode(mstr�ϼ�ID, "��ѯҳ��Ŀ¼")
End Sub

Public Function ShowPageEdit(frmMain As Object, ByVal Key As Long, ByVal lngUpKey As Long) As Boolean

    mKey = Key
    mlngUpKey = lngUpKey

    frmDefQueryPage.Show 1, frmMain
    ShowPageEdit = mOK
End Function

Private Sub cbo_Click()
    cmdOK.Tag = "1"
End Sub

Private Sub cbo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmdOK.SetFocus
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdClear_Click(Index As Integer)
    If UsrPic(Index).Tag <> "" Then
        UsrPic(Index).Tag = ""
        UsrPic(Index).Cls
        cmdOK.Tag = "1"
    End If
End Sub

Private Sub cmdOK_Click()
    If SaveData Then
        mOK = True
        If mKey = 0 Then
            txt.Text = ""
            txtEdit(0).Text = ""
            txtEdit(2).Text = ""
            txtEdit(1).Text = ""
            
            UsrPic(0).Tag = ""
            UsrPic(1).Tag = ""
            lblSize(0).Caption = ""
            lblSize(1).Caption = ""
            UsrPic(0).Cls
            UsrPic(1).Cls
            txtEdit(0).Text = GetMaxLocalCode(txtEdit(3).Tag, "��ѯҳ��Ŀ¼")
            cmdOK.Tag = ""
            txtEdit(0).SetFocus
        Else
            cmdOK.Tag = ""
            Unload Me
        End If
    End If
End Sub


Private Sub cmdOpen_Click(Index As Integer)
    Dim lngKey As Long
    Dim strFilter As String
    Dim strTitle As String
            
    Dim strSQL As String
    Dim blnRe As Boolean
    Dim str���� As String
    Dim strID As String
    Dim strRerurnID As String
    Dim str���� As String
    Dim int����  As Integer
            
            
    If Index = 2 Then
        
        strSQL = "Select ҳ����� AS id,�ϼ���� AS �ϼ�id,ҳ������ AS ����,����,0 as ĩ�� From ��ѯҳ��Ŀ¼ Where (ĩ�� IS NULL OR ĩ��=0)  Start with �ϼ���� is null connect by prior ҳ����� =�ϼ���� "
        
        strID = txtEdit(3).Tag
        str���� = txtEdit(3).Text
        str���� = txtTemp.Text & txtEdit(0).Text
            
        blnRe = frm����ѡ��.ShowTree(strSQL, strID, str����, mstr�ϼ�����, "", Me.Caption, "����ҳ�����", , mstr����)
    
        If blnRe Then       '�µı����Ŀ��
            
            mlngUpKey = Val(strID)
            txtEdit(3).Tag = strID
            txtEdit(3).Text = str����
            Call GetTreeCode(mlngUpKey)
            txtEdit(0).Text = GetMaxLocalCode(strID, "��ѯҳ��Ŀ¼")
            cmdOK.Tag = "1"
        End If
    Else
        strFilter = IIf(Index = 0, "4;0;1;2;3;9", "1;0;2;3;4;9")
        Select Case Index
        Case 0
            strTitle = "���ҳ�汳��"
        Case 1
            strTitle = "���ҳ����������"
        End Select
        If frmPicSelect.OpenPictureBox(Me, strTitle, strFilter, lngKey, mvarSvrPicRange, mvarSvrPicType) Then
            '����ͼƬ��ʾ
            Call ShowPicture(lngKey, Index)
            UsrPic(Index).Tag = lngKey
            cmdOK.Tag = "1"
        End If
    End If
End Sub

Private Sub cmdPos_Click(Index As Integer)
    Select Case Index
    Case 0
        Call frmPosSample.ShowPageSample("ҳ�汳��")
    Case 1
        Call frmPosSample.ShowPageSample("��������")
    End Select
End Sub

Private Sub cmdTest_Click()
    Dim vFileData As New FileSystemObject
    Dim strFile As String
    
    Call MusicClose
    
    
    If cbo.ListIndex < 0 Then Exit Sub
    If cbo.ItemData(cbo.ListIndex) <= 0 Then Exit Sub
    
    '1.���ͼ��Ŀ¼�Ƿ����
    On Error Resume Next
    vFileData.CreateFolder App.Path & "\ͼ��"
    
    '2.��鱾ϵͳ�п���ʹ�õ���ͼƬ
    gstrSQL = "select ���,����,����,�޸����� from ��ѯͼƬԪ�� where ���=[1]"
    Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(cbo.ItemData(cbo.ListIndex)))
    If gRs.BOF Then Exit Sub
    
    strFile = IIf(IsNull(gRs!����), "", gRs!����)
    If strFile <> "" Then Call CheckFileNew(strFile, IIf(IsNull(gRs!����), 0, gRs!����), gRs!���, gRs!�޸�����, vFileData)
            
    Call MusicPlay(strFile)
End Sub

Private Sub Command1_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    DoEvents
    
    '��ʼ������
    lst.AddItem "0-��׼"
    lst.ItemData(lst.NewIndex) = 0
    Call SelectListItem(0)
    
    cbo.AddItem "[��]"
    gstrSQL = "select ���,���� from ��ѯͼƬԪ�� where ����=3"
    Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    If gRs.BOF = False Then
        While Not gRs.EOF
            cbo.AddItem IIf(IsNull(gRs!����), "", gRs!����)
            cbo.ItemData(cbo.NewIndex) = IIf(IsNull(gRs!���), 0, gRs!���)
            gRs.MoveNext
        Wend
    End If
    cbo.ListIndex = 0
    
    If mKey <> 0 Then
        If frmDefQuery.lvw.SelectedItem.Tag = "1" Then
            txt.Enabled = False
            lst.Enabled = False
            If frmDefQuery.lvw.SelectedItem.Text <> "ר�ҽ���" And mKey > 0 Then
                'tbs.TabEnabled(1) = False
                Fra(1).Enabled = False
            End If
        End If
                
        gstrSQL = "select A.�������,A.����,A.����,A.ҳ������,A.ҳ����,A.��������,A.ҳ�汳��,A.��������,A.ҳ�汳��,�������� from ��ѯҳ��Ŀ¼ A where A.ҳ�����=[1]"
        Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mKey)
        If gRs.BOF = False Then
            txt.Text = IIf(IsNull(gRs!ҳ������), "", gRs!ҳ������)
            Call SelectListItem(IIf(IsNull(gRs!ҳ����), 0, gRs!ҳ����))
            
            Call ShowPicture(IIf(IsNull(gRs!ҳ�汳��), 0, gRs!ҳ�汳��), 0)
            Call ShowPicture(IIf(IsNull(gRs!��������), 0, gRs!��������), 1)
            
            UsrPic(0).Tag = IIf(IsNull(gRs!ҳ�汳��), 0, gRs!ҳ�汳��)
            UsrPic(1).Tag = IIf(IsNull(gRs!��������), 0, gRs!��������)
                        
            cbo.ListIndex = FindCboIndex(cbo, IIf(IsNull(gRs!��������), 0, gRs!��������))
            txtEdit(0).Text = IIf(IsNull(gRs!����), "", gRs!����)
            txtEdit(2).Text = IIf(IsNull(gRs!����), "", gRs!����)
            
            txtEdit(1).Text = IIf(IsNull(gRs!�������), "", gRs!�������)
            
            mstr���� = txtEdit(0).Text
        End If
    End If
    
    Call GetTreeCode(mlngUpKey)
    
    cmdOK.Tag = ""
    
    txtEdit(2).Enabled = txt.Enabled
    
    DoEvents
    
    txtEdit(0).SetFocus
    Call SelAll(txtEdit(0))
    
    mblnFirst = False
End Sub

Private Sub Form_Load()
    mblnFirst = True
    mOK = False
        
    lblSize(0).Caption = ""
    lblSize(1).Caption = ""
                
    mvarSvrPicRange = ""
    mvarSvrPicType = ""
    
End Sub

Private Function CheckValid() As Boolean
    txtEdit(0).Text = Trim(txtEdit(0).Text)

    If txtTemp.MaxLength = 0 Then
        If Len(txtEdit(0).Text) = 0 Then
            MsgBox "���벻��Ϊ�ա�", vbExclamation, gstrSysName
            txtEdit(0).SetFocus
            Exit Function
        End If
    Else
        If Len(txtEdit(0).Text) < txtEdit(0).MaxLength Then
            MsgBox "����ĳ��Ȳ�����", vbExclamation, gstrSysName
            txtEdit(0).SetFocus
            Exit Function
        End If
    End If
    If Not IsNumeric(txtEdit(0).Text) Or InStr(txtEdit(0).Text, ",") > 0 Or InStr(txtEdit(0).Text, ".") > 0 Or InStr(txtEdit(0).Text, "-") > 0 Then
        MsgBox "����Ӧ��������ɡ�", vbExclamation, gstrSysName
        txtEdit(0).SetFocus
        Exit Function
    End If
    If Len(Trim(txt.Text)) = 0 Then
        MsgBox "���Ʋ���Ϊ�ա�", vbExclamation, gstrSysName
        txt.Text = ""
        txt.SetFocus
        Exit Function
    End If
    
    CheckValid = True
End Function


Private Function SaveData() As Boolean
    Dim lng��� As Long
    Dim LngStyle As Long
    Dim i As Long
        
    If cmdOK.Tag <> "" Then
        
        If CheckValid = False Then Exit Function
        
        For i = 0 To lst.ListCount - 1
            If lst.Selected(i) Then
                LngStyle = lst.ItemData(i)
                Exit For
            End If
        Next
        If mKey = 0 Then
            lng��� = NextValue("��ѯҳ��Ŀ¼", "ҳ�����")
            gstrSQL = "zl_��ѯҳ��Ŀ¼_insert(" & lng��� & ",'" & txt.Text & "',0," & LngStyle & "," & IIf(Val(UsrPic(1).Tag) = 0, "NULL", Val(UsrPic(1).Tag)) & "," & IIf(Val(UsrPic(0).Tag) = 0, "NULL", Val(UsrPic(0).Tag)) & "," & IIf(cbo.ItemData(cbo.ListIndex) = 0, "NULL", cbo.ItemData(cbo.ListIndex)) & "," & IIf(Val(txtEdit(3).Tag) = 0, "NULL", Val(txtEdit(3).Tag)) & ",1,'" & txtTemp.Text & txtEdit(0).Text & "','" & txtEdit(2).Text & "','" & txtEdit(1).Text & "')"
        Else
            lng��� = mKey
            gstrSQL = "zl_��ѯҳ��Ŀ¼_update(" & mKey & ",'" & txt.Text & "'," & LngStyle & "," & IIf(Val(UsrPic(1).Tag) = 0, "NULL", Val(UsrPic(1).Tag)) & "," & IIf(Val(UsrPic(0).Tag) = 0, "NULL", Val(UsrPic(0).Tag)) & "," & IIf(cbo.ItemData(cbo.ListIndex) = 0, "NULL", cbo.ItemData(cbo.ListIndex)) & "," & IIf(Val(txtEdit(3).Tag) = 0, "NULL", Val(txtEdit(3).Tag)) & ",'" & txtTemp.Text & txtEdit(0).Text & "','" & txtEdit(2).Text & "'," & Len(mstr����) + 1 & ",'" & txtEdit(1).Text & "')"
        End If
                        
        On Error GoTo errHand
        
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        
        
        Call frmDefQuery.RefreshPage(CStr(lng���))
        
    End If
    
    SaveData = True
    Exit Function
errHand:
    If ErrCenter() = -1 Then Resume
    
End Function

Private Sub Form_Unload(Cancel As Integer)
    Cancel = mblnFirst
    If Cancel Then Exit Sub
    
    Call MusicClose
    If cmdOK.Tag = "1" Then
        If MsgBox("��ѯҳ���Ѿ����ģ�ȷ�ϲ�������˳���", vbYesNo + vbDefaultButton2 + vbQuestion, gstrSysName) = vbNo Then Cancel = True
    End If
End Sub

Private Sub lst_ItemCheck(Item As Integer)
    Call SelectListItem(lst.ItemData(Item))
End Sub

Private Sub lst_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
        Exit Sub
    End If
End Sub

Private Sub txt_Change()
    cmdOK.Tag = "1"
End Sub

Private Sub SelectListItem(ByVal Key As Long)
    Dim i As Long
    
    For i = 0 To lst.ListCount - 1
        If lst.ItemData(i) = Key Then
            lst.Selected(i) = True
        Else
            lst.Selected(i) = False
        End If
    Next
End Sub

Private Sub ShowPicture(ByVal PicNo As Long, ByVal Index As Long)
    Dim rs As New ADODB.Recordset
    gstrSQL = "select ���,���,�߶�,���� from ��ѯͼƬԪ�� where ���=[1]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, PicNo)
    If rs.BOF = False Then
        Call UsrPic(Index).ShowPictureByFieldNew(rs!���, rs!��� * Screen.TwipsPerPixelX, rs!�߶� * Screen.TwipsPerPixelY, IIf(IsNull(rs!����), 0, rs!����))
        lblSize(Index).Caption = "���:" & Format(rs!��� * Screen.TwipsPerPixelX / 567, "0.0(����)") & " �߶�:" & Format(rs!�߶� * Screen.TwipsPerPixelY / 567, "0.0(����)")
    End If
    CloseRecord rs
End Sub

Private Sub txt_GotFocus()
    Call SelAll(txt)
    zlCommFun.OpenIme True
End Sub

Private Sub txt_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
        Exit Sub
    Else
        txtEdit(2).Text = zlCommFun.SpellCode(txt.Text)
    End If
    If CheckIsInclude(UCase(Chr(KeyAscii)), "'") = True Then KeyAscii = 0
End Sub

Private Sub txt_LostFocus()
    zlCommFun.OpenIme
End Sub

Private Sub txt_Validate(Cancel As Boolean)
    Cancel = Not StrIsValid(txt.Text, txt.MaxLength)
End Sub

Private Sub txtEdit_Change(Index As Integer)
    cmdOK.Tag = "1"
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    Call SelAll(txtEdit(Index))
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
        If Index = 3 Then SendKeys "{TAB}"
    Else
        If Index = 2 Then KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If Index = 3 And Chr(KeyAscii) = "*" Then Call cmdOpen_Click(2)
    End If
End Sub

Private Sub txtEdit_Validate(Index As Integer, Cancel As Boolean)
    Cancel = Not StrIsValid(txtEdit(Index).Text, txtEdit(Index).MaxLength)
End Sub

Private Sub txtTemp_Change()
    txtEdit(0).Width = txtTemp.Width - TextWidth(txtTemp.Text) - 120
    txtEdit(0).Left = txtTemp.Left + TextWidth(txtTemp.Text) + 60
End Sub
