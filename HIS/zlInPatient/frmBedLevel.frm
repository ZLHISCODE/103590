VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "Msmask32.ocx"
Begin VB.Form frmBedLevel 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���Ĵ�λ�ȼ�"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5715
   Icon            =   "frmBedLevel.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   5715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   435
      TabIndex        =   11
      Top             =   2445
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   2295
      Left            =   120
      TabIndex        =   12
      Top             =   15
      Width           =   5460
      Begin VB.TextBox txt���� 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   3420
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   630
         Width           =   1830
      End
      Begin VB.TextBox txt���� 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   975
         Locked          =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   240
         Width           =   1515
      End
      Begin VB.TextBox txt�Ա� 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   3210
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   240
         Width           =   675
      End
      Begin VB.TextBox txt���� 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   240
         Width           =   690
      End
      Begin VB.TextBox txtסԺ�� 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   975
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   630
         Width           =   1515
      End
      Begin VB.TextBox txt���� 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   975
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1005
         Width           =   1515
      End
      Begin VB.ComboBox cboNew 
         Height          =   300
         Left            =   975
         TabIndex        =   8
         Text            =   "cboNew"
         Top             =   1770
         Width           =   4260
      End
      Begin VB.TextBox txtPre 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   975
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1395
         Width           =   4260
      End
      Begin MSMask.MaskEdBox txtDate 
         Height          =   300
         Left            =   3420
         TabIndex        =   6
         Top             =   1005
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   529
         _Version        =   393216
         AutoTab         =   -1  'True
         MaxLength       =   19
         Format          =   "yyyy-MM-dd HH:mm:ss"
         Mask            =   "####-##-## ##:##:##"
         PromptChar      =   "_"
      End
      Begin VB.ComboBox cboLevel 
         Height          =   300
         Left            =   1095
         Style           =   2  'Dropdown List
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   1770
         Visible         =   0   'False
         Width           =   2460
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ǰ����"
         Height          =   180
         Left            =   2640
         TabIndex        =   21
         Top             =   690
         Width           =   720
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "סԺ��"
         Height          =   180
         Left            =   375
         TabIndex        =   20
         Top             =   690
         Width           =   540
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   540
         TabIndex        =   19
         Top             =   300
         Width           =   360
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�Ա�"
         Height          =   180
         Left            =   2760
         TabIndex        =   18
         Top             =   300
         Width           =   360
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   4125
         TabIndex        =   17
         Top             =   300
         Width           =   360
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�䶯��λ"
         Height          =   180
         Left            =   195
         TabIndex        =   16
         Top             =   1065
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�µȼ�"
         Height          =   180
         Left            =   375
         TabIndex        =   14
         Top             =   1830
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ԭ�ȼ�"
         Height          =   180
         Left            =   375
         TabIndex        =   13
         Top             =   1455
         Width           =   540
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Чʱ��"
         Height          =   180
         Left            =   2640
         TabIndex        =   15
         Top             =   1065
         Width           =   720
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4155
      TabIndex        =   10
      Top             =   2445
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   2895
      TabIndex        =   9
      Top             =   2445
      Width           =   1100
   End
End
Attribute VB_Name = "frmBedLevel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������
Public mlng����ID As Long
Public mlng��ҳID As Long
Public mstr���� As String

Private mrsPatiInfo As ADODB.Recordset
Private mrsBedLevel As ADODB.Recordset

Private Sub cboNew_GotFocus()
    zlControl.TxtSelAll cboNew
End Sub

Private Sub cboNew_KeyPress(KeyAscii As Integer)
    '69273:������,2014-01-03,���ٶ�λ��λ�ȼ�
    Dim lngIdx As Long
    Dim i As Long, iCount As Integer
    Dim strText As String, strResult As String, strFilter As String
    Dim intInputType As Integer '0-�������ȫ����,1-�������ȫ��ĸ,2-����
    Dim strCompents As String 'ƥ�䴮
    Dim rsTemp As ADODB.Recordset
    
    If KeyAscii = 13 Then
        If cboNew.Locked Then
            Call zlCommFun.PressKey(vbKeyTab)
            Exit Sub
        End If
        strText = UCase(cboNew.Text)
        If cboNew.ListIndex <> -1 Then
            '�����б�ʱ,�����ı�������������
            If strText <> cboNew.List(cboNew.ListIndex) Then Call cbo.SetIndex(cboNew.hWnd, -1)
        End If
        If strText = "" Then
            cboNew.ListIndex = -1
        ElseIf cboNew.ListIndex = -1 Then
            strFilter = ""
            '�ȸ��Ƽ�¼��
            Set rsTemp = zlDatabase.zlCopyDataStructure(mrsBedLevel)
            strCompents = Replace(gstrLike, "%", "*") & strText & "*"
            
            If IsNumeric(strText) Then
                intInputType = 0
            ElseIf zlCommFun.IsCharAlpha(strText) Then
                intInputType = 1
            Else
                intInputType = 2
            End If
            
            mrsBedLevel.Filter = strFilter: iCount = 0
            With mrsBedLevel
                If .RecordCount <> 0 Then .MoveFirst
                Do While Not mrsBedLevel.EOF
                    Select Case intInputType
                    Case 0  '�������ȫ����
                        '������������,��Ҫ���:
                        '1.�������ֵ���,��Ҫ������:12 ƥ��000012�������
                        '2.���������,����Ϊ�Ǳ���,ֻ����ƥ��,��������12ƥ��00001201��120001��
                        
                        '��Ҫ�Ǽ�����������������ȫ��ͬ,��ֱ�ӾͶ�λ��������
                        If Nvl(!����) = strText Then strResult = Nvl(!����): iCount = 0: Exit Do
                        
                        '1.�������ֵ���,��Ҫ������:12 ƥ��000012�������,��Ϊ��������кܶ�:��0012,012,000012��.���������ڴ������,��Ҫ����ѡ������ѡ��
                        If Val(Nvl(!����)) = Val(strText) Then
                            If iCount = 0 Then strResult = Nvl(!����)
                            iCount = iCount + 1
                        End If
                        
                        '2.���������,����Ϊ�Ǳ���,ֻ����ƥ��,��������12ƥ��00001201��120001��
                         If Val(Nvl(!����)) Like strText & "*" Then
                            If isCheckBedLevelExists(Nvl(!����)) Then Call zlDatabase.zlInsertCurrRowData(mrsBedLevel, rsTemp)
                         End If
                    Case 1  '�������ȫ��ĸ
                        '����:
                        ' 1.����ļ������,��ֱ�Ӷ�λ
                        ' 2.���ݲ�����ƥ����ͬ����
                        
                        '1.����ļ������,��ֱ�Ӷ�λ
                        If Trim(Nvl(!����)) = strText Then
                            If iCount = 0 Then strResult = Nvl(!����)   '���ܴ��ڶ����ͬ�Ķ��
                            iCount = iCount + 1
                        End If
                        
                        '2.���ݲ�����ƥ����ͬ����
                        If Trim(Nvl(!����)) Like strCompents Then
                            If isCheckBedLevelExists(Nvl(!����)) Then Call zlDatabase.zlInsertCurrRowData(mrsBedLevel, rsTemp)
                        End If
                    Case Else  ' 2-����
                        '����:���ܴ��ں��ֵ����,����������N001���������ZYK01�������
                        '1.����\�������,ֱ�Ӷ�λ
                        '2.������������� ���ݲ�����ƥ����(������ֻ����ƥ��)
                        
                        '1.����\�������,ֱ�Ӷ�λ
                        If Trim(!����) = strText Or Trim(!����) = strText Or Trim(!����) = strText Then
                            If iCount = 0 Then strResult = Nvl(!����)   '���ܴ��ڶ����ͬ�Ķ��
                            iCount = iCount + 1
                        End If
                        
                        '2.������������� ���ݲ�����ƥ����(������ֻ����ƥ��)
                        If Trim(!����) Like strText & "*" Or Trim(Nvl(!����)) Like strCompents Or Trim(Nvl(!����)) Like strCompents Then
                            If isCheckBedLevelExists(Nvl(!����)) Then Call zlDatabase.zlInsertCurrRowData(mrsBedLevel, rsTemp)
                        End If
                    End Select
                    mrsBedLevel.MoveNext
                Loop
            End With
            If iCount > 1 Then strResult = ""
            If strResult = "" And rsTemp.RecordCount = 1 Then strResult = Nvl(rsTemp!����)
            'ֱ�Ӷ�λ
            If strResult <> "" Then
                rsTemp.Close: Set rsTemp = Nothing
                If isCheckBedLevelExists(strResult, True) Then cboNew.SetFocus:  zlCommFun.PressKey vbKeyTab
                Exit Sub
            End If
            
            '��Ҫ����Ƿ��ж������������ļ�¼
            If rsTemp.RecordCount <> 0 Then
                '�Ȱ�ĳ�ַ�ʽ��������
                rsTemp.Sort = "����,����"
                '����ѡ����
                Dim rsReturn As ADODB.Recordset
                If zlDatabase.zlShowListSelect(Me, glngSys, 1130, cboNew, rsTemp, True, "", "", rsReturn) Then
                    If Not rsReturn Is Nothing Then
                        If rsReturn.RecordCount <> 0 Then
                            '���ж�λ
                            If isCheckBedLevelExists(Nvl(rsReturn!����), True) Then
                                cboNew.SetFocus
                                zlCommFun.PressKey vbKeyTab
                                Exit Sub
                            End If
                        End If
                    End If
                Else
                    cboNew.SetFocus
                    Exit Sub
                End If
            Else
                'δ�ҵ�
                rsTemp.Close: Set rsTemp = Nothing
                KeyAscii = 0: cboNew.ListIndex = -1: zlControl.TxtSelAll cboNew: Exit Sub
            End If
            rsTemp.Close: Set rsTemp = Nothing
        End If
        
        If cboNew.ListIndex = -1 Then
            cboNew.Text = ""
            Exit Sub
        Else
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    End If
End Sub

Private Function isCheckBedLevelExists(ByVal str���� As String, Optional blnLocateItem As Boolean = False, Optional ByVal blnLevel As Boolean = True) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��������Ƿ��ڴ�λ�ȼ������б���
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    If blnLevel = True Then
        For i = 0 To cboLevel.ListCount - 1
            If cboLevel.List(i) = str���� Then
                If blnLocateItem Then cboNew.ListIndex = i
                isCheckBedLevelExists = True
                Exit Function
            End If
        Next
    Else
        For i = 0 To cboNew.ListCount - 1
            If cboNew.List(i) = str���� Then
                If blnLocateItem Then cboNew.ListIndex = i
                isCheckBedLevelExists = True
                Exit Function
            End If
        Next
    End If
End Function

Private Sub cboNew_Validate(Cancel As Boolean)
    If isCheckBedLevelExists(cboNew.Text, True, False) = False Then
        cboNew.Text = ""
        cboNew.ListIndex = -1
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("'", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim rsTmp As ADODB.Recordset
    Dim i As Integer, lngLevel As Long
    
    
    gblnOK = False
    Set mrsPatiInfo = GetPatiInfo(mlng����ID, mlng��ҳID)
    Set rsTmp = GetPatiBeds(mlng����ID, mstr����)
    
    With mrsPatiInfo
       txt����.Text = !����
       txt�Ա�.Text = "" & !�Ա�
       txt����.Text = "" & !����
       txtסԺ��.Text = "" & !סԺ��
       txt����.Text = "" & !��ǰ����
       txt����.Text = mstr����
       txtPre.Text = rsTmp!��λ�ȼ�
       lngLevel = Val("" & rsTmp!��λ�ȼ�id)
    End With

    txtDate.Text = Format(zlDatabase.Currentdate(), "yyyy-MM-dd HH:mm:ss")
    
    On Error GoTo errH
    '69273:������,2014-01-03,�ṩ��λ�ǼǵĿ��ٲ���
    gstrSQL = "Select ID,����,����,zlspellcode(����,20) ���� From �շ���ĿĿ¼ Where ���='J' And (����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or ����ʱ�� is NULL) And ID<>[1] Order by ����"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngLevel)
    cboNew.Clear
    cboLevel.Clear: cboLevel.Visible = False
    Set mrsBedLevel = rsTmp.Clone
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            cboNew.AddItem rsTmp!���� & "-" & rsTmp!����
            cboNew.ItemData(i - 1) = rsTmp!ID
            cboLevel.AddItem rsTmp!����
            cboLevel.ItemData(i - 1) = rsTmp!ID
            rsTmp.MoveNext
        Next
        cboNew.ListIndex = 0
    Else
        MsgBox "���ܶ�ȡ��λ�ȼ�����,���ȵ���λ�ȼ����������ã�", vbInformation, gstrSysName
        Unload Me: Exit Sub
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txtDate_GotFocus()
    zlControl.TxtSelAll txtDate
End Sub

Private Sub txtDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If IsDate(txtDate.Text) Then Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub txtDate_LostFocus()
    If Not IsDate(txtDate.Text) Then txtDate.SetFocus
End Sub

Private Sub cmdOK_Click()
    Dim rsTmp As New ADODB.Recordset
    Dim dMax As Date, strSQL As String
    Dim Curdate As Date
    
    If cboNew.ListIndex = -1 Then
        MsgBox "��ѡ���µĴ�λ�ȼ���", vbInformation, gstrSysName
        cboNew.SetFocus: Exit Sub
    End If
    If Not IsDate(txtDate.Text) Then
        MsgBox "������Ϸ�����Чʱ�䣡", vbInformation, gstrSysName
        txtDate.SetFocus: Exit Sub
    End If
    
    dMax = GetMaxDate(mlng����ID, mlng��ҳID)
    If CDate(txtDate.Text) <= dMax Then
        MsgBox "��Чʱ�������ڸò����ϴα䶯ʱ�� " & Format(dMax, "yyyy-MM-dd HH:mm:ss") & " ��", vbInformation, gstrSysName
        txtDate.SetFocus: Exit Sub
    End If
    
    'ʱ�䲻�ܳ�����ǰʱ��̫��(һ����)
    Curdate = zlDatabase.Currentdate
    If CDate(txtDate.Text) > Curdate Then
        If CDate(txtDate.Text) - Curdate > 30 Then
            MsgBox "��Чʱ��ȵ�ǰʱ���ù���,���飡", vbInformation, gstrSysName
            txtDate.SetFocus: Exit Sub
        End If
        If MsgBox("��Чʱ������˵�ǰϵͳʱ��,Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            txtDate.SetFocus: Exit Sub
        End If
    End If
        
    strSQL = "zl_���˱䶯��¼_BedLevel(" & mlng����ID & "," & mlng��ҳID & ",'" & txt����.Text & "'," & _
        cboNew.ItemData(cboNew.ListIndex) & ",To_Date('" & txtDate.Text & "','YYYY-MM-DD HH24:MI:SS')," & _
        "'" & UserInfo.��� & "','" & UserInfo.���� & "')"
    
    On Error GoTo errH
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    gblnOK = True
    Unload Me
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Function ShowMe(frmParent As Object, ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal str���� As String) As Boolean
    mlng����ID = lng����ID
    mlng��ҳID = lng��ҳID
    mstr���� = str����
    
    Me.Show 1, frmParent
    ShowMe = gblnOK
End Function
