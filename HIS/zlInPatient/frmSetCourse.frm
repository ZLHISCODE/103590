VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSetCourse 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��������"
   ClientHeight    =   4395
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5175
   ClipControls    =   0   'False
   Icon            =   "frmSetCourse.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   5175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Visible         =   0   'False
   Begin VB.CommandButton cmdPrintSet 
      Caption         =   "Ӥ�������ӡ����"
      Height          =   345
      Index           =   1
      Left            =   120
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   3600
      Width           =   2055
   End
   Begin VB.CommandButton cmdPrintSet 
      Caption         =   "���������ӡ����"
      Height          =   345
      Index           =   0
      Left            =   120
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   3120
      Width           =   2055
   End
   Begin VB.CheckBox chkInTime 
      Caption         =   "��Ժ��סʱ�����޸���Ժʱ��"
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   2715
      Width           =   2775
   End
   Begin VB.TextBox txtOutTime 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   690
      MaxLength       =   3
      TabIndex        =   7
      Text            =   "30"
      Top             =   2310
      Width           =   525
   End
   Begin VB.TextBox txtInTime 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   690
      MaxLength       =   3
      TabIndex        =   4
      Text            =   "3"
      Top             =   1965
      Width           =   525
   End
   Begin VB.Frame fra����ƹ��� 
      Caption         =   "��ʾ���¿��ҵĴ���ס����"
      Height          =   1875
      Left            =   120
      TabIndex        =   2
      Top             =   75
      Width           =   4935
      Begin VB.ListBox lstDepartments 
         Height          =   1530
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   3
         ToolTipText     =   "Ctrl+Aȫѡ,Ctrl+Cȫ��"
         Top             =   240
         Width           =   4665
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3900
      TabIndex        =   1
      Top             =   3900
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   2730
      TabIndex        =   0
      Top             =   3900
      Width           =   1100
   End
   Begin MSComCtl2.UpDown UDInTime 
      Height          =   300
      Left            =   1215
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1965
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      Value           =   30
      BuddyControl    =   "txtInTime"
      BuddyDispid     =   196623
      OrigLeft        =   2340
      OrigTop         =   210
      OrigRight       =   2580
      OrigBottom      =   450
      Max             =   365
      Min             =   1
      SyncBuddy       =   -1  'True
      BuddyProperty   =   0
      Enabled         =   -1  'True
   End
   Begin MSComCtl2.UpDown UDOutTime 
      Height          =   300
      Left            =   1215
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2310
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      Value           =   30
      BuddyControl    =   "txtOutTime"
      BuddyDispid     =   196622
      OrigLeft        =   1215
      OrigTop         =   2310
      OrigRight       =   1455
      OrigBottom      =   2625
      Max             =   365
      Min             =   1
      SyncBuddy       =   -1  'True
      BuddyProperty   =   0
      Enabled         =   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��ʾ��          �����ڵĳ�Ժ����"
      Height          =   180
      Left            =   120
      TabIndex        =   10
      Top             =   2370
      Width           =   2880
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��ʾ��          �����ڵǼ���Ժ�Ĳ���"
      Height          =   180
      Left            =   120
      TabIndex        =   6
      Top             =   2025
      Width           =   3240
   End
End
Attribute VB_Name = "frmSetCourse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������
Public mstrPrivs As String    'Ȩ�޴�
Public mlngModul As Long      'ģ���

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name
End Sub

Private Sub cmdOK_Click()
    Dim i As Long, strpar As String
    Dim blnSelAll As Boolean


    If txtInTime.Enabled Then
        If Trim(txtInTime.Text) = "" Then
            MsgBox "��������Ҫ��ʾ����Ժʱ�䷶Χ��", vbInformation, gstrSysName
            txtOutTime.SetFocus: Exit Sub
        Else
            zlDatabase.SetPara "��Ժ����", Val(txtInTime.Text), glngSys, mlngModul, IIf(txtInTime.Enabled = True, True, False)
        End If
    End If
    
    If txtOutTime.Enabled Then
        If Trim(txtOutTime.Text) = "" Then
            MsgBox "��������Ҫ��ʾ�ĳ�Ժʱ�䷶Χ��", vbInformation, gstrSysName
            txtOutTime.SetFocus: Exit Sub
        Else
            zlDatabase.SetPara "��Ժ����", Val(txtOutTime.Text), glngSys, mlngModul, IIf(txtOutTime.Enabled = True, True, False)
        End If
    End If
    
    If fra����ƹ���.Enabled Then
        For i = lstDepartments.ListCount - 1 To 0 Step -1
            If lstDepartments.Selected(i) Then
                strpar = strpar & lstDepartments.ItemData(i) & ","
            End If
        Next
        If strpar <> "" Then
            strpar = Left(strpar, Len(strpar) - 1)
            If lstDepartments.ListCount = UBound(Split(strpar, ",")) + 1 Then strpar = "" 'ȫѡ�����޿�������
        End If
        zlDatabase.SetPara "����Ʋ��˿���", strpar, glngSys, mlngModul, IIf(fra����ƹ���.Enabled = True, True, False)
    End If

    '����42701 by ljf
    zlDatabase.SetPara "�����޸���Ժʱ��", chkInTime, glngSys, mlngModul, IIf(chkInTime.Enabled = True, True, False)
    Call InitLocPar(mlngModul)
    gblnOK = True
    Unload Me
End Sub

Private Sub cmdPrintSet_Click(Index As Integer)
    Select Case Index
    
    Case 0
        Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1132_4", Me)
    Case 1
        Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1132_3", Me)
    End Select
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long
    
    If KeyCode = 13 Then cmdOK_Click
    If Shift = vbCtrlMask And fra����ƹ���.Enabled Then
        If KeyCode = vbKeyA Then
            For i = 0 To lstDepartments.ListCount - 1
                lstDepartments.Selected(i) = True
            Next
        ElseIf KeyCode = vbKeyC Then
            For i = 0 To lstDepartments.ListCount - 1
                lstDepartments.Selected(i) = False
            Next
        End If
    End If
    
End Sub

Private Sub Form_Load()
    Dim i As Long, strpar As String
    Dim rsTmp As ADODB.Recordset
    
    gblnOK = False
    
    '����Ʋ��˿���
    Set rsTmp = GetDepts("�ٴ�", "1,2,3")
    Do While Not rsTmp.EOF
        lstDepartments.AddItem rsTmp!���� & "-" & rsTmp!����
        lstDepartments.ItemData(lstDepartments.NewIndex) = rsTmp!ID
        rsTmp.MoveNext
    Loop
    strpar = zlDatabase.GetPara("����Ʋ��˿���", glngSys, mlngModul, "", Array(fra����ƹ���), InStr(mstrPrivs, "��������") > 0)
    If strpar = "" Then
        For i = 0 To lstDepartments.ListCount - 1
            lstDepartments.Selected(i) = True
        Next
    Else
        For i = 0 To lstDepartments.ListCount - 1
            If InStr("," & strpar & ",", "," & lstDepartments.ItemData(i) & ",") > 0 Then lstDepartments.Selected(i) = True
        Next
    End If
    If lstDepartments.ListCount > 0 Then lstDepartments.TopIndex = 0: lstDepartments.ListIndex = 0
    
    txtInTime.Text = Val(zlDatabase.GetPara("��Ժ����", glngSys, mlngModul, "3", Array(txtInTime), InStr(mstrPrivs, "��������") > 0))
    txtOutTime.Text = Val(zlDatabase.GetPara("��Ժ����", glngSys, mlngModul, "30", Array(txtOutTime), InStr(mstrPrivs, "��������") > 0))
    
    chkInTime.Value = IIf(zlDatabase.GetPara("�����޸���Ժʱ��", glngSys, mlngModul, , Array(chkInTime), InStr(mstrPrivs, "��������") > 0) = "1", 1, 0)
End Sub

Private Sub txtInTime_GotFocus()
    zlControl.TxtSelAll txtInTime
End Sub

Private Sub txtInTime_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtOutTime_GotFocus()
    zlControl.TxtSelAll txtOutTime
End Sub

Private Sub txtOutTime_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub
