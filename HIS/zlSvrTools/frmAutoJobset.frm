VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAutoJobset 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�Զ���ҵ����"
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7035
   ControlBox      =   0   'False
   Icon            =   "frmAutoJobset.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   7035
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.PictureBox pic���� 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   2295
      Left            =   960
      ScaleHeight     =   2265
      ScaleWidth      =   4050
      TabIndex        =   19
      Top             =   1170
      Visible         =   0   'False
      Width           =   4080
      Begin VB.Label lbl˵�� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���100��ϵͳ������һ���Զ����۵���ҵ"
         Height          =   180
         Index           =   2
         Left            =   450
         TabIndex        =   27
         Top             =   1740
         Width           =   3330
      End
      Begin VB.Label lbl˵�� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ZL100_USERJOB�Զ�����"
         ForeColor       =   &H00C00000&
         Height          =   180
         Index           =   1
         Left            =   450
         TabIndex        =   26
         Top             =   1980
         Width           =   1890
      End
      Begin VB.Label lbl˵�� 
         BackStyle       =   0  'Transparent
         Caption         =   "�������������岿�����û����룻�Է����������û�������Ҫϵͳ�š�"
         Height          =   345
         Index           =   0
         Left            =   480
         TabIndex        =   25
         Top             =   1005
         Width           =   3345
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   2
         Left            =   90
         TabIndex        =   24
         Top             =   1440
         Width           =   390
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "˵��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   1
         Left            =   90
         TabIndex        =   23
         Top             =   810
         Width           =   390
      End
      Begin VB.Label lbl�û� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "[ϵͳ��]        ����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   690
         TabIndex        =   22
         Top             =   450
         Width           =   2100
      End
      Begin VB.Label lbl�̶� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ZL        _USERJOB"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   480
         TabIndex        =   21
         Top             =   450
         Width           =   1890
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�û��Զ���ҵ��������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   0
         Left            =   90
         TabIndex        =   20
         Top             =   150
         Width           =   1950
      End
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "���²���"
      Height          =   350
      Left            =   5370
      TabIndex        =   28
      Top             =   1560
      Visible         =   0   'False
      Width           =   1100
   End
   Begin VB.CheckBox chk���� 
      Caption         =   "��������"
      Height          =   350
      Left            =   5370
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   1170
      Visible         =   0   'False
      Width           =   1100
   End
   Begin VB.TextBox txtJobName 
      Height          =   300
      Left            =   1320
      TabIndex        =   0
      Top             =   75
      Width           =   3810
   End
   Begin VB.TextBox txtJobComment 
      ForeColor       =   &H00808080&
      Height          =   1230
      Left            =   900
      Locked          =   -1  'True
      MaxLength       =   100
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   870
      Width           =   4215
   End
   Begin VB.CommandButton cmdWhat 
      Caption         =   "��"
      Enabled         =   0   'False
      Height          =   300
      Left            =   4830
      TabIndex        =   1
      Top             =   450
      Width           =   285
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   5370
      TabIndex        =   15
      Top             =   480
      Width           =   1100
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   5370
      TabIndex        =   14
      Top             =   60
      Width           =   1100
   End
   Begin VB.Frame fraPara 
      Caption         =   "ִ�в���"
      Height          =   840
      Left            =   900
      TabIndex        =   12
      Top             =   3690
      Width           =   4215
      Begin VB.TextBox txtPara 
         Height          =   300
         Index           =   0
         Left            =   1035
         TabIndex        =   6
         Top             =   315
         Width           =   2010
      End
      Begin VB.Label lblPara 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "�Ǽ�ʱ��"
         Height          =   180
         Index           =   0
         Left            =   285
         TabIndex        =   13
         Top             =   375
         Width           =   720
      End
   End
   Begin VB.Frame fraCycle 
      Caption         =   "ִ������"
      Height          =   1080
      Left            =   900
      TabIndex        =   9
      Top             =   2535
      Width           =   4215
      Begin MSComCtl2.DTPicker dtpStart 
         Height          =   300
         Left            =   2100
         TabIndex        =   4
         Top             =   645
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "HH:mm:ss"
         Format          =   106168323
         UpDown          =   -1  'True
         CurrentDate     =   37031.0416666667
      End
      Begin VB.ComboBox cboMonth 
         Height          =   300
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   645
         Width           =   900
      End
      Begin VB.ComboBox cboDay 
         Height          =   300
         Left            =   2370
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   645
         Width           =   1030
      End
      Begin VB.ComboBox cboWeek 
         Height          =   300
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   645
         Width           =   1030
      End
      Begin VB.ComboBox cboCycle 
         Height          =   300
         Left            =   2505
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   225
         Width           =   720
      End
      Begin VB.TextBox txtCycle 
         Height          =   300
         Left            =   1080
         TabIndex        =   5
         Top             =   225
         Width           =   1425
      End
      Begin VB.Label lblCycle 
         AutoSize        =   -1  'True
         Caption         =   "ѭ��ʱ��"
         Height          =   180
         Left            =   285
         TabIndex        =   11
         Top             =   270
         Width           =   720
      End
      Begin VB.Label lblStart 
         AutoSize        =   -1  'True
         Caption         =   "ִ��ʱ��"
         Height          =   180
         Left            =   285
         TabIndex        =   10
         Top             =   705
         Width           =   720
      End
   End
   Begin VB.CheckBox chkAutoJob 
      Caption         =   "����Ϊ��̨�Զ���ҵ(&A)"
      Height          =   210
      Left            =   900
      TabIndex        =   3
      Top             =   2190
      Width           =   2850
   End
   Begin VB.Label lblComment 
      AutoSize        =   -1  'True
      Caption         =   "˵��"
      Height          =   180
      Left            =   450
      TabIndex        =   17
      Top             =   900
      Width           =   360
   End
   Begin VB.Label lblJobWhat 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   1320
      TabIndex        =   7
      Top             =   450
      Width           =   3525
   End
   Begin VB.Label lblWhat 
      AutoSize        =   -1  'True
      Caption         =   "����"
      Height          =   180
      Left            =   900
      TabIndex        =   16
      Top             =   510
      Width           =   360
   End
   Begin VB.Image imgMain 
      Height          =   480
      Left            =   255
      Picture         =   "frmAutoJobset.frx":000C
      Top             =   195
      Width           =   480
   End
   Begin VB.Label lblSys 
      AutoSize        =   -1  'True
      Caption         =   "��ҵ"
      Height          =   180
      Left            =   900
      TabIndex        =   8
      Top             =   150
      Width           =   360
   End
   Begin VB.Menu mnuProcedures 
      Caption         =   "Procedure"
      Visible         =   0   'False
      Begin VB.Menu mnuWhat 
         Caption         =   "mnuWhat"
         Index           =   0
      End
   End
End
Attribute VB_Name = "frmAutoJobset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsTemp As New ADODB.Recordset
Dim strSQL As String
Dim intCount As Integer
Dim strOrder As String, strParas As String
Dim aryPara() As String
Private mdateNow As Date

Private Enum DateUnit
    DU_�� = 0
    DU_�� = 1
    DU_�� = 2
    DU_���� = 3
End Enum

Private Sub cboCycle_Click()
    Dim rsTmp As ADODB.Recordset
    Dim lngDay As Long
    Dim lngMonth As Long
    Dim lngMaxDay As Long
    
    Select Case cboCycle.ListIndex
    Case DU_��
        cboMonth.Visible = False
        cboWeek.Visible = False
        cboDay.Visible = False
        dtpStart.Width = 2145
        txtCycle.Width = 1425
        cboCycle.Left = txtCycle.Left + txtCycle.Width
        dtpStart.Left = txtCycle.Left
        
        If cboCycle.Text = cboCycle.Tag Then
            dtpStart.value = dtpStart.Tag
        Else
            dtpStart.value = "2001/5/20 1:00:00"
        End If
    Case DU_��
        cboMonth.Visible = False
        cboWeek.Visible = True
        cboDay.Visible = False
        dtpStart.Width = 1125
        txtCycle.Width = 1425
        cboCycle.Left = txtCycle.Left + txtCycle.Width
        cboWeek.Left = txtCycle.Left
        dtpStart.Left = cboWeek.Left + cboWeek.Width - 20
        
        If cboCycle.Text = cboCycle.Tag Then
            cboWeek.ListIndex = Weekday(CDate(dtpStart.Tag)) - 1
            dtpStart.value = dtpStart.Tag
        Else
            cboWeek.ListIndex = 1
            dtpStart.value = "2001/5/20 1:00:00"
        End If
    Case DU_��
        cboMonth.Visible = False
        cboWeek.Visible = False
        cboDay.Visible = True
        dtpStart.Width = 1125
        txtCycle.Width = 1425
        cboCycle.Left = txtCycle.Left + txtCycle.Width
        cboDay.Left = txtCycle.Left
        dtpStart.Left = cboDay.Left + cboDay.Width - 20
        
        If cboCycle.Text = cboCycle.Tag Then
            '��ȡָ�����������
            lngMaxDay = Right(DateSerial(Year(dtpStart.Tag), Month(dtpStart.Tag) + 1, 0), 2)
            lngDay = Format(dtpStart.Tag, "d")
            If lngDay <= 28 Then
                cboDay.Text = lngDay & "��"
            ElseIf lngDay = lngMaxDay Then
                cboDay.Text = "��ĩ"
            ElseIf lngDay = lngMaxDay - 1 Then
                cboDay.Text = "��ĩ-1"
            ElseIf lngDay = lngMaxDay - 2 Then
                cboDay.Text = "��ĩ-2"
            End If
            dtpStart.value = dtpStart.Tag
        Else
            cboDay.ListIndex = 0
            dtpStart.value = "2001/5/20 1:00:00"
        End If
    Case DU_����
        cboWeek.Visible = False
        cboMonth.Visible = True
        cboDay.Visible = True
        dtpStart.Width = 1125
        txtCycle.Width = 2310
        cboCycle.Left = txtCycle.Left + txtCycle.Width
        cboMonth.Left = txtCycle.Left
        cboDay.Left = cboMonth.Left + cboMonth.Width - 20
        dtpStart.Left = cboDay.Left + cboDay.Width - 20
        
        If cboCycle.Text = cboCycle.Tag Then
            '���ָ�����ǵڼ�����
            lngMonth = Format(dtpStart.Tag, "M") Mod 3 - 1
            If lngMonth = 0 Then
                cboMonth.Text = "��һ��"
            ElseIf lngMonth = 1 Then
                cboMonth.Text = "�ڶ���"
            Else
                lngMonth = 2
                cboMonth.Text = "������"
            End If
            '��ȡָ�����������
            lngMaxDay = Right(DateSerial(Year(CDate(dtpStart.Tag)), Month(CDate(dtpStart.Tag)) + 1, 0), 2)
            lngDay = Format(dtpStart.Tag, "d")
            If lngDay <= 28 Then
                cboDay.Text = lngDay & "��"
            ElseIf lngDay = lngMaxDay Then
                cboDay.Text = "��ĩ"
            ElseIf lngDay = lngMaxDay - 1 Then
                cboDay.Text = "��ĩ-1"
            ElseIf lngDay = lngMaxDay - 2 Then
                cboDay.Text = "��ĩ-2"
            End If
            dtpStart.value = dtpStart.Tag
        Else
            cboMonth.ListIndex = 0
            cboDay.ListIndex = 0
            dtpStart.value = "2001/5/20 1:00:00"
        End If
        
        '���뵱ǰ�����е�һ�µ��·�
        cboMonth.Tag = Format(mdateNow, "M") - lngMonth
    End Select
End Sub

Private Sub chk����_Click()
    pic����.Visible = chk����.value = 1
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim strExecuteTime As String, strQuarterly As String
    Dim rsTmp As ADODB.Recordset
    Dim lngMaxDay As Long
    Dim cnTools As ADODB.Connection
    
    If Trim(lblJobWhat.Caption) = "" Then
        MsgBox "δ������ҵ���ݣ�", vbExclamation, gstrSysName
        Exit Sub
    End If
    If Val(txtCycle.Text) = 0 Then
        MsgBox "δ��ȷ������ҵѭ��ʱ�䣡", vbExclamation, gstrSysName
        txtCycle.SetFocus: Exit Sub
    End If
    
    strParas = ""
    If fraPara.Visible Then
        For intCount = 0 To lblPara.UBound
            If lblPara(intCount).Visible = False Then Exit For
            If Trim(txtPara(intCount).Text) = "" Then
                MsgBox lblPara(intCount).Caption & " ����δָ��ֵ��", vbExclamation, gstrSysName
                Exit Sub
            End If
            strParas = strParas & ";" & lblPara(intCount).Caption & "," & txtPara(intCount).Text
        Next
    End If
    If strParas <> "" Then strParas = Mid(strParas, 2)
    
    '����ȡ����ִ��������Ϣת��Ϊ���������
    Select Case cboCycle.ListIndex
    Case DU_��
        strExecuteTime = Format(mdateNow, "yyyy-MM-dd") & " " & Format(dtpStart.value, "HH:mm:ss")
    Case DU_��
        strExecuteTime = Format(DateAdd("d", cboWeek.ListIndex + 1 - Weekday(mdateNow), mdateNow), "yyyy-MM-dd") & " " & Format(dtpStart.value, "HH:mm:ss")
    Case DU_��
        If cboDay.ListIndex <= 27 Then
            strExecuteTime = Format(mdateNow, "yyyy-MM") & "-" & Val(cboDay.Text) & " " & Format(dtpStart.value, "HH:mm:ss")
        Else
            lngMaxDay = Right(DateSerial(Year(mdateNow), Month(mdateNow) + 1, 0), 2)
            strExecuteTime = Format(mdateNow, "yyyy-MM") & "-" & lngMaxDay - (cboDay.ListCount - cboDay.ListIndex - 1) & " " & Format(dtpStart.value, "HH:mm:ss")
        End If
    Case DU_����
        If cboDay.ListIndex <= 27 Then
            strExecuteTime = Format(mdateNow, "yyyy") & "-" & cboMonth.Tag + cboMonth.ListIndex & "-" & Val(cboDay.Text) & " " & Format(dtpStart.value, "HH:mm:ss")
        Else
            strQuarterly = Format(mdateNow, "yyyy") & "-" & cboMonth.Tag + cboMonth.ListIndex & "-" & "01 11:11:11"
            lngMaxDay = Right(DateSerial(Year(CDate(strQuarterly)), Month(CDate(strQuarterly)) + 1, 0), 2)
            strExecuteTime = Format(mdateNow, "yyyy") & "-" & cboMonth.Tag + cboMonth.ListIndex & "-" & lngMaxDay - (cboDay.ListCount - cboDay.ListIndex - 1) & " " & Format(dtpStart.value, "HH:mm:ss")
        End If
    End Select
    
    If Tag = "ADD" Then
        Dim rsOut As New ADODB.Recordset
        'ȡZlAutoJob���к�
        Set rsOut = OpenCursor(gcnOracle, "ZLTOOLS.B_Runmana.Get_Job_number", Val(lblSys.Tag))
        If rsOut.RecordCount > 0 Then
            strOrder = Nvl(Val(rsOut.Fields(0)), 1)
        Else
            strOrder = 1
        End If
        strSQL = "insert into zlAutoJobs(ϵͳ,����,���,����,˵��,����,����,ִ��ʱ��,���ʱ��,ʱ�䵥λ)" & _
                " values (" & IIf(Val(lblSys.Tag) = 0, "Null", lblSys.Tag) & ",3," & Val(strOrder) & "," & _
                "       '" & txtJobName.Text & "'," & _
                "       '" & txtJobComment.Text & "'," & _
                "       '" & lblJobWhat.Caption & "'," & _
                "       '" & strParas & "'," & _
                "       to_date('" & strExecuteTime & "','yyyy-MM-dd HH24:MI:SS')," & _
                "       " & Val(txtCycle.Text) & _
                "       ,'" & cboCycle.Text & "')"
    Else
        strSQL = "update zlAutoJobs" & _
                " set ����='" & txtJobName.Text & "'," & _
                "     ˵��='" & txtJobComment.Text & "'," & _
                "     ����='" & lblJobWhat.Caption & "'," & _
                "     ����='" & strParas & "'," & _
                "     ִ��ʱ��=to_date('" & strExecuteTime & "','yyyy-MM-dd HH24:MI:SS')," & _
                "     ���ʱ��=" & Val(txtCycle.Text) & "," & _
                "     ʱ�䵥λ='" & cboCycle.Text & "'" & _
                " Where Nvl(ϵͳ,0)=" & Val(lblSys.Tag) & _
                "     and ����=" & Tag & _
                "     and ���=" & txtJobName.Tag
    End If
    err = 0
    On Error Resume Next
    gcnOracle.Execute strSQL
    If err <> 0 Then
        MsgBox "��ҵ���ñ���ʧ�ܣ��������������" & vbNewLine & err.Description, vbExclamation, gstrSysName
        Exit Sub
    End If
    If Tag = "ADD" Then
        '������Ҫ������־
        Call SaveAuditLog(1, "����", "�ڡ�" & Split(frmAutoJobs.cmbSystem.Text, " ")(0) & "������Զ���ҵ��" & txtJobName.Text & "��")
    Else
        '������Ҫ������־
        Call SaveAuditLog(2, "��������", "�޸ġ�" & Split(frmAutoJobs.cmbSystem.Text, " ")(0) & "���е��Զ���ҵ��" & txtJobName.Text & "��")
    End If
    err = 0
    If imgMain.Tag = "ZLTOOLS" Then
        Set cnTools = GetConnection("ZLTOOLS")
        If cnTools Is Nothing Then Exit Sub
    Else
        Set cnTools = gcnOracle
    End If
    If chkAutoJob.value = 1 Then
        If Tag = "ADD" Then                      '����ҵ
            strSQL = "zl" & "_JobSubmit(" & IIf(Val(lblSys.Tag) = 0, "Null", lblSys.Tag) & ",3," & Val(strOrder) & ")"
        ElseIf Val(chkAutoJob.Tag) = 0 Then      '�״�����Ϊ�Զ���ҵ
            strSQL = "zl" & "_JobSubmit(" & IIf(Val(lblSys.Tag) = 0, "Null", lblSys.Tag) & "," & Tag & "," & txtJobName.Tag & ")"
        Else                                        '�޸��Ѿ����õ���ҵ
            strSQL = "zl" & "_JobChange(" & IIf(Val(lblSys.Tag) = 0, "Null", lblSys.Tag) & "," & Tag & "," & txtJobName.Tag & ")"
        End If
        cnTools.Execute strSQL, , adCmdStoredProc
    Else
        If Val(chkAutoJob.Tag) <> 0 Then         'ȡ���Զ���ҵ
            strSQL = "zl" & "_JobRemove(" & IIf(Val(lblSys.Tag) = 0, "Null", lblSys.Tag) & "," & Tag & "," & txtJobName.Tag & ")"
            cnTools.Execute strSQL, , adCmdStoredProc
        End If
    End If
    If err <> 0 Then
        MsgBox "��Ȼ��ҵ���ñ��棬��δ�ܳɹ�����Ϊ�Զ���ҵ���������ݿ�ϵͳ��", vbExclamation, gstrSysName
    End If
    
    Unload Me
End Sub

Private Sub cmdUpdate_Click()
    Dim rsTemp As New ADODB.Recordset
On Error GoTo errHandle
    
    If MsgBox("�Ƿ�������ݹ鵵ת�ƴ����õ�ʱ����²�����", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub

    Set rsTemp = OpenCursor(gcnOracle, "ZLTOOLS.B_Runmana.Get_depict", Val(lblSys.Tag), Val(txtJobName.Tag))
    If rsTemp.RecordCount > 0 Then
        txtPara(0).Text = Val(IIf(IsNull(rsTemp.Fields(0)), "150", rsTemp.Fields(0)))
    Else
        txtPara(0).Text = 150
    End If
    Exit Sub
errHandle:
    MsgBox err.Description, vbCritical, Me.Caption
End Sub

Private Sub cmdWhat_Click()
   Dim cnTools As ADODB.Connection
On Error GoTo errHandle
    If Val(cmdWhat.Tag) = 0 Then
        If imgMain.Tag = "ZLTOOLS" Then
            Set cnTools = GetConnection("ZLTOOLS")
            If cnTools Is Nothing Then Exit Sub
        Else
            Set cnTools = gcnOracle
        End If
        Set rsTemp = cnTools.Execute("SELECT Object_Name  From All_Objects " & vbNewLine & _
                                      "WHERE Object_Type = 'PROCEDURE' AND Object_Name LIKE 'ZL" & CStr(IIf(Val(lblSys.Tag) = 0, "", lblSys.Tag)) & "_USERJOB%' " & vbNewLine & _
                                      " AND Status = 'VALID' AND Owner = '" & CStr(imgMain.Tag) & "'")
        With rsTemp
            Do While Not .EOF
                If .AbsolutePosition - 1 > mnuWhat.UBound Then Load mnuWhat(.AbsolutePosition - 1)
                mnuWhat(.AbsolutePosition - 1).Caption = .Fields(0).value
                mnuWhat(.AbsolutePosition - 1).Visible = True
                .MoveNext
            Loop
            cmdWhat.Tag = .RecordCount
        End With
    End If
    If Val(cmdWhat.Tag) > 0 Then
        PopupMenu mnuProcedures, 2
    Else
        MsgBox "û�п�ѡ�Ĵ洢����", vbExclamation, gstrSysName
    End If
    Exit Sub
errHandle:
    MsgBox err.Description, vbCritical, Me.Caption
End Sub

Private Sub Form_Activate()
    Dim i As Long
    
    If frmAutoJobset.Tag = "2" Then cmdUpdate.Visible = True
    cboCycle.Clear
    cboCycle.addItem "��"
    cboCycle.addItem "��"
    cboCycle.addItem "��"
    cboCycle.addItem "����"
    cboWeek.Clear
    cboWeek.addItem "������"
    cboWeek.addItem "����һ"
    cboWeek.addItem "���ڶ�"
    cboWeek.addItem "������"
    cboWeek.addItem "������"
    cboWeek.addItem "������"
    cboWeek.addItem "������"
    cboMonth.Clear
    cboMonth.addItem "��һ��"
    cboMonth.addItem "�ڶ���"
    cboMonth.addItem "������"
    cboDay.Clear
    For i = 1 To 28
        cboDay.addItem i & "��"
    Next
    cboDay.addItem "��ĩ-2"
    cboDay.addItem "��ĩ-1"
    cboDay.addItem "��ĩ"
    
    '����ǰ���ݿ�ʱ��������
    mdateNow = CurrentDate()
    
    cboCycle.Text = IIf(cboCycle.Tag = "", "��", cboCycle.Tag)
End Sub

Private Sub mnuWhat_Click(Index As Integer)
    On Error GoTo errHandle
    lblJobWhat.Caption = mnuWhat(Index).Caption
    With rsTemp
        If gblnDBA Then
            strSQL = "select rtrim(ltrim(upper(text))) from dba_source where name='" & mnuWhat(Index).Caption & "' and OWNER='" & imgMain.Tag & "'"
        Else
            strSQL = "select rtrim(ltrim(upper(text))) from user_source where name='" & mnuWhat(Index).Caption & "'"
        End If
        If .State = adStateOpen Then .Close
        .Open strSQL, gcnOracle, adOpenKeyset
        strSQL = ""
        Do While Not .EOF
            strSQL = strSQL & " " & Replace(Replace(Replace(Replace(Trim(.Fields(0).value), vbCrLf, " "), vbCr, " "), vbLf, " "), vbTab, " ")
            If InStr(1, strSQL, " AS ") > 0 Then Exit Do
            If InStr(1, strSQL, " IS ") > 0 Then Exit Do
            If InStr(1, strSQL, ")AS ") > 0 Then Exit Do
            If InStr(1, strSQL, ")IS ") > 0 Then Exit Do
            If Right(strSQL, 3) = " AS" Then Exit Do
            If Right(strSQL, 3) = " IS" Then Exit Do
            If Right(strSQL, 3) = ")AS" Then Exit Do
            If Right(strSQL, 3) = ")IS" Then Exit Do
            .MoveNext
        Loop
        strSQL = Replace(Replace(Replace(Replace(strSQL, vbCrLf, " "), vbCr, " "), vbLf, " "), vbTab, " ")
        If InStr(1, strSQL, "(") > 0 Then
            strSQL = Mid(strSQL, InStr(1, strSQL, "(") + 1)
            strSQL = Left(strSQL, InStr(1, strSQL, ")") - 1)
        Else
            strSQL = ""
        End If
        
        For intCount = 0 To lblPara.UBound
            lblPara(intCount).Visible = False
            txtPara(intCount).Visible = False
        Next
    
        If strSQL = "" Then
            Height = fraCycle.Top + fraCycle.Height + 600
            fraPara.Visible = False
        Else
            fraPara.Visible = True
            aryPara = Split(strSQL, ",")
            For intCount = 0 To UBound(aryPara)
                aryPara(intCount) = Trim(aryPara(intCount))
                If intCount > lblPara.UBound Then Load lblPara(intCount)
                If intCount > txtPara.UBound Then Load txtPara(intCount)
                lblPara(intCount).Top = intCount * 400 + 375
                txtPara(intCount).Top = intCount * 400 + 315
                lblPara(intCount).Left = txtPara(0).Left - lblPara(intCount).Width - 45
                txtPara(intCount).Left = txtPara(0).Left
                lblPara(intCount).Caption = Left(aryPara(intCount), InStr(1, aryPara(intCount), " ") - 1)
                txtPara(intCount).Text = ""
                lblPara(intCount).Visible = True
                txtPara(intCount).Visible = True
            Next
            fraPara.Height = (UBound(aryPara) + 1) * 400 + 375
            Height = fraPara.Top + fraPara.Height + 600
        End If
    
    End With
    Exit Sub
errHandle:
    MsgBox err.Description, vbCritical, Me.Caption
End Sub

Private Sub txtCycle_KeyPress(KeyAscii As Integer)
    If Not (InStr("0123456789", Chr(KeyAscii)) > 0 Or KeyAscii = 8) Then
        KeyAscii = 0
    End If
End Sub
