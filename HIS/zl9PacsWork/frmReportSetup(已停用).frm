VERSION 5.00
Begin VB.Form frmReportSetup 
   BorderStyle     =   0  'None
   ClientHeight    =   6795
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7995
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   6795
   ScaleWidth      =   7995
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Frame fraReportSetup 
      BorderStyle     =   0  'None
      Height          =   6615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7935
      Begin VB.Frame fraEditorSetUp 
         Caption         =   "�����ĵ��༭������"
         Height          =   5535
         Left            =   120
         TabIndex        =   34
         Top             =   960
         Width           =   7695
         Begin VB.Frame Frame8 
            Caption         =   "�鿴��ʷ����"
            Height          =   1215
            Left            =   240
            TabIndex        =   35
            Top             =   480
            Width           =   7215
            Begin VB.OptionButton optHistoryReportEditor 
               Caption         =   "PACS����༭��"
               Height          =   255
               Index           =   1
               Left            =   4080
               TabIndex        =   37
               Top             =   600
               Width           =   1695
            End
            Begin VB.OptionButton optHistoryReportEditor 
               Caption         =   "���Ӳ����༭��"
               Height          =   255
               Index           =   0
               Left            =   360
               TabIndex        =   36
               Top             =   600
               Value           =   -1  'True
               Width           =   1695
            End
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "����༭��"
         Height          =   615
         Left            =   120
         TabIndex        =   29
         Top             =   120
         Width           =   7730
         Begin VB.OptionButton optReportEditor 
            Caption         =   "�����ĵ��༭��"
            Height          =   255
            Index           =   2
            Left            =   5640
            TabIndex        =   33
            Top             =   240
            Width           =   1575
         End
         Begin VB.OptionButton optReportEditor 
            Caption         =   "���Ӳ����༭��"
            Height          =   255
            Index           =   0
            Left            =   600
            TabIndex        =   31
            Top             =   240
            Width           =   1575
         End
         Begin VB.OptionButton optReportEditor 
            Caption         =   "PACS����༭��"
            Height          =   255
            Index           =   1
            Left            =   3000
            TabIndex        =   30
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "��������"
         Height          =   4575
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   7695
         Begin VB.CheckBox chkUntreadPrinted 
            Caption         =   "��˴�ӡ���������"
            Height          =   180
            Left            =   600
            TabIndex        =   32
            Top             =   1800
            Width           =   2175
         End
         Begin VB.CheckBox chkSpecialContent 
            Caption         =   "��ʾר�Ʊ������ݣ�"
            Height          =   180
            Left            =   600
            TabIndex        =   28
            Top             =   2280
            Width           =   2055
         End
         Begin VB.ComboBox cboSpecialContent 
            Height          =   300
            Left            =   600
            TabIndex        =   27
            Text            =   "Combo1"
            Top             =   2760
            Width           =   6495
         End
         Begin VB.CheckBox chkExitAfterPrint 
            Caption         =   "��ӡ���˳�"
            Height          =   180
            Left            =   600
            TabIndex        =   26
            Top             =   1320
            Width           =   1335
         End
         Begin VB.Frame Frame3 
            Caption         =   "�����ı�������"
            Height          =   1335
            Left            =   3960
            TabIndex        =   19
            Top             =   960
            Width           =   3255
            Begin VB.TextBox txtAdvice 
               Height          =   270
               Left            =   1560
               TabIndex        =   22
               Top             =   960
               Width           =   1335
            End
            Begin VB.TextBox txtResult 
               Height          =   270
               Left            =   1560
               TabIndex        =   21
               Top             =   600
               Width           =   1335
            End
            Begin VB.TextBox txtCheckView 
               Height          =   270
               Left            =   1560
               TabIndex        =   20
               Top             =   225
               Width           =   1335
            End
            Begin VB.Label Label3 
               Caption         =   "��    �飺"
               Height          =   255
               Left            =   360
               TabIndex        =   25
               Top             =   975
               Width           =   975
            End
            Begin VB.Label Label2 
               Caption         =   "��������"
               Height          =   255
               Left            =   360
               TabIndex        =   24
               Top             =   615
               Width           =   975
            End
            Begin VB.Label Label1 
               Caption         =   "���������"
               Height          =   255
               Left            =   360
               TabIndex        =   23
               Top             =   240
               Width           =   975
            End
         End
         Begin VB.CheckBox chkShowVideoCapture 
            Caption         =   "��ʾ��Ƶ�ɼ�����"
            Height          =   180
            Left            =   600
            TabIndex        =   18
            Top             =   840
            Width           =   2055
         End
         Begin VB.Frame frmShowBigImg 
            Height          =   735
            Left            =   480
            TabIndex        =   14
            Top             =   3480
            Width           =   6735
            Begin VB.OptionButton optBigImgAction 
               Caption         =   "����ƶ�ʱ��ʾ��ͼ���Ŵ���Ϊ��"
               Height          =   180
               Index           =   1
               Left            =   120
               TabIndex        =   17
               Top             =   240
               Value           =   -1  'True
               Width           =   3255
            End
            Begin VB.OptionButton optBigImgAction 
               Caption         =   "������ʾ��ͼ����"
               Height          =   180
               Index           =   2
               Left            =   120
               TabIndex        =   16
               Top             =   480
               Width           =   1815
            End
            Begin VB.ComboBox cboZoom 
               Height          =   300
               ItemData        =   "frmReportSetup.frx":0000
               Left            =   3360
               List            =   "frmReportSetup.frx":0010
               TabIndex        =   15
               Text            =   "1"
               Top             =   200
               Width           =   855
            End
         End
         Begin VB.CheckBox chkShowBigImg 
            Caption         =   "��ʾ��ͼ��"
            Height          =   300
            Left            =   600
            TabIndex        =   13
            Top             =   3240
            Width           =   1335
         End
         Begin VB.TextBox txtMinImageCount 
            Height          =   270
            Left            =   6240
            MaxLength       =   2
            TabIndex        =   12
            Text            =   "8"
            Top             =   360
            Width           =   975
         End
         Begin VB.CheckBox chkShowImage 
            Caption         =   "��ʾ����ͼ������                   ��������ͼ��ʾ������"
            Height          =   180
            Left            =   600
            TabIndex        =   11
            Top             =   420
            Width           =   5415
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "����ʾ�˫����"
         Height          =   855
         Left            =   120
         TabIndex        =   7
         Top             =   5640
         Width           =   2415
         Begin VB.OptionButton optWordDblClick 
            Caption         =   "ֱ��д�뱨��"
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   9
            Top             =   240
            Width           =   1455
         End
         Begin VB.OptionButton optWordDblClick 
            Caption         =   "�򿪴ʾ�༭����"
            Height          =   255
            Index           =   1
            Left            =   360
            TabIndex        =   8
            Top             =   480
            Width           =   1750
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "����ͼ˫����"
         Height          =   855
         Left            =   2520
         TabIndex        =   4
         Top             =   5640
         Width           =   2895
         Begin VB.OptionButton optImageDblClick 
            Caption         =   "��ͼƬ�༭����"
            Height          =   255
            Index           =   1
            Left            =   360
            TabIndex        =   6
            Top             =   480
            Width           =   1750
         End
         Begin VB.OptionButton optImageDblClick 
            Caption         =   "ֱ��д�뱨��"
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   5
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "�ʾ�ģ����ʾ"
         Height          =   855
         Left            =   5400
         TabIndex        =   1
         Top             =   5640
         Width           =   2415
         Begin VB.OptionButton optShowWord 
            Caption         =   "˫������"
            Height          =   180
            Index           =   1
            Left            =   360
            TabIndex        =   3
            Top             =   480
            Width           =   1095
         End
         Begin VB.OptionButton optShowWord 
            Caption         =   "ֱ����ʾ"
            Height          =   180
            Index           =   0
            Left            =   360
            TabIndex        =   2
            Top             =   240
            Width           =   1095
         End
      End
   End
End
Attribute VB_Name = "frmReportSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngDeptID As Long   '����ID
Private mblnRefreshed As Boolean

Public Sub zlRefresh(lngDeptID As Long)
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim lngTemp As Long
    
    mblnRefreshed = True            '���ݱ�ˢ�¹��ˣ����Ա���
    
    mlngDeptID = lngDeptID
    optReportEditor(0).value = True 'Ĭ��ʹ�õ��Ӳ����༭���༭����
    chkShowImage.value = 0          'Ĭ�ϲ���ʾͼ������
    chkShowVideoCapture.value = 0   'Ĭ�ϲ���ʾ��Ƶ�ɼ�����
    chkShowBigImg.value = 0         'Ĭ������ƶ�ʱ����ʾ��ͼ
    optBigImgAction(1).value = True 'Ĭ������ƶ���ʱ����ʾ��ͼ
    frmShowBigImg.Enabled = False   'Ĭ������ƶ�ʱ����ʾ��ͼ
    
    chkSpecialContent.value = 0     'Ĭ�ϲ���ʾר�Ʊ���
    cboSpecialContent.Enabled = False
    cboZoom.Text = 1                'Ĭ�ϷŴ���Ϊ1
    chkExitAfterPrint.value = 0     'Ĭ�ϴ�ӡ���˳�
    optWordDblClick(0).value = True 'Ĭ��˫���ʾ��ֱ��д�뱨��
    optImageDblClick(0).value = True 'Ĭ�ϱ�������ͼ˫����ֱ��д�뱨��
    txtCheckView.Text = "�������"  'Ĭ��Ϊ�������
    txtResult.Text = "������"     'Ĭ��Ϊ������
    txtAdvice.Text = "����"         'Ĭ��Ϊ����
    optShowWord(0).value = True     'Ĭ��Ϊֱ����ʾ�ʾ�ģ��
    chkUntreadPrinted.value = 0     'Ĭ��Ϊ��˴�ӡ���������
     
    On Error GoTo err
    strSql = "select ID ,����ID,������,����ֵ from Ӱ�����̲��� where ����ID = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngDeptID)
    
    While Not rsTemp.EOF
        Select Case rsTemp!������
            Case "����༭��"
                If Nvl(rsTemp!����ֵ, 0) = 0 Then
                    optReportEditor(0).value = True
                ElseIf Nvl(rsTemp!����ֵ, 0) = 1 Then
                    optReportEditor(1).value = True
                Else
                    optReportEditor(2).value = True
                End If
            Case "�鿴��ʷ����"
                If Nvl(rsTemp!����ֵ, 0) = 0 Then
                    optHistoryReportEditor(0).value = True
                Else
                    optHistoryReportEditor(1).value = True
                End If
                
            Case "��ʾ����ͼ��"
                chkShowImage.value = Nvl(rsTemp!����ֵ, 0)
            Case "��������ͼ����"
                txtMinImageCount.Text = Nvl(rsTemp!����ֵ, "8")
            Case "��ʾ��Ƶ�ɼ�"
                chkShowVideoCapture.value = Nvl(rsTemp!����ֵ, 0)
            Case "��ӡ���˳�"
                chkExitAfterPrint.value = Nvl(rsTemp!����ֵ, 0)
            Case "��������ʾ��ͼ"
                lngTemp = Nvl(rsTemp!����ֵ, 0)
                If lngTemp = 0 Then
                    chkShowBigImg.value = 0
                ElseIf lngTemp = 1 Then
                    chkShowBigImg.value = 1
                    optBigImgAction(1).value = True
                Else
                    chkShowBigImg.value = 1
                    optBigImgAction(2).value = True
                End If
                frmShowBigImg.Enabled = IIf(chkShowBigImg.value = 1, True, False)
            Case "�����ͼ�Ŵ���"
                cboZoom.Text = Nvl(rsTemp!����ֵ, 1)
                If Val(cboZoom.Text) = 0 Then cboZoom.Text = 1
            Case "��ʾר�Ʊ���"
                chkSpecialContent.value = Nvl(rsTemp!����ֵ, 0)
                cboSpecialContent.Enabled = IIf(chkSpecialContent.value = 1, True, False)
            Case "ר�Ʊ���ҳ"
                cboSpecialContent.Text = Nvl(rsTemp!����ֵ)
            Case "����ʾ�˫������"
                If Nvl(rsTemp!����ֵ, 0) = 0 Then
                    optWordDblClick(0).value = True
                Else
                    optWordDblClick(1).value = True
                End If
            Case "����ͼ˫������"
                If Nvl(rsTemp!����ֵ, 0) = 0 Then
                    optImageDblClick(0).value = True
                Else
                    optImageDblClick(1).value = True
                End If
            Case "�����������"
                txtCheckView.Text = Nvl(rsTemp!����ֵ, "�������")
            Case "����������"
                txtResult.Text = Nvl(rsTemp!����ֵ, "������")
            Case "��������"
                txtAdvice.Text = Nvl(rsTemp!����ֵ, "����")
            Case "��ʾ�ʾ�ʾ��"
                If Nvl(rsTemp!����ֵ, 0) = 0 Then
                    optShowWord(0).value = True
                Else
                    optShowWord(1).value = True
                End If
            Case "��˴�ӡ���������"
                chkUntreadPrinted.value = Nvl(rsTemp!����ֵ, 0)
        End Select
        rsTemp.MoveNext
    Wend
    
    If optReportEditor(2).value Then
        fraEditorSetUp.Visible = True
        
    Else
        fraEditorSetUp.Visible = False
    End If
    
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume Next
    Call SaveErrLog
End Sub


Public Sub zlSave()
    Dim intMatch As Integer
    Dim strSql As String
    
    On Error GoTo errHand
    
    If mblnRefreshed = False Then Exit Sub          '����û�б�ˢ�£����Բ�����
    
    If optReportEditor(0).value = True Then         '���Ӳ����༭��
        intMatch = 0
    ElseIf optReportEditor(1).value = True Then     'PACS����༭��
        intMatch = 1
    ElseIf optReportEditor(2).value = True Then     '�����ĵ��༭��
        intMatch = 2
    End If
    strSql = "ZL_Ӱ�����̲���_UPDATE( " & mlngDeptID & ", '����༭��','" & intMatch & "')"
    zlDatabase.ExecuteProcedure strSql, Me.Caption
    
    strSql = "ZL_Ӱ�����̲���_UPDATE( " & mlngDeptID & ", '��ʾ����ͼ��','" & chkShowImage.value & "')"
    zlDatabase.ExecuteProcedure strSql, Me.Caption
    
    strSql = "ZL_Ӱ�����̲���_UPDATE( " & mlngDeptID & ", '��������ͼ����','" & txtMinImageCount.Text & "')"
    zlDatabase.ExecuteProcedure strSql, Me.Caption
    
    strSql = "ZL_Ӱ�����̲���_UPDATE( " & mlngDeptID & ", '��ʾ��Ƶ�ɼ�','" & chkShowVideoCapture.value & "')"
    zlDatabase.ExecuteProcedure strSql, Me.Caption
    
    strSql = "ZL_Ӱ�����̲���_UPDATE( " & mlngDeptID & ", '��ӡ���˳�','" & chkExitAfterPrint.value & "')"
    zlDatabase.ExecuteProcedure strSql, Me.Caption
    
    If chkShowBigImg.value = 0 Then
        intMatch = 0
    ElseIf optBigImgAction(1).value = True Then
        intMatch = 1
    Else
        intMatch = 2
    End If
    strSql = "ZL_Ӱ�����̲���_UPDATE( " & mlngDeptID & ", '��������ʾ��ͼ','" & intMatch & "')"
    zlDatabase.ExecuteProcedure strSql, Me.Caption

    strSql = "ZL_Ӱ�����̲���_UPDATE( " & mlngDeptID & ", '�����ͼ�Ŵ���','" & IIf(Val(cboZoom.Text) = 0, 1, Val(cboZoom.Text)) & "')"
    zlDatabase.ExecuteProcedure strSql, Me.Caption
    
    strSql = "ZL_Ӱ�����̲���_UPDATE( " & mlngDeptID & ", '��ʾר�Ʊ���','" & chkSpecialContent.value & "')"
    zlDatabase.ExecuteProcedure strSql, Me.Caption
    
    strSql = "ZL_Ӱ�����̲���_UPDATE( " & mlngDeptID & ", 'ר�Ʊ���ҳ','" & cboSpecialContent.Text & "')"
    zlDatabase.ExecuteProcedure strSql, Me.Caption
    
    If optWordDblClick(0).value = True Then         '����ʾ�˫����ֱ��д�뱨��
        intMatch = 0
    ElseIf optWordDblClick(1).value = True Then     '����ʾ�˫����򿪱༭����
        intMatch = 1
    End If
    strSql = "ZL_Ӱ�����̲���_UPDATE( " & mlngDeptID & ", '����ʾ�˫������','" & intMatch & "')"
    zlDatabase.ExecuteProcedure strSql, Me.Caption
    
    If optImageDblClick(0).value = True Then         '����ͼ˫����ֱ��д�뱨��
        intMatch = 0
    ElseIf optImageDblClick(1).value = True Then     '����ͼ˫�����ͼ��༭����
        intMatch = 1
    End If
    strSql = "ZL_Ӱ�����̲���_UPDATE( " & mlngDeptID & ", '����ͼ˫������','" & intMatch & "')"
    zlDatabase.ExecuteProcedure strSql, Me.Caption
    
    strSql = "ZL_Ӱ�����̲���_UPDATE( " & mlngDeptID & ", '�����������','" & txtCheckView.Text & "')"
    zlDatabase.ExecuteProcedure strSql, Me.Caption
    
    strSql = "ZL_Ӱ�����̲���_UPDATE( " & mlngDeptID & ", '����������','" & txtResult.Text & "')"
    zlDatabase.ExecuteProcedure strSql, Me.Caption
    
    strSql = "ZL_Ӱ�����̲���_UPDATE( " & mlngDeptID & ", '��������','" & txtAdvice.Text & "')"
    zlDatabase.ExecuteProcedure strSql, Me.Caption
    
    If optShowWord(0).value = True Then         'ֱ����ʾ�ʾ�ʾ��
        intMatch = 0
    ElseIf optShowWord(1).value = True Then     '˫���������ʾ�ʾ�ʾ��
        intMatch = 1
    End If
    strSql = "ZL_Ӱ�����̲���_UPDATE( " & mlngDeptID & ", '��ʾ�ʾ�ʾ��','" & intMatch & "')"
    zlDatabase.ExecuteProcedure strSql, Me.Caption
    
    strSql = "ZL_Ӱ�����̲���_UPDATE( " & mlngDeptID & ", '��˴�ӡ���������','" & chkUntreadPrinted.value & "')"
    zlDatabase.ExecuteProcedure strSql, Me.Caption
    
    If optReportEditor(2) Then
        strSql = "ZL_Ӱ�����̲���_UPDATE( " & mlngDeptID & ", '�鿴��ʷ����','" & IIf(optHistoryReportEditor(0).value, 0, 1) & "')"
        zlDatabase.ExecuteProcedure strSql, Me.Caption
    End If
    
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume Next
    Call SaveErrLog
End Sub

Private Sub chkShowBigImg_Click()
    frmShowBigImg.Enabled = IIf(chkShowBigImg.value = 1, True, False)
End Sub

Private Sub chkSpecialContent_Click()
    If chkSpecialContent.value = 1 Then
        cboSpecialContent.Enabled = True
    Else
        cboSpecialContent.Enabled = False
    End If
End Sub

Private Sub Form_Load()
    mblnRefreshed = False
    'װ��ר�Ʊ�������
    cboSpecialContent.Clear
    cboSpecialContent.AddItem (Report_Form_frmReportES)
    cboSpecialContent.AddItem (Report_Form_frmReportPathology)
    cboSpecialContent.AddItem (Report_Form_frmReportUS)
    cboSpecialContent.AddItem (Report_Form_frmReportCustom)
End Sub

Private Sub Form_Resize()
    fraReportSetup.Left = (Me.ScaleWidth - fraReportSetup.Width) / 2
End Sub

Private Sub optBigImgAction_Click(Index As Integer)
    If frmShowBigImg.Enabled = True Then
        cboZoom.Enabled = IIf(Index = 1, True, False)
    Else
        cboZoom.Enabled = False
    End If
End Sub

Private Sub optReportEditor_Click(Index As Integer)
    fraEditorSetUp.Visible = Index = 2
End Sub
