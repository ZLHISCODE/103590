VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#6.9#0"; "ZLIDKIND.OCX"
Begin VB.Form frmClinicPlanStopVisitAndModifyDoctor 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ͣ��"
   ClientHeight    =   5460
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7380
   Icon            =   "frmClinicPlanStopVisitAndModifyDoctor.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   6120
      TabIndex        =   28
      Top             =   720
      Width           =   1100
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   6120
      TabIndex        =   27
      Top             =   240
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   6120
      TabIndex        =   29
      Top             =   3330
      Width           =   1100
   End
   Begin VB.Frame fra����ҽ�� 
      Caption         =   "����ҽ��"
      Height          =   765
      Left            =   60
      TabIndex        =   23
      Top             =   3960
      Width           =   5895
      Begin VB.ComboBox cbo����ҽ�� 
         Height          =   300
         Left            =   1110
         TabIndex        =   25
         Text            =   "����"
         Top             =   300
         Width           =   4575
      End
      Begin zlIDKind.IDKindNew idkDoctor 
         Height          =   315
         Left            =   120
         TabIndex        =   24
         Top             =   270
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         IDKindStr       =   "��|Ժ��ҽ��|0|0|0|0|0||0|0|0;��|Ժ��ҽ��|0|0|0|0|0||0|0|0"
         CaptionAlignment=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontSize        =   9
         FontName        =   "����"
         IDKind          =   -1
         DefaultCardType =   "0"
         NotAutoAppendKind=   -1  'True
         BackColor       =   -2147483633
      End
   End
   Begin VB.Frame fra������Ϣ 
      Caption         =   "������Ϣ"
      Height          =   1185
      Left            =   60
      TabIndex        =   14
      Top             =   2610
      Width           =   5895
      Begin VB.TextBox txt�ϰ�ʱ�� 
         Enabled         =   0   'False
         Height          =   300
         Left            =   840
         TabIndex        =   26
         Top             =   720
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker dtpStart 
         Height          =   300
         Left            =   3150
         TabIndex        =   21
         Top             =   720
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "HH:mm:ss"
         Format          =   184680451
         UpDown          =   -1  'True
         CurrentDate     =   42360.3333333333
      End
      Begin VB.ComboBox cboͣ��ԭ�� 
         Height          =   300
         Left            =   3150
         TabIndex        =   18
         Text            =   "����"
         Top             =   330
         Width           =   2535
      End
      Begin VB.TextBox txt�������� 
         Enabled         =   0   'False
         Height          =   300
         Left            =   840
         TabIndex        =   16
         Text            =   "2016-04-05"
         Top             =   330
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker dtpEnd 
         Height          =   300
         Left            =   4620
         TabIndex        =   22
         Top             =   720
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "HH:mm:ss"
         Format          =   184680451
         UpDown          =   -1  'True
         CurrentDate     =   42360.5
      End
      Begin VB.Label lblTimeRange 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��"
         Height          =   180
         Left            =   4350
         TabIndex        =   30
         Top             =   780
         Width           =   180
      End
      Begin VB.Label lblͣ��ʱ�� 
         AutoSize        =   -1  'True
         Caption         =   "ͣ��ʱ��"
         Height          =   180
         Left            =   2400
         TabIndex        =   20
         Top             =   780
         Width           =   720
      End
      Begin VB.Label lblͣ��ԭ�� 
         AutoSize        =   -1  'True
         Caption         =   "ͣ��ԭ��"
         Height          =   180
         Left            =   2400
         TabIndex        =   17
         Top             =   390
         Width           =   720
      End
      Begin VB.Label lbl�ϰ�ʱ�� 
         AutoSize        =   -1  'True
         Caption         =   "�ϰ�ʱ��"
         Height          =   180
         Left            =   90
         TabIndex        =   19
         Top             =   780
         Width           =   720
      End
      Begin VB.Label lbl�������� 
         AutoSize        =   -1  'True
         Caption         =   "��������"
         Height          =   180
         Left            =   90
         TabIndex        =   15
         Top             =   390
         Width           =   720
      End
   End
   Begin VB.Frame fra��Դ��Ϣ 
      Caption         =   "��Դ������Ϣ"
      Height          =   2295
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   5895
      Begin VB.TextBox txtSignalNO 
         Enabled         =   0   'False
         Height          =   300
         Left            =   810
         TabIndex        =   2
         Text            =   "4"
         Top             =   330
         Width           =   1335
      End
      Begin VB.TextBox txtItem 
         Enabled         =   0   'False
         Height          =   300
         Left            =   810
         TabIndex        =   8
         Text            =   "����ҽʦ��"
         Top             =   1110
         Width           =   4875
      End
      Begin VB.TextBox txtDept 
         Enabled         =   0   'False
         Height          =   300
         Left            =   810
         TabIndex        =   6
         Text            =   "�����ڿ�"
         Top             =   720
         Width           =   4875
      End
      Begin VB.TextBox txtDoctor 
         Enabled         =   0   'False
         Height          =   300
         Left            =   810
         TabIndex        =   10
         Text            =   "����"
         Top             =   1500
         Width           =   4875
      End
      Begin VB.TextBox txt���տ��� 
         Enabled         =   0   'False
         Height          =   300
         Left            =   810
         TabIndex        =   12
         Text            =   "���ϰ�"
         Top             =   1890
         Width           =   1965
      End
      Begin VB.CheckBox chk���� 
         Caption         =   "�Һ�ʱ���뽨��"
         Enabled         =   0   'False
         Height          =   180
         Left            =   3420
         TabIndex        =   13
         Top             =   1935
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.TextBox txt���� 
         Enabled         =   0   'False
         Height          =   300
         Left            =   3420
         TabIndex        =   4
         Text            =   "��ͨ"
         Top             =   330
         Width           =   2265
      End
      Begin VB.Label lbl���տ��� 
         AutoSize        =   -1  'True
         Caption         =   "���տ���"
         Height          =   180
         Left            =   60
         TabIndex        =   11
         Top             =   1950
         Width           =   720
      End
      Begin VB.Label lblDoctor 
         AutoSize        =   -1  'True
         Caption         =   "ҽ��"
         Height          =   180
         Left            =   420
         TabIndex        =   9
         Top             =   1560
         Width           =   360
      End
      Begin VB.Label lblDept 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   180
         Left            =   420
         TabIndex        =   5
         Top             =   780
         Width           =   360
      End
      Begin VB.Label lblItem 
         AutoSize        =   -1  'True
         Caption         =   "��Ŀ"
         Height          =   180
         Left            =   450
         TabIndex        =   7
         Top             =   1170
         Width           =   360
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   180
         Left            =   3030
         TabIndex        =   3
         Top             =   390
         Width           =   360
      End
      Begin VB.Label lblSignalNO 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   180
         Left            =   420
         TabIndex        =   1
         Top             =   390
         Width           =   360
      End
   End
End
Attribute VB_Name = "frmClinicPlanStopVisitAndModifyDoctor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnOk As Boolean
Private mlngModule As Long
Private Enum m_FunType
    F_ͣ�� = 1
    F_ȡ��ͣ�� = 2
    F_���� = 3
    F_ȡ������ = 4
End Enum
Private mbytFun As m_FunType
Private mlng��¼ID As Long '��¼ID
Private mrsStopReason As ADODB.Recordset

'����
Private mblnOnlyԺ��ҽ�� As Boolean '��ֻ����Ժ��ҽ��
Private mbln����ҽ�������� As Boolean
Private mbytԤԼ�嵥���Ʒ�ʽ As Byte
Private mbytԤԼ�嵥��ӡ��ʽ As Byte

Private mblnCboClick As Boolean     '�����cbo��keypress�¼������˵����б��API����:sendmessage,�����ͣ��cbo��,����һ���ַ�,�ƿ�����򰴻س���,
'                                    cbo��ֵ�ᱣ������,�����ᴥ��click�¼�,������Ҫ��validate�¼��е���click�¼�

Public Function ShowMe(frmParent As Form, ByVal lngModule As Long, ByVal bytFun As Byte, _
    ByVal lng��¼ID As Long) As Boolean
    '�������
    '��Σ�
    '   frmParent ������
    '   lngModule ģ���
    '   bytFun ���ܣ�1-ͣ��,2-ȡ��ͣ��,3-����,4-ȡ������
    '   lng��¼ID �����¼ID
    mbytFun = bytFun: mlngModule = lngModule
    mlng��¼ID = lng��¼ID
    
    On Error Resume Next
    If EditBeforCheck(bytFun, lng��¼ID) = False Then Exit Function
    mblnOk = False
    Me.Show 1, frmParent
    ShowMe = mblnOk
End Function

Private Function EditBeforCheck(ByVal bytFun As m_FunType, ByVal lng��¼ID As Long) As Boolean
    '�Գ��ﰲ�Ž������Ƽ��
    Dim strSQL, rsTemp As ADODB.Recordset
    
    On Error GoTo errHandler
    '���ܶ���ʷ�İ��Ž��в���
    strSQL = "Select 1 From �ٴ������¼ A Where ID = [1] And a.��ֹʱ�� < Sysdate"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "���ܶ���ʷ�İ��Ž��в���", lng��¼ID)
    If Not rsTemp.EOF Then
        MsgBox "���ܶ���ʷ�İ��Ž��в�����", vbInformation, gstrSysName
        Exit Function
    End If
    
    If bytFun = F_���� Then
        '���δ����ҽ���ĺ�Դ���������������
        strSQL = "Select 1 From �ٴ������¼ A Where ID = [1] And a.ҽ������ Is Not Null"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "δ����ҽ���ĺ�Դ����������", lng��¼ID)
        If rsTemp.EOF Then
            MsgBox "�ú�Դδ����ҽ�������������������", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    EditBeforCheck = True
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub cbo����ҽ��_Click()
    Dim strSQL As String, blnCancel As Boolean
    Dim rsReturn As ADODB.Recordset
    Dim vRect  As RECT
    
    Err = 0: On Error GoTo errHandler
    If cbo����ҽ��.Text = "��������ҽ��..." Then
        'ѡ��"��������ҽ��..."ʱ������ѡ����
        cbo����ҽ��.ListIndex = -1
        Call GetDoctor(Val(txtDept.Tag), "", True, True, strSQL)  '��ȡSQL���
        vRect = zlControl.GetControlRect(cbo����ҽ��.Hwnd)
        
        Set rsReturn = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "����ҽ��", False, _
                       "", "", False, False, False, vRect.Left, vRect.Top, cbo����ҽ��.Height, blnCancel, True, False)
        If blnCancel Then Exit Sub
        If rsReturn Is Nothing Then Exit Sub
        If rsReturn.EOF Then Exit Sub
        
        With rsReturn
            zlControl.CboLocate cbo����ҽ��, Nvl(!ID), True
            If cbo����ҽ��.ListIndex = -1 Then
                cbo����ҽ��.AddItem Nvl(!����) & IIf(Nvl(!רҵ����ְ��) = "", "", "(" & Nvl(!רҵ����ְ��) & ")"), cbo����ҽ��.ListCount - 1
                cbo����ҽ��.ItemData(cbo����ҽ��.NewIndex) = Val(Nvl(!ID))
                cbo����ҽ��.ListIndex = cbo����ҽ��.NewIndex
            End If
        End With
    End If
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cbo����ҽ��_GotFocus()
    zlControl.TxtSelAll cbo����ҽ��
End Sub

Private Sub cbo����ҽ��_KeyPress(KeyAscii As Integer)
    Dim strSQL As String, blnCancel As Boolean
    Dim rsReturn As ADODB.Recordset
    Dim vRect  As RECT
    Dim strKey As String, strWhere As String
    
    Err = 0: On Error GoTo errHandle
    If KeyAscii <> 13 Then Exit Sub
    If cbo����ҽ��.ListIndex <> -1 Or mblnOnlyԺ��ҽ�� = False Then
        zlCommFun.PressKey vbKeyTab: Exit Sub
    End If
    If Trim(cbo����ҽ��.Text) = "" Then zlCommFun.PressKey vbKeyTab: Exit Sub
    
    'ģ��ƥ��,ѡ��ҽ��
    strKey = gstrLike & Trim(cbo����ҽ��.Text) & "%"
    If zlCommFun.IsCharChinese(Trim(cbo����ҽ��.Text)) Then
         strWhere = " And ���� like [1] "
    ElseIf zlCommFun.IsNumOrChar(Trim(cbo����ҽ��.Text)) Then
         strWhere = " And (���� like upper([1]) or ��� like upper([1]))"
    End If
        
    Call GetDoctor(0, strWhere, False, True, strSQL) '��ȡSQL���
    vRect = zlControl.GetControlRect(cbo����ҽ��.Hwnd)
    Set rsReturn = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "����ҽ��", False, _
                   "", "", False, False, True, vRect.Left, vRect.Top, cbo����ҽ��.Height, blnCancel, True, False, strKey)
    If blnCancel Then Exit Sub
    If rsReturn Is Nothing Then Exit Sub
    If rsReturn.EOF Then Exit Sub
    
    zlControl.CboLocate cbo����ҽ��, Nvl(rsReturn!ID), True
    If cbo����ҽ��.ListIndex = -1 And Nvl(rsReturn!����) <> "��������ҽ��..." Then
        cbo����ҽ��.AddItem Nvl(rsReturn!����) & IIf(Nvl(rsReturn!רҵ����ְ��) = "", "", "(" & Nvl(rsReturn!רҵ����ְ��) & ")"), cbo����ҽ��.ListCount - 1
        cbo����ҽ��.ItemData(cbo����ҽ��.NewIndex) = Val(Nvl(rsReturn!ID))
        cbo����ҽ��.ListIndex = cbo����ҽ��.NewIndex
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub cbo����ҽ��_Validate(Cancel As Boolean)
    If mblnOnlyԺ��ҽ�� Then
        If cbo����ҽ��.ListIndex < 0 Then cbo����ҽ��.Text = ""
    End If
End Sub

Private Sub cboͣ��ԭ��_GotFocus()
    zlControl.TxtSelAll cboͣ��ԭ��
End Sub

Private Sub cboͣ��ԭ��_KeyPress(KeyAscii As Integer)
    Dim strReason As String
    
    If KeyAscii <> vbKeyReturn Then Exit Sub
    If Trim(cboͣ��ԭ��.Text) = "" Then zlCommFun.PressKey vbKeyTab: Exit Sub
    strReason = SearchStopVisitReason(Me, cboͣ��ԭ��, Trim(cboͣ��ԭ��.Text))
    If strReason = "" Then Exit Sub
    
    zlControl.CboLocate cboͣ��ԭ��, strReason
    If cboͣ��ԭ��.ListIndex = -1 Then cboͣ��ԭ��.Text = strReason
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Function IsValied() As Boolean
    Dim arrTime As Variant
    Dim dtStartTime As Date, dtEndTime As Date
    Dim dtStartTimeNew As Date, dtEndTimeNew As Date
    Dim strSQL As String, strWhere As String, rsTemp As ADODB.Recordset
    Dim lngDoctor As Long
    
    Err = 0: On Error GoTo errHandle
    If mbytFun = F_ȡ��ͣ�� Or mbytFun = F_ȡ������ Then
        '�������
        If mbytFun = F_ȡ��ͣ�� Then
            strSQL = "Select 1 From �ٴ������¼ A Where a.ID = [1] And  a.ͣ�￪ʼʱ�� Is Null"
        Else
            strSQL = "Select 1 From �ٴ������¼ A Where a.ID = [1] And  a.���￪ʼʱ�� Is Null"
        End If
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng��¼ID)
        If Not rsTemp.EOF Then
            MsgBox "��ǰ�����ѱ�����ȡ��" & IIf(mbytFun = F_ȡ������, "����", "ͣ��") & "����ˢ�����ݺ�鿴��", vbInformation, gstrSysName
            Exit Function
        End If
        
        If mbytFun = F_ȡ��ͣ�� Then
            strSQL = "Select 1 From �ٴ������¼ A Where a.ID = [1] And  a.ͣ����ֹʱ��< Sysdate"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng��¼ID)
            If Not rsTemp.EOF Then
                MsgBox "ͣ��ʱ�����ֹʱ��С���˵�ǰʱ�䣬���ܽ���ȡ��ͣ�������", vbInformation, gstrSysName
                Exit Function
            End If
        Else
            strSQL = "Select 1 From �ٴ������¼ A Where a.ID = [1] And  a.���￪ʼʱ��< Sysdate"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng��¼ID)
            If Not rsTemp.EOF Then
                MsgBox "����ʱ��Ŀ�ʼʱ��С���˵�ǰʱ�䣬���ܽ���ȡ�����������", vbInformation, gstrSysName
                Exit Function
            End If
        End If
        
        If mbytFun = F_ȡ��ͣ�� Then '���ú�ǰ��Чʱ�䲻���н���
            '��ȥ��ǰ�����ϰ�ʱ��
            strSQL = "Select a.��ʼʱ��, a.��ֹʱ��, a.ͣ�￪ʼʱ��, a.ͣ����ֹʱ��" & vbNewLine & _
                    " From �ٴ������¼ A, �ٴ������¼ B" & vbNewLine & _
                    " Where a.��Դid = b.��Դid And a.�������� = b.�������� And b.Id = [1] And a.Id <> b.Id" & vbNewLine

            '������ú��Ƿ��н���
            strSQL = "Select 1 From " & _
                    "  (Select ��ʼʱ��, ͣ�￪ʼʱ�� As ��ֹʱ�� From (" & strSQL & ") Where ��ʼʱ�� < ͣ�￪ʼʱ�� And ��ֹʱ�� = ͣ����ֹʱ��" & vbNewLine & _
                    "   Union All" & vbNewLine & _
                    "   Select ͣ����ֹʱ�� As ��ʼʱ��, ��ֹʱ�� From (" & strSQL & ") Where ��ʼʱ�� = ͣ�￪ʼʱ�� And ��ֹʱ�� > ͣ����ֹʱ��" & vbNewLine & _
                    "   Union All" & vbNewLine & _
                    "   Select ��ʼʱ��, ͣ�￪ʼʱ�� As ��ֹʱ�� From (" & strSQL & ") Where ��ʼʱ�� < ͣ�￪ʼʱ�� And ��ֹʱ�� > ͣ����ֹʱ��" & vbNewLine & _
                    "   Union All" & vbNewLine & _
                    "   Select ͣ����ֹʱ�� As ��ʼʱ��, ��ֹʱ�� From (" & strSQL & ") Where ��ʼʱ�� < ͣ�￪ʼʱ�� And ��ֹʱ�� > ͣ����ֹʱ��) M, �ٴ������¼ N" & vbNewLine & _
                    " Where m.��ʼʱ�� < n.��ֹʱ�� And m.��ֹʱ�� > n.��ʼʱ�� And n.Id = [1] And Rownum < 2"
            '����ʹ��With��䣬Ҫ����
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng��¼ID)
            If Not rsTemp.EOF Then
                MsgBox "��ǰ�ϰ�ʱ�ε�ʱ�䷶Χ��ú�Դ����Ŀǰ��Ч���ϰ�ʱ�ε�ʱ�䷶Χ�н��棬�㲻��ȡ��ͣ�", vbInformation + vbOKOnly, gstrSysName
                Exit Function
            End If
        End If
        IsValied = True: Exit Function
    End If
    
    '�������
    If mbytFun = F_ͣ�� Then
        strSQL = "Select 1 From �ٴ������¼ A Where a.ID = [1] And  a.ͣ�￪ʼʱ�� Is Not Null"
    Else
        strSQL = "Select 1 From �ٴ������¼ A Where a.ID = [1] And  a.���￪ʼʱ�� is Not Null"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng��¼ID)
    If Not rsTemp.EOF Then
        MsgBox "��ǰ�����ѱ����˽�����" & IIf(mbytFun = F_����, "����", "ͣ��") & "����ˢ�����ݺ�鿴��", vbInformation, gstrSysName
        Exit Function
    End If
    
    If zlControl.TxtCheckInput(cboͣ��ԭ��, "ͣ��ԭ��", 50, False) = False Then Exit Function
    If Trim(cbo����ҽ��.Text) = "" And cbo����ҽ��.Visible Then
        MsgBox "����ҽ������Ϊ�գ�", vbInformation, gstrSysName
        If cbo����ҽ��.Visible And cbo����ҽ��.Enabled Then cbo����ҽ��.SetFocus
        Exit Function
    End If
    If mblnOnlyԺ��ҽ�� Then
        If cbo����ҽ��.ListIndex < 0 And Trim(cbo����ҽ��.Text) <> "" And cbo����ҽ��.Visible Then
            MsgBox "��ѡ���ҽ�������ڣ�����������ҽ����", vbInformation + vbOKOnly, gstrSysName
            If cbo����ҽ��.Visible And cbo����ҽ��.Enabled Then cbo����ҽ��.SetFocus
            Exit Function
        End If
    End If
    
    'ͣ��/����ʱ����
    dtStartTime = Format(dtpStart.Tag, "yyyy-mm-dd hh:mm:ss")
    dtEndTime = Format(dtpEnd.Tag, "yyyy-mm-dd hh:mm:ss")
    dtStartTimeNew = GetWorkTrueDate(dtStartTime, Format(dtStartTime, "yyyy-mm-dd ") & Format(dtpStart.Value, "hh:mm:ss"), True, False)
    dtEndTimeNew = GetWorkTrueDate(dtStartTime, Format(dtStartTime, "yyyy-mm-dd ") & Format(dtpEnd.Value, "hh:mm:ss"))
    If dtStartTimeNew >= dtEndTimeNew Then
        MsgBox IIf(mbytFun = F_����, "����", "ͣ��") & "ʱ�䷶Χ�Ľ���ʱ�������ڿ�ʼʱ�䣡", vbInformation, gstrSysName
        If dtpEnd.Visible And dtpEnd.Enabled Then dtpEnd.SetFocus
        Exit Function
    End If
    If Not ((DateDiff("n", dtStartTime, dtStartTimeNew) >= 0 And DateDiff("n", dtStartTimeNew, dtEndTime) >= 0) _
            And (DateDiff("n", dtStartTime, dtEndTimeNew) >= 0 And DateDiff("n", dtEndTimeNew, dtEndTime) >= 0)) Then
        MsgBox IIf(mbytFun = F_����, "����", "ͣ��") & "ʱ��������ϰ�ʱ��ʱ�䷶Χ(" & Format(dtStartTime, "hh:mm") & "-" & Format(dtEndTime, "hh:mm") & ")�ڣ�", vbInformation, gstrSysName
        If dtpEnd.Visible And dtpEnd.Enabled Then dtpEnd.SetFocus
        Exit Function
    End If
    
    If zlDatabase.Currentdate > dtStartTimeNew Then
        MsgBox IIf(mbytFun = F_����, "����", "ͣ��") & "ʱ��Ŀ�ʼʱ��С���˵�ǰʱ�䣬���ܽ���" & IIf(mbytFun = F_����, "����", "ͣ��") & "������", vbInformation, gstrSysName
        If dtpStart.Visible And dtpStart.Enabled Then dtpStart.SetFocus
        Exit Function
    End If
    
    If mbytFun = F_���� Then
        If mblnOnlyԺ��ҽ�� Then
            strWhere = " And a.ҽ��ID = [4]"
            lngDoctor = cbo����ҽ��.ItemData(cbo����ҽ��.ListIndex)
        Else
            strWhere = " And a.ҽ������ = [5] And a.ҽ��ID Is Null"
        End If
        
        If lngDoctor <> 0 Then
            strSQL = "Select 1 From �ٴ������¼ A Where ID = [1] And Nvl(ҽ��ID,����ҽ��ID)= [2] "
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng��¼ID, lngDoctor)
            If Not rsTemp.EOF Then
                MsgBox "����ҽ������Ϊԭ����ҽ������ѡ������ҽ����", vbInformation, gstrSysName
                If cbo����ҽ��.Visible And cbo����ҽ��.Enabled Then cbo����ҽ��.SetFocus
                Exit Function
            End If
        End If
        
        '�ڸ�ʱ���ڣ�����ҽ�����ܴ��������ĳ��ﰲ��
        '��A[A1,A2],B[B1,B2],��BΪ�ջ���ȫ������A��(A1<=B1,A2>=B2).��ôX[X1,X2]��A-B�н�������
        '(X1>=A1 And X1<=NVL(B1,A2)) Or (X2>=A1 And X2<=NVL(B1,A2)) Or (X1>=NVL(B2,A1) And X1<=A2) Or (X2>=NVL(B2,A1) And X2<=A2)
        strSQL = "Select 1" & vbNewLine & _
                " From �ٴ������¼ A" & vbNewLine & _
                " Where a.�������� = To_Date([1], 'yyyy-mm-dd')" & strWhere & vbNewLine & _
                "       And (([2] Between a.��ʼʱ�� And Nvl(a.ͣ�￪ʼʱ��, a.��ֹʱ��)) Or ([3] Between a.��ʼʱ�� And Nvl(a.ͣ�￪ʼʱ��, a.��ֹʱ��))" & vbNewLine & _
                "       Or ([2] Between Nvl(a.ͣ����ֹʱ��, a.��ʼʱ��) And a.��ֹʱ��) Or ([3] Between Nvl(a.ͣ����ֹʱ��, a.��ʼʱ��) And a.��ֹʱ��))"
        '����������
        strSQL = strSQL & vbNewLine & _
                "       And [2] < Nvl(a.������ֹʱ��, To_Date('1900-01-01','yyyy-mm-dd'))" & vbNewLine & _
                "       And [3] > Nvl(a.���￪ʼʱ��, To_Date('3000-01-01','yyyy-mm-dd'))"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, txt��������.Text, dtStartTimeNew, dtEndTimeNew, _
            lngDoctor, GetReplaceDoctor(Trim(cbo����ҽ��.Text)))
        If Not rsTemp.EOF Then
            MsgBox "����ҽ��������ʱ��(" & Format(dtStartTimeNew, "hh:mm") & "-" & Format(dtEndTimeNew, "hh:mm") & ")��Χ���Ѵ����������ﰲ�ţ���ѡ������ҽ����", vbInformation, gstrSysName
            If cbo����ҽ��.Visible And cbo����ҽ��.Enabled Then cbo����ҽ��.SetFocus
            Exit Function
        End If
        
        '����ҽ��������
        If mbln����ҽ�������� Then
           strSQL = "Select Zl1_Ex_Isdoctorsamelevel(a.ҽ��id, a.ҽ������, [2], [3]) As ��� From �ٴ������¼ A Where ID = [1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng��¼ID, lngDoctor, GetReplaceDoctor(Trim(cbo����ҽ��.Text)))
            If Not rsTemp.EOF Then
                 If Val(Nvl(rsTemp!���)) = -1 Then
                    MsgBox "����ҽ����ְ�񼶱𲻹��������������ѡ������ҽ����", vbInformation, gstrSysName
                    If cbo����ҽ��.Visible And cbo����ҽ��.Enabled Then cbo����ҽ��.SetFocus
                    Exit Function
                End If
            End If
        End If
    End If
    IsValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.Hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub cmdOK_Click()
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim str��¼IDs As String
    Dim strTemp As String
    
    Err = 0: On Error GoTo errHandle
    
    If IsValied() = False Then Exit Sub
    If SaveData() = False Then Exit Sub
    
    '����Ƿ�Ҫ���ԤԼ�嵥
    If mbytFun = F_ͣ�� Or mbytFun = F_���� Then
        If mbytFun = F_ͣ�� Then
            strSQL = "Select a.ID As ��¼ID" & vbNewLine & _
                    " From �ٴ������¼ A, ���˹Һż�¼ B, �ٴ������Դ C" & vbNewLine & _
                    " Where a.Id = b.�����¼id And a.��Դid = c.Id And b.��¼״̬ = 1 And Nvl(b.ִ��״̬, 0) = 0 And a.Id = [1] " & vbNewLine & _
                    "       And (b.��¼���� = 1 And b.����ʱ�� Between a.ͣ�￪ʼʱ�� And a.ͣ����ֹʱ��" & vbNewLine & _
                    "           Or b.��¼���� = 2 And b.ԤԼʱ�� Between a.ͣ�￪ʼʱ�� And a.ͣ����ֹʱ��) And Rownum < 2"
        Else
            strSQL = "Select a.ID As ��¼ID" & vbNewLine & _
                    " From �ٴ������¼ A, ���˹Һż�¼ B, �ٴ������Դ C" & vbNewLine & _
                    " Where a.Id = b.�����¼id And a.��Դid = c.Id And b.��¼״̬ = 1 And Nvl(b.ִ��״̬, 0) = 0  And a.Id = [1] " & vbNewLine & _
                    "       And (b.��¼���� = 1 And b.����ʱ�� Between a.���￪ʼʱ�� And a.������ֹʱ��" & vbNewLine & _
                    "           Or b.��¼���� = 2 And b.ԤԼʱ�� Between a.���￪ʼʱ�� And a.������ֹʱ��) And Rownum < 2"
        End If
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng��¼ID)
        
        If rsTemp Is Nothing Then GoTo UnloadForm:
        If rsTemp.EOF Then GoTo UnloadForm:
    
        Do While Not rsTemp.EOF
            If InStr(strTemp & ",", "," & Nvl(rsTemp!��¼ID) & ",") = 0 Then
                str��¼IDs = str��¼IDs & "," & Nvl(rsTemp!��¼ID)
            End If
            rsTemp.MoveNext
        Loop
        If str��¼IDs <> "" Then str��¼IDs = Mid(str��¼IDs, 2)
        
        If mbytԤԼ�嵥���Ʒ�ʽ = 1 Then
            Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_INSIDE_1114_4", Me, "�����¼IDS=" & str��¼IDs, 3)
        ElseIf mbytԤԼ�嵥���Ʒ�ʽ = 2 Then
            If MsgBox("��ǰ��Դͣ��ʱ���ڴ���ԤԼ��ҺŲ��ˣ��Ƿ�ԤԼ�嵥�����Excel�У�", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_INSIDE_1114_4", Me, "�����¼IDS=" & str��¼IDs, 3)
            End If
        End If
        
        If mbytԤԼ�嵥��ӡ��ʽ = 1 Then
            Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_INSIDE_1114_4", Me, "�����¼IDS=" & str��¼IDs, 2)
        ElseIf mbytԤԼ�嵥��ӡ��ʽ = 2 Then
            If MsgBox("��ǰ��Դͣ��ʱ���ڴ���ԤԼ��ҺŲ��ˣ���ȷ��Ҫ��ӡԤԼ�嵥��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_INSIDE_1114_4", Me, "�����¼IDS=" & str��¼IDs, 2)
            End If
        End If
    End If
UnloadForm:
    mblnOk = True
    Unload Me
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub InitFace()
    Select Case mbytFun
    Case F_ͣ��
        Me.Caption = "ͣ��": Me.Height = 4350
        lblͣ��ԭ��.Caption = "ͣ��ԭ��"
        lblͣ��ʱ��.Caption = "ͣ��ʱ��"
        cbo����ҽ��.Visible = False
        dtpStart.Enabled = True: dtpEnd.Enabled = True
    Case F_ȡ��ͣ��
        Me.Caption = "ȡ��ͣ��": Me.Height = 4350
        lblͣ��ԭ��.Caption = "ͣ��ԭ��"
        lblͣ��ʱ��.Caption = "ͣ��ʱ��"
        cbo����ҽ��.Visible = False
        cboͣ��ԭ��.Enabled = False
        dtpStart.Enabled = False: dtpEnd.Enabled = False
    Case F_����
        Me.Caption = "����": Me.Height = 5250
        lblͣ��ԭ��.Caption = "����ԭ��"
        lblͣ��ʱ��.Caption = "����ʱ��"
        dtpStart.Enabled = True: dtpEnd.Enabled = True
    Case F_ȡ������
        Me.Caption = "ȡ������": Me.Height = 5250
        lblͣ��ԭ��.Caption = "����ԭ��"
        lblͣ��ʱ��.Caption = "����ʱ��"
        cboͣ��ԭ��.Enabled = False
        cbo����ҽ��.Enabled = False
        dtpStart.Enabled = False: dtpEnd.Enabled = False
        idkDoctor.Enabled = False
    End Select
    cmdHelp.Top = Me.ScaleHeight - cmdHelp.Height - 300
End Sub

Private Function InitData() As Boolean
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim lng��ԴId As Long, rsSignalSource As ADODB.Recordset
    Dim dtCurrent As Date
    
    Err = 0: On Error GoTo errHandle
    If mbytFun = F_ͣ�� Or mbytFun = F_���� Then
        strSQL = "Select ����, ����, ����, Nvl(ȱʡ��־,0) As ȱʡ From ����ͣ��ԭ��"
        Set mrsStopReason = zlDatabase.OpenSQLRecord(strSQL, "��ȡ����ͣ��ԭ��")
        With cboͣ��ԭ��
            .Clear
            Do While Not mrsStopReason.EOF
                .AddItem Nvl(mrsStopReason!����) & "-" & Nvl(mrsStopReason!����)
                If Val(Nvl(mrsStopReason!ȱʡ)) = 1 Then .ListIndex = .NewIndex
                mrsStopReason.MoveNext
            Loop
        End With
    End If
    
    If mbytFun = F_���� Or mbytFun = F_ȡ������ Then
        If mblnOnlyԺ��ҽ�� Then
            idkDoctor.IDkindStr = "ҽ��|ҽ��|0|0|0|0|0||0|0|0"
            idkDoctor.ToolTipText = "ֻ��ѡԺ�ڽ���ҽ��"
        Else
            idkDoctor.IDkindStr = "Ժ��ҽ��|Ժ��ҽ��|0|0|0|0|0||0|0|0;Ժ��ҽ��|Ժ��ҽ��|0|0|0|0||0|0|0"
        End If
    End If
    
    '���ذ�����Ϣ
    strSQL = "Select a.Id, a.��Դid, a.��������, a.�ϰ�ʱ��, a.��ʼʱ��, a.��ֹʱ��," & vbNewLine & _
            "        a.ͣ�￪ʼʱ��, a.ͣ����ֹʱ��, a.ͣ��ԭ��," & vbNewLine & _
            "        a.���￪ʼʱ��, a.������ֹʱ��, a.����ҽ��id, a.����ҽ������, b.רҵ����ְ��" & vbNewLine & _
            " From �ٴ������¼ A, ��Ա�� B" & vbNewLine & _
            " Where a.����ҽ��id = b.ID(+) And  a.Id = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng��¼ID)
    If rsTemp.BOF Then Exit Function
    With rsTemp
        lng��ԴId = Nvl(!��ԴID)
        txt��������.Text = Format(Nvl(!��������), "yyyy-mm-dd")
        zlControl.CboLocate cboͣ��ԭ��, Nvl(!ͣ��ԭ��)
        If cboͣ��ԭ��.ListIndex = -1 And Nvl(!ͣ��ԭ��) <> "" Then cboͣ��ԭ��.AddItem Nvl(!ͣ��ԭ��): cboͣ��ԭ��.ListIndex = cboͣ��ԭ��.NewIndex
        '����ҽ���ļ����Ƶ�����
        txt�ϰ�ʱ��.Text = Nvl(!�ϰ�ʱ��)
        If mbytFun = F_ȡ��ͣ�� Then
            dtpStart.Value = Format(Nvl(!ͣ�￪ʼʱ��, "00:00:00"), "hh:mm:ss")
            dtpEnd.Value = Format(Nvl(!ͣ����ֹʱ��, "00:00:00"), "hh:mm:ss")
        ElseIf mbytFun = F_ȡ������ Then
            dtpStart.Value = Format(Nvl(!���￪ʼʱ��, "00:00:00"), "hh:mm:ss")
            dtpEnd.Value = Format(Nvl(!������ֹʱ��, "00:00:00"), "hh:mm:ss")
        Else
            '���ͣ��/���ﵱǰ�Ѵ��ڹҺŵİ��ţ����Ե�ǰʱ��+1����Ϊȱʡʱ��.
            dtCurrent = zlDatabase.Currentdate
            If dtCurrent >= Nvl(!��ʼʱ��, "00:00:00") Then
                dtpStart.Value = Format(DateAdd("n", 1, dtCurrent), "hh:mm:ss")
            Else
                dtpStart.Value = Format(Nvl(!��ʼʱ��, "00:00:00"), "hh:mm:ss")
            End If
            dtpStart.Tag = Format(Nvl(!��ʼʱ��, "00:00:00"), "yyyy-MM-dd hh:mm:ss")
            dtpEnd.Value = Format(Nvl(!��ֹʱ��, "00:00:00"), "hh:mm:ss")
            dtpEnd.Tag = Format(Nvl(!��ֹʱ��, "00:00:00"), "yyyy-MM-dd hh:mm:ss")
        End If
    End With
    
    '��Դ��Ϣ
    strSQL = "Select a.����, a.����, a.����ID, b.���� As ����, c.���� As �շ���Ŀ, a.ҽ������," & vbNewLine & _
            "        Decode(Nvl(a.���տ���״̬, 0), 1, '����ԤԼ', 2, '��ֹԤԼ', 3, '�ܽڼ������ÿ���', '���ϰ�') As ���տ���," & vbNewLine & _
            "        Nvl(a.�Ƿ񽨲���, 0) As ����" & vbNewLine & _
            " From �ٴ������Դ A, ���ű� B, �շ���ĿĿ¼ C" & vbNewLine & _
            " Where a.����id = b.Id And a.��Ŀid = c.Id And a.Id = [1]"
    Set rsSignalSource = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng��ԴId)
    If rsSignalSource.BOF Then Exit Function
    With rsSignalSource
        txtSignalNO.Text = Nvl(!����)
        txt����.Text = Nvl(!����)
        txtDept.Text = Nvl(!����)
        txtDept.Tag = Nvl(!����ID)
        txtItem.Text = Nvl(!�շ���Ŀ)
        txtDoctor.Text = Nvl(!ҽ������)
        txt���տ���.Text = Nvl(!���տ���)
        chk����.Value = Val(Nvl(!����))
    End With
    
    '��������ҽ��
    If mbytFun = F_���� Or mbytFun = F_ȡ������ Then
        If mbytFun = F_���� Then
            Call LoadDoctor(Val(txtDept.Tag))
        Else
            With rsTemp
                If Val(Nvl(rsTemp!����ҽ��id)) = 0 Then idkDoctor.IDKind = 2
                zlControl.CboLocate cbo����ҽ��, Nvl(!����ҽ��id), True
                If cbo����ҽ��.ListIndex = -1 And Nvl(!����ҽ������) <> "" Then
                    cbo����ҽ��.AddItem Nvl(!����ҽ������) & IIf(Nvl(!רҵ����ְ��) = "", "", "(" & Nvl(!רҵ����ְ��) & ")")
                    cbo����ҽ��.ItemData(cbo����ҽ��.NewIndex) = Val(Nvl(!����ҽ��id))
                    cbo����ҽ��.ListIndex = cbo����ҽ��.NewIndex
                End If
            End With
        End If
    End If
    InitData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetDoctor(Optional ByVal lngSectID As Long = 0, Optional ByVal strWhere As String, _
    Optional ByVal blnNotEqualID As Boolean, _
    Optional ByVal blnGetSql As Boolean, Optional ByRef strSQL As String) As ADODB.Recordset
    '�õ�ָ�������µ�����ҽ��������
    '��Σ�
    '   blnNotEqualID - ������ID
    '   strWhere - ���ֻ�ܰ���"[1]"
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    
    strSQL = _
        "Select Distinct c.id,c.���,c.����,c.����,c.רҵ����ְ��" & vbNewLine & _
        " From ��Ա����˵�� a, ������Ա b ,��Ա�� c" & vbNewLine & _
        " Where b.��Աid=c.id And b.��Աid=a.��Աid  And  a.��Ա����=[2]" & vbNewLine & _
        "       And (c.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or c.����ʱ�� Is Null) " & vbNewLine & _
        "       And (c.վ��='" & gstrNodeNo & "' Or c.վ�� is Null)" & vbNewLine & _
                IIf(lngSectID = 0, "", IIf(blnNotEqualID, "  And b.����id <> [3]", "  And b.����id = [3]")) & vbNewLine & _
                strWhere & vbNewLine & _
        " Order By c.����"
        
    If blnGetSql Then
        strSQL = Replace(strSQL, "[2]", "'ҽ��'")
        strSQL = Replace(strSQL, "[3]", lngSectID)
        Exit Function
    End If
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlRegEvent", "", "ҽ��", lngSectID)
    Set GetDoctor = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("':��;��?��", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
End Sub

Private Sub Form_Load()
    mblnOnlyԺ��ҽ�� = Val(zlDatabase.GetPara("ֻ����ѡԺ��ҽ��", glngSys, mlngModule, "0")) = 1
    mbln����ҽ�������� = Val(zlDatabase.GetPara("����ҽ��������", glngSys, mlngModule, "0")) = 1
    mbytԤԼ�嵥���Ʒ�ʽ = Val(zlDatabase.GetPara("ԤԼ�嵥���Ʒ�ʽ", glngSys, mlngModule, "0"))
    mbytԤԼ�嵥��ӡ��ʽ = Val(zlDatabase.GetPara("ԤԼ�嵥��ӡ��ʽ", glngSys, mlngModule, "0"))
    
    Call InitFace
    If InitData() = False Then Unload Me: Exit Sub
    Call SetEnabledBackColor(Me.Controls)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mrsStopReason = Nothing
End Sub

Private Function SaveData() As Boolean
    Dim strSQL As String
    Dim arrTime As Variant
    Dim dtStartTime As Date, dtEndTime As Date
    Dim lngDoctor As Long
    
    Err = 0: On Error GoTo errHandle
    'ͣ��ʱ��
    If mbytFun = F_ͣ�� Or mbytFun = F_���� Then
        dtStartTime = GetWorkTrueDate(dtpStart.Tag, Format(dtpStart.Tag, "yyyy-mm-dd ") & Format(dtpStart.Value, "hh:mm:ss"), True, False)
        dtEndTime = GetWorkTrueDate(dtpStart.Tag, Format(dtpStart.Tag, "yyyy-mm-dd ") & Format(dtpEnd.Value, "hh:mm:ss"))
    End If
    
    Select Case mbytFun
    Case F_ͣ��
        'Zl_�ٴ������¼_Stopvisit
        strSQL = "Zl_�ٴ������¼_Stopvisit("
        '  ��¼id_In   Varchar2,
        strSQL = strSQL & "" & mlng��¼ID & ","
        '  ��ʼʱ��_In Varchar2 := Null,
        strSQL = strSQL & "To_date('" & Format(dtStartTime, "yyyy-MM-dd hh:mm:ss") & "','yyyy-mm-dd hh24:mi:ss'),"
        '  ��ֹʱ��_In Varchar2 := Null,
        strSQL = strSQL & "To_date('" & Format(dtEndTime, "yyyy-MM-dd hh:mm:ss") & "','yyyy-mm-dd hh24:mi:ss'),"
        '  ͣ��ԭ��_In Varchar2 := Null,
        strSQL = strSQL & "'" & NeedName(cboͣ��ԭ��.Text) & "',"
        '  ����Ա_In   Varchar2 := Null,
        strSQL = strSQL & "'" & UserInfo.���� & "',"
        '  ����ʱ��_In Varchar2 := Null,
        strSQL = strSQL & "To_date('" & Format(zlDatabase.Currentdate, "yyyy-MM-dd hh:mm:ss") & "','yyyy-mm-dd hh24:mi:ss'),"
        '  ȡ������_In Number:=0
        strSQL = strSQL & "" & 0 & ")"
    Case F_ȡ��ͣ��
        'Zl_�ٴ������¼_Stopvisit
        strSQL = "Zl_�ٴ������¼_Stopvisit("
        '  ��¼id_In   Varchar2,
        strSQL = strSQL & "" & mlng��¼ID & ","
        '  ��ʼʱ��_In Varchar2 := Null,
        strSQL = strSQL & "" & "NULL" & ","
        '  ��ֹʱ��_In Varchar2 := Null,
        strSQL = strSQL & "" & "NULL" & ","
        '  ͣ��ԭ��_In Varchar2 := Null,
        strSQL = strSQL & "" & "NULL" & ","
        '  ����Ա_In   Varchar2 := Null,
        strSQL = strSQL & "'" & UserInfo.���� & "',"
        '  ����ʱ��_In Varchar2 := Null,
        strSQL = strSQL & "To_date('" & Format(zlDatabase.Currentdate, "yyyy-MM-dd hh:mm:ss") & "','yyyy-mm-dd hh24:mi:ss'),"
        '  ȡ������_In Number:=0
        strSQL = strSQL & "" & 1 & ")"
    Case F_����
        If cbo����ҽ��.ListIndex <> -1 And mblnOnlyԺ��ҽ�� Then
            lngDoctor = cbo����ҽ��.ItemData(cbo����ҽ��.ListIndex)
        End If
        'Zl_�ٴ������¼_Replacedoctor
        strSQL = "Zl_�ٴ������¼_Replacedoctor("
        '  ��¼id_In       Varchar2,
        strSQL = strSQL & "" & mlng��¼ID & ","
        '  ��ʼʱ��_In     Varchar2 := Null,
        strSQL = strSQL & "To_date('" & Format(dtStartTime, "yyyy-MM-dd hh:mm:ss") & "','yyyy-mm-dd hh24:mi:ss'),"
        '  ��ֹʱ��_In Varchar2 := Null,
        strSQL = strSQL & "To_date('" & Format(dtEndTime, "yyyy-MM-dd hh:mm:ss") & "','yyyy-mm-dd hh24:mi:ss'),"
        '  ͣ��ԭ��_In Varchar2 := Null,
        strSQL = strSQL & "'" & NeedName(cboͣ��ԭ��.Text) & "',"
        '  ����ҽ��id_In   Number := Null,
        strSQL = strSQL & "" & ZVal(lngDoctor) & ","
        '  ����ҽ������_In Varchar2 := Null,
        strSQL = strSQL & "'" & GetReplaceDoctor(Trim(cbo����ҽ��.Text)) & "',"
        '  ����Ա����_In   �ٴ�����ͣ���¼.������%Type := Null,
        strSQL = strSQL & "'" & UserInfo.���� & "',"
        '  ����Ա���_In   ��Ա��.���%Type := Null,
        strSQL = strSQL & "'" & UserInfo.��� & "',"
        '  ����ʱ��_In Varchar2 := Null,
        strSQL = strSQL & "To_date('" & Format(zlDatabase.Currentdate, "yyyy-MM-dd hh:mm:ss") & "','yyyy-mm-dd hh24:mi:ss'),"
        '  ȡ������_In Number:=0
        strSQL = strSQL & "" & 0 & ")"
    Case F_ȡ������
        'Zl_�ٴ������¼_Replacedoctor
        strSQL = "Zl_�ٴ������¼_Replacedoctor("
        '  ��¼id_In       Varchar2,
        strSQL = strSQL & "" & mlng��¼ID & ","
        '  ��ʼʱ��_In     Varchar2 := Null,
        strSQL = strSQL & "" & "NULL" & ","
        '  ��ֹʱ��_In     Varchar2 := Null,
        strSQL = strSQL & "" & "NULL" & ","
        '  ͣ��ԭ��_In     Varchar2 := Null,
        strSQL = strSQL & "" & "NULL" & ","
        '  ����ҽ��id_In   Number := Null,
        strSQL = strSQL & "" & "NULL" & ","
        '  ����ҽ������_In Varchar2 := Null,
        strSQL = strSQL & "" & "NULL" & ","
        '  ����Ա����_In   �ٴ�����ͣ���¼.������%Type := Null,
        strSQL = strSQL & "'" & UserInfo.���� & "',"
        '  ����Ա���_In   ��Ա��.���%Type := Null,
        strSQL = strSQL & "'" & UserInfo.��� & "',"
        '  ����ʱ��_In Varchar2 := Null,
        strSQL = strSQL & "To_date('" & Format(zlDatabase.Currentdate, "yyyy-MM-dd hh:mm:ss") & "','yyyy-mm-dd hh24:mi:ss'),"
        '  ȡ������_In Number:=0
        strSQL = strSQL & "" & 1 & ")"
    End Select
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    SaveData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetReplaceDoctor(ByVal strIn As String, Optional strSplit As String = "(") As String
    '���������ҽ������
    GetReplaceDoctor = Mid(strIn, 1, IIf(InStr(strIn, strSplit) = 0, Len(strIn), InStr(strIn, strSplit) - 1))
End Function

Private Sub idkDoctor_ItemClick(index As Integer, objCard As zlIDKind.Card)
    Err = 0: On Error GoTo errHandle
    mblnOnlyԺ��ҽ�� = index = 1
    cbo����ҽ��.Clear
    If mblnOnlyԺ��ҽ�� = False Then Exit Sub
    
    Call LoadDoctor(Val(txtDept.Tag))
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LoadDoctor(Optional ByVal lng����ID As Long)
    '���ݿ���ID����ҽ��
    '˵����
    '   ����IDΪ0�Ǽ�������ҽ��
    Dim strPersons As String
    Dim rsDoctor As ADODB.Recordset
    
    Err = 0: On Error GoTo errHandle
    
    cbo����ҽ��.Clear
    Set rsDoctor = GetDoctor(lng����ID)
    If rsDoctor Is Nothing Then Exit Sub
    Do While Not rsDoctor.EOF
        If InStr("," & strPersons & ",", "," & Nvl(rsDoctor!ID) & ",") = 0 Then
            strPersons = strPersons & "," & Nvl(rsDoctor!ID)
            cbo����ҽ��.AddItem Nvl(rsDoctor!����) & IIf(Nvl(rsDoctor!רҵ����ְ��) = "", "", "(" & Nvl(rsDoctor!רҵ����ְ��) & ")")
            cbo����ҽ��.ItemData(cbo����ҽ��.NewIndex) = Val(Nvl(rsDoctor!ID))
        End If
        rsDoctor.MoveNext
    Loop
    If lng����ID <> 0 Then
        cbo����ҽ��.AddItem "��������ҽ��..."
        cbo����ҽ��.ItemData(cbo����ҽ��.NewIndex) = -1
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub idkDoctor_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

