VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPersonLoanFileter 
   BorderStyle     =   0  'None
   Caption         =   "��������"
   ClientHeight    =   3870
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4050
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   4050
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton cmdˢ�� 
      Caption         =   "����(&F)"
      Height          =   390
      Left            =   2700
      TabIndex        =   28
      Top             =   3105
      Width           =   1050
   End
   Begin VB.PictureBox picRequisition 
      BorderStyle     =   0  'None
      Height          =   2940
      Index           =   0
      Left            =   75
      ScaleHeight     =   2940
      ScaleWidth      =   3855
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   90
      Width           =   3855
      Begin VB.CheckBox chkDate 
         Caption         =   "��ȷ�ϵĽ��"
         Height          =   375
         Index           =   1
         Left            =   75
         TabIndex        =   4
         Top             =   825
         Width           =   1665
      End
      Begin VB.CheckBox chkDate 
         Caption         =   "��ȷ�ϵĽ��"
         Height          =   375
         Index           =   0
         Left            =   75
         TabIndex        =   0
         Top             =   0
         Value           =   1  'Checked
         Width           =   1665
      End
      Begin VB.TextBox txtEdit 
         Height          =   330
         Index           =   0
         Left            =   615
         TabIndex        =   13
         Top             =   2490
         Width           =   3105
      End
      Begin MSComCtl2.DTPicker dtpStartDate 
         Height          =   315
         Index           =   0
         Left            =   615
         TabIndex        =   1
         Top             =   375
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   127926275
         CurrentDate     =   37007
      End
      Begin MSComCtl2.DTPicker dtpEndDate 
         Height          =   315
         Index           =   0
         Left            =   2430
         TabIndex        =   3
         Top             =   375
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   127926275
         CurrentDate     =   37007
      End
      Begin MSComCtl2.DTPicker dtpStartDate 
         Height          =   315
         Index           =   1
         Left            =   615
         TabIndex        =   5
         Top             =   1185
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   127926275
         CurrentDate     =   37007
      End
      Begin MSComCtl2.DTPicker dtpEndDate 
         Height          =   315
         Index           =   1
         Left            =   2430
         TabIndex        =   7
         Top             =   1185
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   127926275
         CurrentDate     =   37007
      End
      Begin MSComCtl2.DTPicker dtpStartDate 
         Height          =   315
         Index           =   2
         Left            =   615
         TabIndex        =   9
         Top             =   1935
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   127926275
         CurrentDate     =   37007
      End
      Begin MSComCtl2.DTPicker dtpEndDate 
         Height          =   315
         Index           =   2
         Left            =   2430
         TabIndex        =   11
         Top             =   1935
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   127926275
         CurrentDate     =   37007
      End
      Begin VB.CheckBox chkDate 
         Caption         =   "��ȡ��ȷ�ϵĽ��"
         Height          =   375
         Index           =   2
         Left            =   75
         TabIndex        =   8
         Top             =   1575
         Width           =   2025
      End
      Begin VB.Label lbl�� 
         AutoSize        =   -1  'True
         Caption         =   "��"
         Height          =   180
         Index           =   0
         Left            =   2070
         TabIndex        =   2
         Top             =   435
         Width           =   180
      End
      Begin VB.Label lbl�� 
         AutoSize        =   -1  'True
         Caption         =   "��"
         Height          =   180
         Index           =   1
         Left            =   2070
         TabIndex        =   6
         Top             =   1245
         Width           =   180
      End
      Begin VB.Label lbl�� 
         AutoSize        =   -1  'True
         Caption         =   "��"
         Height          =   180
         Index           =   2
         Left            =   2070
         TabIndex        =   10
         Top             =   1995
         Width           =   180
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "�����"
         Height          =   180
         Index           =   0
         Left            =   0
         TabIndex        =   12
         Top             =   2565
         Width           =   540
      End
   End
   Begin VB.PictureBox picRequisition 
      BorderStyle     =   0  'None
      Height          =   2865
      Index           =   1
      Left            =   75
      ScaleHeight     =   2865
      ScaleWidth      =   3885
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   180
      Width           =   3885
      Begin VB.TextBox txtEdit 
         Height          =   330
         Index           =   1
         Left            =   615
         TabIndex        =   27
         Top             =   2415
         Width           =   3105
      End
      Begin VB.CheckBox chkOutDate 
         Caption         =   "���������ڲ���"
         Height          =   375
         Index           =   0
         Left            =   75
         TabIndex        =   14
         Top             =   0
         Width           =   1665
      End
      Begin VB.CheckBox chkOutDate 
         Caption         =   "��������ڲ���"
         Height          =   375
         Index           =   1
         Left            =   75
         TabIndex        =   18
         Top             =   825
         Value           =   1  'Checked
         Width           =   1665
      End
      Begin VB.CheckBox chkOutDate 
         Caption         =   "��ȡ�����ڲ���"
         Height          =   375
         Index           =   2
         Left            =   75
         TabIndex        =   22
         Top             =   1575
         Width           =   1665
      End
      Begin MSComCtl2.DTPicker dtpOutStartDate 
         Height          =   315
         Index           =   0
         Left            =   615
         TabIndex        =   15
         Top             =   375
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   127926275
         CurrentDate     =   37007
      End
      Begin MSComCtl2.DTPicker dtpOutEndDate 
         Height          =   315
         Index           =   0
         Left            =   2430
         TabIndex        =   17
         Top             =   375
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   127926275
         CurrentDate     =   37007
      End
      Begin MSComCtl2.DTPicker dtpOutStartDate 
         Height          =   315
         Index           =   1
         Left            =   615
         TabIndex        =   19
         Top             =   1185
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   127926275
         CurrentDate     =   37007
      End
      Begin MSComCtl2.DTPicker dtpOutEndDate 
         Height          =   315
         Index           =   1
         Left            =   2430
         TabIndex        =   21
         Top             =   1185
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   127926275
         CurrentDate     =   37007
      End
      Begin MSComCtl2.DTPicker dtpOutStartDate 
         Height          =   315
         Index           =   2
         Left            =   615
         TabIndex        =   23
         Top             =   1935
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   127926275
         CurrentDate     =   37007
      End
      Begin MSComCtl2.DTPicker dtpOutEndDate 
         Height          =   315
         Index           =   2
         Left            =   2430
         TabIndex        =   25
         Top             =   1935
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   127926275
         CurrentDate     =   37007
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "�����"
         Height          =   180
         Index           =   1
         Left            =   0
         TabIndex        =   26
         Top             =   2490
         Width           =   540
      End
      Begin VB.Label lbl�� 
         AutoSize        =   -1  'True
         Caption         =   "��"
         Height          =   180
         Index           =   5
         Left            =   2070
         TabIndex        =   24
         Top             =   1995
         Width           =   180
      End
      Begin VB.Label lbl�� 
         AutoSize        =   -1  'True
         Caption         =   "��"
         Height          =   180
         Index           =   4
         Left            =   2070
         TabIndex        =   20
         Top             =   1245
         Width           =   180
      End
      Begin VB.Label lbl�� 
         AutoSize        =   -1  'True
         Caption         =   "��"
         Height          =   180
         Index           =   3
         Left            =   2070
         TabIndex        =   16
         Top             =   435
         Width           =   180
      End
   End
End
Attribute VB_Name = "frmPersonLoanFileter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
Option Explicit
Private mArrFilter As Variant
Private mblnRequisition As Boolean   'true-�ҵĽ���¼,false-�ҵĽ����¼
Private mstrPrivs As String
Private mlngModule As Long
Private Enum mTxtIdx
    idx_����� = 1
    idx_����� = 0
End Enum
Private mblnRequisitionChange As Boolean   '�ı����ҵĽ���¼����
Private mblnOutPayChange As Boolean   '�ı����ҵĽ����¼����


'--------------------------------------------------------------------------------------------------------
Public Event zlRefreshCon(ByVal arrFilter As Variant, ByVal blnRequisition As Boolean)

Private Function GetFilter() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:��ȡ������Ϣ
    '���:
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-04-30 11:52:50
    '-----------------------------------------------------------------------------------------------------------
    Dim cllFilter As Collection, strReg As String
    
    '������ѯ����
    Set cllFilter = New Collection
    
    If chkDate(0).Value = 1 Then
        cllFilter.Add Array(Format(dtpStartDate(0).Value, "yyyy-mm-dd") & " 00:00:00", Format(dtpEndDate(0).Value, "yyyy-mm-dd") & " 23:59:59"), "����ʱ��"
    Else
        cllFilter.Add Array("1901-01-01", "1901-01-01"), "����ʱ��"
    End If
    
    If chkDate(1).Value = 1 Then
        cllFilter.Add Array(Format(dtpStartDate(1).Value, "yyyy-mm-dd") & " 00:00:00", Format(dtpEndDate(1).Value, "yyyy-mm-dd") & " 23:59:59"), "���ʱ��"
    Else
        cllFilter.Add Array("1901-01-01", "1901-01-01"), "���ʱ��"
    End If
    
    If chkDate(2).Value = 1 Then
        cllFilter.Add Array(Format(dtpStartDate(2).Value, "yyyy-mm-dd") & " 00:00:00", Format(dtpEndDate(2).Value, "yyyy-mm-dd") & " 23:59:59"), "ȡ��ʱ��"
    Else
        cllFilter.Add Array("1901-01-01", "1901-01-01"), "ȡ��ʱ��"
    End If
    
    
    If chkOutDate(0).Value = 1 Then
        cllFilter.Add Array(Format(dtpOutStartDate(0).Value, "yyyy-mm-dd") & " 00:00:00", Format(dtpOutEndDate(0).Value, "yyyy-mm-dd") & " 23:59:59"), "���-����ʱ��"
    Else
        cllFilter.Add Array("1901-01-01", "1901-01-01"), "���-����ʱ��"
    End If
    
    If chkOutDate(1).Value = 1 Then
        cllFilter.Add Array(Format(dtpOutStartDate(1).Value, "yyyy-mm-dd") & " 00:00:00", Format(dtpOutEndDate(1).Value, "yyyy-mm-dd") & " 23:59:59"), "���-���ʱ��"
    Else
        cllFilter.Add Array("1901-01-01", "1901-01-01"), "���-���ʱ��"
    End If
    
    If chkOutDate(2).Value = 1 Then
        cllFilter.Add Array(Format(dtpOutStartDate(2).Value, "yyyy-mm-dd") & " 00:00:00", Format(dtpOutEndDate(2).Value, "yyyy-mm-dd") & " 23:59:59"), "���-ȡ��ʱ��"
    Else
        cllFilter.Add Array("1901-01-01", "1901-01-01"), "���-ȡ��ʱ��"
    End If
    cllFilter.Add Trim(txtEdit(mTxtIdx.idx_�����)), "�����"
    cllFilter.Add Trim(txtEdit(mTxtIdx.idx_�����)), "�����"
    Set mArrFilter = cllFilter
    
End Function

 
Private Sub cmdˢ��_Click()
    Call GetFilter
    RaiseEvent zlRefreshCon(mArrFilter, blnRequisition)
    If blnRequisition Then
        mblnRequisitionChange = False
    Else
        mblnOutPayChange = False
    End If
End Sub

Private Sub InitData()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ������
    '���:intType:0-�ҵĽ���¼����;1-�ҵĽ����¼����
    '����:���˺�
    '����:2009-09-09 14:41:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    dtpEndDate(0).MaxDate = Format(zlDatabase.Currentdate, "yyyy-mm-dd")
    dtpEndDate(1).MaxDate = dtpEndDate(0).MaxDate
    dtpEndDate(2).MaxDate = dtpEndDate(0).MaxDate

    dtpEndDate(0).Value = dtpEndDate(0).MaxDate
    dtpEndDate(1).Value = dtpEndDate(0).MaxDate
    dtpEndDate(2).Value = dtpEndDate(0).MaxDate
    
    dtpStartDate(0).Value = Format(DateAdd("m", -1, zlDatabase.Currentdate), "yyyy-mm-dd")
    dtpStartDate(1).Value = dtpStartDate(0).Value
    dtpStartDate(2).Value = dtpStartDate(0).Value



    dtpOutEndDate(0).MaxDate = dtpEndDate(0).MaxDate
    dtpOutEndDate(1).MaxDate = dtpOutEndDate(0).MaxDate
    dtpOutEndDate(2).MaxDate = dtpOutEndDate(0).MaxDate

    dtpOutEndDate(0).Value = dtpOutEndDate(0).MaxDate
    dtpOutEndDate(1).Value = dtpOutEndDate(0).MaxDate
    dtpOutEndDate(2).Value = dtpOutEndDate(0).MaxDate
    
    dtpOutStartDate(0).Value = dtpStartDate(0).Value
    dtpOutStartDate(1).Value = dtpOutStartDate(0).Value
    dtpOutStartDate(2).Value = dtpOutStartDate(0).Value
End Sub
 
Private Sub chkDate_Click(Index As Integer)
    dtpStartDate(Index).Enabled = chkDate(Index).Value = 1
    dtpEndDate(Index).Enabled = chkDate(Index).Value = 1
    mblnRequisitionChange = True
End Sub

Private Sub chkDate_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub dtpEndDate_Change(Index As Integer)
     If dtpEndDate(Index).Value > dtpStartDate(Index).MaxDate Then dtpEndDate(Index).Value = dtpStartDate(Index).MaxDate
    
    If dtpEndDate(Index).Value < dtpStartDate(Index).Value Then
        dtpStartDate(Index).Value = dtpEndDate(Index).Value
    End If
End Sub
Private Sub dtpEndDate_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub dtpOutEndDate_Change(Index As Integer)
    If dtpOutEndDate(Index).Value > dtpOutStartDate(Index).MaxDate Then dtpOutEndDate(Index).Value = dtpOutStartDate(Index).MaxDate
    If dtpOutEndDate(Index).Value < dtpOutStartDate(Index).Value Then
        dtpOutStartDate(Index).Value = dtpOutEndDate(Index).Value
    End If
End Sub

Private Sub dtpOutStartDate_Change(Index As Integer)
    mblnOutPayChange = True
    If dtpOutStartDate(Index).Value > dtpOutEndDate(Index).MaxDate Then dtpOutStartDate(Index).Value = dtpOutEndDate(Index).MaxDate
    If dtpOutEndDate(Index).Value < dtpOutStartDate(Index).Value Then
        dtpOutEndDate(Index).Value = dtpOutStartDate(Index).Value
    End If
End Sub

Private Sub dtpStartDate_Change(Index As Integer)
    mblnRequisitionChange = True
    If dtpStartDate(Index).Value > dtpOutEndDate(Index).MaxDate Then dtpStartDate(Index).Value = dtpEndDate(Index).MaxDate
    If dtpEndDate(Index).Value < dtpStartDate(Index).Value Then
        dtpEndDate(Index).Value = dtpStartDate(Index).Value
    End If
End Sub

Private Sub dtpStartDate_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub Form_Load()
    mstrPrivs = gstrPrivs: mlngModule = glngModul
    mblnRequisitionChange = True: mblnOutPayChange = True
End Sub

Private Sub Form_Resize()
        cmdˢ��.Left = Me.ScaleLeft + ScaleWidth - cmdˢ��.Width - 50
End Sub

Private Function Select��Աѡ����(ByVal objCtl As Control, ByVal strSearch As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:����ѡ����
    '���::objCtl-ָ���ؼ�
    '     strSearch-Ҫ����������
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-05-01 14:18:58
    '-----------------------------------------------------------------------------------------------------------
    Dim blnCancel As Boolean, strKey As String, strTittle As String, lngH As Long, strTemp As String
    Dim vRect As RECT
    Dim rsTemp  As ADODB.Recordset
    'zlDatabase.ShowSelect
    '���ܣ�
    '������
    '     frmParent=��ʾ�ĸ�����
    '     strSQL=������Դ,��ͬ����ѡ������SQL�е��ֶ��в�ͬҪ��
    '     bytStyle=ѡ�������
    '       Ϊ0ʱ:�б���:ID,��
    '       Ϊ1ʱ:���η��:ID,�ϼ�ID,����,����(���blnĩ��������Ҫĩ���ֶ�)
    '       Ϊ2ʱ:˫����:ID,�ϼ�ID,����,����,ĩ������ListViewֻ��ʾĩ��=1����Ŀ
    '     strTitle=ѡ������������,Ҳ���ڸ��Ի�����
    '     blnĩ��=������ѡ����(bytStyle=1)ʱ,�Ƿ�ֻ��ѡ��ĩ��Ϊ1����Ŀ
    '     strSeek=��bytStyle<>2ʱ��Ч,ȱʡ��λ����Ŀ��
    '             bytStyle=0ʱ,��ID���ϼ�ID֮��ĵ�һ���ֶ�Ϊ׼��
    '             bytStyle=1ʱ,�����Ǳ��������
    '     strNote=ѡ������˵������
    '     blnShowSub=��ѡ��һ���Ǹ����ʱ,�Ƿ���ʾ�����¼������е���Ŀ(��Ŀ��ʱ����)
    '     blnShowRoot=��ѡ������ʱ,�Ƿ���ʾ������Ŀ(��Ŀ��ʱ����)
    '     blnNoneWin,X,Y,txtH=����ɷǴ�����,X,Y,txtH��ʾ���ý�������������(�������Ļ)�͸߶�
    '     Cancel=���ز���,��ʾ�Ƿ�ȡ��,��Ҫ����blnNoneWin=Trueʱ
    '     blnMultiOne=��bytStyle=0ʱ,�Ƿ񽫶Զ�����ͬ��¼����һ���ж�
    '     blnSearch=�Ƿ���ʾ�к�,�����������кŶ�λ
    '���أ�ȡ��=Nothing,ѡ��=SQLԴ�ĵ��м�¼��
    '˵����
    '     1.ID���ϼ�ID����Ϊ�ַ�������
    '     2.ĩ�����ֶβ�Ҫ����ֵ
    'Ӧ�ã������ڸ������������������Ǻܴ��ѡ����,����ƥ���б�ȡ�
    
    strTittle = "��Աѡ����"
    vRect = zlcontrol.GetControlRect(objCtl.hwnd)
    lngH = objCtl.Height
    
  
    gstrSQL = "" & _
    "   Select Distinct B.ID, B.���, B.����, B.����, B.����, B.��������, B.�Ա�, B.�칫�ҵ绰 " & _
    "   From ��Ա����˵�� A, ��Ա�� B " & _
    "   Where A.��Աid = B.ID And A.��Ա���� In ('����Һ�Ա', '�����շ�Ա', 'Ԥ���տ�Ա', 'סԺ����Ա') " & _
    "         and (b.��� like upper([1]) or b.���� like [1] or b.���� like upper([1]) or b.���� like [1]) " & _
    "   Order By b.���"
    
    strKey = GetMatchingSting(strSearch, False)
    Set rsTemp = zlDatabase.ShowSQLSelect(Me, gstrSQL, 0, strTittle, False, "", "", False, False, True, vRect.Left - 15, vRect.Top, lngH, blnCancel, False, False, strKey)
 
    If blnCancel = True Then
       zlcontrol.ControlSetFocus objCtl, True
        Exit Function
    End If
    
    If rsTemp Is Nothing Then
        ShowMsgbox "û��������������Ա��Ϣ,����!"
       zlcontrol.ControlSetFocus objCtl, True
        Exit Function
    End If
    zlcontrol.ControlSetFocus objCtl, True
    objCtl.Text = Nvl(rsTemp!����)
    objCtl.Tag = Nvl(rsTemp!����)
    zlCommFun.PressKey vbKeyTab
    Select��Աѡ���� = True
End Function
Public Sub Init����()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ���������
    '����:���˺�
    '����:2009-09-09 14:59:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Call InitData
End Sub
Public Property Get GetFilterCon() As Variant
    Call GetFilter
    Set GetFilterCon = mArrFilter
    
End Property

Private Sub txtEdit_Change(Index As Integer)
    txtEdit(Index).Tag = ""
    If Index = mTxtIdx.idx_����� Then
        mblnOutPayChange = True
    Else
        mblnRequisitionChange = True
    End If
End Sub

Private Sub txtEdit_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If txtEdit(Index).Tag <> "" Then zlCommFun.PressKey vbKeyTab: Exit Sub
    If txtEdit(Index).Text = "" Then zlCommFun.PressKey vbKeyTab: Exit Sub
    If Select��Աѡ����(txtEdit(Index), Trim(txtEdit(Index).Text)) = False Then Exit Sub
End Sub
Private Sub chkOutDate_Click(Index As Integer)
    dtpOutStartDate(Index).Enabled = chkOutDate(Index).Value = 1
    dtpOutEndDate(Index).Enabled = chkOutDate(Index).Value = 1
    mblnOutPayChange = True
End Sub

Private Sub chkOutDate_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub
Private Sub dtpOutEndDate_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub
Private Sub dtpOutStartDate_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub
Public Property Get blnRequisition() As Boolean
    blnRequisition = mblnRequisition
End Property
Public Property Let blnRequisition(ByVal vNewValue As Boolean)
    mblnRequisition = vNewValue
    picRequisition(0).Visible = mblnRequisition
    picRequisition(1).Visible = Not mblnRequisition
End Property

Public Property Get IsMyRequistionConChange() As Boolean
   '���������˸ı�
   IsMyRequistionConChange = mblnRequisitionChange
End Property

Public Property Let IsMyRequistionConChange(ByVal vNewValue As Boolean)
    mblnRequisitionChange = vNewValue
End Property

Public Property Get IsMyOutPayConChange() As Boolean
   '���������˸ı�
   IsMyOutPayConChange = mblnOutPayChange
End Property

Public Property Let IsMyOutPayConChange(ByVal vNewValue As Boolean)
    mblnOutPayChange = vNewValue
End Property
Public Sub ReActionFilter(ByVal blnRequisition As Boolean)
    '���½ɻ����
    mblnRequisition = blnRequisition
    cmdˢ��_Click
End Sub
