VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDueFilter 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "��������"
   ClientHeight    =   3600
   ClientLeft      =   -15
   ClientTop       =   270
   ClientWidth     =   7335
   ControlBox      =   0   'False
   Icon            =   "frmDueFilter.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   7335
   StartUpPosition =   2  '��Ļ����
   Begin VB.Frame fra 
      Height          =   3270
      Left            =   120
      TabIndex        =   18
      Top             =   15
      Width           =   5610
      Begin VB.CheckBox chk����ʾǷ�� 
         Caption         =   "����ʾ����Ƿ��Ĳ���"
         Height          =   225
         Left            =   1665
         TabIndex        =   15
         Top             =   2910
         Width           =   2115
      End
      Begin VB.CommandButton cmd���� 
         Height          =   300
         Left            =   5100
         Picture         =   "frmDueFilter.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   14
         TabStop         =   0   'False
         ToolTipText     =   "����(F3)"
         Top             =   2535
         Width           =   330
      End
      Begin VB.TextBox txtUnit 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   1680
         MaxLength       =   20
         TabIndex        =   13
         Top             =   2520
         Width           =   3780
      End
      Begin VB.TextBox Txt���� 
         Height          =   300
         Left            =   1680
         MaxLength       =   64
         TabIndex        =   5
         Top             =   1065
         Width           =   3750
      End
      Begin VB.TextBox txtסԺ�� 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1680
         MaxLength       =   18
         TabIndex        =   7
         Top             =   1425
         Width           =   3750
      End
      Begin VB.TextBox txtInvoice 
         Height          =   300
         Left            =   1680
         TabIndex        =   11
         Top             =   2145
         Width           =   3750
      End
      Begin VB.TextBox txtNO 
         Height          =   300
         Left            =   1680
         MaxLength       =   8
         TabIndex        =   9
         Top             =   1785
         Width           =   3750
      End
      Begin MSComCtl2.DTPicker dtpEnd 
         Height          =   300
         Left            =   1680
         TabIndex        =   3
         Top             =   720
         Width           =   2070
         _ExtentX        =   3651
         _ExtentY        =   529
         _Version        =   393216
         CalendarTitleBackColor=   -2147483647
         CalendarTitleForeColor=   -2147483634
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   183959555
         CurrentDate     =   39083
      End
      Begin MSComCtl2.DTPicker dtpBegin 
         Height          =   300
         Left            =   1680
         TabIndex        =   1
         Top             =   300
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   529
         _Version        =   393216
         CalendarTitleBackColor=   -2147483647
         CalendarTitleForeColor=   -2147483634
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   183959555
         CurrentDate     =   39078
      End
      Begin VB.Label lbl��Լ��λ 
         Caption         =   "��Լ��λ(&H)"
         Height          =   210
         Left            =   600
         TabIndex        =   12
         Top             =   2580
         Width           =   1155
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����(&3)"
         Height          =   180
         Left            =   960
         TabIndex        =   4
         Top             =   1125
         Width           =   630
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "סԺ��(&4)"
         Height          =   180
         Left            =   780
         TabIndex        =   6
         Top             =   1485
         Width           =   810
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����Ʊ�ݺ�(&6)"
         Height          =   180
         Left            =   420
         TabIndex        =   10
         Top             =   2205
         Width           =   1170
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���ʵ��ݺ�(&5)"
         Height          =   180
         Left            =   420
         TabIndex        =   8
         Top             =   1845
         Width           =   1170
      End
      Begin VB.Label lblDateE 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���ʽ���ʱ��(&2)"
         Height          =   180
         Left            =   240
         TabIndex        =   2
         Top             =   780
         Width           =   1350
      End
      Begin VB.Label lblDateB 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���ʿ�ʼʱ��(&1)"
         Height          =   180
         Left            =   240
         TabIndex        =   0
         Top             =   360
         Width           =   1350
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   5970
      TabIndex        =   16
      Top             =   120
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   5970
      TabIndex        =   17
      Top             =   540
      Width           =   1100
   End
End
Attribute VB_Name = "frmDueFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������

Private Sub cmdCancel_Click()
    gblnOK = False
    Hide
End Sub

Private Sub cmdOK_Click()
    Dim DatTmp As Date
    If dtpBegin.Value > dtpEnd.Value Then
        DatTmp = dtpBegin.Value: dtpBegin.Value = dtpEnd.Value: dtpEnd.Value = DatTmp
    End If
    gblnOK = True
    Hide
End Sub

Private Sub cmd����_Click()
    If SelectUnits(txtUnit, "") = False Then Exit Sub
End Sub

Private Sub Form_Activate()
    dtpBegin.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    gblnOK = False
    txtInvoice.MaxLength = gbytFactLength
End Sub


Private Sub txtInvoice_GotFocus()
    zlcontrol.TxtSelAll txtInvoice
End Sub

Private Sub txtInvoice_Validate(Cancel As Boolean)
    txtInvoice.Text = Trim(txtInvoice.Text)
End Sub

Private Sub txtNO_GotFocus()
    zlcontrol.TxtSelAll txtNO
End Sub

Private Sub txtNO_Validate(Cancel As Boolean)
    txtNO.Text = Trim(txtNO.Text)
    If txtNO.Text <> "" Then txtNO.Text = GetFullNO(txtNO.Text, 15)
End Sub

Private Sub txtNO_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii <> 13 Then
        If Not (txtNO.Text = "" Or txtNO.SelLength = Len(txtNO.Text) Or txtNO.SelStart = 0) And _
            InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then
            KeyAscii = 0: Beep: Exit Sub
        End If
    End If
End Sub

Private Sub txtUnit_Change()
    txtUnit.Tag = ""
End Sub

Private Sub txtUnit_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If txtUnit.Tag <> "" Then zlCommFun.PressKey vbKeyTab: Exit Sub
    If txtUnit.Text = "" And txtUnit.Tag = "" Then zlCommFun.PressKey vbKeyTab: Exit Sub
    If SelectUnits(txtUnit, Trim(txtUnit.Text)) = False Then Exit Sub
End Sub
 

Private Sub txt����_GotFocus()
    zlcontrol.TxtSelAll txt����
End Sub

Private Sub txt����_Validate(Cancel As Boolean)
    txt����.Text = Replace(Trim(txt����.Text), "'", "")
End Sub

Private Sub txtסԺ��_GotFocus()
    zlcontrol.TxtSelAll txtסԺ��
End Sub

Private Sub txtסԺ��_Validate(Cancel As Boolean)
    txtסԺ��.Text = Trim(txtסԺ��.Text)
End Sub

Private Sub txtסԺ��_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Function SelectUnits(ByVal objCtl As Control, Optional strKey As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ѡ���Լ��λ
    '���:strKey-����ֵ
    '����:
    '����:
    '����:���˺�
    '����:2011-11-08 15:00:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSql As String
    Dim vRect As RECT, strWhere As String, bytStyle As Byte
    Dim sngX As Single, sngY As Single, lngH As Long
    Dim blnCancel As Boolean
 
    On Error GoTo errH
    bytStyle = 2
    strWhere = " Start with �ϼ�id is null Connect by prior ID=�ϼ�ID"
    If strKey <> "" Then
        strWhere = " Where 1=1 "
        If zlCommFun.IsCharChinese(strKey) Then
            strWhere = strWhere & " And ���� like [1]  Order by ����"
        ElseIf zlCommFun.IsCharAlpha(strKey) Then
            strWhere = strWhere & " And ���� like upper([1]) Order by ����"
        ElseIf zlCommFun.IsNumOrChar(strKey) Then
            strWhere = strWhere & " And ���� like upper([1])  Order by ����"
        Else
            strWhere = strWhere & " And  (���� like [1] or ���� like upper([1]) or ���� like upper([1])) Order by ����"
        End If
        bytStyle = 0
        strKey = gstrLike & strKey & "%"
    End If
    
    strSql = "" & _
    "   Select ID,�ϼ�ID,����,����,����,��ַ, ĩ��,˵��," & _
    "               To_Char(����ʱ��, 'YYYY-MM-DD HH24:MI') ����ʱ�� " & _
    "   From ��Լ��λ" & _
        strWhere
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strKey)
    'ShowSelect:
    '���ܣ��๦��ѡ����
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
    vRect = zlcontrol.GetControlRect(objCtl.hWnd)
    lngH = objCtl.Height
    sngX = vRect.Left - 15: sngY = vRect.Top
    Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSql, bytStyle, "��Լ��λѡ��", IIf(bytStyle = 2, True, False), "", "��ѡ����������ĺ�Լ��λ", IIf(bytStyle = 2, True, False), True, True, sngX, sngY, lngH, blnCancel, False, True, strKey)
    If blnCancel Then
        If objCtl.Enabled And objCtl.Visible Then objCtl.SetFocus
        zlcontrol.TxtSelAll objCtl
        Exit Function
    End If
    If rsTemp Is Nothing Then
        ShowMsgbox "�����ڷ��������ĺ�Լ��λ,����!"
        If objCtl.Enabled And objCtl.Visible Then objCtl.SetFocus
        zlcontrol.TxtSelAll objCtl
        Exit Function
        Exit Function
    End If
    If rsTemp.State <> 1 Then
        If objCtl.Enabled And objCtl.Visible Then objCtl.SetFocus
        zlcontrol.TxtSelAll objCtl
        Exit Function
        Exit Function
    End If
    With rsTemp
        objCtl.Text = Nvl(!����): objCtl.Tag = Nvl(!ID)
    End With
    If objCtl.Enabled And objCtl.Visible Then objCtl.SetFocus
    zlcontrol.TxtSelAll objCtl
    zlCommFun.PressKey vbKeyTab
    SelectUnits = True
    Exit Function
errH:
    If ErrCenter = 1 Then Resume
End Function
