VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmHandBackSearch 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����"
   ClientHeight    =   3675
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5055
   Icon            =   "frmHandBackSearch.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   5055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmdȷ�� 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   2160
      TabIndex        =   19
      Top             =   3120
      Width           =   1100
   End
   Begin VB.CommandButton cmdȡ�� 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3480
      TabIndex        =   18
      Top             =   3120
      Width           =   1100
   End
   Begin VB.Frame fraCondition 
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4815
      Begin VB.TextBox txtInput 
         Height          =   300
         Index           =   2
         Left            =   840
         TabIndex        =   16
         ToolTipText     =   "���������̱��롢���������"
         Top             =   2280
         Width           =   3480
      End
      Begin VB.CommandButton CmdSelecter 
         Caption         =   "��"
         Height          =   300
         Index           =   2
         Left            =   4320
         TabIndex        =   15
         Top             =   2280
         Width           =   255
      End
      Begin VB.TextBox txtInput 
         Height          =   300
         Index           =   1
         Left            =   840
         TabIndex        =   13
         ToolTipText     =   "���빩Ӧ�̱��롢���������"
         Top             =   1800
         Width           =   3480
      End
      Begin VB.CommandButton CmdSelecter 
         Caption         =   "��"
         Height          =   300
         Index           =   1
         Left            =   4320
         TabIndex        =   12
         Top             =   1800
         Width           =   255
      End
      Begin VB.CommandButton CmdSelecter 
         Caption         =   "��"
         Height          =   300
         Index           =   0
         Left            =   4320
         TabIndex        =   9
         Top             =   1320
         Width           =   255
      End
      Begin VB.TextBox txt����NO 
         Height          =   300
         Left            =   2970
         MaxLength       =   8
         TabIndex        =   2
         Top             =   360
         Width           =   1605
      End
      Begin VB.TextBox txt��ʼNo 
         Height          =   300
         Left            =   840
         MaxLength       =   8
         TabIndex        =   1
         Top             =   360
         Width           =   1605
      End
      Begin MSComCtl2.DTPicker dtp��ʼʱ�� 
         Height          =   315
         Left            =   840
         TabIndex        =   5
         Top             =   840
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy��MM��dd��"
         Format          =   166658051
         CurrentDate     =   36263
      End
      Begin MSComCtl2.DTPicker dtp����ʱ�� 
         Height          =   315
         Left            =   2970
         TabIndex        =   6
         Top             =   840
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy��MM��dd��"
         Format          =   166658051
         CurrentDate     =   36263
      End
      Begin VB.TextBox txtInput 
         Height          =   300
         Index           =   0
         Left            =   840
         TabIndex        =   10
         ToolTipText     =   "����ҩƷ���롢���������"
         Top             =   1320
         Width           =   3480
      End
      Begin VB.Label lblInputTxt 
         AutoSize        =   -1  'True
         Caption         =   "������"
         Height          =   180
         Index           =   2
         Left            =   120
         TabIndex        =   17
         Top             =   2340
         Width           =   540
      End
      Begin VB.Label lblInputTxt 
         AutoSize        =   -1  'True
         Caption         =   "��Ӧ��"
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   14
         Top             =   1860
         Width           =   540
      End
      Begin VB.Label lblInputTxt 
         AutoSize        =   -1  'True
         Caption         =   "ҩƷ"
         Height          =   180
         Index           =   0
         Left            =   300
         TabIndex        =   11
         Top             =   1380
         Width           =   360
      End
      Begin VB.Label lbl�� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "��"
         Height          =   180
         Index           =   0
         Left            =   2640
         TabIndex        =   8
         Top             =   900
         Width           =   180
      End
      Begin VB.Label lblʱ�� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   180
         Left            =   300
         TabIndex        =   7
         Top             =   900
         Width           =   360
      End
      Begin VB.Label lbl�� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "��"
         Height          =   180
         Index           =   1
         Left            =   2640
         TabIndex        =   4
         Top             =   420
         Width           =   180
      End
      Begin VB.Label LblNO 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "No"
         Height          =   180
         Left            =   480
         TabIndex        =   3
         Top             =   420
         Width           =   180
      End
   End
End
Attribute VB_Name = "frmHandBackSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngMode As Long            '0-δ���;1-�����
Private mfrmMain As Form            '������
Private mblnChange As Boolean
Private mlng�ⷿID As Long

Private Enum InputType
    ҩƷ = 0
    ��Ӧ�� = 1
    ������ = 2
End Enum

Private Type Type_SQLCondition
    strNO��ʼ As String
    strNO���� As String
    str����ʱ�俪ʼ As String
    str����ʱ����� As String
    str���ʱ�俪ʼ As String
    str���ʱ����� As String
    lngҩƷID As Long
    lng��Ӧ��ID As Long
    str������ As String
End Type

Private SQLCondition As Type_SQLCondition
Private Sub CmdSelecter_Click(Index As Integer)
    Dim RecReturn As ADODB.Recordset
    
    If Index = InputType.ҩƷ Then
        
        Call SetSelectorRS(1, "ҩƷ�⹺������", mlng�ⷿID, mlng�ⷿID, , , , True)
        
'        Set RecReturn = FrmҩƷѡ����.ShowME(Me, 1, 0, mlng�ⷿID, mlng�ⷿID)
        Set RecReturn = frmSelector.showMe(Me, 0, 1, , , , mlng�ⷿID, mlng�ⷿID, mlng�ⷿID, , , , , 2, False)
        
        If RecReturn.RecordCount = 0 Then
            Call zlControl.TxtSelAll(txtInput(Index))
            Exit Sub
        End If
            
        If gintҩƷ������ʾ = 1 Then
            txtInput(Index).Text = "[" & RecReturn!ҩƷ���� & "]" & IIf(IsNull(RecReturn!��Ʒ��), RecReturn!ͨ����, RecReturn!��Ʒ��)
        Else
            txtInput(Index).Text = "[" & RecReturn!ҩƷ���� & "]" & RecReturn!ͨ����
        End If
        txtInput(Index).Tag = RecReturn!ҩƷid
    Else
        If GetTxtInputReturn(Index, txtInput(Index), "") = False Then
            Call zlControl.TxtSelAll(txtInput(Index))
        End If
    End If
End Sub


Private Sub Cmdȡ��_Click()
    Unload Me
End Sub

Private Sub Cmdȷ��_Click()
    If Len(txt��ʼNo.Text) = 8 Then
        SQLCondition.strNO��ʼ = txt��ʼNo.Text
    End If
    
    If Len(txt����NO.Text) = 8 Then
        SQLCondition.strNO���� = txt����NO.Text
    End If

    If mlngMode = 0 Then
        SQLCondition.str����ʱ�俪ʼ = Format(dtp��ʼʱ��.Value, "yyyy-mm-dd") & " 00:00:00"
        SQLCondition.str����ʱ����� = Format(dtp����ʱ��.Value, "yyyy-mm-dd") & " 23:59:59"
    Else
        SQLCondition.str���ʱ�俪ʼ = Format(dtp��ʼʱ��.Value, "yyyy-mm-dd") & " 00:00:00"
        SQLCondition.str���ʱ����� = Format(dtp����ʱ��.Value, "yyyy-mm-dd") & " 23:59:59"
    End If
    
    SQLCondition.lngҩƷID = Val(txtInput(InputType.ҩƷ).Tag)
    SQLCondition.lng��Ӧ��ID = Val(txtInput(InputType.��Ӧ��).Tag)
    SQLCondition.str������ = txtInput(InputType.������).Text
    
    mblnChange = True
    
    Unload Me
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Call ReleaseSelectorRS
End Sub

Private Sub txtInput_GotFocus(Index As Integer)
    zlControl.TxtSelAll txtInput(Index)
End Sub

Public Function GetSearch(ByVal FrmMain As Form, ByVal lngMode As Long, ByVal lng�ⷿID As Long, _
        ByRef strNO��ʼ As String, _
        ByRef strNO���� As String, _
        ByRef str����ʱ�俪ʼ As String, _
        ByRef str����ʱ����� As String, _
        ByRef str���ʱ�俪ʼ As String, _
        ByRef str���ʱ����� As String, _
        ByRef lngҩƷID As Long, _
        ByRef lng��Ӧ��ID As Long, _
        ByRef str������ As String) As Boolean
    
    mblnChange = False
    mlngMode = lngMode
    mlng�ⷿID = lng�ⷿID
    Set mfrmMain = FrmMain
    
    If lngMode = 0 Then
        dtp��ʼʱ��.Value = CDate(str����ʱ�俪ʼ)
        dtp����ʱ��.Value = CDate(str����ʱ�����)
    Else
        dtp��ʼʱ��.Value = CDate(str���ʱ�俪ʼ)
        dtp����ʱ��.Value = CDate(str���ʱ�����)
    End If
    
    Me.Show vbModal, mfrmMain
    
    GetSearch = mblnChange
    
    strNO��ʼ = SQLCondition.strNO��ʼ
    strNO���� = SQLCondition.strNO����
    
    If lngMode = 0 Then
        str����ʱ�俪ʼ = SQLCondition.str����ʱ�俪ʼ
        str����ʱ����� = SQLCondition.str����ʱ�����
    Else
        str���ʱ�俪ʼ = SQLCondition.str���ʱ�俪ʼ
        str���ʱ����� = SQLCondition.str���ʱ�����
    End If
    
    lngҩƷID = SQLCondition.lngҩƷID
    lng��Ӧ��ID = SQLCondition.lng��Ӧ��ID
    str������ = SQLCondition.str������
End Function
Private Sub txtInput_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode <> vbKeyReturn Then Exit Sub
    If Trim(txtInput(Index).Text) = "" Then Exit Sub
    
    Dim sngLeft As Single
    Dim sngTop As Single
    Dim RecReturn As Recordset
    Dim strkey As String
    
    If Index = InputType.ҩƷ Then
        If KeyCode <> vbKeyReturn Then Exit Sub
        If Trim(txtInput(Index).Text) = "" Then Exit Sub
        sngLeft = Me.Left + fraCondition.Left + txtInput(Index).Left
        sngTop = Me.Top + fraCondition.Top + txtInput(Index).Top + txtInput(Index).Height + Me.Height - Me.ScaleHeight '  50
        If sngTop + 3630 > Screen.Height Then
            sngTop = sngTop - txtInput(Index).Height - 3630
        End If
        
        strkey = Trim(txtInput(Index).Text)
        If Mid(strkey, 1, 1) = "[" Then
            If InStr(2, strkey, "]") <> 0 Then
                strkey = Mid(strkey, 2, InStr(2, strkey, "]") - 2)
            Else
                strkey = Mid(strkey, 2)
            End If
        End If
        
        Call SetSelectorRS(1, "ҩƷ�⹺������", mlng�ⷿID, mlng�ⷿID, , , , True)
        
'        Set RecReturn = FrmҩƷ��ѡѡ����.ShowME(Me, 1, , mlng�ⷿID, mlng�ⷿID, strkey, sngLeft, sngTop)
        Set RecReturn = frmSelector.showMe(Me, 1, 1, strkey, sngLeft, sngTop, mlng�ⷿID, mlng�ⷿID, mlng�ⷿID, , , , , 2, False)
        
        If RecReturn.RecordCount = 0 Then
            Call zlControl.TxtSelAll(txtInput(Index))
            Exit Sub
        End If
        
        If gintҩƷ������ʾ = 1 Then
            txtInput(Index).Text = "[" & RecReturn!ҩƷ���� & "]" & IIf(IsNull(RecReturn!��Ʒ��), RecReturn!ͨ����, RecReturn!��Ʒ��)
        Else
            txtInput(Index).Text = "[" & RecReturn!ҩƷ���� & "]" & RecReturn!ͨ����
        End If
        txtInput(Index).Tag = RecReturn!ҩƷid
    Else
        If GetTxtInputReturn(Index, txtInput(Index), Trim(txtInput(Index).Text)) = False Then
            Call zlControl.TxtSelAll(txtInput(Index))
        End If
    End If
End Sub

Private Function GetTxtInputReturn(ByVal intType As Integer, ByVal txtObj As TextBox, ByVal strkey As String) As Boolean
    Dim vRect As RECT
    Dim lngH As Long
    Dim strReturn As String
    
    vRect = zlControl.GetControlRect(txtObj.hWnd)
    lngH = txtObj.Height
    vRect.Left = vRect.Left - 15
    
    strReturn = SelectInput(intType, Trim(strkey), vRect.Left, vRect.Top, lngH)
    
    If strReturn = "" Then Exit Function
        
    txtObj.Tag = Val(Split(strReturn, ";")(0))
    txtObj.Text = Split(strReturn, ";")(1)
    
    GetTxtInputReturn = True
End Function

Private Function SelectInput(ByVal intType As Integer, ByVal strkey As String, ByVal sngX As Single, ByVal sngY As Single, ByVal sngH As Single) As String
    'ѡ������֧�ֶ�ҩƷ����Ӧ�̡������̵�ѡ��
    'intType��0-ҩƷ;1-��Ӧ��;2-������
    'strKey����-ȫ��;�ǿ�-ģ��ƥ��
    'SelectInput����ֵ����-û�ҵ�ƥ���¼;
    '                 �ǿ�-ҩƷ��ҩƷID;ҩƷ����;���;��λ;��װ��
    '                     -��Ӧ�̣���Ӧ��ID;��Ӧ�����ƣ�
    '                     -�����̣�������ID;���������ƣ�
    
    Dim rsTemp As ADODB.Recordset
    Dim blnCancel As Boolean
    Dim strSubUnit As String
    Dim strFindString As String
    Dim strReturn As String
    Dim strSqlҩƷ As String
    
    Err = 0: On Error GoTo ErrHand:
    
    strkey = UCase(Trim(strkey))
    
    Select Case intType
    Case InputType.ҩƷ
        If strkey <> "" Then
            strFindString = " And (B.���� Like [1] OR B.���� Like [2] OR C.���� LIKE [2])"
            If IsNumeric(strkey) Then                         '10,11.����ȫ������ʱֻƥ�����'����ҩƷ,��Ҫƥ�����(����Ϊ3��������)
                If Mid(gtype_UserSysParms.P44_����ƥ��, 1, 1) = "1" Then strFindString = " And (B.���� Like [1] Or B.���� Like [2] And C.����=3)"
            ElseIf zlStr.IsCharAlpha(strkey) Then         '01,11.����ȫ����ĸʱֻƥ�����
                If Mid(gtype_UserSysParms.P44_����ƥ��, 2, 1) = "1" Then strFindString = " And C.���� Like [2] "
            ElseIf zlStr.IsCharChinese(strkey) Then
                strFindString = " And B.���� Like [2] "
            End If
        End If
        
        If strkey = "" Then
            If gintҩƷ������ʾ = 0 Then
                strSqlҩƷ = ",'['||����||']'|| ͨ���� As ҩƷ����"
            ElseIf gintҩƷ������ʾ = 1 Then
                strSqlҩƷ = ",'['||����||']'|| Nvl(��Ʒ��,ͨ����) As ҩƷ����"
            ElseIf gintҩƷ������ʾ = 2 Then
                strSqlҩƷ = ",'['||����||']'|| ͨ���� As ҩƷ����,��Ʒ��"
            End If
        Else
            strSqlҩƷ = ",'['||����||']'|| �������� As ҩƷ����"
        End If
        
        gstrSQL = "Select Rownum As ID, ҩƷid" & strSqlҩƷ & ",���,���� as ������,��Ʒ�� " & _
            " From (Select Distinct A.ҩƷid, B.����, B.��������, B.���� As ͨ����,C.���� As ��Ʒ��, B.���,B.���� " & _
            " From ҩƷ��� A, " & _
            " (Select B.ID, B.����, B.����, B.���,B.����, C.���� As �������� From �շ���ĿĿ¼ B, �շ���Ŀ���� C " & _
            " Where (B.վ�� = [3] Or B.վ�� is Null) And B.ID = C.�շ�ϸĿid And B.��� In ('5', '6', '7') " & strFindString & ") B, �շ���Ŀ���� C " & _
            " Where A.ҩƷid = B.ID And A.ҩƷid = C.�շ�ϸĿid(+) And C.����(+) = 3 "

        gstrSQL = gstrSQL & " Order By B.����)"
        
        Set rsTemp = zlDataBase.ShowSQLSelect(Me, gstrSQL, 0, "ҩƷѡ����", False, "", "ѡ��ҩƷ", False, False, True, sngX, sngY, sngH, blnCancel, False, False, _
                        strkey & "%", "%" & strkey & "%", _
                        gstrNodeNo)
        
        If blnCancel = True Then Exit Function
        
        If rsTemp Is Nothing Then
            strReturn = ""
        Else
            strReturn = rsTemp!ҩƷid & ";" & rsTemp!ҩƷ����
        End If
    Case InputType.��Ӧ��
        gstrSQL = "Select id,����,����,���� From ��Ӧ�� " & _
                  "Where (վ�� = [3] Or վ�� is Null) " & _
                  "  And (To_Char(����ʱ��,'yyyy-MM-dd')='3000-01-01' or ����ʱ�� is null ) And ĩ��=1 " & _
                  "  And (substr(����,1,1)=1 Or Nvl(ĩ��,0)=0) And (���� like [1] or ���� like [2] or ���� like [2])"
        Set rsTemp = zlDataBase.ShowSQLSelect(Me, gstrSQL, 0, "��Ӧ��ѡ����", False, "", "ѡ��Ӧ��", False, False, True, sngX, sngY, sngH, blnCancel, False, False, _
                        strkey & "%", "%" & strkey & "%", _
                        gstrNodeNo)
        
        If blnCancel = True Then Exit Function
        
        If rsTemp Is Nothing Then
            MsgBox "����ֵ��Ч��", vbInformation, gstrSysName
            strReturn = ""
        Else
            strReturn = rsTemp!id & ";" & rsTemp!����
        End If
    Case InputType.������
        gstrSQL = "Select Rownum As ID,����,����,���� From ҩƷ������ " & _
                  "Where (վ�� = [3] Or վ�� is Null) And (���� like [1] Or ���� like [2] Or ���� like [2]) " & _
                  "Order By ����"
        Set rsTemp = zlDataBase.ShowSQLSelect(Me, gstrSQL, 0, "������ѡ����", False, "", "ѡ��������", False, False, True, sngX, sngY, sngH, blnCancel, False, False, _
                        strkey & "%", "%" & strkey & "%", _
                        gstrNodeNo)
        
        If blnCancel = True Then Exit Function
        
        If rsTemp Is Nothing Then
            MsgBox "����ֵ��Ч��", vbInformation, gstrSysName
            strReturn = ""
        Else
            strReturn = rsTemp!id & ";" & rsTemp!����
        End If
    End Select
    
    SelectInput = strReturn
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function



Private Sub txt����NO_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    Call txt����NO_Validate(True)
End Sub
Private Sub txt����NO_Validate(Cancel As Boolean)
    If IsNumeric(txt����NO.Text) Then
        txt����NO.Text = zlCommFun.GetFullNO(txt����NO.Text, 92)
    End If
End Sub


Private Sub txt��ʼNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    Call txt��ʼNo_Validate(True)
End Sub
Private Sub txt��ʼNo_Validate(Cancel As Boolean)
    If IsNumeric(txt��ʼNo.Text) Then
        txt��ʼNo.Text = zlCommFun.GetFullNO(txt��ʼNo.Text, 92)
    End If
End Sub


