VERSION 5.00
Begin VB.Form frmShortcutConfig 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��ݼ�����"
   ClientHeight    =   7170
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9990
   Icon            =   "frmShortcutConfig.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7170
   ScaleWidth      =   9990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton cmdDefault 
      Caption         =   "�ָ�Ĭ��(&D)"
      Height          =   400
      Left            =   4680
      TabIndex        =   5
      Top             =   6600
      Width           =   1215
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "ɾ ��(&D)"
      Height          =   400
      Left            =   6000
      TabIndex        =   3
      Top             =   6600
      Width           =   1215
   End
   Begin VB.CommandButton cmdSure 
      Caption         =   "ȷ ��(&S)"
      Height          =   400
      Left            =   7320
      TabIndex        =   2
      Top             =   6600
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ ��(&C)"
      Height          =   400
      Left            =   8640
      TabIndex        =   1
      Top             =   6600
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   6255
      Left            =   120
      TabIndex        =   0
      Top             =   160
      Width           =   9735
      Begin zl9PACSWork.ucFlexGrid ufgShoftcut 
         Height          =   5895
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   10398
         DefaultCols     =   ""
         IsCopyAdoMode   =   0   'False
         IsEjectConfig   =   -1  'True
         HeadFontCharset =   134
         HeadFontWeight  =   400
         DataFontCharset =   134
         DataFontWeight  =   400
      End
   End
End
Attribute VB_Name = "frmShortcutConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

#Const DebugState = False



Private Const C_STR_SHORTCUT_LIST_COLS As String = "|�˵�����,merge,read,w1800,txtcenter|ID,hide,key,uncfg|��Ŀ,hide,uncfg|ģ���,hide,uncfg|�˵�˵��,read,w2100,uncfg|���Ƽ�,hide,uncfg|�ַ���,hide,uncfg|Ĭ�ϼ�,hide,uncfg|��ݼ�>�����,read,w2100,uncfg|��ǰ��>�����,hide,uncfg|"
Private mstrProject As String
Private mlngMudule As Long

Public blnIsOk As Boolean

Private Declare Function GetKeyboardState Lib "user32" (pbKeyState As Byte) As Long
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer


Private Enum ShiftKeys
    AltKey = &H1
    CtrlKey = &H2
    ShiftKey = &H4
End Enum



Public Sub ShowShortcutConfig(ByVal strProject As String, ByVal lngModule As Long, owner As Object)
    mstrProject = strProject
    mlngMudule = lngModule
    
    blnIsOk = False
    
    Call Me.Show(1, owner)
End Sub

Private Sub cmdCancel_Click()
    blnIsOk = False
    
    Call Me.Hide
End Sub

Private Sub cmdDefault_Click()
On Error GoTo ErrHandle
    Call LoadDefaultShortcut
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub LoadDefaultShortcut()
    Dim i As Long
    Dim strDefaultKey As String
    Dim intShift As Integer
    Dim intKeyCode As Integer
    
    For i = 1 To ufgShoftcut.GridRows - 1
        strDefaultKey = ufgShoftcut.Text(i, "Ĭ�ϼ�")
        
        If strDefaultKey <> "" Then
            intShift = Val(Mid(strDefaultKey, 1, InStr(strDefaultKey, "+") - 1))
            intKeyCode = Val(Mid(strDefaultKey, InStr(strDefaultKey, "+") + 1, 8))
        
            ufgShoftcut.Text(i, "���Ƽ�") = intShift
            ufgShoftcut.Text(i, "�ַ���") = intKeyCode
            ufgShoftcut.Text(i, "��ݼ�") = GetKyeAlias(intKeyCode, intShift)
        Else
            ufgShoftcut.Text(i, "���Ƽ�") = 0
            ufgShoftcut.Text(i, "�ַ���") = 0
            ufgShoftcut.Text(i, "��ݼ�") = ""
        End If
    Next i
End Sub






Private Sub cmdSure_Click()
On Error GoTo ErrHandle
    Call UpdateShortCut
    
    blnIsOk = True
    
    Call Me.Hide
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Form_Load()
On Error GoTo ErrHandle
    Call RestoreWinState(Me, App.ProductName)
    
    blnIsOk = False
    
    #If DebugState = True Then
        mstrProject = "ZL9PACSWORK"
        mlngMudule = 1290
        
        Call InitDebugObject(1290, Me, "zlhis", "HIS")
    #End If
    
    Call InitShoftcutList
    
    Call LoadShortCutData
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub InitShoftcutList()
    '��������
    ufgShoftcut.GridRows = glngStandardRowCount
    '�����и�
    ufgShoftcut.RowHeightMin = glngStandardRowHeight
    
    '��ֹ�Ҽ������б����ô���
    ufgShoftcut.IsEjectConfig = False
    
    ufgShoftcut.IsKeepRows = False
    ufgShoftcut.ColNames = C_STR_SHORTCUT_LIST_COLS
    ufgShoftcut.DefaultColNames = C_STR_SHORTCUT_LIST_COLS
    ufgShoftcut.ColConvertFormat = ""
End Sub


Private Sub LoadShortCutData()
    Dim strSql As String
    
    strSql = "select a.id, a.��Ŀ, a.ģ���, a.�˵�ID, a.�˵�����, a.�˵�˵��, nvl(b.���Ƽ�, a.���Ƽ�) as ���Ƽ�, " & _
             "nvl(b.�ַ���, a.�ַ���) as �ַ���, a.Ĭ�ϼ�, decode(nvl(b.��ݹ���ID,''),'',a.�����,b.�����) as �����, a.������� " & _
             "from ��ݹ�����Ϣ a, (select ��ݹ���ID, ���Ƽ�, �ַ���, ����� from ��ݹ��ܹ��� where �û�id=[1] )b " & _
             "where a.id=b.��ݹ���ID(+) and a.��Ŀ=[2] and a.ģ���=[3] order by a.�������,a.id"
        
    Set ufgShoftcut.AdoData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, UserInfo.ID, UCase(mstrProject), mlngMudule)
    Call ufgShoftcut.RefreshData
    
End Sub


Private Function GetKyeAlias(KeyCode As Integer, Shift As Integer) As String

    Dim strShift As String
    Dim strTemp As String
    
    
    strShift = IIf((Shift And vbCtrlMask) <> 0, "CTRL", "")
    
    strTemp = IIf((Shift And vbShiftMask) <> 0, "SHIFT", "")
    If strTemp <> "" Then
        If strShift <> "" Then strShift = strShift & "+"
        strShift = strShift & strTemp
    End If
    
    strTemp = IIf((Shift And vbAltMask) <> 0, "ALT", "")
    If strTemp <> "" Then
        If strShift <> "" Then strShift = strShift & "+"
        strShift = strShift & strTemp
    End If
    
     
    
             
    strTemp = ""
    If KeyCode >= 48 And KeyCode <= 90 Then
        strTemp = Chr(KeyCode)
        
        If strShift = "" Then strShift = "MENU"
    End If
    
    If KeyCode >= vbKeyF1 And KeyCode <= vbKeyF12 Then
        strTemp = "F" & (KeyCode - 111)
    End If
    
    Select Case KeyCode
        Case vbKeySpace
            strTemp = "SPACE"
    End Select
    
    
    If strTemp <> "" Then
        If strShift <> "" Then strShift = strShift & "+"
        strShift = strShift & strTemp
    End If
    
    GetKyeAlias = strShift
                
End Function


Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Call SaveWinState(Me, App.ProductName)
    err.Clear
End Sub

Private Sub ufgShoftcut_OnKeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    Dim strAlias As String
    
    If Not ufgShoftcut.IsSelectionRow Then Exit Sub
    
    'ɾ����ݼ�
    If KeyCode = vbKeyDelete Then
        ufgShoftcut.Text(ufgShoftcut.SelectionRow, "��ݼ�") = ""
        ufgShoftcut.Text(ufgShoftcut.SelectionRow, "���Ƽ�") = 0
        ufgShoftcut.Text(ufgShoftcut.SelectionRow, "�ַ���") = 0
        
        Exit Sub
    End If
    
    strAlias = GetKyeAlias(KeyCode, Shift)
    If strAlias = "" Then Exit Sub
    
    
    ufgShoftcut.Text(ufgShoftcut.SelectionRow, "��ݼ�") = strAlias
    
    ufgShoftcut.Text(ufgShoftcut.SelectionRow, "���Ƽ�") = Shift
    ufgShoftcut.Text(ufgShoftcut.SelectionRow, "�ַ���") = KeyCode
    
End Sub

Private Sub cmdDelete_Click()
'ɾ����ť ִ��ɾ��ѡ���еĿ�ݼ�����
On Error Resume Next
    
    If Not ufgShoftcut.IsSelectionRow Then Exit Sub
    
    ufgShoftcut.Text(ufgShoftcut.SelectionRow, "��ݼ�") = ""
    ufgShoftcut.Text(ufgShoftcut.SelectionRow, "���Ƽ�") = 0
    ufgShoftcut.Text(ufgShoftcut.SelectionRow, "�ַ���") = 0
    
End Sub


'���¿�ݼ�����
Private Sub UpdateShortCut()
    Dim i As Long
    Dim strSql As String
    
    For i = 1 To ufgShoftcut.GridRows - 1
        If ufgShoftcut.Text(i, "��ݼ�") <> ufgShoftcut.Text(i, "��ǰ��") Then
            strSql = "ZL_��ݼ�_����(" & ufgShoftcut.KeyValue(i) & "," & _
                                            Val(UserInfo.ID) & "," & _
                                            Val(ufgShoftcut.Text(i, "���Ƽ�")) & "," & _
                                            Val(ufgShoftcut.Text(i, "�ַ���")) & ",'" & _
                                            ufgShoftcut.Text(i, "��ݼ�") & "')"
    
            Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
        End If
        '������״̬
        ufgShoftcut.RowState(i) = TDataRowState.Normal
    Next i
    
End Sub



Private Sub ufgShoftcut_OnKeyUp(KeyCode As Integer, Shift As Integer)
'�ж��Ƿ�����ظ��Ŀ�ݼ�
On Error GoTo ErrHandle
    Dim strKeyAlias As String
    Dim lngFindIndex As Long
    
    If Not ufgShoftcut.IsSelectionRow Then Exit Sub
    
    strKeyAlias = ufgShoftcut.Text(ufgShoftcut.SelectionRow, "��ݼ�")
    
    If InStr(strKeyAlias, "MENU+") >= 1 Then Exit Sub
    
    ufgShoftcut.Text(ufgShoftcut.SelectionRow, "��ݼ�") = ""
    
    lngFindIndex = ufgShoftcut.FindRowIndex(strKeyAlias, "��ݼ�", True)
    
    If lngFindIndex > 0 Then
        ufgShoftcut.Text(ufgShoftcut.SelectionRow, "���Ƽ�") = 0
        ufgShoftcut.Text(ufgShoftcut.SelectionRow, "�ַ���") = 0
        
        Call MsgBoxD(Me, "�����ظ��Ŀ�ݼ������������á�", vbOKOnly, Me.Caption)
    Else
        ufgShoftcut.Text(ufgShoftcut.SelectionRow, "��ݼ�") = strKeyAlias
    End If
    
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub
