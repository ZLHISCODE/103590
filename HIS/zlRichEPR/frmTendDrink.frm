VERSION 5.00
Begin VB.Form frmTendDrink 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�����������"
   ClientHeight    =   3825
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6210
   Icon            =   "frmTendDrink.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   6210
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4945
      TabIndex        =   2
      Top             =   3300
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   3720
      TabIndex        =   1
      Top             =   3300
      Width           =   1100
   End
   Begin zlRichEPR.VsfGrid vsf 
      Height          =   2835
      Left            =   135
      TabIndex        =   0
      Top             =   375
      Width           =   5910
      _ExtentX        =   10425
      _ExtentY        =   5001
   End
   Begin VB.Label lblUnit 
      AutoSize        =   -1  'True
      Caption         =   "ml"
      Height          =   180
      Left            =   4725
      TabIndex        =   4
      Top             =   90
      Width           =   180
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      Caption         =   "������ҩ������λ�൱�ڲ���ͬʱ�������Һ��������λ:"
      Height          =   180
      Left            =   105
      TabIndex        =   3
      Top             =   90
      Width           =   4590
   End
End
Attribute VB_Name = "frmTendDrink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

'######################################################################################################################
'�ֲ�������������

Private mblnOK As Boolean
Private mblnDataChanged As Boolean
Private mfrmMain As Form
Private mstrSQL As String

Private Enum mCol
    ͼ��
    ��λ
    ϵ��
    �����ϵ
End Enum


'######################################################################################################################
'�Զ��庯������������

Private Property Let DataChanged(ByVal vData As Boolean)
    mblnDataChanged = vData
End Property

Private Property Get DataChanged() As Boolean
    DataChanged = mblnDataChanged
End Property

Public Function ShowEdit(ByVal frmMain As Form) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:��/��ʾ�༭����,���������������(��ں���)
    '------------------------------------------------------------------------------------------------------------------
    
    mblnOK = False
    
    Set mfrmMain = frmMain
    
    If InitData = False Then GoTo errHand
    If ReadData = False Then GoTo errHand
            
    DataChanged = False
    
    Me.Show 1, frmMain
    
    ShowEdit = mblnOK
    
    Exit Function
    
errHand:
    On Error Resume Next
    DataChanged = False
    Unload Me
End Function

Private Function InitData() As Boolean
    Dim RS As New ADODB.Recordset
    On Error GoTo errHand
    
    With vsf
        .Cols = 0
        .NewColumn "", 255, 4
        .NewColumn "��λ", 1200, 1, , 1
        .NewColumn "ϵ��", 1200, 1, , 1
        .NewColumn "�����ϵ", 1800, 1
        .FixedCols = 1
        .Body.ExtendLastCol = True
        
        vsf.MaxLength(mCol.��λ) = 20
        vsf.MaxLength(mCol.ϵ��) = 10
    End With
    
    '�����������ĵ�λ
    lblUnit.Caption = ""
    mstrSQL = "Select ��Ŀ��λ From �����¼��Ŀ Where ��Ŀ���=7 And ������Ŀ=1"
    Set RS = zlDatabase.OpenSQLRecord(mstrSQL, Me.Caption)
    If RS.BOF = False Then
        lblUnit.Caption = zlCommFun.NVL(RS("��Ŀ��λ"))
    End If
    
    If lblUnit.Caption = "" Then
        ShowSimpleMsg "û��������������Ŀ���������ô�����ϵ��"
        Exit Function
    End If
    
    InitData = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function ReadData() As Boolean
    Dim strValue As String
    Dim lngLoop As Long
    Dim varAry As Variant

    strValue = zlDatabase.GetPara(62, glngSys)
    If strValue <> "" Then
        varAry = Split(strValue, ";")
        For lngLoop = 0 To UBound(varAry)
            If CStr(varAry(lngLoop)) <> "" Then
                
                If vsf.TextMatrix(vsf.Rows - 1, mCol.��λ) <> "" Then vsf.Rows = vsf.Rows + 1
                
                vsf.TextMatrix(vsf.Rows - 1, mCol.��λ) = Split(varAry(lngLoop), ",")(0)
                vsf.TextMatrix(vsf.Rows - 1, mCol.ϵ��) = Val(Split(varAry(lngLoop), ",")(1))
                Call vsf_AfterEdit(vsf.Rows - 1, mCol.��λ)
                
            End If
        Next
    End If
    
    ReadData = True
End Function

Private Function CheckData() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:У��༭���ݵ���Ч��
    '------------------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long
    
    For lngLoop = 1 To vsf.Rows - 1
        If vsf.TextMatrix(lngLoop, mCol.��λ) <> "" Then
            
            If Val(vsf.TextMatrix(lngLoop, mCol.ϵ��)) <= 0 Then
                ShowSimpleMsg "�� " & lngLoop & " �е�ϵ������ȷ���������0����"
                vsf.Row = lngLoop
                vsf.Col = mCol.ϵ��
                vsf.ShowCell vsf.Row, vsf.Col
                Exit Function
            End If
            
        End If
    Next
    
    CheckData = True
    
End Function

Private Function SaveData() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '���ܣ������޸Ļ�����������
    '���أ��ɹ����淵��True�����򷵻�False
    '------------------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long
    Dim strTmp As String
    
    For lngLoop = 1 To vsf.Rows - 1
        If vsf.TextMatrix(lngLoop, mCol.��λ) <> "" Then
            strTmp = strTmp & ";" & vsf.TextMatrix(lngLoop, mCol.��λ) & "," & Val(vsf.TextMatrix(lngLoop, mCol.ϵ��))
        End If
    Next
    
    If strTmp <> "" Then strTmp = Mid(strTmp, 2)
        
    Call zlDatabase.SetPara(62, strTmp, glngSys)
    
    SaveData = True
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
        
    If DataChanged Then
        If CheckData = False Then Exit Sub
        If SaveData = False Then Exit Sub
                
        mblnOK = True
        DataChanged = False
    End If
    
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If DataChanged Then
        Cancel = (MsgBox("����/�޸ĵ����ݱ��뱣������Ч���Ƿ񲻱�����˳���", vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) = vbNo)
    End If
End Sub

Private Sub vsf_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    
    '���»����ϵ�б��ʽ
    
    If vsf.TextMatrix(Row, mCol.��λ) <> "" And Val(vsf.TextMatrix(Row, mCol.ϵ��)) > 0 And lblUnit.Caption <> "" Then
        
        vsf.TextMatrix(Row, mCol.�����ϵ) = "(1)" & vsf.TextMatrix(Row, mCol.��λ) & " = (" & Val(vsf.TextMatrix(Row, mCol.ϵ��)) & ")" & lblUnit.Caption
    Else
        vsf.TextMatrix(Row, mCol.�����ϵ) = ""
    End If
    
    DataChanged = True
    
End Sub

Private Sub vsf_KeyPress(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer, Cancel As Boolean)
    If KeyAscii <> vbKeyReturn Then
        Select Case Col
            Case mCol.ϵ��
                If FilterKeyAscii(KeyAscii, 99, "0123456789.") = 0 Then KeyAscii = 0
        End Select
    End If
End Sub

Private Sub vsf_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii <> vbKeyReturn Then
        Select Case Col
            Case mCol.ϵ��
                If FilterKeyAscii(KeyAscii, 99, "0123456789.") = 0 Then KeyAscii = 0
        End Select
    End If
End Sub
