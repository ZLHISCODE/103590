VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "Msmask32.ocx"
Begin VB.UserControl usrMedicalSum 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   5415
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7935
   LockControls    =   -1  'True
   ScaleHeight     =   5415
   ScaleWidth      =   7935
   Begin VB.Frame fra 
      BackColor       =   &H80000005&
      Height          =   2280
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7110
      Begin VB.PictureBox pic 
         BorderStyle     =   0  'None
         Height          =   330
         Index           =   0
         Left            =   60
         ScaleHeight     =   330
         ScaleWidth      =   5550
         TabIndex        =   1
         Top             =   150
         Width           =   5550
         Begin MSComctlLib.Toolbar cbr 
            Height          =   345
            Index           =   0
            Left            =   465
            TabIndex        =   2
            Top             =   0
            Width           =   4020
            _ExtentX        =   7091
            _ExtentY        =   609
            ButtonWidth     =   1349
            ButtonHeight    =   609
            AllowCustomize  =   0   'False
            Wrappable       =   0   'False
            Style           =   1
            TextAlignment   =   1
            ImageList       =   "ils16"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   3
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "�ռ�"
                  Key             =   "�ռ�"
                  Object.ToolTipText     =   "�ռ����������Ŀ��С��"
                  Object.Tag             =   "�ռ�"
                  ImageIndex      =   2
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "����"
                  Key             =   "����"
                  Object.ToolTipText     =   "������������������"
                  Object.Tag             =   "����"
                  ImageKey        =   "new"
               EndProperty
               BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "���"
                  Key             =   "���"
                  Object.ToolTipText     =   "�����������н���"
                  Object.Tag             =   "���"
                  ImageKey        =   "cls"
               EndProperty
            EndProperty
         End
         Begin VB.Label picTitle 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����"
            Height          =   180
            Index           =   0
            Left            =   60
            TabIndex        =   3
            Top             =   75
            Width           =   360
         End
      End
      Begin zl9CISCore.VsfGrid vsf 
         Height          =   1695
         Left            =   165
         TabIndex        =   4
         Top             =   525
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   2990
      End
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   7260
      Top             =   3000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "usrMedicalSum.ctx":0000
            Key             =   "cls"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "usrMedicalSum.ctx":0296
            Key             =   "search"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "usrMedicalSum.ctx":6AF8
            Key             =   "new"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "usrMedicalSum.ctx":D35A
            Key             =   "newadvice"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "usrMedicalSum.ctx":13BBC
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "usrMedicalSum.ctx":1A41E
            Key             =   "SelAll"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "usrMedicalSum.ctx":20C80
            Key             =   "SelDel"
         EndProperty
      EndProperty
   End
   Begin VB.Frame fra 
      BackColor       =   &H80000005&
      Height          =   1290
      Index           =   1
      Left            =   30
      TabIndex        =   9
      Top             =   2325
      Width           =   6960
      Begin VB.TextBox rtb 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   765
         Left            =   45
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   15
         Top             =   480
         Width           =   1170
      End
      Begin VB.PictureBox pic 
         BorderStyle     =   0  'None
         Height          =   330
         Index           =   1
         Left            =   15
         ScaleHeight     =   330
         ScaleWidth      =   4695
         TabIndex        =   10
         Top             =   105
         Width           =   4695
         Begin MSComctlLib.Toolbar cbr 
            Height          =   345
            Index           =   1
            Left            =   570
            TabIndex        =   11
            Top             =   15
            Width           =   2445
            _ExtentX        =   4313
            _ExtentY        =   609
            ButtonWidth     =   1349
            ButtonHeight    =   609
            AllowCustomize  =   0   'False
            Wrappable       =   0   'False
            Style           =   1
            TextAlignment   =   1
            ImageList       =   "ils16"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   3
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "�ռ�"
                  Key             =   "�ռ�"
                  Object.ToolTipText     =   "�ռ��������Ŀ�Ľ�������"
                  Object.Tag             =   "�ռ�"
                  ImageIndex      =   2
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "����"
                  Key             =   "����"
                  Object.ToolTipText     =   "������Ľ�������ȱʡ����"
                  Object.Tag             =   "����"
                  ImageKey        =   "newadvice"
               EndProperty
               BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "���"
                  Key             =   "���"
                  Object.ToolTipText     =   "�������Ľ�������"
                  Object.Tag             =   "���"
                  ImageKey        =   "cls"
               EndProperty
            EndProperty
         End
         Begin VB.Label picTitle 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����"
            Height          =   180
            Index           =   1
            Left            =   75
            TabIndex        =   12
            Top             =   75
            Width           =   360
         End
      End
   End
   Begin VB.Frame fraOther 
      BackColor       =   &H80000005&
      Height          =   615
      Left            =   165
      TabIndex        =   5
      Top             =   4320
      Width           =   7110
      Begin VB.CommandButton cmd 
         Caption         =   "������Ŀ"
         Height          =   350
         Left            =   2265
         TabIndex        =   16
         Top             =   165
         Visible         =   0   'False
         Width           =   1100
      End
      Begin MSMask.MaskEdBox msk 
         Height          =   240
         Left            =   1245
         TabIndex        =   14
         Top             =   225
         Visible         =   0   'False
         Width           =   990
         _ExtentX        =   1746
         _ExtentY        =   423
         _Version        =   393216
         BorderStyle     =   0
         Appearance      =   0
         ForeColor       =   255
         AutoTab         =   -1  'True
         MaxLength       =   10
         Format          =   "YYYY-MM-DD"
         Mask            =   "####-##-##"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txt 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         ForeColor       =   &H000000C0&
         Height          =   225
         Left            =   4605
         TabIndex        =   13
         Top             =   210
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.CheckBox chk 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "����ʱ��:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   75
         TabIndex        =   7
         Top             =   240
         Width           =   1125
      End
      Begin VB.CheckBox chk 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "�������:"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   3465
         TabIndex        =   6
         Top             =   225
         Width           =   1140
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   4590
         X2              =   5160
         Y1              =   435
         Y2              =   435
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   1260
         X2              =   2175
         Y1              =   465
         Y2              =   465
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��"
         Height          =   180
         Index           =   1
         Left            =   5220
         TabIndex        =   8
         Top             =   255
         Width           =   180
      End
   End
End
Attribute VB_Name = "usrMedicalSum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mstr�Һŵ� As String                    '��紫��
Private mlng����id As Long                      '��紫��
Private mlngҽ��id As Long                      '��紫��
Private mlng����id As Long                      '��紫��

Private mblnMode As Boolean 'Ϊ���Ǳ�ʾ���û����еı༭����ʱ�Ÿ�ֵ
Private mDispMode As Boolean
Private mReturnErrnumber As Long
Private mReturnErrDescription As String
Private mobjParentObject As Object
Private mrsCallBack As New ADODB.Recordset

Private Enum mCol
    �������� = 1
    �쳣���
    ����
    ��Ͻ���
End Enum


Public Function ShowFilterDiagBox(ByVal frmParent As Object, _
                                    ByVal objCmd As Object, _
                                    ByVal strLvw As String, _
                                    ByVal strSavePath As String, _
                                    ByVal strDescrible As String, _
                                    ByVal rsData As ADODB.Recordset, _
                                    ByRef rsResult As ADODB.Recordset, _
                                    Optional ByVal lngCX As Long = 6000, _
                                    Optional ByVal lngCY As Long = 3000, _
                                    Optional ByVal blnFilter As Boolean = True, _
                                    Optional ByVal blnPrompt As Boolean = True, _
                                    Optional ByVal blnMuli As Boolean = False) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����;��ʾ�ı�����ѡ���б�(ֻ���ڱ��ؼ�)
    '------------------------------------------------------------------------------------------------------------------

    Dim objPoint As POINTAPI
    Dim lngX As Long
    Dim lngY As Long
    
    On Error GoTo errHand


    If rsData.BOF Then
        If blnPrompt Then MsgBox "û���ҵ���ƥ��Ľ����", , gstrSysName
        Exit Function                            'û�н����ֱ�ӷ���
    End If
    If rsData.RecordCount = 1 And blnFilter Then GoTo Over                    '��Ϊ��������ң����ֻ��һ������ֱ�ӷ���
        
    Call ClientToScreen(objCmd.hWnd, objPoint)
    lngX = objPoint.x * Screen.TwipsPerPixelX
    lngY = objPoint.y * Screen.TwipsPerPixelY + objCmd.Height

    If frmSelectDialog.ShowSelect(Nothing, 2, rsData, strLvw, strDescrible, lngX, lngY, lngCX, lngCY, , , strSavePath, , False, blnMuli) Then GoTo Over
    
    Exit Function
    
Over:
    
    Set rsResult = rsData
    
    ShowFilterDiagBox = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function InitControl() As Boolean
    
    With vsf
    
        .Cols = 0
        .NewColumn "", 255
        .NewColumn "��������", 2400, 1, "...", 1
        .NewColumn "�쳣���", 3000, 1, , 1
        .NewColumn "����", 600, 1, , 1
        .NewColumn "��Ͻ���", 15, 1, , 1
        .FixedCols = 1
        
        .ColDataType(mCol.����) = flexDTBoolean
        
        .TextMatrix(1, mCol.��������) = "δ���쳣"
        
        .Body.Appearance = flexXPThemes
        .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = True
                
    End With
    
    Set mrsCallBack = Nothing
End Function

Public Property Set ParentObject(vData As Object)
    Set mobjParentObject = vData
End Property

Public Property Get ParentObject() As Object
    Set ParentObject = mobjParentObject
End Property

Private Property Let Modified(vData As Boolean)
    
    On Error Resume Next
    
    If mobjParentObject Is Nothing Then Exit Property
    
    mobjParentObject.Modified = vData
    
End Property

Private Property Get Modified() As Boolean
    
    On Error Resume Next
    
    If mobjParentObject Is Nothing Then Exit Property
    
    Modified = mobjParentObject.Modified
    
End Property



'��������������
Public Sub SetgcnOracle()
    '------------------------------------------------------------------------------------------------------------------
    '�ӿڹ���
    '------------------------------------------------------------------------------------------------------------------
    Call InitCommon(gcnOracle)
    
End Sub

Public Property Get DispMode() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '�ӿڹ���:�Ƿ�Ϊ��ʾģʽ
    '------------------------------------------------------------------------------------------------------------------
    DispMode = mDispMode
End Property

Public Property Let DispMode(ByVal New_DispMode As Boolean)
    '------------------------------------------------------------------------------------------------------------------
    '
    '------------------------------------------------------------------------------------------------------------------
    mDispMode = New_DispMode
    
    ShowUsrControl mlngҽ��id, Not mDispMode
    PropertyChanged "DispMode"
    
    If mDispMode Then
        vsf.Body.Editable = flexEDNone
        
        rtb.Locked = True
        
        cbr(0).Buttons("�ռ�").Enabled = False
        cbr(0).Buttons("����").Enabled = False
        cbr(0).Buttons("���").Enabled = False
                        
        cbr(1).Buttons("�ռ�").Enabled = False
        cbr(1).Buttons("����").Enabled = False
        cbr(1).Buttons("���").Enabled = False
        
        cbr(0).Visible = False
        cbr(1).Visible = False
        
        fraOther.Enabled = False
        
    Else
        cbr(0).Visible = True
        cbr(1).Visible = True
        
        fraOther.Enabled = True
    End If
    
End Property
Public Property Let �Һŵ�(ByVal New_�Һŵ� As String)
    '------------------------------------------------------------------------------------------------------------------
    '���ùҺŵ�
    '------------------------------------------------------------------------------------------------------------------
    
    mstr�Һŵ� = New_�Һŵ�
    
End Property

Public Property Let ����id(ByVal New_����id As String)
    '------------------------------------------------------------------------------------------------------------------
    '���ùҺŵ�
    '------------------------------------------------------------------------------------------------------------------
    
    mlng����id = New_����id
    
End Property

Public Property Get ID���˲���() As Long
    '------------------------------------------------------------------------------------------------------------------
    '���ز��˲���ID
    '------------------------------------------------------------------------------------------------------------------
    
    ID���˲��� = mlng����id
    
End Property

Public Property Let ID���˲���(ByVal New_ID���˲��� As Long)
    '------------------------------------------------------------------------------------------------------------------
    '���ò��˲���ID,�����ò����ǲ��Ǵ���
    '------------------------------------------------------------------------------------------------------------------
    
    mlng����id = New_ID���˲���
    ShowUsrControl mlngҽ��id, Not mDispMode
    
End Property

Public Sub SetDiagItem(ByVal New_ҽ��ID As Long, ByVal New_���ͺ�)
    '------------------------------------------------------------------------------------------------------------------
    '
    '------------------------------------------------------------------------------------------------------------------
    
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    
    mlngҽ��id = New_ҽ��ID
    
End Sub

Public Property Get Getҽ��id() As Long
    '------------------------------------------------------------------------------------------------------------------
    '
    '------------------------------------------------------------------------------------------------------------------
    
    Getҽ��id = mlngҽ��id
        
End Property

Public Property Get Text() As String
    '------------------------------------------------------------------------------------------------------------------
    'Ϊÿһ���ؼ������ı�ת������
    '------------------------------------------------------------------------------------------------------------------
    
    Dim lngLoop As Long
    Dim strTmp As String
    Dim intCount As Integer
    
    On Error GoTo errHand
    
    'ת�����ۼ�¼
    intCount = 0
    strTmp = strTmp & "һ�����ۣ�" & vbCrLf
    For lngLoop = 1 To vsf.Rows - 1
        
        If vsf.TextMatrix(lngLoop, mCol.��������) <> "" Then
            intCount = intCount + 1
            strTmp = strTmp & intCount & "��" & vsf.TextMatrix(lngLoop, mCol.��������) & vbCrLf
        End If
        
    Next
    strTmp = strTmp & vbCrLf
    
    'ת����������
    strTmp = strTmp & "�������飺" & vbCrLf
    strTmp = strTmp & rtb.Text
    
    Text = strTmp
    
    Exit Property
    
errHand:
    
End Property

Public Sub ClearData()
    '------------------------------------------------------------------------------------------------------------------
    '����:�ӿ�
    '����:
    '����:
    '------------------------------------------------------------------------------------------------------------------
    vsf.Rows = 2
    vsf.RowData(1) = 0
    vsf.Cell(flexcpText, 1, 1, vsf.Rows - 1, vsf.Cols - 1) = ""
    
    rtb.Text = ""
End Sub

Public Function SaveData(lng����ID As Long, lng��ҳID As Long, lng����ID As Long, strReturnSQL As String, strError As String) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:�ӿ�
    '����:
    '����:
    '------------------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long
    Dim strTmp As String
    Dim strSQL() As String
    Dim strDate As String
    Dim LngCount As Long
    
    On Error GoTo errHand
    
    ReDim Preserve strSQL(0 To vsf.Rows + 1)
    
    For lngLoop = 1 To vsf.Rows - 1
        If StrIsValid(vsf.TextMatrix(lngLoop, 1), 100) = False Then
            vsf.Row = lngLoop
            vsf.Col = 1
            vsf.ShowCell vsf.Row, vsf.Col
            Exit Function
        End If
    Next
    
    If StrIsValid(rtb.Text, 4000) = False Then
        rtb.SetFocus
    End If
    
    If chk(0).Value = 1 Then
        strDate = Format(msk.Text, "yyyy-MM-dd")
        If IsDate(strDate) = False Then
            strDate = ""
        Else
            strDate = strDate & " 00:00:00"
        End If
    End If
    
    strSQL(0) = "ZL_�����Ա����_DELETE(" & lng����ID & ")"
    LngCount = 0
    
    For lngLoop = 1 To vsf.Rows - 1
        If Trim(vsf.TextMatrix(lngLoop, mCol.��������)) <> "" Then
            
            LngCount = LngCount + 1
            
            strSQL(lngLoop) = "ZL_�����Ա����_INSERT(" & lng����ID & "," & _
                                                        lng��ҳID & "," & _
                                                        lng����ID & "," & _
                                                        "0," & _
                                                        LngCount & ",'" & _
                                                        vsf.TextMatrix(lngLoop, mCol.��������) & "','" & _
                                                        vsf.TextMatrix(lngLoop, mCol.�쳣���) & "'," & _
                                                        "NULL," & _
                                                        Val(vsf.RowData(lngLoop)) & ",NULL,NULL," & _
                                                        Abs(Val(vsf.TextMatrix(lngLoop, mCol.����))) & "," & _
                                                        "'" & vsf.TextMatrix(lngLoop, mCol.��Ͻ���) & "')"
        End If
    Next
    strSQL(lngLoop + 1) = "ZL_�����Ա����_INSERT(" & lng����ID & "," & _
                                                        lng��ҳID & "," & _
                                                        lng����ID & "," & _
                                                        "1," & _
                                                        "1," & _
                                                        "NULL,NULL,'" & _
                                                        rtb.Text & "'," & _
                                                        "NULL," & _
                                                        IIf(strDate = "", "Null", "To_date('" & strDate & "','yyyy-mm-dd hh24:mi:ss')") & "," & _
                                                        IIf(chk(1).Value = 0, "Null", Val(txt.Text)) & "," & _
                                                        "0," & _
                                                        "NULL)"
        
    strTmp = ""
    For lngLoop = 0 To UBound(strSQL)
        If strSQL(lngLoop) <> "" Then
        
            strSQL(lngLoop) = Replace(strSQL(lngLoop), Chr(9), Chr(32))
            
            If strTmp = "" Then
                strTmp = strSQL(lngLoop)
            Else
                strTmp = strTmp & Chr(9) & strSQL(lngLoop)
            End If
        End If
    Next
    
    Dim strCallBack As String
    
    If chk(0).Value = 1 Then
        strCallBack = ""
        If Not (mrsCallBack Is Nothing) Then
            If mrsCallBack.State = adStateOpen Then
                If mrsCallBack.RecordCount > 0 Then
                    mrsCallBack.MoveFirst
                    Do While Not mrsCallBack.EOF
                        strCallBack = strCallBack & "," & Val(mrsCallBack("�嵥id").Value)
                        mrsCallBack.MoveNext
                    Loop
                    
                    If strCallBack <> "" Then strCallBack = Mid(strCallBack, 2)
                    strTmp = strTmp & Chr(9) & "ZL_���ǼǼ�¼_����('" & mstr�Һŵ� & "'," & mlng����id & ",To_Date('" & strDate & "','yyyy-mm-dd hh24:mi:ss'),'" & strCallBack & "')"
                End If
            End If
            
        End If
    Else
        strTmp = strTmp & Chr(9) & "ZL_���ǼǼ�¼_����('" & mstr�Һŵ� & "'," & mlng����id & ",Null,Null)"
    End If
    
    '����SQL���
    strReturnSQL = strTmp
    
    SaveData = True
    
    Exit Function
    
errHand:

    strError = "���ר��ֽ����ʧ�ܣ�"
    
End Function

Private Function ShowOpenList(Optional strText As String, Optional ByVal bytMode As Byte = 1) As Byte
    '------------------------------------------------------------------------------------------------------------------
    '����:���б�ṹ�����Ƽ���걾����
    '����:������2;�ɹ�����1;ȡ������0
    '------------------------------------------------------------------------------------------------------------------
    Dim strLvw As String
    Dim strTitle As String
    Dim strDescrible As String
    
    Dim sglX As Single
    Dim sglY As Single
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errHand
    
    ShowOpenList = 2
    
    strText = "'%" & UCase(strText) & "%'"
    
    If bytMode = 1 Then
        
        strLvw = "����,900,0,1;����,1800,0,0;��Ͻ���,2700,0,0"
        strTitle = "�����۹���"
        strDescrible = "����±���ѡ��һ��������"
        
        strSQL = _
                    "SELECT A.��� AS ID, " & _
                            "A.����, " & _
                            "A.����, " & _
                            "A.�Ƿ񼲲�,A.��Ͻ��� " & _
                    "FROM �����Ͻ��� A " & _
                    "WHERE NVL(ĩ��,0)=1 "
        strSQL = strSQL & " AND (A.���� Like " & strText & " OR A.���� Like " & strText & " OR A.���� Like " & UCase(strText) & ")"
    End If
    
    Call OpenRecord(rs, strSQL, "������")
    If rs.BOF Then
        ShowOpenList = 0
        Exit Function
    End If
    
    If rs.RecordCount = 1 Then GoTo Over
    Call CalcPosition(sglX, sglY, vsf)
    If frmSelectList.ShowSelect(Me, rs, strLvw, sglX + 60, sglY, 9000, 5100, strTitle, strDescrible) Then
        GoTo Over
    End If
    
    Exit Function
    
Over:
    If bytMode = 1 Then
        vsf.RowData(vsf.Row) = zlCommFun.NVL(rs("ID").Value)
        vsf.EditText = zlCommFun.NVL(rs("����").Value)
        vsf.Cell(flexcpData, vsf.Row, vsf.Col) = zlCommFun.NVL(rs("����").Value)
        vsf.TextMatrix(vsf.Row, vsf.Col) = zlCommFun.NVL(rs("����").Value)
        
        vsf.TextMatrix(vsf.Row, mCol.����) = zlCommFun.NVL(rs("�Ƿ񼲲�").Value)
        vsf.TextMatrix(vsf.Row, mCol.��Ͻ���) = zlCommFun.NVL(rs("��Ͻ���").Value)
        
        Call ReadPreVisiteDate(2)
        
    End If
    
    Modified = True
    
    ShowOpenList = 1
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then Resume
End Function


Private Function ShowOpenTree(Optional ByVal bytMode As Byte = 1) As Byte
    '------------------------------------------------------------------------------------------------------------------
    '����:���б�ṹ�����Ƽ���걾����
    '����:������2;�ɹ�����1;ȡ������0
    '------------------------------------------------------------------------------------------------------------------
    Dim strLvw As String
    Dim strTitle As String
    Dim strDescrible As String
    
    Dim sglX As Single
    Dim sglY As Single
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errHand
    
    ShowOpenTree = 2
    
    If bytMode = 1 Then
        strLvw = "����,900,0,1;����,1800,0,0;��Ͻ���,2700,0,0"
        strTitle = "������ѡ��"
        strDescrible = "����±���ѡ��һ��������"
        
        strSQL = "SELECT -1 AS ID," & _
                            "0 AS �ϼ�ID," & _
                            "0 AS ĩ��," & _
                            "'' AS ����," & _
                            "'���з���' AS ����, " & _
                            "Null+0 AS �Ƿ񼲲�,'' As ��Ͻ��� " & _
                    "FROM dual "
                    
        strSQL = strSQL & _
                " UNION ALL " & _
                "SELECT ��� AS ID," & _
                            "DECODE(�ϼ����,NULL,-1,�ϼ����) AS �ϼ�ID," & _
                            "0 AS ĩ��," & _
                            "����," & _
                            "����, " & _
                            "Null+0 AS �Ƿ񼲲�,'' As ��Ͻ��� " & _
                    "FROM �����Ͻ��� " & _
                    "WHERE NVL(ĩ��,0)=0 " & _
                    "START WITH �ϼ���� is NULL CONNECT BY PRIOR ��� = �ϼ���� "
        
        strSQL = strSQL & _
                    "UNION ALL " & _
                    "SELECT A.��� AS ID, " & _
                            "DECODE(�ϼ����,NULL,-1,�ϼ����) AS �ϼ�ID, " & _
                            "1 AS ĩ��, " & _
                            "A.����, " & _
                            "A.����, " & _
                            "A.�Ƿ񼲲�,A.��Ͻ��� " & _
                    "FROM �����Ͻ��� A " & _
                    "WHERE NVL(A.ĩ��,0)=1"
    End If
    
    Call OpenRecord(rs, strSQL, "������")
    
    If rs.BOF Then
        ShowOpenTree = 0
        Exit Function
    End If
        
    Call CalcPosition(sglX, sglY, vsf)
    
    If frmSelectTree.ShowSelect(Screen, rs, sglX, sglY, 9000, 5100, vsf.CellHeight, strTitle, strLvw, strDescrible) Then
        GoTo Over
    End If
    
    Exit Function
    
Over:
    If bytMode = 1 Then
    
        vsf.RowData(vsf.Row) = zlCommFun.NVL(rs("ID").Value)
        vsf.EditText = zlCommFun.NVL(rs("����").Value)
        vsf.Cell(flexcpData, vsf.Row, vsf.Col, vsf.Row, vsf.Col) = zlCommFun.NVL(rs("����").Value)
        vsf.TextMatrix(vsf.Row, vsf.Col) = zlCommFun.NVL(rs("����").Value)
        
        vsf.TextMatrix(vsf.Row, mCol.����) = zlCommFun.NVL(rs("�Ƿ񼲲�").Value, 0)
        vsf.TextMatrix(vsf.Row, mCol.��Ͻ���) = zlCommFun.NVL(rs("��Ͻ���").Value)
        
        Call ReadPreVisiteDate(2)
    End If
    
    Modified = True
    
    ShowOpenTree = 1
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then Resume
End Function

Private Sub CalcPosition(ByRef x As Single, ByRef y As Single, ByVal objBill As Object)
    '------------------------------------------------------------------------------------------------------------------
    '���ܣ� ����X,Y��ʵ�����꣬��������Ļ���������
    '������ X---���غ��������
    '       Y---�������������
    '------------------------------------------------------------------------------------------------------------------
    Dim objPoint As POINTAPI
    
    Call ClientToScreen(objBill.hWnd, objPoint)
    
    x = objPoint.x * 15 + objBill.CellLeft - 45
    y = objPoint.y * 15 + objBill.CellTop + objBill.CellHeight - 30
End Sub

Private Function GetAdvice() As String
    '------------------------------------------------------------------------------------------------------------------
    '����:
    '------------------------------------------------------------------------------------------------------------------
    
    Dim rs As New ADODB.Recordset
    Dim lngLoop As Long
    Dim strSQL As String
        
    On Error GoTo errHand
    
    GetAdvice = ""
    
    For lngLoop = 1 To vsf.Rows - 1
        If Val(vsf.RowData(lngLoop)) > 0 Then
            
            strSQL = "SELECT �ο����� FROM �����Ͻ��� WHERE ��� = " & Val(vsf.RowData(lngLoop))
            Call OpenRecord(rs, strSQL, "������")
            If rs.BOF = False Then
                
                If zlCommFun.NVL(rs("�ο�����").Value) <> "" Then
                    If vsf.TextMatrix(lngLoop, mCol.�쳣���) <> "" Then
                        GetAdvice = GetAdvice & lngLoop & "." & vsf.TextMatrix(lngLoop, mCol.��������) & " {" & vsf.TextMatrix(lngLoop, mCol.�쳣���) & "}��" & vbCrLf
                    Else
                        GetAdvice = GetAdvice & lngLoop & "." & vsf.TextMatrix(lngLoop, mCol.��������) & "��" & vbCrLf
                    End If
                    GetAdvice = GetAdvice & zlCommFun.NVL(rs("�ο�����").Value) & vbCrLf & vbCrLf
                End If
                
            End If
            
        End If
    Next
    
    Exit Function
    
errHand:
        
End Function

Private Sub SetErr(lngErrNum As Long, strErr As String)
    '------------------------------------------------------------------------------------------------------------------
    '���ô��������������
    '���lngErrNum=-1 ��ʾ �ؼ��Լ�����Ĵ���
    '------------------------------------------------------------------------------------------------------------------
    
    mReturnErrnumber = lngErrNum
    mReturnErrDescription = strErr
End Sub

Private Function InDesign() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '���ܣ��жϵ�ǰ���г����Ƿ���VB�Ĺ��̻�����
    '------------------------------------------------------------------------------------------------------------------
    
    On Error Resume Next
    
    Debug.Print 1 / 0
    If Err.Number <> 0 Then Err.Clear: InDesign = True
    
End Function

Private Sub ShowUsrControl(lngKey As Long, Optional ByVal blnEditMode As Boolean = False)
    '------------------------------------------------------------------------------------------------------------------
    '���ܣ��ⲿ������ʾ
    '------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rs As New ADODB.Recordset
    Dim intRow As Integer
    Dim blnSave As Boolean
    
    On Error GoTo errHand
    
    blnSave = Modified
    
    mDispMode = Not blnEditMode
    
    'Begin  <��ʼ������>
    
    If gcnOracle Is Nothing Then SetErr -1, "���Ӷ���û�г�ʼ��": Exit Sub
    If gcnOracle.State <> adStateOpen Then SetErr -1, "���Ӷ���û������": Exit Sub
    
    'End    <��ʼ������>


    'Begin  <��ȡ����>
    
    Call InitControl
    vsf.ExtendLastCol = True
    
    intRow = 0
    
    strSQL = "SELECT DISTINCT A.��¼����, A.��¼���,A.��������,A.�쳣���,A.�ο�����,A.����id,A.�Ƿ񼲲�,A.��Ͻ��� FROM �����Ա���� A WHERE A.����id=" & mlng����id & " ORDER BY A.��¼����,A.��¼���"
    Call OpenRecord(rs, strSQL, "���ר��ֽ")
    If rs.BOF = False Then
        Do While Not rs.EOF
            
            If zlCommFun.NVL(rs("��¼����").Value) = 0 Then
                
                intRow = intRow + 1
                vsf.Rows = intRow + 1
                
                vsf.RowData(intRow) = zlCommFun.NVL(rs("����id").Value)
'                vsf.TextMatrix(intRow, 0) = zlCommFun.Nvl(rs("��¼���").Value) & "��"
                vsf.TextMatrix(intRow, mCol.��������) = zlCommFun.NVL(rs("��������").Value)
                vsf.TextMatrix(intRow, mCol.�쳣���) = zlCommFun.NVL(rs("�쳣���").Value)
                vsf.TextMatrix(intRow, mCol.����) = zlCommFun.NVL(rs("�Ƿ񼲲�").Value, 0)
                vsf.TextMatrix(intRow, mCol.��Ͻ���) = zlCommFun.NVL(rs("��Ͻ���").Value)
                                
            Else
                rtb.Text = zlCommFun.NVL(rs("�ο�����").Value)
            End If
            
            rs.MoveNext
        Loop
        
        strSQL = "Select a.����ʱ��,a.������� From �����Ա���� a,���˲������� b Where a.��첡��id=b.������¼id and b.id=[1]"
        Set rs = OpenSQLRecord(strSQL, "���ר��ֽ", mlng����id)
        If rs.BOF = False Then
            
            If zlCommFun.NVL(rs("����ʱ��")) <> "" Then
                msk.Text = Format(zlCommFun.NVL(rs("����ʱ��")), "yyyy-MM-dd")
                chk(0).Value = 1
            End If
            
            If zlCommFun.NVL(rs("�������"), 0) > 0 Then
                txt.Text = zlCommFun.NVL(rs("�������"), 0)
                chk(1).Value = 1
            End If
        End If
    Else
        Call ReadPreVisiteDate
    End If
    
    'End    <��ȡ����>
        
    Modified = blnSave
    
    Exit Sub
    
errHand:

    If Ambient.UserMode = False Or InDesign = False Then
        SetErr Err.Number, Err.Description
        Exit Sub
    End If
    
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Function EditRefresh(ByVal objVsf As Object) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '���ܣ�
    '------------------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long
    Dim LngCount As Long
    
    On Error GoTo errHand
        
    If MsgBox("�Ƿ�Ҫ�滻ԭ�����ܼ���ۣ�", vbYesNo + vbQuestion + vbDefaultButton1, gstrSysName) = vbYes Then
        vsf.Rows = 2
        vsf.RowData(1) = 0
        vsf.TextMatrix(1, 1) = ""
    End If
    
    For lngLoop = 1 To objVsf.Rows - 1
        If Val(objVsf.RowData(lngLoop)) > 0 Then
            If Abs(Val(objVsf.TextMatrix(lngLoop, 0))) = 1 Then
                
                '���Val(objVsf.RowData(lngLoop))�Ƿ��Ѿ�����
                For LngCount = 0 To vsf.Rows - 1
                    If Trim(vsf.TextMatrix(LngCount, 1)) = Trim(objVsf.TextMatrix(lngLoop, 1)) Then
                        GoTo NextLoop
                    End If
                Next
                
                If Val(vsf.RowData(vsf.Rows - 1)) > 0 Then vsf.Rows = vsf.Rows + 1
                                
                vsf.RowData(vsf.Rows - 1) = Val(objVsf.RowData(lngLoop))
'                vsf.TextMatrix(vsf.Rows - 1, 0) = vsf.Rows - 1 & "��"
                vsf.TextMatrix(vsf.Rows - 1, 1) = objVsf.TextMatrix(lngLoop, 1)
                vsf.TextMatrix(vsf.Rows - 1, 2) = objVsf.TextMatrix(lngLoop, 2)
                vsf.TextMatrix(vsf.Rows - 1, 3) = Abs(Val(objVsf.Cell(flexcpData, lngLoop, 1, lngLoop, 1)))
                
            End If
        End If
        
NextLoop:
        
    Next
    
    Call ReadPreVisiteDate(2)
    
    EditRefresh = True
    
    Exit Function
    
errHand:
    
End Function

Private Function ReadPreVisiteDate(Optional ByVal bytMode As Byte = 1) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '���ܣ�
    '������
    '���أ�
    '------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rs As New ADODB.Recordset
    Dim lngLoop As Long
    Dim strAgain As String
    Dim lngVisit As Long
    
    On Error GoTo errHand
    
    Select Case bytMode
    Case 1      '��ԤԼ�Ǽ�ʱ����ñ�־
    
        strSQL = "Select ������� From ���ǼǼ�¼ Where ����=[1] and ������� Is Not Null"
        Set rs = OpenSQLRecord(strSQL, "������", mstr�Һŵ�)
        If rs.BOF = False Then
            
            txt.Text = zlCommFun.NVL(rs("�������"))
                        
        End If
    Case 2          '�ӵ�ǰ�Ľ�������ȡ�������޼�����ʱ��
        For lngLoop = 1 To vsf.Rows - 1
            If Val(vsf.RowData(lngLoop)) > 0 Then
                strSQL = "Select �������,������*30+sysdate As ����ʱ�� From �����Ͻ��� Where ���=[1]"
                Set rs = OpenSQLRecord(strSQL, "������", Val(vsf.RowData(lngLoop)))
                If rs.BOF = False Then
                    
                    If lngVisit < zlCommFun.NVL(rs("�������"), 0) Then lngVisit = zlCommFun.NVL(rs("�������"), 0)
                    If strAgain < Format(zlCommFun.NVL(rs("����ʱ��")), "yyyy-MM-dd") Then strAgain = Format(zlCommFun.NVL(rs("����ʱ��")), "yyyy-MM-dd")
                    
                End If
            End If
        Next
        
        If strAgain <> "" Then
            msk.Text = strAgain
        Else
            msk.Text = "____-__-__"
        End If
        txt.Text = lngVisit
    End Select
    
    chk(0).Value = IIf(msk.Text <> "" And msk.Text <> "____-__-__", 1, 0)
    chk(1).Value = IIf(Val(txt.Text) > 0, 1, 0)
    
    ReadPreVisiteDate = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    
End Function

Private Function ReadReportResult(Optional ByVal blnAdvice As Boolean = False) As Boolean
'------------------------------------------------------------------------------------------------------------------
    '���ܣ�
    '������
    '���أ�
    '------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rs As New ADODB.Recordset
    Dim lngLoop As Long

    strSQL = "Select y.��������,y.�쳣���,y.����id,y.�Ƿ񼲲�,y.��Ͻ���,y.�ο�����,y.����ʱ��,y.������� " & _
                "From " & _
                "( " & _
                "Select b.ҽ��id,d.����˳�� " & _
                "From ����ҽ����¼ a,�����Ŀҽ�� b,�����Ŀ�嵥 c,�����Ŀ���� d " & _
                "WHERE ������Դ=4 AND a.�Һŵ�=[1] and a.����id=[2] And a.���id Is Null " & _
                      "and a.id=b.ҽ��id " & _
                      "AND b.�嵥id=c.ID and c.������Ŀid=d.������Ŀid " & _
                ") x, " & _
                "( " & _
                "Select Distinct Nvl(a.���id,a.id) As ҽ��id,d.����id,d.��������,d.�쳣���,d.����id,d.�Ƿ񼲲�,d.��Ͻ���,d.��¼����,d.��¼���,d.�ο�����,d.����ʱ��,d.������� " & _
                "From ����ҽ����¼ a,����ҽ������ b,���˲������� c,�����Ա���� d " & _
                "WHERE a.������Դ=4 AND a.�Һŵ�=[1] and a.����id=[2] and d.��¼����=[3] and a.������� In ('C','D') " & _
                      "and a.id=b.ҽ��id and b.����id Is Not Null and c.������¼id=b.����id and d.����id=c.id " & _
                ") y " & _
                "where x.ҽ��id(+)=y.ҽ��id " & _
                "Order By  x.����˳��,y.��¼����,y.��¼���"
                    
    If blnAdvice = False Then
        
        Set rs = OpenSQLRecord(strSQL, "������", mstr�Һŵ�, mlng����id, 0)

        If rs.BOF = False Then
            Do While Not rs.EOF
                
                '���û��,����д
                If zlCommFun.NVL(rs("��������")) <> "" Then
                    vsf.RowData(vsf.Rows - 1) = zlCommFun.NVL(rs("����id"))
'                    vsf.TextMatrix(vsf.Rows - 1, 0) = vsf.Rows & "��"
                    vsf.TextMatrix(vsf.Rows - 1, mCol.��������) = zlCommFun.NVL(rs("��������"))
                    vsf.TextMatrix(vsf.Rows - 1, mCol.�쳣���) = zlCommFun.NVL(rs("�쳣���"))
                    vsf.TextMatrix(vsf.Rows - 1, mCol.����) = zlCommFun.NVL(rs("�Ƿ񼲲�"), 0)
                    vsf.TextMatrix(vsf.Rows - 1, mCol.��Ͻ���) = zlCommFun.NVL(rs("��Ͻ���"))
                    
                End If
                
                vsf.Rows = vsf.Rows + 1
                
                rs.MoveNext
            Loop
        End If
        
        If vsf.Rows > 1 Then vsf.Rows = vsf.Rows - 1
    
    Else
                                    
        Set rs = OpenSQLRecord(strSQL, "������", mstr�Һŵ�, mlng����id, 1)
        If rs.BOF = False Then
            
            If zlCommFun.NVL(rs("����ʱ��")) <> "" Then
                msk.Text = Format(zlCommFun.NVL(rs("����ʱ��")), "yyyy-MM-dd")
                chk(0).Value = 1
                cmd.Visible = True
            End If
            
            If zlCommFun.NVL(rs("�������"), 0) > 0 Then
                txt.Text = zlCommFun.NVL(rs("�������"), 0)
                chk(1).Value = 1
            End If
            
            Do While Not rs.EOF
                
                rtb.Text = rtb.Text & Trim(zlCommFun.NVL(rs("�ο�����"))) & vbCrLf
                
                rs.MoveNext
            Loop
        End If
    End If
    
    ReadReportResult = True
    
    Exit Function
    
errHand:
    
End Function

Private Sub chk_Click(Index As Integer)
    msk.Visible = (chk(0).Value = 1)
    cmd.Visible = (chk(0).Value = 1)
    txt.Visible = (chk(1).Value = 1)
    
    If (msk.Text = "" Or msk.Text = "____-__-__") And msk.Visible Then
        msk.Text = Format(zlDatabase.Currentdate + 90, "yyyy-MM-dd")
    End If
    
    Modified = True
End Sub

Private Sub cmd_Click()
    Dim rs As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "SELECT A.ID,B.ID AS �嵥id,E.�Ǽ�id," & _
                  "DECODE(A.���, 'C', '����', 'D', '���') AS ���," & _
                  "A.����," & _
                  "D.���� as ִ�п���," & _
                  "B.�����۸�,"
    strSQL = strSQL & _
                  "E.�������, " & _
                  "B.�ɼ���ʽid, " & _
                  "B.�ɼ�����id, " & _
                  "B.ִ�п���id, " & _
                  "B.��鲿λ, " & _
                  "B.�������, " & _
                  "B.���۸�,Decode(b.�����۸�,0,0,Null,0,10*B.���۸�/B.�����۸�) As �ۿ�," & _
                  "B.��鲿λid, " & _
                  "B.����걾,Decode(F.�����嵥id,0,0,Null,0,1) As ѡ�� " & _
             "FROM ������ĿĿ¼ A,�����Ŀ�嵥 B,���ű� D,�����Ա���� E,�����Ŀҽ�� F,���ǼǼ�¼ H " & _
            "WHERE B.ִ�п���id=D.ID(+) AND H.ID=E.�Ǽ�id And A.ID = B.������ĿID AND H.����=[1] AND E.�Ǽ�id=B.�Ǽ�id AND E.����id=F.����id AND F.�嵥id=B.ID AND F.����id=[2] and F.�����嵥id Is Null "
    
    strSQL = strSQL & " Order By A.����"

    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "���ר��ֽ", mstr�Һŵ�, mlng����id)
    If ShowFilterDiagBox(Me, cmd, "����,2700,0,0;���,900,0,1;ִ�п���,1500,0,0", "���ר��ֽ\������Ŀѡ��", "����б���ѡ��Ҫ����������Ŀ��", rsData, mrsCallBack, 8790, 4500, , , True) Then
        
    End If
        
End Sub

Private Sub msk_Change()
    Modified = True
End Sub

Private Sub msk_GotFocus()
    zlControl.TxtSelAll msk
End Sub

Private Sub rtb_Change()
    Modified = True
End Sub

Private Sub cbr_ButtonClick(Index As Integer, ByVal Button As MSComctlLib.Button)
    
    Dim lng����ID As Long
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    
    Select Case Index
    Case 0
        Select Case Button.Key
        Case "�ռ�"
            
            If MsgBox("���Ҫ�������Ŀ��������ȡ������", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
            
            If MsgBox("�Ƿ�Ҫ�滻ԭ�����ܼ���ۣ�", vbYesNo + vbQuestion + vbDefaultButton1, gstrSysName) = vbYes Then
                vsf.Rows = 2
                vsf.RowData(1) = 0
                vsf.Cell(flexcpText, 1, 1, vsf.Rows - 1, vsf.Cols - 1) = ""
            End If
                        
            Call ReadReportResult(False)
            Call ReadPreVisiteDate(2)
            
        Case "����"
            
            Call frmMedicalResult.ShowEdit(Me, mlng����id & "'0'" & mstr�Һŵ�)
            
        Case "���"
            
            vsf.Rows = 2
            vsf.RowData(1) = 0
            vsf.Cell(flexcpText, 1, 1, vsf.Rows - 1, vsf.Cols - 1) = ""
            
            Modified = True
        End Select
    Case 1
    
        Select Case Button.Key
        Case "�ռ�"
            If MsgBox("���Ҫ�������Ŀ��������ȡ������", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
            
            If MsgBox("�Ƿ�Ҫ�滻ԭ�����ܼ콨�飿", vbYesNo + vbQuestion + vbDefaultButton1, gstrSysName) = vbYes Then
                rtb.Text = ""
            End If
                        
            Call ReadReportResult(True)
            
        Case "����"
            rtb.Text = GetAdvice
        Case "���"
            rtb.Text = ""
        End Select
        
    End Select
End Sub


Private Sub rtb_GotFocus()
    zlCommFun.OpenIme True
End Sub

Private Sub rtb_LostFocus()
    zlCommFun.OpenIme False
End Sub

Private Sub txt_Change()
    Modified = True
End Sub

Private Sub txt_GotFocus()
    zlControl.TxtSelAll txt
End Sub


Private Sub UserControl_Initialize()
    
    Call InitControl
    
End Sub

Private Sub UserControl_Resize()
    
    On Error Resume Next
        
    With fra(0)
        .Left = 0
        .Top = -90
        .Width = UserControl.Width
    End With
    
    With fra(1)
        .Left = 0
        .Top = fra(0).Top + fra(0).Height - 90
        .Width = fra(0).Width
        .Height = UserControl.Height + 90 - fraOther.Height + 90 - fra(0).Height - 90
    End With
    
    With fraOther
        .Left = fra(1).Left
        .Top = fra(1).Top + fra(1).Height - 90
        .Width = fra(1).Width
    End With
    
    With pic(0)
        .Left = 30
        .Top = 120
        .Width = fra(0).Width - .Left - 45
    End With

    
    With vsf
        .Left = 15
        .Top = pic(0).Top + pic(0).Height
        .Width = fra(0).Width - .Left - 30
    End With
       
    With pic(1)
        .Left = 30
        .Top = 120
        .Width = fra(1).Width - .Left - 45
    End With

                    
    With rtb
        .Left = 15
        .Top = pic(1).Top + pic(1).Height
        .Width = fra(1).Width - .Left - 30
        .Height = fra(1).Height - .Top - 30
    End With
    
End Sub

Private Sub UserControl_Terminate()
    On Error Resume Next
    
    Set mobjParentObject = Nothing
End Sub

Private Sub vsf_AfterDeleteRow(ByVal Row As Long, ByVal Col As Long)
    Call ReadPreVisiteDate(2)
End Sub

Private Sub vsf_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    
    Modified = True
    
End Sub

Private Sub vsf_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    
    If mDispMode Then Exit Sub
    
    Select Case Col
    Case mCol.��������
        
        Call ShowOpenTree(1)
        
    End Select
    
End Sub

Private Sub vsf_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, ByVal ComboList As String, KeyCode As Integer, ByVal Shift As Integer, Cancel As Boolean)
    Dim strSvrText As String
    
    If mDispMode Then Exit Sub
    
    If KeyCode = vbKeyReturn Then
        '����2-�����͵����
        
        If InStr(vsf.EditText, "'") > 0 Then
            KeyCode = 0
            Exit Sub
        End If

        strSvrText = vsf.EditText
        Select Case ShowOpenList(vsf.EditText)
        Case 0
            'û��ƥ�����Ŀ
            vsf.Cell(flexcpData, Row, Col) = strSvrText
            
        Case 1
            'ѡȡ��һ����Ŀ
'            mblnChangeEdit = True
'            Call AdjustEnableState
        Case 2
            'ȡ���˱���ѡ��
            KeyCode = 0
            
            vsf.Cell(flexcpData, Row, Col) = vsf.Cell(flexcpData, Row, Col)
            vsf.TextMatrix(Row, Col) = vsf.Cell(flexcpData, Row, Col)

        End Select
    Else
'        mblnChangeEdit = True
'        Call AdjustEnableState
    End If
End Sub










