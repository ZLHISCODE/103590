VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm������Ŀѡ�������山 
   AutoRedraw      =   -1  'True
   Caption         =   "ҽ����Ŀѡ��"
   ClientHeight    =   5550
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7845
   Icon            =   "frm������Ŀѡ�������山.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   7845
   Begin VB.PictureBox picSplit 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   930
      Left            =   2340
      MousePointer    =   9  'Size W E
      ScaleHeight     =   930
      ScaleWidth      =   45
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1575
      Width           =   45
   End
   Begin VB.PictureBox picCmd 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   585
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   7845
      TabIndex        =   1
      Top             =   4965
      Width           =   7845
      Begin VB.CommandButton cmd���� 
         Caption         =   "����(&N)"
         Height          =   350
         Left            =   2625
         TabIndex        =   10
         ToolTipText     =   "���������ط�����Ŀ��������Ϣ�Ͷ���ҽ�ƻ���"
         Top             =   180
         Width           =   1100
      End
      Begin VB.CommandButton cmdRequery 
         Caption         =   "��Ŀ����"
         Height          =   350
         Left            =   1335
         TabIndex        =   5
         ToolTipText     =   "���������ط�����Ŀ��������Ϣ�Ͷ���ҽ�ƻ���"
         Top             =   180
         Width           =   1100
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "��ӡ�б�"
         Height          =   350
         Left            =   15
         TabIndex        =   4
         Top             =   180
         Width           =   1100
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "ȡ��(&C)"
         Height          =   350
         Left            =   6660
         TabIndex        =   3
         Top             =   180
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "ȷ��(&O)"
         Default         =   -1  'True
         Height          =   350
         Left            =   5400
         TabIndex        =   2
         Top             =   180
         Width           =   1100
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshGrid 
      Height          =   3990
      Left            =   3045
      TabIndex        =   0
      Top             =   390
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   7038
      _Version        =   393216
      FixedCols       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      BandDisplay     =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   45
      Top             =   3900
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������Ŀѡ�������山.frx":0E42
            Key             =   "Detail"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������Ŀѡ�������山.frx":1C94
            Key             =   "Class"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvwClass 
      Height          =   4050
      Left            =   0
      TabIndex        =   7
      Top             =   255
      Width           =   2970
      _ExtentX        =   5239
      _ExtentY        =   7144
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   494
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "img16"
      Appearance      =   1
   End
   Begin VB.Label lblDetail 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "��Ŀ��ϸ(&D)"
      Height          =   240
      Left            =   3060
      TabIndex        =   9
      Top             =   15
      Width           =   4710
   End
   Begin VB.Label lblClass 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "��Ŀ����(&K)"
      Height          =   240
      Left            =   15
      TabIndex        =   8
      Top             =   0
      Width           =   2970
   End
End
Attribute VB_Name = "frm������Ŀѡ�������山"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mint���� As Integer
Private mstrCode As String
Private mstrName As String
Private mblnOK As Boolean

Private mLocalCode As String 'ָ�����
Private mblnFirst As Boolean
'����Ŀ¼�ļ�����
Private Declare Function ExportKA02K3 Lib "YHMdcrAsistntSvr.dll" Alias "_ExportKA02K3@12" (ByVal strYab003 As String, ByVal strFileName As String, ByRef tmpStrut As Struct) As Boolean
'����Ŀ¼�ļ�����
Private Declare Function ExportKA06K1 Lib "YHMdcrAsistntSvr.dll" Alias "_ExportKA06K1@12" (ByVal strYab003 As String, ByVal strFileName As String, ByRef tmpStrut As Struct) As Boolean

Private Declare Function ExportKA03K1 Lib "YHMdcrAsistntSvr.dll" Alias "_ExportKA03K1@12" (ByVal strYab003 As String, ByVal strFileName As String, ByRef tmpStrut As Struct) As Boolean

Private Type Struct
    lngAppCode  As Long   '��־����ִ��״̬���롣����1ʱ��ʾ����ִ������������С��0ʱ��ʾ����ִ���쳣�����
    strErrMsg  As String  '������ִ��״̬����AppCodС��0ʱ����������ִ�е��쳣�������Ϣ��
End Type
Private mbln���� As Boolean
 

Private Sub cmdCancel_Click()
    mblnOK = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If Trim(mshGrid.TextMatrix(mshGrid.Row, 0)) = "" Then
        MsgBox "û��ѡ����Ŀ��", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '����ѡ����Ŀ����
    mstrCode = mshGrid.TextMatrix(mshGrid.Row, 0)
    mstrName = mshGrid.TextMatrix(mshGrid.Row, 1)
    mblnOK = True
    Unload Me
End Sub

Private Function Loadtree() As Boolean
    Dim rsTemp As New ADODB.Recordset, strTemp As String
    Dim tmpNode As Node
    mblnOK = False
    
    On Error GoTo ErrHand:
    
    'װ������
    gstrSQL = "" & _
        "   Select distinct  ����,���� From ҽ����Ŀ���� " & IIf(mbln����, " Where ����='61'", "")
    If rsTemp.State = adStateOpen Then rsTemp.Close
    rsTemp.Open gstrSQL, gcnOracle_CQYB, adOpenStatic, adLockReadOnly
    If rsTemp.EOF = True Then
        MsgBox "ҽ��ǰ�÷�������û��ҽ����Ŀ���࣬�޷�ѡ������ϵͳ����Ա��ϵ��", vbInformation, gstrSysName
        Exit Function
    End If
    
    tvwClass.Nodes.Clear
    Do Until rsTemp.EOF
        Set tmpNode = tvwClass.Nodes.Add(, 4, "K" & Nvl(rsTemp!����), "��" & Nvl(rsTemp("����")) & "��" & Nvl(rsTemp("����")), "Detail", "Detail")
        tmpNode.Sorted = True
        rsTemp.MoveNext
    Loop
    tvwClass.Nodes(1).Selected = True
    Call FillList
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    Call RestoreWinState(Me, App.ProductName)
    Loadtree = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Loadtree = False
End Function
Public Function GetCode(ByVal frmMain As Form, strCode As String, strName As String, Optional bln���� As Boolean = False) As Boolean
    '���ܣ���ȡ����
    '������
    '���أ��ɹ�����True
    mLocalCode = strCode
    mbln���� = bln����
    
    frm������Ŀѡ�������山.Show vbModal, frm������Ŀ
    '����ֵ
    If mblnOK = True Then
        strCode = mstrCode
        strName = mstrName
    End If
    GetCode = mblnOK
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub SetGrdColHead()
    With mshGrid
        .Clear
        .Rows = 2
        .Cols = 35
        .TextMatrix(0, 0) = "��Ʒ����"
        .TextMatrix(0, 1) = "��Ʒ��"
        .TextMatrix(0, 2) = "ҩƷͨ��������"
        .TextMatrix(0, 3) = "ҩƷͨ��Ӣ����"
        .TextMatrix(0, 4) = "��Ʒ������"
        .TextMatrix(0, 5) = "����"
        .TextMatrix(0, 6) = "��װ���"
        .TextMatrix(0, 7) = "ҽԺ�������"
        .TextMatrix(0, 8) = "ҽ������"
        .TextMatrix(0, 9) = "��С��װ��λ"
        .TextMatrix(0, 10) = "��С������λ"
        .TextMatrix(0, 11) = "ÿ���������"
        .TextMatrix(0, 12) = "ָ���۸�"
        .TextMatrix(0, 13) = "�б�۸�"
        .TextMatrix(0, 14) = "����֧���޼�1"
        .TextMatrix(0, 15) = "����֧���޼�2"
        .TextMatrix(0, 16) = "����֧���޼�3"
        .TextMatrix(0, 17) = "ʵ��ִ�м۸�"
        .TextMatrix(0, 18) = "�Ը�����1"
        .TextMatrix(0, 19) = "�Ը�����2"
        .TextMatrix(0, 20) = "�Ը�����3"
        .TextMatrix(0, 21) = "�Ը�����4"
        .TextMatrix(0, 22) = "�Ը�����5"
        .TextMatrix(0, 23) = "�Ը�����6"
        .TextMatrix(0, 24) = "�Ը�����7"
        .TextMatrix(0, 25) = "�Ը�����8"
        .TextMatrix(0, 26) = "�Ը�����9"
        .TextMatrix(0, 27) = "�Ը�����10"
        .TextMatrix(0, 28) = "�Ը�����11"
        .TextMatrix(0, 29) = "�Ը�����12"
        .TextMatrix(0, 30) = "��׼���"
        .TextMatrix(0, 31) = "���������1"
        .TextMatrix(0, 32) = "ƴ��������1"
        .TextMatrix(0, 33) = "��ע"
        .TextMatrix(0, 34) = "Ŀ¼����"
    End With

End Sub
Private Sub FillList()
    '���ܣ���ʾ��ǰ����µ�ҽ����ϸ
    Dim rsTemp As New ADODB.Recordset
    Dim lst As ListItem, fld As ADODB.Field
    Dim str������ As String, blnColSet As Boolean
    Dim lngCol  As Long
    Dim varValue As Variant
    
    Me.MousePointer = vbHourglass
    
    On Error GoTo errHandle
    With tvwClass.SelectedItem
        str������ = IIf(.Key = "Root", "", " And ҽԺ������� ='" & Mid(.Key, 2) & "'")
    End With
    
    
    rsTemp.CursorLocation = adUseClient
    
'    gstrSQL = " select  ��Ʒ����,  ҽԺ�������, ҽ������, ҩƷͨ��������, ҩƷͨ��Ӣ����,��Ʒ��, ��Ʒ������, ������Ŀ���㷽ʽ, ������ʶ, ҽ����ʶ, �Ƿ񴦷���ҩ, ҩƷ��Ӧ֢, ����ҽ��, ����Ȩ��, ����, ��װ���, " & _
             "         ��С��װ��λ, ��С������λ, ÿ���������, ָ���۸�, �б�۸�, ����֧���޼�1, ����֧���޼�2, ����֧���޼�3, ʵ��ִ�м۸�, �Ը�����1, �Ը�����2, �Ը�����3, �Ը�����4, �Ը�����5, �Ը�����6, �Ը�����7, �Ը�����8,  " & _
             "         �Ը�����9, �Ը�����10, �Ը�����11, �Ը�����12, ҽԺʹ��״̬, ����ʹ��״̬, ��׼���,  " & _
             "         ���������1, ���������2, ���������3, ƴ��������1, ƴ��������2, ƴ��������3, ��ע, ҽ���������,������׼���, ҽ�ƻ������, " & _
             "          �޸�ʱ��, Ŀ¼����  " & _
             "  from ҽ��������ĿĿ¼" & _
             "  where 1=1 " & str������
    
    gstrSQL = " select  ��Ʒ����,��Ʒ��, ҩƷͨ��������, ҩƷͨ��Ӣ����, ��Ʒ������, ����, ��װ���,ҽԺ�������, ҽ������, " & _
             "         ��С��װ��λ, ��С������λ, ÿ���������, ָ���۸�, �б�۸�, ����֧���޼�1, ����֧���޼�2, ����֧���޼�3, ʵ��ִ�м۸�, �Ը�����1, �Ը�����2, �Ը�����3, �Ը�����4, �Ը�����5, �Ը�����6, �Ը�����7, �Ը�����8,  " & _
             "         �Ը�����9, �Ը�����10, �Ը�����11, �Ը�����12, ��׼���,  " & _
             "         ���������1, ƴ��������1, ��ע, " & _
             "         Ŀ¼����  " & _
             "  from ҽ��������ĿĿ¼" & _
             "  where 1=1 " & str������
    
    rsTemp.Open gstrSQL, gcnOracle_CQYB, adOpenStatic, adLockReadOnly
    
    If rsTemp.RecordCount = 0 Then
        '������ͷ
        Call SetGrdColHead
    Else
        Set mshGrid.DataSource = rsTemp
    End If
    Me.MousePointer = vbDefault
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
 '   LockWindowUpdate 0
    Me.MousePointer = vbDefault
End Sub

Private Sub cmdPrint_Click()
    If gstrUserName = "" Then Call GetUserInfo
    subPrint 1
End Sub

Private Sub subPrint(bytMode As Byte)
    '����:���д�ӡ,Ԥ���������EXCEL
    '����:bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    Dim objPrint As New zlPrint1Grd
    Dim objRow As New zlTabAppRow
    Dim nod As Node
    
    Set nod = tvwClass.SelectedItem
    Set objPrint.Body = mshGrid
    objPrint.Title.Text = "������Ŀ"
    
    objRow.Add "ҽ�����ࣺ" & nod.Text
    objPrint.UnderAppRows.Add objRow
    
    Set objRow = New zlTabAppRow
    objRow.Add "��ӡ�ˣ�" & gstrUserName
    objRow.Add "��ӡʱ�䣺" & Format(zlDatabase.Currentdate, "yyyy-MM-dd")
    objPrint.BelowAppRows.Add objRow
    If bytMode = 1 Then
      Select Case zlPrintAsk(objPrint)
          Case 1
               zlPrintOrView1Grd objPrint, 1
          Case 2
              zlPrintOrView1Grd objPrint, 2
          Case 3
              zlPrintOrView1Grd objPrint, 3
      End Select
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
End Sub
Private Sub cmdRequery_Click()
    Dim strInPut As String
    Dim rsTemp As New ADODB.Recordset
    Dim bln���� As Boolean
    
    If MsgBox("���������ܻỨ�Ƚϳ���ʱ�䣬�Ƿ������" & vbCrLf & vbCrLf & "����ע�⣬������ֻ����ҽ����Ŀ��ϸ������������Ӧ��ϵ��", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Exit Sub
    End If
    
    If MsgBox("���ν����ز�����", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        bln���� = False
    Else
        bln���� = True
    End If
    
    MousePointer = vbHourglass
    zlCommFun.ShowFlash "ҽ����Ŀѡ�����ڶ�ȡ���ļ��������ȡ������Ŀ��ϸ���ù�����һ���ϳ��Ĺ��̣���ȴ�......��"
    
    If InitInfor_�����山.����������� = "" Then
        ShowMsgbox "����������벻��Ϊ��!"
        Exit Sub
    End If
    
    picCmd.Enabled = False
    tvwClass.Enabled = False
    '��鱾����ȫ�����»�����������(�޸�����)
    If Not bln���� Then
        Call Get����Ŀ¼
    End If
    
    '���ﲡ������:
    Call Get����Ŀ¼
    
    'סԺ����
    Call GetסԺ����Ŀ¼
    
    Me.Caption = "���������ϸ���ݣ����Ժ�..."
    '����װ����ϸ
    Call FillList
    zlCommFun.StopFlash
    MousePointer = vbDefault
    Me.Caption = "ҽ����Ŀѡ��"
    picCmd.Enabled = True
    tvwClass.Enabled = True
End Sub

Private Function Get����Ŀ¼() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    '����:����������ĿĿ¼
    '����:�����ɹ�True,���򷵻�False
    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim strFile As String
    Dim tmpStrut As Struct
    Dim objFile As New FileSystemObject
    Dim objText As TextStream
    Dim StrSQL As String
    Dim strHead As String
    Dim strArr
    Dim strText As String
    Dim strTemp As String
    Dim i As Long
    Dim lngRow As Long
    
    tmpStrut.strErrMsg = Space(5000)
    strFile = App.Path & "\ҽ��"
    Get����Ŀ¼ = False

    Err = 0
    On Error GoTo ErrHand:
    If Not objFile.FolderExists(strFile) Then
        '�������ļ��У��贴��
        objFile.CreateFolder strFile
    End If
    
    strFile = strFile & "\ҽ�Ʒ���Ŀ¼.Txt"
    If Not objFile.FileExists(strFile) Then
        objFile.CreateTextFile strFile
    End If
    
    
    ExportKA02K3 InitInfor_�����山.�����������, strFile, tmpStrut
    
    If tmpStrut.lngAppCode < 0 Then
        ShowMsgbox tmpStrut.strErrMsg
        Exit Function
    End If
    
    strHead = "ZL_ҽ��������ĿĿ¼_UPDATE("
    Set objText = objFile.OpenTextFile(strFile)
    lngRow = 1
    Do While Not objText.AtEndOfStream
        strTemp = Trim(objText.ReadLine)
        strArr = Split(strTemp, vbTab)
        StrSQL = ""
        For i = 0 To UBound(strArr)
            If InStr(1, strArr(i), "'") <> 0 Then
                strTemp = Replace(strArr(i), "'", "��")
                strArr(i) = strTemp
            End If
            If Trim(strArr(i)) = "" Then
                    StrSQL = StrSQL & ",null"
            Else
                If i < 22 Or i >= 41 And i <= 50 Or i >= 52 Then
                    If i >= 43 And i <= 48 Then
                        StrSQL = StrSQL & ",'" & UCase(strArr(i)) & "'"
                    Else
                        StrSQL = StrSQL & ",'" & strArr(i) & "'"
                    End If
                ElseIf i = 51 Then
                    '�޸�ʱ��
                    StrSQL = StrSQL & ",to_date('" & Format(strArr(i), "yyyy-mm-dd") & "','yyyy-mm-dd')"
                Else
                    StrSQL = StrSQL & "," & Val(strArr(i))
                End If
            End If
        Next
        If StrSQL <> "" Then
            StrSQL = strHead & Mid(StrSQL, 2) & ")"
            gcnOracle_CQYB.Execute StrSQL, , adCmdStoredProc
            DoEvents
        End If
        
        Me.Caption = "ҽ��������Ŀ����:�Ѿ������� " & lngRow & "����¼"
        lngRow = lngRow + 1
    Loop
    objText.Close
    Get����Ŀ¼ = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    
End Function
Private Function Get����Ŀ¼() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    '����:��������Ŀ¼
    '����:�����ɹ�True,���򷵻�False
    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim strFile As String
    Dim tmpStrut As Struct
    Dim objFile As New FileSystemObject
    Dim objText As TextStream
    Dim StrSQL As String
    Dim strSQL1 As String
    Dim strHead As String
    Dim strArr
    Dim strText As String
    Dim strTemp As String
    Dim i As Long
    Dim lngRow As Long
    
    tmpStrut.strErrMsg = Space(5000)
    strFile = App.Path & "\ҽ��"
    Get����Ŀ¼ = False
    Err = 0
    On Error GoTo ErrHand:
    If Not objFile.FolderExists(strFile) Then
        '�������ļ��У��贴��
        objFile.CreateFolder strFile
    End If
    
    strFile = strFile & "\ҽ�Ʋ���Ŀ¼.Txt"
    If Not objFile.FileExists(strFile) Then
        objFile.CreateTextFile strFile
    End If
    
    
    ExportKA06K1 InitInfor_�����山.�����������, strFile, tmpStrut
    
    If tmpStrut.lngAppCode < 0 Then
        ShowMsgbox tmpStrut.strErrMsg
        Exit Function
    End If
    
    strHead = "ZL_ҽ������Ŀ¼_UPDATE("
    Set objText = objFile.OpenTextFile(strFile)
    lngRow = 1
    Do While Not objText.AtEndOfStream
        strTemp = Trim(objText.ReadLine)
        strArr = Split(strTemp, vbTab)
        StrSQL = ""
        For i = 0 To UBound(strArr)
            If Trim(strArr(i)) = "" Then
                    StrSQL = StrSQL & ",null"
            Else
                    StrSQL = StrSQL & ",'" & strArr(i) & "'"
            End If
            If i >= 5 Then Exit For
        Next
        If StrSQL <> "" Then
            StrSQL = strHead & "1," & Mid(StrSQL, 2) & ")"
            gcnOracle_CQYB.Execute StrSQL, , adCmdStoredProc
        End If
        Me.Caption = "���ﲡ��Ŀ¼����:�Ѿ������� " & lngRow & "����¼"
        lngRow = lngRow + 1
    Loop
    
    objText.Close
    Get����Ŀ¼ = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    
End Function
Private Function GetסԺ����Ŀ¼() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    '����:����סԺ����Ŀ¼
    '����:�����ɹ�True,���򷵻�False
    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim strFile As String
    Dim tmpStrut As Struct
    Dim objFile As New FileSystemObject
    Dim objText As TextStream
    Dim StrSQL As String
    Dim strSQL1 As String
    Dim strHead As String
    Dim strArr
    Dim strText As String
    Dim strTemp As String
    Dim i As Long
    Dim lngRow As Long
    
    tmpStrut.strErrMsg = Space(5000)
    strFile = App.Path & "\ҽ��"
    GetסԺ����Ŀ¼ = False
    Err = 0
    On Error GoTo ErrHand:
    If Not objFile.FolderExists(strFile) Then
        '�������ļ��У��贴��
        objFile.CreateFolder strFile
    End If
    
    strFile = strFile & "\סԺ����Ŀ¼.Txt"
    If Not objFile.FileExists(strFile) Then
        objFile.CreateTextFile strFile
    End If
    
    
    ExportKA03K1 InitInfor_�����山.�����������, strFile, tmpStrut
    
    If tmpStrut.lngAppCode < 0 Then
        ShowMsgbox tmpStrut.strErrMsg
        Exit Function
    End If
    
    strHead = "ZL_ҽ������Ŀ¼_UPDATE("
    Set objText = objFile.OpenTextFile(strFile)
    lngRow = 1
    
    '1   string  20      ���ֱ���
    '2   string  100     ��������
    '3   string  6       ��ϸ��ʶ���������
    '4   string  10      ������
    
    Do While Not objText.AtEndOfStream
        strTemp = Trim(objText.ReadLine)
        strArr = Split(strTemp, vbTab)
        
        If strTemp <> "" Then
            'ZL_ҽ������Ŀ¼_UPDATE
            StrSQL = "ZL_ҽ������Ŀ¼_UPDATE("
            '  ����_In         In ҽ������Ŀ¼.����%Type,
            StrSQL = StrSQL & "" & 2 & ","
            '  ����_In         In ҽ������Ŀ¼.����%Type,
            StrSQL = StrSQL & "'" & strArr(0) & "',"
            '  ����_In         In ҽ������Ŀ¼.����%Type,
            StrSQL = StrSQL & "'" & strArr(1) & "',"
            '  ֧�����_In     In ҽ������Ŀ¼.֧�����%Type,
            StrSQL = StrSQL & "'" & strArr(2) & "',"
            '  ������_In       In ҽ������Ŀ¼.������%Type,
            StrSQL = StrSQL & "'" & strArr(3) & "',"
            '  ���ֽ���취_In In ҽ������Ŀ¼.���ֽ���취%Type,
            StrSQL = StrSQL & "NULL,"
            '  ���칹������_In In ҽ������Ŀ¼.���칹������%Type
            StrSQL = StrSQL & "NULL)"
            gcnOracle_CQYB.Execute StrSQL, , adCmdStoredProc
        End If
        Me.Caption = "סԺ����Ŀ¼����:�Ѿ������� " & lngRow & "����¼"
        lngRow = lngRow + 1
    Loop
    
    objText.Close
    GetסԺ����Ŀ¼ = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    
End Function
Private Sub cmd����_Click()
    Dim blnReturn As Boolean
    If Me.tvwClass.SelectedItem Is Nothing Then Exit Sub
    
    blnReturn = frm������Ŀ�༭.EditCard(Me, "", Mid(Me.tvwClass.SelectedItem.Key, 2))
    If blnReturn = False Then Exit Sub
     Call FillList
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    If Loadtree = False Then
        Exit Sub
    End If
End Sub
Private Sub Form_Load()
    mblnFirst = True
End Sub

Private Sub Form_Resize()
    lblClass.Top = 0: lblClass.Left = 0: lblClass.Width = tvwClass.Width
    
    On Error Resume Next
    
    tvwClass.Left = 0: tvwClass.Top = lblClass.Top + lblClass.Height
    tvwClass.Height = Me.ScaleHeight - lblClass.Height - picCmd.Height
    
    picSplit.Top = tvwClass.Top
    picSplit.Left = tvwClass.Left + tvwClass.Width
    picSplit.Height = tvwClass.Height
    
    lblDetail.Top = lblClass.Top
    If tvwClass.Visible = True Then
        lblDetail.Left = picSplit.Left + picSplit.Width
    Else
        lblDetail.Left = 0
    End If
    lblDetail.Width = Me.ScaleWidth - lblDetail.Left
        
    With mshGrid
        .Top = tvwClass.Top
        .Left = lblDetail.Left
        .Width = lblDetail.Width
        .Height = tvwClass.Height
    End With
End Sub

Private Sub picCmd_Resize()
    cmdCancel.Left = picCmd.ScaleWidth - cmdCancel.Width * 1.4
    cmdOK.Left = cmdCancel.Left - cmdOK.Width * 1.25
    cmdPrint.Top = cmdOK.Top
    cmdRequery.Top = cmdOK.Top
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName
End Sub

Private Sub mshgrid_DblClick()
    cmdOK_Click
End Sub

Private Sub picSplit_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        If tvwClass.Width + x < 1000 Or mshGrid.Width - x < 1000 Then Exit Sub
        picSplit.Left = picSplit.Left + x
        lblClass.Width = lblClass.Width + x
        tvwClass.Width = tvwClass.Width + x
        
        lblDetail.Left = lblDetail.Left + x
        lblDetail.Width = lblDetail.Width - x
        
        mshGrid.Left = mshGrid.Left + x
        mshGrid.Width = mshGrid.Width - x
    End If
End Sub

Private Sub tvwClass_NodeClick(ByVal Node As MSComctlLib.Node)
    Call FillList
End Sub





