VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm������Ŀѡ�񱱾� 
   AutoRedraw      =   -1  'True
   Caption         =   "ҽ����Ŀѡ��"
   ClientHeight    =   4935
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7845
   Icon            =   "frm������Ŀѡ�񱱾�.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   7845
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
      TabIndex        =   4
      Top             =   4350
      Width           =   7845
      Begin VB.CommandButton cmdRequery 
         Caption         =   "������ϸ"
         Height          =   350
         Left            =   3900
         TabIndex        =   11
         Top             =   150
         Visible         =   0   'False
         Width           =   1100
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "��ӡ�б�"
         Height          =   350
         Left            =   2790
         TabIndex        =   9
         Top             =   150
         Width           =   1100
      End
      Begin VB.TextBox txtFind 
         Height          =   300
         Left            =   1200
         MaxLength       =   6
         TabIndex        =   6
         Top             =   175
         Width           =   1335
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "ȡ��(&C)"
         Height          =   350
         Left            =   6660
         TabIndex        =   8
         Top             =   150
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "ȷ��(&O)"
         Default         =   -1  'True
         Height          =   350
         Left            =   5400
         TabIndex        =   7
         Top             =   150
         Width           =   1100
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "��ϸ����(&F)"
         Height          =   180
         Index           =   0
         Left            =   150
         TabIndex        =   5
         Top             =   240
         Width           =   990
      End
   End
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
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1590
      Width           =   45
   End
   Begin MSComctlLib.ListView lvwDetail 
      Height          =   4050
      Left            =   3060
      TabIndex        =   2
      Top             =   270
      Width           =   4710
      _ExtentX        =   8308
      _ExtentY        =   7144
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "img16"
      SmallIcons      =   "img16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   45
      Top             =   3915
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
            Picture         =   "frm������Ŀѡ�񱱾�.frx":0E42
            Key             =   "Detail"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������Ŀѡ�񱱾�.frx":1C94
            Key             =   "Class"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvwClass 
      Height          =   4050
      Left            =   0
      TabIndex        =   10
      Top             =   270
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
   Begin VB.Label lblClass 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "��Ŀ����(&K)"
      Height          =   240
      Left            =   15
      TabIndex        =   0
      Top             =   15
      Width           =   2970
   End
   Begin VB.Label lblDetail 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "��Ŀ��ϸ(&D)"
      Height          =   240
      Left            =   3060
      TabIndex        =   1
      Top             =   30
      Width           =   4710
   End
End
Attribute VB_Name = "frm������Ŀѡ�񱱾�"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mint���� As Integer
Private mstrCode As String
Private mstrName As String
Private mcnYB As New ADODB.Connection
Private mobjStream As TextStream
Private mobjFileSystem As New FileSystemObject
Private mstrErr As String
Private Const strFile = "C:\DOWNLOAD_ERR.LOG"
Private mErrFile As TextStream
Private mblnOK As Boolean

Private Sub cmdCancel_Click()
    mblnOK = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim rsTemp As New ADODB.Recordset
    If lvwDetail.SelectedItem Is Nothing Then
        MsgBox "û��ѡ����Ŀ��", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '����ѡ����Ŀ����
    mstrCode = lvwDetail.SelectedItem.Text
    
    mblnOK = True
    Unload Me
End Sub

Private Sub GetValueByCol(ByVal strColumnName As String, strValue As String)
    Dim lngCount As Long, lngIndex As Long

    For lngCount = 1 To lvwDetail.ColumnHeaders.Count
        If lvwDetail.ColumnHeaders(lngCount).Text = strColumnName Then
            lngIndex = lngCount
            Exit For
        End If
    Next
    
    If lngIndex > 0 Then
        strValue = lvwDetail.SelectedItem.SubItems(lngIndex - 1)
    End If
End Sub

Public Function GetCode(strCode As String, STRNAME As String, ByVal int���� As Integer) As Boolean
'���ܣ����һ���շ���Ŀ��ҽ������
'������strCode ����Ϊ��������������
'���أ��ɹ�����True
    Dim rsTemp As New ADODB.Recordset, strTemp As String
    Dim strServer As String, strUser As String, strPass As String
    Dim nod As Node
    
    mblnOK = False
    mstrCode = strCode
    mstrName = STRNAME
    mint���� = int����
    
    On Error GoTo ErrH
    
    '���ȶ���������������
    gstrSQL = "Select ������,����ֵ From ���ղ��� Where ����=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, int����)
    Do Until rsTemp.EOF
        strTemp = IIf(IsNull(rsTemp("����ֵ")), "", rsTemp("����ֵ"))
        Select Case rsTemp("������")
            Case "ҽ��������"
                strServer = strTemp
            Case "ҽ���û���"
                strUser = strTemp
            Case "ҽ���û�����"
                strPass = strTemp
        End Select
        rsTemp.MoveNext
    Loop
    If OraDataOpen(mcnYB, strServer, strUser, strPass) = False Then
        Exit Function
    End If
    
    '���շ������з�����ʾҽ����Ŀ
    gstrSQL = "" & _
        " SELECT *" & _
        " FROM (" & _
        "     SELECT B.����,B.����,'0' AS �ϼ�ID" & _
        "     FROM ָ������ A,ָ����ϵ���ձ� B" & _
        "     WHERE A.����='�շ����' AND A.���=B.���" & _
        "     AND SUBSTR(B.����,3,2)='00'" & _
        "     Union" & _
        "     SELECT B.����,B.����,SUBSTR(B.����,1,2)||'00' AS �ϼ�ID" & _
        "     FROM ָ������ A,ָ����ϵ���ձ� B" & _
        "     WHERE A.����='�շ����' AND A.���=B.���" & _
        "     AND SUBSTR(B.����,3,2)<>'00')" & _
        " ORDER BY �ϼ�ID,����"
    If rsTemp.State = adStateOpen Then rsTemp.Close
    rsTemp.Open gstrSQL, mcnYB, adOpenStatic, adLockReadOnly
    
    tvwClass.Nodes.Clear
    Do Until rsTemp.EOF
        If rsTemp("�ϼ�id") = 0 Then
            Set nod = tvwClass.Nodes.Add(, , "K" & rsTemp("����"), "��" & rsTemp("����") & "��" & rsTemp("����"), "Class", "Class")
        Else
            Set nod = tvwClass.Nodes.Add("K" & rsTemp("�ϼ�ID"), tvwChild, "K" & rsTemp("����"), "��" & rsTemp("����") & "��" & rsTemp("����"), "Class", "Class")
        End If
        nod.Sorted = True
        rsTemp.MoveNext
    Loop
    
    tvwClass.Nodes(2).Selected = True
    Call FillList
    
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    Call RestoreWinState(Me, App.ProductName)
    
    frm������Ŀѡ�񱱾�.Show vbModal, frm������Ŀ
    '����ֵ
    If mblnOK = True Then strCode = mstrCode
    GetCode = mblnOK
    Exit Function
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

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
        str������ = Mid(.Text, 2, InStr(.Text, "��") - 2)
    End With
    
    rsTemp.CursorLocation = adUseClient
    '��ʱ���б���ˢ��
    LockWindowUpdate lvwDetail.hwnd
    lvwDetail.ListItems.Clear
    
    If str������ < "0400" Then
        '��ǰѡ���ǵ�һ��ҩƷ���
        Me.lvwDetail.Tag = "Y"
        gstrSQL = "" & _
            " Select A.����,A.��Ŀ,A.����,A.������,A.������λ AS ��λ,B.���� As ���ⲡ,H.���� AS ��Ŀ�ȼ�,A.��׼����,A.�Ը�����,0 �޼�," & _
            " C.���� AS ����ҩ,F.���� AS ����,A.�÷�,A.�ճ�������,D.���� AS ҩƷ����,G.���� AS ����,E.���� AS ʹ�����Ƶȼ�,A.��ע,A.��Ч����" & _
            " From ҩƷĿ¼ A," & _
            "      (Select B.����,B.����" & _
            "      FROM ָ������ A,ָ����ϵ���ձ� B" & _
            "      Where A.����='������ҩ��ʶ' and A.���=B.���) B," & _
            "      (Select B.����,B.����" & _
            "      FROM ָ������ A,ָ����ϵ���ձ� B" & _
            "      Where A.����='����ҩ��־' and A.���=B.���) C," & _
            "      (Select B.����,B.����" & _
            "      FROM ָ������ A,ָ����ϵ���ձ� B" & _
            "      Where A.����='ҩƷ����' and A.���=B.���) D," & _
            "      (Select B.����,B.����" & _
            "      FROM ָ������ A,ָ����ϵ���ձ� B" & _
            "      Where A.����='ʹ�����Ƶȼ�' and A.���=B.���) E,"
        gstrSQL = gstrSQL & _
            "      (Select B.����,B.����" & _
            "      FROM ָ������ A,ָ����ϵ���ձ� B" & _
            "      Where A.����='����' and A.���=B.���) F," & _
            "      (Select B.����,B.����" & _
            "      FROM ָ������ A,ָ����ϵ���ձ� B" & _
            "      Where A.����='����' and A.���=B.���) G," & _
            "      (Select B.����,B.����" & _
            "      FROM ָ������ A,ָ����ϵ���ձ� B" & _
            "      Where A.����='�շ���Ŀ�ȼ�' and A.���=B.���) H" & _
            " Where A.���ⲡ =B.����(+) And A.����ҩ=C.����(+) And A.ҩƷ���� =D.����(+)" & _
            " And A.ʹ�����Ƶȼ�=E.����(+) And A.����=F.����(+) And A.����=G.����(+) AND A.ҩƷ�ȼ�=H.����(+)" & _
            " And A.�շ����='" & str������ & "'"
    Else
        '��ǰѡ���ǵ�һ���������
        Me.lvwDetail.Tag = "Z"
        gstrSQL = "" & _
            " Select A.����,A.����,A.������,A.��λ,B.���� AS ���ⲡ,C.���� AS ��Ŀ�ȼ�,A.��׼����,A.�Ը�����,A.�޼�,A.��ע,A.��Ч����" & _
            "      From ����Ŀ¼ A," & _
            "      (Select B.����,B.����" & _
            "      FROM ָ������ A,ָ����ϵ���ձ� B" & _
            "      Where A.����='������ҩ��ʶ' and A.���=B.���) B," & _
            "      (Select B.����,B.����" & _
            "      FROM ָ������ A,ָ����ϵ���ձ� B" & _
            "      Where A.����='�շ���Ŀ�ȼ�' and A.���=B.���) C" & _
            " Where A.���ⲡ =B.����(+) And A.��Ŀ�ȼ�=C.����(+)" & _
            " And A.�շ����='" & str������ & "'"
    End If
    rsTemp.Open gstrSQL, mcnYB, adOpenStatic, adLockReadOnly
    
    If lvwDetail.ColumnHeaders.Count <> rsTemp.Fields.Count Then
        '���´����ͷ
        blnColSet = True
        lvwDetail.ColumnHeaders.Clear
        For Each fld In rsTemp.Fields
            lvwDetail.ColumnHeaders.Add , , fld.Name, 1000
        Next
    End If
        
    Do Until rsTemp.EOF
        Set lst = lvwDetail.ListItems.Add(, "K" & rsTemp("����"), rsTemp("����"), "Detail", "Detail")
        
        '����ListView�����������ݿ�ȡ��
        For lngCol = 2 To lvwDetail.ColumnHeaders.Count
            varValue = rsTemp(lvwDetail.ColumnHeaders(lngCol).Text).Value
            lst.SubItems(lngCol - 1) = IIf(IsNull(varValue), "", varValue)
        Next
        rsTemp.MoveNext
    Loop
    If blnColSet = True Then
        '���¶��н����˴���
        If lvwDetail.ListItems.Count > 0 Then Call zlControl.LvwSetColWidth(lvwDetail)
    End If
    LockWindowUpdate 0
    Me.MousePointer = vbDefault
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    LockWindowUpdate 0
    Me.MousePointer = vbDefault
End Sub

Private Sub cmdPrint_Click()
'����:���д�ӡ,Ԥ���������EXCEL
'����:bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    If gstrUserName = "" Then Call GetUserInfo
    Dim objPrint As New zlPrintLvw
    
    
    objPrint.Title.Text = "������Ŀ"
    Set objPrint.Body.objData = lvwDetail
    objPrint.UnderAppItems.Add "ҽ�����ࣺ" & tvwClass.SelectedItem.Text
    objPrint.BelowAppItems.Add "��ӡ�ˣ�" & gstrUserName
    objPrint.BelowAppItems.Add "��ӡʱ�䣺" & Format(zlDatabase.Currentdate, "yyyy��MM��dd��")
    Select Case zlPrintAsk(objPrint)
        Case 1
             zlPrintOrViewLvw objPrint, 1
        Case 2
            zlPrintOrViewLvw objPrint, 2
        Case 3
            zlPrintOrViewLvw objPrint, 3
    End Select

End Sub

Private Sub cmdRequery_Click()
    Dim blnExist As Boolean                 'Ŀ¼�Ƿ����
    Dim blnTrans As Boolean, blnȫ�� As Boolean, blnSuccess As Boolean
    Dim strҽ����ĿĿ¼ As String
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    'ѡ��Ŀ¼��������ȫ�����»����������£��Զ�ѡ���Ӧ���ļ�����
    
    '�жϸ���Ŀ¼�Ƿ����趨
    gstrSQL = "Select ����ֵ From ���ղ��� Where ������='ҽ����ĿĿ¼'"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡҽ����ĿĿ¼")
    If Not rsTemp.EOF Then
        If Nvl(rsTemp!����ֵ) <> "" Then
            blnExist = True
            strҽ����ĿĿ¼ = rsTemp!����ֵ
        End If
    End If
    If blnExist = False Then
        MsgBox "�����ڱ������Ĳ��������У�����ҽ����ĿĿ¼��", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '��ʼ��ҽ���м������
    If Not ҽ����ʼ��_����(True) Then Exit Sub
    
    '�жϸ���ģʽ
    blnȫ�� = True
    gstrSQL = "Select 1 From ҩƷĿ¼ Where Rownum<2"
    If rsTemp.State = 1 Then rsTemp.Close
    rsTemp.Open gstrSQL, mcnYB
    If Not rsTemp.EOF Then
        blnȫ�� = (MsgBox("�����Ѵ���ҽ����Ŀ�����ν�Ҫ�����������£�����Ǳ�ʾ�������£�������ʾȫ������", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo)
    End If
    
    'ɾ����־�ļ�
    Set mobjFileSystem = New FileSystemObject
    If mobjFileSystem.FileExists(strFile) Then mobjFileSystem.DeleteFile (strFile)
    Set mErrFile = mobjFileSystem.CreateTextFile(strFile)
    
    mcnYB.BeginTrans
    blnTrans = True
    '׼������ҩƷĿ¼������
    blnSuccess = ҩƷĿ¼_����(blnȫ��, strҽ����ĿĿ¼)
    If blnSuccess = False Then
        mErrFile.Close
        Call zlCommFun.StopFlash
        mcnYB.RollbackTrans
        Me.Caption = "ҽ����Ŀѡ��"
        Exit Sub
    End If
    '׼����������Ŀ¼������
    blnSuccess = ����Ŀ¼_����(blnȫ��, strҽ����ĿĿ¼)
    If blnSuccess = False Then
        mErrFile.Close
        Call zlCommFun.StopFlash
        mcnYB.RollbackTrans
        Me.Caption = "ҽ����Ŀѡ��"
        Exit Sub
    End If
    
    mcnYB.CommitTrans
    mErrFile.Close
    Call zlCommFun.StopFlash
    Me.Caption = "ҽ����Ŀѡ��"
    blnTrans = False
    
    '����װ����ϸ
    Call FillList
    MousePointer = vbDefault
    Me.Caption = "ҽ����Ŀѡ��"
    Exit Sub
errHand:
    If ErrCenter = 1 Then Resume
    If blnTrans Then mcnYB.RollbackTrans
    mErrFile.Close
    Call zlCommFun.StopFlash
End Sub

Private Function ҩƷĿ¼_����(ByVal blnȫ�� As Boolean, ByVal strĿ¼ As String) As Boolean
    Const strҩƷĿ¼_ȫ�� As String = "ypml_all.txt"
    Const strҩƷ����_ȫ�� As String = "ypymml_all.txt"
    Const strҩƷĿ¼_���� As String = "ypml.txt"
    Const strҩƷ����_���� As String = "ypymml.txt"
    Const int���� As Integer = 1
    Const int���� As Integer = 2
    Const strSplit As String = "|"
    Dim strProcedure As String
    Dim arrLine
    Dim lngCol As Long, lngCols As Long
    Dim blnExist As Boolean
    Dim strText As String
    Dim strFileName As String
    Dim strExecute As String
    Dim objStream As TextStream
    Dim objFileSystem As New FileSystemObject
    On Error GoTo errHand
    '�ȴ���ҩƷĿ¼
    '�����ļ�
    strFileName = strĿ¼ & "\" & IIf(blnȫ��, strҩƷĿ¼_ȫ��, strҩƷĿ¼_����)
    If Not objFileSystem.FileExists(strFileName) Then
        blnExist = False
    Else
        '���ļ����и��²���
        blnExist = True
        Set objStream = objFileSystem.OpenTextFile(strFileName)
    End If
    
    Call zlCommFun.ShowFlash("������ȡҩƷ��Ŀ��ϸ", Me)
    '����ҩƷĿ¼ForBJ
    If blnȫ�� Then
        gstrSQL = "zl_ҩƷĿ¼_Clear"
        mcnYB.Execute gstrSQL, , adCmdStoredProc
    End If
    
    strProcedure = "zl_ҩƷĿ¼_"
    If blnExist Then
        Do While Not objStream.AtEndOfStream
            arrLine = Split(objStream.ReadLine, strSplit)
            lngCols = UBound(arrLine)
            gstrSQL = ""
            strExecute = strProcedure & IIf(arrLine(lngCols) = "0", "INSERT", IIf(arrLine(lngCols) = "1", "DELETE", "UPDATE"))
            If arrLine(lngCols) = "1" Then
                gstrSQL = strExecute & "('" & arrLine(int����) & "')"
            Else
                For lngCol = 0 To lngCols - 1 '���һ���ǲ������루0-����;1-ɾ��;2-�޸ģ�
                    If lngCol <> lngCols - 1 Then
                        gstrSQL = gstrSQL & ",'" & UCase(Replace(arrLine(lngCol), "'", "")) & "'"
                    Else
                        strText = arrLine(lngCol)
                        strText = Mid(strText, 1, 4) & "-" & Mid(strText, 5, 2) & "-" & Mid(strText, 7, 2)
                        gstrSQL = gstrSQL & ",to_Date('" & strText & "','yyyy-MM-dd')"
                    End If
                Next
                gstrSQL = Mid(gstrSQL, 2)
                gstrSQL = strExecute & "(" & gstrSQL & ")"
            End If
            mcnYB.Execute gstrSQL, , adCmdStoredProc
            Me.Caption = "ҽ����Ŀѡ��" & Space(10) & "���ڸ��µ�" & objStream.Line - 1 & "��ҩƷ��Ŀ"
        Loop
        objStream.Close
    End If
    
    '�ٴ���ҩƷ����
    '�����ļ�
    strProcedure = "zl_ҩƷ����_"
    strFileName = strĿ¼ & "\" & IIf(blnȫ��, strҩƷ����_ȫ��, strҩƷ����_����)
    If Not objFileSystem.FileExists(strFileName) Then
        blnExist = False
    Else
        '���ļ����и��²���
        blnExist = True
        Set objStream = objFileSystem.OpenTextFile(strFileName)
    End If
    
    Call zlCommFun.ShowFlash("������ȡҩƷ����", Me)
    '����ҩƷĿ¼ForBJ
    If blnȫ�� Then
        gstrSQL = "zl_ҩƷ����_Clear"
        mcnYB.Execute gstrSQL, , adCmdStoredProc
    End If
    
    If blnExist Then
        Do While Not objStream.AtEndOfStream
            arrLine = Split(objStream.ReadLine, strSplit)
            lngCols = UBound(arrLine)
            gstrSQL = ""
            strExecute = strProcedure & IIf(arrLine(lngCols) = "0", "INSERT", IIf(arrLine(lngCols) = "1", "DELETE", "UPDATE"))
            If arrLine(lngCols) = "1" Then
                gstrSQL = strExecute & "('" & arrLine(int����) & "','" & arrLine(int����) & "')"
            Else
                For lngCol = 0 To lngCols - 1 '���һ���ǲ������루0-����;1-ɾ��;2-�޸ģ�
                    gstrSQL = gstrSQL & ",'" & arrLine(lngCol) & "'"
                Next
                gstrSQL = Mid(gstrSQL, 2)
                gstrSQL = strExecute & "(" & gstrSQL & ")"
            End If
            mcnYB.Execute gstrSQL, , adCmdStoredProc
            Me.Caption = "ҽ����Ŀѡ��" & Space(10) & "���ڸ��µ�" & objStream.Line - 1 & "��ҩƷ������Ŀ"
        Loop
        objStream.Close
    End If
    
    ҩƷĿ¼_���� = True
    Exit Function
errHand:
    mstrErr = "��ǰ��:" & objStream.Line - 1 & "�����:" & Err.Number & "������Ϣ:" & Err.Description
    mErrFile.WriteLine mstrErr
    Resume Next
End Function

Private Function ����Ŀ¼_����(ByVal blnȫ�� As Boolean, ByVal strĿ¼ As String) As Boolean
    Const str����Ŀ¼_ȫ�� As String = "zlml_all.txt"
    Const str������ʩĿ¼_ȫ�� As String = "fwssml_all.txt"
    Const str����Ŀ¼_���� As String = "zlml.txt"
    Const str������ʩĿ¼_���� As String = "fwssml.txt"
    Const int���� As Integer = 1
    Const int���� As Integer = 2
    Const strSplit As String = "|"
    Const strProcedure As String = "zl_����Ŀ¼_"
    Dim arrLine
    Dim lngCol As Long, lngCols As Long
    Dim blnExist As Boolean
    Dim strText As String
    Dim strFileName As String
    Dim strExecute As String
    Dim objStream As TextStream
    Dim objFileSystem As New FileSystemObject
    On Error GoTo errHand
    '�����ļ�
    strFileName = strĿ¼ & "\" & IIf(blnȫ��, str����Ŀ¼_ȫ��, str����Ŀ¼_����)
    If Not objFileSystem.FileExists(strFileName) Then
        blnExist = False
    Else
        '���ļ����и��²���
        blnExist = True
        Set objStream = objFileSystem.OpenTextFile(strFileName)
    End If
    
    Call zlCommFun.ShowFlash("������ȡ������Ŀ��ϸ", Me)
    '��������Ŀ¼ForBJ
    If blnȫ�� Then
        gstrSQL = "zl_����Ŀ¼_Clear"
        mcnYB.Execute gstrSQL, , adCmdStoredProc
    End If
    
    If blnExist Then
        Do While Not objStream.AtEndOfStream
            arrLine = Split(objStream.ReadLine, strSplit)
            lngCols = UBound(arrLine)
            gstrSQL = ""
            strExecute = strProcedure & IIf(arrLine(lngCols) = "0", "INSERT", IIf(arrLine(lngCols) = "1", "DELETE", "UPDATE"))
            If arrLine(lngCols) = "1" Then
                gstrSQL = strExecute & "('" & arrLine(int����) & "')"
            Else
                For lngCol = 0 To lngCols - 1 '���һ���ǲ������루0-����;1-ɾ��;2-�޸ģ�
                    If lngCol <> lngCols - 1 Then
                        gstrSQL = gstrSQL & ",'" & Replace(arrLine(lngCol), "'", "") & "'"
                    Else
                        strText = arrLine(lngCol)
                        strText = Mid(strText, 1, 4) & "-" & Mid(strText, 5, 2) & "-" & Mid(strText, 7, 2)
                        gstrSQL = gstrSQL & ",to_Date('" & strText & "','yyyy-MM-dd')"
                    End If
                Next
                gstrSQL = Mid(gstrSQL, 2)
                gstrSQL = strExecute & "(" & gstrSQL & ")"
            End If
            mcnYB.Execute gstrSQL, , adCmdStoredProc
            Me.Caption = "ҽ����Ŀѡ��" & Space(10) & "���ڸ��µ�" & objStream.Line - 1 & "��������Ŀ"
        Loop
        objStream.Close
    End If
    
    '�����ļ�(�������ʩ�����Ʊ�����һ�ű��У���˷�����ʩ����ʱ�����Բ����)
    strFileName = strĿ¼ & "\" & IIf(blnȫ��, str������ʩĿ¼_ȫ��, str������ʩĿ¼_����)
    If Not objFileSystem.FileExists(strFileName) Then
        blnExist = False
    Else
        '���ļ����и��²���
        blnExist = True
        Set objStream = objFileSystem.OpenTextFile(strFileName)
    End If
    
    Call zlCommFun.ShowFlash("������ȡ������ʩ��Ŀ��ϸ", Me)
    '���·�����ʩĿ¼ForBJ
    If blnExist Then
        Do While Not objStream.AtEndOfStream
            arrLine = Split(objStream.ReadLine, strSplit)
            lngCols = UBound(arrLine)
            gstrSQL = ""
            strExecute = strProcedure & IIf(arrLine(lngCols) = "0", "INSERT", IIf(arrLine(lngCols) = "1", "DELETE", "UPDATE"))
            If arrLine(lngCols) = "1" Then
                gstrSQL = strExecute & "('" & arrLine(int����) & "')"
            Else
                For lngCol = 0 To lngCols - 1 '���һ���ǲ������루0-����;1-ɾ��;2-�޸ģ�
                    If lngCol <> lngCols - 1 Then
                        gstrSQL = gstrSQL & ",'" & Replace(arrLine(lngCol), "'", "") & "'"
                    Else
                        strText = arrLine(lngCol)
                        strText = Mid(strText, 1, 4) & "-" & Mid(strText, 5, 2) & "-" & Mid(strText, 7, 2)
                        gstrSQL = gstrSQL & ",to_Date('" & strText & "','yyyy-MM-dd')"
                    End If
                Next
                gstrSQL = Mid(gstrSQL, 2)
                gstrSQL = strExecute & "(" & gstrSQL & ")"
            End If
            mcnYB.Execute gstrSQL, , adCmdStoredProc
            Me.Caption = "ҽ����Ŀѡ��" & Space(10) & "���ڸ��µ�" & objStream.Line - 1 & "��������ʩ��Ŀ"
        Loop
        objStream.Close
    End If
    ����Ŀ¼_���� = True
    Exit Function
errHand:
    mstrErr = "��ǰ��:" & objStream.Line - 1 & "�����:" & Err.Number & "������Ϣ:" & Err.Description
    mErrFile.WriteLine mstrErr
    Resume Next
End Function

Private Sub Form_Load()
    cmdRequery.Visible = True
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
    
    lvwDetail.Top = tvwClass.Top
    lvwDetail.Left = lblDetail.Left
    lvwDetail.Width = lblDetail.Width
    lvwDetail.Height = tvwClass.Height
End Sub

Private Sub picCmd_Resize()
    cmdCancel.Left = picCmd.ScaleWidth - cmdCancel.Width * 1.4
    cmdOK.Left = cmdCancel.Left - cmdOK.Width * 1.25
    cmdPrint.Top = cmdOK.Top
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName
End Sub

Private Sub lvwDetail_DblClick()
    cmdOK_Click
End Sub

Private Sub picSplit_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        If tvwClass.Width + x < 1000 Or lvwDetail.Width - x < 1000 Then Exit Sub
        picSplit.Left = picSplit.Left + x
        lblClass.Width = lblClass.Width + x
        tvwClass.Width = tvwClass.Width + x
        
        lblDetail.Left = lblDetail.Left + x
        lblDetail.Width = lblDetail.Width - x
        
        lvwDetail.Left = lvwDetail.Left + x
        lvwDetail.Width = lvwDetail.Width - x
    End If
End Sub

Private Sub lvwdetail_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Static blnDesc As Boolean
    Static intIdx As Integer
    
    If intIdx = ColumnHeader.Index Then
        blnDesc = Not blnDesc
    Else
        blnDesc = False
    End If
    lvwDetail.SortKey = ColumnHeader.Index - 1
    If blnDesc Then
        lvwDetail.SortOrder = lvwDescending
    Else
        lvwDetail.SortOrder = lvwAscending
    End If
    lvwDetail.Sorted = True
    intIdx = ColumnHeader.Index
    
    If Not lvwDetail.SelectedItem Is Nothing Then lvwDetail.SelectedItem.EnsureVisible
End Sub

Private Sub tvwClass_NodeClick(ByVal Node As MSComctlLib.Node)
    Call FillList
End Sub

Private Sub txtFind_Change()
'���ܣ������û���������ݲ���ƥ�������
    Dim lst As ListItem, lngIndex As Long, lngSubItems As Long
    Dim strFind As String
    
    strFind = UCase(Trim(txtFind.Text))
    If strFind = "" Then Exit Sub
    If lvwDetail.ListItems.Count = 0 Then Exit Sub

    Set lst = lvwDetail.FindItem(strFind, lvwText, , lvwPartial)
    If Not lst Is Nothing Then
        lst.Selected = True
        lst.EnsureVisible
    Else
        '���ı�������������ƥ��
        lngSubItems = lvwDetail.ColumnHeaders.Count - 1
        For Each lst In lvwDetail.ListItems
            For lngIndex = 1 To lngSubItems
                If lst.SubItems(lngIndex) Like strFind & "*" Then
                    lst.Selected = True
                    lst.EnsureVisible
                    Exit Sub
                End If
            Next
            
        Next
    End If
End Sub

Private Sub txtFind_GotFocus()
    zlControl.TxtSelAll txtFind
End Sub

Private Function ReplaceStr(ByVal StrInput As String) As String
    ReplaceStr = Trim(Replace(StrInput, "'", ""))
    ReplaceStr = Replace(ReplaceStr, """", "")
End Function
