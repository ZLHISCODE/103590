VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOpenStudyList 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "�򿪼��"
   ClientHeight    =   5910
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   12090
   Icon            =   "frmOpenStudyList.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   12090
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picPanel 
      Align           =   2  'Align Bottom
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   0
      ScaleHeight     =   975
      ScaleWidth      =   12090
      TabIndex        =   0
      Top             =   4935
      Width           =   12090
      Begin VB.CommandButton cmdCancel 
         Caption         =   "ȡ ��(&S)"
         Height          =   375
         Left            =   10800
         TabIndex        =   2
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdSure 
         Caption         =   "ȷ ��(&S)"
         Height          =   375
         Left            =   9120
         TabIndex        =   1
         Top             =   240
         Width           =   1215
      End
   End
   Begin zl9PACSWork.ucFlexGrid ufgStudyList 
      Height          =   3975
      Left            =   600
      TabIndex        =   3
      Top             =   600
      Width           =   10935
      _ExtentX        =   21405
      _ExtentY        =   8705
      HeadCheckValue  =   1
      IsCopyAdoMode   =   0   'False
      IsEjectConfig   =   -1  'True
      Editable        =   0
      ReadOnly        =   -1  'True
      IsShowPopupMenu =   0   'False
      HeadFontCharset =   134
      HeadFontWeight  =   400
      HeadColor       =   0
      DataFontCharset =   134
      DataFontWeight  =   400
      DataColor       =   -2147483640
      GridLineColor   =   14737632
   End
   Begin MSComctlLib.ImageList Imglist 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOpenStudyList.frx":000C
            Key             =   "����"
            Object.Tag             =   "1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOpenStudyList.frx":05A6
            Key             =   "סԺ"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOpenStudyList.frx":0E80
            Key             =   "����"
            Object.Tag             =   "3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOpenStudyList.frx":0FDA
            Key             =   "Ӱ��"
            Object.Tag             =   "4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOpenStudyList.frx":1754
            Key             =   "��ɫͨ��"
            Object.Tag             =   "6"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOpenStudyList.frx":18AE
            Key             =   "·��"
            Object.Tag             =   "7"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOpenStudyList.frx":1E48
            Key             =   "�޷�"
            Object.Tag             =   "7"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOpenStudyList.frx":21E2
            Key             =   "�շ�"
            Object.Tag             =   "8"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmOpenStudyList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mrsData As ADODB.Recordset

Public mlngModule As Long
Public blnOk As Boolean
Public mblncmd�ѽ� As Boolean, mblncmdδ�� As Boolean, mblncmd�޷� As Boolean, mblncmd���� As Boolean

Private mlngTempCharged As Long


Public Sub ShowStudyWindow(ByVal Cols As String, rsData As ADODB.Recordset, owner As Object, imgList As ImageList)
'��ʾ��鴰��
    Dim strFilter As String
    
    Set mrsData = rsData
        
    'ֻ��ʾ������Ϊ����2�����3��������4�ļ������
    strFilter = "������=2 or ������=3 or ������=4"

    Set ufgStudyList.ImageList = imgList
    
    ufgStudyList.ColNames = Replace(Cols, "btn,", "")   '�ڸ��б��У�����Ҫ��ť
    ufgStudyList.ColConvertFormat = ""
    ufgStudyList.DefaultColNames = ""
    ufgStudyList.IsKeepRows = False
        
    Set ufgStudyList.AdoData = mrsData
    ufgStudyList.AdoFilter = strFilter
    
    Call ufgStudyList.BindData
    
    Call ufgStudyList.HidenCol(ufgStudyList.GetColIndex("·��"))
    Call ufgStudyList.HidenCol(ufgStudyList.GetColIndex("����"))
    Call ufgStudyList.HidenCol(ufgStudyList.GetColIndex("Σ��"))
    Call ufgStudyList.HidenCol(ufgStudyList.GetColIndex("��������"))
    Call ufgStudyList.HidenCol(ufgStudyList.GetColIndex("�����ӡ"))
    Call ufgStudyList.HidenCol(ufgStudyList.GetColIndex("���淢��"))
    
    
    
    If mlngModule = G_LNG_PATHOLSYS_NUM Then    '��ȡ������ִ��״̬
        Call ufgStudyList.HidenCol(ufgStudyList.GetColIndex("����ִ��״̬"))
        Call ufgStudyList.HidenCol(ufgStudyList.GetColIndex("����"))
    Else
        Call ufgStudyList.HidenCol(ufgStudyList.GetColIndex("��Ƭ��ӡ"))
        Call ufgStudyList.HidenCol(ufgStudyList.GetColIndex("�������"))
        Call ufgStudyList.HidenCol(ufgStudyList.GetColIndex("Ӱ������"))
    End If
    
    Call ufgStudyList.LocateRow(1)
    
    '��ʾ����б���
    Call Me.Show(1, owner)
End Sub

Private Sub cmdCancel_Click()
    blnOk = False
    Call Me.Hide
End Sub

Private Sub cmdSure_Click()
    If Not ufgStudyList.IsSelectionRow Then
        Call MsgBoxD(Me, "��ѡ����Ҫ����ͼ��ɼ��ļ���¼��", vbOKOnly, gstrSysName)
        Exit Sub
    End If
    
    blnOk = True
    Call Me.Hide
End Sub

Private Sub Form_Load()
    SetWindowPos Me.hWnd, -1, Me.CurrentX, Me.CurrentY, Me.ScaleWidth, Me.ScaleHeight, 3 '�������ö�
    
    Call RestoreWinState(Me, App.ProductName)
    
    blnOk = False
End Sub

Private Sub Form_Resize()
On Error GoTo ErrHandle
    ufgStudyList.Left = 120
    ufgStudyList.Top = 120
    ufgStudyList.Height = Me.ScaleHeight - picPanel.Height - 240
    ufgStudyList.Width = Me.ScaleWidth - 240
    
    cmdCancel.Left = picPanel.Width - cmdCancel.Width - 120
    cmdSure.Left = cmdCancel.Left - cmdSure.Width - 120
    
    Exit Sub
ErrHandle:
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub ufgStudyList_DblClick()
    If Val(ufgStudyList.CurKeyValue) > 0 Then
        blnOk = True
        Me.Hide
    End If
End Sub



Private Sub ufgStudyList_OnFilterRowData(rsData As ADODB.Recordset, rsClone As ADODB.Recordset, blnFilterOut As Boolean)
    '�ж��Ƿ��Ѿ��շ�
    '"����ҽ������.��¼����"--- 1���շѵģ�2�Ǽ��ʵġ�
    
    'ͨ��"����ҽ������.�Ʒ�״̬"ֱ���ж�,ԭ��ֵ��-1-����Ʒ�;0-δ�Ʒ�;1-�ѼƷѣ����ڼ��ʵ�������������ʵ���������ԭ��ֵ���䡣
    '�����շѵ��ķ��ͼ�¼����������״̬��2-�����շѣ�3-ȫ���շ�
    
    'û�ж�Ӧ���õ�ҽ�������������һ����"-1-����Ʒ�"����û�������շѶ��գ�һ����"0-δ�Ʒ�"������Ȼ�������շѶ��գ�������Ϊ���ͺ��ֹ��Ʒѣ�����ҽ������ȥ���ɡ�
    '"1-�ѼƷ�"���Ƿ���ʱ�����˷��õġ��������˷��õ��ݲ���ʾ�շ��ˣ����ɿ����Ǽ��ʻ��۵������շѻ��۵��������շѻ��۵��Ͷ�����״̬��
    '"2-�����շ�"��ʾ�����շѺͲ����˷ѵ����������û�յ��ꡣ
    
    '���շ���ʾ״̬�����շѣ��޷��ã�δ�շѣ�
    'δ�շ�----
    '1����ҽ�����շѵ��ģ���������������δ�շ�
    '   (1)��һ����ҽ���Ͳ�λҽ���� �Ʒ�״̬ in (1,2)��δ�շ� ------����¼����=1 and �Ʒ�״̬ in (1,2)��
    '���շѣ�
    '1����ҽ���Ǽ��˵����շ�-------����¼����=2��
    '2����ҽ�����շѵ��ģ����������������շ�
    '   (1)�ų�δ�շѺ���һ����ҽ���Ͳ�λҽ���� �Ʒ�״̬ =3 ���շ�-----����¼����=1 and �Ʒ�״̬ = 3��
    '�޷���
    '1����ҽ�����շѵ��ģ����������������޷���
    '   (1)������ҽ���Ͳ�λҽ���� �Ʒ�״̬ in (-1,0)���޷��� ------����¼����=1 and �Ʒ�״̬ in (-1,0)��
    
    
    ' intCharged  '0--δ�շѣ�1--���շѣ�2--�޷���
    
    If Nvl(rsData!���ID) <> "" Then
        '���id��Ϊ��ʱ��˵���鲿λҽ��������Ҫ��ʾ���б���
        blnFilterOut = True
        Exit Sub
    End If

    mlngTempCharged = 2 '�޷���
    
    If Nvl(rsData!��¼����, 2) = 2 Then
        'סԺ�ǼǵĲ��ˣ����û�мƷѣ����Ϊ�޷���
        If Nvl(rsData!�Ʒ�״̬, -1) = 0 Then
            mlngTempCharged = 2
        Else
            mlngTempCharged = 1  '���շ�
        End If
    Else
        If Nvl(rsData!�Ʒ�״̬, -1) = 1 Or Nvl(rsData!�Ʒ�״̬, -1) = 2 Then
            mlngTempCharged = 0      'δ�շ�
        Else        '��ҽ���ļƷ�״̬�� -1,0,3  ��3--���շѣ�-1��0--�޷��ã�
            '��ѯ��ҽ��δ�Ʒѻ����Ѿ��շ��ˣ���Ҫ�鲿λҽ�����շ����������ҽ�����Ѿ��շѣ��������շ�
            
            '��������������շѵģ��ȼ�¼�����շ�
            If Nvl(rsData!�Ʒ�״̬, -1) = 3 Then
                mlngTempCharged = 1      '���շ�
            End If
            
            rsClone.Filter = "���ID = " & Nvl(rsData!ҽ��ID)
            Do While rsClone.EOF = False
                If Nvl(rsClone!�Ʒ�״̬, -1) = 1 Or Nvl(rsClone!�Ʒ�״̬, -1) = 2 Then
                    mlngTempCharged = 0      'δ�շ�

                    Exit Do
                ElseIf Nvl(rsClone!�Ʒ�״̬, -1) = 3 Then
                    mlngTempCharged = 1      '���շ�
                End If

                rsClone.MoveNext
            Loop
            
'            '�Ʒ�״̬��-1-����Ʒ�(ͨ����ִ�к�Ժ��ִ�еĶ�����Ʒ�);0-δ�Ʒ�;1-�ѼƷѣ����շѵ��ݶ�����״̬:2-�����շѣ�3-ȫ���շ�
'            rsClone.Filter = "���ID = " & Nvl(rsData!ҽ��ID) & " and �Ʒ�״̬=1 and �Ʒ�״̬=2"
'            If rsClone.RecordCount > 0 Then
'                mlngTempCharged = 0 'δ�շ�
'            Else
'                rsClone.Filter = "���ID = " & Nvl(rsData!ҽ��ID) & " and �Ʒ�״̬=3"
'                If rsClone.RecordCount > 0 Then mlngTempCharged = 1 '���շ�
'            End If
            
        End If
    End If

    If mlngModule = G_LNG_PATHOLSYS_NUM Then
        If Nvl(rsData!����) > 0 Then mlngTempCharged = 4 '��Ҫ���ѣ��貹�ѵļ��Ҳ��δ�շѵļ��
    End If
    
    If Nvl(rsData!���ID) = "" And ((mblncmd�ѽ� = True And mlngTempCharged = 1) Or (mblncmdδ�� = True And (mlngTempCharged = 0 Or mlngTempCharged = 4)) _
        Or (mblncmd�޷� = True And mlngTempCharged = 2) Or (mblncmd���� = True And mlngTempCharged = 4) _
        Or (mblncmd�ѽ� = False And mblncmdδ�� = False And mblncmd���� = False And mblncmd�޷� = False)) Then
        blnFilterOut = False
        
        Call RowDataConvert(rsData)
    Else
        blnFilterOut = True
    End If
End Sub


Private Sub RowDataConvert(rsData As ADODB.Recordset)
    Dim rsBaby As ADODB.Recordset
    Dim intTxtLen As Long
    
    '���������Ҫ��ʾ������Ҫת�������еĲ���ֵ
    rsData!���뵥 = IIf(Nvl(rsData!���뵥) = "", "��", "��ɨ��")
    rsData!������ = IIf(Val(Nvl(rsData!ִ��״̬)) = 2, "�Ѿܾ�", Decode(Val(Nvl(rsData!���״̬, 0)), -1, "�Ѳ���", 0, "�ѵǼ�", 1, "�ѵǼ�", _
                                                                                2, IIf(Nvl(rsData!�������) <> "", "������", _
                                                                                        IIf(Nvl(rsData!������) = "", "�ѱ���", "������")), _
                                                                                3, IIf(Nvl(rsData!�������) <> "", "������", _
                                                                                        IIf(Nvl(rsData!������) = "", "�Ѽ��", "������")), _
                                                                                4, IIf(Nvl(rsData!�������) <> "", "������", _
                                                                                        IIf(Nvl(rsData!������) <> "", "�����", "�ѱ���")), _
                                                                                5, "�����", "�����"))
                                                                                
    If Nvl(rsData!Ӥ��) <> 0 Then
        gstrSQL = "Select Nvl(A.Ӥ������, B.���� || '֮��' || Trim(To_Char(A.���, '9'))) As Ӥ������, Ӥ���Ա�, ����ʱ��" & vbNewLine & _
                    "From ������������¼ A, ������Ϣ B" & vbNewLine & _
                    "Where A.����id = [1] And A.��ҳid = [2] And A.����id = B.����id And A.��� = [3]"
        
        Set rsBaby = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡӤ����Ϣ", CLng(rsData!����ID), CLng(Nvl(rsData!��ҳID, 0)), CLng(rsData!Ӥ��))
        
        If Not rsBaby.EOF Then
            rsData!���� = rsBaby!Ӥ������
            rsData!�Ա� = Nvl(rsBaby!Ӥ���Ա�)
            rsData!���� = Nvl(rsBaby!����ʱ��)
        End If
    End If
    
    
    If InStr(Nvl(rsData!ҽ������), ":") > 0 Then '�µ�ģʽ������ҽ����������Ϣ�� ����,ִ�б��:��λ(����,����),��λ---
        rsData!��λ���� = Split(Nvl(rsData!ҽ������), ":")(1)
        rsData!ҽ������ = Split(Nvl(rsData!ҽ������), ":")(0)
    End If
    
    
    If Val(Nvl(rsData!����)) <> 0 Then
        rsData!���� = " "
    Else
        rsData!���� = ""
    End If
    
    If mlngTempCharged = 0 Then  'δ�շ�
        rsData!�շ� = ""
    ElseIf mlngTempCharged = 1 Then   '���շ�
        rsData!�շ� = " "
    ElseIf mlngTempCharged = 2 Then    '�޷���
        rsData!�շ� = "  "
    Else
        rsData!�շ� = "   "
    End If
    
    If rsData!��Դ = 1 Then
        rsData!��Դ = "��"
    ElseIf rsData!��Դ = 2 Then
        rsData!��Դ = "ס"
    ElseIf rsData!��Դ = 3 Then
        rsData!��Դ = "��"
    ElseIf rsData!��Դ = 4 Then
        rsData!��Դ = "��"
    End If
End Sub


Private Sub ufgStudyList_OnRefreshRowData(rsBind As ADODB.Recordset, ByVal lngRow As Long)
On Error GoTo ErrHandle
    Dim strTag As String
    Dim strTemp As String
    Dim i As Long
    
    For i = 0 To ufgStudyList.DataGrid.Cols - 1
        Select Case ufgStudyList.DataGrid.TextMatrix(0, i)
                
                
            Case "����"
                If ufgStudyList.Text(lngRow, "����") = " " Then
                    Set ufgStudyList.DataGrid.Cell(flexcpPicture, lngRow, i) = imgList.ListImages("����").Picture
                End If
        
            Case "��Դ"
                strTag = Decode(ufgStudyList.Text(lngRow, "��Դ"), "��", 1, "ס", 2, "��", 3, 4)
                ufgStudyList.DataGrid.Cell(flexcpData, lngRow, i) = strTag
                
                If ufgStudyList.Text(lngRow, "��Դ") = "ס" Then
                    Set ufgStudyList.DataGrid.Cell(flexcpPicture, lngRow, i) = imgList.ListImages("סԺ").Picture
                End If
                
            Case "�շ�" 'TODO:������Ҫ���ǲ��ɷ��õ����
                If ufgStudyList.Text(lngRow, "�շ�") = "" Then  'δ�շ�
                    Set ufgStudyList.DataGrid.Cell(flexcpPicture, lngRow, i) = imgList.ListImages("�޷�").Picture
                ElseIf ufgStudyList.Text(lngRow, "�շ�") = " " Then   '���շ�
                    Set ufgStudyList.DataGrid.Cell(flexcpPicture, lngRow, i) = imgList.ListImages("�շ�").Picture
                ElseIf ufgStudyList.Text(lngRow, "�շ�") = "   " Then   '����
                    Set ufgStudyList.DataGrid.Cell(flexcpPicture, lngRow, i) = imgList.ListImages("����").Picture
                Else '�޷���("  ")
                    '�޷��ò���ʾͼ��
                End If

                
            Case "����" '���Ϊ��ɫͨ��������Ҫ��������ǰ���ͼ��
                If Val(ufgStudyList.Text(lngRow, "��ɫͨ��")) <> 0 Then
                    Set ufgStudyList.DataGrid.Cell(flexcpPicture, lngRow, i) = imgList.ListImages("��ɫͨ��").Picture
                End If
                
            Case GetStudyNumberDisplayName  '���Ż��߲����
                If ufgStudyList.Text(lngRow, "���UID") <> "" Then
                    '����ϵͳ�У�����б��еļ�����ʾΪ�����
                    Set ufgStudyList.DataGrid.Cell(flexcpPicture, lngRow, i) = imgList.ListImages(IIf(mlngModule = G_LNG_PATHOLSYS_NUM, "����", "Ӱ��")).Picture
                End If
                
            Case "������"
                '���ݼ����̣����ò�ͬ����ɫ
                If ufgStudyList.Text(lngRow, "������") = "�Ѿܾ�" Then ufgStudyList.DataGrid.Cell(flexcpBackColor, lngRow, 1, lngRow, ufgStudyList.DataGrid.Cols - 1) = gdblColor�Ѿܾ�
                If ufgStudyList.Text(lngRow, "������") = "�����" Then ufgStudyList.DataGrid.Cell(flexcpBackColor, lngRow, 1, lngRow, ufgStudyList.DataGrid.Cols - 1) = gdblColor�����
                If ufgStudyList.Text(lngRow, "������") = "�ѱ���" Then ufgStudyList.DataGrid.Cell(flexcpBackColor, lngRow, 1, lngRow, ufgStudyList.DataGrid.Cols - 1) = gdblColor�ѱ���
                If ufgStudyList.Text(lngRow, "������") = "�ѵǼ�" Then ufgStudyList.DataGrid.Cell(flexcpBackColor, lngRow, 1, lngRow, ufgStudyList.DataGrid.Cols - 1) = gdblColor�ѵǼ�
                If ufgStudyList.Text(lngRow, "������") = "�Ѽ��" Then ufgStudyList.DataGrid.Cell(flexcpBackColor, lngRow, 1, lngRow, ufgStudyList.DataGrid.Cols - 1) = gdblColor�Ѽ��
                If ufgStudyList.Text(lngRow, "������") = "�����" Then ufgStudyList.DataGrid.Cell(flexcpBackColor, lngRow, 1, lngRow, ufgStudyList.DataGrid.Cols - 1) = gdblColor�����
                If ufgStudyList.Text(lngRow, "������") = "������" Then ufgStudyList.DataGrid.Cell(flexcpBackColor, lngRow, 1, lngRow, ufgStudyList.DataGrid.Cols - 1) = gdblColor������
                If ufgStudyList.Text(lngRow, "������") = "������" Then ufgStudyList.DataGrid.Cell(flexcpBackColor, lngRow, 1, lngRow, ufgStudyList.DataGrid.Cols - 1) = gdblColor������
                If ufgStudyList.Text(lngRow, "������") = "�����" Then ufgStudyList.DataGrid.Cell(flexcpBackColor, lngRow, 1, lngRow, ufgStudyList.DataGrid.Cols - 1) = gdblColor�����
                If ufgStudyList.Text(lngRow, "������") = "�ѱ���" Then ufgStudyList.DataGrid.Cell(flexcpBackColor, lngRow, 1, lngRow, ufgStudyList.DataGrid.Cols - 1) = gdblColor�ѱ���
                If ufgStudyList.Text(lngRow, "������") = "�Ѳ���" Then ufgStudyList.DataGrid.Cell(flexcpBackColor, lngRow, 1, lngRow, ufgStudyList.DataGrid.Cols - 1) = gdblColor�Ѳ���
                                
        End Select
        
    Next i
    
ErrHandle:
    Exit Sub
End Sub


Private Function GetStudyNumberDisplayName() As String
'��ȡ��������ʾ����
    GetStudyNumberDisplayName = IIf(mlngModule = G_LNG_PATHOLSYS_NUM, "�����", "����")
End Function
