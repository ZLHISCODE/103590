VERSION 5.00
Begin VB.Form frmLisStationModifyNo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�޸�������"
   ClientHeight    =   2640
   ClientLeft      =   2760
   ClientTop       =   3756
   ClientWidth     =   5028
   Icon            =   "frmLisStationModifyNo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   5028
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.ComboBox cbo 
      Height          =   276
      Index           =   1
      ItemData        =   "frmLisStationModifyNo.frx":000C
      Left            =   1425
      List            =   "frmLisStationModifyNo.frx":000E
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1482
      Width           =   1944
   End
   Begin VB.TextBox txt 
      ForeColor       =   &H000000FF&
      Height          =   300
      Index           =   1
      Left            =   1410
      MaxLength       =   15
      TabIndex        =   2
      Top             =   1380
      Visible         =   0   'False
      Width           =   1950
   End
   Begin VB.CheckBox ChkEmergency 
      Caption         =   "����"
      Height          =   210
      Left            =   2670
      TabIndex        =   14
      Top             =   1065
      Width           =   675
   End
   Begin VB.ComboBox cbo 
      Height          =   276
      Index           =   0
      Left            =   1428
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1920
      Width           =   1944
   End
   Begin VB.Frame Frame1 
      Height          =   4680
      Left            =   3480
      TabIndex        =   10
      Top             =   -1050
      Width           =   30
   End
   Begin VB.TextBox txt 
      ForeColor       =   &H000000FF&
      Height          =   300
      Index           =   0
      Left            =   1425
      MaxLength       =   15
      TabIndex        =   1
      Top             =   1020
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3735
      TabIndex        =   6
      Top             =   600
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   3735
      TabIndex        =   5
      Top             =   132
      Width           =   1100
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "��������(&B)"
      Height          =   180
      Index           =   2
      Left            =   396
      TabIndex        =   13
      Top             =   1530
      Width           =   996
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "  ������(&N)"
      Height          =   180
      Index           =   4
      Left            =   390
      TabIndex        =   15
      Top             =   1440
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "������̬(&S)"
      Height          =   180
      Left            =   384
      TabIndex        =   7
      Top             =   1968
      Width           =   996
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      ForeColor       =   &H00FF0000&
      Height          =   180
      Index           =   1
      Left            =   1425
      TabIndex        =   12
      Top             =   645
      Width           =   540
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      ForeColor       =   &H00FF0000&
      Height          =   180
      Index           =   0
      Left            =   1425
      TabIndex        =   11
      Top             =   270
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "��������(&N)"
      Height          =   180
      Index           =   3
      Left            =   390
      TabIndex        =   0
      Top             =   1080
      Width           =   990
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "�� �� ��:"
      Height          =   180
      Index           =   1
      Left            =   570
      TabIndex        =   9
      Top             =   645
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "����ʱ��:"
      Height          =   180
      Index           =   0
      Left            =   570
      TabIndex        =   8
      Top             =   270
      Width           =   810
   End
End
Attribute VB_Name = "frmLisStationModifyNo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private mblnOK As Boolean
Private mlngKey As Long
Private mstrSQL As String
Private mstrNewNo As String, mstr�걾��̬ As String, mstr�걾���� As String, mstr�걾��� As String, mstr���� As String
Private mlngDevID As Long

Public Function ShowEdit(ByVal frmMain As Object, ByVal lngKey As Long, ByRef strNewNo As String, str�걾��̬ As String, str�걾���� As String, _
                        str�걾��� As String, str���� As String) As Boolean
    Dim rs As New ADODB.Recordset
    Dim str�걾��̬BK As String
    Dim i As Integer
    mlngKey = lngKey
    
    '��ȡ��Ϣ
    mstrSQL = "SELECT ����ʱ��,�걾���,����ID,�걾��̬,Nvl(�걾���,0) As �걾���,�걾����,���� FROM ����걾��¼ WHERE ID=[1]"
    Set rs = zldatabase.OpenSQLRecord(mstrSQL, Me.Caption, mlngKey)
    If rs.EOF Then Exit Function
    
    lbl(0).Caption = Format(rs("����ʱ��").Value, "YYYY-MM-DD")
    lbl(1).Caption = rs("�걾���").Value
    
    txt(0).Text = rs("�걾���").Value
    txt(1).Enabled = (rs("����").Value & "" <> "")
            
    txt(1).Text = rs("����").Value & ""
    str�걾���� = Nvl(rs("�걾����"))
    
'    txt(1).Text = Nvl(rs("�걾����"))
'    str�걾���� = "" & rs("�걾����")
    ChkEmergency.Value = Nvl(rs("�걾���"), 0)
    
'    lblEmerge.Visible = (rs("�걾���") = 1)
    mlngDevID = zlCommFun.Nvl(rs("����ID"), 0)
    str�걾��̬BK = zlCommFun.Nvl(rs("�걾��̬"))
    '��ʼ�걾��̬
    cbo(0).Clear
    cbo(0).AddItem ""
    mstrSQL = "SELECT ����,���� FROM ����걾��̬"
    Set rs = zldatabase.OpenSQLRecord(mstrSQL, Me.Caption)
    If Not rs.EOF Then Call AddComboData(cbo(0), rs, False)
    If cbo(0).ListCount > 0 Then
        cbo(0).ListIndex = 0
        For i = 0 To cbo(0).ListCount - 1
            If str�걾��̬BK = cbo(0).List(i) Then
                cbo(0).ListIndex = i: Exit For
            End If
        Next
    End If
   ' cbo(0).Text = str�걾��̬BK
    
    '-- 2007-07-05 ��ʼ�걾����
    
    cbo(1).Clear
    mstrSQL = "Select ����,���� From ���Ƽ���걾 Order by ����"
    Set rs = zldatabase.OpenSQLRecord(mstrSQL, Me.Caption)
    If Not rs.EOF Then Call AddComboData(cbo(1), rs)
    If cbo(1).ListCount > 0 Then
        cbo(1).ListIndex = 0
        For i = 0 To cbo(1).ListCount - 1
            Debug.Print cbo(1).List(i)
            If str�걾���� = cbo(1).List(i) Then
                cbo(1).ListIndex = i: Exit For
            End If
        Next
    End If
    
    mblnOK = False
    
    Me.Show 1, frmMain
    
    strNewNo = mstrNewNo
    str�걾��̬ = mstr�걾��̬
    str�걾���� = mstr�걾����
    str�걾��� = mstr�걾���
    str���� = mstr����
    ShowEdit = mblnOK
    
End Function

Private Sub cbo_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim mlngLoop As Integer
    
    If KeyAscii = vbKeyReturn Then
        
'        For mlngLoop = 0 To cbo(Index).ListCount - 1
'            If Mid(cbo(Index).List(mlngLoop), 1, InStr(cbo(Index).List(mlngLoop), "-") - 1) = cbo(Index).Text Then
'                cbo(Index).Text = cbo(Index).List(mlngLoop)
'                Exit For
'            End If
'        Next
        
        zlCommFun.PressKey vbKeyTab
    Else
        If Chr(KeyAscii) = "'" Then KeyAscii = 0
    End If
End Sub

Private Sub cbo_Validate(Index As Integer, Cancel As Boolean)
    Select Case Index
        Case 0
            Cancel = Not StrIsValid(cbo(Index).Text, 50)
        Case Else
            Cancel = Not StrIsValid(cbo(Index).Text, 50)
    End Select
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim rs As New ADODB.Recordset
    
    If Trim(txt(0).Text) = "" Then Exit Sub
    
    If Left(Trim(txt(0).Text), 1) = "0" Then
        MsgBox "�걾��Ч������Ϊ�����ͣ�", vbInformation, gstrSysName
        txt(0).SetFocus
        Exit Sub
    End If
    
    If CheckStrType(Trim(txt(0).Text), 99, "0123456789") = False Then
        MsgBox "�걾��Ч������Ϊ�����ͣ�", vbInformation, gstrSysName
        txt(0).SetFocus
        Exit Sub
    End If
    
    
    '����Ƿ���Ч
    If Val(lbl(1).Caption) <> Val(txt(0)) Then
        mstrSQL = "SELECT 1 FROM ����걾��¼ WHERE ����ʱ�� BETWEEN [2] and [3] " & _
            IIf(mlngDevID = 0, " AND ����id IS NULL ", "AND ����id= [4] ") & " AND �걾���= [1] AND Nvl(�걾���,0)=[5]"
        Set rs = zldatabase.OpenSQLRecord(mstrSQL, Me.Caption, Trim(txt(0).Text), _
            CDate(Format(lbl(0).Caption & " 00:00:00", "yyyy-MM-dd hh:mm:ss")), CDate(Format(lbl(0).Caption & " 23:59:59", "yyyy-MM-dd hh:mm:ss")), mlngDevID, _
            IIf(ChkEmergency.Value = 1, 1, 0))
        If rs.BOF = False Then
            MsgBox "�����õı걾���Ѿ����ڣ����������ã�", vbInformation, gstrSysName
            txt(0).SetFocus
            Exit Sub
        End If
    End If
    
    mstrNewNo = TransSampleNO(Trim(txt(0).Text))
    mstr�걾���� = Trim(cbo(1).Text)
    mstr�걾��̬ = IIf(InStr(cbo(0).Text, "-") > 0, zlCommFun.GetNeedName(cbo(0).Text), cbo(0).Text)
    mstr�걾��� = IIf(ChkEmergency.Value = 1, 1, 0)
    mstr���� = txt(1).Text
    mblnOK = True
    Unload Me
End Sub

Private Sub txt_GotFocus(Index As Integer)
    With txt(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii <> 13 Then
        If Index = 0 Then KeyAscii = FilterKeyAscii(KeyAscii, 99, "0123456789")
    Else
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
    Cancel = Not StrIsValid(txt(Index).Text, txt(Index).MaxLength)
End Sub
