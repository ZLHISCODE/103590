VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSQLEdit 
   AutoRedraw      =   -1  'True
   ClientHeight    =   6240
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7425
   Icon            =   "frmSQLEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   7425
   StartUpPosition =   3  '����ȱʡ
   Begin MSComctlLib.ListView lvwFunc 
      Height          =   4695
      Left            =   2520
      TabIndex        =   9
      Top             =   480
      Visible         =   0   'False
      Width           =   3360
      _ExtentX        =   5927
      _ExtentY        =   8281
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "img16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "������"
         Object.Width           =   5292
      EndProperty
   End
   Begin VB.PictureBox picItem 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1680
      Left            =   0
      ScaleHeight     =   1680
      ScaleWidth      =   7425
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   0
      Width           =   7425
      Begin VB.CommandButton cmdFunc 
         Caption         =   "ѡ��(&S)"
         Height          =   350
         Left            =   4755
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "F3:ѡ��ϵͳ����"
         Top             =   120
         Width           =   1100
      End
      Begin VB.TextBox txt������ 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   1560
         MaxLength       =   50
         TabIndex        =   0
         Top             =   180
         Width           =   3090
      End
      Begin VB.TextBox txt˵�� 
         BackColor       =   &H00FFFFFF&
         Height          =   450
         Left            =   840
         MaxLength       =   250
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   840
         Width           =   5025
      End
      Begin VB.TextBox txt������ 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   840
         MaxLength       =   100
         TabIndex        =   1
         Top             =   495
         Width           =   5025
      End
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   300
         Left            =   840
         ScaleHeight     =   240
         ScaleWidth      =   3795
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   135
         Width           =   3855
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ZL_FUN_"
            ForeColor       =   &H00808080&
            Height          =   180
            Left            =   30
            TabIndex        =   16
            Top             =   15
            Width           =   630
         End
      End
      Begin MSComctlLib.ImageList img16 
         Left            =   6240
         Top             =   255
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSQLEdit.frx":014A
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�����壺���Ҫʹ�ö�̬ʱ����������� zlBeginTime �� zlEndTime ��Ϊ��������"
         ForeColor       =   &H00800000&
         Height          =   180
         Left            =   225
         TabIndex        =   20
         Top             =   1455
         Width           =   6660
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "˵��"
         Height          =   180
         Left            =   405
         TabIndex        =   19
         Top             =   900
         Width           =   360
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������"
         Height          =   180
         Left            =   225
         TabIndex        =   18
         Top             =   195
         Width           =   540
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������"
         Height          =   180
         Left            =   225
         TabIndex        =   17
         Top             =   555
         Width           =   540
      End
   End
   Begin VB.PictureBox picEdit 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      Height          =   3945
      Left            =   15
      ScaleHeight     =   3885
      ScaleWidth      =   7350
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1680
      Width           =   7410
      Begin VB.TextBox txtSQL 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   3885
         Left            =   250
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   3
         Top             =   0
         Width           =   7095
      End
   End
   Begin VB.PictureBox picDown 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   0
      ScaleHeight     =   600
      ScaleWidth      =   7425
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   5640
      Width           =   7425
      Begin VB.PictureBox picCmd 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   420
         ScaleHeight     =   345
         ScaleWidth      =   5400
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   135
         Width           =   5400
         Begin VB.CommandButton cmdCancel 
            Caption         =   "ȡ��(&C)"
            Height          =   350
            Left            =   4305
            TabIndex        =   7
            Top             =   0
            Width           =   1100
         End
         Begin VB.CommandButton cmdOK 
            Caption         =   "ȷ��(&O)"
            Height          =   350
            Left            =   2880
            TabIndex        =   6
            Top             =   0
            Width           =   1100
         End
         Begin VB.CommandButton cmdCompile 
            Caption         =   "����(&M)"
            Height          =   350
            Left            =   15
            TabIndex        =   4
            Top             =   0
            Width           =   1100
         End
         Begin VB.CommandButton cmdPar 
            Caption         =   "����(&P)"
            Height          =   350
            Left            =   1440
            TabIndex        =   5
            Top             =   0
            Width           =   1100
         End
      End
      Begin VB.Frame fra1 
         Height          =   30
         Left            =   -15
         TabIndex        =   12
         Top             =   0
         Width           =   7785
      End
   End
End
Attribute VB_Name = "frmSQLEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'��ڲ�����
Public mblnModi As Boolean
Public mlngSys As Long '��������ϵͳ
Public mstrOwner As String '������������

Private mintNum As Integer '������
Private mobjPars As FuncPars '����������

Private WithEvents objTab As clsTabInput
Attribute objTab.VB_VarHelpID = -1
Private lngSelStart As Long
Private lngSelLen As Long

Private mblnSQLNoDo As Boolean
Private mblnTXTNoDo As Boolean

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdFunc_Click()
    lvwFunc.Visible = True
    lvwFunc.SetFocus
End Sub

Private Sub cmdOK_Click()
    Dim strSQL As String, strPar As String
    Dim i As Integer, arrPar() As String
    Dim StrName As String, j As Integer
    Dim strObject As String, objPar As FuncPar
    
    If Not CheckLen(txt������, txt������.MaxLength, "������", False) Then Exit Sub
    If Not CheckLen(txt������, txt������.MaxLength, "����������", False) Then Exit Sub
    If Not CheckLen(txt˵��, txt˵��.MaxLength, "����˵��") Then Exit Sub
    
    strSQL = Trim(txtSQL.Text)
    StrName = GetFuncName(strSQL)
    If StrName = "" Then
        MsgBox "����������д�������顣", vbInformation, App.Title
        Exit Sub
    End If
    If UCase(StrName) <> UCase("ZL_FUN_" & txt������.Text) Then
        MsgBox "������������д�ĺ������붨��Ĳ�һ�£���������", vbInformation, App.Title
        Exit Sub
    End If
    strSQL = FuncOwnerName(strSQL, StrName, mstrOwner)
    
    If Not CheckFunc(strSQL) Then Exit Sub
    strPar = GetFuncPars(strSQL)
    
    If strPar <> "" And mobjPars.Count = 0 Then
        If MsgBox("���������ж����˲���,��û��Ϊ��Щ��������ȡֵ����,Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbNo Then Exit Sub
    Else
        arrPar = Split(strPar, ";")
        For i = 0 To UBound(arrPar)
            StrName = Split(arrPar(i), ",")(0)
            For j = 1 To mobjPars.Count
                If UCase(mobjPars(j).����) = UCase(StrName) Then
                    Exit For
                End If
            Next
            If j > mobjPars.Count Then
                If MsgBox("���������ж����˲���""" & StrName & """,��û��������ȡֵ����,Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbNo Then Exit Sub
                If MsgBox("Ҫ��������δ����ȡֵ�����Ĳ�����", vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbYes Then Exit For
            End If
        Next
    End If
    
    'ɾ����Ч��������
    For Each objPar In mobjPars
        If InStr(";" & strPar, ";" & objPar.���� & ",") = 0 Then
            mobjPars.Remove "_" & objPar.����
        Else
            If objPar.ȱʡֵ <> "ѡ�������塭" Then
                objPar.���� = ""
                objPar.����SQL = ""
                objPar.�����ֶ� = ""
                objPar.��ϸSQL = ""
                objPar.��ϸ�ֶ� = ""
                If Not objPar.ȱʡֵ Like "*��" Then objPar.ֵ�б� = ""
            End If
        End If
    Next
    
    '������Ȩ��
    For i = 1 To mobjPars.Count
        If mobjPars(i).����SQL <> "" Then
            strObject = strObject & "," & SQLObject(mobjPars(i).����SQL)
        End If
        If mobjPars(i).��ϸSQL <> "" Then
            strObject = strObject & "," & SQLObject(mobjPars(i).��ϸSQL)
        End If
    Next
    strObject = Mid(strObject, 2)
    strObject = CheckObjectPriv(strObject, mstrOwner)
    
    If strObject <> "" Then
        MsgBox "��ǰ�û����������ж����û��Ȩ�޷�����Щ����:" & vbCrLf & vbCrLf & strObject, vbInformation, App.Title
        Exit Sub
    End If
    
    '����
    Screen.MousePointer = 11
    If Not SaveFunc Then
        Screen.MousePointer = 0: Exit Sub
    End If
    Screen.MousePointer = 0
    
    gblnOK = True
    Unload Me
End Sub

Private Sub cmdPar_Click()
    Dim strSQL As String, strPar As String
    Dim blnOK As Boolean, objPars As FuncPars
    
    strSQL = Trim(txtSQL.Text)
    strPar = GetFuncPars(strSQL)
    
    If strPar = "" Then
        Set mobjPars = New FuncPars
        MsgBox "�ں���������û�ж��������", vbInformation, App.Title
        Exit Sub
    End If
    
    blnOK = gblnOK
    
    frmParEdit.mstrPars = strPar
    frmParEdit.mlngSys = mlngSys
    frmParEdit.mstrOwner = mstrOwner
    Set frmParEdit.mobjPars = mobjPars
    frmParEdit.Show 1, Me
    If gblnOK Then
        Call CopyPars(frmParEdit.mobjPars, mobjPars)
        Unload frmParEdit
    End If
    
    gblnOK = blnOK
End Sub

Private Sub cmdCompile_Click()
    Dim strSQL As String, strPar As String
    Dim arrPar() As String, i As Integer
    Dim StrName As String, strText As String
    
    strSQL = Trim(txtSQL.Text)
    If strSQL = "" Then
        MsgBox "�������뺯�����룡", vbInformation, App.Title
        txtSQL.SetFocus: Exit Sub
    End If
    
    StrName = GetFuncName(strSQL)
    If StrName = "" Then
        MsgBox "����������д�������顣", vbInformation, App.Title
        Exit Sub
    End If
    If UCase(StrName) <> UCase("ZL_FUN_" & txt������.Text) Then
        MsgBox "������������д�ĺ������붨��Ĳ�һ�£���������", vbInformation, App.Title
        Exit Sub
    End If
    strSQL = FuncOwnerName(strSQL, StrName, mstrOwner)
    
    If Not mblnModi Then
        grsObject.Filter = "OWNER='" & UCase(mstrOwner) & "' And OBJECT_TYPE='FUNCTION' And OBJECT_NAME='" & UCase(StrName) & "'"
        If Not grsObject.EOF Then
            If MsgBox("ϵͳ���Ѿ�����ͬ���ĺ���,Ҫ��������", vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbNo Then
                If MsgBox("Ҫ��ȡ�Ѵ��ڵĺ��������滻��ǰ��д�Ĵ�����", vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbYes Then
                    strText = GetFunSource(mstrOwner, StrName)
                    strText = GetShowCode(strText, StrName)
                    txtSQL.Text = strText
                End If
                Exit Sub
            End If
        End If
    End If
    
    If Not CheckFunc(strSQL) Then Exit Sub
    
    txtSQL.Tag = txtSQL.Text
    cmdPar.Enabled = True
    cmdOK.Enabled = True
End Sub

Private Function CheckFunc(strSQL As String) As Boolean
'���ܣ���麯���������ȷ��
    Dim arrPar() As String, i As Integer, strPar As String
        
    Screen.MousePointer = 11
    
    On Error Resume Next
        
    'vbCr�����ݿ��л�ת���ɿո�,ֻ����vbLf
    gcnOracle.Execute Replace(strSQL, vbCr, "")
    
    '�����ڲ����󲻻ἤ��Error����
    If gcnOracle.Errors.Count > 0 Then
        Screen.MousePointer = 0
        MsgBox gcnOracle.Errors(0).Description, vbExclamation, App.Title
        Exit Function
    Else
        Screen.MousePointer = 0
        '�������Ϸ���
        strPar = GetFuncPars(strSQL)
        arrPar = Split(strPar, ";")
        
        If InStr(UCase(strPar), "%TYPE") > 0 Then
            MsgBox "�ں����в������������ʹ�ñ��ֶν��ж��壡", vbInformation, App.Title
            lngSelStart = InStr(UCase(txtSQL.Text), "%TYPE")
            lngSelLen = 0: txtSQL.SetFocus: Exit Function
        End If
        
        '��鶯̬ʱ�����
        If InStr(UCase(";" & strPar), ";ZLBEGINTIME,") > 0 And InStr(UCase(";" & strPar), ";ZLENDTIME,") = 0 Then
            MsgBox "��̬ʱ����� zlBeginTime �� zlEndTime �������ʹ�ã�", vbInformation, App.Title
            lngSelStart = InStr(UCase(txtSQL.Text), "ZLBEGINTIME")
            lngSelLen = 0: txtSQL.SetFocus: Exit Function
        End If
        If InStr(UCase(";" & strPar), ";ZLBEGINTIME,") = 0 And InStr(UCase(";" & strPar), ";ZLENDTIME,") > 0 Then
            MsgBox "��̬ʱ����� zlBeginTime �� zlEndTime �������ʹ�ã�", vbInformation, App.Title
            lngSelStart = InStr(UCase(txtSQL.Text), "ZLENDTIME")
            lngSelLen = 0: txtSQL.SetFocus: Exit Function
        End If
        For i = 0 To UBound(arrPar)
            If UCase(Split(arrPar(i), ",")(0)) = "ZLBEGINTIME" And Not UCase(Split(arrPar(i), ",")(1)) Like "*DATE*" Then
                MsgBox "��̬ʱ�����zlBeginTime���붨��Ϊ�������ͣ�", vbInformation, App.Title
                lngSelStart = InStr(UCase(txtSQL.Text), "ZLBEGINTIME") + Len("ZLBEGINTIME")
                lngSelLen = 0: txtSQL.SetFocus: Exit Function
            ElseIf UCase(Split(arrPar(i), ",")(0)) = "ZLENDTIME" And Not UCase(Split(arrPar(i), ",")(1)) Like "*DATE*" Then
                MsgBox "��̬ʱ�����zlEndTime���붨��Ϊ�������ͣ�", vbInformation, App.Title
                lngSelStart = InStr(UCase(txtSQL.Text), "ZLENDTIME") + Len("ZLENDTIME")
                lngSelLen = 0: txtSQL.SetFocus: Exit Function
            End If
        Next
        
        '������������
        For i = 0 To UBound(arrPar)
            If TLen(CStr(Split(arrPar(i), ",")(0))) > 30 Then
                MsgBox "����""" & Split(arrPar(i), ",")(0) & """�ĳ��Ȳ��ܳ���30���ַ���", vbInformation, App.Title
                lngSelStart = InStr(UCase(txtSQL.Text), Split(arrPar(i), ",")(0))
                lngSelLen = 0: txtSQL.SetFocus: Exit Function
            End If
        Next
    End If
    CheckFunc = True
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        If cmdFunc.Enabled And cmdFunc.Visible Then Call cmdFunc_Click
    ElseIf KeyCode = vbKeyEscape Then
        If lvwFunc.Visible Then
            lvwFunc.Visible = False
            txt������.SetFocus
        Else
            cmdCancel_Click
        End If
    End If
End Sub

Private Sub Form_Load()
    Dim strText As String
    
    RestoreWinState Me, App.ProductName
    
    If Not InDesign Then
        glngMinW = 400: glngMinH = 260
        glngMaxW = 1600: glngMaxH = 1200
        glngOldProc = GetWindowLong(hwnd, GWL_WNDPROC)
        Call SetWindowLong(hwnd, GWL_WNDPROC, AddressOf CustomMessage)
    End If
    
    gblnOK = False
    mblnSQLNoDo = False
    mblnTXTNoDo = False
    
    lngSelStart = -1
    
    If mblnModi Then
        Caption = "�޸ĺ���"
        
        cmdFunc.Visible = False
        txt������.Enabled = False
        
        With frmMain.lvw.SelectedItem
            mintNum = Val(Mid(.Key, 2))
            txt������.Text = IIf(UCase(.Text) Like "ZL_FUN_*", Mid(.Text, 8), .Text)
            txt������.Text = .SubItems(2)
            txt˵��.Text = .SubItems(3)
            
            strText = GetFunSource(mstrOwner, .Text)
            Set mobjPars = ReadFuncPars(mlngSys, mintNum)
            strText = GetShowCode(strText, .Text)
        End With
    Else
        Caption = "��������"
        mintNum = NextFuncNo(mlngSys)
        txt������.Text = mintNum
        txt������.Text = ""
        txt˵��.Text = ""
    
        strText = _
            "Create Or Replace Function ZL_FUN_" & txt������.Text & vbCrLf & _
            "Return Number" & vbCrLf & _
            "As" & vbCrLf & _
            "Begin" & vbCrLf & _
            "    Return(0);" & vbCrLf & _
            "End;"
        Set mobjPars = New FuncPars
        
        Call FillUserFunc
    End If
    txtSQL.Tag = strText '��¼ԭ����
    txtSQL.Text = txtSQL.Tag
End Sub

Private Function GetShowCode(strCode As String, strFunc As String) As String
'���ܣ���ʽ����������
    Dim strTmp As String, i As Long
    
    strTmp = strCode
    
    strTmp = Trim(Mid(strTmp, InStr(UCase(strTmp), UCase(strFunc))))
    strTmp = "Create Or Replace Function " & strTmp
    
    '����س�:��vbCrlf���ı����в�����ȷ��ʾ
    strTmp = Replace(strTmp, vbCrLf, Chr(188))
    strTmp = Replace(strTmp, vbCr, Chr(188))
    strTmp = Replace(strTmp, vbLf, Chr(188))
    strTmp = Replace(strTmp, Chr(188), vbCrLf)
    If Left(strTmp, 1) = vbCr Or Left(strTmp, 1) = vbLf Then strTmp = Mid(strTmp, 2)
    If Left(strTmp, 1) = vbCr Or Left(strTmp, 1) = vbLf Then strTmp = Mid(strTmp, 2)
    GetShowCode = strTmp
End Function

Private Sub Form_Resize()
    On Error Resume Next
    
    picEdit.Left = ScaleLeft + 30
    picEdit.Top = ScaleTop + picItem.Height
    picEdit.Width = ScaleWidth - 60
    picEdit.Height = ScaleHeight - picDown.Height - picItem.Height
    
    txtSQL.Left = 0
    txtSQL.Top = 0
    txtSQL.Width = picEdit.ScaleWidth
    txtSQL.Height = picEdit.ScaleHeight
    
    fra1.Width = Width + 60
    
    If ScaleWidth - picCmd.Width - cmdOK.Width / 2 > 300 Then
        picCmd.Left = ScaleWidth - picCmd.Width - cmdOK.Width / 2
    Else
        picCmd.Left = 300
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mobjPars = Nothing
    
    Set objTab = Nothing
    Set gobjTab = Nothing '���������������ͷ�
    
    SaveWinState Me, App.ProductName
    
    If Not InDesign Then Call SetWindowLong(hwnd, GWL_WNDPROC, glngOldProc)
End Sub

Private Sub lvwFunc_DblClick()
    Dim strText As String, StrName As String
    
    If lvwFunc.SelectedItem Is Nothing Then Exit Sub
    
    strText = GetFunSource(mstrOwner, lvwFunc.SelectedItem.Text)
    strText = GetShowCode(strText, lvwFunc.SelectedItem.Text)
    txtSQL.Text = strText
    
    StrName = IIf(UCase(lvwFunc.SelectedItem.Text) Like "ZL_FUN_*", Mid(lvwFunc.SelectedItem.Text, 8), lvwFunc.SelectedItem.Text)
    If UCase(StrName) Like "ZL_FUN*" Then StrName = Mid(StrName, 7)
    If UCase(StrName) Like "ZL_*" Then StrName = Mid(StrName, 4)
    If UCase(StrName) Like "ZL*_*" Then StrName = Mid(StrName, InStr(StrName, "_") + 1)
    
    If txt������.Text = StrName Then
        txt������.Text = StrName
        Call txt������_Change
    Else
        txt������.Text = StrName
    End If
    
    lvwFunc.Visible = False
    txt������.SetFocus
End Sub

Private Sub lvwFunc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call lvwFunc_DblClick
End Sub

Private Sub lvwFunc_LostFocus()
    lvwFunc.Visible = False
End Sub

Private Sub lvwFunc_Validate(Cancel As Boolean)
    lvwFunc.Visible = False
End Sub

Private Sub objTab_sTabKeyDown()
    Dim strNew As String
    Dim lngStart As Long
    
    If Not ActiveControl Is txtSQL Then Exit Sub
    
    If Len(txtSQL.SelText) = 0 Then
        If txtSQL.SelStart > 1 Then
            If Mid(txtSQL.Text, txtSQL.SelStart - 1, 1) = vbTab Then
                txtSQL.SelStart = txtSQL.SelStart - 1
                txtSQL.SelLength = 1
                txtSQL.SelText = ""
            End If
        ElseIf txtSQL.SelStart = 1 And Left(txtSQL.Text, 1) = vbTab Then
            txtSQL.SelStart = 0
            txtSQL.SelLength = 1
            txtSQL.SelText = ""
        End If
    Else '�ɶδ���
        lngStart = txtSQL.SelStart
        
        'ѡ�е�һ�����⴦��(��ʱ��ͷû��vbCr,Ҳû��vbLf)
        strNew = Mid(Replace(vbCr & txtSQL.SelText, vbCr & vbTab, vbCr), 2)
        '������(��ͷ��vbLf)
        strNew = Replace(strNew, vbLf & vbTab, vbLf)
        
        If txtSQL.SelText <> strNew Then
            txtSQL.SelText = strNew
            txtSQL.SelStart = lngStart
            txtSQL.SelLength = Len(strNew)
        End If
    End If
End Sub

Private Sub objTab_TabKeyDown()
    Dim strNew As String
    Dim lngStart As Long
        
    If Not ActiveControl Is txtSQL Then Exit Sub
    
    If Len(txtSQL.SelText) = 0 Then
        txtSQL.SelText = vbTab
    Else '�ɶδ���
        lngStart = txtSQL.SelStart
        
        'ѡ�е�һ�����⴦��(��ʱ��ͷû��vbCr,Ҳû��vbLf)
        strNew = Mid(Replace(vbCrLf & txtSQL.SelText, vbCrLf, vbCrLf & vbTab), 3)
        '�������vbCrLf����
        If Right(strNew, 3) = vbCrLf & vbTab Then strNew = Left(strNew, Len(strNew) - 1)
        
        If txtSQL.SelText <> strNew Then
            txtSQL.SelText = strNew
            txtSQL.SelStart = lngStart
            txtSQL.SelLength = Len(strNew)
        End If
    End If
End Sub

Private Sub txtSQL_Change()
    Dim StrName As String
    
    If mblnSQLNoDo Then Exit Sub
    
    If txtSQL.Text <> txtSQL.Tag Then
        cmdOK.Enabled = False
        cmdPar.Enabled = False
    Else
        cmdOK.Enabled = True
        cmdPar.Enabled = True
    End If
    
    '���ݺ��������Զ�����
    If Not mblnModi And Visible Then
        StrName = GetFuncName(txtSQL.Text)
        If StrName <> "" Then
            mblnTXTNoDo = True
            txt������.Text = IIf(UCase(StrName) Like "ZL_FUN_*", Mid(StrName, 8), StrName)
            mblnTXTNoDo = False
        End If
    End If
End Sub

Private Sub txtSQL_GotFocus()
    If lngSelStart >= 0 Then
        txtSQL.SelStart = lngSelStart
        txtSQL.SelLength = lngSelLen
    End If
    Set objTab = New clsTabInput
End Sub

Private Sub txtSQL_KeyDown(KeyCode As Integer, Shift As Integer)
'Ctrl+A
    If KeyCode = vbKeyA And Shift = 2 Then SelAll txtSQL
End Sub

Private Sub txtSQL_KeyPress(KeyAscii As Integer)
    '
End Sub

Private Sub txtSQL_LostFocus()
    lngSelStart = txtSQL.SelStart
    lngSelLen = txtSQL.SelLength
    Set objTab = Nothing
    Set gobjTab = Nothing '���������������ͷ�
End Sub

Private Sub txt������_Change()
    Dim StrName As String, strTmp As String
    Dim i As Integer
    
    If mblnTXTNoDo Then Exit Sub
    
    '��������ʱ�Զ��޸ĺ�������
    If Visible Then
        strTmp = txtSQL.Text
        StrName = GetFuncName(strTmp)
        If StrName = "" Then Exit Sub
        
        i = InStr(strTmp, StrName)
        Do While i > 0
            strTmp = Left(strTmp, i - 1) & "@#$" & Mid(strTmp, i + Len(StrName))
            i = InStr(strTmp, StrName)
        Loop
        lngSelStart = txtSQL.SelStart
        lngSelLen = 0
        mblnSQLNoDo = True
        txtSQL.Text = Replace(strTmp, "@#$", "ZL_FUN_" & txt������.Text)
        txtSQL.SelStart = lngSelStart
        mblnSQLNoDo = False
    End If
End Sub

Private Sub txt������_GotFocus()
    SelAll txt������
End Sub

Private Sub txt������_KeyPress(KeyAscii As Integer)
    If InStr("'&", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    If KeyAscii = 13 Then SendKeys "{Tab}"
End Sub

Private Sub txt˵��_GotFocus()
    SelAll txt˵��
End Sub

Private Sub txt˵��_KeyPress(KeyAscii As Integer)
    If InStr("'&", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    If KeyAscii = 13 Then SendKeys "{Tab}"
End Sub

Private Sub txt������_GotFocus()
    SelAll txt������
End Sub

Private Sub txt������_KeyPress(KeyAscii As Integer)
    If InStr("'&", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    If KeyAscii = 13 Then SendKeys "{Tab}"
End Sub

Private Function NextFuncNo(lngSys As Long) As Integer
'���ܣ���ȡָ��ϵͳ�е�һ���º�����
'˵�����Զ�ȡȱ��
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = _
        "Select A.������ From zlFunctions A " & _
        "Where A.ϵͳ=" & lngSys & " And Not Exists(Select ������ From zlFunctions B Where B.������=A.������+1 And B.ϵͳ=A.ϵͳ) " & _
        "Order by A.������"
    
    Set rsTmp = zldatabase.OpenSQLRecord(strSQL, "NextFuncNo")
    
    If Not rsTmp.EOF Then
        NextFuncNo = rsTmp!������ + 1
    Else
        NextFuncNo = 1
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function SaveFunc() As Boolean
    Dim i As Integer, arrSQL() As String, objItem As ListItem
    
    ReDim arrSQL(0)
    If mblnModi Then
        arrSQL(UBound(arrSQL)) = "Update zlFunctions Set" & _
            " ������='ZL_FUN_" & txt������.Text & "',������='" & txt������.Text & "',˵��='" & txt˵��.Text & "'" & _
            " Where ϵͳ=" & mlngSys & " And ������=" & mintNum
        
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "Delete From zlFuncPars Where ϵͳ=" & mlngSys & " And ������=" & mintNum
    Else
        arrSQL(UBound(arrSQL)) = "Insert Into zlFunctions(ϵͳ,������,������,������,˵��) Values(" & _
            mlngSys & "," & mintNum & ",'ZL_FUN_" & txt������.Text & "'," & _
            "'" & txt������.Text & "','" & txt˵��.Text & "')"
    End If
    
    If mobjPars.Count > 0 Then
        For i = 1 To mobjPars.Count
            With mobjPars(i)
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "Insert Into zlFuncPars(ϵͳ,������,������,������,������,����,ȱʡֵ," & _
                    "��ʽ,ֵ�б�,����SQL,��ϸSQL,�����ֶ�,��ϸ�ֶ�,����,����) Values(" & mlngSys & "," & mintNum & "," & _
                    i & ",'" & .���� & "','" & .������ & "'," & .���� & ",'" & .ȱʡֵ & "'," & .��ʽ & "," & _
                    "'" & .ֵ�б� & "'," & AdjustStr(.����SQL) & "," & AdjustStr(.��ϸSQL) & ",'" & .�����ֶ� & "','" & .��ϸ�ֶ� & "'," & _
                    "'" & .���� & "','" & .���� & "')"
            End With
        Next
    End If
    
    On Error GoTo errH
    gcnOracle.BeginTrans
    For i = 0 To UBound(arrSQL)
        gcnOracle.Execute arrSQL(i)
    Next
    gcnOracle.CommitTrans
    On Error GoTo 0
    
    With frmMain.lvw
        If mblnModi Then
            Set objItem = .SelectedItem
            objItem.Text = "ZL_FUN_" & txt������.Text
            objItem.SubItems(2) = txt������.Text
            objItem.SubItems(3) = txt˵��.Text
        Else
            Set objItem = .ListItems.Add(, "_" & mintNum, "ZL_FUN_" & txt������.Text, 1, 1)
            objItem.SubItems(1) = mintNum
            objItem.SubItems(2) = txt������.Text
            objItem.SubItems(3) = txt˵��.Text
        End If
        objItem.Selected = True
        .SelectedItem.EnsureVisible
    End With
        
    SaveFunc = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    gcnOracle.RollbackTrans
    Call SaveErrLog
End Function

Private Function FillUserFunc() As Boolean
    Dim i As Integer, strSQL As String
    Dim rsTmp As New ADODB.Recordset
    
    On Error GoTo errH
    
    lvwFunc.ListItems.Clear
    
    strSQL = "Select ��� From zlSystems Where ������=(Select ������ From zlSystems Where ���=" & mlngSys & ")"
    
    strSQL = "Select * From User_Objects" & _
        " Where Object_Type='FUNCTION' And Object_Name Not IN(" & _
        " Select Upper(������) From zlFunctions Where ϵͳ IN(" & strSQL & "))" & _
        " And Object_Name Like 'ZL_FUN_%'" & _
        " Order by Object_Name"
    Set rsTmp = zldatabase.OpenSQLRecord(strSQL, "FillUserFunc")
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            lvwFunc.ListItems.Add , , rsTmp!Object_Name, , 1
            rsTmp.MoveNext
        Next
    End If
    
    FillUserFunc = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
