VERSION 5.00
Begin VB.Form frmParEdit 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��������"
   ClientHeight    =   1440
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7230
   Icon            =   "frmParEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1440
   ScaleWidth      =   7230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.TextBox txtName 
      BackColor       =   &H8000000F&
      Height          =   300
      Index           =   0
      Left            =   150
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   450
      Width           =   1215
   End
   Begin VB.TextBox txtType 
      BackColor       =   &H8000000F&
      Height          =   300
      Index           =   0
      Left            =   2925
      Locked          =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   450
      Width           =   465
   End
   Begin VB.ComboBox cboGroup 
      BackColor       =   &H00FFFFFF&
      Height          =   300
      Index           =   0
      Left            =   5550
      TabIndex        =   2
      Top             =   450
      Width           =   1575
   End
   Begin VB.ComboBox cboValue 
      BackColor       =   &H00FFFFFF&
      Height          =   300
      Index           =   0
      Left            =   3435
      TabIndex        =   1
      Top             =   450
      Width           =   1905
   End
   Begin VB.TextBox txtAlias 
      BackColor       =   &H00FFFFFF&
      Height          =   300
      Index           =   0
      Left            =   1410
      MaxLength       =   20
      TabIndex        =   0
      Top             =   450
      Width           =   1470
   End
   Begin VB.PictureBox picCmd 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   7230
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   825
      Width           =   7230
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "ȡ��(&C)"
         Height          =   350
         Left            =   5580
         TabIndex        =   4
         Top             =   195
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "ȷ��(&O)"
         Height          =   350
         Left            =   4185
         TabIndex        =   3
         Top             =   195
         Width           =   1100
      End
      Begin VB.Frame Frame1 
         Height          =   75
         Left            =   -105
         TabIndex        =   8
         Top             =   30
         Width           =   7875
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "������"
      Height          =   180
      Left            =   487
      TabIndex        =   13
      Top             =   105
      Width           =   540
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "������"
      Height          =   180
      Left            =   6067
      TabIndex        =   12
      Top             =   90
      Width           =   540
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ȱʡֵ"
      Height          =   180
      Left            =   4117
      TabIndex        =   11
      Top             =   105
      Width           =   540
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����"
      Height          =   180
      Left            =   2977
      TabIndex        =   10
      Top             =   120
      Width           =   360
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����"
      Height          =   180
      Left            =   1875
      TabIndex        =   9
      Top             =   120
      Width           =   360
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C00000&
      BorderWidth     =   2
      X1              =   15
      X2              =   8000
      Y1              =   360
      Y2              =   360
   End
End
Attribute VB_Name = "frmParEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mobjPars As FuncPars '��/��
Public mstrPars As String '�룺������
Public mlngSys As Long '�룺ϵͳ
Public mstrOwner As String '�룺������
Private arrCustom() As CustomPar

Private Sub cboGroup_GotFocus(Index As Integer)
    '���»�ȡ��������
    Dim strGroup As String, arrGroup() As String
    Dim i As Integer, strText As String
    
    '�������
    strGroup = ""
    For i = 0 To cboGroup.UBound
        If InStr(strGroup & ",", "," & cboGroup(i).Text & ",") = 0 And cboGroup(i).Text <> "" Then
            strGroup = strGroup & "," & cboGroup(i).Text
        End If
    Next
    If strGroup <> "" Then
        strGroup = Mid(strGroup, 2)
        arrGroup = Split(strGroup, ",")
        
        'Ϊ�����������
        strText = cboGroup(Index).Text
        cboGroup(Index).Clear
        For i = 0 To UBound(arrGroup)
            cboGroup(Index).AddItem arrGroup(i)
        Next
        
        cboGroup(Index).Text = strText
'        If cboGroup(Index).Text = "" Then
'            If Index > 0 Then
'                If cboGroup(Index - 1).Text <> "" Then cboGroup(Index).Text = cboGroup(Index - 1).Text
'            ElseIf Index < cboGroup.UBound Then
'                If cboGroup(Index + 1).Text <> "" Then cboGroup(Index).Text = cboGroup(Index + 1).Text
'            End If
'        End If
        
        SelAll cboGroup(Index)
    End If
End Sub

Private Sub cboGroup_KeyPress(Index As Integer, KeyAscii As Integer)
    If InStr("~`@#$%^&*()=+][}{'"";/?.>,<\|", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub cboValue_Click(Index As Integer)
    Dim tmpPar As FuncPar, blnDo As Boolean, blnOK As Boolean
    
    blnOK = gblnOK
    
    If cboValue(Index).Text Like "*��" Then
        cboValue(Index).ToolTipText = "�� F2 ����" & cboValue(Index).Text
    Else
        cboValue(Index).ToolTipText = ""
    End If

    If Visible Then
        If cboValue(Index).Text = "�̶�ֵ�б�" Then
            frmFixValue.mbytDataType = txtType(Index).Tag
            frmFixValue.mstrParName = IIf(txtAlias(Index).Text = "", txtName(Index).Text, txtAlias(Index).Text)
            frmFixValue.mbytSelType = arrCustom(Index).��ʽ
            
            '���ܴ�ѡ���������л�����
            If InStr(arrCustom(Index).ֵ�б�, "��") > 0 And InStr(arrCustom(Index).ֵ�б�, ",") > 0 Then
                frmFixValue.mstrValue = arrCustom(Index).ֵ�б�
            Else
                frmFixValue.mstrValue = ""
            End If
            On Error Resume Next
            frmFixValue.Show 1, Me
            On Error GoTo 0
            If gblnOK Then
                arrCustom(Index).ֵ�б� = frmFixValue.mstrValue
                arrCustom(Index).��ʽ = frmFixValue.mbytSelType
                Unload frmFixValue
            ElseIf arrCustom(Index).ֵ�б� = "" Then
                 cboValue(Index).Text = ""
            End If
        ElseIf cboValue(Index).Text = "ѡ�������塭" Then
            frmSelValue.mstrSQLList = arrCustom(Index).��ϸSQL
            frmSelValue.mstrSQLTree = arrCustom(Index).����SQL
            frmSelValue.mstrFLDList = arrCustom(Index).��ϸ�ֶ�
            frmSelValue.mstrFLDTree = arrCustom(Index).�����ֶ�
            frmSelValue.mstrObj = arrCustom(Index).����
            '���ܴӹ̶�ֵ�л�����
            frmSelValue.mstrDef = IIf(InStr(arrCustom(Index).ֵ�б�, "��") > 0, "", arrCustom(Index).ֵ�б�)

            frmSelValue.mbytDataType = txtType(Index).Tag
            frmSelValue.mstrParName = IIf(txtAlias(Index).Text = "", txtName(Index).Text, txtAlias(Index).Text)
            frmSelValue.mlngSys = mlngSys
            frmSelValue.mstrOwner = mstrOwner

            frmSelValue.Show 1, Me
            If gblnOK Then
                arrCustom(Index).��ϸSQL = frmSelValue.mstrSQLList
                arrCustom(Index).����SQL = frmSelValue.mstrSQLTree
                arrCustom(Index).��ϸ�ֶ� = frmSelValue.mstrFLDList
                arrCustom(Index).�����ֶ� = frmSelValue.mstrFLDTree
                arrCustom(Index).���� = frmSelValue.mstrObj
                arrCustom(Index).ֵ�б� = frmSelValue.mstrDef
                Unload frmSelValue
            ElseIf arrCustom(Index).��ϸSQL = "" Then
                cboValue(Index).Text = ""
            End If
        End If
    End If
    
    gblnOK = blnOK
End Sub

Private Sub cboValue_GotFocus(Index As Integer)
    SelAll cboValue(Index)
End Sub

Private Sub cboValue_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 And cboValue(Index).Text Like "*��" Then Call cboValue_Click(Index)
End Sub

Private Sub cboValue_KeyPress(Index As Integer, KeyAscii As Integer)
    If InStr("&~`!@#$^""��" & Chr(22), Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim i As Integer, j As Integer, strPar As String
    Dim tmpPar As FuncPar, curPar As FuncPar

    '�������Ϸ���
    For i = 0 To txtName.UBound
        If txtAlias(i).Text = "" Then
            MsgBox "����""" & txtName(i).Text & """û����������ı�����", vbInformation, App.Title
            txtName(i).SetFocus: Exit Sub
        End If
        If TLen(txtAlias(i).Text) > 40 Then
            MsgBox "����""" & txtName(i).Text & """�������Ȳ��ܳ���40���ַ���", vbInformation, App.Title
            txtName(i).SetFocus: Exit Sub
        End If

        For j = 0 To txtName.UBound
            If j <> i And UCase(txtAlias(i).Text) = UCase(txtAlias(j).Text) Then
                MsgBox "����""" & txtName(i).Text & """��������""" & txtName(j).Text & """�����ظ���", vbInformation, App.Title
                txtAlias(j).SetFocus: Exit Sub
            End If
        Next

        If TLen(cboValue(i).Text) > 255 Then
            MsgBox "����""" & txtName(i).Text & """ȱʡֵ���Ȳ��ܳ���250���ַ���", vbInformation, App.Title
            cboValue(i).SetFocus: Exit Sub
        End If
        If TLen(cboGroup(i).Text) > 30 Then
            MsgBox "����""" & txtName(i).Text & """���������Ȳ��ܳ���30���ַ���", vbInformation, App.Title
            cboGroup(i).SetFocus: Exit Sub
        End If

        If cboValue(i).Text <> "" And Not cboValue(i).Text Like "*��" Then
            If Val(txtType(i).Tag) = 1 Then
                If Not IsNumeric(cboValue(i).Text) Then
                    MsgBox "����""" & txtName(i).Text & """ȱʡֵ����Ӧ��Ϊ�����ͣ�", vbInformation, App.Title
                    cboValue(i).SetFocus: Exit Sub
                End If
            ElseIf Val(txtType(i).Tag) = 2 Then
                If Not IsDate(cboValue(i).Text) And cboValue(i).ListIndex = -1 Then
                    MsgBox "����""" & txtName(i).Text & """ȱʡֵ����Ӧ��Ϊ����/ʱ���ͣ�", vbInformation, App.Title
                    cboValue(i).SetFocus: Exit Sub
                End If
            End If
        End If

        '�Զ������ݼ��
        If cboValue(i).Text = "�̶�ֵ�б�" Then
            If arrCustom(i).ֵ�б� = "" Then
                MsgBox "����""" & txtName(i).Text & """��û�ж����ѡ��Ĺ̶�ֵ�б�", vbInformation, App.Title
                cboValue(i).SetFocus: Exit Sub
            End If
            '�������
            Select Case Val(txtType(i).Tag)
                Case 1 '����
                    For j = 0 To UBound(Split(arrCustom(i).ֵ�б�, "|"))
                        If Not IsNumeric(Split(Split(arrCustom(i).ֵ�б�, "|")(j), ",")(1)) Then
                            MsgBox "����""" & txtName(i).Text & """�Ĺ̶�ֵ�б��д��ڷ������Ͱ�ֵ��", vbInformation, App.Title
                            cboValue(i).SetFocus: Exit Sub
                        End If
                    Next
                Case 2 '����
                    For j = 0 To UBound(Split(arrCustom(i).ֵ�б�, "|"))
                        If Not IsDate(Split(Split(arrCustom(i).ֵ�б�, "|")(j), ",")(1)) Then
                            MsgBox "����""" & txtName(i).Text & """�Ĺ̶�ֵ�б��д��ڷ������Ͱ�ֵ��", vbInformation, App.Title
                            cboValue(i).SetFocus: Exit Sub
                        End If
                    Next
            End Select
        End If

        If cboValue(i).Text = "ѡ�������塭" Then
            If arrCustom(i).��ϸSQL = "" Then
                MsgBox "����""" & txtName(i).Text & """��û�ж���ѡ���������ݣ�", vbInformation, App.Title
                cboValue(i).SetFocus: Exit Sub
            End If
            '�������(֮����Ҫ���ж�һ����Ϊ�û����ܸ�������)
            For j = 0 To UBound(Split(arrCustom(i).��ϸ�ֶ�, "|"))
                If InStr(Split(Split(arrCustom(i).��ϸ�ֶ�, "|")(j), ",")(2), "&B") > 0 Then
                    If Val(txtType(i).Tag) = 1 Then
                        Select Case CLng(Split(Split(arrCustom(i).��ϸ�ֶ�, "|")(j), ",")(1))
                            Case adNumeric, adVarNumeric  '����������
                            Case Else '��������
                                If MsgBox("����""" & txtName(i).Text & """��ѡ�������ֶ� [" & Split(Split(arrCustom(i).��ϸ�ֶ�, "|")(j), ",")(0) & "] ����������,Ҫ������", vbQuestion + vbYesNo, App.Title) = vbNo Then
                                    cboValue(i).SetFocus: Exit Sub
                                End If
                        End Select
                    ElseIf Val(txtType(i).Tag) = 2 Then
                        Select Case CLng(Split(Split(arrCustom(i).��ϸ�ֶ�, "|")(j), ",")(1))
                            Case adDBTimeStamp '����������
                            Case Else '��������
                                If MsgBox("����""" & txtName(i).Text & """��ѡ�������ֶ� [" & Split(Split(arrCustom(i).��ϸ�ֶ�, "|")(j), ",")(0) & "] ����������,Ҫ������", vbQuestion + vbYesNo, App.Title) = vbNo Then
                                    cboValue(i).SetFocus: Exit Sub
                                End If
                        End Select
                    End If
                End If
            Next
        End If
    Next

    '��ͬ�������ٺ�����������(��ֻ�������ڵĲ�����������ͬһ��)��ѡ��������������
    j = 0: strPar = ""
    For i = 0 To UBound(arrCustom)
        If arrCustom(i).��ʽ = 1 And cboGroup(i).Text <> "" Then
            MsgBox "����""" & txtName(i).Text & """���������κβ����飬��Ϊ�ò����ǵ�ѡ����ʽ��", vbInformation, App.Title
            cboGroup(i).SetFocus: Exit Sub
        End If
    Next
    For i = 0 To UBound(arrCustom)
        If strPar <> cboGroup(i).Text Then
            If cboGroup(i).Text = "" Then
                If Not (j = 0 Or j > 1) Then
                    MsgBox "ÿ������������Ҫ������������" & txtName(i).Text & "����", vbInformation, App.Title
                    cboGroup(i).SetFocus: Exit Sub
                End If
                strPar = cboGroup(i).Text
                j = 1
            Else
                If j = 0 Or j > 1 Or strPar = "" Then
                    strPar = cboGroup(i).Text
                    j = 1
                Else
                    MsgBox "ÿ������������Ҫ������������" & txtName(i).Text & "����", vbInformation, App.Title
                    cboGroup(i).SetFocus: Exit Sub
                End If
            End If
        Else
            j = j + 1
        End If
    Next
    
    If Not (j = 0 Or j > 1 Or strPar = "") Then
        MsgBox "ÿ������������Ҫ������������" & txtName(i - 1).Text & "����", vbInformation, App.Title
        cboGroup(i - 1).SetFocus
        Exit Sub
    End If

    'ȷ����������
    Set mobjPars = New FuncPars
    For i = 0 To txtName.UBound
        '�����ǰ�Զ��������ݶ����ڲ�ʹ�ã�ͬ���������ݿ⣬�Ա��Ժ�ʹ�á�
        Set curPar = Nothing
        With arrCustom(i)
            Set curPar = mobjPars.Add(cboGroup(i).Text, CByte(i), txtName(i).Text, txtAlias(i).Text, Val(txtType(i).Tag), _
                cboValue(i).Text, .��ʽ, .ֵ�б�, .����SQL, .��ϸSQL, .�����ֶ�, .��ϸ�ֶ�, .����, "_" & txtName(i).Text)
        End With
    Next

    gblnOK = True
    Hide
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0: SendKeys "{Tab}"
    Else
        If Chr(KeyAscii) = "'" Then KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
    Dim intCount As Integer, i As Integer, j As Integer
    Dim strName As String, strType As String
    
    gblnOK = False
    
    intCount = UBound(Split(mstrPars, ";")) + 1
    ReDim arrCustom(intCount - 1) As CustomPar
    
    For i = 0 To intCount - 1
        If i <> 0 Then
            Load txtName(i): txtName(i).Left = txtName(0).Left: txtName(i).Top = txtName(0).Top + 450 * i: txtName(i).Visible = True
            Load txtAlias(i): txtAlias(i).Left = txtAlias(0).Left: txtAlias(i).Top = txtAlias(0).Top + 450 * i: txtAlias(i).TabIndex = txtAlias(0).TabIndex + 4 * i: txtAlias(i).Visible = True
            Load txtType(i): txtType(i).Left = txtType(0).Left: txtType(i).Top = txtType(0).Top + 450 * i:  txtType(i).Visible = True
            Load cboValue(i): cboValue(i).Left = cboValue(0).Left: cboValue(i).Top = cboValue(0).Top + 450 * i: cboValue(i).TabIndex = cboValue(0).TabIndex + 4 * i: cboValue(i).Visible = True
            Load cboGroup(i): cboGroup(i).Left = cboGroup(0).Left: cboGroup(i).Top = cboGroup(0).Top + 450 * i: cboGroup(i).TabIndex = cboGroup(0).TabIndex + 4 * i: cboGroup(i).Visible = True
        End If

        '�̶�����д��ֵ
        strName = Split(Split(mstrPars, ";")(i), ",")(0)
        txtName(i).Text = strName
        strType = Split(Split(mstrPars, ";")(i), ",")(1)
        If UCase(strType) Like "*NUMBER*" Then
            txtType(i).Text = "��ֵ"
            txtType(i).Tag = 1
        ElseIf UCase(strType) Like "*CHAR*" Then
            txtType(i).Text = "�ַ�"
            txtType(i).Tag = 0
        ElseIf UCase(strType) Like "*DATE*" Then
            txtType(i).Text = "����"
            txtType(i).Tag = 2
        End If
        
        'ȱʡ��ֵ
        txtAlias(i).Text = ""
        cboValue(i).Text = ""
        cboValue(i).AddItem "�̶�ֵ�б�"
        cboValue(i).AddItem "ѡ�������塭"
        cboGroup(i).Text = ""
                        
        'ֵ����
        If UCase(strName) = "ZLBEGINTIME" Or UCase(strName) = "ZLENDTIME" Then
            '��̬ʱ������̶�����
            txtAlias(i).Locked = True
            txtAlias(i).TabStop = False
            txtAlias(i).BackColor = txtName(i).BackColor
            cboValue(i).Locked = True
            cboValue(i).TabStop = False
            cboValue(i).BackColor = txtName(i).BackColor
            
            If UCase(strName) = "ZLBEGINTIME" Then
                txtAlias(i).Text = "��ʼʱ��"
            ElseIf UCase(strName) = "ZLENDTIME" Then
                txtAlias(i).Text = "����ʱ��"
            End If
        Else
            txtAlias(i).Locked = False
            txtAlias(i).TabStop = True
            txtAlias(i).BackColor = &HFFFFFF
            cboValue(i).Locked = False
            cboValue(i).TabStop = True
            cboValue(i).BackColor = &HFFFFFF
        End If
        
        '��������ԭ�е�ֵ
        For j = 1 To mobjPars.Count
            If UCase(mobjPars(j).����) = UCase(strName) Then
                If Not txtAlias(i).Locked Then
                    arrCustom(i).��ʽ = mobjPars(j).��ʽ
                    arrCustom(i).ֵ�б� = mobjPars(j).ֵ�б�
                    arrCustom(i).����SQL = mobjPars(j).����SQL
                    arrCustom(i).��ϸSQL = mobjPars(j).��ϸSQL
                    arrCustom(i).�����ֶ� = mobjPars(j).�����ֶ�
                    arrCustom(i).��ϸ�ֶ� = mobjPars(j).��ϸ�ֶ�
                    arrCustom(i).���� = mobjPars(j).����
                    
                    txtAlias(i).Text = mobjPars(j).������
                    If mobjPars(j).ȱʡֵ Like "*��" Then
                        cboValue(i).ListIndex = GetCboIndex(cboValue(i), mobjPars(j).ȱʡֵ)
                    Else
                        cboValue(i).Text = mobjPars(j).ȱʡֵ
                    End If
                End If
                cboGroup(i).Text = mobjPars(j).����
            End If
        Next
    Next
    
    cmdOK.TabIndex = cboGroup(cboGroup.UBound).TabIndex + 1
    cmdCancel.TabIndex = cmdOK.TabIndex + 1
    Height = txtName(txtName.UBound).Top + 1365
End Sub

Private Sub txtAlias_GotFocus(Index As Integer)
    SelAll txtAlias(Index)
End Sub

Private Sub txtAlias_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If Trim(txtAlias(Index).Text) = "" Then txtAlias(Index).Text = txtName(Index).Text
    End If
End Sub

Private Sub txtName_GotFocus(Index As Integer)
    SelAll txtName(Index)
End Sub

Private Sub txtName_KeyPress(Index As Integer, KeyAscii As Integer)
    If InStr("~`@#$%^&*()=+][}{'"";/?.>,<\|", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub txtType_GotFocus(Index As Integer)
    SelAll txtType(Index)
End Sub
