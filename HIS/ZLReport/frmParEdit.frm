VERSION 5.00
Begin VB.Form frmParEdit 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��������"
   ClientHeight    =   1440
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8640
   Icon            =   "frmParEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1440
   ScaleWidth      =   8640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CheckBox chkLock 
      Height          =   225
      Index           =   0
      Left            =   7920
      TabIndex        =   5
      Top             =   480
      Width           =   210
   End
   Begin VB.ComboBox cboGroup 
      BackColor       =   &H00FFFFFF&
      Height          =   300
      Index           =   0
      Left            =   5550
      TabIndex        =   4
      Top             =   450
      Width           =   1905
   End
   Begin VB.ComboBox cboValue 
      BackColor       =   &H00FFFFFF&
      Height          =   300
      Index           =   0
      Left            =   3540
      TabIndex        =   3
      Top             =   450
      Width           =   1905
   End
   Begin VB.ComboBox cboType 
      BackColor       =   &H00FFFFFF&
      Height          =   300
      Index           =   0
      Left            =   2385
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   450
      Width           =   1005
   End
   Begin VB.TextBox txtName 
      BackColor       =   &H00FFFFFF&
      Height          =   300
      Index           =   0
      Left            =   825
      MaxLength       =   20
      TabIndex        =   1
      Top             =   450
      Width           =   1395
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
      ScaleWidth      =   8640
      TabIndex        =   8
      Top             =   825
      Width           =   8640
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "ȡ��(&C)"
         Height          =   350
         Left            =   7320
         TabIndex        =   7
         Top             =   195
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "ȷ��(&O)"
         Height          =   350
         Left            =   5925
         TabIndex        =   6
         Top             =   195
         Width           =   1100
      End
      Begin VB.Frame Frame1 
         Height          =   75
         Left            =   -105
         TabIndex        =   9
         Top             =   30
         Width           =   10000
      End
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ִ��ʱ����"
      Height          =   180
      Left            =   7560
      TabIndex        =   15
      Top             =   90
      Width           =   900
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "������"
      Height          =   180
      Left            =   6000
      TabIndex        =   14
      Top             =   90
      Width           =   540
   End
   Begin VB.Label lblNO 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   180
      Index           =   0
      Left            =   420
      TabIndex        =   0
      Top             =   510
      Width           =   90
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ȱʡֵ"
      Height          =   180
      Left            =   4222
      TabIndex        =   13
      Top             =   105
      Width           =   540
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����"
      Height          =   180
      Left            =   2707
      TabIndex        =   12
      Top             =   120
      Width           =   360
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����"
      Height          =   180
      Left            =   1342
      TabIndex        =   11
      Top             =   120
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "���"
      Height          =   180
      Left            =   360
      TabIndex        =   10
      Top             =   120
      Width           =   360
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C00000&
      BorderWidth     =   2
      X1              =   15
      X2              =   9000
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
Private mobjPars As RPTPars '��/��
Private arrCustom() As CustomPar
Private intPreIdx As Integer

Private mstrSQL As String
Private mlngSys As Long
Private mobjData As RPTData
Private mobjDatas As RPTDatas
Private mblnOK As Boolean
Private mlngReportID As Long

Public Function ShowMe(objParent As Object, ByVal lngSys As Long, objData As RPTData, objDatas As RPTDatas, _
    ByRef objPars As RPTPars, ByRef strSQL As String, ByVal lngReportID As Long) As Boolean
    
    Set mobjPars = objPars
    mlngSys = lngSys
    mstrSQL = strSQL
    Set mobjData = objData
    Set mobjDatas = objDatas
    mlngReportID = lngReportID
    
    Me.Show 1, objParent
    Set objData = mobjData
    Set objDatas = mobjDatas
    Set objPars = mobjPars
    strSQL = mstrSQL
    ShowMe = mblnOK
End Function

Private Sub cboGroup_Change(Index As Integer)
    Dim IntSelStart As Integer
    IntSelStart = cboGroup(Index).SelStart
    cboGroup(Index).Text = UCase(cboGroup(Index).Text)
    arrCustom(Index).���� = cboGroup(Index).Text
    cboGroup(Index).SelStart = IntSelStart
End Sub

Private Sub cboGroup_Click(Index As Integer)
    arrCustom(Index).���� = cboGroup(Index).Text
End Sub

Private Sub cboGroup_GotFocus(Index As Integer)
    '���»�ȡ��������
    Dim strGroup As String, arrGroup
    Dim IntGroup As Integer, IntAdd As Integer
    '�������
    strGroup = ""
    For IntGroup = 0 To UBound(arrCustom)
        If InStr(1, strGroup, "^" & arrCustom(IntGroup).���� & ",") = 0 And arrCustom(IntGroup).���� <> "" Then strGroup = strGroup & "^" & arrCustom(IntGroup).���� & ","
    Next
    If strGroup <> "" Then
        strGroup = Mid(strGroup, 1, Len(strGroup) - 1)
        arrGroup = Split(strGroup, ",")
        
        'Ϊ�����������
        cboGroup(Index).Clear
        For IntAdd = 0 To UBound(arrGroup)
            cboGroup(Index).AddItem Mid(arrGroup(IntAdd), 2)
        Next
        If cboGroup(Index).ListCount <> 0 Then cboGroup(Index).Text = arrCustom(Index).����
        
        cboGroup(Index).SelStart = 0
        cboGroup(Index).SelLength = 1000
    End If
End Sub

Private Sub cboGroup_KeyPress(Index As Integer, KeyAscii As Integer)
    If InStr("~`@#$%^&*()=+][}{'"";/?.>,<\|", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub cboType_Click(Index As Integer)
    'ֻ�������Ͳ����ſ���ʹ�ö�ѡѡ����
    If cboType(Index).ListIndex <> 3 And cboValue(Index).Text = "ѡ�������塭" Then
        arrCustom(Index).��ʽ = 0
    End If
End Sub

Private Sub cboValue_Click(Index As Integer)
    Dim tmpPar As RPTPar, tmpData As RPTData
    Dim blnDo As Boolean
    
    If cboValue(Index).Text Like "*��" Then
        cboValue(Index).ToolTipText = "�� F2 ����" & cboValue(Index).Text
    Else
        cboValue(Index).ToolTipText = ""
    End If
    
    If Visible Then
        If cboValue(Index).Text = "�̶�ֵ�б�" Then
            '����������Դ����ͬ������ֵ�б��ƹ���
            If arrCustom(Index).ֵ�б� = "" Then
                For Each tmpData In mobjDatas
                    If tmpData.���� <> mobjData.���� Then
                        For Each tmpPar In tmpData.Pars
                            If tmpPar.���� = txtName(Index).Text And tmpPar.ȱʡֵ = cboValue(Index).Text Then
                                arrCustom(Index).ֵ�б� = tmpPar.ֵ�б�
                                arrCustom(Index).��ʽ = tmpPar.��ʽ
                                blnDo = True: Exit For
                            End If
                        Next
                    End If
                    If blnDo Then Exit For
                Next
            End If
            
            frmFixValue.bytType = cboType(Index).ListIndex
            frmFixValue.strName = txtName(Index).Text
            frmFixValue.IntSelType = arrCustom(Index).��ʽ
            '���ܴ�ѡ���������л�����
            '���õķָ���
            If InStr(arrCustom(Index).ֵ�б�, "��") > 0 And InStr(arrCustom(Index).ֵ�б�, ",") > 0 Then
                frmFixValue.strValues = arrCustom(Index).ֵ�б�
            End If
            frmFixValue.Show 1, Me
            If gblnOK Then
                arrCustom(Index).ֵ�б� = frmFixValue.strValues
                arrCustom(Index).��ʽ = frmFixValue.IntSelType
                Unload frmFixValue
                
                'ͬʱ������������Դ����ͬ������ֵ�б�
                For Each tmpData In mobjDatas
                    If tmpData.���� <> mobjData.���� Then
                        For Each tmpPar In tmpData.Pars
                            If tmpPar.���� = txtName(Index).Text And tmpPar.ȱʡֵ = cboValue(Index).Text Then
                                tmpPar.ֵ�б� = arrCustom(Index).ֵ�б�
                                tmpPar.��ʽ = arrCustom(Index).��ʽ
                            End If
                        Next
                    End If
                Next
            End If
        ElseIf cboValue(Index).Text = "ѡ�������塭" Then
            '����������Դ����ͬ������ֵ���ƹ���
            If arrCustom(Index).��ϸSQL = "" Then
                For Each tmpData In mobjDatas
                    If tmpData.���� <> mobjData.���� Then
                        For Each tmpPar In tmpData.Pars
                            If tmpPar.���� = txtName(Index).Text And tmpPar.ȱʡֵ = cboValue(Index).Text Then
                                arrCustom(Index).��ϸSQL = tmpPar.��ϸSQL
                                arrCustom(Index).��ϸ�ֶ� = tmpPar.��ϸ�ֶ�
                                arrCustom(Index).����SQL = tmpPar.����SQL
                                arrCustom(Index).�����ֶ� = tmpPar.�����ֶ�
                                arrCustom(Index).���� = tmpPar.����
                                arrCustom(Index).��ʽ = tmpPar.��ʽ
                                arrCustom(Index).ֵ�б� = tmpPar.ֵ�б�
                                blnDo = True: Exit For
                            End If
                        Next
                    End If
                    If blnDo Then Exit For
                Next
            End If
            'ֻ�������Ͳ����ſ���ʹ�ö�ѡѡ����
            If cboType(Index).ListIndex <> 3 Then arrCustom(Index).��ʽ = 0
            
            frmSelValue.mstrSQLList = arrCustom(Index).��ϸSQL
            frmSelValue.mstrSQLTree = arrCustom(Index).����SQL
            frmSelValue.mstrFLDList = arrCustom(Index).��ϸ�ֶ�
            frmSelValue.mstrFLDTree = arrCustom(Index).�����ֶ�
            frmSelValue.mstrObj = arrCustom(Index).����
            '���ܴӹ̶�ֵ�л�����
            frmSelValue.mstrDef = IIF(InStr(arrCustom(Index).ֵ�б�, "��") > 0, "", arrCustom(Index).ֵ�б�)
            
            frmSelValue.mbytType = cboType(Index).ListIndex
            frmSelValue.mstrName = txtName(Index).Text
            frmSelValue.mblnMulti = arrCustom(Index).��ʽ = 1
            frmSelValue.mlngSys = mlngSys
            Set frmSelValue.mobjDatas = mobjDatas
            Set frmSelValue.mobjData = mobjData
            
            frmSelValue.Show 1, Me
            If gblnOK Then
                arrCustom(Index).��ϸSQL = frmSelValue.mstrSQLList
                arrCustom(Index).����SQL = frmSelValue.mstrSQLTree
                arrCustom(Index).��ϸ�ֶ� = frmSelValue.mstrFLDList
                arrCustom(Index).�����ֶ� = frmSelValue.mstrFLDTree
                arrCustom(Index).���� = frmSelValue.mstrObj
                arrCustom(Index).ֵ�б� = frmSelValue.mstrDef
                arrCustom(Index).��ʽ = IIF(frmSelValue.mblnMulti, 1, 0)
                Unload frmSelValue
                
                'ͬʱ������������Դ����ͬ������ֵ
                For Each tmpData In mobjDatas
                    If tmpData.���� <> mobjData.���� Then
                        For Each tmpPar In tmpData.Pars
                            If tmpPar.���� = txtName(Index).Text And tmpPar.ȱʡֵ = cboValue(Index).Text Then
                                tmpPar.��ϸSQL = arrCustom(Index).��ϸSQL
                                tmpPar.����SQL = arrCustom(Index).����SQL
                                tmpPar.��ϸ�ֶ� = arrCustom(Index).��ϸ�ֶ�
                                tmpPar.�����ֶ� = arrCustom(Index).�����ֶ�
                                tmpPar.���� = arrCustom(Index).����
                                tmpPar.ֵ�б� = arrCustom(Index).ֵ�б�
                                tmpPar.��ʽ = arrCustom(Index).��ʽ
                            End If
                        Next
                    End If
                Next
            End If
        End If
    End If
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

Private Sub chkLock_GotFocus(Index As Integer)
    chkLock(Index).Tag = "" & chkLock(Index).BackColor
    chkLock(Index).BackColor = &HC0C0C0
End Sub

Private Sub chkLock_LostFocus(Index As Integer)
    chkLock(Index).BackColor = Val(chkLock(Index).Tag)
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim i As Integer, j As Integer
    Dim tmpPar As RPTPar, tmpData As RPTData
    Dim curPar As RPTPar
    Dim strAddDelBefor As String, strAddDelAfter As String, strInfo As String '�ж��Ƿ�ɾ���������˲���
    Dim strSQL As String, rsTmp As Recordset
    
    If Not CheckFormInput(Me, True) Then Exit Sub
    
    '�������Ϸ���
    For i = 0 To lblNO.UBound
        If txtName(i).Text = "" Then
            MsgBox "�� " & i & " �ĸ�����û������������ƣ�", vbInformation, App.Title
            txtName(i).SetFocus: Exit Sub
        End If
        If TLen(txtName(i).Text) > 20 Then
            MsgBox "�� " & i & " �ĸ��������Ƴ��Ȳ��ܳ���20���ַ���", vbInformation, App.Title
            txtName(i).SetFocus: Exit Sub
        End If
        
        For j = 0 To lblNO.UBound
            If j <> i And txtName(i).Text = txtName(j).Text Then
                MsgBox "�� " & j & " �ĸ������������ " & i & " �ĸ����������ظ���", vbInformation, App.Title
                txtName(j).SetFocus: Exit Sub
            End If
        Next
        
        If TLen(cboValue(i).Text) > 255 Then
            MsgBox "�� " & i & " �ĸ����� [" & txtName(i).Text & "] ȱʡֵ���Ȳ��ܳ���255���ַ���", vbInformation, App.Title
            cboValue(i).SetFocus: Exit Sub
        End If
        If TLen(cboGroup(i).Text) > 30 Then
            MsgBox "�� " & i & " �ĸ����� [" & txtName(i).Text & "] ���������Ȳ��ܳ���30���ַ���", vbInformation, App.Title
            cboGroup(i).SetFocus: Exit Sub
        End If
        
        If cboValue(i).Text <> "" And Not cboValue(i).Text Like "*��" Then
            If cboType(i).ListIndex = 1 Then
                If Not IsNumeric(cboValue(i).Text) Then
                    MsgBox "�� " & i & " �ĸ����� [" & txtName(i).Text & "] ȱʡֵ����Ӧ��Ϊ�����ͣ�", vbInformation, App.Title
                    cboValue(i).SetFocus: Exit Sub
                End If
            ElseIf cboType(i).ListIndex = 2 Then
                If Not IsDate(cboValue(i).Text) And cboValue(i).ListIndex = -1 Then
                    MsgBox "�� " & i & " �ĸ����� [" & txtName(i).Text & "] ȱʡֵ����Ӧ��Ϊ����/ʱ���ͣ�", vbInformation, App.Title
                    cboValue(i).SetFocus: Exit Sub
                End If
            End If
        End If
            
        '�����в������������ͬ�����ͺ���ȱʡֵӦ����ͬ
'        For Each tmpData In frmSQLEdit.objDatas
'            If tmpData.���� <> frmSQLEdit.objData.���� Then
'                For Each tmpPar In tmpData.Pars
'                    If tmpPar.���� = txtName(i).Text And (tmpPar.���� <> cboType(i).ListIndex Or tmpPar.ȱʡֵ <> cboValue(i).Text) Then
'                        MsgBox "�ڱ�����������Դ�з�������ͬ���ƵĲ���""" & txtName(i).Text & """,�����ǵ����ͻ�ȱʡֵ����ͬ��", vbInformation, App.Title: Exit Sub
'                    End If
'                Next
'            End If
'        Next

'        '��鵱ǰ�����Ƿ�����������Դ�Ĳ���ͬ��
'        For Each tmpData In mobjDatas
'            If tmpData.���� <> mobjData.���� Then
'                For Each tmpPar In tmpData.Pars
'                    If UCase(Trim(tmpPar.����)) = UCase(Trim(txtName(i).Text)) Then
'                        MsgBox "��������" & Trim(txtName(i).Text) & "������������Դ�Ĳ��������������飡", vbInformation, App.Title
'                        Exit Sub
'                    End If
'                Next
'            End If
'        Next
        
        '�Զ������ݼ��
        If cboValue(i).Text = "�̶�ֵ�б�" Then
            If arrCustom(i).ֵ�б� = "" Then
                MsgBox "�� " & i & " �ĸ����� [" & txtName(i).Text & "] ��û�ж����ѡ��Ĺ̶�ֵ�б�", vbInformation, App.Title
                cboValue(i).SetFocus: Exit Sub
            End If
            '�������
            Select Case cboType(i).ListIndex
                Case 1 '����
                    '���õķָ���
                    For j = 0 To UBound(Split(arrCustom(i).ֵ�б�, "|"))
                        If Not IsNumeric(Split(Split(arrCustom(i).ֵ�б�, "|")(j), ",")(1)) Then
                            MsgBox "�� " & i & " �ĸ����� [" & txtName(i).Text & "] �Ĺ̶�ֵ�б��д��ڷ������Ͱ�ֵ��", vbInformation, App.Title
                            cboValue(i).SetFocus: Exit Sub
                        End If
                    Next
                Case 2 '����
                    '���õķָ���
                    For j = 0 To UBound(Split(arrCustom(i).ֵ�б�, "|"))
                        If Not IsDate(Split(Split(arrCustom(i).ֵ�б�, "|")(j), ",")(1)) Then
                            MsgBox "�� " & i & " �ĸ����� [" & txtName(i).Text & "] �Ĺ̶�ֵ�б��д��ڷ������Ͱ�ֵ��", vbInformation, App.Title
                            cboValue(i).SetFocus: Exit Sub
                        End If
                    Next
            End Select
        End If
        
        If cboValue(i).Text = "ѡ�������塭" Then
            If arrCustom(i).��ϸSQL = "" Then
                MsgBox "�� " & i & " �ĸ����� [" & txtName(i).Text & "] ��û�ж���ѡ���������ݣ�", vbInformation, App.Title
                cboValue(i).SetFocus: Exit Sub
            End If
            '�������(֮����Ҫ���ж�һ����Ϊ�û����ܸ�������)
            For j = 0 To UBound(Split(arrCustom(i).��ϸ�ֶ�, "|"))
                If InStr(Split(Split(arrCustom(i).��ϸ�ֶ�, "|")(j), ",")(2), "&B") > 0 Then
                    If cboType(i).ListIndex = 1 Then
                        Select Case CLng(Split(Split(arrCustom(i).��ϸ�ֶ�, "|")(j), ",")(1))
                            Case adNumeric, adVarNumeric  '����������
                            Case Else '��������
                                If MsgBox("�� " & i & " �ĸ����� [" & txtName(i).Text & "] ��ѡ�������ֶ� [" & Split(Split(arrCustom(i).��ϸ�ֶ�, "|")(j), ",")(0) & "] ����������,Ҫ������", vbQuestion + vbYesNo, App.Title) = vbNo Then
                                    cboValue(i).SetFocus: Exit Sub
                                End If
                        End Select
                    ElseIf cboType(i).ListIndex = 2 Then
                        Select Case CLng(Split(Split(arrCustom(i).��ϸ�ֶ�, "|")(j), ",")(1))
                            Case adDBTimeStamp '����������
                            Case Else '��������
                                If MsgBox("�� " & i & " �ĸ����� [" & txtName(i).Text & "] ��ѡ�������ֶ� [" & Split(Split(arrCustom(i).��ϸ�ֶ�, "|")(j), ",")(0) & "] ����������,Ҫ������", vbQuestion + vbYesNo, App.Title) = vbNo Then
                                    cboValue(i).SetFocus: Exit Sub
                                End If
                        End Select
                    End If
                End If
            Next
        End If
    Next
    
    '�����������ٺ�����������(��ֻ�������ڵĲ�����������ͬһ��)��ѡ��������������
    Dim IntPar As Integer, StrPar As String, IntUseCount As Integer
    IntUseCount = 0: StrPar = ""
    For IntPar = 0 To UBound(arrCustom)
        If arrCustom(IntPar).��ʽ = 1 And arrCustom(IntPar).���� <> "" Then
            MsgBox "����" & IntPar & "�������������飨��ò�����ѡ��ģʽ�ǵ�ѡ�򣩣�", vbInformation, App.Title
            cboGroup(IntPar).SetFocus
            Exit Sub
        End If
    Next
    For IntPar = 0 To UBound(arrCustom)
        If StrPar <> arrCustom(IntPar).���� Then
            If arrCustom(IntPar).���� = "" Then
                If Not (IntUseCount = 0 Or IntUseCount > 1) Then
                    MsgBox "ÿ������������������������" & IntPar & "����", vbInformation, App.Title
                    cboGroup(IntPar).SetFocus
                    Exit Sub
                End If
                StrPar = arrCustom(IntPar).����
                IntUseCount = 1
            Else
                If IntUseCount = 0 Or IntUseCount > 1 Or StrPar = "" Then
                    StrPar = arrCustom(IntPar).����
                    IntUseCount = 1
                Else
                    MsgBox "ÿ������������������������" & IntPar & "����", vbInformation, App.Title
                    cboGroup(IntPar).SetFocus
                    Exit Sub
                End If
            End If
        Else
            IntUseCount = IntUseCount + 1
        End If
    Next
    If Not (IntUseCount = 0 Or IntUseCount > 1 Or StrPar = "") Then
        MsgBox "ÿ������������������������" & IntPar - 1 & "����", vbInformation, App.Title
        cboGroup(IntPar - 1).SetFocus
        Exit Sub
    End If
    
    '��ȡ֮ǰ�Ĳ���
    For i = 1 To mobjPars.count
        strAddDelBefor = strAddDelBefor & "," & mobjPars.Item(i).����
    Next
    'ȷ����������
    Set mobjPars = New RPTPars
    For i = 0 To lblNO.UBound
        '�����ʹ������Դ�Ĳ����������,�������Ӱ����Ȩ
        If cboValue(i).Text <> "ѡ�������塭" Then
            arrCustom(i).��ϸSQL = ""
            arrCustom(i).��ϸ�ֶ� = ""
            arrCustom(i).����SQL = ""
            arrCustom(i).�����ֶ� = ""
            arrCustom(i).���� = ""
        End If
        Set curPar = Nothing
        Set curPar = mobjPars.Add(arrCustom(i).����, CByte(i), txtName(i).Text, cboType(i).ListIndex, cboValue(i).Text _
                        , arrCustom(i).��ʽ, arrCustom(i).ֵ�б�, arrCustom(i).����SQL, arrCustom(i).��ϸSQL _
                        , arrCustom(i).�����ֶ�, arrCustom(i).��ϸ�ֶ�, arrCustom(i).����, "_" & i, , chkLock(i).Value)
                        
        'ͬʱ�Զ��滻��������Դ����ͬ���Ʋ���������
        For Each tmpData In mobjDatas
            If tmpData.���� <> mobjData.���� Then
                For Each tmpPar In tmpData.Pars
                    If tmpPar.���� = curPar.���� Then
                        tmpPar.��ʽ = curPar.��ʽ
                        tmpPar.���� = curPar.����
                        tmpPar.ȱʡֵ = curPar.ȱʡֵ
                        tmpPar.ֵ�б� = curPar.ֵ�б�
                        tmpPar.��ϸSQL = curPar.��ϸSQL
                        tmpPar.��ϸ�ֶ� = curPar.��ϸ�ֶ�
                        tmpPar.����SQL = curPar.����SQL
                        tmpPar.�����ֶ� = curPar.�����ֶ�
                        tmpPar.���� = curPar.����
                        tmpPar.�Ƿ����� = curPar.�Ƿ�����
                    End If
                Next
            End If
        Next
    Next
    '��ȡ֮ǰ�Ĳ���
    For i = 1 To mobjPars.count
        strAddDelAfter = strAddDelAfter & "," & mobjPars.Item(i).����
    Next
    If strAddDelAfter <> strAddDelBefor Then
        '��ʾ�����˱���ı���
        strSQL = "Select Distinct b.���, b.���� From Zlrptrelation A, zlReports B Where a.����id = b.Id And a.��������id = [1] "
        On Error GoTo errH
        Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, mlngReportID)
        Do While Not rsTmp.EOF
            strInfo = strInfo & vbCrLf & rsTmp!���� & "(" & rsTmp!��� & ")"
            rsTmp.MoveNext
        Loop
        strInfo = Mid(strInfo, 2)
        If strInfo <> "" Then
            MsgBox "���±���������ѯ����������������˲����󣬿�����Ҫ�������±���Ĺ�����Ϣ�����飺" & strInfo, vbInformation, Me.Caption
        End If
    End If
    
    mblnOK = True
    Hide
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0: SendKeys "{Tab}"
End Sub

Private Sub Form_Load()
    Dim intCount As Integer
    Dim i As Integer
    
    mblnOK = False
    intPreIdx = -1
    
    intCount = GetParCount(mstrSQL)
    
    ReDim arrCustom(intCount - 1) As CustomPar
    
    For i = 0 To intCount - 1
        If i <> 0 Then
            Load lblNO(i): lblNO(i).Left = lblNO(0).Left: lblNO(i).Top = lblNO(0).Top + 450 * i: lblNO(i).Visible = True
            Load txtName(i): txtName(i).Left = txtName(0).Left: txtName(i).Top = txtName(0).Top + 450 * i: txtName(i).TabIndex = txtName(0).TabIndex + 5 * i: txtName(i).Visible = True
            Load cboType(i): cboType(i).Left = cboType(0).Left: cboType(i).Top = cboType(0).Top + 450 * i: cboType(i).TabIndex = cboType(0).TabIndex + 5 * i: cboType(i).Visible = True
            Load cboValue(i): cboValue(i).Left = cboValue(0).Left: cboValue(i).Top = cboValue(0).Top + 450 * i: cboValue(i).TabIndex = cboValue(0).TabIndex + 5 * i: cboValue(i).Visible = True
            Load cboGroup(i): cboGroup(i).Left = cboGroup(0).Left: cboGroup(i).Top = cboGroup(0).Top + 450 * i: cboGroup(i).TabIndex = cboGroup(0).TabIndex + 5 * i: cboGroup(i).Visible = True
            Load chkLock(i): chkLock(i).Left = chkLock(0).Left: chkLock(i).Top = chkLock(0).Top + 450 * i: chkLock(i).TabIndex = chkLock(0).TabIndex + 5 * i: chkLock(i).Visible = True: chkLock(i).Value = 0
        End If
        lblNO(i).Caption = i
        cboType(i).AddItem "�ַ�": cboType(i).AddItem "����": cboType(i).AddItem "����": cboType(i).AddItem "������"
        cboValue(i).AddItem "&��ǰ����" '����ʽ
        cboValue(i).AddItem "&��ǰ����ʱ��"
        
        cboValue(i).AddItem "&���쿪ʼʱ��"
        cboValue(i).AddItem "&�������ʱ��"
        cboValue(i).AddItem "&ǰһ�쿪ʼʱ��"
        cboValue(i).AddItem "&ǰһ�����ʱ��"
        cboValue(i).AddItem "&ǰһ��ͬʱ��"
        cboValue(i).AddItem "&��һ��ͬʱ��"
        cboValue(i).AddItem "&��һ�����ʱ��"
        cboValue(i).AddItem "&��һ������"
        
        cboValue(i).AddItem "&ǰһ������"
        cboValue(i).AddItem "&ǰһ������"
        cboValue(i).AddItem "&ǰһ������"
        cboValue(i).AddItem "&ǰһ������"
        
        cboValue(i).AddItem "&��һ������"
        cboValue(i).AddItem "&��һ������"
        cboValue(i).AddItem "&��һ������"
        cboValue(i).AddItem "&��һ������"
        
        cboValue(i).AddItem "&���³�ʱ��"
        cboValue(i).AddItem "&����ĩʱ��"
        cboValue(i).AddItem "&���³�ʱ��"
        cboValue(i).AddItem "&����ĩʱ��"
        cboValue(i).AddItem "&�����ʱ��"
        cboValue(i).AddItem "&����ĩʱ��"
        cboValue(i).AddItem "&�����ʱ��"
        cboValue(i).AddItem "&����ĩʱ��"
        
        '�����Զ�������
        cboValue(i).AddItem "�̶�ֵ�б�"
        cboValue(i).AddItem "ѡ�������塭"
        
        If mobjPars.count >= i + 1 Then '��������ԭ������
            txtName(i).Text = mobjPars("_" & i).����
            cboType(i).ListIndex = mobjPars("_" & i).����
            If Left(mobjPars("_" & i).ȱʡֵ, 1) = "&" Or mobjPars("_" & i).ȱʡֵ Like "*��" Then
                cboValue(i).ListIndex = GetCboIndex(cboValue(i), mobjPars("_" & i).ȱʡֵ)
            Else
                cboValue(i).Text = mobjPars("_" & i).ȱʡֵ
            End If
            chkLock(i).Value = IIF(mobjPars("_" & i).�Ƿ�����, 1, 0)
            
            '�Զ�������
            arrCustom(i).ֵ�б� = mobjPars("_" & i).ֵ�б�
            arrCustom(i).����SQL = mobjPars("_" & i).����SQL
            arrCustom(i).��ϸSQL = mobjPars("_" & i).��ϸSQL
            arrCustom(i).�����ֶ� = mobjPars("_" & i).�����ֶ�
            arrCustom(i).��ϸ�ֶ� = mobjPars("_" & i).��ϸ�ֶ�
            arrCustom(i).���� = mobjPars("_" & i).����
            arrCustom(i).��ʽ = mobjPars("_" & i).��ʽ
            arrCustom(i).���� = mobjPars("_" & i).����
        Else
            txtName(i).Text = ""
            cboType(i).ListIndex = 0
            cboValue(i).Text = ""
        End If
    Next
    Call LoadGroup
    
    Height = txtName(txtName.UBound).Top + 1365
End Sub

Private Sub txtName_GotFocus(Index As Integer)
    SelAll txtName(Index)
End Sub

Private Sub txtName_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim tmpData As RPTData, tmpPar As RPTPar
    
    If KeyCode = 13 And txtName(Index).Text <> "" Then
        For Each tmpData In mobjDatas
            If tmpData.���� <> mobjData.���� Then
                For Each tmpPar In tmpData.Pars
                    If tmpPar.���� = txtName(Index).Text Then
                        cboType(Index).ListIndex = tmpPar.����
                        cboValue(Index).ListIndex = GetCboIndex(cboValue(Index), tmpPar.ȱʡֵ)
                        If cboValue(Index).ListIndex = -1 Then cboValue(Index).Text = tmpPar.ȱʡֵ
                    End If
                Next
            End If
        Next
    End If
End Sub

Private Sub txtName_KeyPress(Index As Integer, KeyAscii As Integer)
    If InStr("~`@#$%^&*()=+][}{'"";/?.>,<\|", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub LoadGroup()
    Dim ItemPar As RPTPar, strGroup As String, arrGroup
    Dim IntGroup As Integer, IntAdd As Integer
    '�������
    strGroup = ""
    For Each ItemPar In mobjPars
        If InStr(1, strGroup, "^" & ItemPar.���� & ",") = 0 And ItemPar.���� <> "" Then strGroup = strGroup & "^" & ItemPar.���� & ","
    Next
    If strGroup <> "" Then
        strGroup = Mid(strGroup, 1, Len(strGroup) - 1)
        arrGroup = Split(strGroup, ",")
        
        'Ϊ�����������
        For IntGroup = 0 To cboGroup.UBound
            cboGroup(IntGroup).Clear
            For IntAdd = 0 To UBound(arrGroup)
                cboGroup(IntGroup).AddItem Mid(arrGroup(IntAdd), 2)
            Next
            If cboGroup(IntGroup).ListCount <> 0 Then cboGroup(IntGroup).Text = arrCustom(IntGroup).����
        Next
    End If
End Sub
