VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmParInput 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2040
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6435
   Icon            =   "frmParInput.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   6435
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picCmd 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   6435
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1365
      Width           =   6435
      Begin VB.CheckBox chkAutoSave 
         Caption         =   "�������ʱ��"
         Height          =   270
         Left            =   195
         TabIndex        =   11
         Top             =   220
         Visible         =   0   'False
         Width           =   1380
      End
      Begin VB.CommandButton cmdDefault 
         Caption         =   "����(&D)"
         Height          =   350
         Left            =   1755
         TabIndex        =   12
         Top             =   180
         Width           =   1100
      End
      Begin VB.Frame Frame1 
         Height          =   75
         Left            =   -150
         TabIndex        =   14
         Top             =   -45
         Width           =   7290
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "ȷ��(&O)"
         Height          =   350
         Left            =   4035
         TabIndex        =   9
         Top             =   180
         Width           =   1100
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "ȡ��(&C)"
         Height          =   350
         Left            =   5160
         TabIndex        =   10
         Top             =   180
         Width           =   1100
      End
   End
   Begin VB.PictureBox picInfo 
      Align           =   1  'Align Top
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   630
      Left            =   0
      ScaleHeight     =   630
      ScaleWidth      =   6435
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   6435
      Begin VB.Frame Frame2 
         Height          =   45
         Left            =   -45
         TabIndex        =   2
         Top             =   570
         Width           =   7000
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "���ڸ������������������ѡ���㱾�β�ѯ����Ҫ������ֵ��"
         ForeColor       =   &H00800000&
         Height          =   360
         Left            =   1080
         TabIndex        =   1
         Top             =   120
         UseMnemonic     =   0   'False
         Width           =   4020
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   390
         Picture         =   "frmParInput.frx":014A
         Top             =   75
         Width           =   480
      End
   End
   Begin VB.PictureBox picPar 
      Align           =   1  'Align Top
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   0
      ScaleHeight     =   705
      ScaleWidth      =   6435
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   630
      Width           =   6435
      Begin VB.CommandButton cmdSelNone 
         Cancel          =   -1  'True
         Caption         =   "ȫ��"
         Height          =   350
         Left            =   5685
         TabIndex        =   19
         Top             =   240
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.CommandButton cmdSelAll 
         Caption         =   "ȫѡ"
         Height          =   350
         Left            =   5040
         TabIndex        =   18
         Top             =   240
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.Frame fraGroup 
         ForeColor       =   &H00800000&
         Height          =   180
         Index           =   0
         Left            =   1125
         TabIndex        =   17
         Top             =   -60
         Visible         =   0   'False
         Width           =   3825
      End
      Begin VB.Frame fra 
         ForeColor       =   &H00800000&
         Height          =   645
         Index           =   0
         Left            =   1125
         TabIndex        =   15
         Top             =   60
         Visible         =   0   'False
         Width           =   3825
         Begin VB.OptionButton opt 
            Caption         =   "#"
            Height          =   180
            Index           =   0
            Left            =   105
            MaskColor       =   &H8000000F&
            TabIndex        =   16
            Top             =   270
            Visible         =   0   'False
            Width           =   1150
         End
      End
      Begin VB.CommandButton cmd 
         Caption         =   "��"
         Height          =   240
         Index           =   0
         Left            =   4425
         TabIndex        =   5
         ToolTipText     =   "�� F2 ��ѡ����"
         Top             =   225
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   0
         Left            =   2250
         TabIndex        =   4
         Top             =   195
         Visible         =   0   'False
         Width           =   2475
      End
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   0
         Left            =   2250
         TabIndex        =   6
         Top             =   195
         Visible         =   0   'False
         Width           =   2460
      End
      Begin MSComCtl2.DTPicker dtp 
         Height          =   300
         Index           =   0
         Left            =   2250
         TabIndex        =   7
         Top             =   195
         Visible         =   0   'False
         Width           =   2460
         _ExtentX        =   4339
         _ExtentY        =   529
         _Version        =   393216
         CalendarTitleBackColor=   12946264
         CalendarTitleForeColor=   16777215
         CustomFormat    =   "yyyy��MM��dd�� HH:mm:ss"
         Format          =   416743427
         CurrentDate     =   36731
      End
      Begin VB.CheckBox chk 
         Caption         =   "#"
         Height          =   195
         Index           =   0
         Left            =   2250
         TabIndex        =   20
         Top             =   240
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������"
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   0
         Left            =   1470
         TabIndex        =   3
         Top             =   255
         Visible         =   0   'False
         Width           =   720
      End
   End
   Begin VB.Menu PopMenu 
      Caption         =   "�����˵�(&P)"
      Visible         =   0   'False
      Begin VB.Menu PopMenu_Cond 
         Caption         =   "����1"
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu split0 
         Caption         =   "-"
      End
      Begin VB.Menu PopMenu_Save 
         Caption         =   "����(&S)"
      End
      Begin VB.Menu PopMenu_Saveas 
         Caption         =   "���Ϊ(&A)..."
      End
      Begin VB.Menu PopMenu_Del 
         Caption         =   "ɾ��(&C)"
         Enabled         =   0   'False
      End
      Begin VB.Menu split1 
         Caption         =   "-"
      End
      Begin VB.Menu PopMenu_Default 
         Caption         =   "ȱʡ(&D)"
      End
   End
End
Attribute VB_Name = "frmParInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mblnReset As Boolean 'I:�Ƿ��û�ѡ����������������
Public mstrTitle As String 'I:�������
Public mobjPars As RPTPars  'IO:������Ψһ�Ĳ�������
Public mobjDefPars As RPTPars '��ǰ����ԭʼ�Ĳ�������,���ڻָ�ȱʡֵ
Public mlngReport As Long   '����ID
Public mblnOK As Boolean
Public mobjRPTDatas As RPTDatas

Private mint������ As Integer '����������
Private mintMenu As Integer   '��ǰѡ��������˵�������
Private mblnMatch As Boolean
Private mintBegin As Integer
Private mintEnd As Integer

Private Sub cbo_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim LngIdx As Long
    
    If KeyAscii = 13 Then KeyAscii = 0: SendKeys "{Tab}": Exit Sub
    If InStr("~`!@#$^&"";|'", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    If mobjPars("_" & lbl(Index).ToolTipText).���� = 1 Then
        If InStr("-0.123456789" & Chr(8) & Chr(3) & Chr(22) & Chr(24) & Chr(26), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
    End If
End Sub

Private Sub chk_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0: SendKeys "{Tab}": Exit Sub
End Sub

Private Sub cmd_Click(Index As Integer)
    Dim tmpPar As RPTPar, str��ϸ���� As String, str������� As String
    Dim frmNewSelect As New frmSelect
    Dim strSQL��ϸ As String, strSQL���� As String
    Dim colValue As New Collection    '�������е�ֵ
    
    For Each tmpPar In mobjPars
        If tmpPar.���� = lbl(Index).ToolTipText Then
            If mblnMatch And txt(Index).Tag = "" Then frmNewSelect.strMatch = txt(Index).Text
            
            If InStr(tmpPar.����, "|") > 0 Then
                str��ϸ���� = Split(tmpPar.����, "|")(0)
                str������� = Split(tmpPar.����, "|")(1)
            End If
            strSQL��ϸ = tmpPar.��ϸSQL
            strSQL���� = tmpPar.����SQL
            Set colValue = GetValues
            Call CheckParsRela(strSQL��ϸ, Nothing, tmpPar.����, True, colValue, mobjPars)
            Call CheckParsRela(strSQL����, Nothing, tmpPar.����, True, colValue, mobjPars)
            frmNewSelect.strSQLList = SQLOwner(RemoveNote(strSQL��ϸ), str��ϸ����)
            frmNewSelect.strSQLTree = SQLOwner(RemoveNote(strSQL����), str�������)
            frmNewSelect.strFLDList = tmpPar.��ϸ�ֶ�
            frmNewSelect.strFLDTree = tmpPar.�����ֶ�
            frmNewSelect.strParName = tmpPar.����
            frmNewSelect.bytType = tmpPar.����
            frmNewSelect.mblnMulti = tmpPar.��ʽ = 1
            frmNewSelect.lngSeekHwnd = cmd(Index).hwnd
            frmNewSelect.mintConnect = GetDBConnectNo(tmpPar, mobjRPTDatas)
            
            On Error Resume Next
            Err.Clear
            
            frmNewSelect.Show 1, Me
            If frmNewSelect.mblnOK Then
                txt(Index).Text = frmNewSelect.strOutDisp
                txt(Index).Tag = frmNewSelect.strOutBand
                Unload frmNewSelect
                
                SendKeys "{Tab}"
                mblnOK = False '�ָ�����ʱ��״̬
            ElseIf mblnMatch Then
                txt(Index).Text = ""
                txt(Index).Tag = ""
            End If
            
            mblnMatch = False
            Exit For
        End If
    Next
    txt(Index).SetFocus
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDefault_Click()
    Call PopupMenu(PopMenu, , cmdDefault.Left, picCmd.Top + cmdDefault.Top + cmdDefault.Height)
End Sub

Private Sub cmdOK_Click()
    Dim i As Integer, j As Integer
    Dim strTmp As String, strDisp As String
    Dim strParName As String, curDate As Date
    
    '�ȼ��Ϸ���
    For i = 1 To lbl.UBound
        strParName = lbl(i).ToolTipText
        
        If mobjPars("_" & strParName).ȱʡֵ = "�̶�ֵ�б�" Then
            Select Case mobjPars("_" & strParName).��ʽ
                Case 0
                    If Trim(cbo(i).Text) = "" Then
                        MsgBox "��ѡ��""" & strParName & """������ֵ��", vbInformation, App.Title
                        If cbo(i).Enabled Then cbo(i).SetFocus
                        Exit Sub
                    End If
                    If GetCboIndex(cbo(i), cbo(i).Text) = -1 Then '�Ƿ���Ϊ����
                        '���ͼ��
                        Select Case mobjPars("_" & strParName).����
                            Case 1
                                If Not IsNumeric(cbo(i).Text) Then
                                    MsgBox "�������""" & strParName & """������ֵ����Ӧ��Ϊ�����ͣ�", vbInformation, App.Title
                                    If cbo(i).Enabled Then cbo(i).SetFocus
                                    Exit Sub
                                End If
                            Case 2
                                If Not IsDate(cbo(i).Text) Then
                                    MsgBox "�������""" & strParName & """������ֵ����Ӧ��Ϊ�����ͣ�", vbInformation, App.Title
                                    If cbo(i).Enabled Then cbo(i).SetFocus
                                    Exit Sub
                                End If
                        End Select
                    End If
            End Select
        ElseIf mobjPars("_" & strParName).ȱʡֵ = "ѡ�������塭" Then
            If Trim(txt(i).Text) = "" Then
                MsgBox "��ѡ��""" & strParName & """������ֵ��", vbInformation, App.Title
                If txt(i).Enabled Then txt(i).SetFocus
                Exit Sub
            End If
            If txt(i).Tag = "" Then '�Ƿ���Ϊ����
                If mobjPars("_" & strParName).ֵ�б� Like "*|*" Then
                    If Split(mobjPars("_" & strParName).ֵ�б�, "|")(0) <> txt(i).Text Then
                        '���ͼ��
                        Select Case mobjPars("_" & strParName).����
                            Case 1
                                If Not IsNumeric(txt(i).Text) Then
                                    MsgBox "�������""" & strParName & """������ֵ����Ӧ��Ϊ�����ͣ�", vbInformation, App.Title
                                    If txt(i).Enabled Then txt(i).SetFocus
                                    Exit Sub
                                End If
                            Case 2
                                If Not IsDate(txt(i).Text) Then
                                    MsgBox "�������""" & strParName & """������ֵ����Ӧ��Ϊ�����ͣ�", vbInformation, App.Title
                                    If txt(i).Enabled Then txt(i).SetFocus
                                    Exit Sub
                                End If
                        End Select
                    Else
                        '����ֵ�붨���ȱʡֵ��ͬ,��ԭΪȱʡֵ
                        txt(i).Tag = Split(mobjPars("_" & strParName).ֵ�б�, "|")(1)
                    End If
                Else
                    '���ͼ��
                    Select Case mobjPars("_" & strParName).����
                        Case 1
                            If Not IsNumeric(txt(i).Text) Then
                                MsgBox "�������""" & strParName & """������ֵ����Ӧ��Ϊ�����ͣ�", vbInformation, App.Title
                                If txt(i).Enabled Then txt(i).SetFocus
                                Exit Sub
                            End If
                        Case 2
                            If Not IsDate(txt(i).Text) Then
                                MsgBox "�������""" & strParName & """������ֵ����Ӧ��Ϊ�����ͣ�", vbInformation, App.Title
                                If txt(i).Enabled Then txt(i).SetFocus
                                Exit Sub
                            End If
                    End Select
                End If
            End If
        Else
            Select Case mobjPars("_" & strParName).����
                Case 0, 3
                    If Trim(txt(i).Text) = "" Then
                        MsgBox "������""" & strParName & """������ֵ��", vbInformation, App.Title
                        If txt(i).Enabled Then txt(i).SetFocus
                        Exit Sub
                    End If
                    If TLen(txt(i).Text) > 4000 Then
                        MsgBox """" & strParName & """������ֵ���Ȳ��ܳ���4000���ַ���", vbInformation, App.Title
                        If txt(i).Enabled Then txt(i).SetFocus
                        Exit Sub
                    End If
                Case 1
                    If Trim(txt(i).Text) = "" Then
                        MsgBox "������""" & strParName & """������ֵ��", vbInformation, App.Title
                        If txt(i).Enabled Then txt(i).SetFocus
                        Exit Sub
                    End If
                    If TLen(txt(i).Text) > 4000 Then
                        MsgBox """" & strParName & """������ֵ���Ȳ��ܳ���4000���ַ���", vbInformation, App.Title
                        If txt(i).Enabled Then txt(i).SetFocus
                        Exit Sub
                    End If
                    If Not IsNumeric(txt(i).Text) Then
                        MsgBox """" & strParName & """������ֵ����Ӧ��Ϊ�����ͣ�", vbInformation, App.Title
                        If txt(i).Enabled Then txt(i).SetFocus
                        Exit Sub
                    End If
                Case 2 '����ʱ�����ֵ���
                    curDate = Currentdate
                    If Not (mobjPars("_" & strParName).ȱʡֵ Like "&��һ*" Or mobjPars("_" & strParName).Reserve Like "&��һ*" Or _
                        mobjPars("_" & strParName).ȱʡֵ Like "&��һ*" Or mobjPars("_" & strParName).Reserve Like "&��һ*" Or _
                        mobjPars("_" & strParName).ȱʡֵ Like "&*����*" Or mobjPars("_" & strParName).Reserve Like "&*����*" Or _
                        mobjPars("_" & strParName).ȱʡֵ Like "&*��ĩ*" Or mobjPars("_" & strParName).ȱʡֵ Like "&*��ĩ*" Or _
                        mobjPars("_" & strParName).Reserve Like "&*��ĩ*" Or mobjPars("_" & strParName).Reserve Like "&*��ĩ*") Then
                        
                        If mobjPars("_" & strParName).ȱʡֵ Like "*ʱ��*" Or mobjPars("_" & strParName).Reserve Like "*ʱ��*" Then
                            If Format(dtp(i).Value, "yyyy-MM-dd HH:mm:ss") > Format(curDate, "yyyy-MM-dd HH:mm:ss") Then
                                MsgBox """" & strParName & """ ������ֵ���ܳ�����ǰʱ�䣡", vbInformation, App.Title
                                If dtp(i).Enabled Then dtp(i).SetFocus
                                Exit Sub
                            End If
                        Else
                            If Format(dtp(i).Value, "yyyy-MM-dd") > Format(curDate, "yyyy-MM-dd") Then
                                MsgBox """" & strParName & """ ������ֵ���ܳ�����ǰ���ڣ�", vbInformation, App.Title
                                If dtp(i).Enabled Then dtp(i).SetFocus
                                Exit Sub
                            End If
                        End If
                    End If
            End Select
        End If
    Next
        
    '��ȡֵ
    For i = 1 To lbl.UBound
        strParName = lbl(i).ToolTipText
        
        If mobjPars("_" & strParName).ȱʡֵ = "�̶�ֵ�б�" Then
            Select Case mobjPars("_" & strParName).��ʽ
                Case 0
                    If GetCboIndex(cbo(i), cbo(i).Text) = -1 Then '�Ƿ���Ϊ����
                        'Reserve�ֶα��汾��������"������ֵ|��ʾֵ"
                        mobjPars("_" & strParName).Reserve = "�̶�ֵ�б�|" & cbo(i).Text
                        mobjPars("_" & strParName).ȱʡֵ = cbo(i).Text
                    Else
                        '�б�ѡ��
                        'Reserve�ֶα��汾��������"������ֵ|��ʾֵ"
                        '���õķָ���
                        mobjPars("_" & strParName).Reserve = "�̶�ֵ�б�|" & cbo(i).Text
                        strTmp = mobjPars("_" & strParName).ֵ�б�
                        For j = 0 To UBound(Split(strTmp, "|"))
                            strDisp = Split(Split(strTmp, "|")(j), ",")(0)
                            If Left(strDisp, 1) = "��" Then strDisp = Mid(strDisp, 2)
                            If strDisp = cbo(i).Text Then
                                mobjPars("_" & strParName).ȱʡֵ = Split(Split(strTmp, "|")(j), ",")(1)
                                Exit For
                            End If
                        Next
                    End If
                Case 1
                    For j = 1 To opt.UBound
                        If opt(j).Container.Index = i Then
                            If opt(j).Value Then
                                'Reserve�ֶα��汾��������"������ֵ|��ʾֵ"
                                mobjPars("_" & strParName).Reserve = "�̶�ֵ�б�|" & opt(j).ToolTipText
                                mobjPars("_" & strParName).ȱʡֵ = opt(j).Tag
                            End If
                        End If
                    Next
                Case 2
                    'Reserve�ֶα��汾��������"������ֵ|��ʾֵ"
                    '���õķָ���
                    strTmp = mobjPars("_" & strParName).ֵ�б�
                    For j = 0 To 1
                        strDisp = Split(Split(strTmp, "|")(j), ",")(0)
                        If chk(i).Value = 0 Then
                            If Left(strDisp, 1) <> "��" Then
                                mobjPars("_" & strParName).Reserve = "�̶�ֵ�б�|" & strDisp
                                mobjPars("_" & strParName).ȱʡֵ = Split(Split(strTmp, "|")(j), ",")(1)
                            End If
                        Else
                            If Left(strDisp, 1) = "��" Then
                                mobjPars("_" & strParName).Reserve = "�̶�ֵ�б�|" & Mid(strDisp, 2)
                                mobjPars("_" & strParName).ȱʡֵ = Split(Split(strTmp, "|")(j), ",")(1)
                            End If
                        End If
                    Next
            End Select
        ElseIf mobjPars("_" & strParName).ȱʡֵ = "ѡ�������塭" Then
            If txt(i).Tag = "" Then '�Ƿ���Ϊ����
                'Reserve�ֶα��汾��������"������ֵ|��ʾֵ"
                mobjPars("_" & strParName).Reserve = "ѡ�������塭|"
                mobjPars("_" & strParName).ȱʡֵ = txt(i).Text
            Else
                '�б�ѡ��
                'Reserve�ֶα��汾��������"������ֵ|��ʾֵ"
                mobjPars("_" & strParName).Reserve = "ѡ�������塭|" & txt(i).Text
                mobjPars("_" & strParName).ȱʡֵ = txt(i).Tag
            End If
        Else
            Select Case mobjPars("_" & strParName).����
                Case 0, 1, 3
                    mobjPars("_" & strParName).ȱʡֵ = txt(i).Text
                Case 2
                    If mobjPars("_" & strParName).ȱʡֵ Like "&*" Then
                        mobjPars("_" & strParName).Reserve = mobjPars("_" & strParName).ȱʡֵ
                    End If
                    mobjPars("_" & strParName).ȱʡֵ = Format(dtp(i).Value, dtp(i).CustomFormat)
                    '���浽ע���
                    If dtp(i).CustomFormat Like "*HH:mm:ss" Then
                        SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.name & mstrTitle, lbl(i).ToolTipText & "ʱ��", Format(dtp(i).Value, "HH:mm:ss")
                    End If
            End Select
        End If
    Next
    
    '������Կ�ʼ����ʱ��(���ܹ�û��)
    If mintBegin <> -1 And mintEnd <> -1 Then
        SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.name & mstrTitle, "AutoSave", chkAutoSave.Value
        SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.name & mstrTitle, "BeginTime", Format(dtp(mintBegin).Value, dtp(mintBegin).CustomFormat)
        SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.name & mstrTitle, "EndTime", Format(dtp(mintEnd).Value, dtp(mintEnd).CustomFormat)
    End If
    
    mblnOK = True
    Hide
End Sub

Private Function GetValues() As Collection
'���ܣ���ȡ���еĽ����ϵĲ���ֵ
    Dim i As Integer, j As Integer
    Dim strParName As String, strTmp As String
    Dim strDisp As String, colValue As New Collection
     
    For i = 1 To lbl.UBound
        strParName = lbl(i).ToolTipText
        
        If mobjPars("_" & strParName).ȱʡֵ = "�̶�ֵ�б�" Then
            Select Case mobjPars("_" & strParName).��ʽ
                Case 0
                    If GetCboIndex(cbo(i), cbo(i).Text) = -1 Then '�Ƿ���Ϊ����
                        'Reserve�ֶα��汾��������"������ֵ|��ʾֵ"
                        colValue.Add cbo(i).Text, "_" & strParName
                    Else
                        '�б�ѡ��
                        'Reserve�ֶα��汾��������"������ֵ|��ʾֵ"
                        '���õķָ���
                        strTmp = mobjPars("_" & strParName).ֵ�б�
                        For j = 0 To UBound(Split(strTmp, "|"))
                            strDisp = Split(Split(strTmp, "|")(j), ",")(0)
                            If Left(strDisp, 1) = "��" Then strDisp = Mid(strDisp, 2)
                            If strDisp = cbo(i).Text Then
                                colValue.Add Split(Split(strTmp, "|")(j), ",")(1), "_" & strParName
                                Exit For
                            End If
                        Next
                    End If
                Case 1
                    For j = 1 To opt.UBound
                        If opt(j).Container.Index = i Then
                            If opt(j).Value Then
                                colValue.Add opt(j).Tag, "_" & strParName
                            End If
                        End If
                    Next
                Case 2
                    'Reserve�ֶα��汾��������"������ֵ|��ʾֵ"
                    '���õķָ���
                    strTmp = mobjPars("_" & strParName).ֵ�б�
                    For j = 0 To 1
                        strDisp = Split(Split(strTmp, "|")(j), ",")(0)
                        If chk(i).Value = 0 Then
                            If Left(strDisp, 1) <> "��" Then
                                colValue.Add Split(Split(strTmp, "|")(j), ",")(1), "_" & strParName
                            End If
                        Else
                            If Left(strDisp, 1) = "��" Then
                                colValue.Add Split(Split(strTmp, "|")(j), ",")(1), "_" & strParName
                            End If
                        End If
                    Next
            End Select
        ElseIf mobjPars("_" & strParName).ȱʡֵ = "ѡ�������塭" Then
            If txt(i).Tag = "" Then '�Ƿ���Ϊ����
                'Reserve�ֶα��汾��������"������ֵ|��ʾֵ"
                colValue.Add txt(i).Text, "_" & strParName
            Else
                '�б�ѡ��
                'Reserve�ֶα��汾��������"������ֵ|��ʾֵ"
                colValue.Add txt(i).Tag, "_" & strParName
            End If
        Else
            Select Case mobjPars("_" & strParName).����
                Case 0, 1, 3
                    colValue.Add txt(i).Text, "_" & strParName
                Case 2
                    strTmp = dtp(i).CustomFormat
                    If strTmp Like "* *:*:*" Then
                        colValue.Add Format(dtp(i).Value, "YYYY-MM-DD hh:mm:ss"), "_" & strParName
                    Else
                        colValue.Add Format(dtp(i).Value, "YYYY-MM-DD"), "_" & strParName
                    End If
            End Select
        End If
    Next
    Set GetValues = colValue
End Function

Private Sub cmdSelAll_Click()
    Dim chkTmp As CheckBox
    
    For Each chkTmp In chk
        If chkTmp.Enabled Then
            chkTmp.Value = 1
        End If
    Next
End Sub

Private Sub cmdSelNone_Click()
    Dim chkTmp As CheckBox
    
    For Each chkTmp In chk
        If chkTmp.Enabled Then
            chkTmp.Value = 0
        End If
    Next
End Sub

Private Sub dtp_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{Tab}"
End Sub

Private Sub dtp_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mblnReset = False
    Set mobjPars = Nothing
    Me.Tag = ""
End Sub

Private Sub opt_GotFocus(Index As Integer)
    If opt(Index).Value Then
        '��������Ŀ���Ǳ��ⰴTAB��ʱ�Զ��л�����һ��ѡ��
        opt(Index).Value = False
        opt(Index).Value = True
    End If
End Sub

Private Sub opt_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0: SendKeys "{Tab}": Exit Sub
End Sub

Private Sub PopMenu_Cond_Click(Index As Integer)
    Dim objCondPars As New RPTPars
    Dim i As Integer
    
    'ִ��LoadCond��Index���Ϊ0��������i
    i = Index
    Set objCondPars = mdlPublic.RPTParsCondExec(mlngReport, Val(PopMenu_Cond(i).Tag), mobjDefPars)
    If Not objCondPars Is Nothing Then
        Call LoadCond(objCondPars)
        mint������ = Val(PopMenu_Cond(i).Tag)
        mintMenu = i
        Call UpdateMenuItemCheck(mintMenu)
    End If
End Sub

Private Sub UpdateMenuItemCheck(ByVal vCondNo As Integer)
    Dim i As Integer
    
    For i = 1 To PopMenu_Cond.count - 1
        PopMenu_Cond(i).Checked = vCondNo = i
    Next
    PopMenu_Default.Checked = vCondNo = 0
End Sub

Private Sub PopMenu_Default_Click()
    mint������ = 0
    mintMenu = 0
    Call LoadCond(mobjDefPars)
    PopMenu_Del.Enabled = False
    Call UpdateMenuItemCheck(mintMenu)
End Sub

Private Sub LoadCond(ByVal objPars As RPTPars)
    Me.Tag = "1"
    LockWindowUpdate Me.hwnd
    Call CopyPars(objPars, mobjPars)
    Call Form_Load
    cmdOK.SetFocus
    LockWindowUpdate 0
    
    PopMenu_Del.Enabled = True
End Sub

Private Sub PopMenu_Del_Click()
    If mdlPublic.RPTParsCondDel(mlngReport, mint������) Then
        mint������ = 0
        mintMenu = 0
        Call LoadCondsMenu
        Call PopMenu_Default_Click
    End If
End Sub

Private Sub PopMenu_Save_Click()
    If mdlPublic.RPTParsCondSave(mlngReport, mint������, mobjPars, mobjDefPars, Me) Then
        If mintMenu = 0 Then
            '��ȱʡ״̬�±��棬����Ϊ����������
            Call PopMenu_Cond_Click(PopMenu_Cond.count - 1)
        Else
            '������״̬�±���
            Call PopMenu_Cond_Click(mintMenu)
        End If
    End If
End Sub

Private Sub PopMenu_Saveas_Click()
    If mdlPublic.RPTParsCondSave(mlngReport, mint������, mobjPars, mobjDefPars, Me, True) Then
        If mintMenu = 0 Then
            '��ȱʡ״̬�±��棬����Ϊ����������
            Call PopMenu_Cond_Click(PopMenu_Cond.count - 1)
        Else
            '������״̬�±���
            Call PopMenu_Cond_Click(mintMenu)
        End If
    End If
End Sub

Private Sub txt_GotFocus(Index As Integer)
    SelAll txt(Index)
End Sub

Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 And txt(Index).ToolTipText <> "" Then
        If cmd(Index).Enabled And cmd(Index).Visible Then Call cmd_Click(Index)
    End If
    If txt(Index).Locked Then Exit Sub
    
    '��Ϊ����ʱ(��ѡ��)�������ֵ��Ϊ��Ϊ����ı�־
    '144=Num;112-123=F1-F12;229=��ʼ���뺺��
    If KeyCode >= 48 And KeyCode <> 144 _
        And Not (KeyCode >= 112 And KeyCode <= 123) Then
        txt(Index).Tag = ""
    End If
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txt(Index).Tag = "" And txt(Index).ToolTipText <> "" Then
            '������ƥ��
            KeyAscii = 0
            If txt(Index).Text <> "" Then
                If cmd(Index).Enabled And cmd(Index).Visible Then
                    mblnMatch = True
                    Call cmd_Click(Index)
                End If
            End If
            Exit Sub
        Else
            '���ƶ�����
            KeyAscii = 0: SendKeys "{Tab}": Exit Sub
        End If
    End If
    
    If txt(Index).Locked Then Exit Sub
    
    If InStr("~`!@#$^&"";|" & Chr(3) & Chr(22), Chr(KeyAscii)) > 0 Then KeyAscii = 0
    
    If txt(Index).ToolTipText = "" And mobjPars("_" & lbl(Index).ToolTipText).���� = 1 Then
        If InStr("-0.123456789" & Chr(8) & Chr(24) & Chr(26), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
    '��Ϊ����ʱ(��ѡ��)�������ֵ��Ϊ��Ϊ����ı�־
    '����ֻ������,������KeyDown�д���
    If KeyAscii < 0 Then txt(Index).Tag = ""
End Sub

Private Sub Form_Load()
    Dim i As Long, j As Long, k As Long
    Dim tmpPar As RPTPar, strTmp As String
    Dim lngCurH As Long, objTmp As Object
    Dim intCurTab As Integer, blnCmd As Boolean
    Dim strGroup As String, objGroup As Object
    Dim strCur As String, strPre As String
    Dim objLoad As Object, blnExist As Boolean
    Dim strBegin As String, strEnd As String
    Dim blnFlag As Boolean
    
    mblnOK = False
    mblnMatch = False
    mintBegin = -1: mintEnd = -1
    Caption = "�������� - " & mstrTitle
    mint������ = 0: mintMenu = 0
    
    'ж�ؿؼ�
    For Each objLoad In lbl
        If objLoad.Index <> 0 Then Unload objLoad
    Next
    For Each objLoad In txt
        If objLoad.Index <> 0 Then Unload objLoad
    Next
    For Each objLoad In cmd
        If objLoad.Index <> 0 Then Unload objLoad
    Next
    For Each objLoad In cbo
        If objLoad.Index <> 0 Then Unload objLoad
    Next
    For Each objLoad In dtp
        If objLoad.Index <> 0 Then Unload objLoad
    Next
    For Each objLoad In opt
        If objLoad.Index <> 0 Then Unload objLoad
    Next
    For Each objLoad In chk
        If objLoad.Index <> 0 Then Unload objLoad
    Next
    For Each objLoad In fra
        If objLoad.Index <> 0 Then Unload objLoad
    Next
    For Each objLoad In fraGroup
        If objLoad.Index <> 0 Then Unload objLoad
    Next
        
    Call LoadCondsMenu
    If Me.Tag = "" Then
        Call UpdateMenuItemCheck(0)
    End If
    
    '���������������
    i = 0: lngCurH = lbl(0).Top
    For Each tmpPar In mobjPars
        i = i + 1
        
        Load lbl(i)
        lbl(i).Caption = tmpPar.���� & "(&" & i & ")"
        lbl(i).ToolTipText = tmpPar.����
        lbl(i).Left = txt(0).Left - lbl(i).Width - 30
        lbl(i).Top = lngCurH
        lbl(i).TabIndex = intCurTab: intCurTab = intCurTab + 1
        lbl(i).Visible = True
        
        If tmpPar.ȱʡֵ = "�̶�ֵ�б�" Then
            If tmpPar.��ʽ = 0 Then '������
                Load cbo(i): Set objTmp = cbo(i)
                If tmpPar.�Ƿ����� Then objTmp.Enabled = False
                cbo(i).TabIndex = intCurTab: intCurTab = intCurTab + 1
                cbo(i).Left = cbo(0).Left: cbo(i).Top = lbl(i).Top - (cbo(i).Height - lbl(i).Height) / 2
                '���õķָ���
                For j = 0 To UBound(Split(tmpPar.ֵ�б�, "|"))
                    strTmp = Split(Split(tmpPar.ֵ�б�, "|")(j), ",")(0)
                    
                    If Left(strTmp, 1) = "��" Then
                        cbo(i).AddItem Mid(strTmp, 2)
                        If cbo(i).ListIndex = -1 Then cbo(i).ListIndex = cbo(i).NewIndex
                    Else
                        cbo(i).AddItem strTmp
                    End If
                    
                    '��������ʱReserve�����"��ʾֵ|��ֵ"
                    '�����ϴ���ʾֵ����λȱʡ��
                    If tmpPar.Reserve Like "*|*" Then
                        If Split(tmpPar.Reserve, "|")(0) = "������" Then
                            '�������ֻ�����˰�ֵ,δ������ʾֵ,�Զ�Ѱ����ʾֵ�����
                            If Split(tmpPar.Reserve, "|")(1) = Split(Split(tmpPar.ֵ�б�, "|")(j), ",")(1) Then
                                cbo(i).ListIndex = cbo(i).NewIndex
                            End If
                        Else
                            If Left(strTmp, 1) = "��" Then
                                If Split(tmpPar.Reserve, "|")(0) = Mid(strTmp, 2) Then cbo(i).ListIndex = cbo(i).NewIndex
                            Else
                                If Split(tmpPar.Reserve, "|")(0) = strTmp Then cbo(i).ListIndex = cbo(i).NewIndex
                            End If
                            
                            '�ϴ���Ϊ�����ֵ��ĳ����ֵ��ͬ,��λ
                            '��Ϊ���ѡ��ֵ�а�ֵ�����ظ�,���Դ˶οɲ�Ҫ
                            If Split(tmpPar.Reserve, "|")(0) = Split(Split(tmpPar.ֵ�б�, "|")(j), ",")(1) Then
                                cbo(i).ListIndex = cbo(i).NewIndex
                            End If
                        End If
                    End If
                Next
                cbo(i).Visible = True
            ElseIf tmpPar.��ʽ = 1 Then '��ѡ��
                Load fra(i): Set objTmp = fra(i)
                fra(i).TabIndex = intCurTab: intCurTab = intCurTab + 1
                fra(i).Left = fra(0).Left: fra(i).Top = lbl(i).Top - 50
                
                lbl(i).Visible = False
                fra(i).Caption = lbl(i).Caption
                                
                j = UBound(Split(tmpPar.ֵ�б�, "|")) + 1 '��ѡ��
                j = CInt((j / 3) + 0.4) '����
                
                fra(i).Height = fra(0).Height + (j - 1) * (opt(0).Height * 1.6) - opt(0).Height * 0.3
                
                blnExist = False
                '���õķָ���
                For j = 0 To UBound(Split(tmpPar.ֵ�б�, "|"))
                    strTmp = Split(Split(tmpPar.ֵ�б�, "|")(j), ",")(0)
                    
                    Load opt(opt.UBound + 1)
                    Set opt(opt.UBound).Container = fra(i)
                    opt(opt.UBound).TabIndex = intCurTab: intCurTab = intCurTab + 1
                    opt(opt.UBound).Tag = Split(Split(tmpPar.ֵ�б�, "|")(j), ",")(1) '��Ű�ֵ
                    If tmpPar.�Ƿ����� Then opt(opt.UBound).Enabled = False
                    
                    If InStr(",0,1,3,", "," & UBound(Split(tmpPar.ֵ�б�, "|")) & ",") > 0 Then
                        'ֻ��1,2,4����������⴦��
                        If j = 0 Or j = 1 Then 'Top
                            opt(opt.UBound).Top = opt(0).Top
                        Else
                            opt(opt.UBound).Top = opt(0).Top + opt(0).Height * 1.6
                        End If
                        If j = 0 Or j = 2 Then 'Left
                            opt(opt.UBound).Left = opt(0).Left + 150
                        Else
                            opt(opt.UBound).Left = opt(0).Left + (opt(0).Width * 1.4 + 60) + 150
                        End If
                        
                        If Left(strTmp, 1) = "��" Then
                            opt(opt.UBound).Caption = GetLenStr(Mid(strTmp, 2), opt(0).Width * 1.4 - 200, Me)
                            opt(opt.UBound).ToolTipText = Mid(strTmp, 2)
                            If Not blnExist Then opt(opt.UBound).Value = True
                        Else
                            opt(opt.UBound).Caption = GetLenStr(strTmp, opt(0).Width * 1.4 - 200, Me)
                            opt(opt.UBound).ToolTipText = strTmp
                        End If
                    Else
                        opt(opt.UBound).Top = opt(0).Top + (CInt(((j + 1) / 3) + 0.4) - 1) * (opt(0).Height * 1.6)
                        opt(opt.UBound).Left = opt(0).Left + (IIF(((j + 1) Mod 3) = 0, 3, ((j + 1) Mod 3)) - 1) * (opt(0).Width + 60)
                        
                        If Left(strTmp, 1) = "��" Then
                            opt(opt.UBound).Caption = GetLenStr(Mid(strTmp, 2), opt(0).Width - 200, Me)
                            opt(opt.UBound).ToolTipText = Mid(strTmp, 2)
                            If Not blnExist Then opt(opt.UBound).Value = True
                        Else
                            opt(opt.UBound).Caption = GetLenStr(strTmp, opt(0).Width - 200, Me)
                            opt(opt.UBound).ToolTipText = strTmp
                        End If
                    End If

                    opt(opt.UBound).Width = TextWidth(opt(opt.UBound).Caption) + 300
                    
                    '��������ʱReserve�����"��ʾֵ|��ֵ"
                    '�����ϴ�ѡ��ֵ����λȱʡ��
                    If tmpPar.Reserve Like "*|*" Then
                        If Split(tmpPar.Reserve, "|")(0) = "������" Then
                            '�������ֻ�����˰�ֵ,δ������ʾֵ,�Զ�Ѱ����ʾֵ�����
                            If Split(tmpPar.Reserve, "|")(1) = Split(Split(tmpPar.ֵ�б�, "|")(j), ",")(1) Then
                                opt(opt.UBound).Value = True: blnExist = True
                            End If
                        Else
                            If Left(strTmp, 1) = "��" Then
                                If Split(tmpPar.Reserve, "|")(0) = Mid(strTmp, 2) Then
                                    opt(opt.UBound).Value = True: blnExist = True
                                End If
                            Else
                                If Split(tmpPar.Reserve, "|")(0) = strTmp Then
                                    opt(opt.UBound).Value = True: blnExist = True
                                End If
                            End If
                        End If
                    End If
                    
                    opt(opt.UBound).Visible = True
                Next
                
                fra(i).ZOrder 1 '����������
                fra(i).Visible = True
            ElseIf tmpPar.��ʽ = 2 Then '������ѡ��
                
                lbl(i).Visible = False
                If cmdSelAll.Tag = "" Then cmdSelAll.Top = lbl(i).Top: cmdSelNone.Top = lbl(i).Top
                cmdSelAll.Visible = True: cmdSelNone.Visible = True: cmdSelAll.Tag = "1"
                Load chk(i): Set objTmp = chk(i)
                If tmpPar.�Ƿ����� Then objTmp.Enabled = False
                chk(i).Caption = lbl(i).Caption
                chk(i).TabIndex = intCurTab: intCurTab = intCurTab + 1
                chk(i).Left = chk(0).Left: chk(i).Top = lbl(i).Top - (chk(i).Height - lbl(i).Height) / 2
                chk(i).Width = TextWidth(chk(i).Caption) + 230
                If tmpPar.���� <> "" Then
                    If k > 0 Then
                        If fra(0).Width + fra(0).Left - chk(k).Left - chk(k).Width > (fra(0).Width - 1550) And chk(i).Width < (fra(0).Width - 1550) Then
                            chk(i).Left = fra(0).Left + 1550
                            blnFlag = True
                        ElseIf fra(0).Width + fra(0).Left - chk(k).Left - chk(k).Width > (fra(0).Width - 2800) And chk(i).Width < (fra(0).Width - 2800) Then
                            chk(i).Left = fra(0).Left + 2800
                            blnFlag = True
                        Else
                            chk(i).Left = fra(0).Left + 300
                        End If
                    Else
                        chk(i).Left = fra(0).Left + 300
                    End If
                    k = i
                End If
                
                If Left(Split(Split(tmpPar.ֵ�б�, "|")(0), ",")(0), 1) = "��" Then chk(i).Value = 1
                '���õķָ���
                For j = 0 To 1
                    strTmp = Split(Split(tmpPar.ֵ�б�, "|")(j), ",")(0)
                    '��������ʱReserve����ϴ���"��ʾֵ|��ֵ"
                    '�����ϴ�ѡ��ֵ����λ����ȱʡ��
                    If tmpPar.Reserve Like "*|*" Then
                        If Split(tmpPar.Reserve, "|")(0) = "������" Then
                            '�������ֻ�����˰�ֵ,δ������ʾֵ,�Զ�Ѱ����ʾֵ�����
                            If Split(tmpPar.Reserve, "|")(1) = Split(Split(tmpPar.ֵ�б�, "|")(j), ",")(1) Then
                                chk(i).Value = IIF(Left(strTmp, 1) = "��", 1, 0)
                            End If
                        Else
                            If Left(strTmp, 1) = "��" Then
                                If Split(tmpPar.Reserve, "|")(0) = Mid(strTmp, 2) Then
                                    chk(i).Value = IIF(Left(strTmp, 1) = "��", 1, 0)
                                End If
                            Else
                                If Split(tmpPar.Reserve, "|")(0) = strTmp Then
                                    chk(i).Value = IIF(Left(strTmp, 1) = "��", 1, 0)
                                End If
                            End If
                        End If
                    End If
                Next
                chk(i).Visible = True
            End If
        ElseIf tmpPar.ȱʡֵ = "ѡ�������塭" Then
            Load txt(i): Set objTmp = txt(i)
            If tmpPar.�Ƿ����� Then objTmp.Enabled = False
            txt(i).TabIndex = intCurTab: intCurTab = intCurTab + 1
            txt(i).Left = txt(0).Left: txt(i).Top = lbl(i).Top - (txt(i).Height - lbl(i).Height) / 2
            txt(i).ToolTipText = "�� F2 ��ѡ����"
            
            blnCmd = True
            If tmpPar.Reserve Like "*|*" Then
                If Split(tmpPar.Reserve, "|")(0) <> "" Then
                    strTmp = ""
                    
                    '�������ֻ�����˰�ֵ,δ������ʾֵ,�Զ�Ѱ����ʾֵ�����
                    If Split(tmpPar.Reserve, "|")(0) = "������" Then
                        If InStr(tmpPar.����, "|") > 0 Then strTmp = Split(tmpPar.����, "|")(0)
                        strTmp = SQLOwner(Replace(RemoveNote(tmpPar.��ϸSQL), "[*]", ""), strTmp)
                        Call CheckParsRela(strTmp, Nothing, tmpPar.����, True, , mobjPars)
                        strTmp = GetDefaultValue(strTmp, tmpPar.��ϸ�ֶ�, Split(tmpPar.Reserve, "|")(1) _
                            , GetDBConnectNo(tmpPar, mobjRPTDatas))
                        If strTmp <> "" Then
                            tmpPar.Reserve = Split(strTmp, "|")(0) & "|" & Split(tmpPar.Reserve, "|")(1)
                        ElseIf tmpPar.ֵ�б� Like "*|*" Then
                            If Split(tmpPar.Reserve, "|")(1) = Split(tmpPar.ֵ�б�, "|")(1) Then
                                tmpPar.Reserve = tmpPar.ֵ�б� '�붨���ȱʡ��ֵ��ͬ
                            End If
                        End If
                    End If

                    '��������ʱReserve�����"��ʾֵ|��ֵ"
                    If Split(tmpPar.Reserve, "|")(0) = "������" Then
                        'û���ҵ�����ȱʡ��ֵ��ͬ,����ʾΪ��ֵ
                        txt(i).Text = Split(tmpPar.Reserve, "|")(1)
                    Else
                        txt(i).Text = Split(tmpPar.Reserve, "|")(0)
                    End If
                    txt(i).Tag = Split(tmpPar.Reserve, "|")(1)
                    
                    '��Ȼ��ȱʡ,�����û��������ѡ�򲻿ɼ�
                    If strTmp = "" Then '����ǰ�������"������"��ʾֵʱ�Ľ��
                        If InStr(tmpPar.����, "|") > 0 Then strTmp = Split(tmpPar.����, "|")(0)
                        strTmp = SQLOwner(Replace(RemoveNote(tmpPar.��ϸSQL), "[*]", ""), strTmp)
                        Call CheckParsRela(strTmp, Nothing, tmpPar.����, True, , mobjPars)
                        strTmp = GetDefaultValue(strTmp, tmpPar.��ϸ�ֶ�, , GetDBConnectNo(tmpPar, mobjRPTDatas))
                    End If
                    
                    If strTmp <> "" Then
                        blnCmd = (CLng((Split(strTmp, "|")(2))) > 0)
                    Else
                        blnCmd = False
                    End If
                Else
                    'ʹ��ȱʡ�����ȱʡֵ
                    If tmpPar.ֵ�б� Like "*|*" Then
                        txt(i).Text = Split(tmpPar.ֵ�б�, "|")(0)
                        txt(i).Tag = Split(tmpPar.ֵ�б�, "|")(1)
                    ElseIf tmpPar.��ϸSQL <> "" Then
                        'ȡ��ϸSQL����е�һ��ֵ,���ֻ��һ��,����ѡ
                        strTmp = ""
                        If InStr(tmpPar.����, "|") > 0 Then strTmp = Split(tmpPar.����, "|")(0)
                        strTmp = SQLOwner(Replace(RemoveNote(tmpPar.��ϸSQL), "[*]", ""), strTmp)
                        Call CheckParsRela(strTmp, Nothing, tmpPar.����, True, , mobjPars)
                        strTmp = GetDefaultValue(strTmp, tmpPar.��ϸ�ֶ�, , GetDBConnectNo(tmpPar, mobjRPTDatas))
                        If strTmp <> "" Then
                            txt(i).Text = Split(strTmp, "|")(0)
                            txt(i).Tag = Split(strTmp, "|")(1)
                            If tmpPar.��ʽ = 1 Then txt(i).Tag = " IN (" & txt(i).Tag & ") "
                            blnCmd = (CLng((Split(strTmp, "|")(2))) > 1)
                        Else
                            blnCmd = False
                        End If
                    End If
                End If
            Else
                If tmpPar.ֵ�б� Like "*|*" Then
                    'ʹ��ȱʡ�����ȱʡֵ
                    txt(i).Text = Split(tmpPar.ֵ�б�, "|")(0)
                    txt(i).Tag = Split(tmpPar.ֵ�б�, "|")(1)
                    
                    '��Ȼ��ȱʡ,�����û��������ѡ�򲻿ɼ�
                    strTmp = ""
                    If InStr(tmpPar.����, "|") > 0 Then strTmp = Split(tmpPar.����, "|")(0)
                    strTmp = SQLOwner(Replace(RemoveNote(tmpPar.��ϸSQL), "[*]", ""), strTmp)
                    Call CheckParsRela(strTmp, Nothing, tmpPar.����, True, , mobjPars)
                    strTmp = GetDefaultValue(strTmp, tmpPar.��ϸ�ֶ�, , GetDBConnectNo(tmpPar, mobjRPTDatas))
                    If strTmp <> "" Then
                        blnCmd = (CLng((Split(strTmp, "|")(2))) > 0)
                    Else
                        blnCmd = False
                    End If
                ElseIf tmpPar.��ϸSQL <> "" Then
                    'ȡ��ϸSQL����е�һ��ֵ,���ֻ��һ��,����ѡ
                    strTmp = ""
                    If InStr(tmpPar.����, "|") > 0 Then strTmp = Split(tmpPar.����, "|")(0)
                    strTmp = SQLOwner(Replace(RemoveNote(tmpPar.��ϸSQL), "[*]", ""), strTmp)
                    Call CheckParsRela(strTmp, Nothing, tmpPar.����, True, , mobjPars)
                    strTmp = GetDefaultValue(strTmp, tmpPar.��ϸ�ֶ�, , GetDBConnectNo(tmpPar, mobjRPTDatas))
                    If strTmp <> "" Then
                        txt(i).Text = Split(strTmp, "|")(0)
                        txt(i).Tag = Split(strTmp, "|")(1)
                        If tmpPar.��ʽ = 1 Then txt(i).Tag = " IN (" & txt(i).Tag & ") "
                        blnCmd = (CLng((Split(strTmp, "|")(2))) > 1)
                    Else
                        blnCmd = False
                    End If
                End If
            End If
                        
            Load cmd(i)
            If tmpPar.�Ƿ����� Then cmd(i).Enabled = False
            cmd(i).Top = txt(i).Top
            cmd(i).Left = txt(i).Left + txt(i).Width + 15
            cmd(i).Height = txt(i).Height
            cmd(i).TabStop = False
            cmd(i).ZOrder
            
            txt(i).Visible = True
            cmd(i).Visible = blnCmd
            
            '�ɷ�����ƥ��
            txt(i).Locked = Not ((InStr(tmpPar.����SQL, "[*]") > 0 Or InStr(tmpPar.��ϸSQL, "[*]") > 0) And blnCmd)
        Else
            If tmpPar.���� = 2 Then
                Load dtp(i): Set objTmp = dtp(i)
                If tmpPar.�Ƿ����� Then objTmp.Enabled = False
                dtp(i).TabIndex = intCurTab: intCurTab = intCurTab + 1
                dtp(i).Left = dtp(0).Left: dtp(i).Top = lbl(i).Top - (dtp(i).Height - lbl(i).Height) / 2
                If InStr(tmpPar.ȱʡֵ, ":") > 0 Or InStr(tmpPar.ȱʡֵ, "ʱ��") > 0 Then
                    dtp(i).CustomFormat = "yyyy��MM��dd�� HH:mm:ss"
                    dtp(i).Width = 2460
                Else
                    dtp(i).CustomFormat = "yyyy��MM��dd��"
                    dtp(i).Width = 1635
                End If
                If tmpPar.ȱʡֵ <> "" Then
                    If Left(tmpPar.ȱʡֵ, 1) = "&" Then
                        dtp(i).Value = GetParVBMacro(tmpPar.ȱʡֵ)
                    Else
                        dtp(i).Value = Format(tmpPar.ȱʡֵ, dtp(i).CustomFormat)
                    End If
                Else
                    dtp(i).Value = Currentdate
                End If
                
'                'ע�����ֵ
'                If dtp(i).CustomFormat Like "*HH:mm:ss" And Left(tmpPar.ȱʡֵ, 1) <> "&" Then
'                    strTmp = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.name & mstrTitle, lbl(i).ToolTipText & "ʱ��", Format(dtp(i).Value, "HH:mm:ss"))
'                    dtp(i).Value = CDate(Format(dtp(i).Value, Left(dtp(i).CustomFormat, InStr(dtp(i).CustomFormat, "HH:mm:ss") - 1)) & strTmp)
'                End If
                
                '�Ƿ�ʼ����ʱ��(�ձ���ʱ��)
                If dtp(i).CustomFormat Like "*HH:mm:ss" Then
                    If tmpPar.���� Like "��ʼ*" Or tmpPar.���� Like "��ʼ*" Then
                        mintBegin = i
                    ElseIf tmpPar.���� Like "����*" Or tmpPar.���� Like "��ֹ*" Then
                        mintEnd = i
                    End If
                End If
                
                dtp(i).Visible = True
            Else
                Load txt(i): Set objTmp = txt(i)
                If tmpPar.�Ƿ����� Then objTmp.Enabled = False
                txt(i).Left = txt(0).Left: txt(i).Top = lbl(i).Top - (txt(i).Height - lbl(i).Height) / 2
                txt(i).TabIndex = intCurTab: intCurTab = intCurTab + 1
                txt(i).Text = tmpPar.ȱʡֵ
                txt(i).Visible = True
            End If
        End If
        If objTmp.name = "fra" Then
            lngCurH = lngCurH + objTmp.Height + 180
        Else
            If blnFlag = False Then
                lngCurH = lngCurH + txt(0).Height + 150
            End If
            blnFlag = False
        End If
        
        lbl(i).Tag = tmpPar.���� & "," & objTmp.name
        If tmpPar.ȱʡֵ = "ѡ�������塭" Then lbl(i).Tag = lbl(i).Tag & ",cmd"
    Next
    cmdOK.TabIndex = intCurTab: intCurTab = intCurTab + 1
    cmdCancel.TabIndex = cmdOK.TabIndex + 1
    
    picPar.Height = lngCurH
    picPar.Visible = Not (lbl.UBound = 0)
    
    k = 0
    '���������
    For i = 1 To lbl.UBound
        strCur = ""
        If strGroup <> CStr(Split(lbl(i).Tag, ",")(0)) And CStr(Split(lbl(i).Tag, ",")(0)) <> "" Then
            Load fraGroup(fraGroup.UBound + 1)
            Set objGroup = fraGroup(fraGroup.UBound)
            objGroup.Caption = CStr(Split(lbl(i).Tag, ",")(0))
            objGroup.Top = lbl(i).Top - 150
            objGroup.ZOrder 1
            objGroup.Visible = True
            
            Select Case CStr(Split(lbl(i).Tag, ",")(1))
                Case "txt"
                    Set objTmp = txt(i)
                Case "cbo"
                    Set objTmp = cbo(i)
                Case "dtp"
                    Set objTmp = dtp(i)
                Case "chk"
                    Set objTmp = chk(i)
                    k = i
            End Select
            
            lngCurH = 195 '��ǰTopλ��
            
            Set objTmp.Container = objGroup
            objTmp.Top = lngCurH + (300 - objTmp.Height) / 2
            If CStr(Split(lbl(i).Tag, ",")(1)) = "chk" Then
                objTmp.Left = 300
            Else
                objTmp.Left = 1250
            End If
            
            Set lbl(i).Container = objGroup
            lbl(i).Top = objTmp.Top + (objTmp.Height - lbl(i).Height) / 2
            lbl(i).Left = objTmp.Left - lbl(i).Width - 30
            lbl(i).Caption = GetLenStr(lbl(i).ToolTipText, 900, Me) & Mid(lbl(i).Caption, InStr(lbl(i).Caption, "("))
            
            If UBound(Split(lbl(i).Tag, ",")) = 2 Then
                Set cmd(i).Container = objGroup
                cmd(i).Top = objTmp.Top + 30
                cmd(i).Left = objTmp.Left + objTmp.Width - cmd(i).Width - 30
            End If

            lngCurH = lngCurH + txt(0).Height + 50 '��ǰTopλ��
        ElseIf strGroup = CStr(Split(lbl(i).Tag, ",")(0)) And CStr(Split(lbl(i).Tag, ",")(0)) <> "" Then
            strCur = "Add"
            Select Case CStr(Split(lbl(i).Tag, ",")(1))
                Case "txt"
                    Set objTmp = txt(i)
                Case "cbo"
                    Set objTmp = cbo(i)
                Case "dtp"
                    Set objTmp = dtp(i)
                Case "chk"
                    Set objTmp = chk(i)
            End Select
            
            Set objTmp.Container = objGroup
            '���Ϊchk�������ж��Ƿ�һ�������ɿؼ�
            If CStr(Split(lbl(i).Tag, ",")(1)) = "chk" Then
                If objGroup.Width - chk(k).Left - chk(k).Width >= (objGroup.Width - 1550) And chk(i).Width < (objGroup.Width - 1550) Then
                    chk(i).Left = 1550
                    blnFlag = True
                ElseIf objGroup.Width - chk(k).Left - chk(k).Width > (objGroup.Width - 2800) And chk(i).Width < (objGroup.Width - 2800) Then
                    chk(i).Left = 2800
                    blnFlag = True
                Else
                    chk(i).Left = 300
                End If
            Else
                objTmp.Left = 1250
            End If
            
            If CStr(Split(lbl(i).Tag, ",")(1)) = "chk" Then
                If objGroup.Width - chk(k).Left - chk(k).Width >= (objGroup.Width - 1550) And chk(i).Width < (objGroup.Width - 1550) Then
                    objTmp.Top = chk(k).Top
                    blnFlag = True
                ElseIf objGroup.Width - chk(k).Left - chk(k).Width > (objGroup.Width - 2800) And chk(i).Width < (objGroup.Width - 2800) Then
                    objTmp.Top = chk(k).Top
                    blnFlag = True
                Else
                    objTmp.Top = lngCurH + (300 - objTmp.Height) / 2
                End If
            Else
                objTmp.Top = lngCurH + (300 - objTmp.Height) / 2
            End If
            If CStr(Split(lbl(i).Tag, ",")(1)) = "chk" Then
                k = i
            End If
            Set lbl(i).Container = objGroup
            lbl(i).Top = objTmp.Top + (objTmp.Height - lbl(i).Height) / 2
            lbl(i).Left = objTmp.Left - lbl(i).Width - 30
            lbl(i).Caption = GetLenStr(lbl(i).ToolTipText, 900, Me) & Mid(lbl(i).Caption, InStr(lbl(i).Caption, "("))
            
            If UBound(Split(lbl(i).Tag, ",")) = 2 Then
                Set cmd(i).Container = objGroup
                cmd(i).Top = objTmp.Top + 30
                cmd(i).Left = objTmp.Left + objTmp.Width - cmd(i).Width - 30
            End If
            
            If blnFlag = False Then
                lngCurH = lngCurH + txt(0).Height + 50 '��ǰTopλ��
            End If
            
            objGroup.Height = objTmp.Top + objTmp.Height + 90  '��߶�
            
            '�ÿ����µ���������ȫ������
            For j = i + 1 To lbl.UBound
                If Split(lbl(j).Tag, ",")(0) <> "fra" Then
                    lbl(j).Top = lbl(j).Top + 60
                    Select Case CStr(Split(lbl(j).Tag, ",")(1))
                        Case "txt"
                            txt(j).Top = txt(j).Top + 60
                        Case "cbo"
                            cbo(j).Top = cbo(j).Top + 60
                        Case "dtp"
                            dtp(j).Top = dtp(j).Top + 60
                        Case "chk"
                            chk(j).Top = chk(j).Top + 60
                    End Select
                    If UBound(Split(lbl(j).Tag, ",")) = 2 Then
                        cmd(j).Top = cmd(j).Top + 60
                    End If
                End If
            Next
        End If
        If strPre = "Add" And strCur = "" Then
            picPar.Height = picPar.Height + 60
        End If
        strPre = strCur
        strGroup = CStr(Split(lbl(i).Tag, ",")(0))
        blnFlag = False
    Next
    
    'û�в����鵫�ж��ѡ��ʱ,��ÿ����
    If fraGroup.UBound = 0 And fra.UBound > 0 Then
        For Each objTmp In fra
            objTmp.Left = txt(0).Left - 400
        Next
    End If
            
    Me.Height = picInfo.Height + IIF(lbl.UBound = 0, 0, picPar.Height) + picCmd.Height + 380
    
    '����ʼ����ʱ��
    If mintBegin <> -1 And mintEnd <> -1 Then
        chkAutoSave.Visible = True
        chkAutoSave.Value = Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.name & mstrTitle, "AutoSave", 0))
        
        '�û����������������ڻָ�ȱʡֵʱ������
        If Not (mblnReset Or Visible) Then
            strBegin = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.name & mstrTitle, "BeginTime", "")
            strEnd = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.name & mstrTitle, "EndTime", "")
                    
            '����ϴ�ѡ���˱���
            If chkAutoSave.Value = 1 And IsDate(strBegin) And IsDate(strEnd) Then
                '���ϴεĽ���ʱ����Ϊ���ο�ʼʱ��(+1s)
                dtp(mintBegin).Value = Format(DateAdd("s", 1, CDate(strEnd)), dtp(mintBegin).CustomFormat)
            End If
        End If
    Else
        chkAutoSave.Visible = False
        cmdDefault.Left = chkAutoSave.Left
    End If
End Sub

Private Sub txt_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        lngTXTProc = GetWindowLong(txt(Index).hwnd, GWL_WNDPROC)
        Call SetWindowLong(txt(Index).hwnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txt_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Call SetWindowLong(txt(Index).hwnd, GWL_WNDPROC, lngTXTProc)
    End If
End Sub

Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
    If txt(Index).Tag = "" And txt(Index).ToolTipText <> "" Then
        'ǿ������ƥ��
        If txt(Index).Text <> "" Then
            If cmd(Index).Enabled And cmd(Index).Visible Then
                mblnMatch = True
                Call cmd_Click(Index)
            End If
            Cancel = True
        End If
    End If
End Sub

Private Sub LoadCondsMenu()
    Dim strSQL As String
    Dim i As Integer
    Dim rsPara As New ADODB.Recordset
    Dim blnRetry As Boolean
    
    If mlngReport = 0 Then Exit Sub
    
    On Error GoTo hErr
    
    'ɾ�������˵�
    For i = 1 To PopMenu_Cond.count - 1
        Unload PopMenu_Cond(i)
    Next
    
    '��װ���û����趨��ȱʡ����
    blnRetry = True
    strSQL = "Select Distinct ������,�������� From zlRptConds Where ����ID=[1] Order by ������"
    Set rsPara = OpenSQLRecord(strSQL, Me.Caption, mlngReport)
    blnRetry = False
    
    With rsPara
        If .RecordCount = 0 Then
            Me.split0.Visible = False
            If mlngReport = 0 Then
                PopMenu_Save.Enabled = False
                PopMenu_Saveas.Enabled = False
                Me.split1.Enabled = False
            End If
        Else
            Me.split0.Visible = True
            PopMenu_Save.Enabled = True
            PopMenu_Saveas.Enabled = True
            Me.split1.Enabled = True
            Do While Not .EOF
                i = .AbsolutePosition
                Load PopMenu_Cond(i)
                PopMenu_Cond(i).Caption = !�������� & "(&" & i & ")"
                PopMenu_Cond(i).Visible = True
                PopMenu_Cond(i).Tag = !������
                .MoveNext
            Loop
        End If
    End With
    
    Exit Sub
    
hErr:
    If blnRetry Then
        If ErrCenter = 1 Then Resume
    Else
        Call ErrCenter
    End If
End Sub
