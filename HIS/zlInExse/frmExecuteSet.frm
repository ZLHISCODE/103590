VERSION 5.00
Begin VB.Form frmExecuteSet 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��������"
   ClientHeight    =   3765
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6555
   Icon            =   "frmExecuteSet.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   6555
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdRegPrint 
      Caption         =   "ִ�еǼǵ���ӡ����"
      Height          =   350
      Left            =   2295
      TabIndex        =   21
      Top             =   2595
      Width           =   1860
   End
   Begin VB.ListBox lst��� 
      Columns         =   2
      ForeColor       =   &H80000012&
      Height          =   2580
      IMEMode         =   3  'DISABLE
      ItemData        =   "frmExecuteSet.frx":058A
      Left            =   165
      List            =   "frmExecuteSet.frx":058C
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   360
      Width           =   2040
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   5235
      TabIndex        =   2
      Top             =   3270
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   4080
      TabIndex        =   1
      Top             =   3270
      Width           =   1100
   End
   Begin VB.Frame Frame2 
      Height          =   120
      Left            =   120
      TabIndex        =   4
      Top             =   3030
      Width           =   6375
   End
   Begin VB.Frame Frame1 
      Caption         =   "������Դ"
      Height          =   2235
      Left            =   2295
      TabIndex        =   3
      Top             =   270
      Width           =   4125
      Begin VB.Frame fra��Դ 
         Height          =   1530
         Index           =   2
         Left            =   2760
         TabIndex        =   17
         Top             =   600
         Width           =   1260
         Begin VB.OptionButton opt2 
            Caption         =   "δ���"
            Height          =   180
            Index           =   0
            Left            =   120
            TabIndex        =   20
            Top             =   240
            Width           =   1020
         End
         Begin VB.OptionButton opt2 
            Caption         =   "�����"
            Height          =   180
            Index           =   1
            Left            =   120
            TabIndex        =   19
            Top             =   525
            Width           =   1020
         End
         Begin VB.OptionButton opt2 
            Caption         =   "���е���"
            Height          =   180
            Index           =   2
            Left            =   120
            TabIndex        =   18
            Top             =   795
            Width           =   1020
         End
      End
      Begin VB.CheckBox chk��Դ 
         Caption         =   "���"
         Height          =   195
         Index           =   2
         Left            =   2760
         TabIndex        =   16
         Top             =   360
         Value           =   1  'Checked
         Width           =   660
      End
      Begin VB.CheckBox chk��Դ 
         Caption         =   "סԺ"
         Height          =   195
         Index           =   1
         Left            =   1440
         TabIndex        =   15
         Top             =   360
         Value           =   1  'Checked
         Width           =   660
      End
      Begin VB.CheckBox chk��Դ 
         Caption         =   "����"
         Height          =   195
         Index           =   0
         Left            =   165
         TabIndex        =   14
         Top             =   360
         Value           =   1  'Checked
         Width           =   660
      End
      Begin VB.Frame fra��Դ 
         Height          =   1530
         Index           =   1
         Left            =   1440
         TabIndex        =   10
         Top             =   600
         Width           =   1260
         Begin VB.OptionButton opt1 
            Caption         =   "���е���"
            Height          =   180
            Index           =   2
            Left            =   120
            TabIndex        =   13
            Top             =   795
            Width           =   1020
         End
         Begin VB.OptionButton opt1 
            Caption         =   "�����"
            Height          =   180
            Index           =   1
            Left            =   120
            TabIndex        =   12
            Top             =   525
            Width           =   1020
         End
         Begin VB.OptionButton opt1 
            Caption         =   "δ���"
            Height          =   180
            Index           =   0
            Left            =   120
            TabIndex        =   11
            Top             =   240
            Width           =   1020
         End
      End
      Begin VB.Frame fra��Դ 
         Height          =   1530
         Index           =   0
         Left            =   165
         TabIndex        =   6
         Top             =   600
         Width           =   1260
         Begin VB.OptionButton opt0 
            Caption         =   "���е���"
            Height          =   180
            Index           =   2
            Left            =   120
            TabIndex        =   9
            Top             =   795
            Width           =   1020
         End
         Begin VB.OptionButton opt0 
            Caption         =   "���շ�"
            Height          =   180
            Index           =   1
            Left            =   120
            TabIndex        =   8
            Top             =   525
            Width           =   1020
         End
         Begin VB.OptionButton opt0 
            Caption         =   "δ�շ�"
            Height          =   180
            Index           =   0
            Left            =   120
            TabIndex        =   7
            Top             =   240
            Width           =   1020
         End
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��Ŀ���"
      Height          =   180
      Left            =   180
      TabIndex        =   5
      Top             =   90
      Width           =   720
   End
End
Attribute VB_Name = "frmExecuteSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mblnOk As Boolean
Public mstrPrivs As String
Public mlngModul As Long

Private Sub chk��Դ_Click(Index As Integer)
    If chk��Դ(0).Value = 0 And chk��Դ(1).Value = 0 And chk��Դ(2).Value = 0 Then
        chk��Դ((Index + 1) Mod 3).Value = 1
    End If
    fra��Դ(Index).Enabled = chk��Դ(Index).Value = 1
    Call SetOptionState

End Sub
Private Sub SetOptionState()
    Dim i As Integer
    
    For i = 0 To 2
        opt0(i).Enabled = fra��Դ(0).Enabled
        opt1(i).Enabled = fra��Դ(1).Enabled
        opt2(i).Enabled = fra��Դ(2).Enabled
    Next
End Sub
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim strTmp As String, i As Integer, j As Integer
    Dim blnHavePrivs As Boolean
    
    blnHavePrivs = InStr(1, mstrPrivs, ";��������;") > 0
    
    For i = 0 To lst���.ListCount - 1
        If lst���.Selected(i) Then
            strTmp = strTmp & ",'" & Chr(lst���.ItemData(i)) & "'"
        End If
    Next
    
    strTmp = Mid(strTmp, 2)
    If strTmp = "" Then
        MsgBox "������ѡ��һ�����", vbInformation, gstrSysName
        lst���.SetFocus: Exit Sub
    End If
    If UBound(Split(strTmp, ",")) + 1 = lst���.ListCount Then strTmp = ""
    
    zlDatabase.SetPara "ҽ��ִ�����", strTmp, glngSys, mlngModul, blnHavePrivs
    
    strTmp = IIf(chk��Դ(0).Value = 1, "1", "0") & IIf(chk��Դ(1).Value = 1, "1", "0") & IIf(chk��Դ(2).Value = 1, "1", "0")
    zlDatabase.SetPara "ҽ��������Դ", strTmp, glngSys, mlngModul, blnHavePrivs
    For j = 0 To 2
        If chk��Դ(j).Value = 1 Then
            strTmp = ""
            For i = 0 To 2
                If j = 0 Then
                    If opt0(i).Value = True Then strTmp = i: Exit For
                ElseIf j = 1 Then
                    If opt1(i).Value = True Then strTmp = i: Exit For
                Else
                    If opt2(i).Value = True Then strTmp = i: Exit For
                End If
            Next
            If strTmp = "" Then strTmp = "2"
            zlDatabase.SetPara Choose(j + 1, "ҽ�����ﵥ������", "ҽ��סԺ��������", "ҽ����쵥������"), strTmp, glngSys, mlngModul, blnHavePrivs
        End If
    Next
    
    Call InitLocPar(mlngModul)
    
    mblnOk = True
    Unload Me
End Sub

Private Sub cmdRegPrint_Click()
    Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1142", Me)
End Sub

Private Sub Form_Load()
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, str��� As String, i As Long, blnParSet As Boolean
    
    mblnOk = False
    blnParSet = InStr(mstrPrivs, ";��������;") > 0
    
    str��� = zlDatabase.GetPara("ҽ��������Դ", glngSys, mlngModul, "111", Array(chk��Դ(0), chk��Դ(1), chk��Դ(2)), blnParSet)
    '���������
    If Len(str���) = 1 Then
        If str��� = "0" Then
            str��� = "111"
        ElseIf str��� = "1" Then
            str��� = "101"
        Else
            str��� = "010"
        End If
    End If
    
    chk��Դ(0).Value = Val(Mid(str���, 1, 1))
    chk��Դ(1).Value = Val(Mid(str���, 2, 1))
    chk��Դ(2).Value = Val(Mid(str���, 3, 1))
    
    i = Val(zlDatabase.GetPara("ҽ�����ﵥ������", glngSys, mlngModul, 2, Array(opt0(0), opt0(1), opt0(2)), blnParSet))
    opt0(i).Value = True
    i = Val(zlDatabase.GetPara("ҽ��סԺ��������", glngSys, mlngModul, 2, Array(opt1(0), opt1(1), opt1(2)), blnParSet))
    opt1(i).Value = True
    i = Val(zlDatabase.GetPara("ҽ����쵥������", glngSys, mlngModul, 2, Array(opt2(0), opt2(1), opt2(2)), blnParSet))
    opt2(i).Value = True
    
    
    fra��Դ(0).Enabled = chk��Դ(0).Value = 1
    fra��Դ(1).Enabled = chk��Դ(1).Value = 1
    fra��Դ(2).Enabled = chk��Դ(2).Value = 1
    
    lst���.Clear
    str��� = zlDatabase.GetPara("ҽ��ִ�����", glngSys, mlngModul, "", Array(lst���), blnParSet)
    Err = 0: On Error GoTo errH:
    strSQL = "Select ����,����,����,�̶�,��� From �շ���Ŀ��� Where ���� Not IN('1','5','6','7','J') Order by ����"
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
    For i = 1 To rsTmp.RecordCount
        lst���.AddItem rsTmp!����
        lst���.ItemData(lst���.NewIndex) = Asc(rsTmp!����)
        If str��� = "" Then
            lst���.Selected(lst���.NewIndex) = True
        Else
            If InStr(str���, "'" & rsTmp!���� & "'") > 0 Then
                lst���.Selected(lst���.NewIndex) = True
            End If
        End If
        rsTmp.MoveNext
    Next
    lst���.ListIndex = 0
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
