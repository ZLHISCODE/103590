VERSION 5.00
Begin VB.Form frm������Ŀ���� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ҽ����Ŀ����"
   ClientHeight    =   3300
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6225
   Icon            =   "frm������Ŀ����.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3300
   ScaleWidth      =   6225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.Frame fraƥ�� 
      Caption         =   "ƥ�䷽ʽ"
      Height          =   1005
      Left            =   2820
      TabIndex        =   15
      Top             =   1530
      Width           =   1845
      Begin VB.OptionButton optMatch 
         Caption         =   "����ƥ��"
         Height          =   180
         Index           =   0
         Left            =   150
         TabIndex        =   16
         Top             =   330
         Value           =   -1  'True
         Width           =   1365
      End
      Begin VB.OptionButton optMatch 
         Caption         =   "����ƥ��"
         Height          =   180
         Index           =   1
         Left            =   150
         TabIndex        =   17
         Top             =   660
         Width           =   1365
      End
   End
   Begin VB.Frame fra��Χ 
      Caption         =   "���ҷ�Χ"
      Height          =   1005
      Left            =   120
      TabIndex        =   12
      Top             =   1530
      Width           =   2235
      Begin VB.OptionButton optClass 
         Caption         =   "�����շ����"
         Height          =   180
         Index           =   1
         Left            =   150
         TabIndex        =   14
         Top             =   660
         Width           =   1875
      End
      Begin VB.OptionButton optClass 
         Caption         =   "��ǰ���"
         Height          =   180
         Index           =   0
         Left            =   150
         TabIndex        =   13
         Top             =   330
         Value           =   -1  'True
         Width           =   1875
      End
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   4950
      TabIndex        =   11
      Top             =   1620
      Width           =   1100
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "�˳�(&E)"
      Height          =   350
      Left            =   4950
      TabIndex        =   10
      Top             =   780
      Width           =   1100
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "��λ(&L)"
      Height          =   350
      Left            =   4950
      TabIndex        =   9
      Top             =   210
      Width           =   1100
   End
   Begin VB.Frame fra���� 
      Caption         =   "��������"
      Height          =   1305
      Left            =   90
      TabIndex        =   18
      Top             =   120
      Width           =   4560
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   3
         Left            =   3180
         MaxLength       =   255
         TabIndex        =   7
         Top             =   750
         Width           =   1185
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   0
         Left            =   870
         MaxLength       =   255
         TabIndex        =   1
         Top             =   330
         Width           =   1035
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   1
         Left            =   3180
         MaxLength       =   255
         TabIndex        =   3
         Top             =   330
         Width           =   1185
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   2
         Left            =   870
         MaxLength       =   255
         TabIndex        =   5
         Top             =   750
         Width           =   1035
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "ҽ������(&Y)"
         Height          =   180
         Index           =   3
         Left            =   2130
         TabIndex        =   6
         Top             =   810
         Width           =   990
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "����(&C)"
         Height          =   180
         Index           =   0
         Left            =   180
         TabIndex        =   0
         Top             =   390
         Width           =   630
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "����(&S)"
         Height          =   180
         Index           =   2
         Left            =   180
         TabIndex        =   4
         Top             =   810
         Width           =   630
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "����(&N)"
         Height          =   180
         Index           =   1
         Left            =   2490
         TabIndex        =   2
         Top             =   390
         Width           =   630
      End
   End
   Begin VB.Label lbl��� 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " �������������"
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   2670
      Width           =   5955
   End
End
Attribute VB_Name = "frm������Ŀ����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mrsFind As New ADODB.Recordset
Dim mint���� As Integer
Dim mblnHIS10 As Boolean

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
   ShowHelp App.ProductName, Me.hwnd, Me.Name
End Sub

Private Sub Form_Load()
    Dim nod As Node
    'ȡ����ǰ���
    mblnHIS10 = IsZLHIS10
    Set nod = frm������Ŀ.tvwMain_S.SelectedItem
    With frm������Ŀ.cmb����
        mint���� = .ItemData(.ListIndex)
    End With
    
    Do Until nod.Parent Is Nothing
        Set nod = nod.Parent
    Loop
    optClass(0).Caption = nod.Text
    optClass(0).Tag = Mid(nod.Key, 2, 1)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mrsFind.State = 1 Then mrsFind.Close
    Set mrsFind = Nothing
End Sub

Private Sub cmdFind_Click()
    Dim strҽ��֧����Ŀ As String
    
    If mrsFind.State = 1 Then
        If Not mrsFind.EOF Then mrsFind.MoveNext
        Call LocateItem
        Exit Sub
    End If
    If IsValid = False Then Exit Sub
    gstrSQL = ""
    If txtEdit(0).Text <> "" Then
        gstrSQL = "upper(A.����) like '" & IIf(optMatch(1).Value = True, "%", "") & UCase(txtEdit(0).Text) & "%' or "
    End If
    If txtEdit(1).Text <> "" Then
        gstrSQL = gstrSQL & "upper(B.����) like '" & IIf(optMatch(1).Value = True, "%", "") & UCase(txtEdit(1).Text) & "%' or "
    End If
    If txtEdit(2).Text <> "" Then
        gstrSQL = gstrSQL & "upper(B.����) like '" & IIf(optMatch(1).Value = True, "%", "") & UCase(txtEdit(2).Text) & "%' or "
    End If
    If txtEdit(3).Text <> "" Then
        strҽ��֧����Ŀ = "(select �շ�ϸĿID,���� from ����֧����Ŀ where upper(��Ŀ����) like '" & IIf(optMatch(1).Value = True, "%", "") & UCase(txtEdit(3).Text) & "%')"
    Else
        strҽ��֧����Ŀ = "����֧����Ŀ"
    End If
    If gstrSQL <> "" Or txtEdit(3).Text <> "" Then
        If gstrSQL <> "" Then
            gstrSQL = " and (" & Mid(gstrSQL, 1, Len(gstrSQL) - 4) & ") "
        End If
    Else
        MsgBox "���������������", vbInformation, gstrSysName
        txtEdit(0).SetFocus
        Exit Sub
    End If
    
    gstrSQL = "select distinct A.���,A.�ϼ�ID,A.ID,A.���� " & _
               " from �շ�ϸĿ A,�շѱ��� B," & strҽ��֧����Ŀ & " C " & _
               " where A.ID =B.�շ�ϸĿID and A.ĩ��=1 " & gstrSQL & _
               IIf(optClass(1).Value = True, "", "and A.���='" & optClass(0).Tag & "'") & _
                " and (A.����ʱ�� is null or A.����ʱ��=to_date('3000-01-01','yyyy-mm-dd')) " & _
                " and A.ID=C.�շ�ϸĿID" & IIf(txtEdit(3).Text <> "", "", "(+)") & " and C.����(+)=" & mint����
    Set mrsFind = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    
    Call LocateItem
End Sub

Private Sub LocateItem()
    Dim rsTemp As New ADODB.Recordset
    Dim lngID As Long
    Dim lngCount As Long
    Dim str���� As String
    
    If mrsFind.RecordCount = 0 Then
        lbl���.Caption = " û���ҵ����ʵ��շ�ϸĿ"
        Beep
        Exit Sub
    End If
    If mrsFind.EOF = True Then
        lbl���.Caption = " �Ѿ���λ�������ҵ��շ�ϸĿ����������������"
        Beep
        Exit Sub
    End If
    lbl���.Caption = "           �ҵ�" & mrsFind.RecordCount & "������Ҫ����շ�ϸĿ��" & vbCrLf & "��ǰ�ǵ�" & mrsFind.AbsolutePosition & "��������Ϊ��" & mrsFind("����")
    
    With frm������Ŀ.tvwMain_S
        lngID = mrsFind("ID")
        If mrsFind!��� = "4" And mblnHIS10 Then
            gstrSQL = "Select B.����ID " & _
                      " From �շ���ĿĿ¼ A, ������ĿĿ¼ B, �������� C " & _
                      " Where A.ID = C.����ID " & _
                      " And B.ID = C.����ID and A.ID=[1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngID)
            If rsTemp.EOF Then
                Exit Sub
            End If
            
            If IsNull(rsTemp!����ID) Then
                .Nodes("G4").Selected = True
            Else
                .Nodes("G4" & rsTemp("����ID")).Selected = True
            End If
        ElseIf mrsFind!��� = "K" And mblnHIS10 Then
            gstrSQL = "Select B.����ID " & _
                      " From ������ĿĿ¼ B, ѪҺ��� C " & _
                      " Where C.Ʒ��ID=B.ID and C.���ID=[1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngID)
            If rsTemp.EOF Then
                Exit Sub
            End If
            
            If IsNull(rsTemp!����ID) Then
                .Nodes("X8").Selected = True
            Else
                .Nodes("X8" & rsTemp("����ID")).Selected = True
            End If
        ElseIf mrsFind("���") = "5" Or mrsFind("���") = "6" Or mrsFind("���") = "7" Then
            '����ϸĿ��ҩƷ,�䶨λҪ����һЩ
            gstrSQL = "select B.���ʷ���,B.��;����ID from ҩƷĿ¼ A,ҩƷ��Ϣ B " & _
                      " Where A.ҩ��ID = B.ҩ��ID And A.ҩƷID =[1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngID)
            
            If rsTemp.EOF Then
                Exit Sub
            End If
            
            Select Case rsTemp("���ʷ���")
                Case "�г�ҩ"
                    str���� = "E6"
                Case "�в�ҩ"
                    str���� = "F7"
                Case Else
                    str���� = "D5"
            End Select
                    
            If IsNull(rsTemp("��;����ID")) Then
                .Nodes(str����).Selected = True
            Else
                .Nodes(str���� & rsTemp("��;����ID")).Selected = True
            End If
            
        Else
            If mblnHIS10 Then
                gstrSQL = " Select ID,����ID From �շ���ĿĿ¼ Where ID=[1]"
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngID)
                
                If rsTemp.EOF Then Exit Sub
                
                .Nodes("CA" & rsTemp("����ID")).Selected = True
            Else
                If IsNull(mrsFind("�ϼ�ID")) Then
                    .Nodes("R" & mrsFind("���")).Selected = True
                Else
                    .Nodes("C" & mrsFind("���") & mrsFind("�ϼ�ID")).Selected = True
                End If
            End If
        End If
        .SelectedItem.EnsureVisible
    End With
    frm������Ŀ.FillSum
        
    lngID = mrsFind("ID")
    With frm������Ŀ.mshSum_S
        For lngCount = 1 To .Rows - 1
            If .RowData(lngCount) = lngID Then
                .Row = lngCount
                .msfObj.TopRow = lngCount
                Exit Sub
            End If
        Next
    End With
    MsgBox "�շ�ϸĿ��" & mrsFind("����") & "���ļ۸�δ���ã����Ѿ������ˡ�", vbInformation, gstrSysName
End Sub

Private Function IsValid() As Boolean
'����:���������йطѱ�������Ƿ���Ч
'����:
'����ֵ:��Ч����True,����ΪFalse
    Dim i As Integer
    Dim strTemp As String
    For i = txtEdit.LBound To txtEdit.UBound
        strTemp = Trim(txtEdit(i).Text)
        If InStr(strTemp, "'") > 0 Then
            MsgBox "���������ݺ��зǷ��ַ���", vbExclamation, gstrSysName
            txtEdit(i).SetFocus
            zlControl.TxtSelAll txtEdit(i)
            Exit Function
        End If
    Next
    IsValid = True
End Function

Private Sub optClass_Click(Index As Integer)
    If mrsFind.State = 1 Then mrsFind.Close
End Sub

Private Sub optClass_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
          zlCommFun.PressKey (vbKeyTab)
    End If
End Sub

Private Sub optMatch_Click(Index As Integer)
    If mrsFind.State = 1 Then mrsFind.Close
End Sub

Private Sub optMatch_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
          zlCommFun.PressKey (vbKeyTab)
    End If
End Sub

Private Sub txtEdit_Change(Index As Integer)
    If mrsFind.State = 1 Then mrsFind.Close
    lbl���.Caption = "  �����Ѹı䣬�����¶�λ"
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
          zlCommFun.PressKey (vbKeyTab)
    End If
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    If Index = 1 Then
        zlCommFun.OpenIme True
    End If
    zlControl.TxtSelAll txtEdit(Index)
End Sub

Private Sub txtEdit_LostFocus(Index As Integer)
    zlCommFun.OpenIme False
End Sub
