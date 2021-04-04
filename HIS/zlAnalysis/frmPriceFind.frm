VERSION 5.00
Begin VB.Form frmPriceFind 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��Ŀ����"
   ClientHeight    =   2940
   ClientLeft      =   45
   ClientTop       =   930
   ClientWidth     =   6225
   Icon            =   "frmPriceFind.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   6225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame fraƥ�� 
      Caption         =   "ƥ�䷽ʽ"
      Height          =   630
      Left            =   105
      TabIndex        =   10
      Top             =   1650
      Width           =   4560
      Begin VB.OptionButton optMatch 
         Caption         =   "����ƥ��"
         Height          =   180
         Index           =   0
         Left            =   825
         TabIndex        =   11
         Top             =   270
         Value           =   -1  'True
         Width           =   1170
      End
      Begin VB.OptionButton optMatch 
         Caption         =   "����ƥ��"
         Height          =   180
         Index           =   1
         Left            =   2355
         TabIndex        =   12
         Top             =   270
         Width           =   1155
      End
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   4950
      TabIndex        =   4
      Top             =   1230
      Width           =   1100
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "�˳�(&E)"
      Height          =   350
      Left            =   4950
      TabIndex        =   3
      Top             =   495
      Width           =   1100
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "��λ(&L)"
      Height          =   350
      Left            =   4950
      TabIndex        =   2
      Top             =   120
      Width           =   1100
   End
   Begin VB.Frame fra���� 
      Caption         =   "��������"
      Height          =   1455
      Left            =   90
      TabIndex        =   13
      Top             =   120
      Width           =   4560
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   0
         Left            =   885
         MaxLength       =   255
         TabIndex        =   6
         Top             =   240
         Width           =   3525
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   1
         Left            =   870
         MaxLength       =   255
         TabIndex        =   8
         Top             =   630
         Width           =   3525
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   2
         Left            =   870
         MaxLength       =   255
         TabIndex        =   1
         Top             =   1020
         Width           =   3525
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "����(&C)"
         Height          =   180
         Index           =   0
         Left            =   180
         TabIndex        =   5
         Top             =   300
         Width           =   630
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "����(&S)"
         Height          =   180
         Index           =   2
         Left            =   180
         TabIndex        =   0
         Top             =   1080
         Width           =   630
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "����(&N)"
         Height          =   180
         Index           =   1
         Left            =   180
         TabIndex        =   7
         Top             =   690
         Width           =   630
      End
   End
   Begin VB.Label lbl��� 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " �������������"
      ForeColor       =   &H8000000D&
      Height          =   510
      Left            =   120
      TabIndex        =   9
      Top             =   2355
      Width           =   5925
   End
End
Attribute VB_Name = "frmPriceFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mrsFind As New ADODB.Recordset


Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hWnd, Me.Name)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mrsFind.State = 1 Then mrsFind.Close
    Set mrsFind = Nothing
End Sub

Private Sub cmdFind_Click()
    If mrsFind.State = 1 Then
        If Not mrsFind.EOF Then mrsFind.MoveNext
        Call LocateItem
        Exit Sub
    End If
    If IsValid = False Then Exit Sub
    gstrSQL = ""
    If txtEdit(0).Text <> "" Then
        gstrSQL = "A.���� like '" & IIf(optMatch(1).Value = True, "%", "") & UCase(txtEdit(0).Text) & "%' or "
    End If
    If txtEdit(1).Text <> "" Then
        gstrSQL = gstrSQL & "B.���� like '" & IIf(optMatch(1).Value = True, "%", "") & txtEdit(1).Text & "%' or "
    End If
    If txtEdit(2).Text <> "" Then
        gstrSQL = gstrSQL & "B.���� like '" & IIf(optMatch(1).Value = True, "%", "") & UCase(txtEdit(2).Text) & "%' or "
    End If
    If gstrSQL <> "" Then
        gstrSQL = Mid(gstrSQL, 1, Len(gstrSQL) - 4)
    Else
        MsgBox "���������������", vbExclamation, gstrSysName
        txtEdit(0).SetFocus
        Exit Sub
    End If
    
    Select Case Mid(frmPriceQuery.tvwMain_S.SelectedItem.Key, 2, 1)
    Case 0
        gstrSQL = "(" & gstrSQL & ") And A.��� not in ('4','5','6','7')"
    Case 1
        gstrSQL = "(" & gstrSQL & ") And A.���='5'"
    Case 2
        gstrSQL = "(" & gstrSQL & ") And A.���='6'"
    Case 3
        gstrSQL = "(" & gstrSQL & ") And A.���='7'"
    Case 7
        gstrSQL = "(" & gstrSQL & ") And A.���='4'"
    End Select
    
    gstrSQL = "select distinct A.����ID,A.ID,A.���� " & _
            " from �շ���ĿĿ¼ A,�շ���Ŀ���� B  " & _
            " where A.ID =B.�շ�ϸĿID(+) And " & gstrSQL & _
            " and (A.����ʱ�� is null or A.����ʱ��=to_date('3000-01-01','yyyy-mm-dd'))" & _
            IIf(frmPriceQuery.mnuViewShowDynamic.Checked, "", " and A.�Ƿ���=0")
    Call OpenRecordset(mrsFind, Me.Caption)
    
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
    lbl���.Caption = "�ҵ�" & mrsFind.RecordCount & "������Ҫ����շ�ϸĿ��" & vbCrLf & "��ǰ�ǵ�" & mrsFind.AbsolutePosition & "��������Ϊ��" & mrsFind("����")
    
    With frmPriceQuery.tvwMain_S
        lngID = mrsFind("ID")
        Select Case Mid(.SelectedItem.Key, 2, 1)
        Case 0
            .Nodes(Mid(.SelectedItem.Key, 1, 2) & mrsFind("����ID")).Selected = True
        Case 1, 2, 3
            gstrSQL = "Select Z.����id From ҩƷ��� T, ������ĿĿ¼ Z Where T.ҩ��id = Z.Id and T.ҩƷID =" & lngID
            Call OpenRecordset(rsTemp, Me.Caption)
            If rsTemp.EOF Then
                Exit Sub
            End If
            .Nodes(Mid(.SelectedItem.Key, 1, 2) & rsTemp("����ID")).Selected = True
        Case 4
            '������������ʵ�����
            Exit Sub
        End Select
        .SelectedItem.EnsureVisible
    End With
    frmPriceQuery.FillSum
        
    lngID = mrsFind("ID")
    With frmPriceQuery.mshSum
        For lngCount = 1 To .Rows - 1
            If .RowData(lngCount) = lngID Then
                .Row = lngCount
                .TopRow = lngCount
                Exit Sub
            End If
        Next
    End With
    MsgBox "��" & mrsFind("����") & "���ļ۸�δ���ã����Ѿ������ˡ�", vbInformation, gstrSysName
End Sub

Private Function IsValid() As Boolean
'����:���������йطѱ�������Ƿ���Ч
'����:
'����ֵ:��Ч����True,����ΪFalse
    Dim i As Integer
    Dim strTemp As String
    For i = 0 To 2
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
          SendKeys "{TAB}"
    End If
End Sub

Private Sub optMatch_Click(Index As Integer)
    If mrsFind.State = 1 Then mrsFind.Close
End Sub

Private Sub optMatch_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
          SendKeys "{TAB}"
    End If
End Sub

Private Sub txtEdit_Change(Index As Integer)
    If mrsFind.State = 1 Then mrsFind.Close
    lbl���.Caption = "  �����Ѹı䣬�����¶�λ"
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
          SendKeys "{TAB}"
    End If
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    '10.35.70�������û����뷨�ᵼ�´����˳����ڣ�ԭ��δ֪����ʱ����
'    If Index = 1 Then
'        zlCommFun.OpenIme True
'    End If
    zlControl.TxtSelAll txtEdit(Index)
End Sub

Private Sub txtEdit_LostFocus(Index As Integer)
'    zlCommFun.OpenIme False
End Sub
