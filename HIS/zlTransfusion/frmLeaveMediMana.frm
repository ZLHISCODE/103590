VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmLeaveMediMana 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�ݴ�ҩƷ����"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9450
   Icon            =   "frmLeaveMediMana.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   9450
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmdCancle 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   8055
      TabIndex        =   2
      Top             =   5190
      Width           =   1100
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   6765
      TabIndex        =   1
      Top             =   5190
      Width           =   1100
   End
   Begin VB.TextBox txtMain 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "gg yyyy""��"" MM""��"" dd""��"""
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2052
         SubFormatType   =   3
      EndProperty
      Height          =   300
      Index           =   9
      Left            =   6630
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   4785
      Width           =   2535
   End
   Begin VB.TextBox txtMain 
      Height          =   300
      Index           =   8
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   4785
      Width           =   1470
   End
   Begin VB.TextBox txtMain 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """��""#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2052
         SubFormatType   =   2
      EndProperty
      Height          =   300
      Index           =   7
      Left            =   750
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   4785
      Width           =   1815
   End
   Begin VB.TextBox txtMain 
      Height          =   300
      Index           =   6
      Left            =   750
      MaxLength       =   200
      TabIndex        =   0
      Top             =   4440
      Width           =   8400
   End
   Begin VB.TextBox txtMain 
      Height          =   300
      Index           =   5
      Left            =   7365
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   495
      Width           =   1755
   End
   Begin VB.TextBox txtMain 
      Height          =   300
      Index           =   4
      Left            =   7365
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   825
      Width           =   1770
   End
   Begin VB.TextBox txtMain 
      Height          =   300
      Index           =   0
      Left            =   840
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   840
      Width           =   1065
   End
   Begin VB.TextBox txtMain 
      Height          =   300
      Index           =   1
      Left            =   2580
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   825
      Width           =   1245
   End
   Begin VB.TextBox txtMain 
      Height          =   300
      Index           =   2
      Left            =   4425
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   825
      Width           =   600
   End
   Begin VB.TextBox txtMain 
      Height          =   300
      Index           =   3
      Left            =   5655
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   825
      Width           =   540
   End
   Begin VSFlex8Ctl.VSFlexGrid vsList 
      Height          =   3180
      Left            =   45
      TabIndex        =   23
      Top             =   1185
      Width           =   9345
      _cx             =   16484
      _cy             =   5609
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16772055
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   2000
      ColWidthMin     =   0
      ColWidthMax     =   5000
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmLeaveMediMana.frx":6852
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   0   'False
      AutoSizeMode    =   1
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   1
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
      Begin VB.TextBox txtEdit 
         Height          =   375
         Left            =   6120
         TabIndex        =   24
         Top             =   2490
         Visible         =   0   'False
         Width           =   1125
      End
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "[��λ����]�ݴ�ҩƷ�Ǽǵ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   2595
      TabIndex        =   22
      Top             =   105
      Width           =   3795
   End
   Begin VB.Label lblMain 
      Caption         =   "�Ǽ�ʱ��"
      Height          =   240
      Index           =   9
      Left            =   5850
      TabIndex        =   21
      Top             =   4845
      Width           =   735
   End
   Begin VB.Label lblMain 
      Caption         =   "������"
      Height          =   225
      Index           =   8
      Left            =   3270
      TabIndex        =   19
      Top             =   4845
      Width           =   570
   End
   Begin VB.Label lblMain 
      Caption         =   "�ϼ�"
      Height          =   240
      Index           =   7
      Left            =   285
      TabIndex        =   17
      Top             =   4845
      Width           =   390
   End
   Begin VB.Label lblMain 
      Caption         =   "ժҪ"
      Height          =   240
      Index           =   6
      Left            =   270
      TabIndex        =   15
      Top             =   4500
      Width           =   390
   End
   Begin VB.Label lblMain 
      Caption         =   "NO"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Index           =   5
      Left            =   6975
      TabIndex        =   14
      Top             =   525
      Width           =   330
   End
   Begin VB.Label lblMain 
      Caption         =   "���տ���"
      Height          =   240
      Index           =   4
      Left            =   6540
      TabIndex        =   12
      Top             =   885
      Width           =   720
   End
   Begin VB.Label lblMain 
      Caption         =   "�����"
      Height          =   240
      Index           =   0
      Left            =   240
      TabIndex        =   10
      Top             =   885
      Width           =   570
   End
   Begin VB.Label lblMain 
      Caption         =   "����"
      Height          =   240
      Index           =   1
      Left            =   2100
      TabIndex        =   9
      Top             =   885
      Width           =   405
   End
   Begin VB.Label lblMain 
      Caption         =   "�Ա�"
      Height          =   240
      Index           =   2
      Left            =   3990
      TabIndex        =   8
      Top             =   885
      Width           =   405
   End
   Begin VB.Label lblMain 
      Caption         =   "����"
      Height          =   240
      Index           =   3
      Left            =   5190
      TabIndex        =   7
      Top             =   885
      Width           =   405
   End
End
Attribute VB_Name = "frmLeaveMediMana"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Enum tMain
    ����� = 0
    ���� = 1
    �Ա� = 2
    ���� = 3
    �ݴ���� = 4
    NO = 5
    ժҪ = 6
    �ϼ� = 7
    ������ = 8
    �������� = 9
End Enum

Public pMediMaster As New MediMaster
Public pintType As Integer  '״̬: 0-�鿴 1-���� 2-�޸� 3-���ĵǼ�

Dim fntStrike As StdFont  'ɾ����

Private Sub init_vsList()
    Dim strHead As String
    If pintType = 1 Then
        '����
        strHead = "ҩƷ��Դ,900,1;ҩƷ���������,2500,1;���,1600,1;��;,550,1;����,750,7;���㵥λ,450,4;����,750,7;���,1000,7;" & _
                  "ҩƷID,0,1;ҽ��ID,0,1;���ͺ�,0,1;������λ,0,1;����ϵ��,0,1;���ﵥλ,0,1;�����װ,0,1;����,0,1;�ɴ�����,0,1"
    ElseIf pintType = 2 Then
        '�޸�
        strHead = "ҩƷ��Դ,900,1;ҩƷ���������,2500,1;���,1600,1;��;,550,1;����,750,7;���㵥λ,450,4;����,750,7;���,1000,7;" & _
                  "ҩƷID,0,1;ҽ��ID,0,1;���ͺ�,0,1;������λ,0,1;����ϵ��,0,1;���ﵥλ,0,1;�����װ,0,1;����,0,1;�ɴ�����,0,1;UPDATE,0,1"
    ElseIf pintType = 3 Then
        '���ĵǼ�
        strHead = "ҩƷ��Դ,900,1;ҩƷ���������,2500,1;���,1600,1;��;,550,1;��������,900,7;ʹ������,900,7;���㵥λ,450,4;����,0,7;���,0,7;" & _
                  "ҩƷID,0,1;ҽ��ID,0,1;���ͺ�,0,1;������λ,0,1;����ϵ��,0,1;���ﵥλ,0,1;�����װ,0,1;����,0,1;�ɴ�����,0,1;����,0,7;Key,0,1"
    Else
        '�鿴
    End If

    vsList.Redraw = flexRDNone
    Call SetVsFlexGridHead(strHead, vsList)
    With vsList
        .ColDataType(.ColIndex("����")) = flexDTCurrency
        .ColFormat(.ColIndex("����")) = "0.00"
        .ColDataType(.ColIndex("����")) = flexDTCurrency
        .ColFormat(.ColIndex("����")) = "0.00"
        .ColDataType(.ColIndex("���")) = flexDTCurrency
        .ColFormat(.ColIndex("���")) = "0.00"
        .Redraw = True
    End With


End Sub

Private Sub cmdCancle_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    Dim iRow As Integer, strErr As String, blnErr As Boolean, int��� As Integer
    Dim blnDelRow As Boolean, iCol As Integer
    strErr = ""
    int��� = 0
    With vsList
        '---ɾ������
        
        .Select .Rows - 1, .Cols - 1
        blnDelRow = True
        Do While blnDelRow = True
            blnDelRow = False
            For iRow = 1 To .Rows - 1
                Select Case .TextMatrix(iRow, .ColIndex("ҩƷ��Դ"))
                Case "ҽ��"
                    If Val(.TextMatrix(iRow, .ColIndex("ҽ��ID"))) = 0 Or Val(.TextMatrix(iRow, .ColIndex("ҩƷID"))) = 0 Then
                        .RemoveItem iRow
                        blnDelRow = True
                        Exit For
                    End If
                Case "Ŀ¼��"
                    If Val(.TextMatrix(iRow, .ColIndex("ҩƷID"))) = 0 Then
                        .RemoveItem iRow
                        blnDelRow = True
                        Exit For
                    End If
                Case "Ŀ¼��"
                    If Val(.TextMatrix(iRow, .ColIndex("���"))) = 0 Then
                        .RemoveItem iRow
                        blnDelRow = True
                        Exit For
                    End If
                Case Else
                    .RemoveItem iRow
                    blnDelRow = True
                    Exit For
                End Select
            Next
        Loop
        
        For iRow = 1 To .Rows - 1
            blnErr = False
            Select Case .TextMatrix(iRow, .ColIndex("ҩƷ��Դ"))
            Case "ҽ��"
                If Val(.TextMatrix(iRow, .ColIndex("ҽ��ID"))) = 0 Then
                    strErr = strErr & "��[" & iRow & "]��������δָ��ҽ����" & vbNewLine
                    blnErr = True
                End If
                If Val(.TextMatrix(iRow, .ColIndex("ҩƷID"))) = 0 Then
                    strErr = strErr & "��[" & iRow & "]��������δָ��ҩƷ��" & vbNewLine
                    blnErr = True
                End If
            Case "Ŀ¼��"
                If Val(.TextMatrix(iRow, .ColIndex("ҩƷID"))) = 0 Then
                    strErr = strErr & "��[" & iRow & "]��������δָ��ҩƷ��" & vbNewLine
                    blnErr = True
                End If
            Case "Ŀ¼��"
                '#
            Case Else
                strErr = strErr & "��[" & iRow & "]��������δָ��ҩƷ��Դ��" & vbNewLine
                blnErr = True
            End Select

            If Val(.TextMatrix(iRow, .ColIndex("���"))) <= 0 Then
                If .TextMatrix(iRow, .ColIndex("ҩƷ��Դ")) = "ҽ��" Then
                    strErr = strErr & "��[" & iRow & "]��������δ�շѡ�" & vbNewLine
                Else
                    strErr = strErr & "��[" & iRow & "]�������󣬽��������㡣" & vbNewLine
                End If
                blnErr = True
            End If
            
            For iCol = .FixedCols To .Cols - 1
               .TextMatrix(iRow, iCol) = DelInvalidChar(.TextMatrix(iRow, iCol), "'")
            Next
        Next
    End With
    If strErr <> "" Then
        MsgBox strErr, vbQuestion, gstrSysName
        Exit Sub
    End If

    If pintType = 1 Then
        Call AddLeveMedi
    ElseIf pintType = 2 Then
        Call UpdateLeveMedi
    ElseIf pintType = 3 Then
        Call UsedLeveMedi
    
    End If
    '�˳�
    Unload Me
End Sub

Private Sub UsedLeveMedi()
    Dim iRow As Integer, dbl�ı��� As Double, curDate As Date
    If pintType <> 3 Then Exit Sub
    curDate = zlDatabase.Currentdate
    With vsList
        For iRow = 1 To .Rows - 1
            If .TextMatrix(iRow, .ColIndex("ҩƷ��Դ")) Like "Ŀ¼*" Then
                dbl�ı��� = Val(.TextMatrix(iRow, .ColIndex("ʹ������")))
                If dbl�ı��� > 0 Then
                    'ֱ������
                    pMediMaster.ժҪ = txtMain(6)
                    Call pMediMaster.InsertUseBill(.TextMatrix(iRow, .ColIndex("Key")), dbl�ı���, curDate)
                End If
            End If
        Next
    End With
    
End Sub

Private Sub UpdateLeveMedi()
    Dim i As Integer
    If pintType = 2 Then
        '�޸�ģʽ
        For i = 1 To pMediMaster.BillCount
            pMediMaster.Remove 1
        Next
        Call pMediMaster.DeleteBill(1)
        Call AddLeveMedi
    End If
End Sub

Private Sub AddLeveMedi()
    '���������ȷ��
    Dim iRow As Integer, strNO As String, int��� As Integer, int��; As Integer
    Dim objBIll As MediBill
    
    On Error GoTo errHandle
    With vsList
        For iRow = 1 To .Rows - 1
            Set objBIll = New MediBill
            int��� = int��� + 1
            objBIll.���� = Val(.TextMatrix(iRow, .ColIndex("����")))
            objBIll.��� = .TextMatrix(iRow, .ColIndex("���"))
            objBIll.������λ = .TextMatrix(iRow, .ColIndex("������λ"))
            objBIll.����ϵ�� = Val(.TextMatrix(iRow, .ColIndex("����ϵ��")))
            objBIll.��� = Val(.TextMatrix(iRow, .ColIndex("���")))
            objBIll.���㵥λ = .TextMatrix(iRow, .ColIndex("���㵥λ"))
            objBIll.���ﵥλ = .TextMatrix(iRow, .ColIndex("���ﵥλ"))
            objBIll.�����װ = Val(.TextMatrix(iRow, .ColIndex("�����װ")))
            objBIll.���� = Val(.TextMatrix(iRow, .ColIndex("����")))
            objBIll.���ϵ�� = 1
            objBIll.ʹ��״̬ = 0
            objBIll.���� = Val(.TextMatrix(iRow, .ColIndex("����")))
            objBIll.��� = int���
            objBIll.ҩƷID = Val(.TextMatrix(iRow, .ColIndex("ҩƷID")))
            objBIll.ҩƷ���� = .TextMatrix(iRow, .ColIndex("ҩƷ���������"))
            objBIll.ҽ��ID = Val(.TextMatrix(iRow, .ColIndex("ҽ��ID")))
            objBIll.���ͺ� = Val(.TextMatrix(iRow, .ColIndex("���ͺ�")))

            Select Case .TextMatrix(iRow, .ColIndex("��;"))
            Case "��Һ"
                int��; = 1
            Case "ע��"
                int��; = 2
            Case "Ƥ��"
                int��; = 3
            Case Else
                int��; = 0
            End Select
            objBIll.ִ�з��� = int��;
            '�ӵ�����
            Call pMediMaster.AddBill(objBIll, int���)
        Next

        'д�����ݿ���
        If pintType = 1 Then
            '����ʱ��ȡNO��
            strNO = pMediMaster.GetNextNo
        ElseIf pintType = 2 Then
            '�޸�ʱ����ԭNO��
            strNO = pMediMaster.NO
        End If

        If strNO <> "" Then
            pMediMaster.ժҪ = Trim(Replace(txtMain(6), "'", ""))
            Call pMediMaster.InsertBill(strNO, zlDatabase.Currentdate)
            pMediMaster.NO = strNO
            txtMain(tMain.NO) = strNO
        Else
            MsgBox "���ݺŴ���,���ܱ�������!", vbQuestion, Me.Caption
            Exit Sub
        End If
        '��������
        
        .AutoSize 0, .Cols - 1

    End With
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub Form_Activate()
    Dim str��λ���� As String
    On Error GoTo errHandle
    str��λ���� = GetSetting("ZLSOFT", "ע����Ϣ", "��λ����", "")
    
    If pMediMaster.����ID = 0 Then Unload Me
    If pintType <> 0 Then
        '����,�޸�
        vsList.Editable = flexEDKbdMouse
    End If
    txtEdit.Visible = False
    txtMain(tMain.NO) = pMediMaster.NO
    txtMain(tMain.��������) = Format(pMediMaster.�Ǽ�ʱ��, "yyyy-MM-dd hh:mm:ss")
    txtMain(tMain.������) = pMediMaster.����Ա
    txtMain(tMain.�����) = pMediMaster.�����
    txtMain(tMain.����) = pMediMaster.����
    txtMain(tMain.����) = pMediMaster.����
    txtMain(tMain.�Ա�) = pMediMaster.�Ա�
    txtMain(tMain.�ݴ����) = pMediMaster.��������
    txtMain(tMain.ժҪ) = pMediMaster.ժҪ
    If pintType <> 1 Then
        Call init_vsList
        Call Fill_vslist
    Else
        If pintType = 1 Then
            vsList.Select 1, 1
        End If
    End If
    If pintType = 3 Then
        lblTitle.Caption = str��λ���� & "�ݴ�ҩƷʹ�õ�"
    Else
        lblTitle.Caption = str��λ���� & "�ݴ�ҩƷ�Ǽǵ�"
    End If
    If Me.ScaleWidth - lblTitle.Width < 0 Then
        lblTitle.Left = 10
    Else
        lblTitle.Left = (Me.ScaleWidth - lblTitle.Width) / 2
    End If
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub Fill_vslist()
    Dim objMediBill As MediBill
    Dim i As Integer, str��; As String, str��Դ As String
    With vsList
        
        For i = 1 To pMediMaster.BillCount
            Set objMediBill = pMediMaster.BillItem(i)
                If objMediBill.���ϵ�� = 1 Then
                    Select Case objMediBill.ִ�з���
                    Case 1
                        str��; = "��Һ"
                    Case 2
                        str��; = "ע��"
                    Case 3
                        str��; = "Ƥ��"
                    Case Else
                        str��; = "����"
                    End Select
                    If objMediBill.ҩƷID = 0 And objMediBill.ҽ��ID = 0 Then
                        str��Դ = "Ŀ¼��"
                    ElseIf objMediBill.ҩƷID <> 0 And objMediBill.ҽ��ID = 0 Then
                        str��Դ = "Ŀ¼��"
                    ElseIf objMediBill.ҩƷID <> 0 And objMediBill.ҽ��ID <> 0 Then
                        str��Դ = "ҽ��"
                    Else
                        str��Դ = "����"
                    End If

                    .TextMatrix(.Rows - 1, .ColIndex("ҩƷ��Դ")) = str��Դ
                    .TextMatrix(.Rows - 1, .ColIndex("ҩƷ���������")) = objMediBill.ҩƷ����
                    .TextMatrix(.Rows - 1, .ColIndex("���")) = objMediBill.���
                    .TextMatrix(.Rows - 1, .ColIndex("��;")) = str��;
                    .TextMatrix(.Rows - 1, .ColIndex("����")) = objMediBill.����
                    .TextMatrix(.Rows - 1, .ColIndex("���㵥λ")) = objMediBill.���㵥λ
                    .TextMatrix(.Rows - 1, .ColIndex("����")) = objMediBill.����
                    .TextMatrix(.Rows - 1, .ColIndex("���")) = objMediBill.���
                    .TextMatrix(.Rows - 1, .ColIndex("ҩƷID")) = objMediBill.ҩƷID
                    .TextMatrix(.Rows - 1, .ColIndex("ҽ��ID")) = objMediBill.ҽ��ID
                    .TextMatrix(.Rows - 1, .ColIndex("���ͺ�")) = objMediBill.���ͺ�
                    .TextMatrix(.Rows - 1, .ColIndex("������λ")) = objMediBill.������λ
                    .TextMatrix(.Rows - 1, .ColIndex("����ϵ��")) = objMediBill.����ϵ��
                    .TextMatrix(.Rows - 1, .ColIndex("���ﵥλ")) = objMediBill.���ﵥλ
                    .TextMatrix(.Rows - 1, .ColIndex("�����װ")) = objMediBill.�����װ
                    .TextMatrix(.Rows - 1, .ColIndex("����")) = objMediBill.����
                    If pintType = 3 Then
                        .TextMatrix(.Rows - 1, .ColIndex("ʹ������")) = 0
                        .TextMatrix(.Rows - 1, .ColIndex("��������")) = objMediBill.���� - objMediBill.��������
                        .TextMatrix(.Rows - 1, .ColIndex("Key")) = objMediBill.��� & "_" & objMediBill.���ϵ�� & "_" & Format(objMediBill.�Ǽ�ʱ��, "yyMMddhhmmss")
                        If objMediBill.���� - objMediBill.�������� <= 0 Or str��Դ = "ҽ��" Then
                            .RemoveItem (.Rows - 1)
                        End If
                    End If
                    '.TextMatrix(.Rows - 1, .ColIndex("�ɴ�����")) = objMediBill.�ɴ�����
                    
                    .Rows = .Rows + 1
                End If
        Next
        If .Rows > 2 Then
            .RemoveItem (.Rows - 1)
        End If
        
        txtMain(tMain.�ϼ�) = Format(.Aggregate(flexSTSum, 1, .ColIndex("���"), .Rows - 1, .ColIndex("���")), "0.00")
        If pintType = 3 Then
            If .Rows < 3 Then
                If .TextMatrix(.Rows - 1, .ColIndex("ҩƷ���������")) = "" Then
                    MsgBox "û�пɵǼǵ��ݴ�ҩƷ!", vbInformation, Me.Caption
                    Me.cmdOk.Enabled = False
                End If
            End If
        End If
        .AutoSize 1, .Cols - 1
    End With
End Sub

Private Sub vsListButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strInput As String
    Dim vRect As RECT, blnCanel As Boolean, strSelectRow As String '����ҽ�����ų���ѡ��ļ�¼
    Dim strPar As String, strType As String, i As Integer
    Dim strNO As String
    
    On Error GoTo errHandle
    
    If pintType = 0 Or pintType = 3 Then Exit Sub
    
    If Col = vsList.ColIndex("ҩƷ���������") Then
        If vsList.TextMatrix(vsList.Row, vsList.ColIndex("ҩƷ��Դ")) = "Ŀ¼��" Then Exit Sub

        'Ŀ¼��
        '--------------------------------------------------------------------------------------
        If vsList.TextMatrix(vsList.Row, vsList.ColIndex("ҩƷ��Դ")) = "Ŀ¼��" Then
            strInput = DelInvalidChar(UCase(Trim(txtEdit)), "'")
            If InStr(strInput, "]") > 0 Then
                strInput = Mid(Split(strInput, "]")(0), 2)
            End If
            If strInput = "" Then
                strSQL = "Select A.ID, A.����, A.����,A.���, A.���㵥λ, B.�ּ�, A.��������, Decode(A.�������, 1, '����', '�����סԺ') As �������," & vbNewLine & _
                    "       A.ִ�п���,C.����ϵ��,C.���ﵥλ,C.�����װ,C.����,D.������λ" & vbNewLine & _
                    "From ҩƷ��Ϣ D,ҩƷ��� C,(Select �ּ�, �շ�ϸĿid,�۸�ȼ� From �շѼ�Ŀ Where ��ֹ���� Is Null Or ��ֹ���� = To_Date('3000-01-01', 'YYYY-MM-DD')) B," & vbNewLine & _
                    "     �շ���ĿĿ¼ A" & vbNewLine & _
                    "Where C.ҩ��ID=D.ҩ��ID And A.Id=C.ҩƷID And A.ID = B.�շ�ϸĿid And Mod(A.�������, 2) = 1 And" & vbNewLine & _
                    "      (A.����ʱ�� Is Null Or A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD')) And" & vbNewLine & _
                    "      A.��� In ( '5','6')" & GetPriceGradeSQL(gstrҩƷ�۸�ȼ�, gstr���ļ۸�ȼ�, gstr��ͨ��Ŀ�۸�ȼ�, "A", "B", "1", "2", "3")
            Else
                strSQL = "Select A.ID, A.����, A.����,A.��� , A.���㵥λ, B.�ּ�, A.��������, Decode(A.�������, 1, '����', '�����סԺ') As �������," & vbNewLine & _
                    "       A.ִ�п���,C.����ϵ��,C.���ﵥλ,C.�����װ,C.����,D.������λ" & vbNewLine & _
                    "From �շ���Ŀ���� E,ҩƷ��Ϣ D,ҩƷ��� C,(Select �ּ�, �շ�ϸĿid,�۸�ȼ� From �շѼ�Ŀ Where ��ֹ���� Is Null Or ��ֹ���� = To_Date('3000-01-01', 'YYYY-MM-DD')) B," & vbNewLine & _
                    "     �շ���ĿĿ¼ A" & vbNewLine & _
                    "Where A.ID = E.�շ�ϸĿid And  C.ҩ��ID=D.ҩ��ID And A.Id=C.ҩƷID And A.ID = B.�շ�ϸĿid And Mod(A.�������, 2) = 1 And" & vbNewLine & _
                    "       (E.���� Like '%" & strInput & "%' Or A.���� Like '%" & strInput & "%' Or A.���� Like '%" & strInput & "%') And E.���� = 1 And" & vbNewLine & _
                    "      (A.����ʱ�� Is Null Or A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD')) And" & vbNewLine & _
                    "      A.��� In ( '5','6')" & GetPriceGradeSQL(gstrҩƷ�۸�ȼ�, gstr���ļ۸�ȼ�, gstr��ͨ��Ŀ�۸�ȼ�, "A", "B", "1", "2", "3")
            End If

            vRect = ZLControl.GetControlRect(txtEdit.hwnd)
            Set rsTmp = New ADODB.Recordset
            Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "ҩƷ", False, "", "ѡ��ҩƷ", False, False, True, _
                                                 vRect.Left, vRect.Top, txtEdit.Height, blnCanel, True, True, gstrҩƷ�۸�ȼ�, gstr���ļ۸�ȼ�, gstr��ͨ��Ŀ�۸�ȼ�)
            If Not blnCanel And rsTmp.State <> 0 Then
                If Not rsTmp.EOF Then
                    With vsList
                        .EditText = Replace("[" & zlCommFun.NVL(rsTmp.Fields("����")) & "] " & zlCommFun.NVL(rsTmp.Fields("����")), "[]", "")
                        .TextMatrix(.Row, .ColIndex("ҩƷ���������")) = Replace("[" & zlCommFun.NVL(rsTmp.Fields("����")) & "] " & zlCommFun.NVL(rsTmp.Fields("����")), "[]", "")
                        .TextMatrix(.Row, .ColIndex("����")) = Format(zlCommFun.NVL(rsTmp.Fields("�ּ�"), 0), "0.00")
                        .TextMatrix(.Row, .ColIndex("���㵥λ")) = zlCommFun.NVL(rsTmp.Fields("���㵥λ"), "")
                        .TextMatrix(.Row, .ColIndex("���")) = zlCommFun.NVL(rsTmp.Fields("���"), "")
                        .TextMatrix(.Row, .ColIndex("ҩƷID")) = zlCommFun.NVL(rsTmp.Fields("ID"), "")
                        .TextMatrix(.Row, .ColIndex("ҽ��ID")) = 0
                        .TextMatrix(.Row, .ColIndex("������λ")) = zlCommFun.NVL(rsTmp.Fields("������λ"), "")
                        .TextMatrix(.Row, .ColIndex("����ϵ��")) = zlCommFun.NVL(rsTmp.Fields("����ϵ��"), "")
                        .TextMatrix(.Row, .ColIndex("���ﵥλ")) = zlCommFun.NVL(rsTmp.Fields("���ﵥλ"), "")
                        .TextMatrix(.Row, .ColIndex("�����װ")) = zlCommFun.NVL(rsTmp.Fields("�����װ"), "")
                    End With
                End If
                Set rsTmp = Nothing
            End If
            txtEdit = ""
        End If

        'ҽ��
        '--------------------------------------------------------------------------------------
        If vsList.TextMatrix(vsList.Row, vsList.ColIndex("ҩƷ��Դ")) = "ҽ��" Then '
            strInput = DelInvalidChar(UCase(Trim(txtEdit)), "'")
            If InStr(strInput, "]") > 0 Then
                strInput = Mid(Split(strInput, "]")(0), 2)
            End If

            'ȡ����ѡ�����
            strSelectRow = ""
            For i = 1 To vsList.Rows - 1
                With vsList
                    If Val(.TextMatrix(i, .ColIndex("ҽ��ID"))) > 0 Then
                        strSelectRow = strSelectRow & Val(.TextMatrix(i, .ColIndex("ҽ��ID"))) & "_" & Val(.TextMatrix(i, .ColIndex("���ͺ�"))) & ","
                    End If
                End With
            Next
            
            strPar = zlDatabase.GetPara("��ʾ��������", glngSys, 1264, "1,1,1,1")
            For i = 0 To 3
                strType = strType & IIf(Val(Split(strPar, ",")(i)) = 1, "," & i, "")
            Next
            
            If pMediMaster.�Һŵ� Like "*_*" Then
                '��������
                strNO = " a.��ҳid = " & Split(pMediMaster.�Һŵ�, "_")(1) & " "
            Else
                '����
                strNO = " a.�Һŵ� = '" & pMediMaster.�Һŵ� & "' "
            End If
            
            If strInput = "" Then
                strSQL = "Select c.���id, i.����ʱ��, i.No, b.ִ�з��� as ��;, i.���ͺ�, i.ҽ��id, d.����, d.����, d.���, nvl(g.������λ,'') as ������λ, nvl(e.����ϵ��,0) ����ϵ��," & vbNewLine & _
                        "            nvl(e.���ﵥλ,'') as ���ﵥλ , nvl(e.�����װ,0) �����װ, nvl(e.����,0) as ����, h.��׼���� As �ּ�, (Nvl(i.��������, 0) / Nvl(e.����ϵ��, 0)) As �ɴ�����, c.�շ�ϸĿid," & vbNewLine & _
                        "            (Nvl(f.����, 0) * Nvl(f.���ϵ��, 0)) As �Ѵ�����, d.���㵥λ, C.��������, e.����ɷ���� " & vbNewLine & _
                        "From ������ü�¼ h," & vbNewLine & _
                        "        ҩƷ��Ϣ g, �ݴ�ҩƷ��¼ f, ҩƷ��� e, �շ���ĿĿ¼ d, ����ҽ����¼ c, ������ĿĿ¼ b, ����ҽ������ i," & vbNewLine & _
                        "        ����ҽ����¼ a" & vbNewLine & _
                        "Where Instr('" & strType & "', nvl(b.ִ�з���,0))>0 And c.id = h.ҽ�����(+) And h.��¼״̬(+)=1 and h.����״̬(+)<>1 And e.ҩ��id = g.ҩ��id And Mod(d.�������, 2) = 1 And" & vbNewLine & _
                        "           (d.����ʱ�� Is Null Or d.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD')) And " & vbNewLine & _
                        "           i.ҽ��id = f.ҽ��id(+) And i.���ͺ� = f.���ͺ�(+) And c.�շ�ϸĿid = e.ҩƷid And c.�շ�ϸĿid = d.Id And" & vbNewLine & _
                        "           (Nvl(I.��������, 0) / Nvl(E.����ϵ��, 0)) - (Nvl(F.����, 0) * Nvl(F.���ϵ��, 0)) > 0 And " & vbNewLine & _
                        "           c.������� In ('5', '6') And a.Id = c.���id And a.������Ŀid = b.Id And i.ҽ��id = c.Id And f.����id(+) = " & pMediMaster.����ID & " And" & vbNewLine & _
                        "           a.����id = " & pMediMaster.����ID & " And a.ִ�п���id = " & pMediMaster.����ID & " And F.���ϵ��(+)=1 And a.������� = 'E' And " & strNO & vbNewLine & _
                        IIf(strSelectRow = "", "", " And Instr('" & strSelectRow & "' ,I.ҽ��ID||'_'||I.���ͺ�||',')<=0 ") & vbNewLine & _
                        "Order By ���ͺ�, ҽ��id,���id"
            Else
                strSQL = "Select c.���id, i.����ʱ��, i.No, b.ִ�з��� as ��;, i.���ͺ�, i.ҽ��id, d.����, d.����, d.���, nvl(g.������λ,'') as ������λ, nvl(e.����ϵ��,0) as ����ϵ��," & vbNewLine & _
                        "            nvl(e.���ﵥλ,'') as ���ﵥλ, nvl(e.�����װ,0) as �����װ, nvl(e.����,0) as ����, h.��׼���� As �ּ�, (Nvl(i.��������, 0) / Nvl(e.����ϵ��, 0)) As �ɴ�����, c.�շ�ϸĿid," & vbNewLine & _
                        "            (Nvl(f.����, 0) * Nvl(f.���ϵ��, 0)) As �Ѵ�����, d.���㵥λ,C.��������, e.����ɷ���� " & vbNewLine & _
                        "From ������ü�¼ h," & vbNewLine & _
                        "        ҩƷ��Ϣ g, �ݴ�ҩƷ��¼ f, ҩƷ��� e, �շ���ĿĿ¼ d, ����ҽ����¼ c, ������ĿĿ¼ b, ����ҽ������ i," & vbNewLine & _
                        "        ����ҽ����¼ a" & vbNewLine & _
                        "Where  Instr('" & strType & "', nvl(b.ִ�з���,0))>0 And c.id = h.ҽ�����(+) And h.��¼״̬(+)=1 And h.����״̬(+)<>1 And e.ҩ��id = g.ҩ��id And Mod(d.�������, 2) = 1 And" & vbNewLine & _
                        "           (d.����ʱ�� Is Null Or d.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD')) And" & vbNewLine & _
                        "                        (Zlspellcode(d.����) Like '%" & strInput & "%' Or d.���� Like '%" & strInput & "%' Or" & vbNewLine & _
                        "                        d.���� Like '%" & strInput & "%') And" & vbNewLine & _
                        "           i.ҽ��id = f.ҽ��id(+) And i.���ͺ� = f.���ͺ�(+) And c.�շ�ϸĿid = e.ҩƷid And c.�շ�ϸĿid = d.Id And" & vbNewLine & _
                        "           (Nvl(I.��������, 0) / Nvl(E.����ϵ��, 0)) - (Nvl(F.����, 0) * Nvl(F.���ϵ��, 0)) > 0 And " & vbNewLine & _
                        "           c.������� In ('5', '6') And a.Id = c.���id And a.������Ŀid = b.Id And i.ҽ��id = c.Id And f.����id(+) = " & pMediMaster.����ID & " And" & vbNewLine & _
                        "           a.����id = " & pMediMaster.����ID & " And a.ִ�п���id = " & pMediMaster.����ID & " And F.���ϵ��(+)=1 And a.������� = 'E' And " & strNO & vbNewLine & _
                        IIf(strSelectRow = "", "", " And Instr('" & strSelectRow & "' ,I.ҽ��ID||'_'||I.���ͺ�||',')<=0 ") & vbNewLine & _
                        "Order By ���ͺ�, ҽ��id,���id"
            End If
            Call frmLeaveSelect.LeaveSelect(Me, strSQL)
            txtEdit = ""
       

        End If
    End If
    Call zlCommFun.PressKey(vbKeyRight)
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If

End Sub

Private Sub Form_Load()
    Call init_vsList
End Sub

Private Sub Form_Unload(Cancel As Integer)
    pintType = 0
    Set pMediMaster = Nothing
End Sub


Private Sub vsList_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim str��Դ
    If Col = vsList.ColIndex("ҩƷ��Դ") Then
        str��Դ = vsList.TextMatrix(Row, Col)
        vsList.Delete
        vsList.TextMatrix(Row, Col) = str��Դ
    End If
End Sub

Private Sub vsList_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)

    Dim blnEdit As Boolean
    Call vsList_BeforeEdit(NewRow, NewCol, blnEdit)
    If blnEdit Then
        vsList.ComboList = ""
        'vsList.FocusRect = flexFocusLight
    Else
        'vsList.FocusRect = flexFocusSolid
        If NewCol = vsList.ColIndex("ҩƷ���������") Then
            vsList.ComboList = "..."
        ElseIf NewCol = vsList.ColIndex("ҩƷ��Դ") Then
            vsList.ComboList = "ҽ��|Ŀ¼��|Ŀ¼��"
        ElseIf NewCol = vsList.ColIndex("��;") Then
            vsList.ComboList = "����|��Һ|ע��|Ƥ��"
        Else
            vsList.ComboList = ""
        End If
    End If

End Sub

Private Sub vsList_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim strEditRow As String
    On Error GoTo errHandle

    If pintType = 1 Then
        With vsList
        Select Case .TextMatrix(Row, .ColIndex("ҩƷ��Դ"))
        Case "Ŀ¼��"
            strEditRow = "," & .ColIndex("ҩƷ��Դ") & "," & .ColIndex("ҩƷ���������") & "," & .ColIndex("����") & "," & .ColIndex("��;") & ","
        Case "Ŀ¼��"
            strEditRow = "," & .ColIndex("ҩƷ��Դ") & "," & .ColIndex("ҩƷ���������") & "," & .ColIndex("���") & "," & .ColIndex("����") & "," & .ColIndex("����") & "," & .ColIndex("���㵥λ") & "," & .ColIndex("��;") & ","
        Case Else
            strEditRow = "," & .ColIndex("ҩƷ��Դ") & "," & .ColIndex("ҩƷ���������") & "," & .ColIndex("����") & ","
        End Select
        End With
    ElseIf pintType = 2 Then
        With vsList
        Select Case .TextMatrix(Row, .ColIndex("ҩƷ��Դ"))
        Case "Ŀ¼��"
            strEditRow = "," & .ColIndex("ҩƷ��Դ") & "," & .ColIndex("ҩƷ���������") & "," & .ColIndex("����") & "," & .ColIndex("��;") & ","
        Case "Ŀ¼��"
            strEditRow = "," & .ColIndex("ҩƷ��Դ") & "," & .ColIndex("ҩƷ���������") & "," & .ColIndex("���") & "," & .ColIndex("����") & "," & .ColIndex("����") & "," & .ColIndex("���㵥λ") & "," & .ColIndex("��;") & ","
        Case Else
            strEditRow = "," & .ColIndex("ҩƷ��Դ") & "," & .ColIndex("ҩƷ���������") & "," & .ColIndex("����") & ","
        End Select
        End With
    ElseIf pintType = 3 Then
        strEditRow = "," & vsList.ColIndex("ʹ������") & ","
    Else
        strEditRow = ""
    End If

    If InStr(strEditRow, "," & Col & ",") <= 0 Then
        Cancel = True
    End If
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If

End Sub


Private Sub vsList_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Call vsListButtonClick(Row, Col)
End Sub

Private Sub vsList_EnterCell()
    On Error GoTo errHandle
    With vsList
    
        If pintType = 0 Or pintType = 3 Then Exit Sub
        If .Col = .ColIndex("ҩƷ���������") And .Row > 0 Then
            If txtEdit.Tag = "False" And InStr("Ŀ¼��,ҽ��", .TextMatrix(.Row, .ColIndex("ҩƷ��Դ"))) > 0 Then
                txtEdit.Left = .CellLeft
                txtEdit.Top = .CellTop
                txtEdit.Height = .CellHeight - 12
                txtEdit.Width = .CellWidth - 12
                txtEdit.Tag = "True"
            End If
        Else
            txtEdit.Tag = "False"
        End If
        Dim blnCancle As Boolean
        Call vsList_BeforeEdit(.Row, .Col, blnCancle)
        If Not blnCancle Then
            Call .CellBorder(vsList.GridColor, 1, 1, 2, 2, 0, 0)
        End If
    End With
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub vsList_KeyDown(KeyCode As Integer, Shift As Integer)

On Error GoTo errHandle
    Dim strLastҩƷ��Դ As String
    With vsList
        If KeyCode = vbKeyReturn Then
            If .Col = .ColIndex("���") And .Row = .Rows - 1 And (pintType = 1 Or pintType = 2) Then
                strLastҩƷ��Դ = .TextMatrix(.Row, .ColIndex("ҩƷ��Դ"))

                .Rows = .Rows + 1
                .Row = .Row + 1

                If strLastҩƷ��Դ <> "" Then
                    .TextMatrix(.Row, .ColIndex("ҩƷ��Դ")) = strLastҩƷ��Դ
                    .Col = .ColIndex("ҩƷ���������")
                Else
                    .Col = .ColIndex("ҩƷ��Դ")
                End If
            Else
                If .Cols > .Col + 1 And .Col <> .ColIndex("���") Then
                    .Col = .Col + 1
                Else
                    If .Rows > .Row + 1 Then
                        .Row = .Row + 1
                        .Col = .ColIndex("ҩƷ���������")
                    End If
                End If
            End If
        End If
        If pintType = 1 Or pintType = 2 Then
            '����
            If KeyCode = vbKeyDelete Then
                If MsgBox("�Ƿ�ɾ����ǰ��?", vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbYes Then
                    If .Rows > 2 Then
                        .RemoveItem (.Row)
                    Else
                        .Delete
                    End If
                End If
            End If
            
        End If
    End With
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub vsList_KeyPress(KeyAscii As Integer)
    If pintType = 0 Or pintType = 3 Then Exit Sub
    With vsList
        If (.Col = .ColIndex("ҩƷ���������")) And KeyAscii = vbKeyReturn Then
            KeyAscii = 0
        Else
            If .Col = .ColIndex("ҩƷ���������") And vsList.ComboList = "..." Then
                If KeyAscii = Asc("*") Then
                    KeyAscii = 0
                    txtEdit.Text = .EditText
                    Call vsList_CellButtonClick(.Row, .Col)
                    txtEdit.Tag = False
                    txtEdit.Visible = False
                Else
                    .ComboList = "" 'ʹ��ť״̬��������״̬
                End If
            End If
        End If
    End With
End Sub

Private Sub vsList_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)

On Error GoTo errHandle
    If pintType = 0 Or pintType = 3 Then Exit Sub
    With vsList
        If Col = .ColIndex("ҩƷ���������") And KeyAscii = vbKeyReturn And .TextMatrix(.Row, .ColIndex("ҩƷ��Դ")) <> "Ŀ¼��" Then
            txtEdit.Text = .EditText
            .EditText = ""
            Call vsListButtonClick(Row, Col)
            txtEdit.Tag = False
            txtEdit.Visible = False
        ElseIf KeyAscii = vbKeyReturn Then
            If .Cols < .Col + 1 And .Col <> .ColIndex("���") Then
                .Col = .Col + 1
            Else
                If .Rows < .Row + 1 Then
                    .Row = .Row + 1
                    .Col = .ColIndex("ҩƷ���������")
                End If
            End If
        End If
    End With

    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub vsList_LeaveCell()
    With vsList
        On Error GoTo errHandle
        If pintType = 1 Or pintType = 2 Then
            If Val(.TextMatrix(.Row, .ColIndex("����"))) <> 0 Then
                .TextMatrix(.Row, .ColIndex("���")) = Format(Val(.TextMatrix(.Row, .ColIndex("����"))) * Val(.TextMatrix(.Row, .ColIndex("����"))), "0.00")
            Else
                .TextMatrix(.Row, .ColIndex("���")) = "0.00"
            End If
        End If
        txtMain(tMain.�ϼ�) = Format(.Aggregate(flexSTSum, 1, .ColIndex("���"), .Rows - 1, .ColIndex("���")), "0.00")
        If .TextMatrix(.Row, .ColIndex("ҩƷ��Դ")) = "" Then .TextMatrix(.Row, .ColIndex("ҩƷ��Դ")) = "ҽ��"
    
        Dim blnCancle As Boolean
        Call vsList_BeforeEdit(.Row, .Col, blnCancle)
        If Not blnCancle Then
            On Error Resume Next
            Call .CellBorder(vsList.GridColor, 0, 0, 0, 0, 0, 0)
        End If
        
    End With
    
    Exit Sub
errHandle:
    If Err.Number = 381 Then Exit Sub
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub vsList_RowColChange()

On Error GoTo errHandle
    With vsList
        If txtEdit.Tag = "True" Then
            txtEdit.Left = .CellLeft
            txtEdit.Top = .CellTop
            txtEdit.Height = .CellHeight - 12
            txtEdit.Width = .CellWidth - 12
        End If
        
        If pintType = 1 Or pintType = 2 Then
            If Val(.TextMatrix(.Row, .ColIndex("����"))) <> 0 Then
                .TextMatrix(.Row, .ColIndex("���")) = Format(Val(.TextMatrix(.Row, .ColIndex("����"))) * Val(.TextMatrix(.Row, .ColIndex("����"))), "0.00")
            Else
                .TextMatrix(.Row, .ColIndex("���")) = "0.00"
            End If
        End If
        txtMain(tMain.�ϼ�) = Format(.Aggregate(flexSTSum, 1, .ColIndex("���"), .Rows - 1, .ColIndex("���")), "0.00")
    End With
    Exit Sub
errHandle:
    If Err.Number = 381 Then Exit Sub
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub vsList_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim rsTmp As New ADODB.Recordset, strSQL As String
    Dim lng�ɴ����� As Long
    On Error GoTo errHandle
    With vsList
    Select Case Col
        Case .ColIndex("����"), .ColIndex("����"), .ColIndex("��� ")
            If IsNumeric(.EditText) = False Then Cancel = True
            If .ColIndex("����") Then
                '����ʱ,������������ܴ��ڿɴ�����
                If pintType = 1 And .TextMatrix(Row, .ColIndex("ҩƷ��Դ")) = "ҽ��" Then
                    If Val(.EditText) > Val(.TextMatrix(Row, .ColIndex("�ɴ�����"))) Then
                        MsgBox "������д���󣬴�ҩƷ���ֻ�ܼĴ� " & Val(.TextMatrix(Row, .ColIndex("�ɴ�����"))) & " " & .TextMatrix(Row, .ColIndex("���㵥λ")), vbQuestion, Me.Caption
                        Cancel = True
                    End If
                End If
                '�޸�ʱ,����ȡ�ɴ�����
                If pintType = 2 And .TextMatrix(Row, .ColIndex("ҩƷ��Դ")) = "ҽ��" Then
'                    strSQL = "Select c.���id, i.����ʱ��, i.No, b.ִ�з��� as ��;, i.���ͺ�, i.ҽ��id, d.����, d.����, d.���, g.������λ, e.����ϵ��," & vbNewLine & _
'                            "            e.���ﵥλ, e.�����װ, e.����, h.�ּ�, (Nvl(i.��������, 0) / Nvl(e.����ϵ��, 0)) As �ɴ�����, c.�շ�ϸĿid," & vbNewLine & _
'                            "            (Nvl(f.����, 0) * Nvl(f.���ϵ��, 0)) As �Ѵ�����, d.���㵥λ" & vbNewLine & _
'                            "From (Select �ּ�, �շ�ϸĿid From �շѼ�Ŀ Where ��ֹ���� Is Null Or ��ֹ���� = To_Date('3000-01-01', 'YYYY-MM-DD')) h," & vbNewLine & _
'                            "        ҩƷ��Ϣ g, �ݴ�ҩƷ��¼ f, ҩƷ��� e, �շ���ĿĿ¼ d, ����ҽ����¼ c, ������ĿĿ¼ b, ����ҽ������ i," & vbNewLine & _
'                            "        ����ҽ����¼ a" & vbNewLine & _
'                            "Where c.�շ�ϸĿid = h.�շ�ϸĿid(+) And e.ҩ��id = g.ҩ��id And Mod(d.�������, 2) = 1 And" & vbNewLine & _
'                            "           (d.����ʱ�� Is Null Or d.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD')) And Nvl(d.�Ƿ���, 0) = 0 And" & vbNewLine & _
'                            "           i.ҽ��id = f.ҽ��id(+) And i.���ͺ� = f.���ͺ�(+) And c.�շ�ϸĿid = e.ҩƷid And c.�շ�ϸĿid = d.Id And" & vbNewLine & _
'                            "           c.������� In ('5', '6') And a.Id = c.���id And a.������Ŀid = b.Id And i.ҽ��id = c.Id And f.����id(+) = [2] And" & vbNewLine & _
'                            "           a.����id = [1] And a.ִ�п���id = [2] And a.������� = 'E' And a.�Һŵ� = [3]" & vbNewLine & _
'                            " And F.���ϵ��(+)=1 And i.ҽ��ID=[4] And i.���ͺ�=[5]  " & vbNewLine & _
'                            "Order By ���ͺ�, ҽ��id,���id"
                    strSQL = "Select c.���id, i.����ʱ��, i.No, b.ִ�з��� as ��;, i.���ͺ�, i.ҽ��id, d.����, d.����, d.���, g.������λ, e.����ϵ��," & vbNewLine & _
                            "            e.���ﵥλ, e.�����װ, e.����, h.��׼���� As �ּ�, (Nvl(i.��������, 0) / Nvl(e.����ϵ��, 0)) As �ɴ�����, c.�շ�ϸĿid," & vbNewLine & _
                            "            (Nvl(f.����, 0) * Nvl(f.���ϵ��, 0)) As �Ѵ�����, d.���㵥λ" & vbNewLine & _
                            "From ������ü�¼ h," & vbNewLine & _
                            "        ҩƷ��Ϣ g, �ݴ�ҩƷ��¼ f, ҩƷ��� e, �շ���ĿĿ¼ d, ����ҽ����¼ c, ������ĿĿ¼ b, ����ҽ������ i," & vbNewLine & _
                            "        ����ҽ����¼ a" & vbNewLine & _
                            "Where c.id = h.ҽ�����(+) And e.ҩ��id = g.ҩ��id And Mod(d.�������, 2) = 1 And" & vbNewLine & _
                            "           (d.����ʱ�� Is Null Or d.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD')) And" & vbNewLine & _
                            "           i.ҽ��id = f.ҽ��id(+) And i.���ͺ� = f.���ͺ�(+) And c.�շ�ϸĿid = e.ҩƷid And c.�շ�ϸĿid = d.Id And" & vbNewLine & _
                            "           c.������� In ('5', '6') And a.Id = c.���id And a.������Ŀid = b.Id And i.ҽ��id = c.Id And f.����id(+) = [2] And" & vbNewLine & _
                            "           a.����id = [1] And a.ִ�п���id = [2] And a.������� = 'E' And a.�Һŵ� = [3]" & vbNewLine & _
                            " And F.���ϵ��(+)=1 And i.ҽ��ID=[4] And i.���ͺ�=[5]  " & vbNewLine & _
                            "Order By ���ͺ�, ҽ��id,���id"
                     Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "����ȡ�ɴ�����", pMediMaster.����ID, pMediMaster.����ID, pMediMaster.�Һŵ�, Val(.TextMatrix(Row, .ColIndex("ҽ��ID"))), Val(.TextMatrix(Row, .ColIndex("���ͺ�"))))
                     If Not rsTmp.EOF Then
                        lng�ɴ����� = rsTmp.Fields("�ɴ�����") - (rsTmp.Fields("�Ѵ�����") - Val(.TextMatrix(Row, .ColIndex("����"))))
                        If Val(.EditText) > lng�ɴ����� Then
                            MsgBox "������д���󣬴�ҩƷ���ֻ�ܼĴ� " & lng�ɴ����� & " " & .TextMatrix(Row, .ColIndex("���㵥λ")), vbQuestion, Me.Caption
                            Cancel = True
                        End If
                     End If
                End If

            End If
        Case .ColIndex("ʹ������")
            If IsNumeric(.EditText) = True Then
                If Val(.TextMatrix(Row, .ColIndex("��������"))) < Val(.EditText) Or Val(.EditText) < 0 Then
                    If Val(.EditText) > 0 Then
                        MsgBox "ʹ���������ܴ��ڿ�������!", vbQuestion, Me.Caption
                    ElseIf Val(.EditText) < 0 Then
                        MsgBox "ʹ����������С��0!", vbQuestion, Me.Caption
                    End If
                    Cancel = True
                End If
            Else
                Cancel = True
            End If
        
        Case .ColIndex("ҩƷ��Դ")
            If InStr(",ҽ��,Ŀ¼��,Ŀ¼��,", "," & .EditText & ",") <= 0 Then Cancel = True
        Case .ColIndex("��;")
            If InStr(",����,��Һ,ע��,Ƥ��,", "," & .EditText & ",") <= 0 Then Cancel = True

    End Select
    End With
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
