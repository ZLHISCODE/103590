VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "VSFLEX8.OCX"
Begin VB.Form frmAutoPrintOther 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�Զ���ӡ��������"
   ClientHeight    =   6675
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9750
   Icon            =   "frmAutoPrintOther.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   9750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      Caption         =   " ����"
      ForeColor       =   &H80000008&
      Height          =   4335
      Left            =   4560
      TabIndex        =   9
      Top             =   1560
      Width           =   5055
      Begin VSFlex8Ctl.VSFlexGrid vsfRptPara 
         Height          =   3975
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   4800
         _cx             =   8467
         _cy             =   7011
         Appearance      =   0
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
         BackColorSel    =   8421376
         ForeColorSel    =   16777215
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483640
         GridColorFixed  =   -2147483640
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   9
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   255
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmAutoPrintOther.frx":030A
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
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
         OwnerDraw       =   0
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
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      Caption         =   " ����"
      ForeColor       =   &H80000008&
      Height          =   4335
      Left            =   120
      TabIndex        =   7
      Top             =   1560
      Width           =   4335
      Begin VSFlex8Ctl.VSFlexGrid vsfRpt 
         Height          =   3975
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   4080
         _cx             =   7197
         _cy             =   7011
         Appearance      =   0
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
         BackColorSel    =   8421376
         ForeColorSel    =   16777215
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483640
         GridColorFixed  =   -2147483640
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   255
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmAutoPrintOther.frx":04AB
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
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
         OwnerDraw       =   0
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
      End
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "���(&O)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   7200
      TabIndex        =   6
      Top             =   6120
      Width           =   1100
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   8400
      TabIndex        =   5
      Top             =   6120
      Width           =   1100
   End
   Begin VB.TextBox txt���� 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   600
      TabIndex        =   3
      Top             =   1140
      Width           =   1845
   End
   Begin VB.Frame Frame1 
      Height          =   90
      Left            =   0
      TabIndex        =   0
      Top             =   960
      Width           =   10660
   End
   Begin VB.Label lblComment 
      AutoSize        =   -1  'True
      Caption         =   "��ʾ��"
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
      Height          =   210
      Left            =   240
      TabIndex        =   11
      Top             =   6190
      Width           =   630
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "*���뱨���������ƽ��в���"
      Height          =   180
      Left            =   2520
      TabIndex        =   4
      Top             =   1200
      Width           =   2430
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "����"
      Height          =   180
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   360
   End
   Begin VB.Label Label1 
      Caption         =   $"frmAutoPrintOther.frx":051C
      ForeColor       =   &H00808000&
      Height          =   735
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   9255
   End
End
Attribute VB_Name = "frmAutoPrintOther"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mstrRpt As String   '����������ı�����Ϣ������,����

Public Sub ShowForm(ByVal frmParent As Object, ByRef strRpt As String)
        
    frmAutoPrintOther.Show vbModal, frmParent
    
    strRpt = mstrRpt
End Sub

Private Sub ShowRpt(Optional ByVal strCodeOrName As String)
    If grsRpt Is Nothing Then Exit Sub
    
    With vsfRptPara
        .rows = 1
    End With
    
    CmdOK.Enabled = False
    mstrRpt = ""
    
    With vsfRpt
        .Redraw = flexRDNone
        .rows = 1
        
        grsRpt.Filter = IIf(strCodeOrName = "", "", "��� Like '*" & strCodeOrName & "*' Or ���� Like '*" & strCodeOrName & "*'")
          
        Do While Not grsRpt.EOF
            .rows = .rows + 1
            
            .TextMatrix(.rows - 1, .ColIndex("���")) = grsRpt!���
            .TextMatrix(.rows - 1, .ColIndex("����")) = grsRpt!����
            
            grsRpt.MoveNext
        Loop
        
        .Redraw = flexRDDirect
        
        '���ֻ��һ�����ݣ���λ������
        If .rows = 2 Then
            If .TextMatrix(1, .ColIndex("���")) <> "" Then
                .Row = 1
            End If
        End If
    End With
End Sub

Private Sub ShowRptPara(ByVal strCode As String)
    Dim n As Integer
    Dim blnCheckҩ�� As Boolean, blnCheck���� As Boolean, blnCheckNO As Boolean, blnCheckOther As Boolean
     
    lblComment.Caption = ""
    
    With vsfRptPara
        .rows = 1
        
        If grsRpt Is Nothing Then Exit Sub
        If grsRptPara Is Nothing Then Exit Sub
        
        grsRptPara.Filter = "���= '" & strCode & "'"
        
        .Redraw = flexRDNone
        
        Do While Not grsRptPara.EOF
            .rows = .rows + 1
            
            .TextMatrix(.rows - 1, .ColIndex("���")) = grsRptPara!���
            .TextMatrix(.rows - 1, .ColIndex("Դ����")) = grsRptPara!����Դ����
            .TextMatrix(.rows - 1, .ColIndex("���")) = grsRptPara!�������
            .TextMatrix(.rows - 1, .ColIndex("��������")) = grsRptPara!��������
            
            grsRptPara.MoveNext
        Loop
        
        .MergeCells = flexMergeRestrictColumns
        .MergeCol(.ColIndex("Դ����")) = True
        
        .Redraw = flexRDDirect
        
        '�������Ƿ���ϸù�����������Ҫ����ֻ�С�ҩ�����������ݡ�����NO������������
        For n = 1 To .rows - 1
            If Trim(.TextMatrix(n, .ColIndex("��������"))) = "ҩ��" Then
                blnCheckҩ�� = True
            End If
            
            If Trim(.TextMatrix(n, .ColIndex("��������"))) = "����" Then
                blnCheck���� = True
            End If
            
            If Trim(.TextMatrix(n, .ColIndex("��������"))) = "NO" Then
                blnCheckNO = True
            End If
            
            '���������������
            If blnCheckOther = False Then
                If InStr(1, ",ҩ��,����,NO,", "," & .TextMatrix(n, .ColIndex("��������")) & ",") = 0 Then
                    blnCheckOther = True
                End If
            End If
        Next
        
        If blnCheckҩ�� = False Or blnCheck���� = False Or blnCheckNO = False Or blnCheckOther = True Then
            lblComment.Caption = "��ʾ��"
            
            If blnCheckҩ�� = False Or blnCheck���� = False Or blnCheckNO = False Then
                lblComment.Caption = lblComment.Caption & "ȱ��"
                
                If blnCheckҩ�� = False Then
                    lblComment.Caption = lblComment.Caption & "��ҩ����"
                End If
                
                If blnCheck���� = False Then
                    lblComment.Caption = lblComment.Caption & IIf(blnCheckҩ�� = False, "��", "") & "�����ݡ�"
                End If
                
                If blnCheckNO = False Then
                    lblComment.Caption = lblComment.Caption & IIf(blnCheckҩ�� = False Or blnCheck���� = False, "��", "") & "��NO��"
                End If
            
                lblComment.Caption = lblComment.Caption & "����"
            End If
            
            If blnCheckOther = True Then
                If blnCheckҩ�� = False Or blnCheck���� = False Or blnCheckNO = False Then
                    lblComment.Caption = lblComment.Caption & "�����Ҳ�������������"
                Else
                    lblComment.Caption = lblComment.Caption & "�����С�ҩ�����������ݡ�����NO�������������"
                End If
            End If
            
            CmdOK.Enabled = False
        Else
            CmdOK.Enabled = True
        End If
    End With
End Sub

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdOK_Click()
    If vsfRpt.Row = 0 Then Exit Sub
    If vsfRpt.TextMatrix(vsfRpt.Row, vsfRpt.ColIndex("���")) = "" Then Exit Sub
    
    mstrRpt = vsfRpt.TextMatrix(vsfRpt.Row, vsfRpt.ColIndex("���")) & "," & vsfRpt.TextMatrix(vsfRpt.Row, vsfRpt.ColIndex("����"))
    
    Unload Me
End Sub


Private Sub Form_Load()
    If grsRpt Is Nothing Then
        Set grsRpt = New ADODB.Recordset
        With grsRpt
            If .State = 1 Then .Close
            .Fields.Append "���", adLongVarChar, 40, adFldIsNullable
            .Fields.Append "����", adLongVarChar, 50, adFldIsNullable
            
            .CursorLocation = adUseClient
            .CursorType = adOpenStatic
            .LockType = adLockOptimistic
            .Open
        End With
        
        gstrSQL = "Select ID, ���, ���� From zlReports Where ϵͳ = 100 Order By ��� "
        Set grsRpt = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ���б���")
    End If
    
    If grsRptPara Is Nothing Then
        Set grsRptPara = New ADODB.Recordset
        With grsRptPara
            If .State = 1 Then .Close
            .Fields.Append "���", adLongVarChar, 40, adFldIsNullable
            .Fields.Append "����Դ����", adLongVarChar, 40, adFldIsNullable
            .Fields.Append "���", adDouble, 2, adFldIsNullable
            .Fields.Append "��������", adLongVarChar, 60, adFldIsNullable
            
            .CursorLocation = adUseClient
            .CursorType = adOpenStatic
            .LockType = adLockOptimistic
            .Open
        End With
        
        gstrSQL = "Select c.���, b.���� As ����Դ����, a.��� As �������, a.���� As �������� " & _
            " From zlRPTPars A, zlRPTDatas B, zlReports C " & _
            " Where c.Id = b.����id And b.Id = a.Դid " & _
            " Order By c.���, b.����, a.��� "
        Set grsRptPara = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ���б������")
    End If
    
    Call ShowRpt
End Sub

Private Sub txt����_Change()
    Call ShowRpt(Trim(txt����.Text))
End Sub

Private Sub vsfRpt_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If OldRow = NewRow Then Exit Sub
    
    With vsfRpt
        If .TextMatrix(NewRow, 0) = "" Then Exit Sub
        
        Call ShowRptPara(.TextMatrix(NewRow, .ColIndex("���")))
    End With
End Sub

