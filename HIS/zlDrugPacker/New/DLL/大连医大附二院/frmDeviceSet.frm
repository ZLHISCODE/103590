VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmDeviceSet 
   Caption         =   "�豸��Ϣ����"
   ClientHeight    =   6420
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11310
   Icon            =   "frmDeviceSet.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6420
   ScaleWidth      =   11310
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picParam 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   2175
      Left            =   120
      ScaleHeight     =   2175
      ScaleWidth      =   11175
      TabIndex        =   3
      Top             =   4200
      Width           =   11175
      Begin VB.CommandButton cmdSetParam 
         Height          =   360
         Left            =   10560
         Picture         =   "frmDeviceSet.frx":6852
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1680
         Width           =   390
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfDeviceParam 
         Height          =   1245
         Left            =   30
         TabIndex        =   4
         Top             =   360
         Width           =   10995
         _cx             =   19394
         _cy             =   2196
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
         BackColorSel    =   16764622
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483634
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmDeviceSet.frx":D0A4
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
         ExplorerBar     =   5
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
      Begin VB.Label lblComment 
         BackColor       =   &H00808080&
         Caption         =   "Ӧ�ò�����ѡ���豸���������ð�ť���á��ı������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   25
         TabIndex        =   9
         Top             =   25
         Width           =   7695
      End
   End
   Begin VB.PictureBox picBase 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   3375
      Left            =   120
      ScaleHeight     =   3375
      ScaleWidth      =   11175
      TabIndex        =   1
      Top             =   720
      Width           =   11175
      Begin VB.CommandButton cmdExit 
         Cancel          =   -1  'True
         Caption         =   "�˳�(&X)"
         Height          =   360
         Left            =   9960
         TabIndex        =   12
         Top             =   2880
         Width           =   990
      End
      Begin VB.CommandButton cmdDel 
         Height          =   360
         Left            =   2415
         Picture         =   "frmDeviceSet.frx":D19F
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "ɾ��"
         Top             =   2880
         Width           =   375
      End
      Begin VB.CommandButton cmdAdd 
         Height          =   360
         Left            =   1650
         Picture         =   "frmDeviceSet.frx":139F1
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "����"
         Top             =   2880
         Width           =   390
      End
      Begin VB.CommandButton cmdEdit 
         Height          =   360
         Left            =   2040
         Picture         =   "frmDeviceSet.frx":1A243
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "�޸�"
         Top             =   2880
         Width           =   390
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfDeviceInfo 
         Height          =   2445
         Left            =   30
         TabIndex        =   2
         Top             =   360
         Width           =   10995
         _cx             =   19394
         _cy             =   4313
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
         BackColorSel    =   16764622
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483634
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   4
         Cols            =   11
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmDeviceSet.frx":20A95
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
         ExplorerBar     =   5
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
      Begin VB.Label lblComment 
         BackColor       =   &H00808080&
         Caption         =   "������Ϣ�����������ť�������豸��˫������л����༭��ť���б༭��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   25
         TabIndex        =   6
         Top             =   30
         Width           =   7695
      End
   End
   Begin VB.Frame fraLine1 
      Height          =   75
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   9135
   End
   Begin VB.Label lblComment 
      Caption         =   "���������úͲο��豸�Ļ�����Ϣ��Ӧ�ò���"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   720
      TabIndex        =   5
      Top             =   240
      Width           =   7695
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmDeviceSet.frx":20D58
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "frmDeviceSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnFirst As Boolean                '�״ν��봰��

Private Sub GetDeviceInfo()
    Dim rsData As ADODB.Recordset
    
    vsfDeviceInfo.Rows = 1
    
    On Error GoTo errHandle
    
    gstrSQL = "Select a.Id, a.����, a.����, a.�ͺ�, a.������, a.ʹ�ò���id, '��' || b.���� || '��' || b.���� As ʹ�ò���, " & _
        " Decode(a.��������, 1, '���ݿ�', 2, 'WebService', 3, '����Ŀ¼', 'δ֪') As ��������, a.��������, a.�Ƿ����� " & _
        " From ҩ����ҩ�豸 A, ���ű� B " & _
        " Where a.ʹ�ò���id = b.Id " & _
        " Order By a.���� "
    Set rsData = gobjComLib.zldatabase.OpenSQLRecord(gstrSQL, "GetDevice")
    
    Do While Not rsData.EOF
        With vsfDeviceInfo
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, .ColIndex("�豸id")) = rsData!ID
            .TextMatrix(.Rows - 1, .ColIndex("����")) = rsData!����
            .TextMatrix(.Rows - 1, .ColIndex("����")) = rsData!����
            .TextMatrix(.Rows - 1, .ColIndex("�ͺ�")) = gobjComLib.zlcommfun.NVL(rsData!�ͺ�)
            .TextMatrix(.Rows - 1, .ColIndex("������")) = gobjComLib.zlcommfun.NVL(rsData!������)
            .TextMatrix(.Rows - 1, .ColIndex("ʹ�ò���")) = rsData!ʹ�ò���
            .TextMatrix(.Rows - 1, .ColIndex("��������")) = rsData!��������
            .TextMatrix(.Rows - 1, .ColIndex("��������")) = gobjComLib.zlcommfun.NVL(rsData!��������)
            .TextMatrix(.Rows - 1, .ColIndex("����")) = gobjComLib.zlcommfun.NVL(rsData!�Ƿ�����, 0)
        End With
                
        rsData.MoveNext
    Loop
    
    If vsfDeviceInfo.Rows = 1 Then vsfDeviceInfo.Rows = 2
    
    vsfDeviceInfo.Row = 1
    
    Exit Sub
errHandle:
    If gobjComLib.ErrCenter = 1 Then Resume
    gstrMessage = Err.Description
End Sub

Private Sub GetDeviceParam(ByVal lng�豸id As Long)
    Dim rsData As ADODB.Recordset
    
    vsfDeviceParam.Rows = 1
    
    On Error GoTo errHandle
    
    gstrSQL = "Select a.����id, a.�豸id, a.����ֵ, b.������, b.������, b.����˵��, b.ȱʡֵ" & vbNewLine & _
        " From ҩ���豸���� A, �Զ���ҩ���� B" & vbNewLine & _
        " Where a.����id = b.Id and a.�豸id=[1] "

    Set rsData = gobjComLib.zldatabase.OpenSQLRecord(gstrSQL, "GetDeviceParam", lng�豸id)
    
    Do While Not rsData.EOF
        With vsfDeviceParam
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, .ColIndex("����id")) = rsData!����id
            .TextMatrix(.Rows - 1, .ColIndex("������")) = rsData!������
            .TextMatrix(.Rows - 1, .ColIndex("����ֵ")) = gobjComLib.zlcommfun.NVL(rsData!����ֵ)
            .TextMatrix(.Rows - 1, .ColIndex("����˵��")) = gobjComLib.zlcommfun.NVL(rsData!����˵��)
        End With
                
        rsData.MoveNext
    Loop
    
    If vsfDeviceParam.Rows = 1 Then vsfDeviceParam.Rows = 2
    
    vsfDeviceParam.Row = 1
    
    Exit Sub
    
errHandle:
    If gobjComLib.ErrCenter = 1 Then Resume
    gstrMessage = Err.Description
End Sub


Private Sub cmdAdd_Click()
    frmDeviceBase.ShowMe Me, 0, 0
    
    Call GetDeviceInfo
    
    vsfDeviceInfo.Row = vsfDeviceInfo.Rows - 1
End Sub

Private Sub cmdDel_Click()
    With vsfDeviceInfo
        If Val(.TextMatrix(.Row, .ColIndex("�豸id"))) > 0 Then
            If MsgBox("�Ƿ�ɾ����" & .TextMatrix(.Row, .ColIndex("����")) & "���豸��", vbInformation + vbYesNo + vbDefaultButton2, GSTR_INTERFACE_NAME) = vbNo Then Exit Sub
            
            gstrSQL = "Zl_ҩ����ҩ�豸_Delete(" & Val(.TextMatrix(.Row, .ColIndex("�豸id"))) & ")"
            Call gobjComLib.zldatabase.ExecuteProcedure(gstrSQL, "ҩ��ע���豸-�޸�")
             
            .RemoveItem (.Row)
            If .Rows = 1 Then
                .Rows = 2
            End If
            Call vsfDeviceInfo_EnterCell
        End If
    End With
End Sub

Private Sub cmdEdit_Click()
    If vsfDeviceInfo.Row > 0 Then
        If Val(vsfDeviceInfo.TextMatrix(vsfDeviceInfo.Row, vsfDeviceInfo.ColIndex("�豸id"))) > 0 Then
            Dim i As Integer
            i = vsfDeviceInfo.Row
            frmDeviceBase.ShowMe Me, Val(vsfDeviceInfo.TextMatrix(vsfDeviceInfo.Row, vsfDeviceInfo.ColIndex("�豸id"))), 1
            Call GetDeviceInfo
            vsfDeviceInfo.Row = i
        End If
    End If
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSetParam_Click()
    If vsfDeviceInfo.Row > 0 Then
        If Val(vsfDeviceInfo.TextMatrix(vsfDeviceInfo.Row, vsfDeviceInfo.ColIndex("�豸id"))) > 0 Then
            frmDeviceParam.ShowMeByDevice Me, Val(vsfDeviceInfo.TextMatrix(vsfDeviceInfo.Row, vsfDeviceInfo.ColIndex("�豸id")))
            
            Call GetDeviceParam(Val(vsfDeviceInfo.TextMatrix(vsfDeviceInfo.Row, vsfDeviceInfo.ColIndex("�豸id"))))
        End If
    End If
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    
    Call GetDeviceInfo
    
    mblnFirst = False
End Sub

Private Sub Form_Load()
    mblnFirst = True
    
    '���أ������ؼ�
    picParam.Visible = False
    
End Sub

Private Sub Form_Resize()
    Const INT_PARAM_AREA = 0             '��������̶��߶�

    If Me.Width < 8000 Then
        Me.Width = 8000
        Exit Sub
    End If
    If Me.Height < 6000 Then
        Me.Height = 6000
        Exit Sub
    End If

    With cmdExit
        .Top = lblComment(0).Top
        .Left = Me.ScaleWidth - .Width - 200
    End With
    
    fraLine1(0).Width = Me.ScaleWidth

    With picBase
        .Top = fraLine1(0).Top + fraLine1(0).Height
        .Left = 0
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight - .Top - INT_PARAM_AREA
    End With
    
    With lblComment(1)
        .Top = 50
        .Left = 100
        .Width = picBase.ScaleWidth - 200
    End With
    
    With vsfDeviceInfo
        .Top = lblComment(1).Top + lblComment(1).Height
        .Left = lblComment(1).Left
        .Width = lblComment(1).Width
        .Height = picBase.ScaleHeight - cmdDel.Height - lblComment(1).Height - 200 - 50
    End With
    
    With cmdAdd
        .Top = vsfDeviceInfo.Top + vsfDeviceInfo.Height + 100
        .Left = vsfDeviceInfo.Left
    End With
    
    With cmdEdit
        .Top = cmdAdd.Top
        .Left = cmdAdd.Left + cmdAdd.Width
    End With
    
    With cmdDel
        .Top = cmdAdd.Top
        .Left = cmdEdit.Left + cmdEdit.Width
    End With
    
    With cmdExit
        .Top = cmdAdd.Top
        .Left = vsfDeviceInfo.Width + vsfDeviceInfo.Left - .Width
    End With
    
    If picParam.Visible = False Then Exit Sub
    
    '���´��벻ִ��
    
    With picParam
        .Top = picBase.Top + picBase.Height
        .Left = 0
        .Width = Me.ScaleWidth
        .Height = INT_PARAM_AREA
    End With
    
    With lblComment(2)
        .Top = 50
        .Left = lblComment(1).Left
        .Width = lblComment(1).Width
    End With
    
    With vsfDeviceParam
        .Top = lblComment(2).Top + lblComment(2).Height
        .Left = lblComment(1).Left
        .Width = lblComment(1).Width
        .Height = picParam.ScaleHeight - lblComment(2).Height - cmdSetParam.Height - 200 - 50
    End With
    
    With cmdSetParam
        .Top = vsfDeviceParam.Top + vsfDeviceParam.Height + 100
        .Left = lblComment(2).Left + lblComment(2).Width - .Width
    End With
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '
End Sub

Private Sub vsfDeviceInfo_EnterCell()
    If picParam.Visible = False Then Exit Sub
    
    With vsfDeviceInfo
        If Val(.TextMatrix(.Row, .ColIndex("�豸id"))) > 0 Then
            Call GetDeviceParam(Val(.TextMatrix(.Row, .ColIndex("�豸id"))))
        Else
            vsfDeviceParam.Clear 1
            vsfDeviceParam.Rows = 2
        End If
    End With
End Sub

