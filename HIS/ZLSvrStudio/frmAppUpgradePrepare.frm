VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "VSFLEX8.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.Unicode.9600.ocx"
Begin VB.Form frmAppUpgradePrepare 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "����׼��"
   ClientHeight    =   8445
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12015
   Icon            =   "frmAppUpgradePrepare.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8445
   ScaleWidth      =   12015
   StartUpPosition =   2  '��Ļ����
   Begin VB.Frame fraLine 
      Height          =   120
      Left            =   5400
      TabIndex        =   9
      Top             =   600
      Width           =   4935
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6960
      Width           =   1100
   End
   Begin VB.CommandButton cmdExec 
      Caption         =   "ִ��(&E)"
      Height          =   350
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6960
      Width           =   1100
   End
   Begin VB.Frame fraCheck 
      BorderStyle     =   0  'None
      Height          =   2415
      Left            =   4560
      TabIndex        =   3
      Top             =   1080
      Width           =   3735
      Begin VSFlex8Ctl.VSFlexGrid vsfShow 
         Height          =   1215
         Index           =   3
         Left            =   360
         TabIndex        =   17
         Top             =   360
         Width           =   2895
         _cx             =   5106
         _cy             =   2143
         Appearance      =   1
         BorderStyle     =   0
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   16777215
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   -1  'True
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
      Begin VB.Label lblCheck 
         AutoSize        =   -1  'True
         Caption         =   "˵����"
         Height          =   180
         Left            =   840
         TabIndex        =   11
         Top             =   2040
         Width           =   540
      End
   End
   Begin VB.Frame fraJob 
      BorderStyle     =   0  'None
      Height          =   2175
      Left            =   8640
      TabIndex        =   2
      Top             =   1200
      Width           =   3495
      Begin VB.CheckBox chkShow 
         Caption         =   "��ʾͣ�õĺ�̨��ҵ"
         Height          =   495
         Index           =   2
         Left            =   120
         TabIndex        =   19
         Top             =   120
         Width           =   2175
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfShow 
         Height          =   855
         Index           =   2
         Left            =   720
         TabIndex        =   16
         Top             =   600
         Width           =   1575
         _cx             =   2778
         _cy             =   1508
         Appearance      =   1
         BorderStyle     =   0
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   16777215
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   -1  'True
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
      Begin VB.Label lblJob 
         AutoSize        =   -1  'True
         Caption         =   "˵����"
         Height          =   180
         Left            =   120
         TabIndex        =   12
         Top             =   1440
         Width           =   540
      End
   End
   Begin VB.Frame fraUser 
      BorderStyle     =   0  'None
      Height          =   2415
      Left            =   6600
      TabIndex        =   1
      Top             =   4080
      Width           =   4695
      Begin VB.CheckBox chkShow 
         Caption         =   "��ʾͣ�õ��û�"
         Height          =   495
         Index           =   1
         Left            =   960
         TabIndex        =   20
         Top             =   1800
         Width           =   2175
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfShow 
         Height          =   975
         Index           =   1
         Left            =   240
         TabIndex        =   15
         Top             =   360
         Width           =   2775
         _cx             =   4895
         _cy             =   1720
         Appearance      =   1
         BorderStyle     =   0
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   -1  'True
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
      Begin VB.Label lblUser 
         AutoSize        =   -1  'True
         Caption         =   "˵����"
         Height          =   180
         Left            =   240
         TabIndex        =   13
         Top             =   1800
         Width           =   540
      End
   End
   Begin VB.Frame fraClient 
      BorderStyle     =   0  'None
      Height          =   3015
      Left            =   240
      TabIndex        =   0
      Top             =   3840
      Width           =   4935
      Begin VB.CheckBox chkShow 
         Caption         =   "��ʾͣ�õĿͻ���"
         Height          =   495
         Index           =   0
         Left            =   240
         TabIndex        =   18
         Top             =   120
         Width           =   2175
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfShow 
         Height          =   1215
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   2295
         _cx             =   4048
         _cy             =   2143
         Appearance      =   1
         BorderStyle     =   0
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   16777215
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   -1  'True
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
      Begin VB.CommandButton cmdkillProcess 
         Caption         =   "�жϿͻ������ӵĽ��̶���(&P)"
         Height          =   350
         Left            =   360
         TabIndex        =   5
         Top             =   1920
         Width           =   2790
      End
      Begin VB.Label lblClient 
         AutoSize        =   -1  'True
         Caption         =   "˵����"
         Height          =   180
         Left            =   120
         TabIndex        =   10
         Top             =   2520
         Width           =   540
      End
   End
   Begin XtremeSuiteControls.TabControl tbcMain 
      Height          =   2415
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   4335
      _Version        =   589884
      _ExtentX        =   7646
      _ExtentY        =   4260
      _StockProps     =   64
   End
   Begin VB.Label lblResult 
      AutoSize        =   -1  'True
      Caption         =   "�����"
      Height          =   180
      Left            =   360
      TabIndex        =   21
      Top             =   7560
      Width           =   540
   End
   Begin VB.Image imgMain 
      Height          =   615
      Left            =   120
      Picture         =   "frmAppUpgradePrepare.frx":6852
      Stretch         =   -1  'True
      Top             =   120
      Width           =   615
   End
   Begin VB.Label lblTip 
      AutoSize        =   -1  'True
      Caption         =   "˵��"
      Height          =   180
      Left            =   840
      TabIndex        =   8
      Top             =   120
      Width           =   360
   End
End
Attribute VB_Name = "frmAppUpgradePrepare"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const SQL_CAPTION = "����ǰ�ü��"
Private Const mstrOracleUser      As String = "'ANONYMOUS','AURORA$JIS$UTILITY$','AURORA$ORB$UNAUTHENTICATED','CTXSYS','DBSNMP','DIP','DMSYS','DVF','DVSYS','EXFSYS','HR','LBACSYS','MDDATA','MDSYS','MGMT_VIEW','OAS_PUBLIC','ODM','ODM_MTR','OE','OGG','OLAPSYS','ORDPLUGINS','ORDSYS','OSE$HTTP$ADMIN','OUTLN','PERFSTAT','PM','QS','QS_ADM','QS_CB','QS_CBADM','QS_CS','QS_ES','QS_OS','QS_WS','REPADMIN','RMAN','SCOTT','SH','SI_INFORMTN_SCHEMA','SYSMAN','TRACESVR','TSMSYS','WEBSYS','WKPROXY','WKSYS','WKUSER','WK_TEST','WMSYS','XDB'"
Private Enum enuTab
    T_�ͻ��� = 0
    T_�û��˺�
    T_��̨��ҵ
    T_����
End Enum
Private mstrSysNum As String
Private mstrUsers As String
Private mbln10g As Boolean
Private mblnFirst As Boolean
Private mstrCondition As String
Private mcnExe As ADODB.Connection
'�ô�����չʾϵͳ����ǰ��һЩ׼����Ϊ�˲�Ӱ�����������������Դ���
Private Sub InitTabControl()
    
    With tbcMain
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .BoldSelected = True
            .BoldSelected = True
            .Color = xtpTabColorDefault
            .ShowIcons = False
        End With
        Call .InsertItem(T_�ͻ���, "�жϿͻ������Ӳ���ֹ��¼", fraClient.hwnd, 0)
        Call .InsertItem(T_�û��˺�, "�����û��˺�", fraUser.hwnd, 0)
        Call .InsertItem(T_��̨��ҵ, "���ú�̨��ҵ", fraJob.hwnd, 0)
        Call .InsertItem(T_����, "Ӱ������Ч�ʵ���������", fraCheck.hwnd, 0)
    End With
End Sub

Private Sub chkShow_Click(Index As Integer)
    Select Case Index
        Case T_�ͻ���
            Call LoadClient
        Case T_�û��˺�
            mblnFirst = True
            Call LoadUser
        Case T_��̨��ҵ
            Call LoadJob
    End Select
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdExec_Click()
    Dim i As Long, j As Long, lngNum As Long, lngInstId As Long
    Dim rsChoose As ADODB.Recordset
    Dim varSQL As Variant
    Dim strErrContent As String, strErr As String, strTemp As String
    Dim strClient As String
    Dim cnTemp As ADODB.Connection

    If ExeCheck = False Then Exit Sub
    On Error GoTo ErrH
    Set rsChoose = CopyNewRec(Nothing, True, , Array("����", adVarChar, 10, Empty, "ͣ��SQL", adVarChar, 500, Empty, _
                    "����", adVarChar, 200, Empty))
    For i = vsfShow.LBound To vsfShow.UBound
        With vsfShow(i)
            For j = 1 To .Rows - 1
                If .Cell(flexcpChecked, j, .ColIndex("ѡ��")) = flexChecked Then
                    If i = T_�ͻ��� Then
                        If InStr(strClient, "'" & .TextMatrix(j, .ColIndex("�ͻ���")) & "'") = 0 Then
                            strClient = IIf(strClient = "", "", strClient & ",") & "'" & .TextMatrix(j, .ColIndex("�ͻ���")) & "'"
                        End If
                        rsChoose.AddNew Array("����", "ͣ��SQL", "����"), Array("�ͻ���", .TextMatrix(j, .ColIndex("ͣ��SQL")), _
                        .TextMatrix(j, .ColIndex("INST_ID")) & "|" & .TextMatrix(j, .ColIndex("��ǰ��־")))
                    ElseIf i = T_�û��˺� Then
                        rsChoose.AddNew Array("����", "ͣ��SQL"), Array("�û��˺�", .TextMatrix(j, .ColIndex("ͣ��SQL")))
                    ElseIf i = T_��̨��ҵ Then
                        If .TextMatrix(j, .ColIndex("����")) = "ϵͳ����" Then
                            rsChoose.AddNew Array("����", "ͣ��SQL", "����"), Array("ϵͳ����", .TextMatrix(j, .ColIndex("ͣ��SQL")), .TextMatrix(j, .ColIndex("����")))
                        '�ǲ�Ʒ�Զ���ҵ���¼��zlUpgradeConfig�У��ʽ���ҵ�ŷ��� ��������
                        ElseIf .TextMatrix(j, .ColIndex("����")) = "�ǲ�Ʒ�Զ���ҵ" Then
                            rsChoose.AddNew Array("����", "ͣ��SQL", "����"), Array("��̨��ҵ", .TextMatrix(j, .ColIndex("ͣ��SQL")), .TextMatrix(j, .ColIndex("��ҵ��")))
                        Else
                            rsChoose.AddNew Array("����", "ͣ��SQL"), Array("��̨��ҵ", .TextMatrix(j, .ColIndex("ͣ��SQL")))
                        End If
                    ElseIf i = T_���� Then
                        If .TextMatrix(j, .ColIndex("����")) = "������" Then
                            rsChoose.AddNew Array("����", "ͣ��SQL", "����"), Array("������", .TextMatrix(j, .ColIndex("ͣ��SQL")), .TextMatrix(j, .ColIndex("����")))
                        Else
                            rsChoose.AddNew Array("����", "ͣ��SQL"), Array("��������", .TextMatrix(j, .ColIndex("ͣ��SQL")))
                        End If
                    End If
                End If
            Next
        End With
    Next
    If rsChoose.RecordCount = 0 Then Unload Me: Exit Sub
    If MsgBox("ȷ��Ҫִ�й�ѡ����Щ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
        Exit Sub
    End If
    Call ShowFlash("����ִ������ѡ���������SQL,���Ժ�...")
    '�����û��˺�
    rsChoose.Filter = "����='�û��˺�'"
    Do While Not rsChoose.EOF
        strErrContent = ""
        varSQL = Split(rsChoose!ͣ��SQL, "�ָ���")
        For i = LBound(varSQL) To UBound(varSQL)
            strErrContent = strErrContent & gclsBase.ExecuteCmdText(varSQL(i), Me.Caption, mcnExe, True)
            '����˺�����ʧ�ܣ��򲻸ı�����ͣ�ñ��ֵ
            If i = 0 And strErrContent <> "" Then Exit For
        Next
        If strErrContent <> "" Then
            lngNum = lngNum + 1
        End If
        rsChoose.MoveNext
    Loop
    strErr = IIf(lngNum = 0, strErr, strErr & "�����û��˺�ʧ��" & lngNum & "��;")
    lngNum = 0
    '����ϵͳ����
    strTemp = ""
    rsChoose.Filter = "����='ϵͳ����'"
    If mbln10g Then
        Call ShowFlash("")
        Set cnTemp = GetConnection("SYS")
        Call ShowFlash("����ִ������ѡ���������SQL,���Ժ�...")
    Else
        Set cnTemp = mcnExe
    End If
    Do While Not rsChoose.EOF
        strErrContent = ""
        strErrContent = gclsBase.ExecuteCmdText(rsChoose!ͣ��SQL, Me.Caption, cnTemp, True)
        If strErrContent = "" Then
            strTemp = IIf(strTemp = "", "", strTemp & ",") & rsChoose!����
        Else
            lngNum = lngNum + 1
        End If
        rsChoose.MoveNext
    Loop
    If strTemp <> "" Then
        gstrSQL = "Update Zlupgradeconfig Set ����='" & strTemp & "' Where ��Ŀ='���õ�ϵͳ����'"
        Call gclsBase.ExecuteCmdText(gstrSQL, Me.Caption, mcnExe)
    End If
    strTemp = ""
    '���ú�̨��ҵ
    rsChoose.Filter = "����='��̨��ҵ'"
    Do While Not rsChoose.EOF
        strErrContent = ""
        varSQL = Split(rsChoose!ͣ��SQL, "�ָ���")
        For i = LBound(varSQL) To UBound(varSQL)
            On Error Resume Next
            '��̨��ҵ������adCmdText����ִ��
            gcnOracle.Execute varSQL(i)
            '��̨��ҵͣ��ʧ�ܣ����ı�����ͣ�ñ��ֵ
            If i = 0 And err.Number <> 0 Then Exit For
        Next
        If err.Number = 0 Then
            '�����ƶ��ڷǲ�Ʒ�Զ���ҵ�����������ҵ�ţ���������
            If "" & rsChoose!���� <> "" Then
                strTemp = IIf(strTemp = "", "", strTemp & ",") & rsChoose!����
            End If
        Else
            err.Clear
            lngNum = lngNum + 1
        End If
        rsChoose.MoveNext
    Loop
    On Error GoTo ErrH
    If strTemp <> "" Then
        '�ǲ�Ʒ��̨��ҵ�赥��������Zlupgradeconfig��
        gstrSQL = "Update Zlupgradeconfig Set ����='" & strTemp & "' Where ��Ŀ='���õĺ�̨��ҵ'"
        Call gclsBase.ExecuteCmdText(gstrSQL, Me.Caption, mcnExe)
    End If
    strErr = IIf(lngNum = 0, strErr, strErr & "ͣ�ú�̨��ҵʧ��" & lngNum & "��;")
    lngNum = 0
    '���ô�����,ֻ�ܱ�������������
    rsChoose.Filter = "����='������'"
    Do While Not rsChoose.EOF
        strErrContent = ""
        varSQL = Split(rsChoose!����, ".")
        If varSQL(0) = UCase(gstrUserName) Then
            Set cnTemp = gcnOracle
        ElseIf varSQL(0) = "ZLTOOLS" Then
            Call ShowFlash("")
            Set cnTemp = GetConnection("ZLTOOLS")
            Call ShowFlash("����ִ������ѡ���������SQL,���Ժ�...")
        Else
            Call ShowFlash("")
            Set cnTemp = GetConnection(varSQL(0))
            Call ShowFlash("����ִ������ѡ���������SQL,���Ժ�...")
        End If
        strErrContent = strErrContent & gclsBase.ExecuteCmdText(rsChoose!ͣ��SQL, Me.Caption, cnTemp, True)
        If strErrContent = "" Then
            gstrSQL = "Insert Into Zltriggers(������, ����) Values ('" & varSQL(0) & "','" & varSQL(1) & "')"
            Call gclsBase.ExecuteCmdText(gstrSQL, Me.Caption, mcnExe)
        Else
            lngNum = lngNum + 1
        End If
        rsChoose.MoveNext
    Loop
    '������������
    rsChoose.Filter = "����='��������'"
    Do While Not rsChoose.EOF
        Call ExeSQL(mcnExe, rsChoose!ͣ��SQL, lngNum)
        rsChoose.MoveNext
    Loop
    strErr = IIf(lngNum = 0, strErr, strErr & "����Ӱ������ִ��Ч�ʵĲ�����ʧ��" & lngNum & "��;")
    lngNum = 0
    '���ÿͻ��˲�ɱ���Ự
    rsChoose.Filter = "����='�ͻ���'"
    If rsChoose.RecordCount > 0 Then
        rsChoose.Sort = "���� desc"
        gstrSQL = "Update Zlclients Set ��ֹʹ�� = 1, ϵͳ�������� = 1 Where Nvl(��ֹʹ��, 0) = 0 and Nvl(ϵͳ��������, 0) = 0 and ����վ in(" & strClient & ")"
        Call gclsBase.ExecuteCmdText(gstrSQL, Me.Caption, mcnExe, True)
        If mbln10g Then
            Do While Not rsChoose.EOF
                varSQL = Split(rsChoose!����, "|")
                '10g���������ҵ�ǰ��־Ϊ0
                If varSQL(1) = 0 Then
                    If lngInstId <> varSQL(0) Then
                        lngInstId = varSQL(0)
                        Set cnTemp = GetInstance(lngInstId)
                        Call ExeSQL(cnTemp, rsChoose!ͣ��SQL, lngNum)
                    Else
                        Call ExeSQL(cnTemp, rsChoose!ͣ��SQL, lngNum)
                    End If
                Else
                    Call ExeSQL(mcnExe, rsChoose!ͣ��SQL, lngNum)
                End If
                rsChoose.MoveNext
            Loop
        Else
            Do While Not rsChoose.EOF
                Call ExeSQL(mcnExe, rsChoose!ͣ��SQL, lngNum)
                rsChoose.MoveNext
            Loop
        End If
    End If
    strErr = IIf(lngNum = 0, strErr, strErr & "���ÿͻ��˲�ɱ���Ựʧ��" & lngNum & "��;")
    Call ShowFlash
    If strErr <> "" Then
        MsgBox strErr, vbExclamation, Me.Caption
    End If
    Unload Me
    Exit Sub
ErrH:
    Call ShowFlash
    MsgBox err.Description, vbExclamation, Me.Caption
End Sub

Private Function ExeCheck() As Boolean
'ִ��ǰ�ļ��
    
    If Not CheckAndAdjustMustTable("zlUpgradeConfig", , False) Then
        MsgBox "�޷��Զ��ؽ�zlUpgradeConfig�����ֹ�ɱ���Ự,֮����ȷ�����ɣ�", vbInformation, gstrSysName
        Exit Function
    End If
    If Not CheckAndAdjustMustTable("Zlclients", "ϵͳ��������", False) Then
        MsgBox "�޷��Զ�����Zlclients�����ֹ�ɱ���Ự,֮����ȷ�����ɣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    ExeCheck = True
End Function

Private Sub ExeSQL(ByVal cnExe As ADODB.Connection, ByVal strSQL As String, ByRef lngNum As Long)
    Dim strErr As String
    
    strErr = gclsBase.ExecuteCmdText(strSQL, Me.Caption, cnExe, True)
    If strErr <> "" Then
        lngNum = lngNum + 1
    End If
End Sub

Private Function GetInstance(ByVal lngInstId As Long) As ADODB.Connection
'����INST_ID��ȡʵ������
    Dim rsTemp As ADODB.Recordset
    Dim cnTemp As ADODB.Connection
    Dim strTemp As String
    
    On Error GoTo ErrH
    gstrSQL = "select a.inst_ID, a.Instance_Name, a.Host_name, b.NAME, b.DBID" & vbNewLine & _
            "  from gv$instance a, gv$database b" & vbNewLine & _
            " where a.INST_ID = b.INST_ID" & vbNewLine & _
            "   and a.INST_ID <> userenv('instance')" & vbNewLine & _
            "   and a.STATUS = 'OPEN' and a.INST_ID=[1]"
    Set rsTemp = gclsBase.OpenSQLRecord(mcnExe, gstrSQL, "��ȡʵ����Ϣ", lngInstId)
    If Not rsTemp.EOF Then
        strTemp = rsTemp!INST_ID & "," & rsTemp!DBID & "," & rsTemp!Instance_Name & "(" & rsTemp!name & ")"
        If frmUserCheckLogin.ShowLogin(UCT_RACInsUser, cnTemp, gstrUserName, "", "", strTemp) Then
            Set GetInstance = cnTemp
        Else
            Set GetInstance = Nothing
        End If
    End If
    Exit Function
ErrH:
    MsgBox err.Description, vbExclamation, "��ȡʵ������"
End Function

Private Sub cmdkillProcess_Click()
    frmKillProcessManage.ShowMe ("0102")
    Call LoadClient
End Sub

Private Sub Form_Load()
    Dim strHead As String
    
    mblnFirst = True
    mstrCondition = ""
    Call ShowFlash("���ڼ�������׼����Ϣ�����Ժ�")
    mbln10g = GetOracleVersion(True, True) < 11
    Call InitTabControl
    lblTip.Caption = "Ϊ�˱�������������Чִ�У�����ǰӦ���ÿͻ��ˣ������û��˺ţ�ͣ�ú�̨��ҵ�������������Ż�����������ݿ������" & vbNewLine & "������ɺ��Զ��ָ����õ���Ŀ����������쳣�жϣ�����������������ִ�в������ָ�����׼���ڼ��������Ŀ����"
    strHead = " ,300,1;�ͻ���,1500,1;IP,1500,1;����,1500,1;Ժ��,1200,1;��;,1000,1;SID,500,1;SERIAL#,800,1;PROGRAM,2000,1;״̬,500,1;INST_ID,0,1;��ǰ��־,0,1;ͣ��SQL,0,1"
    Call iniVsf(strHead, vsfShow(T_�ͻ���))
    strHead = " ,300,1;����,1800,1;�û���,1500,1;����,1200,1;״̬,500,1;ͣ��SQL,0,1"
    Call iniVsf(strHead, vsfShow(T_�û��˺�))
    strHead = " ,300,1;ϵͳ,2000,1;����,1200,1;��ҵ��,800,1;����,2200,1;����,3000,1;�´�ִ��ʱ��,1200,1;״̬,500,1;ͣ��SQL,0,1"
    Call iniVsf(strHead, vsfShow(T_��̨��ҵ))
    strHead = " ,300,1;����,800,1;����,2000,1;��ǰֵ,1000,1;����ֵ,1000,1;Ӱ��˵��,2500,1;����˵��,1500,1;ͣ��SQL,0,1"
    Call iniVsf(strHead, vsfShow(T_����))
    
    Call LoadClient
    Call LoadUser
    Call LoadJob
    Call LoadOther
    Call ChooseAll
    lblClient.Caption = "˵�����������еĿͻ��˿��ܻᵼ�������ű�ִ�б�������ִ�л�����"
    lblUser.Caption = "˵�������δ�����û��ʺţ����ܵ���ɱ���Ŀͻ��˻Ự�������ӵ����ݿ⡣"
    lblJob.Caption = "˵������������ú�̨��ҵ��ϵͳ���ȣ����ܵ��������ű�ִ�б�������ִ�л�����"
    lblCheck.Caption = "˵�������������������ݿ�����������������ܵ���������������������ִ�л�����"
    Call frmColChoose.ClearCol
    vsfShow(T_�û��˺�).Cell(flexcpPicture, 0, vsfShow(T_�û��˺�).ColIndex("����")) = frmColChoose.imgChoose.ListImages("NoFilter").Picture
    cmdkillProcess.Visible = CheckAndAdjustMustTable("zlkillprocess")
    Call ShowFlash("")
    Call FocusRow
End Sub

Private Sub ChooseAll()
'�������ʱ����������ݣ�Ĭ��ȫ����ѡ
    Dim i As Long, j As Long
    
    For i = vsfShow.LBound To vsfShow.UBound
        For j = 0 To vsfShow.Item(i).Rows - 1
            '�����еĲ����Զ��������������ж��Ƿ�ɹ�ѡ
            If vsfShow.Item(i).Cell(flexcpChecked, j, vsfShow.Item(i).ColIndex("ѡ��")) = flexUnchecked Then
                vsfShow.Item(i).Cell(flexcpChecked, j, vsfShow.Item(i).ColIndex("ѡ��")) = flexChecked
            End If
        Next
    Next
    For i = vsfShow.LBound To vsfShow.UBound
        If vsfShow(i).Rows > 1 Then
            tbcMain.Item(i).Selected = True
            Exit For
        Else
            If i = vsfShow.UBound Then
                tbcMain.Item(0).Selected = True
            End If
        End If
    Next
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    lblTip.Move imgMain.Left + imgMain.Width + 100, 150, Me.Width - 200
    tbcMain.Move 0, imgMain.Top + imgMain.Height + 100, Me.Width, Me.Height - imgMain.Height * 3 - 100
    fraLine.Move 0, tbcMain.Top - fraLine.Height, Me.Width
    cmdCancel.Move Me.Width - cmdCancel.Width - 600, tbcMain.Height + tbcMain.Top + 200
    cmdExec.Move cmdCancel.Left - cmdExec.Width - 200, cmdCancel.Top
    lblResult.Move tbcMain.Left + 60, cmdCancel.Top
    Select Case tbcMain.Selected.Index
        Case T_�ͻ���
            vsfShow(T_�ͻ���).Move 50, 0, fraClient.Width - 150, fraClient.Height - 500
            chkShow(T_�ͻ���).Move cmdExec.Left + cmdExec.Width, vsfShow(T_�ͻ���).Height + 50
            cmdkillProcess.Move chkShow(T_�ͻ���).Left - cmdkillProcess.Width - 500, chkShow(T_�ͻ���).Top + 50
            lblClient.Move 60, vsfShow(T_�ͻ���).Height + 200
        Case T_�û��˺�
            vsfShow(T_�û��˺�).Move 50, 0, fraUser.Width - 150, fraUser.Height - 500
            lblUser.Move 60, vsfShow(T_�û��˺�).Height + 200
            chkShow(T_�û��˺�).Move cmdExec.Left, lblUser.Top - 150
        Case T_��̨��ҵ
            vsfShow(T_��̨��ҵ).Move 50, 0, fraJob.Width - 150, fraJob.Height - 500
            lblJob.Move 60, vsfShow(T_��̨��ҵ).Height + 200
            chkShow(T_��̨��ҵ).Move cmdExec.Left, lblJob.Top - 150
        Case T_����
            vsfShow(T_����).Move 50, 0, fraCheck.Width - 150, fraCheck.Height - 500
            lblCheck.Move 60, vsfShow(T_����).Height + 200
    End Select
End Sub

Private Sub tbcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    
    Call Form_Resize
    lblResult.Caption = "�������" & vsfShow(tbcMain.Selected.Index).Rows - 1 & "�����ݡ�"
    Call FocusRow
End Sub

Private Sub FocusRow()
'�������ʱ��vsf����ȫѡ���ˣ��Ƚ�������ڱ��������Խ��������
    Dim lngRow As Long
    
    If vsfShow(tbcMain.Selected.Index).Rows > 1 Then
        If vsfShow(tbcMain.Selected.Index).Row > 0 Then
            lngRow = vsfShow(tbcMain.Selected.Index).Row
        Else
            lngRow = 1
        End If
        vsfShow(tbcMain.Selected.Index).Row = lngRow
        Call vsfShow(tbcMain.Selected.Index).ShowCell(lngRow, 1)
    End If
End Sub

Private Sub iniVsf(strHead As String, vsfData As VSFlexGrid)
    
    Call InitTable(vsfData, strHead)
    
    With vsfData
        .ColKey(0) = "ѡ��"
        .Editable = flexEDKbdMouse
        .ExtendLastCol = True
        .MergeCells = flexMergeRestrictColumns
        .SelectionMode = flexSelectionByRow
        .AllowSelection = False
        .RowHeightMin = 300
        .AllowUserResizing = flexResizeColumns
        .ExplorerBar = flexExSortShow
    End With
End Sub

Private Sub LoadClient()
'���ؿͻ�������
    Dim rsTemp As ADODB.Recordset
    Dim i As Long
    
    Set rsTemp = GetData(T_�ͻ���)
    With vsfShow(T_�ͻ���)
        .Redraw = flexRDNone
        .Rows = 1
        .Rows = rsTemp.RecordCount + 1
        Do While Not rsTemp.EOF
            i = i + 1
            .TextMatrix(i, .ColIndex("�ͻ���")) = "" & rsTemp!�ͻ���
            .TextMatrix(i, .ColIndex("IP")) = "" & rsTemp!IP
            .TextMatrix(i, .ColIndex("����")) = "" & rsTemp!����
            .TextMatrix(i, .ColIndex("Ժ��")) = "" & rsTemp!Ժ��
            .TextMatrix(i, .ColIndex("��;")) = "" & rsTemp!��;
            .TextMatrix(i, .ColIndex("״̬")) = rsTemp!״̬
            .TextMatrix(i, .ColIndex("PROGRAM")) = "" & rsTemp!Program
            .TextMatrix(i, .ColIndex("SID")) = "" & rsTemp!Sid
            .TextMatrix(i, .ColIndex("SERIAL#")) = "" & rsTemp!SERIAL
            .TextMatrix(i, .ColIndex("INST_ID")) = "" & rsTemp!INST_ID
            .TextMatrix(i, .ColIndex("��ǰ��־")) = "" & rsTemp!��ǰ��־
            .TextMatrix(i, .ColIndex("ͣ��SQL")) = "" & rsTemp!ͣ��SQL
            If rsTemp!״̬ = "INACTIVE" Then
                .TextMatrix(i, .ColIndex("ѡ��")) = " "
            Else
                .Cell(flexcpChecked, i, .ColIndex("ѡ��")) = flexUnchecked
            End If
            rsTemp.MoveNext
        Loop
        .TextMatrix(0, 0) = ""
        .Cell(flexcpChecked, 0, 0) = flexUnchecked
        .Redraw = flexRDDirect
    End With
    lblResult.Caption = "�������" & vsfShow(tbcMain.Selected.Index).Rows - 1 & "�����ݡ�"
End Sub

Private Sub LoadUser()
'�����û��˺�
    Dim rsTemp As ADODB.Recordset
    Dim strDept As String
    Dim i As Long
    
    Set rsTemp = GetData(T_�û��˺�)
    With vsfShow(T_�û��˺�)
        .Redraw = flexRDNone
        .Rows = 1
        .Rows = rsTemp.RecordCount + 1
        Do While Not rsTemp.EOF
            i = i + 1
            .TextMatrix(i, .ColIndex("����")) = rsTemp!����
            .TextMatrix(i, .ColIndex("�û���")) = rsTemp!�û���
            .TextMatrix(i, .ColIndex("����")) = rsTemp!����
            .TextMatrix(i, .ColIndex("״̬")) = rsTemp!״̬
            .TextMatrix(i, .ColIndex("ͣ��SQL")) = rsTemp!ͣ��SQL
            If rsTemp!״̬ = "INACTIVE" Then
                .TextMatrix(i, .ColIndex("ѡ��")) = " "
            Else
                .Cell(flexcpChecked, i, .ColIndex("ѡ��")) = flexUnchecked
            End If
            If mblnFirst Then
                If InStr(strDept, "," & rsTemp!���� & ",") = 0 Then
                    strDept = IIf(strDept = "", ",", strDept) & rsTemp!���� & ","
                End If
            End If
            rsTemp.MoveNext
        Loop
        If mblnFirst Then
            .ColData(.ColIndex("����")) = strDept
        End If
        .TextMatrix(0, 0) = ""
        .Cell(flexcpChecked, 0, 0) = flexUnchecked
        .Redraw = flexRDDirect
    End With
    lblResult.Caption = "�������" & vsfShow(tbcMain.Selected.Index).Rows - 1 & "�����ݡ�"
    mblnFirst = False
End Sub

Private Sub LoadJob()
'���غ�̨��ҵ
    Dim rsTemp As ADODB.Recordset
    Dim i As Long
    
    Set rsTemp = GetData(T_��̨��ҵ)
    With vsfShow(T_��̨��ҵ)
        .Redraw = flexRDNone
        .Rows = 1
        .Rows = rsTemp.RecordCount + 1
        Do While Not rsTemp.EOF
            i = i + 1
            .TextMatrix(i, .ColIndex("ϵͳ")) = "" & rsTemp!ϵͳ
            .TextMatrix(i, .ColIndex("����")) = "" & rsTemp!����
            .TextMatrix(i, .ColIndex("����")) = "" & rsTemp!����
            .TextMatrix(i, .ColIndex("����")) = "" & rsTemp!����
            .TextMatrix(i, .ColIndex("��ҵ��")) = "" & rsTemp!��ҵ��
            .TextMatrix(i, .ColIndex("״̬")) = rsTemp!״̬
            .TextMatrix(i, .ColIndex("�´�ִ��ʱ��")) = "" & rsTemp!�´�ִ��ʱ��
            .TextMatrix(i, .ColIndex("ͣ��SQL")) = rsTemp!ͣ��SQL
            If rsTemp!״̬ = "INACTIVE" Then
                .TextMatrix(i, .ColIndex("ѡ��")) = " "
            Else
                .Cell(flexcpChecked, i, .ColIndex("ѡ��")) = flexUnchecked
            End If
            rsTemp.MoveNext
        Loop
        .TextMatrix(0, 0) = ""
        .Cell(flexcpChecked, 0, 0) = flexUnchecked
        .Redraw = flexRDDirect
    End With
    lblResult.Caption = "�������" & vsfShow(tbcMain.Selected.Index).Rows - 1 & "�����ݡ�"
End Sub

Private Sub LoadOther()
'������������
    Dim rsTemp As ADODB.Recordset
    Dim i As Long
    
    Set rsTemp = GetData(T_����)
    With vsfShow(T_����)
        .Redraw = flexRDNone
        .Rows = 1
        .Rows = rsTemp.RecordCount + 1
        If rsTemp.RecordCount > 0 Then rsTemp.MoveFirst
        Do While Not rsTemp.EOF
            i = i + 1
            .TextMatrix(i, .ColIndex("����")) = rsTemp!����
            .TextMatrix(i, .ColIndex("����")) = rsTemp!����
            .TextMatrix(i, .ColIndex("��ǰֵ")) = rsTemp!��ǰֵ
            .TextMatrix(i, .ColIndex("����ֵ")) = rsTemp!����ֵ
            .TextMatrix(i, .ColIndex("Ӱ��˵��")) = rsTemp!Ӱ��˵��
            .TextMatrix(i, .ColIndex("����˵��")) = rsTemp!����˵��
            .TextMatrix(i, .ColIndex("ͣ��SQL")) = "" & rsTemp!ͣ��SQL
            If rsTemp!ͣ��SQL = "" Then
                .TextMatrix(i, 0) = " "
            Else
                .Cell(flexcpChecked, i, 0) = flexUnchecked
            End If
            rsTemp.MoveNext
        Loop
        .TextMatrix(0, 0) = ""
        .Cell(flexcpChecked, 0, 0) = flexUnchecked
        .Redraw = flexRDDirect
    End With
    lblResult.Caption = "�������" & vsfShow(tbcMain.Selected.Index).Rows - 1 & "�����ݡ�"
End Sub

Private Function GetData(ByVal intIndex As Integer) As ADODB.Recordset
'��ȡ����Ҫ������
    Dim rsTemp As ADODB.Recordset
    Dim strKillProcess As String
    
    On Error GoTo ErrH
    Select Case intIndex
        Case T_�ͻ���
            strKillProcess = GetkillProcess
            gstrSQL = "Select b.����վ �ͻ���, b.Ip, Decode(b.վ��, Null, 'ȫԺ', c.����) Ժ��,b.����, a.Program, b.��;, decode(b.��ֹʹ��,0,'ACTIVE',1,'INACTIVE') ״̬, a.Sid," & vbNewLine & _
                "       a.Serial# Serial," & IIf(gblnRac, " a.INST_ID,  Decode(INST_ID, userenv('instance'), 1, 0) ��ǰ��־", "userenv('instance') INST_ID,1 ��ǰ��־") & "," & vbNewLine & _
                "'alter system kill session ' || Chr(39) || a.Sid || ',' || a.Serial# || " & IIf(mbln10g, "", IIf(gblnRac, "',@' || a.INST_ID ||", "',@' || userenv('instance') || ")) & " Chr(39) || ' immediate' ͣ��SQL" & vbNewLine & _
                "From " & IIf(gblnRac, "G", "") & "v$session a, Zlclients b, Zlnodelist c" & vbNewLine & _
                "Where a.Terminal = b.����վ(+) And Upper(a.Program) In (" & strKillProcess & ") And" & vbNewLine & _
                "b.վ��= c.���(+) And a.STATUS != 'KILLED' And a.USERNAME is Not Null And" & vbNewLine & _
                "      (a.Terminal <> Userenv('terminal') Or" & vbNewLine & _
                "      a.Terminal = Userenv('terminal') And Upper(a.Program) Not In ('VB6.EXE', 'ZLSVRSTUDIO.EXE'))" & vbNewLine & _
                IIf(chkShow(T_�ͻ���).value = 0, " And b.��ֹʹ��=0 ", "") & vbNewLine & _
                "Order By INST_ID, a.Terminal, a.Program"
        Case T_�û��˺�
            gstrSQL = "Select Null As ѡ��, e.���� ����, b.�û���, c.����, Decode(a.Account_Status, 'OPEN', 'ACTIVE', 'INACTIVE') ״̬," & vbNewLine & _
                "'alter user ' || b.�û��� || ' account lock �ָ���Update �ϻ���Ա�� Set ϵͳ�������� = 1 Where �û���='||Chr(39)||b.�û���||Chr(39) ͣ��sql" & vbNewLine & _
                "From Dba_Users a, �ϻ���Ա�� b, ��Ա�� c, ������Ա d, ���ű� e" & vbNewLine & _
                "Where a.Username = b.�û��� And b.�û��� <> '" & gstrUserName & "' " & "And b.��Աid = c.Id And c.Id = d.��Աid And d.����id = e.Id And d.ȱʡ=1 " & vbNewLine & _
                IIf(chkShow(T_�û��˺�).value = 0, " And a.Account_Status = 'OPEN' ", "") & mstrCondition & vbNewLine & _
                "Order By ����, b.�û���, c.����"
        Case T_��̨��ҵ
            'ϵͳ����
            If mbln10g Then
                gstrSQL = "Select Null ���, Null ϵͳ, 'ϵͳ����' ����, decode(a.Job_Name,'GATHER_STATS_JOB','�Զ�ͳ����Ϣ�ռ�','�Զ��ֶι���') ����, a.Owner || '.' || a.Job_Name ����, Null ��ҵ��," & vbNewLine & _
                    "       Decode(a.Enabled, 'TRUE', 'ACTIVE', 'FALSE', 'INACTIVE') ״̬,Null �´�ִ��ʱ��," & vbNewLine & _
                    "       'Call dbms_scheduler.disable(' || Chr(39) || a.Owner || '.' || a.Job_Name || Chr(39) || ')' ͣ��sql" & vbNewLine & _
                    "From Dba_Scheduler_Jobs a" & vbNewLine & _
                    "Where a.Job_Name In ('GATHER_STATS_JOB', 'AUTO_SPACE_ADVISOR_JOB')" & IIf(chkShow(T_��̨��ҵ).value = 0, " And a.Enabled = 'TRUE' ", "")
            Else
                gstrSQL = "Select Null ���, Null ϵͳ, 'ϵͳ����' ����,decode(a.Client_Name,'auto optimizer stats collection','�Զ��Ż���ͳ���ռ�','�Զ��ֶι���') ����, a.Client_Name ����, Null ��ҵ��," & vbNewLine & _
                    "       Decode(a.Status, 'ENABLED', 'ACTIVE', 'DISABLED', 'INACTIVE') ״̬,Null �´�ִ��ʱ��," & vbNewLine & _
                    "       'Call DBMS_AUTO_TASK_ADMIN.DISABLE(client_name => ' || Chr(39) || a.Client_Name || Chr(39) ||" & vbNewLine & _
                    "        ',operation => NULL,window_name => NULL)' ͣ��sql" & vbNewLine & _
                    "From Dba_Autotask_Client a" & vbNewLine & _
                    "Where a.Client_Name In ('auto optimizer stats collection', 'auto space advisor')" & IIf(chkShow(T_��̨��ҵ).value = 0, " And a.Status = 'ENABLED' ", "")
            End If
            '��̨��ҵ
            gstrSQL = gstrSQL & " Union All " & "Select c.���, c.���� ϵͳ, Decode(a.����, 1, 'ϵͳ�趨', 2, '����ת��', 3, '�û��Զ���') ����, a.����, a.����, a.��ҵ��," & vbNewLine & _
                "       Decode(b.Broken, 'N', 'ACTIVE', 'INACTIVE') ״̬,b.Next_date �´�ִ��ʱ��," & vbNewLine & _
                " 'Dbms_Job.Broken('||a.��ҵ��||',True)�ָ���Update Zlautojobs Set ϵͳ����ͣ�� = 1 Where ��ҵ��='||a.��ҵ�� ͣ��sql" & vbNewLine & _
                "From Zlautojobs a, User_Jobs b, Zlsystems c" & vbNewLine & _
                "Where b.Job = a.��ҵ�� And a.ϵͳ = c.���" & IIf(chkShow(T_��̨��ҵ).value = 0, " And b.Broken = 'N' ", "") & IIf(mstrSysNum = "", "", " And c.��� In(" & mstrSysNum & ") ")
            '�ǲ�Ʒ�Զ���ҵ
            gstrSQL = gstrSQL & " Union All " & "Select Null ���, Null ϵͳ, '�ǲ�Ʒ�Զ���ҵ' ����, Null ����, a.What ����, a.Job ��ҵ��," & vbNewLine & _
                "       Decode(a.Broken, 'N', 'ACTIVE', 'INACTIVE') ״̬,a.Next_date �´�ִ��ʱ��,'dbms_Job.Broken('||a.Job||',True)' ͣ��sql" & vbNewLine & _
                "From User_Jobs a" & vbNewLine & _
                "Where a.Job Not In (Select ��ҵ�� From Zltools.Zlautojobs) And a.Schema_User Not In (" & mstrOracleUser & ")" & vbNewLine & _
                IIf(chkShow(T_��̨��ҵ).value = 0, " And a.Broken = 'N' ", "") & vbNewLine & _
                "Order By ���, ����"
        Case T_����
            mstrUsers = GetUsers
            Set rsTemp = CopyNewRec(Nothing, True, , _
                        Array("����", adVarChar, 20, Empty, "����", adVarChar, 100, Empty, _
                              "��ǰֵ", adVarChar, 50, Empty, "����ֵ", adVarChar, 50, Empty, _
                              "Ӱ��˵��", adVarChar, 100, Empty, "����˵��", adVarChar, 100, Empty, _
                              "ͣ��SQL", adVarChar, 200, Empty))
            Call CheckSysPara(rsTemp)
            Call CheckDBFile(rsTemp)
            Call CheckTriggers(rsTemp)
            Call CheckPrivs(rsTemp)
            
            Set GetData = rsTemp
    End Select
    If intIndex <> T_���� Then Set GetData = gclsBase.OpenSQLRecord(mcnExe, gstrSQL, Decode(intIndex, T_�ͻ���, "��ȡ�ͻ���", T_�û��˺�, "��ȡ�û��˺�", T_��̨��ҵ, "��ȡ�Զ���ҵ"))
    Exit Function
ErrH:
    MsgBox err.Description, vbExclamation, Me.Caption
End Function

'******************************************************************************************************************
'���ܣ�������ݿ����
'******************************************************************************************************************
Private Sub CheckSysPara(ByRef rsData As ADODB.Recordset)
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo ErrH
    gstrSQL = "Select Name , Value From V$parameter Where Name =[1] And Value =[2]"
    Set rsTemp = gclsBase.OpenSQLRecord(mcnExe, gstrSQL, SQL_CAPTION, "optimizer_index_cost_adj", "100")
    If Not rsTemp.EOF Then rsData.AddNew Array("����", "����", "��ǰֵ", "����ֵ", "Ӱ��˵��", "����˵��", "ͣ��SQL"), _
        Array("���ݿ����", rsTemp!name, rsTemp!value, "20", "ȱʡֵ100�ᵼ�²�Ʒ��������", "", "alter system set " & rsTemp!name & "=20")
    
    Set rsTemp = gclsBase.OpenSQLRecord(mcnExe, gstrSQL, SQL_CAPTION, "optimizer_index_caching", "0")
    If Not rsTemp.EOF Then rsData.AddNew Array("����", "����", "��ǰֵ", "����ֵ", "Ӱ��˵��", "����˵��", "ͣ��SQL"), _
        Array("���ݿ����", rsTemp!name, rsTemp!value, "80", "ȱʡ0�ᵼ�²�Ʒ��������", "", "alter system set " & rsTemp!name & "=80")
        
    Set rsTemp = gclsBase.OpenSQLRecord(mcnExe, gstrSQL, SQL_CAPTION, "O7_DICTIONARY_ACCESSIBILITY", "FALSE")
    If Not rsTemp.EOF Then rsData.AddNew Array("����", "����", "��ǰֵ", "����ֵ", "Ӱ��˵��", "����˵��", "ͣ��SQL"), _
        Array("���ݿ����", rsTemp!name, rsTemp!value, "TRUE", "����ϵͳ��ͼ�޷���Ȩ��Ӱ�������Լ���Ʒ����", "���ֹ�����ΪTRUE���������ݿ�", "")
        
    gstrSQL = "Select Name , Value From V$parameter Where Name = [1] And Zl_To_Number(Value) < [2]"
    Set rsTemp = gclsBase.OpenSQLRecord(mcnExe, gstrSQL, SQL_CAPTION, "log_buffer", "209715200")
    If Not rsTemp.EOF Then rsData.AddNew Array("����", "����", "��ǰֵ", "����ֵ", "Ӱ��˵��", "����˵��", "ͣ��SQL"), _
        Array("���ݿ����", rsTemp!name, Int(Val(rsTemp!value & "") / 1024 / 1024) & "M", ">=200M", "Ӱ��ϵͳ��������������Ч��", "���ֹ��������������ݿ�", "")
    
    Set rsTemp = gclsBase.OpenSQLRecord(mcnExe, gstrSQL, SQL_CAPTION, "parallel_execution_message_size", "8192")
    If Not rsTemp.EOF Then rsData.AddNew Array("����", "����", "��ǰֵ", "����ֵ", "Ӱ��˵��", "����˵��", "ͣ��SQL"), _
        Array("���ݿ����", rsTemp!name, rsTemp!value, ">=8192", "Ӱ��ϵͳ��������ִ��", "���ֹ�����Ϊ8192���������ݿ�", "")
    Exit Sub
ErrH:
    MsgBox err.Description, vbExclamation, Me.Caption
End Sub
'******************************************************************************************************************
'���ܣ������־�ļ�
'******************************************************************************************************************
Private Sub CheckDBFile(ByRef rsData As ADODB.Recordset)
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo ErrH
    gstrSQL = "Select 'INST_ID:' || a.Inst_Id || ',GROUP:' || a.Group# Name,b.Member," & vbNewLine & _
            "       a.Bytes Value" & vbNewLine & _
            "From Gv$log A" & vbNewLine & _
            "Join Gv$logfile B" & vbNewLine & _
            "On (a.Group# = b.Group# And a.Inst_Id = b.Inst_Id)" & vbNewLine & _
            "Where a.Bytes < 104857600" & vbNewLine & _
            "Order By a.Inst_Id, a.Group#, a.Thread#, b.Member"
    Set rsTemp = gclsBase.OpenSQLRecord(mcnExe, gstrSQL, SQL_CAPTION)
    Do While Not rsTemp.EOF
        rsData.AddNew Array("����", "����", "��ǰֵ", "����ֵ", "Ӱ��˵��", "����˵��", "ͣ��SQL"), _
            Array("���ݿ��ļ�", rsTemp!name & "," & GetFileNameByPath(rsTemp!Member & ""), Int(Val(rsTemp!value & "") / 1024 / 1024) & "M", ">=100M", "Ӱ��ϵͳ��������������Ч��", "���ֹ�����Ϊ����100M", "")
        rsTemp.MoveNext
    Loop
    Exit Sub
ErrH:
    MsgBox err.Description, vbExclamation, Me.Caption
End Sub
'******************************************************************************************************************
'���ܣ���鴥����
'******************************************************************************************************************
Private Sub CheckTriggers(ByRef rsData As ADODB.Recordset)
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo ErrH
    'ZLHIS���еĴ��������ý�һ���ж϶���
    gstrSQL = "Select a.Owner, a.Trigger_Name,a.Status From Dba_Triggers A Where a.Status = 'ENABLED' And a.Table_Owner In (" & mstrUsers & ") And a.Trigger_Type <> 'INSTEAD OF'"
    Set rsTemp = gclsBase.OpenSQLRecord(mcnExe, gstrSQL, SQL_CAPTION)
    Do While Not rsTemp.EOF
        rsData.AddNew Array("����", "����", "��ǰֵ", "����ֵ", "Ӱ��˵��", "����˵��", "ͣ��SQL"), _
            Array("������", rsTemp!Owner & "." & rsTemp!trigger_name, "ENABLED", "DISABLED", "Ӱ��ñ�����������ű�ִ��Ч��", "�����ڼ����", _
            "alter trigger " & rsTemp!Owner & "." & rsTemp!trigger_name & " disable")
        rsTemp.MoveNext
    Loop
    Exit Sub
ErrH:
    MsgBox err.Description, vbExclamation, Me.Caption
End Sub
'******************************************************************************************************************
'���ܣ���������û��Ķ���Ȩ��
'******************************************************************************************************************
Private Sub CheckPrivs(ByRef rsData As ADODB.Recordset)
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo ErrH
    '������ʹ�øñ������������
    '���ZLTOOLS��PUBLICȨ��
    gstrSQL = "Select a.Grantee, a.Owner, a.Table_Name, a.Privilege" & vbNewLine & _
            "From (Select 'ZLTOOLS' Grantee, 'SYS' Owner, 'DBA_ROLE_PRIVS' Table_Name, 'SELECT' Privilege From Dual) A" & vbNewLine & _
            "Where Not Exists (Select 1" & vbNewLine & _
            "       From Dba_Tab_Privs C" & vbNewLine & _
            "       Where c.Owner = 'SYS' And (c.Grantee = 'PUBLIC' Or a.Grantee<>'PUBLIC' And c.Grantee = 'ZLTOOLS') And" & vbNewLine & _
            "             c.Table_Name = a.Table_Name And c.Privilege = a.Privilege)"
    Set rsTemp = gclsBase.OpenSQLRecord(mcnExe, gstrSQL, SQL_CAPTION)
    Do While Not rsTemp.EOF
        rsData.AddNew Array("����", "����", "��ǰֵ", "����ֵ", "Ӱ��˵��", "����˵��", "ͣ��SQL"), _
            Array("�����û�����Ȩ��", rsTemp!Grantee & " " & rsTemp!Privilege & " On " & rsTemp!Owner & "." & rsTemp!Table_Name, rsTemp!Privilege, "", "�������ܻ�����쳣���Լ�Ӱ���Ʒʹ��", "", "Grant " & rsTemp!Privilege & " On " & rsTemp!Owner & "." & rsTemp!Table_Name)
        rsTemp.MoveNext
    Loop
    Exit Sub
ErrH:
    MsgBox err.Description, vbExclamation, Me.Caption
End Sub

Private Function GetFileNameByPath(ByVal strFilePath As String) As String
    Dim lngPos As Long
    
    lngPos = InStrRev(strFilePath, "/")
    If lngPos = 0 Then
        lngPos = InStrRev(strFilePath, "\")

    End If
    If lngPos = 0 Then
        GetFileNameByPath = strFilePath
    Else
        GetFileNameByPath = Mid(strFilePath, lngPos + 1)
    End If
End Function

Private Function GetUsers() As String
'���ܣ���ȡ��ѡϵͳ��������
    Dim strTemp As String, strUser As String
    Dim rsTmp As ADODB.Recordset
    
    On Error Resume Next
    gstrSQL = ""
    strTemp = "," & mstrSysNum & ","
    If InStr(strTemp, ",0,") > 0 Then
        gstrSQL = "Select 'ZLTOOLS' ������ From Dual"
        strTemp = Replace(strTemp, ",0,", "")
        strTemp = Mid(strTemp, 1, Len(strTemp) - 1)
    Else
        strTemp = mstrSysNum
    End If
    If strTemp <> "" Then
        gstrSQL = IIf(gstrSQL = "", "", gstrSQL & " Union ") & "Select distinct ������ From Zlsystems Where ��� In (" & strTemp & ")" & vbNewLine & _
                "Union" & vbNewLine & _
                "Select ������ From Zlbakspaces Where ϵͳ In (" & strTemp & ")"
    End If
    If gstrSQL = "" Then Exit Function
    Set rsTmp = gclsBase.OpenSQLRecord(mcnExe, gstrSQL, SQL_CAPTION)
    Do While Not rsTmp.EOF
        strUser = IIf(strUser = "", "", strUser & ",") & "'" & rsTmp!������ & "'"
        rsTmp.MoveNext
    Loop
    GetUsers = strUser
End Function

Private Function GetkillProcess() As String
'��ȡ��Ʒ�еĻỰ
    Dim strKillProcess As String
    Dim rsTemp As ADODB.Recordset
    
    On Error Resume Next
    gstrSQL = "Select Count(1) ���� From Zltools.Zlkillprocess Where Rownum < 2"
    Set rsTemp = gclsBase.OpenSQLRecord(mcnExe, gstrSQL, "zlkillprocess�����ж�")
    If err.Number <> 0 Then
        strKillProcess = ""
    Else
        If rsTemp!���� <> 0 Then
            strKillProcess = "zlkillprocess"
        End If
    End If
    If strKillProcess <> "" Then
        strKillProcess = "Select Upper(����) From Zltools.Zlkillprocess Union All" & vbNewLine & _
                        "Select 'VB6.EXE' From Zltools.Zlkillprocess"
    Else
        strKillProcess = "'ZL9LABPRINTSVR.EXE','ZL9LABRECEIV.EXE','ZL9LABTCPSVR.EXE','ZL9LISCOMM.EXE'," & _
                        "'ZL9WIZARDMAIN.EXE','ZLACTMAIN.EXE','ZLHIS+.EXE','ZLHISCRUST.EXE','ZLLISRECEIVESEND.EXE'," & _
                        "'ZLNEWQUERY.EXE','ZLORCLCONFIG.EXE','ZLPACSBROWSERSTATION.EXE','ZLPACSSRV.EXE'," & _
                        "'ZLPEISAUTOANALYSE.EXE','ZLRPTSQLADJUST.EXE','ZLRUNAS.EXE','ZLSVRNOTICE.EXE'," & _
                        "'ZLSVRSTUDIO.EXE','ZLWIZARDSTART.EXE','VB6.EXE'"
    End If
    GetkillProcess = strKillProcess
End Function

Public Sub ShowMe(ByVal strSys As String, ByVal cnExe As ADODB.Connection)
    mstrSysNum = strSys
    Set mcnExe = cnExe
    Me.Show 1
End Sub

Private Sub vsfShow_AfterEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    Dim i As Long

    With vsfShow(Index)
        If Col = 0 Then
            If Row = 0 Then
                If .Cell(flexcpChecked, 0, 0) = flexChecked Then
                    .Cell(flexcpChecked, 0, 0) = flexChecked
                    For i = 1 To .Rows - 1
                        If .Cell(flexcpChecked, i, 0) = flexUnchecked Then
                            .Cell(flexcpChecked, i, 0) = flexChecked
                        End If
                    Next
                Else
                    .Cell(flexcpChecked, 0, 0) = flexUnchecked
                    For i = 1 To .Rows - 1
                        If .Cell(flexcpChecked, i, 0) = flexChecked Then
                            .Cell(flexcpChecked, i, 0) = flexUnchecked
                        End If
                    Next
                End If
            Else
                If .Cell(flexcpChecked, 0, 0) = flexChecked Then
                    .Cell(flexcpChecked, 0, 0) = flexUnchecked
                End If
                For i = 1 To .Rows - 1
                    If .Cell(flexcpChecked, i, 0) = flexUnchecked Then
                        Exit For
                    Else
                        If i = .Rows - 1 Then
                            .Cell(flexcpChecked, 0, 0) = flexChecked
                        End If
                    End If
                Next
            End If
        End If
    End With
End Sub

Private Sub vsfShow_BeforeEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsfShow(Index)
        If (Col = 0 And .TextMatrix(Row, 0) = " ") Or Col <> 0 Then Cancel = True
    End With
End Sub

Private Sub vsfShow_BeforeSort(Index As Integer, ByVal Col As Long, Order As Integer)
    Dim strFild As String
    
    If Col = 0 Then Order = 0
    If Index = T_�û��˺� And Col = vsfShow(T_�û��˺�).ColIndex("����") Then
        Order = 0
        strFild = "e.����"
        If frmColChoose.ShowMe(vsfShow(T_�û��˺�), strFild, mstrCondition) Then
            Call LoadUser
        End If
    End If
End Sub

Private Sub vsfShow_BeforeUserResize(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = 0 Then Cancel = True
End Sub
