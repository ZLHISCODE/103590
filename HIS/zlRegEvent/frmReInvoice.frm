VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmReInvoice 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Ʊ���ջ�ѡ��"
   ClientHeight    =   4035
   ClientLeft      =   2760
   ClientTop       =   3705
   ClientWidth     =   5280
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4035
   ScaleWidth      =   5280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdȡ�� 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3715
      TabIndex        =   8
      Top             =   3495
      Width           =   1400
   End
   Begin VB.Frame fraTop 
      Height          =   60
      Left            =   0
      TabIndex        =   7
      Top             =   600
      Width           =   5295
   End
   Begin VB.TextBox txtTotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000013&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   480
      IMEMode         =   3  'DISABLE
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Text            =   "0.00"
      Top             =   1200
      Width           =   1755
   End
   Begin VB.TextBox txtThis 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000013&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   480
      IMEMode         =   3  'DISABLE
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Text            =   "0.00"
      ToolTipText     =   "�����ı�ȱʡ���㷽ʽ�Ľ��ʱ�Ų���"
      Top             =   2160
      Width           =   1755
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "ȷ��(&O)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3715
      TabIndex        =   1
      Top             =   3000
      Width           =   1400
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshInvoice 
      Height          =   3090
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   3075
      _ExtentX        =   5424
      _ExtentY        =   5450
      _Version        =   393216
      Rows            =   5
      FixedCols       =   0
      BackColorBkg    =   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   2
      HighLight       =   0
      GridLinesFixed  =   1
      ScrollBars      =   2
      AllowUserResizing=   1
      FormatString    =   "^ ѡ��|^    Ʊ�ݺ�      "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label lblPrompt 
      Caption         =   "����ݱ����˷Ѻϼƺ�ʵ���յ�����Ʊ���ϼ�,ѡ���Ӧ���ջ�Ʊ�ݺ�,ȫѡ��ʾȫ��Ʊ���ջ��ش�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   4935
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�����ܽ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3360
      TabIndex        =   5
      Top             =   840
      Width           =   1200
   End
   Begin VB.Label lblMargin 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�����˷ѽ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3360
      TabIndex        =   4
      Top             =   1800
      Width           =   1440
   End
End
Attribute VB_Name = "frmReInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private mstrInvoices As String
Private mblnChange As Boolean
Private mblnOk As Boolean
Private mblnSelAll As Boolean

Public Function ShowMe(frmParent As Object, ByVal strNO As String, _
    ByVal cur���˽�� As Currency, _
    ByVal cur�����˿� As Currency, _
    ByRef strInvoices As String, _
    ByRef blnSelAll As Boolean, _
    Optional ByVal intƱ�� As Integer = 4) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����˷ѷ�Ʊ��ѡ��
    '���:frmParent-���õĸ�����
    '     strNO-�˷ѵĵ��ݺ�
    '     cur���˽��-���˵Ľ��
    '     cur�����˿�-�����˿��
    '     bln������-�Ƿ񲹽���
    '����:strInVoices-�˷�ѡ��ķ�Ʊ��
    '����:���ȷ��,����true,���򷵻�False
    '����:���˺�
    '����:2010-01-13 10:02:34
    '����:27352
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset, i As Long

    mstrInvoices = ""
    mblnChange = False
    If Mid(strNO, 1, 1) = "," Then strNO = Mid(strNO, 2)
    Set rsTmp = GetInvoice(strNO, intƱ��)
    
    If rsTmp.RecordCount = 0 Then
        strInvoices = ""
        '��δ��ӡ��Ʊ��,ֱ�ӷ���true.
        ShowMe = True: Unload Me
        Exit Function
    End If
    
    If rsTmp.RecordCount = 1 Then
        'ֻ��һ��ʱ,�ջ��ش�
        strInvoices = rsTmp!����
        ShowMe = True
        blnSelAll = True
        Unload Me: Exit Function
    End If

    With mshInvoice
        .Rows = rsTmp.RecordCount + 1
        For i = 1 To rsTmp.RecordCount
            .TextMatrix(i, 0) = "��"
            .TextMatrix(i, 1) = rsTmp!����
            rsTmp.MoveNext
        Next
    End With
    txtTotal.Text = Format(cur���˽��, "0.00")
    txtThis.Text = Format(cur�����˿�, "0.00")
    
    Me.Show 1, frmParent
    strInvoices = mstrInvoices
    blnSelAll = mblnSelAll
    ShowMe = mblnOk
End Function


Private Function GetInvoice(ByVal strNos As String, ByVal intƱ�� As Integer) As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡָ����������Ӧ�ķ�Ʊʹ�ü�
    '����:�������������ĵ��ݷ�Ʊ
    '����:���˺�
    '����:2014-10-10 17:58:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    
    On Error GoTo errH

    strSQL = _
    "   Select A.����" & vbNewLine & _
    "   From Ʊ��ʹ����ϸ A" & vbNewLine & _
    "   Where A.���� = 1 And a.ԭ�� <> 6 " & vbNewLine & _
    "           And A.��ӡid = (Select Max(ID) From Ʊ�ݴ�ӡ���� Where �������� = [2] And NO = [1])" & vbNewLine & _
    "Minus" & vbNewLine & _
    "Select A.����" & vbNewLine & _
    "From Ʊ��ʹ����ϸ A" & vbNewLine & _
    "Where A.���� = 2 And a.ԭ�� <> 6 " & vbNewLine & _
    "   And A.��ӡid = (Select Max(ID) From Ʊ�ݴ�ӡ���� Where �������� = [2] And NO = [1])" & vbNewLine & _
    "Order By ����"

    Set GetInvoice = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNos, intƱ��)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub cmdOK_Click()
    Dim i As Long
    
    With mshInvoice
        
        If .Rows > 1 Then
            mstrInvoices = ""
            For i = 1 To .Rows - 1
                If .TextMatrix(i, 0) = "��" Then
                    mstrInvoices = mstrInvoices & "," & Trim(.TextMatrix(i, 1))
                End If
            Next
            If mstrInvoices = "" Then
                MsgBox "������ѡ��һ��Ʊ��!", vbInformation, gstrSysName
                Exit Sub
            End If
            mstrInvoices = Mid(mstrInvoices, 2)
            
            If .Rows - 1 = UBound(Split(mstrInvoices, ",")) + 1 Then
                If MsgBox("��ȷ��Ҫ�ջ�����Ʊ�ݽ����ش������?", vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Sub
                End If
                mblnSelAll = True
            Else
                If MsgBox("��" & .Rows - 1 & "��Ʊ��,��ѡ�����ջ�" & UBound(Split(mstrInvoices, ",")) + 1 & "��." & vbCrLf & _
                    "��ȷ��Ҫ�ջ���ЩƱ����?" & vbCrLf & mstrInvoices, vbInformation + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
                    Exit Sub
                End If
            End If
        End If
    End With
    mblnChange = False
    mblnOk = True
    Unload Me
End Sub

Private Sub cmdȡ��_Click()
    mblnOk = False: Unload Me
End Sub

Private Sub Form_Load()
    mblnChange = False
    mblnSelAll = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnChange = True Then
        If MsgBox("����������Ʊ��ѡ��ģ�ȷʵҪ�˳���", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Cancel = 1: Exit Sub
        End If
    End If
End Sub

Private Sub mshInvoice_DblClick()
    Dim i As Long
    
    With mshInvoice
        If .Col = 0 Then
            If .Row = 0 Then
                For i = 1 To .Rows - 1
                    .TextMatrix(i, 0) = IIf(.TextMatrix(i, 0) = "", "��", "")
                Next
            Else
                 .TextMatrix(.Row, 0) = IIf(.TextMatrix(.Row, 0) = "", "��", "")
            End If
            mblnChange = True
        End If
    End With
End Sub
