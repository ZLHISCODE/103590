VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmҽ���ʻ����� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����ҽ���ʻ�"
   ClientHeight    =   3075
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6180
   Icon            =   "frmҽ���ʻ�����.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   6180
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4860
      TabIndex        =   17
      Top             =   660
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   4860
      TabIndex        =   18
      Top             =   1440
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   4860
      TabIndex        =   16
      Top             =   210
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Caption         =   "��������"
      Height          =   2835
      Left            =   180
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin MSComCtl2.DTPicker dtpBegin 
         Height          =   300
         Left            =   1320
         TabIndex        =   13
         Top             =   2340
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   65929219
         CurrentDate     =   37405
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   1
         Left            =   1320
         MaxLength       =   20
         TabIndex        =   4
         Top             =   757
         Width           =   2955
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   0
         Left            =   1320
         MaxLength       =   20
         TabIndex        =   2
         Top             =   360
         Width           =   2955
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   2
         Left            =   1320
         MaxLength       =   20
         TabIndex        =   6
         Top             =   1154
         Width           =   2955
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   3
         Left            =   1320
         MaxLength       =   20
         TabIndex        =   8
         Top             =   1551
         Width           =   2955
      End
      Begin VB.CommandButton cmd��λ 
         Caption         =   "��"
         Height          =   300
         Left            =   3990
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   1950
         Width           =   285
      End
      Begin MSComCtl2.DTPicker dtpEnd 
         Height          =   300
         Left            =   2970
         TabIndex        =   15
         Top             =   2340
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   65929219
         CurrentDate     =   37405
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   4
         Left            =   1320
         MaxLength       =   100
         TabIndex        =   10
         Top             =   1950
         Width           =   2655
      End
      Begin VB.Label lbl����ʱ�� 
         AutoSize        =   -1  'True
         Caption         =   "����ʱ��(&T)"
         Height          =   180
         Left            =   255
         TabIndex        =   12
         Top             =   2400
         Width           =   990
      End
      Begin VB.Label lbl�� 
         AutoSize        =   -1  'True
         Caption         =   "��"
         Height          =   180
         Left            =   2700
         TabIndex        =   14
         Top             =   2400
         Width           =   180
      End
      Begin VB.Label lbl˵�� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ҽ����(&Y)"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   1
         Left            =   435
         TabIndex        =   3
         Top             =   810
         Width           =   810
      End
      Begin VB.Label lbl˵�� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "������λ(&U)"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   4
         Left            =   255
         TabIndex        =   9
         Top             =   2010
         Width           =   990
      End
      Begin VB.Label lbl˵�� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "���֤(&R)"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   3
         Left            =   435
         TabIndex        =   7
         Top             =   1620
         Width           =   810
      End
      Begin VB.Label lbl˵�� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����(&N)"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   2
         Left            =   615
         TabIndex        =   5
         Top             =   1200
         Width           =   630
      End
      Begin VB.Label lbl˵�� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����(&A)"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   0
         Left            =   615
         TabIndex        =   1
         Top             =   420
         Width           =   630
      End
   End
End
Attribute VB_Name = "frmҽ���ʻ�����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum �ı�Enum
    Text���� = 0
    Textҽ���� = 1
    Text���� = 2
    Text���֤ = 3
    Text������λ = 4
End Enum

Private mstrFind As String
Private mblnOK As Boolean


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name
End Sub

Private Sub cmdOK_Click()
    Dim lngIndex As Long
    
    mstrFind = ""
    For lngIndex = txtEdit.LBound To txtEdit.UBound
        If Trim(txtEdit(lngIndex).Text) <> "" Then
            Select Case lngIndex
                Case Text����
                    mstrFind = mstrFind & " And A.���� = '" & Trim(UCase(txtEdit(lngIndex).Text)) & "'"
                Case Textҽ����
                    mstrFind = mstrFind & " And A.ҽ���� = '" & Trim(UCase(txtEdit(lngIndex).Text)) & "'"
                Case Text���֤
                    mstrFind = mstrFind & " And P.���֤�� = '" & Trim(txtEdit(lngIndex).Text) & "'"
                Case Text����
                    mstrFind = mstrFind & " And P.���� Like '" & Trim(txtEdit(lngIndex).Text) & "%'"
                Case Text������λ
                    mstrFind = mstrFind & " And P.������λ like '" & Trim(txtEdit(lngIndex).Text) & "%'"
            End Select
        End If
    Next
    mstrFind = mstrFind & " And A.����ʱ��>=to_date('" & Format(dtpBegin.Value, "yyyy-MM-dd") & "','yyyy-MM-dd') And A.����ʱ��<to_date('" & _
                                                        Format(dtpEnd.Value + 1, "yyyy-MM-dd") & "','yyyy-MM-dd')"
    
    mblnOK = True
    Unload Me
End Sub

Public Function GetFind(strFind As String) As Boolean
    dtpEnd.Value = CDate(Format(zlDataBase.Currentdate, "yyyy-MM-dd"))
    dtpBegin = DateAdd("m", -1, dtpEnd.Value)
    dtpBegin.MaxDate = dtpEnd.Value
    
    mblnOK = False
    frmҽ���ʻ�����.Show vbModal, frmҽ���ʻ�
    If mblnOK = True Then
        strFind = mstrFind
    End If
    GetFind = mblnOK
End Function

Private Sub cmd��λ_Click()
    Dim rsTemp As ADODB.Recordset
    Set rsTemp = frmPubSel.ShowSelect(Me, _
            " Select ID,�ϼ�ID,ĩ��,����,����,��ַ,�绰,��������,�ʺ�,��ϵ�� From ��Լ��λ" & _
            " Start With �ϼ�ID is NULL Connect by Prior ID=�ϼ�ID", _
            2, "������λ", , txtEdit(Text������λ).Text)
    If Not rsTemp Is Nothing Then
        txtEdit(Text������λ).Text = rsTemp("����")
        zlControl.TxtSelAll txtEdit(Text������λ)
    Else
        txtEdit(Text������λ).SetFocus
    End If
End Sub

Private Sub dtpEnd_Change()
    dtpBegin.MaxDate = dtpEnd.Value
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey (vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    zlControl.TxtSelAll txtEdit(Index)
    Select Case Index
        Case Text����, Text������λ
            zlCommFun.OpenIme True
        Case Text����, Textҽ����, Text���֤
            zlCommFun.OpenIme False
    End Select
End Sub
