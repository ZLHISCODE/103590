VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl PictureEdit 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BackStyle       =   0  '͸��
   ClientHeight    =   4875
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8145
   KeyPreview      =   -1  'True
   ScaleHeight     =   4875
   ScaleWidth      =   8145
   ToolboxBitmap   =   "PictureEdit.ctx":0000
   Begin zlTableEPR.ColorPicker CForeColor 
      Height          =   2190
      Left            =   5700
      TabIndex        =   8
      Top             =   705
      Visible         =   0   'False
      Width           =   2190
      _ExtentX        =   3863
      _ExtentY        =   3863
   End
   Begin zlTableEPR.ColorPicker CLineColor 
      Height          =   2190
      Left            =   5355
      TabIndex        =   7
      Top             =   420
      Visible         =   0   'False
      Width           =   2190
      _ExtentX        =   3863
      _ExtentY        =   3863
   End
   Begin zlTableEPR.ColorPicker CFillColor 
      Height          =   2190
      Left            =   4980
      TabIndex        =   6
      Top             =   150
      Visible         =   0   'False
      Width           =   2190
      _ExtentX        =   3863
      _ExtentY        =   3863
   End
   Begin VB.PictureBox picBG 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4605
      Left            =   255
      ScaleHeight     =   4605
      ScaleWidth      =   4425
      TabIndex        =   0
      Top             =   120
      Width           =   4425
      Begin VB.PictureBox picDraw 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2985
         Left            =   615
         ScaleHeight     =   199
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   207
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   330
         Width           =   3105
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   270
            Left            =   225
            MaxLength       =   250
            MouseIcon       =   "PictureEdit.ctx":0312
            MousePointer    =   99  'Custom
            MultiLine       =   -1  'True
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   195
            Visible         =   0   'False
            Width           =   180
         End
         Begin VB.TextBox txtTmp 
            Appearance      =   0  'Flat
            Height          =   255
            Left            =   1530
            MultiLine       =   -1  'True
            TabIndex        =   4
            TabStop         =   0   'False
            Tag             =   "������ǰ���������"
            Top             =   135
            Visible         =   0   'False
            Width           =   420
         End
         Begin VB.PictureBox picTxt 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   165
            Left            =   315
            MousePointer    =   1  'Arrow
            ScaleHeight     =   9
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   9
            TabIndex        =   3
            TabStop         =   0   'False
            ToolTipText     =   "�ƶ���˫����������"
            Top             =   135
            Visible         =   0   'False
            Width           =   165
         End
      End
      Begin VB.PictureBox picBuff 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   4260
         Left            =   4800
         Picture         =   "PictureEdit.ctx":0464
         ScaleHeight     =   4230
         ScaleWidth      =   3600
         TabIndex        =   1
         Top             =   300
         Visible         =   0   'False
         Width           =   3630
      End
   End
   Begin MSComctlLib.ImageList imgCur 
      Left            =   5895
      Top             =   3150
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PictureEdit.ctx":1A4B
            Key             =   "Pen"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PictureEdit.ctx":1BAD
            Key             =   "Move"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PictureEdit.ctx":1EC7
            Key             =   "Earse"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PictureEdit.ctx":21E1
            Key             =   "Text"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PictureEdit.ctx":2343
            Key             =   "Sel"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "PictureEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Event pOk()
Public Event pCancel()
Public mParentBar As Object                '������Ĺ�����
Public mDoc As cTableEPR                   '���ڲ�����ȫ�Ķ���,��Ҫ�õ���ǰCELL��Pictures,Marks
Public mselKey As String                   '��ǰѡ��ͼƬ����CELL��Key


Private WithEvents cbsThis As CommandBars
Attribute cbsThis.VB_VarHelpID = -1
Private mlngDrawModeID As Long              '��ǰ��ͼģʽ
Private mlngForeColor As Long               '��ǰѡ�е�������ɫ
Private mlngFillColor As Long               '��ǰѡ��������ɫ
Private mlngLineColor As Long               '��ǰѡ����������ɫ
Private mlngFillStyleID As Long             '��ǰѡ�е������ʽ
Private mlngLineWidthID As Long             '��ǰѡ�е��߿�
Private mlngLineStyleID As Long             '��ǰѡ�е�����
Private mblnInDrawing As Boolean            '�Ƿ��ڻ�ͼģʽ
Private mvarOldPoint As POINTAPI, mvarFirstPoint As POINTAPI
Private mlngSelectedCount As Long
Private mbarTool As CommandBar              '�˵�
Private mvarPolyPoints() As POINTAPI
Private mblnDblClick As Boolean             '�Ƿ�˫��
Private mlngOrgX As Long, mlngOrgY As Long  '��ʼ��������
Public Sub ToolBar_ToolExecute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    cbsThis_Execute Control
End Sub
Public Sub ToolBar_ToolUpdate(ByVal Control As XtremeCommandBars.ICommandBarControl)
    cbsThis_Update Control
End Sub
Private Sub PopupRightButton()
'�����Ҽ��˵�
Dim objPopup As CommandBar
Dim objControl As CommandBarControl
Dim cbpPopup As CommandBarPopup     '��ʱ����
Dim objCustControl As CommandBarControlCustom       '�Զ���ؼ�

    Set objPopup = cbsThis.Add("Popup", xtpBarPopup)
    With objPopup.Controls
        Set objControl = .Add(xtpControlButton, ID_DRAW_SELECT, "ѡ��(&E)")
        Set objControl = .Add(xtpControlButton, ID_DRAW_MOVE, "�ƶ�(&M)"): objControl.Style = xtpButtonIconAndCaption
        Set cbpPopup = .Add(xtpControlButtonPopup, 0, "���"): objControl.Style = xtpButtonIconAndCaption
            Set objControl = cbpPopup.CommandBar.Controls.Add(xtpControlButton, ID_DRAW_LINE, "ֱ��(&L)")
            Set objControl = cbpPopup.CommandBar.Controls.Add(xtpControlButton, ID_DRAW_MLINE, "����(&Z)")
            Set objControl = cbpPopup.CommandBar.Controls.Add(xtpControlButton, ID_DRAW_RECT, "����(&R)")
            Set objControl = cbpPopup.CommandBar.Controls.Add(xtpControlButton, ID_DRAW_MRECT, "�����(&W)")
            Set objControl = cbpPopup.CommandBar.Controls.Add(xtpControlButton, ID_DRAW_CIRCLE, "��Բ(&C)")
        Set objControl = .Add(xtpControlButton, ID_DRAW_TEXT, "����(&T)")
        Set objControl = .Add(xtpControlButton, ID_DRAW_SEQUENCENUMBER, "˳����(&N)")
        Set objControl = .Add(xtpControlButton, ID_DRAW_CLEARNUMBERS, "���˳����(&K)")
        Set objControl = .Add(xtpControlButton, ID_DRAW_DELETE, "ɾ�����(&D)"): objControl.IconId = 325
        Set objControl = .Add(xtpControlButton, ID_DRAW_RESET, "��ձ��(&R)"):             objControl.BeginGroup = True
        
        Set cbpPopup = .Add(xtpControlButtonPopup, ID_DRAW_FILLSTYLE, "�����ʽ"):          cbpPopup.Style = xtpButtonIconAndCaption
            cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_DRAW_FILLNONE, "�����"
            cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_DRAW_FILLALL, "ʵ�����"
            cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_DRAW_FILLH, "�������"
            cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_DRAW_FILLV, "�������"
            cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_DRAW_FILLHV, "�������"
            cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_DRAW_FILLR, "��б�����"
            cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_DRAW_FILLL, "��б�����"
            cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_DRAW_FILLLR, "���������"
        
        Set cbpPopup = .Add(xtpControlButtonPopup, ID_DRAW_LINESTYLE, "����"):          cbpPopup.Style = xtpButtonIconAndCaption
            cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_DRAW_LINECONTINUE, "ʵ��"
            cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_DRAW_LINEDOT, "����"
            cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_DRAW_LINEDASH, "����"
            cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_DRAW_LINEDASHDOT, "�㻮��"
            cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_DRAW_LINEDASHDOT2, "��㻮��"
        
'        Set cbpPopup = .Add(xtpControlButtonPopup, ID_DRAW_LINEWIDTH, "�߿�"):          cbpPopup.Style = xtpButtonIconAndCaption
            cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_DRAW_LINEWIDTH1, "1�����"
            cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_DRAW_LINEWIDTH2, "2�����"
            cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_DRAW_LINEWIDTH3, "3�����"
            cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_DRAW_LINEWIDTH4, "4�����"
            cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_DRAW_LINEWIDTH5, "5�����"
        
        Set cbpPopup = .Add(xtpControlSplitButtonPopup, ID_DRAW_FILLCOLOR, "�����ɫ")
        Set objCustControl = cbpPopup.CommandBar.Controls.Add(xtpControlCustom, 0, ""): objCustControl.Handle = CFillColor.hWnd
        
        Set cbpPopup = .Add(xtpControlSplitButtonPopup, ID_DRAW_LINECOLOR, "������ɫ")
        Set objCustControl = cbpPopup.CommandBar.Controls.Add(xtpControlCustom, 0, ""): objCustControl.Handle = CLineColor.hWnd
'
'        Set cbpPopup = .Add(xtpControlSplitButtonPopup, ID_DRAW_FONTCOLOR, "������ɫ")             '��ʱδ�õ�
'        Set objCustControl = cbpPopup.CommandBar.Controls.Add(xtpControlCustom, 0, ""): objCustControl.Handle = CForeColor.hWnd
        Set objControl = .Add(xtpControlButton, ID_EDIT_DELETE, "���ͼƬ"): objControl.BeginGroup = True
        If UserControl.Extender.Tag = "�ο�ͼ" Then Set objControl = .Add(xtpControlButton, ID_INSERT_PICTURE, "��ͼ(&C)")
    End With
    objPopup.ShowPopup: objPopup.SetIconSize 32, 32
End Sub

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim i As Long
    Select Case Control.ID
    Case ID_DRAW_CLEARNUMBERS '��������Զ����
        Call DeleteSelectedMarks(True)
    Case ID_DRAW_SELECT, ID_DRAW_MOVE, ID_DRAW_LINE, ID_DRAW_MLINE, ID_DRAW_RECT, ID_DRAW_MRECT, ID_DRAW_CIRCLE, ID_DRAW_TEXT, ID_DRAW_SEQUENCENUMBER
        mlngDrawModeID = Control.ID 'ȷ���滭ģʽ
        If mblnInDrawing = False Then GetCurDrawMode
    Case ID_DRAW_DELETE     '���ѡ�б��
        If mblnInDrawing = False Then DeleteSelectedMarks
    Case ID_DRAW_RESET      '������б��
        If mblnInDrawing = False Then DeleteSelectedMarks False, True
    Case ID_DRAW_FILLNONE, ID_DRAW_FILLALL, ID_DRAW_FILLH, ID_DRAW_FILLV, ID_DRAW_FILLHV, ID_DRAW_FILLR, ID_DRAW_FILLL, ID_DRAW_FILLLR
        mlngFillStyleID = Control.ID 'ȷ����䷽ʽ���ػ�
        Call GetCurDrawMode
        Call ChangeLineAndReDraw(1)
    Case ID_DRAW_LINECONTINUE, ID_DRAW_LINEDOT, ID_DRAW_LINEDASH, ID_DRAW_LINEDASHDOT, ID_DRAW_LINEDASHDOT2
        mlngLineStyleID = Control.ID 'ȷ�����Ͳ��ػ�
        GetCurDrawMode
        Call ChangeLineAndReDraw(2)
    Case ID_DRAW_LINEWIDTH1, ID_DRAW_LINEWIDTH2, ID_DRAW_LINEWIDTH3, ID_DRAW_LINEWIDTH4, ID_DRAW_LINEWIDTH5
        Control.Checked = True: mlngLineWidthID = Control.ID 'ȷ���߿��ػ�
        GetCurDrawMode
        Call ChangeLineAndReDraw(3)
    Case ID_DRAW_FILLCOLOR
        Call CFillColor_pOK(False)
    Case ID_DRAW_LINECOLOR
        Call CLineColor_pOK(False)
    Case ID_DRAW_FONTCOLOR
        Call CForeColor_pOK(False)
    Case ID_EDIT_DELETE
        UserControl.Extender.Visible = False
    End Select
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If UserControl.Extender.Visible = False Then Exit Sub
    Select Case Control.ID
        Case ID_DRAW_DELETE
            Control.Enabled = (mlngSelectedCount > 0)
        Case ID_DRAW_SELECT, ID_DRAW_MOVE, ID_DRAW_LINE, ID_DRAW_MLINE, ID_DRAW_RECT, ID_DRAW_MRECT, ID_DRAW_CIRCLE, ID_DRAW_TEXT, ID_DRAW_SEQUENCENUMBER
            If mlngDrawModeID = Control.ID Then Control.Checked = True Else Control.Checked = False
        Case ID_DRAW_FILLNONE, ID_DRAW_FILLALL, ID_DRAW_FILLH, ID_DRAW_FILLV, ID_DRAW_FILLHV, ID_DRAW_FILLR, ID_DRAW_FILLL, ID_DRAW_FILLLR
            If mlngFillStyleID = Control.ID Then Control.Checked = True Else Control.Checked = False   'ȷ����䷽ʽ���ػ�
        Case ID_DRAW_LINECONTINUE, ID_DRAW_LINEDOT, ID_DRAW_LINEDASH, ID_DRAW_LINEDASHDOT, ID_DRAW_LINEDASHDOT2
            If mlngLineStyleID = Control.ID Then Control.Checked = True Else Control.Checked = False 'ȷ�����Ͳ��ػ�
        Case ID_DRAW_LINEWIDTH1, ID_DRAW_LINEWIDTH2, ID_DRAW_LINEWIDTH3, ID_DRAW_LINEWIDTH4, ID_DRAW_LINEWIDTH5
            If mlngLineWidthID = Control.ID Then Control.Checked = True Else Control.Checked = False 'ȷ���߿��ػ�
        Case ID_INSERT_PICTURE
            If UserControl.Extender.Tag <> "�ο�ͼ" Then Control.Visible = False Else Control.Visible = True
    End Select
End Sub



Private Sub picDraw_DblClick()
    mblnDblClick = True
End Sub
Private Sub picDraw_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'��ȡ��ǰ��ͼģʽ��Ϣ
Dim lngKey As Long, i As Long, X1 As Long, Y1 As Long, ary As Variant
    
    mblnDblClick = False
    If Shift = 7 Then DesignDraw True
    If Button = vbRightButton Then Exit Sub
    If Not mblnInDrawing Then Call GetCurDrawMode

    If txt.Visible Then FinishInputText         '�����ı�������

    '��ʼ�����
    Select Case mlngDrawModeID
    Case ID_DRAW_SELECT
        '������ʼ��λ��
        mvarFirstPoint.x = x
        mvarFirstPoint.y = y
        mvarOldPoint.x = x
        mvarOldPoint.y = y
        mblnInDrawing = True
    Case ID_DRAW_MOVE
        '������ʼ��λ��
        mvarFirstPoint.x = x
        mvarFirstPoint.y = y
        mvarOldPoint.x = x
        mvarOldPoint.y = y
        mblnInDrawing = True
    Case ID_DRAW_LINE
        '������ʼ��λ��
        mvarFirstPoint.x = x
        mvarFirstPoint.y = y
        mvarOldPoint.x = x
        mvarOldPoint.y = y

        picDraw.DrawMode = vbInvert
        MoveToEx picDraw.hdc, mvarFirstPoint.x, mvarFirstPoint.y, 0
        LineTo picDraw.hdc, mvarOldPoint.x, mvarOldPoint.y

        mblnInDrawing = True
    Case ID_DRAW_RECT
        '������ʼ��λ��
        mvarFirstPoint.x = x
        mvarFirstPoint.y = y
        mvarOldPoint.x = x
        mvarOldPoint.y = y
        mblnInDrawing = True
    Case ID_DRAW_MLINE
        If mblnInDrawing = False Then
            '������ʼ��λ��
            mvarFirstPoint.x = x
            mvarFirstPoint.y = y
            mvarOldPoint.x = x
            mvarOldPoint.y = y
            ReDim mvarPolyPoints(1 To 1) As POINTAPI
            mvarPolyPoints(1).x = x: mvarPolyPoints(1).y = y
        End If
        mblnInDrawing = True
    Case ID_DRAW_MRECT
        If mblnInDrawing = False Then
            '������ʼ��λ��
            mvarFirstPoint.x = x
            mvarFirstPoint.y = y
            mvarOldPoint.x = x
            mvarOldPoint.y = y
            ReDim mvarPolyPoints(1 To 1) As POINTAPI
            mvarPolyPoints(1).x = x: mvarPolyPoints(1).y = y
        End If
        mblnInDrawing = True
    Case ID_DRAW_CIRCLE
        '������ʼ��λ��
        mvarFirstPoint.x = x
        mvarFirstPoint.y = y
        mvarOldPoint.x = x
        mvarOldPoint.y = y
        mblnInDrawing = True
    Case ID_DRAW_DELETE
        mblnInDrawing = True
    Case ID_DRAW_TEXT
        '���Ƿ�ѡ����ĳһ���ı�
        If mDoc.Cells(mselKey).PicMarkKey <> "" Then
            ary = Split(mDoc.Cells(mselKey).PicMarkKey, "|")
            mDoc.Cells(mselKey).PicMarkKey = ""
            For i = 1 To UBound(ary)
                If mDoc.PicMarks("K" & ary(i)).���� = 0 Then
                    If x > mDoc.PicMarks("K" & ary(i)).X1 And x < mDoc.PicMarks("K" & ary(i)).X2 And y > mDoc.PicMarks("K" & ary(i)).Y1 - 2 And y < mDoc.PicMarks("K" & ary(i)).Y2 - 2 Then
                        lngKey = ary(i)
                    Else
                        mDoc.Cells(mselKey).PicMarkKey = mDoc.Cells(mselKey).PicMarkKey & "|" & ary(i)
                    End If
                Else
                    mDoc.Cells(mselKey).PicMarkKey = mDoc.Cells(mselKey).PicMarkKey & "|" & ary(i)
                End If
            Next i
        End If

        If lngKey > 0 Then 'ѡ��һ�������ı�
            With mDoc.PicMarks("K" & lngKey)
                txt.Font.Name = .����
                txt.Text = .����
                txt.Move .X1, .Y1, (.X2 - .X1), (.Y2 - .Y1)
            End With
            mDoc.PicMarks.Remove "K" & lngKey '��ɾ����ǰ�ı�,Ȼ���ػ����ж���
            Call ReDrawPicMarks '���������
        Else '�½�һ���ı�
            txt.Text = ""
            txt.Top = y: txt.Left = x
            Call GetFitTxtSize(txt, txt.Text, X1, Y1)
            txt.Width = X1 + 10
            txt.Height = Y1 + 6
        End If
        picTxt.Top = txt.Top - picTxt.Height / 2
        picTxt.Left = txt.Left + txt.Width - picTxt.Width / 2
        txt.Visible = True:         picTxt.Visible = True
        txt.SetFocus
    Case ID_DRAW_SEQUENCENUMBER
        If mlngFillColor = 0 Then
            Call SetDrawStyleFromValue(picDraw.hdc, RGB(255, 255, 0), 0, 1, RGB(255, 255, 0), -2)
        Else
            Call SetDrawStyleFromValue(picDraw.hdc, RGB(255, 255, 0), 0, 1, mlngFillColor, -2)
        End If
        Ellipse picDraw.hdc, x - 7, y - 7, x + 7, y + 7
        If mlngLineColor = 0 Then
            Call SetDrawStyleFromValue(picDraw.hdc, vbBlack, 0, 1, vbBlack, -1)
        Else
            Call SetDrawStyleFromValue(picDraw.hdc, mlngLineColor, 0, 1, mlngLineColor, -1)
        End If
        Ellipse picDraw.hdc, x - 7, y - 7, x + 7, y + 7
        Dim Font As New StdFont
        Font.Bold = True
        Dim Num As Long
        Num = GetMaxNum
        Call TextOut(picDraw, Num, IIf(Len(CStr(Num)) > 1, x - 6, x - 2), y - 6, x + 14, y + 14, Font)

        picDraw.Refresh
        '��������
        lngKey = mDoc.PicMarks.Add
        With mDoc.PicMarks("K" & lngKey)
            .���� = Num
            .X1 = x: .Y1 = y
            .X2 = x: .Y2 = y
            .���� = 6
            .��䷽ʽ = -2
            .���ɫ = mlngFillColor
            .�߿� = 1
            .����ɫ = mlngLineColor
            .���� = 1
        End With
        mDoc.Cells(mselKey).PicMarkKey = mDoc.Cells(mselKey).PicMarkKey & "|" & lngKey
        mblnInDrawing = False
    End Select
End Sub

Private Sub picDraw_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    If mblnInDrawing = False Then Exit Sub
    Dim tmpX As Long, tmpY As Long

    '���߻��Ʊ߿�
    Call SetDrawStyleFromValue(picDraw.hdc, mlngLineColor, IIf(gcurPenStyle = 0, 2, gcurPenStyle), gcurPenWidth, mlngFillColor, -1)

    Select Case mlngDrawModeID
    Case ID_DRAW_SELECT
        '���߻��Ʊ߿�
        Call SetDrawStyleFromValue(picDraw.hdc, mlngLineColor, IIf(gcurPenStyle = 0, 2, gcurPenStyle), 1, mlngFillColor, -1)
        '����
        picDraw.DrawMode = vbInvert
        Rectangle picDraw.hdc, mvarFirstPoint.x, mvarFirstPoint.y, mvarOldPoint.x, mvarOldPoint.y
        '����
        Rectangle picDraw.hdc, mvarFirstPoint.x, mvarFirstPoint.y, x, y
        picDraw.Refresh
        '�����µ�ĩβ��λ��
        mvarOldPoint.x = x
        mvarOldPoint.y = y
    Case ID_DRAW_MOVE
        '�ƶ�ѡ�б��
        '����
        tmpX = mvarOldPoint.x - mvarFirstPoint.x: tmpY = mvarOldPoint.y - mvarFirstPoint.y  '��ƫ����
        RefreshSelectedMarks picDraw, tmpX, tmpY    'ˢ��ѡ�еı�ǵ��µ�ַ

        '����
        tmpX = x - mvarFirstPoint.x: tmpY = y - mvarFirstPoint.y
        RefreshSelectedMarks picDraw, tmpX, tmpY    'ˢ��ѡ�еı�ǵ��µ�ַ
        picDraw.Refresh
        mvarOldPoint.x = x
        mvarOldPoint.y = y
    Case ID_DRAW_LINE
        '������ǰ����
        picDraw.DrawMode = vbInvert
        MoveToEx picDraw.hdc, mvarFirstPoint.x, mvarFirstPoint.y, 0
        LineTo picDraw.hdc, mvarOldPoint.x, mvarOldPoint.y

        '�����µ�����
        MoveToEx picDraw.hdc, mvarFirstPoint.x, mvarFirstPoint.y, 0
        LineTo picDraw.hdc, x, y
        picDraw.Refresh
        '�����µ�ĩβ��λ��
        mvarOldPoint.x = x
        mvarOldPoint.y = y
    Case ID_DRAW_RECT
        tmpX = x: tmpY = y
        If Shift = 2 Then '������
            Call ForceSquare(mvarFirstPoint.x, mvarFirstPoint.y, tmpX, tmpY)
        End If
        '����
        picDraw.DrawMode = vbInvert
        Rectangle picDraw.hdc, mvarFirstPoint.x, mvarFirstPoint.y, mvarOldPoint.x, mvarOldPoint.y
        '����
        Rectangle picDraw.hdc, mvarFirstPoint.x, mvarFirstPoint.y, tmpX, tmpY
        picDraw.Refresh
        '�����µ�ĩβ��λ��
        mvarOldPoint.x = tmpX
        mvarOldPoint.y = tmpY
    Case ID_DRAW_MLINE
        picDraw.DrawMode = vbInvert
        MoveToEx picDraw.hdc, mvarFirstPoint.x, mvarFirstPoint.y, 0
        LineTo picDraw.hdc, mvarOldPoint.x, mvarOldPoint.y

        '�����µ�����
        picDraw.DrawMode = vbInvert
        MoveToEx picDraw.hdc, mvarFirstPoint.x, mvarFirstPoint.y, 0
        LineTo picDraw.hdc, x, y
        picDraw.Refresh
        '�����µ�ĩβ��λ��
        mvarOldPoint.x = x
        mvarOldPoint.y = y
    Case ID_DRAW_MRECT
        picDraw.DrawMode = vbInvert
        MoveToEx picDraw.hdc, mvarFirstPoint.x, mvarFirstPoint.y, 0
        LineTo picDraw.hdc, mvarOldPoint.x, mvarOldPoint.y

        '�����µ�����
        picDraw.DrawMode = vbInvert
        MoveToEx picDraw.hdc, mvarFirstPoint.x, mvarFirstPoint.y, 0
        LineTo picDraw.hdc, x, y
        picDraw.Refresh
        '�����µ�ĩβ��λ��
        mvarOldPoint.x = x
        mvarOldPoint.y = y
    Case ID_DRAW_CIRCLE
        tmpX = x: tmpY = y
        If Shift = 2 Then '������
            Call ForceSquare(mvarFirstPoint.x, mvarFirstPoint.y, tmpX, tmpY)
        End If
        '����
        picDraw.DrawMode = vbInvert
        Ellipse picDraw.hdc, mvarFirstPoint.x, mvarFirstPoint.y, mvarOldPoint.x, mvarOldPoint.y
        '����
        Ellipse picDraw.hdc, mvarFirstPoint.x, mvarFirstPoint.y, tmpX, tmpY
        picDraw.Refresh
        '�����µ�ĩβ��λ��
        mvarOldPoint.x = tmpX
        mvarOldPoint.y = tmpY
    End Select
End Sub

Private Sub picDraw_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton And mblnInDrawing = False Then: Call PopupRightButton: Exit Sub '�Ҽ��˵�
    If mblnInDrawing = False Then Exit Sub
    
    '�ָ���䷽ʽ
    Call SetDrawStyleFromValue(picDraw.hdc, mlngLineColor, gcurPenStyle, gcurPenWidth, mlngFillColor, gcurFillStyle)
    Dim tmpX As Long, tmpY As Long
    Dim strTmp As String, i As Long, lngKey As Long

    Select Case mlngDrawModeID
    Case ID_DRAW_SELECT
        '����
        '���߻��Ʊ߿�
        Call SetDrawStyleFromValue(picDraw.hdc, mlngLineColor, IIf(gcurPenStyle = 0, 2, gcurPenStyle), 1, mlngFillColor, -1)
        picDraw.DrawMode = vbInvert
        Rectangle picDraw.hdc, mvarFirstPoint.x, mvarFirstPoint.y, mvarOldPoint.x, mvarOldPoint.y
        picDraw.Refresh
        mblnInDrawing = False

        'ѡ�з�ΧΪ��mvarFirstPoint,mvarOldPoint����
        '�����ж����б������Щ��ѡ�У���������ʾ
        Call HilightSelectMarks(mvarFirstPoint.x, mvarFirstPoint.y, mvarOldPoint.x, mvarOldPoint.y)
    Case ID_DRAW_MOVE
        '�����±�ǣ�ˢ��ͼ��
        tmpX = x - mvarFirstPoint.x: tmpY = y - mvarFirstPoint.y
        SaveSelectedMarks tmpX, tmpY
        Call ReDrawPicMarks
        mblnInDrawing = False
    Case ID_DRAW_LINE
        '������������
        picDraw.DrawMode = vbCopyPen
        MoveToEx picDraw.hdc, mvarFirstPoint.x, mvarFirstPoint.y, 0
        LineTo picDraw.hdc, x, y
        '��������
        lngKey = mDoc.PicMarks.Add
        With mDoc.PicMarks("K" & lngKey)
            .X1 = mvarFirstPoint.x: .Y1 = mvarFirstPoint.y
            .X2 = x: .Y2 = y
            .���� = 1            '0-�ı�,1-����,2,����,3-����,4-�����,5-Բ(��Բ)
            .��䷽ʽ = gcurFillStyle
            .���ɫ = gcurFillColor
            .�߿� = gcurPenWidth
            .����ɫ = gcurPenColor
            .���� = gcurPenStyle
        End With
        mDoc.Cells(mselKey).PicMarkKey = mDoc.Cells(mselKey).PicMarkKey & "|" & lngKey
        mblnInDrawing = False
    Case ID_DRAW_RECT
        tmpX = x: tmpY = y
        If Shift = 2 Then '������
            Call ForceSquare(mvarFirstPoint.x, mvarFirstPoint.y, tmpX, tmpY)
        End If
        '����
        picDraw.DrawMode = vbCopyPen
        Rectangle picDraw.hdc, mvarFirstPoint.x, mvarFirstPoint.y, tmpX, tmpY
        picDraw.Refresh
        '��������
        lngKey = mDoc.PicMarks.Add
        With mDoc.PicMarks("K" & lngKey)
            .X1 = mvarFirstPoint.x: .Y1 = mvarFirstPoint.y
            .X2 = tmpX: .Y2 = tmpY
            .���� = 3            '0-�ı�,1-����,2,����,3-����,4-�����,5-Բ(��Բ)
            .��䷽ʽ = gcurFillStyle
            .���ɫ = gcurFillColor
            .�߿� = gcurPenWidth
            .����ɫ = gcurPenColor
            .���� = gcurPenStyle
        End With
        mDoc.Cells(mselKey).PicMarkKey = mDoc.Cells(mselKey).PicMarkKey & "|" & lngKey
        mblnInDrawing = False
    Case ID_DRAW_MLINE
        If Button = vbRightButton Then '�Ҽ�ȡ����ǰ��ͼ
            Call ReDrawPicMarks
            mblnInDrawing = False
            ReDim mvarPolyPoints(0)
            Exit Sub
        End If

        '���浱ǰ��
        ReDim Preserve mvarPolyPoints(1 To UBound(mvarPolyPoints) + 1) As POINTAPI
        mvarPolyPoints(UBound(mvarPolyPoints)).x = x
        mvarPolyPoints(UBound(mvarPolyPoints)).y = y

        If mblnDblClick And UBound(mvarPolyPoints) >= 2 Then
            '�������ݣ��˳���ͼ
            lngKey = mDoc.PicMarks.Add
            With mDoc.PicMarks("K" & lngKey)
                .���� = 2            '0-�ı�,1-����,2,����,3-����,4-�����,5-Բ(��Բ)
                .��䷽ʽ = gcurFillStyle
                .���ɫ = gcurFillColor
                .�߿� = gcurPenWidth
                .����ɫ = gcurPenColor
                .���� = gcurPenStyle
                For i = 1 To UBound(mvarPolyPoints)
                    If i = 1 Then
                        strTmp = strTmp & CStr(mvarPolyPoints(i).x) & "," & CStr(mvarPolyPoints(i).y)
                    Else
                        strTmp = strTmp & ";" & CStr(mvarPolyPoints(i).x) & "," & CStr(mvarPolyPoints(i).y)
                    End If
                Next i
                .�㼯 = strTmp              '����㼯����
            End With
            mDoc.Cells(mselKey).PicMarkKey = mDoc.Cells(mselKey).PicMarkKey & "|" & lngKey
            mblnInDrawing = False
        End If

        '������������
        picDraw.DrawMode = vbCopyPen
        MoveToEx picDraw.hdc, mvarFirstPoint.x, mvarFirstPoint.y, 0
        LineTo picDraw.hdc, x, y

        mvarFirstPoint.x = x
        mvarFirstPoint.y = y
        mvarOldPoint.x = x
        mvarOldPoint.y = y
    Case ID_DRAW_MRECT
        If Button = vbRightButton Then
            Call ReDrawPicMarks
            mblnInDrawing = False
            ReDim mvarPolyPoints(0)
            Exit Sub
        End If

        '���浱ǰ��
        ReDim Preserve mvarPolyPoints(1 To UBound(mvarPolyPoints) + 1) As POINTAPI
        mvarPolyPoints(UBound(mvarPolyPoints)).x = x
        mvarPolyPoints(UBound(mvarPolyPoints)).y = y

        If mblnDblClick And UBound(mvarPolyPoints) >= 2 Then
            '�������ն����
            picDraw.DrawMode = vbCopyPen
            Polygon picDraw.hdc, mvarPolyPoints(1), UBound(mvarPolyPoints)

            '�������ݣ��˳���ͼ
            lngKey = mDoc.PicMarks.Add
            With mDoc.PicMarks("K" & lngKey)
                .���� = 4            '0-�ı�,1-����,2,����,3-����,4-�����,5-Բ(��Բ)
                .��䷽ʽ = gcurFillStyle
                .���ɫ = gcurFillColor
                .�߿� = gcurPenWidth
                .����ɫ = gcurPenColor
                .���� = gcurPenStyle
                For i = 1 To UBound(mvarPolyPoints)
                    If i = 1 Then
                        strTmp = strTmp & CStr(mvarPolyPoints(i).x) & "," & CStr(mvarPolyPoints(i).y)
                    Else
                        strTmp = strTmp & ";" & CStr(mvarPolyPoints(i).x) & "," & CStr(mvarPolyPoints(i).y)
                    End If
                Next i
                .�㼯 = strTmp              '����㼯����
            End With
            mDoc.Cells(mselKey).PicMarkKey = mDoc.Cells(mselKey).PicMarkKey & "|" & lngKey
            mblnInDrawing = False
        End If

        '������������
        picDraw.DrawMode = vbCopyPen
        MoveToEx picDraw.hdc, mvarFirstPoint.x, mvarFirstPoint.y, 0
        LineTo picDraw.hdc, x, y

        mvarFirstPoint.x = x
        mvarFirstPoint.y = y
        mvarOldPoint.x = x
        mvarOldPoint.y = y
    Case ID_DRAW_CIRCLE
        tmpX = x: tmpY = y
        If Shift = 2 Then '������
            Call ForceSquare(mvarFirstPoint.x, mvarFirstPoint.y, tmpX, tmpY)
        End If
        '����
        picDraw.DrawMode = vbCopyPen
        Ellipse picDraw.hdc, mvarFirstPoint.x, mvarFirstPoint.y, tmpX, tmpY
        picDraw.Refresh
        '��������
        lngKey = mDoc.PicMarks.Add
        With mDoc.PicMarks("K" & lngKey)
            .X1 = mvarFirstPoint.x: .Y1 = mvarFirstPoint.y
            .X2 = tmpX: .Y2 = tmpY
            .���� = 5            '0-�ı�,1-����,2,����,3-����,4-�����,5-Բ(��Բ)
            .��䷽ʽ = gcurFillStyle
            .���ɫ = gcurFillColor
            .�߿� = gcurPenWidth
            .����ɫ = gcurPenColor
            .���� = gcurPenStyle
        End With
        mblnInDrawing = False
        mDoc.Cells(mselKey).PicMarkKey = mDoc.Cells(mselKey).PicMarkKey & "|" & lngKey
    End Select

    picDraw.DrawMode = vbCopyPen
    picDraw.Refresh
End Sub

Private Sub picTxt_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    mlngOrgX = x: mlngOrgY = y
End Sub

Private Sub picTxt_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        If txt.Left + x - mlngOrgX >= 0 And txt.Left + x - mlngOrgX + txt.Width <= picDraw.ScaleWidth Then
            picTxt.Left = picTxt.Left + x - mlngOrgX
            txt.Left = txt.Left + x - mlngOrgX
        End If
        If txt.Top + y - mlngOrgY >= 0 And txt.Top + y - mlngOrgY + txt.Height <= picDraw.ScaleHeight Then
            picTxt.Top = picTxt.Top + y - mlngOrgY
            txt.Top = txt.Top + y - mlngOrgY
        End If
        picDraw.Refresh
    End If
End Sub

Private Sub picTxt_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    txt.SetFocus
End Sub

Private Sub UserControl_Hide()
    If Not mbarTool Is Nothing Then mbarTool.Delete
    Set cbsThis = Nothing
    Set mbarTool = Nothing
    Set mParentBar = Nothing
    Set mDoc = Nothing
    mlngDrawModeID = 0         '��ǰ��ͼģʽ
    mlngForeColor = 0          '��ǰѡ�е�������ɫ
    mlngFillColor = 0          '��ǰѡ��������ɫ
    mlngLineColor = 0         '��ǰѡ����������ɫ
    mlngFillStyleID = 0      '��ǰѡ�е������ʽ
    mlngLineWidthID = 0      '��ǰѡ�е��߿�
    mlngLineStyleID = 0    '��ǰѡ�е�����
    mblnInDrawing = False         '�Ƿ��ڻ�ͼģʽ
    mvarOldPoint.x = 0: mvarOldPoint.y = 0
    mvarFirstPoint.x = 0: mvarFirstPoint.y = 0
    mlngSelectedCount = 0

    ReDim mvarPolyPoints(0)
    mblnDblClick = False           '�Ƿ�˫��
    mlngOrgX = 0: mlngOrgY = 0 '��ʼ��������
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyDelete Then
        Call DeleteSelectedMarks
    End If
    
End Sub

Private Sub UserControl_Resize()
    Call DesignDraw
End Sub

Private Sub UserControl_Show()
    Set cbsThis = mParentBar '���ڹ�����
    If cbsThis Is Nothing Then Exit Sub
    If mDoc.ET = TabET_�������༭ Then
        Dim objControl As CommandBarControl
        Dim cbpPopup As CommandBarPopup     '��ʱ����
        Dim objCustControl As CommandBarControlCustom       '�Զ���ؼ�

        Set mbarTool = cbsThis.Add("Popup", xtpBarBottom)
        With mbarTool.Controls
            Set objControl = .Add(xtpControlButton, ID_DRAW_SELECT, "ѡ��(&E)")
            Set objControl = .Add(xtpControlButton, ID_DRAW_MOVE, "�ƶ�(&M)"): objControl.Style = xtpButtonIconAndCaption
            Set cbpPopup = .Add(xtpControlButtonPopup, 0, "���"): objControl.Style = xtpButtonIconAndCaption
                Set objControl = cbpPopup.CommandBar.Controls.Add(xtpControlButton, ID_DRAW_LINE, "ֱ��(&L)")
                Set objControl = cbpPopup.CommandBar.Controls.Add(xtpControlButton, ID_DRAW_MLINE, "����(&Z)")
                Set objControl = cbpPopup.CommandBar.Controls.Add(xtpControlButton, ID_DRAW_RECT, "����(&R)")
                Set objControl = cbpPopup.CommandBar.Controls.Add(xtpControlButton, ID_DRAW_MRECT, "�����(&W)")
                Set objControl = cbpPopup.CommandBar.Controls.Add(xtpControlButton, ID_DRAW_CIRCLE, "��Բ(&C)")
            Set objControl = .Add(xtpControlButton, ID_DRAW_TEXT, "����(&T)")
            Set objControl = .Add(xtpControlButton, ID_DRAW_SEQUENCENUMBER, "˳����(&N)")
            Set objControl = .Add(xtpControlButton, ID_DRAW_CLEARNUMBERS, "���˳����(&K)")
            Set objControl = .Add(xtpControlButton, ID_DRAW_DELETE, "ɾ�����(&D)"): objControl.IconId = 325
            Set objControl = .Add(xtpControlButton, ID_DRAW_RESET, "��ձ��(&R)"):             objControl.BeginGroup = True
            
            Set cbpPopup = .Add(xtpControlButtonPopup, ID_DRAW_FILLSTYLE, "�����ʽ"):          cbpPopup.Style = xtpButtonIconAndCaption
                cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_DRAW_FILLNONE, "�����"
                cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_DRAW_FILLALL, "ʵ�����"
                cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_DRAW_FILLH, "�������"
                cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_DRAW_FILLV, "�������"
                cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_DRAW_FILLHV, "�������"
                cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_DRAW_FILLR, "��б�����"
                cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_DRAW_FILLL, "��б�����"
                cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_DRAW_FILLLR, "���������"
            
            Set cbpPopup = .Add(xtpControlButtonPopup, ID_DRAW_LINESTYLE, "����"):          cbpPopup.Style = xtpButtonIconAndCaption
                cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_DRAW_LINECONTINUE, "ʵ��"
                cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_DRAW_LINEDOT, "����"
                cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_DRAW_LINEDASH, "����"
                cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_DRAW_LINEDASHDOT, "�㻮��"
                cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_DRAW_LINEDASHDOT2, "��㻮��"
            
    '        Set cbpPopup = .Add(xtpControlButtonPopup, ID_DRAW_LINEWIDTH, "�߿�"):          cbpPopup.Style = xtpButtonIconAndCaption
                cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_DRAW_LINEWIDTH1, "1�����"
                cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_DRAW_LINEWIDTH2, "2�����"
                cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_DRAW_LINEWIDTH3, "3�����"
                cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_DRAW_LINEWIDTH4, "4�����"
                cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_DRAW_LINEWIDTH5, "5�����"
            
            Set cbpPopup = .Add(xtpControlSplitButtonPopup, ID_DRAW_FILLCOLOR, "�����ɫ")
            Set objCustControl = cbpPopup.CommandBar.Controls.Add(xtpControlCustom, 0, ""): objCustControl.Handle = CFillColor.hWnd
            
            Set cbpPopup = .Add(xtpControlSplitButtonPopup, ID_DRAW_LINECOLOR, "������ɫ")
            Set objCustControl = cbpPopup.CommandBar.Controls.Add(xtpControlCustom, 0, ""): objCustControl.Handle = CLineColor.hWnd
    '
    '        Set cbpPopup = .Add(xtpControlSplitButtonPopup, ID_DRAW_FONTCOLOR, "������ɫ")             '��ʱδ�õ�
    '        Set objCustControl = cbpPopup.CommandBar.Controls.Add(xtpControlCustom, 0, ""): objCustControl.Handle = CForeColor.hWnd
            Set objControl = .Add(xtpControlButton, ID_EDIT_DELETE, "���ͼƬ"): objControl.BeginGroup = True
            If UserControl.Extender.Tag = "�ο�ͼ" Then Set objControl = .Add(xtpControlButton, ID_INSERT_PICTURE, "��ͼ(&C)")
        End With
    End If
End Sub
Private Sub DesignDraw(Optional ByVal Stand As Boolean)
    On Error Resume Next
    Err.Clear
    Debug.Print 1 / 0
    If Err.Number <> 0 Or Stand Then '��ƻ�����
        Err.Clear
        picBG.Move 0, 0, ScaleWidth, ScaleHeight: picBG.AutoRedraw = True
        picDraw.Move 0, 0, ScaleWidth, ScaleHeight: picDraw.AutoRedraw = True: Set picDraw.Picture = New StdPicture
        picDraw.PaintPicture picBuff.Picture, 0, 0, ScaleWidth / 15, ScaleHeight / 15
        TextOut picDraw, "���ջ�", 0, ScaleHeight / 2 / 15 - 18, 15, ScaleHeight / 15, Nothing
        TextOut picDraw, Format(Now, "HH:mm:ss"), 15, ScaleHeight / 2 / 15 + 15, ScaleWidth / 15, ScaleHeight / 2 / 15 + 30, Nothing
    Else
        Set picBuff.Picture = New StdPicture
        picBuff.Move ScaleWidth, ScaleHeight
        picBG.Move 0, 0, ScaleWidth, ScaleHeight
        picDraw.Move 0, 0, ScaleWidth, ScaleHeight
    End If
    Err.Clear
End Sub
Private Sub ReDrawPicMarks(Optional blnRaisepOk As Boolean = True)
Dim i As Integer, ary As Variant, srPic As StdPicture        'ԴͼƬ����
    Screen.MousePointer = 11
    Set srPic = mDoc.Pictures("K" & mDoc.Cells(mselKey).PictureKey).OrigPic
    Set picDraw.Picture = New StdPicture
    picDraw.PaintPicture srPic, 0, 0, picDraw.Width / 15, picDraw.Height / 15 '��ԭͼ����
    LockWindowUpdate picDraw.hWnd
    
    If mDoc.Cells(mselKey).PicMarkKey <> "" Then '��ԭͼ�ϻ���
        ary = Split(mDoc.Cells(mselKey).PicMarkKey, "|")
        For i = 1 To UBound(ary)
            ShowPicMark picDraw, mDoc.PicMarks("K" & ary(i))
        Next
        Set picDraw.Picture = picDraw.Image
    End If
    LockWindowUpdate 0
    picDraw.Refresh
    If picDraw.Visible Then picDraw.SetFocus
    Screen.MousePointer = 0
    If blnRaisepOk Then RaiseEvent pOk
End Sub

Private Sub SetCursor(ByVal ID As Long)
'   ��;��  ���õ�ǰ����ꡣ
    picDraw.MousePointer = 99
    Select Case ID
    Case ID_DRAW_SELECT
        Set picDraw.MouseIcon = imgCur.ListImages("Sel").Picture
    Case ID_DRAW_MOVE
        Set picDraw.MouseIcon = imgCur.ListImages("Move").Picture
    Case ID_DRAW_LINE, ID_DRAW_MLINE, ID_DRAW_RECT, ID_DRAW_MRECT, ID_DRAW_CIRCLE
        Set picDraw.MouseIcon = imgCur.ListImages("Pen").Picture
    Case ID_DRAW_TEXT
        Set picDraw.MouseIcon = imgCur.ListImages("Text").Picture
    Case ID_DRAW_SEQUENCENUMBER
        Set picDraw.MouseIcon = imgCur.ListImages("Pen").Picture
    Case Else
        Set picDraw.MouseIcon = imgCur.ListImages("Sel").Picture
    End Select

End Sub
Public Sub GetCurDrawMode()
    'ûֵʱ����ֵ
    If mlngDrawModeID = 0 Then mlngDrawModeID = ID_DRAW_SELECT          '��ǰ��ͼģʽ
    If mlngForeColor = 0 Then mlngForeColor = 0            '��ǰѡ�е�������ɫ
    If mlngFillColor = 0 Then mlngFillColor = 0            '��ǰѡ��������ɫ
    If mlngLineColor = 0 Then mlngLineColor = 0            '��ǰѡ����������ɫ
    If mlngFillStyleID = 0 Then mlngFillStyleID = ID_DRAW_FILLNONE       '��ǰѡ�е������ʽ
    If mlngLineWidthID = 0 Then mlngLineWidthID = ID_DRAW_LINEWIDTH1     '��ǰѡ�е��߿�
    If mlngLineStyleID = 0 Then mlngLineStyleID = ID_DRAW_LINECONTINUE   '��ǰѡ�е�����
    
    SetCursor mlngDrawModeID '���������
    SetDrawStyle picDraw.hdc '���õ�ǰ��ͼģʽ�����ʡ���ˢ��
End Sub
Private Sub SetDrawStyle(hdc As Long)
'��;��  ���ݽ���״̬���õ�ǰ�Ļ��ʵĻ�ˢ��
Dim bytPenW As Byte, vBrush As LOGBRUSH, lngPen As Long, lngBrush As Long

    '�����ԭ�л��ʡ���ˢ
    If glngBrush <> 0 Then DeleteObject glngBrush
    If glngPen <> 0 Then DeleteObject glngPen

    '���������߿�
    Select Case mlngLineWidthID
        Case ID_DRAW_LINEWIDTH1
            bytPenW = 1
        Case ID_DRAW_LINEWIDTH2
            bytPenW = 2
        Case ID_DRAW_LINEWIDTH3
            bytPenW = 3
        Case ID_DRAW_LINEWIDTH4
            bytPenW = 4
        Case ID_DRAW_LINEWIDTH5
            bytPenW = 5
        Case Else
            bytPenW = 1
    End Select
    gcurPenWidth = bytPenW '��¼ԭʼ����

    gcurPenColor = mlngLineColor

    Select Case mlngLineStyleID '����
        Case ID_DRAW_LINECONTINUE
            gcurPenStyle = PS_SOLID
            lngPen = CreatePen(PS_SOLID, bytPenW, mlngLineColor)
        Case ID_DRAW_LINEDOT
            gcurPenStyle = PS_DOT
            lngPen = CreatePen(PS_DOT, bytPenW, mlngLineColor)
        Case ID_DRAW_LINEDASH
            gcurPenStyle = PS_DASH
            lngPen = CreatePen(PS_DASH, bytPenW, mlngLineColor)
        Case ID_DRAW_LINEDASHDOT
            gcurPenStyle = PS_DASHDOT
            lngPen = CreatePen(PS_DASHDOT, bytPenW, mlngLineColor)
        Case ID_DRAW_LINEDASHDOT2
            gcurPenStyle = PS_DASHDOTDOT
            lngPen = CreatePen(PS_DASHDOTDOT, bytPenW, mlngLineColor)
    End Select
    glngPen = SelectObject(picDraw.hdc, lngPen)

    '��ˢ
    vBrush.lbColor = mlngFillColor
    gcurFillColor = vBrush.lbColor
    If mlngFillStyleID = ID_DRAW_FILLNONE Then
        vBrush.lbStyle = BS_NULL
        gcurFillStyle = -1
    ElseIf mlngFillStyleID = ID_DRAW_FILLALL Then
        vBrush.lbStyle = BS_SOLID
        gcurFillStyle = -2
    Else
        vBrush.lbStyle = BS_HATCHED
        If mlngFillStyleID = ID_DRAW_FILLH Then
            vBrush.lbHatch = HS_HORIZONTAL '====
        ElseIf mlngFillStyleID = ID_DRAW_FILLV Then
            vBrush.lbHatch = HS_VERTICAL '||||
        ElseIf mlngFillStyleID = ID_DRAW_FILLHV Then
            vBrush.lbHatch = HS_CROSS '++++
        ElseIf mlngFillStyleID = ID_DRAW_FILLL Then
            vBrush.lbHatch = HS_FDIAGONAL '\\\\
        ElseIf mlngFillStyleID = ID_DRAW_FILLR Then
            vBrush.lbHatch = HS_BDIAGONAL '////
        ElseIf mlngFillStyleID = ID_DRAW_FILLLR Then
            vBrush.lbHatch = HS_DIAGCROSS 'XXXX
        End If
        gcurFillStyle = vBrush.lbHatch
    End If
    lngBrush = CreateBrushIndirect(vBrush)
    glngBrush = SelectObject(picDraw.hdc, lngBrush)
End Sub

Public Sub DeleteSelectedMarks(Optional ByVal blnDelNum As Boolean = False, Optional ByVal blnDelAll As Boolean = False)
'��;��  ɾ����ǰѡ�еı�ǡ�
'������blnDelNum���˳����,blnDelAll������б��,����������ΪFalseʱɾ��ѡ�б��
Dim arrSel As Variant, ary As Variant
Dim i As Integer, j As Integer, strTmp As String

    arrSel = Array()
    If Not (blnDelNum = False Or blnDelAll = False) Then
        If mlngSelectedCount = 0 Or picDraw.Visible = False Then Exit Sub
    End If
    
    If mDoc.Cells(mselKey).PicMarkKey <> "" Then
        ary = Split(mDoc.Cells(mselKey).PicMarkKey, "|")
        For i = 1 To UBound(ary)
            If blnDelNum Then                       '�����Ǳ��
                If mDoc.PicMarks("K" & ary(i)).���� = 6 Then
                    ReDim Preserve arrSel(UBound(arrSel) + 1)
                    arrSel(UBound(arrSel)) = "K" & ary(i)
                Else
                    strTmp = strTmp & "|" & ary(i)
                End If
            ElseIf blnDelAll Then                   '��ձ��
                ReDim Preserve arrSel(UBound(arrSel) + 1)
                arrSel(UBound(arrSel)) = "K" & ary(i)
                strTmp = ""
            Else                                    '���ѡ�б��
                If mDoc.PicMarks("K" & ary(i)).ѡ�� Then
                    ReDim Preserve arrSel(UBound(arrSel) + 1)
                    arrSel(UBound(arrSel)) = "K" & ary(i)
                Else
                    strTmp = strTmp & "|" & ary(i)
                End If
            End If
        Next
    Else
        Exit Sub
    End If

    If MsgBox("ȷ��Ҫɾ��" & IIf(blnDelNum, "˳����", IIf(blnDelAll, "����", "ѡ�е� " & UBound(arrSel) + 1 & " ��")) & "�����", vbOKCancel + vbQuestion, gstrSysName) = vbCancel Then Exit Sub

    mDoc.Cells(mselKey).PicMarkKey = strTmp
'    For i = 0 To UBound(arrSel)'�����ڴ��б����ɾ������ΪҪ�õ���������
'        mDoc.PicMarks.Remove arrSel(i)
'    Next
    
    mlngSelectedCount = 0
    Call ReDrawPicMarks     'ˢ�½����
    picDraw.Visible = True: picDraw.SetFocus
End Sub
Private Sub ChangeLineAndReDraw(ByVal bType As Byte)
'����:ѡ��1��䷽ʽ��2���ͣ�3�߿�,4���ɫ,5��ɫ��6����ɫ���ػ�ͼ��
'����:�ֱ��Ӧ���������е����
Dim ary As Variant, i As Integer
    If mDoc.Cells(mselKey).PicMarkKey = "" Then Exit Sub
    
    If mlngSelectedCount > 0 Then
        ary = Split(mDoc.Cells(mselKey).PicMarkKey, "|")
        For i = 1 To UBound(ary)
            If mDoc.PicMarks("K" & ary(i)).ѡ�� Then
                Select Case bType
                    Case 1
                        mDoc.PicMarks("K" & ary(i)).��䷽ʽ = gcurFillStyle
                    Case 2
                        mDoc.PicMarks("K" & ary(i)).���� = gcurPenStyle
                    Case 3
                        mDoc.PicMarks("K" & ary(i)).�߿� = gcurPenWidth
                    Case 4
                        mDoc.PicMarks("K" & ary(i)).���ɫ = mlngFillColor
                    Case 5
                        mDoc.PicMarks("K" & ary(i)).����ɫ = mlngLineColor
                    Case 6
                        mDoc.PicMarks("K" & ary(i)).����ɫ = mlngForeColor
                End Select
            End If
        Next
        Call ReDrawPicMarks     'ˢ�½����
    End If
End Sub
Private Sub CFillColor_pOK(ByVal ControlSelf As Boolean)
    mlngFillColor = IIf(CFillColor.Color = tomAutoColor, CFillColor.AutoColor, CFillColor.Color)
    Call ChangeLineAndReDraw(4)
    SendKeys "{ESCAPE}"
    If ControlSelf Then SendKeys "{ESCAPE}"
End Sub
Private Sub CForeColor_pOK(ByVal ControlSelf As Boolean)
    mlngForeColor = IIf(CForeColor.Color = tomAutoColor, CForeColor.AutoColor, CForeColor.Color)
    Call ChangeLineAndReDraw(6)
    SendKeys "{ESCAPE}"
    If ControlSelf Then SendKeys "{ESCAPE}"
End Sub
Private Sub CLineColor_pOK(ByVal ControlSelf As Boolean)
    mlngLineColor = IIf(CLineColor.Color = tomAutoColor, CLineColor.AutoColor, CLineColor.Color)
    Call ChangeLineAndReDraw(5)
    SendKeys "{ESCAPE}"
    If ControlSelf Then SendKeys "{ESCAPE}"
End Sub
Private Function GetMaxNum() As Long
'��ȡ�Զ���ŵ����ֵ
Dim ary As Variant, i As Integer, j As Integer
    If mDoc.Cells(mselKey).PicMarkKey = "" Then GetMaxNum = 1: Exit Function
    
    ary = Split(mDoc.Cells(mselKey).PicMarkKey, "|")
    For i = 1 To UBound(ary)
        If mDoc.PicMarks("K" & ary(i)).���� = 6 Then
            If j < CLng(mDoc.PicMarks("K" & ary(i)).����) Then j = CLng(mDoc.PicMarks("K" & ary(i)).����)
        End If
    Next
    GetMaxNum = j + 1
End Function
Private Sub GetFitTxtSize(objMain As Object, strText As String, Optional ByRef Width As Long, Optional ByRef Height As Long, Optional ByRef LineHeight As Long)
'��;��  �����ı���ǰ���ݵĺ��ʳߴ硣
'���أ�w,h�����ߴ�,h2���и߶�
    With objMain
        picTxt.FontName = .FontName
        picTxt.FontSize = .FontSize
        picTxt.FontBold = .FontBold
        picTxt.FontItalic = .FontItalic
        picTxt.FontUnderline = .FontUnderline
        picTxt.FontStrikethru = .FontStrikethru
        If strText = "" Then
            Width = picTxt.TextWidth("AA")
            Height = picTxt.TextHeight("A")
        Else
            Width = picTxt.TextWidth(strText & "A")
            Height = picTxt.TextHeight(strText)
        End If
        LineHeight = picTxt.TextHeight("A")
    End With
End Sub
Public Sub FinishInputText()
'��;��  ��ɵ�ǰ��������
Dim lngKey As Long, tmpFont As New StdFont
    If txt.Visible Then
        '������״̬תΪȷ�����벢�˳�
        If Trim(Replace(txt.Text, vbCrLf, "")) <> "" Then
            '����������
            lngKey = mDoc.PicMarks.Add
            With mDoc.PicMarks("K" & lngKey)
                .���� = 0
                .���� = txt.Text
                .���� = txt.Font.Name
                .X1 = txt.Left
                .Y1 = txt.Top
                .X2 = txt.Left + txt.Width
                .Y2 = txt.Top + txt.Height
                Set tmpFont = txt.Font
                TextOut picDraw, .����, .X1, .Y1, .X2, .Y2, tmpFont
            End With
            mDoc.Cells(mselKey).PicMarkKey = mDoc.Cells(mselKey).PicMarkKey & "|" & lngKey
        End If
        txt.Text = ""
        txt.Visible = False
        picTxt.Visible = False
        RaiseEvent pOk
    End If
End Sub
Private Sub txt_Change()
    Dim w As Long, h2 As Long
    Dim lngLines As Long

    Call GetFitTxtSize(txt, txt.Text, w, , h2)

    If txt.Left + w + 10 <= picDraw.ScaleWidth Then
        txt.Width = w + 10
        picTxt.Left = txt.Left + txt.Width - picTxt.Width / 2
    End If

    lngLines = SendMessage(txt.hWnd, EM_GETLINECOUNT, 0, 0)
    txt.Height = lngLines * h2 + 6
    picTxt.Top = txt.Top - picTxt.Height / 2
    RaiseEvent pOk
End Sub
Private Sub txt_KeyPress(KeyAscii As Integer)
    Dim h2 As Long, lngLines As Long

    If InStr("',%?&", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Beep: Exit Sub '�Ƿ�
    If zlCommFun.ActualLen(txt.Text & Chr(KeyAscii)) > txt.MaxLength Then KeyAscii = 0: Beep: Exit Sub  '����

    If KeyAscii >= 32 Or KeyAscii = 13 Or KeyAscii < 0 Then
        txtTmp.FontSize = txt.FontSize
        txtTmp.FontName = txt.FontName
        txtTmp.FontBold = txt.FontBold
        txtTmp.FontItalic = txt.FontItalic
        txtTmp.FontUnderline = txt.FontUnderline
        txtTmp.FontStrikethru = txt.FontStrikethru
        txtTmp.Width = txt.Width
        txtTmp.Text = Left(txt.Text, txt.SelStart) & IIf(KeyAscii = 13, vbCrLf, Chr(KeyAscii)) & Mid(txt.Text, txt.SelStart + txt.SelLength + 1)
        lngLines = SendMessage(txtTmp.hWnd, EM_GETLINECOUNT, 0, 0)
        Call GetFitTxtSize(txt, "A", , , h2)
        If txt.Top + lngLines * h2 + 6 > picDraw.ScaleHeight Then KeyAscii = 0: Beep
    End If
End Sub

Private Sub txt_Validate(Cancel As Boolean)
    If txt.Left + txt.Width > picDraw.ScaleWidth Or txt.Top + txt.Height > picDraw.Height Then
        Cancel = True
        MsgBox "�����ɼ���Χ��ʾ�����ɱ༭��", vbInformation, gstrSysName
        txt.Visible = False
    End If
End Sub
Public Sub SaveSelectedMarks(x As Long, y As Long)
'��;��  ����ѡ�б�ǵ��������ꡣ
    'objMarks=�����е�ǰ��Ŀ�ı��ͼ����
    'X,Y Ϊ����ƫ��
Dim arrTmp() As String, arrXY() As POINTAPI, aryMark As Variant
Dim i As Integer, j As Integer, strTmp As String
    
    If mDoc.Cells(mselKey).PicMarkKey = "" Then Exit Sub
    aryMark = Split(mDoc.Cells(mselKey).PicMarkKey, "|")
    
    For i = 1 To UBound(aryMark)
        With mDoc.PicMarks("K" & aryMark(i))
            If .ѡ�� Then
                Select Case .����
                    Case 1, 3, 5    '0 �ı�  1 ����  3  ���� 5  Բ
                        .X1 = .X1 + x
                        .Y1 = .Y1 + y
                        .X2 = .X2 + x
                        .Y2 = .Y2 + y
                    Case 2, 4 '����
                        arrTmp = Split(.�㼯, ";")
                        ReDim Preserve arrXY(UBound(arrTmp)) As POINTAPI
                        For j = 0 To UBound(arrTmp)
                            arrXY(j).x = CLng(Split(arrTmp(j), ",")(0)) + x
                            arrXY(j).y = CLng(Split(arrTmp(j), ",")(1)) + y
                        Next
                        strTmp = ""
                        For j = 0 To UBound(arrXY)
                            If j = 0 Then
                                strTmp = strTmp & CStr(arrXY(j).x) & "," & CStr(arrXY(j).y)
                            Else
                                strTmp = strTmp & ";" & CStr(arrXY(j).x) & "," & CStr(arrXY(j).y)
                            End If
                        Next j
                        .�㼯 = strTmp              '����㼯����
                End Select
            End If
        End With
    Next
    Call ReDrawPicMarks
End Sub
Private Sub HilightSelectMarks(X1 As Long, Y1 As Long, X2 As Long, Y2 As Long)
'��;��  �ж����б������Щ��ѡ�У���������ʾ��
Dim ary As Variant
    If mDoc.Cells(mselKey).PicMarkKey = "" Then Exit Sub
    ary = Split(mDoc.Cells(mselKey).PicMarkKey, "|")
    mlngSelectedCount = 0
    '����X1��Y1��X2��Y2��ʹ��(X1,Y1)�������Ͻǣ���(X2,Y2)�������½�
    Dim lTmp As Long
    If X1 > X2 Then
        lTmp = X2
        X2 = X1
        X1 = lTmp
    End If
    If Y1 > Y2 Then
        lTmp = Y2
        Y2 = Y1
        Y1 = lTmp
    End If

    Dim i As Long, j As Long, p As Long, q As Long, lSplit As Long, k As Long
    Dim T As Variant
    Dim lX1 As Long, lY1 As Long, lX2 As Long, lY2 As Long, l As Long
    Dim arrXY() As POINTAPI
    Dim A As Double, b As Double, XX As Double, YY As Double

    i = giGetShiftState()
    If i <> vbShiftMask And i <> vbCtrlMask Then
        '����ס Shift ���� Control ��ѡ��ǡ�
        For i = 1 To UBound(ary)
            mDoc.PicMarks("K" & ary(i)).ѡ�� = False
        Next i
        Call ReDrawPicMarks(False)
    End If
    For i = 1 To UBound(ary)
        With mDoc.PicMarks("K" & ary(i))
            If .���� <> 0 Then
                Call SetDrawStyleFromValue(picDraw.hdc, .����ɫ, .����, .�߿�, .���ɫ, .��䷽ʽ)
            End If
            picDraw.DrawMode = vbInvert
            Select Case .����   '0-�ı�,1-����,2,����,3-����,4-�����,5-Բ(��Բ)
            Case 0
                '
            Case 1
                '�ȿ�����߶ζ˵���һ���ھ����У���ѡ��֮��
                If (.X1 > X1 And .X1 < X2 And .Y1 > Y1 And .Y1 < Y2) Or (.X2 > X1 And .X2 < X2 And .Y2 > Y1 And .Y2 < Y2) Then
                    .ѡ�� = Not .ѡ��
                    MoveToEx picDraw.hdc, .X1, .Y1, 0
                    LineTo picDraw.hdc, .X2, .Y2
                    mlngSelectedCount = mlngSelectedCount + 1
                    GoTo LL
                End If
                '�������ֶ�N(100)�ݣ�ȡ���ھ����У����ʾѡ��
                lSplit = IIf(Abs(.X2 - .X1) > Abs(.Y2 - .Y1), Abs(.X2 - .X1), Abs(.Y2 - .Y1))
                For j = 1 To lSplit
                    p = (j / lSplit) * (.X2 - .X1) + .X1
                    q = (j / lSplit) * (.Y2 - .Y1) + .Y1
                    '(p,q) �ھ�����
                    If p > X1 And p < X2 And q > Y1 And q < Y2 Then
                        .ѡ�� = Not .ѡ��
                        MoveToEx picDraw.hdc, .X1, .Y1, 0
                        LineTo picDraw.hdc, .X2, .Y2
                        mlngSelectedCount = mlngSelectedCount + 1
                        GoTo LL
                    End If
                Next j
            Case 2
                '����
                'ͬ�������߷ֶΣ�ȡ���ھ����У����ʾѡ��
                T = Split(.�㼯, ";")
                For k = 1 To UBound(T)
                    lX1 = CLng(Split(T(k - 1), ",")(0))
                    lY1 = CLng(Split(T(k - 1), ",")(1))
                    lX2 = CLng(Split(T(k), ",")(0))
                    lY2 = CLng(Split(T(k), ",")(1))

                    lSplit = IIf(Abs(lX2 - lX1) > Abs(lY2 - lY1), Abs(lX2 - lX1), Abs(lY2 - lY1))
                    For j = 1 To lSplit
                        p = (j / lSplit) * (lX2 - lX1) + lX1
                        q = (j / lSplit) * (lY2 - lY1) + lY1
                        '(p,q) �ھ�����
                        If p > X1 And p < X2 And q > Y1 And q < Y2 Then
                            .ѡ�� = Not .ѡ��
                            ReDim Preserve arrXY(UBound(T))
                            For l = 0 To UBound(T)
                                arrXY(l).x = CLng(Split(T(l), ",")(0))
                                arrXY(l).y = CLng(Split(T(l), ",")(1))
                            Next
                            Polyline picDraw.hdc, arrXY(0), UBound(T) + 1
                            mlngSelectedCount = mlngSelectedCount + 1
                            GoTo LL
                        End If
                    Next j
                Next k
            Case 3
                '����
                If ����������ཻ(X1, Y1, X2, Y2, .X1, .Y1, .X2, .Y2) Then
                    .ѡ�� = Not .ѡ��
                    Rectangle picDraw.hdc, .X1, .Y1, .X2, .Y2
                    mlngSelectedCount = mlngSelectedCount + 1
                    GoTo LL
                End If
            Case 4
                '�����
                T = Split(.�㼯, ";")
                ReDim Preserve arrXY(UBound(T))
                For l = 0 To UBound(T)
                    arrXY(l).x = CLng(Split(T(l), ",")(0))
                    arrXY(l).y = CLng(Split(T(l), ",")(1))
                Next
                If �����������ཻ(X1, Y1, X2, Y2, arrXY) Then
                    .ѡ�� = Not .ѡ��
                    ReDim Preserve arrXY(UBound(T))
                    For l = 0 To UBound(T)
                        arrXY(l).x = CLng(Split(T(l), ",")(0))
                        arrXY(l).y = CLng(Split(T(l), ",")(1))
                    Next
                    Polygon picDraw.hdc, arrXY(0), UBound(T) + 1
                    mlngSelectedCount = mlngSelectedCount + 1
                    GoTo LL
                End If
            Case 5
                '����4������Բ�н��㣡
                If ��������Բ�ཻ(X1, Y1, X2, Y2, .X1, .Y1, .X2, .Y2) Then
                    .ѡ�� = Not .ѡ��
                    mlngSelectedCount = mlngSelectedCount + 1
                    Ellipse picDraw.hdc, .X1, .Y1, .X2, .Y2
                    GoTo LL
                End If
            End Select
        End With
LL:
    Next i
    GetCurDrawMode
End Sub
Public Sub RefreshSelectedMarks(objPic As PictureBox, x As Long, y As Long)
'��;��  ˢ�µ�ǰѡ�е����б��,�ƶ������еĻ滭
Dim arrTmp() As String, arrXY() As POINTAPI
Dim i As Integer, j As Integer, ary As Variant

    If mDoc.Cells(mselKey).PicMarkKey = "" Then Exit Sub
    ary = Split(mDoc.Cells(mselKey).PicMarkKey, "|")

    LockWindowUpdate objPic.hWnd

    objPic.DrawMode = vbInvert

    '������Ԫ��
    For i = 1 To UBound(ary)
        With mDoc.PicMarks("K" & ary(i))
            If .ѡ�� Then
                If .���� <> 0 Then
                    Call SetDrawStyleFromValue(objPic.hdc, .����ɫ, .����, .�߿�, .���ɫ, .��䷽ʽ)
                End If
                Select Case .����
                    Case 1 '����
                        MoveToEx objPic.hdc, .X1 + x, .Y1 + y, 0
                        LineTo objPic.hdc, .X2 + x, .Y2 + y
                    Case 2 '����
                        arrTmp = Split(.�㼯, ";")
                        For j = 0 To UBound(arrTmp)
                            ReDim Preserve arrXY(j)
                            arrXY(j).x = CLng(Split(arrTmp(j), ",")(0)) + x
                            arrXY(j).y = CLng(Split(arrTmp(j), ",")(1)) + y
                        Next
                        Polyline objPic.hdc, arrXY(0), UBound(arrXY) + 1
                    Case 3 '����
                        Rectangle objPic.hdc, .X1 + x, .Y1 + y, .X2 + x, .Y2 + y
                    Case 4 '�����
                        arrTmp = Split(.�㼯, ";")
                        For j = 0 To UBound(arrTmp)
                            ReDim Preserve arrXY(j)
                            arrXY(j).x = CLng(Split(arrTmp(j), ",")(0)) + x
                            arrXY(j).y = CLng(Split(arrTmp(j), ",")(1)) + y
                        Next
                        Polygon objPic.hdc, arrXY(0), UBound(arrXY) + 1
                    Case 5 'Բ
                        Ellipse objPic.hdc, .X1 + x, .Y1 + y, .X2 + x, .Y2 + y
                End Select
            End If
        End With
    Next
    objPic.Refresh

    GetCurDrawMode

    LockWindowUpdate 0
    Screen.MousePointer = 0
End Sub
Public Sub EditPic(Doc As cTableEPR, ParentBar As Object, ByVal selKey As String)
'����ͼƬ�༭��֮ǰ��ֵ
    Set mDoc = Doc 'ȫ�����������
    Set mParentBar = ParentBar '�˵�����
    mselKey = selKey    '��ǰ��Ԫ��Key   Kxxx
    UserControl_Resize
    
    mblnInDrawing = False
    Call GetCurDrawMode '��ȡ��ǰ��ͼģʽ��Ϣ
    Call ReDrawPicMarks '�ػ�ͼ��ͱ��
End Sub

Public Sub ToolExecute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    cbsThis_Execute Control
End Sub
Public Sub ToolUpdate(ByVal Control As XtremeCommandBars.ICommandBarControl)
    cbsThis_Update Control
End Sub
