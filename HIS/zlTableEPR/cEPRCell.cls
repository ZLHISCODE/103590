VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTabCell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'��Ԫ������
Public Enum CellTypeEnum
    cprCTFixtext = 0        '0-�̶��ı�(���ɱ༭)
    cprCTText = 1           '1-�ı���(�ɱ༭�����ı�)
    cprCTElement = 2        '2-��Ҫ��
    cprCTTextElement = 3    '3-�ı����Ҫ�ػ�ϱ༭
    cprCTPicture = 4        '4-�ο�ͼ
    cprCTReportPic = 5      '5-����ͼ
    cprCTSign = 6           '6-ǩ��
    cprCTRowSign = 7        '7-�п�ǩ��
    cprCTColSign = 8        '8-�п�ǩ��
    cprCTPicMark = 9        '9-ͼƬ���
End Enum
'�������
Public Enum HAlignEnum
    HAlignGeneral = 1       '����
    HALignLeft = 2          '��
    HAlignCenter = 3        '��
    HALignRight = 4         '��
End Enum
'�������,��F1�ؼ���ͬ
Public Enum VAlignEnum
    VALignTop = 1           '��
    VAlignCenter = 2        '��
    VALignBottom = 3        '��
End Enum

Private mvarIndex As Long                         '������
Private mvarKey As String
Private mvarID As Double                          'ID <!--��������-->
Private mvar�ļ�ID As Long                      '�ļ�ID<!--�����ļ��б�ID����������Ŀ¼ID�����Ӳ�����¼ID-->
Private mvar������� As Long                    '
Private mvar�������� As Integer                 '��������<!--0-�̶��ı�(���ɱ༭) 1-�ı���(�ɱ༭�����ı�) 2-��Ҫ�� 3-�ı����Ҫ�ػ�ϱ༭ 4-�ο�ͼ 5-����ͼ 6-ǩ�� 7-�п�ǩ�� 8-�п�ǩ�� 9-ͼ����-->
Private mvar�������� As Boolean                 '�����������<!--��������=0ʱ�̶�Ϊ1,��������Ϊ�༭ʱ�õ�Ԫ���Ƿ�ɲ��� 0-�ɲ���,1-���ɲ���.-->
Private mvar�������� As String                  '<!--��|Ϊ�ָ��ĵ�0ά�洢��Ԫ���������Ҫ�ش洢 �Զ�ת���ı�|����;ǩ���洢 ǩ����ʽ|ǩ������|֤��ID|��ʾ��ǩ|ǩ��ʱ��|��ʾʱ��|ǩ��Ҫ�� ;����Ϊ��-->
Private mvar�����д� As Long                    '<!--����������=3ʱ ��Ҫ��Ϊ�ָ����д�,��0��ʼ������Ϊ��-->
Private mvar�����ı� As String                  '<!--�ı���Ҫ�ء�ǩ��Ϊ��ʾ����;ͼ����¼Ϊ ѡ��|����|����|�㼯|X1|Y1|X2|Y2|���ɫ|��䷽ʽ|����ɫ|����ɫ|����|�߿�|������-->
Private mvar��ʼ�� As Long                      '��ʼ�棬ר���ڡ����Ӳ�����¼��
Private mvar��ֹ�� As Long                      '��ֹ�棬ר���ڡ����Ӳ�����¼��

'����
Private mvarRow As Long                         '�У���1��ʼ���
Private mvarCol As Long                         '�У���1��ʼ���
Private mvarWidth As Integer                        '��Ԫ�����
Private mvarHeight As Integer                      '��Ԫ��߶�
Private mvarFontName As String                  '����>����</����>
Private mvarFontSize As Integer                   '<�ֺ�>9</�ֺ�>
Private mvarFontBold As Boolean                 '<����>False</����>
Private mvarFontItalic As Boolean               '<б��>False</б��>
Private mvarFontUnderline As Boolean            '<�»���>False</�»���>
Private mvarFontStrikeout As Boolean            '<ɾ����>False</ɾ����>
Private mvarFontColor   As OLE_COLOR                 '<������ɫ>vbblack</������ɫ>
Private mvarHAlignment As Integer            '<�������>F1HAlignCenter</�������>
Private mvarVAlignment As Integer            '<�������>F1VAlignCenter</�������>
Private mvarCellLineTop As F1BorderStyleConstants               '<�ϱ߿�>1</�ϱ߿�><!--0-�ޱ߿� 1-�б߿� '4-���� 5- ����-->
Private mvarCellLineBottom As F1BorderStyleConstants           '<�±߿�>1</�±߿�>
Private mvarCellLineLeft As F1BorderStyleConstants             '<��߿�>1</��߿�>
Private mvarCellLineRight As F1BorderStyleConstants           '<�ұ߿�>1</�ұ߿�>
Private mvarCellLineTopColor As OLE_COLOR            '<�ϱ߿���ɫ>vbBlack</�ϱ߿���ɫ>
Private mvarCellLineBottomColor As OLE_COLOR          '<�±߿���ɫ>vbBlack</�±߿���ɫ>
Private mvarCellLineLeftColor As OLE_COLOR          '��߿���ɫ>vbBlack</��߿���ɫ>
Private mvarCellLineRightColor As OLE_COLOR         '<�ұ߿���ɫ>vbBlack</�ұ߿���ɫ>
Private mvarMerge As Boolean                   '�Ƿ�ϲ�                                          if IIf(.Merge, InStr(.MergeRange, ";") > 0, True) Then 'ֻ�кϲ���Ԫ���׸���Ǻϲ���Ԫ�������Ч��Ԫ��
Private mvarMergeRange As String               '�ϲ��������Ϸ���Ԫ��(����,����);�ϲ��������·���Ԫ��(����,����),��:1,2;3,4��ʾ��1��2�кϲ���3��4��;ֻ�кϲ�������׸���Ԫ���кϲ���������,������Ԫ��û��
Private mvarTextKey As String                     '��ϱ༭ʱ��|�ָ����ı�Key,�ӵ�һά���
Private mvarElementKey As String                  'Ҫ�ص�Key,��ϱ༭ʱ��|�ָ��Ķ��Ҫ��Key,�ӵ�һά���
Private mvarPictureKey As String                  'ͼƬ��Key
Private mvarSignKey As String                     'ǩ����Key
Private mvarPicMarkKey As String                  'ͼƬ��ǵ�Key,��|�ָ��Ķ�����Key,�ӵ�һά���
Private mvarProtected As Boolean        '�Ƿ񱣻�
Private mvarToolTipText As String       '��ʾ�ı�

Public Property Get Index() As Long                       '������
    Index = mvarIndex
End Property
Public Property Let Index(ByVal vData As Long)                        '������
    mvarIndex = vData
    mvar������� = mvarIndex
End Property
Public Property Get Key() As String
    Key = mvarKey
End Property
Public Property Let Key(ByVal vData As String)
    mvarKey = vData
End Property
Public Property Get ID() As Double                        'ID
    ID = mvarID
End Property
Public Property Let ID(ByVal vData As Double)                         'ID
    mvarID = vData
End Property
Public Property Get �ļ�ID() As Long                    '�ļ�ID
    �ļ�ID = mvar�ļ�ID
End Property
Public Property Let �ļ�ID(ByVal vData As Long)                    '�ļ�ID
    mvar�ļ�ID = vData
End Property
Public Property Get �������() As Long                    '�ļ�ID
    ������� = mvar�������
End Property
Public Property Let �������(ByVal vData As Long)                    '�ļ�ID
    mvar������� = vData
End Property
Public Property Get ��������() As Integer               '��������<!--0-�̶��ı�(���ɱ༭) 1-�ı���(�ɱ༭�����ı�) 2-��Ҫ�� 3-�ı����Ҫ�ػ�ϱ༭ 4-�ο�ͼ 5-����ͼ 6-ǩ�� 7-�п�ǩ�� 8-�п�ǩ�� 9-ͼ����-->
    �������� = mvar��������
End Property
Public Property Let ��������(ByVal vData As Integer)               '��������<!--0-�̶��ı�(���ɱ༭) 1-�ı���(�ɱ༭�����ı�) 2-��Ҫ�� 3-�ı����Ҫ�ػ�ϱ༭ 4-�ο�ͼ 5-����ͼ 6-ǩ�� 7-�п�ǩ�� 8-�п�ǩ�� 9-ͼ����-->
    mvar�������� = vData
End Property
Public Property Get ��������() As Boolean               '����������� <!--��������=0ʱ�̶�Ϊ1,��������Ϊ�༭ʱ�õ�Ԫ���Ƿ�ɲ��� 0-�ɲ���,1-���ɲ���.-->
    �������� = mvar��������
End Property
Public Property Let ��������(ByVal vData As Boolean)               '�����������
    mvar�������� = vData
End Property
Public Property Get ��������() As String                '<!--��|Ϊ�ָ��ĵ�0ά�洢��Ԫ���������Ҫ�ش洢 �Զ�ת���ı�|����;ǩ���洢 ǩ����ʽ|ǩ������|֤��ID|��ʾ��ǩ|ǩ��ʱ��|��ʾʱ��|ǩ��Ҫ�� ;����Ϊ��-->
    �������� = mvar��������
End Property
Public Property Let ��������(ByVal vData As String)                '<!--��|Ϊ�ָ��ĵ�0ά�洢��Ԫ���������Ҫ�ش洢 �Զ�ת���ı�|����;ǩ���洢 ǩ����ʽ|ǩ������|֤��ID|��ʾ��ǩ|ǩ��ʱ��|��ʾʱ��|ǩ��Ҫ�� ;����Ϊ��-->
    mvar�������� = vData
End Property
Public Property Get �����д�() As Long                  '<!--����������=3ʱ ��Ҫ��Ϊ�ָ����д�,��0��ʼ������Ϊ��-->
    �����д� = mvar�����д�
End Property
Public Property Let �����д�(ByVal vData As Long)                  '<!--����������=3ʱ ��Ҫ��Ϊ�ָ����д�,��0��ʼ������Ϊ��-->
    mvar�����д� = vData
End Property
Public Property Get �����ı�() As String                '<!--�ı���Ҫ�ء�ǩ��Ϊ��ʾ����;ͼ����¼Ϊ ѡ��|����|����|�㼯|X1|Y1|X2|Y2|���ɫ|��䷽ʽ|����ɫ|����ɫ|����|�߿�|������-->
    �����ı� = mvar�����ı�
End Property
Public Property Let �����ı�(ByVal vData As String)                '<!--�ı���Ҫ�ء�ǩ��Ϊ��ʾ����;ͼ����¼Ϊ ѡ��|����|����|�㼯|X1|Y1|X2|Y2|���ɫ|��䷽ʽ|����ɫ|����ɫ|����|�߿�|������-->
    mvar�����ı� = vData
End Property
Public Property Get ��ʼ��() As Long                    '��ʼ�棬ר���ڡ����Ӳ�����¼��
    ��ʼ�� = mvar��ʼ��
End Property
Public Property Let ��ʼ��(ByVal vData As Long)                    '��ʼ�棬ר���ڡ����Ӳ�����¼��
    mvar��ʼ�� = vData
End Property
Public Property Get ��ֹ��() As Long                    '��ֹ�棬ר���ڡ����Ӳ�����¼��
    ��ֹ�� = mvar��ֹ��
End Property
Public Property Let ��ֹ��(ByVal vData As Long)                    '��ֹ�棬ר���ڡ����Ӳ�����¼��
    mvar��ֹ�� = vData
End Property
Public Property Get Row() As Long                       '�У���1��ʼ���
    Row = mvarRow
End Property
Public Property Let Row(ByVal vData As Long)                       '�У���1��ʼ���
    mvarRow = vData
End Property
Public Property Get Col() As Long                       '�У���1��ʼ���
    Col = mvarCol
End Property
Public Property Let Col(ByVal vData As Long)                       '�У���1��ʼ���
    mvarCol = vData
End Property
Public Property Get Width() As Integer                      '��Ԫ�����
    Width = mvarWidth
End Property
Public Property Let Width(ByVal vData As Integer)                     '��Ԫ�����
    mvarWidth = vData
End Property
Public Property Get Height() As Integer                    '��Ԫ��߶�
    Height = mvarHeight
End Property
Public Property Let Height(ByVal vData As Integer)                    '��Ԫ��߶�
    mvarHeight = vData
End Property
Public Property Get FontName() As String                '����>����</����>
    FontName = mvarFontName
End Property
Public Property Let FontName(ByVal vData As String)                '����>����</����>
     mvarFontName = vData
End Property
Public Property Get FontSize() As Integer                  '<�ֺ�>9</�ֺ�>
    FontSize = GetFontSizeNumber(mvarFontSize)
End Property
Public Property Let FontSize(ByVal vData As Integer)                '<�ֺ�>9</�ֺ�>
    mvarFontSize = vData
End Property
Public Property Get FontBold() As Boolean               '<����>False</����>
    FontBold = mvarFontBold
End Property
Public Property Let FontBold(ByVal vData As Boolean)               '<����>False</����>
    mvarFontBold = vData
End Property
Public Property Get FontItalic() As Boolean             '<б��>False</б��>
    FontItalic = mvarFontItalic
End Property
Public Property Let FontItalic(ByVal vData As Boolean)             '<б��>False</б��>
    mvarFontItalic = vData
End Property
Public Property Get FontUnderline() As Boolean          '<�»���>False</�»���>
    FontUnderline = mvarFontUnderline
End Property
Public Property Let FontUnderline(ByVal vData As Boolean)          '<�»���>False</�»���>
    mvarFontUnderline = vData
End Property
Public Property Get FontStrikeout() As Boolean          '<ɾ����>False</ɾ����>
    FontStrikeout = mvarFontStrikeout
End Property
Public Property Let FontStrikeout(ByVal vData As Boolean)          '<ɾ����>False</ɾ����>
    mvarFontStrikeout = vData
End Property
Public Property Get FontColor() As OLE_COLOR                 '<������ɫ>vbblack</������ɫ>
    FontColor = mvarFontColor
End Property
Public Property Let FontColor(ByVal vData As OLE_COLOR)                 '<������ɫ>vbblack</������ɫ>
    mvarFontColor = vData
End Property
Public Property Get HAlignment() As HAlignEnum          '<�������>F1HAlignCenter</�������>
    HAlignment = mvarHAlignment
End Property
Public Property Let HAlignment(ByVal vData As HAlignEnum)          '<�������>F1HAlignCenter</�������>
    mvarHAlignment = vData
End Property
Public Property Get VAlignment() As VAlignEnum          '<�������>F1VAlignCenter</�������>
    VAlignment = mvarVAlignment
End Property
Public Property Let VAlignment(ByVal vData As VAlignEnum)          '<�������>F1VAlignCenter</�������>
    mvarVAlignment = vData
End Property
Public Property Get CellLineTop() As F1BorderStyleConstants            '<�ϱ߿�>1</�ϱ߿�><!--0-�ޱ߿� 1-�б߿�-->
    CellLineTop = mvarCellLineTop
End Property
Public Property Let CellLineTop(ByVal vData As F1BorderStyleConstants)            '<�ϱ߿�>1</�ϱ߿�><!--0-�ޱ߿� 1-�б߿�-->
    mvarCellLineTop = vData
End Property
Public Property Get CellLineBottom() As F1BorderStyleConstants         '<�±߿�>1</�±߿�>
    CellLineBottom = mvarCellLineBottom
End Property
Public Property Let CellLineBottom(ByVal vData As F1BorderStyleConstants)         '<�±߿�>1</�±߿�>
    mvarCellLineBottom = vData
End Property
Public Property Get CellLineLeft() As F1BorderStyleConstants           '<��߿�>1</��߿�>
    CellLineLeft = mvarCellLineLeft
End Property
Public Property Let CellLineLeft(ByVal vData As F1BorderStyleConstants)           '<��߿�>1</��߿�>
    mvarCellLineLeft = vData
End Property
Public Property Get CellLineRight() As F1BorderStyleConstants         '<�ұ߿�>1</�ұ߿�>
    CellLineRight = mvarCellLineRight
End Property
Public Property Let CellLineRight(ByVal vData As F1BorderStyleConstants)         '<�ұ߿�>1</�ұ߿�>
    mvarCellLineRight = vData
End Property

Public Property Get CellLineTopColor() As OLE_COLOR          '<�ϱ߿���ɫ>vbBlack</�ϱ߿���ɫ>
    CellLineTopColor = mvarCellLineTopColor
End Property
Public Property Let CellLineTopColor(ByVal vData As OLE_COLOR)          '<�ϱ߿���ɫ>vbBlack</�ϱ߿���ɫ>
    mvarCellLineTopColor = vData
End Property
Public Property Get CellLineBottomColor() As OLE_COLOR        '<�±߿���ɫ>vbBlack</�±߿���ɫ>
    CellLineBottomColor = mvarCellLineBottomColor
End Property
Public Property Let CellLineBottomColor(ByVal vData As OLE_COLOR)        '<�±߿���ɫ>vbBlack</�±߿���ɫ>
    mvarCellLineBottomColor = vData
End Property
Public Property Get CellLineLeftColor() As OLE_COLOR        '��߿���ɫ>vbBlack</��߿���ɫ>
    CellLineLeftColor = mvarCellLineLeftColor
End Property
Public Property Let CellLineLeftColor(ByVal vData As OLE_COLOR)        '��߿���ɫ>vbBlack</��߿���ɫ>
    mvarCellLineLeftColor = vData
End Property
Public Property Get CellLineRightColor() As OLE_COLOR       '<�ұ߿���ɫ>vbBlack</�ұ߿���ɫ>
    CellLineRightColor = mvarCellLineRightColor
End Property
Public Property Let CellLineRightColor(ByVal vData As OLE_COLOR)       '<�ұ߿���ɫ>vbBlack</�ұ߿���ɫ>
    mvarCellLineRightColor = vData
End Property
Public Property Get Merge() As Boolean                 '�Ƿ�ϲ�
    Merge = mvarMerge
End Property
Public Property Let Merge(ByVal vData As Boolean)                 '�Ƿ�ϲ�
    mvarMerge = vData
End Property
Public Property Get MergeRange() As String                 '�ϲ�����
    MergeRange = mvarMergeRange
End Property
Public Property Let MergeRange(ByVal vData As String)                 '�ϲ�����
    mvarMergeRange = vData
End Property
Public Property Get TextKey() As String                'Ҫ�ص�Key
    TextKey = mvarTextKey
End Property
Public Property Let TextKey(ByVal vData As String)                'Ҫ�ص�Key
    mvarTextKey = vData
End Property
Public Property Get ElementKey() As String                'Ҫ�ص�Key
    ElementKey = mvarElementKey
End Property
Public Property Let ElementKey(ByVal vData As String)                'Ҫ�ص�Key
    mvarElementKey = vData
End Property
Public Property Get PictureKey() As String                'ͼƬ��Key
    PictureKey = mvarPictureKey
End Property
Public Property Let PictureKey(ByVal vData As String)                'ͼƬ��Key
    mvarPictureKey = vData
End Property
Public Property Get SignKey() As String                'ǩ����Key
    SignKey = mvarSignKey
End Property
Public Property Let SignKey(ByVal vData As String)                'ǩ����Key
    mvarSignKey = vData
End Property
Public Property Get PicMarkKey() As String                'ͼƬ��ǵ�Key
    PicMarkKey = mvarPicMarkKey
End Property
Public Property Let PicMarkKey(ByVal vData As String)                'ͼƬ��ǵ�Key
    mvarPicMarkKey = vData
End Property
Public Property Get Protected() As Boolean      '�Ƿ񱣻�
    Protected = mvarProtected
End Property
Public Property Let Protected(ByVal vData As Boolean)      '�Ƿ񱣻�
    mvarProtected = vData
End Property
Public Property Get ToolTipText() As String     '��ʾ�ı�
    ToolTipText = mvarToolTipText
End Property
Public Property Let ToolTipText(ByVal vData As String)     '��ʾ�ı�
    mvarToolTipText = vData
End Property
Public Sub Clone(cNode As cTabCell)
Dim i As Integer
    With cNode
        .Width = Me.Width                     '��Ԫ�����
        .Height = Me.Height                    '��Ԫ��߶�
        .FontName = Me.FontName                '����>����</����>
        .FontSize = Me.FontSize                '<�ֺ�>9</�ֺ�>
        .FontBold = Me.FontBold               '<����>False</����>
        .FontItalic = Me.FontItalic             '<б��>False</б��>
        .FontUnderline = Me.FontUnderline          '<�»���>False</�»���>
        .FontStrikeout = Me.FontStrikeout          '<ɾ����>False</ɾ����>
        .FontColor = Me.FontColor                 '<������ɫ>vbblack</������ɫ>
        .HAlignment = Me.HAlignment           '<�������>F1HAlignCenter</�������>
        .VAlignment = Me.VAlignment          '<�������>F1VAlignCenter</�������>
        .CellLineTop = Me.CellLineBottom            '<�ϱ߿�>1</�ϱ߿�><!--0-�ޱ߿� 1-�б߿�-->
        .CellLineBottom = Me.CellLineBottom         '<�±߿�>1</�±߿�>
        .CellLineLeft = Me.CellLineLeft           '<��߿�>1</��߿�>
        .CellLineRight = Me.CellLineRight         '<�ұ߿�>1</�ұ߿�>
        .CellLineTopColor = Me.CellLineBottomColor          '<�ϱ߿���ɫ>vbBlack</�ϱ߿���ɫ>
        .CellLineBottomColor = Me.CellLineBottomColor        '<�±߿���ɫ>vbBlack</�±߿���ɫ>
        .CellLineLeftColor = Me.CellLineLeftColor        '��߿���ɫ>vbBlack</��߿���ɫ>
        .CellLineRightColor = Me.CellLineRightColor       '<�ұ߿���ɫ>vbBlack</�ұ߿���ɫ>
        .ElementKey = Me.ElementKey                'Ҫ�ص�Key
        .PictureKey = Me.PictureKey                'ͼƬ��Key
        .SignKey = Me.SignKey                   'ǩ����Key
        .PicMarkKey = Me.PicMarkKey              'ͼƬ��ǵ�Key
        .Protected = Me.Protected      '�Ƿ񱣻�
        .ToolTipText = Me.ToolTipText     '��ʾ�ı�
        
        .�ļ�ID = Me.�ļ�ID                  '�ļ�ID<!--�����ļ��б�ID����������Ŀ¼ID�����Ӳ�����¼ID-->
        .�������� = Me.��������
        .�������� = Me.��������
        .�������� = Me.��������
        .�����д� = Me.�����д�
        .�����ı� = Me.�����ı�
        .��ʼ�� = Me.��ʼ��
        .��ֹ�� = Me.��ֹ��
    End With
End Sub

'################################################################################################################
'## ���ܣ�  ��ȡ��Ԫ�������ַ���
'################################################################################################################
Public Property Get ��Ԫ����() As String
Dim strR As String
    strR = strR & Me.Row & "|"                        '�У���1��ʼ���
    strR = strR & Me.Col & "|"                        '�У���1��ʼ���
    strR = strR & Me.Width & "|"                      '��Ԫ�����
    strR = strR & Me.Height & "|"                     '��Ԫ��߶�
    strR = strR & Me.FontName & "|"                 '����>����</����>
    strR = strR & Me.FontSize & "|"                 '<�ֺ�>9</�ֺ�>
    strR = strR & Me.FontBold & "|"                 '<����>False</����>
    strR = strR & Me.FontItalic & "|"               '<б��>False</б��>
    strR = strR & Me.FontUnderline & "|"            '<�»���>False</�»���>
    strR = strR & Me.FontStrikeout & "|"            '<ɾ����>False</ɾ����>
    strR = strR & Me.FontColor & "|"                  '<������ɫ>vbblack</������ɫ>
    strR = strR & Me.HAlignment & "|"            '<�������>F1HAlignCenter</�������>
    strR = strR & Me.VAlignment & "|"           '<�������>F1VAlignCenter</�������>
    strR = strR & Me.CellLineTop & "|"              '<�ϱ߿�>1</�ϱ߿�><!--0-�ޱ߿� 1-�б߿�-->
    strR = strR & Me.CellLineBottom & "|"           '<�±߿�>1</�±߿�>
    strR = strR & Me.CellLineLeft & "|"             '<��߿�>1</��߿�>
    strR = strR & Me.CellLineRight & "|"           '<�ұ߿�>1</�ұ߿�>
    strR = strR & Me.CellLineTopColor & "|"           '<�ϱ߿���ɫ>vbBlack</�ϱ߿���ɫ>
    strR = strR & Me.CellLineBottomColor & "|"         '<�±߿���ɫ>vbBlack</�±߿���ɫ>
    strR = strR & Me.CellLineLeftColor & "|"         '��߿���ɫ>vbBlack</��߿���ɫ>
    strR = strR & Me.CellLineRightColor & "|"        '<�ұ߿���ɫ>vbBlack</�ұ߿���ɫ>
    strR = strR & Me.Merge & "|"                   '�Ƿ�ϲ�
    strR = strR & Me.MergeRange & "|"             '�ϲ��������Ϸ���Ԫ��(����,����);�ϲ��������·���Ԫ��(����,����)
    strR = strR & Me.ElementKey & "|"                 'Ҫ�ص�Key
    strR = strR & Me.PictureKey & "|"                 'ͼƬ��Key
    strR = strR & Me.SignKey & "|"                   'ǩ����Key
    strR = strR & Me.PicMarkKey & "|"               'ͼƬ��ǵ�Key
    strR = strR & Me.Protected & "|"        '�Ƿ񱣻�
    strR = strR & Me.ToolTipText & "|"      '��ʾ�ı�
    
    ��Ԫ���� = strR
End Property

'################################################################################################################
'## ���ܣ�  ͨ��һ���ַ��������õ�Ԫ������
'################################################################################################################
Public Property Let ��Ԫ����(ByRef vData As String)
    On Error Resume Next
    Dim T As Variant
    T = Split(vData, "|")
    Me.Row = T(0)                       '�У���1��ʼ���
    Me.Col = T(1)                       '�У���1��ʼ���
    Me.Width = T(2)                     '��Ԫ�����
    Me.Height = T(3)                    '��Ԫ��߶�
    Me.FontName = T(4)                 '����>����</����>
    Me.FontSize = T(5)                 '<�ֺ�>9</�ֺ�>
    Me.FontBold = T(6)               '<����>False</����>
    Me.FontItalic = T(7)             '<б��>False</б��>
    Me.FontUnderline = T(8)          '<�»���>False</�»���>
    Me.FontStrikeout = T(9)          '<ɾ����>False</ɾ����>
    Me.FontColor = T(10)                 '<������ɫ>vbblack</������ɫ>
    Me.HAlignment = T(11)           '<�������>F1HAlignCenter</�������>
    Me.VAlignment = T(12)           '<�������>F1VAlignCenter</�������>
    Me.CellLineTop = T(13)            '<�ϱ߿�>1</�ϱ߿�><!--0-�ޱ߿� 1-�б߿�-->
    Me.CellLineBottom = T(14)         '<�±߿�>1</�±߿�>
    Me.CellLineLeft = T(15)           '<��߿�>1</��߿�>
    Me.CellLineRight = T(16)         '<�ұ߿�>1</�ұ߿�>
    Me.CellLineTopColor = T(17)          '<�ϱ߿���ɫ>vbBlack</�ϱ߿���ɫ>
    Me.CellLineBottomColor = T(18)        '<�±߿���ɫ>vbBlack</�±߿���ɫ>
    Me.CellLineLeftColor = T(19)        '��߿���ɫ>vbBlack</��߿���ɫ>
    Me.CellLineRightColor = T(20)       '<�ұ߿���ɫ>vbBlack</�ұ߿���ɫ>
    Me.Merge = T(21)                 '�Ƿ�ϲ�
    Me.MergeRange = T(22)            '�ϲ��������Ϸ���Ԫ��(����,����);�ϲ��������·���Ԫ��(����,����)
    Me.ElementKey = T(23)                'Ҫ�ص�Key
    Me.PictureKey = T(24)              'ͼƬ��Key
    Me.SignKey = T(25)                   'ǩ����Key
    Me.PicMarkKey = T(26)             'ͼƬ��ǵ�Key
    Me.Protected = T(27)      '�Ƿ񱣻�
    Me.ToolTipText = T(28)      '��ʾ�ı�
End Property
Public Property Get CellTypeName() As String
    Select Case Me.��������
        Case cprCTFixtext
            CellTypeName = "�̶��ı�"
        Case cprCTText
            CellTypeName = "�ı���"
        Case cprCTElement
            CellTypeName = "��Ҫ��"
        Case cprCTTextElement
            CellTypeName = "��ϱ༭"
        Case cprCTPicture
            CellTypeName = "�ο�ͼ"
        Case cprCTReportPic
            CellTypeName = "����ͼ"
        Case cprCTSign
            CellTypeName = "ǩ��"
        Case cprCTRowSign
            CellTypeName = "�п�ǩ��"
        Case cprCTColSign
            CellTypeName = "�п�ǩ��"
    End Select
End Property
Public Function GetSaveSQL(ByVal editType As TabET, arrSQL As Variant) As Boolean
'����:����̶��ı�,���ı�����,������������
    On Error GoTo errHand
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    Select Case editType
    Case TabET_�����ļ�����
        arrSQL(UBound(arrSQL)) = "Zl_�����ļ��ṹ_Update(" & mvarID & "," & mvar�ļ�ID & _
                                    ",Null," & mvar������� & "," & mvar�������� & ",Null," & IIf(mvar��������, 1, 0) & ",'" & _
                                    mvarRow & "," & mvarCol & "|" & mvar�������� & "',Null,'" & Replace(mvar�����ı�, "'", "'" & "||Chr(39)||" & "'") & "')"
    Case TabET_ȫ��ʾ���༭
        arrSQL(UBound(arrSQL)) = "Zl_������������_Update(" & mvarID & "," & mvar�ļ�ID & _
                                    ",Null," & mvar������� & "," & mvar�������� & ",Null," & IIf(mvar��������, 1, 0) & ",'" & _
                                    mvarRow & "," & mvarCol & "|" & mvar�������� & "',Null,'" & Replace(mvar�����ı�, "'", "'" & "||Chr(39)||" & "'") & "')"
    Case TabET_�������༭, TabET_���������
        arrSQL(UBound(arrSQL)) = "Zl_���Ӳ�������_Update(" & mvarID & "," & mvar�ļ�ID & "," & mvar��ʼ�� & "," & mvar��ֹ�� & _
                                    ",Null," & mvar������� & "," & mvar�������� & ",Null," & IIf(mvar��������, 1, 0) & ",'" & _
                                    mvarRow & "," & mvarCol & "|" & mvar�������� & "',Null,'" & Replace(mvar�����ı�, "'", "'" & "||Chr(39)||" & "'") & "')"
    End Select
    GetSaveSQL = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Public Function clText(ByVal clt As cTabTexts, ByVal lSeq As Long) As cTabText
'���������дζ�ȡ����
Dim i As Integer
    For i = 1 To clt.Count
        If clt(i).�����д� = lSeq And InStr(Me.TextKey & "|", "|" & clt(i).Key & "|") > 0 Then
            Set clText = clt(i)
            Exit Function
        End If
    Next
    'û�ҵ�ʱ
    Set clText = New cTabText
End Function
Public Function clElement(ByVal els As cTabElements, ByVal lSeq As Long) As cTabElement
'���������дζ�ȡ����
Dim i As Integer
    For i = 1 To els.Count
        If els(i).�����д� = lSeq And InStr(Me.ElementKey & "|", "|" & els(i).Key & "|") > 0 Then
            Set clElement = els(i)
            Exit Function
        End If
    Next
End Function
Private Sub Class_Initialize()
    mvarFontName = "����"
    mvarFontSize = 9
    mvarHAlignment = HAlignCenter
    mvarVAlignment = VAlignCenter
    mvarCellLineTop = 1
    mvarCellLineBottom = 1
    mvarCellLineLeft = 1
    mvarCellLineRight = 1
    mvarCellLineTopColor = 0
    mvarCellLineBottomColor = 0
    mvarCellLineLeftColor = 0
    mvarCellLineRightColor = 0
    mvar������� = mvarIndex
End Sub
'################################################################################################################
'## ���ܣ�  ������ж����ID�Ƿ���Ч
'##
'## ������  eEditType :��ǰ�ı༭ģʽ
'##
'## ˵����  ���û��ID�������ݿ�����ȡһ��ΨһID�š���ID=0������ǵ�Ԫ����CellȡID
'################################################################################################################
Public Sub CheckValidID(eEditType As TabET)
    Dim i As Long, strTable As String
    Select Case eEditType
    Case TabET_�����ļ�����
        strTable = "�����ļ��ṹ"
        If Me.ID = 0 Then
            If IIf(Me.Merge, InStr(Me.MergeRange, ";") > 0, True) Then  'ֻ�зǺϲ���Ԫ��ͺϲ���Ԫ����׸���Ԫ��ż�¼ID
                Me.ID = zlDatabase.GetNextId(strTable)
            End If
        End If
    Case TabET_ȫ��ʾ���༭
        strTable = "������������"
        If Me.ID = 0 Then
            If IIf(Me.Merge, InStr(Me.MergeRange, ";") > 0, True) Then  'ֻ�зǺϲ���Ԫ��ͺϲ���Ԫ����׸���Ԫ��ż�¼ID
                Me.ID = zlDatabase.GetNextId(strTable)
            End If
        End If
    Case TabET_�������༭, TabET_���������
        strTable = "���Ӳ�������"
        If Me.ID = 0 Then
            If IIf(Me.Merge, InStr(Me.MergeRange, ";") > 0, True) Then  'ֻ�зǺϲ���Ԫ��ͺϲ���Ԫ����׸���Ԫ��ż�¼ID
                Me.ID = GetEPRContentNextId
            End If
        End If
    End Select
    
End Sub
