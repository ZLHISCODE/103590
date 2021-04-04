VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTimeLineControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum E_TICK_VALUE
    TICK_1 = 1
    TICK_2 = 2
    TICK_3 = 3
    TICK_4 = 4
    TICK_6 = 6
    TICK_8 = 8
    TICK_12 = 12
    TICK_24 = 24
End Enum

Public Type DesignInfoTickRange
    DisplayValue As Integer                   '��ʾֵ
    TickStartTime As String                 'ʱ�̿�ʼʱ���һ����0��0��ʼ
End Type

'ʱ���������ʽ
Public Type TimeLineDesignInfo
    BackgroundColor As String                   '������ɫ "200,255,255,255" A(͸����0-255),R(0-255),G(0-255),B(0-255)û�����ã�Ĭ����ɫ����͸�����ַ���Ҳ������Transparent���ַ���Ĭ��Ҳ��ȫ͸��
    DateTitle As String                         '�����б���"����"
    DateTitleColor  As String
    DateTitleFont  As String
    DateStart As String                          '��ʼ���ڣ�С�ڿ�ʼ������������ݲ���ʾ�� ����Ϊ��
    dateEnd   As String                          '�������ڣ�����������������ݲ���ʾ�� ����Ϊ��
    ShowTick As Boolean                           '�Ƿ���ʾʱ����     true
    ShowFullDate As Boolean                        '��������ʾ��������
    TickTitle As String                         'ʱ���б��� ʱ��
    TickTitleColor  As String
    TickTitleFont  As String
    TickWidth As Integer                          'ʱ�̿���   18
    TickRangeListCount As E_TICK_VALUE                'ʱ�̷�Χ��� ȡֵ��Χ{1,2,3,4,6,8,12,24}
    DateFont As String                  '��������<DateFont>����, 9pt</DateFont>
    TickFont As String                  'ʱ������<TickFont>����, 9pt</TickFont>
    MergePeriodWidth As Integer              '��������ϲ�����ʾ����
    EmptyDataMergeDayCount As Integer        '�����������ݺϲ�
    EmptyDataMergePeriodWidth As Integer     '�����ݺϲ�����ʾ����
    PaddingLeft As Integer                    '�ڲ���߾�
    PaddingTop As Integer                    '�ڲ��ϱ߾�
    PaddingRight As Integer                  '�ڲ��ұ߾�
    PaddingBottom As Integer                 '�ڲ��±߾�
    RowHeight As Integer                     'Ĭ���и�

    MeasureTitleWidth As Integer             '�����߱����п���
    GridMinHeight As Integer                 '�������������С�߶�
    TopFixedSmallRowCount As Integer         '��߶�������ʾ��������
    BottomFixedSmallRowCount As Integer      '��ߵײ�����ʾ��������
    GridYSplitCount As Integer               '�������ֱ�ָ���
    GridYSmallSplitCount As Integer          '�������С��ֱ�ָ���
    Height As Integer                        '��߸߶� (Null��ʾ������)
End Type

Public Type DataInfo
    Value As String
    Time As String              'ʵ��:2015-07-22T11:50:25
    RowNumber As Integer        '����������
    TimeEnd As String           'ʵ��:2015-07-22T11:50:25
    Tag  As String
                                '���ڡ�TextColor������BackgroudColor������FontName������FontSize���ڵ㣬��DataInfo��Ĭ�ϼ̳С�DataItem����ֵ����������þ����Լ���ֵ��
    BackgroundColor As String
    TextColor As String         'Black
    Font As String          '�����������<Font>����, 9.75pt</Font>
    RowIndex As String         '�кţ���Ϊ����Ψһ��ʶ��
                                            '���ڡ�LegendType������ShadowLegendType������BalloonLegendType���ڵ㣬��DataInfo��Ĭ�ϼ̳С�DataItem����ֵ����������þ����Լ���ֵ��
    LegendType As String                    '���ݵ���ʾ����{ʵ��Բ������Բ���ֿ���Բ���㡢�桢H����}
    ShadowLegendType As String              '��Ӱ����ʾ����
    BalloonLegendType As String             '�������ʾ����

    NumberValue As String            '�������ֵ
    ShadowValue As String           '��Ӱ���ֵ<ShadowValue xsi:nil="true" />
    BalloonValue As String          '������ֵ
    Group  As String                '��ͬGroup������ǰ��ʾ�����ߡ�
    Tip As String                   '��ͷ�ײ�������Ϣ����Tip�ڵ㣬��ʾ��ʾ���������ݳ��⣩��Ϊ��ʱ��ʾValue�ڵ��ֵ��
    HotspotFont  As String                  '�ȵ�����
    HotspotColor As String                  '�ȵ���ɫ
    ShowHotspotCursor As Boolean            '�ȵ�ָ��
End Type

Public Type DataItem
    GraphType As String
    Title As String
    TitleColor As String        '��߱���֧����ɫ���á�
    TitleFont As String         '��߱���֧���������á�
    BackgroundColor As String
    TextColor As String           'Ĭ������ Black
    Font As String            ' �����������<Font>����, 9.75pt</Font>
    ListData As Collection    'DataItem��DataInfo�ļ���
    
    SplitString As String          '���ڷָ�����
    SplitCount  As Integer
    Unit As String              '��λ����ʾ�ڡ�Title�����棬�����Զ����� ���ڱ������
    MinValue As Integer         '��߿̶���Сֵ
    MaxValue As Integer         '��߿̶����ֵ
    SplitNum As Integer         '�̶ȷָ���
    SplitScale As Integer       '�Զ���ָ��ֵʱ��SplitNum����Ч������������0,1,2,3,4,4.3,4.6,4.9,5.2����
    Color As String              '��ߵ���ߵ���ɫ
    IsDataDynamicExpansion As Boolean     '�Ƿ�̬��չ���ݣ���̬����λ�ã����ݵ�ǰ��ʾ����������ȷ���̶����ֵ����Сֵ
    ShadowTitle As String                  '��Ӱͼ�����ơ�
    BalloonTitle As String                  '����ͼ������
    LegendType As String                    '���ݵ���ʾ����{ʵ��Բ������Բ���ֿ���Բ���㡢�桢H����}
    ShadowLegendType As String              '��Ӱ����ʾ����
    BalloonLegendType As String             '�������ʾ����
    ShadowColor As String                   '���DataItem�ڵ�  û�и�ֵʹ��Color��ɫ
    BalloonColor As String                  '���DataItem�ڵ�  û�и�ֵʹ��Color��ɫ
    StartDate As String                     '��ʼ���ڣ����ڵ�������ֵ��
    EndDate As String                       '�������ڣ�С������ֵ
    FixedTick As String                     '�̶�ʱ��
    EquantTick As String                    '�ȷ�ʱ��ÿ��ʱ�̵Ĵ�С����FixedTick�����ȡ�
    EquantTickUnit As String                '�ȷ�ʱ��ÿ��ʱ�̵ĵ�λ��Сʱ�����ӣ���
    TickWidth  As Integer                   'ʱ�̵Ŀ���
                                            '��������DataItem�ڵ�
    LineColor As String                      '��������ɫ
    IsCollapse As Boolean                   '�Ƿ��۵�
    ShowHotspotEffect As Boolean              '�ȵ�Ч��
    ShowHotspotCursor As Boolean            '�ȵ�ָ��
    HotspotFont  As String                  '�ȵ����� <HotspotFont>����, 9.75pt, style=Bold, Underline</HotspotFont>
    HotspotColor As String                  '�ȵ���ɫ
    BorderColor As String                   '�Ű��ı��߿���ɫ
    GroupPosition  As String                '����λ��:����,����  Ĭ�Ͼ���
    Effect  As String                       'ȡֵ(��,����
    
    ItemTag As String                       '���ڻ������ݱ�ʶ,��ʱ����ؼ��ڵ�ֵ
End Type

Public Type TimeLineData
    colHeaderData As Collection                 '��ͷ
    colFooterData As Collection                 'ҳ������
    colMeasureData As Collection                '�������
    colMeasureVerticalText As Collection        '��ֱ�ı�
    colCustomTick As Collection                 'ʱ���Զ���
    colDataArea  As Collection                  '���������Զ���
End Type
