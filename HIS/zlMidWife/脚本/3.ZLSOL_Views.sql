--��zlsol�û�����
create or replace view v_sol_inf_checkinroom as
Select a.mid, b."�뷿Ŀ��",b."�뷿ʱ��",b."ҽ�Ʋ���",b."������",b."����֪��֪ͨ��",b."����������",b."̥����",b."̥�Ĵ���",b."��Ĥ���",b."�Ƿ��кϲ�֢",b."����",b."��Һ��",b."����ͨ��",b."�ֲ����",b."����ҩ��",b."����",b."������",b."�Ӱ���"
From SOL_INF_CHECKINROOM a,JSON_TABLE(a.Content,'$' columns(
�뷿Ŀ��       Varchar2(50) PATH '$.�뷿Ŀ��',
�뷿ʱ��       Varchar2(50) PATH '$.�뷿ʱ��',
ҽ�Ʋ���       Varchar2(10) PATH '$.ҽ�Ʋ���',
������       Varchar2(10) PATH '$.������',
����֪��֪ͨ��      Varchar2(10) PATH '$.����֪��֪ͨ��',
����������   Varchar2(20) PATH '$.����������',
̥����        Varchar2(10) PATH '$.̥����',
̥�Ĵ���       Varchar2(10) PATH '$.̥�Ĵ���',
��Ĥ���       Varchar2(10) PATH '$.��Ĥ���',
�Ƿ��кϲ�֢       Varchar2(10) PATH '$.�Ƿ��кϲ�֢',
����         Varchar2(50) PATH '$.����',
��Һ��        Varchar2(10) PATH '$.��Һ��',
����ͨ��       Varchar2(10) PATH '$.����ͨ��',
�ֲ����       Varchar2(50) PATH '$.�ֲ����',
����ҩ��       Varchar2(50) PATH '$.����ҩ��',
����         Varchar2(50) PATH '$.����',
������        Varchar2(50) PATH '$.������',
�Ӱ���               Varchar2(50) PATH '$.�Ӱ���'
)) as b;

create or replace view v_sol_inf_checkoutroom as
Select a.mid, b."OUTROOMTIME",b."����״̬",b."ҽ�Ʋ���",b."������",b."����ͨ��",b."�ֲ����",b."��������",b."�����п���",b."�����пڷ��",b."����ˮ��",b."����Ѫ��",b."�����Ѫ",b."���",b."����ҩ��",b."�������Ա�",b."����",b."��Ժ���",b."������",b."�Ӱ���",b."����",b."ҩ��",b."��ע"
From SOL_INF_CheckOutRoom a,JSON_TABLE(a.Content,'$' columns(
OUTROOMTIME     Varchar2(50) PATH '$.OUTROOMTIME',
����״̬     Varchar2(50) PATH '$.����״̬',
ҽ�Ʋ���     Varchar2(10) PATH '$.ҽ�Ʋ���',
������     Varchar2(10) PATH '$.������',
����ͨ��     Varchar2(10) PATH '$.����ͨ��',
�ֲ����     Varchar2(50) PATH '$.�ֲ����',
��������     Varchar2(20) PATH '$.��������',
�����п���   Varchar2(20) PATH '$.�����п���',
�����пڷ�� Varchar2(10) PATH '$.�����пڷ��',
����ˮ��     Varchar2(10) PATH '$.����ˮ��',
����Ѫ��     Varchar2(10) PATH '$.����Ѫ��',
�����Ѫ     Varchar2(10) PATH '$.�����Ѫ',
���       Number(5) PATH '$.���',
����ҩ��     Varchar2(50) PATH '$.����ҩ��',
�������Ա�       Varchar2(10) PATH '$.�������Ա�',
����       Number(4,2) PATH '$.����',
��Ժ���     Varchar2(10) PATH '$.��Ժ���',
������       Varchar2(20) PATH '$.������',
�Ӱ���       Varchar2(20) PATH '$.�Ӱ���',
����         Varchar2(20) PATH '$.����',
ҩ��         Varchar2(50) PATH '$.ҩ��',
��ע         Varchar2(50) PATH '$.��ע'
)) as b;


create or replace view v_sol_inf_delivery as
Select a.Mid, b."����1", b."���̿�ʼʱ��", b."����ȫ��ʱ��", b."̥�����ʱ��", b."̥�����ʱ��", b."��һ����", b."�ڶ�����", b."��������", b."�������",
       b."��������������", b."����", b."��Ĥ��ʽ", b."��Ĥʱ��", b."��ˮ��״", b."��ˮ��", b."��ˮ��ɫ", b."̥�������ʽ", b."̥�̰��뷽ʽ", b."̥��������",
       b."̥��̥Ĥ����", b."̥�����", b."̥����̬", b."̥�̴�С", b."̥������", b."�������", b."�������", b."�����ٽ�", b."���",b."�ƾ�����", b."�����ʽ",
       b."���̥��λ", b."������С", b."������λ", b."�������˳̶�", b."���������п�", b."�������˷��", b."������������", b."�������˳���", b."�������˲�λ", b."��������״��",
       b."�������˲�λ��С", b."��������Ѫ�״�С", b."��������������", b."����������������", b."������������������״",
       b."��������������ҩ��", b."���������Ȼ���", b."������������̥", b."��������������",b."ĸӤ��Ӵ�����˱ʱ��", b."����Ѫѹ", b."������Ѫ", b."��ʱ��ҩ", b."������ҩ",
       b."�������", b."�������", d."����3",d."������ʱ��", d."������", d."������", d."��¼��"
From Sol_Inf_Delivery A,
     Json_Table(Nvl(a.Newborndetail, '{"����1":"1"}'),
                 '$'
                  Columns(����1 Varchar2(50) Path '$.����1', ���̿�ʼʱ�� Varchar2(19) Path '$.���̿�ʼʱ��',
                          ����ȫ��ʱ�� Varchar2(19) Path '$.����ȫ��ʱ��', ̥�����ʱ�� Varchar2(19) Path '$.̥�����ʱ��',
                          ̥�����ʱ�� Varchar2(19) Path '$.̥�����ʱ��', ��һ���� Varchar2(50) Path '$.��һ����',
                          �ڶ����� Varchar2(50) Path '$.�ڶ�����', �������� Varchar2(50) Path '$.��������',
                          �������������� Varchar2(50) Path '$.��������������', ���� Varchar2(50) Path '$.����',
                          ��Ĥ��ʽ Varchar2(50) Path '$.��Ĥ��ʽ', ��Ĥʱ�� Varchar2(19) Path '$.��Ĥʱ��', ��ˮ��״ Varchar2(50) Path '$.��ˮ��״',
                          ��ˮ�� Varchar2(50) Path '$.��ˮ��', ��ˮ��ɫ Varchar2(50) Path '$.��ˮ��ɫ',
                          ̥�������ʽ Varchar2(50) Path '$.̥�������ʽ', ̥�̰��뷽ʽ Varchar2(50) Path '$.̥�̰��뷽ʽ',
                          ̥�������� Varchar2(50) Path '$.̥��������', ̥��̥Ĥ���� Varchar2(50) Path '$.̥��̥Ĥ����',
                          ̥����� Varchar2(50) Path '$.̥�����', ̥����̬ Varchar2(50) Path '$.̥����̬', ̥�̴�С Varchar2(50) Path '$.̥�̴�С',
                          ̥������ Varchar2(50) Path '$.̥������', ������� Varchar2(50) Path '$.�������', ������� Varchar2(50) Path '$.�������',
                         �����ٽ� Varchar2(50) Path '$.�����ٽ�',��� Varchar2(50) Path '$.���', �ƾ����� Varchar2(50) Path '$.�ƾ�����',
                          �����ʽ Varchar2(50) Path '$.�����ʽ',
                          ���̥��λ Varchar2(50) Path '$.���̥��λ', ������С Varchar2(50) Path '$.������С',
                          ������λ Varchar2(50) Path '$.������λ', �������˳̶� Varchar2(50) Path '$.�������˳̶�',
                          ���������п� Varchar2(50) Path '$.���������п�', �������˷�� Varchar2(50) Path '$.�������˷��',
                          ������������ Varchar2(50) Path '$.������������', �������˳��� Varchar2(50) Path '$.�������˳���',
                          �������˲�λ Varchar2(50) Path '$.�������˲�λ', ��������״�� Varchar2(50) Path '$.��������״��',
                          �������˲�λ��С Varchar2(50) Path '$.�������˲�λ��С', ��������Ѫ�״�С Varchar2(50) Path '$.��������Ѫ�״�С',
                          �������������� Varchar2(50) Path '$.��������������', ���������������� Varchar2(50) Path '$.����������������',
                          ������������������״ Varchar2(50) Path '$.������������������״', ��������������ҩ�� Varchar2(50) Path '$.��������������ҩ��',
                          ���������Ȼ��� Varchar2(50) Path '$.���������Ȼ���', ������������̥ Varchar2(50) Path '$.������������̥',
                          �������������� Varchar2(50) Path '$.��������������',������� Varchar2(50) Path '$.�������', ĸӤ��Ӵ�����˱ʱ�� Varchar2(50) Path '$.ĸӤ��Ӵ�����˱ʱ��',
                          ����Ѫѹ Varchar2(50) Path '$.����Ѫѹ',
                          ������Ѫ Varchar2(50) Path '$.������Ѫ', ��ʱ��ҩ Varchar2(50) Path '$.��ʱ��ҩ', ������ҩ Varchar2(50) Path '$.������ҩ',
                          ������� Varchar2(50) Path '$.�������',������� Varchar2(50) Path '$.�������')) As B,
     Json_Table(Nvl(a.Deliveryinf, '{"����3":"1"}'),
                 '$' Columns(����3 Varchar2(1) Path '$.����3', ������ʱ�� Varchar2(50) Path '$.������ʱ��',
                          ������ Varchar2(50) Path '$.������', ������ Varchar2(50) Path '$.������', ��¼�� Varchar2(50) Path '$.��¼��')) As D;

create or replace view v_sol_inf_equipment as
Select a.mid, b."���м���ǰ",b."���м�����",b."���м�����",b."�������ǰ",b."���������",b."���������",b."ֹѪǯ��ǰ",b."ֹѪǯ����",b."ֹѪǯ����",b."������ǰ",b."��������",b."��������",b."��������ǰ",b."����������",b."����������",b."�������ǰ",b."����������",b."���������",b."ϴ�����ǰ",b."ϴ��������",b."ϴ�������",b."������ǰ",b."���������",b."��������",b."������ǰ",b."��������",b."��������",b."����ǯ��ǰ",b."����ǯ����",b."����ǯ����",b."������ǰ",b."��������",b."��������",b."�γײ�ǰ",b."�γ�����",b."�γײ���",b."����˹��ǰ",b."����˹����",b."����˹����",b."��ǰ��ǰ",b."��ǰ����",b."��ǰ����",b."ɴ����ǰ",b."ɴ������",b."ɴ������",b."��Բǯ��ǰ",b."��Բǯ����",b."��Բǯ����"
From SOL_INF_Equipment a,JSON_TABLE(a.Content,'$' columns(
���м���ǰ   Number(2) PATH '$.���м���ǰ',
���м�����   Number(2) PATH '$.���м�����',
���м�����   Number(2) PATH '$.���м�����',
�������ǰ   Number(2) PATH '$.�������ǰ',
���������   Number(2) PATH '$.���������',
���������   Number(2) PATH '$.���������',
ֹѪǯ��ǰ   Number(2) PATH '$.ֹѪǯ��ǰ',
ֹѪǯ����   Number(2) PATH '$.ֹѪǯ����',
ֹѪǯ����   Number(2) PATH '$.ֹѪǯ����',
������ǰ   Number(2) PATH '$.������ǰ',
��������   Number(2) PATH '$.��������',
��������   Number(2) PATH '$.��������',
��������ǰ   Number(2) PATH '$.��������ǰ',
����������   Number(2) PATH '$.����������',
����������   Number(2) PATH '$.����������',
�������ǰ   Number(2) PATH '$.�������ǰ',
����������   Number(2) PATH '$.����������',
���������   Number(2) PATH '$.���������',
ϴ�����ǰ   Number(2) PATH '$.ϴ�����ǰ',
ϴ��������   Number(2) PATH '$.ϴ��������',
ϴ�������   Number(2) PATH '$.ϴ�������',
������ǰ   Number(2) PATH '$.������ǰ',
���������   Number(2) PATH '$.���������',
��������   Number(2) PATH '$.��������',
������ǰ   Number(2) PATH '$.������ǰ',
��������   Number(2) PATH '$.��������',
��������   Number(2) PATH '$.��������',
����ǯ��ǰ   Number(2) PATH '$.����ǯ��ǰ',
����ǯ����   Number(2) PATH '$.����ǯ����',
����ǯ����   Number(2) PATH '$.����ǯ����',
������ǰ   Number(2) PATH '$.������ǰ',
��������   Number(2) PATH '$.��������',
��������   Number(2) PATH '$.��������',
�γײ�ǰ   Number(2) PATH '$.�γײ�ǰ',
�γ�����   Number(2) PATH '$.�γ�����',
�γײ���   Number(2) PATH '$.�γײ���',
����˹��ǰ   Number(2) PATH '$.����˹��ǰ',
����˹����   Number(2) PATH '$.����˹����',
����˹����   Number(2) PATH '$.����˹����',
��ǰ��ǰ   Number(2) PATH '$.��ǰ��ǰ',
��ǰ����   Number(2) PATH '$.��ǰ����',
��ǰ����   Number(2) PATH '$.��ǰ����',
ɴ����ǰ   Number(2) PATH '$.ɴ����ǰ',
ɴ������   Number(2) PATH '$.ɴ������',
ɴ������   Number(2) PATH '$.ɴ������',
��Բǯ��ǰ   Number(2) PATH '$.��Բǯ��ǰ',
��Բǯ����   Number(2) PATH '$.��Բǯ����',
��Բǯ����   Number(2) PATH '$.��Բǯ����'
)) as b;

--����������
create or replace view v_newborn as
Select a.Mid, b.Bid, b.Babyno, b.Addtime, b.Recorder, a.Name, a.Old, a.Bedno, a.Pno, b.Sex, d."����2",d."��",d."����",d."ͷΧ",d."��Χ",d."һ�������Ӧ",d."һ�������ɫ",d."һ�����Ƥ��",d."һ������ë",d."ͷ������",d."­���ص�",d."̥ͷˮ��Ѫ��",d."̥ͷˮ�״�С",d."ǰض",d."����",d."����",d."��ǻ",d."��",d."���",d."��",d."Ƣ",d."��֫",d."��չ����",d."����",d."��ֳ��", e."����3",e."����1����",e."����5����",e."����10����",e."����1����",e."����5����",e."����10����",e."����1����",e."����5����",e."����10����",e."������1����",e."������5����",e."������10����",e."��ɫ1����",e."��ɫ5����",e."��ɫ10����",e."�ܷ�1����",e."�ܷ�5����",e."�ܷ�10����", f."����4",f."�����ڲ�ʱ�ϲ�֢����ҩ���",f."����ǰ̥�����",f."Ӥ������ʱ�������",f."����ȱ��",f."ĸ��ι��ָ��",f."���"
From Sol_Inf_Puerpera A, Sol_Inf_Newborns B,
     Json_Table(Nvl(b.Newborninf, '{"����2":"2"}'),
                 '$' Columns(����2 Varchar2(50) Path '$.����2', �� Varchar2(50) Path '$.��', ���� Varchar2(50) Path '$.����',
                          ͷΧ Varchar2(50) Path '$.ͷΧ', ��Χ Varchar2(50) Path '$.��Χ', һ�������Ӧ Varchar2(50) Path '$.һ�������Ӧ',
                          һ�������ɫ Varchar2(50) Path '$.һ�������ɫ', һ�����Ƥ�� Varchar2(50) Path '$.һ�����Ƥ��',
                          һ������ë Varchar2(50) Path '$.һ������ë', ͷ������ Varchar2(50) Path '$.ͷ������',
                          ­���ص� Varchar2(50) Path '$.­���ص�', ̥ͷˮ��Ѫ�� Varchar2(50) Path '$.̥ͷˮ��Ѫ��',
                          ̥ͷˮ�״�С Varchar2(50) Path '$.̥ͷˮ�״�С', ǰض Varchar2(50) Path '$.ǰض', ���� Varchar2(50) Path '$.����',
                          ���� Varchar2(50) Path '$.����', ��ǻ Varchar2(50) Path '$.��ǻ', �� Varchar2(50) Path '$.��',
                          ��� Varchar2(50) Path '$.���', �� Varchar2(50) Path '$.��', Ƣ Varchar2(50) Path '$.Ƣ',
                          ��֫ Varchar2(50) Path '$.��֫', ��չ���� Varchar2(50) Path '$.��չ����', ���� Varchar2(50) Path '$.����',
                          ��ֳ�� Varchar2(50) Path '$.��ֳ��')) As D,
     Json_Table(Nvl(b.Newbornscore, '{"����3":"3"}'),
                 '$' Columns(����3 Varchar2(50) Path '$.����3', ����1���� Varchar2(50) Path '$.����1����',
                          ����5���� Varchar2(50) Path '$.����5����', ����10���� Varchar2(50) Path '$.����10����',
                          ����1���� Varchar2(50) Path '$.����1����', ����5���� Varchar2(50) Path '$.����5����',
                          ����10���� Varchar2(50) Path '$.����10����', ����1���� Varchar2(50) Path '$.����1����',
                          ����5���� Varchar2(50) Path '$.����5����', ����10���� Varchar2(50) Path '$.����10����',
                          ������1���� Varchar2(50) Path '$.������1����', ������5���� Varchar2(50) Path '$.������5����',
                          ������10���� Varchar2(50) Path '$.������10����', ��ɫ1���� Varchar2(50) Path '$.��ɫ1����',
                          ��ɫ5���� Varchar2(50) Path '$.��ɫ5����', ��ɫ10���� Varchar2(50) Path '$.��ɫ10����',
                          �ܷ�1���� Varchar2(50) Path '$.�ܷ�1����', �ܷ�5���� Varchar2(50) Path '$.�ܷ�5����',
                          �ܷ�10���� Varchar2(50) Path '$.�ܷ�10����')) As E,
     Json_Table(Nvl(b.Otherinf, '{"����4":"4"}'),
                 '$' Columns(����4 Varchar2(50) Path '$.����4', �����ڲ�ʱ�ϲ�֢����ҩ��� Varchar2(50) Path '$.�����ڲ�ʱ�ϲ�֢����ҩ���  ',
                          ����ǰ̥����� Varchar2(50) Path '$.����ǰ̥�����  ', Ӥ������ʱ������� Varchar2(50) Path '$.Ӥ������ʱ�������  ',
                          ����ȱ�� Varchar2(50) Path '$.����ȱ��  ', ĸ��ι��ָ�� Varchar2(50) Path '$.ĸ��ι��ָ��  ',
                          ��� Varchar2(50) Path '$.���  ')) As F
Where a.Mid = b.Mid
Order By Addtime;

create or replace view v_sol_inf_newborns as
Select a.Mid, b.Bid, b.Babyno, b.Addtime, b.Recorder, a.Name, a.Old, a.Bedno, a.Pno, b.Sex, d."����2",d."��",d."����",d."ͷΧ",d."��Χ",d."һ�������Ӧ",d."һ�������ɫ",d."һ�����Ƥ��",d."һ������ë",d."ͷ������",d."­���ص�",d."̥ͷˮ��Ѫ��",d."̥ͷˮ�״�С",d."ǰض",d."����",d."����",d."��ǻ",d."��",d."���",d."��",d."Ƣ",d."��֫",d."��չ����",d."����",d."��ֳ��", e."����3",e."����1����",e."����5����",e."����10����",e."����1����",e."����5����",e."����10����",e."����1����",e."����5����",e."����10����",e."������1����",e."������5����",e."������10����",e."��ɫ1����",e."��ɫ5����",e."��ɫ10����",e."�ܷ�1����",e."�ܷ�5����",e."�ܷ�10����", f."����4",f."�����ڲ�ʱ�ϲ�֢����ҩ���",f."����ǰ̥�����",f."Ӥ������ʱ�������",f."����ȱ��",f."ĸ��ι��ָ��",f."���"
From Sol_Inf_Puerpera A, Sol_Inf_Newborns B,
     Json_Table(Nvl(b.Newborninf, '{"����2":"2"}'),
                 '$' Columns(����2 Varchar2(50) Path '$.����2', �� Varchar2(50) Path '$.��', ���� Varchar2(50) Path '$.����',
                          ͷΧ Varchar2(50) Path '$.ͷΧ', ��Χ Varchar2(50) Path '$.��Χ', һ�������Ӧ Varchar2(50) Path '$.һ�������Ӧ',
                          һ�������ɫ Varchar2(50) Path '$.һ�������ɫ', һ�����Ƥ�� Varchar2(50) Path '$.һ�����Ƥ��',
                          һ������ë Varchar2(50) Path '$.һ������ë', ͷ������ Varchar2(50) Path '$.ͷ������',
                          ­���ص� Varchar2(50) Path '$.­���ص�', ̥ͷˮ��Ѫ�� Varchar2(50) Path '$.̥ͷˮ��Ѫ��',
                          ̥ͷˮ�״�С Varchar2(50) Path '$.̥ͷˮ�״�С', ǰض Varchar2(50) Path '$.ǰض', ���� Varchar2(50) Path '$.����',
                          ���� Varchar2(50) Path '$.����', ��ǻ Varchar2(50) Path '$.��ǻ', �� Varchar2(50) Path '$.��',
                          ��� Varchar2(50) Path '$.���', �� Varchar2(50) Path '$.��', Ƣ Varchar2(50) Path '$.Ƣ',
                          ��֫ Varchar2(50) Path '$.��֫', ��չ���� Varchar2(50) Path '$.��չ����', ���� Varchar2(50) Path '$.����',
                          ��ֳ�� Varchar2(50) Path '$.��ֳ��')) As D,
     Json_Table(Nvl(b.Newbornscore, '{"����3":"3"}'),
                 '$' Columns(����3 Varchar2(50) Path '$.����3', ����1���� Varchar2(50) Path '$.����1����',
                          ����5���� Varchar2(50) Path '$.����5����', ����10���� Varchar2(50) Path '$.����10����',
                          ����1���� Varchar2(50) Path '$.����1����', ����5���� Varchar2(50) Path '$.����5����',
                          ����10���� Varchar2(50) Path '$.����10����', ����1���� Varchar2(50) Path '$.����1����',
                          ����5���� Varchar2(50) Path '$.����5����', ����10���� Varchar2(50) Path '$.����10����',
                          ������1���� Varchar2(50) Path '$.������1����', ������5���� Varchar2(50) Path '$.������5����',
                          ������10���� Varchar2(50) Path '$.������10����', ��ɫ1���� Varchar2(50) Path '$.��ɫ1����',
                          ��ɫ5���� Varchar2(50) Path '$.��ɫ5����', ��ɫ10���� Varchar2(50) Path '$.��ɫ10����',
                          �ܷ�1���� Varchar2(50) Path '$.�ܷ�1����', �ܷ�5���� Varchar2(50) Path '$.�ܷ�5����',
                          �ܷ�10���� Varchar2(50) Path '$.�ܷ�10����')) As E,
     Json_Table(Nvl(b.Otherinf, '{"����4":"4"}'),
                 '$' Columns(����4 Varchar2(50) Path '$.����4', �����ڲ�ʱ�ϲ�֢����ҩ��� Varchar2(50) Path '$.�����ڲ�ʱ�ϲ�֢����ҩ���  ',
                          ����ǰ̥����� Varchar2(50) Path '$.����ǰ̥�����  ', Ӥ������ʱ������� Varchar2(50) Path '$.Ӥ������ʱ�������  ',
                          ����ȱ�� Varchar2(50) Path '$.����ȱ��  ', ĸ��ι��ָ�� Varchar2(50) Path '$.ĸ��ι��ָ��  ',
                          ��� Varchar2(50) Path '$.���  ')) As F
Where a.Mid = b.Mid
Order By Addtime;

create or replace view v_sol_inf_puerpera as
Select Name, Mid, Old, LPad(Bedno, 10) Bedno, Pno, Diagnosis, Status, Decode(Expectant, 1, '��', '') ����,
       Decode(Checkinroom, 1, '��', '') �뷿, Decode(Birth, 1, '��', '') �ٲ�, Decode(Druglabor, 1, '��', '') ����,
       Decode(Delivery, 1, '��', '') ����, Decode(Newborns, 1, '��', '') ������, Decode(Postpartum, 1, '��', '') ����,
       Decode(Checkoutroom, 1, '��', '') ����,Decode(Equipment, 1, '��', '') ��е,outtime,pid,tid
From Sol_Inf_Puerpera;

create or replace view v_sol_rs_birth as
Select a.mid,b."�Ѵ�",b."����",b."Ѫ��",b."��������ʷ",b."ĩ���¾�",b."Ԥ����",b."��ǰ�ϼ��侶",b."���ռ侶",b."���ǽ�ڼ侶",b."�����⾶",b."���ǻ���",b."���ǹؽ�",b."�����м�",b."������",b."����֢",b."��ǰ��¼����",b."���ʱ��",b."Ѫѹ",b."����",b."����",b."̥����",b."̥����С",b."����������",b."̥λ",b."�ν�",b."��Ĥ���",b."��¶",b."����",b."�����",b."������ʼʱ��",b."��Ĥʱ��",b."��Ժ����"
From SOL_RS_BIRTH a,JSON_TABLE(Nvl(a.CONTENT,'{����:1}'),'$' columns(
�Ѵ�            Number(3)    PATH '$.�Ѵ�',
����            Number(3)    PATH '$.����',
Ѫ��            Varchar2(10) PATH '$.Ѫ��',
��������ʷ      Varchar2(50) PATH '$.��������ʷ',
ĩ���¾�        Varchar2(20) PATH '$.ĩ���¾�',
Ԥ����          Varchar2(20) PATH '$.Ԥ����',
��ǰ�ϼ��侶    Number(5) PATH '$.��ǰ�ϼ��侶',
���ռ侶        Number(5) PATH '$.���ռ侶',
���ǽ�ڼ侶    Number(5) PATH '$.���ǽ�ڼ侶',
�����⾶        Number(5) PATH '$.�����⾶',
���ǻ���        Varchar2(10) PATH '$.���ǻ���',
���ǹؽ�        Varchar2(10) PATH '$.���ǹؽ�',
�����м�        Varchar2(10) PATH '$.�����м�',
������          Varchar2(10) PATH '$.������',
����֢          Varchar2(100) PATH '$.����֢',
��ǰ��¼����    Varchar2(100) PATH '$.��ǰ��¼����',
���ʱ��        Varchar2(20) PATH '$.���ʱ��',
Ѫѹ            Varchar2(10) PATH '$.Ѫѹ',
����            Number(4,2)  PATH '$.����',
����            Varchar2(10) PATH '$.����',
̥����            Varchar2(10) PATH '$.̥����',
̥����С        Number(5,2) PATH '$.̥����С',
����������      Varchar2(10) PATH '$.����������',
̥λ            Varchar2(10) PATH '$.̥λ',
�ν�            Varchar2(10) PATH '$.�ν�',
��Ĥ���        Varchar2(10) PATH '$.��Ĥ���',
��¶            Varchar2(2) PATH '$.��¶',
����            Number(4,2) PATH '$.����',
�����          Varchar2(50) PATH '$.�����',
������ʼʱ��    Varchar2(20) PATH '$.������ʼʱ��',
��Ĥʱ��        Varchar2(20) PATH '$.��Ĥʱ��',
��Ժ����        Varchar2(100) PATH '$.��Ժ����'
)) as b;

create or replace view v_sol_rs_birth_course as
Select  a.courseid,a.mid,b."���ʱ��",b."̥��λ",b."Ѫѹ",b."����",b."����",b."̥����",b."����ǿ��",b."��������",b."�������",b."������",b."����",b."��Ĥ���",b."��¶",b."����",b."�����"
From SOL_RS_BIRTH_COURSE a,JSON_TABLE(a.CONTENT,'$' columns(
���ʱ��        Varchar2(20)  PATH '$.���ʱ��',
̥��λ Varchar2(20)  PATH '$.̥��λ',
Ѫѹ        Varchar2(10) PATH '$.Ѫѹ',
����        Number(4,2)  PATH '$.����',
����        Varchar2(10) PATH '$.����',
̥����        Varchar2(10) PATH '$.̥����',
����ǿ��    Varchar2(10) PATH '$.����ǿ��',
��������  Varchar2(10) PATH '$.��������',
�������  Varchar2(10) PATH '$.�������',
������    Varchar2(10) PATH '$.������',
����        Number(4,2) PATH '$.����',
��Ĥ���    Varchar2(10) PATH '$.��Ĥ���',
��¶        Number(2) PATH '$.��¶'��
����        Varchar2(50) PATH '$.����'��
�����      Varchar2(50) PATH '$.�����'
)) as b;

create or replace view v_sol_rs_druglabor as
Select Mid, To_Char(����, 'YYYY-MM-DD HH24:MI') ����, ����ָ��, �������� from Sol_Rs_Druglabor;

create or replace view v_sol_rs_druglabor_list as
Select a.Mid, a.Courseid ID, b."��¼ʱ��",b."Ѫѹ",b."����",b."̥����",b."����ǿ��",b."��������",b."�������",b."����",b."��¶",b."��ˮ��",b."��ˮ��״",b."����",b."��¼��",b."����",b."����"
From Sol_Rs_Druglabor_List a,
     Json_Table(a.Content,'$' Columns(
     ��¼ʱ�� Varchar2(20) Path '$.��¼ʱ��',
     ���� Number(3,1) Path '$.����',
     ���� Number(3) Path '$.����',
     Ѫѹ Varchar2(7) Path '$.Ѫѹ',
     ���� Number(3) Path '$.����',
     ̥���� Number(3) Path '$.̥����',
     ����ǿ�� Varchar2(10) Path '$.����ǿ��',
     �������� Number(3) Path '$.��������',
     ������� Number(2) Path '$.�������',
     ���� Number(3) Path '$.����',
     ��¶ Varchar2(10) Path '$.��¶',
     ��ˮ�� Number(4) Path '$.��ˮ��',
     ��ˮ��״ Varchar2(10) Path '$.��ˮ��״',
     ���� Varchar2(100) Path '$.����',
     ��¼�� Varchar2(100) Path '$.��¼��')) b;

create or replace view v_sol_rs_expectant as
Select a.mid,a.courseid,b."��¼ʱ��",b."̥��λ",b."Ѫѹ",b."����",b."��Χ",b."̥��������",b."̥��������",b."̥��������",b."̥����",b."��¶",b."����",b."��Ĥ���",b."��ˮ��״",b."����ǿ��",b."��������",b."�������",b."����",b."�����"
From SOL_RS_EXPECTANT a,JSON_TABLE(a.Content,'$' columns(
��¼ʱ��    Varchar2(50) PATH '$.��¼ʱ��',
̥��λ  Varchar2(20) PATH '$.̥��λ',
Ѫѹ     Varchar2(20) PATH '$.Ѫѹ',
����     Number(4,2) PATH '$.����',
��Χ     Varchar2(20) PATH '$.��Χ',
̥��������     Number(3) PATH '$.̥��������',
̥��������     Number(3) PATH '$.̥��������',
̥��������   Number(3) PATH '$.̥��������',
̥���� Number(3) PATH '$.̥����',
��¶     Varchar2(20) PATH '$.��¶',
����     Varchar2(20) PATH '$.����',
��Ĥ���     Varchar2(20) PATH '$.��Ĥ���',
��ˮ��״      Varchar2(20) PATH '$.��ˮ��״',
����ǿ��     Varchar2(20) PATH '$.����ǿ��',
��������       Varchar2(20) PATH '$.��������',
�������       Varchar2(20) PATH '$.�������',
����     Varchar2(100) PATH '$.����',
�����       Varchar2(20) PATH '$.�����'
)) as b;

create or replace view v_sol_rs_postpartum as
Select a.Mid, ��������, �����ʱ��, ���䷽ʽ, ������ʱ��, ������ʱbp, ������ʱ����, ������ʱ��������, ������ʱ������Ѫ, ������ʱһ�����, ����,  ����
From Sol_Rs_Postpartum A,
     Json_Table(a.Content,
                 '$' Columns(�������� varchar2(20) Path '$.��������', �����ʱ�� varchar2(20) Path '$.�����ʱ��', ���䷽ʽ Varchar2(20) Path '$.���䷽ʽ',
                          ������ʱ�� varchar2(20) Path '$.������ʱ��', ������ʱbp varchar2(7) Path '$.������ʱBP', ������ʱ���� Number(3) Path '$.������ʱ����',
                          ������ʱ�������� Number(2) Path '$.������ʱ��������', ������ʱ������Ѫ Number(3) Path '$.������ʱ������Ѫ',
                          ������ʱһ����� Varchar2(10) Path '$.������ʱһ�����', ���� Varchar2(20) Path '$.����', ���� Varchar2(10) Path '$.����'));


create or replace view v_sol_rs_postpartum_list as
Select a.Mid, a.Courseid ID, ��¼ʱ��, ����, �鷿����, ��ͷ, �ӹ�����, �ӹ�ѹʹ, ��¶��, ��¶��ɫ, ��¶��ζ, ��������, ��������, ��������, С��, ���, �������, ǩ��
From Sol_Rs_Postpartum_List A,
     Json_Table(a.Content,
                 '$' Columns(��¼ʱ�� Varchar2(20) Path '$.��¼ʱ��', ���� Number(4) Path '$.����', �鷿���� Varchar2(10) Path '$.�鷿����',
                          ��ͷ Varchar2(50) Path '$.��ͷ', �ӹ����� Number(3) Path '$.�ӹ�����', �ӹ�ѹʹ Varchar2(50) Path '$.�ӹ�ѹʹ',
                          ��¶�� Number(4) Path '$.��¶��', ��¶��ɫ Varchar2(20) Path '$.��¶��ɫ', ��¶��ζ Varchar2(20) Path '$.��¶��ζ',
                          �������� Varchar2(10) Path '$.��������', �������� Varchar2(10) Path '$.��������', �������� Varchar2(50) Path '$.��������',
                          С�� Varchar2(50) Path '$.С��', ��� Varchar2(50) Path '$.���', ������� Varchar2(100) Path '$.�������',
                          ǩ�� Varchar2(100) Path '$.ǩ��'));




