prompt PL/SQL Developer import file
prompt Created on 2012��3��28�� by HJ
set feedback off
set define off
prompt Loading HL7��������...
insert into HL7�������� (ID, IP��ַ, �˿ں�, ��������, ���ͳ�������, �����豸����, ���ճ�������, �����豸����)
values (1, '127.0.0.1', '104', 1, 'MUSE ECG Result 1', 'MEI MUSE', 'ZLHIS', 'HIS001');
insert into HL7�������� (ID, IP��ַ, �˿ں�, ��������, ���ͳ�������, �����豸����, ���ճ�������, �����豸����)
values (2, '127.0.0.1', '1024', 2, 'ZLHIS', 'HIS001', 'MUSE ECG Result 1', 'MEI MUSE');
commit;
prompt 2 records loaded
prompt Loading HL7��Ϣ����...
insert into HL7��Ϣ���� (ID, ����ID, ��������, ��Ϣ����, ��Ϣ����, ��Ϣ�����)
values (1, 1, '�����ĵ���', 'ORU_R01', null, 'MSH|PID|[PV1]|{OBR|[{DG1}]|[{NTE}]}|ZEX|ZPH|[{OBX}]');
insert into HL7��Ϣ���� (ID, ����ID, ��������, ��Ϣ����, ��Ϣ����, ��Ϣ�����)
values (2, 2, '������ҽ��', 'ORM_O01', 'NW', 'MSH|PID|PV1|ORC|OBR');
insert into HL7��Ϣ���� (ID, ����ID, ��������, ��Ϣ����, ��Ϣ����, ��Ϣ�����)
values (3, 2, '����ȡ��ҽ��', 'ORM_O01', 'CA', 'MSH|PID|PV1|ORC|OBR');
insert into HL7��Ϣ���� (ID, ����ID, ��������, ��Ϣ����, ��Ϣ����, ��Ϣ�����)
values (4, 2, '����ɾ��ҽ��', 'ORM_O01', 'DD', 'MSH|PID|PV1|ORC|OBR');
commit;
prompt 4 records loaded
prompt Loading HL7��Ϣ������...
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (268, 4, 'MSH', 1, 'ST', null, '^~\&', 'Encoding Characters �����ַ�');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (2, 2, 'MSH', 1, 'ST', null, '^~\&', 'Encoding Characters �����ַ�');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (3, 2, 'MSH', 2, 'HD', 'MUSE ECG Result 1', 'ZLHIS', 'Sending Application ���ͳ���');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (4, 2, 'MSH', 3, 'HD', 'MEI MUSE', 'HIS001', 'Sending Facility �����豸');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (5, 2, 'MSH', 4, 'HD', 'ZLHIS', 'MUSE ECG Result 1', 'Receiving Application ���ճ���');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (6, 2, 'MSH', 5, 'HD', 'HIS001', 'MEI MUSE', 'Receiving Facility �����豸');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (7, 2, 'MSH', 6, 'TS', null, '[��ǰʱ��]', 'Date/Time Of Message ��Ϣ������/ʱ��');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (8, 2, 'MSH', 7, 'ST', null, null, 'Security ��ȫ��');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (9, 2, 'MSH', 8, 'CM', 'ORM^O01', 'ORM^O01', 'Message Type ��Ϣ����');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (10, 2, 'MSH', 9, 'ST', null, '[��ǰʱ��]', 'Message Control ID ��Ϣ����ID');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (11, 2, 'MSH', 10, 'PT', 'P', 'P', 'Processing ID ����ID');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (12, 2, 'MSH', 11, 'VID', '2.4', '2.4', 'Version ID �汾ID');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (135, 3, 'MSH', 1, 'ST', null, '^~\&', 'Encoding Characters �����ַ�');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (136, 3, 'MSH', 2, 'HD', 'MUSE ECG Result 1', 'ZLHIS', 'Sending Application ���ͳ���');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (137, 3, 'MSH', 3, 'HD', 'MEI MUSE', 'HIS001', 'Sending Facility �����豸');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (138, 3, 'MSH', 4, 'HD', 'ZLHIS', 'MUSE ECG Result 1', 'Receiving Application ���ճ���');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (139, 3, 'MSH', 5, 'HD', 'HIS001', 'MEI MUSE', 'Receiving Facility �����豸');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (140, 3, 'MSH', 6, 'TS', null, '[��ǰʱ��]', 'Date/Time Of Message ��Ϣ������/ʱ��');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (141, 3, 'MSH', 7, 'ST', null, null, 'Security ��ȫ��');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (142, 3, 'MSH', 8, 'CM', 'ORM^O01', 'ORM^O01', 'Message Type ��Ϣ����');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (143, 3, 'MSH', 9, 'ST', null, '[��ǰʱ��]', 'Message Control ID ��Ϣ����ID');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (144, 3, 'MSH', 10, 'PT', 'P', 'P', 'Processing ID ����ID');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (145, 3, 'MSH', 11, 'VID', '2.4', '2.4', 'Version ID �汾ID');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (146, 3, 'PID', 1, 'SI', null, '1', 'Set ID - Patient ID ����ID - ����ID');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (147, 3, 'PID', 2, 'CX', null, null, 'Patient ID ����ID');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (148, 3, 'PID', 3, 'CX', null, '[��ʶ��] ', 'Patient Identifier List ���߱�ʶ����');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (149, 3, 'PID', 4, 'CX', null, '[����ID]', 'Alternate Patient ID ���û���ID');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (150, 3, 'PID', 5, 'XPN', null, '[����]', 'Patient Name ��������');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (151, 3, 'PID', 6, 'XPN', null, null, 'Mother''s Maiden Name ĸ�׵Ļ�ǰ��');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (152, 3, 'PID', 7, 'TS', null, '[��������]', 'Date/Time of Birth ��������');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (153, 3, 'PID', 8, 'IS', null, '[�Ա�]', 'Sex �Ա�');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (154, 3, 'PID', 9, 'XPN', null, '[����]', 'Patient Alias ���߱���');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (155, 3, 'PID', 10, 'CE', null, 'C', 'Race ����');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (156, 3, 'PID', 11, 'XAD', null, '[��ϵ�˵�ַ]', 'Patient Address ����סַ');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (157, 3, 'PID', 12, 'IS', null, null, 'County Code �ش���');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (158, 3, 'PID', 13, 'XTN', null, '[��ͥ�绰]', 'Phone Number - Home ��ͥ�绰');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (159, 3, 'PID', 14, 'XTN', null, '[��ϵ�˵绰]', 'Phone Number - Business ��λ�绰');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (160, 3, 'PID', 15, 'CE', null, null, 'Primary Language ĸ��');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (161, 3, 'PID', 16, 'IS', null, '[����״��]', 'Marital Status ����״��');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (162, 3, 'PID', 17, 'CE', null, null, 'Religion �ڽ�����');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (163, 3, 'PID', 18, 'CX', null, '[���֤��]', 'Patient Account Number�����˺�');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (164, 3, 'PID', 19, 'ST', null, null, 'SSN Number - Patient ������ᱣ�պ���');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (165, 3, 'PV1', 1, 'SI', null, '1', 'Set ID-PV1����ID-PV1');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (166, 3, 'PV1', 2, 'IS', null, '[������Դ]', 'Patient Class �������');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (167, 3, 'PV1', 3, 'PL', null, '[��ǰ��������]^[��������]^[����]^^^^^^[��ǰ��������]', 'Assigned Patient Location ָ������λ��');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (168, 3, 'PV1', 4, 'IS', null, null, 'Admission Type ��Ժ����');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (169, 3, 'PV1', 5, 'CX', null, null, 'Preadmit Number Ԥ����Ժ����');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (170, 3, 'PV1', 6, 'PL', null, null, 'Prior Patient Location ����ԭλ��');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (171, 3, 'PV1', 7, 'XCN', null, '[����ҽ��]', 'Attending Doctor ����ҽ��');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (172, 3, 'PV1', 8, 'XCN', null, null, 'Referring Doctor ת��ҽ��');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (173, 3, 'PV1', 9, 'XCN', null, null, 'Consulting Doctor ����ҽ��');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (174, 3, 'PV1', 10, 'IS', null, null, 'Hospital Service ҽԺ����');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (175, 3, 'PV1', 11, 'PL', null, null, 'Temporary Location ��ʱλ��');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (176, 3, 'PV1', 12, 'IS', null, null, 'Preadmit Test Indicator Ԥ����Ժ�����ʶ��');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (177, 3, 'PV1', 13, 'IS', null, null, 'Readmission Indicator �ٴ���Ժ��ʶ��');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (178, 3, 'PV1', 14, 'IS', null, null, 'Admit Source ��Ժ��Դ');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (179, 3, 'PV1', 15, 'IS', null, null, 'Ambulatory Status ���������߶�״̬');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (180, 3, 'PV1', 16, 'IS', null, null, 'VIP Indicator VIP��ʶ��');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (181, 3, 'PV1', 17, 'XCN', null, null, 'Admitting Doctor ��Ժҽ��');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (182, 3, 'PV1', 18, 'IS', null, null, 'Patient Type ��������');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (183, 3, 'PV1', 19, 'CX', null, '1', 'Visit Number �������');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (184, 3, 'PV1', 20, 'FC', null, null, 'Financial Class ����״�����');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (185, 3, 'PV1', 21, 'IS', null, null, 'Charge Price Indicator ���ü۸��ʶ��');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (186, 3, 'PV1', 22, 'IS', null, null, 'Courtesy Code ��ò����');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (187, 3, 'PV1', 23, 'IS', null, null, 'Credit Rating ���õȼ�');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (188, 3, 'PV1', 24, 'IS', null, null, 'Contract Code ��ͬ����');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (189, 3, 'PV1', 25, 'DT', null, null, 'Contract Effective Date ��ͬ��Ч����');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (190, 3, 'PV1', 26, 'NM', null, null, 'Contract Amount ��ͬ����');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (191, 3, 'PV1', 27, 'NM', null, null, 'Contract Period ��ͬ����');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (192, 3, 'PV1', 28, 'IS', null, null, 'Interest Code ���ʴ���');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (193, 3, 'PV1', 29, 'IS', null, null, 'Transfer to Bad Debt Code תΪ���˴���');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (194, 3, 'PV1', 30, 'DT', null, null, 'Transfer to Bad Debt Date תΪ��������');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (195, 3, 'PV1', 31, 'IS', null, null, 'Bad Debt Agency Code ���˴������');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (196, 3, 'PV1', 32, 'NM', null, null, 'Bad Debt Transfer Amount תΪ�����ܶ�');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (197, 3, 'PV1', 33, 'NM', null, null, 'Bad Debt Recovery Amount ���˻ָ��ܶ�');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (198, 3, 'PV1', 34, 'IS', null, null, 'Delete Account Indicator ɾ���˻���ʶ��');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (199, 3, 'PV1', 35, 'DT', null, null, 'Delete Account Date ɾ���˻�����');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (200, 3, 'PV1', 36, 'IS', null, 'UNK', 'Discharge Disposition ��Ժ����');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (201, 3, 'PV1', 37, 'CM', null, null, 'Discharge to Location ��Ժȥ����λ��');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (202, 3, 'PV1', 38, 'CE', null, null, 'Diet Type ��ʳ����');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (203, 3, 'PV1', 39, 'IS', null, null, 'Servicing Facility �������');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (204, 3, 'PV1', 40, 'IS', null, null, 'Bed Status ��λ״��');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (205, 3, 'PV1', 41, 'IS', null, null, 'Account Status �˻�״��');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (206, 3, 'PV1', 42, 'IS', null, null, 'Pending Location ����λ��');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (207, 3, 'PV1', 43, 'PL', null, null, 'Prior Temporary Location ǰ��ʱλ��');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (208, 3, 'PV1', 44, 'TS', null, '[��Ժʱ��]', 'Admit Date/Time ��Ժ����/ʱ��');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (209, 3, 'PV1', 45, 'TS', null, '[��Ժʱ��]', 'Discharge Date/Time ��Ժ����/ʱ��');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (210, 3, 'PV1', 46, 'NM', null, null, 'Current Patient Balance ��ǰ���߲��');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (211, 3, 'PV1', 47, 'NM', null, null, 'Total Charges �ܷ���');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (212, 3, 'PV1', 48, 'NM', null, null, 'Total Adjustments �ܵ�����');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (213, 3, 'PV1', 49, 'NM', null, null, 'Total Payments ��֧����');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (214, 3, 'PV1', 50, 'CX', null, null, 'Alternate Visit ID ���þ���ID');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (215, 3, 'ORC', 1, 'ID', null, '[��Ϣ����]', 'ҽ��������');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (216, 3, 'ORC', 2, 'EI', null, '[ҽ��ID]', '������ҽ������');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (217, 3, 'ORC', 3, 'EI', null, null, 'ִ����ҽ������');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (218, 3, 'ORC', 4, 'EI', null, null, '������ҽ�������');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (219, 3, 'ORC', 5, 'ID', null, null, 'ҽ��״̬');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (220, 3, 'ORC', 6, 'ID', null, null, 'Ӧ����');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (221, 3, 'ORC', 7, 'TQ', null, null, '����/ʱ��');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (222, 3, 'ORC', 8, 'CM', null, null, '�����ʶ');
commit;
prompt 100 records committed...
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (223, 3, 'ORC', 9, 'TS', null, '[����ʱ��]', '��������/ʱ��');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (224, 3, 'ORC', 10, 'XCN', null, '[����ҽ��]', '¼����');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (225, 3, 'ORC', 11, 'XCN', null, '[У�Ի�ʿ]', 'У����');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (226, 3, 'ORC', 12, 'XCN', null, '[����ҽ��]', 'ҽ���ṩ��');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (231, 3, 'OBR', 1, 'SI', null, '1', '����ID-OBR');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (232, 3, 'OBR', 2, 'EI', null, null, '������ҽ������');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (233, 3, 'OBR', 3, 'EI', null, null, 'ִ����ҽ������');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (234, 3, 'OBR', 4, 'CE', null, '[ҽ������]', 'ͨ�÷����ʶ');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (235, 3, 'OBR', 5, 'ID', null, null, '���ȼ�-OBR');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (236, 3, 'OBR', 6, 'TS', null, '[����ʱ��]', '��������/ʱ��');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (237, 3, 'OBR', 7, 'TS', null, '[����ʱ��]', '�۲�����/ʱ��#');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (238, 3, 'OBR', 8, 'TS', null, null, '�۲��������/ʱ��#');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (239, 3, 'OBR', 9, 'CQ', null, null, '�����ռ���*');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (240, 3, 'OBR', 10, 'XCN', null, null, '�ռ��߱�ʶ*');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (241, 3, 'OBR', 11, 'ID', null, null, '������Ϊ��*');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (242, 3, 'OBR', 12, 'CE', null, null, 'Σ�����ش���');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (243, 3, 'OBR', 13, 'ST', null, null, '����ٴ���Ϣ');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (244, 3, 'OBR', 14, 'TS', null, null, '�յ���������/ʱ��*');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (245, 3, 'OBR', 15, 'CM', null, null, '������Դ');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (246, 3, 'OBR', 16, 'XCN', null, '[����ҽ��]', 'ҽ���ṩ��');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (247, 3, 'OBR', 17, 'XTN', null, null, 'ҽ���طõ绰����');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (248, 3, 'OBR', 18, 'ST', null, null, '�������ֶ�1');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (249, 3, 'OBR', 19, 'ST', null, null, '�������ֶ�2');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (250, 3, 'OBR', 20, 'ST', null, null, 'ִ�����ֶ�1+');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (251, 3, 'OBR', 21, 'ST', null, null, 'ִ�����ֶ�2+');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (252, 3, 'OBR', 22, 'TS', null, null, '�������/״̬�ı�����/ʱ��+');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (253, 3, 'OBR', 23, 'CM', null, null, '�շ�+');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (254, 3, 'OBR', 24, 'ID', null, null, '��Ϸ����ʶID');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (255, 3, 'OBR', 25, 'ID', null, null, '���״̬+');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (256, 3, 'OBR', 26, 'CM', null, null, '������+');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (257, 3, 'OBR', 27, 'TQ', null, null, '����/ʱ��');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (258, 3, 'OBR', 28, 'XCN', null, null, '�����Ҫ��');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (259, 3, 'OBR', 29, 'CM', null, null, '����������');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (260, 3, 'OBR', 30, 'ID', null, null, '�����ж���ʽ');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (261, 3, 'OBR', 31, 'CE', null, null, '�о�ԭ��');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (262, 3, 'OBX', 1, 'SI', null, null, 'Set ID -OBX ����ID -OBX');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (263, 3, 'OBX', 2, 'ID', null, null, 'Value Type ֵ����');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (264, 3, 'OBX', 3, 'CE', null, null, 'Observation Identifier �۲��ʶ��');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (265, 3, 'OBX', 4, 'ST', null, null, 'Observation Sub -ID �۲���ID');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (266, 3, 'OBX', 5, 'TEXT', null, null, 'Observation Value �۲�ֵ');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (269, 4, 'MSH', 2, 'HD', 'MUSE ECG Result 1', 'ZLHIS', 'Sending Application ���ͳ���');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (13, 2, 'PID', 1, 'SI', null, '1', 'Set ID - Patient ID ����ID - ����ID');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (14, 2, 'PID', 2, 'CX', null, null, 'Patient ID ����ID');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (15, 2, 'PID', 3, 'CX', null, '[��ʶ��] ', 'Patient Identifier List ���߱�ʶ����');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (16, 2, 'PID', 4, 'CX', null, '[����ID]', 'Alternate Patient ID ���û���ID');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (17, 2, 'PID', 5, 'XPN', null, '[����]', 'Patient Name ��������');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (18, 2, 'PID', 6, 'XPN', null, null, 'Mother''s Maiden Name ĸ�׵Ļ�ǰ��');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (19, 2, 'PID', 7, 'TS', null, '[��������]', 'Date/Time of Birth ��������');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (20, 2, 'PID', 8, 'IS', null, '[�Ա�]', 'Sex �Ա�');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (21, 2, 'PID', 9, 'XPN', null, '[����]', 'Patient Alias ���߱���');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (22, 2, 'PID', 10, 'CE', null, 'C', 'Race ����');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (23, 2, 'PID', 11, 'XAD', null, '[��ϵ�˵�ַ]', 'Patient Address ����סַ');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (24, 2, 'PID', 12, 'IS', null, null, 'County Code �ش���');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (25, 2, 'PID', 13, 'XTN', null, '[��ͥ�绰]', 'Phone Number - Home ��ͥ�绰');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (26, 2, 'PID', 14, 'XTN', null, '[��ϵ�˵绰]', 'Phone Number - Business ��λ�绰');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (27, 2, 'PID', 15, 'CE', null, null, 'Primary Language ĸ��');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (28, 2, 'PID', 16, 'IS', null, '[����״��]', 'Marital Status ����״��');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (29, 2, 'PID', 17, 'CE', null, null, 'Religion �ڽ�����');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (30, 2, 'PID', 18, 'CX', null, '[���֤��]', 'Patient Account Number�����˺�');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (31, 2, 'PID', 19, 'ST', null, null, 'SSN Number - Patient ������ᱣ�պ���');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (32, 2, 'PV1', 1, 'SI', null, '1', 'Set ID-PV1����ID-PV1');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (33, 2, 'PV1', 2, 'IS', null, '[������Դ]', 'Patient Class �������');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (34, 2, 'PV1', 3, 'PL', null, '[��ǰ��������]^[��������]^[����]^^^^^^[��ǰ��������]', 'Assigned Patient Location ָ������λ��');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (35, 2, 'PV1', 4, 'IS', null, null, 'Admission Type ��Ժ����');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (36, 2, 'PV1', 5, 'CX', null, null, 'Preadmit Number Ԥ����Ժ����');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (37, 2, 'PV1', 6, 'PL', null, null, 'Prior Patient Location ����ԭλ��');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (38, 2, 'PV1', 7, 'XCN', null, '[����ҽ��]', 'Attending Doctor ����ҽ��');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (39, 2, 'PV1', 8, 'XCN', null, null, 'Referring Doctor ת��ҽ��');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (40, 2, 'PV1', 9, 'XCN', null, null, 'Consulting Doctor ����ҽ��');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (41, 2, 'PV1', 10, 'IS', null, null, 'Hospital Service ҽԺ����');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (42, 2, 'PV1', 11, 'PL', null, null, 'Temporary Location ��ʱλ��');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (43, 2, 'PV1', 12, 'IS', null, null, 'Preadmit Test Indicator Ԥ����Ժ�����ʶ��');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (44, 2, 'PV1', 13, 'IS', null, null, 'Readmission Indicator �ٴ���Ժ��ʶ��');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (45, 2, 'PV1', 14, 'IS', null, null, 'Admit Source ��Ժ��Դ');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (46, 2, 'PV1', 15, 'IS', null, null, 'Ambulatory Status ���������߶�״̬');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (47, 2, 'PV1', 16, 'IS', null, null, 'VIP Indicator VIP��ʶ��');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (48, 2, 'PV1', 17, 'XCN', null, null, 'Admitting Doctor ��Ժҽ��');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (49, 2, 'PV1', 18, 'IS', null, null, 'Patient Type ��������');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (50, 2, 'PV1', 19, 'CX', null, '1', 'Visit Number �������');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (51, 2, 'PV1', 20, 'FC', null, null, 'Financial Class ����״�����');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (52, 2, 'PV1', 21, 'IS', null, null, 'Charge Price Indicator ���ü۸��ʶ��');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (53, 2, 'PV1', 22, 'IS', null, null, 'Courtesy Code ��ò����');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (54, 2, 'PV1', 23, 'IS', null, null, 'Credit Rating ���õȼ�');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (55, 2, 'PV1', 24, 'IS', null, null, 'Contract Code ��ͬ����');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (56, 2, 'PV1', 25, 'DT', null, null, 'Contract Effective Date ��ͬ��Ч����');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (57, 2, 'PV1', 26, 'NM', null, null, 'Contract Amount ��ͬ����');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (58, 2, 'PV1', 27, 'NM', null, null, 'Contract Period ��ͬ����');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (59, 2, 'PV1', 28, 'IS', null, null, 'Interest Code ���ʴ���');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (60, 2, 'PV1', 29, 'IS', null, null, 'Transfer to Bad Debt Code תΪ���˴���');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (61, 2, 'PV1', 30, 'DT', null, null, 'Transfer to Bad Debt Date תΪ��������');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (62, 2, 'PV1', 31, 'IS', null, null, 'Bad Debt Agency Code ���˴������');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (63, 2, 'PV1', 32, 'NM', null, null, 'Bad Debt Transfer Amount תΪ�����ܶ�');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (64, 2, 'PV1', 33, 'NM', null, null, 'Bad Debt Recovery Amount ���˻ָ��ܶ�');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (65, 2, 'PV1', 34, 'IS', null, null, 'Delete Account Indicator ɾ���˻���ʶ��');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (66, 2, 'PV1', 35, 'DT', null, null, 'Delete Account Date ɾ���˻�����');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (67, 2, 'PV1', 36, 'IS', null, 'UNK', 'Discharge Disposition ��Ժ����');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (68, 2, 'PV1', 37, 'CM', null, null, 'Discharge to Location ��Ժȥ����λ��');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (69, 2, 'PV1', 38, 'CE', null, null, 'Diet Type ��ʳ����');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (70, 2, 'PV1', 39, 'IS', null, null, 'Servicing Facility �������');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (71, 2, 'PV1', 40, 'IS', null, null, 'Bed Status ��λ״��');
commit;
prompt 200 records committed...
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (72, 2, 'PV1', 41, 'IS', null, null, 'Account Status �˻�״��');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (73, 2, 'PV1', 42, 'IS', null, null, 'Pending Location ����λ��');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (74, 2, 'PV1', 43, 'PL', null, null, 'Prior Temporary Location ǰ��ʱλ��');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (75, 2, 'PV1', 44, 'TS', null, '[��Ժʱ��]', 'Admit Date/Time ��Ժ����/ʱ��');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (76, 2, 'PV1', 45, 'TS', null, '[��Ժʱ��]', 'Discharge Date/Time ��Ժ����/ʱ��');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (77, 2, 'PV1', 46, 'NM', null, null, 'Current Patient Balance ��ǰ���߲��');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (78, 2, 'PV1', 47, 'NM', null, null, 'Total Charges �ܷ���');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (79, 2, 'PV1', 48, 'NM', null, null, 'Total Adjustments �ܵ�����');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (80, 2, 'PV1', 49, 'NM', null, null, 'Total Payments ��֧����');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (81, 2, 'PV1', 50, 'CX', null, null, 'Alternate Visit ID ���þ���ID');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (82, 2, 'ORC', 1, 'ID', null, '[��Ϣ����]', 'ҽ��������');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (83, 2, 'ORC', 2, 'EI', null, '[ҽ��ID]', '������ҽ������');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (84, 2, 'ORC', 3, 'EI', null, null, 'ִ����ҽ������');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (85, 2, 'ORC', 4, 'EI', null, null, '������ҽ�������');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (86, 2, 'ORC', 5, 'ID', null, null, 'ҽ��״̬');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (87, 2, 'ORC', 6, 'ID', null, null, 'Ӧ����');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (88, 2, 'ORC', 7, 'TQ', null, null, '����/ʱ��');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (89, 2, 'ORC', 8, 'CM', null, null, '�����ʶ');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (90, 2, 'ORC', 9, 'TS', null, '[����ʱ��]', '��������/ʱ��');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (91, 2, 'ORC', 10, 'XCN', null, '[����ҽ��]', '¼����');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (92, 2, 'ORC', 11, 'XCN', null, '[У�Ի�ʿ]', 'У����');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (93, 2, 'ORC', 12, 'XCN', null, '[����ҽ��]', 'ҽ���ṩ��');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (270, 4, 'MSH', 3, 'HD', 'MEI MUSE', 'HIS001', 'Sending Facility �����豸');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (271, 4, 'MSH', 4, 'HD', 'ZLHIS', 'MUSE ECG Result 1', 'Receiving Application ���ճ���');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (272, 4, 'MSH', 5, 'HD', 'HIS001', 'MEI MUSE', 'Receiving Facility �����豸');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (273, 4, 'MSH', 6, 'TS', null, '[��ǰʱ��]', 'Date/Time Of Message ��Ϣ������/ʱ��');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (98, 2, 'OBR', 1, 'SI', null, '1', '����ID-OBR');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (99, 2, 'OBR', 2, 'EI', null, null, '������ҽ������');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (100, 2, 'OBR', 3, 'EI', null, null, 'ִ����ҽ������');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (101, 2, 'OBR', 4, 'CE', null, '[ҽ������]', 'ͨ�÷����ʶ');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (102, 2, 'OBR', 5, 'ID', null, null, '���ȼ�-OBR');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (103, 2, 'OBR', 6, 'TS', null, '[����ʱ��]', '��������/ʱ��');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (104, 2, 'OBR', 7, 'TS', null, '[����ʱ��]', '�۲�����/ʱ��#');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (105, 2, 'OBR', 8, 'TS', null, null, '�۲��������/ʱ��#');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (106, 2, 'OBR', 9, 'CQ', null, null, '�����ռ���*');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (107, 2, 'OBR', 10, 'XCN', null, null, '�ռ��߱�ʶ*');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (108, 2, 'OBR', 11, 'ID', null, null, '������Ϊ��*');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (109, 2, 'OBR', 12, 'CE', null, null, 'Σ�����ش���');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (110, 2, 'OBR', 13, 'ST', null, null, '����ٴ���Ϣ');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (111, 2, 'OBR', 14, 'TS', null, null, '�յ���������/ʱ��*');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (112, 2, 'OBR', 15, 'CM', null, null, '������Դ');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (113, 2, 'OBR', 16, 'XCN', null, '[����ҽ��]', 'ҽ���ṩ��');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (114, 2, 'OBR', 17, 'XTN', null, null, 'ҽ���طõ绰����');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (115, 2, 'OBR', 18, 'ST', null, null, '�������ֶ�1');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (116, 2, 'OBR', 19, 'ST', null, null, '�������ֶ�2');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (117, 2, 'OBR', 20, 'ST', null, null, 'ִ�����ֶ�1+');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (118, 2, 'OBR', 21, 'ST', null, null, 'ִ�����ֶ�2+');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (119, 2, 'OBR', 22, 'TS', null, null, '�������/״̬�ı�����/ʱ��+');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (120, 2, 'OBR', 23, 'CM', null, null, '�շ�+');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (121, 2, 'OBR', 24, 'ID', null, null, '��Ϸ����ʶID');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (122, 2, 'OBR', 25, 'ID', null, null, '���״̬+');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (123, 2, 'OBR', 26, 'CM', null, null, '������+');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (124, 2, 'OBR', 27, 'TQ', null, null, '����/ʱ��');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (125, 2, 'OBR', 28, 'XCN', null, null, '�����Ҫ��');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (126, 2, 'OBR', 29, 'CM', null, null, '����������');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (127, 2, 'OBR', 30, 'ID', null, null, '�����ж���ʽ');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (128, 2, 'OBR', 31, 'CE', null, null, '�о�ԭ��');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (129, 2, 'OBX', 1, 'SI', null, null, 'Set ID -OBX ����ID -OBX');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (130, 2, 'OBX', 2, 'ID', null, null, 'Value Type ֵ����');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (131, 2, 'OBX', 3, 'CE', null, null, 'Observation Identifier �۲��ʶ��');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (132, 2, 'OBX', 4, 'ST', null, null, 'Observation Sub -ID �۲���ID');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (133, 2, 'OBX', 5, 'TEXT', null, null, 'Observation Value �۲�ֵ');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (274, 4, 'MSH', 7, 'ST', null, null, 'Security ��ȫ��');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (275, 4, 'MSH', 8, 'CM', 'ORM^O01', 'ORM^O01', 'Message Type ��Ϣ����');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (276, 4, 'MSH', 9, 'ST', null, '[��ǰʱ��]', 'Message Control ID ��Ϣ����ID');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (277, 4, 'MSH', 10, 'PT', 'P', 'P', 'Processing ID ����ID');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (278, 4, 'MSH', 11, 'VID', '2.4', '2.4', 'Version ID �汾ID');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (279, 4, 'PID', 1, 'SI', null, '1', 'Set ID - Patient ID ����ID - ����ID');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (280, 4, 'PID', 2, 'CX', null, null, 'Patient ID ����ID');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (281, 4, 'PID', 3, 'CX', null, '[��ʶ��] ', 'Patient Identifier List ���߱�ʶ����');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (282, 4, 'PID', 4, 'CX', null, '[����ID]', 'Alternate Patient ID ���û���ID');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (283, 4, 'PID', 5, 'XPN', null, '[����]', 'Patient Name ��������');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (284, 4, 'PID', 6, 'XPN', null, null, 'Mother''s Maiden Name ĸ�׵Ļ�ǰ��');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (285, 4, 'PID', 7, 'TS', null, '[��������]', 'Date/Time of Birth ��������');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (286, 4, 'PID', 8, 'IS', null, '[�Ա�]', 'Sex �Ա�');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (287, 4, 'PID', 9, 'XPN', null, '[����]', 'Patient Alias ���߱���');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (288, 4, 'PID', 10, 'CE', null, 'C', 'Race ����');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (289, 4, 'PID', 11, 'XAD', null, '[��ϵ�˵�ַ]', 'Patient Address ����סַ');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (290, 4, 'PID', 12, 'IS', null, null, 'County Code �ش���');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (291, 4, 'PID', 13, 'XTN', null, '[��ͥ�绰]', 'Phone Number - Home ��ͥ�绰');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (292, 4, 'PID', 14, 'XTN', null, '[��ϵ�˵绰]', 'Phone Number - Business ��λ�绰');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (293, 4, 'PID', 15, 'CE', null, null, 'Primary Language ĸ��');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (294, 4, 'PID', 16, 'IS', null, '[����״��]', 'Marital Status ����״��');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (295, 4, 'PID', 17, 'CE', null, null, 'Religion �ڽ�����');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (296, 4, 'PID', 18, 'CX', null, '[���֤��]', 'Patient Account Number�����˺�');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (297, 4, 'PID', 19, 'ST', null, null, 'SSN Number - Patient ������ᱣ�պ���');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (298, 4, 'PV1', 1, 'SI', null, '1', 'Set ID-PV1����ID-PV1');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (299, 4, 'PV1', 2, 'IS', null, '[������Դ]', 'Patient Class �������');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (300, 4, 'PV1', 3, 'PL', null, '[��ǰ��������]^[��������]^[����]^^^^^^[��ǰ��������]', 'Assigned Patient Location ָ������λ��');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (301, 4, 'PV1', 4, 'IS', null, null, 'Admission Type ��Ժ����');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (302, 4, 'PV1', 5, 'CX', null, null, 'Preadmit Number Ԥ����Ժ����');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (303, 4, 'PV1', 6, 'PL', null, null, 'Prior Patient Location ����ԭλ��');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (304, 4, 'PV1', 7, 'XCN', null, '[����ҽ��]', 'Attending Doctor ����ҽ��');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (305, 4, 'PV1', 8, 'XCN', null, null, 'Referring Doctor ת��ҽ��');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (306, 4, 'PV1', 9, 'XCN', null, null, 'Consulting Doctor ����ҽ��');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (307, 4, 'PV1', 10, 'IS', null, null, 'Hospital Service ҽԺ����');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (308, 4, 'PV1', 11, 'PL', null, null, 'Temporary Location ��ʱλ��');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (309, 4, 'PV1', 12, 'IS', null, null, 'Preadmit Test Indicator Ԥ����Ժ�����ʶ��');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (310, 4, 'PV1', 13, 'IS', null, null, 'Readmission Indicator �ٴ���Ժ��ʶ��');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (311, 4, 'PV1', 14, 'IS', null, null, 'Admit Source ��Ժ��Դ');
commit;
prompt 300 records committed...
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (312, 4, 'PV1', 15, 'IS', null, null, 'Ambulatory Status ���������߶�״̬');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (313, 4, 'PV1', 16, 'IS', null, null, 'VIP Indicator VIP��ʶ��');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (314, 4, 'PV1', 17, 'XCN', null, null, 'Admitting Doctor ��Ժҽ��');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (315, 4, 'PV1', 18, 'IS', null, null, 'Patient Type ��������');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (316, 4, 'PV1', 19, 'CX', null, '1', 'Visit Number �������');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (317, 4, 'PV1', 20, 'FC', null, null, 'Financial Class ����״�����');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (318, 4, 'PV1', 21, 'IS', null, null, 'Charge Price Indicator ���ü۸��ʶ��');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (319, 4, 'PV1', 22, 'IS', null, null, 'Courtesy Code ��ò����');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (320, 4, 'PV1', 23, 'IS', null, null, 'Credit Rating ���õȼ�');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (321, 4, 'PV1', 24, 'IS', null, null, 'Contract Code ��ͬ����');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (322, 4, 'PV1', 25, 'DT', null, null, 'Contract Effective Date ��ͬ��Ч����');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (323, 4, 'PV1', 26, 'NM', null, null, 'Contract Amount ��ͬ����');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (324, 4, 'PV1', 27, 'NM', null, null, 'Contract Period ��ͬ����');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (325, 4, 'PV1', 28, 'IS', null, null, 'Interest Code ���ʴ���');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (326, 4, 'PV1', 29, 'IS', null, null, 'Transfer to Bad Debt Code תΪ���˴���');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (327, 4, 'PV1', 30, 'DT', null, null, 'Transfer to Bad Debt Date תΪ��������');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (328, 4, 'PV1', 31, 'IS', null, null, 'Bad Debt Agency Code ���˴������');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (329, 4, 'PV1', 32, 'NM', null, null, 'Bad Debt Transfer Amount תΪ�����ܶ�');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (330, 4, 'PV1', 33, 'NM', null, null, 'Bad Debt Recovery Amount ���˻ָ��ܶ�');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (331, 4, 'PV1', 34, 'IS', null, null, 'Delete Account Indicator ɾ���˻���ʶ��');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (332, 4, 'PV1', 35, 'DT', null, null, 'Delete Account Date ɾ���˻�����');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (333, 4, 'PV1', 36, 'IS', null, 'UNK', 'Discharge Disposition ��Ժ����');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (334, 4, 'PV1', 37, 'CM', null, null, 'Discharge to Location ��Ժȥ����λ��');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (335, 4, 'PV1', 38, 'CE', null, null, 'Diet Type ��ʳ����');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (336, 4, 'PV1', 39, 'IS', null, null, 'Servicing Facility �������');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (337, 4, 'PV1', 40, 'IS', null, null, 'Bed Status ��λ״��');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (338, 4, 'PV1', 41, 'IS', null, null, 'Account Status �˻�״��');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (339, 4, 'PV1', 42, 'IS', null, null, 'Pending Location ����λ��');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (340, 4, 'PV1', 43, 'PL', null, null, 'Prior Temporary Location ǰ��ʱλ��');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (341, 4, 'PV1', 44, 'TS', null, '[��Ժʱ��]', 'Admit Date/Time ��Ժ����/ʱ��');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (342, 4, 'PV1', 45, 'TS', null, '[��Ժʱ��]', 'Discharge Date/Time ��Ժ����/ʱ��');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (343, 4, 'PV1', 46, 'NM', null, null, 'Current Patient Balance ��ǰ���߲��');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (344, 4, 'PV1', 47, 'NM', null, null, 'Total Charges �ܷ���');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (345, 4, 'PV1', 48, 'NM', null, null, 'Total Adjustments �ܵ�����');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (346, 4, 'PV1', 49, 'NM', null, null, 'Total Payments ��֧����');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (347, 4, 'PV1', 50, 'CX', null, null, 'Alternate Visit ID ���þ���ID');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (348, 4, 'ORC', 1, 'ID', null, '[��Ϣ����]', 'ҽ��������');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (349, 4, 'ORC', 2, 'EI', null, '[ҽ��ID]', '������ҽ������');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (350, 4, 'ORC', 3, 'EI', null, null, 'ִ����ҽ������');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (351, 4, 'ORC', 4, 'EI', null, null, '������ҽ�������');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (352, 4, 'ORC', 5, 'ID', null, null, 'ҽ��״̬');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (353, 4, 'ORC', 6, 'ID', null, null, 'Ӧ����');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (354, 4, 'ORC', 7, 'TQ', null, null, '����/ʱ��');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (355, 4, 'ORC', 8, 'CM', null, null, '�����ʶ');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (356, 4, 'ORC', 9, 'TS', null, '[����ʱ��]', '��������/ʱ��');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (357, 4, 'ORC', 10, 'XCN', null, '[����ҽ��]', '¼����');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (358, 4, 'ORC', 11, 'XCN', null, '[У�Ի�ʿ]', 'У����');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (359, 4, 'ORC', 12, 'XCN', null, '[����ҽ��]', 'ҽ���ṩ��');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (364, 4, 'OBR', 1, 'SI', null, '1', '����ID-OBR');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (365, 4, 'OBR', 2, 'EI', null, null, '������ҽ������');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (366, 4, 'OBR', 3, 'EI', null, null, 'ִ����ҽ������');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (367, 4, 'OBR', 4, 'CE', null, '[ҽ������]', 'ͨ�÷����ʶ');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (368, 4, 'OBR', 5, 'ID', null, null, '���ȼ�-OBR');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (369, 4, 'OBR', 6, 'TS', null, '[����ʱ��]', '��������/ʱ��');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (370, 4, 'OBR', 7, 'TS', null, '[����ʱ��]', '�۲�����/ʱ��#');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (371, 4, 'OBR', 8, 'TS', null, null, '�۲��������/ʱ��#');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (372, 4, 'OBR', 9, 'CQ', null, null, '�����ռ���*');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (373, 4, 'OBR', 10, 'XCN', null, null, '�ռ��߱�ʶ*');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (374, 4, 'OBR', 11, 'ID', null, null, '������Ϊ��*');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (375, 4, 'OBR', 12, 'CE', null, null, 'Σ�����ش���');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (376, 4, 'OBR', 13, 'ST', null, null, '����ٴ���Ϣ');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (377, 4, 'OBR', 14, 'TS', null, null, '�յ���������/ʱ��*');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (378, 4, 'OBR', 15, 'CM', null, null, '������Դ');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (379, 4, 'OBR', 16, 'XCN', null, '[����ҽ��]', 'ҽ���ṩ��');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (380, 4, 'OBR', 17, 'XTN', null, null, 'ҽ���طõ绰����');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (381, 4, 'OBR', 18, 'ST', null, null, '�������ֶ�1');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (382, 4, 'OBR', 19, 'ST', null, null, '�������ֶ�2');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (383, 4, 'OBR', 20, 'ST', null, null, 'ִ�����ֶ�1+');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (384, 4, 'OBR', 21, 'ST', null, null, 'ִ�����ֶ�2+');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (385, 4, 'OBR', 22, 'TS', null, null, '�������/״̬�ı�����/ʱ��+');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (386, 4, 'OBR', 23, 'CM', null, null, '�շ�+');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (387, 4, 'OBR', 24, 'ID', null, null, '��Ϸ����ʶID');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (388, 4, 'OBR', 25, 'ID', null, null, '���״̬+');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (389, 4, 'OBR', 26, 'CM', null, null, '������+');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (390, 4, 'OBR', 27, 'TQ', null, null, '����/ʱ��');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (391, 4, 'OBR', 28, 'XCN', null, null, '�����Ҫ��');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (392, 4, 'OBR', 29, 'CM', null, null, '����������');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (393, 4, 'OBR', 30, 'ID', null, null, '�����ж���ʽ');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (394, 4, 'OBR', 31, 'CE', null, null, '�о�ԭ��');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (395, 4, 'OBX', 1, 'SI', null, null, 'Set ID -OBX ����ID -OBX');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (396, 4, 'OBX', 2, 'ID', null, null, 'Value Type ֵ����');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (397, 4, 'OBX', 3, 'CE', null, null, 'Observation Identifier �۲��ʶ��');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (398, 4, 'OBX', 4, 'ST', null, null, 'Observation Sub -ID �۲���ID');
insert into HL7��Ϣ������ (ID, ��ϢID, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������)
values (399, 4, 'OBX', 5, 'TEXT', null, null, 'Observation Value �۲�ֵ');
commit;
prompt 384 records loaded
set feedback on
set define on
prompt Done.
