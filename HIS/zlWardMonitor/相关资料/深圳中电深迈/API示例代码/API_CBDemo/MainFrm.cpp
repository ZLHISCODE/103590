//---------------------------------------------------------------------------

#include <vcl.h>
#include <stdio.h>
#pragma hdrstop

#include "MainFrm.h"
#include <boost/regex.hpp>
//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma resource "*.dfm"
TForm1 *Form1;
//E:\\Development\\CEC\\CecMonitorToHis\\  CecDeviceToHis.dll"//
#define DLL_FILE_NAME "CecDeviceToHis.dll"
//---------------------------------------------------------------------------
__fastcall TForm1::TForm1(TComponent* Owner)
	: TForm(Owner), connected_(false)
{
}
//---------------------------------------------------------------------------
void __fastcall TForm1::Button2Click(TObject *Sender)
{
	#if USE_API
	pfn_show_windows_((long)(void*)Panel1->Handle, 3);
	//bool bret = pfn_select_bedno_(1);
	//bret;
	#else
	CecMonitor1->ShowWindow((long)(void*)Panel1->Handle, 3);
	#endif
}
//---------------------------------------------------------------------------
void __fastcall TForm1::FormCreate(TObject *Sender)
{
	char text[100];
	#if USE_API
	module_ = LoadLibrary(DLL_FILE_NAME);
	sprintf(text, "%d��̬�ⲻ����!", DLL_FILE_NAME);
	if (!module_)
	{
		MessageBox(Handle, text, "��ʾ��Ϣ", MB_OK);
		Close();
	}
	pfn_initialize_ = (PFUN_INITIALIZE)GetProcAddress(module_, "CEC_Initialize");
	pfn_show_windows_ = (PFUN_SHOWWINDOWS)GetProcAddress(module_, "CEC_ShowWindows");
	pfn_uninitialize_ = (PFUN_UNINITIALIZE)GetProcAddress(module_, "CEC_Uninitialize");
	pfn_update_database_ = (PFUN_UPDATEDATABASE)GetProcAddress(module_, "CEC_UpdateDataBase");
	pfn_select_bedno_ = (PFUN_SELECTBEDNO)GetProcAddress(module_, "CEC_SelectBedNo");
	pfn_get_list_benno_ = (PFUN_GETLISTBEDNO)GetProcAddress(module_, "CEC_GetListBedNo");
	pfn_his_set_datatocec_ = (PFUN_HISSETDATATOCEC)GetProcAddress(module_, "CEC_HisSetDataToCec");
	pfn_get_monitor_data_ = (PFUN_GETMONITORDATA)GetProcAddress(module_, "CEC_GetMonitorData");
	pfn_his2devno_ = (PFUN_HIS2DEVNO)GetProcAddress(module_, "CEC_His2DevNo");
	pfn_devno2his_ = (PFUN_DEVNO2HIS)GetProcAddress(module_, "CEC_DevNo2His");
	if (!pfn_initialize_)
	{
		MessageBox(Handle, "CEC_Initialize ����������,����ò��ɹ�!", "��ʾ��Ϣ", MB_OK);
		Close();
	}
	#endif //USE_API

}
//---------------------------------------------------------------------------
void __fastcall TForm1::CecMonitor1MonitorMessage(TObject *Sender,
	  unsigned long nMonitorNo, unsigned long nCmd)
{
	Edit1->Text = IntToStr(nMonitorNo);
	Edit2->Text = IntToHex((__int64)nCmd, 8);
}
//---------------------------------------------------------------------------
void __stdcall TForm1::OnRecvMonitorMsg(unsigned long nMonitorNo, unsigned long nCmd, void* object)
{
	TForm1* pThis = static_cast<TForm1*>(object);
	pThis->Edit1->Text = IntToStr(nMonitorNo);
	pThis->Edit2->Text = IntToHex((__int64)nCmd, 8);
    /*
	#if USE_API
	unsigned short nMainCmd = ((nCmd&0xFFFF0000) >> 16);
	unsigned short nSubCmd = (nCmd&0x0000FFFF);
	unsigned char cBedNo = 0;
	if (nCmd == 0xFF)  //�޸��˴���
	{
       	if (0xFF == Message.LParam)
		{
			AnsiString Info = "�໤�Ǵ���|HIS����|������[������]|סԺ����|";
			Info += Edit3->Text+"|";
			Info += Edit4->Text + "|";
			if (RadioButton1->Checked)
				Info += "0|";
			else if (RadioButton2->Checked)
				Info += "1|";
			Info +="���|����|סԺ����|����|Ѫ��|����ҽ��";

			pfn_update_database_(Message.WParam, 3, Info.c_str());
		}
	}
	else if ((nSubCmd&0x00FF) == 0x01)  //����
	{
		if (0xFF == Message.LParam)
		{
			AnsiString Info = "�໤�Ǵ���|HIS����|������[������]|סԺ����|";
			Info += Edit3->Text+"|";
			Info += Edit4->Text + "|";
			if (RadioButton1->Checked)
				Info += "0|";
			else if (RadioButton2->Checked)
				Info += "1|";
			Info +="���|����|סԺ����|����|Ѫ��|����ҽ��";

			pfn_update_database_(Message.WParam, 3, Info.c_str());
		}
	}
	else if ((nSubCmd&0x00FF) == 0x02)  //ȷ��
	{
		if (0xFF == Message.LParam)
		{
			AnsiString Info = "�໤�Ǵ���|HIS����|������[������]|סԺ����|";
			Info += Edit3->Text+"|";
			Info += Edit4->Text + "|";
			if (RadioButton1->Checked)
				Info += "0|";
			else if (RadioButton2->Checked)
				Info += "1|";
			Info +="���|����|סԺ����|����|Ѫ��|����ҽ��";

			pfn_update_database_(Message.WParam, 3, Info.c_str());
		}
	}
	#endif
	*/
}
void __fastcall TForm1::OnRequestData(TMessage Message)
{
	// wParam Ϊ�໤�Ǳ��, lParam Ϊ������
	Edit1->Text = IntToStr(Message.WParam);
	Edit2->Text = IntToHex((__int64)Message.LParam, 8);
	AnsiString Info;
	if (0xFF == Message.LParam)
	{
		/*Info = "001|��1|002|2009072298|06|";//"�໤�Ǵ���|HIS����|����|������[������]|סԺ����|";
			Info += Edit3->Text+"|";
			Info += Edit4->Text + "|";
			if (RadioButton1->Checked)
				Info += "0|";
			else if (RadioButton2->Checked)
				Info += "1|";
			Info +="175|65|2009-07-22|1982-04-02|1|1|������|530125197810101591|0755-87654321|��������A-107";//"���|����|סԺ����|��������|����|Ѫ��|����ҽ��";

		pfn_update_database_(Message.WParam, 3, Info.c_str()); */
		char buf[200] = "\x0";
		if (pfn_get_monitor_data_)
		{
			if (pfn_get_monitor_data_(Message.WParam, 6, buf))
			{
				buf[strlen(buf)-1] = 0;
				std::string str = &buf[1];
				std::string exp = "[^\|]+";
				boost::regex expression(exp);
				boost::smatch what;
				std::string::const_iterator start = str.begin();
				std::string::const_iterator end = str.end();
				AnsiString field[2];
				int col = 0;
				while(boost::regex_search(start, end, what, expression))
				{
					start = what[0].second;
					field[col++] = what[0].str().c_str();
				}
				edtCaseNo->Text = field[1];
				edtHisNo->Text = field[0];
			}
		}
		
		
	}
	unsigned char cMainCmd = (Message.LParam>>16)&0x000000FF;
	unsigned char cSubCmd = Message.LParam&0x000000FF;
	switch(cMainCmd)
	{
		case 0x0A:
			if (0x01 == cSubCmd)
			{
				 /*Info = "{���ˣ�����|�Ա���|���䣺30|סԺ�ţ�HIS20080907|���ţ�BD20090807}; \
				 {ʱ��|����|������Ŀ|����|����|����}; \
				 {2004/4/20|�ڿ�|Ѫ��|2|100|200}; \
				 {2004/4/20|�ڿ�|Ѫ��|2|100|300}; \
				 {2004/4/20|�ڿ�|Ѫ��|2|100|400};{�ϼ�|900}";  */
				 Info = "���ˣ�����|�Ա���|���䣺30|סԺ�ţ�HIS20080907|���ţ�BD20090807^ \
				 ʱ��|����|������Ŀ|����|����|����^ \
				 2004/4/20|�ڿ�|Ѫ��|2|100|200^ \
				 2004/4/20|�ڿ�|Ѫ��|2|100|300^ \
				 2004/4/20|�ڿ�|Ѫ��|2|100|400^ \
				 �ϼ�|900";
				 if (pfn_his_set_datatocec_)
					pfn_his_set_datatocec_(Message.WParam, Message.LParam, Info.c_str());
				 char text[200] = "\x0";
				 if (pfn_devno2his_(Message.WParam, 3, text))
                     edtCaseNo->Text = text;
			}
			else if (0x02 == cSubCmd)
			{
			}
			else if (0x03 == cSubCmd)
			{
			}
			else if (0x04 == cSubCmd)
			{
			}
			break;
		case 0x0B:
			break;
		case 0x0C:
			break;
		case 0x0D:
			break;
		case 0x0F:
			Info = "��1|abcdef|002|2009072298|06|";//"�໤�Ǵ���|HIS����|����|������[������]|סԺ����|";
			Info += Edit3->Text+"|";
			Info += Edit4->Text + "|";
			if (radgSex->ItemIndex == 0)
				Info += "0|";
			else if (radgSex->ItemIndex == 1)
				Info += "1|";
			Info +="175|65|2009-07-22|1982-04-02|1|1|������|530125197810101591|0755-87654321|��������A-107";//"���|����|סԺ����|��������|����|Ѫ��|����ҽ��";

			pfn_update_database_(Message.WParam, 3, Info.c_str());
			break;
	}
}

void __fastcall TForm1::Panel1Resize(TObject *Sender)
{
	#if USE_API
	//pfn_get_handle_(handle);
	//SendMessage(handle, WM_RESIZE, 0,0);
	//pfn_set_window_pos_(Panel1->Width, Panel1->Height);
	#else
	//CecMonitor1->ShowWindow((long)(void*)Panel1->Handle, 4);
	#endif
}
//---------------------------------------------------------------------------


void __fastcall TForm1::bntConnectClick(TObject *Sender)
{
	if (!connected_)
	{
		#if USE_API
		#if USE_MSG //����Ϣ��ʱ,�ص���������Ϊ������
		if (pfn_initialize_)
			pfn_initialize_(edtIp->Text.c_str(), StrToInt(edtPort->Text),
				(unsigned long)(void*)Panel1->Handle, NULL, this->Handle);
		#else
		if (pfn_initialize_)
			pfn_initialize_(edtIp->Text.c_str(), StrToInt(edtPort->Text),
				(unsigned long)(void*)Panel1->Handle, OnRecvMonitorMsg, (void*)this);
		#endif //USE_MSG
		#else
		CecMonitor1 = new TCecMonitor(this);
		CecMonitor1->OnMonitorMessage = CecMonitor1MonitorMessage;
		wchar_t ip[20];
		Utf8ToUnicode(ip, wcslen(ip), SRV_IP, strlen(SRV_IP));
		CecMonitor1->Initialize(ip, SRV_PORT, (unsigned long)(void*)Panel1->Handle);
		#endif //USE_API
		bntConnect->Caption = "�Ͽ�����";
		connected_ = true;
	}
	else
	{
		#if USE_API
		if (pfn_uninitialize_)
			pfn_uninitialize_();
		#else
		CecMonitor1->Uninitialize();
		if (CecMonitor1)
			delete CecMonitor1;
		#endif
		bntConnect->Caption = "���ӷ���";
		connected_ = false;
	}
}
//---------------------------------------------------------------------------

void __fastcall TForm1::FormClose(TObject *Sender, TCloseAction &Action)
{
	  if (connected_)
	  {
	  	#if USE_API
		if (pfn_uninitialize_)
			pfn_uninitialize_();
		#else
		CecMonitor1->Uninitialize();
		if (CecMonitor1)
			delete CecMonitor1;
		#endif //
	  }
	  #if USE_API
	  if (module_)
	  	FreeLibrary(module_);
	  #endif //USE_API
}
//---------------------------------------------------------------------------


void __fastcall TForm1::BtnListClick(TObject *Sender)
{
	//
	unsigned char list_bedno[200] = "\x0";
	if (pfn_get_list_benno_)
		pfn_get_list_benno_(list_bedno);
	{
		AnsiString message = "���Ӵ���:";
		message += (char*)list_bedno;
		MessageBox(Handle, message.c_str(), "��ʾ��Ϣ", MB_OK);
	}
}
//---------------------------------------------------------------------------

void __fastcall TForm1::btnSelectNoClick(TObject *Sender)
{
	//
	if (pfn_select_bedno_)
	{
		if (pfn_select_bedno_(pfn_his2devno_(radgType->ItemIndex+1, edtSelectCaseNo->Text.c_str())))
		{
		}
		else
		{
		}
	}
}
//---------------------------------------------------------------------------



void __fastcall TForm1::btnDev2HisClick(TObject *Sender)
{
	char buf[50];
	if (pfn_devno2his_(StrToInt(edtSelectCaseNo->Text), radgType->ItemIndex+1, buf))
	{
		MessageBox(Handle,  buf, "��ʾ��Ϣ", MB_OK);
	}
}
//---------------------------------------------------------------------------

