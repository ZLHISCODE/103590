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
	sprintf(text, "%d动态库不存在!", DLL_FILE_NAME);
	if (!module_)
	{
		MessageBox(Handle, text, "提示信息", MB_OK);
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
		MessageBox(Handle, "CEC_Initialize 函数不存在,或调用不成功!", "提示信息", MB_OK);
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
	if (nCmd == 0xFF)  //修改了床号
	{
       	if (0xFF == Message.LParam)
		{
			AnsiString Info = "监护仪床号|HIS床号|病历号[病案号]|住院次数|";
			Info += Edit3->Text+"|";
			Info += Edit4->Text + "|";
			if (RadioButton1->Checked)
				Info += "0|";
			else if (RadioButton2->Checked)
				Info += "1|";
			Info +="身高|体重|住院日期|类型|血型|主治医生";

			pfn_update_database_(Message.WParam, 3, Info.c_str());
		}
	}
	else if ((nSubCmd&0x00FF) == 0x01)  //请求
	{
		if (0xFF == Message.LParam)
		{
			AnsiString Info = "监护仪床号|HIS床号|病历号[病案号]|住院次数|";
			Info += Edit3->Text+"|";
			Info += Edit4->Text + "|";
			if (RadioButton1->Checked)
				Info += "0|";
			else if (RadioButton2->Checked)
				Info += "1|";
			Info +="身高|体重|住院日期|类型|血型|主治医生";

			pfn_update_database_(Message.WParam, 3, Info.c_str());
		}
	}
	else if ((nSubCmd&0x00FF) == 0x02)  //确认
	{
		if (0xFF == Message.LParam)
		{
			AnsiString Info = "监护仪床号|HIS床号|病历号[病案号]|住院次数|";
			Info += Edit3->Text+"|";
			Info += Edit4->Text + "|";
			if (RadioButton1->Checked)
				Info += "0|";
			else if (RadioButton2->Checked)
				Info += "1|";
			Info +="身高|体重|住院日期|类型|血型|主治医生";

			pfn_update_database_(Message.WParam, 3, Info.c_str());
		}
	}
	#endif
	*/
}
void __fastcall TForm1::OnRequestData(TMessage Message)
{
	// wParam 为监护仪编号, lParam 为控制字
	Edit1->Text = IntToStr(Message.WParam);
	Edit2->Text = IntToHex((__int64)Message.LParam, 8);
	AnsiString Info;
	if (0xFF == Message.LParam)
	{
		/*Info = "001|加1|002|2009072298|06|";//"监护仪床号|HIS床号|科室|病历号[病案号]|住院次数|";
			Info += Edit3->Text+"|";
			Info += Edit4->Text + "|";
			if (RadioButton1->Checked)
				Info += "0|";
			else if (RadioButton2->Checked)
				Info += "1|";
			Info +="175|65|2009-07-22|1982-04-02|1|1|刘国华|530125197810101591|0755-87654321|深圳南油A-107";//"身高|体重|住院日期|出生日期|类型|血型|主治医生";

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
				 /*Info = "{病人：张三|性别：男|年龄：30|住院号：HIS20080907|床号：BD20090807}; \
				 {时间|科室|费用项目|数量|单价|费用}; \
				 {2004/4/20|内科|血检|2|100|200}; \
				 {2004/4/20|内科|血检|2|100|300}; \
				 {2004/4/20|内科|血检|2|100|400};{合计|900}";  */
				 Info = "病人：张三|性别：男|年龄：30|住院号：HIS20080907|床号：BD20090807^ \
				 时间|科室|费用项目|数量|单价|费用^ \
				 2004/4/20|内科|血检|2|100|200^ \
				 2004/4/20|内科|血检|2|100|300^ \
				 2004/4/20|内科|血检|2|100|400^ \
				 合计|900";
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
			Info = "普1|abcdef|002|2009072298|06|";//"监护仪床号|HIS床号|科室|病历号[病案号]|住院次数|";
			Info += Edit3->Text+"|";
			Info += Edit4->Text + "|";
			if (radgSex->ItemIndex == 0)
				Info += "0|";
			else if (radgSex->ItemIndex == 1)
				Info += "1|";
			Info +="175|65|2009-07-22|1982-04-02|1|1|刘国华|530125197810101591|0755-87654321|深圳南油A-107";//"身高|体重|住院日期|出生日期|类型|血型|主治医生";

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
		#if USE_MSG //用消息传时,回调函数设置为空则行
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
		bntConnect->Caption = "断开服务";
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
		bntConnect->Caption = "连接服务";
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
		AnsiString message = "连接床号:";
		message += (char*)list_bedno;
		MessageBox(Handle, message.c_str(), "提示信息", MB_OK);
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
		MessageBox(Handle,  buf, "提示信息", MB_OK);
	}
}
//---------------------------------------------------------------------------

