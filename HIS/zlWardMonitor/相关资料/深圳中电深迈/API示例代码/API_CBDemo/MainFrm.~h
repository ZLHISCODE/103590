//---------------------------------------------------------------------------

#ifndef MainFrmH
#define MainFrmH
//---------------------------------------------------------------------------
#include <Classes.hpp>
#include <Controls.hpp>
#include <StdCtrls.hpp>
#include <Forms.hpp>
#include <OleServer.hpp>
#include <ExtCtrls.hpp>
//////////////////////////////////////////////////////////////
typedef void (__stdcall *PFUN_CALLBACE)(unsigned long, unsigned long, void* object);
//////////////////////////////////////////////////////////////
typedef bool (__stdcall *PFUN_INITIALIZE)(char*, unsigned long, unsigned long, PFUN_CALLBACE, void*);
typedef bool (__stdcall *PFUN_SHOWWINDOWS)(unsigned long, unsigned long);
typedef bool (__stdcall *PFUN_UNINITIALIZE)();
typedef bool (__stdcall *PFUN_SETWINDOWPOS)(int, int);
typedef bool (__stdcall *PFUN_UPDATEDATABASE)(unsigned long, unsigned long, char*);
typedef bool (__stdcall *PFUN_SELECTBEDNO)(unsigned long);
typedef bool (__stdcall *PFUN_GETLISTBEDNO)(unsigned char*);
typedef bool (__stdcall *PFUN_HISSETDATATOCEC)(unsigned long, unsigned long, char*);
typedef bool (__stdcall *PFUN_GETMONITORDATA)(unsigned long, unsigned long, char*);
typedef unsigned long (__stdcall *PFUN_HIS2DEVNO)(unsigned long, char*);
typedef bool (__stdcall *PFUN_DEVNO2HIS)(unsigned long, unsigned long, char*);
//////////////////////////////////////////
#define WM_REQUEST_DATA (WM_USER+620)
//////////////////////////////////////////
#define USE_API	1


#if USE_API
#define USE_MSG 1 //调用API,并用消息传数据
#define SRV_IP "192.168.1.254"
#define SRV_PORT 6008
#else
#include "CecMonitorToHis_OCX.h"
#define SRV_IP "192.168.1.254"
#define SRV_PORT 6008
#endif
//---------------------------------------------------------------------------
class TForm1 : public TForm
{
__published:	// IDE-managed Components
	TPanel *Panel1;
	TPanel *Panel2;
	TButton *Button2;
	TEdit *Edit1;
	TEdit *Edit2;
	TLabel *Label1;
	TEdit *Edit3;
	TLabel *Label2;
	TEdit *Edit4;
	TButton *bntConnect;
	TLabel *Label3;
	TEdit *edtIp;
	TLabel *Label4;
	TEdit *edtPort;
	TButton *BtnList;
	TLabel *Label5;
	TLabel *Label6;
	TEdit *edtCaseNo;
	TEdit *edtHisNo;
	TButton *btnSelectNo;
	TEdit *edtSelectCaseNo;
	TRadioGroup *radgType;
	TRadioGroup *radgSex;
	TButton *btnDev2His;
	void __fastcall Button2Click(TObject *Sender);
	void __fastcall FormCreate(TObject *Sender);
	void __fastcall CecMonitor1MonitorMessage(TObject *Sender,
		  unsigned long nMonitorNo, unsigned long nCmd);
	void __fastcall Panel1Resize(TObject *Sender);
	void __fastcall bntConnectClick(TObject *Sender);
	void __fastcall FormClose(TObject *Sender, TCloseAction &Action);
	void __fastcall BtnListClick(TObject *Sender);
	void __fastcall btnSelectNoClick(TObject *Sender);
	void __fastcall btnDev2HisClick(TObject *Sender);
protected:
	void __fastcall OnRequestData(TMessage Message);
	BEGIN_MESSAGE_MAP
		MESSAGE_HANDLER(WM_REQUEST_DATA, TMessage, OnRequestData);
	END_MESSAGE_MAP(TForm)
private:	// User declarations
	#if USE_API
	HMODULE module_;
	PFUN_INITIALIZE pfn_initialize_;
	PFUN_SHOWWINDOWS pfn_show_windows_;
	PFUN_UNINITIALIZE pfn_uninitialize_;
	PFUN_SETWINDOWPOS pfn_set_window_pos_;
	PFUN_UPDATEDATABASE pfn_update_database_;
	PFUN_SELECTBEDNO pfn_select_bedno_;
	PFUN_GETLISTBEDNO pfn_get_list_benno_;
	PFUN_HISSETDATATOCEC pfn_his_set_datatocec_;
	PFUN_GETMONITORDATA pfn_get_monitor_data_;
	PFUN_HIS2DEVNO pfn_his2devno_;
	PFUN_DEVNO2HIS pfn_devno2his_;
	#else
	TCecMonitor *CecMonitor1;
	#endif
	bool connected_;
public:		// User declarations
	__fastcall TForm1(TComponent* Owner);
	static void __stdcall OnRecvMonitorMsg(unsigned long nMonitorNo, unsigned long nCmd, void* object);
};
//---------------------------------------------------------------------------
extern PACKAGE TForm1 *Form1;
//---------------------------------------------------------------------------
#endif


