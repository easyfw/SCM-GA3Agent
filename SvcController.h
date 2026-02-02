//---------------------------------------------------------------------------
#ifndef SvcControllerH
#define SvcControllerH
//---------------------------------------------------------------------------
#include <SysUtils.hpp>
#include <Classes.hpp>
#include <SvcMgr.hpp>
#include <vcl.h>
#include "VaClasses.hpp"
#include "VaComm.hpp"

// OPC Automation 헤더
#include "OPCAutomation_TLB.h"
#include <ExtCtrls.hpp>

using namespace Opcautomation_tlb;

typedef IOPCAutoServerPtr _di_IOPCAutoServer;
typedef IOPCGroupsPtr     _di_IOPCGroups;
typedef IOPCGroupPtr      _di_IOPCGroup;

typedef OPCItemsPtr       _di_IOPCItems;
typedef OPCItemPtr        _di_IOPCItem;


// 프로토콜 상수
#define PROTO_STX       0x02
#define PROTO_ETX       0x03
#define MAX_OPC_ITEMS   500

// 응답 코드
#define RESP_CMD_ACK    0x01
#define RESP_CMD_NAK    0x02
#define RESP_STATUS_OK  0x00
#define RESP_STATUS_CHK 0x01
#define RESP_STATUS_LEN 0x02
#define RESP_STATUS_TMO 0x03
#define RESP_TIMEOUT_MS 5000

#define HK_DEBUG		0		// debug enable
#define	SERVER_SIMULATE	0		// 시뮬레이션 모드

// OPC 아이템 정보 구조체
struct TOPCItemInfo
{
    int         ItemID;
    String      TagName;
    String      DataType;
    String      Description;
    OPCItem*    pItem;          // _di_IOPCItem 대신 OPCItem* 사용
    VARIANT     varValue;
    VARIANT     varPrevValue;
    long        Quality;
    bool        Changed;
};

//---------------------------------------------------------------------------
class TSCM_Ga3Agent : public TService
{
__published:    // IDE-managed Components
	TVaComm *Mycomm;
	TTimer *Timer1;

	void __fastcall Timer1Timer(TObject *Sender);
    void __fastcall ServiceStart(TService *Sender, bool &Started);
    void __fastcall ServiceStop(TService *Sender, bool &Stopped);

private:        // User declarations

    // OPC 관련 (기존)
    _di_IOPCAutoServer OPCServer;
    _di_IOPCGroups     OPCGroups;
    _di_IOPCGroup      MyGroup;
    _di_IOPCItems      MyItems;

     // === INI 설정 변수 ===
    int m_nComPort;         // COM 포트 번호 (숫자만)
    int m_nBaudRate;        // 통신 속도
    int m_nTimeInterval;    // 타이머 간격 (ms)
        
    // 아이템 배열
    TOPCItemInfo    m_Items[MAX_OPC_ITEMS];
    int             m_ItemCount;

    // 시리얼 통신 관련
    bool            m_bCommOpened;
    BYTE            m_SendBuffer[1024];
    bool            m_bFirstSend;

	// 응답 관련
	int             m_nRetryCount;
	int             m_nMaxRetries;
	bool            m_bWaitingResponse;
    DWORD           m_dwLastSendTick;       // 마지막 전송 시간
	DWORD           m_dwHeartbeatInterval;  // Heartbeat 주기 (ms)

    // 로그
    TCHAR gbuf[65535];

    // 내부 함수 - 기존
    void __fastcall LogMessage(String msg);
    String __fastcall VariantToString(const tagVARIANT &v);
    int __fastcall GetQualityCode(long quality);

        // === 설정 로드 함수 ===
    void __fastcall LoadSettings();

    // 내부 함수 - CSV 로드
    bool __fastcall LoadItemConfig(String filename);
    
    // 내부 함수 - 시리얼 통신
    bool __fastcall InitSerialPort(int portNum, int baudRate);
    void __fastcall CloseSerialPort();
    BYTE __fastcall CalcChecksum(BYTE* data, int len);
    int __fastcall BuildPacket(BYTE* buffer);
    void __fastcall SendToESP32(int changeCount = 0, bool isHeartbeat = false);

    // 내부 함수 - 값 비교
    bool __fastcall IsValueChanged(int index);
    bool __fastcall HasAnyChanges();
    long __fastcall VariantToLong(const VARIANT &v);

	bool __fastcall WaitForResponse(int timeoutMs);
	void __fastcall HandleSendFailure();
 

public:         // User declarations

	__fastcall TSCM_Ga3Agent(TComponent* Owner);
	TServiceController __fastcall GetServiceController(void);

	friend void __stdcall ServiceController(unsigned CtrlCode);
};
//---------------------------------------------------------------------------
extern PACKAGE TSCM_Ga3Agent *SCM_Ga3Agent;
//---------------------------------------------------------------------------
#endif
