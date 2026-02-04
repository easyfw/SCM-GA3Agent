//---------------------------------------------------------------------------
#include "SvcController.h"
//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma link "VaClasses"
#pragma link "VaComm"
#pragma resource "*.dfm"

TSCM_Ga3Agent *SCM_Ga3Agent;

// 로그 파일 관리용 전역 변수
static int g_LogFileIndex = 0;
static bool g_bFirstRun = true;

//---------------------------------------------------------------------------
__fastcall TSCM_Ga3Agent::TSCM_Ga3Agent(TComponent* Owner)
	: TService(Owner)
{
    this->OnStart = ServiceStart;
    this->OnStop  = ServiceStop;
    lstrcpy(gbuf, "[SCM-GA3 Service Log]\r\n");

    m_ItemCount = 0;
    m_bCommOpened = false;
    m_bFirstSend = true;

    // 응답 관련 초기화
    m_nRetryCount = 0;
    m_nMaxRetries = 3;
    m_bWaitingResponse = false;
    
    // === Heartbeat 관련 초기화 추가 ===
    m_dwLastSendTick = 0;
    m_dwHeartbeatInterval = 5000;  // 5초 (필요시 조정)

    // === INI 설정 기본값 ===
    m_nComPort = 3;
    m_nBaudRate = 115200;
    m_nTimeInterval = 5000;
}

TServiceController __fastcall TSCM_Ga3Agent::GetServiceController(void)
{
	return (TServiceController) ServiceController;
}

void __stdcall ServiceController(unsigned CtrlCode)
{
	SCM_Ga3Agent->Controller(CtrlCode);
}

//
void __fastcall TSCM_Ga3Agent::LogMessage(String msg)
{
    HANDLE hFile;
    DWORD dwBytesWritten;
    DWORD dwFileSize;
    SYSTEMTIME st;
    String logFileName;
    
    String exePath = ExtractFilePath(ParamStr(0));
    String logBasePath = exePath + "logsave";

    while (true)
    {
        if (g_LogFileIndex == 0) logFileName = logBasePath + ".txt";
        else logFileName = logBasePath + "_" + IntToStr(g_LogFileIndex) + ".txt";

        hFile = CreateFile(
            logFileName.c_str(),
            GENERIC_WRITE,
            FILE_SHARE_READ,
            NULL,
            OPEN_ALWAYS,
            FILE_ATTRIBUTE_NORMAL,
            NULL
        );

        if (hFile == INVALID_HANDLE_VALUE) return;

        dwFileSize = GetFileSize(hFile, NULL);

        if (dwFileSize > 60000)
        {
            CloseHandle(hFile);
            g_LogFileIndex++;
            g_bFirstRun = false;
            continue;
        }
        break;
    }

    SetFilePointer(hFile, 0, NULL, FILE_END);

    if (g_bFirstRun)
    {
        if (dwFileSize > 0)
        {
            String blankLine = "\r\n";
            WriteFile(hFile, blankLine.c_str(), blankLine.Length(), &dwBytesWritten, NULL);
        }
        g_bFirstRun = false;
    }

    GetLocalTime(&st);
    
    // === 변경: 날짜 제거, 시간만 출력 (HH:MM:SS) ===
    String timeStr;
    timeStr.printf("[%02d:%02d:%02d] ", st.wHour, st.wMinute, st.wSecond);

    AnsiString finalMsg = timeStr + msg + "\r\n";
    WriteFile(hFile, finalMsg.c_str(), finalMsg.Length(), &dwBytesWritten, NULL);
    CloseHandle(hFile);
}

//---------------------------------------------------------------------------
// VARIANT를 문자열로 변환
//---------------------------------------------------------------------------
String __fastcall TSCM_Ga3Agent::VariantToString(const tagVARIANT &v)
{
    switch (v.vt)
    {
        case VT_EMPTY:  return "(EMPTY)";
        case VT_NULL:   return "(NULL)";
        case VT_I1:     return IntToStr((int)v.cVal);
        case VT_UI1:    return IntToStr((int)v.bVal);
        case VT_I2:     return IntToStr(v.iVal);
        case VT_UI2:    return IntToStr(v.uiVal);
        case VT_I4:     return IntToStr(v.lVal);
        case VT_UI4:    return IntToStr((int)v.ulVal);
        case VT_INT:    return IntToStr(v.intVal);
        case VT_UINT:   return IntToStr((int)v.uintVal);
        case VT_R4:     return FloatToStrF(v.fltVal, ffFixed, 10, 3);
        case VT_R8:     return FloatToStrF(v.dblVal, ffFixed, 15, 6);
        case VT_BOOL:   return (v.boolVal == VARIANT_TRUE) ? "TRUE" : "FALSE";
        case VT_BSTR:   return String(v.bstrVal);
        default:        return "(VT=" + IntToStr(v.vt) + ")";
    }
}

//---------------------------------------------------------------------------
// Quality 코드 변환
//---------------------------------------------------------------------------
int __fastcall TSCM_Ga3Agent::GetQualityCode(long quality)
{
    int majorQuality = quality & 0xC0;

    if (majorQuality == 0xC0) return 0;   // Good
    else if (majorQuality == 0x40) return 1;   // Uncertain
    else
    {
        switch (quality)
        {
            case 0x08: return 2;    // Not Connected
            case 0x18: return 3;    // Comm Failure
            case 0x0C: return 4;    // Device Failure
            case 0x10: return 5;    // Sensor Failure
            default:   return 9;    // Other Error
        }
    }
}

//---------------------------------------------------------------------------
// VARIANT를 long으로 변환 (비교용)
//---------------------------------------------------------------------------
long __fastcall TSCM_Ga3Agent::VariantToLong(const VARIANT &v)
{
    switch (v.vt)
    {
        case VT_I1:   return (long)v.cVal;
        case VT_UI1:  return (long)v.bVal;
        case VT_I2:   return (long)v.iVal;
        case VT_UI2:  return (long)v.uiVal;
        case VT_I4:   return v.lVal;
        case VT_UI4:  return (long)v.ulVal;
        case VT_INT:  return (long)v.intVal;
        case VT_UINT: return (long)v.uintVal;
        case VT_R4:   return (long)(v.fltVal * 1000);
        case VT_R8:   return (long)(v.dblVal * 1000);
        case VT_BOOL: return v.boolVal ? 1 : 0;
		case VT_DATE: return (long)((v.date - 25569.0) * 86400.0); // OLE DATE (1899-12-30 기준) → Unix timestamp (1970-01-01 기준)
        default:      return 0;
    }
}

//---------------------------------------------------------------------------
// INI 설정 파일 로드
//---------------------------------------------------------------------------
void __fastcall TSCM_Ga3Agent::LoadSettings()
{
    String IniPath = ExtractFilePath(ParamStr(0)) + "oem_setting.ini";
    
    TIniFile *ini = new TIniFile(IniPath);
    try
    {
        // [Communication] 섹션
        String comStr = ini->ReadString("Communication", "COM_Port", "COM3");

        // "COM17" -> 17 추출
        if (comStr.UpperCase().Pos("COM") == 1) m_nComPort = StrToIntDef(comStr.SubString(4, comStr.Length() - 3), 17);
        else m_nComPort = StrToIntDef(comStr, 17);

        m_nBaudRate = ini->ReadInteger("Communication", "BaudRate", 115200);

        // [Agent] 섹션
        m_nTimeInterval = ini->ReadInteger("Agent", "TimeInterval", 5000);
        LogMessage("CFG: COM" + IntToStr(m_nComPort) + " " + IntToStr(m_nBaudRate) + " T:" + IntToStr(m_nTimeInterval));
    }
    __finally
    {
        delete ini;
    }
}

//---------------------------------------------------------------------------
// CSV 파일에서 아이템 설정 로드
//---------------------------------------------------------------------------
bool __fastcall TSCM_Ga3Agent::LoadItemConfig(String filename)
{
    TStringList *lines = new TStringList();
    m_ItemCount = 0;

    try
    {
        if (!FileExists(filename))
        {
            LogMessage("Config file not found: " + filename);
            delete lines;
            return false;
        }

        lines->LoadFromFile(filename);
        LogMessage("Loading config: " + filename + " (" + IntToStr(lines->Count) + " lines)");

        for (int i = 1; i < lines->Count && m_ItemCount < MAX_OPC_ITEMS; i++)
        {
            String line = lines->Strings[i].Trim();
            if (line.IsEmpty() || line[1] == '#')
                continue;

            // 수동 CSV 파싱 (StrictDelimiter 대신)
            String col0 = "";
            String col1 = "";
            String col2 = "";
            String col3 = "";
            int colIndex = 0;
            String temp = "";

            for (int j = 1; j <= line.Length(); j++)
            {
                if (line[j] == ',')
                {
                    switch (colIndex)
                    {
                        case 0: col0 = temp.Trim(); break;
                        case 1: col1 = temp.Trim(); break;
                        case 2: col2 = temp.Trim(); break;
                        case 3: col3 = temp.Trim(); break;
                    }
                    temp = "";
                    colIndex++;
                }
                else
                {
                    temp += line[j];
                }
            }
            // 마지막 컬럼
            switch (colIndex)
            {
                case 0: col0 = temp.Trim(); break;
                case 1: col1 = temp.Trim(); break;
                case 2: col2 = temp.Trim(); break;
                case 3: col3 = temp.Trim(); break;
            }

            if (!col0.IsEmpty() && !col1.IsEmpty() && !col2.IsEmpty())
            {
                m_Items[m_ItemCount].ItemID = StrToIntDef(col0, 0);
                m_Items[m_ItemCount].TagName = col1;
                m_Items[m_ItemCount].DataType = col2.UpperCase();
                m_Items[m_ItemCount].Description = col3;  // 빈 문자열이면 그냥 빈 문자열
                m_Items[m_ItemCount].pItem = NULL;
                m_Items[m_ItemCount].Quality = 0;
                m_Items[m_ItemCount].Changed = false;

                VariantInit(&m_Items[m_ItemCount].varValue);
                VariantInit(&m_Items[m_ItemCount].varPrevValue);

                LogMessage("  Item[" + IntToStr(m_ItemCount) + "]: ID=" +
                           IntToStr(m_Items[m_ItemCount].ItemID) +
                           ", Tag=" + m_Items[m_ItemCount].TagName +
                           ", Type=" + m_Items[m_ItemCount].DataType);

                m_ItemCount++;
            }
        }
        LogMessage("Loaded " + IntToStr(m_ItemCount) + " items from config.");
    }
    catch (Exception &ex)
    {
        LogMessage("Error loading config: " + ex.Message);
        delete lines;
        return false;
    }

    delete lines;
    return (m_ItemCount > 0);
}

//---------------------------------------------------------------------------
// 시리얼 포트 초기화
//---------------------------------------------------------------------------
bool __fastcall TSCM_Ga3Agent::InitSerialPort(int portNum, int baudRate)
{
    try
    {
        if (Mycomm == NULL)
        {
            LogMessage("Error: Mycomm component is NULL");
            return false;
        }

        // 이미 열려있으면 닫기
        if (Mycomm->Active()) Mycomm->Close();

        // 포트 설정
        Mycomm->PortNum = portNum;

        // BaudRate는 enum 타입이므로 변환 필요
        switch (baudRate)
        {
            case 9600:   Mycomm->Baudrate = br9600;   break;
            case 19200:  Mycomm->Baudrate = br19200;  break;
            case 38400:  Mycomm->Baudrate = br38400;  break;
            case 57600:  Mycomm->Baudrate = br57600;  break;
            case 115200: Mycomm->Baudrate = br115200; break;
            default:     Mycomm->Baudrate = br115200; break;
        }

        Mycomm->Databits = db8;
        Mycomm->Stopbits = sb1;
        Mycomm->Parity = paNone;

        // 포트 열기
        Mycomm->Open();

        if (Mycomm->Active())
        {
            m_bCommOpened = true;
            LogMessage("Serial port COM" + IntToStr(portNum) +
                       " opened at " + IntToStr(baudRate) + " bps");
            return true;
        }
        else
        {
            LogMessage("Failed to open COM" + IntToStr(portNum));
            return false;
        }
    }
    catch (Exception &ex)
    {
        LogMessage("Serial port error: " + ex.Message);
        return false;
    }
}

//---------------------------------------------------------------------------
// 시리얼 포트 닫기
//---------------------------------------------------------------------------
void __fastcall TSCM_Ga3Agent::CloseSerialPort()
{
    try
    {
        if (Mycomm && Mycomm->Active())
        {
            Mycomm->Close();
            m_bCommOpened = false;
            LogMessage("Serial port closed.");
        }
    }
    catch (Exception &ex)
    {
        LogMessage("Error closing serial port: " + ex.Message);
    }
}

//---------------------------------------------------------------------------
// 체크섬 계산 (XOR)
//---------------------------------------------------------------------------
BYTE __fastcall TSCM_Ga3Agent::CalcChecksum(BYTE* data, int len)
{
    BYTE checksum = 0;
    for (int i = 0; i < len; i++)
    {
        checksum ^= data[i];
    }
    return checksum;
}

//---------------------------------------------------------------------------
// 값 변경 여부 확인
//---------------------------------------------------------------------------
bool __fastcall TSCM_Ga3Agent::IsValueChanged(int index)
{
    if (index < 0 || index >= m_ItemCount)
        return false;

    long currVal = VariantToLong(m_Items[index].varValue);
    long prevVal = VariantToLong(m_Items[index].varPrevValue);

    return (currVal != prevVal);
}

//---------------------------------------------------------------------------
// 변경된 아이템이 있는지 확인
//---------------------------------------------------------------------------
bool __fastcall TSCM_Ga3Agent::HasAnyChanges()
{
    bool hasChange = false;
    for (int i = 0; i < m_ItemCount; i++)
    {
        if (IsValueChanged(i))
        {
            m_Items[i].Changed = true;
            hasChange = true;
        }
    }
    return hasChange;
}

//---------------------------------------------------------------------------
// 패킷 생성
// 프로토콜: [STX][LEN_L][LEN_H][CNT][ID_L][ID_H][Q][VAL0][VAL1][VAL2][VAL3]...[CHK][ETX]
//---------------------------------------------------------------------------
int __fastcall TSCM_Ga3Agent::BuildPacket(BYTE* buffer)
{
    int pos = 0;

    // STX
    buffer[pos++] = PROTO_STX;

    // Length (나중에 채움)
    int lenPos = pos;
    pos += 2;

    // Item Count
    buffer[pos++] = (BYTE)m_ItemCount;

    // 각 아이템 데이터
    for (int i = 0; i < m_ItemCount; i++)
    {
        // Item ID (2 bytes, Little Endian)
        WORD itemId = (WORD)m_Items[i].ItemID;
        buffer[pos++] = (BYTE)(itemId & 0xFF);
        buffer[pos++] = (BYTE)((itemId >> 8) & 0xFF);

        // Quality (1 byte)
        buffer[pos++] = (BYTE)GetQualityCode(m_Items[i].Quality);

        // Value (4 bytes, Little Endian)
        long value = VariantToLong(m_Items[i].varValue);
        buffer[pos++] = (BYTE)(value & 0xFF);
        buffer[pos++] = (BYTE)((value >> 8) & 0xFF);
        buffer[pos++] = (BYTE)((value >> 16) & 0xFF);
        buffer[pos++] = (BYTE)((value >> 24) & 0xFF);
    }

    // Length 채우기 (STX 제외, Checksum/ETX 제외한 데이터 길이)
    WORD dataLen = pos - 3;
    buffer[lenPos] = (BYTE)(dataLen & 0xFF);
    buffer[lenPos + 1] = (BYTE)((dataLen >> 8) & 0xFF);

    // Checksum (STX 다음부터 데이터 끝까지)
    buffer[pos] = CalcChecksum(&buffer[1], pos - 1);
    pos++;

    // ETX
    buffer[pos++] = PROTO_ETX;

    return pos;
}


//---------------------------------------------------------------------------
// ESP32로 데이터 전송 (전체 교체 - 응답처리 + 컴팩트 로그)
//---------------------------------------------------------------------------
void __fastcall TSCM_Ga3Agent::SendToESP32(int changeCount, bool isHeartbeat)
{
    if (!m_bCommOpened || Mycomm == NULL || !Mycomm->Active())
    {
        LogMessage("E:COM not ready");
        return;
    }

    try
    {
        // 패킷 생성
        int packetLen = BuildPacket(m_SendBuffer);

        // 수신 버퍼 비우기
        while (Mycomm->ReadBufUsed() > 0)
        {
            BYTE dummy;
            Mycomm->ReadBuf(&dummy, 1);
        }
    
		// 로그 출력 시
#if	HK_DEBUG
	    // 상세 로그 (HEX 덤프 등)
    	String hexDump = "TX: ";
	    for (int i = 0; i < packetLen; i++) hexDump += IntToHex(m_SendBuffer[i], 2) + " ";
	    LogMessage(hexDump);
#endif
        // 전송
        Mycomm->WriteBuf(m_SendBuffer, packetLen);

        // 컴팩트 로그 생성
        // 형식: D:5 TX:43 OK / D(HB):5 TX:43 OK / D:5(C:2) TX:43 FAIL
        String logMsg = "D";
        if (isHeartbeat) logMsg += "(HB)";
        logMsg += ":" + IntToStr(m_ItemCount);
        if (changeCount > 0) logMsg += "(C:" + IntToStr(changeCount) + ")";
        logMsg += " TX:" + IntToStr(packetLen);

        // 응답 대기
        if (WaitForResponse(RESP_TIMEOUT_MS))
        {
            // 성공
            logMsg += " OK";
            for (int i = 0; i < m_ItemCount; i++)
            {
                VariantCopy(&m_Items[i].varPrevValue, &m_Items[i].varValue);
                m_Items[i].Changed = false;
            }
            m_nRetryCount = 0;
        }
        else
        {
            // 실패
            logMsg += " FAIL";
//            HandleSendFailure();
        }
        
        LogMessage(logMsg);
    }
    catch (Exception &ex)
    {
        LogMessage("E:" + ex.Message);
//        HandleSendFailure();
    }
}

//---------------------------------------------------------------------------
// 서비스 시작
//---------------------------------------------------------------------------
void __fastcall TSCM_Ga3Agent::ServiceStart(TService *Sender, bool &Started)
{
    if (Timer1) Timer1->Enabled = false;

    // STA 모드로 COM 초기화 (OPC Automation은 STA 필요)
    CoInitialize(NULL);

    LogMessage("SVC START");
    Started = true;

    try
    {
        // 0. INI 설정 로드
        LoadSettings();

        // 1. CSV 설정 파일 로드
        String exePath = ExtractFilePath(ParamStr(0));
        String configFile = exePath + "oem_param.csv";

        if (!LoadItemConfig(configFile))
        {
            LogMessage("CFG: default");

            m_ItemCount = 5;
            m_Items[0].ItemID = 1; m_Items[0].TagName = "Random.Int1";  m_Items[0].DataType = "INT";
            m_Items[1].ItemID = 2; m_Items[1].TagName = "Random.Int2";  m_Items[1].DataType = "INT";
            m_Items[2].ItemID = 3; m_Items[2].TagName = "Random.Int4";  m_Items[2].DataType = "INT";
            m_Items[3].ItemID = 4; m_Items[3].TagName = "Random.Real4"; m_Items[3].DataType = "REAL";
            m_Items[4].ItemID = 5; m_Items[4].TagName = "Random.Real8"; m_Items[4].DataType = "REAL";

            for (int i = 0; i < m_ItemCount; i++)
            {
                m_Items[i].pItem = NULL;
                m_Items[i].Quality = 0;
                m_Items[i].Changed = false;
                VariantInit(&m_Items[i].varValue);
                VariantInit(&m_Items[i].varPrevValue);
            }
        }

        // 2. 시리얼 포트 초기화
        if (!InitSerialPort(m_nComPort, m_nBaudRate))
        {
            LogMessage("COM FAIL");
        }

        // 3. OPC 서버 연결
        OPCServer = CoOPCServer::Create();
#if SERVER_SIMULATE
        OPCServer->Connect(WideString("Matrikon.OPC.Simulation.1"), TNoParam());
#else
        OPCServer->Connect(WideString("Schneider-Aut.OFS.2"), TNoParam());
#endif
        LogMessage("OPC OK");

        // 4. 그룹 설정
        OPCGroups = OPCServer->OPCGroups;
        OPCGroups->DefaultGroupIsActive = true;
        OPCGroups->DefaultGroupUpdateRate = 1000;

        IOPCGroup *tempGroup = NULL;
        OPCGroups->Add(TVariant(WideString("TestGroup")), &tempGroup);
        MyGroup = tempGroup;

        MyGroup->IsActive = true;
        MyGroup->IsSubscribed = true;
        MyGroup->set_IsActive(VARIANT_TRUE);
        MyGroup->set_IsSubscribed(VARIANT_TRUE);
        MyGroup->set_UpdateRate(1000);

        MyItems = MyGroup->OPCItems;

        // 5. 아이템 등록
        int regCount = 0;
        for (int i = 0; i < m_ItemCount; i++)
        {
            OPCItem *tempItem = NULL;
            try
            {
                MyItems->AddItem(WideString(m_Items[i].TagName),
                                 m_Items[i].ItemID, &tempItem);
                m_Items[i].pItem = tempItem;

                if (tempItem != NULL)
                {
                    long serverHandle = tempItem->get_ServerHandle();
                    long clientHandle = tempItem->get_ClientHandle();
                    LogMessage("  [" + IntToStr(i) + "] SH=" + IntToStr(serverHandle) + " CH=" + IntToStr(clientHandle));
                }
                regCount++;
            }
            catch (Exception &e)
            {
                LogMessage("  [" + IntToStr(i) + "] AddItem FAIL: " + e.Message);
                m_Items[i].pItem = NULL;
            }
        }
        LogMessage("ITEM:" + IntToStr(regCount) + "/" + IntToStr(m_ItemCount));

        // 6. 초기 데이터 읽기 대기 (OPC 서버가 데이터 준비할 시간)
        Sleep(2000);

        // 7. 초기 값 읽기 테스트 (같은 스레드에서)
        for (int i = 0; i < m_ItemCount; i++)
        {
            if (m_Items[i].pItem != NULL)
            {
                OPCItem* pItem = (OPCItem*)m_Items[i].pItem;
                
                VARIANT varValue, varQuality, varTimestamp;
                VariantInit(&varValue);
                VariantInit(&varQuality);
                VariantInit(&varTimestamp);
                
                try
                {
                    HRESULT hr = pItem->Read(2, &varValue, &varQuality, &varTimestamp);
                    LogMessage("INIT RD[" + IntToStr(i) + "] hr=" + IntToHex((int)hr, 8) +
                               " vt=" + IntToStr(varValue.vt) +
                               " V=" + VariantToString(varValue));

                    if (SUCCEEDED(hr))
                    {
                        VariantCopy(&m_Items[i].varValue, &varValue);
                        VariantCopy(&m_Items[i].varPrevValue, &varValue);
                        
                        if (varQuality.vt == VT_I2)
                            m_Items[i].Quality = varQuality.iVal;
                        else if (varQuality.vt == VT_I4)
                            m_Items[i].Quality = varQuality.lVal;
                        else
                            m_Items[i].Quality = 192;
                    }
                }
                catch (Exception &e)
                {
                    LogMessage("INIT RD ERR[" + IntToStr(i) + "]: " + e.Message);
                }
                
                VariantClear(&varValue);
                VariantClear(&varQuality);
                VariantClear(&varTimestamp);
            }
        }

        m_bFirstSend = true;
        m_dwLastSendTick = 0;

        // 8. 타이머 시작
        if (Timer1)
        {
            Timer1->Interval = m_nTimeInterval;
            Timer1->Enabled = true;
		    LogMessage("Timer1 enabled, Interval=" + IntToStr(Timer1->Interval));
        }
		else
		{
		    LogMessage("ERROR: Timer1 is NULL!");
		}

        LogMessage("SVC READY");
    }
    catch (Exception &ex)
    {
        LogMessage("E:" + ex.Message);
    }
}


//---------------------------------------------------------------------------
// 서비스 종료
//---------------------------------------------------------------------------
void __fastcall TSCM_Ga3Agent::ServiceStop(TService *Sender, bool &Stopped)
{
    LogMessage("SVC STOP");

    if (Timer1) Timer1->Enabled = false;

    for (int i = 0; i < m_ItemCount; i++)
    {
        VariantClear(&m_Items[i].varValue);
        VariantClear(&m_Items[i].varPrevValue);
    }

    CloseSerialPort();

    try
    {
        if (OPCServer)
        {
            OPCServer->Disconnect();
        }
    }
    catch (Exception &ex)
    {
        LogMessage("E:" + ex.Message);
    }

    CoUninitialize();

    Stopped = true;
    LogMessage("SVC END");
}

///---------------------------------------------------------------------------
// 타이머 이벤트 (전체 교체 - Heartbeat + 컴팩트 로그 + COM 재연결)
//---------------------------------------------------------------------------
//---------------------------------------------------------------------------
// 타이머 이벤트 - SyncRead 사용 버전
//---------------------------------------------------------------------------
//---------------------------------------------------------------------------
// 타이머 이벤트 - OPCItem.Read() 사용 버전
//---------------------------------------------------------------------------
void __fastcall TSCM_Ga3Agent::Timer1Timer(TObject *Sender)
{
    Timer1->Enabled = false;

    try
    {
        //------------------------------------------------------------------
        // 1. OPC 데이터 읽기
        //------------------------------------------------------------------
        if ((IUnknown*)OPCServer != NULL && m_ItemCount > 0)
        {
            for (int i = 0; i < m_ItemCount; i++)
            {
                if (m_Items[i].pItem != NULL)
                {
                    OPCItem* pItem = (OPCItem*)m_Items[i].pItem;
                    
                    try
                    {
                        // 방법 1: Read() 사용
                        VARIANT varValue, varQuality, varTimestamp;
                        VariantInit(&varValue);
                        VariantInit(&varQuality);
                        VariantInit(&varTimestamp);
                        
                        HRESULT hr = pItem->Read(2, &varValue, &varQuality, &varTimestamp);
                        
                        if (SUCCEEDED(hr))
                        {
                            VariantCopy(&m_Items[i].varValue, &varValue);
                            
                            if (varQuality.vt == VT_I2) m_Items[i].Quality = varQuality.iVal;
                            else if (varQuality.vt == VT_I4) m_Items[i].Quality = varQuality.lVal;
                            else m_Items[i].Quality = 192;

                            // 디버그 로그 (필요시 주석 해제)
                            // LogMessage("RD[" + IntToStr(i) + "] V=" + VariantToString(varValue));
                        }
                        else
                        {
                            // Read 실패 시 get_Value 시도
                            pItem->get_Value(&m_Items[i].varValue);
                            pItem->get_Quality(&m_Items[i].Quality);

                            // 디버그 로그
                            // LogMessage("GV[" + IntToStr(i) + "] vt=" + IntToStr(m_Items[i].varValue.vt));
                        }
                        
                        VariantClear(&varValue);
                        VariantClear(&varQuality);
                        VariantClear(&varTimestamp);
                    }
                    catch (Exception &e)
                    {
                        LogMessage("RD ERR[" + IntToStr(i) + "]: " + e.Message);
                        m_Items[i].Quality = 0;
                    }
                }
            }

            //------------------------------------------------------------------
            // 2. 변경 여부 확인 및 변경 개수 카운트
            //------------------------------------------------------------------
            int changeCount = 0;
            for (int i = 0; i < m_ItemCount; i++)
            {
                if (IsValueChanged(i))
                {
                    m_Items[i].Changed = true;
                    changeCount++;
                }
            }
            bool hasChanges = (changeCount > 0);
            
            //------------------------------------------------------------------
            // 3. Heartbeat 타임아웃 확인
            //------------------------------------------------------------------
            DWORD dwNow = GetTickCount();
            bool heartbeatTimeout = false;
            
            if (m_dwLastSendTick == 0)
            {
                heartbeatTimeout = true;
            }
            else
            {
                DWORD elapsed;
                if (dwNow >= m_dwLastSendTick)
                    elapsed = dwNow - m_dwLastSendTick;
                else
                    elapsed = (0xFFFFFFFF - m_dwLastSendTick) + dwNow + 1;
                
                if (elapsed >= m_dwHeartbeatInterval)
                    heartbeatTimeout = true;
            }

            //------------------------------------------------------------------
            // 4. 전송 조건: 최초 OR 변경 OR Heartbeat
            //------------------------------------------------------------------
            if (m_bFirstSend || hasChanges || heartbeatTimeout)
            {
                bool isHB = heartbeatTimeout && !hasChanges && !m_bFirstSend;
                
                SendToESP32(changeCount, isHB);
                m_dwLastSendTick = GetTickCount();

                m_bFirstSend = false;
            }
        }
    }
    catch (Exception &e)
    {
        LogMessage("E:" + e.Message);
    }

    Timer1->Enabled = true;
}

//---------------------------------------------------------------------------
// WaitForResponse 함수 (로그 제거)
//---------------------------------------------------------------------------
bool __fastcall TSCM_Ga3Agent::WaitForResponse(int timeoutMs)
{
    if (!m_bCommOpened || Mycomm == NULL || !Mycomm->Active())
        return false;

    BYTE respBuffer[5];
    int respIndex = 0;
    DWORD startTick = GetTickCount();

    m_bWaitingResponse = true;

    while (GetTickCount() - startTick < (DWORD)timeoutMs)
    {
        if (Mycomm->ReadBufUsed() > 0)
        {
            BYTE b;
            if (Mycomm->ReadBuf(&b, 1) == 1)
            {
                // STX 감지
                if (b == PROTO_STX && respIndex == 0)
                {
                    respBuffer[respIndex++] = b;
                }
                else if (respIndex > 0 && respIndex < 5)
                {
                    respBuffer[respIndex++] = b;
                    
                    // 5바이트 수신 완료
                    if (respIndex == 5)
                    {
                        m_bWaitingResponse = false;
                        
                        // ETX 확인
                        if (respBuffer[4] != PROTO_ETX)
                            return false;
                        
                        // 체크섬 확인
                        BYTE calcChk = respBuffer[1] ^ respBuffer[2];
                        if (calcChk != respBuffer[3])
                            return false;
                        
                        // 응답 코드 확인
                        BYTE cmd = respBuffer[1];
                        BYTE status = respBuffer[2];
                        
                        return (cmd == RESP_CMD_ACK && status == RESP_STATUS_OK);
                    }
                }
            }
        }
        Sleep(10);
    }
    
    m_bWaitingResponse = false;
    return false;
}

//---------------------------------------------------------------------------
void __fastcall TSCM_Ga3Agent::HandleSendFailure()
{
    m_nRetryCount++;
    
    if (m_nRetryCount >= m_nMaxRetries)
    {
        LogMessage("Reconn...");
        
        CloseSerialPort();
        Sleep(1000);
        
        if (InitSerialPort(m_nComPort, m_nBaudRate))
        {
            LogMessage("COM OK");
            m_nRetryCount = 0;
        }
        else
        {
            LogMessage("COM FAIL");
        }
    }
}


//---------------------------------------------------------------------------
