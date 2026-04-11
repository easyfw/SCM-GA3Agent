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

    m_nRetryCount = 0;
    m_nMaxRetries = 3;
    m_bWaitingResponse = false;
    
    m_dwLastSendTick = 0;
    m_dwHeartbeatInterval = 5000;

    // No-change counter (log suppression)
    m_nNoChangeCount = 0;

    m_nComPort = 3;
    m_nBaudRate = 115200;
    m_nTimeInterval = 5000;

    // === WinCutPlus 초기화 ===
    m_sProdBasePath = "C:\\WinCutPlus\\prod\\";
    m_sCurrentProdFile = "";
    m_nLastLineCount = 0;
    m_sLastProdDate = "";
    m_bProdNewData = false;
    m_ProdRecord.FieldCount = 0;
    m_ProdRecord.RawLine = "";
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
    
    String timeStr;
    timeStr.printf("[%02d:%02d:%02d] ", st.wHour, st.wMinute, st.wSecond);

    AnsiString finalMsg = timeStr + msg + "\r\n";
    WriteFile(hFile, finalMsg.c_str(), finalMsg.Length(), &dwBytesWritten, NULL);
    CloseHandle(hFile);
}

//---------------------------------------------------------------------------
// WriteStatusFile - overwrite single-line status for no-change cycles
//---------------------------------------------------------------------------
void __fastcall TSCM_Ga3Agent::WriteStatusFile(String msg)
{
    String exePath = ExtractFilePath(ParamStr(0));
    String statusPath = exePath + "logsave_status.txt";

    HANDLE hFile = CreateFile(
        statusPath.c_str(),
        GENERIC_WRITE,
        FILE_SHARE_READ,
        NULL,
        CREATE_ALWAYS,
        FILE_ATTRIBUTE_NORMAL,
        NULL
    );
    if (hFile == INVALID_HANDLE_VALUE) return;

    SYSTEMTIME st;
    GetLocalTime(&st);

    String timeStr;
    timeStr.printf("[%04d-%02d-%02d %02d:%02d:%02d] ",
        st.wYear, st.wMonth, st.wDay,
        st.wHour, st.wMinute, st.wSecond);

    AnsiString finalMsg = timeStr + msg + "\r\n";
    DWORD dwBytesWritten;
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
        case VT_DATE: return (long)((v.date - 25569.0) * 86400.0);
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
        if (comStr.UpperCase().Pos("COM") == 1)
            m_nComPort = StrToIntDef(comStr.SubString(4, comStr.Length() - 3), 17);
        else
            m_nComPort = StrToIntDef(comStr, 17);

        m_nBaudRate = ini->ReadInteger("Communication", "BaudRate", 115200);

        // [Agent] 섹션
        m_nTimeInterval = ini->ReadInteger("Agent", "TimeInterval", 5000);

        // [WinCutPlus] 섹션
        m_sProdBasePath = ini->ReadString("WinCutPlus", "ProdPath",
                                          "C:\\WinCutPlus\\prod\\");
        if (!m_sProdBasePath.IsEmpty() &&
            m_sProdBasePath[m_sProdBasePath.Length()] != '\\')
            m_sProdBasePath += "\\";

        LogMessage("CFG: COM" + IntToStr(m_nComPort) +
                   " " + IntToStr(m_nBaudRate) +
                   " T:" + IntToStr(m_nTimeInterval) +
                   " Prod:" + m_sProdBasePath);
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
                m_Items[m_ItemCount].Description = col3;
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

        if (Mycomm->Active()) Mycomm->Close();

        Mycomm->PortNum = portNum;

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
// 오늘 날짜의 WinCutPlus 생산 파일 경로
// ※ 실제 WinCutPlus 파일 구조에 맞춰 경로 패턴 수정 필요
//---------------------------------------------------------------------------
String __fastcall TSCM_Ga3Agent::GetTodayProdFilePath()
{
    SYSTEMTIME st;
    GetLocalTime(&st);

    // WinCutPlus 실제 구조: prod\년\월\일 (앞자리 0 없음, 확장자 없음)
    // 예: C:\WinCutPlus\prod\2026\4\6
    String dateStr = IntToStr(st.wYear) + "\\" +
                     IntToStr(st.wMonth) + "\\" +
                     IntToStr(st.wDay);

    String dateKey = IntToStr(st.wYear) + IntToStr(st.wMonth) + IntToStr(st.wDay);

    // 날짜가 바뀌면 행 카운터 리셋
    if (dateKey != m_sLastProdDate)
    {
        m_sLastProdDate = dateKey;
        m_nLastLineCount = 0;
        LogMessage("PROD: new date " + dateStr);
    }

    return m_sProdBasePath + dateStr;
}

//---------------------------------------------------------------------------
// 생산 파일 신규 행 감지
//---------------------------------------------------------------------------
bool __fastcall TSCM_Ga3Agent::CheckProdFileNewLine()
{
    m_bProdNewData = false;

    String filePath = GetTodayProdFilePath();

    if (!FileExists(filePath))
        return false;

    TStringList *lines = new TStringList();
    try
    {
        // WinCutPlus가 쓰는 중에도 읽을 수 있도록 공유 모드로 열기
        TFileStream *fs = NULL;
        try
        {
            fs = new TFileStream(filePath, fmOpenRead | fmShareDenyNone);
            lines->LoadFromStream(fs);
        }
        __finally
        {
            delete fs;
        }

        int currentLineCount = lines->Count;

        // 최초: 현재 행 수만 기록
        if (m_nLastLineCount == 0 && currentLineCount > 0)
        {
            m_nLastLineCount = currentLineCount;
            m_sCurrentProdFile = filePath;
            LogMessage("PROD: init " + IntToStr(currentLineCount) + " lines");
            delete lines;
            return false;
        }

        // 신규 행 감지
        if (currentLineCount > m_nLastLineCount)
        {
            int newLineIdx = currentLineCount - 1;
            String newLine = lines->Strings[newLineIdx].Trim();

            if (!newLine.IsEmpty())
            {
                if (ParseProdLine(newLine, m_ProdRecord))
                {
                    m_bProdNewData = true;
                    LogMessage("PROD: +" + IntToStr(currentLineCount - m_nLastLineCount) +
                               " new, last=[" + newLine.SubString(1, 60) + "]");
                }
            }

            m_nLastLineCount = currentLineCount;
            m_sCurrentProdFile = filePath;

            delete lines;
            return m_bProdNewData;
        }

        m_nLastLineCount = currentLineCount;
    }
    catch (Exception &ex)
    {
        LogMessage("PROD ERR: " + ex.Message);
    }

    delete lines;
    return false;
}

//---------------------------------------------------------------------------
// 생산 행 CSV 파싱
//---------------------------------------------------------------------------
bool __fastcall TSCM_Ga3Agent::ParseProdLine(const String &line, TProdRecord &rec)
{
    rec.RawLine = line;
    rec.FieldCount = 0;

    String temp = "";
    for (int i = 1; i <= line.Length() && rec.FieldCount < MAX_PROD_FIELDS; i++)
    {
        if (line[i] == ',' || line[i] == '\t' || line[i] == ';')
        {
            rec.Fields[rec.FieldCount++] = temp.Trim();
            temp = "";
        }
        else
        {
            temp += line[i];
        }
    }
    if (!temp.IsEmpty() && rec.FieldCount < MAX_PROD_FIELDS)
    {
        rec.Fields[rec.FieldCount++] = temp.Trim();
    }

    return (rec.FieldCount > 0);
}

//---------------------------------------------------------------------------
// 패킷 생성 (Alarm only)
// [STX][LEN_L][LEN_H][MSG_TYPE][CNT][ID_L][ID_H][Q][V0-V3]...[CHK][ETX]
//---------------------------------------------------------------------------
int __fastcall TSCM_Ga3Agent::BuildPacket(BYTE* buffer, BYTE msgType)
{
    int pos = 0;

    buffer[pos++] = PROTO_STX;

    int lenPos = pos;
    pos += 2;

    // MSG_TYPE
    buffer[pos++] = msgType;

    // Item Count
    buffer[pos++] = (BYTE)m_ItemCount;

    for (int i = 0; i < m_ItemCount; i++)
    {
        WORD itemId = (WORD)m_Items[i].ItemID;
        buffer[pos++] = (BYTE)(itemId & 0xFF);
        buffer[pos++] = (BYTE)((itemId >> 8) & 0xFF);

        buffer[pos++] = (BYTE)(m_Items[i].Quality & 0xFF);

        long value = VariantToLong(m_Items[i].varValue);
        buffer[pos++] = (BYTE)(value & 0xFF);
        buffer[pos++] = (BYTE)((value >> 8) & 0xFF);
        buffer[pos++] = (BYTE)((value >> 16) & 0xFF);
        buffer[pos++] = (BYTE)((value >> 24) & 0xFF);
    }

    WORD dataLen = pos - 3;
    buffer[lenPos] = (BYTE)(dataLen & 0xFF);
    buffer[lenPos + 1] = (BYTE)((dataLen >> 8) & 0xFF);

    buffer[pos] = CalcChecksum(&buffer[1], pos - 1);
    pos++;

    buffer[pos++] = PROTO_ETX;

    return pos;
}

//---------------------------------------------------------------------------
// 복합 패킷 (Alarm + Prod)
// [STX][LEN_L][LEN_H][0x03]
// [ALARM_CNT][alarm items...]
// [PROD_FIELD_CNT][FLEN][FDATA]...[FLEN][FDATA]
// [CHK][ETX]
//---------------------------------------------------------------------------
int __fastcall TSCM_Ga3Agent::BuildCombinedPacket(BYTE* buffer, const TProdRecord &rec)
{
    int pos = 0;

    buffer[pos++] = PROTO_STX;

    int lenPos = pos;
    pos += 2;

    buffer[pos++] = MSG_TYPE_ALARM_PROD;

    // Part 1: Alarm
    buffer[pos++] = (BYTE)m_ItemCount;

    for (int i = 0; i < m_ItemCount; i++)
    {
        WORD itemId = (WORD)m_Items[i].ItemID;
        buffer[pos++] = (BYTE)(itemId & 0xFF);
        buffer[pos++] = (BYTE)((itemId >> 8) & 0xFF);

        buffer[pos++] = (BYTE)(m_Items[i].Quality & 0xFF);

        long value = VariantToLong(m_Items[i].varValue);
        buffer[pos++] = (BYTE)(value & 0xFF);
        buffer[pos++] = (BYTE)((value >> 8) & 0xFF);
        buffer[pos++] = (BYTE)((value >> 16) & 0xFF);
        buffer[pos++] = (BYTE)((value >> 24) & 0xFF);
    }

    // Part 2: Prod
    buffer[pos++] = (BYTE)rec.FieldCount;

    for (int f = 0; f < rec.FieldCount; f++)
    {
        AnsiString fieldData = AnsiString(rec.Fields[f]);
        int len = fieldData.Length();
        if (len > MAX_PROD_FIELD_LEN) len = MAX_PROD_FIELD_LEN;

        if (pos + 1 + len > 4000)
        {
            LogMessage("PROD PKT: truncated at field " + IntToStr(f));
            break;
        }

        buffer[pos++] = (BYTE)len;
        if (len > 0)
        {
            memcpy(&buffer[pos], fieldData.c_str(), len);
            pos += len;
        }
    }

    WORD dataLen = pos - 3;
    buffer[lenPos] = (BYTE)(dataLen & 0xFF);
    buffer[lenPos + 1] = (BYTE)((dataLen >> 8) & 0xFF);

    buffer[pos] = CalcChecksum(&buffer[1], pos - 1);
    pos++;

    buffer[pos++] = PROTO_ETX;

    return pos;
}

//---------------------------------------------------------------------------
// ESP32로 데이터 전송
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
        int packetLen;

        // 생산 데이터가 있으면 복합 패킷, 없으면 Alarm만
        if (m_bProdNewData)
        {
            packetLen = BuildCombinedPacket(m_SendBuffer, m_ProdRecord);
        }
        else
        {
            packetLen = BuildPacket(m_SendBuffer, MSG_TYPE_ALARM);
        }

        // 수신 버퍼 비우기
        while (Mycomm->ReadBufUsed() > 0)
        {
            BYTE dummy;
            Mycomm->ReadBuf(&dummy, 1);
        }
    
#if	HK_DEBUG
    	String hexDump = "TX: ";
	    for (int i = 0; i < packetLen; i++) hexDump += IntToHex(m_SendBuffer[i], 2) + " ";
	    LogMessage(hexDump);
#endif
        Mycomm->WriteBuf(m_SendBuffer, packetLen);

        String logMsg = "D";
        if (isHeartbeat) logMsg += "(HB)";
        if (m_bProdNewData) logMsg += "(P)";
        logMsg += ":" + IntToStr(m_ItemCount);
        if (changeCount > 0) logMsg += "(C:" + IntToStr(changeCount) + ")";
        logMsg += " TX:" + IntToStr(packetLen);

        if (WaitForResponse(RESP_TIMEOUT_MS))
        {
            logMsg += " OK";
            for (int i = 0; i < m_ItemCount; i++)
            {
                VariantCopy(&m_Items[i].varPrevValue, &m_Items[i].varValue);
                m_Items[i].Changed = false;
            }
            m_nRetryCount = 0;
            
            // 전송 성공 시 Prod 플래그 클리어
            m_bProdNewData = false;
        }
        else
        {
            logMsg += " FAIL";
        }
        
        // Log to file only when data changed (OPC or Prod) or error.
        // No-change heartbeats just update the overwrite status file.
        if (changeCount > 0 || m_bProdNewData || logMsg.Pos("FAIL") > 0)
        {
            m_nNoChangeCount = 0;
            LogMessage(logMsg);
        }
        else
        {
            m_nNoChangeCount++;
            WriteStatusFile("NoChange:" + IntToStr(m_nNoChangeCount)
                          + " " + logMsg);
        }
    }
    catch (Exception &ex)
    {
        LogMessage("E:" + ex.Message);
    }
}

//---------------------------------------------------------------------------
// 서비스 시작
//---------------------------------------------------------------------------
void __fastcall TSCM_Ga3Agent::ServiceStart(TService *Sender, bool &Started)
{
    if (Timer1) Timer1->Enabled = false;

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

        // 6. 초기 데이터 읽기 대기
        Sleep(2000);

        // 7. 초기 값 읽기
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

        // 8. WinCutPlus 생산 파일 초기화
        m_sCurrentProdFile = GetTodayProdFilePath();
        if (FileExists(m_sCurrentProdFile))
        {
            TStringList *initLines = new TStringList();
            try
            {
                TFileStream *fs = new TFileStream(m_sCurrentProdFile,
                                                  fmOpenRead | fmShareDenyNone);
                try
                {
                    initLines->LoadFromStream(fs);
                }
                __finally
                {
                    delete fs;
                }

                m_nLastLineCount = initLines->Count;
                LogMessage("PROD INIT: " + m_sCurrentProdFile +
                           " (" + IntToStr(m_nLastLineCount) + " lines)");
            }
            catch (Exception &ex)
            {
                LogMessage("PROD INIT ERR: " + ex.Message);
                m_nLastLineCount = 0;
            }
            delete initLines;
        }
        else
        {
            LogMessage("PROD: file not yet: " + m_sCurrentProdFile);
            m_nLastLineCount = 0;
        }

        m_bFirstSend = true;
        m_dwLastSendTick = 0;

        // 9. 타이머 시작
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

//---------------------------------------------------------------------------
// 타이머 이벤트 - Alarm OPC + WinCutPlus 감지 통합
//---------------------------------------------------------------------------
void __fastcall TSCM_Ga3Agent::Timer1Timer(TObject *Sender)
{
    Timer1->Enabled = false;

    try
    {
        //------------------------------------------------------------------
        // 1. OPC Alarm 데이터 읽기
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
                        }
                        else
                        {
                            pItem->get_Value(&m_Items[i].varValue);
                            pItem->get_Quality(&m_Items[i].Quality);
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
            // 2. Alarm 변경 여부 확인
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
            bool hasAlarmChanges = (changeCount > 0);

            //------------------------------------------------------------------
            // 3. WinCutPlus 생산 파일 신규 행 감지
            //------------------------------------------------------------------
            bool hasProdData = CheckProdFileNewLine();

            //------------------------------------------------------------------
            // 4. Heartbeat 타임아웃 확인
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
            // 5. 전송 조건: 최초 OR Alarm변경 OR 생산데이터 OR Heartbeat
            //------------------------------------------------------------------
            if (m_bFirstSend || hasAlarmChanges || hasProdData || heartbeatTimeout)
            {
                bool isHB = heartbeatTimeout && !hasAlarmChanges && !hasProdData && !m_bFirstSend;
                
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
// WaitForResponse
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
                if (b == PROTO_STX && respIndex == 0)
                {
                    respBuffer[respIndex++] = b;
                }
                else if (respIndex > 0 && respIndex < 5)
                {
                    respBuffer[respIndex++] = b;
                    
                    if (respIndex == 5)
                    {
                        m_bWaitingResponse = false;
                        
                        if (respBuffer[4] != PROTO_ETX)
                            return false;
                        
                        BYTE calcChk = respBuffer[1] ^ respBuffer[2];
                        if (calcChk != respBuffer[3])
                            return false;
                        
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
