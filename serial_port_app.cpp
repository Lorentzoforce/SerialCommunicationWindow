#include <windows.h>
#include <commctrl.h>
#include <string>
#include <vector>
#include <thread>
#include <mutex>
#include <chrono>
#include <sstream>
#include <windowsx.h>   // 为 GET_X_LPARAM / GET_Y_LPARAM 提供宏
#include <queue>
#include <condition_variable>
#include <atomic>

// === Speech-to-Serial dependencies [ADD] ===
#include <vector>
#include <string>
//#include <sapi.h>
//#include <sphelper.h>   // CSpEvent, SpEnumTokens helpers
#pragma warning(push)
#pragma warning(disable: 4996) // sphelper.h 内部用到了已弃用的 GetVersionExW；仅对本次包含静默
#include <sphelper.h>
#pragma warning(pop)
#pragma comment(lib, "sapi.lib")


#define IDC_COMBO_PORT 1001    // ID for the ComboBox to select serial ports
#define IDC_BUTTON_ESTABLISH 1002  // ID for the Establish Connection button
#define IDC_BUTTON_RELEASE 1003    // ID for the Release Connection button
#define IDC_BUTTON_CONNECT 1004    // ID for the Connect button
#define IDC_BUTTON_DISCONNECT 1005 // ID for the Disconnect button
#define IDC_EDIT_OUTPUT 1006   // ID for the output Edit control
#define IDC_LABEL_OUTPUT 1007  // ID for the label above output Edit
#define IDC_EDIT_INPUT 1008    // ID for the first input Edit control
#define IDC_BUTTON_SEND 1009   // ID for the first Send button
#define IDC_EDIT_INPUT2 1010   // ID for the second input Edit control
#define IDC_BUTTON_SEND2 1011  // ID for the second Send button
#define IDC_EDIT_INPUT3 1012   // ID for the third input Edit control
#define IDC_BUTTON_SEND3 1013  // ID for the third Send button
#define IDC_BUTTON_CLEAR 1014  // ID for the Clear Output button
#define IDC_LABEL_STATUS 1015  // ID for the connection status label
#define IDC_COMBO_DELAY_MODE 1016 // ID for the delay mode ComboBox
#define IDC_EDIT_DELAY_TIME 1017  // ID for the delay time Edit control
#define IDC_LABEL_DELAY_TIME 1018 // ID for the delay time label
#define IDC_EDIT_DELAY_TEXT 1019  // ID for the delay text Edit control
#define IDC_BUTTON_DELAY_SEND 1020 // ID for the delay send button
#define IDC_LABEL_REMAINING 1021   // ID for the remaining time label
#define IDC_BUTTON_DELAY_STOP 1022 // ID for the delay stop button
#define IDC_BUTTON_STOP_RECEIVING 1023 // ID for the Stop Receiving button
#define IDC_LABEL_RECEIVING_STATUS 1024 // ID for the receiving status label
#define IDC_BUTTON_REFRESH 1025 // ID for the Refresh COM ports button
#define IDC_BUTTON_W 1026      // ID for the W button
#define IDC_BUTTON_A 1027      // ID for the A button
#define IDC_BUTTON_S 1028      // ID for the S button
#define IDC_BUTTON_D 1029      // ID for the D button
#define IDC_COMBO_KEYMODE 1030 // ID for the key mode ComboBox
#define IDC_LABEL_WASD_EXAMPLE 1031 // ID for the WASD example label
#define IDC_COMBO_BAUDRATE 1032 // ID for the Baud Rate ComboBox
#define IDC_LABEL_OUTPUT_FULL 1033 // ID for the output full status label
#define WM_UPDATE_OUTPUT (WM_USER + 1)  // Custom message for updating output asynchronously
#define WM_UPDATE_REMAINING (WM_USER + 2) // Custom message for updating remaining time


// === Speech panel control IDs [ADD] ===
#define IDC_GROUP_SPEECH          3000
#define IDC_ST_SPEECH_STATUS      3001
#define IDC_CB_MIC                3002
#define IDC_BTN_REFRESH_MIC       3003
#define IDC_BTN_TOGGLE_SPEECH     3004
#define IDC_ST_LAST_SENT          3005

// SAPI 通过窗口消息回调识别事件
#define WM_RECOEVENT             (WM_APP + 201)

HINSTANCE hInst;         // Instance handle of the application
HWND hComboPort, hButtonEstablish, hButtonRelease, hButtonConnect, hButtonDisconnect, hEditOutput, hLabelOutput;
HWND hEditInput, hButtonSend, hEditInput2, hButtonSend2, hEditInput3, hButtonSend3, hButtonClear, hLabelStatus;
HWND hComboDelayMode, hEditDelayTime, hLabelDelayTime, hEditDelayText, hButtonDelaySend, hLabelRemaining, hButtonDelayStop;
HWND hButtonStopReceiving, hLabelReceivingStatus;
HWND hButtonRefresh, hButtonW, hButtonA, hButtonS, hButtonD, hComboKeyMode, hLabelWASDExample;
HWND hComboBaudRate, hLabelOutputFull; // Handle for Baud Rate ComboBox and output full status
HWND hFrameCOM, hFrameFixedInput, hFrameOutput, hFrameCustomInput, hFrameDelayInput; // Handles for static frames
HWND hComboMouseMode;    // 新的鼠标模式下拉框

HANDLE hSerial = INVALID_HANDLE_VALUE;  // Handle for the serial port
bool isRunning = false;  // Flag to indicate if serial port is active
bool isReceiving = true; // Flag to indicate if receiving data
std::mutex mtx;          // Mutex for thread-safe access to shared data
std::thread* serialThread = nullptr;  // Pointer to the serial reading thread
std::thread* delayThread = nullptr;   // Pointer to the delay thread
std::wstring outputBuffer;  // Buffer to accumulate serial output data

// === 统一的输入事件队列与发送线程 ===
std::queue<std::string> inputEventQueue;
std::mutex inputQueueMutex;
std::condition_variable inputQueueCV;
std::thread* inputSendThread = nullptr;
bool inputThreadRunning = true;

// 鼠标移动限频（全局 + 窗口焦点共用）
std::atomic<DWORD> g_lastMouseMoveTick{ 0 };
constexpr DWORD g_mouseMoveIntervalMs = 16;   // ~60Hz

// 统一队列的最大长度，防止极端场景堆积
constexpr size_t kMaxQueueLen = 1000;

std::thread* mouseSendThread = nullptr;
bool mouseThreadRunning = false;
bool keyThreadRunning = false;
bool delayRunning = false;  // Flag to indicate if delay is active
bool delayLoop = false;     // Flag to indicate loop mode
int delayTime = 0;          // Delay time in seconds
std::string delayText;      // Text to send on delay

HHOOK hKeyboardHook = NULL; // 全局键盘钩子句柄
HHOOK hMouseHook = NULL; // 全局鼠标钩子
int keyMode = 0;            // 0: Ignore, 1: Window focus, 2: Global
int mouseMode = 0;       // 0: Ignore, 1: Window focus, 2: Global



// === Speech panel handles [ADD] ===
HWND hGroupSpeech = nullptr;
HWND hSpeechStatus = nullptr;
HWND hMicCombo = nullptr;
HWND hBtnRefreshMic = nullptr;
HWND hBtnToggleSpeech = nullptr;
HWND hLastSentLabel = nullptr;

// === Speech recognition state [ADD] ===
bool g_comInited = false;           // 是否已 CoInitialize
bool g_speechEnabled = false;       // 功能是否开启
std::vector<std::wstring> g_micTokenIds; // 与下拉框索引一一对应：音频输入 tokenId

// SAPI COM 接口指针（使用原生指针，注意释放）
ISpRecognizer* g_cpRecognizer = nullptr;
ISpRecoContext* g_cpRecoContext = nullptr;
ISpRecoGrammar* g_cpDictGrammar = nullptr;
ISpObjectToken* g_cpAudioToken = nullptr;
ISpAudio* g_cpAudioIn = nullptr;

// Function to enumerate available serial ports
std::vector<std::wstring> EnumerateSerialPorts() {
    std::vector<std::wstring> ports;
    for (int i = 1; i <= 256; ++i) {
        std::wstring portName = L"\\\\.\\COM" + std::to_wstring(i);
        HANDLE hPort = CreateFileW(portName.c_str(), GENERIC_READ | GENERIC_WRITE, 0, NULL, OPEN_EXISTING, 0, NULL);
        if (hPort != INVALID_HANDLE_VALUE) {
            ports.push_back(L"COM" + std::to_wstring(i));
            CloseHandle(hPort);
        }
    }
    return ports;
}
// Function to send data to the serial port
bool SendToSerial(const std::string& data, int count = 1) {
    if (hSerial == INVALID_HANDLE_VALUE) {
        MessageBoxW(NULL, L"Serial port not opened", L"Error", MB_OK | MB_ICONERROR);
        return false;
    }

    std::string dataWithSemicolon = data + ";";
    DWORD bytesWritten;
    for (int i = 0; i < count && i < 3; ++i) {
        if (!WriteFile(hSerial, dataWithSemicolon.c_str(), dataWithSemicolon.length(), &bytesWritten, NULL)) {
            MessageBoxW(NULL, L"Failed to write to serial port", L"Error", MB_OK | MB_ICONERROR);
            return false;
        }
        //Sleep(100);
    }
    return true;
}

//Voice control*********************************************************************
// 将宽字串粗略转成 ANSI（非 ASCII 字符用 '?' 替代）
static std::string Narrow(const std::wstring& w) {
    std::string s; s.reserve(w.size());
    for (wchar_t c : w) s.push_back((c < 128) ? char(c) : '?');
    return s;
}

// 更新面板 UI：状态文本与按钮文字
static void UpdateSpeechUI() {
    if (hSpeechStatus) {
        SetWindowTextW(hSpeechStatus, g_speechEnabled ? L"Status: Enabled" : L"Status: Disabled");
    }
    if (hBtnToggleSpeech) {
        SetWindowTextW(hBtnToggleSpeech, g_speechEnabled ? L"Stop" : L"Start");
    }
}
//Voice control main logic


// 重新枚举录音设备（SAPI Audio-In tokens），填充下拉框
static void PopulateMicList() {
    if (!hMicCombo) return;
    SendMessageW(hMicCombo, CB_RESETCONTENT, 0, 0);
    g_micTokenIds.clear();

    IEnumSpObjectTokens* pEnum = nullptr;
    ULONG count = 0;
    if (SUCCEEDED(SpEnumTokens(SPCAT_AUDIOIN, nullptr, nullptr, &pEnum)) &&
        SUCCEEDED(pEnum->GetCount(&count))) {
        for (ULONG i = 0; i < count; ++i) {
            ISpObjectToken* pTok = nullptr;
            ULONG fetched = 0;
            if (pEnum->Next(1, &pTok, &fetched) == S_OK && pTok) {
                // 友好名
                LPWSTR desc = nullptr;
                SpGetDescription(pTok, &desc);
                // tokenId
                LPWSTR id = nullptr;
                pTok->GetId(&id);

                if (desc && id) {
                    SendMessageW(hMicCombo, CB_ADDSTRING, 0, (LPARAM)desc);
                    g_micTokenIds.emplace_back(id);
                }

                if (desc) ::CoTaskMemFree(desc);
                if (id)   ::CoTaskMemFree(id);
                pTok->Release();
            }
        }
    }
    if (pEnum) pEnum->Release();

    if (!g_micTokenIds.empty()) SendMessageW(hMicCombo, CB_SETCURSEL, 0, 0);
}

// 停止语音识别并释放资源
static void StopSpeech(HWND hwnd) {
    if (g_cpDictGrammar) { g_cpDictGrammar->SetDictationState(SPRS_INACTIVE); }
    if (g_cpDictGrammar) { g_cpDictGrammar->Release(); g_cpDictGrammar = nullptr; }
    if (g_cpRecoContext) { g_cpRecoContext->Release(); g_cpRecoContext = nullptr; }
    if (g_cpRecognizer) { g_cpRecognizer->SetInput(nullptr, TRUE); g_cpRecognizer->Release(); g_cpRecognizer = nullptr; }
    if (g_cpAudioIn) { g_cpAudioIn->Release(); g_cpAudioIn = nullptr; }
    if (g_cpAudioToken) { g_cpAudioToken->Release(); g_cpAudioToken = nullptr; }

    g_speechEnabled = false;
    UpdateSpeechUI();
    if (hLastSentLabel) SetWindowTextW(hLastSentLabel, L"Last sent: (none)");
}

// 启动语音识别（使用下拉框的当前设备）
static bool StartSpeech(HWND hwnd) {
    if (g_speechEnabled) return true;

    if (!g_comInited) {
        if (FAILED(CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED))) return false;
        g_comInited = true;
    }

    // 读取当前选择的设备
    LRESULT idx = SendMessageW(hMicCombo, CB_GETCURSEL, 0, 0);
    if (idx == CB_ERR || idx < 0 || (size_t)idx >= g_micTokenIds.size()) {
        MessageBoxW(hwnd, L"Please choose a microphone first.", L"Speech", MB_OK | MB_ICONWARNING);
        return false;
    }

    // 创建识别器（进程内）
    if (FAILED(CoCreateInstance(CLSID_SpInprocRecognizer, nullptr, CLSCTX_ALL,
        IID_ISpRecognizer, (void**)&g_cpRecognizer))) {
        MessageBoxW(hwnd, L"Failed to create recognizer.", L"Speech", MB_OK | MB_ICONERROR);
        return false;
    }

    //强制选择英文
    ISpObjectToken* pRecognizerToken = nullptr;
    //IEnumSpObjectTokens* pEnum = nullptr;
    //if (SUCCEEDED(SpEnumTokens(SPCAT_RECOGNIZERS, nullptr, nullptr, &pEnum))) {
    //    ULONG count = 0;
    //    if (SUCCEEDED(pEnum->GetCount(&count))) {
    //        for (ULONG i = 0; i < count; ++i) {
    //            ISpObjectToken* pTok = nullptr;
    //            if (pEnum->Next(1, &pTok, nullptr) == S_OK && pTok) {
    //                LPWSTR lang = nullptr;
    //                pTok->GetStringValue(L"Language", &lang);
    //                if (lang && wcsstr(lang, L"409")) { // 409 = en-US
    //                    pRecognizerToken = pTok;
    //                    break; // 找到英语识别器
    //                }
    //                if (lang) CoTaskMemFree(lang);
    //                pTok->Release();
    //            }
    //        }
    //    }
    //    pEnum->Release();
    //}
    //if (pRecognizerToken) {
    //    g_cpRecognizer->SetRecognizer(pRecognizerToken);
    //    pRecognizerToken->Release();
    //}
    if (SUCCEEDED(SpFindBestToken(SPCAT_RECOGNIZERS, L"Language=409", nullptr, &pRecognizerToken))) {
        g_cpRecognizer->SetRecognizer(nullptr); // 清空旧配置
        g_cpRecognizer->SetRecognizer(pRecognizerToken);
        pRecognizerToken->Release();
    }

    // 用所选 token 创建音频输入对象并绑定到识别器
    if (FAILED(SpGetTokenFromId(g_micTokenIds[(size_t)idx].c_str(), &g_cpAudioToken))) {
        MessageBoxW(hwnd, L"Failed to get selected mic token.", L"Speech", MB_OK | MB_ICONERROR);
        StopSpeech(hwnd);
        return false;
    }
    if (FAILED(SpCreateObjectFromToken(g_cpAudioToken, &g_cpAudioIn))) {
        MessageBoxW(hwnd, L"Failed to create audio object from token.", L"Speech", MB_OK | MB_ICONERROR);
        StopSpeech(hwnd);
        return false;
    }
    if (FAILED(g_cpRecognizer->SetInput(g_cpAudioIn, TRUE))) {
        MessageBoxW(hwnd, L"Failed to set recognizer input.", L"Speech", MB_OK | MB_ICONERROR);
        StopSpeech(hwnd);
        return false;
    }

    // 创建识别上下文，改为窗口消息通知
    if (FAILED(g_cpRecognizer->CreateRecoContext(&g_cpRecoContext))) {
        MessageBoxW(hwnd, L"Failed to create reco context.", L"Speech", MB_OK | MB_ICONERROR);
        StopSpeech(hwnd);
        return false;
    }
    if (FAILED(g_cpRecoContext->SetNotifyWindowMessage(hwnd, WM_RECOEVENT, 0, 0))) {
        MessageBoxW(hwnd, L"Failed to bind reco notifications.", L"Speech", MB_OK | MB_ICONERROR);
        StopSpeech(hwnd);
        return false;
    }
    // 仅关心识别结果事件（也可加上 SPEI_FALSE_RECOGNITION 根据需要）
    g_cpRecoContext->SetInterest(SPFEI(SPEI_RECOGNITION), SPFEI(SPEI_RECOGNITION));

    // 加载并激活“听写”模式
    if (FAILED(g_cpRecoContext->CreateGrammar(1, &g_cpDictGrammar))) {
        MessageBoxW(hwnd, L"Failed to create grammar.", L"Speech", MB_OK | MB_ICONERROR);
        StopSpeech(hwnd);
        return false;
    }
    if (FAILED(g_cpDictGrammar->LoadDictation(nullptr, SPLO_STATIC))) {
        MessageBoxW(hwnd, L"Failed to load dictation.", L"Speech", MB_OK | MB_ICONERROR);
        StopSpeech(hwnd);
        return false;
    }
    if (FAILED(g_cpDictGrammar->SetDictationState(SPRS_ACTIVE))) {
        MessageBoxW(hwnd, L"Failed to activate dictation.", L"Speech", MB_OK | MB_ICONERROR);
        StopSpeech(hwnd);
        return false;
    }

    g_speechEnabled = true;
    UpdateSpeechUI();
    return true;
}


void InputSendThread() {
    while (inputThreadRunning) {
        std::unique_lock<std::mutex> lock(inputQueueMutex);
        inputQueueCV.wait(lock, [] {
            return !inputEventQueue.empty() || !inputThreadRunning;
            });

        while (!inputEventQueue.empty()) {
            std::string msg = std::move(inputEventQueue.front());
            inputEventQueue.pop();
            lock.unlock();  // 立刻放开锁，允许钩子/窗口线程继续入队

            if (isRunning && hSerial != INVALID_HANDLE_VALUE) {
                // 统一通过 SendToSerial，加分号的事交给它做
                SendToSerial(msg);
            }

            lock.lock();    // 继续批量取
        }
    }
}
//Voice control*********************************************************************

// Function to update ComboBox with available serial ports
void UpdateSerialPorts(HWND hwnd) {
    SendMessageW(hComboPort, CB_RESETCONTENT, 0, 0); // Clear existing items
    auto ports = EnumerateSerialPorts();
    for (const auto& port : ports) {
        SendMessageW(hComboPort, CB_ADDSTRING, 0, (LPARAM)port.c_str());
    }
    if (!ports.empty()) {
        SendMessageW(hComboPort, CB_SETCURSEL, 0, 0);
    }
}

// Function to initialize and configure the serial port
bool InitSerialPort(const std::wstring& portName, HWND hwnd) {
    hSerial = CreateFileW((L"\\\\.\\" + portName).c_str(), GENERIC_READ | GENERIC_WRITE, 0, NULL, OPEN_EXISTING, 0, NULL);
    if (hSerial == INVALID_HANDLE_VALUE) {
        MessageBoxW(NULL, L"Failed to open serial port", L"Error", MB_OK | MB_ICONERROR);
        return false;
    }

    DCB dcb = { 0 };
    dcb.DCBlength = sizeof(DCB);
    GetCommState(hSerial, &dcb);

    int sel = SendMessageW(hComboBaudRate, CB_GETCURSEL, 0, 0);
    DWORD baudRates[] = { CBR_300, CBR_1200, CBR_2400, CBR_4800, CBR_9600, CBR_19200, CBR_38400, CBR_57600, CBR_115200 };
    dcb.BaudRate = (sel != CB_ERR && sel < 9) ? baudRates[sel] : CBR_57600;
    dcb.ByteSize = 8;
    dcb.Parity = NOPARITY;
    dcb.StopBits = ONESTOPBIT;
    if (!SetCommState(hSerial, &dcb)) {
        MessageBoxW(NULL, L"Failed to set serial port state", L"Error", MB_OK | MB_ICONERROR);
        CloseHandle(hSerial);
        hSerial = INVALID_HANDLE_VALUE;
        return false;
    }

    COMMTIMEOUTS timeouts = { 0 };
    timeouts.ReadIntervalTimeout = MAXDWORD;
    timeouts.ReadTotalTimeoutConstant = 0;
    timeouts.ReadTotalTimeoutMultiplier = 0;
    if (!SetCommTimeouts(hSerial, &timeouts)) {
        MessageBoxW(NULL, L"Failed to set serial port timeouts", L"Error", MB_OK | MB_ICONERROR);
        CloseHandle(hSerial);
        hSerial = INVALID_HANDLE_VALUE;
        return false;
    }
    return true;
}

// Thread function to read data from serial port
void SerialReadThread(HWND hwnd) {
    char buffer[256];
    DWORD bytesRead;

    while (isRunning) {
        if (isReceiving) {
            if (ReadFile(hSerial, buffer, sizeof(buffer) - 1, &bytesRead, NULL) && bytesRead > 0) {
                buffer[bytesRead] = '\0';
                std::wstring wBuffer(buffer, buffer + bytesRead);
                {
                    std::lock_guard<std::mutex> lock(mtx); // mtx只保护outputBuffer
                    outputBuffer += wBuffer;
                }
                PostMessageW(hwnd, WM_UPDATE_OUTPUT, 0, 0);  // 立刻刷新
            }
        }
        else {
            Sleep(10);
        }
    }
}

// Thread function for delay sending
void DelaySendThread(HWND hwnd) {
    int remaining = delayTime;
    while (delayRunning && remaining > 0) {
        Sleep(1000);
        remaining--;
        PostMessageW(hwnd, WM_UPDATE_REMAINING, remaining, 0);
    }
    if (delayRunning && hSerial != INVALID_HANDLE_VALUE) {
        SendToSerial(delayText);
    }
    if (!delayLoop) {
        delayRunning = false;
        PostMessageW(hwnd, WM_UPDATE_REMAINING, 0, 0);
    }
}

// 键盘钩子回调函数
LRESULT CALLBACK KeyboardProc(int nCode, WPARAM wParam, LPARAM lParam) {
    if (nCode >= 0 && wParam == WM_KEYDOWN) {
        if (isRunning && hSerial != INVALID_HANDLE_VALUE && keyMode == 2) { // 全局模式
            const KBDLLHOOKSTRUCT* kb = reinterpret_cast<KBDLLHOOKSTRUCT*>(lParam);
            wchar_t keyName[32];
            UINT scanCode = MapVirtualKey(kb->vkCode, MAPVK_VK_TO_VSC);
            if (GetKeyNameTextW(scanCode << 16, keyName, 32) > 0) {
                std::wstring wsKey(keyName);
                std::string keyStr(wsKey.begin(), wsKey.end());
                std::lock_guard<std::mutex> lk(inputQueueMutex);
                if (inputEventQueue.size() < kMaxQueueLen) {
                    inputEventQueue.push(".KEY." + keyStr);
                    inputQueueCV.notify_one();
                }
            }
        }
    }
    return CallNextHookEx(hKeyboardHook, nCode, wParam, lParam);
}
// 鼠标钩子回调
LRESULT CALLBACK MouseProc(int nCode, WPARAM wParam, LPARAM lParam) {
    if (nCode >= 0 && isRunning && hSerial != INVALID_HANDLE_VALUE && mouseMode == 2) {
        const MSLLHOOKSTRUCT* ms = reinterpret_cast<MSLLHOOKSTRUCT*>(lParam);
        std::string msg;
        DWORD now = GetTickCount();

        switch (wParam) {
        case WM_LBUTTONDOWN: msg = ".MOUSE.LEFT_DOWN"; break;
        case WM_LBUTTONUP:   msg = ".MOUSE.LEFT_UP";   break;
        case WM_RBUTTONDOWN: msg = ".MOUSE.RIGHT_DOWN";break;
        case WM_RBUTTONUP:   msg = ".MOUSE.RIGHT_UP";  break;
        case WM_MBUTTONDOWN: msg = ".MOUSE.MIDDLE_DOWN"; break;
        case WM_MBUTTONUP:   msg = ".MOUSE.MIDDLE_UP";   break;
        case WM_MOUSEWHEEL:
            msg = (GET_WHEEL_DELTA_WPARAM(ms->mouseData) > 0) ? ".MOUSE.WHEEL_UP" : ".MOUSE.WHEEL_DOWN";
            break;
        case WM_MOUSEMOVE:
            if (now - g_lastMouseMoveTick.load() >= g_mouseMoveIntervalMs) {
                std::ostringstream oss;
                oss << ".MOUSE.MOVE:" << ms->pt.x << "," << ms->pt.y;
                msg = oss.str();
                g_lastMouseMoveTick.store(now);
            }
            break;
        }

        if (!msg.empty()) {
            std::lock_guard<std::mutex> lk(inputQueueMutex);
            if (inputEventQueue.size() < kMaxQueueLen) {
                inputEventQueue.push(std::move(msg));
                inputQueueCV.notify_one();
            }
        }
    }
    return CallNextHookEx(hMouseHook, nCode, wParam, lParam);
}


// Window procedure
LRESULT CALLBACK WndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    switch (msg) {
    case WM_CREATE: {
        INITCOMMONCONTROLSEX icex = { sizeof(INITCOMMONCONTROLSEX), ICC_STANDARD_CLASSES };
        InitCommonControlsEx(&icex);

        // COM and Baud Rate Frame
        hFrameCOM = CreateWindowW(L"STATIC", L"", WS_CHILD | WS_VISIBLE | SS_ETCHEDFRAME,
            20, 0, 400, 355, hwnd, NULL, hInst, NULL);
        CreateWindowW(L"STATIC", L"COM Settings", WS_CHILD | WS_VISIBLE | SS_LEFT,
            30, 10, 100, 20, hwnd, NULL, hInst, NULL);
        CreateWindowW(L"STATIC", L"Select COM Port", WS_CHILD | WS_VISIBLE | SS_LEFT,
            30, 40, 150, 20, hwnd, NULL, hInst, NULL);
        hComboPort = CreateWindowW(L"COMBOBOX", NULL, WS_CHILD | WS_VISIBLE | CBS_DROPDOWNLIST | WS_VSCROLL,
            30, 60, 120, 150, hwnd, (HMENU)IDC_COMBO_PORT, hInst, NULL);
        hButtonRefresh = CreateWindowW(L"BUTTON", L"Refresh", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
            160, 60, 70, 20, hwnd, (HMENU)IDC_BUTTON_REFRESH, hInst, NULL);
        hButtonEstablish = CreateWindowW(L"BUTTON", L"Establish Connection", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
            30, 90, 330, 30, hwnd, (HMENU)IDC_BUTTON_ESTABLISH, hInst, NULL);
        hButtonRelease = CreateWindowW(L"BUTTON", L"Release Connection", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
            30, 130, 330, 30, hwnd, (HMENU)IDC_BUTTON_RELEASE, hInst, NULL);
        hLabelStatus = CreateWindowW(L"STATIC", L"Status: Disconnected", WS_CHILD | WS_VISIBLE | SS_LEFT,
            30, 170, 330, 20, hwnd, (HMENU)IDC_LABEL_STATUS, hInst, NULL);
        CreateWindowW(L"STATIC", L"Select Baud Rate", WS_CHILD | WS_VISIBLE | SS_LEFT,
            30, 200, 150, 20, hwnd, NULL, hInst, NULL);
        hComboBaudRate = CreateWindowW(L"COMBOBOX", NULL, WS_CHILD | WS_VISIBLE | CBS_DROPDOWNLIST | WS_VSCROLL,
            30, 220, 120, 200, hwnd, (HMENU)IDC_COMBO_BAUDRATE, hInst, NULL);
        const wchar_t* baudRates[] = { L"300", L"1200", L"2400", L"4800", L"9600", L"19200",
                                       L"38400", L"57600", L"115200" };
        for (int i = 0; i < 9; i++) {
            SendMessageW(hComboBaudRate, CB_ADDSTRING, 0, (LPARAM)baudRates[i]);
        }
        SendMessageW(hComboBaudRate, CB_SETCURSEL, 7, 0); // Default to 57600

        // Output Frame
        hFrameOutput = CreateWindowW(L"STATIC", L"", WS_CHILD | WS_VISIBLE | SS_ETCHEDFRAME,
            420, 0, 510, 355, hwnd, NULL, hInst, NULL);
        CreateWindowW(L"STATIC", L"Output Display", WS_CHILD | WS_VISIBLE | SS_LEFT,
            430, 10, 150, 20, hwnd, NULL, hInst, NULL);
        hLabelOutput = CreateWindowW(L"STATIC", L"Arduino Serial Output", WS_CHILD | WS_VISIBLE | SS_LEFT,
            430, 40, 450, 20, hwnd, (HMENU)IDC_LABEL_OUTPUT, hInst, NULL);
        hEditOutput = CreateWindowW(L"EDIT", NULL, WS_CHILD | WS_VISIBLE | WS_VSCROLL | ES_MULTILINE | ES_AUTOVSCROLL,
            430, 60, 450, 205, hwnd, (HMENU)IDC_EDIT_OUTPUT, hInst, NULL);
        hLabelReceivingStatus = CreateWindowW(L"STATIC", L"Receiving: Yes", WS_CHILD | WS_VISIBLE | SS_LEFT,
            430, 275, 200, 20, hwnd, (HMENU)IDC_LABEL_RECEIVING_STATUS, hInst, NULL);
        hLabelOutputFull = CreateWindowW(L"STATIC", L"Output not full", WS_CHILD | WS_VISIBLE | SS_LEFT,
            430, 305, 450, 20, hwnd, (HMENU)IDC_LABEL_OUTPUT_FULL, hInst, NULL);
        hButtonClear = CreateWindowW(L"BUTTON", L"Clear Output", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
            430, 325, 90, 30, hwnd, (HMENU)IDC_BUTTON_CLEAR, hInst, NULL);
        hButtonStopReceiving = CreateWindowW(L"BUTTON", L"Start/Stop", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
            730, 325, 100, 30, hwnd, (HMENU)IDC_BUTTON_STOP_RECEIVING, hInst, NULL);

        // Fixed Input Frame
        hFrameFixedInput = CreateWindowW(L"STATIC", L"", WS_CHILD | WS_VISIBLE | SS_ETCHEDFRAME,
            20, 370, 400, 200, hwnd, NULL, hInst, NULL);
        CreateWindowW(L"STATIC", L"Fixed Input Controls", WS_CHILD | WS_VISIBLE | SS_LEFT,
            30, 380, 150, 20, hwnd, NULL, hInst, NULL);
        hButtonConnect = CreateWindowW(L"BUTTON", L"Start", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
            30, 410, 80, 30, hwnd, (HMENU)IDC_BUTTON_CONNECT, hInst, NULL);
        hButtonDisconnect = CreateWindowW(L"BUTTON", L"Stop", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
            120, 410, 80, 30, hwnd, (HMENU)IDC_BUTTON_DISCONNECT, hInst, NULL);
        CreateWindowW(L"STATIC", L"WASD Controls", WS_CHILD | WS_VISIBLE | SS_LEFT,
            30, 450, 100, 20, hwnd, NULL, hInst, NULL);
        hButtonW = CreateWindowW(L"BUTTON", L"W", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
            90, 505, 40, 30, hwnd, (HMENU)IDC_BUTTON_W, hInst, NULL);
        hButtonA = CreateWindowW(L"BUTTON", L"A", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
            50, 535, 40, 30, hwnd, (HMENU)IDC_BUTTON_A, hInst, NULL);
        hButtonS = CreateWindowW(L"BUTTON", L"S", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
            90, 535, 40, 30, hwnd, (HMENU)IDC_BUTTON_S, hInst, NULL);
        hButtonD = CreateWindowW(L"BUTTON", L"D", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
            130, 535, 40, 30, hwnd, (HMENU)IDC_BUTTON_D, hInst, NULL);
        hComboKeyMode = CreateWindowW(L"COMBOBOX", NULL, WS_CHILD | WS_VISIBLE | CBS_DROPDOWNLIST | WS_VSCROLL,
            135, 445, 280, 100, hwnd, (HMENU)IDC_COMBO_KEYMODE, hInst, NULL);
        SendMessageW(hComboKeyMode, CB_ADDSTRING, 0, (LPARAM)L"Ignore all keyboard input");
        SendMessageW(hComboKeyMode, CB_ADDSTRING, 0, (LPARAM)L"Enable Keyboard Input(Window focus only)");
        SendMessageW(hComboKeyMode, CB_ADDSTRING, 0, (LPARAM)L"Enable Keyboard Input(Global, background)");
        SendMessageW(hComboKeyMode, CB_SETCURSEL, 0, 0); // Default to Ignore
        hLabelWASDExample = CreateWindowW(L"STATIC", L"e.g., W sends: .KEY.W;", WS_CHILD | WS_VISIBLE | SS_LEFT,
            180, 510, 200, 30, hwnd, (HMENU)IDC_LABEL_WASD_EXAMPLE, hInst, NULL);
        CreateWindowW(L"STATIC", L"Mouse Controls Mode", WS_CHILD | WS_VISIBLE | SS_LEFT,
            30, 480, 150, 20, hwnd, NULL, hInst, NULL); // 标签

        hComboMouseMode = CreateWindowW(L"COMBOBOX", NULL,
            WS_CHILD | WS_VISIBLE | CBS_DROPDOWNLIST | WS_VSCROLL,
            135, 475, 250, 100, hwnd, (HMENU)2001, hInst, NULL); // ID=2001，避免和其他冲突
        SendMessageW(hComboMouseMode, CB_ADDSTRING, 0, (LPARAM)L"Ignore all mouse input");
        SendMessageW(hComboMouseMode, CB_ADDSTRING, 0, (LPARAM)L"Enable mouse tracking (Window focus only)");
        SendMessageW(hComboMouseMode, CB_ADDSTRING, 0, (LPARAM)L"Enable mouse tracking (Global, background)");
        SendMessageW(hComboMouseMode, CB_SETCURSEL, 0, 0); // 默认忽略


        // Custom Input Frame
        hFrameCustomInput = CreateWindowW(L"STATIC", L"", WS_CHILD | WS_VISIBLE | SS_ETCHEDFRAME,
            20, 580, 400, 130, hwnd, NULL, hInst, NULL);
        CreateWindowW(L"STATIC", L"Custom Text Input", WS_CHILD | WS_VISIBLE | SS_LEFT,
            30, 590, 150, 20, hwnd, NULL, hInst, NULL);
        hEditInput = CreateWindowW(L"EDIT", NULL, WS_CHILD | WS_VISIBLE | WS_BORDER | ES_AUTOHSCROLL,
            30, 620, 150, 20, hwnd, (HMENU)IDC_EDIT_INPUT, hInst, NULL);
        hButtonSend = CreateWindowW(L"BUTTON", L"Send", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
            190, 620, 40, 20, hwnd, (HMENU)IDC_BUTTON_SEND, hInst, NULL);
        hEditInput2 = CreateWindowW(L"EDIT", NULL, WS_CHILD | WS_VISIBLE | WS_BORDER | ES_AUTOHSCROLL,
            30, 650, 150, 20, hwnd, (HMENU)IDC_EDIT_INPUT2, hInst, NULL);
        hButtonSend2 = CreateWindowW(L"BUTTON", L"Send", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
            190, 650, 40, 20, hwnd, (HMENU)IDC_BUTTON_SEND2, hInst, NULL);
        hEditInput3 = CreateWindowW(L"EDIT", NULL, WS_CHILD | WS_VISIBLE | WS_BORDER | ES_AUTOHSCROLL,
            30, 680, 150, 20, hwnd, (HMENU)IDC_EDIT_INPUT3, hInst, NULL);
        hButtonSend3 = CreateWindowW(L"BUTTON", L"Send", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
            190, 680, 40, 20, hwnd, (HMENU)IDC_BUTTON_SEND3, hInst, NULL);

        // Delay Input Frame
        hFrameDelayInput = CreateWindowW(L"STATIC", L"", WS_CHILD | WS_VISIBLE | SS_ETCHEDFRAME,
            420, 580, 510, 180, hwnd, NULL, hInst, NULL);
        CreateWindowW(L"STATIC", L"Delay Input Controls", WS_CHILD | WS_VISIBLE | SS_LEFT,
            430, 590, 150, 20, hwnd, NULL, hInst, NULL);
        CreateWindowW(L"STATIC", L"Delay Mode", WS_CHILD | WS_VISIBLE | SS_LEFT,
            430, 620, 100, 20, hwnd, NULL, hInst, NULL);
        hComboDelayMode = CreateWindowW(L"COMBOBOX", NULL, WS_CHILD | WS_VISIBLE | CBS_DROPDOWNLIST | WS_VSCROLL,
            530, 620, 100, 100, hwnd, (HMENU)IDC_COMBO_DELAY_MODE, hInst, NULL);
        SendMessageW(hComboDelayMode, CB_ADDSTRING, 0, (LPARAM)L"Single");
        SendMessageW(hComboDelayMode, CB_ADDSTRING, 0, (LPARAM)L"Loop");
        SendMessageW(hComboDelayMode, CB_SETCURSEL, 0, 0);
        hLabelDelayTime = CreateWindowW(L"STATIC", L"Delay Time (s)", WS_CHILD | WS_VISIBLE | SS_LEFT,
            430, 650, 100, 20, hwnd, (HMENU)IDC_LABEL_DELAY_TIME, hInst, NULL);
        hEditDelayTime = CreateWindowW(L"EDIT", NULL, WS_CHILD | WS_VISIBLE | WS_BORDER | ES_AUTOHSCROLL,
            530, 650, 100, 20, hwnd, (HMENU)IDC_EDIT_DELAY_TIME, hInst, NULL);
        CreateWindowW(L"STATIC", L"Delay Text", WS_CHILD | WS_VISIBLE | SS_LEFT,
            430, 680, 100, 20, hwnd, NULL, hInst, NULL);
        hEditDelayText = CreateWindowW(L"EDIT", NULL, WS_CHILD | WS_VISIBLE | WS_BORDER | ES_AUTOHSCROLL,
            530, 680, 100, 20, hwnd, (HMENU)IDC_EDIT_DELAY_TEXT, hInst, NULL);
        hButtonDelaySend = CreateWindowW(L"BUTTON", L"Delay Send", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
            430, 710, 100, 30, hwnd, (HMENU)IDC_BUTTON_DELAY_SEND, hInst, NULL);
        hLabelRemaining = CreateWindowW(L"STATIC", L"Remaining: 0s", WS_CHILD | WS_VISIBLE | SS_LEFT,
            540, 715, 100, 20, hwnd, (HMENU)IDC_LABEL_REMAINING, hInst, NULL);
        hButtonDelayStop = CreateWindowW(L"BUTTON", L"Stop", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
            540, 710, 100, 30, hwnd, (HMENU)IDC_BUTTON_DELAY_STOP, hInst, NULL);

        UpdateSerialPorts(hwnd);
        //voice control part*************************************************************
        // 初始化 COM（用于设备枚举；启动/停止里也会做兜底）
        if (!g_comInited && SUCCEEDED(CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED))) {
            g_comInited = true;
        }

        // === Speech panel (右侧面板) [ADD] ===
        // === Speech panel (右侧面板) [REPLACE EXISTING SPEECH PANEL CREATION] ===
// 与 Fixed Input Controls 框同一竖直高度 (y = 370)

        hGroupSpeech = CreateWindowExW(
            0, L"BUTTON", L"Speech to Serial",
            WS_CHILD | WS_VISIBLE | BS_GROUPBOX,
            440, 370, 450, 200, hwnd, (HMENU)IDC_GROUP_SPEECH, hInst, nullptr);

        hSpeechStatus = CreateWindowExW(
            0, L"STATIC", L"Status: Disabled",
            WS_CHILD | WS_VISIBLE,
            440 + 12, 370 + 24, 450 - 24, 20,
            hwnd, (HMENU)IDC_ST_SPEECH_STATUS, hInst, nullptr);

        hMicCombo = CreateWindowExW(
            WS_EX_CLIENTEDGE, L"COMBOBOX", L"",
            WS_CHILD | WS_VISIBLE | CBS_DROPDOWNLIST,
            440 + 12, 370 + 52, 450 - 120, 250,
            hwnd, (HMENU)IDC_CB_MIC, hInst, nullptr);

        hBtnRefreshMic = CreateWindowExW(
            0, L"BUTTON", L"Refresh",
            WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
            440 + 450 - 100, 370 + 52, 88, 24,
            hwnd, (HMENU)IDC_BTN_REFRESH_MIC, hInst, nullptr);

        hBtnToggleSpeech = CreateWindowExW(
            0, L"BUTTON", L"Start",
            WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
            440 + 12, 370 + 90, 100, 28,
            hwnd, (HMENU)IDC_BTN_TOGGLE_SPEECH, hInst, nullptr);

        hLastSentLabel = CreateWindowExW(
            0, L"STATIC", L"Last sent: (none)",
            WS_CHILD | WS_VISIBLE,
            440 + 12, 370 + 132, 450 - 24, 20,
            hwnd, (HMENU)IDC_ST_LAST_SENT, hInst, nullptr);

        // 填充设备列表
        PopulateMicList();
        UpdateSpeechUI();

        // 启动统一输入发送线程（伴随程序全生命周期）
        inputThreadRunning = true;
        inputSendThread = new std::thread(InputSendThread);
        break;
    }
    case WM_COMMAND: {
        if (LOWORD(wParam) == IDC_BUTTON_ESTABLISH && HIWORD(wParam) == BN_CLICKED) {
            if (!isRunning) {
                wchar_t port[10];
                int sel = SendMessageW(hComboPort, CB_GETCURSEL, 0, 0);
                if (sel != CB_ERR) {
                    SendMessageW(hComboPort, CB_GETLBTEXT, sel, (LPARAM)port);
                    if (InitSerialPort(port, hwnd)) {
                        isRunning = true;
                        isReceiving = true;
                        serialThread = new std::thread(SerialReadThread, hwnd);
                        SendMessageW(hLabelStatus, WM_SETTEXT, 0, (LPARAM)L"Status: Connected");
                        SendMessageW(hLabelReceivingStatus, WM_SETTEXT, 0, (LPARAM)L"Receiving: Yes");
                    }
                }
                else {
                    MessageBoxW(NULL, L"Please select a serial port", L"Error", MB_OK | MB_ICONERROR);
                }
            }
        }
        else if (LOWORD(wParam) == IDC_BUTTON_RELEASE && HIWORD(wParam) == BN_CLICKED) {
            if (isRunning) {
                isRunning = false;
                isReceiving = false;
                if (serialThread) {
                    serialThread->join();
                    delete serialThread;
                    serialThread = nullptr;
                }
                if (hSerial != INVALID_HANDLE_VALUE) {
                    CloseHandle(hSerial);
                    hSerial = INVALID_HANDLE_VALUE;
                }
                SendMessageW(hLabelStatus, WM_SETTEXT, 0, (LPARAM)L"Status: Disconnected");
                SendMessageW(hLabelReceivingStatus, WM_SETTEXT, 0, (LPARAM)L"Receiving: No");
            }
        }
        else if (LOWORD(wParam) == IDC_BUTTON_REFRESH && HIWORD(wParam) == BN_CLICKED) {
            UpdateSerialPorts(hwnd);
        }
        else if (LOWORD(wParam) == IDC_BUTTON_CONNECT && HIWORD(wParam) == BN_CLICKED) {
            if (isRunning && hSerial != INVALID_HANDLE_VALUE) {
                SendToSerial("start");
            }
        }
        else if (LOWORD(wParam) == IDC_BUTTON_DISCONNECT && HIWORD(wParam) == BN_CLICKED) {
            if (isRunning && hSerial != INVALID_HANDLE_VALUE) {
                SendToSerial("stop");
            }
        }
        else if (LOWORD(wParam) == IDC_BUTTON_STOP_RECEIVING && HIWORD(wParam) == BN_CLICKED) {
            if (isRunning) {
                isReceiving = !isReceiving;
                SendMessageW(hLabelReceivingStatus, WM_SETTEXT, 0, (LPARAM)(isReceiving ? L"Receiving: Yes" : L"Receiving: No"));
            }
        }
        else if (LOWORD(wParam) == IDC_BUTTON_W && HIWORD(wParam) == BN_CLICKED) {
            if (isRunning && hSerial != INVALID_HANDLE_VALUE) {
                SendToSerial(".KEY.W");
            }
        }
        else if (LOWORD(wParam) == IDC_BUTTON_A && HIWORD(wParam) == BN_CLICKED) {
            if (isRunning && hSerial != INVALID_HANDLE_VALUE) {
                SendToSerial(".KEY.A");
            }
        }
        else if (LOWORD(wParam) == IDC_BUTTON_S && HIWORD(wParam) == BN_CLICKED) {
            if (isRunning && hSerial != INVALID_HANDLE_VALUE) {
                SendToSerial(".KEY.S");
            }
        }
        else if (LOWORD(wParam) == IDC_BUTTON_D && HIWORD(wParam) == BN_CLICKED) {
            if (isRunning && hSerial != INVALID_HANDLE_VALUE) {
                SendToSerial(".KEY.D");
            }
        }
        else if (LOWORD(wParam) == IDC_BUTTON_SEND && HIWORD(wParam) == BN_CLICKED) {
            if (isRunning && hSerial != INVALID_HANDLE_VALUE) {
                wchar_t buffer[256];
                int len = SendMessageW(hEditInput, WM_GETTEXT, 256, (LPARAM)buffer);
                if (len > 0) {
                    std::wstring ws(buffer, len);
                    std::string data(ws.begin(), ws.end());
                    SendToSerial(data);
                }
            }
        }
        else if (LOWORD(wParam) == IDC_BUTTON_SEND2 && HIWORD(wParam) == BN_CLICKED) {
            if (isRunning && hSerial != INVALID_HANDLE_VALUE) {
                wchar_t buffer[256];
                int len = SendMessageW(hEditInput2, WM_GETTEXT, 256, (LPARAM)buffer);
                if (len > 0) {
                    std::wstring ws(buffer, len);
                    std::string data(ws.begin(), ws.end());
                    SendToSerial(data);
                }
            }
        }
        else if (LOWORD(wParam) == IDC_BUTTON_SEND3 && HIWORD(wParam) == BN_CLICKED) {
            if (isRunning && hSerial != INVALID_HANDLE_VALUE) {
                wchar_t buffer[256];
                int len = SendMessageW(hEditInput3, WM_GETTEXT, 256, (LPARAM)buffer);
                if (len > 0) {
                    std::wstring ws(buffer, len);
                    std::string data(ws.begin(), ws.end());
                    SendToSerial(data);
                }
            }
        }
        else if (LOWORD(wParam) == IDC_BUTTON_CLEAR && HIWORD(wParam) == BN_CLICKED) {
            SendMessageW(hEditOutput, WM_SETTEXT, 0, (LPARAM)L"");
            SendMessageW(hLabelOutputFull, WM_SETTEXT, 0, (LPARAM)L"Output not full");
        }
        else if (LOWORD(wParam) == IDC_BUTTON_DELAY_SEND && HIWORD(wParam) == BN_CLICKED) {
            if (!isRunning) {
                MessageBoxW(NULL, L"Connection not established", L"Error", MB_OK | MB_ICONERROR);
                return 0;
            }
            wchar_t timeBuffer[256], textBuffer[256];
            int timeLen = SendMessageW(hEditDelayTime, WM_GETTEXT, 256, (LPARAM)timeBuffer);
            int textLen = SendMessageW(hEditDelayText, WM_GETTEXT, 256, (LPARAM)textBuffer);
            if (timeLen == 0) return 0;
            std::wstring wsTime(timeBuffer, timeLen);
            std::string timeStr(wsTime.begin(), wsTime.end());
            int validTime = 0;
            for (char c : timeStr) if (isdigit(c)) validTime = validTime * 10 + (c - '0');
            if (validTime <= 0) return 0;
            std::wstring wsText(textBuffer, textLen);
            if (textLen == 0) {
                MessageBoxW(NULL, L"Please enter text to send", L"Error", MB_OK | MB_ICONERROR);
                return 0;
            }
            delayText = std::string(wsText.begin(), wsText.end());
            delayTime = validTime;
            delayLoop = SendMessageW(hComboDelayMode, CB_GETCURSEL, 0, 0) == 1;
            delayRunning = true;
            if (delayThread) {
                delayThread->join();
                delete delayThread;
            }
            delayThread = new std::thread(DelaySendThread, hwnd);
        }
        else if (LOWORD(wParam) == IDC_BUTTON_DELAY_STOP && HIWORD(wParam) == BN_CLICKED) {
            if (delayRunning) {
                delayRunning = false;
                if (delayThread) {
                    delayThread->join();
                    delete delayThread;
                    delayThread = nullptr;
                }
                SendMessageW(hLabelRemaining, WM_SETTEXT, 0, (LPARAM)L"Remaining: 0s");
            }
        }
        // 键盘模式切换
        else if (LOWORD(wParam) == IDC_COMBO_KEYMODE && HIWORD(wParam) == CBN_SELCHANGE) {
            int newKeyMode = (int)SendMessageW(hComboKeyMode, CB_GETCURSEL, 0, 0);
            if (newKeyMode != keyMode) {
                if (hKeyboardHook) { UnhookWindowsHookEx(hKeyboardHook); hKeyboardHook = NULL; }
                if (newKeyMode == 2 && isRunning && hSerial != INVALID_HANDLE_VALUE) {
                    hKeyboardHook = SetWindowsHookEx(WH_KEYBOARD_LL, KeyboardProc, hInst, 0);
                    if (!hKeyboardHook) MessageBoxW(NULL, L"Failed to install keyboard hook", L"Error", MB_OK | MB_ICONERROR);
                }
                keyMode = newKeyMode;
                if (keyMode == 1) SetFocus(hwnd);
            }
        }
        // 鼠标模式切换（ComboBox ID 2001）
        else if (LOWORD(wParam) == 2001 && HIWORD(wParam) == CBN_SELCHANGE) {
            int newMouseMode = (int)SendMessageW(hComboMouseMode, CB_GETCURSEL, 0, 0);
            if (newMouseMode != mouseMode) {
                if (hMouseHook) { UnhookWindowsHookEx(hMouseHook); hMouseHook = NULL; }
                if (newMouseMode == 2 && isRunning && hSerial != INVALID_HANDLE_VALUE) {
                    hMouseHook = SetWindowsHookEx(WH_MOUSE_LL, MouseProc, hInst, 0);
                    if (!hMouseHook) MessageBoxW(NULL, L"Failed to install mouse hook", L"Error", MB_OK | MB_ICONERROR);
                }
                mouseMode = newMouseMode;
                if (mouseMode == 1) SetFocus(hwnd);
            }
        }
        else if (LOWORD(wParam) == IDC_BTN_REFRESH_MIC && HIWORD(wParam) == BN_CLICKED) {
            PopulateMicList();
            return 0;
        }
        else if (LOWORD(wParam) == IDC_BTN_TOGGLE_SPEECH && HIWORD(wParam) == BN_CLICKED) {
            if (g_speechEnabled) StopSpeech(hwnd);
            else                 StartSpeech(hwnd);
            return 0;
        }
        break;

    }
    case WM_KEYDOWN: {
        if (isRunning && hSerial != INVALID_HANDLE_VALUE && keyMode == 1) { // 窗口焦点模式
            SetFocus(hwnd);
            wchar_t keyName[32];
            UINT scanCode = MapVirtualKey((UINT)wParam, MAPVK_VK_TO_VSC);
            if (GetKeyNameTextW(scanCode << 16, keyName, 32) > 0) {
                std::wstring wsKey(keyName);
                std::string keyStr(wsKey.begin(), wsKey.end());
                std::lock_guard<std::mutex> lk(inputQueueMutex);
                if (inputEventQueue.size() < kMaxQueueLen) {
                    inputEventQueue.push(".KEY." + keyStr);
                    inputQueueCV.notify_one();
                }
            }
        }
        break;
    }
    case WM_LBUTTONDOWN: {
        if (mouseMode == 1 && isRunning && hSerial != INVALID_HANDLE_VALUE) {
            std::lock_guard<std::mutex> lk(inputQueueMutex);
            if (inputEventQueue.size() < kMaxQueueLen) {
                inputEventQueue.push(".MOUSE.LEFT_DOWN");
                inputQueueCV.notify_one();
            }
        }
    } break;

    case WM_LBUTTONUP: {
        if (mouseMode == 1 && isRunning && hSerial != INVALID_HANDLE_VALUE) {
            std::lock_guard<std::mutex> lk(inputQueueMutex);
            if (inputEventQueue.size() < kMaxQueueLen) {
                inputEventQueue.push(".MOUSE.LEFT_UP");
                inputQueueCV.notify_one();
            }
        }
    } break;

    case WM_RBUTTONDOWN: {
        if (mouseMode == 1 && isRunning && hSerial != INVALID_HANDLE_VALUE) {
            std::lock_guard<std::mutex> lk(inputQueueMutex);
            if (inputEventQueue.size() < kMaxQueueLen) {
                inputEventQueue.push(".MOUSE.RIGHT_DOWN");
                inputQueueCV.notify_one();
            }
        }
    } break;

    case WM_RBUTTONUP: {
        if (mouseMode == 1 && isRunning && hSerial != INVALID_HANDLE_VALUE) {
            std::lock_guard<std::mutex> lk(inputQueueMutex);
            if (inputEventQueue.size() < kMaxQueueLen) {
                inputEventQueue.push(".MOUSE.RIGHT_UP");
                inputQueueCV.notify_one();
            }
        }
    } break;

    case WM_MBUTTONDOWN: {
        if (mouseMode == 1 && isRunning && hSerial != INVALID_HANDLE_VALUE) {
            std::lock_guard<std::mutex> lk(inputQueueMutex);
            if (inputEventQueue.size() < kMaxQueueLen) {
                inputEventQueue.push(".MOUSE.MIDDLE_DOWN");
                inputQueueCV.notify_one();
            }
        }
    } break;

    case WM_MBUTTONUP: {
        if (mouseMode == 1 && isRunning && hSerial != INVALID_HANDLE_VALUE) {
            std::lock_guard<std::mutex> lk(inputQueueMutex);
            if (inputEventQueue.size() < kMaxQueueLen) {
                inputEventQueue.push(".MOUSE.MIDDLE_UP");
                inputQueueCV.notify_one();
            }
        }
    } break;

    case WM_MOUSEMOVE: {
        if (mouseMode == 1 && isRunning && hSerial != INVALID_HANDLE_VALUE) {
            DWORD now = GetTickCount();
            if (now - g_lastMouseMoveTick.load() >= g_mouseMoveIntervalMs) {
                int x = GET_X_LPARAM(lParam);
                int y = GET_Y_LPARAM(lParam);
                std::ostringstream oss; oss << ".MOUSE.MOVE:" << x << "," << y;

                std::lock_guard<std::mutex> lk(inputQueueMutex);
                if (inputEventQueue.size() < kMaxQueueLen) {
                    inputEventQueue.push(oss.str());
                    inputQueueCV.notify_one();
                }
                g_lastMouseMoveTick.store(now);
            }
        }
    } break;

    case WM_MOUSEWHEEL: {
        if (mouseMode == 1 && isRunning && hSerial != INVALID_HANDLE_VALUE) {
            short delta = GET_WHEEL_DELTA_WPARAM(wParam);
            std::lock_guard<std::mutex> lk(inputQueueMutex);
            if (inputEventQueue.size() < kMaxQueueLen) {
                inputEventQueue.push(delta > 0 ? ".MOUSE.WHEEL_UP" : ".MOUSE.WHEEL_DOWN");
                inputQueueCV.notify_one();
            }
        }
    } break;
    case WM_UPDATE_OUTPUT: {
        std::wstring localBuffer;
        {
            std::lock_guard<std::mutex> lock(mtx);
            localBuffer.swap(outputBuffer);
        }
        if (!localBuffer.empty()) {
            SendMessageW(hEditOutput, EM_SETSEL, -1, -1);
            SendMessageW(hEditOutput, EM_REPLACESEL, 0, (LPARAM)localBuffer.c_str());
            SendMessageW(hEditOutput, EM_SCROLLCARET, 0, 0);
            if (localBuffer.length() > 1000) {
                SendMessageW(hLabelOutputFull, WM_SETTEXT, 0, (LPARAM)L"Output full");
            }
            else {
                SendMessageW(hLabelOutputFull, WM_SETTEXT, 0, (LPARAM)L"Output not full");
            }
        }
        break;
    }
    case WM_UPDATE_REMAINING: {
        int remaining = (int)wParam;
        std::wstringstream ss;
        ss << L"Remaining: " << remaining << L"s";
        SendMessageW(hLabelRemaining, WM_SETTEXT, 0, (LPARAM)ss.str().c_str());
        if (remaining <= 0 && delayRunning && delayLoop) {
            if (hSerial != INVALID_HANDLE_VALUE) SendToSerial(delayText);
            delayTime = SendMessageW(hEditDelayTime, WM_GETTEXTLENGTH, 0, 0) > 0 ? delayTime : 0;
            if (delayThread) {
                delayThread->join();
                delete delayThread;
            }
            delayThread = new std::thread(DelaySendThread, hwnd);
        }
        break;
    }
    case WM_DESTROY: {
        // 先卸载全局钩子
        if (hMouseHook != NULL) { UnhookWindowsHookEx(hMouseHook); hMouseHook = NULL; }
        if (hKeyboardHook != NULL) { UnhookWindowsHookEx(hKeyboardHook); hKeyboardHook = NULL; }

        // 停统一输入发送线程
        if (inputSendThread) {
            inputThreadRunning = false;
            inputQueueCV.notify_all();
            inputSendThread->join();
            delete inputSendThread;
            inputSendThread = nullptr;
        }
        // 下面保留你原有的串口/接收/延迟线程清理…
        if (isRunning) {
            isRunning = false;
            isReceiving = false;
            if (serialThread) {
                serialThread->join();
                delete serialThread;
            }
            if (hSerial != INVALID_HANDLE_VALUE) {
                CloseHandle(hSerial);
            }
        }
        if (delayRunning) {
            delayRunning = false;
            if (delayThread) {
                delayThread->join();
                delete delayThread;
            }
        }
        if (g_speechEnabled) StopSpeech(hwnd);
        if (g_comInited) { CoUninitialize(); g_comInited = false; }
        PostQuitMessage(0);
        break;
    }
    case WM_RECOEVENT: {
        if (g_speechEnabled && g_cpRecoContext) {
            CSpEvent evt;
            while (evt.GetFrom(g_cpRecoContext) == S_OK) {
                if (evt.eEventId == SPEI_RECOGNITION) {
                    ISpRecoResult* pRes = evt.RecoResult();
                    if (pRes) {
                        LPWSTR pwszText = nullptr;
                        if (SUCCEEDED(pRes->GetText(SP_GETWHOLEPHRASE, SP_GETWHOLEPHRASE, TRUE, &pwszText, nullptr)) && pwszText) {
                            // 串口消息格式：.SPEECH.<text>;
                            std::wstring wmsg = L".SPEECH.";
                            wmsg += pwszText;
                            std::string send = Narrow(wmsg) + ";";
                            SendToSerial(send);

                            // 界面显示“最后一次发送”
                            std::wstring wlabel = L"Last sent: ";
                            wlabel += wmsg;
                            wlabel += L";";
                            if (hLastSentLabel) SetWindowTextW(hLastSentLabel, wlabel.c_str());

                            ::CoTaskMemFree(pwszText);
                        }
                    }
                }
            }
        }
        return 0;
    }
    default:
        return DefWindowProcW(hwnd, msg, wParam, lParam);
    }
    return 0;
}

// Main entry point
int WINAPI wWinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, PWSTR pCmdLine, int nCmdShow) {
    hInst = hInstance;

    WNDCLASSEXW wc = { 0 };
    wc.cbSize = sizeof(WNDCLASSEXW);
    wc.lpfnWndProc = WndProc;
    wc.hInstance = hInstance;
    wc.lpszClassName = L"SerialPortApp";
    wc.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);
    RegisterClassExW(&wc);

    HWND hwnd = CreateWindowW(L"SerialPortApp", L"Serial Port Control", WS_OVERLAPPEDWINDOW,
        CW_USEDEFAULT, CW_USEDEFAULT, // 位置（系统默认）
        1080, 900,// <-- 宽度在前高度在后（默认大小）
        NULL, NULL, hInstance, NULL);


    if (!hwnd) {
        MessageBoxW(NULL, L"Window creation failed", L"Error", MB_OK | MB_ICONERROR);
        return 1;
    }

    ShowWindow(hwnd, nCmdShow);
    UpdateWindow(hwnd);

    MSG msg;
    while (GetMessageW(&msg, NULL, 0, 0)) {
        TranslateMessage(&msg);
        DispatchMessageW(&msg);
    }

    return (int)msg.wParam;
}