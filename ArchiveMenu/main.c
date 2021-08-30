#pragma comment(linker,"\"/manifestdependency:type='win32' \
name='Microsoft.Windows.Common-Controls' version='6.0.0.0' \
processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")
#include <initguid.h>
#include <Windows.h>
#include <Shlwapi.h>
#include <Shlobj_core.h>
#pragma comment(lib, "Shlwapi.lib")
#include <tlhelp32.h>

#define CLASS_NAME L"ArchiveMenu"
#define OPEN_NAME L"&Open archive"
#define EXTRACT_NAME L"&Extract to \"%s\\\""
#define OPEN_CMD L"\"C:\\Program Files\\7-Zip\\7zFM.exe\" %s"
#define EXTRACT_CMD L"\"C:\\Program Files\\7-Zip\\7zG.exe\" x -o\"%s\" -spe %s"

DEFINE_GUID(__uuidof_TaskbarList,
    0x56FDF344,
    0xFD6D, 0x11d0, 0x95, 0x8A,
    0x00, 0x60, 0x97, 0xC9, 0xA0, 0x90
);
DEFINE_GUID(__uuidof_ITaskbarList,
    0x56FDF342,
    0xFD6D, 0x11d0, 0x95, 0x8A,
    0x00, 0x60, 0x97, 0xC9, 0xA0, 0x90
);

DWORD WINAPI SendCommandLineThread(HWND hWnd)
{
    COPYDATASTRUCT st;
    st.cbData = wcslen(PathGetArgsW(GetCommandLineW())) * sizeof(TCHAR);
    st.lpData = PathGetArgsW(GetCommandLineW());
    st.dwData = 1;
    SendMessage(
        FindWindow(L"ArchiveMenuWindowExplorer", NULL),
        WM_COPYDATA,
        hWnd,
        &st
    );
    TerminateProcess(
        OpenProcess(
            PROCESS_TERMINATE,
            FALSE,
            GetCurrentProcessId()
        ),
        0
    );
}

int WINAPI wWinMain(
    _In_ HINSTANCE hInstance,
    _In_opt_ HINSTANCE hPrevInstance,
    _In_ LPWSTR lpCmdLine,
    _In_ int nShowCmd
)
{
    SetProcessDpiAwarenessContext(
        DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2
    );

    HRESULT hr = CoInitialize(NULL);
    ITaskbarList* pTaskList = NULL;
    hr = CoCreateInstance(
        &__uuidof_TaskbarList,
        NULL,
        CLSCTX_ALL,
        &__uuidof_ITaskbarList,
        (void**)(&pTaskList)
    );
    hr = pTaskList->lpVtbl->HrInit(pTaskList);

    WNDCLASS wc = { 0 };
    wc.style = CS_DBLCLKS;
    wc.lpfnWndProc = DefWindowProc;
    wc.hbrBackground = (HBRUSH)GetStockObject(WHITE_BRUSH);
    wc.hInstance = GetModuleHandle(NULL);
    wc.lpszClassName = CLASS_NAME;
    wc.hCursor = LoadCursorW(NULL, IDC_ARROW);
    RegisterClass(&wc);

    HWND hWnd = CreateWindowEx(
        0,
        CLASS_NAME,
        0,
        WS_POPUP,
        0,
        0,
        0,
        0,
        0,
        0,
        hInstance,
        NULL
    );
    ShowWindow(hWnd, SW_SHOW);
    hr = pTaskList->lpVtbl->DeleteTab(pTaskList, hWnd);

    BOOL canShowToast = FALSE;
    PROCESSENTRY32 pe32 = { 0 };
    pe32.dwSize = sizeof(PROCESSENTRY32);
    HANDLE hSnapshot = CreateToolhelp32Snapshot(
        TH32CS_SNAPPROCESS,
        0
    );
    if (Process32First(hSnapshot, &pe32) == TRUE)
    {
        do
        {
            if (!wcscmp(pe32.szExeFile, TEXT("explorer.exe")))
            {
                canShowToast = TRUE;
                break;
            }
        } while (Process32Next(hSnapshot, &pe32) == TRUE);
    }

    CloseHandle(hSnapshot);

    if (canShowToast && FindWindow(L"ArchiveMenuWindowExplorer", NULL))
    {
        CreateThread(0, 0, SendCommandLineThread, hWnd, 0, 0);

        MSG msg = { 0 };
        while (GetMessage(&msg, NULL, 0, 0))
        {
            TranslateMessage(&msg);
            DispatchMessage(&msg);
        }

        TerminateProcess(
            OpenProcess(
                PROCESS_TERMINATE,
                FALSE,
                GetCurrentProcessId()
            ),
            0
        );
    }

    POINT pt;
    GetCursorPos(&pt);

    HMENU hMenu = CreatePopupMenu();
    TCHAR buffer[MAX_PATH + 100];
    TCHAR filename[MAX_PATH];
    ZeroMemory(filename, MAX_PATH * sizeof(TCHAR));
    memcpy(filename, PathGetArgsW(GetCommandLineW()), wcslen(PathGetArgsW(GetCommandLineW())) * sizeof(TCHAR));
    PathUnquoteSpacesW(filename);
    PathRemoveExtensionW(filename);
    PathStripPathW(filename);
    wsprintf(buffer, EXTRACT_NAME, filename);

    InsertMenu(hMenu, 0, MF_BYPOSITION | MF_STRING, 1, buffer);
    InsertMenu(hMenu, 0, MF_BYPOSITION | MF_STRING, 2, OPEN_NAME);
    BOOL res = TrackPopupMenu(
        hMenu,
        TPM_RETURNCMD,
        pt.x - 10,
        pt.y - 10,
        0,
        hWnd,
        0
    );
    if (res == 1 || res == 2)
    {
        TCHAR buffer[MAX_PATH + 100];
        ZeroMemory(buffer, (MAX_PATH + 100) * sizeof(TCHAR));
        if (res == 2)
        {
            wsprintf(buffer, OPEN_CMD, PathGetArgsW(GetCommandLineW()));
        }
        else if (res == 1)
        {
            TCHAR path[MAX_PATH];
            ZeroMemory(path, MAX_PATH * sizeof(TCHAR));
            memcpy(path, PathGetArgsW(GetCommandLineW()), wcslen(PathGetArgsW(GetCommandLineW())) * sizeof(TCHAR));
            PathUnquoteSpacesW(path);
            PathRemoveExtensionW(path);
            wsprintf(buffer, EXTRACT_CMD, path, PathGetArgsW(GetCommandLineW()));
        }
        STARTUPINFO info = { sizeof(info) };
        PROCESS_INFORMATION processInfo;
        BOOL b = CreateProcess(
            NULL,
            buffer,
            NULL,
            NULL,
            TRUE,
            CREATE_UNICODE_ENVIRONMENT,
            NULL,
            NULL,
            &info,
            &processInfo
        );
    }

    return 0;
}