// TaskBarAppMenu.cpp : Defines the entry point for the application.
//

#include "framework.h"
#include "TaskBarAppMenu.h"

#define MAX_LOADSTRING 100

// Global Variables:
HINSTANCE hInst;                                // current instance
WCHAR szTitle[MAX_LOADSTRING];                  // The title bar text
WCHAR szWindowClass[MAX_LOADSTRING];            // the main window class name

// Forward declarations of functions included in this code module:
ATOM                MyRegisterClass(HINSTANCE hInstance);
BOOL                InitInstance(HINSTANCE, int);
LRESULT CALLBACK    WndProc(HWND, UINT, WPARAM, LPARAM);
INT_PTR CALLBACK    About(HWND, UINT, WPARAM, LPARAM);








#include <windows.h>
#include <shobjidl.h>
#include <propkey.h>
#include <propvarutil.h>
#include <shlobj.h>
#include <string>
#include <vector>

// 链接必要的库
#pragma comment(lib, "propsys.lib")
#pragma comment(lib, "shell32.lib")

// 辅助函数：创建一个简单的 ShellLink (快捷方式对象)
HRESULT CreateShellLink(PCWSTR pszPath, PCWSTR pszArguments, PCWSTR pszTitle, IShellLink** ppsl)
{
    IShellLink* psl = nullptr;
    HRESULT hr = CoCreateInstance(CLSID_ShellLink, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&psl));

    if (SUCCEEDED(hr))
    {
        psl->SetPath(pszPath);
        if (pszArguments)
        {
            psl->SetArguments(pszArguments);
        }

        // 设置标题（显示在菜单上的文字）
        IPropertyStore* pps = nullptr;
        hr = psl->QueryInterface(IID_PPV_ARGS(&pps));
        if (SUCCEEDED(hr))
        {

            PROPVARIANT var;
            hr = InitPropVariantFromString(pszTitle, &var);
            if (SUCCEEDED(hr))
            {
                hr = pps->SetValue(PKEY_Title, var);
                PropVariantClear(&var);
            }
            pps->Release();
        }

        if (SUCCEEDED(hr))
        {
            *ppsl = psl;
        }
        else
        {
            psl->Release();
        }
    }
    return hr;
}

// 核心函数：构建并应用 Jump List
HRESULT BuildJumpListWithSubMenu()
{
    ICustomDestinationList* pcdl = nullptr;
    HRESULT hr = CoCreateInstance(CLSID_DestinationList, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&pcdl));

    if (FAILED(hr))
    {
        return hr;
    }

    // 1. 开始构建列表，获取最大条目数和类别集合
    UINT cMaxSlots = 0;
    IObjectArray* paoRemoved = nullptr;
    hr = pcdl->BeginList(&cMaxSlots, IID_PPV_ARGS(&paoRemoved));

    if (SUCCEEDED(hr))
    {
        // ---------------------------------------------------------
        // 步骤 A: 创建二级菜单中的子项
        // ---------------------------------------------------------
        IObjectCollection* pocSubItems = nullptr;
        hr = CoCreateInstance(CLSID_EnumerableObjectCollection, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&pocSubItems));

        if (SUCCEEDED(hr))
        {
            // 子项 1: 记事本
            IShellLink* pslNotepad = nullptr;
            if (SUCCEEDED(CreateShellLink(L"C:\\Windows\\System32\\notepad.exe", nullptr, L"打开记事本", &pslNotepad)))
            {
                pocSubItems->AddObject(pslNotepad);
                pslNotepad->Release();
            }

            // 子项 2: 计算器
            IShellLink* pslCalc = nullptr;
            if (SUCCEEDED(CreateShellLink(L"C:\\Windows\\System32\\calc.exe", nullptr, L"打开计算器", &pslCalc)))
            {
                pocSubItems->AddObject(pslCalc);
                pslCalc->Release();
            }

            // 子项 3：弹出一个菜单IDC_TASKBARAPPMENU
			IShellLink* pslMenu = nullptr;
			if (SUCCEEDED(CreateShellLink(L"rundll32.exe", L"shell32.dll,Control_RunDLL", L"打开控制面板", &pslMenu)))
			{
				pocSubItems->AddObject(pslMenu);
				pslMenu->Release();
			}

            // ---------------------------------------------------------
            // 步骤 B: 创建分类并将子项集合添加进去
            // ---------------------------------------------------------
            pcdl->AppendCategory(L"我的工具集", pocSubItems);

            pocSubItems->Release();
        }

        // 提交更改
        pcdl->CommitList();

        if (paoRemoved)
        {
            paoRemoved->Release();
        }
    }

    pcdl->Release();

    return hr;
}

int APIENTRY wWinMain(_In_ HINSTANCE hInstance,
                     _In_opt_ HINSTANCE hPrevInstance,
                     _In_ LPWSTR    lpCmdLine,
                     _In_ int       nCmdShow)
{
    // 1. 初始化 COM 库
    CoInitialize(nullptr);

    // 2. 设置应用程序的用户模型 ID (AppUserModelID)
    // 修正：使用 GetCurrentProcess 直接设置当前进程的 AppUserModelID
    //HRESULT hr = SetCurrentProcessExplicitAppUserModelID(L"Company.Product.MyApp1");

    // 3. 构建带有二级菜单的跳转列表
    //if (SUCCEEDED(hr))
    {
        BuildJumpListWithSubMenu();
    }

    UNREFERENCED_PARAMETER(hPrevInstance);
    UNREFERENCED_PARAMETER(lpCmdLine);

    // TODO: Place code here.

    // Initialize global strings
    LoadStringW(hInstance, IDS_APP_TITLE, szTitle, MAX_LOADSTRING);
    LoadStringW(hInstance, IDC_TASKBARAPPMENU, szWindowClass, MAX_LOADSTRING);
    MyRegisterClass(hInstance);

    // Perform application initialization:
    if (!InitInstance (hInstance, nCmdShow))
    {
        return FALSE;
    }

    HACCEL hAccelTable = LoadAccelerators(hInstance, MAKEINTRESOURCE(IDC_TASKBARAPPMENU));

    MSG msg;

    // Main message loop:
    while (GetMessage(&msg, nullptr, 0, 0))
    {
        if (!TranslateAccelerator(msg.hwnd, hAccelTable, &msg))
        {
            TranslateMessage(&msg);
            DispatchMessage(&msg);
        }
    }

    CoUninitialize();

    return (int) msg.wParam;
}



//
//  FUNCTION: MyRegisterClass()
//
//  PURPOSE: Registers the window class.
//
ATOM MyRegisterClass(HINSTANCE hInstance)
{
    WNDCLASSEXW wcex;

    wcex.cbSize = sizeof(WNDCLASSEX);

    wcex.style          = CS_HREDRAW | CS_VREDRAW;
    wcex.lpfnWndProc    = WndProc;
    wcex.cbClsExtra     = 0;
    wcex.cbWndExtra     = 0;
    wcex.hInstance      = hInstance;
    wcex.hIcon          = LoadIcon(hInstance, MAKEINTRESOURCE(IDI_TASKBARAPPMENU));
    wcex.hCursor        = LoadCursor(nullptr, IDC_ARROW);
    wcex.hbrBackground  = (HBRUSH)(COLOR_WINDOW+1);
    wcex.lpszMenuName   = MAKEINTRESOURCEW(IDC_TASKBARAPPMENU);
    wcex.lpszClassName  = szWindowClass;
    wcex.hIconSm        = LoadIcon(wcex.hInstance, MAKEINTRESOURCE(IDI_SMALL));

    return RegisterClassExW(&wcex);
}

//
//   FUNCTION: InitInstance(HINSTANCE, int)
//
//   PURPOSE: Saves instance handle and creates main window
//
//   COMMENTS:
//
//        In this function, we save the instance handle in a global variable and
//        create and display the main program window.
//
BOOL InitInstance(HINSTANCE hInstance, int nCmdShow)
{
   hInst = hInstance; // Store instance handle in our global variable

   HWND hWnd = CreateWindowW(szWindowClass, szTitle, WS_OVERLAPPEDWINDOW,
      CW_USEDEFAULT, 0, CW_USEDEFAULT, 0, nullptr, nullptr, hInstance, nullptr);

   if (!hWnd)
   {
      return FALSE;
   }

   ShowWindow(hWnd, nCmdShow);
   UpdateWindow(hWnd);

   return TRUE;
}

//
//  FUNCTION: WndProc(HWND, UINT, WPARAM, LPARAM)
//
//  PURPOSE: Processes messages for the main window.
//
//  WM_COMMAND  - process the application menu
//  WM_PAINT    - Paint the main window
//  WM_DESTROY  - post a quit message and return
//
//
LRESULT CALLBACK WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam)
{
    switch (message)
    {
    case WM_COMMAND:
        {
            int wmId = LOWORD(wParam);
            // Parse the menu selections:
            switch (wmId)
            {
            case IDM_ABOUT:
                DialogBox(hInst, MAKEINTRESOURCE(IDD_ABOUTBOX), hWnd, About);
                break;
            case IDM_EXIT:
                DestroyWindow(hWnd);
                break;
            default:
                return DefWindowProc(hWnd, message, wParam, lParam);
            }
        }
        break;
    case WM_PAINT:
        {
            PAINTSTRUCT ps;
            HDC hdc = BeginPaint(hWnd, &ps);
            // TODO: Add any drawing code that uses hdc here...
            EndPaint(hWnd, &ps);
        }
        break;
    case WM_DESTROY:
        PostQuitMessage(0);
        break;
    default:
        return DefWindowProc(hWnd, message, wParam, lParam);
    }
    return 0;
}

// Message handler for about box.
INT_PTR CALLBACK About(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam)
{
    UNREFERENCED_PARAMETER(lParam);
    switch (message)
    {
    case WM_INITDIALOG:
        return (INT_PTR)TRUE;

    case WM_COMMAND:
        if (LOWORD(wParam) == IDOK || LOWORD(wParam) == IDCANCEL)
        {
            EndDialog(hDlg, LOWORD(wParam));
            return (INT_PTR)TRUE;
        }
        break;
    }
    return (INT_PTR)FALSE;
}
