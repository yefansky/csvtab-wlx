#include <windows.h>
#include <iostream>

// Your ListLoad function declaration
HWND APIENTRY ListLoad(HWND hListerWnd, char* fileToLoad, int showFlags);

LRESULT CALLBACK WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam) {

    switch (message) {
        // Handle window messages as needed
    case WM_DESTROY:
        PostQuitMessage(0);
        break;
    default:
        return DefWindowProc(hWnd, message, wParam, lParam);
    }
    return 0;
}

//const char* cpszTestFile = "H:\\sword3-products_Classic\\trunk\\client\\settings\\skill\\skills.tab";
const char* cpszTestFile = "h:\\sword3-products_Classic\\trunk\\client\\settings\\NpcTemplate.tab";

int APIENTRY wWinMain(_In_ HINSTANCE hInstance,
    _In_opt_ HINSTANCE hPrevInstance,
    _In_ LPWSTR lpCmdLine,
    _In_ int nCmdShow)
{
    HMODULE dllHandle = LoadLibrary(TEXT("csvtab-wlx.dll"));
    if (dllHandle) {
        using ListLoadFunc = HWND(APIENTRY*)(HWND, const char*, int);
        ListLoadFunc listLoad = reinterpret_cast<ListLoadFunc>(GetProcAddress(dllHandle, "ListLoad"));
        if (listLoad) {
            // Create a window for testing
            WNDCLASS wc = { 0 };
            wc.lpfnWndProc = WndProc;
            wc.hInstance = GetModuleHandle(NULL);
            wc.lpszClassName = L"MyWindowClass";
            RegisterClass(&wc);

            int newWidth = 1000;
            int newHeight = 800;

            HWND hMainWindow = CreateWindow(
                L"MyWindowClass", L"My Window",
                WS_OVERLAPPEDWINDOW,
                CW_USEDEFAULT, CW_USEDEFAULT, newWidth, newHeight,
                NULL, NULL, GetModuleHandle(NULL), NULL);

            if (hMainWindow) {
                ShowWindow(hMainWindow, nCmdShow);
                UpdateWindow(hMainWindow);

                // Call ListLoad function with the created window handle
                HWND result = listLoad(hMainWindow, cpszTestFile, 1);

                if (result) {
                    // Process the result or perform additional tests

                    // SetWindowPos parameters
                    SetWindowPos(result, NULL, 0, 0, newWidth -20, newHeight - 20, SWP_NOMOVE | SWP_NOZORDER | SWP_NOACTIVATE);

                    // Enter the message loop
                    MSG msg;
                    while (GetMessage(&msg, NULL, 0, 0)) {
                        TranslateMessage(&msg);
                        DispatchMessage(&msg);
                    }
                }
                else {
                    std::cerr << "ListLoad function returned NULL." << std::endl;
                }

                // Clean up
                DestroyWindow(hMainWindow);
            }
        }
        else {
            std::cerr << "Failed to get function pointer." << std::endl;
        }
        FreeLibrary(dllHandle);
    }
    else {
        std::cerr << "Failed to load DLL." << std::endl;
    }

    return 0;
}
