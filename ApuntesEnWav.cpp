// Copyright (c) 2025 Anibal Zanutti
// Licensed under the MIT License. See LICENSE file for details.

#include <windows.h>
#include <initguid.h>
#include <sapi.h>
#include <commdlg.h>
#include <fstream>
#include <string>
#include <shellapi.h>
#pragma comment(lib, "ole32.lib")
#pragma comment(lib, "oleaut32.lib")
#pragma comment(lib, "uuid.lib")
#pragma comment(lib, "winnm.lib")
#pragma comment(lib, "sapi.lib")

#define IDC_EDIT        101
#define IDC_BTN_LEER    102
#define IDC_BTN_ABRIR   103
#define IDC_BTN_BORRAR  104
#define IDC_BTN_GUARDAR 105
#define IDC_BTN_CERRAR  106

HWND hEdit;
ISpVoice* pVoice = NULL;
bool modoGrafico = true;
// {C31ADBAE-527F-4ff5-A230-F62BB61FF70C}
 DEFINE_GUID(SPDFID_WaveFormatEx, 0xc31adbae, 0x527f, 0x4ff5, 0xa2, 0x30, 0xf6, 0x2b, 0xb6, 0x1f, 0xf7, 0x0c);

void GuardarTextoComoWav(const std::wstring& texto, const std::wstring& archivo) {
    HRESULT hr;
    ISpVoice* pVoice = nullptr;
    ISpStream* pStream = nullptr;

    hr = CoCreateInstance(CLSID_SpVoice, NULL, CLSCTX_ALL, IID_ISpVoice, (void**)&pVoice);
    if (FAILED(hr)) {
        MessageBoxW(NULL, L"No se pudo crear instancia de voz", L"Error", MB_OK);
        return;
    }

    hr = CoCreateInstance(CLSID_SpStream, NULL, CLSCTX_ALL, IID_ISpStream, (void**)&pStream);
    if (FAILED(hr)) {
        MessageBoxW(NULL, L"No se pudo crear instancia de stream", L"Error", MB_OK);
        pVoice->Release();
        return;
    }

    WAVEFORMATEX wfx = {0};
    wfx.wFormatTag = WAVE_FORMAT_PCM;
    wfx.nChannels = 1;
    wfx.nSamplesPerSec = 16000;
    wfx.wBitsPerSample = 16;
    wfx.nBlockAlign = (wfx.nChannels * wfx.wBitsPerSample) / 8;
    wfx.nAvgBytesPerSec = wfx.nBlockAlign * wfx.nSamplesPerSec;

    hr = pStream->BindToFile(
        archivo.c_str(),
        SPFM_CREATE_ALWAYS,
        &SPDFID_WaveFormatEx,
        &wfx,
        sizeof(WAVEFORMATEX)
    );
    if (FAILED(hr)) {
        wchar_t buffer[256];
        swprintf(buffer, 256, L"Error al ligar el stream. HRESULT = 0x%08X", hr);
        MessageBoxW(NULL, buffer, L"Error", MB_OK);

        pStream->Release();
        pVoice->Release();
        return;
    }

    pVoice->SetOutput(pStream, TRUE);
    pVoice->Speak(texto.c_str(), SPF_DEFAULT, NULL);

    pStream->Close();
    pStream->Release();
    pVoice->Release();

    MessageBoxW(NULL, L"¡Archivo WAV guardado exitosamente!", L"Listo", MB_OK);
}

// Conversión de string a wstring (ANSI → Unicode)
std::wstring StringToWString(const std::string& str) {
    int size_needed = MultiByteToWideChar(CP_ACP, 0, str.c_str(), -1, NULL, 0);
    std::wstring wstr(size_needed, 0);
    MultiByteToWideChar(CP_ACP, 0, str.c_str(), -1, &wstr[0], size_needed);
    return wstr;
}

// Función para leer texto
void LeerTexto(const std::string& texto) {
    if (pVoice && !texto.empty()) {
        std::wstring wtexto = StringToWString(texto);
        pVoice->Speak(wtexto.c_str(), SPF_DEFAULT, NULL);
    }
}

// Función para abrir archivo de texto (ANSI)
std::string AbrirArchivo(HWND hwnd) {
    OPENFILENAME ofn;
    char szFile[260] = {0};

    ZeroMemory(&ofn, sizeof(ofn));
    ofn.lStructSize = sizeof(ofn);
    ofn.hwndOwner = hwnd;
    ofn.lpstrFilter = "Archivos de texto\0*.TXT\0";
    ofn.lpstrFile = szFile;
    ofn.nMaxFile = sizeof(szFile);
    ofn.Flags = OFN_PATHMUSTEXIST | OFN_FILEMUSTEXIST;

    if (GetOpenFileName(&ofn)) {
        std::ifstream file(szFile);
        std::string contenido((std::istreambuf_iterator<char>(file)), std::istreambuf_iterator<char>());
        return contenido;
    }
    return "";
}

LRESULT CALLBACK WndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    switch (msg) {
        case WM_CREATE:
            hEdit = CreateWindow(
                "EDIT", "",
                WS_CHILD | WS_VISIBLE | WS_BORDER |
                ES_LEFT | ES_MULTILINE | ES_AUTOVSCROLL | WS_VSCROLL | ES_WANTRETURN,
                20, 20, 360, 100,
                hwnd, (HMENU)IDC_EDIT, NULL, NULL
            );
            CreateWindow("BUTTON", "Leer", WS_CHILD | WS_VISIBLE,
                         20, 130, 60, 20, hwnd, (HMENU)IDC_BTN_LEER, NULL, NULL);
            CreateWindow("BUTTON", "Abrir y leer", WS_CHILD | WS_VISIBLE,
                         90, 130, 80, 20, hwnd, (HMENU)IDC_BTN_ABRIR, NULL, NULL);
            CreateWindow("BUTTON", "Borrar", WS_CHILD | WS_VISIBLE,
                         180, 130, 60, 20, hwnd, (HMENU)IDC_BTN_BORRAR, NULL, NULL);
            CreateWindow("BUTTON", "Guardar", WS_CHILD | WS_VISIBLE,
                         250, 130, 60, 20, hwnd, (HMENU)IDC_BTN_GUARDAR, NULL, NULL);
            CreateWindow("BUTTON", "Cerrar", WS_CHILD | WS_VISIBLE,
                         320, 130, 60, 20, hwnd, (HMENU)IDC_BTN_CERRAR, NULL, NULL);
            break;

        case WM_COMMAND:
            if (LOWORD(wParam) == IDC_BTN_LEER) { // Botón Leer
                char buffer[2048];
                GetWindowText(hEdit, buffer, sizeof(buffer));
                LeerTexto(buffer);
            }
            if (LOWORD(wParam) == IDC_BTN_ABRIR) { // Botón Abrir y Leer
                std::string texto = AbrirArchivo(hwnd);
                if (!texto.empty()) {
                    // Reemplazar '\n' por "\r\n" para el control EDIT
                    size_t pos = 0;
                    while ((pos = texto.find('\n', pos)) != std::string::npos) {
                        texto.replace(pos, 1, "\r\n");
                        pos += 2; // Avanza el cursor después del reemplazo
                    }
                    SetWindowText(hEdit, texto.c_str());
                    LeerTexto(texto);
                }
            }
            if (LOWORD(wParam) == IDC_BTN_BORRAR) { // Botón Cerrar
                    SetWindowText(hEdit, "");
            }
            if (LOWORD(wParam) == IDC_BTN_GUARDAR) { // Botón Cerrar
                    char buffer[2048];
                    GetWindowText(hEdit, buffer, sizeof(buffer));

                    // Pide al usuario la ruta para guardar el WAV
                    char szFile[MAX_PATH] = {0};
                    OPENFILENAME ofn = {0};
                    ofn.lStructSize = sizeof(ofn);
                    ofn.hwndOwner = hwnd;
                    ofn.lpstrFilter = "Archivo WAV (*.wav)\0*.wav\0";
                    ofn.lpstrFile = szFile;
                    ofn.nMaxFile = MAX_PATH;
                    ofn.Flags = OFN_OVERWRITEPROMPT | OFN_PATHMUSTEXIST;

                    if (GetSaveFileName(&ofn)) {
                        std::wstring texto = StringToWString(std::string(buffer));
                        std::wstring warchivo = StringToWString(std::string(szFile));
                        GuardarTextoComoWav(texto, warchivo);
                    }
            }
            if (LOWORD(wParam) == IDC_BTN_CERRAR) { // Botón Cerrar
                    MessageBox(
                        NULL,
                        "Programado por Anibal Zanutti \n Copyright (c) 2025.",
                        "Gracias por usar el programa",
                        MB_OK | MB_ICONINFORMATION
                    );
                    PostQuitMessage(0);
            }
            break;

        case WM_CTLCOLOREDIT: {
            HBRUSH hBrushEditFondo = NULL;
            HDC hdcEdit = (HDC)wParam;
            HWND hwndEditControl = (HWND)lParam;

            if (hwndEditControl == hEdit) {
                SetBkColor(hdcEdit, RGB(255, 248, 220));   // Fondo suave Cornsilk
                SetTextColor(hdcEdit, RGB(100, 10, 0)); // Texto Color Vino Oscuro
                return (INT_PTR)hBrushEditFondo;
            }
            break;
        }
        case WM_DESTROY:
            if(modoGrafico) {
                MessageBox(
                    NULL,
                    "Programado por Anibal Zanutti \n Copyright (c) 2025.",
                    "Gracias por usar el programa",
                    MB_OK | MB_ICONINFORMATION
                );
            }
            if (pVoice) pVoice->Release();
            PostQuitMessage(0);
            break;

        default:
            return DefWindowProc(hwnd, msg, wParam, lParam);
    }
    return 0;
}

int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance,
                   LPSTR lpCmdLine, int nCmdShow) {

    // Si se pasan argumentos: usar como modo consola
        if (__argc > 1) {
        modoGrafico = false; // ← MODO CONSOLA detectado
        std::string texto;
        bool guardar = false;
        std::string archivo_salida;

        if (__argc >= 4 && std::string(__argv[1]) == "-sf") {
            // Leer texto desde archivo y guardar a WAV
            std::ifstream archivo(__argv[2]);
            if (!archivo.is_open()) {
                MessageBox(NULL, "No se pudo abrir el archivo", "Error", MB_ICONERROR);
                return -1;
            }
            texto = std::string((std::istreambuf_iterator<char>(archivo)),
                                std::istreambuf_iterator<char>());
            guardar = true;
            archivo_salida = __argv[3];

        } else if (__argc >= 4 && std::string(__argv[1]) == "-s") {
            // Leer texto directo y guardar a WAV
            texto = __argv[2];
            guardar = true;
            archivo_salida = __argv[3];

        } else if (__argc >= 3 && std::string(__argv[1]) == "-f") {
            // Solo leer desde archivo, no guardar
            std::ifstream archivo(__argv[2]);
            if (!archivo.is_open()) {
                MessageBox(NULL, "No se pudo abrir el archivo", "Error", MB_ICONERROR);
                return -1;
            }
            texto = std::string((std::istreambuf_iterator<char>(archivo)),
                                std::istreambuf_iterator<char>());

        } else {
            // Leer frase desde línea de comandos
            for (int i = 1; i < __argc; ++i) {
                if (i > 1) texto += " ";
                texto += __argv[i];
            }
        }

        if (guardar) {
            ::CoInitialize(NULL); // ← NECESARIO para usar COM
            std::wstring wtexto = StringToWString(texto);
            std::wstring warchivo = StringToWString(archivo_salida);
            GuardarTextoComoWav(wtexto, warchivo);
            ::CoUninitialize();
        } else {
            ::CoInitialize(NULL);
            HRESULT hr = CoCreateInstance(CLSID_SpVoice, NULL, CLSCTX_ALL, IID_ISpVoice, (void**)&pVoice);
            if (SUCCEEDED(hr)) {
                LeerTexto(texto);
                pVoice->Release();
            }
            ::CoUninitialize();
        }

        return 0;
    }

    // Modo ventana (sin argumentos)
    ::CoInitialize(NULL);
    HRESULT hr = CoCreateInstance(CLSID_SpVoice, NULL, CLSCTX_ALL,
                                  IID_ISpVoice, (void**)&pVoice);

    if (FAILED(hr)) {
        MessageBox(NULL, "No se pudo inicializar SAPI", "Error", MB_ICONERROR);
        return -1;
    }

    WNDCLASS wc = {0};
    wc.lpfnWndProc = WndProc;
    wc.hInstance = hInstance;
    wc.hbrBackground = CreateSolidBrush(RGB(255, 140, 0)); // Dark Orange
    wc.lpszClassName = "Ventana";

    RegisterClass(&wc);

    HWND hwnd = CreateWindow("Ventana", "Apuntes de de Texto a Voz (.wav)",
                             WS_OVERLAPPEDWINDOW | WS_VISIBLE,
                             100, 100, 420, 200, NULL, NULL, hInstance, NULL);

    MessageBox(
        NULL,
        "Modo Consola (opcional):\n\n"
        "  Leer archivo de texto:\n"
        "    Programa.exe -f archivo.txt\n\n"
        "  Leer texto directo:\n"
        "    Programa.exe frase a leer...\n\n"
        "  Guardar texto como WAV:\n"
        "    Programa.exe -s \"texto\" salida.wav\n\n"
        "  Guardar archivo como WAV:\n"
        "    Programa.exe -sf archivo.txt salida.wav\n",
        "Instrucciones de Consola",
        MB_OK | MB_ICONINFORMATION
    );

    MSG msg = {0};
    while (GetMessage(&msg, NULL, 0, 0)) {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }

    CoUninitialize();
    return 0;
}
