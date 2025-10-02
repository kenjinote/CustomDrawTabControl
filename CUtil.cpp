#include "CUtil.h"
#include <Windows.h>

bool CUtil::IsSystemInDarkTheme() {
    HKEY hKey;
    if (RegOpenKeyExW(HKEY_CURRENT_USER, L"Software\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize", 0, KEY_READ, &hKey) == ERROR_SUCCESS)
    {
        DWORD dwValue = 1;
        DWORD dwSize = sizeof(dwValue);
        if (RegQueryValueExW(hKey, L"AppsUseLightTheme", NULL, NULL, (LPBYTE)&dwValue, &dwSize) == ERROR_SUCCESS)
        {
            RegCloseKey(hKey);
            return dwValue == 0;
        }
        RegCloseKey(hKey);
    }
    return false;
}
