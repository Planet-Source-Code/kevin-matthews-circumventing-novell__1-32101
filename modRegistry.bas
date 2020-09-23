Attribute VB_Name = "modRegistry"
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As Any, ByVal fuWinIni As Long) As Long
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const REG_SZ = 1                         ' Unicode nul terminated string
Public Const REG_DWORD = 4                      ' 32-bit number
Public Const ERROR_SUCCESS = 0&
Public Declare Function GetUserNameAPI Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long 'Get the username
Public Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
'--------------------------------------------------
'--------------------------------------------------
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long

Public Declare Function FindWindowEx Lib "user32" _
    Alias "FindWindowExA" (ByVal hWnd1 As Long, _
    ByVal hWnd2 As Long, ByVal lpsz1 As String, _
    ByVal lpsz2 As String) As Long
   Public Const SPI_SCREENSAVERRUNNING = 97
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, _
    ByVal nCmdShow As Long) As Long

Public Sub CreateKey(hKey As Long, strPath As String)

  Dim hCurKey As Long
  Dim lRegResult As Long

    lRegResult = RegCreateKey(hKey, strPath, hCurKey)

    If lRegResult <> ERROR_SUCCESS Then
        ' there is a problem
    End If

    lRegResult = RegCloseKey(hCurKey)

End Sub

Public Function GetSettingString(hKey As Long, strPath As String, strValue As String, Optional Default As String) As String

  Dim hCurKey As Long
  Dim lValueType As Long
  Dim strBuffer As String
  Dim lDataBufferSize As Long
  Dim intZeroPos As Integer
  Dim lRegResult As Long

    ' Set up default value
    If Not IsEmpty(Default) Then
        GetSettingString = Default
      Else
        GetSettingString = ""
    End If

    ' Open the key and get length of string
    lRegResult = RegOpenKey(hKey, strPath, hCurKey)
    lRegResult = RegQueryValueEx(hCurKey, strValue, 0&, lValueType, ByVal 0&, lDataBufferSize)

    If lRegResult = ERROR_SUCCESS Then

        If lValueType = REG_SZ Then
            ' initialise string buffer and retrieve string
            strBuffer = String$(lDataBufferSize, " ")
            lRegResult = RegQueryValueEx(hCurKey, strValue, 0&, 0&, ByVal strBuffer, lDataBufferSize)

            ' format string
            intZeroPos = InStr(strBuffer, Chr$(0))
            If intZeroPos > 0 Then
                GetSettingString = Left$(strBuffer, intZeroPos - 1)
              Else
                GetSettingString = strBuffer
            End If

        End If

      Else
        ' there is a problem
    End If

    lRegResult = RegCloseKey(hCurKey)

End Function

Public Sub SaveSettingString(hKey As Long, strPath As String, strValue As String, strData As String)

  Dim hCurKey As Long
  Dim lRegResult As Long

    lRegResult = RegCreateKey(hKey, strPath, hCurKey)

    lRegResult = RegSetValueEx(hCurKey, strValue, 0, REG_SZ, ByVal strData, Len(strData))

    If lRegResult <> ERROR_SUCCESS Then
        'there is a problem
    End If

    lRegResult = RegCloseKey(hCurKey)

End Sub

Public Function GetSettingLong(ByVal hKey As Long, ByVal strPath As String, ByVal strValue As String, Optional Default As Long) As Long

  Dim lRegResult As Long
  Dim lValueType As Long
  Dim lBuffer As Long
  Dim lDataBufferSize As Long
  Dim hCurKey As Long

    ' Set up default value
    If Not IsEmpty(Default) Then
        GetSettingLong = Default
      Else
        GetSettingLong = 0
    End If

    lRegResult = RegOpenKey(hKey, strPath, hCurKey)
    lDataBufferSize = 4       ' 4 bytes = 32 bits = long

    lRegResult = RegQueryValueEx(hCurKey, strValue, 0&, lValueType, lBuffer, lDataBufferSize)

    If lRegResult = ERROR_SUCCESS Then

        If lValueType = REG_DWORD Then
            GetSettingLong = lBuffer
        End If

      Else
        'there is a problem
    End If

    lRegResult = RegCloseKey(hCurKey)

End Function

Public Sub SaveSettingLong(ByVal hKey As Long, ByVal strPath As String, ByVal strValue As String, ByVal lData As Long)

  Dim hCurKey As Long
  Dim lRegResult As Long

    lRegResult = RegCreateKey(hKey, strPath, hCurKey)

    lRegResult = RegSetValueEx(hCurKey, strValue, 0&, REG_DWORD, lData, 4)

    If lRegResult <> ERROR_SUCCESS Then
        'there is a problem
    End If

    lRegResult = RegCloseKey(hCurKey)

End Sub


