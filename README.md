<div align="center">

## fGetSystemInfo


</div>

### Description

MS stipulates that OS Version Info must be obtained "correctly" in their Windows2000 Application Specifications. This is the way.

It also uses api's to get the OS path, get the Windows Temp Dir and to generate a unique temp file name.

This is a .BAS file with a Sub Main() so it should compile easily. It generates the info, writes to a temp file and launches notepad with the info. No forms. You can easily hash through it to pull out what you need.
 
### More Info
 
None--

Just copy the entire source to a .bas file and launch. No forms needed. step through to pull out what you want.

Compiled under VB5/6 and ran on WinNT4 (server) and Windows2000 Professional. Don't know about the 9.x kernel, but it should be fine.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Brian Cidern](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/brian-cidern.md)
**Level**          |Advanced
**User Rating**    |4.7 (14 globes from 3 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/brian-cidern-fgetsysteminfo__1-6072/archive/master.zip)

### API Declarations

```
' used for dwPlatformId
Const VER_PLATFORM_WIN32s = 0
Const VER_PLATFORM_WIN32_WINDOWS = 1
Const VER_PLATFORM_WIN32_NT = 2
' used for wSuiteMask
Const VER_SUITE_BACKOFFICE = 4
Const VER_SUITE_DATACENTER = 128
Const VER_SUITE_ENTERPRISE = 2
Const VER_SUITE_SMALLBUSINESS = 1
Const VER_SUITE_SMALLBUSINESS_RESTRICTED = 32
Const VER_SUITE_TERMINAL = 16
' used for wProductType
Const VER_NT_WORKSTATION = 1
Const VER_NT_DOMAIN_CONTROLLER = 2
Const VER_NT_SERVER = 3
Const MAX_PATH = 260
Private Type OSVERSIONINFOEX
  dwOSVersionInfoSize As Long
  dwMajorVersion As Long
  dwMinorVersion As Long
  dwBuildNumber As Long
  dwPlatformId As Long
  szCSDVersion As String * 128
  wServicePackMajor As Integer
  wServicePackMinor As Integer
  wSuiteMask As Integer
  wProductType As Byte
  wReserved As Byte
End Type
Private Declare Function GetVersionEx _
  Lib "kernel32" _
  Alias "GetVersionExA" ( _
  lpVersionInformation As OSVERSIONINFOEX _
  ) As Long
Private Declare Function GetWindowsDirectory _
  Lib "kernel32" _
  Alias "GetWindowsDirectoryA" ( _
  ByVal lpBuffer As String, _
  ByVal nSize As Long _
  ) As Long
Private Declare Function GetTempPath _
  Lib "kernel32" _
  Alias "GetTempPathA" ( _
  ByVal nBufferLength As Long, _
  ByVal lpBuffer As String _
  ) As Long
Private Declare Function GetTempFileName _
  Lib "kernel32" _
  Alias "GetTempFileNameA" ( _
  ByVal lpszPath As String, _
  ByVal lpPrefixString As String, _
  ByVal wUnique As Long, _
  ByVal lpTempFileName As String _
  ) As Long
Dim sSystemInfo As String
Dim OSVI As OSVERSIONINFOEX
```


### Source Code

```
Function fGetTempFile() As String
  Dim sTempDir As String
  sTempDir = fDirCheck(fGetTempDir())
  Dim sPrefix As String
  sPrefix = ""
  Dim lUnique As Long
  lUnique = 0
  Dim lRet As Long
  Dim sBuf As String * 512
  lRet = GetTempFileName(sTempDir, sPrefix, lUnique, sBuf)
  If InStr(1, sBuf, Chr(0)) > 0 Then
    fGetTempFile = _
    Left(sBuf, InStr(1, sBuf, Chr(0)) - 1)
  Else
    fGetTempFile = ""
  End If
End Function
Function fGetWinDir() As String
  Dim lRet As Long
  Dim lSize As Long
  Dim sBuf As String * MAX_PATH
  lSize = MAX_PATH
  lRet = GetWindowsDirectory(ByVal sBuf, ByVal lSize)
  If InStr(1, sBuf, Chr(0)) > 0 Then
    fGetWinDir = Left(sBuf, InStr(1, sBuf, Chr(0)) - 1)
  Else
    fGetWinDir = ""
  End If
End Function
Function fDirCheck(sDirName As String) As String
  fDirCheck = IIf(Right(sDirName, 1) = "\", _
  sDirName, sDirName & "\")
End Function
Function fGetTempDir() As String
  Dim lRet As Long
  Dim lSize As Long
  Dim sBuf As String * MAX_PATH
  lSize = MAX_PATH
  lRet = GetTempPath(ByVal lSize, sBuf)
  If InStr(1, sBuf, Chr(0)) > 0 Then
    fGetTempDir = Left(sBuf, InStr(2, sBuf, Chr(0)) - 1)
  Else
    fGetTempDir = ""
  End If
End Function
Function fGetSystemInfo() As Boolean
  Dim lRet As Long
  Dim iNullPos As Integer
  Dim colProdSuites As Collection
  Dim vCurrProdSuite As Variant
  OSVI.dwOSVersionInfoSize = Len(OSVI)
  OSVI.szCSDVersion = Space(128)
  lRet = GetVersionEx(OSVI)
  If lRet = 0 Then
    MsgBox ("Error" & vbCrLf & _
        Err.LastDllError & " - " & Err.Description)
    fGetSystemInfo = False
    Exit Function
  End If
  ' For major version number, minor version number,
  ' and build number, convert the value returned into
  ' a string.
  sSystemInfo = "Major Version: " & _
         Str(OSVI.dwMajorVersion) & vbCrLf
  sSystemInfo = sSystemInfo + "Minor Version: " & _
         Str(OSVI.dwMinorVersion) & vbCrLf
  sSystemInfo = sSystemInfo + "Build Number: " & _
         Str(OSVI.dwBuildNumber) & vbCrLf
  ' To determine the specific platform, use the
  ' constants you declared to evaluate dwPlatformId.
  ' Depending on the platform, check dwBuildNumber
  ' to determine the specific platform.
  sSystemInfo = sSystemInfo + "Platform: "
  Select Case OSVI.dwPlatformId
    Case VER_PLATFORM_WIN32s
      sSystemInfo = sSystemInfo & _
             "Win32s on Windows 3.1" & vbCrLf
    Case VER_PLATFORM_WIN32_WINDOWS
      sSystemInfo = sSystemInfo & _
      IIf(OSVI.dwBuildNumber = 0, _
      "Windows 98", "Windows 95") & vbCrLf
    Case VER_PLATFORM_WIN32_NT
      sSystemInfo = sSystemInfo & _
      IIf(OSVI.dwMajorVersion < 5, _
      "Windows NT", "Windows 2000") & vbCrLf
  End Select
  ' To determine service pack information, use the
  ' constants you declared to evaluate dwPlatformId.
  ' Depending on the platform, check szCSDVersion
  ' to determine the specific service pack information.
  Select Case OSVI.dwPlatformId
    Case VER_PLATFORM_WIN32s
      sSystemInfo = sSystemInfo & _
             "No additional info on " & _
             "Win32s on Windows 3.1." & vbCrLf
    Case VER_PLATFORM_WIN32_WINDOWS
      sSystemInfo = sSystemInfo & _
             "Additional OS Info: " & _
             OSVI.szCSDVersion & vbCrLf
    Case VER_PLATFORM_WIN32_NT
      If Asc(Left$(OSVI.szCSDVersion, 1)) = 0 Then
        ' leftmost char = null, this is an
        ' empty string
        sSystemInfo = sSystemInfo & _
               "Service Pack Install " & _
               "Info: No Service Pack " & _
               "Installed" & vbCrLf
      Else
        ' find the null char in the string
        iNullPos = InStr(OSVI.szCSDVersion, Chr(0))
        sSystemInfo = sSystemInfo & _
               "Service Pack Install " & _
               "Info: " & _
               Left$(OSVI.szCSDVersion, _
               iNullPos - 1) & vbCrLf
      End If
  End Select
  ' For major service pack, major and minor
  ' version numbers, convert the values returned
  ' into a string.
  sSystemInfo = sSystemInfo & "Service Pack Version: "
  sSystemInfo = sSystemInfo & _
         CStr(OSVI.wServicePackMajor) & "." & _
         CStr(OSVI.wServicePackMinor) & vbCrLf
  ' To determine which product suite components are
  ' installed evaluate wSuiteMask and compare the value
  ' against the constants declared for the various
  ' product suites. Add information to the colProdSuite
  ' collection based on which product suites are installed.
  ' This this value is a set of bit flags. Test against
  ' each bit mask, add found items to a VB collection
  Set colProdSuites = New Collection
  If (OSVI.wSuiteMask And VER_SUITE_BACKOFFICE) = VER_SUITE_BACKOFFICE Then
    colProdSuites.Add "Microsoft BackOffice components are installed."
  End If
  If (OSVI.wSuiteMask And VER_SUITE_DATACENTER) = VER_SUITE_DATACENTER Then
    colProdSuites.Add "Windows 2000 Datacenter Server is installed."
  End If
  If (OSVI.wSuiteMask And VER_SUITE_ENTERPRISE) = VER_SUITE_ENTERPRISE Then
    colProdSuites.Add "Windows 2000 Advanced Server is installed."
  End If
  If (OSVI.wSuiteMask And VER_SUITE_SMALLBUSINESS) = VER_SUITE_SMALLBUSINESS Then
    colProdSuites.Add "Microsoft Small Business Server is installed."
  End If
  If (OSVI.wSuiteMask And VER_SUITE_SMALLBUSINESS_RESTRICTED) = VER_SUITE_SMALLBUSINESS_RESTRICTED Then
    colProdSuites.Add "Microsoft Small Business Server is installed " & "with the restrictive client license in force."
  End If
  If (OSVI.wSuiteMask And VER_SUITE_TERMINAL) = VER_SUITE_TERMINAL Then
    colProdSuites.Add "Terminal Services is installed."
  End If
  ' list all product suites available
  ' that were added to the collection object
  sSystemInfo = sSystemInfo & "Product Suites: " & vbCrLf
  For Each vCurrProdSuite In colProdSuites
    sSystemInfo = sSystemInfo & vbCrLf & vbTab & vCurrProdSuite
  Next
  ' To determine the product type, use the constants you declared to
  ' evaluate wProductType.
  sSystemInfo = sSystemInfo & "Product Type: "
  Select Case OSVI.wProductType
    Case VER_NT_WORKSTATION
      sSystemInfo = sSystemInfo & "Windows 2000 Professional"
    Case VER_NT_DOMAIN_CONTROLLER
      sSystemInfo = sSystemInfo & "Windows 2000 domain controller"
    Case VER_NT_SERVER
      sSystemInfo = sSystemInfo & "Windows 2000 Server"
  End Select
  fGetSystemInfo = True
End Function
Sub Main()
  If fGetSystemInfo() Then
    Dim sTmpFile As String
    sTmpFile = fGetTempFile
    Open sTmpFile For Output As #1
      Print #1, sSystemInfo
    Close #1
    Dim sCmd As String
    sCmd = fDirCheck(fGetWinDir()) & "Notepad.exe " & sTmpFile
    Dim vRet As Variant
    vRet = Shell(sCmd, vbNormalFocus)
  End If
End Sub
```

