Attribute VB_Name = "modWinVer"
'***************************************************
'윈도우 이름, 버전 이름, 그리고 버전 번호를 구하는 함수 입니다.
'MSDN을 참조해서 만들었습니다.
'제작 : 리바이
'수정 : YJSoft
'일시 : 2011.01.22
'***************************************************
 
Option Compare Binary
Option Explicit
 
'함수 선언부
 
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFOEX) As Long
Private Declare Sub GetSystemInfo Lib "kernel32" (lpSystemInfo As SYSTEM_INFO)
Private Declare Function GetProductInfo Lib "kernel32" (ByVal dwOSMajorVersion As Long, ByVal dwOSMinorVersion As Long, ByVal dwSpMajorVersion As Long, ByVal dwSpMinorVersion As Long, ByRef pdwReturnedProductType As Long) As Long
Private Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Const SM_SERVERR2 = 89
'구조체 타입 선언부
 
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
 
Private Type SYSTEM_INFO
    wProcessorArchitecture As Integer
    wReserved As Integer
    dwPageSize As Long
    lpMinimumApplicationAddress As Long
    lpMaximumApplicationAddress As Long
    dwActiveProcessorMask As Long
    dwNumberOfProcessors As Long
    dwProcessorType As Long
    dwAllocationGranularity As Long
    wProcessorLevel As Integer
    wProcessorRevision As Integer
End Type
 
'상수 선언부
 
Private Const VER_NT_WORKSTATION = &O1
Private Const VER_NT_DOMAIN_CONTROLLER = &H2
Private Const VER_NT_SERVER = &H3
 
Private Const VER_SUITE_BACKOFFICE = &H4 'Microsoft BackOffice
Private Const VER_SUITE_BLADE = &H400 'Windows Server 2003 Web Edition
Private Const VER_SUITE_COMPUTE_SERVER = &H4000 'Windows Server 2003 Compute Cluster Edition
Private Const VER_SUITE_DATACENTER = &H80 'Windows Server 2008 Datacenter, Windows Server 2003 Datacenter Edition, or Windows 2000 Datacenter Server
Private Const VER_SUITE_ENTERPRISE = &H2 'Windows Server 2008 Enterprise, Windows Server 2003 Enterprise Edition, or Windows 2000 Advanced Server
Private Const VER_SUITE_EMBEDDEDNT = &H40 'Windows XP Embedded
Private Const VER_SUITE_PERSONAL = &H200 'Windows Vista Home Premium, Windows Vista Home Basic, or Windows XP Home Edition
Private Const VER_SUITE_SINGLEUSERTS = &H100 'Remote Desktop
Private Const VER_SUITE_SMALLBUSINESS = &H1 'Microsoft Small Business Server
Private Const VER_SUITE_SMALLBUSINESS_RESTRICTED = &H20 'Microsoft Small Business Server
Private Const VER_SUITE_STORAGE_SERVER = &H2000 'Windows Storage Server 2003 R2 or Windows Storage Server 2003
Private Const VER_SUITE_TERMINAL = &H10 'Terminal Services
Private Const VER_SUITE_WH_SERVER = &H8000 'Windows Home Server
 
Private Const VER_PLATFORM_WIN32s = 0
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VER_PLATFORM_WIN32_NT = 2
 
Private Const PROCESSOR_ARCHITECTURE_AMD64 = 9 'x64 (AMD Or Intel)
Private Const PROCESSOR_ARCHITECTURE_IA64 = 6 'Intel Itanium - based
Private Const PROCESSOR_ARCHITECTURE_INTEL = 0 'x86
Private Const PROCESSOR_ARCHITECTURE_UNKNOWN = &HFFFF 'Unknown
 
Private Const PRODUCT_BUSINESS = &H6 'Business
Private Const PRODUCT_BUSINESS_N = &H10 'Business N
Private Const PRODUCT_CLUSTER_SERVER = &H12 'Cluster Server Edition
Private Const PRODUCT_DATACENTER_SERVER = &H8 'Server Datacenter Edition (full installation)
Private Const PRODUCT_DATACENTER_SERVER_CORE = &HC 'Server Datacenter Edition (core installation)
Private Const PRODUCT_DATACENTER_SERVER_CORE_V = &H27 'Server Datacenter Edition without Hyper-V (core installation)
Private Const PRODUCT_DATACENTER_SERVER_V = &H25 'Server Datacenter Edition without Hyper-V (full installation)
Private Const PRODUCT_ENTERPRISE = &H4 'Enterprise
Private Const PRODUCT_ENTERPRISE_E = &H46 'Not supported
Private Const PRODUCT_ENTERPRISE_N = &H1B 'Enterprise N
Private Const PRODUCT_ENTERPRISE_SERVER = &HA 'Server Enterprise Edition (full installation)
Private Const PRODUCT_ENTERPRISE_SERVER_CORE = &HE 'Server Enterprise Edition (core installation)
Private Const PRODUCT_ENTERPRISE_SERVER_CORE_V = &H29 'Server Enterprise Edition without Hyper-V (core installation)
Private Const PRODUCT_ENTERPRISE_SERVER_IA64 = &HF 'Server Enterprise Edition for Itanium-based Systems
Private Const PRODUCT_ENTERPRISE_SERVER_V = &H26 'Server Enterprise Edition without Hyper-V (full installation)
Private Const PRODUCT_HOME_BASIC = &H2 'Home Basic
Private Const PRODUCT_HOME_BASIC_E = &H43 'Not supported
Private Const PRODUCT_HOME_BASIC_N = &H5 'Home Basic N
Private Const PRODUCT_HOME_PREMIUM = &H3 'Home Premium
Private Const PRODUCT_HOME_PREMIUM_E = &H44 'Not supported
Private Const PRODUCT_HOME_PREMIUM_N = &H1A 'Home Premium N
Private Const PRODUCT_HOME_PREMIUM_SERVER = &H22 'Windows Home Server 2011
Private Const PRODUCT_HOME_SERVER = &H13 'Windows Storage Server 2008 R2 Essentials
Private Const PRODUCT_HYPERV = &H2A 'Microsoft Hyper-V Server
Private Const PRODUCT_MEDIUMBUSINESS_SERVER_MANAGEMENT = &H1E 'Windows Essential Business Server Management Server
Private Const PRODUCT_MEDIUMBUSINESS_SERVER_MESSAGING = &H20 'Windows Essential Business Server Messaging Server
Private Const PRODUCT_MEDIUMBUSINESS_SERVER_SECURITY = &H1F 'Windows Essential Business Server Security Server
Private Const PRODUCT_PROFESSIONAL = &H30 'Professional
Private Const PRODUCT_PROFESSIONAL_E = &H45 'Not supported
Private Const PRODUCT_PROFESSIONAL_N = &H31 'Professional N
Private Const PRODUCT_SB_SOLUTION_SERVER = &H32 'Windows Small Business Server 2011 Essentials
Private Const PRODUCT_SERVER_FOR_SMALLBUSINESS = &H18 'Windows Server 2008 for Windows Essential Server Solutions
Private Const PRODUCT_SERVER_FOR_SMALLBUSINESS_V = &H23 'Windows Server 2008 without Hyper-V for Windows Essential Server Solutions
Private Const PRODUCT_SERVER_FOUNDATION = &H21 'Server Foundation
Private Const PRODUCT_SMALLBUSINESS_SERVER = &H9 'Small Business Server
Private Const PRODUCT_SMALLBUSINESS_SERVER_PREMIUM = &H19 'Small Business Server Premium Edition
Private Const PRODUCT_SOLUTION_EMBEDDEDSERVER = &H38 'Windows MultiPoint Server
Private Const PRODUCT_STANDARD_SERVER = &H7 'Server Standard Edition (full installation)
Private Const PRODUCT_STANDARD_SERVER_CORE = &HD 'Server Standard Edition (core installation)
Private Const PRODUCT_STANDARD_SERVER_CORE_V = &H28 'Server Standard Edition without Hyper-V (core installation)
Private Const PRODUCT_STANDARD_SERVER_V = &H24 'Server Standard Edition without Hyper-V (full installation)
Private Const PRODUCT_STARTER = &HB 'Starter
Private Const PRODUCT_STARTER_E = &H42 'Not supported
Private Const PRODUCT_STARTER_N = &H2F 'Starter N
Private Const PRODUCT_STORAGE_ENTERPRISE_SERVER = &H17 'Storage Server Enterprise Edition
Private Const PRODUCT_STORAGE_EXPRESS_SERVER = &H14 'Storage Server Express Edition
Private Const PRODUCT_STORAGE_STANDARD_SERVER = &H15 'Storage Server Standard Edition
Private Const PRODUCT_STORAGE_WORKGROUP_SERVER = &H16 'Storage Server Workgroup Edition
Private Const PRODUCT_ULTIMATE = &H1 'Ultimate
Private Const PRODUCT_ULTIMATE_E = &H47 'Not supported
Private Const PRODUCT_ULTIMATE_N = &H1C 'Ultimate N
Private Const PRODUCT_UNDEFINED = &H0 'An unknown product
Private Const PRODUCT_WEB_SERVER = &H11 'Web Server Edition (full installation)
Private Const PRODUCT_WEB_SERVER_CORE = &H1D 'Web Server Edition (core installation)
 
'힘수 본체
 
Public Function fGetWindowVersion() As String
    Dim tVersionInfo As OSVERSIONINFOEX
    Dim tSysemInfo As SYSTEM_INFO
    Dim sOSClass As String 'OS 분류
    Dim sOSName As String 'OS 이름 : (예) Windows 7
    Dim sOSVersionName As String 'OS 버전 이름 : (예) Ultimate
    Dim sServicePackName As String '서비스 팩 이름 : (예) Sevice Pack 1
    Dim sCoreNumbers As String '코어 수 : (예) Dual-Core
    Dim sFullVersion As String '풀 버전 : (예) 6.1.7601
    Dim iVersionName As Long
    Dim iReturn As Long
    
    Call ZeroMemory(tVersionInfo, Len(tVersionInfo))
    Call ZeroMemory(tSysemInfo, Len(tSysemInfo))
    
    tVersionInfo.dwOSVersionInfoSize = Len(tVersionInfo)
    
    iReturn = GetVersionEx(tVersionInfo)
    If (iReturn = 0) Then
        fGetWindowVersion = "Unknown"
        Exit Function
    End If
    
    Select Case tVersionInfo.dwPlatformId
        Case VER_PLATFORM_WIN32s: 'WIN32s(3.1)
            sOSClass = "Windows 32s"
        Case VER_PLATFORM_WIN32_WINDOWS: '9x
            sOSClass = "Windows 95/98/ME"
        Case VER_PLATFORM_WIN32_NT: 'NT
            sOSClass = "Windows NT"
    End Select
        
    Call GetSystemInfo(tSysemInfo)
    
    Select Case tSysemInfo.dwNumberOfProcessors
        Case 1: sCoreNumbers = "Single-Core"
        Case 2: sCoreNumbers = "Dual-Core"
        Case 3: sCoreNumbers = "Triple-Core"
        Case 4: sCoreNumbers = "Quad-Core"
        Case 6: sCoreNumbers = "Hexa-Core"
        Case 8: sCoreNumbers = "Octa-Core"
        Case 12: sCoreNumbers = "Magni-Core"
        Case Else: sCoreNumbers = tSysemInfo.dwNumberOfProcessors & "-Core"
    End Select
    
    Select Case tVersionInfo.dwMajorVersion
        Case 3: 'WINNT 구버전
            Select Case tVersionInfo.dwMinorVersion
                Case 0: sOSName = "Windows NT3"
                Case 1: sOSName = "Windows NT 3.1"
                Case 5: sOSName = "Windows NT 3.5"
                Case 51: sOSName = "Windows NT 3.51"
            End Select
        Case 4: '9x,NT4
            Select Case tVersionInfo.dwPlatformId
                Case VER_PLATFORM_WIN32_WINDOWS: '9x
                    Select Case tVersionInfo.dwMinorVersion
                        Case 0: sOSName = "Windows 95"
                        Case 10: sOSName = "Windows 98"
                        Case 90: sOSName = "Windows ME"
                        Case Else: sOSName = "Windows 98"
                    End Select
                Case VER_PLATFORM_WIN32_NT: 'NT4
                    If (tVersionInfo.wProductType = VER_NT_WORKSTATION) Then
                        sOSVersionName = "Worksation"
                    ElseIf (tVersionInfo.wProductType = VER_NT_SERVER) Then
                        If (tVersionInfo.wSuiteMask & VER_SUITE_ENTERPRISE) Then
                            sOSVersionName = "Server Enterprise"
                        Else
                            sOSVersionName = "Server Standard"
                        End If
                    End If
                   
                    sOSName = "Windows NT4" & " " & sOSVersionName
            End Select
        Case 5: 'WINNT4 이상
            IsAboveNT = True '투명화 사용 가능
            Select Case tVersionInfo.dwMinorVersion
                Case 0: 'Windows 2000
                    If (tVersionInfo.wProductType = VER_NT_WORKSTATION) Then
                        sOSVersionName = "Professional"
                    Else
                        If (tVersionInfo.wSuiteMask & VER_SUITE_DATACENTER) Then
                            sOSVersionName = "Datacenter Server"
                        ElseIf (tVersionInfo.wSuiteMask & VER_SUITE_ENTERPRISE) Then
                            sOSVersionName = "Advanced Server"
                        Else
                            sOSVersionName = "Server"
                        End If
                    End If
                    sOSName = "Windows 2000" & " " & sOSVersionName
                Case 1: 'Windows XP
                    If (tVersionInfo.wSuiteMask And VER_SUITE_PERSONAL = VER_SUITE_PERSONAL) Then
                        sOSVersionName = "Home Edition"
                    Else
                        sOSVersionName = "Professional"
                    End If
                    sOSName = "Windows XP" & " " & sOSVersionName
                Case 2:
                    If (tVersionInfo.wProductType <> VER_NT_WORKSTATION) Then
                        If (tSysemInfo.wProcessorArchitecture = PROCESSOR_ARCHITECTURE_IA64) Then
                            If (tVersionInfo.wSuiteMask & VER_SUITE_DATACENTER) Then
                                sOSVersionName = "Datacenter Edition for Itanium-based Systems"
                            ElseIf (tVersionInfo.wSuiteMask & VER_SUITE_ENTERPRISE) Then
                                sOSVersionName = "Enterprise Edition for Itanium-based Systems"
                            End If
                        ElseIf (tSysemInfo.wProcessorArchitecture = PROCESSOR_ARCHITECTURE_AMD64) Then
                            If (tVersionInfo.wSuiteMask & VER_SUITE_DATACENTER) Then
                                sOSVersionName = "Datacenter x64 Edition"
                            ElseIf (tVersionInfo.wSuiteMask & VER_SUITE_ENTERPRISE) Then
                                sOSVersionName = "Enterprise x64 Edition"
                            Else
                                sOSVersionName = "Standard x64 Edition"
                            End If
                        Else
                            If (tVersionInfo.wSuiteMask & VER_SUITE_COMPUTE_SERVER) Then
                                sOSVersionName = "Compute Cluster Edition"
                            ElseIf (tVersionInfo.wSuiteMask & VER_SUITE_DATACENTER) Then
                                sOSVersionName = "Datacenter Edition"
                            ElseIf (tVersionInfo.wSuiteMask & VER_SUITE_ENTERPRISE) Then
                                sOSVersionName = "Enterprise Edition"
                            ElseIf (tVersionInfo.wSuiteMask & VER_SUITE_BLADE) Then
                                sOSVersionName = "Web Edition"
                            Else
                                sOSVersionName = "Standard Edition"
                            End If
                        End If
                    End If
                        
                    If (GetSystemMetrics(SM_SERVERR2)) Then
                        sOSName = "Windows Server 2003 R2" & " " & sOSVersionName
                    ElseIf (tVersionInfo.wSuiteMask & VER_SUITE_STORAGE_SERVER) Then
                        sOSName = "Windows Storage Server 2003" & " " & sOSVersionName
                    ElseIf (tVersionInfo.wSuiteMask & VER_SUITE_WH_SERVER) Then
                        sOSName = "Windows Home Server" & " " & sOSVersionName
                    ElseIf (tVersionInfo.wProductType = VER_NT_WORKSTATION And tSysemInfo.wProcessorArchitecture = PROCESSOR_ARCHITECTURE_AMD64) Then
                        sOSName = "Windows XP Professional x64 Edition"
                    Else
                        sOSName = "Windows Server 2003" & " " & sOSVersionName
                    End If
            End Select
        Case 6:
            IsAboveNT = True 'Bugfix:issue6
            iVersionName = GetProductInfo(tVersionInfo.dwMajorVersion, tVersionInfo.dwMinorVersion, 0, 0, iVersionName)
            
            Select Case iVersionName
                Case PRODUCT_ULTIMATE:
                    sOSVersionName = "Ultimate Edition"
                Case PRODUCT_PROFESSIONAL:
                    sOSVersionName = "Professional"
                Case PRODUCT_HOME_PREMIUM:
                    sOSVersionName = "Home Premium Edition"
                Case PRODUCT_HOME_BASIC:
                    sOSVersionName = "Home Basic Edition"
                Case PRODUCT_ENTERPRISE:
                    sOSVersionName = "Enterprise Edition"
                Case PRODUCT_BUSINESS:
                    sOSVersionName = "Business Edition"
                Case PRODUCT_STARTER:
                    sOSVersionName = "Starter Edition"
                Case PRODUCT_CLUSTER_SERVER:
                    sOSVersionName = "Cluster Server Edition"
                Case PRODUCT_DATACENTER_SERVER:
                    sOSVersionName = "Datacenter Edition"
                Case PRODUCT_DATACENTER_SERVER_CORE:
                    sOSVersionName = "Datacenter Edition (Core Installation)"
                Case PRODUCT_ENTERPRISE_SERVER:
                    sOSVersionName = "Enterprise Edition"
                Case PRODUCT_ENTERPRISE_SERVER_CORE:
                    sOSVersionName = "Enterprise Edition (Core Installation)"
                Case PRODUCT_ENTERPRISE_SERVER_IA64:
                    sOSVersionName = "Enterprise Edition for Itanium-based Systems"
                Case PRODUCT_SMALLBUSINESS_SERVER:
                    sOSVersionName = "Small Business Server"
                Case PRODUCT_SMALLBUSINESS_SERVER_PREMIUM:
                    sOSVersionName = "Small Business Server Premium Edition"
                Case PRODUCT_STANDARD_SERVER:
                    sOSVersionName = "Standard Edition"
                Case PRODUCT_STANDARD_SERVER_CORE:
                    sOSVersionName = "Standard Edition (Core Installation)"
                Case PRODUCT_WEB_SERVER:
                    sOSVersionName = "Web Server Edition"
                Case Else: sOSVersionName = ""
            End Select
            
            If (tSysemInfo.wProcessorArchitecture = PROCESSOR_ARCHITECTURE_AMD64) Then
                If (sOSVersionName <> "") Then
                    sOSVersionName = sOSVersionName & " "
                End If
                sOSVersionName = sOSVersionName & "64-bit"
            ElseIf (tSysemInfo.wProcessorArchitecture = PROCESSOR_ARCHITECTURE_INTEL) Then
                If (sOSVersionName <> "") Then
                    sOSVersionName = sOSVersionName & " "
                End If
                sOSVersionName = sOSVersionName & "32-bit"
            End If
    
            Select Case tVersionInfo.dwMinorVersion
                Case 0:
                    If (tVersionInfo.wProductType = VER_NT_WORKSTATION) Then
                        sOSName = "Windows Vista" & " " & sOSVersionName
                    Else
                        sOSName = "Windows Server 2008" & " " & sOSVersionName
                    End If
                Case 1:
                    If (tVersionInfo.wProductType = VER_NT_WORKSTATION) Then
                        sOSName = "Windows 7" & " " & sOSVersionName
                    Else
                        sOSName = "Windows Server 2008 R2" & " " & sOSVersionName
                    End If
                Case 2:
                    sOSName = "Windows 8" & " " & sOSVersionName
                    
                Case 3:
                    sOSName = "Windows 9" & " " & sOSVersionName
            End Select
    End Select
        
    '*******************
    '아래와 결과가 같다.
    '*******************
    If (0) Then
    If (InStr(tVersionInfo.szCSDVersion, Chr(0)) <> 0) Then
        sServicePackName = Left(tVersionInfo.szCSDVersion, InStr(tVersionInfo.szCSDVersion, Chr(0)) - 1)
    Else
        sServicePackName = tVersionInfo.szCSDVersion
    End If
    
    If (sServicePackName <> "") Then
        sOSName = sOSName & " (" & sServicePackName & ")"
    End If
    End If
    
    If (1) Then
    If (tVersionInfo.wServicePackMajor > 0) Then
        sServicePackName = " (Service Pack " & tVersionInfo.wServicePackMajor & ")"
        sOSName = sOSName & sServicePackName
    End If
    End If
    
    sFullVersion = "Version : " & tVersionInfo.dwMajorVersion & "." & tVersionInfo.dwMinorVersion & "." & tVersionInfo.dwBuildNumber
    
    fGetWindowVersion = sOSName & vbCrLf & sFullVersion & vbCrLf & sCoreNumbers
End Function
