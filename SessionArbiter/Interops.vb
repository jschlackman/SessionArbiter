Public Class Interops

    'Win32 constants
    Public Const WM_POWERBROADCAST As Integer = &H218
    Public Const WM_QUERYENDSESSION As Integer = &H11
    Public Const WM_QUIT As Integer = &H12
    Public Const WM_CLOSE As Integer = &H10
    Public Const PBT_POWERSETTINGCHANGE As Integer = &H8013
    Public Const DEVICE_NOTIFY_WINDOW_HANDLE As Integer = &H0

    ''' <summary>
    ''' GUID value to request receipt of lid close messages.
    ''' </summary>
    ''' <remarks></remarks>
    Public Shared GUID_LIDCLOSE_ACTION As New Guid("{0xBA3E0F4D, 0xB817, 0x4094, {0xA2, 0xD1, 0xD5, 0x63, 0x79, 0xE6, 0xA0, 0xF3}}")
    Public Shared GUID_POWERSCHEME_PERSONALITY As New Guid("{0x245D8541, 0x3943, 0x4422, {0xB0, 0x25, 0x13, 0xA7, 0x84, 0xF6, 0x79, 0xB7}}")

    Declare Function PostMessage Lib "user32.dll" (ByVal hWnd As IntPtr, ByVal Msg As UInteger, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As Boolean

    ''' <summary>
    ''' Registers the application to receive power setting notifications for the specific power setting event.
    ''' </summary>
    ''' <param name="hRecipient">Handle indicating where the power setting notifications are to be sent. For interactive applications, the Flags parameter should be zero, and the hRecipient parameter should be a window handle. For services, the Flags parameter should be one, and the hRecipient parameter should be a SERVICE_STATUS_HANDLE as returned from RegisterServiceCtrlHandlerEx.</param>
    ''' <param name="PowerSettingGuid">The GUID of the power setting for which notifications are to be sent.</param>
    ''' <param name="Flags">DEVICE_NOTIFY_WINDOW_HANDLE to send messages to a window handle, or DEVICE_NOTIFY_SERVICE_HANDLE to send to service.</param>
    ''' <returns>Notification handle for unregistering for power notifications. If the function fails, the return value is NULL. To get extended error information, call GetLastError.</returns>
    ''' <remarks>http://www.pinvoke.net/default.aspx/user32/RegisterPowerSettingNotification.html</remarks>
    Declare Function RegisterPowerSettingNotification Lib "user32.dll" (ByVal hRecipient As IntPtr, ByRef PowerSettingGuid As Guid, ByVal Flags As Int32) As IntPtr

    ''' <summary>
    ''' Unregisters the power setting notification.
    ''' </summary>
    ''' <param name="handle">The handle returned from the RegisterPowerSettingNotification function.</param>
    ''' <returns>If the function succeeds, the return value is nonzero. If the function fails, the return value is zero. To get extended error information, call GetLastError.</returns>
    ''' <remarks>Not currently used as program should terminates when the lid is closed and a logoff is triggered.</remarks>
    Declare Function UnregisterPowerSettingNotification Lib "user32.dll" (ByVal handle As IntPtr) As Boolean

    ''' <summary>
    ''' Logs off the interactive user, shuts down the system, or shuts down and restarts the system.
    ''' </summary>
    ''' <param name="uFlags">The shutdown type.</param>
    ''' <param name="dwReason">The reason for initiating the shutdown. This parameter must be one of the system shutdown reason codes.</param>
    ''' <returns>If the function succeeds, the return value is nonzero. Because the function executes asynchronously, a nonzero return value indicates that the shutdown has been initiated. It does not indicate whether the shutdown will succeed. It is possible that the system, the user, or another application will abort the shutdown. If the function fails, the return value is zero. To get extended error information, call GetLastError.</returns>
    ''' <remarks>http://msdn.microsoft.com/en-us/library/windows/desktop/aa376868%28v=vs.85%29.aspx</remarks>
    Declare Function ExitWindowsEx Lib "user32.dll" (ByVal uFlags As UInteger, ByVal dwReason As UInteger) As UInteger

    ''' <summary>
    ''' Shutdown types used by ExitWindowsEx in the uFlags parameter.
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum ExitWindows As UInteger
        LogOff = &H0
        ShutDown = &H1
        HybridShutdown = &H400000 'New for Windows 8
        Reboot = &H2
        PowerOff = &H8
        RestartApps = &H40
        'Note: only one of the following may be specified
        Force = &H4
        ForceIfHung = &H10
    End Enum

    ''' <summary>
    ''' The shutdown reason codes are used by the ExitWindowsEx and InitiateSystemShutdownEx functions in the dwReason parameter.
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum ShutdownReason As UInteger
        MajorApplication = &H40000
        MajorHardware = &H10000
        MajorLegacyApi = &H70000
        MajorOperatingSystem = &H20000
        MajorOther = &H0
        MajorPower = &H60000
        MajorSoftware = &H30000
        MajorSystem = &H50000

        MinorBlueScreen = &HF
        MinorCordUnplugged = &HB
        MinorDisk = &H7
        MinorEnvironment = &HC
        MinorHardwareDriver = &HD
        MinorHotfix = &H11
        MinorHung = &H5
        MinorInstallation = &H2
        MinorMaintenance = &H1
        MinorMMC = &H19
        MinorNetworkConnectivity = &H14
        MinorNetworkCard = &H9
        MinorOther = &H0
        MinorOtherDriver = &HE
        MinorPowerSupply = &HA
        MinorProcessor = &H8
        MinorReconfig = &H4
        MinorSecurity = &H13
        MinorSecurityFix = &H12
        MinorSecurityFixUninstall = &H18
        MinorServicePack = &H10
        MinorServicePackUninstall = &H16
        MinorTermSrv = &H20
        MinorUnstable = &H6
        MinorUpgrade = &H3
        MinorWMI = &H15

        FlagUserDefined = &H40000000
        FlagPlanned = &H80000000&
    End Enum

End Class
