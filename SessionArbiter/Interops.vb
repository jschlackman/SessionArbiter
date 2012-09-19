Public Class Interops

    'Win32 constants
    Public Const WM_POWERBROADCAST As Integer = &H218
    Public Const PBT_POWERSETTINGCHANGE As Integer = &H8013
    Public Const DEVICE_NOTIFY_WINDOW_HANDLE As Integer = &H0

    Public Structure POWERBROADCAST_SETTING
        Public PowerSetting As Guid
        Public DataLength As UInteger
        Public Data As Byte
    End Structure


    ''' <summary>
    ''' GUID value to request receipt of lid close messages.
    ''' </summary>
    ''' <remarks></remarks>

    Public Shared GUID_LIDSWITCH_STATE_CHANGE As New Guid("{0xBA3E0F4D, 0xB817, 0x4094, {0xA2, 0xD1, 0xD5, 0x63, 0x79, 0xE6, 0xA0, 0xF3}}")
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

End Class
