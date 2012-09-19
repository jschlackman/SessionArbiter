Imports SessionArbiter.Interops
Imports System.Windows.Forms

''' <summary>
''' Listener to process power notification messages.
''' </summary>
''' <remarks></remarks>
Class LidListener
    Inherits NativeWindow

    ''' <summary>
    ''' Delegate type for lid close event
    ''' </summary>
    ''' <remarks></remarks>
    Public Delegate Sub OnLidCloseEvent()

    ''' <summary>
    ''' Internal variable for lid close delegate
    ''' </summary>
    ''' <remarks></remarks>
    Private dLidClose As OnLidCloseEvent

    ''' <summary>
    ''' Internal variable to store the lid event callback handle.
    ''' </summary>
    ''' <remarks></remarks>
    Private hCallback As IntPtr

    ''' <summary>
    ''' Stores the date and time the program started listening for messages.
    ''' </summary>
    ''' <remarks></remarks>
    Public Started As DateTime

    ''' <summary>
    ''' Whether to force the logoff (default is true).
    ''' </summary>
    ''' <remarks></remarks>
    Private ForceLogoff As Boolean

    ''' <summary>
    ''' Register to receive lid close notifications.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub RegisterForPowerNotifications()
        hCallback = RegisterPowerSettingNotification(Me.Handle, GUID_LIDCLOSE_ACTION, DEVICE_NOTIFY_WINDOW_HANDLE)
    End Sub

    ''' <summary>
    ''' Unregister the application from receiving lid close notifications.
    ''' </summary>
    ''' <returns>Success status.</returns>
    ''' <remarks></remarks>
    Public Function UnregisterForPowerNotifications() As Boolean
        Return UnregisterPowerSettingNotification(hCallback)
    End Function


    Public Sub New(ByVal OnLidClose As OnLidCloseEvent)

        MyBase.New()

        'Dummy parameters used to create window
        Dim oParams As New CreateParams

        With oParams
            .X = 100
            .Y = 100
            .Height = 100
            .Width = 100
        End With

        'Create the listener window
        Me.CreateHandle(oParams)

        'Register this application to receive power notification messages.
        RegisterForPowerNotifications()
        Started = Now

        'Always force the logoff if the /noforce parameter is not specified
        ForceLogoff = True

        dLidClose = OnLidClose

        For Each sArg In My.Application.CommandLineArgs
            If sArg.Trim.ToLower = "/noforce" Then ForceLogoff = False
        Next

#If DEBUG Then
            If Not ForceLogoff Then Console.WriteLine("/noforce parameter specified. User will not be forced to log off.")

            Console.WriteLine("Listener started: " & Started.ToString & " with handle " & Me.Handle.ToString)
#End If


    End Sub


    ''' <summary>
    ''' Processes Windows messages. Modified to react to power notification messages.
    ''' </summary>
    ''' <param name="m">Message type recevied.</param>
    ''' <remarks></remarks>
    Protected Overrides Sub WndProc(ByRef m As System.Windows.Forms.Message)

#If DEBUG Then
            Console.WriteLine("Message received: " & m.Msg)
#End If

        'Respond to power notification
        If m.Msg = WM_POWERBROADCAST Then
            OnPowerBroadcast(m)
        End If

        'Handle everything else
        MyBase.WndProc(m)

    End Sub

    ''' <summary>
    ''' Reacts to the power notification message
    ''' </summary>.
    ''' <param name="m">Type of power notification message recevied.</param>
    ''' <remarks></remarks>
    Private Sub OnPowerBroadcast(ByRef m As System.Windows.Forms.Message)

        If m.WParam.ToInt32 = PBT_POWERSETTINGCHANGE Then

#If DEBUG Then
                Console.Write("Power setting change: " & m.LParam.ToString & " at " & Now.ToString)
#End If
            'Ignore any notification within 1 second of startup.
            'This is a fudge since Windows sends a message immediately after registering the listener for some reason, even if the lid
            'has not been closed.
            If Started.AddSeconds(1) < Now Then

#If DEBUG Then
                    Console.WriteLine(", logging user off now.")
#End If

                'Log the user off
                dLidClose.Invoke()

            Else
                'Power notification received immediately after startup, so ignore it.
#If DEBUG Then
                    Console.WriteLine(" (ignored).")
#End If
            End If

        End If

    End Sub

End Class
