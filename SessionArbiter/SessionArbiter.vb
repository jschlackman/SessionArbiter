Public Class SessionArbiter

#If DEBUG Then
    Private EventLog As New EventLog
#Else
    Inherits ServiceBase
#End If

    ''' <summary>
    ''' Registry path from HKLM to the SessionArbiter parameters.
    ''' </summary>
    ''' <remarks></remarks>
    Private Const sServiceParamsKey As String = "SYSTEM\CurrentControlSet\services\SessionArbiter\Parameters"

    Private WithEvents oCheckTimer As Timers.Timer

    ''' <summary>
    ''' Standard synchronisation period in milliseconds (may be ignored if limits demand more frequent checking).
    ''' </summary>
    ''' <remarks></remarks>
    Public StandardCheckPeriod As UInteger

    ''' <summary>
    ''' Whether to ignore RDS policy configuration when determining session limits.
    ''' </summary>
    ''' <remarks></remarks>
    Public IgnoreRDSPolicy As Boolean

    ''' <summary>
    ''' Loads and stores configured session limits as they apply to a given user on the current computer.
    ''' </summary>
    ''' <remarks></remarks>
    Public Class SessionLimits

        ''' <summary>
        ''' Standard registry path to RDS policies
        ''' </summary>
        ''' <remarks></remarks>
        Private Const sPolicyKey As String = "Software\Policies\Microsoft\Windows NT\Terminal Services"

        ''' <summary>
        ''' Maximum amount of time in milliseconds that the user's session can be active before the session is automatically disconnected or ended. 0 indicates no limit.
        ''' </summary>
        ''' <remarks></remarks>
        Public Active As UInteger
        ''' <summary>
        ''' Maximum amount of time in milliseconds that an active session can be idle (without user input) before the session is automatically disconnected or ended. 0 indicates no limit.
        ''' </summary>
        ''' <remarks>Not currently supported for local console sessions.</remarks>
        Public Idle As UInteger
        ''' <summary>
        ''' Maximum amount of time in milliseconds that a disconnected user session is kept active on the workstation. 0 indicates no limit.
        ''' </summary>
        ''' <remarks></remarks>
        Public Disconnected As UInteger
        Public TerminateSession As Boolean

        ''' <summary>
        ''' Returns a sensible check period in milliseconds based on the configured session limits.
        ''' </summary>
        ''' <remarks></remarks>
        Public ReadOnly Property CheckPeriod As UInteger
            Get
                Dim iPeriod As UInteger

                iPeriod = Idle \ 2
                If Active > 0 Then iPeriod = Math.Min(iPeriod, Active \ 2)
                If Disconnected > 0 Then iPeriod = Math.Min(iPeriod, Disconnected \ 2)

                Return iPeriod
            End Get
        End Property


        ''' <summary>
        ''' Reads limits for current machine, optionally ignoring any RDS policy settings.
        ''' </summary>
        ''' <param name="IgnoreRDSPolicy">Specify True to ignore any configured RDS policies.</param>
        ''' <remarks></remarks>
        Public Sub New(Optional ByVal IgnoreRDSPolicy As Boolean = False)
            InitLimits()
            'Service params have lowest precedence
            GetFromServiceParams()
            'Skip loading from machine policy if we are igonring it.
            If Not IgnoreRDSPolicy Then GetFromMachinePolicy()
        End Sub

        ''' <summary>
        ''' Readins limits for the given user on the current machine, optionally ignoring any RDS policy settings.
        ''' </summary>
        ''' <param name="UserAccount">User account to read limits for, in the format DOMAIN\username</param>
        ''' <param name="IgnoreRDSPolicy">Specify True to ignore any configured RDS policies.</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal UserAccount As String, Optional ByVal IgnoreRDSPolicy As Boolean = False)
            InitLimits()
            'Service params have lowest precedence
            GetFromServiceParams()
            'Skip loading from machine and user policy if we are igonring it.
            If Not IgnoreRDSPolicy Then
                GetFromUserPolicy(UserAccount)
                GetFromMachinePolicy()
            End If
        End Sub

        ''' <summary>
        ''' Set default limits.
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub InitLimits()
            Active = 0
            Idle = 0
            Disconnected = 0
        End Sub

        ''' <summary>
        ''' Load limits from the local computer RDS policies.
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub GetFromMachinePolicy()

            'Try to get read-only access to the machine policy.
            Dim oMachinePolicy As RegistryKey = Registry.LocalMachine.OpenSubKey(sPolicyKey, False)

            'Read the machine policy if it exists.
            If Not oMachinePolicy Is Nothing Then
                GetFromRegistry(oMachinePolicy)
                oMachinePolicy.Close()
            End If

        End Sub

        ''' <summary>
        ''' Load limits from the specified user's RDS policies (user must be logged on for this to succeed)
        ''' </summary>
        ''' <param name="UserAccount">User account to read limits for, in the format DOMAIN\username</param>
        ''' <remarks></remarks>
        Private Sub GetFromUserPolicy(ByVal UserAccount As String)
            Dim oUserPolicy As RegistryKey = GetUserRegistry(UserAccount).OpenSubKey(sPolicyKey, False)

            'Read the user policy if it exists.
            If Not oUserPolicy Is Nothing Then
                GetFromRegistry(oUserPolicy)
                oUserPolicy.Close()
            End If

        End Sub

        ''' <summary>
        ''' Load limits from the service parameters key in the system registry.
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub GetFromServiceParams()

            'Try to get read-only access to the machine policy.
            Dim oServiceParams As RegistryKey = Registry.LocalMachine.OpenSubKey(sServiceParamsKey, False)

            'Read the machine policy if it exists.
            If Not oServiceParams Is Nothing Then
                GetFromRegistry(oServiceParams)
                oServiceParams.Close()
            End If

        End Sub

        ''' <summary>
        ''' Load limits from the specified registry key.
        ''' </summary>
        ''' <param name="SettingsRoot">Reference to an open registry key to read limits from.</param>
        ''' <remarks></remarks>
        Private Sub GetFromRegistry(ByVal SettingsRoot As RegistryKey)

            'Get the active session limit.
            If Not (SettingsRoot.GetValue("MaxConnectionTime") Is Nothing) Then
                Active = SettingsRoot.GetValue("MaxConnectionTime")
                'If configured value read but is invalid (less than 1 minute), set no limit.
                If Active < 60000 Then
                    Active = 0
                End If
            End If

            'Get the disconnected session limit.
            If Not (SettingsRoot.GetValue("MaxDisconnectionTime") Is Nothing) Then
                Disconnected = SettingsRoot.GetValue("MaxDisconnectionTime")
                'If configured value read but is invalid (less than 1 minute), set no limit.
                If Disconnected < 60000 Then
                    Disconnected = 0
                End If
            End If

            'Get the idle session limit.
            If Not (SettingsRoot.GetValue("MaxIdleTime") Is Nothing) Then
                Idle = SettingsRoot.GetValue("MaxIdleTime")
                'If configured value read but is invalid (less than 1 minute), set no limit.
                If Idle < 60000 Then
                    Idle = 0
                End If
            End If

            'Get the idle session limit.
            If Not (SettingsRoot.GetValue("fResetBroken") Is Nothing) Then
                TerminateSession = SettingsRoot.GetValue("fResetBroken")
            End If

        End Sub


        ''' <summary>
        ''' Open the HKEY_USERS subkey for a given named user account.
        ''' </summary>
        ''' <param name="UserAccount">User account to find the registry key of.</param>
        ''' <returns>Read-only reference to registry key if it can be found, otherwise Nothing.</returns>
        ''' <remarks></remarks>
        Private Function GetUserRegistry(ByVal UserAccount As String) As RegistryKey

            Dim oReturn As RegistryKey

            Try
                Dim oUser As New Security.Principal.NTAccount(UserAccount)
                Dim sUserSID As String = oUser.Translate(GetType(Security.Principal.SecurityIdentifier)).ToString

                'Open handle to subkey
                oReturn = Registry.Users.OpenSubKey(sUserSID, False)
                'Close handle to root key, no longer needed.
                Registry.Users.Close()

            Catch ex As Exception
                oReturn = Nothing
            End Try

            Return oReturn

        End Function

    End Class

    ''' <summary>
    ''' Counts down 2 minutes, then executes a specified action against a session.
    ''' </summary>
    ''' <remarks></remarks>
    Public Class CountdownTimer

        Implements IDisposable

        Public Delegate Sub CountdownAction(ByVal Session As ITerminalServicesSession)

        ''' <summary>
        ''' Internal reference to the session we are counting down.
        ''' </summary>
        ''' <remarks></remarks>
        Private oSession As ITerminalServicesSession
        ''' <summary>
        ''' Connection state that originally started this countdown.
        ''' </summary>
        ''' <remarks></remarks>
        Private oTriggerState As ConnectionState
        ''' <summary>
        ''' Whether the session will be logged off or just disconnected.
        ''' </summary>
        ''' <remarks></remarks>
        Private bLogoff As Boolean
        ''' <summary>
        ''' Reference to subroutine that will actually peform the logoff or disconnect.
        ''' </summary>
        ''' <remarks></remarks>
        Private dAction As CountdownAction

        Private WithEvents oTimer As Timers.Timer

        ''' <summary>
        ''' Create a new countdown timer.
        ''' </summary>
        ''' <param name="Session">Session to countdown for.</param>
        ''' <param name="WillLogoff">Determines whether the session will be logged off or just disconnected.</param>
        ''' <param name="PerformAction">Reference to subroutine that will actually peform the logoff or disconnect.</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal Session As ITerminalServicesSession, ByVal WillLogoff As Boolean, ByVal PerformAction As CountdownAction)
            oSession = Session
            oTriggerState = Session.ConnectionState
            bLogoff = WillLogoff
            dAction = PerformAction

            '2 minute timer
            oTimer = New Timers.Timer(180000)
            oTimer.AutoReset = False

            'Warn the user they are about to be disconnected/logged off
            Dim sMessage As String = "Your logon session has been " & Session.ConnectionState.ToString.ToLower & " for the maximum allowed time. You will be "

            If WillLogoff Then
                sMessage &= "logged off"
            Else
                sMessage &= "disconnected"
            End If

            sMessage &= " in 2 minutes. Save your work now."

            'Start the clock.
            oTimer.Enabled = True

            'Send the warning message.
            Session.MessageBox(sMessage)

        End Sub

        ''' <summary>
        ''' Runs when the countdown expires.
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Private Sub OnJobTimer(ByVal sender As System.Object, ByVal e As System.Timers.ElapsedEventArgs) Handles oTimer.Elapsed
            'Check the session is still valid
            If Not oSession Is Nothing Then
                'If the connection state is unchanged, perform the correct action
                If oSession.ConnectionState = oTriggerState Then
                    dAction.Invoke(oSession)
                End If
            End If
            'We are done with this countdown timer.
            Me.Dispose()
        End Sub

        Public Sub Dispose() Implements IDisposable.Dispose
            oTimer.Dispose()
        End Sub

    End Class


    ''' <summary>
    ''' Application Event IDs for SessionArbiter
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum ArbiterEvent
        ''' <summary>
        ''' Reserved for automatic log events.
        ''' </summary>
        ''' <remarks></remarks>
        Reserved = 0
        ''' <summary>
        ''' Successfully logged off a disconnected session.
        ''' </summary>
        ''' <remarks></remarks>
        LogoffDisconnected = 1
        ''' <summary>
        ''' Failed to logoff a disconnected session.
        ''' </summary>
        ''' <remarks></remarks>
        LogoffDisconnectedFailed = 2
    End Enum

    ''' <summary>
    ''' Initialise a new SessionArbiter service instance.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New()

#If Not Debug Then

        MyBase.New()

        ServiceName = "SessionArbiter"
        CanStop = True
        CanPauseAndContinue = False
        'Turn off superfluous application log messages.
        AutoLog = False

#End If
        'Start the arbiter
        StartArbiter()
    End Sub

    ''' <summary>
    ''' Starts the main Session Arbiter functionality.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub StartArbiter()
        oCheckTimer = New Timers.Timer

        'Read initial config from registry.
        ReadServiceParameters()

        'First sync after 5 seconds
        oCheckTimer.Interval = 5000
    End Sub

#If Not Debug Then

    ''' <summary>
    ''' Executes when a Start command is issued to the service.
    ''' </summary>
    ''' <param name="args">Command line arguments passed on start.</param>
    ''' <remarks></remarks>
    Protected Overrides Sub OnStart(args() As String)
        MyBase.OnStart(args)
        oCheckTimer.Enabled = True
    End Sub

    ''' <summary>
    ''' Executes when a Stop command is issued to the service.
    ''' </summary>
    ''' <remarks></remarks>
    Protected Overrides Sub OnStop()
        MyBase.OnStop()
        oCheckTimer.Enabled = False
    End Sub

#End If

    ''' <summary>
    ''' Main entry point when service execution begins.
    ''' </summary>
    ''' <remarks></remarks>
    Public Shared Sub Main()

#If DEBUG Then
        Console.WriteLine("Starting debug mode...")
        Dim oDebug As New SessionArbiter()
        oDebug.oCheckTimer.Enabled = True
        Console.ReadKey()
#Else
        ServiceBase.Run(New SessionArbiter)
#End If
    End Sub

    ''' <summary>
    ''' Runs the session checks whenever the configured timer period elapses.
    ''' </summary>
    ''' <param name="sender">Timer object that has triggered this event.</param>
    ''' <param name="e">Event arguments from the Timer object.</param>
    ''' <remarks></remarks>
    Private Sub OnJobTimer(ByVal sender As System.Object, ByVal e As System.Timers.ElapsedEventArgs) Handles oCheckTimer.Elapsed


        'Read current config from registry, in case it has changed since last check.
        ReadServiceParameters()
        oCheckTimer.Interval = StandardCheckPeriod

        Dim oManager As ITerminalServicesManager = New TerminalServicesManager

        Using oServer As ITerminalServer = oManager.GetLocalServer
            oServer.Open()

            For Each oSession As ITerminalServicesSession In oServer.GetSessions

                'Only check sessions with usernames (ignores Services session).
                If Not oSession.UserAccount Is Nothing Then

                    Dim oLimits As New SessionLimits(oSession.UserAccount.Value, IgnoreRDSPolicy)

                    'Set next sync to run after configured sync period.
                    If oLimits.CheckPeriod > 0 Then
                        'Interval should be whichever is shorter, the standard period or the determined sensible period for this session's limits.
                        oCheckTimer.Interval = Math.Min(StandardCheckPeriod, oLimits.CheckPeriod)
                    End If

                    'If this session is disconnected and there is a time limit
                    If (oSession.ConnectionState = ConnectionState.Disconnected) And (oLimits.Disconnected > 0) Then
                        'Check if it has been disconnected for longer than the configured limit
                        If oSession.DisconnectTime < Now.AddMilliseconds(-oLimits.Disconnected) Then
                            LogoffSession(oSession)
                        End If
                    End If

                    'If this session is active and there is a time limit
                    If (oSession.ConnectionState = ConnectionState.Active) And (oLimits.Active > 0) Then
                        'Check if it has been logged in for longer than the configured limit
                        If oSession.LoginTime < Now.AddMilliseconds(-oLimits.Active) Then
                            'Start counting down to logoff/disconnect
                            StartCountdown(oSession, oLimits.TerminateSession)
                        End If
                    End If

                    'NOTE: There is no handler for idle time as the RDS API does not currently return idle time for local console sessions.

                End If

            Next

            oServer.Close()
        End Using

    End Sub

    
    ''' <summary>
    ''' Start counting down a session to log off or disconnect.
    ''' </summary>
    ''' <param name="Session">Session to log off or disconnect.</param>
    ''' <param name="WillLogoff">Whether to log the session off when  the countdown expires, or just disconnect it.</param>
    ''' <remarks></remarks>
    Private Sub StartCountdown(ByVal Session As ITerminalServicesSession, ByVal WillLogoff As Boolean)
        'Start a countdown, passing in the session to log off, the type of action, and the subroutine delegate
        If WillLogoff Then
            Dim oCountdown As New CountdownTimer(Session, True, AddressOf LogoffSession)
        Else
            Dim oCountdown As New CountdownTimer(Session, False, AddressOf DisconnectSession)
        End If

    End Sub

    ''' <summary>
    ''' Log off a session immediately and make an event log entry.
    ''' </summary>
    ''' <param name="Session">Session to log off.</param>
    ''' <remarks></remarks>
    Private Sub LogoffSession(ByVal Session As ITerminalServicesSession)
        Dim sLogMessage As String = Session.ConnectionState.ToString.ToLower & " session for user " & Session.UserAccount.Value & "."

        Select Case Session.ConnectionState
            Case ConnectionState.Disconnected
                sLogMessage &= " Session was disconnected at " & Session.DisconnectTime & "."
            Case ConnectionState.Active
                sLogMessage &= " Session logged in at " & Session.LoginTime & "."
            Case ConnectionState.Idle
                sLogMessage &= " Session was idle since " & Session.LastInputTime & "."
        End Select

        Try
            'Attempt to log the session off
            Session.Logoff(True)
            'Log success
            WriteEventLogEntry("Successfully logged off " & sLogMessage, EventLogEntryType.Information, ArbiterEvent.LogoffDisconnected)
        Catch ex As Exception
            'Log failure
            WriteEventLogEntry("Failed to log off " & sLogMessage & vbCrLf & vbCrLf & ex.Message, EventLogEntryType.Warning, ArbiterEvent.LogoffDisconnectedFailed)
        End Try
    End Sub


    ''' <summary>
    ''' Disconnect a session immediately and make an event log entry.
    ''' </summary>
    ''' <param name="Session">Session to disconnect.</param>
    ''' <remarks></remarks>
    Private Sub DisconnectSession(ByVal Session As ITerminalServicesSession)
        Dim sLogMessage As String = Session.ConnectionState.ToString.ToLower & " session for user " & Session.UserAccount.Value & "."

        Select Case Session.ConnectionState
            Case ConnectionState.Active
                sLogMessage &= " Session logged in at " & Session.LoginTime & "."
            Case ConnectionState.Idle
                sLogMessage &= " Session was idle since " & Session.LastInputTime & "."
        End Select

        Try
            'Attempt to disconnect the session.
            Session.Disconnect(True)
            'Log success
            WriteEventLogEntry("Successfully disconnected " & sLogMessage, EventLogEntryType.Information, ArbiterEvent.LogoffDisconnected)
        Catch ex As Exception
            'Log failure
            WriteEventLogEntry("Failed to disconnect " & sLogMessage & vbCrLf & vbCrLf & ex.Message, EventLogEntryType.Warning, ArbiterEvent.LogoffDisconnectedFailed)
        End Try
    End Sub

    ''' <summary>
    ''' Attempt to read configurable service parameters from the registry.
    ''' </summary>
    ''' <remarks>Loads default values for any parameters that cannot be read.</remarks>
    Private Sub ReadServiceParameters()

        Const sCheckPeriodValue As String = "CheckPeriod"
        Const sIgnoreRDSPolicy As String = "IgnorePolicy"
        Const iDefaultCheckPeriod As Integer = 900000 'Default is 15 minutes (90000ms)
        Const bDefaultIgnorePolicy As Boolean = False

        Dim oRegKey As RegistryKey = Nothing

        Try
            'Open the SyncDesc key with read-only access.
            oRegKey = Registry.LocalMachine.OpenSubKey(sServiceParamsKey, False)

            'Get the check period
            Try
                StandardCheckPeriod = oRegKey.GetValue(sCheckPeriodValue)
                'If configured value read and is valid (1 minute or greater), set it as the check period.
                If StandardCheckPeriod < 60000 Then
                    StandardCheckPeriod = 0
                End If
            Catch ex As Exception
                'Use default if a setting could not be read
                StandardCheckPeriod = iDefaultCheckPeriod
            End Try

            'Get whether to ignore RDS policy
            Try
                IgnoreRDSPolicy = oRegKey.GetValue(sIgnoreRDSPolicy)
            Catch ex As Exception
                'Use default if a setting could not be read
                IgnoreRDSPolicy = bDefaultIgnorePolicy
            End Try

        Catch ex As Exception
            'Use default values for everything if there was a problem opening the registry.
            StandardCheckPeriod = iDefaultCheckPeriod
            IgnoreRDSPolicy = bDefaultIgnorePolicy
        Finally
            If Not oRegKey Is Nothing Then oRegKey.Close()
        End Try

    End Sub

    ''' <summary>
    ''' Writes a message to the application event log.
    ''' </summary>
    ''' <param name="Message">Message to log.</param>
    ''' <param name="Level">Event level.</param>
    ''' <param name="EventID">Event ID.</param>
    ''' <remarks></remarks>
    Private Sub WriteEventLogEntry(ByVal Message As String, ByVal Level As EventLogEntryType, ByVal EventID As Integer)

        Const sEventSource As String = "Session Arbiter"

        'Create event log source if it doesn't already exist.
        If Not EventLog.SourceExists(sEventSource) Then
            EventLog.CreateEventSource(sEventSource, "Application")
        End If

        'Make sure we have exclusive access to the log writer.
        SyncLock EventLog
            'Write the log entry.
            EventLog.WriteEntry(sEventSource, Message, Level, EventID)
        End SyncLock

    End Sub

End Class
