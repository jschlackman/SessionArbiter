Imports System.Windows.Forms
Imports System.Threading

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

    ''' <summary>
    ''' Session limit check timer
    ''' </summary>
    ''' <remarks></remarks>
    Private WithEvents oCheckTimer As Timers.Timer

    ''' <summary>
    ''' Background listener thread for lid events
    ''' </summary>
    ''' <remarks></remarks>
    Private ListenThread As Thread

    Private oListener As LidListener

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
    ''' Whether to put the computer into Suspend after a user is logged off due to the lid closing.
    ''' </summary>
    ''' <remarks></remarks>
    Public SuspendOnLidCloseAtLogonScreen As Boolean


    ''' <summary>
    ''' Loads and stores configured session limits as they apply to a given user on the current computer.
    ''' </summary>
    ''' <remarks></remarks>
    Public Class SessionLimits

        ''' <summary>
        ''' Standard registry path to RDS policies
        ''' </summary>
        ''' <remarks></remarks>
        Private Const sPolicyKeyRDS As String = "Software\Policies\Microsoft\Windows NT\Terminal Services"
        ''' <summary>
        ''' Registry path to 
        ''' </summary>
        ''' <remarks></remarks>
        Private Const sPolicyKeySA As String = "Software\Policies\SessionArbiter"

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
        ''' <summary>
        ''' Whether to log off a session instead of just disconnecting it.
        ''' </summary>
        ''' <remarks></remarks>
        Public TerminateSession As Boolean
        ''' <summary>
        ''' Whether to log off the session wheh the lid is closed
        ''' </summary>
        ''' <remarks></remarks>
        Public ReadOnly Property LogoffOnLidClose As Boolean
            Get
                Return (iLogoff = 1)
            End Get
        End Property

        ''' <summary>
        ''' Logoff on lid close policy state
        ''' </summary>
        ''' <remarks>0 indicates not configured (don't log off), 1 indicates log off, 2 indicates the policy is disabled (don't logoff).</remarks>
        Private iLogoff As UInteger
        ''' <summary>
        ''' Whether to ignore RDS policy when reading registry settings
        ''' </summary>
        ''' <remarks></remarks>
        Private bIgnoreRDS As Boolean

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
            'Flag to skip loading from machine RDS policy if we are ignoring it.
            bIgnoreRDS = IgnoreRDSPolicy
            'Load machine policy
            GetFromMachinePolicy()
        End Sub

        ''' <summary>
        ''' Reads limits for the given user on the current machine, optionally ignoring any RDS policy settings.
        ''' </summary>
        ''' <param name="UserAccount">User account to read limits for, in the format DOMAIN\username</param>
        ''' <param name="IgnoreRDSPolicy">Specify True to ignore any configured RDS policies.</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal UserAccount As String, Optional ByVal IgnoreRDSPolicy As Boolean = False)
            InitLimits()
            'Service params have lowest precedence
            GetFromServiceParams()
            'Skip loading from machine and user policy if we are ignoring it.
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

            Dim oMachinePolicy As RegistryKey

            'First get RDS policy
            If Not bIgnoreRDS Then
                'Try to get read-only access to the machine policy.
                oMachinePolicy = Registry.LocalMachine.OpenSubKey(sPolicyKeyRDS, False)

                'Read the machine policy if it exists.
                If Not oMachinePolicy Is Nothing Then
                    GetRDSFromRegistry(oMachinePolicy)
                    oMachinePolicy.Close()
                End If
            End If

            'Now get the Session Arbiter policy
            oMachinePolicy = Registry.LocalMachine.OpenSubKey(sPolicyKeySA, False)

            'Read the machine policy if it exists.
            If Not oMachinePolicy Is Nothing Then
                GetSAFromRegistry(oMachinePolicy)
                oMachinePolicy.Close()
            End If

        End Sub

        ''' <summary>
        ''' Load limits from the specified user's RDS policies (user must be logged on for this to succeed)
        ''' </summary>
        ''' <param name="UserAccount">User account to read limits for, in the format DOMAIN\username</param>
        ''' <remarks></remarks>
        Private Sub GetFromUserPolicy(ByVal UserAccount As String)
            Dim oUserKey As RegistryKey = Nothing

            'Get the user's registry hive and open the RDS policy subkey
            If GetUserRegistry(UserAccount, oUserKey) Then

                Dim oUserPolicy As RegistryKey

                If Not bIgnoreRDS Then
                    'First get RDS policy
                    oUserPolicy = oUserKey.OpenSubKey(sPolicyKeyRDS, False)
                    'Read the user policy if it exists.
                    If Not oUserPolicy Is Nothing Then
                        GetRDSFromRegistry(oUserPolicy)
                        'Make sure the key is closed or User Profile Service will complain.
                        oUserPolicy.Close()
                    End If
                End If

                'Now get Session Arbiter policy
                oUserPolicy = oUserKey.OpenSubKey(sPolicyKeySA, False)
                'Read the user policy if it exists.
                If Not oUserPolicy Is Nothing Then
                    GetSAFromRegistry(oUserPolicy)
                    'Make sure the key is closed or User Profile Service will complain.
                    oUserPolicy.Close()
                End If

            End If

            'Make sure the user's root key is closed too.
            oUserKey.Close()

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
                GetRDSFromRegistry(oServiceParams)
                GetSAFromRegistry(oServiceParams)
                oServiceParams.Close()
            End If

        End Sub

        ''' <summary>
        ''' Load RDS limits from the specified registry key.
        ''' </summary>
        ''' <param name="SettingsRoot">Reference to an open registry key to read limits from.</param>
        ''' <remarks></remarks>
        Private Sub GetRDSFromRegistry(ByVal SettingsRoot As RegistryKey)

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

            'Gets whether to log off or disconnect idle sessions.
            If Not (SettingsRoot.GetValue("fResetBroken") Is Nothing) Then
                TerminateSession = SettingsRoot.GetValue("fResetBroken")
            End If

        End Sub


        ''' <summary>
        ''' Load RDS limits from the specified registry key.
        ''' </summary>
        ''' <param name="SettingsRoot">Reference to an open registry key to read limits from.</param>
        ''' <remarks></remarks>
        Private Sub GetSAFromRegistry(ByVal SettingsRoot As RegistryKey)

            Dim iSuspend As UInteger

            'Get whether to log off on lid close
            If Not (SettingsRoot.GetValue("LogoffOnLidClose") Is Nothing) Then
                iSuspend = SettingsRoot.GetValue("LogoffOnLidClose")
                'If configured value read but is invalid (>2), set to 0.
                If iSuspend > 2 Then
                    iSuspend = 0
                End If
            End If

            iLogoff = Math.Max(iSuspend, iLogoff)

        End Sub

        ''' <summary>
        ''' Open the HKEY_USERS subkey for a given named user account.
        ''' </summary>
        ''' <param name="UserAccount">User account to find the registry key of.</param>
        ''' <returns>Read-only reference to registry key if it can be found, otherwise Nothing.</returns>
        ''' <remarks></remarks>
        Private Function GetUserRegistry(ByVal UserAccount As String, ByRef UserKey As RegistryKey) As Boolean

            Dim bReturn As Boolean

            Try
                Dim oUser As New Security.Principal.NTAccount(UserAccount)
                Dim sUserSID As String = oUser.Translate(GetType(Security.Principal.SecurityIdentifier)).ToString

                'Open handle to subkey
                UserKey = Registry.Users.OpenSubKey(sUserSID, False)
                'Close handle to root key, no longer needed.
                Registry.Users.Close()

                'Success
                bReturn = True
            Catch ex As Exception
                bReturn = False
            End Try

            Return bReturn

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
        ''' <summary>
        ''' System power state altered.
        ''' </summary>
        ''' <remarks></remarks>
        PowerState = 3
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

        'Create a background listerner thread to process lid events
        ListenThread = New Thread(AddressOf RunListenThread)
        ListenThread.IsBackground = True

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
        ListenThread.Start()
    End Sub

    ''' <summary>
    ''' Executes when a Stop command is issued to the service.
    ''' </summary>
    ''' <remarks></remarks>
    Protected Overrides Sub OnStop()
        MyBase.OnStop()
        oCheckTimer.Enabled = False
        ListenThread.Abort()

        'Stop listening for power notifications
        oListener.UnregisterForPowerNotifications()

    End Sub


#End If

    ''' <summary>
    ''' Main entry point when service execution begins.
    ''' </summary>
    ''' <remarks></remarks>
    Public Shared Sub Main()

#If DEBUG Then
        Console.WriteLine("Starting debug mode.")
        Dim oDebug As New SessionArbiter()
        oDebug.oCheckTimer.Enabled = True
        Console.Read()
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
        Dim dLimitTime As Date

        Using oServer As ITerminalServer = oManager.GetLocalServer
            oServer.Open()

#If DEBUG Then
            Console.WriteLine("Starting session check at " & Now)
#End If

            For Each oSession As ITerminalServicesSession In oServer.GetSessions

                'Only check sessions with usernames (ignores Services session).
                If Not oSession.UserAccount Is Nothing Then

#If DEBUG Then
                    Console.WriteLine("Checking user " & oSession.UserAccount.Value)
#End If

                    Dim oLimits As New SessionLimits(oSession.UserAccount.Value, IgnoreRDSPolicy)

#If DEBUG Then
                    Console.WriteLine("Session limits for this user:")
                    Console.WriteLine("Active: " & oLimits.Active)
                    Console.WriteLine("Disconnected: " & oLimits.Disconnected)
                    Console.WriteLine("Idle: " & oLimits.Idle)
                    Console.WriteLine("Logoff on lid close: " & oLimits.LogoffOnLidClose)
#End If

                    'Set next sync to run after configured sync period.
                    If oLimits.CheckPeriod > 0 Then
                        'Interval should be whichever is shorter, the standard period or the determined sensible period for this session's limits.
                        oCheckTimer.Interval = Math.Min(StandardCheckPeriod, oLimits.CheckPeriod)
                    End If

                    'If this session is disconnected and there is a time limit
                    If (oSession.ConnectionState = ConnectionState.Disconnected) And (oLimits.Disconnected > 0) Then

                        'Determine when the limit will be/was reached
                        dLimitTime = oSession.DisconnectTime.Value.AddMilliseconds(oLimits.Disconnected)
#If DEBUG Then
                        Console.WriteLine("Session will reach disconnected limit at " & dLimitTime)
#End If

                        'Check if session has been disconnected for longer than or equal to the configured limit
                        If Now >= dLimitTime Then
                            LogoffSession(oSession)
                        Else
                            'Make sure the limit cannot be reached before the next check
                            CheckInterval(dLimitTime)
                        End If

                    End If

                    'If this session is active and there is a time limit
                    If (oSession.ConnectionState = ConnectionState.Active) And (oLimits.Active > 0) Then

                        'Determine when the limit will be/was reached
                        dLimitTime = oSession.LoginTime.Value.AddMilliseconds(oLimits.Active)

#If DEBUG Then
                        Console.WriteLine("Session will reach active limit at " & dLimitTime)
#End If

                        'Check if session has been logged in for longer than or equal to the configured limit
                        If Now >= dLimitTime Then
                            'Start counting down to logoff/disconnect
                            StartCountdown(oSession, oLimits.TerminateSession)
                        Else
                            'Make sure the limit cannot be reached before the next check
                            CheckInterval(dLimitTime)
                        End If

                    End If

                    'NOTE: There is no handler for idle time as the RDS API does not currently return idle time for local console sessions.

                End If

            Next

            oServer.Close()
        End Using

    End Sub

    ''' <summary>
    ''' Checks if a limit is due to be reached before the next check, and adjusts the next check time accordingly
    ''' </summary>
    ''' <param name="dLimitTime">Time that the session limit is due to expire.</param>
    ''' <remarks></remarks>
    Private Sub CheckInterval(ByVal dLimitTime As Date)

        'If the session limit is due to be reached before the next scheduled check
        If dLimitTime < Now.AddMilliseconds(oCheckTimer.Interval) Then
            'Adjust the next check so that it coincides with the time the limit is due to be reached
            oCheckTimer.Interval = (dLimitTime - Now).TotalMilliseconds
        End If

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
    Private Function LogoffSession(ByVal Session As ITerminalServicesSession) As Boolean

        Dim sLogMessage As String = Session.ConnectionState.ToString.ToLower & " session for user " & Session.UserAccount.Value & "."
        Dim bSuccess As Boolean = True

        Select Case Session.ConnectionState
            Case ConnectionState.Disconnected
                sLogMessage &= " Session was disconnected at " & Session.DisconnectTime & "."
            Case ConnectionState.Active
                sLogMessage &= " Session logged in at " & Session.LoginTime & "."
            Case ConnectionState.Idle
                sLogMessage &= " Session was idle since " & Session.LastInputTime & "."
        End Select

        Try
            'Attempt to log the session off and WAITS for it to finish before continuing.
            Session.Logoff(True)
            'Log success
            WriteEventLogEntry("Successfully logged off " & sLogMessage, EventLogEntryType.Information, ArbiterEvent.LogoffDisconnected)
        Catch ex As Exception
            bSuccess = False
            'Log failure
            WriteEventLogEntry("Failed to log off " & sLogMessage & vbCrLf & vbCrLf & ex.Message, EventLogEntryType.Warning, ArbiterEvent.LogoffDisconnectedFailed)
        End Try

        Return bSuccess
    End Function


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
        Const sSuspendOnLidCloseAtLogonScreen As String = "SuspendOnGinaLidClose"
        Const iDefaultCheckPeriod As UInteger = 900000 'Default is 15 minutes (900s)
        Const bDefaultIgnorePolicy As Boolean = False
        Const bSuspendOnLidCloseAtLogonScreen As Boolean = False

        Dim oRegKey As RegistryKey = Nothing

        Try
            'Open the SyncDesc key with read-only access.
            oRegKey = Registry.LocalMachine.OpenSubKey(sServiceParamsKey, False)

            'Get the check period
            Try
                StandardCheckPeriod = oRegKey.GetValue(sCheckPeriodValue)
                'If configured value read and is valid (1 minute or greater), set it as the check period.
                If StandardCheckPeriod < 60000 Then
                    StandardCheckPeriod = iDefaultCheckPeriod
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


            'Get whether to suspend after a logoff due to lod closure
            Try
                SuspendOnLidCloseAtLogonScreen = oRegKey.GetValue(sSuspendOnLidCloseAtLogonScreen)
            Catch ex As Exception
                'Use default if a setting could not be read
                SuspendOnLidCloseAtLogonScreen = bSuspendOnLidCloseAtLogonScreen
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

    ''' <summary>
    ''' Starts a listener window for lid events
    ''' </summary>
    ''' <remarks></remarks>
    Sub RunListenThread()
        'Listen for lid events
        oListener = New LidListener(AddressOf LogoffActiveSessions)
        Application.Run()
    End Sub


    ''' <summary>
    ''' Logs off any sessions that are currently Active (on a client workstation, only the console session can be Active).
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub LogoffActiveSessions()

        ReadServiceParameters()

#If DEBUG Then
        Console.WriteLine("Checking sessions to log off.")
#End If

        Dim oManager As ITerminalServicesManager = New TerminalServicesManager

        Using oServer As ITerminalServer = oManager.GetLocalServer
            oServer.Open()

            For Each oSession As ITerminalServicesSession In oServer.GetSessions

                'Only check sessions with usernames (ignores Services session).
                If Not oSession.UserAccount Is Nothing Then

                    'Get session parameters
                    Dim oLimits As New SessionLimits(oSession.UserAccount.Value, IgnoreRDSPolicy)

                    If oLimits.LogoffOnLidClose Then
                        'Only logoff the session if it is active
                        If oSession.ConnectionState = ConnectionState.Active Then

                            'Logoff this session
                            LogoffSession(oSession)

                        End If

                    End If
                End If

            Next

            'Suspend if requested
            If SuspendOnLidCloseAtLogonScreen Then
                WriteEventLogEntry("Entering sleep mode.", EventLogEntryType.Information, ArbiterEvent.PowerState)
                Application.SetSuspendState(PowerState.Suspend, True, False)
            End If

        End Using

    End Sub

End Class
