Public Class SessionArbiter

    Inherits ServiceBase
    Private WithEvents JobTimer As Timers.Timer

    ''' <summary>
    ''' Synchronisation period in milliseconds.
    ''' </summary>
    ''' <remarks></remarks>
    Public CheckPeriod As Integer
    ''' <summary>
    ''' Maximum time in minutes that a session can be disconnected before logoff.
    ''' </summary>
    ''' <remarks></remarks>
    Public DisconnectedLimit As Integer


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
        MyBase.New()

        ServiceName = "SessionArbiter"
        CanStop = True
        CanPauseAndContinue = False
        'Turn off superfluous application log messages.
        AutoLog = False
        JobTimer = New Timers.Timer

        'Read config from registry.
        ReadServiceParameters()

        'First sync after 5 seconds
        JobTimer.Interval = 5000
    End Sub

    ''' <summary>
    ''' Executes when a Start command is issued to the service.
    ''' </summary>
    ''' <param name="args">Command line arguments passed on start.</param>
    ''' <remarks></remarks>
    Protected Overrides Sub OnStart(args() As String)
        MyBase.OnStart(args)
        JobTimer.Enabled = True
    End Sub

    ''' <summary>
    ''' Executes when a Stop command is issued to the service.
    ''' </summary>
    ''' <remarks></remarks>
    Protected Overrides Sub OnStop()
        MyBase.OnStop()
        JobTimer.Enabled = False
    End Sub

    ''' <summary>
    ''' Main entry point when service execution begins.
    ''' </summary>
    ''' <remarks></remarks>
    Public Shared Sub Main()
        ServiceBase.Run(New SessionArbiter)
    End Sub

    ''' <summary>
    ''' Synchronises the computer description with Active Directory. Runs whenever the configured timer period elapses.
    ''' </summary>
    ''' <param name="sender">Timer object that has triggered this event.</param>
    ''' <param name="e">Event arguments from the Timer object.</param>
    ''' <remarks></remarks>
    Private Sub CheckSessions(ByVal sender As System.Object, ByVal e As System.Timers.ElapsedEventArgs) Handles JobTimer.Elapsed

        Dim oManager As ITerminalServicesManager = New TerminalServicesManager
        Dim sLogMessage As String = ""

        Using oServer As ITerminalServer = oManager.GetLocalServer
            oServer.Open()

            For Each oSession As ITerminalServicesSession In oServer.GetSessions

                'Only check sessions with usernames (ignores Services session).
                If Not oSession.UserName Is Nothing Then

                    'If this session is disconnected
                    If oSession.ConnectionState = ConnectionState.Disconnected Then
                        'Check if it has been disconnected for longer than the configured limit
                        If oSession.DisconnectTime < Now.AddMinutes(-DisconnectedLimit) Then
                            sLogMessage = String.Format("user {0}. Session was disconnected at {1}.", oSession.UserAccount.Value, oSession.DisconnectTime)
                            Try
                                'Attempt to log the session off
                                oSession.Logoff(True)
                                'Log success
                                WriteEventLogEntry("Successfully logged off " & sLogMessage, EventLogEntryType.Information, ArbiterEvent.LogoffDisconnected)
                            Catch ex As Exception
                                'Log failure
                                WriteEventLogEntry("Failed to log off " & sLogMessage & vbCrLf & vbCrLf & ex.Message, EventLogEntryType.Warning, ArbiterEvent.LogoffDisconnectedFailed)
                            End Try
                        End If
                    End If
                End If

            Next

            oServer.Close()
        End Using

        'Read config from registry.
        ReadServiceParameters()

        'Set next sync to run after configured sync period.
        JobTimer.Interval = CheckPeriod

    End Sub

    ''' <summary>
    ''' Attempt to read configurable service parameters from the registry.
    ''' </summary>
    ''' <remarks>Loads default values for any parameters that cannot be read.</remarks>
    Private Sub ReadServiceParameters()

        Const sSyncDescKey As String = "SYSTEM\CurrentControlSet\services\SessionArbiter\Parameters"
        Const sCheckPeriodValue As String = "CheckPeriod"
        Const sDisconnectedLimitValue As String = "DisconnectedLimit"
        Const iDefaultCheckPeriod As Integer = 900000 'Default is 15 minutes (90000ms)
        Const iDefaultDisconnectedLimit As Integer = 600 'Default is 10 hours (600 minutes)

        Dim oRegKey As RegistryKey = Nothing

        Try
            'Open the SyncDesc key with read-only access.
            oRegKey = Registry.LocalMachine.OpenSubKey(sSyncDescKey, False)

            'Get the check period
            Dim iCheckPeriod As Integer = oRegKey.GetValue(sCheckPeriodValue)
            'If configured value read and is valid (1 minute or greater), set it as the check period.
            If iCheckPeriod > 60000 Then
                CheckPeriod = iCheckPeriod
            Else
                'Otherwise, use default value.
                CheckPeriod = iDefaultCheckPeriod
            End If

            'Get the disconnected limit.
            Dim iDisconnectedLimit As Integer = oRegKey.GetValue(sDisconnectedLimitValue)
            'If configured value read and is valid (1 minute or greater), set it as the disconneced limit.
            If iDisconnectedLimit > 1 Then
                DisconnectedLimit = iDisconnectedLimit
            Else
                'Otherwise, use default value.
                DisconnectedLimit = iDefaultDisconnectedLimit
            End If

        Catch ex As Exception
            'Use default values for everything if there was a problem opening the registry.
            CheckPeriod = iDefaultCheckPeriod
            DisconnectedLimit = iDefaultDisconnectedLimit
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
