Imports System.Serviceprocess
Imports Microsoft.Win32
Imports System.IO
Imports System.Management


Public Class Remote_Shell_Command_Service
    Inherits System.Serviceprocess.ServiceBase

    Dim TimerInterval As String = ""
    Dim MACAddress As String = ""
    Dim Proxy As String = ""
    Dim CommandURL As String = ""
    Dim ResponsesURL As String = ""
    Dim RespondURL As String = ""

#Region " Component Designer generated code "

    Public Sub New()
        MyBase.New()

        ' This call is required by the Component Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call
    End Sub

    'UserService overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    ' The main entry point for the process
    <MTAThread()> _
    Shared Sub Main()
        Dim ServicesToRun() As System.Serviceprocess.ServiceBase

        ' More than one NT Service may run within the same process. To add
        ' another service to this process, change the following line to
        ' create a second service object. For example,
        '
        '   ServicesToRun = New System.Serviccurrentprocessess.ServiceBase () {New Service1, New MySecondUserService}
        '
        ServicesToRun = New System.Serviceprocess.ServiceBase() {New Remote_Shell_Command_Service}

        System.ServiceProcess.ServiceBase.Run(ServicesToRun)
    End Sub

    'Required by the Component Designer
    Private components As System.ComponentModel.IContainer

    ' NOTE: The following procedure is required by the Component Designer
    ' It can be modified using the Component Designer.  
    ' Do not modify it using the code editor.
    Friend WithEvents Main_Timer As System.Timers.Timer
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.Main_Timer = New System.Timers.Timer
        CType(Me.Main_Timer, System.ComponentModel.ISupportInitialize).BeginInit()
        '
        'Main_Timer
        '
        Me.Main_Timer.Enabled = True
        Me.Main_Timer.Interval = 60000
        '
        'Remote_Shell_Command_Service
        '
        Me.ServiceName = "DOS_Shell_Rollout"
        CType(Me.Main_Timer, System.ComponentModel.ISupportInitialize).EndInit()

    End Sub

#End Region

    Protected Overrides Sub OnStart(ByVal args() As String)
        ' Add code here to start your service. This method should set things
        ' in motion so your service can do its work.
        Try
            Activity_Logger("Service Successfully Started")
            Code_Execute()
            Main_Timer.Start()
        Catch ex As Exception
            Error_Handler(ex, "OnStart")
        End Try
    End Sub

    Protected Overrides Sub OnStop()
        ' Add code here to perform any tear-down necessary to stop your service.
        Try
            Main_Timer.Stop()
            Activity_Logger("Service Successfully Stopped")
        Catch ex As Exception
            Error_Handler(ex, "OnStop")
        End Try
    End Sub

    Function GetMACAddress() As String
        Dim result As String = ""
        Try

            Dim mc As System.Management.ManagementClass
            Dim mo As ManagementObject
            mc = New ManagementClass("Win32_NetworkAdapterConfiguration")
            Dim moc As ManagementObjectCollection = mc.GetInstances()
            Dim found As Integer = 0
            For Each mo In moc
                If mo.Item("IPEnabled") = True Then
                    If found = 0 Then
                        result = (mo.Item("MacAddress").ToString())
                    End If
                    found = found + 1
                End If
            Next
            If found = 0 Then
                result = ("Fail. No IPEnabled MAC addresses located.")
            End If
        Catch ex As Exception
            result = "Fail. Check Error Log for details."
            Error_Handler(ex, "GetMACAddress")
        End Try
        Return result
    End Function

    Private Sub GetVariables()
        Try
            Dim oldTimerInterval As String = TimerInterval
            Dim oldMACAddress As String = MACAddress
            Dim oldProxy As String = Proxy
            Dim oldCommandURL As String = CommandURL
            Dim oldResponsesURL As String = ResponsesURL
            Dim oldRespondURL As String = RespondURL

            TimerInterval = ""
            MACAddress = ""
            Proxy = ""
            CommandURL = ""
            ResponsesURL = ""
            RespondURL = ""

            MACAddress = GetMACAddress()
            If Not MACAddress = oldMACAddress Then
                Activity_Logger("MAC Address: " & MACAddress)
            End If
            If MACAddress.StartsWith("Fail") = True Then
                MACAddress = ""
            End If

            Dim finfo As FileInfo = New FileInfo((System.Environment.SystemDirectory & "\").Replace("\\", "\") & "Remote Shell Command Service.ini")
            If finfo.Exists = True Then
                Dim filereader As StreamReader = New StreamReader((System.Environment.SystemDirectory & "\").Replace("\\", "\") & "Remote Shell Command Service.ini")
                While filereader.Peek() > -1
                    Dim lineread As String = filereader.ReadLine()
                    If lineread.StartsWith("TimerInterval") Then
                        TimerInterval = lineread.Replace("TimerInterval=", "")
                    End If
                    If lineread.StartsWith("Proxy") Then
                        Proxy = lineread.Replace("Proxy=", "")
                    End If

                    If lineread.StartsWith("CommandURL") Then
                        CommandURL = lineread.Replace("CommandURL=", "")
                    End If
                    If lineread.StartsWith("ResponsesURL") Then
                        ResponsesURL = lineread.Replace("ResponsesURL=", "")
                    End If
                    If lineread.StartsWith("RespondURL") Then
                        RespondURL = lineread.Replace("RespondURL=", "")
                    End If
                End While
                filereader.Close()
                filereader = Nothing
            Else
                    Activity_Logger("No Config File Located: " & (System.Environment.SystemDirectory & "\").Replace("\\", "\") & "Remote Shell Command Service.ini")
                    Error_Handler("No Config File Located: " & (System.Environment.SystemDirectory & "\").Replace("\\", "\") & "Remote Shell Command Service.ini")
            End If
            finfo = Nothing
            If Not TimerInterval = oldTimerInterval Then
                Activity_Logger("Timer Inteval: " & TimerInterval & " secs")
                Main_Timer.Interval = TimerInterval * 1000
            End If
            If Not Proxy = oldProxy Then
                Activity_Logger("Proxy: " & Proxy)
            End If
            If Not CommandURL = oldCommandURL Then
                Activity_Logger("Command URL: " & CommandURL)
            End If
            If Not ResponsesURL = oldResponsesURL Then
                Activity_Logger("Responses URL: " & ResponsesURL)
            End If
            If Not RespondURL = oldRespondURL Then
                Activity_Logger("Respond URL: " & RespondURL)
            End If
        Catch ex As Exception
            Error_Handler(ex, "GetVariables")
        End Try
    End Sub

    Private Sub Error_Handler(ByVal ex As Exception, Optional ByVal identifier_msg As String = "")
        Try
            Dim dir As DirectoryInfo = New DirectoryInfo((System.Environment.SystemDirectory & "\").Replace("\\", "\") & "Remote Shell Command Service\Error Logs")
            If dir.Exists = False Then
                dir.Create()
            End If
            Dim filewriter As StreamWriter = New StreamWriter((System.Environment.SystemDirectory & "\").Replace("\\", "\") & "Remote Shell Command Service\Error Logs\" & Format(Now(), "yyyyMMdd") & "_Error_Log.txt", True)
            filewriter.WriteLine("#" & Format(Now(), "dd/MM/yyyy hh:mm:ss tt") & " - " & identifier_msg & ":" & ex.ToString)
            filewriter.Flush()
            filewriter.Close()
        Catch exc As Exception
            Dim mylog As New EventLog
            If Not mylog.SourceExists("Remote Shell Command Service") Then
                mylog.CreateEventSource("Remote Shell Command Service", "Remote Shell Command Service Log")
            End If
            mylog.Source = "Remote Shell Command Service"
            mylog.WriteEntry("Remote Shell Command Service Log", "Error Handler Failure: " & exc.ToString, EventLogEntryType.Error)
            mylog.Close()
        End Try
    End Sub

    Private Sub Error_Handler(ByVal identifier_msg As String)
        Try
            Dim dir As DirectoryInfo = New DirectoryInfo((System.Environment.SystemDirectory & "\").Replace("\\", "\") & "Remote Shell Command Service\Error Logs")
            If dir.Exists = False Then
                dir.Create()
            End If
            Dim filewriter As StreamWriter = New StreamWriter((System.Environment.SystemDirectory & "\").Replace("\\", "\") & "Remote Shell Command Service\Error Logs\" & Format(Now(), "yyyyMMdd") & "_Error_Log.txt", True)
            filewriter.WriteLine("#" & Format(Now(), "dd/MM/yyyy hh:mm:ss tt") & " - " & identifier_msg)
            filewriter.Flush()
            filewriter.Close()
        Catch ex As Exception
            Error_Handler(ex, "Message Error Handler")
        End Try
    End Sub

    Public Sub Activity_Logger(ByVal message As String)
        Try
            Dim dir As DirectoryInfo = New DirectoryInfo((System.Environment.SystemDirectory & "\").Replace("\\", "\") & "Remote Shell Command Service\Activity Logs")
            If dir.Exists = False Then
                dir.Create()
            End If
            Dim filewriter As StreamWriter = New StreamWriter((System.Environment.SystemDirectory & "\").Replace("\\", "\") & "Remote Shell Command Service\Activity Logs\" & Format(Now(), "yyyyMMdd") & "_Activity_Log.txt", True)
            filewriter.WriteLine("#" & Format(Now(), "dd/MM/yyyy hh:mm:ss tt") & " - " & message)
            filewriter.Flush()
            filewriter.Close()
        Catch ex As Exception
            Error_Handler(ex, "Activity_Logger")
        End Try
    End Sub


    Private Sub Main_Timer_Elapsed(ByVal sender As System.Object, ByVal e As System.Timers.ElapsedEventArgs) Handles Main_Timer.Elapsed
        Try
            Code_Execute()
        Catch ex As Exception
            Error_Handler(ex, "Code_Execute")
        End Try
    End Sub

    Function WebRequest(ByVal url As String) As String
        Try
            Dim sServer As String
            sServer = Trim(url)
            Dim HttpWReq As System.Net.HttpWebRequest = _
                      CType(System.Net.WebRequest.Create(sServer), System.Net.HttpWebRequest)

            Dim proxyObject As New System.Net.WebProxy(Proxy, True)
            HttpWReq.Proxy = proxyObject
            Dim HttpWResp As System.Net.HttpWebResponse = _
               CType(HttpWReq.GetResponse(), System.Net.HttpWebResponse)
            HttpWResp.Close()
            HttpWResp = Nothing
            HttpWReq = Nothing
        Catch ex As Exception
            Error_Handler(ex, "WebRequest")
        End Try
        Return ""
    End Function

    Function WebRequestDoWork(ByVal url As String, ByVal mac As String) As Boolean
        Dim dowork As Boolean = True
        Try
            Dim sServer As String
            sServer = Trim(url)

            Dim HttpWReq As System.Net.HttpWebRequest = _
                      CType(System.Net.WebRequest.Create(sServer), System.Net.HttpWebRequest)

            Dim proxyObject As New System.Net.WebProxy(Proxy, True)
            HttpWReq.Proxy = proxyObject

            Dim HttpWResp As System.Net.HttpWebResponse = _
               CType(HttpWReq.GetResponse(), System.Net.HttpWebResponse)


            Dim streamer As System.IO.StreamReader = New System.IO.StreamReader(HttpWResp.GetResponseStream, System.Text.Encoding.ASCII, False, 512)

            Dim stringtoanalyze As String
            Dim substring As String
            Dim acceptable As String
            Dim addeditem As Boolean
            Dim linetocheck As String = ""
            While streamer.Peek() <> -1
                linetocheck = streamer.ReadLine
                If linetocheck.Replace("<br>", "") = mac Then
                    dowork = False
                    Exit While
                End If
            End While


            streamer.Close()
            HttpWResp.Close()
            streamer = Nothing
            HttpWResp = Nothing
            HttpWReq = Nothing
        Catch ex As Exception
            Error_Handler(ex, "WebRequest")
        End Try
        Return dowork

    End Function

    Public Sub Code_Execute()
        Main_Timer.Stop()
        Try

            GetVariables()
            If Not Proxy = "" And Not MACAddress = "" And Not CommandURL = "" And Not ResponsesURL = "" And Not RespondURL = "" Then
                Dim urlstring As String = CommandURL

                Dim sServer As String
                sServer = Trim(urlstring)

                Dim HttpWReq As System.Net.HttpWebRequest = _
                          CType(System.Net.WebRequest.Create(sServer), System.Net.HttpWebRequest)

                Dim proxyObject As New System.Net.WebProxy(Proxy, True)
                HttpWReq.Proxy = proxyObject

                Dim HttpWResp As System.Net.HttpWebResponse = _
                   CType(HttpWReq.GetResponse(), System.Net.HttpWebResponse)

                Dim streamer As System.IO.StreamReader = New System.IO.StreamReader(HttpWResp.GetResponseStream, System.Text.Encoding.ASCII, False, 512)

                Dim stringtoanalyze As String
                Dim substring As String
                Dim acceptable As String
                Dim addeditem As Boolean
                While streamer.Peek() <> -1
                    'Activity_Logger(streamer.ReadLine)
                    If streamer.ReadLine = "<!--start command-->" Then
                        Try


                            Dim idno As String = streamer.ReadLine.Replace("<br>", "")
                            'ApplicationLauncher(streamer.ReadLine.Replace("<br>", ""))
                            Dim progtorun As String = streamer.ReadLine.Replace("<br>", "")
                            If WebRequestDoWork(ResponsesURL & "?SC_ID=" & idno, MACAddress) = True Then
                                Activity_Logger("Executing: " & progtorun)
                                WebRequest(RespondURL & "?Page_Action=create&SC_ID=" & idno & "&SC_ResponseMAC=" & MACAddress & "")
                                DosShellCommand(progtorun)
                            End If
                        Catch ex As Exception
                            Error_Handler(ex)
                        End Try
                    End If
                End While

                streamer.Close()
                HttpWResp.Close()
                streamer = Nothing
                HttpWResp = Nothing
                HttpWReq = Nothing
            End If
        Catch ex As Exception
            Error_Handler(ex, "Code_Execute")
        End Try
        Main_Timer.Start()
    End Sub

    Private Function DosShellCommand(ByVal AppToRun As String) As String
        Dim s As String = ""
        Try
            Dim myProcess As Process = New Process

            myProcess.StartInfo.FileName = "cmd.exe"
            myProcess.StartInfo.UseShellExecute = False
            myProcess.StartInfo.CreateNoWindow = True
            myProcess.StartInfo.RedirectStandardInput = True
            myProcess.StartInfo.RedirectStandardOutput = True
            myProcess.StartInfo.RedirectStandardError = True
            myProcess.Start()
            Dim sIn As StreamWriter = myProcess.StandardInput
            sIn.AutoFlush = True

            Dim sOut As StreamReader = myProcess.StandardOutput
            Dim sErr As StreamReader = myProcess.StandardError
            sIn.Write(AppToRun & _
               System.Environment.NewLine)
            sIn.Write("exit" & System.Environment.NewLine)
            s = sOut.ReadToEnd()
            If Not myProcess.HasExited Then
                myProcess.Kill()
            End If

            'MessageBox.Show("The 'dir' command window was closed at: " & myProcess.ExitTime & "." & System.Environment.NewLine & "Exit Code: " & myProcess.ExitCode)

            sIn.Close()
            sOut.Close()
            sErr.Close()
            myProcess.Close()
            'MessageBox.Show(s)
        Catch ex As Exception
            Error_Handler(ex, "DOSShellCommand")
        End Try
        Return s
    End Function
End Class
