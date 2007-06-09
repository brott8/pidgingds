Imports GoogleDesktopAPILib
Imports System.Runtime.InteropServices
Imports System.Runtime.Serialization.Formatters.Binary
Imports Microsoft.Win32
Imports System.IO

Public Class clsPidgin
    Const kLogginEnableMessage As String = "In order to allow chat indexing" & vbCrLf & "please enable HTML logging" & vbCrLf & "under Tools->Preferences->Logging"
    Const kComponentGuid As String = "{eec82554-1893-4b09-9f73-623fe2150751}"
    'System.Guid.NewGuid.ToString()
    'Const kPerPassportSettings As String = "Software\\Microsoft\\MSNMessenger\\PerPassportSettings"
    Const kWindowsRunRegKey As String = "Software\\Microsoft\\Windows\\CurrentVersion\\Run"

    Const kMessengerSetupMessage As String = "GDS Messenger Component Setup"
    Const kTrackerName As String = "pidginGDS.dat"

    'Const kPidginPath As String = "C:\Documents and Settings\user\Application Data\.gaim\logs"
    Private kPidginPath As String = System.Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) & "\.purple\logs"

    Private setting As New My.MySettings
    Private user_name As String = ""
    Private buddy_name As String = ""

    Private MessageQueue As New Queue

    ''' <summary>
    ''' Holds per IM message information 
    ''' </summary>
    Private Structure PidginMessageData
        Public mdate As DateTime
        Public fromFriendlyName As String
        Public toFriendlyName As String
        Public title As String
        'Public sender As String
        Public message As String
        Public sessionId As Integer
    End Structure

    ''' <summary>
    ''' Registers and unregisters the component with GDS
    ''' </summary>
    ''' <param name="install">set to true to register, false otherwise</param>
    ''' <returns>true on success</returns>
    Friend Function DoRegistration(ByVal install As Boolean) As Boolean
        Try
            ' register/unregister the component
            Dim gdsReg As GoogleDesktopIndexingComponentRegisterClass = New GoogleDesktopIndexingComponentRegisterClass()

            If (install) Then
                Dim result As Boolean = MsgBox(kLogginEnableMessage, , kMessengerSetupMessage)

                If (Not result) Then Return False

                Try
                    ' register with GDS
                    Dim componentDesc As Object() = New Object(5) {"Title", "Pidgin GDS Indexer", "Description", "Captures and indexes Pidgin Messenger conversations", "Icon", "no icon"}
                    gdsReg.RegisterIndexingComponent(kComponentGuid, componentDesc)

                    ' add the program to Run
                    Dim key As RegistryKey = Registry.CurrentUser.OpenSubKey(kWindowsRunRegKey, True)
                    key.SetValue(kMessengerSetupMessage, Directory.GetCurrentDirectory() + "\" + Process.GetCurrentProcess().ProcessName + ".exe")

                    MsgBox("Registration completed successfully!", , kMessengerSetupMessage)
                Catch e As COMException
                    ' check if not already registered
                    If Hex(e.ErrorCode) <> "80040006" Then   'CLng("&H80040006") ' E_COMPONENT_ALREADY_REGISTERED
                        MsgBox(e.Message, , "COM Exception")
                        Return False
                    Else
                        MsgBox("Component already registered!", , kMessengerSetupMessage)
                    End If
                End Try

            Else
                ' remove the program from Run
                Dim key As RegistryKey = Registry.CurrentUser.OpenSubKey(kWindowsRunRegKey, True)
                key.DeleteValue(kMessengerSetupMessage, False)

                Try

                    gdsReg.UnregisterIndexingComponent(kComponentGuid)

                Catch e As COMException

                    ' check if not already unregistered
                    If Hex(e.ErrorCode) <> "80040002" Then        ' E_COMPONENT_NOT_REGISTERED
                        MsgBox("Component is not registered." & vbCrLf & e.Message, , "COM Exception")
                        Return False
                    End If
                End Try

                MsgBox("Component unregistered successfully!", , kMessengerSetupMessage)

            End If
        Catch e As Exception
            MsgBox(e.Message, , "Exception")
            Return False
        End Try

        Return True
    End Function

    Friend Sub ResetTrackingData()
        If Directory.Exists(kPidginPath) = False Then Exit Sub

        For Each dirService As String In Directory.GetDirectories(kPidginPath)
            For Each dirUserName As String In Directory.GetDirectories(dirService)
                For Each dirEntry As String In Directory.GetDirectories(dirUserName)
                    For Each fileEntry As String In Directory.GetFiles(dirEntry, "pidginGDS.dat", SearchOption.TopDirectoryOnly)
                        Try
                            File.SetAttributes(fileEntry, FileAttributes.Normal)
                            File.Delete(fileEntry)
                        Catch ex As Exception
                            MsgBox("Unable to remove: " & fileEntry & vbCrLf & ex.Message, MsgBoxStyle.Critical)
                        End Try
                    Next
                Next
            Next
        Next
    End Sub

    ''' <summary>
    ''' Captures Pidgin conversations
    ''' </summary>
    Friend Sub CaptureMessages()
        Dim hashTracker As Hashtable
        Dim lastIndexed As DateTime
        Dim sessionID As Integer
        Const sleepLength As Integer = 60000
        Const pacerSleep As Integer = 80

        '' first get the location of the log files
        'Dim key As RegistryKey = Registry.CurrentUser.OpenSubKey(kPerPassportSettings, True)
        'If (key Is Not Nothing AndAlso key.SubKeyCount > 0) Then
        '    ' enumerate through all passports and save settings
        '    Dim subKeys As String() = key.GetSubKeyNames()
        'passportList = new Passport[subKeys.Length];

        'for (int i=0; i < subKeys.Length; i++) 
        '{
        '  RegistryKey subKey = key.OpenSubKey(subKeys[i]);
        '  passportList[i] = new Passport(subKeys[i], (string)subKey.GetValue(kMessageLogPath));

        '        ' check if we have a value for last indexed
        '  string lastIndexed = (string)subKey.GetValue(kMessageLogLastIndexed);   
        '  subKey.Close();
        '  if (lastIndexed != null)
        '  {
        '    passportList[i].LogLastIndexed = DateTime.Parse(lastIndexed);
        '  }
        '}
        'key.Close();
        '        End If

        ' periodically check which log files timestamps have changed
        ' for each log file that has changed - parse it and send the new messages to GDS
        While (True)
            Dim nowUpdated As DateTime = DateTime.Now
            Dim needTimeUpdate As Boolean = False

            ' check if a chat log exists for this user
            If (Not Directory.Exists(kPidginPath)) Then
                Threading.Thread.Sleep(sleepLength)
                Continue While
            End If

            For Each dirService As String In Directory.GetDirectories(kPidginPath)
                For Each dirUserName As String In Directory.GetDirectories(dirService)
                    For Each dirEntry As String In Directory.GetDirectories(dirUserName)
                        'If dirEntry.EndsWith("UFMike45") = False Then Continue For
                        'If it is a chatroom log, then skip
                        If Path.GetDirectoryName(dirEntry).StartsWith("cr-") = True Then Continue For

                        hashTracker = Nothing

                        If File.Exists(Path.Combine(dirEntry, kTrackerName)) = True Then
                            hashTracker = HashDeserialize(dirEntry)
                        Else
                            hashTracker = New Hashtable
                        End If

                        For Each fileEntry As String In Directory.GetFiles(dirEntry)
                            'If fileEntry.EndsWith("2007-04-30.232931-0400EDT.html") = False Then Continue For
                            If (fileEntry.EndsWith(".html")) Then
                                Threading.Thread.Sleep(pacerSleep) 'Slow Down!!

                                'To prevent the queue from growing too memory intensive, wait until it gets trimmed
                                While MessageQueue.Count > 400
                                    Threading.Thread.Sleep(sleepLength)
                                End While

                                sessionID = GetSessionID(fileEntry)

                                'If hashTracker.ContainsKey(sessionID) Then
                                Try
                                    lastIndexed = hashTracker.Item(sessionID)
                                Catch
                                    lastIndexed = ""
                                End Try

                                If (File.GetLastWriteTime(fileEntry).CompareTo(lastIndexed) > 0) Then
                                    'needTimeUpdate = True
                                    ParseAndSendMessages(fileEntry, lastIndexed)
                                    hashTracker.Item(sessionID) = DateTime.Now
                                    HashSerialize(hashTracker, dirEntry)
                                End If
                            End If
                        Next
                    Next
                Next
            Next

            '' update time stamp of last index check
            'If (needTimeUpdate) Then
            '    'Mark as last updated
            '    setting.LastIndexed = nowUpdated
            '    setting.Save()
            'End If

            ' sleep for some time
            Threading.Thread.Sleep(sleepLength)
        End While
    End Sub

    ''' <summary>
    ''' Parses an IM log file and sends the information to GDS
    ''' </summary>
    ''' <param name="logFile">The IM conversations log file</param>
    ''' <param name="lastIndexed">messages older than this will not be sent to GDS</param>
    Private Sub ParseAndSendMessages(ByVal logFile As String, ByVal lastIndexed As DateTime)
        Dim fs As FileStream = Nothing
        Dim sr As StreamReader = Nothing
        Dim messageData As PidginMessageData = Nothing
        Const pacerSleepRead As Integer = 75

        Try
            Dim title As String = ""
            'Dim sender As String = ""
            Dim di As DirectoryInfo
            Dim logDate As DateTime
            Dim lastLogDate As DateTime = Path.GetFileNameWithoutExtension(logFile).Substring(0, 10)

            fs = New FileStream(logFile, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
            sr = New StreamReader(fs)

            Dim reader As String = sr.ReadLine()

            If reader Is Nothing Then Exit Sub

            Dim findTitle As Integer = reader.IndexOf("<title>Conversation with")

            If findTitle < 0 Then
                Exit Sub
            Else
                Dim findEndTitle As Integer = reader.IndexOf("</title>", findTitle)
                title = reader.Substring(findTitle + 7, findEndTitle - findTitle - 7)
            End If

            'First get the user and buddy names based on the coloring
            'The names in the log must equal what we tell GDS or else erratic formatting can occur
            user_name = ""
            buddy_name = ""
            GetUserAndBuddy(sr)

            di = New DirectoryInfo(Path.GetDirectoryName(logFile))

            If user_name = "" Then user_name = di.Parent.Name
            If buddy_name = "" Then buddy_name = di.Name

            'Calculate a unique session ID
            Dim sessionID As Integer = GetSessionID(logFile)

            'Advance to the second line
            If sr.ReadLine() Is Nothing Then Exit Sub

            Do While sr.Peek() >= 0
                reader = sr.ReadLine
                Threading.Thread.Sleep(pacerSleepRead) 'Slow Down!!

                Try
                    If reader.IndexOf("<font color=""#16569E"">") < 0 AndAlso reader.IndexOf("<font color=""#A82F2F"">") < 0 Then
                        Continue Do
                    End If

                    ' check the date of the message - if older skip
                    Dim findtime As Integer = reader.IndexOf("(")
                    If findtime < 0 Then Continue Do

                    Dim findendtime As Integer = reader.IndexOf(")", findtime)
                    If findendtime < 0 Then Continue Do

                    'If not a timestamp, skip
                    If DateTime.TryParse(reader.Substring(findtime + 1, findendtime - findtime - 1), logDate) = False Then Continue Do

                    'Add a year, month, and day to the time
                    logDate = New DateTime(lastLogDate.Year, lastLogDate.Month, lastLogDate.Day, logDate.Hour, logDate.Minute, logDate.Second)

                    'If the timestamp has flipped to the next day, then reflect in variable
                    If logDate.CompareTo(lastLogDate) < 0 Then logDate = logDate.AddDays(1)

                    lastLogDate = logDate

                    ' if older than the last indexing time, skip the message
                    If logDate.CompareTo(lastIndexed) <= 0 Then
                        Continue Do
                    End If

                    ''Get the sender
                    'Dim startSender As Integer = reader.IndexOf("<b>", findtime)
                    'Dim endSender As Integer = reader.IndexOf(":</b>", startSender)
                    'sender = reader.Substring(startSender + 3, endSender - startSender - 3)

                    'Remove the original sender name so that we can make sure it matches the folders sender name
                    'reader = reader.Remove(startSender + 3, endSender - startSender - 3)

                    'If sender.Replace(" ", "").ToLower = di.Parent.Name.Replace(" ", "").ToLower Then
                    '    reader.Remove(startSender + 3, endSender - startSender - 3)
                    'End If

                    'Chop off the sender from the message
                    'reader = reader.Substring(endSender + 13)

                    ' set up messagedata
                    messageData = New PidginMessageData
                    'messageData.fromFriendlyName = di.Parent.Name
                    'messageData.toFriendlyName = di.Name
                    'messageData.sender = sender
                    messageData.fromFriendlyName = user_name
                    messageData.toFriendlyName = buddy_name
                    messageData.mdate = logDate.ToUniversalTime
                    messageData.title = title
                    messageData.sessionId = sessionID

                    'Get rid of the linebreak at the end of each line
                    If reader.EndsWith("<br/>") Then reader = reader.Substring(0, reader.Length - 5)

                    'Move the auto reply message to ensure the buddyname is next to the colon
                    If reader.IndexOf(" &lt;AUTO-REPLY&gt;:</b></font>") >= 0 Then
                        reader = reader.Replace(" &lt;AUTO-REPLY&gt;:</b></font>", ": &lt;AUTO-REPLY&gt;</b></font>")
                    End If

                    'At this point check if we have "buddyname:" - if not then skip it because GDS is picky
                    If reader.IndexOf(user_name & ":") < 0 AndAlso reader.IndexOf(buddy_name & ":") < 0 Then Continue Do

                    messageData.message = reader

                    'Put each message in the queue to be processed
                    MessageQueue.Enqueue(messageData)

                    'SendMessageData(messageData) ' send this message to GDS for indexing
                Catch
                    'If there are problems with this one, just skip the record
                    Continue Do
                End Try
            Loop
        Catch
            'Skip the log if we have serious errors
        Finally
            Try
                sr.Close()
                fs.Close()
            Catch

            End Try
        End Try

    End Sub

    ''' <summary>
    ''' Parses out the user and buddy names based on color
    ''' </summary>
    ''' <param name="sr">streamreader object pointing to log file</param>
    Private Sub GetUserAndBuddy(ByVal sr As StreamReader)
        Dim contents As String = sr.ReadToEnd()
        Dim startSender As Integer
        Dim endSender As Integer

        Dim findBuddy As Integer = contents.IndexOf("color=""#A82F2F"">")
        'Get the sender
        If findBuddy < 0 Then
            buddy_name = ""
        Else
            startSender = contents.IndexOf("<b>", findBuddy)
            endSender = contents.IndexOf(":</b>", startSender)
            buddy_name = contents.Substring(startSender + 3, endSender - startSender - 3)
            'Filter out auto reply message
            If buddy_name.IndexOf(" &lt;AUTO-REPLY&gt;") >= 0 Then
                buddy_name = buddy_name.Replace(" &lt;AUTO-REPLY&gt;", "")
            End If
        End If

        Dim findUser As Integer = contents.IndexOf("color=""#16569E"">")
        'Get the sender
        If findUser < 0 Then
            user_name = ""
        Else
            startSender = contents.IndexOf("<b>", findUser)
            endSender = contents.IndexOf(":</b>", startSender)
            user_name = contents.Substring(startSender + 3, endSender - startSender - 3)
            'Filter out auto reply message
            If user_name.IndexOf(" &lt;AUTO-REPLY&gt;") >= 0 Then
                user_name = user_name.Replace(" &lt;AUTO-REPLY&gt;", "")
            End If
        End If

        'Set the streamreader back to the beginning of the file
        sr.BaseStream.Position = 0
    End Sub

    ''' <summary>
    ''' Processes the message queue at a controlled pace to send to GDS
    ''' </summary>
    Friend Sub ProcessQueue()
        Const shortSleepLength As Integer = 500
        Const longSleepLength As Integer = 8000
        Const trimQueueLength As Integer = 7

        Dim messageData As PidginMessageData = Nothing
        Dim fullQueue As Boolean = False  'This flag tells if the queue should eventually be trimmed for memory saving

        While (True)
            Try
                If MessageQueue.Count > trimQueueLength AndAlso fullQueue = False Then fullQueue = True
                'Retrieve and remove the next available message (FIFO)
                messageData = MessageQueue.Dequeue()
                'And send the message to GDS
                While SendMessageData(messageData)
                    'This will retry if indicated
                    Threading.Thread.Sleep(60000 * 4) 'Go to sleep for 4 minutes and try again
                End While

                If MessageQueue.Count < trimQueueLength AndAlso fullQueue = True Then
                    MessageQueue.TrimToSize()
                    fullQueue = False
                End If
            Catch exEmpty As InvalidOperationException
                'If the queue is empty, continue checking
            Catch ex As Exception
                'Other error
            Finally
                messageData = Nothing
                If MessageQueue.Count > 0 Then
                    Threading.Thread.Sleep(shortSleepLength)
                Else
                    Threading.Thread.Sleep(longSleepLength)
                End If
            End Try
        End While
    End Sub

    ''' <summary>
    ''' Sends the conversation data to GDS using COM Interop
    ''' </summary>
    ''' <param name="messageData">PidginMessageData struct containing the IM message data</param>
    ''' <returns>Retry value - whether the caller should attempt to retry the send</returns>
    Private Function SendMessageData(ByVal messageData As PidginMessageData) As Boolean

        Try
            'Instantiate the class
            Dim gdsClass As New GoogleDesktopClass()

            ' create the event
            Dim gdsEventDisp As Object = gdsClass.CreateEvent(kComponentGuid, "Google.Desktop.IM")
            'Dim gdsEventDisp As Object = gdsClass.CreateEvent(kComponentGuid, "Google.Desktop.TextFile")
            Dim gdsEvent As IGoogleDesktopEvent = CType(gdsEventDisp, IGoogleDesktopEvent)

            ' add IM event properties

            ' "message_time" property
            'gdsEvent.AddProperty("message_time", messageData.mdate.ToUniversalTime().ToOADate)
            gdsEvent.AddProperty("message_time", messageData.mdate)

            ' "format" property
            gdsEvent.AddProperty("format", "text/html")

            ' "content" property
            gdsEvent.AddProperty("content", messageData.message)

            ' "user_name" property
            gdsEvent.AddProperty("user_name", messageData.fromFriendlyName)

            ' "buddy_name" property
            gdsEvent.AddProperty("buddy_name", messageData.toFriendlyName)

            ' "conversation_id" property
            gdsEvent.AddProperty("conversation_id", messageData.sessionId)

            ' "title" property - use the buddy name
            gdsEvent.AddProperty("title", messageData.title)

            ' send the event real-time
            gdsEvent.Send(CLng("&H00000001"))

        Catch e As COMException

            ' protect some valid error results
            Dim error1 As String = Hex(e.ErrorCode)

            Select Case error1 'Google Desktop is not running
                Case "80040010"
                    Dim ret As Integer = MsgBox("Google Desktop appears to be off or not accepting new index entries." & vbCrLf & _
                        "Would you like to close the PidginGDS indexer?" & vbCrLf & _
                        "If ""NO"", then will check again in 4 minutes.", MsgBoxStyle.YesNo)
                    If ret = vbYes Then End
                    Return True 'Return true to try again

                Case "80040002"
                    MsgBox("If you have not yet done so, you MUST run this application with the register flag" & _
                        vbCrLf & "PidginGDS.exe -register" & vbCrLf & e.Message & vbCrLf & vbCrLf & _
                        "Plugin will now exit.", MsgBoxStyle.Critical, "COM Exception")
                    End

                Case "80040005", "80040008", "80040009"
                    Return False 'These are valid errors according to what I hear

                Case Else
                    MsgBox("Error: " & e.Message & vbCrLf & "HRESULT: " & error1 & vbCrLf & "Please report this message to the developer." & vbCrLf & _
                        "Plugin will now exit.", MsgBoxStyle.Critical, "COM Exception")
                    End
            End Select

        End Try

        Return False
    End Function

    Private Sub HashSerialize(ByVal hash As Hashtable, ByVal write_dir As String)
        Dim fs As Stream = New FileStream(Path.Combine(write_dir, kTrackerName), FileMode.Create)
        Dim bf As BinaryFormatter = New BinaryFormatter
        bf.Serialize(fs, hash)
        fs.Close()
    End Sub

    Private Function HashDeserialize(ByVal read_dir As String) As Hashtable
        Dim fs As Stream = New FileStream(Path.Combine(read_dir, kTrackerName), FileMode.Open)
        Dim bf As BinaryFormatter = New BinaryFormatter()

        Dim hash As Hashtable = CType(bf.Deserialize(fs), Hashtable)
        fs.Close()

        Return hash
    End Function

    Private Function GetSessionID(ByVal logfile As String) As Integer
        Dim di As DirectoryInfo = New DirectoryInfo(Path.GetDirectoryName(logfile))

        'Calculate a unique session ID
        'Use the CRC of username + buddyname + logfilename
        Dim crc32 As New clsCRC32
        Return crc32.GetCrc32(di.Parent.Name & di.Name & Path.GetFileNameWithoutExtension(logfile))
    End Function

    Public Sub New()
        If setting.LogLocation.Trim <> "" Then
            If Directory.Exists(setting.LogLocation) Then
                kPidginPath = setting.LogLocation
            Else
                MsgBox("Your specified log path: """ & setting.LogLocation & """ does not exist!" & _
                    vbCrLf & vbCrLf & "Path has been reset to: " & kPidginPath, MsgBoxStyle.Exclamation)
                setting.LogLocation = kPidginPath
                setting.Save()
            End If
        End If
    End Sub
End Class
