Module Module1

    Sub Main()

        Dim proclist() As Process = Process.GetProcessesByName(Process.GetCurrentProcess().ProcessName)
        Dim argsv As String() = Command.Split(",")
        Dim pidgin As clsPidgin = Nothing

        'Trim each parameter
        For x As Integer = 0 To argsv.Length - 1
            argsv(x) = argsv(x).Trim
        Next

        If proclist.Length > 1 Then
            If argsv.Length > 0 Then
                Select Case argsv(0)
                    Case "-unregister", "-stop"

                    Case Else
                        MsgBox("This process is already running." & vbCrLf & _
                            "Please use the ""-stop"" parameter to end the process before performing this action.", MsgBoxStyle.Exclamation)
                        End
                End Select
            Else
                MsgBox("PidginGDS is already running!", MsgBoxStyle.Exclamation)
                End
            End If
        End If

        ' check command line args
        If argsv.Length > 0 Then
            Select Case argsv(0).Trim

                Case "-register"
                    pidgin = New clsPidgin
                    ' register and if successful launch the component
                    If pidgin.DoRegistration(True) = False Then
                        End
                    End If

                Case "-unregister"
                    pidgin = New clsPidgin
                    For Each p As Process In proclist
                        If p.Id <> Process.GetCurrentProcess().Id Then
                            p.Kill()
                            p.WaitForExit()
                        End If
                    Next
                    pidgin.DoRegistration(False)
                    End

                Case "-stop"
                    If proclist.Length <= 1 Then
                        MsgBox("Process is not running", MsgBoxStyle.Information)
                        End
                    End If
                    For Each p As Process In proclist
                        If p.Id <> Process.GetCurrentProcess().Id Then
                            p.Kill()
                            p.WaitForExit()
                        End If
                    Next
                    MsgBox("Successfully stopped process", MsgBoxStyle.Information)
                    End

                Case "-reset"
                    Dim setting As New My.MySettings
                    setting.Reset()

                    pidgin = New clsPidgin
                    pidgin.ResetTrackingData()

                    MsgBox("Settings reset", MsgBoxStyle.Information)
                    End

                Case "-upgrade"
                    Dim setting As New My.MySettings
                    setting.Upgrade()
                    MsgBox("Settings upgraded", MsgBoxStyle.Information)
                    End

                Case "-logpath"
                    If argsv.Length > 1 Then
                        Dim setting As New My.MySettings
                        setting.LogLocation = argsv(1).Replace("""", "")
                        setting.Save()
                        MsgBox("Log location set to: """ & setting.LogLocation & """", MsgBoxStyle.Information)
                        End
                    End If

                Case ""
                    pidgin = New clsPidgin

                Case Else
                    MsgBox("Invalid parameters. Please try again.", MsgBoxStyle.Critical)
                    End

            End Select

        End If

        'If the process is already running, then quit
        If proclist.Length > 1 Then
            End
        End If

        ' index existing IM messages and monitor and index new conversations
        Try
            Dim threadQueue As New Threading.Thread(AddressOf pidgin.ProcessQueue)
            threadQueue.Start()

            pidgin.CaptureMessages()
        Catch ex As Exception
            MsgBox("Class not initialized." & vbCrLf & ex.Message, MsgBoxStyle.Critical)
            End
        End Try

        'Dim mdata As New clsPidgin.PidginMessageData

        'pidgin.SendMessageData(mdata)

    End Sub

End Module
