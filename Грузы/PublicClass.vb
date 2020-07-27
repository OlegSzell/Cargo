Public Class PublicClass

    Private Shared WithEvents myTimer As New Timer()
    Private Shared alarmCounter As Integer = 1
    Private Shared exitFlag As Boolean = False
    Private Shared Счетчик As Integer
    Private Shared ВремяПоказа As New List(Of String)

    ' This is the method to run when the timer is raised.
    Private Shared Sub TimerEventProcessor(myObject As Object, ByVal myEventArgs As EventArgs) Handles myTimer.Tick
        myTimer.Stop()

        Obrabotka()


        '' Displays a message box asking whether to continue running the timer.
        'If MessageBox.Show("Continue running?", "Count is: " & alarmCounter,
        '                    MessageBoxButtons.YesNo) = DialogResult.Yes Then
        '    ' Restarts the timer and increments the counter.
        '    alarmCounter += 1
        '    myTimer.Enabled = True
        'Else
        '    ' Stops the timer.
        '    exitFlag = True
        'End If
    End Sub
    Public Shared Property _Календарь As Dictionary(Of String, String)

    Public Shared Sub Obrabotka()




        Dim f As DateTime = DateTime.Now
        Dim f1 As String = f.Hour

        If Счетчик = 0 Or f1 = "18" Then

            exitFlag = True
            Exit Sub
        End If
        Dim fks As String
        Dim g As String
        For Each fk As KeyValuePair(Of String, String) In _Календарь
            If fk.Key.Length = 5 Then
                fks = Left(fk.Key, 2)
            Else
                fks = Left(fk.Key, 1)
            End If


            If fks = f1 And Not ВремяПоказа.Contains(fk.Key) Then
                g = fk.Value
                myTimer.Stop()
                Dim n As New КалендарьМодалноеОкно(g)
                n.ShowDialog()
                myTimer.Enabled = True
                If КалендарьПовтор = False Then
                    Счетчик -= 1
                    ВремяПоказа.Add(fk.Key)
                End If

            End If

        Next


    End Sub


    Sub New(ByVal d As Dictionary(Of String, String))
        _Календарь = d
        Счетчик = _Календарь.Count
        ВремяПоказа.Add("25")
        'Obrabotka()

        '' Adds the event and the event handler for the method that will
        '' process the timer event to the timer.

        '' Sets the timer interval to 5 seconds.
        '' myTimer.Interval = 5000
        ''myTimer.Start()

        myTimer.Interval = 1200000 '20 минут
        myTimer.Start()


        ' Runs the timer, and raises the event.
        While exitFlag = False
            ' Processes all the events in the queue.
            Application.DoEvents()
        End While

    End Sub

End Class



