Imports System.Data.SqlClient

Public Class SQLDependency2
    Dim changeCount As Integer = 0
    Dim queueName = "SELECT Маршрут FROM РейсыКлиента"

    Private Sub ExecuteWatchingQuery()
        Using connection As SqlConnection = New SqlConnection(ConString)
            connection.Open()
            Using command = New SqlCommand("SELECT Маршрут FROM РейсыКлиента", connection)

                'Dim sqlDependency = New SqlDependency(command)
                'AddHandler sqlDependency.OnChange, AddressOf OnDatabaseChange
                'command.ExecuteReader()

                Dim dependency As New SqlDependency(command)
                AddHandler dependency.OnChange, AddressOf OnDatabaseChange
                'ListBox1.Items.Clear()

                Using reader = command.ExecuteReader()
                    'While (reader.Read())
                    '    ListBox1.Items.Add(reader.GetValue(0))
                    'End While
                End Using
            End Using
        End Using
    End Sub
    Private Sub OnDatabaseChange(ByVal sender As Object, ByVal e As SqlNotificationEventArgs)
        Try
            Dim info As SqlNotificationInfo = e.Info
            If SqlNotificationInfo.Insert.Equals(info) Or SqlNotificationInfo.Update.Equals(info) Or SqlNotificationInfo.Delete.Equals(info) Then
                changeCount += 1
                Invoke(Sub()
                           Label1.Text = changeCount & " changes have occurred."
                           With ListBox1.Items
                               .Clear()
                               .Add("Info:   " & e.Info.ToString())
                               .Add("Source: " & e.Source.ToString())
                               .Add("Type:   " & e.Type.ToString())
                           End With
                       End Sub)
            End If
            ExecuteWatchingQuery()
        Catch ex As Exception
            MsgBox(ex.StackTrace)
        End Try
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        SqlDependency.Stop(ConString)
        SqlDependency.Start(ConString)
        ExecuteWatchingQuery()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        'Using db As New dbAllDataContext()
        '    Dim f As New РейсыКлиента
        '    f.Маршрут = "ntntntntnt"
        '    db.РейсыКлиента.InsertOnSubmit(f)
        '    db.SubmitChanges()
        'End Using
        Updates3(stroka:="insert into РейсыКлиента (Маршрут) VALUES ('ntntntntnt')")
    End Sub

    Private Sub SQLDependency2_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub


End Class