Imports WMPDXMLib
Imports WMPLib

Public Class Radio
    Dim f As New WindowsMediaPlayer
    Private Sub Radio_Load(sender As Object, e As EventArgs) Handles MyBase.Load



    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If f.status.Contains("Воспроизведение") Then
            f.controls.stop()
            f = New WindowsMediaPlayer
        End If
        'f.URL = "http:\\www.europaplus.ua\store\public\europaplus64.m3u"

        f.URL = "https://rusradio.hostingradio.ru/rusradio128.mp3"
        f.controls.play()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If f.status.Contains("Воспроизведение") Then
            f.controls.stop()
            f = New WindowsMediaPlayer
        End If
        'f.URL = "http:\\www.europaplus.ua\store\public\europaplus64.m3u"

        f.URL = "http://listen1.myradio24.com:9000/5967"
        f.controls.play()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If f.status.Contains("Воспроизведение") Then
            f.controls.stop()
            f = New WindowsMediaPlayer
        End If

        'f.URL = "http:\\www.europaplus.ua\store\public\europaplus64.m3u"

        f.URL = "http://air.volna.top/HypeFM"
        f.controls.play()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If f.status.Contains("Воспроизведение") Then
            f.controls.stop()
        End If
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click

        If f.status.Contains("Воспроизведение") Then
            f.controls.stop()
            f = New WindowsMediaPlayer
        End If
        f.URL = "http://air.radiorecord.ru:8102/deep_320"
        f.controls.play()


    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        If f.status.Contains("Воспроизведение") Then
            f.controls.stop()
            f = New WindowsMediaPlayer
        End If
        f.URL = "http://air.radiorecord.ru:8102/brks_320"
        f.controls.play()
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click

        If f.status.Contains("Воспроизведение") Then
            f.controls.stop()
            f = New WindowsMediaPlayer
        End If
        f.URL = "http://185.220.35.56:8000/128"
        f.controls.play()


    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click

        If f.status.Contains("Воспроизведение") Then
            f.controls.stop()
            f = New WindowsMediaPlayer
        End If
        f.URL = "http://zaycevfm.cdnvideo.ru/ZaycevFM_club_128.mp3"
        f.controls.play()
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        If f.status.Contains("Воспроизведение") Then
            f.controls.stop()
            f = New WindowsMediaPlayer
        End If
        f.URL = "http://zaycevfm.cdnvideo.ru/ZaycevFM_pop_128.mp3"
        f.controls.play()
    End Sub
End Class