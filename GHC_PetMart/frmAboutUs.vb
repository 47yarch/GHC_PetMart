Imports MySql.Data.MySqlClient
Imports WMPLib
Public Class frmAboutUs

    ' ---- Controls ----
    Private pbBanner As New PictureBox()      ' Large picture / animated .gif banner (500 x 180 px)

    ' Stat cards (replaces the small picture boxes)
    Private lblStat1Num As New Label()
    Private lblStat1Cap As New Label()
    Private lblStat2Num As New Label()
    Private lblStat2Cap As New Label()
    Private lblStat3Num As New Label()
    Private lblStat3Cap As New Label()

    Private lblTitle As New Label()
    Private lblSub As New Label()
    Private rtbAbout As New RichTextBox()
    Private lblDevHeader As New Label()
    Private lblDevelopers As New Label()
    Private btnClose As New Button()
    Private player As WindowsMediaPlayer

    ' Handler kept as a field so AddHandler/RemoveHandler refer to the exact same delegate
    Private bannerFrameHandler As EventHandler

    Private Sub InitAboutUs() Handles MyBase.Load
        Me.Text = "GHC Pet Mart - About Us"
        Me.Size = New Size(560, 650)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.FormBorderStyle = FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.BackColor = Color.White

        ' ===== Large picture (.gif banner) =====
        pbBanner.Size = New Size(500, 180)
        pbBanner.Location = New Point(20, 20)
        pbBanner.BackColor = Color.WhiteSmoke
        pbBanner.BorderStyle = BorderStyle.FixedSingle
        pbBanner.SizeMode = PictureBoxSizeMode.Zoom
        Me.Controls.Add(pbBanner)

        ' TODO: point this to your own .gif file, e.g.:
        ' pbBanner.Image = Image.FromFile(Application.StartupPath & "\Images\banner.gif")
        LoadImageSafely(pbBanner, Application.StartupPath & "\Images\banner.gif")

        ' ===== Stat cards row (Customers / Products / Categories) =====
        Dim statY As Integer = 210
        BuildStatCard(lblStat1Num, lblStat1Cap, "Customers", New Point(40, statY))
        BuildStatCard(lblStat2Num, lblStat2Cap, "Products", New Point(210, statY))
        BuildStatCard(lblStat3Num, lblStat3Cap, "Categories", New Point(380, statY))

        LoadStats()

        ' ===== Title =====
        lblTitle.Text = "GHC PET MART"
        lblTitle.Font = New Font("Arial", 16, FontStyle.Bold)
        lblTitle.ForeColor = Color.FromArgb(63, 81, 181)
        lblTitle.AutoSize = True
        lblTitle.Location = New Point(180, 300)
        Me.Controls.Add(lblTitle)

        lblSub.Text = "Your One-Stop Shop for Happy, Healthy Pets"
        lblSub.Font = New Font("Arial", 9, FontStyle.Italic)
        lblSub.ForeColor = Color.Gray
        lblSub.AutoSize = True
        lblSub.Location = New Point(140, 330)
        Me.Controls.Add(lblSub)

        ' ===== About the shop =====
        rtbAbout.Location = New Point(30, 360)
        rtbAbout.Size = New Size(480, 130)
        rtbAbout.Font = New Font("Arial", 9.5!)
        rtbAbout.ReadOnly = True
        rtbAbout.BorderStyle = BorderStyle.FixedSingle
        rtbAbout.BackColor = Color.WhiteSmoke
        rtbAbout.Text =
            "GHC Pet Mart is a point-of-sale and inventory management system built " &
            "to help small pet shops run their day-to-day operations with ease. " &
            "From tracking products and suppliers to processing sales and monitoring " &
            "stock levels, GHC Pet Mart brings everything a growing pet store needs " &
            "into one simple, easy-to-use application." & vbCrLf & vbCrLf &
            "This system was developed as a school project to apply the principles " &
            "of database-driven application development using VB.NET and MySQL."
        Me.Controls.Add(rtbAbout)

        ' ===== Developers =====
        lblDevHeader.Text = "Developed by:"
        lblDevHeader.Font = New Font("Arial", 10, FontStyle.Bold)
        lblDevHeader.ForeColor = Color.FromArgb(63, 81, 181)
        lblDevHeader.AutoSize = True
        lblDevHeader.Location = New Point(30, 505)
        Me.Controls.Add(lblDevHeader)

        lblDevelopers.Text =
            "Agustin, Jeshaiah D." & vbCrLf &
            "Sabulao, John Michael P." & vbCrLf & vbCrLf &
            "BSIT-T-1B-T" & vbCrLf &
            "Technological University of the Philippines - Taguig"
        lblDevelopers.Font = New Font("Arial", 9.5!)
        lblDevelopers.ForeColor = Color.Black
        lblDevelopers.AutoSize = True
        lblDevelopers.Location = New Point(30, 528)
        Me.Controls.Add(lblDevelopers)

        ' ===== Close button =====
        btnClose.Text = "CLOSE"
        btnClose.Size = New Size(120, 38)
        btnClose.Location = New Point(390, 570)
        btnClose.BackColor = Color.FromArgb(63, 81, 181)
        btnClose.ForeColor = Color.White
        btnClose.FlatStyle = FlatStyle.Flat
        btnClose.Font = New Font("Arial", 10, FontStyle.Bold)
        AddHandler btnClose.Click, AddressOf OnCloseClick
        Me.Controls.Add(btnClose)

        ' Start animating the banner if it turned out to be an animated .gif
        StartBannerAnimation()
        StartBannerMusic()
    End Sub

    ''' <summary>
    ''' Creates one stat "card": a big bold number on top of a small caption,
    ''' both centered under a shared 150px-wide slot.
    ''' </summary>
    Private Sub BuildStatCard(numLabel As Label, capLabel As Label, caption As String, location As Point)
        numLabel.Text = "0"
        numLabel.Font = New Font("Arial", 20, FontStyle.Bold)
        numLabel.ForeColor = Color.FromArgb(63, 81, 181)
        numLabel.AutoSize = False
        numLabel.TextAlign = ContentAlignment.MiddleCenter
        numLabel.Size = New Size(150, 40)
        numLabel.Location = location
        Me.Controls.Add(numLabel)

        capLabel.Text = caption
        capLabel.Font = New Font("Arial", 9)
        capLabel.ForeColor = Color.Gray
        capLabel.AutoSize = False
        capLabel.TextAlign = ContentAlignment.MiddleCenter
        capLabel.Size = New Size(150, 20)
        capLabel.Location = New Point(location.X, location.Y + 40)
        Me.Controls.Add(capLabel)
    End Sub

    ''' <summary>
    ''' Pulls the live counts from the database, matching the same queries
    ''' used on the dashboard (frmMain). Defaults to "0" if the DB isn't reachable.
    ''' </summary>
    Private Sub LoadStats()
        lblStat1Num.Text = GetCount("SELECT COUNT(*) FROM tbl_users WHERE role = 'Customer'")
        lblStat2Num.Text = GetCount("SELECT COUNT(*) FROM tbl_products WHERE status <> 'Discontinued'")
        lblStat3Num.Text = GetCount("SELECT COUNT(*) FROM tbl_categories")
    End Sub

    Private Function GetCount(query As String) As String
        Try
            Using con As MySqlConnection = GetConn()
                Dim cmd As New MySqlCommand(query, con)
                Dim result As Object = cmd.ExecuteScalar()
                Return If(result IsNot Nothing, Convert.ToInt32(result).ToString(), "0")
            End Using
        Catch ex As Exception
            Return "0"
        End Try
    End Function

    ''' <summary>
    ''' Loads an image into a PictureBox only if the file actually exists,
    ''' so the form still opens cleanly before you've added your own pictures.
    ''' </summary>
    Private Sub LoadImageSafely(box As PictureBox, path As String)
        Try
            If IO.File.Exists(path) Then
                box.Image = Image.FromFile(path)
            End If
        Catch ex As Exception
            ' Ignore bad/missing image files so the form still loads
        End Try
    End Sub

    ''' <summary>
    ''' Animated .gif files need ImageAnimator to advance their frames;
    ''' a PictureBox will only show the first frame without this.
    ''' </summary>
    Private Sub StartBannerAnimation()
        If pbBanner.Image IsNot Nothing AndAlso ImageAnimator.CanAnimate(pbBanner.Image) Then
            bannerFrameHandler = New EventHandler(AddressOf OnBannerFrameChanged)
            ImageAnimator.Animate(pbBanner.Image, bannerFrameHandler)
            AddHandler pbBanner.Paint, AddressOf OnBannerPaint
        End If
    End Sub

    Private Sub OnBannerFrameChanged(sender As Object, e As EventArgs)
        If pbBanner.IsHandleCreated Then
            pbBanner.Invalidate()
        End If
    End Sub

    Private Sub OnBannerPaint(sender As Object, e As PaintEventArgs)
        If pbBanner.Image IsNot Nothing Then
            ImageAnimator.UpdateFrames(pbBanner.Image)
        End If
    End Sub

    Private Sub StartBannerMusic()

        Try

            Dim musicPath As String = IO.Path.Combine(Application.StartupPath, "Audio", "banner.mp3")

            If IO.File.Exists(musicPath) Then

                player = New WindowsMediaPlayer()

                player.URL = musicPath
                player.settings.volume = 70
                player.settings.setMode("loop", True)
                player.controls.play()

            Else

                MessageBox.Show("Music file not found:" & vbCrLf & musicPath)

            End If

        Catch ex As Exception

            MessageBox.Show(ex.Message)

        End Try

    End Sub

    Private Sub OnCloseClick(sender As Object, e As EventArgs)
        Me.Close()
    End Sub

    Private Sub frmAboutUs_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing

        If pbBanner.Image IsNot Nothing AndAlso bannerFrameHandler IsNot Nothing Then
            ImageAnimator.StopAnimate(pbBanner.Image, bannerFrameHandler)
        End If

        If player IsNot Nothing Then
            player.controls.stop()
            player.close()
            player = Nothing
        End If

    End Sub

End Class