Public Class frmSplash

    Private counter As Integer = 0

    Private prg As New ProgressBar()
    Private lblPercent As New Label()
    Private lblStatus As New Label()
    Private tmr As New Timer()

    Private Sub frmSplash_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.Text = "Loading..."
        Me.Size = New Size(600, 400)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.FormBorderStyle = FormBorderStyle.None
        Me.BackColor = Color.White
        Me.Opacity = 0

        Dim pic As New PictureBox()

        pic.Size = New Size(130, 130)
        pic.Location = New Point(235, 30)
        pic.SizeMode = PictureBoxSizeMode.Zoom

        Try
            pic.Image = Image.FromFile(Application.StartupPath & "\Images\logo.png")
        Catch
        End Try

        Me.Controls.Add(pic)

        Dim lblTitle As New Label()

        lblTitle.Text = "GHC PET MART"
        lblTitle.Font = New Font("Arial", 22, FontStyle.Bold)
        lblTitle.AutoSize = True
        lblTitle.Location = New Point(165, 170)

        Me.Controls.Add(lblTitle)

        Dim lblTag As New Label()

        lblTag.Text = "One Home for Fur, Feathers, and Fins."
        lblTag.Font = New Font("Arial", 10, FontStyle.Italic)
        lblTag.ForeColor = Color.Gray
        lblTag.AutoSize = True
        lblTag.Location = New Point(150, 205)

        Me.Controls.Add(lblTag)

        lblStatus.Text = "Initializing..."
        lblStatus.Font = New Font("Arial", 9)
        lblStatus.ForeColor = Color.DimGray
        lblStatus.AutoSize = True
        lblStatus.Location = New Point(100, 265)

        Me.Controls.Add(lblStatus)

        prg.Size = New Size(400, 20)
        prg.Location = New Point(100, 290)
        prg.Maximum = 100

        Me.Controls.Add(prg)

        lblPercent.Text = "0%"
        lblPercent.Font = New Font("Arial", 9, FontStyle.Bold)
        lblPercent.AutoSize = True
        lblPercent.Location = New Point(510, 291)

        Me.Controls.Add(lblPercent)


        Dim lblCopy As New Label()

        lblCopy.Text = "© 2026 GHC Pet Mart"
        lblCopy.ForeColor = Color.Gray
        lblCopy.Font = New Font("Arial", 8)
        lblCopy.AutoSize = True
        lblCopy.Location = New Point(230, 350)

        Me.Controls.Add(lblCopy)

        tmr.Interval = 30
        AddHandler tmr.Tick, AddressOf LoadingProgress
        tmr.Start()

    End Sub

    Private Sub LoadingProgress(sender As Object, e As EventArgs)

        If Me.Opacity < 1 Then
            Me.Opacity += 0.05
        End If

        counter += 2

        If counter > 100 Then counter = 100

        prg.Value = counter

        lblPercent.Text = counter & "%"

        Select Case counter

            Case 0 To 20
                lblStatus.Text = "Loading System..."

            Case 21 To 40
                lblStatus.Text = "Loading Products..."

            Case 41 To 60
                lblStatus.Text = "Connecting Database..."

            Case 61 To 80
                lblStatus.Text = "Preparing User Interface..."

            Case 81 To 99
                lblStatus.Text = "Almost Ready..."

            Case 100
                lblStatus.Text = "Welcome!"

        End Select

        If counter >= 100 Then

            tmr.Stop()

            Dim login As New frmLogin()

            login.Show()

            Me.Hide()

        End If

    End Sub

End Class