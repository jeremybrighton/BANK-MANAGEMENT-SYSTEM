' ============================================================
'  BANK MANAGEMENT SYSTEM - frmDashboard.vb
'  Main navigation hub with live statistics tiles
' ============================================================
Imports System.Drawing
Imports System.Windows.Forms

Public Class frmDashboard
    Inherits Form

    Private WithEvents btnCustomers As Button
    Private WithEvents btnTransactions As Button
    Private WithEvents btnEmployees As Button
    Private WithEvents btnReports As Button
    Private WithEvents btnSearch As Button
    Private WithEvents btnLogout As Button
    Private WithEvents refreshTimer As System.Windows.Forms.Timer

    Private pnlHeader As Panel
    Private pnlStats As Panel
    Private lblWelcome As Label
    Private lblDateTime As Label

    ' Stat tiles
    Private tileTotalCustomers As Panel
    Private tileTotalBalance As Panel
    Private tileTodayTxn As Panel
    Private tileActivePct As Panel

    Public Sub New()
        BuildUI()
        RefreshStats()
    End Sub

    Private Sub BuildUI()
        Me.Text = APP_TITLE & " — Dashboard"
        Me.Size = New Size(900, 650)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.BackColor = Color.FromArgb(245, 247, 250)
        Me.Font = New Font("Segoe UI", 9.5F)
        Me.FormBorderStyle = FormBorderStyle.FixedSingle
        Me.MaximizeBox = False

        ' ── Header ────────────────────────────────────────────
        pnlHeader = New Panel()
        pnlHeader.Size = New Size(900, 80)
        pnlHeader.BackColor = Color.FromArgb(15, 76, 129)

        Dim lblBank As New Label()
        lblBank.Text = "🏦 " & BANK_NAME
        lblBank.Font = New Font("Segoe UI", 16, FontStyle.Bold)
        lblBank.ForeColor = Color.White
        lblBank.AutoSize = True
        lblBank.Location = New Point(24, 24)

        lblWelcome = New Label()
        lblWelcome.Font = New Font("Segoe UI", 9)
        lblWelcome.ForeColor = Color.FromArgb(180, 210, 240)
        lblWelcome.AutoSize = True
        lblWelcome.Location = New Point(24, 54)
        lblWelcome.Text = $"Welcome, {CurrentEmployee?.FullName} ({CurrentEmployee?.Role})"

        lblDateTime = New Label()
        lblDateTime.Font = New Font("Segoe UI", 9)
        lblDateTime.ForeColor = Color.FromArgb(180, 210, 240)
        lblDateTime.AutoSize = True
        lblDateTime.Location = New Point(700, 30)

        btnLogout = New Button()
        btnLogout.Text = "⬡ Logout"
        btnLogout.Size = New Size(100, 30)
        btnLogout.Location = New Point(780, 25)
        btnLogout.BackColor = Color.FromArgb(200, 60, 60)
        btnLogout.ForeColor = Color.White
        btnLogout.FlatStyle = FlatStyle.Flat
        btnLogout.FlatAppearance.BorderSize = 0
        btnLogout.Cursor = Cursors.Hand

        pnlHeader.Controls.AddRange(New Control() {lblBank, lblWelcome, lblDateTime, btnLogout})

        ' ── Stat Tiles ────────────────────────────────────────
        pnlStats = New Panel()
        pnlStats.Size = New Size(900, 130)
        pnlStats.Location = New Point(0, 80)
        pnlStats.BackColor = Color.FromArgb(235, 240, 248)
        pnlStats.Padding = New Padding(20, 16, 20, 16)

        tileTotalCustomers = MakeTile("Total Customers", "—", Color.FromArgb(15, 76, 129))
        tileTotalCustomers.Location = New Point(20, 16)

        tileTotalBalance = MakeTile("Total Deposits", "—", Color.FromArgb(22, 130, 90))
        tileTotalBalance.Location = New Point(230, 16)

        tileTodayTxn = MakeTile("Today's Transactions", "—", Color.FromArgb(150, 90, 20))
        tileTodayTxn.Location = New Point(440, 16)

        tileActivePct = MakeTile("Active Accounts", "—", Color.FromArgb(100, 40, 150))
        tileActivePct.Location = New Point(650, 16)

        pnlStats.Controls.AddRange(New Control() {tileTotalCustomers, tileTotalBalance, tileTodayTxn, tileActivePct})

        ' ── Navigation Grid ───────────────────────────────────
        Dim navLabel As New Label()
        navLabel.Text = "QUICK NAVIGATION"
        navLabel.Font = New Font("Segoe UI", 8.5F, FontStyle.Bold)
        navLabel.ForeColor = Color.FromArgb(100, 110, 130)
        navLabel.Location = New Point(30, 230)
        navLabel.AutoSize = True

        btnCustomers = MakeNavButton("👥  Customer Accounts",
            "Add, view, modify & delete customer accounts",
            Color.FromArgb(15, 76, 129), New Point(30, 250))

        btnTransactions = MakeNavButton("💰  Transactions",
            "Deposit, Withdraw & Transfer funds",
            Color.FromArgb(22, 130, 90), New Point(330, 250))

        btnSearch = MakeNavButton("🔍  Search Account",
            "Find customer accounts by number or name",
            Color.FromArgb(150, 90, 20), New Point(630, 250))

        btnEmployees = MakeNavButton("🔑  Employee Management",
            "Register and manage bank staff accounts",
            Color.FromArgb(100, 40, 150), New Point(30, 420))

        btnReports = MakeNavButton("📊  Reports & Analytics",
            "Full financial reports and audit trails",
            Color.FromArgb(15, 100, 130), New Point(330, 420))

        ' Status bar
        Dim pnlStatus As New Panel()
        pnlStatus.Size = New Size(900, 30)
        pnlStatus.Location = New Point(0, 615)
        pnlStatus.BackColor = Color.FromArgb(210, 220, 235)

        Dim lblStatus As New Label()
        lblStatus.Text = $"  {BANK_NAME}  |  {APP_TITLE}  |  Data: {DataFolder}"
        lblStatus.Font = New Font("Segoe UI", 8)
        lblStatus.ForeColor = Color.FromArgb(80, 90, 110)
        lblStatus.Dock = DockStyle.Fill
        pnlStatus.Controls.Add(lblStatus)

        Me.Controls.AddRange(New Control() {
            pnlHeader, pnlStats, navLabel,
            btnCustomers, btnTransactions, btnSearch,
            btnEmployees, btnReports, pnlStatus})

        ' Live clock
        refreshTimer = New System.Windows.Forms.Timer()
        refreshTimer.Interval = 1000
        refreshTimer.Start()
    End Sub

    Private Function MakeTile(title As String, value As String, color As Color) As Panel
        Dim p As New Panel()
        p.Size = New Size(195, 95)
        p.BackColor = Color.White

        Dim accent As New Panel()
        accent.Size = New Size(6, 95)
        accent.BackColor = color
        accent.Dock = DockStyle.Left

        Dim lblTitle As New Label()
        lblTitle.Text = title.ToUpper()
        lblTitle.Font = New Font("Segoe UI", 7.5F, FontStyle.Bold)
        lblTitle.ForeColor = Color.FromArgb(120, 130, 150)
        lblTitle.Location = New Point(14, 14)
        lblTitle.AutoSize = True
        lblTitle.Tag = "title"

        Dim lblValue As New Label()
        lblValue.Text = value
        lblValue.Font = New Font("Segoe UI", 18, FontStyle.Bold)
        lblValue.ForeColor = color
        lblValue.Location = New Point(14, 40)
        lblValue.AutoSize = True
        lblValue.Tag = "value"

        p.Controls.AddRange(New Control() {accent, lblTitle, lblValue})
        Return p
    End Function

    Private Sub SetTileValue(tile As Panel, value As String)
        For Each c As Control In tile.Controls
            If c.Tag IsNot Nothing AndAlso c.Tag.ToString() = "value" Then
                c.Text = value
                Return
            End If
        Next
    End Sub

    Private Function MakeNavButton(title As String, subtitle As String, color As Color, loc As Point) As Button
        Dim btn As New Button()
        btn.Size = New Size(260, 140)
        btn.Location = loc
        btn.BackColor = color
        btn.ForeColor = Color.White
        btn.FlatStyle = FlatStyle.Flat
        btn.FlatAppearance.BorderSize = 0
        btn.Cursor = Cursors.Hand
        btn.Text = title & vbCrLf & vbCrLf & subtitle
        btn.Font = New Font("Segoe UI", 10, FontStyle.Regular)
        btn.TextAlign = ContentAlignment.MiddleCenter
        AddHandler btn.MouseEnter, Sub()
            btn.BackColor = ControlPaint.Light(color, 0.15F)
        End Sub
        AddHandler btn.MouseLeave, Sub()
            btn.BackColor = color
        End Sub
        Return btn
    End Function

    Private Sub RefreshStats()
        Try
            Dim report = DataAccess.GenerateSummaryReport()
            Dim txns = DataAccess.LoadAllTransactions()
            Dim todayCount = txns.Count(Function(t) t.Timestamp.Date = DateTime.Today)

            SetTileValue(tileTotalCustomers, report.TotalCustomers.ToString())
            SetTileValue(tileTotalBalance, FormatMoney(report.TotalDeposits))
            SetTileValue(tileTodayTxn, todayCount.ToString())
            SetTileValue(tileActivePct, report.ActiveAccounts.ToString() & " Active")
        Catch
        End Try
    End Sub

    ' ── Event Handlers ────────────────────────────────────────

    Private Sub refreshTimer_Tick(sender As Object, e As EventArgs) Handles refreshTimer.Tick
        lblDateTime.Text = DateTime.Now.ToString("ddd dd MMM yyyy   HH:mm:ss")
    End Sub

    Private Sub btnCustomers_Click(sender As Object, e As EventArgs) Handles btnCustomers.Click
        Dim f As New frmCustomers()
        f.ShowDialog()
        RefreshStats()
    End Sub

    Private Sub btnTransactions_Click(sender As Object, e As EventArgs) Handles btnTransactions.Click
        Dim f As New frmTransactions()
        f.ShowDialog()
        RefreshStats()
    End Sub

    Private Sub btnSearch_Click(sender As Object, e As EventArgs) Handles btnSearch.Click
        Dim f As New frmSearch()
        f.ShowDialog()
    End Sub

    Private Sub btnEmployees_Click(sender As Object, e As EventArgs) Handles btnEmployees.Click
        If Not IsAdminSession Then
            MessageBox.Show("Admin or Manager access required.", "Access Denied",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If
        Dim f As New frmEmployees()
        f.ShowDialog()
    End Sub

    Private Sub btnReports_Click(sender As Object, e As EventArgs) Handles btnReports.Click
        Dim f As New frmReports()
        f.ShowDialog()
    End Sub

    Private Sub btnLogout_Click(sender As Object, e As EventArgs) Handles btnLogout.Click
        If MessageBox.Show("Log out of NexaBank?", "Confirm Logout",
                           MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            WriteAudit("LOGOUT", $"Employee {CurrentEmployee?.FullName} logged out.")
            CurrentEmployee = Nothing
            IsAdminSession = False
            refreshTimer.Stop()
            Dim login As New frmLogin()
            login.Show()
            Me.Close()
        End If
    End Sub

    Protected Overrides Sub OnFormClosing(e As FormClosingEventArgs)
        refreshTimer.Stop()
        If CurrentEmployee IsNot Nothing Then
            WriteAudit("LOGOUT", $"Dashboard closed by {CurrentEmployee.FullName}.")
        End If
        MyBase.OnFormClosing(e)
    End Sub

End Class
