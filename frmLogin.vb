' ============================================================
'  BANK MANAGEMENT SYSTEM - frmLogin.vb
'  Secure employee login with PIN masking and lockout
' ============================================================
Imports System.Drawing
Imports System.Windows.Forms

Public Class frmLogin
    Inherits Form

    ' ── Controls ──────────────────────────────────────────────
    Private WithEvents btnLogin As Button
    Private WithEvents btnExit As Button
    Private WithEvents txtEmployeeID As TextBox
    Private WithEvents txtPin As TextBox
    Private WithEvents lblTitle As Label
    Private WithEvents lblSubtitle As Label
    Private WithEvents lblEmpID As Label
    Private WithEvents lblPin As Label
    Private WithEvents lblAttempts As Label
    Private WithEvents pnlHeader As Panel
    Private WithEvents pnlForm As Panel
    Private WithEvents chkShowPin As CheckBox
    Private WithEvents picLogo As PictureBox

    ' ── State ─────────────────────────────────────────────────
    Private failedAttempts As Integer = 0
    Private Const MAX_ATTEMPTS As Integer = 3
    Private lockoutTimer As System.Windows.Forms.Timer

    Public Sub New()
        InitialiseStorage()
        BuildUI()
    End Sub

    Private Sub BuildUI()
        Me.Text = APP_TITLE & " - Login"
        Me.Size = New Size(480, 580)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.FormBorderStyle = FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.BackColor = Color.FromArgb(245, 247, 250)
        Me.Font = New Font("Segoe UI", 9.5F)

        ' ── Header Panel ──────────────────────────────────────
        pnlHeader = New Panel()
        pnlHeader.Size = New Size(480, 160)
        pnlHeader.Location = New Point(0, 0)
        pnlHeader.BackColor = Color.FromArgb(15, 76, 129)

        lblTitle = New Label()
        lblTitle.Text = "🏦 " & BANK_NAME
        lblTitle.Font = New Font("Segoe UI", 22, FontStyle.Bold)
        lblTitle.ForeColor = Color.White
        lblTitle.AutoSize = True
        lblTitle.Location = New Point(40, 40)

        lblSubtitle = New Label()
        lblSubtitle.Text = "Bank Management System — Staff Portal"
        lblSubtitle.Font = New Font("Segoe UI", 10)
        lblSubtitle.ForeColor = Color.FromArgb(180, 210, 240)
        lblSubtitle.AutoSize = True
        lblSubtitle.Location = New Point(40, 100)

        pnlHeader.Controls.Add(lblTitle)
        pnlHeader.Controls.Add(lblSubtitle)

        ' ── Form Panel ────────────────────────────────────────
        pnlForm = New Panel()
        pnlForm.Size = New Size(380, 320)
        pnlForm.Location = New Point(50, 190)
        pnlForm.BackColor = Color.White
        pnlForm.Padding = New Padding(30)

        Dim shadow As Panel = New Panel()
        shadow.Size = New Size(382, 322)
        shadow.Location = New Point(49, 191)
        shadow.BackColor = Color.FromArgb(200, 200, 210)

        ' Employee ID field
        lblEmpID = New Label()
        lblEmpID.Text = "Employee ID"
        lblEmpID.Font = New Font("Segoe UI", 9, FontStyle.Bold)
        lblEmpID.ForeColor = Color.FromArgb(60, 60, 80)
        lblEmpID.Location = New Point(30, 30)
        lblEmpID.AutoSize = True

        txtEmployeeID = New TextBox()
        txtEmployeeID.Size = New Size(320, 32)
        txtEmployeeID.Location = New Point(30, 52)
        txtEmployeeID.Font = New Font("Segoe UI", 11)
        txtEmployeeID.BorderStyle = BorderStyle.FixedSingle
        txtEmployeeID.PlaceholderText = "e.g. EMP-001"
        txtEmployeeID.CharacterCasing = CharacterCasing.Upper

        ' PIN field
        lblPin = New Label()
        lblPin.Text = "PIN (4-6 digits)"
        lblPin.Font = New Font("Segoe UI", 9, FontStyle.Bold)
        lblPin.ForeColor = Color.FromArgb(60, 60, 80)
        lblPin.Location = New Point(30, 100)
        lblPin.AutoSize = True

        txtPin = New TextBox()
        txtPin.Size = New Size(320, 32)
        txtPin.Location = New Point(30, 122)
        txtPin.Font = New Font("Segoe UI", 11)
        txtPin.BorderStyle = BorderStyle.FixedSingle
        txtPin.PasswordChar = "●"c
        txtPin.MaxLength = 6

        chkShowPin = New CheckBox()
        chkShowPin.Text = "Show PIN"
        chkShowPin.Location = New Point(30, 162)
        chkShowPin.AutoSize = True
        chkShowPin.ForeColor = Color.FromArgb(100, 100, 120)

        ' Attempts label
        lblAttempts = New Label()
        lblAttempts.Text = ""
        lblAttempts.Font = New Font("Segoe UI", 8.5F)
        lblAttempts.ForeColor = Color.FromArgb(200, 50, 50)
        lblAttempts.Location = New Point(30, 188)
        lblAttempts.Size = New Size(320, 20)

        ' Login button
        btnLogin = New Button()
        btnLogin.Text = "LOGIN  →"
        btnLogin.Size = New Size(320, 42)
        btnLogin.Location = New Point(30, 218)
        btnLogin.BackColor = Color.FromArgb(15, 76, 129)
        btnLogin.ForeColor = Color.White
        btnLogin.Font = New Font("Segoe UI", 11, FontStyle.Bold)
        btnLogin.FlatStyle = FlatStyle.Flat
        btnLogin.FlatAppearance.BorderSize = 0
        btnLogin.Cursor = Cursors.Hand

        pnlForm.Controls.AddRange(New Control() {
            lblEmpID, txtEmployeeID,
            lblPin, txtPin,
            chkShowPin, lblAttempts, btnLogin})

        ' Exit button
        btnExit = New Button()
        btnExit.Text = "Exit System"
        btnExit.Size = New Size(120, 32)
        btnExit.Location = New Point(180, 530)
        btnExit.BackColor = Color.FromArgb(200, 50, 50)
        btnExit.ForeColor = Color.White
        btnExit.FlatStyle = FlatStyle.Flat
        btnExit.FlatAppearance.BorderSize = 0
        btnExit.Cursor = Cursors.Hand

        Me.Controls.Add(pnlHeader)
        Me.Controls.Add(shadow)
        Me.Controls.Add(pnlForm)
        Me.Controls.Add(btnExit)

        Me.AcceptButton = btnLogin
    End Sub

    ' ── Event Handlers ────────────────────────────────────────

    Private Sub chkShowPin_CheckedChanged(sender As Object, e As EventArgs) Handles chkShowPin.CheckedChanged
        txtPin.PasswordChar = If(chkShowPin.Checked, Nothing, "●"c)
    End Sub

    Private Sub btnLogin_Click(sender As Object, e As EventArgs) Handles btnLogin.Click
        Dim empID As String = txtEmployeeID.Text.Trim()
        Dim pin As String = txtPin.Text.Trim()

        If String.IsNullOrEmpty(empID) OrElse String.IsNullOrEmpty(pin) Then
            ShowError("Please enter both Employee ID and PIN.")
            Return
        End If

        Dim emp = DataAccess.AuthenticateEmployee(empID, pin)

        If emp IsNot Nothing Then
            CurrentEmployee = emp
            IsAdminSession = (emp.Role = "Admin" OrElse emp.Role = "Manager")
            WriteAudit("LOGIN", $"Employee {emp.FullName} logged in successfully.")

            Dim dashboard As New frmDashboard()
            dashboard.Show()
            Me.Hide()
        Else
            failedAttempts += 1
            Dim remaining = MAX_ATTEMPTS - failedAttempts
            WriteAudit("FAILED_LOGIN", $"Failed attempt for ID={empID}. Attempts={failedAttempts}")

            If failedAttempts >= MAX_ATTEMPTS Then
                ShowError("Too many failed attempts. System locked for 30 seconds.")
                LockForm()
            Else
                ShowError($"Invalid Employee ID or PIN. {remaining} attempt(s) remaining.")
            End If
            txtPin.Clear()
            txtPin.Focus()
        End If
    End Sub

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        If MessageBox.Show("Exit NexaBank?", "Confirm", MessageBoxButtons.YesNo,
                           MessageBoxIcon.Question) = DialogResult.Yes Then
            Application.Exit()
        End If
    End Sub

    Private Sub ShowError(msg As String)
        lblAttempts.Text = "⚠ " & msg
    End Sub

    Private Sub LockForm()
        btnLogin.Enabled = False
        txtEmployeeID.Enabled = False
        txtPin.Enabled = False

        Dim countdown As Integer = 30
        lockoutTimer = New System.Windows.Forms.Timer()
        lockoutTimer.Interval = 1000
        AddHandler lockoutTimer.Tick, Sub()
            countdown -= 1
            ShowError($"Locked. Please wait {countdown}s...")
            If countdown <= 0 Then
                lockoutTimer.Stop()
                failedAttempts = 0
                btnLogin.Enabled = True
                txtEmployeeID.Enabled = True
                txtPin.Enabled = True
                lblAttempts.Text = ""
                txtEmployeeID.Clear()
                txtPin.Clear()
                txtEmployeeID.Focus()
            End If
        End Sub
        lockoutTimer.Start()
    End Sub

    Protected Overrides Sub OnFormClosing(e As FormClosingEventArgs)
        If e.CloseReason = CloseReason.UserClosing AndAlso Not IsDisposed Then
            If Not Application.OpenForms.Cast(Of Form)().Any(Function(f) TypeOf f Is frmDashboard) Then
                Application.Exit()
            End If
        End If
        MyBase.OnFormClosing(e)
    End Sub

End Class
