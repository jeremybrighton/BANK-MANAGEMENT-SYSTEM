' ============================================================
'  BANK MANAGEMENT SYSTEM - frmTransactions.vb
'  Deposit, Withdrawal, and Transfer operations
' ============================================================
Imports System.Drawing
Imports System.Windows.Forms

Public Class frmTransactions
    Inherits Form

    Private WithEvents tabControl As TabControl
    Private WithEvents btnClose As Button
    Private pnlHeader As Panel

    Public Sub New()
        BuildUI()
    End Sub

    Private Sub BuildUI()
        Me.Text = "Banking Transactions"
        Me.Size = New Size(560, 580)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.BackColor = Color.FromArgb(245, 247, 250)
        Me.Font = New Font("Segoe UI", 9.5F)
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False

        pnlHeader = New Panel()
        pnlHeader.Size = New Size(560, 50)
        pnlHeader.BackColor = Color.FromArgb(22, 130, 90)
        Dim lblH As New Label()
        lblH.Text = "💰  Banking Transactions"
        lblH.Font = New Font("Segoe UI", 13, FontStyle.Bold)
        lblH.ForeColor = Color.White
        lblH.AutoSize = True
        lblH.Location = New Point(16, 13)
        pnlHeader.Controls.Add(lblH)

        tabControl = New TabControl()
        tabControl.Location = New Point(10, 60)
        tabControl.Size = New Size(530, 450)
        tabControl.Font = New Font("Segoe UI", 10)

        ' Build tabs
        tabControl.TabPages.Add(BuildDepositTab())
        tabControl.TabPages.Add(BuildWithdrawTab())
        tabControl.TabPages.Add(BuildTransferTab())

        btnClose = New Button()
        btnClose.Text = "✕  Close"
        btnClose.Size = New Size(110, 34)
        btnClose.Location = New Point(430, 520)
        btnClose.BackColor = Color.FromArgb(160, 50, 50)
        btnClose.ForeColor = Color.White
        btnClose.FlatStyle = FlatStyle.Flat
        btnClose.FlatAppearance.BorderSize = 0
        btnClose.Cursor = Cursors.Hand

        Me.Controls.AddRange(New Control() {pnlHeader, tabControl, btnClose})
    End Sub

    ' ── Deposit Tab ───────────────────────────────────────────
    Private Function BuildDepositTab() As TabPage
        Dim tab As New TabPage("  ⬇ Deposit  ")
        tab.BackColor = Color.FromArgb(245, 250, 246)

        Dim lblAcct As New Label() With {.Text = "Account Number", .Location = New Point(20, 24), .AutoSize = True, .Font = New Font("Segoe UI", 9, FontStyle.Bold)}
        Dim txtAcct As New TextBox() With {.Size = New Size(300, 28), .Location = New Point(20, 46), .Font = New Font("Segoe UI", 11)}

        Dim lblAmt As New Label() With {.Text = "Amount to Deposit (KES)", .Location = New Point(20, 90), .AutoSize = True, .Font = New Font("Segoe UI", 9, FontStyle.Bold)}
        Dim txtAmt As New TextBox() With {.Size = New Size(200, 28), .Location = New Point(20, 112), .Font = New Font("Segoe UI", 11)}

        Dim lblNotes As New Label() With {.Text = "Notes / Reference", .Location = New Point(20, 156), .AutoSize = True, .Font = New Font("Segoe UI", 9, FontStyle.Bold)}
        Dim txtNotes As New TextBox() With {.Size = New Size(480, 28), .Location = New Point(20, 178), .Font = New Font("Segoe UI", 10)}

        Dim pnlInfo As New Panel() With {.Size = New Size(480, 80), .Location = New Point(20, 220), .BackColor = Color.FromArgb(230, 245, 235), .BorderStyle = BorderStyle.FixedSingle}
        Dim lblInfo As New Label() With {.Text = "Account info will appear here after lookup.", .Font = New Font("Segoe UI", 9), .ForeColor = Color.FromArgb(40, 100, 60), .Location = New Point(8, 8), .Size = New Size(462, 62)}
        pnlInfo.Controls.Add(lblInfo)

        Dim btnLookup As New Button() With {
            .Text = "🔍 Lookup", .Size = New Size(100, 28), .Location = New Point(330, 46),
            .BackColor = Color.FromArgb(15, 76, 129), .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat, .Cursor = Cursors.Hand}
        btnLookup.FlatAppearance.BorderSize = 0

        Dim btnDeposit As New Button() With {
            .Text = "✔  DEPOSIT", .Size = New Size(200, 42), .Location = New Point(20, 320),
            .BackColor = Color.FromArgb(22, 130, 90), .ForeColor = Color.White,
            .Font = New Font("Segoe UI", 11, FontStyle.Bold),
            .FlatStyle = FlatStyle.Flat, .Cursor = Cursors.Hand}
        btnDeposit.FlatAppearance.BorderSize = 0

        Dim lblResult As New Label() With {.Text = "", .Location = New Point(20, 374), .Size = New Size(480, 28), .Font = New Font("Segoe UI", 9.5F, FontStyle.Bold)}

        AddHandler btnLookup.Click, Sub()
            Dim c = DataAccess.FindCustomerByAccount(txtAcct.Text.Trim())
            If c Is Nothing Then
                lblInfo.Text = "⚠ Account not found."
                lblInfo.ForeColor = Color.FromArgb(180, 40, 40)
            Else
                lblInfo.Text = $"✔  Name: {c.FullName}     Type: {c.AccountType}     Status: {c.Status}{vbCrLf}" &
                               $"     Current Balance: {FormatMoney(c.Balance)}"
                lblInfo.ForeColor = Color.FromArgb(20, 100, 50)
            End If
        End Sub

        AddHandler btnDeposit.Click, Sub()
            Dim accNum = txtAcct.Text.Trim()
            Dim amtStr = txtAmt.Text.Trim()
            Dim amt As Decimal

            If String.IsNullOrEmpty(accNum) Then
                ShowMsg(lblResult, "⚠ Please enter an account number.", False)
                Return
            End If
            If Not Decimal.TryParse(amtStr, amt) OrElse amt <= 0 Then
                ShowMsg(lblResult, "⚠ Enter a valid amount greater than 0.", False)
                Return
            End If

            Dim result = DataAccess.Deposit(accNum, amt, txtNotes.Text.Trim())
            ShowMsg(lblResult, If(result.Success, "✔ " & result.Message, "✘ " & result.Message), result.Success)
            If result.Success Then
                txtAmt.Clear()
                txtNotes.Clear()
                ' Refresh info
                Dim c = DataAccess.FindCustomerByAccount(accNum)
                If c IsNot Nothing Then
                    lblInfo.Text = $"✔  Name: {c.FullName}     New Balance: {FormatMoney(c.Balance)}"
                    lblInfo.ForeColor = Color.FromArgb(20, 100, 50)
                End If
            End If
        End Sub

        tab.Controls.AddRange(New Control() {
            lblAcct, txtAcct, btnLookup,
            lblAmt, txtAmt, lblNotes, txtNotes,
            pnlInfo, btnDeposit, lblResult})
        Return tab
    End Function

    ' ── Withdraw Tab ──────────────────────────────────────────
    Private Function BuildWithdrawTab() As TabPage
        Dim tab As New TabPage("  ⬆ Withdraw  ")
        tab.BackColor = Color.FromArgb(250, 246, 245)

        Dim lblAcct As New Label() With {.Text = "Account Number", .Location = New Point(20, 24), .AutoSize = True, .Font = New Font("Segoe UI", 9, FontStyle.Bold)}
        Dim txtAcct As New TextBox() With {.Size = New Size(300, 28), .Location = New Point(20, 46), .Font = New Font("Segoe UI", 11)}

        Dim pnlInfo As New Panel() With {.Size = New Size(480, 80), .Location = New Point(20, 84), .BackColor = Color.FromArgb(245, 235, 232), .BorderStyle = BorderStyle.FixedSingle}
        Dim lblInfo As New Label() With {.Text = "Account info will appear here after lookup.", .Font = New Font("Segoe UI", 9), .ForeColor = Color.FromArgb(100, 60, 40), .Location = New Point(8, 8), .Size = New Size(462, 62)}
        pnlInfo.Controls.Add(lblInfo)

        Dim btnLookup As New Button() With {
            .Text = "🔍 Lookup", .Size = New Size(100, 28), .Location = New Point(330, 46),
            .BackColor = Color.FromArgb(15, 76, 129), .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat, .Cursor = Cursors.Hand}
        btnLookup.FlatAppearance.BorderSize = 0

        Dim lblAmt As New Label() With {.Text = "Amount to Withdraw (KES)", .Location = New Point(20, 180), .AutoSize = True, .Font = New Font("Segoe UI", 9, FontStyle.Bold)}
        Dim txtAmt As New TextBox() With {.Size = New Size(200, 28), .Location = New Point(20, 202), .Font = New Font("Segoe UI", 11)}

        Dim lblNotes As New Label() With {.Text = "Notes / Reference", .Location = New Point(20, 246), .AutoSize = True, .Font = New Font("Segoe UI", 9, FontStyle.Bold)}
        Dim txtNotes As New TextBox() With {.Size = New Size(480, 28), .Location = New Point(20, 268), .Font = New Font("Segoe UI", 10)}

        Dim btnWithdraw As New Button() With {
            .Text = "✔  WITHDRAW", .Size = New Size(200, 42), .Location = New Point(20, 318),
            .BackColor = Color.FromArgb(180, 80, 20), .ForeColor = Color.White,
            .Font = New Font("Segoe UI", 11, FontStyle.Bold),
            .FlatStyle = FlatStyle.Flat, .Cursor = Cursors.Hand}
        btnWithdraw.FlatAppearance.BorderSize = 0

        Dim lblResult As New Label() With {.Text = "", .Location = New Point(20, 374), .Size = New Size(480, 28), .Font = New Font("Segoe UI", 9.5F, FontStyle.Bold)}

        AddHandler btnLookup.Click, Sub()
            Dim c = DataAccess.FindCustomerByAccount(txtAcct.Text.Trim())
            If c Is Nothing Then
                lblInfo.Text = "⚠ Account not found."
                lblInfo.ForeColor = Color.FromArgb(180, 40, 40)
            Else
                lblInfo.Text = $"✔  Name: {c.FullName}     Type: {c.AccountType}     Status: {c.Status}{vbCrLf}" &
                               $"     Available Balance: {FormatMoney(c.Balance)}"
                lblInfo.ForeColor = Color.FromArgb(100, 60, 20)
            End If
        End Sub

        AddHandler btnWithdraw.Click, Sub()
            Dim accNum = txtAcct.Text.Trim()
            Dim amt As Decimal
            If String.IsNullOrEmpty(accNum) Then
                ShowMsg(lblResult, "⚠ Please enter an account number.", False)
                Return
            End If
            If Not Decimal.TryParse(txtAmt.Text.Trim(), amt) OrElse amt <= 0 Then
                ShowMsg(lblResult, "⚠ Enter a valid amount greater than 0.", False)
                Return
            End If

            Dim confirm = MessageBox.Show($"Withdraw {FormatMoney(amt)} from account {accNum}?",
                                          "Confirm Withdrawal", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If confirm = DialogResult.No Then Return

            Dim result = DataAccess.Withdraw(accNum, amt, txtNotes.Text.Trim())
            ShowMsg(lblResult, If(result.Success, "✔ " & result.Message, "✘ " & result.Message), result.Success)
            If result.Success Then
                txtAmt.Clear()
                txtNotes.Clear()
                Dim c = DataAccess.FindCustomerByAccount(accNum)
                If c IsNot Nothing Then
                    lblInfo.Text = $"✔  Name: {c.FullName}     New Balance: {FormatMoney(c.Balance)}"
                End If
            End If
        End Sub

        tab.Controls.AddRange(New Control() {
            lblAcct, txtAcct, btnLookup, pnlInfo,
            lblAmt, txtAmt, lblNotes, txtNotes,
            btnWithdraw, lblResult})
        Return tab
    End Function

    ' ── Transfer Tab ──────────────────────────────────────────
    Private Function BuildTransferTab() As TabPage
        Dim tab As New TabPage("  ↔ Transfer  ")
        tab.BackColor = Color.FromArgb(246, 245, 252)

        Dim lblFrom As New Label() With {.Text = "Source Account Number", .Location = New Point(20, 24), .AutoSize = True, .Font = New Font("Segoe UI", 9, FontStyle.Bold)}
        Dim txtFrom As New TextBox() With {.Size = New Size(280, 28), .Location = New Point(20, 46), .Font = New Font("Segoe UI", 11)}

        Dim lblTo As New Label() With {.Text = "Destination Account Number", .Location = New Point(20, 90), .AutoSize = True, .Font = New Font("Segoe UI", 9, FontStyle.Bold)}
        Dim txtTo As New TextBox() With {.Size = New Size(280, 28), .Location = New Point(20, 112), .Font = New Font("Segoe UI", 11)}

        Dim lblAmt As New Label() With {.Text = "Amount to Transfer (KES)", .Location = New Point(20, 156), .AutoSize = True, .Font = New Font("Segoe UI", 9, FontStyle.Bold)}
        Dim txtAmt As New TextBox() With {.Size = New Size(200, 28), .Location = New Point(20, 178), .Font = New Font("Segoe UI", 11)}

        Dim lblNotes As New Label() With {.Text = "Transfer Reference / Notes", .Location = New Point(20, 222), .AutoSize = True, .Font = New Font("Segoe UI", 9, FontStyle.Bold)}
        Dim txtNotes As New TextBox() With {.Size = New Size(480, 28), .Location = New Point(20, 244), .Font = New Font("Segoe UI", 10)}

        Dim btnTransfer As New Button() With {
            .Text = "↔  TRANSFER FUNDS", .Size = New Size(220, 42), .Location = New Point(20, 294),
            .BackColor = Color.FromArgb(70, 40, 150), .ForeColor = Color.White,
            .Font = New Font("Segoe UI", 11, FontStyle.Bold),
            .FlatStyle = FlatStyle.Flat, .Cursor = Cursors.Hand}
        btnTransfer.FlatAppearance.BorderSize = 0

        Dim lblResult As New Label() With {.Text = "", .Location = New Point(20, 354), .Size = New Size(480, 40), .Font = New Font("Segoe UI", 9.5F, FontStyle.Bold)}

        AddHandler btnTransfer.Click, Sub()
            Dim fromAcc = txtFrom.Text.Trim()
            Dim toAcc = txtTo.Text.Trim()
            Dim amt As Decimal

            If String.IsNullOrEmpty(fromAcc) OrElse String.IsNullOrEmpty(toAcc) Then
                ShowMsg(lblResult, "⚠ Enter both source and destination account numbers.", False)
                Return
            End If
            If Not Decimal.TryParse(txtAmt.Text.Trim(), amt) OrElse amt <= 0 Then
                ShowMsg(lblResult, "⚠ Enter a valid transfer amount.", False)
                Return
            End If

            Dim confirm = MessageBox.Show($"Transfer {FormatMoney(amt)} from {fromAcc} to {toAcc}?",
                                          "Confirm Transfer", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If confirm = DialogResult.No Then Return

            Dim result = DataAccess.Transfer(fromAcc, toAcc, amt, txtNotes.Text.Trim())
            ShowMsg(lblResult, If(result.Success, "✔ " & result.Message, "✘ " & result.Message), result.Success)
            If result.Success Then
                txtAmt.Clear()
                txtNotes.Clear()
            End If
        End Sub

        tab.Controls.AddRange(New Control() {
            lblFrom, txtFrom, lblTo, txtTo,
            lblAmt, txtAmt, lblNotes, txtNotes,
            btnTransfer, lblResult})
        Return tab
    End Function

    Private Sub ShowMsg(lbl As Label, msg As String, success As Boolean)
        lbl.Text = msg
        lbl.ForeColor = If(success, Color.FromArgb(0, 130, 60), Color.FromArgb(180, 40, 40))
    End Sub

    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

End Class
