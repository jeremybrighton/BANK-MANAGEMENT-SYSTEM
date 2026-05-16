' ============================================================
'  BANK MANAGEMENT SYSTEM - frmReports.vb
'  Comprehensive reports: summary, customers, transactions, audit
' ============================================================
Imports System.Drawing
Imports System.Windows.Forms
Imports System.IO
Imports System.Text

Public Class frmReports
    Inherits Form

    Private WithEvents tabReports As TabControl
    Private WithEvents btnClose As Button
    Private WithEvents btnExport As Button

    Public Sub New()
        BuildUI()
    End Sub

    Private Sub BuildUI()
        Me.Text = "Reports & Analytics"
        Me.Size = New Size(980, 660)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.BackColor = Color.FromArgb(245, 247, 250)
        Me.Font = New Font("Segoe UI", 9.5F)

        Dim pnlH As New Panel() With {.Size = New Size(980, 50), .BackColor = Color.FromArgb(15, 100, 130)}
        Dim lH As New Label() With {.Text = "📊  Reports & Analytics", .Font = New Font("Segoe UI", 13, FontStyle.Bold), .ForeColor = Color.White, .AutoSize = True, .Location = New Point(16, 13)}
        pnlH.Controls.Add(lH)

        tabReports = New TabControl()
        tabReports.Location = New Point(10, 60)
        tabReports.Size = New Size(960, 545)
        tabReports.Font = New Font("Segoe UI", 10)

        tabReports.TabPages.Add(BuildSummaryTab())
        tabReports.TabPages.Add(BuildCustomersTab())
        tabReports.TabPages.Add(BuildTransactionsTab())
        tabReports.TabPages.Add(BuildAuditTab())

        btnExport = New Button() With {.Text = "💾 Export to CSV", .Size = New Size(140, 32), .Location = New Point(720, 617), .BackColor = Color.FromArgb(22, 130, 90), .ForeColor = Color.White, .FlatStyle = FlatStyle.Flat, .Cursor = Cursors.Hand}
        btnExport.FlatAppearance.BorderSize = 0

        btnClose = New Button() With {.Text = "✕ Close", .Size = New Size(110, 32), .Location = New Point(868, 617), .BackColor = Color.FromArgb(140, 50, 50), .ForeColor = Color.White, .FlatStyle = FlatStyle.Flat, .Cursor = Cursors.Hand}
        btnClose.FlatAppearance.BorderSize = 0

        Me.Controls.AddRange(New Control() {pnlH, tabReports, btnExport, btnClose})
    End Sub

    ' ── Summary Report Tab ────────────────────────────────────
    Private Function BuildSummaryTab() As TabPage
        Dim tab As New TabPage("  📈 Summary  ")
        tab.BackColor = Color.White

        Dim rpt = DataAccess.GenerateSummaryReport()
        Dim txns = DataAccess.LoadAllTransactions()
        Dim todayTxns = txns.Where(Function(t) t.Timestamp.Date = DateTime.Today).ToList()

        Dim y As Integer = 20
        Dim AddRow = Sub(label As String, value As String, color As Color)
            Dim pnl As New Panel() With {.Size = New Size(900, 48), .Location = New Point(20, y), .BackColor = Color.FromArgb(248, 250, 255)}
            Dim accent As New Panel() With {.Size = New Size(5, 48), .BackColor = color, .Dock = DockStyle.Left}
            Dim lblL As New Label() With {.Text = label, .Font = New Font("Segoe UI", 10, FontStyle.Bold), .ForeColor = Color.FromArgb(60, 70, 90), .Location = New Point(14, 13), .AutoSize = True}
            Dim lblV As New Label() With {.Text = value, .Font = New Font("Segoe UI", 12, FontStyle.Bold), .ForeColor = color, .AutoSize = True}
            lblV.Location = New Point(480, 11)
            pnl.Controls.AddRange(New Control() {accent, lblL, lblV})
            tab.Controls.Add(pnl)
            y += 56
        End Sub

        Dim lblTitle As New Label() With {.Text = $"SUMMARY REPORT  —  Generated: {DateTime.Now:dd MMM yyyy HH:mm}", .Font = New Font("Segoe UI", 10, FontStyle.Bold), .ForeColor = Color.FromArgb(80, 90, 110), .Location = New Point(20, y), .AutoSize = True}
        tab.Controls.Add(lblTitle)
        y += 36

        AddRow("Total Registered Customers", rpt.TotalCustomers.ToString("N0"), Color.FromArgb(15, 76, 129))
        AddRow("Active Accounts", rpt.ActiveAccounts.ToString("N0"), Color.FromArgb(22, 130, 90))
        AddRow("Frozen Accounts", rpt.FrozenAccounts.ToString("N0"), Color.FromArgb(180, 100, 20))
        AddRow("Closed Accounts", rpt.ClosedAccounts.ToString("N0"), Color.FromArgb(160, 50, 50))
        AddRow("Total Funds in Bank (KES)", FormatMoney(rpt.TotalBalance), Color.FromArgb(15, 100, 130))
        AddRow("Total All-time Deposits (KES)", FormatMoney(rpt.TotalDeposits), Color.FromArgb(22, 130, 90))
        AddRow("Total All-time Withdrawals (KES)", FormatMoney(rpt.TotalWithdrawals), Color.FromArgb(160, 50, 50))
        AddRow("Net Bank Position (KES)", FormatMoney(rpt.TotalDeposits - rpt.TotalWithdrawals), Color.FromArgb(70, 40, 150))
        AddRow("Total Transaction Count", rpt.TransactionCount.ToString("N0"), Color.FromArgb(15, 76, 129))
        AddRow("Today's Transactions", todayTxns.Count.ToString("N0"), Color.FromArgb(100, 70, 160))

        Return tab
    End Function

    ' ── All Customers Tab ─────────────────────────────────────
    Private Function BuildCustomersTab() As TabPage
        Dim tab As New TabPage("  👥 All Customers  ")
        tab.BackColor = Color.White

        Dim dgv = BuildDgv(New String() {"Account No.", "Name", "Phone", "Email", "ID No.", "Type", "Balance", "Status", "Opened"})
        dgv.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(15, 76, 129)

        Dim customers = DataAccess.LoadAllCustomers()
        customers = customers.OrderBy(Function(c) c.FullName).ToList()

        For Each c In customers
            Dim i = dgv.Rows.Add()
            dgv.Rows(i).Cells(0).Value = c.AccountNumber
            dgv.Rows(i).Cells(1).Value = c.FullName
            dgv.Rows(i).Cells(2).Value = c.PhoneNumber
            dgv.Rows(i).Cells(3).Value = c.Email
            dgv.Rows(i).Cells(4).Value = c.IDNumber
            dgv.Rows(i).Cells(5).Value = c.AccountType
            dgv.Rows(i).Cells(6).Value = FormatMoney(c.Balance)
            dgv.Rows(i).Cells(7).Value = c.Status
            dgv.Rows(i).Cells(8).Value = c.DateOpened.ToString("dd MMM yyyy")
            If c.Status = "Active" Then
                dgv.Rows(i).Cells(7).Style.ForeColor = Color.FromArgb(0, 140, 60)
                dgv.Rows(i).Cells(7).Style.Font = New Font("Segoe UI", 9, FontStyle.Bold)
            End If
        Next

        Dim lblTotal As New Label() With {.Text = $"Total: {customers.Count} customer(s)  |  Sum Balance: {FormatMoney(customers.Sum(Function(c) c.Balance))}", .Location = New Point(5, 3), .AutoSize = True, .Font = New Font("Segoe UI", 8.5F, FontStyle.Bold)}
        tab.Controls.AddRange(New Control() {lblTotal, dgv})
        Return tab
    End Function

    ' ── All Transactions Tab ──────────────────────────────────
    Private Function BuildTransactionsTab() As TabPage
        Dim tab As New TabPage("  💰 Transactions  ")
        tab.BackColor = Color.White

        Dim dgv = BuildDgv(New String() {"Txn ID", "Account", "Type", "Amount", "Bal Before", "Bal After", "Date & Time", "By", "Notes"})
        dgv.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(22, 130, 90)

        Dim txns = DataAccess.LoadAllTransactions()
        txns = txns.OrderByDescending(Function(t) t.Timestamp).ToList()

        For Each t In txns
            Dim i = dgv.Rows.Add()
            dgv.Rows(i).Cells(0).Value = t.TransactionID
            dgv.Rows(i).Cells(1).Value = t.AccountNumber
            dgv.Rows(i).Cells(2).Value = t.TransactionType
            dgv.Rows(i).Cells(3).Value = FormatMoney(t.Amount)
            dgv.Rows(i).Cells(4).Value = FormatMoney(t.BalanceBefore)
            dgv.Rows(i).Cells(5).Value = FormatMoney(t.BalanceAfter)
            dgv.Rows(i).Cells(6).Value = t.Timestamp.ToString("dd MMM yyyy HH:mm:ss")
            dgv.Rows(i).Cells(7).Value = t.PerformedBy
            dgv.Rows(i).Cells(8).Value = t.Notes

            Dim isCredit = t.TransactionType = "Deposit" OrElse t.TransactionType = "Transfer-In"
            dgv.Rows(i).Cells(3).Style.ForeColor = If(isCredit, Color.FromArgb(0, 140, 60), Color.FromArgb(180, 40, 40))
            dgv.Rows(i).Cells(3).Style.Font = New Font("Segoe UI", 9, FontStyle.Bold)
        Next

        Dim totalDep = txns.Where(Function(t) t.TransactionType = "Deposit").Sum(Function(t) t.Amount)
        Dim totalWith = txns.Where(Function(t) t.TransactionType = "Withdrawal").Sum(Function(t) t.Amount)
        Dim lblTotal As New Label() With {
            .Text = $"Total Transactions: {txns.Count}  |  Deposits: {FormatMoney(totalDep)}  |  Withdrawals: {FormatMoney(totalWith)}",
            .Location = New Point(5, 3), .AutoSize = True, .Font = New Font("Segoe UI", 8.5F, FontStyle.Bold)}
        tab.Controls.AddRange(New Control() {lblTotal, dgv})
        Return tab
    End Function

    ' ── Audit Log Tab ─────────────────────────────────────────
    Private Function BuildAuditTab() As TabPage
        Dim tab As New TabPage("  🔒 Audit Log  ")
        tab.BackColor = Color.White

        Dim rtb As New RichTextBox()
        rtb.Location = New Point(5, 24)
        rtb.Size = New Size(950, 475)
        rtb.Font = New Font("Consolas", 9)
        rtb.ReadOnly = True
        rtb.BackColor = Color.FromArgb(20, 20, 30)
        rtb.ForeColor = Color.FromArgb(180, 220, 140)
        rtb.BorderStyle = BorderStyle.None

        Try
            If File.Exists(AuditPath) Then
                Dim lines = File.ReadAllLines(AuditPath)
                Dim recent = lines.Reverse().Take(500).Reverse().ToArray()
                rtb.Text = String.Join(vbCrLf, recent)
                rtb.SelectionStart = rtb.Text.Length
                rtb.ScrollToCaret()
            Else
                rtb.Text = "No audit log found."
            End If
        Catch ex As Exception
            rtb.Text = "Error reading audit log: " & ex.Message
        End Try

        Dim lblHint As New Label() With {.Text = "Showing last 500 audit entries (newest at bottom):", .Location = New Point(5, 5), .AutoSize = True, .Font = New Font("Segoe UI", 8.5F, FontStyle.Bold)}
        tab.Controls.AddRange(New Control() {lblHint, rtb})
        Return tab
    End Function

    ' ── Shared grid builder ───────────────────────────────────
    Private Function BuildDgv(columns() As String) As DataGridView
        Dim dgv As New DataGridView() With {
            .Location = New Point(5, 25),
            .Size = New Size(950, 478),
            .BackgroundColor = Color.White,
            .BorderStyle = BorderStyle.None,
            .RowHeadersVisible = False,
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            .AllowUserToAddRows = False,
            .AllowUserToDeleteRows = False,
            .ReadOnly = True,
            .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
            .ColumnHeadersHeight = 36,
            .ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing}

        dgv.EnableHeadersVisualStyles = False
        dgv.ColumnHeadersDefaultCellStyle.ForeColor = Color.White
        dgv.ColumnHeadersDefaultCellStyle.Font = New Font("Segoe UI", 9.5F, FontStyle.Bold)
        dgv.DefaultCellStyle.SelectionBackColor = Color.FromArgb(210, 225, 245)
        dgv.DefaultCellStyle.SelectionForeColor = Color.Black
        dgv.RowTemplate.Height = 28

        For Each col As String In columns
            Dim dc As New DataGridViewTextBoxColumn()
            dc.HeaderText = col : dc.ReadOnly = True
            dgv.Columns.Add(dc)
        Next
        Return dgv
    End Function

    ' ── Export to CSV ─────────────────────────────────────────
    Private Sub btnExport_Click(sender As Object, e As EventArgs) Handles btnExport.Click
        Using sfd As New SaveFileDialog()
            sfd.Filter = "CSV files (*.csv)|*.csv"
            sfd.FileName = $"NexaBank_Report_{DateTime.Now:yyyyMMdd_HHmmss}.csv"
            If sfd.ShowDialog() = DialogResult.OK Then
                Try
                    Dim sb As New StringBuilder()
                    sb.AppendLine("NexaBank Report — Generated: " & DateTime.Now.ToString("dd MMM yyyy HH:mm:ss"))
                    sb.AppendLine()

                    ' Summary
                    Dim rpt = DataAccess.GenerateSummaryReport()
                    sb.AppendLine("=== SUMMARY ===")
                    sb.AppendLine($"Total Customers,{rpt.TotalCustomers}")
                    sb.AppendLine($"Active Accounts,{rpt.ActiveAccounts}")
                    sb.AppendLine($"Frozen Accounts,{rpt.FrozenAccounts}")
                    sb.AppendLine($"Closed Accounts,{rpt.ClosedAccounts}")
                    sb.AppendLine($"Total Balance,{rpt.TotalBalance:F2}")
                    sb.AppendLine($"Total Deposits,{rpt.TotalDeposits:F2}")
                    sb.AppendLine($"Total Withdrawals,{rpt.TotalWithdrawals:F2}")
                    sb.AppendLine($"Transactions,{rpt.TransactionCount}")
                    sb.AppendLine()

                    ' Customers
                    sb.AppendLine("=== CUSTOMERS ===")
                    sb.AppendLine("AccountNumber,Name,Phone,Email,IDNumber,Type,Balance,Status,DateOpened")
                    For Each c In DataAccess.LoadAllCustomers()
                        sb.AppendLine(c.ToCsvLine())
                    Next
                    sb.AppendLine()

                    ' Transactions
                    sb.AppendLine("=== TRANSACTIONS ===")
                    sb.AppendLine("TxnID,Account,Type,Amount,BalBefore,BalAfter,Timestamp,By,Notes")
                    For Each t In DataAccess.LoadAllTransactions()
                        sb.AppendLine(t.ToCsvLine())
                    Next

                    File.WriteAllText(sfd.FileName, sb.ToString())
                    MsgBox($"Report exported to:{vbCrLf}{sfd.FileName}", MsgBoxStyle.Information, "Export Successful")
                Catch ex As Exception
                    MsgBox("Export failed: " & ex.Message, MsgBoxStyle.Critical)
                End Try
            End If
        End Using
    End Sub

    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub
End Class
