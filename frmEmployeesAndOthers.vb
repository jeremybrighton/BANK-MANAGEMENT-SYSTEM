' ============================================================
'  BANK MANAGEMENT SYSTEM - frmEmployees.vb
'  Register and manage bank staff (Admin only)
' ============================================================
Imports System.Drawing
Imports System.Windows.Forms

Public Class frmEmployees
    Inherits Form

    Private WithEvents dgvEmployees As DataGridView
    Private WithEvents btnRegister As Button
    Private WithEvents btnEdit As Button
    Private WithEvents btnDeactivate As Button
    Private WithEvents btnResetPin As Button
    Private WithEvents btnRefresh As Button
    Private WithEvents btnClose As Button
    Private lblCount As Label

    Public Sub New()
        BuildUI()
        LoadEmployeesGrid()
    End Sub

    Private Sub BuildUI()
        Me.Text = "Employee Management"
        Me.Size = New Size(850, 540)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.BackColor = Color.FromArgb(245, 247, 250)
        Me.Font = New Font("Segoe UI", 9.5F)

        Dim pnlHead As New Panel() With {.Size = New Size(850, 50), .BackColor = Color.FromArgb(100, 40, 150)}
        Dim lblH As New Label() With {.Text = "🔑  Employee Management", .Font = New Font("Segoe UI", 13, FontStyle.Bold), .ForeColor = Color.White, .AutoSize = True, .Location = New Point(16, 13)}
        pnlHead.Controls.Add(lblH)

        dgvEmployees = New DataGridView()
        dgvEmployees.Location = New Point(10, 60)
        dgvEmployees.Size = New Size(830, 380)
        dgvEmployees.BackgroundColor = Color.White
        dgvEmployees.BorderStyle = BorderStyle.None
        dgvEmployees.RowHeadersVisible = False
        dgvEmployees.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        dgvEmployees.MultiSelect = False
        dgvEmployees.AllowUserToAddRows = False
        dgvEmployees.AllowUserToDeleteRows = False
        dgvEmployees.ReadOnly = True
        dgvEmployees.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        dgvEmployees.ColumnHeadersHeight = 36
        dgvEmployees.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        dgvEmployees.EnableHeadersVisualStyles = False
        dgvEmployees.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(100, 40, 150)
        dgvEmployees.ColumnHeadersDefaultCellStyle.ForeColor = Color.White
        dgvEmployees.ColumnHeadersDefaultCellStyle.Font = New Font("Segoe UI", 9.5F, FontStyle.Bold)
        dgvEmployees.DefaultCellStyle.SelectionBackColor = Color.FromArgb(225, 210, 245)
        dgvEmployees.DefaultCellStyle.SelectionForeColor = Color.Black
        dgvEmployees.RowTemplate.Height = 30

        Dim cols() As String = {"Employee ID", "Full Name", "Role", "Email", "Phone", "Registered", "Active"}
        For Each col As String In cols
            Dim dc As New DataGridViewTextBoxColumn()
            dc.HeaderText = col
            dc.ReadOnly = True
            dgvEmployees.Columns.Add(dc)
        Next

        Dim pnlBot As New Panel() With {.Size = New Size(850, 48), .Location = New Point(0, 450), .BackColor = Color.FromArgb(235, 240, 248)}
        lblCount = New Label() With {.Text = "0 employees", .Location = New Point(14, 14), .AutoSize = True, .Font = New Font("Segoe UI", 9)}

        btnRegister = MkBtn("➕ Register", Color.FromArgb(22, 130, 90), New Point(370, 9))
        btnEdit = MkBtn("✏ Edit", Color.FromArgb(15, 76, 129), New Point(480, 9))
        btnDeactivate = MkBtn("⛔ Deactivate", Color.FromArgb(180, 50, 50), New Point(570, 9))
        btnResetPin = MkBtn("🔐 Reset PIN", Color.FromArgb(120, 80, 30), New Point(680, 9))
        btnRefresh = MkBtn("↻", Color.FromArgb(80, 90, 110), New Point(760, 9))
        btnClose = MkBtn("✕ Close", Color.FromArgb(140, 50, 50), New Point(790, 9))

        pnlBot.Controls.AddRange(New Control() {lblCount, btnRegister, btnEdit, btnDeactivate, btnResetPin, btnRefresh, btnClose})
        Me.Controls.AddRange(New Control() {pnlHead, dgvEmployees, pnlBot})
    End Sub

    Private Function MkBtn(text As String, color As Color, loc As Point) As Button
        Dim b As New Button()
        b.Text = text : b.Size = New Size(100, 30) : b.Location = loc
        b.BackColor = color : b.ForeColor = Color.White
        b.FlatStyle = FlatStyle.Flat : b.FlatAppearance.BorderSize = 0
        b.Cursor = Cursors.Hand : b.Font = New Font("Segoe UI", 8.5F)
        Return b
    End Function

    Private Sub LoadEmployeesGrid()
        dgvEmployees.Rows.Clear()
        Dim all = DataAccess.LoadAllEmployees()
        For Each emp In all
            Dim idx = dgvEmployees.Rows.Add()
            Dim row = dgvEmployees.Rows(idx)
            row.Cells(0).Value = emp.EmployeeID
            row.Cells(1).Value = emp.FullName
            row.Cells(2).Value = emp.Role
            row.Cells(3).Value = emp.Email
            row.Cells(4).Value = emp.Phone
            row.Cells(5).Value = emp.DateRegistered.ToString("dd MMM yyyy")
            row.Cells(6).Value = If(emp.IsActive, "✔ Yes", "✘ No")
            row.Cells(6).Style.ForeColor = If(emp.IsActive, Color.FromArgb(0, 140, 60), Color.FromArgb(180, 40, 40))
            row.Cells(6).Style.Font = New Font("Segoe UI", 9, FontStyle.Bold)
            If Not emp.IsActive Then row.DefaultCellStyle.ForeColor = Color.Gray
        Next
        lblCount.Text = $"{all.Count} employee(s)"
    End Sub

    Private Function GetSelectedEmpID() As String
        If dgvEmployees.SelectedRows.Count = 0 Then Return Nothing
        Return dgvEmployees.SelectedRows(0).Cells(0).Value?.ToString()
    End Function

    Private Sub btnRegister_Click(sender As Object, e As EventArgs) Handles btnRegister.Click
        Dim f As New frmEmployeeDetail(Nothing)
        If f.ShowDialog() = DialogResult.OK Then LoadEmployeesGrid()
    End Sub

    Private Sub btnEdit_Click(sender As Object, e As EventArgs) Handles btnEdit.Click
        Dim id = GetSelectedEmpID()
        If id Is Nothing Then MsgBox("Select an employee first.", MsgBoxStyle.Information) : Return
        Dim emp = DataAccess.FindEmployeeByID(id)
        Dim f As New frmEmployeeDetail(emp)
        If f.ShowDialog() = DialogResult.OK Then LoadEmployeesGrid()
    End Sub

    Private Sub btnDeactivate_Click(sender As Object, e As EventArgs) Handles btnDeactivate.Click
        Dim id = GetSelectedEmpID()
        If id Is Nothing Then MsgBox("Select an employee first.", MsgBoxStyle.Information) : Return
        If id = CurrentEmployee?.EmployeeID Then MsgBox("Cannot deactivate your own account.", MsgBoxStyle.Warning) : Return

        Dim emp = DataAccess.FindEmployeeByID(id)
        If emp Is Nothing Then Return
        emp.IsActive = Not emp.IsActive
        DataAccess.SaveEmployee(emp)
        WriteAudit("TOGGLE_EMPLOYEE", $"EmployeeID={id}, IsActive={emp.IsActive}")
        LoadEmployeesGrid()
    End Sub

    Private Sub btnResetPin_Click(sender As Object, e As EventArgs) Handles btnResetPin.Click
        Dim id = GetSelectedEmpID()
        If id Is Nothing Then MsgBox("Select an employee first.", MsgBoxStyle.Information) : Return
        Dim newPin = InputBox("Enter new 4-6 digit PIN for employee " & id & ":", "Reset PIN")
        If String.IsNullOrEmpty(newPin) Then Return
        If newPin.Length < 4 OrElse newPin.Length > 6 OrElse Not newPin.All(Function(c) Char.IsDigit(c)) Then
            MsgBox("PIN must be 4-6 digits.", MsgBoxStyle.Warning) : Return
        End If
        Dim emp = DataAccess.FindEmployeeByID(id)
        emp.PinHash = HashPin(newPin)
        DataAccess.SaveEmployee(emp)
        WriteAudit("RESET_PIN", $"EmployeeID={id}")
        MsgBox("PIN reset successfully.", MsgBoxStyle.Information)
    End Sub

    Private Sub btnRefresh_Click(sender As Object, e As EventArgs) Handles btnRefresh.Click
        LoadEmployeesGrid()
    End Sub

    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub
End Class

' ── Employee Detail Form ───────────────────────────────────────
Public Class frmEmployeeDetail
    Inherits Form

    Private ReadOnly _existing As EmployeeRecord
    Private WithEvents btnSave As Button
    Private WithEvents btnCancel As Button
    Private txtName As TextBox, txtEmail As TextBox, txtPhone As TextBox
    Private txtPin As TextBox, txtPin2 As TextBox
    Private cmbRole As ComboBox
    Private lblEmpID As Label

    Public Sub New(existing As EmployeeRecord)
        _existing = existing
        BuildUI()
        If existing IsNot Nothing Then PopulateFields()
    End Sub

    Private Sub BuildUI()
        Dim isEdit = _existing IsNot Nothing
        Me.Text = If(isEdit, "Edit Employee", "Register New Employee")
        Me.Size = New Size(440, 500)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False : Me.MinimizeBox = False
        Me.BackColor = Color.FromArgb(245, 247, 250)
        Me.Font = New Font("Segoe UI", 9.5F)

        Dim pnlH As New Panel() With {.Size = New Size(440, 50), .BackColor = Color.FromArgb(100, 40, 150)}
        Dim lH As New Label() With {.Text = If(isEdit, "✏ Edit Employee", "🔑 Register Employee"), .Font = New Font("Segoe UI", 12, FontStyle.Bold), .ForeColor = Color.White, .AutoSize = True, .Location = New Point(14, 13)}
        pnlH.Controls.Add(lH)

        Dim y As Integer = 65 : Const G As Integer = 56
        Dim MkL = Function(t As String, pt As Point) As Label
                      Return New Label() With {.Text = t, .Location = pt, .AutoSize = True, .Font = New Font("Segoe UI", 8.5F, FontStyle.Bold), .ForeColor = Color.FromArgb(60, 60, 80)}
                  End Function
        Dim MkT = Function(pt As Point) As TextBox
                      Return New TextBox() With {.Location = pt, .Size = New Size(380, 28), .Font = New Font("Segoe UI", 10)}
                  End Function

        lblEmpID = New Label() With {.Text = "(Auto-assigned)", .Location = New Point(20, y + 20), .AutoSize = True, .Font = New Font("Segoe UI", 10, FontStyle.Bold), .ForeColor = Color.FromArgb(100, 40, 150)}
        y += G
        txtName = MkT(New Point(20, y + 20)) : y += G
        cmbRole = New ComboBox() With {.Location = New Point(20, y + 20), .Size = New Size(200, 28), .DropDownStyle = ComboBoxStyle.DropDownList}
        cmbRole.Items.AddRange(New Object() {"Teller", "Manager", "Admin"})
        cmbRole.SelectedIndex = 0
        y += G
        txtEmail = MkT(New Point(20, y + 20)) : y += G
        txtPhone = MkT(New Point(20, y + 20)) : y += G
        txtPin = MkT(New Point(20, y + 20))
        txtPin.PasswordChar = "●"c : txtPin.MaxLength = 6 : y += G
        txtPin2 = MkT(New Point(20, y + 20))
        txtPin2.PasswordChar = "●"c : txtPin2.MaxLength = 6 : y += G + 5

        Dim labels() As Control = {
            MkL("Employee ID (Auto)", New Point(20, 65)),
            MkL("Full Name *", New Point(20, 65 + G)),
            MkL("Role *", New Point(20, 65 + G * 2)),
            MkL("Email *", New Point(20, 65 + G * 3)),
            MkL("Phone *", New Point(20, 65 + G * 4)),
            MkL(If(isEdit, "New PIN (leave blank = no change)", "PIN (4-6 digits) *"), New Point(20, 65 + G * 5)),
            MkL("Confirm PIN *", New Point(20, 65 + G * 6))
        }

        btnSave = New Button() With {.Text = If(isEdit, "💾 Save", "✔ Register"), .Size = New Size(160, 38), .Location = New Point(20, y), .BackColor = Color.FromArgb(100, 40, 150), .ForeColor = Color.White, .FlatStyle = FlatStyle.Flat, .Font = New Font("Segoe UI", 10, FontStyle.Bold), .Cursor = Cursors.Hand}
        btnSave.FlatAppearance.BorderSize = 0
        btnCancel = New Button() With {.Text = "✕ Cancel", .Size = New Size(120, 38), .Location = New Point(200, y), .BackColor = Color.FromArgb(140, 50, 50), .ForeColor = Color.White, .FlatStyle = FlatStyle.Flat, .Cursor = Cursors.Hand}
        btnCancel.FlatAppearance.BorderSize = 0

        Me.Controls.Add(pnlH)
        Me.Controls.Add(lblEmpID)
        Me.Controls.AddRange(labels)
        Me.Controls.AddRange(New Control() {txtName, cmbRole, txtEmail, txtPhone, txtPin, txtPin2, btnSave, btnCancel})
        Me.AcceptButton = btnSave : Me.CancelButton = btnCancel
    End Sub

    Private Sub PopulateFields()
        lblEmpID.Text = _existing.EmployeeID
        txtName.Text = _existing.FullName
        cmbRole.SelectedItem = _existing.Role
        txtEmail.Text = _existing.Email
        txtPhone.Text = _existing.Phone
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        If String.IsNullOrWhiteSpace(txtName.Text) OrElse String.IsNullOrWhiteSpace(txtEmail.Text) OrElse String.IsNullOrWhiteSpace(txtPhone.Text) Then
            MsgBox("All fields are required.", MsgBoxStyle.Warning) : Return
        End If

        Dim isEdit = _existing IsNot Nothing
        Dim pinProvided = Not String.IsNullOrEmpty(txtPin.Text)

        If (Not isEdit OrElse pinProvided) AndAlso pinProvided Then
            If txtPin.Text <> txtPin2.Text Then MsgBox("PINs do not match.", MsgBoxStyle.Warning) : Return
            If txtPin.Text.Length < 4 OrElse Not txtPin.Text.All(Function(c) Char.IsDigit(c)) Then
                MsgBox("PIN must be 4-6 numeric digits.", MsgBoxStyle.Warning) : Return
            End If
        ElseIf Not isEdit AndAlso Not pinProvided Then
            MsgBox("PIN is required for new employees.", MsgBoxStyle.Warning) : Return
        End If

        If isEdit Then
            _existing.FullName = txtName.Text.Trim()
            _existing.Role = cmbRole.SelectedItem.ToString()
            _existing.Email = txtEmail.Text.Trim()
            _existing.Phone = txtPhone.Text.Trim()
            If pinProvided Then _existing.PinHash = HashPin(txtPin.Text)
            DataAccess.SaveEmployee(_existing)
            WriteAudit("EDIT_EMPLOYEE", _existing.EmployeeID)
            MsgBox("Employee updated.", MsgBoxStyle.Information)
        Else
            Dim emp As New EmployeeRecord With {
                .EmployeeID = GenerateEmployeeID(),
                .FullName = txtName.Text.Trim(),
                .Role = cmbRole.SelectedItem.ToString(),
                .PinHash = HashPin(txtPin.Text),
                .Email = txtEmail.Text.Trim(),
                .Phone = txtPhone.Text.Trim(),
                .DateRegistered = DateTime.Now,
                .IsActive = True
            }
            DataAccess.SaveEmployee(emp)
            WriteAudit("REGISTER_EMPLOYEE", emp.EmployeeID)
            MsgBox($"Employee registered.{vbCrLf}ID: {emp.EmployeeID}", MsgBoxStyle.Information)
        End If

        Me.DialogResult = DialogResult.OK
        Me.Close()
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Me.DialogResult = DialogResult.Cancel : Me.Close()
    End Sub
End Class

' ============================================================
'  frmSearch - Account search
' ============================================================
Public Class frmSearch
    Inherits Form

    Private WithEvents txtSearch As TextBox
    Private WithEvents btnSearch As Button
    Private WithEvents btnClose As Button
    Private WithEvents dgvResults As DataGridView

    Public Sub New()
        BuildUI()
    End Sub

    Private Sub BuildUI()
        Me.Text = "Search Customer Accounts"
        Me.Size = New Size(820, 520)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.BackColor = Color.FromArgb(245, 247, 250)
        Me.Font = New Font("Segoe UI", 9.5F)

        Dim pnlH As New Panel() With {.Size = New Size(820, 50), .BackColor = Color.FromArgb(150, 90, 20)}
        Dim lH As New Label() With {.Text = "🔍  Account Search", .Font = New Font("Segoe UI", 13, FontStyle.Bold), .ForeColor = Color.White, .AutoSize = True, .Location = New Point(16, 13)}
        pnlH.Controls.Add(lH)

        Dim lblHint As New Label() With {.Text = "Search by account number, name, email, phone, or ID number:", .Location = New Point(14, 64), .AutoSize = True, .Font = New Font("Segoe UI", 9, FontStyle.Bold)}

        txtSearch = New TextBox() With {.Location = New Point(14, 84), .Size = New Size(580, 28), .Font = New Font("Segoe UI", 11), .PlaceholderText = "Type to search..."}

        btnSearch = New Button() With {.Text = "Search", .Size = New Size(100, 28), .Location = New Point(602, 84), .BackColor = Color.FromArgb(150, 90, 20), .ForeColor = Color.White, .FlatStyle = FlatStyle.Flat, .Cursor = Cursors.Hand}
        btnSearch.FlatAppearance.BorderSize = 0

        dgvResults = New DataGridView() With {.Location = New Point(10, 122), .Size = New Size(790, 330), .BackgroundColor = Color.White, .BorderStyle = BorderStyle.None, .RowHeadersVisible = False, .SelectionMode = DataGridViewSelectionMode.FullRowSelect, .AllowUserToAddRows = False, .AllowUserToDeleteRows = False, .ReadOnly = True, .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill, .ColumnHeadersHeight = 36, .ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing}
        dgvResults.EnableHeadersVisualStyles = False
        dgvResults.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(150, 90, 20)
        dgvResults.ColumnHeadersDefaultCellStyle.ForeColor = Color.White
        dgvResults.ColumnHeadersDefaultCellStyle.Font = New Font("Segoe UI", 9.5F, FontStyle.Bold)
        dgvResults.DefaultCellStyle.SelectionBackColor = Color.FromArgb(245, 220, 190)
        dgvResults.DefaultCellStyle.SelectionForeColor = Color.Black
        dgvResults.RowTemplate.Height = 30

        For Each col As String In {"Account No.", "Name", "Phone", "Email", "ID No.", "Type", "Balance", "Status"}
            Dim dc As New DataGridViewTextBoxColumn()
            dc.HeaderText = col : dc.ReadOnly = True
            dgvResults.Columns.Add(dc)
        Next

        btnClose = New Button() With {.Text = "✕ Close", .Size = New Size(100, 30), .Location = New Point(700, 462), .BackColor = Color.FromArgb(140, 50, 50), .ForeColor = Color.White, .FlatStyle = FlatStyle.Flat, .Cursor = Cursors.Hand}
        btnClose.FlatAppearance.BorderSize = 0

        Me.Controls.AddRange(New Control() {pnlH, lblHint, txtSearch, btnSearch, dgvResults, btnClose})
        Me.AcceptButton = btnSearch
    End Sub

    Private Sub RunSearch()
        dgvResults.Rows.Clear()
        Dim q = txtSearch.Text.ToLower().Trim()
        If String.IsNullOrEmpty(q) Then Return

        Dim all = DataAccess.LoadAllCustomers()
        Dim found = all.FindAll(Function(c)
            Return c.AccountNumber.Contains(q) OrElse
                   c.FullName.ToLower().Contains(q) OrElse
                   c.Email.ToLower().Contains(q) OrElse
                   c.PhoneNumber.Contains(q) OrElse
                   c.IDNumber.ToLower().Contains(q)
        End Function)

        For Each c In found
            Dim i = dgvResults.Rows.Add()
            dgvResults.Rows(i).Cells(0).Value = c.AccountNumber
            dgvResults.Rows(i).Cells(1).Value = c.FullName
            dgvResults.Rows(i).Cells(2).Value = c.PhoneNumber
            dgvResults.Rows(i).Cells(3).Value = c.Email
            dgvResults.Rows(i).Cells(4).Value = c.IDNumber
            dgvResults.Rows(i).Cells(5).Value = c.AccountType
            dgvResults.Rows(i).Cells(6).Value = FormatMoney(c.Balance)
            dgvResults.Rows(i).Cells(7).Value = c.Status
        Next

        If found.Count = 0 Then
            MsgBox("No matching accounts found.", MsgBoxStyle.Information)
        End If
    End Sub

    Private Sub btnSearch_Click(sender As Object, e As EventArgs) Handles btnSearch.Click
        RunSearch()
    End Sub

    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub
End Class

' ============================================================
'  frmTransactionHistory - per-account statement
' ============================================================
Public Class frmTransactionHistory
    Inherits Form

    Private ReadOnly _accNum As String

    Public Sub New(accountNumber As String)
        _accNum = accountNumber
        BuildUI()
    End Sub

    Private Sub BuildUI()
        Dim customer = DataAccess.FindCustomerByAccount(_accNum)
        Me.Text = $"Transaction History — {_accNum}"
        Me.Size = New Size(860, 560)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.BackColor = Color.FromArgb(245, 247, 250)
        Me.Font = New Font("Segoe UI", 9.5F)

        Dim pnlH As New Panel() With {.Size = New Size(860, 70), .BackColor = Color.FromArgb(15, 76, 129)}
        Dim lH As New Label() With {.Text = $"📋  Statement: {customer?.FullName} ({_accNum})", .Font = New Font("Segoe UI", 12, FontStyle.Bold), .ForeColor = Color.White, .AutoSize = True, .Location = New Point(14, 10)}
        Dim lBal As New Label() With {.Text = $"Current Balance: {FormatMoney(customer?.Balance)}", .Font = New Font("Segoe UI", 10), .ForeColor = Color.FromArgb(180, 210, 240), .AutoSize = True, .Location = New Point(14, 38)}
        pnlH.Controls.AddRange(New Control() {lH, lBal})

        Dim dgv As New DataGridView() With {.Location = New Point(10, 80), .Size = New Size(840, 420), .BackgroundColor = Color.White, .BorderStyle = BorderStyle.None, .RowHeadersVisible = False, .SelectionMode = DataGridViewSelectionMode.FullRowSelect, .AllowUserToAddRows = False, .AllowUserToDeleteRows = False, .ReadOnly = True, .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill}
        dgv.ColumnHeadersHeight = 36
        dgv.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        dgv.EnableHeadersVisualStyles = False
        dgv.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(15, 76, 129)
        dgv.ColumnHeadersDefaultCellStyle.ForeColor = Color.White
        dgv.ColumnHeadersDefaultCellStyle.Font = New Font("Segoe UI", 9.5F, FontStyle.Bold)
        dgv.DefaultCellStyle.SelectionBackColor = Color.FromArgb(210, 225, 245)
        dgv.DefaultCellStyle.SelectionForeColor = Color.Black
        dgv.RowTemplate.Height = 28

        For Each col As String In {"Date & Time", "Type", "Amount", "Bal Before", "Bal After", "By", "Notes"}
            Dim dc As New DataGridViewTextBoxColumn()
            dc.HeaderText = col : dc.ReadOnly = True
            dgv.Columns.Add(dc)
        Next

        Dim txns = DataAccess.LoadTransactionsByAccount(_accNum)
        txns = txns.OrderByDescending(Function(t) t.Timestamp).ToList()

        For Each t In txns
            Dim i = dgv.Rows.Add()
            dgv.Rows(i).Cells(0).Value = t.Timestamp.ToString("dd MMM yyyy HH:mm:ss")
            dgv.Rows(i).Cells(1).Value = t.TransactionType
            dgv.Rows(i).Cells(2).Value = FormatMoney(t.Amount)
            dgv.Rows(i).Cells(3).Value = FormatMoney(t.BalanceBefore)
            dgv.Rows(i).Cells(4).Value = FormatMoney(t.BalanceAfter)
            dgv.Rows(i).Cells(5).Value = t.PerformedBy
            dgv.Rows(i).Cells(6).Value = t.Notes

            Select Case t.TransactionType
                Case "Deposit", "Transfer-In"
                    dgv.Rows(i).Cells(2).Style.ForeColor = Color.FromArgb(0, 140, 60)
                    dgv.Rows(i).Cells(2).Style.Font = New Font("Segoe UI", 9, FontStyle.Bold)
                Case "Withdrawal", "Transfer-Out"
                    dgv.Rows(i).Cells(2).Style.ForeColor = Color.FromArgb(180, 40, 40)
                    dgv.Rows(i).Cells(2).Style.Font = New Font("Segoe UI", 9, FontStyle.Bold)
            End Select
        Next

        Dim btnClose As New Button() With {.Text = "✕ Close", .Size = New Size(110, 32), .Location = New Point(740, 512), .BackColor = Color.FromArgb(140, 50, 50), .ForeColor = Color.White, .FlatStyle = FlatStyle.Flat, .Cursor = Cursors.Hand}
        btnClose.FlatAppearance.BorderSize = 0
        AddHandler btnClose.Click, Sub() Me.Close()

        Dim lblCount As New Label() With {.Text = $"  {txns.Count} transaction(s)", .Location = New Point(14, 518), .AutoSize = True, .Font = New Font("Segoe UI", 9)}

        Me.Controls.AddRange(New Control() {pnlH, dgv, btnClose, lblCount})
    End Sub
End Class
