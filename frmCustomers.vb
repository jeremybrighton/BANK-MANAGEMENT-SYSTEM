' ============================================================
'  BANK MANAGEMENT SYSTEM - frmCustomers.vb
'  Full CRUD for customer accounts with validation
' ============================================================
Imports System.Drawing
Imports System.Windows.Forms

Public Class frmCustomers
    Inherits Form

    Private WithEvents dgvCustomers As DataGridView
    Private WithEvents btnAdd As Button
    Private WithEvents btnEdit As Button
    Private WithEvents btnDelete As Button
    Private WithEvents btnRefresh As Button
    Private WithEvents btnClose As Button
    Private WithEvents btnViewHistory As Button
    Private WithEvents txtFilter As TextBox
    Private WithEvents cmbStatusFilter As ComboBox
    Private pnlTop As Panel
    Private pnlBottom As Panel
    Private lblCount As Label

    Public Sub New()
        BuildUI()
        LoadCustomersGrid()
    End Sub

    Private Sub BuildUI()
        Me.Text = "Customer Account Management"
        Me.Size = New Size(1050, 620)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.BackColor = Color.FromArgb(245, 247, 250)
        Me.Font = New Font("Segoe UI", 9.5F)

        ' ── Top toolbar ───────────────────────────────────────
        pnlTop = New Panel()
        pnlTop.Size = New Size(1050, 55)
        pnlTop.BackColor = Color.FromArgb(15, 76, 129)

        Dim lblTitle As New Label()
        lblTitle.Text = "👥  Customer Accounts"
        lblTitle.Font = New Font("Segoe UI", 13, FontStyle.Bold)
        lblTitle.ForeColor = Color.White
        lblTitle.AutoSize = True
        lblTitle.Location = New Point(14, 15)

        txtFilter = New TextBox()
        txtFilter.PlaceholderText = "🔍 Search by name or account..."
        txtFilter.Size = New Size(220, 28)
        txtFilter.Location = New Point(580, 14)
        txtFilter.Font = New Font("Segoe UI", 9.5F)

        cmbStatusFilter = New ComboBox()
        cmbStatusFilter.Size = New Size(120, 28)
        cmbStatusFilter.Location = New Point(808, 14)
        cmbStatusFilter.DropDownStyle = ComboBoxStyle.DropDownList
        cmbStatusFilter.Items.AddRange(New Object() {"All Status", "Active", "Frozen", "Closed"})
        cmbStatusFilter.SelectedIndex = 0

        pnlTop.Controls.AddRange(New Control() {lblTitle, txtFilter, cmbStatusFilter})

        ' ── DataGridView ──────────────────────────────────────
        dgvCustomers = New DataGridView()
        dgvCustomers.Location = New Point(10, 65)
        dgvCustomers.Size = New Size(1030, 440)
        dgvCustomers.BackgroundColor = Color.White
        dgvCustomers.BorderStyle = BorderStyle.None
        dgvCustomers.RowHeadersVisible = False
        dgvCustomers.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        dgvCustomers.MultiSelect = False
        dgvCustomers.AllowUserToAddRows = False
        dgvCustomers.AllowUserToDeleteRows = False
        dgvCustomers.ReadOnly = True
        dgvCustomers.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        dgvCustomers.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        dgvCustomers.ColumnHeadersHeight = 36
        StyleGrid(dgvCustomers)

        ' Define columns
        Dim cols As String() = {"Account No.", "Full Name", "Phone", "Email", "ID Number", "Account Type", "Balance", "Status", "Date Opened"}
        For Each col As String In cols
            Dim dc As New DataGridViewTextBoxColumn()
            dc.HeaderText = col
            dc.ReadOnly = True
            dgvCustomers.Columns.Add(dc)
        Next

        ' ── Bottom panel ──────────────────────────────────────
        pnlBottom = New Panel()
        pnlBottom.Size = New Size(1050, 50)
        pnlBottom.Location = New Point(0, 515)
        pnlBottom.BackColor = Color.FromArgb(235, 240, 248)

        lblCount = New Label()
        lblCount.Text = "0 records"
        lblCount.Font = New Font("Segoe UI", 9)
        lblCount.ForeColor = Color.FromArgb(80, 90, 110)
        lblCount.Location = New Point(14, 15)
        lblCount.AutoSize = True

        btnAdd = MakeButton("➕ Add Account", Color.FromArgb(22, 130, 90), New Point(430, 10))
        btnEdit = MakeButton("✏ Edit", Color.FromArgb(15, 76, 129), New Point(570, 10))
        btnDelete = MakeButton("🗑 Delete", Color.FromArgb(200, 50, 50), New Point(668, 10))
        btnViewHistory = MakeButton("📋 History", Color.FromArgb(100, 70, 160), New Point(766, 10))
        btnRefresh = MakeButton("↻ Refresh", Color.FromArgb(80, 90, 110), New Point(864, 10))
        btnClose = MakeButton("✕ Close", Color.FromArgb(140, 50, 50), New Point(962, 10))

        pnlBottom.Controls.AddRange(New Control() {
            lblCount, btnAdd, btnEdit, btnDelete, btnViewHistory, btnRefresh, btnClose})

        Me.Controls.AddRange(New Control() {pnlTop, dgvCustomers, pnlBottom})
    End Sub

    Private Sub StyleGrid(dgv As DataGridView)
        dgv.EnableHeadersVisualStyles = False
        dgv.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(30, 90, 150)
        dgv.ColumnHeadersDefaultCellStyle.ForeColor = Color.White
        dgv.ColumnHeadersDefaultCellStyle.Font = New Font("Segoe UI", 9.5F, FontStyle.Bold)
        dgv.ColumnHeadersDefaultCellStyle.SelectionBackColor = Color.FromArgb(30, 90, 150)
        dgv.DefaultCellStyle.SelectionBackColor = Color.FromArgb(210, 225, 245)
        dgv.DefaultCellStyle.SelectionForeColor = Color.Black
        dgv.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(248, 250, 255)
        dgv.GridColor = Color.FromArgb(220, 225, 235)
        dgv.DefaultCellStyle.Font = New Font("Segoe UI", 9)
        dgv.RowTemplate.Height = 30
    End Sub

    Private Function MakeButton(text As String, color As Color, loc As Point) As Button
        Dim btn As New Button()
        btn.Text = text
        btn.Size = New Size(90, 30)
        btn.Location = loc
        btn.BackColor = color
        btn.ForeColor = Color.White
        btn.FlatStyle = FlatStyle.Flat
        btn.FlatAppearance.BorderSize = 0
        btn.Font = New Font("Segoe UI", 8.5F)
        btn.Cursor = Cursors.Hand
        Return btn
    End Function

    Private Sub LoadCustomersGrid()
        dgvCustomers.Rows.Clear()
        Dim all = DataAccess.LoadAllCustomers()

        ' Apply filters
        Dim filter As String = txtFilter.Text.ToLower().Trim()
        Dim statusFilter As String = If(cmbStatusFilter.SelectedIndex <= 0, "", cmbStatusFilter.SelectedItem.ToString())

        Dim filtered = all.FindAll(Function(c)
            Dim nameMatch = String.IsNullOrEmpty(filter) OrElse
                            c.FullName.ToLower().Contains(filter) OrElse
                            c.AccountNumber.Contains(filter)
            Dim statusMatch = String.IsNullOrEmpty(statusFilter) OrElse c.Status = statusFilter
            Return nameMatch AndAlso statusMatch
        End Function)

        For Each c In filtered
            Dim idx = dgvCustomers.Rows.Add()
            Dim row = dgvCustomers.Rows(idx)
            row.Cells(0).Value = c.AccountNumber
            row.Cells(1).Value = c.FullName
            row.Cells(2).Value = c.PhoneNumber
            row.Cells(3).Value = c.Email
            row.Cells(4).Value = c.IDNumber
            row.Cells(5).Value = c.AccountType
            row.Cells(6).Value = FormatMoney(c.Balance)
            row.Cells(7).Value = c.Status
            row.Cells(8).Value = c.DateOpened.ToString("dd MMM yyyy")

            ' Colour-code status
            Select Case c.Status
                Case "Active"
                    row.Cells(7).Style.ForeColor = Color.FromArgb(0, 140, 70)
                    row.Cells(7).Style.Font = New Font("Segoe UI", 9, FontStyle.Bold)
                Case "Frozen"
                    row.Cells(7).Style.ForeColor = Color.FromArgb(150, 90, 0)
                Case "Closed"
                    row.Cells(7).Style.ForeColor = Color.FromArgb(180, 40, 40)
                    row.DefaultCellStyle.ForeColor = Color.FromArgb(160, 160, 170)
            End Select
        Next

        lblCount.Text = $"{filtered.Count} of {all.Count} records"
    End Sub

    Private Function GetSelectedAccountNumber() As String
        If dgvCustomers.SelectedRows.Count = 0 Then Return Nothing
        Return dgvCustomers.SelectedRows(0).Cells(0).Value?.ToString()
    End Function

    ' ── Events ────────────────────────────────────────────────

    Private Sub txtFilter_TextChanged(sender As Object, e As EventArgs) Handles txtFilter.TextChanged
        LoadCustomersGrid()
    End Sub

    Private Sub cmbStatusFilter_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbStatusFilter.SelectedIndexChanged
        LoadCustomersGrid()
    End Sub

    Private Sub btnAdd_Click(sender As Object, e As EventArgs) Handles btnAdd.Click
        Dim f As New frmCustomerDetail(Nothing)
        If f.ShowDialog() = DialogResult.OK Then
            LoadCustomersGrid()
        End If
    End Sub

    Private Sub btnEdit_Click(sender As Object, e As EventArgs) Handles btnEdit.Click
        Dim accNum = GetSelectedAccountNumber()
        If accNum Is Nothing Then
            MessageBox.Show("Please select a customer first.", "No Selection", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If
        Dim customer = DataAccess.FindCustomerByAccount(accNum)
        Dim f As New frmCustomerDetail(customer)
        If f.ShowDialog() = DialogResult.OK Then
            LoadCustomersGrid()
        End If
    End Sub

    Private Sub btnDelete_Click(sender As Object, e As EventArgs) Handles btnDelete.Click
        If Not IsAdminSession Then
            MessageBox.Show("Only Admin users can delete accounts.", "Access Denied",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        Dim accNum = GetSelectedAccountNumber()
        If accNum Is Nothing Then
            MessageBox.Show("Please select a customer first.", "No Selection", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If

        Dim customer = DataAccess.FindCustomerByAccount(accNum)
        If customer.Balance > 0 Then
            MessageBox.Show($"Cannot delete account with balance of {FormatMoney(customer.Balance)}. Withdraw funds first.",
                            "Cannot Delete", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        If MessageBox.Show($"Permanently delete account {accNum} ({customer.FullName})?" & vbCrLf &
                           "This cannot be undone.",
                           "Confirm Delete", MessageBoxButtons.YesNo,
                           MessageBoxIcon.Warning) = DialogResult.Yes Then
            If DataAccess.DeleteCustomer(accNum) Then
                MessageBox.Show("Account deleted.", "Done", MessageBoxButtons.OK, MessageBoxIcon.Information)
                LoadCustomersGrid()
            Else
                MessageBox.Show("Failed to delete. Please try again.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        End If
    End Sub

    Private Sub btnViewHistory_Click(sender As Object, e As EventArgs) Handles btnViewHistory.Click
        Dim accNum = GetSelectedAccountNumber()
        If accNum Is Nothing Then
            MessageBox.Show("Please select a customer first.", "No Selection", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If
        Dim f As New frmTransactionHistory(accNum)
        f.ShowDialog()
    End Sub

    Private Sub btnRefresh_Click(sender As Object, e As EventArgs) Handles btnRefresh.Click
        LoadCustomersGrid()
    End Sub

    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub dgvCustomers_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvCustomers.CellDoubleClick
        btnEdit_Click(sender, e)
    End Sub

End Class

' ============================================================
'  frmCustomerDetail - Add / Edit single customer
' ============================================================
Public Class frmCustomerDetail
    Inherits Form

    Private ReadOnly _existing As CustomerRecord  ' Nothing = Add mode

    Private WithEvents btnSave As Button
    Private WithEvents btnCancel As Button

    Private txtName As TextBox
    Private txtPhone As TextBox
    Private txtAddress As TextBox
    Private txtEmail As TextBox
    Private txtIDNumber As TextBox
    Private txtInitialDeposit As TextBox
    Private cmbAccountType As ComboBox
    Private cmbStatus As ComboBox
    Private lblAccountNum As Label

    Public Sub New(existing As CustomerRecord)
        _existing = existing
        BuildUI()
        If existing IsNot Nothing Then PopulateFields()
    End Sub

    Private Sub BuildUI()
        Dim isEdit = _existing IsNot Nothing
        Me.Text = If(isEdit, "Edit Customer Account", "New Customer Account")
        Me.Size = New Size(520, 580)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.BackColor = Color.FromArgb(245, 247, 250)
        Me.Font = New Font("Segoe UI", 9.5F)

        ' Header
        Dim pnlHead As New Panel()
        pnlHead.Size = New Size(520, 50)
        pnlHead.BackColor = Color.FromArgb(15, 76, 129)

        Dim lblHead As New Label()
        lblHead.Text = If(isEdit, "✏  Edit Account", "➕  New Customer Account")
        lblHead.Font = New Font("Segoe UI", 12, FontStyle.Bold)
        lblHead.ForeColor = Color.White
        lblHead.AutoSize = True
        lblHead.Location = New Point(16, 13)
        pnlHead.Controls.Add(lblHead)

        Dim y As Integer = 70
        Const GAP As Integer = 56

        ' Account Number (read-only in edit, auto in add)
        Dim lblAN = MakeLbl("Account Number", New Point(20, y))
        lblAccountNum = New Label()
        lblAccountNum.Size = New Size(460, 28)
        lblAccountNum.Location = New Point(20, y + 20)
        lblAccountNum.Font = New Font("Segoe UI", 10.5F, FontStyle.Bold)
        lblAccountNum.ForeColor = Color.FromArgb(15, 76, 129)
        lblAccountNum.Text = If(isEdit, _existing.AccountNumber, "(Auto-generated on save)")
        y += GAP

        txtName = MakeTxt(New Point(20, y + 20), 460)
        y += GAP

        txtPhone = MakeTxt(New Point(20, y + 20), 460)
        y += GAP

        txtEmail = MakeTxt(New Point(20, y + 20), 460)
        y += GAP

        txtIDNumber = MakeTxt(New Point(20, y + 20), 220)
        y += GAP

        txtAddress = MakeTxt(New Point(20, y + 20), 460)
        y += GAP

        cmbAccountType = New ComboBox()
        cmbAccountType.Size = New Size(200, 28)
        cmbAccountType.Location = New Point(20, y + 20)
        cmbAccountType.DropDownStyle = ComboBoxStyle.DropDownList
        cmbAccountType.Items.AddRange(New Object() {"Savings", "Current", "Fixed Deposit"})
        cmbAccountType.SelectedIndex = 0

        cmbStatus = New ComboBox()
        cmbStatus.Size = New Size(140, 28)
        cmbStatus.Location = New Point(240, y + 20)
        cmbStatus.DropDownStyle = ComboBoxStyle.DropDownList
        cmbStatus.Items.AddRange(New Object() {"Active", "Frozen", "Closed"})
        cmbStatus.SelectedIndex = 0
        If Not isEdit Then cmbStatus.Enabled = False  ' New accounts always Active

        y += GAP

        Dim lblDep As New Label()
        lblDep.Text = If(isEdit, "", "Initial Deposit (KES)")
        lblDep.Font = New Font("Segoe UI", 9, FontStyle.Bold)
        lblDep.ForeColor = Color.FromArgb(60, 60, 80)
        lblDep.Location = New Point(20, y)
        lblDep.AutoSize = True

        txtInitialDeposit = New TextBox()
        txtInitialDeposit.Size = New Size(200, 28)
        txtInitialDeposit.Location = New Point(20, y + 20)
        txtInitialDeposit.Text = "0.00"
        txtInitialDeposit.Enabled = Not isEdit

        y += GAP + 10

        btnSave = New Button()
        btnSave.Text = If(isEdit, "💾  Save Changes", "✔  Create Account")
        btnSave.Size = New Size(200, 38)
        btnSave.Location = New Point(20, y)
        btnSave.BackColor = Color.FromArgb(22, 130, 90)
        btnSave.ForeColor = Color.White
        btnSave.FlatStyle = FlatStyle.Flat
        btnSave.FlatAppearance.BorderSize = 0
        btnSave.Font = New Font("Segoe UI", 10, FontStyle.Bold)
        btnSave.Cursor = Cursors.Hand

        btnCancel = New Button()
        btnCancel.Text = "✕  Cancel"
        btnCancel.Size = New Size(120, 38)
        btnCancel.Location = New Point(240, y)
        btnCancel.BackColor = Color.FromArgb(160, 50, 50)
        btnCancel.ForeColor = Color.White
        btnCancel.FlatStyle = FlatStyle.Flat
        btnCancel.FlatAppearance.BorderSize = 0
        btnCancel.Font = New Font("Segoe UI", 10)
        btnCancel.Cursor = Cursors.Hand

        ' Labels using y coordinates baked into label array for clarity
        Dim lbls() As Label = {
            MakeLbl("Account Number", New Point(20, 70)),
            MakeLbl("Full Name *", New Point(20, 70 + GAP)),
            MakeLbl("Phone Number *", New Point(20, 70 + GAP * 2)),
            MakeLbl("Email Address *", New Point(20, 70 + GAP * 3)),
            MakeLbl("National ID / Passport Number *", New Point(20, 70 + GAP * 4)),
            MakeLbl("Physical Address *", New Point(20, 70 + GAP * 5)),
            MakeLbl("Account Type", New Point(20, 70 + GAP * 6)),
            MakeLbl("Status", New Point(240, 70 + GAP * 6)),
            lblDep
        }

        Me.Controls.Add(pnlHead)
        Me.Controls.Add(lblAccountNum)
        Me.Controls.AddRange(lbls)
        Me.Controls.AddRange(New Control() {
            txtName, txtPhone, txtEmail, txtIDNumber, txtAddress,
            cmbAccountType, cmbStatus,
            txtInitialDeposit, btnSave, btnCancel})

        Me.AcceptButton = btnSave
        Me.CancelButton = btnCancel
    End Sub

    Private Function MakeLbl(text As String, loc As Point) As Label
        Dim l As New Label()
        l.Text = text
        l.Font = New Font("Segoe UI", 8.5F, FontStyle.Bold)
        l.ForeColor = Color.FromArgb(60, 60, 80)
        l.Location = loc
        l.AutoSize = True
        Return l
    End Function

    Private Function MakeTxt(loc As Point, width As Integer) As TextBox
        Dim t As New TextBox()
        t.Location = loc
        t.Size = New Size(width, 28)
        t.Font = New Font("Segoe UI", 10)
        t.BorderStyle = BorderStyle.FixedSingle
        Return t
    End Function

    Private Sub PopulateFields()
        txtName.Text = _existing.FullName
        txtPhone.Text = _existing.PhoneNumber
        txtEmail.Text = _existing.Email
        txtIDNumber.Text = _existing.IDNumber
        txtAddress.Text = _existing.Address
        cmbAccountType.SelectedItem = _existing.AccountType
        cmbStatus.SelectedItem = _existing.Status
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        ' Validate
        If String.IsNullOrWhiteSpace(txtName.Text) OrElse
           String.IsNullOrWhiteSpace(txtPhone.Text) OrElse
           String.IsNullOrWhiteSpace(txtEmail.Text) OrElse
           String.IsNullOrWhiteSpace(txtIDNumber.Text) OrElse
           String.IsNullOrWhiteSpace(txtAddress.Text) Then
            MessageBox.Show("All fields marked with * are required.", "Validation Error",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        ' Validate email basic format
        If Not txtEmail.Text.Contains("@") OrElse Not txtEmail.Text.Contains(".") Then
            MessageBox.Show("Please enter a valid email address.", "Validation Error",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        Dim initialDeposit As Decimal = 0
        If _existing Is Nothing Then
            If Not Decimal.TryParse(txtInitialDeposit.Text, initialDeposit) OrElse initialDeposit < 0 Then
                MessageBox.Show("Please enter a valid initial deposit amount.", "Validation Error",
                                MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If
        End If

        If _existing Is Nothing Then
            ' ADD new
            Dim c As New CustomerRecord With {
                .AccountNumber = GenerateAccountNumber(),
                .FullName = txtName.Text.Trim(),
                .PhoneNumber = txtPhone.Text.Trim(),
                .Email = txtEmail.Text.Trim(),
                .IDNumber = txtIDNumber.Text.Trim(),
                .Address = txtAddress.Text.Trim(),
                .AccountType = cmbAccountType.SelectedItem.ToString(),
                .Status = "Active",
                .Balance = initialDeposit,
                .DateOpened = DateTime.Now
            }
            DataAccess.SaveCustomer(c)

            If initialDeposit > 0 Then
                DataAccess.Deposit(c.AccountNumber, initialDeposit, "Initial deposit on account opening")
            End If

            WriteAudit("CREATE_ACCOUNT", $"AccNum={c.AccountNumber}, Name={c.FullName}")
            MessageBox.Show($"Account created successfully!{vbCrLf}Account Number: {c.AccountNumber}",
                            "Account Created", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Else
            ' EDIT existing
            _existing.FullName = txtName.Text.Trim()
            _existing.PhoneNumber = txtPhone.Text.Trim()
            _existing.Email = txtEmail.Text.Trim()
            _existing.IDNumber = txtIDNumber.Text.Trim()
            _existing.Address = txtAddress.Text.Trim()
            _existing.AccountType = cmbAccountType.SelectedItem.ToString()
            _existing.Status = cmbStatus.SelectedItem.ToString()
            DataAccess.SaveCustomer(_existing)
            WriteAudit("EDIT_ACCOUNT", $"AccNum={_existing.AccountNumber}")
            MessageBox.Show("Account updated successfully.", "Saved", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If

        Me.DialogResult = DialogResult.OK
        Me.Close()
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Me.DialogResult = DialogResult.Cancel
        Me.Close()
    End Sub

End Class
