# 🏦 NexaBank Management System
### Visual Basic .NET — Bank Management System
**Submission Date: 18th May 2026**

---

## 📋 System Overview

NexaBank Management System is a full-featured, professional bank management application built in **pure Visual Basic .NET** using Windows Forms. The system persists all data to **CSV flat-file databases** and includes complete audit logging.

---

## 🗂️ Project Files

| File | Purpose |
|---|---|
| `Program.vb` | Application entry point (`Sub Main`) |
| `Module1.vb` | Global constants, helpers, session state, security |
| `Models.vb` | Data record classes: `CustomerRecord`, `EmployeeRecord`, `TransactionRecord` |
| `DataAccess.vb` | All file I/O — CRUD operations for customers, employees, transactions |
| `frmLogin.vb` | Secure PIN login with lockout after 3 failed attempts |
| `frmDashboard.vb` | Main menu with live stats tiles and navigation |
| `frmCustomers.vb` | Customer account management (Add/Edit/Delete/View) + `frmCustomerDetail` |
| `frmTransactions.vb` | Deposit, Withdrawal, and Transfer operations |
| `frmEmployeesAndOthers.vb` | Employee management, Search, Transaction History |
| `frmReports.vb` | Reports: Summary, All Customers, All Transactions, Audit Log, CSV Export |
| `NexaBankMS.vbproj` | Project file (targets .NET 6 Windows) |

---

## ✅ Requirements Fulfilled

| Requirement | Implementation |
|---|---|
| 1. Register employees with unique PIN | `frmEmployees` → `frmEmployeeDetail` — SHA-256 hashed PINs, unique Employee IDs |
| 2. Admin: Add / View / Delete / Modify customers | `frmCustomers` (admin role guards delete) |
| 3. Deposit / Withdraw / Update bank details | `frmTransactions` tabs + Customer Edit |
| 4. Account creation captures all fields | Name, Account No., Phone, Address, Email, ID No., Account Type |
| 5. Deposit / Withdraw by account number + amount | Quick transaction forms with account lookup |
| 6. Reports: All customers, totals, withdrawals | `frmReports` — Summary, Customers, Transactions, Audit |
| 7. Data saved to file | CSV flat-files in `%AppData%\NexaBank\` |
| 8. Extra features | SHA-256 PIN hashing, login lockout, fund transfers, account statements, status management, CSV export, audit trail, real-time clock |

---

## 🔑 Default Admin Credentials

| Field | Value |
|---|---|
| Employee ID | `EMP-001` |
| PIN | `1234` |
| Role | Admin |

> ⚠️ **Change the default PIN immediately after first login** via Employee Management → Reset PIN.

---

## 🏗️ Architecture

```
┌─────────────────────────────────────────────────────┐
│                  PRESENTATION LAYER                  │
│  frmLogin → frmDashboard → (all feature forms)      │
├─────────────────────────────────────────────────────┤
│                  BUSINESS LOGIC                      │
│  Module1 (helpers, security, session)               │
├─────────────────────────────────────────────────────┤
│                   DATA ACCESS LAYER                  │
│  DataAccess.vb — all CSV read/write operations      │
├─────────────────────────────────────────────────────┤
│                    DATA STORAGE                      │
│  %AppData%\NexaBank\                                │
│  ├── customers.csv                                  │
│  ├── employees.csv                                  │
│  ├── transactions.csv                               │
│  └── audit_log.txt                                  │
└─────────────────────────────────────────────────────┘
```

---

## 🚀 How to Run

### Prerequisites
- Visual Studio 2022 (Community or higher)
- .NET 6 SDK with Windows Desktop workload

### Steps
1. Open Visual Studio 2022
2. **File → Open → Project/Solution**
3. Select `NexaBankMS.vbproj`
4. Press **F5** or click **▶ Start**
5. Login with `EMP-001` / PIN `1234`

### Alternative (Command Line)
```bash
cd BankManagementSystem
dotnet run
```

---

## 🔐 Security Features

- **SHA-256 PIN Hashing** — PINs are never stored in plaintext
- **Login Lockout** — 30-second lockout after 3 failed attempts
- **Role-Based Access** — Admin/Manager vs Teller permissions
- **Full Audit Trail** — Every action logged with timestamp and employee ID
- **Atomic Transactions** — Balance updates and transaction records saved together

---

## 📊 Features Summary

### Customer Management
- Open new accounts (Savings / Current / Fixed Deposit)
- Edit all customer details
- Freeze / Unfreeze / Close accounts
- Delete accounts (Admin only, zero balance required)
- Filter and search by name or account number

### Transactions
- **Deposit** — Credit funds to any active account
- **Withdrawal** — Debit funds with balance check
- **Transfer** — Move funds between accounts atomically
- **Account Lookup** — See name and balance before transacting
- **Transaction History** — Full statement per account

### Employee Management (Admin/Manager)
- Register new employees with roles (Teller / Manager / Admin)
- Edit employee details
- Activate / Deactivate accounts
- Reset PINs securely

### Reports
- **Summary** — Key metrics at a glance
- **All Customers** — Full roster with balances
- **All Transactions** — Complete ledger with colour-coding
- **Audit Log** — System-level event trail
- **Export** — One-click CSV export of all data

---

## 📁 Data Files Location

```
Windows: C:\Users\{YourName}\AppData\Roaming\NexaBank\
```

---

*NexaBank Management System — Built with Visual Basic .NET | Windows Forms | CSV Data Persistence*
