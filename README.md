# GHC Pet Mart

> A desktop-based pet shop management and point-of-sale system built with **VB.NET Windows Forms**, **MySQL**, and **.NET Framework 4.7.2**.

[![Language](https://img.shields.io/badge/Language-VB.NET-blue)](https://learn.microsoft.com/en-us/dotnet/visual-basic/)
[![Platform](https://img.shields.io/badge/Platform-Windows-0078D6)](https://www.microsoft.com/windows)
[![Database](https://img.shields.io/badge/Database-MySQL-4479A1)](https://www.mysql.com/)
[![Framework](https://img.shields.io/badge/.NET%20Framework-4.7.2-512BD4)](https://learn.microsoft.com/en-us/dotnet/framework/)

## Table of Contents

- [About the Project](#about-the-project)
- [Project Details](#project-details)
- [Features](#features)
- [Technology Stack](#technology-stack)
- [Project Structure](#project-structure)
- [Requirements](#requirements)
- [Installation and Setup](#installation-and-setup)
  - [1. Clone the Repository](#1-clone-the-repository)
  - [2. Install the Prerequisites](#2-install-the-prerequisites)
  - [3. Prepare the MySQL Database](#3-prepare-the-mysql-database)
  - [4. Configure the Database Connection](#4-configure-the-database-connection)
  - [5. Open and Restore the Project](#5-open-and-restore-the-project)
  - [6. Build and Run](#6-build-and-run)
- [How to Use the System](#how-to-use-the-system)
- [Interactive Setup Checklist](#interactive-setup-checklist)
- [Troubleshooting](#troubleshooting)
- [Known Setup Notes](#known-setup-notes)
- [Contributors](#contributors)
- [Academic Information](#academic-information)
- [License](#license)

## About the Project

**GHC Pet Mart** is a Windows desktop application designed to support common pet shop operations in one place. It combines customer-facing shopping and order workflows with employee/admin tools for managing products, stock, customers, suppliers, purchases, sales, expenses, and reports.

The project uses Windows Forms for its graphical interface and MySQL for data storage. The application contains separate forms for the different tasks, with database access shared through `DBConnection.vb`.

## Project Details

| Item | Details |
|---|---|
| Project name | GHC Pet Mart |
| Application type | Windows desktop application |
| Main language | Visual Basic .NET (VB.NET) |
| UI framework | Windows Forms |
| Target framework | .NET Framework 4.7.2 |
| Database | MySQL |
| Database name configured in source | `ghc_petmart_3nf` |
| IDE | Visual Studio (Windows desktop development workload) |

## Features

The repository includes forms and code for the following areas. Availability of each workflow may depend on a correctly configured database and user role.

<details>
<summary><strong>Shopping and customer accounts</strong></summary>

- User login and account registration
- Product browsing in the shop interface
- Product search, category filtering, and sorting
- Shopping cart
- Customer order history
- Customer profile and account-related screens
- Order/receipt-related screens

</details>

<details>
<summary><strong>Point of Sale (POS)</strong></summary>

- Search for products
- Add selected products and quantities to a cart
- Select a customer when applicable
- Calculate totals
- Enter the amount paid and calculate change
- Generate/show receipt-related information

</details>

<details>
<summary><strong>Inventory and administration</strong></summary>

- Product management, including add/edit/delete workflows
- Inventory viewing and stock management
- Customer management
- Employee management
- Supplier management
- Purchase orders
- Sales history and transaction editing
- Expense tracking
- Application settings

</details>

<details>
<summary><strong>Reports</strong></summary>

- Report types include sales, inventory, customers, suppliers, and total income
- Report periods include daily, weekly, monthly, annual, and custom date ranges
- PDF export functionality is included in the report form

</details>

## Technology Stack

- **VB.NET** — application logic
- **Windows Forms** — desktop graphical interface
- **MySQL** — relational database
- **MySql.Data** — .NET MySQL connector
- **iTextSharp** — PDF report generation
- **Crystal Reports assemblies** — referenced by the project
- **.NET Framework 4.7.2** — target runtime

## Project Structure

The main solution and project files are inside the `GHC_PetMart/` directory.

```text
GHC_PetMart/
├── GHC_PetMart2.slnx
└── GHC_PetMart/
    ├── GHC_PetMart.vbproj
    ├── App.config
    ├── DBConnection.vb
    ├── ReceiptGenerator.vb
    ├── frmLogin.vb
    ├── frmRegister.vb
    ├── frmMain.vb
    ├── frmShop.vb
    ├── frmCart.vb
    ├── frmPOS.vb
    ├── frmProducts.vb
    ├── frmInventory.vb
    ├── frmCustomers.vb
    ├── frmSuppliers.vb
    ├── frmPurchaseOrders.vb
    ├── frmSalesHistory.vb
    ├── frmReports.vb
    └── ...
```

The project also contains `.Designer.vb` and `.resx` files for Windows Forms UI layouts and resources, plus `packages.config` for NuGet dependencies.

## Requirements

Before running the application, make sure you have:

- A Windows PC
- [Visual Studio](https://visualstudio.microsoft.com/) with the **.NET desktop development** workload
- MySQL Server, with a local MySQL service running
- A MySQL client such as [MySQL Workbench](https://dev.mysql.com/downloads/workbench/) or another SQL client
- Git, if you plan to clone/push using the command line
- The database schema and any required seed/sample data for `ghc_petmart_3nf`

> **Important:** The source connection settings identify the database name, but the repository archive inspected for this README did not contain a database `.sql` dump. You must obtain the correct schema/backup from the project authors or an existing project setup before the application can work fully. Do not assume that creating an empty database is sufficient.

## Installation and Setup

### 1. Clone the Repository

Open Terminal, PowerShell, or Git Bash and run:

```bash
git clone https://github.com/47yarch/GHC_PetMart.git
cd GHC_PetMart
```

Alternatively, open the repository page and choose **Code → Download ZIP**, then extract it.

### 2. Install the Prerequisites

1. Install Visual Studio on Windows.
2. In Visual Studio Installer, select **.NET desktop development**.
3. Install MySQL Server and start the MySQL service.
4. (Optional) Install MySQL Workbench to manage the database visually.

### 3. Prepare the MySQL Database

1. Obtain the project's actual SQL schema/backup and any required sample data from the repository or project authors.
2. Open the `.sql` file in MySQL Workbench or your preferred MySQL client.
3. Connect to the intended MySQL server.
4. Execute/import the script.
5. Confirm that the database named **`ghc_petmart_3nf`** exists and contains the tables expected by the application.

If you already have the database on your computer, verify its name and connection details before proceeding.

<details>
<summary><strong>Optional: create the database container only (this does not create the tables)</strong></summary>

If you are preparing a fresh local server, this SQL creates only an empty database:

```sql
CREATE DATABASE ghc_petmart_3nf;
```

This is **not enough to run the application** until the required tables, relationships, and any required initial records have been imported.

</details>

### 4. Configure the Database Connection

1. In Visual Studio, open `GHC_PetMart/DBConnection.vb`.
2. Find the connection settings near the top of the file.
3. Set them to match your local MySQL setup.

The current source defaults are:

```vb
Public Const SERVER As String = "localhost"
Public Const DATABASE As String = "ghc_petmart_3nf"
Public Const DBUSER As String = "root"
Public Const DBPASS As String = ""
```

For example, if your local MySQL account has a password, update `DBPASS` to that password. If your server uses a different host, port, database name, or username, adjust the connection string accordingly.

**Security note:** These are development defaults. Avoid committing real passwords or production credentials to a public repository. For a deployed application, store secrets outside source control and use a least-privilege database account.

### 5. Open and Restore the Project

1. Open Visual Studio.
2. Choose **Open a project or solution**.
3. Browse to the repository folder and open `GHC_PetMart/GHC_PetMart2.slnx` if supported by your Visual Studio version. If that solution format is not supported, open `GHC_PetMart/GHC_PetMart/GHC_PetMart.vbproj` directly.
4. Allow Visual Studio to load the project.
5. If prompted, restore NuGet packages.
6. Check that the project's target framework is **.NET Framework 4.7.2**.
7. Resolve any missing third-party references using the versions specified in `packages.config` and the project file.

The project references third-party components including MySQL Connector/NET, iTextSharp, and Crystal Reports assemblies. Some Crystal Reports components may require the matching SAP Crystal Reports runtime/development installation on the machine. Install only the version compatible with the project's references.

### 6. Build and Run

1. In Visual Studio, select **Build → Rebuild Solution** (or rebuild the project if you opened the `.vbproj` directly).
2. Review the **Error List** and resolve missing packages/references first.
3. Set the Windows Forms project as the startup project if Visual Studio has not done so.
4. Press **F5** to run with debugging, or **Ctrl+F5** to run without debugging.
5. Use an existing account from the database, or register an account if the database and registration workflow are configured.

The exact first screen and available navigation depend on the application's startup settings and account role.

## How to Use the System

<details>
<summary><strong>Customer workflow</strong></summary>

1. Launch the application.
2. Sign in with an existing customer account, or use the registration screen if available.
3. Open the shop screen.
4. Search for products or browse by category; use sorting/filter controls if available.
5. Open a product and add items to the cart.
6. Review quantities and cart contents.
7. Continue through the order/checkout flow shown by the application.
8. Check the order history screen to review previous orders.

</details>

<details>
<summary><strong>Employee/admin workflow</strong></summary>

1. Sign in with an account that has the appropriate role.
2. Open the main dashboard.
3. Use **Products** to add, edit, search, filter, or remove products.
4. Use **Inventory** to review stock.
5. Use **Customers**, **Employees**, and **Suppliers** to manage related records.
6. Use **Purchase Orders** to manage restocking workflows.
7. Use **Point of Sale** to search for products, add quantities to the cart, select a customer if needed, enter the amount paid, and complete the transaction.
8. Use **Sales History** to review transactions.
9. Use **Expenses** to record business expenses.
10. Use **Reports** to choose a report type and period, generate the report, and export it to PDF if required.
11. Log out when finished.

Only use administrative functions with an account authorized for that role. Labels and exact workflows can differ depending on the current implementation and database contents.

</details>

## Interactive Setup Checklist

Tick items off as you complete the setup. This checklist is interactive in GitHub Markdown when viewed or edited in the repository.

- [ ] Repository cloned or ZIP extracted
- [ ] Visual Studio installed with **.NET desktop development**
- [ ] MySQL Server installed and running
- [ ] Correct `ghc_petmart_3nf` database schema imported
- [ ] Database connection values checked in `DBConnection.vb`
- [ ] NuGet packages restored
- [ ] Required Crystal Reports components installed, if needed
- [ ] Solution/project builds successfully
- [ ] Application starts without a database connection error
- [ ] Login/registration tested with a valid account
- [ ] Main customer and admin workflows tested

## Troubleshooting

<details>
<summary><strong>“Unable to connect to any of the specified MySQL hosts” or connection errors</strong></summary>

- Confirm that the MySQL service is running.
- Check `SERVER`, `DATABASE`, `DBUSER`, and `DBPASS` in `DBConnection.vb`.
- Verify the database exists and that the MySQL user has permission to access it.
- Check the server port and firewall settings if MySQL is not using the default local configuration.

</details>

<details>
<summary><strong>“Unknown database 'ghc_petmart_3nf'”</strong></summary>

The database has not been created on the server, or its name differs from the configured value. Import the correct project database backup and confirm the name matches `DATABASE` in `DBConnection.vb`.

</details>

<details>
<summary><strong>“Table doesn't exist” or SQL errors when opening a screen</strong></summary>

An empty database is not sufficient. Import the project's actual schema and required seed data. Confirm that table and column names match the queries used by the application. Ask the project authors for the SQL dump if it is not included in the repository.

</details>

<details>
<summary><strong>Missing NuGet packages or assembly references</strong></summary>

- Restore NuGet packages in Visual Studio.
- Check `packages.config` and the `.vbproj` references for expected versions.
- Rebuild after restore.
- If a third-party DLL is referenced by a local path, verify that the expected package is installed at that path.

</details>

<details>
<summary><strong>Crystal Reports-related errors</strong></summary>

The project references Crystal Reports assemblies. If those references are unavailable, install the compatible SAP Crystal Reports developer/runtime components required by the project, then reopen and rebuild it. Match the runtime architecture (32-bit/64-bit) to the application's build and installed components.

</details>

<details>
<summary><strong>Login does not work</strong></summary>

- Confirm the database has the expected user/account records.
- Check whether registration requirements are satisfied.
- Use an account created for this application; do not assume a default admin username or password unless the project authors explicitly provided one.
- Check the Visual Studio Output window and any application error messages.

</details>

## Known Setup Notes

- The connection configuration in `DBConnection.vb` currently uses a local MySQL server, database `ghc_petmart_3nf`, user `root`, and a blank password. Change these values to match your own environment.
- The database schema/backup was not present in the source archive reviewed when preparing this README. A working database backup is therefore a required setup dependency.
- The project targets **.NET Framework 4.7.2**, not modern .NET (such as .NET 8).
- The code references MySQL Connector/NET and PDF/reporting dependencies; restore/install compatible dependencies before building.
- Do not publish private database credentials, customer data, or real account passwords in source control.

## Contributors

- **Agustin, Jeshaiah D.**
- **Sabulao, John Michael P.**

## Academic Information

**Program:** Bachelor of Science in Information Technology (BSIT – Full IT)  
**Year Level:** First Year  
**Semester:** Third Semester  
**Institution:** Technological University of the Philippines – Taguig

## License

No license file was identified in the source archive reviewed for this README. Unless the repository owner adds a license, reuse and redistribution permissions are not explicitly granted. Add a `LICENSE` file if you want to specify the terms under which others may use this project.

---

<p align="center">
  <strong>GHC Pet Mart</strong><br>
  <em>Academic project — BSIT, First Year, Third Semester</em>
</p>
