InvoiceGenerator (VB.NET WinForms)

Description
-----------
InvoiceGenerator is a lightweight Windows Forms application written in VB.NET that creates, previews, prints and stores invoices (proforma/tax/quotation). It supports a trial/licensing workflow, local invoice persistence using an Access database (in AppData), printing via PrintPreview, and a simple client identifier for license binding.

Key Features
------------
- Create and edit invoice headers (Client, Address, Date, Invoice Type).
- Add/remove invoice line items with automatic calculations (unit price, qty, total).
- Print preview and print invoices (custom PrintDocument rendering).
- Persist invoices to a local Access database (invoices.accdb) located in the application AppData folder.
- Trial mode with days-left UI and license installation (license.lic) bound to a local Client ID.
- Auto-generated invoice serial (persisted counter + GUID) and client ID (persisted GUID in AppData).

Prerequisites
-------------
- Windows OS
- Visual Studio (2017/2019/2022) or MSBuild-compatible environment
- .NET Framework compatible with the project (open the ClassProject.vbproj to confirm target framework)
- Microsoft Access Database Engine (ACE) for OleDb (Provider=Microsoft.ACE.OLEDB.12.0). Install "Microsoft Access Database Engine 2010/2016 Redistributable" if not present.

Important File Locations
------------------------
- App data folder used by the app: %APPDATA%\InvoiceGenerator
  - license.lic          -> license file (text: payload + signature)
  - clientid.txt         -> generated client identifier (GUID)
  - invoices.accdb       -> Access database created/used by the app
  - serial_counter.txt   -> persisted invoice serial counter

Build and Run
-------------
1. Open the solution in Visual Studio: ClassProject\ClassProject.vbproj (or open the solution file if present).
2. Restore NuGet packages if any and ensure project references are resolved.
3. Build the project (Build -> Build Solution).
4. Run (F5) to start the WinForms application.

Notes on Licensing and Trial
---------------------------
- The application supports a built-in free trial. When the trial expires the UI is disabled until a valid license is installed.
- A valid license file (license.lic) must be installed to enable saving and editing invoices and to re-enable the UI after trial expiry.
- License files are expected to be an RSA-signed payload (see LicenseGenerator.vb for structure and signing tools included in the repo).
- Use the "Load License" button or place license.lic into the AppData folder to install a license.

Database and Compatibility
--------------------------
- The app uses Microsoft.ACE.OLEDB.12.0 to create and access an Access database. If you cannot create/read invoices, verify the Access Database Engine is installed and the provider is available on the running machine.
- The project attempts to create tables (Invoices, InvoiceItems) automatically on first run.

Troubleshooting
---------------
- If PrintPreview or printing does not work, confirm default printer drivers are installed and supported.
- If the app cannot create the Access DB, ensure the process has write permissions to %APPDATA% and the ACE OLEDB provider is installed.
- If license validation fails, verify the license file format and that the Client ID shown in the header matches the licensed client ID.

Contributing
------------
- Fork the repo, create a feature branch, and submit pull requests.
- Keep UI changes minimal and prefer background tasks for long-running operations (database, file IO).

License
-------
Project licensing is not included in this README. Check repository root for a LICENSE file or contact the project owner.

Contact
-------
For issues and improvements open an issue on the repository or contact the project owner.
