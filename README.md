# Wallen Invoice Manager

A Tkinter-based invoice management application for automating invoice creation and field management with PDF templates.

## Features

- **Invoice Creation**: Create new invoices from fillable PDF templates
- **Quick Notes**: Create plain-text customer notes with address/date naming
- **Field Management**: Automatically extract and manage PDF form fields
- **Bulk Processing**: Edit and organize multiple invoices efficiently
- **Logo Display**: Professional header with company logo
- **Status Tracking**: Real-time status updates during operations
- **File Organization**: Automatic invoice numbering and categorization

## Requirements

- Python 3.14+
- Tkinter (included with Python)
- pypdf 6.14.2
- Pillow 12.3.0

## Installation

1. Clone the repository:
```bash
git clone https://github.com/rgwallen/kelly.git
cd kelly
```

2. Create a virtual environment (optional but recommended):
```bash
python -m venv .venv
.venv\Scripts\activate  # On Windows
```

3. Install dependencies:
```bash
pip install pypdf pillow
```

## Usage

Run the application:
```bash
python wallen_invoice.py
```

### Email Notification Setup

The app can send an email to admin@wallenac.com whenever a new invoice is created.

If this app folder is synchronized via OneDrive, use this pattern:

1. Keep non-secret defaults in `smtp_settings.json` (shared).
2. Keep machine-specific values and passwords in `%APPDATA%\\WallenInvoice\\smtp_settings.local.json` (local only).
3. Use **Test Email Settings** on each machine after setup.

Preferred method for all systems:

1. Edit `smtp_settings.json` in the app folder.
2. Choose `email_provider`: `smtp`, `graph`, or `auto`.
3. Fill SMTP or Graph settings.
3. Use **Test Email Settings** inside the app.

`smtp_settings.json` fields:

```json
{
  "email_provider": "auto",
  "smtp_host": "smtp.yourprovider.com",
  "smtp_port": "587",
  "smtp_use_tls": "true",
  "email_from": "noreply@wallenac.com",
  "email_username": "smtp-user",
  "email_password": "smtp-password",
  "graph_tenant_id": "your-tenant-guid",
  "graph_client_id": "your-app-client-id",
  "graph_client_secret": "your-app-client-secret",
  "graph_scope": "https://graph.microsoft.com/.default"
}
```

Microsoft Graph mode requirements (`email_provider: graph`):

1. Azure App Registration with client secret.
2. Application permission `Mail.Send` granted with admin consent.
3. `email_from` mailbox exists and is allowed for app-based send.

Optional override method (advanced):

Environment variables can still override file settings:

```powershell
$env:WALLEN_SMTP_HOST="smtp.yourprovider.com"
$env:WALLEN_SMTP_PORT="587"
$env:WALLEN_SMTP_USE_TLS="true"
$env:WALLEN_EMAIL_FROM="noreply@wallenac.com"
$env:WALLEN_EMAIL_USERNAME="smtp-user"
$env:WALLEN_EMAIL_PASSWORD="smtp-password"
$env:WALLEN_EMAIL_PROVIDER="graph"
$env:WALLEN_GRAPH_TENANT_ID="your-tenant-guid"
$env:WALLEN_GRAPH_CLIENT_ID="your-app-client-id"
$env:WALLEN_GRAPH_CLIENT_SECRET="your-app-client-secret"
$env:WALLEN_GRAPH_SCOPE="https://graph.microsoft.com/.default"
```

Notes:
- If SMTP host is not set in either `smtp_settings.json` or `WALLEN_SMTP_HOST`, email notifications are skipped.
- Settings precedence is: environment variable -> local AppData file -> shared app folder file.
- In `auto` mode, Graph is used first when fully configured; otherwise SMTP is used.
- The notification recipient is fixed at `admin@wallenac.com` in the app.
- Use **Test Email Settings** in the app to run provider checks and send checks.
- Use **Copy Email Diagnostics** to copy the latest provider checks for support tickets or email.
- The status panel shows an SMTP health indicator that checks automatically at startup.

Local machine file example (`%APPDATA%\\WallenInvoice\\smtp_settings.local.json`):

```json
{
  "email_provider": "smtp",
  "smtp_host": "smtp.yourprovider.com",
  "smtp_port": "587",
  "smtp_use_tls": "true",
  "email_from": "noreply@wallenac.com",
  "email_username": "smtp-user",
  "email_password": "smtp-password",
  "graph_tenant_id": "",
  "graph_client_id": "",
  "graph_client_secret": "",
  "graph_scope": "https://graph.microsoft.com/.default"
}
```

### Workflow

1. **Create New Invoice**: Select address, click "Create New Invoice"
2. **Edit Existing Invoice**: Select invoice from list, click "Edit Existing Invoice"
3. **Create New Note**: Create and open a plain `.txt` note in the Invoices folder
4. **Test Email Settings**: Validate email provider from the current machine
5. **Copy Email Diagnostics**: Copy latest provider check details to clipboard
6. **PDF Editor**: Make changes in the PDF editor window
7. **Confirm and Save**: Close the PDF and confirm in the app dialog

## Project Structure

```
kelly/
├── wallen_invoice.py              # Main application
├── Invoice Template/              # PDF templates
├── Invoices/                      # Created invoices
├── Processed/                     # Completed invoices
├── Logs/                          # Application logs
└── Assets/                        # Company logo and images
```

## Future Enhancements

- Ollama ML integration for smart field population
- Advanced invoice analytics
- Email delivery integration
- Multi-template support

## License

MIT License - See LICENSE file for details

## Support

For issues or questions, please open an issue on GitHub.
