# Wallen Invoice Manager

A Tkinter-based invoice management application for automating invoice creation and field management with PDF templates.

## Features

- **Invoice Creation**: Create new invoices from fillable PDF templates
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

### Workflow

1. **Create New Invoice**: Select address, click "Create New Invoice"
2. **Edit Existing Invoice**: Select invoice from list, click "Edit Existing Invoice"
3. **PDF Editor**: Make changes in the PDF editor window
4. **Auto-Save**: Invoice automatically returns to manager when PDF closes

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
