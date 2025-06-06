# Kobo Annotation Exporter

A tool to export annotations from Kobo e-readers to Joplin and other formats.

## Features

- Export annotations from Kobo e-readers
- Export to Joplin with proper formatting and colors
- Export to Markdown files with customizable templates
- Support for both regular annotations and markup annotations
- Automatic device detection
- Configurable export settings

## Installation

### From Source

1. Clone this repository
2. Install the required dependencies:
   ```bash
   pip install -r requirements.txt
   ```
3. Run the application:
   ```bash
   python export_annotations.py
   ```

### From Executable

1. Download the latest release from the releases page
2. Extract the ZIP file
3. Run `kae.exe`

## Usage

1. Connect your Kobo e-reader to your computer
2. Launch the application
3. Select the book you want to export annotations from
4. Choose your export format:
   - Joplin: Exports directly to your Joplin notebook
   - Markdown: Exports to a markdown file using the template

### Export Formats

#### Joplin Export
- Requires Joplin to be running
- Exports annotations with proper formatting and colors
- Supports both regular and markup annotations
- Automatically organizes annotations by book

#### Markdown Export
- Exports annotations to a markdown file
- Uses customizable templates
- Supports regular annotations only (markup annotations are not supported)
- Includes book metadata and annotation details
- Preserves highlight colors

## Configuration

The application uses a `config.json` file to store settings. You can configure:
- Joplin API token
- Joplin server URL
- Joplin port
- Enable/disable Joplin export

## Requirements

- Python 3.8 or higher
- Kobo e-reader
- Joplin (for Joplin export)
- Required Python packages (see requirements.txt)

## Development

### Building from Source

1. Install PyInstaller:
   ```bash
   pip install pyinstaller
   ```
2. Build the executable:
   ```bash
   pyinstaller --onefile --windowed --name kae export_annotations.py
   ```

## Version History

### 0.5.1 "Erasmus" (pre)
- Added markdown export functionality
- Added support for customizable export templates
- Improved error handling
- Fixed various bugs

### 0.5.0 "Erasmus"
- Initial release
- Basic Joplin export functionality
- Device detection
- Configuration management

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request. 