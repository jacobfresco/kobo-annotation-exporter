# Kobo Annotation Exporter

A Python application that exports annotations from Kobo e-readers to Joplin notes. This tool allows you to easily transfer your book highlights, notes, and markup annotations from your Kobo device to Joplin for better organization and access.

## Stable release
This branch is the repo for [release 0.2.2](https://github.com/jacobfresco/kobo-annotation-exporter/releases/tag/0.2.2). Branch is protected.

## Features

- Export highlights and notes from Kobo e-readers to Joplin
- Support for markup annotations (SVG and JPG files)
- Automatic detection of Kobo devices
- Easy-to-use two-panel interface:
  - Top panel: List of books with annotation counts
  - Bottom panel: Annotations for the selected book
- Proper chapter title handling for annotations
- Configurable Joplin notebook destination
- Chronological ordering of annotations within notes

## Note Formatting

Annotations are exported with the following format:
- Chapter titles: `## Chapter Title`
- Timestamps: `### Timestamp`
- Regular annotations: Wrapped in code blocks
- Markup annotations: Images with associated text
- Annotations are separated by `---`

## Requirements

- Python 3.7 or higher
- Joplin desktop application
- Kobo e-reader device
- Windows operating system (for device detection)

## Installation

### Option 1: Using the executable (recommended)
1. Download the latest release from the releases page
2. Extract the zip file
3. Run `Kobo Annotation Exporter.exe`

### Option 2: From source
1. Clone this repository:
```bash
git clone https://github.com/jacobfresco/kobo-annotation-exporter.git
cd kobo-annotation-exporter
```

2. Install the required packages:
```bash
pip install -r requirements.txt
```

3. Configure Joplin:
   - Open Joplin
   - Enable the Web Clipper service
   - Copy your API token from the Web Clipper options
   - Create a notebook where you want to store your annotations

4. Configure the application:
   - Copy `config.json.default` to `config.json`
   - Edit `config.json` with your settings:
     - `joplin_api_token`: Your Joplin API token
     - `notebook_id`: The ID of the notebook where annotations will be stored
     - `web_clipper`: Configuration for the Joplin Web Clipper service
       - `url`: The base URL (usually "http://localhost")
       - `port`: The port number (usually 41184)

## Building the Executable

If you want to build the executable yourself:

1. Install the required packages:
```bash
pip install -r requirements.txt
```

2. Build the executable:
```bash
pyinstaller kobo_annotation_exporter.spec
```

The executable will be created in the `dist` directory.

## Usage

1. Connect your Kobo device to your computer
2. Run the application
3. Select your Kobo device from the dropdown menu
4. In the top panel, select a book to view its annotations
5. In the bottom panel, select the annotations you want to export
6. Click "Export to Joplin"

## Configuration

The application stores its configuration in a `config.json` file. You can manually edit this file or use the Settings dialog in the application.

Required settings:
- `joplin_api_token`: Your Joplin API token
- `notebook_id`: The ID of the notebook where annotations will be stored
- `web_clipper`: Configuration for the Joplin Web Clipper service
  - `url`: The base URL (usually "http://localhost")
  - `port`: The port number (usually 41184)

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request. 
