# Kobo Annotation Exporter

An application to export annotations from your Kobo e-reader to Joplin.

## Features

- Automatic detection of connected Kobo devices
- Display of books and their annotations
- Support for different types of annotations:
  - Text annotations (highlights and notes)
  - Markup annotations (handwritten markings)
- Preview of markup annotations
- Direct export to Joplin
- Option to save markup annotations as images
- Automatic configuration on first use

## Requirements

- Python 3.6 or higher
- Joplin desktop application with Web Clipper enabled
- A Kobo e-reader

## Installation

1. Download the latest release of the application
2. Install the required Python packages:
   ```
   pip install -r requirements.txt
   ```

## Configuration

On first use, the application will show a configuration window. You need to fill in the following details:

1. **Joplin API Token**
   - Open Joplin
   - Go to Tools > Options > Web Clipper
   - Enable the Web Clipper
   - Copy the authorization token

2. **Notebook ID**
   - Right-click on the notebook in Joplin where you want to store your annotations
   - Select "Copy notebook ID"

3. **Web Clipper URL and Port**
   - Default is set to http://localhost:41184
   - These settings can be adjusted in the Joplin Web Clipper options

## Usage

1. Start the application
2. Connect your Kobo e-reader to your computer
3. Select your Kobo device from the dropdown list
4. Choose a book from the list to view its annotations
5. Select the annotations you want to export
6. Click "Export to Joplin" to export text annotations
7. For markup annotations:
   - Click "Preview Image" to view the markup
   - Choose "Export to Joplin" to export the markup to Joplin
   - Or choose "Save Image" to save the markup as an image

## Notes

- The application checks every 5 seconds for connected or disconnected Kobo devices
- Markup annotations are only supported for books in KEPUB format
- Text annotations are exported as markdown in Joplin
- Markup annotations are exported as images in Joplin

## Troubleshooting

If you encounter issues:

1. Check if Joplin is running and the Web Clipper is enabled
2. Verify that your API token is correct
3. Ensure your Kobo device is properly connected
4. Check if your notebook ID is correct

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request. 