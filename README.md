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
- Customizable highlight colors for text annotations
- Customizable annotation template for Joplin notes

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

### Customization Files

The application uses two additional configuration files for customization:

1. **highlight_colors.json**
   - Defines the colors for different highlight types
   - Format:
     ```json
     {
       "0": { "background": "#FFFFFF", "foreground": "#000000" },
       "1": { "background": "#FFFF00", "foreground": "#000000" }
     }
     ```
   - The number keys correspond to the highlight color index in Kobo
   - Each color has a background and foreground (text) color

2. **annotation_template.md**
   - Defines the template for how annotations are formatted in Joplin
   - Available placeholders:
     - `%chapter_title%`: Chapter title
     - `%anno_date%`: Annotation date
     - `%anno_time%`: Annotation time
     - `%anno_type%`: Annotation type
     - `%highlight_background%`: Background color
     - `%highlight_foreground%`: Text color
     - `%anno_text%`: The annotation text

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
- Text annotations are exported as markdown in Joplin using the configured template
- Markup annotations are exported as images in Joplin
- Highlight colors are preserved in the exported annotations
- Multiple annotations from the same book are combined into a single Joplin note

## Troubleshooting

If you encounter issues:

1. Check if Joplin is running and the Web Clipper is enabled
2. Verify that your API token is correct
3. Ensure your Kobo device is properly connected
4. Check if your notebook ID is correct
5. Verify that highlight_colors.json and annotation_template.md are properly formatted

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request. 