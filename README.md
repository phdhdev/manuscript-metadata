# Figma Code Generator - Word Add-in

[![GitHub](https://img.shields.io/badge/GitHub-phdhdev%2Ffigma--code--generator-blue)](https://github.com/phdhdev/figma-code-generator)

A Microsoft Word Add-in that generates unique 6-digit Figma codes in the format `XXX-XXX` and inserts them at the cursor position while preserving existing text formatting.

## Features

- **Generate Unique Codes**: Creates random 6-digit codes in format `123-456`
- **Document-Wide Uniqueness Check**: Scans the entire document to ensure no duplicate codes
- **Format Preservation**: Inserts codes using the existing font, font size, and color at cursor position
- **Real-Time Statistics**: Shows total codes in document and codes generated in current session
- **User-Friendly Interface**: Clean, modern design with clear status messages

## How It Works

1. Click "Generate Unique Code" to create a new code
2. The add-in scans your entire document for existing codes (format: XXX-XXX)
3. Generates a unique code that doesn't exist in the document
4. Place your cursor anywhere in the document
5. Click "Insert at Cursor" to add the code with existing formatting

## Installation & Setup

### Prerequisites

- Microsoft Word (Desktop version)
- Node.js (v14 or higher)
- npm

### Step 1: Install Dependencies

```bash
npm install -g office-addin-dev-server
```

### Step 2: Create Project Structure

Create a folder for your add-in and place all the files in it:

```
figma-code-generator/
├── manifest.xml
├── taskpane.html
├── taskpane.js
├── commands.html
├── package.json
└── assets/ (optional - for icons)
```

### Step 3: Create Simple Icons (Optional)

You can create placeholder icon files or use simple images. Place them in an `assets` folder:
- `icon-16.png` (16x16 pixels)
- `icon-32.png` (32x32 pixels)
- `icon-64.png` (64x64 pixels)
- `icon-80.png` (80x80 pixels)

### Step 4: Update Manifest

Edit `manifest.xml` and change the `<Id>` to a unique GUID. You can generate one at https://guidgenerator.com/

```xml
<Id>YOUR-UNIQUE-GUID-HERE</Id>
```

### Step 5: Start the Development Server

Navigate to your project folder and run:

```bash
npm start
```

This will start a local server at `https://localhost:3000`

### Step 6: Sideload the Add-in in Word

#### For Windows:

1. Open Word
2. Go to **Insert** tab → **Get Add-ins** → **My Add-ins** → **Shared Folder**
3. Browse to your project folder and select `manifest.xml`

#### For Mac:

1. Open Word
2. Go to **Insert** tab → **Add-ins** → **My Add-ins**
3. Click "Upload Add-in" and select `manifest.xml`

#### Alternative Method (Network Share):

1. Create a network share folder
2. Copy `manifest.xml` to the network share
3. Add the network share path to Word's trusted add-in catalogs
4. Restart Word and load the add-in from **My Add-ins**

## Usage

### Generating a Code

1. Click the "Generate Unique Code" button in the task pane
2. The add-in will:
   - Scan your entire document for existing codes
   - Generate a random 6-digit code (XXX-XXX format)
   - Ensure the code is unique (not already in the document)
   - Display the code in the panel

### Inserting a Code

1. Place your cursor where you want the code inserted
2. Click "Insert at Cursor"
3. The code will be inserted with the same formatting as the surrounding text:
   - Font family (e.g., Arial, Times New Roman)
   - Font size (e.g., 12pt, 14pt)
   - Font color (e.g., black, blue, red)

### Statistics

The add-in displays:
- **Codes in Doc**: Total number of codes currently in the document
- **Generated**: Number of codes you've generated in this session

## Technical Details

### Code Format

- Format: `XXX-XXX` (e.g., `123-456`, `789-012`)
- Each part is a 3-digit number (100-999)
- Total possible combinations: 810,000 unique codes

### Uniqueness Check

The add-in uses Word's built-in search functionality with wildcards:
- Pattern: `\d{3}-\d{3}` (matches any 6-digit code with hyphen)
- Searches the entire document body
- Generates new codes until a unique one is found

### Format Preservation

When inserting code, the add-in:
1. Reads the current selection's font properties
2. Inserts the code as text
3. Applies the original formatting to the new text

## Troubleshooting

### Add-in doesn't appear in Word

- Ensure the development server is running (`npm start`)
- Check that manifest.xml is properly loaded
- Try restarting Word

### "Unable to generate unique code" error

- This occurs if 1000 attempts fail to find a unique code
- Usually means your document has many codes already
- Try using a different code format or clearing old codes

### Code doesn't preserve formatting

- Make sure you have text selected or cursor in formatted text
- The add-in copies formatting from the cursor position
- If at the start of document, default formatting is used

### Certificate/SSL errors

- The development server uses self-signed certificates
- You may need to trust the certificate in your system
- Follow prompts when starting the server

## Customization

### Change Code Format

Edit the `generateCode()` function in `taskpane.js`:

```javascript
function generateCode() {
    // Current: XXX-XXX
    const part1 = Math.floor(Math.random() * 900) + 100;
    const part2 = Math.floor(Math.random() * 900) + 100;
    return `${part1}-${part2}`;
    
    // Alternative: XXXX (4 digits, no hyphen)
    // return String(Math.floor(Math.random() * 9000) + 1000);
}
```

And update the search pattern in `getAllCodesInDocument()`:

```javascript
// Current pattern for XXX-XXX
const searchResults = body.search("\\d{3}-\\d{3}", { matchWildcards: true });

// Pattern for XXXX (4 digits)
// const searchResults = body.search("\\d{4}", { matchWildcards: true });
```

### Change Colors/Styling

Edit the `<style>` section in `taskpane.html` to customize:
- Button colors
- Header gradient
- Font sizes
- Border radius

### Add Prefix/Suffix

Modify the code generation or insertion:

```javascript
// In insertCode() function
const codeWithPrefix = "CODE-" + currentCode;
const insertedRange = selection.insertText(codeWithPrefix, Word.InsertLocation.replace);
```

## Development

### Project Structure

- **manifest.xml**: Add-in configuration and metadata
- **taskpane.html**: User interface layout and styling
- **taskpane.js**: Core functionality (generation, checking, insertion)
- **commands.html**: Required by Office.js framework
- **package.json**: Node.js project configuration

### Key Functions

- `generateCode()`: Creates random XXX-XXX format
- `getAllCodesInDocument()`: Searches document for existing codes
- `generateUniqueCode()`: Generates code not in document
- `insertCode()`: Inserts code with format preservation
- `updateDocumentStats()`: Updates code count display

## Browser Support

The add-in works in Word on:
- Windows Desktop
- Mac Desktop
- Word Online (with limitations on sideloading)

## License

MIT License - feel free to modify and use as needed!

## Support

For issues with:
- **Office Add-ins**: https://docs.microsoft.com/office/dev/add-ins/
- **Word JavaScript API**: https://docs.microsoft.com/javascript/api/word
- **This Add-in**: https://github.com/phdhdev/figma-code-generator/issues

## Version History

- **1.0.0** (2025): Initial release
  - Basic code generation
  - Document-wide uniqueness checking
  - Format preservation
  - Statistics display
