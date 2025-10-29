<div align="center">

# Claude for Excel

**AI-powered Excel assistant powered by Anthropic's Claude**

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![TypeScript](https://img.shields.io/badge/TypeScript-5.2-blue?logo=typescript)](https://www.typescriptlang.org/)
[![React](https://img.shields.io/badge/React-18-blue?logo=react)](https://reactjs.org/)
[![Vite](https://img.shields.io/badge/Vite-5-646CFF?logo=vite)](https://vitejs.dev/)
[![Anthropic Claude](https://img.shields.io/badge/Claude-3.5%20Sonnet-orange)](https://www.anthropic.com/)

[Features](#-features) • [Installation](#-installation) • [Usage](#-usage) • [Security](#-security--privacy) • [Contributing](#-contributing)

---

</div>

An AI-powered Excel add-in that brings the power of Claude AI directly into your spreadsheets. Analyze data, generate formulas, format tables, and get intelligent assistance without leaving Excel.

<div align="center">

**[📸 Demo Screenshot Coming Soon]**

*Chat with Claude directly in Excel's sidebar • Context-aware assistance • Instant Excel tools*

</div>

## ✨ Features

- 🤖 **Natural Language Interface** - Ask questions about your data in plain English
- 📊 **Smart Data Analysis** - Automatic analysis of selected cells and ranges
- ⚡ **Quick Tools Menu** - 17+ built-in Excel tools for common operations
- 🎨 **Professional UI** - Clean, modern interface matching Excel's design
- 🔒 **Privacy First** - Your API key stays local, data only sent to Anthropic
- 📝 **Command Palette** - Quick access to 18 pre-built prompts with `/` commands
- 🖼️ **Image Support** - Attach images and PDFs for analysis (up to 10 files)
- 💾 **Context Aware** - Automatically includes selected Excel data in conversations
- ⌨️ **Keyboard Shortcuts** - Fast navigation and actions
- 🛠️ **Excel Tools** - Format tables, freeze panes, sort data, insert rows/columns, and more

## 🚀 Quick Start

| Workflow | Steps |
|----------|-------|
| **Analyze Data** | 1. Select cells in Excel<br>2. Type `/analyze` or ask "What patterns do you see?"<br>3. Get AI-powered insights |
| **Generate Formula** | 1. Select your data range<br>2. Type `/formula` or describe what you need<br>3. Claude generates the formula |
| **Format Table** | 1. Select data range<br>2. Click **⋮** menu → "Format as Table"<br>3. Instant professional formatting |
| **Clean Data** | 1. Select messy data<br>2. Type `/clean` or describe issues<br>3. Get cleanup suggestions |
| **Stop Generation** | Press the **Stop** button (appears during response) to cancel |

## 📊 What You Can Do

<table>
<tr>
<td width="50%">

**📈 Data Analysis**
- Pattern recognition
- Statistical insights
- Trend identification
- Outlier detection

**🧮 Formula Generation**
- Complex calculations
- VLOOKUP/XLOOKUP
- Conditional formulas
- Array formulas

</td>
<td width="50%">

**🎨 Formatting & Styling**
- Professional table formatting
- Conditional formatting rules
- Cell styling suggestions
- Auto-fit columns/rows

**🔧 Data Operations**
- Clean and normalize data
- Find and remove duplicates
- Transpose data
- Sort and filter

</td>
</tr>
</table>

## 🔒 Security & Privacy

✅ **No hardcoded API keys** - You provide your own Anthropic API key
✅ **Local storage only** - API key stored in Office document settings (never sent to external servers)
✅ **No tracking** - No analytics, no telemetry, no data collection
✅ **Direct API calls** - Communication goes directly to Anthropic's API
✅ **Open source** - Full transparency, inspect the code yourself

**Important:** Your Anthropic API key is stored locally in Excel's document settings and is never transmitted to any server except Anthropic's official API (api.anthropic.com).

## 📋 Prerequisites

- **Excel Desktop** (Windows or macOS)
  - Windows: Excel 2016 or later
  - macOS: Excel 2016 or later
- **Node.js** 18+ and npm
- **Anthropic API Key** - Get one at [console.anthropic.com](https://console.anthropic.com/account/keys)

## 🚀 Installation

### 1. Clone the Repository

```bash
git clone https://github.com/heyimjames/excel-claude-sidebar.git
cd excel-claude-sidebar
```

### 2. Install Dependencies

```bash
npm install
```

### 3. Generate SSL Certificate

Office Add-ins require HTTPS in development. Generate a self-signed certificate:

```bash
npx office-addin-dev-certs install
```

This creates a certificate in the `.certs` folder and installs it in your system's trust store.

### 4. Start Development Server

```bash
npm run dev
```

This starts:
- Vite dev server at `https://localhost:3000`
- Hot module replacement (HMR) enabled
- Source maps for debugging

### 5. Sideload the Add-in

#### macOS
1. Open Excel
2. Go to **Insert** > **Add-ins** > **My Add-ins**
3. Choose **Add a Custom Add-in**
4. Select the `manifest.xml` file from this project
5. Click **OK**

#### Windows
1. Open Excel
2. Go to **Insert** > **Add-ins** > **My Add-ins**
3. Select **SHARED FOLDER**
4. Browse to the folder containing `manifest.xml`
5. Select the manifest and click **Insert**

### 6. Configure API Key

When you first open the add-in:
1. Click **Show Claude** in the Excel ribbon
2. Enter your Anthropic API key
3. Click **Save API Key**

Your key is saved locally in Excel's document settings.

## 💡 Usage

### Basic Chat

1. Select cells in Excel (optional - provides context)
2. Type your question in the input box
3. Press **Enter** to send

### Command Palette

Type `/` in the input to open the command palette with 18 pre-built prompts:

| Command | Description |
|---------|-------------|
| `/analyze` | Analyze selected data and provide insights |
| `/summarize` | Create a summary of the data |
| `/chart` | Get suggestions for chart types |
| `/format` | Apply formatting to cells |
| `/formula` | Generate Excel formulas |
| `/explain` | Explain patterns in the data |
| `/clean` | Clean up and normalize data |
| `/sort` | Sort data recommendations |
| `/filter` | Apply filtering suggestions |
| `/conditional` | Create conditional formatting rules |
| `/vlookup` | Generate VLOOKUP formulas |
| `/table` | Convert range to structured table |
| `/transpose` | Transpose rows and columns |
| `/duplicate` | Find and remove duplicates |
| `/calculate` | Perform calculations |
| `/validate` | Add data validation rules |
| `/merge` | Merge cells or data |
| `/pivot` | Create pivot table suggestions |

### Quick Tools Menu (⋮)

Click the **three-dot menu** in the header for instant Excel actions:

| Category | Tool | Description |
|----------|------|-------------|
| **Formatting** | AutoFit Columns & Rows | Auto-fit all columns and rows to content |
| | Format as Table | Convert selection to a table |
| | Clear Formatting | Remove all formatting from selection |
| **Insert/Delete** | Insert Row Above | Insert new row above selection |
| | Insert Column Left | Insert new column to the left |
| | Delete Selected Rows | Delete the selected rows |
| | Delete Selected Columns | Delete the selected columns |
| **Freeze Panes** ✓ | Freeze Top Row | Keep top row visible when scrolling |
| | Freeze First Column | Keep first column visible when scrolling |
| | Unfreeze Panes | Remove all freeze panes |
| **Sorting** | Sort Ascending | Sort selection A to Z |
| | Sort Descending | Sort selection Z to A |
| **Navigation** | Select All Data | Select all used cells in worksheet |
| | Go to Cell A1 | Jump to the top-left cell |

> ✓ = Toggleable tools show an orange checkmark when active

### Image Attachments

Attach images or PDFs for Claude to analyze:
- Click the **📎 paperclip icon** to upload files
- Or **drag and drop** files into the input area
- Or **paste** images from clipboard (Cmd/Ctrl + V)
- Supports: JPEG, PNG, GIF, WebP, PDF
- Max 10 files per message, 5MB per image, 10MB per PDF

### Keyboard Shortcuts

| Shortcut (Mac) | Shortcut (Windows) | Action |
|----------------|-------------------|--------|
| `⌘K` | `Ctrl+K` | Focus message input |
| `⌘L` | `Ctrl+L` | Clear chat history |
| `⇧?` | `Shift+?` | Show keyboard shortcuts |
| `Enter` | `Enter` | Send message |
| `Shift+Enter` | `Shift+Enter` | New line in message |
| `/` | `/` | Open command palette |
| `Esc` | `Esc` | Clear input or close palette |

## 🏗️ Project Structure

```
excel-claude-sidebar/
├── src/
│   └── taskpane/
│       ├── components/      # React components
│       │   ├── ChatInterface.tsx
│       │   ├── MessageInput.tsx
│       │   ├── Message.tsx
│       │   ├── Settings.tsx
│       │   ├── ToolsMenu.tsx
│       │   └── ...
│       ├── hooks/          # Custom React hooks
│       │   ├── useClaudeChat.ts
│       │   ├── useExcelContext.ts
│       │   └── ...
│       ├── lib/            # Utilities & types
│       │   ├── anthropic.ts
│       │   ├── commands.ts
│       │   └── types.ts
│       ├── styles/         # CSS files
│       └── App.tsx         # Main app component
├── assets/                 # Icons and images
├── manifest.xml           # Office Add-in manifest
├── vite.config.ts         # Vite configuration
└── package.json
```

## 🛠️ Development

### Available Scripts

```bash
npm run dev          # Start dev server with HMR
npm run build        # Build for production
npm run preview      # Preview production build
npm run lint         # Run ESLint
npm run type-check   # TypeScript type checking
npm run clean        # Remove dist folder
```

### Environment Variables

Create a `.env` file for development (optional):

```bash
# Development only - for testing without Office
VITE_ANTHROPIC_API_KEY=your_api_key_here
```

**Note:** In production, API keys are entered by users via the Settings UI and stored in Office document settings.

## 🔧 Troubleshooting

### Certificate Issues

If you see certificate warnings:

```bash
# Reinstall certificate
npx office-addin-dev-certs install --force
```

### Add-in Not Loading

1. Check dev server is running (`npm run dev`)
2. Verify HTTPS certificate is trusted
3. Clear Office cache:
   - macOS: `~/Library/Containers/com.microsoft.Excel/Data/Documents/wef`
   - Windows: `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef`
4. Restart Excel

### API Key Not Saving

API keys are stored in Office document settings. If not saving:
1. Make sure you have write permissions for the Excel file
2. Try saving the Excel file first
3. Re-enter the API key in Settings

### Context Not Loading

If Excel context isn't being captured:
1. Select cells before sending message
2. Check Console for errors (F12 in Excel)
3. Try refreshing the add-in

## 💰 API Usage & Costs

This add-in uses Anthropic's Claude API. You're responsible for your own API usage and costs:

- API calls are billed by Anthropic based on token usage
- See [Anthropic Pricing](https://www.anthropic.com/pricing) for current rates
- Monitor usage in your [Anthropic Console](https://console.anthropic.com)

## 🤝 Contributing

Contributions welcome! Please:

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Make your changes
4. Test thoroughly
5. Commit your changes (`git commit -m 'Add amazing feature'`)
6. Push to the branch (`git push origin feature/amazing-feature`)
7. Open a Pull Request

## 👨‍💻 Credits

**Built by**
- [James Frewin](https://twitter.com/jamesfrewin1)
  - [Twitter](https://twitter.com/jamesfrewin1)
  - [LinkedIn](https://linkedin.com/in/jamesfrewin)
  - [GitHub](https://github.com/heyimjames)

**Design by**
- [OCTOBER Design Studio](https://www.octoberwip.com)

## 📜 License

MIT License - see LICENSE file for details

## 🆘 Support

- 🐛 **Issues**: [GitHub Issues](https://github.com/heyimjames/excel-claude-sidebar/issues)
- 💬 **Questions**: [GitHub Discussions](https://github.com/heyimjames/excel-claude-sidebar/discussions)
- 🐦 **Updates**: Follow [@jamesfrewin1](https://twitter.com/jamesfrewin1)

## 🔗 Links

- [Anthropic Console](https://console.anthropic.com)
- [Office Add-ins Documentation](https://learn.microsoft.com/en-us/office/dev/add-ins/)
- [Claude API Documentation](https://docs.anthropic.com/)

---

**Disclaimer:** This is an unofficial third-party add-in. Not affiliated with Microsoft or Anthropic. Use at your own risk.
