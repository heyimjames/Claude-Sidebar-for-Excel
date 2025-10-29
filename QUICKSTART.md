# Excel Claude Sidebar - Quick Start Guide

## 🎉 Your AI-Powered Excel Assistant is Ready!

This project is a fully functional Excel add-in that integrates Claude AI to help you work with your spreadsheets through natural language.

## 🚀 Getting Started

### 1. Start the Development Server

```bash
npm run dev
```

This will start the Vite dev server at `https://localhost:3000` with HTTPS.

### 2. Load the Add-in in Excel

Open a new terminal and run:

```bash
npm start
```

This will:
- Open Excel (or use an existing instance)
- Sideload the add-in
- Show the "Claude AI" button in the Home ribbon

### 3. Configure Your API Key

1. Click the "Show Claude" button in Excel's Home ribbon
2. The sidebar will open
3. Enter your Anthropic API key (get one from https://console.anthropic.com)
4. Click "Get Started"

### 4. Start Chatting!

You can now interact with Claude to:
- **Read data**: "What's in cells A1 through A10?"
- **Analyze**: "Summarize the data in this sheet"
- **Edit**: "Fill column B with the squares of column A"
- **Format**: "Make the header row bold with a blue background"
- **Create charts**: "Create a line chart from the data in A1:B10"
- **Apply formulas**: "Add a SUM formula at the bottom of column C"

## 🛠️ Available Tools

Claude has access to these Excel operations:

### Data Operations
- `read_range` - Read cell values, formulas, and formats
- `write_range` - Write data to cells
- `get_selection` - Get currently selected cells
- `get_workbook_info` - View all worksheets

### Formatting & Structure
- `format_range` - Apply colors, fonts, number formats
- `create_table` - Create formatted tables
- `insert_rows` / `delete_rows` - Modify row structure
- `sort_range` - Sort data by columns

### Formulas & Charts
- `apply_formula` - Add Excel formulas
- `create_chart` - Generate visualizations

## 📁 Project Structure

```
Excel Claude Sidebar/
├── src/
│   ├── taskpane/
│   │   ├── components/
│   │   │   ├── ChatInterface.tsx    # Main chat UI
│   │   │   ├── Message.tsx          # Message bubbles
│   │   │   ├── MessageInput.tsx     # Input field
│   │   │   └── ApiKeySetup.tsx      # Setup screen
│   │   ├── hooks/
│   │   │   ├── useClaudeChat.ts     # Claude integration
│   │   │   └── useExcelTools.ts     # Excel operations
│   │   ├── lib/
│   │   │   ├── excel-tools.ts       # Tool definitions
│   │   │   └── types.ts             # TypeScript types
│   │   ├── styles/                  # CSS files
│   │   ├── App.tsx                  # Root component
│   │   └── index.tsx                # Entry point
│   └── commands.html                # Ribbon commands
├── manifest.xml                     # Add-in manifest
├── package.json
└── vite.config.ts
```

## 🎨 Features

### Beautiful Modern UI
- Clean, minimalist design
- Gradient purple theme
- Smooth animations
- Fluent UI components

### Smart AI Integration
- Powered by Claude 3.5 Sonnet
- Function calling for Excel operations
- Context-aware responses
- Error handling

### Secure & Private
- API key stored in Excel workbook settings
- Never shared or exposed
- All processing through Anthropic's secure API

## 🔧 Development Commands

- `npm run dev` - Start development server
- `npm run build` - Build for production
- `npm start` - Debug in Excel
- `npm stop` - Stop debugging
- `npm run validate` - Validate manifest

## 📝 Example Conversations

**User**: "What data do I have in this sheet?"

**Claude**: *Uses `get_selection` and `read_range` tools to analyze your data and provides a summary*

---

**User**: "Create a chart showing monthly sales"

**Claude**: *Uses `create_chart` tool to generate a visualization*

---

**User**: "Make the header row stand out"

**Claude**: *Uses `format_range` to apply bold text and background color*

## 🐛 Troubleshooting

### Build Issues
- Make sure all dependencies are installed: `npm install`
- Clear build cache: `rm -rf dist && npm run build`

### Add-in Not Loading
- Ensure dev server is running: `npm run dev`
- Check that Excel is accessing `https://localhost:3000`
- Trust the development certificates

### API Key Issues
- Verify your Anthropic API key is valid
- Check that it starts with `sk-ant-`
- Ensure you have credits in your Anthropic account

## 🚢 Production Deployment

1. Build the project: `npm run build`
2. Host the `dist/` folder on a web server with HTTPS
3. Update `manifest.xml` with your production URLs
4. Distribute the manifest to users

## 🎯 Next Steps

Consider adding:
- [ ] More advanced chart types
- [ ] PivotTable support
- [ ] Multi-sheet operations
- [ ] Export to different formats
- [ ] Custom Excel functions
- [ ] Conversation history persistence

## 📚 Resources

- [Office Add-ins Documentation](https://docs.microsoft.com/office/dev/add-ins)
- [Excel JavaScript API](https://docs.microsoft.com/javascript/api/excel)
- [Anthropic Claude Documentation](https://docs.anthropic.com)
- [Fluent UI React](https://react.fluentui.dev)

---

**Built with ❤️ using Claude Code**
