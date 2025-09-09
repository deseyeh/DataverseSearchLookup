# Dataverse Search Lookup PCF Control

A modern Power Platform Component Framework (PCF) control that provides intelligent search functionality for Dataverse records using the latest search APIs.

![Dataverse Search Lookup Demo](DataverseSearchLookup/assets/dataversesearch.gif)

## 🚀 Features

- **Modern Search API**: Uses Dataverse's latest `searchquery` endpoint for fast, relevant results
- **Intelligent Token Replacement**: Dynamically replaces field tokens in descriptions with actual record data
- **Fluent UI Design**: Built with Microsoft's Fluent UI React components for consistent styling
- **Flexible Card Appearance**: Multiple visual styles (filled, subtle, outline, filled-alternative)
- **Auto-Detection**: Automatically detects primary name fields for different entity types
- **Type-Safe**: Full TypeScript implementation with proper type definitions
- **Responsive**: Works seamlessly across different form factors

## 📦 Installation

1. Download the latest solution from releases
2. Import the solution into your Power Platform environment
3. Publish customizations

## 🔧 Configuration

### Properties

| Property            | Type            | Required | Description                                                     |
| ------------------- | --------------- | -------- | --------------------------------------------------------------- |
| **lookupField**     | Lookup.Simple   | ✅       | The lookup field to bind the control to                         |
| **searchParameter** | SingleLine.Text | ✅       | The search term/parameter to find records                       |
| **description**     | Text            | ✅       | Description text that supports token replacement                |
| **appearance**      | Enum            | ❌       | Card visual style (filled, subtle, outline, filled-alternative) |
| **searchCount**     | Text            | ❌       | Output property showing search result count                     |

The control automatically detects primary name fields for entities

## 💡 Usage Examples

### Basic Configuration

1. Add the control to a lookup field on your form
2. Set the **searchParameter** to a field that contains search terms
3. Configure the **description** with {{}} wrapped token placeholders

### Token Replacement

Use tokens in the description to dynamically show record information:

```
Found match: {{account.name}} - Created: {{account.createdon}}
Link to record: {{contact.account.name}}
```

The control will replace these tokens with actual field values from the current record.

### Search Integration

The control searches for records based on the `searchParameter` value and automatically:

- Finds the best matching record using Dataverse search
- Creates clickable links to the found records
- Updates the description with real-time data

## 🛠️ Technical Details

### Architecture

- **React Components**: Built with modern React hooks and functional components
- **Fluent UI**: Uses Microsoft's design system for consistent UX
- **TypeScript**: Fully typed for better development experience
- **Modern APIs**: Leverages latest Dataverse search capabilities

## 🔍 How It Works

1. **Search Trigger**: When the search parameter changes, the control initiates a search
2. **API Call**: Uses Dataverse's `searchquery` endpoint for intelligent search
3. **Result Processing**: Extracts the best matching record with relevance scoring
4. **Token Replacement**: Replaces description tokens with actual field values
5. **UI Rendering**: Displays results in a Fluent UI card with actions
6. **User Interaction**: Users can confirm selection to populate the lookup field

## 🎨 Customization

### Card Appearances

- **filled**: Solid background (default)
- **subtle**: Light background
- **outline**: Border only
- **filled-alternative**: Alternative solid style

### Styling

The control uses Fluent UI's theming system and automatically adapts to:

- Light/dark themes
- Platform color schemes
- Responsive breakpoints

## 🚀 Performance

- **Efficient Search**: Only searches when term length ≥ 2 characters
- **Caching**: Leverages browser and API caching mechanisms
- **Minimal Renders**: Optimized React rendering with proper state management
- **Lazy Loading**: Components load only when needed

## 🔧 Development

### Prerequisites

- Node.js 16+
- Power Platform CLI
- TypeScript
- React development knowledge

### Build Commands

```bash
# Install dependencies
npm install

# Build the control
npm run build



## 🤝 Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Add tests if applicable
5. Submit a pull request

## 📄 License

This project is licensed under the MIT License - see the LICENSE file for details.

## 🆘 Support

For issues, questions, or contributions:

- Create an issue in the repository
- Check existing documentation
- Review the example implementations

## 🔄 Version History

- Enhanced token replacement
- Improved error handling
- Better TypeScript support

---

**Made with ❤️ for the Power Platform community**
```
