# DocXFilter v3.0

[![Python](https://img.shields.io/badge/Python-3.8%2B-blue.svg)](https://python.org)
[![Streamlit](https://img.shields.io/badge/Streamlit-1.28%2B-red.svg)](https://streamlit.io)
[![License](https://img.shields.io/badge/License-Proprietary-orange.svg)]()

> **Advanced multi-pattern search and analytics tool for DocXScan Excel outputs**

DocXFilter is a professional-grade document analysis application that enables powerful search capabilities across large document collections. Built specifically for processing DocXScan Excel outputs, it provides multi-pattern search, token management, and comprehensive analytics.

[[DocXFilter](https://docxfilter.streamlit.app/)](https://docxfilter.streamlit.app/)

## ✨ Key Features

### 🔍 **Advanced Search Engine**
- **Multi-pattern search** with AND/OR logic
- **Real-time search** with instant results feedback
- **Context highlighting** with multiple color coding
- **Regex pattern discovery** for automatic token detection

### 🏷️ **Token Management System**
- **JSON token definitions** import and management
- **Individual token selection** with visual feedback
- **Bulk operations** for efficient workflow
- **Search and filter** tokens by name or description

### 📊 **Analytics & Insights**
- **Document collection overview** with size and content metrics
- **Pattern distribution analysis** with interactive charts
- **Content quality assessment** and statistics
- **Auto-discovery** of document patterns

### 📤 **Professional Export**
- **Excel export** with detailed match counts
- **Summary sheets** with search metadata
- **Timestamped reports** for documentation
- **Match statistics** per document

## 🚀 Quick Start

### Prerequisites
```bash
Python 3.8+
pip install streamlit pandas plotly openpyxl
```

### Installation
1. **Clone or download** the DocXFilter script
2. **Install dependencies:**
   ```bash
   pip install streamlit pandas plotly openpyxl
   ```
3. **Run the application:**
   ```bash
   streamlit run docxfilter.py
   ```

### Basic Usage
1. **Upload Excel File** - Load your DocXScan Excel output
2. **Upload Token Definitions** (Optional) - JSON file with predefined search patterns
3. **Add Search Terms** - Either manually or select from imported tokens
4. **Choose Search Mode** - AND (all terms) or OR (any term)
5. **View Results** - Instant feedback with document matches
6. **Export Data** - Generate Excel reports with detailed analytics

## 📋 Input File Formats

### Excel File Requirements
Your DocXScan Excel file must contain:
- **File Name** column - Document identifiers
- **Full Contents** column - Complete document text
- **Size (KB)** column (optional) - File size information

### Token Definitions JSON Format
```json
{
  "<<FileService.": "Fileservice",
  "</ff>": "Page Break",
  "<fontsize": "Font Size",
  "PROMTINTO(": "PROMTINTO",
  "{ATTY": "ESIGN",
  "<<Special.": "SPECIAL"
}
```

**Structure:**
- **Key:** Exact token to search for (preserves special characters)
- **Value:** Human-readable description/name

## 🎯 Search Capabilities

### Search Modes
- **AND Mode:** Documents must contain ALL search terms
- **OR Mode:** Documents must contain ANY search term

### Pattern Discovery
DocXFilter automatically discovers common patterns:
- `<<text>>` - Angle bracket patterns
- `{TEXT}` - Curly brace patterns  
- `[TEXT]` - Square bracket patterns
- `<tag>` - HTML-style tags

### Context Highlighting
- **Multiple colors** for different search terms
- **Context preview** showing surrounding text
- **Full document highlighting** with scrollable view

## 📊 Analytics Features

### Document Overview
- Total document count and size metrics
- Content length distribution charts
- Document statistics and quality metrics
- Content density analysis

### Pattern Insights
- Imported token performance tracking
- Auto-discovered pattern analysis
- Interactive scatter plots for pattern distribution
- Document occurrence statistics

## 🎨 User Interface

### Professional Design
- **Clean, modern interface** with professional color scheme
- **Clear navigation** with highlighted active tabs
- **Instant feedback** for all user actions
- **Responsive design** for different screen sizes

### Key Sections
1. **Search Tab** - Main search interface with token management
2. **Analytics Tab** - Document and pattern insights
3. **Export Tab** - Report generation and download

## 📤 Export Capabilities

### Excel Reports Include
- **Search Results Sheet** - Matched documents with statistics
- **Summary Sheet** - Search metadata and configuration
- **Match Counts** - Per-token occurrence statistics
- **Export Metadata** - Timestamp and search parameters

### Export Features
- **Timestamped filenames** for organization
- **Multiple sheets** for comprehensive data
- **Match summaries** for quick analysis
- **Professional formatting** for reports

## 🔧 Technical Details

### Architecture
- **Streamlit-based** web application
- **Pandas** for data processing and analysis
- **Plotly** for interactive visualizations
- **OpenpyXL** for Excel file handling
- **Regex** for pattern matching and discovery

### Performance Features
- **Cached operations** for improved speed
- **Vectorized search** for large datasets
- **Optimized regex** compilation
- **Efficient memory usage** for large files

### Browser Compatibility
- Chrome, Firefox, Safari, Edge
- Responsive design for desktop and tablet
- No mobile optimization (desktop application)

## 🛠️ Advanced Usage

### Custom Token Files
Create JSON files with your organization's specific tokens:
```json
{
  "<<merge": "Mail merge field",
  "{CLIENT_NAME}": "Client name variable",
  "SIGNATURE_BLOCK": "Signature placeholder"
}
```

### Bulk Operations
- **Add All Visible** - Add all filtered tokens at once
- **Clear Selected** - Remove multiple tokens efficiently
- **Search and Filter** - Find specific tokens quickly

### Search Strategies
- Start with **broad terms** for general document filtering
- Use **specific tokens** for precise pattern matching
- Combine **AND/OR modes** for different search approaches
- Leverage **auto-discovered patterns** for comprehensive coverage

## 📝 Best Practices

### File Organization
- Use **descriptive filenames** for Excel exports
- Organize **token definition files** by project or department
- Maintain **consistent naming** conventions

### Search Optimization
- Test **individual tokens** before bulk operations
- Use **AND mode** for precise filtering
- Use **OR mode** for broad discovery
- Review **context previews** before full document analysis

### Performance Tips
- **Filter large datasets** before complex searches
- Use **specific terms** rather than very broad patterns
- **Clear search terms** between different analyses
- **Export results** for offline analysis

## 🐛 Troubleshooting

### Common Issues

**No documents found:**
- Verify Excel file has correct column names
- Check search terms for typos
- Try OR mode instead of AND mode

**Token file not loading:**
- Ensure JSON format is valid
- Check file encoding (UTF-8 recommended)
- Verify file extension is .json

**Performance issues:**
- Clear browser cache
- Restart Streamlit application
- Check file sizes (large files may be slow)

**Export problems:**
- Ensure sufficient disk space
- Check browser download permissions
- Try different export filename

## 📄 License

Copyright 2025 Hrishik Kunduru. All rights reserved.

This software is proprietary and confidential. Unauthorized copying, distribution, or use is strictly prohibited.

## 👤 Author

**Hrishik Kunduru**
- Professional document analysis solutions
- Advanced search and analytics tools

## 🔄 Version History

### v2.1 (Current)
- Professional UI redesign with improved navigation
- Enhanced token management with individual selection
- Instant search results feedback
- Optimized performance and code structure
- Improved export functionality

### v2.0
- Multi-pattern search with AND/OR logic
- Token import and management system
- Comprehensive analytics dashboard
- Auto-pattern discovery
- Enhanced export capabilities

### v1.0
- Basic document search functionality
- Excel file processing
- Simple pattern matching

---

For support or feature requests, please contact the developer.
Hrishik Kunduru
hkunduru@raslg.com
[https://hrishik](https://hrishik.netlify.app/)
