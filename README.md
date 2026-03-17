# Gemini Addin Toolkit 🚀

A powerful, dark-themed AI assistant for Microsoft Excel, powered by **Google Gemini 2.5 Flash**. 

This toolkit allows you to interact with your spreadsheets using natural language. Whether you need to generate complex formulas, clean messy data, or create insightful charts, the Gemini Addin Toolkit brings state-of-the-art AI directly into your Excel workflow.

![App Screenshot](assets/icon-128.png)

## ✨ Features

- **Natural Language Data Entry**: "Add a row for a new employee named Sarah in HR with a salary of 75,000"
- **Formula Generation**: "Calculate the average of column B and put it in C1"
- **Data Cleaning**: One-click "Clean" to remove duplicates and trim whitespace
- **Smart Charts**: Generate Bar, Line, Pie, and Scatter charts by simply asking
- **Formula Explainer**: Understand complex legacy formulas with a single click
- **Selection Awareness**: Automatically understands context from your currently selected cells
- **Dark Mode UI**: A premium, glassmorphic interface designed for productivity

## 🛠️ Tech Stack

- **Frontend**: HTML5, Vanilla CSS, JavaScript (ES6+)
- **Build Tool**: Webpack & Babel
- **Integration**: Office.js API
- **AI Brain**: Google Gemini 1.5/2.5 Flash via REST API

## 🚀 Getting Started

### 1. Prerequisites
- [Node.js](https://nodejs.org/) installed
- [Microsoft Excel](https://www.microsoft.com/en-us/microsoft-365/excel) (Desktop or Web)

### 2. Get a Gemini API Key
1. Visit [Google AI Studio](https://aistudio.google.com/)
2. Create a new API Key (Free Tier supported!)

### 3. Installation
```bash
# Clone the repository
git clone https://github.com/YOUR_USERNAME/Gemini-Addin-Toolkit.git
cd Gemini-Addin-Toolkit

# Install dependencies
npm install
```

### 4. Running the Toolkit
```bash
# Start the development server
npm start
```
This will launch the Webpack Dev Server and automatically open Excel with the add-in sideloaded.

## ⚙️ Development

If you need to check which Gemini models are available for your specific API key, run:
```bash
node check_models.js YOUR_API_KEY
```

## 📄 License
MIT License - feel free to use and modify for your own projects!
