# Priority Based Budgeting Report Generator

A comprehensive web application for generating Priority Based Budgeting reports from Excel data files. This tool allows users to upload budget request data, apply filters, and generate professional reports for public budget documents.

## Features

- **Excel File Upload**: Support for .xlsx and .xls files with drag-and-drop functionality
- **Advanced Filtering**: Filter by Fund, Department, Division, Program, Request Type, and Status
- **Real-time Statistics**: Display key metrics including total requests and amounts
- **Visual Analytics**: Charts showing department breakdown and program alignment
- **Professional Reports**: Generate formatted reports with executive summaries
- **Word Export**: Download reports as Word documents for inclusion in budget documents
- **Responsive Design**: Works on desktop and mobile devices

## Data Structure

The application expects an Excel file with the following tabs:
- **Request Summary**: Main budget request data
- **Personnel**: Personnel-related budget items
- **NonPersonnel**: Non-personnel budget items  
- **Request Q&A**: Additional context and questions/answers
- **Budget Summary**: Detailed budget breakdowns

## Installation

1. Clone this repository:
```bash
git clone https://github.com/yourusername/priority-based-budgeting-app.git
cd priority-based-budgeting-app
```

2. Install dependencies:
```bash
npm install
```

3. Start the development server:
```bash
npm start
```

4. Open your browser to `http://localhost:3000`

## Usage

1. **Upload Data**: Drag and drop or click to upload your Excel budget file
2. **Apply Filters**: Use the filter dropdowns to narrow down requests
3. **Review Statistics**: View the summary cards showing filtered totals
4. **Generate Report**: Click "Generate Report" to create the formatted document
5. **Export**: Download the report as a Word document

## Deployment to Render.com

1. Push your code to GitHub
2. Connect your GitHub repository to Render.com
3. Set the build command: `npm install`
4. Set the start command: `npm start`
5. Deploy the application

## Technologies Used

- **Frontend**: HTML5, CSS3, JavaScript (ES6+)
- **Backend**: Node.js, Express.js
- **File Processing**: SheetJS (xlsx library)
- **Charts**: Chart.js
- **File Upload**: Multer
- **Styling**: Custom CSS with modern design patterns

## File Structure

```
priority-based-budgeting-app/
├── public/
│   ├── index.html          # Main application page
│   ├── styles.css          # Application styles
│   └── app.js             # Frontend JavaScript
├── server.js              # Express server
├── package.json           # Node.js dependencies
├── .gitignore            # Git ignore rules
└── README.md             # This file
```

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Test thoroughly
5. Submit a pull request

## License

MIT License - see LICENSE file for details

## Support

For questions or issues, please create an issue in the GitHub repository.