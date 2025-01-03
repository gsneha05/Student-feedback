const path = require('path');
const os = require('os');
const fs = require('fs');
const xlsx = require('xlsx');
const express = require('express');
const bodyParser = require('body-parser');

const app = express();
app.use(bodyParser.urlencoded({ extended: true }));
app.use(bodyParser.json());

// Path to save the Excel file on the Desktop
const desktopPath = path.join(os.homedir(), 'Desktop');
const filePath = path.join(desktopPath, 'feedbacks.xlsx');

// Function to write feedback to the Excel file
function writeFeedbackToExcel(feedback) {
    let workbook;
    let worksheet;

    // Check if the file exists
    if (fs.existsSync(filePath)) {
        // Read the existing Excel file and get the worksheet
        workbook = xlsx.readFile(filePath);
        worksheet = workbook.Sheets['Feedbacks'];
    } else {
        // Create a new workbook and worksheet if the file does not exist
        workbook = xlsx.utils.book_new();
        worksheet = xlsx.utils.aoa_to_sheet([
            ['Name', 'Courses', 'Rating', 'Feedback', 'AIPredictions', 'Confidence']
        ]);
        xlsx.utils.book_append_sheet(workbook, worksheet, 'Feedbacks');
    }

    // Convert the worksheet to JSON to append the new feedback
    const data = xlsx.utils.sheet_to_json(worksheet, { header: 1 });

    // Create a new row with empty columns for AIPredictions and Confidence
    const newRow = [feedback.name, feedback.courses.join(', '), feedback.rating, feedback.feedback, '', ''];
    data.push(newRow);

    // Convert the updated data back to a worksheet and overwrite the existing worksheet
    const newWorksheet = xlsx.utils.aoa_to_sheet(data);
    workbook.Sheets['Feedbacks'] = newWorksheet;  // Overwrite the existing worksheet

    // Write the updated workbook to the file
    xlsx.writeFile(workbook, filePath);
}

// Handle form submission
app.post('/submit-feedback', (req, res) => {
    const feedback = {
        name: req.body.name,
        courses: Array.isArray(req.body.course) ? req.body.course : [req.body.course],
        rating: req.body.rating,
        feedback: req.body.feedback
    };

    writeFeedbackToExcel(feedback);
    res.send('Feedback submitted successfully!');
});

// Serve the static HTML file
app.use(express.static('public'));

// Start the server
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`Server is running on port ${PORT}`);
});
