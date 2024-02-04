const express = require('express');
const bodyParser = require('body-parser');
const xlsx = require('xlsx');
const fs = require('fs');
const cors = require('cors');

const app = express();
const port = process.env.PORT || 3001;

app.use(bodyParser.json());
app.use(bodyParser.urlencoded({ extended: true }));
app.use(cors());

app.post('/submit-form', (req, res) => {
  const formData = req.body;
  const fileName = 'form_data.xlsx';
  const fileExists = fs.existsSync(fileName);
  let workbook;

  if (fileExists) {
    const fileContent = fs.readFileSync(fileName);
    workbook = xlsx.read(fileContent, { type: 'buffer' });
  } else {
    workbook = xlsx.utils.book_new();
    // Create a new worksheet with header row
    const header = ['FirstName', 'LastName', 'Email', 'Phone', 'Message', 'Agreement']; // Add your form fields here
    // const header = [''];
    const worksheet = xlsx.utils.json_to_sheet([header]);
    xlsx.utils.book_append_sheet(workbook, worksheet, 'Form Data');
  }

  // Get the existing worksheet
  const worksheet = workbook.Sheets['Form Data'];

  // Extract data from the worksheet
  const data = xlsx.utils.sheet_to_json(worksheet, { header: 1 });
  console.log('data '+data);
  console.log('formdata '+formData);

  // Add new form data to the existing data
  data.push(Object.values(formData));

  // Update the worksheet with the new data
  xlsx.utils.sheet_add_aoa(worksheet, data);

  // Write the updated workbook to the Excel file
  const excelBuffer = xlsx.write(workbook, { bookType: 'xlsx', type: 'buffer' });
  fs.writeFileSync(fileName, excelBuffer);

  res.status(200).send('Form data submitted successfully');

  fs.closeSync(fs.openSync(fileName, 'r+'));
});

app.listen(port, () => {
  console.log(`Server is running on port ${port}`);
});





// const express = require('express');
// const bodyParser = require('body-parser');
// const xlsx = require('xlsx');
// const fs = require('fs');
// const cors = require('cors');

// const app = express();
// const port = 3001;

// app.use(bodyParser.json());
// app.use(bodyParser.urlencoded({ extended: true }));
// app.use(cors());

// app.post('/submit-form', (req, res) => {
//   const formData = req.body;
//   const fileName = 'form_data.xlsx';
//   const fileExists = fs.existsSync(fileName);
//   let workbook;

//   if (fileExists) {
//     const fileContent = fs.readFileSync(fileName);
//     workbook = xlsx.read(fileContent, { type: 'buffer' });
//   } else {
//     workbook = xlsx.utils.book_new();
//     const header = ['FirstName', 'LastName', 'Email', 'Phone', 'Message', 'Agreement'];
//     const worksheet = xlsx.utils.json_to_sheet([header], { header: header });
//     xlsx.utils.book_append_sheet(workbook, worksheet, 'Form Data');
//   }

//   // Get the existing worksheet
//   const worksheet = workbook.Sheets['Form Data'];

//   // Extract data from the worksheet
//   const data = xlsx.utils.sheet_to_json(worksheet, { header: 1 });

//   // Check if the data array contains the header
//   const hasHeader = data.length > 0 && data[0].includes('FirstName');

//   // Remove the header from the data array if present
//   if (hasHeader) {
//     data.shift();
//   }

//   // Add new form data to the existing data
//   data.push(Object.values(formData));

//   // Update the worksheet with the new data
//   xlsx.utils.sheet_add_aoa(worksheet, data);

//   // Write the updated workbook to the Excel file
//   const excelBuffer = xlsx.write(workbook, { bookType: 'xlsx', type: 'buffer' });
//   fs.writeFileSync(fileName, excelBuffer);

//   res.status(200).send('Form data submitted successfully');

//   fs.closeSync(fs.openSync(fileName, 'r+'));
// });

// app.listen(port, () => {
//   console.log(`Server is running on port ${port}`);
// });
