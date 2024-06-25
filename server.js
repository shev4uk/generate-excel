import express from 'express';
import xl from 'excel4node';
import cors from 'cors';

const app = express();
const port = 3000;

// Use CORS middleware
app.use(cors());

app.use(express.json());

app.post('/generate-excel', async (req, res) => {
  const jsonData = req.body;

  // Create a new instance of a Workbook
  const wb = new xl.Workbook();

  // Add a Worksheet
  const ws = wb.addWorksheet('Sheet 1');

  // Create a reusable style
  const headerStyle = wb.createStyle({
    font: {
      color: '#FFFFFF',
      size: 12,
      bold: true,
    },
    fill: {
      type: 'pattern',
      patternType: 'solid',
      fgColor: '#4F81BD',
    },
    alignment: {
      horizontal: 'center',
    },
  });

  const cellStyle = wb.createStyle({
    font: {
      color: '#000000',
      size: 12,
    },
    numberFormat: '#,##0.00; (#,##0.00); -',
  });

  const {headers, rowsData} = jsonData;

const columnNames = Object.values(headers);
    const cellKeys = Object.keys(headers);
    columnNames.forEach((header, index) => {
      ws.cell(1, index + 1).string(header);
    });

    rowsData.forEach((data, rowIndex) => {
      cellKeys.forEach((key, colIndex) => {
        const cellValue = data[key];
        if (typeof cellValue === 'number') {
          ws.cell(rowIndex + 2, colIndex + 1).number(cellValue);
        } else {
          ws.cell(rowIndex + 2, colIndex + 1).string(cellValue);
        }
      });
    });

  const buffer = await wb.writeToBuffer();
  const base64Buffer = buffer.toString('base64');

  res.writeHead(200, {
    'Content-Type': 'application/json',
  });
  res.end(JSON.stringify({
    status: 'success',
    data: base64Buffer,
  }));
});

app.listen(port, () => {
  console.log(`Server is running on port ${port}`);
});
