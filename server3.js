import express from 'express';
import bodyParser from 'body-parser';
import xl from 'excel4node';
import path from 'path';
import { fileURLToPath } from 'url';
import { dirname } from 'path';

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

const app = express();
const port = 3000;

app.use(bodyParser.json());

app.post('/generate-excel', (req, res) => {
    const { headers, claimPaymentDetail, claimPaymentDetails, explanation, status } = req.body;

    const wb = new xl.Workbook();
    const ws = wb.addWorksheet('Sheet 1');

    const titleStyle = wb.createStyle({
        font: {
            bold: true,
            size: 14,
        },
    });

    const headerStyle = wb.createStyle({
        font: {
            bold: true,
            size: 12,
        },
    });

    const subHeaderStyle = wb.createStyle({
        font: {
            bold: true,
        },
    });

    // Add "Claim details" title
    ws.cell(1, 1).string(headers.titles.details).style(titleStyle);

    // Add Claim details
    const detailHeaders = headers.details;
    const claimDetails = [
        [detailHeaders.claimNumber, claimPaymentDetail.claimNumber],
        [detailHeaders.providerName, claimPaymentDetail.providerName],
        [detailHeaders.startDate, claimPaymentDetail.serviceStartDate],
        [detailHeaders.endDate, claimPaymentDetail.serviceEndDate],
        [detailHeaders.paidDate, claimPaymentDetail.paidDate || ''],
        [detailHeaders.paidTo, claimPaymentDetail.paidTo],
        [detailHeaders.amountTowards, `$${claimPaymentDetail.amountAppliedTowardsDeductible}`],
        ["Associated Authorization", claimPaymentDetail.claimNumber],
        [detailHeaders.status, status],
    ];

    let row = 2;
    claimDetails.forEach((detail) => {
        ws.cell(row, 1).string(detail[0]).style(headerStyle);
        ws.cell(row, 2).string(detail[1]);
        row++;
    });

    // Add "Payment details" title
    ws.cell(row, 1).string(headers.titles.payment).style(titleStyle);
    row++;

    // Add payment details headers
    const tableHeaders = headers.table;
    const paymentHeaders = [
        tableHeaders.serviceDate,
        tableHeaders.typeOfService,
        tableHeaders.totalCharge,
        tableHeaders.paidAmount,
        tableHeaders.amountAllowed,
        tableHeaders.appliedToDeductible,
        tableHeaders.copayment,
        tableHeaders.coinsurance,
        tableHeaders.eob,
    ];

    paymentHeaders.forEach((header, index) => {
        ws.cell(row, index + 1).string(header).style(subHeaderStyle);
    });

    row++;

    // Add payment details data
    claimPaymentDetails.forEach((item) => {
        ws.cell(row, 1).string(item.dateOfService);
        ws.cell(row, 2).string(item.serviceDesc);
        ws.cell(row, 3).number(item.claimedAmount);
        ws.cell(row, 4).number(item.amountPaid);
        ws.cell(row, 5).number(item.allowedAmount);
        ws.cell(row, 6).number(item.deductibleAmount);
        ws.cell(row, 7).number(item.copayAmount);
        ws.cell(row, 8).number(item.coinsuranceAmount);
        ws.cell(row, 9).string(item.eopCodes);
        row++;
    });

    // Add Explanation of Benefits
    ws.cell(row, 1).string(headers.titles.eob).style(titleStyle);
    row++;

    ws.cell(row, 1).string(explanation.replace(/<\/?[^>]+(>|$)/g, ""));

    // Set column widths based on content
    const colWidths = [20, 50, 20, 20, 20, 20, 20, 20, 20];
    colWidths.forEach((width, index) => {
        ws.column(index + 1).setWidth(width);
    });

    wb.write('ClaimsDetails.xlsx', res);
});

app.use(express.static(path.join(__dirname, 'public')));

app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, '', 'index3.html'));
});

app.listen(port, () => {
    console.log(`Server running at http://localhost:${port}`);
});
