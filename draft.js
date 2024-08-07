const wb = new xl.Workbook();
    const ws = wb.addWorksheet('Sheet 1');

    const headerStyle = wb.createStyle({
        fill: {
            type: 'pattern',
            patternType: 'solid',
            fgColor: '#4472C4',
        },
        font: {
            color: '#FFFFFF',
            bold: true,
        },
    });

    const subHeaderStyle = wb.createStyle({
        fill: {
            type: 'pattern',
            patternType: 'solid',
            fgColor: '#BDD7EE',
        },
        font: {
            bold: true,
        },
    });

    const captionStyle = wb.createStyle({
        fill: {
            type: 'pattern',
            patternType: 'solid',
            fgColor: '#f3fcf8',
        },
        font: {
            bold: true,
        },
    });

    const colWidths = [0, 0, 0, 0];

    const setMaxWidth = (colIndex, value) => {
        if (value) {
            const length = value.toString().length;
            if (length > colWidths[colIndex]) {
                colWidths[colIndex] = length;
            }
        }
    };

    ws.cell(1, 1).string(headers.memberName).style(captionStyle);
    ws.cell(1, 2).string(headers.memberId).style(captionStyle);
    ws.cell(1, 3).string(headers.benefitEffectiveDate).style(captionStyle);

    ws.cell(2, 1).string(caption.memberName).style(captionStyle);
    ws.cell(2, 2).string(caption.memberId).style(captionStyle);
    ws.cell(2, 3).string(caption.benefitEffectiveDate).style(captionStyle);

    let row = 4;
    
    rowsData.forEach((item, index) => {
        const header1 = item.primaryLocName;
        const coveredStatus = item.covered;
        setMaxWidth(0, header1);
        setMaxWidth(1, coveredStatus);
        ws.cell(row, 1).string(header1).style(headerStyle);
        ws.cell(row, 2).string('').style(headerStyle);
        ws.cell(row, 3).string('').style(headerStyle);
        ws.cell(row, 4).string(coveredStatus).style(headerStyle);

        row++;

        item.parentLevelOfCareList.forEach((loc) => {
            const subHeader = loc.parentLocName;
            setMaxWidth(0, subHeader);
            ws.cell(row, 1).string(subHeader).style(subHeaderStyle);
            ws.cell(row, 2).string('').style(subHeaderStyle);
            ws.cell(row, 3).string(headers.inNetworkLoc).style(subHeaderStyle);
            ws.cell(row, 4).string(headers.outNetworkLoc).style(subHeaderStyle);
            row++;

            const details = [
                [headers.coveredStr, loc.inNetworkLoc.coveredStr, loc.outNetworkLoc.coveredStr],
                [headers.authPercentReqStr, loc.inNetworkLoc.authPercentReqStr, loc.outNetworkLoc.authPercentReqStr],
                [headers.tier1CoInsurance, loc.inNetworkLoc.tier1CoInsurance, loc.outNetworkLoc.tier1CoInsurance],
                [headers.tier1Copay, loc.inNetworkLoc.tier1Copay, loc.outNetworkLoc.tier1Copay],
                [headers.deductibleFamilyStr, loc.inNetworkLoc.deductibleFamilyStr, loc.outNetworkLoc.deductibleFamilyStr],
                [headers.deductibleIndStr, loc.inNetworkLoc.deductibleIndStr, loc.outNetworkLoc.deductibleIndStr],
                [headers.oopFamilyStr, loc.inNetworkLoc.oopFamilyStr, loc.outNetworkLoc.oopFamilyStr],
                [headers.oopIndStr, loc.inNetworkLoc.oopIndStr, loc.outNetworkLoc.oopIndStr]
            ];

            details.forEach((detail) => {
                setMaxWidth(1, detail[0]);
                setMaxWidth(2, detail[1]);
                setMaxWidth(3, detail[2]);
                
                ws.cell(row, 2).string(detail[0]);
                ws.cell(row, 3).string(detail[1]);
                ws.cell(row, 4).string(detail[2]);
                row++;
            });
            
            row++;
        });
    });

    colWidths.forEach((width, index) => {
        ws.column(index + 1).setWidth(width + 2);
    });
