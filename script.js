document.getElementById('fileUpload').addEventListener('change', function(event) {
    const file = event.target.files[0];
    if (file) {
        const reader = new FileReader();
        reader.onload = async function(e) {
            const data = new Uint8Array(e.target.result);
            const workbook = new ExcelJS.Workbook();
            await workbook.xlsx.load(data.buffer);
            const worksheet = workbook.getWorksheet(1);

            // Set print titles to first 3 rows
            worksheet.pageSetup.printTitlesRow = '1:7';

            // Set narrow margins
            worksheet.pageSetup.margins = {
                left: 0.25,
                right: 0.25,
                top: 0.75,
                bottom: 0.75,
                header: 0.3,
                footer: 0.3
            };

            // Set page width to fit one page wide
            worksheet.pageSetup.fitToPage = true;
            worksheet.pageSetup.fitToWidth = 1;
            worksheet.pageSetup.fitToHeight = 0;

            // Delete rows 4, 5, and 6
            worksheet.spliceRows(4, 3);

            // // Define the unmergeCells function
            // const unmergeCells = (range) => {
            //     if (worksheet.getCell(range.split(':')[0]).isMerged) {
            //         worksheet.unMergeCells(range);
            //     }
            // };

            // Unmerge cells in Rows 1, 2, and 3 if they are merged from A to D
            // unmergeCells('A1:I1');
            // unmergeCells('A2:I2');
            // unmergeCells('A3:I3');

            // Merge cells in Rows 1, 2, and 3 from A to I
            // worksheet.mergeCells('A1:I1');
            // worksheet.mergeCells('A2:I2');
            // worksheet.mergeCells('A3:I3');

            // // Track already unmerged ranges
            // const unmergedRanges = new Set(); 

            // // Unmerge all merged cells in Row 5
            // for (let col = 1; col <= 9; col++) {
            //     const cell = worksheet.getCell(5, col);
                
            //     if (cell.isMerged) {
            //         const mergeRange = cell.master.address; // Get the merged range

            //         // Unmerge only if we haven't done so for this range
            //         if (!unmergedRanges.has(mergeRange)) {
            //             console.log(`Unmerging cells in range: ${mergeRange}`);
            //             worksheet.unMergeCells(mergeRange);
            //             unmergedRanges.add(mergeRange); // Store the range to prevent duplicate unmerge
            //         }
            //     }
            // }

            // Now, safely merge B5:E5
            // try {
            //     worksheet.mergeCells('B5:E5');
            //     console.log("Successfully merged B5:E5");
            // } catch (error) {
            //     console.error("Error merging B5:E5:", error);
            // }
            

            // Merge cells in Rows 5, from B to E, and from F to I
            // unmergeCells('B5:I5');
            // worksheet.mergeCells('B6:E6');
            // worksheet.mergeCells('F6:I6');

            function formatDate(dateStr) {
                let parts = dateStr.split(' ');e
                if (parts.length !== 3) return 'Invalid Date'; 
                let month = parts[0];
                let day = parts[1].replace(',', ''); 
                let year = parts[2];
                let monthNumber;
                switch (month) {
                    case 'January': monthNumber = 1; break;
                    case 'February': monthNumber = 2; break;
                    case 'March': monthNumber = 3; break;
                    case 'April': monthNumber = 4; break;
                    case 'May': monthNumber = 5; break;
                    case 'June': monthNumber = 6; break;
                    case 'July': monthNumber = 7; break;
                    case 'August': monthNumber = 8; break;
                    case 'September': monthNumber = 9; break;
                    case 'October': monthNumber = 10; break;
                    case 'November': monthNumber = 11; break;
                    case 'December': monthNumber = 12; break;
                    default: return 'Invalid Month';
                }
                return `${monthNumber}/${day}/${year}`;
            }

            cellValue = worksheet.getCell('A3').value;
            date = cellValue.substring(6).trim();

            proper_date = formatDate(date);

            worksheet.getCell('B5').value = "                                                          Month Ending";
            worksheet.getCell('C5').value = null;
            worksheet.getCell('D5').value = null;
            worksheet.getCell('E5').value = null;
            worksheet.getCell('F5').value = "                                                          Period Ending";
            worksheet.getCell('G5').value = null;
            worksheet.getCell('H5').value = null;
            worksheet.getCell('I5').value = null;

            worksheet.getCell('B6').value = `                                                          ${proper_date}`;
            worksheet.getCell('C6').value = null;
            worksheet.getCell('D6').value = null;
            worksheet.getCell('E6').value = null;
            worksheet.getCell('F6').value = `                                                          ${proper_date}`;
            worksheet.getCell('G6').value = null;
            worksheet.getCell('H6').value = null;
            worksheet.getCell('I6').value = null;

            // Center align the merged cells in Rows 1, 2, and 3
            worksheet.getCell('A1').alignment = { horizontal: 'center', vertical: 'middle' };
            worksheet.getCell('A2').alignment = { horizontal: 'center', vertical: 'middle' };
            worksheet.getCell('A3').alignment = { horizontal: 'center', vertical: 'middle' };

            // Wrap text for all cells in Column J
            worksheet.getColumn('J').eachCell((cell) => {
                cell.alignment = { wrapText: true };
                cell.border = {};
            });

            // Set the header for Column J
            const varianceHeader = worksheet.getCell('J6');
            varianceHeader.value = 'Variance Comments';

            // Apply styles to the header
            varianceHeader.font = { bold: true };
            varianceHeader.alignment = { horizontal: 'center', vertical: 'middle' };
            varianceHeader.border = { bottom: { style: 'thin' } };

            // Set the width of Column J
            worksheet.getColumn('J').width = 40;

            // Center align all cells in rows 9 and below
            worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
                if (rowNumber >= 7) {
                    row.eachCell((cell) => {
                        cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
                    });
                }
            });

            // Left align all cells in Column A for rows 4 and below
            worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
                if (rowNumber >= 4) {
                    const cell = row.getCell('A');
                    cell.alignment = { horizontal: 'left', vertical: 'middle', wrapText: true };
                }
            });

            // Remove cell with the text "Created on:"
            worksheet.eachRow((row) => {
                row.eachCell((cell) => {
                    if (cell.value === 'Created on:') {
                        cell.value = null;
                    }
                });
            });

            // Get the value of cell A1 and A3 to use as the filename
            let fileNameA1 = worksheet.getCell('A1').value || '';
            let fileNameA3 = worksheet.getCell('A3').value || '';
            if (typeof fileNameA3 === 'string' && fileNameA3.startsWith('As of ')) {
                fileNameA3 = fileNameA3.substring(6).trim();
            }
            const fileName = `${fileNameA1} - Full Reporting - ${fileNameA3}`.trim() || 'modified';

            // Output the modified content to the console
            const buffer = await workbook.xlsx.writeBuffer();
            const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
            const url = URL.createObjectURL(blob);
            const downloadLink = document.getElementById('downloadLink');
            const downloadButton = document.getElementById('downloadButton');
            downloadButton.href = url;
            downloadButton.download = `${fileName}.xlsx`;
            downloadLink.style.display = 'block';
        };
        reader.readAsArrayBuffer(file);
    }
});