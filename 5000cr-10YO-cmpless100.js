const puppeteer = require('puppeteer');
const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');
const pdf = require('html-pdf');

const getTimestamp = () => {
    const now = new Date();
    const day = String(now.getDate()).padStart(2, '0');
    const month = String(now.getMonth() + 1).padStart(2, '0');
    const year = String(now.getFullYear()).slice(-2);
    const hours = String(now.getHours()).padStart(2, '0');
    const minutes = String(now.getMinutes()).padStart(2, '0');
    return `${day}${month}${year}-${hours}${minutes}`;
};

(async () => {
    const reportDir = path.join(__dirname, 'Report');
    if (!fs.existsSync(reportDir)) {
        fs.mkdirSync(reportDir);
    }

    const browser = await puppeteer.launch({ headless: false });

    try {
        const page = await browser.newPage();
        const timestamp = getTimestamp();
        const excelFilePath = path.join(reportDir, `5000cr-10YO-cmpless100_${timestamp}.xlsx`);
        const pdfFilePath = path.join(reportDir, `5000cr-10YO-cmpless100_${timestamp}.pdf`);

        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('Report');

        // Styling
        const headerStyle = {
            font: { name: 'Calibri', color: { argb: 'FFFF00' }, size: 14 },
            fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: '0F9D58' } },
            border: {
                top: { style: 'thin', color: { argb: '000000' } },
                left: { style: 'thin', color: { argb: '000000' } },
                bottom: { style: 'thin', color: { argb: '000000' } },
                right: { style: 'thin', color: { argb: '000000' } }
            }
        };

        // Add custom headers with styling
        // const customHeaders = [
        //     'Market cap 500 cr',
        //     '10 year old with 10% profit',
        //     'Current market price is less 50/-'
        // ];
        // customHeaders.forEach((header) => {
        //     const row = worksheet.addRow([header]);
        //     row.eachCell(cell => cell.style = { ...headerStyle, font: { ...headerStyle.font, size: 14 } });
        // });

        // worksheet.addRow([]); // Add an empty row for separation

        // Add main header row with cell-wise styling
        const mainHeaderRow = worksheet.addRow(['Name', 'CMP', 'All time high', 'Change%']);
        mainHeaderRow.eachCell(cell => {
            cell.font = { name: 'Calibri', color: { argb: 'FFFF00' }, size: 14 }; // Yellow font color and font size 14
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '0F9D58' } }; // Green background color
            cell.border = {
                top: { style: 'thin', color: { argb: '000000' } },
                left: { style: 'thin', color: { argb: '000000' } },
                bottom: { style: 'thin', color: { argb: '000000' } },
                right: { style: 'thin', color: { argb: '000000' } }
            };
        });

        let pageIndex = 1;
        let previousPageData = null;

        while (true) {
            const url = `https://www.screener.in/screens/2181096/cap-5000cr-price-under-100/?page=${pageIndex}`;
            //const url = `https://www.screener.in/screens/1890383/cap-500-profit-10year-cmp-50-less/?page=${pageIndex}`;
            console.log(`Fetching data from ${url}...`);

            await page.goto(url, { waitUntil: 'networkidle2', timeout: 0 });
            await page.waitForSelector('body > main > div.card.card-large > div.responsive-holder.fill-card-width > table > tbody > tr');

            const companies = await page.evaluate(() => {
                const rows = document.querySelectorAll('body > main > div.card.card-large > div.responsive-holder.fill-card-width > table > tbody > tr');
                let result = [];

                rows.forEach(row => {
                    const nameElement = row.querySelector('td:nth-child(2) > a');
                    const cmpElement = row.querySelector('td:nth-child(3)');
                    const allTimeHighElement = row.querySelector('td:nth-child(13)');

                    if (nameElement && cmpElement && allTimeHighElement) {
                        const name = nameElement.innerText.trim();
                        const cmp = parseFloat(cmpElement.innerText.trim().replace(/,/g, ''));
                        const allTimeHigh = parseFloat(allTimeHighElement.innerText.trim().replace(/,/g, ''));
                        const change = (((allTimeHigh - cmp) / allTimeHigh) * 100).toFixed(2);

                        result.push({
                            name,
                            cmp: cmp.toFixed(2),
                            allTimeHigh: allTimeHigh.toFixed(2),
                            change
                        });
                    }
                });

                return result;
            });

            if (!companies || companies.length === 0 || (previousPageData && JSON.stringify(companies) === JSON.stringify(previousPageData))) {
                console.log('No more data available or duplicate data found. Stopping the process.');
                break;
            }

            const filteredCompanies = companies.filter(company => company.change >= 50 && company.change <= 60);
            //const filteredCompanies = companies.filter(company => company.change >= 30 && company.change <= 70);

            filteredCompanies.forEach(company => {
                const row = worksheet.addRow([company.name, company.cmp, company.allTimeHigh, company.change]);
                row.eachCell(cell => {
                    cell.border = {
                        top: { style: 'thin', color: { argb: '000000' } },
                        left: { style: 'thin', color: { argb: '000000' } },
                        bottom: { style: 'thin', color: { argb: '000000' } },
                        right: { style: 'thin', color: { argb: '000000' } }
                    };
                });
            });

            previousPageData = companies;
            pageIndex++;
        }

        console.log(`Excel file successfully created: ${excelFilePath}`);
        await workbook.xlsx.writeFile(excelFilePath);

        console.log(`Excel file saved: ${excelFilePath}`);

        // Convert Excel to HTML
        const html = `
            <html>
            <head>
                <style>
                    table { border-collapse: collapse; width: 100%; }
                    .header td {
            padding: 8px;
            text-align: center;
            border: 1px solid #ddd;
            background-color: #0F9D58;
            color: #FFFF00;
            font-family: Calibri, sans-serif;
            font-size: 24px;
        }
                    th, td { border: 1px solid black; padding: 8px; text-align: left; font-size: 14px; font-family: Calibri; }
                    th { background-color: #0F9D58; color: #FFFF00; }
                </style>
            </head>
            <body>
            <div class="header">
        <table>
            <tr>
                <td>Market cap More Than 5000cr | 10 Year Old with 10% profit | CMP is less 100/-</td>
                
            </tr>
            <tr>
                
            </tr>
        </table>
    </div>
                <table>
<tr>
&nbsp;
</tr>
                    <tr>
                       <th>Name</th>
                        <th>CMP</th>
                        <th>All time high</th>
                         <th>Change between 50% to 60% </th>
                    </tr>
                    ${worksheet.getSheetValues().slice(2).map(row => `
                        <tr>
                            ${row.map(cell => `<td>${cell}</td>`).join('')}
                        </tr>
                    `).join('')}
                </table>
            </body>
            </html>
        `;

        const htmlFilePath = path.join(reportDir, `5000cr-10YO-cmpless100_${timestamp}.html`);
        fs.writeFileSync(htmlFilePath, html);

        console.log(`HTML file created: ${htmlFilePath}`);

        // const pdfOptions = {
        //     format: 'A4',
        //     printBackground: true
        // };

        // pdf.create(html, pdfOptions).toFile(pdfFilePath, (err, res) => {
        //     if (err) return console.log(err);
        //     console.log(`PDF file created: ${pdfFilePath}`);
        // });

    } catch (error) {
        console.error('Error:', error);
    } finally {
        await browser.close();
    }
})();
