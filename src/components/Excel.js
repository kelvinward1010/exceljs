import React, { useState } from 'react'
import { datafake } from '../data';
import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';

function Excel() {

    const data = datafake;
    const [excelFilterEnabled, setExcelFilterEnabled] = useState(true)

    const workSheetName = 'Work Sheet Test Name';
    const workBookName = 'MyWorkBook';
    const myInputId = 'myInput';

    const handleExport = async () => {

        //Change name file download
        const myInput = document.getElementById(myInputId);
        const fileName = myInput.value || workBookName;

        // const workbook = new ExcelJS.Workbook();
        const htmlTable = document.getElementById('myTable');
 
        // Create a new workbook and worksheet using ExcelJS
        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet(workSheetName, {properties:{tabColor:{argb:'FFC0000'}}});

        // gộp dònng theo ô trên excel 
        worksheet.mergeCells('A2', 'I2');

        // create some style for sheet
        //const worksheet = workbook.addWorksheet('My Sheet', {properties:{tabColor:{argb:'FFC0000'}}}, {views:[{state: 'frozen', xSplit: 4, ySplit:1}]});
        worksheet.getRow(1).font = { bold: true };

        // const customCell = worksheet.getCell("A2");
        // customCell.font = {
        //     name: "Comic Sans MS",
        //     family: 4,
        //     size: 5,
        //     underline: true,
        //     bold: true
        // };
 
        // Add the column names to the worksheet
        const headerRow = worksheet.addRow([]);
        worksheet.getRow(4).font = { bold: true };
        const headerCells = htmlTable.getElementsByTagName('th');
            for (let i = 0; i < headerCells.length; i++) {
                headerRow.getCell(i + 1).value = headerCells[i].innerText;
        }
 
        // Add the HTML table data to the worksheet
        const rows = htmlTable.getElementsByTagName('tr');
        for (let i = 0; i < rows.length; i++) {
            const cells = rows[i].getElementsByTagName('td');
            const rowData = [];

            for (let j = 0; j < cells.length; j++) {
                rowData.push(cells[j].innerText);
            }
            worksheet.addRow(rowData);
        }

        worksheet.columns.forEach(column => {
            column.width = headerCells.length + 10
            column.alignment = { horizontal: 'left' };
        });

        // auto create a row fillter
        // worksheet.autoFilter = {
        //     from: {
        //       row: 4,
        //       column: 1
        //     },
        //     to: {
        //       row: 4,
        //       column: 10
        //     }
        // };

        // worksheet.eachRow({ includeEmpty: false }, row => {
        //     // store each cell to currentCell
        //     const currentCell = row._cells;
    
        //     // loop through currentCell to apply border only for the non-empty cell of excel
        //     currentCell.forEach(singleCell => {
        //       // store the cell address i.e. A1, A2, A3, B1, B2, B3, ...
        //       const cellAddress = singleCell._address;
    
        //       // apply border
        //       worksheet.getCell(cellAddress).border = {
        //         top: { style: 'thin' },
        //         left: { style: 'thin' },
        //         bottom: { style: 'thin' },
        //         right: { style: 'thin' }
        //       };
        //     });
        // });

 
        // Generate a blob object from the workbook and download it as an attachment
        const buf = await workbook.xlsx.writeBuffer();
        // workbook.xlsx.writeBuffer().then(function (buffer) {
        //     const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        // });
        saveAs(new Blob([buf]), `${fileName}.xlsx`);
    }

    return (
        <div>
            <div>
                Export to excel from table
                <br />
                <br />
                Export to : <input id={myInputId} defaultValue={workBookName} /> .xlsx
            </div>
            <table id="myTable">
                <thead>
                    <tr>
                        <th>ID</th>
                        <th>FirstName</th>
                        <th>LastName</th>
                        <th>Prefix</th>
                        <th>Position</th>
                        <th>BirthDate</th>
                        <th>HireDate</th>
                        <th>State</th>
                        <th>City</th>
                        <th>SaleAmount</th>
                    </tr>
                </thead>
                <tbody>
                    {data.map((item,idx) => (
                        <tr key={idx}>
                            <td>{item.ID}</td>
                            <td>{item.FirstName}</td>
                            <td>{item.LastName}</td>
                            <td>{item.Prefix}</td>
                            <td>{item.Position}</td>
                            <td>{item.BirthDate}</td>
                            <td>{item.HireDate}</td>
                            <td>{item.Notes}</td>
                            <td>{item.State}</td>
                            <td>{item.City}</td>
                            <td>{item.SaleAmount}</td>
                        </tr>
                    ))}
                </tbody>
            </table>
            <button onClick={handleExport}>Export to Excel</button>
        </div>
    )
}

export default Excel
