import Excel from 'exceljs';
import { saveAs } from 'file-saver';

const columns = [
    { header: 'ID', key: 'id' },
    { header: 'First Name', key: 'firstName' },
    { header: 'Last Name', key: 'lastName' },
    { header: 'Purchase Price', key: 'purchasePrice' },
    { header: 'Payments Made', key: 'paymentsMade' }
  ];
  
  const data = [
    {
        id: 1,
        firstName: 'Kylie',
        lastName: 'James',
        purchasePrice: 1000,
        paymentsMade: 900
    },
    {
        id: 2,
        firstName: 'Harry',
        lastName: 'Peake',
        purchasePrice: 1000,
        paymentsMade: 1000
    },
    {
        id: 3,
        firstName: 'parker',
        lastName: 'bie',
        purchasePrice: 1000,
        paymentsMade: 1000
    },
    {
        id: 4,
        firstName: 'kevin',
        lastName: 'love',
        purchasePrice: 1099,
        paymentsMade: 1340
    },
  ];
  
  const workSheetName = 'Worksheet-1';
  const workBookName = 'MyWorkBook';
  const myInputId = 'myInput';
  
  export default function Excel1() {
    const workbook = new Excel.Workbook();
    
  
    const saveExcel = async () => {
      try {
        const table = document.getElementById(myInputId);
        const fileName = table.value || workBookName;
  
        // creating one worksheet in workbook and stylesheet
        const worksheet = workbook.addWorksheet(workSheetName, {properties:{tabColor:{argb:'FFC0000'}}}, {
            pageSetup:{fitToPage: true, fitToHeight: 5, fitToWidth: 7}
        });

        // merge by top-left, bottom-right
        worksheet.mergeCells('A2', 'I2');
        //worksheet.mergeCells('A4:B5');

        // node something in a cell
        //worksheet.getCell('A1').note = 'Hello, ExcelJS!';

        // worksheet.addTable({
        //     name: 'MyTable',
        //     ref: 'A11',
        //     headerRow: true,
        //     totalsRow: true,
        //     style: {
        //       theme: 'TableStyleDark3',
        //       showRowStripes: true,
        //     },
        //     columns: [
        //       {name: 'Date', totalsRowLabel: 'Totals:', filterButton: true},
        //       {name: 'Amount', totalsRowFunction: 'sum', filterButton: false},
        //     ],
        //     rows: [
        //       [new Date('2019-07-20'), 70.10],
        //       [new Date('2019-07-21'), 70.60],
        //       [new Date('2019-07-22'), 70.10],
        //     ],
        // });

        
  
        // add worksheet columns
        // each columns contains header and its mapping key from data
        worksheet.columns = columns;
  
        // updated the font for first row.
        worksheet.getRow(1).font = { bold: true };
  
        // loop through all of the columns and set the alignment with width.
        worksheet.columns.forEach(column => {
          column.width = column.header.length + 20;
          column.alignment = { horizontal: 'center' };
          //column.font = 'Code';
          column.style = {font:{bold: true, name: 'Comic Sans MS'}};
        });

        // style detail for each column
        // worksheet.columns = [
        //     { header: 'First Name', key: 'firstName', width: 60 },
        //     { header: 'Last Name', key: 'lastName', width: 32, style: { font: { name: 'Arial Black' } } },
        // ];

        

        // worksheet.columns = [
        //     { header: 'First Name', key: 'firstName', width: 50, },
        // ];

        //const rows = worksheet.getRows(2, 50); // start and length
        

        // take info row
        //const row1 = worksheet.getRow(1);
        // Set a specific row height
        //row1.height = 42.5;

        // make row hidden
        //row.hidden = true;
        

        // loop through data and add each one to worksheet
        data.forEach(singleData => {
            worksheet.addRow(singleData);
        });


        // add a row for table
        worksheet.addRow({ 
            id: 6,
            firstName: 'kevin11111',
            lastName: 'love1111',
            purchasePrice: 1099,
            paymentsMade: 1340
        });
  
        // loop through all of the rows and set the outline style.
        worksheet.eachRow({ includeEmpty: false }, row => {
          // store each cell to currentCell
          const currentCell = row._cells;

          row.height = 50

          
  
          // loop through currentCell to apply border only for the non-empty cell of excel
          currentCell.forEach(singleCell => {
            // store the cell address i.e. A1, A2, A3, B1, B2, B3, ...
            const cellAddress = singleCell._address;
  
            //apply border
            worksheet.getCell(cellAddress).border = {
                top: {style:'double', color: {argb:'FF00FF00'}},
                left: {style:'double', color: {argb:'FF00FF00'}},
                bottom: {style:'double', color: {argb:'FFE580FF'}},
                right: {style:'double', color: {argb:'FF00FF00'}}
            };


            worksheet.getCell('A2').border = {
                diagonal: {up: true, down: true, style:'thick', color: {argb:'FFFF0000'}}
            }

            worksheet.getCell('A4').fill = {
                type: 'gradient',
                gradient: 'angle',
                degree: 0,
                stops: [
                  {position:0, color:{argb:'FF0000FF'}},
                  {position:0.5, color:{argb:'FFFFFFFF'}},
                  {position:1, color:{argb:'FF0000FF'}}
                ]
            };

            // worksheet.getCell('A3').value = {
            //     'richText': [
            //       {'font': {'size': 12,'color': {'theme': 0},'name': 'Calibri','family': 2,'scheme': 'minor'},'text': 'This is '},
            //       {'font': {'italic': true,'size': 12,'color': {'theme': 0},'name': 'Calibri','scheme': 'minor'},'text': 'a'},
            //       {'font': {'size': 12,'color': {'theme': 1},'name': 'Calibri','family': 2,'scheme': 'minor'},'text': ' '},
            //       {'font': {'size': 12,'color': {'argb': 'FFFF6600'},'name': 'Calibri','scheme': 'minor'},'text': 'colorful'},
            //       {'font': {'size': 12,'color': {'theme': 1},'name': 'Calibri','family': 2,'scheme': 'minor'},'text': ' text '},
            //       {'font': {'size': 12,'color': {'argb': 'FFCCFFCC'},'name': 'Calibri','scheme': 'minor'},'text': 'with'},
            //       {'font': {'size': 12,'color': {'theme': 1},'name': 'Calibri','family': 2,'scheme': 'minor'},'text': ' in-cell '},
            //       {'font': {'bold': true,'size': 12,'color': {'theme': 1},'name': 'Calibri','family': 2,'scheme': 'minor'},'text': 'format'}
            //     ]
            // };
      

            //APP a cell address
            //worksheet.getCell('D1').alignment = { wrapText: true };

          });
        });

        worksheet.getCell('A11').value = 5;



        //Conditional Formatting⬆
        // worksheet.addConditionalFormatting({
        //     ref: 'A1:E7', //row and column
        //     rules: [
        //       {
        //         type: 'expression',
        //         formulae: ['MOD(ROW()+COLUMN(),2)=0'],
        //         style: {fill: {type: 'pattern', pattern: 'solid', bgColor: {argb: 'FF00FF00'}}},
        //       }
        //     ]
        // })

        
  
        // write the content using writeBuffer
        const buf = await workbook.xlsx.writeBuffer();
  
        // download the processed file
        saveAs(new Blob([buf]), `${fileName}.xlsx`);
      } catch (error) {
        console.error('<<<ERRROR>>>', error);
        console.error('Something Went Wrong', error.message);
      } finally {
        // removing worksheet's instance to create new one
        workbook.removeWorksheet(workSheetName);
      }
    };
  
    return (
      <>
        <style>
          {`
            table, th, td {
              border: 1px solid black;
              border-collapse: collapse;
              textAlign: center;
            }
             th, td { 
               padding: 4px;
             }
          `}
        </style>
        <div style={{ textAlign: 'center' }}>
          <div>
            Export to excel from table
            <br />
            <br />
            Export to : <input id={myInputId} defaultValue={workBookName} /> .xlsx
          </div>
  
          <br />
          <div>
            <button onClick={saveExcel}>Export</button>
          </div>
  
          <br />
  
          <div>
            <table style={{ margin: '0 auto' }}>
            <thead>
              <tr>
                {columns.map((item) => {
                  return <th key={item.key}>{item.header}</th>;
                })}
              </tr>
            </thead>
            <tbody>
              {data.map((uniqueData,idx) => {
                return (
                  <tr key={idx}>
                    <td>{uniqueData.id}</td>
                    <td>{uniqueData.firstName}</td>
                    <td>{uniqueData.lastName}</td>
                    <td>{uniqueData.purchasePrice}</td>
                    <td>{uniqueData.paymentsMade}</td>
                  </tr>
                );
              })}
            </tbody>
            </table>
          </div>
        </div>
      </>
    );
}