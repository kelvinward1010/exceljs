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
  
  const workSheetName = 'Tờ khai hàng hóa nhập khẩu';
  const workBookName = 'Tờ khai hàng hóa nhập khẩu';
  const myInputId = 'myInput';
  
  export default function Excel2() {
    const workbook = new Excel.Workbook();
    
  
    const saveExcel = async () => {
      try {
        const table = document.getElementById(myInputId);
        const fileName = table.value || workBookName;
  
        // creating one worksheet in workbook and stylesheet
        const worksheet = workbook.addWorksheet(workSheetName, {properties:{tabColor:{argb:'FFC0000'}}}, {
            pageSetup:{fitToPage: true, fitToHeight: 5, fitToWidth: 7}
        });

        worksheet.mergeCells('F3', 'X3');
        worksheet.mergeCells('C4', 'D4');
        worksheet.mergeCells('E4', 'K4');
        worksheet.mergeCells('L4', 'Q4');
        worksheet.mergeCells('L5', 'X5');
        worksheet.mergeCells('I6', 'K6');
        worksheet.mergeCells('AA4:AH5');
        worksheet.mergeCells('P6', 'T6');
        worksheet.mergeCells('U6', 'AC6');
        worksheet.mergeCells('AE6', 'AH6');
        worksheet.mergeCells('L7', 'S7');
        worksheet.mergeCells('T7', 'AC7');
        worksheet.mergeCells('AE7', 'AH7');
        worksheet.mergeCells('G8', 'K8');
        worksheet.mergeCells('R8', 'U8');
        worksheet.mergeCells('V8', 'AC8');
        worksheet.mergeCells('AE8', 'AH8');
        worksheet.mergeCells('D10', 'G10');
        worksheet.mergeCells('H10', 'AH10');
        worksheet.mergeCells('D11', 'G11');
        worksheet.mergeCells('H11:AH12');
        worksheet.mergeCells('D13', 'G13');
        worksheet.mergeCells('H13', 'AH13');
        worksheet.mergeCells('D14', 'G14');
        worksheet.mergeCells('H14:AH15');
        worksheet.mergeCells('H16', 'AF16');
        worksheet.mergeCells('H18', 'AA18');
        worksheet.mergeCells('H19:AH20');
        worksheet.mergeCells('H22', 'P22');
        worksheet.mergeCells('H23', 'AF23');
        worksheet.mergeCells('D24', 'G24');
        worksheet.mergeCells('H24', 'AE24');
        worksheet.mergeCells('D25:F26');
        worksheet.mergeCells('H25', 'T25');
        worksheet.mergeCells('U25', 'AG25');
        worksheet.mergeCells('H26', 'T26');
        worksheet.mergeCells('U26', 'AG26');
        worksheet.mergeCells('H27', 'AE27');
        worksheet.mergeCells('G29', 'Y29');
        worksheet.mergeCells('AF29', 'AH29');
        worksheet.mergeCells('E30', 'P30');
        worksheet.mergeCells('U30', 'X30');
        worksheet.mergeCells('Z30', 'AH30');
        worksheet.mergeCells('D31', 'P31');
        worksheet.mergeCells('U31', 'X31');
        worksheet.mergeCells('Z31', 'AH31');
        worksheet.mergeCells('D32', 'P32');
        worksheet.mergeCells('U32', 'X32');
        worksheet.mergeCells('Z32', 'AH32');
        worksheet.mergeCells('D33', 'P33');
        worksheet.mergeCells('D34', 'P34');
        worksheet.mergeCells('T34', 'X34');
        worksheet.mergeCells('Z34', 'AH34');
        worksheet.mergeCells('D35', 'P35');
        worksheet.mergeCells('U35', 'AH35');
        worksheet.mergeCells('K36', 'N36');
        worksheet.mergeCells('K37', 'N37');
        worksheet.mergeCells('K38', 'N38');
        worksheet.mergeCells('R39', 'X39');
        worksheet.mergeCells('Y39', 'AH39');
        worksheet.mergeCells('R40', 'X40');
        worksheet.mergeCells('Y40', 'AH40');
        worksheet.mergeCells('J41', 'AH41');
        const title = worksheet.getRow(3);
        title.height = 40;

        //const dobCol = worksheet.getColumn(3);
        //dobCol.values = ['Date of Birth'];
        //worksheet.getCell('A3').value = { formula: 'A1+A2', result: 7 };
        worksheet.getCell('F3', 'X3').value = 'Tờ khai hàng hóa nhập khẩu (Thông quan)';
        worksheet.getCell('F3', 'X3').alignment = { vertical: 'middle', horizontal: 'center' };
        worksheet.getCell('F3', 'X3').font = {
            name: 'Arial',
            family: 2,
            size: 20,
            bold: true
        }
        
        worksheet.getCell('C4', 'D4').value = 'Số tờ khai'
        worksheet.getCell('C4', 'D4').font = {
            name: 'Arial',
            family: 2,
            size: 10,
        };

        worksheet.getCell('E4', 'K4').value = '104517461160';
        worksheet.getCell('E4', 'K4').alignment = { vertical: 'middle', horizontal: 'center' };

        worksheet.getCell('L4', 'Q4').value ='Số tờ khai đầu tiên';
        worksheet.getCell('L4', 'Q4').alignment = { vertical: 'middle', horizontal: 'center' };

        worksheet.getCell('AA4:AH6').value = '0451746116'
        worksheet.getCell('AA4:AH6').alignment = { vertical: 'middle', horizontal: 'center' };

        worksheet.getCell('C5').value = 'Số tờ khai tạm nhập tái xuất tương ứng';

        worksheet.getCell('C6').value = "Mã phân loại kiểm tra";

        worksheet.getCell('I6', 'K6').value = '1';

        worksheet.getCell('L6').value = 'Mã loại hình';

        worksheet.getCell('P6', 'T6').value = 'E21  2 [4]';
        worksheet.getCell('P6', 'T6').alignment = { vertical: 'middle', horizontal: 'center' };

        worksheet.getCell('U6', 'AC6').value ='Mã số hàng hóa đại diện của tờ khai';
        worksheet.getCell('U6', 'AC6').alignment = { vertical: 'middle', horizontal: 'center' };

        worksheet.getCell('AE6', 'AH6').value ='8507';
        worksheet.getCell('AE6', 'AH6').alignment = { vertical: 'middle', horizontal: 'center' };

        worksheet.getCell('C7').value ='Tên cơ quan Hải quan tiếp nhập tờ khai';

        worksheet.getCell('L7', 'S7').value ='BACGIANGBN';
        worksheet.getCell('L7', 'S7').alignment = { vertical: 'middle', horizontal: 'center' };
        
        worksheet.getCell('T7', 'AC7').value ='Mã bộ phận xử lý tờ khai';
        worksheet.getCell('T7', 'AC7').alignment = { vertical: 'middle', horizontal: 'right' };

        worksheet.getCell('AE7', 'AH7').value ='00';
        worksheet.getCell('AE7', 'AH7').alignment = { vertical: 'middle', horizontal: 'left' };

        worksheet.getCell('C8').value ='Ngày đăng ký';

        worksheet.getCell('G8', 'K8').value ='08/02/2022  14:42:07';
        worksheet.getCell('G8', 'K8').alignment = { vertical: 'middle', horizontal: 'left' };

        worksheet.getCell('L8').value ='Ngày thay đổi đăng ký';

        worksheet.getCell('V8', 'AC8').value ='Thời hạn tái nhập/ tái xuất';
        worksheet.getCell('V8', 'AC8').alignment = { vertical: 'middle', horizontal: 'right' };

        worksheet.getCell('C9').value ='Người nhập khẩu';

        worksheet.getCell('D10', 'G10').value ='Mã';
        worksheet.getCell('H10', 'AH10').value ='2400859625';

        worksheet.getCell('D11', 'G11').value ='Tên';

        worksheet.getCell('H11:AH12').value ='CÔNG TY TNHH SEOJIN VIỆT NAM';
        worksheet.getCell('H11:AH12').alignment = { vertical: 'top', horizontal: 'left' };

        worksheet.getCell('D13', 'G13').value ='Mã bưu chính';
        worksheet.getCell('H13', 'AH13').value ='(+84) 43';

        worksheet.getCell('D14', 'G14').value ='Địa chỉ';
        worksheet.getCell('H14:AH15').value = 'Lô B1, B2, B3, B6, B7 KCN Song Khê Nội Hoàng(phía Bắc), xã Song Khê, TP Bắc Giang, Tỉnh Bắc Giang, Việt Nam';
        worksheet.getCell('H14:AH15').alignment = { vertical: 'top', horizontal: 'left' };
       
        worksheet.getCell('D16').value ='Số điện thoại';
        worksheet.getCell('H16', 'AF16').value ='0966862023';

        worksheet.getCell('C17').value ='Người ủy thác nhập khẩu';

        worksheet.getCell('D18').value ='Mã';

        worksheet.getCell('D19').value ='Tên';

        worksheet.getCell('C21').value ='Người xuất khẩu';

        worksheet.getCell('D22').value ='Mã';

        worksheet.getCell('D23').value ='Tên';
        worksheet.getCell('H23', 'AF23').value ='AST TECHNOLOGIES CO.,  LTD.';

        worksheet.getCell('D24', 'G24').value ='Mã bưu chính';

        worksheet.getCell('D25:F26').value ='Địa chỉ';

        worksheet.getCell('H25', 'T25').value ='80, SAPYEONG-DAERO';
        worksheet.getCell('U25', 'AG25').value ='SEOCHO-GU';

        worksheet.getCell('H26', 'T26').value ='SEOUL';
        worksheet.getCell('U25', 'AG25').value ='REPUBLIC OF KOREA';

        worksheet.getCell('D27').value ='Mã nước';
        worksheet.getCell('H27', 'AE27').value ='KR';

        worksheet.getCell('C28').value ='Người ủy thác xuất khẩu';

        worksheet.getCell('C29').value ='Đại lý hải quan';
        worksheet.getCell('AA29').value ='Mã nhân viên hải quan';

        worksheet.getCell('C30').value ='Số vận đơn';
        worksheet.getCell('R30').value ='Địa điểm lưu kho';
        worksheet.getCell('U30', 'X30').value ='03TGS04';
        worksheet.getCell('Z30', 'AH30').value ='CTY CP CONTAINER VN';

        worksheet.getCell('C31').value ='1';
        worksheet.getCell('C31').alignment = { vertical: 'middle', horizontal: 'center' };
        worksheet.getCell('D31','P31').value ='050122ED02201002';
        worksheet.getCell('R31').value ='Địa điểm dỡ hàng';
        worksheet.getCell('U31', 'X31').value ='VNGEE';
        worksheet.getCell('Z31', 'AH31').value ='GREEN PORT (HP)';

        worksheet.getCell('C32').value ='2';
        worksheet.getCell('C32').alignment = { vertical: 'middle', horizontal: 'center' };
        worksheet.getCell('R32').value ='Địa điểm xếp hàng';
        worksheet.getCell('U32', 'X32').value ='CNXMN';
        worksheet.getCell('Z32', 'AH32').value ='XIAMEN';

        worksheet.getCell('C33').value ='3';
        worksheet.getCell('C33').alignment = { vertical: 'middle', horizontal: 'center' };
        worksheet.getCell('R33').value ='Phương tiện vận chuyển';

        worksheet.getCell('C34').value ='4';
        worksheet.getCell('C34').alignment = { vertical: 'middle', horizontal: 'center' };
        worksheet.getCell('T34', 'X34').value ='9999';
        worksheet.getCell('T34', 'X34').alignment = { horizontal: 'right' };
        worksheet.getCell('Z34', 'AH34').value ='SITC OSAKA / 22045';

        worksheet.getCell('C35').value ='5';
        worksheet.getCell('C35').alignment = { vertical: 'middle', horizontal: 'center' };
        worksheet.getCell('R35').value ='Ngày hàng đến';
        worksheet.getCell('U35', 'AH35').value ='07/02/2022';
        worksheet.getCell('U35', 'AH35').alignment = { vertical: 'middle', horizontal: 'center' };

        worksheet.getCell('C36').value ='Số lượng';
        worksheet.getCell('K36', 'N36').value ='480';
        worksheet.getCell('K36', 'N36').alignment = { vertical: 'middle', horizontal: 'right' };
        worksheet.getCell('P36').value ='PK';
        worksheet.getCell('R36').value ='Kí hiệu và số liệu';

        worksheet.getCell('C37').value ='Tổng trọng lượng hàng (Gross)';
        worksheet.getCell('K37', 'N37').value ='179.840';
        worksheet.getCell('K37', 'N37').alignment = { vertical: 'middle', horizontal: 'right' };
        worksheet.getCell('P37').value ='KGM';

        worksheet.getCell('C38').value ='Số lượng container';
        worksheet.getCell('K38', 'N38').value ='10';
        worksheet.getCell('K38', 'N38').alignment = { vertical: 'middle', horizontal: 'right' };

        worksheet.getCell('R39', 'X39').value ='Ngày được phép nhập kho đầu tiên';

        worksheet.getCell('R40', 'X40').value ='Mã văn bản pháp quy khác';
        worksheet.getCell('Y40', 'AH40').value ='MO';
        worksheet.getCell('Y40', 'AH40').alignment = { vertical: 'middle', horizontal: 'center' };

        worksheet.getCell('C41').value ='Số hóa đơn';

        const row9 = worksheet.getRows(9, 1);
        row9.forEach(item => item.border ={top: { style: 'thin' }})

        const row21 = worksheet.getRows(21, 1);
        row21.forEach(item => item.border ={top: { style: 'thin' },})

        const row30 = worksheet.getRows(30, 1);
        row30.forEach(item => item.border ={top: { style: 'thin' },})

        const row41 = worksheet.getRows(41, 1);
        row41.forEach(item => item.border ={top: { style: 'thin' },})

        worksheet.getCell('P30').border = {right: { style: 'thin'}, top: { style: 'thin'}}
        for(let i = 31; i <= 40; i++){
          worksheet.getCell(`P${i}`).border = {right: { style: 'thin'}}
        }
        // const dobCol = worksheet.getColumn(16);
        // const t = dobCol._worksheet._columns.slice(30,33)
        // t.border = {right: { style: 'thin' }}
        
        

        // width column
        const w1 = worksheet.getColumn(1)
        w1.width = 1;
        const w2 = worksheet.getColumn(2)
        w2.width = 1;
        const w5 = worksheet.getColumn(5)
        w5.width = 1;
        const w12 = worksheet.getColumn(12)
        w12.width = 2;
        const w14 = worksheet.getColumn(14)
        w14.width = 2;
        const w15 = worksheet.getColumn(15)
        w15.width = 2;
        const w17 = worksheet.getColumn(17)
        w17.width = 1;
        const w18 = worksheet.getColumn(18)
        w18.width = 1;
        const w24 = worksheet.getColumn(23)
        w24.width = 1;
        const w25 = worksheet.getColumn(25)
        w25.width = 1;
        const w26 = worksheet.getColumn(26)
        w26.width = 2;
        const w28 = worksheet.getColumn(28)
        w28.width = 1;
        const w30 = worksheet.getColumn(30)
        w30.width = 1;
        const w33 = worksheet.getColumn(33)
        w33.width = 2;
        const w35 = worksheet.getColumn(35)
        w35.width = 1;

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
        </div>
      </>
    );
}