import React from 'react';
const ExcelJs = require("exceljs");

function ExportExcelButton (){
  function onClick(){
    const workbook = new ExcelJs.Workbook(); // 創建試算表檔案
    const sheet = workbook.addWorksheet('Page1',{views: [{showGridLines: false}]}); //在檔案中新增工作表 參數放自訂名稱

	
		sheet.addTable({
			name: 'Customer_Object',  // 表格內看不到的，算是key值，讓你之後想要針對這個table去做額外設定的時候，可以指定到這個table
			ref: 'A1',
			columns: [{name:'XXXXX Limited'}],
			rows: [['ABC Plaze,'],['Hong Kong, Hong Kong'],['Website:'],['Phone:'],[''],['Prepare by: Testing'],[''],
			['Customer'],['Testing solutions Limited:'],['Testing Wan'],['testing@testing.com'],['+852 12311923'],['General Manager ']]
		});
		
		sheet.addTable({
			name: 'Data_Object',  // 表格內看不到的，算是key值，讓你之後想要針對這個table去做額外設定的時候，可以指定到這個table
			ref: 'D3', // 從A1開始
			headerRow: false,
			columns: [{name:'empty'},{name:'dataName'},{name:'dataDetail'}],
			rows: [['','DATA','8/7/2022'],['','INVOICE','INV-INO-202207'],['','CUSTOMER ID','AAA001'],['','VALID UNTIL','7/8/2022']]
		});
		
		sheet.addTable({
			name: 'Product_Array',  // 表格內看不到的，算是key值，讓你之後想要針對這個table去做額外設定的時候，可以指定到這個table
			ref: 'A16', // 從A1開始
			style: {
				showRowStripes: true
			},
			columns: [{name: 'DESCRIPTION'}, { name: 'second'}, { name: 'UNIT PRICE'}],
			rows: [['',''],['小明', '', '0987654321'],['小美' ,'' ,'0912345678']]
		});
		sheet.getColumn(1).width=50;
		sheet.getColumn(2).width=50;
		const objectTable=sheet.getTable('Data_Object')
		const column =objectTable.getColumn(1);
		column.style={font:{bold: true}};
		column.filterButton = true;
		objectTable.commit();
    // 表格裡面的資料都填寫完成之後，訂出下載的callback function
		// 異步的等待他處理完之後，創建url與連結，觸發下載
	  	workbook.xlsx.writeBuffer().then((content) => {
		const link = document.createElement("a");
	    const blobData = new Blob([content], {
	      type: "application/vnd.ms-excel;charset=utf-8;"
	    });
	    link.download = '測試的試算表.xlsx';
	    link.href = URL.createObjectURL(blobData);
	    link.click();
	  });
	}
  return (
      <button onClick={onClick}> Download </button>
    )
};
export default ExportExcelButton;