import React from 'react';
const ExcelJs = require("exceljs");
//input data
const product_Description=['Testing Product 001 (0000-0000-0000-0000)','Testing Product 002 (0000-0000-0000-0000)',
				'Testing Product 003 (0000-0000-0000-0000)'];
const product_UnitPrice=['10.00','33.69','21.13'];
const product_QTY=['1','1','1'];

const titleFont={
	bold:true,
	color: {argb:'FFFFFFFF'},
};
const linkStyle = {
	underline: true,
	color: { argb: 'FF0000FF' },
};
function ExportExcelButton (){
  function onClick(){
    const workbook = new ExcelJs.Workbook(); 
    const sheet = workbook.addWorksheet('Page1',{views: [{showGridLines: false}]});
	
	sheet.getColumn(1).width=35;
	sheet.getColumn(2).width=35;
	sheet.getColumn(3).width=10;
	sheet.getColumn(4).width=6;
	sheet.getColumn(6).width=15;
  	const row=sheet.getRow(1);
	row.height=40;
	function Merge(start,end,content,align){
		sheet.mergeCells(start+':'+end);
		sheet.getCell(start).value=content;
		sheet.getCell(start).alignment={vertical: 'middle',horizontal: align};
	}

	function addDataObject(destination,content,align,thickness){
		sheet.getCell(destination).border = {
			top: {style:thickness},
			left: {style:thickness},
			bottom: {style:thickness},
			right: {style:thickness}
		};
			sheet.getCell(destination).value=content;
			sheet.getCell(destination).alignment={horizontal:align};
		}
	function fillBackgroundColor(destination,color){
		sheet.getCell(destination).fill={  
			type:'pattern',
			pattern:'solid',
			fgColor:{argb: color},
			bgColor:{argb: color}
		};
	}
	function addTitle(destination,context,align){
		sheet.getCell(destination).value=context;
		fillBackgroundColor(destination,'FF1F497D');
		sheet.getCell(destination).font=titleFont;
		sheet.getCell(destination).alignment={horizontal: align};
	}

	function addProductDetail(destination,content,align){
		sheet.getCell(destination).border = {
			left: {style:'thin'},
			right: {style:'thin'}
		};
		sheet.getCell(destination).value=content;
		sheet.getCell(destination).numFmt='00.00'
		sheet.getCell(destination).alignment={vertical: 'middle',horizontal:align};
	}	
	function fillInRow(row,color){
		fillBackgroundColor('A'+row,color);
		fillBackgroundColor('C'+row,color);
		fillBackgroundColor('D'+row,color);
		fillBackgroundColor('F'+row,color);
	}
	function addProductArray(description,unitPrice,QTY,amount,row){
		
		Merge('A'+row,'B'+row,description,'left');
		addProductDetail('C'+row,unitPrice,'right');
		Merge('D'+row,'E'+row,QTY,'center');
		addProductDetail('F'+row,amount,'right');
		if(row%2===1)
			{
				fillInRow(row,'FFF2F2F2');
				fillBackgroundColor('D'+row,'FFFFFFFF');
			}
	}
	function totalAmount(){
		let sum=0;
		for(let i=0;i<ArrayOfProductObject.length;i++)
		{sum+=parseFloat(ArrayOfProductObject[i].unitPrice)*parseFloat(ArrayOfProductObject[i].quantity).toFixed(2);}
		return sum;
	}
		//Customer Object
		sheet.getCell('A1').value='XXXXX Limited';
		sheet.getCell('A1').font={
			size: 20,
			color: { argb: 'FF366092' },
		};
		addTitle('A9','CUSTOMER');
		
		sheet.getCell('A1').alignment={vertical: 'middle',horizontal: "center"};
		sheet.getCell('A2').value='ABC Plaze,';
		sheet.getCell('A3').value='Hong Kong, Hong Kong';
		sheet.getCell('A4').value='Website:';
		sheet.getCell('A5').value='Phone:';
		sheet.getCell('A7').value='Prepare by: Testing';
		sheet.getCell('A10').value='Testing solutions Limited:';
		sheet.getCell('A11').value='Testing Wan';
		sheet.getCell('A12').value={ text: 'testing@testing.com', hyperlink: 'testing@testing.com' };
		sheet.getCell('A12').font=linkStyle;
		sheet.getCell('A13').value='+852 12311923';
		sheet.getCell('A14').value='General Manager ';
		//Data_Object 
		Merge('D1','F1','Invoice ','right')
		sheet.getCell('D1').font={
			size: 26,
			color: { argb: 'FF8DB4E2' },
		};
		Merge('D3','E3','Data ','right')
		Merge('D4','E4','INVOICE# ','right')
		Merge('D5','E5','CUSTOMER ID ','right')
		Merge('D6','E6','VALID UNTIL ','right')
		addDataObject('F3','8/7/2022','center','medium');
		addDataObject('F4','INV-INO-202207','center','medium');
		addDataObject('F5','AAA001','center','medium');
		addDataObject('F6','7/8/2022','center','medium');
		//add Product Array title	
		Merge('A16','B16','Invoice ','left');
		addTitle('A16','DESCRIPTION','left');
		addTitle('C16','UNIT PRICE','center');
		Merge('D16','E16','QTY ','center');
		addTitle('D16','QTY','center');
		addTitle('F16','AMOUNT','center');
		//add Product Array context
		addProductArray('','','','',17);
		fillBackgroundColor('D17','FFF2F2F2');
		for(let i=0;i<ArrayOfProductObject.length;i++){
			addProductArray(ArrayOfProductObject[i].description,ArrayOfProductObject[i].unitPrice,ArrayOfProductObject[i].quantity,
				parseFloat(ArrayOfProductObject[i].unitPrice)*parseFloat(ArrayOfProductObject[i].quantity).toFixed(2),18+i);
			const row=sheet.getRow(18+i);
			row.height=25;
			if(i===ArrayOfProductObject.length-1){
				addProductArray('','','','',18+i+1);
				fillInRow(18+i+1,'FFD9D9D9');
				printTermsAndConditions(18+i+3);
				printBillingInfo(18+i+2);
			}
		}
		//TermsAndConditions
		function printTermsAndConditions(row){
			Merge('A'+row,'C'+row,'','left');
			addTitle('A'+row,'TERMS AND CONDITIONS','left');
			Merge('A'+(row+1),'C'+(row+1),'1.Payment Terms: 30 Day after invoice','left');
			Merge('A'+(row+2),'C'+(row+2),'2.Deposit: 1 monthly consumption(Waived)','left');
			Merge('A'+(row+3),'C'+(row+3),'3.Please fax or mail the signed price quote to the address above','left');
			Merge('A'+(row+4),'C'+(row+4),'Customer Acceptance(sign below);','left');
			sheet.getCell('A'+(row+4)).font={italic:true};
			Merge('A'+(row+5),'C'+(row+5),'Bank info:XX Bank//XX 銀行','left');
			Merge('A'+(row+6),'C'+(row+6),'','left');
			sheet.getCell('A'+(row+6)).value='              Testing Limited'; 
			sheet.getCell('A'+(row+6)).font={bold: true};
			Merge('A'+(row+7),'C'+(row+7),'','left');
			sheet.getCell('A'+(row+7)).value='          Acot No:12345678';
			Merge('A'+(row+8),'C'+(row+8),'','left');
			for(let i=0;i<8;i++)
			{
				sheet.getCell('A'+(row+i)).border=
				{
					right: {style:'thin'}
				}
			}
				sheet.getCell('A'+(row+8)).border={
					right: {style:'thin'},
					bottom: {style:'thin'}
				}
			}
		//Billing info
		function printBillingInfo(row){
			sheet.getCell('E'+row).value='Subtotal';
			sheet.getCell('F'+row).value=totalAmount();
			sheet.getCell('F'+row).numFmt='_("USD"* #,##0.00_);_("USD"* (#,##0.00);_("$"* "-"??_);_(@_)';
			sheet.getCell('E'+(row+1)).value='Off';
			addDataObject('F'+(row+1),discount,'right','thin');
			sheet.getCell('F'+(row+1)).numFmt='00.000%'
			sheet.getCell('E'+(row+2)).value='Discount';
			sheet.getCell('E'+(row+2)).border={
				bottom:{style:'double'},
			};
			addDataObject('F'+(row+2),sheet.getCell('F'+row).value*discount,'right','thin');
			sheet.getCell('F'+(row+2)).numFmt='_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)';
			sheet.getCell('F'+(row+2)).border={
				top: {style:'thin'},
				left: {style:'thin'},
				right: {style:'thin'},
				bottom:{style:'double'},
			};
			sheet.getCell('E'+(row+3)).value='TOTAL';
			sheet.getCell('E'+(row+3)).font={bold: true,};
			sheet.getCell('F'+(row+3)).value=sheet.getCell('F'+row).value-sheet.getCell('F'+(row+2)).value;
			sheet.getCell('F'+(row+3)).numFmt='_("USD"* #,##0.00_);_("USD"* (#,##0.00);_("$"* "-"??_);_(@_)';
		}

    // output
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
