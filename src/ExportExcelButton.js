import React from 'react';
const ExcelJs = require("exceljs");
//input product data
const ArrayOfProductObject=[
	{
		description:'Testing Product 001 (0000-0000-0000-0000)',
		unitPrice:'10.00',
		quantity:'1',
		DiscountRate:15,
	},
	{
		description:'Testing Product 002 (0000-0000-0000-0000)',
		unitPrice:'33.69',
		quantity:'1',
		DiscountRate:15,
	},
	{
		description:'Testing Product 003 (0000-0000-0000-0000)',
		unitPrice:'21.13',
		quantity:'1',
		DiscountRate:12,
	},
]
// discount amount
let discount=0.2;
let exchangeRate=7.9;
const titleFont={
	bold:true,
	color: {argb:'FFFFFFFF'},
};
function ExportExcelButton (){
  function onClick(){
    const workbook = new ExcelJs.Workbook(); 
    const sheet = workbook.addWorksheet('Page1',{views: [{showGridLines: false}]});
	
	sheet.getColumn(1).width=40;
	sheet.getColumn(2).width=40;
	sheet.getColumn(3).width=15;
	sheet.getColumn(4).width=5;
	sheet.getColumn(5).width=1;
	sheet.getColumn(6).width=12;
	sheet.getColumn(7).width=12;
	sheet.getColumn(8).width=22;
  	const row=sheet.getRow(1);
	row.height=42;
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
	function addContextAndAlign(cell,context,aligns){
		sheet.getCell(cell).value=context;
		sheet.getCell(cell).alignment={vertical: 'middle',horizontal: aligns};
	}
	function totalAmount(){
		let sum=0;
		for(let i=0;i<ArrayOfProductObject.length;i++)
		{sum+=parseFloat(ArrayOfProductObject[i].unitPrice)*parseFloat(ArrayOfProductObject[i].quantity).toFixed(2);}
		return sum;
	}
		//Customer Object
		Merge('A1','C1','XXXXX Limited')
		addContextAndAlign('A1','XXXXX Limited','left')
		sheet.getCell('A1').alignment = { 
			indent: 5,
			vertical: 'middle'
		};
		sheet.getCell('A1').font={
			size: 20
		};
		addTitle('A9','CUSTOMER');
		sheet.getCell('A2').value='Unit XXX XXX Centre, ';
		sheet.getCell('A3').value='Hong Kong, Hong Kong';
		sheet.getCell('A4').value='Website:';
		sheet.getCell('A5').value='Phone:';
		sheet.getCell('A7').value='Prepare by: Testing';
		sheet.getCell('A10').value='BiB Solutions';
		//Data_Object 
		Merge('E1','H1','Invoice ','right')
		sheet.getCell('E1').font={
			size: 26,
			color: { argb: 'FF8DB4E2' },
		};
		addContextAndAlign('E3','Data ','right')
		addContextAndAlign('E4','INVOICE# ','right')
		addContextAndAlign('E5','CUSTOMER ID ','right')
		addContextAndAlign('E6','VALID UNTIL ','right')
		addDataObject('H3','8/7/2022','center','thin');
		addDataObject('H4','INV-INO-202207','center','thin');
		addDataObject('H5','AAA001','center','thin');
		addDataObject('H6','7/8/2022','center','thin');
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
			Merge('A'+row,'B'+row,'','left');
			addTitle('A'+row,'TERMS AND CONDITIONS','left');
			sheet.getCell('A'+(row+1)).value='1.Payment Terms: 14 Day after invoice';
			sheet.getCell('A'+(row+2)).value='2.Exchange Rate :'+exchangeRate;
			sheet.getCell('A'+(row+3)).value='Bank info : XXX Bank /XX銀行';
			sheet.getCell('A'+(row+4)).value='                 XXX Company Limited ';
			sheet.getCell('A'+(row+5)).value='                 Acct No : 6549813';
			Merge('A'+(row+8),'C'+(row+8),'','left');
			for(let i=0;i<6;i++)
			{
				sheet.getCell('B'+(row+i)).border=
				{
					right: {style:'thin'},
				}
			}
				sheet.getCell('A'+(row+6)).border={
					bottom: {style:'thin'}
				}
				sheet.getCell('B'+(row+6)).border={
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