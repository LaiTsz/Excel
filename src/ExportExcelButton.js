import React from 'react';
const ExcelJs = require("exceljs");
// exchange rate
let exchangeRate=7.9;
//input product data
const ArrayOfProductObject=[
	{
		description:'Alibaba Cloud - UID: 3165563165565656',
		unitPrice:'521.45',
		quantity:'1',
		DiscountRate:0.15,
	},
	{
		description:'Alibaba Cloud - UID: 5445646898989563',
		unitPrice:'12.54',
		quantity:'1',
		DiscountRate:0.15,
	},
	{
		description:'Alibaba Cloud - UID: 5445646898989563',
		unitPrice:'56183.12',
		quantity:'1',
		DiscountRate:0.15,
	},
	{
		description:'Azure - XX-03(sd54-56as-461e-33cd-12dsbd44)',
		unitPrice:'1386.15',
		quantity:'1',
		DiscountRate:0.12,
	},
]

const titleFont={
	bold:true,
	color: {argb:'FFFFFFFF'},
};
function ExportExcelButton (){
  function onClick(){
    const workbook = new ExcelJs.Workbook(); 
    const sheet = workbook.addWorksheet('Page1',{views: [{showGridLines: false}]});
	sheet.eachRow(function(row, rowNumber){
		row.eachCell( function(cell, colNumber){
			if(cell.value)
				row.getCell(colNumber).font = {size:10};
		});
	});
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
		fillBackgroundColor(destination,'FF3B4E87');
		sheet.getCell(destination).font=titleFont;
		sheet.getCell(destination).alignment={ wrapText: true,vertical: 'middle',horizontal: align};
	}

	function addProductDetail(destination,content,align){
		sheet.getCell(destination).border = {
			left: {style:'thin'},
			right: {style:'thin'}
		};
		if(content!==0){
		sheet.getCell(destination).value=content;
		sheet.getCell(destination).numFmt='0.00'
		sheet.getCell(destination).alignment={vertical: 'middle',horizontal:align};
		}
		
	}	
	function fillInRow(row,color){
		fillBackgroundColor('A'+row,color);
		fillBackgroundColor('C'+row,color);
		fillBackgroundColor('D'+row,color);
		fillBackgroundColor('F'+row,color);
		fillBackgroundColor('G'+row,color);
		fillBackgroundColor('H'+row,color);
	}
	function addProductArray(description,unitPrice,QTY,amount,discount,row){
		
		Merge('A'+row,'B'+row,description,'left');
		addProductDetail('C'+row,unitPrice,'center');
		Merge('D'+row,'E'+row,QTY,'center');
		addProductDetail('F'+row,discount,'center');
		sheet.getCell('F'+row).numFmt='00%'
		addProductDetail('G'+row,amount*discount,'center');
		addProductDetail('H'+row,amount*(1-discount),'center');
		if(row%2===1)
			{
				fillInRow(row,'FFF2F2F2');
			}
	}
	function addContextAndAlign(cell,context,aligns){
		sheet.getCell(cell).value=context;
		sheet.getCell(cell).alignment={vertical: 'middle',horizontal: aligns};
	}
	function totalAmount(){
		let sum=0;
		for(let i=0;i<ArrayOfProductObject.length;i++)
		{sum+=parseFloat(ArrayOfProductObject[i].unitPrice)*parseFloat(ArrayOfProductObject[i].quantity);}
		return sum;
	}
	function discountAmount(){
		let sum=0;
		for(let i=0;i<ArrayOfProductObject.length;i++)
		{sum+=parseFloat(ArrayOfProductObject[i].unitPrice)*parseFloat(ArrayOfProductObject[i].quantity)
			*(ArrayOfProductObject[i].DiscountRate);}
		return sum;
	}
		//Customer Object
		Merge('A1','C1','XXX Company Limited')
		addContextAndAlign('A1','XXX Company Limited','left')
		sheet.getCell('A1').alignment = { 
			indent: 5,
			vertical: 'middle'
		};
		sheet.getCell('A1').font={
			name: 'Arial',
			size: 20,
			color:{argb:'BF2C3A65'}
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
			name: 'Arial',
			size: 26,
			color: { argb: '997A8DC5' },
			bold:true
		};
		sheet.getCell('E1').alignment={vertical: 'bottom',horizontal:'right'};
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
		addTitle('F16','Discount Rate','center');
		addTitle('G16','Discount Amount','center');
		addTitle('H16','Discounted Price','center');
		
		//add Product Array context
		addProductArray('','','','','',17);
		fillBackgroundColor('D17','FFF2F2F2');
		addProductArray('Servier Period : 1 Jul 2022 - 31 Jul 2022','','','','',18);
		sheet.getRow(18).height=25;
		sheet.getCell('A18').font={underline: true,bold: true};
		addProductArray('','','','','',19);
		for(let i=0;i<11;i++){
			if(i<ArrayOfProductObject.length){
				addProductArray(ArrayOfProductObject[i].description,ArrayOfProductObject[i].unitPrice,ArrayOfProductObject[i].quantity,
				parseFloat(ArrayOfProductObject[i].unitPrice)*parseFloat(ArrayOfProductObject[i].quantity).toFixed(2),
				ArrayOfProductObject[i].DiscountRate,20+i);
				
			}
			else
			{
				addProductArray('','','','','',20+i);
			}
			const row=sheet.getRow(20+i);
			row.height=46;
			if(i===10){
				addProductArray('','','','','',20+i+1);
				fillInRow(20+i+1,'FFD9D9D9');
				printTermsAndConditions(20+i+4);
				printBillingInfo(20+i+2);
			}
		}
		//TermsAndConditions
		function printTermsAndConditions(row){
			Merge('A'+row,'B'+row,'','left');
			addTitle('A'+row,'TERMS AND CONDITIONS','left');
			sheet.getCell('A'+(row+1)).value='1.Payment Terms: 14 Day after invoice';
			sheet.getCell('A'+(row+2)).value='2.Exchange Rate : '+exchangeRate;
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
			addContextAndAlign('B'+row,'Subtotal : ','right')
			sheet.getCell('C'+row).value=totalAmount();
			sheet.getCell('C'+row).numFmt='#,##0.00;##0.00';
			sheet.getCell('G'+row).value=discountAmount();
			sheet.getCell('H'+row).value=totalAmount()-discountAmount();
			sheet.getCell('H'+row).numFmt='"USD "#,##0.00;##0.00';
			sheet.getCell('H'+(row+1)).value=(totalAmount()-discountAmount())*exchangeRate;
			sheet.getCell('H'+(row+1)).numFmt='"HKD "#,##0.00;##0.00';
			sheet.getCell('H'+(row+1)).font={size: 14}
			sheet.getCell('H'+(row+1)).border={
				bottom:{style:'double'},
			};
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