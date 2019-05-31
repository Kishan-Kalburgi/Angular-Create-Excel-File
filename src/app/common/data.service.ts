import { Injectable } from '@angular/core';
import * as FileSaver from 'file-saver';
import * as XLSX from 'xlsx';
import { getViewData } from '@angular/core/src/render3/state';

const EXCEL_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8';
const EXCEL_EXTENSION = '.xlsx';

@Injectable({
  providedIn: 'root'
})
export class DataService {

  constructor() { }
  generateXLSX2() {
    const wb: XLSX.WorkBook = XLSX.utils.book_new();
    // XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');

  }
  gData = [{
		"Attachment": "/docp/supply-chain/v1/invoices/50493873/attachments/64385743;/docp/supply-chain/v1/invoices/50493873/attachments/64384546;/docp/supply-chain/v1/invoices/50493873/attachments/64385624;/docp/supply-chain/v1/invoices/50493873/attachments/64385675;/docp/supply-chain/v1/invoices/50493873/attachments/64385722;/docp/supply-chain/v1/invoices/50493873/attachments/64385437;",
		"InvoiceId": "50493873",
		"InvoiceNumber": "20034-1",
		"VendorNumber": "1070053076"
	},
	{
		"Attachment": "/docp/supply-chain/v1/invoices/145258599/attachments/179820775;",
		"InvoiceId": "145258599",
		"InvoiceNumber": "4607",
		"VendorNumber": "1070048215"
	}
]
  generateXLSX() {
    var counter = 2;
    const worksheet = XLSX.utils.json_to_sheet([
      { A: 'InvoiceId', B: 'InvoiceNumber', C: 'VendorNumber', D: 'Attachment' }
    ], {skipHeader: true });
this.gData.forEach(element => {
  let temp ={
    A : element["InvoiceId"],
    B : element["InvoiceNumber"],
    C : element["VendorNumber"]
  }

  let attachData = element.Attachment.split(';');
  for(let i = 68; i < (68 + attachData.length); i++ ) {
    temp[String.fromCharCode(i)] = attachData[i-68];
  }

  console.log(temp);

  var row = XLSX.utils.sheet_add_json(worksheet, [
    temp
  ], {skipHeader: true, origin: -1 });

  for(let i = 68; i < (68 + attachData.length); i++ ) {
    row[String.fromCharCode(i) + counter].l = { Target: attachData[i-68], Tooltip:"Find us @ SheetJS.com!" };
  }
  counter++;
});


    // /* Append row */
    // var row = XLSX.utils.sheet_add_json(worksheet, [
    //   { A: 4, B: 5, C: 6, D: 7, E: 8, F: 9, G: 0 }
    // ], {skipHeader: true, origin: -1 });

    // row['G2'].l = { Target:"http://sheetjs.com", Tooltip:"Find us @ SheetJS.com!" };
    // console.log(row);
    // /* Append row */
    // var row2 = XLSX.utils.sheet_add_json(worksheet, [
    //   { A: 4, B: 5, C: 6, D: 7, E: 8, F: 9, G: 0 }
    // ], {skipHeader: true, origin: -1 });

    // row2['G3'].l = { Target:"http://google.com", Tooltip:"Find us @ sdfmsdlm.com!" };


    const workbook: XLSX.WorkBook = { Sheets: { 'data': worksheet }, SheetNames: ['data'] };
    const excelBuffer: any = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
    this.saveAsExcelFile(excelBuffer, 'Sample');
  }
  private saveAsExcelFile(buffer: any, fileName: string): void {
    const data: Blob = new Blob([buffer], { type: EXCEL_TYPE });
    FileSaver.saveAs(data, fileName + '_export_' + new Date().getTime() + EXCEL_EXTENSION);
  }
}

