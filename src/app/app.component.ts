import { Component } from '@angular/core';
import * as XLSX from 'xlsx';
import { saveAs } from 'file-saver';
import { read, utils, WorkBook } from 'xlsx';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css'],
})
export class AppComponent {
  jsonData: any; // store json data
  fileName:any;
  sheetName:any;

  onFileSelected(event: any) {
    const file: File = event.target.files[0];
    this.fileName =file.name;
    const fileReader: FileReader = new FileReader();

    fileReader.onload = (e: any) => {
      const arrayBuffer: ArrayBuffer = e.target.result;
      const data: Uint8Array = new Uint8Array(arrayBuffer);
      const workbook: WorkBook = XLSX.read(data, { type: 'array' });

      const worksheetName: string = workbook.SheetNames[0];
      this.sheetName = worksheetName;
      const worksheet: any = workbook.Sheets[worksheetName];

      let jsonData: any = XLSX.utils.sheet_to_json(worksheet, {
        raw: true,
        defval: '',
        blankrows: false,
        header: 1,
      });
      console.log(jsonData);
      const updatedData = jsonData.map((row: any, index: number) => {
        if (index === 0)
          return { ...row, ValidationStatus: 'ValidationStatus' };
        if (
          row[0] === 'WLP_LIFE_PRO_MAXIMA' &&
          row[1] === 18 &&
          row[2] === 'F'
        ) {
          return { ...row, ValidationStatus: 'PASS' };
        }
        return row;
      });
      jsonData = updatedData;
      console.log('Update Data :', jsonData);
      this.jsonData = jsonData;
    };

    fileReader.readAsArrayBuffer(file);
  }

  // create and upload excel file
  createExcelFile() {
    console.log('create excel file');
    console.log(this.jsonData)
    const worksheet: XLSX.WorkSheet = XLSX.utils.json_to_sheet(this.jsonData, {
      skipHeader: true,
    });
    const workbook: XLSX.WorkBook = {
      Sheets: { [this.sheetName]: worksheet },
      SheetNames: [this.sheetName],
    };
    const excelBuffer: any = XLSX.write(workbook, {
      bookType: 'xlsx',
      type: 'array',
    });
    this.saveExcelFile(excelBuffer, `update ${this.fileName}`);

    console.log('create excel  file finished');
  }

  // save & download excel file
  saveExcelFile(buffer: any, fileName: string) {
    const data: Blob = new Blob([buffer], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8',
    });
    saveAs(data, `${fileName}_export_${new Date().getTime()}.xlsx`);
  }
}
