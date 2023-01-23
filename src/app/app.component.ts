import { Component } from '@angular/core';
import * as xlsx from 'xlsx';
import { IUser } from './models/user.model';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})
export class AppComponent {
  title = 'angular-import-export-excel';

  excelData: IUser[] | any[] = [];
  fileName = 'ExcelSheet.xlsx';

  constructor() { }

  ReadExcel(event: any) {

    let file = event.target.files[0];
    let fileReader = new FileReader();
    fileReader.readAsBinaryString(file);

    fileReader.onload = (e) => {

      var workBook = xlsx.read(fileReader.result, { type: "binary" });
      var sheetNames = workBook.SheetNames;

      this.excelData = xlsx.utils.sheet_to_json(workBook.Sheets[sheetNames[0]]);
    }

  }


  ExportExcel() {

    // pass here the table id

    let element = document.getElementById("tableUsers");
    const ws: xlsx.WorkSheet = xlsx.utils.table_to_sheet(element);

    // generate workBook and add WorkSheet

    const wb: xlsx.WorkBook = xlsx.utils.book_new();
    xlsx.utils.book_append_sheet(wb, ws, "Feuil 1");

    // save to file

    xlsx.writeFile(wb, this.fileName)
  }
}
