import { Component } from '@angular/core';
import { of } from 'rxjs';
import * as XLSX from 'xlsx';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})
export class AppComponent {
  title = 'json-excel';
  selectedFile:any;
  //json:any;
  json = [{rollno:1, name:"Abhijit"}];

  onFileChanged(event:Event) {
    let element = <HTMLInputElement>event.target;
    if(element != null)
    {
      this.selectedFile = element.files != null ? element.files[0] : null;
      const fileReader = new FileReader();
      fileReader.readAsText(this.selectedFile, "UTF-8");
      fileReader.onload = () => {
      if(fileReader.result != null)
        this.json = JSON.parse(fileReader.result?.toString());
      }
      fileReader.onerror = (error) => {
        console.log(error);     
       } 
    }
  }

  static toExportFileName(excelFileName: string): string {
    return `${excelFileName}_export_${new Date().getTime()}.xlsx`;
  }

  export(){
    console.log(this.json);
    this.exportAsExcelFile(this.json, "Sample.xlsx");
  }

  public exportAsExcelFile(json: any[], excelFileName: string): void {
    const worksheet: XLSX.WorkSheet = XLSX.utils.json_to_sheet(json);
    const workbook: XLSX.WorkBook = {Sheets: {'data': worksheet}, SheetNames: ['data']};
    XLSX.writeFile(workbook, AppComponent.toExportFileName(excelFileName));
  }

}
