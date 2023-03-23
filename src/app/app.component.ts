import { Component } from '@angular/core';
import { ExcelService } from './excel-service';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})
export class AppComponent {
  title = 'excel-generator';

  constructor(private excelService: ExcelService) {}

  downloadExcel() {
    this.excelService.download(); 
  }
}
