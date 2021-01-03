import {Component} from '@angular/core';
import {ViewChild} from '@angular/core';
import {NgxCsvParser} from 'projects/ngx-csv-parser/src/public-api';
import {NgxCSVParserError} from 'projects/ngx-csv-parser/src/public-api';
import {jsPDF} from 'jspdf/dist/jspdf.debug';
import {formatCurrency, formatNumber} from '@angular/common';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.scss']
})

export class AppComponent {

  csvRecords: any[] = [];
  headers: any[] = [];
  filetype = 'application/json, .csv, .xml, .tsv, .xls';
  maxFileSize = 10000000;
  contents: any = null;
  filename: string;
  exportColumns: any[];
  _selectedColumns: any[];
  cols: any[];

  constructor(private ngxCsvParser: NgxCsvParser) {
    this.exportColumns = [
      {title: 'Gesellschaft', dataKey: 'Gesellschaft'},
      {title: 'Vertragsnummer', dataKey: 'Vertragsnummer'},
      {title: 'Sparte', dataKey: 'Sparte'},
      {title: 'Beginn', dataKey: 'Vertragsbeginn'},
      {title: 'Ende', dataKey: 'Vertragsende'},
      {title: 'Bruttobeitrag', dataKey: 'Bruttobeitrag'},
      {title: 'Zahlweise', dataKey: 'Zahlweise'}
    ];
  }

  @ViewChild('fileImportInput') fileImportInput: any;
  public myUploader(event, form) {
    console.log('Reading file...');
    for (const file of event.files) {
      const dataset = this.readCSVFile(file);
      console.log('onUpload: ', dataset);
    }
    form.clear();
  }

  private readFile(file: File) {
    const reader: FileReader = new FileReader();
    reader.onload = () => {
      console.log('readFile: ', reader.result);
      this.contents = reader.result;
    };
    reader.readAsText(file);
    this.filename = file.name;
  }

  private readCSVFile(file: File) {
    let isFirstRow = true;
    this.ngxCsvParser.parse(file, {header: true, delimiter: ';'})
    .pipe().subscribe((result: Array<any>) => {

      if (isFirstRow) {
        console.log('Headers', result);
        this.headers = result;
        isFirstRow = false;
      } else {
        console.log('Result', result);
        this.csvRecords = result;
      }
    }, (error: NgxCSVParserError) => {
      console.log('Error', error);
    });
  }

  initCols() {
    this.headers.forEach(header => {
        this.cols.push({field: header, header: header});
      }
    );
    console.log('Cols: ' + this.cols);
  }

  exportPdf1() {
    const doc = new jsPDF();
    doc.save('vertragsuebersicht.pdf');
  }

  exportPdf() {
    import('jspdf').then(jsPDF => {
      import('jspdf-autotable').then(x => {
        const doc = new jsPDF.default('l');
        const options = {
          align: 'left'
        };
        doc.text('Vertragsübersicht', 120, 20, options);
        doc.autoTable(
          this.exportColumns,
          this.csvRecords,
          {
            startY: 25,
            styles: {
              cellWidth: 'auto',
              minCellHeight: 9
            },
            headStyles: {
              fillColor: '#0091D5'
            },
            columnStyles: {
              3: {halign: 'center'},
              4: {halign: 'center'},
              5: {halign: 'right'},
              6: {halign: 'right'}
            },
            didParseCell: (data) => {
              if (data.section === 'body') {
                switch (data.column.dataKey) {
                  case 'Bruttobeitrag':
                    let content = data.cell.text[0];
                    content = content.replace(',', '.')
                    data.cell.text = formatCurrency(content, 'de-DE', '€');
                    break;
                }
              }
            }
          }
        );
        doc.save('vertragsuebersicht.pdf');
      });
    });
  }


  exportExcel() {
    import('xlsx').then(xlsx => {
      const worksheet = xlsx.utils.json_to_sheet(this.csvRecords);
      const workbook = {Sheets: {data: worksheet}, SheetNames: ['data']};
      const excelBuffer: any = xlsx.write(workbook, {bookType: 'xlsx', type: 'array'});
      this.saveAsExcelFile(excelBuffer, 'vertragsuebersicht');
    });
  }

  saveAsExcelFile(buffer: any, fileName: string): void {
    import('file-saver').then(FileSaver => {
      const EXCEL_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8';
      const EXCEL_EXTENSION = '.xlsx';
      const data: Blob = new Blob([buffer], {
        type: EXCEL_TYPE
      });
      FileSaver.saveAs(data, fileName + '_export_' + new Date().getTime() + EXCEL_EXTENSION);
    });
  }

}
