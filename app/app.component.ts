import { Component, ViewChild, ElementRef } from '@angular/core';
import * as XLSX from 'xlsx';

//import { Observable } from 'rxjs/Observable';
@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css'],
})
export class AppComponent {
  @ViewChild('table1') table: ElementRef;
  responseheader: any;
  responsedata: any;
  reportname: any = 'Energy Cost Report';
  response: any = {
    data: [
      {
        Date: '2022-09-19',
        Timestamp: '23:00 to 24:00',
        'Main Incomer': [
          {
            'Energy Consumption (kWh)': '3893',
            'Cost (Rs.)': '3893',
          },
        ],
        rowSpan: 2,
      },
      {
        Date: '2022-09-19',
        Timestamp: '22:00 to 23:00',
        'Main Incomer': [
          {
            'Energy Consumption (kWh)': '3823',
            'Cost (Rs.)': '3823',
          },
        ],
      },
      {
        Date: '2022-09-19',
        Timestamp: '21:00 to 22:00',
        'Main Incomer': [
          {
            'Energy Consumption (kWh)': '4324',
            'Cost (Rs.)': '4324',
          },
        ],
      },
      {
        Date: '2022-09-19',
        Timestamp: '20:00 to 21:00',
        'Main Incomer': [
          {
            'Energy Consumption (kWh)': '4659',
            'Cost (Rs.)': '4659',
          },
        ],
      },
      {
        Date: '2022-09-19',
        Timestamp: '19:00 to 20:00',
        'Main Incomer': [
          {
            'Energy Consumption (kWh)': '4666',
            'Cost (Rs.)': '4666',
          },
        ],
      },
      {
        Date: '2022-09-19',
        Timestamp: '18:00 to 19:00',
        'Main Incomer': [
          {
            'Energy Consumption (kWh)': '4408',
            'Cost (Rs.)': '4408',
          },
        ],
      },
      {
        Date: '2022-09-19',
        Timestamp: '17:00 to 18:00',
        'Main Incomer': [
          {
            'Energy Consumption (kWh)': '4640',
            'Cost (Rs.)': '4640',
          },
        ],
      },
      {
        Date: '2022-09-19',
        Timestamp: '16:00 to 17:00',
        'Main Incomer': [
          {
            'Energy Consumption (kWh)': '4884',
            'Cost (Rs.)': '4884',
          },
        ],
      },
      {
        Date: '2022-09-19',
        Timestamp: '15:00 to 16:00',
        'Main Incomer': [
          {
            'Energy Consumption (kWh)': '4961',
            'Cost (Rs.)': '4961',
          },
        ],
      },
      {
        Date: '2022-09-19',
        Timestamp: '14:00 to 15:00',
        'Main Incomer': [
          {
            'Energy Consumption (kWh)': '4883',
            'Cost (Rs.)': '4883',
          },
        ],
      },
      {
        Date: '2022-09-19',
        Timestamp: '13:00 to 14:00',
        'Main Incomer': [
          {
            'Energy Consumption (kWh)': '5089',
            'Cost (Rs.)': '5089',
          },
        ],
      },
      {
        Date: '2022-09-19',
        Timestamp: '12:00 to 13:00',
        'Main Incomer': [
          {
            'Energy Consumption (kWh)': '4705',
            'Cost (Rs.)': '4705',
          },
        ],
      },
      {
        Date: '2022-09-19',
        Timestamp: '11:00 to 12:00',
        'Main Incomer': [
          {
            'Energy Consumption (kWh)': '4935',
            'Cost (Rs.)': '4935',
          },
        ],
      },
      {
        Date: '2022-09-19',
        Timestamp: '10:00 to 11:00',
        'Main Incomer': [
          {
            'Energy Consumption (kWh)': '4567',
            'Cost (Rs.)': '4567',
          },
        ],
      },
      {
        Date: '2022-09-19',
        Timestamp: '09:00 to 10:00',
        'Main Incomer': [
          {
            'Energy Consumption (kWh)': '4378',
            'Cost (Rs.)': '4378',
          },
        ],
      },
      {
        Date: '2022-09-19',
        Timestamp: '08:00 to 09:00',
        'Main Incomer': [
          {
            'Energy Consumption (kWh)': '4367',
            'Cost (Rs.)': '4367',
          },
        ],
      },
      {
        Date: '2022-09-19',
        Timestamp: '07:00 to 08:00',
        'Main Incomer': [
          {
            'Energy Consumption (kWh)': '2877',
            'Cost (Rs.)': '2877',
          },
        ],
      },
    ],
    headers: [
      {
        name: 'Time',
        subheaders: [],
        displayName: 'Time',
        rowSpan: 2,
      },
      {
        name: 'Main Incomer',
        displayName: 'Main Incomer',
        subheaders: [
          {
            name: 'Energy Consumption (kWh)',
            displayName: 'Energy Consumption (kWh)',
          },
          {
            name: 'Cost (Rs.)',
            displayName: 'Cost (Rs.)',
          },
        ],
        rowSpan: 1,
        colSpan: 2,
      },
    ],
  };
  constructor() {
    this.responseheader = this.response['headers'];
    this.responsedata = this.response['data'];
    console.log('responseheader', this.responseheader);
    console.log('responsedata', this.responsedata);
  }
  fireEvent() {
    const ws: XLSX.WorkSheet = XLSX.utils.table_to_sheet(
      this.table.nativeElement
    );

    /* new format */
    var fmt = '0.00';
    /* change cell format of range B2:D4 */
    var range = { s: { r: 1, c: 1 }, e: { r: 2, c: 100000 } };
    for (var R = range.s.r; R <= range.e.r; ++R) {
      for (var C = range.s.c; C <= range.e.c; ++C) {
        var cell = ws[XLSX.utils.encode_cell({ r: R, c: C })];
        if (!cell || cell.t != 'n') continue; // only format numeric cells
        cell.z = fmt;
      }
    }
    const wb: XLSX.WorkBook = XLSX.utils.book_new();
    //  wb.write('A2', 'Insert an image in a cell:')
    XLSX.utils.sheet_add_json(ws, ['Insert an image in a cell:']);
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
    var fmt = '@';

    //ws.insert_image('B2', 'python.png')
    //  wb.Sheets['Sheet1']['F'] = fmt;

    /* save to file */
    XLSX.writeFile(wb, this.reportname + '.xls');
  }
}
