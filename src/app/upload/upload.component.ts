import { Component, OnInit } from '@angular/core';

import { Workbook } from 'exceljs';
import * as fs from 'file-saver';

@Component({
  selector: 'app-upload',
  templateUrl: './upload.component.html',
  styleUrls: ['./upload.component.css']
})
export class UploadComponent implements OnInit {
  
  data: product[] = [
    { id: 1, name: "Nivia Graffiti Basketball", brand: "Nivia", color: "Mixed", price: 391.00 },
    { id: 2, name: "Strauss Official Basketball", brand: "Strauss", color: "Orange", price: 391.00 },
    { id: 3, name: "Spalding Rebound Rubber Basketball", brand: "Spalding", color: "Brick", price: 675.00 },
    { id: 4, name: "Cosco Funtime Basket Ball, Size 6 ", brand: "Cosco", color: "Orange", price: 300.00 },
    { id: 5, name: "Nike Dominate 8P Basketball", brand: "Nike", color: "brick", price: 1295 },
    { id: 6, name: "Nivia Europa Basketball", brand: "Nivia", color: "Orange", price: 280.00 }
  ]
  constructor() { }
  
  ngOnInit() {
    
    let workbook = new Workbook();
  let worksheet = workbook.addWorksheet('ProductSheet');
    worksheet.columns = [
      { header: 'Id', key: 'id', width: 10 },
      { header: 'Name', key: 'name', width: 32 },
      { header: 'Brand', key: 'brand', width: 10 },
      { header: 'Color', key: 'color', width: 10 },
      { header: 'Price', key: 'price', width: 10, style: { font: { name: 'Arial Black', size:10} } },
    ];  
    this.data.forEach(e => {
      worksheet.addRow({id: e.id, name: e.name, brand:e.brand, color:e.color, price:e.price },"n");
    });
   
    workbook.xlsx.writeBuffer().then((data) => {
      let blob = new Blob([data], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      fs.saveAs(blob, 'ProductData.xlsx');
    })
   
  }
   
  
} 
export interface product {
  id: number,
  name: string
  brand: string,
  color: string,
  price:number
}

