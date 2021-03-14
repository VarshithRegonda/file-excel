import {Injectable} from '@angular/core'
import * as Excel  from 'exceljs'
import *  as fs from 'file-saver'

const EXCEL_TYPE = "application/vnd.openxmlformats-officedocument.spreedsheetml.sheet;UTF=8"
const EXCEL_EXTENTION = ".xlsx"

@Injectable({
    providedIn:'root'
})
export class ExcelService{
    constructor(){}

    public exportAsExcelFile(
         reportHeading:string,
         reportSubHeading: string,
         headersArray:any[],
         json:any[],
         footerData:any,
         excelFileName:string,
         sheetName:string



    ){
        const header= headersArray;
        const data = json;
       
        /*Create workbook and worksheet  */

        const workbook=new Excel.Workbook();
        workbook.creator='employee workbook'
        workbook.lastModifiedBy='employee coder'
        workbook.created=new Date()
        const worksheet=workbook.addWorksheet(sheetName)

        
        /* Add Header Row */
        worksheet.addRow([])
        worksheet.mergeCells('A1:'+this.numToAlpha(header.length-1)+1)
        worksheet.getCell('A1').value=reportHeading;
        worksheet.getCell('A1').alignment={horizontal:'center'}
        worksheet.getCell('A1').font={size:15,bold:true}
        



        if(reportSubHeading!==''){
            worksheet.addRow([])
            worksheet.mergeCells('A2:' +this.numToAlpha(header.length-1)+'1')
            worksheet.getCell('A2').value=reportSubHeading;
            worksheet.getCell('A2').alignment={horizontal:'center'}
            worksheet.getCell('A2').font={size:12,bold:false}
        }
        worksheet.addRow([])
        /* Add Header Row */
        const headerRow = worksheet.addRow(header)
        //cell style and border style 
        headerRow.eachCell((cell,index) =>{
            cell.fill={
                type: 'pattern',
                pattern:'solid',
                fgColor:{argb:'FFFFFF00'},
                bgColor:{argb:'FF00FFF'}
            };
            cell.border={top:{style:'thin'},left:{style:'thin'},bottom:{style:'thin'},right:{style:'thin'}}
            cell.font={size:12,bold:true}

            worksheet.getColumn(index).width=header[index-1].length<20 ? 20 : header[index-1].length;

        });
        //Get all columns from Json
        let columnArray:any[];
        for(const key in json){
            if(json.hasOwnProperty(key)){
                columnArray=Object.keys(json[key]);
            }
        }
        //Add Data and conditional Formating 
        data.forEach((element:any) => {
            const eachRow: any[]=[];
            columnArray.forEach((column)=>{
                eachRow.push(element[column])
            })
            if(element.isDeleted==='Y'){
                const deltaRow =worksheet.addRow(eachRow)
                deltaRow.eachCell((cell)=>{
                    cell.font={name:'Calbirl',family:4,size:11,bold:false,strike:true}

                });                                
            }else{
                worksheet.addRow(eachRow)
            }
        });  
             workbook.xlsx.writeBuffer().then((data:any)=>{
                const blob = new Blob([data], {
                    type:
                      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                  });
                let  url = window.URL.createObjectURL(blob)
                let a =document.createElement("a")
                document.body.appendChild(a)
                a.download='export.xlsx'
                a.click()
                window.URL.revokeObjectURL(url);
                a.remove();

                fs.saveAs(blob,excelFileName+EXCEL_EXTENTION)
             })


              
    }
    private numToAlpha(num:number){
        let alpha= " "
        for(;num>=0;num= parseInt((num/26).toString(),10)-1){
            alpha=String.fromCharCode(num%26+0x41)+alpha

    }
    return alpha
    }
}
