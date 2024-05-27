// excelService.js
const ExcelJS = require('exceljs');

class ExcelService {

    static async createWorkBook(filepath, worksheets) {
        const workbook = new ExcelJS.Workbook(); // creating a workbook

        //Creating the worksheets
        worksheets.forEach(sheet => {
            const worksheet = workbook.addWorksheet(sheet.sheetname) // creating sheets
            //Adding heading columns
            worksheet.columns=sheet.columns
          
        });
        await workbook.xlsx.writeFile(filepath); //Writing to the file
        console.log("File created successfully")
    }



    static async addData(data,filename,sheetname)
    {
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(filename);
        const worksheet = workbook.getWorksheet(sheetname);
        

        worksheet.addRow([data.MobileNumber,
            data.FullName,
            data.EmailAddress,
            data.EventName,
            data.LikedEventManagement,
            data.LikedEventOrganization,
            data.LikedValueAddition,
            data.ContentRating,
            data.ManagementRating,
            data.InformationRating,
            data.toRecommend,
            data.Remarks,
            data.Rating])
        // worksheet.addRow(data)
        await workbook.xlsx.writeFile(filename)
    }


    static async readExcelFile(filename,sheetname) {
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(filename);
        const worksheet = workbook.getWorksheet(sheetname);

        const rows = [];
        worksheet.eachRow((row) => {
            rows.push(row.values);
        });

        return rows;
    }
}

module.exports = ExcelService;
