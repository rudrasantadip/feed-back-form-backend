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
        const worksheet = workbook.getWorksheet("Flutter Fiesta");
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
        console.log(worksheet.columns)
        await workbook.xlsx.writeFile(filename)
    }


    static async createExcelFile(data, filename) {
        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('Data');

        worksheet.columns = [
            { header: 'Name', key: 'name', width: 20 },
            { header: 'Age', key: 'age', width: 10 },
        ];

        // Assuming data is an array of objects
        data.forEach((row) => {
            worksheet.addRow(row);
        });

        await workbook.xlsx.writeFile(filename);

        console.log(worksheet.columns)
        console.log('Excel file created successfully:', filename);
    }

    static async readExcelFile(filename) {
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(filename);
        const worksheet = workbook.getWorksheet("Data");

        const rows = [];
        worksheet.eachRow((row) => {
            rows.push(row.values);
        });

        return rows;
    }
}

module.exports = ExcelService;
