const { test: Approve_Request } = require('@playwright/test');
const ExcelJS_Approve_Request = require('exceljs');

Approve_Request.only("Approve Request", async ({ page }) => {
    Approve_Request.setTimeout(250000);
    const workbook = new ExcelJS_Approve_Request.Workbook();
    await workbook.xlsx.readFile("C:\\Users\\Vivo\\Desktop\\Test_Project\\tests\\14_Data_Approve_Request.xlsx");

    await page.goto('http://localhost:8083/sci_mju_lifelonglearning/', { waitUntil: 'domcontentloaded' });
    await page.waitForFunction(() => document.querySelector('title')?.textContent === 'Science MJU LifeLong Learning');
    await page.goto('http://localhost:8083/sci_mju_lifelonglearning/loginAdmin');

    let username = "seg123";
    await page.waitForSelector("//input[@type='text'][@name='username']", { visible: true });
    await page.fill("//input[@type='text'][@name='username']", username);

    let password = "1234";
    await page.waitForSelector("//input[@type='password'][@name='password']", { visible: true });
    await page.fill("//input[@type='password'][@name='password']", password);

    const submit = await page.$("//input[@type='submit' and @value='เข้าสู่ระบบ']");
    await submit.scrollIntoViewIfNeeded();
    await submit.click()

    
    await page.goto("http://localhost:8083/sci_mju_lifelonglearning/course/seg123/view_request_open_course/3");

    const worksheet = workbook.getWorksheet(1);
    let row = 1; // เริ่มต้นที่แถวที่ 2

    for (let round = 0; round < 2; round++) { 

        const button = await page.$("//input[@type='button'][@value='ยืนยันคำร้องขอ']");
        await button.scrollIntoViewIfNeeded();
        await button.click();

        page.once('dialog', async dialog => {
            const alertMessage = dialog.message();
            console.log('Alert message:', alertMessage);
            worksheet.getCell(`F${row}`).value = alertMessage;
            //await dialog.accept();//OK
            await dialog.dismiss(); // กด Cancel
            await workbook.xlsx.writeFile("C:\\Users\\Vivo\\Desktop\\Test_Project\\tests\\14_Data_Approve_Request.xlsx");
        });

        const cellD = worksheet.getCell(`D${row}`).value;
        const cellF = worksheet.getCell(`F${row}`).value;
        worksheet.getCell(`G${row}`).value = cellD === cellF ? 'True' : 'False';
        await workbook.xlsx.writeFile("C:\\Users\\Vivo\\Desktop\\Test_Project\\tests\\14_Data_Approve_Request.xlsx");

        row++; // ไปยังแถวถัดไป
    }
});
