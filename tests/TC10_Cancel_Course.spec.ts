const { test: Cancel_Request_open_course } = require('@playwright/test');
const ExcelJS_Cancel_Request_open_course = require('exceljs');

Cancel_Request_open_course.only("Cancel Request open course", async ({ page }) => {
    Cancel_Request_open_course.setTimeout(100000);

    const workbook = new ExcelJS_Cancel_Request_open_course.Workbook();
    await workbook.xlsx.readFile("C:\\Users\\Vivo\\Desktop\\Test_Project\\tests\\10_Data_Cancel_Course.xlsx");

    await page.goto('http://localhost:8083/sci_mju_lifelonglearning/', { waitUntil: 'domcontentloaded' });
    await page.waitForFunction(() => document.querySelector('title')?.textContent === 'Science MJU LifeLong Learning');
    await page.goto('http://localhost:8083/sci_mju_lifelonglearning/loginLecturer');

    let username = "Lecturer01";
    await page.waitForSelector("//input[@type='text'][@name='username']", { visible: true });
    await page.fill("//input[@type='text'][@name='username']", username);

    let password = "1234";
    await page.waitForSelector("//input[@type='password'][@name='password']", { visible: true });
    await page.fill("//input[@type='password'][@name='password']", password);

    const submit = await page.$("//input[@type='submit' and @value='เข้าสู่ระบบ']");
    await submit.scrollIntoViewIfNeeded();
    await submit.click();

    await page.goto('http://localhost:8083/sci_mju_lifelonglearning/lecturer/Lecturer01/list_request_open_course');
    await page.waitForTimeout(1000);

    const worksheet = workbook.getWorksheet(1);

    let row = 2; // เริ่มต้นที่แถวที่ 2
    let round = 0; // ตั้งค่าตัวนับรอบ

    while (round < 2) {

    const button = await page.$("//body[1]/div[1]/table[1]/tbody[1]/tr[1]/td[1]/div[1]/div[1]/table[1]/tbody[1]/tr[1]/td[7]/div[1]/div[2]/button[1]");
    //await button.scrollIntoViewIfNeeded();
    await button.click();

    page.once('dialog', async dialog => {
        const alertMessage = dialog.message();
        console.log('Alert message:', alertMessage);
        worksheet.getCell(`E${row}`).value = alertMessage;
        //await dialog.accept();//OK
         await dialog.dismiss();//Cancel
         await workbook.xlsx.writeFile("C:\\Users\\Vivo\\Desktop\\Test_Project\\tests\\10_Data_Cancel_Course.xlsx");
    });
    const cellD = worksheet.getCell(`D${row}`).value;
    const cellE = worksheet.getCell(`E${row}`).value;
    worksheet.getCell(`F${row}`).value = cellD === cellE ? 'True' : 'False';
    await workbook.xlsx.writeFile("C:\\Users\\Vivo\\Desktop\\Test_Project\\tests\\10_Data_Cancel_Course.xlsx");

    row++; // ไปยังแถวถัดไป
    round++; // เพิ่มตัวนับรอบ
    continue;
 }    
});
