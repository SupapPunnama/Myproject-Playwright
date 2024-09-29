const { test: Delete_Activity_News } = require('@playwright/test');
const ExcelJS_Delete_Activity_News = require('exceljs');

Delete_Activity_News.only("Delete Activity News", async ({ page }) => {
    Delete_Activity_News.setTimeout(250000);
    const workbook = new ExcelJS_Delete_Activity_News.Workbook();
    await workbook.xlsx.readFile("C:\\Users\\Vivo\\Desktop\\Test_Project\\tests\\13_Data_Delete_Activity_News.xlsx");

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

    await page.goto("http://localhost:8083/sci_mju_lifelonglearning/course/public/list_activity");

    const worksheet = workbook.getWorksheet(1);
    let row = 1; // เริ่มต้นที่แถวที่ 2

    for (let round = 0; round < 2; round++) { // แก้ไขเป็นลูป for ที่ทำงานรอบเดียว

        const button = await page.$("//tbody/tr[1]/td[5]/button[1]");
        await button.scrollIntoViewIfNeeded();
        await button.click();

        page.once('dialog', async dialog => {
            const alertMessage = dialog.message();
            console.log('Alert message:', alertMessage);
            worksheet.getCell(`D${row}`).value = alertMessage;
            //await dialog.accept();//OK
            await dialog.dismiss(); // กด Cancel
            await workbook.xlsx.writeFile("C:\\Users\\Vivo\\Desktop\\Test_Project\\tests\\13_Data_Delete_Activity_News.xlsx");
        });

        const cellD = worksheet.getCell(`D${row}`).value;
        const cellC = worksheet.getCell(`C${row}`).value;
        worksheet.getCell(`F${row}`).value = cellD === cellC ? 'True' : 'False';
        await workbook.xlsx.writeFile("C:\\Users\\Vivo\\Desktop\\Test_Project\\tests\\13_Data_Delete_Activity_News.xlsx");

        row++; // ไปยังแถวถัดไป
    }
});
