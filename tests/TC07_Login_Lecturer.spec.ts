const { test: Lecturer } = require('@playwright/test');
const Lecturer_ExcelJS = require('exceljs');

Lecturer.only("Login Member user", async ({ page }) => {
    Lecturer.setTimeout(120000);

    const workbook = new Lecturer_ExcelJS.Workbook();
    await workbook.xlsx.readFile("C:\\Users\\Vivo\\Desktop\\Test_Project\\tests\\07_Data_Login_Lecturer.xlsx");

    await page.goto('http://localhost:8083/sci_mju_lifelonglearning/', { waitUntil: 'domcontentloaded' });
    await page.waitForFunction(() => document.querySelector('title')?.textContent === 'Science MJU LifeLong Learning');
    await page.goto('http://localhost:8083/sci_mju_lifelonglearning/loginLecturer');

    const worksheet = workbook.getWorksheet(1);
    let row = 2; // Start at row 2
    let round = 0;
    
    while (round < 2) { 
        await page.reload();

        const username = worksheet.getCell(`B${row}`).value;
        if (username) {
            await page.waitForSelector("//input[@type='text'][@name='username']", { visible: true, timeout: 6000 });
            await page.fill("//input[@type='text'][@name='username']", username.toString());
        }

        const password = worksheet.getCell(`C${row}`).value;
        if (password) {
            await page.waitForSelector("//input[@type='password'][@name='password']", { visible: true, timeout: 6000 });
            await page.fill("//input[@type='password'][@name='password']", password.toString());
        }

        const nextBtn = await page.$('input[type="submit"][value="เข้าสู่ระบบ"]');
        await nextBtn.scrollIntoViewIfNeeded();
        await nextBtn.click();

        await page.waitForTimeout(3000); 
        const currentUrl = page.url();
        if (currentUrl === 'http://localhost:8083/sci_mju_lifelonglearning/') {
            console.log('เข้าสู่ระบบสำเร็จ');
            worksheet.getCell(`F${row}`).value = 'เข้าสู่ระบบสำเร็จ';
            await page.click('a[href="/sci_mju_lifelonglearning/doLogout"]');  
            await page.goto('http://localhost:8083/sci_mju_lifelonglearning/loginLecturer', { waitUntil: 'domcontentloaded' });  
        } else {
            console.log('เข้าสู่ระบบไม่สำเร็จ');
            worksheet.getCell(`F${row}`).value = 'เข้าสู่ระบบไม่สำเร็จ';
            await page.goto('http://localhost:8083/sci_mju_lifelonglearning/loginLecturer', { waitUntil: 'domcontentloaded' });  
        }

        //เงื่อนไขการตรวจสอบ
        const valueF = worksheet.getCell(`F${row}`).value;
        const valueG = worksheet.getCell(`G${row}`).value;
            if (valueF === valueG) {
                console.log(`Row ${row}: True`);
                worksheet.getCell(`H${row}`).value = 'Pass';
            } else {
                console.log(`Row ${row}: False`);
                worksheet.getCell(`H${row}`).value = 'Fail';
            }
        await workbook.xlsx.writeFile("C:\\Users\\Vivo\\Desktop\\Test_Project\\tests\\07_Data_Login_Lecturer.xlsx");
        await page.waitForTimeout(1000);

        row++; 
        round++;
    }
});