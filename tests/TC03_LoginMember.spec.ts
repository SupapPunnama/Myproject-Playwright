const { test: user } = require('@playwright/test');
const user_ExcelJS = require('exceljs');

user.only("Login Member user", async ({ page }) => {
    user.setTimeout(120000); // Increase the overall test timeout

    const workbook = new user_ExcelJS.Workbook();
    await workbook.xlsx.readFile("C:\\Users\\Vivo\\Desktop\\Test_Project\\tests\\03_Data_Loginmember.xlsx");

    await page.goto('http://localhost:8083/sci_mju_lifelonglearning/', { waitUntil: 'domcontentloaded' });
    await page.waitForFunction(() => document.querySelector('title')?.textContent === 'Science MJU LifeLong Learning');
    await page.goto('http://localhost:8083/sci_mju_lifelonglearning/loginMember');

    const worksheet = workbook.getWorksheet(1);
    let row = 2; // Start at row 2
    let round = 0;

    while (round < 6) { // Loop 3 rounds
        await page.reload();

        const username = worksheet.getCell(`B${row}`).value;
        if (username) {
            await page.waitForSelector('input[name="username"]', { visible: true, timeout: 6000 });
            await page.fill('input[name="username"]', username.toString());
        }

        const password = worksheet.getCell(`C${row}`).value;
        if (password) {
            await page.waitForSelector('input[name="password"]', { visible: true, timeout: 6000 });
            await page.fill('input[name="password"]', password.toString());
        }

        const nextBtn = await page.$('input[type="submit"][value="เข้าสู่ระบบ"]');
        await nextBtn.scrollIntoViewIfNeeded();
        await nextBtn.click();


        await page.waitForTimeout(3000); 
        const currentUrl = page.url();
        if (currentUrl === 'http://localhost:8083/sci_mju_lifelonglearning/') {
            console.log('เข้าสู่ระบบสำเร็จ');
            worksheet.getCell(`F${row}`).value = 'Teue';
            await page.click('a[href="/sci_mju_lifelonglearning/doLogout"]');  
            await page.goto('http://localhost:8083/sci_mju_lifelonglearning/loginMember', { waitUntil: 'domcontentloaded' });  
        } else {
            console.log('เข้าสู่ระบบไม่สำเร็จ');
            worksheet.getCell(`F${row}`).value = 'False';
            await page.goto('http://localhost:8083/sci_mju_lifelonglearning/loginMember', { waitUntil: 'domcontentloaded' });  
        }
        row++; 
        round++;
    }
    await workbook.xlsx.writeFile("C:\\Users\\Vivo\\Desktop\\Test_Project\\tests\\03_Data_Loginmember.xlsx");
});

