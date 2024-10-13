const { test: Add_Activity_News } = require('@playwright/test');
const ExcelJS_Add_Activity_News = require('exceljs');

Add_Activity_News.only("Add Activity News", async ({ page }) => {
    Add_Activity_News.setTimeout(400000);
    const workbook = new ExcelJS_Add_Activity_News.Workbook();
    await workbook.xlsx.readFile("C:\\Users\\Vivo\\Desktop\\Test_Project\\tests\\11_Data_Add_Activity_News.xlsx");

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
    await submit.click();

    await page.goto('http://localhost:8083/sci_mju_lifelonglearning/course/public/add_activity');

    const worksheet = workbook.getWorksheet(1);
    let row = 2; // เริ่มต้นที่แถวที่ 2
    let round = 0; // ตั้งค่าตัวนับรอบ

    while (round < 22) {
        await page.reload();
        const ac_name = worksheet.getCell(`B${row}`).value; // ชื่อข่าวสาร
        if (ac_name) {
            await page.waitForSelector('#ac_name', { visible: true });
            await page.fill('#ac_name', ac_name.toString());
        }

        const ac_data = worksheet.getCell(`C${row}`).value; // ข้อมูลข่าวสาร
        if (ac_data) {
            await page.waitForSelector("//div[@class='ql-editor ql-blank']", { visible: true });
            await page.fill("//div[@class='ql-editor ql-blank']", ac_data.toString());
        }

        const ac_img = worksheet.getCell(`D${row}`).value; // รูปภาพข่าวสาร
        if (ac_img) {
            await page.waitForSelector('#ac_img', { visible: true });
            await page.setInputFiles('#ac_img', ac_img.toString());

            // รออะเลิท 3 วินาที
            try {
                const dialog = await page.waitForEvent('dialog', { timeout: 3000 });
                console.log(`พบอะเลิท: ${dialog.message()}`);
                worksheet.getCell(`G${row}`).value = dialog.message();
                await dialog.accept(); // หรือ dismiss ตามต้องการ
                await workbook.xlsx.writeFile("C:\\Users\\Vivo\\Desktop\\Test_Project\\tests\\11_Data_Add_Activity_News.xlsx");
                
                row++;
                round++;
                continue;

            } catch (error) {
                console.log('ไม่มีอะเลิท');
            }
        }

        // คลิกปุ่มบันทึก
        const nextBtn = await page.$("//input[@type='submit'][@value='บันทึก']");
        if (nextBtn) {
            await nextBtn.scrollIntoViewIfNeeded();
            await nextBtn.click();
            await page.waitForTimeout(1000);

            try {
                // รอ dialog หลังจากคลิกปุ่มบันทึก
                const dialog = await page.waitForEvent('dialog', { timeout: 3000 });
                console.log('Alert message:', dialog.message());
                worksheet.getCell(`G${row}`).value = dialog.message();
                await dialog.accept(); // หรือ dismiss ถ้ามีการยืนยัน
            } catch (error) {
                console.log('ไม่มีอะเลิทหลังจากบันทึก');
            }

            const invalidAcName = await page.textContent('#invalidAcName');
            const invalidAcDetail = await page.textContent('#invalidAcDetail');
            const invalidAcImg = await page.textContent('#invalidAcImg');

            if (!invalidAcName.trim() && !invalidAcDetail.trim() && !invalidAcImg.trim()) {
                console.log('ฟิลด์ทั้งหมดถูกต้อง บันทึกสำเร็จ');
                worksheet.getCell(`G${row}`).value = 'คุณแน่ใจหรือไม่ว่าต้องการเพิ่มข่าวสารนี้?';
            } else {
                console.log('ข้อผิดพลาดหน้า:', invalidAcName, invalidAcDetail, invalidAcImg);
                worksheet.getCell(`G${row}`).value = invalidAcName || invalidAcDetail || invalidAcImg;

                row++;
                round++;
                continue;
            }

            const valueE = worksheet.getCell(`E${row}`).value;
            const valueG = worksheet.getCell(`G${row}`).value;
            if (valueE === valueG) {
                console.log(`Row ${row}: True`);
                worksheet.getCell(`H${row}`).value = 'Pass';
            } else {
                console.log(`Row ${row}: False`);
                worksheet.getCell(`H${row}`).value = 'Fail';
            }

            await workbook.xlsx.writeFile("C:\\Users\\Vivo\\Desktop\\Test_Project\\tests\\11_Data_Add_Activity_News.xlsx");
        }

        row++; // ไปยังแถวถัดไป
        round++; // เพิ่มตัวนับรอบ
        continue;
    }

    await workbook.xlsx.writeFile("C:\\Users\\Vivo\\Desktop\\Test_Project\\tests\\11_Data_Add_Activity_News.xlsx");
});
