const { test: Edit_Profile, expect } = require('@playwright/test');
const ExcelJS2 = require('exceljs');

Edit_Profile.only("Edit Profile", async ({ page }) => {
    Edit_Profile.setTimeout(2000000); // Set test timeout to a higher value

    const workbook = new ExcelJS2.Workbook();
    await workbook.xlsx.readFile("C:\\Users\\Vivo\\Desktop\\Test_Project\\tests\\04_Data_EditProfile.xlsx");

    await page.goto('http://localhost:8083/sci_mju_lifelonglearning/', { waitUntil: 'load', timeout: 60000 });
    await page.waitForFunction(() => document.querySelector('title')?.textContent === 'Science MJU LifeLong Learning', { timeout: 60000 });
    await page.goto('http://localhost:8083/sci_mju_lifelonglearning/loginMember', { waitUntil: 'load', timeout: 60000 });

    await page.waitForSelector('input[name="username"]', { visible: true, timeout: 60000 });
    await page.fill('input[name="username"]', '1234');

    await page.waitForSelector('input[name="password"]', { visible: true, timeout: 60000 });
    await page.fill('input[name="password"]', '1234');

    const loginBtn = await page.$('input[type="submit"][value="เข้าสู่ระบบ"]');
    await loginBtn.scrollIntoViewIfNeeded();
    await Promise.all([
        page.waitForNavigation({ waitUntil: 'load', timeout: 60000 }),
        loginBtn.click()
    ]);

    const worksheet = workbook.getWorksheet(1);
    let row = 2; // Start at row 2
    let round = 0;

    // Handle dialogs globally
    page.on('dialog', async dialog => {
        const alertMessage = dialog.message();
        console.log('Alert message:', alertMessage);
        worksheet.getCell(`J${row}`).value = alertMessage;

        await dialog.accept();

        // Compare cell H and J, then store True or False in K
        const cellH = worksheet.getCell(`H${row}`).value;
        const cellJ = worksheet.getCell(`J${row}`).value;
        worksheet.getCell(`K${row}`).value = cellH === cellJ ? 'True' : 'False';

        // Save Excel changes
        await workbook.xlsx.writeFile("C:\\Users\\Vivo\\Desktop\\Test_Project\\tests\\04_Data_EditProfile.xlsx");

        row++; // Move to the next row after handling the dialog
        round++;
    });

    while (round <= 45) { // Loop 45 rounds
        const processFlag = worksheet.getCell(`L${row}`).value; // ตรวจสอบค่าในเซลล์ L{row}
        console.log(processFlag);

        if (processFlag !== 'Yes') {
            row++;
            round++;
            continue;
        }

        await page.reload({ waitUntil: 'load', timeout: 60000 });
        await page.goto('http://localhost:8083/sci_mju_lifelonglearning/member/1234/edit_profile', { waitUntil: 'load', timeout: 60000 });

        const firstName = worksheet.getCell(`B${row}`).value;
        if (firstName) {
            await page.waitForSelector('#firstName', { visible: true, timeout: 60000 });
            await page.fill('#firstName', firstName.toString());
        }

        const lastName = worksheet.getCell(`C${row}`).value;
        if (lastName) {
            await page.waitForSelector('//input[@type="text"][@name="lastName"]', { visible: true, timeout: 60000 });
            await page.fill('//input[@type="text"][@name="lastName"]', lastName.toString());
        }

        const birthdayCellValue = worksheet.getCell(`D${row}`).value;
        if (birthdayCellValue) {
            try {
                const dateParts = birthdayCellValue.split('/');
                if (dateParts.length === 3) {
                    const year = (parseInt(dateParts[2]) - 543).toString();
                    const month = dateParts[1].padStart(2, '0');
                    const day = dateParts[0].padStart(2, '0');
                    const formattedBirthday = `${year}-${month}-${day}`;
                    await page.waitForSelector('#datePicker', { visible: true, timeout: 60000 });
                    await page.fill('#datePicker', formattedBirthday);
                } else {
                    console.error(`Invalid date format in cell D${row}: ${birthdayCellValue}`);
                }
            } catch (error) {
                console.error(`Error parsing date value in cell D${row}: ${birthdayCellValue}`, error);
            }
        }

        const tel = worksheet.getCell(`E${row}`).value;
        if (tel) {
            await page.waitForSelector('#tel', { visible: true, timeout: 60000 });
            await page.fill('#tel', tel.toString());
        }

        const email = worksheet.getCell(`F${row}`).value;
        if (email) {
            await page.waitForSelector('#email', { visible: true, timeout: 60000 });
            await page.fill('#email', email.toString());
        }

        const education = worksheet.getCell(`G${row}`).value;
        if (education) {
            await page.waitForSelector('#education', { visible: true, timeout: 60000 });
            const optionValues = await page.$$eval('#education option', options =>
                options.map(option => option.textContent?.trim() || '')
            );
            if (optionValues.includes(education.trim())) {
                await page.selectOption('#education', { label: education.trim() });
            } else {
                console.log(`ตัวเลือก '${education}' ไม่มีในรายการ`);
            }
        }

        const nextButton = await page.$('input[type="submit"][value="บันทึก"]');
        await nextButton.scrollIntoViewIfNeeded();
        await nextButton.click();

        // Short wait for any possible dialog to appear
        await page.waitForTimeout(2000);

        console.log('บันทึกสำเร็จ');
        worksheet.getCell(`J${row}`).value = 'บันทึกสำเร็จ';

        const cellH = worksheet.getCell(`H${row}`).value;
        const cellJ = worksheet.getCell(`J${row}`).value;
        worksheet.getCell(`K${row}`).value = cellH === cellJ ? 'True' : 'False';

        // Save Excel changes
        await workbook.xlsx.writeFile("C:\\Users\\Vivo\\Desktop\\Test_Project\\tests\\04_Data_EditProfile.xlsx");

        row++;
        round++;
    }

    await workbook.xlsx.writeFile("C:\\Users\\Vivo\\Desktop\\Test_Project\\tests\\04_Data_EditProfile.xlsx");
});
