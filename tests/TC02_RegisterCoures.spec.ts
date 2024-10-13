const { test: RegisterCourse } = require('@playwright/test');
const ExcelJS = require('exceljs');

RegisterCourse.only("Register Courses Test", async ({ page }) => {
    RegisterCourse.setTimeout(2000000);
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile("C:\\Users\\Vivo\\Desktop\\Test_Project\\tests\\02_Data_Register.xlsx");

    await page.goto('http://localhost:8083/sci_mju_lifelonglearning/', { waitUntil: 'domcontentloaded' });
    await page.waitForFunction(() => document.querySelector('title')?.textContent === 'Science MJU LifeLong Learning');
    await page.goto('http://localhost:8083/sci_mju_lifelonglearning/register_member');

    const worksheet = workbook.getWorksheet(1);

    let row = 2; 
    let round = 0; 

    while (round <= 71) { 
        const processFlag = worksheet.getCell(`Q${row}`).value; 
        console.log(processFlag);

        if (processFlag !== 'Yes') {
            round++;
            row++;
            continue;
        }

        await page.reload();
        const idCardValue = worksheet.getCell(`B${row}`).value; // เลขบัตรประชาชน
        if (idCardValue) {
            await page.waitForSelector('#idcard', { visible: true });
            await page.fill('#idcard', idCardValue.toString());
        }

        const firstName = worksheet.getCell(`C${row}`).value; // ชื่อจริง
        if (firstName) {
            await page.waitForSelector('#firstName', { visible: true });
            await page.fill('#firstName', firstName.toString());
        }

        const lastName = worksheet.getCell(`D${row}`).value; // นามสกุล
        if (lastName) {
            await page.waitForSelector('#lastName', { visible: true });
            await page.fill('#lastName', lastName.toString());
        }

        const genderValue = worksheet.getCell(`E${row}`).value; // เพศ
        if (genderValue) {
            const gender = genderValue.toString();
            if (gender === 'เลือกเพศชาย') {
                const maleRadioButton = await page.locator("//td//input[@value='ชาย']");
                if (await maleRadioButton.count() > 0) {
                    await maleRadioButton.first().check();
                }
            } else if (gender === 'เลือกเพศหญิง') {
                const femaleRadioButton = await page.locator("//td//input[@value='หญิง']");
                if (await femaleRadioButton.count() > 0) {
                    await femaleRadioButton.first().check();
                }
            }
        }

        const email = worksheet.getCell(`F${row}`).value; // อีเมล
        if (email) {
            await page.waitForSelector('#email', { visible: true });
            await page.fill('#email', email.toString());
        }

        const nextBtn = await page.$('#nextBtn');
        await nextBtn.scrollIntoViewIfNeeded();
        await nextBtn.click();

        // รอข้อความแสดงข้อผิดพลาดหรือข้อความสำเร็จ
        await page.waitForTimeout(1000);

        const invalidIdCard = await page.textContent('#invalidIdCard');
        const invalidFirstname = await page.textContent('#invalidFirstname');
        const invalidLastName = await page.textContent('#invalidLastName');
        const invalidGender = await page.textContent('#invalidGender');
        const invalidEmail = await page.textContent('#invalidEmail');

        if (!invalidIdCard.trim() && !invalidFirstname.trim() && !invalidLastName.trim() && !invalidGender.trim() && !invalidEmail.trim()) {
            console.log('ฟิลด์ทั้งหมดถูกต้อง ดำเนินการต่อ');
            
            const birthdayCellValue = worksheet.getCell(`G${row}`).value; // วันเกิด 
            if (birthdayCellValue) {
                try {
                    const dateParts = birthdayCellValue.split('/');
                    const buddhistYear = parseInt(dateParts[2], 10);
                    const christianYear = buddhistYear - 543; // แปลงปีพุทธศักราชเป็นคริสต์ศักราช
                    const formattedBirthday = `${christianYear}-${dateParts[1].padStart(2, '0')}-${dateParts[0].padStart(2, '0')}`;
                    await page.waitForSelector('#datePicker', { visible: true });
                    await page.fill('#datePicker', formattedBirthday);
                } catch (error) {
                    console.error(`Error parsing date value in cell G${row}: ${birthdayCellValue}`, error);
                }
            }

            const tel = worksheet.getCell(`H${row}`).value; // เบอร์โทร
            if (tel) {
                await page.waitForSelector('#tel', { visible: true });
                await page.fill('#tel', tel.toString());
            }

            const education = worksheet.getCell(`I${row}`).value; // สถานศึกษา
            if (education) {
                await page.waitForSelector('#education', { visible: true });

                const optionValues = await page.$$eval('#education option', options =>
                    options.map(option => option.textContent?.trim() || '')
                );

                if (optionValues.includes(education.trim())) {
                    await page.selectOption('#education', { label: education.trim() });
                } else {
                    console.log(`ตัวเลือก '${education}' ไม่มีในรายการ`);
                }
            }

            const username = worksheet.getCell(`J${row}`).value; // ชื่อผู้ใช้
            if (username) {
                await page.waitForSelector('#username', { visible: true });
                await page.fill('#username', username.toString());

                // คลิกปุ่มตรวจสอบ
                await page.waitForSelector('a#link', { visible: true });
                await page.click('a#link');
            }

            const password = worksheet.getCell(`K${row}`).value; // รหัสผ่าน
            if (password) {
                await page.waitForSelector('#password', { visible: true });
                await page.fill('#password', password.toString());
            }

            const confirmPassword = worksheet.getCell(`L${row}`).value; // ยืนยันรหัสผ่าน
            if (confirmPassword) {
                await page.waitForSelector('#confirmPassword');
                await page.fill('#confirmPassword', confirmPassword.toString());
                await page.press('#confirmPassword', 'Enter')
                await page.waitForTimeout(3000);
            }

            const nextBtn = await page.$('#nextBtn');
            await nextBtn.scrollIntoViewIfNeeded();
            await nextBtn.click();
            await page.waitForTimeout(3000);

            // ตรวจสอบฟิลด์เพิ่มเติม
            const invalidBirthday = await page.textContent('#invalidBirthday');
            const invalidTel = await page.textContent('#invalidTel');
            const invalidEducation = await page.textContent('#invalidEducation');
            const status = await page.textContent('#status');
            const invalidPassword = await page.textContent('#invalidPassword');
            const passwordMatchMessage = await page.textContent('#passwordMatchMessage');

            if (passwordMatchMessage.trim() !== "Passwords do not match." && status.trim() === "บัญชีนี้ใช้งานได้" &&  !invalidBirthday.trim()
                && !invalidTel.trim() && !invalidEducation.trim() && !invalidPassword.trim()) {
                console.log('บันทึกสำเร็จ');
                worksheet.getCell(`O${row}`).value = 'บันทึกสำเร็จ';
            } else {
                console.log('ข้อผิดพลาดหน้า 2:', invalidBirthday, invalidTel, invalidEducation, status, invalidPassword, passwordMatchMessage);
                worksheet.getCell(`O${row}`).value = invalidBirthday || invalidTel || invalidEducation || status || invalidPassword || passwordMatchMessage || 'ข้อผิดพลาดไม่รู้จัก';
            }

        } else {
            console.log('ข้อผิดพลาดหน้า 1:', invalidIdCard, invalidFirstname, invalidLastName, invalidGender, invalidEmail);
            worksheet.getCell(`O${row}`).value = invalidIdCard || invalidFirstname || invalidLastName || invalidGender || invalidEmail || 'ข้อผิดพลาดไม่รู้จัก';

        }

        const valueM = worksheet.getCell(`M${row}`).value;
        const valueO = worksheet.getCell(`O${row}`).value;
        worksheet.getCell(`P${row}`).value = valueM === valueO ? 'Pass' : 'Fail';
        await workbook.xlsx.writeFile("C:\\Users\\Vivo\\Desktop\\Test_Project\\tests\\02_Data_Register.xlsx");

        row++; 
        round++; 
        continue;
    }
    await workbook.xlsx.writeFile("C:\\Users\\Vivo\\Desktop\\Test_Project\\tests\\02_Data_Register.xlsx");
});
