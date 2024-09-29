
const { test: Register } = require('@playwright/test');
const ExcelJS2 = require('exceljs');

Register.only("Register Courses", async ({ page }) => {
    Register.setTimeout(120000);
    const workbook = new ExcelJS2.Workbook();
    await workbook.xlsx.readFile("C:\\Users\\Vivo\\Desktop\\Test_Project\\tests\\02_Data_Register.xlsx");

    await page.goto('http://localhost:8083/sci_mju_lifelonglearning/', { waitUntil: 'domcontentloaded' });
    await page.waitForFunction(() => document.querySelector('title')?.textContent === 'Science MJU LifeLong Learning');
    await page.goto('http://localhost:8083/sci_mju_lifelonglearning/register_member');

    const worksheet = workbook.getWorksheet(1);
    let row = 2;
    let round = 0; 

    while (row <= worksheet.rowCount && round < 6) { 
        await page.reload(); 

        const idCardValue = worksheet.getCell(`B${row}`).value; 
        if (idCardValue) {
            await page.fill('#idcard', idCardValue.toString());
        }

        const firstName = worksheet.getCell(`C${row}`).value; // First name
        if (firstName) {
            await page.waitForSelector('#firstName', { visible: true });
            await page.fill('#firstName', firstName.toString());
        }

        const lastName = worksheet.getCell(`D${row}`).value; // Last name
        if (lastName) {
            await page.waitForSelector('#lastName', { visible: true });
            await page.fill('#lastName', lastName.toString());
        }

        const gender = worksheet.getCell(`E${row}`).value; // Gender
        if (gender && gender.toString().toLowerCase() === 'เลือกเพศชาย') {
            const maleRadioButton = await page.locator("//td//input[@value='ชาย']");
            if (await maleRadioButton.count() > 0) {
                await maleRadioButton.first().check();
            }
        }

        const email = worksheet.getCell(`F${row}`).value; // Email
        if (email) {
            await page.waitForSelector('#email', { visible: true });
            await page.fill('#email', email.toString());
        }

        await page.click('#nextBtn');

        // Check ในส่วนของเลขบัตร์ประชาชน
        const invalidIdCard = await page.textContent('#invalidIdCard');
        //const excelCellValue = worksheet.getCell(`M${row}`).value || '';

        if (invalidIdCard.trim() === '') {
            console.log('บันทึกสำเร็จ');
            worksheet.getCell(`O${row}`).value = 'บันทึกสำเร็จ';
        } else {
            console.log(invalidIdCard);

            
            const m2CellValue = worksheet.getCell(`M${row}`).value || ''; // Value in Excel cell M2
            const comparisonResult = invalidIdCard.trim() === m2CellValue.toString().trim();

         
            worksheet.getCell(`O${row}`).value = comparisonResult ? 'มีเออเร่อตรงกับผลลัพธ์ที่คาดหวัง' : 'มีเออเร่อแต่ไม่ตรงกับผลลัพธ์ที่คาดหวัง';
        }

        
        row++;
        round++; 
    }
    // Save Excel file after completing all loops
    await workbook.xlsx.writeFile("C:\\Users\\Vivo\\Desktop\\Test_Project\\tests\\02_Data_Register.xlsx");

    
});




const { test: MemFirstNameTest } = require('@playwright/test'); 
const ExcelJS3 = require('exceljs');

MemFirstNameTest.only("Register Courses - FirstName Test", async ({ page }) => { // ทดสอบในส่วนของชื่อจริง
    const workbook = new ExcelJS3.Workbook();
    await workbook.xlsx.readFile("C:\\Users\\Vivo\\Desktop\\Test_Project\\tests\\02_Data_Register.xlsx");

    await page.goto('http://localhost:8083/sci_mju_lifelonglearning/', { waitUntil: 'domcontentloaded' });
    await page.waitForFunction(() => document.querySelector('title')?.textContent === 'Science MJU LifeLong Learning');
    await page.goto('http://localhost:8083/sci_mju_lifelonglearning/register_member');

    const worksheet = workbook.getWorksheet(1);
    let row = 8;
    let round = 0; 

    while (row <= worksheet.rowCount && round < 9) { 
        await page.reload(); 

        const idCardValue = worksheet.getCell(`B${row}`).value; 
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


        const maleRadioButton = await page.locator("//td//input[@value='ชาย']");
        if (await maleRadioButton.count() > 0) {
            await page.waitForSelector("//td//input[@value='ชาย']", { visible: true });
            await maleRadioButton.first().check();
        }
        
        

        const email = worksheet.getCell(`F${row}`).value; // อีเมล
        if (email) {
            await page.waitForSelector('#email', { visible: true });
            await page.fill('#email', email.toString());
        }

        await page.click('#nextBtn');

        // ตรวจสอบเงื่อนไขในส่วนของชื่อจริง
        const invalidFirstname = await page.textContent('#invalidFirstname');
        const excelCellValue = worksheet.getCell(`M${row}`).value || '';

        if (invalidFirstname.trim() === '') {
            console.log('บันทึกสำเร็จ');
            worksheet.getCell(`O${row}`).value = 'บันทึกสำเร็จ';
        } else {
            console.log(invalidFirstname);

            const m2CellValue = worksheet.getCell(`M${row}`).value || ''; // ค่าใน Excel cell M2
            const comparisonResult = invalidFirstname.trim() === m2CellValue.toString().trim();

            worksheet.getCell(`O${row}`).value = comparisonResult ? 'มีเออเร่อตรงกับผลลัพธ์ที่คาดหวัง' : 'มีเออเร่อแต่ไม่ตรงกับผลลัพธ์ที่คาดหวัง';
        }

        // เซฟไฟล์ Excel
        row++;
        round++; // เพิ่มตัวแปรนับรอบ
    }
    // เซฟไฟล์ Excel หลังจากวนลูปทั้งหมดเสร็จสิ้น
    await workbook.xlsx.writeFile("C:\\Users\\Vivo\\Desktop\\Test_Project\\tests\\02_Data_Register.xlsx");
}, 600000); // ตั้งค่า timeout ของการทดสอบเป็น 120 วินาที (2 นาที)



const { test: MemlastNameTest } = require('@playwright/test'); 
const ExcelJS4 = require('exceljs');

MemlastNameTest.only("Register Courses LastName Test", async ({ page }) => {
    const workbook = new ExcelJS4.Workbook();
    await workbook.xlsx.readFile("C:\\Users\\Vivo\\Desktop\\Test_Project\\tests\\02_Data_Register.xlsx");

    await page.goto('http://localhost:8083/sci_mju_lifelonglearning/', { waitUntil: 'domcontentloaded' });
    await page.waitForFunction(() => document.querySelector('title')?.textContent === 'Science MJU LifeLong Learning');
    await page.goto('http://localhost:8083/sci_mju_lifelonglearning/register_member');

    const worksheet = workbook.getWorksheet(1);
    let row = 17; // เริ่มต้นที่แถวที่ 17
    let round = 0;

    while (round < 8) { // วนลูป 8 รอบ
        await page.reload();

        const idCardValue = worksheet.getCell(`B${row}`).value;
        if (idCardValue) {
            await page.waitForSelector('#idcard', { visible: true });
            await page.fill('#idcard', idCardValue.toString());
        }

        const firstName = worksheet.getCell(`C${row}`).value;
        if (firstName) {
            await page.waitForSelector('#firstName', { visible: true });
            await page.fill('#firstName', firstName.toString());
        }

        const lastName = worksheet.getCell(`D${row}`).value;
        if (lastName) {
            await page.waitForSelector('#lastName', { visible: true });
            await page.fill('#lastName', lastName.toString());
        }

        const maleRadioButton = await page.locator("//td//input[@value='ชาย']");
        if (await maleRadioButton.count() > 0) {
            await page.waitForSelector("//td//input[@value='ชาย']", { visible: true });
            await maleRadioButton.first().check();
        }

        const email = worksheet.getCell(`F${row}`).value;
        if (email) {
            await page.waitForSelector('#email', { visible: true });
            await page.fill('#email', email.toString());
        }

        await page.click('#nextBtn');

        // ตรวจสอบเงื่อนไขในส่วนของนามสกุลจริง
        const invalidFirstname = await page.textContent('#invalidLastName'); // แก้ไขการตรวจสอบเป็น firstName
        const excelCellValue = worksheet.getCell(`M${row}`).value || '';

        if (invalidFirstname.trim() === '') {
            console.log('บันทึกสำเร็จ');
            worksheet.getCell(`O${row}`).value = 'บันทึกสำเร็จ';
        } else {
            console.log(invalidFirstname);

            const m2CellValue = worksheet.getCell(`M${row}`).value || '';
            const comparisonResult = invalidFirstname.trim() === m2CellValue.toString().trim();

            worksheet.getCell(`O${row}`).value = comparisonResult ? 'มีเออเร่อตรงกับผลลัพธ์ที่คาดหวัง' : 'มีเออเร่อแต่ไม่ตรงกับผลลัพธ์ที่คาดหวัง';
        }

        // เซฟไฟล์ Excel
        row++;
        round++; // เพิ่มตัวแปรนับรอบ
    }

    // เซฟไฟล์ Excel หลังจากวนลูปทั้งหมดเสร็จสิ้น
    await workbook.xlsx.writeFile("C:\\Users\\Vivo\\Desktop\\Test_Project\\tests\\02_Data_Register.xlsx");
}, 600000);


const { test: MemgenderTest } = require('@playwright/test'); 
const ExcelJS5 = require('exceljs');

MemgenderTest.only("Register Courses Gender Test", async ({ page }) => {
    const workbook = new ExcelJS5.Workbook();
    await workbook.xlsx.readFile("C:\\Users\\Vivo\\Desktop\\Test_Project\\tests\\02_Data_Register.xlsx");

    await page.goto('http://localhost:8083/sci_mju_lifelonglearning/', { waitUntil: 'domcontentloaded' });
    await page.waitForFunction(() => document.querySelector('title')?.textContent === 'Science MJU LifeLong Learning');
    await page.goto('http://localhost:8083/sci_mju_lifelonglearning/register_member');

    const worksheet = workbook.getWorksheet(1);
    let row = 25; // เริ่มต้นที่แถวที่ 25
    let round = 0;

    while (round < 3) { // วนลูป 3 รอบ
        await page.reload();

        const idCardValue = worksheet.getCell(`B${row}`).value;
        if (idCardValue) {
            await page.waitForSelector('#idcard', { visible: true });
            await page.fill('#idcard', idCardValue.toString());
        }

        const firstName = worksheet.getCell(`C${row}`).value;
        if (firstName) {
            await page.waitForSelector('#firstName', { visible: true });
            await page.fill('#firstName', firstName.toString());
        }

        const lastName = worksheet.getCell(`D${row}`).value;
        if (lastName) {
            await page.waitForSelector('#lastName', { visible: true });
            await page.fill('#lastName', lastName.toString());
        }

        const genderValue = worksheet.getCell(`E${row}`).value; // ดึงค่าเพศจากเซลล์ E
        if (genderValue) {
            if (genderValue.toString() === 'เลือกเพศชาย') {
                const maleRadioButton = await page.locator("//td//input[@value='ชาย']");
                if (await maleRadioButton.count() > 0) {
                    await page.waitForSelector("//td//input[@value='ชาย']", { visible: true });
                    await maleRadioButton.first().check();
                }
            } else if (genderValue.toString() === 'เลือกเพศหญิง') {
                const femaleRadioButton = await page.locator("//td//input[@value='หญิง']");
                if (await femaleRadioButton.count() > 0) {
                    await page.waitForSelector("//td//input[@value='หญิง']", { visible: true });
                    await femaleRadioButton.first().check();
                }
            }
        }

        const email = worksheet.getCell(`F${row}`).value;
        if (email) {
            await page.waitForSelector('#email', { visible: true });
            await page.fill('#email', email.toString());
        }

        await page.click('#nextBtn');

        //ตรวจสอบเงื่อนไขในส่วนของเพศ
        const invalidFirstname = await page.textContent('#invalidGender'); // แก้ไขการตรวจสอบเป็น gender
        const excelCellValue = worksheet.getCell(`M${row}`).value || '';

        if (invalidFirstname.trim() === '') {
            console.log('บันทึกสำเร็จ');
            worksheet.getCell(`O${row}`).value = 'บันทึกสำเร็จ';
        } else {
            console.log(invalidFirstname);

            const m2CellValue = worksheet.getCell(`M${row}`).value || '';
            const comparisonResult = invalidFirstname.trim() === m2CellValue.toString().trim();

            worksheet.getCell(`O${row}`).value = comparisonResult ? 'มีเออเร่อตรงกับผลลัพธ์ที่คาดหวัง' : 'มีเออเร่อแต่ไม่ตรงกับผลลัพธ์ที่คาดหวัง';
        }

        // เซฟไฟล์ Excel
        row++;
        round++; // เพิ่มตัวแปรนับรอบ
    }

   // เซฟไฟล์ Excel หลังจากวนลูปทั้งหมดเสร็จสิ้น
    await workbook.xlsx.writeFile("C:\\Users\\Vivo\\Desktop\\Test_Project\\tests\\02_Data_Register.xlsx");
}, 600000);



const { test: EmailTest } = require('@playwright/test'); 
const ExcelJS6 = require('exceljs');

EmailTest.only("Register Courses Email Test", async ({ page }) => {
    const workbook = new ExcelJS6.Workbook();
    await workbook.xlsx.readFile("C:\\Users\\Vivo\\Desktop\\Test_Project\\tests\\02_Data_Register.xlsx");

    await page.goto('http://localhost:8083/sci_mju_lifelonglearning/', { waitUntil: 'domcontentloaded' });
    await page.waitForFunction(() => document.querySelector('title')?.textContent === 'Science MJU LifeLong Learning');
    await page.goto('http://localhost:8083/sci_mju_lifelonglearning/register_member');

    const worksheet = workbook.getWorksheet(1);
    let row = 28; // เริ่มต้นที่แถวที่ 28
    let round = 0;

    while (round < 9) { // วนลูป 9 รอบ
        await page.reload();

        const idCardValue = worksheet.getCell(`B${row}`).value;
        if (idCardValue) {
            await page.waitForSelector('#idcard', { visible: true });
            await page.fill('#idcard', idCardValue.toString());
        }

        const firstName = worksheet.getCell(`C${row}`).value;
        if (firstName) {
            await page.waitForSelector('#firstName', { visible: true });
            await page.fill('#firstName', firstName.toString());
        }

        const lastName = worksheet.getCell(`D${row}`).value;
        if (lastName) {
            await page.waitForSelector('#lastName', { visible: true });
            await page.fill('#lastName', lastName.toString());
        }

        const genderValue = worksheet.getCell(`E${row}`).value; // ดึงค่าเพศจากเซลล์ E
        if (genderValue) {
            if (genderValue.toString() === 'เลือกเพศชาย') {
                const maleRadioButton = await page.locator("//td//input[@value='ชาย']");
                if (await maleRadioButton.count() > 0) {
                    await page.waitForSelector("//td//input[@value='ชาย']", { visible: true });
                    await maleRadioButton.first().check();
                }
            } else if (genderValue.toString() === 'เลือกเพศหญิง') {
                const femaleRadioButton = await page.locator("//td//input[@value='หญิง']");
                if (await femaleRadioButton.count() > 0) {
                    await page.waitForSelector("//td//input[@value='หญิง']", { visible: true });
                    await femaleRadioButton.first().check();
                }
            }
        }

        const email = worksheet.getCell(`F${row}`).value;
        if (email) {
            await page.waitForSelector('#email', { visible: true });
            await page.fill('#email', email.toString());
        }

        await page.click('#nextBtn');

        //ตรวจสอบเงื่อนไขในส่วนของEmail
        const invalidFirstname = await page.textContent('#invalidEmail'); // แก้ไขการตรวจสอบเป็น Email
        const excelCellValue = worksheet.getCell(`M${row}`).value || '';

        if (invalidFirstname.trim() === '') {
            console.log('บันทึกสำเร็จ');
            worksheet.getCell(`O${row}`).value = 'บันทึกสำเร็จ';
        } else {
            console.log(invalidFirstname);

            const m2CellValue = worksheet.getCell(`M${row}`).value || '';
            const comparisonResult = invalidFirstname.trim() === m2CellValue.toString().trim();

            worksheet.getCell(`O${row}`).value = comparisonResult ? 'มีเออเร่อตรงกับผลลัพธ์ที่คาดหวัง' : 'มีเออเร่อแต่ไม่ตรงกับผลลัพธ์ที่คาดหวัง';
        }

        // เซฟไฟล์ Excel
        row++;
        round++; // เพิ่มตัวแปรนับรอบ
    }

   // เซฟไฟล์ Excel หลังจากวนลูปทั้งหมดเสร็จสิ้น
    await workbook.xlsx.writeFile("C:\\Users\\Vivo\\Desktop\\Test_Project\\tests\\02_Data_Register.xlsx");
}, 600000);




const { test: BirthdayTest } = require('@playwright/test');
const ExcelJS7 = require('exceljs');

BirthdayTest.only("Register Courses Birthday Test", async ({ page }) => {
    const workbook = new ExcelJS7.Workbook();
    await workbook.xlsx.readFile("C:\\Users\\Vivo\\Desktop\\Test_Project\\tests\\02_Data_Register.xlsx");

    await page.goto('http://localhost:8083/sci_mju_lifelonglearning/', { waitUntil: 'domcontentloaded' });
    await page.waitForFunction(() => document.querySelector('title')?.textContent === 'Science MJU LifeLong Learning');
    await page.goto('http://localhost:8083/sci_mju_lifelonglearning/register_member');

    const worksheet = workbook.getWorksheet(1);
    let row = 37; // เริ่มต้นที่แถวที่ 37
    let round = 0;

    while (round < 4) { // วนลูป 4 รอบ
        await page.reload(); // โหลดหน้าเว็บไซต์ใหม่
        const idCardValue = worksheet.getCell(`B${row}`).value;
        if (idCardValue) {
            await page.waitForSelector('#idcard', { visible: true, timeout: 60000 });
            await page.fill('#idcard', idCardValue.toString(), { timeout: 60000 });
        }

        const firstName = worksheet.getCell(`C${row}`).value;
        if (firstName) {
            await page.waitForSelector('#firstName', { visible: true, timeout: 60000 });
            await page.fill('#firstName', firstName.toString(), { timeout: 60000 });
        }

        const lastName = worksheet.getCell(`D${row}`).value;
        if (lastName) {
            await page.waitForSelector('#lastName', { visible: true, timeout: 60000 });
            await page.fill('#lastName', lastName.toString(), { timeout: 60000 });
        }

        const genderValue = worksheet.getCell(`E${row}`).value; // ดึงค่าเพศจากเซลล์ E
        if (genderValue) {
            if (genderValue.toString() === 'เลือกเพศชาย') {
                const maleRadioButton = await page.locator("//td//input[@value='ชาย']");
                if (await maleRadioButton.count() > 0) {
                    await page.waitForSelector("//td//input[@value='ชาย']", { visible: true, timeout: 60000 });
                    await maleRadioButton.first().check();
                }
            } else if (genderValue.toString() === 'เลือกเพศหญิง') {
                const femaleRadioButton = await page.locator("//td//input[@value='หญิง']");
                if (await femaleRadioButton.count() > 0) {
                    await page.waitForSelector("//td//input[@value='หญิง']", { visible: true, timeout: 60000 });
                    await femaleRadioButton.first().check();
                }
            }
        }

        const email = worksheet.getCell(`F${row}`).value; // Email
        if (email) {
            await page.waitForSelector('#email', { visible: true, timeout: 60000 });
            await page.fill('#email', email.toString(), { timeout: 60000 });
        }

        await page.click('#nextBtn');


        const birthdayCellValue = worksheet.getCell(`G${row}`).value; // วันเกิด
        if (birthdayCellValue) {
            try {
                const dateParts = birthdayCellValue.split('/');
                const formattedBirthday = `${dateParts[2]}-${dateParts[0].padStart(2, '0')}-${dateParts[1].padStart(2, '0')}`;
                await page.waitForSelector('#datePicker', { visible: true, timeout: 60000 });
                await page.fill('#datePicker', formattedBirthday, { timeout: 60000 });
            } catch (error) {
                console.error(`Error parsing date value in cell G${row}: ${birthdayCellValue}`, error);
            }
        }

     
        const tel = worksheet.getCell(`H${row}`).value; // เบอร์โทร
        if (tel) {
            await page.waitForSelector('#tel', { visible: true, timeout: 60000 });
            await page.fill('#tel', tel.toString(), { timeout: 60000 });
        }

        const education = worksheet.getCell(`I${row}`).value; // สถานศึกษา
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

        const username = worksheet.getCell(`J${row}`).value; // ชื่อผู้ใช้
        if (username) {
            await page.waitForSelector('#username', { visible: true, timeout: 60000 });
            await page.fill('#username', username.toString(), { timeout: 60000 });
        }

        const password = worksheet.getCell(`K${row}`).value; // รหัสผ่าน
        if (password) {
            await page.waitForSelector('#password', { visible: true, timeout: 60000 });
            await page.fill('#password', password.toString(), { timeout: 60000 });
        }

        const confirmPassword = worksheet.getCell(`L${row}`).value; // ยืนยันรหัสผ่าน
        if (confirmPassword) {
            await page.waitForSelector('#confirmPassword', { visible: true, timeout: 60000 });
            await page.fill('#confirmPassword', confirmPassword.toString(), { timeout: 60000 });
        }

        await page.click('#nextBtn');



        // รอจนกว่าข้อความในหน้าเว็บที่ตำแหน่ง #displayDatePicker จะเปลี่ยนแปลง
        const actualValue = await page.$eval('#displayDatePicker', el => el.textContent.trim());
        const expectedValue = worksheet.getCell(`M${row}`).value; // ค่าเปรียบเทียบที่คาดหวัง

        if (actualValue === "NaN/NaN/NaN" || actualValue === null) {
        if (actualValue === expectedValue) {
        console.log("ตรงตามผลลัพธ์ที่คาดหวัง");
            worksheet.getCell(`O${row}`).value = "ตรงตามผลลัพธ์ที่คาดหวัง";
        } else {
        console.log("ไม่ตรงตามผลลัพธ์ที่คาดหวัง");
            worksheet.getCell(`O${row}`).value = "ไม่ตรงตามผลลัพธ์ที่คาดหวัง";
        }
        } else {
        console.log("บันทึกสำเร็จ");
            worksheet.getCell(`O${row}`).value = "บันทึกสำเร็จ";
        }
 
        row++;
        round++; // เพิ่มตัวแปรนับรอบ
    }

    // เซฟไฟล์ Excel หลังจากวนลูปทั้งหมดเสร็จสิ้น
    await workbook.xlsx.writeFile("C:\\Users\\Vivo\\Desktop\\Test_Project\\tests\\02_Data_Register.xlsx");
}, 16000000);



const { test: PhonTest } = require('@playwright/test');
const ExcelJS8 = require('exceljs');

PhonTest.only("Register Courses Phon Test", async ({ page }) => {
    const workbook = new ExcelJS8.Workbook();
    await workbook.xlsx.readFile("C:\\Users\\Vivo\\Desktop\\Test_Project\\tests\\02_Data_Register.xlsx");

    await page.goto('http://localhost:8083/sci_mju_lifelonglearning/', { waitUntil: 'domcontentloaded' });
    await page.waitForFunction(() => document.querySelector('title')?.textContent === 'Science MJU LifeLong Learning');
    await page.goto('http://localhost:8083/sci_mju_lifelonglearning/register_member');

    const worksheet = workbook.getWorksheet(1);
    let row = 41; // เริ่มต้นที่แถวที่ 41
    let round = 0;

    while (round < 8) { // วนลูป 8 รอบ
        await page.reload(); // โหลดหน้าเว็บไซต์ใหม่

        const idCardValue = worksheet.getCell(`B${row}`).value;
        if (idCardValue) {
            await page.waitForSelector('#idcard', { visible: true});
            await page.fill('#idcard', idCardValue.toString());
        }

        const firstName = worksheet.getCell(`C${row}`).value;
        if (firstName) {
            await page.waitForSelector('#firstName', { visible: true });
            await page.fill('#firstName', firstName.toString());
        }

        const lastName = worksheet.getCell(`D${row}`).value;
        if (lastName) {
            await page.waitForSelector('#lastName', { visible: true });
            await page.fill('#lastName', lastName.toString());
        }

        const genderValue = worksheet.getCell(`E${row}`).value; // ดึงค่าเพศจากเซลล์ E
        if (genderValue) {
            if (genderValue.toString() === 'เลือกเพศชาย') {
                const maleRadioButton = await page.locator("//td//input[@value='ชาย']");
                if (await maleRadioButton.count() > 0) {
                    await page.waitForSelector("//td//input[@value='ชาย']", { visible: true });
                    await maleRadioButton.first().check();
                }
            } else if (genderValue.toString() === 'เลือกเพศหญิง') {
                const femaleRadioButton = await page.locator("//td//input[@value='หญิง']");
                if (await femaleRadioButton.count() > 0) {
                    await page.waitForSelector("//td//input[@value='หญิง']", { visible: true });
                    await femaleRadioButton.first().check();
                }
            }
        }

        const email = worksheet.getCell(`F${row}`).value; // Email
        if (email) {
            await page.waitForSelector('#email', { visible: true });
            await page.fill('#email', email.toString());
        }

        const nextBtn = await page.$('#nextBtn');
        await nextBtn.scrollIntoViewIfNeeded();
        await nextBtn.click();

        const birthdayCellValue = worksheet.getCell(`G${row}`).value; // วันเกิด
        if (birthdayCellValue) {
            try {
                const dateParts = birthdayCellValue.split('/');
                const formattedBirthday = `${dateParts[2]}-${dateParts[0].padStart(2, '0')}-${dateParts[1].padStart(2, '0')}`;
                await page.waitForSelector('#datePicker', { visible: true});
                await page.fill('#datePicker', formattedBirthday);
            } catch (error) {
                console.error(`Error parsing date value in cell G${row}: ${birthdayCellValue}`, error);
            }
        }

        const tel = worksheet.getCell(`H${row}`).value; // เบอร์โทร
        if (tel) {
            await page.waitForSelector('#tel', { visible: true, timeout: 3000 });
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
        }

        const password = worksheet.getCell(`K${row}`).value; // รหัสผ่าน
        if (password) {
            await page.waitForSelector('#password', { visible: true });
            await page.fill('#password', password.toString());
        }

        const confirmPassword = worksheet.getCell(`L${row}`).value; // ยืนยันรหัสผ่าน
        if (confirmPassword) {
            await page.waitForSelector('#confirmPassword', { visible: true });
            await page.fill('#confirmPassword', confirmPassword.toString());
        }

        await nextBtn.scrollIntoViewIfNeeded();
        await nextBtn.click();


        let actualValue = "";
        try {
            actualValue = await page.$eval('#invalidTel', el => el.textContent.trim());
        } catch (error) {
            console.error(`Error retrieving actual value for row ${row}:`, error);
        }
        
        const expectedValue = worksheet.getCell(`M${row}`).value;
        const excelValue = worksheet.getCell(`H${row}`).value;
        
        let resultMessage = "";
        if (!excelValue) {
            resultMessage = "บันทึกไม่สำเร็จ";  // Print message if Excel value is empty
        } else if (actualValue !== "") {
            if (actualValue === expectedValue) {
                resultMessage = "ตรงตามผลลัพธ์ที่คาดหวัง";
            } else {
                resultMessage = `ไม่ตรงตามผลลัพธ์ที่คาดหวัง: ${actualValue}`;
            }
        } else {
            resultMessage = "บันทึกสำเร็จ";
        }
        
        console.log(resultMessage);
        worksheet.getCell(`O${row}`).value = resultMessage;
        
        row++;
        round++;
        
    }
    await workbook.xlsx.writeFile("C:\\Users\\Vivo\\Desktop\\Test_Project\\tests\\02_Data_Register.xlsx");  
}, 60000);



const { test: educationTest } = require('@playwright/test');
const ExcelJS9 = require('exceljs');

educationTest.only("Register Courses education Test", async ({ page }) => {
    const workbook = new ExcelJS9.Workbook();
    await workbook.xlsx.readFile("C:\\Users\\Vivo\\Desktop\\Test_Project\\tests\\02_Data_Register.xlsx");

    await page.goto('http://localhost:8083/sci_mju_lifelonglearning/', { waitUntil: 'domcontentloaded' });
    await page.waitForFunction(() => document.querySelector('title')?.textContent === 'Science MJU LifeLong Learning');
    await page.goto('http://localhost:8083/sci_mju_lifelonglearning/register_member');

    const worksheet = workbook.getWorksheet(1);
    let row = 49; // เริ่มต้นที่แถวที่ 49
    let round = 0;

    while (round < 1) { // วนลูป 1 รอบ
        await page.reload(); // โหลดหน้าเว็บไซต์ใหม่

        const idCardValue = worksheet.getCell(`B${row}`).value;
        if (idCardValue) {
            await page.waitForSelector('#idcard', { visible: true});
            await page.fill('#idcard', idCardValue.toString());
        }

        const firstName = worksheet.getCell(`C${row}`).value;
        if (firstName) {
            await page.waitForSelector('#firstName', { visible: true });
            await page.fill('#firstName', firstName.toString());
        }

        const lastName = worksheet.getCell(`D${row}`).value;
        if (lastName) {
            await page.waitForSelector('#lastName', { visible: true });
            await page.fill('#lastName', lastName.toString());
        }

        const genderValue = worksheet.getCell(`E${row}`).value; // ดึงค่าเพศจากเซลล์ E
        if (genderValue) {
            if (genderValue.toString() === 'เลือกเพศชาย') {
                const maleRadioButton = await page.locator("//td//input[@value='ชาย']");
                if (await maleRadioButton.count() > 0) {
                    await page.waitForSelector("//td//input[@value='ชาย']", { visible: true });
                    await maleRadioButton.first().check();
                }
            } else if (genderValue.toString() === 'เลือกเพศหญิง') {
                const femaleRadioButton = await page.locator("//td//input[@value='หญิง']");
                if (await femaleRadioButton.count() > 0) {
                    await page.waitForSelector("//td//input[@value='หญิง']", { visible: true });
                    await femaleRadioButton.first().check();
                }
            }
        }

        const email = worksheet.getCell(`F${row}`).value; // Email
        if (email) {
            await page.waitForSelector('#email', { visible: true });
            await page.fill('#email', email.toString());
        }

        const nextBtn = await page.$('#nextBtn');
        await nextBtn.scrollIntoViewIfNeeded();
        await nextBtn.click();

        const birthdayCellValue = worksheet.getCell(`G${row}`).value; // วันเกิด
        if (birthdayCellValue) {
            try {
                const dateParts = birthdayCellValue.split('/');
                const formattedBirthday = `${dateParts[2]}-${dateParts[0].padStart(2, '0')}-${dateParts[1].padStart(2, '0')}`;
                await page.waitForSelector('#datePicker', { visible: true});
                await page.fill('#datePicker', formattedBirthday);
            } catch (error) {
                console.error(`Error parsing date value in cell G${row}: ${birthdayCellValue}`, error);
            }
        }

        const tel = worksheet.getCell(`H${row}`).value; // เบอร์โทร
        if (tel) {
            await page.waitForSelector('#tel', { visible: true, timeout: 3000 });
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
        }

        const password = worksheet.getCell(`K${row}`).value; // รหัสผ่าน
        if (password) {
            await page.waitForSelector('#password', { visible: true });
            await page.fill('#password', password.toString());
        }

        const confirmPassword = worksheet.getCell(`L${row}`).value; // ยืนยันรหัสผ่าน
        if (confirmPassword) {
            await page.waitForSelector('#confirmPassword', { visible: true });
            await page.fill('#confirmPassword', confirmPassword.toString());
        }

        await nextBtn.scrollIntoViewIfNeeded();
        await nextBtn.click();


        let actualValue = "";
        try {
            actualValue = await page.$eval('#invalidEducation', el => el.textContent.trim());
        } catch (error) {
            console.error(`Error retrieving actual value for row ${row}:`, error);
        }
        
        const expectedValue = worksheet.getCell(`M${row}`).value;
        let resultMessage = "";
        if (!actualValue) {
            resultMessage = "บันทึกสำเร็จ";  // Print message if actualValue is empty
        } else if (actualValue === expectedValue) {
            resultMessage = "ตรงตามผลลัพธ์ที่คาดหวัง";
        } else {
            resultMessage = `มีเอ่อเร่อแต่ไม่ตรงกับที่คาดหวัง: ${actualValue}`;
        }
        console.log(resultMessage);
        worksheet.getCell(`O${row}`).value = resultMessage;
        
        row++;
        round++;
        
    }
    await workbook.xlsx.writeFile("C:\\Users\\Vivo\\Desktop\\Test_Project\\tests\\02_Data_Register.xlsx");  
}, 60000);



const { test: usernameTest } = require('@playwright/test');
const ExcelJS10 = require('exceljs');

usernameTest.only("Register Courses username Test", async ({ page }) => {
    usernameTest.setTimeout(60000);
    const workbook = new ExcelJS10.Workbook();
    await workbook.xlsx.readFile("C:\\Users\\Vivo\\Desktop\\Test_Project\\tests\\02_Data_Register.xlsx");

    await page.goto('http://localhost:8083/sci_mju_lifelonglearning/', { waitUntil: 'domcontentloaded' });
    await page.waitForFunction(() => document.querySelector('title')?.textContent === 'Science MJU LifeLong Learning');
    await page.goto('http://localhost:8083/sci_mju_lifelonglearning/register_member');

    const worksheet = workbook.getWorksheet(1);
    let row = 50; // เริ่มต้นที่แถวที่ 50
    let round = 0;

    while (round < 10) { // วนลูป 10 รอบ
        await page.reload(); // โหลดหน้าเว็บไซต์ใหม่

        // รหัสบัตรประชาชน
        await page.waitForSelector('#idcard', { visible: true , timeout: 60000});
        await page.fill('#idcard', '1509970158094');    

        // ชื่อจริง
        await page.waitForSelector('#firstName', { visible: true , timeout: 60000});
        await page.fill('#firstName', 'สุภาพ');

        // นามสกุลจริง
        await page.waitForSelector('#lastName', { visible: true, timeout: 60000 }); // เพิ่ม timeout
        await page.fill('#lastName', 'ปุนนะมา');

        // เพศ
        await page.waitForSelector("//td//input[@value='ชาย']", { visible: true , timeout: 60000});
        await page.click("//td//input[@value='ชาย']");

        // Email
        await page.waitForSelector('#email', { visible: true , timeout: 60000});
        await page.fill('#email', 'supap_1998@gmail.com');

        // ปุ่มต่อไป
        await page.waitForSelector('#nextBtn', { visible: true , timeout: 60000});
        await page.click('#nextBtn');

        // วันเกิด
        await page.waitForSelector('#datePicker', { visible: true, timeout: 60000 }); // เพิ่ม timeout
        await page.fill('#datePicker', '1998-07-03');

        // เบอร์โทร
        await page.waitForSelector('#tel', { visible: true , timeout: 60000});
        await page.fill('#tel', '0959256362');

        // สถานศึกษา
        await page.waitForSelector('#education', { visible: true , timeout: 60000});
        await page.selectOption('#education', 'ระดับอาชีวศึกษา');

        // ชื่อผู้ใช้
        const username = worksheet.getCell(`J${row}`).value; 
        if (username) {
            await page.waitForSelector('#username', { visible: true , timeout: 60000});
            await page.fill('#username', username.toString());
        }

        // รหัสผ่าน
        await page.waitForSelector('#password', { visible: true , timeout: 60000});
        await page.fill('#password', 'Sunz_.');
        
        // ยืนยันรหัสผ่าน
        await page.waitForSelector('#confirmPassword', { visible: true , timeout: 60000});
        await page.fill('#confirmPassword', 'Sunz_.');

        // ตรวจสอบ
        await page.waitForSelector('#link', { visible: true, timeout: 60000 }); // เพิ่ม timeout
        await page.click('#link');

        // ตรวจสอบสถานะ
        await page.waitForSelector('#status', { visible: true, timeout: 60000 }); // เพิ่ม timeout
        const actualValue = await page.$eval('#status', el => el.textContent.trim());

        const expectedValue = worksheet.getCell(`M${row}`).value;
        let resultMessage = "";
        if (actualValue === expectedValue) {
            resultMessage = "ตรงตามผลลัพธ์ที่คาดหวัง";
        } else {
            resultMessage = `ไม่ตรงตามผลลัพธ์ที่คาดหวัง: ${actualValue}`;
        }
        console.log(resultMessage);
        worksheet.getCell(`O${row}`).value = resultMessage;

        row++;
        round++;
    }

    // Save the workbook after modifications
    await workbook.xlsx.writeFile("C:\\Users\\Vivo\\Desktop\\Test_Project\\tests\\02_Data_Register.xlsx");
}, 120000); // เพิ่ม timeout สำหรับการทดสอบ



const { test: password } = require('@playwright/test');
const ExcelJS11 = require('exceljs');

password.only("Register Courses password Test", async ({ page }) => {
    password.setTimeout(90000);
    const workbook = new ExcelJS11.Workbook();
    await workbook.xlsx.readFile("C:\\Users\\Vivo\\Desktop\\Test_Project\\tests\\02_Data_Register.xlsx");

    await page.goto('http://localhost:8083/sci_mju_lifelonglearning/', { waitUntil: 'domcontentloaded' });
    await page.waitForFunction(() => document.querySelector('title')?.textContent === 'Science MJU LifeLong Learning');
    await page.goto('http://localhost:8083/sci_mju_lifelonglearning/register_member');

    const worksheet = workbook.getWorksheet(1);
    let row = 60; // เริ่มต้นที่แถวที่ 60
    let round = 0;

    while (round < 10) { // วนลูป 10 รอบ
        await page.reload(); // โหลดหน้าเว็บไซต์ใหม่

        const idCardValue = worksheet.getCell(`B${row}`).value;
        if (idCardValue) {
            await page.waitForSelector('#idcard', { visible: true});
            await page.fill('#idcard', idCardValue.toString());
        }

        const firstName = worksheet.getCell(`C${row}`).value;
        if (firstName) {
            await page.waitForSelector('#firstName', { visible: true });
            await page.fill('#firstName', firstName.toString());
        }

        const lastName = worksheet.getCell(`D${row}`).value;
        if (lastName) {
            await page.waitForSelector('#lastName', { visible: true });
            await page.fill('#lastName', lastName.toString());
        }

        const genderValue = worksheet.getCell(`E${row}`).value; // ดึงค่าเพศจากเซลล์ E
        if (genderValue) {
            if (genderValue.toString() === 'เลือกเพศชาย') {
                const maleRadioButton = await page.locator("//td//input[@value='ชาย']");
                if (await maleRadioButton.count() > 0) {
                    await page.waitForSelector("//td//input[@value='ชาย']", { visible: true });
                    await maleRadioButton.first().check();
                }
            } else if (genderValue.toString() === 'เลือกเพศหญิง') {
                const femaleRadioButton = await page.locator("//td//input[@value='หญิง']");
                if (await femaleRadioButton.count() > 0) {
                    await page.waitForSelector("//td//input[@value='หญิง']", { visible: true });
                    await femaleRadioButton.first().check();
                }
            }
        }

        const email = worksheet.getCell(`F${row}`).value; // Email
        if (email) {
            await page.waitForSelector('#email', { visible: true });
            await page.fill('#email', email.toString());
        }

        const nextBtn = await page.$('#nextBtn');
        await nextBtn.scrollIntoViewIfNeeded();
        await nextBtn.click();

        const birthdayCellValue = worksheet.getCell(`G${row}`).value; // วันเกิด
        if (birthdayCellValue) {
            try {
                const dateParts = birthdayCellValue.split('/');
                const formattedBirthday = `${dateParts[2]}-${dateParts[0].padStart(2, '0')}-${dateParts[1].padStart(2, '0')}`;
                await page.waitForSelector('#datePicker', { visible: true});
                await page.fill('#datePicker', formattedBirthday);
            } catch (error) {
                console.error(`Error parsing date value in cell G${row}: ${birthdayCellValue}`, error);
            }
        }

        const tel = worksheet.getCell(`H${row}`).value; // เบอร์โทร
        if (tel) {
            await page.waitForSelector('#tel', { visible: true, timeout: 3000 });
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
        }

        const password = worksheet.getCell(`K${row}`).value; // รหัสผ่าน
        if (password) {
            await page.waitForSelector('#password', { visible: true });
            await page.fill('#password', password.toString());
        }

        const confirmPassword = worksheet.getCell(`L${row}`).value; // ยืนยันรหัสผ่าน
        if (confirmPassword) {
            await page.waitForSelector('#confirmPassword', { visible: true });
            await page.fill('#confirmPassword', confirmPassword.toString());
        }

        await nextBtn.scrollIntoViewIfNeeded();
        await nextBtn.click();

        const actualValue = await page.$eval('#invalidPassword', el => el.textContent.trim());
        const expectedValue = worksheet.getCell(`M${row}`).value;

        //เงื่อนไข
        if (actualValue === '' && expectedValue !== '') {
            console.log(`Row ${row}: ไม่ตรงกับผลลัพธ์ที่คาดหวัง`);
        } else if (actualValue !== '' && actualValue === expectedValue) {
            console.log(`Row ${row}: ตรงกับผลลัพธ์ที่คาดหวัง`);
        } else {
            console.log(`Row ${row}: ไม่ตรงกับผลลัพธ์ที่คาดหวัง`);
        }
        row++;
        round++;      
    }
    await workbook.xlsx.writeFile("C:\\Users\\Vivo\\Desktop\\Test_Project\\tests\\02_Data_Register.xlsx");  
}, 60000);


const { test: confirmPassword } = require('@playwright/test');
const ExcelJS12 = require('exceljs');

confirmPassword.only("Register Courses confirmPassword Test", async ({ page }) => {
    confirmPassword.setTimeout(120000);
    const workbook = new ExcelJS12.Workbook();
    await workbook.xlsx.readFile("C:\\Users\\Vivo\\Desktop\\Test_Project\\tests\\02_Data_Register.xlsx");

    await page.goto('http://localhost:8083/sci_mju_lifelonglearning/', { waitUntil: 'domcontentloaded' });
    await page.waitForFunction(() => document.querySelector('title')?.textContent === 'Science MJU LifeLong Learning');
    await page.goto('http://localhost:8083/sci_mju_lifelonglearning/register_member');

    const worksheet = workbook.getWorksheet(1);
    let row = 70; // เริ่มต้นที่แถวที่ 70
    let round = 0;

    while (round < 3) { // วนลูป 3 รอบ
        await page.reload(); // โหลดหน้าเว็บไซต์ใหม่

        const idCardValue = worksheet.getCell(`B${row}`).value;
        if (idCardValue) {
            await page.waitForSelector('#idcard', { visible: true});
            await page.fill('#idcard', idCardValue.toString());
        }

        const firstName = worksheet.getCell(`C${row}`).value;
        if (firstName) {
            await page.waitForSelector('#firstName', { visible: true });
            await page.fill('#firstName', firstName.toString());
        }

        const lastName = worksheet.getCell(`D${row}`).value;
        if (lastName) {
            await page.waitForSelector('#lastName', { visible: true });
            await page.fill('#lastName', lastName.toString());
        }

        const genderValue = worksheet.getCell(`E${row}`).value; // ดึงค่าเพศจากเซลล์ E
        if (genderValue) {
            if (genderValue.toString() === 'เลือกเพศชาย') {
                const maleRadioButton = await page.locator("//td//input[@value='ชาย']");
                if (await maleRadioButton.count() > 0) {
                    await page.waitForSelector("//td//input[@value='ชาย']", { visible: true });
                    await maleRadioButton.first().check();
                }
            } else if (genderValue.toString() === 'เลือกเพศหญิง') {
                const femaleRadioButton = await page.locator("//td//input[@value='หญิง']");
                if (await femaleRadioButton.count() > 0) {
                    await page.waitForSelector("//td//input[@value='หญิง']", { visible: true });
                    await femaleRadioButton.first().check();
                }
            }
        }

        const email = worksheet.getCell(`F${row}`).value; // Email
        if (email) {
            await page.waitForSelector('#email', { visible: true });
            await page.fill('#email', email.toString());
        }

        const nextBtn = await page.$('#nextBtn');
        await nextBtn.scrollIntoViewIfNeeded();
        await nextBtn.click();

        const birthdayCellValue = worksheet.getCell(`G${row}`).value; // วันเกิด
        if (birthdayCellValue) {
            try {
                const dateParts = birthdayCellValue.split('/');
                const formattedBirthday = `${dateParts[2]}-${dateParts[0].padStart(2, '0')}-${dateParts[1].padStart(2, '0')}`;
                await page.waitForSelector('#datePicker', { visible: true});
                await page.fill('#datePicker', formattedBirthday);
            } catch (error) {
                console.error(`Error parsing date value in cell G${row}: ${birthdayCellValue}`, error);
            }
        }

        const tel = worksheet.getCell(`H${row}`).value; // เบอร์โทร
        if (tel) {
            await page.waitForSelector('#tel', { visible: true, timeout: 3000 });
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
            await page.waitForTimeout(3000);
        }

            //await nextBtn.scrollIntoViewIfNeeded();
            //await nextBtn.click();

            await page.waitForSelector('#passwordMatchMessage',{ visible: true });
            const actualValue = await page.$eval('#passwordMatchMessage', el => el.textContent.trim());
            const expectedValue = worksheet.getCell(`M${row}`).value;

            // เงื่อนไข
            if (actualValue === expectedValue) {
            console.log(`Row ${row}: ตรงกับผลลัพธ์ที่คาดหวัง`);
            } else {
            console.log(`Row ${row}: ไม่ตรงกับผลลัพธ์ที่คาดหวัง`);
        }
        row++;
        round++;      
    }
    //await workbook.xlsx.writeFile("C:\\Users\\Vivo\\Desktop\\Test_Project\\tests\\02_Data_Register.xlsx");  
}, 60000);
