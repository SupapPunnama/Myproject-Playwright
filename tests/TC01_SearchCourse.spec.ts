import { connect } from "http2";

const { test : search_course } = require('@playwright/test');
const search_course_ExcelJS = require('exceljs');

search_course.only("Search Courses", async ({ page }) => {
    const workbook = new search_course_ExcelJS.Workbook();
    search_course.setTimeout(250000);
    await workbook.xlsx.readFile("C:\\Users\\Vivo\\Desktop\\Test_Project\\tests\\01_Data_Search_Course.xlsx");
    const worksheet = workbook.getWorksheet(1);

    // เข้าสู่เว็บไซต์และหน้า Search Course
    await page.goto('http://localhost:8083/sci_mju_lifelonglearning/', { waitUntil: 'domcontentloaded' });
    await page.waitForFunction(() => document.querySelector('title')?.textContent === 'Science MJU LifeLong Learning');
    await page.goto('http://localhost:8083/sci_mju_lifelonglearning/search_course');

    let row = 2; 
    let round = 0; 
    while (row <= 6) { 

        await page.reload();

        const Datasearch_course = worksheet.getCell(`B${row}`).value; // ข้อมูลที่ต้องการเซิร์ช
        if (Datasearch_course) {
            await page.waitForSelector('#searchInput', { visible: true });
            await page.fill('#searchInput', Datasearch_course.toString());
            await page.press('#searchInput', 'Enter')
        }

        await page.waitForTimeout(9000); // รอผลลัพธ์การค้นหา
        const DataPageResult = worksheet.getCell(`C${row}`).value; // ข้อมูลคอร์สที่คาดว่าจะเจอ
        const elementHandle = await page.$('//div[@class="block col-lg-3 col-md-6 wow zoomIn"]');

        // ตรวจสอบว่าพบผลลัพธ์การค้นหาหรือไม่
        if (elementHandle) {
            const DataPagesearch_course = await elementHandle.textContent(); 
            if (DataPagesearch_course.includes(DataPageResult)) {
                console.log(`Row ${row}: Pass`);
                worksheet.getCell(`E${row}`).value = 'Pass';
            } else {
                console.log(`Row ${row}: Fail`);
                worksheet.getCell(`E${row}`).value = 'Fail';
            }
        } else {
            console.log(`Row ${row}: Fail - No search result`);
            worksheet.getCell(`E${row}`).value = 'Fail'; 
        }

        row++; 
        round++; 
        continue;
    }

    await workbook.xlsx.writeFile("C:\\Users\\Vivo\\Desktop\\Test_Project\\tests\\01_Data_Search_Course.xlsx");
});
