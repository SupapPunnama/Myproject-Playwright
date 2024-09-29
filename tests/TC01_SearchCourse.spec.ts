const { test } = require('@playwright/test');
const test_ExcelJS = require('exceljs');

test.only("Search Courses", async ({ page }) => {
    const workbook = new test_ExcelJS.Workbook();
    test.setTimeout(250000);
    await workbook.xlsx.readFile("C:\\Users\\Vivo\\Desktop\\Test_Project\\tests\\01_Data_Search_Course.xlsx");
    const worksheet = workbook.getWorksheet(1);

    // เข้าสู่เว็บไซต์และหน้า Search Course
    await page.goto('http://localhost:8083/sci_mju_lifelonglearning/', { waitUntil: 'domcontentloaded' });
    await page.waitForFunction(() => document.querySelector('title')?.textContent === 'Science MJU LifeLong Learning');
    await page.goto('http://localhost:8083/sci_mju_lifelonglearning/search_course');

    let row = 2; 
    let round = 0; 
    while (row <= 6) {  // ปรับเป็นทำงานตั้งแต่แถว 2 ถึงแถว 6

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
            const DataPagesearch_course = await elementHandle.textContent(); // ดึงข้อความจากองค์ประกอบ
            if (DataPagesearch_course.includes(DataPageResult)) {
                console.log(`Row ${row}: True`);
                worksheet.getCell(`E${row}`).value = 'True';
            } else {
                console.log(`Row ${row}: False`);
                worksheet.getCell(`E${row}`).value = 'False';
            }
        } else {
            console.log(`Row ${row}: False - No search result found`);
            worksheet.getCell(`E${row}`).value = 'False'; // กรณีที่ไม่พบผลลัพธ์การค้นหา
        }

        row++; // ไปยังแถวถัดไป
        round++; // เพิ่มตัวนับรอบ
    }

    await workbook.xlsx.writeFile("C:\\Users\\Vivo\\Desktop\\Test_Project\\tests\\01_Data_Search_Course.xlsx");
});
