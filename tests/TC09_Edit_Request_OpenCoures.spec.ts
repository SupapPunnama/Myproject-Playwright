const { test: Edit_Request_open_course } = require('@playwright/test');
const ExcelJS_Edit_Request_open_course = require('exceljs');

Edit_Request_open_course.only("Edit Request open course", async ({ page }) => {
    Edit_Request_open_course.setTimeout(350000);

    const workbook = new ExcelJS_Edit_Request_open_course.Workbook();
    await workbook.xlsx.readFile("C:\\Users\\Vivo\\Desktop\\Test_Project\\tests\\09_Data_Edit_Open_Course.xlsx");

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

    //await page.goto('http://localhost:8083/sci_mju_lifelonglearning/lecturer/Lecturer01/list_request_open_course');
    await page.goto('http://localhost:8083/sci_mju_lifelonglearning/lecturer/Lecturer01/2/update_page');


    const worksheet = workbook.getWorksheet(1);
    let row = 2; // เริ่มต้นที่แถวที่ 2
    let round = 0; // ตั้งค่าตัวนับรอบ

    while (round < 85) { //85

        const processFlag = worksheet.getCell(`W${row}`).value; // ตรวจสอบค่าในเซลล์ 
        console.log(processFlag);

        if (processFlag !== 'Yes') {
            round++;
            row++;
            continue;
        }

        await page.reload();

        const course_name = worksheet.getCell(`B${row}`).value; // หลักสูตร
        if (course_name) {
            const course = course_name.toString();
            if (course === 'หลักสูตรทั้งหมด') {
                const Course_1_RadioButton = await page.locator("//td//input[@value='หลักสูตรทั้งหมด']");
                if (await Course_1_RadioButton.count() > 0) {
                    await Course_1_RadioButton.first().check();
                }
            } else if (course === 'หลักสูตรอบรมระยะสั้น') {
                const Course_2_RadioButton = await page.locator("//td//input[@value='หลักสูตรอบรมระยะสั้น']");
                if (await Course_2_RadioButton.count() > 0) {
                    await Course_2_RadioButton.first().check();
                }
            } else if (course === 'Non-Degree') {
                const Course_3_RadioButton = await page.locator("//td//input[@value='Non-Degree']");
                if (await Course_3_RadioButton.count() > 0) {
                    await Course_3_RadioButton.first().check();
                }
            }

            const inputGroupSelect02 = worksheet.getCell(`C${row}`).value; // เลือกหลักสูตร
            if (inputGroupSelect02) {
                await page.waitForSelector('#inputGroupSelect02', { visible: true });
                const optionValues = await page.$$eval('#inputGroupSelect02 option', options =>
                    options.map(option => option.textContent?.trim() || '')
                );

                if (optionValues.includes(inputGroupSelect02.trim())) {
                    await page.selectOption('#inputGroupSelect02', { label: inputGroupSelect02.trim() });
                } else {
                    console.log(`ตัวเลือก '${inputGroupSelect02}' ไม่มีในรายการ`);
                }
            }
        }

        const convertBuddhistToGregorian = (buddhistDate) => {
            if (!buddhistDate) return null;

            try {
                const dateParts = buddhistDate.split('/');
                if (dateParts.length !== 3) throw new Error('Invalid date format');

                const buddhistYear = parseInt(dateParts[2], 10);
                const christianYear = buddhistYear - 543;

                // ตรวจสอบและสร้างวันที่ในรูปแบบ YYYY-MM-DD
                const formattedDate = `${christianYear}-${dateParts[1].padStart(2, '0')}-${dateParts[0].padStart(2, '0')}`;
                if (!/^\d{4}-\d{2}-\d{2}$/.test(formattedDate)) throw new Error('Date format incorrect');

                return formattedDate;
            } catch (error) {
                console.error(`Error converting date: ${buddhistDate}`, error);
                return null;
            }
        };




        const startRegister = convertBuddhistToGregorian(worksheet.getCell(`D${row}`).value); // วันเปิดรับสมัคร
        if (startRegister) {
            await page.fill('#startRegister', startRegister);
        }

        const endRegister = convertBuddhistToGregorian(worksheet.getCell(`E${row}`).value); // วันปิดรับสมัคร
        if (endRegister) {
            await page.fill('#endRegister', endRegister);
        }

        const quantity = worksheet.getCell(`F${row}`).value; // จำนวนนักเรียน
        if (quantity) {
            const validquantity = quantity.toString().replace(/\D/g, '');
            await page.waitForSelector('#quantity', { visible: true });
            if (validquantity) {
                await page.waitForSelector('#quantity', { visible: true });
                await page.fill('#quantity', quantity);
            } else {
                console.log(`จำนวนคนไม่ถูกต้องในแถว ${row}`);
            }

        }

        const startPayment = convertBuddhistToGregorian(worksheet.getCell(`G${row}`).value); // วันเริ่มชำระเงิน
        if (startPayment) {
            await page.fill('#startPayment', startPayment);
        }

        const endPayment = convertBuddhistToGregorian(worksheet.getCell(`H${row}`).value); // วันสิ้นสุดชำระเงิน
        if (endPayment) {
            await page.fill('#endPayment', endPayment);
        }

        const applicationResult = convertBuddhistToGregorian(worksheet.getCell(`I${row}`).value); // วันประกาศผลการสมัค
        if (applicationResult) {
            await page.fill('#applicationResult', applicationResult);
        }


        const nextBtn = await page.$('#nextBtn');
        await nextBtn.scrollIntoViewIfNeeded();
        await nextBtn.click();

        await page.waitForTimeout(1000);

        const invalidCourse_Select = await page.textContent('#invalidCourse_Select');
        const invalidStartRegister = await page.textContent('#invalidStartRegister');
        const invalidEndRegister = await page.textContent('#invalidEndRegister');
        const invalidQuantity = await page.textContent('#invalidQuantity');
        const invalidStartPayment = await page.textContent('#invalidStartPayment');
        const invalidEndPayment = await page.textContent('#invalidEndPayment');
        const invalidApplicationResult = await page.textContent('#invalidApplicationResult');

        if (!invalidCourse_Select.trim() && !invalidStartRegister.trim() && !invalidEndRegister.trim() && !invalidQuantity.trim() &&
            !invalidStartPayment.trim() && !invalidEndPayment.trim() && !invalidApplicationResult.trim()) {
            console.log('ฟิลด์ทั้งหมดถูกต้อง ดำเนินการต่อหน้า 2');


            const days = ['mo', 'tu', 'we', 'th', 'fr', 'sa', 'su'];
            const studyDayMapping = {
                'วันจันทร์': 'mo',
                'วันอังคาร': 'tu',
                'วันพุธ': 'we',
                'วันพฤหัสบดี': 'th',
                'วันศุกร์': 'fr',
                'วันเสาร์': 'sa',
                'วันอาทิตย์': 'su'
            };

            const studyDay = worksheet.getCell(`J${row}`).value; // วันที่เรียน
            if (studyDay && studyDayMapping[studyDay]) {
                const dayAbbreviation = studyDayMapping[studyDay];
                const dayRadioButton = await page.locator(`//label[contains(text(),'${studyDay}')]`);

                if (await dayRadioButton.count() > 0) {
                    await dayRadioButton.first().click();

                    const startTime = worksheet.getCell(`K${row}`).value; // เวลาเริ่มเรียน
                    const endTime = worksheet.getCell(`L${row}`).value; // เวลาเลิกเรียน

                    if (startTime && endTime) {
                        const startTimeStr = startTime.toString().trim();
                        const endTimeStr = endTime.toString().trim();

                        // รอให้ element แสดงผล
                        await page.waitForSelector(`#start_chk_${dayAbbreviation}`, { state: 'visible' });

                        // ตรวจสอบว่า element ไม่ถูกปิดใช้งาน (disabled)
                        const startSelect = await page.locator(`#start_chk_${dayAbbreviation}`);
                        if (!(await startSelect.isDisabled())) {
                            await page.selectOption(`#start_chk_${dayAbbreviation}`, { label: startTimeStr });
                        }

                        await page.waitForSelector(`#end_chk_${dayAbbreviation}`, { state: 'visible' });
                        const endSelect = await page.locator(`#end_chk_${dayAbbreviation}`);
                        if (!(await endSelect.isDisabled())) {
                            await page.selectOption(`#end_chk_${dayAbbreviation}`, { label: endTimeStr });
                        }
                    } else {
                        console.log(`เวลาที่เริ่มและสิ้นสุดไม่ถูกต้องในแถว ${row}`);
                    }
                } else {
                    console.log(`ไม่พบตัวเลือกวัน ${studyDay} ในแถว ${row}`);
                }
            } else {
                console.log(`วันเรียนไม่ถูกต้องในแถว ${row}`);
            }



            const startStudyDate = convertBuddhistToGregorian(worksheet.getCell(`M${row}`).value); // วันเริ่มเรียน
            if (startStudyDate) {
                await page.fill('#startStudyDate', startStudyDate);

            }

            const endStudyDate = convertBuddhistToGregorian(worksheet.getCell(`N${row}`).value); // วันสิ้นสุดการเรียน
            if (endStudyDate) {
                await page.fill('#endStudyDate', endStudyDate);

            }


            const nextBtn = await page.$('#nextBtn');
            await nextBtn.scrollIntoViewIfNeeded();
            await nextBtn.click();

            await page.waitForTimeout(1000);

            const invalid_chk_mo = await page.textContent('#invalid_chk_mo');
            const invalid_chk_tu = await page.textContent('#invalid_chk_tu');
            const invalid_chk_we = await page.textContent('#invalid_chk_we');
            const invalid_chk_th = await page.textContent('#invalid_chk_th');
            const invalid_chk_fr = await page.textContent('#invalid_chk_fr');
            const invalid_chk_sa = await page.textContent('#invalid_chk_sa');
            const invalid_chk_su = await page.textContent('#invalid_chk_su');

            const invalidStartStudyDate = await page.textContent('#invalidStartStudyDate');
            const invalidEndStudyDate = await page.textContent('#invalidEndStudyDate');

            if (!invalid_chk_mo.trim() && !invalid_chk_tu.trim() && !invalid_chk_we.trim() && !invalid_chk_th.trim() && !invalid_chk_fr.trim()
                && !invalid_chk_sa.trim() && !invalid_chk_su.trim() && !invalidStartStudyDate.trim() && !invalidEndStudyDate.trim()) {
                console.log('ฟิลด์ทั้งหมดถูกต้อง ดำเนินการต่อหน้า 3');
        


                /*  const type_teach = worksheet.getCell(`O${row}`).value; // ประเภทการเรียน
                  if (type_teach) {
                      const Type = type_teach.toString();
  
                      try {
                          // ตรวจสอบและแสดง select ถ้ามันถูกซ่อนไว้
                          const selectVisible = await page.evaluate(() => {
                              const select = document.querySelector('/html[1]/body[1]/form[1]/div[3]/table[1]/tbody[1]/tr[1]/td[2]/div[1]/div[1]/select[1]') as HTMLElement; // แคสต์เป็น HTMLElement
                              if (select) { // เช็คว่ามี select อยู่หรือไม่
                                  if (getComputedStyle(select).display === 'none') {
                                      select.style.display = 'block'; // เปลี่ยนให้เป็น block เพื่อให้มองเห็นได้
                                  }
                                  return true; // select มองเห็นได้แล้ว
                              }
                              return false; // ไม่มี select หรือไม่สามารถมองเห็นได้
                          });
  
                          // ถ้า select มองไม่เห็นจะทำการโยน error
                          if (!selectVisible) {
                              throw new Error(`Cannot make select element visible for row ${row}`);
                          }
  
                          // เลือกตัวเลือกตามค่าใน Excel
                          if (Type === 'แบบที่ 1 เรียนร่วมกับนักศึกษาในหลักสูตร') {
                              await page.selectOption("/html[1]/body[1]/form[1]/div[3]/table[1]/tbody[1]/tr[1]/td[2]/div[1]/div[1]/select[1]", { label: 'แบบที่ 1 เรียนร่วมกับนักศึกษาในหลักสูตร' });
                          } else if (Type === 'แบบที่ 2 แยกกลุ่มเรียนโดยเฉพาะ') {
                              await page.selectOption("/html[1]/body[1]/form[1]/div[3]/table[1]/tbody[1]/tr[1]/td[2]/div[1]/div[1]/select[1]", { label: 'แบบที่ 2 แยกกลุ่มเรียนโดยเฉพาะ' });
                          } else if (Type === 'จัดการเรียนการสอนร่วมกับทั้งแบบที่ 1 และแบบที่ 2') {
                              await page.selectOption("/html[1]/body[1]/form[1]/div[3]/table[1]/tbody[1]/tr[1]/td[2]/div[1]/div[1]/select[1]", { label: 'จัดการเรียนการสอนร่วมกับทั้งแบบที่ 1 และแบบที่ 2' });
                          } else {
                              throw new Error(`ไม่มี type ตำแหน่ง id='type_teach' ใน Excel row ${row}`);
                          }
  
                      } catch (error) {
                          // บันทึกข้อผิดพลาดลงในเซลล์ U{row}
                          console.error(error.message);
                          worksheet.getCell(`U${row}`).value = error.message;
                          row++;
                          round++;
                          continue; // ไปยังแถวถัดไป
                      }
                  }*/


                // แบบฟิกส์
                const fixedTypeTeach = 'แบบที่ 1 เรียนร่วมกับนักศึกษาในหลักสูตร';

                try {
                    // ตรวจสอบและแสดง select ถ้ามันถูกซ่อนไว้
                    const selectVisible = await page.evaluate(() => {
                        const select = document.querySelector('select[id="type_teach"]') as HTMLElement; // แคสต์เป็น HTMLElement
                        if (select) { // เช็คว่ามี select อยู่หรือไม่
                            if (getComputedStyle(select).display === 'none') {
                                select.style.display = 'block'; // เปลี่ยนให้เป็น block เพื่อให้มองเห็นได้
                            }
                            return true; // select มองเห็นได้แล้ว
                        }
                        return false; // ไม่มี select หรือไม่สามารถมองเห็นได้
                    });
                    // ถ้า select มองไม่เห็นจะทำการโยน error
                    if (!selectVisible) {
                        throw new Error(`Cannot make select element visible for row ${row}`);
                    }
                    // เลือกตัวเลือกตามค่าที่ฟิกซ์
                    await page.selectOption('select[id="type_teach"]', { label: fixedTypeTeach });

                } catch (error) {
                    // บันทึกข้อผิดพลาดลงในเซลล์ U{row}
                    console.error(error.message);
                    worksheet.getCell(`U${row}`).value = error.message;
                    row++;
                    round++;
                    continue;
                }


                const Locetor_Study = worksheet.getCell(`P${row}`).value; // ตรวจสอบรูปแบบการเรียน
                if (Locetor_Study) {
                    const Locetor = Locetor_Study.toString();
                    await page.waitForSelector('#type_learn', { visible: true });
                    if (Locetor === 'เรียนออนไลน์') {
                        await page.selectOption('#type_learn', { label: 'เรียนออนไลน์' });

                        const link_mooc = worksheet.getCell(`Q${row}`).value; // ตรวจสอบลิงก์ถ้าเรียนออนไลน์
                        if (link_mooc) {
                            await page.waitForSelector('#link_mooc', { visible: true });
                            await page.fill('#link_mooc', link_mooc.toString());
                        }
                    } else if (Locetor === 'เรียนที่มหาวิทยาลัย') {
                        await page.selectOption('#type_learn', { label: 'เรียนที่มหาวิทยาลัย' });

                        const building = worksheet.getCell(`R${row}`).value; // กรอกสถานที่เรียนถ้าเรียนที่มหาวิทยาลัย
                        if (building) {
                            await page.waitForSelector('#building', { visible: true });
                            await page.fill('#building', building.toString());
                        }

                    }
                }

                const submitButton = await page.$('#nextBtn');
                await submitButton.scrollIntoViewIfNeeded();
                await submitButton.click();



                // ตรวจสอบฟิลด์เพิ่มเติม
                const invalidTypeTeach = await page.textContent('#invalidTypeTeach');
                const invalidTypeLearn = await page.textContent('#invalidTypeLearn');
                const invalidLinkMooc = await page.textContent('#invalidLinkMooc');
                const invalidLocation = await page.textContent('#invalidLocation');

                if (!invalidTypeTeach.trim() && !invalidTypeLearn.trim() && !invalidLinkMooc.trim() && !invalidLocation.trim()) {
                    console.log('บันทึกสำเร็จ');
                    worksheet.getCell(`U${row}`).value = 'บันทึกสำเร็จ';

                } else {
                    console.log('ข้อผิดพลาดหน้า 3:', invalidTypeTeach, invalidTypeLearn, invalidLinkMooc, invalidLocation);
                    worksheet.getCell(`U${row}`).value = invalidTypeTeach || invalidTypeLearn || invalidLinkMooc || invalidLocation || 'ข้อผิดพลาดไม่รู้จัก';

                }


            } else {
                console.log('ข้อผิดพลาดหน้า 2:', invalid_chk_mo, invalid_chk_tu, invalid_chk_we, invalid_chk_th, invalid_chk_fr, invalid_chk_sa,
                    invalid_chk_su, invalidStartStudyDate, invalidEndStudyDate);
                worksheet.getCell(`U${row}`).value = invalid_chk_mo || invalid_chk_tu || invalid_chk_we || invalid_chk_th || invalid_chk_fr ||
                    invalid_chk_sa || invalid_chk_su || invalidStartStudyDate || invalidEndStudyDate || 'ข้อผิดพลาดไม่รู้จัก';
            }


        } else {
            console.log('ข้อผิดพลาดหน้า 1:', invalidCourse_Select, invalidStartRegister, invalidEndRegister, invalidQuantity,
                invalidStartPayment, invalidEndPayment, invalidApplicationResult);
            worksheet.getCell(`U${row}`).value = invalidCourse_Select || invalidStartRegister || invalidEndRegister || invalidQuantity
            invalidStartPayment || invalidEndPayment || invalidApplicationResult || 'ข้อผิดพลาดไม่รู้จัก';

        }

        const valueS = worksheet.getCell(`S${row}`).value;
        const valueU = worksheet.getCell(`U${row}`).value;
        worksheet.getCell(`V${row}`).value = valueS === valueU ? 'True' : 'False';
        await workbook.xlsx.writeFile("C:\\Users\\Vivo\\Desktop\\Test_Project\\tests\\09_Data_Edit_Open_Course.xlsx");

        row++;
        round++;
        continue;
    }

    await workbook.xlsx.writeFile("C:\\Users\\Vivo\\Desktop\\Test_Project\\tests\\09_Data_Edit_Open_Course.xlsx");
});