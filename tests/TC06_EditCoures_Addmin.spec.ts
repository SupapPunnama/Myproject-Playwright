const { test: Edit_Course_Admin } = require('@playwright/test');
const ExcelJS_Edit_coures_Admin = require('exceljs');

Edit_Course_Admin.only("Edit Course Admin", async ({ page }) => {
    Edit_Course_Admin.setTimeout(2000000);
    const workbook = new ExcelJS_Edit_coures_Admin.Workbook();
    await workbook.xlsx.readFile("C:\\Users\\Vivo\\Desktop\\Test_Project\\tests\\06_Data_Edit_Coures_Admin.xlsx");

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

    await page.goto('http://localhost:8083/sci_mju_lifelonglearning/course/C003/edit_course');


    const worksheet = workbook.getWorksheet(1);
    let row = 2; // เริ่มต้นที่แถวที่ 2
    let round = 0; // ตั้งค่าตัวนับรอบ

    while (round < 107) {

        const processFlag = worksheet.getCell(`X${row}`).value; // ตรวจสอบค่าในเซลล์ X{row}
        console.log(processFlag);

        if (processFlag !== 'Yes') {
            row++;
            round++;
            continue;
        }

        await page.reload();


        const course_type = worksheet.getCell(`B${row}`).value; // ประเภทหลักสูตร
        if (course_type) {
            // รอให้ select element ปรากฏ
            await page.waitForSelector('#course_type', { visible: true });
            // ตรวจสอบว่ามี option ที่ตรงกับ course_type หรือไม่
            const optionExists = await page.$eval('#course_type', (select, course_type) => {
                return Array.from(select.options).some(option =>
                    (option as HTMLOptionElement).textContent?.trim() === course_type.trim()
                );
            }, course_type);

            if (optionExists) {
                // ถ้ามี option ที่ตรงกัน ให้เลือก option นั้น
                await page.selectOption('#course_type', { label: course_type.trim() });
            } else {
                console.log(`ตัวเลือก '${course_type}' ไม่มีในรายการ`);
            }
        }

        const major_id = worksheet.getCell(`C${row}`).value; // สาขา
        if (major_id) {
            await page.waitForSelector('#major_id', { visible: true });
            const optionValues = await page.$$eval('#major_id option', options =>
                options.map(option => option.textContent?.trim() || '')
            );

            if (optionValues.includes(major_id.trim())) {
                await page.selectOption('#major_id', { label: major_id.trim() });
            } else {
                console.log(`ตัวเลือก '${major_id}' ไม่มีในรายการ`);
            }
        }


        const course_name = worksheet.getCell(`D${row}`).value; // ชื่อหลักสูตร
        if (course_name) {
            await page.waitForSelector('#course_name', { visible: true });
            await page.fill('#course_name', course_name.toString());
            await page.waitForSelector('a#link', { visible: true });
            await page.click('a#link');
        }


        const certificateName = worksheet.getCell(`E${row}`).value; // ชื่อเกียรติบัตร
        if (certificateName) {
            await page.waitForSelector('#certificateName', { visible: true });
            await page.fill('#certificateName', certificateName.toString());
        }


        const course_img = worksheet.getCell(`F${row}`).value; // รูปภาพหลักสูตร
        if (course_img) {
            await page.waitForSelector('#fileInput', { visible: true });
            await page.setInputFiles('#fileInput', course_img.toString());

            // รออะเลิท 3 วินาที
            try {
                const alertMessage = await page.waitForEvent('dialog', { timeout: 3000 });
                console.log(`พบอะเลิท: ${alertMessage.message()}`);
                worksheet.getCell(`V${row}`).value = alertMessage.message();
                await alertMessage.accept();

                row++;
                round++;
                continue;

            } catch (error) {
                console.log('ไม่มีอะเลิท');
            }

        }

        const course_principle = worksheet.getCell(`G${row}`).value; // หลักการและเหตุผล
        if (course_principle) {
            await page.waitForSelector('#editor', { visible: true });
            await page.fill('#editor', course_principle.toString());
        }


        const nextBtn = await page.$('#nextBtn');
        await nextBtn.scrollIntoViewIfNeeded();
        await nextBtn.click();

        const invalidCourseType = await page.textContent('#invalidCourseType');
        const invalidMajor = await page.textContent('#invalidMajor');
        const status = await page.textContent('#status');
        const invalidCertificateName = await page.textContent('#invalidCertificateName');
        const invalidImg = await page.textContent('#invalidImg');
        const invalidPrinciple = await page.textContent('#invalidPrinciple');

        if (status.trim() === "สามารถใช้งานได้" && !invalidCourseType.trim() && !invalidMajor.trim() && !invalidCertificateName.trim()
            && !invalidImg.trim() && !invalidPrinciple.trim()) {
            console.log('ฟิลด์ทั้งหมดถูกต้อง ดำเนินการต่อหน้า 2');

            const course_object = worksheet.getCell(`H${row}`).value; // วัตถุประสงค์
            if (course_object) {
                await page.waitForSelector('#course_object', { visible: true });
                await page.fill('#course_object', course_object.toString());
            }

            const CallFeeValue = worksheet.getCell(`I${row}`).value; // อัตราค่าบริการ
            if (CallFeeValue) {
                const CallFee = CallFeeValue.toString();
                if (CallFee === 'ไม่มีค่าธรรมเนียม') {
                    const NoncallFeeRadioButton = await page.locator("//td//input[@value='ไม่มีค่าธรรมเนียม']");
                    if (await NoncallFeeRadioButton.count() > 0) {
                        await NoncallFeeRadioButton.first().check();
                    }
                } else if (CallFee === 'มีค่าธรรมเนียม') {
                    const CallFeeRadioButton = await page.locator("//td//input[@value='มีค่าธรรมเนียม']");
                    if (await CallFeeRadioButton.count() > 0) {
                        await CallFeeRadioButton.first().check();

                        const course_fee = worksheet.getCell(`J${row}`).value; // จำนวนเงิน
                        if (course_fee) {
                            const validCourseFee = course_fee.toString().replace(/\D/g, ''); // ลบอักขระที่ไม่ใช่ตัวเลข
                            if (validCourseFee) {
                                await page.waitForSelector('#course_fee', { visible: true });
                                await page.fill('#course_fee', validCourseFee);
                            } else {
                                console.log(`ค่าธรรมเนียมที่ไม่ถูกต้องในแถว ${row}`);
                            }
                        }

                    }
                }
            }
            const course_totalHours = worksheet.getCell(`K${row}`).value; // เวลาในการเรียน
            if (course_totalHours) {
                // ลบอักขระที่ไม่ใช่ตัวเลขเพื่อให้เข้ากับ input[type=number]
                const validtotalHours = course_totalHours.toString().replace(/\D/g, ''); // ลบตัวอักษรที่ไม่ใช่ตัวเลข
                if (validtotalHours) {
                    await page.waitForSelector('#course_totalHours', { visible: true });
                    await page.fill('#course_totalHours', validtotalHours); // กรอกข้อมูลที่เป็นตัวเลข
                } else {
                    console.log(`เวลาเรียนไม่ถูกต้องในแถว ${row}`);
                }
            }

            const course_file = worksheet.getCell(`L${row}`).value; // เอกสารหลักสูตร
            if (course_file) {
                await page.waitForSelector('#course_file', { visible: true });
                await page.setInputFiles('#course_file', course_file.toString());

                // รออะเลิท 3 วินาที
                try {
                    const alertMessage = await page.waitForEvent('dialog', { timeout: 3000 });
                    console.log(`พบอะเลิท: ${alertMessage.message()}`);
                    worksheet.getCell(`V${row}`).value = alertMessage.message();
                    await alertMessage.accept();


                    row++;
                    round++;
                    continue;

                } catch (error) {
                    console.log('ไม่มีอะเลิท');
                }

            }


            const floatingTextarea3 = worksheet.getCell(`M${row}`).value; // กลุ่มเป้าหมายอาชีพ
            if (floatingTextarea3) {
                await page.waitForSelector('#floatingTextarea3', { visible: true });
                await page.fill('#floatingTextarea3', floatingTextarea3.toString());
            }

            const nextBtn = await page.$('#nextBtn');
            await nextBtn.scrollIntoViewIfNeeded();
            await nextBtn.click();

            const invalidObjective = await page.textContent('#invalidObjective');
            const invalidCourseFee = await page.textContent('#invalidCourseFee');
            const invalidCourseTotalHours = await page.textContent('#invalidCourseTotalHours');
            const invalidCourseFile = await page.textContent('#invalidCourseFile');
            const invalidCourseTargetOccupation = await page.textContent('#invalidCourseTargetOccupation');

            if (!invalidObjective.trim() && !invalidCourseFee.trim() && !invalidCourseTotalHours.trim() && !invalidCourseFile.trim()
                && !invalidCourseTargetOccupation.trim()) {
                console.log('ฟิลด์ทั้งหมดถูกต้อง ดำเนินการต่อหน้า 3');

                const prefix = worksheet.getCell(`N${row}`).value; // คำนำหน้า
                if (prefix) {
                    await page.waitForSelector('#prefix', { visible: true });
                    await page.selectOption('#prefix', { label: prefix.toString() });
                }

                const fname_contacts = worksheet.getCell(`O${row}`).value; // ชื่อผู้ติดต่อ
                if (fname_contacts) {
                    await page.waitForSelector('#fname_contacts', { visible: true });
                    await page.fill('#fname_contacts', fname_contacts.toString());
                }

                const lname_contacts = worksheet.getCell(`P${row}`).value; // นามสกุลผู้ติดต่อ
                if (lname_contacts) {
                    await page.waitForSelector('#lname_contacts', { visible: true });
                    await page.fill('#lname_contacts', lname_contacts.toString());
                }

                const faculty = worksheet.getCell(`Q${row}`).value; // คณะ
                if (faculty) {
                    await page.waitForSelector('#faculty', { visible: true });
                    await page.fill('#faculty', faculty.toString());
                }

                const phone_contacts = worksheet.getCell(`R${row}`).value; // เบอร์โทรศัพท์
                if (phone_contacts) {
                    const validPhoneNumber = phone_contacts.toString().replace(/\D/g, ''); // ลบตัวอักษรที่ไม่ใช่ตัวเลข
                    if (validPhoneNumber) {
                        await page.waitForSelector('#phone_contacts', { visible: true });
                        await page.fill('#phone_contacts', validPhoneNumber); // กรอกข้อมูลที่เป็นตัวเลขเท่านั้น
                    } else {
                        console.log(`เบอร์โทรศัพท์ไม่ถูกต้องในแถว ${row}`);
                    }
                }


                const email_contacts = worksheet.getCell(`S${row}`).value; // อีเมล
                if (email_contacts) {
                    await page.waitForSelector('#email_contacts', { visible: true });
                    await page.fill('#email_contacts', email_contacts.toString());
                }

                const nextBtn = await page.$('#nextBtn');
                await nextBtn.scrollIntoViewIfNeeded();
                await nextBtn.click();

                const invalidSelectPrefix = await page.textContent('#invalidSelectPrefix');
                const invalidFNameContacts = await page.textContent('#invalidFNameContacts');
                const invalidLNameContacts = await page.textContent('#invalidLNameContacts');
                const invalidFaculty = await page.textContent('#invalidFaculty');
                const invalidPhoneContacts = await page.textContent('#invalidPhoneContacts');
                const invalidEmailContacts = await page.textContent('#invalidEmailContacts');

                if (!invalidSelectPrefix.trim() && !invalidFNameContacts.trim() && !invalidLNameContacts.trim() && !invalidFaculty.trim()
                    && !invalidPhoneContacts.trim() && !invalidEmailContacts.trim()) {
                    console.log('บันทึกสำเร็จ');
                    worksheet.getCell(`V${row}`).value = 'บันทึกสำเร็จ';
                } else {
                    console.log('ข้อผิดพลาดหน้า 3:', invalidSelectPrefix, invalidFNameContacts, invalidLNameContacts, invalidFaculty,
                        invalidPhoneContacts, invalidEmailContacts);
                    worksheet.getCell(`V${row}`).value = invalidSelectPrefix || invalidFNameContacts || invalidLNameContacts || invalidFaculty
                        || invalidPhoneContacts || invalidEmailContacts || 'ข้อผิดพลาดไม่รู้จัก หน้า3';
                }
            } else {
                console.log('ข้อผิดพลาดหน้า 2:', invalidObjective, invalidCourseFee, invalidCourseTotalHours, invalidCourseFile, invalidCourseTargetOccupation);
                worksheet.getCell(`V${row}`).value = invalidObjective || invalidCourseFee || invalidCourseTotalHours || invalidCourseFile
                    || invalidCourseTargetOccupation || 'ข้อผิดพลาดไม่รู้จัก หน้า2';
            }
        } else {
            console.log('ข้อผิดพลาดหน้า 1:', invalidCourseType, invalidMajor, invalidCertificateName, invalidImg, invalidPrinciple);
            worksheet.getCell(`V${row}`).value = invalidCourseType || invalidMajor || invalidCertificateName || invalidImg || invalidPrinciple || 'ข้อผิดพลาดไม่รู้จัก หน้า1';
        }


        const cellT = worksheet.getCell(`T${row}`).value;
        const cellV = worksheet.getCell(`V${row}`).value;
        worksheet.getCell(`W${row}`).value = cellT === cellV ? 'Pass' : 'Fail';
        await workbook.xlsx.writeFile("C:\\Users\\Vivo\\Desktop\\Test_Project\\tests\\05_Data_AddCoures_Admin.xlsx");

        row++;
        round++;
        continue;
    }
    await workbook.xlsx.writeFile("C:\\Users\\Vivo\\Desktop\\Test_Project\\tests\\06_Data_Edit_Coures_Admin.xlsx");
});