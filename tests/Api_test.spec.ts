import {test, expect} from '@playwright/test';

test.only('API GET Require', async ({ request }) => {
    const response = await request.get('') //ใส่ลิ้ง API ที่ต้องการทดสอบ (เมธอท GET)
    expect(response.status()).toBe(200)
    const text = await response.text();
    expect(text).toContain('') //ชือ คอส, ชื่อ คน ที่เอามาทดสอบ
    console.log(await response.json());//ปริ้นผลลัพธ์การทดสอบออกมา
})

test.only('API POST Require', async ({ request }) => {
    //ใส่ลิ้ง API และ ข้อมูล ที่ต้องการทดสอบ (เมธอท POST)
    const response = await request.post('',{
        data : {
            "Test1" : "Test",
            "Test2" : "Test"
        }
    }) 
    expect(response.status()).toBe(201)
    const text = await response.text();
    expect(text).toContain('') //ชือ คอส, ชื่อ คน ที่เอามาทดสอบ
    console.log(await response.json());//ปริ้นผลลัพธ์การทดสอบออกมา

})

test.only('API PUT Require', async ({ request }) => {
    //ใส่ลิ้ง API และ ข้อมูล ที่ต้องการทดสอบ (เมธอท PUT)
    const response = await request.put('',{
        data : {
            "Test1" : "Test",
            "Test2" : "Test"
        }
    }) 

    expect(response.status()).toBe(202)
    const text = await response.text();
    expect(text).toContain('') //ชือ คอส, ชื่อ คน ที่เอามาทดสอบ
    console.log(await response.json());//ปริ้นผลลัพธ์การทดสอบออกมา
    
})

test.only('API DELETE Require', async ({ request }) => {
    const response = await request.delete('') //ใส่ลิ้ง API ที่ต้องการทดสอบ DELETE
    expect(response.status()).toBe(203)

})