import {test, expect} from '@playwright/test';

test.only('API GET Require', async ({ request }) => {

    const response = await request.get('') //ใส่ลิ้ง API ที่ต้องการทดสอบ (เมธอท GET)
    expect(response.status()).toBe(200)

    const text = await response.text();
    expect(text).toContain('') //ชือ คอส, ชื่อ คน ที่เอามาทดสอบ

    console.log(await response.json());

})