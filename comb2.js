const puppeteer = require('puppeteer');

const waitForTimeout = ms => new Promise(res => setTimeout(res, ms));

(async () => {
    const browser = await puppeteer.launch({ headless: false });
    const page = await browser.newPage();
    await page.setUserAgent('Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36');

    // 1. Truy cập trang collection
    await page.goto('https://forms.office.com/Pages/DesignPageV2.aspx?collectionid=p0nj1k3ofrkcsv6cnnrrml', { waitUntil: 'networkidle2' });

    if (page.url().includes('login.microsoftonline.com')) {
        console.log('Vui lòng đăng nhập trong 60 giây...');
        await waitForTimeout(60000);
    }

    // 2. Đợi các form hiện ra
    await page.waitForSelector('[data-automation-id="itemContainer"]', { timeout: 60000 });
    const formElements = await page.$$('div[data-automation-id="itemContainer"]');

    const formList = [];

    for (let i = 0; i < formElements.length; i++) {
        const formEl = formElements[i];

        const title = await formEl.$eval('[data-automation-id="detailTitle"]', el => el.textContent.trim());

        const [formPage] = await Promise.all([
            new Promise(resolve => page.once('popup', resolve)),
            formEl.click()
        ]);

        await formPage.waitForNavigation({ waitUntil: 'networkidle2' });
        const formUrl = formPage.url();

        try {
            await formPage.waitForFunction(() => window.location.href.includes('DesignPageV2.aspx'), { timeout: 30000 });

            // Mở chỉnh sửa nếu có
            await formPage.evaluate(() => {
                const editBtn = document.querySelector('button[data-automation-id="editFormButton"]') ||
                                document.querySelector('button[title*="Edit"]') ||
                                document.querySelector('button[title*="Preview"]');
                if (editBtn) editBtn.click();
            });
            await waitForTimeout(5000);

            // Đợi câu hỏi xuất hiện
            await formPage.waitForSelector('div[data-automation-id="questionTitle"] span.text-format-content', { timeout: 30000 });

            // Trích xuất câu hỏi và đáp án
            const questions = await formPage.evaluate(() => {
                const result = [];
                const questionElements = document.querySelectorAll('div[data-automation-id="questionTitle"]');

                for (const questionElement of questionElements) {
                    const questionText = questionElement.querySelector('span.text-format-content')?.textContent.trim();
                    if (!questionText) continue;

                    if (['Họ và tên', 'HỌ VÀ TÊN NHÂN VIÊN', 'Mã nhân viên', 'SỐ NHÂN VIÊN'].includes(questionText)) continue;

                    let current = questionElement;
                    let answerContainer = null;

                    while (current) {
                        current = current.nextElementSibling;
                        if (current?.matches('div[role="radiogroup"], div[role="group"]')) {
                            answerContainer = current;
                            break;
                        }
                        if (current?.matches('div[data-automation-id="questionTitle"]')) break;
                    }

                    if (!answerContainer) {
                        let parent = questionElement.parentElement;
                        while (parent) {
                            answerContainer = parent.querySelector('div[role="radiogroup"], div[role="group"]');
                            if (answerContainer) break;
                            parent = parent.parentElement;
                        }
                    }

                    if (!answerContainer) continue;

                    const type = answerContainer.getAttribute('role') === 'radiogroup' ? 'single' : 'multiple';
                    const answers = Array.from(answerContainer.querySelectorAll('span.text-format-content'))
                        .map(el => el.textContent.trim())
                        .filter(Boolean);

                    result.push({ question: questionText, answers, type });
                }

                return result;
            });

            formList.push({ title, url: formUrl, questions });
        } catch (e) {
            console.error('Lỗi khi xử lý form:', title, formUrl, e);
        } finally {
            await formPage.close();
        }
    }

    console.log('\n✅ Kết quả:');
    console.dir(formList, { depth: null });

    await browser.close();
})();
