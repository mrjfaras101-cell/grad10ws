// دالة مساعدة لإرسال الأحداث والوسوم المخصصة إلى Microsoft Clarity
function trackClarityEvent(eventName, tagKey, tagValue) {
    if (typeof window.clarity === 'function') {
        window.clarity('event', eventName);
        if (tagKey && tagValue !== undefined) {
            window.clarity('set', tagKey, String(tagValue));
        }
    }
}

// التبديل بين تبويبات أوراق العمل
function filterWorksheets(target, btnElement) {
    trackClarityEvent(`tab_${target}`, 'active_worksheet', target);

    // تحديث حالة الأزرار
    const buttons = document.querySelectorAll('.tab-btn');
    buttons.forEach(btn => btn.classList.remove('active'));
    if (btnElement) {
        btnElement.classList.add('active');
    }

    const sections = document.querySelectorAll('.worksheet');
    sections.forEach(sec => {
        if (target === 'all') {
            sec.classList.remove('hidden-tab');
        } else if (sec.id === target) {
            sec.classList.remove('hidden-tab');
        } else {
            sec.classList.add('hidden-tab');
        }
    });
}

// التحقق التلقائي من فراغات الكود لكل ورقة عمل
function checkAnswers(wsNum) {
    const inputs = document.querySelectorAll(`.ws-${wsNum}-input`);
    let correctCount = 0;

    inputs.forEach(input => {
        const userVal = input.value.trim().replace(/\s+/g, ' ');
        const acceptedAnswers = input.dataset.ans.split('|').map(a => a.trim());

        // لحقل رسالة الترحيب في ورقة 1 نقبل أي نص بين علامتي تنصيص أو غير فارغ
        if (input.dataset.anytext === 'true') {
            if (userVal.length >= 2) {
                input.classList.remove('wrong');
                input.classList.add('correct');
                correctCount++;
            } else {
                input.classList.remove('correct');
                input.classList.add('wrong');
            }
            return;
        }

        if (acceptedAnswers.includes(userVal)) {
            input.classList.remove('wrong');
            input.classList.add('correct');
            correctCount++;
        } else {
            input.classList.remove('correct');
            input.classList.add('wrong');
        }
    });

    const banner = document.getElementById(`feedback-ws-${wsNum}`);
    banner.style.display = 'block';
    if (correctCount === inputs.length) {
        trackClarityEvent(`ws_${wsNum}_completed_100`, `ws_${wsNum}_score`, `${correctCount}/${inputs.length}`);
        banner.style.background = '#dcfce7';
        banner.style.color = '#14532d';
        banner.style.border = '1px solid #22c55e';
        banner.innerHTML = `🎉 أحسنت يا بطل! جميع إجاباتك صحيحة (${correctCount} من ${inputs.length}). مشروعك البرمجي جاهز للتشغيل!`;
    } else {
        trackClarityEvent(`ws_${wsNum}_check_attempt`, `ws_${wsNum}_score`, `${correctCount}/${inputs.length}`);
        banner.style.background = '#fef2f2';
        banner.style.color = '#7f1d1d';
        banner.style.border = '1px solid #ef4444';
        banner.innerHTML = `💡 حصلت على (${correctCount} من ${inputs.length}). راجع الحقول الملونة بالأحمر واستعن بالتلميحات الذكية!`;
    }
}

// تعبئة الحل النموذجي تلقائياً داخل الفراغات
function fillSampleAnswers(wsNum) {
    trackClarityEvent(`ws_${wsNum}_revealed_solution`, `ws_${wsNum}_used_hint_solution`, 'true');
    const inputs = document.querySelectorAll(`.ws-${wsNum}-input`);
    inputs.forEach(input => {
        const firstAns = input.dataset.ans.split('|')[0];
        input.value = firstAns;
        input.classList.remove('wrong');
        input.classList.add('correct');
    });
    const banner = document.getElementById(`feedback-ws-${wsNum}`);
    banner.style.display = 'block';
    banner.style.background = '#eff6ff';
    banner.style.color = '#1e3a8a';
    banner.style.border = '1px solid #3b82f6';
    banner.innerHTML = `✨ تم عرض الإجابات النموذجية في الفراغات للمراجعة والتعلم الذاتي.`;
}

// إعادة تعيين الفراغات
function resetAnswers(wsNum) {
    trackClarityEvent(`ws_${wsNum}_reset`);
    const inputs = document.querySelectorAll(`.ws-${wsNum}-input`);
    inputs.forEach(input => {
        input.value = '';
        input.classList.remove('correct', 'wrong');
    });
    const banner = document.getElementById(`feedback-ws-${wsNum}`);
    banner.style.display = 'none';
}

// محاكي شاشة التنفيذ (ورقة عمل 1 - كاشير المقصف بالدينار الأردني)
function runSim1() {
    trackClarityEvent('simulator_run_ws1');
    const name = document.getElementById('sim1-name').value.trim() || 'أحمد';
    const price = parseFloat(document.getElementById('sim1-price').value) || 4;
    const qty = parseInt(document.getElementById('sim1-qty').value) || 3;

    let total = price * qty;
    let html = `*** أهلاً بك في كاشير مقصف مدرسة صويلح الثانوية للبنين ***<br>`;
    html += `أدخل اسمك: <span class="user-input">${name}</span><br>`;
    html += `أدخل سعر الوجبة بالدينار الأردني: <span class="user-input">${price}</span><br>`;
    html += `أدخل عدد الوجبات: <span class="user-input">${qty}</span><br>`;
    html += `المجموع قبل الخصم: ${total.toFixed(1)} دينار أردني<br>`;

    if (total >= 10) {
        total = total - 2;
        html += `🎉 مبروك يا ${name}! حصلت على خصم بقيمة 2 دينار أردني!<br>`;
    } else {
        html += `شكراً لك يا ${name}، لا يوجد خصم لهذه الفاتورة (أقل من 10 دنانير).<br>`;
    }
    html += `<strong>المبلغ النهائي المطلوب دفعه: ${total.toFixed(1)} دينار أردني</strong>`;

    document.getElementById('term-body-1').innerHTML = html;
}

// محاكي شاشة التنفيذ (ورقة عمل 2 - المرشد الأكاديمي)
function runSim2() {
    trackClarityEvent('simulator_run_ws2');
    const g1 = parseInt(document.getElementById('sim2-g1').value) || 92;
    const g2 = parseInt(document.getElementById('sim2-g2').value) || 95;
    const g3 = parseInt(document.getElementById('sim2-g3').value) || 98;
    const absence = parseInt(document.getElementById('sim2-abs').value);
    const absVal = isNaN(absence) ? 0 : absence;

    const avg = (g1 + g2 + g3) / 3;
    let gradeText = '';
    if (avg >= 90) {
        gradeText = 'ممتاز';
    } else if (avg >= 75) {
        gradeText = 'جيد جداً';
    } else if (avg >= 50) {
        gradeText = 'ناجح';
    } else {
        gradeText = 'يحتاج إلى تحسين';
    }

    let html = `أدخل علامات الاختبارات الثلاثة: <span class="user-input">${g1} ، ${g2} ، ${g3}</span><br>`;
    html += `أدخل عدد أيام الغياب: <span class="user-input">${absVal}</span><br>`;
    html += `معدلك الحسابي هو: ${avg.toFixed(1)}<br>`;
    html += `التقدير: <strong>${gradeText}</strong><br>`;

    if (avg >= 90 && absVal === 0) {
        html += `🏆 مبارك! تستحق شهادة التفوق والانضباط ومكافأة 25 ديناراً أردنياً.`;
    } else {
        html += `ℹ️ ملاحظة: شهادة التفوق ومكافأة الـ 25 ديناراً تتطلب معدلاً 90 فأكثر وغياباً يساوي 0.`;
    }

    document.getElementById('term-body-2').innerHTML = html;
}

// محاكي شاشة التنفيذ (ورقة عمل 3 - بوابة مدينة الألعاب بالدينار الأردني)
function runSim3() {
    trackClarityEvent('simulator_run_ws3');
    const age = parseInt(document.getElementById('sim3-age').value) || 15;
    const height = parseInt(document.getElementById('sim3-height').value) || 160;
    const ticket = parseInt(document.getElementById('sim3-ticket').value) || 104;

    let html = `أدخل عمرك: <span class="user-input">${age}</span> | طولك (سم): <span class="user-input">${height}</span> | رقم تذكرتك: <span class="user-input">${ticket}</span><br>`;

    if (age >= 12 && height >= 140) {
        html += `✅ مسموح لك بركوب قطار المغامرات السريع!<br>`;
        if (age < 16 || age >= 60) {
            html += `🎟️ شملتك فئة الخصم! سعر تذكرتك: <strong>3 دنانير أردنية</strong>.<br>`;
        } else {
            html += `🎟️ سعر تذكرتك الأساسي: <strong>5 دنانير أردنية</strong>.<br>`;
        }

        if (ticket % 2 === 0) {
            html += `🎁 رقم تذكرتك (${ticket}) زوجي! ربحت قسيمة مجانية بقيمة 2 دينار أردني.`;
        } else {
            html += `رقم تذكرتك (${ticket}) فردي، نتمنى لك حظاً أوفر في السحب القادم!`;
        }
    } else {
        html += `❌ عذراً، شروط الأمان (العمر 12+ والطول 140+ سم) غير مستوفاة.`;
    }

    document.getElementById('term-body-3').innerHTML = html;
}

// نسخ كود الحل إلى الحافظة
function copySolutionCode(elementId, btn) {
    const codeText = document.getElementById(elementId).innerText;
    navigator.clipboard.writeText(codeText).then(() => {
        const oldText = btn.innerText;
        btn.innerText = '✅ تم النسخ!';
        setTimeout(() => {
            btn.innerText = oldText;
        }, 2000);
    });
}

// طباعة أوراق العمل (مع خيار تضمين أو استثناء صفحة الحلول)
function printWorksheets(includeSolutions) {
    const solSection = document.getElementById('ws-solutions');
    if (!includeSolutions) {
        solSection.classList.add('hide-in-print');
    } else {
        solSection.classList.remove('hide-in-print');
    }
    window.print();
}

// تحميل ذكي لملفات الوورد سواء كانت في المجلد الرئيسي أو داخل مجلد word-files
async function smartDownload(event, filename) {
    event.preventDefault();
    trackClarityEvent('word_file_download', 'downloaded_file', filename);
    const directUrl = `./${filename}`;
    const folderUrl = `./word-files/${filename}`;
    try {
        const res = await fetch(directUrl, { method: 'HEAD' });
        if (res.ok) {
            triggerDownload(directUrl, filename);
            return false;
        }
    } catch (e) {}
    try {
        const res2 = await fetch(folderUrl, { method: 'HEAD' });
        if (res2.ok) {
            triggerDownload(folderUrl, filename);
            return false;
        }
    } catch (e) {}
    // في حال لم يرفع المستخدم ملفات الوورد بعد على GitHub، نوجهه للرابط المباشر
    window.location.href = directUrl;
    return false;
}

function triggerDownload(url, filename) {
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
}

