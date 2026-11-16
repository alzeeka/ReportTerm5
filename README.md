# ReportTerm5 — مرجع قوالب تقارير التدريب التعاوني (الفصل الخامس المعتمد)

هذا المستودع هو **المرجع الموثوق والمعياري** لإنتاج تقارير التدريب التعاوني بصيغة DOCX وفق القالب الجديد المعتمد من جامعة حائل (كلية علوم وهندسة الحاسب الآلي)، مع الحفاظ على أعلى معايير التنسيق والتكامل البرمجي وحقول وورد التلقائية (SEQ & TOC).

---

## 1. بنية المستودع

```
ReportTerm5/
├── README.md                          ← هذا الملف (المرجع المعياري الشامل والبرومبت الذهبي)
├── template.docx                      ← القالب الأساسي بصيغة DOCX
├── template_term5.docx                ← قالب Term 5 الجديد الجاهز للاستخدام
├── report_builder_core.py             ← المحرك البرمجي لدوال التنسيق، الجداول، والصور، وحقول XML
├── generate_report.py                 ← سكربت التوليد الآلي لتقرير مشروع رعاية (Riaya)
├── make_clean_template.py             ← سكربت تفريغ القالب وتحديث مجلد Tamplate/
├── update_word_fields.py              ← سكربت تحديث حقول Word وتوليد الفهارس وأرقام الصفحات آلياً
├── Tamplate/                          ← مجلد القالب المفكوك (XML) المعتمد الجديد
│   ├── [Content_Types].xml
│   ├── _rels/
│   ├── docProps/
│   └── word/
│       ├── document.xml               ← هيكل المستند الجديد (الغلاف + الفصول الأربعة + المراجع)
│       ├── styles.xml                 ← الأنماط المعتمدة (12pt + line spacing 1.5)
│       ├── settings.xml               ← يحتوي <w:updateFields w:val="true"/>
│       ├── theme/theme1.xml
│       ├── header1.xml                ← رأس الصفحة (فارغ)
│       └── footer1.xml                ← تذييل الصفحة (رقم الصفحة فقط)
└── FinalReportsDOCX/                  ← مجلد التقارير النهائية المنتجة
    └── Riaya_Final_Report.docx        ← ✅ التقرير النهائي الشامل لمشروع رعاية (56 صفحة)
```

---

## 2. أبرز الفروقات والتحديثات في قالب Term 5 مقارنة بالقالب القديم

| العنصر | القالب القديم (ReportLMM) | القالب الجديد المعتمد (ReportTerm5) |
|---|---|---|
| **صفحة الغلاف** | الغلاف التقليدي لتقرير Lamia | **غلاف Term 5 الجديد:** شعار جامعة حائل في المنتصف، اسم الجامعة، الكلية، القسم، اسم جهة التدريب ببنط 26pt، المشرفة الأكاديمية ببنط 20pt، بيانات الطالبة والرقم الجامعي. |
| **هيكل الفصول** | فصول التدريب التعاوني القديمة | **فهرس الفصول الأربعة المعتمد:**<br>• Chapter I: INTRODUCTION (الحديث عن جهة التدريب حصراً)<br>• Chapter II: TASKS & EXPERIENCES (دمج المهام دون ذكر الأشهر)<br>• Chapter III: COOP PROJECT (المشروع التقني مع المعايير الدولية)<br>• Chapter IV: CONCLUSION (الخاتمة والتوصيات)<br>• References (المراجع) |
| **الملحق (Appendix)** | كان موجوداً ويحتوي على قصاصات كود | **تم حذفه بالكامل (Deleted)** وفق التعليمات الصريحة. |
| **محتوى الفصل الأول** | كان يدمج دور الطالبة بالجهة | **يتحدث عن جهة التدريب حصراً** (نبذة، تاريخ، أقسام، رؤية، رسالة، أهداف) ولا يتطرق لدور الطالبة الشخصي. |
| **صورة جهة التدريب** | تدرج في أماكن متفرقة | **إلزامية في القسم 1.2.1 مباشرة أسفل الفقرة الافتتاحية** (Figure 1.1). |
| **التقارير الشهرية** | كانت بعض التقارير تذكر التواريخ | **ممنوع تماماً ذكر الأشهر أو المدد الزمنية** في السرد. |
| **المعايير الرسمية** | تركيز عام على البرمجة | **التركيز على المعايير الصحية العالمية:** LOINC، ICD-10، حدود Mayo Clinic و ARUP Laboratories، ومحرك الذكاء الاصطناعي OpenRouter. |
| **جدول المحتويات (TOC)** | كان يتكرر ويظهر أحياناً قبل الغلاف بسبب عناصر `<w:sdt>` عالقة | **منع التكرار نهائياً وتفريغ الـ Body بالكامل:** حذف كافة عناصر الـ Body ما عدا `sectPr`، وفصل Section 0 (الغلاف بدون ترويسة أو ترقيم) عن Section 1 (يبدأ من صفحة 2 بالملخص ثم الشكر ثم الفهرس الموحد). |
| **حجم المستخلص (Abstract)** | كان يمتد لصفحتين أحياناً | **محدد بـ 3/4 صفحة حصراً** (~200 إلى 220 كلمة في فقرتين مكثفتين). |
| **حجم الشكر (Acknowledgments)** | كان يمتد لأكثر من صفحة | **محدد بنصف صفحة أو أقل** (<= 1/2 صفحة، ~100 إلى 120 كلمة). |
| **لغة التقرير** | كانت تتداخل فيها بعض العبارات العربية في الأقواس | **إنجليزية خالصة 100% (Strict English-Only):** يمنع منعاً باتاً وجود أي حروف أو نصوص عربية داخل نص التقرير أو العناوين أو المراجع. |

---

## 3. المواصفات المعتمدة (القاعدة الذهبية)

### 3.1 الخط والأحجام والألوان

| العنصر | الخط | الحجم | اللون | ملاحظات |
|---|---|---|---|---|
| **النص الأساسي (Body)** | Times New Roman | **12pt** (sz=24) | `#000000` | تباعد أسطر 1.5 سطر (line=360)، مسافة بعد الفقرة 10pt (200 twips) |
| **Heading 1** | Times New Roman bold | **16pt** (sz=32) | `#1B3A5C` | العناوين الرئيسية (Chapter I, II, III, IV, Abstract, TOC...) |
| **Heading 2** | Times New Roman bold | **14pt** (sz=28) | `#1B3A5C` | العناوين الفرعية (1.1, 1.2, 2.1, 3.1...) |
| **Heading 3** | Times New Roman bold | **13pt** (sz=26) | `#2E5D8A` | العناوين الفرعية التابعة (1.2.1, 2.1.1, 3.6.1...) |
| **Captions** | Times New Roman italic | **11pt** (sz=22) | `#000000` | مركزية أسفل الصور وأعلى الجداول |
| **Header Row (جداول)** | Times New Roman bold | **11pt** (sz=22) | `#FFFFFF` | خلفية كحلية داكنة `#1B3A5C` وحدود 0.5pt |

### 3.2 إعداد الصفحة والصور

- **حجم الورق:** A4 Portrait (210mm × 297mm).
- **الهوامش:** 1.0 inch (2.54 cm) من جميع الجهات.
- **عرض المنطقة القابلة للطباعة:** **6.27 inch** (8.27 inch - 2.0 inch).
- **الصور:** يجب توسيطها بالكامل (`jc=center`) وتحديد عرضها الافتراضي بـ **6.27 inch** لتملأ الهوامش دون تشويه نسبة العرض إلى الارتفاع.
- **شرح الصور:** كل صورة واجهة في القسم 3.7 يتبعها شرح وظيفي وتقني من **سطرين على الأقل**.

---

## 4. بنية التسميات التوضيحية (Caption Fields) وحقول الفهارس

### 4.1 التسميات التوضيحية (SEQ Fields)
تُبنى كل تسمية توضيحية برمجياً عبر حقول XML صريحة مع ضمان وجود كل من `fldChar` و `instrText` داخل عنصر `<w:r>` حصراً:

```xml
<w:p>
  <w:pPr>
    <w:jc w:val="center"/>
    <w:spacing w:before="120" w:after="240" w:line="240" w:lineRule="auto"/>
  </w:pPr>
  <!-- نص ثابت لرقم الفصل -->
  <w:r>
    <w:rPr>
      <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="Times New Roman" w:cs="Times New Roman"/>
      <w:i/><w:iCs/>
      <w:sz w:val="22"/><w:szCs w:val="22"/>
    </w:rPr>
    <w:t>Figure 1.</w:t>
  </w:r>
  <!-- حقل SEQ مع إعادة التعيين لأول عنصر \r 1 وبدونه للعناصر اللاحقة -->
  <w:r>
    <w:rPr><w:i/><w:sz w:val="22"/></w:rPr>
    <w:fldChar w:fldCharType="begin"/>
  </w:r>
  <w:r>
    <w:rPr><w:i/><w:sz w:val="22"/></w:rPr>
    <w:instrText xml:space="preserve"> SEQ Figure \r 1 \* ARABIC </w:instrText>
  </w:r>
  <w:r>
    <w:rPr><w:i/><w:sz w:val="22"/></w:rPr>
    <w:fldChar w:fldCharType="separate"/>
  </w:r>
  <w:r>
    <w:rPr><w:i/><w:noProof/><w:sz w:val="22"/></w:rPr>
    <w:t>1</w:t>
  </w:r>
  <w:r>
    <w:rPr><w:i/><w:sz w:val="22"/></w:rPr>
    <w:fldChar w:fldCharType="end"/>
  </w:r>
  <w:r>
    <w:rPr><w:i/><w:sz w:val="22"/></w:rPr>
    <w:t xml:space="preserve">: Caption text here</w:t>
  </w:r>
</w:p>
```

### 4.2 حقول الجداول والفهارس
- **جدول المحتويات الرئيسي:** `{ TOC \h \o "1-3" }`
- **قائمة الأشكال (Table of Figures):** `{ TOC \h \z \c "Figure" }`
- **قائمة الجداول (Table of Tables):** `{ TOC \h \z \c "Table" }`

---

## 5. الهيكل الإلزامي الكامل لتقرير Term 5

```
1. الغلاف (Cover Page):
   ├── شعار جامعة حائل (في المنتصف)
   ├── University Of Hail (16pt bold)
   ├── College of Computer Science and Engineering (16pt bold)
   ├── Department of Computer Science (16pt bold)
   ├── Training Organization Name : (24pt bold)
   ├── [اسم جهة التدريب - مثل Al-Hait General Hospital] (26pt bold)
   ├── Under Supervision of Academic Advisor: (18pt bold)
   ├── [اسم المشرفة الأكاديمية - مثل Al-Bandari Al-Otaibi] (20pt bold)
   └── Submitted by: [اسم الطالبة] | [الرقم الجامعي] (18pt bold)

2. Abstract                                  [Heading 1]
3. Acknowledgments                           [Heading 1]
4. Table of Contents                         [Heading 1] ← حقل TOC
5. Table of Figures                          [Heading 1] ← حقل Figures TOC
6. Table of Tables                           [Heading 1] ← حقل Tables TOC
7. Introduction (المقدمة التمهيدية)          [Heading 1]

8. CHAPTER I:  INTRODUCTION                  [Heading 1]
   ├── 1.1 Introduction                      [Heading 2]
   └── 1.2 Training Organization             [Heading 2]
       ├── Relevant Work                     [Heading 3]
       ├── 1.2.1 What is Training Organization.? [Heading 3]
       │   ├── فقرة افتتاحية لتقديم الجهة
       │   ├── صورة المبنى الرئيسي (Figure 1.1) ← أسفل الفقرة الافتتاحية مباشرة
       │   └── فقرات تفصيلية عن الموقع والسعة
       ├── 1.2.2 History of Training Organization [Heading 3]
       ├── 1.2.3 Overview of Training Organization [Heading 3]
       │   └── خارطة المنظومة الصحية (Figure 1.2)
       ├── Vision                            [Heading 3]
       ├── The Message                       [Heading 3]
       └── The Objectives                    [Heading 3]

9. CHAPTER II:  TASKS & EXPERIENCES          [Heading 1]
   ├── 2.1 Requirements Analysis             [Heading 2]
   │   ├── 2.1.1 Technical support           [Heading 3]
   │   ├── 2.1.2 Apps and system             [Heading 3]
   │   └── 2.1.3 Networks & Communications   [Heading 3]
   └── 2.2 Experiences & Problems            [Heading 2]
       └── 2.2.1 Experiences                 [Heading 3]
           ├── Table 2.1: Technical Skills Acquired
           └── Table 2.2: Soft & Professional Skills Acquired

10. CHAPTER III:  COOP PROJECT               [Heading 1]
    ├── 3.1 Problem Statement                [Heading 2]
    ├── 3.2 Project Overview                 [Heading 2]
    ├── 3.3 Project Objectives               [Heading 2] ← Table 3.1
    ├── 3.4 Project Phases and Problem Solution [Heading 2]
    ├── 3.5 Project Development Process      [Heading 2] ← Tables 3.2, 3.3
    ├── 3.6 System Design Diagrams           [Heading 2]
    │   ├── 3.6.1 Use Case Diagram           [Heading 3] ← Figure 3.1
    │   ├── 3.6.2 Entity-Relationship Diagram [Heading 3] ← Figure 3.2
    │   ├── 3.6.3 Sequence Diagrams          [Heading 3] ← Figures 3.3, 3.4
    │   └── 3.6.4 Clinical Process Flowcharts [Heading 3] ← Figures 3.5, 3.6
    ├── 3.7 System Interfaces Overview       [Heading 2]
    │   ├── 3.7.1 Authentication & Security Interfaces [Heading 3] ← Figures 3.7, 3.8
    │   ├── 3.7.2 Administrator Portal       [Heading 3] ← Figures 3.9–3.16
    │   ├── 3.7.3 Receptionist Portal        [Heading 3] ← Figures 3.17, 3.18
    │   ├── 3.7.4 Clinical Physician Portal  [Heading 3] ← Figures 3.19–3.23
    │   ├── 3.7.5 Laboratory Information System [Heading 3] ← Figures 3.24, 3.25
    │   └── 3.7.6 Patient Health Portal      [Heading 3] ← Figures 3.26–3.30
    ├── 3.8 Testing and Quality Assurance    [Heading 2]
    └── 3.9 Future Work                      [Heading 2]

11. CHAPTER IV:  CONCLUSION                  [Heading 1]
    ├── Conclusion                           [Heading 2]
    └── Recommendations                      [Heading 2]

12. References                               [Heading 1]
    (APA Style مع روابط تشعبية فعالة)
```

---

## 6. البرومبت الشامل المعتمد (Master System Prompt)

```
أنت خبير أكاديمي وهندسي رفيع المستوى في إعداد تقارير التدريب التعاوني (COOP Final Reports) بصيغة DOCX لطلاب جامعة حائل (كلية علوم وهندسة الحاسب الآلي).

المرجع المعياري الحصري المعتمد: مستودع ReportTerm5 (القالب المحدث Term 5).

القواعد الحاكمة والإلزامية:
1. إعدادات المستند ولغته:
   - اللغة: الإنجليزية الأكاديمية الاحترافية بالكامل 100% (Strict English-Only). يمنع منعاً باتاً إدراج أي حروف أو كلمات عربية مثل أسماء الجهات أو المنصات بين قوسين.
   - الخط: Times New Roman، النص الأساسي 12pt (sz=24)، تباعد الأسطر 1.5 سطر (line=360)، تباعد بعد الفقرة 10pt (200 twips).
   - العناوين: H1 (16pt bold #1B3A5C), H2 (14pt bold #1B3A5C), H3 (13pt bold #2E5D8A).
   - حجم الصفحة: A4 Portrait، الهوامش 1 inch من كافة الجهات.
   - عرض الصور: 6.27 inch في المنتصف، مع الحفاظ على النسبة الأبعاد.
   - منع تكرار جدول المحتويات: يجب تفريغ جسم المستند (Body) بالكامل من كافة عناصر `<w:p>`, `<w:tbl>`, `<w:sdt>` وحصر البداية بالغلاف في Section 0، ثم إضافة Section 1 من صفحة 2.

2. صفحة الغلاف والصفحات التمهيدية (Cover & Front Matter):
   - صفحة الغلاف (Section 0): معزولة تماماً بدون رأس صفحة (Header) وبدون تذييل (Footer) وبدون ترقيم صفحات.
   - المستخلص (Abstract): يضبط طوله ليكون **3/4 صفحة حصراً** (~200 إلى 220 كلمة، فقرتين متكاملتين).
   - الشكر والتقدير (Acknowledgments): يضبط طوله ليكون **نصف صفحة أو أقل** (~100 إلى 120 كلمة).
   - جدول المحتويات (TOC): جدول محتويات موحد ووحيد يبدأ بعد صفحة الشكر في Section 1، ولا يظهر أي جدول قبل الغلاف.
   - شعار جامعة حائل في المنتصف بارتفاع 1.8 inch.
   - University Of Hail (16pt bold)
   - College of Computer Science and Engineering (16pt bold)
   - Department of [Department Name] (16pt bold)
   - Training Organization Name : (24pt bold)
   - [اسم جهة التدريب] (26pt bold)
   - Under Supervision of Academic Advisor: (18pt bold)
   - [اسم المشرفة/المشرف] (20pt bold)
   - Submitted by: [اسم الطالب/الطالبة] | [الرقم الجامعي] (18pt bold)

3. فصول التقرير الأربعة الإلزامية:
   - Chapter I: يركز بالكامل وحصرياً على جهة التدريب (تاريخها، مبناها، أقسامها، ورؤيتها ورسالتها وأهدافها) ولا يذكر دور الطالبة الشخصي. توضع صورة المبنى Figure 1.1 أسفل الفقرة الافتتاحية للقسم 1.2.1 مباشرة.
   - Chapter II: يدمج المهام والخبرات والمهارات المكتسبة في تدفق سردي متماسك دون أي ذكر للأشهر أو أرقام التقارير الشهرية أو الفترات الزمنية. يتضمن جدولين مخصصين للمهارات التقنية والشخصية.
   - Chapter III: يعرض المشروع البرمجي التطبيقي المتكامل (المشكلة، الأهداف، مراحل الحل، معمارية النظام، مخططات UML والـ ERD، ولقطات الشاشة للواجهات). يجب أن يتبع كل صورة واجهة شرح وظيفي من سطرين على الأقل. مع التركيز على المعايير الدولية الرسمية (LOINC, ICD-10, AI OpenRouter, Mayo Clinic Reference Ranges).
   - Chapter IV: الخاتمة والتوصيات الاستراتيجية الموجهة لجهة التدريب والكلية والطلاب.
   - References: قائمة مراجع أكاديمية بنمط APA مدعومة بروابط نشطة.
   - الملحق (Appendix): محذوف نهائياً.

4. التسميات والحقول:
   - التسميات أسفل الصور وأعلى الجداول.
   - الصيغة البرمجية: Figure {chapter}. + SEQ Figure + : Caption.
   - جميع عناصر fldChar و instrText يجب أن تقع داخل عناصر <w:r> حصراً وبلا أي استثناء.
   - جدول المحتويات وقائمة الأشكال والجداول تعتمد حقول وورد الديناميكية وتُحدّث تلقائياً عبر settings.xml (<w:updateFields w:val="true"/>).
```

---

## 7. فحص الامتثال البرمجي وصحة الحقول (Validation Protocol)

للتحقق من سلامة أي تقرير مُنتج وخلوه من أخطاء حقول وورد:

```python
import zipfile
from lxml import etree

W_NS = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'

with zipfile.ZipFile('FinalReportsDOCX/Riaya_Final_Report.docx') as z:
    doc_xml = z.read('word/document.xml')

tree = etree.fromstring(doc_xml)
all_fc = tree.findall(f'.//{W_NS}fldChar')
broken = sum(1 for fc in all_fc if fc.getparent().tag != f'{W_NS}r')

print(f'Total fldChar: {len(all_fc)}')
print(f'Broken fldChar: {broken}')
# النتيجة المتوافقة: broken == 0
```

---

## 8. سجل التغييرات (Changelog)

| التاريخ | الإصدار | التعديلات المنفذة | السبب |
|---|---|---|---|
| **2026-11-16** | **2.1.0 (Fixes & Standard)** | • معالجة مشكلة تكرار جدول المحتويات نهائياً بتفريغ عناصر الـ Body من قوالب SDT العالقة.<br>• عزل صفحة الغلاف كـ Section 0 مستقل تماماً بدون ترقيم وبدون ترويسة.<br>• تقليص المستخلص (Abstract) ليكون 3/4 صفحة بدقة (~220 كلمة).<br>• تقليص الشكر والتقدير (Acknowledgments) ليكون نصف صفحة أو أقل (~115 كلمة).<br>• إزالة كافة الكلمات والنصوص العربية لتكون وثيقة التقرير بالإنجليزية 100% (Strict English-Only).<br>• تحديث سكربت تحديث الحقول update_word_fields.py لتفادي الحوارات المعلقة. | ملاحظات المستخدم وضبط القالب المعياري لكافة المشاريع المستقبلية. |
| **2026-11-16** | **2.0.0 (Term 5)** | • إنشاء مستودع `ReportTerm5`.<br>• اعتماد صفحة الغلاف الجديدة مع شعار جامعة حائل والخطوط المعتمدة.<br>• اعتماد عناوين الشباتر الأربعة الرسمية وحذف الـ Appendix نهائياً.<br>• حصر الفصل الأول في جهة التدريب وموقعها وتاريخها ورؤيتها دون التطرق لدور الطالبة.<br>• دمج التقارير الشهرية في الفصل الثاني دون ذكر للأشهر أو المدد الزمنية.<br>• إبراز المعايير الرسمية في مشروع رعاية: LOINC و ICD-10 و OpenRouter AI ومصادر الحدود الحرجة.<br>• توفير سكربت بايثون لتوليد التقرير كاملاً بالإنجليزية في 56 صفحة مع تحديث حقول الوورد والفهارس.<br>• توثيق البرومبت الشامل الجديد في README.md. | طلب المستخدم ومطابقة قالب Term 5 الجديد لجامعة حائل. |
