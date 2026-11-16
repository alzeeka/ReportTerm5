# ReportTerm5 — مرجع قوالب تقارير التدريب التعاوني (الفصل الخامس المعتمد)

هذا المستودع هو **المرجع الموثوق والمعياري** لإنتاج تقارير التدريب التعاوني بصيغة DOCX وفق القالب الجديد المعتمد من جامعة حائل (كلية علوم وهندسة الحاسب الآلي)، مع الحفاظ على أعلى معايير التنسيق والتكامل البرمجي وحقول وورد التلقائية (SEQ & TOC).

---

## 1. بنية المستودع

```
ReportTerm5/
├── README.md                          ← هذا الملف (المرجع المعياري الشامل والبرومبت الذهبي)
├── MASTER_PROMPT.md                   ← البرومبت الشامل الجاهز للنسخ المباشر لنماذج الذكاء الاصطناعي
├── theme_customizer.html              ← 🎨 الأداة البصرية التفاعلية لتخصيص حدود الصفحة وألوان الجداول وتوليد البرومبت
├── template.docx                      ← القالب الأساسي بصيغة DOCX
├── template_term5.docx                ← قالب Term 5 الجديد الجاهز للاستخدام
├── report_builder_core.py             ← المحرك البرمجي لدوال التنسيق، الجداول، والصور، وحقول XML
├── generate_report.py                 ← سكربت التوليد الآلي لتقرير مشروع رعاية (Riaya)
├── make_clean_template.py             ← سكربت تفريغ القالب وتحديث مجلد Tamplate/
├── update_word_fields.py              ← سكربت تحديث حقول Word وتوليد الفهارس وأرقام الصفحات آلياً
├── agile_methodology.png              ← المخطط الدائري المعتمد لمنهجية أجايل (6 مراحل تكرارية)
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
    └── Riaya_Final_Report.docx        ← ✅ التقرير النهائي الشامل لمشروع رعاية (53 صفحة)
```

---

## 2. أبرز الفروقات والتحديثات في قالب Term 5 مقارنة بالقالب القديم

| العنصر | القالب القديم (ReportLMM) | القالب الجديد المعتمد (ReportTerm5) |
|---|---|---|
| **صفحة الغلاف** | الغلاف التقليدي لتقرير Lamia | **غلاف Term 5 الجديد:** شعار جامعة حائل في المنتصف، اسم الجامعة، الكلية، القسم، اسم جهة التدريب ببنط 26pt، المشرفة الأكاديمية ببنط 20pt، بيانات الطالبة والرقم الجامعي. |
| **هيكل الفصول** | فصول التدريب التعاوني القديمة | **فهرس الفصول الأربعة المعتمد:**<br>• Chapter I: INTRODUCTION (الحديث عن جهة التدريب حصراً)<br>• Chapter II: TASKS & EXPERIENCES (دمج المهام دون ذكر الأشهر)<br>• Chapter III: COOP PROJECT (المشروع التقني مع المعايير الدولية ومنهجية أجايل)<br>• Chapter IV: CONCLUSION (الخاتمة والتوصيات)<br>• References (المراجع) |
| **ترقيم الصفحات** | ترقيم عربي مستمر لكافة الصفحات | **معمارية الأقسام الثلاثة (3-Section Architecture):**<br>• الغلاف (Section 0): معزول وبدون ترقيم.<br>• الصفحات التمهيدية (Section 1): أرقام لاتينية صغيرة (`i, ii, iii...`).<br>• فصول التقرير (Section 2): أرقام عربية قياسية (`1, 2, 3...`) تبدأ من 1. |
| **الملحق (Appendix)** | كان موجوداً ويحتوي على قصاصات كود | **تم حذفه بالكامل (Deleted)** وفق التعليمات الصريحة. |
| **محتوى الفصل الأول** | كان يدمج دور الطالبة بالجهة | **يتحدث عن جهة التدريب حصراً** (نبذة، تاريخ، أقسام، رؤية، رسالة، أهداف) ولا يتطرق لدور الطالبة الشخصي. |
| **صورة جهة التدريب** | تدرج في أماكن متفرقة | **إلزامية في القسم 1.2.1 مباشرة أسفل الفقرة الافتتاحية** (Figure 1.1). |
| **التقارير الشهرية** | كانت بعض التقارير تذكر التواريخ | **ممنوع تماماً ذكر الأشهر أو المدد الزمنية** في السرد. |
| **فحص الـ Endpoints** | كان يذكر فحص نقاط API | **تم حذف فحص Endpoints وقسم 2.1.2** لكونها غير متناسبة مع نطاق تدريب الطلاب بالمستشفيات. |
| **منع التحاليل المكررة** | كان محدد بـ 48 ساعة | **معدل إلى منع التحاليل المكررة لمدة 30 يوماً أو حسب إعدادات المنصة** لواقعيته الإكلينيكية. |
| **المعايير الرسمية** | تركيز عام على البرمجة | **التركيز على المعايير الصحية العالمية:** LOINC، ICD-10، حدود المختبرات المعتمدة، محرك الذكاء الاصطناعي OpenRouter، ومنهجية Agile Development. |
| **جدول المحتويات (TOC)** | كان يتكرر ويظهر أحياناً قبل الغلاف بسبب عناصر `<w:sdt>` عالقة | **منع التكرار نهائياً وتفريغ الـ Body بالكامل:** تفريغ الـ Body وعزل Section 0 عن Section 1. |
| **حجم المستخلص (Abstract)** | كان يمتد لصفحتين أحياناً | **محدد بـ 3/4 صفحة حصراً** (~200 إلى 220 كلمة في فقرتين مكثفتين دون عبارات متطلبات التخرج العامة). |
| **حجم الشكر (Acknowledgments)** | كان يمتد لأكثر من صفحة | **محدد بنصف صفحة أو أقل** (<= 1/2 صفحة، ~100 إلى 120 كلمة مع لقب Dr.). |
| **لغة التقرير** | كانت تتداخل فيها بعض العبارات العربية في الأقواس | **إنجليزية خالصة 100% (Strict English-Only):** يمنع وجود أي حروف أو نصوص عربية داخل متن التقرير. |
| **الإشارة للجداول والصور** | أرقام ثابتة في الشرح تسبب تضارباً | **إشارة وصفية عامة في الشرح** (مثل: in the table below, as shown in the diagram) لتفادي التضارب مع حقول SEQ. |

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
    <w:instrText xml:space="preserve"> SEQ Figure \s 1 </w:instrText>
  </w:r>
  <w:r>
    <w:rPr><w:i/><w:sz w:val="22"/></w:rPr>
    <w:fldChar w:fldCharType="separate"/>
  </w:r>
  <w:r>
    <w:rPr><w:i/><w:sz w:val="22"/></w:rPr>
    <w:fldChar w:fldCharType="end"/>
  </w:r>
  <w:r>
    <w:rPr><w:i/><w:sz w:val="22"/></w:rPr>
    <w:t>: Caption Text</w:t>
  </w:r>
</w:p>
```

### 4.2 بنية الأقسام وترقيم الصفحات اللاتيني والعربي (3-Section Setup)

1. **Section 0 (Cover Page):**
   ينتهي بفاصل مقطعي مرتبط بخصائص القسم 0 بدون أرقام صفحات:
   ```python
   p_cover_last._p.get_or_add_pPr().append(parse_xml(s0_xml))
   ```
2. **Section 1 (Preliminary Pages: Abstract -> Introduction):**
   ينتهي بفاصل مقطعي يحمل خاصية الترقيم اللاتيني الصغير:
   ```xml
   <w:pgNumType w:fmt="lowerRoman" w:start="1"/>
   ```
3. **Section 2 (Main Chapters: Chapter I -> References):**
   يحمل جسم المستند الرئيسي خاصية إعادة تعيين الترقيم العربي من 1:
   ```xml
   <w:pgNumType w:fmt="decimal" w:start="1"/>
   ```

---

## 5. هيكل الفصول المعتمد (The Approved 4-Chapter Hierarchy)

```
1. Cover Page                                 [Section 0: بدون ترويسة، بدون تذييل، بدون ترقيم]
2. Abstract                                   [Section 1: يبدأ الترقيم اللاتيني i، 3/4 صفحة]
3. Acknowledgments                            [نصف صفحة أو أقل، Dr. Academic Advisor]
4. Table of Contents                          [حقل وورد تلقائي وحيد: TOC \h \o "1-3"]
5. Table of Figures                           [حقل وورد تلقائي: TOC \h \z \c "Figure"]
6. Table of Tables                            [حقل وورد تلقائي: TOC \h \z \c "Table"]
7. Introduction                               [نهاية Section 1 بفاصل مقطعي]

8. CHAPTER I: INTRODUCTION                    [Section 2: يبدأ الترقيم العربي 1]
   1.1 Introduction
   1.2 Training Organization
       Relevant Work
       1.2.1 What is Training Organization.?  [أسفله مباشرة: Figure 1.1]
       1.2.2 History of Training Organization
       1.2.3 Overview of Training Organization
             Vision
             The Message
             The Objectives

9. CHAPTER II: TASKS & EXPERIENCES
   2.1 Requirements Analysis
       2.1.1 Technical support
       2.1.2 Networks & Communications
   2.2 Experiences & Problems
       2.2.1 Experiences                      [جدول المهارات التقنية + جدول المهارات الشخصية]

10. CHAPTER III: COOP PROJECT
    3.1 Problem Statement
    3.2 Project Overview
    3.3 Project Objectives
    3.4 Project Phases and Problem Solution
        3.4.1 Agile Software Development Methodology [مخطط الأجايل الدائري agile_methodology.png]
        3.4.2 Project Development Phases
    3.5 Project Development Process           [PHP 8.3, MySQL, Bootstrap 5, منع التكرار 30 يوماً]
    3.6 System Design Diagrams                [Use Case, Chen ERD مع 14 جدولاً، Sequence, Activity]
    3.7 System Interfaces Overview            [لقطات الشاشة مع شرح وظيفي لا يقل عن سطرين لكل واجهة]
    3.8 Testing                               [Unit, Integration, OWASP Security Auditing]

11. CHAPTER IV: CONCLUSION & RECOMMENDATIONS
    Conclusion                                [ملخص الإنجازات والدروس المستفادة]
    Recommendations                           [توصيات للجهة والكلية والطلاب]

12. References                                [APA Style مع روابط تشعبية فعالة]
```

---

## 6. فحص الامتثال البرمجي وصحة الحقول (Validation Protocol)

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

## 7. سجل التغييرات (Changelog)

| التاريخ | الإصدار | التعديلات المنفذة | السبب |
|---|---|---|---|
| **2026-11-16** | **2.3.0 (3-Section Roman & Agile Standard)** | • تفعيل معمارية الأقسام الثلاثة: الغلاف (Section 0) غير مرقم، الصفحات التمهيدية (Section 1) بترقيم لاتيني صغير (`i, ii, iii...`)، والفصول الرئيسية (Section 2) بترقيم عربي (`1, 2, 3...`) يبدأ من 1.<br>• حذف عبارة متطلبات التخرج العامة من المستخلص.<br>• تعديل فترة منع إجراء التحاليل المكررة إلى "30 يوماً أو حسب إعدادات المنصة" عبر التقرير بالكامل.<br>• ترقية بيئة العمل المعتمدة إلى PHP 8.3.<br>• حذف قسم 2.1.2 Apps and system واستبعاد فحص Endpoint APIs لكونها غير متناسبة مع نطاق تدريب الطلاب بالمستشفيات، وتعديل ترقيم شبكات الاتصال إلى 2.1.2.<br>• منع كتابة أرقام ثابتة للجداول في الشرح واعتماد الإشارة الوصفية العامة حصراً.<br>• تعميم مقاييس الأداء والأمان دون ذكر أرقام ميلي ثانية ثابتة أو تحديد عدد الثغرات.<br>• توضيح معمارية قاعدة البيانات الفعلية (14 جدولاً معيارياً 3NF) مع إبراز الكيانات الأساسية في مخطط Chen ERD.<br>• إضافة القسم 3.4.1 لمنهجية أجايل (Agile Development) مع إدراج المخطط الدائري ذي المراحل الست (`agile_methodology.png`). | ملاحظات المستخدم وضبط القالب النهائي لكافة المشاريع في المستودع. |
| **2026-11-16** | **2.2.0 (Advisor & Image Rules)** | • إضافة اللقب الأكاديمي Dr. أمام اسم المشرفة الأكاديمية (`Dr. Al-Bandari Al-Otaibi`).<br>• حذف قسم الأعمال المستقبلية (3.9 Future Work) نهائياً من التقرير والقالب.<br>• منع كتابة أرقام الأشكال الثابتة (Figure X.X) في شروحات النصوص واعتماد الإشارة الوصفية العامة لتفادي التضارب عند تحديث حقول SEQ.<br>• تحديث القالب وسكربتات التوليد وإعادة توليد النسخ المعتمدة بنجاح. | ملاحظات المستخدم وضبط القالب للمشاريع المستقبلية. |
| **2026-11-16** | **2.1.0 (Fixes & Standard)** | • معالجة مشكلة تكرار جدول المحتويات نهائياً بتفريغ عناصر الـ Body من قوالب SDT العالقة.<br>• عزل صفحة الغلاف كـ Section 0 مستقل تماماً بدون ترقيم وبدون ترويسة.<br>• تقليص المستخلص (Abstract) ليكون 3/4 صفحة بدقة (~220 كلمة).<br>• تقليص الشكر والتقدير (Acknowledgments) ليكون نصف صفحة أو أقل (~115 كلمة).<br>• إزالة كافة الكلمات والنصوص العربية لتكون وثيقة التقرير بالإنجليزية 100% (Strict English-Only).<br>• تحديث سكربت تحديث الحقول update_word_fields.py لتفادي الحوارات المعلقة. | ملاحظات المستخدم وضبط القالب المعياري لكافة المشاريع المستقبلية. |
| **2026-11-16** | **2.0.0 (Term 5)** | • إنشاء مستودع `ReportTerm5`.<br>• اعتماد صفحة الغلاف الجديدة مع شعار جامعة حائل والخطوط المعتمدة.<br>• اعتماد عناوين الشباتر الأربعة الرسمية وحذف الـ Appendix نهائياً.<br>• حصر الفصل الأول في جهة التدريب وموقعها وتاريخها ورؤيتها دون التطرق لدور الطالبة.<br>• دمج التقارير الشهرية في الفصل الثاني دون ذكر للأشهر أو المدد الزمنية.<br>• إبراز المعايير الرسمية في مشروع رعاية: LOINC و ICD-10 و OpenRouter AI ومصادر الحدود الحرجة.<br>• توفير سكربت بايثون لتوليد التقرير كاملاً بالإنجليزية مع تحديث حقول الوورد والفهارس.<br>• توثيق البرومبت الشامل الجديد في README.md. | طلب المستخدم ومطابقة قالب Term 5 الجديد لجامعة حائل. |
