<div align="center">

<table width="100%" cellpadding="0" cellspacing="0">
  <tr>
    <td align="center" width="25%">
      <img src="Karkh1_Vocational_Edu_Dept_Logo.png"
           style="max-width:120px; width:100%; height:auto;" />
    </td>
    <td align="center" width="35%">
      <h2>نظام إدارة الحضور والانصراف</h2>
      <p>مشروع تسجيل حضور وانصراف الطلاب</p>
    </td>
    <td align="center" width="40%">
      <img src="Zaytoon_Vocational_Logo.png"
           style="max-width:320px; width:100%; height:auto;" />
    </td>
  </tr>
</table>

</div>

<hr />


# ⟡ الأعدادية ┊ اعدادية الزيتون المهنية للبنين 
# ⟡ dev ┊ BERLIN  

⟡⟡⟡──────────────────────────────────────────────⟡⟡⟡

# العربية
⟡ ينظم حضور وانصراف الطلبة داخل المدرسة بصورة واضحة وسريعة.  
⟡ كل العمل يعتمد على اختيار المرحلة ثم اختيار القسم من القوائم.  
⟡ بعد الاختيار تظهر أسماء الطلبة الخاصة بهذا الصف فقط.  
⟡ البحث بالاسم يحدد الطالب خلال ثوانٍ وبأقل مجهود.  

⟡ التسجيل اليدوي يتم عبر اختيار الطالب ثم تنفيذ التسجيل بضغطة واحدة.  
⟡ داخل اليوم نفسه يتبدل التسجيل تلقائيًا بين حضور وانصراف.  
⟡ يمكن التبديل بين عرض الطلبة وعرض السجلات من نفس الواجهة.  
⟡ السجلات تظهر في جدول منظم مع التاريخ والوقت لكل عملية.  

⟡ إضافة طالب جديدة متاحة في أي وقت مع تحديد المرحلة والقسم.  
⟡ نقل طالب أو مجموعة طلاب بين الأقسام والمراحل متاح من الواجهة.  
⟡ عند النقل يتم تحديث بيانات الطالب مع الحفاظ على ربط البطاقة.  

⟡ التسجيل عبر RFID يعمل عند تشغيل القارئ وربطه بـ Arduino.  
⟡ البطاقة المعروفة تسجل فورًا دون سؤال أو خطوات إضافية.  
⟡ البطاقة الجديدة تتطلب ربطًا باسم موجود أو إدخال اسم جديد.  
⟡ بعد الربط يتم الحفظ ثم التسجيل مباشرة ضمن نفس السياق.  

⟡ التصدير اليومي ينشئ Excel منسق لكل مرحلة ولكل قسم.  
⟡ الملف يوضح الحضور والانصراف، ويضع “غياب” عند عدم تسجيل حضور.  
⟡ حساب الغياب يقرأ ملفات Excel في records ويحسب أيام “غياب”.  
⟡ إن كانت ملفات الطلاب فارغة، تُنشأ أسماء عربية وهمية تلقائيًا.  

### ◈ تفاصيل منظّمة
▸ السياق المعتمد  
⟶ مرحلة + قسم، ثم كل العمليات تُطبق ضمن هذا الصف فقط.  

▸ مفاتيح الحفظ  
⟶ مفتاح الطالب: `الاسم|المرحلة|القسم`  
⟶ مفتاح اليوم: `YYYY-MM-DD`  

▸ مسارات التخزين  
⟶ `data/cards.json`  ┊ ربط UID مع بيانات الطالب.  
⟶ `data/students/<stage>_<dept>.json`  ┊ طلاب كل صف.  
⟶ `data/attendance/<stage>_<dept>.json`  ┊ سجلات كل صف.  
⟶ `records/<stage>/<dept>/<YYYY-MM-DD>.xlsx`  ┊ تقارير يومية.  

▸ القارئ والبطاقات  
⟶ يعمل عبر Serial بسرعة 9600 مع كشف منفذ تلقائي أو يدوي.  
⟶ UID يُوحّد قبل المعالجة لمنع اختلاف التنسيق.  

▸ التصدير والغياب  
⟶ بعد التصدير يتم تصفير سجلات اليوم داخل data/attendance.  
⟶ الغياب يُحسب من records فقط لضمان ثبات النتائج.  

▸ الشعارات  
⟶ ضع الصور بجانب `main.py` مباشرة.  
⟶ الأسماء المطلوبة:  
  ⟿ `Zaytoon_Vocational_Logo.png`  
  ⟿ `Karkh1_Vocational_Edu_Dept_Logo.png`  

▸ التشغيل  
⟶ تثبيت: `pip install -r requirements.txt`  
⟶ تشغيل: `python main.py`  

</div>

⟡⟡⟡──────────────────────────────────────────────⟡⟡⟡

## English 

### ◈ Comprehensive feature coverage
⟡ The application organises attendance and departure with a clear class scope.  
⟡ All actions are bound to a stage–department context, selected up front.  
⟡ The student list remains strictly limited to the chosen class context.  
⟡ Name search enables immediate access with minimal operational friction.  

⟡ Manual logging is executed via selection and a single action.  
⟡ Within the same day, the system automatically alternates event types.  
⟡ The interface switches between a student view and a records view.  
⟡ Records are presented as a structured table with dates and timestamps.  

⟡ Students can be added at any time with explicit class assignment.  
⟡ Individuals or groups can be transferred across classes seamlessly.  
⟡ RFID bindings are preserved and updated during transfers.  

⟡ RFID logging is supported via Arduino once the reader is enabled.  
⟡ Known cards log instantly without additional prompts or interruptions.  
⟡ Unknown cards trigger a guided binding step to an existing or new name.  
⟡ After binding, the event is recorded immediately in the same workflow.  

⟡ Daily Excel reports are produced per stage and department, fully formatted.  
⟡ Absence counting is derived from the Excel archive stored in records.  
⟡ If student files are empty, Arabic dummy names are generated automatically.  

### ◈ Structured notes
▸ Identifiers  
⟶ student_key: `name|stage|department`  
⟶ date_key: `YYYY-MM-DD`  

▸ Paths  
⟶ `data/cards.json`  
⟶ `data/students/<stage>_<dept>.json`  
⟶ `data/attendance/<stage>_<dept>.json`  
⟶ `records/<stage>/<dept>/<YYYY-MM-DD>.xlsx`  

▸ Run  
⟶ Install: `pip install -r requirements.txt`  
⟶ Start: `python main.py`  

⟡⟡⟡──────────────────────────────────────────────⟡⟡⟡

## Deutsch

### ◈ Vollständige Funktionsabdeckung
⟡ Die Anwendung strukturiert Anwesenheit und Abgang in einem Klassen-Kontext.  
⟡ Sämtliche Aktionen sind an Stufe und Abteilung gebunden, vorab ausgewählt.  
⟡ Die Schülerliste bleibt konsequent auf den gewählten Kontext begrenzt.  
⟡ Eine Namenssuche ermöglicht schnelle Auswahl und sofortige Erfassung.  

⟡ Manuelle Einträge erfolgen per Auswahl und einem einzigen Klick.  
⟡ Am selben Tag schaltet das System automatisch zwischen Eintragsarten um.  
⟡ Die Oberfläche wechselt zwischen Schüleransicht und Datensatzansicht.  
⟡ Datensätze erscheinen tabellarisch mit Datum und Zeitstempeln.  

⟡ Schüler lassen sich jederzeit hinzufügen oder gruppiert verschieben.  
⟡ RFID-Zuordnungen bleiben erhalten und werden beim Verschieben angepasst.  
⟡ RFID-Erfassung über Arduino protokolliert bekannte Karten sofort.  
⟡ Unbekannte Karten werden geführt zugeordnet und danach direkt erfasst.  

⟡ Excel-Berichte entstehen täglich pro Stufe und Abteilung, sauber formatiert.  
⟡ Fehlzeiten werden aus dem Excel-Archiv im Ordner records abgeleitet.  
⟡ Bei leeren Dateien werden arabische Dummy-Namen automatisch erzeugt.  

### ◈ Strukturhinweise
▸ Schlüssel  
⟶ student_key: `name|stage|department`  
⟶ date_key: `YYYY-MM-DD`  

▸ Pfade  
⟶ `data/cards.json`  
⟶ `data/students/<stage>_<dept>.json`  
⟶ `data/attendance/<stage>_<dept>.json`  
⟶ `records/<stage>/<dept>/<YYYY-MM-DD>.xlsx`  

▸ Start  
⟶ Installieren: `pip install -r requirements.txt`  
⟶ Starten: `python main.py`  


















