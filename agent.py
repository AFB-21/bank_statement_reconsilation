import os
import sys
import re
import json
import time
from typing import List
from difflib import get_close_matches

# ==========================================
# 1. تنظيف متغيرات البيئة لمنع خطأ 401 UNAUTHENTICATED
# ==========================================
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

for env_var in [
    "GOOGLE_APPLICATION_CREDENTIALS",
    "GCLOUD_PROJECT",
    "GOOGLE_CLOUD_PROJECT",
]:
    os.environ.pop(env_var, None)

import pandas as pd
from pydantic import BaseModel, Field
from google import genai
from google.genai import types, errors
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ==========================================
# 2. إعدادات النظام والمفاتيح (Configuration)
# ==========================================
API_KEY = (os.environ.get("GEMINI_API_KEY") or os.environ.get("GOOGLE_API_KEY") or "").strip()
if not API_KEY:
    raise ValueError("GEMINI_API_KEY is not set. Please set it in your .env file or environment variables.")
MODEL_NAME = "gemini-3.5-flash-lite"      # الموديل الأساسي (خفيف وسريع وحصته واسعة)
FALLBACK_MODELS = ["gemini-3.5-flash", "gemini-3.5-flash-lite"]  # موديلات بديلة عند نفاد الحصة
EXCEL_GUIDE_PATH = "الوحدات الحسابية و اكوادها المؤسسية.xlsx"
CACHE_FILE_PATH = "unit_aliases_cache.json"
OUTPUT_EXCEL_PATH = "التقرير_النهائي_للوحدات_الحسابية.xlsx"

# تهيئة العميل مع إجبار المكتبة على عدم استخدام Vertex AI أو OAuth2
client = genai.Client(api_key=API_KEY.strip(), vertexai=False)

# ==========================================
# 3. إعداد الطباعة بالعربية في الـ Terminal (اختياري وآمن)
# ==========================================
try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

try:
    import arabic_reshaper
    from bidi.algorithm import get_display

    def cprint(text: str):
        """طباعة النص العربي بشكل سليم وغير معكوس في الـ Terminal"""
        reshaped = arabic_reshaper.reshape(str(text))
        print(get_display(reshaped))
except ImportError:
    def cprint(text: str):
        """في حال عدم تثبيت مكتبات العربي يطبع بشكل عادي دون توقف"""
        print(text)


# ==========================================
# 4. هياكل البيانات (Pydantic Schemas) لضمان دقة مخرجات الـ AI
# ==========================================
class TransactionItem(BaseModel):
    page: int = Field(description="رقم الصفحة في كشف الحساب")
    date: str = Field(description="تاريخ القيد بصيغة YYYY/MM/DD")
    gp: str = Field(description="رقم المدفوعة الذي يبدأ بـ GP")
    credit: float = Field(description="المبلغ الدائن (المحول) كرقم عشري")
    sender_pdf: str = Field(
        description="اسم الوحدة الحسابية المرسلة المكتوب بعد كلمة (الراسل) بدون أقواس"
    )


class PageExtractionResult(BaseModel):
    transactions: List[TransactionItem]


class AIMatchResult(BaseModel):
    matched_name: str = Field(
        description="اسم الوحدة الحسابية المطابق من القائمة المعطاة"
    )
    moasase_no: str = Field(description="الكود المؤسسي الفرعي")
    moasase_no_main: str = Field(description="الكود المؤسسي الرئيسي")
    note: str = Field(description="توضيح مختصر لسبب اختيار هذه الوحدة")


# ==========================================
# 5. دالة الاتصال الذكي مع إعادة المحاولة والتبديل التلقائي (503 / 429)
# ==========================================
def generate_with_retry(contents, config, model_name=MODEL_NAME, max_retries=5):
    """الاتصال الذكي مع دعم إعادة المحاولة والتبديل التلقائي للموديلات البديلة عند استنفاد الحصة أو ضغط السيرفر"""
    models_to_try = [model_name] + [m for m in FALLBACK_MODELS if m != model_name]

    last_error = None
    for current_model in models_to_try:
        for attempt in range(max_retries):
            try:
                return client.models.generate_content(
                    model=current_model,
                    contents=contents,
                    config=config,
                )
            except errors.ServerError as e:
                # خطأ 503 / 500 (ضغط مؤقت على سيرفرات جوجل)
                wait_time = (attempt + 1) * 4
                cprint(
                    f"⏳ ضغط مؤقت على ({current_model})، انتظار {wait_time} ثانية ({attempt + 1}/{max_retries})..."
                )
                time.sleep(wait_time)
                last_error = e
            except errors.ClientError as e:
                err_str = str(e)
                if "429" in err_str:
                    # فحص ما إذا كان الخطأ استنفاد للحصة اليومية للموديل
                    if "RESOURCE_EXHAUSTED" in err_str or "quota" in err_str.lower() or "free_tier" in err_str:
                        cprint(f"⚠️ نفدت حصة الموديل ({current_model}). جاري التبديل لموديل بديل...")
                        last_error = e
                        break  # الانتقال فوراً للموديل التالي دون إضاعة الوقت في الانتظار
                    else:
                        # تجاوز مؤقت للطلبات في الدقيقة (RPM)
                        wait_time = 15
                        cprint(f"⏳ تجاوز معدل الطلبات في الدقيقة، انتظار {wait_time} ثانية...")
                        time.sleep(wait_time)
                        last_error = e
                else:
                    raise e

    raise Exception(
        f"❌ فشل الاتصال بالسيرفر بعد تجربة الموديلات المتاحة. تفاصيل الخطأ: {last_error}"
    )


# ==========================================
# 6. دوال تنظيف النصوص والذاكرة المؤقتة (Helpers & Cache)
# ==========================================
def normalize_arabic(text: str) -> str:
    # """توحيد الحروف العربية وإزالة الكلمات الزائدة لضمان دقة المطابقة"""
    """Normalizing arabic and removing extra words """
    if not isinstance(text, str):
        return ""
    text = text.strip()
    text = re.sub(r"[أإآ]", "ا", text)
    text = re.sub(r"ة", "ه", text)
    text = re.sub(r"ى", "ي", text)
    text = re.sub(r"^(و\s*\.\s*ح|وح|و\.ح|الوحدة الحسابية|لـ)\s+", "", text)
    text = re.sub(r"\b(و\s*\.\s*ح|وح|و\.ح)\b", "", text)
    text = re.sub(r"^ال", "", text)
    text = re.sub(r"[\.\-\(\)\/،,_]", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def load_cache() -> dict:
    if os.path.exists(CACHE_FILE_PATH):
        with open(CACHE_FILE_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}


def save_cache(cache_data: dict):
    with open(CACHE_FILE_PATH, "w", encoding="utf-8") as f:
        json.dump(cache_data, f, ensure_ascii=False, indent=2)


# ==========================================
# 7. استخراج العمليات من ملف الـ PDF مباشرة (Inline Bytes)
# ==========================================
def extract_transactions_from_pdf(pdf_path: str) -> List[dict]:
    # cprint(f"📄 جاري قراءة وتحليل ملف الـ PDF مباشرة: {pdf_path} ...")
    print(f"📄 Reading and analyzing pdf file {pdf_path} ...")

    # قراءة ملف الـ PDF كـ Bytes مباشرة من الجهاز لتفادي خطأ 401 في FileService
    with open(pdf_path, "rb") as f:
        pdf_bytes = f.read()

    pdf_part = types.Part.from_bytes(data=pdf_bytes, mime_type="application/pdf")

    prompt = """
    أنت محلل مالي دقيق لكشوف حسابات منظومة الدفع الإلكتروني الحكومية (e-finance).
    قم باستخراج جميع العمليات المالية الدائنة (التي فيها "المبلغ الدائن" أكبر من 0.00) والتي تتبع "الباب الموازني : الدائنون" فقط.
    
    قواعد صارمة:
    1. تجاهل أي حركة مدينة (المبلغ المدين أكبر من 0 مثل حركات المناقلة بالخصم).
    2. تجاهل تماماً أي جدول تحت عنوان "الباب الموازني : دائنون تمويل ذاتي" في نهاية الكشف.
    3. استخرج رقم الصفحة المطبوع في رأس كل صفحة بدقة.
    4. استخرج اسم الجهة الراسلة المكتوب بعد كلمة (الراسل ...) داخل عمود تفاصيل العمليات.
    5. تأكد من عدم إسقاط أي حركة دائنة في أي صفحة.
    """

    response = generate_with_retry(
        contents=[pdf_part, prompt],
        config=types.GenerateContentConfig(
            response_mime_type="application/json",
            response_schema=PageExtractionResult,
            temperature=0.0,
        ),
    )

    result_data = json.loads(response.text)
    transactions = result_data.get("transactions", [])
    # cprint(f"✅ تم استخراج {len(transactions)} عملية تحويل من {pdf_path}")
    print(f"✅ Extracted {len(transactions)} transaction from {pdf_path}")

    return transactions


# ==========================================
# 8. محرك المطابقة الذكي (Hybrid Matching Engine)
# ==========================================
class AccountingUnitMatcher:

    def __init__(self, excel_path: str):
        # cprint("📊 جاري تحميل دليل الوحدات الحسابية من الإكسيل...")
        print("📊 uploading the codes for wahadat...")

        self.df_guide = pd.read_excel(excel_path)
        self.df_guide["wehda_hesabia"] = (
            self.df_guide["wehda_hesabia"].astype(str).str.strip()
        )
        self.df_guide["norm_name"] = self.df_guide["wehda_hesabia"].apply(
            normalize_arabic
        )
        self.cache = load_cache()

    def ask_ai_agent(
        self, pdf_sender_name: str, candidates_df: pd.DataFrame
    ) -> dict:
        # cprint(f"🤖 استدعاء الـ AI Agent لمطابقة اسم غير مباشر: '{pdf_sender_name}'")
        print(f"🤖 calling AI agent to reconsile a complex wehda name: '{pdf_sender_name}'")

        candidates_list = candidates_df[
            ["wehda_hesabia", "moasase_no", "moasase_no_main"]
        ].to_dict(orient="records")

        prompt = f"""
        أنت خبير في دليل الوحدات الحسابية الحكومية المصرية.
        ورد في كشف الحساب البنكي تحويل من وحدة مكتوبة باسم: "{pdf_sender_name}"
        
        المطلوب اختيار الوحدة الحسابية الصحيحة المطابقة لها من القائمة التالية المستخرجة من دليل الإكسيل:
        {json.dumps(candidates_list, ensure_ascii=False, indent=2)}
        
        معلومات مساعدة للمطابقة الإدارية:
        - "جامعة العاصمة" في مستشفى بدر هي المسمى المرتبط بـ "مستشفى بدر العام" التابع لجامعة حلوان (30900201).
        - "مستشفى كفر شكر التخصصي" مسجل بالدليل باسم "و.ح مستشفي كفر شكر المركزي بمحافظة القليوبية" (13107201).
        - "مستشفى التضامن ببورسعيد" أصبحت تتبع "و.ح مجمع الشفاء الطبى" (32110504) بهيئة الرعاية الصحية ببورسعيد.
        - إذا كان الاسم لصناديق خاصة بمستشفيات جامعية ولا يوجد حساب صناديق منفصل للمستشفى، اختر الوحدة الحسابية لمستشفيات تلك الجامعة.
        """

        response = generate_with_retry(
            contents=prompt,
            config=types.GenerateContentConfig(
                response_mime_type="application/json",
                response_schema=AIMatchResult,
                temperature=0.0,
            ),
        )
        return json.loads(response.text)

    def match(self, pdf_sender_name: str) -> dict:
        # 1. فحص الذاكرة المؤقتة أولاً (Cache)
        if pdf_sender_name in self.cache:
            return self.cache[pdf_sender_name]

        norm_s = normalize_arabic(pdf_sender_name)

        # 2. المرحلة الأولى: تطابق نصي مباشر بعد التنظيف (Exact Match)
        exact = self.df_guide[self.df_guide["norm_name"] == norm_s]
        if len(exact) == 1:
            r = exact.iloc[0]
            res = {
                "excel_name": r["wehda_hesabia"],
                "moasase_no": str(r["moasase_no"]),
                "moasase_no_main": str(r["moasase_no_main"]),
                "note": "تطابق تام",
            }
            self.cache[pdf_sender_name] = res
            save_cache(self.cache)
            return res

        # 3. المرحلة الثانية: احتواء نصي وحيد (Single Substring Match)
        contains = self.df_guide[
            self.df_guide["norm_name"].apply(
                lambda x: (norm_s in x or x in norm_s)
                if (x and len(x) > 5)
                else False
            )
        ]
        if len(contains) == 1:
            r = contains.iloc[0]
            res = {
                "excel_name": r["wehda_hesabia"],
                "moasase_no": str(r["moasase_no"]),
                "moasase_no_main": str(r["moasase_no_main"]),
                "note": "تطابق تام",
            }
            self.cache[pdf_sender_name] = res
            save_cache(self.cache)
            return res

        # 4. المرحلة الثالثة: تجميع أقرب المرشحين وإرسالهم للـ AI Agent
        stop_words = {
            "مستشفي",
            "اداره",
            "مديريه",
            "منطقه",
            "الشئون",
            "الصحيه",
            "العامه",
            "هيئه",
            "التامين",
            "الصحي",
            "جامعه",
            "كليه",
            "مركز",
            "الطبيه",
            "المركزي",
            "العام",
        }
        words = [
            w for w in norm_s.split() if len(w) > 3 and w not in stop_words
        ]
        candidates = pd.DataFrame()
        if words:
            pattern = "|".join([re.escape(w) for w in words])
            candidates = self.df_guide[
                self.df_guide["norm_name"].str.contains(pattern, na=False)
            ].head(25)

        if len(candidates) < 5:
            close_names = get_close_matches(
                norm_s, self.df_guide["norm_name"].tolist(), n=15, cutoff=0.35
            )
            fuzzy_candidates = self.df_guide[
                self.df_guide["norm_name"].isin(close_names)
            ]
            candidates = (
                pd.concat([candidates, fuzzy_candidates])
                .drop_duplicates(subset=["moasase_no"])
                .head(30)
            )

        ai_res = self.ask_ai_agent(pdf_sender_name, candidates)
        res = {
            "excel_name": ai_res["matched_name"],
            "moasase_no": str(ai_res["moasase_no"]),
            "moasase_no_main": str(ai_res["moasase_no_main"]),
            "note": ai_res["note"],
        }
        self.cache[pdf_sender_name] = res
        save_cache(self.cache)
        return res


# ==========================================
# 9. تصدير وتنسيق ملف الإكسيل الاحترافي
# ==========================================
def export_to_excel(df_all: pd.DataFrame, output_path: str):
    # cprint("📝 جاري إنشاء وتنسيق ملف الإكسيل النهائي...")
    print("📝 building and formatting the final excel file...")


    grouped_list = []
    for (excel_name, moasase_no), group in df_all.groupby(
        ["excel_name", "moasase_no"], sort=False
    ):
        pages = sorted(group["page"].unique().tolist())
        first_row = group.iloc[0]
        grouped_list.append({
            "sender_pdf": first_row["sender_pdf"],
            "excel_name": excel_name,
            "moasase_no": str(moasase_no),
            "moasase_no_main": str(first_row["moasase_no_main"]),
            "count": len(group),
            "total_credit": group["credit"].sum(),
            "pages_str": ", ".join(map(str, pages)),
            "note": first_row["note"],
        })
    df_grouped = pd.DataFrame(grouped_list)

    wb = openpyxl.Workbook()
    ws1 = wb.active
    ws1.title = "ملخص الوحدات الحسابية"
    ws1.views.sheetView[0].rightToLeft = True

    ws2 = wb.create_sheet(title="تفاصيل كافة التحويلات")
    ws2.views.sheetView[0].rightToLeft = True

    header_fill = PatternFill(
        start_color="1F4E78", end_color="1F4E78", fill_type="solid"
    )
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    zebra_fill = PatternFill(
        start_color="F2F6F9", end_color="F2F6F9", fill_type="solid"
    )
    white_fill = PatternFill(
        start_color="FFFFFF", end_color="FFFFFF", fill_type="solid"
    )
    warn_fill = PatternFill(
        start_color="FFF2CC", end_color="FFF2CC", fill_type="solid"
    )
    total_fill = PatternFill(
        start_color="D9E1F2", end_color="D9E1F2", fill_type="solid"
    )

    regular_font = Font(name="Calibri", size=11, bold=False)
    bold_font = Font(name="Calibri", size=11, bold=True)
    code_font = Font(name="Consolas", size=11, bold=True, color="1F4E78")

    thin_side = Side(style="thin", color="D9D9D9")
    border_cell = Border(
        left=thin_side, right=thin_side, top=thin_side, bottom=thin_side
    )
    border_total = Border(
        top=Side(style="thin", color="1F4E78"),
        bottom=Side(style="double", color="1F4E78"),
        left=thin_side,
        right=thin_side,
    )

    # --- الورقة الأولى (الملخص) ---
    headers1 = [
        "م",
        "اسم الوحدة المرسلة (بالـ PDF)",
        "اسم الوحدة المطابق (بالإكسيل)",
        "الكود المؤسسي الفرعي",
        "الكود المؤسسي الرئيسي",
        "عدد التحويلات",
        "إجمالي المبالغ المحولة (ج.م)",
        "أرقام الصفحات",
        "ملاحظات المطابقة",
    ]
    ws1.append(headers1)
    for col in range(1, len(headers1) + 1):
        c = ws1.cell(row=1, column=col)
        c.fill, c.font = header_fill, header_font
        c.alignment = Alignment(
            horizontal="center", vertical="center", wrap_text=True
        )
    ws1.row_dimensions[1].height = 30

    for idx, r in df_grouped.iterrows():
        r_num = idx + 2
        ws1.append([
            idx + 1,
            r["sender_pdf"],
            r["excel_name"],
            r["moasase_no"],
            r["moasase_no_main"],
            r["count"],
            r["total_credit"],
            r["pages_str"],
            r["note"],
        ])
        fill = (
            warn_fill
            if r["note"] != "تطابق تام"
            else (zebra_fill if idx % 2 == 1 else white_fill)
        )
        for col in range(1, len(headers1) + 1):
            c = ws1.cell(row=r_num, column=col)
            c.fill, c.border, c.font = fill, border_cell, regular_font
            if col in [1, 6, 8]:
                c.alignment = Alignment(horizontal="center", vertical="center")
            elif col in [4, 5]:
                c.alignment, c.font = (
                    Alignment(horizontal="center", vertical="center"),
                    code_font,
                )
            elif col == 7:
                c.number_format, c.alignment, c.font = (
                    "#,##0.00",
                    Alignment(horizontal="right", vertical="center"),
                    bold_font,
                )
            else:
                c.alignment = Alignment(
                    horizontal="right", vertical="center", wrap_text=True
                )
        ws1.row_dimensions[r_num].height = 22

    tot_r1 = len(df_grouped) + 2
    ws1.append([
        "الإجمالي",
        f"{len(df_grouped)} وحدة حسابية",
        "",
        "",
        "",
        f"=SUM(F2:F{tot_r1-1})",
        f"=SUM(G2:G{tot_r1-1})",
        "",
        "",
    ])
    for col in range(1, len(headers1) + 1):
        c = ws1.cell(row=tot_r1, column=col)
        c.fill, c.font, c.border = total_fill, bold_font, border_total
        c.alignment = Alignment(
            horizontal="right" if col == 7 else "center", vertical="center"
        )
        if col == 7:
            c.number_format = "#,##0.00"

    for i, w in enumerate([6, 45, 48, 22, 22, 14, 22, 18, 45], 1):
        ws1.column_dimensions[get_column_letter(i)].width = w
    ws1.freeze_panes = "A2"
    ws1.auto_filter.ref = f"A1:I{tot_r1-1}"

    # --- الورقة الثانية (التفاصيل) ---
    headers2 = [
        "م",
        "رقم الصفحة",
        "تاريخ القيد",
        "رقم المدفوعة (GP)",
        "المبلغ المحول (ج.م)",
        "اسم الوحدة المرسلة (بالـ PDF)",
        "اسم الوحدة المطابق (بالإكسيل)",
        "الكود المؤسسي الفرعي",
        "الكود المؤسسي الرئيسي",
        "ملاحظات المطابقة",
    ]
    ws2.append(headers2)
    for col in range(1, len(headers2) + 1):
        c = ws2.cell(row=1, column=col)
        c.fill, c.font = header_fill, header_font
        c.alignment = Alignment(
            horizontal="center", vertical="center", wrap_text=True
        )
    ws2.row_dimensions[1].height = 30

    for idx, r in df_all.iterrows():
        r_num = idx + 2
        ws2.append([
            idx + 1,
            r["page"],
            r["date"],
            r["gp"],
            r["credit"],
            r["sender_pdf"],
            r["excel_name"],
            r["moasase_no"],
            r["moasase_no_main"],
            r["note"],
        ])
        fill = (
            warn_fill
            if r["note"] != "تطابق تام"
            else (zebra_fill if idx % 2 == 1 else white_fill)
        )
        for col in range(1, len(headers2) + 1):
            c = ws2.cell(row=r_num, column=col)
            c.fill, c.border, c.font = fill, border_cell, regular_font
            if col in [1, 2, 3, 4]:
                c.alignment = Alignment(horizontal="center", vertical="center")
            elif col in [8, 9]:
                c.alignment, c.font = (
                    Alignment(horizontal="center", vertical="center"),
                    code_font,
                )
            elif col == 5:
                c.number_format, c.alignment, c.font = (
                    "#,##0.00",
                    Alignment(horizontal="right", vertical="center"),
                    bold_font,
                )
            else:
                c.alignment = Alignment(
                    horizontal="right", vertical="center", wrap_text=True
                )
        ws2.row_dimensions[r_num].height = 22

    tot_r2 = len(df_all) + 2
    ws2.append([
        "الإجمالي",
        "",
        "",
        f"{len(df_all)} عملية",
        f"=SUM(E2:E{tot_r2-1})",
        "",
        "",
        "",
        "",
        "",
    ])
    for col in range(1, len(headers2) + 1):
        c = ws2.cell(row=tot_r2, column=col)
        c.fill, c.font, c.border = total_fill, bold_font, border_total
        c.alignment = Alignment(
            horizontal="right" if col == 5 else "center", vertical="center"
        )
        if col == 5:
            c.number_format = "#,##0.00"

    for i, w in enumerate([6, 12, 14, 22, 22, 45, 48, 22, 22, 45], 1):
        ws2.column_dimensions[get_column_letter(i)].width = w
    ws2.freeze_panes = "A2"
    ws2.auto_filter.ref = f"A1:J{tot_r2-1}"

    wb.save(output_path)
    # cprint(f"🎉 تم حفظ التقرير النهائي بنجاح في: {output_path}")
    print(f"🎉 Report saved successfully at: {output_path}")



# ==========================================
# 10. نقطة التشغيل الرئيسية (Main Execution)
# ==========================================
if __name__ == "__main__":
    pdf_files = ["1.pdf", "2.pdf", "3.pdf"]

    matcher = AccountingUnitMatcher(EXCEL_GUIDE_PATH)
    all_transactions = []

    for pdf_file in pdf_files:
        if os.path.exists(pdf_file):
            # حفظ نسخة كاش من استخراج كل ملف PDF حتى لا نعيد قراءته لو حدث انقطاع
            pdf_cache_file = f"{pdf_file}_extracted.json"
            if os.path.exists(pdf_cache_file):
                # cprint(
                #     f"📂 تم العثور على بيانات مستخرجة مسبقاً لـ {pdf_file}، جاري قراءتها فوراً..."
                # )
                print(
                                    f"📂 found previously extracted data for {pdf_file}، reading now..."
                                )
                with open(pdf_cache_file, "r", encoding="utf-8") as f:
                    trans_list = json.load(f)
            else:
                trans_list = extract_transactions_from_pdf(pdf_file)
                with open(pdf_cache_file, "w", encoding="utf-8") as f:
                    json.dump(trans_list, f, ensure_ascii=False, indent=2)

            for t in trans_list:
                match_info = matcher.match(t["sender_pdf"])
                t.update(match_info)
                all_transactions.append(t)
        else:
            # cprint(f"⚠️ الملف {pdf_file} غير موجود في المجلد الحالي.")
            print(f"⚠️ File {pdf_file} not found in the current directory.")


    if all_transactions:
        df_final = pd.DataFrame(all_transactions)
        print("\n" + "=" * 50)
        # cprint(f"📈 إجمالي عدد العمليات المستخرجة: {len(df_final)}")
        print(f"📈 total transactions extracted: {len(df_final)}")

        # cprint(
        #     f"💰 إجمالي المبالغ الدائنة (المحولة): {df_final['credit'].sum():,.2f} ج.م"
        # )
        print(
                    f"💰 total amount of money transfered: {df_final['credit'].sum():,.2f} ج.م"
                )
        print("=" * 50 + "\n")
        export_to_excel(df_final, OUTPUT_EXCEL_PATH)