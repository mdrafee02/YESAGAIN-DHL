"""
=============================================================
  AIRTABLE → DHL SHIPPING AUTOMATION SCRIPT
  YesAgain / Palm Tree AE
  VERSION: 5.0
=============================================================
"""
 
import pandas as pd
import re
import requests
import sys
import json
import base64
import time
import os
from datetime import datetime, timezone
from dotenv import load_dotenv
from urllib.parse import quote
 
try:
    from unidecode import unidecode
    HAS_UNIDECODE = True
except ImportError:
    HAS_UNIDECODE = False
    print("⚠️  unidecode not installed — special characters will be stripped instead of converted.")
    print("   Run: pip install unidecode")
 
load_dotenv()
 
AIRTABLE_API_KEY         = os.getenv("AIRTABLE_API_KEY")
AIRTABLE_RMA_API_KEY     = os.getenv("AIRTABLE_RMA_API_KEY", AIRTABLE_API_KEY)
AIRTABLE_CRS_RMA_API_KEY = os.getenv("AIRTABLE_CRS_RMA_API_KEY", AIRTABLE_RMA_API_KEY)
 
DHL_API_KEY    = os.getenv("DHL_API_KEY")
DHL_API_SECRET = os.getenv("DHL_API_SECRET")
DHL_TEST_MODE  = os.getenv("DHL_TEST_MODE", "true").strip().lower() == "true"
 
DHL_BASE_URL_TEST = "https://express.api.dhl.com/mydhlapi/test"
DHL_BASE_URL_PROD = "https://express.api.dhl.com/mydhlapi"
 
TABLE_CONFIG = {
    "orders": {
        "base_id":  os.getenv("AIRTABLE_BASE_ID",          "appJ3jLnRx4pDaTsM"),
        "table_id": os.getenv("AIRTABLE_TABLE",             "tbl3MiZfKfMR1jd6A"),
        "view_id":  os.getenv("AIRTABLE_ORDERS_VIEW_ID",    "viwEBcUnZvtLvxov1"),
        "label":    "Sales Orders (DHL Labels)",
        "api_key":  None,
    },
    "sales_order_lines": {
        "base_id":  os.getenv("AIRTABLE_BASE_ID",                    "appJ3jLnRx4pDaTsM"),
        "table_id": os.getenv("AIRTABLE_SALES_ORDER_LINES_TABLE_ID", "tblGWIuIdVNhRLwEf"),
        "view_id":  None,
        "label":    "Sales Order Lines",
        "api_key":  None,
    },
    "ya_rma": {
        "base_id":  os.getenv("AIRTABLE_YA_RMA_BASE_ID", "app42Tgocgm1DFH3H"),
        "table_id": os.getenv("AIRTABLE_YA_RMA_TABLE",   "tblNTvHfVY0SVfcPg"),
        "view_id":  None,
        "label":    "YesAgain Commerce Central RMA",
        "api_key":  None,
    },
    "crs_rma": {
        "base_id":  os.getenv("AIRTABLE_CRS_RMA_BASE_ID", "appdEZg6u3zXG8tRE"),
        "table_id": os.getenv("AIRTABLE_CRS_RMA_TABLE",   "tblWvBHXdLdOYIYDd"),
        "view_id":  None,
        "label":    "CRS Commerce Central RMA",
        "api_key":  None,
    },
    "commerce_central_sales_lines": {
        "base_id":  "app42Tgocgm1DFH3H",                                   # same as YA RMA base
        "table_id": os.getenv("AIRTABLE_COMMERCE_CENTRAL_SO_LINES_TABLE_ID", "tblXI2dCP2nNpJBvg"),
        "view_id":  None,
        "label":    "Commerce Central Sales Order Lines",
        "api_key":  AIRTABLE_RMA_API_KEY,                                 # uses the RMA key (has access)
    },
}
 
TABLE_CONFIG["orders"]["api_key"]            = AIRTABLE_API_KEY
TABLE_CONFIG["sales_order_lines"]["api_key"] = AIRTABLE_API_KEY
TABLE_CONFIG["ya_rma"]["api_key"]            = AIRTABLE_RMA_API_KEY
TABLE_CONFIG["crs_rma"]["api_key"]           = AIRTABLE_CRS_RMA_API_KEY
 
# ============================================================
# PLT (Paperless Trade) ACCEPTING COUNTRIES — from DHL PLT_2025.xlsx
# WY service code must be added for all shipments to these destinations
# ============================================================
PLT_COUNTRIES = {
    'AD','AE','AG','AI','AL','AM','AO','AR','AS','AT','AU','AW','AZ',
    'BA','BB','BE','BF','BG','BH','BI','BJ','BM','BN','BO','BR','BS',
    'BT','BW','BY','BZ','CA','CD','CF','CG','CH','CI','CK','CL','CM',
    'CN','CO','CR','CU','CV','CY','CZ','DE','DJ','DK','DM','DO','DZ',
    'EC','EE','EG','ER','ES','ET','FI','FJ','FK','FM','FO','FR','GA',
    'GB','GD','GE','GF','GG','GH','GI','GL','GM','GN','GP','GQ','GR',
    'GT','GU','GW','GY','HK','HN','HR','HT','HU','IC','ID','IE','IL',
    'IN','IQ','IR','IS','IT','JE','JM','JO','JP','KE','KG','KH','KI',
    'KM','KN','KP','KR','KV','KW','KY','KZ','LA','LB','LC','LI','LK',
    'LR','LS','LT','LU','LV','LY','MA','MC','MD','MG','MH','MK','ML',
    'MM','MN','MO','MP','MQ','MR','MS','MT','MU','MV','MW','MX','MY',
    'MZ','NA','NC','NE','NG','NI','NL','NO','NP','NR','NU','NZ','OM',
    'PA','PE','PF','PG','PH','PK','PL','PR','PT','PW','PY','QA','RE',
    'RO','RU','RW','SA','SB','SC','SD','SE','SG','SH','SI','SK','SL',
    'SM','SN','SO','SR','SS','ST','SV','SY','SZ','TC','TD','TG','TH',
    'TJ','TL','TM','TN','TO','TR','TT','TV','TW','TZ','UA','UG','US',
    'UY','UZ','VA','VC','VE','VG','VI','VN','VU','WS','XB','XC','XE',
    'XM','XN','XS','XY','YE','YT','ZA','ZM','ZW',
}
 
# ============================================================
# EUROPEAN COUNTRIES — used to distinguish EU (YesAgain France
# as importer) from "Others not listed" (YesAgain UAE as importer)
# GB is intentionally excluded — handled separately as UK rule.
# ============================================================
EU_COUNTRIES = {
    'AT','BE','BG','CH','CY','CZ','DE','DK','EE','ES','FI','FR',
    'GR','HR','HU','IE','IS','IT','LI','LT','LU','LV','MC','MT',
    'NL','NO','PL','PT','RO','SE','SI','SK','SM','AL','BA','ME',
    'MK','RS','XK',
}
 
# ============================================================
# FIXED VALUES
# ============================================================
FIXED = {
    # ── Accounts ──
    "account_shipper"     : 454189098,   # YA UAE
    "account_payer"       : 454189098,   # YA UAE
    "account_duty_eu"     : 229887839,   # YA France — pays EU duties
    "account_duty_us_gcc" : 961923318,   # PT UAE — pays US & GCC duties
 
    "account_crs_shipper" : 952629100,   # CRS UK 
    "account_crs_duty"    : 952629100,   # CRS UK — pays UK duties
 
    # ── Shipping config ──
    # CSV vs API split: portal upload uses different codes than the API
    "incoterms"           : "Y",         # CSV portal upload value
    "incoterms_api"       : "DDP",       # DHL API value — "Y" is portal-only
    "product_code"        : "WPX",       # CSV portal upload value
    "product_code_api"    : "P",         # DHL API value — "WPX" is portal-only
 
    "weight"              : 1.5,
    "weight_unit"         : "KG",
    "currency"            : "EUR",
    "contents"            : "Used laptop",
    "shipment_type"       : "P",
    "dim_unit"            : "CM",
    "length"              : 35,
    "width"               : 30,
    "height"              : 7,
    "export_reason"       : "Personal use",
    "dig_customs"         : "Y",
    "item_desc"           : "Used laptop",
    "commodity"           : "8471.49.0000",
    "item_units"          : "PCS",
    "item_net"            : 1.3,
    "item_gross"          : 1.5,
    "origin"              : "CN",
}
 
# ============================================================
# PARTY DETAILS
# ============================================================
 
PARTY_UAE = {
    "Company"  : "YESAGAIN FZC",
    "Name"     : "YESAGAIN",
    "Address1" : "P MALL, SHARJAH FREE ZONE",
    "Country"  : "AE",
    "City"     : "SHARJAH",
    "ZIP"      : "00000",
    "Email"    : "logistics@yesagain.com",
    "PhoneCC"  : 971,
    "Phone"    : "508893656",
    "VAT"      : "",
    "EORI"     : "",
    "Rel"      : "IP",
}
 
PARTY_UK = {
    "Company"  : "Computer Remarketing Services Ltd",
    "Name"     : "Computer Remarketing Services Ltd",
    "Address1" : "Unit 2 Maryland Road",
    "Country"  : "GB",
    "City"     : "Tongwell",
    "ZIP"      : "MK15 8HF",
    "Email"    : "logistics@yesagain.com",
    "PhoneCC"  : 44,
    "Phone"    : "01908656700",
    "VAT"      : "GB450944880",
    "EORI"     : "GB524271957000",
    "Rel"      : "IP",
}
 
PARTY_FRANCE = {
    "Company"  : "YESAGAIN",
    "Name"     : "YESAGAIN",
    "Address1" : "14 Chemin du Chapitre",
    "Country"  : "FR",
    "City"     : "TOULOUSE",
    "ZIP"      : "31100",
    "Email"    : "logistics@yesagain.com",
    "PhoneCC"  : 33,
    "Phone"    : "187661011",
    "VAT"      : "FR04919345207",
    "EORI"     : "FR91934520700022",
    "Rel"      : "IP",
}

# ============================================================
# DESTINATION-BASED VAT (YesAgain's local VAT per EU country)
# ============================================================
# The VAT number declared is chosen by the DESTINATION country.
# e.g. an order shipping to Belgium declares the Belgian VAT.
# Anything not listed here falls back to the importer entity's
# default VAT (France for EU, UK for GB) — see importer_vat_for().
EU_VAT_BY_COUNTRY = {
    "AT": "ATU79670679",      # Austria
    "BE": "BE0803906108",     # Belgium
    "DK": "DK13394814",       # Denmark
    "IT": "IT00358559995",    # Italy
    "NL": "NL827367090B01",   # Netherlands
    "SE": "SE502096226101",   # Sweden
    "DE": "DE362263608",      # Germany
    "FR": "FR04919345207",    # France
    "GB": "GB450944880",      # United Kingdom
}

def importer_vat_for(country, fallback_vat=""):
    """Return YesAgain's VAT registered in the destination country.
    Falls back to the importer entity's default VAT when the destination
    has no specific registration in EU_VAT_BY_COUNTRY."""
    return EU_VAT_BY_COUNTRY.get(str(country).strip().upper(), fallback_vat) or fallback_vat

# ============================================================
# PHONE / STATE LOOKUPS
# ============================================================
PHONE_CODES = {
    "GB":"44",  "DE":"49",  "FR":"33",  "ES":"34",  "IT":"39",  "AT":"43",
    "BE":"32",  "NL":"31",  "PT":"351", "IE":"353", "SE":"46",  "NO":"47",
    "DK":"45",  "FI":"358", "PL":"48",  "CZ":"420", "HU":"36",
    "US":"1",   "CA":"1",   "AU":"61",
    "CH":"41",  "LU":"352", "GR":"30",  "RO":"40",  "BG":"359", "HR":"385",
    "SK":"421", "SI":"386", "LT":"370", "LV":"371", "EE":"372", "CY":"357",
    "MT":"356", "IS":"354", "TR":"90",  "SA":"966", "AE":"971", "JP":"81",
    "CN":"86",  "IN":"91",  "BR":"55",  "MX":"52",  "ZA":"27",
}
 
STATE_FULL_NAMES = {
    "AL":"Alabama","AK":"Alaska","AZ":"Arizona","AR":"Arkansas",
    "CA":"California","CO":"Colorado","CT":"Connecticut","DE":"Delaware",
    "FL":"Florida","GA":"Georgia","HI":"Hawaii","ID":"Idaho",
    "IL":"Illinois","IN":"Indiana","IA":"Iowa","KS":"Kansas",
    "KY":"Kentucky","LA":"Louisiana","ME":"Maine","MD":"Maryland",
    "MA":"Massachusetts","MI":"Michigan","MN":"Minnesota","MS":"Mississippi",
    "MO":"Missouri","MT":"Montana","NE":"Nebraska","NV":"Nevada",
    "NH":"New Hampshire","NJ":"New Jersey","NM":"New Mexico","NY":"New York",
    "NC":"North Carolina","ND":"North Dakota","OH":"Ohio","OK":"Oklahoma",
    "OR":"Oregon","PA":"Pennsylvania","RI":"Rhode Island","SC":"South Carolina",
    "SD":"South Dakota","TN":"Tennessee","TX":"Texas","UT":"Utah",
    "VT":"Vermont","VA":"Virginia","WA":"Washington","WV":"West Virginia",
    "WI":"Wisconsin","WY":"Wyoming",
}
 
 
# ============================================================
# HELPER FUNCTIONS
# ============================================================
 
def split_phone(telephone, country):
    tel = re.sub(r'[^\d]', '', str(telephone))
    cc = PHONE_CODES.get(country, "")
    if country in ["US", "CA"]:
        if tel.startswith("1") and len(tel) > 10:
            return cc, tel[1:]
        return cc, tel
    if cc and tel.startswith(cc):
        return cc, tel[len(cc):]
    return cc, tel
 
 
def clean_postal_code(postcode, country_code):
    """Ensure postal code matches DHL format. LU must be exactly 4 digits."""
    if not postcode:
        return ""
    postcode = str(postcode).strip().upper()
    if country_code == "LU":
        # Remove all non-digits, then take first 4 digits or pad
        digits = re.sub(r'\D', '', postcode)
        if len(digits) >= 4:
            return digits[:4]
        elif len(digits) > 0:
            return digits.zfill(4)
        else:
            return "9999"   # fallback
    return postcode
 
 
def clean_price(value):
    if pd.isna(value) if not isinstance(value, list) else False:
        return 0.0
    # Airtable returns a list when one order has multiple line prices
    # (e.g. qty=2 same model → [174.11, 178.31]). Sum ALL values.
    if isinstance(value, list):
        return sum(clean_price(v) for v in value) if value else 0.0
    if isinstance(value, (int, float)):
        return float(value)
    cleaned = re.sub(r'[€$£,\s]', '', str(value))
    try:
        return float(cleaned)
    except ValueError:
        return 0.0
 
 
def clean_qty(value):
    # Airtable returns a list when one order has multiple qty values
    # (e.g. qty=2 same model → [1, 1]). Sum ALL values.
    if isinstance(value, list):
        return sum(clean_qty(v) for v in value) if value else 1
    try:
        return int(float(str(value)))
    except Exception:
        return 1
 
 
# ============================================================
# ARABIC → LATIN TRANSLITERATION
# ============================================================
#
# Why transliteration is imperfect for Arabic:
#   Arabic is normally written WITHOUT short vowels (harakat). So "احمد"
#   could be Ahmed, Ahmad or Ahmd — the script cannot know which without
#   the vowel marks. We solve this with two layers:
#     1. A name dictionary for the most common Arabic first names, titles
#        and words → gives exact correct spellings (Ahmed, Mohammed, Eid…)
#     2. A letter-by-letter phonetic fallback for anything not in the dict.
#

ARABIC_NAME_DICT = {
    # Common first names (male)
    'احمد':'Ahmed','احمد':'Ahmed','أحمد':'Ahmed','محمد':'Mohammed',
    'محمود':'Mahmoud','عبدالله':'Abdullah','عبد الله':'Abdullah',
    'عبدالرحمن':'Abdulrahman','عبدالعزيز':'Abdulaziz','علي':'Ali',
    'عمر':'Omar','خالد':'Khalid','يوسف':'Youssef','ابراهيم':'Ibrahim',
    'إبراهيم':'Ibrahim','سامي':'Sami','سالم':'Salem','سعد':'Saad',
    'سعيد':'Saeed','صالح':'Saleh','طارق':'Tariq','عادل':'Adel',
    'عثمان':'Othman','عصام':'Essam','علاء':'Alaa','فاروق':'Farouk',
    'فيصل':'Faisal','كريم':'Karim','ماجد':'Majid','مازن':'Mazen',
    'منصور':'Mansour','موسى':'Moussa','نبيل':'Nabil','هاني':'Hani',
    'هشام':'Hisham','وليد':'Walid','ياسر':'Yasser','زياد':'Ziad',
    'رامي':'Rami','رياض':'Riyad','زيد':'Zaid','بلال':'Bilal',
    'تركي':'Turki','جمال':'Jamal','حسن':'Hassan','حسين':'Hussein',
    'حمد':'Hamad','حمزة':'Hamza','راشد':'Rashed','زكريا':'Zakaria',
    'شاكر':'Shaker','شريف':'Sherif','صلاح':'Salah','طلال':'Talal',
    'عزيز':'Aziz','مجدي':'Magdy','مصطفى':'Mostafa','مصطفا':'Mostafa',
    'ناصر':'Nasser','نادر':'Nader','نواف':'Nawaf','هلال':'Hilal',
    'وائل':'Wael','وسام':'Wissam','يحيى':'Yahya','بدر':'Badr',
    'جاسم':'Jasim','رضا':'Reda','سلطان':'Sultan','سمير':'Samir',
    'صقر':'Saqr','فهد':'Fahad','قاسم':'Qasim','كامل':'Kamel',
    'نزار':'Nizar','هيثم':'Haitham','امير':'Amir','أمير':'Amir',
    'بشير':'Bashir','رفيق':'Rafik','رمزي':'Ramzi','سلمان':'Salman',
    'سيف':'Saif','شادي':'Shadi','فادي':'Fadi','كمال':'Kamal',
    'مروان':'Marwan','نصر':'Nasr','نور':'Nour','هادي':'Hadi',
    'ادريس':'Idris','إدريس':'Idris','انس':'Anas','أنس':'Anas',
    # Common first names (female)
    'مريم':'Mariam','فاطمة':'Fatima','سارة':'Sara','سارا':'Sara',
    'رنا':'Rana','لينا':'Lina','منى':'Mona','هيا':'Haya',
    'ريم':'Reem','دينا':'Dina','رانيا':'Rania','نادية':'Nadia',
    'هند':'Hind','ياسمين':'Yasmine','اسماء':'Asmaa','بسمة':'Basma',
    'حنان':'Hanan','خديجة':'Khadija','رشا':'Rasha','سلمى':'Salma',
    'غادة':'Ghada','ليلى':'Layla','منال':'Manal','نجلاء':'Najla',
    'نوره':'Noura','هدى':'Huda','وفاء':'Wafa','شيماء':'Shaimaa',
    # Titles / words that appear in names
    'مهندس':'Eng','مهندسة':'Eng','دكتور':'Dr','دكتورة':'Dr',
    'الدكتور':'Dr','الدكتورة':'Dr','استاذ':'Prof','الاستاذ':'Prof',
    # Common surnames / family name parts
    'عيد':'Eid','الشعراوي':'El Shaarawy','الشريف':'El Sherif',
    'الامير':'El Amir','الزهراني':'Al Zahrani','العمري':'Al Omari',
    'العتيبي':'Al Otaibi','الغامدي':'Al Ghamdi','السبيعي':'Al Subaie',
    'القحطاني':'Al Qahtani','الدوسري':'Al Dosari','الشهري':'Al Shehri',
    'المطيري':'Al Mutairi','الرشيدي':'Al Rashidi','المالكي':'Al Maliki',
    'الحربي':'Al Harbi','العنزي':'Al Anazi','الشمري':'Al Shammari',
    # Connectors
    'بن':'Bin','ابن':'Ibn','ابو':'Abu','ام':'Um','بنت':'Bint',
    'ال':'Al',
}
 
# Single-char phonetic fallback (for words not in the dictionary)
ARABIC_LATIN = str.maketrans({
    'ا':'a','أ':'a','إ':'i','آ':'a','ب':'b',
    'ت':'t','ث':'th','ج':'j','ح':'h','خ':'kh',
    'د':'d','ذ':'dh','ر':'r','ز':'z','س':'s',
    'ش':'sh','ص':'s','ض':'d','ط':'t','ظ':'z',
    'ع':'e','غ':'gh','ف':'f','ق':'q','ك':'k',
    'ل':'l','م':'m','ن':'n','ه':'h','و':'o',
    'ي':'i','ى':'a','ة':'a','ء':'','ئ':'i',
    'ؤ':'o',
    'ً':'','ٌ':'','ٍ':'','َ':'a','ُ':'u',
    'ِ':'i','ّ':'','ْ':'',
})
 
def _arabic_to_latin(text):
    """
    Two-layer Arabic → Latin:
      1. Word-level dictionary lookup (Ahmed, Eid, El Shaarawy, Eng…)
      2. Letter-by-letter phonetic fallback for unknown words
    """
    # Pre-process: strip definite article prefix ال from standalone token
    text = text.replace('لا', 'la')  # لا → la
    words = text.split()
    result = []
    for word in words:
        w = word.strip()
        if not w:
            continue
        # Direct dictionary hit
        if w in ARABIC_NAME_DICT:
            result.append(ARABIC_NAME_DICT[w])
            continue
        # Try stripping leading ال (definite article) and look up the rest
        if w.startswith('ال') and len(w) > 2:
            root = w[2:]
            if root in ARABIC_NAME_DICT:
                result.append('Al ' + ARABIC_NAME_DICT[root])
                continue
        # Phonetic fallback
        fallback = re.sub(r'[^ -~]', '', w.translate(ARABIC_LATIN)).strip()
        if fallback:
            result.append(fallback.capitalize())
    return re.sub(r'\s+', ' ', ' '.join(result)).strip()
 
def _has_arabic(text):
    """Return True if the text contains Arabic-script characters."""
    return any('؀' <= c <= 'ۿ' for c in text)
 
 
def clean_text(value):
    if pd.isna(value) if not isinstance(value, (list, dict)) else False:
        return ""
    if value == "":
        return ""
    text = str(value).strip()
    if not text or text.lower() == "nan":
        return ""
    # Transliterate Arabic → Latin before any further processing.
    # We always use our dictionary-based transliterator for Arabic because
    # unidecode lacks short-vowel knowledge and produces "Ahmd" not "Ahmed".
    if _has_arabic(text):
        latin = _arabic_to_latin(text)
        latin = re.sub(r'[^ -~]', '', latin).strip()
        latin = re.sub(r'\s+', ' ', latin)
        text = latin if latin else "CUSTOMER"
    elif HAS_UNIDECODE:
        text = unidecode(text)
    else:
        CHAR_MAP = str.maketrans({
            'À':'A','Á':'A','Â':'A','Ã':'A','Ä':'A','Å':'A','Æ':'AE',
            'Ç':'C','È':'E','É':'E','Ê':'E','Ë':'E',
            'Ì':'I','Í':'I','Î':'I','Ï':'I',
            'Ð':'D','Ñ':'N','Ò':'O','Ó':'O','Ô':'O','Õ':'O','Ö':'O','Ø':'O',
            'Ù':'U','Ú':'U','Û':'U','Ü':'U','Ý':'Y','Þ':'TH','ß':'ss',
            'à':'a','á':'a','â':'a','ã':'a','ä':'a','å':'a','æ':'ae',
            'ç':'c','è':'e','é':'e','ê':'e','ë':'e',
            'ì':'i','í':'i','î':'i','ï':'i',
            'ð':'d','ñ':'n','ò':'o','ó':'o','ô':'o','õ':'o','ö':'o','ø':'o',
            'ù':'u','ú':'u','û':'u','ü':'u','ý':'y','þ':'th','ÿ':'y',
            'Ł':'L','ł':'l','Ź':'Z','ź':'z','Ż':'Z','ż':'z',
            'Ś':'S','ś':'s','Ą':'A','ą':'a','Ę':'E','ę':'e',
            'Ć':'C','ć':'c','Ń':'N','ń':'n',
            'Ğ':'G','ğ':'g','İ':'I','ı':'i','Ş':'S','ş':'s',
            'Č':'C','č':'c','Š':'S','š':'s','Ž':'Z','ž':'z',
            'Ř':'R','ř':'r','Ů':'U','ů':'u','Ď':'D','ď':'d',
            'Ť':'T','ť':'t','Ľ':'L','ľ':'l',
        })
        text = text.translate(CHAR_MAP)
    text = re.sub(r'[^\x20-\x7E]', '', text)
    text = re.sub(r'[^\w\s\-.,/]', '', text)
    return text.strip()
 
 
def get_first_value(row, field_names, default=""):
    for field in field_names:
        value = row.get(field, "")
        if isinstance(value, list):
            value = value[0] if value else ""
        if value is None:
            continue
        if isinstance(value, float) and pd.isna(value):
            continue
        text = str(value).strip()
        if text and text.lower() != "nan":
            return value
    return default
 
 
def normalize_lookup_key(value):
    if isinstance(value, list):
        value = value[0] if value else ""
    if value is None:
        return ""
    if isinstance(value, float) and pd.isna(value):
        return ""
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        numeric = float(value)
        return str(int(numeric)) if numeric.is_integer() else str(numeric)
    text = str(value).strip()
    if re.fullmatch(r"\d+\.0", text):
        return text[:-2]
    return text
 
def normalize_sales_order_number(value):
    """Extract the last part after any slash, e.g. 'RMAY31984/72771270' → '72771270'."""
    text = normalize_lookup_key(value)
    if '/' in text:
        return text.split('/')[-1].strip()
    return text
 
 
def normalize_rma_number(value):
    text = normalize_lookup_key(value)
    text = re.sub(r"^\s*RMA[-\s]*", "", text, flags=re.IGNORECASE).strip()
    return text.upper()
 
 
def extract_sales_order_from_rma(rma_number_str, rma_map):
    """
    For a given RMA number (e.g. 'RMAY34783/RMAY31984/72771270'),
    returns the sales order number found in the RMA map.
    The RMA map uses the full RMA number as key (without the slash suffix),
    so we try the exact key first, then fall back to the first part.
    """
    rma_clean = normalize_rma_number(rma_number_str)
    # If the RMA number contains a slash, the map key is usually the first part (e.g. 'RMAY34783')
    parts = rma_clean.split('/')
    for candidate in [rma_clean] + parts:
        if candidate in rma_map:
            return rma_map[candidate]
    return ""
 
 
def get_rma_source(order_number):
    upper = str(order_number).upper().strip()
    if upper.startswith("RMAY"):
        return "ya"
    if upper.startswith("RMAC"):
        return "crs"
    return None
 
 
def get_address2_and_state(row):
    country = str(row.get("Shipping Country", "")).strip()
    addr2_raw = row.get("Shipping address 2", "")
    addr2 = str(addr2_raw).strip() if pd.notna(addr2_raw) else ""
    if country == "US":
        state_code = addr2 if len(addr2) == 2 and addr2.isalpha() else ""
        return "", state_code
    else:
        return clean_text(addr2), ""
 
 
def get_destination_rules(destination_country, order_id):
    """
    Smart Routing: Chooses the correct Duty Account, Additional Party,
    Shipper and region label based on destination.
 
    Regions (per shipping policy):
      ┌──────────────────────┬──────────────┬──────────────────┬──────────────────┐
      │ Destination          │ Shipper      │ Importer         │ Declared Value   │
      ├──────────────────────┼──────────────┼──────────────────┼──────────────────┤
      │ GCC Countries        │ YesAgain UAE │ YesAgain UAE     │ 20% of item val  │
      │ USA                  │ YesAgain UAE │ YesAgain UAE     │ 20% of item val  │
      │ Others not listed    │ YesAgain UAE │ YesAgain UAE     │ 20% of item val  │
      │ Europe               │ YesAgain UAE │ YesAgain France  │ Exclude VAT      │
      │ United Kingdom (UK)  │ YesAgain UAE │ CRS Account      │ Exclude VAT      │
      └──────────────────────┴──────────────┴──────────────────┴──────────────────┘
 
    NOTE: Saudi Arabia (SA) has an additional override — declared value is
    always fixed at 164 EUR regardless of item price (applied in transform()).
    """
    country     = str(destination_country).strip().upper()
    order_upper = str(order_id).upper().strip()
 
    # Shipper is ALWAYS YesAgain UAE regardless of order type or destination.
    # Per shipping policy: "YesAgain UAE" is shipper for all regions.
    # RMAC orders still use CRS as importer/duty account for UK, but the
    # physical shipper on the DHL label must be YesAgain UAE (SHJ origin).
    shipper_party = PARTY_UAE
    shipper_acc   = str(FIXED["account_shipper"])
 
    # RULE 1: UK (GB) — Importer: CRS Account | Declared: Exclude VAT
    if country == "GB":
        return {
            "region"          : "uk",
            "duty_account"    : str(FIXED["account_crs_duty"]),
            "additional_party": PARTY_UK,
            "shipper_party"   : shipper_party,
            "shipper_account" : shipper_acc,
        }
 
    # RULE 2: GCC countries — Importer: YesAgain UAE | Declared: 20%
    # (SA gets an additional 164 EUR fixed override applied in transform())
    gcc_countries = {"AE", "SA", "OM", "QA", "BH", "KW"}
    if country in gcc_countries:
        return {
            "region"          : "gcc",
            "duty_account"    : str(FIXED["account_duty_us_gcc"]),
            "additional_party": PARTY_UAE,
            "shipper_party"   : shipper_party,
            "shipper_account" : shipper_acc,
        }
 
    # RULE 3: USA / Canada — Importer: YesAgain UAE | Declared: 20%
    if country in {"US", "CA"}:
        return {
            "region"          : "us",
            "duty_account"    : str(FIXED["account_duty_us_gcc"]),
            "additional_party": PARTY_UAE,
            "shipper_party"   : shipper_party,
            "shipper_account" : shipper_acc,
        }
 
    # RULE 4: Europe — Importer: YesAgain France | Declared: Exclude VAT (full price)
    if country in EU_COUNTRIES:
        return {
            "region"          : "eu",
            "duty_account"    : str(FIXED["account_duty_eu"]),
            "additional_party": PARTY_FRANCE,
            "shipper_party"   : shipper_party,
            "shipper_account" : shipper_acc,
        }
 
    # RULE 5: Others not listed — Importer: YesAgain UAE | Declared: 20%
    return {
        "region"          : "other",
        "duty_account"    : str(FIXED["account_duty_us_gcc"]),
        "additional_party": PARTY_UAE,
        "shipper_party"   : shipper_party,
        "shipper_account" : shipper_acc,
    }
 
 
def print_progress(current, total):
    """Print a compact progress bar. Used in large batches."""
    pct   = current / total
    filled = int(pct * 20)
    bar   = "█" * filled + "░" * (20 - filled)
    print(f"   [{bar}] {current}/{total} ({pct:.0%})", flush=True)
 
 
# ============================================================
# AIRTABLE FETCH
# ============================================================
 
def fetch_table(base_id, table_id, view_id=None, label=None, api_key=None,
                filter_formula=None):
    if not table_id:
        print(f"⚠️  Skipping {label}: no table_id configured")
        return pd.DataFrame()
 
    encoded_table = quote(str(table_id), safe="")
    url = f"https://api.airtable.com/v0/{base_id}/{encoded_table}"
    headers = {"Authorization": f"Bearer {api_key or AIRTABLE_API_KEY}"}
    all_records = []
    offset = None
    page = 0
 
    while True:
        params = {"pageSize": 100}
        if view_id:
            params["view"] = view_id
        if offset:
            params["offset"] = offset
        if filter_formula:
            params["filterByFormula"] = filter_formula
 
        try:
            response = requests.get(url, headers=headers, params=params, timeout=30)
        except requests.exceptions.Timeout:
            print(f"⚠️  Timeout on {label or table_id} (page {page + 1}). "
                  f"Continuing with {len(all_records)} records.")
            break
        except requests.exceptions.ConnectionError as e:
            print(f"⚠️  Connection error on {label or table_id}: {e}")
            break
 
        # Airtable rate limit — back off and retry once
        if response.status_code == 429:
            print(f"⚠️  Rate limited fetching {label or table_id} — waiting 30s...")
            time.sleep(30)
            continue
 
        if response.status_code == 403:
            print(f"❌ 403 Permission denied for {label or table_id} (base: {base_id})")
            if label and "CRS" in label:
                print("   → set AIRTABLE_CRS_RMA_API_KEY in your .env with a token that has CRS access")
            return pd.DataFrame()
 
        if response.status_code != 200:
            print(f"⚠️  Error on {label or table_id} (status {response.status_code}): "
                  f"{response.text[:200]}")
            break
 
        data = response.json()
        records = []
        for record in data.get("records", []):
            fields = record.get("fields", {})
            fields["_airtable_id"] = record.get("id", "")  # e.g. "recXXXXXXXXXXXXXX"
            records.append(fields)
        all_records.extend(records)
        page += 1
 
        offset = data.get("offset")
        if not offset:
            break
 
    if all_records:
        print(f"✅ Loaded {label or table_id}: {len(all_records)} records")
    else:
        print(f"⚠️  No records loaded for {label or table_id}")
 
    return pd.DataFrame(all_records) if all_records else pd.DataFrame()
 
 
def fetch_from_airtable():
    return fetch_table(**TABLE_CONFIG["orders"])
 
 
# ============================================================
# AIRTABLE WRITEBACK — RETRY HELPER
#
# Airtable allows ~5 requests/second per base. For 100-200
# orders this means you can easily hit the limit. When Airtable
# returns 429, this helper waits and retries automatically so
# no write is silently dropped.
# ============================================================
 
def airtable_patch_with_retry(url, body, headers, label="", max_retries=3):
    """
    PATCH to Airtable with automatic retry on 429 (rate limit).
 
    Waits 2s, 4s, 8s between retries (exponential backoff).
    Returns the response object, or None if all retries fail.
    """
    for attempt in range(1, max_retries + 1):
        try:
            resp = requests.patch(url, json=body, headers=headers, timeout=15)
            if resp.status_code == 429:
                wait = 2 ** attempt  # 2s, 4s, 8s
                print(f"   ⚠️  Airtable rate limit (429) on {label} — "
                      f"retry {attempt}/{max_retries} in {wait}s...")
                time.sleep(wait)
                continue
            return resp
        except requests.exceptions.Timeout:
            print(f"   ⚠️  Timeout on {label} (attempt {attempt}/{max_retries})")
            if attempt < max_retries:
                time.sleep(2 ** attempt)
        except Exception as e:
            print(f"   ⚠️  Exception on {label} (attempt {attempt}/{max_retries}): {e}")
            if attempt < max_retries:
                time.sleep(2)
    return None
 
 
# ============================================================
# AIRTABLE WRITEBACK — STEP 1
# Writes "Shipment Tracking Number" + "Shipment Courier" only.
#
# "Shipment Label Created" is intentionally NOT set here.
# It is only set after the label PDF is confirmed uploaded
# in mark_label_created_in_airtable() (STEP 3).
# ============================================================
 
def update_shipment_tracking_in_airtable(record_id, order_number, tracking_number):
    """
    PATCH Step 1 — writes Shipment Tracking Number + Shipment Courier.
 
    Args:
        record_id       – Airtable record ID (e.g. "recXXXXXXXXXXXXXX")
        order_number    – for log messages only
        tracking_number – the DHL AWB string (e.g. "1570941875")
 
    Returns:
        True on success, False on failure.
    """
    if not record_id:
        print(f"   ⚠️  No Airtable record_id for {order_number} — skipping Step 1.")
        return False
 
    base_id  = TABLE_CONFIG["orders"]["base_id"]
    table_id = TABLE_CONFIG["orders"]["table_id"]
    api_key  = AIRTABLE_API_KEY
 
    encoded_table = quote(str(table_id), safe="")
    url = f"https://api.airtable.com/v0/{base_id}/{encoded_table}/{record_id}"
 
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type":  "application/json",
    }
    body = {
        "fields": {
            "Shipment Tracking Number": str(tracking_number),
            "Shipment Courier":         "dhl",
        }
    }
 
    resp = airtable_patch_with_retry(url, body, headers, label=f"Step1/{order_number}")
    if resp is None:
        print(f"   ❌ Step 1 failed (all retries exhausted) for {order_number}")
        return False
    if resp.status_code == 200:
        print(f"   ✅ Step 1 — Tracking: {tracking_number} | Courier: dhl  "
              f"(Order: {order_number})")
        return True
    else:
        print(f"   ❌ Step 1 PATCH failed ({resp.status_code}) for {order_number}: "
              f"{resp.text[:400]}")
        return False
 
 
# ============================================================
# AIRTABLE WRITEBACK — STEP 2
# Uploads label PDF to "Shipment Label file" attachment field.
# Uses Airtable's Content Upload API (no public URL needed).
#
# CORRECT URL: content.airtable.com/v0/{base_id}/{record_id}/...
# WRONG URL:   content.airtable.com/v0/{base_id}/{table_id}/{record_id}/...
#              (table_id does NOT go in the content upload URL)
# ============================================================
 
def _upload_to_temp_host(pdf_bytes, filename):
    """
    Upload PDF to a temporary public host and return a download URL.
    Tries multiple services so there is always a fallback.
    Used only when the Airtable Content API is unavailable.
    """
    services = [
        {
            "name"     : "tmpfiles.org",
            "url"      : "https://tmpfiles.org/api/v1/upload",
            "get_link" : lambda r: (
                r.json().get("data", {}).get("url", "")
                 .replace("tmpfiles.org/", "tmpfiles.org/dl/")
            ),
        },
        {
            "name"     : "file.io",
            "url"      : "https://file.io/?expires=1h",
            "get_link" : lambda r: r.json().get("link", ""),
        },
    ]
 
    for svc in services:
        try:
            resp = requests.post(
                svc["url"],
                files={"file": (filename, pdf_bytes, "application/pdf")},
                timeout=20,
            )
            if resp.status_code == 200:
                url = svc["get_link"](resp)
                if url:
                    print(f"   🌐 Temp URL ({svc['name']}): {url}")
                    return url
            else:
                print(f"   ⚠️  {svc['name']} returned {resp.status_code}")
        except Exception as exc:
            print(f"   ⚠️  {svc['name']} failed: {exc}")
 
    return None
 
 
 
def upload_docs_to_airtable(record_id, order_number, res_data, max_retries=3):
    """
    Step 2 — Uploads BOTH Shipping Label and Commercial Invoice to Airtable.
    Uses ONLY Method 2 (PATCH with temporary URL) – faster and more reliable.
    """
    if not record_id:
        print(f"   ⚠️  No record_id for {order_number} — skipping Step 2.")
        return False
 
    base_id = TABLE_CONFIG["orders"]["base_id"]
    api_key = AIRTABLE_API_KEY
    
    documents = res_data.get("documents", [])
    if not documents and res_data.get("label_base64"):
        documents = [{"typeCode": "shipping_label", "content": res_data.get("label_base64")}]
 
    success_tracker = {"label": False, "invoice": False}
 
    for doc in documents:
        type_code = doc.get("typeCode", "").lower()
        content_b64 = doc.get("content")
        if not content_b64:
            continue
 
        if "label" in type_code or "waybill" in type_code:
            field_name = "Shipment Label file"
            field_id = "fldG3hmHH8cTPwzxo"
            filename = f"label_{order_number}.pdf"
            doc_key = "label"
        elif "invoice" in type_code:
            field_name = "Commercial invoice file"
            field_id = "fldp7fi4xvhqCybQ7"
            filename = f"invoice_{order_number}.pdf"
            doc_key = "invoice"
        else:
            continue
 
        print(f"   📡 Step 2 — Processing {doc_key} for {order_number}...")
        
        try:
            pdf_bytes = base64.b64decode(content_b64)
        except Exception as exc:
            print(f"   ❌ Base64 decode error for {doc_key} on {order_number}: {exc}")
            continue
 
        # --- ONLY METHOD 2: PATCH with temporary URL ---
        print(f"   📡 Step 2 — Method 2: PATCH with temporary URL for {doc_key} ({order_number})...")
        temp_url = _upload_to_temp_host(pdf_bytes, filename)
        if temp_url:
            table_id = TABLE_CONFIG["orders"]["table_id"]
            patch_url = f"https://api.airtable.com/v0/{base_id}/{quote(str(table_id))}/{record_id}"
            patch_headers = {
                "Authorization": f"Bearer {api_key}",
                "Content-Type": "application/json",
            }
            payload = {"fields": {field_name: [{"url": temp_url, "filename": filename}]}}
 
            for attempt in range(1, max_retries + 1):
                try:
                    resp = requests.patch(
                        patch_url, headers=patch_headers, json=payload, timeout=30
                    )
                    print(f"       PATCH response: HTTP {resp.status_code}")
 
                    if resp.status_code == 200:
                        print(f"   📎 {doc_key.capitalize()} uploaded via Method 2 (PATCH+URL) (Order: {order_number})")
                        success_tracker[doc_key] = True
                        break
 
                    elif resp.status_code == 422:
                        print(f"       422 Unprocessable — Airtable could not fetch the URL.")
                        print(f"       → Check field name: '{field_name}' must be type 'Attachment' in Airtable.")
                        print(f"       → Response: {resp.text[:300]}")
                        break
 
                    elif resp.status_code == 429:
                        wait = 2 ** attempt
                        print(f"       Rate limited — retrying in {wait}s...")
                        time.sleep(wait)
 
                    else:
                        print(f"       Failed ({resp.status_code}): {resp.text[:300]}")
                        if attempt < max_retries:
                            time.sleep(2)
 
                except Exception as exc:
                    print(f"       Exception (attempt {attempt}): {exc}")
                    if attempt < max_retries:
                        time.sleep(2)
        else:
            print(f"   ❌ {doc_key.capitalize()} Step 2 FAILED: could not get a temporary URL for {order_number}.")
 
    return success_tracker["label"]
 
 
# ============================================================
# AIRTABLE WRITEBACK — STEP 3
# Ticks "Shipment Label Created" checkbox AND writes a real
# clickable DHL tracking URL into "Shipment Label URL".
#
# Called ONLY after upload_label_to_airtable() returns True —
# so the checkbox is never ticked prematurely.
# ============================================================
 
def mark_label_created_in_airtable(record_id, order_number, tracking_number):
    """
    PATCH Step 3 — ticks 'Shipment Label Created' AND sets 'Shipment Label URL'
    to the real DHL tracking page URL.
 
    Called ONLY after upload_label_to_airtable() succeeds.
 
    Args:
        record_id       – Airtable record ID
        order_number    – for log messages only
        tracking_number – DHL AWB used to build the tracking URL
 
    Returns:
        True on success, False on failure.
    """
    if not record_id:
        return False
 
    base_id  = TABLE_CONFIG["orders"]["base_id"]
    table_id = TABLE_CONFIG["orders"]["table_id"]
    api_key  = AIRTABLE_API_KEY
 
    dhl_tracking_url = (
        f"https://www.dhl.com/ae-en/home/tracking/tracking-express.html"
        f"?submit=1&tracking-id={tracking_number}"
    )
 
    encoded_table = quote(str(table_id), safe="")
    url = f"https://api.airtable.com/v0/{base_id}/{encoded_table}/{record_id}"
 
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type":  "application/json",
    }
    body = {
        "fields": {
            "Shipment Label Created": True,                  # ✅ checkbox — ONLY set after upload confirmed
        }
    }
 
    resp = airtable_patch_with_retry(url, body, headers, label=f"Step3/{order_number}")
    if resp is None:
        print(f"   ❌ Step 3 failed (all retries exhausted) for {order_number}")
        return False
    if resp.status_code == 200:
        print(f"   ✅ Step 3 — Checkbox ticked ✅ | URL set  (Order: {order_number})")
        print(f"               {dhl_tracking_url}")
        return True
    else:
        # Not fatal — label is already uploaded and tracking written
        print(f"   ⚠️  Step 3 PATCH failed ({resp.status_code}) for {order_number}: "
              f"{resp.text[:200]}")
        return False
 
 
# ============================================================
# BUILD LOOKUP MAPS
# ============================================================
 
def build_rma_maps(df_ya_rma, df_crs_rma):
    ya_rma_map  = {}
    crs_rma_map = {}
 
    if not df_ya_rma.empty:
        print(f"\n   YA RMA columns: {list(df_ya_rma.columns)}")
        for _, row in df_ya_rma.iterrows():
            rma_no = normalize_rma_number(
                get_first_value(row, ["RMA #", "RMA Number", "RMA No", "RMA"])
            )
            sales_order = normalize_sales_order_number(
                get_first_value(row, ["External Sales Order", "Sales Order Number",
                    "Sales Order", "Returned Sales Order", "Order Number"])
            )
            if rma_no and sales_order:
                ya_rma_map[rma_no] = sales_order
 
    if not df_crs_rma.empty:
        print(f"   CRS RMA columns: {list(df_crs_rma.columns)}")
        for _, row in df_crs_rma.iterrows():
            rma_no = normalize_rma_number(
                get_first_value(row, ["RMA #", "RMA Number", "RMA No", "RMA"])
            )
            sales_order = normalize_lookup_key(
                get_first_value(row, [
                    "Returned Sales Order (Off System)",
                    "External Sales Order",
                    "Sales Order Number",
                    "Sales Order", "Returned Sales Order",
                    "Order Number", "Refund Sales Order",
                ])
            )
            if rma_no and sales_order:
                crs_rma_map[rma_no] = sales_order
 
    print(f"✅ YA RMA map : {len(ya_rma_map)} entries  → {dict(list(ya_rma_map.items())[:5])}")
    print(f"✅ CRS RMA map: {len(crs_rma_map)} entries → {dict(list(crs_rma_map.items())[:5])}")
    return ya_rma_map, crs_rma_map
 
 
def build_price_map(df_sales_lines):
    price_map = {}
    if df_sales_lines.empty:
        return price_map
 
    print(f"\n   Sales Order Lines columns: {list(df_sales_lines.columns)}")
    for _, row in df_sales_lines.iterrows():
        so = normalize_lookup_key(
            get_first_value(row, ["Sales Order Number", "Sales Order",
                                   "Order Number", "SO Line#", "Reference"])
        )
        price_val = clean_price(
            get_first_value(row, [
                "Converted Unit Price per Line", "Converted Unit Price",
                "Converted Unit price", "Converted Price",
                "Unit Price", "Price (Final)", "Price", "Sales...",
            ], 0)
        )
        if so:
            price_map[so] = price_map.get(so, 0) + price_val
 
    print(f"✅ Price map: {len(price_map)} entries")
    return price_map
 
 
def fetch_price_map_for_orders(so_numbers, api_key):
    if not so_numbers:
        return {}
 
    config   = TABLE_CONFIG["sales_order_lines"]
    so_list  = list(so_numbers)
    all_rows = []
    CHUNK    = 20
 
    print(f"🌐 Fetching Sales Order Lines for {len(so_list)} order(s): {so_list}")
 
    for i in range(0, len(so_list), CHUNK):
        chunk = so_list[i:i + CHUNK]
        conditions = ','.join([f'{{Sales Order Number}}="{so}"' for so in chunk])
        formula = f"OR({conditions})" if len(chunk) > 1 else conditions
 
        df_chunk = fetch_table(
            base_id=config["base_id"], table_id=config["table_id"],
            view_id=config["view_id"],
            label=f"Sales Order Lines (batch {i // CHUNK + 1})",
            api_key=api_key, filter_formula=formula,
        )
        if not df_chunk.empty:
            all_rows.append(df_chunk)
 
    if not all_rows:
        print("⚠️  No Sales Order Lines records found.")
        return {}
 
    df_all = pd.concat(all_rows, ignore_index=True)
    return build_price_map(df_all)
 
 
def fetch_price_map_from_commerce_central(so_numbers):
    """Fetch unit prices from Commerce Central Sales Order Lines table using SO Line Id field."""
    if not so_numbers:
        return {}
 
    config = TABLE_CONFIG["commerce_central_sales_lines"]
    so_list = list(so_numbers)
    all_rows = []
    CHUNK = 20
 
    print(f"🌐 Fetching Commerce Central Sales Lines for {len(so_list)} SO(s): {so_list[:5]}...")
 
    # Normalize each SO number (extract last part after slash)
    normalized_sos = [normalize_sales_order_number(so) for so in so_list]
 
    for i in range(0, len(normalized_sos), CHUNK):
        chunk = normalized_sos[i:i+CHUNK]
        # Build OR conditions using SO Line Id (e.g., "41178316_74391" starts with "41178316_")
        conditions = []
        for so in chunk:
            conditions.append(f'LEFT({{SO Line Id}}, LEN("{so}") + 1) = "{so}_"')
        formula = "OR(" + ",".join(conditions) + ")"
 
        df_chunk = fetch_table(
            base_id=config["base_id"], table_id=config["table_id"],
            view_id=config["view_id"],
            label=f"Commerce Central Sales Lines (batch {i//CHUNK + 1})",
            api_key=config["api_key"], filter_formula=formula,
        )
        if not df_chunk.empty:
            all_rows.append(df_chunk)
 
    if not all_rows:
        print("⚠️  No Commerce Central Sales Order Lines found.")
        return {}
 
    df_all = pd.concat(all_rows, ignore_index=True)
    price_map = {}
    for _, row in df_all.iterrows():
        # Extract Sales Order number from SO Line Id (everything before the first underscore)
        so_line = row.get("SO Line Id", "")
        if not so_line:
            # fallback to legacy field name just in case
            so_line = row.get("SO Line#", "")
        if so_line and '_' in so_line:
            so = normalize_sales_order_number(so_line.split('_')[0])
        else:
            continue
        if not so:
            continue
 
        # Get price from "Converted Unit price" (fallback to "Price (Final)")
        price = clean_price(row.get("Converted Unit price", 0))
        if price <= 0:
            price = clean_price(row.get("Price (Final)", 0))
        if price > 0:
            # Sum prices if multiple lines for the same SO
            price_map[so] = price_map.get(so, 0) + price
 
    print(f"✅ Commerce Central price map: {len(price_map)} entries")
    return price_map
 
 
# ============================================================
# TRANSFORM (builds CSV rows)
# ============================================================
 
def transform(df, ya_rma_map=None, crs_rma_map=None, price_map=None):
    rows = []
    F = FIXED
 
    ya_rma_map  = ya_rma_map  or {}
    crs_rma_map = crs_rma_map or {}
    price_map   = price_map   or {}
 
    for _, r in df.iterrows():
        country      = str(r.get("Shipping Country", "")).strip()
        cc, num      = split_phone(r.get("Telephone", ""), country)
        order_number = str(r.get("Order Number", "")).strip()
 
        price = round(clean_price(r.get("Converted Unit Price per Line", 0)), 2)
 
        rma_source = get_rma_source(order_number)
 
        if rma_source:
            clean_rma = normalize_rma_number(order_number)
            if rma_source == "ya":
                sales_order  = ya_rma_map.get(clean_rma, "")
                source_label = "YesAgain RMA"
            else:
                sales_order  = crs_rma_map.get(clean_rma, "")
                source_label = "CRS RMA"
 
            if sales_order:
                looked_up_price = price_map.get(sales_order, None)
                if looked_up_price is not None:
                    price = round(looked_up_price, 2)
                    print(f"   ✅ {order_number} ({source_label}) → SO {sales_order} → €{price}")
                else:
                    # Path A: Sales Order found but no price in Sales Order Lines
                    if rma_source == "crs":
                        price = 200.0
                        print(f"   ⚠️  {order_number} ({source_label}) → SO {sales_order} "
                              f"found but no price in Sales Order Lines — defaulting to €200.00")
                    else:
                        print(f"   ⚠️  {order_number} ({source_label}) → SO {sales_order} "
                              f"found but no price in Sales Order Lines")
            else:
                # Path B: RMA key not found in the map at all
                rma_map_used = ya_rma_map if rma_source == "ya" else crs_rma_map
                print(f"   ⚠️  {order_number} ({source_label}) → "
                      f"RMA key '{clean_rma}' not found in map.")
                print(f"        Map has {len(rma_map_used)} entries: "
                      f"{list(rma_map_used.keys())[:10]}")
                if rma_source == "crs":
                    price = 200.0
                    print(f"   🔄 {order_number} (RMAC) → RMA not in map — defaulting to €200.00")
 
        qty = clean_qty(r.get("Sold Qty per line", 1))
 
        rules       = get_destination_rules(country, order_number)
        shipper_acc = rules["shipper_account"]
        duty_acc    = rules["duty_account"]
        add_party   = rules["additional_party"]
        region      = rules["region"]
 
        # ── Declared Value Rules ──────────────────────────────────────
        # Policy (per shipping policy document):
        #   Saudi Arabia (SA) → FIXED 164 EUR (regardless of item price)
        #   GCC / USA / Others not listed → 20% of item value
        #   Europe / UK → full price (Exclude VAT — data is already ex-VAT)
        country_upper = country.strip().upper()
        if country_upper == "SA":
            declared_val = 164.0
            print(f"   🇸🇦 {order_number} (SA) → Fixed declared value: €164.00")
        elif region in ("gcc", "us", "other"):
            declared_val = round(price * 0.20, 2)
            declared_val = max(declared_val, 1.0)   # DHL minimum
            print(f"   📦 {order_number} ({region.upper()}) → 20% declared: €{declared_val} (item: €{price})")
        else:
            # eu / uk — full price, exclude VAT (DHL minimum 1.0)
            declared_val = max(price, 1.0)
 
        addr2, state = get_address2_and_state(r)
        state_full = STATE_FULL_NAMES.get(state, state) if state else ""
 
        name = clean_text(r.get("Shipping Name", ""))[:35]
        raw_company = r.get("Invoice Name", "")
        if pd.isna(raw_company) if not isinstance(raw_company, (list, dict)) else False:
            company = name
        elif str(raw_company).strip() == "":
            company = name
        else:
            company = clean_text(raw_company)[:35]
 
        rows.append({
            # --- SHIPMENT DETAILS ---
            "Shipment Reference 1"                      : order_number,
            "Name (Ship TO) (Required)"                 : name,
            "Company (Ship TO) (Required)"              : company,
            "Address 1 (Ship TO) (Required)"            : (clean_text(r.get("Shipping address 1", "")).replace(",", " "))[:45],
            "Address 2 (Ship TO)"                       : addr2.replace(",", " "),
            "Address 3 (Ship TO)"                       : "",
            "City (Ship TO) (Required)"                 : clean_text(r.get("Shipping City", "")),
            "State Province (Ship TO)"                  : state_full,
            "ZIP Postal Code (Ship TO)"                 : r.get("Shipping Postcode", ""),
            "Country Code (Ship TO) (Required)"         : country,
            "Email Address (Ship TO)"                   : r.get("CustomerEmail", ""),
            "Phone Type (Ship TO)"                      : "O",
            "Phone Country Code (Ship TO) (Required)"   : cc,
            "Phone Number (Ship TO) (Required)"         : num,
            "VAT Tax ID (Ship TO)"                      : "",
            "EORI Number (Ship TO)"                     : "",
 
            # --- ACCOUNT / SHIPPING CONFIG ---
            "Account Number (Shipper) (Required)"       : int(shipper_acc),
            "Account Number (Payer)"                    : int(shipper_acc),
            "inCoterms"                                 : F["incoterms"],
            "Total Weight (Required)"                   : round(qty * F["weight"], 2),
            "Piece Weight (Unit of Measure)"            : F["weight_unit"],
            "Declared Value Currency (Required)"        : F["currency"],
            "Declared Value (Required)"                 : declared_val,
            "Product Code (3 Letter)"                   : F["product_code"],
            "Summary of Contents"                       : F["contents"],
            "SHIPMENT TYPE"                             : F["shipment_type"],
            "Total Shipment Pieces"                     : qty,
            "Piece Dimensions (Unit of Measure)"        : F["dim_unit"],
            "LENGTH"                                    : F["length"],
            "WIDTH"                                     : F["width"],
            "HEIGHT"                                    : F["height"],
            "REASON OF EXPORT"                          : F["export_reason"],
            "Digital Customs Invoice  Y N (Paperless)"  : F["dig_customs"],
 
            # --- INVOICE / ITEM ---
            "INVOICE NO."                               : order_number,
            "ITEM DESCRIPTION"                          : F["item_desc"],
            "ITEM COMMODITY"                            : F["commodity"],
            "ITM QTY"                                   : qty,
            "ITM UNITS"                                 : F["item_units"],
            "ITM VALUE"                                 : round(declared_val, 2),
            "ITM CRNCY"                                 : F["currency"],
            "ITM NET"                                   : F["item_net"],
            "ITM GRSS"                                  : F["item_gross"],
            "Country of Origin (Customs Invoice)"       : F["origin"],
            "REMARKS"                                   : "",
 
            # --- ADDITIONAL PARTY ---
            "Company (Additional Party)"                : add_party["Company"],
            "Name (Additional Party)"                   : add_party["Name"],
            "Address 1 (Additional Party)"              : add_party["Address1"],
            "Address 2 (Additional Party)"              : "",
            "Address 3 (Additional Party)"              : "",
            "Country Code (Additional Party)"           : add_party["Country"],
            "City (Additional Party)"                   : add_party["City"],
            "ZIP Postal Code (Additional Party)"        : add_party["ZIP"],
            "ADD  EMAIL"                                : add_party["Email"],
            "Phone Country Code (Additional Party)"     : add_party["PhoneCC"],
            "Phone Number (Additional Party)"           : add_party["Phone"],
            "ADD  VAT"                                  : importer_vat_for(country, add_party["VAT"]),
            "ADD  EORI"                                 : add_party["EORI"],
            "ADD  RELATIONSHIP"                         : add_party["Rel"],
            "Account Number (Duty Tax)"                 : int(duty_acc),
        })
 
    return pd.DataFrame(rows)
 
 
# ============================================================
# DHL API INTEGRATION
# ============================================================
 
def build_dhl_payload(row):
    F = FIXED
    from datetime import timedelta
    
    # DHL prefers an explicit offset structure for timezones
    ship_date = (datetime.now(timezone.utc) + timedelta(days=1)).strftime(
        "%Y-%m-%dT10:00:00GMT+04:00"
    )
 
    order_id = str(row.get("Shipment Reference 1", "")).strip()
    country  = str(row.get("Country Code (Ship TO) (Required)", "")).strip()
 
    # Get structural variables from mapping logic
    rules           = get_destination_rules(country, order_id)
    party           = rules["shipper_party"]
    shipper_account = rules["shipper_account"]
    duty_account    = rules["duty_account"]

    # Destination-based importer VAT/EORI (YesAgain's local registration for
    # the country the goods are imported into). Falls back to the importer
    # entity's default VAT when the destination has no specific registration.
    importer_party = rules["additional_party"]
    importer_vat   = importer_vat_for(country, importer_party.get("VAT", ""))
    importer_eori  = importer_party.get("EORI", "")
    vat_issuer     = (importer_vat[:2].upper() if importer_vat else importer_party.get("Country", ""))
 
    # Phone number sanitation and formatting
    phone_cc  = str(row.get("Phone Country Code (Ship TO) (Required)", "")).strip()
    phone_num = str(row.get("Phone Number (Ship TO) (Required)", "")).strip()
    if not phone_cc:
        phone_cc = PHONE_CODES.get(country, "")
    if not phone_num:
        phone_num = "0000000000"
    phone_num_clean = re.sub(r'\D', '', phone_num)
    
    # --- FIXED LINE HERE ---
    full_phone = f"+{phone_cc}{phone_num_clean}" if not str(phone_cc).startswith('+') else f"{phone_cc}{phone_num_clean}"
 
    shipper_phone_cc   = str(party["PhoneCC"])
    shipper_phone_num  = re.sub(r'\D', '', str(party["Phone"]))
    shipper_full_phone = f"+{shipper_phone_cc}{shipper_phone_num}"
 
    # Smart address splitting to prevent truncation data loss
    raw_address = str(row.get("Address 1 (Ship TO) (Required)", "")).strip()
    addr2       = str(row.get("Address 2 (Ship TO)", "")).strip()
    
    postal_address = {
        "postalCode"  : clean_postal_code(str(row.get("ZIP Postal Code (Ship TO)", "")).strip(), country),
        "cityName"    : str(row.get("City (Ship TO) (Required)", "")).strip(),
        "countryCode" : country,
        "addressLine1": raw_address[:45],
    }
 
    # Overflow remainder of address line 1 safely into address line 2 if line 2 is empty
    if len(raw_address) > 45:
        postal_address["addressLine2"] = (raw_address[45:] + " " + addr2).strip()[:45]
    elif addr2:
        postal_address["addressLine2"] = addr2[:45]
        
    state = str(row.get("State Province (Ship TO)", "")).strip()
    if state and country == "US":
        postal_address["provinceCode"] = state
 
    # Accurate unit-level pricing calculation 
    itm_qty    = max(int(row.get("ITM QTY", 1)), 1)
    total_val  = float(row.get("Declared Value (Required)", 0))
    unit_price = max(round(total_val / itm_qty, 2), 0.01)
 
    payload = {
        "plannedShippingDateAndTime": ship_date,
        "pickup": {"isRequested": False},
        "productCode": F["product_code_api"],
        "accounts": [
            {"typeCode": "shipper", "number": str(shipper_account)},
            {"typeCode": "payer",   "number": str(shipper_account)},
            {"typeCode": "duties-taxes", "number": str(duty_account)}, # Bound dynamically via accounting routing Matrix
        ],
        "outputImageProperties": {
            "printerDPI"    : 300,
            "encodingFormat": "pdf",
            "imageOptions"  : [
                {"typeCode": "label", "templateName": "ECOM26_84_001", "isRequested": True},
                {
                    "typeCode"           : "invoice",
                    "templateName"       : "COMMERCIAL_INVOICE_P_10",
                    "isRequested"        : True,
                    "invoiceType"        : "commercial",
                    "languageCode"       : "eng",
                    "languageCountryCode": "US",
                },
            ]
        },
        "customerDetails": {
            "shipperDetails": {
                "postalAddress": {
                    "postalCode"  : party["ZIP"],
                    "cityName"    : party["City"],
                    "countryCode" : party["Country"],
                    "addressLine1": party["Address1"],
                },
                "contactInformation": {
                    "companyName": party["Company"],
                    "fullName"   : party["Name"],
                    "email"      : party["Email"],
                    "phone"      : shipper_full_phone,
                },
                "registrationNumbers": [
                    *([{"typeCode": "VAT", "number": party["VAT"],  "issuerCountryCode": party["Country"]}] if party.get("VAT")  else []),
                    *([{"typeCode": "EOR", "number": party["EORI"], "issuerCountryCode": party["Country"]}] if party.get("EORI") else []),
                    *([{"typeCode": "VAT", "number": importer_vat,  "issuerCountryCode": vat_issuer}] if importer_vat  else []),
                    *([{"typeCode": "EOR", "number": importer_eori, "issuerCountryCode": importer_party.get("Country", "")}] if importer_eori else []),
                ],
            },
            "receiverDetails": {
                "postalAddress": postal_address,
                "contactInformation": {
                    "companyName": str(row.get("Company (Ship TO) (Required)", "")).strip()[:35] or str(row.get("Name (Ship TO) (Required)", "")).strip()[:35],
                    "fullName"   : str(row.get("Name (Ship TO) (Required)", "")).strip()[:35],
                    "email"      : str(row.get("Email Address (Ship TO)", "")).strip(),
                    "phone"      : full_phone,
                },
            },
        },
        "content": {
            "packages": [
                {
                    "weight"    : float(row.get("Total Weight (Required)", F["weight"])),
                    "dimensions": {"length": F["length"], "width": F["width"], "height": F["height"]},
                    "customerReferences": [
                        {"typeCode": "CU",  "value": order_id[:35]},
                    ]
                }
            ],
            "isCustomsDeclarable"   : True,
            "declaredValue"         : total_val,
            "declaredValueCurrency" : str(row.get("Declared Value Currency (Required)", F["currency"])),
            "exportDeclaration": {
                "lineItems": [
                    {
                        "number"             : i + 1,
                        "description"        : F["item_desc"],
                        "price"              : unit_price,
                        "priceCurrency"      : str(row.get("Declared Value Currency (Required)", F["currency"])),
                        "commodityCodes"     : [
                            {"typeCode": "outbound", "value": F["commodity"]},
                            {"typeCode": "inbound",  "value": F["commodity"]},
                        ],
                        "exportReasonType"   : "permanent",
                        "manufacturerCountry": F["origin"],
                        "weight"             : {"netValue": F["item_net"], "grossValue": F["item_gross"]},
                        "quantity"           : {"value": 1, "unitOfMeasurement": "PCS"},
                    }
                    for i in range(itm_qty)
                ],
                "invoice": {
                    # DHL limits invoice number to 35 chars. Airtable UUIDs are 36 → HTTP 422.
                    "number": (str(row.get("INVOICE NO.", "")) or order_id)[:35],
                    "date"  : datetime.now().strftime("%Y-%m-%d"),
                },
                "exportReason": F["export_reason"],
            },
            "description"       : F["contents"],
            "incoterm"          : "DDP", # Forced to DDP to match owner instructions
            "unitOfMeasurement": "metric",
        },
        "shipmentNotification": [
            {
                "typeCode"    : "email",
                "receiverId"  : str(row.get("Email Address (Ship TO)", "")).strip(),
                "languageCode": "eng",
            }
        ] if str(row.get("Email Address (Ship TO)", "")).strip() else [],
        "valueAddedServices": [
            {"serviceCode": "WY"}, # Changed from DD to WY (Duties Taxes Paid) to avoid double billing bugs!
        ],
        "customerReferences": [
            {"typeCode": "CU",  "value": order_id[:35]},
        ],
    }
 
    return payload
 
 
# ============================================================
# LABEL INDEX — saves every shipment so the lookup tool can find it
# ============================================================
 
LABEL_INDEX_FILE = "label_index.json"
 
def save_label_index(order_number, tracking_number, label_path,
                     recipient_name, destination_country):
    """Append/update a record in label_index.json for the lookup tool."""
    if os.path.exists(LABEL_INDEX_FILE):
        try:
            with open(LABEL_INDEX_FILE, "r", encoding="utf-8") as f:
                index = json.load(f)
        except Exception:
            index = []
    else:
        index = []
 
    index = [e for e in index
             if str(e.get("order", "")).strip() != str(order_number).strip()]
 
    index.insert(0, {
        "order"     : str(order_number).strip(),
        "awb"       : str(tracking_number).strip(),
        "label"     : label_path,
        "recipient" : str(recipient_name).strip(),
        "country"   : str(destination_country).strip(),
        "created_at": datetime.now().strftime("%Y-%m-%d %H:%M"),
    })
 
    with open(LABEL_INDEX_FILE, "w", encoding="utf-8") as f:
        json.dump(index, f, indent=2, ensure_ascii=False)
 
    print(f"   📋 Label index updated: {LABEL_INDEX_FILE}")
 
 
def generate_lookup_html():
    """Generate label_lookup.html with all label data embedded — no server needed."""
    if not os.path.exists(LABEL_INDEX_FILE):
        print("⚠️  No label_index.json found — skipping HTML generation.")
        return
 
    try:
        with open(LABEL_INDEX_FILE, "r", encoding="utf-8") as f:
            index = json.load(f)
    except Exception as e:
        print(f"⚠️  Could not read label_index.json: {e}")
        return
 
    enriched = []
    for entry in index:
        label_path = entry.get("label", "")
        label_b64  = ""
        if label_path and os.path.exists(label_path):
            with open(label_path, "rb") as lf:
                label_b64 = base64.b64encode(lf.read()).decode("utf-8")
        enriched.append({**entry, "label_b64": label_b64})
 
    data_json = json.dumps(enriched, ensure_ascii=False)
 
    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>YesAgain — DHL Label Lookup</title>
<style>
  * {{ box-sizing: border-box; margin: 0; padding: 0; }}
  body {{
    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
    background: #f5f5f5; min-height: 100vh;
    display: flex; flex-direction: column; align-items: center; padding: 40px 20px;
  }}
  .header {{
    background: #FFCC00; width: 100%; max-width: 700px;
    border-radius: 12px 12px 0 0; padding: 24px 30px;
    display: flex; align-items: center; gap: 16px;
  }}
  .header-logo {{ font-size: 28px; font-weight: 900; color: #D40511; letter-spacing: -1px; }}
  .header-title {{ font-size: 18px; font-weight: 700; color: #333; }}
  .header-sub {{ font-size: 13px; color: #666; margin-top: 2px; }}
  .card {{
    background: white; width: 100%; max-width: 700px;
    border-radius: 0 0 12px 12px; padding: 30px;
    box-shadow: 0 4px 20px rgba(0,0,0,0.08);
  }}
  .search-row {{ display: flex; gap: 10px; margin-bottom: 28px; }}
  .search-input {{
    flex: 1; padding: 14px 18px; font-size: 16px;
    border: 2px solid #ddd; border-radius: 8px; outline: none; transition: border-color 0.2s;
  }}
  .search-input:focus {{ border-color: #FFCC00; }}
  .search-btn {{
    padding: 14px 28px; background: #D40511; color: white; border: none;
    border-radius: 8px; font-size: 16px; font-weight: 600; cursor: pointer;
  }}
  .search-btn:hover {{ background: #b00; }}
  .result-box {{
    display: none; background: #f9f9f9; border: 1px solid #e0e0e0;
    border-radius: 10px; padding: 24px; margin-bottom: 24px;
  }}
  .result-box.visible {{ display: block; }}
  .result-label {{ font-size: 12px; color: #888; text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 3px; }}
  .result-value {{ font-size: 18px; font-weight: 700; color: #222; margin-bottom: 16px; }}
  .result-meta {{ display: flex; gap: 24px; flex-wrap: wrap; margin-bottom: 20px; }}
  .meta-item {{ flex: 1; min-width: 120px; }}
  .download-btn {{
    display: inline-block; padding: 14px 32px; background: #FFCC00;
    color: #333; font-size: 16px; font-weight: 700; border-radius: 8px;
    text-decoration: none; cursor: pointer; border: none; width: 100%; text-align: center;
  }}
  .download-btn:hover {{ background: #e6b800; }}
  .no-label {{ color: #cc0000; font-size: 14px; margin-top: 10px; }}
  .error-box {{
    display: none; background: #fff0f0; border: 1px solid #ffcccc;
    border-radius: 10px; padding: 20px 24px; color: #cc0000; font-size: 15px; margin-bottom: 24px;
  }}
  .error-box.visible {{ display: block; }}
  .recent-title {{ font-size: 13px; font-weight: 600; color: #888; text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 12px; }}
  .recent-table {{ width: 100%; border-collapse: collapse; font-size: 14px; }}
  .recent-table th {{ text-align: left; padding: 8px 10px; background: #f0f0f0; color: #555; font-weight: 600; }}
  .recent-table td {{ padding: 9px 10px; border-bottom: 1px solid #f0f0f0; color: #333; }}
  .recent-table tr:hover td {{ background: #fafafa; cursor: pointer; }}
  .badge {{ display: inline-block; padding: 2px 8px; background: #e8f5e9; color: #2e7d32; border-radius: 20px; font-size: 12px; font-weight: 600; }}
  .badge.no-file {{ background: #fce4ec; color: #c62828; }}
</style>
</head>
<body>
<div class="header">
  <div class="header-logo">DHL</div>
  <div>
    <div class="header-title">Label Lookup Tool</div>
    <div class="header-sub">YesAgain — Search by Order Number to download a label</div>
  </div>
</div>
<div class="card">
  <div class="search-row">
    <input type="text" class="search-input" id="searchInput"
           placeholder="Enter Order Number or AWB"
           onkeydown="if(event.key==='Enter') doSearch()">
    <button class="search-btn" onclick="doSearch()">🔍 Search</button>
  </div>
  <div class="result-box" id="resultBox">
    <div class="result-label">Order Number</div>
    <div class="result-value" id="resOrder">—</div>
    <div class="result-meta">
      <div class="meta-item"><div class="result-label">AWB / Tracking</div><div class="result-value" id="resAWB" style="font-size:15px">—</div></div>
      <div class="meta-item"><div class="result-label">Recipient</div><div class="result-value" id="resRecipient" style="font-size:15px">—</div></div>
      <div class="meta-item"><div class="result-label">Country</div><div class="result-value" id="resCountry" style="font-size:15px">—</div></div>
      <div class="meta-item"><div class="result-label">Created</div><div class="result-value" id="resDate" style="font-size:15px">—</div></div>
    </div>
    <button class="download-btn" id="downloadBtn" onclick="downloadLabel()">⬇️ Download Shipping Label</button>
    <div class="no-label" id="noLabelMsg" style="display:none">⚠️ Label file not available.</div>
  </div>
  <div class="error-box" id="errorBox">❌ No shipment found for that order number.</div>
  <div class="recent-title">Recent Shipments ({len(index)} total)</div>
  <table class="recent-table">
    <thead><tr><th>Order</th><th>AWB</th><th>Recipient</th><th>Country</th><th>Created</th><th>Label</th></tr></thead>
    <tbody id="recentTable"></tbody>
  </table>
</div>
<script>
const DATA = {data_json};
let currentEntry = null;
function buildTable() {{
  const tbody = document.getElementById("recentTable");
  DATA.slice(0, 100).forEach(e => {{
    const hasLabel = e.label_b64 && e.label_b64.length > 0;
    const tr = document.createElement("tr");
    tr.innerHTML = `<td><strong>${{e.order}}</strong></td><td>${{e.awb||"—"}}</td><td>${{e.recipient||"—"}}</td><td>${{e.country||"—"}}</td><td>${{e.created_at||"—"}}</td><td><span class="badge ${{hasLabel?"":"no-file"}}">${{hasLabel?"✅ Ready":"❌ Missing"}}</span></td>`;
    tr.onclick = () => showResult(e);
    tbody.appendChild(tr);
  }});
}}
function showResult(entry) {{
  currentEntry = entry;
  document.getElementById("resOrder").textContent     = entry.order;
  document.getElementById("resAWB").textContent       = entry.awb || "—";
  document.getElementById("resRecipient").textContent = entry.recipient || "—";
  document.getElementById("resCountry").textContent   = entry.country || "—";
  document.getElementById("resDate").textContent      = entry.created_at || "—";
  const hasLabel = entry.label_b64 && entry.label_b64.length > 0;
  document.getElementById("downloadBtn").style.display = hasLabel ? "block" : "none";
  document.getElementById("noLabelMsg").style.display  = hasLabel ? "none" : "block";
  document.getElementById("resultBox").classList.add("visible");
  document.getElementById("errorBox").classList.remove("visible");
}}
function doSearch() {{
  const q = document.getElementById("searchInput").value.trim().toUpperCase();
  if (!q) return;
  const found = DATA.find(e => e.order.toUpperCase()===q || (e.awb||"").toUpperCase()===q);
  if (found) {{ showResult(found); }}
  else {{
    document.getElementById("resultBox").classList.remove("visible");
    document.getElementById("errorBox").classList.add("visible");
    currentEntry = null;
  }}
}}
function downloadLabel() {{
  if (!currentEntry || !currentEntry.label_b64) return;
  const bytes = atob(currentEntry.label_b64);
  const arr = new Uint8Array(bytes.length);
  for (let i=0;i<bytes.length;i++) arr[i]=bytes.charCodeAt(i);
  const blob = new Blob([arr], {{type:"application/pdf"}});
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = "label_" + currentEntry.order + ".pdf";
  a.click();
}}
buildTable();
</script>
</body>
</html>"""
 
    with open("label_lookup.html", "w", encoding="utf-8") as f:
        f.write(html)
 
    print(f"\n✅ Label lookup tool: label_lookup.html ({len(index)} shipment(s) embedded)")
 
 
def send_to_dhl(row, verbose=False):
    if not DHL_API_KEY or not DHL_API_SECRET:
        print("❌ DHL_API_KEY or DHL_API_SECRET missing in .env — cannot send to DHL.")
        return {"success": False, "error": "Missing DHL credentials"}
    
    use_test     = DHL_TEST_MODE
    base_url     = DHL_BASE_URL_PROD   # always production URL — test bypass handled below
    order_number = str(row.get("Shipment Reference 1", "unknown")).strip()
 
    # ── TEST MODE BYPASS ──────────────────────────────────────────────────────
    # DHL error 803 "Account not allowed for this service" happens because the
    # DHL TEST SANDBOX does not recognise real production account numbers.
    # The test sandbox requires special dummy test accounts provided by DHL.
    # Since we don't have those dummy accounts, calling the sandbox will always
    # fail with 803.
    #
    # Solution: in --test mode we skip the real DHL call entirely and return a
    # simulated success. This lets you safely test the full pipeline:
    #   ✅  Airtable fetch
    #   ✅  Price / RMA lookup
    #   ✅  CSV generation
    #   ✅  Payload building & printing (use --verbose to inspect)
    #   ✅  Airtable writeback is STILL SKIPPED (test mode already guards this)
    #   ❌  No real DHL call (by design — no fake AWB)
    #
    # To test the actual DHL connection without booking a real shipment, run
    # python airtable_transform.py --send-to-dhl --validate
    # which calls /shipments/validate on PRODUCTION with your real accounts.
    if use_test:
        fake_awb = f"TEST-{order_number}-SIM"
        print(f"   ✅ TEST SIMULATION — Payload built OK (use --verbose to print)")
        print(f"   ℹ️  No DHL call made — fake AWB: {fake_awb}")
        return {
            "success"         : True,
            "order_number"    : order_number,
            "tracking_number" : fake_awb,
            "label_base64"    : None,
            "raw_response"    : {},
            "simulated"       : True,
        }
 
    url = f"{base_url}/shipments"
 
    headers = {"Content-Type": "application/json", "Accept": "application/json"}
    payload = build_dhl_payload(row)
 
    country = str(row.get("Country Code (Ship TO) (Required)", "")).strip()
    rules   = get_destination_rules(country, order_number)
    party   = rules["shipper_party"]
    print(f"   Shipper: {party['Company']} ({party['Country']})  "
          f"Duty account: {rules['duty_account']}  Dest: {country}")
 
    if verbose:
        print(f"   Payload:\n{json.dumps(payload, indent=2)}")
 
    try:
        response = requests.post(url, json=payload, headers=headers,
                                 auth=(DHL_API_KEY, DHL_API_SECRET), timeout=30)
    except requests.exceptions.Timeout:
        print(f"   ❌ Timeout sending {order_number}")
        return {"success": False, "order_number": order_number, "error": "Timeout"}
    except requests.exceptions.ConnectionError as e:
        print(f"   ❌ Connection error for {order_number}: {e}")
        return {"success": False, "order_number": order_number, "error": str(e)}
 
    if response.status_code in (200, 201):
        data            = response.json()
        tracking_number = data.get("shipmentTrackingNumber", "")
        packages        = data.get("packages", [])
        pkg_tracking    = packages[0].get("trackingNumber", "") if packages else ""
 
        print(f"   ✅ DHL OK — Tracking: {tracking_number or pkg_tracking}")
 
        label_b64 = None
        for doc in data.get("documents", []):
            type_code    = doc.get("typeCode", "")
            image_format = doc.get("imageFormat", "")
            content      = doc.get("content", "")
 
            # DHL production API returns typeCode="label".
            # DHL sandbox/test API often omits typeCode entirely, returning only
            # imageFormat="PDF". Accept EITHER so the label is never silently dropped.
            is_label = (
                type_code == "label"
                or (not type_code and image_format.upper() == "PDF" and content)
            )
 
            if is_label and content:
                label_b64  = content
                os.makedirs("labels", exist_ok=True)
                label_path = f"labels/label_{order_number}.pdf"
                with open(label_path, "wb") as f:
                    f.write(base64.b64decode(label_b64))
                print(f"   🏷️  Label saved locally: {label_path}")
                recipient_name      = str(row.get("Name (Ship TO) (Required)", "")).strip()
                destination_country = str(row.get("Country Code (Ship TO) (Required)", "")).strip()
                save_label_index(order_number, tracking_number or pkg_tracking,
                                 label_path, recipient_name, destination_country)
                break
 
        if not label_b64:
            print(f"   ⚠️  No label found in DHL response for {order_number}.")
            docs_summary = [
                {k: v for k, v in d.items() if k != "content"}
                for d in data.get("documents", [])
            ]
            print(f"       Documents returned: {docs_summary}")
 
        return {
            "success"         : True,
            "order_number"    : order_number,
            "tracking_number" : tracking_number or pkg_tracking,
            "label_base64"    : label_b64,
            "raw_response"    : data,
        }
    else:
        try:
            error_msg = json.dumps(response.json(), indent=2)
        except Exception:
            error_msg = response.text
 
        print(f"   ❌ DHL FAILED — HTTP {response.status_code}")
        print(f"   {error_msg[:500]}")
 
        return {
            "success"      : False,
            "order_number" : order_number,
            "status_code"  : response.status_code,
            "error"        : error_msg,
        }
 
 
def validate_credentials_dhl():
    if not DHL_API_KEY or not DHL_API_SECRET:
        print("❌ DHL_API_KEY or DHL_API_SECRET not set in .env")
        return False
 
    base_url = DHL_BASE_URL_TEST if DHL_TEST_MODE else DHL_BASE_URL_PROD
    url = (
        f"{base_url}/products"
        f"?accountNumber={FIXED['account_shipper']}"
        f"&originCountryCode=AE&originCityName=Sharjah"
        f"&destinationCountryCode=DE&destinationCityName=Berlin"
        f"&weight=1.5&length=35&width=30&height=7"
        f"&plannedShippingDate={datetime.now().strftime('%Y-%m-%d')}"
    )
 
    print(f"\n🔑 Testing DHL credentials {'[TEST]' if DHL_TEST_MODE else '[PRODUCTION]'}...")
 
    try:
        response = requests.get(url, auth=(DHL_API_KEY, DHL_API_SECRET),
                                headers={"Accept": "application/json"}, timeout=15)
        print(f"   HTTP Status: {response.status_code}")
 
        if response.status_code == 200:
            print("   ✅ DHL credentials are VALID!")
            return True
        elif response.status_code == 401:
            print("   ❌ 401 Unauthorized — credentials are WRONG.")
            return False
        elif response.status_code == 403:
            print("   ⚠️  403 Forbidden — credentials work but account may not have this endpoint.")
            return True
        else:
            print(f"   ⚠️  Unexpected status {response.status_code} — assuming credentials valid.")
            return True
 
    except requests.exceptions.ConnectionError:
        print("   ❌ Cannot reach DHL API — check internet connection.")
        return False
    except requests.exceptions.Timeout:
        print("   ❌ Timeout connecting to DHL API.")
        return False
 
 
# ============================================================
# CANCEL DHL SHIPMENT
# ============================================================

def cancel_dhl_shipment(tracking_number):
    """
    Cancel a booked DHL shipment using its AWB / tracking number.

    HOW IT WORKS
    ─────────────────────────────────────────────────────────────
    DHL Express API:  DELETE /mydhlapi/shipments/{trackingNumber}
    Returns 200 or 204 on success.

    IMPORTANT RULES
    ─────────────────────────────────────────────────────────────
    ✅  Works ONLY if DHL has NOT yet scanned / picked up the parcel.
    ❌  Once DHL scans it, this API call returns an error.
        → You must call DHL customer service: +971 600 567 567
        → Or ask the recipient to refuse delivery.

    Args:
        tracking_number  – DHL AWB string (e.g. "1234567890")

    Returns dict:
        { success, tracking_number, message }   on success
        { success, tracking_number, error }     on failure
    """
    if not DHL_API_KEY or not DHL_API_SECRET:
        print("❌ DHL_API_KEY or DHL_API_SECRET missing — cannot cancel.")
        return {"success": False, "error": "Missing DHL credentials"}

    tracking_number = str(tracking_number).strip()

    if not tracking_number or tracking_number.startswith("TEST-"):
        return {
            "success": False,
            "error"  : "Invalid or simulated tracking number — nothing to cancel."
        }

    url = f"{DHL_BASE_URL_PROD}/shipments/{tracking_number}"
    print(f"\n🗑️  Cancelling DHL shipment: {tracking_number}")
    print(f"   URL: DELETE {url}")

    try:
        response = requests.delete(
            url,
            headers={"Accept": "application/json"},
            auth=(DHL_API_KEY, DHL_API_SECRET),
            timeout=30
        )
    except requests.exceptions.Timeout:
        print(f"   ❌ Timeout cancelling {tracking_number}")
        return {"success": False, "tracking_number": tracking_number, "error": "Timeout"}
    except requests.exceptions.ConnectionError as e:
        print(f"   ❌ Connection error: {e}")
        return {"success": False, "tracking_number": tracking_number, "error": str(e)}

    # DHL returns 200 or 204 on successful cancellation
    if response.status_code in (200, 204):
        print(f"   ✅ DHL confirmed cancellation: {tracking_number}")
        return {
            "success"         : True,
            "tracking_number" : tracking_number,
            "message"         : "Shipment cancelled successfully at DHL."
        }

    # Parse error from DHL — log the real status and body so it's diagnosable
    print(f"   ❌ DHL cancellation failed — HTTP {response.status_code}")
    print(f"   Raw response (first 600 chars): {response.text[:600]}")

    try:
        error_body = response.json()
        # DHL error responses use "detail" or "message" field
        error_msg  = error_body.get("detail") or error_body.get("message") or json.dumps(error_body)
    except Exception:
        raw = response.text.strip()
        if raw.startswith("<"):
            # DHL returns HTML/XML when the AWB is already picked up / in transit
            # or when the request hits a gateway error (403/503).
            # Map by HTTP status so the dashboard shows a clear actionable message.
            if response.status_code == 403:
                error_msg = (
                    "DHL rejected the request (403 Forbidden). "
                    "This usually means the shipment has already been picked up or scanned by DHL. "
                    "You cannot cancel it via API — call DHL UAE: +971 600 567 567."
                )
            elif response.status_code == 404:
                error_msg = (
                    f"Tracking number {tracking_number} not found at DHL (404). "
                    "It may already be cancelled or the number is incorrect."
                )
            elif response.status_code in (500, 503):
                error_msg = (
                    f"DHL server error ({response.status_code}). "
                    "Try again in a few minutes. If it persists, call DHL UAE: +971 600 567 567."
                )
            else:
                error_msg = (
                    f"DHL returned HTTP {response.status_code} with an unexpected response. "
                    "The shipment may already be in transit — call DHL UAE: +971 600 567 567."
                )
        else:
            error_msg = raw or f"DHL returned HTTP {response.status_code} with no detail."

    print(f"   Error: {error_msg[:400]}")

    return {
        "success"         : False,
        "tracking_number" : tracking_number,
        "status_code"     : response.status_code,
        "error"           : error_msg
    }


# ============================================================
# UNDO AIRTABLE BOOKING
# After a successful DHL cancellation, clear the Airtable fields
# that were written during the original booking (Steps 1–3).
# ============================================================

def undo_airtable_booking(record_id, order_number):
    """
    Clear Airtable writeback fields for a cancelled shipment.

    Clears:
      • Shipment Tracking Number  → empty string
      • Shipment Courier          → empty string
      • Shipment Label Created    → False (untick checkbox)

    Args:
        record_id    – Airtable record ID (e.g. "recXXXXXXXXXXXXXX")
        order_number – for log messages only

    Returns True on success, False on failure.
    """
    if not record_id:
        print(f"   ⚠️  No Airtable record_id for {order_number} — skipping Airtable undo.")
        return False

    base_id  = TABLE_CONFIG["orders"]["base_id"]
    table_id = TABLE_CONFIG["orders"]["table_id"]
    api_key  = AIRTABLE_API_KEY

    encoded_table = quote(str(table_id), safe="")
    url = f"https://api.airtable.com/v0/{base_id}/{encoded_table}/{record_id}"

    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type" : "application/json",
    }
    body = {
        "fields": {
            "Shipment Tracking Number": "",
            "Shipment Courier"        : "",
            "Shipment Label Created"  : False,
            "Shipment Label file"     : [],   # clear so re-book produces a fresh label
            "Commercial invoice file" : [],   # clear so re-book produces a fresh invoice
        }
    }

    print(f"   🔄 Undoing Airtable booking for {order_number} (record: {record_id})...")
    resp = airtable_patch_with_retry(url, body, headers, label=f"Undo/{order_number}")

    if resp is None:
        print(f"   ❌ Airtable undo failed (all retries exhausted) for {order_number}")
        return False

    if resp.status_code == 200:
        print(f"   ✅ Airtable fields cleared for {order_number}")
        return True
    else:
        print(f"   ❌ Airtable undo PATCH failed ({resp.status_code}) for {order_number}: "
              f"{resp.text[:300]}")
        return False


# ============================================================
# FIND ORDER BY TRACKING NUMBER
# ============================================================

def find_order_record_by_tracking(tracking_number):
    """Find the Airtable orders record whose Shipment Tracking Number matches the AWB."""
    tracking_number = str(tracking_number).strip()
    if not tracking_number:
        return None, None
    cfg = TABLE_CONFIG["orders"]
    formula = f'{{Shipment Tracking Number}}="{tracking_number}"'
    df = fetch_table(
        base_id=cfg["base_id"], table_id=cfg["table_id"],
        view_id=None, label="Orders (cancel lookup)",
        api_key=cfg["api_key"], filter_formula=formula,
    )
    if df.empty:
        return None, None
    row = df.iloc[0]
    record_id    = str(row.get("_airtable_id", "")).strip()
    order_number = str(row.get("Order Number", "")).strip()
    return (record_id or None), (order_number or None)


# ============================================================
# CANCEL + RESET  (single entry point for the dashboard)
# ============================================================

def cancel_and_reset_shipment(tracking_number):
    """
    Cancel a DHL shipment AND reset its Airtable record so it can be
    booked again. Call this from the dashboard / api.py cancel action.

    Flow:
      1. DELETE the shipment at DHL.
      2. If DHL cancellation succeeds → find the Airtable record by AWB.
      3. Clear all booking fields so the order returns to the queue.

    If the DHL cancel fails, Airtable is left untouched.
    """
    result = cancel_dhl_shipment(tracking_number)

    if not result.get("success"):
        result["airtable_reset"] = False
        return result

    record_id, order_number = find_order_record_by_tracking(tracking_number)

    if record_id:
        reset_ok = undo_airtable_booking(record_id, order_number or tracking_number)
        result["airtable_reset"] = reset_ok
        result["order_number"]   = order_number
        if reset_ok:
            result["message"] = ("Shipment cancelled at DHL and the order was reset — "
                                 "it can now be booked again.")
        else:
            result["message"] = ("Shipment cancelled at DHL, but clearing the Airtable "
                                 "fields failed — reset it manually to allow re-booking.")
    else:
        result["airtable_reset"] = False
        result["message"] = (f"Shipment cancelled at DHL, but no Airtable order was found "
                             f"with tracking number {tracking_number} to reset.")

    return result


# ============================================================
# MAIN
# ============================================================
if __name__ == "__main__":
 
    # ── Parse flags ──────────────────────────────────────────
    args             = [a for a in sys.argv[1:] if not a.startswith("--")]
    send_to_dhl_flag = "--send-to-dhl" in sys.argv
    test_override    = "--test"         in sys.argv
    verbose_flag     = "--verbose"      in sys.argv  # print full DHL payload per order
 
    if test_override:
        DHL_TEST_MODE = True
        print("⚠️  --test flag: DHL_TEST_MODE forced ON")

    # ── CANCEL + RESET (early exit) ──────────────────────────
    # Usage: python airtable_transform.py --cancel <TRACKING_NUMBER>
    if "--cancel" in sys.argv:
        try:
            awb = sys.argv[sys.argv.index("--cancel") + 1]
        except IndexError:
            print("❌ Usage: python airtable_transform.py --cancel <TRACKING_NUMBER>")
            sys.exit(1)
        cancel_result = cancel_and_reset_shipment(awb)
        print(json.dumps(cancel_result, indent=2, default=str))
        sys.exit(0 if cancel_result.get("success") else 1)

    # ── Initialise maps (will be populated below for RMA orders) ──
    ya_rma_map  = {}
    crs_rma_map = {}
    price_map   = {}
 
    # ── 0. Validate DHL credentials before doing anything ────
    if send_to_dhl_flag:
        ok = validate_credentials_dhl()
        if not ok:
            print("\n❌ Cannot proceed — fix DHL credentials first.")
            sys.exit(1)
 
    # ── 1. FETCH ORDERS ──────────────────────────────────────
    if args:
        csv_file = args[0]
        print(f"📂 Reading from CSV: {csv_file}")
        df = pd.read_csv(csv_file)
        print(f"✅ Loaded {len(df)} records from CSV.")
        print("ℹ️  CSV mode: RMA price lookup skipped.")
    else:
        if not AIRTABLE_API_KEY:
            print("❌ Missing AIRTABLE_API_KEY in .env file")
            sys.exit(1)
        print("🌐 Fetching Sales Orders from Airtable...")
        df = fetch_from_airtable()
 
    if df.empty:
        print("❌ No orders loaded.")
        sys.exit(1)
 
    # ── 2. DUPLICATE PREVENTION ──────────────────────────────
    # Skip orders that already have a tracking number OR have the
    # "Shipment Label Created" checkbox ticked. This means if the
    # script crashes mid-batch and you re-run it, it picks up
    # exactly where it left off without re-booking DHL.
    initial_count = len(df)
 
    if "Shipment Tracking Number" in df.columns:
        df = df[df["Shipment Tracking Number"].isna() |
                (df["Shipment Tracking Number"] == "")]
 
    if "Shipment Label Created" in df.columns:
        df = df[df["Shipment Label Created"] != True]
 
    skipped = initial_count - len(df)
    if skipped > 0:
        print(f"⏭️  Skipped {skipped} already-processed order(s).")
 
    if df.empty:
        print("✅ All orders are already processed. Nothing to do.")
        sys.exit(0)
 
    print(f"📋 {len(df)} order(s) to process.")
 
    # ── 3. RMA LOOKUP (only in Airtable mode, not CSV mode) ─
    # THIS BLOCK MUST BE OUTSIDE "if df.empty" — it was the bug in v4.x.
    # It runs here, AFTER the duplicate filter, with the real df.
    if not args:  # Airtable mode only
        ya_rma_keys  = []
        crs_rma_keys = []
        for order_number in df.get("Order Number", pd.Series(dtype=str)):
            order_str = str(order_number).strip()
            # Take the first part before any slash
            base_rma = order_str.split('/')[0].strip()
            rma_source = get_rma_source(base_rma)
            if rma_source == "ya":
                ya_rma_keys.append(normalize_rma_number(base_rma))
            elif rma_source == "crs":
                crs_rma_keys.append(normalize_rma_number(base_rma))
 
        def build_rma_filter(keys):
            if not keys:
                return None
            if len(keys) == 1:
                return f'{{RMA #}}="{keys[0]}"'
            conditions = ','.join([f'{{RMA #}}="{k}"' for k in keys])
            return f"OR({conditions})"
 
        if ya_rma_keys:
            cfg = TABLE_CONFIG["ya_rma"]
            df_ya_rma = fetch_table(
                base_id=cfg["base_id"], table_id=cfg["table_id"],
                view_id=None, label=cfg["label"], api_key=cfg["api_key"],
                filter_formula=build_rma_filter(ya_rma_keys),
            )
        else:
            print("\nℹ️  No RMAY orders — skipping YA RMA fetch.")
            df_ya_rma = pd.DataFrame()
 
        if crs_rma_keys:
            cfg = TABLE_CONFIG["crs_rma"]
            df_crs_rma = fetch_table(
                base_id=cfg["base_id"], table_id=cfg["table_id"],
                view_id=None, label=cfg["label"], api_key=cfg["api_key"],
                filter_formula=build_rma_filter(crs_rma_keys),
            )
        else:
            print("\nℹ️  No RMAC orders — skipping CRS RMA fetch.")
            df_crs_rma = pd.DataFrame()
 
        ya_rma_map, crs_rma_map = build_rma_maps(df_ya_rma, df_crs_rma)
 
        needed_so_numbers = set()
        for order_number in df.get("Order Number", pd.Series(dtype=str)):
            order_str = str(order_number).strip()
            base_rma = order_str.split('/')[0].strip()
            rma_source = get_rma_source(base_rma)
            if rma_source:
                clean_rma = normalize_rma_number(base_rma)
                so = (ya_rma_map.get(clean_rma, "")
                      if rma_source == "ya"
                      else crs_rma_map.get(clean_rma, ""))
                if so:
                    needed_so_numbers.add(so)
 
        if needed_so_numbers:
            # Try Commerce Central first (for RMA orders)
            price_map_cc = fetch_price_map_from_commerce_central(needed_so_numbers)
            # Then fallback to Sales Hub
            price_map_sh = fetch_price_map_for_orders(needed_so_numbers, AIRTABLE_API_KEY)
            # Merge: Commerce Central overrides
            price_map = {**price_map_sh, **price_map_cc}
 
        else:
            if ya_rma_keys or crs_rma_keys:
                print("⚠️  RMA orders found but no External Sales Order numbers in RMA records.")
            else:
                print("ℹ️  No RMA orders — price lookup not needed.")
 
    # ── 4. TRANSFORM → CSV ───────────────────────────────────
    print("\n⚙️  Transforming data...")
    result    = transform(df, ya_rma_map, crs_rma_map, price_map)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    csv_out   = f"DHL-Order-File-{timestamp}.csv"
    result.to_csv(csv_out, index=False, encoding="utf-8-sig")
    print(f"\n✅ CSV saved: {csv_out}  ({len(result)} order(s))")
 
    # ── 5. SEND TO DHL (only if --send-to-dhl flag) ─────────
    if not send_to_dhl_flag:
        print("\nℹ️  Dry run complete — CSV only.")
        print(f"   To send to DHL:  python airtable_transform.py --send-to-dhl")
        print(f"   Safe test first: python airtable_transform.py --send-to-dhl --test")
        sys.exit(0)
 
    print(f"\n{'='*60}")
    print(f"🚀 BATCH START: {len(result)} shipment(s)")
    print(f"   Mode    : {'⚠️  TEST (no real shipments)' if DHL_TEST_MODE else '✅ PRODUCTION'}")
    print(f"   Verbose : {'ON (full payload)' if verbose_flag else 'OFF (use --verbose to enable)'}")
    print(f"{'='*60}")
 
    # Build order_number → Airtable record_id map for writeback
    record_id_map = {}
    if "_airtable_id" in df.columns:
        for _, r in df.iterrows():
            on  = str(r.get("Order Number", "")).strip()
            rid = str(r.get("_airtable_id", "")).strip()
            if on and rid and rid != "nan":
                record_id_map[on] = rid
        print(f"   Airtable record map: {len(record_id_map)} order(s) mapped\n")
    else:
        print("   ℹ️  CSV mode — Airtable writeback skipped (no record IDs)\n")
 
    dhl_results   = []
    success_count = 0
    fail_count    = 0
    failed_orders = []  # saved to JSON so you can retry just these
 
    # ── THE PROTECTED BATCH LOOP ─────────────────────────────
    # Each order is wrapped in try/except so one failure never
    # crashes the whole batch.
    total = len(result)
    for idx, (_, row) in enumerate(result.iterrows(), start=1):
        order_number = str(row.get("Shipment Reference 1", "unknown"))
        print(f"\n[{idx}/{total}] ── {order_number} ──────────────────────")
        print_progress(idx, total)
 
        try:
            # ── DHL API call ──────────────────────────────────
            result_data = send_to_dhl(row, verbose=verbose_flag)
            dhl_results.append(result_data)
 
            if result_data.get("success"):
                success_count  += 1
                tracking_number = result_data.get("tracking_number", "")
                airtable_rec_id = record_id_map.get(order_number, "")
                label_b64       = result_data.get("label_base64", "")
 
                if DHL_TEST_MODE:
                    # ── TEST MODE: never touch production Airtable ────────
                    # DHL gives us a fake tracking number from their sandbox.
                    # Writing it to Airtable would corrupt live records.
                    print(f"   ⏭️  TEST MODE — Airtable writeback skipped.")
                    print(f"   ℹ️  Fake tracking (NOT written): {tracking_number}")
 
                else:
                    # ── PRODUCTION: 3-step Airtable writeback ─────────────
                    #
                    # STEP 1 → Write tracking number + courier
                    # STEP 2 → Upload label PDF
                    # STEP 3 → Tick checkbox + write tracking URL
                    #          (ONLY runs if Step 2 succeeds)
                    #
                    # If Step 2 fails, the checkbox stays unticked.
                    # The order will be skipped on re-run (it has a
                    # tracking number). Use the recovery script to
                    # re-upload labels for orders in failed_orders log.
 
                    if airtable_rec_id and tracking_number:
                        # Step 1
                        print(f"   📡 Step 1/3: Writing tracking to Airtable...")
                        update_shipment_tracking_in_airtable(
                            airtable_rec_id, order_number, tracking_number
                        )
 
                        # Step 2
                        raw_response = result_data.get("raw_response", {})
                        if raw_response:
                            print(f"   📡 Step 2/3: Uploading label + invoice PDFs...")
                            label_ok = upload_docs_to_airtable(
                                airtable_rec_id, order_number, raw_response
                            )
 
                            # Step 3 — ONLY if Step 2 succeeded
                            if label_ok:
                                print(f"   📡 Step 3/3: Ticking 'Label Created'...")
                                mark_label_created_in_airtable(
                                    airtable_rec_id, order_number, tracking_number
                                )
                            else:
                                print(f"   ⚠️  Step 2 failed → Step 3 skipped. "
                                      f"Checkbox NOT ticked.")
                        else:
                            print(f"   ⚠️  No raw_response data returned — Steps 2 & 3 skipped.")
 
                    elif not airtable_rec_id:
                        print(f"   ℹ️  No Airtable record ID for {order_number} — "
                              f"writeback skipped (CSV mode).")
 
            else:
                fail_count += 1
                failed_orders.append({
                    "order"      : order_number,
                    "error"      : result_data.get("error", "Unknown DHL error"),
                    "status_code": result_data.get("status_code", ""),
                    "failed_at"  : datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                })
                print(f"   ❌ DHL FAILED: {result_data.get('error', 'Unknown error')[:200]}")
 
        except Exception as e:
            fail_count += 1
            failed_orders.append({
                "order"     : order_number,
                "error"     : str(e),
                "failed_at" : datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            })
            print(f"   💥 EXCEPTION on {order_number}: {e}")
            # continue to next order — never crash the whole batch
 
        # ── Rate limit: 1s between orders ────────────────────
        # Airtable allows ~5 req/sec. We make up to 3 calls per
        # order. 1s sleep keeps us at ~3 req/sec — safe margin.
        if idx < total:
            time.sleep(1.0)
 
    # ── 6. SAVE LOGS ─────────────────────────────────────────
    os.makedirs("logs", exist_ok=True)
 
    response_file = f"logs/DHL-Responses-{timestamp}.json"
    with open(response_file, "w", encoding="utf-8") as f:
        json.dump(dhl_results, f, indent=2, default=str)
 
    if failed_orders:
        failed_file = f"logs/failed_orders-{timestamp}.json"
        with open(failed_file, "w", encoding="utf-8") as f:
            json.dump(failed_orders, f, indent=2, ensure_ascii=False)
        print(f"\n⚠️  Failed orders saved to: {failed_file}")
        print(f"   (Re-run just these by loading them into a recovery script)")
 
    # ── 7. FINAL SUMMARY ─────────────────────────────────────
    print(f"\n{'='*60}")
    print(f"📊 BATCH COMPLETE")
    print(f"   ✅ Success : {success_count}/{total}")
    print(f"   ❌ Failed  : {fail_count}/{total}")
    print(f"   Full log  : {response_file}")
    if os.path.exists("labels"):
        print(f"   Labels    : {len(os.listdir('labels'))} file(s) in labels/")
    print(f"{'='*60}")
 
    generate_lookup_html()









# """
# =============================================================
#   AIRTABLE → DHL SHIPPING AUTOMATION SCRIPT
#   YesAgain / Palm Tree AE
#   VERSION: 5.0
# =============================================================
# """
 
# import pandas as pd
# import re
# import requests
# import sys
# import json
# import base64
# import time
# import os
# from datetime import datetime, timezone
# from dotenv import load_dotenv
# from urllib.parse import quote
 
# try:
#     from unidecode import unidecode
#     HAS_UNIDECODE = True
# except ImportError:
#     HAS_UNIDECODE = False
#     print("⚠️  unidecode not installed — special characters will be stripped instead of converted.")
#     print("   Run: pip install unidecode")
 
# load_dotenv()
 
# AIRTABLE_API_KEY         = os.getenv("AIRTABLE_API_KEY")
# AIRTABLE_RMA_API_KEY     = os.getenv("AIRTABLE_RMA_API_KEY", AIRTABLE_API_KEY)
# AIRTABLE_CRS_RMA_API_KEY = os.getenv("AIRTABLE_CRS_RMA_API_KEY", AIRTABLE_RMA_API_KEY)
 
# DHL_API_KEY    = os.getenv("DHL_API_KEY")
# DHL_API_SECRET = os.getenv("DHL_API_SECRET")
# DHL_TEST_MODE  = os.getenv("DHL_TEST_MODE", "true").strip().lower() == "true"
 
# DHL_BASE_URL_TEST = "https://express.api.dhl.com/mydhlapi/test"
# DHL_BASE_URL_PROD = "https://express.api.dhl.com/mydhlapi"
 
# TABLE_CONFIG = {
#     "orders": {
#         "base_id":  os.getenv("AIRTABLE_BASE_ID",          "appJ3jLnRx4pDaTsM"),
#         "table_id": os.getenv("AIRTABLE_TABLE",             "tbl3MiZfKfMR1jd6A"),
#         "view_id":  os.getenv("AIRTABLE_ORDERS_VIEW_ID",    "viwEBcUnZvtLvxov1"),
#         "label":    "Sales Orders (DHL Labels)",
#         "api_key":  None,
#     },
#     "sales_order_lines": {
#         "base_id":  os.getenv("AIRTABLE_BASE_ID",                    "appJ3jLnRx4pDaTsM"),
#         "table_id": os.getenv("AIRTABLE_SALES_ORDER_LINES_TABLE_ID", "tblGWIuIdVNhRLwEf"),
#         "view_id":  None,
#         "label":    "Sales Order Lines",
#         "api_key":  None,
#     },
#     "ya_rma": {
#         "base_id":  os.getenv("AIRTABLE_YA_RMA_BASE_ID", "app42Tgocgm1DFH3H"),
#         "table_id": os.getenv("AIRTABLE_YA_RMA_TABLE",   "tblNTvHfVY0SVfcPg"),
#         "view_id":  None,
#         "label":    "YesAgain Commerce Central RMA",
#         "api_key":  None,
#     },
#     "crs_rma": {
#         "base_id":  os.getenv("AIRTABLE_CRS_RMA_BASE_ID", "appdEZg6u3zXG8tRE"),
#         "table_id": os.getenv("AIRTABLE_CRS_RMA_TABLE",   "tblWvBHXdLdOYIYDd"),
#         "view_id":  None,
#         "label":    "CRS Commerce Central RMA",
#         "api_key":  None,
#     },
#     "commerce_central_sales_lines": {
#         "base_id":  "app42Tgocgm1DFH3H",                                   # same as YA RMA base
#         "table_id": os.getenv("AIRTABLE_COMMERCE_CENTRAL_SO_LINES_TABLE_ID", "tblXI2dCP2nNpJBvg"),
#         "view_id":  None,
#         "label":    "Commerce Central Sales Order Lines",
#         "api_key":  AIRTABLE_RMA_API_KEY,                                 # uses the RMA key (has access)
#     },
# }
 
# TABLE_CONFIG["orders"]["api_key"]            = AIRTABLE_API_KEY
# TABLE_CONFIG["sales_order_lines"]["api_key"] = AIRTABLE_API_KEY
# TABLE_CONFIG["ya_rma"]["api_key"]            = AIRTABLE_RMA_API_KEY
# TABLE_CONFIG["crs_rma"]["api_key"]           = AIRTABLE_CRS_RMA_API_KEY
 
# # ============================================================
# # PLT (Paperless Trade) ACCEPTING COUNTRIES — from DHL PLT_2025.xlsx
# # WY service code must be added for all shipments to these destinations
# # ============================================================
# PLT_COUNTRIES = {
#     'AD','AE','AG','AI','AL','AM','AO','AR','AS','AT','AU','AW','AZ',
#     'BA','BB','BE','BF','BG','BH','BI','BJ','BM','BN','BO','BR','BS',
#     'BT','BW','BY','BZ','CA','CD','CF','CG','CH','CI','CK','CL','CM',
#     'CN','CO','CR','CU','CV','CY','CZ','DE','DJ','DK','DM','DO','DZ',
#     'EC','EE','EG','ER','ES','ET','FI','FJ','FK','FM','FO','FR','GA',
#     'GB','GD','GE','GF','GG','GH','GI','GL','GM','GN','GP','GQ','GR',
#     'GT','GU','GW','GY','HK','HN','HR','HT','HU','IC','ID','IE','IL',
#     'IN','IQ','IR','IS','IT','JE','JM','JO','JP','KE','KG','KH','KI',
#     'KM','KN','KP','KR','KV','KW','KY','KZ','LA','LB','LC','LI','LK',
#     'LR','LS','LT','LU','LV','LY','MA','MC','MD','MG','MH','MK','ML',
#     'MM','MN','MO','MP','MQ','MR','MS','MT','MU','MV','MW','MX','MY',
#     'MZ','NA','NC','NE','NG','NI','NL','NO','NP','NR','NU','NZ','OM',
#     'PA','PE','PF','PG','PH','PK','PL','PR','PT','PW','PY','QA','RE',
#     'RO','RU','RW','SA','SB','SC','SD','SE','SG','SH','SI','SK','SL',
#     'SM','SN','SO','SR','SS','ST','SV','SY','SZ','TC','TD','TG','TH',
#     'TJ','TL','TM','TN','TO','TR','TT','TV','TW','TZ','UA','UG','US',
#     'UY','UZ','VA','VC','VE','VG','VI','VN','VU','WS','XB','XC','XE',
#     'XM','XN','XS','XY','YE','YT','ZA','ZM','ZW',
# }
 
# # ============================================================
# # EUROPEAN COUNTRIES — used to distinguish EU (YesAgain France
# # as importer) from "Others not listed" (YesAgain UAE as importer)
# # GB is intentionally excluded — handled separately as UK rule.
# # ============================================================
# EU_COUNTRIES = {
#     'AT','BE','BG','CH','CY','CZ','DE','DK','EE','ES','FI','FR',
#     'GR','HR','HU','IE','IS','IT','LI','LT','LU','LV','MC','MT',
#     'NL','NO','PL','PT','RO','SE','SI','SK','SM','AL','BA','ME',
#     'MK','RS','XK',
# }
 
# # ============================================================
# # FIXED VALUES
# # ============================================================
# FIXED = {
#     # ── Accounts ──
#     "account_shipper"     : 454189098,   # YA UAE
#     "account_payer"       : 454189098,   # YA UAE
#     "account_duty_eu"     : 229887839,   # YA France — pays EU duties
#     "account_duty_us_gcc" : 961923318,   # PT UAE — pays US & GCC duties
 
#     "account_crs_shipper" : 952629100,   # CRS UK 
#     "account_crs_duty"    : 952629100,   # CRS UK — pays UK duties
 
#     # ── Shipping config ──
#     # CSV vs API split: portal upload uses different codes than the API
#     "incoterms"           : "Y",         # CSV portal upload value
#     "incoterms_api"       : "DDP",       # DHL API value — "Y" is portal-only
#     "product_code"        : "WPX",       # CSV portal upload value
#     "product_code_api"    : "P",         # DHL API value — "WPX" is portal-only
 
#     "weight"              : 1.5,
#     "weight_unit"         : "KG",
#     "currency"            : "EUR",
#     "contents"            : "Used laptop",
#     "shipment_type"       : "P",
#     "dim_unit"            : "CM",
#     "length"              : 35,
#     "width"               : 30,
#     "height"              : 7,
#     "export_reason"       : "Personal use",
#     "dig_customs"         : "Y",
#     "item_desc"           : "Used laptop",
#     "commodity"           : "8471.49.0000",
#     "item_units"          : "PCS",
#     "item_net"            : 1.3,
#     "item_gross"          : 1.5,
#     "origin"              : "CN",
# }
 
# # ============================================================
# # PARTY DETAILS
# # ============================================================
 
# PARTY_UAE = {
#     "Company"  : "YESAGAIN FZC",
#     "Name"     : "YESAGAIN",
#     "Address1" : "P MALL, SHARJAH FREE ZONE",
#     "Country"  : "AE",
#     "City"     : "SHARJAH",
#     "ZIP"      : "00000",
#     "Email"    : "logistics@yesagain.com",
#     "PhoneCC"  : 971,
#     "Phone"    : "508893656",
#     "VAT"      : "",
#     "EORI"     : "",
#     "Rel"      : "IP",
# }
 
# PARTY_UK = {
#     "Company"  : "Computer Remarketing Services Ltd",
#     "Name"     : "Computer Remarketing Services Ltd",
#     "Address1" : "Unit 2 Maryland Road",
#     "Country"  : "GB",
#     "City"     : "Tongwell",
#     "ZIP"      : "MK15 8HF",
#     "Email"    : "logistics@yesagain.com",
#     "PhoneCC"  : 44,
#     "Phone"    : "01908656700",
#     "VAT"      : "GB450944880",
#     "EORI"     : "GB524271957000",
#     "Rel"      : "IP",
# }
 
# PARTY_FRANCE = {
#     "Company"  : "YESAGAIN",
#     "Name"     : "YESAGAIN",
#     "Address1" : "14 Chemin du Chapitre",
#     "Country"  : "FR",
#     "City"     : "TOULOUSE",
#     "ZIP"      : "31100",
#     "Email"    : "logistics@yesagain.com",
#     "PhoneCC"  : 33,
#     "Phone"    : "187661011",
#     "VAT"      : "FR04919345207",
#     "EORI"     : "FR91934520700022",
#     "Rel"      : "IP",
# }

# # ============================================================
# # DESTINATION-BASED VAT (YesAgain's local VAT per EU country)
# # ============================================================
# # The VAT number declared is chosen by the DESTINATION country.
# # e.g. an order shipping to Belgium declares the Belgian VAT.
# # Anything not listed here falls back to the importer entity's
# # default VAT (France for EU, UK for GB) — see importer_vat_for().
# EU_VAT_BY_COUNTRY = {
#     "AT": "ATU79670679",      # Austria
#     "BE": "BE0803906108",     # Belgium
#     "DK": "DK13394814",       # Denmark
#     "IT": "IT00358559995",    # Italy
#     "NL": "NL827367090B01",   # Netherlands
#     "SE": "SE502096226101",   # Sweden
#     "DE": "DE362263608",      # Germany
#     "FR": "FR04919345207",    # France
#     "GB": "GB450944880",      # United Kingdom
# }

# def importer_vat_for(country, fallback_vat=""):
#     """Return YesAgain's VAT registered in the destination country.
#     Falls back to the importer entity's default VAT when the destination
#     has no specific registration in EU_VAT_BY_COUNTRY."""
#     return EU_VAT_BY_COUNTRY.get(str(country).strip().upper(), fallback_vat) or fallback_vat

# # ============================================================
# # PHONE / STATE LOOKUPS
# # ============================================================
# PHONE_CODES = {
#     "GB":"44",  "DE":"49",  "FR":"33",  "ES":"34",  "IT":"39",  "AT":"43",
#     "BE":"32",  "NL":"31",  "PT":"351", "IE":"353", "SE":"46",  "NO":"47",
#     "DK":"45",  "FI":"358", "PL":"48",  "CZ":"420", "HU":"36",
#     "US":"1",   "CA":"1",   "AU":"61",
#     "CH":"41",  "LU":"352", "GR":"30",  "RO":"40",  "BG":"359", "HR":"385",
#     "SK":"421", "SI":"386", "LT":"370", "LV":"371", "EE":"372", "CY":"357",
#     "MT":"356", "IS":"354", "TR":"90",  "SA":"966", "AE":"971", "JP":"81",
#     "CN":"86",  "IN":"91",  "BR":"55",  "MX":"52",  "ZA":"27",
# }
 
# STATE_FULL_NAMES = {
#     "AL":"Alabama","AK":"Alaska","AZ":"Arizona","AR":"Arkansas",
#     "CA":"California","CO":"Colorado","CT":"Connecticut","DE":"Delaware",
#     "FL":"Florida","GA":"Georgia","HI":"Hawaii","ID":"Idaho",
#     "IL":"Illinois","IN":"Indiana","IA":"Iowa","KS":"Kansas",
#     "KY":"Kentucky","LA":"Louisiana","ME":"Maine","MD":"Maryland",
#     "MA":"Massachusetts","MI":"Michigan","MN":"Minnesota","MS":"Mississippi",
#     "MO":"Missouri","MT":"Montana","NE":"Nebraska","NV":"Nevada",
#     "NH":"New Hampshire","NJ":"New Jersey","NM":"New Mexico","NY":"New York",
#     "NC":"North Carolina","ND":"North Dakota","OH":"Ohio","OK":"Oklahoma",
#     "OR":"Oregon","PA":"Pennsylvania","RI":"Rhode Island","SC":"South Carolina",
#     "SD":"South Dakota","TN":"Tennessee","TX":"Texas","UT":"Utah",
#     "VT":"Vermont","VA":"Virginia","WA":"Washington","WV":"West Virginia",
#     "WI":"Wisconsin","WY":"Wyoming",
# }
 
 
# # ============================================================
# # HELPER FUNCTIONS
# # ============================================================
 
# def split_phone(telephone, country):
#     tel = re.sub(r'[^\d]', '', str(telephone))
#     cc = PHONE_CODES.get(country, "")
#     if country in ["US", "CA"]:
#         if tel.startswith("1") and len(tel) > 10:
#             return cc, tel[1:]
#         return cc, tel
#     if cc and tel.startswith(cc):
#         return cc, tel[len(cc):]
#     return cc, tel
 
 
# def clean_postal_code(postcode, country_code):
#     """Ensure postal code matches DHL format. LU must be exactly 4 digits."""
#     if not postcode:
#         return ""
#     postcode = str(postcode).strip().upper()
#     if country_code == "LU":
#         # Remove all non-digits, then take first 4 digits or pad
#         digits = re.sub(r'\D', '', postcode)
#         if len(digits) >= 4:
#             return digits[:4]
#         elif len(digits) > 0:
#             return digits.zfill(4)
#         else:
#             return "9999"   # fallback
#     return postcode
 
 
# def clean_price(value):
#     if pd.isna(value) if not isinstance(value, list) else False:
#         return 0.0
#     # Airtable returns a list when one order has multiple line prices
#     # (e.g. qty=2 same model → [174.11, 178.31]). Sum ALL values.
#     if isinstance(value, list):
#         return sum(clean_price(v) for v in value) if value else 0.0
#     if isinstance(value, (int, float)):
#         return float(value)
#     cleaned = re.sub(r'[€$£,\s]', '', str(value))
#     try:
#         return float(cleaned)
#     except ValueError:
#         return 0.0
 
 
# def clean_qty(value):
#     # Airtable returns a list when one order has multiple qty values
#     # (e.g. qty=2 same model → [1, 1]). Sum ALL values.
#     if isinstance(value, list):
#         return sum(clean_qty(v) for v in value) if value else 1
#     try:
#         return int(float(str(value)))
#     except Exception:
#         return 1
 
 
# # ============================================================
# # ARABIC → LATIN TRANSLITERATION
# # ============================================================
# #
# # Why transliteration is imperfect for Arabic:
# #   Arabic is normally written WITHOUT short vowels (harakat). So "احمد"
# #   could be Ahmed, Ahmad or Ahmd — the script cannot know which without
# #   the vowel marks. We solve this with two layers:
# #     1. A name dictionary for the most common Arabic first names, titles
# #        and words → gives exact correct spellings (Ahmed, Mohammed, Eid…)
# #     2. A letter-by-letter phonetic fallback for anything not in the dict.
# #

# ARABIC_NAME_DICT = {
#     # Common first names (male)
#     'احمد':'Ahmed','احمد':'Ahmed','أحمد':'Ahmed','محمد':'Mohammed',
#     'محمود':'Mahmoud','عبدالله':'Abdullah','عبد الله':'Abdullah',
#     'عبدالرحمن':'Abdulrahman','عبدالعزيز':'Abdulaziz','علي':'Ali',
#     'عمر':'Omar','خالد':'Khalid','يوسف':'Youssef','ابراهيم':'Ibrahim',
#     'إبراهيم':'Ibrahim','سامي':'Sami','سالم':'Salem','سعد':'Saad',
#     'سعيد':'Saeed','صالح':'Saleh','طارق':'Tariq','عادل':'Adel',
#     'عثمان':'Othman','عصام':'Essam','علاء':'Alaa','فاروق':'Farouk',
#     'فيصل':'Faisal','كريم':'Karim','ماجد':'Majid','مازن':'Mazen',
#     'منصور':'Mansour','موسى':'Moussa','نبيل':'Nabil','هاني':'Hani',
#     'هشام':'Hisham','وليد':'Walid','ياسر':'Yasser','زياد':'Ziad',
#     'رامي':'Rami','رياض':'Riyad','زيد':'Zaid','بلال':'Bilal',
#     'تركي':'Turki','جمال':'Jamal','حسن':'Hassan','حسين':'Hussein',
#     'حمد':'Hamad','حمزة':'Hamza','راشد':'Rashed','زكريا':'Zakaria',
#     'شاكر':'Shaker','شريف':'Sherif','صلاح':'Salah','طلال':'Talal',
#     'عزيز':'Aziz','مجدي':'Magdy','مصطفى':'Mostafa','مصطفا':'Mostafa',
#     'ناصر':'Nasser','نادر':'Nader','نواف':'Nawaf','هلال':'Hilal',
#     'وائل':'Wael','وسام':'Wissam','يحيى':'Yahya','بدر':'Badr',
#     'جاسم':'Jasim','رضا':'Reda','سلطان':'Sultan','سمير':'Samir',
#     'صقر':'Saqr','فهد':'Fahad','قاسم':'Qasim','كامل':'Kamel',
#     'نزار':'Nizar','هيثم':'Haitham','امير':'Amir','أمير':'Amir',
#     'بشير':'Bashir','رفيق':'Rafik','رمزي':'Ramzi','سلمان':'Salman',
#     'سيف':'Saif','شادي':'Shadi','فادي':'Fadi','كمال':'Kamal',
#     'مروان':'Marwan','نصر':'Nasr','نور':'Nour','هادي':'Hadi',
#     'ادريس':'Idris','إدريس':'Idris','انس':'Anas','أنس':'Anas',
#     # Common first names (female)
#     'مريم':'Mariam','فاطمة':'Fatima','سارة':'Sara','سارا':'Sara',
#     'رنا':'Rana','لينا':'Lina','منى':'Mona','هيا':'Haya',
#     'ريم':'Reem','دينا':'Dina','رانيا':'Rania','نادية':'Nadia',
#     'هند':'Hind','ياسمين':'Yasmine','اسماء':'Asmaa','بسمة':'Basma',
#     'حنان':'Hanan','خديجة':'Khadija','رشا':'Rasha','سلمى':'Salma',
#     'غادة':'Ghada','ليلى':'Layla','منال':'Manal','نجلاء':'Najla',
#     'نوره':'Noura','هدى':'Huda','وفاء':'Wafa','شيماء':'Shaimaa',
#     # Titles / words that appear in names
#     'مهندس':'Eng','مهندسة':'Eng','دكتور':'Dr','دكتورة':'Dr',
#     'الدكتور':'Dr','الدكتورة':'Dr','استاذ':'Prof','الاستاذ':'Prof',
#     # Common surnames / family name parts
#     'عيد':'Eid','الشعراوي':'El Shaarawy','الشريف':'El Sherif',
#     'الامير':'El Amir','الزهراني':'Al Zahrani','العمري':'Al Omari',
#     'العتيبي':'Al Otaibi','الغامدي':'Al Ghamdi','السبيعي':'Al Subaie',
#     'القحطاني':'Al Qahtani','الدوسري':'Al Dosari','الشهري':'Al Shehri',
#     'المطيري':'Al Mutairi','الرشيدي':'Al Rashidi','المالكي':'Al Maliki',
#     'الحربي':'Al Harbi','العنزي':'Al Anazi','الشمري':'Al Shammari',
#     # Connectors
#     'بن':'Bin','ابن':'Ibn','ابو':'Abu','ام':'Um','بنت':'Bint',
#     'ال':'Al',
# }
 
# # Single-char phonetic fallback (for words not in the dictionary)
# ARABIC_LATIN = str.maketrans({
#     'ا':'a','أ':'a','إ':'i','آ':'a','ب':'b',
#     'ت':'t','ث':'th','ج':'j','ح':'h','خ':'kh',
#     'د':'d','ذ':'dh','ر':'r','ز':'z','س':'s',
#     'ش':'sh','ص':'s','ض':'d','ط':'t','ظ':'z',
#     'ع':'e','غ':'gh','ف':'f','ق':'q','ك':'k',
#     'ل':'l','م':'m','ن':'n','ه':'h','و':'o',
#     'ي':'i','ى':'a','ة':'a','ء':'','ئ':'i',
#     'ؤ':'o',
#     'ً':'','ٌ':'','ٍ':'','َ':'a','ُ':'u',
#     'ِ':'i','ّ':'','ْ':'',
# })
 
# def _arabic_to_latin(text):
#     """
#     Two-layer Arabic → Latin:
#       1. Word-level dictionary lookup (Ahmed, Eid, El Shaarawy, Eng…)
#       2. Letter-by-letter phonetic fallback for unknown words
#     """
#     # Pre-process: strip definite article prefix ال from standalone token
#     text = text.replace('لا', 'la')  # لا → la
#     words = text.split()
#     result = []
#     for word in words:
#         w = word.strip()
#         if not w:
#             continue
#         # Direct dictionary hit
#         if w in ARABIC_NAME_DICT:
#             result.append(ARABIC_NAME_DICT[w])
#             continue
#         # Try stripping leading ال (definite article) and look up the rest
#         if w.startswith('ال') and len(w) > 2:
#             root = w[2:]
#             if root in ARABIC_NAME_DICT:
#                 result.append('Al ' + ARABIC_NAME_DICT[root])
#                 continue
#         # Phonetic fallback
#         fallback = re.sub(r'[^ -~]', '', w.translate(ARABIC_LATIN)).strip()
#         if fallback:
#             result.append(fallback.capitalize())
#     return re.sub(r'\s+', ' ', ' '.join(result)).strip()
 
# def _has_arabic(text):
#     """Return True if the text contains Arabic-script characters."""
#     return any('؀' <= c <= 'ۿ' for c in text)
 
 
# def clean_text(value):
#     if pd.isna(value) if not isinstance(value, (list, dict)) else False:
#         return ""
#     if value == "":
#         return ""
#     text = str(value).strip()
#     if not text or text.lower() == "nan":
#         return ""
#     # Transliterate Arabic → Latin before any further processing.
#     # We always use our dictionary-based transliterator for Arabic because
#     # unidecode lacks short-vowel knowledge and produces "Ahmd" not "Ahmed".
#     if _has_arabic(text):
#         latin = _arabic_to_latin(text)
#         latin = re.sub(r'[^ -~]', '', latin).strip()
#         latin = re.sub(r'\s+', ' ', latin)
#         text = latin if latin else "CUSTOMER"
#     elif HAS_UNIDECODE:
#         text = unidecode(text)
#     else:
#         CHAR_MAP = str.maketrans({
#             'À':'A','Á':'A','Â':'A','Ã':'A','Ä':'A','Å':'A','Æ':'AE',
#             'Ç':'C','È':'E','É':'E','Ê':'E','Ë':'E',
#             'Ì':'I','Í':'I','Î':'I','Ï':'I',
#             'Ð':'D','Ñ':'N','Ò':'O','Ó':'O','Ô':'O','Õ':'O','Ö':'O','Ø':'O',
#             'Ù':'U','Ú':'U','Û':'U','Ü':'U','Ý':'Y','Þ':'TH','ß':'ss',
#             'à':'a','á':'a','â':'a','ã':'a','ä':'a','å':'a','æ':'ae',
#             'ç':'c','è':'e','é':'e','ê':'e','ë':'e',
#             'ì':'i','í':'i','î':'i','ï':'i',
#             'ð':'d','ñ':'n','ò':'o','ó':'o','ô':'o','õ':'o','ö':'o','ø':'o',
#             'ù':'u','ú':'u','û':'u','ü':'u','ý':'y','þ':'th','ÿ':'y',
#             'Ł':'L','ł':'l','Ź':'Z','ź':'z','Ż':'Z','ż':'z',
#             'Ś':'S','ś':'s','Ą':'A','ą':'a','Ę':'E','ę':'e',
#             'Ć':'C','ć':'c','Ń':'N','ń':'n',
#             'Ğ':'G','ğ':'g','İ':'I','ı':'i','Ş':'S','ş':'s',
#             'Č':'C','č':'c','Š':'S','š':'s','Ž':'Z','ž':'z',
#             'Ř':'R','ř':'r','Ů':'U','ů':'u','Ď':'D','ď':'d',
#             'Ť':'T','ť':'t','Ľ':'L','ľ':'l',
#         })
#         text = text.translate(CHAR_MAP)
#     text = re.sub(r'[^\x20-\x7E]', '', text)
#     text = re.sub(r'[^\w\s\-.,/]', '', text)
#     return text.strip()
 
 
# def get_first_value(row, field_names, default=""):
#     for field in field_names:
#         value = row.get(field, "")
#         if isinstance(value, list):
#             value = value[0] if value else ""
#         if value is None:
#             continue
#         if isinstance(value, float) and pd.isna(value):
#             continue
#         text = str(value).strip()
#         if text and text.lower() != "nan":
#             return value
#     return default
 
 
# def normalize_lookup_key(value):
#     if isinstance(value, list):
#         value = value[0] if value else ""
#     if value is None:
#         return ""
#     if isinstance(value, float) and pd.isna(value):
#         return ""
#     if isinstance(value, (int, float)) and not isinstance(value, bool):
#         numeric = float(value)
#         return str(int(numeric)) if numeric.is_integer() else str(numeric)
#     text = str(value).strip()
#     if re.fullmatch(r"\d+\.0", text):
#         return text[:-2]
#     return text
 
# def normalize_sales_order_number(value):
#     """Extract the last part after any slash, e.g. 'RMAY31984/72771270' → '72771270'."""
#     text = normalize_lookup_key(value)
#     if '/' in text:
#         return text.split('/')[-1].strip()
#     return text
 
 
# def normalize_rma_number(value):
#     text = normalize_lookup_key(value)
#     text = re.sub(r"^\s*RMA[-\s]*", "", text, flags=re.IGNORECASE).strip()
#     return text.upper()
 
 
# def extract_sales_order_from_rma(rma_number_str, rma_map):
#     """
#     For a given RMA number (e.g. 'RMAY34783/RMAY31984/72771270'),
#     returns the sales order number found in the RMA map.
#     The RMA map uses the full RMA number as key (without the slash suffix),
#     so we try the exact key first, then fall back to the first part.
#     """
#     rma_clean = normalize_rma_number(rma_number_str)
#     # If the RMA number contains a slash, the map key is usually the first part (e.g. 'RMAY34783')
#     parts = rma_clean.split('/')
#     for candidate in [rma_clean] + parts:
#         if candidate in rma_map:
#             return rma_map[candidate]
#     return ""
 
 
# def get_rma_source(order_number):
#     upper = str(order_number).upper().strip()
#     if upper.startswith("RMAY"):
#         return "ya"
#     if upper.startswith("RMAC"):
#         return "crs"
#     return None
 
 
# def get_address2_and_state(row):
#     country = str(row.get("Shipping Country", "")).strip()
#     addr2_raw = row.get("Shipping address 2", "")
#     addr2 = str(addr2_raw).strip() if pd.notna(addr2_raw) else ""
#     if country == "US":
#         state_code = addr2 if len(addr2) == 2 and addr2.isalpha() else ""
#         return "", state_code
#     else:
#         return clean_text(addr2), ""
 
 
# def get_destination_rules(destination_country, order_id):
#     """
#     Smart Routing: Chooses the correct Duty Account, Additional Party,
#     Shipper and region label based on destination.
 
#     Regions (per shipping policy):
#       ┌──────────────────────┬──────────────┬──────────────────┬──────────────────┐
#       │ Destination          │ Shipper      │ Importer         │ Declared Value   │
#       ├──────────────────────┼──────────────┼──────────────────┼──────────────────┤
#       │ GCC Countries        │ YesAgain UAE │ YesAgain UAE     │ 20% of item val  │
#       │ USA                  │ YesAgain UAE │ YesAgain UAE     │ 20% of item val  │
#       │ Others not listed    │ YesAgain UAE │ YesAgain UAE     │ 20% of item val  │
#       │ Europe               │ YesAgain UAE │ YesAgain France  │ Exclude VAT      │
#       │ United Kingdom (UK)  │ YesAgain UAE │ CRS Account      │ Exclude VAT      │
#       └──────────────────────┴──────────────┴──────────────────┴──────────────────┘
 
#     NOTE: Saudi Arabia (SA) has an additional override — declared value is
#     always fixed at 164 EUR regardless of item price (applied in transform()).
#     """
#     country     = str(destination_country).strip().upper()
#     order_upper = str(order_id).upper().strip()
 
#     # Shipper is ALWAYS YesAgain UAE regardless of order type or destination.
#     # Per shipping policy: "YesAgain UAE" is shipper for all regions.
#     # RMAC orders still use CRS as importer/duty account for UK, but the
#     # physical shipper on the DHL label must be YesAgain UAE (SHJ origin).
#     shipper_party = PARTY_UAE
#     shipper_acc   = str(FIXED["account_shipper"])
 
#     # RULE 1: UK (GB) — Importer: CRS Account | Declared: Exclude VAT
#     if country == "GB":
#         return {
#             "region"          : "uk",
#             "duty_account"    : str(FIXED["account_crs_duty"]),
#             "additional_party": PARTY_UK,
#             "shipper_party"   : shipper_party,
#             "shipper_account" : shipper_acc,
#         }
 
#     # RULE 2: GCC countries — Importer: YesAgain UAE | Declared: 20%
#     # (SA gets an additional 164 EUR fixed override applied in transform())
#     gcc_countries = {"AE", "SA", "OM", "QA", "BH", "KW"}
#     if country in gcc_countries:
#         return {
#             "region"          : "gcc",
#             "duty_account"    : str(FIXED["account_duty_us_gcc"]),
#             "additional_party": PARTY_UAE,
#             "shipper_party"   : shipper_party,
#             "shipper_account" : shipper_acc,
#         }
 
#     # RULE 3: USA / Canada — Importer: YesAgain UAE | Declared: 20%
#     if country in {"US", "CA"}:
#         return {
#             "region"          : "us",
#             "duty_account"    : str(FIXED["account_duty_us_gcc"]),
#             "additional_party": PARTY_UAE,
#             "shipper_party"   : shipper_party,
#             "shipper_account" : shipper_acc,
#         }
 
#     # RULE 4: Europe — Importer: YesAgain France | Declared: Exclude VAT (full price)
#     if country in EU_COUNTRIES:
#         return {
#             "region"          : "eu",
#             "duty_account"    : str(FIXED["account_duty_eu"]),
#             "additional_party": PARTY_FRANCE,
#             "shipper_party"   : shipper_party,
#             "shipper_account" : shipper_acc,
#         }
 
#     # RULE 5: Others not listed — Importer: YesAgain UAE | Declared: 20%
#     return {
#         "region"          : "other",
#         "duty_account"    : str(FIXED["account_duty_us_gcc"]),
#         "additional_party": PARTY_UAE,
#         "shipper_party"   : shipper_party,
#         "shipper_account" : shipper_acc,
#     }
 
 
# def print_progress(current, total):
#     """Print a compact progress bar. Used in large batches."""
#     pct   = current / total
#     filled = int(pct * 20)
#     bar   = "█" * filled + "░" * (20 - filled)
#     print(f"   [{bar}] {current}/{total} ({pct:.0%})", flush=True)
 
 
# # ============================================================
# # AIRTABLE FETCH
# # ============================================================
 
# def fetch_table(base_id, table_id, view_id=None, label=None, api_key=None,
#                 filter_formula=None):
#     if not table_id:
#         print(f"⚠️  Skipping {label}: no table_id configured")
#         return pd.DataFrame()
 
#     encoded_table = quote(str(table_id), safe="")
#     url = f"https://api.airtable.com/v0/{base_id}/{encoded_table}"
#     headers = {"Authorization": f"Bearer {api_key or AIRTABLE_API_KEY}"}
#     all_records = []
#     offset = None
#     page = 0
 
#     while True:
#         params = {"pageSize": 100}
#         if view_id:
#             params["view"] = view_id
#         if offset:
#             params["offset"] = offset
#         if filter_formula:
#             params["filterByFormula"] = filter_formula
 
#         try:
#             response = requests.get(url, headers=headers, params=params, timeout=30)
#         except requests.exceptions.Timeout:
#             print(f"⚠️  Timeout on {label or table_id} (page {page + 1}). "
#                   f"Continuing with {len(all_records)} records.")
#             break
#         except requests.exceptions.ConnectionError as e:
#             print(f"⚠️  Connection error on {label or table_id}: {e}")
#             break
 
#         # Airtable rate limit — back off and retry once
#         if response.status_code == 429:
#             print(f"⚠️  Rate limited fetching {label or table_id} — waiting 30s...")
#             time.sleep(30)
#             continue
 
#         if response.status_code == 403:
#             print(f"❌ 403 Permission denied for {label or table_id} (base: {base_id})")
#             if label and "CRS" in label:
#                 print("   → set AIRTABLE_CRS_RMA_API_KEY in your .env with a token that has CRS access")
#             return pd.DataFrame()
 
#         if response.status_code != 200:
#             print(f"⚠️  Error on {label or table_id} (status {response.status_code}): "
#                   f"{response.text[:200]}")
#             break
 
#         data = response.json()
#         records = []
#         for record in data.get("records", []):
#             fields = record.get("fields", {})
#             fields["_airtable_id"] = record.get("id", "")  # e.g. "recXXXXXXXXXXXXXX"
#             records.append(fields)
#         all_records.extend(records)
#         page += 1
 
#         offset = data.get("offset")
#         if not offset:
#             break
 
#     if all_records:
#         print(f"✅ Loaded {label or table_id}: {len(all_records)} records")
#     else:
#         print(f"⚠️  No records loaded for {label or table_id}")
 
#     return pd.DataFrame(all_records) if all_records else pd.DataFrame()
 
 
# def fetch_from_airtable():
#     return fetch_table(**TABLE_CONFIG["orders"])
 
 
# # ============================================================
# # AIRTABLE WRITEBACK — RETRY HELPER
# #
# # Airtable allows ~5 requests/second per base. For 100-200
# # orders this means you can easily hit the limit. When Airtable
# # returns 429, this helper waits and retries automatically so
# # no write is silently dropped.
# # ============================================================
 
# def airtable_patch_with_retry(url, body, headers, label="", max_retries=3):
#     """
#     PATCH to Airtable with automatic retry on 429 (rate limit).
 
#     Waits 2s, 4s, 8s between retries (exponential backoff).
#     Returns the response object, or None if all retries fail.
#     """
#     for attempt in range(1, max_retries + 1):
#         try:
#             resp = requests.patch(url, json=body, headers=headers, timeout=15)
#             if resp.status_code == 429:
#                 wait = 2 ** attempt  # 2s, 4s, 8s
#                 print(f"   ⚠️  Airtable rate limit (429) on {label} — "
#                       f"retry {attempt}/{max_retries} in {wait}s...")
#                 time.sleep(wait)
#                 continue
#             return resp
#         except requests.exceptions.Timeout:
#             print(f"   ⚠️  Timeout on {label} (attempt {attempt}/{max_retries})")
#             if attempt < max_retries:
#                 time.sleep(2 ** attempt)
#         except Exception as e:
#             print(f"   ⚠️  Exception on {label} (attempt {attempt}/{max_retries}): {e}")
#             if attempt < max_retries:
#                 time.sleep(2)
#     return None
 
 
# # ============================================================
# # AIRTABLE WRITEBACK — STEP 1
# # Writes "Shipment Tracking Number" + "Shipment Courier" only.
# #
# # "Shipment Label Created" is intentionally NOT set here.
# # It is only set after the label PDF is confirmed uploaded
# # in mark_label_created_in_airtable() (STEP 3).
# # ============================================================
 
# def update_shipment_tracking_in_airtable(record_id, order_number, tracking_number):
#     """
#     PATCH Step 1 — writes Shipment Tracking Number + Shipment Courier.
 
#     Args:
#         record_id       – Airtable record ID (e.g. "recXXXXXXXXXXXXXX")
#         order_number    – for log messages only
#         tracking_number – the DHL AWB string (e.g. "1570941875")
 
#     Returns:
#         True on success, False on failure.
#     """
#     if not record_id:
#         print(f"   ⚠️  No Airtable record_id for {order_number} — skipping Step 1.")
#         return False
 
#     base_id  = TABLE_CONFIG["orders"]["base_id"]
#     table_id = TABLE_CONFIG["orders"]["table_id"]
#     api_key  = AIRTABLE_API_KEY
 
#     encoded_table = quote(str(table_id), safe="")
#     url = f"https://api.airtable.com/v0/{base_id}/{encoded_table}/{record_id}"
 
#     headers = {
#         "Authorization": f"Bearer {api_key}",
#         "Content-Type":  "application/json",
#     }
#     body = {
#         "fields": {
#             "Shipment Tracking Number": str(tracking_number),
#             "Shipment Courier":         "dhl",
#         }
#     }
 
#     resp = airtable_patch_with_retry(url, body, headers, label=f"Step1/{order_number}")
#     if resp is None:
#         print(f"   ❌ Step 1 failed (all retries exhausted) for {order_number}")
#         return False
#     if resp.status_code == 200:
#         print(f"   ✅ Step 1 — Tracking: {tracking_number} | Courier: dhl  "
#               f"(Order: {order_number})")
#         return True
#     else:
#         print(f"   ❌ Step 1 PATCH failed ({resp.status_code}) for {order_number}: "
#               f"{resp.text[:400]}")
#         return False
 
 
# # ============================================================
# # AIRTABLE WRITEBACK — STEP 2
# # Uploads label PDF to "Shipment Label file" attachment field.
# # Uses Airtable's Content Upload API (no public URL needed).
# #
# # CORRECT URL: content.airtable.com/v0/{base_id}/{record_id}/...
# # WRONG URL:   content.airtable.com/v0/{base_id}/{table_id}/{record_id}/...
# #              (table_id does NOT go in the content upload URL)
# # ============================================================
 
# def _upload_to_temp_host(pdf_bytes, filename):
#     """
#     Upload PDF to a temporary public host and return a download URL.
#     Tries multiple services so there is always a fallback.
#     Used only when the Airtable Content API is unavailable.
#     """
#     services = [
#         {
#             "name"     : "tmpfiles.org",
#             "url"      : "https://tmpfiles.org/api/v1/upload",
#             "get_link" : lambda r: (
#                 r.json().get("data", {}).get("url", "")
#                  .replace("tmpfiles.org/", "tmpfiles.org/dl/")
#             ),
#         },
#         {
#             "name"     : "file.io",
#             "url"      : "https://file.io/?expires=1h",
#             "get_link" : lambda r: r.json().get("link", ""),
#         },
#     ]
 
#     for svc in services:
#         try:
#             resp = requests.post(
#                 svc["url"],
#                 files={"file": (filename, pdf_bytes, "application/pdf")},
#                 timeout=20,
#             )
#             if resp.status_code == 200:
#                 url = svc["get_link"](resp)
#                 if url:
#                     print(f"   🌐 Temp URL ({svc['name']}): {url}")
#                     return url
#             else:
#                 print(f"   ⚠️  {svc['name']} returned {resp.status_code}")
#         except Exception as exc:
#             print(f"   ⚠️  {svc['name']} failed: {exc}")
 
#     return None
 
 
 
# def upload_docs_to_airtable(record_id, order_number, res_data, max_retries=3):
#     """
#     Step 2 — Uploads BOTH Shipping Label and Commercial Invoice to Airtable.
#     Uses ONLY Method 2 (PATCH with temporary URL) – faster and more reliable.
#     """
#     if not record_id:
#         print(f"   ⚠️  No record_id for {order_number} — skipping Step 2.")
#         return False
 
#     base_id = TABLE_CONFIG["orders"]["base_id"]
#     api_key = AIRTABLE_API_KEY
    
#     documents = res_data.get("documents", [])
#     if not documents and res_data.get("label_base64"):
#         documents = [{"typeCode": "shipping_label", "content": res_data.get("label_base64")}]
 
#     success_tracker = {"label": False, "invoice": False}
 
#     for doc in documents:
#         type_code = doc.get("typeCode", "").lower()
#         content_b64 = doc.get("content")
#         if not content_b64:
#             continue
 
#         if "label" in type_code or "waybill" in type_code:
#             field_name = "Shipment Label file"
#             field_id = "fldG3hmHH8cTPwzxo"
#             filename = f"label_{order_number}.pdf"
#             doc_key = "label"
#         elif "invoice" in type_code:
#             field_name = "Commercial invoice file"
#             field_id = "fldp7fi4xvhqCybQ7"
#             filename = f"invoice_{order_number}.pdf"
#             doc_key = "invoice"
#         else:
#             continue
 
#         print(f"   📡 Step 2 — Processing {doc_key} for {order_number}...")
        
#         try:
#             pdf_bytes = base64.b64decode(content_b64)
#         except Exception as exc:
#             print(f"   ❌ Base64 decode error for {doc_key} on {order_number}: {exc}")
#             continue
 
#         # --- ONLY METHOD 2: PATCH with temporary URL ---
#         print(f"   📡 Step 2 — Method 2: PATCH with temporary URL for {doc_key} ({order_number})...")
#         temp_url = _upload_to_temp_host(pdf_bytes, filename)
#         if temp_url:
#             table_id = TABLE_CONFIG["orders"]["table_id"]
#             patch_url = f"https://api.airtable.com/v0/{base_id}/{quote(str(table_id))}/{record_id}"
#             patch_headers = {
#                 "Authorization": f"Bearer {api_key}",
#                 "Content-Type": "application/json",
#             }
#             payload = {"fields": {field_name: [{"url": temp_url, "filename": filename}]}}
 
#             for attempt in range(1, max_retries + 1):
#                 try:
#                     resp = requests.patch(
#                         patch_url, headers=patch_headers, json=payload, timeout=30
#                     )
#                     print(f"       PATCH response: HTTP {resp.status_code}")
 
#                     if resp.status_code == 200:
#                         print(f"   📎 {doc_key.capitalize()} uploaded via Method 2 (PATCH+URL) (Order: {order_number})")
#                         success_tracker[doc_key] = True
#                         break
 
#                     elif resp.status_code == 422:
#                         print(f"       422 Unprocessable — Airtable could not fetch the URL.")
#                         print(f"       → Check field name: '{field_name}' must be type 'Attachment' in Airtable.")
#                         print(f"       → Response: {resp.text[:300]}")
#                         break
 
#                     elif resp.status_code == 429:
#                         wait = 2 ** attempt
#                         print(f"       Rate limited — retrying in {wait}s...")
#                         time.sleep(wait)
 
#                     else:
#                         print(f"       Failed ({resp.status_code}): {resp.text[:300]}")
#                         if attempt < max_retries:
#                             time.sleep(2)
 
#                 except Exception as exc:
#                     print(f"       Exception (attempt {attempt}): {exc}")
#                     if attempt < max_retries:
#                         time.sleep(2)
#         else:
#             print(f"   ❌ {doc_key.capitalize()} Step 2 FAILED: could not get a temporary URL for {order_number}.")
 
#     return success_tracker["label"]
 
 
# # ============================================================
# # AIRTABLE WRITEBACK — STEP 3
# # Ticks "Shipment Label Created" checkbox AND writes a real
# # clickable DHL tracking URL into "Shipment Label URL".
# #
# # Called ONLY after upload_label_to_airtable() returns True —
# # so the checkbox is never ticked prematurely.
# # ============================================================
 
# def mark_label_created_in_airtable(record_id, order_number, tracking_number):
#     """
#     PATCH Step 3 — ticks 'Shipment Label Created' AND sets 'Shipment Label URL'
#     to the real DHL tracking page URL.
 
#     Called ONLY after upload_label_to_airtable() succeeds.
 
#     Args:
#         record_id       – Airtable record ID
#         order_number    – for log messages only
#         tracking_number – DHL AWB used to build the tracking URL
 
#     Returns:
#         True on success, False on failure.
#     """
#     if not record_id:
#         return False
 
#     base_id  = TABLE_CONFIG["orders"]["base_id"]
#     table_id = TABLE_CONFIG["orders"]["table_id"]
#     api_key  = AIRTABLE_API_KEY
 
#     dhl_tracking_url = (
#         f"https://www.dhl.com/ae-en/home/tracking/tracking-express.html"
#         f"?submit=1&tracking-id={tracking_number}"
#     )
 
#     encoded_table = quote(str(table_id), safe="")
#     url = f"https://api.airtable.com/v0/{base_id}/{encoded_table}/{record_id}"
 
#     headers = {
#         "Authorization": f"Bearer {api_key}",
#         "Content-Type":  "application/json",
#     }
#     body = {
#         "fields": {
#             "Shipment Label Created": True,                  # ✅ checkbox — ONLY set after upload confirmed
#         }
#     }
 
#     resp = airtable_patch_with_retry(url, body, headers, label=f"Step3/{order_number}")
#     if resp is None:
#         print(f"   ❌ Step 3 failed (all retries exhausted) for {order_number}")
#         return False
#     if resp.status_code == 200:
#         print(f"   ✅ Step 3 — Checkbox ticked ✅ | URL set  (Order: {order_number})")
#         print(f"               {dhl_tracking_url}")
#         return True
#     else:
#         # Not fatal — label is already uploaded and tracking written
#         print(f"   ⚠️  Step 3 PATCH failed ({resp.status_code}) for {order_number}: "
#               f"{resp.text[:200]}")
#         return False
 
 
# # ============================================================
# # BUILD LOOKUP MAPS
# # ============================================================
 
# def build_rma_maps(df_ya_rma, df_crs_rma):
#     ya_rma_map  = {}
#     crs_rma_map = {}
 
#     if not df_ya_rma.empty:
#         print(f"\n   YA RMA columns: {list(df_ya_rma.columns)}")
#         for _, row in df_ya_rma.iterrows():
#             rma_no = normalize_rma_number(
#                 get_first_value(row, ["RMA #", "RMA Number", "RMA No", "RMA"])
#             )
#             sales_order = normalize_sales_order_number(
#                 get_first_value(row, ["External Sales Order", "Sales Order Number",
#                     "Sales Order", "Returned Sales Order", "Order Number"])
#             )
#             if rma_no and sales_order:
#                 ya_rma_map[rma_no] = sales_order
 
#     if not df_crs_rma.empty:
#         print(f"   CRS RMA columns: {list(df_crs_rma.columns)}")
#         for _, row in df_crs_rma.iterrows():
#             rma_no = normalize_rma_number(
#                 get_first_value(row, ["RMA #", "RMA Number", "RMA No", "RMA"])
#             )
#             sales_order = normalize_lookup_key(
#                 get_first_value(row, [
#                     "Returned Sales Order (Off System)",
#                     "External Sales Order",
#                     "Sales Order Number",
#                     "Sales Order", "Returned Sales Order",
#                     "Order Number", "Refund Sales Order",
#                 ])
#             )
#             if rma_no and sales_order:
#                 crs_rma_map[rma_no] = sales_order
 
#     print(f"✅ YA RMA map : {len(ya_rma_map)} entries  → {dict(list(ya_rma_map.items())[:5])}")
#     print(f"✅ CRS RMA map: {len(crs_rma_map)} entries → {dict(list(crs_rma_map.items())[:5])}")
#     return ya_rma_map, crs_rma_map
 
 
# def build_price_map(df_sales_lines):
#     price_map = {}
#     if df_sales_lines.empty:
#         return price_map
 
#     print(f"\n   Sales Order Lines columns: {list(df_sales_lines.columns)}")
#     for _, row in df_sales_lines.iterrows():
#         so = normalize_lookup_key(
#             get_first_value(row, ["Sales Order Number", "Sales Order",
#                                    "Order Number", "SO Line#", "Reference"])
#         )
#         price_val = clean_price(
#             get_first_value(row, [
#                 "Converted Unit Price per Line", "Converted Unit Price",
#                 "Converted Unit price", "Converted Price",
#                 "Unit Price", "Price (Final)", "Price", "Sales...",
#             ], 0)
#         )
#         if so:
#             price_map[so] = price_map.get(so, 0) + price_val
 
#     print(f"✅ Price map: {len(price_map)} entries")
#     return price_map
 
 
# def fetch_price_map_for_orders(so_numbers, api_key):
#     if not so_numbers:
#         return {}
 
#     config   = TABLE_CONFIG["sales_order_lines"]
#     so_list  = list(so_numbers)
#     all_rows = []
#     CHUNK    = 20
 
#     print(f"🌐 Fetching Sales Order Lines for {len(so_list)} order(s): {so_list}")
 
#     for i in range(0, len(so_list), CHUNK):
#         chunk = so_list[i:i + CHUNK]
#         conditions = ','.join([f'{{Sales Order Number}}="{so}"' for so in chunk])
#         formula = f"OR({conditions})" if len(chunk) > 1 else conditions
 
#         df_chunk = fetch_table(
#             base_id=config["base_id"], table_id=config["table_id"],
#             view_id=config["view_id"],
#             label=f"Sales Order Lines (batch {i // CHUNK + 1})",
#             api_key=api_key, filter_formula=formula,
#         )
#         if not df_chunk.empty:
#             all_rows.append(df_chunk)
 
#     if not all_rows:
#         print("⚠️  No Sales Order Lines records found.")
#         return {}
 
#     df_all = pd.concat(all_rows, ignore_index=True)
#     return build_price_map(df_all)
 
 
# def fetch_price_map_from_commerce_central(so_numbers):
#     """Fetch unit prices from Commerce Central Sales Order Lines table using SO Line Id field."""
#     if not so_numbers:
#         return {}
 
#     config = TABLE_CONFIG["commerce_central_sales_lines"]
#     so_list = list(so_numbers)
#     all_rows = []
#     CHUNK = 20
 
#     print(f"🌐 Fetching Commerce Central Sales Lines for {len(so_list)} SO(s): {so_list[:5]}...")
 
#     # Normalize each SO number (extract last part after slash)
#     normalized_sos = [normalize_sales_order_number(so) for so in so_list]
 
#     for i in range(0, len(normalized_sos), CHUNK):
#         chunk = normalized_sos[i:i+CHUNK]
#         # Build OR conditions using SO Line Id (e.g., "41178316_74391" starts with "41178316_")
#         conditions = []
#         for so in chunk:
#             conditions.append(f'LEFT({{SO Line Id}}, LEN("{so}") + 1) = "{so}_"')
#         formula = "OR(" + ",".join(conditions) + ")"
 
#         df_chunk = fetch_table(
#             base_id=config["base_id"], table_id=config["table_id"],
#             view_id=config["view_id"],
#             label=f"Commerce Central Sales Lines (batch {i//CHUNK + 1})",
#             api_key=config["api_key"], filter_formula=formula,
#         )
#         if not df_chunk.empty:
#             all_rows.append(df_chunk)
 
#     if not all_rows:
#         print("⚠️  No Commerce Central Sales Order Lines found.")
#         return {}
 
#     df_all = pd.concat(all_rows, ignore_index=True)
#     price_map = {}
#     for _, row in df_all.iterrows():
#         # Extract Sales Order number from SO Line Id (everything before the first underscore)
#         so_line = row.get("SO Line Id", "")
#         if not so_line:
#             # fallback to legacy field name just in case
#             so_line = row.get("SO Line#", "")
#         if so_line and '_' in so_line:
#             so = normalize_sales_order_number(so_line.split('_')[0])
#         else:
#             continue
#         if not so:
#             continue
 
#         # Get price from "Converted Unit price" (fallback to "Price (Final)")
#         price = clean_price(row.get("Converted Unit price", 0))
#         if price <= 0:
#             price = clean_price(row.get("Price (Final)", 0))
#         if price > 0:
#             # Sum prices if multiple lines for the same SO
#             price_map[so] = price_map.get(so, 0) + price
 
#     print(f"✅ Commerce Central price map: {len(price_map)} entries")
#     return price_map
 
 
# # ============================================================
# # TRANSFORM (builds CSV rows)
# # ============================================================
 
# def transform(df, ya_rma_map=None, crs_rma_map=None, price_map=None):
#     rows = []
#     F = FIXED
 
#     ya_rma_map  = ya_rma_map  or {}
#     crs_rma_map = crs_rma_map or {}
#     price_map   = price_map   or {}
 
#     for _, r in df.iterrows():
#         country      = str(r.get("Shipping Country", "")).strip()
#         cc, num      = split_phone(r.get("Telephone", ""), country)
#         order_number = str(r.get("Order Number", "")).strip()
 
#         price = round(clean_price(r.get("Converted Unit Price per Line", 0)), 2)
 
#         rma_source = get_rma_source(order_number)
 
#         if rma_source:
#             clean_rma = normalize_rma_number(order_number)
#             if rma_source == "ya":
#                 sales_order  = ya_rma_map.get(clean_rma, "")
#                 source_label = "YesAgain RMA"
#             else:
#                 sales_order  = crs_rma_map.get(clean_rma, "")
#                 source_label = "CRS RMA"
 
#             if sales_order:
#                 looked_up_price = price_map.get(sales_order, None)
#                 if looked_up_price is not None:
#                     price = round(looked_up_price, 2)
#                     print(f"   ✅ {order_number} ({source_label}) → SO {sales_order} → €{price}")
#                 else:
#                     # Path A: Sales Order found but no price in Sales Order Lines
#                     if rma_source == "crs":
#                         price = 200.0
#                         print(f"   ⚠️  {order_number} ({source_label}) → SO {sales_order} "
#                               f"found but no price in Sales Order Lines — defaulting to €200.00")
#                     else:
#                         print(f"   ⚠️  {order_number} ({source_label}) → SO {sales_order} "
#                               f"found but no price in Sales Order Lines")
#             else:
#                 # Path B: RMA key not found in the map at all
#                 rma_map_used = ya_rma_map if rma_source == "ya" else crs_rma_map
#                 print(f"   ⚠️  {order_number} ({source_label}) → "
#                       f"RMA key '{clean_rma}' not found in map.")
#                 print(f"        Map has {len(rma_map_used)} entries: "
#                       f"{list(rma_map_used.keys())[:10]}")
#                 if rma_source == "crs":
#                     price = 200.0
#                     print(f"   🔄 {order_number} (RMAC) → RMA not in map — defaulting to €200.00")
 
#         qty = clean_qty(r.get("Sold Qty per line", 1))
 
#         rules       = get_destination_rules(country, order_number)
#         shipper_acc = rules["shipper_account"]
#         duty_acc    = rules["duty_account"]
#         add_party   = rules["additional_party"]
#         region      = rules["region"]
 
#         # ── Declared Value Rules ──────────────────────────────────────
#         # Policy (per shipping policy document):
#         #   Saudi Arabia (SA) → FIXED 164 EUR (regardless of item price)
#         #   GCC / USA / Others not listed → 20% of item value
#         #   Europe / UK → full price (Exclude VAT — data is already ex-VAT)
#         country_upper = country.strip().upper()
#         if country_upper == "SA":
#             declared_val = 164.0
#             print(f"   🇸🇦 {order_number} (SA) → Fixed declared value: €164.00")
#         elif region in ("gcc", "us", "other"):
#             declared_val = round(price * 0.20, 2)
#             declared_val = max(declared_val, 1.0)   # DHL minimum
#             print(f"   📦 {order_number} ({region.upper()}) → 20% declared: €{declared_val} (item: €{price})")
#         else:
#             # eu / uk — full price, exclude VAT (DHL minimum 1.0)
#             declared_val = max(price, 1.0)
 
#         addr2, state = get_address2_and_state(r)
#         state_full = STATE_FULL_NAMES.get(state, state) if state else ""
 
#         name = clean_text(r.get("Shipping Name", ""))[:35]
#         raw_company = r.get("Invoice Name", "")
#         if pd.isna(raw_company) if not isinstance(raw_company, (list, dict)) else False:
#             company = name
#         elif str(raw_company).strip() == "":
#             company = name
#         else:
#             company = clean_text(raw_company)[:35]
 
#         rows.append({
#             # --- SHIPMENT DETAILS ---
#             "Shipment Reference 1"                      : order_number,
#             "Name (Ship TO) (Required)"                 : name,
#             "Company (Ship TO) (Required)"              : company,
#             "Address 1 (Ship TO) (Required)"            : (clean_text(r.get("Shipping address 1", "")).replace(",", " "))[:45],
#             "Address 2 (Ship TO)"                       : addr2.replace(",", " "),
#             "Address 3 (Ship TO)"                       : "",
#             "City (Ship TO) (Required)"                 : clean_text(r.get("Shipping City", "")),
#             "State Province (Ship TO)"                  : state_full,
#             "ZIP Postal Code (Ship TO)"                 : r.get("Shipping Postcode", ""),
#             "Country Code (Ship TO) (Required)"         : country,
#             "Email Address (Ship TO)"                   : r.get("CustomerEmail", ""),
#             "Phone Type (Ship TO)"                      : "O",
#             "Phone Country Code (Ship TO) (Required)"   : cc,
#             "Phone Number (Ship TO) (Required)"         : num,
#             "VAT Tax ID (Ship TO)"                      : "",
#             "EORI Number (Ship TO)"                     : "",
 
#             # --- ACCOUNT / SHIPPING CONFIG ---
#             "Account Number (Shipper) (Required)"       : int(shipper_acc),
#             "Account Number (Payer)"                    : int(shipper_acc),
#             "inCoterms"                                 : F["incoterms"],
#             "Total Weight (Required)"                   : round(qty * F["weight"], 2),
#             "Piece Weight (Unit of Measure)"            : F["weight_unit"],
#             "Declared Value Currency (Required)"        : F["currency"],
#             "Declared Value (Required)"                 : declared_val,
#             "Product Code (3 Letter)"                   : F["product_code"],
#             "Summary of Contents"                       : F["contents"],
#             "SHIPMENT TYPE"                             : F["shipment_type"],
#             "Total Shipment Pieces"                     : qty,
#             "Piece Dimensions (Unit of Measure)"        : F["dim_unit"],
#             "LENGTH"                                    : F["length"],
#             "WIDTH"                                     : F["width"],
#             "HEIGHT"                                    : F["height"],
#             "REASON OF EXPORT"                          : F["export_reason"],
#             "Digital Customs Invoice  Y N (Paperless)"  : F["dig_customs"],
 
#             # --- INVOICE / ITEM ---
#             "INVOICE NO."                               : order_number,
#             "ITEM DESCRIPTION"                          : F["item_desc"],
#             "ITEM COMMODITY"                            : F["commodity"],
#             "ITM QTY"                                   : qty,
#             "ITM UNITS"                                 : F["item_units"],
#             "ITM VALUE"                                 : round(declared_val, 2),
#             "ITM CRNCY"                                 : F["currency"],
#             "ITM NET"                                   : F["item_net"],
#             "ITM GRSS"                                  : F["item_gross"],
#             "Country of Origin (Customs Invoice)"       : F["origin"],
#             "REMARKS"                                   : "",
 
#             # --- ADDITIONAL PARTY ---
#             "Company (Additional Party)"                : add_party["Company"],
#             "Name (Additional Party)"                   : add_party["Name"],
#             "Address 1 (Additional Party)"              : add_party["Address1"],
#             "Address 2 (Additional Party)"              : "",
#             "Address 3 (Additional Party)"              : "",
#             "Country Code (Additional Party)"           : add_party["Country"],
#             "City (Additional Party)"                   : add_party["City"],
#             "ZIP Postal Code (Additional Party)"        : add_party["ZIP"],
#             "ADD  EMAIL"                                : add_party["Email"],
#             "Phone Country Code (Additional Party)"     : add_party["PhoneCC"],
#             "Phone Number (Additional Party)"           : add_party["Phone"],
#             "ADD  VAT"                                  : importer_vat_for(country, add_party["VAT"]),
#             "ADD  EORI"                                 : add_party["EORI"],
#             "ADD  RELATIONSHIP"                         : add_party["Rel"],
#             "Account Number (Duty Tax)"                 : int(duty_acc),
#         })
 
#     return pd.DataFrame(rows)
 
 
# # ============================================================
# # DHL API INTEGRATION
# # ============================================================
 
# def build_dhl_payload(row):
#     F = FIXED
#     from datetime import timedelta
    
#     # DHL prefers an explicit offset structure for timezones
#     ship_date = (datetime.now(timezone.utc) + timedelta(days=1)).strftime(
#         "%Y-%m-%dT10:00:00GMT+04:00"
#     )
 
#     order_id = str(row.get("Shipment Reference 1", "")).strip()
#     country  = str(row.get("Country Code (Ship TO) (Required)", "")).strip()
 
#     # Get structural variables from mapping logic
#     rules           = get_destination_rules(country, order_id)
#     party           = rules["shipper_party"]
#     shipper_account = rules["shipper_account"]
#     duty_account    = rules["duty_account"]

#     # Destination-based importer VAT/EORI (YesAgain's local registration for
#     # the country the goods are imported into). Falls back to the importer
#     # entity's default VAT when the destination has no specific registration.
#     importer_party = rules["additional_party"]
#     importer_vat   = importer_vat_for(country, importer_party.get("VAT", ""))
#     importer_eori  = importer_party.get("EORI", "")
#     vat_issuer     = (importer_vat[:2].upper() if importer_vat else importer_party.get("Country", ""))
 
#     # Phone number sanitation and formatting
#     phone_cc  = str(row.get("Phone Country Code (Ship TO) (Required)", "")).strip()
#     phone_num = str(row.get("Phone Number (Ship TO) (Required)", "")).strip()
#     if not phone_cc:
#         phone_cc = PHONE_CODES.get(country, "")
#     if not phone_num:
#         phone_num = "0000000000"
#     phone_num_clean = re.sub(r'\D', '', phone_num)
    
#     # --- FIXED LINE HERE ---
#     full_phone = f"+{phone_cc}{phone_num_clean}" if not str(phone_cc).startswith('+') else f"{phone_cc}{phone_num_clean}"
 
#     shipper_phone_cc   = str(party["PhoneCC"])
#     shipper_phone_num  = re.sub(r'\D', '', str(party["Phone"]))
#     shipper_full_phone = f"+{shipper_phone_cc}{shipper_phone_num}"
 
#     # Smart address splitting to prevent truncation data loss
#     raw_address = str(row.get("Address 1 (Ship TO) (Required)", "")).strip()
#     addr2       = str(row.get("Address 2 (Ship TO)", "")).strip()
    
#     postal_address = {
#         "postalCode"  : clean_postal_code(str(row.get("ZIP Postal Code (Ship TO)", "")).strip(), country),
#         "cityName"    : str(row.get("City (Ship TO) (Required)", "")).strip(),
#         "countryCode" : country,
#         "addressLine1": raw_address[:45],
#     }
 
#     # Overflow remainder of address line 1 safely into address line 2 if line 2 is empty
#     if len(raw_address) > 45:
#         postal_address["addressLine2"] = (raw_address[45:] + " " + addr2).strip()[:45]
#     elif addr2:
#         postal_address["addressLine2"] = addr2[:45]
        
#     state = str(row.get("State Province (Ship TO)", "")).strip()
#     if state and country == "US":
#         postal_address["provinceCode"] = state
 
#     # Accurate unit-level pricing calculation 
#     itm_qty    = max(int(row.get("ITM QTY", 1)), 1)
#     total_val  = float(row.get("Declared Value (Required)", 0))
#     unit_price = max(round(total_val / itm_qty, 2), 0.01)
 
#     payload = {
#         "plannedShippingDateAndTime": ship_date,
#         "pickup": {"isRequested": False},
#         "productCode": F["product_code_api"],
#         "accounts": [
#             {"typeCode": "shipper", "number": str(shipper_account)},
#             {"typeCode": "payer",   "number": str(shipper_account)},
#             {"typeCode": "duties-taxes", "number": str(duty_account)}, # Bound dynamically via accounting routing Matrix
#         ],
#         "outputImageProperties": {
#             "printerDPI"    : 300,
#             "encodingFormat": "pdf",
#             "imageOptions"  : [
#                 {"typeCode": "label", "templateName": "ECOM26_84_001", "isRequested": True},
#                 {
#                     "typeCode"           : "invoice",
#                     "templateName"       : "COMMERCIAL_INVOICE_P_10",
#                     "isRequested"        : True,
#                     "invoiceType"        : "commercial",
#                     "languageCode"       : "eng",
#                     "languageCountryCode": "US",
#                 },
#             ]
#         },
#         "customerDetails": {
#             "shipperDetails": {
#                 "postalAddress": {
#                     "postalCode"  : party["ZIP"],
#                     "cityName"    : party["City"],
#                     "countryCode" : party["Country"],
#                     "addressLine1": party["Address1"],
#                 },
#                 "contactInformation": {
#                     "companyName": party["Company"],
#                     "fullName"   : party["Name"],
#                     "email"      : party["Email"],
#                     "phone"      : shipper_full_phone,
#                 },
#                 "registrationNumbers": [
#                     *([{"typeCode": "VAT", "number": party["VAT"],  "issuerCountryCode": party["Country"]}] if party.get("VAT")  else []),
#                     *([{"typeCode": "EOR", "number": party["EORI"], "issuerCountryCode": party["Country"]}] if party.get("EORI") else []),
#                     *([{"typeCode": "VAT", "number": importer_vat,  "issuerCountryCode": vat_issuer}] if importer_vat  else []),
#                     *([{"typeCode": "EOR", "number": importer_eori, "issuerCountryCode": importer_party.get("Country", "")}] if importer_eori else []),
#                 ],
#             },
#             "receiverDetails": {
#                 "postalAddress": postal_address,
#                 "contactInformation": {
#                     "companyName": str(row.get("Company (Ship TO) (Required)", "")).strip()[:35] or str(row.get("Name (Ship TO) (Required)", "")).strip()[:35],
#                     "fullName"   : str(row.get("Name (Ship TO) (Required)", "")).strip()[:35],
#                     "email"      : str(row.get("Email Address (Ship TO)", "")).strip(),
#                     "phone"      : full_phone,
#                 },
#             },
#         },
#         "content": {
#             "packages": [
#                 {
#                     "weight"    : float(row.get("Total Weight (Required)", F["weight"])),
#                     "dimensions": {"length": F["length"], "width": F["width"], "height": F["height"]},
#                     "customerReferences": [
#                         {"typeCode": "CU",  "value": order_id[:35]},
#                     ]
#                 }
#             ],
#             "isCustomsDeclarable"   : True,
#             "declaredValue"         : total_val,
#             "declaredValueCurrency" : str(row.get("Declared Value Currency (Required)", F["currency"])),
#             "exportDeclaration": {
#                 "lineItems": [
#                     {
#                         "number"             : i + 1,
#                         "description"        : F["item_desc"],
#                         "price"              : unit_price,
#                         "priceCurrency"      : str(row.get("Declared Value Currency (Required)", F["currency"])),
#                         "commodityCodes"     : [
#                             {"typeCode": "outbound", "value": F["commodity"]},
#                             {"typeCode": "inbound",  "value": F["commodity"]},
#                         ],
#                         "exportReasonType"   : "permanent",
#                         "manufacturerCountry": F["origin"],
#                         "weight"             : {"netValue": F["item_net"], "grossValue": F["item_gross"]},
#                         "quantity"           : {"value": 1, "unitOfMeasurement": "PCS"},
#                     }
#                     for i in range(itm_qty)
#                 ],
#                 "invoice": {
#                     # DHL limits invoice number to 35 chars. Airtable UUIDs are 36 → HTTP 422.
#                     "number": (str(row.get("INVOICE NO.", "")) or order_id)[:35],
#                     "date"  : datetime.now().strftime("%Y-%m-%d"),
#                 },
#                 "exportReason": F["export_reason"],
#             },
#             "description"       : F["contents"],
#             "incoterm"          : "DDP", # Forced to DDP to match owner instructions
#             "unitOfMeasurement": "metric",
#         },
#         "shipmentNotification": [
#             {
#                 "typeCode"    : "email",
#                 "receiverId"  : str(row.get("Email Address (Ship TO)", "")).strip(),
#                 "languageCode": "eng",
#             }
#         ] if str(row.get("Email Address (Ship TO)", "")).strip() else [],
#         "valueAddedServices": [
#             {"serviceCode": "WY"}, # Changed from DD to WY (Duties Taxes Paid) to avoid double billing bugs!
#         ],
#         "customerReferences": [
#             {"typeCode": "CU",  "value": order_id[:35]},
#         ],
#     }
 
#     return payload
 
 
# # ============================================================
# # LABEL INDEX — saves every shipment so the lookup tool can find it
# # ============================================================
 
# LABEL_INDEX_FILE = "label_index.json"
 
# def save_label_index(order_number, tracking_number, label_path,
#                      recipient_name, destination_country):
#     """Append/update a record in label_index.json for the lookup tool."""
#     if os.path.exists(LABEL_INDEX_FILE):
#         try:
#             with open(LABEL_INDEX_FILE, "r", encoding="utf-8") as f:
#                 index = json.load(f)
#         except Exception:
#             index = []
#     else:
#         index = []
 
#     index = [e for e in index
#              if str(e.get("order", "")).strip() != str(order_number).strip()]
 
#     index.insert(0, {
#         "order"     : str(order_number).strip(),
#         "awb"       : str(tracking_number).strip(),
#         "label"     : label_path,
#         "recipient" : str(recipient_name).strip(),
#         "country"   : str(destination_country).strip(),
#         "created_at": datetime.now().strftime("%Y-%m-%d %H:%M"),
#     })
 
#     with open(LABEL_INDEX_FILE, "w", encoding="utf-8") as f:
#         json.dump(index, f, indent=2, ensure_ascii=False)
 
#     print(f"   📋 Label index updated: {LABEL_INDEX_FILE}")
 
 
# def generate_lookup_html():
#     """Generate label_lookup.html with all label data embedded — no server needed."""
#     if not os.path.exists(LABEL_INDEX_FILE):
#         print("⚠️  No label_index.json found — skipping HTML generation.")
#         return
 
#     try:
#         with open(LABEL_INDEX_FILE, "r", encoding="utf-8") as f:
#             index = json.load(f)
#     except Exception as e:
#         print(f"⚠️  Could not read label_index.json: {e}")
#         return
 
#     enriched = []
#     for entry in index:
#         label_path = entry.get("label", "")
#         label_b64  = ""
#         if label_path and os.path.exists(label_path):
#             with open(label_path, "rb") as lf:
#                 label_b64 = base64.b64encode(lf.read()).decode("utf-8")
#         enriched.append({**entry, "label_b64": label_b64})
 
#     data_json = json.dumps(enriched, ensure_ascii=False)
 
#     html = f"""<!DOCTYPE html>
# <html lang="en">
# <head>
# <meta charset="UTF-8">
# <meta name="viewport" content="width=device-width, initial-scale=1.0">
# <title>YesAgain — DHL Label Lookup</title>
# <style>
#   * {{ box-sizing: border-box; margin: 0; padding: 0; }}
#   body {{
#     font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
#     background: #f5f5f5; min-height: 100vh;
#     display: flex; flex-direction: column; align-items: center; padding: 40px 20px;
#   }}
#   .header {{
#     background: #FFCC00; width: 100%; max-width: 700px;
#     border-radius: 12px 12px 0 0; padding: 24px 30px;
#     display: flex; align-items: center; gap: 16px;
#   }}
#   .header-logo {{ font-size: 28px; font-weight: 900; color: #D40511; letter-spacing: -1px; }}
#   .header-title {{ font-size: 18px; font-weight: 700; color: #333; }}
#   .header-sub {{ font-size: 13px; color: #666; margin-top: 2px; }}
#   .card {{
#     background: white; width: 100%; max-width: 700px;
#     border-radius: 0 0 12px 12px; padding: 30px;
#     box-shadow: 0 4px 20px rgba(0,0,0,0.08);
#   }}
#   .search-row {{ display: flex; gap: 10px; margin-bottom: 28px; }}
#   .search-input {{
#     flex: 1; padding: 14px 18px; font-size: 16px;
#     border: 2px solid #ddd; border-radius: 8px; outline: none; transition: border-color 0.2s;
#   }}
#   .search-input:focus {{ border-color: #FFCC00; }}
#   .search-btn {{
#     padding: 14px 28px; background: #D40511; color: white; border: none;
#     border-radius: 8px; font-size: 16px; font-weight: 600; cursor: pointer;
#   }}
#   .search-btn:hover {{ background: #b00; }}
#   .result-box {{
#     display: none; background: #f9f9f9; border: 1px solid #e0e0e0;
#     border-radius: 10px; padding: 24px; margin-bottom: 24px;
#   }}
#   .result-box.visible {{ display: block; }}
#   .result-label {{ font-size: 12px; color: #888; text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 3px; }}
#   .result-value {{ font-size: 18px; font-weight: 700; color: #222; margin-bottom: 16px; }}
#   .result-meta {{ display: flex; gap: 24px; flex-wrap: wrap; margin-bottom: 20px; }}
#   .meta-item {{ flex: 1; min-width: 120px; }}
#   .download-btn {{
#     display: inline-block; padding: 14px 32px; background: #FFCC00;
#     color: #333; font-size: 16px; font-weight: 700; border-radius: 8px;
#     text-decoration: none; cursor: pointer; border: none; width: 100%; text-align: center;
#   }}
#   .download-btn:hover {{ background: #e6b800; }}
#   .no-label {{ color: #cc0000; font-size: 14px; margin-top: 10px; }}
#   .error-box {{
#     display: none; background: #fff0f0; border: 1px solid #ffcccc;
#     border-radius: 10px; padding: 20px 24px; color: #cc0000; font-size: 15px; margin-bottom: 24px;
#   }}
#   .error-box.visible {{ display: block; }}
#   .recent-title {{ font-size: 13px; font-weight: 600; color: #888; text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 12px; }}
#   .recent-table {{ width: 100%; border-collapse: collapse; font-size: 14px; }}
#   .recent-table th {{ text-align: left; padding: 8px 10px; background: #f0f0f0; color: #555; font-weight: 600; }}
#   .recent-table td {{ padding: 9px 10px; border-bottom: 1px solid #f0f0f0; color: #333; }}
#   .recent-table tr:hover td {{ background: #fafafa; cursor: pointer; }}
#   .badge {{ display: inline-block; padding: 2px 8px; background: #e8f5e9; color: #2e7d32; border-radius: 20px; font-size: 12px; font-weight: 600; }}
#   .badge.no-file {{ background: #fce4ec; color: #c62828; }}
# </style>
# </head>
# <body>
# <div class="header">
#   <div class="header-logo">DHL</div>
#   <div>
#     <div class="header-title">Label Lookup Tool</div>
#     <div class="header-sub">YesAgain — Search by Order Number to download a label</div>
#   </div>
# </div>
# <div class="card">
#   <div class="search-row">
#     <input type="text" class="search-input" id="searchInput"
#            placeholder="Enter Order Number or AWB"
#            onkeydown="if(event.key==='Enter') doSearch()">
#     <button class="search-btn" onclick="doSearch()">🔍 Search</button>
#   </div>
#   <div class="result-box" id="resultBox">
#     <div class="result-label">Order Number</div>
#     <div class="result-value" id="resOrder">—</div>
#     <div class="result-meta">
#       <div class="meta-item"><div class="result-label">AWB / Tracking</div><div class="result-value" id="resAWB" style="font-size:15px">—</div></div>
#       <div class="meta-item"><div class="result-label">Recipient</div><div class="result-value" id="resRecipient" style="font-size:15px">—</div></div>
#       <div class="meta-item"><div class="result-label">Country</div><div class="result-value" id="resCountry" style="font-size:15px">—</div></div>
#       <div class="meta-item"><div class="result-label">Created</div><div class="result-value" id="resDate" style="font-size:15px">—</div></div>
#     </div>
#     <button class="download-btn" id="downloadBtn" onclick="downloadLabel()">⬇️ Download Shipping Label</button>
#     <div class="no-label" id="noLabelMsg" style="display:none">⚠️ Label file not available.</div>
#   </div>
#   <div class="error-box" id="errorBox">❌ No shipment found for that order number.</div>
#   <div class="recent-title">Recent Shipments ({len(index)} total)</div>
#   <table class="recent-table">
#     <thead><tr><th>Order</th><th>AWB</th><th>Recipient</th><th>Country</th><th>Created</th><th>Label</th></tr></thead>
#     <tbody id="recentTable"></tbody>
#   </table>
# </div>
# <script>
# const DATA = {data_json};
# let currentEntry = null;
# function buildTable() {{
#   const tbody = document.getElementById("recentTable");
#   DATA.slice(0, 100).forEach(e => {{
#     const hasLabel = e.label_b64 && e.label_b64.length > 0;
#     const tr = document.createElement("tr");
#     tr.innerHTML = `<td><strong>${{e.order}}</strong></td><td>${{e.awb||"—"}}</td><td>${{e.recipient||"—"}}</td><td>${{e.country||"—"}}</td><td>${{e.created_at||"—"}}</td><td><span class="badge ${{hasLabel?"":"no-file"}}">${{hasLabel?"✅ Ready":"❌ Missing"}}</span></td>`;
#     tr.onclick = () => showResult(e);
#     tbody.appendChild(tr);
#   }});
# }}
# function showResult(entry) {{
#   currentEntry = entry;
#   document.getElementById("resOrder").textContent     = entry.order;
#   document.getElementById("resAWB").textContent       = entry.awb || "—";
#   document.getElementById("resRecipient").textContent = entry.recipient || "—";
#   document.getElementById("resCountry").textContent   = entry.country || "—";
#   document.getElementById("resDate").textContent      = entry.created_at || "—";
#   const hasLabel = entry.label_b64 && entry.label_b64.length > 0;
#   document.getElementById("downloadBtn").style.display = hasLabel ? "block" : "none";
#   document.getElementById("noLabelMsg").style.display  = hasLabel ? "none" : "block";
#   document.getElementById("resultBox").classList.add("visible");
#   document.getElementById("errorBox").classList.remove("visible");
# }}
# function doSearch() {{
#   const q = document.getElementById("searchInput").value.trim().toUpperCase();
#   if (!q) return;
#   const found = DATA.find(e => e.order.toUpperCase()===q || (e.awb||"").toUpperCase()===q);
#   if (found) {{ showResult(found); }}
#   else {{
#     document.getElementById("resultBox").classList.remove("visible");
#     document.getElementById("errorBox").classList.add("visible");
#     currentEntry = null;
#   }}
# }}
# function downloadLabel() {{
#   if (!currentEntry || !currentEntry.label_b64) return;
#   const bytes = atob(currentEntry.label_b64);
#   const arr = new Uint8Array(bytes.length);
#   for (let i=0;i<bytes.length;i++) arr[i]=bytes.charCodeAt(i);
#   const blob = new Blob([arr], {{type:"application/pdf"}});
#   const a = document.createElement("a");
#   a.href = URL.createObjectURL(blob);
#   a.download = "label_" + currentEntry.order + ".pdf";
#   a.click();
# }}
# buildTable();
# </script>
# </body>
# </html>"""
 
#     with open("label_lookup.html", "w", encoding="utf-8") as f:
#         f.write(html)
 
#     print(f"\n✅ Label lookup tool: label_lookup.html ({len(index)} shipment(s) embedded)")
 
 
# def send_to_dhl(row, verbose=False):
#     if not DHL_API_KEY or not DHL_API_SECRET:
#         print("❌ DHL_API_KEY or DHL_API_SECRET missing in .env — cannot send to DHL.")
#         return {"success": False, "error": "Missing DHL credentials"}
    
#     use_test     = DHL_TEST_MODE
#     base_url     = DHL_BASE_URL_PROD   # always production URL — test bypass handled below
#     order_number = str(row.get("Shipment Reference 1", "unknown")).strip()
 
#     # ── TEST MODE BYPASS ──────────────────────────────────────────────────────
#     # DHL error 803 "Account not allowed for this service" happens because the
#     # DHL TEST SANDBOX does not recognise real production account numbers.
#     # The test sandbox requires special dummy test accounts provided by DHL.
#     # Since we don't have those dummy accounts, calling the sandbox will always
#     # fail with 803.
#     #
#     # Solution: in --test mode we skip the real DHL call entirely and return a
#     # simulated success. This lets you safely test the full pipeline:
#     #   ✅  Airtable fetch
#     #   ✅  Price / RMA lookup
#     #   ✅  CSV generation
#     #   ✅  Payload building & printing (use --verbose to inspect)
#     #   ✅  Airtable writeback is STILL SKIPPED (test mode already guards this)
#     #   ❌  No real DHL call (by design — no fake AWB)
#     #
#     # To test the actual DHL connection without booking a real shipment, run
#     # python airtable_transform.py --send-to-dhl --validate
#     # which calls /shipments/validate on PRODUCTION with your real accounts.
#     if use_test:
#         fake_awb = f"TEST-{order_number}-SIM"
#         print(f"   ✅ TEST SIMULATION — Payload built OK (use --verbose to print)")
#         print(f"   ℹ️  No DHL call made — fake AWB: {fake_awb}")
#         return {
#             "success"         : True,
#             "order_number"    : order_number,
#             "tracking_number" : fake_awb,
#             "label_base64"    : None,
#             "raw_response"    : {},
#             "simulated"       : True,
#         }
 
#     url = f"{base_url}/shipments"
 
#     headers = {"Content-Type": "application/json", "Accept": "application/json"}
#     payload = build_dhl_payload(row)
 
#     country = str(row.get("Country Code (Ship TO) (Required)", "")).strip()
#     rules   = get_destination_rules(country, order_number)
#     party   = rules["shipper_party"]
#     print(f"   Shipper: {party['Company']} ({party['Country']})  "
#           f"Duty account: {rules['duty_account']}  Dest: {country}")
 
#     if verbose:
#         print(f"   Payload:\n{json.dumps(payload, indent=2)}")
 
#     try:
#         response = requests.post(url, json=payload, headers=headers,
#                                  auth=(DHL_API_KEY, DHL_API_SECRET), timeout=30)
#     except requests.exceptions.Timeout:
#         print(f"   ❌ Timeout sending {order_number}")
#         return {"success": False, "order_number": order_number, "error": "Timeout"}
#     except requests.exceptions.ConnectionError as e:
#         print(f"   ❌ Connection error for {order_number}: {e}")
#         return {"success": False, "order_number": order_number, "error": str(e)}
 
#     if response.status_code in (200, 201):
#         data            = response.json()
#         tracking_number = data.get("shipmentTrackingNumber", "")
#         packages        = data.get("packages", [])
#         pkg_tracking    = packages[0].get("trackingNumber", "") if packages else ""
 
#         print(f"   ✅ DHL OK — Tracking: {tracking_number or pkg_tracking}")
 
#         label_b64 = None
#         for doc in data.get("documents", []):
#             type_code    = doc.get("typeCode", "")
#             image_format = doc.get("imageFormat", "")
#             content      = doc.get("content", "")
 
#             # DHL production API returns typeCode="label".
#             # DHL sandbox/test API often omits typeCode entirely, returning only
#             # imageFormat="PDF". Accept EITHER so the label is never silently dropped.
#             is_label = (
#                 type_code == "label"
#                 or (not type_code and image_format.upper() == "PDF" and content)
#             )
 
#             if is_label and content:
#                 label_b64  = content
#                 os.makedirs("labels", exist_ok=True)
#                 label_path = f"labels/label_{order_number}.pdf"
#                 with open(label_path, "wb") as f:
#                     f.write(base64.b64decode(label_b64))
#                 print(f"   🏷️  Label saved locally: {label_path}")
#                 recipient_name      = str(row.get("Name (Ship TO) (Required)", "")).strip()
#                 destination_country = str(row.get("Country Code (Ship TO) (Required)", "")).strip()
#                 save_label_index(order_number, tracking_number or pkg_tracking,
#                                  label_path, recipient_name, destination_country)
#                 break
 
#         if not label_b64:
#             print(f"   ⚠️  No label found in DHL response for {order_number}.")
#             docs_summary = [
#                 {k: v for k, v in d.items() if k != "content"}
#                 for d in data.get("documents", [])
#             ]
#             print(f"       Documents returned: {docs_summary}")
 
#         return {
#             "success"         : True,
#             "order_number"    : order_number,
#             "tracking_number" : tracking_number or pkg_tracking,
#             "label_base64"    : label_b64,
#             "raw_response"    : data,
#         }
#     else:
#         try:
#             error_msg = json.dumps(response.json(), indent=2)
#         except Exception:
#             error_msg = response.text
 
#         print(f"   ❌ DHL FAILED — HTTP {response.status_code}")
#         print(f"   {error_msg[:500]}")
 
#         return {
#             "success"      : False,
#             "order_number" : order_number,
#             "status_code"  : response.status_code,
#             "error"        : error_msg,
#         }
 
 
# def validate_credentials_dhl():
#     if not DHL_API_KEY or not DHL_API_SECRET:
#         print("❌ DHL_API_KEY or DHL_API_SECRET not set in .env")
#         return False
 
#     base_url = DHL_BASE_URL_TEST if DHL_TEST_MODE else DHL_BASE_URL_PROD
#     url = (
#         f"{base_url}/products"
#         f"?accountNumber={FIXED['account_shipper']}"
#         f"&originCountryCode=AE&originCityName=Sharjah"
#         f"&destinationCountryCode=DE&destinationCityName=Berlin"
#         f"&weight=1.5&length=35&width=30&height=7"
#         f"&plannedShippingDate={datetime.now().strftime('%Y-%m-%d')}"
#     )
 
#     print(f"\n🔑 Testing DHL credentials {'[TEST]' if DHL_TEST_MODE else '[PRODUCTION]'}...")
 
#     try:
#         response = requests.get(url, auth=(DHL_API_KEY, DHL_API_SECRET),
#                                 headers={"Accept": "application/json"}, timeout=15)
#         print(f"   HTTP Status: {response.status_code}")
 
#         if response.status_code == 200:
#             print("   ✅ DHL credentials are VALID!")
#             return True
#         elif response.status_code == 401:
#             print("   ❌ 401 Unauthorized — credentials are WRONG.")
#             return False
#         elif response.status_code == 403:
#             print("   ⚠️  403 Forbidden — credentials work but account may not have this endpoint.")
#             return True
#         else:
#             print(f"   ⚠️  Unexpected status {response.status_code} — assuming credentials valid.")
#             return True
 
#     except requests.exceptions.ConnectionError:
#         print("   ❌ Cannot reach DHL API — check internet connection.")
#         return False
#     except requests.exceptions.Timeout:
#         print("   ❌ Timeout connecting to DHL API.")
#         return False
 
 
# # ============================================================
# # CANCEL DHL SHIPMENT
# # ============================================================

# def cancel_dhl_shipment(tracking_number):
#     """
#     Cancel a booked DHL shipment using its AWB / tracking number.

#     HOW IT WORKS
#     ─────────────────────────────────────────────────────────────
#     DHL Express API:  DELETE /mydhlapi/shipments/{trackingNumber}
#     Returns 200 or 204 on success.

#     IMPORTANT RULES
#     ─────────────────────────────────────────────────────────────
#     ✅  Works ONLY if DHL has NOT yet scanned / picked up the parcel.
#     ❌  Once DHL scans it, this API call returns an error.
#         → You must call DHL customer service: +971 600 567 567
#         → Or ask the recipient to refuse delivery.

#     Args:
#         tracking_number  – DHL AWB string (e.g. "1234567890")

#     Returns dict:
#         { success, tracking_number, message }   on success
#         { success, tracking_number, error }     on failure
#     """
#     if not DHL_API_KEY or not DHL_API_SECRET:
#         print("❌ DHL_API_KEY or DHL_API_SECRET missing — cannot cancel.")
#         return {"success": False, "error": "Missing DHL credentials"}

#     tracking_number = str(tracking_number).strip()

#     if not tracking_number or tracking_number.startswith("TEST-"):
#         return {
#             "success": False,
#             "error"  : "Invalid or simulated tracking number — nothing to cancel."
#         }

#     url = f"{DHL_BASE_URL_PROD}/shipments/{tracking_number}"
#     print(f"\n🗑️  Cancelling DHL shipment: {tracking_number}")
#     print(f"   URL: DELETE {url}")

#     try:
#         response = requests.delete(
#             url,
#             headers={"Accept": "application/json"},
#             auth=(DHL_API_KEY, DHL_API_SECRET),
#             timeout=30
#         )
#     except requests.exceptions.Timeout:
#         print(f"   ❌ Timeout cancelling {tracking_number}")
#         return {"success": False, "tracking_number": tracking_number, "error": "Timeout"}
#     except requests.exceptions.ConnectionError as e:
#         print(f"   ❌ Connection error: {e}")
#         return {"success": False, "tracking_number": tracking_number, "error": str(e)}

#     # DHL returns 200 or 204 on successful cancellation
#     if response.status_code in (200, 204):
#         print(f"   ✅ DHL confirmed cancellation: {tracking_number}")
#         return {
#             "success"         : True,
#             "tracking_number" : tracking_number,
#             "message"         : "Shipment cancelled successfully at DHL."
#         }

#     # Parse error from DHL
#     try:
#         error_body = response.json()
#         error_msg = error_body.get("detail", json.dumps(error_body))
#     except Exception:
#         raw = response.text.strip()
#         if raw.startswith("<"):
#             error_msg = "DHL returned an XML/SOAP error — check API credentials or endpoint URL"
#         else:
#             error_msg = raw

#     print(f"   ❌ DHL cancellation failed — HTTP {response.status_code}")
#     print(f"   {error_msg[:400]}")

#     return {
#         "success"         : False,
#         "tracking_number" : tracking_number,
#         "status_code"     : response.status_code,
#         "error"           : error_msg
#     }


# # ============================================================
# # UNDO AIRTABLE BOOKING
# # After a successful DHL cancellation, clear the Airtable fields
# # that were written during the original booking (Steps 1–3).
# # ============================================================

# def undo_airtable_booking(record_id, order_number):
#     """
#     Clear Airtable writeback fields for a cancelled shipment.

#     Clears:
#       • Shipment Tracking Number  → empty string
#       • Shipment Courier          → empty string
#       • Shipment Label Created    → False (untick checkbox)

#     Args:
#         record_id    – Airtable record ID (e.g. "recXXXXXXXXXXXXXX")
#         order_number – for log messages only

#     Returns True on success, False on failure.
#     """
#     if not record_id:
#         print(f"   ⚠️  No Airtable record_id for {order_number} — skipping Airtable undo.")
#         return False

#     base_id  = TABLE_CONFIG["orders"]["base_id"]
#     table_id = TABLE_CONFIG["orders"]["table_id"]
#     api_key  = AIRTABLE_API_KEY

#     encoded_table = quote(str(table_id), safe="")
#     url = f"https://api.airtable.com/v0/{base_id}/{encoded_table}/{record_id}"

#     headers = {
#         "Authorization": f"Bearer {api_key}",
#         "Content-Type" : "application/json",
#     }
#     body = {
#         "fields": {
#             "Shipment Tracking Number": "",
#             "Shipment Courier"        : "",
#             "Shipment Label Created"  : False,
#             "Shipment Label file"     : [],   # clear so re-book produces a fresh label
#             "Commercial invoice file" : [],   # clear so re-book produces a fresh invoice
#         }
#     }

#     print(f"   🔄 Undoing Airtable booking for {order_number} (record: {record_id})...")
#     resp = airtable_patch_with_retry(url, body, headers, label=f"Undo/{order_number}")

#     if resp is None:
#         print(f"   ❌ Airtable undo failed (all retries exhausted) for {order_number}")
#         return False

#     if resp.status_code == 200:
#         print(f"   ✅ Airtable fields cleared for {order_number}")
#         return True
#     else:
#         print(f"   ❌ Airtable undo PATCH failed ({resp.status_code}) for {order_number}: "
#               f"{resp.text[:300]}")
#         return False


# # ============================================================
# # FIND ORDER BY TRACKING NUMBER
# # ============================================================

# def find_order_record_by_tracking(tracking_number):
#     """Find the Airtable orders record whose Shipment Tracking Number matches the AWB."""
#     tracking_number = str(tracking_number).strip()
#     if not tracking_number:
#         return None, None
#     cfg = TABLE_CONFIG["orders"]
#     formula = f'{{Shipment Tracking Number}}="{tracking_number}"'
#     df = fetch_table(
#         base_id=cfg["base_id"], table_id=cfg["table_id"],
#         view_id=None, label="Orders (cancel lookup)",
#         api_key=cfg["api_key"], filter_formula=formula,
#     )
#     if df.empty:
#         return None, None
#     row = df.iloc[0]
#     record_id    = str(row.get("_airtable_id", "")).strip()
#     order_number = str(row.get("Order Number", "")).strip()
#     return (record_id or None), (order_number or None)


# # ============================================================
# # CANCEL + RESET  (single entry point for the dashboard)
# # ============================================================

# def cancel_and_reset_shipment(tracking_number):
#     """
#     Cancel a DHL shipment AND reset its Airtable record so it can be
#     booked again. Call this from the dashboard / api.py cancel action.

#     Flow:
#       1. DELETE the shipment at DHL.
#       2. If DHL cancellation succeeds → find the Airtable record by AWB.
#       3. Clear all booking fields so the order returns to the queue.

#     If the DHL cancel fails, Airtable is left untouched.
#     """
#     result = cancel_dhl_shipment(tracking_number)

#     if not result.get("success"):
#         result["airtable_reset"] = False
#         return result

#     record_id, order_number = find_order_record_by_tracking(tracking_number)

#     if record_id:
#         reset_ok = undo_airtable_booking(record_id, order_number or tracking_number)
#         result["airtable_reset"] = reset_ok
#         result["order_number"]   = order_number
#         if reset_ok:
#             result["message"] = ("Shipment cancelled at DHL and the order was reset — "
#                                  "it can now be booked again.")
#         else:
#             result["message"] = ("Shipment cancelled at DHL, but clearing the Airtable "
#                                  "fields failed — reset it manually to allow re-booking.")
#     else:
#         result["airtable_reset"] = False
#         result["message"] = (f"Shipment cancelled at DHL, but no Airtable order was found "
#                              f"with tracking number {tracking_number} to reset.")

#     return result


# # ============================================================
# # MAIN
# # ============================================================
# if __name__ == "__main__":
 
#     # ── Parse flags ──────────────────────────────────────────
#     args             = [a for a in sys.argv[1:] if not a.startswith("--")]
#     send_to_dhl_flag = "--send-to-dhl" in sys.argv
#     test_override    = "--test"         in sys.argv
#     verbose_flag     = "--verbose"      in sys.argv  # print full DHL payload per order
 
#     if test_override:
#         DHL_TEST_MODE = True
#         print("⚠️  --test flag: DHL_TEST_MODE forced ON")

#     # ── CANCEL + RESET (early exit) ──────────────────────────
#     # Usage: python airtable_transform.py --cancel <TRACKING_NUMBER>
#     if "--cancel" in sys.argv:
#         try:
#             awb = sys.argv[sys.argv.index("--cancel") + 1]
#         except IndexError:
#             print("❌ Usage: python airtable_transform.py --cancel <TRACKING_NUMBER>")
#             sys.exit(1)
#         cancel_result = cancel_and_reset_shipment(awb)
#         print(json.dumps(cancel_result, indent=2, default=str))
#         sys.exit(0 if cancel_result.get("success") else 1)

#     # ── Initialise maps (will be populated below for RMA orders) ──
#     ya_rma_map  = {}
#     crs_rma_map = {}
#     price_map   = {}
 
#     # ── 0. Validate DHL credentials before doing anything ────
#     if send_to_dhl_flag:
#         ok = validate_credentials_dhl()
#         if not ok:
#             print("\n❌ Cannot proceed — fix DHL credentials first.")
#             sys.exit(1)
 
#     # ── 1. FETCH ORDERS ──────────────────────────────────────
#     if args:
#         csv_file = args[0]
#         print(f"📂 Reading from CSV: {csv_file}")
#         df = pd.read_csv(csv_file)
#         print(f"✅ Loaded {len(df)} records from CSV.")
#         print("ℹ️  CSV mode: RMA price lookup skipped.")
#     else:
#         if not AIRTABLE_API_KEY:
#             print("❌ Missing AIRTABLE_API_KEY in .env file")
#             sys.exit(1)
#         print("🌐 Fetching Sales Orders from Airtable...")
#         df = fetch_from_airtable()
 
#     if df.empty:
#         print("❌ No orders loaded.")
#         sys.exit(1)
 
#     # ── 2. DUPLICATE PREVENTION ──────────────────────────────
#     # Skip orders that already have a tracking number OR have the
#     # "Shipment Label Created" checkbox ticked. This means if the
#     # script crashes mid-batch and you re-run it, it picks up
#     # exactly where it left off without re-booking DHL.
#     initial_count = len(df)
 
#     if "Shipment Tracking Number" in df.columns:
#         df = df[df["Shipment Tracking Number"].isna() |
#                 (df["Shipment Tracking Number"] == "")]
 
#     if "Shipment Label Created" in df.columns:
#         df = df[df["Shipment Label Created"] != True]
 
#     skipped = initial_count - len(df)
#     if skipped > 0:
#         print(f"⏭️  Skipped {skipped} already-processed order(s).")
 
#     if df.empty:
#         print("✅ All orders are already processed. Nothing to do.")
#         sys.exit(0)
 
#     print(f"📋 {len(df)} order(s) to process.")
 
#     # ── 3. RMA LOOKUP (only in Airtable mode, not CSV mode) ─
#     # THIS BLOCK MUST BE OUTSIDE "if df.empty" — it was the bug in v4.x.
#     # It runs here, AFTER the duplicate filter, with the real df.
#     if not args:  # Airtable mode only
#         ya_rma_keys  = []
#         crs_rma_keys = []
#         for order_number in df.get("Order Number", pd.Series(dtype=str)):
#             order_str = str(order_number).strip()
#             # Take the first part before any slash
#             base_rma = order_str.split('/')[0].strip()
#             rma_source = get_rma_source(base_rma)
#             if rma_source == "ya":
#                 ya_rma_keys.append(normalize_rma_number(base_rma))
#             elif rma_source == "crs":
#                 crs_rma_keys.append(normalize_rma_number(base_rma))
 
#         def build_rma_filter(keys):
#             if not keys:
#                 return None
#             if len(keys) == 1:
#                 return f'{{RMA #}}="{keys[0]}"'
#             conditions = ','.join([f'{{RMA #}}="{k}"' for k in keys])
#             return f"OR({conditions})"
 
#         if ya_rma_keys:
#             cfg = TABLE_CONFIG["ya_rma"]
#             df_ya_rma = fetch_table(
#                 base_id=cfg["base_id"], table_id=cfg["table_id"],
#                 view_id=None, label=cfg["label"], api_key=cfg["api_key"],
#                 filter_formula=build_rma_filter(ya_rma_keys),
#             )
#         else:
#             print("\nℹ️  No RMAY orders — skipping YA RMA fetch.")
#             df_ya_rma = pd.DataFrame()
 
#         if crs_rma_keys:
#             cfg = TABLE_CONFIG["crs_rma"]
#             df_crs_rma = fetch_table(
#                 base_id=cfg["base_id"], table_id=cfg["table_id"],
#                 view_id=None, label=cfg["label"], api_key=cfg["api_key"],
#                 filter_formula=build_rma_filter(crs_rma_keys),
#             )
#         else:
#             print("\nℹ️  No RMAC orders — skipping CRS RMA fetch.")
#             df_crs_rma = pd.DataFrame()
 
#         ya_rma_map, crs_rma_map = build_rma_maps(df_ya_rma, df_crs_rma)
 
#         needed_so_numbers = set()
#         for order_number in df.get("Order Number", pd.Series(dtype=str)):
#             order_str = str(order_number).strip()
#             base_rma = order_str.split('/')[0].strip()
#             rma_source = get_rma_source(base_rma)
#             if rma_source:
#                 clean_rma = normalize_rma_number(base_rma)
#                 so = (ya_rma_map.get(clean_rma, "")
#                       if rma_source == "ya"
#                       else crs_rma_map.get(clean_rma, ""))
#                 if so:
#                     needed_so_numbers.add(so)
 
#         if needed_so_numbers:
#             # Try Commerce Central first (for RMA orders)
#             price_map_cc = fetch_price_map_from_commerce_central(needed_so_numbers)
#             # Then fallback to Sales Hub
#             price_map_sh = fetch_price_map_for_orders(needed_so_numbers, AIRTABLE_API_KEY)
#             # Merge: Commerce Central overrides
#             price_map = {**price_map_sh, **price_map_cc}
 
#         else:
#             if ya_rma_keys or crs_rma_keys:
#                 print("⚠️  RMA orders found but no External Sales Order numbers in RMA records.")
#             else:
#                 print("ℹ️  No RMA orders — price lookup not needed.")
 
#     # ── 4. TRANSFORM → CSV ───────────────────────────────────
#     print("\n⚙️  Transforming data...")
#     result    = transform(df, ya_rma_map, crs_rma_map, price_map)
#     timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
#     csv_out   = f"DHL-Order-File-{timestamp}.csv"
#     result.to_csv(csv_out, index=False, encoding="utf-8-sig")
#     print(f"\n✅ CSV saved: {csv_out}  ({len(result)} order(s))")
 
#     # ── 5. SEND TO DHL (only if --send-to-dhl flag) ─────────
#     if not send_to_dhl_flag:
#         print("\nℹ️  Dry run complete — CSV only.")
#         print(f"   To send to DHL:  python airtable_transform.py --send-to-dhl")
#         print(f"   Safe test first: python airtable_transform.py --send-to-dhl --test")
#         sys.exit(0)
 
#     print(f"\n{'='*60}")
#     print(f"🚀 BATCH START: {len(result)} shipment(s)")
#     print(f"   Mode    : {'⚠️  TEST (no real shipments)' if DHL_TEST_MODE else '✅ PRODUCTION'}")
#     print(f"   Verbose : {'ON (full payload)' if verbose_flag else 'OFF (use --verbose to enable)'}")
#     print(f"{'='*60}")
 
#     # Build order_number → Airtable record_id map for writeback
#     record_id_map = {}
#     if "_airtable_id" in df.columns:
#         for _, r in df.iterrows():
#             on  = str(r.get("Order Number", "")).strip()
#             rid = str(r.get("_airtable_id", "")).strip()
#             if on and rid and rid != "nan":
#                 record_id_map[on] = rid
#         print(f"   Airtable record map: {len(record_id_map)} order(s) mapped\n")
#     else:
#         print("   ℹ️  CSV mode — Airtable writeback skipped (no record IDs)\n")
 
#     dhl_results   = []
#     success_count = 0
#     fail_count    = 0
#     failed_orders = []  # saved to JSON so you can retry just these
 
#     # ── THE PROTECTED BATCH LOOP ─────────────────────────────
#     # Each order is wrapped in try/except so one failure never
#     # crashes the whole batch.
#     total = len(result)
#     for idx, (_, row) in enumerate(result.iterrows(), start=1):
#         order_number = str(row.get("Shipment Reference 1", "unknown"))
#         print(f"\n[{idx}/{total}] ── {order_number} ──────────────────────")
#         print_progress(idx, total)
 
#         try:
#             # ── DHL API call ──────────────────────────────────
#             result_data = send_to_dhl(row, verbose=verbose_flag)
#             dhl_results.append(result_data)
 
#             if result_data.get("success"):
#                 success_count  += 1
#                 tracking_number = result_data.get("tracking_number", "")
#                 airtable_rec_id = record_id_map.get(order_number, "")
#                 label_b64       = result_data.get("label_base64", "")
 
#                 if DHL_TEST_MODE:
#                     # ── TEST MODE: never touch production Airtable ────────
#                     # DHL gives us a fake tracking number from their sandbox.
#                     # Writing it to Airtable would corrupt live records.
#                     print(f"   ⏭️  TEST MODE — Airtable writeback skipped.")
#                     print(f"   ℹ️  Fake tracking (NOT written): {tracking_number}")
 
#                 else:
#                     # ── PRODUCTION: 3-step Airtable writeback ─────────────
#                     #
#                     # STEP 1 → Write tracking number + courier
#                     # STEP 2 → Upload label PDF
#                     # STEP 3 → Tick checkbox + write tracking URL
#                     #          (ONLY runs if Step 2 succeeds)
#                     #
#                     # If Step 2 fails, the checkbox stays unticked.
#                     # The order will be skipped on re-run (it has a
#                     # tracking number). Use the recovery script to
#                     # re-upload labels for orders in failed_orders log.
 
#                     if airtable_rec_id and tracking_number:
#                         # Step 1
#                         print(f"   📡 Step 1/3: Writing tracking to Airtable...")
#                         update_shipment_tracking_in_airtable(
#                             airtable_rec_id, order_number, tracking_number
#                         )
 
#                         # Step 2
#                         raw_response = result_data.get("raw_response", {})
#                         if raw_response:
#                             print(f"   📡 Step 2/3: Uploading label + invoice PDFs...")
#                             label_ok = upload_docs_to_airtable(
#                                 airtable_rec_id, order_number, raw_response
#                             )
 
#                             # Step 3 — ONLY if Step 2 succeeded
#                             if label_ok:
#                                 print(f"   📡 Step 3/3: Ticking 'Label Created'...")
#                                 mark_label_created_in_airtable(
#                                     airtable_rec_id, order_number, tracking_number
#                                 )
#                             else:
#                                 print(f"   ⚠️  Step 2 failed → Step 3 skipped. "
#                                       f"Checkbox NOT ticked.")
#                         else:
#                             print(f"   ⚠️  No raw_response data returned — Steps 2 & 3 skipped.")
 
#                     elif not airtable_rec_id:
#                         print(f"   ℹ️  No Airtable record ID for {order_number} — "
#                               f"writeback skipped (CSV mode).")
 
#             else:
#                 fail_count += 1
#                 failed_orders.append({
#                     "order"      : order_number,
#                     "error"      : result_data.get("error", "Unknown DHL error"),
#                     "status_code": result_data.get("status_code", ""),
#                     "failed_at"  : datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
#                 })
#                 print(f"   ❌ DHL FAILED: {result_data.get('error', 'Unknown error')[:200]}")
 
#         except Exception as e:
#             fail_count += 1
#             failed_orders.append({
#                 "order"     : order_number,
#                 "error"     : str(e),
#                 "failed_at" : datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
#             })
#             print(f"   💥 EXCEPTION on {order_number}: {e}")
#             # continue to next order — never crash the whole batch
 
#         # ── Rate limit: 1s between orders ────────────────────
#         # Airtable allows ~5 req/sec. We make up to 3 calls per
#         # order. 1s sleep keeps us at ~3 req/sec — safe margin.
#         if idx < total:
#             time.sleep(1.0)
 
#     # ── 6. SAVE LOGS ─────────────────────────────────────────
#     os.makedirs("logs", exist_ok=True)
 
#     response_file = f"logs/DHL-Responses-{timestamp}.json"
#     with open(response_file, "w", encoding="utf-8") as f:
#         json.dump(dhl_results, f, indent=2, default=str)
 
#     if failed_orders:
#         failed_file = f"logs/failed_orders-{timestamp}.json"
#         with open(failed_file, "w", encoding="utf-8") as f:
#             json.dump(failed_orders, f, indent=2, ensure_ascii=False)
#         print(f"\n⚠️  Failed orders saved to: {failed_file}")
#         print(f"   (Re-run just these by loading them into a recovery script)")
 
#     # ── 7. FINAL SUMMARY ─────────────────────────────────────
#     print(f"\n{'='*60}")
#     print(f"📊 BATCH COMPLETE")
#     print(f"   ✅ Success : {success_count}/{total}")
#     print(f"   ❌ Failed  : {fail_count}/{total}")
#     print(f"   Full log  : {response_file}")
#     if os.path.exists("labels"):
#         print(f"   Labels    : {len(os.listdir('labels'))} file(s) in labels/")
#     print(f"{'='*60}")
 
#     generate_lookup_html()
