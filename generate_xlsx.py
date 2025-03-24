import pandas as pd
import random
import json

# Number of rows to generate
num_rows = 2000

# Define column names directly in the code to ensure it works without dependencies
column_names = [
  "סטאטוס",
  "יחידה",
  "תת יחידה",
  "מיקום",
  "תיעדוף ",
  "שם הפרויקט",
  "מס' פרויקט מיגון",
  "מס' פרויקט מולטימדיה",
  "מס' פרויקט תקנ\"מ          ",
  "מוביל פרוייקט תקשוב אכ\"א",
  "חברת מולטימדיה",
  "סוג הפרויקט                                  (תקנ\"מ, מיגון, מולטימדיה)",
  "מהות המשימה",
  "  תוכנית עבודה                            (כן, לא)",
  "גוף מתקצב",
  "אישור מצו\"ב",
  "תאריך רשום בטבלה",
  "תאריך פתיחת פרויקט",
  "תאריך קבלת דמ\"צ",
  "קבלת דמ\"צ",
  "נפתח פרוייקט",
  "תוקצב",
  "שנת עבודה",
  "אומדן ציוד תקשוב דולרי",
  "אומדן ציוד תקשוב שיקלי",
  "אומדן שקלי ",
  "אומדן מולטימדיה שקלי",
  "אומדן דולרי חוש\"ן",
  "תוספת אומדן שיקלי",
  "תוספת אומדן דולרי",
  "סה\"כ אומדן שיקלי",
  "סה\"כ אומדן דולרי",
  "תאריך סיום בינוי",
  "סיום תקשוב",
  "צפי סיום",
  "סיום בפועל",
  "הערות"
]

# Try to load from JSON if it exists (optional)
try:
    with open("projects.json", "r", encoding="utf-8") as f:
        projects_template = json.load(f)
        column_names = projects_template["סטאטוס פרויקטים"]
        print("Loaded column names from projects.json")
except FileNotFoundError:
    print("Using default column names (projects.json not found)")

# Define massively expanded realistic data for different columns
unit_names = [f'יחידה {i+1}' for i in range(50)] + ['חיל התקשוב', 'אגף טכנולוגיה ולוגיסטיקה', 'פיקוד המרכז', 'אגף התכנון', 'חיל האוויר', 'חיל הים', 'זרוע היבשה', 'אגף המודיעין', 'מפקדת העומק', 'פיקוד הצפון', 'פיקוד הדרום', 'פיקוד העורף'] + [f'חטיבה {i+1}' for i in range(30)] + [f'אוגדה {i+1}' for i in range(20)] # 100+ options
sub_units = [f'מדור {i+1}' for i in range(100)] + ['מדור תכנון', 'מדור רכש', 'יחידת ביצוע', 'מדור בקרה', 'מחלקת פיתוח', 'מחלקת יישום', 'צוות תמיכה', 'גף טכנולוגיות', 'פלוגת קשר', 'סיירת מטכ"ל', 'שייטת 13', 'יחידה 8200', 'צוות ניהול פרויקטים', 'מחלקת תשתיות', 'גף תקציבים', 'צוות הדרכה'] + [f'פלוגה {i+1}' for i in range(50)] # 200+ options
locations = [f'בסיס {i+1}' for i in range(200)] + ['בסיס תל השומר', 'קריה - תל אביב', 'מחנה גלילות', 'בסיס נבטים', 'בסיס חצרים', 'בסיס רמון', 'בסיס פלמחים', 'בסיס עציון', 'בסיס צאלים', 'בסיס טכני חיפה', 'בסיס אשדוד', 'בסיס נחשונים', 'בסיס חיל האוויר', 'בסיס הצי', 'מחנה אימונים', 'שטח כינוס', 'עמדת פיקוד קדמית'] # 200+ options
priorities = ['גבוה', 'בינוני', 'נמוך', 'דחוף', 'קריטי', 'שגרתי', 'מיידי', 'חשוב', 'פחות דחוף', 'ללא דחיפות'] + [f'עדיפות {i+1}' for i in range(20)] # 30+ options
project_names_base = ['הקמת תשתיות תקשוב', 'שדרוג מערכות מידע', 'פרויקט אבטחת סייבר', 'התקנת מולטימדיה מתקדמת', 'הטמעת מערכת ERP', 'פיתוח אפליקציה ייעודית', 'הקמת חדר בקרה מבצעי', 'שיפור תשתית רשת', 'החלפת ציוד קצה', 'הדרכת משתמשים', 'פיתוח מערכת שליטה ובקרה', 'הקמת מערך תקשורת לוויינית', 'שדרוג מערכות הגנה', 'הטמעת טכנולוגיות ענן', 'פיתוח פלטפורמת ניהול ידע', 'הקמת מרכז נתונים', 'שיפור מערכות ניטור ובקרה', 'החלפת תשתיות חשמל', 'הקמת מערכת כיבוי אש', 'שיפוץ מבנים', 'רכש ציוד מחשוב', 'התקנת מערכות אזעקה', 'שדרוג מערכות מיזוג אוויר', 'הקמת מערכת גיבוי אנרגיה'] # 24 base names
# Expanded variations to ensure we can generate 2000 unique names
project_variations = [
    'מרכזי', 'חדש', 'מתקדם', 'מאובטח', 'ניסיוני', 'דחוף', 'מותאם', 'משולב', 'רחב היקף', 'מקומי', 
    'ארצי', 'יישובי', 'אסטרטגי', 'טקטי', 'מבצעי', 'לוגיסטי', 'פנים-ארגוני', 'חוץ-ארגוני', 'ראשוני', 'מתמשך', 
    'עתידי', 'דחוף ביותר', 'בסיסי', 'מורחב', 'נקודתי', 'מיוחד', 'גדול', 'קטן', 'מהיר', 'איכותי',
    'חסכוני', 'יעיל', 'חדשני', 'פורץ דרך', 'תקני', 'מבוקר', 'מאובטח היטב', 'בעל נראות גבוהה', 'משודרג', 'מתואם',
    'רב-שנתי', 'קצר טווח', 'ארוך טווח', 'בעל השפעה', 'מרובה משתתפים', 'מרובה שלבים', 'איטי', 'זריז', 'מדויק', 'גמיש'
] # 50 variations - combinations = 24 * 50 = 1200 unique names

# Add combinations with two variations to ensure we have enough unique names
project_combinations = []
for var1 in project_variations[:10]:  # Use first 10 variations
    for var2 in project_variations[10:20]:  # Use next 10 variations
        project_combinations.append(f"{var1} {var2}")

companies = [f'חברה {i+1}' for i in range(50)] + ['בזק', 'סלקום', 'פרטנר', 'HOT', 'סיסקו', 'מוטורולה', 'אלביט מערכות', 'רפאל', 'תעשייה אווירית', 'נס טכנולוגיות', 'אורקל', 'מיקרוסופט', 'IBM', 'HP', 'Dell', 'אמזון', 'גוגל', 'פייסבוק', 'אפל', 'אינטל', 'TSMC', 'סמסונג', 'LG', 'סוני', 'פיליפס', 'GE', 'סימנס', 'בואינג', 'לוקהיד מרטין', 'נרל דיינמיקס', 'נורת\'רופ גרומן', 'ריית\'און', 'ASML', 'Applied Materials', 'Lam Research', 'KLA', 'Teradyne']
project_types = ['תקנ"מ', 'מיגון', 'מולטימדיה', 'תשתית', 'סייבר', 'תוכנה', 'חומרה', 'הדרכה', 'לוגיסטיקה', 'בינוי', 'שיפוץ', 'רכש', 'אבטחה', 'אנרגיה', 'תקשורת', 'בקרה', 'ניטור', 'פיתוח', 'יישום', 'תחזוקה'] # 20+ types
task_descriptions = [f'משימה {i+1}' for i in range(200)] + [ # 200+ descriptions
    'הרחבת כיסוי רשת אלחוטית', 'שיפור אבטחת נתונים בבסיס', 'פריסת סיבים אופטיים', 'חיבור מערכות שליטה ובקרה',
    'הקמת מערך גיבוי נתונים', 'שדרוג מערכות הפעלה', 'הטמעת מערכת ניהול משתמשים', 'הגברת רוחב פס',
    'התקנת מערכות וידאו קונפרנס', 'שיפור מערכות תקשורת לוויינית', 'הקמת מוקד שירות ותמיכה טכנית',
    'החלפת שרתים מרכזיים', 'הקמת מערכת ניהול זהויות', 'שדרוג מערכות גיבוי חשמל', 'התקנת מערכות קירור חדשות',
    'שיפור תשתית אינטרנט', 'הקמת רשת תקשורת פנימית', 'שדרוג מערכות סינון תעבורה', 'התקנת מערכות ניטור רשת',
    'הטמעת מערכות הגנה מפני תקיפות סייבר', 'פיתוח מערכת לניהול פרויקטים', 'הקמת פורטל מידע פנים-ארגוני',
    'שדרוג מערכות CRM', 'הטמעת מערכת BI', 'פיתוח מערכת לניהול משאבי אנוש', 'הקמת מערכת ניהול מלאי',
    'שדרוג מערכות ERP קיימות', 'הטמעת מערכת לניהול מסמכים', 'פיתוח מערכת לניהול ידע', 'הקמת מערכת לניהול נכסים',
    'שדרוג מערכות לניהול לקוחות', 'הטמעת מערכת לניהול ספקים', 'פיתוח מערכת לניהול שרשרת אספקה', 'הקמת מערכת לניהול אירועים',
    'שדרוג מערכות לניהול תקלות', 'הטמעת מערכת לניהול שינויים', 'פיתוח מערכת לניהול סיכונים', 'הקמת מערכת לניהול איכות',
    'שדרוג מערכות לניהול תהליכים', 'הטמעת מערכת לניהול ביצועים', 'פיתוח מערכת לניהול תקציב', 'הקמת מערכת לניהול משאבים',
    'שדרוג מערכות לניהול ישיבות', 'הטמעת מערכת לניהול משימות', 'פיתוח מערכת לניהול לו"ז', 'הקמת מערכת לניהול תורים',
    'שדרוג מערכות לניהול פניות', 'הטמעת מערכת לניהול משוב', 'פיתוח מערכת לניהול סקרים', 'הקמת מערכת לניהול משאבים אנושיים'
]
funding_sources = [f'גוף מימון {i+1}' for i in range(50)] + ['משרד הביטחון', 'צה"ל', 'אגף התקשוב', 'משרד האוצר', 'תקציב פנימי', 'תרומות', 'שיתוף פעולה אזרחי-צבאי', 'קרן מחקר ופיתוח', 'גיוס המונים', 'השקעה פרטית', 'מימון ממשלתי', 'תקציב ייעודי', 'הלוואות בנקאיות', 'מענקים', 'חסויות מסחריות'] # 60+ sources
notes = [f'הערה {i+1}' for i in range(100)] + ['נדרש אישור נוסף', 'בהמתנה לתקצוב', 'מחכה לאישור ספק', 'בוצע בהצלחה', 'בשלב בדיקות', 'עיכוב בלוחות זמנים', 'חריגה מהתקציב', 'התקדמות טובה', 'דורש תיאום מול גורמים נוספים', 'אושר עקרונית', 'הושלם חלקית', 'ממתין להחלטה', 'בדיקה טכנית מתבצעת', 'התקבל אישור עקרוני', 'נדרש אישור סופי', 'בשלבי סיום', 'מוקפא זמנית', 'מבוטל', 'הועבר לטיפול גורם אחר', 'בבחינה מחדש', 'התקבל תקציב', 'התקבל אישור תקציבי', 'התקבל אישור סופי ממשרד הביטחון', 'התקבל אישור ממשרד האוצר', 'התקבל אישור מאגף התקשוב', 'התקבל אישור מפיקוד מרכז', 'התקבל אישור מפיקוד צפון', 'התקבל אישור מפיקוד דרום', 'התקבל אישור מפיקוד העורף', 'התקבל אישור מחיל האוויר', 'התקבל אישור מחיל הים', 'התקבל אישור מזרוע היבשה', 'התקבל אישור מאגף המודיעין'] # 130+ notes
yes_no_options = ['כן', 'לא', 'בטיפול', 'בהמתנה', 'מאושר', 'לא מאושר', 'תלוי', 'בבדיקה', 'בהכנה', 'סוכם', 'נדחה', 'אושר עקרונית'] # 12 options
project_statuses = ['בביצוע', 'הושלם', 'בתכנון', 'הוקפא', 'בהמתנה לאישור', 'מושק', 'סוכל', 'מושהה', 'ממתין לתקציב', 'אושר סופית', 'אושר עקרונית', 'נדחה', 'בבדיקה ראשונית', 'בשלבי הקמה', 'בשלבי פיתוח', 'בשלבי יישום', 'בשלבי סיום', 'בשלבי מסירה', 'בשלבי תחזוקה', 'בשלבי שיפוץ', 'בשלבי שדרוג', 'בשלבי פירוק', 'בשלבי סגירה'] # 24 statuses

def generate_date(start_year=2020, end_year=2026):
    month = random.randint(1, 12)
    day = random.randint(1, 28) # Keep it simple
    year = random.randint(start_year, end_year)
    return f"{day:02d}/{month:02d}/{year}"

# Function to generate unique project names
used_project_names = set()
def generate_unique_project_name():
    attempts = 0
    while attempts < 1000:  # Safety to prevent infinite loop
        base_name = random.choice(project_names_base)
        
        # Randomly decide to use a single variation or a combined variation
        if random.random() < 0.3 and project_combinations:  # 30% chance to use combination
            variation = random.choice(project_combinations)
            project_name = f"{base_name} {variation}"
        else:
            variation = random.choice(project_variations)
            project_name = f"{base_name} {variation}"
            
            # Sometimes add a location for more uniqueness
            if random.random() < 0.4:  # 40% chance to add location
                location = random.choice(locations)
                project_name = f"{project_name} - {location}"
        
        if project_name not in used_project_names:
            used_project_names.add(project_name)
            return project_name
        attempts += 1
    
    # If we failed to generate a unique name, create one with a unique identifier
    return f"פרויקט מיוחד #{len(used_project_names) + 1} - {base_name}"

# Ensure we can generate at least 2000 unique project names
potential_combinations = len(project_names_base) * len(project_variations) + len(project_combinations) * len(project_names_base)
print(f"Potential unique project name combinations: {potential_combinations}")
if potential_combinations < num_rows:
    print(f"Warning: May not generate {num_rows} unique names, only {potential_combinations} combinations available")
    print("Using additional strategies to ensure uniqueness")

# Generate fake project data
data = {}
data[column_names[0]] = [random.choice(project_statuses) for _ in range(num_rows)] # 'סטאטוס'
data[column_names[1]] = [random.choice(unit_names) for _ in range(num_rows)] # 'יחידה'
data[column_names[2]] = [random.choice(sub_units) for _ in range(num_rows)] # 'תת יחידה'
data[column_names[3]] = [random.choice(locations) for _ in range(num_rows)] # 'מיקום'
data[column_names[4]] = [random.choice(priorities) for _ in range(num_rows)] # 'תיעדוף '
data[column_names[5]] = [generate_unique_project_name() for _ in range(num_rows)] # 'שם הפרויקט' - UNIQUE
data[column_names[6]] = [random.randint(1000, 9999) for _ in range(num_rows)] # "מס' פרויקט מיגון"
data[column_names[7]] = [random.randint(1000, 9999) for _ in range(num_rows)] # "מס' פרויקט מולטימדיה"
data[column_names[8]] = [random.randint(1000, 9999) for _ in range(num_rows)] # "מס' פרויקט תקנ"מ"
data[column_names[9]] = [f'רס״ן {random.choice(["דוד כהן", "יעל לוי", "אורן ישראל", "נועה ברק", "רותם כהן", "אביב לוי", "גלעד מזרחי", "שירה אוחיון", "יוסי ביטון", "מיכל כהן", "איתי לוי", "רועי בר"]) }' for _ in range(num_rows)] # "מוביל פרוייקט תקשוב אכ\"א" - More names
data[column_names[10]] = [random.choice(companies) for _ in range(num_rows)] # "חברת תקשורת/מולטימדיה"
data[column_names[11]] = [random.choice(project_types) for _ in range(num_rows)] # "סוג הפרויקט"
data[column_names[12]] = [random.choice(task_descriptions) for _ in range(num_rows)] # "מהות המשימה"
data[column_names[13]] = [random.choice(yes_no_options) for _ in range(num_rows)] # "תוכנית עבודה" - More options
data[column_names[14]] = [random.choice(funding_sources) for _ in range(num_rows)] # "גוף מתקצב"
data[column_names[15]] = [random.choice(yes_no_options) for _ in range(num_rows)] # "אישור ב"מ /מצו"ב" - More options
data[column_names[16]] = [generate_date(2020, 2023) for _ in range(num_rows)] # "תאריך רשום בטבלה"
data[column_names[17]] = [generate_date(2020, 2024) for _ in range(num_rows)] # "תאריך פתיחת פרויקט"
data[column_names[18]] = [generate_date(2021, 2024) for _ in range(num_rows)] # "תאריך קבלת דמ"צ"
data[column_names[19]] = [random.choice(yes_no_options) for _ in range(num_rows)] # "קבלת דמ"צ" - More options
data[column_names[20]] = [random.choice(yes_no_options) for _ in range(num_rows)] # "נפתח פרוייקט" - More options
data[column_names[21]] = [random.choice(yes_no_options) for _ in range(num_rows)] # "תוקצב" - More options
data[column_names[22]] = [random.randint(2019, 2025) for _ in range(num_rows)] # "שנת עבודה"
data[column_names[23]] = [random.randint(10000, 500000) for _ in range(num_rows)] # "אומדן ציוד תקשוב דולרי"
data[column_names[24]] = [random.randint(50000, 2000000) for _ in range(num_rows)] # "אומדן ציוד תקשוב שיקלי"
data.setdefault(column_names[25], [random.randint(1000, 500000) for _ in range(num_rows)]) # "אומדן שקלי "
data[column_names[26]] = [random.randint(5000, 500000) for _ in range(num_rows)] # "אומדן מולטימדיה שקלי"
data[column_names[27]] = [random.randint(5000, 500000) for _ in range(num_rows)] # "אומדן דולרי חוש"ן"
data[column_names[28]] = [random.randint(1000, 100000) for _ in range(num_rows)] # "תוספת אומדן שיקלי"
data[column_names[29]] = [random.randint(1000, 50000) for _ in range(num_rows)] # "תוספת אומדן דולרי"
data[column_names[30]] = [random.randint(50000, 4000000) for _ in range(num_rows)] # "סה"כ אומדן שיקלי"
data[column_names[31]] = [random.randint(10000, 1000000) for _ in range(num_rows)] # "סה"כ אומדן דולרי"
data[column_names[32]] = [generate_date(2022, 2026) for _ in range(num_rows)] # "תאריך סיום בינוי"
data[column_names[33]] = [generate_date(2022, 2026) for _ in range(num_rows)] # "סיום תקשוב"
data[column_names[34]] = [generate_date(2023, 2026) for _ in range(num_rows)] # "צפי סיום"
data[column_names[35]] = [generate_date(2023, 2026) for _ in range(num_rows)] # "סיום בפועל"
data[column_names[36]] = [random.choice(notes) for _ in range(num_rows)] # "הערות"

# Create DataFrame
df = pd.DataFrame(data)

# Check if we have 2000 unique project names
unique_projects = df[column_names[5]].nunique()
print(f"Generated {unique_projects} unique project names out of {num_rows} rows")

# Save the data to an Excel file
output_file = "projects_data.xlsx"
try:
    df.to_excel(output_file, index=False, engine='openpyxl')
    print(f"✔ קובץ נוצר בהצלחה: {output_file}")
except Exception as e:
    print(f"Error creating Excel file: {e}")
    # Try with xlsxwriter as alternative
    try:
        df.to_excel(output_file, index=False, engine='xlsxwriter')
        print(f"✔ קובץ נוצר בהצלחה (with xlsxwriter): {output_file}")
    except Exception as e2:
        print(f"Error with xlsxwriter: {e2}")