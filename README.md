# המרת טופסי דיווח איכות מים — כלי קליטה

כלי להמרת טופסי דיווח איכות מים של אתרי דלק לפורמט קליטה למערכת המידע של רשות המים.

## התקנה

```bash
pip install streamlit openpyxl pandas
```

## מבנה הקבצים

```
project/
├── app.py                        # ממשק Streamlit
├── convert_report_to_intake.py   # מנוע ההמרה (backend)
├── טבלת_פרמטרים.xlsx              # טבלת מיפוי פרמטרים (חובה)
├── lab_codes.csv                  # קודי מעבדות
├── sampler_codes.csv              # קודי חברות דיגום
└── well_codes_memory.csv          # זיכרון קודי קידוחים (נוצר אוטומטית)
```

## הרצה

### ממשק גרפי (Streamlit)
```bash
streamlit run app.py
```

### שורת פקודה
```bash
# רגיל
python convert_report_to_intake.py file1.xlsx file2.xlsx

# עם השלמת קודי קידוח אינטראקטיבית
python convert_report_to_intake.py file.xlsx --interactive

# עם בדיקה היסטורית
python convert_report_to_intake.py file.xlsx --historical historical_data.xlsx

# כל האפשרויות
python convert_report_to_intake.py file.xlsx \
  --params טבלת_פרמטרים.xlsx \
  --labs lab_codes.csv \
  --samplers sampler_codes.csv \
  --historical historical_data.xlsx \
  --interactive \
  --output קליטה.xlsx \
  --error-report שגיאות.xlsx
```

## תהליך העבודה בממשק

1. **העלאת קבצי ייחוס** (סרגל צד) — טבלת פרמטרים, קודי מעבדות, קודי חברות דיגום. נטענים אוטומטית אם נמצאים בתיקיית הפרויקט.
2. **העלאת טופסי דיווח** — קובץ בודד או מרובים.
3. **השלמת קודי קידוח** — אם חסרים, הממשק מציג שדות להקלדה. קודים שהוקלדו נשמרים לזיכרון.
4. **תצוגה מקדימה** — טבלה עם השורות שנוצרו.
5. **הורדה** — קובץ קליטה + דוח שגיאות.

## ולידציות

- תאריך חסר/לא תקין
- קוד קידוח חסר (עם השלמה אינטראקטיבית + זיכרון)
- קוד מעבדה/חברת דיגום — התאמה מקורבת לרשימות הייחוס
- קוד פרמטר לא מוכר
- ערך ECFD >= 100 (חשד ליחידות שגויות)
- חריגה היסטורית — 2 סדרי גודל ביחס לנתון הקודם
