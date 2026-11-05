# OpenDocs — התקנה וניהול לפי מוצר Office

```text
deploy/
  Word/          התקנה, הסרה, זיהוי ואריזה של Word בלבד
  Excel/         התקנה, הסרה, זיהוי ואריזה של Excel בלבד
  PowerPoint/    התקנה, הסרה, זיהוי ואריזה של PowerPoint בלבד
  Outlook/       התקנה, הסרה, זיהוי ואריזה של Outlook בלבד
  common/        המקור המשותף למנגנון ההתקנה והבדיקה
  tools/         כלי תחזוקה ואריזה משותפים
  dist/          חבילות ההפצה שנוצרו
```

בכל מוצר יש `Install.ps1`, ‏`Uninstall.ps1`, ‏`Test-Installed.ps1`, ‏`Detect-SCCM.ps1`, ‏`Check-Prerequisites.ps1`, ‏`Install-Prerequisites.ps1`, ‏`Get-Prerequisites.ps1`, ‏`Package.ps1`, ‏`product.json`, ‏`README.md` ועותק מקומי של `common/`.

כל תיקיית מוצר כוללת גם עותק של מתקיני התלויות ל־Windows x64. תלויות x86 ניתן להכין דרך Get-Prerequisites.ps1. הקבצים הגדולים מוחרגים מ־Git. תעודות פרטיות לעולם אינן מועתקות לחבילות.

## יצירת חבילה מלאה

מתוך שורש המאגר, למשל ל־Word בלבד:

```powershell
.\deploy\Word\Package.ps1 -CertificateThumbprint YOUR_SIGNING_CERTIFICATE_THUMBPRINT
```

או לכל המוצרים, כל אחד בחבילה נפרדת:

```powershell
.\deploy\tools\Package-All.ps1 -App All -CertificateThumbprint YOUR_SIGNING_CERTIFICATE_THUMBPRINT
```

האריזה מאתרת את VS2022 במחשב, בונה Release ומכינה תיקיית הפצה חדשה עם הסקריפטים, תעודת המפרסם הציבורית, קובצי התוסף ותלויותיו. תיקיית הפצה כזו עצמאית ואפשר להעתיק אותה לתחנה או למקור התוכן של SCCM. נתיב הפלט מוצג בסיום. חבילות קיימות אינן נמחקות אוטומטית.

תיקיות המקור שהוכנו כעת מסומנות `PackageReady=false`: יש בהן סקריפטים ותלויות, אך קובצי התוספים והתעודות נוצרים בבנייה המלאה. אין בתיקיות מקור אלה חבילת Release חתומה מוכנה להפצה. `-ScriptsOnly` מיועד להכנת מבנה/תלויות בלבד, ולא לפריסה של תוסף.

## עבודה עם מוצר בודד

מתוך חבילת Word מוכנה:

```powershell
.\Install.ps1
.\Test-Installed.ps1
.\Uninstall.ps1
```

התקנה והסרה לכל המשתמשים דורשות מנהל/SYSTEM. מצב `-CurrentUser` חייב להיות עקבי בין ההתקנה, הבדיקה וההסרה. הסרה אינה מסירה תלויות, תעודה משותפת או שיחות/הגדרות משתמש, כדי לא לפגוע במוצרים האחרים. ההתקנה אינה מתקינה את Microsoft Office עצמו.

ב־SCCM צור Application נפרד לכל מוצר. פקודות ההתקנה וההסרה מופיעות ב־README של אותו מוצר. השתמש ב־Detect-SCCM.ps1 **מהחבילה המלאה** כ־PowerShell Detection Method: הסקריפט עצמאי וגרסת המינימום מוטבעת בו באריזה. לעומת זאת Test-Installed.ps1 מחזיר דוח JSON לאבחון ואינו מתאים להדבקה ככלל זיהוי.

`3010` מציין צורך באתחול והרצה חוזרת לפני התקנת התוסף. `1618` מציין שיישום Office פתוח ויש לתזמן ניסיון חוזר. `1603` מציין כשל התקנה/הסרה. הזיהוי מבצע בדיקות קבצים, גרסה, רישום ותלויות ואינו מחליף בדיקת טעינה אמיתית בתוך Office.

נקודות הכניסה הישנות הוסרו. מתקינים ומסירים רק דרך תיקיית המוצר. README.md הוא הקובץ היחיד בשורש deploy; יתר הקבצים מסודרים בתיקיות.

## תחזוקה ובדיקות

לאחר שינוי ב־deploy/common הרץ:

```powershell
.\deploy\tools\Sync-Products.ps1
.\tools\Test-ProductDeployment.ps1
.\tools\Test-Prerequisites.ps1
```

Sync מעדכן את עותקי התמיכה ואת סקריפטי SCCM העצמאיים. הוא אינו מתקין תוספים ואינו משנה אמון בתעודות. כל Package.ps1 במאגר מפעיל אותו אוטומטית לפני אריזה.


## סביבה מנותקת

ראו [מדריך AIR GAPPED](docs/AIR-GAPPED.md).

