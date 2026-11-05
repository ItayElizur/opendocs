# OpenDocs for Word

תיקייה עצמאית לניהול תוסף Word. כל הפקודות כאן מטפלות רק במוצר הזה.

| קובץ | שימוש |
|---|---|
| Install.ps1 | התקנה / התקנה חוזרת לתיקון; תלויות חסרות מותקנות אוטומטית מהחבילה |
| Uninstall.ps1 | הסרת התוסף בלבד; שומר תעודות, תלויות ונתוני משתמש |
| Test-Installed.ps1 | דוח JSON: קבצים, גרסה, רישום, תלויות ובעיות |
| Detect-SCCM.ps1 | סקריפט זיהוי עצמאי ל־SCCM, ללא תלות בקבצים לידו |
| Check-Prerequisites.ps1 | בדיקת תלויות בלבד |
| Install-Prerequisites.ps1 | התקנת תלויות בלבד |
| Get-Prerequisites.ps1 | הורדת מתקיני Microsoft למחשב הבנייה, ללא הפעלתם |
| Package.ps1 | יצירת חבילת הפצה מלאה של Word |
| product.json | שם המוצר, גרסה והאם החבילה מכילה קבצים שנבנו |
| common/ | עותק מקומי של מנגנון ההתקנה והבדיקה |
| prerequisites/ | מתקיני Microsoft ו־manifest.json עם SHA256 |

## הכנת חבילת הפצה במחשב הבנייה

```powershell
.\Get-Prerequisites.ps1
.\Package.ps1 -CertificateThumbprint YOUR_SIGNING_CERTIFICATE_THUMBPRINT
```

האריזה מאתרת Visual Studio 2022, משחזרת חבילות npm/NuGet, בונה רק את Word ב־Release, ומעתיקה את התוסף, התעודה הציבורית והתלויות לתיקיית הפצה חדשה תחת deploy/dist/Word/. אין צורך לתקן נתיב MSBuild בקוד.

אפשר גם לארוז פלט שכבר נבנה ונחתם:

```powershell
.\Package.ps1 -SkipBuild -PayloadDirectory C:\Built\WordAiAddIn -PublicCertificatePath C:\Certificates\Publisher.cer
```

`-ScriptsOnly` מכין סקריפטים ותלויות בלבד. חבילה כזו מסומנת PackageReady=false ואינה ניתנת להתקנת התוסף. בתיקיית המקור לא קיימים עדיין קובצי תוסף או תעודת חתימה: יש לבנות חבילה מלאה לפני פריסה.

## על תחנת היעד — מתוך חבילת ההפצה המוכנה

הרץ PowerShell ‏64 ביט כמנהל להתקנה לכל משתמשי המחשב:

```powershell
.\Install.ps1
.\Test-Installed.ps1
.\Uninstall.ps1
```

`-CurrentUser` מתקין או מסיר למשתמש הנוכחי בלבד. יש להשתמש באותו היקף בהתקנה, בדיקה והסרה. `-InstallRoot` הוא תיקיית האב OpenDocs, לא תיקיית WordAiAddIn עצמה. סקריפטי התקנה/הסרה/התקנת תלויות תומכים ב־`-WhatIf`.

ל־Test-Installed: קוד 0 = Ready, קוד 1 = חסר רכיב או נדרשת בדיקה, קוד 2 = שגיאת בדיקה. Installed ו־Ready הם שדות נפרדים. הבדיקה אינה מוכיחה טעינה אמיתית בתוך Office ואינה מאמתת אמון בתעודה.

ההסרה אינה סוגרת מסמכים בכוח. אם Word פתוח, התקנה/הסרה מחזירות 1618. תלויות שמבקשות אתחול מחזירות 3010; יש לאתחל ולהריץ שוב את ההתקנה. לא מבוצע אתחול אוטומטי.

## SCCM

- Application / Script Installer, ‏Install for system.
- התקנה: `powershell.exe -NoProfile -ExecutionPolicy Bypass -File Install.ps1`
- הסרה: `powershell.exe -NoProfile -ExecutionPolicy Bypass -File Uninstall.ps1`
- זיהוי: הדבק את **כל תוכן Detect-SCCM.ps1 מהחבילה שנבנתה** לתוך PowerShell Detection Method. אין בו תלות במיקום תיקיית החבילה.
- הזיהוי מחזיר קוד 0 עם פלט רק כשהקבצים, הרישום, הגרסה והתלויות תקינים; כשהמוצר חסר הוא מחזיר 0 ללא פלט. שגיאת בדיקה מחזירה קוד 1.
- באריזה מלאה גרסת המינימום מוטבעת אוטומטית בסקריפט הזיהוי.
- קודים: 0 הצלחה, 3010 נדרש אתחול והרצה חוזרת, 1618 ניסיון חוזר לאחר סגירת Office, 1603 כשל.
- הגדר 1618 כ־Fast Retry. התקנת התלויות בלבד אינה מצב Installed של התוסף.

הרכיבים המשותפים נוצרים מ־deploy/common באמצעות deploy/tools/Sync-Products.ps1. שינוי במנגנון עושים במקור המשותף ואז מריצים tools/Sync-Products.ps1. חבילת ההפצה עצמאית: אין צורך בתיקיות המוצרים האחרים או במחשב הפיתוח.


