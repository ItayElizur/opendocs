# התקנה בסביבת AIR GAPPED

## הכנה במחשב בנייה מחובר

Microsoft Office עצמו חייב להיות מותקן בתחנות היעד. במחשב הבנייה דרושים VS2022 וכלי הבנייה של הפרויקט, Node/npm ותעודת חתימת קוד עם מפתח פרטי. אם גם מחשב הבנייה מנותק, יש להכין מראש בנפרד את כלי הבנייה וחבילות NuGet/npm; הסקריפט אינו מספק מטמון offline לבנייה.

מתוך שורש המאגר, לדוגמה Word:

```powershell
.\deploy\Word\Get-Prerequisites.ps1 -Architecture x64
.\deploy\Word\Package.ps1 -CertificateThumbprint YOUR_THUMBPRINT -OutDir C:\Staging\OpenDocs-Word-Release
```

בחר תיקיית פלט חדשה. האריזה בונה Release חתום ומצרפת את מתקיני .NET Framework 4.8, VSTO ו-WebView2, עם manifest וחתימות/hash לבדיקה. x64 הוא ארכיטקטורת Windows, לא בהכרח Office. ל-Windows x86 יש להכין x86 ולהעביר גם ‎-PrerequisiteArchitecture x86 לאריזה. ARM64 אינו נתמך בחבילה זו.

העבר את **כל תיקיית הפלט** לרשת המנותקת בתהליך ההעברה המקובל בארגון. אל תעביר רק Install.ps1 או את תיקיית המקור deploy/Word. בדוק שב-product.json מופיעים PackageReady=true וגרסה, ובחבילה נמצאים WordAiAddIn, common, prerequisites והתעודה הציבורית. פלט ScriptsOnly אינו ניתן לפריסת התוסף. המפתח הפרטי אינו מועבר לתחנות.

## תעודות

תאם עם IT את אמון התחנות בתעודת המפרסם ובשרשרת שלה, בהתאם למדיניות Office/GPO. דאג מראש גם לשורשי Microsoft ולשרשראות הנדרשות לאימות מתקיני התלויות ללא אינטרנט; אין לבטל בדיקת חתימות כדי לעקוף כשל. מדיניות בדיקת ביטול תעודות עשויה לדרוש שירותי CRL פנימיים או הכנת נתוני אמון מתאימים בארגון. תעודת HTTPS של שרת ה-AI היא נפרדת מתעודת חתימת הקוד.

## התקנה בתחנה

סגור את Word. פתח PowerShell כמנהל, מתוך תיקיית הפלט שהועברה:

```powershell
.\Check-Prerequisites.ps1
.\Install.ps1
.\Test-Installed.ps1
```

Check-Prerequisites הוא אבחון בלבד; תוצאה של תלות חסרה לפני ההתקנה צפויה. Install מתקין אוטומטית את התלויות החסרות מהקבצים המקומיים, ללא הורדה. כברירת מחדל ההתקנה לכל משתמשי המחשב. קוד 3010 דורש אתחול ואז הרצת Install.ps1 שוב: בשלב זה התוסף עצמו עדיין לא הותקן. אין אתחול אוטומטי. קוד 1618 מחייב סגירת Office וניסיון חוזר. לאחר Ready=true בדוח, פתח Word וודא שהחלונית נטענת בפועל.

להסרה: ‎.\Uninstall.ps1 כמנהל. התלויות המשותפות ונתוני המשתמש אינם מוסרים.

## SCCM ברשת המנותקת

צור Application לכל מוצר והגדר את **תיקיית הפלט המלאה** כמקור התוכן. הפץ ל-Distribution Points הפנימיים. פקודת התקנה והסרה, ב-PowerShell ‏64 ביט ותחת SYSTEM:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File Install.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File Uninstall.ps1
```

הדבק את מלוא Detect-SCCM.ps1 מהחבילה כ-PowerShell Detection Method. גרסת המינימום מוטבעת באריזה. הגדר 3010 כצורך באתחול, ותזמן התקנה חוזרת לאחריו; אל תסתמך על אתחול בלבד להשלמת ההתקנה. תזמן ניסיון חוזר ל-1618. במקרה של מדיניות שמחייבת סקריפטים חתומים, חתום גם על סקריפטי PowerShell לפי נוהל הארגון; ExecutionPolicy Bypass אינו עוקף GPO/WDAC/AppLocker.

המבנה הנוכחי הוא חבילת התקנה באמצעות PowerShell, ואינו מייצר MSI.

## שימוש ב-AI ללא אינטרנט

התקנה offline אינה מספקת מודל AI. נדרש שרת AI נגיש ברשת הפנימית ומודל מותקן בו. הגדר ב-OpenDocs את ה-provider, כתובת השרת, המודל והאימות המתאימים; ודא אמון בתעודת HTTPS שלו. שירותי ענן לא יהיו זמינים ברשת מנותקת.
