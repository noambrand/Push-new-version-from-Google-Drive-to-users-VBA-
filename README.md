# Push-new-version-from-Google-Drive-to-users-VBA-
Push new version from Google Drive to users (3 methods in VBA)
Method1: Push new file version from Google Drive to users (VBA)
Method2: Push new code from Google Drive to original users file (VBA)
Method3: Push new code from a shared path on local network to original users file (VBA)
See project on stackoverflow: https://stackoverflow.com/questions/71829652/push-new-version-from-google-drive-to-users-vba/71829732#71829732

Existing solutions disadvantages that this solution solves:

● Some solutions require saving user's emails and mailing multiple users. If someone shares the file whoever receives the file will not receive version updates.
● Some solutions require the developer to register to a Zapeir or Integrate account in order to configure webhooks.
● Some solutions require a fixed filename (the new file name cannot be taken from Google Drive).
● Some solutions require the use of Google API which includes a complicated set of permissions that have to be configured(authentication with token issuance and secret code). Since in our case the file is shared publicly, the need for such permissions can be avoided, thus a simpler solution can be implemented.

How does it work?
The original file downloads a TXT file (It didn't work for me without downloading the google doc as TXT) from Google docs by a permanent link which contains the following data: 
Newest version number; New link to the new file version; The updates in the new version.
If there is a newer version upon opening the file the user will be notified about its existence, and the updates it contains, and ask permission to download the new version from Google Drive to the same file path as the original file.
Credit to part of method1 to Florian Lindstaedt posted:
https://www.linkedin.com/pulse/20140608044541-54506939-how-to-recall-an-old-excel-spreadsheet-version-control-with-vba

You have built a great Excel sheet. You share it and whoever gets it loves it and it gets handed around even more - you don't even know to whom.
Then it happens - something needs to be changed in the file: Some value changes in the worksheet, Some value was hardcoded and can't be changed by the user,
you think of another helpful feature, the database it connects to moves to a new server, you find a mistake, How do you let everyone know? How do you tell the users of your file that there is a newer version available if you don't even know who those users are? 
Maybe you are too lazy to collect and manage a user's mailing list.
Noam Brand noambbb@gmail.com



עדכון גרסאות תוכנה מקומיות מגוגל דרייב באמצעות VBA

בנית גיליון אקסל מעולה. אתה משתף אותו ומי שמקבל אותו אוהב אותו ומשתף אותו הלאה,  אתה אפילו לא יודע למי.
ואז זה קורה - צריך לשנות משהו בקובץ: שינוי כלשהו בגליון, קוד שצריך לשנות ,אתה חושב על עוד פיצ'ר מועיל, אתה מוצא באג, בסיס הנתונים עובר לשרת חדש והקובץ יפסיק לעבוד אם לא יעודכן, איך אתה מודיע לכולםשיש גרסה חדשה יותר זמינה אם אתה אפילו לא יודע מיהם אותם משתמשים?
או אולי אתה עצלן מכדי לנהל רשימת מיילים של משתמשים ורוצה דרך קלה יותר.
 
פתרונות קיימים לעדכון גרסאות וחסרונותיהם אשר קובץ זה פותר:

● חלק מהפתרונות מצריכים לאסוף מיילים של המשתמשים (אנשים חוששים לפרטיותם וקבלת ספאם), לנהל רשימת מיילים עם דיוור משתמשים מרובים (דורש השקעת זמן ולעיתים המיילים עלולים להגיע לתיבת הספאם בטעות) . אם מישהו ישתף את הקובץ ללא העברת המייל שלו הוא לא יקבל עדכוני גרסאות

● חלק מהפתרונות ב VBA דורשים יצירת חשבון בזאפייר או אינטגרומט עם webhook.

● חלק מהפתרונות ב VBA  דורשים ששם הקובץ יהיה קבוע (הפתרונות לא יודעים לשלוף את שם הקובץ החדש מגוגל דרייב).

● חלק מהפתרונות ב VBA  מחייבים שימוש ב  Google API הכולל מערך ההרשאות מסובך: דורש הקמת חשבון מיוחד עם מנגנון אימות, הנפקת token  וקוד סודי. כיוון שמדובר בקבצים שבכל אופן משותפים באופן פומבי ניתן להמנע מהצורך בהרשאות מעין אלו.

איך שיטה1 פועלת?
הקובץ המקורי מוריד מלינק קבוע לקובץ Google docs  קובץ בפורמט  txt למחשב המכיל בתוכו את הנתונים הבאים:
מספר הגרסה חדשה ביותר; לינק חדש לגרסה החדשה; החידושים שנוספו בגרסה החדשה.
במידה וישנה גרסה חדשה יותר תוצג הודעה למשתמש אודות קיום גרסה חדשה יותר, החידושים שנוספו בגרסה החדשה ותשאל אותו האם הוא מעוניין להוריד את הגרסה החדשה. אם המשתמש יאשר היא תוריד את הגרסה החדשה מגוגל דרייב.  
עד כאן מדובר בהחלפת הקובץ כולו בקובץ עדכני, בנוסף מוצעים עוד 2 דרכים נוספות לביצוע העדכון:
איך שיטה2 פועלת?
עדכון מגוגל דרייב רק לקוד ולא לקובץ כולו (קובץ bas שנשמר בגוגל דרייב ומורד ומחליף את הקיים בקובץ המקורי של המשתמש).
איך שיטה3 פועלת?
עדכון מנתיב מקומי רק לקוד ולא לקובץ כולו (קובץ bas שנשמר בנתיב מקומי מחליף את הקיים בקובץ המקורי של המשתמש).

נ.ב הצלחתי לרשום הודעה למשתמש עם פרטי העדכון מקובץ TXT רק באנגלית, בעברית התוכן מוצג כג'יבריש, אם אתם יודעים איך לסדר את זה אשמח שתשלחו לי מייל noambbb@gmail.com  או שתעשו לפרויקט fork).

לעיון בפרויקט בסטאקאוברפלואו: https://stackoverflow.com/questions/71829652/push-new-version-from-google-drive-to-users-vba/71829732#71829732
