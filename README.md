# excel-filter
Slice excel table and save in new excel files by key


At first nedd install modules:
## openpyxl
## re

By the next you need change path to work file excel ("D:/excel/test/Врачи.xlsx"), and change folder where you want to save new files. ("D:/excel/test/1/" + cleanName + '.xlsx')

Этот скрипт поможет разбить на файлы excel таблицу по определенной ячейке(ключу).
Скрипт собирает коллекции с одинм ключем и пересохраняет их в отдельные файлы подписав соответсвующим ключом.
Для работы необходим интерпритор Python 3.7 и несколько модулей (openpyxl - для работы с excel и re - для того чтобы получить "чистые имена"(без запрещенных знаков Windows) у новыъ файлов).
