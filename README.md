# Загрузчик SQL

Консольная программа на C++ для загрузки CSV или XLSX в SQL-базу через ODBC.

## Что умеет

- читает `.csv`;
- читает все листы из `.xlsx`;
- поддерживает пути с русскими буквами;
- определяет количество строк и столбцов;
- нормализует названия колонок для SQL;
- автоматически определяет типы колонок: `BIGINT`, `DOUBLE PRECISION`, `BOOLEAN`, `DATE`, `TIMESTAMP`, `TEXT`;
- создает SQL-таблицу;
- загружает строки;
- пишет лог в консоль и в `sql_loader.log`;
- работает с PostgreSQL, MySQL, SQL Server и произвольной ODBC-строкой.
- в интерактивном режиме сначала спрашивает параметры SQL-сервера, затем открывает окно выбора файла Windows.

## Сборка

Если `g++` виден в терминале:

```powershell
g++ main.cpp -std=c++20 -static-libgcc -static-libstdc++ -lodbc32 -lcomdlg32 -o sql_loader.exe
```

Если `g++` не виден, можно собрать полным путем:

```powershell
& 'C:\Users\krepo\AppData\Local\Microsoft\WinGet\Packages\BrechtSanders.WinLibs.POSIX.UCRT_Microsoft.Winget.Source_8wekyb3d8bbwe\mingw64\bin\g++.exe' main.cpp -std=c++20 -static-libgcc -static-libstdc++ -lodbc32 -lcomdlg32 -o sql_loader.exe
```

Если рядом с `exe` нужна DLL:

```powershell
Copy-Item 'C:\Users\krepo\AppData\Local\Microsoft\WinGet\Packages\BrechtSanders.WinLibs.POSIX.UCRT_Microsoft.Winget.Source_8wekyb3d8bbwe\mingw64\bin\libwinpthread-1.dll' .
```

## Запуск с интерфейсом

```powershell
.\sql_loader.exe
```

Программа сама спросит:

- тип базы данных;
- параметры подключения;
- путь к `.csv` или `.xlsx` через окно выбора файла Windows;
- имя таблицы для CSV;
- выполнить только проверку или сразу загрузить данные.

Для XLSX отдельное имя таблицы не спрашивается: каждый лист Excel загружается в отдельную SQL-таблицу, а имя таблицы берется из названия листа.

Путь можно вставлять с русскими буквами:

```powershell
.\sql_loader.exe --input "C:\Users\krepo\OneDrive\Рабочий стол\Log test\1.csv" --dry-run
```

Если имя файла написано кириллицей, имя таблицы по умолчанию может стать `imported_data`. В таком случае лучше указать таблицу вручную латиницей:

```powershell
.\sql_loader.exe --input "C:\данные\опрос.csv" --table survey_2026 --dry-run
```

## PostgreSQL

Для подключения к PostgreSQL нужен установленный ODBC-драйвер PostgreSQL.

Проверка без изменения базы:

```powershell
.\sql_loader.exe --input data.csv --db postgres --host localhost --port 5432 --dbname sociology_survey --user postgres --password YOUR_PASSWORD --dry-run
```

Реальная загрузка:

```powershell
.\sql_loader.exe --input data.csv --db postgres --host localhost --port 5432 --dbname sociology_survey --user postgres --password YOUR_PASSWORD
```

Пересоздать таблицу перед загрузкой:

```powershell
.\sql_loader.exe --input data.csv --db postgres --host localhost --port 5432 --dbname sociology_survey --user postgres --password YOUR_PASSWORD --drop-existing
```

## Другие СУБД

MySQL:

```powershell
.\sql_loader.exe --input data.csv --db mysql --host localhost --port 3306 --dbname mydb --user root --password YOUR_PASSWORD
```

SQL Server:

```powershell
.\sql_loader.exe --input data.csv --db sqlserver --host localhost --port 1433 --dbname mydb --user sa --password YOUR_PASSWORD
```

Произвольная ODBC-строка:

```powershell
.\sql_loader.exe --input data.csv --connstr "DRIVER={PostgreSQL ODBC Driver(UNICODE)};SERVER=localhost;PORT=5432;DATABASE=sociology_survey;UID=postgres;PWD=YOUR_PASSWORD;"
```

## Важные замечания

- Первая строка CSV/XLSX считается строкой заголовков.
- CSV с многострочными полями и лишними `;` в конце строк обрабатывается автоматически.
- Для `.xlsx` читаются все листы. Пустые листы пропускаются.
- Для `.xlsx` работает правило: один лист Excel = одна таблица SQL.
- Название SQL-таблицы для XLSX берется из названия листа Excel и нормализуется: пробелы заменяются на `_`, небезопасные символы убираются.
- Старый формат `.xls` не поддерживается. Его нужно сохранить как `.xlsx` или `.csv`.
- Если таблица уже существует, используйте `--drop-existing`.
- Для реальной загрузки нужен ODBC-драйвер выбранной СУБД.
