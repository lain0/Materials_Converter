# Materials_Converter is a script that converts file 'materials.xml' to  *.xlsx Excel Table.
File 'materials.xml' is a Unicut software file that includes information about working parameters
of a Unimach laser cutting cnc machines. From a version 1.0.7 Unicut have a built-in instrumet for exporting parameters,
but not for all of them. That is why this script was born.

Needed openpyxl library
and lxml library


Программа разбирает xml файл "materials.xml", в котором
содержится информация о режимах обработки станка лазерной резки
семейства Unimach и выводит параметры в Excel таблицу "Materials(XX)кВт",
где XX-выбранная мощность излучателя, которую надо предварительно ввести или задать в переменной Power.
Каждому режиму обработки соответствует отдельный лист таблицы. Имя листа
таблицы это ID режима в xml файле. Для навигации на первом листе таблицы
помещено оглавление с гиперссылками на имена режимов, как они указаны
в станке. Листы таблицы сразу могут быть распечатаны, т.к. все настройки печати уже заданы.

## Installation
```python -m pip install -r requirements.txt```

## RUN
```python3 ./Materials_Converter.py```
