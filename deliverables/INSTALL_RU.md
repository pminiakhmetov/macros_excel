# Как установить рабочий макрос (CIP / DDP / CIP&DDP)

1. Откройте `macros_test.xls` в Excel.
2. Нажмите `Alt + F11` (VBA Editor).
3. В меню: `File -> Import File...` и импортируйте `deliverables/Cisco_DeliveryMode_Patch.bas`.
4. В существующем модуле, где сейчас `GO_CISCO`, удалите старую процедуру `GO_CISCO` (или переименуйте её), чтобы не было дубля имени.
5. На листе `MACROS` добавьте dropdown (Form Control):
   - значения: `CIP,DDP,CIP&DDP`
   - LinkedCell: `B10`
6. Убедитесь, что кнопка `GO` назначена на макрос `GO_CISCO` (из модуля `Cisco_DeliveryMode_Patch`).
7. Сохраните файл как `xlsm`/`xls` с макросами.

## Что делает патч
- `CIP`: обычная генерация (как раньше).
- `DDP`: генерация по DDP-ветке, с маркером `_DDP` через комментарий в имени.
- `CIP&DDP`: за один запуск создаёт 2 файла из одного выбранного исходника:
  - CIP (`..._CIP&DDP`)
  - DDP (`..._CIP&DDP_DDP`)

Это обеспечивает наличие `CIP&DDP` в строке отчёта/имени и `_DDP` у DDP-файла.


## Автоматическая сборка готового файла
См. `deliverables/README_READY_WORKBOOK_RU.md` и скрипт `deliverables/build_ready_workbook.ps1`.
