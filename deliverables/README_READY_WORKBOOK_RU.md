# Готовый файл с макросом: автоматическая сборка

Я не могу физически открыть Excel в этой Linux-среде и сохранить бинарный `.xlsm` напрямую. Поэтому сделал **автосборщик**, который создаёт готовый файл на вашей Windows-машине в 1 шаг.

## Что запускать

```powershell
powershell -ExecutionPolicy Bypass -File .\deliverables\build_ready_workbook.ps1 -InputWorkbook .\macros_test.xls -PatchModule .\deliverables\Cisco_DeliveryMode_Patch.bas -OutputWorkbook .\macros_test_ready.xlsm
```

## Результат

После выполнения получите:
- `macros_test_ready.xlsm` — готовый файл с импортированным рабочим макросом `GO_CISCO` для режимов `CIP / DDP / CIP&DDP`.

## Что внутри уже настроено

- В `MACROS!B10` выставлен default `CIP`.
- Добавлены подсказки значений `CIP/DDP/CIP&DDP` в `MACROS!Z1:Z3`.

## Важно

Если кнопка `GO` в книге привязана к старому `GO_CISCO`, откройте `Assign Macro` и выберите `Cisco_DeliveryMode_Patch.GO_CISCO`.
