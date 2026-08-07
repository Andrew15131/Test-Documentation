# Placeholder для TestPlan_v1.0.xlsx

В этом репозитории ожидается файл TestPlan_v1.0.xlsx по TC-GH-003.

Обновлено: добавлен скрипт для генерации Excel-файла.

Как сгенерировать и загрузить реальный TestPlan_v1.0.xlsx:

1. Убедитесь, что у вас установлен Python 3 и пакет openpyxl:

   pip install openpyxl

2. Запустите скрипт из корня репозитория:

   python3 scripts/generate_testplan.py

   В результате будет создан файл `TestPlan_v1.0.xlsx` в текущей директории.

3. Загрузите сгенерированный файл в репозиторий через веб-интерфейс GitHub:
   Add file → Upload files → выберите `TestPlan_v1.0.xlsx` → Commit.

Альтернативно, вы можете выполнить коммит локально и создать PR:

   git checkout -b add-testplan
   git add TestPlan_v1.0.xlsx
   git commit -m "Add generated TestPlan_v1.0.xlsx"
   git push origin add-testplan

Если файл не должен быть публичным, вместо загрузки в репозиторий разместите его во внутреннем хранилище и обновите README/placeholder соответствующей ссылкой.
