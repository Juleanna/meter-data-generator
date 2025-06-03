# 🚀 Гід по деплою на GitHub

## Крок 1: Підготовка файлів

Переконайтеся, що у вас є всі необхідні файли:
```
meter-data-generator/
├── meter_generator.py      # Основна програма
├── requirements.txt        # Залежності Python
├── README.md              # Документація
├── .gitignore            # Файли для ігнорування Git
├── build.bat             # Скрипт збірки для Windows
├── LICENSE               # Ліцензія (створити окремо)
└── screenshots/          # Папка для скріншотів (опціонально)
```

## Крок 2: Створення репозиторію на GitHub

1. Перейдіть на [GitHub.com](https://github.com)
2. Натисніть зелену кнопку **"New"** або **"Create repository"**
3. Заповніть форму:
   - **Repository name**: `meter-data-generator`
   - **Description**: `Professional meter data generator with GUI and Excel export`
   - **Public** або **Private** (на ваш вибір)
   - ✅ **Add a README file** (не відмічайте, у нас вже є README)
   - **Add .gitignore**: None (у нас вже є)
   - **Choose a license**: MIT License (рекомендовано)

## Крок 3: Локальне налаштування Git

Відкрийте командний рядок (CMD) або PowerShell у папці з проектом:

```bash
# Ініціалізація Git репозиторію
git init

# Додавання файлів до відстеження
git add .

# Перший коміт
git commit -m "🎉 Початковий реліз - Генератор даних лічильників

✨ Функції:
- Графічний інтерфейс користувача
- Підтримка 1-фазних та 3-фазних лічильників  
- Генерація 1200 записів з реалістичними даними
- Експорт в Excel з діаграмами
- Автоматичний аналіз напруги по годинах"

# Налаштування основної гілки
git branch -M main

# Підключення до віддаленого репозиторію (замініть YOUR_USERNAME на ваш GitHub логін)
git remote add origin https://github.com/YOUR_USERNAME/meter-data-generator.git

# Завантаження на GitHub
git push -u origin main
```

## Крок 4: Створення релізу з EXE файлом

### Збірка EXE файлу:
```bash
# Встановлення PyInstaller
pip install pyinstaller

# Створення EXE файлу
pyinstaller --onefile --windowed --name "MeterGenerator" --add-data "README.md;." meter_generator.py
```

### Створення релізу на GitHub:
1. Перейдіть в ваш репозиторій на GitHub
2. Натисніть **"Releases"** → **"Create a new release"**
3. Заповніть форму:
   - **Tag version**: `v1.0.0`
   - **Release title**: `🎉 Meter Data Generator v1.0.0`
   - **Description**:
   ```markdown
   ## 🎯 Що нового в v1.0.0
   
   ### ✨ Нові функції:
   - Графічний інтерфейс користувача на Tkinter
   - Генерація реалістичних даних напруги
   - Підтримка 1-фазних та 3-фазних лічильників
   - Експорт в Excel з автоматичними діаграмами
   - Погодинний аналіз мін/макс/середніх значень
   
   ### 📦 Завантаження:
   - **MeterGenerator.exe** - готова програма для Windows (не потребує Python)
   - **Source code** - вихідний код для розробників
   
   ### 🔧 Системні вимоги:
   - Windows 7/10/11
   - 50 MB вільного місця
   
   ### 🚀 Швидкий старт:
   1. Завантажте MeterGenerator.exe
   2. Запустіть програму
   3. Налаштуйте параметри лічильника
   4. Згенеруйте дані та експортуйте в Excel
   ```

4. Завантажте EXE файл через **"Attach binaries"**
5. Натисніть **"Publish release"**

## Крок 5: Налаштування GitHub Pages (опціонально)

Для створення веб-сторінки проекту:

1. В налаштуваннях репозиторію → **Pages**
2. **Source**: Deploy from a branch
3. **Branch**: main / (root)
4. **Save**

Ваша сторінка буде доступна за адресою: `https://YOUR_USERNAME.github.io/meter-data-generator`

## Крок 6: Додавання тем та міток

### Теми (Topics):
В налаштуваннях репозиторію додайте теми:
```
python, gui, tkinter, excel, data-generator, meter, electricity, pandas, openpyxl, windows
```

### Мітки для Issues:
Створіть мітки для класифікації проблем:
- `bug` (червона) - Щось працює неправильно
- `enhancement` (синя) - Нова функція або покращення
- `documentation` (зелена) - Покращення документації
- `help wanted` (жовта) - Потрібна допомога спільноти

## Крок 7: Створення CI/CD (GitHub Actions)

Створіть файл `.github/workflows/build.yml`:

```yaml
name: Build and Release

on:
  push:
    tags:
      - 'v*'

jobs:
  build:
    runs-on: windows-latest
    
    steps:
    - uses: actions/checkout@v3
    
    - name: Set up Python
      uses: actions/setup-python@v3
      with:
        python-version: '3.8'
    
    - name: Install dependencies
      run: |
        python -m pip install --upgrade pip
        pip install -r requirements.txt
        pip install pyinstaller
    
    - name: Build executable
      run: |
        pyinstaller --onefile --windowed --name "MeterGenerator" meter_generator.py
    
    - name: Create Release
      uses: softprops/action-gh-release@v1
      with:
        files: dist/MeterGenerator.exe
      env:
        GITHUB_TOKEN: ${{ secrets.GITHUB_TOKEN }}
```

## 🎉 Готово!

Ваш проект тепер опублікований на GitHub з:
- ✅ Професійною документацією
- ✅ Готовим EXE файлом для користувачів
- ✅ Автоматичною збіркою релізів
- ✅ Красивою GitHub сторінкою

### Подальші кроки:
1. Додайте скріншоти інтерфейсу в папку `screenshots/`
2. Створіть CONTRIBUTING.md з правилами участі
3. Додайте unit тести
4. Налаштуйте автоматичні оновлення через GitHub Releases API