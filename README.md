# Word AI OCR Plugin

Automate document layout and text reconstruction from photos or scans directly into Microsoft Word using the power of AI.

[Russian version below / Русская версия ниже]

---

## 🚀 Concept: "Bring Your Own AI"
This plugin acts as a bridge between powerful LLMs (like ChatGPT, Claude, or Gemini) and Microsoft Word. Instead of paying for expensive OCR APIs, you use your existing AI chat:
1. **Send image** to your favorite AI with our specialized instruction.
2. **Get structured JSON** that describes the document's layout (tables, lists, fonts).
3. **Paste JSON** into the plugin to instantly reconstruct the document in Word.

## ✨ Key Features
- **Accurate Tables**: Reconstructs complex table grids with content.
- **Lists & Hierarchy**: Handles numbered and bulleted lists, including nested levels.
- **Styling**: Preserves font families (e.g., Times New Roman), sizes, bold/italic, and alignment.
- **Robust Layout**: Handles multi-line paragraphs and complex indentations.

## 🛠 Installation (Production)
1. Download the `manifest-release.xml` file from this repository.
2. Place it in a local folder (e.g., `C:\Addins`).
3. Share that folder (Right-click -> Properties -> Sharing -> Share).
4. In **Word**, go to: `File` -> `Options` -> `Trust Center` -> `Trust Center Settings` -> `Trusted Add-in Catalogs`.
5. Enter the network path of your shared folder (e.g., `\\Your-PC\Addins`) and click **Add catalog**.
6. Check **Show in menu**, click OK, and restart Word.
7. Go to `Insert` -> `My Add-ins` -> `Shared Folder` and add **AI Word OCR**.

---

# Плагин Word AI OCR

Автоматизируйте перенос верстки и текста с фотографий или сканов документов прямо в Microsoft Word с помощью ИИ.

## 🚀 Концепция: "Bring Your Own AI"
Этот плагин служит мостом между мощными языковыми моделями (ChatGPT, Claude, Gemini) и Microsoft Word.
1. **Отправьте изображение** вашему ИИ вместе со специальной инструкцией (доступна в плагине).
2. **Получите структурированный JSON**, описывающий верстку (таблицы, списки, шрифты).
3. **Вставьте JSON** в плагин, чтобы мгновенно отрисовать документ в Word.

## ✨ Основные возможности
- **Точные таблицы**: Воссоздает сложные табличные сетки вместе с содержимым.
- **Списки и иерархия**: Поддерживает нумерованные и маркированные списки, включая вложенные уровни.
- **Форматирование**: Сохраняет шрифты (например, Times New Roman), размеры, жирность/курсив и выравнивание.
- **Сложная верстка**: Корректно обрабатывает переносы строк и отступы.

## 🛠 Установка
1. Скачайте файл `manifest-release.xml`.
2. Добавьте его в Word через "Каталоги доверенных надстроек" (подробная инструкция выше).
3. Плагин будет подгружаться напрямую с GitHub Pages, сервер на вашем ПК не требуется.

---

## 👨‍💻 Development
If you want to run the project locally:
1. `npm install`
2. `npm start` (for dev mode) or `START_AI_PLUGIN.bat` (for local production mode).
