<div align="center">
  <img src="addin_icon.png" width="128" height="128" alt="Word AI OCR Plugin" />
  <h1>Word AI OCR Plugin</h1>
  
![AI Word OCR Plugin Interface](showcase/screenshot.jpg)
</div>

Automate document layout and text reconstruction from photos or scans directly into Microsoft Word **without any paid APIs or API keys**.

---

## 🛠 Installation

The plugin is distributed as a self-installing script that configures Microsoft Word automatically.

1. Download **`INSTALL_PLUGIN.bat`** from the [latest Release](https://github.com/mantukin/word-ai-ocr/releases).
2. **Run the file**.
   *(It will download the manifest and register it in the Windows Registry)*.
3. Open **Word**.
4. Go to **Home (Главная)** ribbon tab.
5. Click the **Add-ins (Надстройки)** button.
6. Click **More Add-ins (Другие надстройки)** at the bottom of the popup.
7. Switch to the **SHARED FOLDER (ОБЩАЯ ПАПКА)** tab.
8. Select **Word AI OCR** and click **Add**.

*Note: In older Word versions, go to **Insert (Вставка)** -> **My Add-ins (Мои надстройки)** instead.*

---

## 🚀 The "No-API" Philosophy
Unlike traditional OCR solutions that require expensive subscriptions or complex API integrations (like GPT-4 Vision API), this plugin is built for the **standard user interface**:
- **ZERO API Costs**: You don't need to pay for tokens or manage API keys.
- **Works Everywhere**: Use any AI you already have access to (ChatGPT Free/Plus, Claude, Gemini, local LLMs).
- **Privacy & Control**: You decide which AI to trust with your data.

### How it works:
1. **Open your favorite AI chat** (web or mobile app).
2. **Send your document image** with the instruction provided by the plugin.
3. **Copy the JSON response** and paste it into Word. The plugin handles the heavy lifting of drawing tables, lists, and styles.

## ✨ Key Features
- **Complex Tables**: Reconstructs exact table grids with content.
- **Lists & Hierarchy**: Handles numbered and bulleted lists, including nested levels.
- **True Styling**: Preserves font families (e.g., Times New Roman), sizes, bold/italic, and alignment.
- **Seamless Flow**: No background servers needed for the release version.

---

## 👨‍💻 Development
To run the project locally (for contributors):
1. `npm install`
2. `npm start` (starts dev server on localhost:3000)
