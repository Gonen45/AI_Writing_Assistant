# Local AI Writing Assistant

A system-wide, privacy-first AI writing tool powered by local Large Language Models (LLMs) via **Ollama**. Built with **AutoHotkey v2**, this application integrates seamlessly into your Windows environment, allowing you to rewrite, translate, summarize, and fix text in any application using a custom right-click context menu or keyboard shortcuts.

## ✨ Key Features

* **System-Wide Context Menu:** Intercepts the native right-click menu when text is selected, offering AI-powered text manipulation directly in your browser, IDE, or word processor.
* **Review & Edit UI:** Instead of blindly pasting AI-generated text, a custom GUI pops up allowing you to compare the original text with the AI's output, tweak it, and apply it.
* **Dynamic Model Management:** Automatically fetches installed models directly from the Ollama API (`/api/tags`), allowing you to switch models on the fly via the system tray.
* **Algorithmic Layout Fix (No-LLM):** Instantly fixes English/Hebrew keyboard gibberish (e.g., typing "שלום" when the keyboard is set to English and getting "akuo") using a pure character mapping algorithm, completely bypassing LLM latency.
* **Robust JSON & Unicode Parsing:** Safely escapes user text for the API and decodes tricky Unicode sequences (like `\u05e9` for Hebrew characters) returned by the model.
* **Interaction Logging:** Saves a local history of your prompts and the AI's responses to a JSON file for later review.

## 🚀 Capabilities

* **Rewrite:** Enhances grammar, clarity, and professionalism.
* **Translate:** Automatically detects the language (English ↔ Hebrew) and translates accurately.
* **Summarize:** Extracts key points into bulleted lists.
* **Tone Adjustments:** Quickly rewrite text to be Formal, Casual, or Assertive.
* **Explain:** Simplifies complex jargon into plain language.

## 🛠️ Prerequisites

1. **[AutoHotkey v2](https://www.autohotkey.com/):** Required to run the script.
2. **[Ollama](https://ollama.com/):** Must be installed and running in the background.
3. **An LLM Model:** Pull your preferred model via Ollama (e.g., `ollama run hf.co/dicta-il/DictaLM-3.0-1.7B-Thinking-GGUF:Q8_0`).

## 📥 Installation & Usage

1. Clone this repository.
2. Make sure Ollama is running.
3. Double-click the `.ahk` script file to launch the assistant.
4. An 'H' icon will appear in your system tray. You can right-click it to test the Ollama connection, view your history, or switch the active model.
5. Highlight text anywhere on your computer and press `Ctrl+Alt+R` to rewrite, or right-click to open the AI context menu.

## ⌨️ Shortcuts

| Shortcut | Action |
| :--- | :--- |
| `Ctrl + Alt + R` | Rewrite & Improve |
| `Ctrl + Alt + G` | Fix Grammar Only |
| `Ctrl + Alt + T` | Translate (EN ↔ HE) |
| `Ctrl + Alt + S` | Summarize |
| `Ctrl + Alt + E` | Explain |
| `Ctrl + Alt + F` | Fix Keyboard Gibberish |
| `Ctrl + Alt + W` | Cycle Tone (Formal -> Casual -> Assertive) |

## 🏗️ Technical Architecture

* **API Communication:** Utilizes `WinHttp.WinHttpRequest.5.1` for synchronous HTTP POST/GET requests to the local Ollama server.
* **Regex Extraction:** Implements regular expressions to parse model reasoning tags (e.g., `<result>...</result>`) ensuring clean output without the model's internal thought process.
* **Asynchronous UX:** Features a non-blocking loading overlay to provide visual feedback while waiting for the local API to respond.
