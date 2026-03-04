#Requires AutoHotkey v2.0

; ==========================================
; CONFIGURATION
; ==========================================
OLLAMA_MODEL   := "hf.co/dicta-il/DictaLM-3.0-1.7B-Thinking-GGUF:Q8_0"  ; changed at runtime by model switcher
OLLAMA_URL     := "http://localhost:11434/api/generate"
OLLAMA_TIMEOUT := 60000
LOG_ENABLED    := true
LOG_FILE       := A_ScriptDir "\ollama_history.json"

; ==========================================
; TRAY MENU
; ==========================================
A_TrayMenu.Delete()
A_TrayMenu.Add("Test Ollama Connection", TrayTestConnection)
A_TrayMenu.Add("View History",           TrayViewHistory)
A_TrayMenu.Add("Switch Model",           TraySwitchModel)
A_TrayMenu.Add()
A_TrayMenu.Add("Reload Script", (*) => Reload())
A_TrayMenu.Add("Exit",          (*) => ExitApp())
A_TrayMenu.Default := "Test Ollama Connection"
TraySetIcon("imageres.dll", 96)

TrayTestConnection(*) {
    ok := PingOllama()
    MsgBox(ok ? "OK - Ollama is running and reachable."
               : "FAIL - Could not reach Ollama.`nMake sure it is running.",
           "Connection Test", ok ? "Iconi" : "Icon!")
}

TrayViewHistory(*) {
    if !FileExist(LOG_FILE) {
        MsgBox "No history file found yet.", "History"
        return
    }
    Run 'notepad.exe "' LOG_FILE '"'
}

; Fetch installed model names from Ollama /api/tags
GetInstalledModels() {
    http := ComObject("WinHttp.WinHttpRequest.5.1")
    try {
        http.Open("GET", "http://localhost:11434/api/tags", false)
        http.SetTimeouts(5000, 5000, 5000, 5000)
        http.Send()
    } catch {
        return []
    }
    raw := http.ResponseText
    models := []
    pos := 1
    while pos := RegExMatch(raw, '"name"\s*:\s*"([^"]+)"', &m, pos) {
        models.Push(m[1])
        pos += StrLen(m[0])
    }
    return models
}

TraySwitchModel(*) {
    global OLLAMA_MODEL
    models := GetInstalledModels()
    if (models.Length = 0) {
        MsgBox "Could not fetch models.`nIs Ollama running?", "Switch Model", "Icon!"
        return
    }

    ; Build a numbered list for the InputBox
    list := ""
    loop models.Length
        list .= A_Index . ". " . models[A_Index] . "`n"

    h := 60 + models.Length * 18
    result := InputBox("Installed models:`n`n" . list . "`nEnter number to select:", "Switch Model", "w420 h" . h)
    if (result.Result != "OK" || Trim(result.Value) = "")
        return

    choice := Integer(Trim(result.Value))
    if (choice < 1 || choice > models.Length) {
        MsgBox "Invalid selection.", "Switch Model", "Icon!"
        return
    }

    OLLAMA_MODEL := models[choice]
    MsgBox "Model switched to:`n" . OLLAMA_MODEL, "Switch Model", "Iconi"
}

; ==========================================
; UI: Loading overlay
; ==========================================
ShowLoading(action := "") {
    global loadingGui
    if IsSet(loadingGui)
        try loadingGui.Destroy()

    switch action {
        case "rewrite":
            icon := "✍️"
            msg  := "Rewriting..."
        case "grammar":
            icon := "🔍"
            msg  := "Fixing grammar..."
        case "translate":
            icon := "🌐"
            msg  := "Translating..."
        case "summarise":
            icon := "📋"
            msg  := "Summarising..."
        case "explain":
            icon := "💡"
            msg  := "Explaining..."
        case "formal":
            icon := "👔"
            msg  := "Rewriting (Formal)..."
        case "casual":
            icon := "😊"
            msg  := "Rewriting (Casual)..."
        case "assertive":
            icon := "💪"
            msg  := "Rewriting (Assertive)..."
        default:
            icon := "🤖"
            msg  := "Processing..."
    }

    loadingGui := Gui("-Caption +AlwaysOnTop +ToolWindow +Border")
    loadingGui.BackColor := "1A1A2E"
    loadingGui.MarginX := 12
    loadingGui.MarginY := 8
    loadingGui.SetFont("s16", "Segoe UI")
    loadingGui.Add("Text", "x10 y8 w30 h30 Background1A1A2E", icon)
    loadingGui.SetFont("s11 bold cE0E0FF", "Segoe UI")
    loadingGui.Add("Text", "x46 y12 w230 Background1A1A2E", msg)
    loadingGui.Show("NoActivate xCenter y20 w296 h48")

    ; Capsule shape
    hwnd := loadingGui.Hwnd
    hRgn := DllCall("CreateRoundRectRgn","int",0,"int",0,"int",297,"int",49,"int",48,"int",48,"ptr")
    DllCall("SetWindowRgn","ptr",hwnd,"ptr",hRgn,"int",1)
    WinSetTransparent(240, "ahk_id " hwnd)
}

HideLoading() {
    global loadingGui
    if IsSet(loadingGui)
        loadingGui.Destroy()
}

; ==========================================
; UI: Editable result window
; ==========================================
ShowResultEditor(original, aiResult, title := "Review AI Result") {
    editorGui := Gui("+AlwaysOnTop +Resize", "  " . title)
    editorGui.BackColor := "1E1E2E"
    editorGui.MarginX := 14
    editorGui.MarginY := 10

    ; Original label + box
    editorGui.SetFont("s8 bold c6E96F8", "Segoe UI")
    editorGui.Add("Text", "x14 y10 w570", "ORIGINAL")
    editorGui.SetFont("s9 cC8C8E0", "Consolas")
    origBox := editorGui.Add("Edit", "x14 y26 w572 h75 ReadOnly Wrap -VScroll Background2A2A3E cC8C8E0", original)

    ; Divider
    editorGui.Add("Text", "x14 y108 w572 h1 Background3A3A5E", "")

    ; Result label + box
    editorGui.SetFont("s8 bold c6E96F8", "Segoe UI")
    editorGui.Add("Text", "x14 y116 w400", "AI RESULT  —  edit before applying")
    editorGui.SetFont("s10 cE0E0FF", "Segoe UI")
    resultBox := editorGui.Add("Edit", "x14 y134 w572 h180 Wrap Background2A2A3E cE0E0FF", aiResult)

    ; Footer divider
    editorGui.Add("Text", "x0 y322 w620 h1 Background3A3A5E", "")

    ; Buttons
    editorGui.SetFont("s9 bold", "Segoe UI")
    btnApply   := editorGui.Add("Button", "x14 y330 w100 h28 Default", "✓  Apply")
    btnDiscard := editorGui.Add("Button", "x122 y330 w90  h28",        "✕  Discard")

    editorGui.Show("w600 h370")

    approved  := false
    finalText := ""

    ApplyFn(*) {
        finalText := resultBox.Value
        approved  := true
        editorGui.Destroy()
    }

    btnApply.OnEvent("Click",   ApplyFn)
    btnDiscard.OnEvent("Click", (*) => editorGui.Destroy())
    editorGui.OnEvent("Close",  (*) => editorGui.Destroy())

    WinWaitClose("ahk_id " editorGui.Hwnd)
    return approved ? finalText : ""
}
; ==========================================
; PING
; ==========================================
PingOllama() {
    http := ComObject("WinHttp.WinHttpRequest.5.1")
    try {
        http.Open("GET", "http://localhost:11434", false)
        http.SetTimeouts(3000, 3000, 3000, 3000)
        http.Send()
        return true
    } catch {
        return false
    }
}

; ==========================================
; JSON escape — user text only
; ==========================================
EscapeUserText(str) {
    str := StrReplace(str, "\",  "\\")
    str := StrReplace(str, '"',  '\"')
    str := StrReplace(str, "`r", "")
    str := StrReplace(str, "`n", "\n")
    str := StrReplace(str, "`t", "\t")
    return str
}

; ==========================================
; LOGGING
; ==========================================
LogInteraction(hotkey, original, result) {
    if !LOG_ENABLED
        return
    timestamp := FormatTime(, "yyyy-MM-dd HH:mm:ss")
    entry := '{"time":"' . timestamp . '","hotkey":"' . hotkey
           . '","input":"'  . EscapeUserText(original)
           . '","output":"' . EscapeUserText(result) . '"}' . "`n"
    FileAppend entry, LOG_FILE, "UTF-8"
}

; ==========================================
; ==========================================
; EXTRACT AI RESULT
; ==========================================
ExtractResult(rawResponse) {
    ; Step 1: decode \uXXXX unicode escapes (Hebrew chars etc.)
    raw := rawResponse
    while RegExMatch(raw, "\\\\u([0-9a-fA-F]{4})", &um)
        raw := StrReplace(raw, um[0], Chr(Integer("0x" . um[1])))

    ; Step 2: remove ALL newline variants — Ollama splits tags across lines
    raw := StrReplace(raw, "\n", "")
    raw := StrReplace(raw, "\r", "")
    raw := StrReplace(raw, "`n", "")
    raw := StrReplace(raw, "`r", "")

    ; Step 3: decode escaped slash
    raw := StrReplace(raw, "\/", "/")

    ; Step 4: try to find result tags — handles both <r> and <r>
    if RegExMatch(raw, 'i)<r(?:esult)?>([\s\S]*?)</r(?:esult)?>', &m)
        return Trim(m[1], " `t`n`r")

    ; Step 5: no tags found — extract the raw "response" field value as fallback
    ; The model wrote the answer without tags — just use it directly
    if RegExMatch(raw, '"response"\s*:\s*"((?:[^"\\]|\\.)*)"', &jm) {
        fallback := jm[1]
        fallback := StrReplace(fallback, "\n", "`n")
        fallback := StrReplace(fallback, "\/", "/")
        fallback := StrReplace(fallback, "\`"", "`"")
        fallback := StrReplace(fallback, "\\", "\")
        fallback := Trim(fallback, " `t`n`r")
        if (StrLen(fallback) > 0)
            return fallback
    }

    return ""
}










; CORE ENGINE
; ==========================================
CallOllama(instruction, textToProcess, action := "") {
    ShowLoading(action)

    prompt   := instruction . '\n\nText:\n' . EscapeUserText(textToProcess)
    jsonBody := '{"model":"' . OLLAMA_MODEL . '","prompt":"' . prompt . '","stream":false}'

    http := ComObject("WinHttp.WinHttpRequest.5.1")
    try {
        http.Open("POST", OLLAMA_URL, false)
        http.SetTimeouts(OLLAMA_TIMEOUT, OLLAMA_TIMEOUT, OLLAMA_TIMEOUT, OLLAMA_TIMEOUT)
        http.SetRequestHeader("Content-Type", "application/json")
        http.Send(jsonBody)
    } catch as e {
        HideLoading()
        MsgBox "Failed to connect to Ollama.`nIs it running?`n`nError: " e.Message,
               "Connection Error", "Icon!"
        return ""
    }

    response := http.ResponseText

    ; decode unicode escapes Ollama emits
    response := StrReplace(response, "\u003c", "<")
    response := StrReplace(response, "\u003e", ">")
    response := StrReplace(response, "\u0026", "&")

    HideLoading()

    out := ExtractResult(response)
    if (out != "") {
        out := StrReplace(out, "\n",  "`n")
        out := StrReplace(out, '\"',  '"')
        out := StrReplace(out, "\\",  "\")
        ; ensure Q:/A: pairs always start on their own line
        out := StrReplace(out, "Q:", "`nQ:")
        out := StrReplace(out, "A:", "`nA:")
        ; remove duplicate blank lines
        out := RegExReplace(out, "(`n){3,}", "`n`n")
        out := Trim(out, " `t`n`r")
        return out
    }

    ; ExtractResult returned "" — truly nothing usable found
    MsgBox "No result tags found in response.`n`nRaw output:`n" response,
           "Parse Warning", "Icon!"
    return ""
}

; ==========================================
; CLIPBOARD HELPER
; ==========================================
GetSelection() {
    saved := A_Clipboard
    A_Clipboard := ""
    Send "^c"
    if ClipWait(2)
        return A_Clipboard
    return (saved != "") ? saved : ""
}

; ==========================================
; ACTION DISPATCHER — called by both hotkeys and context menu
; ==========================================
RunAction(action, original) {
    if (original = "")
        return

    switch action {
        case "rewrite":
            instruction := 'You are an expert bilingual text editor (English and Hebrew).\nRewrite the provided text to be professional, clear, and grammatically correct, while maintaining the original meaning and language.\n1. Thinking phase: Analyse the text.\n2. Output phase: Return ONLY the final rewritten text wrapped strictly inside <result> and </result> tags.'
            label := "Rewrite -- Review & Edit"
            paste := true

        case "grammar":
            instruction := 'You are a careful proofreader. Fix ONLY spelling mistakes and grammatical errors in the text. Do NOT change the style, tone, vocabulary choice, or structure. Keep every sentence as close to the original as possible.\n1. Thinking phase: Identify errors.\n2. Output phase: Return ONLY the corrected text inside <result> and </result> tags.'
            label := "Grammar Fix -- Review & Edit"
            paste := true

        case "translate":
            instruction := 'You are a professional translator. Your ONLY job is to translate the given text.\nSTRICT RULES:\n- If the input is English: output the full Hebrew translation and nothing else.\n- If the input is Hebrew: output the full English translation and nothing else.\n- NEVER copy or echo back the original text.\n- NEVER add explanations, comments, or extra words.\n- NEVER refuse. Always translate.\nStep 1: Identify the input language.\nStep 2: Write the complete translation.\nStep 3: Wrap it inside <result> and </result> tags. Output nothing outside those tags.'
            label := "Translation -- Review & Edit"
            paste := true

                        case "summarise":
            instruction := 'You are an expert summariser.\nCRITICAL: plain text ONLY — no markdown, no asterisks, no bold, no hashtags, no backticks.\nWrite the summary in this exact structure:\n[One sentence overview]\n\n• [Key point one — one sentence]\n\n• [Key point two — one sentence]\n\n• [Key point three — one sentence]\n\n• [Add more bullets as needed, max 6]\n\n[One sentence conclusion]\nSame language as input. Nothing else. Wrap in <r> and </r>.'
            label := "Summary"
            paste := false

        case "explain":
            instruction := 'You are a patient teacher. Explain the text using a simple Q&A format.\nCRITICAL: Output plain text ONLY. No markdown, no asterisks, no bold, no hashtags, no backticks.\nUse this exact format:\nQ: [First question about a concept]\nA: [Answer in 1-3 plain sentences]\n\nQ: [Next question]\nA: [Answer]\n\n[Continue for all key concepts, max 5 pairs]\nRespond in the SAME language as the original text. Output ONLY the Q&A. Wrap inside <r> and </r> tags.'
            label := "Explanation"
            paste := false

        case "formal":
            instruction := 'You are an expert writer. Rewrite the text in a formal, professional tone suitable for business correspondence. Keep the original meaning and language.\n1. Thinking phase: Plan the rewrite.\n2. Output phase: Return ONLY the rewritten text inside <result> and </result> tags.'
            label := "Formal Rewrite -- Review & Edit"
            paste := true

        case "casual":
            instruction := 'You are an expert writer. Rewrite the text in a friendly, casual, conversational tone. Keep the original meaning and language.\n1. Thinking phase: Plan the rewrite.\n2. Output phase: Return ONLY the rewritten text inside <result> and </result> tags.'
            label := "Casual Rewrite -- Review & Edit"
            paste := true

        case "assertive":
            instruction := 'You are an expert writer. Rewrite the text in a confident, direct, assertive tone with no hedging language. Keep the original meaning and language.\n1. Thinking phase: Plan the rewrite.\n2. Output phase: Return ONLY the rewritten text inside <result> and </result> tags.'
            label := "Assertive Rewrite -- Review & Edit"
            paste := true

        case "fixlang":
            FixKeyboardLayout(original)
            return

        default:
            return
    }

    result := CallOllama(instruction, original, action)
    if (result = "")
        return

    LogInteraction(action, original, result)
    approved := ShowResultEditor(original, result, label)
    if (paste && approved != "") {
        A_Clipboard := approved
        Sleep 150
        Send "^v"
    }
}

; ==========================================
; FIX KEYBOARD LAYOUT — pure character map, no LLM
; Detects which direction the text is garbled and converts it.
; ==========================================
FixKeyboardLayout(text) {
    ; Standard Israeli Hebrew keyboard layout mappings
    ; English keys -> Hebrew chars (when typed with Hebrew layout active)
    engToHeb := Map(
        "q","/",  "w","'",  "e","ק",  "r","ר",  "t","א",
        "y","ט",  "u","ו",  "i","ן",  "o","ם",  "p","פ",
        "a","ש",  "s","ד",  "d","ג",  "f","כ",  "g","ע",
        "h","י",  "j","ח",  "k","ל",  "l","ך",
        "z","ז",  "x","ס",  "c","ב",  "v","ה",  "b","נ",
        "n","מ",  "m","צ",
        ",","ת",  ".","ץ",  ";","ף",  "/",".",
        "Q","/",  "W","'",  "E","ק",  "R","ר",  "T","א",
        "Y","ט",  "U","ו",  "I","ן",  "O","ם",  "P","פ",
        "A","ש",  "S","ד",  "D","ג",  "F","כ",  "G","ע",
        "H","י",  "J","ח",  "K","ל",  "L","ך",
        "Z","ז",  "X","ס",  "C","ב",  "V","ה",  "B","נ",
        "N","מ",  "M","צ"
    )

    ; Hebrew chars -> English keys (reverse map)
    hebToEng := Map()
    for eng, heb in engToHeb
        if !hebToEng.Has(heb)   ; avoid overwriting (Q/q both map to same heb)
            hebToEng[heb] := eng

    ; Detect direction: count how many chars have a mapping in each direction
    engScore := 0
    hebScore := 0
    loop StrLen(text) {
        ch := SubStr(text, A_Index, 1)
        if engToHeb.Has(ch)
            engScore++
        if hebToEng.Has(ch)
            hebScore++
    }

    ; Convert whichever direction scores higher
    result := ""
    if (engScore >= hebScore) {
        ; Typed English chars on Hebrew layout -> convert to Hebrew
        loop StrLen(text) {
            ch := SubStr(text, A_Index, 1)
            result .= engToHeb.Has(ch) ? engToHeb[ch] : ch
        }
    } else {
        ; Typed Hebrew chars on English layout -> convert to English
        loop StrLen(text) {
            ch := SubStr(text, A_Index, 1)
            result .= hebToEng.Has(ch) ? hebToEng[ch] : ch
        }
    }

    LogInteraction("fixlang", text, result)
    approved := ShowResultEditor(text, result, "Fix Keyboard Layout -- Review & Edit")
    if (approved != "") {
        A_Clipboard := approved
        Sleep 150
        Send "^v"
    }
}

; ==========================================
; RIGHT-CLICK CONTEXT MENU (system-wide)
; Intercepts right-click, checks if text is selected.
; If yes -> shows AI menu instead of normal context menu.
; If no  -> passes right-click through normally.
; ==========================================
~^RButton:: {   ; Ctrl + Right-click to open AI menu
    MouseGetPos(&mx, &my)
    KeyWait "RButton"

    aiMenu := Menu()
    aiMenu.Add("Rewrite && Improve",  (*) => RunAction("rewrite",   GetSelection()))
    aiMenu.Add("Fix Grammar Only",     (*) => RunAction("grammar",   GetSelection()))
    aiMenu.Add("Translate EN<->HE",    (*) => RunAction("translate", GetSelection()))
    aiMenu.Add("Fix Keyboard Gibberish",  (*) => RunAction("fixlang",   GetSelection()))
    aiMenu.Add("Summarise",            (*) => RunAction("summarise", GetSelection()))
    aiMenu.Add("Explain This",         (*) => RunAction("explain",   GetSelection()))
    aiMenu.Add()
    toneMenu := Menu()
    toneMenu.Add("Formal",    (*) => RunAction("formal",    GetSelection()))
    toneMenu.Add("Casual",    (*) => RunAction("casual",    GetSelection()))
    toneMenu.Add("Assertive", (*) => RunAction("assertive", GetSelection()))
    aiMenu.Add("Rewrite Tone", toneMenu)
    aiMenu.Add()
    aiMenu.Add("(Cancel / Normal Right-click)", (*) => Click("Right " mx " " my))

    aiMenu.Show(mx, my)
}

; ==========================================
; HOTKEYS (still available as before)
; ==========================================
^!r:: RunAction("rewrite",   GetSelection())
^!g:: RunAction("grammar",   GetSelection())
^!t:: RunAction("translate", GetSelection())
^!s:: RunAction("summarise", GetSelection())
^!e:: RunAction("explain",   GetSelection())
^!f:: RunAction("fixlang",   GetSelection())

; Tone cycle hotkey
toneIndex := 0
toneKeys  := ["formal", "casual", "assertive"]
^!w:: {
    global toneIndex, toneKeys
    toneIndex := Mod(toneIndex, 3) + 1
    RunAction(toneKeys[toneIndex], GetSelection())
}
