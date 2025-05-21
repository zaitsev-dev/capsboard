#Requires AutoHotkey v2.0
#SingleInstance Force

; Constants for Windows messages
WM_INPUTLANGCHANGEREQUEST := 0x0050
INPUTLANGCHANGE_FORWARD := 0x0002

; Registry key for settings
SETTINGS_KEY := "HKEY_CURRENT_USER\Software\CapsBoard"

; Global variable for switching mode
global useWinSpaceMode := LoadModePreference()

; Function to load mode preference from registry
LoadModePreference() {
    try {
        return RegRead(SETTINGS_KEY, "UseWinSpaceMode") == "1"
    } catch {
        return false
    }
}

; Function to save mode preference to registry
SaveModePreference() {
    try {
        RegWrite(useWinSpaceMode ? "1" : "0", "REG_SZ", SETTINGS_KEY, "UseWinSpaceMode")
    } catch as err {
        MsgBox "Failed to save mode preference: " . err.Message
    }
}

; Function to create tray menu
CreateTrayMenu() {
    A_TrayMenu.Delete()  ; Clear default menu
    A_TrayMenu.Add("Start with Windows", ToggleStartup)
    A_TrayMenu.Add("Alternative Mode", ToggleSwitchMode)
    if (useWinSpaceMode)
        A_TrayMenu.Check("Alternative Mode")
    A_TrayMenu.Add("Exit", ExitScript)
    A_TrayMenu.Default := "Start with Windows"  ; Set default option
}

; Function to toggle switching mode
ToggleSwitchMode(*) {
    global useWinSpaceMode
    useWinSpaceMode := !useWinSpaceMode
    SaveModePreference()
    CreateTrayMenu()
    UpdateStartupMenu()
}

; Function to toggle startup
ToggleStartup(*) {
    startupEnabled := IsStartupEnabled()
    if (startupEnabled) {
        DisableStartup()
    } else {
        EnableStartup()
    }
    UpdateStartupMenu()
}

; Function to check if startup is enabled and update menu
UpdateStartupMenu() {
    startupEnabled := IsStartupEnabled()
    A_TrayMenu.Rename("Start with Windows", startupEnabled ? "Disable Startup" : "Enable Startup")
}

; Function to check if startup is enabled
IsStartupEnabled() {
    try {
        startupKey := "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run"
        return RegRead(startupKey, "CapsBoard") != ""
    } catch {
        return false
    }
}

; Function to enable startup
EnableStartup() {
    try {
        startupKey := "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run"
        scriptPath := A_ScriptFullPath
        RegWrite(scriptPath, "REG_SZ", startupKey, "CapsBoard")
    } catch as err {
        MsgBox "Failed to enable startup: " . err.Message
    }
}

; Function to disable startup
DisableStartup() {
    try {
        startupKey := "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run"
        RegDelete(startupKey, "CapsBoard")
    } catch as err {
        MsgBox "Failed to disable startup: " . err.Message
    }
}

; Function to exit script
ExitScript(*) {
    ExitApp
}

; Function to switch keyboard layout using PostMessage
SwitchKeyboardLayout() {
    global useWinSpaceMode
    if (useWinSpaceMode) {
        Send "#{Space}"
        return
    }

    try {
        ; Get the current keyboard layout
        currentLayout := DllCall("GetKeyboardLayout", "UInt", 0, "UInt")
        
        ; Get the number of keyboard layouts
        layoutCount := DllCall("GetKeyboardLayoutList", "UInt", 0, "Ptr", 0, "UInt")
        
        if (layoutCount > 1) {
            ; Create buffer for layout list
            layoutList := Buffer(layoutCount * 8)
            
            ; Get the list of keyboard layouts
            DllCall("GetKeyboardLayoutList", "UInt", layoutCount, "Ptr", layoutList, "UInt")
            
            ; Find the next layout
            nextLayout := 0
            found := false
            
            Loop layoutCount {
                layout := NumGet(layoutList, (A_Index - 1) * 8, "UInt")
                if (layout == currentLayout) {
                    found := true
                } else if (found) {
                    nextLayout := layout
                    break
                }
            }
            
            ; If we haven't found the next layout, start from the beginning
            if (!nextLayout) {
                nextLayout := NumGet(layoutList, 0, "UInt")
            }
            
            ; Get the foreground window
            hwnd := WinExist("A")
            
            ; Send the layout change request
            DllCall("PostMessage", 
                "Ptr", hwnd, 
                "UInt", WM_INPUTLANGCHANGEREQUEST, 
                "UInt", INPUTLANGCHANGE_FORWARD, 
                "Ptr", 0)
        }
    } catch as err {
        MsgBox "Error switching keyboard layout: " . err.Message
    }
}

; Remap CapsLock to switch keyboard layout
CapsLock:: {
    SwitchKeyboardLayout()
}

; Prevent CapsLock from being sent to the system
CapsLock Up:: {
    return
}

; Initialize menu
CreateTrayMenu()
UpdateStartupMenu()
