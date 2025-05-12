#Requires AutoHotkey v2.0
#SingleInstance Force

; Constants for Windows messages
WM_INPUTLANGCHANGEREQUEST := 0x0050
INPUTLANGCHANGE_FORWARD := 0x0002

; Create tray menu
A_TrayMenu.Delete()  ; Clear default menu
A_TrayMenu.Add("Start with Windows", ToggleStartup)
A_TrayMenu.Add("Exit", ExitScript)
A_TrayMenu.Default := "Start with Windows"  ; Set default option

; Check if startup is enabled and update menu
UpdateStartupMenu() {
    startupEnabled := IsStartupEnabled()
    A_TrayMenu.Rename("Start with Windows", startupEnabled ? "Disable Startup" : "Enable Startup")
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

; Initialize startup menu state
UpdateStartupMenu()

; Remap CapsLock to switch keyboard layout
CapsLock:: {
    SwitchKeyboardLayout()
}

; Prevent CapsLock from being sent to the system
CapsLock Up:: {
    return
}
