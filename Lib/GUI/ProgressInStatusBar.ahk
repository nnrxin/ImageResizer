; ==================================================================================================
; Creates a Progress control embedded in a StatusBar
; Parameters:
;     SB       -  StatusBar object
;     Value    -  Initial position of the Progress (integer)
;                 Default: 0
;     Part     -  Number of the StatusBar part to embed the Progress (integer)
;                 Default: 1
;     Options  -  Object containing one or more of the Progress class properties like
;                 BackColor, BarColor, Range, etc.   
; Return value:
;     Returns on object of class ProgressInStatusBar.Progress on success.
;     The object has the following properties and methods:
;        AdjustSize()   -  Adjusts the size of the control after changes to the StatusBar.
;        SetMarquee()   -  Sets marquee mode of the control on or off.
;        StepIt()       -  Increments the end position of the bar by the current value of Step.
;        BackColor      -  Retrieves or sets the controls background color.
;        BarColor       -  Retrieves or sets the controls bar color.
;        State          -  Retrieves or sets the state of the control.
;        Step           -  Retrieves or sets the increment used by StepIt()- Default: 10
;        Enabled        -  Retrieves or sets the interaction state of the control. 
;        Range          -  Retrieves or sets the range of the control - Default: 0-100
;        Value          -  Retrieves or sets value of the control (end position of the bar) 
;        Visible        -  Retrieves or sets the visibility state of the control.
; License:
;        The Unlicense  -> https://unlicense.org/
; ==================================================================================================
Class ProgressInStatusBar {
    ; Construction and destruction ==================================================================
    Static Call(SB, Value := 0, Part := 1, Options := "") {
       Static PBStyle := 0x50000001 ; forced dwStyle WS_CHILD|WS_VISIBLE|PBS_SMOOTH
       If !(SB Is Gui.StatusBar)
          Throw TypeError("Parameter SB must be a Gui.Statusbar but is " . Type(SB), -1)
       Local HSB := SB.Hwnd
       Local NumOfParts := SendMessage(0x0406, 0, 0, HSB) ; SB_GETPARTS
       If (Part < 1) || (Part > NumOfParts)
          Throw ValueError("Parameter Part is invalid", -1, Part)
       Local PartIndex := Part - 1
       ; Get the segment's dimensions
       Local RC := Buffer(16, 0)
       If !SendMessage(0x040A, PartIndex, RC.Ptr, HSB) ; SB_GETRECT
          Throw Error("Couldn't get the rectangle of part " . Part, -1)
       ControlSetStyle("+0x02000000", HSB) ; WS_CLIPCHILDREN
       Local PBX := NumGet(RC,  0, "Int"),
             PBY := NumGet(RC,  4, "Int"),
             PBW := NumGet(RC,  8, "Int") - PBX,
             PBH := NumGet(RC, 12, "Int") - PBY
       Local HPB := DllCall("User32.dll\CreateWindowEx",
                            "UInt", 0,
                            "Str", "msctls_progress32",
                            "Ptr", 0,
                            "UInt", PBStyle,
                            "Int", PBX,
                            "Int", PBY,
                            "Int", PBW,
                            "Int", PBH,
                            "Ptr", HSB,
                            "Ptr", 0,
                            "Ptr", 0,
                            "Ptr", 0,
                            "UInt")
       ; Create a Progress object
       Local PB := ProgressInStatusBar.Progress(HPB, HSB, PartIndex)
       ; Set the initial pos if specified
       If IsInteger(Value) && Integer(Value > 0)
          PB.Value := Value
       ; Evaluate the options, if any
       If IsObject(Options) && (Type(Options) = "Object") {
          For K, V In Options.OwnProps() {
             If PB.HasProp(K)
                PB.%K% := V
          }
       }
       Return PB
    }
    ; ===============================================================================================
    Class Progress {
       __New(HWnd, HSB, PartIndex) {
          This.Hwnd := HWnd
          This.HSB := HSB
          This.PartIndex := PartIndex
       }
        ; --------------------------------------------------------------------------------------------
       __Delete() {
          If This.HasProp("Hwnd")
             DllCall("DestroyWindow", "Ptr", This.Hwnd)
       }
       ; ============================================================================================
       ; Own methods ================================================================================
       ; ============================================================================================
       ; Recalculates the size of the control after the size of the Statusbar changed
       AdjustSize() {
          ; SB_GETRECT = 0x040A
          Local RC := Buffer(16, 0)
          If !SendMessage(0x040A, This.PartIndex , RC.Ptr, This.HSB)
             Throw Error("Coudn't get the rectangle of part " . (This.PartIndex + 1), -1)
          Local PBX := NumGet(RC,  0, "Int"),
                PBY := NumGet(RC,  4, "Int"),
                PBW := NumGet(RC,  8, "Int") - PBX,
                PBH := NumGet(RC, 12, "Int") - PBY
          DllCall("User32.dll\MoveWindow",
                  "Ptr", This.Hwnd,
                  "Int", PBX,                
                  "Int", PBY,                
                  "Int", PBW,                
                  "Int", PBH,                
                  "UInt", True)  
       }
       ; --------------------------------------------------------------------------------------------
       ; Sets marquee mode of the control on or off.
       ;  MS -  Time, in milliseconds, between marquee animation updates.
       ;        Pass 0 to reset to the default (30 milliseconds). 
       SetMarquee(Marquee := True, MS := 0) {
          ; PBM_SETMARQUEE = 0x040A, PBS_MARQUEE = 0x08
          ControlSetStyle((!!Marquee ? "+" : "-") . "0x08", This.Hwnd)         
          Return SendMessage(0x040A, !!Marquee, MS, This.Hwnd)
       }
       ; --------------------------------------------------------------------------------------------
       ; Increment the value/pos of the bar by one step.
       StepIt() {
          ; PBM_STEPIT = 0x0405
          Return SendMessage(0x0405, 0, 0, This.Hwnd)
       }   
       ; ============================================================================================
       ; Own Properties =============================================================================
       ; ============================================================================================
       ; Retrieves or sets the background color of the control. The colour must be specified as HTML
       ; color name or RGB integer value (e.g. 0xFFFF00). Pass an empty string ("") to reset to the
       ; default color.
       BackColor {
          ; PBM_GETBKCOLOR = 0x040E, PBM_SETBKCOLOR = 0x2001, CLR_DEFAULT = 0xFF000000
          Get {
             Local BkGColor := SendMessage(0x040E, 0, 0, This.Hwnd) & 0xFF000000
             Return (BkGColor = 0xFF000000 ? "" : This.RGB(BkGColor))
          }
          Set {
             DllCall("UxTheme.dll\SetWindowTheme", "Ptr", This.Hwnd, "Str", "", "Str", "")
             SendMessage(0x2001, 0, Value = "" ? 0xFF000000 : This.BGR(Value), This.Hwnd)
          }
       }
       ; --------------------------------------------------------------------------------------------
       ; Retrieves or sets the bar color of the control. For details see BackColor.
       BarColor {
          ; PBM_GETBARCOLOR = 0x040F, PBM_SETBARCOLOR = 0x0409, CLR_DEFAULT = 0xFF000000
          Get {
             Local BarColor := SendMessage(0x040F, 0, 0, This.Hwnd) & 0xFF000000
             Return (BarColor = 0xFF000000 ? "" : This.RGB(BarColor))
          }
          Set {
             DllCall("UxTheme.dll\SetWindowTheme", "Ptr", This.Hwnd, "Str", "", "Str", "")
             SendMessage(0x0409, 0, Value = "" ? 0xFF000000 : This.BGR(Value), This.Hwnd) ; 
          }
       }
       ; --------------------------------------------------------------------------------------------
       ; Retrieves or sets the state of the control
       State {
          ; PBM_GETSTATE = 0x0411, PBM_SETSTATE = 0x0410
          ; PBST_NORMAL = 1, PBST_ERROR = 2, PBST_PAUSED = 3
          Get => SendMessage(0x0411, 0, 0, This.Hwnd)
          Set {
             If IsInteger(Value) && (Value > 0) && (Value < 4) {
                DllCall("LockWindowUpdate", "Ptr", This.Hwnd)
                SendMessage(0x0410, 1, 0, This.Hwnd)
                If !(Value = 1)
                   SendMessage(0x0410, Integer(Value), 0, This.Hwnd)
                DllCall("LockWindowUpdate", "Ptr", 0)
                Return 1
             }
             Return 0
          }
       }
       ; --------------------------------------------------------------------------------------------
       ; Retrieves or sets the increment used by StepIt()
       Step {
          ; PBM_GETSTEP = 0x040D, PBM_SETSTEP = 0x0404
          Get => SendMessage(0x040D, 0, 0, This.Hwnd) << 32 >> 32
          Set => (IsInteger(Value) ? SendMessage(0x0404, Value, 0, This.Hwnd) : "")
       }
       ; ============================================================================================
       ; Gui.Control like properties ============================================================
       ; ============================================================================================
       ; Retrieves or sets the interaction state of the control.
       Enabled {
          Get => ControlGetEnabled(This.Hwnd)
          Set => ControlSetEnabled(Value ? 1 : 0, This.Hwnd)
       }
       ; --------------------------------------------------------------------------------------------
       ; Retrieves or sets the value range of the control.
       ; https://www.autohotkey.com/docs/v2/lib/GuiControls.htm#Progress_Options -> Range
       Range {
          ; PBM_GETRANGE = 0x0407, PBM_SETRANGE32 = 0x0406
          Get {
             Local PBRANGE := Buffer(8, 0)
             SendMessage(0x0407, 0, PBRANGE.Ptr, This.Hwnd)
             Return NumGet(PBRANGE, 0, "Int") . "-" . NumGet(PBRANGE, 4, "Int")
          }
          Set {
             Local S1 := 1, S2 := 1, Min, Max
             If (SubStr(Value, 1, 1) = "-") {
                Value := SubStr(Value, 2)
                S1 := -1
             }
             Local P := InStr(Value, "-")
             If (P = 0)
                Return False
             If (SubStr(Value, P + 1, 1) = "-") {
                Value := SubStr(Value, 1, P) . SubStr(Value, P + 2)
                S2 := -1
             }
             Local R := StrSplit(Value, "-")
             If !(R.Length = 2) || !IsInteger(R[1]) || !IsInteger(R[2])
                Return 0
             LocaL Min := R[1] * S1
             Local Max := R[2] * S2
             If !(Min < Max)
                Return 0
             Return SendMessage(0x0406, Min, Max, This.Hwnd)
          }
       }
       ; --------------------------------------------------------------------------------------------
       ; Retrieves or sets value of the control (end position of the bar) 
       Value {
          ; PBM_GETPOS = 0x0408, PBM_SETPOS = 0x0402
          Get => SendMessage(0x0408, 0, 0, This.Hwnd) << 32 >> 32
          Set => SendMessage(0x0402, Value, 0, This.Hwnd)
       }
       ; --------------------------------------------------------------------------------------------
       ; ; Retrieves or sets the visibility state of the control.
       Visible {
          Get => ControlGetVisible(This.Hwnd)
          Set => (Value ? ControlShow(This.Hwnd) : ControlHide(This.Hwnd))
       }        
       ; ============================================================================================
       ; Internal functions =========================================================================
       ; ============================================================================================
       BGR(RGB) { ; Converts a numeric RGB value or a HTML color name to BGR
          Static HTML := {BLACK:  0x000000, SILVER: 0xC0C0C0, GRAY:   0x808080, WHITE:   0xFFFFFF
                        , MAROON: 0x000080, RED:    0x0000FF, PURPLE: 0x800080, FUCHSIA: 0xFF00FF
                        , GREEN:  0x008000, LIME:   0x00FF00, OLIVE:  0x008080, YELLOW:  0x00FFFF
                        , NAVY:   0x800000, BLUE:   0xFF0000, TEAL:   0x808000, AQUA:    0xFFFF00}
          If HTML.HasProp(RGB)
             Return HTML.%RGB%
          Return ((RGB & 0xFF0000) >> 16) + (RGB & 0x00FF00) + ((RGB & 0x0000FF) << 16)
       }
       ; --------------------------------------------------------------------------------------------
       RGB(BGR) { ; Converts a numeric BGR value to RGB
          Return ((BGR & 0xFF0000) >> 16) + (BGR & 0x00FF00) + ((BGR & 0x0000FF) << 16)
       }
    }
 }