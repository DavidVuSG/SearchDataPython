 
 
; ------------------------- FORMAT VERSION
#Persistent

; Kích hoạt hotkey Alt+` để chèn văn bản "_001"
!`::
    Send, _001
return

; Kích hoạt hotkey Alt+SHIFT+` để chèn văn bản "_101"
!+`::
    Send, _101
return

; Kích hoạt hotkey Alt+2 để chèn văn bản "_002"
!2::
    Send, _002
return

; Kích hoạt hotkey Alt+3 để chèn văn bản "_003"
!3::
    Send, _003
return

; Kích hoạt hotkey Alt+shift+2 để chèn văn bản "_003"
!+2::
    Send, _201
return


; ----------------------------- ENDING FORMAT VERSION



; ------------------ INSERT CODE

!+p:: ; Alt + SHIFT+ P to activate; CODE : HIGHLIGH ROW AND COLS IN EXCEL VBA
code := "
(
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    If Target.Cells.Count > 1 Then Exit Sub
    Application.ScreenUpdating = False
    'Clear the color of all cells
    Cells.Interior.ColorIndex = 0
    With Target
        'Highlight row and column of the selected cell
        .EntireRow.Interior.ColorIndex = 42
        .EntireColumn.Interior.ColorIndex = 44
    End With
    Application.ScreenUpdating = True
End Sub
)"
Clipboard := code
Send "^v"
Return



; ------------------ ENDING INSERT CODE




;---------------------------- FORMAT TIME-----------

#X:: ; FORMAT TYPE: YYYYMMDD ĐỂ LÀM INSERT PHIẾU HÀNG: WIN+X
FormatTime, xx,, yyyyMMdd
SendInput, %xx%
return






+#x::  ; WIN+SHIFT+x ; ĐẦY ĐỦ NGÀY THÁNG, GIỜ
{
    FormatTime, xx, , dd/MM/yyyy, HH:mm:ss
    SendInput, %xx%
    return
}


#c:: ; WIN+C, insert ngày hiện tại
{
    FormatTime, xx, , dd-MM-yyyy 
    SendInput, %xx%
    return
}


#z:: ; win+z, insert ngày hiện tại theo định dạng DD.MM.YYYY 
{ 
	FormatTime, DateString,, dd.MM.yyyy 
	SendInput, %DateString% 
	return 
}


#!z:: ; Win + Alt + Z → insert yesterday (DD.MM.YYYY)
{
    Date := A_Now
    EnvAdd, Date, -1, Days
    FormatTime, Yesterday, %Date%, dd.MM.yyyy
    SendInput, %Yesterday%
    return
}



#!x::  ; WIN+ALT+x ; NGÀY THEO FORMAT US: M/D/YY
{
    FormatTime, xx, , M/d/yy
    SendInput, %xx%
    return
}


; ----------------- TẠO MÃ KIỂM SOÁT THEO NGÀY HIỆN TẠI: ĐỊNH DẠNG DWDMWMY --- UPDATE RULE
 
#NoEnv  ; Đề phòng lỗi môi trường
SendMode Input  ; Tăng tốc độ nhập
SetWorkingDir %A_ScriptDir%  ; Đảm bảo rằng thư mục làm việc là nơi lưu script

; Định nghĩa hàm lấy ngày trong tuần (thứ 2 = 2, thứ 3 = 3, ...)
Weekday(Date) {
    FormatTime, DayOfWeek, %Date%, dddd
    Days := {"Monday": 2, "Tuesday": 3, "Wednesday": 4, "Thursday": 5, "Friday": 6, "Saturday": 7, "Sunday": 1}
    return Days[DayOfWeek]
}

#!c::  ; Phím tắt Win+Alt+C
{
    FormatTime, Day, , dd
    FormatTime, Month, , MM
    FormatTime, Year, , yyyy
    Weekday := Weekday(A_Now)
    
    ; Lấy chữ số cuối của năm
    LastDigit := SubStr(Year, 4, 1)
    ; Chuyển sang số rồi tính công thức mới
    LastDigitYear := Mod(LastDigit + 5, 10) + 1

    DateFormat := SubStr(Day, 1, 1) . Weekday . SubStr(Day, 2, 1) . SubStr(Month, 1, 1) . Weekday . SubStr(Month, 2, 1) . LastDigitYear
    SendInput, %DateFormat%
    return
}




 
;----------------------------------- END: TẠO MÃ KIỂM SOÁT   ------------------


; ------------------------------------------- SOFTWARE-------------------------------------------------------

; - ---------------------------------EXCEL

#+q:: ; Win+Shift+Q
{
    Run, "C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE"
    Return  
}

; ------------------------------------ WORD

#+h:: ; WIN+SHIFT+h
{
	Run, "C:\Program Files (x86)\Microsoft Office\root\Office16\WINWORD.EXE"
	Return
}



; ------------------------------------ POWERPOINT

#+k:: ; WIN+SHIFT+k
{
	Run, "C:\Program Files (x86)\Microsoft Office\root\Office16\POWERPNT.EXE"
	Return
}


; ------------------------------------ ACCESS

#+J:: ; WIN+SHIFT+J
{
	Run, "C:\Program Files (x86)\Microsoft Office\root\Office16\MSACCESS.EXE"
	Return
}


#+U:: ; WIN+SHIFT+U
{
	Run, "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"
	Return
}




; ------------------------------------- SIMPLE NOTES

#+a:: ; Win+Shift+A
{
    Run, "C:\Program Files (x86)\Simnet\Simple Sticky Notes\ssn.exe"
    Return
}


; --------------------------------------- EDGE


#+e:: ; Win+Shift+E
{
    Run, "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
    Return 
}


; ---------------------------------------- CHROME
#+R:: ; WIN+SHIFT+R 
{
	RUN, "C:\Program Files\Google\Chrome\Application\chrome.exe"
	RETURN
}


; ---------------------------------ZALO

#+z:: ; Win+Shift+Z
{
    Run, "C:\Users\nhan\AppData\Local\Programs\Zalo\Zalo.exe"
    Return
}


#+l:: ; Win+shift+l: UNIKEY AND LIGHTSHOT
{
    Run, "C:\Program Files (x86)\Skillbrains\lightshot\Lightshot.exe"
    Run, "C:\Users\nhan\Documents\UniKeyNT.exe"
    return
}



#+m:: ; WIN+SHIFT+M
{
	Run, "C:\Program Files\SAP\FrontEnd\SAPGUI\saplogon.exe"
	Return
}



#+b:: ; WIN+SHIFT+B
{
	Run, "C:\Windows\System32\calc.exe"
	Return
}


#!+R:: ; WIN+SHIFT+ALT+R
{
	RUN, "C:\Program Files\Google\Chrome\Application\chrome.exe"
	RETURN
}



#+O:: ; wn+shift+O TO OPEN SIGNAL MESSENGER
{
	RUN, "C:\Users\nhan\AppData\Local\Programs\signal-desktop\Signal.exe"
	RETURN
}
; -----------------------------------------------------END:  sOFTWARE












; --------------------------------------- export files

#Persistent

; Kích hoạt hotkey Alt+Shift+b
!+b::

   ; Kích hoạt cửa sổ Edge 
   WinActivate, ahk_class Chrome_WidgetWin_1

    ; Bước 1: Di chuyển chuột đến vị trí (398, 194) và click: -------------------------check ALL
    MouseMove 344, 199
    Click
    

    ; Bước 2: Di chuyển chuột đến vị trí (524, 163) và click, sau đó đợi 1.2 giây------- CHeck EXPORT TO EXCEL
    MouseMove 524, 163
    Click
    Sleep 4200

    ; Bước 3: Di chuyển chuột đến vị trí (427, 168) và click ------------------------- CHECK NEXT TO 
    MouseMove 380, 168
    Click
    Sleep 4200

     

return


#Persistent


!+n::
; Loop 5 lần thực hiện phím tắt Alt+Shift+B
Loop, 5
{
    Send, !+b  ; Gửi tổ hợp phím Alt+Shift+B
    Sleep, 3500  ; Đợi 1 giây trước khi lặp lại
}

return


; -----------------------------------------end export files




; --------------------------------------- Copy and Replace b36.xls

; Copy selected file in Explorer, rename to b36.xls, and overwrite destination
; Hotkey: win+shift+f

#+f::
    selectedFile := Explorer_GetSelection()
    if (selectedFile = "")
    {
        MsgBox, 48, Error, No file selected!
        return
    }

    destFile := "C:\Users\nhan\Documents\Python Notebook\ProjectSQLPython\b36.xls"

    FileCopy, %selectedFile%, %destFile%, 1  ; 1 = overwrite
    MsgBox, 64, Done, Copied and replaced as b36.xls in ProjectSQLPython folder!
return

; Function to get selected file from Explorer
Explorer_GetSelection() {
    for window in ComObjCreate("Shell.Application").Windows
    {
        if (window.hwnd = WinActive("A"))
        {
            sel := window.document.SelectedItems
            if (sel.Count = 0)
                return ""
            return sel.Item(0).Path
        }
    }
    return ""
}

; -------------------------------------------- end copy and replace b36.xls


; --------------------------------------- Copy and Replace b37.xls

; Copy selected file in Explorer, rename to b37.xls, and overwrite destination
; Hotkey: win+shift+g

#+g::
    selectedFile := Explorer_GetSelection()
    if (selectedFile = "")
    {
        MsgBox, 48, Error, No file selected!
        return
    }

    destFile := "C:\Users\nhan\Documents\Python Notebook\ProjectSQLPython\b37.xls"

    FileCopy, %selectedFile%, %destFile%, 1  ; 1 = overwrite
    MsgBox, 64, Done, Copied and replaced as b37.xls in ProjectSQLPython folder!
return



; -------------------------------------------- end copy and replace b37.xls