Dim Num,Cube,Sum 
Num = InputBox("Enter the number")
Num = CInt(Num)
NumTemp = Num
Cube = 0 
N = 0 
Sum = 0

'used timer to get the stat and end time of script

StartTime = Timer()
While (NumTemp > 0)
    N = NumTemp Mod 10
    Cube = N*N*N
    Sum =  Sum + Cube
    NumTemp = NumTemp\10 
Wend
If Sum=Num Then
    msgbox(CStr(Num) +" is armstrong number")
Else 
    msgbox(CStr(Num) +" is not a armstrong number")
End If
EndTime = Timer()

msgbox("This logic takes " +CStr(EndTime-StartTime) +" secs" )