Dim speech,greeting,randomstatement
Dim currenthour
Dim max,min
Dim randomno
max=3
min=1

Function getrandomno(min,max)
    Randomize
    getrandomno=(Int((max-min+1)*Rnd+min))
End Function

currenthour=Hour(Now())

if currenthour<=12 then
    randomno=getrandomno(min,max)
    Select Case randomno
        Case 1
	randomstatement="You're looking sharp today."
        Case 2
	randomstatement="Rise and shine!"
        Case Else
	randomstatement="Had your coffee yet?"
    End Select

    greeting="Good morning, "
end if

if currenthour>12 and currenthour <=18 then
    randomno=getrandomno(min,max)
    Select Case randomno
        Case 1
	randomstatement="Hard at work, I see."
        Case 2
	randomstatement="Lovely weather!"
        Case Else
	randomstatement="I love your hair."
    End Select

    greeting="Good afternoon, "
end if

if currenthour>18 then
    randomno=getrandomno(min,max)
    Select Case randomno
        Case 1
	randomstatement="Isn't it a little late?"
        Case 2
	randomstatement="Time for bed!"
        Case Else
	randomstatement="The moon's out tonight."
    End Select

    greeting="Good evening, "
end if

speech=greeting & ", Teochew Thunder. " & randomstatement
Set VObj=CreateObject("sapi.spvoice")
VObj.Speak speech
