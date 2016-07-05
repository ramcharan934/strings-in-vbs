dim i,j,str,tempstr,tempstr2,display,count,copystr
str=inputbox("enter the text to count the frequency of the characters")
copystr=str
strlength=len(str)
'msgbox(strlength)
i=1
while(i<=strlength)
	tempstr=mid(str,i,1) 
	j=1
	count=0
	if(tempstr<>"0")then
		while(j<=strlength )
			tempstr2=mid(str,j,1) 
			if(tempstr=tempstr2)then
				count=count+1
			end if
			j=j+1
		wend
	str=replace(str,tempstr,"0")
	'msgbox str
	display=display&" "&tempstr&" - "&" "&count&vbcr
	end if
	
	
i=i+1
wend
msgbox "'"&copystr&"'"&" has"&vbcr&display
