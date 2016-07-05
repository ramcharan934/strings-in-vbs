'demo string

set fso = createobject("scripting.filesystemobject")
filename="c:\my\note4.txt"
if(fso.fileexists(filename))then 
	set opentxtfile = fso.opentextfile(filename,1)
	msgbox "text opened"
else
	msgbox "no input txt fileexists"
end if

i=1
do
	tempstr= opentxtfile.readline()
	temparray=split(tempstr,",")
	j=1
	for each temp in temparray
		msgbox temp
		j=j+1
	next
	i=i+1	
loop until(opentxtfile.atendofstream)

opentxtfile.close()