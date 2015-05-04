import xlsxwriter
import sys ,os
import re
print """\
                                                  __    __ 
                                                 /  |  /  |
 _______   _____  ____    ______    ______       ## |  ## |
/       \ /     \/    \  /      \  /      \      ##  \/##/ 
#######  |###### ####  | ######  |/######  |      ##  ##<  
## |  ## |## | ## | ## | /    ## |## |  ## |       ####  \ 
## |  ## |## | ## | ## |/####### |## |__## |      ## /##  |
## |  ## |## | ## | ## |##    ## |##    ##/_____ ## |  ## |
##/   ##/ ##/  ##/  ##/  #######/ #######/      |##/   ##/ 
                                  ## |   ######/           
                                  ## |                     
                                  ##/
                                      v2 Text Converter

                    Nmap Text output to Excel Converter !!

                                         - Coded By R4m

                                       
"""
try:
    print "Please wait .....  "
    if (sys.argv[1].isspace() or sys.argv[2].isspace() or sys.argv[1]=="" or sys.argv[2]=="" or (os.path.isdir(sys.argv[1]) or os.path.isdir(sys.argv[2]))):
        print "Usage : nmap_x.py <Path to Nmap Text Output> <Path for Excel Export> "
    elif (sys.argv==-1)
        #print "Usage : nmap_x.py <Path to Nmap Text Output> <Path for Excel Export> "
    #else:
    files=open("a.txt.txt",'r')
    xl=xlsxwriter.Workbook("Nmapreport.xlsx")
    setbord=xl.add_format()
    setbord.set_border()
    setbord.set_align('center')
    setbord.set_align('vcenter')
    color=xl.add_format()
    color.set_border()
    color.set_align('center')
    color.set_align('vcenter')
    color.set_pattern(1)
    color.set_bg_color('#FFC000')
    color.set_bold()
    wrksht=xl.add_worksheet('Report')
    wrksht.set_column(1,1,30)
    wrksht.set_column(1,4,30)
    wrksht.write(0,0,"S.NO",color)
    wrksht.write(0,1,"IP Address",color)
    wrksht.write(0,2,"Port",color)
    wrksht.write(0,3,"Protocol",color)
    wrksht.write(0,4,"State",color)
    wrksht.write(0,5,"Service",color)

            
    row=1
    col=0
    S_no=1
    start="i"
    end="u"
    pattern = r'^(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)$'
    while True :
        
        while True:
            end=files.tell()
            strs=files.readline()
            start=files.tell()
            if strs.startswith("Nmap scan report for "):
                ip=strs.split()[4]
                check=strs.split()[0]
                break
            elif (start==end):
                raise Exception("")
                
        merge_start=row+1
        while True:
                strs=files.readline()
                check=strs.split()[0]
                if check=="PORT":
                        break
        strs=files.readline()
        while (strs!="\n"):
            
            store=strs.split()
            
            port=store[0].__getslice__(0,store[0].find("/"))
            if(port.isdigit()):
                protocol_used=store[0].__getslice__(store[0].find("/")+1,store[0].__len__())
                service=store[2]
                state=store[1]                                                                                          
                wrksht.write(row,2,int(port),setbord)
                wrksht.write(row,3,protocol_used.upper(),setbord)
                wrksht.write(row,4,state,setbord)
                wrksht.write(row,5,service,setbord)
                strs=files.readline()
                
                row+=1
            else:
                break
        merge_end=row
        if merge_end - merge_start!= 0:
            wrksht.merge_range("A"+str(merge_start)+":"+"A"+str(merge_end),S_no,setbord)
            wrksht.merge_range("B"+str(merge_start)+":"+"B"+str(merge_end),ip,setbord)
        else:
            wrksht.write(row-1,0,S_no,setbord)
            wrksht.write(row-1,1,ip,setbord)
        S_no+=1
        end=files.tell()
except Exception as e:
    print "Job Successfully Done !!! :) :D"
    xl.close()
    files.close()

