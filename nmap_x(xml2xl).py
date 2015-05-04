from elementtree import ElementTree
import xlsxwriter
import sys ,os 



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

                    Nmap XML output to Excel Converter !!

                                         - Coded By SuRam123
"""



try :
    if sys.argv.__len__() <=1 or sys.argv[1].isspace() or sys.argv[2].isspace() or sys.argv[1]=='' or sys.argv[2]=='' :
        print "Usage : nmap_x.py <Path to Nmap XML Output> <Path for Excel Export> \n eg: nmap_xv2.py C:\nmap\nmap_out.xml C:\nmap\Excelreport.xlsx"
        


    else :
        print "Please wait .....  "
        arg2=sys.argv[2]
        #copy from nmap8000
        if sys.argv[2].__getitem__(sys.argv[2].__len__()-1)=="/" or sys.argv[2].__getitem__(sys.argv[2].__len__()-1)=="\\" :
            arg2+="default.xlsx"
        xl=xlsxwriter.Workbook(arg2)
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
        wrksht.write(0,0,"S.NO",color)
        wrksht.write(0,1,"IP Address",color)
        wrksht.write(0,2,"Port",color)
        wrksht.write(0,3,"Protocol",color)
        wrksht.write(0,4,"Service",color)
        root=ElementTree.parse(sys.argv[1]).getroot()
                
        row=1
        col=0

        host=root.findall("host")
        for i in range(host.__len__()) :
            merge_start=row+1
            ip_address=host[i].getchildren()[1].attrib["addr"]
            ports=host[i].getchildren()[3].findall("port")
            for j in range(ports.__len__() -1):
                port_no=ports[j].attrib["portid"]
                protocol_used=ports[j].attrib["protocol"]
                state=ports[j].getchildren()[0].attrib["state"]
                service=ports[j].getchildren()[1].attrib["name"]
                col=2
                wrksht.write(row,col,int(port_no),setbord)
                col+=1
                wrksht.write(row,col,protocol_used.upper(),setbord)
                col+=1
                col+=1
                wrksht.write(row,col,state.upper(),setbord)
                wrksht.write(row,col,service,setbord)
                row+=1
            merge_end=row
            wrksht.merge_range("A"+str(merge_start)+":"+"A"+str(merge_end),i+1,setbord)
            wrksht.merge_range("B"+str(merge_start)+":"+"B"+str(merge_end),ip_address,setbord)
           
            
        xl.close()
        print "Job Successful :) :D "
except Exception:
    try:
        xl.close()
    except Exception:
        pass
    print "Incorrect Arguments !!"
    try :
        os.remove(arg2)
        
    except :
            pass
    print "Usage : nmap_x.py <Path to Nmap XML Output> <Path for Excel Export> \n eg: nmap_xv2.py C:\nmap\nmap_out.xml C:\nmap\Excelreport.xlsx"
   

