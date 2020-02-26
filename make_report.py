#!/usr/bin/python3

import openpyxl
import os
import shutil
import sys

def read_excel_city(fname, sname, city):   
    print("read the data file for the defined city")    
    #open the workbook
    wb    = openpyxl.load_workbook(fname)
    sheet = wb[sname]

    max_row = sheet.max_row
    max_col = sheet.max_column

    #get the headers (as a key for a dict)
    city_data    = {}
    keys         = []    
    maincity_col = 0
    for i in range(1,max_col+1,1):  #add 1 b/c of range
        key = sheet.cell(row=1,column=i).value
        keys.append(key)
        if key == 'maincity':
            maincity_col = i
        #endif
    #endfor

    if 'maincity' in keys:
        found_city = False
        city_row   = 0
        for i in range(1,max_row+1,1): #add 1 b/c of range
            maincity = sheet.cell(row=i,column=maincity_col).value
            if maincity == city:
                found_city = True
                city_row   = i
            #endif
        #endor
        if found_city:
            for i in range(1,max_col+1,1): #add 1 b/c of range
                val = sheet.cell(row=city_row,column=i).value
                city_data[keys[i-1]]=val
            #endfor
        else:
            print("Error - maincity (" + city + ") not found - check your spelling.")
            sys.exit()
        #endif
    else:
        print("Error - column header 'maincity' not found - check xlsx file.")
        sys.exit()
    #endif
    return(city_data)
#enddef

def read_codebook(fname,sname,code):

    #translate code from Results sheet to that of Codebook
    #I already changed some that ended in RWJ
    if code == "sdg3v8_lifeExpectancy":
        code = "sdg3v8_le"
    #endif
    if code == "sdg4v4_HSgrad":
        code = "sdg4v4_Hsgrad"
    #endif
    
    print("reading the codebook")    
    #open the workbook
    wb    = openpyxl.load_workbook(fname)
    sheet = wb[sname]

    max_row = sheet.max_row
    max_col = sheet.max_column

    #get the headers (as a key for a dict)
    code_data     = {}
    keys          = []    
    indicator_col = 0
    for i in range(1,max_col+1,1):  #add 1 b/c of range
        key = sheet.cell(row=1,column=i).value
        keys.append(key)
        if key == 'Indicator':
            indicator_col = i
        #endif
    #endfor
    
    if 'Indicator' in keys:
        found_code = False
        code_row   = 0
        for i in range(1,max_row+1,1): #add 1 b/c of range
            maincode = sheet.cell(row=i,column=indicator_col).value
            if maincode == code:
                found_code = True
                code_row   = i
            #endif
        #endfor
        
        if found_code:
            for i in range(1,max_col+1,1): #add 1 b/c of range
                val = sheet.cell(row=code_row,column=i).value
                code_data[keys[i-1]]=val
            #endfor
        else:
            print("Error - Indicator (" + code + ") not found - check your spelling.")
            sys.exit()
        #endif
    else:
        print("Error - column header 'Indicator' not found - check xlsx file.")
        sys.exit()
    #endif

#    for k,v in code_data.items():
#        print(k,":",v)
    
    
    return(code_data)
#enddef

#here's where the BIG math is done
def get_color_status(code_data, city_data, target):
    city_val = city_data[target]
    print(target,city_val)
    color = "red" #default
    
    sorder    = code_data['Sort Order']
    to_orange = code_data['To Orange']
    to_yellow = code_data['To Yellow']
    to_green  = code_data['To Green']

    print(city_val,sorder, to_orange,to_yellow,to_green)

    if city_val is None:
        color = "gray"
    else:    
        if "asc" in sorder:  #ascending, higher is better
            if city_val >= to_green:
                color = "green"
            elif to_green > city_val and city_val >= to_yellow:
                color = "yellow"
            elif to_yellow > city_val and city_val >= to_orange:
                color = "orange"
            #else color remains red
            #endif
        else:                #descending, lower is better
            if city_val <= to_green:
                color = "green"
            elif to_green < city_val and city_val <= to_yellow:
                color = "yellow"
            elif to_yellow < city_val and city_val <= to_orange:
                color = "orange"
            #else color is still red        
            #endif
        #endif
    #endif
    
#    print(color)

    return(color)
#enddef

def fix_val(nval):
    if nval is None:
        nval = ""
    else:
        nval = round(nval,1)
    #endif
    return(nval)
#enddef


#----------------------------------------------------------------
# main latex generator
#----------------------------------------------------------------
def make_goal_figs(rep_dir,city_data):        

    for i in range(1,18,1):
        sdgx = "sdg"+str(i)
        print("create figure for SDG: " + sdgx)

        fig_data = []
        
        sdgxv = sdgx + "v"
        for k,v in city_data.items():
            if k.startswith(sdgxv):
                print("get data for target:" + k)
                code_data = get_code_data(k)
                target = k
                color  = get_color_status(code_data,city_data,target)
                iname  = code_data['Indicator name']
                val    = city_data[target]  #measured value
                nval   = city_data["n_"+k]  #normalized value - used to calc goal score
                val  = fix_val(val)
                nval = fix_val(nval)
                
                fig_data.append([target,color,iname,code_data,nval,val])
            #endif
        #endfor

        print("making the figure for SDG"+str(i))
        #MAKE THE FIGURE HERE
        figtex = ""
        figtex = figtex + "\\begin{table}[ht]\n"
        figtex = figtex + "\\begin{tabularx}{\\textwidth}{\n"
        for j in range(0,len(fig_data)):
            figtex = figtex + "    	            X>{\\columncolor{white}[0.5\\tabcolsep]}\n"
        #endfor
        figtex = figtex + "    	            X}\n"
        figtex = figtex + "\\toprule\n"
        figtex = figtex + "\\multirow{2}{*}{\includegraphics[height=2em]{../figs/Goal"+str(i)+".pdf}} \n"
        for k in range(0,len(fig_data)):
            iname = fig_data[k][2]
            figtex = figtex + " & " + iname
        #endfor
        figtex = figtex + "\\\\ \n"
        for k in range(0,len(fig_data)):
            color = fig_data[k][1]
            nval  = fig_data[k][4]
            val   = fig_data[k][5]
            figtex = figtex + " & " + "\\cellcolor{"+color+"} " + str(val)
        #endfor                
        figtex = figtex + "\\\\ \n"

        figtex = figtex + "\\bottomrule\n"
        figtex = figtex + "\\end{tabularx}\n"
        figtex = figtex + "\\end{table}\n"

        tfname = os.path.join(rep_dir,"tab-sdg"+str(i)+".tex")
        if os.path.exists(tfname):
            os.remove(tfname)
        #endif
        with open(tfname,"w") as fh:
            fh.writelines(figtex)
        #endwith
        print("...done making figure\n")


        print("making the text here...")
        sectex = ""        
        for k in range(0,len(fig_data)):
            target    = fig_data[k][0]
            code_data = fig_data[k][3]
            iname  = code_data['Indicator name']
            desc   = code_data['Description']
            sdg    = code_data['SDG']
            sdga   = code_data['SDG Alignment']
            source = code_data['Source']
            rat    = code_data['Threshold Rationale']

            #make the comment file for this goal (to add comments on local performance)
            cfname = os.path.join(rep_dir,"comments")
            if os.path.exists(cfname):
                pass
            else:
                os.mkdir(cfname)
            #endif
            cfname = os.path.join(cfname,"sdg"+str(sdg)+"-"+target+".tex")
            if os.path.exists(cfname):
                pass
            else:
                os.mknod(cfname)
            #endif
            
            sectex = sectex + "\\subsection{SDG " + str(sdg) + " " + sdga + " " + iname + "}\n"
            sectex = sectex + "\\begin{labeling}{DESCRIPTION  }\n"

            if u"\u2265" in desc:
                print(desc)
                desc = desc.replace(u'\u2265',"$\geq$")  #>= symbol
                print(desc)
            #endif
            
            if desc[-1] in [".","!","?"]:
                pass
            else:
                desc = desc + "."
            #endif

            source = source.replace("&", "\&")
            
            
            sectex = sectex + "\\item [Description] " + desc + "\n"
            sectex = sectex + "\\item [Source] " + source + "\n"
            sectex = sectex + "\\item [Threshold] " + rat + "\n"
            sectex = sectex + "\\item [Comments] \\input{./comments/sdg"+str(sdg)+"-"+ target  +"}\n"
            sectex = sectex + "\\end{labeling}\n"

        #endfor
        sectex = sectex + "\\clearpage\n"

        sfname = os.path.join(rep_dir,"sec-sdg"+str(i)+".tex")
        if os.path.exists(sfname):
            os.remove(sfname)
        #endif
        with open(sfname,"w") as fh:
            fh.writelines(sectex)
        #endwith        
        print("...done making the text\n")
        
    #endfor

#enddef

#----------------------------------------------------------------
def make_summary_fig(rep_dir,city_data):
    print("TODO - make the summary table")
    print("probably call an aux routine that is used in the target_figs")

    cfmt = ">{\\columncolor{white}[0.5\\tabcolsep]}"
    
    sumtex = ""
    sumtex = sumtex + "\\begin{table}[ht]\n"
    sumtex = sumtex + "\\setlength{\\tabcolsep}{2pt}\n"
    
    sumtex = sumtex + "\\centering\n"
    sumtex = sumtex + "\\caption{XX}\n"
    sumtex = sumtex + "\\label{tab:summary}\n"
    sumtex = sumtex + "\\begin{tabular}{ "+cfmt + "\n"
    for i in range(1,18):
        sumtex = sumtex + "                c"+cfmt+"\n"
    #endfor
    sumtex = sumtex + "                c}\n"
    sumtex = sumtex + "\\toprule\n"
    sumtex = sumtex + "\\includegraphics[width=\\allgoalwidth]{../figs/AllGoals.pdf} & \n"    
    for i in range(1,17):
        sumtex = sumtex + "\\includegraphics[width=\\allgoalwidth]{../figs/Goal"+str(i)+".pdf} & \n"
    #endfor
    sumtex = sumtex + "\\includegraphics[width=\\allgoalwidth]{../figs/Goal17.pdf} \\\\ \n"

    sumtex = sumtex + "\\midrule\n"

    #the goal summary
    for i in range(0,17):
        sumtex = sumtex + "\\cellcolor{red}" + str(i) + "  & \n"
    #endfor
    sumtex = sumtex + "\\cellcolor{red}17  \\\\ \n"
    sumtex = sumtex + "\\addlinespace[3ex]\n"

    #the targets
    for i in range(0,17):
        sumtex = sumtex + "\\cellcolor{red}" + str(i) + "  & \n"
    #endfor
    sumtex = sumtex + "\\cellcolor{red}17   \\\\ \n"

    #DO MORE HERE
    
    sumtex = sumtex + "\\bottomrule\n"
    sumtex = sumtex + "\\end{tabular}\n"
    sumtex = sumtex + "\\end{table}\n"

    sfname = os.path.join(rep_dir,"summary-table.tex")
    if os.path.exists(sfname):
        os.remove(sfname)
    #endif
    with open(sfname,"w") as fh:
        fh.writelines(sumtex)
    #endwith
    print("...done making summary table")
    
#enddef

#----------------------------------------------------------------
def get_code_data(code):
    data_dir  = 'data'
    fname     = '2019USCitiesIndexResults.xlsx'
    fname     = os.path.join(os.getcwd(),os.path.join(data_dir,fname))

    sname     = 'Codebook'
    code_data = read_codebook(fname,sname,code)
    return code_data
#enddef

#----------------------------------------------------------------
def get_city_data(city):
    data_dir  = 'data'
    fname     = '2019USCitiesIndexResults.xlsx'
    fname     = os.path.join(os.getcwd(),os.path.join(data_dir,fname))

    sname     = 'Results'
    city_data = read_excel_city(fname, sname, city)
    return city_data
#enddef

#----------------------------------------------------------------
def main(city):
    #create report diretory, if needed
    dname   = city[0:3].lower()
    rep_dir = 'report-'+dname
    rep_dir = os.path.join(os.getcwd(),rep_dir)
    print(rep_dir)
    if os.path.exists(rep_dir):
        pass
    else:
        os.mkdir(rep_dir)
    #endif

    #if it doesn't exist, copy the template report
    rep_file = os.path.join(rep_dir,dname+"-sdg.tex")
    if os.path.exists(rep_file):
        pass
    else:
        rep_temp = os.path.join('report-template','template-sdg.tex')
        shutil.copyfile(rep_temp,rep_file)
    #endif

    #get the data and generate the report
    city_data = get_city_data(city)

#    make_goal_figs(rep_dir,city_data)
    #make_target_figs or words that go in each subsection
    make_summary_fig(rep_dir,city_data)
    
#enddef

if __name__ == "__main__":
    city = 'Augusta'
    main(city)
#endif
