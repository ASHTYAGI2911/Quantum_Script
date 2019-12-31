import pandas as pd
import numpy as np
import itertools
import re

path = ""

pd.set_option('display.max_colwidth', -1)
datamap = pd.read_excel(path + "Datamap.xlsx", dtype="str")

len(datamap)

output = pd.Series()
tabs = pd.Series()


########### This is the main part of the code

j = itertools.count()
t = itertools.count()
i = 0
while i<len(datamap):
    
    """ Basic set for all types of questions l card, Question text & Base text and some extra liens """
    
    dummy = datamap["START_DATA"][i]                                  # record
    questiontxt = re.split(r":|\[|\]", dummy)
    

    """ Different types of question setup  """
    i = i + 1
    
    dummy = datamap["START_DATA"][i]                                  # Columns Values
    coluns_values = re.split(r":|-|\[|\]", dummy)
    
    if dummy.startswith("-"):
        output.at[next(j)] = "l " + str.upper(questiontxt[1])
        tabs.at[next(t)] = "tab " + str.upper(questiontxt[1]) + " &ban"
        output.at[next(j)] = "ttl" + str.upper(questiontxt[1]) + "." + str.upper(questiontxt[3])
        output.at[next(j)] = "*include ibase;base=TOTAL RESPONDENTS"
        output.at[next(j)] = ""
        start = coluns_values[1]
        end = start
        i = i + 2
        for k in range(i, len(datamap)):
            if datamap["START_DATA"][k] == "NaN9999":
                break
            else:
                dummy = datamap["START_DATA"][k]
                sidestub = re.split(r"=", dummy)
                output.at[next(j)] = "n01" + str(sidestub[1]) + "               ;c=c" + str(start) + "'" + str(sidestub[0]) + "'"
        i = k    
        output.at[next(j)] = ""
        output.at[next(j)] = "*include sigma"
        output.at[next(j)] = ""
        output.at[next(j)] = "C-------------------------------------------------------------------------------------------"
    elif dummy.startswith("("):
        output.at[next(j)] = "l " + str.upper(questiontxt[1])
        tabs.at[next(t)] = "tab " + str.upper(questiontxt[1]) + " &ban"
        output.at[next(j)] = "ttl" + str.upper(questiontxt[1]) + "." + str.upper(questiontxt[3])
        output.at[next(j)] = "*include ibase;base=TOTAL RESPONDENTS"
        output.at[next(j)] = ""
        start = coluns_values[0].split("(",1)
        start = start[1]
        end = coluns_values[1].split(")",1)
        end = end[0]
        i = i + 1
        dummy1 = datamap["START_DATA"][i]
        i = i + 1
        dummy2 = datamap["START_DATA"][i]
        if dummy2.startswith('NaN9999'):
            output.at[next(j)] = "n25;inc=c(" + str(start) + "," + str(end) + ");c=c(" + str(start) + "," + str(end) + ").ge.1"
            output.at[next(j)] = "n12Mean ;dec=2;nodsp"
            output.at[next(j)] = "n01Median ;median;inc=c(" + str(start) + "," + str(end) + ");c=c(" + str(start) + "," + str(end) + ").ge.1;dec=0;notstat;nonz;ntot;nodsp"
        else:
            k = i
            for k in range(i, len(datamap)):
                if datamap["START_DATA"][k] == "NaN9999":
                    break
                else:
                    dummy = datamap["START_DATA"][k]
                    sidestub = re.split(r"=", dummy)
                    output.at[next(j)] = "n01" + str(sidestub[1]) + "               ;c=c(" + str(start) + "," + str(end) + ").in.(" + str(sidestub[0]) + ")"
            i = k
        output.at[next(j)] = ""
        output.at[next(j)] = "*include sigma"
        output.at[next(j)] = ""
        output.at[next(j)] = "C-------------------------------------------------------------------------------------------"
        output.at[next(j)] = ""
    elif dummy.startswith("Values:"):
        i = i + 1
        emptyList = {}
        counter = 0
        k = i
        dummy = datamap["START_DATA"][k]
        sidestub = re.split(r"=", dummy)
        while sidestub[0].isnumeric():
            counter = counter + 1
            emptyList[sidestub[0]] = sidestub[1]
            k = k + 1
            dummy = datamap["START_DATA"][k]
            sidestub = re.split(r"=", dummy)
        i = k
        if emptyList.get("0") == "Unchecked" and counter == 2:
            output.at[next(j)] = "l " + str.upper(questiontxt[0])
            tabs.at[next(t)] = "tab " + str.upper(questiontxt[0]) + " &ban"
            output.at[next(j)] = "ttl" + str.upper(questiontxt[0]) + "." + str.upper(questiontxt[1])
            output.at[next(j)] = "*include ibase;base=TOTAL RESPONDENTS"
            output.at[next(j)] = ""
            k = i
            for k in range(i, len(datamap)):
                if datamap["START_DATA"][k] == "NaN9999":
                    break
                else:
                    dummy = datamap["START_DATA"][k]
                    sidestub = re.split(r"\[|\]|\(|\)", dummy)
                    output.at[next(j)] = "n01" + str(sidestub[2]) + "               ;c=c" + str.upper(sidestub[3]) +  "'1'"
            i = k
            output.at[next(j)] = ""
            output.at[next(j)] = "*include sigma"
            output.at[next(j)] = ""
            output.at[next(j)] = "C-------------------------------------------------------------------------------------------"
            output.at[next(j)] = ""
        elif emptyList.get("1") == "Yes" and (counter == 2 or counter == 3):
            p = i
            k = i
            dummy = datamap["START_DATA"][k]
            sidestub = re.split(r"=", dummy)
            
            file = open(path + str.upper(questiontxt[0]), "w")
            file.write("l " + str.upper(questiontxt[0]) + "_&N ;c=&filt\n")
            file.write("ttl" + str.upper(questiontxt[0]) + "_&N." + str.upper(questiontxt[1]) + " - &TXT\n")
            file.write("*include ibase;base=TOTAL RESPONDENTS\n")
            file.write("\n")
            dummy = datamap["START_DATA"][k]
            sidestub = re.split(r"\[|\]|\(|\)", dummy)
            
            for key, value in emptyList.items():
                file.write("n01" + str(value) + "               ;c=ca00'" + str(key) + "'\n")
                    
            if counter >= 2 and counter <= 3:
                file.write("\n")
                file.write("n01NET: YES    ;c=ca00'1';ntot\n")
            file.write("\n")
            file.write("*include sigma\n")
            file.write("\n")
            file.write("C-------------------------------------------------------------------------------------------\n")
            file.write("\n")
            file.close()
            m=1
            for k in range(i, len(datamap)):
                if datamap["START_DATA"][k] == "NaN9999":
                    break
                else:
                    dummy = datamap["START_DATA"][k]
                    sidestub = re.split(r"\[|\]|\(|\)", dummy)
                    output.at[next(j)] = "*include " + str.upper(questiontxt[0]) + ";col(a)=" + str.upper(sidestub[3]) + ";filt=1.eq.1;N=" + str(m) + ";TXT=" + str.upper(sidestub[2])
                    tabs.at[next(t)] = "tab " + str.upper(questiontxt[0]) + "_" + str(m) + " &ban"
                m+=1
            output.at[next(j)] = ""
            output.at[next(j)] = "*include " + str.upper(questiontxt[0]) + "_YesSum;qid=" + str.upper(questiontxt[0]) + "_YesSum;TXT=YES SUMMARY;code=1"
            output.at[next(j)] = ""
            output.at[next(j)] = "C-------------------------------------------------------------------------------------------"
            output.at[next(j)] = ""
            
            """ Creating a summary table """
            tabs.at[next(t)] = "tab " + str.upper(questiontxt[0]) + "_YesSum" + " &ban"
            file = open(path + str.upper(questiontxt[0]) + "_YesSum", "w")
            file.write("l " + "&qid\n")
            file.write("ttl" + str.upper(questiontxt[0]) + "." + str.upper(questiontxt[1]) + " - &TXT\n")
            file.write("*include ibase;base=TOTAL RESPONDENTS\n")
            file.write("\n")
            for s in range(p, len(datamap)):
                if datamap["START_DATA"][s] == "NaN9999":
                    break
                else:
                    dummy = datamap["START_DATA"][s]
                    sidestub = re.split(r"\[|\]|\(|\)", dummy)
                    file.write("*include fbox;cols=c" + str(sidestub[3]) + ";flt=1.eq.1;stmt=" + str(sidestub[2]) + "\n")

            file.write("\n")
            file.write("C-------------------------------------------------------------------------------------------")
            file.write("\n")
            file.close()
            
            i = k
        elif (counter >= 4 and counter <= 7):
            p = i
            k = i
            dummy = datamap["START_DATA"][k]
            sidestub = re.split(r"=", dummy)
            
            file = open(path + str.upper(questiontxt[0]), "w")
            file.write("l " + str.upper(questiontxt[0]) + "_&N ;c=&filt\n")
            file.write("ttl" + str.upper(questiontxt[0]) + "_&N." + str.upper(questiontxt[1]) + " - &TXT\n")
            file.write("*include ibase;base=TOTAL RESPONDENTS\n")
            file.write("\n")
            dummy = datamap["START_DATA"][k]
            sidestub = re.split(r"\[|\]|\(|\)", dummy)
            
            for key, value in emptyList.items():
                file.write("n01" + str(value) + " ()               ;c=ca00'" + str(key) + "';fac=" + str(key) + "\n")
                    
            if counter == 5 or counter == 6 or counter == 7:
                file.write("\n")
                file.write("n01NET: TOP 2 BOX (4-5)       ;c=ca00'12';ntot;nofac\n")
                file.write("n01NET: BOTTOM 2 BOX (1-2)    ;c=ca00'45';ntot;nofac\n")
                file.write("\n")
                file.write("n12Mean                       ;dec=1;nodsp\n")
                file.write("n17Standard Deviation         ;dec=2;nodsp\n")
                file.write("n19Standard Error             ;dec=2;nodsp\n")
            file.write("\n")
            file.write("*include sigma\n")
            file.write("\n")
            file.write("C-------------------------------------------------------------------------------------------\n")
            file.write("\n")
            file.close()
            m=1
            for k in range(i, len(datamap)):
                if datamap["START_DATA"][k] == "NaN9999":
                    break
                else:
                    dummy = datamap["START_DATA"][k]
                    sidestub = re.split(r"\[|\]|\(|\)", dummy)
                    output.at[next(j)] = "*include " + str.upper(questiontxt[0]) + ";col(a)=" + str.upper(sidestub[3]) + ";filt=1.eq.1;N=" + str(m) + ";TXT=" + str.upper(sidestub[2])
                    tabs.at[next(t)] = "tab " + str.upper(questiontxt[0]) + "_" + str(m) + " &ban"
                m+=1
            output.at[next(j)] = ""
            output.at[next(j)] = "*include " + str.upper(questiontxt[0]) + "_Sum;qid=" + str.upper(questiontxt[0]) + "_SumT2;TXT=TOP 2 BOX SUMMARY;code=12"
            output.at[next(j)] = ""
            output.at[next(j)] = "C-------------------------------------------------------------------------------------------"
            output.at[next(j)] = ""
            
            """ Creating a summary table """
            tabs.at[next(t)] = "tab " + str.upper(questiontxt[0]) + "_SumT2" + " &ban"
            file = open(path + str.upper(questiontxt[0]) + "_Sum", "w")
            file.write("l " + "&qid\n")
            file.write("ttl" + str.upper(questiontxt[0]) + "." + str.upper(questiontxt[1]) + " - &TXT\n")
            file.write("*include ibase;base=TOTAL RESPONDENTS\n")
            file.write("\n")
            for s in range(p, len(datamap)):
                if datamap["START_DATA"][s] == "NaN9999":
                    break
                else:
                    dummy = datamap["START_DATA"][s]
                    sidestub = re.split(r"\[|\]|\(|\)", dummy)
                    file.write("*include fbox;cols=c" + str(sidestub[3]) + ";flt=1.eq.1;stmt=" + str(sidestub[2]) + "\n")

            file.write("\n")
            file.write("C-------------------------------------------------------------------------------------------")
            file.write("\n")
            file.close()
            i = k
        elif (counter >= 8 and counter <= 11):
            p = i
            k = i
            dummy = datamap["START_DATA"][k]
            sidestub = re.split(r"=", dummy)
            
            file = open(path + str.upper(questiontxt[0]), "w")
            file.write("l " + str.upper(questiontxt[0]) + "_&N ;c=&filt\n")
            file.write("ttl" + str.upper(questiontxt[0]) + "_&N." + str.upper(questiontxt[1]) + " - &TXT\n")
            file.write("*include ibase;base=TOTAL RESPONDENTS\n")
            file.write("\n")
            dummy = datamap["START_DATA"][k]
            sidestub = re.split(r"\[|\]|\(|\)", dummy)
            
            for key, value in emptyList.items():
                file.write("n01" + str(value) + " ()               ;c=c(a00,a01).in.(" + str(key) + ");fac=" + str(key) + "\n")
                    
            if counter >= 10:
                file.write("\n")
                file.write("n01NET: TOP 3 BOX (8-10)       ;c=c(a00,a01).in.(8,9,10);ntot;nofac\n")
                file.write("n01NET: TOP 2 BOX (9-10)       ;c=c(a00,a01).in.(9,10);ntot;nofac\n")
                file.write("n01NET: BOTTOM 2 BOX (1-2)    ;c=c(a00,a01).in.(1,2);ntot;nofac\n")
                file.write("n01NET: BOTTOM 3 BOX (1-3)    ;c=c(a00,a01).in.(1,2,3);ntot;nofac\n")
                file.write("\n")
                file.write("n12Mean                       ;dec=1;nodsp\n")
                file.write("n17Standard Deviation         ;dec=2;nodsp\n")
                file.write("n19Standard Error             ;dec=2;nodsp\n")
            file.write("\n")
            file.write("*include sigma\n")
            file.write("\n")
            file.write("C-------------------------------------------------------------------------------------------\n")
            file.write("\n")
            file.close()
            m=1
            for k in range(i, len(datamap)):
                if datamap["START_DATA"][k] == "NaN9999":
                    break
                else:
                    dummy = datamap["START_DATA"][k]
                    sidestub = re.split(r"\[|\]|\(|\)", dummy)
                    output.at[next(j)] = "*include " + str.upper(questiontxt[0]) + ";col(a)=" + str.upper(sidestub[3]) + ";filt=1.eq.1;N=" + str(m) + ";TXT=" + str.upper(sidestub[2])
                    tabs.at[next(t)] = "tab " + str.upper(questiontxt[0]) + "_" + str(m) + " &ban"
                m+=1
            output.at[next(j)] = ""
            output.at[next(j)] = "*include " + str.upper(questiontxt[0]) + "_Sum;qid=" + str.upper(questiontxt[0]) + "_SumT2;TXT=TOP 2 BOX SUMMARY;code=12"
            output.at[next(j)] = ""
            output.at[next(j)] = "C-------------------------------------------------------------------------------------------"
            output.at[next(j)] = ""
            
            """ Creating a summary table """
            tabs.at[next(t)] = "tab " + str.upper(questiontxt[0]) + "_SumT2" + " &ban"
            file = open(path + str.upper(questiontxt[0]) + "_Sum", "w")
            file.write("l " + "&qid\n")
            file.write("ttl" + str.upper(questiontxt[0]) + "." + str.upper(questiontxt[1]) + " - &TXT\n")
            file.write("*include ibase;base=TOTAL RESPONDENTS\n")
            file.write("\n")
            for s in range(p, len(datamap)):
                if datamap["START_DATA"][s] == "NaN9999":
                    break
                else:
                    dummy = datamap["START_DATA"][s]
                    sidestub = re.split(r"\[|\]|\(|\)", dummy)
                    file.write("*include Ffav;cols=c" + str(sidestub[3]) + ";flt=1.eq.1;stmt=" + str(sidestub[2]) + "\n")

            file.write("\n")
            file.write("C-------------------------------------------------------------------------------------------")
            file.write("\n")
            file.close()
            i = k
        elif counter > 11:
            k = i
            dummy = datamap["START_DATA"][k]
            sidestub = re.split(r"=", dummy)
            
            file = open(path + str.upper(questiontxt[0]), "w")
            file.write("l " + str.upper(questiontxt[0]) + "_&N ;c=&filt\n")
            file.write("ttl" + str.upper(questiontxt[0]) + "_&N." + str.upper(questiontxt[1]) + " - &TXT\n")
            file.write("*include ibase;base=TOTAL RESPONDENTS\n")
            file.write("\n")
            dummy = datamap["START_DATA"][k]
            sidestub = re.split(r"\[|\]|\(|\)", dummy)
            
            for key, value in emptyList.items():
                if counter >= 5 and counter <= 9:
                    file.write("n01" + str(value) + "               ;c=ca00'" + str(key) + "';fac=" + str(key) + "\n")
                elif counter >= 0 and counter <= 4:
                    file.write("n01" + str(value) + "               ;c=ca00'" + str(key) + "'\n")
                elif counter >= 10:
                    file.write("n01" + str(value) + "               ;c=c(a00,a01).in.(" + str(key) + ");fac=" + str(key) + "\n")
                    
            if counter == 5 or counter == 6 or counter == 7:
                file.write("\n")
                file.write("n01NET: TOP 2 BOX (4-5)       ;c=ca00'45';ntot;nofac\n")
                file.write("n01NET: BOTTOM 2 BOX (1-2)    ;c=ca00'12';ntot;nofac\n")
                file.write("\n")
                file.write("n12Mean                       ;dec=1;nodsp\n")
                file.write("n17Standard Deviation         ;dec=2;nodsp\n")
                file.write("n19Standard Error             ;dec=2;nodsp\n")
            elif counter >= 10:
                file.write("\n")
                file.write("n01NET: TOP 3 BOX (8-10)       ;c=c(a00,a01).in.(8,9,10);ntot;nofac\n")
                file.write("n01NET: TOP 2 BOX (9-10)       ;c=c(a00,a01).in.(9,10);ntot;nofac\n")
                file.write("n01NET: BOTTOM 2 BOX (1-2)    ;c=c(a00,a01).in.(1,2);ntot;nofac\n")
                file.write("n01NET: BOTTOM 3 BOX (1-3)    ;c=c(a00,a01).in.(1,2,3);ntot;nofac\n")
                file.write("\n")
                file.write("n12Mean                       ;dec=1;nodsp\n")
                file.write("n17Standard Deviation         ;dec=2;nodsp\n")
                file.write("n19Standard Error             ;dec=2;nodsp\n")
            file.write("\n")
            file.write("*include sigma\n")
            file.write("\n")
            file.write("C-------------------------------------------------------------------------------------------\n")
            file.write("\n")
            file.close()
            m=1
            for k in range(i, len(datamap)):
                if datamap["START_DATA"][k] == "NaN9999":
                    break
                else:
                    dummy = datamap["START_DATA"][k]
                    sidestub = re.split(r"\[|\]|\(|\)", dummy)
                    output.at[next(j)] = "*include " + str.upper(questiontxt[0]) + ";col(a)=" + str.upper(sidestub[3]) + ";filt=1.eq.1;N=" + str(m) + ";TXT=" + str.upper(sidestub[2])
                    tabs.at[next(t)] = "tab " + str.upper(questiontxt[0]) + "_" + str(m) + " &ban"
                m+=1
            output.at[next(j)] = ""
            output.at[next(j)] = "C-------------------------------------------------------------------------------------------"
            output.at[next(j)] = ""
            i = k
       
    i = i + 1


output.to_excel(path + "axes.xlsx")
tabs.to_excel(path + "tabs.xlsx")
