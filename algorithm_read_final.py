# -*- coding: utf-8 -*-
"""
Created on Thu Feb  1 13:18:42 2024

@author: ffs
"""

import os
import math
import numpy as np
import sys

#open PowerFActory
sys.path.append(r'C:\Program Files\DIgSILENT\PowerFactory 2023 SP6\Python\3.9') #correct accordingly 
import powerfactory as pf
app = pf.GetApplication()
if app is None:
    raise Exception('getting Powerfactory application failed')
app.Show()


#define project name and study case    
projName = 'XXXX' #correct accordingly
#study_case = '01 - Load Flow'

#activate project
project = app.ActivateProject(projName)
proj = app.GetActiveProject()

#define limits
V_Max=1.05
V_Min=0.95
SOEC_maximum=40
SOEC_minimum=0.5
Loading_max=100

#execute multiple power flows
time_lf_ref=int(000000); #initial time for the 1st loadflow
day_counter=1 #start at teh 1st day of the year
t_lf=60; #minutes between runs of the load flow 
time_step=0;
N_terminals=203; #number of terminlas
N_busbar=29
N_lines=15; #number of lines


SOEC_optimisation_file = open(r"XXX.txt",'r') #read the value obtained from the optimisation
SOEC_optimisation_data=SOEC_optimisation_file.readlines() #input all the data into a list
SOEC_load = app.GetCalcRelevantObjects('SOEC.Elmlod')[0]

for time_step in range(8736): #the range is the number of hours that the simulation should run
    SOEC_load.plini_Watts = float(SOEC_optimisation_data[time_step]) #reads the number from the file, however, the number is in Watts
    SOEC_load.plini = SOEC_load.plini_Watts/1000#writes the load in MW at the general load equivalent ot he SOEC device. The value is inputted from Matlab
    print(SOEC_load.plini)
            
    #change the load for the load flow
    Loads=app.GetCalcRelevantObjects('*.ElmLod')
    name_load = Loads[0].GetAttribute('loc_name') # get name of the load
       
    #method to change weeks, days or hours
    multi_time=app.GetCalcRelevantObjects('*.SetTime')
 
    #to change the day of the week, and hour but not minutes or seconds
    oSetTime=app.GetFromStudyCase('SetTime') # get object that allows to set time
    multi_time[0].dayofyear=day_counter #select the day
    time_lf_temp=(time_lf_ref+time_step*t_lf-24*60*(day_counter-1)) # calculate the the time, here it is assumed periods of t_lf minutes between load flow runs
    hours=math.floor(time_lf_temp/60); #calculate the hours by rounding down
    minutes=time_lf_temp-hours*60; # calculate the minutes
    if len(str(hours)) == 1: # check the length of the hours, which must be 2: hhmmss
        hours_string = '0'+str(hours)
    else: hours_string = str(hours)
    
    if len(str(minutes)) == 1: # check the length of the minutes, which must be 2: hhmmss
        minutes_string = '0'+str(minutes)
    else: minutes_string = str(minutes)
    

    time_lf=hours_string+minutes_string+'00';
    oSetTime.cTime=time_lf # set the time as a string in the following format 'hhmmss'
    
    status_volt = 1
    while status_volt==1:
        #get load flow object and execute
        oLoadflow=app.GetFromStudyCase('ComLdf') #get load flow object
        oLoadflow.Execute() #execute load flow
    
    #the loads and the generators are not used for anything right now. It is just in case of being necessary
             
        #get the voltage 
        Busbar_posicao=0 #help counter to deal with the fact that there are more terminals than busbars
        Volt = app.GetCalcRelevantObjects('*.ElmTerm')
        txt_file = open("voltage_busbars.txt","a")         #create txt file and write the information in the file
        txt_file.write("\nday"), txt_file.write('{}'.format(day_counter)),txt_file.write(" - hour: "),txt_file.write('{}'.format(hours)), txt_file.write(":"), txt_file.write('{}'.format(minutes_string)), txt_file.write("\n") #write headhing for each time
        Voltages_pu=np.zeros(N_busbar)
        for y in range(N_terminals): #Read the busbar voltage from PF for each bubar
            if Volt[y].GetAttribute("e:iUsage") == 0:     #not all terminals are busbars. So, this checks which ones are busbars and should be recorded
                Voltages_pu[Busbar_posicao]=Volt[y].GetAttribute("m:u")  #appends the value at the end of the list    
                txt_file.write("V"), txt_file.write('{}'.format(y+1)), txt_file.write(" = %.4f \n" %Voltages_pu[Busbar_posicao]) #write voltage 
                Busbar_posicao=Busbar_posicao+1
        txt_file.close()
        y=0
        print(Voltages_pu) #not necessary, it is just to show it in Python window  
        
        Current = app.GetCalcRelevantObjects('*.ElmLne')
        txt_file = open("Current_lines.txt","a")         #create txt file and write the information in the file
        txt_file.write("\nday"), txt_file.write('{}'.format(day_counter)),txt_file.write(" - hour: "),txt_file.write('{}'.format(hours)), txt_file.write(":"), txt_file.write('{}'.format(minutes_string)), txt_file.write("\n") #write headhing for each time
        Line_loading=np.zeros(N_lines)
        for y in range(N_lines): #Read the busbar voltage from PF for each bubar
            Line_loading[y]=Current[y].GetAttribute("c:loading")  #appends the value at the end of the list    
            txt_file.write("Loading"), txt_file.write('{}'.format(y+1)), txt_file.write(" = %.4f \n" %Line_loading[y]) #write voltage 
        txt_file.close()
        y=0
        print(Line_loading) #not necessary, it is just to show it in Python window    
        
        #check if htere is overcurrent or over/under votlage
        if any(Voltages_pu>=V_Min) and any(Voltages_pu)<=V_Max:
            status_volt=0
            if any(Line_loading>Loading_max):
                print ('warning: there is an overcurrent')
        if any(Voltages_pu>V_Max):
            print('voltage too high') #inform that the voltage is too high. SOEC should increase
            if SOEC_load.plini<SOEC_maximum:
                SOEC_load.plini=SOEC_load.plini*1.1 #soec increases 10%
                if SOEC_load.plini<SOEC_minimum:
                    SOEC_load.plini=SOEC_minimum
                if SOEC_load.plini>SOEC_maximum:
                    SOEC_load.plini=SOEC_maximum #if it goes over the maxium with teh 10%, make it equal to the maximum
        if any(Voltages_pu<V_Min):
            print('voltage too low') #inform that the voltage is too low. SOEC should decrease
            if SOEC_load.plini>SOEC_minimum:
                SOEC_load.plini=SOEC_load.plini*0.9 #soec decreases 10%
                if SOEC_load.plini<SOEC_minimum:
                    SOEC_load.plini=SOEC_minimum
                if SOEC_load.plini>SOEC_maximum:
                    SOEC_load.plini=SOEC_maximum 
                    
    if hours == 23: # it needs to start a new day and to reset the hour counter
        day_counter = day_counter+1        
                      