import iesve        # the VE api
import pprint
from ies_file_picker import IesFilePicker
import time
import tkinter as tk
import tkinter.messagebox as messagebox
import xlsxwriter
from xlsxwriter.utility import xl_rowcol_to_cell
from xlsxwriter.utility import xl_range
import os
import datetime
import time
from pprint import pprint as pp

os.getlogin()
current_date = datetime.date.today()                      
user = os.path.join('..','Documents and Settings',os.getlogin(),'Desktop')

#### CREATE POP UPS FOR SELECTION OF ROOMS AND RESULTS FILES #########
def generate_window(project, ve_folder, results_reader, room_groups):
    
    
    # works out user name
    os.getlogin()
    current_date = datetime.date.today()                      
    userpath = os.path.join('..','Documents and Settings',os.getlogin(),'Desktop')
    
    user_name = []
    backwards_user_name = []
    
    for n in range(9,30):
        if userpath[-n] == "\\":
            break
        else:
            backwards_user_name = str(backwards_user_name) + userpath[-n]
    user_name_length = len(backwards_user_name)
 
    for n in range (0,user_name_length):
        user_name = str(user_name) + backwards_user_name[-n]
        user_name_new = str(user_name[3:-1] ) 
        user_name_new2 = user_name_new.replace("."," ")
    
    # works out model name from file path
    backwards_model_name = []
    model_name = []
    for n in range(2,30):
        if project.path[-n] == "\\":
            break
        else:
            backwards_model_name = str(backwards_model_name) + project.path[-n]
    model_name_length = len(backwards_model_name)    
 
    for n in range (0,model_name_length):
        model_name = str(model_name) + backwards_model_name[-n]
    model_name_new = model_name[3:-1]
    
    class Window(tk.Frame):
        def __init__(self, master=None):
            tk.Frame.__init__(self, master)
            self.project = project
            self.project_folder = project.path
            self.ve_folder = ve_folder
            self.save_file_name = (current_date,"HGHL")
            self.results_reader = results_reader
            self.room_groups = room_groups
            self.master = master
            self.init_window()

        # Creation of window
        def init_window(self):
            self.master.title("HGHL results")
            self.master.columnconfigure(0, weight=1)
            self.master.rowconfigure(0, weight=1)
            self.master.grid()

            instructions_label = tk.Label(self, text='Click the button below to create a new grouping scheme in your model.\nManually add the rooms you wish to be analysed to the group "Analyse HGHL Results" \n(make sure not to include shading or adjacent buildings).\n (Note: Check the assumptions on the front sheet of the results for correctness)')
            instructions_label.grid(row=0, sticky=tk.W)
            tk.Label(self, text='').grid(row=1, sticky=tk.W)
            tk.Button(self, text="Create Grouping Scheme", command=self.create_grouping).grid(row=2, sticky=tk.W)
            tk.Label(self, text=' ').grid(row=3, sticky=tk.W)

            tk.Label(self, text=' ').grid(row=4, sticky=tk.W)
            tk.Label(self, text='Select Vista results files for both .htg and .clg').grid(row=5, sticky=tk.W)

            print(self.project_folder)
            path = self.project_folder + 'Vista'        
            files = os.listdir(path)
            htg_aps_files = []
            clg_aps_files = []

            for htg_file in files:
                htg_file = htg_file.split('.')
                if htg_file[-1] == 'htg':
                    htg_file = ('.').join(htg_file)
                    htg_aps_files.append(htg_file)        
            
            for clg_file in files:
                clg_file = clg_file.split('.')
                if clg_file[-1] == 'clg':
                    clg_file = ('.').join(clg_file)
                    clg_aps_files.append(clg_file)          

            #sets listbox, allows selection of multiple files
            self.listbox = tk.Listbox(self, selectmode = tk.MULTIPLE)

            for htg_file in htg_aps_files:
                self.listbox.insert(tk.END, htg_file)
                
            for clg_file in clg_aps_files:
                self.listbox.insert(tk.END, clg_file)            
            
            self.listbox.select_set(0)
            self.listbox.grid(row=6, sticky='nsew')
            tk.Label(self, text=' ').grid(row=7, sticky=tk.W)

            tk.Label(self, text='HGHL results will be added to an Excel sheet that will be saved in the model folder').grid(row=8, sticky=tk.W)
            tk.Label(self, text='Name the Excel file below:').grid(row=9, sticky=tk.W)
            
            self.save_file_entry_box = tk.Entry(self)
            self.save_file_entry_box.insert(0, self.save_file_name)
            self.save_file_entry_box.grid(row=10, sticky='ew')
            tk.Label(self, text=' ').grid(row=11, sticky=tk.W)

            # creating a button instance
            tk.Button(self, text="Run Calculation", command=self.run_calc).grid(row=12, sticky=tk.W)
            
            self.columnconfigure(0, weight=1)
            self.rowconfigure(6, weight=1)
            self.grid(row=0, column=0, sticky='nsew')
            
        def create_grouping(self):
            """function activated by the 'Create Grouping Scheme' button. Tests to see if the grouping scheme already
            exists. If it does not exist, it creates it. If it does exist a popup message is displayed to the user"""
            schemes = self.room_groups.get_grouping_schemes()
            new_grouping_scheme_needed = True
            for scheme in schemes:
                if scheme['name'] == 'HGHL Analysis':
                    new_grouping_scheme_needed = False

            if new_grouping_scheme_needed:
                scheme_index = self.room_groups.create_grouping_scheme('HGHL Analysis')
                room_groups.create_room_group(scheme_index, "Analyse HGHL Results")
                room_groups.create_room_group(scheme_index, "Do Not Analyse")

                tk.messagebox.showinfo("Grouping scheme", "Manually select the rooms you wish to be analysed and add them to the group \'Analyse HGHL\'")
            else:
                tk.messagebox.showinfo("Grouping scheme already exists", "Grouping scheme already exists")


        def run_calc(self):
            """create the excel workbook, runs the mains calculation functions and writes data to the excel sheet"""

            # Gets a list of the rooms to be analysed from the grouping scheme
            schemes = self.room_groups.get_grouping_schemes()
            scheme_handle = False
            for scheme in schemes:
                if scheme['name'] == 'HGHL Analysis':
                    scheme_handle = scheme['handle']

            if scheme_handle == False:
                tk.messagebox.showinfo("Grouping scheme", "Create grouping scheme and assign rooms before running calculation")
                return

            hghl_group = self.room_groups.get_room_groups(scheme_handle)
            rooms_to_be_analysed = []

            for group in hghl_group:
                if group['name'] == 'Analyse HGHL Results':
                    rooms_to_be_analysed = group['rooms']

            # if the grouping scheme is empty, a popup is displayed to the user telling them to put some rooms into the group
            if not rooms_to_be_analysed:
                tk.messagebox.showerror("Room group error", "You must manually add some rooms to the room group \'Analyse HL Results\'")
                return
                
            selected = self.listbox.curselection()            

            #takes the first selected name, is this always the heating file?
            htg_file_name = self.listbox.get(selected[0])
            print('htg file selected:',htg_file_name)
            if not htg_file_name:
                tk.messagebox.showinfo(".htg error", "No .htg file selected. Please select a .htg file.")
                return
                
            # takes the second active name, is this always the cooling file?                
            clg_file_name = self.listbox.get(selected[1])
            print('clg file selected:', clg_file_name)
            if not clg_file_name:
                tk.messagebox.showinfo(".clg error", "No .clg file selected. Please select a .clg file.")
                return
            
            self.save_file_name = self.save_file_entry_box.get()
            print('Excel File name = \t\t' + self.save_file_name)
            
            # create excel workbook
            workbook = xlsxwriter.Workbook(self.project_folder + '\\' + self.save_file_name + '.xlsx')
           
            # create excel work sheet  
            sheet1 = workbook.add_worksheet('Notes')           
            sheet2 = workbook.add_worksheet('Heating loads')
            sheet3 = workbook.add_worksheet('Cooling loads')   
            
            # insert images to worksheet
            sheet1.insert_image('A2', 'aecom.png', {'x_scale': 0.3, 'y_scale': 0.3})
            sheet2.insert_image('A2', 'aecom.png', {'x_scale': 0.3, 'y_scale': 0.3})
            sheet3.insert_image('A2', 'aecom.png', {'x_scale': 0.3, 'y_scale': 0.3})      
            
            # set Notes column widths
            sheet1.set_column('B:C', 13)            
            
            # set heating loads column widths
            sheet2.set_column('A:A', 30)
            sheet2.set_column('B:D', 15)
            sheet2.set_column('E:I', 19)  
            sheet2.set_column('J:O', 17) 
            
            # set cooling loads column widths
            sheet3.set_column('A:A', 30)
            sheet3.set_column('B:B', 15)
            sheet3.set_column('C:D', 13)
            sheet3.set_column('E:J', 15)
            sheet3.set_column('K:L', 17)  
            sheet3.set_column('N:O', 17)  
            
             # Add a bold format to use to highlight cells.
            titleformat = workbook.add_format({'bold': True}) 
            
            sidecolumnformat = workbook.add_format({'bold': False, 'bottom': False , 'top': False, 'right': True, 'left':True, 'text_wrap':False})
            sidecolumnformat.set_border(style=1)  
            
            #length of list for number of rooms
            list_rooms_length = len(rooms_to_be_analysed)
            list_rooms_length = list_rooms_length + 6
            string_length = str(list_rooms_length)
            
            # Add Border to sheet - not sure if this does anything
            border_format = workbook.add_format({'bold': False, 'bottom': False, 'top': False, 'right': True, 'left':True, 'text_wrap':False})
            border_format.set_right(2) 
            border_format.set_bottom(1)
            sheet2.conditional_format( 'I7:I' + string_length , { 'type': 'no_blanks' , 'format' : border_format })
            sheet3.conditional_format( 'L7:L' + string_length , { 'type': 'no_blanks' , 'format' : border_format })
            
            #formating for the columns
            dnt_format = workbook.add_format({'bold': False, 'bottom': False, 'top': False, 'right': False, 'left':False, 'text_wrap':False})
            dnt_format.set_right(2)
            dnt_format.set_left(2)
            dnt_format.set_bottom(1)
            sheet2.conditional_format( 'A7:I' + string_length, { 'type': 'no_blanks' , 'format' : dnt_format })
            sheet3.conditional_format( 'A7:L' + string_length, { 'type': 'no_blanks' , 'format' : dnt_format })
            
            headingformat = workbook.add_format({'bold': True, 'bottom': True, 'top': True, 'right': True,'text_wrap':True})
            headingformat.set_align('vcenter')
            headingformat.set_border(style=2)
            
            nonboldheadingformat = workbook.add_format({'bold': False, 'bottom': True, 'top': True, 'right': True,'text_wrap':True})
            
            sheet2.conditional_format('A6:L6',{'type':'no_blanks','format':headingformat})
            sheet3.conditional_format('A6:L6',{'type':'no_blanks','format':headingformat})          
                
                    
                        
            # Freeze heading panes
            sheet2.freeze_panes(6, 4)
            sheet3.freeze_panes(6, 4)
            
            #Write text to sheets
            sheet1.write('B5','Note:', titleformat)
            sheet1.write('B6','- No fresh air loads allowed for in calculation ')
            sheet1.write('B7','- No cold start allowed for in the heating loads')
            sheet1.write('B8','- No safety margin added to results')
            sheet1.write('B9','- No systems delivery or efficiency losses included in the results')
            sheet1.write('B10','- No solar shading from surrounding buildings in heat gain calculations')
            
            sheet1.write('B13','QA', titleformat)
            sheet1.write('B14','Model name:')
            sheet1.write('C14', model_name_new)
            sheet1.write('B15','Created by:')
            sheet1.write('C15',user_name_new2)
            sheet1.write('B16','Checked by:')
            sheet1.write('B20', 'HGHL results script update version: 20191111')
            sheet2.write('B2', 'Heating loads', titleformat)
            sheet2.write('B3','Results file name:') 
            sheet2.write('B4','Model name:') 
            sheet3.write('B2', 'Cooling loads', titleformat)
            sheet3.write('B3','Results file name:') 
            sheet3.write('B4','Model name:')

            
####################### HEAT LOSS ####################
             
            # run main calculation functions for heating
            results_reader_file = results_reader.open(htg_file_name)
            results_file_name = htg_file_name
            numberofrooms = len(rooms_to_be_analysed)            
          
            def get_hl_data(results_reader_file, rooms_to_be_analysed, htg_file):
                hl_data = []
                all_steady_state_heating_plant_load_kw = []
                rooms = results_reader_file.get_room_list()                              
                
                
                # loops through every room in the model
                for room_number, room in enumerate(rooms):
                       
                    # unpack the tuples of room data that is returned from the ResultsReader function. 'a' and
                    #  'b' and not used but had to be assigned to something
                    name, room_id, a, b = room
                    name, volume, room_area, room_volume = room
                    rounded_room_area=(round(room_area ,2))
                                                                            
                    # prints progress information to the console for the user
                    if room_number % 10 == 0 and room_number + 10 < len(rooms):
                        print('Analysing rooms: ' + str(room_number + 1) + ' - ' + str(room_number + 10) + ' of ' + str(len(rooms)) + ' in ' + htg_file)
                    elif room_number % 10 == 0 and room_number + 10 >= len(rooms):
                        print('Analysing rooms: ' + str(room_number + 1) + ' - ' + str(len(rooms)) + ' of ' + str(len(rooms)) + ' in ' + htg_file)
                    
                                            
                    # checks if the current room is part of the 'rooms_to_be_analysed' list. Only performs calculations on the rooms that the user placed in the room group
                    if room_id in rooms_to_be_analysed:
                        np_air_temp = results_reader_file.get_room_results(room_id, 'Room air temperature','Air temperature', 'z') 
                        air_temp = float(np_air_temp)
                        air_temp = round(air_temp,2)
                        
                        np_dry_resultant_temp = results_reader_file.get_room_results(room_id, 'Comfort temperature', 'Dry resultant temperature','z')
                        dry_resultant_temp = float(np_dry_resultant_temp)
                        dry_resultant_temp = round(dry_resultant_temp,2)
                        
                        np_external_conduction_gain = results_reader_file.get_room_results(room_id, 'Conduction from ext elements', 'External conduction gain','z')
                        np_external_conduction_gain_kw = np_external_conduction_gain / 1000
                        external_conduction_gain_kw = float(np_external_conduction_gain_kw)
                        external_conduction_gain_kw = round(external_conduction_gain_kw,2)
                        
                        np_internal_conduction_gain = results_reader_file.get_room_results(room_id, 'Conduction from int surfaces','Internal conduction gain', 'z')
                        np_internal_conduction_gain_kw = np_internal_conduction_gain / 1000
                        internal_conduction_gain = float(np_internal_conduction_gain_kw)
                        internal_conduction_gain_kw = round(internal_conduction_gain,2)
                        
                        np_infiltration_gain = results_reader_file.get_room_results(room_id, 'Infiltration gain','Infiltration gain', 'z')
                        np_infiltration_gain_kw = np_infiltration_gain / 1000
                        infiltration_gain = float(np_infiltration_gain_kw)
                        infiltration_gain_kw = round(infiltration_gain,2)
                        
                        np_steady_state_heating_plant_load = results_reader_file.get_room_results(room_id, 'Room units steady state htg load','Steady state heating plant load', 'z')
                        np_steady_state_heating_plant_load_kw = np_steady_state_heating_plant_load / 1000
                        steady_state_heating_plant_load = float(np_steady_state_heating_plant_load_kw)
                        steady_state_heating_plant_load_kw = round(steady_state_heating_plant_load,2)
                        steady_state_heating_plant_load_kw_for_total= round(steady_state_heating_plant_load,4)
                        
                        all_steady_state_heating_plant_load_kw.append(steady_state_heating_plant_load_kw_for_total) 
                        totalsshpl = round(sum(all_steady_state_heating_plant_load_kw),2)
                                              
                        room_hl_data = [name, rounded_room_area, air_temp, dry_resultant_temp, external_conduction_gain_kw, internal_conduction_gain_kw, infiltration_gain_kw, steady_state_heating_plant_load_kw,totalsshpl]
                        hl_data.append(room_hl_data)  
                
                return hl_data                
            hl_data = get_hl_data(results_reader_file, rooms_to_be_analysed, htg_file_name)
            
                             
          
 ###################### COOLING ######################################
            # run main calculation functions
            results_reader_file = results_reader.open(clg_file_name)
            results_file_name = clg_file_name
            numberofrooms = len(rooms_to_be_analysed)
            
            def get_hg_data(results_reader_file, rooms_to_be_analysed, clg_file):
                hg_data = []
                rooms = results_reader_file.get_room_list()
                
                total_combine_space_con = [] 
                
                # loops through every room in the model
                for room_number, room in enumerate(rooms):
                    
                    # unpack the tuples of room data that is returned from the ResultsReader function. 'a' and
                    #  'b' and not used but had to be assigned to something
                    name, room_id, a, b = room
                    name, volume, room_area, room_volume = room
                    rounded_room_area=(round(room_area ,2))
                                        
                    # prints progress information to the console for the user
                    if room_number % 10 == 0 and room_number + 10 < len(rooms):
                        print('Analysing rooms: ' + str(room_number + 1) + ' - ' + str(room_number + 10) + ' of ' + str(len(rooms)) + ' in ' + clg_file)
                    elif room_number % 10 == 0 and room_number + 10 >= len(rooms):
                        print('Analysing rooms: ' + str(room_number + 1) + ' - ' + str(len(rooms)) + ' of ' + str(len(rooms)) + ' in ' + clg_file)
                        
                       
                    # checks if the current room is part of the 'rooms_to_be_analysed' list. Only performs calculations on the rooms that the user placed in the room group
                    if room_id in rooms_to_be_analysed:
                        

                        np_air_temp = results_reader_file.get_room_results(room_id, 'Room air temperature','Air temperature', 'z')                                 
                        np_dry_resultant_temp = results_reader_file.get_room_results(room_id, 'Comfort temperature', 'Dry resultant temperature','z')                                          
                        np_internal_gain = results_reader_file.get_room_results(room_id, 'Casual gains','Internal gain', 'z')
                        np_solar_gain = results_reader_file.get_room_results(room_id, 'Window solar gains', 'Solar gain','z')
                        np_conduction_gain = results_reader_file.get_room_results(room_id, 'Conduction gain', 'Conduction gain','z')
                        np_infiltration_gain = results_reader_file.get_room_results(room_id, 'Infiltration gain','Infiltration gain', 'z')                        
                        
                        np_space_conditioning_sensible = results_reader_file.get_room_results(room_id, 'System plant etc. gains',  'Space conditioning sensible','z')
                        np_space_conditioning_sensible = np_space_conditioning_sensible 
                        
                        #add room's hourly space con results to total list of each room's data
                        combine_space_con = np_space_conditioning_sensible
                        total_combine_space_con.append(combine_space_con)                
                    
                       
                        peak = min(np_space_conditioning_sensible)                      

                        for hournumber, spacecon in enumerate (np_space_conditioning_sensible):                            
                            if spacecon == peak:
                                peak_hour = hournumber
                                
                                #takes into account if 0 space con so 0 min peak, selects first data point as ies does, instead of last
                                if peak_hour == 119:
                                   new_peak_hour = 0
                                else:
                                    new_peak_hour = peak_hour
  
                        hours = list(range(0,120) )
                        month = 0
                        time = 0
                        
                        if new_peak_hour == 0:
                            month = "May"
                        elif 0<= new_peak_hour <=23:
                            month = "May"
                        elif 24 <= new_peak_hour <=47:
                                month = "June"
                        elif 48 <= new_peak_hour <=71:
                            month = "July"
                        elif 72<= new_peak_hour <=95:
                            month = "August"
                        elif 96<= new_peak_hour <=118:
                            month = "September"
                        else: 
                            month = "May"
                      
                        peak_date = month
                                                
                        midnight = [0, 24, 48, 72, 96, 119]
                        oneam= [i for i in midnight]
                        twoam = [i +1 for i in midnight]
                        threeam = [i +2 for i in midnight]
                        fouram = [i +3 for i in midnight]
                        fiveam = [i +4 for i in midnight]
                        sixam = [i +5 for i in midnight]
                        sevenam = [i +6 for i in midnight]
                        eightam = [i +7 for i in midnight]
                        nineam = [i +8 for i in midnight]
                        tenam = [i +9 for i in midnight]
                        elevenam = [i +10 for i in midnight]
                        twelveam = [i +11 for i in midnight]
                        onepm = [i +12 for i in midnight]
                        twopm = [i +13 for i in midnight]
                        threepm = [i +14 for i in midnight]
                        fourpm = [i +15 for i in midnight]
                        fivepm = [i +16 for i in midnight]
                        sixpm = [i +17 for i in midnight]
                        sevenpm = [i +18 for i in midnight]
                        eightpm = [i +19 for i in midnight]
                        ninepm = [i +20 for i in midnight]
                        tenpm = [i +21 for i in midnight]
                        elevenpm = [i +22 for i in midnight]
                        twelvepm = [i +23 for i in midnight]                        
                      
                        if new_peak_hour in oneam:
                            time = "01:00"
                        elif new_peak_hour in twoam:
                            time = "02:00"
                        elif new_peak_hour in threeam:
                            time = "03:00"                        
                        elif new_peak_hour in fouram:
                            time = "04:00"
                        elif new_peak_hour in fiveam:
                            time = "05:00"
                        elif new_peak_hour in sixam:
                            time = "06:00"                        
                        elif new_peak_hour in sevenam:
                            time = "07:00"
                        elif new_peak_hour in eightam:
                            time = "08:00"
                        elif new_peak_hour in nineam:
                            time = "09:00"
                        elif new_peak_hour in tenam:
                            time = "10:00"
                        elif new_peak_hour in elevenam:
                            time = "11:00"
                        elif new_peak_hour in twelveam:
                            time = "12:00"               
                        elif new_peak_hour in onepm:
                            time = "13:00"
                        elif new_peak_hour in twopm:
                            time = "14:00"
                        elif new_peak_hour in threepm:
                            time = "15:00"
                        elif new_peak_hour in fourpm:
                            time = "16:00"
                        elif new_peak_hour in fivepm:
                            time = "17:00"
                        elif new_peak_hour in sixpm:
                            time = "18:00"
                        elif new_peak_hour in sevenpm:
                            time = "19:00"
                        elif new_peak_hour in eightpm:
                            time = "20:00"
                        elif new_peak_hour in ninepm:
                            time = "21:00"
                        elif new_peak_hour in tenpm:
                            time = "22:00"
                        elif new_peak_hour in elevenpm:
                            time = "23:00"
                        elif new_peak_hour in twelvepm:
                            time = "24:00"
                        else:
                            time = "01:00"
                       
                        
                        peak_hour_time = time
                        
                        day = 24
                       
                        for n in range(0,365):
                            peak_hour_data=np_space_conditioning_sensible[(n*day):(n*day)]                       
                        
                        for air_temp_number, airtemp in enumerate (np_air_temp):
                            if air_temp_number == new_peak_hour:
                                peak_air_temp = float(airtemp)
                                peak_air_temp = round(peak_air_temp,2)                                
                      
                        for dry_resultant_number, dryresultanttemp in enumerate (np_dry_resultant_temp):
                            if dry_resultant_number == new_peak_hour:
                                peak_dry_res_temp = float(dryresultanttemp)
                                peak_dry_res_temp = round(peak_dry_res_temp,2)
                                
                        for internal_gain_number, internalgain in enumerate (np_internal_gain):
                            if internal_gain_number == new_peak_hour:
                                peak_internal_gain = float(internalgain/1000)
                                peak_internal_gain = round(peak_internal_gain,2)
                                
                        for solar_gain_number, solargain in enumerate (np_solar_gain):
                            if solar_gain_number == new_peak_hour:
                                peak_solar_gain = float(solargain/1000)
                                peak_solar_gain = round(peak_solar_gain,2)       

                        for conduction_gain_number, conductiongain in enumerate (np_conduction_gain):
                            if conduction_gain_number == new_peak_hour:
                                peak_conduction_gain = float(conductiongain/1000)
                                peak_conduction_gain = round(peak_conduction_gain,2)   
  
                        for infiltration_gain_number, infiltrationgain in enumerate (np_infiltration_gain):
                            if infiltration_gain_number == new_peak_hour:
                                peak_infiltration_gain = float(infiltrationgain/1000)
                                peak_infiltration_gain = round(peak_infiltration_gain,2)  
                            
                        peak_space_con = float(peak/1000)
                        peak_space_con = round(peak_space_con,2)
                        
                 
                        room_hg_data = [name, rounded_room_area, peak_date, peak_hour_time, peak_air_temp, peak_dry_res_temp, peak_internal_gain, peak_solar_gain, peak_conduction_gain, peak_infiltration_gain, peak_space_con]
                        hg_data.append(room_hg_data) 
                        

            
                combined_space_con_hourly = []
                for n in range(0,118):
                    room_hour_ordered = []
                    for room_number_to_combine, combine_value in enumerate (total_combine_space_con):
                        room_hour_ordered.append(combine_value[n])#for n in range(0,119)
                        combined_room_hour_ordered = sum(i for i in room_hour_ordered)

                    combined_space_con_hourly.append(combined_room_hour_ordered)
           
                peak_combined_space_con = min(combined_space_con_hourly)
                             
                for combined_hournumber, combined_spacecon in enumerate (combined_space_con_hourly):                            
                    if combined_spacecon == peak_combined_space_con:
                        combined_peak_hour = combined_hournumber
                                
                        #takes into account if 0 space con so 0 min peak, selects first data point as ies does, instead of last
                        if combined_peak_hour == 119:
                            combined_new_peak_hour = 0
                        else:
                            combined_new_peak_hour = combined_peak_hour
  
  
                hours = list(range(0,120) )
                combined_month = 0
                combined_time = 0
                        
                if combined_new_peak_hour == 0:
                            combined_month = "May"
                elif 0<= combined_new_peak_hour <=23:
                            combined_month = "May"
                elif 24 <= combined_new_peak_hour <=47:
                                combined_month = "June"
                elif 48 <= combined_new_peak_hour <=71:
                            combined_month = "July"
                elif 72<= combined_new_peak_hour <=95:
                            combined_month = "August"
                elif 96<= combined_new_peak_hour <=118:
                            combined_month = "September"
                else: 
                    combined_month = "May"
                      
                combined_peak_date = combined_month                        
                        
                midnight = [0, 24, 48, 72, 96, 119]
                oneam= [i for i in midnight]
                twoam = [i +1 for i in midnight]
                threeam = [i +2 for i in midnight]
                fouram = [i +3 for i in midnight]
                fiveam = [i +4 for i in midnight]
                sixam = [i +5 for i in midnight]
                sevenam = [i +6 for i in midnight]
                eightam = [i +7 for i in midnight]
                nineam = [i +8 for i in midnight]
                tenam = [i +9 for i in midnight]
                elevenam = [i +10 for i in midnight]
                twelveam = [i +11 for i in midnight]
                onepm = [i +12 for i in midnight]
                twopm = [i +13 for i in midnight]
                threepm = [i +14 for i in midnight]
                fourpm = [i +15 for i in midnight]
                fivepm = [i +16 for i in midnight]
                sixpm = [i +17 for i in midnight]
                sevenpm = [i +18 for i in midnight]
                eightpm = [i +19 for i in midnight]
                ninepm = [i +20 for i in midnight]
                tenpm = [i +21 for i in midnight]
                elevenpm = [i +22 for i in midnight]
                twelvepm = [i +23 for i in midnight]                        
                      
                if combined_new_peak_hour in oneam:
                    combined_time = "01:00"
                elif combined_new_peak_hour in twoam:
                    combined_time = "02:00"
                elif combined_new_peak_hour in threeam:
                    combined_time = "03:00"                        
                elif combined_new_peak_hour in fouram:
                    combined_time = "04:00"
                elif combined_new_peak_hour in fiveam:
                    combined_time = "05:00"
                elif combined_new_peak_hour in sixam:
                    combined_time = "06:00"                        
                elif combined_new_peak_hour in sevenam:
                    combined_time = "07:00"
                elif combined_new_peak_hour in eightam:
                    combined_time = "08:00"
                elif combined_new_peak_hour in nineam:
                    combined_time = "09:00"
                elif combined_new_peak_hour in tenam:
                    combined_time = "10:00"
                elif combined_new_peak_hour in elevenam:
                    combined_time = "11:00"
                elif combined_new_peak_hour in twelveam:
                    combined_time = "12:00"               
                elif combined_new_peak_hour in onepm:
                    combined_time = "13:00"
                elif combined_new_peak_hour in twopm:
                    combined_time = "14:00"
                elif combined_new_peak_hour in threepm:
                    combined_time = "15:00"
                elif combined_new_peak_hour in fourpm:
                    combined_time = "16:00"
                elif combined_new_peak_hour in fivepm:
                    combined_time = "17:00"
                elif combined_new_peak_hour in sixpm:
                    combined_time = "18:00"
                elif combined_new_peak_hour in sevenpm:
                    combined_time = "19:00"
                elif combined_new_peak_hour in eightpm:
                    combined_time = "20:00"
                elif combined_new_peak_hour in ninepm:
                    combined_time = "21:00"
                elif combined_new_peak_hour in tenpm:
                    combined_time = "22:00"
                elif combined_new_peak_hour in elevenpm:
                    combined_time = "23:00"
                elif combined_new_peak_hour in twelvepm:
                    combined_time = "24:00"
                else:
                    combined_time = "01:00"                       
                
 
                combined_peak_hour_time = combined_time
                
                combined_space_con_data = [ peak_combined_space_con, combined_peak_date,combined_peak_hour_time]
                hg_data.append(combined_space_con_data)

                return hg_data

               
            hg_data = get_hg_data(results_reader_file, rooms_to_be_analysed, clg_file_name)

            
            combined_peak_spacecon_data = hg_data[-1]
            combined_peak_spacecon_kW = round((combined_peak_spacecon_data[0] / 1000),2)            
                                  
####################### WRITE DATA TO EXCEL ################         
                                   
            print('Writing results to Excel Sheet')
            
            
            # write hl column headings in sheet 2
            hl_heading = ['Room Name', 'Room Area (m\u00b2)', 'Air temperature (\u2070C)', 'Dry resultant temperature (\u2070C)', 'External conduction gain (kW)', 'Internal conduction gain (kW)', 'Infiltration gain (kW)', 'Steady state heating plant load (kW)', 'Steady state heating plant load (W/m\u00b2)']
            sheet2.write_row(5, 0, hl_heading, headingformat)
            
            sheet2.write(5,10, 'Peak heating demand (kW)', headingformat)
            peak_demand = hl_data[-1]
            peak_demand_value = peak_demand [8]  

            sheet2.write(6,10, peak_demand_value, headingformat)
            
            sheet2.write(5,11, 'Peak heating demand (W/m\u00b2)', headingformat)
            peak_demand_per_m2 =  "=ROUND((k7/SUM(B7:B1000))*1000,2)"
            
            sheet2.write(6,11, peak_demand_per_m2, headingformat)

            # write hl results data
            
            #bottom of the table line
            bottomformat = workbook.add_format({'bold': False, 'bottom': False, 'top': False, 'right': True, 'left':True, 'text_wrap':False})
            bottomformat.set_bottom(2)
            bottomformat.set_right(2)
            
            y = 6
            for row in hl_data:
                sheet2.write_row(y, 0, row, sidecolumnformat)
                y += 1
                if y == 6+numberofrooms:
                    break 
                    

                    
            for row_num in range(6, (6+numberofrooms)):  
                heating_plant_load_cell = xl_rowcol_to_cell(row_num, 7)
                room_area_cell = xl_rowcol_to_cell(row_num, 1)
                heating_plant_load_area = '=ROUND(((%s *1000)/ %s),2)' % (heating_plant_load_cell, room_area_cell)
                sheet2.write_formula(row_num,8, heating_plant_load_area,sidecolumnformat)
            
            sheet2.write_string(2, 2, results_file_name)          
              
            
            # write hg column headings in sheet 3
            hg_heading = ['Room Name', 'Room Area (m\u00b2)', 'Peak date', 'Peak time', 'Air temperature (\u2070C)', 'Dry resultant temperature (\u2070C)', 'Internal gain (kW)', 'Solar gain (kW)', 'Conduction gain (kW)', 'Infiltration gain (kW)', 'Space conditioning sensible (kW)', 'Space conditioning sensible (W/m\u00b2)']
            sheet3.write_row(5, 0, hg_heading, headingformat)
            
            sheet3.write(5,13, 'Peak cooling demand (kW)', headingformat)
            sheet3.write(6,13,combined_peak_spacecon_kW,headingformat)   
            sheet3.write(7,13,combined_peak_spacecon_data[1],headingformat) 
            sheet3.write(8,13,combined_peak_spacecon_data[2],headingformat) 
            
            sheet3.write(5,14, 'Peak cooling demand (W/m\u00b2)', headingformat)
            peak_demand_per_m2 =  "=ROUND((n7/SUM(B7:B1000))*1000,2)"            
            sheet3.write(6,14, peak_demand_per_m2,headingformat)

            # write hg results data
            y = 6
            for row in hg_data:
                sheet3.write_row(y, 0, row, sidecolumnformat)
                y += 1
                if y == 6+numberofrooms:
                    break
            
            
            for row_num in range(6, (6+numberofrooms)): 
                w_per_m2_cell = xl_rowcol_to_cell(row_num, 1) 
                spacecon_sensible_cell = xl_rowcol_to_cell(row_num, 10) 
                w_per_m2 = '=ROUND(((%s * 1000) / %s),2)' %(spacecon_sensible_cell, w_per_m2_cell)
                sheet3.write_formula(row_num,11, w_per_m2, sidecolumnformat) 

            sheet2.write_string(2, 2, htg_file_name)
            sheet2.write_string(3, 2, model_name_new)
            sheet3.write_string(2, 2, clg_file_name)
            sheet3.write_string(3, 2, model_name_new)

                                  
            try:
                workbook.close()
            except PermissionError as e:
                print("Couldn't close workbook: ", e)
            os.startfile(self.project_folder + '\\' + self.save_file_name + '.xlsx')
            root.destroy()

    root = tk.Tk()
    pp = Window(root)
    root.mainloop()

if __name__ == '__main__':
    project = iesve.VEProject.get_current_project()
    ve_folder = iesve.get_application_folder()
    results_reader = iesve.ResultsReader
    room_groups = iesve.RoomGroups()

    # generate the tkinter GUI
    generate_window(project, ve_folder, results_reader, room_groups)


