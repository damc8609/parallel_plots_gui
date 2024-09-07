#need to install pyawr

#pip install pywin32
#cd AppData\Local\Programs\Python\Python39\Lib\site-gridages\win32com\client (if using less than Python 3.9.5)
#python .\makepy.py
#select AWR Design Environment

#pip install pyawr

from operator import truediv
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
from tkinter import ttk
import pyawr.mwoffice as mwo
import os
import time
import ast
import pandas as pd
import numpy as np
import mopa
import threading
import multiprocessing
import webbrowser
import paxplot
import sys
import matplotlib.pyplot as plt


app = mopa.app.create_dashboard()
def run_parallel_plots():
     app.run_server()

def launch_parallel_plots():
     webbrowser.open("http://127.0.0.1:8050")

#awrde = mwo.CMWOffice(version='17.0') #may need to account if user has multiple installs

##Design Notes
#Most SAX Basic commands are Project.XXXX
#In Python, we need to use awrde variable as follows: awrde.Project.XXXXX
#can also do project = awrde.Project        #Create "project" variable as a convenience item and then do project.XXXXX similar to SAX Basic

# NumSchem = awrde.Project.Schematics.Count
# for s_idx in range(NumSchem):
#     schem = awrde.Project.Schematics[s_idx]
#     print(schem.Name)
     
t1 = threading.Thread(target=run_parallel_plots)
# t1 = multiprocessing.Process(target=run_parallel_plots)
t1.daemon = True
t1.start()

print("Launching GUI")
awrde = None
modify_measurement_button = None
lp_script_button = None
choose_file_button = None
export_mdf_box = None
isExampleProject = False
graph_name = ""
measurement_name = ""
frame_bg_color = "#f0f0f0"
# frame_bg_color = "#000080"
widget_pad_default = 10
frame_pad_default = 20
output_directory_path = os.getcwd() + "/"
schematic_list = []
current_schematic = ""
current_tabular = ""
check_button_correlator = []
temp_elements = []
tabular_dropdown = None
schematic_dropdown = None
graph_index = 1
file_creation_label = None
debug = True

if len(sys.argv) == 1 or sys.argv[1] != "--debug":
     debug = False

def debug_print(msg):
     if debug == True:
          print(msg)

def step_1_instructions():
     step_1_instructions_string = "Option 1: Open \"Parallel_Plot_Example_New.emp\" in AWR."
     step_1_instructions_string = step_1_instructions_string + " Enter values in the boxes below and click \"Run Load Pull Script\" to launch Load Pull GUI for schematic \"Load Pull Template\".\n\n"
     step_1_instructions_string = step_1_instructions_string + "Option 2: Convert MDF file to .txt file containing tabular data for PAE or output power."
     messagebox.showinfo(title="Step 1 Instructions", message=step_1_instructions_string)
    
def step_2_instructions():
     step_2_instructions_string = "Enter values below to run manual sweep for schematic \"Manual SP and LP Setup\"."
     step_2_instructions_string = step_2_instructions_string + " Each parameter is optional. Use the checkboxes to enable each different variable sweep in AWR."
     step_2_instructions_string = step_2_instructions_string + " It is not possible to modify just one parameter for each sweep. Make sure Start, Stop, and Step are all valid numbers for desired sweeps.\n\n"
     step_2_instructions_string = step_2_instructions_string + " Enter desired graph name and click \"Run Manual Sweep\" to generate tabular data for PAE and output power (other measurements will be added later)."
     step_2_instructions_string = step_2_instructions_string + " Once tabular data has been created, enter a file name to export trace data to a .TXT, if desired."
     messagebox.showinfo(title="Step 2 Instructions", message=step_2_instructions_string)


def populate_sweep_variables(*args):
     global check_button_correlator
     global temp_elements
     global cTableContainer

     for element in temp_elements:
          element.grid_forget()
     schem = awrde.Project.Schematics(current_schematic.get())
     swpVarNames = []
     swpVarIDs = []
     for element in schem.Elements:
          debug_print(element.Name)
          if "SWPVAR" in element.Name:
               swpVarNames.append(element.Parameters('VarName').ValueAsString)
               swpVarIDs.append(element.Name)
     debug_print(swpVarNames)
     check_button_correlator = []
     current_row = 4
     for value in swpVarNames:
          toggled = IntVar()
          swpVarToggle = Checkbutton(master = sweeps_global_frame, text = value, variable =toggled, justify="left", command=toggle_checkbox)
          swpVarStart = Entry(master=sweeps_global_frame)
          swpVarStop = Entry(master=sweeps_global_frame)
          swpVarStep = Entry(master=sweeps_global_frame)
          check_button_correlator.append((toggled, swpVarIDs[swpVarNames.index(value)], swpVarStart, swpVarStop, swpVarStep))
          swpVarToggle.grid(row=current_row, column=0, padx=widget_pad_default, pady=widget_pad_default)
          swpVarStart.grid(row=current_row, column=1, padx=widget_pad_default, pady=widget_pad_default)
          swpVarStop.grid(row=current_row, column=2, padx=widget_pad_default, pady=widget_pad_default)
          swpVarStep.grid(row=current_row, column=3, padx=widget_pad_default, pady=widget_pad_default)
          temp_elements.append(swpVarToggle)
          temp_elements.append(swpVarStart)
          temp_elements.append(swpVarStop)
          temp_elements.append(swpVarStep)

          current_row=current_row+1
     toggle_checkbox()
     populate_tabulars()
     cTableContainer.update_idletasks()
     sweeps_global_frame.update()
     cTableContainer.config(scrollregion=sweeps_global_frame.bbox())


def change_output_directory():
     global output_directory_path
     output_directory_path = filedialog.askdirectory()
     debug_print(output_directory_path)
     output_directory_path = output_directory_path + "/"

def export_mdf_file():
     global root
     popup = Toplevel(root)
     popup.resizable(False, False)
     popup.title("Choose MDF and File Name")

     mdf_label = Label(master = popup, text="MDF File Name (in AWR)")
     mdf_entry = Entry(master = popup)
     destination_label = Label(master = popup, text="Destination File Name")
     destination_entry = Entry(master = popup)

     mdf_label.grid(column=0, row=0, padx=widget_pad_default, pady=widget_pad_default)
     mdf_entry.grid(column=0, row=1, padx=widget_pad_default, pady=widget_pad_default)
     destination_label.grid(column=2, row=0, padx=widget_pad_default, pady=widget_pad_default)
     destination_entry.grid(column=2, row=1, padx=widget_pad_default, pady=widget_pad_default)

     def export_mdf():

          print("Exporting MDF...")
          destination_filename = destination_entry.get()
          if (not destination_filename.endswith(".mdf")):
               destination_filename = destination_filename + ".mdf"
          awrde.Project.DataFiles(mdf_entry.get()).Export(output_directory_path + destination_filename)
          popup.destroy()

     export_button = Button(master=popup, text="Export", command=export_mdf)
     export_button.grid(row = 2, column = 1, padx=widget_pad_default, pady=widget_pad_default)

def convert_mdf_file():
     global root
     global output_directory_path
     mdf_file_name = filedialog.askopenfilename(filetypes = [("MDF Files","*.mdf *.mdif")])
     debug_print(mdf_file_name)
     if mdf_file_name == "":
          return
     popup = Toplevel(root)
     # popup.geometry("200x100")
     popup.resizable(False, False)
     popup.title("Select Data to Extract")


     pout_checked = BooleanVar()
     pae_checked = BooleanVar()



     generate_plot = BooleanVar()
     calculate_q = BooleanVar()
     time_remaining_var = StringVar()
     current_file_var = StringVar()

     pout_check_box = Checkbutton(master = popup, text = "Pout", variable=pout_checked)
     pout_check_box.grid(row = 1, column = 1, padx=50, pady=10)
     pae_check_box = Checkbutton(master = popup, text = "PAE", variable=pae_checked)
     pae_check_box.grid(row = 2, column = 1, padx=50, pady=10)
     output_file_label = Label(text="Output File Name: ", master = popup)
     output_file_entry = Entry(width=50, master=popup)
     output_file_label.grid(row=3, column = 0, padx=10)
     output_file_entry.grid(row=3, column = 1)
     q_check_box = Checkbutton(master=popup, text = "Calculate Q?", variable = calculate_q)
     q_check_box.grid(row = 4, column = 0, pady=5)
     plot_check_box = Checkbutton(master=popup, text = "Generate Plot?", variable = generate_plot)
     plot_check_box.grid(row=5, column=0, pady=5)
     conversion_progress_bar = ttk.Progressbar(master=popup)
     conversion_progress_bar.grid(row = 6, column = 1)
     current_file_label = Label(master=popup, textvariable=current_file_var)
     current_file_label.grid(row=7, column=1)
     time_remaining_label = Label(master=popup, textvariable=time_remaining_var)
     time_remaining_label.grid(row=8, column=1)

     def parse_mdf():

          calculations_to_perform = []
          if pout_checked.get():
               calculations_to_perform.append("Output_Power")
          if pae_checked.get():
               calculations_to_perform.append("PAE")

          for calc in calculations_to_perform:
               current_file_var.set("Generating file for " + calc)
               converted_file_name = output_file_entry.get()
               if converted_file_name == "":
                    converted_file_name = mdf_file_name.split('/')[-1] + "_Converted"
               if len(calculations_to_perform) > 1:
                    converted_file_name = converted_file_name + "_" + calc
               if not converted_file_name.endswith(".txt"):
                    converted_file_name = converted_file_name + ".txt"
               debug_print(converted_file_name)
               output_file = open(output_directory_path+converted_file_name, "a")
               
               def file_print(message):
                    print(message)
                    output_file.write(message + "\n")

               print("Converting...")
               
               output_file.write(calc + " : Frequency\t")

               isSourcePull = False
               isLoadPull = False
               df = pd.read_csv(mdf_file_name, skip_blank_lines=False, sep='^')
               # df = pd.read_csv(mdf_file_name, skip_blank_lines=False)
               #print (df["!Data File Info"])
               row = 1
               sweep_points = -1
               z0 = -1
               n_harm = -1
               num_vars = 0
               parsedToData = False

               data = {}
               # header = ""
               startFreq = -1

               #create dictionary containing list of sweep vars
               #have list of indices that corresponds to where we are in each list
               #increment frontmost index until it is greater than the length of corresponding list -> reset back to 0 and increment next index
               #repeat

               #currently assuming order in !Sweep Info will be identical to order in VAR<> section

               swpVarNames = []
               swpVarValues = []
               swpVarIndices = []

               while (not parsedToData):
                    row_string = str(df["!Data File Info"][row])
                    # print(row_string)
                    if "! Sweep Variable:" in row_string:
                         if "iGammaL1" in row_string:
                              isLoadPull = True
                         if "iGammaS1" in row_string:
                              isSourcePull = True
                         num_vars = num_vars + 1
                         split_sweep = row_string.split(' ')
                         swpVarNames.append(split_sweep[4])
                         swpVarIndices.append(0)
                         if split_sweep[-1][0] == "[" and split_sweep[-1][-1] == "]":
                              swpVarValues.append(ast.literal_eval(split_sweep[-1]))
                         else:
                              numDummyVariables = int(split_sweep[5][1:-1])
                              dummyList = []
                              for i in range(numDummyVariables):
                                   dummyList.append(None)
                              swpVarValues.append(dummyList)
                         if swpVarNames[-1] == "Pwr":
                              for i in range(len(swpVarValues[-1])):
                                   swpVarValues[-1][i] = swpVarValues[-1][i] + 30 #conversion from dBW to dBm
                    if "Total Sweep Points" in row_string:
                         sweep_points = int(row_string.split()[-1])
                    if "Number of Recorded Harmonics" in row_string:
                         n_harm = int(row_string.split()[-1])
                    if "BEGIN HEADER" in row_string:
                         parameters = df["!Data File Info"][row+2].split()
                         n_harm = int(parameters[1])
                         z0 = float(parameters[2])

                    if "END<>" in row_string:
                         parsedToData = True
                    row = row + 1

               debug_print(swpVarNames)
               debug_print(swpVarValues)

               freq_position = swpVarNames.index("Freq") #should always be 0 but generalized this anyways
               if freq_position == -1:
                    print("ERROR: No frequency information found in MDF")
                    return
               
               plot_labels = ["Freq (GHz)"]
               plot_lines = []

               source_string_mag = "Mag_Gamma_Source"
               source_string_ang = "Ang_Gamma_Source"
               load_string_mag = "Mag_Gamma_Load"
               load_string_ang = "Ang_Gamma_Load"

               if calculate_q.get():
                    source_string_mag = source_string_mag + "_"
                    source_string_ang = source_string_ang + "_"
                    load_string_mag = load_string_mag + "_"
                    load_string_ang = load_string_ang + "_"

               if isSourcePull:
                    plot_labels.append(source_string_mag)
                    plot_labels.append(source_string_ang)
               if isLoadPull:
                    plot_labels.append(load_string_mag)
                    plot_labels.append(load_string_ang)

               deduction = 0
               if "Pwr" in swpVarNames:
                    deduction = deduction + 1
               if "iGammaL1" in swpVarNames:
                    deduction = deduction + 1
               if len(swpVarValues[freq_position]) == 1:
                    deduction = deduction + 1

               start_time = time.perf_counter()
               points_done = 0
                    
               step = 1 + n_harm + (num_vars + 2 - deduction)
               start = row + 1 + (num_vars + 2 - deduction) #because iGammaL1 does not show up in !Sweep Variable, assuming this is always the case
               for i in range(start, start + 1 + (sweep_points*step), step):
                    #for j in range(n_harm):
                    j = 0 #remove and add for loop if adding harmonics
                    debug_print(str(df["!Data File Info"][i+j]))
                    if str(df["!Data File Info"][i+j]) == "nan":
                         continue
                    if str(df["!Data File Info"][i+j])[0] == "!":
                         break
                    ab_waves = df["!Data File Info"][i+j].split()

                    a1_real = float(ab_waves[1])
                    # file_print("A1 real is" + str(a1_real))
                    a1_imag = float(ab_waves[2])
                    b1_real = float(ab_waves[3])
                    b1_imag = float(ab_waves[4])
                    a2_real = float(ab_waves[5])
                    a2_imag = float(ab_waves[6])
                    b2_real = float(ab_waves[7])
                    b2_imag = float(ab_waves[8])
                              
                    a1 = a1_real+a1_imag*1j
                    b1 = b1_real+b1_imag*1j
                    a2 = a2_real+a2_imag*1j
                    b2 = b2_real+b2_imag*1j

                    if int(ab_waves[0]) == 1: #always true?
                         vdd = float(ab_waves[11])
                         idd = float(ab_waves[12])
                         vgg = float(ab_waves[9])
                         igg = float(ab_waves[10])
     
                    mag_a1 = np.absolute(a1)
                    mag_a2 = np.absolute(a2)
                    mag_b1 = np.absolute(b1)
                    mag_b2 = np.absolute(b2)

                    if isSourcePull:
                         source_gamma_real = float(ab_waves[13])
                         source_gamma_imag = float(ab_waves[14])
                         mag1 = np.absolute(source_gamma_real+(source_gamma_imag*1j))
                         debug_print("Source Gamma Real is" + str(source_gamma_real))
                         debug_print("Source Gamma Imag is" + str(source_gamma_imag))
                         ang1 = np.arctan2(source_gamma_imag, source_gamma_real)
                         debug_print("Arctan yields" + str(ang1))
                    else:
                         mag1=0
                         ang1=0
                         
                              
                    # mag1 = mag_a1/mag_b1
                    mag2 = mag_a2/mag_b2

                    debug_print("A/B mags")
                    debug_print(mag_a1)
                    debug_print(mag_b1)
                    debug_print(mag_a2)
                    debug_print(mag_b2)

                    debug_print("Mags")
                    debug_print(mag1)
                    debug_print(mag2)

                    # file_print("Gamma should be " + str(mag_a2/mag_b2))
                    # ang1 = np.arctan(a1_imag/a1_real)-np.arctan(b1_imag/b1_real)
                    # ang2 = np.arctan(a2_imag/a2_real)-np.arctan(b2_imag/b2_real)
                    ang2 = np.arctan2(a2_imag, a2_real)-np.arctan2(b2_imag, b2_real)

                    # file_print("Ang1 should be " + str(ang1))
                    # file_print("Ang2 should be " + str(ang2))

                    ang1_deg = 180*ang1/np.pi
                    ang2_deg = 180*ang2/np.pi

                    debug_print("Ang1 in degrees is be " + str(ang1_deg))
                    debug_print("Ang2 in degrees is be " + str(ang2_deg))

                    trans_wave = a1-b1
                    # print(trans_wave/a1)
                    output_wave = b2
                    # print("Scalar gain equals:", np.absolute(output_wave/a1))
                    # print("Logarithmic Gain equals:", 20*np.log10(np.absolute(output_wave/a1)), " dB")
                    p_out = 0.5*((mag_b2**2)-(mag_a2**2))
                    p_in = 0.5*(mag_a1**2)
                    # print(p_out)
                    # print(p_in)

                    plot_line = []
                    
                    frequency = str(int(swpVarValues[freq_position][swpVarIndices[freq_position]]/1000000000))

                    plot_line.append(int(swpVarValues[freq_position][swpVarIndices[freq_position]]/1000000000))
                    if isSourcePull:
                         plot_line.append(round(mag1, 5))
                         plot_line.append(round(ang1_deg, 5))
                    if isLoadPull:
                         plot_line.append(round(mag2, 5))
                         plot_line.append(round(ang2_deg, 5))
                    if data == {}:
                         startFreq = frequency

                    for swpVarIndex in range(len(swpVarNames)):
                              if swpVarValues[swpVarIndex][swpVarIndices[swpVarIndex]] == None or swpVarNames[swpVarIndex] == "Freq":
                                   continue
                              else:
                                   if swpVarNames[swpVarIndex] not in plot_labels:
                                        plot_labels.append(swpVarNames[swpVarIndex])
                                   plot_line.append(round(swpVarValues[swpVarIndex][swpVarIndices[swpVarIndex]], 5))

                    def ang_fix(ang):
                         if ang < 0:
                              return ang + 180
                         return ang
                              
                    if frequency == startFreq:
                         if isSourcePull:
                              # header = header + "Source_Gamma_Mag = " + str(round(mag1, 5)) + " Source_Gamma_Ang = " + str(round(ang_fix(ang1_deg), 5)) + " "
                              output_file.write(source_string_mag + " = " + str(round(mag1, 5)) + " " + source_string_ang + " = " + str(round(ang_fix(ang1_deg), 5)) + " ")
                         if isLoadPull:
                              # header = header + "Load_Gamma_Mag = " + str(round(mag2, 5)) + " Load_Gamma_Ang = " + str(round(ang_fix(ang2_deg), 5)) + " "
                              output_file.write(load_string_mag + " = " + str(round(mag2, 5)) + " " + load_string_ang + " = " + str(round(ang_fix(ang2_deg), 5)) + " ")
                         for swpVarIndex in range(len(swpVarNames)):
                              if swpVarValues[swpVarIndex][swpVarIndices[swpVarIndex]] == None or swpVarNames[swpVarIndex] == "Freq":
                                   continue
                              else:
                                   # header = header + swpVarNames[swpVarIndex] + " = " + str(round(swpVarValues[swpVarIndex][swpVarIndices[swpVarIndex]], 5)) + " "
                                   output_file.write(swpVarNames[swpVarIndex] + " = " + str(round(swpVarValues[swpVarIndex][swpVarIndices[swpVarIndex]], 5)) + " ")
                         # header = header + "\t"
                         output_file.write("\t")
                    if frequency not in data.keys():
                         data[frequency] = ""

                         
                    p_delivered = 0.5*((mag_a1**2)-(mag_b1**2))
                              
                    if calc == "Output_Power":
                         if "Output_Power" not in plot_labels:
                              plot_labels.append("Output_Power")
                         p_out_dbm = 10*np.log10(0.5*((mag_b2**2)-(mag_a2**2)))+30
                         p_in_dbm = 10*np.log10(0.5*(mag_a1**2))+30 #Possibly just missing Zo
                                   
                         data[frequency] = data[frequency] + "\t" + str(round(p_out_dbm, 4))
                         plot_line.append(round(p_out_dbm, 4))
                                   
                    # file_print("Freq is " + str(df["!Data File Info"][i-3]))
                    # file_print("Output power equals:" + str(p_out_dbm) + " dBm") #Missing Zo, should be (b2^2)/2Zo
                    #For load pull, may need to do (b2^2 - a2^2)/2Zo, may not be correct, need to figure out how AWR is normalizing the wave
                    #Most important is factoring in a waves
                    # print(np.absolute(a1))
                    # print(0.5*(np.absolute(a1)**2))

                    # file_print("Input power equals:" + str(10*np.log10(0.5*(mag_a1**2))+30) + " dBm")
                    # print("New G_trans equals:", p_out_db - p_in_db)

                    elif calc == "PAE":
                         if "PAE" not in plot_labels:
                              plot_labels.append("PAE")
                         pdc = (vdd * idd)#+(vgg*igg)
                         pae = ((p_out)-(p_delivered))/pdc
                                   
                         data[frequency] = data[frequency] + "\t" + str(round(pae*100, 4))
                         plot_line.append(round(pae*100, 4))
                         # file_print("PAE is " + str(pae*100) + "%")
                         # print()

                    for k in range(len(swpVarIndices)):
                         if swpVarIndices[k] + 1 != len(swpVarValues[k]):
                              swpVarIndices[k] = swpVarIndices[k] + 1
                              break
                         else:
                              swpVarIndices[k] = 0
                    plot_lines.append(plot_line)
                    debug_print("Step is " +  str(100/float(sweep_points)))

                    points_done = points_done + 1
                    if (points_done % 100 == 0):
                         current_time = time.perf_counter()
                         time_remaining_estimated = round(((current_time-start_time)*(sweep_points-points_done))/points_done)
                         time_remaining_var.set("Estimated Time Remaining: " + str(time_remaining_estimated) + " seconds")

                    conversion_progress_bar.step(100/float(sweep_points))
                    conversion_progress_bar.update_idletasks()


               print("Parsing complete")
               #print all data to .txt in proper format

               # file_print(header[:-2]) #removing last tab and space from header
               output_file.close()
               output_file = open(output_directory_path+converted_file_name, "ab")
               output_file.seek(-2, 2)
               output_file.truncate()
               output_file.close()
               output_file = open(output_directory_path+converted_file_name, "a")
               output_file.write("\n")
               for frequency in data.keys():
                    file_print(frequency + data[frequency])
               output_file.close()
          popup.destroy()

          if generate_plot.get():
               debug_print("Before culling, there are " + str(len(plot_lines)) + " plot lines with a length of " + str(len(plot_lines[0])) +  " and " + str(len(plot_labels)) + " labels")
               i = 0
               while i < len(plot_lines[0]):
                    unique_values = []
                    for j in range(len(plot_lines)):
                         if plot_lines[j][i] not in unique_values:
                              unique_values.append(plot_lines[j][i])
                    if len(unique_values) == 1:
                         plot_labels.pop(i)
                         for j in range(len(plot_lines)):
                              plot_lines[j].pop(i)
                    else:
                         i = i + 1
               debug_print("Plot labels are:")
               debug_print(plot_labels)
               paxfig = paxplot.pax_parallel(n_axes=len(plot_labels))
               debug_print("After culling, there are " + str(len(plot_lines)) + " plot lines with a length of " + str(len(plot_lines[0])) +  " and " + str(len(plot_labels)) + " labels")
               paxfig.plot(plot_lines)
               paxfig.set_labels(plot_labels)
               plt.show()


     done_button = Button(master=popup, text="Convert", command=parse_mdf)
     done_button.grid(row = 5, column = 2, padx=10)



    

def input_checker(entry):
    if entry.get() == "":
         return 0
    try:
         g = float(entry.get())
         return 1
    except:
         print("Invalid parameter input: " + entry.get())
         return -1


def run_load_pull_script():
    global awrde
    global mag1_entry
    global mag2_entry
    global ang1_entry
    global ang2_entry
    global vdd_entry
    global vgg_entry

    for routine in awrde.GlobalScripts('Load_Pull').Routines:
        debug_print(routine)
    # for routine in awrde.GlobalScripts('Open_Project_Directory').Routines:
    #     print(routine)
    # awrde.GlobalScripts('Import_Load_Pull_Files').Routines('Import_Load_Pull_Files').Run()
    # awrde.GlobalScripts('Load_Pull').Routines('importLoadPullTemplate').RunWithArgs([False, True])
    # createTemplate = True
    # for schem in awrde.Project.Schematics:
    #     if 'Load_Pull_Template' in schem.Name:
    #         createTemplate = False
    # if createTemplate:
    #      awrde.GlobalScripts('Load_Pull').Routines('Create_Load_Pull_Template').Run()
    #      time.sleep(5)

    schem = awrde.Project.Schematics('Load_Pull_Template')

    #MODIFY ELEMENT PARAMS DIRECTLY
    # elem = schem.Elements('HBTUNER3.SourceTuner')

    # elem.Parameters('Mag1').ValueAsDouble = mag1_entry.get()
    # elem.Parameters('Mag2').ValueAsDouble = mag2_entry.get()
    # elem.Parameters('Ang1').ValueAsDouble = ang1_entry.get()
    # elem.Parameters('Ang2').ValueAsDouble = ang2_entry.get()

    #GENERIC VERSION FOR IF WE MAKE .EMP BETTER
    # for param in [mag1_entry, ang1_entry, mag2_entry, ang2_entry, vdd_entry, vgg_entry]:
    #     if param.get() != "":
    #         try:
    #             param_val = float(param.get())

    #             schem.Equations('mag1').Expression = "mag1 = " + mag1_entry.get() + "\r\nang1 = " + ang1_entry.get()
    #             schem.Equations('mag2').Expression = "mag2 = " + mag2_entry.get() + "\r\nang2 = " + ang2_entry.get()

    #         except:
    #             print("Invalid parameter input:" + param)
    #             print("Type "+ type(param.get()) " is not valid; Entered input was \"" + param.get() + "\"")

    #PROBLEMS:
    #HAVE TO SET MAG AND ANG TOGETHER BECAUSE THEY ARE IN THE SAME EQUATION
    # WOULD RESULT IN BUG IF TRYING TO CHANGE MAG1 BUT NOT ANG1 THROUGH GUI

    


        
    if (input_checker(mag1_entry) and input_checker(ang1_entry)):
        schem.Equations('mag1').Expression = "mag1 = " + mag1_entry.get() + "\r\nang1 = " + ang1_entry.get()
    if (input_checker(mag2_entry) and input_checker(ang2_entry)):    
        schem.Equations('mag2').Expression = "mag2 = " + mag2_entry.get() + "\r\nang2 = " + ang2_entry.get()
    if (input_checker(vdd_entry)):
        schem.Equations('iVd').Expression = "iVd = " + vdd_entry.get()
    if (input_checker(vgg_entry)):    
        schem.Elements('DCVS.VGS').Parameters('V').ValueAsDouble = vgg_entry.get()


    awrde.GlobalScripts('Load_Pull').Routines('Load_Pull').Run()

    # awrde.GlobalScripts('Load_Pull').Routines('doLoadPull').Run()
    # for i in range(0, len(schem.Elements)):
    #     print(schem.Elements[i])
    # return

def run_manual_sweep():
     global awrde
     global graph_name
     global measurement_name
     global graph_name_entry
     global graph_index

     global modify_measurement_button
     # print(modify_measurement_button.state)
     modify_measurement_button.config(state = 'normal')
     # print(modify_measurement_button.state)

     schem = awrde.Project.Schematics(current_schematic.get())

     #Apply sweep variables
     for i in range(len(check_button_correlator)):
          if (schem.Elements(check_button_correlator[i][1]).Enabled and input_checker(check_button_correlator[i][2]) and input_checker(check_button_correlator[i][3]) and input_checker(check_button_correlator[i][4])):
               
            schem.Elements(check_button_correlator[i][1]).Parameters('Values').ValueAsString = "swpstp(" + check_button_correlator[i][2].get() + "," + check_button_correlator[i][3].get() + "," + check_button_correlator[i][4].get() + ")"
     graph_name = graph_name_entry.get()
     if graph_name == "":
          graph_name = "Tabular " + str(graph_index)
          graph_index = graph_index + 1
     if (awrde.Project.Graphs.Exists(graph_name)):
          #Prompt for overwrite if possible
          awrde.Project.Graphs.Remove(graph_name)
     awrde.Project.Graphs.Add(graph_name, 4) #I believe 4 corresponds to Tabular graph per AWR scripting guide, mwGT_Tabular is undefined
     #Look at Page 389 of AWR Scripting Guide for info on how to retrieve other measurements (MeasurementInfos collection) (not ideal)
     #Graph has InvokeCommand method -> see line 10065 of mwoffice.py
     #Possible command name:? Name = GLOBAL!CE_Cookbook.copy_MWO_TB_And_Add_Graphs	Category = Macros
     #Name = ViewMeasurementEditor	Category = View

     #Comes from Find_Command_Names global script with nothing as input
     # awrde.Project.Graphs(graph_name).StartCommand("ViewMeasurementEditor", "awrde") #This is a valid command! Why won't it show the window?
     # print("?")
     if measurement_name == "":
          measurement_name = current_schematic.get() + ".AP_HB:PAE(PORT_1,PORT_2)"
     measurement_split = measurement_name.split(":")
     awrde.Project.Graphs(graph_name).Measurements.Add(measurement_split[0], measurement_split[1])
     awrde.Project.Simulator.Analyze() #Equivalent of clicking simulate in AWR
     populate_tabulars()
     print("Manual Sweep Complete")
     

     #Potentially add CSimulator.AnalyzeOpenGraphs if we want to automate this instead of have the user do it
def modify_measurement():
     global awrde
     global graph_name
     global measurement_name

     for nd in awrde.Project.TreeView.Nodes("Project").Children("Graphs").Children(graph_name).Children("Measurements").Children:
          if (nd.Name == measurement_name):
               awrde.Project.TreeView.Visible = True
               nd.Selected = True
               awrde.StartCommand("PropertiesObject", 0)
               # awrde.StartCommand("ProjectAddMeasurement", 0)
               # awrde.StartCommand("AddMeasurement", 0)
               # awrde.StartCommand("Add Measurement", 0)
               awrde.Project.Simulator.Analyze()
               
     measurement_name = awrde.Project.TreeView.Nodes("Project").Children("Graphs").Children(graph_name).Children("Measurements").Children[0].Name


def create_txt_file():
     global filename_entry
     global awrde
     global output_directory_path
     global file_creation_label
     global sweeps_global_frame

     current_graph_name = current_tabular.get()
     filename = filename_entry.get()
     if filename == "":
          filename = current_graph_name + "_Tabular_Data"
     if len(filename) > 32: 
          filename = filename[0:32]
     if (not filename.endswith(".txt")):
        filename = filename + ".txt"
     awrde.Project.Graphs(current_graph_name).ExportTraceData(output_directory_path+filename)

     messagebox.showinfo(title="Success", message=str(filename) + " successfully created!")

def toggle_checkbox():
     schem = awrde.Project.Schematics(current_schematic.get())
     for pairing in check_button_correlator:
          if pairing[0].get():
               schem.Elements(pairing[1]).Enabled = True
          else:
               schem.Elements(pairing[1]).Enabled = False

def goto_main_menu():
     global method_global_frame
     global direct_method_frame
     global specific_method_frame
     global load_pull_tutorial_frame
     global sweeps_global_frame

     sweeps_global_frame.pack_forget()
     cTableContainer.pack_forget()
     direct_method_frame.pack_forget()
     specific_method_frame.pack_forget()
     load_pull_tutorial_frame.pack_forget()

     method_global_frame.pack(padx=frame_pad_default, pady=frame_pad_default)

def goto_direct_frame():
     global method_global_frame
     global direct_method_frame

     method_global_frame.pack_forget()
     direct_method_frame.pack(padx=frame_pad_default, pady=frame_pad_default)

def goto_specific_frame():
     global method_global_frame
     global specific_method_frame

     method_global_frame.pack_forget()
     specific_method_frame.pack(padx=frame_pad_default, pady=frame_pad_default)

def goto_tutorial():
     global sweeps_global_frame
     global load_pull_tutorial_frame
     load_pull_tutorial_frame.pack(padx=frame_pad_default, pady=frame_pad_default)
     sweeps_global_frame.pack_forget()

def goto_sweeps():
     global sweeps_global_frame
     global load_pull_tutorial_frame
     global sbVerticalScrollBar
     global cTableContainer

     sbVerticalScrollBar.pack(fill=Y, side=RIGHT, expand=FALSE)
     cTableContainer.pack(fill=BOTH, side=LEFT, expand=TRUE)
     cTableContainer.create_window(0, 0, window=sweeps_global_frame, anchor=NW)

     populate_schematic_list()
     
     if isExampleProject:
          goto_tutorial_button.grid(row=102, column=0, padx=widget_pad_default, pady=widget_pad_default)
     
     load_pull_tutorial_frame.pack_forget()

def populate_tabulars():
     global sweeps_global_frame
     global current_tabular
     global awrde
     global tabular_dropdown

     if tabular_dropdown != None:
          tabular_dropdown.grid_forget()

     tabular_list = []

     for graph in awrde.Project.Graphs:
          debug_print(graph.Name)
          if graph.Type == 4: #Tabular
               tabular_list.append(graph.Name)

     current_tabular = StringVar()
     current_tabular.set(tabular_list[0])
     tabular_dropdown = OptionMenu(
          sweeps_global_frame,
          current_tabular,
          *tabular_list
     )

     tabular_dropdown.grid(row=100, column=2, padx=widget_pad_default, pady=widget_pad_default)

def launch_example_project():
     global awrde
     global root
     global specific_method_frame
     global load_pull_tutorial_frame
     global isExampleProject
     isExampleProject = True
     awrde = mwo.CMWOffice()
     awrde.Open(Filename=os.getcwd() + "/Parallel_Plot_Example_File.emp")
     specific_method_frame.pack_forget()
     load_pull_tutorial_frame.pack(padx=frame_pad_default, pady=frame_pad_default)

def emp_file_dialog():
     global awrde
     global root
     global specific_method_frame
     global sweeps_global_frame
     project_name = filedialog.askopenfilename(filetypes = [("AWR Project Files","*.emp")])
     if (project_name == ""):
          return
     awrde = mwo.CMWOffice()
     awrde.Open(Filename=project_name)
     specific_method_frame.pack_forget()
     goto_sweeps()


def populate_schematic_list():
     global awrde
     global schematic_list
     global schematic_dropdown
     global current_schematic
     global schematic_list_subframe

     schematic_list = []
     for schematic in awrde.Project.Schematics:
          debug_print(schematic.Name)
          schematic_list.append(schematic.Name)

     if schematic_dropdown != None:
          schematic_dropdown.grid_forget()


     current_schematic = StringVar()
     current_schematic.set(schematic_list[0])
     schematic_dropdown = OptionMenu(
          schematic_list_subframe,
          current_schematic,
          *schematic_list
     )

     current_schematic.trace_add('write', populate_sweep_variables)

     schematic_dropdown.grid(row=0, column=0)
     populate_sweep_variables()
    
class Root(Tk):

    def __init__(self):
        super(Root,self).__init__()
 
        self.title("Parallel Plots GUI")
        #self.minsize(600,400)

        ##METHOD FRAME

        global method_global_frame
        method_global_frame = Frame(bg=frame_bg_color)

        method_label = Label(master=method_global_frame, text="Choose Conversion Method")

        direct_method_button = Button(
            master = method_global_frame,
            text="Direct Method",
            width=20,
            height=3,
            command=goto_direct_frame
        )

        specific_method_button = Button(
            master = method_global_frame,
            text="Specific Method",
            width=20,
            height=3,
            command=goto_specific_frame
        )

        output_directory_button = Button(
             master=method_global_frame,
             text="Choose Output Directory for Generated Files",
             width=40,
             height=3,
             command=change_output_directory
        )
        
        method_label.grid(row=0, column=1, padx=widget_pad_default, pady=widget_pad_default)
        direct_method_button.grid(row = 1, column = 0, padx=widget_pad_default, pady=widget_pad_default)
        specific_method_button.grid(row= 1, column = 2, padx=widget_pad_default, pady=widget_pad_default)
        output_directory_button.grid(row=2, column=1, padx=widget_pad_default, pady=widget_pad_default)

        #DIRECT FRAME

        global direct_method_frame
        direct_method_frame = Frame(bg=frame_bg_color)

        main_menu_button_1 = Button(
            master = direct_method_frame,
            text="Main Menu",
            width=10,
            height=1,
            command=goto_main_menu
        )

        convert_mdf_file_button = Button(
            master = direct_method_frame,
            text="Convert MDF File",
            width=20,
            height=3,
            command=convert_mdf_file
        )

        parallel_plots_button = Button(
            master = direct_method_frame,
            text="Open Parallel Plots",
            width=20,
            height=3,
            command=launch_parallel_plots
        )

        main_menu_button_1.grid(row = 0, column = 0, padx=widget_pad_default, pady=widget_pad_default)
        convert_mdf_file_button.grid(row = 1, column = 0, padx=widget_pad_default, pady=widget_pad_default)
        parallel_plots_button.grid(row= 1, column = 1, padx=widget_pad_default, pady=widget_pad_default)

        #SPECIFIC FRAME

        global specific_method_frame

        specific_method_frame = Frame(bg=frame_bg_color)

        main_menu_button_2 = Button(
            master = specific_method_frame,
            text="Main Menu",
            width=10,
            height=1,
            command=goto_main_menu
        )

        example_project_button = Button(
            master = specific_method_frame,
            text="Launch Example Project",
            width=20,
            height=3,
            command=launch_example_project  
        )

        global choose_file_button
        choose_file_button = Button(
            master = specific_method_frame,
            text="Choose .EMP",
            width=20,
            height=3,
            command=emp_file_dialog
        )

        main_menu_button_2.grid(row = 0, column = 0, padx=widget_pad_default, pady=widget_pad_default)
        example_project_button.grid(row=1, column=0, padx=widget_pad_default, pady=widget_pad_default)
        choose_file_button.grid(row=1, column=1, padx=widget_pad_default, pady=widget_pad_default)

        #SWEEPS FRAME

        #CREDIT TO JOE HUTCH FOR THIS SCROLLABLE FRAME SOLUTION: https://www.joehutch.com/posts/tkinter-dynamic-scroll-area/

        global sbVerticalScrollBar
        global cTableContainer
        cTableContainer = Canvas(width=750, height=700)

        sbVerticalScrollBar = Scrollbar()

        global sweeps_global_frame
        sweeps_global_frame = Frame( master = cTableContainer, bg=frame_bg_color)

        cTableContainer.config(yscrollcommand=sbVerticalScrollBar.set, highlightthickness=0)
        sbVerticalScrollBar.config(orient=VERTICAL, command=cTableContainer.yview)
        def _on_mousewheel(event):
          cTableContainer.yview_scroll(int(-1*(event.delta/120)), "units")

        cTableContainer.bind_all("<MouseWheel>", _on_mousewheel)



        global goto_tutorial_button
        goto_tutorial_button = Button(
            master = sweeps_global_frame,
            text="Previous",
            width=10,
            height=3,
            command=goto_tutorial
        )

        main_menu_button_3 = Button(
            master = sweeps_global_frame,
            text="Main Menu",
            width=10,
            height=1,
            command=goto_main_menu
        )

        start_sweep_label = Label(master = sweeps_global_frame, text="Start")
        end_sweep_label = Label(master = sweeps_global_frame, text="End")
        step_sweep_label = Label(master = sweeps_global_frame, text="Step")

        parallel_plots_button_2 = Button(
            master = sweeps_global_frame,
            text="Open Parallel Plots",
            width=20,
            height=3,
            command=launch_parallel_plots
        )


        populate_sweep_variables_button = Button(
            master = sweeps_global_frame,
            text="Repopulate Sweep Variables",
            width=25,
            height=3,
            command=populate_sweep_variables
        )

        global schematic_list_subframe
        global update_schematic_list_button

        schematic_list_subframe = Frame(master=sweeps_global_frame)

        update_schematic_list_button = Button(
            master = schematic_list_subframe,
            text=u'\u21BB',
            width=1,
            height=1,
            command=populate_schematic_list  
        )

        update_schematic_list_button.grid(row=0, column=1)

        run_manual_sweep_button = Button(
            master = sweeps_global_frame,
            text="Run Manual Sweep",
            width=20,
            height=3,
            command=run_manual_sweep
        )

        create_txt_file_button = Button(
            master = sweeps_global_frame,
            text="Create File",
            width=20,
            height=3,
            command=create_txt_file
        )
        
        global modify_measurement_button
        modify_measurement_button = Button(
            master = sweeps_global_frame,
            text="Modify Measurement",
            width=20,
            height=3,
            state='disabled',
            command=modify_measurement
        )

       
        
        global filename_entry
        global graph_name_entry

        graph_name_label = Label(master = sweeps_global_frame, text="Tabular Graph Name:")
        graph_name_entry = Entry(master = sweeps_global_frame)

        filename_label = Label(master = sweeps_global_frame, text="File Name:")
        filename_entry = Entry(master = sweeps_global_frame)
        schematic_label = Label(master=sweeps_global_frame, text="Choose Schematic:")
        tabular_label = Label(master=sweeps_global_frame, text="Tabular File to Export:")

        schematic_label.grid(row=0, column=0)

        main_menu_button_3.grid(row=0, column=3, padx=widget_pad_default, pady=widget_pad_default)
        schematic_list_subframe.grid(row=1, column = 0, padx=widget_pad_default, pady=widget_pad_default)
        parallel_plots_button_2.grid(row=1, column=3, padx=widget_pad_default, pady=widget_pad_default)
        populate_sweep_variables_button.grid(row=2, column=0, padx=widget_pad_default, pady=widget_pad_default)
        
        start_sweep_label.grid(row=3, column=1, padx=widget_pad_default, pady=widget_pad_default)
        end_sweep_label.grid(row=3, column=2, padx=widget_pad_default, pady=widget_pad_default)
        step_sweep_label.grid(row=3, column=3, padx=widget_pad_default, pady=widget_pad_default)

        #All variables start at row 4, there is room for 95 different SWPVARS
        

        graph_name_label.grid(row=99, column=0)
        graph_name_entry.grid(row=100, column=0, padx=widget_pad_default, pady=widget_pad_default)
        run_manual_sweep_button.grid(row=100, column=1, padx=widget_pad_default, pady=widget_pad_default)
        modify_measurement_button.grid(row=101, column=1, padx=widget_pad_default, pady=widget_pad_default)
        tabular_label.grid(row = 99, column=2)
        #dropdown box row 100 column 2
        filename_label.grid(row=101, column=2)
        filename_entry.grid(row=102, column=2)
        create_txt_file_button.grid(row=102, column=3, padx=widget_pad_default, pady=widget_pad_default)

        #EXAMPLE PROJECT TUTORIAL FRAME

        global load_pull_tutorial_frame
        load_pull_tutorial_frame = Frame(bg=frame_bg_color)

        main_menu_button_4 = Button(
            master = load_pull_tutorial_frame,
            text="Main Menu",
            width=10,
            height=1,
            command=goto_main_menu
        )

        goto_sweeps_button = Button(
            master = load_pull_tutorial_frame,
            text="Next",
            width=10,
            height=3,
            command=goto_sweeps
        )


        global lp_script_button
        lp_script_button = Button(
            master = load_pull_tutorial_frame,
            text="Run Load Pull Script",
            width=20,
            height=3,
            command=run_load_pull_script,
        )

        export_mdf_file_button = Button(
            master = load_pull_tutorial_frame,
            text="Export MDF",
            width=20,
            height=3,
            command=export_mdf_file
        )


        mag1_label = Label(master = load_pull_tutorial_frame, text="Mag1")
        mag2_label = Label(master = load_pull_tutorial_frame, text="Mag2")
        ang1_label = Label(master = load_pull_tutorial_frame, text="Ang1")
        ang2_label = Label(master = load_pull_tutorial_frame, text="Ang2")
        vdd_label = Label(master = load_pull_tutorial_frame, text="Vdd")
        vgg_label = Label(master = load_pull_tutorial_frame, text="Vgg")

        global mag1_entry
        global mag2_entry
        global ang1_entry
        global ang2_entry
        global vdd_entry
        global vgg_entry
        mag1_entry = Entry(master = load_pull_tutorial_frame)
        mag2_entry = Entry(master = load_pull_tutorial_frame)
        ang1_entry = Entry(master = load_pull_tutorial_frame)
        ang2_entry = Entry(master = load_pull_tutorial_frame)
        vdd_entry = Entry(master = load_pull_tutorial_frame)
        vgg_entry = Entry(master = load_pull_tutorial_frame)

        main_menu_button_4.grid(row=0,column=0, padx=widget_pad_default, pady=widget_pad_default)
        mag1_label.grid(row=1,column=0, padx=widget_pad_default, pady=widget_pad_default)
        mag1_entry.grid(row=2,column=0, padx=widget_pad_default, pady=widget_pad_default)
        mag2_label.grid(row=3,column=0, padx=widget_pad_default, pady=widget_pad_default)
        mag2_entry.grid(row=4,column=0, padx=widget_pad_default, pady=widget_pad_default)
        ang1_label.grid(row=1,column=1, padx=widget_pad_default, pady=widget_pad_default)
        ang1_entry.grid(row=2,column=1, padx=widget_pad_default, pady=widget_pad_default)
        ang2_label.grid(row=3,column=1, padx=widget_pad_default, pady=widget_pad_default)
        ang2_entry.grid(row=4,column=1, padx=widget_pad_default, pady=widget_pad_default)
        vdd_label.grid(row=1,column=2, padx=widget_pad_default, pady=widget_pad_default)
        vdd_entry.grid(row=2,column=2, padx=widget_pad_default, pady=widget_pad_default)
        vgg_label.grid(row=3,column=2, padx=widget_pad_default, pady=widget_pad_default)
        vgg_entry.grid(row=4,column=2, padx=widget_pad_default, pady=widget_pad_default)
        export_mdf_file_button.grid(row = 6, column = 0, padx=widget_pad_default, pady=widget_pad_default)
        lp_script_button.grid(row=5, column=0, padx=widget_pad_default, pady=widget_pad_default)
        goto_sweeps_button.grid(row=6, column=2, padx=widget_pad_default, pady=widget_pad_default)
        
        #### LAUNCH FIRST WINDOW

        method_global_frame.pack(padx=frame_pad_default, pady=frame_pad_default)
    
global root


root = Root()
root.resizable(False, False)
root.mainloop()

# t1.terminate()

#TODO :

#Fix newline issue with setting mag/ang params - would be helpful if they were all separate equations but they are not

