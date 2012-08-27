#!/usr/bin/env python
##This is a refactored Python version of my spatial pattern analysis program, published 
##in Morgan et al., 2012. The program examines the distribution of cellular or other 
##populations to determine whether they are randomly spaced or instead more or less closely 
##clustered together than expected. This version assumes no layer information (i.e. a tissue 
##sample without cortical layers). For a non-homogeneous, layered sample, please see 
##SpatialPatternRefactor.py.

##Summary of method:
##This program looks at the recorded x, y-coordinates of each cell in a region of interest.
##If the cell is far enough from the boundaries of the ROI that we can investigate the full
##distance range around it without running into ROI edge effects, the program calculates the 
##distance to all of the other cells of the appropriate cell class in the sample. It then
##marks each matching instance at the distance between the cells and all of the farther 
##distances within the analysis range.

##The program then generates simulations of cellular distribution, which have randomized cell
##locations. It compares an average of these simulation values to the actual distribution of 
##cells to detect inhomogeneities in their organization.

####################################################################
##Section 1: setting up your computer to run this program
##This program is written in Python 2.6-2.7. It also uses the xlwt addon library to make Excel 
##spreadsheets.

##Step 1: Install Python 2.7.3 from here: http://www.python.org/download/releases/2.7.3/
##On the page, select the Windows MSI installer (x86 if you have 32-bit Windows installed,
##x86-64 if you have 64-bit Windows installed.)
##I suggest using the default option, which will install Python to c:/Python27

##Step 2: Install the xlwt library from here: http://pypi.python.org/pypi/xlwt/
##Use the program WinRAR to unzip the files to a directory
##Go to "run" in the start menu and type cmd
##Type cd c:\directory_where_xlwt_was_unzipped_to
##Type setup.py install

##Step 3: Copy this program into the c:/Python27 directory
##You can also put it into another directory that is added to the correct PATH.

####################################################################
##Section 2: file preparation for this program (and related scripts)

##Record coordinates of all cells belonging to the populations of interest in a rectangular
##counting frame in Stereo Investigator. Take the raw cellular coordinates and save as a 
##.txt file. Set your variables in the section below, and run the program.

####################################################################
##User set variables here

#Directory containing input files
directory = "C:\Users\John Morgan\Documents\sp_datafiles"

#Save cell coordinate data from Stereo Investigator as a tab-delimited text file
#There should be a header, leave it in
#This test file is actually from a layered sample. Note the strong clustering
#this program reports with this input. This is because the layering effectively
#packs many of the cells in dense clusters relative to a non-layered distribution.
inputfile = "NoLayer_B4925.txt"

#This is the name of the .xls file that the program will save when finished. 
outputfile = "nolayer_test"

#This is the name of the .xls file that the program will save when finished. 
outputfile = "test"

##This variable adjusts the cell types that are being compared. Input the values in the
##first column of your saved .txt file.
##In the demo run, neuron = 1, microglia = 3 
##To look at the properties of a single population, set both of these values the same.
cell1 = 1
cell2 = 3

##This variable sets distance from the ROI boundary at which seed cells will be
##excluded from analysis to avoid edge effects
exclude_dist = 100

##This variable adjusts the studied distance and interval in the cleaned output file.
##Do not exceed the excluded distance or you will have edge effects distorting your
##results!
analysis_dist = 100

##This variable adjusts the number of interrogation points over the analysis range.
##It is strongly suggested to match analysis_dist unless you are looking to cut noise
##with larger distance bins.
interval_num = 100

##The number of simulations to run; the results are averaged and then the clustering
##values are divided by the simulation results to produce a clustering ratio.
##In Morgan et al., 2012, I ran 200 simulations/condition. Experimentation indicated 
##that this number of simulations resulted in <1% variability in results.
sim_run_num = 5

####################################################################
##Program begins here

import csv
import xlwt
import math
import random
import time
print time.clock()

#Load and clean up file, output is sp_data which is a list of all cells 
#sp_data format is [[celltype1, xcoord1, ycoord1],[celltype2, xcoord2, ycoord2], etc]
def loadfile():
    path = directory + "\\" + inputfile
    input_file_obj = open(path, "r") 
    csv_read = csv.reader(input_file_obj, dialect = csv.excel_tab)
    sp_data = []
    for line in csv_read:
        sp_data.append(line[0:3])
    sp_data = sp_data[1:]
    for cell in sp_data:
        cell[0], cell[1], cell[2] = int(cell[0]), float(cell[1]), float(cell[2])
    return sp_data

#This function finds the max and min x and y ROI boundaries in the data file.
#The data file is modified so that the distance of each cell from these boundaries is recorded
#in positions cell[3] - cell[6].
def boundaries(sp_data):
    xmin, ymin, xmax, ymax = sp_data[0][1], sp_data[0][2], sp_data[0][1], sp_data[0][2]
    for cell in sp_data:
        if cell[1] < xmin:
            xmin = cell[1]
        if cell[1] > xmax:
            xmax = cell[1]
        if cell[2] < ymin:
            ymin = cell[2]
        if cell[2] > ymax:
            ymax = cell[2]
    for cell in sp_data:
        xmin_dist = abs(cell[1] - xmin)
        xmax_dist = abs(xmax - cell[1])
        ymin_dist = abs(cell[2] - ymin)
        ymax_dist = abs(ymax - cell[2])
        cell.append(xmin_dist)
        cell.append(xmax_dist)
        cell.append(ymin_dist)
        cell.append(ymax_dist)
    return sp_data, xmin, xmax, ymin, ymax

#Generate clustering values
#This function is the main speed culprit when launched from Aptana or Eclipse/PyDev.
#The "for cell/for compare_cell" loop typically runs ~1-2 million times per function call
#with a data input length of roughly 2500-3000 cells.
#On my PC, with input case B4925 and celltypes set to 1/3:
#Aptana and Eclipse loop in ~.5 sec, while IDLE loops in 6-7 sec.
#Between loops: Aptana/Eclipse <.0001 sec, IDLE: 6-7 sec.
#Given that I will loop through this 1200-6000 times/case doing various calcs, that's huge!
def cluster(sp_data, cell1, cell2):
    print "cluster in: " + str(time.clock())
    data_cluster = []
    for unused in range(0, analysis_dist + 1):
        data_cluster.append(0.)
    for cell in sp_data:
        if cell[0] == cell1:
            if cell[3] > exclude_dist and cell[4] > exclude_dist and cell[5] > exclude_dist and cell[6] > exclude_dist:
                #setting these variables here shaves ~7-8% off runtime
                xloc = cell[1]
                yloc = cell[2]
                for compare_cell in sp_data:
                    if compare_cell[0] == cell2:
                        dist = math.sqrt((xloc - compare_cell[1])**2 + (yloc - compare_cell[2])**2)
                        if dist > 0 and dist <= analysis_dist:
                            array_target = int(math.ceil(dist * analysis_dist / interval_num))
                            for insert in range (array_target, analysis_dist + 1):
                                data_cluster[insert] += 1
    print "cluster out: " + str(time.clock())
    return data_cluster

#Average together the results of the two runs (one from the "perspective" of each cell type)
def cluster_average(cluster1, cluster2):
    for interval in range(0, analysis_dist+1):
        cluster1[interval] = (float(cluster1[interval]) + float(cluster2[interval]))/2
    return cluster1

#Make a simulated version of the cell distribution with random locations
def simulation_gen(sp_data, xmin, xmax, ymin, ymax):
    sim_data = []
    for cell in sp_data:
        sim_data.append([cell[0], random.uniform(xmin, xmax), random.uniform(ymin, ymax), cell[3]])
    return sim_data

#Modified version of boundaries function so as not to reset boundaries smaller in simulation runs
#There is probably a better way to refactor all of the boundaries functions
def simulation_boundaries(sim_data, xmin, xmax, ymin, ymax):
    for cell in sim_data:
        xmin_dist = abs(cell[1] - xmin)
        xmax_dist = abs(xmax - cell[1])
        ymin_dist = abs(cell[2] - ymin)
        ymax_dist = abs(ymax - cell[2])
        cell.append(xmin_dist)
        cell.append(xmax_dist)
        cell.append(ymin_dist)
        cell.append(ymax_dist)
    return sim_data

#This is the main function that runs simulations of cellular location
def simulation_iterate(sim_run_num, sp_data_mod, cell1, cell2, xmin, xmax, ymin, ymax):
    simulation_track = []
    for unused in range(0, analysis_dist + 1):
        simulation_track.append(0)
    for run_count in range(0, sim_run_num):
        print "simulation run " + str(run_count + 1)
        simulation_raw = simulation_boundaries(simulation_gen(sp_data_mod, xmin, xmax, ymin, ymax), xmin, xmax, ymin, ymax)
        if cell1 == cell2:
            simulation_cluster = cluster(simulation_raw, cell1, cell1)
        else:
            simulation_cluster = cluster_average(cluster(simulation_raw, cell1, cell2), cluster(simulation_raw, cell2, cell1))
        for location in range(0,analysis_dist + 1):
            simulation_track[location] = simulation_track[location] + simulation_cluster[location]
    for location in range(0, analysis_dist + 1):
        simulation_track[location] = simulation_track[location] / sim_run_num
    return simulation_track

#Use simulation output to correct density-correct clustering data
def simulation_correct(data_cluster, simulation_cluster):
    corrected_output = []
    for unused in range(0, analysis_dist + 1):
        corrected_output.append(0)
    for location in range(0, analysis_dist + 1):
        try:
            corrected_output[location] = data_cluster[location] / simulation_cluster[location]
        except:
            pass
    return corrected_output

sp_data = loadfile()
sp_data_mod, xmin, xmax, ymin, ymax = boundaries(sp_data)

print "data cluster run"
if cell1 == cell2:
    data_cluster = cluster(sp_data_mod, cell1, cell1)
else:
    data_cluster = cluster_average(cluster(sp_data_mod, cell1, cell2), cluster(sp_data_mod, cell2, cell1))
print "raw clustering value: "
print data_cluster

simulation_cluster = simulation_iterate(sim_run_num, sp_data_mod, cell1, cell2, xmin, xmax, ymin, ymax)
print "simulation clustering value:"
print simulation_cluster

sp_output = simulation_correct(data_cluster, simulation_cluster)
print "output clustering value: "
print sp_output

print "run time: " + str(time.clock())

#set up worksheet to write to
book = xlwt.Workbook(encoding="utf-8")
sheet1 = book.add_sheet("Python Sheet 1")

##populate excel worksheet
for location in range(0, analysis_dist + 1):
    sheet1.write(0, location, (str(location) + " um"))
    sheet1.write(1, location, sp_output[location])

#save the spreadsheet
savepath = directory + "\\" + outputfile + ".xls"
book.save(savepath)

