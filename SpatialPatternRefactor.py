# This is a refactored Python version of my spatial pattern analysis program, 
# published in Morgan et al., 2012.  The program examines the distribution of 
# cellular populations in a region of cortex or other layered tissue to 
# determine whether they are randomly spaced or instead more or less closely 
# clustered together than expected.  For a homogeneous, non-layered sample, 
# please see SpatialPattern_NoLayer.py.

# Summary of method:
# This program looks at the recorded x, y-coordinates of each cell in a region 
# of interest.  If the cell is far enough from the boundaries of the ROI that 
# we can investigate the full distance range around it without running into 
# ROI edge effects, the program calculates the distance to all of the other 
# cells of the appropriate cell class in the sample.  It then marks each 
# matching instance at the distance between the cells and all of the farther 
# distances within the analysis range.

# The program then generates simulations of cellular distribution, which have 
# randomized cell locations *by layer*.  It compares an average of these 
# simulation values to the actual distribution of cells to detect 
# inhomogeneities in their organization.

##############################################################################
# Section 1: setting up your computer to run this program
# This program is written in Python 2.6-2.7.  It also uses the xlwt addon 
# library to make Excel spreadsheets.

# Step 1: Install Python 2.7.3 from here: 
# http://www.python.org/download/releases/2.7.3/
# On the page, select the Windows MSI installer (x86 if you have 32-bit 
# Windows installed, x86-64 if you have 64-bit Windows installed.)
# I suggest using the default option, which will install Python to c:/Python27

# Step 2: Install the xlwt library from here: 
# http://pypi.python.org/pypi/xlwt/
# Use the program WinRAR to unzip the files to a directory
# Go to "run" in the start menu and type cmd
# Type cd c:\directory_where_xlwt_was_unzipped_to
# Type setup.py install

# Step 3: Copy this program into the c:/Python27 directory
# You can also put it into any directory that is added to the correct PATH.

##############################################################################
# Section 2: file preparation for this program (and related scripts)

# Record coordinates of all cells belonging to the populations of interest in 
# a rectangular counting frame in Stereo Investigator.  Take the raw cellular 
# coordinates and save as a .txt file. Open the text file.  Leave the header 
# in.  Delete the information in columns 4 and 5.  Add layer information for 
# each cell in column 4.  (Do a y-coordinate sort on the data, then fill in 
# the column with reference to the Stereo Investigator file.)  Save, set your 
# variables in the section below, and run the program.

##############################################################################
# User set variables here

# Directory containing input files
directory = r"C:\Users\John Morgan\My Documents\sp_datafiles"

# Name of cell coordinate data file from Stereo Investigator. 
# There should be a header, leave it in.
# Delete z-axis data and input layer information for each cell.
# Save as a tab-delimited text file.
inputfile = r"B4925.txt"

# This is the name of the .xls file that the program will save when finished. 
outputfile = "test"

# This variable adjusts the cell types that are being compared.  Input the 
# values in the first column of your saved .txt file.  In the demo run, 
# neuron = 1, microglia = 3.  To look at the properties of a single 
# population, set both of these values the same.
cell1 = 1
cell2 = 3

# Number of cortical layers, this is used to randomize cell distributions by 
# layer
layer_num = 6

# This variable sets distance from the ROI boundary at which seed cells will 
# be excluded from analysis to avoid edge effects
exclude_dist = 100

# This variable adjusts the studied distance and interval in the cleaned 
# output file.  Do not exceed the excluded distance or you will have edge 
# effects distorting your results!
analysis_dist = 100

# This variable adjusts the number of interrogation points over the analysis 
# range.  It is strongly suggested to match analysis_dist unless you are 
# attempting to reduce noise with larger distance bins.
interval_num = 100

# The number of simulations to run; the results are averaged and then the 
# clustering values are divided by them to produce a clustering ratio.
# In Morgan et al., 2012, I ran 200 simulations/condition.  Experimentation 
# indicated that this number of simulations resulted in very low variability 
# in results.  1000 simulations may be ideal if you have the time/power.
sim_run_num = 5

##############################################################################
# Program begins here

# Import modules to handle a tab-delimited text file and produce .xls output
import csv
import math
import random
import time

import xlwt


# Load and cleanup file. Output is sp_data, which contains all cells. 
# Output is [[celltype1, xcoord1, ycoord1],[celltype2, xcoord2, ycoord2], etc]
def loadfile():
    path = directory + "\\" + inputfile
    myfileobj = open(path,"r") 
    csv_read = csv.reader(myfileobj,dialect = csv.excel_tab)
    sp_data = []
    for line in csv_read:
        sp_data.append(line[0:4])
    sp_data = sp_data[1:]
    for cell in sp_data:
        cell[0], cell[1] = int(cell[0]), float(cell[1]) 
        cell[2], cell[3] = float(cell[2]), int(cell[3])
    return sp_data


# This function finds the max and min x and y ROI boundaries in the data file.
# The data file is modified so that the distance of each cell from these 
# boundaries is recorded in positions cell[4] - cell[7].
def boundaries(sp_data):
    (xmin, ymin, xmax, ymax) = (sp_data[0][1], sp_data[0][2], 
                                sp_data[0][1], sp_data[0][2])
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


# This function sets boundaries by layer so that when random cell location 
# simulations are generated, they are performed by layer.  This is necessary 
# because cell density varies by layer.
def layer_ybound(sp_data_mod, ymin, ymax):
    ybound_list = []
    for layer in xrange(layer_num):
        layer_list = []
        for cell in sp_data_mod:
            if cell[3] == layer + 1:
                layer_list.append(cell[2])
        layer_max = layer_list[0]
        layer_min = layer_list[0]
        for layer_cell in layer_list:
            if layer_cell > layer_max:
                layer_max = layer_cell
            if layer_cell < layer_min:
                layer_min = layer_cell
        ybound_list.append([layer_max, layer_min, layer + 1])
    # Set layer boundaries by averaging the min from one layer with the max 
    # from the next layer.
    for layer in xrange(layer_num - 1):
        try:
            layer_bound = (ybound_list[layer][1] + 
                           ybound_list[layer + 1][0]) / 2
            (ybound_list[layer][1], ybound_list[layer + 1][0]) = (layer_bound, 
                                                                  layer_bound)
            ybound_list[0][0], ybound_list[0][layer_num - 1] = ymax, ymin
        except:
            pass     
    return ybound_list


# Generate clustering values.
# This function is the main time sink of the program right now. 
def cluster(sp_data, cell1, cell2):
    print "cluster in: " + str(time.clock())
    raw_cluster = [0.] * (analysis_dist)
    for cell in sp_data:
        if cell[0] == cell1:
            if (cell[4] > exclude_dist and cell[5] > exclude_dist and 
                cell[6] > exclude_dist and cell[7] > exclude_dist):
                # Setting these variables here shaves ~7-8% off runtime
                xloc = cell[1]
                yloc = cell[2]
                for compare_cell in sp_data:
                    if compare_cell[0] == cell2:
                        dist = math.sqrt((xloc - compare_cell[1])**2 + 
                                         (yloc - compare_cell[2])**2)
                        if dist > 0 and dist < analysis_dist:
                            array_target = int(math.ceil(dist * 
                                                         (analysis_dist - 1) / 
                                                         interval_num))
                            for insert in range (array_target, analysis_dist):
                                raw_cluster[insert] += 1
    print "cluster out: " + str(time.clock())
    return raw_cluster


# Average together the results of the two runs (one from the "perspective" of 
# each cell type).
def cluster_average(cluster1, cluster2):
    for interval in xrange(analysis_dist):
        cluster1[interval] = (float(cluster1[interval]) + 
                              float(cluster2[interval]))/2
    return cluster1


# Make a simulated version of the cell distribution with random locations
def sim_gen(sp_data, xmin, xmax, ybound_list):
    sim_data = []
    for cell in sp_data:
        yrand = random.uniform(ybound_list[cell[3]-1][0], 
                               ybound_list[cell[3]-1][1])
        sim_data.append([cell[0], random.uniform(xmin, xmax), yrand, 
                          cell[3]])
    return sim_data


# Modified version of boundaries function so as not to reset boundaries smaller 
# in simulation runs
def sim_boundaries(sim_data, xmin, xmax, ymin, ymax):
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


# This is the main function that runs simulations of cellular location
def sim_iterate(sim_run_num, sp_data_mod, cell1, cell2, xmin, xmax, ymin, 
                ymax, ybound_list):
    sim_track = [0.] * (analysis_dist) 
    for runcount in xrange(sim_run_num):
        print runcount + 1
        sim_raw = sim_boundaries(sim_gen(sp_data_mod, xmin, xmax, 
                                         ybound_list), xmin, xmax, 
                                                       ymin, ymax)
        if cell1 == cell2:
            sim_cluster = cluster(sim_raw, cell1, cell1)
        else:
            sim_cluster = cluster_average(cluster(sim_raw, cell1, cell2), 
                                          cluster(sim_raw, cell2, cell1))

        for location in xrange(analysis_dist):
            sim_track[location] = (sim_track[location] + 
                                       sim_cluster[location])
    for location in xrange(analysis_dist):
        sim_track[location] = sim_track[location] / sim_run_num
    return sim_track


# Use simulation output to density-correct clustering data
def sim_correct(raw_cluster, sim_cluster):
    corrected_output = [0.] * (analysis_dist)
    for location in xrange(analysis_dist):
        try:
            corrected_output[location] = (raw_cluster[location] / 
                                          sim_cluster[location])
        except:
            pass
    return corrected_output


print time.clock()
# Add an extra column to analysis distance so program runs from zero to 
# analysis distance, *inclusive*.
analysis_dist += 1 

sp_data = loadfile()
sp_data_mod, xmin, xmax, ymin, ymax = boundaries(sp_data)

if cell1 == cell2:
    raw_cluster = cluster(sp_data_mod, cell1, cell1)
else:
    raw_cluster = cluster_average(cluster(sp_data_mod, cell1, cell2), 
                                  cluster(sp_data_mod, cell2, cell1))
print "raw clustering value: "
print raw_cluster

ybound_list = layer_ybound(sp_data_mod, ymin, ymax)
sim_cluster = sim_iterate(sim_run_num, sp_data_mod, cell1, cell2, 
                          xmin, xmax, ymin, ymax, ybound_list)
print "simulation clustering value:"
print sim_cluster

sp_output = sim_correct(raw_cluster, sim_cluster)
print "output clustering value: "
print sp_output

print "run time: " + str(time.clock())

# Set up worksheet to write to
book = xlwt.Workbook(encoding="utf-8")
sheet1 = book.add_sheet("Python Sheet 1")

# Populate excel worksheet with headers and results
for location in xrange(analysis_dist):
    sheet1.write(0, location, (str(location) + " um"))
    sheet1.write(1, location, sp_output[location])

# Save the spreadsheet
savepath = directory + "\\" + outputfile + ".xls"
book.save(savepath)