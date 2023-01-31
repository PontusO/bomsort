#!/usr/bin/env python3
"""
Module Docstring
"""
import sys                        # https://docs.python.org/3/library/sys.html
import argparse                   # https://docs.python.org/3/library/argparse.html
import csv                        # https://docs.python.org/3/library/csv.html
import xlsxwriter                 # https://xlsxwriter.readthedocs.io/
import copy                       # https://docs.python.org/3/library/copy.html
from natsort import natsorted, ns # https://natsort.readthedocs.io/en/master/
from operator import itemgetter   # https://docs.python.org/3/library/operator.html

__author__ = "Pontus Oldberg"
__version__ = "0.1.0"
__license__ = "MIT"

# Indexes used in the main BOM
BOM_DES = 0   # Designator
BOM_X   = 1   # X Value
BOM_Y   = 2   # Y Value
BOM_A   = 3   # A (rotation) Value
BOM_CMB = 4   # This is for the combined "value" and "package" fields in the BOM.
              # This field is used in the internal main BOM.
BOM_VAL = 4   # Component value (Used only in the original file)
BOM_PAC = 5   # Component package (Used only in the original file)

# Indexes used in parts list (created by create_component_list)
PART_DES = 0  # Designator of the first instance of the part in this entry.
PART_X   = 1  # X Value
PART_Y   = 2  # Y Value
PART_A   = 3  # A (rotation) Value
PART_CMB = 4  # This is for the combined "value" and "package" fields in the BOM.
PART_QTY = 5  # The qty of this specific part
PART_FDR = 6  # The feeder number for this part.

# Indexes for splitted combination field
IX_VAL  = 0   # Component value
IX_PAC  = 1   # Component package

#
# This function takes the full complete BOM list and returns a list of unique parts with
# an added quantity field.
def create_component_list(bomlist):
  # Create an empty list where we place the counted components.
  counted_parts = []
  # Now go through all the components in the supplied BOM
  for part in bomlist:
    # Assume that the current part should be insterted into the count list
    insert_flag = True
    # Go through the list of counted parts
    for counted_part in counted_parts:
      # Check if the current part in the bomlist is already added.
      if counted_part[PART_CMB] == part[BOM_CMB]:
        # If so, indicated that it should not be added.
        insert_flag = False
        # Instead just increase the count
        counted_part[PART_QTY] = counted_part[PART_QTY] + 1
        break
    if insert_flag:
      # Creata a copy of the list item to avoid refencing the original list
      item = list(part)
      # Add the QTY field to the entry.
      item.append(1)
      # And finally add it to the list of counted items.
      counted_parts.append(item)
  return counted_parts

# Optimization function for constructed parts lists.
# Optimization primarily exists to fix multiple parts being sorted together after
# having been added to the list. This function adheres to the feeder table
# constraints (Positions 0-37).
# bomlist is passed by assignment (in this case by reference) and is operated on
# directly by this function. If a change was made to the parts list the function
# returns true so that the caller can choose to call it again should it be
# necessary.
def optimize_part_list(bomlist, cnt):
  updated = False
  for i in range(cnt-1):
    if bomlist[i][4] == bomlist[i+1][4]:
      updated = True
      if i <= 17:   # Lower part of feeder table, try to exchange [i] part with the next feeder down.
        if i == 0:  # Special case if we're looking at the first positon, then we need to exchange [i+1] with next feeder up.
          tmp = bomlist[i+2]
          bomlist[i+2] = bomlist[i+1]
          bomlist[i+1] = tmp
        else:       # Normal case on the lower part of the feeder table
          tmp = bomlist[i-1]
          bomlist[i-1] = bomlist[i]
          bomlist[i] = tmp
      else:
        if i == 36: # Special case if we are looking at the last position, then we need to exchange [i] with next feeder down
          tmp = bomlist[i-1]
          bomlist[i-1] = bomlist[i]
          bomlist[i] = tmp
        else:       # Normal case on the upper part of the feeder table
          tmp = bomlist[i+2]
          bomlist[i+2] = bomlist[i+1]
          bomlist[i+1] = tmp
  return updated

#
# The main part of the show
# This is where program execution starts after it has been started from the command line.
#
# This program was created for use with the SMT production line at Invector Labs.
#
# It can be used to create a component list that can be imported into the SMT800 pick and
# place machines. It can also generate a list of used components for the project (inventory
# pick list) to simplify the preparation. It can also produce a suggested sort order for
# the pick and place machines.
#
# The algorithm for suggesting the sort order in the pick and place machine currently only
# supports one machine and the front row.
#
def main():
  print("BOM Manager utility V0.2")
  print("(C) 2019-2022 Invector Embedded Systems AB, Written by Pontus Oldberg")
  
  # The argument parser is used to read and interpret the options given on a command line.
  #
  # For instance if we want to create a suggested sort order for a particular project we
  # type the following in a command shell:
  #
  #   bomsort.py bom.txt -f bom.feeders
  #
  # The first part (bomsort.py) makes sure that this program is called. The first argument
  # (bom.txt) is the original BOM file generated by the CAD software. 
  # The second part (-f bom.feeders) is an option that tells this program that we want to
  # create a feederlist file called "bom.feeders". This file will contain the sorted list
  # that can be used by the pick and place operator to determine what components to put
  # on a particular feeder position.
  #
  parser = argparse.ArgumentParser() 
  parser.add_argument('-b', '--bom', action='store')
  parser.add_argument('-s', '--sorted', action='store_true')
  parser.add_argument('-p', '--parts', action='store')
  parser.add_argument('-f', '--feederlist', action='store')
  parser.add_argument('Input File', nargs='*') # use '+' for 1 or more args (instead of 0 or more)
  parsed = parser.parse_args()
  
  # This just makes sure you have enough arguments given on the command line.
  if len(sys.argv) < 2:
    print("You have to specify a file name !")
    exit(1)

  # Create an empty list that will hold the main BOM.    
  v = []
  try:
    # Get the filename from the input arguments.
    filename = sys.argv[1]
    # Open the file
    f = open(filename, "r")
    # Read the file line by line and parse the input as we go through the file
    for line in f:
      # Do not include test points or fiducialshere
      if not (line.startswith("TP") or line.startswith("FID")):
        # Sometimes CAD libraries use commas in component values and/or descriptions. This does
        # not work well when creating and exporting comma separated files. So here we simply
        # replace any detected commas with decimal points.
        noComma = line.replace(",", ".")
        # Split the current line as well as strip it from any leading or ending white spaces
        ml = noComma.strip().split()
        # Create a temporary empty list
        nl = []
        # Add items from the file to our internal BOM list.
        nl.append(ml[BOM_DES])  # Designator (U1, C2 etc)
        nl.append(ml[BOM_X])  # X value
        nl.append(ml[BOM_Y])  # Y value
        nl.append(ml[BOM_A])  # A (rotation) value
        # Merge the value and description to create a unique component identifer
        # If we did not do this and just used the component value there would be a conflict if
        # we had for instance a 0402 0.1uF and a 0603 0.1uF on the same board. 
        nl.append(ml[BOM_VAL] + "|" + ml[BOM_PAC])
        # Add this entry to the main BOM.
        v.append(nl)
    f.close()
  except  Exception as e: 
    print("Could not open file %s" % filename)
    print(e)

  #
  # Generate a BOM if requested
  if parsed.bom:
    print("Generating parts list for the KAYO pick and place machine !")
    # If the user requested the list to be sorted we do this before generating the part list.
    if parsed.sorted is True:
      # Do a natural sort of the main BOM and copy the result to vs
      vs = natsorted(v, key=lambda bom: bom[BOM_DES])
    else:
      # No sorting requested, just copy the original main BOM
      vs = v
    
    # Create a comma separated (CSV) text file
    with open(parsed.bom, 'w') as csvfile:
      # Create CSV field names for this file
      writer = csv.DictWriter(csvfile, fieldnames=['Part', 'X', 'Y', 'A', 'Description', 'Package'])
      writer.writerow({'Part': 'Part', 'X': 'X', 'Y': 'Y', 'A': 'A', 'Description': 'Description', 'Package': 'Package'})
      # Write the data to the CSV file
      for row in vs:
        # Split the values again that we combined initially.
        info = row[BOM_CMB].split("|")
        writer.writerow({'Part': row[BOM_DES], 'X': row[BOM_X], 'Y': row[BOM_Y], 'A': row[BOM_A], 'Description': info[IX_VAL], 'Package': info[IX_PAC]})
    
  #
  # Generate a list of used components if requested
  if parsed.parts:
    print("Generating a list of used parts (Inventory picking list) !")
    vs = create_component_list(v)
    cl = sorted(vs, key=itemgetter(BOM_DES), reverse=False)

    # Create an excel file 
    row = 1
    with open(parsed.parts, 'w') as csvfile:
      # Create CSV field names for this file
      writer = csv.DictWriter(csvfile, fieldnames=['Part', 'Description', 'Package', 'Count'])
      writer.writerow({'Part': 'Part', 'Description': 'Description', 'Package': 'Package', 'Count': 'Count'})
      # Write the data to the CSV file
      for i in cl:
        info = i[4].split("|")
        writer.writerow({'Part': row, 'Description': info[IX_VAL], 'Package': info[IX_PAC], 'Count': i[PART_QTY]})
        row = row + 1
        
  # Generating a suggested feeder list
  if parsed.feederlist:
    print("Generating feeder list !")

    # Feeder width table
    # This table defines the different available feeder widths and how they are
    # distributed on the feeder table. Each entry in this list of lists contain
    # the following information:
    #   * The width of the feeder, in mm (Example 8, 12, 16 etc)
    #   * How many 8mm positions that is occupied above the base feeder location.
    #     For instance, an 8mm feeder does not occupy any feeder locations above
    #     the base position. In this case another 8mm feeder can be positioned
    #     directly above. A 12mm feeder always occupy the next location as well
    #     as the base postion.
    #   * How many 8mm positions that the feeder occupy below the base feeder 
    #     location.
    #     Like the parameter above but below the base location.
    # For instance, an 8mm feeder would have the following defintion:
    #   [8, 0, 0]
    # and a 12mm feeder would look like this:
    #   [12, 1, 0]
    #
    feeder_widths = [[8, 0, 0],[12, 1, 0],[16, 1, 0]]

    # Feeder Distribution table
    # This table is used to translate from the (reverse) sorted list to the feeder
    # table in the pick and place machine. For instance the 0'th item in the part
    # list (after it has been reverse sorted) corresponds to the 19'th position on
    # the feeder table. 
    l1_ptrn = [19,18,20,17,21,16,22,15,23,14,24,13,25,12,26,11,27,10,28,9,29,8,30,7,31,6,32,5,33,4,34,3,35,2,36,1,37,0]

    # Create a list of the components used in the project
    col = create_component_list(v)

    # Determine how many parts to process. We do this to make sure that we are not
    # processing more parts than the number of available feeder slots and if the
    # number of parts in the parts list is less than available feeder slots.
    if len(col) > len(l1_ptrn):
      cnt = len(l1_ptrn)
    else:
      cnt = len(col)

    # Here we do some funky stuff =)
    #
    # First thing we do is check if there are any parts that are over represented.
    # If so we add a position to the list, creating another feeder entry with the
    # same part. This will enable the pnp machine to double pick the same part
    # which can increase mounting speed.
    # Optimization of the feeder list is done later so here we can just add freely.
    #
    # Currently the algorithm checks to see if the number of parts for a specific
    # component is higher than half the total number of components in the project.
    # In this case we add another copy of this component.
    # Currently this is just a guesstimate that this is an approprate ratio and
    # this will likely evolve over time.
    adders = []
    for cmp in col:
      if cmp[PART_QTY] > cnt / 2:
        qty = cmp[PART_QTY]
        cmp[PART_QTY] = int(cmp[PART_QTY] / 2)
        item = list(cmp)
        item[PART_QTY] = qty - cmp[PART_QTY]
        adders.append(item)
    # Now add in any parts that were created.
    for each in adders:
      col.append(each)

    # Create a new list sorted in reverse order based on number of parts per component.
    cl = sorted(col, key=itemgetter(PART_QTY), reverse=True)
    fl = []
    # This is the algorithm for distributing the used parts onto different feeders.

    for i in range(cnt):
      cl[i].append(l1_ptrn[i])     # Append feeder number for this part
      fl.append(cl[i])
    # Sort feeder list (fl) in feeder order.
    fls = sorted(fl, key=itemgetter(PART_FDR))

    # Now we can try to optimize the list.
    while optimize_part_list(fls, cnt):
      print("optimizing......")

    # Create a CSV file to write the sorted feeder list.
    with open(parsed.feederlist, 'w') as csvfile:
      writer = csv.DictWriter(csvfile, fieldnames=['Feeder', 'Description', 'Package', 'Count'])
      writer.writerow({'Feeder': 'Feeder', 'Description': 'Description', 'Package': 'Package', 'Count': 'Count'})
      for entry in fls:
        info = entry[PART_CMB].split("|");
        writer.writerow({'Feeder': entry[PART_FDR]+1, 'Description': info[IX_VAL], 'Package': info[IX_PAC], 'Count': entry[PART_QTY]})

    
if __name__ == "__main__":
  """ This is executed when run from the command line """
  main()
