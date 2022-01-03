'''
ExcelPeek v0.2 - Copyright 2022 James Slaughter,
This file is part of ExcelPeek v0.2.

ExcelPeek v0.2 is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.

ExcelPeek v0.2 is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License for more details.

You should have received a copy of the GNU General Public License
along with ExcelPeek v0.2.  If not, see <http://www.gnu.org/licenses/>.
'''

'''
controller.py - This file is responsible for keeping global settings available through class properties
'''

#python imports
import imp
import sys

'''
controller
Class: This class is is responsible for keeping global settings available through class properties
'''
class controller:
    '''
    Constructor
    '''
    def __init__(self):

        self.debug = False#Boolean value, if set to true, debug lines will print to the console.
        self.file = ''#input from the --file cmd line flag denoting target file to investigate.
        self.list = False#Boolean value, if set to true, will list all worksheets in the target workbook.
        self.sheet = ''##input from the --sheet cmd line flag denoting which desired worksheet to investigate.
        self.sheetstats = False#Boolean value, if set to true, will provide info on the desired worksheet
        self.dump = False#Boolean value, if set to true, will dump all non-empty cells in range for the desired worksheet to output
        self.analyze = False#Boolean value, if set to true, will cycle through all sheets and output the particulars
        self.wb = ''#Object for the workbook being investigated.
        self.range = ''#input from the --range cmd line flag denoting the range of cells that can be read.
        self.output = ''#input from the --output cmd line flag denoting the desired location and name of the output file.
