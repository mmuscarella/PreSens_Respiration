################################################################################
#                                                                              #
#  Test Script for Analysis of PreSens Respiration Data                        #
#   Version 2.0                                                                #
#  Written By: Mario Muscarella                                                #
#  Last Update: 28 Jan 2014                                                    #
#                                                                              #
#  Use this file to test the PreSens.Respiration fucntion on local machine     #
#  And as a template to create your own analysis                               #
#                                                                              #
################################################################################

setwd('C:/Users/Mario Muscarella/Documents/PhD Files/Projects/2013-PreSens/Analysis Script/')
rm(list=ls())

# Inport the function from source file
source("PreSensInteractiveRegression.r")

################################################################################
# Examples ##################################################################### 
################################################################################

# Example txt analysis
PreSens.Respiration("Example.txt", "Example.Out.txt")

# Example excel analysis (with data in column format)
PreSens.Respiration("Example.xlsx", "Example.Out.txt", "Columns")

#
# Each of the above examples should open 
