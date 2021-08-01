# VBA projects
This repo contains 1 project and 3 mini-projects of slightly shorter length. The project ideas were adopted from the Excel/VBA specialization offered in Coursera. Each project consists of one .xlsm file (Excel Macro-enabled notebook). Note that all projects have been designed to work only with Windows versions of Excel, due to major differences in the way Windows and MacOS versions of Excel implement the GetOpenFileName method.

## Grade Manager project
This project implements a setup tool that consolidates and combines student grades from multiple class section roster grade books into one master workbook. The master workbook, once synchronized with section files, will contain each student's grade from every assignment. The number of assignments are determined during initialization and can be changed subsequently. The project assumes that the class roster (number of students) does not change from initialization, and that section files are always provided in the same pre-determined format.

Features:
* Initializes a new directory containing a master grades workbook (given a fixed roster), and section files containing student grades from each section in a consistent format
* Ability to set a default number of each type of assignment (homework, exam, lab) during initialization
* Ability to add/delete assignments
* Ability to create dated backups for the master workbook
* Ability to synchronize the master workbook with section files at any time
* Ability to search for or replace a chosen student's grade for a chosen assignment
* Input validation

## Currency Converter
This mini-project implements a user form that will enable the use of real-time exchange rates to convert currency from one unit to another.

Features:
* Queries exchange rate data from xe.com/currencytables/ to convert currency value. Default data used is data from two days before the current date
* Ability to change the date to convert currency value using any historical rate
* Ability to generate a plot of exchange rates over the past 30 days
* Input validation

Note that depending on system settings, the program may not work perfectly on certain devices.

## Regression Toolbox
This mini-project implements a user form that allows the user to create a custom linear regression model using a user-hypothesized form of the model.

Features:
* Accepts up to 4 custom univariate functions
* Fits x-y data provided on the spreadsheet to the chosen model and outputs model parameters and the adjusted R-squared value
* Option to plot experimental data (as markers, no line) and model predictions (no markers, solid smooth line) on the same plot.
* Input validation

## Profitability Analysis
This mini-project implements a user form that allows the user to simulate a profitability analysis based on net present value (NPV) of a given proposed capital project. It assumes a fixed set of input variables with fixed distribution types, but the distribution parameters can be customized.

Features:
* Ability to input custom distribution parameters (eg. mean and standard deviation for a normally distributed input variable)
* Ability to change the number of simulations
* Outputs percentage of simulations that are "profitable" (NPV > 0), and outputs results of the simulation in a histogram
