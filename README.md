Excel-Monte-Carlo-Simulation-and-Optimisation
=============================================

Excel Monte Carlo Simulation and Optimisation

Background
This is a simple application that varies particular values of a spreadsheet according to a distribution and report on dependent cells. In essence, you are flexing a static spreadsheet model. The distributions, currently, are equal probability, Normal distribution, or series distribution. The result is a new separate spreadsheet with each iteration of the model on a single line. This data can then, of course, be used for graphing etc.

Usage
(a) import xlm macro into spreadsheet or your workspace
(b) in the target sheet, for cells which are variable, include a comment against the cell as follows

=SeriesFunc(RC[-1],1,0.5,11,0.6,21,0.7,31,0.8,41,0.9)



or 
=EqualProbFunc(RC[-1],-0.5, +0.5)

Here 50% above, or 50% below cell value

or
=NormFunc(RC[-1])

Normal distribution about cell value.

(c) for cells which are dependent variables or need to be reported, colour them red
The application will try to use the cell to the immediate left as the label

(d) change the number of iterations of the simulation

(e) run the macro, which will produce another sheet with the name "Flexing Sheet1" or whatever the target sheet was called
All the cells which are coloured red in step (c) report for each simulation. Each column is a reported variable. 



