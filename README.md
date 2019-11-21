# Google-Sheets-MC-Simulator
Google Apps Script utility to run Monte Carlo simulations

This script draws MC simulations from a google sheet. Variables are set up as random variables in the spreadsheet, and calculations laid out just as usual in the spreadsheet.

The script recalculates the sheet, thereby drawing new random numbers, and collects the results of specified cells.

|| A        | B           | C  |
|----| ------------- |:-------------:| -------:|
|1| 14.34323      | 123.3423 | 1245.3 |
|2| 12.34432  |       |    |
|3| Cells to evaluate |      |   |
|4| **A1** |     |   |
|5| **A2** |      |   |

e.g. select cells A4 to A5, Addons -> MCsim -> Calculate MC Simulations. A new spreadsheet containing n_mc_smps random estimates of the selected cells is created. 
n_mc_smps is hardcoded at the moment.

This tool allows usage of standard spreadsheet calculations, at the cost of a terrible runtime.
