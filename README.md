# OctopusAnalysis
Octopus smart meter export analysis.

Export the smart meter data from your octopus account.

From my account I get csv files for electricity or gas.  The csv files have column headings as follows:

Electricity export: "Consumption (kWh)" " Start" " End" note the leading whitespace for the Start and End column headings.
Gas export: "Consumption (mÂ³)" " Start" " End" note the leading whitespace for the Start and End column headings.

The date/time format in the csv files appears to be eg: 2022-09-03T00:30:00+01:00

At the start of the code under the if __name__ == '__main__': declaration you will find variable for the expected column headers and time format.
You will also find the gas_calorific value used to convert the gas use in to KWh.

Probably this script could work with data from any energy company smart meter export, but the file format might be different.  You may be able to make it work by setting the initial variables to the new format.  It could of course be possible to create settings for various energy companies and the automatically detect which format is in use, or use a commandline option.  I don't have files from any other energy company export so I'm not able to do this.

The script looks for and removes day data for any day which does not have 24-hours of recorded use (for zero use the file normaly contains zeros for the relevant time periods.  This is done because missing day/hour data can skew the minimum and mean hourly/daily calculations.  If you don't want them deleted run the script with the -n or --nodelete option and they will be reported but remain in the analysis.  I found quite a lot of gas use data was missing from my export, and a small amount of electricity data.  No idea why that happened.

Running the script will produce png image saves of a couple of graphs and create a docx file with those images inserted and a couple of tables of data.  The docx file is quite basic.  

Overall the script provides some useful output to help you see your energy use.  Compare winter/summer differences, this may allow you to work out how much energy you use for hot water or heating (if you have gas then the summer use will mostly/entirely be for hot water).  You can also see how much you use overnight/various times of day which could help you work out what size solar panels you need, if a battery storage is good for you and what size you may need etc.  Seeing the peak gas heating use may help you workout what size gas boiler you really need, or perhaps what size heatpump you may need etc.



