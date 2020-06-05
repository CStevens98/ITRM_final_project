# Author: Chiara Stevens
# This is a python program for reading a cell in an excel file

import pandas as pd
import openpyxl


def main():

    xl = pd.ExcelFile('Sex-disaggregated COVID-19 data tracker.xlsx')
    df = xl.parse(0)

    # Creating variables specific to Italy
    # The first [] is the column and the [1] is the row number for Italy
    total_cases = (df['Total cases'][1])
    date = (df['Date'][1])
    total_deaths = (df['Total deaths'][1])
    deaths_male_perc = (df['deaths (% male)'][1])
    deaths_female_perc = (df['deaths (% female)'][1])
    total_male_deaths = (int(deaths_male_perc) / 100) * int(total_deaths)
    total_female_deaths = (int(deaths_female_perc) / 100) * int(total_deaths)

    print("The total number of cases in Italy is {0} as of {1}.".format(total_cases, date))
    print("With a total number of {0} deaths.".format(total_deaths))
    print("Where {0} is the total deaths of males and {1} is the total death of females.".format(total_male_deaths, total_female_deaths))


main()
