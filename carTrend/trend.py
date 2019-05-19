import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import numpy as np 
import pandas as pd
import pandas.plotting._converter as pandacnv

def pricing_plot(date, pricing, raw_input):

    fig, ax = plt.subplots()
    ax.scatter(date, pricing)

    # x-axis formatting to show major months with minor ticks being individual days
    ax.xaxis.set_major_locator(mdates.MonthLocator(interval=1))
    ax.xaxis.set_minor_locator(mdates.DayLocator())
    ax.xaxis.set_major_formatter(mdates.DateFormatter('%m-%Y'))
    fig.autofmt_xdate()

    # Formatting to give plot information
    plot_title = 'Craigslist Listings of ' + raw_input.upper()  + ' Over Time'

    ax.set(title = plot_title)
    ax.set(xlabel = "Date", ylabel = "Pricing ($)")
    plt.xticks(rotation=45)
    ax.grid(True)
    plt.show()
    

if __name__ == "__main__":
    # Takes user input of vehicle
    raw_input = input('What vehicle? ')
    file_name = raw_input.replace(' ', '_') + '.xlsx'
    
    # Pandas doesn't import all converters,  so
    pandacnv.register()

    # Use Pandas to read and sort xlsx file of scrapped listings
    df = pd.read_excel(file_name)
    df = df.sort_values('Date')

    # Convert dates to work with matplotlib plotting
    df['Date'] = pd.to_datetime(df['Date'])
    df['Date'] = df['Date'].tolist()

    # Pass data to be plotted in scatter plot
    pricing_plot(df['Date'], df['Listed Price'], raw_input)
    
    
