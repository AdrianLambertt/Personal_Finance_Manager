import pandas as pd
import gspread
from gspread_dataframe import set_with_dataframe


#Data formatting for the dataframe to remove
#NA values and also remove £ (conversion breaks the ASCII..?)
def formatter(df):
    df.fillna("0", inplace=True)
    df.loc[:,"Paid out"] = pd.to_numeric(df.loc[:,"Paid out"].str.replace('�', ''))
    df.loc[:,"Paid in"] = pd.to_numeric(df.loc[:,"Paid in"].str.replace('�', ''))
    df.loc[:,"Balance"] = pd.to_numeric(df.loc[:,"Balance"].str.replace('�', ''))
    return df


#Finds total sum spent per company
def totalPerCompany(df):
    payCompanies = {}

    for index, row in df.iterrows():
        if row["Paid out"] > 0:
            name = row["Description"]
            price = row["Paid out"]
            if name in payCompanies:
                payCompanies[name] += price
            else:
                payCompanies[name] = price
    return payCompanies



#Combines the data pulled from the CSV to create Dataframes ready for the sheets API.
def createGSheets(totalCompany, totalExpense, totalGain, startingBalance, endingBalance):
    companyDF = pd.DataFrame.from_dict(totalCompany, orient='index', columns=['Total Paid out']).reset_index().sort_values(by='Total Paid out', ascending=False)
    companyDF.columns = ['Company', 'Total Paid out']

    infoDF = pd.DataFrame({"Starting Balance" : startingBalance, "Ending Balance": endingBalance, "Total Paid in": totalGain, "Total Paid out": totalExpense}, index=[0])
    return infoDF, companyDF


#verifies and Connects to Google Sheets then updates the Spreadsheet 
def setGSheets(infoDF, companyDF):
    try:
        sa = gspread.service_account()
    except:
        RuntimeError("Service account file unaccesible")

    try:
        sh = sa.open("Finance Manager")
    except:
        RuntimeError("Error connecting to Google Sheets")


    wks = sh.worksheet("Sheet1")
    wks.clear()

    next_row = (len(infoDF.index) +4)
    set_with_dataframe(wks, infoDF, include_index=False, include_column_header=True, row=1, col=1)
    set_with_dataframe(wks, companyDF, include_index=False, include_column_header=True, row=next_row, col=1)


def main():
    df = pd.read_csv('Statement Jun.csv')
    formatter(df)

    #Find management information
    totalCompany = totalPerCompany(df)
    totalExpense = df["Paid out"].sum()
    totalGain = df["Paid in"].sum()
    startingBalance = df["Balance"].iloc[0]
    endingBalance = df["Balance"].iloc[-1]

    print("Data sucessfully formatted from CSV! \n")
    infoDF, companyDF = createGSheets(totalCompany, totalExpense, totalGain, startingBalance, endingBalance)
    setGSheets(infoDF, companyDF)

if __name__ == "__main__":
    main()