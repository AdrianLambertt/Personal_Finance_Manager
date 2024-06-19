import pandas as pd
import gspread
import decimal as d


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


#TODO finish this function for sheets output
#This function will combine the data pulled from the CSV to create a Dataframe ready for the sheets API.
def createGSheets(totalCompany, totalExpense, TotalGain, startingBalance, endingBalance):
    formatDF = pd.DataFrame.from_dict(totalCompany, orient='index', columns=['Total Paid']).reset_index().sort_values(by='Total Paid', ascending=False)
    formatDF.columns = ['Company', 'Total Paid']


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
    sheetsDF = createGSheets(totalCompany, totalExpense, totalGain, startingBalance, endingBalance)

    #sa = gspread.service_account()
    #sh = sa.open("Finance Manager")


if __name__ == "__main__":
    main()