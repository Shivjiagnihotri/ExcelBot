import pandas as pd


# IF YOU WANT TO LOAD THE EXCEL SHEET DIRECTLY FROM YOUR COMPUTER, GIVE THE FILE NAME IN BELOW LINE
df = pd.read_excel("Onboarding Form(1-1) (4).xlsx")

#IF YOU WANT TO LOAD THE EXCEL SHEET FROM GOOGLE FORM, CREATE A GOOGLE FORM, AND GET THE SPREADSHEETS'S LINK
#url = 'https://1drv.ms/x/s!As0NRhKJW7hMiGvmgA65EzsH2LQ2?e=5ENiwH'

#df = pd.read_excel(url)

greeting = "Greetings! I am your Excel Bot."
print(greeting)

greeting3 = "\n 1. Search \n 2. Show Columns \n 3. Onboarding Link \n 4. Stop"
print(greeting3)


# Define a function to take text input and generate a response
def talk():

    # Get input from the user
    prompt = input("\nYou: ")
    if "hi" and "hey" in prompt.lower():
        print("Hello")
    elif "how are you" and "how are you doing" in prompt.lower():
        print("Hello, I am good.") 
    elif "thanks" and "thank you" in prompt.lower():
        print("My pleasure!")
    elif "search the excel" in prompt.lower():

        print("\nSure!")
        
        # Check if the prompt contains the "search the excel" keyword
        if "search the excel" in prompt.lower() and "stop" not in prompt.lower():
            # Ask the user for the column to search in
            print("\nWhat column do you want to search in?")
            df.columns = map(str.lower, df.columns)
            # Get the column to search in from the user's input

            column = input("You: ").lower()
            if column not in df.columns:
                print(f"Column {column} not found in the Excel sheet!")
                talk()
            else: print("Type the", column, "of the person.")
                
            # Get the value to search for from the user's input
            value = input("You: ").lower()
            para1 = ("Sure, searching the excel sheet for the", column, "of", value)
            print(para1)

            # Find the corresponding row in the Excel file
            row = df.loc[df[column].str.lower() == value.lower()]
            

            if not row.empty:
                # Ask the user for the column to return the value from
                print("What value do you want to get?")

                # Get the column to return the value from from the user's input
                return_column = input("You: ").lower()
                if return_column not in df.columns:
                    print(f"Column {return_column} not found in the Excel sheet!")
                    talk()
                else: 
                    # Extract the value from the row
                    return_value = row[return_column].values[0]

                    # Output the return value
                    answer = f"The {return_column} of the row where {column} is {value} is ['{return_value}']."
                    print("AI says:", answer)
                    print("Results Found!")
                    talk()
            
            else:
                # Check if the specified column is present in the dataframe
                if column in df.columns:
                    answer = f"I'm sorry, I could not find {value} in the {column} column."
                else:
                    answer = "The column or the row is not present."
                print("AI says:", answer)
                print("Results Not Found!")
                talk()
                # break
    
    elif "onboarding link" in prompt.lower():
        print("Sure, this is the link: ")
        print("\nhttps://forms.office.com/Pages/DesignPageV2.aspx?origin=NeoPortalPage&subpage=design&id=Wq6idgCfa0-V7V0z13xNYSwcjiwSIm9EvFGzPjL2_5FUMThIRlNPRkVTRkhNNUc3U0w5RkRMUzI0OC4u&analysis=true")
        talk()

    elif "columns" in prompt.lower():
        para=("Total", len(df.columns), "columns found!")
        print(para)
        print("The columns are as follows:- ", df.columns.values)
        talk()
        
    elif "stop" in prompt.lower():
        # Output a goodbye message and stop the program
        answer = "Goodbye!"
        print("AI says:", answer)
        False
        

    else:
        print("AI says: Invalid Response! Kindly try again.")
        talk()

while True:
    talk()