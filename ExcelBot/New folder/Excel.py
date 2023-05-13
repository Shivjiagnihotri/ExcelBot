import pandas as pd


# IF YOU WANT TO LOAD THE EXCEL SHEET DIRECTLY FROM YOUR COMPUTER, GIVE THE FILE NAME IN BELOW LINE
df = pd.read_excel("Onboarding Form(1-1) (4).xlsx")

#IF YOU WANT TO LOAD THE EXCEL SHEET FROM GOOGLE FORM, CREATE A GOOGLE FORM, AND GET THE SPREADSHEETS'S LINK
#url = 'https://1drv.ms/x/s!As0NRhKJW7hMiGvmgA65EzsH2LQ2?e=5ENiwH'

#df = pd.read_excel(url)

greeting = "\nGreetings! I am your Excel Bot."
print(greeting)

greeting3 = "\n 1. Search \n 2. Show Columns \n 3. Onboarding Link \n 4. Stop"
print(greeting3)


# Define a function to take text input and generate a response
def talk():

    # Get input from the user
    prompt = input("\nHuman: ")
    
    if "how are you" in prompt.lower() or "how are you doing" in prompt.lower() or "hi, how are you" in prompt.lower() or "hello, how are you" in prompt.lower():
        print("Chatbot: Hello, I am doing well, thank you for asking.")
    elif "hi" in prompt.lower() or "hey" in prompt.lower() or "hello" in prompt.lower():
        print("Chatbot: Hello!")
    elif "thanks" in prompt.lower() or "thank you" in prompt.lower():
        print("Chatbot: My pleasure!")
    elif "search the excel" in prompt.lower():

        print("\nChatbot: Sure!")
        
        # Check if the prompt contains the "search the excel" keyword
        if "search" in prompt.lower() and "stop" not in prompt.lower():
            # Ask the user for the column to search in
            print("Chatbot: What column do you want to search in?")
            df.columns = map(str.lower, df.columns)
            # Get the column to search in from the user's input

            column = input("Human: ").lower()
            if column not in df.columns:
                print(f"Chatbot: Column {column} not found in the Excel sheet!")
                talk()
            else: print("Chatbot: Type the", column, "of the person.")
                
            # Get the value to search for from the user's input
            value = input("Human: ").lower()
            para1 = ("Sure, searching the excel sheet for the", column, "of", value)
            print(para1)

            # Find the corresponding row in the Excel file
            row = df.loc[df[column].str.lower() == value.lower()]
            

            if not row.empty:
                # Ask the user for the column to return the value from
                print("Chatbot: What value do you want to get?")

                # Get the column to return the value from from the user's input
                return_column = input("Human: ").lower()
                if return_column not in df.columns:
                    print(f"Chatbot: Column {return_column} not found in the Excel sheet!")
                    talk()
                else: 
                    # Extract the value from the row
                    return_value = row[return_column].values[0]

                    # Output the return value
                    answer = f"The {return_column} of the row where {column} is {value} is ['{return_value}']."
                    print("Chatbot:", answer)
                    print("Chatbot: Results Found!")
                    talk()
            
            else:
                # Check if the specified column is present in the dataframe
                if column in df.columns:
                    answer = f"I'm sorry, I could not find {value} in the {column} column."
                else:
                    answer = "The column or the row is not present."
                print("Chatbot: ", answer)
                print("Chatbot: Results Not Found!")
                talk()
                # break
    
    elif "onboarding link" in prompt.lower():
        print("Chatbot: Sure, this is the link: ")
        print("\nhttps://forms.office.com/Pages/DesignPageV2.aspx?origin=NeoPortalPage&subpage=design&id=Wq6idgCfa0-V7V0z13xNYSwcjiwSIm9EvFGzPjL2_5FUMThIRlNPRkVTRkhNNUc3U0w5RkRMUzI0OC4u&analysis=true")
        talk()

    elif "columns" in prompt.lower():
        para=("Total", len(df.columns), "columns found!")
        print(para)
        print("Chatbot: The columns are as follows:- ", df.columns.values)
        talk()
    
    elif "stop" in prompt.lower():
        # Output a goodbye message and stop the program
        answer = "Goodbye!"
        print("Chatbot:", answer)
        # Return "stop" to indicate that the user wants to stop the program
        return "stop"

    else:
        print("Chatbot: Invalid Response! Kindly try again.")
        talk()

while True:
    # Call the talk() function and get its return value
    result = talk()
    # Check if the user wants to stop the program
    if result == "stop":
        # Exit the while loop
        break
    