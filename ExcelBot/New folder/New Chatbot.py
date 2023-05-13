import pandas as pd


# IF YOU WANT TO LOAD THE EXCEL SHEET DIRECTLY FROM YOUR COMPUTER, GIVE THE FILE NAME IN BELOW LINE
df = pd.read_excel("Onboarding Form(1-1) (4).xlsx")

#IF YOU WANT TO LOAD THE EXCEL SHEET FROM GOOGLE FORM, CREATE A GOOGLE FORM, AND GET THE SPREADSHEETS'S LINK
#url = 'https://1drv.ms/x/s!As0NRhKJW7hMiGvmgA65EzsH2LQ2?e=5ENiwH'

#df = pd.read_excel(url)

greeting = "\nHow can I help you? You can choose the below option number."
print(greeting)
greeting2 = "\n 1. New Onboarding \n 2. New Onboarding Status \n 3. New R2D2 Req creation \n 4. Status of R2D2 req \n 5. help"
print(greeting2)
result = "start"

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

    elif "1" in prompt.lower():
        print("Chatbot: Sure, please submit the below onboarding form: ")
        print("\nhttps://forms.office.com/Pages/DesignPageV2.aspx?origin=NeoPortalPage&subpage=design&id=Wq6idgCfa0-V7V0z13xNYSwcjiwSIm9EvFGzPjL2_5FUMThIRlNPRkVTRkhNNUc3U0w5RkRMUzI0OC4u&analysis=true")
        talk()
    
    elif "2" in prompt.lower():
        print("Chatbot: Sure, please share the associate's complete email address:")
        column = "Email" #Change the colomn name here if you get any key or value error
        value = input("Human: ").lower()
        if "@" not in value:
            print("Chatbot: Invalid email")
        else:
            # filter rows with email value
            filtered_rows = df[df[column] == str(value)]

            # check if any row matches the email provided
            if filtered_rows.empty:
                print("No data found for email: ", value)
            else:
                # get the first row of filtered rows
                row_data = filtered_rows.iloc[0]

                print("The onboarding status is as follows:\n")
                # iterate over all columns and print column name with its value
                for column_name, value in row_data.iteritems():
                    print(column_name + ":" + str(value))

    elif "3" in prompt.lower():
        print("Chatbot: Under construction")
    
    elif "4" in prompt.lower():
        print("Chatbot: Under construction")
    
    elif "5" in prompt.lower() or "help" in prompt.lower():

        print("Chatbot: Hello, I am your Chatbot. \n I am able to perform majorly four tasks." )
        print(greeting2)
        print("\n some other commands I support:\n 6. show columns --returns the total columns present in the onboarding form \n 7. advanced search --advanced tool to search the form \n 8. stop or quit --ends the program")

    elif "6" in prompt.lower() or "column" in prompt.lower():
        para=("Total", len(df.columns), "columns found!")
        print(para)
        print("Chatbot:", df.columns.values)
        talk()
    
    # Check if the prompt contains the "search the excel" keyword
    elif  "7" in prompt.lower() or "advanced search" in prompt.lower():
        
        print("\nChatbot: Sure!")
        
        # Ask the user for the column to search in
        print("Chatbot: What column do you want to search in?")
        df.columns = map(str.lower, df.columns)
        
        # Get the column to search in from the user's input
        column = input("Human: ").lower()
        if column not in df.columns:
            print(f"Chatbot: Column {column} not found in the Excel sheet!")
            #talk() #phase 1, here i get the column
        else: 
            print("Chatbot: Type the", column, "of the person.")
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
                    # Extract the value from the row(HERE I CROSSED, MEANS BY USING THE ONE I GOT OTHER VALUE)
                    return_value = row[return_column].values[0]

                    # Output the return value
                    answer = f"The {return_column} of the row where {column} is {value} is ['{return_value}']."
                    print("Chatbot:", answer)
                    print("Chatbot: Results Found!")
                    talk()
            
            else:
                # Check if the specified column is present in the dataframe
                if column in df.columns:
                    answer = f"Person with this {column} is not present."
                else:
                    answer = "No such column is present."
                print("Chatbot: ", answer)
                print("Chatbot: Results Not Found!")
                talk()
    
    elif "8" in prompt.lower() or "stop" in prompt.lower() or "bye" in prompt.lower() or "exit" in prompt.lower() or "quit" in prompt.lower():
        # Output a goodbye message and stop the program
        answer = "Goodbye!"
        print("Chatbot:", answer)
        # Return "stop" to indicate that the user wants to stop the program
        return "stop"

    else:
        print("Chatbot: I Cannot recognize what you say, other than what i am designed for ;) . \n Start by choosing any of the above option number")
        talk()

while True:
    # Check if the user wants to stop the program
    if result == "stop":
        # Exit the while loop
        break
    # Call the talk() function and get its return value
    else: 
        result = talk()
    