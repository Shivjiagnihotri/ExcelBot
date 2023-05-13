import pandas as pd
import gradio as gr

# IF YOU WANT TO LOAD THE EXCEL SHEET DIRECTLY FROM YOUR COMPUTER, GIVE THE FILE NAME IN BELOW LINE
df = pd.read_excel("Onboarding Form(1-1) (4).xlsx")

#IF YOU WANT TO LOAD THE EXCEL SHEET FROM GOOGLE FORM, CREATE A GOOGLE FORM, AND GET THE SPREADSHEETS'S LINK
#url = 'https://1drv.ms/x/s!As0NRhKJW7hMiGvmgA65EzsH2LQ2?e=5ENiwH'

#df = pd.read_excel(url)

greeting = "\nGreetings! I am your Excel bot. I can search data inside any Excel sheet. \n(Also by using the CSV file link of Google Form responses.)"

greeting2 = "\nCurrently I support Four commands"

greeting3 = "\n 1. Search the Excel \n 2. Show Columns \n 3. Onboarding Link \n 4. Stop"

# Define a function to take text input and generate a response
def talk(text):

    # Get input from the user
    prompt = text
    response = ""

    if "search the excel" in prompt.lower():
        response += "\nSure!"

        # Check if the prompt contains the "search the excel" keyword
        if "search the excel" in prompt.lower() and "stop" not in prompt.lower():
            # Ask the user for the column to search in
            response += "\nWhat column do you want to search in?"
            df.columns = map(str.lower, df.columns)
            # Get the column to search in from the user's input

            column = input("You: ").lower()
            if column not in df.columns:
                response += f"\nColumn {column} not found in the Excel sheet!"
            else:
                response += f"\nType the {column} of the person."

            # Get the value to search for from the user's input
            value = input("You: ").lower()
            para1 = ("Sure, searching the excel sheet for the", column, "of", value)
            response += f"\n{para1}"

            # Find the corresponding row in the Excel file
            row = df.loc[df[column].str.lower() == value.lower()]

            if not row.empty:
                # Ask the user for the column to return the value from
                response += "\nWhat value do you want to get?"

                # Get the column to return the value from from the user's input
                return_column = input("You: ").lower()
                if return_column not in df.columns:
                    response += f"\nColumn {return_column} not found in the Excel sheet!"
                else:
                    # Extract the value from the row
                    return_value = row[return_column].values[0]

                    # Output the return value
                    answer = f"The {return_column} of the row where {column} is {value} is ['{return_value}']."
                    response += f"\nAI says: {answer}\nResults Found!"
            else:
                # Check if the specified column is present in the dataframe
                if column in df.columns:
                    answer = f"I'm sorry, I could not find {value} in the {column} column."
                else:
                    answer = "The column or the row is not present."
                response += f"\nAI says: {answer}\nResults Not Found!"
            # break

    elif "onboarding link" in prompt.lower():
        response += "https://forms.gle/rdXJxd4La4nM72Qj8\n"
        response += "What else can I help you with?"
    
    elif "columns" in prompt.lower():
        response += f"Total {len(df.columns)} columns found!\n"
        response += f"The columns are as follows:- {', '.join(df.columns)}\n"
        response += "What else can I help you with?"
        
    elif "stop" in prompt.lower():
        response += "Goodbye!"
        
    else:
        response += "Invalid response! Kindly try again.\n"
        response += "What else can I help you with?"

    return response

# Create a Gradio interface for the Excel bot
chatbot = gr.Interface(
    fn=talk,
    inputs=gr.inputs.Textbox(label="What can I help you with today?"),
    outputs=gr.outputs.Textbox(),
    examples=[
        ["Search the Excel sheet for the name of John.", "What value do you want to get?", "Email"],
        ["Show me the columns in the Excel sheet.", "", ""],
        ["Stop the program.", "", ""]
    ],
    title="Excel Bot",
    description="A chatbot that can search data inside an Excel sheet.",
    article="This is a demo of a chatbot that can search data inside an Excel sheet. You can ask the bot to search for a value in a specific row and column, or to show you the columns in the sheet. To use the bot, type your query in the input box and click \"Submit\". The bot will then process your query and display the result in the output box.")

# Launch the interface
chatbot.launch(share=True)
