import openai
import pandas as pd
import gradio as gr

#Set up the OpenAI API client with your API key
#openai.api_key = "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"

#Load the Excel file from computer
#df = pd.read_excel("Recruitment_Details.xlsx")

#Load the Excel file live from google responses
url = 'https://docs.google.com/spreadsheets/d/1On7Zf8Q5GD-0ReKGYndJNbLMbt_a4yEn8OqKjMXBeQ8/export?format=csv'

df = pd.read_csv(url)

greeting = "Hello, how can I assist you?"

#Define a function to take text input and generate a response
def talk(prompt):
    # Get input from the user
    # prompt = input("You: ")
    if "search the excel" in prompt.lower():

        response = "\nSure!"
        # Check if the prompt contains the "search the excel" keyword
        while "search the excel" in prompt.lower() and "stop" not in prompt.lower():
            # Ask the user for the column to search in
            response += "\nWhat column do you want to search in?"
            # Get the column to search in from the user's input
            column = gr.inputs.Textbox(label="Column")
            value = gr.inputs.Textbox(label="Value")

            # Ask the user for the value to search for
            #print("What information do you want to search for?")

            # Get the value to search for from the user's input
            # value = input("You: ").lower()
            response += "Sure, searching the excel sheet for the" , column, "of" , value

            # Find the corresponding row in the Excel file
            row = df.loc[df[column].str.lower() == value.lower()]

            if not row.empty:
                # Ask the user for the column to return the value from
                response += "Total", len(df.columns), "columns found!"
                response += "The columns are as follows:- ", df.columns.values
                response += "What value do you want to get?"
                # Get the column to return the value from from the user's input
                return_column = gr.inputs.Textbox(label="Return Column")
                response += "Searching the", return_column, "inside the excel sheet"

                # Extract the value from the row
                return_value = row[return_column].values[0]

                # Output the return value
                answer = f"The {return_column} of the row where {column} is {value} is ['{return_value}']."
                response += "AI says:", answer
                response += "Results Found!"
            else:
                # Check if the specified column is present in the dataframe
                if column in df.columns:
                    answer = f"I'm sorry, I could not find {value} in the {column} column."
                else:
                    answer = "The column or the row is not present."
                response += "AI says:", answer
                response += "Results Not Found!"
                break

    elif "onboarding link" in prompt.lower():
        response = "https://forms.gle/rdXJxd4La4nM72Qj8"

    elif "stop" in prompt.lower():
        # Output a goodbye message and stop the program
        response = "Goodbye!"
    else:
        response = "Invalid Response! Kindly try again."

    return response
#Set up the Gradio interface
import gradio as gr

iface = gr.Interface(fn=talk,
inputs=gr.inputs.Textbox(label="You: "),
outputs=gr.outputs.Textbox(label="AI:"))

#Launch the interface
iface.launch()
