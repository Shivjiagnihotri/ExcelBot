import pandas as pd
import gradio as gr

# Load the Excel sheet
df = pd.read_excel("Onboarding Form(1-1) (4).xlsx")

# Define a function to search the Excel sheet
def search_excel(column, value, return_column):
    # Find the corresponding row in the Excel file
    row = df.loc[df[column].str.lower() == value.lower()]

    if not row.empty:
        # Extract the value from the row
        return_value = row[return_column].values[0]

        # Output the return value
        answer = f"The {return_column} of the row where {column} is {value} is ['{return_value}']."
        return answer, True

    else:
        # Check if the specified column is present in the dataframe
        if column in df.columns:
            answer = f"I'm sorry, I could not find {value} in the {column} column."
        else:
            answer = "The column or the row is not present."
        return answer, False

# Define a function to generate a response
def excel_bot(text):
    # Check the input text for relevant commands
    if "search the excel" in text.lower():
        # Ask the user for the column to search in
        column_prompt = "What column do you want to search in?"
        column_options = list(df.columns)
        column = gr.inputs.Dropdown(choices=column_options, label=column_prompt)

        # Ask the user for the value to search for
        value_prompt = "Type the value you want to search for."
        value = gr.inputs.Textbox(label=value_prompt)

        # Ask the user for the column to return the value from
        return_column_prompt = "What value do you want to get?"
        return_column_options = list(df.columns)
        return_column = gr.inputs.Dropdown(choices=return_column_options, label=return_column_prompt)

        # Search the Excel sheet
        result, found = search_excel(column.value, value.value, return_column.value)

        # Output the result
        if found:
            output = f"{result}\n Results found!"
        else:
            output = f"{result}\n Results not found!"

    elif "onboarding link" in text.lower():
        # Output the onboarding link
        output = "https://forms.gle/rdXJxd4La4nM72Qj8"

    elif "columns" in text.lower():
        # Output the list of columns in the Excel sheet
        columns = ", ".join(list(df.columns))
        output = f"Total {len(df.columns)} columns found!\n The columns are as follows: {columns}"

    elif "stop" in text.lower():
        # Output a goodbye message and stop the program
        output = "Goodbye!"
        return output
        js = "window.close();"
        return gr.outputs.HTML(output), gr.outputs.Javascript(js)

    else:
        # Output an error message
        output = "Invalid Response! Kindly try again."

    return output

# Create a Gradio interface for the Excel bot
chatbot = gr.Interface(
    fn=excel_bot,
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
