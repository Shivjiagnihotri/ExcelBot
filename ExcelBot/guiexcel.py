import openai
import pandas as pd
import tkinter as tk

# Set up the OpenAI API client with your API key
openai.api_key = "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"

# Load the Excel file from computer
#df = pd.read_excel("Recruitment_Details.xlsx")

#Load the Excel file live from google responses
url = 'https://docs.google.com/spreadsheets/d/1On7Zf8Q5GD-0ReKGYndJNbLMbt_a4yEn8OqKjMXBeQ8/export?format=csv'

df = pd.read_csv(url)

# Create a Tkinter window
root = tk.Tk()
root.title("Excel Bot")

# Define a function to handle button clicks
def search_excel():
    # Get the user's inputs from the entry fields
    column = column_entry.get().lower()
    value = value_entry.get().lower()

    # Find the corresponding row in the Excel file
    row = df.loc[df[column].str.lower() == value.lower()]

    if not row.empty:
        # Get the user's desired return column from the dropdown menu
        return_column = return_dropdown.get().lower()

        # Extract the value from the row
        return_value = row[return_column].values[0]

        # Output the return value in the result label
        result_label.config(text=f"The {return_column} of the row where {column} is {value} is {return_value}.")
    else:
        # Display an error message in the result label
        result_label.config(text=f"I'm sorry, I could not find {value} in the {column} column.")

# Create the GUI elements
greeting_label = tk.Label(root, text="Hello, how can I assist you?")
greeting_label.pack(pady=10)

column_label = tk.Label(root, text="What column do you want to search in?")
column_label.pack(pady=5)

column_entry = tk.Entry(root)
column_entry.pack()

value_label = tk.Label(root, text="Type the value of the person.")
value_label.pack(pady=5)

value_entry = tk.Entry(root)
value_entry.pack()

return_label = tk.Label(root, text="What value do you want to get?")
return_label.pack(pady=5)

return_options = df.columns.tolist()
return_dropdown = tk.StringVar(root)
return_dropdown.set(return_options[0])
return_menu = tk.OptionMenu(root, return_dropdown, *return_options)
return_menu.pack()

search_button = tk.Button(root, text="Search", command=search_excel)
search_button.pack(pady=10)

result_label = tk.Label(root, text="")
result_label.pack(pady=10)

# Start the Tkinter event loop
root.mainloop()
