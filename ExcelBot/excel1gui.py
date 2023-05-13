import openai
import pandas as pd
import tkinter as tk

# Set up the OpenAI API client with your API key
openai.api_key = "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"

# Load the Excel file live from google responses
url = 'https://docs.google.com/spreadsheets/d/1On7Zf8Q5GD-0ReKGYndJNbLMbt_a4yEn8OqKjMXBeQ8/export?format=csv'
df = pd.read_csv(url)

def get_response():
    global df
    user_input = entry.get().lower()
    if "search the excel" in user_input:
        response_text.set("Sure!")
        while "search the excel" in user_input and "stop" not in user_input:
            response_text.set("What column do you want to search in?")
            window.update()

            column = column_entry.get().lower()
            response_text.set("Type the {} of the person.".format(column))
            window.update()

            value = value_entry.get().lower()
            response_text.set("Sure, searching the excel sheet for the {} of {}.".format(column, value))
            window.update()

            row = df.loc[df[column].str.lower() == value.lower()]

            if not row.empty:
                response_text.set("Total {} columns found! The columns are as follows: {}. What value do you want to get?".format(len(df.columns), df.columns.values))
                window.update()

                return_column = return_entry.get().lower()
                response_text.set("Searching the {} inside the excel sheet.".format(return_column))
                window.update()

                return_value = row[return_column].values[0]

                response_text.set("The {} of the row where {} is {} is ['{}'].".format(return_column, column, value, return_value))
                window.update()
            else:
                if column in df.columns:
                    response_text.set("I'm sorry, I could not find {} in the {} column.".format(value, column))
                else:
                    response_text.set("The column or the row is not present.")
                window.update()
                break
    elif "onboarding link" in user_input:
        response_text.set("https://forms.gle/rdXJxd4La4nM72Qj8")
    elif "stop" in user_input:
        response_text.set("Goodbye!")
        window.update()
        window.after(2000, window.destroy)
    else:
        response_text.set("Invalid Response! Kindly try again.")
        window.update()

# Create the GUI
window = tk.Tk()
window.title("ExcelBot")
window.geometry("600x400")

# Create the input field and button
entry = tk.Entry(window, width=50)
entry.place(relx=0.5, rely=0.3, anchor="center")
button = tk.Button(window, text="Send", command=get_response)
button.place(relx=0.7, rely=0.3, anchor="center")

# Create the labels and entry fields for searching the Excel file
column_label = tk.Label(window, text="What column do you want to search in?")
column_label.place(relx=0.25, rely=0.4, anchor="center")
column_entry = tk.Entry(window, width=20)
column_entry.place(relx=0.5, rely=0.4, anchor="center")
value_label = tk.Label(window,text="Type the value of the person:")
value_label.place(relx=0.25, rely=0.5, anchor="center")
value_entry = tk.Entry(window, width=20)
value_entry.place(relx=0.5, rely=0.5, anchor="center")

def search_excel():
    # Get the column to search in from the user's input
    column = column_entry.get().lower()
    print("Type the", column, "of the person.")
    
    # Get the value to search for from the user's input
    value = value_entry.get().lower()
    print("Sure, searching the excel sheet for the" , column, "of" , value)
    
    # Find the corresponding row in the Excel file
    row = df.loc[df[column].str.lower() == value.lower()]
    
    if not row.empty:
        # Get the column to return the value from from the user's input
        return_column = return_entry.get().lower()
        print("Searching the", return_column, "inside the excel sheet")
        
        # Extract the value from the row
        return_value = row[return_column].values[0]
        
        # Output the return value
        answer = f"The {return_column} of the row where {column} is {value} is ['{return_value}']."
        print("AI says:", answer)
        result_label.config(text=answer)
    else:
        # Check if the specified column is present in the dataframe
        if column in df.columns:
            answer = f"I'm sorry, I could not find {value} in the {column} column."
        else:
            answer = "The column or the row is not present."
        print("AI says:", answer)
        result_label.config(text=answer)


#Create the button to search the Excel file
search_button = tk.Button(window, text="Search", command=search_excel)
search_button.place(relx=0.5, rely=0.6, anchor="center")

#Create a label to display the search results
results_label = tk.Label(window, text="")
results_label.place(relx=0.5, rely=0.8, anchor="center")

#Run the window
window.mainloop()
