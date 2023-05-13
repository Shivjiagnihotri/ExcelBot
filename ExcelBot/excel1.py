import openai
import pandas as pd

# Set up the OpenAI API client with your API key
openai.api_key = "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"

# Load the Excel file from computer
#df = pd.read_excel("Recruitment_Details.xlsx")

#Load the Excel file live from google responses
url = 'https://docs.google.com/spreadsheets/d/1On7Zf8Q5GD-0ReKGYndJNbLMbt_a4yEn8OqKjMXBeQ8/export?format=csv'

df = pd.read_csv(url)

greeting = "Hello, how can I assist you?"
print(greeting)

# Define a function to take text input and generate a response
def talk():

    # Get input from the user
    prompt = input("You: ")
    if "search the excel" in prompt.lower():

        print("\nSure!")
        # Check if the prompt contains the "search the excel" keyword
        while "search the excel" in prompt.lower() and "stop" not in prompt.lower():
            # Ask the user for the column to search in
            print("\nWhat column do you want to search in?")

            df.columns = map(str.lower, df.columns)
            # Get the column to search in from the user's input
            column = input("You: ").lower()
            print("Type the", column, "of the person.")

            # Ask the user for the value to search for
            #print("What information do you want to search for?")

            # Get the value to search for from the user's input
            value = input("You: ").lower()
            print("Sure, searching the excel sheet for the" , column, "of" , value)

            # Find the corresponding row in the Excel file
            row = df.loc[df[column].str.lower() == value.lower()]

            if not row.empty:
                # Ask the user for the column to return the value from
                print("Total", len(df.columns), "columns found!")
                print("The columns are as follows:- ", df.columns.values)
                print("What value do you want to get?")

                # Get the column to return the value from from the user's input
                return_column = input("You: ").lower()
                print("Searching the", return_column, "inside the excel sheet")

                # Extract the value from the row
                return_value = row[return_column].values[0]

                # Output the return value
                answer = f"The {return_column} of the row where {column} is {value} is ['{return_value}']."
                print("AI says:", answer)
                print("Results Found!")
            else:
                # Check if the specified column is present in the dataframe
                if column in df.columns:
                    answer = f"I'm sorry, I could not find {value} in the {column} column."
                else:
                    answer = "The column or the row is not present."
                print("AI says:", answer)
                print("Results Not Found!")
                break
    
    elif "onboarding link" in prompt.lower():
        print("https://forms.gle/rdXJxd4La4nM72Qj8")
    
    elif "stop" in prompt.lower():
        # Output a goodbye message and stop the program
        answer = "Goodbye!"
        print("AI says:", answer)

    else:
        print("AI says: Invalid Response! Kindly try again.")

# Continuously prompt the user for input until they say "stop"
while True:
    talk()


#CHATGPT ELSE PART BELOW. COPY PASTE THIS CODE AND USE A VALID API TO ACTIVATE CHATGPT REPLY.
# else:
#         # Generate a response using the OpenAI API
#         response = openai.Completion.create(
#             engine="text-davinci-003",
#             prompt=prompt,
#             max_tokens=1024,
#             temperature=0.5,
#         )
#         answer = response.choices[0].text.strip()
#         # Output the response
#         print("AI says:", answer)
