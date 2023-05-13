import openai
import pandas as pd
import pyttsx3

# Set up the OpenAI API client with your API key
openai.api_key = "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"

# Set up the text-to-speech engine
engine = pyttsx3.init()

# Load the Excel file from computer
#df = pd.read_excel("Recruitment_Details.xlsx")

#Load the Excel file live from google responses
url = 'https://docs.google.com/spreadsheets/d/1On7Zf8Q5GD-0ReKGYndJNbLMbt_a4yEn8OqKjMXBeQ8/export?format=csv'

df = pd.read_csv(url)

greeting = "Greetings! I am your Excel bot. I can search data inside any Excel sheet. \n (Also by using the CSV file link of Google Form responses.)"
print(greeting)
engine.say(greeting)
engine.runAndWait()
greeting2 = "Currently I support three commands"
print(greeting2)
# engine.say(greeting2)
# engine.runAndWait()
greeting3 = "1. Search the excel \n 2. Give the onboarding link \n 3. Stop"
print(greeting3)
# engine.say("1., Search the excel, 2. Give the onboarding link, and 3. Stop")
# engine.runAndWait()
# Define a function to take text input and generate a response
def talk():

    # Get input from the user
    prompt = input("You: ")
    if "search the excel" in prompt.lower():

        print("\nSure!")
        engine.say("Sure!")
        engine.runAndWait()
        # Check if the prompt contains the "search the excel" keyword
        if "search the excel" in prompt.lower() and "stop" not in prompt.lower():
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
            para1=("Sure, searching the excel sheet for the" , column, "of" , value)
            print(para1)
            engine.say(para1)
            engine.runAndWait()

            # Find the corresponding row in the Excel file
            row = df.loc[df[column].str.lower() == value.lower()]
            if not row.empty:
                # Ask the user for the column to return the value from
                para=("Total", len(df.columns), "columns found!")
                print(para)
                engine.say(para)
                engine.runAndWait()
                print("The columns are as follows:- ", df.columns.values)
                print("What value do you want to get?")
                engine.say("What value do you want to get?")
                engine.runAndWait()

                # Get the column to return the value from from the user's input
                return_column = input("You: ").lower()
                para3=("Searching the", return_column, "inside the excel sheet")
                print(para3)
                engine.say(para3)
                engine.runAndWait()

                # Extract the value from the row
                return_value = row[return_column].values[0]

                # Output the return value
                answer = f"The {return_column} of the row where {column} is {value} is ['{return_value}']."
                print("AI says:", answer)
                engine.say(answer)
                engine.runAndWait()
                print("Results Found!")
            
            else:
                # Check if the specified column is present in the dataframe
                if column in df.columns:
                    answer = f"I'm sorry, I could not find {value} in the {column} column."
                    engine.say(answer)
                    engine.runAndWait()
                else:
                    answer = "The column or the row is not present."
                    engine.say(answer)
                    engine.runAndWait()
                print("AI says:", answer)
                print("Results Not Found!")
                # break
    
    elif "onboarding link" in prompt.lower():
        print("https://forms.gle/rdXJxd4La4nM72Qj8")
        engine.say("Here is the onboarding link")
        engine.runAndWait()
        
    
    elif "stop" in prompt.lower():
        # Output a goodbye message and stop the program
        answer = "Goodbye!"
        print("AI says:", answer)
        engine.say("Have a good day!")
        engine.runAndWait()
        breakpoint()

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
