import openai
import pandas as pd
import pyttsx3
import speech_recognition as sr

# Set up the OpenAI API client with your API key
openai.api_key = "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"

# Set up the text-to-speech engine
engine = pyttsx3.init()

# Load the Excel file into a DataFrame
df = pd.read_excel("Recruitment_Details.xlsx")

# Define a function to take text input and generate a spoken response
def talk():
    # Output the greeting message as spoken text
    greeting = "Hello, how can I assist you?"
    print(greeting)
    engine.say(greeting)
    engine.runAndWait()

    # Get input from the user
    prompt = input("You: ")

    try:
        # Check if the prompt contains the "search the excel" keyword
        if "search the excel" in prompt.lower():
            # Ask the user for the column to search in
            engine.say("Sure, what column do you want to search in?")
            engine.runAndWait()

            # Get the column to search in from the user's input
            column = input("You: ").lower()
            print("You want to search in the", column, "column.")

            # Ask the user for the value to search for
            engine.say("What value do you want to search for?")
            engine.runAndWait()

            # Get the value to search for from the user's input
            value = input("You: ").lower()
            print("You want to search for", value, "in the", column, "column.")

            # Find the corresponding row in the Excel file
            row = df.loc[df[column].str.lower() == value]

            if not row.empty:
                # Ask the user for the column to return the value from
                engine.say("Sure, what column do you want to get the value from?")
                engine.runAndWait()

                # Get the column to return the value from from the user's input
                return_column = input("You: ").lower()
                print("You want to get the value from the", return_column, "column.")

                # Extract the value from the row
                return_value = row[return_column].values[0]

                # Output the return value as spoken text
                answer = f"The {return_column} of the row where {column} is {value} is {return_value}."
                print("AI says:", answer)
                engine.say(answer)
                engine.runAndWait()

            else:
                # Output a message indicating that the value was not found in the specified column
                answer = f"I'm sorry, I could not find {value} in the {column} column."
                print("AI says:", answer)
                engine.say(answer)
                engine.runAndWait()

        else:
            # Generate a response using the OpenAI API
            response = openai.Completion.create(
                engine="text-davinci-003",
                prompt=prompt,
                max_tokens=1024,
                temperature=0.5,
            )
            answer = response.choices[0].text.strip()

            # Output the response as spoken text
            print("AI says:", answer)
            engine.say(answer)
            engine.runAndWait()

    except sr.UnknownValueError:
        # Handle speech recognition errors
        print("Sorry, I didn't understand that.")
        engine.say("Sorry, I didn't understand that.")
        engine.runAndWait()

    except sr.RequestError as e:
        print("Could not request results from speech recognition service; {0}".format(e))
        engine.say("Sorry, I'm having trouble connecting to the speech recognition service. Please try again later.")
        engine.runAndWait()

# Continuously prompt the user for input until they say "stop"
while True:
    talk()
    prompt = input("You: ")
    if prompt.lower() == "stop":
        break