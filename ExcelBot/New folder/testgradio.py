import gradio as gr
import pandas as pd
from transformers import AutoModelForCausalLM, AutoTokenizer
import torch


# IF YOU WANT TO LOAD THE EXCEL SHEET DIRECTLY FROM YOUR COMPUTER, GIVE THE FILE NAME IN BELOW LINE
df = pd.read_excel("Onboarding Form(1-1) (4).xlsx")

#IF YOU WANT TO LOAD THE EXCEL SHEET FROM GOOGLE FORM, CREATE A GOOGLE FORM, AND GET THE SPREADSHEETS'S LINK
#url = 'https://1drv.ms/x/s!As0NRhKJW7hMiGvmgA65EzsH2LQ2?e=5ENiwH'

#df = pd.read_excel(url)

greeting = "\nGreetings! I am your Excel bot. I can search data inside any Excel sheet. \n(Also by using the CSV file link of Google Form responses.)"

greeting2 = "\nCurrently I support Four commands"

greeting3 = "\n 1. Search the Excel \n 2. Show Columns \n 3. Onboarding Link \n 4. Stop"

tokenizer = AutoTokenizer.from_pretrained("microsoft/DialoGPT-medium")
model = AutoModelForCausalLM.from_pretrained("microsoft/DialoGPT-medium")

def predict(input, history=[]):
    # tokenize the new input sentence
    new_user_input_ids = tokenizer.encode(input + tokenizer.eos_token, return_tensors='pt')

    # append the new user input tokens to the chat history
    bot_input_ids = torch.cat([torch.LongTensor(history), new_user_input_ids], dim=-1)

    # generate a response 
    history = model.generate(bot_input_ids, max_length=1000, pad_token_id=tokenizer.eos_token_id).tolist()


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
        # convert the tokens to text, and then split the responses into lines
        response = tokenizer.decode(history[0]).split("<|endoftext|>")
        response = [(response[i], response[i+1]) for i in range(0, len(response)-1, 2)]  # convert to tuples of list
        return response, history

with gr.Blocks() as demo:
    chatbot = gr.Chatbot()
    state = gr.State([])

    with gr.Row():
        txt = gr.Textbox(show_label=False, placeholder="Enter text and press enter").style(container=False)

    txt.submit(predict, [txt, state], [chatbot, state])
            
if __name__ == "__main__":
    demo.launch()




