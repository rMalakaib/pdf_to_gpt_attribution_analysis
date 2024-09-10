import openai
import json
import pandas as pd

# Set up the OpenAI API key
openai.api_key = ""

CATEGORIES = "Name,MOIC"

def format_data_into_table(data):
    responses = []

    # Prepare the prompt
    prompt = "formulate the data into a list delimited by | based off these categories: {CATEGORIES}. DO NOT CHANGE THE CATEGORY NAMES. An example would be: John Doe 3.5x Jane Smith 2.8x TURNS INTO John Doe | 3.5x, Jane Smith | 2.8x \n\n"
    for item in data:
        prompt += f"{item}\n"
    
    for segment in data:
        # Call the ChatGPT API
        
        response = openai.chat.completions.create(
            model="gpt-3.5-turbo",
            messages=[
                {"role": "system", "content": prompt},
                {"role": "user", "content": segment[0]}
            ]
        )
        
        # # Extract the response text
        table = json.loads(response.json())
        responses.append(table["choices"][0]["message"]["content"])

    return responses

# Example usage

# data = "John Doe 3.5x"

data = [
    ["John Doe 3.5x"],
    ["Jane Smith 2.8x"],
    ["Sam Johnson 4.2x"]
]

def formatting(table):

    chart = []

    for row in table:
        item = row.split("|")
        chart.append(item)

    column = CATEGORIES.split(",")

    df = pd.DataFrame(chart,columns=column)


    df.to_csv("gpt.csv", index=False)
            
table = format_data_into_table(data)
formatting(table)

