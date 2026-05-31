from transformers import pipeline

print("Loading Qwen...")

llm = pipeline(

    "text-generation",

    model="Qwen/Qwen2.5-3B-Instruct",

    device_map="auto"

)

print("Qwen Loaded")


def ask_qwen(query,columns):

    prompt=f"""

You convert user queries
into dataframe instructions.

Columns:

{columns}

Allowed outputs:

FILTER column > value

FILTER column < value

FILTER column = value

COUNT

AVERAGE column

TOP n

Examples:

days required more than 20

FILTER Days Required > 20

count rows

COUNT

average progress

AVERAGE Progress

top 5

TOP 5

User:

{query}

Return ONLY instruction.

"""

    messages=[

        {

            "role":"user",

            "content":prompt

        }

    ]

    output=llm(

        messages,

        max_new_tokens=40

    )

    return output[0][
        "generated_text"
    ][-1][
        "content"
    ]