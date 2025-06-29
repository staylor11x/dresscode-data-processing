import os
import pandas as pd
import openai
import time

# Load your OpenAI API key
openai.api_key = os.getenv("OPENAI_API_KEY")

# Prompt template for few-shot classification
PROMPT_TEMPLATE = """You are a helpful assistant that classifies education-related labels.

Here is a list of terms. For each, respond with either "Subject" or "Not a Subject".

Example:
- Mathematics: Subject
- Total: Not a Subject
- Chemistry: Subject
- All Candidates: Not a Subject

Now classify the following:

{terms}

Respond in this format:
- Term: Label
"""

def batch_terms(terms, batch_size=20):
    """Yield batches of subject names."""
    for i in range(0, len(terms), batch_size):
        yield terms[i:i + batch_size]

def classify_batch(batch):
    term_block = "\n".join(batch)
    prompt = PROMPT_TEMPLATE.format(terms=term_block)
    
    try:
        response = openai.chat.completions.create(
            model="gpt-3.5-turbo",  # You can downgrade to "gpt-3.5-turbo" if needed
            messages=[
                {"role": "system", "content": "You are a subject classification assistant."},
                {"role": "user", "content": prompt}
            ],
            temperature=0
        )
        return response.choices[0].message.content

    except openai.error.OpenAIError as e:
        print(f"Error: {e}")
        return None

def parse_response(text):
    results = []
    for line in text.splitlines():
        if ':' in line:
            try:
                term, label = line.split(':', 1)
                results.append((term.strip(), label.strip()))
            except ValueError:
                continue
    return results

def main():
    input_file = "subject_names.csv"
    output_file = "subject_labels.csv"

    df = pd.read_csv(input_file)
    terms = df["Subject"].dropna().unique().tolist()

    final_results = []

    for batch in batch_terms(terms):
        print(f"🔍 Classifying batch of {len(batch)} terms...")
        result_text = classify_batch(batch)
        if result_text:
            parsed = parse_response(result_text)
            final_results.extend(parsed)
        time.sleep(1.5)  # Be kind to the API

    result_df = pd.DataFrame(final_results, columns=["Subject", "Label"])
    result_df.to_csv(output_file, index=False)
    print(f"✅ Labels saved to {output_file}")

if __name__ == "__main__":
    main()
