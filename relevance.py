import os
from typing import Dict, List, Tuple, Any
from openai import AzureOpenAI
from dotenv import load_dotenv
from docx import Document
import pandas as pd
import requests
import json
import sys
import re
load_dotenv()

client = AzureOpenAI(
    api_version="2024-02-01",
    azure_endpoint="https://ai-pacepivotai486018497712.openai.azure.com/openai/deployments/gpt-4o/chat/completions?api-version=2023-03-15-preview",
    api_key=os.environ.get("AZURE_OPENAI_API_KEY")
)

def import_data(file_path):
    try:
        # Attempt to read the Excel file
        df = pd.read_excel(file_path, engine="openpyxl")
        keywords = df.iloc[:, 0].tolist()
        return keywords
    except ImportError as e:
        print(f"Error: {e}")
        print("Please install the 'openpyxl' library using 'pip install openpyxl'")
        return None
    except Exception as e:
        print(f"Error reading file: {e}")
        return None
    

def fetch_asin_data(asin: str, region: str):
        url = "https://parazun-amazon-data.p.rapidapi.com/product/"
        querystring = {"asin": asin, "region": region}
        headers = {
            "x-rapidapi-key": os.environ.get("RAPIDAPI_KEY"),
            "x-rapidapi-host": "parazun-amazon-data.p.rapidapi.com"
        }
        response = requests.get(url, headers=headers, params=querystring)
        data = response.json()

        title = data["title"]
        features = data.get("features", [])
        images = data["images"]
        subtitle = data.get("subtitle", "")
        brand = data.get("brand", "")
        answered_questions = data.get("answered_questions", "")
        description = data.get("description", "")
        overview = data.get("overview", "")
        desc = f"{brand} {title} {subtitle} {features} {description} {overview}"
        return desc

def relevance_check(keywords: List[str], desc: str) -> pd.DataFrame:
    prompt = f"""You are a keyword relevance checker with a deep understanding of customer search patterns on Amazon. You are provided a list of keywords: **{keywords}** and a product description: **{desc}** which contains the brand, title, subtitle, features, description, and overview of the product.

Your tasks are as follows:
1. Extract and accurately classify each keyword into one of three categories:
   - **Brand:** Keywords related to brand names.
   - **Shop Intent:** Keywords indicating an intent to purchase or shop.
   - **Browse Intent:** Keywords indicating an intent to browse general information or options.
2. Assign a relevance score to each keyword, based on its relevance to the product (0 being not relevant at all, 5 being highly relevant).
3. Group the keywords according to search intent cohorts.

The output should be in JSON format with the following format for each keyword:
- **Keyword:** The keyword itself.
- **Relevance Score:** A numerical score from 0 to 10.
- **Category:** One of the three categories (Brand, Shop Intent, Browse Intent).
- **Search Intent Cohort Classification:** 
"""

    try:
        # Send request to OpenAI
        response = client.chat.completions.create(
            model="gpt-4",
            messages=[{"role": "system", "content": prompt}],
            response_format={ "type": "json_object" }, 
            temperature=0.0,
        )

        # Extract the content
        relevant = response.choices[0].message.content.strip()

        # Print the relevant content for debugging
        print("Relevant Content from OpenAI:")
        print(relevant)
    except Exception as e:
        print(f"Error: {e}")
        return None
        
    return relevant

def clean_to_df(relevant_content: str) -> pd.DataFrame:
    try:
        # Clean the relevant content and ensure proper structure
        relevant_content = relevant_content.replace("**", "")
        relevant_content = re.sub(r"(\w+):", r'"\1":', relevant_content)  # Add quotes around keys if missing
        relevant_content = re.sub(r",(\s*)}", r"}", relevant_content)  # Remove trailing commas before closing braces
        relevant_content = f'[{relevant_content}]'  # Convert to JSON array if necessary

        # Parse the cleaned JSON content
        keyword_data = json.loads(relevant_content)

        # If keyword_data contains a list of dictionaries inside "keywords", we need to extract those
        if isinstance(keyword_data, list) and 'keywords' in keyword_data[0]:
            # Extract the 'keywords' column (which contains the nested data)
            keywords_data = [item['keywords'] for item in keyword_data]
            
            # Flatten the nested list of dictionaries into individual rows
            flat_keywords_data = [keyword for sublist in keywords_data for keyword in sublist]

            # Normalize the flattened data to create the DataFrame
            df = pd.json_normalize(flat_keywords_data)
        else:
            # If the data isn't nested, normalize the content directly
            df = pd.json_normalize(keyword_data)

        # Ensure columns are properly structured
        if "Keyword" not in df.columns:
            df['Keyword'] = df.apply(lambda row: row.get('Keyword', ''), axis=1)
        if "Relevance Score" not in df.columns:
            df['Relevance Score'] = df.apply(lambda row: row.get('Relevance Score', 0), axis=1)
        if "Category" not in df.columns:
            df['Category'] = df.apply(lambda row: row.get('Category', ''), axis=1)
        if "Search Intent Cohort Classification" not in df.columns:
            df['Search Intent Cohort Classification'] = df.apply(lambda row: row.get('Search Intent Cohort Classification', ''), axis=1)

        print("Constructed DataFrame:")
        print(df)
        return df

    except (json.JSONDecodeError, ValueError) as e:
        print(f"Failed to parse and convert to DataFrame: {e}")
        return pd.DataFrame(columns=['Keyword', 'Relevance Score', 'Category', 'Search Intent Cohort Classification'])


if __name__ == "__main__":
    asin = "B009GCTZWC"
    region = "IN"
    file_path = "/mnt/c/Users/arun/Downloads/final search terms/hindi alphabet.xlsx"
    desc = fetch_asin_data(asin, region)    
    kw = import_data(file_path)
    rel = relevance_check(kw, desc)
    rel_df=clean_to_df(rel)
    rel_df.to_excel("relevance.xlsx", index=False)
    




