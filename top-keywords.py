import pandas as pd

# ------- LOAD EXCEL FILE -------
input_file = '/Users/harshita/Desktop/Work/CodeWork/Email templates/keyword-frequency.xlsx'
df = pd.read_excel(input_file)

# ------- REMOVE EMPTY SEARCH KEYS -------
# Filter out rows where SEARCH_KEY is empty
df = df[df['SEARCH_KEY'].notna() & (df['SEARCH_KEY'].str.strip() != '')]

# ------- USER INPUT FOR TOP N SEARCH KEYS -------
while True:
    try:
        # Ask the user how many top search keys they want
        user_input = input("🔢 Enter the number of top search keys to display (Press Enter for default 20): ")
        
        # Check if user presses Enter without any input
        if not user_input.strip():
            top_n = 20  # Default value
            print("✅ No input provided. Using default value: 20")
            break
        
        # Convert to integer if input is provided
        top_n = int(user_input)
        
        if top_n <= 0:
            print("❗ Please enter a positive number.")
        else:
            break
    except ValueError:
        print("❗ Invalid input. Please enter a valid integer.")

# ------- FILTER TOP N SEARCH KEYS -------
# Get the top N search keys based on FREQUENCY
top_keywords_df = df.head(top_n)

# ------- SAVE TO NEW EXCEL FILE -------
output_file = '/Users/harshita/Desktop/Work/CodeWork/Email templates/top-searchkeys.xlsx'
top_keywords_df.to_excel(output_file, index=False)

print(f"✅ Top {top_n} search keys saved successfully at: {output_file}")

