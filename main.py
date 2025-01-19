import sys
import numpy as np
import pandas as pd

def main():
    # Check if the correct number of arguments are provided
    if len(sys.argv) != 5:
        print("Error: Incorrect number of arguments.")
        print("Usage: python topsis.py <InputDataFile> <Weights> <Impacts> <ResultFileName>")
        sys.exit(1)
    

    # Read arguments from command line
    input_file = sys.argv[1]
    weights = np.array(list(map(int, sys.argv[2].split(","))))
    impacts = sys.argv[3]
    output_file = sys.argv[4]

    # Read the CSV file into a DataFrame
    try:
        df = pd.read_excel(input_file)
    except FileNotFoundError:
        print(f"Error: Input file '{input_file}' not found.")
        sys.exit(1)

    # Store the column names with numeric values
    cols = df.columns[1:]

    # Calculate the square root of the sum of squares for each column
    root_sum_sq = np.sqrt(np.sum(df[cols] ** 2, axis=0))

    # Divide each value in the DataFrame by the root sum of squares
    df[cols] = df[cols] / root_sum_sq

    # Assign the output file name
    output_file = "102203816-result.csv"

    # Normalize the weights
    normalized_weights = weights / np.sum(weights)

    # Multiply each column with the normalized weights
    df[cols] = df[cols] * normalized_weights

    # Split the impacts string into a list
    impacts_list = impacts.split(",")

    # Calculate best and worst values for each column based on impact
    best_values = df[cols].max(axis=0)
    worst_values = df[cols].min(axis=0)


    # Adjust best and worst values based on impact
    for i, impact in enumerate(impacts_list):
        if impact == "-":
            best_values[i], worst_values[i] = worst_values[i], best_values[i]

    
    # Calculate Euclidean distance from best values
    df['best_dist'] = np.sqrt(np.sum((df[cols] - best_values) ** 2, axis=1))

    # Calculate Euclidean distance from worst values
    df['worst_dist'] = np.sqrt(np.sum((df[cols] - worst_values) ** 2, axis=1))

    # Calculate performance score
    df['Topsis Score'] = df['worst_dist'] / (df['best_dist'] + df['worst_dist'])

    # Calculate ranks based on performance score
    df['Rank'] = df['Topsis Score'].rank(ascending=False).astype(int)

    # Save the results to the output CSV file
    df.to_csv(output_file, index=False)



if __name__ == '__main__':
    main()












