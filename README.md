# TOPSIS Implementation

This project provides a **TOPSIS (Technique for Order of Preference by Similarity to Ideal Solution)** implementation in Python. The purpose of the algorithm is to rank alternatives based on their similarity to an ideal solution, considering multiple criteria. The user can input a dataset with multiple criteria, apply weights and impacts to each criterion, and receive a ranked output of alternatives.

---

## Features

- Accepts a dataset in CSV format or Excel (`.xlsx`).
- Allows the user to specify weights for each criterion.
- Allows the user to specify whether each criterion has a positive or negative impact.
- Computes the **TOPSIS** score and ranks the alternatives.
- Outputs the results in a new CSV file.

---

## Requirements

Before running the program, ensure that you have the following libraries installed:

- Python 3.x
- **NumPy** (for numerical computations)
- **Pandas** (for data handling and manipulation)

You can install the required libraries using **pip**:

```bash
pip install numpy pandas
