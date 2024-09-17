# Statistics Functions Solver

The **Statistics Functions Solver** is a Python project designed to simplify the process of solving complex statistics problems. It is particularly useful for college-level statistics assignments and calculations.

The main Python file is located inside the `Statistics Question Solver` folder.

## Features

The program includes a variety of functions to perform statistical calculations and operations on data. It supports importing data from `.xlsx` files and directly performing computations on the data.

## Functions

1. **Table Maker**
   - Creates a table in `.xlsx` file format for further operations.
   
2. **Table Reader**
   - Reads an existing table from an `.xlsx` file for analysis.
   
3. **Mean**
   - Computes and prints the mean of the data in a table.
   
4. **Linear Percentile**
   - Calculates the position corresponding to a given percentile and sample size.
   
5. **Table Percentile**
   - Calculates the percentile directly from the table data.
   
6. **Sample Variance**
   - Computes the variance of the data in a table.
   
7. **Sample Standard Deviation**
   - Computes the standard deviation by taking the square root of the sample variance.
   
8. **(Not Updated Yet)**
   
9. **Expected Value**
   - Computes the expected value from the table data, similar to the mean function.
   
10. **Variance**
    - Computes the variance from the table data.
    
11. **Binomial Random Variable**
    - Calculates the binomial random variable, expected value, and variance.
    
12. **Standard Normal Probability**
    - Determines the standardized normal probability using an image with predefined values located in the same directory as the `.py` file.
    
13. **Standard Normal Probability using Central Limit Theorem**
    - Calculates the standardized normal probability using the Central Limit Theorem.
    
14. **(Not Updated Yet)**

## Usage

1. **Table Creation and Reading**
   - Use the **Table Maker** function to create a table if you don’t have one.
   - Use the **Table Reader** function to import and read data from an existing `.xlsx` file.

2. **Statistical Calculations**
   - Call the relevant function (e.g., **Mean**, **Variance**, **Sample Standard Deviation**) to perform statistical operations on the imported table data.

3. **Special Calculations**
   - For probability and distribution-related functions, ensure the required images or data files are available in the same directory as the `.py` file.

## Customization

The program may require customization to suit specific use cases or formats. Modify the code as needed to adapt to different data structures or additional functionality.

## Installation

Ensure you have Python installed. You may also need to install additional packages such as `openpyxl` for handling `.xlsx` files.

To install required packages, you can use:
```bash
pip install openpyxl
