# Ko_Sweep

Combine displacement, area, and vertical efficiency to get the sweep efficiency. Use the Koval Function to fit the results and find the best Koval value.

The Koval Function is defined as:

$$
F(C) = \frac{K_o C}{C(K_o - 1) + 1}
$$

## Usage

1. **Package Installation**

   To install the required packages, run:

   ```bash
   pip install numpy pandas matplotlib scipy scikit-learn openpyxl

2. **Input Data**
   
   Input Ko value for displacement, vertical, and horizontal efficiency in corresponding columns in 'Dataset.xlsx'
