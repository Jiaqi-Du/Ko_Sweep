# Ko_Sweep

Combine displacement, area, and vertical efficiency to get the sweep efficiency. Use the Koval Function to fit the results and find the best Koval value.

The Koval Function is defined as:

$$
F(C) = \frac{K_o C}{C(K_o - 1) + 1}
$$

## Installation and Use

- **1. Package Installation**

   To install the required packages, run the following:

   ```bash
   pip install numpy pandas matplotlib scipy scikit-learn openpyxl

- **2. Input Data for Dataset.xlsx**
   

   **KoD:** Ko value for displacement efficiency

   **KoA:** Ko value for area sweep efficiency

   **KoV:** Ko value for vertical sweep efficiency
   
   **&#916; S:** The flowing fluid, calculated as $\ 1 - S_{or} - S_{wr}$ where:
  
  &nbsp;&nbsp;&nbsp;&nbsp; S<sub>or</sub>: residual oil saturation
  
  &nbsp;&nbsp;&nbsp;&nbsp; S<sub>wr</sub>: irreducible water saturation

  **tD:** dimensionless time, 1 tD means the volume of injected fluid is equal to the pore volume
