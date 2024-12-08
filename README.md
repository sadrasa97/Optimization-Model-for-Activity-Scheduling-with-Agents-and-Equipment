# Optimization Model for Activity Scheduling with Agents and Equipment

This project defines and solves an optimization problem for assigning agents and equipment to various activities while maximizing overall quality and minimizing costs, using **Pyomo**. The project also includes data generation and visualization capabilities.

---

## Table of Contents
1. [Features](#features)
2. [Getting Started](#getting-started)
   - [Prerequisites](#prerequisites)
   - [Installation](#installation)
   - [Usage](#usage)
3. [Project Structure](#project-structure)
4. [Input Data Description](#input-data-description)
5. [Optimization Problem](#optimization-problem)
6. [Results](#results)
7. [License](#license)

---

## Features

- **Data Generation**: Randomly generate synthetic input data for testing.
- **Data Loading**: Load multi-scenario data from Excel files.
- **Optimization**: Use Pyomo to solve a complex assignment problem involving agents, equipment, and activities across multiple scenarios.
- **Output Results**: Retrieve and visualize results for decision-making.

---

## Getting Started

### Prerequisites

Ensure you have the following installed:
- Python (>= 3.7)
- Required libraries: `pandas`, `numpy`, `pyomo`, `xlsxwriter`

Install missing dependencies using:

```bash
pip install pandas numpy pyomo xlsxwriter
```

### Installation

1. Clone this repository:
   ```bash
   git clone https://github.com/username/optimization-scheduling.git
   ```
2. Navigate to the project directory:
   ```bash
   cd optimization-scheduling
   ```

---

## Usage

### Step 1: Generate Input Data
Run the following script to generate synthetic input data and save it to an Excel file:
```bash
python generate_data.py
```
This creates a file named `Generated_data.xlsx` with the required input data.

### Step 2: Load Input Data
The script includes functionality to load data from the generated Excel file into a structured dictionary.

### Step 3: Solve the Optimization Model
Run the main script to solve the optimization problem:
```bash
python solve_model.py
```
This will output the agent assignments, equipment allocations, and time allocations to the console.

---

## Project Structure

```
├── generate_data.py     # Script for generating synthetic input data
├── solve_model.py       # Script for defining and solving the optimization model
├── requirements.txt     # Python package dependencies
└── README.md            # Project documentation
```

---

## Input Data Description

The input data consists of the following:

| Parameter      | Description                                             |
|----------------|---------------------------------------------------------|
| `Dh`           | Importance of each activity.                           |
| `Tmin`, `Tmax` | Minimum and maximum time required for activities.       |
| `Kh`           | Fixed daily cost for activities.                       |
| `Bhj`          | Quality scores for agents for each activity.           |
| `Rhe`          | Quality scores for equipment for each activity.        |
| `Wij`          | Collaboration factor between agents.                   |
| `G`            | Fixed equipment cost per activity and scenario.        |
| `Beta`         | Variable cost factor per activity, equipment, and scenario. |

---

## Optimization Problem

The optimization model seeks to:
1. Maximize **quality** scores for agents and equipment assignments.
2. Minimize **fixed and variable costs**.
3. Maximize **collaboration scores** between agents.

### Decision Variables:
- `X[h, j, s]`: Binary variable for assigning agent `j` to activity `h` in scenario `s`.
- `Y[h, e, s]`: Binary variable for assigning equipment `e` to activity `h` in scenario `s`.
- `T[h, s]`: Time allocated for activity `h` in scenario `s`.
- `Z[h, i, j, s]`: Collaboration binary variable for agents `i` and `j` in activity `h`, scenario `s`.
- `Aux[h, e, s]`: Auxiliary variable for cost calculation.

---

## Results

After solving the model, the following outputs are generated:
1. **Agent Assignments**: Which agent is assigned to each activity in each scenario.
2. **Equipment Assignments**: Which equipment is assigned to each activity in each scenario.
3. **Time Allocations**: Time allocated for each activity in each scenario.

Results are displayed in the console for quick review.

---

## License

This project is licensed under the MIT License. See the LICENSE file for details.

---

Feel free to contribute by opening issues or submitting pull requests! 🚀