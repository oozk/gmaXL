# General Morphological Analysis in Microsoft Excel
---
### About the project
This project implements a light-weight version of [General Morphological Analysis](https://en.wikipedia.org/wiki/Morphological_analysis_%28problem-solving%29) using Microsoft Excel VBA.

General Morphological Analysis is a method developed by Fritz Zwicky to explore all possible solutions to a multi-dimensional, non-quantified, complex problem. An introduction to the methodology is presented in [this video](https://www.youtube.com/watch?v=x4zAniSP0FY) and a detailed description of the method is available at [this link](https://www.swemorph.com/ma.html).

### Usage
1. Go to the **ReadMe** worksheet.
2. Click on **Maintain Attributes**.
3. In **Attributes** worksheet, maintain a free form table starting in cell A1. Enter the names of all the *attributes* as the table header in the first row, and then maintain the corresponding values (conditions) as rows under each attribute.
4. Once finished, go back to **ReadMe** worksheet and click on **Initialize Cross Consistency Matrix**.
5. Maintain yellow highlighted cells in **CrossConsistency** tab either as ***1***, i.e. valid combination, or as ***-1***, i.e. invalid combination.
6. Once **CrossConsistency** worksheet is fully maintained, go back to **ReadMe** worksheet and click on **Generate Solutions**.
7. The tool will generate all feasible solutions and list them out in **Solutions** tab.
8. You can review and analyze feasible solutions in **Solutions** tab, and further refine the relationship between different attributes repeating the steps as of step 5.

### Microsoft Excel Implementation
In the embedded VBA modules, you will find a series of Class Modules that capture and process the key parameters and the output of the GMA methodology. The critical subset are the following:

| Class Module | Description | Corresponding Worksheet |
| ------ | ------ |------ |
| Attributes | Manages the dimensions of the problem | Reads from *Attributes* worksheet |
| Constraints | Manages the cross-consistency matrix  | Reads from *CrossConsistency* worksheet |
| Solutions | Generates feasible solution space| Output is written to *Solutions* worksheet |
| Worktabs | Manages the tabs in the Excel workbook | Creates, deletes and updates above worksheets |

Three VBA subroutines establish the core algorithm of the methodology:

| Sub | Description | Corresponding Worksheet |
| ------ | ------ |------ |
| S01_Initialize_Attributes() | Initializes the workbook and creates *Attributes* worksheet | *Attributes* |
| S02_Initialize_Matrix() | Creates the cross-consistency matrix based on the values maintained in the *Attributes* worksheet |*CrossConsistency*|
| S03_Generate_Solutions() | Generates all valid solutions to the problem based on the constraints in the *CrossConsistency* worksheet | *Solutions* |

Sub named *TraverseSolutionTree* recursively iterates through the combination of all attribute values to identify feasible combinations as possible valid solutions to the problem.

Pull requests are welcome.

### To Do List
 - Improve project documentation
 - Improve user instructions
 - Implement automatic reflection of Attributes/Condition updates to the populated Cross Consistency Matrix
 - Simplify Cross Consistency Matrix maintenance
 - Improve solution output analysis (pivot tables?)

### License
MIT
