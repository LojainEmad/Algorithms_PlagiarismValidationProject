# Plagiarism Validation

## Project Requirements

### Required Implementation

| **Requirement** | **Performance** |
|-----------------|-----------------|
| **1. Read a html file (or excel) containing N matching pairs.** <br> Each pair consists of file1 path, file2 path, file1 hyperlink, file2 hyperlink, similarity percentage (%) of each file | **Time:** should be **bounded by O(N)**, N is the number of matching pairs |
| **2. Construct the Graph** | **Time:** should be **bounded by O(N)**, N is the number of matching pairs |
| **3. Find ALL groups and their statistics** | **Time:** should be **bounded by O(N)**, N is the number of matching pairs |
| **4. Refine each group, by finding its max spanning tree** | **Time:** should be bounded by:<br> ![formula](https://latex.codecogs.com/svg.latex?O\left(\sum_{c=1}^{\#groups}\left(N_c\log\log\left(N_c\right)+M_c^2\right)\right))<br> Where ![Nc](https://latex.codecogs.com/svg.latex?N_c) and ![Mc](https://latex.codecogs.com/svg.latex?M_c) are the number of pairs & files of each group |
| **5. Output**<br> a. Group statistics file.<br> b. Refined matching pairs file. | |
