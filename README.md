# Student Group Assignment Optimizer

## Overview
This project provides a solution for efficiently and fairly assigning students into groups based on their friendship preferences.

Each student can list up to two preferred friends they'd like to be grouped with.  
The algorithm builds a **friendship graph**, applies **graph theory** and **network flow-inspired techniques**, and distributes students into a predefined number of balanced groups while:

- Maximizing the number of students placed with at least one chosen friend,
- Balancing group sizes as evenly as possible,
- Respecting specific constraints (for example, students who must not be in the same group).

The final result achieves a high success rate in satisfying students' preferences while maintaining fairness and organizational structure.

---

## Technologies Used
- **Python 3**
- **NetworkX** (graph algorithms)
- **Pandas** (data manipulation)
- **OpenPyXL** (Excel file export)
- **Matplotlib** (optional: for visualizing success rates)

---

## How It Works
1. Load students and their friendship selections from an Excel file.
2. Build a friendship graph connecting students based on their selections.
3. Split large components into manageable clusters if necessary.
4. Assign students to groups, prioritizing friend connections and balanced group sizes.
5. Optionally, fine-tune by moving students between groups to further maximize friend satisfaction.
6. Export the final group assignments to a colored Excel file.

---

## Project Structure
- `load_students()`: Reads students and friends from Excel.
- `build_friend_graph()`: Constructs a graph representing student friendships.
- `assign_students_to_groups()`: Core logic for group assignment based on friendships and balancing.
- `post_friend_balancing()`: Improves assignments by rebalancing friends post-assignment.
- `export_assignments_to_excel()`: Saves the final results into a formatted Excel sheet.

---

## Usage
1. Prepare an Excel file with:
    - A sheet named **"students"** containing:
        - `student`, `friend1`, `friend2` columns.
    - (Optional) A sheet named **"constraints"** with pairs of students who must not be grouped together.
2. Update the filename and sheet names in the code if needed.
3. Run the main script to generate group assignments.
4. Review the generated Excel output with final groupings.

---

## Notes
- The project is flexible and can easily be adapted to support different numbers of groups, students, or friend preferences.
- No personal data is exposed — results can be exported using pseudonyms if needed.

---

## License
This project is open-source and free to use.

