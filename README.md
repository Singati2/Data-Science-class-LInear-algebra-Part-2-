# Math for Data Science — Linear Algebra, Week 2

Class slides for **Solving Systems of Linear Equations**.

| File | What it is |
|---|---|
| `Linear_Algebra_Week2_Slides.pptx` | The deck (57 slides, 16:9), editable in PowerPoint / Keynote / Google Slides |
| `Linear_Algebra_Week2_Slides.pdf` | The same deck as a PDF for projecting or sharing |
| `Linear_Algebra_Week2_Lecture_Notes.pdf` | Lecture notes: one entry per slide (slide image, what to say, full worked arithmetic) |
| `lecture_notes.py`, `build_notes_pdf.py` | Source of the notes and the script that builds the notes PDF |
| `build_slides.py` | Python source that generates the deck (`python3 build_slides.py`, needs `python-pptx`) |
| `assets_rank_compression.png` | Low-rank image-compression figure used on the rank slide (built from matplotlib's bundled public-domain sample photo) |

## Contents

0. Machine-learning motivation — neural networks as matrix operations; sound recognition, music generation, linear regression
1. Solving systems of linear equations — non-singular, singular (redundant / contradictory), three unknowns
2. Matrix row reduction — coefficient matrix, elementary row operations, effect on the determinant
3. Rank of a matrix — image compression by rank, pieces of information, solution space, rank and singularity
4. Row echelon form — definition, pivots, computing rank
5. Reduced row echelon form — definition and examples
6. Gaussian elimination — augmented matrix, pivoting, back substitution, singular cases

Each concept opens with a formal **definition** box. Three in-class **quizzes** (with solution slides) are built in:
after non-singular systems, after singular systems, and after rank. Short **data science connection** notes link the
definitions to regression, collinearity and numerical solvers.

## Notes

- Adapted from the DeepLearning.AI "Math for Machine Learning" Week 2 material (CC BY-SA), with the
  branding removed, formal definitions added, and the machine-learning motivation rebuilt with original figures.
- The Gaussian-elimination example uses `4a + b = 4` as its third equation (the source deck listed
  `4a + b = −1`, which is inconsistent with the walkthrough it shows; the solution `a = 1, b = 0, c = −1`
  requires the constant 4).
