"""Lecture notes, one entry per slide of Linear_Algebra_Week2_Slides.pptx.

NOTES[n] = (short title, list of paragraphs). Paragraphs may use <b>, <i>.
Used by build_notes_pdf.py (handout) and build_slides.py (speaker notes).
"""

NOTES = {
1: ("Cover", [
    "Welcome back. Week 1 introduced matrices, singular and non-singular systems, and the determinant. This week we learn to actually "
    "<b>solve</b> systems of linear equations, and to turn that solution procedure into a machine-friendly algorithm on matrices.",
    "By the end of the session students should be able to: solve a square system by elimination; recognise redundant and contradictory "
    "systems; carry out row reduction on a matrix; define and compute rank, row echelon form and reduced row echelon form; and run "
    "Gaussian elimination on an augmented matrix."]),
2: ("Roadmap", [
    "Walk through the six topics in order. The first half (topics 1 to 3) is procedural: how to solve a system and why the moves are legal. "
    "The second half (topics 4 to 6) is conceptual: rank measures information, echelon forms are the standardised end product of "
    "elimination, and Gaussian elimination packages everything into one algorithm.",
    "Mention the three quizzes so students know pauses are coming. Each quiz is two or three minutes of individual work followed by a "
    "solution slide."]),
3: ("Motivation divider", [
    "Before the mathematics, spend two or three minutes on why a data scientist cares. The one-sentence version: every model you will "
    "fit stores its knowledge in matrices, and fitting the model means solving for the entries of those matrices."]),
4: ("Neural networks are matrix operations", [
    "A neural network takes an input vector (pixel values, audio samples, tabular features), multiplies it by a weight matrix W<sub>1</sub>, "
    "adds a bias and applies a simple non-linearity, then repeats with W<sub>2</sub>, W<sub>3</sub>, and so on until a prediction comes out.",
    "Two take-aways for the class. First, <b>using</b> a trained network is nothing but repeated matrix multiplication, so matrices are the "
    "data structure of machine learning. Second, <b>training</b> a network means finding the entries of these matrices, which is a very large "
    "system of equations solved approximately. The tools of this week (solving systems, row reduction, rank) are the exact versions of the "
    "ideas that training algorithms approximate."]),
5: ("Where this shows up in data science", [
    "Three concrete applications. <b>Sound recognition</b>: hydrophones record underwater sound; a network classifies which species is "
    "present, which lets ecologists monitor habitats without disturbing them. <b>Music generation</b>: a model first compresses audio into "
    "discrete codes, then learns the patterns of a genre well enough to produce new pieces. Both are neural networks, so both are stacks of "
    "matrices.",
    "<b>Linear regression</b> is the simplest model and the most direct link to today. Minimising squared error leads to the normal "
    "equations X<sup>T</sup>X β = X<sup>T</sup>y, a system of linear equations in the unknown coefficients β. Whether that system has one "
    "solution, infinitely many, or none is exactly the question we answer this week."]),
6: ("Part 1 divider", [
    "Part 1 works entirely with equations, no matrices yet. The goal is to make the elimination procedure feel natural before we encode it."]),
7: ("Linear systems: definition", [
    "Read the definition slowly. A <b>linear equation</b> is a weighted sum of the variables set equal to a constant; the key restriction is "
    "that variables appear only to the first power and never multiplied together. So 3a − b + 2c = 7 is linear while ab = 1 or a² = 4 is not.",
    "A <b>system</b> is several such equations about the same unknowns, and a <b>solution</b> must satisfy all of them simultaneously. "
    "Point out the two examples on the right: the 2×2 system whose solution is (8, 2), and the 3×3 system we will solve later.",
    "Data science connection: the normal equations of least squares are a linear system of exactly this form, with the regression "
    "coefficients as the unknowns."]),
8: ("A system and its solution", [
    "Use the word problem first: an apple and a banana cost $10, an apple and two bananas cost $12. Most students will see immediately "
    "that the extra banana costs $2, hence the apple costs $8. The system is a + b = 10, a + 2b = 12 and the solved system is a = 8, b = 2.",
    "The slide names the three kinds of manipulation we are allowed: swapping equations, adding equations, and multiplying an equation by a "
    "constant. The whole of elimination is a disciplined way of applying these three moves."]),
9: ("Equivalent systems and the legal operations", [
    "Two systems are <b>equivalent</b> when they have the same solution set. The three operations (swap, scale by a non-zero constant, add a "
    "multiple of one equation to another) are legal precisely because each is reversible, so no solution is gained or lost.",
    "Work the two mini examples: multiplying a + b = 10 by 7 gives 7a + 7b = 70, which is the same line; adding a + b = 10 and "
    "2a + 3b = 22 gives 3a + 4b = 32, which every solution of the original pair also satisfies. Ask why scaling by zero is forbidden: it "
    "would erase the equation and could not be undone."]),
10: ("Worked example: 2×2", [
    "System: 5a + b = 17 and 4a − 3b = 6. Step 1, divide each equation by its coefficient of a: a + 0.2b = 3.4 and a − 0.75b = 1.5. "
    "Step 2, subtract the first from the second: −0.95b = −1.9, so b = 2. Step 3, back-substitute: a + 0.2(2) = 3.4 gives a = 3.",
    "Emphasise the pattern, because it is the whole algorithm: make the leading coefficients match, subtract to eliminate a variable, solve "
    "the smaller system, substitute back. Check: 5(3) + 2 = 17 and 4(3) − 3(2) = 6."]),
11: ("A coefficient that is already zero", [
    "System: 5a + b = 17 and 3b = 6. There is no a to eliminate from the second equation because its coefficient is already 0. "
    "So b = 2 immediately, then a + 0.2(2) = 3.4 gives a = 3.",
    "The lesson: zeros in the right places are the <i>goal</i> of elimination. When we move to matrices, a zero below the diagonal means "
    "one less step of work."]),
12: ("Quiz 1", [
    "Give two to three minutes. System: 2a + 5b = 46 and 8a + b = 32. Encourage students to pick whichever variable is easier to eliminate; "
    "b is convenient here because its coefficient in the second equation is 1."]),
13: ("Quiz 1 solution", [
    "Multiply the second equation by 5: 40a + 5b = 160. Subtract the first: 38a = 114, so a = 3. Back-substitute: 8(3) + b = 32 gives b = 8. "
    "Check both equations: 2(3) + 5(8) = 46 and 8(3) + 8 = 32.",
    "Students who substituted b = 32 − 8a into the first equation reach the same answer. Point out that substitution and elimination are "
    "the same operations in a different order."]),
14: ("Singular system: redundant", [
    "System: a + b = 10 and 2a + 2b = 20. Dividing the second by 2 gives a + b = 10 twice; subtracting yields 0 = 0. The second equation "
    "carried no new information, so only one constraint remains for two unknowns.",
    "Describe the solutions: choose any a = x, then b = 10 − x. One free parameter, so infinitely many solutions. Connect to Week 1: this "
    "is the singular case with determinant 1·2 − 1·2 = 0."]),
15: ("Singular system: contradictory", [
    "System: a + b = 10 and 2a + 2b = 24. Dividing gives a + b = 12, and subtracting the first equation produces 0 = 2, which is false. "
    "No pair (a, b) can satisfy both equations.",
    "Geometrically the two equations are parallel lines. Note that the coefficient matrix is the same as on the previous slide; only the "
    "constants differ. Singularity is a property of the coefficients; whether the singular system has many or no solutions depends on the "
    "constants."]),
16: ("Classifying systems", [
    "Collect the vocabulary. <b>Consistent</b>: at least one solution. <b>Inconsistent</b>: none. For square systems, <b>non-singular</b> means "
    "exactly one solution and <b>singular</b> covers both the redundant case (infinitely many) and the contradictory case (none).",
    "The three cards are the three pictures to keep in mind: two lines crossing at a point; two descriptions of the same line; two parallel "
    "lines."]),
17: ("Quiz 2", [
    "System: 5a + b = 11 and 10a + 2b = 22. Ask students to look before computing. Two minutes."]),
18: ("Quiz 2 solution", [
    "The second equation is exactly twice the first, so after dividing by 2 and subtracting we get 0 = 0. The system is singular and "
    "redundant: infinitely many solutions, described by b = 11 − 5a for any real a. Give examples: (0, 11), (1, 6), (2, 1).",
    "If someone computed the determinant, 5·2 − 1·10 = 0 confirms singularity, but the determinant alone cannot distinguish redundant from "
    "contradictory; for that we need the constants."]),
19: ("Three variables, step 1", [
    "System: a + b + 2c = 12, 3a − 3b − c = 3, 2a − b + 6c = 24. Divide each equation by its coefficient of a: a + b + 2c = 12, "
    "a − b − ⅓c = 1, a − ½b + 3c = 12. Subtract the first from the other two: −2b − 7⁄3 c = −11 and −3⁄2 b + c = 0.",
    "The variable a now appears only in the first equation. What remains below it is a 2×2 system in b and c, which we already know how "
    "to solve. This is the recursive heart of elimination."]),
20: ("Three variables, step 2", [
    "Divide the last two equations by their coefficient of b: b + 7⁄6 c = 11⁄2 and b − 2⁄3 c = 0. Subtract the second from the third: "
    "(−2⁄3 − 7⁄6)c = −11⁄2, that is −11⁄6 c = −11⁄2, so c = 3.",
    "The system is now triangular: the first equation involves a, b, c; the second b, c; the third c alone. Triangular systems are solved by "
    "reading from the bottom up."]),
21: ("Three variables, step 3: back substitution", [
    "From c = 3, the second equation gives b + 7⁄6·3 = 11⁄2, that is b + 7⁄2 = 11⁄2, so b = 2. The first gives a + 2 + 6 = 12, so a = 4. "
    "Solution (a, b, c) = (4, 2, 3).",
    "Check all three original equations: 4 + 2 + 6 = 12; 12 − 6 − 3 = 3; 8 − 2 + 18 = 24. Checking is cheap and catches arithmetic slips."]),
22: ("Part 2 divider", [
    "We now strip away the variable names and work on the numbers alone. Nothing new happens mathematically; the notation becomes compact "
    "enough for a computer."]),
23: ("From systems to matrices", [
    "The <b>coefficient matrix</b> has one row per equation and one column per variable. For 5a + b = 17, 4a − 3b = 6 the matrix is "
    "[5 1; 4 −3]. The elimination steps from the earlier slide turn it into [1 0.2; 0 1] (upper triangular, called <b>row echelon form</b>) "
    "and finally into the identity [1 0; 0 1] (<b>reduced row echelon form</b>).",
    "Read the three matrices against the three systems above them so students see that the matrix is just bookkeeping for the equations. "
    "The constants are left out here; they return with the augmented matrix in Part 6."]),
24: ("Singular systems as matrices", [
    "The redundant system a + b = 10, 2a + 2b = 20 has matrix [1 1; 2 2]; elimination gives [1 1; 0 0]. The quiz system 5a + b = 11, "
    "10a + 2b = 22 gives [5 1; 10 2] and then [1 0.2; 0 0]. The all-zero system stays all zero.",
    "Highlight the rule: a singular matrix produces a row of zeros during elimination. That row is the matrix version of the equation 0 = 0 "
    "(or 0 = something, once constants are attached)."]),
25: ("Elementary row operations", [
    "Define the three <b>elementary row operations</b>: swap R<sub>i</sub> ↔ R<sub>j</sub>; scale R<sub>i</sub> ← k·R<sub>i</sub> with "
    "k ≠ 0; replace R<sub>i</sub> ← R<sub>i</sub> + k·R<sub>j</sub>. They are the three legal moves on equations, rewritten for rows.",
    "State the theorem: these operations do not change the solution set and do not change singularity. The reason is reversibility; each "
    "operation has an inverse of the same type (swap again, scale by 1/k, subtract k·R<sub>j</sub>)."]),
26: ("Row operations and the determinant", [
    "Start from [5 1; 4 3] with determinant 15 − 4 = 11. Swapping the rows gives [4 3; 5 1] with determinant 4 − 15 = −11: the sign flips. "
    "Multiplying row 1 by 10 gives [50 10; 4 3] with determinant 150 − 40 = 110 = 10·11. Adding row 1 to row 2 gives [5 1; 9 4] with "
    "determinant 20 − 9 = 11: unchanged.",
    "The point for this week: a zero determinant stays zero and a non-zero one stays non-zero under all three operations, which is the "
    "determinant's way of saying that row operations preserve singularity."]),
27: ("Part 3 divider", [
    "Rank is the first genuinely new concept. Motivate it with the question: when do two equations really give two pieces of information?"]),
28: ("Compressing images: reducing rank", [
    "A grayscale image is a matrix of pixel intensities. This one is 300 × 256 and has rank 256. A rank-k approximation keeps only the k "
    "most important 'pieces of information' and rebuilds the picture from k row-patterns and k column-patterns.",
    "Walk along the panel: rank 1 is horizontal and vertical stripes; rank 5 shows a silhouette; rank 15 is recognisable; rank 50 is close to "
    "the original using a fraction of the numbers. The method behind it (the singular value decomposition) comes later in the course, but "
    "the concept of rank as information content is what we define now. The same idea underlies principal component analysis and "
    "recommender systems."]),
29: ("Systems of information", [
    "A non-mathematical warm-up. 'The dog is black. The cat is orange.' is two sentences and two facts: rank 2. 'The dog is black' twice is "
    "two sentences but one fact: rank 1. 'The dog. The dog.' says nothing: rank 0.",
    "Rank counts independent pieces of information, not the number of statements. Keep this slogan; every later definition of rank is a "
    "precise version of it."]),
30: ("Rank of a 2×2 system", [
    "Translate the sentences into equations. a + b = 0 and a + 2b = 0 are two independent constraints: rank 2. a + b = 0 and 2a + 2b = 0 "
    "say the same thing: rank 1. 0a + 0b = 0 twice says nothing: rank 0.",
    "The rank of the matrix is defined as the rank of the system it represents. Students should notice that the rank-1 and rank-0 matrices "
    "are exactly the singular ones from Week 1."]),
31: ("Rank and the solution space", [
    "For a homogeneous system (all constants 0) the <b>solution space</b> is the set of all solutions. Rank 2: only (0, 0) works, a single "
    "point, dimension 0. Rank 1: every point on the line b = −a works, dimension 1. Rank 0: every point in the plane works, dimension 2.",
    "The formula: rank = number of variables − dimension of the solution space, here rank = 2 − dim. This is a first glimpse of the "
    "rank–nullity theorem. Draw the three pictures on the board if students have not met 'dimension' informally."]),
32: ("Rank and singularity", [
    "Put the two ideas together: an n × n matrix is non-singular exactly when its rank is n (<b>full rank</b>). Any smaller rank means a "
    "row that carries no new information, hence a zero row after elimination, hence a zero determinant.",
    "Data science connection: in regression, if two feature columns are proportional (perfectly collinear), X<sup>T</sup>X loses rank, "
    "becomes singular, and the least-squares coefficients are not uniquely determined. Software then warns or drops a column."]),
33: ("Quiz 3", [
    "Matrix 1 is [5 1; −1 3]; Matrix 2 is [2 −1; −6 3]. Ask for the rank of each, using the solution space of the homogeneous system or any "
    "other method. Three minutes."]),
34: ("Quiz 3 solution", [
    "Matrix 1: the system 5a + b = 0, −a + 3b = 0 gives b = −5a, then −a − 15a = −16a = 0, so a = b = 0. Solution space is a point, "
    "dimension 0, rank 2 − 0 = 2. Non-singular, and indeed det = 15 + 1 = 16.",
    "Matrix 2: the second row is −3 times the first, so both equations say 2a − b = 0. Solution space is the line b = 2a, dimension 1, "
    "rank 1. Singular, det = 6 − 6 = 0."]),
35: ("Rank of a 3×3 system", [
    "Four systems with three equations each. System 1 (a + b + c, a + 2b + c, a + b + 2c, all = 0) has three independent equations: rank 3. "
    "System 2 (a + b + c, a + b + 2c, a + b + 3c) has only two independent ones, since the third is 2·(second) − (first): rank 2. "
    "System 3 (rows proportional to 1, 1, 1) has one: rank 1. The zero system has rank 0.",
    "Counting independent equations by eye is already awkward at 3×3. That is the cue for the next slide."]),
36: ("An easier way to compute the rank", [
    "Answer: bring the matrix to row echelon form and count the pivots. Equivalently, count the ones on the diagonal of the reduced row "
    "echelon form. The next two parts define these forms precisely and show that elimination always reaches them."]),
37: ("Part 4 divider", [
    "Row echelon form is the standard shape that elimination produces. Once students can recognise it, rank becomes a counting exercise."]),
38: ("Row echelon form: definition", [
    "A matrix is in <b>row echelon form</b> if all zero rows are at the bottom and, in every non-zero row, the first non-zero entry (the "
    "<b>pivot</b>) sits strictly to the right of the pivot in the row above. The pivots therefore form a staircase descending to the right. "
    "The <b>rank</b> is the number of pivots.",
    "Point at the two 5×5 patterns: pivots 2, 1, 3, −5, 1 on the diagonal give rank 5; pivots 3, 1, −4 followed by two zero rows give rank 3. "
    "Note: pivots need not equal 1 in general; in this class we normalise them to 1 for tidiness, and this changes nothing."]),
39: ("Reaching row echelon form, step by step", [
    "Matrix [5 1; 4 −3]. Step 1: divide each row by its leftmost coefficient to get [1 0.2; 1 −0.75]. Step 2: subtract row 1 from row 2 "
    "to get [1 0.2; 0 −0.95]. Step 3: divide row 2 by −0.95 to get [1 0.2; 0 1], which is in row echelon form with two pivots.",
    "These are the same three moves as the 2×2 worked example in Part 1, now written on rows. Students should be able to say which "
    "elementary row operation each step is."]),
40: ("Row echelon form, singularity and rank", [
    "Three matrices and their echelon forms. [5 1; 4 −3] → [1 0.2; 0 1]: two pivots, rank 2, non-singular. [5 1; 10 2] → [1 0.2; 0 0]: one "
    "pivot, rank 1, singular. [0 0; 0 0] stays zero: no pivots, rank 0, singular.",
    "Full rank means a pivot in every row; a missing pivot means a zero row, which is the fingerprint of a singular matrix."]),
41: ("Row echelon form of a 3×3 system", [
    "The 3×3 system from Part 1, now as the matrix [1 1 2; 3 −3 −1; 2 −1 6]. Operations: R<sub>2</sub> ← R<sub>2</sub> − 3R<sub>1</sub> gives "
    "[0 −6 −7]; R<sub>3</sub> ← R<sub>3</sub> − 2R<sub>1</sub> gives [0 −3 2]; then R<sub>3</sub> ← 2R<sub>3</sub> − R<sub>2</sub> gives "
    "[0 0 11]. Echelon form [1 1 2; 0 −6 −7; 0 0 11], three pivots, rank 3.",
    "Carrying the constants along (12, 3, 24) gives 12, −33, 33 and the triangular system a + b + 2c = 12, −6b − 7c = −33, 11c = 33, so "
    "c = 3 as before. The last operation is two legal moves in one (scale R<sub>3</sub> by 2, then subtract R<sub>2</sub>); it avoids fractions."]),
42: ("Normalizing the pivots", [
    "Dividing each non-zero row by its pivot (here ÷3, ÷1, ÷(−4)) is a scaling operation, so it is legal. The staircase, the number of "
    "pivots and the rank are unchanged; only the convention changes. We adopt pivots equal to 1 because it makes the next step, reduced "
    "row echelon form, cleaner."]),
43: ("More examples", [
    "[1 1 1; 1 2 1; 1 1 2]: subtract row 1 from rows 2 and 3 to get [1 1 1; 0 1 0; 0 0 1], rank 3, non-singular. "
    "[1 1 1; 1 1 2; 1 1 3]: the same step gives [1 1 1; 0 0 1; 0 0 2]; then R<sub>3</sub> ← R<sub>3</sub> − 2R<sub>2</sub> gives a zero "
    "row, so rank 2. [1 1 1; 2 2 2; 3 3 3]: subtracting 2R<sub>1</sub> and 3R<sub>1</sub> leaves two zero rows, rank 1.",
    "The middle example shows that one round of elimination may not be enough; keep going until the staircase condition holds in every row."]),
44: ("Pivots count the rank", [
    "The four 3×3 matrices from the rank slide, each with its row echelon form and highlighted pivots: 3, 2, 1 and 0 pivots, matching the "
    "ranks 3, 2, 1, 0 we found by counting independent equations. This closes the loop: rank by information content and rank by pivot count "
    "agree."]),
45: ("Part 5 divider", [
    "Reduced row echelon form continues elimination upwards. It is the form in which the solution can be read off directly."]),
46: ("Reduced row echelon form: definition", [
    "A matrix is in <b>reduced row echelon form</b> if it is in row echelon form, every pivot equals 1, and every pivot is the only non-zero "
    "entry in its column (zeros above as well as below). Unlike row echelon form, this form is <b>unique</b> for each matrix.",
    "For a non-singular square matrix the reduced row echelon form is the identity. In the 5×5 rank-3 pattern, the non-pivot columns keep "
    "their entries; those columns correspond to free variables of the system."]),
47: ("Reduced row echelon form: 2×2", [
    "From [1 0.2; 0 1], the only entry above a pivot is 0.2. Apply R<sub>1</sub> ← R<sub>1</sub> − 0.2R<sub>2</sub>: 0.2 × [0 1] = [0 0.2] "
    "and [1 0.2] − [0 0.2] = [1 0]. Result [1 0; 0 1], the identity.",
    "In the language of equations, this is substituting b = 2 back into a + 0.2b = 3.4 to isolate a. Back substitution and 'clearing above "
    "the pivots' are the same thing."]),
48: ("Reduced row echelon form: general recipe", [
    "Two steps after reaching row echelon form. First divide each row by its pivot so every pivot is 1. Then, for each pivot, subtract "
    "multiples of its row from the rows above to turn every entry above the pivot into 0.",
    "Pivot columns become columns of the identity matrix; columns without a pivot keep whatever entries remain, and each such column is a "
    "free variable. The number of pivots, and so the rank, never changes."]),
49: ("Reduced row echelon form: 3×3 example", [
    "Start from [1 2 3; 0 1 4; 0 0 1]. R<sub>1</sub> ← R<sub>1</sub> − 2R<sub>2</sub> gives [1 0 −5]. R<sub>1</sub> ← R<sub>1</sub> + 5R<sub>3</sub> "
    "gives [1 0 0]. R<sub>2</sub> ← R<sub>2</sub> − 4R<sub>3</sub> gives [0 1 0]. The result is the identity, so the matrix has rank 3.",
    "Either order works (last pivot upwards or first pivot downwards) because clearing a column never disturbs a column already finished."]),
50: ("Part 6 divider", [
    "Gaussian elimination is the complete algorithm: everything from Parts 1 to 5, applied to the augmented matrix so the constants travel "
    "with the coefficients."]),
51: ("The augmented matrix", [
    "The <b>augmented matrix</b> [A | b] appends the constants column to the coefficient matrix. For 2a − b + c = 1, 2a + 2b + 4c = −2, "
    "4a + b = 4 it is [2 −1 1 | 1; 2 2 4 | −2; 4 1 0 | 4]. Note the 0 for the missing c in the third equation.",
    "Row operations act on whole rows including the constant, which is exactly what we did to equations in Part 1. The vertical bar is just "
    "a reminder of where the equals signs were."]),
52: ("Pivoting on the first column", [
    "R<sub>1</sub> ← ½R<sub>1</sub> makes the first pivot 1: [1 −½ ½ | ½]. Then R<sub>2</sub> ← R<sub>2</sub> − 2R<sub>1</sub>: "
    "[2 2 4 | −2] − [2 −1 1 | 1] = [0 3 3 | −3]. And R<sub>3</sub> ← R<sub>3</sub> − 4R<sub>1</sub>: [4 1 0 | 4] − [4 −2 2 | 2] = [0 3 −2 | 2].",
    "Say out loud what 'pivoting on a column' means: choose the pivot, scale it to 1, and use it to zero out everything below it."]),
53: ("Pivoting on the second and third columns", [
    "R<sub>2</sub> ← ⅓R<sub>2</sub> gives [0 1 1 | −1]. R<sub>3</sub> ← R<sub>3</sub> − 3R<sub>2</sub>: [0 3 −2 | 2] − [0 3 3 | −3] = "
    "[0 0 −5 | 5]. R<sub>3</sub> ← −⅕R<sub>3</sub> gives [0 0 1 | −1]. The augmented matrix is now in row echelon form: "
    "[1 −½ ½ | ½; 0 1 1 | −1; 0 0 1 | −1].",
    "As equations: a − ½b + ½c = ½, b + c = −1, c = −1. Three pivots, rank 3, so the system is non-singular and has a unique solution."]),
54: ("Back substitution", [
    "Clear above the third pivot: R<sub>2</sub> ← R<sub>2</sub> − R<sub>3</sub> gives [0 1 0 | 0]; R<sub>1</sub> ← R<sub>1</sub> − ½R<sub>3</sub> "
    "gives [1 −½ 0 | 1]. Clear above the second pivot: R<sub>1</sub> ← R<sub>1</sub> + ½R<sub>2</sub> gives [1 0 0 | 1]. The left block is "
    "the identity and the right column is the solution: a = 1, b = 0, c = −1.",
    "Check in the original equations: 2 − 0 − 1 = 1; 2 + 0 − 4 = −2; 4 + 0 = 4. (Historical note for the instructor: the source material "
    "listed the third equation as 4a + b = −1, which does not fit this walkthrough; 4a + b = 4 is the consistent version.)"]),
55: ("Gaussian elimination on a singular system", [
    "Augmented matrix [1 2 −1 | 5; 2 4 5 | 1; 3 6 4 | 6]. R<sub>2</sub> ← R<sub>2</sub> − 2R<sub>1</sub> gives [0 0 7 | −9]; "
    "R<sub>3</sub> ← R<sub>3</sub> − 3R<sub>1</sub> gives [0 0 7 | −9]; R<sub>3</sub> ← R<sub>3</sub> − R<sub>2</sub> gives [0 0 0 | 0]. "
    "The zero row with constant 0 reads 0 = 0: infinitely many solutions (b is free).",
    "Change the last constant from 6 to 10 and the same steps give [0 0 0 | 4], that is 0 = 4: no solution. The coefficient matrix is "
    "singular either way; the constants column decides between 'many' and 'none'. This is the matrix form of the redundant and "
    "contradictory cases from Part 1."]),
56: ("Gaussian elimination: the algorithm", [
    "Summarise the four steps: (1) form the augmented matrix; (2) pivot column by column to reach row echelon form; (3) clear above the "
    "pivots to reach reduced row echelon form, or back-substitute; (4) if a zero row appears, its constant tells you whether there are "
    "infinitely many solutions (constant 0) or none (constant non-zero); otherwise read the unique solution from the last column.",
    "Data science connection: this is what numpy.linalg.solve, scipy.linalg.lu and R's solve() do, with rows swapped so the largest "
    "available pivot is used (partial pivoting) for numerical stability. Fitting a linear model calls such a routine on the normal "
    "equations or an equivalent factorisation."]),
57: ("Recap", [
    "Read the eight definitions aloud and ask students to supply an example of each from today's slides. Suggested exit question: give a "
    "3×3 matrix of rank 2 and explain how you know its rank without computing a determinant.",
    "Next week builds on this: vectors, linear combinations and the geometry behind rank and the solution space."]),
}
