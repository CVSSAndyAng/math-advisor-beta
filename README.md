# SG Math Reasoning Tutor

A reasoning-first AI mathematics tutor for Singapore secondary
mathematics, built with Streamlit and the Gemini API.

## Supported curriculum

The tutor supports the 2027 Singapore-Cambridge Secondary Education
Certificate (SEC) mathematics tracks:

-   G1 Mathematics (K110)
-   G2 Mathematics (K210)
-   G3 Mathematics (K310)
-   G2 Additional Mathematics (K232)
-   G3 Additional Mathematics (K341)

The existing 2026 O-Level and N-Level Mathematics options are retained
as transition tracks.

## What the tutor does

The tutor is designed to analyse how a student works through a
mathematics problem rather than checking only the final answer.

It can:

-   accept typed questions, images and PDFs;
-   detect multiple questions in an uploaded worksheet or examination
    page;
-   check whether a question is complete, mathematically feasible and
    internally consistent before analysing a solution;
-   read student working from a separate upload;
-   read student working that is already written on the same uploaded
    question image or PDF;
-   provide guided solving when no student solution is supplied;
-   analyse handwritten work;
-   identify the first material logic break;
-   distinguish conceptual, procedural, arithmetic and presentation
    errors;
-   provide progressive hints and advice;
-   generate adaptive Near Transfer, Varied Context and Stretch
    practice;
-   keep a student in the same practice category until mastery is
    demonstrated;
-   check every required subpart before allowing progression;
-   independently verify calculable mathematics using Gemini code
    execution where appropriate;
-   support class/batch analysis of multiple student submissions and
    identify common trends.

## Question upload modes

After supplying a question, choose one of three modes:

### Separate student solution

Use this when the question and the student's solution are in different
files or when the student will type their working.

### Student solution is already on the question upload

Use this when the uploaded photo or PDF contains both the printed
question and the student's handwritten or annotated solution.

### No student solution --- guide me to solve it

Use this when the student has not attempted the question.

The tutor first checks question feasibility, then guides the student
through:

1.  the goal and known information;
2.  a diagnostic starting question;
3.  progressive hints;
4.  solution steps revealed one at a time;
5.  a final verified answer that remains hidden until the end.

## Adaptive practice

Practice progresses through:

**Near Transfer → Varied Context → Stretch**

The next category remains locked until the current category is secure.

If a student makes a non-secure attempt, the tutor remains in the same
category and generates another question focused on the same reasoning
gap.

For multi-part questions, all required parts must be completed correctly
before the attempt can count as secure.

## Mathematics input

Students can use:

-   the visual equation editor;
-   handwritten working on an iPad using Apple Pencil, stylus or finger;
-   typed text working;
-   uploaded handwritten images or PDFs;
-   editable tables;
-   interactive graph and coordinate tools.

The handwriting pad uses an explicit **Save handwriting** action so
drawing does not repeatedly refresh the Streamlit app.

## Interactive mathematics tools

For graph, function and coordinate-geometry questions, the tutor can
provide an interactive workspace with tools such as:

-   point;
-   line;
-   segment;
-   ray;
-   vector;
-   midpoint;
-   parallel and perpendicular lines;
-   circle;
-   polygon;
-   angle measurement;
-   distance measurement;
-   pan and zoom;
-   undo, delete and clear.

Students can also insert and fill editable tables and optionally display
a function graph.

## Visual explanations

Visual step-by-step simulation is generated only when graphics genuinely
help, such as:

-   geometry;
-   coordinate geometry;
-   graphs and functions;
-   trigonometry involving diagrams;
-   bearings;
-   transformations and constructions;
-   3D solids;
-   isometric and orthographic-view questions.

Ordinary algebra, indices, standard form and other non-graphical
questions do not need a visual simulation.

For 3D questions, the tutor can reconstruct and explore solid forms
where the uploaded information is sufficiently reliable.

## Geometry safeguards

For shaded-region, composite-area, perimeter and arc-length questions,
the tutor must identify the relevant boundaries before formulating the
mathematical expression.

The intended reasoning sequence is:

1.  identify the required region;
2.  list the outer boundary lines/arcs/curves;
3.  identify excluded or internal boundaries;
4.  check that the boundary description is consistent;
5.  formulate the area/perimeter expression;
6.  independently verify the calculation.

The tutor should not guess a familiar formula merely because a diagram
resembles a common textbook example.

## Question feasibility

Before student-work analysis, the tutor checks for:

-   missing information;
-   contradictory givens;
-   ambiguous wording;
-   impossible values;
-   cropped or missing diagrams;
-   unclear tables or graphs;
-   suspected typographical errors;
-   inconsistent geometric relationships;
-   whether every subpart is answerable.

If a blocking issue is detected, marking is stopped until the question
is clarified.

For uploaded diagrams, the app can highlight relevant regions associated
with feasibility warnings.

## Accuracy architecture

The tutor uses multiple layers rather than relying on a single model
response:

1.  question feasibility;
2.  multimodal question/diagram interpretation;
3.  independent mathematical verification;
4.  student-working analysis;
5.  adaptive follow-up practice.

Gemini code execution is used as a verification aid for suitable
computational work such as arithmetic, algebra, trigonometry, coordinate
calculations, matrices and statistics.

Code execution does not replace careful visual interpretation of
diagrams.

## Class trends / batch analysis

The Class Trends workflow allows several student solutions to be
uploaded for the same question.

The question is verified once, then submissions are analysed separately.

The summary can identify patterns such as:

-   common misconceptions;
-   common first logic-break steps;
-   recurring presentation errors;
-   areas of secure understanding;
-   concepts that may need whole-class reteaching.

Avoid including unnecessary student-identifying information in filenames
or uploaded work.

## Fast and Full analysis

Where available:

-   **Fast** prioritises reasoning analysis and mathematical
    verification and delays optional visual-generation work.
-   **Full** includes the richer visual explanation workflow.

This helps reduce waiting time when visual reconstruction is not
immediately needed.

## Running on Streamlit Community Cloud

The repository should contain at least:

``` text
app.py
gemini_service.py
offline_engine.py
requirements.txt
README.md
```

Deploy `app.py` as the Streamlit main file.

### Gemini secret

In Streamlit Community Cloud, add the Gemini API key under **App
settings → Secrets**:

``` toml
GEMINI_API_KEY = "your-api-key"
```

Do not commit the API key to GitHub.

## Dependencies

Dependencies are installed automatically by Streamlit from
`requirements.txt`.

The project uses the current `google-genai` Python SDK rather than the
legacy `google.generativeai` package.

## Privacy

Student work may contain personal information.

Before uploading real student work:

-   remove names where they are unnecessary;
-   remove NRICs and other identifiers;
-   avoid unnecessary school/class identifiers;
-   follow applicable school, parent/guardian and organisational
    policies;
-   review the data-use terms for the Gemini API tier being used.

## Important limitation

The tutor is an educational support tool, not an official marking
authority.

AI interpretation of handwriting, diagrams and unusual solution methods
can be imperfect. For high-stakes assessment decisions, verify feedback
against the official syllabus, examination marking guidance or a
qualified teacher.
