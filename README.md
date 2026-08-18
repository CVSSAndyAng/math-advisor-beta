# Math Advisor Beta

**Math Advisor** is a beta mathematics teaching and assessment-support application for Singapore secondary Mathematics. It supports the configured 2027 SEC G1/G2/G3 Mathematics and Additional Mathematics workflows, with transition support for O-/N-Level use where available in the app.

> **Beta educational tool:** Math Advisor is not an official SEAB or MOE product. Generated solutions, diagrams, assessment papers and marking schemes should be reviewed by a teacher before formal use.

## Main features

### Analyse and advise
- Upload or enter a mathematics question.
- Upload student working separately, or analyse working already visible in the question image.
- Advise how to solve a question when no student solution is supplied.
- Question-feasibility checking can be used or bypassed.
- Guided help supports hints or full solutions.
- Mathematical expressions use MathIO-style equation presentation in the web app.
- Geometry and graph questions can use diagrams, axes and gridlines where appropriate.

### Student working
- Equation-editor and text-working modes.
- Mathematics keyboard designed for desktop, tablet and phone use.
- Multiple working steps.
- Mathematical working is kept separate from ordinary explanatory text.

### Paper setter
The teacher workflow combines paper setting and full-paper solution generation.

Teachers can:
1. **Create a new assessment paper**, or
2. **Upload an existing question paper** to generate worked solutions and a teacher marking scheme.

For new papers, teachers can choose the syllabus/track, assessment type, total marks, number of main questions, duration, topic scope and other available settings.

A **reference format paper is optional**:
- If supplied, it guides structure, numbering, mark placement, command-word style and difficulty progression.
- If omitted, Math Advisor uses the selected syllabus/settings and built-in Singapore secondary Mathematics assessment conventions.

Mark allocation is flexible. The application can reconcile small inconsistencies between part marks, question marks and the requested paper total instead of rejecting an otherwise usable paper.

### Generated papers
Current paper-generation rules include:
- Times New Roman, 11 pt for normal paper text.
- New main questions start on new pages.
- Diagrams are numbered sequentially as Figure 1, Figure 2, etc.
- Diagram labels/captions use Times New Roman styling.
- Mathematical notation should be symbolic rather than written out in prose.
- Greek symbols such as `theta` are converted to symbols such as `θ`.
- Expressions such as `y = x squared divided by (2x+1)` are normalised toward symbolic equation form.
- Word output uses editable Microsoft Word Equation Editor/OMML structures where supported.
- Supported fractions, roots, powers, vectors and integrals are converted into equation objects.
- Generated diagrams are checked against question wording where possible.
- Unverified generated diagrams may be withheld and flagged for teacher review rather than causing the whole paper to fail.

### Graphs and geometry
- 2D geometry preserves mathematical aspect ratios where required.
- Function questions use a deterministic local graph renderer from explicit equations rather than treating blank axes as a complete graph.
- Multiple explicit functions can be plotted on the same axes.
- 3D rendering supports configured primitives such as cones, cylinders and spheres.
- Graph and geometry output should always be checked during the beta period.

### Offline practice
- No-credit syllabus-generated practice remains available without a Gemini call.
- Offline questions support MathIO-style mathematical display.
- Expressions are kept together horizontally rather than split into separate vertical fragments.
- Hints and student-working input are available.

## Repository files

The deployed repository must contain **all local Python modules imported by `app.py`**.

Typical structure:

```text
math-advisor-beta/
├── app.py
├── gemini_service.py
├── offline_engine.py
├── requirements.txt
├── README.md
└── other local modules imported by app.py
```

If a local module is missing, Streamlit may fail with `ModuleNotFoundError`.

## Streamlit Community Cloud deployment

1. Push the complete project to GitHub.
2. Connect Streamlit Community Cloud to GitHub.
3. For a private repository, grant Streamlit access to that repository.
4. Deploy the `main` branch with `app.py` as the main file.
5. Use **Python 3.12** for the current beta unless the complete dependency set has been separately tested on a newer version.
6. In **App settings → Secrets**, add:

```toml
GEMINI_API_KEY = "your-key-here"
```

7. Never commit the Gemini API key to GitHub.
8. Deploy/reboot and inspect the Streamlit logs for missing modules or dependency errors.

## Updating the beta

1. Commit replacement files to the GitHub `main` branch.
2. Allow Streamlit to redeploy automatically, or reboot from **Manage app**.
3. Regenerate papers after changes to equation rendering, diagrams, formatting or generation logic. Existing Word files are not retroactively changed.
4. Run the regression checks below.

## Recommended beta regression checks

Test:
- Standard form and indices, e.g. `7.2 × 10^-4`.
- Fractions, roots and powers.
- Matrices.
- Vectors and over-arrow notation.
- Definite integrals.
- `θ`, angle notation and degree symbols.
- Trigonometric functions and graphs.
- Multiple curves/lines on one graph.
- Circle geometry, tangents and intersecting chords.
- Sectors and shaded regions.
- Coordinate geometry.
- 2D and 3D diagrams.
- Maths-keyboard typing and save persistence on desktop and iPad/mobile.
- Paper generation with and without a reference-format paper.
- Uploaded full-paper solutions and marking scheme.
- Flexible total-mark reconciliation.
- Word equation rendering.
- Sequential figure numbering and question page breaks.

## Important beta limitations

- AI-generated mathematical work still requires teacher verification.
- Diagram generation is a higher-risk area and should be checked against the question statement.
- Generated marking schemes are teacher drafts, not official SEAB/MOE marking schemes.
- Gemini API quotas/rate limits can temporarily prevent online generation.
- Offline practice is no-credit practice and does not replace formal assessment.
- Word equation compatibility can vary between Microsoft Word and other office suites.

## Privacy and API keys

Do not place API keys, passwords or other secrets in the repository. Use Streamlit Secrets for deployment credentials.

Review your institution's requirements before uploading identifiable student work or assessment materials to a cloud-hosted service.

## Beta status

This repository contains an actively tested **Beta** version of Math Advisor. Features and output formats may change as mathematical accuracy, diagram reliability, MathIO rendering, paper generation and mobile input are improved.
