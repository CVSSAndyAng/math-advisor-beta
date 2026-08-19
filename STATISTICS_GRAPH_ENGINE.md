# Statistics Graph Engine + strict graph audit

Supported statistics diagrams:
- cumulative-frequency curves
- histograms
- frequency polygons
- box-and-whisker plots
- scatter plots
- line graphs
- bar charts

The engine creates deterministic 300-DPI PNG images for the Streamlit preview and
downloaded Word papers.

For cumulative-frequency curves it validates:
- one frequency per class interval;
- cumulative values aligned with class boundaries;
- non-decreasing cumulative totals;
- final cumulative frequency equals the total sample size.

Question/solution behaviour:
- If students are asked to draw/construct/complete a graph, the question can show
  blank axes or a blank grid without the completed curve.
- If the question says a graph is already shown or students must read values from it,
  the completed graph is required.

Function graph fail-safe:
- General forms such as y = a sin(bx+c) are not plot-ready.
- Graph-reading questions require a fully numeric hidden graph_equations function.
- Blank axes are rejected for graph-reading questions.
- Blank axes remain valid only when the student is explicitly asked to draw/sketch/plot.

Regenerate papers after deployment; existing generated drafts do not contain the new
statistics graph data or stricter hidden graph specification.
