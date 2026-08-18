# Interactive Visual Explanation Update

This update adds step-by-step interactive visual explanations for geometry, coordinate graphs, trigonometry, mensuration and 3D solids.

## Student experience

- After a successful Gemini reasoning analysis, visual topics automatically receive a second, constrained visual-planning pass.
- 2D geometry and coordinate/graph questions are reconstructed as interactive JSXGraph scenes.
- 3D solid questions are reconstructed as interactive Three.js scenes that can be rotated and zoomed on iPad.
- Each explanation has Previous/Next controls and highlights only the sides, points, angles, diagonals or regions needed at that step.
- For 3D questions, the explanation is prompted to reveal the relevant 2D cross-section/triangle before using Pythagoras or trigonometry.
- Mathematical calculations continue to render through the existing MathIO/equation view.

## Reliability safeguards

- Gemini returns declarative geometry primitives only. It cannot inject JavaScript or executable graph expressions.
- Visual reconstruction is suppressed when confidence is low or the uploaded diagram is too ambiguous/cropped.
- Schematic diagrams explicitly state that they are not to scale unless the question provides a scale.
- The normal verified reasoning analysis remains available if the optional visual-generation call fails or reaches quota.

## Browser libraries

- JSXGraph 2.4.0 for 2D geometry/graphs.
- Three.js r185 with OrbitControls for 3D rotation/zoom.
- Both load client-side from pinned CDN URLs; no additional Python package is required.
