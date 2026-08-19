# Text/Math separation, Question Keyboard and GeoGebra update

## Goal/guidance rendering
The mixed renderer now stops an equation before English connectors such as:
`for`, `where`, `when`, `with`, `over`, `from`, `on`, `using`, and `given that`.

Example:
- Text: `Draw the graph of the trigonometric function`
- MathIO: `y = sin(3x)`
- Text: `for a standard domain.`

## Word equations
`\left` and `\right` are removed safely before OMML parsing, preventing corrupted
text such as `ight)`. The words `pi` and `theta` are normalised to mathematical
symbols.

## Question input
Analyse and typed-algebra workflows now offer:
- Text
- Text + math keyboard

The second mode uses the existing MathLive equation editor for structured mathematics.

## GeoGebra
Graph/geometry questions now have an expandable `Open GeoGebra to illustrate the question`
section with links to:
- GeoGebra Graphing
- GeoGebra Geometry

GeoGebra opens in a new browser tab. Uploaded images/student work are not automatically
sent to GeoGebra.

The existing deterministic renderers remain responsible for generated paper PNG output.
