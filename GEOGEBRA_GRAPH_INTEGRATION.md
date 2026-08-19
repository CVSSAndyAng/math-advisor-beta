# GeoGebra Graph Integration

Math Advisor now uses GeoGebra as the primary browser-side renderer for explicit `y=f(x)` questions in Paper Setter.

## Behaviour
1. Generated function equations are sanitized and sent to the official GeoGebra Graphing app.
2. GeoGebra draws the curve with axes and gridlines.
3. The Graphics View is exported as a 300-DPI PNG.
4. The PNG is returned to Streamlit and cached for the current generated paper.
5. The downloaded Word question paper uses the captured GeoGebra PNG.
6. If GeoGebra is unavailable, Math Advisor uses the existing local deterministic graph renderer.

## Streamlit requirement
This implementation uses Streamlit Components v2. If your `requirements.txt` pins an older Streamlit release that does not expose `st.components.v2`, upgrade Streamlit to a current release.

## Network
The user's browser must be able to load:
`https://www.geogebra.org/apps/deployggb.js`

Only the sanitized mathematical function expression and graph viewport are sent to the GeoGebra browser applet; student uploads are not sent through this graph component.
