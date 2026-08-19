# Graph rendering fail-safe

The Word paper no longer depends on the asynchronous GeoGebra PNG capture.

For every explicit `y=f(x)` question:

1. GeoGebra is still used as the preferred interactive graph preview.
2. If its PNG has already been captured, that PNG is inserted in Word.
3. If the capture is still `0/1 ready`, Math Advisor immediately constructs a new
   graph scene from the exact equation and renders the curve locally.
4. This local graph does not depend on Gemini returning `diagram_scene_2d`.
5. Therefore a generated paper should no longer contain blank axes simply because
   the GeoGebra browser callback was not ready.

Regenerate the paper after deployment; previously downloaded Word files do not change.
