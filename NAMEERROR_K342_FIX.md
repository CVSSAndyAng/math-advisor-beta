# NameError fix + K342 removal

Fixed:
- Student-working GeoGebra no longer references the undefined local variable `question_text`.
- It now resolves the active question through a safe Streamlit session-state helper.
- K342 has been removed from the app track selector and offline-practice mappings.
- K232 remains available.

Validation:
- app.py: syntax checked
- offline_engine.py: syntax checked
- gemini_service.py: syntax checked
- learning_outcomes_data.py: syntax checked
