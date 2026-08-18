# Math Advisor Beta

## Release
Beta build prepared on 2026-08-18.

## Included
- Math Advisor tutoring workflow
- Paper Setter + uploaded-paper solutions/marking-scheme workflow
- Optional reference format paper
- MathIO rendering improvements
- Offline practice
- Diagram validation and tolerant diagram handling
- Flexible mark allocation
- Times New Roman 11 pt generated-paper styling
- Sequential Figure numbering

## Streamlit Cloud deployment
1. Put these files in the root of your GitHub repository.
2. Ensure `requirements.txt` is committed.
3. In Streamlit Community Cloud, deploy `app.py`.
4. Add `GEMINI_API_KEY` under **App settings → Secrets**.
5. Redeploy/reboot the app.
6. Verify the page title shows **Math Advisor Beta**.

## Beta warning
Generated solutions, diagrams and marking schemes should be reviewed by a teacher before formal assessment use.
