# Smart Teacher PPT Maker

A small Streamlit web app for generating editable PowerPoint lessons for Mathematics (Classes 6–10). It works in a mobile or desktop browser once deployed.

## What it does
- Choose Class 6–10, curriculum, language, slide count and design style.
- Enter a chapter/topic and optional teaching instructions.
- Optionally paste textbook/reference material for closer alignment.
- Uses the OpenAI Responses API to draft slide content.
- Creates a downloadable `.pptx` file.

## Important
- This starter app is not hosted yet. It must be deployed once before it can be opened like a website on your phone and computer.
- You need an OpenAI API key and API billing/credits. A ChatGPT subscription does not automatically include API usage.
- The app does not store the API key; it is entered for the current session. Do not share your key.
- AI-generated mathematics, proofs and examples must be checked by the teacher. It does not guarantee verbatim NCERT questions unless source text is supplied and verified.

## Deploy it free (hosting may have plan limits)
1. Create/sign in to GitHub.
2. Upload `app.py` and `requirements.txt` to a new repository.
3. Open Streamlit Community Cloud and connect that repository.
4. Choose `app.py` as the app entry point and deploy.
5. Open the deployed app URL on your Android phone, tablet or computer. You can add it to your home screen.

## Run on your own computer
1. Install Python 3.10 or newer.
2. Open terminal in this folder.
3. Run:
   `pip install -r requirements.txt`
4. Run:
   `streamlit run app.py`
5. Open the local URL shown in the terminal.

## Files
- `app.py` — application
- `requirements.txt` — Python packages
