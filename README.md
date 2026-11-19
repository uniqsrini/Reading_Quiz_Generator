# Generative AI and RAG powered Reading Quiz Generator

Exam Buddy is a desktop application that generates multiple-choice quizzes from PDF teaching materials using Google Gemini. It provides a Tkinter GUI, an interactive quiz mode, and Word document exports.

## Features

- Native desktop GUI (Tkinter)
- Upload/select a PDF document
- Configure question count (1–50), difficulty, and topic focus
- Retrieval-augmented generation (RAG) using ChromaDB
- Gemini-powered structured MCQ generation
- Interactive quiz with score and explanations
- Export **all questions, options, correct answers, and explanations** to `.docx`
- Export a student version containing questions and options only

## Requirements

- Python 3.10 or newer
- A Google Gemini API key
- Windows, macOS, or Linux desktop with Tkinter installed

> The application is a desktop GUI. If running Jupyter remotely, the window will appear on the remote machine rather than on your local desktop.

## Installation

### 1. Clone the repository

```bash
git clone https://github.com/YOUR-USERNAME/Reading_Quiz_Generator.git
cd Reading_Quiz_Generator
```

### 2. Create and activate a virtual environment

Windows:

```powershell
py -m venv .venv
.venv\Scripts\activate
```

macOS/Linux:

```bash
python3 -m venv .venv
source .venv/bin/activate
```

### 3. Install dependencies

```bash
python -m pip install --upgrade pip
pip install -r requirements.txt
```

### 4. Run the application

```bash
python app.py
```

Enter your Google Gemini API key in the application window. The key is used locally and is not saved to a configuration file by this application.

## Usage

1. Enter your Gemini API key.
2. Select a text-based PDF.
3. Choose the number of questions.
4. Select a difficulty level.
5. Optionally enter a topic or chapter to focus on.
6. Click **GENERATE QUIZ**.
7. After generation, choose:
   - **Download All Questions & Answers (.docx)** — includes the answer key and explanations.
   - **Download Questions Only (.docx)** — suitable for distributing to students.
   - **Take Interactive Quiz** — answer questions and view your score.

## Notes and limitations

- Scanned/image-only PDFs may require OCR before use.
- The generated questions should be reviewed by an instructor before formal assessment.
- Model availability and structured-output support depend on the Google AI account and installed `langchain-google-genai` version.
- The model identifier is configured in `app.py` as `gemini-3.8-flash`. Change `MODEL_NAME` if your Google AI account exposes a different supported model identifier.
- Embedding requests and generation requests may incur Google API usage charges.
- Large PDFs may require more memory, processing time, and API quota.

## Troubleshooting

### Tkinter is unavailable

On some Linux distributions, install the system Tk package (for example, `python3-tk`) and try again.

### Model not found or unsupported

Check that the configured model identifier is available to your API key. Update `MODEL_NAME` in `app.py` if necessary.

### PDF has no readable text

The current PDF loader extracts embedded text. Run OCR on scanned PDFs before importing them.

### Dependency/import errors

Use a fresh virtual environment and reinstall:

```bash
pip install -r requirements.txt
```

## Repository structure

```text
exam-buddy/
├── app.py
├── requirements.txt
├── README.md
├── .gitignore
└── LICENSE
```

## License

This project is distributed under the MIT License. See `LICENSE`.
