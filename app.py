"""
Exam Buddy - AI-powered PDF quiz generator
Desktop application using Tkinter, Google Gemini, LangChain, and Chroma.
"""

import os
import threading
import queue
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter.scrolledtext import ScrolledText
from datetime import datetime
import uuid

from pydantic import BaseModel, Field
from typing import List

from docx import Document
from docx.shared import Pt, Inches

from langchain_community.document_loaders import PyPDFLoader
from langchain_text_splitters import RecursiveCharacterTextSplitter
from langchain_google_genai import (
    ChatGoogleGenerativeAI,
    GoogleGenerativeAIEmbeddings,
)
from langchain_chroma import Chroma


APP_NAME = "Exam Buddy"
MODEL_NAME = "gemini-3.8-flash"
EMBEDDING_MODEL = "models/gemini-embedding-001"

quiz_data = None
selected_pdf = None
result_queue = queue.Queue()


class QuestionSchema(BaseModel):
    question: str = Field(description="The complete examination question")
    options: List[str] = Field(description="Exactly four answer options")
    correct_answer: str = Field(description="Correct option letter: A, B, C, or D")
    explanation: str = Field(description="Explanation of why the answer is correct")


class QuizSchema(BaseModel):
    questions: List[QuestionSchema]


def generate_quiz(api_key, pdf_path, number_of_questions, difficulty, topic):
    """Read a PDF, retrieve relevant passages, and generate validated MCQs."""
    os.environ["GOOGLE_API_KEY"] = api_key

    documents = PyPDFLoader(pdf_path).load()
    if not documents:
        raise ValueError("No readable content was found in the PDF.")

    full_text = "\n".join(doc.page_content for doc in documents)
    if len(full_text.strip()) < 100:
        raise ValueError(
            "The PDF contains insufficient readable text. Scanned PDFs may need OCR."
        )

    splitter = RecursiveCharacterTextSplitter(
        chunk_size=1200, chunk_overlap=200
    )
    chunks = splitter.split_documents(documents)
    if not chunks:
        raise ValueError("Could not divide the PDF into text sections.")

    embeddings = GoogleGenerativeAIEmbeddings(
        model=EMBEDDING_MODEL,
        google_api_key=api_key,
    )

    collection_name = "quiz_" + uuid.uuid4().hex[:12]
    vectorstore = Chroma.from_documents(
        documents=chunks,
        embedding=embeddings,
        collection_name=collection_name,
    )

    try:
        search_query = (
            f"Find detailed educational information, definitions, principles, "
            f"equations and examples about: {topic}"
            if topic.strip()
            else "Find the most important concepts, definitions, principles, "
                 "equations and examples in this document."
        )

        retrieved_docs = vectorstore.similarity_search(
            search_query, k=min(50, len(chunks))
        )
        context = "\n\n".join(doc.page_content for doc in retrieved_docs)
        if len(context.strip()) < 500:
            context = full_text[:50000]

        llm = ChatGoogleGenerativeAI(
            model=MODEL_NAME,
            google_api_key=api_key,
            temperature=0.3,
        )
        structured_llm = llm.with_structured_output(QuizSchema)

        topic_instruction = (
            f"Concentrate on this topic: {topic}."
            if topic.strip()
            else "Cover important topics throughout the PDF."
        )

        prompt = f"""
You are an experienced university-level examination question setter.

Generate exactly {number_of_questions} multiple-choice questions using
ONLY the supplied PDF content.

DIFFICULTY: {difficulty}
TOPIC: {topic_instruction}

STRICT REQUIREMENTS:
1. Generate exactly {number_of_questions} questions.
2. Each question must have exactly four distinct options.
3. There must be exactly one correct answer.
4. correct_answer must be only A, B, C, or D.
5. Provide a concise, accurate explanation.
6. Avoid duplicate or substantially similar questions.
7. Do not introduce facts absent from the PDF.
8. Use numerical, conceptual, and application questions where appropriate.
9. Ensure the correct answer matches the corresponding option.

PDF CONTENT:
{context}

Return the result in the required structured format.
"""
        quiz = structured_llm.invoke(prompt)
        if not quiz or not quiz.questions:
            raise ValueError("Gemini did not return any questions.")

        if len(quiz.questions) != number_of_questions:
            raise ValueError(
                f"Requested {number_of_questions} questions, but Gemini returned "
                f"{len(quiz.questions)}. Please try again."
            )

        for index, question in enumerate(quiz.questions, start=1):
            if len(question.options) != 4:
                raise ValueError(
                    f"Question {index} does not have exactly four options."
                )
            if len(set(question.options)) != 4:
                raise ValueError(f"Question {index} contains duplicate options.")

            answer = question.correct_answer.strip().upper()
            if answer not in ("A", "B", "C", "D"):
                raise ValueError(
                    f"Question {index} has invalid answer: "
                    f"{question.correct_answer}"
                )
            question.correct_answer = answer
            if not question.question.strip():
                raise ValueError(f"Question {index} is empty.")

        return quiz
    finally:
        try:
            vectorstore.delete_collection()
        except Exception:
            pass


def download_all_questions_answers(quiz):
    """Export all questions, options, answers, and explanations to DOCX."""
    if quiz is None or not quiz.questions:
        messagebox.showwarning("No Questions", "Please generate a quiz first.")
        return

    filepath = filedialog.asksaveasfilename(
        title="Save all questions and answers",
        defaultextension=".docx",
        filetypes=[("Word Document", "*.docx")],
        initialfile="Complete_Quiz_Questions_and_Answers.docx",
    )
    if not filepath:
        return

    try:
        doc = Document()
        section = doc.sections[0]
        section.top_margin = Inches(0.7)
        section.bottom_margin = Inches(0.7)
        section.left_margin = Inches(0.8)
        section.right_margin = Inches(0.8)

        doc.styles["Normal"].font.name = "Arial"
        doc.styles["Normal"].font.size = Pt(10)

        doc.add_heading("Complete Quiz – Questions and Answers", level=0)
        doc.add_paragraph(f"Total Questions: {len(quiz.questions)}")
        doc.add_paragraph(
            "Includes all questions, four options, correct answers, and explanations."
        )
        doc.add_paragraph(
            "Generated: " + datetime.now().strftime("%d %B %Y, %H:%M")
        )
        doc.add_page_break()

        for i, question in enumerate(quiz.questions, start=1):
            doc.add_heading(f"Question {i}", level=2)
            p = doc.add_paragraph()
            p.paragraph_format.space_after = Pt(6)
            p.add_run(question.question).bold = True

            doc.add_paragraph("Options:")
            for j, option in enumerate(question.options):
                letter = chr(65 + j)
                p = doc.add_paragraph()
                p.paragraph_format.left_indent = Inches(0.25)
                p.add_run(f"{letter}. ").bold = True
                p.add_run(option)

            answer_letter = question.correct_answer.upper()
            answer_index = ord(answer_letter) - ord("A")
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(5)
            p.add_run("Correct Answer: ").bold = True
            p.add_run(f"{answer_letter}. {question.options[answer_index]}")

            if question.explanation:
                p = doc.add_paragraph()
                p.add_run("Explanation: ").bold = True
                p.add_run(question.explanation)

            doc.add_paragraph("─" * 65)

        doc.save(filepath)
        messagebox.showinfo(
            "Download Complete", f"All questions and answers saved to:\n{filepath}"
        )
    except Exception as exc:
        messagebox.showerror("DOCX Export Error", str(exc))


def download_questions_only(quiz):
    """Export questions and options without the answer key."""
    if quiz is None or not quiz.questions:
        messagebox.showwarning("No Questions", "Please generate a quiz first.")
        return

    filepath = filedialog.asksaveasfilename(
        title="Save questions",
        defaultextension=".docx",
        filetypes=[("Word Document", "*.docx")],
        initialfile="Quiz_Questions.docx",
    )
    if not filepath:
        return

    try:
        doc = Document()
        doc.add_heading("Examination Questions", level=0)
        for i, question in enumerate(quiz.questions, start=1):
            doc.add_heading(f"Question {i}", level=2)
            doc.add_paragraph(question.question)
            for j, option in enumerate(question.options):
                doc.add_paragraph(f"{chr(65 + j)}. {option}")
            doc.add_paragraph("")
        doc.save(filepath)
        messagebox.showinfo("Download Complete", f"Questions saved to:\n{filepath}")
    except Exception as exc:
        messagebox.showerror("DOCX Export Error", str(exc))


def open_exam_window(quiz):
    if not quiz or not quiz.questions:
        return

    window = tk.Toplevel(root)
    window.title("Exam Buddy - Take Quiz")
    window.geometry("850x700")
    window.minsize(650, 500)

    canvas = tk.Canvas(window)
    scrollbar = ttk.Scrollbar(window, orient="vertical", command=canvas.yview)
    frame = ttk.Frame(canvas, padding=15)
    frame.bind(
        "<Configure>",
        lambda event: canvas.configure(scrollregion=canvas.bbox("all")),
    )
    canvas_window = canvas.create_window((0, 0), window=frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)
    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")
    canvas.bind(
        "<Configure>",
        lambda event: canvas.itemconfig(canvas_window, width=event.width),
    )

    tk.Label(frame, text="Exam Buddy", font=("Arial", 20, "bold")).pack(pady=10)
    tk.Label(
        frame, text=f"Total Questions: {len(quiz.questions)}"
    ).pack(pady=5)

    user_answers = []
    for i, question in enumerate(quiz.questions):
        qframe = ttk.LabelFrame(frame, text=f"Question {i + 1}", padding=12)
        qframe.pack(fill="x", pady=10)
        tk.Label(
            qframe,
            text=question.question,
            font=("Arial", 11, "bold"),
            wraplength=730,
            justify="left",
        ).pack(anchor="w", pady=5)

        answer_var = tk.StringVar(value="")
        user_answers.append(answer_var)
        for j, option in enumerate(question.options):
            letter = chr(65 + j)
            ttk.Radiobutton(
                qframe,
                text=f"{letter}. {option}",
                variable=answer_var,
                value=letter,
            ).pack(anchor="w", pady=3)

    def submit_exam():
        unanswered = [
            i + 1 for i, answer in enumerate(user_answers) if not answer.get()
        ]
        if unanswered:
            messagebox.showwarning(
                "Incomplete Quiz",
                "Please answer all questions. Unanswered: "
                + ", ".join(map(str, unanswered)),
                parent=window,
            )
            return

        score = 0
        results = []
        for i, question in enumerate(quiz.questions):
            selected = user_answers[i].get()
            correct = question.correct_answer.upper()
            is_correct = selected == correct
            score += int(is_correct)
            results.append(
                {
                    "question": question.question,
                    "selected": selected,
                    "correct": correct,
                    "explanation": question.explanation,
                    "is_correct": is_correct,
                }
            )
        window.destroy()
        show_results(score, len(quiz.questions), results)

    ttk.Button(frame, text="Submit Quiz", command=submit_exam).pack(pady=20)


def show_results(score, total, results):
    window = tk.Toplevel(root)
    window.title("Quiz Results")
    window.geometry("850x650")

    tk.Label(
        window, text="Quiz Results", font=("Arial", 20, "bold")
    ).pack(pady=15)
    percentage = score / total * 100 if total else 0
    tk.Label(
        window,
        text=f"Your Score: {score}/{total} ({percentage:.1f}%)",
        font=("Arial", 16, "bold"),
    ).pack(pady=10)

    canvas = tk.Canvas(window)
    scrollbar = ttk.Scrollbar(window, orient="vertical", command=canvas.yview)
    frame = ttk.Frame(canvas, padding=15)
    frame.bind(
        "<Configure>",
        lambda event: canvas.configure(scrollregion=canvas.bbox("all")),
    )
    canvas.create_window((0, 0), window=frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)
    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    for i, item in enumerate(results, start=1):
        status = "CORRECT" if item["is_correct"] else "INCORRECT"
        color = "green" if item["is_correct"] else "red"
        tk.Label(
            frame, text=f"Question {i}: {status}",
            font=("Arial", 11, "bold"), fg=color,
        ).pack(anchor="w", pady=(10, 2))
        tk.Label(
            frame, text=item["question"], wraplength=750, justify="left"
        ).pack(anchor="w")
        tk.Label(frame, text=f"Your answer: {item['selected']}").pack(anchor="w")
        tk.Label(
            frame, text=f"Correct answer: {item['correct']}",
            font=("Arial", 10, "bold"),
        ).pack(anchor="w")
        if item["explanation"]:
            tk.Label(
                frame,
                text=f"Explanation: {item['explanation']}",
                wraplength=750,
                justify="left",
            ).pack(anchor="w", pady=(2, 8))


def generation_worker(api_key, pdf_path, count, difficulty, topic):
    try:
        result_queue.put(("status", "Reading PDF document..."))
        result_queue.put(("status", "Creating embeddings and searching the PDF..."))
        quiz = generate_quiz(api_key, pdf_path, count, difficulty, topic)
        result_queue.put(("success", quiz))
    except Exception as exc:
        result_queue.put(("error", str(exc)))


def check_queue():
    global quiz_data
    try:
        while True:
            message_type, message = result_queue.get_nowait()
            if message_type == "status":
                status_label.config(text=message)
            elif message_type == "success":
                quiz_data = message
                progress.stop()
                status_label.config(
                    text=f"Successfully generated {len(quiz_data.questions)} questions."
                )
                generate_button.config(state="normal")
                download_all_button.config(state="normal")
                download_questions_button.config(state="normal")
                take_quiz_button.config(state="normal")
                messagebox.showinfo(
                    "Quiz Generated",
                    f"{len(quiz_data.questions)} questions generated successfully.",
                )
            elif message_type == "error":
                progress.stop()
                status_label.config(text="Quiz generation failed.")
                generate_button.config(state="normal")
                messagebox.showerror("Generation Error", message)
    except queue.Empty:
        pass

    if generate_button["state"] == "disabled":
        root.after(200, check_queue)


def browse_pdf():
    global selected_pdf
    filepath = filedialog.askopenfilename(
        title="Select PDF", filetypes=[("PDF Documents", "*.pdf")]
    )
    if filepath:
        selected_pdf = filepath
        pdf_label.config(text=os.path.basename(filepath))


def toggle_api_visibility():
    api_entry.config(show="" if show_api_var.get() else "*")


def start_generation():
    api_key = api_entry.get().strip()
    if not api_key:
        messagebox.showwarning("Missing API Key", "Please enter your Gemini API key.")
        return
    if not selected_pdf or not os.path.exists(selected_pdf):
        messagebox.showwarning("Missing PDF", "Please select a valid PDF file.")
        return

    try:
        count = int(question_count.get())
        if not 1 <= count <= 50:
            raise ValueError
    except ValueError:
        messagebox.showwarning(
            "Invalid Question Count", "Enter a number between 1 and 50."
        )
        return

    difficulty = difficulty_combo.get()
    topic = topic_text.get("1.0", "end").strip()

    generate_button.config(state="disabled")
    download_all_button.config(state="disabled")
    download_questions_button.config(state="disabled")
    take_quiz_button.config(state="disabled")
    status_label.config(text="Starting quiz generation...")
    progress.start(10)

    thread = threading.Thread(
        target=generation_worker,
        args=(api_key, selected_pdf, count, difficulty, topic),
        daemon=True,
    )
    thread.start()
    root.after(200, check_queue)


def create_app():
    global root, api_entry, show_api_var, pdf_label
    global question_count, difficulty_combo, topic_text
    global generate_button, progress, status_label
    global download_all_button, download_questions_button, take_quiz_button

    root = tk.Tk()
    root.title(f"{APP_NAME} - AI Quiz Maker")
    root.geometry("750x780")
    root.minsize(650, 650)
    root.configure(bg="#f2f4f8")

    header = tk.Frame(root, bg="#243b55", padx=20, pady=20)
    header.pack(fill="x")
    tk.Label(
        header, text="EXAM BUDDY", font=("Arial", 24, "bold"),
        bg="#243b55", fg="white",
    ).pack()
    tk.Label(
        header, text="AI-Powered PDF Quiz Generator", font=("Arial", 11),
        bg="#243b55", fg="white",
    ).pack(pady=5)

    main = tk.Frame(root, bg="white", padx=25, pady=20)
    main.pack(fill="both", expand=True, padx=20, pady=15)

    tk.Label(
        main, text="Google Gemini API Key", font=("Arial", 11, "bold"),
        bg="white",
    ).pack(anchor="w", pady=(5, 5))
    api_entry = tk.Entry(main, show="*", font=("Arial", 11), width=65)
    api_entry.pack(fill="x", pady=5)

    show_api_var = tk.BooleanVar(value=False)
    tk.Checkbutton(
        main, text="Show API key", variable=show_api_var,
        command=toggle_api_visibility, bg="white",
    ).pack(anchor="w")

    tk.Label(
        main, text="Select PDF Document", font=("Arial", 11, "bold"),
        bg="white",
    ).pack(anchor="w", pady=(15, 5))
    pdf_frame = tk.Frame(main, bg="white")
    pdf_frame.pack(fill="x")
    tk.Button(
        pdf_frame, text="Browse PDF", command=browse_pdf,
        bg="#2874a6", fg="white", padx=15, pady=7,
    ).pack(side="left")
    pdf_label = tk.Label(
        pdf_frame, text="No PDF selected", bg="white", fg="#555555",
        wraplength=400, justify="left",
    )
    pdf_label.pack(side="left", padx=12)

    tk.Label(
        main, text="Number of Questions (1–50)",
        font=("Arial", 11, "bold"), bg="white",
    ).pack(anchor="w", pady=(15, 5))
    question_count = tk.Spinbox(main, from_=1, to=50, width=10, font=("Arial", 11))
    question_count.delete(0, "end")
    question_count.insert(0, "10")
    question_count.pack(anchor="w")

    tk.Label(
        main, text="Question Difficulty", font=("Arial", 11, "bold"),
        bg="white",
    ).pack(anchor="w", pady=(15, 5))
    difficulty_combo = ttk.Combobox(
        main,
        values=["Beginner", "Intermediate", "Advanced", "Expert"],
        state="readonly",
        width=25,
    )
    difficulty_combo.set("Intermediate")
    difficulty_combo.pack(anchor="w")

    tk.Label(
        main, text="Topic / Chapter Focus (Optional)",
        font=("Arial", 11, "bold"), bg="white",
    ).pack(anchor="w", pady=(15, 5))
    topic_text = ScrolledText(main, height=4, font=("Arial", 10), wrap="word")
    topic_text.pack(fill="x", pady=5)
    topic_text.insert("1.0", "Leave blank to cover the entire PDF.")

    generate_button = tk.Button(
        main, text="GENERATE QUIZ", command=start_generation,
        font=("Arial", 12, "bold"), bg="#2874a6", fg="white",
        activebackground="#1b4f72", activeforeground="white",
        padx=20, pady=10, cursor="hand2",
    )
    generate_button.pack(pady=(18, 10))

    progress = ttk.Progressbar(main, mode="indeterminate", length=400)
    progress.pack(pady=5)
    status_label = tk.Label(
        main, text="Ready. Select a PDF to begin.",
        font=("Arial", 10), bg="white", fg="#333333", wraplength=600,
    )
    status_label.pack(pady=8)

    tk.Label(
        main, text="After generating the quiz:",
        font=("Arial", 11, "bold"), bg="white",
    ).pack(anchor="w", pady=(10, 5))

    download_all_button = tk.Button(
        main, text="Download All Questions & Answers (.docx)",
        command=lambda: download_all_questions_answers(quiz_data),
        font=("Arial", 10, "bold"), bg="#217346", fg="white",
        padx=10, pady=8, state="disabled", cursor="hand2",
    )
    download_all_button.pack(fill="x", pady=4)

    download_questions_button = tk.Button(
        main, text="Download Questions Only (.docx)",
        command=lambda: download_questions_only(quiz_data),
        font=("Arial", 10), bg="#5d6d7e", fg="white",
        padx=10, pady=7, state="disabled", cursor="hand2",
    )
    download_questions_button.pack(fill="x", pady=4)

    take_quiz_button = tk.Button(
        main, text="Take Interactive Quiz",
        command=lambda: open_exam_window(quiz_data),
        font=("Arial", 10, "bold"), bg="#8e44ad", fg="white",
        padx=10, pady=8, state="disabled", cursor="hand2",
    )
    take_quiz_button.pack(fill="x", pady=4)

    tk.Label(
        root, text="Powered by Google Gemini", font=("Arial", 9),
        bg="#f2f4f8", fg="#666666",
    ).pack(pady=5)

    root.mainloop()


if __name__ == "__main__":
    create_app()
