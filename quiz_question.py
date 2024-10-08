import csv
from docx import Document
from enum import Enum
import re
import os

# lastest
# Enum for answer choices
class AnswerEnum(Enum):
    A = 1
    B = 2
    C = 3
    D = 4


class TrueFalseEnum(Enum):
    TRUE = 1
    FALSE = 2


def get_answer_value(text):
    """
    Convert A) to 1 || B. to 2 || True to 1 || False to 2
    """
    if text.lower() in ["true", "false"]:
        return TrueFalseEnum[text.upper()].value

    match = re.match(r"^[\(\[]?([A-Da-d])[\.\)\]]?", text.strip())
    if match:
        return AnswerEnum[match.group(1).upper()].value

    # Return the text if it doesn't match the expected patterns
    return text + " "


# Initialize the folder for the document files
questions_folder = "questions"

# Iterate over all DOCX files in the questions folder
for filename in os.listdir(questions_folder):
    if filename.endswith(".docx"):
        # Prepare the paths for the DOCX and CSV files
        docx_path = os.path.join(questions_folder, filename)
        csv_file = filename.replace(".docx", ".csv")
        csv_path = os.path.join(questions_folder, csv_file)

        print(f"Processing: {filename}")
        processed_questions = []
        current_question = []

        # Read the document
        doc = Document(docx_path)
        for para in doc.paragraphs:
            item = para.text.strip()

            # Skip empty paragraphs
            if not item:
                continue

            # Check if the paragraph starts with a number indicating a new question
            if item[0].isdigit() and current_question:
                # If a new question starts, save the current question
                processed_questions.append(current_question)
                current_question = [
                    para
                ]  # Start a new question (save paragraph object for later processing)
            else:
                current_question.append(para)

        # Append the last question after the loop if it exists
        if current_question:
            processed_questions.append(current_question)

        # Write all processed questions to the CSV file
        with open(csv_path, mode="w", newline="", encoding="utf-8") as file:
            writer = csv.writer(file, quoting=csv.QUOTE_ALL)

            # Write header with an additional 'Explanation' column and Question Number
            writer.writerow(
                [
                    "S.No",
                    "Question",
                    "Option A",
                    "Option B",
                    "Option C",
                    "Option D",
                    "Answer",
                    "Explanation",
                ]
            )

            for question_paragraphs in processed_questions:
                
                question_number = question_paragraphs[0].text.strip().split(" ", 1)[0]
                question_text = question_paragraphs[0].text.strip().split(" ", 1)[1]
                
                question_paragraphs[0].text = "Question"

                answer = ""
                explanation = ""

                # Process each paragraph in the current question group
                answer_found = False
                for para in question_paragraphs:
                    for run in para.runs:
                        if run.italic and not answer_found:
                            current_answer = run.text.strip()  # Capture the italic answer
                            current_answer = current_answer.split(" ", 1)[0]
                            answer = get_answer_value(current_answer)
                            answer_found = True
                        elif 'Correct Answer' in run.text:
                            answer = run.text.split(":", 1)[-1]
                            answer = get_answer_value(answer)
                        else:
                            explanation += run.text.strip() + " "
                            section_index = explanation.find("Section") or explanation.find("Lesson")
                            if section_index != -1:
                                explanation = explanation[section_index:].strip()

                # Check if it's a True/False type question
                if "True or False" in question_text:
                    question_parts = question_text.split(": ", 1)

                    main_question = (
                        question_parts[1] if len(question_parts) > 1 else question_text
                    )

                    main_question = (
                        main_question.replace("true", "")
                        .replace("false", "")
                        .replace("True", "")
                        .replace("False", "")
                        .strip()
                    )

                    # Remove "True or False" from the question text
                    if main_question.startswith("True or False"):
                        main_question = main_question.replace("True or False: ", "")

                    # Prepare the row with the True/False options
                    row = [
                        question_number,
                        f"{main_question}",
                        "True",
                        "False",
                        "",
                        "",
                        answer,
                        explanation,
                    ]
                else:
                    row = [question_number, question_text]  # Start with the question number and question text
                    if len(question_paragraphs) == 5:
                        options = [
                            opt.text.strip().split(" ", 1)[1]
                            for opt in question_paragraphs[1:4]
                        ]  # Get up to 3 options
                        row.extend(
                            options + [""] * (4 - len(options))
                        )  # Pad with empty strings if needed

                        row.extend([answer, explanation])
                    elif len(question_paragraphs) == 3:
                        options = ['True', 'False']
                        row.extend(
                            options + [""] * (4 - len(options))
                        )  # Pad with empty strings if needed

                        row.extend([answer, explanation])
                    else:
                        options = [
                            opt.text.strip().split(" ", 1)[1]
                            for opt in question_paragraphs[1:5]
                        ]  # Get up to 4 options
                        row.extend(
                            options + [""] * (4 - len(options))
                        )
                        row.extend([answer, explanation])

                writer.writerow(row)

        print(f"Data has been written to {csv_file}")

print("Processing completed for all files.")
