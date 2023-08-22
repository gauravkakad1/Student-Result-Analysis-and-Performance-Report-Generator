import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import csv

# import convert
import os

# from convert import Convert_extract_data
import PyPDF2
import csv
import re

# from PIL import Image, ImageTk


selected_file = None


def Convert_extract_data(f):
    # reading pdf file
    global selected_file
    selected_file = f

    with open(selected_file, "rb") as pdfFile:
        reader = PyPDF2.PdfReader(pdfFile)
        n = len(reader.pages)
        for i in range(n):
            page1 = reader.pages[i]
            text = page1.extract_text()
            # print(text)

            with open("SampleFile.txt", "a") as f:
                f.write(text)
            # replace unwanted word from the txt file
            with open("SampleFile.txt", "r") as f:
                text = f.read()
            text = text.replace("NAME :", "  NAME ")
            text = text.replace(",", " ")
            text = text.replace("*", " ")
            text = text.replace(":", "  ")
            text = text.replace("%", "% ")
            text = text.replace("ORD", " ORD")
            text = text.replace("CLG", " CLG")
            text = text.replace("SGPA", "  SGPA")
            text = text.replace("/", "  ")

            with open("SampleFile.txt", "w") as f:
                f.write(text)

    with open("SampleFile.txt", "r") as file:
        lines = file.readlines()
    # remove unwanted line from txt file
    with open("SampleFile.txt", "w") as output:
        for line in lines:
            if (
                ("PAGE") not in line
                and ("COLLEGE") not in line
                and ("BRANCH") not in line
                and ("CONFIDENTIAL") not in line
                and ("SEM") not in line
            ):
                if not line.startswith("."):
                    line = re.sub(
                        r"\s{2,}", ",", line
                    )  # replace multiple spaces by comma
                    line = re.sub(r"\.{2,}", ".", line)
                    output.write(line)
    # converting txt file into csv using comma as delimiter
    input_file = "SampleFile.txt"
    output_file = "Result.csv"
    delimiter = ","
    with open(input_file, "r") as f_in, open(output_file, "w", newline="") as f_out:
        csv_writer = csv.writer(f_out, delimiter=",")
        for line in f_in:
            row = line.strip().split(delimiter)
            csv_writer.writerow(row)

        # extracting required data from extracted csv file and representing it in csv file for further operations

    with open("SampleFile.txt", "w") as f:
        f.write(
            "SEAT NO,NAME,SGPA,DM_ISE,DM_ESE,DM_TOTAL,DM_TOT%,DM_GRD,DM_TUT_TW,DM_TUT_TOT%,DM_TUT_GRD,LDCO_ISE,LDCO_ESE,LDCO_TOTAL,LDCO_TOT%,LDCO_GRD,DSA_ISE,DSA_ESE,DSA_TOTAL,DSA_TOT%,DSA_GRADE,OOP_ISE,OOP_ESE,OOP_TOTAL,OOP_TOT%,OOP_GRADE,BCN_ISE,BCN_ESE,BCN_TOTAL,BCN_TOT%,BCN_GRADE,LDCO_LAB_TW,LDCO_LAB_PR,LDCO_LAB_TOT%,LDCO_LAB_GRD,DSA_LAB_TW,DSA_LAB_PR,DSA_LAB_TOT%,DSA_LAB_GRD,OOP_LAB_TW,OOP_LAB_PR,OOP_LAB_TOT%,OOP_LAB_GRD,SSL_LAB_TW,SSL_LAB_TOT%,SSL_LAB_GRD,AUDIT_GRD\n"
        )
    with open("Result.csv", "r", newline="") as csvfile:
        reader = csv.reader(csvfile)

        for row in reader:
            if "SEAT NO." in row:
                seat_no_index = row.index("SEAT NO.")
                if seat_no_index < len(row) - 1:
                    seat_no = row[seat_no_index + 1]
                    name = row[seat_no_index + 3]

            elif "214441 DISCRETE MATHEMATICS" in row and "03" in row:
                dm_index = row.index("214441 DISCRETE MATHEMATICS")
                if dm_index < len(row) - 1:
                    dm_ise = row[dm_index + 1]
                    dm_ese = row[dm_index + 3]
                    dm_total = row[dm_index + 5]
                    dm_tot_per = row[dm_index + 10]
                    dm_grd = row[dm_index + 12]

            elif "214441 DISCRETE MATHEMATICS" in row and "01" in row:
                dm_tut_index = row.index("214441 DISCRETE MATHEMATICS")
                if dm_tut_index < len(row) - 1:
                    dm_tut_tw = row[dm_tut_index + 4]
                    dm_tut_tot_per = row[dm_tut_index + 8]
                    dm_tut_grd = row[dm_tut_index + 10]

            elif "214442 LOGIC DESIGN & COMP. ORG." in row:
                ldco_index = row.index("214442 LOGIC DESIGN & COMP. ORG.")
                if ldco_index < len(row) - 1:
                    ldco_ise = row[ldco_index + 1]
                    ldco_ese = row[ldco_index + 3]
                    ldco_total = row[ldco_index + 5]
                    ldco_tot_per = row[ldco_index + 10]
                    ldco_grd = row[ldco_index + 12]

            elif "214443 DATA STRUCTURES & ALGO." in row:
                dsa_index = row.index("214443 DATA STRUCTURES & ALGO.")
                if dsa_index < len(row) - 1:
                    dsa_ise = row[dsa_index + 1]
                    dsa_ese = row[dsa_index + 3]
                    dsa_total = row[dsa_index + 5]
                    dsa_tot_per = row[dsa_index + 10]
                    dsa_grd = row[dsa_index + 12]

            elif "214444 OBJECT ORIENTED PROGRAMMING" in row:
                oop_index = row.index("214444 OBJECT ORIENTED PROGRAMMING")
                if oop_index < len(row) - 1:
                    oop_ise = row[oop_index + 1]
                    oop_ese = row[oop_index + 3]
                    oop_total = row[oop_index + 5]
                    oop_tot_per = row[oop_index + 10]
                    oop_grd = row[oop_index + 12]

            elif "214445 BASIC OF COMPUTER NETWORK" in row:
                bcn_index = row.index("214445 BASIC OF COMPUTER NETWORK")
                if bcn_index < len(row) - 1:
                    bcn_ise = row[bcn_index + 1]
                    bcn_ese = row[bcn_index + 3]
                    bcn_total = row[bcn_index + 5]
                    bcn_tot_per = row[bcn_index + 10]
                    bcn_grd = row[bcn_index + 12]

            elif "214446 LOGIC DESIGN COMP. ORG. LAB" in row:
                ldco_lab_index = row.index("214446 LOGIC DESIGN COMP. ORG. LAB")
                if ldco_lab_index < len(row) - 1:
                    ldco_lab_tw = row[ldco_lab_index + 4]
                    ldco_lab_pr = row[ldco_lab_index + 6]
                    ldco_lab_tot_per = row[ldco_lab_index + 9]
                    ldco_lab_grd = row[ldco_lab_index + 11]

            elif "214447 DATA STRUCTURES & ALGO. LAB" in row:
                dsa_lab_index = row.index("214447 DATA STRUCTURES & ALGO. LAB")
                if dsa_lab_index < len(row) - 1:
                    dsa_lab_tw = row[dsa_lab_index + 4]
                    dsa_lab_pr = row[dsa_lab_index + 6]
                    dsa_lab_tot_per = row[dsa_lab_index + 9]
                    dsa_lab_grd = row[dsa_lab_index + 11]

            elif "214448 OBJECT ORIENTED PROG. LAB" in row:
                oop_lab_index = row.index("214448 OBJECT ORIENTED PROG. LAB")
                if oop_lab_index < len(row) - 1:
                    oop_lab_tw = row[oop_lab_index + 4]
                    oop_lab_pr = row[oop_lab_index + 6]
                    oop_lab_tot_per = row[oop_lab_index + 9]
                    oop_lab_grd = row[oop_lab_index + 11]

            elif "214449 SOFT SKILL LAB" in row:
                ssl_lab_index = row.index("214449 SOFT SKILL LAB")
                if ssl_lab_index < len(row) - 1:
                    ssl_lab_tw = row[ssl_lab_index + 4]
                    ssl_lab_tot_per = row[ssl_lab_index + 8]
                    ssl_lab_grd = row[ssl_lab_index + 10]

            elif "214450C LANGUAGE STUDY- JPN MODULE" in row:
                audit_index = row.index("214450C LANGUAGE STUDY- JPN MODULE")
                if audit_index < len(row) - 1:
                    audit_grd = row[audit_index + 9]

            elif "SGPA1" in row:
                sgpa_index = row.index("SGPA1")
                if sgpa_index < len(row) - 1:
                    sgpa = row[sgpa_index + 1]

                with open("SampleFile.txt", "a") as f:
                    f.write(
                        seat_no
                        + ","
                        + name
                        + ","
                        + sgpa
                        + ","
                        + dm_ise
                        + ","
                        + dm_ese
                        + ","
                        + dm_total
                        + ","
                        + dm_tot_per
                        + ","
                        + dm_grd
                        + ","
                        + dm_tut_tw
                        + ","
                        + dm_tut_tot_per
                        + ","
                        + dm_tut_grd
                        + ","
                        + ldco_ise
                        + ","
                        + ldco_ese
                        + ","
                        + ldco_total
                        + ","
                        + ldco_tot_per
                        + ","
                        + ldco_grd
                        + ","
                        + dsa_ise
                        + ","
                        + dsa_ese
                        + ","
                        + dsa_total
                        + ","
                        + dsa_tot_per
                        + ","
                        + dsa_grd
                        + ","
                        + oop_ise
                        + ","
                        + oop_ese
                        + ","
                        + oop_total
                        + ","
                        + oop_tot_per
                        + ","
                        + oop_grd
                        + ","
                        + bcn_ise
                        + ","
                        + bcn_ese
                        + ","
                        + bcn_total
                        + ","
                        + bcn_tot_per
                        + ","
                        + bcn_grd
                        + ","
                        + ldco_lab_tw
                        + ","
                        + ldco_lab_pr
                        + ","
                        + ldco_lab_tot_per
                        + ","
                        + ldco_lab_grd
                        + ","
                        + dsa_lab_tw
                        + ","
                        + dsa_lab_pr
                        + ","
                        + dsa_lab_tot_per
                        + ","
                        + dsa_lab_grd
                        + ","
                        + oop_lab_tw
                        + ","
                        + oop_lab_pr
                        + ","
                        + oop_lab_tot_per
                        + ","
                        + oop_lab_grd
                        + ","
                        + ssl_lab_tw
                        + ","
                        + ssl_lab_tot_per
                        + ","
                        + ssl_lab_grd
                        + ","
                        + audit_grd
                        + "\n"
                    )

    input_file = "SampleFile.txt"
    output_file = "Result.csv"
    delimiter = ","
    with open(input_file, "r") as f_in, open(output_file, "w", newline="") as f_out:
        csv_writer = csv.writer(f_out, delimiter=",")
        for line in f_in:
            row = line.strip().split(delimiter)
            csv_writer.writerow(row)

    global csv_file
    csv_file = output_file

    global data
    data = load_data(csv_file)


def select_pdf_file():
    # Prompt the user to select a PDF file
    file_path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
    file_name = os.path.basename(file_path)
    global selected_file
    selected_file = file_name
    Convert_extract_data(selected_file)
    if file_path:
        pdf_file_entry.delete(0, tk.END)
        pdf_file_entry.insert(tk.END, file_path)


def display_csv_data():
    # Get the filename from the entry widget
    # filename = 'Result.csv'
    global filename

    if filename:
        try:
            rows = len(data)
            cols = len(data[0])
            # Get the screen width and height
            screen_width = root.winfo_screenwidth()
            screen_height = root.winfo_screenheight()

            # Calculate the size of the table based on screen resolution
            table_width = int(screen_width * 0.8)
            table_height = int(screen_height * 0.8)

            # Set the size and position of the root window
            root.geometry(
                f"{table_width}x{table_height}+{int(screen_width - table_width)}+{int(screen_height - table_height)}"
            )

            # Create a frame to hold the table and scrollbars
            table_frame = tk.Frame(root)
            table_frame.pack(pady=10, fill=tk.BOTH, expand=True)

            # Create a canvas for the table
            canvas = tk.Canvas(table_frame)
            canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

            # Create a vertical scrollbar
            y_scrollbar = tk.Scrollbar(
                table_frame, orient=tk.VERTICAL, command=canvas.yview
            )
            y_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

            # Create a horizontal scrollbar
            x_scrollbar = tk.Scrollbar(root, orient=tk.HORIZONTAL, command=canvas.xview)
            x_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)

            # Configure the canvas to work with the scrollbars
            canvas.configure(
                yscrollcommand=y_scrollbar.set, xscrollcommand=x_scrollbar.set
            )
            canvas.bind(
                "<Configure>",
                lambda e: canvas.configure(scrollregion=canvas.bbox("all")),
            )

            # Create a frame inside the canvas to hold the table content
            table_content_frame = tk.Frame(canvas)
            canvas.create_window((0, 0), window=table_content_frame, anchor="nw")

            # Create table headers
            for j in range(cols):
                header_label = tk.Label(
                    table_content_frame, text=data[0][j], relief=tk.RIDGE, width=35
                )
                header_label.grid(row=0, column=j, padx=1, pady=1)

            # Populate table rows
            for i in range(1, rows):
                # Check if the row has the expected number of columns
                if len(data[i]) == cols:
                    for j in range(cols):
                        cell_label = tk.Label(
                            table_content_frame,
                            text=data[i][j],
                            relief=tk.RIDGE,
                            width=35,
                        )
                        cell_label.grid(row=i, column=j, padx=1, pady=1)
                else:
                    print(f"Ignoring row {i} due to inconsistent number of columns.")

            # Configure the canvas to scroll
            canvas.configure(scrollregion=canvas.bbox("all"))

            # Create a button to save the data as an Excel file
            save_button = tk.Button(
                root,
                text="Save as Excel",
                command=lambda: save_as_excel(data),
                bg="#EEEEAD",
                fg="black",
                font=("Arial", 10, "normal"),
                relief=tk.RAISED,
                bd=3,
                padx=10,
                pady=5,
            )
            save_button.pack(pady=10)

        except Exception as e:
            tk.messagebox.showerror(
                "Error", f"An error occurred while displaying the data: {e}"
            )
    else:
        tk.messagebox.showwarning("Warning", "Please enter a CSV filename.")


def load_data(f):
    global filename
    filename = f
    data = []
    with open(filename, "r") as file:
        csv_reader = csv.reader(file)
        for row in csv_reader:
            if row:
                data.append(row)
    return data


def save_as_excel(data):
    # Prompt the user to select a save location
    save_path = filedialog.asksaveasfilename(
        defaultextension=".csv", filetypes=[("CSV Files", "*.csv")]
    )

    if save_path:
        try:
            with open(save_path, "w", newline="") as file:
                writer = csv.writer(file)
                writer.writerows(data)
            messagebox.showinfo(
                "Success", "Table data saved to Excel file successfully!"
            )
        except Exception as e:
            messagebox.showerror(
                "Error", f"An error occurred while saving the file: {e}"
            )


# Specify the CSV file path
filename = csv_file = None

# Load the data from the CSV file
data = None


# Convert_extract_data()

# Create the main window
root = tk.Tk()
root.title("Student Result Analysis and Performance Report Generator")
root.geometry("800x600")
root.configure(bg="lightblue")


# # Load and display the first image
# image1 = Image.open("taelogo.png")  # Replace "image1.png" with the path to your image file
# photo1 = ImageTk.PhotoImage(image1)
# label1 = tk.Label(root, image=photo1)
# label1.grid(row=0, column=0)

# Display the title
title_label = tk.Label(
    root,
    text="Student Result Analysis and Performance Report Generator",
    font=("Arial", 16, "bold"),
)
title_label.pack(padx=10)

# # Load and display the second image
# image2 = Image.open("agradenaac.jpg")  # Replace "image2.png" with the path to your image file
# photo2 = ImageTk.PhotoImage(image2)
# label2 = tk.Label(root, image=photo2)
# label2.grid(row=0, column=2)


# Create a label and entry for selecting the PDF file
pdf_file_label = tk.Label(root, text="PDF File:")
pdf_file_label.pack(pady=10)

pdf_file_entry = tk.Entry(root, width=50)
pdf_file_entry.pack(pady=5)

pdf_file_button = tk.Button(
    root,
    text="Select PDF File",
    command=select_pdf_file,
    bg="#EEEEAD",
    fg="black",
    font=("Arial", 10, "normal"),
    relief=tk.RAISED,
    bd=3,
    padx=10,
    pady=5,
)
pdf_file_button.pack(pady=5)

# Create a button to display the CSV data
display_button = tk.Button(
    root,
    text="Display Data",
    command=display_csv_data,
    bg="#EEEEAD",
    fg="black",
    font=("Arial", 10, "normal"),
    relief=tk.RAISED,
    bd=3,
    padx=10,
    pady=5,
)
display_button.pack(pady=10)

# Create a button to download the table as an Excel file
# download_button = tk.Button(root, text="Download as Excel", command=download_as_excel)
# download_button.pack(pady=5)

root.mainloop()
