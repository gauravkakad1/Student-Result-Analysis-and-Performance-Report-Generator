import PyPDF2
import csv

def convert_pdf_to_csv():
    input_pdf_file = "2019PatternResult.pdf"
    output_txt_file = "SampleFile.txt"
    output_csv_file = "Result.csv"

    with open(input_pdf_file, "rb") as pdf_file, open(output_txt_file, "w") as txt_file:
        reader = PyPDF2.PdfReader(pdf_file)
        for page in reader.pages:
            text = page.extract_text()
            text = (
                text.replace("NAME :", "  NAME ")
                .replace(",", " ")
                .replace("*", " ")
                .replace(":", "  ")
            )
            text = (
                text.replace("%", "% ")
                .replace("ORD", " ORD")
                .replace("CLG", " CLG")
                .replace("SGPA", "  SGPA")
            )
            text = text.replace("/", "  ")
            txt_file.write(text)

    with open(output_txt_file, "r") as file:
        lines = file.readlines()

    with open(output_txt_file, "w") as output:
        for line in lines:
            if all(
                key not in line
                for key in ["PAGE", "COLLEGE", "BRANCH", "CONFIDENTIAL", "SEM"]
            ) and not line.startswith("."):
                line = line.replace("  ", ",")
                line = line.replace("..", ".")
                output.write(line)

    with open(output_txt_file, "r") as f_in, open(
        output_csv_file, "w", newline=""
    ) as f_out:
        csv_writer = csv.writer(f_out, delimiter=",")
        for line in f_in:
            row = line.strip().split(",")
            csv_writer.writerow(row)


# def Convert_to_ReqTable() :
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
                if dm_tut_index < len(row) - 1 :
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

                with open("SampleFile.txt", "a") as f :
                    f.write(seat_no + "," + name + "," + sgpa + "," 
                            +
                            dm_ise + "," +
                            dm_ese + "," + dm_total + "," + dm_tot_per + "," + dm_grd + "," 
                            + dm_tut_tw + "," + dm_tut_tot_per  + "," + dm_tut_grd + "," 
                            + ldco_ise + "," + ldco_ese + "," + ldco_total + ","+ ldco_tot_per + "," + ldco_grd + "," 
                            + dsa_ise + "," + dsa_ese + "," + dsa_total + "," + dsa_tot_per + "," + dsa_grd + "," 
                            + oop_ise + "," + oop_ese + "," + oop_total + "," + oop_tot_per + ","+ oop_grd + "," 
                            + bcn_ise + "," + bcn_ese + "," + bcn_total + "," + bcn_tot_per + ","+ bcn_grd + ","
                            + ldco_lab_tw + "," + ldco_lab_pr + "," + ldco_lab_tot_per + "," + ldco_lab_grd + ","
                            + dsa_lab_tw + "," + dsa_lab_pr + "," + dsa_lab_tot_per + "," + dsa_lab_grd + ","
                            + oop_lab_tw + "," + oop_lab_pr + "," + oop_lab_tot_per + "," + oop_lab_grd + ","
                            + ssl_lab_tw + "," + ssl_lab_tot_per + "," + ssl_lab_grd + ","
                            + audit_grd 
                            +"\n")

    input_file = "SampleFile.txt"
    output_file = "REQ_Table.csv"
    delimiter = ","
    with open(input_file, "r") as f_in, open(output_file, "w", newline="") as f_out:
        csv_writer = csv.writer(f_out, delimiter=",")
        for line in f_in:
            row = line.strip().split(delimiter)
            csv_writer.writerow(row)


if __name__ == "__main__":
    convert_pdf_to_csv()
    # Convert_to_ReqTable()
