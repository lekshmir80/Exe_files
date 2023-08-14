import tkinter as tk
from tkinter import filedialog, messagebox
import openpyxl
import pandas as pd
import jinja2
import os
import logging
import re


class ExcelKeywordApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Parser 1.0")
        self.input_file_path = tk.StringVar()
        self.selected_file_label = tk.StringVar()
        self.signal_component_dict = {}
        self.d_f = None
        self.create_widgets()
        self.log_file_path = "app_log.txt"
        logging.basicConfig(filename=self.log_file_path, level=logging.INFO,
                            format='%(asctime)s %(levelname)s : %(message)s')
        self.log = logging.getLogger("ExcelKeywordApp")
        self.log.setLevel(logging.INFO)
        self.log.info("Action: Logging started")

    def log(self, message):
        logging.info(message)

    def create_widgets(self):
        # tk.Label(self.root, text="Select Input File:").pack()
        tk.Button(self.root, text="Choose File",
                  command=self.browse_file).pack()
        tk.Label(self.root, textvariable=self.selected_file_label).pack()
        tk.Label(self.root, text="Output:").pack()

        self.output_text = tk.Text(self.root, height=10, width=40)
        self.output_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        # if not self.output_scrollbar:
        self.output_scrollbar = tk.Scrollbar(self.root,
                                             command=self.output_text.yview)
        self.output_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.output_text.configure(yscrollcommand=self.output_scrollbar.set)

        tk.Button(self.root, text="Convert", command=self.convert).pack()
        tk.Button(self.root, text="Close", command=self.close_app).pack()

    def read_config_file(self):
        # Read the Excel file
        excel_file = "LabCar_Keywords.xlsx"
        d_f = pd.read_excel(excel_file)
        # Convert columns to dictionary
        self.signal_component_dict = dict(zip(d_f['Signal'], d_f['Component']))
        self.log.info("Action: Configuration reads completed")

    def browse_file(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("Excel Files", "*.xlsx;")])
        self.input_file_path.set(file_path)
        if file_path:
            self.d_f = pd.read_excel(file_path)
        self.log.info(f"Action: Input file selected, File: {file_path}")
        if self.input_file_path.get():
            self.selected_file_label.set(f"Selected File: {file_path}")

    def extract_test_case_number(self, test_case_id):
        match = re.search(r'\d+', test_case_id)
        return match.group() if match else None

    def convert(self):
        input_path = self.input_file_path.get()
        if not input_path:
            self.output_text.insert(tk.END,
                                    "Please select an input file first.\n")
            self.log.info("Action: Please select an input file first.")
            return
        if self.d_f is None:
            self.output_text.insert(tk.END,
                                    "Please select an input file first.\n")
            self.log.info("Action: Please select an input file first.")
            return
        try:
            self.read_config_file()
            result = {}
            final_res = []
            # Iterate through rows and perform comparison
            for index, row in self.d_f.iterrows():
                row['Reference'] = row[
                    'Test Case ID']  # Prepend component to test case ID
                if isinstance(row['Test Steps'], str):
                    designation_words = row['Test Steps'].split()
                    word_check = True
                    for word in designation_words:
                        if word in self.signal_component_dict:
                            component = self.signal_component_dict.get(word)
                            if component and word_check:
                                if result.get(component) is None:
                                    result[component] = {'High': [],
                                                         'Medium': [],
                                                         'Low': []
                                                         }
                                word_check = False
                                priority = row['Test Priority']
                                row['Test Case ID'] = f'TC-{component}-'
                                result[component][priority].append(
                                    row.to_dict())
                        elif word in self.signal_component_dict.keys():
                            final_res.append(row)
                else:
                    final_res.append(row)
            self.generate_output(result, final_res)

        except KeyError as e:
            self.output_text.insert(tk.END, f"Invalid Input File - The input file does not contain the {e} column.\n")
            self.log.error(f"Invalid Input File - The input file does not contain the {e} column.\n")
        except Exception as e:
            self.output_text.insert(tk.END, f"Error: {str(e)}\n")
            self.log.error(f"Error: {str(e)}\n")

    def generate_output(self, result, final_res):
        try:
            output_data = []
            for component, priorities in result.items():
                if any(rows for rows in priorities.values()):
                    output_data.append([component, '', '', '', '', '', '', '', ''])
                    ic = 00
                    for priority, rows in priorities.items():

                        for row in rows:
                            ic += 1
                            output_data.append(['', row['Reference'],
                                                row['Test Case ID'] + str(ic),
                                                row['Test Priority'],
                                                row['Test Type'],
                                                row['Description'],
                                                row['Preconditions'],
                                                row['Test Steps'],
                                                row['Expected Results']])
            output_df = pd.DataFrame(output_data, columns=['Component',
                                                        'Reference',
                                                        'Test Case ID',
                                                        'Test Priority',
                                                        'Test Type',
                                                        'Description',
                                                        'Preconditions',
                                                        'Test Steps',
                                                        'Expected Results'])
            failed_df = pd.DataFrame(final_res,
                                    columns=['Test Case ID', 'Test Priority',
                                            'Test Type', 'Description',
                                            'Preconditions', 'Test Steps',
                                            'Expected Results'])

            # Write the DataFrames to the same sheet using ExcelWriter,csv and html
            self.process_excel(output_df, failed_df)
            self.process_csv(output_df, failed_df)
            self.process_html(output_df, failed_df)
            # write to dialogue box
            formatted_output = self.format_output_data(output_data)
            # Insert the formatted output into the Text widget
            self.output_text.insert(tk.END, formatted_output + "\n\n")
            self.output_text.see(tk.END)  # Scroll to the end

            message = f"Reports Generated \n CSV File: {os.path.abspath('output_results.csv')} \n Excel File: {os.path.abspath('output_results.xlsx')} \n HTML File: {os.path.abspath('output_results.html')}"
            messagebox.showinfo("Files Generated", message)
            self.log.info(
                f"Action: Generated results in csv,xlsx and html formatCSV File: {os.path.abspath('output_results.csv')} \n Excel File: {os.path.abspath('output_results.xlsx')} \n HTML File: {os.path.abspath('output_results.html')}")
        except KeyError as e:
            self.output_text.insert(tk.END, f"Invalid  File - The  file does not contain the {e} column.\n")
            self.log.error(f"Invalid  File - The  file does not contain the {e} column.\n")
        except Exception as e:
            self.output_text.insert(tk.END, f"Error: {str(e)}\n")
            self.log.error(f"Error: {str(e)}\n")

    def process_excel(self, output_df, failed_df):
        with pd.ExcelWriter('output_results.xlsx',
                            engine='xlsxwriter') as writer:
            # Write the first DataFrame with header
            output_df.to_excel(writer, sheet_name='Sheet1', startrow=0,
                               index=False, header=True)
            # Write the caption for df2
            writer.sheets['Sheet1'].write(output_df.shape[0] + 2, 0,
                                          'Failed Conditions',
                                          writer.book.add_format(
                                              {'bold': True}))
            # Write the second DataFrame without header, starting below the first DataFrame
            failed_df.to_excel(writer, sheet_name='Sheet1',
                               startrow=output_df.shape[0] + 3, index=False,
                               header=True)

    def process_csv(self, output_df, failed_df):
        # Write the DataFrames to the same CSV file
        with open('output_results.csv', 'w') as f:
            # Remove spaces within data in DataFrames
            output_df = output_df.applymap(
                lambda x: x.replace(' ', '') if isinstance(x, str) else x)
            failed_df = failed_df.applymap(
                lambda x: x.replace(' ', '') if isinstance(x, str) else x)

            # Write the header and data of the first DataFrame
            output_df.to_csv(f, index=False)

            # Write the header and data of the second DataFrame
            f.write("Failed Conditions")
            failed_df.to_csv(f, index=False)

    def process_html(self, output_df, failed_df):
        # Write the combined HTML to a file
        with open('output_results.html', 'w') as f:
            # Convert DataFrames to HTML
            html_df1 = output_df.to_html(index=False)
            html_df2 = failed_df.to_html(index=False)

            # Combine HTML strings with a separator
            combined_html = f"{html_df1}<hr><b><caption>Failed Conditions</caption></b>{html_df2}"  # You can adjust the separator as needed

            f.write(combined_html)

    def format_output_data(self, output_data):
        formatted_output = ""
        for row in output_data:
            if row[0]:  # Component name
                formatted_output += f"\n{row[0]}"
            else:  # Test case details
                formatted_output += "\n" + "\t".join(row[1:]) + "\n"

        return formatted_output

    def close_app(self):
        self.root.destroy()
        self.log.info(
            "Action:{star}Application closed{star}".format(star='*' * 20))


def main():
    # Start Application
    root = tk.Tk()
    app = ExcelKeywordApp(root)
    app.log.info(
        "Action:{star}Application Started{star}".format(star='*' * 20))
    root.mainloop()


if __name__ == "__main__":
    main()
