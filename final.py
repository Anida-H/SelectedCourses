import tkinter as tk
from tkinter import ttk, messagebox, END
import csv

class CourseSelectionApp(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("Course Selection App")
        self.geometry("1000x500")
        self.minsize(600, 400)
        self.maxsize(1200, 800)
        self.config(bg="green")
        self.create_widgets()

    def create_widgets(self):
        # Labels
        label_font = ("Helvetica", 12, "bold")
        label_color = "black"
        label_bg_color = "white"

        enter_filepath_label = tk.Label(self, text="Enter the file path:", bg=label_bg_color, width=15, font=label_font,
                                        fg=label_color)
        enter_filepath_label.grid(row=2, column=1, padx=(5, 10), pady=(20, 0), sticky="W")

        year_label = tk.Label(self, text="Year", bg=label_bg_color, width=17, font=label_font, fg=label_color)
        year_label.grid(row=4, column=1, padx=(5, 10), pady=(20, 0), sticky="W")

        department_label = tk.Label(self, width=17, text="Department:", bg=label_bg_color, font=label_font,
                                    fg=label_color)
        department_label.grid(row=4, column=3, padx=(5, 10), pady=(20, 0), sticky="W")

        warnings_label = tk.Label(self, width=17, text="Warnings:", bg=label_bg_color, font=label_font, fg=label_color)
        warnings_label.grid(row=6, column=1, padx=(0, 0), pady=(50, 0), sticky="W")

        courses_label = tk.Label(self, width=17, text="Courses:", bg=label_bg_color, font=label_font, fg=label_color)
        courses_label.grid(row=6, column=4, padx=(0, 0), pady=(50, 0), sticky="W")


        entry_font = ("Helvetica", 10)
        entry_bg_color = "lightgray"

        self.entry_filepath_entry = tk.Entry(self, font=entry_font, bg=entry_bg_color)
        self.entry_filepath_entry.grid(row=2, column=2, padx=(0, 0), pady=(20, 0), columnspan=2, sticky="WE")

        self.department_label_entry = tk.Entry(self, font=entry_font, bg=entry_bg_color)
        self.department_label_entry.grid(row=4, column=4, padx=(0, 0), pady=(20, 0), columnspan=2)

        self.year_combobox = ttk.Combobox(self, values=["1", "2", "3", "4", "5", " "], font=entry_font,
                                          state="readonly")
        self.year_combobox.grid(row=4, column=2, padx=(5, 10), pady=(20, 0))


        button_bg_color = "#4CAF50"  # Green
        button_fg_color = "white"

        display_button = tk.Button(self, text="Display", bg=button_bg_color, fg=button_fg_color,
                                   command=self.display_courses)
        display_button.grid(row=5, column=1, sticky="E", padx=(0, 10), pady=(50, 0))

        clear_button = tk.Button(self, text="Clear", bg=button_bg_color, fg=button_fg_color, command=self.clear)
        clear_button.grid(row=5, column=2, sticky="WE", padx=(0, 10), pady=(50, 0))

        save_button = tk.Button(self, text="Save", bg=button_bg_color, fg=button_fg_color, command=self.save)
        save_button.grid(row=5, column=3, sticky="WE", padx=(0, 10), pady=(50, 0))

        listbox_font = ("Helvetica", 10)
        listbox_bg_color = "lightgray"

        self.warnings = tk.Listbox(self, width=60, font=listbox_font, bg=listbox_bg_color)
        self.warnings.grid(row=7, column=1, columnspan=3, pady=(50, 0), sticky="EW")

        self.courses = tk.Listbox(self, width=60, font=listbox_font, bg=listbox_bg_color)
        self.courses.bind("<Double-1>", self.course_warning)
        self.courses.grid(row=7, column=4, columnspan=3, padx=(10, 0), pady=(50, 0), sticky="EW")

    def display_courses(self):
        filepath = self.entry_filepath_entry.get()
        try:
            with open('C:/Users/Admin/Desktop/ICTFinal/sampledata.csv', encoding="unicode_escape") as file:

                courses = file.readlines()

            department = self.department_label_entry.get().upper()
            year = self.year_combobox.get()

            matching_courses = [course.strip() for course in courses if self.is_matching_course(course, department, year)]

            self.courses.delete(0, END)
            self.courses.insert('end', *matching_courses)

        except FileNotFoundError:
            messagebox.showerror("Error", "File not found.")

    def is_matching_course(self, course, department, year):
        tokens = course.split(' ')
        return (
                tokens[0].startswith(department) and
                tokens[0].endswith(department) and
                tokens[1].startswith(year)
        )

    def clear(self):
        widgets_to_clear = [
            self.entry_filepath_entry,
            self.year_combobox,
            self.department_label_entry,
            self.courses,
            self.warnings
        ]

        for widget in widgets_to_clear:
            widget.delete(0, END)

    list = []

    def course_warning(self, course):
        selected_courses = self.courses.curselection()

        if self.warnings.size() >= 6:
            messagebox.showinfo("Information", "You can choose only 6 courses")
        else:
            for index in selected_courses:
                course_info = self.courses.get(index).split(',')
                course_name = course_info[2].split(' ')[0]

                if len(selected_courses) <= 6 and f"Added {course_info[0]}" not in self.warnings.get(0, END):
                    self.warnings.insert(0, f"Added {course_info[0]}")

            if selected_courses:
                self.list.extend([self.courses.get(index) for index in selected_courses])

    def save(self):
        file_path = "C:\\Users\\Admin\\OneDrive\\Desktop\\Timetable.xlsx"
        try:
            with open(file_path, 'w', newline='') as file:
                csv_writer = csv.writer(file)
                for row in self.list:
                    csv_writer.writerow(row.split(", "))

            messagebox.showinfo("Information", "Data saved successfully.")
        except Exception as e:
            messagebox.showerror("Error", f"Error saving data: {e}")

if __name__ == "__main__":
    app = CourseSelectionApp()
    app.mainloop()
