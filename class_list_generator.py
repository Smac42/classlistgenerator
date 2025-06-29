
import pandas as pd
import random
from collections import defaultdict
from tkinter import Tk, simpledialog, messagebox, filedialog
import os

REQUIRED_COLUMNS = {
    "Name", "Gender", "Program", "NeedsSA", "NeedsIRT", "Behaviour", "EAL", "TogetherGroup"
}

VALID_PROGRAMS = {"French", "English"}

def load_students(file_path):
    df = pd.read_excel(file_path)

    missing = REQUIRED_COLUMNS - set(df.columns)
    if missing:
        raise ValueError(f"Missing required columns: {missing}")

    # Validate Program values
    invalid_programs = df[~df["Program"].str.capitalize().isin(VALID_PROGRAMS)]
    if not invalid_programs.empty:
        raise ValueError(f"Invalid Program values found: {invalid_programs['Program'].unique()}")

    # Normalize Program values (capitalize)
    df["Program"] = df["Program"].str.capitalize()

    # Normalize boolean columns
    for col in ["NeedsSA", "NeedsIRT", "Behaviour", "EAL"]:
        df[col] = df[col].astype(bool)

    # Fill TogetherGroup with -1 if missing
    df["TogetherGroup"] = df["TogetherGroup"].fillna(-1).astype(int)

    return df

def initialize_classes(num_classes):
    return [[] for _ in range(num_classes)]

def get_balanced_groups(df, group_by_columns):
    groups = defaultdict(list)
    for _, row in df.iterrows():
        key = tuple(row[col] for col in group_by_columns)
        groups[key].append(row)
    for key in groups:
        random.shuffle(groups[key])
    return groups

def get_apart_groups_gui():
    root = Tk()
    root.withdraw()
    messagebox.showinfo(
        "Class Generator",
        "Enter 'Apart' groups, one group per line, names separated by commas.\nExample:\nAlice,Bob\nCarol,David"
    )
    apart_raw = simpledialog.askstring("Apart Groups", "")
    apart = [line.strip().split(",") for line in apart_raw.strip().splitlines()] if apart_raw else []
    # Trim whitespace on names
    apart = [[name.strip() for name in group] for group in apart]
    return apart

def assign_students(groups, num_classes, sa_class_ids, irt_class_ids, eal_class_ids, apart_groups=[]):
    class_lists = initialize_classes(num_classes)
    class_counters = [0] * num_classes
    student_class = {}

    # Assign TogetherGroup students together
    together_groups = defaultdict(list)
    for student in [s for sublist in groups.values() for s in sublist]:
        tg = student["TogetherGroup"]
        if tg != -1:
            together_groups[tg].append(student)

    for tg, students in together_groups.items():
        assigned_class = random.choice(range(num_classes))
        for s in students:
            student_class[s["Name"]] = assigned_class
            class_lists[assigned_class].append(s)
            class_counters[assigned_class] += 1

    # Assign apart groups from names given in popup (list of lists of names)
    for apart_group in apart_groups:
        for i, name in enumerate(apart_group):
            class_id = i % num_classes
            # Only assign if not already assigned (e.g., not in together)
            if name in student_class:
                continue
            student_class[name] = class_id
            # Find the student object to append to class_lists
            found_students = [s for sublist in groups.values() for s in sublist if s["Name"] == name]
            if found_students:
                class_lists[class_id].append(found_students[0])
                class_counters[class_id] += 1
            else:
                print(f"Warning: Student name '{name}' in apart group not found in data.")

    # Assign remaining students
    for group_key, students in groups.items():
        for student in students:
            if student["Name"] in student_class:
                continue

            needs_sa = bool(student.get("NeedsSA", False))
            needs_irt = bool(student.get("NeedsIRT", False))
            needs_eal = bool(student.get("EAL", False))
            behaviour = bool(student.get("Behaviour", False))

            eligible_classes = set(range(num_classes))
            if needs_sa:
                eligible_classes &= set(sa_class_ids)
            if needs_irt:
                eligible_classes &= set(irt_class_ids)
            if needs_eal:
                eligible_classes &= set(eal_class_ids)

            if not eligible_classes:
                raise ValueError(f"No eligible class for student {student['Name']} with support needs.")

            # Behaviour students spread evenly
            if behaviour:
                behaviour_counts = [
                    sum(1 for s in class_lists[i] if s["Behaviour"]) for i in eligible_classes
                ]
                min_behaviour_class = list(eligible_classes)[behaviour_counts.index(min(behaviour_counts))]
                class_id = min_behaviour_class
            else:
                class_id = min(eligible_classes, key=lambda i: class_counters[i])

            student_class[student["Name"]] = class_id
            class_lists[class_id].append(student)
            class_counters[class_id] += 1

    return class_lists

def export_to_excel(class_lists, output_file="class_rosters.xlsx"):
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        for idx, class_list in enumerate(class_lists):
            df = pd.DataFrame(class_list)
            df.to_excel(writer, sheet_name=f"Class_{idx+1}", index=False)
    print(f"Exported to {output_file}")

if __name__ == "__main__":
    try:
        root = Tk()
        root.withdraw()
        file_path = filedialog.askopenfilename(title="Select Excel File", filetypes=[("Excel files", "*.xlsx *.xls")])
        if not file_path or not os.path.exists(file_path):
            raise FileNotFoundError("No file selected or file not found.")
        df = load_students(file_path)
        num_classes = simpledialog.askinteger("Number of Classes", "Enter number of classes:", minvalue=1)
        apart_groups = get_apart_groups_gui()

        def spread_indexes(count):
            return list(set(i % num_classes for i in range(count)))

        sa_class_ids = spread_indexes(len(df[df["NeedsSA"] == True]))
        irt_class_ids = spread_indexes(len(df[df["NeedsIRT"] == True]))
        eal_class_ids = spread_indexes(len(df[df["EAL"] == True]))

        group_by = ["Gender", "Program"]
        balanced_groups = get_balanced_groups(df, group_by)

        class_lists = assign_students(balanced_groups, num_classes, sa_class_ids, irt_class_ids, eal_class_ids, apart_groups)

        save_path = filedialog.asksaveasfilename(title="Save Excel File As", defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        export_to_excel(class_lists, output_file=save_path)

        messagebox.showinfo("Success", f"Class lists saved to {save_path}")

    except Exception as e:
        messagebox.showerror("Error", str(e))
