import pandas as pd
import matplotlib.pyplot as plt
import tkinter as tk
from tkinter import ttk
from tkinter import StringVar
from tkinter import filedialog, messagebox
import os
import random

#Global vraiables
jobs = None
problem = 'Single'

def load_jobs_from_excel(file_path):
    return pd.read_excel(file_path)

def calculate_completion_times(machines, jobs, sequence,rule):
    machine_times = [0] * machines
    job_completion_times = {}
    machine_assignments = [None] * machines  # To track which job is assigned to which machine
    machine_index = 0
    if rule == "Wrap-Around":
        for job in sequence:
            start_time = machine_times[machine_index]
            end_time = start_time + jobs.loc[job, 'process time']
            machine_times[machine_index] = end_time
            job_completion_times[job] = end_time
            machine_assignments[machine_index] = job
            machine_index = (machine_index+1) % machines
    else:
        for job in sequence:
            machine_index = machine_times.index(min(machine_times))
            start_time = machine_times[machine_index]
            end_time = start_time + jobs.loc[job, 'process time']
            machine_times[machine_index] = end_time
            job_completion_times[job] = end_time
            
            machine_assignments[machine_index] = job
    return job_completion_times

def calculate_objectives(jobs, sequence, job_completion_times):
    makespan = max(job_completion_times.values())
    total_completion_time = sum(job_completion_times.values())
    tardiness = [max(0, job_completion_times[job] - jobs.loc[job, 'due date']) for job in sequence]
    total_tardiness = sum(tardiness)
    weight_completion = [job_completion_times[job] * jobs.loc[job, 'weight'] for job in sequence]
    weighted_completion = sum(weight_completion)
    lateness = [1 if job_completion_times[job] > jobs.loc[job, 'due date'] else 0 for job in sequence]
    total_lateness = sum(lateness)
    weighted_lateness = sum(jobs.loc[job, 'weight'] * lateness[i] for i, job in enumerate(sequence))
    weighted_tardiness = sum(jobs.loc[job, 'weight'] * tardiness[i] for i, job in enumerate(sequence))

    result = {"Makespan":makespan,"Total completion time":total_completion_time,"Total tardiness":total_tardiness,"Total weighted completion":weighted_completion,
              "Total lateness (Uj)":total_lateness,"Total weighted lateness (WjUj)":weighted_lateness,"Total weighted tardiness (WjTj)":weighted_tardiness}
    return result

def apply_dispatching_rule(jobs, rule):
    try:
        if rule == "SPT":
            return jobs.sort_values('process time', ascending=True).index.tolist()
        elif rule == "WSPT":
            jobs['priority'] = jobs['process time'] / jobs['weight']
            return jobs.sort_values('priority').index.tolist()
        elif rule == "LPT":
            return jobs.sort_values('process time', ascending=False).index.tolist()
        elif rule == "EDD":
            return jobs.sort_values('due date').index.tolist()
        elif rule == "ERD":
            return jobs.sort_values('release date').index.tolist()
        elif rule == "Wrap-Around":
            wrap_job = jobs.copy()
            return wrap_job.sort_index().index.tolist()
        else:
            raise ValueError("Invalid rule.")
    except KeyError as e:
        raise ValueError(f"Missing required columns in job data: {str(e)}")

def generate_gantt_chart(machines, jobs, sequence, rule):
    machine_times = [0] * machines
    fig, ax = plt.subplots(figsize=(10, 6))
    machine_index = 0
    if rule == "Wrap-Around":
        for job in sequence:
            start_time = machine_times[machine_index]
            end_time = start_time + jobs.loc[job, 'process time']
            ax.barh(f"Machine {machine_index + 1}", end_time - start_time, left=start_time, edgecolor='black')
            ax.text((start_time + end_time) / 2, machine_index, f"Job {job}", ha='center', va='center', color='white')
            machine_times[machine_index] = end_time
            machine_index = (machine_index+1) % machines
    else:
        for job in sequence:
            machine_index = machine_times.index(min(machine_times))
            start_time = machine_times[machine_index]
            end_time = start_time + jobs.loc[job, 'process time']
            ax.barh(f"Machine {machine_index + 1}", end_time - start_time, left=start_time, edgecolor='black')
            ax.text((start_time + end_time) / 2, machine_index, f"Job {job}", ha='center', va='center', color='white')
            machine_times[machine_index] = end_time

    ax.set_xlabel("Time")
    ax.set_title(f"Gantt Chart ({rule} Rule)")
    plt.show()

def generate_gantt_chart_flowshop(machines, jobs, sequence,method):
    machine_times = [0] * machines
    job_completion_times = {job: 0 for job in sequence}

    fig, ax = plt.subplots(figsize=(10, 6))

    for job in sequence:
        start_times = []
        end_times = []

        for machine_index in range(machines):
            start_time = max(
                machine_times[machine_index],
                job_completion_times[job] if machine_index == 0 else end_times[machine_index - 1]
            )
            end_time = start_time + jobs.loc[job, 'process time']

            start_times.append(start_time)
            end_times.append(end_time)

            ax.barh(
                f"Machine {machine_index + 1}",
                end_time - start_time,
                left=start_time,
                edgecolor='black',
                color=f"C{job % 10}"
            )
            ax.text(
                (start_time + end_time) / 2,
                machine_index,
                f"Job {job}",
                ha='center',
                va='center',
                color='white'
            )

            machine_times[machine_index] = end_time

        job_completion_times[job] = end_times[-1]

    ax.set_xlabel("Time")
    ax.set_ylabel("Machines")
    ax.set_title("Gantt Chart for Flow Shop Scheduling for"+method)
    ax.invert_yaxis()
    plt.tight_layout()
    plt.show()

def generate_table(table_frame,tree,result):
    for item in tree.get_children():
        tree.delete(item)
    table_frame.pack_forget()
    tree.pack_forget()
    tree.heading("Key", text="Objectives")
    tree.heading("Value", text="Values")
    tree.column("Key", width=200)
    tree.column("Value", width=300)
    tree.pack_forget()
    for key, value in result.items():
        tree.insert("", tk.END, values=(key, value))
    tree.pack(fill=tk.BOTH, expand=True)
    table_frame.pack()

#Step5 initial solutions:
def generate_initial_solutions(jobs):
    rules = ["SPT", "LPT", "Wrap-Around"]
    initial_solutions = {}

    for rule in rules:
        try:
            # Her kural için sıralama uygulayın ve sonucu saklayın
            sequence = apply_dispatching_rule(jobs, rule)
            initial_solutions[rule] = sequence
        except Exception as e:
            print(f"Error applying rule {rule}: {e}")
            initial_solutions[rule] = []  # Hata durumunda boş bir liste döndür

def random_swap(sequence, num_neighborhood):
    num_neighborhood += 1
    swapped = sequence.copy()
    i, j = random.sample(range(len(swapped)), 2)
    swapped[i], swapped[j] = swapped[j], swapped[i]
    return swapped,num_neighborhood

def local_search(jobs, rule, machines, num_neighborhood,initial_sequence, method,iterations=500, initial_threshold=0):
    best_sequence = initial_sequence
    best_completion_times = calculate_completion_times(machines, jobs, best_sequence, rule)
    best_makespan = max(best_completion_times.values())

    current_sequence = best_sequence
    current_makespan = best_makespan

    threshold = initial_threshold
    for i in range(iterations):
        new_sequence,num_neighborhood = random_swap(current_sequence, num_neighborhood)
        new_completion_times = calculate_completion_times(machines, jobs, new_sequence, rule)
        new_makespan = max(new_completion_times.values())

        if new_makespan < best_makespan:
            best_sequence = new_sequence
            best_completion_times = new_completion_times
            best_makespan = new_makespan
        if method == "Meta-Heuristic":
            if new_makespan < current_makespan + threshold:
                current_sequence = new_sequence
                current_makespan = new_makespan

        # Decrease threshold dynamically
        threshold *= 0.95  # Reduce by 5% each iteration

    return best_sequence, best_makespan, best_completion_times,num_neighborhood

def find_best_solution(jobs, machines,method,threshold=0):
    rules = ["SPT", "LPT", "Wrap-Around"]
    best_overall_makespan = float('inf')
    best_overall_sequence = None
    sequences_rules = {}
    num_neighborhood = 0
    
    for rule in rules:
        initial_sequence = apply_dispatching_rule(jobs, rule)
        seq_name = "Initial_sequence_"+rule
        sequences_rules[seq_name] = initial_sequence

        best_sequence, best_makespan, best_completion_times,num_neighborhood = local_search(
            jobs, rule, machines, num_neighborhood,initial_sequence, method,iterations=500, initial_threshold=threshold
        )

        if best_makespan < best_overall_makespan:
            best_overall_sequence = best_sequence
            best_overall_makespan = best_makespan
            best_overall_time = best_completion_times
    print(f"Number of neighborhood structures is {num_neighborhood}")
    return best_overall_sequence, best_overall_time, sequences_rules







#flowshop scheduling

def calculate_flowshop_completion_times(jobs, machines, sequence):
    completion_times = {}
    num_machines = machines
    num_jobs = len(sequence)

    # Initialize completion times table
    times = [[0] * (num_jobs + 1) for _ in range(num_machines + 1)]

    for i, job in enumerate(sequence, start=1):
        for m in range(1, num_machines + 1):
            times[m][i] = max(times[m - 1][i], times[m][i - 1]) + jobs.loc[job, 'process time']

    for idx, job in enumerate(sequence):
        completion_times[job] = [times[m][idx + 1] for m in range(1, num_machines + 1)]

    return completion_times

def apply_dispatching_rule_flowshop(jobs, rule):
    if rule == "SPT":
        return jobs.sort_values(by='process time').index.tolist()
    elif rule == "LPT":
        return jobs.sort_values(by='process time', ascending=False).index.tolist()
    elif rule == "EDD":  # Earliest Due Date
        return jobs.sort_values(by='due date').index.tolist()
    elif rule == "WSPT":  # Weighted Shortest Processing Time
        return jobs.sort_values(by=lambda x: jobs.loc[x, 'process time'] / jobs.loc[x, 'weight']).index.tolist()
    else:
        raise ValueError("Unknown rule")



def local_search_flowshop(jobs, rule, machines, num_neighborhood,initial_sequence, method, iterations=500, initial_threshold=0):
    best_sequence = initial_sequence
    best_completion_times = calculate_flowshop_completion_times(jobs, machines, best_sequence)
    best_makespan = max(best_completion_times[best_sequence[-1]])

    current_sequence = best_sequence
    current_makespan = best_makespan

    threshold = initial_threshold
    for i in range(iterations):
        new_sequence,num_neighborhood = random_swap(current_sequence,num_neighborhood)
        new_completion_times = calculate_flowshop_completion_times(jobs, machines, new_sequence)
        new_makespan = max(new_completion_times[new_sequence[-1]])

        if new_makespan < best_makespan:
            best_sequence = new_sequence
            best_completion_times = new_completion_times
            best_makespan = new_makespan
        if method == "Meta-Heuristic":
            if new_makespan < current_makespan + threshold:
                current_sequence = new_sequence
                current_makespan = new_makespan

        # Decrease threshold dynamically
        threshold *= 0.95  # Reduce by 5% each iteration

    return best_sequence, best_makespan, best_completion_times,num_neighborhood

def find_best_solution_flowshop(jobs, machines, method, threshold=0):
    rules = ["SPT", "LPT", "EDD"]
    best_overall_makespan = float('inf')
    best_overall_sequence = None
    sequences_rules = {}
    num_neighborhood = 0

    for rule in rules:
        initial_sequence = apply_dispatching_rule_flowshop(jobs, rule)
        seq_name = "Initial_sequence_" + rule
        sequences_rules[seq_name] = initial_sequence
        
        best_sequence, best_makespan, best_completion_times,num_neighborhood = local_search_flowshop(
            jobs, rule, machines,num_neighborhood, initial_sequence, method, iterations=500, initial_threshold=threshold
        )

        if best_makespan < best_overall_makespan:
            best_overall_sequence = best_sequence
            best_overall_makespan = best_makespan
            best_overall_time = best_completion_times

    print(f"Number of neighborhood structures is {num_neighborhood}")
    return best_overall_sequence, best_overall_time, sequences_rules

def calculate_objectives_flowshop(jobs, sequence, job_completion_times):
    makespan = max(job_completion_times[job][-1] for job in sequence)
    total_completion_time = sum(job_completion_times[job][-1] for job in sequence)
    tardiness = [
        max(0, job_completion_times[job][-1] - jobs.loc[job, 'due date']) for job in sequence
    ]
    total_tardiness = sum(tardiness)

    weight_completion = [
        job_completion_times[job][-1] * jobs.loc[job, 'weight'] for job in sequence
    ]
    weighted_completion = sum(weight_completion)
    lateness = [
        1 if job_completion_times[job][-1] > jobs.loc[job, 'due date'] else 0 for job in sequence
    ]
    total_lateness = sum(lateness)
    weighted_lateness = sum(
        jobs.loc[job, 'weight'] * lateness[i] for i, job in enumerate(sequence)
    )
    weighted_tardiness = sum(
        jobs.loc[job, 'weight'] * tardiness[i] for i, job in enumerate(sequence)
    )

    result = {
        "Makespan": makespan,
        "Total completion time": total_completion_time,
        "Total tardiness": total_tardiness,
        "Total weighted completion": weighted_completion,
        "Total lateness (Uj)": total_lateness,
        "Total weighted lateness (WjUj)": weighted_lateness,
        "Total weighted tardiness (WjTj)": weighted_tardiness,
    }
    return result



def main():


    def open_file():
        global jobs
        try:
            file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
            if file_path:
                jobs = load_jobs_from_excel(file_path)
                jobs.columns = jobs.columns.str.strip()
                jobs.index = jobs['Job Number']
                required_columns = ['Job Number', 'process time', 'due date', 'weight']
                missing_columns = [col for col in required_columns if col not in jobs.columns]
                if missing_columns:
                    messagebox.showerror("Missing Columns", f"Excel file is missing the following columns: {', '.join(missing_columns)}")
                    jobs = None
                    return
                messagebox.showinfo("File Loaded", "Excel file loaded successfully!")
                file_name = os.path.basename(file_path)
                excel_label.config(text=f"Selected excel is {file_name}")
        except Exception as e:
            messagebox.showerror("Error Loading File", f"An error occurred while loading the file: {str(e)}")

    def apply_rule():
        global jobs
        if jobs is None:
            messagebox.showwarning("No File Loaded", "Please load an Excel file first.")
            return

        try:
            if problem == 'Single':
                machines = 1
            else:
                machines = int(machine_count.get())
                if machines <= 0:
                    raise ValueError("Number of machines must be greater than 0.")
        except ValueError as e:
            messagebox.showerror("Invalid Input", "Please enter a valid number of machines.")
            return

        rule = rule_var.get()
        if not rule :
            if problem != "Flowshop":
                messagebox.showwarning("No Rule Selected", "Please select a dispatching rule.")
                return
            

        try:
            method = method_var.get()
            
            if problem == "Flowshop":
                if method == "Local Search":
                    best_sequence,best_overall_time,sequences_rules = find_best_solution_flowshop(jobs,machines,method)
                    results = calculate_objectives_flowshop(jobs, best_sequence, best_overall_time)
                    results["Local Search Sequence"] = best_sequence
                    results.update(sequences_rules)
                    generate_table(table_frame,tree,results)
                    generate_gantt_chart_flowshop(machines, jobs, best_sequence,method)
                else:
                    thresh = int(thresh_count.get())
                    if thresh < 0:
                        raise ValueError("Threshold can not be negative.")
                    elif thresh == None:
                        raise ValueError("Threshold is invalid")
                    else:
                        best_sequence,best_overall_time,sequences_rules = find_best_solution_flowshop(jobs,machines,method,threshold=thresh)
                        results = calculate_objectives_flowshop(jobs, best_sequence, best_overall_time)
                        results["Meta-Heuristic Sequence"] = best_sequence
                        results.update(sequences_rules)
                        generate_table(table_frame,tree,results)
                        generate_gantt_chart_flowshop(machines, jobs, best_sequence,method)
            else:
                if method == "None" or None:
                    initial_sequence = apply_dispatching_rule(jobs, rule)
                    job_completion_times = calculate_completion_times(machines, jobs, initial_sequence, rule)
                    results = calculate_objectives(jobs, initial_sequence, job_completion_times)
                    
                    results[rule+" Sequence"] = initial_sequence
                    generate_table(table_frame,tree,results)
                    generate_gantt_chart(machines, jobs, initial_sequence, rule)
                elif method == "Local Search":
                    best_sequence,best_overall_time,sequences_rules = find_best_solution(jobs,machines,method)
                    results = calculate_objectives(jobs, best_sequence, best_overall_time)
                    results["Local Search Sequence"] = best_sequence
                    results.update(sequences_rules)
                    generate_table(table_frame,tree,results)
                    generate_gantt_chart(machines, jobs, best_sequence, method)
                elif method == "Meta-Heuristic":
                    thresh = int(thresh_count.get())
                    if thresh < 0:
                        raise ValueError("Threshold can not be negative.")
                    elif thresh == None:
                        raise ValueError("Threshold is invalid")
                    else:
                        best_sequence,best_overall_time,sequences_rules = find_best_solution(jobs,machines,method,threshold=thresh)
                        results = calculate_objectives(jobs, best_sequence, best_overall_time)
                        results["Meta-Heuristic Sequence"] = best_sequence
                        results.update(sequences_rules)
                        generate_table(table_frame,tree,results)
                        generate_gantt_chart(machines, jobs, best_sequence, method)

        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")

    def all_root_forget():
        #Sıralama karışmasın diye ilk başta forgetting yapıyoruz.
        rule_frame.pack_forget()
        rule_label.grid_forget()
        rule_dropdown.grid_forget()
        rule_start.pack_forget()
        machines_num.pack_forget()
        parallel_frame.pack_forget()
        machines_num.grid_forget()
        machine_count.grid_forget()
        table_frame.pack_forget()
        method_dropdown.grid_forget()
        method_label.grid_forget()
        method_frame.pack_forget()
        thresh_frame.pack_forget()
        thresh_label.grid_forget()
        thresh_count.grid_forget()
    
    def all_root_get():
        rule_frame.pack()
        rule_label.grid(row=0, column=0, padx=5, pady=5)
        rule_dropdown.grid(row=0, column=1, padx=5, pady=5)
        

    
        

    def single_root():
        global problem
        style.configure("Single.TButton", background="green",  foreground="green")
        style.configure("Parallel.TButton", background="#abebc6", foreground="black")
        style.configure("Flowshop.TButton", background="#abebc6",  foreground="black")
        problem_type.config(text="Single mod is selected.")
        problem = 'Single'
        all_root_forget()
        rule_var.set(models[0])
        rule_dropdown.config(values=models[:5])
        all_root_get()
        rule_start.pack(pady=5)
        


    def parallel_root():
        global problem
        style.configure("Parallel.TButton", background="green", foreground="green")
        style.configure("Single.TButton", background="#abebc6",  foreground="black")
        style.configure("Flowshop.TButton", background="#abebc6",  foreground="black")
        problem_type.config(text="Parallel mod is selected.")
        problem = 'Parallel'
        all_root_forget()
        parallel_frame.pack()
        machines_num.grid(row=0, column=0, padx=5, pady=5)
        machine_count.grid(row=0, column=1, padx=5, pady=5)
        rule_var.set(models[3])
        rule_dropdown.config(values=models[-4:])
        all_root_get()
        method_frame.pack()
        method_label.grid(row=1, column=0, padx=5, pady=5)
        method_dropdown.grid(row=1, column=1, padx=5, pady=5)
        method_dropdown.config(values=methods)
        rule_start.pack(pady=5)
        
        
    def hide_or_show(method):
        if method == "Meta-Heuristic":
            thresh_frame.pack()
            thresh_label.grid(row=0, column=0, padx=5, pady=5)
            thresh_count.grid(row=0, column=1, padx=5, pady=5)
        else:
            thresh_frame.pack_forget()
            thresh_label.grid_forget()
            thresh_count.grid_forget()

    def flowshop_root():
        global problem
        style.configure("Parallel.TButton", background="#abebc6", foreground="black")
        style.configure("Single.TButton", background="#abebc6",  foreground="black")
        style.configure("Flowshop.TButton", background="green",  foreground="green")
        problem_type.config(text="Flowshop mod is selected.")
        problem = 'Flowshop'
        all_root_forget()
        parallel_frame.pack()
        
        machines_num.grid(row=0, column=0, padx=5, pady=5)
        machine_count.grid(row=0, column=1, padx=5, pady=5)
        method_frame.pack()
        method_label.grid(row=1, column=0, padx=5, pady=5)
        method_dropdown.grid(row=1, column=1, padx=5, pady=5)
        method_var.set(methods[2])
        method_dropdown.config(values=methods[-2:])
        rule_start.pack(pady=5)

    root = tk.Tk()
    root.title("Machine Scheduling Tool")
    root.geometry("600x700")
    root.configure(bg="#abebc6")
    style = ttk.Style(root)
    style.configure("TButton", font=("Arial", 10),background="#abebc6")
    tk.Label(root, text="Machine Scheduling Project - Group6", bg="#abebc6",font=("Arial", 18, "bold")).pack(pady=10)

    #Single ya da Paralel olup olmadığını kontrol ediyorum
    button_frame = tk.Frame(root,bg="#abebc6")
    button_frame.pack()
    single_button = ttk.Button(button_frame, text="Single Machine", command=single_root,style="Single.TButton")
    single_button.grid(row=0, column=0, padx=10, pady=5)
    parallel_button = ttk.Button(button_frame, text="Parallel Machine", command=parallel_root,style="Parallel.TButton")
    parallel_button.grid(row=0, column=1, padx=10, pady=5)
    flowshop_button = ttk.Button(button_frame, text="Flowshop", command=flowshop_root,style="Flowshop.TButton")
    flowshop_button.grid(row=0, column=2, padx=10, pady=5)
    problem_type = tk.Label(root, text="Please select a machine type problem.", bg="#abebc6",wraplength=400,font=("Arial", 14, "bold"))
    problem_type.pack(pady=10)

    #Excel dosyası yükletiyorum.
    excel_frame = tk.Frame(root,bg="#abebc6",bd=1, relief="solid")
    excel_frame.pack()
    ttk.Button(excel_frame, text="Load Excel File", command=open_file).grid(row=0, column=0, padx=10, pady=5)
    excel_label = tk.Label(excel_frame, bg="#abebc6",text="Please select an excel file.")
    excel_label.grid(row=0, column=1, padx=10, pady=5)

    #Modelleri Tanımlıyorum
    models = ["WSPT", "EDD","ERD","LPT","SPT","Wrap-Around","None"]
    selected_option = StringVar()
    selected_option.set(models[0])

    #Methodları tanımlıyorum
    methods = ["None","Meta-Heuristic","Local Search"]

    #Single ya da paralel machine için ayarlar yaptırıyorum. Eğer paralel olursa aşağıdakiler gözükecek.
    parallel_frame = tk.Frame(root,bg="#abebc6")
    machines_num = tk.Label(parallel_frame, bg="#abebc6",text="Enter Number of Machines:") 
    machine_count = tk.Entry(parallel_frame)

    #Her ikiside olursa aşağıdakiler gözükecek
    rule_frame = tk.Frame(root,bg="#abebc6")
    rule_label = tk.Label(rule_frame, bg="#abebc6",text="Select Dispatching Rule:")
    rule_var = tk.StringVar()
    rule_dropdown = ttk.Combobox(rule_frame, textvariable=rule_var, values=models, state="readonly")
    rule_start = ttk.Button(root, text="Start Calculation", command=apply_rule)

    #Method seçimi 
    method_frame = tk.Frame(root,bg="#abebc6")
    method_label = tk.Label(method_frame, bg="#abebc6",text="Select Method:")
    method_var = tk.StringVar()
    method_var.set(methods[0])
    method_dropdown = ttk.Combobox(method_frame, textvariable=method_var, values=methods, state="readonly")

    thresh_frame = tk.Frame(root,bg="#abebc6")
    thresh_label = tk.Label(thresh_frame,bg="#abebc6",text="Threshold")
    thresh_count = tk.Entry(thresh_frame)

    def on_selection_change(*args): 
        hide_or_show(method_var.get())

    method_var.trace_add("write", on_selection_change)

    #Tablo için frame
    table_frame = tk.Frame(root,bg="#abebc6")
    columns = ("Key", "Value")
    tree = ttk.Treeview(table_frame, columns=columns, show="headings",height=12)


    

    root.mainloop()

if __name__ == "__main__":
    main()
