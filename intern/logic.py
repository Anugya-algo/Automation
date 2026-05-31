import datetime
import os
import re
import json
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter.scrolledtext import ScrolledText
import numpy as np
import zipfile
import tempfile
import io

from llm import ask_qwen

import pandas as pd
import matplotlib.pyplot as plt

# Optional local LLM support via llama-cpp-python (ggml models)
try:
    from llama_cpp import Llama
except Exception:
    Llama = None


class ExcelManager:

    LOG_FILE = "activity.log"
    QUERY_HINT = "e.g. count rows | top 10 | show columns Name, Status where Status = Open"

    def __init__(self, root):
        self.root = root
        self.root.title("Excel Analytics Tool")
        self.root.geometry("1450x850")
        self.root.minsize(1000, 650)

        self.df = None
        self.path = None
        self.log_path = os.path.join(os.path.abspath(os.getcwd()), self.LOG_FILE)
        self.llm_model_path = None
        self.report_history = []
        self.config_path = os.path.join(os.path.abspath(os.getcwd()), "config.json")
        
        self.load_config()
        self.create_widgets()
        self.log_activity("Application started.")


    def create_widgets(self):
        style = ttk.Style(self.root)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure("TButton", padding=6)
        style.configure("Treeview", rowheight=24)

        self.create_menu()

        toolbar = ttk.Frame(self.root, padding=(10, 10, 10, 0))
        toolbar.pack(fill=tk.X)

        button_frame = ttk.Frame(toolbar)
        button_frame.pack(side=tk.LEFT, anchor="w")

        ttk.Button(button_frame, text="Load Excel", command=self.load_excel).pack(side=tk.LEFT, padx=4)
        ttk.Button(button_frame, text="Add Row", command=self.add_row).pack(side=tk.LEFT, padx=4)
        ttk.Button(button_frame, text="Delete", command=self.delete_row).pack(side=tk.LEFT, padx=4)
        ttk.Button(button_frame, text="Save", command=self.save_excel).pack(side=tk.LEFT, padx=4)
        ttk.Button(button_frame, text="Generate Report", command=self.generate_report).pack(side=tk.LEFT, padx=4)

        query_frame = ttk.Frame(toolbar)
        query_frame.pack(side=tk.LEFT, padx=(20, 0), fill=tk.X, expand=True)

        ttk.Label(query_frame, text="Query").pack(side=tk.LEFT, padx=(0, 4))
        self.search_var = tk.StringVar()
        self.search_entry = ttk.Entry(query_frame, textvariable=self.search_var, width=50)
        self.search_entry.pack(side=tk.LEFT, padx=(0, 4), fill=tk.X, expand=True)
        self.search_entry.insert(0, self.QUERY_HINT)
        self.search_entry.bind("<FocusIn>", self.clear_search_placeholder)
        self.search_entry.bind("<FocusOut>", self.restore_search_placeholder)

        ttk.Button(query_frame, text="Search", command=self.search).pack(side=tk.LEFT, padx=(0, 4))
        ttk.Button(query_frame, text="Reset", command=self.reset_view).pack(side=tk.LEFT)

        data_frame = ttk.Frame(self.root, padding=(10, 10, 0, 0))
        data_frame.pack(fill=tk.BOTH, expand=True)

        self.tree = ttk.Treeview(data_frame, show="headings")
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.tree.bind("<Double-1>", self.edit_cell)

        tree_scroll_y = ttk.Scrollbar(data_frame, orient=tk.VERTICAL, command=self.tree.yview)
        tree_scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.configure(yscrollcommand=tree_scroll_y.set)

        tree_scroll_x = ttk.Scrollbar(self.root, orient=tk.HORIZONTAL, command=self.tree.xview)
        tree_scroll_x.pack(fill=tk.X, padx=10)
        self.tree.configure(xscrollcommand=tree_scroll_x.set)

        log_frame = ttk.LabelFrame(self.root, text="Activity Log", padding=(10, 10, 10, 10))
        log_frame.pack(fill=tk.BOTH, padx=10, pady=(10, 10))

        self.log_text = ScrolledText(log_frame, height=8, state="disabled", wrap=tk.WORD)
        self.log_text.pack(fill=tk.BOTH, expand=True)

        # Assistant panel
        self.create_assistant_panel()

        self.status_var = tk.StringVar(value="Ready")
        status_bar = ttk.Label(self.root, textvariable=self.status_var, relief=tk.SUNKEN, anchor="w")
        status_bar.pack(fill=tk.X, side=tk.BOTTOM)

    def create_menu(self):
        menu_bar = tk.Menu(self.root)
        file_menu = tk.Menu(menu_bar, tearoff=False)
        file_menu.add_command(label="Load Excel...", command=self.load_excel)
        file_menu.add_command(label="Save", command=self.save_excel)
        file_menu.add_command(label="Reset view", command=self.reset_view)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.root.quit)
        menu_bar.add_cascade(label="File", menu=file_menu)

        help_menu = tk.Menu(menu_bar, tearoff=False)
        help_menu.add_command(label="Query examples", command=self.show_query_examples)
        menu_bar.add_cascade(label="Help", menu=help_menu)

        self.root.config(menu=menu_bar)

    def create_assistant_panel(self):
        frame = ttk.LabelFrame(self.root, text="Assistant", padding=(10, 6))
        frame.pack(fill=tk.BOTH, padx=10, pady=(0, 10))

        self.assistant_text = ScrolledText(frame, height=6, state="disabled", wrap=tk.WORD)
        self.assistant_text.pack(fill=tk.BOTH, expand=True, pady=(6, 6))

        entry_row = ttk.Frame(frame)
        entry_row.pack(fill=tk.X)
        self.assistant_var = tk.StringVar()
        self.assistant_entry = ttk.Entry(entry_row, textvariable=self.assistant_var)
        self.assistant_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6))
        self.assistant_entry.bind("<Return>", lambda e: self.send_assistant())
        ttk.Button(entry_row, text="Ask", command=self.send_assistant).pack(side=tk.RIGHT)

    def load_config(self):
        """Auto-load LLM model path from config.json or environment variable."""
        try:
            if os.path.exists(self.config_path):
                with open(self.config_path, "r", encoding="utf-8") as f:
                    cfg = json.load(f)
                    model_path = cfg.get("llm_model_path")
                    if model_path and os.path.exists(model_path):
                        self.llm_model_path = model_path
                        return
        except Exception:
            pass
        
        # Try environment variable
        model_path = os.environ.get("LLM_MODEL_PATH")
        if model_path and os.path.exists(model_path):
            self.llm_model_path = model_path

    def save_config(self):
        """Save LLM model path and other settings to config.json."""
        cfg = {
            "llm_model_path": self.llm_model_path,
        }
        try:
            with open(self.config_path, "w", encoding="utf-8") as f:
                json.dump(cfg, f, indent=2)
        except Exception:
            pass

    def set_llm_model(self):
        path = filedialog.askopenfilename(title="Select LLM model file", filetypes=[("Model files", "*.bin;*.pt;*.ggml*;*.bin")])
        if not path:
            return
        self.llm_model_path = path
        self.save_config()
        self.log_activity(f"LLM model set to {os.path.basename(path)}")

    def send_assistant(self):
        prompt = self.assistant_var.get().strip()
        if not prompt:
            return
        self.assistant_var.set("")
        self.append_assistant_message("User", prompt)
        try:
            resp = self.ask_local_llm(prompt)
            self.append_assistant_message("Assistant", resp)
            self.log_activity("Assistant answered a question.")
        except Exception as e:
            messagebox.showerror("Assistant Error", str(e))
            self.log_activity(f"Assistant error: {e}")

    def append_assistant_message(self, who, text):
        self.assistant_text.config(state="normal")
        self.assistant_text.insert(tk.END, f"{who}: {text}\n\n")
        self.assistant_text.see(tk.END)
        self.assistant_text.config(state="disabled")

    def ask_local_llm(self, user_prompt, max_tokens=256):
        if Llama is None:
            raise RuntimeError("Local LLM support not available. Install 'llama-cpp-python' and retry.")
        if not self.llm_model_path:
            raise RuntimeError("LLM model not configured. Set LLM_MODEL_PATH environment variable or add llm_model_path to config.json.")

        # Build a small prompt with dataset context (columns + sample rows)
        ctx_parts = ["You are a helpful assistant for analyzing tabular data."]
        try:
            cols = list(self.df.columns)
            ctx_parts.append("Columns: " + ", ".join(cols))
            sample = self.df.head(5).to_dict(orient="records")
            ctx_parts.append("Sample rows: ")
            for r in sample:
                ctx_parts.append(str(r))
        except Exception:
            pass
        ctx_parts.append("User question: ")
        ctx_parts.append(user_prompt)
        prompt = "\n".join(ctx_parts)

        # Call local Llama model (silently in background)
        llm = Llama(model_path=self.llm_model_path)
        try:
            out = llm.create(prompt=prompt, max_tokens=max_tokens)
            # llama-cpp-python returns choices with 'text' or 'content'
            text = None
            if isinstance(out, dict):
                choices = out.get("choices") or []
                if choices:
                    text = choices[0].get("text") or choices[0].get("content")
            if text is None:
                text = str(out)
            return text.strip()
        finally:
            try:
                llm.__del__()
            except Exception:
                pass

    def show_query_examples(self):
        messagebox.showinfo(
            "Query examples",
            "Examples:\n"
            "- count rows\n"
            "- top 10\n"
            "- show columns Name, Status where Status = Open\n"
            "- filter Assigned To by Alice\n"
            "- sum of Hours\n"
            "- Revenue > 1000"
        )

    def clear_search_placeholder(self, event):
        if self.search_entry.get() == self.QUERY_HINT:
            self.search_entry.delete(0, tk.END)

    def restore_search_placeholder(self, event):
        if not self.search_entry.get().strip():
            self.search_entry.insert(0, self.QUERY_HINT)

    def log_activity(self, message):
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        text = f"[{timestamp}] {message}"
        try:
            with open(self.log_path, "a", encoding="utf-8") as handle:
                handle.write(text + "\n")
        except OSError:
            pass

        self.log_text.config(state="normal")
        self.log_text.insert(tk.END, text + "\n")
        self.log_text.see(tk.END)
        self.log_text.config(state="disabled")
        self.set_status(message)

    def set_status(self, message):
        self.status_var.set(message)

    def load_excel(self):
        path = filedialog.askopenfilename(
            filetypes=[
                ("Excel Files", "*.xlsx *.xls"),
                ("CSV Files", "*.csv"),
                ("All Files", "*.*"),
            ]
        )
        if not path:
            return

        try:
            raw = pd.read_csv(path, header=None) if path.lower().endswith(".csv") else pd.read_excel(path, header=None)
        except Exception as error:
            messagebox.showerror("Load Error", str(error))
            self.log_activity(f"Failed to load file: {error}")
            return

        header_row = 0
        keywords = ["project name", "task name", "assigned to"]

        for i, row in raw.iterrows():
            values = [str(x).lower() if pd.notna(x) else "" for x in row]
            score = sum(1 for keyword in keywords if any(keyword in value for value in values))
            if score >= 2:
                header_row = i
                break

        try:
            self.df = pd.read_csv(path, header=header_row) if path.lower().endswith(".csv") else pd.read_excel(path, header=header_row)
        except Exception as error:
            messagebox.showerror("Load Error", str(error))
            self.log_activity(f"Failed to parse file: {error}")
            return

        self.path = path
        self.df = self.df.dropna(how="all")
        self.df = self.df.loc[:, ~self.df.columns.astype(str).str.contains("unnamed", case=False)]
        self.df.columns = [str(c).strip() for c in self.df.columns]
        self.df = self.df.reset_index(drop=True)

        if self.df.empty:
            messagebox.showwarning("Empty Data", "The loaded file contains no usable rows.")

        self.show_df(self.df)
        self.log_activity(f"Loaded file '{os.path.basename(path)}' with {len(self.df)} rows.")

    def show_df(self, data):
        if data is None:
            return

        self.tree.delete(*self.tree.get_children())
        self.tree.config(columns=list(data.columns))
        self.tree["show"] = "headings"

        for column_name in data.columns:
            self.tree.heading(column_name, text=column_name)
            self.tree.column(column_name, width=160, minwidth=80, anchor=tk.W)

        if data.empty:
            return

        for row in data.itertuples(index=False, name=None):
            self.tree.insert("", tk.END, values=row)

    def add_row(self):
        if self.df is None:
            messagebox.showwarning("No Data", "Load a spreadsheet before adding rows.")
            return

        popup = tk.Toplevel(self.root)
        popup.title("Add Row")
        popup.grab_set()

        entries = {}
        for index, column_name in enumerate(self.df.columns):
            ttk.Label(popup, text=column_name).grid(row=index, column=0, sticky=tk.W, padx=6, pady=4)
            entry = ttk.Entry(popup, width=40)
            entry.grid(row=index, column=1, padx=6, pady=4)
            entries[column_name] = entry

        def save():
            new_row = [entries[column].get() for column in self.df.columns]
            self.df.loc[len(self.df)] = new_row
            self.show_df(self.df)
            self.log_activity("Added a new row.")
            popup.destroy()

        ttk.Button(popup, text="Insert", command=save).grid(row=len(self.df.columns), columnspan=2, pady=10)

    def delete_row(self):
        if self.df is None:
            messagebox.showwarning("No Data", "Load a spreadsheet before deleting rows.")
            return

        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("No Selection", "Please select a row to delete.")
            return

        idx = self.tree.index(selected[0])
        self.df = self.df.drop(idx).reset_index(drop=True)
        self.show_df(self.df)
        self.log_activity(f"Deleted row {idx + 1}.")

    def edit_cell(self, event):
        if self.df is None:
            return

        selected = self.tree.selection()
        if not selected:
            return

        item = selected[0]
        row_index = self.tree.index(item)
        col_index = int(self.tree.identify_column(event.x)[1:]) - 1
        current_values = self.tree.item(item, "values")

        popup = tk.Toplevel(self.root)
        popup.title("Edit Cell")
        popup.grab_set()

        entry = ttk.Entry(popup, width=40)
        entry.insert(0, current_values[col_index])
        entry.pack(padx=10, pady=10)

        def save():
            self.df.iat[row_index, col_index] = entry.get()
            self.show_df(self.df)
            self.log_activity(f"Updated row {row_index + 1}, column {self.df.columns[col_index]}.")
            popup.destroy()

        ttk.Button(popup, text="Update", command=save).pack(pady=(0, 10))

    def search(self):
        if self.df is None:
            messagebox.showwarning("No Data", "Load a spreadsheet before searching.")
            return

        query = self.search_var.get().strip()
        if not query or query == self.QUERY_HINT:
            self.reset_view()
            return

        from llm import ask_qwen

        normalized_query = ask_qwen(
            query,
            list(
                self.df.columns
        )
        )

        print(
            "LLM Output:",
            normalized_query
        )

        try:
            result = self.run_query(normalized_query)
            if isinstance(result, pd.DataFrame):
                self.show_df(result)
                self.log_activity(f"Search executed: {query}")
            else:
                messagebox.showinfo("Result", str(result))
                self.log_activity(f"Search executed: {query}")
        except Exception as error:
            messagebox.showerror("Query Error", str(error))
            self.log_activity(f"Search failed: {query}")

    def run_query(self, query):
        if not query or not query.strip():
            raise ValueError("Please enter a query.")

        q = self.normalize_query(query)
        q_lower = q.lower()

        if q_lower == "count rows":
            return f"Rows: {len(self.df)}"

        top_match = re.match(r"top\s+(\d+)", q_lower)
        if top_match:
            count = int(top_match.group(1))
            return self.df.head(count)

        columns_where = re.match(r"(?:show|display)\s+(?:columns?\s+)?(.+?)\s+where\s+(.+)", q, re.I)
        if columns_where:
            columns = self.resolve_columns(columns_where.group(1))
            filtered = self.apply_condition(self.df, columns_where.group(2))
            return filtered.loc[:, columns]

        columns_only = re.match(r"(?:show|display)\s+(?:columns?\s+)?(.+)$", q, re.I)
        if columns_only:
            columns = self.resolve_columns(columns_only.group(1))
            return self.df.loc[:, columns]

        filter_match = re.match(r"(?:filter|show)\s+(.+?)\s+(?:by|with|where)\s+(.+)", q, re.I)
        if filter_match:
            column_hint = filter_match.group(1).strip()
            value = filter_match.group(2).strip()
            column_name = self.find_column(column_hint)
            if column_name is None:
                raise ValueError(f"Column not found: {column_hint}")
            return self.apply_condition(self.df, f"{column_name} = {value}")

        stat_match = re.match(r"(sum|average|avg|min|max)\s+of\s+(.+)", q_lower)
        if stat_match:
            op = stat_match.group(1)
            column_hint = stat_match.group(2).strip()
            column_name = self.find_column(column_hint)
            if column_name is None:
                raise ValueError(f"Column not found: {column_hint}")
            series = pd.to_numeric(
                self.df[column_name].astype(str).str.replace("%", "", regex=False).str.extract(r"([-+]?\d*\.?\d+)")[0],
                errors="coerce",
            )
            if series.empty:
                raise ValueError(f"No numeric values found in column '{column_name}'.")
            if op in ("average", "avg"):
                return f"Average {column_name}: {series.mean():.4g}"
            if op == "sum":
                return f"Sum of {column_name}: {series.sum():.4g}"
            if op == "min":
                return f"Min of {column_name}: {series.min():.4g}"
            if op == "max":
                return f"Max of {column_name}: {series.max():.4g}"

        if any(symbol in q for symbol in [">=", "<=", "!=", "==", ">", "<", "="]):
            return self.apply_condition(self.df, q)

        mask = self.df.astype(str).apply(lambda x: x.str.contains(q, case=False, na=False))
        filtered = self.df[mask.any(axis=1)]
        if filtered.empty:
            raise ValueError("No rows match the query.")
        return filtered

    def resolve_columns(self, raw_columns):
        raw_columns = raw_columns.strip()
        if not raw_columns:
            return list(self.df.columns)

        column_names = [c.strip() for c in raw_columns.split(",") if c.strip()]
        if not column_names:
            return list(self.df.columns)

        resolved = []
        for name in column_names:
            column = self.find_column(name)
            if column is None:
                raise ValueError(f"Column not found: {name}")
            resolved.append(column)
        return resolved

    def find_column(self, name_hint):
        name_hint = str(name_hint).strip().lower()
        if not name_hint:
            return None
        for column in self.df.columns:
            if column.lower() == name_hint:
                return column
        for column in self.df.columns:
            if name_hint in column.lower():
                return column
        return None

    def normalize_query(self, query):
        q = str(query).strip()
        q = re.sub(r"[.?!]+$", "", q)
        replacements = [
            (r"\b(greater than or equal to|at least|>=)\b", ">="),
            (r"\b(less than or equal to|at most|<=)\b", "<="),
            (r"\b(more than|greater than)\b", ">"),
            (r"\b(less than|smaller than)\b", "<"),
            (r"\b(not equal to|does not equal|is not|!=)\b", "!="),
            (r"\b(equal to|equals|is equal to|is)\b", "="),
        ]
        for pattern, replacement in replacements:
            q = re.sub(pattern, replacement, q, flags=re.I)
        q = re.sub(r"\s+", " ", q).strip()
        return q

    def parse_numeric_value(self, raw_value):
        raw_value = str(raw_value).strip().replace("%", "").replace(",", "")
        match = re.search(r"([-+]?\d*\.?\d+)", raw_value)
        if not match:
            raise ValueError(f"Could not parse numeric value from '{raw_value}'.")
        return float(match.group(1))

    def apply_condition(self, data, condition):
        if not condition or not str(condition).strip():
            raise ValueError("Please provide a condition.")

        condition = self.normalize_query(condition)

        condition = str(condition).strip()
        operator_match = re.search(r"(>=|<=|!=|==|=|>|<)", condition)
        if not operator_match:
            raise ValueError("Condition must include one of: >=, <=, !=, ==, =, >, <.")

        raw_column, raw_operator, raw_value = re.split(r"(>=|<=|!=|==|=|>|<)", condition, maxsplit=1)
        column_name = self.find_column(raw_column.strip())
        if column_name is None:
            raise ValueError(f"Column not found: '{raw_column.strip()}'.")

        if raw_operator in ("=", "==", "!="):
            value_text = raw_value.strip().strip('"\'')
            series = data[column_name].astype(str).str.strip()
            if raw_operator == "!=":
                return data[series.str.lower() != value_text.lower()]
            return data[series.str.lower() == value_text.lower()]

        numeric_series = pd.to_numeric(
            data[column_name].astype(str).str.replace("%", "", regex=False).str.extract(r"([-+]?\d*\.?\d+)")[0],
            errors="coerce",
        )
        numeric_value = self.parse_numeric_value(raw_value)

        if raw_operator == ">":
            mask = numeric_series > numeric_value
        elif raw_operator == "<":
            mask = numeric_series < numeric_value
        elif raw_operator == ">=":
            mask = numeric_series >= numeric_value
        elif raw_operator == "<=":
            mask = numeric_series <= numeric_value
        else:
            raise ValueError(f"Unsupported operator: {raw_operator}")

        return data[mask]

    def reset_view(self):
        if self.df is not None:
            self.show_df(self.df)
            self.log_activity("Reset to full dataset view.")

    def generate_report(self):
        if self.df is None:
            messagebox.showwarning("No Data", "Load a spreadsheet before generating a report.")
            return

        popup = tk.Toplevel(self.root)
        popup.title("Generate Report")
        popup.geometry("520x350")
        popup.grab_set()

        ttk.Label(popup, text="Columns (comma separated)").pack(anchor=tk.W, padx=10, pady=(10, 4))
        cols_entry = ttk.Entry(popup, width=60)
        cols_entry.pack(padx=10, pady=4)

        ttk.Label(popup, text="Optional condition").pack(anchor=tk.W, padx=10, pady=(10, 4))
        cond_entry = ttk.Entry(popup, width=60)
        cond_entry.pack(padx=10, pady=4)

        plot_type = tk.StringVar(value="bar")
        ttk.Label(popup, text="Plot type").pack(anchor=tk.W, padx=10, pady=(10, 4))
        ttk.Combobox(
            popup,
            textvariable=plot_type,
            values=["line", "bar", "hist", "box"],
            state="readonly",
        ).pack(padx=10, pady=4)

        ttk.Label(popup, text="Report insights").pack(anchor=tk.W, padx=10, pady=(10, 4))
        insights_box = ScrolledText(popup, height=6, state="disabled", wrap=tk.WORD)
        insights_box.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))

        def create():
            try:
                data = self.df.copy()
                condition = cond_entry.get().strip()
                if condition:
                    data = self.apply_condition(data, condition)

                selected_columns = self.resolve_columns(cols_entry.get())
                report = data.loc[:, selected_columns]
                if report.empty:
                    raise ValueError("No results match the selected columns and condition.")

                self.show_df(report)
                self.log_activity(f"Generated report with columns {', '.join(selected_columns)}.")

                insights_text = self.get_report_insights(report, selected_columns, condition)
                insights_box.config(state="normal")
                insights_box.delete("1.0", tk.END)
                insights_box.insert(tk.END, insights_text)
                insights_box.config(state="disabled")

                # add to history
                self.add_report_to_history(report, insights_text)

                numeric = self.make_numeric(report)
                if numeric.empty:
                    messagebox.showinfo("Report", "Report created but no numeric columns were available for plotting.")
                    return

                fig = plt.figure(figsize=(9, 5))
                numeric.plot(kind=plot_type.get() or "bar")
                plt.tight_layout()
                plt.show()

                # Controls: Save CSV, Save insights, Save plot, Heatmap, Group-by, Export bundle, History
                ctrl_frame = ttk.Frame(popup)
                ctrl_frame.pack(fill=tk.X, padx=10, pady=(6, 10))

                def _save_csv():
                    self.save_report_csv(report)

                def _save_insights():
                    self.save_insights_text(insights_text)

                def _save_plot():
                    self.save_plot_from_numeric(numeric, plot_type.get())

                def _heatmap():
                    self.show_correlation_heatmap(numeric)

                def _groupby():
                    self.open_groupby_dialog(report)

                def _export_bundle():
                    self.export_report_bundle(report, insights_text, numeric)

                def _history():
                    self.show_report_history()

                ttk.Button(ctrl_frame, text="Save CSV", command=_save_csv).pack(side=tk.LEFT, padx=4)
                ttk.Button(ctrl_frame, text="Save insights", command=_save_insights).pack(side=tk.LEFT, padx=4)
                ttk.Button(ctrl_frame, text="Save plot", command=_save_plot).pack(side=tk.LEFT, padx=4)
                ttk.Button(ctrl_frame, text="Heatmap", command=_heatmap).pack(side=tk.LEFT, padx=4)
                ttk.Button(ctrl_frame, text="Group-by", command=_groupby).pack(side=tk.LEFT, padx=4)
                ttk.Button(ctrl_frame, text="Export bundle", command=_export_bundle).pack(side=tk.LEFT, padx=4)
                ttk.Button(ctrl_frame, text="History", command=_history).pack(side=tk.LEFT, padx=4)
            except Exception as error:
                messagebox.showerror("Report Error", str(error))
                self.log_activity(f"Report generation failed: {error}")

        ttk.Button(popup, text="Generate", command=create).pack(pady=(10, 10))

    def make_numeric(self, data):
        numeric = data.copy()
        for column in numeric.columns:
            numeric[column] = pd.to_numeric(
                numeric[column].astype(str).str.replace("%", "", regex=False).str.extract(r"([-+]?\d*\.?\d+)")[0],
                errors="coerce",
            )
        return numeric.dropna(axis=1, how="all")

    def get_report_insights(self, report, selected_columns, condition):
        rows = len(report)
        columns = list(report.columns)
        insights = [f"This report contains {rows} row(s) and {len(columns)} selected column(s)."]
        if condition:
            insights.append(f"Condition applied: {condition}")
        if rows == 0:
            insights.append("No data remains after applying the filters.")
            return "\n".join(insights)

        numeric = self.make_numeric(report)
        if numeric.empty:
            insights.append("No numeric columns are available for summary.")
            return "\n".join(insights)

        insights.append("Numeric summary:")
        summary_columns = numeric.columns[:4]
        for col in summary_columns:
            col_data = numeric[col].dropna()
            if col_data.empty:
                continue
            insights.append(
                f"- {col}: min={col_data.min():.4g}, max={col_data.max():.4g}, mean={col_data.mean():.4g}, count={len(col_data)}"
            )

        if len(numeric.columns) > 1:
            insights.append("Multiple numeric columns are present, so the report can be used to compare trends and outliers.")
        else:
            insights.append("This report is useful for checking the distribution and totals of the selected numeric column.")

        return "\n".join(insights)

    def add_report_to_history(self, report, insights_text):
        entry = {
            "timestamp": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "report": report.copy(),
            "insights": insights_text,
        }
        self.report_history.insert(0, entry)
        self.report_history = self.report_history[:20]

    def save_report_csv(self, report):
        path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV", "*.csv")])
        if not path:
            return
        try:
            report.to_csv(path, index=False)
            messagebox.showinfo("Saved", f"Report saved to {path}")
            self.log_activity(f"Report CSV saved: {os.path.basename(path)}")
        except Exception as e:
            messagebox.showerror("Save Error", str(e))
            self.log_activity(f"Save report failed: {e}")

    def save_insights_text(self, text):
        path = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Text", "*.txt")])
        if not path:
            return
        try:
            with open(path, "w", encoding="utf-8") as f:
                f.write(text)
            messagebox.showinfo("Saved", f"Insights saved to {path}")
            self.log_activity(f"Insights saved: {os.path.basename(path)}")
        except Exception as e:
            messagebox.showerror("Save Error", str(e))
            self.log_activity(f"Save insights failed: {e}")

    def save_plot_from_numeric(self, numeric, plot_type):
        path = filedialog.asksaveasfilename(defaultextension=".png", filetypes=[("PNG", "*.png"), ("JPEG", "*.jpg")])
        if not path:
            return
        try:
            fig = plt.figure(figsize=(9, 5))
            numeric.plot(kind=plot_type or "bar")
            plt.tight_layout()
            fig.savefig(path)
            plt.close(fig)
            messagebox.showinfo("Saved", f"Plot saved to {path}")
            self.log_activity(f"Plot saved: {os.path.basename(path)}")
        except Exception as e:
            messagebox.showerror("Save Error", str(e))
            self.log_activity(f"Save plot failed: {e}")

    def show_correlation_heatmap(self, numeric):
        if numeric.empty:
            messagebox.showinfo("Heatmap", "No numeric data to show heatmap.")
            return
        corr = numeric.corr()
        fig, ax = plt.subplots(figsize=(6, 5))
        cax = ax.imshow(corr.values, cmap="RdYlBu", vmin=-1, vmax=1)
        ax.set_xticks(np.arange(len(corr.columns)))
        ax.set_yticks(np.arange(len(corr.columns)))
        ax.set_xticklabels(corr.columns, rotation=45, ha="right")
        ax.set_yticklabels(corr.columns)
        fig.colorbar(cax, ax=ax, fraction=0.046, pad=0.04)
        plt.title("Correlation heatmap")
        plt.tight_layout()
        plt.show()

        notes = []
        for i, a in enumerate(corr.columns):
            for j, b in enumerate(corr.columns):
                if i >= j:
                    continue
                val = corr.iat[i, j]
                if abs(val) >= 0.8:
                    notes.append(f"High correlation ({val:.2f}) between {a} and {b}")
        if notes:
            messagebox.showinfo("Heatmap insights", "\n".join(notes))

    def open_groupby_dialog(self, report):
        if report is None or report.empty:
            messagebox.showwarning("Group-by", "No report data to group.")
            return
        popup = tk.Toplevel(self.root)
        popup.title("Group By")
        popup.geometry("420x320")
        popup.grab_set()

        ttk.Label(popup, text="Group by column").pack(anchor=tk.W, padx=10, pady=(10, 4))
        group_col = tk.StringVar()
        ttk.Combobox(popup, textvariable=group_col, values=list(report.columns), state="readonly").pack(padx=10, pady=4)

        ttk.Label(popup, text="Aggregation").pack(anchor=tk.W, padx=10, pady=(10, 4))
        agg_var = tk.StringVar(value="sum")
        ttk.Combobox(popup, textvariable=agg_var, values=["sum", "mean", "count"], state="readonly").pack(padx=10, pady=4)

        ttk.Label(popup, text="Columns to aggregate (comma separated)").pack(anchor=tk.W, padx=10, pady=(10, 4))
        cols_entry = ttk.Entry(popup, width=50)
        cols_entry.pack(padx=10, pady=4)

        def do_groupby():
            g = group_col.get().strip()
            if not g:
                messagebox.showwarning("Group-by", "Select a group-by column.")
                return
            aggs = [c.strip() for c in cols_entry.get().split(",") if c.strip()]
            if not aggs:
                aggs = list(self.make_numeric(report).columns)
            try:
                if agg_var.get() == "count":
                    res = report.groupby(g).size().reset_index(name="count")
                else:
                    res = getattr(report.groupby(g)[aggs], agg_var.get())().reset_index()
                self.show_df(res)
                self.log_activity(f"Group-by {g} with {agg_var.get()} on {', '.join(aggs)}")
                self.add_report_to_history(res, f"Group-by {g} agg {agg_var.get()} on {', '.join(aggs)}")
                popup.destroy()
            except Exception as e:
                messagebox.showerror("Group-by Error", str(e))

        ttk.Button(popup, text="Apply", command=do_groupby).pack(pady=10)

    def export_report_bundle(self, report, insights_text, numeric):
        path = filedialog.asksaveasfilename(defaultextension=".zip", filetypes=[("Zip", "*.zip")])
        if not path:
            return
        try:
            with tempfile.TemporaryDirectory() as td:
                csv_path = os.path.join(td, "report.csv")
                report.to_csv(csv_path, index=False)
                insights_path = os.path.join(td, "insights.txt")
                with open(insights_path, "w", encoding="utf-8") as f:
                    f.write(insights_text)
                plot_path = os.path.join(td, "plot.png")
                if not numeric.empty:
                    fig = plt.figure(figsize=(9, 5))
                    numeric.plot(kind="bar")
                    plt.tight_layout()
                    fig.savefig(plot_path)
                    plt.close(fig)
                with zipfile.ZipFile(path, "w") as zf:
                    zf.write(csv_path, arcname="report.csv")
                    zf.write(insights_path, arcname="insights.txt")
                    if os.path.exists(plot_path):
                        zf.write(plot_path, arcname="plot.png")
            messagebox.showinfo("Exported", f"Bundle exported to {path}")
            self.log_activity(f"Exported bundle: {os.path.basename(path)}")
        except Exception as e:
            messagebox.showerror("Export Error", str(e))
            self.log_activity(f"Export bundle failed: {e}")

    def show_report_history(self):
        popup = tk.Toplevel(self.root)
        popup.title("Report History")
        popup.geometry("600x400")
        popup.grab_set()

        lb = tk.Listbox(popup)
        lb.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(10, 0), pady=10)
        for i, e in enumerate(self.report_history):
            lb.insert(tk.END, f"{i+1}: {e['timestamp']} - {len(e['report'])} rows")

        detail = ScrolledText(popup, state="disabled")
        detail.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=(6, 10), pady=10)

        def on_select(evt):
            sel = lb.curselection()
            if not sel:
                return
            idx = sel[0]
            e = self.report_history[idx]
            detail.config(state="normal")
            detail.delete("1.0", tk.END)
            detail.insert(tk.END, f"Timestamp: {e['timestamp']}\n\n")
            detail.insert(tk.END, e['insights'])
            detail.config(state="disabled")

        lb.bind("<<ListboxSelect>>", on_select)

        def _reload():
            sel = lb.curselection()
            if not sel:
                return
            idx = sel[0]
            e = self.report_history[idx]
            self.show_df(e['report'])
            popup.destroy()

        def _export_selected():
            sel = lb.curselection()
            if not sel:
                return
            idx = sel[0]
            e = self.report_history[idx]
            self.export_report_bundle(e['report'], e['insights'], self.make_numeric(e['report']))

        btn_frame = ttk.Frame(popup)
        btn_frame.pack(fill=tk.X, padx=10, pady=(0, 10))
        ttk.Button(btn_frame, text="Reload", command=_reload).pack(side=tk.LEFT, padx=4)
        ttk.Button(btn_frame, text="Export selected", command=_export_selected).pack(side=tk.LEFT, padx=4)

    def save_excel(self):
        if self.df is None:
            messagebox.showwarning("No Data", "Load a spreadsheet before saving.")
            return

        if not self.path:
            self.path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel Files", "*.xlsx"), ("CSV Files", "*.csv"), ("All Files", "*.*")],
            )
            if not self.path:
                return

        try:
            if self.path.lower().endswith(".csv"):
                self.df.to_csv(self.path, index=False)
            else:
                self.df.to_excel(self.path, index=False)
            messagebox.showinfo("Saved", f"Saved to {self.path}")
            self.log_activity(f"Saved data to '{os.path.basename(self.path)}'.")
        except Exception as error:
            messagebox.showerror("Save Error", str(error))
            self.log_activity(f"Save failed: {error}")


if __name__ == "__main__":
    root = tk.Tk()
    ExcelManager(root)
    root.mainloop()
