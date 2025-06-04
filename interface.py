import os
import tkinter as tk
from tkinter import filedialog, messagebox
from compare import compare_bom
import tkinter.messagebox as messagebox
import os
from pandas import read_excel
from openpyxl.utils.exceptions import InvalidFileException


def check_bom_comparison(file_1: str, file_2: str, destination_file: str) -> bool:
    """
    Compares two BOM files and handles potential errors with appropriate message boxes.

    Args:
        file_1 (str): Path to first BOM file
        file_2 (str): Path to second BOM file
        destination_file (str): Path to destination file

    Returns:
        bool: True if comparison was successful and BOMs differ, False if identical or error occurred

    Raises:
        None: All errors are caught and handled with message boxes
    """
    try:
        # Verify file existence
        if not os.path.isfile(file_1):
            messagebox.showerror("File Error", f"First BOM file not found: {file_1}")
            return False
        if not os.path.isfile(file_2):
            messagebox.showerror("File Error", f"Second BOM file not found: {file_2}")
            return False

        # Verify output path is writable
        try:
            output_dir = os.path.dirname(destination_file)
            if output_dir and not os.access(output_dir, os.W_OK):
                messagebox.showerror("Permission Error", f"No write permission for output directory: {output_dir}")
                return False
        except OSError as e:
            messagebox.showerror("Path Error", f"Invalid output path: {str(e)}")
            return False

        # Attempt BOM comparison
        try:
            if not compare_bom(file_1, file_2, destination_file):
                messagebox.showinfo("Same BOMs", "BOM files are identical")
                return False
            return True

        except PermissionError as e:
            messagebox.showerror("Permission Error", f"Permission denied accessing files: {str(e)}")
            return False
        except ValueError as e:
            messagebox.showerror("Data Error", f"Invalid BOM structure: {str(e)}")
            return False
        except KeyError as e:
            messagebox.showerror("Data Error", f"Missing required column: {str(e)}")
            print(f"Missing required column error: {str(e)}")
            return False
        except InvalidFileException as e:
            messagebox.showerror("File Error", f"Invalid Excel file format: {str(e)}")
            return False
        except read_excel.exceptions.ReadError as e:
            messagebox.showerror("File Error", f"Error reading Excel file: {str(e)}")
            return False
        except Exception as e:
            messagebox.showerror("Error", f"Unexpected error during BOM comparison: {str(e)}")
            return False

    except Exception as e:
        messagebox.showerror("Critical Error", f"Failed to initialize comparison: {str(e)}")
        return False


def profile_startup():
    # Your existing code
    def browse_file_1():
        file_path = filedialog.askopenfilename(filetypes=[("Excel or Text files", "*.xlsx;*.xls;*.txt")])
        entry_file_1.delete(0, tk.END)
        entry_file_1.insert(0, file_path)

    def browse_file_2():
        file_path = filedialog.askopenfilename(filetypes=[("Excel or Text files", "*.xlsx;*.xls;*.txt")])
        entry_file_2.delete(0, tk.END)
        entry_file_2.insert(0, file_path)

    def browse_destination():
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx;*.xls")])
        entry_file.delete(0, tk.END)
        entry_file.insert(0, file_path)

    def compare_files(simple_bom_mode: bool):
        """
        Compares two BOM files from GUI inputs and handles potential errors with message boxes.

        Returns:
            bool: True if comparison was successful and BOMs differ, False if identical or error occurred
        """
        try:
            file_1 = entry_file_1.get()
            file_2 = entry_file_2.get()
            destination_file = entry_file.get()

            # Validate input fields
            if not file_1 or not file_2 or not destination_file:
                messagebox.showwarning("Warning", "All fields must be filled.")
                return False

            # Validate file extensions (allow .txt for input files)
            if simple_bom_mode:
                if not (file_1.endswith((".xlsx", ".xls", ".txt"))):
                    messagebox.showwarning("Warning", "File 1 must be a valid Excel or TXT file.")
                    return False
                if not (file_2.endswith((".xlsx", ".xls", ".txt"))):
                    messagebox.showwarning("Warning", "File 2 must be a valid Excel or TXT file.")
                    return False
            else:
                if not (file_1.endswith((".xlsx", ".xls"))):
                    messagebox.showwarning("Warning", "File 1 must be a valid Excel file for architecture comparison.")
                    return False
                if not (file_2.endswith((".xlsx", ".xls"))):
                    messagebox.showwarning("Warning", "File 2 must be a valid Excel file for architecture comparison.")
                    return False
            if not (destination_file.endswith(".xlsx") or destination_file.endswith(".xls")):
                messagebox.showwarning("Warning", "Destination File must be a valid Excel file.")
                return False

            # Verify file existence
            if not os.path.isfile(file_1):
                messagebox.showerror("File Error", f"First BOM file not found: {file_1}")
                return False
            if not os.path.isfile(file_2):
                messagebox.showerror("File Error", f"Second BOM file not found: {file_2}")
                return False

            # Verify output path is writable
            try:
                output_dir = os.path.dirname(destination_file)
                if output_dir and not os.access(output_dir, os.W_OK):
                    messagebox.showerror("Permission Error", f"No write permission for output directory: {output_dir}")
                    return False
            except OSError as e:
                messagebox.showerror("Path Error", f"Invalid output path: {str(e)}")
                return False

            # Attempt BOM comparison
            try:
                if not compare_bom(file_1, file_2, destination_file, simple_bom_mode=simple_bom_mode):
                    messagebox.showinfo("Same BOMs", "BOM files are identical")
                    return False
                return True

            except PermissionError as e:
                if "Error: Cannot save file" in str(e):
                    messagebox.showerror("Permission Error", str(e))
                else:
                    messagebox.showerror(
                        "Permission Error", f"Please close compared files before trying again \n\n {str(e)}")
                return False
            except ValueError as e:
                if "Sheet 'BOM' not found in the workbook" in str(e):
                    messagebox.showerror(
                        "Sheet Error", "'BOM' sheet not present in Excel file: Light BOM compare is only available for BOM extracted from Creo and formatted with 'BOM refresh' or BOM extracted from Oracle in a txt format")
                else:
                    messagebox.showerror("Data Error", f"Invalid BOM structure: \n \n {str(e)}")
                return False
            except KeyError as e:
                messagebox.showerror("Data Error", f"Missing required column: \n \n {str(e)}")
                return False
            except InvalidFileException as e:
                messagebox.showerror("File Error", f"Invalid Excel file format: \n \n {str(e)}")
                return False
            except read_excel.exceptions.ReadError as e:
                messagebox.showerror("File Error", f"Error reading Excel file: \n \n {str(e)}")
                return False
            except Exception as e:
                messagebox.showerror("Error", f"Unexpected error during BOM comparison: \n \n {str(e)}")
                return False

        except Exception as e:
            messagebox.showerror("Critical Error", f"Failed to initialize comparison: \n \n {str(e)}")
            return False

    def quit_app():
        root.destroy()

    def find_two_xlsx_files():
        """
        Looks in current directory for exactly two xlsx files.
        Returns tuple of file paths if found, empty tuple otherwise.
        """
        xlsx_files = [f for f in os.listdir() if f.endswith('.xlsx')]

        if len(xlsx_files) == 2:
            return (os.path.abspath(xlsx_files[0]), os.path.abspath(xlsx_files[1]))
        return ()

    root = tk.Tk()
    root.title("BOM Compare")

    label_file_1 = tk.Label(root, text="File 1:")
    label_file_1.grid(row=0, column=0, padx=10, pady=10)
    entry_file_1 = tk.Entry(root, width=50)
    entry_file_1.grid(row=0, column=1, padx=10, pady=10)
    label_file_2 = tk.Label(root, text="File 2:")
    label_file_2.grid(row=1, column=0, padx=10, pady=10)
    entry_file_2 = tk.Entry(root, width=50)
    entry_file_2.grid(row=1, column=1, padx=10, pady=10)

    def browse_files(target_entry):
        files = filedialog.askopenfilenames(filetypes=[("Excel or Text files", "*.xlsx;*.xls;*.txt")])
        print(files)
        if len(files) == 2:
            target_entry.delete(0, tk.END)
            target_entry.insert(0, files[0])
            other_entry = entry_file_2 if target_entry == entry_file_1 else entry_file_1
            other_entry.delete(0, tk.END)
            other_entry.insert(0, files[1])
            move_path_view()
        else:
            target_entry.delete(0, tk.END)
            target_entry.insert(0, files[0])
            move_path_view()

    button_browse_1 = tk.Button(root, text="Browse", command=lambda: browse_files(entry_file_1))
    button_browse_1.grid(row=0, column=2, padx=10, pady=10)

    button_browse_2 = tk.Button(root, text="Browse", command=lambda: browse_files(entry_file_2))
    button_browse_2.grid(row=1, column=2, padx=10, pady=10)

    def move_path_view():
        """
        Moves the view of the entry field to the end of the text.
        This is useful when the text is too long to fit in the entry field.
        """
        entry_file_1.xview_moveto(1)
        entry_file_2.xview_moveto(1)

    def swap_paths():
        path1 = entry_file_1.get()
        path2 = entry_file_2.get()
        entry_file_1.delete(0, tk.END)
        entry_file_2.delete(0, tk.END)
        entry_file_1.insert(0, path2)
        entry_file_2.insert(0, path1)
        move_path_view()  # Move view to end of text after swap

    button_swap = tk.Button(root, text="⇅", command=swap_paths)
    button_swap.grid(row=0, column=3, rowspan=2, padx=10, pady=10)

    # Auto-populate entries if exactly two xlsx files found
    xlsx_paths = find_two_xlsx_files()
    if xlsx_paths:
        entry_file_1.insert(0, xlsx_paths[0])
        entry_file_2.insert(0, xlsx_paths[1])

        move_path_view()  # Move view to end of text

    label_file = tk.Label(root, text="Destination File:")
    label_file.grid(row=2, column=0, padx=10, pady=10)
    default_path = f"C:\\Users\\{os.getlogin()}\\Downloads\\compare_result.xlsx"
    entry_file = tk.Entry(root, width=50)
    entry_file.insert(0, default_path)
    entry_file.grid(row=2, column=1, padx=10, pady=10)
    button_browse_file = tk.Button(root, text="Browse", command=browse_destination)
    button_browse_file.grid(row=2, column=2, padx=10, pady=10)

    # Create a frame for buttons to center them
    button_frame = tk.Frame(root)
    button_frame.grid(row=3, column=0, columnspan=3, pady=10)

    button_compare = tk.Button(button_frame, text="Compare Architecture",
                               command=lambda: compare_files(simple_bom_mode=False))
    button_compare.pack(side=tk.LEFT, padx=25)  # 25px padding on each side = 50px spacing

    button_simple_compare = tk.Button(button_frame, text="Simple BOM Compare",
                                      command=lambda: compare_files(simple_bom_mode=True))
    button_simple_compare.pack(side=tk.LEFT, padx=25)

    button_quit = tk.Button(root, text="Quit", command=quit_app)
    button_quit.grid(row=4, column=0, columnspan=3, padx=10, pady=10)

    root.mainloop()


if __name__ == "__main__":
    profile_startup()
