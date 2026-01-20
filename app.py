"""
Hymn PDF Combiner - Complete application.
"""

import tkinter as tk
from tkinter import messagebox
from pathlib import Path
import threading
import re


class ModernStyle:
    """Modern styling configuration for GUI."""
    
    # Color scheme
    PRIMARY = "#27AE60"     # Green
    SECONDARY = "#2ECC71"   # Lighter green
    SUCCESS = "#10B981"       # Professional green
    WARNING = "#F59E0B"       # Professional orange
    ERROR = "#E74C3C"       # Red for errors
    BACKGROUND = "#ECF0F1"  # Light gray background
    SURFACE = "#FFFFFF"     # White surface
    TEXT = "#2C3E50"        # Dark gray text
    TEXT_LIGHT = "#7F8C8D"  # Medium gray text
    BORDER = "#BDC3C7"      # Light gray border
    
    # Fonts
    TITLE_FONT = ("Segoe UI", 20, "bold")
    HEADER_FONT = ("Segoe UI", 12, "bold")
    LABEL_FONT = ("Segoe UI", 11)
    BUTTON_FONT = ("Segoe UI", 10, "bold")
    STATUS_FONT = ("Consolas", 9)


def parse_hymn_numbers(raw: str) -> list[int]:
    """Parse hymn numbers from user input."""
    raw = (raw or "").strip()
    if not raw:
        return []

    tokens = re.split(r"[\s,;]+", raw)
    nums: list[int] = []
    for tok in tokens:
        if not tok:
            continue

        m = re.fullmatch(r"(\d{1,4})-(\d{1,4})", tok)
        if m:
            start = int(m.group(1))
            end = int(m.group(2))
            step = 1 if end >= start else -1
            nums.extend(list(range(start, end + step, step)))
            continue

        if re.fullmatch(r"\d{1,4}", tok):
            nums.append(int(tok))

    return nums


def pdf_paths_for_hymn(pdf_dir: Path, hymn_number: int) -> list[Path]:
    """Find all PDF files for a given hymn number."""
    if not pdf_dir.exists():
        return []

    pattern = re.compile(rf"^{hymn_number}(?:\.|_).+\.pdf$", re.IGNORECASE)
    exact = pdf_dir / f"{hymn_number}.pdf"

    matches: list[Path] = []
    if exact.exists():
        matches.append(exact)

    for p in pdf_dir.glob(f"{hymn_number}_*.pdf"):
        if pattern.match(p.name):
            matches.append(p)

    def sort_key(p: Path):
        m = re.search(r"_PAGE(\d+)", p.stem, flags=re.IGNORECASE)
        page = int(m.group(1)) if m else 0
        return (page, p.name.lower())

    return sorted(set(matches), key=sort_key)


def merge_pdfs(input_paths: list[Path], output_path: Path) -> None:
    """Merge multiple PDF files into one."""
    try:
        from pypdf import PdfMerger
    except ImportError:
        try:
            from PyPDF2 import PdfMerger
        except ImportError:
            raise ImportError("Install 'pypdf' (recommended) or 'PyPDF2' to merge PDFs")

    output_path.parent.mkdir(parents=True, exist_ok=True)

    merger = PdfMerger()
    try:
        for p in input_paths:
            merger.append(str(p))
        with output_path.open("wb") as f:
            merger.write(f)
    finally:
        try:
            merger.close()
        except Exception:
            pass


class HymnCombinerApp:
    """Main application class for Hymn PDF Combiner."""
    
    def __init__(self, root):
        self.root = root
        self.setup_window()
        self.setup_paths()
        self.setup_ui()
        
    def setup_window(self):
        """Setup main window properties."""
        self.root.title("Hymn PDF Combiner")
        self.root.geometry("650x630")  # Slightly increased height for progress bar
        self.root.resizable(False, False)
        self.root.configure(bg=ModernStyle.BACKGROUND)
        
        # Center window on screen
        self.center_window()
        
    def setup_paths(self):
        """Setup application paths."""
        self.base_dir = Path(__file__).resolve().parent
        self.pdf_dir = self.base_dir / "pdf"
        self.output_dir = self.base_dir / "output"
        
    def setup_ui(self):
        """Setup user interface."""
        # Main container
        main_container = tk.Frame(self.root, bg=ModernStyle.BACKGROUND)
        main_container.pack(fill=tk.BOTH, expand=True, padx=30, pady=20)

        # Create sections
        self.create_header(main_container)
        self.create_input_section(main_container)
        self.create_progress_section(main_container)
        
    def create_header(self, parent):
        """Create header section."""
        header_frame = tk.Frame(parent, bg=ModernStyle.PRIMARY, relief="flat", bd=0)
        header_frame.pack(fill=tk.X, pady=(0, 20))
        
        title_container = tk.Frame(header_frame, bg=ModernStyle.PRIMARY)
        title_container.pack(pady=25)
        
        icon_label = tk.Label(title_container, text="📚", 
                           font=("Segoe UI", 28), 
                           bg=ModernStyle.PRIMARY, fg="white")
        icon_label.pack(side=tk.LEFT, padx=(0, 15))
        
        title_label = tk.Label(title_container, 
                             text="HYMN PDF Combiner",
                             font=ModernStyle.TITLE_FONT,
                             bg=ModernStyle.PRIMARY, fg="white")
        title_label.pack(side=tk.LEFT)
        
        info_frame = tk.Frame(header_frame, bg=ModernStyle.PRIMARY)
        info_frame.pack(pady=(20, 25))
        
        folder_text = f"📁 PDF: {self.pdf_dir.name}/  📤 Output: {self.output_dir.name}/"
        folder_label = tk.Label(info_frame, text=folder_text,
                              font=ModernStyle.LABEL_FONT,
                              bg=ModernStyle.PRIMARY, fg="#E0F2FE")
        folder_label.pack()
        
    def create_input_section(self, parent):
        """Create input section."""
        container = tk.Frame(parent, bg=ModernStyle.BACKGROUND)
        container.pack(fill=tk.BOTH, expand=True, padx=30, pady=10)
        
        # Hymn numbers input with better alignment
        hymn_label = tk.Label(container, text="Hymn Numbers",
                            font=ModernStyle.HEADER_FONT,
                            bg=ModernStyle.BACKGROUND, fg=ModernStyle.TEXT)
        hymn_label.grid(row=0, column=0, columnspan=2, sticky=tk.W, pady=(0, 10))
        
        # Hymn input and clear button
        self.hymn_entry = tk.Entry(container, 
                                  font=ModernStyle.LABEL_FONT,
                                  relief="solid", bd=2,
                                  bg="white",
                                  highlightthickness=0,
                                  highlightbackground=ModernStyle.BORDER,
                                  highlightcolor=ModernStyle.PRIMARY)
        self.hymn_entry.grid(row=1, column=0, sticky=tk.EW, padx=(0, 5), pady=(0, 5))
        
        clear_btn = tk.Button(container, text="Clear",
                            font=ModernStyle.BUTTON_FONT,
                            bg=ModernStyle.ERROR, fg="white",
                            relief="flat", bd=0,
                            width=8,
                            command=self.clear_hymn_input,
                            cursor="hand2")
        clear_btn.grid(row=1, column=1, sticky=tk.W, padx=(0, 0), pady=(0, 5))
        
        # Examples label below input
        examples_label = tk.Label(container, 
                               text="💡 Enter: 1 2 3, 1-5, or 1-3,8,10-12",
                               font=ModernStyle.LABEL_FONT,
                               bg=ModernStyle.BACKGROUND, fg=ModernStyle.TEXT_LIGHT)
        examples_label.grid(row=2, column=0, columnspan=2, sticky=tk.W, pady=(5, 15))
        
        # Output filename and merge button
        filename_label = tk.Label(container, text="Output Filename",
                                font=ModernStyle.HEADER_FONT,
                                bg=ModernStyle.BACKGROUND, fg=ModernStyle.TEXT)
        filename_label.grid(row=3, column=0, columnspan=2, sticky=tk.W, pady=(0, 10))
        
        self.filename_entry = tk.Entry(container,
                                      font=ModernStyle.LABEL_FONT,
                                      relief="solid", bd=2,
                                      bg="white",
                                      highlightthickness=0,
                                      highlightbackground=ModernStyle.BORDER,
                                      highlightcolor=ModernStyle.PRIMARY)
        self.filename_entry.grid(row=4, column=0, sticky=tk.EW, padx=(0, 5), pady=(0, 15))
        self.filename_entry.insert(0, "combined.pdf")
        
        # Merge button
        self.merge_button = tk.Button(
            container,
            text="🚀 Merge",
            font=ModernStyle.BUTTON_FONT,
            bg=ModernStyle.PRIMARY,
            fg="white",
            relief="flat",
            bd=0,
            width=8,
            command=self.start_merge,
            cursor="hand2",
        )
        self.merge_button.grid(row=4, column=1, sticky=tk.W, padx=(0, 0), pady=(0, 15))
        
        # Configure grid weights for proper expansion
        container.grid_columnconfigure(0, weight=1)
        container.grid_columnconfigure(1, weight=0)
        
        # Bind hover effects
        self.merge_button.bind("<Enter>", lambda e: self.merge_button.config(bg=ModernStyle.SECONDARY))
        self.merge_button.bind("<Leave>", lambda e: self.merge_button.config(bg=ModernStyle.PRIMARY))
        
    def create_progress_section(self, parent):
        """Create progress bar section with green theme."""
        progress_container = tk.Frame(parent, bg=ModernStyle.BACKGROUND)
        progress_container.pack(fill=tk.X, padx=30, pady=(15, 10))
        
        # Progress label
        self.progress_label = tk.Label(progress_container, 
                                      text="Status: Ready",
                                      font=ModernStyle.LABEL_FONT,
                                      bg=ModernStyle.BACKGROUND, fg=ModernStyle.TEXT)
        self.progress_label.pack(anchor=tk.W, pady=(0, 5))
        
        # Progress bar frame
        progress_frame = tk.Frame(progress_container, 
                                 bg="#E8F5E8",  # Light green background
                                 relief="solid",
                                 bd=1,
                                 height=12)
        progress_frame.pack(fill=tk.X, pady=(0, 5))
        progress_frame.pack_propagate(False)
        
        # Progress bar fill with green color
        self.progress_bar = tk.Frame(progress_frame, 
                                    bg=ModernStyle.PRIMARY,  # Green color
                                    height=8)
        self.progress_bar.place(x=1, y=2, width=0, height=8)
        
        # Progress percentage
        self.progress_percent = tk.Label(progress_container, 
                                        text="0%",
                                        font=ModernStyle.STATUS_FONT,
                                        bg=ModernStyle.BACKGROUND, fg=ModernStyle.SUCCESS)
        self.progress_percent.pack(anchor=tk.E)
        
    def center_window(self):
        """Center window on screen."""
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')
        
    def clear_hymn_input(self):
        """Clear the hymn number input field."""
        self.hymn_entry.delete(0, tk.END)
        self.hymn_entry.focus()
        
    def start_merge(self):
        """Start PDF merging process with validation."""
        hymn_input = self.hymn_entry.get().strip()
        if not hymn_input:
            messagebox.showerror("Error", "Please enter hymn numbers.")
            return
            
        output_filename = self.filename_entry.get().strip()
        if not output_filename:
            messagebox.showerror("Error", "Please enter output filename.")
            return
            
        if not output_filename.lower().endswith(".pdf"):
            output_filename += ".pdf"
        
        # Update UI state
        self.merge_button.config(
            text="Merging...",
            bg=ModernStyle.WARNING,
            state=tk.DISABLED
        )
        
        # Initialize progress bar
        self.update_progress(0, "Starting merge process...")
        
        # Clear hymn input for new task
        self.hymn_entry.delete(0, tk.END)
        
        # Run merge in separate thread
        thread = threading.Thread(target=self.merge_pdfs, args=(hymn_input, output_filename))
        thread.daemon = True
        thread.start()
        
    def update_progress(self, percentage, status_text):
        """Update progress bar and status text."""
        self.root.after(0, self._update_progress_ui, percentage, status_text)
        
    def _update_progress_ui(self, percentage, status_text):
        """Update progress UI elements (called from main thread)."""
        # Update status text
        self.progress_label.config(text=f"Status: {status_text}")
        
        # Update progress bar
        self.root.update_idletasks()
        progress_frame_width = self.progress_bar.master.winfo_width()
        if progress_frame_width > 1:
            bar_width = int((percentage / 100) * (progress_frame_width - 2))
            self.progress_bar.place(x=1, y=2, width=bar_width, height=8)
        
        # Update percentage text
        if percentage == 100:
            self.progress_percent.config(text="Done", fg=ModernStyle.SUCCESS)
        else:
            self.progress_percent.config(text=f"{percentage}%")
            
    def merge_pdfs(self, hymn_input, output_filename):
        """Perform actual PDF merging with progress tracking."""
        try:
            # Parse hymn numbers
            self.update_progress(10, "Parsing hymn numbers...")
            hymn_numbers = parse_hymn_numbers(hymn_input)
            if not hymn_numbers:
                self.update_progress(100, "Error - No valid hymn numbers")
                self.root.after(0, lambda: messagebox.showerror("Error", "No valid hymn numbers found."))
                return
                
            # Setup output path
            self.update_progress(20, "Setting up output path...")
            output_path = Path(output_filename)
            if not output_path.is_absolute():
                output_path = self.output_dir / output_path.name
                
            # Find PDF files
            self.update_progress(30, "Finding PDF files...")
            missing = []
            to_merge = []
            
            for i, n in enumerate(hymn_numbers):
                paths = pdf_paths_for_hymn(self.pdf_dir, n)
                if not paths:
                    missing.append(n)
                    continue
                to_merge.extend(paths)
                # Update progress during file search
                progress = 30 + int((i + 1) / len(hymn_numbers) * 20)
                self.update_progress(progress, f"Finding PDF files... ({i + 1}/{len(hymn_numbers)} hymns checked)")
                
            if not to_merge:
                self.update_progress(100, "Error - No PDF files found")
                error_msg = "No PDF files found for the provided hymn numbers."
                if missing:
                    error_msg += f"\nMissing: {', '.join(map(str, missing))}"
                self.root.after(0, lambda: messagebox.showerror("Error", error_msg))
                return
            
            # Merge PDFs
            self.update_progress(60, "Merging PDF files...")
            merge_pdfs(to_merge, output_path)
            self.update_progress(90, "Finalizing merge...")
            
            # Prepare success message
            success_msg = f"PDF merged successfully!\n\n"
            success_msg += f"Total hymns processed: {len(hymn_numbers)}\n"
            success_msg += f"Total PDF files merged: {len(to_merge)}\n\n"
            success_msg += f"Saved to: {output_path.absolute()}\n"
            success_msg += f"File size: {output_path.stat().st_size:,} bytes\n\n"
            success_msg += f"Would you like to open the output folder to see your merged PDF?"
            
            if missing:
                success_msg += f"\n\nNote: Missing hymn PDFs: {', '.join(map(str, missing))}"
            
            # Mark as Done before showing message box
            self.update_progress(100, "Done")
                
            # Show success message with option to open folder
            result = messagebox.askyesno("Success!", success_msg)
            if result:
                import subprocess
                import platform
                
                try:
                    if platform.system() == "Windows":
                        subprocess.run(["explorer", str(output_path.parent)], check=False)
                    elif platform.system() == "Darwin":  # macOS
                        subprocess.run(["open", str(output_path.parent)], check=False)
                    else:  # Linux
                        subprocess.run(["xdg-open", str(output_path.parent)], check=False)
                except Exception as e:
                    messagebox.showinfo("Info", f"Could not open folder automatically. Please check:\n{output_path.parent}")
            
        except ImportError as e:
            self.update_progress(100, "Error - Missing dependency")
            error_msg = f"Missing dependency: {e}\n\nInstall with: pip install pypdf"
            self.root.after(0, lambda: messagebox.showerror("Dependency Error", error_msg))
        except Exception as e:
            self.update_progress(100, "Error - Merge failed")
            self.root.after(0, lambda: messagebox.showerror("Error", f"Failed to merge PDFs: {e}"))
        finally:
            # Restore UI state
            self.root.after(0, self.restore_ui_state)
            
    def restore_ui_state(self):
        """Restore UI state after merge completion."""
        self.merge_button.config(
            text="🚀 Merge",
            bg=ModernStyle.PRIMARY,
            state=tk.NORMAL
        )
        # Reset progress bar to zero
        self.update_progress(0, "Ready")


def main():
    """Main entry point."""
    try:
        print("Starting Hymn PDF Combiner...")
        root = tk.Tk()
        app = HymnCombinerApp(root)
        root.mainloop()
    except Exception as e:
        print(f"Error: {e}")
        messagebox.showerror("Error", f"Failed to start application: {e}")


if __name__ == "__main__":
    main()
