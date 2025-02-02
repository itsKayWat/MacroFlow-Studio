import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import json
import os
import pyautogui
import time
import keyboard
import win32gui
import win32con
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from PIL import Image, ImageTk

class MacroRecorder:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Macro Recorder")
        
        # Dark theme colors
        self.colors = {
            'bg': '#1E1E1E',
            'fg': '#FFFFFF',
            'button': '#00A6ED',
            'frame': '#2D2D2D'
        }
        
        self.root.configure(bg=self.colors['bg'])
        self.root.geometry("800x600")
        
        # Initialize macro steps
        self.macro_steps = []
        self.current_step = 0
        
        # Get script directory for relative paths
        self.script_dir = os.path.dirname(os.path.abspath(__file__))
        
        # Initialize directories relative to script location
        self.ref_dir = self.script_dir  # Default reference directory is script location
        self.save_dir = os.path.join(self.ref_dir, 'ref_images')
        self.log_dir = os.path.join(self.ref_dir, 'logs')
        self.video_dir = os.path.join(self.ref_dir, 'ref_videos')
        self.screenshot_dir = os.path.join(self.ref_dir, 'screenshots')
        
        # Create directories if they don't exist
        for directory in [self.save_dir, self.log_dir, self.video_dir, self.screenshot_dir]:
            if not os.path.exists(directory):
                os.makedirs(directory)
        
        self.setup_ui()
        self.load_macro()

    def setup_ui(self):
        # Main container
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Left panel - Step list
        left_panel = ttk.Frame(main_frame)
        left_panel.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Step list with scrollbar
        self.step_list = tk.Listbox(left_panel, bg=self.colors['frame'], fg=self.colors['fg'])
        scrollbar = ttk.Scrollbar(left_panel, orient=tk.VERTICAL, command=self.step_list.yview)
        self.step_list.configure(yscrollcommand=scrollbar.set)
        self.step_list.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Right panel - Step configuration
        right_panel = ttk.Frame(main_frame)
        right_panel.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=(10, 0))
        
        # Action type selector
        action_frame = ttk.LabelFrame(right_panel, text="Action Type")
        action_frame.pack(fill=tk.X, pady=5)
        
        self.action_var = tk.StringVar()
        actions = [
            "Open Website",
            "Click Reference Image",
            "Type Text",
            "Press Key",
            "Wait",
            "Mouse Move",
            "Click",
            "Right Click",
            "Double Click",
            "Copy Text",
            "Paste Text",
            "Wait for Image",
            "Custom JavaScript",
            "CSS Selector Click",
            "XPath Click",
            "Load Reference Point"
        ]
        
        self.action_combo = ttk.Combobox(action_frame, textvariable=self.action_var, values=actions)
        self.action_combo.pack(fill=tk.X, pady=5)
        self.action_combo.bind('<<ComboboxSelected>>', self.update_parameter_frame)
        
        # Parameters frame
        self.param_frame = ttk.LabelFrame(right_panel, text="Parameters")
        self.param_frame.pack(fill=tk.X, pady=5)
        
        # Buttons frame
        button_frame = ttk.Frame(right_panel)
        button_frame.pack(fill=tk.X, pady=5)
        
        ttk.Button(button_frame, text="Add Step", command=self.add_step).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Remove Step", command=self.remove_step).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Move Up", command=self.move_step_up).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Move Down", command=self.move_step_down).pack(side=tk.LEFT, padx=5)
        
        # Control buttons
        control_frame = ttk.Frame(right_panel)
        control_frame.pack(fill=tk.X, pady=10)
        
        ttk.Button(control_frame, text="Save Macro", command=self.save_macro).pack(side=tk.LEFT, padx=5)
        ttk.Button(control_frame, text="Load Macro", command=self.load_macro).pack(side=tk.LEFT, padx=5)
        ttk.Button(control_frame, text="Run Macro", command=self.run_macro).pack(side=tk.LEFT, padx=5)
        
        # Status bar
        self.status_var = tk.StringVar()
        status_bar = ttk.Label(self.root, textvariable=self.status_var, relief=tk.SUNKEN)
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)

    def update_parameter_frame(self, event=None):
        # Clear existing parameters
        for widget in self.param_frame.winfo_children():
            widget.destroy()
            
        action = self.action_var.get()
        
        # Common reference selector frame - available for all actions
        ref_frame = ttk.LabelFrame(self.param_frame, text="Reference Target (Optional)")
        ref_frame.pack(fill=tk.X, pady=5)
        
        # Reference source selector (renamed from type)
        ttk.Label(ref_frame, text="Reference Source:").pack()
        self.ref_type_var = tk.StringVar()
        ref_sources = ["None", "Image", "Video", "CSS/HTML", "Text", "Coordinates"]
        self.ref_type_combo = ttk.Combobox(ref_frame, 
                                          textvariable=self.ref_type_var,
                                          values=ref_sources)
        self.ref_type_combo.set("None")
        self.ref_type_combo.pack(fill=tk.X, pady=2)
        
        # Reference file selector
        ref_file_frame = ttk.Frame(ref_frame)
        ref_file_frame.pack(fill=tk.X, pady=2)
        self.ref_path_entry = ttk.Entry(ref_file_frame)
        self.ref_path_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Button(ref_file_frame, text="Browse", 
                   command=self.browse_reference).pack(side=tk.RIGHT, padx=5)
        
        # Preview button
        ttk.Button(ref_frame, text="Preview Reference",
                   command=self.preview_reference).pack(pady=2)
        
        # Action-specific parameters
        param_frame = ttk.LabelFrame(self.param_frame, text="Action Parameters")
        param_frame.pack(fill=tk.X, pady=5)
        
        if action == "Open Website":
            ttk.Label(param_frame, text="URL:").pack()
            self.url_entry = ttk.Entry(param_frame)
            self.url_entry.pack(fill=tk.X)
            
        elif action == "Click Reference Image":
            ttk.Label(param_frame, text="Confidence (0-1):").pack()
            self.confidence_entry = ttk.Entry(param_frame)
            self.confidence_entry.insert(0, "0.9")
            self.confidence_entry.pack(fill=tk.X)
        
        elif action == "Type Text":
            ttk.Label(param_frame, text="Text to Type:").pack()
            self.text_entry = ttk.Entry(param_frame)
            self.text_entry.pack(fill=tk.X)
            
        elif action == "Wait":
            ttk.Label(param_frame, text="Seconds:").pack()
            self.wait_entry = ttk.Entry(param_frame)
            self.wait_entry.insert(0, "1")
            self.wait_entry.pack(fill=tk.X)
        
        # ... other action types ...

    def add_step(self):
        action = self.action_var.get()
        params = self.get_current_parameters()
        
        step = {
            "action": action,
            "params": params
        }
        
        self.macro_steps.append(step)
        self.update_step_list()
        self.status_var.set(f"Added step: {action}")

    def get_current_parameters(self):
        action = self.action_var.get()
        params = {}
        
        # Get reference parameters if set
        if self.ref_type_var.get() != "None":
            params["reference"] = {
                "type": self.ref_type_var.get(),
                "path": self.ref_path_entry.get()
            }
        
        # Get action-specific parameters
        if action == "Click Reference Image":
            params["confidence"] = float(self.confidence_entry.get())
        elif action == "Type Text":
            params["text"] = self.text_entry.get()
        elif action == "Wait":
            params["seconds"] = float(self.wait_entry.get())
        
        return params

    def update_step_list(self):
        self.step_list.delete(0, tk.END)
        for i, step in enumerate(self.macro_steps):
            self.step_list.insert(tk.END, f"{i+1}. {step['action']}")

    def save_macro(self):
        try:
            with open("macro_config.json", "w") as f:
                json.dump(self.macro_steps, f, indent=4)
            self.status_var.set("Macro saved successfully")
        except Exception as e:
            self.status_var.set(f"Error saving macro: {str(e)}")

    def load_macro(self):
        try:
            if os.path.exists("macro_config.json"):
                with open("macro_config.json", "r") as f:
                    self.macro_steps = json.load(f)
                self.update_step_list()
                self.status_var.set("Macro loaded successfully")
        except Exception as e:
            self.status_var.set(f"Error loading macro: {str(e)}")

    def run_macro(self):
        self.status_var.set("Running macro...")
        self.root.update()
        
        try:
            for i, step in enumerate(self.macro_steps):
                self.current_step = i
                self.step_list.selection_clear(0, tk.END)
                self.step_list.selection_set(i)
                self.step_list.see(i)
                
                self.execute_step(step)
                
            self.status_var.set("Macro completed successfully")
        except Exception as e:
            self.status_var.set(f"Error running macro: {str(e)}")

    def execute_step(self, step):
        action = step["action"]
        params = step["params"]
        
        # Handle reference if present
        reference = params.get("reference")
        if reference:
            self.current_reference = self.load_reference_data(reference["path"], reference["type"])
        
        # Execute action
        if action == "Click Reference Image":
            if self.current_reference and isinstance(self.current_reference, str):  # Image path
                location = pyautogui.locateOnScreen(
                    self.current_reference,
                    confidence=params.get("confidence", 0.9)
                )
                if location:
                    pyautogui.click(pyautogui.center(location))
                else:
                    raise Exception(f"Could not find reference image")
        
        # ... rest of action handling ...

    def remove_step(self):
        selection = self.step_list.curselection()
        if selection:
            index = selection[0]
            del self.macro_steps[index]
            self.update_step_list()

    def move_step_up(self):
        selection = self.step_list.curselection()
        if selection and selection[0] > 0:
            index = selection[0]
            self.macro_steps[index], self.macro_steps[index-1] = \
                self.macro_steps[index-1], self.macro_steps[index]
            self.update_step_list()
            self.step_list.selection_set(index-1)

    def move_step_down(self):
        selection = self.step_list.curselection()
        if selection and selection[0] < len(self.macro_steps) - 1:
            index = selection[0]
            self.macro_steps[index], self.macro_steps[index+1] = \
                self.macro_steps[index+1], self.macro_steps[index]
            self.update_step_list()
            self.step_list.selection_set(index+1)

    def preview_reference(self):
        """Preview the selected reference point"""
        ref_name = self.ref_path_entry.get()
        ref_type = self.ref_type_var.get()
        
        try:
            if ref_type == "Image":
                # Load and display image reference
                img_path = ref_name
                if os.path.exists(img_path):
                    img = Image.open(img_path)
                    self.show_preview_window(img, f"Image Reference: {ref_name}")
                else:
                    raise FileNotFoundError(f"Image reference not found: {ref_name}")
                
            elif ref_type == "Video":
                # Load video reference
                video_path = ref_name
                if os.path.exists(video_path):
                    os.startfile(video_path)  # Open with default video player
                else:
                    raise FileNotFoundError(f"Video reference not found: {ref_name}")
                
            elif ref_type == "CSS/HTML":
                # Load CSS/HTML reference
                selector_path = ref_name
                if os.path.exists(selector_path):
                    with open(selector_path, 'r') as f:
                        selector = f.read()
                    self.show_text_preview(selector, f"CSS/HTML Reference: {ref_name}")
                else:
                    raise FileNotFoundError(f"CSS/HTML reference not found: {ref_name}")
                
            elif ref_type == "Text":
                # Load text reference
                text_path = ref_name
                if os.path.exists(text_path):
                    with open(text_path, 'r') as f:
                        text = f.read()
                    self.show_text_preview(text, f"Text Reference: {ref_name}")
                else:
                    raise FileNotFoundError(f"Text reference not found: {ref_name}")
                
            elif ref_type == "Coordinates":
                # Load coordinate reference
                coord_path = ref_name
                if os.path.exists(coord_path):
                    with open(coord_path, 'r') as f:
                        coords = json.load(f)
                    self.show_text_preview(str(coords), f"Coordinates: {ref_name}")
                else:
                    raise FileNotFoundError(f"Coordinates not found: {ref_name}")
        
        except Exception as e:
            messagebox.showerror("Preview Error", str(e))

    def show_preview_window(self, image, title):
        """Show image preview in a new window"""
        preview = tk.Toplevel(self.root)
        preview.title(title)
        preview.geometry("400x400")
        
        # Create canvas for image
        canvas = tk.Canvas(preview)
        canvas.pack(fill=tk.BOTH, expand=True)
        
        # Resize image to fit window while maintaining aspect ratio
        img_width, img_height = image.size
        scale = min(380/img_width, 380/img_height)
        new_width = int(img_width * scale)
        new_height = int(img_height * scale)
        image = image.resize((new_width, new_height), Image.Resampling.LANCZOS)
        
        # Display image
        photo = ImageTk.PhotoImage(image)
        canvas.create_image(200, 200, image=photo, anchor=tk.CENTER)
        preview.photo = photo  # Keep reference to prevent garbage collection

    def show_text_preview(self, text, title):
        """Show text preview in a new window"""
        preview = tk.Toplevel(self.root)
        preview.title(title)
        preview.geometry("400x400")
        
        # Create text widget
        text_widget = tk.Text(preview, wrap=tk.WORD)
        text_widget.pack(fill=tk.BOTH, expand=True)
        
        # Insert text and make read-only
        text_widget.insert('1.0', text)
        text_widget.configure(state='disabled')

    def load_reference_data(self, ref_name, ref_type):
        """Load reference data based on type"""
        try:
            if ref_type == "Image":
                return ref_name
            elif ref_type == "Video":
                return ref_name
            elif ref_type == "CSS/HTML":
                return ref_name
            elif ref_type == "Text":
                return ref_name
            elif ref_type == "Coordinates":
                return ref_name
        except Exception as e:
            raise Exception(f"Error loading reference {ref_name}: {str(e)}")

    def browse_reference(self):
        """Browse for reference file based on selected type"""
        ref_type = self.ref_type_var.get()
        if ref_type == "None":
            return
        
        # Set up file dialog options based on reference type
        file_types = {
            "Image": [("PNG files", "*.png")],
            "Video": [("MP4 files", "*.mp4")],
            "CSS/HTML": [("Text files", "*.txt")],
            "Text": [("Text files", "*.txt")],
            "Coordinates": [("JSON files", "*.json")]
        }
        
        # Set initial directory based on reference type
        init_dirs = {
            "Image": self.save_dir,
            "Video": self.video_dir,
            "CSS/HTML": os.path.join(self.ref_dir, "selectors"),
            "Text": os.path.join(self.ref_dir, "text"),
            "Coordinates": os.path.join(self.ref_dir, "coordinates")
        }
        
        filename = filedialog.askopenfilename(
            title=f"Select {ref_type} Reference",
            initialdir=init_dirs.get(ref_type, self.ref_dir),
            filetypes=file_types.get(ref_type, [("All files", "*.*")])
        )
        
        if filename:
            self.ref_path_entry.delete(0, tk.END)
            self.ref_path_entry.insert(0, filename)

    def create_menu(self):
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        
        # Help menu
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Help", menu=help_menu)
        help_menu.add_command(label="Instructions", command=self.show_instructions)
        help_menu.add_command(label="About", command=self.show_about)

    def show_instructions(self):
        instructions = """
Macro Recorder Instructions:

1. Reference Sources:
   • Image - For clicking or verifying visual elements
   • Video - For reviewing complex sequences
   • CSS/HTML - For web element selection
   • Text - For content verification
   • Coordinates - For specific screen locations

2. Creating a Macro:
   • Select an Action Type from dropdown
   • Choose a Reference Source if needed
   • Browse and select your reference file
   • Preview to verify the reference
   • Configure any action parameters
   • Click Add Step

3. Managing Steps:
   • Add Step - Add the configured step
   • Remove Step - Delete selected step
   • Move Up/Down - Reorder steps
   • Save Macro - Save your sequence
   • Load Macro - Load existing sequence
   • Run Macro - Execute all steps

4. Tips:
   • Use Reference Capture tool first to gather references
   • Test steps individually before running full sequence
   • Add Wait steps between actions for reliability
   • Preview references before adding steps
   • Save your macro regularly

Keyboard Shortcuts:
• Ctrl+S - Save Macro
• Ctrl+O - Load Macro
• F5 - Run Macro
• Delete - Remove Step
• Ctrl+Up/Down - Move Step
"""
        
        help_window = tk.Toplevel(self.root)
        help_window.title("Macro Recorder Instructions")
        help_window.geometry("600x700")
        
        # Use same theme as main window
        help_window.configure(bg=self.colors['bg'])
        
        # Create text widget with scrollbar
        text_frame = ttk.Frame(help_window)
        text_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        scrollbar = ttk.Scrollbar(text_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        text_widget = tk.Text(text_frame, 
                             wrap=tk.WORD, 
                             yscrollcommand=scrollbar.set,
                             bg=self.colors['frame'],
                             fg=self.colors['fg'],
                             font=('Arial', 10))
        text_widget.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        scrollbar.config(command=text_widget.yview)
        
        # Insert instructions
        text_widget.insert('1.0', instructions)
        text_widget.configure(state='disabled')  # Make read-only

    def show_about(self):
        about_text = """
Macro Recorder v1.0

A companion tool to Reference Capture for creating automated sequences.
Works with captured reference points to build automated workflows.

• Works with Reference Capture tool
• Supports multiple reference sources
• Visual and web automation
• Customizable action sequences
"""
        
        messagebox.showinfo("About Macro Recorder", about_text)

    def run(self):
        self.root.mainloop()

if __name__ == "__main__":
    app = MacroRecorder()
    app.run()