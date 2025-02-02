import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pyautogui
import time
import os
import cv2
import keyboard
from PIL import Image, ImageGrab, ImageTk
import win32gui
import win32con
import sys
import json
import numpy as np

# Hide console window on launch
if sys.platform == 'win32':
    console_window = win32gui.GetForegroundWindow()
    win32gui.ShowWindow(console_window, win32con.SW_HIDE)

class ReferenceImageCaptureGUI:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Reference Image Capture Utility")
        self.root.attributes('-topmost', True)
        
        # Dark Logitech theme colors
        self.colors = {
            'bg': '#1E1E1E',
            'fg': '#FFFFFF',
            'button': '#00A6ED',
            'button_hover': '#0088CC',
            'frame': '#2D2D2D'
        }
        
        # Configure window
        self.root.configure(bg=self.colors['bg'])
        self.root.geometry("600x400")
        
        # Create menu bar
        self.create_menu()
        
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
                
        self.capturing = False
        self.current_element = None
        
        self.setup_ui()
        self.setup_keyboard_listener()
        self.update_mouse_position()

    def create_menu(self):
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        
        menubar.add_command(label="Copy")
        menubar.add_command(label="Log")
        menubar.add_command(label="Help")

    def setup_ui(self):
        # Menu and checkbox remain at top
        delay_frame = tk.Frame(self.root, bg=self.colors['bg'])
        delay_frame.pack(fill=tk.X, padx=5, pady=2)
        self.delay_var = tk.BooleanVar(value=True)
        tk.Checkbutton(delay_frame, text="3 Sec. Button Delay", 
                      variable=self.delay_var,
                      bg=self.colors['bg'],
                      fg=self.colors['fg'],
                      selectcolor=self.colors['frame']).pack(side=tk.LEFT)
        
        # Target and Reference Settings Frame
        settings_frame = tk.Frame(self.root, bg=self.colors['bg'])
        settings_frame.pack(fill=tk.X, padx=5, pady=2)
        
        # Add separator above settings
        separator1 = tk.Frame(settings_frame, height=2, bg=self.colors['button'])
        separator1.pack(fill=tk.X, pady=5)
        
        # Target Application
        target_frame = tk.Frame(settings_frame, bg=self.colors['bg'])
        target_frame.pack(fill=tk.X, pady=2)
        tk.Label(target_frame, text="Target Application:", 
                bg=self.colors['bg'], fg=self.colors['fg']).pack(side=tk.LEFT)
        self.target_path = tk.Entry(target_frame, bg=self.colors['frame'], fg=self.colors['fg'])
        self.target_path.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        tk.Button(target_frame, text="Browse", command=self.browse_target,
                 bg=self.colors['button'], fg=self.colors['fg']).pack(side=tk.RIGHT)
        
        # Reference Destination
        ref_frame = tk.Frame(settings_frame, bg=self.colors['bg'])
        ref_frame.pack(fill=tk.X, pady=2)
        tk.Label(ref_frame, text="Reference Destination:", 
                bg=self.colors['bg'], fg=self.colors['fg']).pack(side=tk.LEFT)
        self.ref_path = tk.Entry(ref_frame, bg=self.colors['frame'], fg=self.colors['fg'])
        self.ref_path.insert(0, self.ref_dir)
        self.ref_path.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        tk.Button(ref_frame, text="Browse", command=self.browse_ref_dir,
                 bg=self.colors['button'], fg=self.colors['fg']).pack(side=tk.RIGHT)
        
        # Add separator below settings
        separator2 = tk.Frame(settings_frame, height=2, bg=self.colors['button'])
        separator2.pack(fill=tk.X, pady=5)
        
        # XY Position with standardized sizes
        xy_frame = tk.Frame(self.root, bg=self.colors['bg'])
        xy_frame.pack(fill=tk.X, padx=5, pady=2)
        tk.Label(xy_frame, text="XY Position", width=10, anchor='w',
                bg=self.colors['bg'], fg=self.colors['fg']).pack(side=tk.LEFT)
        self.xy_entry = tk.Entry(xy_frame, width=15, bg=self.colors['frame'], fg=self.colors['fg'])
        self.xy_entry.pack(side=tk.LEFT, padx=5)
        tk.Button(xy_frame, text="Copy XY (F2)", width=12,
                 command=lambda: self.copy_value(self.xy_entry.get()),
                 bg=self.colors['button'], fg=self.colors['fg']).pack(side=tk.LEFT, padx=2)
        tk.Button(xy_frame, text="Log XY (F6)", width=12,
                 command=lambda: self.log_value(self.xy_entry.get()),
                 bg=self.colors['button'], fg=self.colors['fg']).pack(side=tk.LEFT, padx=2)
        
        # RGB Color with standardized sizes
        rgb_frame = tk.Frame(self.root, bg=self.colors['bg'])
        rgb_frame.pack(fill=tk.X, padx=5, pady=2)
        tk.Label(rgb_frame, text="RGB Color", width=10, anchor='w',
                bg=self.colors['bg'], fg=self.colors['fg']).pack(side=tk.LEFT)
        self.rgb_entry = tk.Entry(rgb_frame, width=15, bg=self.colors['frame'], fg=self.colors['fg'])
        self.rgb_entry.pack(side=tk.LEFT, padx=5)
        tk.Button(rgb_frame, text="Copy RGB (F3)", width=12,
                 command=lambda: self.copy_value(self.rgb_entry.get()),
                 bg=self.colors['button'], fg=self.colors['fg']).pack(side=tk.LEFT, padx=2)
        tk.Button(rgb_frame, text="Log RGB (F7)", width=12,
                 command=lambda: self.log_value(self.rgb_entry.get()),
                 bg=self.colors['button'], fg=self.colors['fg']).pack(side=tk.LEFT, padx=2)
        
        # RGB as Hex with standardized sizes
        hex_frame = tk.Frame(self.root, bg=self.colors['bg'])
        hex_frame.pack(fill=tk.X, padx=5, pady=2)
        tk.Label(hex_frame, text="RGB as Hex", width=10, anchor='w',
                bg=self.colors['bg'], fg=self.colors['fg']).pack(side=tk.LEFT)
        self.hex_entry = tk.Entry(hex_frame, width=15, bg=self.colors['frame'], fg=self.colors['fg'])
        self.hex_entry.pack(side=tk.LEFT, padx=5)
        tk.Button(hex_frame, text="Copy RGB Hex (F4)", width=12,
                 command=lambda: self.copy_value(self.hex_entry.get()),
                 bg=self.colors['button'], fg=self.colors['fg']).pack(side=tk.LEFT, padx=2)
        tk.Button(hex_frame, text="Log RGB Hex (F8)", width=12,
                 command=lambda: self.log_value(self.hex_entry.get()),
                 bg=self.colors['button'], fg=self.colors['fg']).pack(side=tk.LEFT, padx=2)
        
        # Color preview with standardized width
        color_frame = tk.Frame(self.root, bg=self.colors['bg'])
        color_frame.pack(fill=tk.X, padx=5, pady=2)
        tk.Label(color_frame, text="Color", width=10, anchor='w',
                bg=self.colors['bg'], fg=self.colors['fg']).pack(side=tk.LEFT)
        color_box_frame = tk.Frame(color_frame, bg=self.colors['bg'])
        color_box_frame.pack(side=tk.LEFT, padx=5)
        self.color_preview = tk.Canvas(color_box_frame, width=120, height=25,
                                     bg='white', highlightthickness=1,
                                     highlightbackground=self.colors['fg'])
        self.color_preview.pack(fill=tk.BOTH, expand=True)
        
        # Create main content frame for two-column layout
        content_frame = tk.Frame(self.root, bg=self.colors['bg'])
        content_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=2)
        
        # Left column for controls (narrower)
        left_column = tk.Frame(content_frame, bg=self.colors['bg'], width=160)
        left_column.pack(side=tk.LEFT, fill=tk.Y)
        left_column.pack_propagate(False)
        
        # XY Origin with standardized size
        origin_frame = tk.Frame(left_column, bg=self.colors['bg'])
        origin_frame.pack(fill=tk.X, pady=2)
        tk.Label(origin_frame, text="XY Origin", width=10, anchor='w',
                bg=self.colors['bg'], fg=self.colors['fg']).pack(side=tk.LEFT)
        self.origin_entry = tk.Entry(origin_frame, width=12, bg=self.colors['frame'], fg=self.colors['fg'])
        self.origin_entry.insert(0, "0, 0")
        self.origin_entry.pack(side=tk.LEFT, padx=5)

        # Control Buttons
        button_frame = tk.Frame(left_column, bg=self.colors['bg'])
        button_frame.pack(fill=tk.X, pady=2)
        
        # Create Start/Stop Capture button
        self.capture_button = tk.Button(button_frame, text="Start Capture",
                                      command=self.toggle_capture,
                                      bg=self.colors['button'], fg=self.colors['fg'])
        self.capture_button.pack(side=tk.TOP, pady=1, fill=tk.X)
        
        # Other buttons
        tk.Button(button_frame, text="Show Images",
                 command=self.show_captured_images,
                 bg=self.colors['button'], fg=self.colors['fg']).pack(side=tk.TOP, pady=1, fill=tk.X)
        tk.Button(button_frame, text="Launch Macro",
                 command=self.launch_macro,
                 bg=self.colors['button'], fg=self.colors['fg']).pack(side=tk.TOP, pady=1, fill=tk.X)

        # Right column (empty space)
        right_column = tk.Frame(content_frame, bg=self.colors['bg'])
        right_column.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5)

        # Status text area below both columns
        self.status_text = tk.Text(self.root, height=6, bg=self.colors['frame'], fg=self.colors['fg'])
        self.status_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Save options at bottom
        save_frame = tk.Frame(self.root, bg=self.colors['bg'])
        save_frame.pack(fill=tk.X, side=tk.BOTTOM, padx=5, pady=5)
        
        # Log save (first)
        log_frame = tk.Frame(save_frame, bg=self.colors['bg'])
        log_frame.pack(fill=tk.X, pady=2)
        self.log_path = tk.Entry(log_frame, bg=self.colors['frame'], fg=self.colors['fg'])
        self.log_path.insert(0, os.path.join(self.log_dir, "capture_log.txt"))
        self.log_path.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        tk.Button(log_frame, text="Save Log",
                 bg=self.colors['button'], fg=self.colors['fg']).pack(side=tk.RIGHT)
        
        # Screenshot save (second)
        screenshot_frame = tk.Frame(save_frame, bg=self.colors['bg'])
        screenshot_frame.pack(fill=tk.X, pady=2)
        self.screenshot_path = tk.Entry(screenshot_frame, bg=self.colors['frame'], fg=self.colors['fg'])
        self.screenshot_path.insert(0, os.path.join(self.screenshot_dir, "screenshot.png"))
        self.screenshot_path.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        tk.Button(screenshot_frame, text="Save Screenshot",
                 bg=self.colors['button'], fg=self.colors['fg']).pack(side=tk.RIGHT)

    def copy_value(self, value):
        self.root.clipboard_clear()
        self.root.clipboard_append(value)
        self.update_status(f"Copied: {value}")

    def log_value(self, value):
        # Implementation of log_value method
        pass

    def browse_target(self):
        """Browse for target application"""
        filename = filedialog.askopenfilename(
            title="Select Target Application",
            filetypes=[("Executable files", "*.exe"), ("All files", "*.*")]
        )
        if filename:
            self.target_path.insert(0, filename)
            self.update_status(f"Target set to: {filename}")
            
            # Try to find and focus the target window
            def enum_window_callback(hwnd, target_name):
                if win32gui.IsWindowVisible(hwnd):
                    window_title = win32gui.GetWindowText(hwnd)
                    if target_name in window_title:
                        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                        win32gui.SetForegroundWindow(hwnd)
                        return False
                return True
            
            win32gui.EnumWindows(enum_window_callback, os.path.basename(filename))

    def launch_macro(self):
        if not self.target_path.get():
            messagebox.showerror("Error", "Please select a target application first")
            return
            
        if not os.path.exists('macro_runner.py'):
            self.create_macro_runner()
            
        os.system('start pythonw macro_runner.py')
        self.update_status("Launched macro runner")
    
    def create_macro_runner(self):
        macro_code = '''
import pyautogui
import time
import json
import os
import win32gui
import win32con
import sys
from PIL import ImageGrab

# Hide console window
if sys.platform == 'win32':
    console_window = win32gui.GetForegroundWindow()
    win32gui.ShowWindow(console_window, win32con.SW_HIDE)

class MacroRunner:
    def __init__(self):
        self.load_config()
        
    def load_config(self):
        with open('ref_config.json', 'r') as f:
            self.config = json.load(f)
            
    def focus_target(self):
        """Focus the target application window"""
        windows = []
        win32gui.EnumWindows(lambda hwnd, windows: windows.append((hwnd, win32gui.GetWindowText(hwnd))), windows)
        
        for hwnd, title in windows:
            if self.config['target_exe'].split('\\')[-1].lower() in title.lower():
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                win32gui.SetForegroundWindow(hwnd)
                return True
        return False
        
    def find_and_click(self, element_name, confidence=0.8):
        """Find and click an element using its reference image"""
        try:
            location = pyautogui.locateOnScreen(f'ref_images/{element_name}.png', 
                                              confidence=confidence)
            if location:
                center = pyautogui.center(location)
                pyautogui.click(center)
                return True
            return False
        except Exception as e:
            print(f"Error finding {element_name}: {e}")
            return False
            
    def run(self):
        if not self.focus_target():
            print("Could not find target application window")
            return
            
        time.sleep(1)  # Wait for window to focus
        
        # Example macro sequence using captured reference points
        for element_name in self.config['elements']:
            if self.find_and_click(element_name):
                print(f"Clicked {element_name}")
                time.sleep(0.5)
            else:
                print(f"Could not find {element_name}")

if __name__ == "__main__":
    runner = MacroRunner()
    runner.run()
'''
        with open('macro_runner.py', 'w') as f:
            f.write(macro_code.strip())
    
    def setup_keyboard_listener(self):
        """Setup keyboard and mouse bindings"""
        # Change binding for inspection window to Ctrl + Shift + Left Click
        self.root.bind('<Control-Shift-Button-1>', lambda e: None)  # Let system handle inspection
        
        # Capture screenshot on Ctrl + Left Click
        self.root.bind('<Control-Button-1>', self.capture_at_cursor)
        
        # Window toggling
        self.root.bind('<Alt-Tab>', self.toggle_window_focus)
        self.root.bind('<F11>', self.toggle_window_focus)

    def toggle_capture(self):
        """Toggle capture mode on/off"""
        if not self.capturing:
            # Start capture mode
            self.capturing = True
            self.capture_button.configure(text="Stop Capture", bg='#FF4444')  # Red color
            
            # Launch target application if set
            target_app = self.target_path.get()
            if target_app:
                try:
                    # First minimize our utility window
                    self.root.iconify()
                    
                    # Find and focus target window
                    target_windows = []
                    def enum_window_callback(hwnd, target_windows):
                        if win32gui.IsWindowVisible(hwnd):
                            window_title = win32gui.GetWindowText(hwnd)
                            if target_app in window_title:
                                target_windows.append(hwnd)
                    win32gui.EnumWindows(enum_window_callback, target_windows)
                    
                    if target_windows:
                        # Focus existing window
                        target_hwnd = target_windows[0]
                        win32gui.ShowWindow(target_hwnd, win32con.SW_RESTORE)
                        win32gui.SetForegroundWindow(target_hwnd)
                    else:
                        # Try to launch the application
                        if os.path.exists(target_app):
                            os.startfile(target_app)
                        self.update_status(f"Launched target application: {target_app}")
                        
                    # Store target window handle for toggling
                    self.target_hwnd = target_windows[0] if target_windows else None
                    
                except Exception as e:
                    self.update_status(f"Error launching target application: {str(e)}")
            
            self.update_status("Capture mode started - Use Ctrl + Left Click to capture")
        else:
            # Stop capture mode
            self.capturing = False
            self.capture_button.configure(text="Start Capture", bg=self.colors['button'])
            self.update_status("Capture mode stopped")
            
            # Restore utility window
            self.root.deiconify()
            self.root.lift()
            self.root.focus_force()

    def capture_at_cursor(self, event=None):
        """Capture a square region around the mouse cursor"""
        if not self.capturing:
            return  # Only capture if in capture mode
            
        try:
            # Stop event propagation to prevent inspection window
            if event:
                event.widget.stop_propagation = True
            
            # Get current mouse position
            x, y = pyautogui.position()
            
            # Define capture region (20x20 pixels around cursor)
            region = (x-10, y-10, 20, 20)
            
            # Temporarily hide cursor for clean capture
            pyautogui.FAILSAFE = False
            original_visibility = win32gui.ShowCursor(False)
            
            # Capture the region
            screenshot = ImageGrab.grab(bbox=region)
            
            # Restore cursor
            win32gui.ShowCursor(True)
            pyautogui.FAILSAFE = True
            
            # Save the capture
            timestamp = time.strftime("%Y%m%d-%H%M%S")
            filename = f"capture_{timestamp}.png"
            save_path = os.path.join(self.save_dir, filename)
            screenshot.save(save_path)
            
            self.update_status(f"Captured region at ({x}, {y}) - saved as {filename}")
            
        except Exception as e:
            self.update_status(f"Capture failed: {str(e)}")
            
        finally:
            win32gui.ShowCursor(True)
            pyautogui.FAILSAFE = True
            return "break"  # Prevent event from propagating

    def toggle_window_focus(self, event=None):
        """Toggle between utility and target application"""
        if not hasattr(self, 'target_hwnd') or not self.target_hwnd:
            return
            
        try:
            current_hwnd = win32gui.GetForegroundWindow()
            utility_hwnd = self.root.winfo_id()
            
            # If utility is current window, switch to target
            if current_hwnd == utility_hwnd:
                self.root.iconify()  # Minimize utility
                win32gui.ShowWindow(self.target_hwnd, win32con.SW_RESTORE)
                win32gui.SetForegroundWindow(self.target_hwnd)
            # If target is current window, switch to utility
            elif current_hwnd == self.target_hwnd:
                win32gui.ShowWindow(self.target_hwnd, win32con.SW_MINIMIZE)
                self.root.deiconify()  # Restore utility
                self.root.lift()
                self.root.focus_force()
            else:
                # If neither is focused, bring target to front
                win32gui.ShowWindow(self.target_hwnd, win32con.SW_RESTORE)
                win32gui.SetForegroundWindow(self.target_hwnd)
            
        except Exception as e:
            self.update_status(f"Error toggling window focus: {str(e)}")

    def update_mouse_position(self):
        """Update mouse position display every 100ms"""
        try:
            x, y = pyautogui.position()
            color = pyautogui.pixel(x, y)
            hex_color = '#{:02x}{:02x}{:02x}'.format(color[0], color[1], color[2])
            
            self.xy_entry.delete(0, tk.END)
            self.xy_entry.insert(0, f"{x},{y}")
            
            self.rgb_entry.delete(0, tk.END)
            self.rgb_entry.insert(0, f"{color[0]},{color[1]},{color[2]}")
            
            self.hex_entry.delete(0, tk.END)
            self.hex_entry.insert(0, hex_color)
            
            self.color_preview.configure(bg=hex_color)
        except:
            pass
        finally:
            self.root.after(100, self.update_mouse_position)
        
    def update_status(self, message, clear=False):
        self.status_text.configure(state='normal')
        if clear:
            self.status_text.delete(1.0, tk.END)
        self.status_text.insert(tk.END, message + "\n")
        self.status_text.see(tk.END)
        self.status_text.configure(state='disabled')
        
    def show_captured_images(self):
        images = []
        for filename in os.listdir(self.save_dir):
            if filename.endswith('.png'):
                element_name = filename[:-4]
                img = cv2.imread(f'{self.save_dir}/{filename}')
                if img is not None:
                    height, width = img.shape[:2]
                    images.append(f"{element_name}: {width}x{height} pixels")
        
        if images:
            self.update_status("\nCaptured Reference Images:\n" + "\n".join(images))
        else:
            self.update_status("\nNo reference images found")
        
    def run(self):
        self.root.mainloop()

    def save_recording(self):
        if hasattr(self, 'current_recording') and self.current_recording is not None:
            try:
                save_path = self.recording_path.get()
                self.current_recording.write(save_path)
                self.update_status(f"Recording saved to: {save_path}")
            except Exception as e:
                self.update_status(f"Error saving recording: {str(e)}")
        else:
            self.update_status("No recording available to save")

    def start_recording(self):
        """Start screen recording"""
        try:
            x, y = pyautogui.position()
            region = (x-50, y-50, 100, 100)  # Record 100x100 area around mouse
            self.current_recording = cv2.VideoWriter(
                self.recording_path.get(),
                cv2.VideoWriter_fourcc(*'mp4v'),
                30.0, (100, 100)
            )
            self.recording = True
            self.update_status("Recording started...")
        except Exception as e:
            self.update_status(f"Error starting recording: {str(e)}")

    def stop_recording(self):
        """Stop screen recording"""
        if hasattr(self, 'current_recording') and self.current_recording is not None:
            self.current_recording.release()
            self.recording = False
            self.update_status("Recording stopped")

    def toggle_recording(self, event=None):
        """Toggle screen recording on/off"""
        if not self.recording:
            self.start_recording()
        else:
            self.stop_recording()

    def browse_ref_dir(self):
        """Browse for reference directory"""
        directory = filedialog.askdirectory(
            title="Select Reference Directory",
            initialdir=self.ref_dir
        )
        if directory:
            self.ref_dir = directory
            self.ref_path.delete(0, tk.END)
            self.ref_path.insert(0, directory)
            
            # Update all subdirectories
            self.save_dir = os.path.join(self.ref_dir, 'ref_images')
            self.log_dir = os.path.join(self.ref_dir, 'logs')
            self.video_dir = os.path.join(self.ref_dir, 'ref_videos')
            self.screenshot_dir = os.path.join(self.ref_dir, 'screenshots')
            
            # Create directories if they don't exist
            for directory in [self.save_dir, self.log_dir, self.video_dir, self.screenshot_dir]:
                if not os.path.exists(directory):
                    os.makedirs(directory)
            
            # Update paths in entry fields
            self.log_path.delete(0, tk.END)
            self.log_path.insert(0, os.path.join(self.log_dir, "capture_log.txt"))
            
            self.screenshot_path.delete(0, tk.END)
            self.screenshot_path.insert(0, os.path.join(self.screenshot_dir, "screenshot.png"))
            
            self.recording_path.delete(0, tk.END)
            self.recording_path.insert(0, os.path.join(self.video_dir, "recording.mp4"))
            
            self.update_status(f"Reference directory set to: {self.ref_dir}")

    def update_preview(self, image=None):
        """Update the preview canvas with captured image"""
        if image:
            # Convert PIL Image to PhotoImage
            photo = ImageTk.PhotoImage(image)
            # Keep a reference to prevent garbage collection
            self.current_preview = photo
            # Update canvas
            self.preview_canvas.delete("all")
            self.preview_canvas.create_image(0, 0, anchor="nw", image=photo)
        else:
            self.preview_canvas.delete("all")

    def get_current_application(self):
        """Get the currently focused application"""
        try:
            hwnd = win32gui.GetForegroundWindow()
            window_title = win32gui.GetWindowText(hwnd)
            
            # Don't include our utility or file explorer
            if (window_title != self.root.title() and 
                "File Explorer" not in window_title):
                self.target_path.insert(0, window_title)
                self.update_status(f"Current application set to: {window_title}")
                return hwnd
            
        except Exception as e:
            self.update_status(f"Error getting current application: {str(e)}")
        
        return None

if __name__ == "__main__":
    app = ReferenceImageCaptureGUI()
    app.run() 