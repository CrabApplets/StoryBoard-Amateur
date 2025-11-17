#!/usr/bin/env python3

"""StoryBoard Amateur
Copyright (c) 2025 Johanner Corrales, Lauren Oquendo, Paulo Cao Suarez,  Danilo Bodden, Carlos Jerak

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), 
to deal with the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, 
sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the 
following conditions: 

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software. 

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, 
INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A 
PARTICULAR PURPOSE, AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR 
COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES, OR OTHER LIABILITY, WHETHER IN 
AN ACTION OF CONTRACT, TORT, OR OTHERWISE, ARISING FROM, OUT OF, OR IN CONNECTION 
WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
"""


import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from PIL import Image, ImageTk
import json
import csv
from datetime import datetime
import os
from typing import List, Dict, Optional

# Try to import reportlab for PDF generation
try:
    from reportlab.lib.pagesizes import letter
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image as RLImage, Table, TableStyle
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import inch
    from reportlab.lib import colors
    from reportlab.lib.enums import TA_CENTER
    REPORTLAB_AVAILABLE = True
except ImportError:
    REPORTLAB_AVAILABLE = False
    TA_CENTER = 1  # Fallback value
    print("Warning: reportlab not available. PDF export will be disabled.")


class StoryboardScene:
    """Represents a single storyboard scene"""
    
    def __init__(self, scene_id: str, title: str = "", description: str = ""):
        self.scene_id = scene_id
        self.title = title
        self.description = description
        self.image_path = ""
        self.film_tip = ""
        self.edit_tip = ""
        # Changed from single clip_type to list of clip_types for multiple selection
        self.clip_types = ["Video"]  # List of: Video, Still, Audio, Title
        self.length = 0  # in seconds
        self.timestamp = datetime.now().isoformat()
        # New fields for audio and video file associations
        self.audio_path = ""  # Path to associated audio file
        self.video_path = ""  # Path to associated video file
        
        # For backward compatibility, maintain clip_type property
        self._clip_type = "Video"
    
    @property
    def clip_type(self):
        """Backward compatibility: return first clip type or empty string"""
        return self.clip_types[0] if self.clip_types else ""
    
    @clip_type.setter
    def clip_type(self, value):
        """Backward compatibility: set first clip type"""
        if value and value not in self.clip_types:
            self.clip_types = [value]
        elif not value:
            self.clip_types = []
    
    def to_dict(self) -> Dict:
        """Convert scene to dictionary for serialization"""
        return {
            'scene_id': self.scene_id,
            'title': self.title,
            'description': self.description,
            'image_path': self.image_path,
            'film_tip': self.film_tip,
            'edit_tip': self.edit_tip,
            'clip_type': self.clip_type,  # For backward compatibility
            'clip_types': self.clip_types,  # New field for multiple selections
            'length': self.length,
            'timestamp': self.timestamp,
            'audio_path': self.audio_path,
            'video_path': self.video_path
        }
    
    @classmethod
    def from_dict(cls, data: Dict) -> 'StoryboardScene':
        """Create scene from dictionary"""
        scene = cls(data['scene_id'], data.get('title', ''), data.get('description', ''))
        scene.image_path = data.get('image_path', '')
        scene.film_tip = data.get('film_tip', '')
        scene.edit_tip = data.get('edit_tip', '')
        
        # Handle both old and new format
        if 'clip_types' in data:
            # Use the saved clip_types list (can be empty list)
            scene.clip_types = data.get('clip_types', [])
        else:
            # Legacy format: single clip_type
            old_clip_type = data.get('clip_type', 'Video')
            # If old_clip_type is empty or None, use empty list; otherwise use the value
            scene.clip_types = [old_clip_type] if old_clip_type else []
        
        scene.length = data.get('length', 0)
        scene.timestamp = data.get('timestamp', datetime.now().isoformat())
        scene.audio_path = data.get('audio_path', '')
        scene.video_path = data.get('video_path', '')
        return scene


class StoryboardProject:
    """Represents a complete storyboard project"""
    
    def __init__(self, name: str = "Untitled Storyboard"):
        self.name = name
        self.scenes: List[StoryboardScene] = []
        self.project_path = ""
        self.created = datetime.now().isoformat()
        self.modified = datetime.now().isoformat()
        self.current_order = 1
        # New fields for PDF header
        self.project_title = ""  # Custom title for PDF header
        self.creators = []  # Creator names for PDF header (list)
    
    def add_scene(self, scene: StoryboardScene):
        """Add a scene to the project"""
        self.scenes.append(scene)
        self.modified = datetime.now().isoformat()
    
    def remove_scene(self, scene_id: str):
        """Remove a scene by ID"""
        self.scenes = [s for s in self.scenes if s.scene_id != scene_id]
        self.modified = datetime.now().isoformat()
    
    def get_scene(self, scene_id: str) -> Optional[StoryboardScene]:
        """Get a scene by ID"""
        for scene in self.scenes:
            if scene.scene_id == scene_id:
                return scene
        return None
    
    def to_dict(self) -> Dict:
        """Convert project to dictionary for serialization"""
        return {
            'name': self.name,
            'scenes': [scene.to_dict() for scene in self.scenes],
            'project_path': self.project_path,
            'created': self.created,
            'modified': self.modified,
            'current_order': self.current_order,
            'project_title': self.project_title,
            'creators': self.creators,
            'theme': getattr(self, 'theme', 'Default')  # Save theme
        }
    
    @classmethod
    def from_dict(cls, data: Dict) -> 'StoryboardProject':
        """Create project from dictionary"""
        project = cls(data.get('name', 'Untitled Storyboard'))
        project.project_path = data.get('project_path', '')
        project.created = data.get('created', datetime.now().isoformat())
        project.modified = data.get('modified', datetime.now().isoformat())
        project.current_order = data.get('current_order', 1)
        project.project_title = data.get('project_title', '')
        # Handle both old string format and new list format for creators
        creators_data = data.get('creators', [])
        if isinstance(creators_data, str):
            # Legacy format: split string into list, or keep empty list if empty string
            project.creators = [name.strip() for name in creators_data.split(',') if name.strip()] if creators_data else []
        else:
            # New format: already a list
            project.creators = creators_data if isinstance(creators_data, list) else []
        
        # Load theme
        project.theme = data.get('theme', 'Default')
        
        for scene_data in data.get('scenes', []):
            scene = StoryboardScene.from_dict(scene_data)
            project.add_scene(scene)
        
        return project


class StoryboardApp:
    """Main application class with improved UI"""
    
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("StoryBoard Amateur")
        
        # Get screen dimensions for centering and scaling
        self.screen_width = self.root.winfo_screenwidth()
        self.screen_height = self.root.winfo_screenheight()
        
        # Calculate scaling factor based on screen physical size
        self.scale_factor = self.calculate_scale_factor()
        
        # Debug information
        print(f"Screen: {self.screen_width}x{self.screen_height}")
        print(f"Scale factor: {self.scale_factor:.2f}")
        
        # Set window size as a function of screen size with scaling
        # Use smaller percentages for smaller screens
        if self.scale_factor > 1.2:  # Small screen (laptop)
            width_percent = 0.6  # Smaller window on laptop
            height_percent = 0.5
        else:  # Large screen (desktop)
            width_percent = 0.7  # Normal window on desktop
            height_percent = 0.6
            
        self.window_width = max(int(self.screen_width * width_percent), int(1000 * self.scale_factor))
        self.window_height = max(int(self.screen_height * height_percent), int(600 * self.scale_factor))
        
        # Center the window on screen
        x = (self.screen_width - self.window_width) // 2
        y = (self.screen_height - self.window_height) // 2
        self.root.geometry(f"{self.window_width}x{self.window_height}+{x}+{y}")
        self.root.minsize(800, 600)
        # Window background will be set by theme
        
        # Keyboard shortcuts
        self.root.bind('<Control-s>', lambda e: self.save_project())
        self.root.bind('<Control-n>', lambda e: self.new_project())
        
        # Current project
        self.current_project = StoryboardProject()
        
        # Tip mode - now filters by Film Tips or Edit Tips
        self.current_tip_mode = "Film Tips"
        
        # Theme support
        self.current_theme = "Default"
        self.themes = {
            "Default": {
                "bg_main": "#E0E0E0",
                "bg_header": "#4A90E2", 
                "bg_drag": "#D0D0D0",
                "bg_status": "#C0C0C0",
                "text_header": "white",
                "text_drag": "#666666",
                "text_main": "black"
            },
            "Dark": {
                "bg_main": "#2C2C2C",
                "bg_header": "#1A1A1A",
                "bg_drag": "#404040", 
                "bg_status": "#333333",
                "text_header": "white",
                "text_drag": "#CCCCCC",
                "text_main": "white"
            },
            "Ocean": {
                "bg_main": "#E6F3FF",
                "bg_header": "#0066CC",
                "bg_drag": "#B3D9FF",
                "bg_status": "#80BFFF", 
                "text_header": "white",
                "text_drag": "#003366",
                "text_main": "black"
            },
            "Forest": {
                "bg_main": "#E6F7E6",
                "bg_header": "#228B22",
                "bg_drag": "#B3E6B3",
                "bg_status": "#90EE90",
                "text_header": "white", 
                "text_drag": "#006400",
                "text_main": "black"
            },
            "Sunset": {
                "bg_main": "#FFF0E6",
                "bg_header": "#FF6600",
                "bg_drag": "#FFB366",
                "bg_status": "#FF9933",
                "text_header": "white",
                "text_drag": "#CC3300",
                "text_main": "black"
            }
        }
        
        # Load film tips from CSV
        self.film_tips = self.load_film_tips()
        
        # UI Components
        self.setup_ui()
        self.setup_menu()
        
        # Create default scene
        self.create_new_scene()
        
        # Show project properties dialog when app starts
        self.root.after(100, self.show_project_properties)
    
    def calculate_scale_factor(self):
        """Calculate scaling factor based on screen DPI and physical size"""
        try:
            # Get screen DPI (dots per inch)
            dpi = self.root.winfo_fpixels('1i')  # Get DPI from tkinter
            
            # Calculate diagonal size in inches
            diagonal_pixels = (self.screen_width**2 + self.screen_height**2)**0.5
            diagonal_inches = diagonal_pixels / dpi
            
            # Reference: 24-inch monitor at 1920x1080 (typical desktop)
            # This gives us a baseline for "normal" screen size
            reference_diagonal = 24.0  # inches
            reference_dpi = 96  # Standard DPI
            
            # Calculate scale based on screen size relative to reference
            # Smaller screens get larger scale (bigger elements)
            # Larger screens get smaller scale (smaller elements)
            size_scale = reference_diagonal / diagonal_inches
            
            # Also consider DPI - higher DPI screens can have smaller elements
            dpi_scale = reference_dpi / dpi
            
            # Combine both factors with reduced impact
            scale = (size_scale * 0.3) + (dpi_scale * 0.2) + 0.5
            
            # Clamp scale to more conservative bounds (0.6 to 1.4)
            scale = max(0.6, min(1.4, scale))
            
            return scale
            
        except Exception as e:
            # Fallback to resolution-based scaling if DPI detection fails
            print(f"DPI detection failed, using fallback: {e}")
            base_width = 1920  # Fallback reference
            scale = self.screen_width / base_width
            return max(0.7, min(1.3, scale))
    
    def scale_size(self, size):
        """Scale a size value by the current scale factor"""
        return int(size * self.scale_factor)
    
    def scale_font(self, base_size, weight='normal'):
        """Get scaled font size based on screen width"""
        scaled_size = int(base_size * self.scale_factor)
        
        if weight == 'bold':
            return ('Arial', scaled_size, 'bold')
        else:
            return ('Arial', scaled_size)
    
    def setup_ui(self):
        """Setup the main user interface with improved layout"""
        # Get current theme colors
        theme = self.themes[self.current_theme]
        
        # Set window background to theme color
        self.root.configure(bg=theme["bg_main"])
        
        # Main container with theme background
        main_frame = tk.Frame(self.root, bg=theme["bg_main"])
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Header with theme colors
        header_frame = tk.Frame(main_frame, bg=theme["bg_header"], height=self.scale_size(60))
        header_frame.pack(fill=tk.X, padx=0, pady=0)
        header_frame.pack_propagate(False)
        
        # Header content
        header_content = tk.Frame(header_frame, bg=theme["bg_header"])
        header_content.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # Time display - show total runtime
        self.time_label = tk.Label(header_content, text="Duration: 00:00:00", bg=theme["bg_header"], fg=theme["text_header"], 
                             font=self.scale_font(12, 'bold'))
        self.time_label.pack(side=tk.LEFT, padx=self.scale_size(10))
        
        # Tip mode selector
        tip_mode_frame = tk.Frame(header_content, bg=theme["bg_header"])
        tip_mode_frame.pack(side=tk.LEFT, padx=(self.scale_size(20), 0))
        
        tk.Label(tip_mode_frame, text="Tip Types:", bg=theme["bg_header"], fg=theme["text_header"], 
                font=self.scale_font(12, 'bold')).pack(side=tk.LEFT, padx=(0, self.scale_size(5)))
        
        self.tip_mode_var = tk.StringVar(value="Film Tips")
        tip_mode_menu = tk.OptionMenu(tip_mode_frame, self.tip_mode_var, "Film Tips", "Edit Tips",
                                      command=self.update_tip_mode)
        tip_mode_menu.config(bg=theme["bg_header"], fg=theme["text_header"], font=self.scale_font(12), 
                            relief=tk.FLAT, bd=0, highlightthickness=0)
        tip_mode_menu.pack(side=tk.LEFT)
        
        # Header buttons
        button_frame = tk.Frame(header_content, bg=theme["bg_header"])
        button_frame.pack(side=tk.RIGHT)
        
        tk.Button(button_frame, text="Project Properties", bg=theme["bg_header"], fg=theme["text_header"], 
                 font=self.scale_font(12), relief=tk.FLAT, command=self.show_project_properties).pack(side=tk.LEFT, padx=self.scale_size(2))
        
        tk.Button(button_frame, text="Add New Clip", bg=theme["bg_header"], fg=theme["text_header"], 
                 font=self.scale_font(12), relief=tk.FLAT, command=self.create_new_scene).pack(side=tk.LEFT, padx=self.scale_size(2))
        
        
        # Main content area
        content_frame = tk.Frame(main_frame, bg=theme["bg_main"])
        content_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Create scrollable frame for storyboard segments
        self.create_storyboard_area(content_frame)
        
        # Status bar
        self.status_var = tk.StringVar()
        self.status_var.set("Ready")
        status_frame = tk.Frame(main_frame, bg=theme["bg_status"], height=self.scale_size(25))
        status_frame.pack(side=tk.BOTTOM, fill=tk.X)
        status_frame.pack_propagate(False)
        
        # Status bar content
        status_content = tk.Frame(status_frame, bg=theme["bg_status"])
        status_content.pack(fill=tk.BOTH, expand=True, padx=5)
        
        # Status bar is now just for status messages
    
    def create_storyboard_area(self, parent):
        """Create the main storyboard area with scrollable content"""
        # Get current theme colors
        theme = self.themes[self.current_theme]
        
        # Create canvas and scrollbar for scrollable area
        canvas = tk.Canvas(parent, bg=theme["bg_main"], highlightthickness=0)
        scrollbar = ttk.Scrollbar(parent, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg=theme["bg_main"])
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Bind canvas resize to update scrollable frame width
        def on_canvas_configure(event):
            canvas_width = event.width
            canvas.itemconfig(canvas.find_all()[0], width=canvas_width)
        canvas.bind('<Configure>', on_canvas_configure)
        
        # Update scrollable frame width when canvas is created
        def update_frame_width():
            canvas_width = canvas.winfo_width()
            if canvas_width > 1:  # Make sure canvas has been rendered
                canvas.itemconfig(canvas.find_all()[0], width=canvas_width)
        canvas.after(100, update_frame_width)
        
        # Pack canvas and scrollbar
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Store references
        self.canvas = canvas
        self.scrollable_frame = scrollable_frame
        
        # Bind mousewheel to canvas - but only when not editing text widgets
        def _on_mousewheel(event):
            # Check if the event widget is a text widget that should handle its own scrolling
            widget = event.widget
            if isinstance(widget, (tk.Text, scrolledtext.ScrolledText)):
                # Let the text widget handle its own scrolling
                return
            
            # Check if we're inside a text widget
            current_widget = widget
            while current_widget:
                if isinstance(current_widget, (tk.Text, scrolledtext.ScrolledText)):
                    # We're inside a text widget, let it handle scrolling
                    return
                current_widget = current_widget.master
            
            # Only scroll the canvas if we're not in a text widget
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        
        canvas.bind_all("<MouseWheel>", _on_mousewheel)
    
    def create_scene_widget(self, scene, order_number):
        """Create a storyboard scene widget"""
        # Get current theme colors
        theme = self.themes[self.current_theme]
        
        # Main scene frame - make it span full width
        scene_frame = tk.Frame(self.scrollable_frame, bg=theme["bg_main"], relief=tk.RAISED, bd=1)
        scene_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Create a dedicated drag handle at the top
        drag_handle = tk.Frame(scene_frame, bg=theme["bg_drag"], height=self.scale_size(30))
        drag_handle.pack(fill=tk.X, padx=self.scale_size(2), pady=self.scale_size(2))
        drag_handle.pack_propagate(False)  # Maintain fixed height
        
        # Add drag handle label
        drag_label = tk.Label(drag_handle, text="⋮⋮ Drag to reorder", bg=theme["bg_drag"], 
                            font=self.scale_font(10), fg=theme["text_drag"])
        drag_label.pack(side=tk.LEFT, padx=self.scale_size(5), pady=self.scale_size(2))
        
        # Insert buttons at the top
        top_insert_btn = tk.Button(drag_handle, text="↑ Insert", bg=theme["bg_drag"], 
                                 font=self.scale_font(9, 'bold'), fg=theme["text_drag"], relief=tk.RAISED,
                                 command=lambda: self.insert_scene_before(scene))
        top_insert_btn.pack(side=tk.RIGHT, padx=self.scale_size(3), pady=self.scale_size(3))
        
        bottom_insert_btn = tk.Button(drag_handle, text="↓ Insert", bg=theme["bg_drag"], 
                                     font=self.scale_font(9, 'bold'), fg=theme["text_drag"], relief=tk.RAISED,
                                     command=lambda: self.insert_scene_after(scene))
        bottom_insert_btn.pack(side=tk.RIGHT, padx=self.scale_size(3), pady=self.scale_size(3))
        
        # Bind drag events only to the drag handle
        self.bind_drag_to_widget(drag_handle, scene)
        
        # Top row: Order (far left) + Length and Clip Types (aligned with content area)
        top_row_frame = tk.Frame(scene_frame, bg=theme["bg_main"])
        top_row_frame.pack(fill=tk.X, padx=self.scale_size(5), pady=(self.scale_size(5), 0))
        
        # Order section on the far left
        order_frame = tk.Frame(top_row_frame, bg=theme["bg_main"])
        order_frame.pack(side=tk.LEFT)
        
        tk.Label(order_frame, text="Order:", bg=theme["bg_main"], fg=theme["text_main"], font=self.scale_font(10, 'bold')).pack(side=tk.LEFT)
        self.order_var = tk.StringVar(value=str(order_number))
        order_entry = tk.Entry(order_frame, textvariable=self.order_var, font=self.scale_font(10), width=self.scale_size(5), bg=theme["bg_main"], fg=theme["text_main"])
        order_entry.pack(side=tk.LEFT, padx=(self.scale_size(5), 0))
        order_entry.bind('<FocusOut>', lambda e: self.update_scene_order(scene, order_entry))
        order_entry.bind('<Return>', lambda e: self.update_scene_order(scene, order_entry))
        
        # Calculate image width for spacer calculation
        image_width = self.scale_size(250)
        
        # Estimate Order width (label "Order:" + entry width + padding)
        # "Order:" label is approximately 50 pixels, entry is 5 chars * ~8 pixels = 40, padding = 5
        estimated_order_width = self.scale_size(95)  # Approximate width of Order section
        
        # Spacer to align Length label with Clip Title label below
        # Both top_row_frame and main_content_frame have same left padding (5)
        # Clip Title starts at: image_frame left pad + image width + image_frame right pad + spacing
        # image_frame has padx=5 (both sides), left_frame has right padx=5 for spacing
        image_frame_left_pad = self.scale_size(5)  # image_frame left padding
        image_frame_right_pad = self.scale_size(5)  # image_frame right padding (from padx=5)
        spacing_to_content = self.scale_size(5)  # left_frame right padding (spacing between image and content)
        total_distance_to_clip_title = image_frame_left_pad + image_width + image_frame_right_pad + spacing_to_content
        
        # Spacer should be total distance minus Order width (since Order is already taking that space)
        spacer_width = total_distance_to_clip_title - estimated_order_width
        
        spacer = tk.Frame(top_row_frame, bg=theme["bg_main"], width=spacer_width)
        spacer.pack(side=tk.LEFT)
        spacer.pack_propagate(False)
        
        # Length and Clip Types section (will align with content area below)
        length_type_frame = tk.Frame(top_row_frame, bg=theme["bg_main"])
        length_type_frame.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # Length section - label should align with Clip Title label
        length_frame = tk.Frame(length_type_frame, bg=theme["bg_main"])
        length_frame.pack(side=tk.LEFT)
        
        tk.Label(length_frame, text="  Length (hh:mm:ss):", bg=theme["bg_main"], fg=theme["text_main"], font=self.scale_font(10, 'bold')).pack(side=tk.LEFT)
        # Initialize with hh:mm:ss format
        initial_length = self.seconds_to_hms(scene.length) if hasattr(scene, 'length') else "00:00:00"
        self.length_var = tk.StringVar(value=initial_length)
        length_entry = tk.Entry(length_frame, textvariable=self.length_var, font=self.scale_font(10), width=self.scale_size(12), bg=theme["bg_main"], fg=theme["text_main"])
        length_entry.pack(side=tk.LEFT, padx=(self.scale_size(5), 0))
        length_entry.bind('<Return>', lambda e: self.finish_editing_duration(scene, e))
        length_entry.bind('<FocusOut>', lambda e: self.finish_editing_duration(scene, e))
        
        # Clip Types section on the right - all horizontal
        clip_types_frame = tk.Frame(length_type_frame, bg=theme["bg_main"])
        clip_types_frame.pack(side=tk.LEFT, padx=(self.scale_size(20), 0))
        
        tk.Label(clip_types_frame, text="Clip Types:", bg=theme["bg_main"], fg=theme["text_main"], font=self.scale_font(10, 'bold')).pack(side=tk.LEFT)
        
        # Create dictionary to store checkbox variables and widgets
        self.clip_type_vars = {}
        self.clip_type_checkboxes = {}
        clip_types = ["Video", "Still", "Audio", "Title"]
        
        for clip_type in clip_types:
            # Create IntVar for each checkbox
            self.clip_type_vars[clip_type] = tk.IntVar()
            # Set value based on whether this type is selected
            self.clip_type_vars[clip_type].set(1 if clip_type in scene.clip_types else 0)
            
            # Create a container frame for checkbox + label
            checkbox_container = tk.Frame(clip_types_frame, bg=theme["bg_main"])
            checkbox_container.pack(side=tk.LEFT, padx=(self.scale_size(5), 0))
            
            # Checkbox square and checkmark use fixed colors (unaffected by theme)
            # Create checkbox WITHOUT text so we can control checkmark and text separately
            selectcolor = "#FFFFFF"  # Always white background for checkbox square
            checkbox_fg = "#000000"  # Always black for checkmark only
            
            # Create checkbox without text label
            cb = tk.Checkbutton(checkbox_container, text="", variable=self.clip_type_vars[clip_type],
                               bg=theme["bg_main"], fg=checkbox_fg,
                               selectcolor=selectcolor, activebackground=theme["bg_main"],
                               activeforeground=checkbox_fg,
                               command=lambda s=scene: self.update_clip_types(s))
            
            # Explicitly configure selectcolor and fg again after creation to ensure they stick
            try:
                cb.configure(selectcolor="#FFFFFF", fg="#000000", activeforeground="#000000")
            except:
                pass
            
            cb.pack(side=tk.LEFT)
            
            # Create separate label for text with theme-aware color
            label = tk.Label(checkbox_container, text=clip_type, bg=theme["bg_main"], 
                           fg=theme["text_main"], font=self.scale_font(10))
            label.pack(side=tk.LEFT, padx=(self.scale_size(2), 0))
            
            # Store both checkbox and label
            self.clip_type_checkboxes[clip_type] = {'checkbox': cb, 'label': label}
        
        # Main content area below top row: Image (left) and Content (right)
        main_content_frame = tk.Frame(scene_frame, bg=theme["bg_main"])
        main_content_frame.pack(fill=tk.BOTH, expand=True, padx=self.scale_size(5), pady=self.scale_size(5))
        
        # Left side - Image area (no film strip)
        left_frame = tk.Frame(main_content_frame, bg=theme["bg_main"])
        left_frame.pack(side=tk.LEFT, fill=tk.Y, padx=(0, self.scale_size(5)))
        
        # Image placeholder with scaled sizing
        image_frame = tk.Frame(left_frame, bg=theme["bg_main"])
        image_frame.pack(side=tk.LEFT, padx=self.scale_size(5))
        
        # Image display - create a placeholder image
        placeholder_image = Image.new('RGB', (image_width, self.scale_size(180)), 'white')
        placeholder_photo = ImageTk.PhotoImage(placeholder_image)
        
        image_label = tk.Label(image_frame, image=placeholder_photo, bg='white',
                              relief=tk.SUNKEN, bd=1)
        image_label.pack()
        image_label.image = placeholder_photo  # Keep reference
        
        # Image load button
        load_btn = tk.Button(image_frame, text="Load Image", bg=theme["bg_main"], fg=theme["text_main"], font=self.scale_font(10),
                           command=lambda: self.load_scene_image(scene, image_label, image_width, self.scale_size(180)))
        load_btn.pack(pady=self.scale_size(2))
        
        # Load existing image if scene has one
        if hasattr(scene, 'image_path') and scene.image_path and os.path.exists(scene.image_path):
            try:
                self.load_image_display(scene.image_path, image_label, image_width, self.scale_size(180))
            except Exception as e:
                print(f"Failed to reload image for scene {scene.title}: {e}")
                # Keep the placeholder if image fails to reload
        
        # Store image label reference and dimensions
        scene.image_widget = image_label
        scene.image_width = image_width
        scene.image_height = self.scale_size(180)
        
        # Linked files display below the image
        links_frame = tk.Frame(image_frame, bg=theme["bg_main"])
        links_frame.pack(fill=tk.X, pady=(self.scale_size(5), 0))
        
        linked_files = []
        if hasattr(scene, 'audio_path') and scene.audio_path:
            linked_files.append("Audio")
        if hasattr(scene, 'video_path') and scene.video_path:
            linked_files.append("Video")
        
        if linked_files:
            links_text = f"Linked: {', '.join(linked_files)}"
            links_color = "#4A90E2"  # Blue for linked files
        else:
            links_text = "No files linked"
            links_color = "#999999"  # Gray for no files
        
        self.links_label = tk.Label(links_frame, text=links_text, bg=theme["bg_main"], 
                                   fg=links_color, font=self.scale_font(9, 'italic'))
        self.links_label.pack(anchor=tk.W)
        
        # Right side - Content and Buttons container
        right_container = tk.Frame(main_content_frame, bg=theme["bg_main"])
        right_container.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, self.scale_size(5)))
        
        # Content area (left side of right container)
        right_frame = tk.Frame(right_container, bg=theme["bg_main"])
        right_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Clip title
        title_frame = tk.Frame(right_frame, bg=theme["bg_main"])
        title_frame.pack(fill=tk.X, pady=(0, self.scale_size(5)))
        
        tk.Label(title_frame, text="Clip Title:", bg=theme["bg_main"], fg=theme["text_main"], font=self.scale_font(10, 'bold')).pack(anchor=tk.W)
        self.title_var = tk.StringVar(value=scene.title)
        title_entry = tk.Entry(title_frame, textvariable=self.title_var, font=self.scale_font(11), bg=theme["bg_main"], fg=theme["text_main"])
        title_entry.pack(fill=tk.X, pady=self.scale_size(2))
        title_entry.bind('<KeyRelease>', lambda e: self.update_scene_title(scene))
        
        # Description
        desc_frame = tk.Frame(right_frame, bg=theme["bg_main"])
        desc_frame.pack(fill=tk.X, pady=(0, self.scale_size(5)))
        
        tk.Label(desc_frame, text="Description:", bg=theme["bg_main"], fg=theme["text_main"], font=self.scale_font(10, 'bold')).pack(anchor=tk.W)
        self.desc_text = scrolledtext.ScrolledText(desc_frame, height=self.scale_size(3), font=self.scale_font(11), bg=theme["bg_main"], fg=theme["text_main"])
        self.desc_text.pack(fill=tk.X, pady=self.scale_size(2))
        self.desc_text.insert(1.0, scene.description)
        self.desc_text.bind('<KeyRelease>', lambda e: self.update_scene_description(scene))
        
        # Tips section - responsive layout based on screen size
        tips_frame = tk.Frame(right_frame, bg=theme["bg_main"])
        tips_frame.pack(fill=tk.BOTH, expand=True)
        
        # Use single column for small screens (like your laptop 1920x1080)
        if self.screen_width < 2000:  # Small screen - use single column
            # Film Tip
            film_tip_frame = tk.Frame(tips_frame, bg=theme["bg_main"])
            film_tip_frame.pack(fill=tk.BOTH, expand=True, pady=(0, self.scale_size(5)))
            
            tk.Label(film_tip_frame, text="Film Tip:", bg=theme["bg_main"], fg=theme["text_main"], font=self.scale_font(10, 'bold')).pack(anchor=tk.W)
            self.film_tip_text = scrolledtext.ScrolledText(film_tip_frame, height=self.scale_size(3), font=self.scale_font(11), bg=theme["bg_main"], fg=theme["text_main"])
            self.film_tip_text.pack(fill=tk.BOTH, pady=self.scale_size(2))
            self.film_tip_text.insert(1.0, scene.film_tip)
            self.film_tip_text.bind('<KeyRelease>', lambda e: self.update_film_tip(scene))
            
            # Edit Tip
            edit_tip_frame = tk.Frame(tips_frame, bg=theme["bg_main"])
            edit_tip_frame.pack(fill=tk.BOTH, expand=True)
            
            tk.Label(edit_tip_frame, text="Edit Tip:", bg=theme["bg_main"], fg=theme["text_main"], font=self.scale_font(10, 'bold')).pack(anchor=tk.W)
            self.edit_tip_text = scrolledtext.ScrolledText(edit_tip_frame, height=self.scale_size(3), font=self.scale_font(11), bg=theme["bg_main"], fg=theme["text_main"])
            self.edit_tip_text.pack(fill=tk.BOTH, pady=self.scale_size(2))
            self.edit_tip_text.insert(1.0, scene.edit_tip)
            self.edit_tip_text.bind('<KeyRelease>', lambda e: self.update_edit_tip(scene))
        else:  # Large screen - use two columns
            # Film Tip
            film_tip_frame = tk.Frame(tips_frame, bg=theme["bg_main"])
            film_tip_frame.grid(row=0, column=0, sticky='nsew', padx=(0, self.scale_size(5)))
            
            tk.Label(film_tip_frame, text="Film Tip:", bg=theme["bg_main"], fg=theme["text_main"], font=self.scale_font(10, 'bold')).pack(anchor=tk.W)
            self.film_tip_text = scrolledtext.ScrolledText(film_tip_frame, height=self.scale_size(4), font=self.scale_font(11), bg=theme["bg_main"], fg=theme["text_main"])
            self.film_tip_text.pack(fill=tk.BOTH, pady=self.scale_size(2))
            self.film_tip_text.insert(1.0, scene.film_tip)
            self.film_tip_text.bind('<KeyRelease>', lambda e: self.update_film_tip(scene))
            
            # Edit Tip
            edit_tip_frame = tk.Frame(tips_frame, bg=theme["bg_main"])
            edit_tip_frame.grid(row=0, column=1, sticky='nsew', padx=(self.scale_size(5), 0))
            
            tk.Label(edit_tip_frame, text="Edit Tip:", bg=theme["bg_main"], fg=theme["text_main"], font=self.scale_font(10, 'bold')).pack(anchor=tk.W)
            self.edit_tip_text = scrolledtext.ScrolledText(edit_tip_frame, height=self.scale_size(4), font=self.scale_font(11), bg=theme["bg_main"], fg=theme["text_main"])
            self.edit_tip_text.pack(fill=tk.BOTH, pady=self.scale_size(2))
            self.edit_tip_text.insert(1.0, scene.edit_tip)
            self.edit_tip_text.bind('<KeyRelease>', lambda e: self.update_edit_tip(scene))
            
            # Configure grid weights for proper expansion
            tips_frame.grid_columnconfigure(0, weight=1)
            tips_frame.grid_columnconfigure(1, weight=1)
        
        # Action buttons (right side of the right container)
        action_frame = tk.Frame(right_container, bg=theme["bg_main"])
        action_frame.pack(side=tk.RIGHT, fill=tk.Y, padx=self.scale_size(5), pady=self.scale_size(5))
        
        # Adjust button width based on screen size
        if self.screen_width < 2000:  # Small screen - smaller buttons
            button_width = self.scale_size(8)
        else:  # Large screen - normal buttons
            button_width = self.scale_size(12)
        
        # Configure grid weights for equal vertical space allocation (1/3 each)
        action_frame.grid_rowconfigure(0, weight=1)  # Tip button row
        action_frame.grid_rowconfigure(1, weight=1)  # Audio/Video buttons row  
        action_frame.grid_rowconfigure(2, weight=1)  # Delete button row
        action_frame.grid_columnconfigure(0, weight=1)  # Single column
        
        # Tip button (1/3 of vertical space)
        tip_button = tk.Button(action_frame, text="Tip", bg=theme["bg_main"], fg=theme["text_main"], font=self.scale_font(11, 'bold'),
                command=lambda: self.show_tip_dialog(scene))
        tip_button.grid(row=0, column=0, sticky='nsew', padx=self.scale_size(2), pady=self.scale_size(2))
        
        # Audio and Video buttons side by side (1/3 of vertical space)
        audio_video_frame = tk.Frame(action_frame, bg=theme["bg_main"])
        audio_video_frame.grid(row=1, column=0, sticky='nsew', padx=self.scale_size(2), pady=self.scale_size(2))
        audio_video_frame.grid_rowconfigure(0, weight=1)
        audio_video_frame.grid_columnconfigure(0, weight=1)
        audio_video_frame.grid_columnconfigure(1, weight=1)
        
        link_audio_button = tk.Button(audio_video_frame, text="Link Audio", bg=theme["bg_main"], fg=theme["text_main"], font=self.scale_font(11, 'bold'),
                command=lambda: self.link_audio_file(scene))
        link_audio_button.grid(row=0, column=0, sticky='nsew', padx=(0, 1))
        
        link_video_button = tk.Button(audio_video_frame, text="Link Video", bg=theme["bg_main"], fg=theme["text_main"], font=self.scale_font(11, 'bold'),
                command=lambda: self.link_video_file(scene))
        link_video_button.grid(row=0, column=1, sticky='nsew', padx=(1, 0))
        
        # Delete button (1/3 of vertical space)
        delete_button = tk.Button(action_frame, text="Delete", bg=theme["bg_main"], fg=theme["text_main"], font=self.scale_font(11, 'bold'),
                command=lambda: self.delete_scene(scene))
        delete_button.grid(row=2, column=0, sticky='nsew', padx=self.scale_size(2), pady=self.scale_size(2))
        
        # Store widget references
        scene.widgets = {
            'frame': scene_frame,
            'title_var': self.title_var,
            'desc_text': self.desc_text,
            'film_tip_text': self.film_tip_text,
            'edit_tip_text': self.edit_tip_text,
            'clip_type_vars': self.clip_type_vars,  # Updated for checkboxes
            'clip_type_checkboxes': self.clip_type_checkboxes,  # Checkbox widgets
            'length_var': self.length_var,
            'order_var': self.order_var,
            'links_label': self.links_label  # Store the links status label
        }
        
        # Update links label for existing linked files
        self.update_links_label(scene)
        
        # No drag binding on content areas - only on drag handle
        
        return scene_frame
    
    def setup_menu(self):
        """Setup the application menu"""
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        
        # File menu
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="New Project", command=self.new_project)
        file_menu.add_command(label="Open Project", command=self.open_project)
        file_menu.add_command(label="Save Project", command=self.save_project)
        file_menu.add_command(label="Save Project As...", command=self.save_project_as)
        file_menu.add_separator()
        file_menu.add_command(label="Export Timeline to PDF...", command=self.export_timeline_pdf)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.root.quit)
        
        # View menu
        view_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="View", menu=view_menu)
        view_menu.add_command(label="Refresh View", command=self.refresh_view)
        view_menu.add_separator()
        
        # Theme submenu
        theme_menu = tk.Menu(view_menu, tearoff=0)
        view_menu.add_cascade(label="Theme", menu=theme_menu)
        theme_menu.add_command(label="Default", command=lambda: self.apply_theme("Default"))
        theme_menu.add_command(label="Dark", command=lambda: self.apply_theme("Dark"))
        theme_menu.add_command(label="Ocean", command=lambda: self.apply_theme("Ocean"))
        theme_menu.add_command(label="Forest", command=lambda: self.apply_theme("Forest"))
        theme_menu.add_command(label="Sunset", command=lambda: self.apply_theme("Sunset"))
        
        # Help menu
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Help", menu=help_menu)
        help_menu.add_command(label="About", command=self.show_about)
    
    def create_new_scene(self):
        """Create a new scene"""
        scene_id = f"scene_{len(self.current_project.scenes) + 1}"
        scene = StoryboardScene(scene_id, f"Scene {len(self.current_project.scenes) + 1}")
        self.current_project.add_scene(scene)
        self.refresh_scene_display()
        self.status_var.set(f"Created new scene: {scene.title}")
    
    def refresh_scene_display(self):
        """Refresh the display of all scenes"""
        # Clear existing widgets
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()
        
        # Recreate all scene widgets
        for i, scene in enumerate(self.current_project.scenes):
            self.create_scene_widget(scene, i + 1)
        
        # Update total runtime
        self.update_total_runtime()
        
        # Update canvas scroll region
        self.canvas.update_idletasks()
        bbox = self.canvas.bbox("all")
        if bbox:
            # Check if content fits in canvas
            canvas_height = self.canvas.winfo_height()
            content_height = bbox[3] - bbox[1]  # bottom - top
            
            if content_height <= canvas_height:
                # Content fits, disable scrolling
                self.canvas.configure(scrollregion=(0, 0, 0, 0))
            else:
                # Content overflows, enable scrolling
                self.canvas.configure(scrollregion=bbox)
    
    def load_scene_image(self, scene, image_label, max_width, max_height):
        """Load an image for a scene with proper scaling"""
        file_path = filedialog.askopenfilename(
            title="Select Image",
            filetypes=[
                ("Image files", "*.png *.jpg *.jpeg *.gif *.bmp *.tiff"),
                ("All files", "*.*")
            ]
        )
        
        if file_path:
            scene.image_path = file_path
            self.load_image_display(file_path, image_label, max_width, max_height)
            self.status_var.set(f"Loaded image: {os.path.basename(file_path)}")
    
    def load_image_display(self, image_path: str, image_label, max_width, max_height):
        """Load and display an image scaled to fill the box exactly"""
        try:
            # Open image
            image = Image.open(image_path)
            
            # Resize image to fill the box exactly
            image = image.resize((max_width, max_height), Image.Resampling.LANCZOS)
            
            # Convert to PhotoImage
            photo = ImageTk.PhotoImage(image)
            
            # Update the label
            image_label.config(image=photo)
            image_label.image = photo  # Keep a reference
            
        except Exception as e:
            messagebox.showerror("Image Error", f"Could not load image: {str(e)}")
            # Create error placeholder
            error_image = Image.new('RGB', (max_width, max_height), 'red')
            error_photo = ImageTk.PhotoImage(error_image)
            image_label.config(image=error_photo)
            image_label.image = error_photo
    
    def update_scene_title(self, scene):
        """Update scene title"""
        scene.title = scene.widgets['title_var'].get()
        self.status_var.set(f"Updated scene: {scene.title}")
    
    def seconds_to_hms(self, seconds):
        """Convert seconds to hh:mm:ss format"""
        if seconds < 0:
            seconds = 0
        hours = int(seconds // 3600)
        minutes = int((seconds % 3600) // 60)
        secs = int(seconds % 60)
        return f"{hours:02d}:{minutes:02d}:{secs:02d}"
    
    def hms_to_seconds(self, hms_string):
        """Convert hh:mm:ss format to seconds"""
        try:
            parts = hms_string.strip().split(':')
            if len(parts) == 3:
                hours, minutes, seconds = parts
                total_seconds = int(hours) * 3600 + int(minutes) * 60 + int(seconds)
                return total_seconds
            elif len(parts) == 2:
                minutes, seconds = parts
                total_seconds = int(minutes) * 60 + int(seconds)
                return total_seconds
            else:
                # Try to parse as plain seconds
                return float(hms_string)
        except (ValueError, IndexError):
            return 0
    
    def update_scene_length(self, scene):
        """Update scene length and total runtime - only called when editing is complete"""
        try:
            hms_text = scene.widgets['length_var'].get().strip()
            if not hms_text:
                scene.length = 0
                scene.widgets['length_var'].set("00:00:00")
            else:
                # Convert hh:mm:ss to seconds
                length_seconds = self.hms_to_seconds(hms_text)
                scene.length = length_seconds
                
                # Update the display to show proper format only after editing is done
                formatted_time = self.seconds_to_hms(length_seconds)
                scene.widgets['length_var'].set(formatted_time)
            
            self.update_total_runtime()
            formatted_display = self.seconds_to_hms(scene.length)
            self.status_var.set(f"Updated scene length: {formatted_display}")
        except (ValueError, AttributeError):
            # Invalid format, reset to 0
            scene.length = 0
            scene.widgets['length_var'].set("00:00:00")
            self.update_total_runtime()
    
    def finish_editing_duration(self, scene, event=None):
        """Handle when user finishes editing duration (Enter key or focus out)"""
        if event is None:
            self.update_scene_length(scene)
        elif hasattr(event, 'type') and str(event.type) == '9':  # FocusOut event
            self.update_scene_length(scene)
        elif hasattr(event, 'keysym') and event.keysym == 'Return':  # Enter key
            self.update_scene_length(scene)
    
    def update_scene_order(self, scene, order_entry):
        """Update scene order and reorder scenes if needed"""
        try:
            new_order = int(scene.widgets['order_var'].get())
            if new_order < 1 or new_order > len(self.current_project.scenes):
                # Invalid order, revert to current position
                current_index = self.current_project.scenes.index(scene)
                scene.widgets['order_var'].set(str(current_index + 1))
                self.status_var.set("Invalid order number - reverted")
                return
            
            # Get current index
            current_index = self.current_project.scenes.index(scene)
            new_index = new_order - 1  # Convert to 0-based index
            
            if current_index != new_index:
                # Move scene to new position
                scene_to_move = self.current_project.scenes.pop(current_index)
                self.current_project.scenes.insert(new_index, scene_to_move)
                
                # Refresh display to show new order
                self.refresh_scene_display()
                self.status_var.set(f"Moved scene to position {new_order}")
            else:
                self.status_var.set("Scene already at that position")
                
        except ValueError:
            # Invalid number, revert to current position
            current_index = self.current_project.scenes.index(scene)
            scene.widgets['order_var'].set(str(current_index + 1))
            self.status_var.set("Invalid order number - reverted")
    
    def start_drag(self, event, scene):
        """Start dragging a scene"""
        self.drag_start_y = event.y_root
        self.dragged_scene = scene
        self.drag_start_index = self.current_project.scenes.index(scene)
        
        # Visual feedback - highlight the dragged scene
        scene.widgets['frame'].config(relief=tk.SUNKEN, bd=2)
    
    def drag_motion(self, event, scene):
        """Handle drag motion"""
        if not hasattr(self, 'dragged_scene') or self.dragged_scene != scene:
            return
        
        # Calculate which scene we're hovering over
        current_y = event.y_root
        delta_y = current_y - self.drag_start_y
        
        # Find the target scene based on mouse position
        target_scene = self.find_target_scene(event)
        if target_scene and target_scene != scene:
            # Determine which half we're hovering over
            target_frame = target_scene.widgets['frame']
            target_y = target_frame.winfo_rooty()
            target_height = target_frame.winfo_height()
            target_middle = target_y + (target_height // 2)
            
            # Visual feedback - different colors for top/bottom half
            if event.y_root <= target_middle:
                # Top half - will insert BEFORE
                target_scene.widgets['frame'].config(relief=tk.RAISED, bd=3, bg='#E6F3FF')
            else:
                # Bottom half - will insert AFTER
                target_scene.widgets['frame'].config(relief=tk.RAISED, bd=3, bg='#FFF0E6')
    
    def end_drag(self, event, scene):
        """End dragging and reorder if needed"""
        if not hasattr(self, 'dragged_scene') or self.dragged_scene != scene:
            return
        
        # Reset visual feedback
        scene.widgets['frame'].config(relief=tk.RAISED, bd=1, bg='#E0E0E0')
        
        # Find target scene
        target_scene = self.find_target_scene(event)
        current_index = self.current_project.scenes.index(scene)
        
        if target_scene and target_scene != scene:
            # Determine if we're dragging to top or bottom half of target
            target_frame = target_scene.widgets['frame']
            target_y = target_frame.winfo_rooty()
            target_height = target_frame.winfo_height()
            target_middle = target_y + (target_height // 2)
            
            # Get target index
            target_index = self.current_project.scenes.index(target_scene)
            
            # Remove scene from current position
            scene_to_move = self.current_project.scenes.pop(current_index)
            
            # Adjust target index if we removed a scene before the target
            if current_index < target_index:
                target_index -= 1
            
            # If dragging to bottom half, insert after the target
            if event.y_root > target_middle:
                target_index += 1
            
            # Insert at new position
            self.current_project.scenes.insert(target_index, scene_to_move)
            
            # Refresh display and update order numbers
            self.refresh_scene_display()
            self.status_var.set(f"Moved scene to position {target_index + 1}")
            
        elif not target_scene:
            # No target scene found - check if we're dragging to the end
            # Get the last scene's position
            if self.current_project.scenes:
                last_scene = self.current_project.scenes[-1]
                if hasattr(last_scene, 'widgets') and 'frame' in last_scene.widgets:
                    last_frame = last_scene.widgets['frame']
                    last_y = last_frame.winfo_rooty() + last_frame.winfo_height()
                    
                    # If we're dragging below the last scene, move to end
                    if event.y_root > last_y:
                        # Remove scene from current position
                        scene_to_move = self.current_project.scenes.pop(current_index)
                        
                        # Add to end
                        self.current_project.scenes.append(scene_to_move)
                        
                        # Refresh display and update order numbers
                        self.refresh_scene_display()
                        self.status_var.set(f"Moved scene to position {len(self.current_project.scenes)}")
        
        # Clean up
        if hasattr(self, 'dragged_scene'):
            delattr(self, 'dragged_scene')
        if hasattr(self, 'drag_start_y'):
            delattr(self, 'drag_start_y')
        if hasattr(self, 'drag_start_index'):
            delattr(self, 'drag_start_index')
    
    def find_target_scene(self, event):
        """Find which scene the mouse is over"""
        # Get the widget under the mouse
        widget = event.widget.winfo_containing(event.x_root, event.y_root)
        
        # Walk up the widget hierarchy to find a scene frame
        while widget:
            if hasattr(widget, 'master') and widget.master:
                # Check if this widget belongs to a scene
                for scene in self.current_project.scenes:
                    if hasattr(scene, 'widgets') and 'frame' in scene.widgets:
                        if widget == scene.widgets['frame'] or widget.winfo_parent() == str(scene.widgets['frame'].winfo_id()):
                            return scene
                widget = widget.master
            else:
                break
        
        return None
    
    def bind_drag_to_widget(self, widget, scene):
        """Bind drag events to a widget and all its children"""
        # Bind to the widget itself
        widget.bind('<Button-1>', lambda e: self.start_drag(e, scene))
        widget.bind('<B1-Motion>', lambda e: self.drag_motion(e, scene))
        widget.bind('<ButtonRelease-1>', lambda e: self.end_drag(e, scene))
        widget.bind('<Enter>', lambda e: widget.config(cursor='hand2'))
        widget.bind('<Leave>', lambda e: widget.config(cursor=''))
        
        # Recursively bind to all child widgets
        for child in widget.winfo_children():
            if isinstance(child, (tk.Frame, tk.Label, tk.Entry, tk.Text, tk.Canvas)):
                self.bind_drag_to_widget(child, scene)
    
    def update_total_runtime(self):
        """Calculate and display total runtime"""
        total_seconds = sum(getattr(scene, 'length', 0) for scene in self.current_project.scenes)
        
        hours = int(total_seconds // 3600)
        minutes = int((total_seconds % 3600) // 60)
        seconds = int(total_seconds % 60)
        
        runtime_text = f"{hours:02d}:{minutes:02d}:{seconds:02d}"
        self.time_label.config(text="Duration: " + runtime_text)
    
    def update_tip_mode(self, selected_mode):
        """Update tip mode (Film Tips or Edit Tips)"""
        self.status_var.set(f"Tip mode changed to: {selected_mode}")
        # Store the current tip mode for use by tip functionality
        self.current_tip_mode = selected_mode
    
    def update_scene_description(self, scene):
        """Update scene description"""
        scene.description = scene.widgets['desc_text'].get(1.0, tk.END).strip()
    
    def update_film_tip(self, scene):
        """Update film tip"""
        scene.film_tip = scene.widgets['film_tip_text'].get(1.0, tk.END).strip()
    
    def update_edit_tip(self, scene):
        """Update edit tip"""
        scene.edit_tip = scene.widgets['edit_tip_text'].get(1.0, tk.END).strip()
    
    def update_clip_types(self, scene):
        """Update clip types from checkboxes"""
        # Get all selected clip types from checkboxes
        selected_types = []
        if 'clip_type_vars' in scene.widgets:
            for clip_type, var in scene.widgets['clip_type_vars'].items():
                if var.get() == 1:  # Checkbox is checked
                    selected_types.append(clip_type)
        
        # Update the scene's clip_types list - allow empty list if none selected
        scene.clip_types = selected_types  # Can be empty list if no checkboxes are selected
        
        # Update status
        if scene.clip_types:
            types_str = ", ".join(scene.clip_types)
            self.status_var.set(f"Updated clip types: {types_str}")
        else:
            self.status_var.set("No clip types selected")
    
    def load_film_tips(self):
        """Load film tips from CSV file"""
        # New structure: Topics, Sub-topics, Description, Tip Type
        # Tip Type is either "Film Tip" or "Edit Tip"
        tips = {}  # Structure: {topic: {subtopic: {'description': str, 'tip_type': str}}}
        csv_path = os.path.join(os.path.dirname(__file__), "AllFilmTips.csv")
        
        try:
            with open(csv_path, 'r', encoding='utf-8') as file:
                reader = csv.DictReader(file)
                
                for row in reader:
                    topic = row['Topics'].strip()
                    subtopic = row['Sub-topics'].strip()
                    description = row['Description'].strip()
                    tip_type = row['Tip Type'].strip()
                    
                    if not topic or not subtopic or not tip_type:
                        continue
                    
                    if topic not in tips:
                        tips[topic] = {}
                    
                    tips[topic][subtopic] = {
                        'description': description,
                        'tip_type': tip_type
                    }
                
        except FileNotFoundError:
            print(f"Warning: Could not find {csv_path}")
        except Exception as e:
            print(f"Error loading film tips: {e}")
            import traceback
            traceback.print_exc()
        
        return tips
    
    def show_tip_dialog(self, scene):
        """Show tip dialog for a scene"""
        # Create tip selection dialog
        tip_dialog = tk.Toplevel(self.root)
        tip_dialog.title("Select Tip")
        tip_dialog.geometry("500x400")
        tip_dialog.configure(bg='#E0E0E0')
        tip_dialog.transient(self.root)
        tip_dialog.grab_set()
        
        # Center the dialog
        tip_dialog.update_idletasks()
        x = (tip_dialog.winfo_screenwidth() // 2) - (500 // 2)
        y = (tip_dialog.winfo_screenheight() // 2) - (400 // 2)
        tip_dialog.geometry(f"500x400+{x}+{y}")
        
        # Main frame
        main_frame = tk.Frame(tip_dialog, bg='#E0E0E0')
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Title
        title_label = tk.Label(main_frame, text="Select Tip (Film or Edit)", 
                              font=('Arial', 14, 'bold'), bg='#E0E0E0')
        title_label.pack(pady=(0, 10))
        
        # Topic selection
        topic_frame = tk.Frame(main_frame, bg='#E0E0E0')
        topic_frame.pack(fill=tk.X, pady=5)
        
        tk.Label(topic_frame, text="Topic:", font=('Arial', 10, 'bold'), 
                bg='#E0E0E0').pack(anchor=tk.W)
        
        self.tip_topic_var = tk.StringVar()
        topic_combo = ttk.Combobox(topic_frame, textvariable=self.tip_topic_var, 
                                  font=('Arial', 9), state='readonly')
        topic_combo.pack(fill=tk.X, pady=2)
        
        # Subtopic selection
        subtopic_frame = tk.Frame(main_frame, bg='#E0E0E0')
        subtopic_frame.pack(fill=tk.X, pady=5)
        
        tk.Label(subtopic_frame, text="Subtopic:", font=('Arial', 10, 'bold'), 
                bg='#E0E0E0').pack(anchor=tk.W)
        
        self.tip_subtopic_var = tk.StringVar()
        subtopic_combo = ttk.Combobox(subtopic_frame, textvariable=self.tip_subtopic_var, 
                                     font=('Arial', 9), state='readonly')
        subtopic_combo.pack(fill=tk.X, pady=2)
        
        # Description preview
        desc_frame = tk.Frame(main_frame, bg='#E0E0E0')
        desc_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        tk.Label(desc_frame, text="Description:", font=('Arial', 10, 'bold'), 
                bg='#E0E0E0').pack(anchor=tk.W)
        
        desc_text = scrolledtext.ScrolledText(desc_frame, height=8, font=('Arial', 9))
        desc_text.pack(fill=tk.BOTH, expand=True, pady=2)
        
        # Buttons
        button_frame = tk.Frame(main_frame, bg='#E0E0E0')
        button_frame.pack(fill=tk.X, pady=10)
        
        def update_topics():
            """Update topic list - filtered by current tip mode (Film Tips or Edit Tips)"""
            current_mode = self.current_tip_mode  # "Film Tips" or "Edit Tips"
            target_tip_type = "Film Tip" if current_mode == "Film Tips" else "Edit Tip"
            
            # Filter topics that have at least one subtopic of the target tip type
            available_topics = []
            for topic in self.film_tips.keys():
                for subtopic, tip_data in self.film_tips[topic].items():
                    if tip_data.get('tip_type') == target_tip_type:
                        available_topics.append(topic)
                        break  # Found at least one matching subtopic, add topic and move on
            
            topic_combo['values'] = available_topics
            if available_topics:
                topic_combo.set(available_topics[0])
                update_subtopics()
        
        def update_subtopics():
            """Update subtopic list based on selected topic - filtered by tip type"""
            selected_topic = self.tip_topic_var.get()
            if not selected_topic or selected_topic not in self.film_tips:
                return
            
            current_mode = self.current_tip_mode  # "Film Tips" or "Edit Tips"
            target_tip_type = "Film Tip" if current_mode == "Film Tips" else "Edit Tip"
            
            # Filter subtopics by tip type
            available_subtopics = []
            for subtopic, tip_data in self.film_tips[selected_topic].items():
                if tip_data.get('tip_type') == target_tip_type:
                    available_subtopics.append(subtopic)
            
            subtopic_combo['values'] = available_subtopics
            if available_subtopics:
                subtopic_combo.set(available_subtopics[0])
                update_description()
        
        def update_description():
            """Update description based on selected subtopic"""
            selected_topic = self.tip_topic_var.get()
            selected_subtopic = self.tip_subtopic_var.get()
            
            if not selected_topic or not selected_subtopic:
                return
            
            if selected_topic not in self.film_tips:
                return
            
            if selected_subtopic not in self.film_tips[selected_topic]:
                return
            
            tip_data = self.film_tips[selected_topic][selected_subtopic]
            description = tip_data.get('description', '')
            tip_type = tip_data.get('tip_type', 'Film Tip')
            
            # Show tip type in description preview
            desc_text.delete(1.0, tk.END)
            desc_text.insert(1.0, f"Tip Type: {tip_type}\n\n{description}")
        
        def apply_tip():
            """Apply the selected tip to the scene"""
            selected_topic = self.tip_topic_var.get()
            selected_subtopic = self.tip_subtopic_var.get()
            
            if not selected_topic or not selected_subtopic:
                messagebox.showwarning("Warning", "Please select a tip first.")
                return
            
            if selected_topic not in self.film_tips:
                messagebox.showwarning("Warning", "Invalid topic selected.")
                return
            
            if selected_subtopic not in self.film_tips[selected_topic]:
                messagebox.showwarning("Warning", "Invalid subtopic selected.")
                return
            
            tip_data = self.film_tips[selected_topic][selected_subtopic]
            description = tip_data.get('description', '').strip()
            tip_type = tip_data.get('tip_type', 'Film Tip')
            
            if not description:
                messagebox.showwarning("Warning", "No description available for this tip.")
                return
            
            # Apply tip to the correct box based on Tip Type
            tip_text = f"{selected_subtopic}: {description}"
            
            if tip_type == "Film Tip":
                # Update the film tip text area
                if hasattr(scene, 'widgets') and 'film_tip_text' in scene.widgets:
                    scene.widgets['film_tip_text'].delete(1.0, tk.END)
                    scene.widgets['film_tip_text'].insert(1.0, tip_text)
                    scene.film_tip = tip_text
                    self.status_var.set(f"Applied film tip: {selected_subtopic}")
            elif tip_type == "Edit Tip":
                # Update the edit tip text area
                if hasattr(scene, 'widgets') and 'edit_tip_text' in scene.widgets:
                    scene.widgets['edit_tip_text'].delete(1.0, tk.END)
                    scene.widgets['edit_tip_text'].insert(1.0, tip_text)
                    scene.edit_tip = tip_text
                    self.status_var.set(f"Applied edit tip: {selected_subtopic}")
            else:
                messagebox.showwarning("Warning", f"Unknown tip type: {tip_type}")
                return
            
            tip_dialog.destroy()
        
        # Bind events
        topic_combo.bind('<<ComboboxSelected>>', lambda e: update_subtopics())
        subtopic_combo.bind('<<ComboboxSelected>>', lambda e: update_description())
        
        # Buttons
        tk.Button(button_frame, text="Apply Tip", bg='#4A90E2', fg='white', 
                 font=('Arial', 10, 'bold'), command=apply_tip).pack(side=tk.LEFT, padx=(0, 10))
        tk.Button(button_frame, text="Cancel", bg='#C0C0C0', fg='black', 
                 font=('Arial', 10), command=tip_dialog.destroy).pack(side=tk.LEFT)
        
        # Initialize the dialog
        update_topics()
    
    def update_links_label(self, scene):
        """Update the links status label for a scene"""
        if hasattr(scene, 'widgets') and 'links_label' in scene.widgets:
            linked_files = []
            if hasattr(scene, 'audio_path') and scene.audio_path:
                linked_files.append("Audio")
            if hasattr(scene, 'video_path') and scene.video_path:
                linked_files.append("Video")
            
            if linked_files:
                links_text = f"Linked: {', '.join(linked_files)}"
                links_color = "#4A90E2"  # Blue for linked files
            else:
                links_text = "No files linked"
                links_color = "#999999"  # Gray for no files
            
            scene.widgets['links_label'].config(text=links_text, fg=links_color)
    
    def link_audio_file(self, scene):
        """Link an audio file to the scene"""
        file_path = filedialog.askopenfilename(
            title="Select Audio File",
            filetypes=[
                ("Audio files", "*.mp3 *.wav *.ogg *.m4a *.aac *.flac"),
                ("All files", "*.*")
            ]
        )
        
        if file_path:
            scene.audio_path = file_path
            self.status_var.set(f"Linked audio: {os.path.basename(file_path)}")
            self.update_links_label(scene)
    
    def link_video_file(self, scene):
        """Link a video file to the scene"""
        file_path = filedialog.askopenfilename(
            title="Select Video File",
            filetypes=[
                ("Video files", "*.mp4 *.avi *.mov *.mkv *.wmv *.flv *.webm"),
                ("All files", "*.*")
            ]
        )
        
        if file_path:
            scene.video_path = file_path
            self.status_var.set(f"Linked video: {os.path.basename(file_path)}")
            self.update_links_label(scene)
    
    
    def insert_scene_before(self, scene):
        """Insert a new scene before the current one"""
        # Find the index of the current scene
        try:
            index = self.current_project.scenes.index(scene)
            new_scene = StoryboardScene(f"scene_{len(self.current_project.scenes) + 1}", "New Scene")
            self.current_project.scenes.insert(index, new_scene)
            self.refresh_scene_display()
            self.status_var.set("Inserted new scene before current")
        except ValueError:
            self.create_new_scene()
    
    def insert_scene_after(self, scene):
        """Insert a new scene after the current one"""
        # Find the index of the current scene
        try:
            index = self.current_project.scenes.index(scene)
            new_scene = StoryboardScene(f"scene_{len(self.current_project.scenes) + 1}", "New Scene")
            self.current_project.scenes.insert(index + 1, new_scene)
            self.refresh_scene_display()
            self.status_var.set("Inserted new scene after current")
        except ValueError:
            self.create_new_scene()
    
    def delete_scene(self, scene):
        """Delete a scene"""
        if messagebox.askyesno("Delete Scene", f"Are you sure you want to delete '{scene.title}'?"):
            self.current_project.remove_scene(scene.scene_id)
            self.refresh_scene_display()
            self.status_var.set("Scene deleted")
    
    def new_project(self):
        """Create a new project"""
        if messagebox.askyesno("New Project", "Create a new project? Unsaved changes will be lost."):
            self.current_project = StoryboardProject()
            
            # Show project properties dialog for new project
            self.show_project_properties()
            
            # Create initial scene for new projects
            initial_scene = StoryboardScene("scene_1", "Scene 1")
            self.current_project.add_scene(initial_scene)
            
            self.refresh_scene_display()
            self.status_var.set("New project created with initial scene")
    
    def open_project(self):
        """Open an existing project"""
        file_path = filedialog.askopenfilename(
            title="Open Project",
            filetypes=[
                ("JSON files", "*.json"),
                ("All files", "*.*")
            ]
        )
        
        if file_path:
            try:
                if file_path.endswith('.json'):
                    self.load_json_project(file_path)
                else:
                    messagebox.showerror("Error", "Unsupported file format")
                    return
                
                self.current_project.project_path = file_path
                self.refresh_scene_display()
                self.status_var.set(f"Opened project: {os.path.basename(file_path)}")
                
            except Exception as e:
                messagebox.showerror("Error", f"Could not open project: {str(e)}")
    
    def save_project(self):
        """Save the current project"""
        if self.current_project.project_path:
            self.save_project_to_path(self.current_project.project_path)
        else:
            self.save_project_as()
    
    def save_project_as(self):
        """Save the project with a new name"""
        file_path = filedialog.asksaveasfilename(
            title="Save Project As",
            defaultextension=".json",
            filetypes=[
                ("JSON files", "*.json"),
                ("All files", "*.*")
            ]
        )
        
        if file_path:
            self.save_project_to_path(file_path)
            self.current_project.project_path = file_path
            self.status_var.set(f"Project saved: {os.path.basename(file_path)}")
    
    def save_project_to_path(self, file_path: str):
        """Save project to a specific path"""
        try:
            # Save current theme to project before saving
            self.current_project.theme = self.current_theme
            
            if file_path.endswith('.json'):
                with open(file_path, 'w', encoding='utf-8') as f:
                    json.dump(self.current_project.to_dict(), f, indent=2, ensure_ascii=False)
            else:
                messagebox.showerror("Error", "Unsupported file format")
                return
            
            self.status_var.set(f"Project saved: {os.path.basename(file_path)}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Could not save project: {str(e)}")
    
    def load_json_project(self, file_path: str):
        """Load project from JSON file"""
        with open(file_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        self.current_project = StoryboardProject.from_dict(data)
        
        # Load and apply theme if saved in project
        if hasattr(self.current_project, 'theme') and self.current_project.theme:
            if self.current_project.theme in self.themes:
                self.apply_theme(self.current_project.theme)
    
    def export_timeline_pdf(self):
        """Export the entire timeline to a PDF on a standard 8.5x11 sheet"""
        if not REPORTLAB_AVAILABLE:
            messagebox.showerror("Error", "PDF export requires reportlab library. Please install it with: pip install reportlab")
            return
        
        if not self.current_project.scenes:
            messagebox.showwarning("Warning", "No scenes to export. Please add some scenes first.")
            return
        
        # Ask user for save location
        file_path = filedialog.asksaveasfilename(
            title="Export Timeline to PDF",
            defaultextension=".pdf",
            filetypes=[
                ("PDF files", "*.pdf"),
                ("All files", "*.*")
            ]
        )
        
        if not file_path:
            return
        
        try:
            self.create_pdf_timeline(file_path)
            self.status_var.set(f"Timeline exported to PDF: {os.path.basename(file_path)}")
            messagebox.showinfo("Success", f"Timeline exported successfully to:\n{file_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Could not export PDF: {str(e)}")
            print(f"PDF export error: {e}")
            import traceback
            traceback.print_exc()
    
    def create_pdf_timeline(self, file_path: str):
        """Create PDF with timeline elements arranged top to bottom"""
        doc = SimpleDocTemplate(file_path, pagesize=letter,
                              rightMargin=72, leftMargin=72,
                              topMargin=72, bottomMargin=18)
        
        # Container for the 'Flowable' objects
        elements = []
        
        # Use current theme (what's active in UI) - prioritize over saved theme
        pdf_theme_name = self.current_theme
        if pdf_theme_name not in self.themes:
            pdf_theme_name = "Default"
        
        # Check if Default theme - if so, use standard colors (no theming)
        use_theme = pdf_theme_name != "Default"
        
        if use_theme:
            pdf_theme = self.themes[pdf_theme_name]
            
            # Convert hex colors or color names to reportlab colors
            def hex_to_color(color_value):
                """Convert hex color string or color name to reportlab color"""
                # Handle color names
                color_map = {
                    "white": colors.white,
                    "black": colors.black,
                    "red": colors.red,
                    "green": colors.green,
                    "blue": colors.blue,
                }
                if color_value.lower() in color_map:
                    return color_map[color_value.lower()]
                
                # Handle hex colors
                if not color_value.startswith('#'):
                    color_value = '#' + color_value
                try:
                    return colors.HexColor(color_value)
                except:
                    # Fallback to black if color conversion fails
                    return colors.black
            
            theme_bg = hex_to_color(pdf_theme["bg_main"])
            theme_text = hex_to_color(pdf_theme["text_main"])
            theme_header = hex_to_color(pdf_theme["bg_header"])
            theme_header_text = hex_to_color(pdf_theme["text_header"])
        else:
            # Default theme - use standard colors (white background, black text)
            theme_bg = colors.white
            theme_text = colors.black
            theme_header = colors.white
            theme_header_text = colors.black
        
        # Determine text colors based on background
        # Creators and duration are on white background, so use black text
        # Title has header background, so use header text color
        # Table content has theme background, so use theme text color
        creators_text_color = colors.black  # Always black on white background
        duration_text_color = colors.black  # Always black on white background
        
        # Get styles
        styles = getSampleStyleSheet()
        if use_theme:
            title_style = ParagraphStyle('CustomTitle', parent=styles['Heading1'], alignment=TA_CENTER,
                                        textColor=theme_header_text, backColor=theme_header)
        else:
            # Default theme - no background color, just black text
            title_style = ParagraphStyle('CustomTitle', parent=styles['Heading1'], alignment=TA_CENTER,
                                        textColor=colors.black)
        creators_style = ParagraphStyle('Creators', parent=styles['Heading3'], alignment=TA_CENTER,
                                       textColor=creators_text_color)
        duration_style = ParagraphStyle('Duration', parent=styles['Heading3'], alignment=TA_CENTER,
                                       textColor=duration_text_color)
        heading_style = ParagraphStyle('CustomHeading', parent=styles['Heading2'],
                                       textColor=theme_text)
        normal_style = ParagraphStyle('Normal', parent=styles['Normal'], textColor=theme_text)
        
        # Add project header with custom title and creators
        # Use custom title if available, otherwise fall back to project name
        pdf_title = self.current_project.project_title if self.current_project.project_title else self.current_project.name
        elements.append(Paragraph(f"<b>{pdf_title}</b>", title_style))
        
        # Add creators if specified
        if self.current_project.creators:
            creators_text = ', '.join(self.current_project.creators)
            elements.append(Paragraph(f"<i>{creators_text}</i>", creators_style))
        
        # Add total duration below creators
        total_seconds = sum(scene.length for scene in self.current_project.scenes)
        total_duration = self.seconds_to_hms(total_seconds)
        elements.append(Paragraph(f"<b>Total Duration: {total_duration}</b>", duration_style))
        
        elements.append(Spacer(1, 20))
        
        # Add scenes in order
        for i, scene in enumerate(self.current_project.scenes, 1):
            # Image column (if available)
            image_path = scene.image_path if scene.image_path and os.path.exists(scene.image_path) else None
            
            if image_path:
                try:
                    # Use original image directly - much simpler and more reliable
                    img = RLImage(image_path, width=2*inch, height=1.5*inch)
                    image_cell = img
                except Exception as e:
                    print(f"Could not load image {image_path}: {e}")
                    image_cell = "<i>Image not available</i>"
            else:
                image_cell = "<i>No image</i>"
            
            # Linked files section (below image)
            linked_files_list = []
            if scene.audio_path:
                linked_files_list.append(f"<b>Audio:</b> {os.path.basename(scene.audio_path)}")
            if scene.video_path:
                linked_files_list.append(f"<b>Video:</b> {os.path.basename(scene.video_path)}")
            
            # Linked files section (below image) - only show if files are linked
            if linked_files_list:
                linked_files_text = "<br/>".join(linked_files_list)
                linked_files_cell = Paragraph(linked_files_text, normal_style)
            else:
                # No files linked - use empty string instead of "No files linked"
                linked_files_cell = ""
            
            # Text content - Length and Clip Types on same horizontal line (only show Clip Types if any are selected)
            # Get clip types - check if list exists and is not empty
            if hasattr(scene, 'clip_types') and scene.clip_types:
                clip_types_str = ", ".join(scene.clip_types)
            elif hasattr(scene, 'clip_type') and scene.clip_type:
                # Fallback to single clip_type for backward compatibility
                clip_types_str = scene.clip_type
            else:
                clip_types_str = None
            
            # Convert seconds to hh:mm:ss format for display
            length_display = self.seconds_to_hms(scene.length)
            
            # Create horizontal layout for Length and Clip Types (only include Clip Types if present)
            if clip_types_str:
                length_clip_types_text = f"<b>Length:</b> {length_display} &nbsp;&nbsp;&nbsp;&nbsp; <b>Clip Types:</b> {clip_types_str}"
            else:
                # No clip types selected - only show Length
                length_clip_types_text = f"<b>Length:</b> {length_display}"
            
            # Build scene info with spacing between elements (extra <br/> for new line spacing)
            scene_info = [
                f"<b>Clip Title:</b> {scene.title or '<i>No title</i>'}",
                length_clip_types_text,
                f"<b>Description:</b> {scene.description or '<i>No description</i>'}",
            ]
            
            if scene.film_tip:
                scene_info.append(f"<b>Film Tip:</b> {scene.film_tip}")
            
            if scene.edit_tip:
                scene_info.append(f"<b>Edit Tip:</b> {scene.edit_tip}")
            
            # Join with double <br/> to add spacing between each element
            scene_text = "<br/><br/>".join(scene_info)
            text_cell = Paragraph(scene_text, normal_style)
            
            # Create table conditionally based on whether files are linked
            # If files are linked: 2 rows (image+linked files on left, text on right)
            # If no files linked: 1 row (image on left, text on right) - allows proper image centering
            if linked_files_cell:
                # Row 1: Image (left) | Text content (right)
                # Row 2: Linked files (left) | (empty)
                scene_data = [
                    [image_cell, text_cell],
                    [linked_files_cell, ""]
                ]
                has_linked_files = True
            else:
                # Single row: Image (left) | Text content (right)
                scene_data = [
                    [image_cell, text_cell]
                ]
                has_linked_files = False
            
            # Create table for this scene with box border around the whole thing
            # Remove fixed row heights to allow content to determine size naturally
            scene_table = Table(scene_data, colWidths=[2.5*inch, 4.5*inch])
            
            # Build table style - conditionally apply theme colors
            table_style = [
                ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                # Center image both horizontally and vertically in its cell
                ('ALIGN', (0, 0), (0, 0), 'CENTER'),
                ('VALIGN', (0, 0), (0, 0), 'MIDDLE'),
                ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
                ('FONTSIZE', (0, 0), (-1, -1), 9),
            ]
            
            # Add linked files alignment only if linked files exist
            if has_linked_files:
                table_style.extend([
                    # Align linked files to top left
                    ('ALIGN', (0, 1), (0, 1), 'LEFT'),
                    ('VALIGN', (0, 1), (0, 1), 'TOP'),
                ])
            
            # Only apply theme background colors if not using Default theme
            if use_theme:
                table_style.extend([
                    ('BACKGROUND', (0, 0), (0, -1), theme_bg),  # Image column background
                    ('BACKGROUND', (1, 0), (1, -1), theme_bg),  # Text column background
                ])
            
            table_style.extend([
                # Increase padding to prevent text overlap
                ('BOTTOMPADDING', (0, 0), (-1, -1), 15),
                ('TOPPADDING', (0, 0), (-1, -1), 15),
                ('LEFTPADDING', (0, 0), (-1, -1), 10),
                ('RIGHTPADDING', (0, 0), (-1, -1), 10),
                # Set minimum row height for first row to ensure image cell has enough height
                # The row height will be determined by the taller of: image cell or text cell
                ('MINIMUMHEIGHT', (0, 0), (0, 0), 1.5*inch),
                # Ensure text cell can expand vertically and determines row height
                ('VALIGN', (1, 0), (1, 0), 'TOP'),
                # Make sure the image cell expands to match the row height for proper centering
                ('VALIGN', (0, 0), (0, 0), 'MIDDLE'),
                # Add single box border around the entire scene
                ('BOX', (0, 0), (-1, -1), 1, colors.black),
                # Add vertical line between image and text columns
                ('LINEBEFORE', (1, 0), (1, -1), 1, colors.black),
                # Completely remove internal grid lines (except the vertical line we just added)
                ('INNERGRID', (0, 0), (-1, -1), 0, colors.transparent),
            ])
            
            # Only remove line between rows if we have linked files (2 rows)
            if has_linked_files:
                table_style.append(('LINEBELOW', (0, 0), (-1, 0), 0, colors.transparent))
            
            scene_table.setStyle(TableStyle(table_style))
            
            elements.append(scene_table)
            elements.append(Spacer(1, 30))  # Increased space between scenes
        
        # Build PDF
        doc.build(elements)
    
    def refresh_view(self):
        """Refresh the view"""
        self.refresh_scene_display()
        self.status_var.set("View refreshed")
    
    def apply_theme(self, theme_name):
        """Apply a color theme to the application"""
        if theme_name not in self.themes:
            return
        
        self.current_theme = theme_name
        theme = self.themes[theme_name]
        
        # Update main window background
        self.root.configure(bg=theme["bg_main"])
        
        # Force update all widgets with comprehensive color mapping
        self.force_update_all_widgets(theme)
        
        # Force update header bar specifically
        self.force_update_header_bar(theme)
        
        # Force update status bar specifically
        self.force_update_status_bar(theme)
        
        # Force update canvas and scrollable frame backgrounds
        self.force_update_canvas_backgrounds(theme)
        
        # Refresh the scene display to apply new colors
        self.refresh_scene_display()
        self.status_var.set(f"Theme changed to: {theme_name}")
    
    def force_update_all_widgets(self, theme):
        """Force update all widgets with comprehensive theme application"""
        try:
            # Update all widgets recursively with aggressive color mapping
            for widget in self.root.winfo_children():
                self.force_update_widget(widget, theme)
        except Exception as e:
            print(f"Error updating widgets: {e}")
    
    def force_update_header_bar(self, theme):
        """Force update the header bar with theme colors"""
        try:
            # Find and update the time label
            if hasattr(self, 'time_label'):
                self.time_label.configure(bg=theme["bg_header"], fg=theme["text_header"])
            
            # Force update all header elements
            for widget in self.root.winfo_children():
                if isinstance(widget, tk.Frame):
                    self.force_update_header_widgets(widget, theme)
        except Exception as e:
            print(f"Error updating header: {e}")
    
    def force_update_status_bar(self, theme):
        """Force update the status bar with theme colors"""
        try:
            # Find and update status bar frames
            for widget in self.root.winfo_children():
                if isinstance(widget, tk.Frame):
                    self.force_update_status_widgets(widget, theme)
        except Exception as e:
            print(f"Error updating status bar: {e}")
    
    def force_update_status_widgets(self, widget, theme):
        """Force update status bar widgets specifically"""
        try:
            # Check if this is a status frame (has status background color)
            if hasattr(widget, 'cget'):
                try:
                    current_bg = widget.cget('bg')
                    if current_bg in ['#C0C0C0', '#333333', '#80BFFF', '#90EE90', '#FF9933']:
                        widget.configure(bg=theme["bg_status"])
                        # Update all children of status frames
                        for child in widget.winfo_children():
                            self.force_update_status_child(child, theme)
                except:
                    pass
            
            # Recursively check child frames
            for child in widget.winfo_children():
                if isinstance(child, tk.Frame):
                    self.force_update_status_widgets(child, theme)
        except:
            pass
    
    def force_update_status_child(self, widget, theme):
        """Force update status bar child widgets"""
        try:
            if isinstance(widget, tk.Frame):
                widget.configure(bg=theme["bg_status"])
                # Recursively update children
                for child in widget.winfo_children():
                    self.force_update_status_child(child, theme)
        except:
            pass
    
    def force_update_canvas_backgrounds(self, theme):
        """Force update canvas and scrollable frame backgrounds"""
        try:
            # Update canvas background
            if hasattr(self, 'canvas'):
                self.canvas.configure(bg=theme["bg_main"])
            
            # Update scrollable frame background
            if hasattr(self, 'scrollable_frame'):
                self.scrollable_frame.configure(bg=theme["bg_main"])
        except Exception as e:
            print(f"Error updating canvas backgrounds: {e}")
    
    def force_update_header_widgets(self, widget, theme):
        """Force update header widgets specifically"""
        try:
            # Check if this is a header frame (has blue background or similar)
            if hasattr(widget, 'cget'):
                try:
                    current_bg = widget.cget('bg')
                    if current_bg in ['#4A90E2', '#1A1A1A', '#0066CC', '#228B22', '#FF6600']:
                        widget.configure(bg=theme["bg_header"])
                        # Update all children of header frames
                        for child in widget.winfo_children():
                            self.force_update_header_child(child, theme)
                except:
                    pass
            
            # Recursively check child frames
            for child in widget.winfo_children():
                if isinstance(child, tk.Frame):
                    self.force_update_header_widgets(child, theme)
        except:
            pass
    
    def force_update_header_child(self, widget, theme):
        """Force update header child widgets"""
        try:
            if isinstance(widget, tk.Label):
                widget.configure(bg=theme["bg_header"], fg=theme["text_header"])
            elif isinstance(widget, tk.Button):
                widget.configure(bg=theme["bg_header"], fg=theme["text_header"])
            elif isinstance(widget, tk.OptionMenu):
                widget.config(bg=theme["bg_header"], fg=theme["text_header"])
                if hasattr(widget, 'menu'):
                    widget['menu'].config(bg=theme["bg_header"], fg=theme["text_header"])
            elif isinstance(widget, tk.Frame):
                widget.configure(bg=theme["bg_header"])
                # Recursively update children
                for child in widget.winfo_children():
                    self.force_update_header_child(child, theme)
        except:
            pass
    
    def force_update_widget(self, widget, theme):
        """Force update a single widget and all its children"""
        try:
            # Update frames
            if isinstance(widget, tk.Frame):
                if hasattr(widget, 'configure'):
                    try:
                        current_bg = widget.cget('bg')
                        # Map any existing color to theme colors
                        if current_bg in ['#E0E0E0', '#4A90E2', '#D0D0D0', '#C0C0C0', '#2C2C2C', '#1A1A1A', '#404040', '#333333']:
                            if current_bg in ['#E0E0E0', '#2C2C2C']:
                                widget.configure(bg=theme["bg_main"])
                            elif current_bg in ['#4A90E2', '#1A1A1A']:
                                widget.configure(bg=theme["bg_header"])
                            elif current_bg in ['#D0D0D0', '#404040']:
                                widget.configure(bg=theme["bg_drag"])
                            elif current_bg in ['#C0C0C0', '#333333']:
                                widget.configure(bg=theme["bg_status"])
                    except:
                        pass
            
            # Update labels
            elif isinstance(widget, tk.Label):
                if hasattr(widget, 'configure'):
                    try:
                        current_bg = widget.cget('bg')
                        if current_bg in ['#4A90E2', '#D0D0D0', '#E0E0E0', '#1A1A1A', '#404040', '#2C2C2C']:
                            if current_bg in ['#4A90E2', '#1A1A1A']:
                                widget.configure(bg=theme["bg_header"], fg=theme["text_header"])
                            elif current_bg in ['#D0D0D0', '#404040']:
                                widget.configure(bg=theme["bg_drag"], fg=theme["text_drag"])
                            elif current_bg in ['#E0E0E0', '#2C2C2C']:
                                widget.configure(bg=theme["bg_main"], fg=theme["text_main"])
                    except:
                        pass
            
            # Update buttons
            elif isinstance(widget, tk.Button):
                if hasattr(widget, 'configure'):
                    try:
                        current_bg = widget.cget('bg')
                        if current_bg in ['#D0D0D0', '#E0E0E0', '#4A90E2', '#404040', '#2C2C2C', '#1A1A1A']:
                            if current_bg in ['#D0D0D0', '#404040']:
                                widget.configure(bg=theme["bg_drag"], fg=theme["text_drag"])
                            elif current_bg in ['#E0E0E0', '#2C2C2C']:
                                widget.configure(bg=theme["bg_main"], fg=theme["text_main"])
                            elif current_bg in ['#4A90E2', '#1A1A1A']:
                                widget.configure(bg=theme["bg_header"], fg=theme["text_header"])
                    except:
                        pass
            
            # Update OptionMenus
            elif isinstance(widget, tk.OptionMenu):
                try:
                    widget.config(bg=theme["bg_header"], fg=theme["text_header"])
                    if hasattr(widget, 'menu'):
                        widget['menu'].config(bg=theme["bg_header"], fg=theme["text_header"])
                except:
                    pass
            
            # Update ScrolledText widgets
            elif hasattr(widget, 'tag_configure'):  # ScrolledText widgets
                try:
                    widget.configure(bg=theme["bg_main"], fg=theme["text_main"])
                except:
                    pass
            
            # Update Checkbuttons - preserve selectcolor (white) and fg (black) but update bg
            elif isinstance(widget, tk.Checkbutton):
                if hasattr(widget, 'configure'):
                    try:
                        # Preserve the selectcolor (white) and fg (black) while updating bg
                        # The checkmark color follows fg, so we must keep it black
                        widget.configure(bg=theme["bg_main"], fg="#000000",  # Always black for checkmark
                                       selectcolor="#FFFFFF",  # Always keep white
                                       activebackground=theme["bg_main"],
                                       activeforeground="#000000")  # Always black even when active
                    except:
                        pass
            
            # Recursively update all child widgets
            for child in widget.winfo_children():
                self.force_update_widget(child, theme)
                
        except Exception as e:
            # Skip widgets that can't be updated
            pass
    
    def update_widget_colors(self, widget, theme):
        """Recursively update widget colors based on theme"""
        try:
            # Update frame backgrounds
            if isinstance(widget, tk.Frame):
                if hasattr(widget, 'cget') and 'bg' in widget.configure():
                    current_bg = widget.cget('bg')
                    # Map old colors to new theme colors
                    if current_bg in ['#E0E0E0', '#4A90E2', '#D0D0D0', '#C0C0C0']:
                        if current_bg == '#E0E0E0':
                            widget.configure(bg=theme["bg_main"])
                        elif current_bg == '#4A90E2':
                            widget.configure(bg=theme["bg_header"])
                        elif current_bg == '#D0D0D0':
                            widget.configure(bg=theme["bg_drag"])
                        elif current_bg == '#C0C0C0':
                            widget.configure(bg=theme["bg_status"])
            
            # Update label colors
            elif isinstance(widget, tk.Label):
                if hasattr(widget, 'cget') and 'bg' in widget.configure():
                    current_bg = widget.cget('bg')
                    if current_bg in ['#4A90E2', '#D0D0D0', '#E0E0E0']:
                        if current_bg == '#4A90E2':
                            widget.configure(bg=theme["bg_header"], fg=theme["text_header"])
                        elif current_bg == '#D0D0D0':
                            widget.configure(bg=theme["bg_drag"], fg=theme["text_drag"])
                        elif current_bg == '#E0E0E0':
                            widget.configure(bg=theme["bg_main"], fg=theme["text_main"])
            
            # Update button colors
            elif isinstance(widget, tk.Button):
                if hasattr(widget, 'cget') and 'bg' in widget.configure():
                    current_bg = widget.cget('bg')
                    if current_bg in ['#D0D0D0', '#E0E0E0', '#4A90E2']:
                        if current_bg == '#D0D0D0':
                            widget.configure(bg=theme["bg_drag"], fg=theme["text_drag"])
                        elif current_bg == '#E0E0E0':
                            widget.configure(bg=theme["bg_main"], fg=theme["text_main"])
                        elif current_bg == '#4A90E2':
                            widget.configure(bg=theme["bg_header"], fg=theme["text_header"])
            
            # Update OptionMenu colors
            elif isinstance(widget, tk.OptionMenu):
                if hasattr(widget, 'config'):
                    widget.config(bg=theme["bg_header"], fg=theme["text_header"])
            
            # Recursively update child widgets
            for child in widget.winfo_children():
                self.update_widget_colors(child, theme)
                
        except Exception as e:
            # Skip widgets that can't be updated
            pass
    
    def update_header_colors(self, theme):
        """Specifically update header bar colors and text"""
        try:
            # Find and update the time label
            if hasattr(self, 'time_label'):
                self.time_label.configure(bg=theme["bg_header"], fg=theme["text_header"])
            
            # Find and update the tip mode menu
            if hasattr(self, 'tip_mode_var'):
                # Find the OptionMenu widget and update its colors
                for widget in self.root.winfo_children():
                    self.update_option_menu_colors(widget, theme)
                    
        except Exception as e:
            # Skip if widgets don't exist yet
            pass
    
    def update_option_menu_colors(self, widget, theme):
        """Update OptionMenu colors and fonts"""
        try:
            if isinstance(widget, tk.OptionMenu):
                widget.config(bg=theme["bg_header"], fg=theme["text_header"], font=self.scale_font(12))
                # Update the menu button colors
                widget['menu'].config(bg=theme["bg_header"], fg=theme["text_header"])
            
            # Recursively check child widgets
            for child in widget.winfo_children():
                self.update_option_menu_colors(child, theme)
        except Exception as e:
            pass
    
    def show_project_properties(self):
        """Show project properties dialog for editing title and creators"""
        # Create project properties dialog
        prop_dialog = tk.Toplevel(self.root)
        prop_dialog.title("Project Properties")
        prop_dialog.geometry("600x480")
        prop_dialog.configure(bg='#E0E0E0')
        prop_dialog.transient(self.root)
        prop_dialog.grab_set()
        
        # Center the dialog
        prop_dialog.update_idletasks()
        x = (prop_dialog.winfo_screenwidth() // 2) - (600 // 2)
        y = (prop_dialog.winfo_screenheight() // 2) - (480 // 2)
        prop_dialog.geometry(f"600x480+{x}+{y}")
        
        # Main frame
        main_frame = tk.Frame(prop_dialog, bg='#E0E0E0')
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Title
        title_label = tk.Label(main_frame, text="Project Properties", 
                              font=('Arial', 14, 'bold'), bg='#E0E0E0')
        title_label.pack(pady=(0, 20))
        
        # Project Title
        title_frame = tk.Frame(main_frame, bg='#E0E0E0')
        title_frame.pack(fill=tk.X, pady=10)
        
        tk.Label(title_frame, text="PDF Title:", font=('Arial', 10, 'bold'), 
                bg='#E0E0E0').pack(anchor=tk.W)
        
        self.project_title_var = tk.StringVar(value=self.current_project.project_title)
        title_entry = tk.Entry(title_frame, textvariable=self.project_title_var, 
                              font=('Arial', 10), width=70)
        title_entry.pack(fill=tk.X, pady=5)
        
        # Creators
        creators_frame = tk.Frame(main_frame, bg='#E0E0E0')
        creators_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        tk.Label(creators_frame, text="Creators (one per line):", font=('Arial', 10, 'bold'), 
                bg='#E0E0E0').pack(anchor=tk.W)
        
        # Convert creators list to string format for display in text widget
        creators_text = '\n'.join(self.current_project.creators) if self.current_project.creators else ''
        
        creators_text_widget = tk.Text(creators_frame, height=6, font=('Arial', 10), width=70)
        creators_text_widget.insert('1.0', creators_text)
        creators_text_widget.pack(fill=tk.X, pady=5)
        
        # Help text
        help_label = tk.Label(creators_frame, 
                             text="(Enter one creator name per line. These will appear at the top of the PDF export)", 
                             font=('Arial', 8), fg='gray', bg='#E0E0E0')
        help_label.pack(anchor=tk.W)
        
        # Buttons
        button_frame = tk.Frame(main_frame, bg='#E0E0E0')
        button_frame.pack(fill=tk.X, pady=(20, 0))
        
        def save_properties():
            """Save the project properties"""
            title = self.project_title_var.get()
            creators_text = creators_text_widget.get('1.0', tk.END).strip()
            creators_list = [name.strip() for name in creators_text.split('\n') if name.strip()]
            
            # Update the project properties
            self.current_project.project_title = title
            self.current_project.creators = creators_list
            
            # Debug output
            print(f"Project title set to: '{title}'")
            print(f"Creators set to: {creators_list}")
            
            self.status_var.set(f"Project properties updated - Title: '{title}', Creators: {len(creators_list)}")
            prop_dialog.destroy()
        
        tk.Button(button_frame, text="Set", bg='#4A90E2', fg='white', 
                 font=('Arial', 10, 'bold'), command=save_properties).pack(side=tk.LEFT, padx=(0, 10))
        tk.Button(button_frame, text="Cancel", bg='#C0C0C0', fg='black', 
                 font=('Arial', 10), command=prop_dialog.destroy).pack(side=tk.LEFT)
        
        # Focus the first entry field
        title_entry.focus_set()

    def show_about(self):
        """Show about dialog"""
        about_text = """StoryBoard Amateur
Copyright (c) 2025 Johanner Corrales, Lauren Oquendo, Paulo Cao Suarez,  Danilo Bodden, Carlos Jerak

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), 
to deal with the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, 
sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the 
following conditions: 

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software. 

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, 
INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A 
PARTICULAR PURPOSE, AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR 
COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES, OR OTHER LIABILITY, WHETHER IN 
AN ACTION OF CONTRACT, TORT, OR OTHERWISE, ARISING FROM, OUT OF, OR IN CONNECTION 
WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
"""
        
        messagebox.showinfo("About StoryBoard Amateur", about_text)
    
    def run(self):
        """Run the application"""
        self.root.mainloop()


def main():
    """Main entry point"""
    app = StoryboardApp()
    app.run()


if __name__ == "__main__":
    main()
