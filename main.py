import tkinter as tk
from tkinter import ttk, Menu, messagebox, Toplevel
import os
import threading
import json
import numpy as np

from pynput_handler import PynputHandler
from pynput import keyboard, mouse

class MacroApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Advanced Macro Recorder")
        self.geometry("900x600")

        self.default_macro_folder = "macros_json"
        self.create_default_macro_folder()
        
        self.pynput_handler = PynputHandler()
        self.capture_listener = None
        self.is_editing = False

        self.protocol("WM_DELETE_WINDOW", self.on_app_close)
        self.create_menu()

        self.main_panes = ttk.PanedWindow(self, orient=tk.HORIZONTAL)
        self.main_panes.pack(fill=tk.BOTH, expand=True)
        self.left_frame = ttk.Frame(self.main_panes, width=600)
        self.main_panes.add(self.left_frame, weight=1)
        self.right_frame = ttk.Frame(self.main_panes, width=300)
        self.main_panes.add(self.right_frame, weight=0)
        
        self.create_left_panel_widgets()
        self.create_right_panel_widgets()

        self.status_label = tk.Label(self.left_frame, text="Ready.", anchor="w")
        self.status_label.pack(fill=tk.X, padx=10, pady=5)

        self.start_key, self.stop_key, self.play_key = None, None, None
        self.recording_delay = 0.1 # Default value
        
        self.load_settings()
        self.load_macros()

    def update_hotkeys(self):
        hotkeys_map = {}
        if self.start_key and self.start_key == self.stop_key:
            hotkeys_map[self.start_key] = self.toggle_recording
        else:
            if self.start_key: hotkeys_map[self.start_key] = self.start_recording
            if self.stop_key: hotkeys_map[self.stop_key] = self.stop_recording
        if self.play_key: hotkeys_map[self.play_key] = self.play_macro
        self.pynput_handler.set_hotkeys(hotkeys_map)

    def toggle_recording(self):
        if self.pynput_handler.is_recording:
            self.stop_recording()
        else:
            self.start_recording()
        
    def on_app_close(self):
        self.pynput_handler.stop_listeners()
        self.destroy()

    def create_menu(self):
        menubar = Menu(self); self.config(menu=menubar)
        file_menu = Menu(menubar, tearoff=False)
        menubar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="New Macro", command=self.new_macro)
        file_menu.add_separator()
        file_menu.add_command(label="Settings", command=self.show_settings)

    def create_default_macro_folder(self):
        if not os.path.exists(self.default_macro_folder):
            os.makedirs(self.default_macro_folder)

    def create_left_panel_widgets(self):
        control_frame = ttk.Frame(self.left_frame)
        control_frame.pack(fill=tk.X, padx=10, pady=(10, 5))
        self.record_button = tk.Button(control_frame, text="Record", command=self.start_recording)
        self.record_button.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=2)
        self.stop_button = tk.Button(control_frame, text="Stop", state=tk.DISABLED, command=self.stop_recording)
        self.stop_button.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=2)
        self.play_button = tk.Button(control_frame, text="Play", command=self.play_macro)
        self.play_button.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=2)
        self.action_tree = ttk.Treeview(self.left_frame, columns=("Action", "Value", "Duration"), show="headings")
        self.action_tree.heading("Action", text="Action Name"); self.action_tree.heading("Value", text="Value/Location"); self.action_tree.heading("Duration", text="Duration")
        self.action_tree.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        self.action_tree.bind("<Delete>", self.delete_treeview_item)

    def create_right_panel_widgets(self):
        macro_label = ttk.Label(self.right_frame, text="Macros", font=("Helvetica", 14))
        macro_label.pack(pady=(10, 5))
        self.macro_listbox = tk.Listbox(self.right_frame)
        self.macro_listbox.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        self.macro_listbox.bind("<<ListboxSelect>>", self.on_macro_select)
        self.macro_listbox.bind("<Double-Button-1>", self.rename_macro)
        self.macro_listbox.bind("<Button-3>", self.show_context_menu)
        button_frame = ttk.Frame(self.right_frame)
        button_frame.pack(fill=tk.X, padx=10, pady=(0, 10))
        self.add_button = tk.Button(button_frame, text="Add", command=self.new_macro)
        self.add_button.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(0, 5))
        self.delete_button = tk.Button(button_frame, text="Delete", command=self.confirm_delete)
        self.delete_button.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(5, 0))

    def show_context_menu(self, event):
        selection = self.macro_listbox.nearest(event.y)
        if selection != -1:
            self.macro_listbox.selection_clear(0, tk.END); self.macro_listbox.selection_set(selection)
        context_menu = Menu(self.macro_listbox, tearoff=False)
        context_menu.add_command(label="New Macro", command=self.new_macro)
        if selection != -1:
            context_menu.add_command(label="Rename", command=self.rename_macro)
            context_menu.add_command(label="Delete", command=self.confirm_delete)
        context_menu.tk_popup(event.x_root, event.y_root)

    def rename_macro(self, event=None):
        if self.is_editing: return
        selected_indices = self.macro_listbox.curselection()
        if not selected_indices: return
        self.editing_index = selected_indices[0]
        macro_name = self.macro_listbox.get(self.editing_index)
        self.is_editing = True
        bbox = self.macro_listbox.bbox(self.editing_index)
        if not bbox: return
        self.editing_entry = tk.Entry(self.macro_listbox, bd=0)
        self.editing_entry.insert(0, macro_name)
        self.editing_entry.place(x=bbox[0], y=bbox[1], width=bbox[2], height=bbox[3])
        self.editing_entry.focus_set()
        self.editing_entry.bind("<Return>", self.confirm_rename)
        self.editing_entry.bind("<FocusOut>", self.confirm_rename)

    def confirm_rename(self, event=None):
        if not hasattr(self, 'editing_entry') or not self.editing_entry: return
        new_name = self.editing_entry.get().strip()
        old_name = self.macro_listbox.get(self.editing_index)
        if new_name and new_name != old_name:
            self.macro_listbox.delete(self.editing_index)
            self.macro_listbox.insert(self.editing_index, new_name)
            old_path = os.path.join(self.default_macro_folder, f"{old_name}.json")
            new_path = os.path.join(self.default_macro_folder, f"{new_name}.json")
            if os.path.exists(new_path):
                messagebox.showwarning("Warning", f"A macro named '{new_name}' already exists.")
                self.macro_listbox.delete(self.editing_index)
                self.macro_listbox.insert(self.editing_index, old_name)
            elif os.path.exists(old_path): os.rename(old_path, new_path)
        self.editing_entry.destroy()
        self.is_editing = False

    def on_macro_select(self, event=None):
        selected_indices = self.macro_listbox.curselection()
        if not selected_indices:
            self.action_tree.delete(*self.action_tree.get_children())
            return
        macro_name = self.macro_listbox.get(selected_indices[0])
        file_path = os.path.join(self.default_macro_folder, f"{macro_name}.json")
        self.action_tree.delete(*self.action_tree.get_children())
        if not os.path.exists(file_path): return
        try:
            with open(file_path, "r") as f: events = json.load(f).get("actions", [])
            if not events: return
            i = 0
            while i < len(events):
                event = events[i]
                row_values = ()
                if i + 1 < len(events):
                    next_event = events[i+1]
                    if (event['type'] == 'key_press' and next_event['type'] == 'key_release' and event['data'][0] == next_event['data'][0]):
                        duration = next_event['time'] - event['time']
                        row_values = ("Key Click", event['data'][0].strip("'"), f"{duration:.3f}s")
                        i += 2; self.action_tree.insert("", "end", values=row_values); continue
                    if (event['type'] == 'mouse_click' and event['data'][3] and next_event['type'] == 'mouse_click' and not next_event['data'][3] and event['data'][2] == next_event['data'][2]):
                        duration = next_event['time'] - event['time']
                        x, y, button, _ = event['data']
                        row_values = ("Mouse Click", f"({x}, {y}) {button}", f"{duration:.3f}s")
                        i += 2; self.action_tree.insert("", "end", values=row_values); continue
                if event['type'] == 'mouse_move':
                    start_event = event
                    j = i + 1
                    while j < len(events) and events[j]['type'] == 'mouse_move': j += 1
                    end_event = events[j - 1]
                    row_values = ("Mouse Drag", f"From {start_event['data']} to {end_event['data']}", f"{end_event['time'] - start_event['time']:.2f}s")
                    i = j
                elif event['type'] == 'mouse_scroll':
                    start_event = event
                    j = i + 1
                    total_dx, total_dy = start_event['data'][2], start_event['data'][3]
                    start_dir = (np.sign(total_dx), np.sign(total_dy))
                    while j < len(events) and events[j]['type'] == 'mouse_scroll':
                        next_dx, next_dy = events[j]['data'][2], events[j]['data'][3]
                        if (np.sign(next_dx), np.sign(next_dy)) == start_dir:
                            total_dx += next_dx; total_dy += next_dy; j += 1
                        else: break
                    row_values = ("Mouse Scroll", f"dx={total_dx}, dy={total_dy}", f"{events[j-1]['time'] - start_event['time']:.2f}s")
                    i = j
                else:
                    duration = 0
                    if i > 0: duration = event['time'] - events[i-1]['time']
                    if event['type'] == 'key_press': row_values = ("Key Press", event['data'][0].strip("'"), f"{duration:.3f}s")
                    elif event['type'] == 'key_release': row_values = ("Key Release", event['data'][0].strip("'"), f"{duration:.3f}s")
                    elif event['type'] == 'mouse_click' and event['data'][3]: row_values = ("Mouse Press", f"({event['data'][0]}, {event['data'][1]}) {event['data'][2]}", f"{duration:.3f}s")
                    elif event['type'] == 'mouse_click' and not event['data'][3]: row_values = ("Mouse Release", f"({event['data'][0]}, {event['data'][1]}) {event['data'][2]}", f"{duration:.3f}s")
                    i += 1
                if row_values: self.action_tree.insert("", "end", values=row_values)
        except (json.JSONDecodeError, IOError) as e:
            messagebox.showerror("Error", f"Could not load macro file: {e}")
    
    def save_settings(self):
        def key_to_str(key):
            if key is None: return None
            if isinstance(key, keyboard.Key): return f"keyboard.Key.{key.name}"
            if isinstance(key, keyboard.KeyCode): return f"keyboard.KeyCode.from_char('{key.char}')"
            return str(key)
        
        self.recording_delay = self.delay_var.get()
        self.pynput_handler.set_recording_delay(self.recording_delay)
        
        settings = {
            "start_key": key_to_str(self.start_key), 
            "stop_key": key_to_str(self.stop_key), 
            "play_key": key_to_str(self.play_key),
            "recording_delay": self.recording_delay
        }
        try:
            with open("settings.json", "w") as f: json.dump(settings, f, indent=4)
        except IOError as e: messagebox.showerror("Error", f"Error saving settings: {e}")
        self.update_hotkeys()
        if hasattr(self, 'settings_window') and self.settings_window.winfo_exists():
            self.settings_window.destroy()

    def load_settings(self):
        try:
            with open("settings.json", "r") as f:
                settings = json.load(f)
                self.start_key = eval(settings.get("start_key")) if settings.get("start_key") else None
                self.stop_key = eval(settings.get("stop_key")) if settings.get("stop_key") else None
                self.play_key = eval(settings.get("play_key")) if settings.get("play_key") else None
                self.recording_delay = float(settings.get("recording_delay", 0.1))
        except (FileNotFoundError, json.JSONDecodeError, ValueError):
            self.start_key, self.stop_key, self.play_key = None, None, None
            self.recording_delay = 0.1
        finally:
            self.update_hotkeys()
            self.pynput_handler.set_recording_delay(self.recording_delay)
    
    def show_settings(self):
        self.settings_window = Toplevel(self); self.settings_window.title("Settings")
        self.settings_window.geometry("300x260"); self.settings_window.grab_set()

        def key_to_display_str(key):
            if key is None: return "None"
            return str(key).replace("Key.", "").replace("'", "")
            
        self.start_key_var = tk.StringVar(value=key_to_display_str(self.start_key))
        self.stop_key_var = tk.StringVar(value=key_to_display_str(self.stop_key))
        self.play_key_var = tk.StringVar(value=key_to_display_str(self.play_key))
        self.delay_var = tk.DoubleVar(value=self.recording_delay)

        tk.Label(self.settings_window, text="Start/Toggle Recording Key:").pack(pady=(10,0))
        start_label = tk.Label(self.settings_window, textvariable=self.start_key_var, relief="solid", bd=1, width=20); start_label.pack(pady=2)
        start_label.bind("<Button-1>", lambda e: self.start_key_capture("start"))
        
        tk.Label(self.settings_window, text="Stop Recording Key (if different):").pack()
        stop_label = tk.Label(self.settings_window, textvariable=self.stop_key_var, relief="solid", bd=1, width=20); stop_label.pack(pady=2)
        stop_label.bind("<Button-1>", lambda e: self.start_key_capture("stop"))
        
        tk.Label(self.settings_window, text="Play Macro Key:").pack()
        play_label = tk.Label(self.settings_window, textvariable=self.play_key_var, relief="solid", bd=1, width=20); play_label.pack(pady=2)
        play_label.bind("<Button-1>", lambda e: self.start_key_capture("play"))

        tk.Label(self.settings_window, text="Recording Start Delay (s):").pack(pady=(10,0))
        delay_entry = tk.Entry(self.settings_window, textvariable=self.delay_var, width=10, justify='center'); delay_entry.pack()

        save_button = tk.Button(self.settings_window, text="Save", command=self.save_settings, font=("Helvetica", 10, "bold")); save_button.pack(pady=15, ipadx=10, ipady=2)

    def start_key_capture(self, key_type):
        if self.capture_listener and self.capture_listener.is_alive(): return
        self.pynput_handler.pause()
        self.current_key_to_set = key_type
        capture_label = tk.Label(self.settings_window, text="Press any key...", fg="red", bg="#f0f0f0", font=("Helvetica", 12, "bold"))
        capture_label.pack(pady=5, fill="x", ipady=5)
        def on_press(key):
            if self.current_key_to_set == "start": self.start_key = key
            elif self.current_key_to_set == "stop": self.stop_key = key
            elif self.current_key_to_set == "play": self.play_key = key
            self.start_key_var.set(str(self.start_key).replace("Key.", "").replace("'", "")); self.stop_key_var.set(str(self.stop_key).replace("Key.", "").replace("'", "")); self.play_key_var.set(str(self.play_key).replace("Key.", "").replace("'", ""))
            capture_label.destroy()
            self.pynput_handler.resume()
            self.capture_listener = None
            return False
        self.capture_listener = keyboard.Listener(on_press=on_press)
        self.capture_listener.start()

    def delete_treeview_item(self, event=None):
        for item in self.action_tree.selection(): self.action_tree.delete(item)
    
    def load_macros(self):
        self.macro_listbox.delete(0, tk.END)
        for file_name in sorted(os.listdir(self.default_macro_folder)):
            if file_name.endswith('.json'): self.macro_listbox.insert(tk.END, os.path.splitext(file_name)[0])

    def new_macro(self):
        i = 1
        while True:
            new_name = f"macro{i}"
            if new_name not in self.macro_listbox.get(0, tk.END): break
            i += 1
        file_path = os.path.join(self.default_macro_folder, f"{new_name}.json")
        try:
            with open(file_path, "w") as f: json.dump({"actions": []}, f)
            self.macro_listbox.insert(tk.END, new_name)
            self.macro_listbox.selection_clear(0, tk.END)
            self.macro_listbox.selection_set(tk.END)
            self.on_macro_select(None)
        except IOError as e: messagebox.showerror("Error", f"Could not create new macro file: {e}")

    def confirm_delete(self):
        selected_indices = self.macro_listbox.curselection()
        if not selected_indices: return
        macro_name = self.macro_listbox.get(selected_indices[0])
        if messagebox.askyesno("Confirm Delete", f"Delete macro '{macro_name}'?"):
            file_path = os.path.join(self.default_macro_folder, f"{macro_name}.json")
            if os.path.exists(file_path):
                try:
                    os.remove(file_path); self.macro_listbox.delete(selected_indices[0]); self.action_tree.delete(*self.action_tree.get_children())
                except OSError as e: messagebox.showerror("Error", f"Error deleting file: {e}")

    def start_recording(self):
        if not self.macro_listbox.curselection():
            messagebox.showinfo("Recording Error", "Please select a macro to record into.")
            return
        self.action_tree.delete(*self.action_tree.get_children())
        self.pynput_handler.start_recording()
        self.record_button.config(state=tk.DISABLED); self.stop_button.config(state=tk.NORMAL); self.play_button.config(state=tk.DISABLED)
        self.status_label.config(text="Recording...")

    def stop_recording(self):
        recorded_events = self.pynput_handler.stop_recording()
        self.record_button.config(state=tk.NORMAL); self.stop_button.config(state=tk.DISABLED); self.play_button.config(state=tk.NORMAL)
        self.status_label.config(text="Recording stopped.")
        selected_index = self.macro_listbox.curselection()
        if selected_index:
            macro_name = self.macro_listbox.get(selected_index[0])
            file_path = os.path.join(self.default_macro_folder, f"{macro_name}.json")
            save_thread = threading.Thread(target=self._save_events_to_file, args=(file_path, recorded_events.copy()), daemon=True)
            save_thread.start()

    def _save_events_to_file(self, file_path, events):
        try:
            with open(file_path, "w") as f: json.dump({"actions": events}, f, indent=4)
            self.after(0, self.on_macro_select, None)
        except IOError as e:
            self.after(0, messagebox.showerror, "Save Error", f"Failed to save macro: {e}")
            
    def play_macro(self):
        if not self.macro_listbox.curselection():
            messagebox.showinfo("Playback", "Please select a macro to play.")
            return
        self.record_button.config(state=tk.DISABLED); self.play_button.config(state=tk.DISABLED)
        self.status_label.config(text="Playing...")
        playback_thread = threading.Thread(target=self._run_playback, daemon=True)
        playback_thread.start()

    def _run_playback(self):
        selected_indices = self.macro_listbox.curselection()
        if not selected_indices:
            self.after(0, self.on_playback_finished)
            return
        macro_name = self.macro_listbox.get(selected_indices[0])
        file_path = os.path.join(self.default_macro_folder, f"{macro_name}.json")
        try:
            with open(file_path, "r") as f: events_to_play = json.load(f).get("actions", [])
            if events_to_play: self.pynput_handler.play_macro(events_to_play)
        except (IOError, json.JSONDecodeError) as e:
            self.after(0, messagebox.showerror, "Playback Error", f"Cannot load macro: {e}")
        finally:
            self.after(0, self.on_playback_finished)

    def on_playback_finished(self):
        self.record_button.config(state=tk.NORMAL); self.play_button.config(state=tk.NORMAL)
        self.status_label.config(text="Playback finished.")

if __name__ == "__main__":
    app = MacroApp()
    app.mainloop()