import webbrowser
import time
import os
import platform
import threading
import customtkinter as ctk
import subprocess
import win32com.client
import win32gui
import win32con

# Set appearance mode and default color theme for customtkinter
ctk.set_appearance_mode("dark")  # Options: "dark", "light", "system"
ctk.set_default_color_theme("blue")  # Options: "blue", "green", "dark-blue"

# Global variables for tracking progress
TOTAL_TASKS = 0
COMPLETED_TASKS = 0
PROGRESS_WINDOW = None
PROGRESS_BAR = None
PROGRESS_LABEL = None
STATUS_LABEL = None
TASK_WEIGHTS = {}  # Dictionary to store weight of each task

def calculate_progress():
    """
    Calculates the current progress percentage based on completed weighted tasks
    
    Returns:
        int: Progress percentage (0-100)
    """
    global TOTAL_TASKS
    
    if TOTAL_TASKS == 0:
        return 0
    
    return min(99, int((COMPLETED_TASKS / TOTAL_TASKS) * 100))

def update_progress(step_text, task_name=None, is_completed=False):
    """
    Updates the progress bar and status text
    
    Args:
        step_text (str): Text describing the current step
        task_name (str): Name of the task being performed/completed
        is_completed (bool): Whether this task is now complete
    """
    global COMPLETED_TASKS, PROGRESS_BAR, PROGRESS_LABEL, STATUS_LABEL, TASK_WEIGHTS
    
    # Update task completion if a task was specified
    if task_name and is_completed and task_name in TASK_WEIGHTS:
        COMPLETED_TASKS += TASK_WEIGHTS[task_name]
    
    # Update UI if it exists
    if PROGRESS_WINDOW and PROGRESS_WINDOW.winfo_exists():
        try:
            progress_value = calculate_progress()
            PROGRESS_BAR.set(progress_value / 100)  # customtkinter uses 0.0-1.0 range
            PROGRESS_LABEL.configure(text=f"{progress_value}% Complete")
            STATUS_LABEL.configure(text=step_text)
            PROGRESS_WINDOW.update()
        except Exception:
            # Window might have been closed
            pass

def register_tasks(tasks_dict):
    """
    Registers the tasks and their weights for progress tracking
    
    Args:
        tasks_dict (dict): Dictionary mapping task names to their weight values
    """
    global TASK_WEIGHTS, TOTAL_TASKS
    
    TASK_WEIGHTS = tasks_dict
    TOTAL_TASKS = sum(tasks_dict.values())

def minimize_window(window_title, delay=1):
    """
    Finds and minimizes a window with the given title.
    
    Args:
        window_title (str): Part of the window title to match
        delay (int): Time to wait before attempting to minimize
    """
    # Wait for the window to open
    time.sleep(delay)
    
    def minimize_callback(hwnd, titles):
        window_text = win32gui.GetWindowText(hwnd).lower()
        title_to_match = titles[0].lower()
        
        # More verbose logging to debug matching issues
        if title_to_match in window_text and win32gui.IsWindowVisible(hwnd):
            update_progress(f"Minimizing: {window_text}")
            try:
                win32gui.ShowWindow(hwnd, win32con.SW_MINIMIZE)
                titles[1] = True  # Mark that we found and minimized a window
            except Exception as e:
                update_progress(f"Error minimizing window: {e}")
        elif title_to_match and len(window_text) > 3 and win32gui.IsWindowVisible(hwnd):
            print(f"Saw window: '{window_text}' (no match for '{title_to_match}')")
    
    # Use list to track if we found a match
    found_match = [window_title, False]
    win32gui.EnumWindows(lambda hwnd, param: minimize_callback(hwnd, param), found_match)
    
    if not found_match[1]:
        update_progress(f"Looking for window: '{window_title}'")
        # Try a more aggressive approach with a broader search
        if "firefox" in window_title.lower():
            alternate_titles = ["Mozilla", "Firefox"]
        elif "chrome" in window_title.lower():
            alternate_titles = ["Google Chrome", "Chrome"]
        elif "edge" in window_title.lower():
            alternate_titles = ["Microsoft Edge", "Edge"]
        elif "http" in window_title.lower():
            alternate_titles = ["Mozilla", "Firefox", "Chrome", "Edge"]
        else:
            alternate_titles = []
            
        for alt_title in alternate_titles:
            alt_found = [alt_title, False]
            win32gui.EnumWindows(lambda hwnd, param: minimize_callback(hwnd, param), alt_found)
            if alt_found[1]:
                update_progress(f"Minimized window: {alt_title}")
                break

def open_microsoft_teams():
    """
    Opens Microsoft Teams application and minimizes it.
    """
    system = platform.system()
    update_progress("Launching Microsoft Teams...")
    
    try:
        if system == 'Windows':
            # Launch ms-teams via Run command
            os.system("start ms-teams:")
            
            # Minimize the Teams window
            time.sleep(5)  # Wait longer for Teams to load
            minimize_window("Microsoft Teams", 2)
            
            update_progress("Microsoft Teams ready (minimized)", "teams_launch", True)
            return True
            
        else:
            update_progress(f"Unsupported OS for Teams: {system}", "teams_launch", True)
            return False
            
    except Exception as e:
        update_progress(f"Error opening Microsoft Teams: {e}", "teams_launch", True)
        return False

def open_links_in_firefox(urls):
    """
    Opens multiple URLs in Firefox browser in minimized state.
   
    Args:
        urls (list): List of URLs to open
    """
    global COMPLETED_TASKS, TASK_WEIGHTS
    
    # Use firefox if specified, otherwise use default browser
    browser_name = 'firefox' if 'firefox' in webbrowser._browsers else None
    
    update_progress(f"Preparing to open {len(urls)} websites...")
    
    # First, minimize existing browser windows to ensure we start clean
    if browser_name == 'firefox':
        minimize_window("Firefox", 1)
    else:
        minimize_window("Chrome", 1)
        minimize_window("Edge", 1)
    
    # Calculate weight per site
    site_weight = TASK_WEIGHTS.get("browser_sites", 0) / max(1, len(urls))
    
    # Open the first URL with new window
    update_progress(f"Opening first site: {urls[0].split('//')[1].split('/')[0]}")
    if browser_name:
        browser = webbrowser.get(browser_name)
        browser.open(urls[0])
    else:
        webbrowser.open(urls[0])
    
    # Wait longer before minimizing - browsers can take time to fully launch
    time.sleep(5) #5 seconds delay to ensure the browser is fully loaded
    
    # Try to minimize the browser with different potential titles
    if browser_name == 'firefox':
        minimize_window("Mozilla Firefox", 1)
        minimize_window("Firefox", 1)
    else:
        minimize_window("http", 1)
        minimize_window("Chrome", 1)
        minimize_window("Edge", 1)
    
    # Mark first site as complete - partial completion of the task
    COMPLETED_TASKS += site_weight
    update_progress(f"First site loaded: {urls[0].split('//')[1].split('/')[0]}")
    
    # Open remaining URLs in new tabs
    for i, url in enumerate(urls[1:], 2):
        site_name = url.split('//')[1].split('/')[0]
        update_progress(f"Opening site {i}/{len(urls)}: {site_name}")
        if browser_name:
            browser.open_new_tab(url)
        else:
            webbrowser.open_new_tab(url)
        # Small delay to prevent overwhelming the browser
        time.sleep(0.5)
        
        # Update progress for each site
        COMPLETED_TASKS += site_weight
        update_progress(f"Site loaded: {site_name}")
    
    # Final attempt to minimize after all tabs are opened
    time.sleep(2)
    if browser_name == 'firefox':
        minimize_window("Mozilla Firefox", 1)
    else:
        minimize_window("http", 1)
        minimize_window("Chrome", 1)
        minimize_window("Edge", 1)
    
    update_progress("All websites loaded and minimized", "browser_sites", True)

def open_word_documents(file_paths):
    """
    Opens multiple Word documents and ensures each one is minimized.
   
    Args:
        file_paths (list): List of paths to Word documents
    """
    global COMPLETED_TASKS, TASK_WEIGHTS
    
    update_progress(f"Preparing {len(file_paths)} Word documents...")
    
    system = platform.system()
    if system != 'Windows':
        update_progress(f"Error: Unsupported operating system: {system}", "word_docs", True)
        return False
    
    try:
        # Start a single Word application instance
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = True
        
        # Calculate weight per document
        doc_weight = TASK_WEIGHTS.get("word_docs", 0) / max(1, len(file_paths))
        
        # Open each document and explicitly minimize its window
        for i, file_path in enumerate(file_paths, 1):
            if not os.path.exists(file_path):
                update_progress(f"Error: File not found: {os.path.basename(file_path)}")
                # Still count this toward completion
                COMPLETED_TASKS += doc_weight
                continue
                
            file_name = os.path.basename(file_path)
            update_progress(f"Opening document {i}/{len(file_paths)}: {file_name}")
            try:
                # Open the document
                doc = word.Documents.Open(file_path)
                
                # Give Word a moment to fully open the document
                time.sleep(1)
                
                # Get the current active window for this document
                active_window = word.ActiveWindow
                active_window.WindowState = 2  # 2 = minimized
                
                update_progress(f"Document ready: {file_name}")
                
            except Exception as e:
                update_progress(f"Error with document: {file_name}")
                # Fallback method
                os.startfile(file_path)
                # Try to minimize the window
                minimize_window(file_name, 3)
            
            # Update progress for each document
            COMPLETED_TASKS += doc_weight
            update_progress(f"Document {i}/{len(file_paths)} processed")
        
        update_progress("All documents ready", "word_docs", True)
        return True
        
    except Exception as e:
        update_progress(f"Error initializing Word")
        # Fallback: open documents individually
        doc_weight = TASK_WEIGHTS.get("word_docs", 0) / max(1, len(file_paths))
        
        for i, file_path in enumerate(file_paths, 1):
            if os.path.exists(file_path):
                file_name = os.path.basename(file_path)
                update_progress(f"Opening: {file_name}")
                os.startfile(file_path)
                minimize_window(file_name, 3)
                
                # Update progress for each document
                COMPLETED_TASKS += doc_weight
                update_progress(f"Document {i}/{len(file_paths)} processed")
                
        update_progress("All documents processed with fallback method", "word_docs", True)
        return False

def create_progress_ui():
    """
    Creates a sleek, modern progress bar UI using customtkinter
    """
    global PROGRESS_WINDOW, PROGRESS_BAR, PROGRESS_LABEL, STATUS_LABEL
    
    # Create the main window
    ctk.set_appearance_mode("dark")  # Set dark mode
    root = ctk.CTk()
    root.title("Workspace Startup")
    PROGRESS_WINDOW = root
    
    # Set window size and position it in the center of the screen
    window_width = 600
    window_height = 300
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    x_position = (screen_width - window_width) // 2
    y_position = (screen_height - window_height) // 2
    root.geometry(f"{window_width}x{window_height}+{x_position}+{y_position}")
    
    # Remove window borders (CustomTkinter already has a modern look)
    root.overrideredirect(True)
    
    # Make window appear on top
    root.attributes('-topmost', True)
    
    # Create a main frame
    main_frame = ctk.CTkFrame(root)
    main_frame.pack(fill="both", expand=True, padx=10, pady=10)
    
    # Add greeting text
    greeting_label = ctk.CTkLabel(
        main_frame, 
        text="Workspace Initialization", 
        font=ctk.CTkFont(family="Segoe UI", size=22, weight="bold")
    )
    greeting_label.pack(pady=(40, 20))
    
    # Add progress bar (progress_bar in customtkinter uses 0-1 range)
    PROGRESS_BAR = ctk.CTkProgressBar(
        main_frame,
        width=400,
        height=10,
        corner_radius=5
    )
    PROGRESS_BAR.set(0)  # Set initial value to 0
    PROGRESS_BAR.pack(pady=15)
    
    # Add percentage label
    PROGRESS_LABEL = ctk.CTkLabel(
        main_frame,
        text="0% Complete",
        font=ctk.CTkFont(family="Segoe UI", size=14, weight="bold")
    )
    PROGRESS_LABEL.pack(pady=5)
    
    # Add status message
    STATUS_LABEL = ctk.CTkLabel(
        main_frame,
        text="Initializing...",
        font=ctk.CTkFont(family="Segoe UI", size=12),
        wraplength=500
    )
    STATUS_LABEL.pack(pady=15)
    
    # Add fade-in effect when showing
    root.attributes('-alpha', 0.0)
    root.update()
    for i in range(0, 11):
        root.attributes('-alpha', i/10)
        root.update()
        time.sleep(0.03)
    
    return root

def complete_progress_ui():
    """
    Updates the UI to show completion and adds a close button
    """
    global PROGRESS_WINDOW, PROGRESS_BAR, PROGRESS_LABEL, STATUS_LABEL
    
    if not PROGRESS_WINDOW or not PROGRESS_WINDOW.winfo_exists():
        return
    
    # Set a final task for UI completion itself
    update_progress("Finalizing workspace setup...", "ui_completion", True)
    
    # Now update to 100% - this happens only at the very end
    PROGRESS_BAR.set(1.0)  # customtkinter uses 0.0-1.0 range
    PROGRESS_LABEL.configure(text="100% Complete")
    STATUS_LABEL.configure(text="Workspace ready. Welcome Amir!")
    
    def close_window():
        # Add fade-out effect
        for i in range(10, 0, -1):
            PROGRESS_WINDOW.attributes('-alpha', i/10)
            PROGRESS_WINDOW.update()
            time.sleep(0.02)
        PROGRESS_WINDOW.destroy()
    
    # Add a close button
    close_button = ctk.CTkButton(
        PROGRESS_WINDOW,
        text="Begin",
        command=close_window,
        font=ctk.CTkFont(family="Segoe UI", size=14, weight="bold"),
        corner_radius=8,
        height=38,
        width=120,
        hover=True,
        fg_color="#e94560",  # Custom button color
        hover_color="#ba181b"  # Darker on hover
    )
    close_button.place(relx=0.5, rely=0.85, anchor="center")
    
    # Auto-close after 15 seconds
    PROGRESS_WINDOW.after(15000, close_window)

def startup_sequence():
    """
    Runs the startup sequence in a separate thread
    """
    # List of URLs to open
    links = [
        "enter website link",
    ]
   
    # List of Word documents
    word_documents = [
        r"link:\to\yourfile",
    ]
    
    # Register tasks and their weights (total should equal 100)
    register_tasks({
        "initialization": 5,      # Initial setup
        "browser_sites": 35,      # Opening websites (weighted heavily as browser startup is slow)
        "word_docs": 30,          # Opening documents
        "teams_launch": 20,       # Teams application
        "ui_completion": 10       # Final UI updates and completion steps
    })
    
    # Mark initialization as complete
    update_progress("Starting workspace initialization...", "initialization", True)
    
    # Small delay to let the UI render first
    time.sleep(1)
    
    # Open the links in Firefox
    update_progress("Starting browser initialization...")
    open_links_in_firefox(links)
   
    # Open the Word documents
    update_progress("Preparing document workspace...")
    open_word_documents(word_documents)
    
    # Open Microsoft Teams
    update_progress("Connecting communication systems...")
    open_microsoft_teams()
    
    # Complete the progress UI (marks the final task as complete and shows 100%)
    complete_progress_ui()

if __name__ == "__main__":
    # Create and show the progress UI
    root = create_progress_ui()
    
    # Start the startup sequence in a separate thread
    threading.Thread(target=startup_sequence, daemon=True).start()
    
    # Start the main loop for the UI
    root.mainloop()
