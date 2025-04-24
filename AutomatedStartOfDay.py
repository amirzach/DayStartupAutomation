import webbrowser
import time
import os
import platform
import tkinter as tk
from tkinter import font as tkfont
from tkinter import PhotoImage

def open_microsoft_teams():
    """
    Opens Microsoft Teams application.
    Handles different paths based on operating system.
    """
    system = platform.system()
    print("Attempting to open Microsoft Teams...")
    
    try:
        if system == 'Windows':
            
            #launch ms-teams via Run command
            os.system("start ms-teams:")
            print("Attempted to launch Teams via protocol handler")
            print("Successful launch of Teams via protocol handler")
            return True
            
        else:
            print(f"Unsupported operating system: {system}")
            return False
            
    except Exception as e:
        print(f"Error opening Microsoft Teams: {e}")
        return False

def open_links_in_firefox(urls):
    """
    Opens multiple URLs in Firefox browser.
   
    Args:
        urls (list): List of URLs to open
    """
    # Use firefox if specified, otherwise use default browser
    browser = webbrowser.get('firefox') if 'firefox' in webbrowser._browsers else webbrowser.get()
   
    print(f"Opening {len(urls)} links in Firefox...")
   
    # Open each URL with a small delay
    for i, url in enumerate(urls):
        print(f"Opening link {i+1}: {url}")
        browser.open_new_tab(url)
        # Small delay to prevent overwhelming the browser
        time.sleep(1)
   
    print("All links opened successfully!")

def open_word_document(file_path):
    """
    Opens a Word document at the specified path.
   
    Args:
        file_path (str): Full path to the Word document
    """
    if not os.path.exists(file_path):
        print(f"Error: File not found at {file_path}")
        return False
   
    print(f"Opening Word document: {file_path}")
   
    system = platform.system()
    try:
        if system == 'Windows':
            os.startfile(file_path)
        else:  # Linux
            print("Error, OS not supported")
            return False
       
        print("Word document opened successfully!")
        return True
    except Exception as e:
        print(f"Error opening document: {e}")
        return False

def show_nice_day_popup():
    """
    Shows an aesthetically attractive popup wishing the user a nice day.
    """
    # Create the main window
    root = tk.Tk()
    root.title("Daily Greeting")
    
    # Set window size and position it in the center of the screen
    window_width = 600
    window_height = 250
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    x_position = (screen_width - window_width) // 2
    y_position = (screen_height - window_height) // 2
    root.geometry(f"{window_width}x{window_height}+{x_position}+{y_position}")
    
    # Set window properties
    root.configure(bg="#f0f4f8")
    root.overrideredirect(True)  # Remove window borders
    
    # Create a frame with rounded corners effect
    frame = tk.Frame(root, bg="#ffffff", bd=1, relief=tk.SOLID)
    frame.place(relx=0.5, rely=0.5, anchor=tk.CENTER, relwidth=0.9, relheight=0.85)
    
    # Add greeting text
    title_font = tkfont.Font(family="Arial", size=16, weight="bold")
    greeting_label = tk.Label(frame, text="Have a wonderful and productive day!", 
                            font=title_font, bg="#ffffff", fg="#2c3e50")
    greeting_label.pack(pady=(30, 10))
    
    # Add a motivational message
    message_font = tkfont.Font(family="Arial", size=12)
    message = "Everything is ready.\nHope you have a lovely work day!"
    message_label = tk.Label(frame, text=message, font=message_font, 
                           bg="#ffffff", fg="#7f8c8d", justify=tk.CENTER)
    message_label.pack(pady=10)
    
    # Add a close button
    def close_window():
        root.destroy()
    
    button_frame = tk.Frame(frame, bg="#ffffff")
    button_frame.pack(pady=20)
    
    close_button = tk.Button(button_frame, text="Thank you!", command=close_window,
                           bg="#3498db", fg="white", font=("Arial", 10),
                           relief=tk.FLAT, padx=15, pady=5,
                           activebackground="#2980b9", activeforeground="white")
    close_button.pack()
    
    # Auto-close after 10 seconds
    root.after(10000, close_window)
    
    # Make window appear on top
    root.attributes('-topmost', True)
    
    # Start the main loop
    root.mainloop()

if __name__ == "__main__":
    # List of URLs to open
    links = [
        "https://your own links",
        "https://your own links2",
    ]
   
    # Literal path to your Word documents, can be changed to list if planning to open more than one file
    ms_word = r"path to your Msword file"
   
    # Open the links in Firefox
    open_links_in_firefox(links)
   
    # Open the Word documents
    open_word_document(ms_word)
    
    # Open Microsoft Teams
    open_microsoft_teams()    
    
    # Show the nice day popup after everything is opened
    show_nice_day_popup()
