import subprocess
import sys
import os

def launch_dashboard():
    print("Starting the Streamlit Dashboard... Please wait.")
    
    # 1. Get the exact Python executable path being used right now
    python_exe = sys.executable
    
    # 2. Get the absolute path to your Comp.Py file to avoid missing file errors
    current_dir = os.path.dirname(os.path.abspath(__file__))
    app_path = os.path.join(current_dir, "Comp.py")
    
    # 3. Build the terminal command automatically
    command = [python_exe, "-m", "streamlit", "run", app_path]
    
    # 4. Execute the command
    try:
        subprocess.run(command)
    except KeyboardInterrupt:
        print("\nDashboard stopped.")
    except Exception as e:
        print(f"An error occurred: {e}")

if __name__ == "__main__":
    launch_dashboard()