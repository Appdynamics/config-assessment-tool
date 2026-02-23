import sys
import os
import multiprocessing
from streamlit.web import cli as st_cli

def resolve_path(path):
    if getattr(sys, 'frozen', False):
        base_path = sys._MEIPASS
    else:
        base_path = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

    return os.path.join(base_path, path)

def run_ui():
    """
    Launch the Streamlit UI.
    """
    # Path to the frontend script within the bundle
    # We will ensure 'frontend' folder is included in the bundle via spec file
    script_path = resolve_path(os.path.join("frontend", "frontend.py"))

    if not os.path.exists(script_path):
        print(f"Error: Frontend script not found at {script_path}")
        sys.exit(1)

    # Ensure the directory containing 'backend' package is in sys.path
    # In frozen app, sys._MEIPASS is usually where bundled packages live.
    if getattr(sys, 'frozen', False):
        # sys._MEIPASS is the root of the bundle content (_internal)
        # If 'backend' package is in sys._MEIPASS, then sys._MEIPASS should be in PYTHONPATH
        # Streamlit spawns a new process or execs in current?
        # Streamlit runner executes code. It might mess with sys.path or use its own.
        # But here we are calling st_cli.main() in-process.
        # However, we should verify sys.path has sys._MEIPASS. It usually does for frozen apps.
        pass

    # Add backend directory to sys.path to support imports like 'from api...' inside Engine.py
    # Engine.py seems to rely on 'backend/' folder being in path because it imports 'api' directly.
    # In bundle: _internal/backend/api
    # This logic is now handled globally in __main__ block

    # Streamlit CLI expects sys.argv to start with "streamlit"
    # We construct the arguments: "streamlit run path/to/frontend.py [args]"
    # Currently we don't pass extra args to frontend, but we could filter sys.argv
    sys.argv = [
        "streamlit",
        "run",
        script_path,
        "--global.developmentMode=false"
    ]

    print("Starting Config Assessment Tool UI...")
    sys.exit(st_cli.main())

def run_backend():
    """
    Run the Backend CLI.
    """
    # backend.backend.main is a click command.
    # We can invoke it.
    # Note: backend.py uses click, which parses sys.argv.
    # We might need to ensure backend module is importable.

    # In the spec file, we will bundle the app such that 'backend' is a package or in path.
    # Let's import it here.
    try:
        import backend.backend
        # Click's main() will handle sys.argv parsing/execution and sys.exit
        # Force prog_name to be "config-assessment-tool" instead of script name
        backend.backend.main(prog_name="config-assessment-tool")
    except ImportError as e:
        print(f"Error importing backend: {e}")
        sys.exit(1)

if __name__ == "__main__":
    multiprocessing.freeze_support()

    # Add backend directory to sys.path
    if getattr(sys, 'frozen', False):
        backend_path = os.path.join(sys._MEIPASS, 'backend')
        if backend_path not in sys.path:
            sys.path.append(backend_path)

        # Set SSL_CERT_FILE to the bundled certifi cacert.pem
        # PyInstaller bundles certifi's cacert.pem in 'certifi/cacert.pem' relative to _MEIPASS
        cert_path = os.path.join(sys._MEIPASS, 'certifi', 'cacert.pem')
        if os.path.exists(cert_path):
            os.environ['SSL_CERT_FILE'] = cert_path
            os.environ['REQUESTS_CA_BUNDLE'] = cert_path
            # print(f"Setting SSL CA Bundle to: {cert_path}") # Debug only

    # Check for UI flags
    if "--ui" in sys.argv or "--run" in sys.argv:
        run_ui()
    # Intercept help to show UI options before backend options
    elif "--help" in sys.argv:
        print("Config Assessment Tool Bundle")
        print("Usage:")
        print("  Launch UI Mode:     config-assessment-tool --ui")
        print("  Launch Backend CLI: config-assessment-tool [OPTIONS]")
        print("\nBackend CLI Help:")
        run_backend()
    else:
        run_backend()
