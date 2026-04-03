import click
import os
import sys
import subprocess

PLUGIN_DIR = "plugins"

@click.group()
def cli():
    pass

def is_remote_session():
    """Check if the script is running in a remote SSH session."""
    return any(key in os.environ for key in ["SSH_CLIENT", "SSH_TTY", "SSH_CONNECTION"])

@cli.command(name="list")
def list_plugins():
    """List available plugins"""
    if not os.path.exists(PLUGIN_DIR):
        print("No plugins directory found.")
        return

    plugins = [d for d in os.listdir(PLUGIN_DIR) if os.path.isdir(os.path.join(PLUGIN_DIR, d)) and not d.startswith('__')]
    if not plugins:
        print("No plugins found.")
    else:
        print("Available plugins:")
        for p in plugins:
            plugin_path = os.path.join(PLUGIN_DIR, p)
            main_file = os.path.join(plugin_path, "main.py")
            readme_file = os.path.join(plugin_path, "README.md")

            if os.path.exists(main_file):
                if os.path.exists(readme_file):
                    if is_remote_session():
                         status = f"(run: config-assessment-tool --plugin docs {p})"
                    else:
                        try:
                            import tempfile
                            import html

                            temp_dir = tempfile.gettempdir()
                            html_filename = f"cat_plugin_{p}_readme.html"
                            html_path = os.path.join(temp_dir, html_filename)

                            with open(readme_file, 'r', encoding='utf-8') as rf:
                                content = rf.read()

                            content_escaped = html.escape(content)
                            html_content = f"""<!DOCTYPE html>
<html>
<head>
<title>{p} Documentation</title>
<style>body {{ font-family: monospace; padding: 20px; }}</style>
</head>
<body>
<h3>Documentation for {p}</h3>
<hr>
<pre style="white-space: pre-wrap; word-wrap: break-word;">
{content_escaped}
</pre>
</body>
</html>"""
                            with open(html_path, 'w', encoding='utf-8') as wf:
                                wf.write(html_content)

                            abs_link_path = os.path.abspath(html_path)
                        except Exception:
                            abs_link_path = os.path.abspath(readme_file)

                        # "see docs" link
                        link = f"\033]8;;file://{abs_link_path}\033\\see docs\033]8;;\033\\"
                        status = f"({link})"
                else:
                    status = "(ready-to-use)"
            else:
                status = "(Missing main.py)"

            print(f"  - {p} {status}")

@cli.command()
@click.argument('name')
@click.argument('args', nargs=-1)
def start(name, args):
    """Start a plugin"""
    plugin_path = os.path.join(PLUGIN_DIR, name)
    if not os.path.exists(plugin_path):
        print(f"Plugin '{name}' not found.")
        sys.exit(1)

    main_file = os.path.join(plugin_path, "main.py")
    if not os.path.exists(main_file):
        print(f"Plugin '{name}' does not have a main.py entry point.")
        sys.exit(1)

    # Revert to local logic for pipenv/virtualenv handling or just usage of current python
    # We will assume plugins manage dependencies via requirements.txt if they are standalone
    # and we do a simple check here similar to before.

    python_executable = sys.executable
    requirements_file = os.path.join(plugin_path, "requirements.txt")

    if os.path.exists(requirements_file):
        print(f"Loading plugin: {name},   detected standalone plugin. will not run as part of this process")
        print(f"Dependency file found: {requirements_file}")
        venv_dir = os.path.join(plugin_path, ".venv")
        # Handle Windows vs Unix venv paths
        if sys.platform == "win32":
            venv_python = os.path.join(venv_dir, "Scripts", "python.exe")
        else:
            venv_python = os.path.join(venv_dir, "bin", "python")

        if not os.path.exists(venv_dir):
            print(f"Creating isolated environment for {name}...")
            subprocess.check_call([sys.executable, "-m", "venv", venv_dir])

            print(f"Installing dependencies from requirements.txt...")
            subprocess.check_call([venv_python, "-m", "pip", "install", "-r", requirements_file])
            print("Dependencies installed.")

        python_executable = venv_python

    print(f"Starting plugin {name}...", flush=True)
    cmd = [python_executable, main_file] + list(args)

    # We must preserve environment variables
    env = os.environ.copy()
    # Add plugin dir to PYTHONPATH so it can import its own modules
    # And current PYTHONPATH
    current_pythonpath = env.get("PYTHONPATH", "")
    env["PYTHONPATH"] = f"{plugin_path}:{current_pythonpath}"

    prefix = f"plugin ({name}): "

    try:
        # Use Popen to capture output line by line
        with subprocess.Popen(
            cmd,
            env=env,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,  # Merge stderr into stdout
            text=True,
            bufsize=1,  # Line buffered
            universal_newlines=True
        ) as process:
            # Print output with prefix
            for line in process.stdout:
                print(f"{prefix}{line}", end='')

            # Wait for process to finish
            process.wait()

            if process.returncode != 0:
                print(f"\nPlugin exited with error code {process.returncode}")
                # We don't necessarily exit here, just report it

    except KeyboardInterrupt:
        print("\nPlugin execution interrupted.")
    except Exception as e:
        print(f"Error executing plugin: {e}")

@cli.command(name="docs")
@click.argument('name')
def show_docs(name):
    """Show documentation for a plugin in the terminal"""
    plugin_path = os.path.join(PLUGIN_DIR, name)
    readme_file = os.path.join(plugin_path, "README.md")

    if os.path.exists(readme_file):
        try:
            # Try to use rich for better formatting if available, otherwise just print
            from rich.console import Console
            from rich.markdown import Markdown
            console = Console()
            with open(readme_file, 'r', encoding='utf-8') as f:
                md = Markdown(f.read())
            console.print(md)
        except ImportError:
            # Fallback to plain text
            with open(readme_file, 'r', encoding='utf-8') as f:
                print(f.read())
    else:
        print(f"No documentation found for plugin '{name}'.")

if __name__ == "__main__":
    cli()
