import os
import sys

# Makes plugin safer to import even if Flask isn't installed
try:
    from flask import Flask, render_template_string, request, redirect, url_for, flash
except ImportError:
    Flask = None

if Flask:
    app = Flask(__name__)
    app.secret_key = 'demo_secret'

    # Simple HTML template with styles for a file explorer look
    TEMPLATE = """
    <!doctype html>
    <html>
    <head>
        <title>CAT Plugin - Flask Explorer</title>
        <style>
            body { font-family: sans-serif; padding: 20px; background-color: #f8f9fa; }
            .container { display: flex; background: white; padding: 20px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); height: 80vh; }
            .sidebar { width: 250px; border-right: 1px solid #dee2e6; padding-right: 20px; margin-right: 20px; flex-shrink: 0; }
            .main { flex: 1; overflow-y: auto; }
            .file-list { list-style: none; padding: 0; }
            .file-item { padding: 8px; border-bottom: 1px solid #eee; display: flex; align-items: center; }
            .file-item:hover { background-color: #e9ecef; }
            .file-icon { margin-right: 10px; }
            .folder { font-weight: bold; color: #007bff; text-decoration: none; }
            .file { color: #495057; }
            .btn { display: block; width: 100%; padding: 10px; margin-bottom: 10px; border: none; border-radius: 4px; cursor: pointer; color: white; font-size: 14px; text-align: center; }
            .btn-primary { background-color: #007bff; }
            .btn-primary:hover { background-color: #0056b3; }
            .btn-success { background-color: #28a745; }
            .btn-success:hover { background-color: #218838; }
            .btn-info { background-color: #17a2b8; }
            .btn-info:hover { background-color: #138496; }
            .flash { padding: 15px; background: #d4edda; color: #155724; margin-bottom: 20px; border: 1px solid #c3e6cb; border-radius: 4px; }
            .breadcrumb { margin-bottom: 15px; font-size: 14px; color: #666; }
        </style>
    </head>
    <body>
        <h1>Plugin Demo: Flask File Explorer</h1>
        
        {% with messages = get_flashed_messages() %}
          {% if messages %}
            <div class="flash">
            {% for message in messages %}
              {{ message }}<br>
            {% endfor %}
            </div>
          {% endif %}
        {% endwith %}

        <div class="container">
            <div class="sidebar">
                <h3>Control Panel</h3>
                <form action="{{ url_for('action') }}" method="post">
                    <input type="hidden" name="current_path" value="{{ current_path }}">
                    <button name="btn" value="analyze" class="btn btn-primary">⚡ Run Analysis</button>
                    <button name="btn" value="backup" class="btn btn-success">💾 Backup Files</button>
                    <button name="btn" value="refresh" class="btn btn-info">🔄 Refresh View</button>
                </form>
                <hr>
                <p><small>This sidebar demonstrates how a plugin can expose custom actions that interact with the file system context.</small></p>
            </div>
            <div class="main">
                <h3>File Explorer</h3>
                <div class="breadcrumb">Path: {{ current_path }}</div>
                
                <div class="file-item">
                    <span class="file-icon">⬆️</span>
                    <a href="{{ url_for('index', path=parent_path) }}" class="folder">.. (Parent Directory)</a>
                </div>
                
                <ul class="file-list">
                {% for item in items %}
                    <li class="file-item">
                        {% if item.is_dir %}
                            <span class="file-icon">📁</span>
                            <a class="folder" href="{{ url_for('index', path=item.path) }}">{{ item.name }}</a>
                        {% else %}
                            <span class="file-icon">📄</span>
                            <span class="file">{{ item.name }}</span>
                        {% endif %}
                    </li>
                {% endfor %}
                </ul>
            </div>
        </div>
    </body>
    </html>
    """

    @app.route('/')
    def index():
        path = request.args.get('path', os.getcwd())
        if not os.path.exists(path):
            flash(f"Path does not exist: {path}")
            path = os.getcwd()

        items = []
        try:
            with os.scandir(path) as it:
                for entry in it:
                    items.append({
                        'name': entry.name,
                        'path': entry.path,
                        'is_dir': entry.is_dir()
                    })
            items.sort(key=lambda x: (not x['is_dir'], x['name']))
        except PermissionError:
            flash(f"Permission denied accessing {path}")
            # go back to parent or current if possible
            items = []

        parent_path = os.path.dirname(os.path.abspath(path))

        return render_template_string(TEMPLATE, items=items, current_path=os.path.abspath(path), parent_path=parent_path)

    @app.route('/action', methods=['POST'])
    def action():
        btn = request.form.get('btn')
        path = request.form.get('current_path', os.getcwd())

        if btn == 'analyze':
            flash(f"Analysis started on directory: {path}")
        elif btn == 'backup':
            flash(f"Simulating backup for: {path}")
        elif btn == 'refresh':
            flash("View refreshed.")

        return redirect(url_for('index', path=path))
else:
    app = None

def main():
    print("Starting Flask File Explorer Plugin...")

    if not Flask:
        print("This plugin requires Flask. Please install it with: pip install flask")
        sys.exit(1)

    port = 5001  # Default to 5001 to avoid macOS AirPlay conflict on 5000
    if len(sys.argv) > 1:
        try:
            port = int(sys.argv[1])
        except ValueError:
            pass

    print(f"Open your browser at http://127.0.0.1:{port}")

    try:
        app.run(host='0.0.0.0', port=port, debug=True)
    except OSError as e:
        if "Address already in use" in str(e):
            print(f"\nError: Port {port} is already in use.")
            print("Try running with a different port argument, e.g.:")
            print(f"  ./config-assessment-tool.sh --plugin start demo_flask_cli {port+1}")
        else:
            raise

if __name__ == "__main__":
    main()
