import os

from flask import Flask, render_template_string
from dotenv import load_dotenv
from identity.flask import Auth

# Load variables from .env
load_dotenv()

# Create Flask app
app = Flask(__name__)

# Flask session configuration
app.config["SECRET_KEY"] = "secret-session-key"
app.config["SESSION_TYPE"] = "filesystem"

# Initialise Microsoft Entra ID authentication
auth = Auth(
    app,
    authority=os.getenv("AUTHORITY"),
    client_id=os.getenv("CLIENT_ID"),
    client_credential=os.getenv("CLIENT_SECRET"),
    redirect_uri="http://localhost:5000/getAToken",
)

# Simple dashboard page that displays sign-in information
HTML_PAGE = """
<!DOCTYPE html>
<html>
<head>
    <title>Finance Dashboard</title>
</head>
<body>

<h1>Finance Dashboard</h1>

<p style="color: green;">
    <strong>Status:</strong> Authenticated via Entra ID
</p>

<p>
    <strong>Welcome:</strong> {{ user.get('name', 'Unknown') }}
</p>

<p>
    <strong>Email/UPN:</strong>
    {{ user.get('preferred_username', 'Unknown') }}
</p>

<p>
    <a href="{{ url_for('identity.logout') }}">Sign Out</a>
</p>

</body>
</html>
"""

# Displays the homepage
@app.route("/")
@auth.login_required
def index(*, context):
    user = context["user"]

    return render_template_string(
        HTML_PAGE,
        user=user
    )


if __name__ == "__main__":
    app.run(port=5000, debug=True)