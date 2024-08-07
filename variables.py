SECRET_NAME = "MSGraph"
SCOPES = "Mail.Send"
BASE_GRAPH_URL = "https://graph.microsoft.com/v1.0"
SUBJECT = "Test HTML Email using modern authentication"
BODY = """
<html>
  <body>
    <h1>This is a test email</h1>
    <p>This is a test email body in HTML format.</p>
    <p><strong>Enjoy your day!</strong></p>
  </body>
</html>
"""
RECIPIENTS = ["firstname.lastname@example.com"]
