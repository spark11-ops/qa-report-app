print("STARTING TEST SERVER")

from flask import Flask
app = Flask(__name__)

@app.route("/")
def home():
    return "Test server working!"

print("BEFORE RUN")
app.run(host="0.0.0.0", port=5000, debug=True)
print("AFTER RUN")
