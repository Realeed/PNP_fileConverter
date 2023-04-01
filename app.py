from flask import Flask, render_template, request, send_file

app = Flask(__name__)
app.secret_key = "secret"

@app.route("/", methods = ["POST", "GET"])
def main():
    return render_template('index.html')



if __name__  == "__main__":
    app.run(host='0.0.0.0', port=80, debug=True)