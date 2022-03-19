from flask import Flask, render_template, url_for, request,  jsonify
import sys, time, random

from selenium.webdriver.support.expected_conditions import text_to_be_present_in_element



app = Flask(__name__)
@app.route("/")
@app.route("/home")

def home():
    return render_template('index.html')


# -------------Main Dislay Functions---------------------

@app.route('/home', methods = ['POST'])
def result():
    output= request.form.to_dict()
    setup = True

    return render_template('index.html', setup = setup)

@app.route('/screenshot', methods = ['GET','POST'])
def screenshot():
    output= request.form.to_dict()
    screen = True
        
    return render_template('index.html', screen = screen)


@app.route('/move_to_directory', methods = ['GET','POST'])
def move_to_directory():
    output= request.form.to_dict()
    directory = True

    return render_template('index.html', directory = directory)


# ---------------- Secondary Funtions -------------------

@app.route('/setup_banner', methods = ['GET','POST'])
def setup_banner():
    from setup_main import run_prog
    run_prog()


@app.route('/screenshot_maker', methods = ['GET','POST'])
def screenshot_maker():
    from screenshot_main import make_screenshot
    make_screenshot()
    


@app.route('/directory_move', methods = ['GET','POST'])
def directory_move():
    from move_to_directory import move_to_dir
    move_to_dir()


# -----------------------------------------------------------------------------------------#



if __name__ == "__main__":
    app.run(debug= True, port= 5001)