<h1>Report Generator & Mailer</h1>
A Web Application - PDF Report Generator using Flask and AJAX

This is a web application which creates PDF reports for induvidual students based on an inputted Excel File - specifically made for JSS Academy of Technical Education, Bangalore.<br>
It is useful in academic settings where teachers needed to create separate performance reports for students and it was too cumbersome.
This web application takes in an excel file with information of all students and their performance and attendance, and generates induvidual reports for each one.
It can also mail the generated reports to emails associated to every student.

How to use:
1. Install the latest version of python from: https://www.python.org/downloads/
2. Select the "Add Python to PATH" checkbox during installation.
3. Download ```jQuery.js``` and place it in the js folder of the directory
4. Open the Terminal in the folder where the files are located.
5. Run the ```app.py``` file in command prompt on Windows or Terminal on Linux/UNIX using ```python3 app.py```
6. Once the code is running, Open your browser and in the address bar, type: ```localhost:8000```
7. Now you can use the web application!

NOTE: The reason why this application has not been converted into an actual production application and kept in a development environment is because it is unable to handle multiple connections, and is better off used on induvidual's computers as a local server.
