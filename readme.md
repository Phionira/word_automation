
<img src="https://res.cloudinary.com/durbs4yfq/image/upload/v1694684619/logo/square_h2aswm.png" width="200" height="200" style="">

# Automated Letter Sending Script

## Introduction

This document provides instructions for setting up a script that automatically builds personalized letters and sends them to multiple schools.

## Prerequisites

- Python 3.6 or higher
- A text editor (e.g., Visual Studio Code, Sublime Text)
- An email account for sending the letters
- Microsoft word and Microsoft Excel (preferably 2019 or Office 365)
- pip (Python package installer)
- virtualenv (Python virtual environment)
- Git

## Steps
## Steps

1. **Clone the Repository**: Clone the Python repository from GitHub using the following command in your terminal:

```bash
git clone https://github.com/Phionira/word_automation.git
```
2. **Navigate to the Project Directory**: Use the cd command to navigate to the project directory:
```bash
cd word_automation
```
3. **Create a Virtual Environment**: Set up a virtual environment using the virtualenv command:
```bash
python -m venv env
```
4. **Activate the Virtual Environment**: Activate the virtual environment using the following command:
- On Windows:
```bash
.\env\Scripts\activate
```
- On Unix or MacOS:
```bash
source env/bin/activate
```

5. **Install the Dependencies**: Install the project dependencies from the requirements.txt file using the pip install command:
```bash
pip install -r requirements.txt
```
 
6. **Run the Script**: Finally, run the word_automation.py script using the python command:
```bash
python word_automation.py
```

** Note: ** This script requires Please reach out the IT Department for further information about modifying the template and excel documents. 
1. **Create a Template Letter**: Write a template for the letter with placeholders for the personalized information.

2. **Prepare the Data**: Create a CSV file with the names of the schools and the personalized information you want to include in the letters.

3. **Write the Script**: Write a Python script that reads the data from the CSV file, fills in the placeholders in the letter template, and sends the letters via email.

4. **Test the Script**: Run the script with a small subset of data to ensure it works as expected.

5. **Run the Script**: Once you're confident the script is working correctly, run it with the full data set.

## Conclusion

With this script, you can automate the process of sending personalized letters to multiple schools. This can save you a significant amount of time and ensure that each school receives a personalized letter.

## Contact

For more information or assistance, please contact the IT department.

soporte@lms.phionira.com
