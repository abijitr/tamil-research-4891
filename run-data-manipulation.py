import subprocess

# Specify the path to the Python file you want to run
python_file_path = 'formant-split.py'

# Run the Python file
print('running: ', python_file_path)
subprocess.run(['python', python_file_path])