# Project Title

## Steps to Setup and Run the Project

### 1. Install Requirements
First, install the necessary packages using `requirements.txt`:
```
pip install -r requirements.txt
```

### 2, Create a .env File
Create a .env file for environment variables and add the following line, replace with tour ADO token [The ADO token should have the following permissions: Code (Read & Write), Packaging (Read), Service Connections (Read & query), Work Items (Read)]:
```
PAT=<ADO_PERSONAL_ACCESS_TOKEN>
```

### 3. Run the Python Script
Finally, run the Python script test.py:
```
python generate_project_status.py
```
