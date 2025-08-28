# Face Recognition Authentication System

This is a flexible face recognition authentication system that can be used standalone or integrated with a virtual assistant.

## Setup

1. Install the required packages:
```bash
pip install -r requirements.txt
```

## Usage

### 1. Collect Face Data
Run the data collection script to create your face dataset:
```bash
python collect_face_data.py
```
- Enter your name when prompted
- The script will capture 50 images of your face
- Press 'q' to quit early if needed

### 2. Train the Model
After collecting face data, train the recognition model:
```bash
python train_model.py
```
This will create a `face_encodings.pkl` file with your face data.

### 3. Authentication
To test the authentication:
```bash
python face_auth.py
```
- Enter your username (same as used in data collection)
- The system will try to authenticate your face

### Integration with Virtual Assistant
To use this system in your virtual assistant, import the authentication module:

```python
from face_auth import verify_user

# In your virtual assistant code
def start_assistant():
    username = "your_username"
    if verify_user(username):
        # Start virtual assistant
        print("Access granted!")
    else:
        print("Access denied!")
```

## Features
- Flexible and modular design
- Easy integration with other applications
- Configurable confidence threshold for authentication
- Real-time face detection and recognition
