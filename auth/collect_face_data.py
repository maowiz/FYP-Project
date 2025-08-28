import cv2
import os
import time

def create_dataset(name, num_samples=200):
    """
    Capture face images and save them to dataset folder
    Args:
        name: Name of the person
        num_samples: Number of face samples to collect
    """
    # Create dataset directory if it doesn't exist
    dataset_dir = "dataset"
    if not os.path.exists(dataset_dir):
        os.makedirs(dataset_dir)
    
    # Create person's directory
    person_dir = os.path.join(dataset_dir, name)
    if not os.path.exists(person_dir):
        os.makedirs(person_dir)

    # Initialize camera
    cap = cv2.VideoCapture(0)
    face_cascade = cv2.CascadeClassifier(cv2.data.haarcascades + 'haarcascade_frontalface_default.xml')
    
    count = 0
    print(f"Collecting face data for {name}. Press 'q' to quit.")
    
    while count < num_samples:
        ret, frame = cap.read()
        if not ret:
            print("Failed to grab frame")
            break
            
        gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
        faces = face_cascade.detectMultiScale(gray, 1.3, 5)
        
        for (x, y, w, h) in faces:
            cv2.rectangle(frame, (x, y), (x+w, y+h), (255, 0, 0), 2)
            # Save the captured face
            if count < num_samples:
                face_img = frame[y:y+h, x:x+w]
                save_path = os.path.join(person_dir, f"{name}_{count}.jpg")
                cv2.imwrite(save_path, face_img)
                count += 1
                print(f"Captured image {count}/{num_samples}")
                time.sleep(0.1)  # Small delay between captures
        
        cv2.imshow('Collecting Face Data', frame)
        if cv2.waitKey(1) & 0xFF == ord('q'):
            break
    
    cap.release()
    cv2.destroyAllWindows()
    print(f"Dataset collection completed for {name}")

if __name__ == "__main__":
    name = input("Enter your name: ")
    create_dataset(name)
