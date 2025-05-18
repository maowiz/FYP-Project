import face_recognition
import os
import pickle
import numpy as np

def train_face_recognition():
    """
    Train face recognition model by encoding all faces in the dataset
    Returns:
        dict: Dictionary containing face encodings and names
    """
    known_face_encodings = []
    known_face_names = []
    dataset_dir = "dataset"

    # Loop through each person's directory
    for person_name in os.listdir(dataset_dir):
        person_dir = os.path.join(dataset_dir, person_name)
        if os.path.isdir(person_dir):
            # Process each image in person's directory
            for image_name in os.listdir(person_dir):
                image_path = os.path.join(person_dir, image_name)
                try:
                    # Load and encode face
                    face_image = face_recognition.load_image_file(image_path)
                    face_encodings = face_recognition.face_encodings(face_image)
                    
                    if face_encodings:
                        known_face_encodings.append(face_encodings[0])
                        known_face_names.append(person_name)
                        print(f"Processed {image_path}")
                except Exception as e:
                    print(f"Error processing {image_path}: {str(e)}")

    # Save the encodings
    data = {
        "encodings": known_face_encodings,
        "names": known_face_names
    }
    
    with open("face_encodings.pkl", "wb") as f:
        pickle.dump(data, f)
    
    print("Training completed and model saved!")
    return data

if __name__ == "__main__":
    train_face_recognition()
