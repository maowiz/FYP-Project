import face_recognition
import os
import pickle

def generate_face_encodings(images_dir="faces"):
    """
    Generate face encodings from images in the specified directory
    Args:
        images_dir: Directory containing face images (named as person_name.jpg/png)
    """
    known_face_encodings = []
    known_face_names = []
    
    # Create faces directory if it doesn't exist
    if not os.path.exists(images_dir):
        os.makedirs(images_dir)
        print(f"Created {images_dir} directory. Please add face images named as person_name.jpg")
        return
    
    # Load images and generate encodings
    for image_file in os.listdir(images_dir):
        if image_file.lower().endswith(('.png', '.jpg', '.jpeg')):
            # Get person name from filename (remove extension)
            person_name = os.path.splitext(image_file)[0]
            
            # Load image file
            image_path = os.path.join(images_dir, image_file)
            image = face_recognition.load_image_file(image_path)
            
            # Get face encodings
            face_encodings = face_recognition.face_encodings(image)
            
            if len(face_encodings) > 0:
                # Use first face found in image
                face_encoding = face_encodings[0]
                known_face_encodings.append(face_encoding)
                known_face_names.append(person_name)
                print(f"Added face encoding for {person_name}")
            else:
                print(f"No face found in {image_file}")
    
    if known_face_encodings:
        # Save encodings to file
        data = {
            "encodings": known_face_encodings,
            "names": known_face_names
        }
        with open("face_encodings.pkl", "wb") as f:
            pickle.dump(data, f)
        print("\nFace encodings saved successfully!")
        print(f"Total faces encoded: {len(known_face_names)}")
        print("Names:", ", ".join(known_face_names))
    else:
        print("\nNo faces were encoded. Please add images to the faces directory.")

if __name__ == "__main__":
    generate_face_encodings()
