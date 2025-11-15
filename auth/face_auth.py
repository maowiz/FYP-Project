#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import cv2
import os
import shutil
import pickle
import numpy as np
from pathlib import Path

class FaceAuthSystem:
    def __init__(self, dataset_path="dataset", model_path="models"):
        self.dataset_path = dataset_path
        self.model_path = model_path
        
        # Create directories
        Path(self.dataset_path).mkdir(exist_ok=True)
        Path(self.model_path).mkdir(exist_ok=True)
        
        # Haar Cascade - SUPER FAST and lightweight
        self.face_cascade = cv2.CascadeClassifier(
            cv2.data.haarcascades + 'haarcascade_frontalface_default.xml'
        )
        
        self.face_database = {}
        self.load_database()
    
    def extract_features(self, face_img):
        """Extract lightweight features from face"""
        # Convert to grayscale if needed
        if len(face_img.shape) == 3:
            face_gray = cv2.cvtColor(face_img, cv2.COLOR_BGR2GRAY)
        else:
            face_gray = face_img
        
        # Resize to standard size
        face_gray = cv2.resize(face_gray, (100, 100))
        
        # Calculate histogram (very fast feature)
        hist = cv2.calcHist([face_gray], [0], None, [256], [0, 256])
        hist = cv2.normalize(hist, hist).flatten()
        
        # Also use downsampled image as feature
        small = cv2.resize(face_gray, (25, 25)).flatten()
        
        # Combine features
        features = np.concatenate([hist, small / 255.0])
        
        return features
    
    def detect_faces(self, frame):
        """Fast face detection"""
        gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
        # Optimized for speed
        faces = self.face_cascade.detectMultiScale(
            gray, scaleFactor=1.2, minNeighbors=5, 
            minSize=(80, 80), flags=cv2.CASCADE_SCALE_IMAGE
        )
        return faces, gray
    
    def compare_faces(self, features1, features2):
        """Compare two face feature vectors"""
        # Cosine similarity (fast)
        similarity = np.dot(features1, features2) / (
            np.linalg.norm(features1) * np.linalg.norm(features2)
        )
        return similarity * 100  # Convert to percentage
    
    def enhance_image(self, img):
        """Enhance image quality"""
        # Apply CLAHE (Contrast Limited Adaptive Histogram Equalization)
        if len(img.shape) == 3:
            # Convert to LAB color space
            lab = cv2.cvtColor(img, cv2.COLOR_BGR2LAB)
            l, a, b = cv2.split(lab)
            
            # Apply CLAHE to L channel
            clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8))
            l = clahe.apply(l)
            
            # Merge and convert back
            enhanced = cv2.merge([l, a, b])
            enhanced = cv2.cvtColor(enhanced, cv2.COLOR_LAB2BGR)
        else:
            clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8))
            enhanced = clahe.apply(img)
        
        return enhanced
    
    def add_new_person(self, name, num_images=199):
        """Capture and save face images in HIGH QUALITY COLOR"""
        print(f"\n{'═'*60}")
        print(f"  ADDING NEW PERSON: {name.upper()}")
        print(f"{'═'*60}")
        
        person_dir = Path(self.dataset_path) / name
        person_dir.mkdir(exist_ok=True)
        
        cap = cv2.VideoCapture(0)
        if not cap.isOpened():
            print("❌ Error: Cannot access webcam!")
            return False
        
        # ===== HIGH QUALITY CAMERA SETTINGS =====
        cap.set(cv2.CAP_PROP_FRAME_WIDTH, 1280)
        cap.set(cv2.CAP_PROP_FRAME_HEIGHT, 720)
        cap.set(cv2.CAP_PROP_FPS, 30)
        cap.set(cv2.CAP_PROP_FOURCC, cv2.VideoWriter_fourcc(*'MJPG'))
        cap.set(cv2.CAP_PROP_BRIGHTNESS, 128)
        cap.set(cv2.CAP_PROP_CONTRAST, 32)
        cap.set(cv2.CAP_PROP_SATURATION, 64)
        cap.set(cv2.CAP_PROP_AUTOFOCUS, 1)
        cap.set(cv2.CAP_PROP_AUTO_EXPOSURE, 1)
        
        print("\n📸 INSTRUCTIONS:")
        print("  • Keep your face centered in the green box")
        print("  • Slowly move your head: left, right, up, down")
        print("  • Try different expressions")
        print("  • Press 'q' to cancel\n")
        
        # Let camera warm up and adjust
        print("⏳ Warming up camera...")
        for _ in range(30):
            cap.read()
        
        input("👉 Press ENTER to start capturing...")
        
        captured = 0
        frame_count = 0
        
        print("\n📊 Capturing HIGH QUALITY images...")
        
        while captured < num_images:
            ret, frame = cap.read()
            if not ret:
                break
            
            faces, gray = self.detect_faces(frame)
            display = frame.copy()
            
            for (x, y, w, h) in faces:
                # Draw box
                cv2.rectangle(display, (x, y), (x+w, y+h), (0, 255, 0), 2)
                
                # Capture every 3rd frame for variety
                if frame_count % 3 == 0 and len(faces) == 1:
                    # ===== SAVE COLOR IMAGE WITH ENHANCEMENT =====
                    face_roi_color = frame[y:y+h, x:x+w]
                    
                    # Resize to good quality (300x300)
                    face_roi_color = cv2.resize(face_roi_color, (300, 300), 
                                                interpolation=cv2.INTER_CUBIC)
                    
                    # Enhance quality
                    face_roi_color = self.enhance_image(face_roi_color)
                    
                    # Save with high quality JPEG
                    img_path = person_dir / f"{name}_{captured}.jpg"
                    cv2.imwrite(str(img_path), face_roi_color, 
                               [cv2.IMWRITE_JPEG_QUALITY, 95])
                    
                    captured += 1
                    
                    # Progress bar
                    progress = int((captured / num_images) * 40)
                    bar = '█' * progress + '░' * (40 - progress)
                    percent = (captured / num_images) * 100
                    print(f'  [{bar}] {captured}/{num_images} ({percent:.1f}%)', end='\r')
            
            # Display info
            cv2.putText(display, f"Captured: {captured}/{num_images}", 
                       (10, 30), cv2.FONT_HERSHEY_SIMPLEX, 0.7, (0, 255, 0), 2)
            cv2.putText(display, "Press 'q' to quit", 
                       (10, 60), cv2.FONT_HERSHEY_SIMPLEX, 0.6, (0, 0, 255), 2)
            
            if len(faces) == 0:
                cv2.putText(display, "No face detected!", 
                           (10, 90), cv2.FONT_HERSHEY_SIMPLEX, 0.6, (0, 0, 255), 2)
            elif len(faces) > 1:
                cv2.putText(display, "Multiple faces detected!", 
                           (10, 90), cv2.FONT_HERSHEY_SIMPLEX, 0.6, (0, 0, 255), 2)
            
            cv2.imshow('Adding New Person - Face Auth', display)
            
            frame_count += 1
            
            if cv2.waitKey(1) & 0xFF == ord('q'):
                print("\n\n⚠️  Capture cancelled!")
                break
        
        cap.release()
        cv2.destroyAllWindows()
        
        print(f"\n\n✅ Successfully captured {captured} HIGH QUALITY COLOR images!")
        
        if captured >= 50:
            print("🔄 Processing face data...")
            self.process_person_data(name)
            return True
        else:
            print("❌ Not enough images captured. Need at least 50.")
            return False
    
    def process_person_data(self, name):
        """Process and create face template"""
        print("\n" + "─"*60)
        print(f"🧠 PROCESSING: {name}")
        print("─"*60)
        
        person_dir = Path(self.dataset_path) / name
        features_list = []
        
        for img_file in person_dir.glob('*.jpg'):
            img = cv2.imread(str(img_file))
            if img is not None:
                features = self.extract_features(img)
                features_list.append(features)
        
        if len(features_list) > 0:
            # Average all features to create template
            avg_features = np.mean(features_list, axis=0)
            self.face_database[name] = avg_features
            
            # Save database
            self.save_database()
            
            print(f"  ✓ Processed {len(features_list)} images")
            print(f"✅ {name} added to database!")
            print("─"*60)
            return True
        
        return False
    
    def rebuild_database(self):
        """Rebuild face database from all persons"""
        print("\n" + "─"*60)
        print("🔄 REBUILDING DATABASE...")
        print("─"*60)
        
        self.face_database = {}
        
        for person_dir in Path(self.dataset_path).iterdir():
            if person_dir.is_dir():
                self.process_person_data(person_dir.name)
        
        print(f"\n✅ Database rebuilt! Total persons: {len(self.face_database)}")
        print("─"*60)
    
    def delete_person(self, name=None):
        """Delete a person from the system"""
        print(f"\n{'═'*60}")
        print("  🗑️  DELETE PERSON")
        print(f"{'═'*60}")
        
        # Get list of persons
        persons = [d.name for d in Path(self.dataset_path).iterdir() if d.is_dir()]
        
        if not persons:
            print("\n❌ No persons registered yet!")
            return False
        
        # Show list
        print("\n👥 Registered Persons:")
        for i, person in enumerate(persons, 1):
            num_images = len(list((Path(self.dataset_path) / person).glob('*.jpg')))
            print(f"  {i}. {person} ({num_images} images)")
        print(f"  0. Cancel")
        
        # Get choice
        if name is None:
            choice = input(f"\n👉 Enter number to delete (0-{len(persons)}): ").strip()
            
            if not choice.isdigit():
                print("❌ Invalid input!")
                return False
            
            choice = int(choice)
            
            if choice == 0:
                print("❌ Deletion cancelled.")
                return False
            
            if choice < 1 or choice > len(persons):
                print("❌ Invalid choice!")
                return False
            
            name = persons[choice - 1]
        
        # Confirm deletion
        print(f"\n⚠️  WARNING: You are about to delete '{name}'")
        confirm = input("   Type 'YES' to confirm deletion: ").strip()
        
        if confirm != 'YES':
            print("❌ Deletion cancelled.")
            return False
        
        # Delete folder
        person_dir = Path(self.dataset_path) / name
        try:
            shutil.rmtree(person_dir)
            print(f"  ✓ Deleted folder: {person_dir}")
        except Exception as e:
            print(f"❌ Error deleting folder: {e}")
            return False
        
        # Remove from database
        if name in self.face_database:
            del self.face_database[name]
            self.save_database()
            print(f"  ✓ Removed from database")
        
        print(f"\n✅ '{name}' has been completely removed!")
        print("═"*60)
        return True
    
    def save_database(self):
        """Save face database"""
        db_file = Path(self.model_path) / "face_database.pkl"
        with open(db_file, 'wb') as f:
            pickle.dump(self.face_database, f)
    
    def load_database(self):
        """Load face database"""
        db_file = Path(self.model_path) / "face_database.pkl"
        
        if db_file.exists():
            try:
                with open(db_file, 'rb') as f:
                    self.face_database = pickle.load(f)
                if self.face_database:
                    print(f"✅ Database loaded! Persons: {', '.join(self.face_database.keys())}")
                return True
            except Exception as e:
                print(f"⚠️  Error loading database: {e}")
        return False
    
    def authenticate(self):
        """Live face authentication"""
        print(f"\n{'═'*60}")
        print("  🔐 FACE AUTHENTICATION")
        print(f"{'═'*60}")
        
        if not self.face_database:
            print("❌ No registered faces! Please add persons first.")
            return False
        
        print("\n📸 Position your face in the frame")
        print("👉 Press 'q' to exit\n")
        
        input("Press ENTER to start authentication...")
        
        cap = cv2.VideoCapture(0)
        
        # High quality settings
        cap.set(cv2.CAP_PROP_FRAME_WIDTH, 1280)
        cap.set(cv2.CAP_PROP_FRAME_HEIGHT, 720)
        cap.set(cv2.CAP_PROP_FPS, 30)
        cap.set(cv2.CAP_PROP_FOURCC, cv2.VideoWriter_fourcc(*'MJPG'))
        
        if not cap.isOpened():
            print("❌ Cannot access webcam!")
            return False
        
        # Camera warm-up
        for _ in range(15):
            cap.read()
        
        auth_success = False
        auth_name = None
        confidence_threshold = 75
        
        while True:
            ret, frame = cap.read()
            if not ret:
                break
            
            faces, gray = self.detect_faces(frame)
            display = frame.copy()
            
            for (x, y, w, h) in faces:
                face_roi_color = frame[y:y+h, x:x+w]
                
                # Extract features
                features = self.extract_features(face_roi_color)
                
                # Compare with database
                best_match = None
                best_score = 0
                
                for name, stored_features in self.face_database.items():
                    score = self.compare_faces(features, stored_features)
                    if score > best_score:
                        best_score = score
                        best_match = name
                
                # Check threshold
                if best_score >= confidence_threshold:
                    name = best_match
                    color = (0, 255, 0)
                    status = "AUTHORIZED"
                    auth_success = True
                    auth_name = name
                else:
                    name = "Unknown"
                    color = (0, 0, 255)
                    status = "DENIED"
                
                # Display
                cv2.rectangle(display, (x, y), (x+w, y+h), color, 2)
                cv2.putText(display, f"{name}", 
                           (x, y-35), cv2.FONT_HERSHEY_SIMPLEX, 0.8, color, 2)
                cv2.putText(display, f"{status} ({best_score:.0f}%)", 
                           (x, y-10), cv2.FONT_HERSHEY_SIMPLEX, 0.6, color, 2)
            
            cv2.putText(display, "Press 'q' to quit", 
                       (10, 30), cv2.FONT_HERSHEY_SIMPLEX, 0.6, (255, 255, 255), 2)
            
            cv2.imshow('Face Authentication', display)
            
            if cv2.waitKey(1) & 0xFF == ord('q'):
                break
        
        cap.release()
        cv2.destroyAllWindows()
        
        if auth_success:
            print(f"\n✅ AUTHENTICATED! Welcome, {auth_name}!")
        else:
            print("\n❌ AUTHENTICATION FAILED!")
        
        return auth_success
    
    def list_persons(self):
        """Show registered persons"""
        print(f"\n{'═'*60}")
        print("  👥 REGISTERED PERSONS")
        print(f"{'═'*60}")
        
        persons = [d for d in Path(self.dataset_path).iterdir() if d.is_dir()]
        
        if not persons:
            print("  No persons registered yet.")
        else:
            for i, person_dir in enumerate(persons, 1):
                num_images = len(list(person_dir.glob('*.jpg')))
                in_db = "✓" if person_dir.name in self.face_database else "✗"
                print(f"  {i}. {person_dir.name:<20} ({num_images} images) [{in_db}]")
        
        print("═"*60)


def print_header():
    """Print ASCII header"""
    print("\n" + "╔" + "═"*58 + "╗")
    print("║" + " "*15 + "FACE AUTHENTICATION SYSTEM" + " "*17 + "║")
    print("║" + " "*15 + "(HIGH QUALITY EDITION)" + " "*18 + "║")
    print("╚" + "═"*58 + "╝")


def print_menu():
    """Print menu"""
    print("\n┌" + "─"*58 + "┐")
    print("│  [1] 🔐 Authenticate Face (Live)                        │")
    print("│  [2] ➕ Add New Person                                  │")
    print("│  [3] 🗑️  Delete Person                                  │")
    print("│  [4] 🔄 Rebuild Database                                │")
    print("│  [5] 👥 List Registered Persons                         │")
    print("│  [6] 🚪 Exit                                            │")
    print("└" + "─"*58 + "┘")


def main():
    """Main program"""
    print_header()
    
    # Initialize system
    auth_system = FaceAuthSystem(
        dataset_path="dataset",
        model_path="models"
    )
    
    while True:
        print_menu()
        choice = input("\n👉 Enter your choice (1-6): ").strip()
        
        if choice == '1':
            auth_system.authenticate()
        
        elif choice == '2':
            name = input("\n👤 Enter person's name: ").strip()
            if name and name.replace('_', '').isalnum():
                num_imgs = input("📸 Number of images (default 199): ").strip()
                num_imgs = int(num_imgs) if num_imgs.isdigit() else 199
                auth_system.add_new_person(name, num_imgs)
            else:
                print("❌ Invalid name! Use only letters, numbers, and underscores.")
        
        elif choice == '3':
            auth_system.delete_person()
        
        elif choice == '4':
            auth_system.rebuild_database()
        
        elif choice == '5':
            auth_system.list_persons()
        
        elif choice == '6':
            print("\n" + "═"*60)
            print("  👋 Thank you for using Face Authentication System!")
            print("═"*60 + "\n")
            break
        
        else:
            print("\n❌ Invalid choice! Please enter 1-6.")
        
        if choice in ['1', '2', '3', '4', '5']:
            input("\n⏸️  Press ENTER to continue...")


if __name__ == "__main__":
    main()