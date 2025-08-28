import face_recognition
import cv2
import pickle
import numpy as np
import time
from typing import Optional, Tuple
from pathlib import Path

class FaceAuthenticator:
    def __init__(self, model_path: str = None, speech=None):
        """
        Initialize face authenticator with speech integration.
        Args:
            model_path: Path to the trained face encodings file.
            speech: Speech object for audio feedback (optional, from virtual assistant).
        """
        print("[FaceAuth] Initializing FaceAuthenticator...")
        if model_path is None:
            # Use the absolute path to the face_encodings.pkl file
            model_path = Path(__file__).parent.absolute() / "face_encodings.pkl"
        print(f"[FaceAuth] Model path: {model_path}")
        self.known_face_encodings = []
        self.known_face_names = []
        self.speech = speech
        self.model_path = model_path
        if not self.load_encodings():
            raise Exception("Failed to load face encodings. Please train the model first.")
        self.frame_resize_scale = 1.0  # Default: no resizing; set < 1.0 for optimization
        self.min_brightness = 50  # Minimum frame brightness (0-255) to detect low light
        print("[FaceAuth] FaceAuthenticator initialized.")

    def load_encodings(self) -> bool:
        """Load the trained face encodings from file."""
        print("[FaceAuth] Loading face encodings...")
        try:
            if not self.model_path.exists():
                print(f"[FaceAuth] Error: Face encodings file '{self.model_path}' not found.")
                if self.speech:
                    self.speech.speak("Face encodings file not found. Please train the model first.")
                return False
            with open(self.model_path, "rb") as f:
                data = pickle.load(f)
                self.known_face_encodings = data["encodings"]
                self.known_face_names = data["names"]
            print(f"[FaceAuth] Loaded {len(self.known_face_names)} face encodings.")
            if self.speech:
                self.speech.speak("Face recognition model loaded successfully.")
            return True
        except Exception as e:
            print(f"[FaceAuth] Error loading face encodings: {str(e)}")
            if self.speech:
                self.speech.speak(f"Error loading face recognition model: {str(e)}")
            return False

    def check_brightness(self, frame: np.ndarray) -> bool:
        """
        Check if the frame is too dark for reliable face detection.
        Args:
            frame: Input frame from webcam.
        Returns:
            bool: True if brightness is sufficient, False if too dark.
        """
        gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
        avg_brightness = np.mean(gray)
        print(f"[FaceAuth] Frame brightness: {avg_brightness:.2f}")
        return avg_brightness >= self.min_brightness

    def preprocess_frame(self, frame: np.ndarray) -> np.ndarray:
        """
        Preprocess frame for face detection: optional resize and convert to RGB.
        Args:
            frame: Input frame from webcam.
        Returns:
            Processed frame.
        """
        if self.frame_resize_scale < 1.0:
            frame = cv2.resize(frame, (0, 0), fx=self.frame_resize_scale, fy=self.frame_resize_scale)
        return cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)

    def authenticate(self, confidence_threshold: float = 0.42, timeout_seconds: float = 7.0) -> Tuple[Optional[str], str]:
        """
        Authenticate a person using face recognition and return their name and status.
        Args:
            confidence_threshold: Minimum confidence level for authentication (lower = stricter).
            timeout_seconds: Maximum time to attempt authentication.
        Returns:
            Tuple[str, str]: (Name of the recognized person or None, status message).
        """
        print("[FaceAuth] Starting authentication...")
        # Initialize webcam
        cap = cv2.VideoCapture(0)
        if not cap.isOpened():
            print("[FaceAuth] Error: Could not open webcam.")
            if self.speech:
                self.speech.speak("Could not open webcam. Please check your camera.")
            return None, "webcam_error"
        
        # Variables for verification
        consecutive_matches = 0
        required_matches = 3  # Need 3 consecutive matches to confirm identity
        last_matched_name = None
        verification_frames = []  # Store last few frames for verification
        start_time = time.time()
        frame_count = 0
        process_every_n_frames = 1  # Process every frame; increase to 2 for optimization
        
        try:
            while time.time() - start_time < timeout_seconds:
                ret, frame = cap.read()
                if not ret:
                    print("[FaceAuth] Warning: Failed to capture frame.")
                    if self.speech:
                        self.speech.speak("Failed to capture frame. Please ensure camera is working.")
                    continue
                
                frame_count += 1
                print(f"[FaceAuth] Frame {frame_count}, Time elapsed: {time.time() - start_time:.2f}s")
                # Skip processing for some frames to improve performance (if enabled)
                if frame_count % process_every_n_frames != 0:
                    continue
                
                # Check brightness
                if not self.check_brightness(frame):
                    cv2.putText(frame, "Too dark, please improve lighting", (10, 30),
                               cv2.FONT_HERSHEY_SIMPLEX, 0.9, (0, 0, 255), 2)
                    if self.speech and frame_count % 10 == 0:
                        self.speech.speak("It's too dark. Please improve lighting.")
                    consecutive_matches = 0
                    verification_frames.clear()
                    cv2.imshow("Authentication", frame)
                    cv2.waitKey(1)
                    continue
                
                # Preprocess frame
                rgb_frame = self.preprocess_frame(frame)
                
                # Scale face locations if resized
                scale = 1.0 / self.frame_resize_scale
                face_locations = face_recognition.face_locations(rgb_frame)
                face_locations = [(int(top * scale), int(right * scale), int(bottom * scale), int(left * scale))
                                 for top, right, bottom, left in face_locations]
                
                # Handle face detection scenarios
                if len(face_locations) == 0:
                    cv2.putText(frame, "No face detected", (10, 30),
                               cv2.FONT_HERSHEY_SIMPLEX, 0.9, (0, 0, 255), 2)
                    if self.speech and frame_count % 10 == 0:
                        self.speech.speak("Please face the camera.")
                    consecutive_matches = 0
                    verification_frames.clear()
                elif len(face_locations) > 1:
                    cv2.putText(frame, "Multiple faces detected", (10, 30),
                               cv2.FONT_HERSHEY_SIMPLEX, 0.9, (0, 0, 255), 2)
                    if self.speech and frame_count % 10 == 0:
                        self.speech.speak("Please ensure only one person is in the frame.")
                    consecutive_matches = 0
                    verification_frames.clear()
                else:
                    # Single face detected
                    face_encodings = face_recognition.face_encodings(frame, [face_locations[0]])
                    if not face_encodings:
                        continue
                    face_encoding = face_encodings[0]
                    
                    # Store frame for verification
                    verification_frames.append(frame)
                    if len(verification_frames) > required_matches:
                        verification_frames.pop(0)
                    
                    # Compare with known faces
                    face_distances = face_recognition.face_distance(self.known_face_encodings, face_encoding)
                    if len(face_distances) == 0:
                        continue
                    
                    best_match_index = np.argmin(face_distances)
                    best_match_distance = face_distances[best_match_index]
                    
                    if best_match_distance < confidence_threshold:
                        name = self.known_face_names[best_match_index]
                        
                        # Check consistency
                        if name == last_matched_name:
                            consecutive_matches += 1
                        else:
                            consecutive_matches = 1
                            last_matched_name = name
                        
                        # Draw feedback
                        y1, x2, y2, x1 = face_locations[0]
                        color = (0, 255, 0) if consecutive_matches >= required_matches else (0, 255, 255)
                        cv2.rectangle(frame, (x1, y1), (x2, y2), color, 2)
                        cv2.putText(frame, f"Verifying: {name} ({consecutive_matches}/{required_matches})",
                                   (x1, y1-10), cv2.FONT_HERSHEY_SIMPLEX, 0.9, color, 2)
                        
                        # Draw progress bar
                        progress = min(1.0, (time.time() - start_time) / timeout_seconds)
                        bar_width = int(200 * progress)
                        cv2.rectangle(frame, (10, frame.shape[0] - 30), (10 + bar_width, frame.shape[0] - 10),
                                     (0, 255, 0), -1)
                        
                        # Final verification
                        if consecutive_matches >= required_matches:
                            all_verified = True
                            for verify_frame in verification_frames[-required_matches:]:
                                verify_locations = face_recognition.face_locations(verify_frame)
                                if len(verify_locations) == 1:
                                    verify_encoding = face_recognition.face_encodings(verify_frame, [verify_locations[0]])[0]
                                    verify_distance = face_recognition.face_distance([face_encoding], verify_encoding)[0]
                                    if verify_distance >= confidence_threshold:
                                        all_verified = False
                                        break
                            
                            if all_verified:
                                cv2.putText(frame, f"Welcome, {name}!", (10, 30),
                                           cv2.FONT_HERSHEY_SIMPLEX, 1.2, (0, 255, 0), 2)
                                cv2.imshow("Authentication", frame)
                                cv2.waitKey(1000)  # Show success for 1 second
                                if self.speech:
                                    self.speech.speak(f"Welcome, {name}!")
                                cap.release()
                                cv2.destroyAllWindows()
                                return name, "success"
                    else:
                        consecutive_matches = 0
                        cv2.putText(frame, "Face not recognized", (10, 30),
                                   cv2.FONT_HERSHEY_SIMPLEX, 0.9, (0, 0, 255), 2)
                        if self.speech and frame_count % 10 == 0:
                            self.speech.speak("Face not recognized. Please try again.")
                
                # Display frame
                cv2.imshow("Authentication", frame)
                if cv2.waitKey(1) & 0xFF == ord("q"):
                    cap.release()
                    cv2.destroyAllWindows()
                    return None, "user_cancelled"
            
            # Timeout reached
            print("[FaceAuth] Authentication timed out.")
            if self.speech:
                self.speech.speak("Authentication timed out. Please try again.")
            cap.release()
            cv2.destroyAllWindows()
            return None, "timeout"
        
        except Exception as e:
            print(f"[FaceAuth] Error during authentication: {str(e)}")
            if self.speech:
                self.speech.speak(f"Error during authentication: {str(e)}")
            cap.release()
            cv2.destroyAllWindows()
            return None, "error"
        finally:
            # Ensure camera is released even if an unexpected error occurs
            if 'cap' in locals() and cap.isOpened():
                cap.release()
            cv2.destroyAllWindows()

    def verify_user(self, speech=None) -> Tuple[Optional[str], str]:
        """
        Function to be called by virtual assistant for user verification.
        Args:
            speech: Speech object for audio feedback (optional).
        Returns:
            Tuple[str, str]: (Name of the verified user or None, status message).
        """
        if speech and not self.speech:
            self.speech = speech
        return self.authenticate()

if __name__ == "__main__":
    recognized_name, status = FaceAuthenticator().verify_user()
    if recognized_name:
        print(f"Welcome back, {recognized_name}! (Status: {status})")
    else:
        print(f"Authentication failed - User not recognized (Status: {status})")