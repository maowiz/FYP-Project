import cv2

cap = cv2.VideoCapture(0)

if not cap.isOpened():
    print("Cannot open camera")
    exit()

try:
    while True:
        ret, frame = cap.read()
        if not ret:
            print("Can't receive frame (stream end?). Exiting ...")
            break

        cv2.imshow('Test Frame', frame)
        if cv2.waitKey(1) == ord('q'):
            break
except KeyboardInterrupt:
    print("Interrupted by user.")

cap.release()
cv2.destroyAllWindows()
