import win32com.client
from cvzone.HandTrackingModule import HandDetector
import cv2
import time

# Open PowerPoint presentation
Application = win32com.client.Dispatch("PowerPoint.Application")
Presentation = Application.Presentations.Open("C:\\Users\\Lenovo\\Desktop\\presentation.ppt")
print(Presentation.Name)
Presentation.SlideShowSettings.Run()

# Parameters
width, height = 900, 720
gestureThreshold = 300
pause = False

# Camera Setup
cap = cv2.VideoCapture(0)
cap.set(3, width)
cap.set(4, height)

# Hand Detector
detectorHand = HandDetector(detectionCon=0.8, maxHands=1)

# Variables
lastActionTime = time.time()
delay = 3  # 3 seconds delay for action cooldown
annotations = []
annotationNumber = -1
annotationStart = False
actionCooldown = False

# Helper function to clear all annotations
def clearAnnotations(slide):
    shapes = slide.Shapes
    for shape in shapes:
        if shape.Type == 9:  # msoLine
            shape.Delete()

while True:
    # Get image frame
    success, img = cap.read()
    imgCurrent = img.copy()
    # Find the hand and its landmarks
    hands, img = detectorHand.findHands(img)  # with draw

    if hands:
        hand = hands[0]
        cx, cy = hand["center"]
        lmList = hand["lmList"]  # List of 21 Landmark points
        fingers = detectorHand.fingersUp(hand)  # List of which fingers are up
        indexFinger = lmList[8][:2]  # Tip of the index finger

        # Print detected finger states
        print(fingers)

        currentTime = time.time()

        # If current time exceeds the cooldown delay, reset actionCooldown
        if currentTime - lastActionTime > delay:
            actionCooldown = False

        # Ensure actions are performed only if not in cooldown
        if not actionCooldown:
            if cy <= gestureThreshold:  # If hand is at the height of the face
                # Check if gesture for next slide (all fingers up) is stable
                if fingers == [1, 1, 1, 1, 1]:
                    print("Next Slide")
                    Presentation.SlideShowWindow.View.Next()
                    lastActionTime = currentTime
                    actionCooldown = True
                    annotations = []
                    annotationStart = False

                # Check if gesture for previous slide (thumb and pinky up) is stable
                elif fingers == [1, 0, 0, 0, 1]:
                    print("Previous Slide")
                    Presentation.SlideShowWindow.View.Previous()
                    lastActionTime = currentTime
                    actionCooldown = True
                    annotations = []
                    annotationStart = False

                # Check if gesture for drawing (index finger up) is stable
                elif fingers == [0, 1, 0, 0, 0]:
                    if not annotationStart:
                        annotationStart = True
                        annotationNumber += 1
                        annotations.append([])
                    annotations[-1].append(indexFinger)
                    cv2.circle(imgCurrent, indexFinger, 12, (0, 0, 255), cv2.FILLED)

                    # Draw on PowerPoint slide immediately
                    slide = Presentation.SlideShowWindow.View.Slide
                    shapes = slide.Shapes
                    if len(annotations[-1]) > 1:
                        pt1 = annotations[-1][-2]
                        pt2 = annotations[-1][-1]
                        shapes.AddLine(pt1[0], pt1[1], pt2[0], pt2[1]).Line.ForeColor.RGB = 255

                # Check if gesture for full erase (index, middle, ring fingers up) is stable
                elif fingers == [0, 1, 1, 1, 0]:
                    print("Full Erase")
                    clearAnnotations(Presentation.SlideShowWindow.View.Slide)
                    annotations = []
                    annotationStart = False
                    lastActionTime = currentTime
                    actionCooldown = True

                # Check if gesture for pause/resume (index and middle fingers up) is stable
                elif fingers == [0, 1, 1, 0, 0]:
                    if not actionCooldown:
                        if not pause:
                            print("Pause Slide Show")
                            Presentation.SlideShowWindow.View.State = 2  # Pause
                            pause = True
                        else:
                            print("Resume Slide Show")
                            Presentation.SlideShowWindow.View.State = 1  # Resume
                            pause = False
                        lastActionTime = currentTime
                        actionCooldown = True

    # Draw on the video feed
    for annotation in annotations:
        for j in range(len(annotation)):
            if j != 0:
                cv2.line(imgCurrent, annotation[j - 1], annotation[j], (0, 0, 200), 12)

    cv2.imshow("Image", img)

    key = cv2.waitKey(1)
    if key == ord('q'):
        break

cap.release()
cv2.destroyAllWindows()