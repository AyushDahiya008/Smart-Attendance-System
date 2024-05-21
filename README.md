This project is a Smart Attendance System that uses facial recognition to mark the attendance of students automatically. By leveraging several Python libraries, the system can detect faces in real-time using a webcam, match these faces with a preloaded set of known images, and log attendance data into an Excel file. This system provides an efficient and accurate way to track attendance, eliminating the need for manual processes.

OUTPUT : 


https://github.com/AyushDahiya008/Smart-Attendance-System/assets/158088706/69e7d484-a20d-43b3-a945-f02bc227528f



Libraries Used : 


[1] face_recognition

Explanation: This library is used for face detection and recognition. It allows the loading of images, encoding of faces, and comparing face encodings to identify individuals.
Usage: It is the core library used to identify and recognize faces from the webcam feed and the preloaded images.


[2] cv2 (OpenCV)

Explanation: OpenCV is an open-source computer vision library that allows image and video processing.
Usage: It is used to capture video from the webcam, resize frames for faster processing, and display the video with detected faces.


[3] numpy

Explanation: NumPy is a fundamental package for scientific computing with Python. It provides support for arrays and matrices, along with a collection of mathematical functions.
Usage: It is used for handling and manipulating image data efficiently, such as resizing frames and converting image formats.


[4] os

Explanation: The os module provides a way of using operating system-dependent functionality, such as reading or writing to the filesystem.
Usage: It is used to read the current folder path and handle file paths for loading images.


[5] xlwt

Explanation: xlwt is a library for writing data to Excel files.
Usage: It is used to create a new Excel workbook and write attendance data to it.


[6] xlrd

Explanation: xlrd is a library for reading data from Excel files.
Usage: It is used to read an existing Excel file to copy its contents when updating attendance data.


[7] xlutils.copy

Explanation: This module allows copying and editing Excel files read by xlrd.
Usage: It is used to copy the structure and data of an existing Excel workbook to update attendance records.



Importance of the Project : 
The Smart Attendance System automates the attendance marking process, which is traditionally done manually. This automation ensures higher accuracy, saves time, and minimizes errors. It is particularly useful in educational institutions, workplaces, and any environment where regular attendance tracking is necessary. By using facial recognition, the system also enhances security and ensures that the attendance data is reliable and tamper-proof.




How It Works :
1) Initialization
The script begins by capturing video from the default webcam using OpenCV (cv2.VideoCapture(0)).

2)Loading Known Images
A set of known images with corresponding student names is loaded using face_recognition.load_image_file. Each image is then encoded into a format that can be used for comparison (face_recognition.face_encodings).

3)Setting Up Attendance Sheet
The Excel workbook is prepared for logging attendance. If an attendance sheet for the current subject does not exist, a new sheet is created. The date and student names are initialized.

4)Face Detection and Recognition
The system processes video frames captured by the webcam. For every frame, it detects faces and encodes them.
Each detected face encoding is compared with the known face encodings using face_recognition.compare_faces and face_recognition.face_distance.
If a match is found, the corresponding student name is retrieved.

5)Logging Attendance
If a detected face matches a known face and has not been marked as present already, the system logs the student's name as "Present" in the Excel sheet.

6)Displaying Results
The system displays the video feed with rectangles drawn around detected faces and labels indicating the recognized student names.
Saving Data
The attendance data is saved to the Excel workbook after each detection to ensure that the data is up-to-date and not lost if the program is interrupted.
User Interaction
The system runs in a loop until the user decides to quit (by pressing 'q'), ensuring continuous attendance tracking during the session.




Conclusion :
This Smart Attendance System demonstrates an efficient application of computer vision and data management libraries in Python. By automating attendance tracking, it ensures accuracy and saves time, making it a valuable tool for various organizations. The integration of face recognition technology adds a layer of security and convenience, making the process seamless and reliable.
