import face_recognition
import cv2
import numpy as np
import os
import xlwt
from xlwt import Workbook
from datetime import date
import xlrd, xlwt
from xlutils.copy import copy as xl_copy

CurrentFolder = os.getcwd() #Read current folder path

image1 = r'C:\Users\ayush\Desktop\Smart_Attendence_System (1)\UI21CS01_AALEKH.jpeg'
image2 = r'C:\Users\ayush\Desktop\Smart_Attendence_System (1)\UI21CS02_ABHIDNYA.jpeg'
image3 = r'C:\Users\ayush\Desktop\Smart_Attendence_System (1)\UI21CS03_Abhilasha.png'
image4 = r'C:\Users\ayush\Desktop\Smart_Attendence_System (1)\UI21CS04_ABHISHEK.jpeg'
image5 = r'C:\Users\ayush\Desktop\Smart_Attendence_System (1)\UI21CS05_AERIN.jpeg'
image7 = r'C:\Users\ayush\Desktop\Smart_Attendence_System (1)\UI21CS07_AKASH_THAPA.jpeg'
image8 = r'C:\Users\ayush\Desktop\Smart_Attendence_System (1)\UI21CS08_AMIT.jpeg'
image9 = r'C:\Users\ayush\Desktop\Smart_Attendence_System (1)\UI21CS09_ANKIT.jpeg'
image10 = r'C:\Users\ayush\Desktop\Smart_Attendence_System (1)\UI21CS10_ANKIT_GONDHA.jpeg'
image12 = r'C:\Users\ayush\Downloads\UI21CS12_ASHMIT.jpeg'
image13 = r'C:\Users\ayush\Desktop\Smart_Attendence_System (1)\UI21CS13_AYUSH.png'
image14 = r'C:\Users\ayush\Desktop\Smart_Attendence_System (1)\UI21CS14_SWAROOP.png'
image15 = r'C:\Users\ayush\Desktop\Smart_Attendence_System (1)\UI21CS15_DEEPAK_DANGI.jpeg'
image16 = r'C:\Users\ayush\Desktop\Smart_Attendence_System (1)\UI21CS16_DIVYANSH.png'
image17 = r'C:\Users\ayush\Desktop\Smart_Attendence_System (1)\UI21CS17_GAURAV_SINGH.jpeg'
image18 = r'C:\Users\ayush\Desktop\Smart_Attendence_System (1)\UI21CS18_GAURAV_SWAMI.jpeg'
image19 = r'C:\Users\ayush\Desktop\Smart_Attendence_System (1)\UI21CS19_HAMESH.jpeg'
image20 = r'C:\Users\ayush\Desktop\Smart_Attendence_System (1)\UI21CS20_HARSH.jpeg'
image21 = r'C:\Users\ayush\Desktop\Smart_Attendence_System (1)\UI21CS21_HARSHIT.jpg'
image22 = r'C:\Users\ayush\Desktop\Smart_Attendence_System (1)\UI21CS22_CHANAKYA.jpeg'
image24 = r'C:\Users\ayush\Desktop\Smart_Attendence_System (1)\UI21CS24_JEET.jpeg'
image25 = r'C:\Users\ayush\Desktop\Smart_Attendence_System (1)\UI21CS25_JEJURKAR.jpeg'
image26 = r'C:\Users\ayush\Desktop\Smart_Attendence_System (1)\UI21CS26_JINESH.jpeg'
image27 = r'C:\Users\ayush\Desktop\Smart_Attendence_System (1)\UI21CS27_JITANSHU.jpeg'
image28 = r'C:\Users\ayush\Desktop\Smart_Attendence_System (1)\ UI21CS28_KALPAN.png'
image29 = r'C:\Users\ayush\Desktop\Smart_Attendence_System (1)\UI21CS29_KARTIKEY.jpeg'
image30 = r'C:\Users\ayush\Desktop\Smart_Attendence_System (1)\UI21CS30_KRISHNA.jpeg'
image31 = r'C:\Users\ayush\Desktop\Smart_Attendence_System (1)\ UI21CS31_YOGESH.png'
image32 = r'C:\Users\ayush\Desktop\Smart_Attendence_System (1)\UI21CS32_KUNAL.jpeg'
image33 = r'C:\Users\ayush\Desktop\Smart_Attendence_System (1)\UI21CS33_MANAS.jpeg'
image34 = r'C:\Users\ayush\Desktop\Smart_Attendence_System (1)\UI21CS34_NAKUL.jpeg'
image36 = r'C:\Users\ayush\Desktop\Smart_Attendence_System (1)\UI21CS36_NISCHAY.jpeg'
image37 = r'C:\Users\ayush\Desktop\Smart_Attendence_System (1)\UI21CS37_PANKAJ (3).jpeg'











# This is a demo of running face recognition on live video from your webcam. It's a little more complicated than the
# other example, but it includes some basic performance tweaks to make things run a lot faster:
#   1. Process each video frame at 1/4 resolution (though still display it at full resolution)
#   2. Only detect faces in every other frame of video.

# PLEASE NOTE: This example requires OpenCV (the `cv2` library) to be installed only to read from your webcam.
# OpenCV is not required to use the face_recognition library. It's only required if you want to run this
# specific demo. If you have trouble installing it, try any of the other demos that don't require it instead.

# Get a reference to webcam #0 (the default one)
video_capture = cv2.VideoCapture(0)

# Load a sample picture and learn how to recognize it.


# Load a second sample picture and learn how to recognize it.
person1_name = "UI21CS01_AALEKH"
person1_image = face_recognition.load_image_file(image1)
person1_face_encoding = face_recognition.face_encodings(person1_image)[0]

person2_name = "UI21CS02_ABHIDNYA"
person2_image = face_recognition.load_image_file(image2)
person2_face_encoding = face_recognition.face_encodings(person2_image)[0]

person3_name = "UI21CS03_ABHILASHA"
person3_image = face_recognition.load_image_file(image3)
person3_face_encoding = face_recognition.face_encodings(person3_image)[0]

person4_name = "UI21CS04_ABHISHEK"
person4_image = face_recognition.load_image_file(image4)
person4_face_encoding = face_recognition.face_encodings(person4_image)[0]

person5_name = "UI21CS05_AERIN"
person5_image = face_recognition.load_image_file(image5)
person5_face_encoding = face_recognition.face_encodings(person5_image)[0]

person7_name = "UI21CS07_AKASH_THAPA"
person7_image = face_recognition.load_image_file(image7)
person7_face_encoding = face_recognition.face_encodings(person7_image)[0]

person8_name = "UI21CS08_AMIT"
person8_image = face_recognition.load_image_file(image8)
person8_face_encoding = face_recognition.face_encodings(person8_image)[0]

person9_name = "UI21CS09_ANKIT"
person9_image = face_recognition.load_image_file(image9)
person9_face_encoding = face_recognition.face_encodings(person9_image)[0]

person10_name = "UI21CS10_ANKIT_GONDHA"
person10_image = face_recognition.load_image_file(image10)
person10_face_encoding = face_recognition.face_encodings(person10_image)[0]

person12_name = "UI21CS12_ASHMIT"
person12_image = face_recognition.load_image_file(image12)
person12_face_encoding = face_recognition.face_encodings(person12_image)[0]

person13_name = "UI21CS13_AYUSH"
person13_image = face_recognition.load_image_file(image13)
person13_face_encoding = face_recognition.face_encodings(person13_image)[0]

person14_name = "UI21CS14_SWAROOP"
person14_image = face_recognition.load_image_file(image14)
person14_face_encoding = face_recognition.face_encodings(person14_image)[0]

person15_name = "UI21CS15_DEEPAK"
person15_image = face_recognition.load_image_file(image15)
person15_face_encoding = face_recognition.face_encodings(person15_image)[0]

person16_name = "UI21CS16_DIVYANSH"
person16_image = face_recognition.load_image_file(image16)
person16_face_encoding = face_recognition.face_encodings(person16_image)[0]

person17_name = "UI21CS17__GAURAV_SINGH"
person17_image = face_recognition.load_image_file(image17)
person17_face_encoding = face_recognition.face_encodings(person17_image)[0]

person18_name = "UI21CS18_GAURAV_SWAMI"
person18_image = face_recognition.load_image_file(image18)
person18_face_encoding = face_recognition.face_encodings(person18_image)[0]

person19_name = "UI21CS19_HAMESH"
person19_image = face_recognition.load_image_file(image19)
person19_face_encoding = face_recognition.face_encodings(person19_image)[0]

person20_name = "UI21CS20_HARSH"
person20_image = face_recognition.load_image_file(image20)
person20_face_encoding = face_recognition.face_encodings(person20_image)[0]

person21_name = "UI21CS21_HARSHIT_PANJWANI"
person21_image = face_recognition.load_image_file(image21)
person21_face_encoding = face_recognition.face_encodings(person21_image)[0]

person22_name = "UI21CS22_CHANAKYA"
person22_image = face_recognition.load_image_file(image22)
person22_face_encoding = face_recognition.face_encodings(person22_image)[0]

person24_name = "UI21CS24_JEET"
person24_image = face_recognition.load_image_file(image24)
person24_face_encoding = face_recognition.face_encodings(person24_image)[0]

person25_name = "UI21CS25_JEJURKAR"
person25_image = face_recognition.load_image_file(image25)
person25_face_encoding = face_recognition.face_encodings(person25_image)[0]

person26_name = "UI21CS26_JINESH"
person26_image = face_recognition.load_image_file(image26)
person26_face_encoding = face_recognition.face_encodings(person26_image)[0]

person27_name = "UI21CS27_JITANSHU"
person27_image = face_recognition.load_image_file(image27)
person27_face_encoding = face_recognition.face_encodings(person27_image)[0]

person28_name = "UI21CS28_KALPAN"
person28_image = face_recognition.load_image_file(image28)
person28_face_encoding = face_recognition.face_encodings(person28_image)[0]

person29_name = "UI21CS29_KARTIKEY"
person29_image = face_recognition.load_image_file(image29)
person29_face_encoding = face_recognition.face_encodings(person29_image)[0]

person30_name = "UI21CS30_KRISHNA"
person30_image = face_recognition.load_image_file(image30)
person30_face_encoding = face_recognition.face_encodings(person30_image)[0]

person31_name = "UI21CS31_YOGESH"
person31_image = face_recognition.load_image_file(image31)
person31_face_encoding = face_recognition.face_encodings(person31_image)[0]

person32_name = "UI21CS32_KUNAL"
person32_image = face_recognition.load_image_file(image32)
person32_face_encoding = face_recognition.face_encodings(person32_image)[0]

person33_name = "UI21CS33_MANAS"
person33_image = face_recognition.load_image_file(image33)
person33_face_encoding = face_recognition.face_encodings(person33_image)[0]

person34_name = "UI21CS34_NAKUL"
person34_image = face_recognition.load_image_file(image34)
person34_face_encoding = face_recognition.face_encodings(person34_image)[0]

person36_name = "UI21CS36_NISCHAY"
person36_image = face_recognition.load_image_file(image36)
person36_face_encoding = face_recognition.face_encodings(person36_image)[0]

person37_name = "UI21CS37_PANKAJ"
person37_image = face_recognition.load_image_file(image37)
person37_face_encoding = face_recognition.face_encodings(person37_image)[0]






# Create arrays of known face encodings and their names
known_face_encodings = [
    person1_face_encoding,
    person2_face_encoding,
    person3_face_encoding,
    person4_face_encoding,
    person5_face_encoding,
    person7_face_encoding,
    person8_face_encoding,
    person9_face_encoding,
    person10_face_encoding,
    person12_face_encoding,
    person13_face_encoding,
    person14_face_encoding,
    person15_face_encoding,
    person16_face_encoding,
    person17_face_encoding,
    person18_face_encoding,
    person19_face_encoding,
    person20_face_encoding,
    person21_face_encoding,
    person22_face_encoding,
    person24_face_encoding,
    person25_face_encoding,
    person26_face_encoding,
    person27_face_encoding,
    person29_face_encoding,
    person30_face_encoding,
    person31_face_encoding,
    person32_face_encoding,
    person33_face_encoding,
    person34_face_encoding,
    person36_face_encoding,
    person37_face_encoding
    
    
    
  
    
]
known_face_names = [
    person1_name,
    person2_name,
    person3_name,
    person4_name,
    person5_name,
    person7_name,
    person8_name,
    person9_name,
    person10_name,
    person12_name,
    person13_name,
    person14_name,
    person15_name,
    person16_name,
    person17_name,
    person18_name,
    person19_name,
    person20_name,
    person21_name,
    person22_name,
    person24_name,
    person25_name,
    person26_name,
    person27_name,
    person28_name,
    person29_name,
    person30_name,
    person31_name,
    person32_name,
    person33_name,
    person34_name,
    person36_name,
    person37_name
    
    
    
    
]

# Initialize some variables
face_locations = []
face_encodings = []
face_names = []
process_this_frame = True

rb = xlrd.open_workbook('SMART_ATTENDANCE.xls', formatting_info=True) 
wb = xl_copy(rb)
inp = input('Please give current subject lecture name')
face_names = []
sheet1 = wb.add_sheet(inp)
sheet1.write(0, 0, 'Name/Date')
sheet1.write(0, 1, str(date.today()))
row=1
col=0
already_attendence_taken = ""
while True:
            # Grab a single frame of video
            ret, frame = video_capture.read()

            # Resize frame of video to 1/4 size for faster face recognition processing
            small_frame = cv2.resize(frame, (0, 0), fx=0.25, fy=0.25)

            # Convert the image from BGR color (which OpenCV uses) to RGB color (which face_recognition uses)
            rgb_small_frame = np.ascontiguousarray(small_frame[:, :, ::-1])
            face12 =[]
            # Only process every other frame of video to save time
            if process_this_frame:
                # Find all the faces and face encodings in the current frame of video
                face_locations = face_recognition.face_locations(rgb_small_frame)
                face_encodings = face_recognition.face_encodings(rgb_small_frame, face_locations)

                for face_encoding in face_encodings:
                    # See if the face is a match for the known face(s)
                    matches = face_recognition.compare_faces(known_face_encodings, face_encoding)
                    name = "Unknown"

                    # # If a match was found in known_face_encodings, just use the first one.
                    # if True in matches:
                    #     first_match_index = matches.index(True)
                    #     name = known_face_names[first_match_index]

                    # Or instead, use the known face with the smallest distance to the new face
                    face_distances = face_recognition.face_distance(known_face_encodings, face_encoding)
                    best_match_index = np.argmin(face_distances)
                    if matches[best_match_index]:
                        name = known_face_names[best_match_index]

                    print(name)
                    print(face_names)
                    face12.append(name)
                    if(( name not in face_names) and (name != "Unknown") and (already_attendence_taken != name)):
                     face_names.append(name)
                     sheet1.write(row, col, name )
                     col =col+1
                     sheet1.write(row, col, "Present" )
                     row = row+1
                     col = 0
                     print("attendence taken")
                     wb.save('SMART_ATTENDANCE.xls')
                     already_attendence_taken = name
                    else:
                     print("next student")
                        
            process_this_frame = not process_this_frame


            # Display the results
            for (top, right, bottom, left), name in zip(face_locations, face12):
                # Scale back up face locations since the frame we detected in was scaled to 1/4 size
                top *= 4
                right *= 4
                bottom *= 4
                left *= 4

                # Draw a box around the face
                cv2.rectangle(frame, (left, top), (right, bottom), (0, 0, 255), 2)

                # Draw a label with a name below the face
                cv2.rectangle(frame, (left, bottom - 35), (right, bottom), (0, 0, 255), cv2.FILLED)
                font = cv2.FONT_HERSHEY_DUPLEX
                print(name)
                cv2.putText(frame, name, (left + 6, bottom - 6), font, 1.0, (255, 255, 255), 1)

            # Display the resulting image
            cv2.imshow('Video', frame)

            # Hit 'q' on the keyboard to quit!
            if cv2.waitKey(1) & 0xff==ord('q'):   
                print("data save")
                break

# Release handle to the webcam
video_capture.release()
cv2.destroyAllWindows()
