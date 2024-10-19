import cv2
import numpy as np
from imutils.perspective import four_point_transform #เมื่อใช้ imutils.perspective เราจำเป็นต้องลง scipy ด้วยเนื่องจากเป็นโมดูลที่จำเป็นสำหรับ imutils.perspective ทำการลงด้วย pip install scipy



def preprocessing(contents):

    nparr = np.fromstring(contents, np.uint8)

    imRGB = cv2.imdecode(nparr, cv2.IMREAD_COLOR)
    print(imRGB.shape)

    imGray = cv2.cvtColor(imRGB, cv2.COLOR_BGR2GRAY)
    print(imGray.shape)

    blur = cv2.GaussianBlur(imGray, (3, 3), 0)

    image_path = scan_detection(blur, imRGB)

    img = image_smoothening(image_path)

    return img

    
    

def scan_detection(blur, imRGB):

    image_path = '..\\uploads\\scan_detection.jpg'

    height = imRGB.shape[0]
    width = imRGB.shape[1]

    contour_ratio = ((width * height) * 0.5)


    #สร้าง numpy array ที่เก็บพิกัดของ 4 มุมของรูปภาพต้นฉบับ
    #ใช้เป็นค่าเริ่มต้นสำหรับการวาดเส้นขอบ ของรูปภาพต้นฉบับ ในกรณีที่ไม่พบรูปทรงที่ต้องการ
    #โดยเเต่ละมุมจะมีตำเเหน่ง x,y ทั้งหมด 4 มุม
    
    document_contour = np.array([[0, 0], [width, 0], [width, height], [0, height]])
    print(document_contour)
    

    thresh = cv2.threshold(blur, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)[1]


    #contour คือเส้นโค้งที่เชื่อมจุดต่อเนื่องทั้งหมด โดยมีสีหรือความเข้มเท่ากัน
    #contour เป็นเครื่องมือที่มีประโยชน์สำหรับการวิเคราะห์รูปร่าง การตรวจจับเเละการจดจำวัตถุ
    #contour เเต่ละตัวจะเป็น numpy array ของพิกัด x,y ของจุดขอบเขตของรูปร่าง

    contours = cv2.findContours(thresh, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)[0] #ค้นหารูปทรง ซึ่งจะเก็บรูปทรงทั้งหมดที่พบ เป็น List
    
    contours = sorted(contours, key=cv2.contourArea, reverse=True)


    max_area = 0
    for contour in contours:

        area = cv2.contourArea(contour) #คำนวณพื้นที่ของ contour นั้นๆ
        
        if area > contour_ratio:

            peri = cv2.arcLength(contour, True) #หาความยาวเส้นโค้ง
            approx = cv2.approxPolyDP(contour, 0.015 * peri, True)

            if area > max_area and len(approx) == 4:

                max_area = area
                document_contour = approx

    # cv2.drawContours(imRGB, [document_contour], -1, (255, 0, 255), 3)

    warped = four_point_transform(imRGB, document_contour.reshape(4, 2))
    cv2.imwrite(image_path, warped)

    return image_path




def image_smoothening(image_path):

    imRGB = cv2.imread(image_path)
    height = imRGB.shape[0]
    width = imRGB.shape[1]
    
    if height < 4000 and width < 4000:
        imRGB = cv2.resize(imRGB, None, fx=2, fy=2, interpolation=cv2.INTER_CUBIC)
    
    imGray = cv2.cvtColor(imRGB, cv2.COLOR_BGR2GRAY)

    thresh = cv2.threshold(imGray, 127, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)[1]

    thresh = thresh[10:thresh.shape[0] - 10, 10:thresh.shape[1] - 10]

    # cv2.imwrite('..//uploads//thresh.jpg', thresh)

    return thresh