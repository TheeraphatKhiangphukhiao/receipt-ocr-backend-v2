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

    remove_shadow = shadow_extraction(image_path)

    dst_dilate = image_smoothening(remove_shadow)

    return dst_dilate

    
    

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




def image_smoothening(remove_shadow):
    
    imGray = cv2.cvtColor(remove_shadow, cv2.COLOR_BGR2GRAY)

    #blur = cv2.GaussianBlur(imGray, (3, 3), 0)

    thresh = cv2.threshold(imGray, 127, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)[1]

    #cv2.imwrite('..\\uploads\\thresh.jpg', thresh)


    kernel = np.array([[0, 1, 0], [1, 1, 1], [0, 1, 0]], np.uint8)
    print(kernel)

    dst_dilate = cv2.dilate(thresh, kernel, iterations=1)

    height, width = dst_dilate.shape
    dst_dilate = dst_dilate[10:height - 10, 10:width - 10]

    #cv2.imwrite('..\\uploads\\dst_dilate.jpg', dst_dilate)

    return dst_dilate



def shadow_extraction(image_path):

    imRGB = cv2.imread(image_path)

    bgr_planes = cv2.split(imRGB) #ทำการเเยก channel ของสี

    res_md = []
    result_planes = []

    for plane in bgr_planes:

        dilated_img = cv2.dilate(plane, np.ones((7, 7), np.uint8))
        bg_img = cv2.medianBlur(dilated_img, 21)
        
        diff_img = 255-cv2.absdiff(plane, bg_img)

        res_md.append(bg_img)
        result_planes.append(diff_img)

    remove_shadow = cv2.merge(result_planes) #ทำการรวมสีเเต่ละ channel เข้าด้วยกัน

    #cv2.imwrite('..\\uploads\\remove_shadow.jpg', remove_shadow)

    return remove_shadow
