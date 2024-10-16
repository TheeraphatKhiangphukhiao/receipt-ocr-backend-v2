from typing import Union
from fastapi import FastAPI, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse #สำหรับส่งไฟล์กลับไปเเบบ asynchronously
from pydantic import BaseModel
from openpyxl import Workbook #library ที่ออกเเบบมาเพื่อทำงานกับ excel โดยเฉพาะ โดยสามารถทำงานต่างๆที่เกี่ยวข้องกับ excel ได้เเทบทุกอย่าง
from openpyxl.styles import NamedStyle, Font, PatternFill, Border, Side #สำหรับตกเเต่ง font ในตารางข้อมูล 
import pytesseract
import re
import pre_image #นำเข้าโมดูลจากไฟล์อื่นๆ



app = FastAPI()


origins = [ #รายชื่อเเหล่งที่มาที่ควรได้รับอนุญาตให้ทำการร้องขอข้ามเเหล่งที่มา
    "https://receipt-ocr-app-8c0a9.web.app",
    "http://localhost:5173",
    "http://202.28.34.197:8027"
]

app.add_middleware( #CORS or Cross-Origin Resource Sharing
    CORSMiddleware,
    allow_origins=origins,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)



class Item(BaseModel):
    item1: str
    item2: str
    item3: str
    item4: str
    item5: str
    item6: str
    item7: str
    item8: str

class Result(BaseModel):
    result: list[Item]



@app.post("/write/excel")
async def write_excel(result: Result):

    file_path = "..\\write-file\\receipt.xlsx"


    try:
        print(result)
        rows = result.result
        print(rows)


        file = Workbook() #Workbook คือ object ที่หน้าตาเหมือนกับ Excel สร้างมาเพื่อเป็นตัวเเทนของไฟล์ Excel ที่เราต้องการ อ่าน-เขียน
        
        print(file)
        print(file.worksheets) #ถามถึงเเผ่นงาน
        print(file.sheetnames) #ถามชื่อเเผ่นงาน

        sheet = file['Sheet'] 

        custom_style = NamedStyle(name='custom style')
        
        custom_style.font = Font(name='TH SarabunPSK', size=14, bold=True) #กำหนด font
        custom_style.fill = PatternFill('solid', fgColor='FFFFCC')
        custom_style.border = Border(left=Side(border_style='thin', 
                                            color='B2B2B2'), 
                                            right=Side(border_style='thin', 
                                                        color='B2B2B2'), 
                                                        top=Side(border_style='thin', 
                                                                color='B2B2B2'),
                                                                bottom=Side(border_style='thin',
                                                                            color='B2B2B2')
                                    )

        i = 1
        for row in rows:

            sheet.cell(i, 1).value = row.item1 ; sheet.cell(i, 1).style = custom_style
            sheet.cell(i, 2).value = row.item2 ; sheet.cell(i, 2).style = custom_style
            sheet.cell(i, 3).value = row.item3 ; sheet.cell(i, 3).style = custom_style
            sheet.cell(i, 4).value = row.item4 ; sheet.cell(i, 4).style = custom_style
            sheet.cell(i, 5).value = row.item5 ; sheet.cell(i, 5).style = custom_style
            sheet.cell(i, 6).value = row.item6 ; sheet.cell(i, 6).style = custom_style
            sheet.cell(i, 7).value = row.item7 ; sheet.cell(i, 7).style = custom_style
            sheet.cell(i, 8).value = row.item8 ; sheet.cell(i, 8).style = custom_style

            i += 1
        

        cell_length = [0, 0, 0, 0, 0, 0, 0, 0]
        
        count = 1
        for row in rows:

            if count > 1:

                if len(row.item1) > cell_length[0]:
                    cell_length[0] = len(row.item1)

                if len(row.item2) > cell_length[1]:
                    cell_length[1] = len(row.item2)

                if len(row.item3) > cell_length[2]:
                    cell_length[2] = len(row.item3)

                if len(row.item4) > cell_length[3]:
                    cell_length[3] = len(row.item4)

                if len(row.item5) > cell_length[4]:
                    cell_length[4] = len(row.item5)

                if len(row.item6) > cell_length[5]:
                    cell_length[5] = len(row.item6)

                if len(row.item7) > cell_length[6]:
                    cell_length[6] = len(row.item7)

                if len(row.item8) > cell_length[7]:
                    cell_length[7] = len(row.item8)

            else:
                cell_length[0] = len(row.item1)
                cell_length[1] = len(row.item2)
                cell_length[2] = len(row.item3)
                cell_length[3] = len(row.item4)
                cell_length[4] = len(row.item5)
                cell_length[5] = len(row.item6)
                cell_length[6] = len(row.item7)
                cell_length[7] = len(row.item8)

            count += 1

        print(cell_length) #เเสดงความยาวสูงสุดของเเต่ละ cell

        columns_names = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H']
        
        for i in range(len(columns_names)):

            sheet.column_dimensions[columns_names[i]].width = (cell_length[i] + 5) #ขยายความกว้างของ column
        
        
        file.save(file_path) #บันทึกออกมาเป็นไฟล์ Excel

        return FileResponse(file_path)
    
    except Exception:
        print('error')




@app.post("/image/tesseract")
async def extract_receipt_information(file: UploadFile):

    contents = await file.read() #อ่านข้อมูลทั้งหมดจากไฟล์
    img = pre_image.preprocessing(contents)


    result = [] #ประกาศตัวเเปรสำหรับเก็บข้อมูลของใบเสร็จตาม pattern ที่กำหนด
    current_index = 0
    company_name = ''

    bigc_pattern = r'^\d+.\d+\s+\d{13}'
    makro_pattern = r'^\d+\s+\d{13}'
    lotus_pattern = r'^\d{14}\s+\w+'

    try:
        text = pytesseract.image_to_string(img, lang='tha+eng', config='--psm 6') #เเปลงรูปภาพใบเสร็จไปเป็น text
        
        text_line = text.split('\n') #เเบ่งบรรทัดตามการขึ้นบรรทัดใหม่ \n

        for i in range(len(text_line)):

            if re.match(bigc_pattern, text_line[i]):
                print('Hello big c : ', text_line[i])

                company_name = 'บริษัท บิ๊กซี ซูเปอร์เซ็นเตอร์ จำกัด (มหาชน)'
                current_index = i

                break

            elif re.match(makro_pattern, text_line[i]):
                print('Hello makro : ', text_line[i])

                company_name = 'บริษัท ซีพี เเอ็กซ์ตร้า จำกัด (มหาชน)'
                current_index = i

                break

            elif re.match(lotus_pattern, text_line[i]):
                print('Hello lotus : ', text_line[i])

                company_name = 'บริษัท เอก-ชัย ดีสทริบิวชั่น ซิสเทม จำกัด'
                current_index = i

                break
            else:
                company_name = 'not found'

        if company_name != 'not found':

            #กำหนดชื่อบริษัท
            result.append({
                "item1": company_name
            })

            #กำหนด pattern สำหรับเก็บข้อมูล
            result.append({
                "item1": "จำนวน",
                "item2": "รหัสสินค้า",
                "item3": "รายการสินค้า",
                "item4": "หน่วยบรรจุ",
                "item5": "ราคาต่อหน่วย",
                "item6": "ส่วนลด บาท",
                "item7": "รหัสภาษี",
                "item8": "จำนวนเงิน (รวม VAT)"
            })

            print('ตำเเหน่งของ array ปัจจุบัน = ', current_index)
            print(text_line[current_index])

            result = await search_important_data(result, text_line, current_index, bigc_pattern, makro_pattern, lotus_pattern)
    
        else:
            print('รูปภาพไม่ถูกต้อง ต้องเป็น makro bigc lotus เท่านั้น')

            result = []

    except Exception:
        result = []
    
    
    return {
        "result": result
    }



async def search_important_data(result, text_line, current_index, bigc_pattern, makro_pattern, lotus_pattern):
    payment_amount = 0 #ประกาศตัวเเปรสำหรับเก็บ ยอดเงินชำระ
    count = 0

    i = current_index

    while i < len(text_line):

        number = 0 #ประกาศตัวเเปรสำหรับเก็บจำนวนของบรรทัดที่จะบวกไปหาข้อมูลที่เกินของรายการสินค้าในเเถวนั้นๆ
        over = "" #สำหรับเก็บข้อมูลที่เกินของรายการสินค้าในเเถวนั้นๆออกมา
        product_name = "" #ประกาศตัวเเปรสำหรับสร้างรายการสินค้าที่มีการ join ข้อมูลใน List

        if count >= 3:
            print('วนลูปจนถึงเเถวที่ไม่ต้องการ ทำการหยุดลูป โดยที่ count = ', count)
            break


        if re.match(bigc_pattern, text_line[i]):

            words = text_line[i].split() #เเบ่งข้อความตามการเว้นวรรค
            print(words)


            count2 = 1 #สำหรับนับจำนวนของบรรทัดที่จะบวกไปหาข้อมูลที่เกินของรายการสินค้าในเเถวนั้นๆ
            while True:
                if text_line[i+count2] == "": #ตรวจสอบว่าบรรทัดต่อๆไปของเเถวนั้นๆมีค่าว่างหรือไม่
                    count2 += 1 #ถ้ามีค่าว่างที่ให้ count2 บวกไปทีละ 1 เเล้ววนลูปไปเรื่อยๆจนกว่าจะเจอเเถวที่มีข้อมูล
                else:
                    number += count2 #ถ้าเจอเเถวมีข้อมูลก็ให้เก็บจำนวนของบรรทัดที่จะบวกไปหาข้อมูลที่เกินของรายการสินค้าในเเถวนั้นๆ เเละหยุดการทำงานของลูป
                    break


            if (not re.match(bigc_pattern, text_line[i+number])) and (not re.match(r'ออกแทนใบกํากับภาษีอย่างย่อ\s+เลขที่\s+\d+', text_line[i+number])):
                over = text_line[i+number] #นำข้อมูลที่เกินของรายการสินค้าในเเถวนั้นๆออกมา
                product_name = " ".join(words[2:-3]) + over #นำข้อมูลที่เกินมาต่อเข้ากับรายการสินค้าของตัวมัน

                i += number

            else:
                print("กรณีที่ไม่มีข้อความเกินจนขึ้นบรรทัดใหม่")
                product_name = " ".join(words[2:-3]) #เพิ่มรายการสินค้า, การ join หมายความว่านำข้อมูลใน List มารวมกันเเละเเทนที่ช่องที่ต่อกันด้วย " " หรือจะใส่ "-"


            result.append({
                "item1": words[0], #เพิ่มจำนวน
                "item2": words[1], #เพิ่มรหัสสินค้า
                "item3": product_name, #เพิ่มรายการสินค้า
                "item4": "", #เนื่องจากใบเสร็จ big c ไม่มี column สำหรับข้อมูลหน่วยบรรจุ ดังนั้นจึงใส่ค่าว่าง
                "item5": words[-3], #เพิ่มราคาต่อหน่วย (รวม VAT), การใช้ตัวเลขติดลบในการเข้าถึงสมาชิกของ List จะเป็นการเข้าถึงสมาชิกจากท้ายสุดมา -3 หมายถึงสมาชิกตัวที่สามจากด้านท้ายสุดของ List
                "item6": words[-2], #เพิ่มส่วนลด บาท, -2 หมายถึงสมาชิกตัวที่ 2 จากด้านท้ายสุดของ List
                "item7": "", #เนื่องจากใบเสร็จ big c ไม่มี column สำหรับข้อมูล VAT CODE ดังนั้นจึงใส่ค่าว่าง
                "item8": words[-1] #เพิ่มจำนวนเงิน (รวม VAT), -1 หมายถึงสมาชิกตัวเเรกจากด้านท้ายสุดของ List
            })

            payment_amount += float(words[-1]) #ทำการหาผลรวมสำหรับ ยอดเงินชำระ


        elif re.match(makro_pattern, text_line[i]):
            
            words = text_line[i].split() #เเบ่งข้อความตามการเว้นวรรค
            print(words)

            result.append({
                "item1": words[0], #เพิ่มจำนวน
                "item2": words[1], #เพิ่มรหัสสินค้า
                "item3": " ".join(words[2:-4]), #เพิ่มรายการสินค้า, การ join หมายความว่านำข้อมูลใน List มารวมกันเเละเเทนที่ช่องที่ต่อกันด้วย " " หรือจะใส่ "-"
                "item4": words[-4], #เพิ่มหน่วยบรรจุ, -4 หมายถึงสมาชิกตัวที่ 4 จากด้านท้ายสุดของ List
                "item5": words[-3], #เพิ่มราคาต่อหน่วย (รวม VAT), การใช้ตัวเลขติดลบในการเข้าถึงสมาชิกของ List จะเป็นการเข้าถึงสมาชิกจากท้ายสุดมา -3 หมายถึงสมาชิกตัวที่สามจากด้านท้ายสุดของ List
                "item6": "", #เนื่องจากใบเสร็จ makro ไม่มี column สำหรับข้อมูล ส่วนลด บาท ดังนั้นจึงใส่ค่าว่าง
                "item7": words[-2], #เพิ่มข้อมูล VAT CODE, -2 หมายถึงสมาชิกตัวที่สองจากด้านท้ายสุดของ List
                "item8": words[-1] #เพิ่มจำนวนเงิน (รวม VAT), -1 หมายถึงสมาชิกตัวเเรกจากด้านท้ายสุดของ List
            })

            payment_amount += float(words[-1]) #ทำการหาผลรวมสำหรับ ยอดเงินชำระ

        elif re.match(lotus_pattern, text_line[i]):
            
            words = text_line[i].split() #เเบ่งข้อความตามการเว้นวรรค
            print(words)

            result.append({
                "item1": words[-4], #เพิ่มจำนวน
                "item2": words[0], #เพิ่มรหัสสินค้า
                "item3": " ".join(words[1:-4]), #เพิ่มรายการสินค้า, การ join หมายความว่านำข้อมูลใน List มารวมกันเเละเเทนที่ช่องที่ต่อกันด้วย " " หรือจะใส่ "-"
                "item4": "", #เพิ่มหน่วยบรรจุ, -4 หมายถึงสมาชิกตัวที่ 4 จากด้านท้ายสุดของ List
                "item5": words[-3], #เพิ่มราคาต่อหน่วย (รวม VAT), การใช้ตัวเลขติดลบในการเข้าถึงสมาชิกของ List จะเป็นการเข้าถึงสมาชิกจากท้ายสุดมา -3 หมายถึงสมาชิกตัวที่สามจากด้านท้ายสุดของ List
                "item6": "", #เนื่องจากใบเสร็จ makro ไม่มี column สำหรับข้อมูล ส่วนลด บาท ดังนั้นจึงใส่ค่าว่าง
                "item7": "", #เพิ่มข้อมูล VAT CODE, -2 หมายถึงสมาชิกตัวที่สองจากด้านท้ายสุดของ List
                "item8": words[-2] + " บ" #เพิ่มจำนวนเงิน (รวม VAT), -1 หมายถึงสมาชิกตัวเเรกจากด้านท้ายสุดของ List
            })

            payment_amount += float(words[-2]) #ทำการหาผลรวมสำหรับ ยอดเงินชำระ

        else:
            count += 1


        i += 1

    #เพิ่มข้อมูล ยอดเงินชำระ
    result.append({
        "item1": "ยอดเงินชำระ",
        "item2": "",
        "item3": "",
        "item4": "",
        "item5": "",
        "item6": "",
        "item7": "",
        "item8": "{:.2f}".format(payment_amount)
    })

    return result