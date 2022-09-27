from pptx import Presentation
import json
import os
import qrcode

user_input_choices = "1. Generate a ppt based on an existing URL prefix.\n2. Create the URL prefix.\n3. Change the URL prefix.\n4. End\n"

def generate_ppt(URL_object):

  course_name = input("Enter the course name.\n")
  special_characters = '\/:*?|"<>'
  for z in special_characters:
    if z in course_name:
      print("Course name cannot contain any of the follower characters: ", special_characters)
      print("File is not saved, returning to main page.")
      return
  if course_name.isspace():
    print("Your course name cannot only contains whitespace.")
    print("File is not saved, returning to main page.")
    return

  new_course_name = "%20".join(course_name.split())
  string_URL = URL_object["URL"]+"?"+URL_object["field_id"]+"="+new_course_name
  
  prs = Presentation("template.pptx")
  slide = prs.slides[0]

  shapes =[shape for shape in slide.shapes]
  string = old_QR_code = None

  for i in shapes:
    if i.name == "URL":
      string = i
    elif i.name == "QRCode":
      old_QR_code = i

  if string == None:
    print("No text box for URL found. Returning to main page.")
    return 
  if old_QR_code == None:
    print("No image box for QR code found. Returning to main page.")
    return

  cur_text = string.text_frame.paragraphs[0].runs[0].text
  new_text = cur_text.replace(cur_text,string_URL)
  string.text_frame.paragraphs[0].runs[0].text = new_text
  
  img = qrcode.make(string_URL)
  img_file_name = course_name + ".png"
  img.save(img_file_name)

  with open(img_file_name, 'rb') as f:
    rImgBlob = f.read()

  imgPic = old_QR_code._pic
  imgRID = imgPic.xpath('./p:blipFill/a:blip/@r:embed')[0]
  imgPart = slide.part.related_part(imgRID)
  imgPart._blob = rImgBlob

  os.remove(img_file_name)
  prs.save(course_name + ".pptx")
  print(course_name, ".pptx is saved to the following path: ", os.path.abspath(os.getcwd()))

def create_pre_fix_json():
  print("You need to enter the URL and field ID. You can use shift + CTRL + V to paste.")
  URL = input("Enter the URL: ")
  field_id = input("Enter the field ID: ")
  dic ={
    "URL": URL,
    "field_id": field_id
  }
  URL_object = json.dumps(dic, indent=4)
  with open("pre_fix.json", "w") as outfile:
    outfile.write(URL_object)
    print("pre_fix.json saved.")
    
def change_pre_fix_json():
  try:
    with open('pre_fix.json', 'r') as openFile:
          URL_object = json.load(openFile)

    print("URL : ", URL_object["URL"] , " field ID: ", URL_object["field_id"])
    user_input = input("You want to change \n1. URL\n2. field ID. \n3 return \n")

    if user_input == "1":
      # change URL
      URL = input("New URL: ")
      dic ={
        "URL": URL,
        "field_id": URL_object["field_id"]
      }

      URL_object = json.dumps(dic, indent=4)
      with open("pre_fix.json", "w") as outfile:
        outfile.write(URL_object)
        print("pre_fix.json saved.")

    elif user_input == "2":
      # change field id
      field_id = input("New field id: ")
      dic ={
        "URL": URL_object["URL"],
        "field_id": field_id
      }

      URL_object = json.dumps(dic, indent=4)
      with open("pre_fix.json", "w") as outfile:
        outfile.write(URL_object)
        print("pre_fix.json saved.")

    elif user_input == '3':
      return

    else:
      print("Wrong input.")

  except:
    print("No pre_fix.json file found in directory. You can choose 2 to create a pre_fix.json file.")

if __name__ == "__main__":
  while(1):
    print("#################################################################")
    user_input = input(user_input_choices)

    if user_input == '1':
      try:
        with open('pre_fix.json', 'r') as openFile:
          URL_object = json.load(openFile)
        
      except:
        print("No pre_fix.json file found in directory. You can choose 2 to create a pre_fix.json file.")
        continue
      generate_ppt(URL_object)

    elif user_input == '2':
      # ask for URL pre-fix
      create_pre_fix_json()
      
    elif user_input == '3':
      # change URL pre-fix
      change_pre_fix_json()
    
    elif user_input == '4':
      break
    else:
      print("You can only choose one from 1-4. ")
      