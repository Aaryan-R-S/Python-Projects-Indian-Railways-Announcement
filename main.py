# ------------------------ 1 ---------------------- imports
import pandas
from gtts import gTTS 
import os

# ------------------------ 2 ---------------------- speak
def speak(announcement):
    lang = 'hi'
    speechObj = gTTS(text=announcement, lang=lang, slow=False)
    speechObj.save("announce.mp3")
    os.system("start announce.mp3") 

# ------------------------ 3 ---------------------- annoucement
def skeleton(filename, no):
    b = None            #from
    d = None            #via
    f = None            #to
    h = None            # no name
    j = None            # platform
    excel = pandas.read_excel(filename)
    # print(excel)
    for i, item in excel.iterrows():
        if item['train_no'] == no:
            b = item['from']
            d = item['via']
            f = item['to']
            h = item['train_no']+" "+item['train_name']
            j = item['platform']

    announcement = f"Yatrigan! kripya! dhyan! dijiye!{b} se chlkar!!! {d} ke raste!!! {f} ko jaane waali!!! Gaadi sankhya {h}!!! kuch hi samay me!! platform kramank!! {j}!! pe! aa rahi hai.."
    
    return announcement

# ------------------------ 3 ---------------------- run
if __name__ == "__main__":
    no = input("\nEnter Train No.\n ")
    announcement = skeleton("train.xlsx", no)
    print("Announcing...\n")
    speak(announcement)
