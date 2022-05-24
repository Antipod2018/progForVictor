import urllib.request
import requests, time, openpyxl, time, datetime
import vk_api

with open("config.txt", 'r', encoding='utf-8') as f:
    config = f.read()
f.close()
config = config.split('\n')
token = config[0]
group_id = int(config[1])
album_id = int(config[2])
sheet_path = config[3]
time_d = int(config[4])
delta = datetime.timedelta(hours=time_d)
now = datetime.datetime.now()
t = now - delta

print('Время сейчас', t.date(), t.time() )

session = vk_api.VkApi(token=token)
vk = session.get_api()

def add_photo(user_id, album_id, text, attch_url):
    try:
        img = urllib.request.urlopen(attch_url).read()
        out = open("system\\img.jpg", "wb")
        out.write(img)
        out.close()
        upload_url = session.method("photos.getUploadServer", {"album_id": album_id, "group_id": user_id})['upload_url']
        request = requests.post(upload_url, files={'file': open("system\\img.jpg", "rb")})
    except Exception:
        session.method("wall.post",
                       {"owner_id": -user_id, "from_group": 1, "message": text})
        print('ссылка говно или кодер говно, фотки не будет')
        return 0
    time.sleep(5)
    if not request.json()['photos_list']:
        print('ссылка говно, фотки не будет')
        return 0


    save_photo = session.method("photos.save", {
        "album_id": album_id,
        "group_id": user_id,
        "server": request.json()['server'],
        "photos_list": request.json()['photos_list'],
        "hash": request.json()['hash']
    })

    attc_name = str("photo" + str(save_photo[0]['owner_id']) + "_" + str(save_photo[0]['id']))
    session.method("wall.post",
                           {"owner_id": -user_id, "from_group": 1, "message": text, "attachments": attc_name})
    return 1

if __name__ == '__main__':




    while(time.time()):
        wb = openpyxl.open(sheet_path, read_only=True)
        wb1 = openpyxl.open('system\\не_трогать.xlsx', read_only=False)
        sheet = wb.active
        sheet1 = wb1.active

        for row in range(2, sheet.max_row + 1):
            t = False
            for i in range(4):
                if sheet[row][i].value != sheet1[row][i].value:
                    t = True
                    sheet1[row][i].value = sheet[row][i].value
                    sheet1.cell(row=row, column=5, value=1)


                    wb1.save('system\\не_трогать.xlsx')
        wb.close()
        now = datetime.datetime.now()


        for row in range(2, sheet.max_row + 1):
            if sheet1[row][0].value or sheet1[row][1].value:
                if sheet1[row][2].value and sheet1[row][3].value and sheet1[row][4].value:
                    try: d =datetime.datetime.combine(sheet1[row][2].value.date(), sheet1[row][3].value)
                    except Exception:
                        d = datetime.datetime.combine(sheet1[row][2].value.date(), sheet1[row][3].value.time())

                    if d < now - delta:
                        add_t = add_photo(group_id, album_id, sheet1[row][0].value, sheet1[row][1].value)
                        if add_t:
                            print("Пост под номером ", row, ' в таблице, был добавлен с фото')
                        else:
                            print("Пост под номером ", row, ' в таблице, был добавлен без фото')
                        sheet1[row][4].value = 0
                        wb1.save('system\\не_трогать.xlsx')

        time.sleep(20)



