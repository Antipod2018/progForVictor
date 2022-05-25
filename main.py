import requests, time, openpyxl, time, datetime, urllib
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
now = datetime.datetime.now() - delta


print('Время сейчас', now.date(), now.time() )

session = vk_api.VkApi(token=token)
vk = session.get_api()


def add_photo(user_id, album_id, text, attch_url):
    try:
        img = urllib.request.urlopen(attch_url).read()
        out = open("img.jpg", "wb")
        out.write(img)
        out.close()
        upload_url = session.method("photos.getUploadServer", {"album_id": album_id, "group_id": user_id})['upload_url']
        request = requests.post(upload_url, files={'file': open("img.jpg", "rb")})
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


VS=[[],[]]
time_d_post = datetime.timedelta(minutes=1)

while True:
    wb = openpyxl.open('Primer.xlsx', read_only=True)
    sheet = wb.active
    for row in range(2, sheet.max_row + 1):
        try:
            for i in range(len(sheet[row])):
                if sheet[row][i].value != VS[row][i]:
                    print('В ',row,' строке внесены изменения')

                    for j in range(len(VS[row])):
                        VS[row][j]=sheet[row][j].value

        except IndexError:
            VS.append([])
            for i in range(len(sheet[row])):
                VS[row].append('')
                VS[row][i]=sheet[row][i].value
    wb.close()
    now = datetime.datetime.now() - delta
    for row in range(2, sheet.max_row + 1):
        try:
            if VS[row][2] and VS[row][3]:

                if VS[row][0] or VS[row][1]:
                    time_date = datetime.datetime.combine(VS[row][2], VS[row][3])
                    if now > time_date and abs(now - time_date) < time_d_post:
                        add_t = add_photo(group_id, album_id, VS[row][0], VS[row][1])
                        if add_t:
                            print("Пост под номером ", row, ' в таблице, был добавлен с фото')
                        else:
                            print("Пост под номером ", row, ' в таблице, был добавлен без фото')

                        time.sleep(61)
        except TypeError:
            print('Поход в строке ', row, ' не правильно указана дата или время')

    time.sleep(5)